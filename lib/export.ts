// Builds the Excel file from the user's own import sheet (public/template.xlsx),
// so fonts, colours, widths, header rows, merged CHA groups and the "Format" row stay exactly as in their format.
import { COLS, type Row } from './rules';

const LAST = 36; // column AJ
const F: Record<string, (r: number, e: number) => string> = {
  inr: r => `H${r}*F${r}`, basic: r => `M${r}-O${r}`, tDuty: (r, e) => `SUM(N${r}:N${e})`, tIgst: (r, e) => `SUM(O${r}:O${e})`,
  dutyIgst: r => `U${r}+T${r}`, frTotal: r => `X${r}*W${r}`, tds: r => `Z${r}*1%`, net: r => `AG${r}-AH${r}`,
};
const DATES = new Set(['invDate', 'boeDate', 'chaDate']);
const MERGE_COLS = ['Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI']; // per CHA invoice group
const toDate = (s: string) => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s); return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : s; };

export async function buildWorkbook(ExcelJS: any, tpl: ArrayBuffer, rows: Row[], withRemarks: boolean) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(tpl);
  const ws = wb.worksheets[0];
  let fmt = 0; ws.eachRow((r: any, n: number) => { if (String(r.getCell(1).value) === 'Format') fmt = n; });
  if (!fmt) throw new Error('template.xlsx: "Format" row not found');
  const snap = (n: number) => ({ h: ws.getRow(n).height, cells: Array.from({ length: LAST }, (_, i) => { const c = ws.getRow(n).getCell(i + 1); return { v: c.value, s: JSON.parse(JSON.stringify(c.style || {})) }; }) });
  const dataRow = snap(3), fmtRow = snap(fmt);

  [...ws.model.merges].forEach((m: string) => ws.unMergeCells(m));
  for (let r = 3; r <= fmt; r++) for (let c = 1; c <= LAST; c++) { const cell = ws.getRow(r).getCell(c); cell.value = null; cell.style = {}; }

  const ajStyle = JSON.parse(JSON.stringify(fmtRow.cells[LAST - 1].s)); ajStyle.alignment = { vertical: 'center', wrapText: false, horizontal: 'left' };
  const RED = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };

  rows.forEach((row, i) => {
    const r = i + 3; let j = i;
    while (j + 1 < rows.length && rows[j + 1].c.cha === undefined && rows[j + 1].c.tDuty === undefined) j++;
    const e = j + 3, first = row.c.tDuty !== undefined || row.c.cha !== undefined;
    COLS.forEach(([k, col]) => {
      const cell = ws.getCell(`${col}${r}`);
      cell.style = JSON.parse(JSON.stringify(dataRow.cells[colNum(col) - 1].s));
      if (k === 'remark') { cell.style = JSON.parse(JSON.stringify(ajStyle)); return; }
      let val = row.v[k]; if (val === undefined || val === null || val === '') return;
      if (DATES.has(k) && typeof val === 'string') val = toDate(val);
      if (k === 'awb' && /^\d{1,15}$/.test(String(val))) val = Number(val);
      cell.value = F[k] ? ({ formula: F[k](r, e), result: val } as any) : val;
      if (withRemarks && row.c[k] && row.c[k].s !== 'OK') cell.fill = RED;
    });
    for (let c = 1; c <= LAST; c++) { const cell = ws.getRow(r).getCell(c); if (!cell.style || !Object.keys(cell.style).length) cell.style = JSON.parse(JSON.stringify(dataRow.cells[c - 1].s)); }
    if (withRemarks && row.v.remark && row.v.remark !== 'All checks passed') ws.getCell(`AJ${r}`).value = row.v.remark;
    if (first && j > i) MERGE_COLS.forEach(c => ws.mergeCells(`${c}${r}:${c}${e}`));
  });

  const f = rows.length + 3;
  fmtRow.cells.forEach((c, i) => { const cell = ws.getRow(f).getCell(i + 1); cell.value = c.v; cell.style = c.s; });
  if (fmtRow.h) ws.getRow(f).height = fmtRow.h;
  return wb;
}

function colNum(col: string) { return col.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0); }
