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

export type Meta = { supplier: string; total: number; matched: number; counts: [string, number][]; warnings: string[] };
const NUM: Record<string, string> = {
  usd: '#,##0.00', rate: '0.00', inr: '#,##0.00', gw: '#,##0.00', duty: '#,##0', basic: '#,##0', igst: '#,##0',
  tDuty: '#,##0', tIgst: '#,##0', dutyIgst: '#,##0', chWt: '#,##0.##', frRate: '#,##0.##', frTotal: '#,##0.00', clrRate: '#,##0.##',
  clr: '#,##0.00', fuel: '#,##0.00', dem: '#,##0.00', hand: '#,##0.00', gst: '#,##0.00', invAmt: '#,##0.00', tds: '#,##0.00', net: '#,##0.00', qty: '#,##0',
};
const BAND = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF2FB' } };

export async function buildWorkbook(ExcelJS: any, tpl: ArrayBuffer, rows: Row[], withRemarks: boolean, meta?: Meta) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(tpl);
  const ws = wb.worksheets[0];
  let fmt = 0; ws.eachRow((r: any, n: number) => { if (String(r.getCell(1).value) === 'Format') fmt = n; });
  if (!fmt) throw new Error('template.xlsx: "Format" row not found');
  const snap = (n: number) => ({ h: ws.getRow(n).height, cells: Array.from({ length: LAST }, (_, i) => { const c = ws.getRow(n).getCell(i + 1); return { v: c.value, s: JSON.parse(JSON.stringify(c.style || {})) }; }) });
  const dataRow = snap(3), fmtRow = snap(fmt); // Format row is only used for its styling, never written to the output

  // Row 1 only explains where each column comes from (Ruby / BOE / CHA Details ...): not part of the output
  for (let c = 1; c <= LAST; c++) ws.getRow(1).getCell(c).value = null;

  [...ws.model.merges].forEach((m: string) => ws.unMergeCells(m));
  for (let r = 3; r <= fmt; r++) for (let c = 1; c <= LAST; c++) { const cell = ws.getRow(r).getCell(c); cell.value = null; cell.style = {}; }

  const ajStyle = JSON.parse(JSON.stringify(fmtRow.cells[LAST - 1].s)); ajStyle.alignment = { vertical: 'center', wrapText: false, horizontal: 'left' };
  const RED = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };

  let band = -1;
  rows.forEach((row, i) => {
    const r = i + 3; let j = i;
    while (j + 1 < rows.length && rows[j + 1].c.cha === undefined && rows[j + 1].c.tDuty === undefined) j++;
    const e = j + 3, first = row.c.tDuty !== undefined || row.c.cha !== undefined;
    if (first) band++;
    ws.getRow(r).height = 24;
    COLS.forEach(([k, col]) => {
      const cell = ws.getCell(`${col}${r}`);
      cell.style = JSON.parse(JSON.stringify(dataRow.cells[colNum(col) - 1].s));
      if (k === 'remark') { cell.style = JSON.parse(JSON.stringify(ajStyle)); return; }
      let val = row.v[k]; if (val === undefined || val === null || val === '') return;
      if (DATES.has(k) && typeof val === 'string') val = toDate(val);
      if (k === 'awb' && /^\d{1,15}$/.test(String(val))) val = Number(val);
      cell.value = F[k] ? ({ formula: F[k](r, e), result: val } as any) : val;
      if (NUM[k]) cell.numFmt = NUM[k];
      cell.alignment = { vertical: 'middle', horizontal: NUM[k] ? 'right' : k === 'sno' || DATES.has(k) || k === 'awb' ? 'center' : 'left', wrapText: false };
      if (withRemarks && row.c[k] && row.c[k].s !== 'OK') cell.fill = RED;
    });
    for (let c = 1; c <= LAST; c++) {
      const cell = ws.getRow(r).getCell(c);
      if (!cell.style || !Object.keys(cell.style).length) cell.style = JSON.parse(JSON.stringify(dataRow.cells[c - 1].s));
      cell.alignment = { ...(cell.alignment || {}), vertical: 'middle' };
      const isErr = withRemarks && cell.fill && (cell.fill as any).fgColor?.argb === 'FFFFC7CE';
      if (band % 2 === 1 && !isErr && c !== LAST) cell.fill = BAND;
    }
    if (withRemarks && row.v.remark && row.v.remark !== 'All checks passed') { const a = ws.getCell(`AJ${r}`); a.value = row.v.remark; a.font = { name: 'Times New Roman', size: 12, color: { argb: 'FFC00000' } }; }
    if (first && j > i) MERGE_COLS.forEach(c => ws.mergeCells(`${c}${r}:${c}${e}`));
  });

  if (withRemarks) ws.getColumn('AJ').width = 90;
  if (meta) {
    const sm = wb.addWorksheet('Summary', { views: [{ showGridLines: false }] });
    sm.columns = [{ width: 4 }, { width: 38 }, { width: 70 }];
    const T = (c: string, v: any, o: any = {}) => { const x = sm.getCell(c); x.value = v; x.font = { name: 'Times New Roman', size: 13, ...o.font }; if (o.fill) x.fill = o.fill; x.alignment = { vertical: 'middle', wrapText: true }; return x; };
    sm.mergeCells('B2:C2'); T('B2', 'Import Sheet Check Summary', { font: { size: 20, bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } } }); sm.getRow(2).height = 36;
    const kv: [string, any][] = [['Supplier', meta.supplier], ['Checked on', new Date().toLocaleString('en-IN')], ['BOE rows checked', meta.total], ['Rows fully matched', meta.matched], ['Rows with issues', meta.total - meta.matched]];
    kv.forEach(([k, v], i) => { const r = 4 + i; T(`B${r}`, k, { font: { bold: true }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } } }); T(`C${r}`, v); sm.getCell(`C${r}`).alignment = { horizontal: 'left' }; });
    let r = 4 + kv.length + 1;
    if (meta.counts.length) { sm.mergeCells(`B${r}:C${r}`); T(`B${r}`, 'Issues by attribute (number of rows)', { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } } }); r++; meta.counts.forEach(([k, n]) => { T(`B${r}`, k); T(`C${r}`, n); sm.getCell(`C${r}`).alignment = { horizontal: 'left' }; r++; }); r++; }
    if (meta.warnings.length) { sm.mergeCells(`B${r}:C${r}`); T(`B${r}`, 'Warnings', { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBF8F00' } } }); r++; meta.warnings.forEach(w => { sm.mergeCells(`B${r}:C${r}`); T(`B${r}`, w); r++; }); }
  }
  return wb;
}

function colNum(col: string) { return col.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0); }
