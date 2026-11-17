// Reads shipper invoice (.xls/.xlsx), BOE (PDF) and CHA invoice/debit note (PDF)
// fully inside the browser. No server, no AI: fixed labels + positions.
import type { Ship, Boe, Cha } from './rules';

export type Kind = 'ship' | 'boe' | 'cha';
type It = { s: string; r: string; x: number; y: number; w: number };
export type Page = { items: It[]; lines: string[] };

const num = (s?: string | null) => { if (s == null) return undefined; const n = Number(String(s).replace(/,/g, '')); return Number.isFinite(n) ? n : undefined; };
const r2 = (x: number) => Math.round(x * 100) / 100;
const MON: Record<string, string> = { jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12' };

export function toIso(v: any): string | undefined {
  if (v instanceof Date && !isNaN(+v)) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
  const s = String(v ?? '').trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})[-\/. ]([A-Za-z]{3})[A-Za-z]*[-\/. ,]+(\d{4})$/); if (m && MON[m[2].toLowerCase()]) return `${m[3]}-${MON[m[2].toLowerCase()]}-${m[1].padStart(2, '0')}`;
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/); if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  return undefined;
}

/* ------------------------------ PDF text ------------------------------ */
export function buildPage(raw: { str: string; transform: number[]; width?: number }[]): Page {
  const items: It[] = raw.filter(i => i.str && i.str.trim()).map(i => ({ s: i.str.replace(/\s+/g, ' ').trim(), r: i.str, x: i.transform[4], y: i.transform[5], w: i.width ?? 0 }));
  const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
  const lines: string[] = []; let cur: It[] = []; let cy = 0;
  const flush = () => {
    cur.sort((a, b) => a.x - b.x);
    let t = '', prev: It | undefined;
    for (const it of cur) {
      // Some PDFs split a word into several items with no gap: glue those, put a space only across real gaps
      const gap = prev ? it.x - (prev.x + prev.w) : 0;
      if (prev && gap > 1.5 && !/\s$/.test(t) && !/^\s/.test(it.r)) t += ' ';
      t += it.r; prev = it;
    }
    if (t.trim()) lines.push(t.replace(/\s+/g, ' ').trim());
    cur = [];
  };
  for (const it of sorted) { if (cur.length && Math.abs(it.y - cy) > 2.5) flush(); if (!cur.length) cy = it.y; cur.push(it); }
  flush();
  return { items, lines };
}

export async function pdfToPages(buf: ArrayBuffer): Promise<Page[]> {
  // @ts-ignore legacy build has no bundled types path
  const pdfjs: any = await import('pdfjs-dist/legacy/build/pdf.mjs');
  if (typeof window !== 'undefined') pdfjs.GlobalWorkerOptions.workerSrc = '/pdf.worker.min.mjs';
  const doc = await pdfjs.getDocument({ data: new Uint8Array(buf.slice(0)), useSystemFonts: true }).promise;
  const pages: Page[] = [];
  try {
    for (let i = 1; i <= doc.numPages; i++) { const p = await doc.getPage(i); pages.push(buildPage((await p.getTextContent()).items)); p.cleanup(); }
  } finally { await doc.destroy(); }
  return pages;
}

// Value cell to the right of a label cell (same row). wrap: also join cells wrapped directly below it.
function valuesAfter(pg: Page, label: RegExp, wrap = false, join = ''): string[] {
  const out: string[] = [];
  for (const l of pg.items) {
    if (!label.test(l.s)) continue;
    const c = pg.items.filter(i => i !== l && Math.abs(i.y - l.y) < 2 && i.x > l.x).sort((a, b) => a.x - b.x)[0];
    if (!c) continue;
    let v = c.s, y = c.y;
    while (wrap) {
      const n = pg.items.find(i => Math.abs(i.x - c.x) < 1.5 && i.y < y - 8 && i.y > y - 11);
      if (!n) break; v += join + n.s; y = n.y;
    }
    out.push(v);
  }
  return out;
}
const firstOf = (pages: Page[], label: RegExp, wrap = false, join = '') => { for (const p of pages) { const v = valuesAfter(p, label, wrap, join)[0]; if (v) return v; } return undefined; };

/* --------------------------------- BOE --------------------------------- */
export function parseBoe(pages: Page[]): Boe[] {
  const groups: Page[][] = [];
  for (const p of pages) {
    const isStart = p.lines.some(l => /Courier Bill Of Entry\s*-?\s*XIII/i.test(l)) && p.lines.some(l => /Current Status of the CBE/i.test(l));
    if (isStart || !groups.length) groups.push([p]); else groups[groups.length - 1].push(p);
  }
  const out: Boe[] = [];
  for (const g of groups) {
    const lines = g.flatMap(p => p.lines);
    if (!lines.some(l => /Current Status of the CBE/i.test(l))) continue;
    const text = lines.join('\n');
    const o = text.match(/OOC ISSUED on\s+(\d{2})-(\d{2})-(\d{4})/i);
    let igst = 0, found = false;
    for (const l of lines) { const m = l.match(/^\d+\s+IGST\s+[\d.]+\s+[\d.]+\s+[\d.]+\s+([\d.]+)$/); if (m) { igst += Number(m[1]); found = true; } }
    let airline: string | undefined, flight: string | undefined;
    const hi = lines.findIndex(l => /^Airlines\s+Flight No/i.test(l));
    if (hi >= 0 && lines[hi + 1]) { const m = lines[hi + 1].match(/^(.*?)\s+([A-Z0-9]{2}\s?\d{2,4})\s+[A-Z]{3}\s+[A-Z]{3}\b/); airline = m?.[1]?.trim(); flight = m?.[2]?.replace(/\s+/g, ''); }
    const invNos = [...new Set(lines.map(l => l.match(/Invoice Number\s*:\s*(\S+)/)?.[1]).filter(Boolean) as string[])];
    out.push({
      boeNo: firstOf(g, /^CBE-XIII Number$/i, true, ''),
      oocDate: o ? `${o[3]}-${o[2]}-${o[1]}` : undefined,
      invoiceNo: firstOf(g, /^Invoice Number\s*:?$/i) ?? invNos[0],
      shipper: firstOf(g, /^Name of Consignor\s*:?$/i, true, ' '),
      exchangeRate: num(firstOf(g, /^Rate of Exchange\s*:?$/i)),
      duty: num(firstOf(g, /^Duty\(Rs\.?\)\s*:?$/i)),
      igst: found ? r2(igst) : undefined,
      hawb: firstOf(g, /^HAWB Number\s*:?$/i),
      airline, flight,
    });
  }
  return out;
}

/* --------------------------------- CHA --------------------------------- */
const AMT = /\b\d[\d,]*\.\d{2}\b/g;
export function parseCha(pages: Page[]): Cha[] {
  const out: Cha[] = [];
  for (const p of pages) {
    const text = p.lines.join('\n'), flat = p.lines.join(' ');
    const isDebit = /DEBIT NOTE NO\s*:/.test(flat), isTax = !isDebit && /TAX INVOICE/.test(flat) && /INVOICE NO\s*:/.test(flat);
    if (!isDebit && !isTax) continue;
    const m = flat.match(/(?:INVOICE|DEBIT NOTE) NO\s*:\s*(.+?)\s+DATE\s*:\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    const c: Cha = { chaName: p.lines[0]?.trim(), invoiceNo: m?.[1].replace(/\s+/g, ''), invoiceDate: m ? `${m[4]}-${m[3].padStart(2, '0')}-${m[2].padStart(2, '0')}` : undefined };
    const notes: string[] = [];
    const hs = text.search(/HAWB NO/i), he = text.search(/Flight\s*:/i);
    c.awbs = hs >= 0 ? (text.slice(hs, he > hs ? he : hs + 600).match(/\b\d{10,12}\b/g) ?? []) : [];
    const fl = flat.match(/Flight\s*:\s*(.+?)(?:\s{2,}|$)/)?.[1] ?? p.lines.find(l => /^Flight\s*:/i.test(l))?.replace(/^Flight\s*:\s*/i, '');
    if (fl) c.flightText = fl.trim();
    if (fl) c.airline = fl.match(/^(.*?)\s+[A-Z0-9]{2}\s?\d{3,4}\b/)?.[1]?.trim() || fl.trim();
    c.chargeableWeight = num(flat.match(/CHARGEABLE WT\.?\s*:\s*([\d.,]+)/i)?.[1]);

    const start = p.lines.findIndex(l => /^PARTICULARS/.test(l));
    const end = p.lines.findIndex((l, i) => i > start && /^Total\b/.test(l));
    const body = start >= 0 ? p.lines.slice(start + 1, end > start ? end : undefined) : [];
    const sumRow = (k: 'clearance' | 'fuel' | 'demurrage' | 'handling' | 'dutyIgst' | 'totalFreight', v: number) => { c[k] = r2((c[k] ?? 0) + v); };
    let gst = 0, gstSeen = false;
    for (const l of body) {
      const g = l.match(/^(?:SGST|CGST|IGST)\b.*?@\s*[\d.]+\s*%\s+([\d,]+\.\d{2})\s*$/i);
      if (g) { gst += num(g[1])!; gstSeen = true; continue; }
      const row = l.match(/^\d+\.\s+(.*)$/); if (!row) continue;
      const fr = row[1].match(/Freight Charges\s*=?\s*INR\s*([\d.]+)\s*x\s*([\d.]+)\s*kg/i);
      const amts = (fr ? row[1].replace(fr[0], '') : row[1]).match(AMT)?.map(a => num(a)!) ?? [];
      const a = amts.reduce((t, x) => t + x, 0), has = amts.length > 0;
      if (fr) { c.freightRate = num(fr[1]); if (c.chargeableWeight === undefined) c.chargeableWeight = num(fr[2]); }
      if (/CUSTOM DUTY/i.test(row[1]) && /AMENDMENT/i.test(row[1])) { if (has) notes.push(`Duty/GST amendment charges of ${a} are billed (no column in the import sheet)`); }
      else if (/CUSTOM DUTY/i.test(row[1])) { if (has) sumRow('dutyIgst', a); }
      else if (/Freight/i.test(row[1])) { if (has) sumRow('totalFreight', a); }
      else if (/Origin Handling/i.test(row[1])) { if (has) sumRow('handling', a); }
      else if (/Demurrage|Detention/i.test(row[1])) { if (has) sumRow('demurrage', a); }
      else if (/Fuel/i.test(row[1])) { if (has) sumRow('fuel', a); }
      else if (/D\.?O\.?\s*\+?\s*Endorsement/i.test(row[1])) { if (has) notes.push(`Airlines D.O. + Endorsement fee of ${a} is billed (no column in the import sheet)`); }
      else if (/Clearance/i.test(row[1])) { if (has) sumRow('clearance', a); }
      else if (has) notes.push(`Unrecognised charge line "${row[1].replace(AMT, '').replace(/\s+/g, ' ').trim()}" of ${a}`);
    }
    if (gstSeen) c.gst = r2(gst);
    const tot = p.lines.map(l => l.match(/(?:^|\s)TOTAL\s+([\d,]+\.\d{2})\s*$/)?.[1]).find(Boolean);
    if (isDebit) c.debitNoteTotal = num(tot); else c.taxInvoiceTotal = num(tot);
    if (notes.length) c.notes = notes;
    out.push(c);
  }
  return out;
}

/* ----------------------------- Shipper invoice ----------------------------- */
export function parseShip(rows: any[][]): Ship {
  const cell = (r: any[], i: number) => (r[i] == null ? '' : String(r[i]).trim());
  const first = (re: RegExp) => { for (const r of rows) for (const v of r) { const m = String(v ?? '').match(re); if (m) return m; } return null; };
  const s: Ship = {};
  s.invoiceNo = first(/^Invoice Number\s*:\s*(.+)$/im)?.[1].trim();
  const d = first(/^Date\s*:\s*(.+)$/im)?.[1]; s.invoiceDate = d ? toIso(d.trim()) : undefined;
  s.shipper = first(/Shipper\/Exporter\s*:\s*\n?\s*([^\n]+)/i)?.[1].trim();
  const cur = first(/^Currency\s*:\s*(\w+)/im)?.[1];
  const hi = rows.findIndex(r => r.some(v => /^TOTAL PRICE/i.test(String(v ?? '').trim())) && r.some(v => /^QTY/i.test(String(v ?? '').trim())));
  if (hi >= 0) {
    const h = rows[hi], at = (re: RegExp) => h.findIndex(v => re.test(String(v ?? '').trim()));
    const qi = at(/^QTY/i), ui = at(/^TOTAL PRICE/i), gi = at(/^G\.?\s?W/i);
    let q = 0, u = 0, g = 0, tq: number | undefined, tu: number | undefined, n = 0;
    for (const r of rows.slice(hi + 1)) {
      if (/^TOTAL/i.test(cell(r, 0))) { tq = num(cell(r, qi)); tu = num(cell(r, ui)); break; }
      const a = num(cell(r, qi)), b = num(cell(r, ui)); if (a === undefined && b === undefined) continue;
      q += a ?? 0; u += b ?? 0; n++; if (gi >= 0) g += num(cell(r, gi)) ?? 0;
    }
    if (n) {
      s.qty = r2(q); s.usd = r2(u); s.grossWeight = gi >= 0 ? r2(g) : undefined;
      const diff: string[] = [];
      if (tq !== undefined && Math.abs(tq - q) > 0.01) diff.push(`TOTAL row qty ${tq} vs item sum ${r2(q)}`);
      if (tu !== undefined && Math.abs(tu - u) > 0.01) diff.push(`TOTAL row amount ${tu} vs item sum ${r2(u)}`);
      if (diff.length) s.note = `Invoice ${s.invoiceNo ?? ''}: ${diff.join('; ')}`;
    }
  }
  if (cur && cur.toUpperCase() !== 'USD') s.note = `${s.note ? s.note + '; ' : ''}Invoice ${s.invoiceNo ?? ''} currency is ${cur}, not USD`;
  return s;
}

/* ------------------------------ entry point ------------------------------ */
export async function extractFile(kind: Kind, f: File): Promise<any[]> {
  const buf = await f.arrayBuffer(), b = new Uint8Array(buf);
  const pdf = String.fromCharCode(...b.slice(0, 4)) === '%PDF';
  const xls = (b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0) || (b[0] === 0x50 && b[1] === 0x4b);
  if (kind === 'ship') {
    if (!xls) throw new Error('Shipper invoice must be an Excel file (.xls / .xlsx)');
    const XLSX: any = await import('xlsx');
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    let best: Ship | undefined;
    for (const n of wb.SheetNames) {
      const s = parseShip(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }));
      if (s.invoiceNo && s.qty !== undefined) { best = s; break; }
    }
    if (!best) throw new Error('Could not find an invoice (invoice number / item table) in this Excel file');
    return [best];
  }
  if (!pdf) throw new Error('This file must be a PDF');
  const pages = await pdfToPages(buf);
  if (!pages.some(p => p.items.length)) throw new Error('This PDF has no readable text (looks scanned). Use the original PDF from the portal/CHA');
  if (kind === 'boe') { const r = parseBoe(pages); if (!r.length) throw new Error('No Bill of Entry (CBE-XIII) found in this PDF'); return r; }
  const r = parseCha(pages); if (!r.length) throw new Error('No CHA tax invoice / debit note found in this PDF'); return r;
}
