export type Ship = { note?: string; shipper?: string; invoiceNo?: string; invoiceDate?: string; qty?: number; usd?: number; grossWeight?: number };
export type Boe = { invoiceNos?: string[]; boeNo?: string; oocDate?: string; invoiceNo?: string; shipper?: string; exchangeRate?: number; duty?: number; igst?: number; hawb?: string; airline?: string; flight?: string };
export type Cha = { flightText?: string; notes?: string[]; chaName?: string; invoiceNo?: string; invoiceDate?: string; dutyIgst?: number; chargeableWeight?: number; freightRate?: number; totalFreight?: number; clearanceRate?: number; clearance?: number; fuel?: number; demurrage?: number; handling?: number; gst?: number; taxInvoiceTotal?: number; debitNoteTotal?: number; airline?: string; awbs?: string[] };
export type St = 'OK' | 'ERROR' | 'MISSING';
export type Chk = { s: St; m: string };
export type Row = { v: Record<string, any>; c: Record<string, Chk> };

// [key, Excel column, label, source, rule]  (mirrors the "Format" row of the import sheet)
export const COLS: [string, string, string, string, string][] = [
['sno','A','S.NO','—','One row per BOE'],
['shipper','B','Shipper Company Name','Shipper Invoice','Must match the selected supplier'],
['invNo','C','Invoice No','Shipper Invoice','Must match the invoice number on the BOE'],
['invDate','D','Shipper Invoice Date','Shipper Invoice','Must not be after BOE date'],
['qty','E','Qty','Shipper Invoice','Qty sum'],
['usd','F','USD Amount','Shipper Invoice','Total price sum'],
['rate','H','Rate of Exchange','BOE','Same across the batch'],
['inr','I','INR Amount','Formula','H × F'],
['gw','J','Gross Weight (kg)','Shipper Invoice','G.W (KG) sum'],
['boeNo','K','BOE No.','BOE','CBE-XIII number, unique'],
['boeDate','L','BOE Date','BOE','OOC issued on; not before invoice date'],
['duty','M','Custom + Duty','BOE','Duty (Rs.) total incl. IGST'],
['basic','N','Duty','Formula','M − O'],
['igst','O','IGST','BOE','IGST (sum of all items)'],
['awb','P','AWB No','BOE','HAWB number; must be on CHA invoice'],
['cha','Q','CHA','CHA Invoice','CHA name'],
['chaNo','R','CHA Invoice No.','CHA Invoice','Invoice number'],
['chaDate','S','CHA Invoice Date','CHA Invoice','Not before BOE date'],
['tDuty','T','Custom Duty','Formula','Sum of N per CHA invoice'],
['tIgst','U','IGST (Custom)','Formula','Sum of O per CHA invoice'],
['dutyIgst','V','Duty + IGST as per CHA','Formula','U + T; compared with duty on CHA invoice'],
['chWt','W','Chargeable Weight','CHA Invoice','Must be ≥ total gross weight'],
['frRate','X','Freight Rate','CHA Invoice','Freight charges per kg'],
['airline','Y','Flight Name','BOE','Compared with CHA invoice'],
['frTotal','Z','Total Freight','Formula','X × W'],
['clrRate','AA','Clearance/Warehouse Rate','CHA Invoice','Rate (information only)'],
['clr','AB','Clearance/Warehouse','CHA Invoice','Amount'],
['fuel','AC','Fuel/Detention','CHA Invoice','Amount'],
['dem','AD','Demurrage','CHA Invoice','Amount'],
['hand','AE','Handling','CHA Invoice','Amount'],
['gst','AF','GST on Freight & Clearance','Formula','Total freight × 18%'],
['invAmt','AG','Invoice Amount','CHA Invoice','Tax invoice + debit note; must equal V+Z+AB+AC+AD+AE+AF'],
['tds','AH','TDS','Formula','Z × 1%'],
['net','AI','Net Payable','Formula','AG − AH'],
['remark','AJ','Remark','Formula','All issues for the row'],
];

const norm = (s?: string) => (s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const r2 = (x: number) => Math.round(x * 100) / 100;
const sum = (a: (number | undefined)[]) => a.reduce<number>((t, x) => t + (x ?? 0), 0);
const ok = (m = 'OK'): Chk => ({ s: 'OK', m });
const bad = (m: string): Chk => ({ s: 'ERROR', m });
const miss = (w: string): Chk => ({ s: 'MISSING', m: `${w} not found in the uploaded document` });
const req = (x: any, w: string) => (x === undefined || x === '' ? miss(w) : ok());
const pos = (x: number | undefined, w: string): Chk => (x === undefined ? miss(w) : x > 0 ? ok() : bad(`${w} must be greater than 0`));
const vs = (label: string, exp: number, got: number | undefined, tol = 1): Chk =>
  got === undefined ? { s: 'MISSING', m: `${label} not found on CHA invoice` }
  : Math.abs(exp - got) <= tol ? ok(`Matches CHA invoice (${got})`)
  : bad(`Calculated ${r2(exp)} but CHA invoice shows ${got} (difference ${r2(exp - got)})`);

export function run(supplier: string, ships: Ship[], boes: Boe[], chasIn: Cha[]) {
  const warnings: string[] = [];
  const chas = chasIn.map(c => ({ ...c }));
  // Debit notes are merged into the tax invoice they belong to
  const dn = chas.filter(c => c.taxInvoiceTotal === undefined && c.debitNoteTotal !== undefined);
  const base = chas.filter(c => !dn.includes(c));
  dn.forEach(d => {
    const t = base.find(c => c.awbs?.some(a => d.awbs?.some(x => norm(x) === norm(a)))) ?? base[0];
    if (!t) { warnings.push('Debit note uploaded without a CHA tax invoice'); return; }
    t.debitNoteTotal = (t.debitNoteTotal ?? 0) + (d.debitNoteTotal ?? 0);
    (['dutyIgst', 'totalFreight', 'clearance', 'fuel', 'demurrage', 'handling', 'gst'] as const).forEach(k => { const s = sum([t[k], d[k]]); if (s) t[k] = s; });
    (['freightRate', 'chargeableWeight', 'airline'] as const).forEach(k => { if (t[k] === undefined && d[k] !== undefined) (t as any)[k] = d[k]; });
    t.notes = [...(t.notes ?? []), ...(d.notes ?? [])];
  });

  [...new Set(base.flatMap(c => (c.notes ?? []).map(n => `CHA invoice ${c.invoiceNo ?? ''}: ${n}`)))].forEach(w => warnings.push(w));
  ships.forEach(s => { if (s.note) warnings.push(s.note); });
  const used = new Set<Ship>();
  const items = boes.map((b, i) => {
    const sh = ships.find(s => norm(s.invoiceNo) !== '' && norm(s.invoiceNo) === norm(b.invoiceNo))
      ?? (ships.length === boes.length && !used.has(ships[i]) ? ships[i] : undefined);
    if (sh) used.add(sh);
    const ch = base.find(c => c.awbs?.some(a => norm(a) === norm(b.hawb))) ?? (base.length === 1 ? base[0] : undefined);
    return { b, sh, ch };
  }).sort((x, y) => (x.ch ? base.indexOf(x.ch) : 999) - (y.ch ? base.indexOf(y.ch) : 999));
  ships.filter(s => !used.has(s)).forEach(s => warnings.push(`Shipper invoice ${s.invoiceNo ?? '(no number)'} has no matching BOE`));

  const rates = new Set(boes.map(b => b.exchangeRate).filter(x => x !== undefined));
  const seen = new Set<Cha>();
  const rows: Row[] = items.map(({ b, sh, ch }, i) => {
    const v: Record<string, any> = {}, c: Record<string, Chk> = {};
    const put = (k: string, val: any, chk: Chk) => { v[k] = val; c[k] = chk; };
    const shipper = sh?.shipper ?? b.shipper, invNo = sh?.invoiceNo ?? b.invoiceNo;
    const n1 = norm(shipper), n2 = norm(supplier);
    put('sno', i + 1, ok());
    put('shipper', shipper, !n1 ? miss('Shipper name') : n1.includes(n2) || n2.includes(n1) ? ok('Matches selected supplier') : bad(`Document shows "${shipper}" but selected supplier is "${supplier}"`));
    put('invNo', invNo, (b.invoiceNos?.length ?? 0) > 1 ? bad(`This BOE has items from ${b.invoiceNos!.length} invoices (${b.invoiceNos!.join(', ')}); the sheet expects one invoice per BOE`) : !sh ? bad(`No shipper invoice file matched this BOE (BOE shows invoice ${b.invoiceNo ?? 'not found'})`) : !invNo ? miss('Invoice number') : b.invoiceNo && norm(b.invoiceNo) !== norm(invNo) ? bad(`BOE shows invoice ${b.invoiceNo}`) : ok('Matches BOE'));
    put('invDate', sh?.invoiceDate, !sh?.invoiceDate ? miss('Invoice date') : b.oocDate && sh.invoiceDate > b.oocDate ? bad('Invoice date is after BOE date') : ok());
    put('qty', sh?.qty, pos(sh?.qty, 'Quantity'));
    put('usd', sh?.usd, pos(sh?.usd, 'USD amount'));
    put('rate', b.exchangeRate, b.exchangeRate !== undefined && rates.size > 1 ? bad(`Rate differs between BOEs (${[...rates].join(', ')})`) : pos(b.exchangeRate, 'Exchange rate'));
    const inr = b.exchangeRate !== undefined && sh?.usd !== undefined ? r2(b.exchangeRate * sh.usd) : undefined;
    put('inr', inr, inr === undefined ? miss('Rate or USD amount') : ok('H × F'));
    put('gw', sh?.grossWeight, pos(sh?.grossWeight, 'Gross weight'));
    put('boeNo', b.boeNo, !b.boeNo ? miss('BOE number') : boes.filter(x => norm(x.boeNo) === norm(b.boeNo)).length > 1 ? bad('Duplicate BOE number') : ok());
    put('boeDate', b.oocDate, !b.oocDate ? miss('OOC date') : sh?.invoiceDate && b.oocDate < sh.invoiceDate ? bad('BOE date is before invoice date') : ok());
    put('duty', b.duty, b.duty === undefined ? miss('Duty') : b.igst !== undefined && b.duty < b.igst ? bad('Duty is less than IGST') : ok());
    const basic = b.duty !== undefined && b.igst !== undefined ? r2(b.duty - b.igst) : undefined;
    put('basic', basic, basic === undefined ? miss('Duty or IGST') : ok('M − O'));
    put('igst', b.igst, req(b.igst, 'IGST'));
    put('awb', b.hawb, !b.hawb ? miss('HAWB number') : ch?.awbs?.length && !ch.awbs.some(a => norm(a) === norm(b.hawb)) ? bad('HAWB not listed on CHA invoice') : ok());

    if (!ch) put('cha', undefined, bad('No CHA invoice matched this BOE (HAWB not found on any CHA invoice)'));
    else if (!seen.has(ch)) {
      seen.add(ch);
      const g = items.filter(x => x.ch === ch);
      const T = sum(g.map(x => x.b.duty !== undefined && x.b.igst !== undefined ? x.b.duty - x.b.igst : undefined));
      const U = sum(g.map(x => x.b.igst)), V = r2(T + U);
      const W = ch.chargeableWeight, X = ch.freightRate, Z = W !== undefined && X !== undefined ? r2(W * X) : undefined;
      const gw = sum(g.map(x => x.sh?.grossWeight));
      put('cha', ch.chaName, req(ch.chaName, 'CHA name'));
      put('chaNo', ch.invoiceNo, req(ch.invoiceNo, 'CHA invoice number'));
      put('chaDate', ch.invoiceDate, !ch.invoiceDate ? miss('CHA invoice date') : g.some(x => x.b.oocDate && (ch.invoiceDate as string) < x.b.oocDate) ? bad('CHA invoice date is before a BOE date') : ok());
      put('tDuty', r2(T), ok(`Sum of duty for ${g.length} BOE(s)`));
      put('tIgst', r2(U), ok(`Sum of IGST for ${g.length} BOE(s)`));
      put('dutyIgst', V, vs('Duty + IGST', V, ch.dutyIgst));
      const wc = pos(W, 'Chargeable weight');
      put('chWt', W, wc.s === 'OK' && (W as number) < gw ? bad(`Chargeable weight is less than total gross weight (${r2(gw)} kg)`) : wc);
      put('frRate', X, pos(X, 'Freight rate'));
      const names = [...new Set(g.map(x => x.b.airline).filter(Boolean) as string[])];
      const disp = [...new Set(g.map(x => [x.b.airline, x.b.flight].filter(Boolean).join(' ')).filter(Boolean))].join(', ');
      const a2 = ch.airline, a2f = norm(ch.flightText ?? ch.airline);
      const wrong = a2 ? names.filter(n => !(a2f.includes(norm(n)) || norm(n).includes(norm(a2)))) : [];
      put('airline', disp || a2, !(disp || a2) ? miss('Airline') : wrong.length ? bad(`BOE shows ${wrong.join(', ')}, CHA invoice shows ${ch.flightText ?? a2}`) : ok());
      put('frTotal', Z, Z === undefined ? miss('Weight or rate') : ch.totalFreight !== undefined ? vs('Total freight', Z, ch.totalFreight) : ok('X × W'));
      ([['clrRate', 'clearanceRate'], ['clr', 'clearance'], ['fuel', 'fuel'], ['dem', 'demurrage'], ['hand', 'handling']] as const)
        .forEach(([k, f]) => put(k, ch[f], ok(ch[f] === undefined ? 'Not billed (treated as 0)' : 'OK')));
      const G = Z === undefined ? undefined : r2(Z * 0.18);
      put('gst', G, G === undefined ? miss('Total freight') : ch.gst === undefined ? ok('Freight × 18% (GST not read from CHA invoice)') : vs('GST on freight', G, ch.gst));
      const AG = ch.taxInvoiceTotal === undefined ? undefined : r2(ch.taxInvoiceTotal + (ch.debitNoteTotal ?? 0));
      const exp = Z === undefined ? undefined : r2(V + Z + sum([ch.clearance, ch.fuel, ch.demurrage, ch.handling]) + (G ?? 0));
      put('invAmt', AG, AG === undefined ? miss('CHA invoice total') : exp === undefined ? miss('Freight data needed to compute expected amount')
        : Math.abs(AG - exp) <= 1 ? ok(`Matches calculated amount ${exp}`)
        : bad(`CHA invoice total ${AG} does not match calculated ${exp} (V+Z+clearance+fuel+demurrage+handling+GST); difference ${r2(AG - exp)}`));
      const TDS = Z === undefined ? undefined : r2(Z * 0.01);
      put('tds', TDS, TDS === undefined ? miss('Total freight') : ok('Z × 1%'));
      const ready = AG !== undefined && TDS !== undefined;
      put('net', ready ? r2(AG - TDS) : undefined, ready ? ok('AG − AH') : miss('Invoice amount or TDS'));
    }
    const errs = Object.entries(c).filter(([, x]) => x.s !== 'OK').map(([k, x]) => `${COLS.find(z => z[0] === k)![2]}: ${x.m}`);
    v.remark = errs.length ? errs.join(' | ') : 'All checks passed';
    c.remark = errs.length ? bad(`${errs.length} issue(s)`) : ok();
    return { v, c };
  });
  return { rows, warnings };
}
