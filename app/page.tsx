'use client';
import { useEffect, useState } from 'react';
import { COLS, run, type Row } from '../lib/rules';
import { extractFile } from '../lib/extract';

type Kind = 'ship' | 'boe' | 'cha';
const SLOTS: { k: Kind; t: string; h: string }[] = [
  { k: 'ship', t: '1. Shipper invoice(s)', h: 'Commercial invoice Excel (.xls / .xlsx), one file per invoice' },
  { k: 'boe', t: '2. Bill of Entry (BOE)', h: 'CBE-XIII with OOC (PDF), one file per BOE' },
  { k: 'cha', t: '3. CHA invoice + debit note', h: 'Tax invoice and debit note from the CHA (PDF)' },
];
const SEED = 'JICHENG TRADING CO., LIMITED';
const clean = (o: any) => Object.fromEntries(Object.entries(o || {}).filter(([, v]) => v !== null));
const fmt = (v: any) => (v == null ? '' : typeof v === 'number' ? v.toLocaleString('en-IN') : String(v));

// Keep only CHA-invoice groups where every check of every row passed
function matchedRows(rows: Row[]): Row[] {
  const out: Row[] = []; let cur: Row[] = [];
  const flush = () => { if (cur.length && cur.every(r => Object.values(r.c).every(x => x.s === 'OK'))) out.push(...cur); cur = []; };
  rows.forEach(r => { if (r.c.tDuty !== undefined || r.c.cha !== undefined) flush(); cur.push(r); });
  flush();
  return out.map((r, i) => ({ ...r, v: { ...r.v, sno: i + 1 } }));
}

export default function Home() {
  const [sup, setSup] = useState<string[]>([SEED]);
  const [supplier, setSupplier] = useState('');
  const [custom, setCustom] = useState('');
  const [files, setFiles] = useState<Record<Kind, File[]>>({ ship: [], boe: [], cha: [] });
  const [res, setRes] = useState<{ rows: Row[]; warnings: string[] } | null>(null);
  const [busy, setBusy] = useState('');
  const [errs, setErrs] = useState<string[]>([]);
  const [all, setAll] = useState(false);

  useEffect(() => {
    try { const s = JSON.parse(localStorage.getItem('suppliers') || '[]'); if (Array.isArray(s)) setSup([SEED, ...s.filter((x: string) => x !== SEED)]); } catch {}
  }, []);

  function addSupplier() {
    const n = custom.trim().slice(0, 120); if (!n) return;
    const l = [...new Set([...sup, n])];
    setSup(l); setSupplier(n); setCustom('');
    try { localStorage.setItem('suppliers', JSON.stringify(l.filter(x => x !== SEED))); } catch {}
  }

  async function check() {
    setErrs([]); setRes(null);
    const out: Record<Kind, any[]> = { ship: [], boe: [], cha: [] };
    const problems: string[] = [];
    const jobs = (Object.keys(files) as Kind[]).flatMap(k => files[k].map(f => ({ k, f })));
    let n = 0;
    try {
      for (const { k, f } of jobs) {
        setBusy(`Reading ${f.name} (${++n}/${jobs.length})`);
        await new Promise(r => setTimeout(r, 0)); // let the screen update
        try { out[k].push(...(await extractFile(k, f)).map(clean)); }
        catch (e: any) { problems.push(`${f.name}: ${e?.message || 'could not be read'}`); }
      }
      setErrs(problems);
      if (out.boe.length && out.cha.length) setRes(run(supplier, out.ship, out.boe, out.cha));
      else problems.length || setErrs(['Nothing could be read from the uploaded files']);
    } finally { setBusy(''); }
  }

  async function exportXlsx(rows: Row[], withRemarks: boolean, name: string) {
    try {
      const [{ default: ExcelJS }, { buildWorkbook }] = await Promise.all([import('exceljs'), import('../lib/export')]);
      const tpl = await (await fetch('/template.xlsx')).arrayBuffer();
      const wb = await buildWorkbook(ExcelJS, tpl, rows, withRemarks);
      const buf = await wb.xlsx.writeBuffer();
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })); a.download = name; a.click();
    } catch (e: any) { setErrs(x => [...x, `Excel export failed: ${e?.message || e}`]); }
  }

  const ready = supplier && files.ship.length && files.boe.length && files.cha.length && !busy;
  const issues = res?.rows.flatMap((r, i) => Object.entries(r.c).filter(([k, x]) => k !== 'remark' && (all || x.s !== 'OK')).map(([k, x]) => ({ i, k, x, v: r.v[k] }))) ?? [];
  const bad = issues.filter(x => x.x.s !== 'OK').length;
  const okRows = res ? matchedRows(res.rows) : [];

  return (
    <main>
      <h1>Import Sheet Checker</h1>
      <div className="mu">Validate shipper invoice, BOE and CHA invoice against the import sheet format.</div>

      <div className="card">
        <h2>Step 1: Supplier</h2>
        <div className="row">
          <select value={supplier} onChange={e => setSupplier(e.target.value)} aria-label="Supplier">
            <option value="">Select supplier…</option>
            {sup.map(s => <option key={s}>{s}</option>)}
          </select>
          <span className="mu">or</span>
          <input type="text" placeholder="Add supplier manually" value={custom} maxLength={120} onChange={e => setCustom(e.target.value)} onKeyDown={e => e.key === 'Enter' && addSupplier()} />
          <button className="sec" onClick={addSupplier}>Add</button>
        </div>
      </div>

      <div className="card">
        <h2>Step 2: Upload documents (runs in your browser, files are not uploaded anywhere)</h2>
        <div className="grid">
          {SLOTS.map(s => (
            <div key={s.k}>
              <b>{s.t}</b><div className="mu">{s.h}</div>
              <input type="file" multiple accept={s.k === 'ship' ? '.xls,.xlsx' : '.pdf'} disabled={!supplier}
                onChange={e => setFiles({ ...files, [s.k]: Array.from(e.target.files || []) })} />
              <div className="mu">{files[s.k].length} file(s) selected</div>
            </div>
          ))}
        </div>
        <div className="row" style={{ marginTop: 14 }}>
          <button disabled={!ready} onClick={check}>Run checks</button>
          {busy && <span className="mu">{busy}</span>}
        </div>
        {errs.map(e => <div key={e} className="err">✖ {e}</div>)}
      </div>

      {res && (
        <>
          <div className="card">
            <div className="row" style={{ justifyContent: 'space-between' }}>
              <h2>Matched: {okRows.length} of {res.rows.length} BOE row(s)</h2>
              <div className="row">
                <button disabled={!okRows.length} onClick={() => exportXlsx(okRows, false, 'Import_Sheet_checked.xlsx')}>Download Excel (matched rows)</button>
                <button className="sec" onClick={() => exportXlsx(res.rows, true, 'Import_Sheet_all_rows_with_remarks.xlsx')}>Download all rows + remarks</button>
              </div>
            </div>
            {res.warnings.map(w => <div key={w} className="err">⚠ {w}</div>)}
            {!okRows.length && <div className="mu">No row matched fully. See the remarks below for what needs correction.</div>}
            <div className="scroll"><table>
              <thead><tr>{COLS.map(c => <th key={c[0]} title={`${c[3]}: ${c[4]}`}>{c[1]} · {c[2]}</th>)}</tr></thead>
              <tbody>{okRows.map((r, i) => <tr key={i}>{COLS.map(c => <td key={c[0]} className={r.c[c[0]]?.s} title={r.c[c[0]]?.m}>{c[0] === 'remark' ? (r.c.remark.s === 'OK' ? 'All checks passed' : r.c.remark.m) : fmt(r.v[c[0]])}</td>)}</tr>)}</tbody>
            </table></div>
          </div>

          <div className="card">
            <div className="row" style={{ justifyContent: 'space-between' }}>
              <h2>Attribute-wise remarks ({bad} issue{bad === 1 ? '' : 's'})</h2>
              <label><input type="checkbox" checked={all} onChange={e => setAll(e.target.checked)} /> Show passed checks too</label>
            </div>
            <div className="scroll"><table>
              <thead><tr><th>Row</th><th>Attribute</th><th>Value</th><th>Status</th><th>Remark</th></tr></thead>
              <tbody>{issues.map(({ i, k, x, v }) => { const c = COLS.find(z => z[0] === k)!; return (
                <tr key={`${i}-${k}`}><td>{i + 1}</td><td>{c[1]} · {c[2]}</td><td>{fmt(v)}</td><td className={x.s}>{x.s}</td><td className="m">{x.m}</td></tr>); })}</tbody>
            </table></div>
          </div>
        </>
      )}
    </main>
  );
}
