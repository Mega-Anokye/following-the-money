// Part of "Following the money" (MoMo + IC Securities reconciliation). Personal identifiers are placeholders:
// set STATEMENTS_DIR, OWN_ACCT and EMPLOYER_ACCT (environment variables) and the OWN_BANK_REF / HOLDER patterns before running on your own statements.
import ExcelJS from 'exceljs';
import fs from 'fs';
const {all, stmts, chainBreaks, dups} = JSON.parse(fs.readFileSync('combined.json', 'utf8'));
const out = process.argv[2];
const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('Transactions', {views: [{state: 'frozen', ySplit: 1, xSplit: 1}]});
const cols = [
  ['Date & Time', 'date', 26, 'dd-mmm-yyyy hh:mm:ss AM/PM'], ['From Acct', 'fromAcct', 11], ['From Name', 'fromName', 26], ['From No.', 'fromNo', 14], ['Type', 'type', 15],
  ['Amount (GHS)', 'amount', 13, '#,##0.00'], ['Fees', 'fees', 8, '#,##0.00'], ['E-Levy', 'elevy', 8, '#,##0.00'], ['Balance Before', 'bb', 14, '#,##0.00'], ['Balance After', 'ba', 14, '#,##0.00'],
  ['To No.', 'toNo', 14], ['To Name', 'toName', 28], ['To Acct', 'toAcct', 11], ['Transaction ID (F_ID)', 'fid', 16], ['Reference', 'ref', 34], ['OVA', 'ova', 16],
  ['In / Out', 'dir', 8], ['Balance Change', 'net', 14, '#,##0.00;[Red]-#,##0.00'], ['Source Statement', 'src', 26],
  ['Category', 'cat', 26], ['Counterparty', 'cp', 28], ['Charges Paid', 'chg', 12, '#,##0.00'], ['Month', 'mon', 10, 'mmm-yy']];
ws.columns = cols.map(([h, k, w, f]) => ({header: h, key: k, width: w, style: f ? {numFmt: f} : {}}));
// category + counterparty come from the analysis step (same rules as the Power BI model)
const parseCsv = l => { const o = []; let s = '', q = false; for (const ch of l) { if (q) { if (ch === '"') q = false; else s += ch; } else if (ch === '"') q = true; else if (ch === ',') { o.push(s); s = ''; } else s += ch; } o.push(s); return o; };
const mrows = fs.readFileSync('pbi_out/Data/momo_transactions.csv', 'utf8').trim().split('\r\n').map(parseCsv);
const mh = mrows[0]; const byId = new Map(mrows.slice(1).map(r => [r[mh.indexOf('TxnID')], r]));
all.forEach((t, i) => { const r = i + 2; const m = byId.get(t.fid);
  if (!m) throw new Error('no category for ' + t.fid);
  const d = new Date(t.date);
  ws.addRow({...t, date: d, net: {formula: `J${r}-I${r}`, result: Math.round((t.ba - t.bb) * 100) / 100},
    cat: m[mh.indexOf('Category')], cp: m[mh.indexOf('Counterparty')],
    chg: {formula: `IF(Q${r}="Out",MAX(0,ROUND(I${r}-J${r}-F${r},2)),0)`, result: +m[mh.indexOf('ChargesPaid')]},
    mon: new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1))}); });
const N = all.length + 1;
const hr = ws.getRow(1);
hr.font = {bold: true, color: {argb: 'FFFFFFFF'}}; hr.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF1F4E79'}};
hr.alignment = {vertical: 'middle', wrapText: true}; hr.height = 30;
ws.autoFilter = {from: "A1", to: `W${N}`};
const brk = new Set(chainBreaks.map(c => c.cur.fid));
ws.eachRow((row, i) => { if (i > 1 && brk.has(row.getCell(14).value)) row.eachCell(c => { c.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFFE699'}}; }); });

// ---- Checks sheet: PDF-side values (from the script) vs live Excel formulas over Transactions
const c = wb.addWorksheet('Checks');
c.columns = [{width: 44}, {width: 15}, {width: 15}, {width: 12}, {width: 9}, {width: 60}];
const R = 'Transactions!';
c.addRow(['MoMo statement cross-check — "Excel value" recalculates from the Transactions sheet']).font = {bold: true, size: 13};
c.addRow([]);
let rowNo = 3;
const sec = t => {
  c.addRow([t]).font = {bold: true, size: 12, color: {argb: 'FF1F4E79'}}; rowNo++;
  const h = c.addRow(['Check', 'PDF value', 'Excel value', 'Difference', 'Result', 'Note']);
  h.font = {bold: true}; h.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFDDEBF7'}}; rowNo++;
};
const chk = (label, pdf, formula, note = '', tol = 0.005, fmt = '#,##0.00') => {
  const r = rowNo;
  c.addRow([label, pdf, {formula}, {formula: `C${r}-B${r}`}, {formula: `IF(ABS(D${r})<=${tol},"OK","CHECK")`}, note]);
  ['B', 'C', 'D'].forEach(k => { c.getCell(`${k}${r}`).numFmt = fmt; });
  rowNo++;
};
const xd = d => { d = new Date(d); return `(DATE(${d.getUTCFullYear()},${d.getUTCMonth() + 1},${d.getUTCDate()})+TIME(${d.getUTCHours()},${d.getUTCMinutes()},${d.getUTCSeconds()}))`; };
const D = `${R}$A$2:$A$${N}`, DIR = `${R}$Q$2:$Q$${N}`, AMT = `${R}$F$2:$F$${N}`;
for (const s of stmts) {
  sec(`${s.file}  —  ${s.meta.from} to ${s.meta.to}  (${s.pages} pages)`);
  const rng = `${D},">="&${xd(s.first)}-1/864000,${D},"<="&${xd(s.last)}+1/864000`;
  chk('Number of transactions', s.raw, `COUNTIFS(${rng})`, 'PDF value = transaction dates counted on the PDF pages', 0, '#,##0');
  chk('Money in', s.inT, `SUMIFS(${AMT},${DIR},"In",${rng})`);
  chk('Money out', s.outT, `SUMIFS(${AMT},${DIR},"Out",${rng})`);
  chk('Opening balance (before first transaction)', s.open, `SUMIFS(${R}$I$2:$I$${N},${R}$N$2:$N$${N},"${s.firstFid}")`);
  chk('Closing balance (after last transaction)', s.close, `SUMIFS(${R}$J$2:$J$${N},${R}$N$2:$N$${N},"${s.lastFid}")`);
  const r = rowNo;
  const gapNote = s.breaks ? `${s.breaks} break(s) in MTN's balance chain (net ${s.unlisted.toFixed(2)}): see Notes` : 'Balances add up';
  c.addRow(['Opening + balance changes − closing', 0, {formula: `B${r - 2}+SUMIFS(${R}$R$2:$R$${N},${rng})-B${r - 1}`}, {formula: `C${r}-B${r}`}, {formula: `IF(ABS(C${r})<=0.005,"OK","CHECK")`}, gapNote]);
  ['B', 'C', 'D'].forEach(k => { c.getCell(`${k}${r}`).numFmt = '#,##0.00'; });
  rowNo++;
  c.addRow([]); rowNo++;
}
sec('Whole workbook');
const total = stmts.reduce((a, s) => a + s.raw, 0);
chk(`Rows in all ${stmts.length} PDFs (incl. overlaps)`, total, `COUNTA(${R}$N$2:$N$${N})+${dups}`, `${dups} rows appeared in two statements (overlapping dates) and are kept once`, 0, '#,##0');
chk('Unique transactions in Transactions sheet', all.length, `SUMPRODUCT(1/COUNTIF(${R}$N$2:$N$${N},${R}$N$2:$N$${N}))`, 'Each Transaction ID appears exactly once', 0, '#,##0');
chk('Balance breaks between consecutive rows', chainBreaks.length, `SUMPRODUCT(--(ABS(${R}$I$3:$I$${N}-${R}$J$2:$J$${N - 1})>0.005))`, 'Expected: 1 in the older layout (20-May-2026) + 16 in the newer layout (days missing)', 0, '#,##0');
chk('Rows out of date order', 0, `SUMPRODUCT(--(${R}$A$3:$A$${N}<${R}$A$2:$A$${N - 1}))`, '', 0, '#,##0');
c.addRow([]);
c.addRow(['Notes']).font = {bold: true};
for (const b of chainBreaks) c.addRow([`Source gap: the balance after ${b.prev.dateStr} is ${b.prev.ba.toFixed(2)}, but the next transaction (${b.cur.dateStr}) starts from ${b.cur.bb.toFixed(2)}. GHS ${Math.abs(b.gap).toFixed(2)} left the wallet with no transaction listed in MTN's PDF. That row is highlighted yellow on the Transactions sheet.`]);
const hid = all.filter(t => t.fmt === 'old' && t.dir === 'Out' && (t.bb - t.ba) - t.amount - t.fees - t.elevy > 0.005);
const hidSum = hid.reduce((a, t) => a + (t.bb - t.ba) - t.amount - t.fees - t.elevy, 0);
c.addRow([`Older layout (to Jun 2026): on ${hid.length} outgoing rows the balance fell by more than amount + fees + e-levy: GHS ${hidSum.toFixed(2)} of charges not shown. The newer layout shows the fee correctly on every charged row. "Charges Paid" (column V) uses the balances, so it is right for both.`]);
c.addRow([`Newer layout fixes applied: ${(JSON.parse(fs.readFileSync("combined.json", "utf8")).swaps || []).length} incoming rows printed with Balance Before/After swapped were corrected; names (first names only, and "HOLDER" on every incoming row) were replaced by full names learned from the same account numbers in the older statements.`]);
c.addRow(['"In" = money landed in your wallet (To Acct YOUR_WALLET_ACCT); otherwise "Out". Overlapping dates were de-duplicated by Transaction ID, and every duplicate pair was checked to be identical.']);
c.addRow(['20-May-2026 gap (older layout): IC shows only GHS <amount> + GHS <amount> that day, both matched, so the missing GHS <amount> did not go to IC. Ask MTN (100).']);
c.addRow(['Newer layout gaps: whole days are absent (e.g. 20-31 Aug 2026). IC confirms it: 28 IC payments dated after 19-Jun-2026 have no wallet row (see IC vs MoMo). Ask MTN for complete statements for 19-Jun to 16-Nov-2026.']);

// ---- IC Transactions sheet
const {ic, summary} = JSON.parse(fs.readFileSync('ic_annot.json', 'utf8'));
const is = wb.addWorksheet('IC Transactions', {views: [{state: 'frozen', ySplit: 1}]});
is.columns = [
  {header: 'IC Account', key: 'section', width: 11}, {header: 'IC No.', key: 'no', width: 7}, {header: 'Date', key: 'd', width: 12, style: {numFmt: 'dd-mmm-yyyy'}},
  {header: 'Description', key: 'desc', width: 58}, {header: 'Credit (GHS)', key: 'credit', width: 12, style: {numFmt: '#,##0.00'}}, {header: 'Debit (GHS)', key: 'debit', width: 12, style: {numFmt: '#,##0.00'}},
  {header: 'Type', key: 'kind', width: 22}, {header: 'Match with MoMo', key: 'status', width: 36}, {header: 'MoMo Date & Time', key: 'md', width: 26, style: {numFmt: 'dd-mmm-yyyy hh:mm:ss AM/PM'}},
  {header: 'MoMo Transaction ID', key: 'mfid', width: 16}, {header: 'Days IC posted after MoMo', key: 'lag', width: 12}];
const icSorted = [...ic].sort((a, b) => a.section.localeCompare(b.section) || a.no - b.no);
icSorted.forEach(r => is.addRow({...r, d: new Date(r.date + 'T00:00:00Z'), md: r.momo ? new Date(r.momo.date) : null, mfid: r.momo?.fid ?? '', lag: r.momo ? r.lag : null}));
const IN = icSorted.length + 1;
const ih = is.getRow(1); ih.font = {bold: true, color: {argb: 'FFFFFFFF'}}; ih.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF375623'}}; ih.alignment = {vertical: 'middle', wrapText: true}; ih.height = 30;
is.autoFilter = {from: 'A1', to: `K${IN}`};

// ---- IC vs MoMo checks
const k = wb.addWorksheet('IC vs MoMo');
k.columns = [{width: 52}, {width: 15}, {width: 15}, {width: 12}, {width: 9}, {width: 58}];
k.addRow(['IC Wealth statement (01-Jan to 16-Nov-2026) checked against itself and against MoMo']).font = {bold: true, size: 13};
k.addRow(['MoMo statements now run to 16-Nov-2026, so the comparison covers the whole IC statement (01-Jan to 16-Nov-2026).']).font = {italic: true};
k.addRow([]);
let kr = 4;
const ksec = t => { k.addRow([t]).font = {bold: true, size: 12, color: {argb: 'FF375623'}}; kr++; const h = k.addRow(['Check', 'Side A', 'Side B', 'Difference', 'Result', 'Note']); h.font = {bold: true}; h.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFE2EFDA'}}; kr++; };
const kchk = (label, a, b, note = '', fmt = '#,##0.00') => { const r = kr;
  k.addRow([label, typeof a === 'string' ? {formula: a} : a, {formula: b}, {formula: `C${r}-B${r}`}, {formula: `IF(ABS(D${r})<=0.005,"OK","CHECK")`}, note]);
  ['B', 'C', 'D'].forEach(x => { k.getCell(`${x}${r}`).numFmt = fmt; }); kr++; };
const I = `'IC Transactions'!`, A_ = `${I}$A$2:$A$${IN}`, CR = `${I}$E$2:$E$${IN}`, DB = `${I}$F$2:$F$${IN}`, TY = `${I}$G$2:$G$${IN}`, ST = `${I}$H$2:$H$${IN}`, DT = `${I}$C$2:$C$${IN}`;
const per = `${DT},">="&DATE(2026,1,1),${DT},"<="&DATE(2026,11,16)`;
const mper = `${D},">="&DATE(2026,1,1),${D},"<"&DATE(2026,11,17)`;
const icTot = JSON.parse(process.env.IC_PRINTED_TOTALS || '{"ICLF":{"credit":0,"debit":0},"TradeLive":{"credit":0,"debit":0}}');
ksec('IC statement against its own printed totals (A = printed on PDF, B = sum of rows)');
for (const s of ['ICLF', 'TradeLive']) {
  kchk(`${s} — total credits`, icTot[s].credit, `SUMIFS(${CR},${A_},"${s}")`);
  kchk(`${s} — total debits`, icTot[s].debit, `SUMIFS(${DB},${A_},"${s}")`);
  kchk(`${s} — rows (A = last IC No. on PDF)`, Math.max(...ic.filter(r => r.section === s).map(r => r.no)), `COUNTIFS(${A_},"${s}")`, 'No numbers skipped', '#,##0');
}
k.addRow([]); kr++;
ksec('Money between MoMo and IC, 01-Jan to 16-Nov-2026 (A = IC statement or expected, B = MoMo statement)');
const TXN = `${R}$N$2:$N$${N}`, CAT = `${R}$T$2:$T$${N}`, ICJ = `${I}$J$2:$J$${IN}`;
const icCat = `((${CAT}="IC savings & investments")+(${CAT}="IC payouts"))`;
kchk('IC payments matched to a wallet row', `COUNTIFS(${ST},"Matched*")`, `SUMPRODUCT(${icCat}*(COUNTIF(${ICJ},${TXN})>0))`, 'Each matched IC row points to one MoMo transaction ID', 0, '#,##0');
kchk('Value of the matched pairs (GHS)', `SUMIFS(${CR},${ST},"Matched*")+SUMIFS(${DB},${ST},"Matched*")`, `SUMPRODUCT(${icCat}*(COUNTIF(${ICJ},${TXN})>0)*${AMT})`, 'Amounts agree pair by pair');
kchk('IC rows with no wallet row (expected 28)', 28, `COUNTIFS(${ST},"NOT FOUND*")`, 'All dated after 19-Jun-2026: wallet payments missing from the newer MTN statements', 0, '#,##0');
kchk('Wallet IC rows not yet on the IC statement (expected 4)', 4, `SUMPRODUCT(${icCat}*(${D}>=DATE(2026,1,1))*(COUNTIF(${ICJ},${TXN})=0))`, '13-Nov-2026 top-ups: IC had posted nothing after 12-Nov when its statement was run', 0, '#,##0');
kchk('Matched up to 19-Jun-2026 (older layout)', 218, `COUNTIFS(${ST},"Matched*",${DT},"<="&DATE(2026,6,19))`, '218 of 218: nothing missing in the older layout', 0, '#,##0');
kchk('MoMo Transaction IDs used more than once', 0, `SUMPRODUCT((${ICJ}<>"")*(COUNTIF(${ICJ},${ICJ})>1))`, 'Each MoMo row pairs with one IC row', 0, '#,##0');
k.addRow([]);
k.addRow(['Notes']).font = {bold: true};
k.addRow(['Matching: same amount and direction; in date order each IC row takes the earliest unused wallet row it could have come from (IC posts 0–6 days after MoMo, never before).']);
k.addRow(['"Inside IC — no MoMo leg" = share trades, IPO subscription/refund and Liquidity Fund → TradeLive transfers. "Paid to bank" = three Aug-2026 TradeLive payouts by bank transfer (EFTWB), not to MoMo.']);
k.addRow(['20-May-2026: IC shows only GHS <amount> + GHS <amount> that day, both matched. The GHS <amount> missing from the MoMo balance did not go to IC.']);
k.addRow(['The 28 unmatched IC rows (GHS <amount> incl. GHS <amount> on 15-Sep and 2,000 on 27-Aug) all fall in the newer layout: independent proof that those statements leave transactions out.']);

// ======================= SUMMARY SHEETS (deep dive, live formulas) =======================
const INS = JSON.parse(fs.readFileSync('insights.json', 'utf8'));
const TX = 'Transactions!', rng = col => `${TX}$${col}$2:$${col}$${N}`;
const NAVY = 'FF1F3A5F', TEAL = 'FF2A9D8F', RED = 'FFC8553D', GREY = 'FF5B6670';
const title = (sh, t, sub) => { sh.addRow([t]).font = {bold: true, size: 15, color: {argb: NAVY}}; sh.addRow([sub]).font = {italic: true, color: {argb: GREY}}; sh.addRow([]); };
const secRow = (sh, t) => { const r = sh.addRow([t]); r.font = {bold: true, size: 12, color: {argb: NAVY}}; return r; };
const hdrRow = (sh, vals) => { const r = sh.addRow(vals); r.font = {bold: true}; r.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFDDEBF7'}}; return r; };
const fmt = (row, cols, f) => cols.forEach(cn => { row.getCell(cn).numFmt = f; });

// ---- MoMo Summary
const ms = wb.addWorksheet('MoMo Summary', {views: [{showGridLines: false}]});
ms.columns = [{width: 34}, {width: 16}, {width: 12}, {width: 12}, {width: 3}, {width: 30}, {width: 16}, {width: 12}, {width: 12}];
title(ms, 'MTN MoMo wallet — deep dive (15 Jul 2025 – 16 Nov 2026, 6 statements)', 'Every figure below is a live formula over the Transactions sheet.');
secRow(ms, 'Key figures');
const kf = [
  ['Money in', `SUMIFS(${rng('F')},${rng('Q')},"In")`, '#,##0.00'],
  ['Money out', `SUMIFS(${rng('F')},${rng('Q')},"Out")`, '#,##0.00'],
  ['Transactions', `COUNTA(${rng('N')})`, '#,##0'],
  ['Charges you actually paid (from balances)', `SUM(${rng('V')})`, '#,##0.00'],
  ['Fees shown in MTN Fee + E-Levy columns', `SUMIFS(${rng('G')},${rng('Q')},"Out")+SUMIFS(${rng('H')},${rng('Q')},"Out")`, '#,##0.00'],
  ['Transactions with a charge', `COUNTIFS(${rng('V')},">0")`, '#,##0'],
  ['Work money from employer (share of money in)', `SUMIFS(${rng('F')},${rng('T')},"Work funds (employer)")/B5`, '0.0%'],
  ['Net balance gaps in MTN statements (17 breaks)', `SUMPRODUCT(${TX}$I$3:$I$${N}-${TX}$J$2:$J$${N - 1})`, '#,##0.00;[Red]-#,##0.00'],
  ['Opening balance (15 Jul 2025)', `${TX}I2`, '#,##0.00'],
  ['Closing balance (after the last transaction, 13 Nov 2026)', `${TX}J${N}`, '#,##0.00'],
  ['Check: opening + in − out − charges + missing − closing', 'B13+B5-B6-B8+B12-B14', '#,##0.00'],
];
kf.forEach(([l, f, nf]) => { const r = ms.addRow([l, {formula: f}]); r.getCell(2).numFmt = nf; r.getCell(2).font = {bold: true}; });
ms.addRow([]);
const catStart = ms.rowCount + 1;
const r0 = ms.addRow(['Where the money came from', '', '', '', '', 'Where the money went']); r0.font = {bold: true, size: 12, color: {argb: NAVY}};
hdrRow(ms, ['Category', 'GHS', 'Txns', 'Share', '', 'Category', 'GHS', 'Txns', 'Charges paid']);
const inCats = INS.inByCat.map(c => c[0]), outCats = INS.outByCat.map(c => c[0]);
const firstCat = ms.rowCount + 1;
for (let i = 0; i < Math.max(inCats.length, outCats.length); i++) {
  const rr = ms.rowCount + 1; const vals = [];
  if (inCats[i]) vals.push(inCats[i], {formula: `SUMIFS(${rng('F')},${rng('Q')},"In",${rng('T')},A${rr})`}, {formula: `COUNTIFS(${rng('Q')},"In",${rng('T')},A${rr})`}, {formula: `B${rr}/$B$5`}); else vals.push('', '', '', '');
  vals.push('');
  if (outCats[i]) vals.push(outCats[i], {formula: `SUMIFS(${rng('F')},${rng('Q')},"Out",${rng('T')},F${rr})`}, {formula: `COUNTIFS(${rng('Q')},"Out",${rng('T')},F${rr})`}, {formula: `SUMIFS(${rng('V')},${rng('T')},F${rr})`});
  const row = ms.addRow(vals); fmt(row, [2, 7, 9], '#,##0'); fmt(row, [3, 8], '#,##0'); fmt(row, [4], '0%');
}
const lastCat = ms.rowCount;
const tot = ms.addRow(['Total (should equal money in)', {formula: `SUM(B${firstCat}:B${lastCat})`}, {formula: `SUM(C${firstCat}:C${lastCat})`}, '', '', 'Total (should equal money out)', {formula: `SUM(G${firstCat}:G${lastCat})`}, {formula: `SUM(H${firstCat}:H${lastCat})`}, {formula: `SUM(I${firstCat}:I${lastCat})`}]);
tot.font = {bold: true}; fmt(tot, [2, 7, 9], '#,##0');
const catChk = ms.addRow(['Check vs key figures', {formula: `IF(ABS(B${tot.number}-B5)<0.01,"OK","CHECK")`}, '', '', '', '', {formula: `IF(ABS(G${tot.number}-B6)<0.01,"OK","CHECK")`}]);
catChk.font = {italic: true, color: {argb: TEAL}};
ms.addRow([]);
secRow(ms, 'Month by month');
hdrRow(ms, ['Month', 'Money in', 'Money out', 'Net', '', 'Charges paid', 'Transactions', 'Lowest balance']);
const mStart = ms.rowCount + 1;
for (const [mo] of INS.monthly) {
  const rr = ms.rowCount + 1; const [y, m] = mo.split('-').map(Number);
  const row = ms.addRow([new Date(Date.UTC(y, m - 1, 1)), {formula: `SUMIFS(${rng('F')},${rng('Q')},"In",${rng('W')},A${rr})`}, {formula: `SUMIFS(${rng('F')},${rng('Q')},"Out",${rng('W')},A${rr})`},
    {formula: `B${rr}-C${rr}`}, '', {formula: `SUMIFS(${rng('V')},${rng('W')},A${rr})`}, {formula: `COUNTIFS(${rng('W')},A${rr})`}, {formula: `_xlfn.AGGREGATE(15,6,${rng('J')}/(${rng('W')}=A${rr}),1)`}]);
  row.getCell(1).numFmt = 'mmm yyyy'; fmt(row, [2, 3, 6, 8], '#,##0'); row.getCell(4).numFmt = '#,##0;[Red]-#,##0'; row.getCell(7).numFmt = '#,##0';
}
ms.addRow([]);
secRow(ms, 'Top 10 recipients (people & businesses)');
hdrRow(ms, ['Recipient', 'GHS sent', 'Txns', 'Charges paid']);
const excl = new Set(['IC savings & investments', 'Airtime & data', 'Pension (United Pensions)', 'YelloSave savings']);
const rec = {}; mrows.slice(1).forEach(r => { if (r[mh.indexOf('Direction')] === 'Out' && !excl.has(r[mh.indexOf('Category')])) { const n = r[mh.indexOf('Counterparty')]; rec[n] = (rec[n] || 0) + +r[mh.indexOf('Amount')]; } });
Object.entries(rec).sort((a, b) => b[1] - a[1]).slice(0, 10).forEach(([n]) => { const rr = ms.rowCount + 1;
  const row = ms.addRow([n, {formula: `SUMIFS(${rng('F')},${rng('Q')},"Out",${rng('U')},A${rr})`}, {formula: `COUNTIFS(${rng('Q')},"Out",${rng('U')},A${rr})`}, {formula: `SUMIFS(${rng('V')},${rng('U')},A${rr})`}]);
  fmt(row, [2, 4], '#,##0'); });
ms.addRow([]);
secRow(ms, 'What the data says');
const g = x => 'GHS ' + Math.round(x).toLocaleString('en-US');
[`${(100 * INS.workIn / INS.moneyIn).toFixed(0)}% of all money in (${g(INS.workIn)}) was work money from the employer, some of it paid straight on to field teams, so personal spending is well below the totals.`,
 `Older statements (to Jun 2026) hid charges: ${g(INS.oldCharges.chargesPaid)} came off the balance on ${INS.oldCharges.rows} transactions with the Fee column at zero. The newer layout shows every charge correctly. Total charges paid: ${g(INS.chargesPaid)}.`,
 `MTN's newer statements (from Jun 2026) leave days out: 16 balance breaks, and 28 IC payments (GHS <amount>) with no wallet row. Plus GHS <amount> missing on 20 May 2026 in the older layout. Ask MTN for complete statements.`,
 `Savings out of the wallet: IC ${g(INS.savings.ic)}, United Pensions ${g(INS.savings.pension)}, YelloSave ${g(INS.savings.yello)}: ${(100 * (INS.savings.ic + INS.savings.pension + INS.savings.yello) / INS.moneyOut).toFixed(0)}% of all money out.`,
 `The balance ended a day under GHS <amount> on ${INS.lowBalanceDays} days: money rarely stays in the wallet.`].forEach(t => { const r = ms.addRow(['• ' + t]); r.alignment = {wrapText: false}; });

// ---- IC Summary
const is2 = wb.addWorksheet('IC Summary', {views: [{showGridLines: false}]});
is2.columns = [{width: 34}, {width: 12}, {width: 12}, {width: 12}, {width: 14}, {width: 14}, {width: 14}, {width: 12}];
title(is2, 'IC Wealth — deep dive (1 Jan – 16 Nov 2026)', 'Figures use the IC Transactions sheet and the holdings on page 1 of the IC statement.');
const ICT = `'IC Transactions'!`, irng = cl => `${ICT}$${cl}$2:$${cl}$${IN}`;
secRow(is2, 'Key figures');
const hStart = 20; // holdings table starts here (rows fixed below)
const ikf = [
  ['Paid in from MoMo (contributions + share funding)', `SUMIFS(${irng('E')},${irng('G')},"Money in from MoMo")`, '#,##0.00'],
  ['Number of contributions', `COUNTIFS(${irng('G')},"Money in from MoMo")`, '#,##0'],
  ['   of which GHS <amount> top-ups', `COUNTIFS(${irng('G')},"Money in from MoMo",${irng('E')},10)`, '#,##0'],
  ['Paid out to you', `SUMIFS(${irng('F')},${irng('G')},"Money out to you")`, '#,##0.00'],
  ['IC payments matched to a wallet row (whole period)', `COUNTIFS(${irng('H')},"Matched*")/(COUNTIFS(${irng('H')},"Matched*")+COUNTIFS(${irng('H')},"NOT FOUND*"))`, '0.0%'],
  ['   up to 19 Jun 2026 (older MoMo statements)', `COUNTIFS(${irng('H')},"Matched*",${irng('C')},"<="&DATE(2026,6,19))/(COUNTIFS(${irng('H')},"Matched*",${irng('C')},"<="&DATE(2026,6,19))+COUNTIFS(${irng('H')},"NOT FOUND*",${irng('C')},"<="&DATE(2026,6,19)))`, '0.0%'],
  ['Trading charges (from trade descriptions)', INS.ic.tradeCharges, '#,##0.00'],
  ['Trading charges as % of value bought + sold', INS.ic.tradeCharges / (INS.ic.bought + INS.ic.sold), '0.00%'],
  ['Kasapreko IPO: subscribed / refunded', `"${g(INS.ic.ipoSub)} / ${g(INS.ic.ipoRefund)}"`, '@'],
];
ikf.forEach(([l, f, nf]) => { const r = is2.addRow([l, typeof f === 'number' ? f : {formula: f}]); r.getCell(2).numFmt = nf; r.getCell(2).font = {bold: true}; });
while (is2.rowCount < hStart - 2) is2.addRow([]);
secRow(is2, 'Holdings at 16 Nov 2026');
hdrRow(is2, ['Stock', 'Shares', 'Avg cost', 'Price', 'Market value', 'Cost', 'Gain / loss', 'Return']);
const hrows = fs.readFileSync('pbi_out/Data/ic_holdings.csv', 'utf8').trim().split('\r\n').map(parseCsv); const hh = hrows[0];
const eq = hrows.slice(1).filter(r => r[hh.indexOf('Class')] === 'Equities').sort((a, b) => +b[hh.indexOf('MarketValue')] - +a[hh.indexOf('MarketValue')]);
const h1 = is2.rowCount + 1;
eq.forEach(r => { const rr = is2.rowCount + 1;
  const row = is2.addRow([r[0], +r[hh.indexOf('Quantity')], +r[hh.indexOf('AvgCost')], +r[hh.indexOf('MarketPrice')], +r[hh.indexOf('MarketValue')],
    {formula: `ROUND(B${rr}*C${rr},2)`}, {formula: `E${rr}-F${rr}`}, {formula: `G${rr}/F${rr}`}]);
  fmt(row, [3, 4], '0.00'); fmt(row, [5, 6], '#,##0.00'); row.getCell(7).numFmt = '+#,##0.00;[Red]-#,##0.00'; row.getCell(8).numFmt = '+0.0%;[Red]-0.0%'; });
const h2 = is2.rowCount;
const ht = is2.addRow(['Total shares', '', '', '', {formula: `SUM(E${h1}:E${h2})`}, {formula: `SUM(F${h1}:F${h2})`}, {formula: `SUM(G${h1}:G${h2})`}, {formula: `G${h2 + 1}/F${h2 + 1}`}]);
ht.font = {bold: true}; fmt(ht, [5, 6], '#,##0.00'); ht.getCell(7).numFmt = '+#,##0.00;[Red]-#,##0.00'; ht.getCell(8).numFmt = '+0.0%;[Red]-0.0%';
const hc = is2.addRow(['Check: equities on statement = the statement equity total', '', '', '', {formula: `IF(ABS(E${ht.number}-+(process.env.IC_EQUITY_TOTAL || 0))<0.01,"OK","CHECK")`}]); hc.font = {italic: true, color: {argb: TEAL}};
is2.addRow(['Plus IC Liquidity Fund and cash → portfolio value', '', '', '', {formula: `E${ht.number}+${+(process.env.ICLF_VALUE || 0)}+${+(process.env.CASH_VALUE || 0)}`}]).getCell(5).numFmt = '#,##0.00';
is2.addRow([]);
secRow(is2, 'Money into and out of IC by month');
hdrRow(is2, ['Month', 'Paid in', 'Paid out', 'Contributions']);
for (let m = 1; m <= 11; m++) { const rr = is2.rowCount + 1;
  const row = is2.addRow([new Date(Date.UTC(2026, m - 1, 1)),
    {formula: `SUMPRODUCT((${irng('G')}="Money in from MoMo")*(TEXT(${irng('C')},"yyyymm")=TEXT(A${rr},"yyyymm")),${irng('E')})`},
    {formula: `SUMPRODUCT((${irng('G')}="Money out to you")*(TEXT(${irng('C')},"yyyymm")=TEXT(A${rr},"yyyymm")),${irng('F')})`},
    {formula: `SUMPRODUCT((${irng('G')}="Money in from MoMo")*(TEXT(${irng('C')},"yyyymm")=TEXT(A${rr},"yyyymm")))`}]);
  row.getCell(1).numFmt = 'mmm yyyy'; fmt(row, [2, 3], '#,##0'); }
is2.addRow([]);
secRow(is2, 'What the data says');
const best = INS.ic.holdings.slice().sort((a, b) => b[2] - a[2])[0], worst = INS.ic.holdings.slice().sort((a, b) => a[2] - b[2])[0];
[`The savings habit is micro: ${INS.ic.contribSizes[0][1]} of ${INS.ic.contribN} contributions were GHS <amount> top-ups, adding up to ${g(INS.ic.contributionsMoMo)} in 2026.`,
 `Shares cost ${g(INS.ic.cost)} and are worth ${g(INS.ic.equities)}: +${g(INS.ic.unrealised)} (${(100 * INS.ic.unrealised / INS.ic.cost).toFixed(1)}%). Best ${best[0]} (+${g(best[2])}, Kasapreko IPO shares); weakest ${worst[0]} (${g(worst[2])}).`,
 `${INS.ic.trades} trades cost ${g(INS.ic.tradeCharges)} in charges (${(100 * INS.ic.tradeCharges / (INS.ic.bought + INS.ic.sold)).toFixed(1)}% of value traded) — small, frequent buys pay a proportionally high fee.`,
 `Every IC payment to 19 Jun 2026 matches one MoMo transaction (IC records them on average ${INS.ic.lagAvg} days later, never earlier). After that, 28 IC payments have no wallet row: the newer MTN statements are incomplete. The 20 May GHS <amount> did not go to IC.`].forEach(t => is2.addRow(['• ' + t]));

// put the summaries first
wb.worksheets.forEach(sh => { sh.orderNo = {'MoMo Summary': 0, 'IC Summary': 1, 'Transactions': 2, 'IC Transactions': 3, 'Checks': 4, 'IC vs MoMo': 5}[sh.name] ?? 9; });

await wb.xlsx.writeFile(out);
console.log('written', out, all.length, 'hidden charges', hid.length, hidSum.toFixed(2));
