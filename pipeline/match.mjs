// Part of "Following the money" (MoMo + IC Securities reconciliation). Personal identifiers are placeholders:
// set STATEMENTS_DIR, OWN_ACCT and EMPLOYER_ACCT (environment variables) and the OWN_BANK_REF / HOLDER patterns before running on your own statements.
import fs from 'fs';
const ic = JSON.parse(fs.readFileSync('ic.json', 'utf8'));
const {all} = JSON.parse(fs.readFileSync('combined.json', 'utf8'));
const START = '2026-01-01', END = '2026-11-16';
const day = d => new Date(d.slice(0, 10) + 'T00:00:00Z').getTime() / 864e5;
const IC_ACCT = '65077476';
const isIcMomo = t => (t.dir === 'Out' && (t.ova === 'ICS3.sp' || t.toAcct === IC_ACCT)) || (t.dir === 'In' && (t.fromAcct === IC_ACCT || /ICS3/.test(t.fromName) || /^tradelive$/i.test(t.ref.trim())));
const mm = all.filter(isIcMomo).map(t => ({...t, d: new Date(t.date).toISOString().slice(0, 10)})).filter(t => t.d >= START && t.d <= END);

for (const r of ic) {
  r.amt = r.credit || r.debit;
  if (r.section === 'ICLF' && /^Contribution/.test(r.desc)) r.kind = 'Money in from MoMo';
  else if (r.section === 'ICLF' && /^Withdrawal/.test(r.desc)) r.kind = 'Withdrawal';
  else if (r.section === 'TradeLive' && /^Funding for Purchase of shares$/.test(r.desc)) r.kind = 'Money in from MoMo';
  else if (r.section === 'TradeLive' && /From IC Liquidity Fund/.test(r.desc)) r.kind = 'Internal IC transfer';
  else if (r.section === 'TradeLive' && /payout/i.test(r.desc) && /EFTWB/.test(r.desc)) r.kind = 'Money out to bank';
  else if (r.section === 'TradeLive' && /payout/i.test(r.desc)) r.kind = 'Money out to you';
  else r.kind = 'Trade / other inside IC';
}
// ICLF withdrawals that fund TradeLive are internal
const fund = ic.filter(r => r.kind === 'Internal IC transfer');
for (const w of ic.filter(r => r.kind === 'Withdrawal')) {
  const k = fund.findIndex(f => Math.abs(f.credit - w.debit) < 0.005 && Math.abs(day(f.date) - day(w.date)) <= 3 && !f.paired);
  if (k >= 0) { w.kind = 'Internal IC transfer'; fund[k].paired = true; } else w.kind = 'Money out to you';
}
// match external moves to MoMo; pass 1 tight window, pass 2 wider window for late postings
const used = new Set();
const ext = ic.filter(r => /MoMo|you/.test(r.kind) && r.date >= START && r.date <= END).sort((a, b) => a.date.localeCompare(b.date) || a.no - b.no);
// FIFO per (direction, amount): oldest MoMo row pairs with oldest IC row
const groups = {};
for (const r of ext) (groups[(r.kind === 'Money in from MoMo' ? 'Out' : 'In') + '|' + r.amt] ??= {ic: [], mm: []}).ic.push(r);
for (const t of mm) (groups[t.dir + '|' + t.amount] ??= {ic: [], mm: []}).mm.push(t);
for (const [k, g] of Object.entries(groups)) {
  g.mm.sort((a, b) => new Date(a.date) - new Date(b.date));
  if (g.ic.length !== g.mm.length) console.log('count differs for', k, 'IC', g.ic.length, 'MoMo', g.mm.length);
  // in date order, each IC row takes the earliest unused MoMo row that IC could have posted it from:
  // IC posts 0-6 days after MoMo (1 day tolerance either side). Equivalent to FIFO when nothing is missing.
  g.ic.sort((a, b) => a.date.localeCompare(b.date) || a.no - b.no);
  for (const r of g.ic) {
    const t = g.mm.find(t => !used.has(t.fid) && day(r.date) - day(t.d) >= -1 && day(r.date) - day(t.d) <= 7);
    if (!t) continue; used.add(t.fid); r.momo = {fid: t.fid, dateStr: t.dateStr, date: t.date}; r.lag = day(r.date) - day(t.d);
  }
}
const lags = ext.filter(r => r.momo).map(r => r.lag);
console.log('lag range (IC date minus MoMo date):', Math.min(...lags), 'to', Math.max(...lags), '| >7 days:', lags.filter(l => l > 7).length);
for (const r of ic) {
  if (r.kind === 'Internal IC transfer') r.status = 'Inside IC — no MoMo leg';
  else if (r.kind === 'Money out to bank') r.status = 'Paid to bank — no MoMo leg';
  else if (r.kind === 'Trade / other inside IC') r.status = 'Inside IC — no MoMo leg';
  else if (r.date > END) r.status = 'After MoMo statements end';
  else if (r.momo) r.status = r.lag > 4 ? `Matched (IC posted ${r.lag} days later)` : 'Matched';
  else r.status = 'NOT FOUND in MoMo';
}
const unMM = mm.filter(t => !used.has(t.fid));
const cnt = {}; ic.forEach(r => cnt[r.status] = (cnt[r.status] || 0) + 1);
console.log(cnt); console.log('MoMo IC rows', mm.length, 'unmatched MoMo', unMM.length, unMM.map(t => t.dateStr + ' ' + t.amount));
const s = a => +a.reduce((x, y) => x + y, 0).toFixed(2);
const summary = {
  icIn: s(ext.filter(r => r.kind === 'Money in from MoMo').map(r => r.amt)), icInN: ext.filter(r => r.kind === 'Money in from MoMo').length,
  icOut: s(ext.filter(r => r.kind === 'Money out to you').map(r => r.amt)), icOutN: ext.filter(r => r.kind === 'Money out to you').length,
  mmOut: s(mm.filter(t => t.dir === 'Out').map(t => t.amount)), mmOutN: mm.filter(t => t.dir === 'Out').length,
  mmIn: s(mm.filter(t => t.dir === 'In').map(t => t.amount)), mmInN: mm.filter(t => t.dir === 'In').length, unMM: unMM.length,
};
console.log(summary);
fs.writeFileSync('ic_annot.json', JSON.stringify({ic, summary, momoIcFids: mm.map(t => t.fid)}));
