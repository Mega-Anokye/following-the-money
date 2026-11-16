// Part of "Following the money" (MoMo + IC Securities reconciliation). Personal identifiers are placeholders:
// set STATEMENTS_DIR, OWN_ACCT and EMPLOYER_ACCT (environment variables) and the OWN_BANK_REF / HOLDER patterns before running on your own statements.
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import fs from 'fs';
const F = './statements/IC-Wealth-account-statement-2026-01-01-to-2026-11-16.pdf';
const doc = await pdfjs.getDocument({data: new Uint8Array(fs.readFileSync(F)), verbosity: 0}).promise;
const num = s => +s.replace(/,/g, '');
const rows = [], totals = {};
for (let p = 3; p <= doc.numPages; p++) {
  const items = (await (await doc.getPage(p)).getTextContent()).items.filter(i => i.str.trim())
    .map(i => ({s: i.str.trim(), x: i.transform[4], y: i.transform[5]}));
  const section = items.find(i => /Transactions$/.test(i.s)).s.replace(' Transactions', '');
  const body = items.filter(i => i.y < 600 && i.y > 60);
  const anchors = body.filter(i => i.x < 68 && /^\d+$/.test(i.s)).sort((a, b) => b.y - a.y);
  for (const a of anchors) {
    const line = body.filter(i => Math.abs(i.y - a.y) < 2);
    const date = line.find(i => i.x > 68 && i.x < 100).s;
    const credit = num(line.find(i => i.x > 430 && i.x < 470).s), debit = num(line.find(i => i.x > 500).s);
    const desc = body.filter(i => i.x > 130 && i.x < 420 && i.y <= a.y + 1 && i.y > a.y - 15 && !anchors.some(o => o !== a && Math.abs(o.y - i.y) < 2)).sort((x, y) => y.y - x.y).map(i => i.s).join(' ');
    const [d, m, yr] = date.split('/');
    rows.push({section, no: +a.s, date: `${yr}-${m}-${d}`, desc, credit, debit, page: p});
  }
  // section totals line: two numbers with no '#' anchor on the line
  const tot = body.filter(i => i.x > 420 && /^[\d,]+\.\d\d$/.test(i.s) && !anchors.some(a => Math.abs(a.y - i.y) < 2));
  if (tot.length === 2) totals[section] = {credit: num(tot.sort((a, b) => a.x - b.x)[0].s), debit: num(tot[1].s)};
}
// sanity: numbering continuous, totals match
for (const sec of ['ICLF', 'TradeLive']) {
  const r = rows.filter(x => x.section === sec);
  const gaps = r.filter((x, i) => x.no !== i + 1);
  const c = r.reduce((a, x) => a + x.credit, 0), d = r.reduce((a, x) => a + x.debit, 0);
  console.log(sec, 'rows', r.length, 'numbering gaps', gaps.length, '| credit', c.toFixed(2), 'vs total', totals[sec]?.credit, '| debit', d.toFixed(2), 'vs total', totals[sec]?.debit);
}
fs.writeFileSync('ic.json', JSON.stringify(rows));
