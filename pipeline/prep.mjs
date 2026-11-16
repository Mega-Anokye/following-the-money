// Part of "Following the money" (MoMo + IC Securities reconciliation). Personal identifiers are placeholders:
// set STATEMENTS_DIR, OWN_ACCT and EMPLOYER_ACCT (environment variables) and the OWN_BANK_REF / HOLDER patterns before running on your own statements.
// Builds model tables (private + anonymised/scaled public) and prints deep-dive figures with sanity checks.
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import fs from 'fs';
import path from 'path';
const LOW_BALANCE = +(process.env.LOW_BALANCE || 0); // low-balance threshold, in real GHS

const OUT = process.argv[2];               // folder that receives Data (private) and Data_Public
const SCALE = Number(process.env.PUBLIC_SCALE || 1);                      // public version: every GHS amount x this (not disclosed in report)
const {all} = JSON.parse(fs.readFileSync('combined.json', 'utf8'));
const {ic} = JSON.parse(fs.readFileSync('ic_annot.json', 'utf8'));
const fails = []; const ok = (c, m) => { if (!c) { fails.push(m); console.log('FAIL', m); } };
const r2 = x => Math.round(x * 100) / 100;
const sum = a => r2(a.reduce((s, x) => s + x, 0));

// ---------------- MoMo categorisation ----------------
const OWN = (process.env.OWN_ACCT || 'YOUR_WALLET_ACCT');
const workRef = new RegExp(process.env.WORK_REF_PATTERN || 'WORK_REFERENCE', 'i');
// Account numbers are stable across both statement layouts (merchant codes and names are not), so key on them.
const ACCT = {IC: '65077476', PENSION: ['51345771', '51345789'], YELLOSAVE: '63788210', AIRTIME: '39011161', EMPLOYER: (process.env.EMPLOYER_ACCT || 'EMPLOYER_ACCT')};
// wallets of the other networks' push services, learned from older rows (where the merchant code still said mmipush)
const PUSH_ACCTS = new Set(all.filter(t => t.dir === 'Out' && /mmipush/.test(t.ova)).map(t => t.toAcct));
function category(t) {
  const ref = t.ref || '', from = t.fromName || '', ova = t.ova || '';
  if (t.dir === 'In') {
    if (t.fromAcct === ACCT.EMPLOYER || /MULTIPRO/i.test(from) || (/interpush/i.test(from) && /multipro/i.test(ref)) || workRef.test(ref)) return 'Work funds (employer)';
    if (t.fromAcct === ACCT.IC || /ICS3/.test(from) || ova === 'ICS3.sp' || /^tradelive$/i.test(ref.trim())) return 'IC payouts';
    if (/OWN_BANK_REF|ACCOUNT HOLDER/.test(ref) || /BKMBPULL/i.test(ova)) return 'From bank accounts';
    if (t.type === 'CASH_IN') return 'Cash deposits';
    if (t.type === 'ADJUSTMENT' || /momointerest|ysavedep|NS PAYOUTS|OPSWALLET/i.test(from + ' ' + ova)) return 'Interest, refunds & savings payouts';
    if (/INTEROPERABILITY|MobileApp|Interpush|OVA|^mmipull/i.test(from) || /^mmipull/i.test(ova)) return 'From other networks & banks';
    return 'From people';
  }
  if (ova === 'ICS3.sp' || t.toAcct === ACCT.IC) return 'IC savings & investments';
  if (/^uptmyown/.test(ova) || ACCT.PENSION.includes(t.toAcct)) return 'Pension (United Pensions)';
  if (/yelloAD/.test(ova) || t.toAcct === ACCT.YELLOSAVE) return 'YelloSave savings';
  if (workRef.test(ref)) return 'Work payouts';
  if (ova === 'MTNONLINEAIRTIMEVENDOR' || t.toAcct === ACCT.AIRTIME) return 'Airtime & data';
  if (t.type === 'CASH_OUT') return 'Cash withdrawals';
  if (/mmipush/.test(ova) || PUSH_ACCTS.has(t.toAcct)) return 'Transfers to other networks';
  if (t.type === 'TRANSFER' || t.type === 'PAYMENT_SEND') return 'Transfers to people';
  return 'Merchants & bills';
}
const INST = {'IC savings & investments': 'IC Securities', 'IC payouts': 'IC Securities', 'Airtime & data': 'MTN airtime',
  'Pension (United Pensions)': 'United Pensions', 'YelloSave savings': 'YelloSave'};
// service accounts identified from their merchant codes in the older statements
const KNOWN = {'58138232': 'OPS personal wallet', '47812298': 'YelloSave payouts', '48878389': 'MTN MoMo interest'};
const last4 = a => '••' + String(a || '').slice(-4);
const counterpartyOf = t => {
  if (KNOWN[t.dir === 'In' ? t.fromAcct : t.toAcct]) return KNOWN[t.dir === 'In' ? t.fromAcct : t.toAcct];
  if (t.dir === 'In' && /BKMBPULL/i.test(t.ova || '')) return 'Access Bank (mobile banking)';
  // newer layout prints the sender's (first) name in the TO NAME column of incoming rows
  if (t.dir === 'In' && !t.fromName && t.fmt === 'new' && t.toName && !/^HOLDER/i.test(t.toName)) return `${t.toName.trim()} (first name only)`;
  if (t.dir === 'In' && !t.fromName && !(t.ref && t.ref.length > 2) && !/OWN_BANK_REF|ACCOUNT HOLDER/.test(t.ref || '') && !/BKMBPULL/i.test(t.ova)) return `Unnamed wallet ${last4(t.fromAcct)}`;
  if (t.dir === 'In') return (t.fromName && !/^(mmipull|\?)$/i.test(t.fromName) ? t.fromName
    : (/OWN_BANK_REF|ACCOUNT HOLDER/.test(t.ref) || /BKMBPULL/i.test(t.ova) ? 'Own bank account' : (/^[A-Z][^,]+,/.test(t.ref || '') ? t.ref.split(',')[0] : (t.ref && t.ref.length > 2 ? t.ref : `Account ${t.fromAcct}`)))).trim();
  if (/mmipush/.test(t.ova) || PUSH_ACCTS.has(t.toAcct)) return /^[^,]+,[0-9]+/.test(t.ref || '') ? t.ref.split(',')[0].trim() : 'Other network (no name on statement)';
  if (t.toName && t.toName.trim() && t.toName !== '0') return t.toName.trim();
  return ((t.ref || '').replace(/IlI.*/, '').trim().slice(0, 40)) || `Unnamed wallet ${last4(t.toAcct)}`;
};
const momo = all.map((t, i) => {
  const change = r2(t.ba - t.bb);
  const hidden = t.dir === 'Out' ? Math.max(0, r2((t.bb - t.ba) - t.amount)) : 0;
  const unlisted = i === 0 ? 0 : r2(t.bb - all[i - 1].ba);
  const d = new Date(t.date);
  return {Seq: i + 1, TxnID: t.fid, DateTime: d.toISOString().slice(0, 19), Date: d.toISOString().slice(0, 10), Direction: t.dir, Type: t.type,
    Category: category(t), Counterparty: INST[category(t)] || counterpartyOf(t), Amount: t.amount, FeeShown: r2(t.fees + t.elevy), ChargesPaid: hidden,
    BalanceBefore: t.bb, BalanceAfter: t.ba, BalanceChange: change, UnlistedMovement: unlisted, Statement: t.src.replace('.pdf', ''), Layout: t.fmt === 'new' ? 'New (from Jun 2026)' : 'Old', BalanceGap: Math.abs(unlisted) >= 0.005 ? 'Gap before this row' : 'No', Charged: hidden > 0 ? 'Yes' : 'No'};
});
// sanity: category totals add back to in/out totals; every row categorised; balance identity
ok(momo.every(m => m.Category), 'uncategorised MoMo rows');
const inTot = sum(momo.filter(m => m.Direction === 'In').map(m => m.Amount)), outTot = sum(momo.filter(m => m.Direction === 'Out').map(m => m.Amount));
const catSum = d => sum(Object.values(momo.filter(m => m.Direction === d).reduce((a, m) => (a[m.Category] = (a[m.Category] || 0) + m.Amount, a), {})));
ok(Math.abs(catSum('In') - inTot) < 0.01 && Math.abs(catSum('Out') - outTot) < 0.01, 'category totals != direction totals');
const opening = momo[0].BalanceBefore, closing = momo.at(-1).BalanceAfter;
const totalChange = sum(momo.map(m => m.BalanceChange)), totalUnlisted = sum(momo.map(m => m.UnlistedMovement));
ok(Math.abs(r2(opening + totalChange + totalUnlisted) - closing) < 0.01, 'opening + changes + unlisted != closing');
ok(Math.abs(r2(inTot - outTot - sum(momo.map(m => m.ChargesPaid))) - totalChange) < 0.01, 'in - out - charges paid != total change');
ok(momo.filter(m => m.Direction === 'Out').every(m => r2(m.BalanceBefore - m.BalanceAfter - m.Amount) >= -0.005), 'an outgoing row fell by less than its amount');

// ---------------- IC: trades + holdings ----------------
const trade = /^(Bought|Sold) ([\d,]+) (\w+) at ([\d.,]+) for a consideration of ([\d,.]+) and total charges(?: of)? ([\d,.]+)/;
const icRows = ic.map(r => {
  const m = r.desc.replace(/\s+/g, ' ').match(trade);
  const o = {Account: r.section, No: r.no, Date: r.date, Description: r.desc.replace(/\s+/g, ' '), Credit: r.credit, Debit: r.debit,
    Type: r.kind, MatchStatus: r.status, MoMoTxnID: r.momo?.fid ?? '', LagDays: r.momo ? r.lag : '',
    Action: '', Ticker: '', Quantity: '', Price: '', Consideration: '', TradeCharges: ''};
  if (m) {
    const n = s => +s.replace(/,/g, '');
    Object.assign(o, {Action: m[1] === 'Bought' ? 'Buy' : 'Sell', Ticker: m[3], Quantity: n(m[2]), Price: n(m[4]), Consideration: n(m[5]), TradeCharges: n(m[6])});
    const cash = m[1] === 'Bought' ? r.debit : r.credit;
    ok(Math.abs(r2(m[1] === 'Bought' ? o.Consideration + o.TradeCharges : o.Consideration - o.TradeCharges) - cash) < 0.011, `trade cash mismatch IC ${r.section} ${r.no}`);
    o.Type = m[1] === 'Bought' ? 'Share purchase' : 'Share sale';
  } else if (/Kasapreko IPO/.test(r.desc)) o.Type = /Refund/.test(r.desc) ? 'IPO refund' : 'IPO subscription';
  else if (/^Deposited/.test(r.desc)) { o.Type = 'IPO shares received'; const q = r.desc.match(/Deposited ([\d,]+) of (\w+)/); o.Ticker = q[2]; o.Quantity = +q[1].replace(/,/g, ''); }
  return o;
});
ok(icRows.filter(r => r.Type === 'Trade / other inside IC').length === 0, 'unclassified IC rows: ' + icRows.filter(r => r.Type === 'Trade / other inside IC').map(r => r.No).join(','));

const F = './statements/IC-Wealth-account-statement-2026-01-01-to-2026-11-16.pdf';
const doc = await pdfjs.getDocument({data: new Uint8Array(fs.readFileSync(F)), verbosity: 0}).promise;
const items = (await (await doc.getPage(1)).getTextContent()).items.filter(i => i.str.trim()).map(i => ({s: i.str.trim(), x: i.transform[4], y: Math.round(i.transform[5])}));
const lines = {}; items.forEach(i => (lines[i.y] ??= []).push(i));
const holdings = [];
for (const L of Object.values(lines)) {
  const t = L.sort((a, b) => a.x - b.x).map(i => i.s);
  if (t.length === 6 && /\(TradeLive Portfolio\)/.test(t[0])) {
    const n = s => +s.replace(/,/g, '');
    holdings.push({Asset: t[0].replace(' (TradeLive Portfolio)', ''), Class: 'Equities', ISIN: t[1], Quantity: n(t[2]), AvgCost: n(t[3]), MarketPrice: n(t[4]), MarketValue: n(t[5])});
  }
}
holdings.push({Asset: 'IC Liquidity Fund', Class: 'ICLF', ISIN: '', Quantity: '', AvgCost: '', MarketPrice: +(process.env.ICLF_PRICE || 0), MarketValue: +(process.env.ICLF_VALUE || 0)});
holdings.push({Asset: 'Cash (GHS)', Class: 'Cash', ISIN: '', Quantity: '', AvgCost: '', MarketPrice: '', MarketValue: +(process.env.CASH_VALUE || 0)});
holdings.forEach(h => { if (h.Class === 'Equities') { h.CostBasis = r2(h.Quantity * h.AvgCost); h.UnrealisedGain = r2(h.MarketValue - h.CostBasis);
  ok(Math.abs(r2(h.Quantity * h.MarketPrice) - h.MarketValue) < 0.02, `holding value ${h.Asset}`); } else { h.CostBasis = ''; h.UnrealisedGain = ''; } });
const eqMV = sum(holdings.filter(h => h.Class === 'Equities').map(h => h.MarketValue));
ok(Math.abs(eqMV - +(process.env.IC_EQUITY_TOTAL || 0)) < 0.01, `equities total ${eqMV} vs the statement equity total on statement`);
ok(Math.abs(sum(holdings.map(h => h.MarketValue)) - +(process.env.IC_PORTFOLIO_TOTAL || 0)) < 0.01, 'portfolio total vs the statement portfolio total');
// share counts: 2026 buys - sells + IPO deposit vs holdings (difference = shares held before 2026)
const net = {}; icRows.filter(r => r.Ticker).forEach(r => net[r.Ticker] = (net[r.Ticker] || 0) + (r.Action === 'Sell' ? -r.Quantity : r.Quantity));
const qtyCheck = holdings.filter(h => h.Class === 'Equities').map(h => {
  const k = h.Asset === 'KASA' ? (net.KASA || 0) + (net.KPLC || 0) : (net[h.Asset] || 0);
  return {Asset: h.Asset, Held: h.Quantity, Net2026: k, Before2026: h.Quantity - k};
});
const soldOut = Object.entries(net).filter(([k, q]) => !holdings.some(h => h.Asset === k) && k !== 'KPLC').map(([k, q]) => `${k}:${q}`);

// ---------------- write CSVs (private + public) ----------------
const csv = (rows, cols) => [cols.join(','), ...rows.map(r => cols.map(c => { const v = r[c] ?? ''; const s = String(v);
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(','))].join('\r\n');
const MCOLS = ['Seq', 'TxnID', 'DateTime', 'Date', 'Direction', 'Type', 'Category', 'Counterparty', 'Amount', 'FeeShown', 'ChargesPaid', 'BalanceBefore', 'BalanceAfter', 'BalanceChange', 'UnlistedMovement', 'Statement', 'Layout', 'BalanceGap', 'Charged'];
const ICOLS = ['Account', 'No', 'Date', 'Description', 'Credit', 'Debit', 'Type', 'MatchStatus', 'MoMoTxnID', 'LagDays', 'Action', 'Ticker', 'Quantity', 'Price', 'Consideration', 'TradeCharges'];
const HCOLS = ['Asset', 'Class', 'ISIN', 'Quantity', 'AvgCost', 'MarketPrice', 'MarketValue', 'CostBasis', 'UnrealisedGain'];
const write = (dir, m, i, h) => { fs.mkdirSync(dir, {recursive: true});
  fs.writeFileSync(path.join(dir, 'momo_transactions.csv'), csv(m, MCOLS)); fs.writeFileSync(path.join(dir, 'ic_transactions.csv'), csv(i, ICOLS)); fs.writeFileSync(path.join(dir, 'ic_holdings.csv'), csv(h, HCOLS)); };
write(path.join(OUT, 'Data'), momo, icRows, holdings);

// public: people -> Contact NNN, institutions -> generic, txn IDs re-keyed, money x SCALE
const isPerson = (m) => /^(From people|Transfers to people|Cash withdrawals|Cash deposits|Work payouts|Transfers to other networks)$/.test(m.Category);
const aliases = new Map(); const alias = n => { if (!aliases.has(n)) aliases.set(n, `Contact ${String(aliases.size + 1).padStart(3, '0')}`); return aliases.get(n); };
const generic = {'Work funds (employer)': 'Employer', 'From bank accounts': 'Bank account', 'IC payouts': 'IC Securities', 'IC savings & investments': 'IC Securities',
  'Airtime & data': 'MTN airtime', 'Pension (United Pensions)': 'Pension provider', 'YelloSave savings': 'YelloSave', 'Interest, refunds & savings payouts': 'MTN / savings provider', 'From other networks & banks': 'Other network / bank'};
const s = x => (x === '' || x === null || x === undefined) ? x : r2(x * SCALE);
const idMap = new Map(); const pid = id => { if (!idMap.has(id)) idMap.set(id, 'T' + String(idMap.size + 1).padStart(5, '0')); return idMap.get(id); };
const momoPub = momo.map(m => ({...m, TxnID: pid(m.TxnID), Counterparty: generic[m.Category] || (isPerson(m) ? alias(m.Counterparty) : (m.Category === 'Merchants & bills' ? 'Merchant ' + alias(m.Counterparty).slice(8) : alias(m.Counterparty))),
  Amount: s(m.Amount), FeeShown: s(m.FeeShown), ChargesPaid: s(m.ChargesPaid), BalanceBefore: s(m.BalanceBefore), BalanceAfter: s(m.BalanceAfter),
  BalanceChange: r2(s(m.BalanceAfter) - s(m.BalanceBefore)), UnlistedMovement: s(m.UnlistedMovement), Statement: m.Statement.replace(/.*\((\d)\)/, 'Statement $1')}));
// public balances: scale each real balance directly (so none can go negative or drift), then derive amount and
// charge from the scaled balances so every identity still holds exactly. Rounding moves an amount by at most GHS <amount>.
{ momoPub.forEach((m, i) => { const real = momo[i];
    m.BalanceBefore = s(real.BalanceBefore); m.BalanceAfter = s(real.BalanceAfter);
    m.BalanceChange = r2(m.BalanceAfter - m.BalanceBefore);
    m.UnlistedMovement = i === 0 ? 0 : r2(m.BalanceBefore - momoPub[i - 1].BalanceAfter);
    if (m.Direction === 'In') { m.Amount = m.BalanceChange; m.ChargesPaid = 0; }
    else { let chg = r2(-m.BalanceChange - m.Amount); if (chg < 0 || real.ChargesPaid === 0) { m.Amount = r2(-m.BalanceChange); chg = 0; } m.ChargesPaid = chg; } });
  ok(momoPub.every(m => m.BalanceBefore >= 0 && m.BalanceAfter >= 0), 'public: negative balance');
  ok(momoPub.filter(m => Math.abs(m.UnlistedMovement) >= 0.005).length === momo.filter(m => Math.abs(m.UnlistedMovement) >= 0.005).length, 'public: gap count differs from private');
  ok(Math.abs(r2(momoPub[0].BalanceBefore + sum(momoPub.map(m => m.BalanceChange + m.UnlistedMovement)) - momoPub.at(-1).BalanceAfter)) < 0.005, 'public balance chain');
  ok(momoPub.every((m, i) => i === 0 || Math.abs(r2(m.BalanceBefore - momoPub[i - 1].BalanceAfter - m.UnlistedMovement)) < 0.005), 'public row-to-row chain');
  ok(momoPub.every(m => m.Direction === 'Out' || Math.abs(m.BalanceChange - m.Amount) < 0.005), 'public: inflows must change balance by the amount'); }
const icPub = icRows.map(r => ({...r, Description: r.Description.replace(/Account_Holder/g, 'Client').replace(/[\d,]+\.\d\d/g, v => s(+v.replace(/,/g, '')).toFixed(2)),
  Credit: s(r.Credit), Debit: s(r.Debit), MoMoTxnID: r.MoMoTxnID ? pid(r.MoMoTxnID) : '', Quantity: r.Quantity === '' ? '' : Math.round(r.Quantity * SCALE),
  Consideration: s(r.Consideration), TradeCharges: s(r.TradeCharges)}));
const holdPub = holdings.map(h => ({...h, Quantity: h.Quantity === '' ? '' : Math.round(h.Quantity * SCALE), MarketValue: s(h.MarketValue), CostBasis: s(h.CostBasis), UnrealisedGain: s(h.UnrealisedGain)}));
write(path.join(OUT, 'Data_Public'), momoPub, icPub, holdPub);
// public sanity: no real names / phone numbers / IDs leak
const pubText = ['momo_transactions.csv', 'ic_transactions.csv', 'ic_holdings.csv'].map(f => fs.readFileSync(path.join(OUT, 'Data_Public', f), 'utf8')).join('\n');
const leaks = [...new Set(momo.filter(m => isPerson(m) || m.Category === 'From people').map(m => m.Counterparty).filter(n => n.length > 3 && !Object.values(generic).includes(n) && pubText.includes(n)))];
ok(leaks.length === 0, 'names leaked into public data: ' + leaks.slice(0, 5).join(' | '));
ok(!/2335\d{8}|YOUR_WALLET_ACCT|HOLDER_SURNAME|HOLDER_FIRSTNAME/i.test(pubText), 'phone/account/own name leaked into public data');
ok(!momo.some(m => pubText.includes(m.TxnID)), 'real MoMo transaction ID leaked');

// ---------------- deep-dive figures ----------------
const by = (rows, key, val) => Object.entries(rows.reduce((a, r) => (a[key(r)] = (a[key(r)] || 0) + val(r), a), {})).map(([k, v]) => [k, r2(v)]).sort((a, b) => b[1] - a[1]);
const months = by(momo, m => m.Date.slice(0, 7), m => 0).map(x => x[0]).sort();
const ins = momo.filter(m => m.Direction === 'In'), outs = momo.filter(m => m.Direction === 'Out');
const work = sum(ins.filter(m => m.Category === 'Work funds (employer)').map(m => m.Amount)), workOut = sum(outs.filter(m => m.Category === 'Work payouts').map(m => m.Amount));
const insight = {
  period: `${momo[0].Date} to ${momo.at(-1).Date}`, txns: momo.length, months: months.length, moneyIn: inTot, moneyOut: outTot, opening, closing,
  feeShown: sum(momo.map(m => m.FeeShown)), chargesPaid: sum(momo.map(m => m.ChargesPaid)), chargedRows: momo.filter(m => m.ChargesPaid > 0).length, chargesByCat: by(momo.filter(m => m.ChargesPaid > 0), m => m.Category, m => m.ChargesPaid),
  unlisted: totalUnlisted,
  breaks: momo.filter(m => Math.abs(m.UnlistedMovement) >= 0.005).map(m => ({date: m.Date, gap: m.UnlistedMovement, layout: m.Layout})),
  oldPeriod: {txns: momo.filter(m => m.Layout === 'Old').length, to: momo.filter(m => m.Layout === 'Old').at(-1).Date},
  newPeriod: {txns: momo.filter(m => m.Layout !== 'Old').length, from: momo.find(m => m.Layout !== 'Old').Date,
    feeShown: sum(momo.filter(m => m.Layout !== 'Old').map(m => m.FeeShown)), chargesPaid: sum(momo.filter(m => m.Layout !== 'Old').map(m => m.ChargesPaid)),
    daysWithRows: new Set(momo.filter(m => m.Layout !== 'Old').map(m => m.Date)).size},
  oldCharges: {feeShown: sum(momo.filter(m => m.Layout === 'Old').map(m => m.FeeShown)), chargesPaid: sum(momo.filter(m => m.Layout === 'Old').map(m => m.ChargesPaid)),
    rows: momo.filter(m => m.Layout === 'Old' && m.ChargesPaid > 0).length},
  savings: {ic: sum(outs.filter(m => m.Category === 'IC savings & investments').map(m => m.Amount)), pension: sum(outs.filter(m => m.Category === 'Pension (United Pensions)').map(m => m.Amount)),
    yello: sum(outs.filter(m => m.Category === 'YelloSave savings').map(m => m.Amount))},
  monthlyCharges: by(momo, m => m.Date.slice(0, 7), m => m.ChargesPaid).sort(),
  inByCat: by(ins, m => m.Category, m => m.Amount), outByCat: by(outs, m => m.Category, m => m.Amount),
  monthly: months.map(mo => [mo, sum(ins.filter(m => m.Date.startsWith(mo)).map(m => m.Amount)), sum(outs.filter(m => m.Date.startsWith(mo)).map(m => m.Amount))]),
  workIn: work, workOut, airtimeN: outs.filter(m => m.Category === 'Airtime & data').length,
  lowBalanceDays: new Set(momo.filter(m => m.BalanceAfter < LOW_BALANCE).map(m => m.Date)).size,
  ic: { contributionsMoMo: sum(icRows.filter(r => r.Type === 'Money in from MoMo').map(r => r.Credit)), contribN: icRows.filter(r => r.Type === 'Money in from MoMo').length,
    payouts: sum(icRows.filter(r => r.Type === 'Money out to you').map(r => r.Debit)), payoutN: icRows.filter(r => r.Type === 'Money out to you').length,
    bought: sum(icRows.filter(r => r.Action === 'Buy').map(r => r.Consideration)), sold: sum(icRows.filter(r => r.Action === 'Sell').map(r => r.Consideration)),
    tradeCharges: sum(icRows.filter(r => r.Action).map(r => r.TradeCharges)), trades: icRows.filter(r => r.Action).length,
    ipoSub: sum(icRows.filter(r => r.Type === 'IPO subscription').map(r => r.Debit)), ipoRefund: sum(icRows.filter(r => r.Type === 'IPO refund').map(r => r.Credit)),
    equities: eqMV, cost: sum(holdings.filter(h => h.Class === 'Equities').map(h => h.CostBasis)), unrealised: sum(holdings.filter(h => h.Class === 'Equities').map(h => h.UnrealisedGain)),
    holdings: holdings.filter(h => h.Class === 'Equities').map(h => [h.Asset, h.MarketValue, h.UnrealisedGain]).sort((a, b) => b[1] - a[1]),
    iclfMonthly: by(icRows.filter(r => r.Account === 'ICLF' && r.Type === 'Money in from MoMo'), r => r.Date.slice(0, 7), r => r.Credit).sort(),
    contribSizes: by(icRows.filter(r => r.Type === 'Money in from MoMo'), r => String(r.Credit), r => 1).slice(0, 6),
    lagAvg: r2(icRows.filter(r => r.LagDays !== '').reduce((a, r) => a + r.LagDays, 0) / icRows.filter(r => r.LagDays !== '').length),
    qtyCheck, soldOut },
};
fs.writeFileSync('insights.json', JSON.stringify(insight, null, 1));
console.log(JSON.stringify(insight, null, 1).slice(0, 6000));
console.log('FAILS:', fails.length, fails);
