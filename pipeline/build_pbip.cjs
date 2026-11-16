// Part of "Following the money" (MoMo + IC Securities reconciliation). Personal identifiers are placeholders:
// set STATEMENTS_DIR, OWN_ACCT and EMPLOYER_ACCT (environment variables) and the OWN_BANK_REF / HOLDER patterns before running on your own statements.
// Writes a complete Power BI project (semantic model in TMDL + two PBIR report pages).
// usage: node build_pbip.cjs <rootFolder> <name> <dataFolderAbsWithTrailingBackslash> <private|public>
const fs = require('fs'), path = require('path'), crypto = require('crypto');
const LOW_BALANCE = +(process.env.LOW_BALANCE || 0); // low-balance threshold, in real GHS

const [ROOT, NAME, DATA, VERSION] = process.argv.slice(2);
const PUBLIC = VERSION === 'public', SCALE = PUBLIC ? Number(process.env.PUBLIC_SCALE || 1) : 1;
const RPT = path.join(ROOT, `${NAME}.Report`), SM = path.join(ROOT, `${NAME}.SemanticModel`);
const SCH = 'https://developer.microsoft.com/json-schemas/fabric';
const w = (f, o) => { fs.mkdirSync(path.dirname(f), {recursive: true}); fs.writeFileSync(f, typeof o === 'string' ? o : JSON.stringify(o, null, 2)); };
const uid = () => crypto.randomUUID(), id = () => crypto.randomBytes(10).toString('hex');
const I = require('./insights.json');
const ICA = JSON.parse(fs.readFileSync('ic_annot.json', 'utf8')).ic.filter(x => x.status.startsWith('NOT'));
const NF = ICA.length, NF_SUM = ICA.reduce((a, x) => a + x.amt, 0);
const ghs = x => 'GHS ' + (x * SCALE).toLocaleString('en-US', {minimumFractionDigits: 0, maximumFractionDigits: 0});

// ======================= SEMANTIC MODEL =======================
const T = '\t';
const col = (name, type, o = {}) => [`${T}column ${q(name)}`, `${T}${T}dataType: ${type}`,
  o.fmt ? `${T}${T}formatString: ${o.fmt}` : null, o.hidden ? `${T}${T}isHidden` : null, `${T}${T}lineageTag: ${uid()}`,
  `${T}${T}summarizeBy: ${o.sum || 'none'}`, `${T}${T}sourceColumn: ${o.calc ? `[${name}]` : name}`, o.sortBy ? `${T}${T}sortByColumn: ${o.sortBy}` : null,
  o.calc ? `${T}${T}isNameInferred` : null, '', `${T}${T}annotation SummarizationSetBy = Automatic`, ''].filter(x => x !== null).join('\n');
function q(n) { return /^[A-Za-z_][A-Za-z0-9_]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`; }
const mq = (file, types) => [
  'let',
  `    Source = Csv.Document(File.Contents(DataFolder & "${file}"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),`,
  '    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),',
  '    Blanks = Table.TransformColumns(Promoted, {}, each if _ = "" then null else _),',
  `    Typed = Table.TransformColumnTypes(Blanks, {${types.map(([c, t]) => `{"${c}", ${t}}`).join(', ')}}, "en-US")`,
  'in', '    Typed'].map(l => `${T}${T}${T}${T}${l}`).join('\n');
const MT = {string: 'type text', int64: 'Int64.Type', double: 'type number', dateTime: 'type datetime', date: 'type date'};
function table(name, desc, file, cols) {
  const types = cols.map(c => [c[0], c[1] === 'dateTime' && c[2]?.dateOnly ? 'type date' : MT[c[1]]]);
  return [`/// ${desc}`, `table ${q(name)}`, `${T}lineageTag: ${uid()}`, '',
    ...cols.map(c => col(c[0], c[1], c[2] || {})),
    `${T}partition ${q(name)} = m`, `${T}${T}mode: import`, `${T}${T}source =`, mq(file, types), '',
    `${T}annotation PBI_ResultType = Table`, ''].join('\n');
}
const S = 'string', N = 'double', D = 'dateTime', Z = 'int64';
const money = {fmt: '"GHS "#,0.00', sum: 'sum'};
const tables = {
  MoMo: table('MoMo', 'One row per MTN MoMo transaction, 4 statements merged and de-duplicated by transaction ID.', 'momo_transactions.csv', [
    ['Seq', Z, {hidden: true}], ['TxnID', S], ['DateTime', D, {fmt: 'dd mmm yyyy hh:nn'}], ['Date', D, {dateOnly: true, fmt: 'dd mmm yyyy'}],
    ['Direction', S], ['Type', S], ['Category', S], ['Counterparty', S], ['Amount', N, money], ['FeeShown', N, money], ['ChargesPaid', N, money],
    ['BalanceBefore', N, {fmt: '"GHS "#,0.00'}], ['BalanceAfter', N, {fmt: '"GHS "#,0.00'}], ['BalanceChange', N, money], ['UnlistedMovement', N, money], ['Statement', S], ['Layout', S], ['BalanceGap', S], ['Charged', S]]),
  IC: table('IC', 'One row per IC Wealth transaction (ICLF and TradeLive), Jan to Nov 2026, with its matched MoMo transaction.', 'ic_transactions.csv', [
    ['Account', S], ['No', Z], ['Date', D, {dateOnly: true, fmt: 'dd mmm yyyy'}], ['Description', S], ['Credit', N, money], ['Debit', N, money], ['Type', S],
    ['MatchStatus', S], ['MoMoTxnID', S], ['LagDays', Z, {sum: 'none'}], ['Action', S], ['Ticker', S], ['Quantity', Z, {sum: 'sum'}], ['Price', N],
    ['Consideration', N, money], ['TradeCharges', N, money]]),
  Holdings: table('Holdings', 'Portfolio holdings as at 16 Nov 2026 (IC statement page 1).', 'ic_holdings.csv', [
    ['Asset', S], ['Class', S], ['ISIN', S], ['Quantity', Z, {sum: 'sum'}], ['AvgCost', N, {fmt: '#,0.00'}], ['MarketPrice', N, {fmt: '#,0.00'}],
    ['MarketValue', N, money], ['CostBasis', N, money], ['UnrealisedGain', N, money]]),
};
tables.Date = [`/// Calendar for both sources.`, `table Date`, `${T}lineageTag: ${uid()}`, `${T}dataCategory: Time`, '',
  col('Date', 'dateTime', {calc: true, fmt: 'dd mmm yyyy'}).replace(`${T}${T}isNameInferred`, `${T}${T}isKey\n${T}${T}isNameInferred`),
  col('Month', 'string', {calc: true, sortBy: 'MonthKey'}), col('MonthKey', 'int64', {calc: true, hidden: true}), col('Year', 'int64', {calc: true}),
  `${T}partition Date = calculated`, `${T}${T}mode: import`, `${T}${T}source =`,
  `${T}${T}${T}${T}ADDCOLUMNS ( CALENDAR ( DATE ( 2025, 7, 1 ), DATE ( 2026, 11, 30 ) ), "Month", FORMAT ( [Date], "MMM yy" ), "MonthKey", YEAR ( [Date] ) * 100 + MONTH ( [Date] ), "Year", YEAR ( [Date] ) )`, '',
  `${T}annotation PBI_Id = ${id()}`, ''].join('\n');

const M = [];
const ms = (name, expr, fmt, folder, desc) => M.push([desc ? `${T}/// ${desc}` : null, `${T}measure ${q(name)} = ${expr}`, `${T}${T}formatString: ${fmt}`,
  `${T}${T}displayFolder: ${folder}`, `${T}${T}lineageTag: ${uid()}`, ''].filter(Boolean).join('\n'));
const G = '"GHS "#,0', G2 = '"GHS "#,0.00', NUM = '#,0', PCT = '0.0%';
// MoMo
ms('Money In', 'CALCULATE ( SUM ( MoMo[Amount] ), KEEPFILTERS ( MoMo[Direction] = "In" ) )', G, 'MoMo');
ms('Money Out', 'CALCULATE ( SUM ( MoMo[Amount] ), KEEPFILTERS ( MoMo[Direction] = "Out" ) )', G, 'MoMo');
ms('Net Flow', '[Money In] - [Money Out]', G, 'MoMo');
ms('Transactions', 'COUNTROWS ( MoMo ) + 0', NUM, 'MoMo');
ms('Charges Paid', 'SUM ( MoMo[ChargesPaid] ) + 0', G2, 'MoMo', 'Charges that actually left the wallet: balance fall minus amount.');
ms('Fees Shown by MTN', 'SUM ( MoMo[FeeShown] ) + 0', G2, 'MoMo', 'Fee + E-Levy columns on the statement (these were paid by merchants, not deducted from the wallet).');
ms('Charged Transactions', 'CALCULATE ( COUNTROWS ( MoMo ), MoMo[ChargesPaid] > 0 ) + 0', NUM, 'MoMo');
ms('Charges % of Money Out', 'DIVIDE ( [Charges Paid], [Money Out] )', '0.00%', 'MoMo');
ms('Unlisted Movement', 'SUM ( MoMo[UnlistedMovement] ) + 0', G2, 'MoMo', 'Net balance change with no transaction on the statement (gaps in MTN statements).');
ms('Balance Breaks', 'CALCULATE ( COUNTROWS ( MoMo ), MoMo[UnlistedMovement] <> 0 ) + 0', NUM, 'MoMo', 'Rows whose balance-before does not equal the previous balance-after.');
ms('Opening Balance', 'VAR s = MIN ( MoMo[Seq] ) RETURN CALCULATE ( MAX ( MoMo[BalanceBefore] ), MoMo[Seq] = s, REMOVEFILTERS ( MoMo[Direction], MoMo[Category] ) )', G2, 'MoMo');
ms('Closing Balance', 'VAR s = MAX ( MoMo[Seq] ) RETURN CALCULATE ( MAX ( MoMo[BalanceAfter] ), MoMo[Seq] = s, REMOVEFILTERS ( MoMo[Direction], MoMo[Category] ) )', G2, 'MoMo');
ms('Balance Check', '[Opening Balance] + [Money In] - [Money Out] - [Charges Paid] + [Unlisted Movement] - [Closing Balance]', G2, 'Checks', 'Should be 0 with no filters.');
ms('Work Funds In', 'CALCULATE ( [Money In], KEEPFILTERS ( MoMo[Category] = "Work funds (employer)" ) )', G, 'MoMo');
ms('Work Share of Inflows', 'DIVIDE ( [Work Funds In], [Money In] )', PCT, 'MoMo');
ms('Days Below GHS <amount>', `COUNTROWS ( FILTER ( VALUES ( MoMo[Date] ), CALCULATE ( MIN ( MoMo[BalanceAfter] ) ) < ${(LOW_BALANCE * SCALE).toFixed(2)} ) ) + 0`, NUM, 'MoMo', 'Days the balance fell under GHS <amount> (scaled in the public version).');
// IC
ms('IC Paid In from MoMo', 'CALCULATE ( SUM ( IC[Credit] ), KEEPFILTERS ( IC[Type] = "Money in from MoMo" ) )', G, 'IC');
ms('IC Paid Out to You', 'CALCULATE ( SUM ( IC[Debit] ), KEEPFILTERS ( IC[Type] = "Money out to you" ) )', G, 'IC');
ms('IC Contributions', 'CALCULATE ( COUNTROWS ( IC ), KEEPFILTERS ( IC[Type] = "Money in from MoMo" ) ) + 0', NUM, 'IC');
ms('Shares Bought', 'CALCULATE ( SUM ( IC[Consideration] ), KEEPFILTERS ( IC[Action] = "Buy" ) ) + 0', G, 'IC');
ms('Shares Sold', 'CALCULATE ( SUM ( IC[Consideration] ), KEEPFILTERS ( IC[Action] = "Sell" ) ) + 0', G, 'IC');
ms('Trading Charges', 'SUM ( IC[TradeCharges] ) + 0', G2, 'IC');
ms('Trading Cost %', 'DIVIDE ( [Trading Charges], [Shares Bought] + [Shares Sold] )', '0.00%', 'IC');
ms('Trades', 'CALCULATE ( COUNTROWS ( IC ), IC[Action] <> BLANK () ) + 0', NUM, 'IC');
ms('Portfolio Value', 'SUM ( Holdings[MarketValue] ) + 0', G2, 'IC');
ms('Equity Value', 'CALCULATE ( SUM ( Holdings[MarketValue] ), KEEPFILTERS ( Holdings[Class] = "Equities" ) ) + 0', G2, 'IC');
ms('Cost Basis', 'SUM ( Holdings[CostBasis] ) + 0', G2, 'IC');
ms('Unrealised Gain', 'SUM ( Holdings[UnrealisedGain] ) + 0', '"+GHS "#,0.00;"-GHS "#,0.00', 'IC');
ms('Unrealised Return %', 'DIVIDE ( [Unrealised Gain], [Cost Basis] )', '+0.0%;-0.0%', 'IC');
ms('MoMo Rows Listed', 'COUNTROWS ( MoMo ) + 0', NUM, 'Detail');
ms('IC Rows Listed', 'COUNTROWS ( IC ) + 0', NUM, 'Detail');
ms('IC Rows to Match', 'CALCULATE ( COUNTROWS ( IC ), KEEPFILTERS ( IC[Type] IN { "Money in from MoMo", "Money out to you" } ), KEEPFILTERS ( IC[Date] <= DATE ( 2026, 11, 16 ) ) ) + 0', NUM, 'Checks');
ms('IC Rows Matched', 'CALCULATE ( [IC Rows to Match], LEFT ( IC[MatchStatus], 7 ) = "Matched" ) + 0', NUM, 'Checks');
ms('MoMo Match Rate', 'DIVIDE ( [IC Rows Matched], [IC Rows to Match] )', '0.0%', 'Checks');
ms('Avg Posting Lag (days)', 'AVERAGE ( IC[LagDays] )', '0.0', 'Checks');
tables._Measures = [`/// All measures.`, `table _Measures`, `${T}lineageTag: ${uid()}`, '', ...M,
  `${T}column Value`, `${T}${T}dataType: string`, `${T}${T}isHidden`, `${T}${T}lineageTag: ${uid()}`, `${T}${T}summarizeBy: none`, `${T}${T}isNameInferred`, `${T}${T}sourceColumn: [Value]`, '',
  `${T}partition _Measures = calculated`, `${T}${T}mode: import`, `${T}${T}source = ROW ( "Value", BLANK () )`, ''].join('\n');

fs.rmSync(SM, {recursive: true, force: true});
const SD = path.join(SM, 'definition');
w(path.join(SD, 'database.tmdl'), 'database\n\tcompatibilityLevel: 1606\n\n');
w(path.join(SD, 'model.tmdl'), ['model Model', '\tculture: en-US', '\tdefaultPowerBIDataSourceVersion: powerBI_V3', '\tsourceQueryCulture: en-US',
  '\tdataAccessOptions', '\t\tlegacyRedirects', '\t\treturnErrorValuesAsNull', '',
  'annotation __PBI_TimeIntelligenceEnabled = 0', '', 'annotation PBI_QueryOrder = ["DataFolder","MoMo","IC","Holdings"]', '',
  ...Object.keys(tables).map(t => `ref table ${q(t)}`), '', 'ref cultureInfo en-US', ''].join('\n'));
w(path.join(SD, 'expressions.tmdl'), `/// Folder holding momo_transactions.csv, ic_transactions.csv and ic_holdings.csv (keep the trailing backslash)\nexpression DataFolder = "${DATA.replace(/"/g, '""')}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n\tlineageTag: ${uid()}\n\n\tannotation PBI_ResultType = Text\n`);
w(path.join(SD, 'relationships.tmdl'), [
  `relationship ${uid()}`, '\tfromColumn: MoMo.Date', '\ttoColumn: Date.Date', '',
  `relationship ${uid()}`, '\tfromColumn: IC.Date', '\ttoColumn: Date.Date', '',
  `relationship ${uid()}`, '\tfromColumn: IC.Ticker', '\ttoColumn: Holdings.Asset', ''].join('\n'));
w(path.join(SD, 'cultures', 'en-US.tmdl'), 'cultureInfo en-US\n\n\tlinguisticMetadata =\n\t\t\t{\n\t\t\t  "Version": "1.0.0",\n\t\t\t  "Language": "en-US"\n\t\t\t}\n\t\tcontentType: json\n\n');
for (const [n, t] of Object.entries(tables)) w(path.join(SD, 'tables', `${n}.tmdl`), t);

// ======================= REPORT =======================
const INK = '#1F2933', MUTED = '#5B6670', IN_C = '#2A9D8F', OUT_C = '#C8553D', NAVY = '#1F3A5F';
const lit = v => ({expr: {Literal: {Value: v}}});
const str = s => lit(`'${s.replace(/'/g, "''")}'`);
const color = c => ({solid: {color: str(c)}});
const C = (t, c, as) => ({f: {Column: {Expression: {SourceRef: {Entity: t}}, Property: c}}, as});
const Me = (m, as) => ({f: {Measure: {Expression: {SourceRef: {Entity: '_Measures'}}, Property: m}}, as});
const proj = ({f, as}) => { const p = f.Column || f.Measure; const o = {field: f, queryRef: `${p.Expression.SourceRef.Entity}.${p.Property}`, nativeQueryRef: p.Property}; if (as) o.displayName = as; return o; };
function inFilter(t, c, values, negate = false) {
  const cond = {In: {Expressions: [{Column: {Expression: {SourceRef: {Source: 'q'}}, Property: c}}], Values: values.map(v => [{Literal: {Value: `'${v.replace(/'/g, "''")}'`}}])}};
  return {name: id(), field: C(t, c).f, type: 'Categorical', howCreated: 'User',
    filter: {Version: 2, From: [{Name: 'q', Entity: t, Type: 0}], Where: [{Condition: negate ? {Not: {Expression: cond}} : cond}]}};
}
function visual(pos, type, roles, o = {}) {
  const v = {visualType: type};
  if (roles) { v.query = {queryState: Object.fromEntries(Object.entries(roles).map(([r, f]) => [r, {projections: f.map(proj)}]))};
    if (o.sort) v.query.sortDefinition = {sort: [{field: o.sort[0].f, direction: o.sort[1]}], isDefaultSort: true}; }
  v.objects = o.objects || {};
  const vco = {title: o.title ? [{properties: {show: lit('true'), text: str(o.title), fontSize: lit('11D'), bold: lit('true'), fontColor: color(INK)}}] : [{properties: {show: lit('false')}}]};
  if (o.subtitle) vco.subTitle = [{properties: {show: lit('true'), text: str(o.subtitle), fontSize: lit('9D'), fontColor: color(MUTED)}}];
  v.visualContainerObjects = vco;
  const out = {$schema: `${SCH}/item/report/definition/visualContainer/1.0.0/schema.json`, name: id(), position: {x: pos[0], y: pos[1], z: 0, width: pos[2], height: pos[3], tabOrder: 0}, visual: v};
  if (o.filters) out.filterConfig = {filters: o.filters};
  return out;
}
const labels = {labels: [{properties: {show: lit('true'), color: color(INK), fontSize: lit('8D')}}]};
const axes = {categoryAxis: [{properties: {showAxisTitle: lit('false'), labelColor: color(INK), fontSize: lit('8D')}}], valueAxis: [{properties: {showAxisTitle: lit('false'), show: lit('false')}}]};
const kpi = (pos, m, title, sub, c = NAVY) => visual(pos, 'card', {Values: [Me(m)]}, {title, subtitle: sub,
  objects: {categoryLabels: [{properties: {show: lit('false')}}], labels: [{properties: {fontSize: lit('18D'), color: color(c)}}]}});
const bar = (pos, cat, m, title, sub, c, o = {}) => visual(pos, 'clusteredBarChart', Object.assign({Category: [cat], Y: [m]}, o.tooltips ? {Tooltips: o.tooltips} : {}),
  {title, subtitle: sub, sort: [m, 'Descending'], filters: o.filters, objects: Object.assign({}, labels, axes, {dataPoint: [{properties: {fill: color(c)}}]})});
const textbox = (pos, paras) => visual(pos, 'textbox', null, {objects: {general: [{properties: {paragraphs: paras.map(p =>
  typeof p === 'string' ? {textRuns: [{value: p, textStyle: {fontSize: '9pt', color: INK}}]} : p)}}]}});
const header = (title, sub, wd = 740) => textbox([16, 4, wd, 58], [
  {textRuns: [{value: title, textStyle: {fontWeight: 'bold', fontSize: '17pt', color: INK}}]},
  {textRuns: [{value: sub, textStyle: {fontSize: '9.5pt', color: MUTED}}]}]);
// dropdown slicer; same-named groups stay in sync across pages
const slicer = (pos, field, group) => { const s = visual(pos, 'slicer', {Values: [field]}, {objects: {
  data: [{properties: {mode: str('Dropdown')}}], general: [{properties: {selfFilterEnabled: lit('true')}}],
  selection: [{properties: {selectAllCheckboxEnabled: lit('true'), singleSelect: lit('false')}}],
  header: [{properties: {show: lit('true'), fontColor: color(INK), textSize: lit('9D')}}]}});
  s.visual.syncGroup = {groupName: group, fieldChanges: true, filterChanges: true}; return s; };
const slicerBar = list => list.map(([f, g], i) => slicer([764 + i * 168, 8, 160, 54], f, g));
// ---------- navigation: buttons, bookmarks, drill-through (Distributor Profiling pattern) ----------
const CHIP = '#EAF0F7';
const sel = {selector: {id: 'default'}};
function actionButton(pos, label, link) {
  const v = visual(pos, 'actionButton', null, {objects: {
    icon: [{properties: {show: lit('false')}}, Object.assign({properties: {shapeType: str('blank')}}, sel)],
    fill: label ? [{properties: {show: lit('true')}}, Object.assign({properties: {fillColor: color(CHIP), transparency: lit('0D')}}, sel)] : [{properties: {show: lit('false')}}],
    outline: [{properties: {show: lit('false')}}],
    text: label ? [{properties: {show: lit('true')}}, Object.assign({properties: {text: str(label), fontColor: color(INK), fontSize: lit('10D'), bold: lit('true')}}, sel)] : [{properties: {show: lit('false')}}]}});
  const vco = v.visual.visualContainerObjects;
  vco.visualLink = [{properties: Object.assign({show: lit('true')}, link)}];
  vco.border = [{properties: {show: lit('false')}}];
  if (!label) vco.background = [{properties: {show: lit('false')}}];
  return v;
}
const MOMO_DETAIL = 'momo_detail', IC_DETAIL = 'ic_detail', MOMO_PAGE = 'momo', IC_PAGE = 'ic';
// page-level filters on the detail pages that the card bookmarks switch on/off
const DETAIL_FILTERS = {
  [MOMO_DETAIL]: {gap: ['fltGap', 'MoMo', 'BalanceGap'], charged: ['fltCharged', 'MoMo', 'Charged']},
  [IC_DETAIL]: {match: ['fltMatch', 'IC', 'MatchStatus'], action: ['fltAction', 'IC', 'Action']}};
const pageFilter = (page, key, values) => { const [name, t, c] = DETAIL_FILTERS[page][key];
  const f = values ? inFilter(t, c, values) : {field: C(t, c).f, type: 'Categorical', howCreated: 'User'}; return Object.assign(f, {name}); };
const bookmarks = [];
function bookmark(key, displayName, page, filterKey, values) {
  const name = 'bm_' + key;
  const byExpr = Object.keys(DETAIL_FILTERS[page]).map(k => { const f = pageFilter(page, k, k === filterKey ? values : null);
    const o = {name: f.name, type: 'Categorical', expression: f.field, howCreated: 1}; if (f.filter) o.filter = f.filter; return o; });
  bookmarks.push({$schema: `${SCH}/item/report/definition/bookmark/1.0.0/schema.json`, displayName, name,
    explorationState: {version: '1.3', activeSection: page, sections: {[page]: {filters: {byExpr}, visualContainers: {}}}}});
  return name;
}
const BM = {
  momoAll: bookmark('momo_all', 'MoMo detail: all', MOMO_DETAIL),
  gaps: bookmark('momo_gaps', 'MoMo detail: balance gaps', MOMO_DETAIL, 'gap', ['Gap before this row']),
  charged: bookmark('momo_charged', 'MoMo detail: charged transactions', MOMO_DETAIL, 'charged', ['Yes']),
  icAll: bookmark('ic_all', 'IC detail: all', IC_DETAIL),
  unmatched: bookmark('ic_unmatched', 'IC detail: no wallet row', IC_DETAIL, 'match', ['NOT FOUND in MoMo']),
  trades: bookmark('ic_trades', 'IC detail: share trades', IC_DETAIL, 'action', ['Buy', 'Sell'])};
// a card with an invisible button on top that opens the detail page with one filter set
const linkedKpi = (pos, m, title, sub, bm, c, tip) => [kpi(pos, m, title, sub, c),
  actionButton(pos, null, {type: str('Bookmark'), bookmark: str(bm), tooltip: str(tip || 'Click to list these transactions')})];
// "List selected" carries a clicked bar/row into the detail page (greyed out until something is selected)
const listSelected = (pos, page) => actionButton(pos, 'List selected →', {type: str('Drillthrough'), drillthroughSection: str(page),
  tooltip: str('List the transactions behind the selected bar or row'), disabledTooltip: str('Click a bar or row first, then this lists its transactions')});
const openDetail = (pos, bm, label) => actionButton(pos, label, {type: str('Bookmark'), bookmark: str(bm), tooltip: str('Open every transaction (filters on this page carry over)')});
const drillFilter = (t, c) => ({name: `dt_${t}_${c}`, field: C(t, c).f, type: 'Categorical', howCreated: 'Drillthrough'});
const tableText = {columnHeaders: [{properties: {fontSize: lit('8D'), wordWrap: lit('true'), bold: lit('true'), fontColor: color(INK)}}], values: [{properties: {fontSize: lit('8D'), fontColor: color(INK)}}]};
const head = t => ({textRuns: [{value: t, textStyle: {fontWeight: 'bold', fontSize: '10.5pt', color: NAVY}}]});
const bullet = t => ({textRuns: [{value: '• ' + t, textStyle: {fontSize: '9pt', color: INK}}]});
const bulletS = t => ({textRuns: [{value: '• ' + t, textStyle: {fontSize: '8pt', color: INK}}]});
const pct = (a, b) => (100 * a / b).toFixed(0) + '%';
const ic = I.ic;
const who = PUBLIC ? 'Anonymised portfolio version: names masked and all GHS amounts scaled by a hidden factor, so patterns and checks hold.' : 'Private version, real figures.';

const pages = [];
// ---------- Page 1: MTN MoMo ----------
pages.push({name: 'momo', displayName: 'MTN MoMo wallet', visuals: [
  header('MTN MoMo wallet', `15 Jul 2025 – 16 Nov 2026 · ${I.txns.toLocaleString()} transactions, 6 statements, reconciled. ${PUBLIC ? 'Anonymised.' : 'Private.'} Click a bar, then List selected →; or click a red card.`, 470),
  listSelected([494, 10, 130, 24], MOMO_DETAIL), openDetail([494, 38, 130, 24], BM.momoAll, 'All transactions'),
  ...slicerBar([[C('Date', 'Month', 'Month'), 'sync_month'], [C('MoMo', 'Direction', 'In / Out'), 'sync_dir'], [C('MoMo', 'Category', 'Category'), 'sync_cat']]),
  kpi([16, 68, 200, 92], 'Money In', 'Money in', 'Selected period'),
  kpi([224, 68, 200, 92], 'Money Out', 'Money out', 'Selected period'),
  kpi([432, 68, 200, 92], 'Work Share of Inflows', 'Work money passing through', 'Share of inflows from employer'),
  ...linkedKpi([640, 68, 200, 92], 'Charges Paid', 'Charges paid', 'Click to list charged rows', BM.charged, OUT_C, 'List every transaction that carried a charge'),
  ...linkedKpi([848, 68, 200, 92], 'Balance Breaks', 'Gaps in MTN statements', 'Click to list the gaps', BM.gaps, OUT_C, 'List the rows where MTN\'s balance chain breaks'),
  kpi([1056, 68, 208, 92], 'Days Below GHS <amount>', 'Low-balance days', `Balance under ${ghs(LOW_BALANCE)}`),
  visual([16, 168, 620, 262], 'clusteredColumnChart', {Category: [C('Date', 'Month', 'Month')], Y: [Me('Money In', 'In'), Me('Money Out', 'Out')]},
    {title: 'Money in vs money out by month', subtitle: 'In and out move together: the wallet is a pass-through, not a store of money',
      sort: [C('Date', 'Month'), 'Ascending'], objects: Object.assign({}, axes, {legend: [{properties: {show: lit('true'), position: str('TopRight'), labelColor: color(INK)}}],
        dataPoint: [{properties: {fill: color(IN_C)}, selector: {metadata: '_Measures.Money In'}}, {properties: {fill: color(OUT_C)}, selector: {metadata: '_Measures.Money Out'}}]})}),
  bar([644, 168, 620, 262], C('MoMo', 'Category', 'Category'), Me('Money Out', 'Money out'), 'Where the money went', 'Outgoing money by category', OUT_C,
    {tooltips: [Me('Transactions'), Me('Charges Paid')]}),
  bar([16, 438, 300, 274], C('MoMo', 'Category', 'Category'), Me('Money In', 'Money in'), 'Where the money came from', 'Incoming money by category', IN_C,
    {tooltips: [Me('Transactions')]}),
  visual([324, 438, 312, 274], 'tableEx', {Values: [C('MoMo', 'Counterparty', PUBLIC ? 'Recipient (masked)' : 'Recipient'), Me('Money Out', 'Sent'), Me('Transactions', 'Txns'), Me('Charges Paid', 'Charges')]},
    {title: 'Top recipients', subtitle: 'People and businesses paid, by money sent', sort: [Me('Money Out'), 'Descending'],
      filters: [inFilter('MoMo', 'Category', ['IC savings & investments', 'Airtime & data', 'Pension (United Pensions)', 'YelloSave savings'], true), inFilter('MoMo', 'Direction', ['Out'])],
      objects: {columnHeaders: [{properties: {fontSize: lit('8D'), bold: lit('true'), fontColor: color(INK)}}], values: [{properties: {fontSize: lit('8D'), fontColor: color(INK)}}]}}),
  visual([644, 438, 312, 274], 'tableEx', {Values: [C('MoMo', 'Counterparty', PUBLIC ? 'Sender (masked)' : 'Sender'), Me('Money In', 'Received'), Me('Transactions', 'Txns')]},
    {title: 'Top senders', subtitle: 'Who paid money in, by amount', sort: [Me('Money In'), 'Descending'],
      filters: [inFilter('MoMo', 'Direction', ['In'])],
      objects: {columnHeaders: [{properties: {fontSize: lit('8D'), bold: lit('true'), fontColor: color(INK)}}], values: [{properties: {fontSize: lit('8D'), fontColor: color(INK)}}]}}),
  textbox([964, 438, 300, 274], [head('What the data says'),
    bulletS(`${pct(I.workIn, I.moneyIn)} of all money in (${ghs(I.workIn)}) was work money from the employer. The wallet is largely a pass-through, so personal spending is well below the totals.`),
    bulletS(`MTN’s newer statements (from Jun 2026) leave days out: ${I.breaks.filter(b => b.layout !== 'Old').length} balance breaks, confirmed by ${NF} IC payments (${ghs(NF_SUM)}) with no wallet row. The older ones have one gap: ${ghs(-I.breaks.find(b => b.layout === 'Old').gap)} on 20 May 2026.`),
    bulletS(`${ghs(I.chargesPaid)} in charges came off the balance. Older statements showed none of the ${ghs(I.oldCharges.chargesPaid)} charged to Jun 2026. The newer layout itemises every fee.`),
    bulletS(`${pct(I.savings.ic + I.savings.pension + I.savings.yello, I.moneyOut)} of money out was saved: IC ${ghs(I.savings.ic)}, pension ${ghs(I.savings.pension)}, YelloSave ${ghs(I.savings.yello)}.`),
    bulletS(`The balance ended a day under ${ghs(LOW_BALANCE)} on ${I.lowBalanceDays} days. Money rarely stays in the wallet.`)]),
]});
// ---------- Page 2: IC Wealth ----------
const best = ic.holdings.slice().sort((a, b) => b[2] - a[2])[0], worst = ic.holdings.slice().sort((a, b) => a[2] - b[2])[0];
pages.push({name: 'ic', displayName: 'IC Wealth', visuals: [
  header('IC Wealth', `1 Jan – 16 Nov 2026 · Liquidity Fund + TradeLive shares, each payment matched to MoMo. ${PUBLIC ? 'Anonymised.' : 'Private.'} Click a stock or bar, then List selected →.`, 470),
  listSelected([494, 10, 130, 24], IC_DETAIL), openDetail([494, 38, 130, 24], BM.icAll, 'All IC transactions'),
  ...slicerBar([[C('Date', 'Month', 'Month'), 'sync_month'], [C('IC', 'Account', 'IC account'), 'sync_icacct'], [C('IC', 'Type', 'Transaction type'), 'sync_ictype']]),
  kpi([16, 68, 200, 92], 'Portfolio Value', 'Portfolio value', 'At 16 Nov 2026 (not filtered)'),
  kpi([224, 68, 200, 92], 'Unrealised Gain', 'Gain on shares held', 'Value minus cost (not filtered)', IN_C),
  kpi([432, 68, 200, 92], 'Unrealised Return %', 'Return on cost', 'Shares still held'),
  kpi([640, 68, 200, 92], 'IC Paid In from MoMo', 'Paid in from MoMo', 'Contributions + share funding'),
  ...linkedKpi([848, 68, 200, 92], 'Trading Cost %', 'Trading cost', 'Click to list the trades', BM.trades, OUT_C, 'List every share purchase and sale'),
  ...linkedKpi([1056, 68, 208, 92], 'MoMo Match Rate', 'Matched to a wallet row', 'Click to list the 28 missing', BM.unmatched, IN_C, 'List IC payments with no wallet row'),
  visual([16, 168, 620, 262], 'clusteredColumnChart', {Category: [C('Date', 'Month', 'Month')], Y: [Me('IC Paid In from MoMo', 'Paid in'), Me('IC Paid Out to You', 'Paid out')]},
    {title: 'Money into and out of IC by month', subtitle: 'Small daily top-ups, then lump-sum payouts',
      sort: [C('Date', 'Month'), 'Ascending'], objects: Object.assign({}, axes, {legend: [{properties: {show: lit('true'), position: str('TopRight'), labelColor: color(INK)}}],
        dataPoint: [{properties: {fill: color(IN_C)}, selector: {metadata: '_Measures.IC Paid In from MoMo'}}, {properties: {fill: color(OUT_C)}, selector: {metadata: '_Measures.IC Paid Out to You'}}]})}),
  bar([644, 168, 620, 262], C('Holdings', 'Asset', 'Stock'), Me('Unrealised Gain', 'Gain / loss'), 'Gain or loss by stock', 'Market value minus what was paid (weighted average cost)', NAVY,
    {tooltips: [Me('Equity Value', 'Market value'), Me('Cost Basis', 'Cost')], filters: [inFilter('Holdings', 'Class', ['Equities'])]}),
  visual([16, 438, 520, 274], 'tableEx', {Values: [C('Holdings', 'Asset', 'Stock'), C('Holdings', 'Quantity', 'Shares'), C('Holdings', 'AvgCost', 'Avg cost'),
    C('Holdings', 'MarketPrice', 'Price'), Me('Equity Value', 'Value'), Me('Unrealised Gain', 'Gain / loss')]},
    {title: 'Holdings', subtitle: 'Shares held at 16 Nov 2026', sort: [Me('Equity Value'), 'Descending'], filters: [inFilter('Holdings', 'Class', ['Equities'])],
      objects: {columnHeaders: [{properties: {fontSize: lit('8D'), bold: lit('true'), fontColor: color(INK)}}], values: [{properties: {fontSize: lit('8D'), fontColor: color(INK)}}]}}),
  textbox([544, 438, 720, 274], [head('What the data says'),
    bullet(`The savings habit is micro: ${ic.contribSizes[0][1]} of ${ic.contribN} contributions were ${ghs(+ic.contribSizes[0][0])} top-ups. Together they added ${ghs(ic.contributionsMoMo)} in 2026.`),
    bullet(`Shares cost ${ghs(ic.cost)} and are worth ${ghs(ic.equities)}: ${ic.unrealised >= 0 ? '+' : ''}${ghs(ic.unrealised)} (${(100 * ic.unrealised / ic.cost).toFixed(1)}%). Best: ${best[0]} (${ghs(best[2])}${best[0] === 'KASA' ? ', Kasapreko IPO shares' : ''}). Weakest: ${worst[0]} (${ghs(worst[2])}).`),
    bullet(`${ic.trades} trades cost ${ghs(ic.tradeCharges)} in charges, ${(100 * ic.tradeCharges / (ic.bought + ic.sold)).toFixed(1)}% of value traded. Small, frequent buys pay a proportionally high fee.`),
    bullet(`IPO: ${ghs(ic.ipoSub)} subscribed, ${ghs(ic.ipoRefund)} refunded (partial allotment).`),
    bullet(`Every IC payment to 19 Jun 2026 matches a wallet transaction. After that, ${NF} payments (${ghs(NF_SUM)}) have none: proof the newer MTN statements are incomplete. IC posts on average ${ic.lagAvg} days after MoMo.`)]),
]});

// ---------- Page 3: MoMo transactions (drill-through target) ----------
const MOMO_DRILL = [['MoMo', 'Category'], ['MoMo', 'Direction'], ['MoMo', 'Counterparty'], ['MoMo', 'Statement'], ['MoMo', 'Layout'], ['Date', 'Month']];
pages.push({name: MOMO_DETAIL, displayName: 'MoMo transactions', visuals: [
  header('MoMo transactions', 'Every wallet transaction. Arrives filtered from the summary page; the month, direction and category filters stay in sync.', 470),
  actionButton([494, 10, 130, 24], 'Clear card filter', {type: str('Bookmark'), bookmark: str(BM.momoAll)}),
  actionButton([494, 38, 130, 24], '← Back', {type: str('PageNavigation'), navigationSection: str(MOMO_PAGE)}),
  ...slicerBar([[C('Date', 'Month', 'Month'), 'sync_month'], [C('MoMo', 'Direction', 'In / Out'), 'sync_dir'], [C('MoMo', 'Category', 'Category'), 'sync_cat']]),
  kpi([16, 68, 200, 92], 'MoMo Rows Listed', 'Transactions listed'),
  kpi([224, 68, 200, 92], 'Money In', 'Money in'), kpi([432, 68, 200, 92], 'Money Out', 'Money out'),
  kpi([640, 68, 200, 92], 'Charges Paid', 'Charges paid', '', OUT_C), kpi([848, 68, 200, 92], 'Unlisted Movement', 'Balance gaps (net)', 'Money moved with nothing listed', OUT_C),
  slicer([1056, 68, 208, 92], C('MoMo', 'Counterparty', PUBLIC ? 'Counterparty (masked)' : 'Counterparty'), 'sync_cp'),
  visual([16, 168, 1248, 544], 'tableEx', {Values: [C('MoMo', 'DateTime', 'Date & time'), C('MoMo', 'Direction', 'In/Out'), C('MoMo', 'Category', 'Category'),
    C('MoMo', 'Counterparty', 'Counterparty'), C('MoMo', 'Type', 'Type'), C('MoMo', 'Amount', 'Amount'), C('MoMo', 'ChargesPaid', 'Charges'), C('MoMo', 'FeeShown', 'Fee shown'),
    C('MoMo', 'BalanceBefore', 'Balance before'), C('MoMo', 'BalanceAfter', 'Balance after'), C('MoMo', 'UnlistedMovement', 'Gap before row'),
    C('MoMo', 'Statement', 'Statement'), C('MoMo', 'TxnID', 'MTN transaction ID')]},
    {title: 'Transactions', subtitle: 'Newest first. "Gap before row" = balance moved with no transaction listed (MTN statement gap).', sort: [C('MoMo', 'DateTime'), 'Descending'], objects: tableText}),
], filters: [...Object.keys(DETAIL_FILTERS[MOMO_DETAIL]).map(k => pageFilter(MOMO_DETAIL, k)), ...MOMO_DRILL.map(([t, c]) => drillFilter(t, c))],
  pageBinding: {name: 'dt_momo', type: 'Drillthrough', referenceScope: 'Default', acceptsFilterContext: 'Default'}});
// ---------- Page 4: IC transactions (drill-through target) ----------
const IC_DRILL = [['IC', 'Type'], ['IC', 'Account'], ['IC', 'MatchStatus'], ['IC', 'Ticker'], ['Holdings', 'Asset'], ['Date', 'Month']];
pages.push({name: IC_DETAIL, displayName: 'IC transactions', visuals: [
  header('IC transactions', 'Every IC Liquidity Fund and TradeLive row with its matched MoMo transaction. Filters stay in sync with the IC page.', 470),
  actionButton([494, 10, 130, 24], 'Clear card filter', {type: str('Bookmark'), bookmark: str(BM.icAll)}),
  actionButton([494, 38, 130, 24], '← Back', {type: str('PageNavigation'), navigationSection: str(IC_PAGE)}),
  ...slicerBar([[C('Date', 'Month', 'Month'), 'sync_month'], [C('IC', 'Account', 'IC account'), 'sync_icacct'], [C('IC', 'Type', 'Transaction type'), 'sync_ictype']]),
  kpi([16, 68, 200, 92], 'IC Rows Listed', 'IC rows listed'),
  kpi([224, 68, 200, 92], 'IC Paid In from MoMo', 'Paid in from MoMo'), kpi([432, 68, 200, 92], 'IC Paid Out to You', 'Paid out to MoMo'),
  kpi([640, 68, 200, 92], 'Trading Charges', 'Trading charges', '', OUT_C),
  slicer([848, 68, 200, 92], C('IC', 'MatchStatus', 'Match with MoMo'), 'sync_match'),
  slicer([1056, 68, 208, 92], C('IC', 'Ticker', 'Stock'), 'sync_ticker'),
  visual([16, 168, 1248, 544], 'tableEx', {Values: [C('IC', 'Date', 'Date'), C('IC', 'Account', 'Account'), C('IC', 'Type', 'Type'), C('IC', 'Description', 'Description'),
    C('IC', 'Credit', 'Credit'), C('IC', 'Debit', 'Debit'), C('IC', 'Ticker', 'Stock'), C('IC', 'TradeCharges', 'Trade charges'),
    C('IC', 'MatchStatus', 'Match with MoMo'), C('IC', 'MoMoTxnID', 'MoMo transaction ID'), C('IC', 'LagDays', 'Days after MoMo')]},
    {title: 'IC transactions', subtitle: '"NOT FOUND in MoMo" = the wallet paid it but MTN\'s statement does not list it.', sort: [C('IC', 'Date'), 'Descending'], objects: tableText}),
], filters: [...Object.keys(DETAIL_FILTERS[IC_DETAIL]).map(k => pageFilter(IC_DETAIL, k)), ...IC_DRILL.map(([t, c]) => drillFilter(t, c))],
  pageBinding: {name: 'dt_ic', type: 'Drillthrough', referenceScope: 'Default', acceptsFilterContext: 'Default'}});

fs.rmSync(RPT, {recursive: true, force: true});
const RD = path.join(RPT, 'definition');
for (const p of pages) {
  w(path.join(RD, 'pages', p.name, 'page.json'), {$schema: `${SCH}/item/report/definition/page/1.0.0/schema.json`, name: p.name, displayName: p.displayName,
    displayOption: 'FitToPage', height: 720, width: 1280, objects: {background: [{properties: {color: color('#F5F6F8'), transparency: lit('0D')}}]},
    ...(p.filters ? {filterConfig: {filters: p.filters}} : {}),
    // drill-through needs one parameter per drill field, bound to its filter by name (as Power BI Desktop saves it)
    ...(p.pageBinding ? {pageBinding: Object.assign({}, p.pageBinding, {parameters: (p.filters || []).filter(f => f.howCreated === 'Drillthrough')
      .map(f => ({name: id(), boundFilter: f.name, asAggregation: false, qnaSingleSelectRequired: false, fieldExpr: f.field}))})} : {})});
  p.visuals.forEach((v, i) => { v.position.z = i * 1000; v.position.tabOrder = i * 1000; w(path.join(RD, 'pages', p.name, 'visuals', v.name, 'visual.json'), v); });
}
w(path.join(RD, 'pages', 'pages.json'), {$schema: `${SCH}/item/report/definition/pagesMetadata/1.0.0/schema.json`, pageOrder: pages.map(p => p.name), activePageName: pages[0].name});
bookmarks.forEach(b => w(path.join(RD, 'bookmarks', `${b.name}.bookmark.json`), b));
w(path.join(RD, 'bookmarks', 'bookmarks.json'), {$schema: `${SCH}/item/report/definition/bookmarksMetadata/1.0.0/schema.json`, items: bookmarks.map(b => ({name: b.name}))});
w(path.join(RD, 'version.json'), {$schema: `${SCH}/item/report/definition/versionMetadata/1.0.0/schema.json`, version: '2.0.0'});
w(path.join(RD, 'report.json'), {$schema: `${SCH}/item/report/definition/report/1.0.0/schema.json`,
  themeCollection: {customTheme: {name: 'MoneyReport.json', reportVersionAtImport: '5.55', type: 'RegisteredResources'}}, layoutOptimization: 'None',
  resourcePackages: [{name: 'RegisteredResources', type: 'RegisteredResources', items: [{name: 'MoneyReport.json', path: 'MoneyReport.json', type: 'CustomTheme'}]}],
  settings: {useStylableVisualContainerHeader: true, defaultDrillFilterOtherVisuals: true, allowChangeFilterTypes: true, useEnhancedTooltips: true, useDefaultAggregateDisplayName: true, defaultFilterActionIsDataFilter: true}});
w(path.join(RPT, 'StaticResources', 'RegisteredResources', 'MoneyReport.json'), {name: 'Money Report', dataColors: [NAVY, IN_C, '#E9A23B', OUT_C, '#6C7A89', '#8AB17D', '#9B6A9E', '#4F8FC0'],
  background: '#FFFFFF', foreground: INK, tableAccent: NAVY,
  visualStyles: {'*': {'*': {border: [{show: true, color: {solid: {color: '#E1E4E8'}}, radius: 6}], background: [{show: true, color: {solid: {color: '#FFFFFF'}}, transparency: 0}]}},
    textbox: {'*': {border: [{show: false}], background: [{show: false}]}}}});
w(path.join(RPT, 'definition.pbir'), {$schema: `${SCH}/item/report/definitionProperties/2.0.0/schema.json`, version: '4.0', datasetReference: {byPath: {path: `../${NAME}.SemanticModel`}}});
const platform = type => ({$schema: `${SCH}/gitIntegration/platformProperties/2.0.0/schema.json`, metadata: {type, displayName: NAME}, config: {version: '2.0', logicalId: uid()}});
w(path.join(RPT, '.platform'), platform('Report')); w(path.join(SM, '.platform'), platform('SemanticModel'));
w(path.join(SM, 'definition.pbism'), {$schema: `${SCH}/item/semanticModel/definitionProperties/1.0.0/schema.json`, version: '4.2', settings: {}});
w(path.join(ROOT, `${NAME}.pbip`), {$schema: `${SCH}/pbip/pbipProperties/1.0.0/schema.json`, version: '1.0', artifacts: [{report: {path: `${NAME}.Report`}}], settings: {enableAutoRecovery: true}});
w(path.join(ROOT, '.gitignore'), '**/.pbi/localSettings.json\n**/.pbi/cache.abf\n');
console.log(`${NAME}: ${pages.map(p => `${p.displayName} ${p.visuals.length} visuals`).join(', ')}; ${M.length} measures; data = ${DATA}`);
