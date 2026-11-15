// Run with: node tests/recon.test.js
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
for (const f of ['vendor/sheet-engine.js', 'vendor/sheet-io.js', 'src/recon.js', 'src/sample.js']) vm.runInThisContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), { filename: f });

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}
const side = (s) => ({ name: s.name, headers: s.rows[0], rows: s.rows.slice(1), firstDataRow: 2 });

const now = new Date(2026, 10, 15, 12);
const S = ReconSample.build(now);
const A = side(S.a), B = side(S.b);

console.log('Mapping suggestion');
const cfg = Recon.suggestMapping(A, B);
test('key: Tx hash ↔ Transaction ID (despite 0x and case)', () => assert.deepStrictEqual(cfg.keys, [{ a: 'Tx hash', b: 'Transaction ID' }]));
test('amount: Amount ↔ Quantity, exact for crypto decimals', () => {
  const c = cfg.compares.find(c => c.type === 'number');
  assert.deepStrictEqual([c.a, c.b, c.tol, c.abs], ['Amount', 'Quantity', 0, false]);
});
test('date: Booked date ↔ Settled at, ±1 day', () => {
  const c = cfg.compares.find(c => c.type === 'date');
  assert.deepStrictEqual([c.a, c.b, c.tol], ['Booked date', 'Settled at', 1]);
});
test('group: Asset ↔ Coin', () => assert.deepStrictEqual(cfg.group, { a: 'Asset', b: 'Coin' }));

console.log('Reconciliation');
const r = Recon.reconcile(A, B, cfg);
const by = (s) => r.results.filter(x => x.status === s);
test('no setup problems', () => assert.deepStrictEqual(r.problems, []));
test('amount break on USDC tx 3 (-10)', () => {
  const m = by('mismatch').find(x => x.diffs.some(d => d.field.startsWith('Amount')));
  assert.ok(m && m.diffs[0].diff === -10, JSON.stringify(by('mismatch').map(x => x.diffs)));
});
test('date break 3 days', () => assert.ok(by('mismatch').some(x => x.diffs.some(d => d.note === '+3 days'))));
test('asset break ETH vs WETH', () => assert.ok(by('mismatch').some(x => x.diffs.some(d => d.field.startsWith('Asset')))));
test('1-day settlement lag is within tolerance', () => assert.strictEqual(by('mismatch').length, 3, JSON.stringify(by('mismatch').map(x => [x.a.row, x.diffs.map(d => d.note)]))));
test('tx 10 missing on the statement', () => assert.strictEqual(by('missing_in_b').length, 2));   // tx 10 + the duplicate booking
test('tx 99 never booked', () => assert.strictEqual(by('missing_in_a').length, 1));
test('duplicate booking flagged', () => assert.strictEqual(r.stats.duplicatesA, 1));
test('row without ID matched on values', () => assert.strictEqual(by('matched_no_key').length, 1));
test('everything else matched', () => assert.strictEqual(r.stats.matched, 11));
test('totals per asset show the breaks', () => {
  const t = Object.fromEntries(r.totals.map(x => [x.group, x.diff]));
  assert.strictEqual(t.USDC, -15010);   // -10 amount break, -15000 duplicate booking
  assert.deepStrictEqual([t.WETH, t.ETH], [-0.75, 1.65]);   // tx 8 booked as ETH, settled as WETH; tx 99 (+0.9) never booked
});

console.log('Edge cases');
test('opposite signs → absolute compare suggested', () => {
  const a = side({ name: 'A', rows: [['Ref', 'Amount'], ['R1', -10], ['R2', -20.5], ['R3', -3]] });
  const b = side({ name: 'B', rows: [['Reference', 'Value'], ['R1', 10], ['R2', 20.5], ['R3', 3]] });
  const c = Recon.suggestMapping(a, b);
  const res = Recon.reconcile(a, b, c);
  assert.strictEqual(res.stats.matched, 3, JSON.stringify(c));
});
test('no shared ID → matches on values', () => {
  const a = side({ name: 'A', rows: [['Date', 'Amount'], ['01.11.2026', '100,00'], ['02.11.2026', '50,00']] });
  const b = side({ name: 'B', rows: [['Valuta', 'Betrag'], ['2026-11-02', 50], ['2026-11-01', 100], ['2026-11-03', 7]] });
  const res = Recon.reconcile(a, b, { keys: [], compares: [{ a: 'Date', b: 'Valuta', type: 'date', tol: 0 }, { a: 'Amount', b: 'Betrag', type: 'number', tol: 0.005 }] });
  assert.deepStrictEqual([res.stats.matched, res.stats.missingInA], [2, 1]);
});
test('unknown column is a setup problem, not a crash', () => {
  const res = Recon.reconcile(A, B, { keys: [{ a: 'Nope', b: 'Transaction ID' }] });
  assert.ok(res.problems[0].includes('Nope'));
});
test('large: 20k vs 20k rows in under 2 s', () => {
  const ra = [['ID', 'Amt']], rb = [['ID', 'Amt']];
  for (let i = 0; i < 20000; i++) { ra.push(['T' + i, i]); rb.push(['T' + (i + 5), i + 5]); }
  const t0 = Date.now();
  const res = Recon.reconcile(side({ name: 'A', rows: ra }), side({ name: 'B', rows: rb }), { keys: [{ a: 'ID', b: 'ID' }], compares: [{ a: 'Amt', b: 'Amt', type: 'number' }] });
  const ms = Date.now() - t0;
  assert.ok(ms < 2000 && res.stats.matched === 19995 && res.stats.missingInB === 5, ms + ' ms ' + JSON.stringify(res.stats));
});


test('bank export with opposite signs → flip suggested', () => {
  const a = side({ name: 'Bank', rows: [['Valuta', 'Betrag', 'Referenz'], ['01.10.2026', '-1.200,00', 'R-1'], ['02.10.2026', '3.500,00', 'R-2'], ['03.10.2026', '-45,90', 'R-3'], ['05.10.2026', '-89,00', 'R-4']] });
  const b = side({ name: 'Ledger', rows: [['Booking date', 'Amount', 'Bank ref'], ['2026-10-01', 1200, 'R-1'], ['2026-10-02', -3500, 'R-2'], ['2026-10-04', 45.9, 'R-3'], ['2026-10-06', 29.99, 'R-5']] });
  const c = Recon.suggestMapping(a, b);
  const res = Recon.reconcile(a, b, c);
  assert.deepStrictEqual([c.compares[0].flip, res.stats.matched, res.stats.mismatch, res.stats.missingInB, res.stats.missingInA], [true, 3, 0, 1, 1], JSON.stringify(c));
  assert.strictEqual(res.totals[0].diff, 59.01);   // +89 (bank only) − 29.99 (ledger only, sign flipped)
});

console.log('Workbook export');
{
  vm.runInThisContext(fs.readFileSync(path.join(__dirname, '..', 'src/export.js'), 'utf8'));
  const bytes = ReconExport.build(A, B, cfg, r, { ranAt: now });
  test('xlsx is a zip', () => assert.deepStrictEqual(Array.from(bytes.slice(0, 2)), [0x50, 0x4b]));
  const csv = ReconExport.breaksCsv(A, B, cfg, r);
  test('breaks csv lists every non-match', () => assert.strictEqual(csv.trim().split('\r\n').length - 1, r.results.filter(x => x.status !== 'matched').length + r.duplicates.length));
  if (process.argv[2]) { fs.writeFileSync(process.argv[2], bytes); console.log('       wrote ' + process.argv[2]); }
}
console.log('\n' + passed + ' tests passed' + (process.exitCode ? ', some FAILED' : ''));
