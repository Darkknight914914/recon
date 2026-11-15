/**
 * Recon — two-sided reconciliation engine.
 *
 * Pure JavaScript, no DOM. Needs SheetEngine (toNumber, toDate, display) from
 * vendor/sheet-engine.js for value parsing.
 *
 * A side is { name, headers: [...], rows: [[...]], firstDataRow }.
 * A config is:
 * {
 *   keys:     [{ a: 'Tx hash', b: 'Transaction ID' }],          // match on these (all must be equal)
 *   compares: [{ a: 'Amount', b: 'Quantity', type: 'number', tol: 0.00000001, abs: false },
 *              { a: 'Date',   b: 'Settled',  type: 'date',   tol: 1 },   // days
 *              { a: 'Asset',  b: 'Coin',     type: 'text' }],
 *   group:    { a: 'Asset', b: 'Coin' },                         // optional: totals per group
 *   amount:   { a: 'Amount', b: 'Quantity' },                    // optional: column summed for totals
 *   fallback: true          // pair leftovers on all compare fields when the key doesn't match
 * }
 */
var Recon = (function () {
  'use strict';
  var E = typeof SheetEngine !== 'undefined' ? SheetEngine : require('../vendor/sheet-engine.js');

  // ---------------------------------------------------------------------------
  // helpers
  // ---------------------------------------------------------------------------
  function blank(v) { return v === null || v === undefined || (typeof v === 'string' && v.trim() === ''); }

  function normKey(v) {
    if (blank(v)) return '';
    if (typeof v === 'number') return String(v);
    if (Object.prototype.toString.call(v) === '[object Date]') return E.display(v);
    return String(v).trim().replace(/\s+/g, ' ').toLowerCase().replace(/^0x/, '');
  }

  function normText(v) { return blank(v) ? '' : E.display(v).trim().replace(/\s+/g, ' ').toLowerCase(); }

  function colIndex(headers, name) {
    var k = String(name || '').trim().toLowerCase();
    for (var i = 0; i < headers.length; i++) if (String(headers[i]).trim().toLowerCase() === k) return i;
    return -1;
  }

  function round(n, dp) { var f = Math.pow(10, dp == null ? 8 : dp); return Math.round(n * f) / f; }

  function fmt(n) {
    if (typeof n !== 'number' || !isFinite(n)) return String(n);
    var s = round(n, 8).toString();
    return s;
  }

  // ---------------------------------------------------------------------------
  // compare one field
  // ---------------------------------------------------------------------------
  /** Returns null when equal (within tolerance), else { a, b, diff, note }. */
  function compareField(c, va, vb) {
    var ba = blank(va), bb = blank(vb);
    if (ba && bb) return null;
    if (ba || bb) return { a: va, b: vb, note: ba ? 'Empty on A' : 'Empty on B' };
    if (c.type === 'number') {
      var na = E.toNumber(va), nb = E.toNumber(vb);
      if (isNaN(na) || isNaN(nb)) return { a: va, b: vb, note: 'Not a number' };
      if (c.flip) nb = -nb;
      if (c.abs) { na = Math.abs(na); nb = Math.abs(nb); }
      var d = round(nb - na, 10);
      return Math.abs(d) <= (c.tol || 0) + 1e-12 ? null : { a: va, b: vb, diff: d, note: 'Differs by ' + (d > 0 ? '+' : '') + fmt(d) };
    }
    if (c.type === 'date') {
      var da = E.toDate(va), db = E.toDate(vb);
      if (!da || !db) return { a: va, b: vb, note: 'Not a date' };
      var dayA = Date.UTC(da.getFullYear(), da.getMonth(), da.getDate());
      var dayB = Date.UTC(db.getFullYear(), db.getMonth(), db.getDate());
      var days = Math.round((dayB - dayA) / 86400000);
      return Math.abs(days) <= (c.tol || 0) ? null : { a: va, b: vb, diff: days, note: (days > 0 ? '+' : '') + days + ' day' + (Math.abs(days) === 1 ? '' : 's') };
    }
    return normText(va) === normText(vb) ? null : { a: va, b: vb, note: 'Different value' };
  }

  // ---------------------------------------------------------------------------
  // reconcile
  // ---------------------------------------------------------------------------
  function reconcile(A, B, cfg) {
    var problems = [];
    function idx(side, name, what) {
      var i = colIndex(side.headers, name);
      if (i < 0) problems.push(what + ': column "' + name + '" not found in ' + side.name);
      return i;
    }
    var keys = (cfg.keys || []).map(function (k) { return { a: idx(A, k.a, 'Match on'), b: idx(B, k.b, 'Match on'), nameA: k.a, nameB: k.b }; });
    var comps = (cfg.compares || []).map(function (c) {
      return { a: idx(A, c.a, 'Compare'), b: idx(B, c.b, 'Compare'), nameA: c.a, nameB: c.b, type: c.type || 'text', tol: +c.tol || 0, abs: !!c.abs, flip: !!c.flip };
    });
    var group = cfg.group && cfg.group.a ? { a: idx(A, cfg.group.a, 'Totals by'), b: idx(B, cfg.group.b, 'Totals by') } : null;
    var amount = cfg.amount && cfg.amount.a ? { a: idx(A, cfg.amount.a, 'Totals of'), b: idx(B, cfg.amount.b, 'Totals of'), abs: !!cfg.amount.abs, flip: !!cfg.amount.flip } : null;
    if (!keys.length && !comps.length) problems.push('Choose at least one column to match on.');
    if (problems.length) return { problems: problems };

    function rowsOf(side) {
      var out = [];
      side.rows.forEach(function (r, i) {
        if (!r || !r.some(function (v) { return !blank(v); })) return;
        out.push({ side: side, r: r, row: (side.firstDataRow || 2) + i, matched: false });
      });
      return out;
    }
    var ra = rowsOf(A), rb = rowsOf(B);

    function keyOf(rec, which) {
      if (!keys.length) return '';
      var parts = keys.map(function (k) { return normKey(rec.r[k[which]]); });
      return parts.every(function (p) { return p === ''; }) ? '' : parts.join('␟');
    }

    var results = [];           // one entry per pair or unpaired row
    var dupA = 0, dupB = 0;

    function diffs(x, y) {
      var out = [];
      comps.forEach(function (c) {
        var d = compareField(c, x.r[c.a], y.r[c.b]);
        if (d) { d.field = c.nameA === c.nameB ? c.nameA : c.nameA + ' / ' + c.nameB; d.colA = c.a; d.colB = c.b; out.push(d); }
      });
      return out;
    }

    // 1) key matching
    if (keys.length) {
      var mapB = {};
      rb.forEach(function (rec) { var k = keyOf(rec, 'b'); if (!k) return; (mapB[k] = mapB[k] || []).push(rec); });
      var seenA = {};
      ra.forEach(function (rec) {
        var k = keyOf(rec, 'a');
        if (!k) return;
        var dup = seenA[k];
        seenA[k] = (seenA[k] || 0) + 1;
        if (dup) { dupA++; rec.duplicate = true; }
        var cands = mapB[k];
        if (!cands) return;
        var pick = null;
        for (var i = 0; i < cands.length; i++) if (!cands[i].matched) { pick = cands[i]; break; }
        if (!pick) return;
        rec.matched = pick.matched = true;
        var ds = diffs(rec, pick);
        results.push({ status: ds.length ? 'mismatch' : 'matched', a: rec, b: pick, diffs: ds, via: 'key' });
      });
      Object.keys(mapB).forEach(function (k) { mapB[k].forEach(function (rec, j) { if (j > 0) { dupB++; rec.duplicate = true; } }); });
    }

    // 2) fallback: pair leftovers on all compare fields (within tolerance)
    if ((cfg.fallback || !keys.length) && comps.length) {
      var leftB = rb.filter(function (r) { return !r.matched; });
      ra.forEach(function (rec) {
        if (rec.matched) return;
        for (var i = 0; i < leftB.length; i++) {
          var cand = leftB[i];
          if (cand.matched) continue;
          if (!diffs(rec, cand).length) {
            rec.matched = cand.matched = true;
            results.push({ status: keys.length ? 'matched_no_key' : 'matched', a: rec, b: cand, diffs: [], via: 'values' });
            break;
          }
        }
      });
    }

    // 3) leftovers
    ra.forEach(function (rec) { if (!rec.matched) results.push({ status: 'missing_in_b', a: rec, b: null, diffs: [] }); });
    rb.forEach(function (rec) { if (!rec.matched) results.push({ status: 'missing_in_a', a: null, b: rec, diffs: [] }); });

    // duplicates as their own flag (rows may also be matched/missing)
    var dupRows = ra.filter(function (r) { return r.duplicate; }).map(function (r) { return { status: 'duplicate', a: r, b: null, diffs: [] }; })
      .concat(rb.filter(function (r) { return r.duplicate; }).map(function (r) { return { status: 'duplicate', a: null, b: r, diffs: [] }; }));

    // 4) totals
    var totals = null;
    if (amount && amount.a >= 0 && amount.b >= 0) {
      var g = {};
      function addTo(rec, which) {
        var key = group ? (E.display(rec.r[group[which]]).trim() || '(blank)') : 'All';
        var gk = key.toUpperCase();
        var t = g[gk] || (g[gk] = { group: key, a: 0, b: 0, countA: 0, countB: 0 });
        var n = E.toNumber(rec.r[amount[which]]);
        if (isNaN(n)) return;
        if (amount.flip && which === 'b') n = -n;
        if (amount.abs) n = Math.abs(n);
        t[which] += n; t[which === 'a' ? 'countA' : 'countB']++;
      }
      ra.forEach(function (r) { addTo(r, 'a'); });
      rb.forEach(function (r) { addTo(r, 'b'); });
      totals = Object.keys(g).sort().map(function (k) {
        var t = g[k]; t.a = round(t.a, 8); t.b = round(t.b, 8); t.diff = round(t.b - t.a, 8); return t;
      });
    }

    var order = { mismatch: 0, missing_in_b: 1, missing_in_a: 2, matched_no_key: 3, matched: 4 };
    results.sort(function (x, y) {
      return (order[x.status] - order[y.status]) || ((x.a ? x.a.row : 1e9) - (y.a ? y.a.row : 1e9)) || ((x.b ? x.b.row : 1e9) - (y.b ? y.b.row : 1e9));
    });

    var count = function (s) { return results.filter(function (r) { return r.status === s; }).length; };
    return {
      problems: [],
      results: results,
      duplicates: dupRows,
      totals: totals,
      stats: {
        rowsA: ra.length, rowsB: rb.length,
        matched: count('matched'), matchedNoKey: count('matched_no_key'), mismatch: count('mismatch'),
        missingInB: count('missing_in_b'), missingInA: count('missing_in_a'),
        duplicatesA: dupA, duplicatesB: dupB
      }
    };
  }

  // ---------------------------------------------------------------------------
  // suggest a mapping from the two sheets
  // ---------------------------------------------------------------------------
  var ROLES = {
    key: /(^|\b)(id|ref|reference|referenz|tx|txid|tx id|tx hash|txn|hash|transaction id|transaction|trade id|order id|invoice|rechnung|beleg|booking|buchung|external id)(\b|$)/i,
    amount: /(amount|amt|quantity|qty|value|betrag|summe|net|gross|volume|balance|saldo)/i,
    date: /(date|datum|time|timestamp|settled|booked|created|value date|valuta|executed)/i,
    group: /(asset|coin|currency|ccy|währung|token|symbol|instrument|account|konto|wallet)/i
  };

  function profile(side) {
    var hs = side.headers.map(function (h) { return String(h).trim(); });
    return hs.map(function (h, c) {
      var vals = side.rows.map(function (r) { return r[c]; }).filter(function (v) { return !blank(v); });
      var n = vals.length || 1;
      var nums = vals.filter(function (v) { return !isNaN(E.toNumber(v)); }).length / n;
      var dates = vals.filter(function (v) { return !!E.toDate(v); }).length / n;
      var set = {};
      vals.forEach(function (v) { set[normKey(v)] = true; });
      var distinct = Object.keys(set).length;
      return { name: h, c: c, filled: vals.length, nums: nums, dates: dates, distinctRatio: distinct / n, distinct: distinct, set: set };
    });
  }

  function overlap(pa, pb) {
    var small = pa.distinct < pb.distinct ? pa : pb, big = small === pa ? pb : pa;
    if (!small.distinct) return 0;
    var hit = 0;
    for (var k in small.set) if (big.set[k]) hit++;
    return hit / small.distinct;
  }

  function nameScore(a, b, role) {
    var x = a.toLowerCase(), y = b.toLowerCase();
    if (x === y) return 1;
    if (role && ROLES[role].test(a) && ROLES[role].test(b)) return 0.7;
    var ta = x.split(/[^a-z0-9äöü]+/).filter(Boolean), tb = y.split(/[^a-z0-9äöü]+/).filter(Boolean);
    var common = ta.filter(function (t) { return tb.indexOf(t) > -1; }).length;
    return common ? 0.4 + 0.2 * common / Math.max(ta.length, tb.length) : 0;
  }

  function suggestMapping(A, B) {
    var pa = profile(A), pb = profile(B);
    var cfg = { keys: [], compares: [], group: null, amount: null, fallback: true };
    if (!pa.length || !pb.length) return cfg;

    // key: mostly-distinct columns whose values overlap the most
    var best = null;
    pa.forEach(function (x) {
      if (x.distinctRatio < 0.9 || x.filled < 2) return;
      pb.forEach(function (y) {
        if (y.distinctRatio < 0.9 || y.filled < 2) return;
        if (x.nums > 0.9 && y.nums > 0.9 && !(ROLES.key.test(x.name) || ROLES.key.test(y.name))) return; // amounts, not IDs
        if (x.dates > 0.9 || y.dates > 0.9) return;
        var s = overlap(x, y) + 0.3 * nameScore(x.name, y.name, 'key');
        if (!best || s > best.s) best = { s: s, a: x.name, b: y.name };
      });
    });
    if (best && best.s >= 0.35) cfg.keys.push({ a: best.a, b: best.b });

    function pick(role, test) {
      var bestP = null;
      pa.forEach(function (x) {
        if (!test(x)) return;
        pb.forEach(function (y) {
          if (!test(y)) return;
          var s = nameScore(x.name, y.name, role) + (ROLES[role].test(x.name) ? 0.2 : 0) + (ROLES[role].test(y.name) ? 0.2 : 0);
          if (role === 'group') s += overlap(x, y);
          if (!bestP || s > bestP.s) bestP = { s: s, a: x.name, b: y.name };
        });
      });
      return bestP && bestP.s >= 0.4 ? bestP : null;
    }
    var used = function (n, side) { return cfg.keys.some(function (k) { return k[side] === n; }); };
    var amt = pick('amount', function (p) { return p.nums >= 0.9 && p.dates < 0.5 && !used(p.name, p === pa[p.c] ? 'a' : 'b'); });
    var dt = pick('date', function (p) { return p.dates >= 0.9; });
    var grp = pick('group', function (p) { return p.nums < 0.5 && p.dates < 0.5 && p.distinct <= 30 && p.distinctRatio < 0.5; });

    if (amt) {
      var tol = 0.005;
      var decimals = function (side, name) {
        var c = colIndex(side.headers, name), max = 0;
        side.rows.forEach(function (r) {
          var n = E.toNumber(r[c]);
          if (isNaN(n)) return;
          var s = String(n), d = s.indexOf('.') > -1 ? s.length - s.indexOf('.') - 1 : 0;
          if (/e-/.test(s)) d = 10;
          max = Math.max(max, d);
        });
        return max;
      };
      var dp = Math.max(decimals(A, amt.a), decimals(B, amt.b));
      if (dp > 2) tol = 0;
      // opposite signs (ledger debit vs statement credit) → compare absolute values
      var sumSign = function (side, name) {
        var c = colIndex(side.headers, name), neg = 0, tot = 0;
        side.rows.forEach(function (r) { var n = E.toNumber(r[c]); if (!isNaN(n) && n !== 0) { tot++; if (n < 0) neg++; } });
        return tot ? neg / tot : 0;
      };
      // Sign convention: compare the pairs that share an ID. Bank exports often show
      // an outgoing payment as -100 where the ledger books +100.
      var flip = false, abs = false;
      if (cfg.keys.length) {
        var ka = colIndex(A.headers, cfg.keys[0].a), kb = colIndex(B.headers, cfg.keys[0].b);
        var ca = colIndex(A.headers, amt.a), cb = colIndex(B.headers, amt.b), byKey = {};
        B.rows.forEach(function (r) { var k = normKey(r[kb]); if (k && !(k in byKey)) byKey[k] = E.toNumber(r[cb]); });
        var same = 0, opp = 0;
        A.rows.forEach(function (r) {
          var k = normKey(r[ka]), na = E.toNumber(r[ca]);
          if (!k || !(k in byKey) || isNaN(na) || isNaN(byKey[k]) || na === 0) return;
          if (Math.abs(na - byKey[k]) <= tol + 1e-12) same++;
          else if (Math.abs(na + byKey[k]) <= tol + 1e-12) opp++;
        });
        flip = opp > same;
      } else {
        abs = Math.abs(sumSign(A, amt.a) - sumSign(B, amt.b)) > 0.6;
      }
      cfg.compares.push({ a: amt.a, b: amt.b, type: 'number', tol: tol, abs: abs, flip: flip });
      cfg.amount = { a: amt.a, b: amt.b, abs: abs, flip: flip };
    }
    if (grp) { cfg.compares.push({ a: grp.a, b: grp.b, type: 'text' }); cfg.group = { a: grp.a, b: grp.b }; }
    if (dt) cfg.compares.push({ a: dt.a, b: dt.b, type: 'date', tol: 1 });
    return cfg;
  }

  return { reconcile: reconcile, suggestMapping: suggestMapping, compareField: compareField, normKey: normKey };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = Recon;
