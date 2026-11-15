/**
 * Recon — builds the reconciliation workbook (.xlsx) and a breaks CSV.
 * Uses the dependency-free xlsx writer in vendor/sheet-io.js (SheetIO.buildXlsx).
 */
var ReconExport = (function () {
  'use strict';
  var E = SheetEngine, W = SheetIO;

  var LABEL = {
    mismatch: 'Break', missing_in_b: 'Missing in B', missing_in_a: 'Missing in A',
    matched_no_key: 'Matched without ID', matched: 'Matched', duplicate: 'Duplicate'
  };

  function v(x) { return x === undefined || x === null ? '' : x; }
  function isDate(x) { return Object.prototype.toString.call(x) === '[object Date]'; }
  function cell(x) { return isDate(x) ? x : (typeof x === 'number' || typeof x === 'boolean') ? x : String(v(x)); }

  function explain(res) {
    if (res.status === 'mismatch') return res.diffs.map(function (d) { return d.field + ': ' + d.note; }).join('; ');
    if (res.status === 'missing_in_b') return res.a && res.a.duplicate ? 'Booked twice in A; this copy has no counterpart in B' : 'In A, not found in B';
    if (res.status === 'missing_in_a') return res.b && res.b.duplicate ? 'Appears twice in B; this copy has no counterpart in A' : 'In B, not found in A';
    if (res.status === 'matched_no_key') return 'No matching ID; paired on values';
    if (res.status === 'duplicate') return 'Same ID appears more than once on this side';
    return '';
  }

  function head(rows, styles, r, n) { for (var c = 0; c < n; c++) styles[r + ',' + c] = 'head'; }

  function build(A, B, cfg, out, meta) {
    meta = meta || {};
    var sheets = [];
    var st = out.stats;

    // Summary
    var sum = [
      ['Reconciliation', ''],
      ['Side A', A.name], ['Side B', B.name],
      ['Matched on', (cfg.keys || []).map(function (k) { return k.a + ' = ' + k.b; }).join(', ') || 'values only'],
      ['Compared', (cfg.compares || []).map(function (c) {
        return c.a + ' = ' + c.b + (c.type === 'number' && c.tol ? ' (±' + c.tol + ')' : '') + (c.type === 'date' && c.tol ? ' (±' + c.tol + ' days)' : '') + (c.abs ? ' (sign ignored)' : '') + (c.flip ? ' (opposite sign)' : '');
      }).join(', ')],
      ['Run at', meta.ranAt || new Date()],
      ['', ''],
      ['Rows in A', st.rowsA], ['Rows in B', st.rowsB],
      ['Matched', st.matched], ['Matched without ID', st.matchedNoKey], ['Breaks', st.mismatch],
      ['Missing in B', st.missingInB], ['Missing in A', st.missingInA],
      ['Duplicates in A', st.duplicatesA], ['Duplicates in B', st.duplicatesB]
    ];
    var sumStyles = { '0,0': 'head', '0,1': 'head' };
    [11, 12, 13].forEach(function (r) { if (sum[r][1]) sumStyles[r + ',1'] = 'err'; });
    [14, 15].forEach(function (r) { if (sum[r][1]) sumStyles[r + ',1'] = 'warn'; });
    sheets.push({ name: 'Summary', rows: sum, styles: sumStyles, widths: [22, 70] });

    // Totals
    if (out.totals) {
      var tot = [[cfg.group ? cfg.group.a + ' / ' + cfg.group.b : 'Group', 'Total A', 'Total B', 'Difference (B − A)', 'Rows A', 'Rows B']];
      var tStyles = {};
      head(tot, tStyles, 0, 6);
      out.totals.forEach(function (t, i) {
        tot.push([t.group, t.a, t.b, t.diff, t.countA, t.countB]);
        if (t.diff !== 0) tStyles[(i + 1) + ',3'] = 'err';
      });
      sheets.push({ name: 'Totals', rows: tot, styles: tStyles, freezeRow: 1, widths: [18, 16, 16, 18, 9, 9] });
    }

    // Exceptions (everything that is not a clean match), then duplicates
    var keyA = (cfg.keys || [])[0] ? cfg.keys[0].a : '';
    var exc = [['Status', 'A row', 'B row', 'Key', 'Explanation']];
    var compCols = (cfg.compares || []);
    compCols.forEach(function (c) { exc[0].push('A: ' + c.a, 'B: ' + c.b); });
    var eStyles = {}, links = [];
    head(exc, eStyles, 0, exc[0].length);
    var ia = function (name) { return A.headers.map(function (h) { return String(h).trim().toLowerCase(); }).indexOf(String(name).trim().toLowerCase()); };
    var ib = function (name) { return B.headers.map(function (h) { return String(h).trim().toLowerCase(); }).indexOf(String(name).trim().toLowerCase()); };
    var kA = keyA ? ia(keyA) : -1, kB = keyA ? ib(cfg.keys[0].b) : -1;
    var list = out.results.filter(function (r) { return r.status !== 'matched'; }).concat(out.duplicates);
    list.forEach(function (res) {
      var key = res.a && kA >= 0 ? res.a.r[kA] : res.b && kB >= 0 ? res.b.r[kB] : '';
      var row = [LABEL[res.status], res.a ? res.a.row : '', res.b ? res.b.row : '', cell(key), explain(res)];
      compCols.forEach(function (c) {
        row.push(res.a ? cell(res.a.r[ia(c.a)]) : '', res.b ? cell(res.b.r[ib(c.b)]) : '');
      });
      var r = exc.length;
      exc.push(row);
      eStyles[r + ',0'] = res.status === 'matched_no_key' || res.status === 'duplicate' ? 'warn' : 'err';
      if (res.status === 'mismatch') {
        res.diffs.forEach(function (d) {
          var k = compCols.findIndex(function (c) { return (c.a === c.b ? c.a : c.a + ' / ' + c.b) === d.field; });
          if (k > -1) { eStyles[r + ',' + (5 + 2 * k)] = 'err'; eStyles[r + ',' + (6 + 2 * k)] = 'err'; }
        });
      }
    });
    if (exc.length === 1) exc.push(['No breaks', '', '', '', 'Both sides agree']);
    sheets.push({ name: 'Breaks', rows: exc, styles: eStyles, freezeRow: 1, widths: [18, 7, 7, 30, 46].concat(compCols.map(function () { return 16; }).concat(compCols.map(function () { return 16; }))) });

    // Matched pairs
    var mat = [['A row', 'B row', 'Key', 'How']];
    var mStyles = {};
    head(mat, mStyles, 0, 4);
    out.results.filter(function (r) { return r.status === 'matched' || r.status === 'matched_no_key'; }).forEach(function (res) {
      mat.push([res.a.row, res.b.row, cell(kA >= 0 ? res.a.r[kA] : ''), res.via === 'key' ? 'ID' : 'Values']);
    });
    sheets.push({ name: 'Matched', rows: mat, styles: mStyles, freezeRow: 1, widths: [7, 7, 30, 8] });

    // The two sides, with problem cells colored
    function sideSheet(S, which) {
      var rows = [S.headers.slice()].concat(S.rows);
      var styles = {};
      head(rows, styles, 0, S.headers.length);
      var first = S.firstDataRow || 2;
      function mark(rec, cols, style) {
        if (!rec) return;
        var r = rec.row - first + 1;
        cols.forEach(function (c) { if (c >= 0 && styles[r + ',' + c] !== 'err') styles[r + ',' + c] = style; });
      }
      var all = which === 'a' ? function (res) { return res.a; } : function (res) { return res.b; };
      var idx = which === 'a' ? ia : ib;
      out.results.forEach(function (res) {
        var rec = all(res);
        if (!rec) return;
        if (res.status === 'missing_in_b' || res.status === 'missing_in_a') mark(rec, S.headers.map(function (h, c) { return c; }), 'err');
        else if (res.status === 'mismatch') {
          res.diffs.forEach(function (d) { mark(rec, [which === 'a' ? d.colA : d.colB], 'err'); });
        } else if (res.status === 'matched_no_key') mark(rec, [kA >= 0 && which === 'a' ? kA : kB], 'warn');
      });
      out.duplicates.forEach(function (res) { mark(all(res), [which === 'a' ? kA : kB], 'warn'); });
      return { name: S.name, rows: rows, styles: styles, freezeRow: 1, widths: S.headers.map(function (h) { return Math.max(10, String(h).length + 4); }) };
    }
    sheets.push(sideSheet(A, 'a'), sideSheet(B, 'b'));
    return W.buildXlsx(sheets).bytes;
  }

  function breaksCsv(A, B, cfg, out) {
    var q = function (x) {
      var s = isDate(x) ? E.display(x) : String(v(x));
      if (/^[=+\-@]/.test(s)) s = "'" + s;
      return /[",;\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    var lines = [['Status', 'A row', 'B row', 'Explanation'].join(',')];
    out.results.filter(function (r) { return r.status !== 'matched'; }).concat(out.duplicates).forEach(function (res) {
      lines.push([LABEL[res.status], res.a ? res.a.row : '', res.b ? res.b.row : '', explain(res)].map(q).join(','));
    });
    return '﻿' + lines.join('\r\n');
  }

  return { build: build, breaksCsv: breaksCsv, explain: explain, LABEL: LABEL };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = ReconExport;
