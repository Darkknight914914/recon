(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var E = SheetEngine, W = SheetIO;
  var PAGE = 200;
  var EXCEL = ['xlsx', 'xlsm', 'xlsb', 'xls', 'xltx', 'xltm', 'xlt', 'ods', 'fods'], TEXT = ['csv', 'tsv', 'txt'];

  // side: { file, format, sheets: [{name, rows, headerRow, textInput}], idx, sample }
  var state = { a: null, b: null, cfg: null, out: null, filter: 'breaks', shown: PAGE, problem: '', note: '' };
  var downloads;

  function el(tag, attrs, kids) {
    var e = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (k === 'text') e.textContent = attrs[k];
      else if (k === 'class') e.className = attrs[k];
      else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), attrs[k]);
      else if (attrs[k] !== false && attrs[k] != null) e.setAttribute(k, attrs[k] === true ? '' : attrs[k]);
    }
    (kids || []).forEach(function (c) { if (c != null) e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
    return e;
  }
  function plural(n, w) { return n.toLocaleString() + ' ' + w + (n === 1 ? '' : 's'); }
  function store(k, v) {
    try { if (v === undefined) return JSON.parse(localStorage.getItem(k) || 'null'); localStorage.setItem(k, JSON.stringify(v)); }
    catch (e) { return null; }
  }
  function fmtNum(n) {
    if (typeof n !== 'number' || !isFinite(n)) return '';
    return n.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 8 });
  }

  // ---------------------------------------------------------------------------
  // sides
  // ---------------------------------------------------------------------------
  function sheetOf(s) { return s.sheets[s.idx]; }
  function sideData(s) {
    var sh = sheetOf(s);
    var width = sh.rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    return {
      name: s.sample ? s.label : (s.sheets.length > 1 || s.sameFile ? sh.name : s.file.replace(/\.[^.]+$/, '')),
      headers: W.namedHeaders(sh.rows[sh.headerRow] || [], width),
      rows: W.dataOnly(sh.rows, sh.headerRow),
      firstDataRow: sh.headerRow + 2,
      textInput: sh.textInput
    };
  }
  function signature() {
    return ['a', 'b'].map(function (k) { return sideData(state[k]).headers.join('|'); }).join('||');
  }

  function loadSample() {
    var s = ReconSample.build(new Date());
    state.a = { file: 'Sample', format: 'sample', sample: true, label: s.a.name, sheets: [{ name: s.a.name, rows: s.a.rows, headerRow: 0 }], idx: 0 };
    state.b = { file: 'Sample', format: 'sample', sample: true, label: s.b.name, sheets: [{ name: s.b.name, rows: s.b.rows, headerRow: 0 }], idx: 0 };
    state.note = '';
    remap(true);
  }

  function ext(n) { var m = /\.([^.]+)$/.exec(n || ''); return m ? m[1].toLowerCase() : ''; }

  function readFile(file) {
    var x = ext(file.name);
    if (TEXT.indexOf(x) > -1) {
      return file.arrayBuffer().then(function (buf) {
        var t = new TextDecoder('utf-8').decode(buf);
        if (t.indexOf('\uFFFD') > -1) { try { t = new TextDecoder('windows-1252').decode(buf); } catch (e) { /* keep */ } }
        return [{ name: file.name.replace(/\.[^.]+$/, '').slice(0, 31) || 'Data', rows: W.parseCsv(t).rows, textInput: true }];
      });
    }
    if (EXCEL.indexOf(x) > -1 || !x) {
      if (!window.XLSX) return Promise.reject(new Error('the Excel reader did not load. Check your connection and reload, or save the sheet as CSV'));
      return file.arrayBuffer().then(function (buf) {
        var wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellNF: true });
        var d1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
        return wb.SheetNames.map(function (n) { return { name: n, rows: W.sheetToRows(XLSX, wb.Sheets[n], d1904), textInput: false }; });
      });
    }
    return Promise.reject(new Error('it is not a spreadsheet. Use .xlsx, .xls, .ods or .csv'));
  }

  function openInto(which, file) {
    state.problem = '';
    renderBanner('Reading ' + file.name + '…');
    readFile(file).then(function (sheets) {
      sheets = sheets.filter(function (s) { return s.rows.some(function (r) { return r.some(function (v) { return v !== '' && v != null; }); }); });
      if (!sheets.length) throw new Error('it has no data');
      sheets.forEach(function (s) { s.headerRow = W.detectHeaderRow(s.rows); });
      var other = which === 'a' ? 'b' : 'a';
      state[which] = { file: file.name, format: ext(file.name), sheets: sheets, idx: 0 };
      state.note = '';
      // One workbook with two or more sheets: use the next sheet for the other side if it still shows the sample.
      if (sheets.length > 1 && state[other].sample) {
        state[other] = { file: file.name, format: ext(file.name), sheets: sheets, idx: 1, sameFile: true };
        state[which].sameFile = true;
        state.note = file.name + ' has ' + sheets.length + ' sheets, so Recon put "' + sheets[0].name + '" on side ' + which.toUpperCase() + ' and "' + sheets[1].name + '" on side ' + other.toUpperCase() + '. Change either below.';
      } else if (state[other].sample) {
        state.note = 'Side ' + which.toUpperCase() + ' loaded. Now open the file for side ' + other.toUpperCase() + '; until then it still shows sample data.';
      }
      remap(false);
    }).catch(function (e) {
      state.problem = 'Could not read "' + file.name + '": ' + (e && e.message || e) + '.';
      render();
    });
  }

  // ---------------------------------------------------------------------------
  // mapping + running
  // ---------------------------------------------------------------------------
  function remap(force) {
    var A = sideData(state.a), B = sideData(state.b);
    var saved = !force && !state.a.sample && !state.b.sample ? store('recon.map.' + signature()) : null;
    state.cfg = saved || Recon.suggestMapping(A, B);
    state.cfgFrom = saved ? 'saved' : 'suggested';
    if (!state.cfg.keys.length && !state.cfg.compares.length) {
      var ha = A.headers, hb = B.headers;
      if (ha.length && hb.length) state.cfg.keys.push({ a: ha[0], b: hb[0] });
    }
    state.filter = 'breaks'; state.shown = PAGE;
    run();
  }

  function run() {
    var A = sideData(state.a), B = sideData(state.b);
    state.A = A; state.B = B;
    state.out = Recon.reconcile(A, B, state.cfg);
    if (!state.a.sample && !state.b.sample) store('recon.map.' + signature(), state.cfg);
    render();
  }

  // ---------------------------------------------------------------------------
  // rendering
  // ---------------------------------------------------------------------------
  function render() { renderBanner(); renderSources(); renderSetup(); renderResults(); }

  function renderBanner(text) {
    var b = $('banner');
    b.innerHTML = '';
    if (text) { b.className = 'note'; b.textContent = text; return; }
    if (state.problem) { b.className = 'note problem'; b.textContent = state.problem; return; }
    if (state.note) { b.className = 'note sample'; b.textContent = state.note; return; }
    if (state.a.sample && state.b.sample) {
      b.className = 'note sample';
      b.appendChild(el('b', { text: 'Sample data: ' }));
      b.appendChild(document.createTextNode('an internal crypto ledger against a custody statement, with a short amount, a late settlement, a wrong asset, a missing and an unbooked transaction, a double booking and a row without an ID. Open your own files for side A and B, or one workbook that has both sheets.'));
      return;
    }
    b.className = 'note';
    b.textContent = state.cfgFrom === 'saved' ? 'Matching setup restored from the last time you reconciled files with these columns.' : 'Recon suggested how to match these files from their columns and values. Check “How to match” below.';
  }

  function renderSources() {
    ['a', 'b'].forEach(function (k) {
      var s = state[k], box = $(k === 'a' ? 'srcA' : 'srcB');
      var sh = sheetOf(s);
      box.innerHTML = '';
      var data = k === 'a' ? state.A : state.B;
      box.appendChild(el('div', { class: 'head' }, [
        el('span', { class: 'tag', text: 'Side ' + k.toUpperCase() }),
        el('button', { type: 'button', class: 'btn small', text: s.sample ? 'Open file for ' + k.toUpperCase() : 'Open another file', onclick: function () { $(k === 'a' ? 'fileA' : 'fileB').click(); } })
      ]));
      box.appendChild(el('div', { class: 'file', text: s.sample ? s.label + ' (sample)' : s.file }));
      var ctl = el('div', { class: 'ctl' });
      var sel = el('select', { id: 'sheet' + k, 'aria-label': 'Sheet for side ' + k.toUpperCase(), disabled: s.sheets.length < 2 ? true : false });
      s.sheets.forEach(function (x, i) { sel.appendChild(el('option', { value: i, text: x.name })); });
      sel.value = s.idx;
      sel.addEventListener('change', function () { s.idx = +this.value; state.note = ''; remap(false); });
      var hr = el('input', { type: 'number', id: 'hdr' + k, min: '1', value: sh.headerRow + 1, 'aria-label': 'Header row for side ' + k.toUpperCase() });
      hr.addEventListener('change', function () {
        var v = Math.round(+this.value) - 1;
        if (isNaN(v)) return;
        sh.headerRow = Math.max(0, Math.min(sh.rows.length - 1, v));
        remap(false);
      });
      ctl.appendChild(el('label', { class: 'lbl' }, ['Sheet', sel]));
      ctl.appendChild(el('label', { class: 'lbl' }, ['Header row', hr]));
      box.appendChild(ctl);
      var n = data ? data.rows.filter(function (r) { return r.some(function (v) { return v !== '' && v != null; }); }).length : 0;
      box.appendChild(el('div', { class: 'meta', text: plural(n, 'row') + ' · ' + plural(data ? data.headers.length : 0, 'column') }));
    });
  }

  function options(sel, headers, value, allowNone) {
    sel.innerHTML = '';
    if (allowNone) sel.appendChild(el('option', { value: '', text: '(none)' }));
    headers.forEach(function (h) { sel.appendChild(el('option', { value: h, text: h })); });
    sel.value = value || '';
    return sel;
  }

  function renderSetup() {
    var A = state.A, B = state.B, cfg = state.cfg;
    var keys = $('keys'); keys.innerHTML = '';
    cfg.keys.forEach(function (k, i) {
      var sa = options(el('select', { class: 'a', 'aria-label': 'Match column in A' }), A.headers, k.a);
      var sb = options(el('select', { class: 'b', 'aria-label': 'Match column in B' }), B.headers, k.b);
      sa.addEventListener('change', function () { k.a = this.value; run(); });
      sb.addEventListener('change', function () { k.b = this.value; run(); });
      keys.appendChild(el('div', { class: 'pair key' }, [sa, el('span', { class: 'eq', text: '=' }), sb,
        el('button', { type: 'button', class: 'icon-btn', text: '✕', 'aria-label': 'Remove', onclick: function () { cfg.keys.splice(i, 1); run(); } })]));
    });
    if (!cfg.keys.length) keys.appendChild(el('div', { class: 'muted', style: 'font-size:14px', text: 'No ID column: rows are paired when all compared values agree.' }));
    $('keyHint').textContent = cfg.keys.length ? 'IDs are matched ignoring case, spaces and a leading 0x.' : '';

    var cmps = $('compares'); cmps.innerHTML = '';
    cfg.compares.forEach(function (c, i) {
      var sa = options(el('select', { class: 'a', 'aria-label': 'Column in A' }), A.headers, c.a);
      var sb = options(el('select', { class: 'b', 'aria-label': 'Column in B' }), B.headers, c.b);
      var type = el('select', { 'aria-label': 'Compare as' });
      [['number', 'Number'], ['date', 'Date'], ['text', 'Text']].forEach(function (t) { type.appendChild(el('option', { value: t[0], text: t[1] })); });
      type.value = c.type;
      var tol = el('input', { type: 'number', min: '0', step: 'any', value: c.tol || 0, 'aria-label': c.type === 'date' ? 'Tolerance in days' : 'Tolerance', title: c.type === 'date' ? 'Allowed difference in days' : 'Allowed difference', disabled: c.type === 'text' ? true : false });
      var sign = el('select', { 'aria-label': 'Sign convention', title: 'Bank exports often show a payment as −100 where the ledger books +100', disabled: c.type !== 'number' ? true : false });
      [['same', 'Same sign'], ['flip', 'Opposite sign'], ['abs', 'Ignore sign']].forEach(function (t) { sign.appendChild(el('option', { value: t[0], text: t[1] })); });
      sign.value = c.flip ? 'flip' : c.abs ? 'abs' : 'same';
      sa.addEventListener('change', function () { c.a = this.value; run(); });
      sb.addEventListener('change', function () { c.b = this.value; run(); });
      type.addEventListener('change', function () { c.type = this.value; c.tol = this.value === 'date' ? 1 : 0; run(); });
      tol.addEventListener('change', function () { c.tol = Math.max(0, +this.value || 0); run(); });
      sign.addEventListener('change', function () {
        c.flip = this.value === 'flip'; c.abs = this.value === 'abs';
        if (cfg.amount && cfg.amount.a === c.a && cfg.amount.b === c.b) { cfg.amount.flip = c.flip; cfg.amount.abs = c.abs; }
        run();
      });
      cmps.appendChild(el('div', { class: 'pair' }, [sa, el('span', { class: 'eq', text: '=' }), sb, type,
        el('label', { class: 'lbl tol', style: 'display:flex;gap:4px;align-items:center' + (c.type === 'text' ? ';visibility:hidden' : '') }, [c.type === 'date' ? '± days' : '±', tol]),
        c.type === 'number' ? sign : el('span', { 'aria-hidden': 'true' }),
        el('button', { type: 'button', class: 'icon-btn', text: '✕', 'aria-label': 'Remove', onclick: function () { cfg.compares.splice(i, 1); run(); } })]));
    });

    options($('amtA'), A.headers, cfg.amount && cfg.amount.a, true);
    options($('amtB'), B.headers, cfg.amount && cfg.amount.b, true);
    options($('grpA'), A.headers, cfg.group && cfg.group.a, true);
    options($('grpB'), B.headers, cfg.group && cfg.group.b, true);
    $('fallback').checked = !!cfg.fallback;
  }

  var FILTERS = [
    ['breaks', 'All breaks', function (r) { return r.status !== 'matched'; }],
    ['mismatch', 'Value breaks', function (r) { return r.status === 'mismatch'; }],
    ['missing_in_b', 'Missing in B', function (r) { return r.status === 'missing_in_b'; }],
    ['missing_in_a', 'Missing in A', function (r) { return r.status === 'missing_in_a'; }],
    ['duplicate', 'Duplicates', function (r) { return r.status === 'duplicate'; }],
    ['matched_no_key', 'Paired without ID', function (r) { return r.status === 'matched_no_key'; }],
    ['matched', 'Matched', function (r) { return r.status === 'matched'; }]
  ];

  function listFor(f) {
    var out = state.out;
    var all = out.results.concat(out.duplicates);
    var def = FILTERS.filter(function (x) { return x[0] === f; })[0];
    return all.filter(def[2]);
  }

  function renderResults() {
    var out = state.out;
    var pb = $('problems');
    if (out.problems && out.problems.length) {
      pb.hidden = false; pb.textContent = out.problems.join(' ');
      ['tiles', 'totals', 'items', 'filters', 'totFoot'].forEach(function (id) { $(id).innerHTML = ''; });
      return;
    }
    pb.hidden = true;
    var st = out.stats;
    var tiles = $('tiles'); tiles.innerHTML = '';
    [['matched', st.matched + st.matchedNoKey, 'matched', 'ok'],
     ['mismatch', st.mismatch, 'value breaks', st.mismatch ? 'err' : 'ok'],
     ['missing_in_b', st.missingInB, 'in A, missing in B', st.missingInB ? 'err' : 'ok'],
     ['missing_in_a', st.missingInA, 'in B, missing in A', st.missingInA ? 'err' : 'ok'],
     ['duplicate', st.duplicatesA + st.duplicatesB, 'duplicate IDs', st.duplicatesA + st.duplicatesB ? 'warn' : 'ok']
    ].forEach(function (t) {
      tiles.appendChild(el('button', { type: 'button', class: 'tile ' + t[3], 'aria-pressed': String(state.filter === t[0]),
        onclick: function () { state.filter = t[0]; state.shown = PAGE; renderResults(); } }, [el('b', { text: t[1].toLocaleString() }), el('span', { text: t[2] })]));
    });

    // totals
    var tw = $('totals'); tw.innerHTML = '';
    var foot = $('totFoot');
    if (out.totals && out.totals.length) {
      var t = el('table', { class: 'totals' });
      t.appendChild(el('thead', {}, [el('tr', {}, [el('th', { text: state.cfg.group ? state.cfg.group.a : '' }), el('th', { text: 'A' }), el('th', { text: 'B' }), el('th', { text: 'Difference' })])]));
      var tb = el('tbody');
      out.totals.forEach(function (x) {
        tb.appendChild(el('tr', {}, [el('td', { text: x.group }), el('td', { text: fmtNum(x.a) }), el('td', { text: fmtNum(x.b) }),
          el('td', { class: x.diff === 0 ? 'good' : 'bad', text: x.diff === 0 ? '0 ✓' : (x.diff > 0 ? '+' : '') + fmtNum(x.diff) })]));
      });
      t.appendChild(tb);
      tw.appendChild(t);
      var off = out.totals.filter(function (x) { return x.diff !== 0; }).length;
      foot.textContent = off ? plural(off, 'group') + ' out of balance. The breaks list shows the rows behind each difference.' : 'All groups balance.';
    } else {
      foot.textContent = 'Choose an amount column under “Totals of” to compare totals.';
    }

    // filters
    var fb = $('filters'); fb.innerHTML = '';
    FILTERS.forEach(function (f) {
      var n = listFor(f[0]).length;
      if (!n && f[0] !== 'breaks' && f[0] !== state.filter) return;
      fb.appendChild(el('button', { type: 'button', 'aria-pressed': String(state.filter === f[0]), text: f[1] + ' ' + n.toLocaleString(),
        onclick: function () { state.filter = f[0]; state.shown = PAGE; renderResults(); } }));
    });

    var ul = $('items'); ul.innerHTML = '';
    var items = listFor(state.filter);
    if (!items.length) {
      ul.appendChild(el('li', { class: 'empty', text: state.filter === 'breaks' ? 'No breaks. Both sides agree.' : 'Nothing in this group.' }));
    } else {
      items.slice(0, state.shown).forEach(function (res) { ul.appendChild(item(res)); });
      if (items.length > state.shown) ul.appendChild(el('li', { class: 'more' }, [el('button', { type: 'button', class: 'btn', text: 'Show ' + Math.min(PAGE, items.length - state.shown) + ' more of ' + (items.length - state.shown).toLocaleString(), onclick: function () { state.shown += PAGE; renderResults(); } })]));
    }
    renderDownloads();
  }

  function keyText(res) {
    var k = state.cfg.keys[0];
    if (!k) return '';
    var ia = state.A.headers.indexOf(k.a), ib = state.B.headers.indexOf(k.b);
    var v = res.a && ia >= 0 ? res.a.r[ia] : res.b && ib >= 0 ? res.b.r[ib] : '';
    return E.display(v);
  }

  function item(res) {
    var cls = res.status === 'matched' ? 'ok' : res.status === 'matched_no_key' || res.status === 'duplicate' ? 'warn' : '';
    var li = el('li', { class: 'item ' + cls });
    var key = keyText(res);
    li.appendChild(el('div', { class: 'line1' }, [
      el('span', { class: 'badge', text: ReconExport.LABEL[res.status] }),
      key ? el('span', { class: 'key', text: key }) : el('span', { class: 'key muted', text: 'no ID' }),
      el('span', { class: 'rows', text: [res.a ? 'A row ' + res.a.row : '', res.b ? 'B row ' + res.b.row : ''].filter(Boolean).join(' · ') })
    ]));
    var why = ReconExport.explain(res);
    if (why && res.status !== 'mismatch') li.appendChild(el('div', { class: 'why', text: why }));
    // side-by-side compare of the compared columns
    var cmps = state.cfg.compares;
    if (cmps.length) {
      var t = el('table');
      t.appendChild(el('tr', {}, [el('th', { text: '' }), el('th', { text: 'A' }), el('th', { text: 'B' }), el('th', { text: '' })]));
      cmps.forEach(function (c) {
        var ia = state.A.headers.indexOf(c.a), ib = state.B.headers.indexOf(c.b);
        var field = c.a === c.b ? c.a : c.a + ' / ' + c.b;
        var d = res.diffs.filter(function (x) { return x.field === field; })[0];
        t.appendChild(el('tr', { class: d ? 'diff' : '' }, [
          el('th', { text: field }),
          el('td', { class: 'a', text: res.a && ia >= 0 ? E.display(res.a.r[ia]) : '—' }),
          el('td', { class: 'b', text: res.b && ib >= 0 ? E.display(res.b.r[ib]) : '—' }),
          el('td', { class: 'd', text: d ? d.note : '' })
        ]));
      });
      li.appendChild(el('div', { class: 'cmp' }, [t]));
    }
    return li;
  }

  // ---------------------------------------------------------------------------
  // downloads
  // ---------------------------------------------------------------------------
  function getDownloads() {
    if (downloads !== undefined) return Promise.resolve(downloads);
    if (!window.claude || typeof window.claude.use !== 'function') { downloads = 'local'; return Promise.resolve(downloads); }
    return window.claude.use('downloads').then(function (d) { downloads = d || null; renderDownloads(); return downloads; }, function () { downloads = null; renderDownloads(); return null; });
  }
  function renderDownloads() {
    var none = downloads === null || !!(state.out.problems && state.out.problems.length);
    $('dlXlsx').disabled = none; $('dlCsv').disabled = none;
    if (downloads === null) $('dlHint').textContent = 'Downloads are not available in this view. Open the page in a browser to save the workbook.';
  }
  function save(name, data, type) {
    return getDownloads().then(function (d) {
      if (!d) return;
      if (d === 'local') {
        var url = URL.createObjectURL(new Blob([data], { type: type }));
        var a = el('a', { href: url, download: name });
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
        return;
      }
      return d.save({ filename: name, data: new Blob([data], { type: type }) }).catch(function (e) {
        if (e && e.code === 'declined') return;
        $('dlHint').textContent = 'The file could not be saved here (' + (e && e.code || 'error') + ').';
      });
    });
  }
  function baseName() {
    if (state.a.sample && state.b.sample) return 'sample-recon';
    var n = function (s) { return (s.sameFile ? sheetOf(s).name : s.file.replace(/\.[^.]+$/, '')).replace(/[^\w\- ]+/g, '').trim().slice(0, 30) || 'side'; };
    return 'recon-' + n(state.a) + '-vs-' + n(state.b);
  }
  $('dlXlsx').addEventListener('click', function () {
    save(baseName() + '.xlsx', ReconExport.build(state.A, state.B, state.cfg, state.out, { ranAt: new Date() }), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  });
  $('dlCsv').addEventListener('click', function () {
    save(baseName() + '-breaks.csv', ReconExport.breaksCsv(state.A, state.B, state.cfg, state.out), 'text/csv;charset=utf-8');
  });

  // ---------------------------------------------------------------------------
  // events
  // ---------------------------------------------------------------------------
  ['a', 'b'].forEach(function (k) {
    $(k === 'a' ? 'fileA' : 'fileB').addEventListener('change', function () { if (this.files[0]) openInto(k, this.files[0]); this.value = ''; });
  });
  $('autoMap').addEventListener('click', function () { state.note = ''; remap(true); });
  $('addKey').addEventListener('click', function () { state.cfg.keys.push({ a: state.A.headers[0], b: state.B.headers[0] }); run(); });
  $('addCmp').addEventListener('click', function () { state.cfg.compares.push({ a: state.A.headers[0], b: state.B.headers[0], type: 'text', tol: 0 }); run(); });
  $('fallback').addEventListener('change', function () { state.cfg.fallback = this.checked; run(); });
  function totalsChanged() {
    var aa = $('amtA').value, ab = $('amtB').value, ga = $('grpA').value, gb = $('grpB').value;
    var match = state.cfg.compares.filter(function (c) { return c.a === aa && c.b === ab; })[0] || {};
    state.cfg.amount = aa && ab ? { a: aa, b: ab, abs: !!match.abs, flip: !!match.flip } : null;
    state.cfg.group = ga && gb ? { a: ga, b: gb } : null;
    run();
  }
  ['amtA', 'amtB', 'grpA', 'grpB'].forEach(function (id) { $(id).addEventListener('change', totalsChanged); });

  // Drop a file: left half of the window → A, right half → B.
  window.addEventListener('dragover', function (e) { e.preventDefault(); });
  window.addEventListener('drop', function (e) {
    e.preventDefault();
    var files = e.dataTransfer && e.dataTransfer.files;
    if (!files || !files.length) return;
    if (files.length >= 2) { openInto('a', files[0]); setTimeout(function () { openInto('b', files[1]); }, 50); return; }
    openInto(e.clientX < window.innerWidth / 2 ? 'a' : 'b', files[0]);
  });

  loadSample();
  if (window.claude && typeof window.claude.use === 'function') getDownloads(); else downloads = 'local';
})();
