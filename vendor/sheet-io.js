/**
 * Sheet Web — file handling that doesn't need the DOM:
 * CSV parsing, SheetJS sheet → rows, header detection, rules-sheet import,
 * and a small dependency-free .xlsx writer (highlighted copy + report).
 */
var SheetIO = (function () {
  'use strict';

  var RULES_SHEET = 'Sheet Rules';
  var REPORT_SHEET = 'Sheet Report';
  var RULE_HEADERS = ['Enabled', 'Sheet', 'Column', 'Check', 'Parameter', 'Severity', 'Description'];

  // ---------------------------------------------------------------------------
  // CSV / TSV
  // ---------------------------------------------------------------------------
  function detectDelimiter(text) {
    var first = text.split(/\r?\n/).slice(0, 5).join('\n');
    var best = ',', bestCount = -1;
    [',', ';', '\t', '|'].forEach(function (d) {
      var inQ = false, n = 0;
      for (var i = 0; i < first.length; i++) {
        var ch = first.charAt(i);
        if (ch === '"') inQ = !inQ;
        else if (!inQ && ch === d) n++;
      }
      if (n > bestCount) { best = d; bestCount = n; }
    });
    return best;
  }

  function parseCsv(text) {
    text = String(text).replace(/^\uFEFF/, '');
    var sepLine = text.match(/^sep=(.)\r?\n/i);
    var d;
    if (sepLine) { d = sepLine[1]; text = text.slice(sepLine[0].length); }
    else d = detectDelimiter(text);
    var rows = [], row = [], field = '', inQ = false, i = 0, n = text.length;
    while (i < n) {
      var ch = text.charAt(i);
      if (inQ) {
        if (ch === '"') {
          if (text.charAt(i + 1) === '"') { field += '"'; i += 2; continue; }
          inQ = false; i++; continue;
        }
        field += ch; i++; continue;
      }
      if (ch === '"' && field === '') { inQ = true; i++; continue; }
      if (ch === d) { row.push(field); field = ''; i++; continue; }
      if (ch === '\r' || ch === '\n') {
        row.push(field); rows.push(row); row = []; field = '';
        if (ch === '\r' && text.charAt(i + 1) === '\n') i++;
        i++; continue;
      }
      field += ch; i++;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return { rows: rows, delimiter: d };
  }

  // ---------------------------------------------------------------------------
  // SheetJS worksheet → 2D array of JS values (Date for date cells)
  // ---------------------------------------------------------------------------
  function serialToDate(XLSX, v, date1904) {
    if (XLSX && XLSX.SSF && XLSX.SSF.parse_date_code) {
      var p = XLSX.SSF.parse_date_code(v, { date1904: !!date1904 });
      if (p) return new Date(p.y, p.m - 1, p.d, p.H || 0, p.M || 0, Math.floor(p.S || 0));
    }
    var ms = Math.round((v - (date1904 ? -1462 : 0)) * 86400000);
    var base = new Date(Date.UTC(1899, 11, 30) + ms);
    return new Date(base.getUTCFullYear(), base.getUTCMonth(), base.getUTCDate(), base.getUTCHours(), base.getUTCMinutes(), base.getUTCSeconds());
  }

  function isDateFormat(XLSX, z) {
    if (!z || z === 'General') return false;
    if (XLSX && XLSX.SSF && XLSX.SSF.is_date) {
      try { return XLSX.SSF.is_date(z); } catch (e) { /* fall through */ }
    }
    return /(^|[^"\\])[dmyhs]/i.test(String(z).replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, ''));
  }

  function sheetToRows(XLSX, ws, date1904) {
    if (!ws || !ws['!ref']) return [];
    var range = XLSX.utils.decode_range(ws['!ref']);
    var rows = [];
    // Always start at A1 so row/column numbers match what Excel shows.
    for (var r = 0; r <= range.e.r; r++) {
      var row = [];
      for (var c = 0; c <= range.e.c; c++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        var v = '';
        if (cell) {
          if (cell.t === 'n') v = isDateFormat(XLSX, cell.z) ? serialToDate(XLSX, cell.v, date1904) : cell.v;
          else if (cell.t === 'd') v = cell.v instanceof Date ? cell.v : new Date(cell.v);
          else if (cell.t === 'b') v = !!cell.v;
          else if (cell.t === 'e') v = cell.w || '#ERROR!';
          else if (cell.t === 's' || cell.t === 'str') v = cell.v == null ? '' : String(cell.v);
          else if (cell.v != null) v = cell.v;
        }
        row.push(v);
      }
      rows.push(row);
    }
    // drop trailing empty rows
    while (rows.length && rows[rows.length - 1].every(blank)) rows.pop();
    return rows;
  }

  function blank(v) { return v === '' || v === null || v === undefined || (typeof v === 'string' && v.trim() === ''); }

  /**
   * The header is the widest mostly-text row near the top (title rows such as
   * "Sales report Q3" or "Generated: 18.10.2026" are narrower and get skipped).
   */
  function detectHeaderRow(rows) {
    var best = -1, bestWidth = 0, cands = [];
    for (var i = 0; i < Math.min(rows.length, 30); i++) {
      var filled = rows[i].filter(function (v) { return !blank(v); });
      if (filled.length < 2) continue;
      var texty = filled.filter(function (v) { return typeof v === 'string' && isNaN(Number(String(v).replace(',', '.'))); }).length;
      if (texty < filled.length * 0.6) continue;
      cands.push({ i: i, w: filled.length });
      if (filled.length > bestWidth) bestWidth = filled.length;
    }
    for (var k = 0; k < cands.length; k++) if (cands[k].w >= bestWidth * 0.8) { best = cands[k].i; break; }
    return best < 0 ? 0 : best;
  }

  /** Column names for checking: blanks become "Column F", repeats become "Amount (2)". */
  function namedHeaders(row, width) {
    var seen = {}, out = [];
    for (var c = 0; c < Math.max(width || 0, (row || []).length); c++) {
      var h = row && row[c] != null ? SheetEngine.display(row[c]).replace(/\s+/g, ' ').trim() : '';
      if (!h) h = 'Column ' + columnLetter(c + 1);
      var key = h.toLowerCase(), base = h, n = 2;
      while (seen[key]) { h = base + ' (' + n++ + ')'; key = h.toLowerCase(); }
      seen[key] = true;
      out.push(h);
    }
    return out;
  }

  var TOTAL_RE = /^(grand\s+)?(total|totals|sum|summe|gesamt|gesamtsumme|subtotal|sub-total|zwischensumme|teilsumme|endsumme|insgesamt|overall)\b/i;

  /**
   * Splits a report into data rows and the rows around them:
   *  kind[i] = 'data' | 'total' | 'subtotal' | 'note' | 'blank' (rows after the header only).
   */
  function classifyRows(rows, headerRow) {
    var width = (rows[headerRow] || []).filter(function (v) { return !blank(v); }).length;
    var kinds = [];
    for (var i = headerRow + 1; i < rows.length; i++) {
      var r = rows[i] || [];
      var filled = r.filter(function (v) { return !blank(v); });
      if (!filled.length) { kinds.push('blank'); continue; }
      var firstText = filled.filter(function (v) { return typeof v === 'string'; })[0];
      if (firstText && TOTAL_RE.test(String(firstText).trim()) && filled.indexOf(firstText) <= 2) {
        kinds.push(/sub|zwischen|teil/i.test(firstText) ? 'subtotal' : 'total');
        continue;
      }
      if (width >= 3 && filled.length === 1 && typeof filled[0] === 'string' && filled[0].length > 20 && r.indexOf(filled[0]) === 0) { kinds.push('note'); continue; }
      kinds.push('data');
    }
    return kinds;
  }

  /** Rows for the engine: only data rows keep their values; others become empty (row numbers stay right). */
  function dataOnly(rows, headerRow) {
    var kinds = classifyRows(rows, headerRow);
    return rows.slice(headerRow + 1).map(function (r, k) { return kinds[k] === 'data' ? r : []; });
  }

  /** Reads a cell like "F112" from the same sheet (used for "column total = Total row"). */
  function cellValue(rows, ref) {
    var m = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(String(ref).trim());
    if (!m) return undefined;
    var col = 0;
    m[1].toUpperCase().split('').forEach(function (ch) { col = col * 26 + ch.charCodeAt(0) - 64; });
    var r = rows[+m[2] - 1];
    return r ? r[col - 1] : undefined;
  }

  /** Rules stored in a "Sheet Rules" sheet (same layout as the Google Sheets add-on). */
  function rulesFromRows(rows) {
    if (!rows || rows.length < 2) return [];
    var head = rows[0].map(function (h) { return String(h).trim().toLowerCase(); });
    var ix = {};
    RULE_HEADERS.forEach(function (h) { ix[h.toLowerCase()] = head.indexOf(h.toLowerCase()); });
    if (ix.check < 0) return [];
    var get = function (r, k) { return ix[k] >= 0 && r[ix[k]] != null ? r[ix[k]] : ''; };
    var out = [];
    rows.slice(1).forEach(function (r) {
      var check = String(get(r, 'check')).trim().toLowerCase();
      if (!check) return;
      var en = get(r, 'enabled');
      out.push({
        enabled: !(en === false || /^(false|no|nein|0)$/i.test(String(en).trim())),
        sheet: String(get(r, 'sheet')).trim(),
        column: String(get(r, 'column')).trim(),
        check: check,
        param: SheetEngine.display(get(r, 'parameter')).trim(),
        severity: String(get(r, 'severity')).trim().toLowerCase() === 'warning' ? 'warning' : 'error',
        description: String(get(r, 'description')).trim()
      });
    });
    return out;
  }

  function columnLetter(n) {          // 1-based
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  // ---------------------------------------------------------------------------
  // Minimal .xlsx writer (store-only ZIP, inline strings, a few fills)
  // ---------------------------------------------------------------------------
  var CRC_TABLE = (function () {
    var t = new Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function utf8(s) {
    if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(s);
    var out = [], i = 0;
    s = unescape(encodeURIComponent(s));
    for (; i < s.length; i++) out.push(s.charCodeAt(i));
    return new Uint8Array(out);
  }

  function zip(files) {               // files: [{name, data: Uint8Array}]
    var chunks = [], central = [], offset = 0;
    var now = new Date();
    var dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    var dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    function u16(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF); }
    function u32(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF); }
    files.forEach(function (f) {
      var name = utf8(f.name), crc = crc32(f.data), size = f.data.length;
      var h = [];
      u32(h, 0x04034B50); u16(h, 20); u16(h, 0x0800); u16(h, 0); u16(h, dosTime); u16(h, dosDate);
      u32(h, crc); u32(h, size); u32(h, size); u16(h, name.length); u16(h, 0);
      chunks.push(new Uint8Array(h), name, f.data);
      var c = [];
      u32(c, 0x02014B50); u16(c, 20); u16(c, 20); u16(c, 0x0800); u16(c, 0); u16(c, dosTime); u16(c, dosDate);
      u32(c, crc); u32(c, size); u32(c, size); u16(c, name.length); u16(c, 0); u16(c, 0); u16(c, 0); u16(c, 0);
      u32(c, 0); u32(c, offset);
      central.push(new Uint8Array(c), name);
      offset += h.length + name.length + size;
    });
    var cdSize = central.reduce(function (s, a) { return s + a.length; }, 0);
    var e = [];
    u32(e, 0x06054B50); u16(e, 0); u16(e, 0); u16(e, files.length); u16(e, files.length);
    u32(e, cdSize); u32(e, offset); u16(e, 0);
    var all = chunks.concat(central, [new Uint8Array(e)]);
    var total = all.reduce(function (s, a) { return s + a.length; }, 0);
    var out = new Uint8Array(total), pos = 0;
    all.forEach(function (a) { out.set(a, pos); pos += a.length; });
    return out;
  }

  function xmlEsc(s) {
    return String(s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // style ids (see STYLES)
  var S = { plain: 0, head: 1, err: 2, warn: 3, date: 4, dateErr: 5, dateWarn: 6, link: 7, cfg: 8 };
  var STYLES =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>' +
    '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="11"/><name val="Calibri"/></font>' +
    '<font><u/><sz val="11"/><color rgb="FF1F5F99"/><name val="Calibri"/></font></fonts>' +
    '<fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFAD4D0"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFEBB0"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFE3EAF3"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFEADCF5"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="9">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>' +
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>';

  function isDate(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }

  function dateSerial(d) {
    return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  /**
   * sheet: { name, rows, styles?: {"r,c": styleKey}, freezeRow?: n, widths?: [chars],
   *          links?: [{r, c, location}] }   (r, c zero-based)
   */
  function sheetXml(sheet) {
    var rows = sheet.rows || [];
    var maxC = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    var x = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'];
    if (rows.length && maxC) x.push('<dimension ref="A1:' + columnLetter(maxC) + rows.length + '"/>');
    x.push('<sheetViews><sheetView workbookViewId="0">');
    if (sheet.freezeRow) {
      x.push('<pane ySplit="' + sheet.freezeRow + '" topLeftCell="A' + (sheet.freezeRow + 1) + '" activePane="bottomLeft" state="frozen"/>');
    }
    x.push('</sheetView></sheetViews>');
    if (sheet.widths && sheet.widths.length) {
      x.push('<cols>');
      sheet.widths.forEach(function (w, i) {
        x.push('<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + Math.max(6, Math.min(70, w)) + '" customWidth="1"/>');
      });
      x.push('</cols>');
    }
    x.push('<sheetData>');
    var styles = sheet.styles || {};
    rows.forEach(function (row, r) {
      var cells = [];
      row.forEach(function (v, c) {
        var ref = columnLetter(c + 1) + (r + 1);
        var st = styles[r + ',' + c];
        var sid = st ? S[st] : 0;
        if (v === '' || v === null || v === undefined) {
          if (sid) cells.push('<c r="' + ref + '" s="' + sid + '"/>');
          return;
        }
        if (isDate(v)) {
          var ds = st === 'err' ? S.dateErr : st === 'warn' ? S.dateWarn : S.date;
          cells.push('<c r="' + ref + '" s="' + ds + '"><v>' + dateSerial(v) + '</v></c>');
        } else if (typeof v === 'number' && isFinite(v)) {
          cells.push('<c r="' + ref + '"' + (sid ? ' s="' + sid + '"' : '') + '><v>' + v + '</v></c>');
        } else if (typeof v === 'boolean') {
          cells.push('<c r="' + ref + '" t="b"' + (sid ? ' s="' + sid + '"' : '') + '><v>' + (v ? 1 : 0) + '</v></c>');
        } else {
          var s = String(v);
          if (s.length > 32767) s = s.slice(0, 32767);
          cells.push('<c r="' + ref + '" t="inlineStr"' + (sid ? ' s="' + sid + '"' : '') + '><is><t xml:space="preserve">' + xmlEsc(s) + '</t></is></c>');
        }
      });
      x.push('<row r="' + (r + 1) + '">' + cells.join('') + '</row>');
    });
    x.push('</sheetData>');
    if (sheet.links && sheet.links.length) {
      x.push('<hyperlinks>');
      sheet.links.forEach(function (l) {
        x.push('<hyperlink ref="' + columnLetter(l.c + 1) + (l.r + 1) + '" location="' + xmlEsc(l.location) + '" display="' + xmlEsc(l.display || '') + '"/>');
      });
      x.push('</hyperlinks>');
    }
    x.push('</worksheet>');
    return x.join('');
  }

  function safeSheetName(name, used) {
    var base = String(name || 'Sheet').replace(/[\[\]:*?\/\\]/g, ' ').replace(/^'+|'+$/g, '').trim().slice(0, 31) || 'Sheet';
    var n = base, k = 2;
    while (used[n.toLowerCase()]) { var suf = ' (' + k++ + ')'; n = base.slice(0, 31 - suf.length) + suf; }
    used[n.toLowerCase()] = true;
    return n;
  }

  function quoteSheet(name) { return "'" + name.replace(/'/g, "''") + "'"; }

  function buildXlsx(sheets) {
    var used = {};
    sheets = sheets.map(function (s) { var o = {}; for (var k in s) o[k] = s[k]; o.name = safeSheetName(s.name, used); return o; });
    var files = [];
    var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      sheets.map(function (s, i) {
        return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      }).join('') + '</Types>';
    files.push({ name: '[Content_Types].xml', data: utf8(ct) });
    files.push({ name: '_rels/.rels', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>') });
    files.push({ name: 'xl/workbook.xml', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<bookViews><workbookView/></bookViews><sheets>' +
      sheets.map(function (s, i) { return '<sheet name="' + xmlEsc(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join('') +
      '</sheets></workbook>') });
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map(function (s, i) {
        return '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
      }).join('') +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>') });
    files.push({ name: 'xl/styles.xml', data: utf8(STYLES) });
    sheets.forEach(function (s, i) { files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: utf8(sheetXml(s)) }); });
    return { bytes: zip(files), names: sheets.map(function (s) { return s.name; }) };
  }

  function widthsFor(rows, maxRows) {
    var w = [];
    rows.slice(0, maxRows || 200).forEach(function (r) {
      r.forEach(function (v, c) {
        var len = isDate(v) ? 10 : String(v == null ? '' : v).length;
        w[c] = Math.max(w[c] || 8, len + 2);
      });
    });
    return w;
  }

  /**
   * Builds the marked-up workbook.
   * book: { sheets: [{ name, rows, headerRow }] }, rules: [...], issues: [...] from the engine.
   */
  function buildReport(book, rules, issues) {
    var marks = {};                                   // sheetName -> {"r,c": style}
    issues.forEach(function (i) {
      if (!i.row || i.colIndex < 0) return;
      var m = marks[i.sheet] = marks[i.sheet] || {};
      var key = (i.row - 1) + ',' + i.colIndex;
      if (m[key] !== 'err') m[key] = i.severity === 'warning' ? 'warn' : 'err';
    });

    var used = {}, finalNames = {};
    var out = [];
    book.sheets.forEach(function (s) {
      if (s.name === RULES_SHEET || s.name === REPORT_SHEET) return;
      var nm = safeSheetName(s.name, used);
      finalNames[s.name] = nm;
      var styles = {};
      var m = marks[s.name] || {};
      for (var k in m) styles[k] = m[k];
      var hr = s.headerRow || 0;
      (s.rows[hr] || []).forEach(function (h, c) { if (!styles[hr + ',' + c]) styles[hr + ',' + c] = 'head'; });
      out.push({ name: nm, rows: s.rows, styles: styles, freezeRow: hr + 1, widths: widthsFor(s.rows) });
    });

    var rep = [['Severity', 'Sheet', 'Cell', 'Column', 'Value', 'Problem', 'Rule']];
    var repStyles = {}, links = [];
    for (var c = 0; c < 7; c++) repStyles['0,' + c] = 'head';
    issues.forEach(function (i, k) {
      var r = k + 1;
      var sheet = finalNames[i.sheet] || i.sheet;
      var cell = i.row ? (i.colIndex >= 0 ? columnLetter(i.colIndex + 1) : 'A') + i.row : '';
      rep.push([i.severity === 'config' ? 'rule problem' : i.severity, sheet, cell, i.column || '',
        isDate(i.value) ? i.value : (i.value == null ? '' : i.value), i.message, i.rule]);
      repStyles[r + ',0'] = i.severity === 'warning' ? 'warn' : i.severity === 'config' ? 'cfg' : 'err';
      if (cell && finalNames[i.sheet]) {
        repStyles[r + ',2'] = 'link';
        links.push({ r: r, c: 2, location: quoteSheet(sheet) + '!' + cell, display: cell });
      }
    });
    if (!issues.length) rep.push(['', '', '', '', '', 'No problems found', '']);
    out.push({ name: REPORT_SHEET, rows: rep, styles: repStyles, freezeRow: 1, widths: [12, 16, 8, 18, 22, 60, 40], links: links });

    var rr = [RULE_HEADERS.slice()];
    var rStyles = {};
    for (var j = 0; j < RULE_HEADERS.length; j++) rStyles['0,' + j] = 'head';
    rules.forEach(function (r) {
      rr.push([r.enabled !== false, r.sheet || '', r.column || '', r.check, r.param == null ? '' : String(r.param), r.severity || 'error', r.description || '']);
    });
    out.push({ name: RULES_SHEET, rows: rr, styles: rStyles, freezeRow: 1, widths: [9, 16, 20, 16, 22, 10, 44] });

    return buildXlsx(out).bytes;
  }

  // ---------------------------------------------------------------------------
  // Starter rules from the data itself, so a new file gets checked right away.
  // Each suggestion is something the column almost always does; the rare rows
  // that break it are what the user wants to see.
  // ---------------------------------------------------------------------------
  /** "ACCT_856_976" → skeleton "A_9_9" with per-run lengths, for format consistency. */
  function shapeRuns(s) {
    var runs = [], re = /([A-Z]+)|([a-z]+)|(\d+)|([^A-Za-z\d])/g, m;
    while ((m = re.exec(s)) !== null) {
      if (m[1]) runs.push(['A', m[1].length]);
      else if (m[2]) runs.push(['a', m[2].length]);
      else if (m[3]) runs.push(['9', m[3].length]);
      else runs.push([m[4], 1]);
    }
    // merge repeated identical punctuation into one run so "a  b" and "a b" differ
    return runs;
  }
  function escRe(ch) { return ch.replace(/[.*+?^${}()|[\]\\\/]/g, '\\$&'); }

  /** Regex covering the dominant format of a text column, or null. */
  function formatPattern(texts) {
    if (texts.length < 10) return null;
    var groups = {};
    texts.forEach(function (t) {
      if (t.length > 40) return;
      var runs = shapeRuns(t);
      var key = runs.map(function (r) { return r[0]; }).join('');
      var g = groups[key] || (groups[key] = { n: 0, runs: runs.map(function (r) { return [r[0], r[1], r[1]]; }) });
      g.n++;
      runs.forEach(function (r, i) { g.runs[i][1] = Math.min(g.runs[i][1], r[1]); g.runs[i][2] = Math.max(g.runs[i][2], r[1]); });
    });
    var keys = Object.keys(groups);
    if (!keys.length) return null;
    var top = keys.reduce(function (a, b) { return groups[b].n > groups[a].n ? b : a; });
    var g = groups[top];
    // If one exact length pattern covers 90 %, use exact lengths so "ACCT_1_2_3_4" stands out.
    var exact = {};
    texts.forEach(function (t) {
      var runs = shapeRuns(t);
      if (runs.map(function (r) { return r[0]; }).join('') !== top) return;
      var k = runs.map(function (r) { return r[1]; }).join(',');
      exact[k] = (exact[k] || 0) + 1;
    });
    var topExact = Object.keys(exact).reduce(function (a, b) { return exact[b] > exact[a] ? b : a; });
    if (exact[topExact] >= texts.length * 0.9) {
      var lens = topExact.split(',').map(Number);
      g = { n: exact[topExact], runs: g.runs.map(function (r, i) { return [r[0], lens[i], lens[i]]; }) };
    }
    if (g.n < texts.length * 0.9) return null;
    if (!/9/.test(top) || !/[^Aa9]|9/.test(top)) return null;              // codes with digits, not free text
    if (/ /.test(top) && top.replace(/[^ ]/g, '').length > 2) return null; // looks like a sentence
    return '^' + g.runs.map(function (r) {
      var cls = r[0] === 'A' ? '[A-Z]' : r[0] === 'a' ? '[a-z]' : r[0] === '9' ? '\\d' : escRe(r[0]);
      var cnt = r[1] === r[2] ? (r[1] === 1 ? '' : '{' + r[1] + '}') : '{' + r[1] + ',' + r[2] + '}';
      return cls + cnt;
    }).join('') + '$';
  }

  /**
   * Report-level suggestions that need the full sheet (header + total rows).
   * rows: the whole sheet; headerRow: index of the header row.
   */
  function suggestForSheet(sheetName, rows, headerRow) {
    var width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    var hs = namedHeaders(rows[headerRow], width);
    var out = suggestRules(sheetName, hs, dataOnly(rows, headerRow));
    var kinds = classifyRows(rows, headerRow);
    var totals = [];
    kinds.forEach(function (k, j) { if (k === 'total') totals.push(headerRow + 1 + j); });
    if (totals.length === 1) {
      var tr = rows[totals[0]];
      hs.forEach(function (h, c) {
        var v = tr[c];
        if (typeof v !== 'number' && isNaN(SheetEngine.toNumber(v))) return;
        if (blank(v)) return;
        var isNumCol = out.some(function (r) { return r.column === h && (r.check === 'range' || r.check === 'number' || r.check === 'outlier' || r.check === 'decimals'); });
        if (!isNumCol) return;
        var ref = columnLetter(c + 1) + (totals[0] + 1);
        out.push({ sheet: sheetName, column: h, check: 'sum_equals', param: ref, severity: 'error',
          description: h + ' must add up to the total in ' + ref, enabled: true, suggested: true });
      });
    }
    return out;
  }

  function suggestRules(sheetName, headers, rows) {
    var data = rows.filter(function (r) { return r.some(function (v) { return !blank(v); }); });
    var n = data.length;
    if (!n) return [];
    var out = [];
    function add(column, check, param, severity, description) {
      out.push({ sheet: sheetName, column: column, check: check, param: param || '', severity: severity || 'error', description: description, enabled: true, suggested: true });
    }
    headers.forEach(function (h, c) {
      h = String(h == null ? '' : h).trim();
      if (!h) return;
      var vals = data.map(function (r) { return r[c]; });
      var filled = vals.filter(function (v) { return !blank(v); });
      if (!filled.length) return;
      var share = function (pred) { return filled.filter(pred).length / filled.length; };
      var name = h.toLowerCase();

      var nums = filled.map(function (v) { return SheetEngine.toNumber(v); });
      var numShare = nums.filter(function (x) { return !isNaN(x); }).length / filled.length;
      var dateShare = share(function (v) { return !!SheetEngine.toDate(v); });
      var boolShare = share(function (v) { return typeof v === 'boolean' || /^(true|false|wahr|falsch|yes|no|ja|nein)$/i.test(String(v).trim()); });
      var texts = filled.map(function (v) { return SheetEngine.display(v).trim(); });
      var distinct = {};
      texts.forEach(function (t) { var k = t.toLowerCase(); distinct[k] = (distinct[k] || 0) + 1; });
      var distinctCount = Object.keys(distinct).length;

      // completeness: a column that is almost always filled should always be filled
      if (filled.length / n >= 0.8 && filled.length < n) add(h, 'required', '', 'error', h + ' is usually filled; flag empty cells');
      else if (filled.length === n && n >= 3) add(h, 'required', '', 'error', h + ' must be filled');

      // identifiers: ID-like name and (almost) all values distinct
      var idName = /(^|[\s_\-.#])(id|nr|no|number|nummer|ref|reference|referenz|key|code)([\s_\-.#]|$)/i.test(name) || /(id|nr|nummer|number|no\.?)$/i.test(name.replace(/[\s_\-.#]+$/, ''));
      var repeats = filled.length - distinctCount;
      if (idName && filled.length >= 3 && (distinctCount / filled.length >= 0.8 || repeats <= Math.max(2, filled.length * 0.05))) {
        add(h, 'unique', '', 'error', h + ' must not repeat');
      }

      if (share(function (v) { return /^[A-Z]{2}\d{2}\s?[A-Z0-9]/i.test(String(v)) && String(v).replace(/\s/g, '').length >= 15; }) >= 0.8 || /\biban\b/.test(name)) {
        add(h, 'iban', '', 'error', h + ' must be a valid IBAN');
        return;
      }
      var textVals = filled.filter(function (v) { return typeof v === 'string'; });
      if (textVals.length >= filled.length * 0.5 && textVals.some(function (v) { return /^\s|\s$|\S\s{2,}\S/.test(v); })) {
        add(h, 'no_extra_spaces', '', 'warning', h + ' should have no stray spaces');
      }

      if (/\b(bic|swift)\b/.test(name) || share(function (v) { return /^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$/.test(String(v).trim()); }) >= 0.9) {
        add(h, 'bic', '', 'error', h + ' must be a valid BIC');
        return;
      }
      if (share(function (v) { return /@/.test(String(v)); }) >= 0.8 || /e-?mail/.test(name)) {
        add(h, 'email', '', 'warning', h + ' should be a valid e-mail address');
        return;
      }
      if (boolShare >= 0.9 && distinctCount <= 2) {
        add(h, 'allowed', 'TRUE, FALSE', 'error', h + ' must be TRUE or FALSE');
        return;
      }
      if (dateShare >= 0.9 && numShare < 0.5) {
        add(h, 'date', '', 'error', h + ' must be a valid date');
        if (/(created|booked|booking|paid|posted|erstellt|gebucht|bezahlt|timestamp|date of)/.test(name)) add(h, 'not_future', '', 'error', h + ' cannot be in the future');
        return;
      }
      if (numShare >= 0.9) {
        if (numShare < 1 || filled.some(function (v) { return typeof v !== 'number'; })) add(h, 'number', '', 'error', h + ' must be a number');
        var nonNeg = nums.filter(function (x) { return !isNaN(x) && x >= 0; }).length / filled.length;
        var pctName = /(%|percent|prozent|rate|quote|ratio|anteil|share|margin)/.test(name);
        var okNums = nums.filter(function (x) { return !isNaN(x); });
        var inUnit = okNums.filter(function (x) { return x >= 0 && x <= 1; }).length / okNums.length;
        var inPct = okNums.filter(function (x) { return x >= 0 && x <= 100; }).length / okNums.length;
        if (pctName && inUnit >= 0.95) add(h, 'range', '0..1', 'error', h + ' must be between 0 and 1 (0–100 %)');
        else if (pctName && inPct >= 0.95 && !/rate/.test(name)) add(h, 'range', '0..100', 'error', h + ' must be between 0 and 100');
        else if (nonNeg >= 0.95) add(h, 'range', '0..', 'error', h + ' must not be negative');
        var money = /(amount|balance|interest|tax|fee|price|cost|total|eur|usd|betrag|saldo|zins|steuer|gebühr|preis|summe)/.test(name);
        if (money) add(h, 'decimals', '2', 'warning', h + ' should have at most 2 decimals (cents)');
        var distinctNums = {};
        okNums.forEach(function (x) { distinctNums[x] = true; });
        if (!idName && okNums.length >= 20 && Object.keys(distinctNums).length >= 10) {
          add(h, 'outlier', '3', 'warning', 'Unusually large or small ' + h);
        }
        var numCounts = {};
        nums.forEach(function (x) { if (!isNaN(x)) numCounts[x] = (numCounts[x] || 0) + 1; });
        var keys = Object.keys(numCounts);
        var common = keys.filter(function (k) { return numCounts[k] >= 2; });
        var rareN = keys.length - common.length;
        var wholeNums = okNums.every(function (x) { return Math.round(x) === x; });
        if (wholeNums && !idName && !money && filled.length >= 10 && keys.length >= 2 && keys.length <= 8 && rareN <= 2 && common.length >= 2) {
          add(h, 'allowed', common.map(Number).sort(function (a, b) { return a - b; }).join(', '), 'warning', h + ' should be one of the usual values');
        }
        return;
      }
      var pattern = formatPattern(texts);
      if (pattern) {
        add(h, 'regex', pattern, 'error', h + ' should follow the same format as the other values');
      }

      // short codes (currency, country, status) used throughout the file
      var shortCodes = texts.every(function (t) { return t.length <= 12; });
      var commonCodes = Object.keys(distinct).filter(function (k) { return distinct[k] >= 2; });
      var rareCodes = Object.keys(distinct).length - commonCodes.length;
      if (shortCodes && filled.length >= 10 && commonCodes.length >= 1 && commonCodes.length <= 12 && rareCodes <= 3 &&
          (rareCodes > 0 || commonCodes.length <= 6) && /(currency|währung|country|land|status|type|typ|segment|category|kategorie|code|unit|einheit)/.test(name) ||
          shortCodes && filled.length >= 10 && commonCodes.length === 1 && rareCodes <= 2) {
        var keep = {};
        texts.forEach(function (t) { var k = t.toLowerCase(); if (distinct[k] >= 2 && !keep[k]) keep[k] = t; });
        add(h, 'allowed', Object.keys(keep).map(function (k) { return keep[k]; }).join(', '), rareCodes ? 'warning' : 'error', h + ' must be ' + (commonCodes.length === 1 ? keep[commonCodes[0]] : 'one of the usual values'));
        return;
      }
      // short code lists: a handful of values, each used several times → typos stand out
      if (filled.length >= 10 && distinctCount <= 12 && distinctCount < filled.length / 3) {
        var common = Object.keys(distinct).filter(function (k) { return distinct[k] >= 2; });
        var rare = Object.keys(distinct).filter(function (k) { return distinct[k] === 1; });
        if (rare.length && rare.length <= 3 && common.length >= 2) {
          var shown = {};
          texts.forEach(function (t) { if (distinct[t.toLowerCase()] >= 2 && !shown[t.toLowerCase()]) shown[t.toLowerCase()] = t; });
          add(h, 'allowed', Object.keys(shown).map(function (k) { return shown[k]; }).join(', '), 'warning', h + ' should be one of the values used elsewhere in the column');
        }
      }
      if (/(name|customer|client|kunde|bank|company|firma|supplier|lieferant|vendor)/.test(name) && distinctCount > 1) {
        add(h, 'fuzzy_duplicate', '', 'warning', 'Same ' + h + ' written differently');
      }
    });

    // start/end date pairs, e.g. "Placement Start Date" and "Deposit Closing Date"
    var dateCols = [];
    headers.forEach(function (h, c) {
      h = String(h == null ? '' : h).trim();
      var vals = data.map(function (r) { return r[c]; }).filter(function (v) { return !blank(v); });
      if (h && vals.length && vals.filter(function (v) { return !!SheetEngine.toDate(v); }).length / vals.length >= 0.9) dateCols.push({ h: h, c: c });
    });
    var startRe = /(start|begin|beginn|placement|open|opening|from|von|valid from|issue|ausgabe|eröffnung)/i;
    var endRe = /(end|ende|closing|close|maturity|fällig|until|bis|to\b|expiry|ablauf|valid to)/i;
    dateCols.filter(function (d) { return startRe.test(d.h); }).forEach(function (st) {
      dateCols.filter(function (d) { return d !== st && endRe.test(d.h) && !startRe.test(d.h); }).forEach(function (en) {
        var pairs = data.filter(function (r) { return SheetEngine.toDate(r[st.c]) && SheetEngine.toDate(r[en.c]); });
        if (!pairs.length) return;
        var ok = pairs.filter(function (r) { return SheetEngine.toDate(r[en.c]) >= SheetEngine.toDate(r[st.c]); }).length / pairs.length;
        if (ok >= 0.8) add(en.h, 'compare', '>= ' + st.h, 'error', en.h + ' cannot be before ' + st.h);
      });
    });
    return out;
  }

  function reportCsv(issues) {
    var q = function (v) {
      var s = isDate(v) ? SheetEngine.display(v) : String(v == null ? '' : v);
      if (/^[=+\-@]/.test(s)) s = "'" + s;          // keep spreadsheet apps from evaluating it
      return /[",;\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    var lines = [['Severity', 'Sheet', 'Cell', 'Column', 'Value', 'Problem', 'Rule'].join(',')];
    issues.forEach(function (i) {
      var cell = i.row ? (i.colIndex >= 0 ? columnLetter(i.colIndex + 1) : 'A') + i.row : '';
      lines.push([i.severity, i.sheet, cell, i.column, i.value, i.message, i.rule].map(q).join(','));
    });
    return '\uFEFF' + lines.join('\r\n');
  }

  return {
    RULES_SHEET: RULES_SHEET, REPORT_SHEET: REPORT_SHEET,
    parseCsv: parseCsv, sheetToRows: sheetToRows, detectHeaderRow: detectHeaderRow,
    rulesFromRows: rulesFromRows, columnLetter: columnLetter, suggestRules: suggestRules, suggestForSheet: suggestForSheet,
    namedHeaders: namedHeaders, classifyRows: classifyRows, dataOnly: dataOnly, cellValue: cellValue, formatPattern: formatPattern,
    buildXlsx: buildXlsx, buildReport: buildReport, reportCsv: reportCsv, crc32: crc32
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = SheetIO;
