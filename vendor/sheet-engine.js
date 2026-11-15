/**
 * Sheet — rule engine.
 *
 * Pure JavaScript with no Apps Script services, so the same file runs inside
 * Google Sheets (Apps Script V8) and in Node for tests.
 *
 * A rule is a plain object:
 *   { sheet, column, check, param, severity: 'error' | 'warning', description }
 * A table is:
 *   { sheetName, headers: [...], rows: [[...], ...], firstDataRow: 2 }
 */
var SheetEngine = (function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // Check catalogue (also used for the sidebar help and the Claude prompt)
  // ---------------------------------------------------------------------------
  var CHECKS = {
    required:        { scope: 'cell',   param: '', help: 'Cell must not be empty.' },
    number:          { scope: 'cell',   param: '', help: 'Must be a real number. Flags numbers stored as text such as "1.250,00".' },
    range:           { scope: 'cell',   param: 'min..max, e.g. 0..5, 0.. or ..1000', help: 'Number must lie within the range (inclusive).' },
    decimals:        { scope: 'cell',   param: 'max decimal places, e.g. 2', help: 'Number may have at most N decimal places.' },
    date:            { scope: 'cell',   param: '', help: 'Must be a valid date (date cell, YYYY-MM-DD or DD.MM.YYYY).' },
    not_future:      { scope: 'cell',   param: '', help: 'Date must not be in the future.' },
    compare:         { scope: 'cell',   param: 'operator and column, e.g. >= Start date', help: 'Compare with another column in the same row (>, >=, <, <=, =, !=).' },
    allowed:         { scope: 'cell',   param: 'comma-separated list, e.g. EUR, USD', help: 'Value must be one of the listed values (case-insensitive).' },
    regex:           { scope: 'cell',   param: 'regular expression, e.g. ^P-\\d{4}$', help: 'Value must match the pattern.' },
    length:          { scope: 'cell',   param: 'exact or min..max, e.g. 11 or 8..11', help: 'Text length must match.' },
    email:           { scope: 'cell',   param: '', help: 'Must look like an e-mail address.' },
    iban:            { scope: 'cell',   param: '', help: 'Valid IBAN: country length and checksum (catches typos).' },
    bic:             { scope: 'cell',   param: '', help: 'Valid BIC/SWIFT format.' },
    unique:          { scope: 'column', param: 'optional extra columns, e.g. Date, Currency', help: 'No duplicates, optionally for a combination of columns.' },
    fuzzy_duplicate: { scope: 'column', param: '', help: 'Near-duplicates, e.g. "Müller & Söhne GmbH" vs "Mueller und Soehne".' },
    sum_equals:      { scope: 'column', param: 'number or cell, e.g. 12500 or Bank!B2', help: 'Column total must equal a target (reconciliation).' },
    freshness:       { scope: 'column', param: 'days, e.g. 7', help: 'Newest date in the column must be at most N days old.' },
    min_rows:        { scope: 'table',  param: 'number, e.g. 10', help: 'Sheet must have at least N data rows. No column needed.' },
    no_extra_spaces: { scope: 'cell',   param: '', help: 'No leading, trailing or double spaces (they break lookups and matching).' },
    outlier:         { scope: 'column', param: 'sensitivity, e.g. 3 (higher = fewer flags)', help: 'Numbers far outside the usual range of the column (interquartile rule).' }
  };

  function checkNames() { return Object.keys(CHECKS); }

  function checkList() {
    return Object.keys(CHECKS).map(function (k) {
      return { name: k, scope: CHECKS[k].scope, param: CHECKS[k].param, help: CHECKS[k].help };
    });
  }

  // ---------------------------------------------------------------------------
  // Value helpers
  // ---------------------------------------------------------------------------
  function isDate(v) { return Object.prototype.toString.call(v) === '[object Date]'; }

  function isBlank(v) {
    return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  }

  /** Parses numbers, including German-formatted text like "1.250,00". */
  function toNumber(v) {
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    if (v === null || v === undefined || typeof v === 'boolean' || isDate(v)) return NaN;
    var s = String(v).trim().replace(/[\s€$£%]/g, '');
    if (s === '') return NaN;
    var lastDot = s.lastIndexOf('.');
    var lastComma = s.lastIndexOf(',');
    if (lastDot > -1 && lastComma > -1) {
      s = lastComma > lastDot ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    } else if (lastComma > -1) {
      s = (s.match(/,/g) || []).length === 1 ? s.replace(',', '.') : s.replace(/,/g, '');
    }
    if (!/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s)) return NaN;
    return parseFloat(s);
  }

  /** Accepts Date objects, YYYY-MM-DD[ hh:mm[:ss]] and DD.MM.YYYY / DD/MM/YYYY. */
  function toDate(v) {
    if (isDate(v)) return isNaN(v.getTime()) ? null : v;
    if (typeof v !== 'string') return null;
    var s = v.trim();
    var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) return makeDate(+m[1], +m[2], +m[3], m[4], m[5], m[6]);
    m = s.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) return makeDate(+m[3], +m[2], +m[1], m[4], m[5], m[6]);
    return null;
  }

  function makeDate(y, mo, d, h, mi, se) {
    var dt = new Date(y, mo - 1, d, +(h || 0), +(mi || 0), +(se || 0));
    if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return null;
    return dt;
  }

  function pad(n) { return (n < 10 ? '0' : '') + n; }

  function display(v) {
    if (isDate(v)) return v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
    if (v === null || v === undefined) return '';
    return String(v);
  }

  function fmtNum(n) { return String(Math.round(n * 100) / 100); }

  // ---------------------------------------------------------------------------
  // Identifier validation
  // ---------------------------------------------------------------------------
  var IBAN_LENGTHS = {
    AD: 24, AT: 20, BE: 16, BG: 22, CH: 21, CY: 28, CZ: 24, DE: 22, DK: 18, EE: 20,
    ES: 24, FI: 18, FR: 27, GB: 22, GI: 23, GR: 27, HR: 21, HU: 28, IE: 22, IS: 26,
    IT: 27, LI: 21, LT: 20, LU: 20, LV: 21, MC: 27, MT: 31, NL: 18, NO: 15, PL: 28,
    PT: 25, RO: 24, SE: 24, SI: 19, SK: 24, SM: 27
  };

  function mod97(str) {
    var rem = 0;
    for (var i = 0; i < str.length; i++) {
      var ch = str.charAt(i);
      var code = /[0-9]/.test(ch) ? ch : String(ch.charCodeAt(0) - 55);
      for (var j = 0; j < code.length; j++) rem = (rem * 10 + (+code.charAt(j))) % 97;
    }
    return rem;
  }

  function ibanProblem(v) {
    var s = String(v).replace(/\s+/g, '').toUpperCase();
    if (!/^[A-Z]{2}\d{2}[A-Z0-9]{8,30}$/.test(s)) return 'Not in IBAN format';
    var expected = IBAN_LENGTHS[s.slice(0, 2)];
    if (expected && s.length !== expected) {
      return s.slice(0, 2) + ' IBANs have ' + expected + ' characters, this one has ' + s.length;
    }
    if (mod97(s.slice(4) + s.slice(0, 4)) !== 1) return 'IBAN checksum is wrong (probably a typo)';
    return null;
  }

  /** Builds a valid IBAN from a country code and BBAN (used for demo data). */
  function makeIban(country, bban) {
    var check = 98 - mod97(bban + country + '00');
    return country + pad(check) + bban;
  }

  // ---------------------------------------------------------------------------
  // Fuzzy normalisation for near-duplicate names
  // ---------------------------------------------------------------------------
  var STOP_WORDS = ['gmbh', 'ag', 'kg', 'ug', 'ohg', 'gbr', 'mbh', 'co', 'ltd', 'limited', 'inc',
    'llc', 'plc', 'se', 'sa', 'sarl', 'bv', 'haftungsbeschrankt', 'und', 'and', 'the'];

  function fuzzyKey(v) {
    var s = String(v).toLowerCase()
      .replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ß/g, 'ss');
    if (s.normalize) s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    var tokens = s.replace(/[^a-z0-9]+/g, ' ').trim().split(/\s+/).filter(function (t) {
      return t && STOP_WORDS.indexOf(t) === -1;
    });
    tokens.sort();
    return tokens.join(' ');
  }

  // ---------------------------------------------------------------------------
  // Parameter parsing
  // ---------------------------------------------------------------------------
  function ConfigError(msg) { this.message = msg; this.isConfig = true; }

  function parseRange(p, what) {
    var m = String(p).trim().match(/^(-?[\d.,]*)\s*\.\.\s*(-?[\d.,]*)$/);
    if (!m || (m[1] === '' && m[2] === '')) throw new ConfigError(what + ' must look like 0..5, 0.. or ..5');
    var min = m[1] === '' ? null : toNumber(m[1]);
    var max = m[2] === '' ? null : toNumber(m[2]);
    if ((min !== null && isNaN(min)) || (max !== null && isNaN(max))) throw new ConfigError(what + ' contains something that is not a number');
    return { min: min, max: max };
  }

  function parsePositiveInt(p, what) {
    var n = parseInt(String(p).trim(), 10);
    if (isNaN(n) || n < 0) throw new ConfigError(what + ' must be a whole number');
    return n;
  }

  function lookupColumn(headerIdx, name) {
    var key = String(name).trim().toLowerCase();
    return Object.prototype.hasOwnProperty.call(headerIdx, key) ? headerIdx[key] : -1;
  }

  function comparable(v) {
    var d = toDate(v);
    if (d) return { type: 'date', v: d.getTime() };
    var n = toNumber(v);
    if (!isNaN(n)) return { type: 'number', v: n };
    return { type: 'text', v: String(v).trim().toLowerCase() };
  }

  var OPS = {
    '>':  function (a, b) { return a > b; },
    '>=': function (a, b) { return a >= b; },
    '<':  function (a, b) { return a < b; },
    '<=': function (a, b) { return a <= b; },
    '=':  function (a, b) { return a === b; },
    '!=': function (a, b) { return a !== b; }
  };

  // ---------------------------------------------------------------------------
  // Cell checks: return a function(value, rowArray) -> message | null
  // ---------------------------------------------------------------------------
  function buildCellTest(rule, headerIdx, now, opts) {
    opts = opts || {};
    var p = rule.param === undefined || rule.param === null ? '' : String(rule.param).trim();
    switch (rule.check) {
      case 'required':
        return function (v) { return isBlank(v) ? 'Missing value' : null; };

      case 'number':
        return function (v) {
          if (typeof v === 'number') return null;
          if (isNaN(toNumber(v))) return 'Not a number';
          return opts.textInput ? null : 'Number stored as text';   // CSV cells are always text
        };

      case 'range': {
        var r = parseRange(p, 'Range');
        return function (v) {
          var n = toNumber(v);
          if (isNaN(n)) return 'Not a number, so the range cannot be checked';
          if (r.min !== null && n < r.min) return 'Below the minimum of ' + r.min;
          if (r.max !== null && n > r.max) return 'Above the maximum of ' + r.max;
          return null;
        };
      }

      case 'decimals': {
        var places = parsePositiveInt(p, 'Decimals');
        var f = Math.pow(10, places);
        return function (v) {
          var n = toNumber(v);
          if (isNaN(n)) return 'Not a number';
          return Math.abs(n * f - Math.round(n * f)) > 1e-6 ? 'More than ' + places + ' decimal places' : null;
        };
      }

      case 'date':
        return function (v) { return toDate(v) ? null : 'Not a valid date'; };

      case 'not_future': {
        var endOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
        return function (v) {
          var d = toDate(v);
          if (!d) return 'Not a valid date';
          return d > endOfToday ? 'Date is in the future' : null;
        };
      }

      case 'compare': {
        var m = p.match(/^(>=|<=|!=|==|=|>|<)\s*(.+)$/);
        if (!m) throw new ConfigError('Compare needs an operator and a column, e.g. ">= Start date"');
        var op = m[1] === '==' ? '=' : m[1];
        var otherName = m[2].trim();
        var otherIdx = lookupColumn(headerIdx, otherName);
        if (otherIdx < 0) throw new ConfigError('Column "' + otherName + '" not found');
        return function (v, row) {
          var o = row[otherIdx];
          if (isBlank(o)) return null;
          var a = comparable(v), b = comparable(o);
          if (a.type !== b.type) return 'Cannot compare with ' + otherName + ' ("' + display(o) + '")';
          return OPS[op](a.v, b.v) ? null : 'Should be ' + op + ' ' + otherName + ' (' + display(o) + ')';
        };
      }

      case 'allowed': {
        var list = p.split(/[,;|]/).map(function (s) { return s.trim(); }).filter(Boolean);
        if (!list.length) throw new ConfigError('List the allowed values, e.g. "EUR, USD"');
        var lower = list.map(function (s) { return s.toLowerCase(); });
        var listNums = list.map(function (s) { return /^-?[\d.,]+$/.test(s) ? toNumber(s) : NaN; });
        return function (v) {
          if (lower.indexOf(display(v).trim().toLowerCase()) > -1) return null;
          var x = typeof v === 'number' ? v : (/^\s*-?[\d.,]+\s*$/.test(String(v)) ? toNumber(v) : NaN);
          if (!isNaN(x) && listNums.some(function (y) { return !isNaN(y) && Math.abs(x - y) < 1e-9; })) return null;
          return 'Not an allowed value (allowed: ' + list.join(', ') + ')';
        };
      }

      case 'regex': {
        var re;
        try { re = new RegExp(p); } catch (e) { throw new ConfigError('Invalid pattern: ' + e.message); }
        return function (v) { return re.test(display(v)) ? null : 'Does not match the expected pattern'; };
      }

      case 'length': {
        var lr = /\.\./.test(p) ? parseRange(p, 'Length') : (function () {
          var n = parsePositiveInt(p, 'Length');
          return { min: n, max: n };
        })();
        return function (v) {
          var len = display(v).trim().length;
          if (lr.min !== null && len < lr.min) return 'Too short (' + len + ' characters, expected ' + (lr.min === lr.max ? lr.min : 'at least ' + lr.min) + ')';
          if (lr.max !== null && len > lr.max) return 'Too long (' + len + ' characters, expected ' + (lr.min === lr.max ? lr.max : 'at most ' + lr.max) + ')';
          return null;
        };
      }

      case 'email':
        return function (v) {
          return /^[^\s@]+@[^\s@]+\.[^\s@.]{2,}$/.test(display(v).trim()) ? null : 'Not a valid e-mail address';
        };

      case 'iban':
        return function (v) { return ibanProblem(v); };

      case 'no_extra_spaces':
        return function (v) {
          if (typeof v !== 'string') return null;
          if (/^\s|\s$/.test(v)) return 'Leading or trailing space';
          if (/\S\s{2,}\S/.test(v)) return 'Double space inside the value';
          return null;
        };

      case 'bic':
        return function (v) {
          return /^[A-Z]{4}[A-Z]{2}[A-Z0-9]{2}([A-Z0-9]{3})?$/.test(display(v).replace(/\s+/g, '').toUpperCase()) ? null : 'Not a valid BIC';
        };
    }
    return null;
  }

  // ---------------------------------------------------------------------------
  // Running rules
  // ---------------------------------------------------------------------------
  function describe(rule) {
    if (rule.description) return rule.description;
    return [rule.column, rule.check, rule.param].filter(function (x) { return x !== undefined && x !== null && String(x) !== ''; }).join(' ');
  }

  function configIssue(rule, message) {
    return {
      sheet: rule.sheet || '', row: null, column: rule.column || '', colIndex: -1, value: '',
      check: rule.check, severity: 'config', message: 'Rule problem: ' + message, rule: describe(rule)
    };
  }

  function issue(rule, row, colIdx, value, message) {
    return {
      sheet: rule.sheet || '', row: row, column: rule.column || '', colIndex: colIdx, value: value,
      check: rule.check, severity: rule.severity === 'warning' ? 'warning' : 'error',
      message: message, rule: describe(rule)
    };
  }

  /** opts: { now: Date, textInput: true when every cell is text (CSV) } */
  function runRules(table, rules, opts) {
    opts = opts || {};
    var now = opts.now || new Date();
    var first = table.firstDataRow || 2;
    var headers = table.headers || [];
    var rows = table.rows || [];

    var headerIdx = {};
    headers.forEach(function (h, i) {
      var k = String(h).trim().toLowerCase();
      if (k && !Object.prototype.hasOwnProperty.call(headerIdx, k)) headerIdx[k] = i;
    });

    var dataRows = [];
    rows.forEach(function (r, i) {
      if (r.some(function (v) { return !isBlank(v); })) dataRows.push(i);
    });

    var issues = [];

    rules.forEach(function (rawRule) {
      var rule = {};
      for (var k in rawRule) rule[k] = rawRule[k];
      rule.sheet = rule.sheet || table.sheetName || '';
      rule.check = String(rule.check || '').trim().toLowerCase();
      rule.severity = String(rule.severity || 'error').trim().toLowerCase();
      try {
        var def = CHECKS[rule.check];
        if (!def) throw new ConfigError('Unknown check "' + rule.check + '"');
        var colIdx = -1;
        if (def.scope !== 'table') {
          if (isBlank(rule.column)) throw new ConfigError('No column given');
          colIdx = lookupColumn(headerIdx, rule.column);
          if (colIdx < 0) throw new ConfigError('Column "' + rule.column + '" not found in sheet "' + rule.sheet + '"');
          rule.column = String(headers[colIdx]).trim();
        }

        if (def.scope === 'cell') {
          var test = buildCellTest(rule, headerIdx, now, opts);
          dataRows.forEach(function (i) {
            var v = rows[i][colIdx];
            if (rule.check !== 'required' && isBlank(v)) return;
            var msg = test(v, rows[i]);
            if (msg) issues.push(issue(rule, first + i, colIdx, v, msg));
          });
        } else {
          runColumnRule(rule, colIdx, rows, dataRows, headerIdx, first, now, issues);
        }
      } catch (e) {
        if (e && e.isConfig) issues.push(configIssue(rule, e.message));
        else throw e;
      }
    });

    issues.sort(sortIssues);
    return { issues: issues, stats: statsFor(issues, rules.length, dataRows.length) };
  }

  function runColumnRule(rule, colIdx, rows, dataRows, headerIdx, first, now, issues) {
    var p = rule.param === undefined || rule.param === null ? '' : String(rule.param).trim();
    switch (rule.check) {
      case 'unique': {
        var cols = [colIdx];
        var extra = p ? p.split(/[,;]/).map(function (s) { return s.trim(); }).filter(Boolean) : [];
        extra.forEach(function (name) {
          var ix = lookupColumn(headerIdx, name);
          if (ix < 0) throw new ConfigError('Column "' + name + '" not found');
          cols.push(ix);
        });
        var seen = {};
        dataRows.forEach(function (i) {
          var parts = cols.map(function (c) {
            var v = rows[i][c];
            if (isDate(v)) return 'd' + v.getTime();
            return isBlank(v) ? '' : String(v).trim().toLowerCase();
          });
          if (parts.every(function (x) { return x === ''; })) return;
          var key = parts.join('\u241f');
          if (Object.prototype.hasOwnProperty.call(seen, key)) {
            issues.push(issue(rule, first + i, colIdx, rows[i][colIdx], 'Duplicate of row ' + seen[key]));
          } else {
            seen[key] = first + i;
          }
        });
        return;
      }

      case 'fuzzy_duplicate': {
        // Group values that normalise to the same key; within a group the most
        // common exact spelling is "right" and every other spelling is flagged.
        // Exact repeats (the same customer on many rows) are not flagged.
        var groups = {};
        dataRows.forEach(function (i) {
          var v = rows[i][colIdx];
          if (isBlank(v)) return;
          var key = fuzzyKey(v);
          if (!key) return;
          var exact = display(v).trim().replace(/\s+/g, ' ');
          var g = groups[key] || (groups[key] = { order: [], variants: {} });
          if (!g.variants[exact]) { g.variants[exact] = { count: 0, firstRow: first + i, rows: [] }; g.order.push(exact); }
          g.variants[exact].count++;
          g.variants[exact].rows.push({ row: first + i, value: v });
        });
        Object.keys(groups).forEach(function (key) {
          var g = groups[key];
          if (g.order.length < 2) return;
          var main = g.order.reduce(function (best, x) { return g.variants[x].count > g.variants[best].count ? x : best; }, g.order[0]);
          g.order.forEach(function (x) {
            if (x === main) return;
            g.variants[x].rows.forEach(function (r) {
              issues.push(issue(rule, r.row, colIdx, r.value, 'Looks like the same as row ' + g.variants[main].firstRow + ' ("' + main + '")'));
            });
          });
        });
        return;
      }

      case 'sum_equals': {
        var targetRaw = rule.resolvedParam !== undefined && rule.resolvedParam !== null ? rule.resolvedParam : p;
        var target = toNumber(targetRaw);
        if (isNaN(target)) throw new ConfigError('Target "' + p + '" is not a number or a readable cell');
        var sum = 0, skipped = 0;
        dataRows.forEach(function (i) {
          var v = rows[i][colIdx];
          if (isBlank(v)) return;
          var n = toNumber(v);
          if (isNaN(n)) skipped++; else sum += n;
        });
        var diff = sum - target;
        if (Math.abs(diff) > 0.005) {
          issues.push(issue(rule, null, colIdx, fmtNum(sum),
            'Column total ' + fmtNum(sum) + ' does not match target ' + fmtNum(target) +
            ' (difference ' + (diff > 0 ? '+' : '') + fmtNum(diff) + ')' +
            (skipped ? '; ' + skipped + ' non-numeric cell(s) ignored' : '')));
        }
        return;
      }

      case 'freshness': {
        var days = parsePositiveInt(p, 'Freshness');
        var newest = null;
        dataRows.forEach(function (i) {
          var d = toDate(rows[i][colIdx]);
          if (d && (!newest || d > newest)) newest = d;
        });
        if (!newest) {
          issues.push(issue(rule, null, colIdx, '', 'No valid dates found'));
          return;
        }
        var age = Math.floor((now.getTime() - newest.getTime()) / 86400000);
        if (age > days) {
          issues.push(issue(rule, null, colIdx, display(newest),
            'Newest entry is from ' + display(newest) + ' (' + age + ' days old, limit ' + days + ')'));
        }
        return;
      }

      case 'outlier': {
        var k = p === '' ? 3 : toNumber(p);
        if (isNaN(k) || k <= 0) throw new ConfigError('Sensitivity must be a positive number, e.g. 3');
        var vals = [];
        dataRows.forEach(function (i) { var x = toNumber(rows[i][colIdx]); if (!isNaN(x)) vals.push(x); });
        if (vals.length < 8) return;
        var sorted = vals.slice().sort(function (a, b) { return a - b; });
        var q = function (f) { var pos = (sorted.length - 1) * f, lo = Math.floor(pos), hi = Math.ceil(pos); return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo); };
        var q1 = q(0.25), q3 = q(0.75), iqr = q3 - q1;
        if (iqr <= 0) return;
        var lowB = q1 - k * iqr, highB = q3 + k * iqr;
        dataRows.forEach(function (i) {
          var x = toNumber(rows[i][colIdx]);
          if (isNaN(x)) return;
          if (x < lowB || x > highB) issues.push(issue(rule, first + i, colIdx, rows[i][colIdx],
            'Unusual value: most values are between ' + fmtNum(q1) + ' and ' + fmtNum(q3)));
        });
        return;
      }

      case 'min_rows': {
        var minRows = parsePositiveInt(p, 'Minimum rows');
        if (dataRows.length < minRows) {
          issues.push(issue(rule, null, -1, String(dataRows.length),
            'Only ' + dataRows.length + ' data rows, expected at least ' + minRows));
        }
        return;
      }
    }
  }

  var SEV_ORDER = { config: 0, error: 1, warning: 2 };

  function sortIssues(a, b) {
    if (SEV_ORDER[a.severity] !== SEV_ORDER[b.severity]) return SEV_ORDER[a.severity] - SEV_ORDER[b.severity];
    if (a.sheet !== b.sheet) return a.sheet < b.sheet ? -1 : 1;
    return (a.row || 0) - (b.row || 0);
  }

  function statsFor(issues, ruleCount, rowCount) {
    var s = { rules: ruleCount, rows: rowCount, errors: 0, warnings: 0, config: 0 };
    issues.forEach(function (i) {
      if (i.severity === 'error') s.errors++;
      else if (i.severity === 'warning') s.warnings++;
      else s.config++;
    });
    return s;
  }

  function issueKey(i) {
    return [i.sheet, i.row, i.column, i.check, display(i.value), i.message].join('|');
  }

  // ---------------------------------------------------------------------------
  // Validating rules that come from outside (e.g. Claude)
  // ---------------------------------------------------------------------------
  function validateRules(candidates, headersBySheet) {
    var sheetNames = Object.keys(headersBySheet);
    var out = [];
    (candidates || []).forEach(function (c) {
      if (!c || !CHECKS[String(c.check || '').toLowerCase()]) return;
      var check = String(c.check).toLowerCase();
      var sheet = sheetNames.filter(function (n) { return n.toLowerCase() === String(c.sheet || '').trim().toLowerCase(); })[0];
      if (!sheet && sheetNames.length === 1) sheet = sheetNames[0];
      if (!sheet) return;
      var column = '';
      if (CHECKS[check].scope !== 'table') {
        column = (headersBySheet[sheet] || []).filter(function (h) {
          return String(h).trim().toLowerCase() === String(c.column || '').trim().toLowerCase();
        })[0];
        if (!column) return;
        column = String(column).trim();
      }
      out.push({
        sheet: sheet, column: column, check: check,
        param: c.param === undefined || c.param === null ? '' : String(c.param),
        severity: String(c.severity || '').toLowerCase() === 'warning' ? 'warning' : 'error',
        description: c.description ? String(c.description) : ''
      });
    });
    return out;
  }

  // ---------------------------------------------------------------------------
  // Offline plain-language parser (English + German). Claude does this better;
  // this keeps the add-on useful without an API key.
  // ---------------------------------------------------------------------------
  function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  function findColumns(text, headers) {
    var t = text.toLowerCase();
    var found = [];
    headers.map(function (h) { return String(h).trim(); }).filter(Boolean)
      .sort(function (a, b) { return b.length - a.length; })
      .forEach(function (h) {
        var re = new RegExp('(^|[^a-z0-9äöüß])' + escapeRe(h.toLowerCase()) + '(?=$|[^a-z0-9äöüß])', 'g');
        var m;
        while ((m = re.exec(t)) !== null) {
          var start = m.index + m[1].length, end = start + h.length;
          var overlaps = found.some(function (f) { return start < f.end && end > f.start; });
          if (!overlaps) found.push({ header: h, start: start, end: end });
        }
      });
    found.sort(function (a, b) { return a.start - b.start; });
    var seen = {}, out = [];
    found.forEach(function (f) { if (!seen[f.header]) { seen[f.header] = true; out.push(f.header); } });
    return out;
  }

  function textNum(s) {
    if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) return parseFloat(s.replace(/,/g, ''));
    return toNumber(s);
  }

  var NUM = '(-?\\d+(?:[.,]\\d+)*)';

  function parseRuleText(text, headers) {
    if (isBlank(text)) return null;
    var original = String(text).trim();
    var t = original.toLowerCase();
    var cols = findColumns(original, headers || []);
    var column = cols[0] || '';
    var severity = /\b(warn|warning|warnung|hinweis)\b/.test(t) ? 'warning' : 'error';
    var m;

    function rule(check, param, needsColumn) {
      if (needsColumn !== false && !column) return null;
      return { column: needsColumn === false ? '' : column, check: check, param: param === undefined ? '' : String(param), severity: severity, description: original };
    }

    if ((m = t.match(new RegExp('(?:at least|mindestens|min(?:imum)?)\\s*' + NUM + '\\s*(?:data\\s*)?(?:rows|zeilen|entries|einträge|records|lines)'))))
      return rule('min_rows', textNum(m[1]), false);

    if (/(updat|refresh|aktualis|fresh|older|älter|within|innerhalb|newest|neueste|latest)/.test(t) &&
        (m = t.match(/(\d+)\s*(?:days?|tagen?)/)))
      return rule('freshness', m[1]);

    if (/(required|mandatory|must not be (?:empty|blank|missing)|can'?t be (?:empty|blank)|cannot be (?:empty|blank)|never (?:empty|blank)|always (?:be )?filled|must be filled|pflicht|darf nicht leer|muss ausgefüllt|nie leer)/.test(t))
      return rule('required');

    if (/(similar|look[- ]?alike|near[- ]?duplicate|spelling|spelled|ähnlich|tippfehler|schreibweise)/.test(t))
      return rule('fuzzy_duplicate');

    if (/(unique|no duplicates?|not (?:be )?duplicated|eindeutig|keine (?:duplikate|doppelten)|doppelt|only once|nur einmal)/.test(t)) {
      var extra = cols.length > 1 && /(combination|together|per|with|and|kombination|zusammen|pro|und)/.test(t) ? cols.slice(1).join(', ') : '';
      return rule('unique', extra);
    }

    if (/\b(sum|total|summe|gesamt)/.test(t) && /(equal|match|=|gleich|entspricht|entsprech|stimm|abgleich|reconcile)/.test(t)) {
      var ref = original.match(/((?:'[^']+'|[A-Za-z0-9_]+)![A-Z]{1,3}\d+)|\b([A-Z]{1,3}\d{1,6})\b/);
      if (ref) return rule('sum_equals', ref[0]);
      m = t.match(new RegExp('(?:equal|match|=|gleich|entspricht|entsprech\\w*|stimmt|reconcile)[^\\d-]*' + NUM));
      if (m) return rule('sum_equals', textNum(m[1]));
    }

    if (cols.length > 1) {
      var op = null;
      if (/(on or after|not before|nicht vor|am oder nach|>=|at least as (?:late|large|high))/.test(t)) op = '>=';
      else if (/(on or before|not after|nicht nach|am oder vor|<=|at most|no later than|not later than)/.test(t)) op = '<=';
      else if (/(!=|not equal|different from|ungleich|nicht gleich)/.test(t)) op = '!=';
      else if (/(\bafter\b|later than|greater than|more than|higher than|larger than|\bnach\b|größer|später|höher|>)/.test(t)) op = '>';
      else if (/(\bbefore\b|earlier than|less than|smaller than|lower than|\bvor\b|kleiner|früher|niedriger|<)/.test(t)) op = '<';
      else if (/(equal|same as|gleich|identisch|=)/.test(t)) op = '=';
      if (op) return rule('compare', op + ' ' + cols[1]);
    }

    if (/(extra spaces?|trailing spaces?|leading spaces?|leerzeichen|whitespace)/.test(t)) return rule('no_extra_spaces');
    if (/(outlier|unusual|abnormal|ausreißer|auffällig|ungewöhnlich)/.test(t)) return rule('outlier', '3');
    if (/\biban\b/.test(t)) return rule('iban');
    if (/\b(bic|swift)\b/.test(t)) return rule('bic');
    if (/e-?mail/.test(t)) return rule('email');
    if (/(future|zukunft)/.test(t)) return rule('not_future');

    if ((m = t.match(/(\d+)\s*(?:decimal|decimals|decimal places|nachkommastellen|nachkomma|dezimalstellen)/)))
      return rule('decimals', m[1]);

    if ((m = t.match(/(\d+)\s*(?:characters|chars|zeichen|digits|stellen|ziffern)/)))
      return rule('length', m[1]);

    if ((m = original.match(/(?:match(?:es)?|pattern|format|muster)\s*:?\s*(\^\S+)/i)))
      return rule('regex', m[1]);

    if ((m = original.match(/(?:one of|either|must be in|only|allowed(?: values)?(?: are)?|erlaubt(?: sind)?|nur|entweder)\s*:?\s*(.+)$/i))) {
      var list = m[1].replace(/[.!]+$/, '').replace(/\s+(?:or|oder)\s+/gi, ', ')
        .split(/[,;|/]/).map(function (s) { return s.trim(); }).filter(Boolean);
      if (list.length) return rule('allowed', list.join(', '));
    }

    if ((m = t.match(new RegExp('(?:between|zwischen)\\s+' + NUM + '\\s*(?:%\\s*)?(?:and|und|bis|-)\\s*' + NUM))))
      return rule('range', textNum(m[1]) + '..' + textNum(m[2]));

    var min = null, max = null;
    if ((m = t.match(new RegExp('(?:at least|minimum|min\\.?|mindestens|not less than|nicht weniger als|>=)\\s*' + NUM)))) min = textNum(m[1]);
    if ((m = t.match(new RegExp('(?:at most|maximum|max\\.?|höchstens|maximal|not more than|nicht mehr als|<=)\\s*' + NUM)))) max = textNum(m[1]);
    if (min === null && /(positive|not (?:be )?negative|non-negative|nicht negativ|no negative|keine negativen)/.test(t)) min = 0;
    if (min !== null || max !== null) return rule('range', (min === null ? '' : min) + '..' + (max === null ? '' : max));

    if (/\b(number|numeric|zahl|numerisch)\b/.test(t)) return rule('number');
    if (/\b(date|datum)\b/.test(t) && /(valid|gültig|real|proper|korrekt)/.test(t)) return rule('date');

    return null;
  }

  return {
    CHECKS: CHECKS,
    checkNames: checkNames,
    checkList: checkList,
    runRules: runRules,
    configIssue: configIssue,
    issueKey: issueKey,
    validateRules: validateRules,
    parseRuleText: parseRuleText,
    display: display,
    toNumber: toNumber,
    toDate: toDate,
    ibanProblem: ibanProblem,
    makeIban: makeIban,
    fuzzyKey: fuzzyKey
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = SheetEngine;
