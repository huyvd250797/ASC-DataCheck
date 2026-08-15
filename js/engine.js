/* ASC-DataCheck — Validation Engine.
   Không phụ thuộc DOM: có thể chạy trong trình duyệt hoặc Node (dùng cho test). */
(function (root) {
  'use strict';
  var U = root.ASC.util;

  var MAX_ISSUES_PER_RULE = 8000;   // trần an toàn cho tệp rất lớn
  var CHUNK = 2000;

  function nextTick() { return new Promise(function (r) { setTimeout(r, 0); }); }

  /* Vòng lặp chia lô, có tiến độ và có thể hủy */
  function chunkLoop(total, body, onProgress, token, phase) {
    var i = 0;
    function step() {
      if (token && token.cancelled) return Promise.reject(new Error('CANCELLED'));
      var end = Math.min(i + CHUNK, total);
      for (; i < end; i++) body(i);
      if (onProgress) onProgress({ phase: phase, done: i, total: total });
      if (i >= total) return Promise.resolve();
      return nextTick().then(step);
    }
    return total === 0 ? Promise.resolve() : step();
  }

  function colLetter(n) {
    var s = '';
    n = n + 1;
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
    return s;
  }

  /* ================= DATASET ================= */
  /**
   * @param aoa mảng 2 chiều lấy từ sheet (giữ nguyên kiểu: Date, number, string)
   * @param headerRowIdx chỉ số dòng tiêu đề trong aoa (0-based)
   */
  function buildDataset(aoa, headerRowIdx, opt) {
    opt = opt || {};
    var notes = [];
    var headerRaw = aoa[headerRowIdx] || [];
    var width = 0;
    for (var r = 0; r < aoa.length; r++) width = Math.max(width, (aoa[r] || []).length);

    var headers = [], seen = {};
    for (var c = 0; c < width; c++) {
      var raw = headerRaw[c];
      var name = U.normHeader(U.toStr(raw));
      var invalid = false;
      if (name === '' || /^unnamed[:_ ]/i.test(name) || /^column\d+$/i.test(name)) {
        invalid = true;
        name = 'Cột ' + colLetter(c);
      }
      var key = U.normKey(name);
      var dup = false;
      if (seen[key]) { dup = true; } else { seen[key] = 1; }
      headers.push({ index: c, raw: U.toStr(raw), name: name, key: key, letter: colLetter(c), invalidHeader: invalid, duplicateHeader: dup });
    }

    // Bỏ các cột đuôi hoàn toàn trống (không header, không dữ liệu)
    while (headers.length) {
      var last = headers[headers.length - 1];
      if (!last.invalidHeader) break;
      var any = false;
      for (var rr = headerRowIdx + 1; rr < aoa.length; rr++) {
        if (!U.isBlank((aoa[rr] || [])[last.index])) { any = true; break; }
      }
      if (any) break;
      headers.pop();
    }

    var rows = [], rowNo = [], skippedEmpty = 0, footerRows = [];
    for (var i = headerRowIdx + 1; i < aoa.length; i++) {
      var src = aoa[i] || [];
      var filled = 0, firstVal = null;
      for (var k = 0; k < headers.length; k++) {
        if (!U.isBlank(src[k])) { filled++; if (firstVal === null) firstVal = src[k]; }
      }
      if (filled === 0) { skippedEmpty++; if (!opt.keepEmptyRows) continue; }
      // Dòng tổng cộng ở cuối bảng
      if (filled <= 2 && i >= aoa.length - 3) {
        var txt = U.stripDiacritics(U.toStr(firstVal)).toLowerCase();
        if (/^(tong|total|cong|sum)\b/.test(txt) || txt.indexOf('tong cong') >= 0) {
          footerRows.push(i + 1); continue;
        }
      }
      var row = new Array(headers.length);
      for (var c2 = 0; c2 < headers.length; c2++) row[c2] = src[c2] === undefined ? null : src[c2];
      rows.push(row);
      rowNo.push(i + 1); // số dòng thật trong Excel (1-based)
    }

    if (skippedEmpty) notes.push({ kind: 'info', text: 'Đã bỏ qua ' + U.fmtInt(skippedEmpty) + ' dòng trống.' });
    if (footerRows.length) notes.push({ kind: 'info', text: 'Đã bỏ qua dòng tổng cộng ở cuối bảng (dòng ' + footerRows.join(', ') + ').' });

    return { headers: headers, rows: rows, rowNo: rowNo, notes: notes, headerRowIdx: headerRowIdx };
  }

  /* ================= ISSUE ================= */
  var _seq = 0;
  function issue(o) {
    o.id = ++_seq;
    if (!o.severity) o.severity = 'error';
    return o;
  }
  function resetSeq() { _seq = 0; }

  /* ================= PROFILE ================= */
  function profileColumn(ds, colIdx) {
    var total = ds.rows.length, blank = 0, minLen = Infinity, maxLen = 0, sumLen = 0, counted = 0;
    var types = { number: 0, date: 0, text: 0, boolean: 0 };
    var freq = new Map();
    var numeric = [];
    for (var i = 0; i < total; i++) {
      var v = ds.rows[i][colIdx];
      if (U.isBlank(v)) { blank++; continue; }
      var s = U.toStr(v);
      var t = U.cellType(v);
      if (types[t] !== undefined) types[t]++;
      var L = s.trim().length;
      minLen = Math.min(minLen, L); maxLen = Math.max(maxLen, L); sumLen += L; counted++;
      var kf = s.trim();
      freq.set(kf, (freq.get(kf) || 0) + 1);
      if (t === 'number') { var n = U.toNumber(v); if (!isNaN(n)) numeric.push(n); }
    }
    var dupValues = 0, dupCells = 0;
    freq.forEach(function (c) { if (c > 1) { dupValues++; dupCells += c; } });
    var top = [];
    freq.forEach(function (c, v) { top.push([v, c]); });
    top.sort(function (a, b) { return b[1] - a[1]; });

    var stat = {
      column: ds.headers[colIdx] ? ds.headers[colIdx].name : '',
      index: colIdx, total: total, blank: blank, filled: total - blank,
      unique: freq.size, duplicateValues: dupValues, duplicateCells: dupCells,
      minLength: counted ? minLen : 0, maxLength: maxLen,
      avgLength: counted ? Math.round((sumLen / counted) * 10) / 10 : 0,
      types: types, top: top.slice(0, 20), distinctCount: freq.size
    };
    if (numeric.length) {
      numeric.sort(function (a, b) { return a - b; });
      stat.numeric = {
        count: numeric.length, min: numeric[0], max: numeric[numeric.length - 1],
        q1: numeric[Math.floor(numeric.length * 0.25)], q3: numeric[Math.floor(numeric.length * 0.75)],
        avg: Math.round((numeric.reduce(function (a, b) { return a + b; }, 0) / numeric.length) * 100) / 100
      };
    }
    stat.dominantType = 'text';
    var best = -1;
    Object.keys(types).forEach(function (t) { if (types[t] > best) { best = types[t]; stat.dominantType = t; } });
    return stat;
  }

  /* ================= QUICK CHECK ================= */
  var CODE_HINT = /(^|[^a-z])(ma|mssv|msv|cccd|cmnd|code|id|so)([^a-z]|$)/;

  function looksLikeCode(header) {
    var k = U.normKey(header);
    return CODE_HINT.test(' ' + U.stripDiacritics(header).toLowerCase() + ' ') ||
      /^(ma|msv|mssv|cccd|cmnd|code|id)/.test(k);
  }

  function quickCheck(ds, opt, onProgress, token) {
    opt = opt || {};
    resetSeq();
    var issues = [], profiles = [];
    var sheet = opt.sheet || '';
    var nCols = ds.headers.length;

    return chunkLoop(nCols, function (c) {
      var h = ds.headers[c];
      var p = profileColumn(ds, c);
      profiles.push(p);

      if (h.duplicateHeader) {
        issues.push(issue({
          sheet: sheet, rowIndex: -1, rowNo: ds.headerRowIdx + 1, col: c, column: h.name, value: h.name,
          severity: 'error', ruleType: 'DUPLICATE_HEADER',
          message: 'Có nhiều hơn một cột cùng tên "' + h.name + '". Kết quả ghép cột có thể sai.',
          suggestion: 'Đổi tên một trong hai cột trong tệp Excel rồi tải lại.'
        }));
      }
      if (h.invalidHeader && p.filled > 0) {
        issues.push(issue({
          sheet: sheet, rowIndex: -1, rowNo: ds.headerRowIdx + 1, col: c, column: h.name, value: '',
          severity: 'warning', ruleType: 'INVALID_HEADER',
          message: 'Cột ' + h.letter + ' không có tiêu đề nhưng vẫn có dữ liệu.',
          suggestion: 'Bổ sung tiêu đề cho cột hoặc xóa cột nếu không dùng.'
        }));
      }
      if (p.filled === 0) {
        issues.push(issue({
          sheet: sheet, rowIndex: -1, rowNo: ds.headerRowIdx + 1, col: c, column: h.name, value: '',
          severity: 'warning', ruleType: 'EMPTY_COLUMN',
          message: 'Cột "' + h.name + '" hoàn toàn trống.',
          suggestion: 'Kiểm tra lại tệp nguồn hoặc bỏ cột này khi import.'
        }));
        return;
      }

      var fillRate = p.filled / Math.max(1, p.total);
      var isCode = looksLikeCode(h.name);

      // Mã đang được lưu dạng số -> nguy cơ mất số 0 đầu
      if (isCode && p.types.number > 0 && p.types.number / p.filled > 0.9) {
        issues.push(issue({
          sheet: sheet, rowIndex: -1, rowNo: ds.headerRowIdx + 1, col: c, column: h.name, value: '',
          severity: 'warning', ruleType: 'LEADING_ZERO',
          message: 'Cột "' + h.name + '" là cột mã nhưng đang được Excel lưu dạng số. Số 0 ở đầu có thể đã mất.',
          suggestion: 'Yêu cầu khách hàng định dạng cột này là Text rồi xuất lại tệp.'
        }));
      }

      // Kiểu dữ liệu lẫn lộn
      var kinds = ['number', 'date', 'text'].filter(function (t) { return p.types[t] > 0; });
      if (kinds.length > 1) {
        var majority = kinds.reduce(function (a, b) { return p.types[a] >= p.types[b] ? a : b; });
        var minorityCount = p.filled - p.types[majority];
        if (minorityCount / p.filled < 0.35) {
          issues.push(issue({
            sheet: sheet, rowIndex: -1, rowNo: ds.headerRowIdx + 1, col: c, column: h.name, value: '',
            severity: 'warning', ruleType: 'MIXED_TYPE',
            message: 'Cột "' + h.name + '" có kiểu dữ liệu lẫn lộn: ' + kinds.map(function (t) {
              return t + ' ' + Math.round(p.types[t] / p.filled * 100) + '%';
            }).join(', ') + '.',
            suggestion: 'Chuẩn hóa về một kiểu dữ liệu trước khi import.'
          }));
        }
      }

      // Duyệt từng ô của cột
      var seenVals = new Map();
      var flagDup = isCode || (p.unique / Math.max(1, p.filled)) > 0.92;
      for (var i = 0; i < ds.rows.length; i++) {
        var v = ds.rows[i][c];
        var blank = U.isBlank(v);
        if (blank) {
          if (fillRate >= 0.8 && issues.length < MAX_ISSUES_PER_RULE) {
            issues.push(issue({
              sheet: sheet, rowIndex: i, rowNo: ds.rowNo[i], col: c, column: h.name, value: '',
              severity: 'warning', ruleType: 'BLANK',
              message: 'Ô trống trong cột "' + h.name + '" (cột này có ' + Math.round(fillRate * 100) + '% dòng có dữ liệu).',
              suggestion: 'Bổ sung dữ liệu hoặc xác nhận cột này được phép để trống.'
            }));
          }
          continue;
        }
        var s = U.toStr(v);
        if (U.hasEdgeSpace(s)) {
          issues.push(issue({
            sheet: sheet, rowIndex: i, rowNo: ds.rowNo[i], col: c, column: h.name, value: s,
            severity: 'warning', ruleType: 'WHITESPACE',
            message: 'Giá trị có khoảng trắng ở đầu hoặc cuối.',
            suggestion: s.trim()
          }));
        } else if (U.hasDoubleSpace(s)) {
          issues.push(issue({
            sheet: sheet, rowIndex: i, rowNo: ds.rowNo[i], col: c, column: h.name, value: s,
            severity: 'warning', ruleType: 'DOUBLE_SPACE',
            message: 'Giá trị có khoảng trắng kép ở giữa.',
            suggestion: s.replace(/\s+/g, ' ').trim()
          }));
        }
        if (typeof v === 'number' && Math.abs(v) >= 1e11 && isCode) {
          issues.push(issue({
            sheet: sheet, rowIndex: i, rowNo: ds.rowNo[i], col: c, column: h.name, value: s,
            severity: 'warning', ruleType: 'SCIENTIFIC',
            message: 'Số quá dài, Excel có thể hiển thị dạng khoa học và làm sai giá trị.',
            suggestion: 'Định dạng cột là Text trong tệp nguồn.'
          }));
        }
        if (flagDup) {
          var key = s.trim().toLowerCase();
          if (seenVals.has(key)) seenVals.get(key).push(i);
          else seenVals.set(key, [i]);
        }
      }

      if (flagDup) {
        seenVals.forEach(function (rowsIdx, key) {
          if (rowsIdx.length < 2) return;
          var rowNos = rowsIdx.map(function (r) { return ds.rowNo[r]; });
          rowsIdx.forEach(function (r) {
            issues.push(issue({
              sheet: sheet, rowIndex: r, rowNo: ds.rowNo[r], col: c, column: h.name,
              value: U.toStr(ds.rows[r][c]),
              severity: 'warning', ruleType: 'DUPLICATE',
              message: 'Giá trị xuất hiện ' + rowsIdx.length + ' lần trong cột "' + h.name + '".',
              suggestion: 'Xác nhận đây là dữ liệu trùng hay là hai bản ghi khác nhau.',
              relatedRows: rowNos
            }));
          });
        });
      }

      // Giá trị bất thường trong cột số (IQR)
      if (p.numeric && p.numeric.count > 20) {
        var iqr = p.numeric.q3 - p.numeric.q1;
        if (iqr > 0) {
          var lo = p.numeric.q1 - 3 * iqr, hi = p.numeric.q3 + 3 * iqr, flagged = 0;
          for (var j = 0; j < ds.rows.length && flagged < 200; j++) {
            var nv = ds.rows[j][c];
            if (U.isBlank(nv)) continue;
            var num = U.toNumber(nv);
            if (isNaN(num)) continue;
            if (num < lo || num > hi) {
              flagged++;
              issues.push(issue({
                sheet: sheet, rowIndex: j, rowNo: ds.rowNo[j], col: c, column: h.name, value: U.toStr(nv),
                severity: 'info', ruleType: 'OUTLIER',
                message: 'Giá trị nằm xa dải phổ biến của cột (' + p.numeric.q1 + ' – ' + p.numeric.q3 + ').',
                suggestion: 'Kiểm tra lại xem có nhập nhầm không.'
              }));
            }
          }
        }
      }
    }, onProgress, token, 'Đang phân tích cột').then(function () {
      return { issues: issues, profiles: profiles };
    });
  }

  /* ================= RULE VALIDATION ================= */

  function sev(rules, key, dflt) {
    return rules[key + 'Severity'] || rules.severity || dflt || 'error';
  }

  /**
   * Suy luận thứ tự ngày/tháng của cả cột.
   * Chỉ cần một giá trị có phần đầu > 12 là biết chắc cột đang là dd/MM.
   */
  function inferDateOrder(ds, colIdx) {
    var dmy = 0, mdy = 0, seen = 0;
    for (var i = 0; i < ds.rows.length; i++) {
      var v = ds.rows[i][colIdx];
      if (v instanceof Date || typeof v === 'number') { seen++; continue; }
      var s = U.toStr(v).trim();
      var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
      if (!m) continue;
      seen++;
      var a = Number(m[1]), b = Number(m[2]);
      if (a > 12 && b <= 12) dmy++;
      else if (b > 12 && a <= 12) mdy++;
    }
    if (dmy && mdy) return 'MIXED';
    if (dmy) return 'DMY';
    if (mdy) return 'MDY';
    return seen ? 'UNKNOWN' : 'NONE';
  }

  function buildReferenceSet(ds, colIdx) {
    var set = new Set();
    for (var i = 0; i < ds.rows.length; i++) {
      var v = ds.rows[i][colIdx];
      if (U.isBlank(v)) continue;
      set.add(U.toStr(v).trim().toLowerCase());
    }
    return set;
  }

  /**
   * @param ds dataset của sheet đang kiểm tra
   * @param ruleSet bộ rule
   * @param mapping {fieldKey: colIdx}
   * @param refs {fieldKey: Set<string>} danh mục tham chiếu đã dựng sẵn
   */
  function validate(ds, ruleSet, mapping, refs, onProgress, token, opt) {
    opt = opt || {};
    resetSeq();
    var sheet = opt.sheet || '';
    var issues = [];
    var today = new Date(); today.setHours(23, 59, 59, 999);

    var fields = ruleSet.fields.filter(function (f) { return mapping[f.key] !== undefined && mapping[f.key] !== null && mapping[f.key] >= 0; });

    // Xác định định dạng ngày ở mức cột, để không cảnh báo lặp trên từng ô
    var dateOrders = {};
    fields.forEach(function (f) {
      if ((f.rules || {}).type !== 'date') return;
      var col = mapping[f.key];
      var ord = opt.dateOrder && opt.dateOrder !== 'auto' ? opt.dateOrder : inferDateOrder(ds, col);
      var header = ds.headers[col] ? ds.headers[col].name : f.label;
      if (ord === 'MIXED') {
        issues.push(issue({
          sheet: sheet, rowIndex: -1, rowNo: ds.headerRowIdx + 1, col: col, column: header, field: f.key,
          value: '', severity: 'error', ruleType: 'DATE_AMBIGUOUS',
          message: 'Cột "' + header + '" chứa cả hai kiểu ngày dd/MM và MM/dd. Không thể xác định đúng ngày cho mọi dòng.',
          suggestion: 'Yêu cầu khách hàng xuất lại cột này theo một định dạng duy nhất, tốt nhất là dd/MM/yyyy.'
        }));
        dateOrders[f.key] = 'DMY';
      } else if (ord === 'UNKNOWN') {
        issues.push(issue({
          sheet: sheet, rowIndex: -1, rowNo: ds.headerRowIdx + 1, col: col, column: header, field: f.key,
          value: '', severity: 'warning', ruleType: 'DATE_AMBIGUOUS',
          message: 'Cột "' + header + '" không có giá trị nào cho biết ngày hay tháng đứng trước. Hệ thống đang hiểu theo dd/MM/yyyy.',
          suggestion: 'Xác nhận định dạng ngày với khách hàng trước khi import.'
        }));
        dateOrders[f.key] = 'DMY';
      } else if (ord === 'NONE') {
        dateOrders[f.key] = 'DMY';
      } else {
        dateOrders[f.key] = ord;
      }
    });
    var uniqueMaps = {};   // fieldKey -> Map(value -> [rowIndex])
    fields.forEach(function (f) { if (f.rules.unique) uniqueMaps[f.key] = new Map(); });
    var groupMaps = (ruleSet.uniqueGroups || []).map(function () { return new Map(); });

    function colOf(key) { return mapping[key]; }
    function valOf(row, key) { var c = colOf(key); return c === undefined || c < 0 ? null : row[c]; }
    function push(o) { if (issues.length < MAX_ISSUES_PER_RULE * 4) issues.push(issue(o)); }

    function checkField(f, row, i) {
      var c = colOf(f.key), rules = f.rules || {}, v = row[c];
      var header = ds.headers[c] ? ds.headers[c].name : f.label;
      var base = { sheet: sheet, rowIndex: i, rowNo: ds.rowNo[i], col: c, column: header, field: f.key };
      var raw = U.toStr(v);
      var blank = U.isBlank(v);

      if (blank) {
        if (rules.required) {
          push(Object.assign({}, base, {
            value: '', severity: sev(rules, 'required', 'error'), ruleType: 'REQUIRED',
            message: f.label + ' bắt buộc nhập.', suggestion: 'Bổ sung giá trị cho ' + f.label.toLowerCase() + '.'
          }));
        } else if (rules.conditionalRequired) {
          var cond = rules.conditionalRequired;
          var other = U.toStr(valOf(row, cond.field)).trim();
          var match = (cond.equals || []).some(function (x) { return U.normKey(x) === U.normKey(other); });
          if (match) {
            push(Object.assign({}, base, {
              value: '', severity: 'error', ruleType: 'CONDITIONAL',
              message: f.label + ' bắt buộc khi ' + cond.field + ' = "' + other + '".',
              suggestion: 'Bổ sung ' + f.label.toLowerCase() + ' hoặc kiểm tra lại giá trị của ' + cond.field + '.'
            }));
          }
        }
        return;
      }

      var s = raw;
      var trimmed = s.trim();

      if (rules.trim !== false && U.hasEdgeSpace(s)) {
        push(Object.assign({}, base, {
          value: s, severity: 'warning', ruleType: 'WHITESPACE',
          message: 'Giá trị có khoảng trắng đầu/cuối.', suggestion: trimmed
        }));
      }
      if (rules.noDoubleSpace && U.hasDoubleSpace(s)) {
        push(Object.assign({}, base, {
          value: s, severity: 'warning', ruleType: 'DOUBLE_SPACE',
          message: 'Giá trị có khoảng trắng kép.', suggestion: s.replace(/\s+/g, ' ').trim()
        }));
      }
      if (rules.notAllNumeric && /^[\d\s.,-]+$/.test(trimmed)) {
        push(Object.assign({}, base, {
          value: s, severity: 'error', ruleType: 'DATA_TYPE',
          message: f.label + ' không được toàn số.', suggestion: 'Kiểm tra lại dữ liệu cột này, có thể bị lệch cột.'
        }));
      }
      if (rules.minLength !== undefined && trimmed.length < rules.minLength) {
        push(Object.assign({}, base, {
          value: s, severity: sev(rules, 'length', 'error'), ruleType: 'LENGTH',
          message: f.label + ' ngắn hơn ' + rules.minLength + ' ký tự (hiện ' + trimmed.length + ').', suggestion: ''
        }));
      }
      if (rules.maxLength !== undefined && trimmed.length > rules.maxLength) {
        push(Object.assign({}, base, {
          value: s, severity: sev(rules, 'length', 'error'), ruleType: 'LENGTH',
          message: f.label + ' dài ' + trimmed.length + ' ký tự, vượt giới hạn ' + rules.maxLength + '.',
          suggestion: 'Rút gọn hoặc kiểm tra lại dữ liệu.'
        }));
      }

      switch (rules.type) {
        case 'number': {
          var n = U.toNumber(v);
          if (isNaN(n)) {
            push(Object.assign({}, base, {
              value: s, severity: sev(rules, 'type', 'error'), ruleType: 'DATA_TYPE',
              message: f.label + ' phải là số, giá trị hiện tại không đọc được thành số.', suggestion: ''
            }));
          } else {
            if (rules.integer && !Number.isInteger(n)) {
              push(Object.assign({}, base, {
                value: s, severity: sev(rules, 'type', 'error'), ruleType: 'DATA_TYPE',
                message: f.label + ' phải là số nguyên.', suggestion: String(Math.round(n))
              }));
            }
            if (rules.min !== undefined && n < rules.min || rules.max !== undefined && n > rules.max) {
              push(Object.assign({}, base, {
                value: s, severity: sev(rules, 'range', 'error'), ruleType: 'RANGE',
                message: f.label + ' phải nằm trong khoảng ' + rules.min + ' – ' + rules.max + '.', suggestion: ''
              }));
            }
          }
          break;
        }
        case 'date': {
          var d = U.parseDate(v, { order: dateOrders[f.key] || 'auto' });
          if (!d.ok) {
            push(Object.assign({}, base, {
              value: s, severity: sev(rules, 'type', 'error'), ruleType: 'INVALID_DATE',
              message: d.reason === 'invalid'
                ? 'Ngày không tồn tại trên lịch.'
                : 'Không đọc được giá trị này thành ngày.',
              suggestion: 'Dùng định dạng dd/MM/yyyy, ví dụ 14/08/2026.'
            }));
          } else {
            if (rules.notFuture && d.date > today) {
              push(Object.assign({}, base, {
                value: s, severity: 'error', ruleType: 'DATE_FUTURE',
                message: f.label + ' lớn hơn ngày hiện tại.', suggestion: 'Kiểm tra lại năm.'
              }));
            }
            if (rules.minDate && d.date < new Date(rules.minDate)) {
              push(Object.assign({}, base, {
                value: s, severity: 'warning', ruleType: 'DATE_RANGE',
                message: f.label + ' nhỏ hơn mốc hợp lý (' + rules.minDate + ').', suggestion: 'Kiểm tra lại năm.'
              }));
            }
          }
          break;
        }
        case 'email': {
          if (!U.RE.email.test(trimmed)) {
            push(Object.assign({}, base, {
              value: s, severity: sev(rules, 'type', 'error'), ruleType: 'FORMAT',
              message: 'Email sai định dạng.', suggestion: 'Ví dụ hợp lệ: ten.sv@truong.edu.vn'
            }));
          }
          break;
        }
        case 'phone': {
          var digits = trimmed.replace(/[\s.()-]/g, '');
          if (!U.RE.phoneVN.test(digits)) {
            push(Object.assign({}, base, {
              value: s, severity: sev(rules, 'type', 'warning'), ruleType: 'FORMAT',
              message: 'Số điện thoại không đúng dạng 0xxxxxxxxx hoặc +84xxxxxxxxx.',
              suggestion: 'Kiểm tra lại, hoặc bỏ qua nếu là số quốc tế.'
            }));
          }
          break;
        }
      }

      if (rules.regex) {
        var re = new RegExp(rules.regex);
        if (!re.test(trimmed)) {
          push(Object.assign({}, base, {
            value: s, severity: sev(rules, 'regex', 'error'), ruleType: 'FORMAT',
            message: rules.regexMessage || (f.label + ' không đúng định dạng quy định.'), suggestion: ''
          }));
        }
      }

      if (rules.allowedValues && rules.allowedValues.length) {
        var nk = U.normKey(trimmed);
        var hit = rules.allowedValues.some(function (a) { return U.normKey(a) === nk; });
        if (!hit) {
          var mapped = rules.valueMap ? rules.valueMap[nk] || rules.valueMap[trimmed.toLowerCase()] : null;
          push(Object.assign({}, base, {
            value: s, severity: sev(rules, 'allowed', 'warning'), ruleType: 'ALLOWED_VALUES',
            message: 'Giá trị "' + trimmed + '" không thuộc danh sách cho phép: ' + rules.allowedValues.join(', ') + '.',
            suggestion: mapped ? mapped : 'Ánh xạ về một trong các giá trị cho phép.'
          }));
        }
      }

      if (rules.reference && refs && refs[f.key]) {
        if (!refs[f.key].set.has(trimmed.toLowerCase())) {
          push(Object.assign({}, base, {
            value: s, severity: sev(rules, 'reference', 'error'), ruleType: 'REFERENCE_EXISTS',
            message: 'Giá trị không tồn tại trong ' + (rules.reference.label || 'danh mục tham chiếu') +
              ' (' + refs[f.key].sheet + '.' + refs[f.key].column + ').',
            suggestion: 'Kiểm tra lại giá trị hoặc bổ sung bản ghi vào danh mục.'
          }));
        }
      }

      if (rules.unique) {
        var uk = rules.caseInsensitive ? trimmed.toLowerCase() : trimmed;
        var m = uniqueMaps[f.key];
        if (m.has(uk)) m.get(uk).push(i); else m.set(uk, [i]);
      }
    }

    function checkCross(row, i) {
      (ruleSet.crossFields || []).forEach(function (cf) {
        var lc = colOf(cf.left), rc = colOf(cf.right);
        if (lc === undefined || rc === undefined || lc < 0 || rc < 0) return;
        var lv = row[lc], rv = row[rc];
        if (U.isBlank(lv) || U.isBlank(rv)) return;
        var a, b;
        if (cf.type === 'date') {
          var da = U.parseDate(lv, { order: dateOrders[cf.left] || 'auto' });
          var db = U.parseDate(rv, { order: dateOrders[cf.right] || 'auto' });
          if (!da.ok || !db.ok) return;
          a = da.date.getTime(); b = db.date.getTime();
        } else {
          a = U.toNumber(lv); b = U.toNumber(rv);
          if (isNaN(a) || isNaN(b)) return;
        }
        var bad = (cf.op === '<=' && !(a <= b)) || (cf.op === '<' && !(a < b)) ||
          (cf.op === '>=' && !(a >= b)) || (cf.op === '>' && !(a > b));
        if (bad) {
          push({
            sheet: sheet, rowIndex: i, rowNo: ds.rowNo[i], col: rc,
            column: ds.headers[rc] ? ds.headers[rc].name : cf.right, field: cf.right,
            value: U.toStr(rv), severity: cf.severity || 'error', ruleType: 'CROSS_FIELD',
            message: cf.message || (cf.left + ' phải ' + cf.op + ' ' + cf.right + '.'),
            suggestion: 'Kiểm tra lại cặp giá trị ' + cf.left + ' / ' + cf.right + '.'
          });
        }
      });
    }

    function checkGroups(row, i) {
      (ruleSet.uniqueGroups || []).forEach(function (g, gi) {
        var parts = [], ok = true;
        g.fields.forEach(function (k) {
          var c = colOf(k);
          if (c === undefined || c < 0) { ok = false; return; }
          var v = row[c];
          if (U.isBlank(v)) ok = false;
          parts.push(U.toStr(v).trim().toLowerCase());
        });
        if (!ok) return;
        var key = parts.join('\u0001');
        var m = groupMaps[gi];
        if (m.has(key)) m.get(key).push(i); else m.set(key, [i]);
      });
    }

    return chunkLoop(ds.rows.length, function (i) {
      var row = ds.rows[i];
      for (var f = 0; f < fields.length; f++) checkField(fields[f], row, i);
      checkCross(row, i);
      checkGroups(row, i);
    }, onProgress, token, 'Đang kiểm tra dữ liệu').then(function () {
      // Phát hiện trùng sau khi đã quét hết
      fields.forEach(function (f) {
        var m = uniqueMaps[f.key]; if (!m) return;
        var c = colOf(f.key);
        var header = ds.headers[c] ? ds.headers[c].name : f.label;
        m.forEach(function (rowsIdx, key) {
          if (rowsIdx.length < 2) return;
          var rowNos = rowsIdx.map(function (r) { return ds.rowNo[r]; });
          rowsIdx.forEach(function (r) {
            push({
              sheet: sheet, rowIndex: r, rowNo: ds.rowNo[r], col: c, column: header, field: f.key,
              value: U.toStr(ds.rows[r][c]),
              severity: f.rules.uniqueSeverity || f.rules.severity || 'error', ruleType: 'DUPLICATE',
              message: f.label + ' bị trùng, xuất hiện ' + rowsIdx.length + ' lần (dòng ' +
                rowNos.slice(0, 6).join(', ') + (rowNos.length > 6 ? '…' : '') + ').',
              suggestion: 'Giữ lại một bản ghi hoặc sửa mã cho đúng.',
              relatedRows: rowNos
            });
          });
        });
      });

      (ruleSet.uniqueGroups || []).forEach(function (g, gi) {
        var m = groupMaps[gi];
        var c = colOf(g.fields[0]);
        m.forEach(function (rowsIdx) {
          if (rowsIdx.length < 2) return;
          var rowNos = rowsIdx.map(function (r) { return ds.rowNo[r]; });
          rowsIdx.forEach(function (r) {
            push({
              sheet: sheet, rowIndex: r, rowNo: ds.rowNo[r], col: c,
              column: ds.headers[c] ? ds.headers[c].name : g.fields[0], field: g.fields.join('+'),
              value: g.fields.map(function (k) { return U.toStr(ds.rows[r][colOf(k)]); }).join(' | '),
              severity: 'error', ruleType: 'DUPLICATE_GROUP',
              message: g.message || ('Trùng tổ hợp ' + g.fields.join(' + ') + '.'),
              suggestion: 'Kiểm tra các dòng: ' + rowNos.slice(0, 8).join(', ') + '.',
              relatedRows: rowNos
            });
          });
        });
      });

      return { issues: issues };
    });
  }

  /* ================= SUMMARY & INDEX ================= */
  function summarize(ds, issues) {
    var errorRows = new Set(), warnRows = new Set();
    var byRule = {}, byColumn = {};
    var e = 0, w = 0, inf = 0;
    issues.forEach(function (is) {
      if (is.severity === 'error') { e++; if (is.rowIndex >= 0) errorRows.add(is.rowIndex); }
      else if (is.severity === 'warning') { w++; if (is.rowIndex >= 0) warnRows.add(is.rowIndex); }
      else inf++;
      var rk = is.ruleType;
      if (!byRule[rk]) byRule[rk] = { rule: rk, error: 0, warning: 0, info: 0, total: 0 };
      byRule[rk][is.severity]++; byRule[rk].total++;
      var ck = is.column || '(bảng)';
      if (!byColumn[ck]) byColumn[ck] = { column: ck, col: is.col, error: 0, warning: 0, info: 0, total: 0 };
      byColumn[ck][is.severity]++; byColumn[ck].total++;
    });
    var total = ds.rows.length;
    var invalid = errorRows.size;
    var quality = total ? Math.round((1 - invalid / total) * 1000) / 10 : 100;
    return {
      totalRows: total, validRows: total - invalid, invalidRows: invalid,
      warningRows: warnRows.size, errorCount: e, warningCount: w, infoCount: inf,
      qualityScore: quality,
      errorRowSet: errorRows, warnRowSet: warnRows,
      issuesByRule: Object.keys(byRule).map(function (k) { return byRule[k]; }).sort(function (a, b) { return b.total - a.total; }),
      issuesByColumn: Object.keys(byColumn).map(function (k) { return byColumn[k]; }).sort(function (a, b) { return b.total - a.total; })
    };
  }

  /* Chỉ mục ô -> lỗi, phục vụ tô màu grid */
  function indexIssues(issues) {
    var cell = new Map(), row = new Map();
    issues.forEach(function (is) {
      if (is.rowIndex < 0) return;
      var ck = is.rowIndex + '|' + is.col;
      var cur = cell.get(ck);
      if (!cur) { cur = { severity: is.severity, list: [] }; cell.set(ck, cur); }
      cur.list.push(is);
      if (is.severity === 'error') cur.severity = 'error';
      else if (is.severity === 'warning' && cur.severity !== 'error') cur.severity = 'warning';
      var rr = row.get(is.rowIndex);
      if (!rr) { rr = { error: 0, warning: 0, info: 0 }; row.set(is.rowIndex, rr); }
      rr[is.severity]++;
    });
    return { cell: cell, row: row };
  }

  root.ASC = root.ASC || {};
  root.ASC.engine = {
    buildDataset: buildDataset, profileColumn: profileColumn, quickCheck: quickCheck,
    validate: validate, summarize: summarize, indexIssues: indexIssues,
    buildReferenceSet: buildReferenceSet, colLetter: colLetter, looksLikeCode: looksLikeCode,
    inferDateOrder: inferDateOrder
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = root.ASC.engine;
})(typeof window !== 'undefined' ? window : globalThis);
