/* ASC-DataCheck — Xuất kết quả (ExcelJS) */
(function (root) {
  'use strict';
  var U = root.ASC.util;

  var FILL = {
    error: 'FFFFD4D6', warning: 'FFFDE7C3', header: 'FF1F3350', valid: 'FFDFF5E9'
  };
  var MAX_NOTES = 600;

  function baseName(fileName) {
    return String(fileName || 'DuLieu').replace(/\.[^.]+$/, '');
  }

  function cellValue(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    return v;
  }

  function styleHeader(ws, count) {
    var row = ws.getRow(1);
    row.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: FILL.header } };
    row.alignment = { vertical: 'middle' };
    row.height = 20;
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: Math.max(1, count) } };
  }

  function autoWidth(ws, headers, rows, maxSample) {
    var widths = headers.map(function (h) { return String(h).length + 3; });
    var n = Math.min(rows.length, maxSample || 300);
    for (var i = 0; i < n; i++) {
      for (var c = 0; c < headers.length; c++) {
        var L = U.toStr(rows[i][c]).length + 2;
        if (L > widths[c]) widths[c] = L;
      }
    }
    ws.columns.forEach(function (col, i) { col.width = Math.max(9, Math.min(46, widths[i] || 12)); });
  }

  /** ctx: {fileName, sheetName, dataset, issues, summary, ruleSetName, mode, project} */
  function buildSummarySheet(wb, ctx) {
    var ws = wb.addWorksheet('DATACHECK_SUMMARY');
    var s = ctx.summary;
    var rows = [
      ['Tệp', ctx.fileName],
      ['Sheet', ctx.sheetName],
      ['Chế độ kiểm tra', ctx.mode === 'quick' ? 'Kiểm tra nhanh' : 'Kiểm tra theo bộ rule'],
      ['Bộ rule', ctx.ruleSetName || '—'],
      ['Dự án', ctx.project || '—'],
      ['Thời điểm kiểm tra', U.fmtDateTime(new Date())],
      [],
      ['Tổng số dòng', s.totalRows],
      ['Dòng hợp lệ', s.validRows],
      ['Dòng có lỗi', s.invalidRows],
      ['Số lỗi', s.errorCount],
      ['Số cảnh báo', s.warningCount],
      ['Điểm chất lượng (%)', s.qualityScore],
      [],
      ['LỖI THEO LOẠI', '']
    ];
    rows.forEach(function (r) { ws.addRow(r); });
    s.issuesByRule.forEach(function (r) {
      var label = (root.ASC.rules.typeLabels[r.rule] || r.rule);
      ws.addRow([label, r.total]);
    });
    ws.getColumn(1).width = 30; ws.getColumn(2).width = 42;
    ws.getRow(1).font = { bold: true };
    ws.getRow(15).font = { bold: true };
    return ws;
  }

  function buildErrorSheet(wb, ctx) {
    var ws = wb.addWorksheet('DATACHECK_ERRORS');
    var headers = ['STT', 'Sheet', 'Dòng', 'Cột', 'Giá trị', 'Mức độ', 'Loại rule', 'Nội dung', 'Gợi ý xử lý'];
    ws.addRow(headers);
    var sevText = { error: 'ERROR', warning: 'WARNING', info: 'INFO' };
    ctx.issues.forEach(function (is, i) {
      var r = ws.addRow([
        i + 1, is.sheet || ctx.sheetName, is.rowIndex < 0 ? '(tiêu đề)' : is.rowNo,
        is.column || '', U.toStr(is.value), sevText[is.severity] || is.severity,
        is.ruleType, is.message, is.suggestion || ''
      ]);
      var color = is.severity === 'error' ? 'FFC0272D' : (is.severity === 'warning' ? 'FF9A6206' : 'FF3E5C8A');
      r.getCell(6).font = { bold: true, color: { argb: color } };
    });
    styleHeader(ws, headers.length);
    ws.columns.forEach(function (c, i) { c.width = [6, 14, 8, 20, 26, 11, 20, 52, 40][i] || 16; });
    ws.getColumn(8).alignment = { wrapText: true, vertical: 'top' };
    ws.getColumn(9).alignment = { wrapText: true, vertical: 'top' };
    return ws;
  }

  function buildDataSheet(wb, ctx, opt) {
    opt = opt || {};
    var ds = ctx.dataset;
    var name = (ctx.sheetName || 'DATA').substring(0, 28).replace(/[\\\/\?\*\[\]:]/g, '_');
    var ws = wb.addWorksheet(name || 'DATA');
    var headers = ds.headers.map(function (h) { return h.name; });
    var withStatus = opt.withStatus !== false;
    ws.addRow(withStatus ? headers.concat(['DataCheck_Status', 'DataCheck_Message']) : headers);

    var rowFilter = opt.rowFilter || function () { return true; };
    var index = ctx.index;
    var notes = 0;
    var written = [];

    for (var i = 0; i < ds.rows.length; i++) {
      if (!rowFilter(i)) continue;
      var src = ds.rows[i];
      var vals = src.map(cellValue);
      var rstat = index.row.get(i);
      var status = rstat ? (rstat.error ? 'ERROR' : (rstat.warning ? 'WARNING' : 'INFO')) : 'VALID';
      var msgs = [];
      if (withStatus) {
        for (var c = 0; c < ds.headers.length; c++) {
          var hit = index.cell.get(i + '|' + c);
          if (hit) hit.list.forEach(function (is) { msgs.push(is.column + ': ' + is.message); });
        }
        vals = vals.concat([status, msgs.slice(0, 6).join(' | ')]);
      }
      var row = ws.addRow(vals);
      written.push(i);

      if (withStatus && rstat) {
        for (var c2 = 0; c2 < ds.headers.length; c2++) {
          var hit2 = index.cell.get(i + '|' + c2);
          if (!hit2) continue;
          var cell = row.getCell(c2 + 1);
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hit2.severity === 'error' ? FILL.error : FILL.warning } };
          if (notes < MAX_NOTES) {
            cell.note = { texts: [{ text: hit2.list.map(function (x) { return x.message; }).join('\n') }] };
            notes++;
          }
        }
        var st = row.getCell(ds.headers.length + 1);
        st.font = { bold: true, color: { argb: status === 'ERROR' ? 'FFC0272D' : (status === 'WARNING' ? 'FF9A6206' : 'FF177245') } };
      }
      // Định dạng ngày
      for (var c3 = 0; c3 < ds.headers.length; c3++) {
        if (src[c3] instanceof Date) row.getCell(c3 + 1).numFmt = 'dd/mm/yyyy';
      }
    }
    styleHeader(ws, headers.length + (withStatus ? 2 : 0));
    autoWidth(ws, withStatus ? headers.concat(['DataCheck_Status', 'DataCheck_Message']) : headers, ds.rows, 200);
    if (withStatus) {
      ws.getColumn(headers.length + 1).width = 16;
      ws.getColumn(headers.length + 2).width = 50;
    }
    return { ws: ws, count: written.length };
  }

  function save(wb, name) {
    return wb.xlsx.writeBuffer().then(function (buf) {
      U.download(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name);
      return name;
    });
  }

  function newWorkbook() {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'ASC-DataCheck';
    wb.created = new Date();
    return wb;
  }

  var Exporter = {
    /** Tệp gần giống bản gốc, có tô màu ô lỗi + sheet lỗi + sheet tổng hợp */
    excelWithErrors: function (ctx) {
      var wb = newWorkbook();
      buildDataSheet(wb, ctx, { withStatus: true });
      buildErrorSheet(wb, ctx);
      buildSummarySheet(wb, ctx);
      return save(wb, baseName(ctx.fileName) + '_DataCheck.xlsx');
    },

    invalidRows: function (ctx) {
      var wb = newWorkbook();
      var idx = ctx.index;
      var res = buildDataSheet(wb, ctx, {
        withStatus: true,
        rowFilter: function (i) { var r = idx.row.get(i); return !!(r && (r.error || r.warning)); }
      });
      buildErrorSheet(wb, ctx);
      buildSummarySheet(wb, ctx);
      if (res.count === 0) return Promise.reject(new Error('Không có dòng lỗi để xuất.'));
      return save(wb, baseName(ctx.fileName) + '_InvalidRows.xlsx');
    },

    validRows: function (ctx) {
      var wb = newWorkbook();
      var idx = ctx.index;
      var res = buildDataSheet(wb, ctx, {
        withStatus: false,
        rowFilter: function (i) { var r = idx.row.get(i); return !(r && r.error); }
      });
      if (res.count === 0) return Promise.reject(new Error('Không có dòng hợp lệ để xuất.'));
      return save(wb, baseName(ctx.fileName) + '_ValidRows.xlsx');
    },

    errorReport: function (ctx) {
      var wb = newWorkbook();
      buildSummarySheet(wb, ctx);
      buildErrorSheet(wb, ctx);
      return save(wb, baseName(ctx.fileName) + '_ErrorReport.xlsx');
    },

    /** Văn bản gọn để dán vào Teams/Zalo/Email */
    errorsToText: function (ctx, limit) {
      var lines = [];
      lines.push('ASC-DATACHECK — ' + ctx.fileName + ' / sheet ' + ctx.sheetName);
      lines.push('Kiểm tra lúc ' + U.fmtDateTime(new Date()) +
        ' — ' + U.fmtInt(ctx.summary.totalRows) + ' dòng, ' +
        U.fmtInt(ctx.summary.errorCount) + ' lỗi, ' + U.fmtInt(ctx.summary.warningCount) + ' cảnh báo.');
      lines.push('');
      var list = ctx.issues.slice(0, limit || 500);
      list.forEach(function (is, i) {
        lines.push((i + 1) + '. Dòng ' + (is.rowIndex < 0 ? '(tiêu đề)' : is.rowNo) + ' - ' + (is.column || '') +
          (is.severity === 'warning' ? ' [Cảnh báo]' : (is.severity === 'info' ? ' [Gợi ý]' : '')));
        lines.push('   Giá trị: ' + (U.toStr(is.value) === '' ? '(trống)' : U.toStr(is.value)));
        lines.push('   Lỗi: ' + is.message);
        if (is.suggestion) lines.push('   Gợi ý: ' + is.suggestion);
      });
      if (ctx.issues.length > list.length) {
        lines.push('');
        lines.push('… còn ' + U.fmtInt(ctx.issues.length - list.length) + ' mục nữa. Xuất tệp Excel để xem đầy đủ.');
      }
      return lines.join('\n');
    }
  };

  root.ASC = root.ASC || {};
  root.ASC.exporter = Exporter;
})(typeof window !== 'undefined' ? window : globalThis);
