/* Test luồng export bằng Node: dựng dữ liệu -> validate -> xuất 4 loại tệp -> đọc lại kiểm tra.
   Chạy: node tools/test-export.js  */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

globalThis.ExcelJS = require('exceljs');
require('../js/util.js');
require('../js/rules.js');
require('../js/engine.js');

const U = globalThis.ASC.util, E = globalThis.ASC.engine, R = globalThis.ASC.rules;
const OUT = path.join(__dirname, '..', '.tmp-export');
fs.mkdirSync(OUT, { recursive: true });

// Thay hàm tải xuống của trình duyệt bằng ghi tệp
U.download = function (blob, name) {
  return blob.arrayBuffer().then(b => fs.writeFileSync(path.join(OUT, name), Buffer.from(b)));
};
require('../js/exporter.js');
const X = globalThis.ASC.exporter;

const wb = XLSX.readFile(path.join(__dirname, '..', 'sample', 'SinhVien_EPU_Dot1.xlsx'), { cellDates: true });
const aoa = XLSX.utils.sheet_to_json(wb.Sheets['SinhVien'], { header: 1, raw: true, defval: null, blankrows: true });
const ds = E.buildDataset(aoa, 0, {});
const rs = R.byId('SINHVIEN_ASC');
const alias = { MaSinhVien: 'MSSV', HoTen: 'Họ và tên', NgaySinh: 'Ngày sinh', GioiTinh: 'Giới tính', MaLop: 'Lớp', Email: 'Email', SoDienThoai: 'SĐT', CCCD: 'CCCD', TrangThai: 'Trạng thái', NgayTotNghiep: 'Ngày tốt nghiệp' };
const mapping = {};
Object.keys(alias).forEach(k => { const i = ds.headers.findIndex(h => h.name === alias[k]); if (i >= 0) mapping[k] = i; });

const aoaLop = XLSX.utils.sheet_to_json(wb.Sheets['LopHoc'], { header: 1, raw: true, defval: null, blankrows: true });
const dsLop = E.buildDataset(aoaLop, 0, {});
const refs = { MaLop: { sheet: 'LopHoc', column: 'MaLop', set: E.buildReferenceSet(dsLop, 0) } };

let fail = 0;
function check(name, ok, extra) {
  ok ? null : fail++;
  console.log(`${ok ? '  ok  ' : ' FAIL '} ${name}${extra ? ' — ' + extra : ''}`);
}

E.validate(ds, rs, mapping, refs, null, null, { sheet: 'SinhVien' }).then(out => {
  const summary = E.summarize(ds, out.issues);
  const index = E.indexIssues(out.issues);
  const ctx = {
    fileName: 'SinhVien_EPU_Dot1.xlsx', sheetName: 'SinhVien', dataset: ds, issues: out.issues,
    summary, index, mode: 'rule', ruleSetName: rs.name, project: 'EPU'
  };

  const t0 = Date.now();
  return X.excelWithErrors(ctx)
    .then(() => { console.log('  → excelWithErrors: ' + (Date.now() - t0) + 'ms'); return X.invalidRows(ctx); })
    .then(() => X.validRows(ctx))
    .then(() => X.errorReport(ctx))
    .then(() => {
      const f1 = path.join(OUT, 'SinhVien_EPU_Dot1_DataCheck.xlsx');
      const w1 = XLSX.readFile(f1);
      check('Tệp đánh dấu lỗi mở được', true);
      check('Có sheet DATACHECK_ERRORS', w1.SheetNames.includes('DATACHECK_ERRORS'), w1.SheetNames.join(', '));
      check('Có sheet DATACHECK_SUMMARY', w1.SheetNames.includes('DATACHECK_SUMMARY'));
      const rows1 = XLSX.utils.sheet_to_json(w1.Sheets['SinhVien'], { header: 1 });
      check('Đủ số dòng dữ liệu', rows1.length === ds.rows.length + 1, rows1.length + ' dòng (kể cả tiêu đề)');
      check('Có cột DataCheck_Status', rows1[0].includes('DataCheck_Status'));
      const errRows = XLSX.utils.sheet_to_json(w1.Sheets['DATACHECK_ERRORS'], { header: 1 });
      check('Sheet lỗi đủ bản ghi', errRows.length === out.issues.length + 1, (errRows.length - 1) + '/' + out.issues.length);

      const w2 = XLSX.readFile(path.join(OUT, 'SinhVien_EPU_Dot1_InvalidRows.xlsx'));
      const inv = XLSX.utils.sheet_to_json(w2.Sheets['SinhVien'], { header: 1 });
      check('Chỉ xuất dòng có lỗi/cảnh báo', inv.length - 1 === index.row.size, (inv.length - 1) + '/' + index.row.size);

      const w3 = XLSX.readFile(path.join(OUT, 'SinhVien_EPU_Dot1_ValidRows.xlsx'));
      const val = XLSX.utils.sheet_to_json(w3.Sheets['SinhVien'], { header: 1 });
      check('Dòng hợp lệ đúng số lượng', val.length - 1 === summary.validRows, (val.length - 1) + '/' + summary.validRows);

      const txt = X.errorsToText(ctx, 5);
      check('Văn bản copy có nội dung lỗi', /Dòng \d+ - /.test(txt));
      console.log('\n--- Trích văn bản copy ---\n' + txt.split('\n').slice(0, 9).join('\n'));

      fs.readdirSync(OUT).forEach(f => console.log('  tệp: ' + f + ' — ' + (fs.statSync(path.join(OUT, f)).size / 1024).toFixed(0) + ' KB'));
      console.log(`\n${fail === 0 ? '✓ EXPORT ĐẠT' : '✕ EXPORT HỎNG'}\n`);
      process.exit(fail ? 1 : 0);
    });
}).catch(e => { console.error(e); process.exit(1); });
