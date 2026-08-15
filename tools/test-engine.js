/* Test nhanh Validation Engine bằng Node (không cần trình duyệt).
   Chạy: node tools/test-engine.js  */
const XLSX = require('xlsx');
const path = require('path');

globalThis.window = undefined;
require('../js/util.js');
require('../js/rules.js');
require('../js/engine.js');
const U = globalThis.ASC.util, E = globalThis.ASC.engine, R = globalThis.ASC.rules;

let pass = 0, fail = 0;
function check(name, actual, expected) {
  const ok = actual === expected;
  ok ? pass++ : fail++;
  console.log(`${ok ? '  ok  ' : ' FAIL '} ${name}: ${actual}${ok ? '' : ' (mong đợi ' + expected + ')'}`);
}
function near(name, actual, expected, tol) {
  const ok = Math.abs(actual - expected) <= tol;
  ok ? pass++ : fail++;
  console.log(`${ok ? '  ok  ' : ' FAIL '} ${name}: ${actual}${ok ? '' : ' (mong đợi ~' + expected + ')'}`);
}

/* ---------- 1. Hàm ngày tháng ---------- */
console.log('\n[1] Ngày tháng');
check('31/02/2026 không hợp lệ', U.parseDate('31/02/2026').ok, false);
check('14/08/2026 hợp lệ', U.fmtDate(U.parseDate('14/08/2026').date), '14/08/2026');
check('2026-08-14 hợp lệ', U.fmtDate(U.parseDate('2026-08-14').date), '14/08/2026');
check('25/12/2005 không mơ hồ', U.parseDate('25/12/2005').ambiguous, false);
check('01/02/2026 bị mơ hồ', U.parseDate('01/02/2026').ambiguous, true);
check('Excel serial 45200', U.fmtDate(U.serialToDate(45200)), '01/10/2023');
check('Số điện thoại 0912345678', U.RE.phoneVN.test('0912345678'), true);
check('Email sai abc@@gmail', U.RE.email.test('abc@@gmail'), false);
check('normKey Mã sinh viên', U.normKey('Mã sinh viên'), 'masinhvien');

/* ---------- 2. Dựng dataset từ tệp cấu trúc xấu ---------- */
console.log('\n[2] Cấu trúc bảng');
const wbM = XLSX.readFile(path.join(__dirname, '..', 'sample', 'DuLieu_Cau_Truc_Xau.xlsx'), { cellDates: true });
const aoaM = XLSX.utils.sheet_to_json(wbM.Sheets['DanhSach'], { header: 1, raw: true, defval: null, blankrows: true });
const dsM = E.buildDataset(aoaM, 2, {});
check('Header ở dòng 3 -> 5 cột', dsM.headers.length, 5);
check('Bỏ dòng trống + dòng tổng cộng', dsM.rows.length, 5);
check('Phát hiện trùng tên cột', dsM.headers.filter(h => h.duplicateHeader).length, 1);
check('Số dòng Excel của bản ghi đầu', dsM.rowNo[0], 4);

/* ---------- 3. Kiểm tra nhanh ---------- */
console.log('\n[3] Kiểm tra nhanh');
const wb = XLSX.readFile(path.join(__dirname, '..', 'sample', 'SinhVien_EPU_Dot1.xlsx'), { cellDates: true });
const aoa = XLSX.utils.sheet_to_json(wb.Sheets['SinhVien'], { header: 1, raw: true, defval: null, blankrows: true });
const ds = E.buildDataset(aoa, 0, {});
check('10.000 dòng dữ liệu', ds.rows.length, 10000);

const t0 = Date.now();
E.quickCheck(ds, { sheet: 'SinhVien' }, null, null).then(res => {
  const by = {};
  res.issues.forEach(i => { by[i.ruleType] = (by[i.ruleType] || 0) + 1; });
  console.log('  → ' + res.issues.length + ' phát hiện trong ' + (Date.now() - t0) + 'ms:', by);
  near('Khoảng trắng thừa ~100', by.WHITESPACE || 0, 100, 3);
  check('Không có cảnh báo ngày mơ hồ tràn lan', (by.DATE_AMBIGUOUS || 0) < 3, true);
  check('Có phát hiện ô trống MSSV', (by.BLANK || 0) >= 50, true);
  check('Có phát hiện giá trị trùng', (by.DUPLICATE || 0) >= 100, true);

  /* ---------- 4. Kiểm tra theo rule ---------- */
  console.log('\n[4] Kiểm tra theo bộ rule');
  const rs = R.byId('SINHVIEN_ASC');
  const mapping = {};
  const alias = {
    MaSinhVien: 'MSSV', HoTen: 'Họ và tên', NgaySinh: 'Ngày sinh', GioiTinh: 'Giới tính',
    MaLop: 'Lớp', Email: 'Email', SoDienThoai: 'SĐT', CCCD: 'CCCD',
    TrangThai: 'Trạng thái', NgayTotNghiep: 'Ngày tốt nghiệp'
  };
  Object.keys(alias).forEach(k => {
    const idx = ds.headers.findIndex(h => h.name === alias[k]);
    if (idx >= 0) mapping[k] = idx;
  });
  check('Ghép đủ 10 field', Object.keys(mapping).length, 10);

  const aoaLop = XLSX.utils.sheet_to_json(wb.Sheets['LopHoc'], { header: 1, raw: true, defval: null, blankrows: true });
  const dsLop = E.buildDataset(aoaLop, 0, {});
  const refs = { MaLop: { sheet: 'LopHoc', column: 'MaLop', set: E.buildReferenceSet(dsLop, 0) } };

  const t1 = Date.now();
  return E.validate(ds, rs, mapping, refs, null, null, { sheet: 'SinhVien' }).then(out => {
    const by2 = {};
    out.issues.forEach(i => { by2[i.ruleType] = (by2[i.ruleType] || 0) + 1; });
    console.log('  → ' + out.issues.length + ' phát hiện trong ' + (Date.now() - t1) + 'ms:', by2);
    check('Thiếu mã sinh viên = 50', by2.REQUIRED, 50);
    // Số bản ghi trùng thật sự (các field có rule unique: MSSV, Email, CCCD)
    let expDup = 0;
    ['MaSinhVien', 'Email', 'CCCD'].forEach(k => {
      const cnt = new Map();
      ds.rows.forEach(r => {
        const v = U.toStr(r[mapping[k]]).trim().toLowerCase();
        if (v) cnt.set(v, (cnt.get(v) || 0) + 1);
      });
      cnt.forEach(c => { if (c > 1) expDup += c; });
    });
    check('Mã trùng khớp số thực tế trong tệp', by2.DUPLICATE, expDup);
    check('Ngày sinh không hợp lệ = 20', by2.INVALID_DATE, 20);
    check('Email/SĐT sai định dạng = 35', by2.FORMAT, 35);
    check('Mã lớp không có trong danh mục = 30', by2.REFERENCE_EXISTS, 30);
    let expWs = 0;
    ds.rows.forEach(r => { if (U.hasEdgeSpace(U.toStr(r[mapping.MaSinhVien]))) expWs++; });
    check('Khoảng trắng thừa khớp số thực tế', by2.WHITESPACE, expWs);
    check('Giới tính ngoài danh sách = 20', by2.ALLOWED_VALUES, 20);
    check('Thiếu ngày tốt nghiệp theo điều kiện = 10', by2.CONDITIONAL, 10);
    check('Họ tên toàn số = 5', by2.DATA_TYPE, 5);

    const sum = E.summarize(ds, out.issues);
    console.log('  → Tổng hợp:', {
      rows: sum.totalRows, valid: sum.validRows, errors: sum.errorCount,
      warnings: sum.warningCount, quality: sum.qualityScore + '%'
    });
    check('Chất lượng nằm trong 95-99%', sum.qualityScore > 95 && sum.qualityScore < 99.5, true);

    const idx = E.indexIssues(out.issues);
    check('Chỉ mục ô có dữ liệu', idx.cell.size > 200, true);

    /* ---------- 5. Rule của sheet LopHoc ---------- */
    console.log('\n[5] Cross-field trên sheet LopHoc');
    const rsLop = R.byId('LOPHOC_ASC');
    const mapLop = {};
    dsLop.headers.forEach((h, i) => { if (rsLop.fields.find(f => f.key === h.name)) mapLop[h.name] = i; });
    return E.validate(dsLop, rsLop, mapLop, {}, null, null, { sheet: 'LopHoc' }).then(o2 => {
      const cf = o2.issues.filter(i => i.ruleType === 'CROSS_FIELD');
      check('Bắt được ngày bắt đầu > ngày kết thúc', cf.length, 1);
      console.log(`\n${fail === 0 ? '✓ TẤT CẢ ĐỀU ĐẠT' : '✕ CÓ TEST HỎNG'} — ${pass} đạt, ${fail} hỏng\n`);
      process.exit(fail ? 1 : 0);
    });
  });
}).catch(e => { console.error(e); process.exit(1); });
