/* Sinh tệp Excel mẫu để thử ASC-DataCheck.
   Chạy: npm i xlsx && node tools/make-sample.js   */
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const OUT = path.join(__dirname, '..', 'sample');
fs.mkdirSync(OUT, { recursive: true });

let seed = 20260814;
function rnd() { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; }
function pick(a) { return a[Math.floor(rnd() * a.length)]; }

const HO = ['Nguyễn', 'Trần', 'Lê', 'Phạm', 'Hoàng', 'Vũ', 'Đặng', 'Bùi', 'Đỗ', 'Hồ', 'Ngô', 'Dương'];
const DEM = ['Văn', 'Thị', 'Hữu', 'Đức', 'Minh', 'Thanh', 'Quang', 'Ngọc', 'Xuân', 'Hải'];
const TEN = ['An', 'Bình', 'Chi', 'Dũng', 'Giang', 'Hà', 'Hùng', 'Khánh', 'Linh', 'Mai', 'Nam', 'Phong', 'Quân', 'Sơn', 'Trang', 'Tú', 'Việt', 'Yến'];
const NGANH = [['CNTT', 'Công nghệ thông tin'], ['KTPM', 'Kỹ thuật phần mềm'], ['HTTT', 'Hệ thống thông tin'], ['DTVT', 'Điện tử viễn thông'], ['QTKD', 'Quản trị kinh doanh'], ['KTOAN', 'Kế toán']];

/* ---------- Sheet LopHoc ---------- */
const lopRows = [['MaLop', 'TenLop', 'MaNganh', 'Khoa', 'NienKhoa', 'SiSo', 'NgayBatDau', 'NgayKetThuc']];
const maLops = [];
NGANH.forEach(([ma, ten]) => {
  for (let k = 1; k <= 7; k++) {
    const code = `D22${ma}${String(k).padStart(2, '0')}`;
    maLops.push(code);
    lopRows.push([code, `${ten} ${k} - K22`, ma, ten, '2022-2026', 30 + Math.floor(rnd() * 25), '05/09/2022', '30/06/2026']);
  }
});
// một lớp có ngày bắt đầu sau ngày kết thúc
lopRows[3][6] = '30/06/2026'; lopRows[3][7] = '05/09/2022';

/* ---------- Sheet SinhVien ---------- */
const N = 10000;
const svRows = [['STT', 'MSSV', 'Họ và tên', 'Ngày sinh', 'Giới tính', 'Lớp', 'Email', 'SĐT', 'CCCD', 'Trạng thái', 'Ngày tốt nghiệp']];
const codes = [];
for (let i = 1; i <= N; i++) {
  const code = 'SV' + String(20000 + i);
  codes.push(code);
  const ten = `${pick(HO)} ${pick(DEM)} ${pick(TEN)}`;
  const ngay = `${String(1 + Math.floor(rnd() * 28)).padStart(2, '0')}/${String(1 + Math.floor(rnd() * 12)).padStart(2, '0')}/${2002 + Math.floor(rnd() * 4)}`;
  svRows.push([
    i, code, ten, ngay, rnd() > 0.45 ? 'Nam' : 'Nữ', pick(maLops),
    `${code.toLowerCase()}@sv.epu.edu.vn`, '09' + String(10000000 + Math.floor(rnd() * 89999999)),
    String(100000000000 + Math.floor(rnd() * 899999999999)),
    rnd() > 0.02 ? 'Đang học' : 'Bảo lưu', ''
  ]);
}
const R = i => svRows[i]; // 1-based (dòng 1 là header)

// 50 mã trùng
for (let k = 0; k < 50; k++) R(120 + k * 37)[1] = R(100 + k * 37)[1];
// 50 mã trống
for (let k = 0; k < 50; k++) R(300 + k * 41)[1] = '';
// 20 ngày sinh không hợp lệ
const badDates = ['31/02/2005', '32/01/2005', '15/13/2004', '00/05/2003', 'không rõ'];
for (let k = 0; k < 20; k++) R(500 + k * 53)[3] = badDates[k % badDates.length];
// 20 email sai
for (let k = 0; k < 20; k++) R(700 + k * 59)[6] = k % 2 ? 'abc' : 'sv' + k + '@@gmail';
// 30 mã lớp không tồn tại
for (let k = 0; k < 30; k++) R(900 + k * 61)[5] = 'D22XX' + String(10 + k);
// 100 khoảng trắng thừa
for (let k = 0; k < 100; k++) { const r = R(1500 + k * 43); r[1] = ' ' + r[1] + ' '; }
// 20 giới tính chưa chuẩn hóa
for (let k = 0; k < 20; k++) R(2500 + k * 67)[4] = k % 2 ? 'Male' : 'Female';
// 15 số điện thoại sai định dạng
for (let k = 0; k < 15; k++) R(3000 + k * 71)[7] = '12345';
// 10 trạng thái tốt nghiệp thiếu ngày tốt nghiệp
for (let k = 0; k < 10; k++) R(3500 + k * 73)[9] = 'Tốt nghiệp';
// 5 họ tên toàn số (lệch cột)
for (let k = 0; k < 5; k++) R(4000 + k * 79)[2] = '123456';

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(svRows), 'SinhVien');
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(lopRows), 'LopHoc');
XLSX.writeFile(wb, path.join(OUT, 'SinhVien_EPU_Dot1.xlsx'));

/* ---------- Tệp cấu trúc xấu ---------- */
const messy = [
  ['DANH SÁCH SINH VIÊN LỚP D22CNTT01', null, null, null, null],
  ['(Kèm theo công văn số 145/ĐHEPU ngày 02/08/2026)', null, null, null, null],
  ['STT', 'Mã SV', 'Họ và tên', 'Mã lớp', 'Mã lớp'],
  [1, '00123', 'Nguyễn Văn An', 'D22CNTT01', 'D22CNTT01'],
  [2, '00124', 'Trần Thị Bình', 'D22CNTT01', 'D22CNTT01'],
  [null, null, null, null, null],
  [3, 123, 'Lê Hữu Cường', 'D22CNTT01', 'D22CNTT01'],
  [4, 124, '', 'D22CNTT02', 'D22CNTT02'],
  [5, 125, 'Phạm  Ngọc   Dung', 'D22CNTT01', 'D22CNTT01'],
  ['Tổng cộng: 5', null, null, null, null]
];
const wsM = XLSX.utils.aoa_to_sheet(messy);
wsM['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } }];
const wb2 = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb2, wsM, 'DanhSach');
XLSX.writeFile(wb2, path.join(OUT, 'DuLieu_Cau_Truc_Xau.xlsx'));

console.log('Đã tạo tệp mẫu trong', OUT);
