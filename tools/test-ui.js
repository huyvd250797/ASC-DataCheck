/* Smoke test giao diện bằng jsdom: mở app, tải tệp mẫu, chọn sheet, chạy kiểm tra, xem kết quả.
   Chạy: npm i jsdom && node tools/test-ui.js  */
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const ROOT = path.join(__dirname, '..');
let fail = 0;
function check(name, ok, extra) { ok ? null : fail++; console.log(`${ok ? '  ok  ' : ' FAIL '} ${name}${extra ? ' — ' + extra : ''}`); }
const wait = ms => new Promise(r => setTimeout(r, ms));

(async function () {
  const dom = await JSDOM.fromFile(path.join(ROOT, 'index.html'), {
    runScripts: 'dangerously', resources: 'usable', pretendToBeVisual: true
  });
  const { window } = dom;
  const doc = window.document;
  const $ = s => doc.querySelector(s);

  window.addEventListener('error', e => { console.error('LỖI JS:', e.error && e.error.stack || e.message); fail++; });
  await new Promise(r => window.addEventListener('load', r));
  await wait(1600); // qua splash

  console.log('\n[1] Khởi động');
  check('Không còn splash', !$('#splash'));
  check('Modal hướng dẫn tự mở lần đầu', !$('#helpModal').classList.contains('hidden'));
  check('Khóa cuộn nền khi mở modal', doc.body.classList.contains('modal-open'));

  // Click ra ngoài modal: KHÔNG được đóng
  $('#helpModal').dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
  await wait(50);
  check('Click ra ngoài không đóng modal', !$('#helpModal').classList.contains('hidden'));
  $('#helpCloseBtn').click();
  await wait(50);
  check('Nút Đóng đóng được modal', $('#helpModal').classList.contains('hidden'));
  check('Mở lại bằng icon ?', ($('#helpBtn').click(), !$('#helpModal').classList.contains('hidden')));
  $('#helpClose').click();

  console.log('\n[2] Đổi giao diện sáng/tối');
  const before = doc.documentElement.getAttribute('data-theme');
  $('#themeBtn').click();
  check('Toggle theme đổi trạng thái', doc.documentElement.getAttribute('data-theme') !== before);
  $('#themeBtn').click();

  console.log('\n[3] Tải tệp Excel');
  const buf = fs.readFileSync(path.join(ROOT, 'sample', 'SinhVien_EPU_Dot1.xlsx'));
  const file = new window.File([new Uint8Array(buf)], 'SinhVien_EPU_Dot1.xlsx', {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const input = $('#fileInput');
  Object.defineProperty(input, 'files', { value: [file], configurable: true });
  input.dispatchEvent(new window.Event('change'));
  await wait(3000);

  check('Chuyển sang màn hình sheet', $('#viewSheet').classList.contains('active'));
  check('Hiện tên tệp', $('#fbName').textContent.includes('SinhVien_EPU_Dot1'));
  check('Liệt kê 2 sheet', doc.querySelectorAll('#sheetGrid .sheet-card').length === 2);

  console.log('\n[4] Chọn sheet và xem trước');
  doc.querySelectorAll('#sheetGrid .sheet-card')[0].click();
  await wait(1200);
  check('Dòng tiêu đề tự nhận là dòng 1', $('#headerRowSel').value === '0');
  check('Grid xem trước có render dòng', doc.querySelectorAll('#previewHost .grid-row').length > 0,
    doc.querySelectorAll('#previewHost .grid-row').length + ' dòng');
  check('Có ghi chú cấu trúc', $('#structureNotes').textContent.includes('10.000'));

  console.log('\n[5] Kiểm tra nhanh');
  $('#modeQuick').click();
  await wait(4000);
  check('Chuyển sang màn hình kết quả', $('#viewResult').classList.contains('active'));
  check('Có thẻ tổng hợp', doc.querySelectorAll('#summaryCards .scard').length === 5);
  check('Cột trái liệt kê loại lỗi', doc.querySelectorAll('#rail .rail-item').length > 3,
    doc.querySelectorAll('#rail .rail-item').length + ' mục');
  check('Grid kết quả có render', doc.querySelectorAll('#gridHost .grid-row').length > 0);
  check('Thanh trạng thái có số liệu', $('#sbRows').textContent === '10.000', $('#sbRows').textContent);

  console.log('\n[6] Lọc theo loại lỗi');
  const totalBefore = $('#resultCount').textContent;
  doc.querySelector('#rail .rail-item').click();
  await wait(300);
  check('Bộ lọc đổi số dòng hiển thị', $('#resultCount').textContent !== totalBefore,
    totalBefore + ' → ' + $('#resultCount').textContent);
  check('Hiện chip bộ lọc', !!$('#activeFilterHost .filter-chip'));
  check('Panel chi tiết có nội dung', $('#detail').textContent.length > 40);
  check('Ô lỗi được tô màu trong grid', doc.querySelectorAll('#gridHost .gc.warn, #gridHost .gc.err').length > 0,
    doc.querySelectorAll('#gridHost .gc.warn, #gridHost .gc.err').length + ' ô');
  $('#activeFilterHost .filter-chip button').click();
  await wait(200);
  check('Bỏ lọc trở lại đầy đủ', $('#resultCount').textContent === totalBefore);

  console.log('\n[7] Tab lỗi và tìm kiếm');
  doc.querySelector('#resultTabs .tab[data-tab="error"]').click();
  await wait(300);
  const errCount = $('#resultCount').textContent;
  check('Tab Lỗi lọc được', errCount !== totalBefore, errCount);
  doc.querySelector('#resultTabs .tab[data-tab="all"]').click();
  await wait(200);
  $('#searchInput').value = 'SV20500';
  $('#searchInput').dispatchEvent(new window.Event('input'));
  await wait(500);
  check('Tìm kiếm ra kết quả', /^[1-9]/.test($('#resultCount').textContent), $('#resultCount').textContent);
  $('#searchInput').value = '';
  $('#searchInput').dispatchEvent(new window.Event('input'));
  await wait(400);

  console.log('\n[8] Chi tiết ô lỗi và thống kê cột');
  doc.querySelector('#rail .rail-item').click();
  await wait(300);
  const badCell = doc.querySelector('#gridHost .gc.warn, #gridHost .gc.err');
  badCell.dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
  await wait(200);
  check('Panel hiện chi tiết lỗi', $('#detail').querySelector('.sev-badge') !== null);
  doc.querySelectorAll('#gridHost .grid-head .gh-cell')[2].click();
  await wait(300);
  check('Panel hiện thống kê cột', $('#detail').textContent.includes('Giá trị khác nhau'));

  console.log('\n[9] Lịch sử và modal export');
  $('#exportBtn').click();
  await wait(200);
  check('Modal export mở', !$('#exportModal').classList.contains('hidden'));
  check('Có 5 lựa chọn export', doc.querySelectorAll('#exportBody .export-opt').length === 5);
  $('#exportCloseBtn').click();
  check('Modal export đóng', $('#exportModal').classList.contains('hidden'));

  console.log('\n[10] Kiểm tra theo bộ rule');
  doc.querySelector('.brand').click();
  await wait(200);
  $('#modeRule').click();
  await wait(1500);
  check('Vào màn hình ghép cột', $('#viewMapping').classList.contains('active'));
  check('Chọn đúng bộ rule sinh viên', $('#ruleSetSel').value === 'SINHVIEN_ASC', $('#ruleSetSel').value);
  const mapped = Array.from(doc.querySelectorAll('#mapBody .map-status.ok')).length;
  check('Ghép tự động được ≥ 8 field', mapped >= 8, mapped + ' field');
  check('Tự chọn danh mục tham chiếu LopHoc', $('#refBody').textContent.includes('Sẽ đối chiếu'));
  check('Nút kiểm tra được bật', !$('#runValidationBtn').disabled);

  $('#runValidationBtn').click();
  await wait(5000);
  check('Ra màn hình kết quả', $('#viewResult').classList.contains('active'));
  check('Có lỗi tham chiếu mã lớp', $('#rail').textContent.includes('Không có trong danh mục'));
  check('Trạng thái sẵn sàng là "chưa import được"', $('#readiness').textContent.includes('Chưa import'));
  const hist = doc.querySelectorAll('#recentList .recent-row').length;
  check('Lịch sử được ghi lại', hist >= 1, hist + ' phiên');

  console.log(`\n${fail === 0 ? '✓ GIAO DIỆN ĐẠT' : '✕ CÓ LỖI GIAO DIỆN'} — ${fail} lỗi\n`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
