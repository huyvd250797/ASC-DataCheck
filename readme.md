# ASC-DATACHECK v1.0.0

Công cụ kiểm tra, phân tích và chuẩn hóa dữ liệu Excel trước khi import hoặc migrate vào hệ thống quản trị đào tạo.

**Validate faster. Import safer.**

---

## Chạy ứng dụng

Không cần build, không cần cài đặt gì:

```
Mở tệp index.html bằng Chrome hoặc Edge
```

Toàn bộ thư viện đã nằm sẵn trong `vendor/`, app chạy được cả khi không có mạng.

Nếu muốn chạy qua HTTP (khuyến nghị khi deploy nội bộ hoặc khi trình duyệt chặn `localStorage` với `file://`):

```bash
python3 -m http.server 8080     # rồi mở http://localhost:8080
```

Deploy nội bộ: copy nguyên thư mục lên IIS/nginx/SharePoint — đây là static site thuần, không cần backend.

## Dữ liệu mẫu

| Tệp | Dùng để thử |
|---|---|
| `sample/SinhVien_EPU_Dot1.xlsx` | 10.000 sinh viên + sheet danh mục lớp, đã cài sẵn lỗi: 50 mã trùng, 50 mã trống, 20 ngày sinh sai, 20 email sai, 30 mã lớp không có trong danh mục, 100 khoảng trắng thừa, 20 giới tính chưa chuẩn hóa, 10 dòng thiếu ngày tốt nghiệp theo điều kiện |
| `sample/DuLieu_Cau_Truc_Xau.xlsx` | Tiêu đề ở dòng 3, ô gộp, dòng trống giữa bảng, dòng "Tổng cộng", trùng tên cột, mã bị Excel biến thành số |

Luồng thử nhanh: mở app → kéo `SinhVien_EPU_Dot1.xlsx` vào → chọn sheet `SinhVien` → **Kiểm tra theo bộ rule** → bộ rule "Sinh viên — chuẩn ASC" tự chọn, cột tự ghép, danh mục lớp tự đối chiếu → **Kiểm tra dữ liệu**.

## Chức năng V1

**Đọc tệp** — .xlsx/.xlsm/.xls/.csv, kéo thả, chọn sheet, tự nhận dòng tiêu đề (chọn lại được), bỏ dòng trống và dòng tổng cộng, cảnh báo ô gộp/cột ẩn/trùng tên cột. Tệp gốc không bao giờ bị sửa.

**Hai chế độ kiểm tra**

- *Kiểm tra nhanh*: không cần cấu hình — ô trống, giá trị trùng, khoảng trắng thừa/kép, kiểu dữ liệu lẫn lộn, cột mã bị lưu dạng số (nguy cơ mất số 0 đầu), số dài dạng khoa học, cột trống, giá trị bất thường (IQR), thống kê từng cột.
- *Kiểm tra theo bộ rule*: required, unique, unique tổ hợp, kiểu dữ liệu, ngày, độ dài, khoảng giá trị, regex, danh sách giá trị cho phép, tham chiếu sang sheet khác, cross-field, bắt buộc có điều kiện.

**Ghép cột** — tự động theo tên field, nhãn và alias; hiển thị trạng thái đã ghép / bỏ qua / chưa ghép; chặn chạy khi field bắt buộc chưa ghép. Danh mục tham chiếu tự dò sheet và cột, sửa tay được.

**Kết quả** — thẻ tổng hợp, tab Tất cả/Lỗi/Cảnh báo/Hợp lệ, cột trái phân loại lỗi theo rule và theo cột (bấm để lọc), grid ảo hóa tô màu ô lỗi, panel chi tiết trả lời đủ 5 câu: lỗi ở đâu, giá trị nào, sai điều kiện gì, mức độ nào, xử lý ra sao. Trạng thái sẵn sàng: *Chưa import được* / *Sẵn sàng — còn cảnh báo* / *Dữ liệu sẵn sàng import*.

**Xuất** — Excel có đánh dấu lỗi (tô màu ô + ghi chú + cột `DataCheck_Status` + sheet `DATACHECK_ERRORS` + `DATACHECK_SUMMARY`), chỉ dòng lỗi, chỉ dòng hợp lệ, báo cáo lỗi, và sao chép danh sách lỗi dạng văn bản để dán vào Teams/Zalo/email.

**Khác** — dark/light mode, icon `?` luôn hiện ở header mở hướng dẫn nhanh (modal chỉ đóng bằng X hoặc nút Đóng, khóa cuộn nền, tự mở lần đầu), lịch sử kiểm tra, phím tắt `Ctrl+F` tìm kiếm, `Ctrl+E` xuất, `?` mở hướng dẫn, `Ctrl+C` sao chép ô, `Ctrl+Shift+C` sao chép dòng.

## Quyền riêng tư

Dữ liệu được đọc và xử lý hoàn toàn trong trình duyệt. Không có backend, không có request mạng, không log nội dung ô. Lịch sử chỉ lưu metadata (tên tệp, số dòng, số lỗi, thời điểm) trong `localStorage`, xóa được bằng nút "Xóa lịch sử".

## Hiệu năng

Đo trên tệp mẫu 10.000 dòng × 11 cột:

| Việc | Thời gian |
|---|---|
| Kiểm tra nhanh (đủ profiling) | ~340 ms |
| Kiểm tra theo bộ rule (10 field + tham chiếu) | ~370 ms |
| Xuất Excel có đánh dấu lỗi | ~3 s |

Grid chỉ render những dòng đang nhìn thấy nên cuộn vẫn mượt với 50.000+ dòng. Validation chạy theo lô 2.000 dòng và nhả luồng giữa các lô, có tiến độ và nút Hủy.

## Cấu trúc mã nguồn

```
index.html              khung giao diện, các view và modal
css/styles.css          toàn bộ style, biến màu cho dark/light
js/util.js              chuẩn hóa chuỗi, đọc ngày, nhận kiểu dữ liệu, toast, storage
js/rules.js             bộ rule preset (dữ liệu thuần, xuất/nhập JSON được)
js/engine.js            validation engine — thuần logic, không đụng DOM
js/grid.js              data grid ảo hóa
js/exporter.js          xuất Excel bằng ExcelJS
js/app.js               điều phối màn hình
vendor/                 SheetJS (đọc) và ExcelJS (ghi có định dạng)
tools/                  script sinh dữ liệu mẫu và bộ test
```

`engine.js` không phụ thuộc DOM và không giữ trạng thái toàn cục, nên có thể đưa nguyên vào Web Worker khi app được phục vụ qua HTTP. V1 chưa dùng Worker vì Worker không chạy được với `file://` — thay vào đó validation chia lô và nhả luồng nên giao diện không đứng.

## Chạy test

Cần Node 18+ và ba gói chỉ dùng cho test:

```bash
npm install                 # xlsx, exceljs, jsdom
npm run sample              # sinh lại tệp mẫu
npm test                    # engine + export + giao diện
```

- `tools/test-engine.js` — 31 phép kiểm: ngày tháng, cấu trúc bảng, kiểm tra nhanh, đủ 9 loại rule, cross-field, điểm chất lượng.
- `tools/test-export.js` — xuất 4 loại tệp rồi đọc lại, đối chiếu số dòng và số bản ghi lỗi.
- `tools/test-ui.js` — chạy toàn bộ luồng giao diện bằng jsdom: tải tệp, chọn sheet, kiểm tra, lọc, xem chi tiết, ghép cột, xuất.

## Bộ rule

Mỗi bộ rule là dữ liệu thuần, xuất/nhập được bằng JSON ngay trên màn hình ghép cột, không cần sửa mã nguồn. Một field trông như sau:

```json
{
  "key": "MaSinhVien",
  "label": "Mã sinh viên",
  "aliases": ["MSSV", "Mã SV", "StudentCode"],
  "rules": {
    "required": true,
    "unique": true,
    "caseInsensitive": true,
    "maxLength": 50,
    "regex": "^[A-Za-z0-9._-]+$",
    "regexMessage": "Mã sinh viên chỉ nên gồm chữ, số và . _ -"
  }
}
```

Rule hỗ trợ: `required`, `unique` + `caseInsensitive` + `uniqueSeverity`, `trim`, `noDoubleSpace`, `notAllNumeric`, `minLength`/`maxLength`, `type` (`number`/`date`/`email`/`phone`), `integer`, `min`/`max`, `notFuture`, `minDate`, `regex` + `regexMessage`, `allowedValues` + `valueMap`, `reference`, `conditionalRequired`, `severity`. Ngoài field còn có `uniqueGroups` (trùng tổ hợp) và `crossFields` (so sánh hai cột).

Định dạng ngày được suy luận ở mức cột: chỉ cần một giá trị có phần đầu lớn hơn 12 là biết chắc cột đang là dd/MM. Nếu cả cột không có manh mối, app cảnh báo một lần cho cả cột thay vì cảnh báo từng ô; nếu cột lẫn cả hai kiểu thì báo lỗi.

## Chưa làm trong V1

Theo đúng phạm vi đã chốt: chưa có backend, đăng nhập, cloud, kết nối SQL Server, AI, Auto Fix, Rule Builder trên giao diện, so sánh hai lần kiểm tra, và chưa tích hợp ASC-GenScript/ASC-Config.

Định hướng tiếp theo:

- **v1.1** — lưu rule vào IndexedDB, so sánh hai lần kiểm tra (đã sửa được bao nhiêu, còn lại bao nhiêu), danh sách distinct value đầy đủ, gắn thẻ dự án cho lịch sử.
- **v1.5** — Auto Fix các lỗi an toàn (trim, gộp khoảng trắng, chuẩn hóa enum, chuẩn hóa ngày) kèm Undo, Rule Builder trên giao diện, rule riêng theo dự án.
- **v2** — đưa validation vào Web Worker khi chạy qua HTTP, chuyển dòng hợp lệ sang ASC-GenScript, gộp vào ASC-Toolbox.
