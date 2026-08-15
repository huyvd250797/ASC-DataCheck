/* ASC-DataCheck — Rule Set preset.
   Mỗi Rule Set mô tả bằng dữ liệu thuần (JSON-able) để có thể xuất/nhập, không hard-code trong engine. */
(function (root) {
  'use strict';

  var GENDER_MAP = {
    'male': 'Nam', 'm': 'Nam', 'nam': 'Nam', '1': 'Nam', 'boy': 'Nam',
    'female': 'Nữ', 'f': 'Nữ', 'nu': 'Nữ', '0': 'Nữ', 'girl': 'Nữ',
    'other': 'Khác', 'khac': 'Khác'
  };

  var RULE_SETS = [
    {
      id: 'SINHVIEN_ASC', name: 'Sinh viên — chuẩn ASC', version: '1.0',
      desc: 'Dùng cho danh sách sinh viên nhập học, migrate hồ sơ sinh viên. Mã lớp được đối chiếu với sheet danh mục lớp nếu có.',
      fields: [
        { key: 'MaSinhVien', label: 'Mã sinh viên', aliases: ['MSSV', 'Mã SV', 'Mã sinh viên', 'StudentCode', 'MaSV'],
          rules: { required: true, unique: true, caseInsensitive: true, trim: true, minLength: 1, maxLength: 50, regex: '^[A-Za-z0-9._-]+$', regexMessage: 'Mã sinh viên chỉ nên gồm chữ, số và các ký tự . _ -' } },
        { key: 'HoTen', label: 'Họ và tên', aliases: ['Họ tên', 'Họ và tên', 'FullName', 'Tên sinh viên', 'HoVaTen'],
          rules: { required: true, trim: true, noDoubleSpace: true, notAllNumeric: true, maxLength: 100 } },
        { key: 'NgaySinh', label: 'Ngày sinh', aliases: ['Ngày sinh', 'NTNS', 'DOB', 'BirthDate', 'Ngay sinh'],
          rules: { required: true, type: 'date', notFuture: true, minDate: '1940-01-01' } },
        { key: 'GioiTinh', label: 'Giới tính', aliases: ['Giới tính', 'Gender', 'Phái', 'Sex'],
          rules: { allowedValues: ['Nam', 'Nữ', 'Khác'], valueMap: GENDER_MAP, severity: 'warning' } },
        { key: 'MaLop', label: 'Mã lớp', aliases: ['Lớp', 'Mã lớp', 'MaLopHoc', 'Class', 'ClassCode', 'Lớp học'],
          rules: { required: true, trim: true, reference: { sheetHint: 'LopHoc', columnHint: 'MaLop', label: 'Danh mục lớp' } } },
        { key: 'Email', label: 'Email', aliases: ['Email', 'E-mail', 'Thư điện tử', 'Mail'],
          rules: { type: 'email', unique: true, uniqueSeverity: 'warning', trim: true } },
        { key: 'SoDienThoai', label: 'Số điện thoại', aliases: ['SĐT', 'Điện thoại', 'Phone', 'SoDT', 'Số điện thoại'],
          rules: { type: 'phone', severity: 'warning', trim: true } },
        { key: 'CCCD', label: 'CCCD/CMND', aliases: ['CCCD', 'CMND', 'Căn cước', 'Số CCCD', 'IDCard'],
          rules: { regex: '^\\d{9}$|^\\d{12}$', regexMessage: 'CCCD/CMND phải là 9 hoặc 12 chữ số.', unique: true, uniqueSeverity: 'warning', severity: 'warning' } },
        { key: 'TrangThai', label: 'Trạng thái', aliases: ['Trạng thái', 'Status', 'TinhTrang', 'Tình trạng'],
          rules: { allowedValues: ['Đang học', 'Bảo lưu', 'Thôi học', 'Tốt nghiệp'], severity: 'warning' } },
        { key: 'NgayTotNghiep', label: 'Ngày tốt nghiệp', aliases: ['Ngày tốt nghiệp', 'NgayTN', 'GraduationDate'],
          rules: { type: 'date', conditionalRequired: { field: 'TrangThai', equals: ['Tốt nghiệp'] } } }
      ],
      uniqueGroups: [],
      crossFields: [{ left: 'NgaySinh', op: '<=', right: 'NgayTotNghiep', type: 'date', message: 'Ngày tốt nghiệp phải sau ngày sinh.' }]
    },

    {
      id: 'LOPHOC_ASC', name: 'Lớp học — chuẩn ASC', version: '1.0',
      desc: 'Danh mục lớp. Thường được dùng làm danh mục tham chiếu cho sheet sinh viên.',
      fields: [
        { key: 'MaLop', label: 'Mã lớp', aliases: ['Mã lớp', 'MaLop', 'ClassCode', 'Lớp'],
          rules: { required: true, unique: true, caseInsensitive: true, trim: true, maxLength: 50 } },
        { key: 'TenLop', label: 'Tên lớp', aliases: ['Tên lớp', 'TenLop', 'ClassName'],
          rules: { required: true, trim: true, maxLength: 150 } },
        { key: 'MaNganh', label: 'Mã ngành', aliases: ['Ngành', 'Mã ngành', 'MaNganh', 'Major'],
          rules: { trim: true } },
        { key: 'Khoa', label: 'Khoa', aliases: ['Khoa', 'Faculty', 'Đơn vị'], rules: { trim: true } },
        { key: 'NienKhoa', label: 'Niên khóa', aliases: ['Niên khóa', 'NienKhoa', 'Khóa', 'Year'], rules: { trim: true } },
        { key: 'SiSo', label: 'Sĩ số', aliases: ['Sĩ số', 'SiSo', 'Số lượng'], rules: { type: 'number', integer: true, min: 0, max: 300, severity: 'warning' } },
        { key: 'NgayBatDau', label: 'Ngày bắt đầu', aliases: ['Ngày bắt đầu', 'TuNgay', 'StartDate'], rules: { type: 'date' } },
        { key: 'NgayKetThuc', label: 'Ngày kết thúc', aliases: ['Ngày kết thúc', 'DenNgay', 'EndDate'], rules: { type: 'date' } }
      ],
      uniqueGroups: [],
      crossFields: [{ left: 'NgayBatDau', op: '<=', right: 'NgayKetThuc', type: 'date', message: 'Ngày bắt đầu phải nhỏ hơn hoặc bằng ngày kết thúc.' }]
    },

    {
      id: 'MONHOC_ASC', name: 'Môn học — chuẩn ASC', version: '1.0',
      desc: 'Danh mục môn học/học phần.',
      fields: [
        { key: 'MaMonHoc', label: 'Mã môn học', aliases: ['Mã môn', 'Mã học phần', 'MaMonHoc', 'SubjectCode', 'MaHP'],
          rules: { required: true, unique: true, caseInsensitive: true, trim: true, maxLength: 50 } },
        { key: 'TenMonHoc', label: 'Tên môn học', aliases: ['Tên môn', 'Tên học phần', 'TenMonHoc', 'SubjectName'],
          rules: { required: true, trim: true, maxLength: 200 } },
        { key: 'SoTinChi', label: 'Số tín chỉ', aliases: ['Số tín chỉ', 'STC', 'TinChi', 'Credits'],
          rules: { required: true, type: 'number', integer: true, min: 1, max: 30 } },
        { key: 'LoaiMonHoc', label: 'Loại môn học', aliases: ['Loại môn', 'LoaiMonHoc', 'Hình thức'],
          rules: { allowedValues: ['Bắt buộc', 'Tự chọn'], severity: 'warning' } },
        { key: 'Khoa', label: 'Khoa quản lý', aliases: ['Khoa', 'Bộ môn', 'Đơn vị quản lý'], rules: { trim: true } }
      ],
      uniqueGroups: [], crossFields: []
    },

    {
      id: 'DIEM_ASC', name: 'Điểm — chuẩn ASC', version: '1.0',
      desc: 'Bảng điểm theo học kỳ. Mã sinh viên và mã môn học được đối chiếu với các sheet danh mục tương ứng.',
      fields: [
        { key: 'MaSinhVien', label: 'Mã sinh viên', aliases: ['MSSV', 'Mã SV', 'Mã sinh viên', 'StudentCode'],
          rules: { required: true, trim: true, reference: { sheetHint: 'SinhVien', columnHint: 'MaSinhVien', label: 'Danh sách sinh viên' } } },
        { key: 'MaMonHoc', label: 'Mã môn học', aliases: ['Mã môn', 'Mã học phần', 'MaMonHoc', 'MaHP'],
          rules: { required: true, trim: true, reference: { sheetHint: 'MonHoc', columnHint: 'MaMonHoc', label: 'Danh mục môn học' } } },
        { key: 'HocKy', label: 'Học kỳ', aliases: ['Học kỳ', 'HocKy', 'Kỳ', 'Semester'], rules: { required: true, trim: true } },
        { key: 'DiemSo', label: 'Điểm', aliases: ['Điểm', 'DiemSo', 'Điểm tổng kết', 'Score', 'Mark'],
          rules: { type: 'number', min: 0, max: 10 } },
        { key: 'LanHoc', label: 'Lần học', aliases: ['Lần học', 'LanHoc', 'Lần thi'], rules: { type: 'number', integer: true, min: 1, max: 10, severity: 'warning' } }
      ],
      uniqueGroups: [{ fields: ['MaSinhVien', 'MaMonHoc', 'HocKy'], message: 'Trùng tổ hợp sinh viên + môn học + học kỳ.' }],
      crossFields: []
    },

    {
      id: 'CANBO_ASC', name: 'Cán bộ — chuẩn ASC', version: '1.0',
      desc: 'Danh sách cán bộ, giảng viên, nhân sự đơn vị.',
      fields: [
        { key: 'MaCanBo', label: 'Mã cán bộ', aliases: ['Mã CB', 'Mã cán bộ', 'MaCanBo', 'StaffCode', 'Mã nhân sự'],
          rules: { required: true, unique: true, caseInsensitive: true, trim: true, maxLength: 50 } },
        { key: 'HoTen', label: 'Họ và tên', aliases: ['Họ tên', 'Họ và tên', 'FullName'],
          rules: { required: true, trim: true, noDoubleSpace: true, notAllNumeric: true } },
        { key: 'NgaySinh', label: 'Ngày sinh', aliases: ['Ngày sinh', 'DOB'], rules: { type: 'date', notFuture: true, minDate: '1930-01-01' } },
        { key: 'GioiTinh', label: 'Giới tính', aliases: ['Giới tính', 'Gender', 'Phái'],
          rules: { allowedValues: ['Nam', 'Nữ', 'Khác'], valueMap: GENDER_MAP, severity: 'warning' } },
        { key: 'MaDonVi', label: 'Mã đơn vị', aliases: ['Đơn vị', 'Mã đơn vị', 'Phòng ban', 'Khoa', 'MaDonVi'],
          rules: { trim: true, reference: { sheetHint: 'DonVi', columnHint: 'MaDonVi', label: 'Danh mục đơn vị' } } },
        { key: 'Email', label: 'Email', aliases: ['Email', 'Mail'], rules: { type: 'email', unique: true, uniqueSeverity: 'warning' } },
        { key: 'SoDienThoai', label: 'Số điện thoại', aliases: ['SĐT', 'Điện thoại', 'Phone'], rules: { type: 'phone', severity: 'warning' } }
      ],
      uniqueGroups: [], crossFields: []
    }
  ];

  /* Nhãn tiếng Việt cho từng loại rule — dùng cho chip ở màn ghép cột */
  var RULE_LABELS = {
    required: 'Bắt buộc', unique: 'Không trùng', trim: 'Không thừa khoảng trắng',
    noDoubleSpace: 'Không khoảng trắng kép', notAllNumeric: 'Không toàn số',
    minLength: 'Độ dài tối thiểu', maxLength: 'Độ dài tối đa', regex: 'Định dạng riêng',
    allowedValues: 'Giá trị cho phép', reference: 'Tham chiếu danh mục', type: 'Kiểu dữ liệu',
    min: 'Giá trị nhỏ nhất', max: 'Giá trị lớn nhất', integer: 'Số nguyên',
    notFuture: 'Không ở tương lai', minDate: 'Ngày nhỏ nhất', conditionalRequired: 'Bắt buộc có điều kiện'
  };

  /* Mô tả loại lỗi hiển thị ở cột trái */
  var RULE_TYPE_LABELS = {
    REQUIRED: 'Thiếu dữ liệu bắt buộc',
    DUPLICATE: 'Giá trị trùng',
    DUPLICATE_GROUP: 'Trùng tổ hợp',
    DATA_TYPE: 'Sai kiểu dữ liệu',
    INVALID_DATE: 'Ngày không hợp lệ',
    DATE_AMBIGUOUS: 'Ngày không xác định được định dạng',
    DATE_FUTURE: 'Ngày ở tương lai',
    DATE_RANGE: 'Ngày ngoài khoảng cho phép',
    LENGTH: 'Sai độ dài',
    RANGE: 'Giá trị ngoài khoảng',
    FORMAT: 'Sai định dạng',
    ALLOWED_VALUES: 'Không thuộc danh sách cho phép',
    REFERENCE_EXISTS: 'Không có trong danh mục',
    ADVANCED_REFERENCE: 'Không khớp tham chiếu nâng cao',
    CROSS_FIELD: 'Sai logic giữa các cột',
    CONDITIONAL: 'Thiếu dữ liệu theo điều kiện',
    WHITESPACE: 'Khoảng trắng thừa',
    DOUBLE_SPACE: 'Khoảng trắng kép',
    BLANK: 'Ô trống',
    MIXED_TYPE: 'Kiểu dữ liệu lẫn lộn',
    LEADING_ZERO: 'Mã có thể mất số 0 đầu',
    SCIENTIFIC: 'Số dài hiển thị dạng khoa học',
    DUPLICATE_HEADER: 'Trùng tên cột',
    EMPTY_COLUMN: 'Cột trống hoàn toàn',
    INVALID_HEADER: 'Tên cột không hợp lệ',
    OUTLIER: 'Giá trị bất thường'
  };

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  root.ASC = root.ASC || {};
  root.ASC.rules = {
    presets: RULE_SETS,
    labels: RULE_LABELS,
    typeLabels: RULE_TYPE_LABELS,
    byId: function (id) {
      var all = (root.ASC.rules.custom || []).concat(RULE_SETS);
      for (var i = 0; i < all.length; i++) if (all[i].id === id) return all[i];
      return null;
    },
    clone: clone,
    custom: []
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = root.ASC.rules;
})(typeof window !== 'undefined' ? window : globalThis);
