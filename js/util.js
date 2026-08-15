/* ASC-DataCheck — tiện ích dùng chung (không phụ thuộc DOM ở phần lõi) */
(function (root) {
  'use strict';
  var U = {};

  /* ---------- DOM ---------- */
  U.$ = function (sel, ctx) { return (ctx || document).querySelector(sel); };
  U.$$ = function (sel, ctx) { return Array.prototype.slice.call((ctx || document).querySelectorAll(sel)); };
  U.h = function (tag, attrs, children) {
    var e = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      if (k === 'class') e.className = attrs[k];
      else if (k === 'text') e.textContent = attrs[k];
      else if (k === 'html') e.innerHTML = attrs[k];
      else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2).toLowerCase(), attrs[k]);
      else if (attrs[k] !== null && attrs[k] !== undefined) e.setAttribute(k, attrs[k]);
    });
    (children || []).forEach(function (c) { if (c) e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
    return e;
  };
  U.esc = function (s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  };

  /* ---------- Chuỗi ---------- */
  U.stripDiacritics = function (s) {
    return String(s)
      .replace(/đ/g, 'd').replace(/Đ/g, 'D')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  };
  // Khóa so khớp: bỏ dấu, bỏ ký tự không phải chữ/số, lowercase
  U.normKey = function (s) {
    return U.stripDiacritics(String(s === null || s === undefined ? '' : s))
      .toLowerCase().replace(/[^a-z0-9]/g, '');
  };
  // Chuẩn hóa hiển thị header: gộp khoảng trắng, trim
  U.normHeader = function (s) {
    return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ' ').trim();
  };
  U.hasEdgeSpace = function (s) { return typeof s === 'string' && s !== s.trim() && s.trim() !== ''; };
  U.hasDoubleSpace = function (s) { return typeof s === 'string' && /\S\s{2,}\S/.test(s); };

  /* ---------- Số & định dạng ---------- */
  U.fmtInt = function (n) { return (n === null || n === undefined || isNaN(n)) ? '—' : Number(n).toLocaleString('vi-VN'); };
  U.fmtBytes = function (b) {
    if (!b && b !== 0) return '—';
    if (b < 1024) return b + ' B';
    if (b < 1048576) return (b / 1024).toFixed(1) + ' KB';
    return (b / 1048576).toFixed(1) + ' MB';
  };
  U.pad2 = function (n) { return (n < 10 ? '0' : '') + n; };
  U.fmtDate = function (d) {
    if (!(d instanceof Date) || isNaN(d)) return '';
    return U.pad2(d.getDate()) + '/' + U.pad2(d.getMonth() + 1) + '/' + d.getFullYear();
  };
  U.fmtDateTime = function (d) {
    return U.fmtDate(d) + ' ' + U.pad2(d.getHours()) + ':' + U.pad2(d.getMinutes());
  };

  /* ---------- Giá trị rỗng ---------- */
  U.DEFAULT_EMPTY_TOKENS = ['', '-', '--', 'n/a', 'na', 'null', 'nil', 'none', '#n/a', '.'];
  U.isBlank = function (v, tokens) {
    if (v === null || v === undefined) return true;
    if (v instanceof Date) return false;
    if (typeof v === 'number') return false;
    var s = String(v).trim().toLowerCase();
    if (s === '') return true;
    return (tokens || U.DEFAULT_EMPTY_TOKENS).indexOf(s) >= 0;
  };
  U.toStr = function (v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return U.fmtDate(v);
    return String(v);
  };

  /* ---------- Ngày tháng ---------- */
  var DAYS_IN_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  U.isLeap = function (y) { return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0; };
  U.isRealDate = function (d, m, y) {
    if (!(y >= 1000 && y <= 9999) || m < 1 || m > 12 || d < 1) return false;
    var max = DAYS_IN_MONTH[m - 1] + ((m === 2 && U.isLeap(y)) ? 1 : 0);
    return d <= max;
  };
  // Excel serial -> Date (hệ 1900, đã tính lỗi 29/02/1900 của Excel)
  U.serialToDate = function (n) {
    if (typeof n !== 'number' || !isFinite(n)) return null;
    if (n < 1 || n > 2958465) return null;
    var ms = Math.round((n - 25569) * 86400 * 1000);
    var d = new Date(ms);
    return isNaN(d) ? null : new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  };

  /**
   * Đọc giá trị ngày.
   * @param v  giá trị ô
   * @param opt {order:'DMY'|'MDY'|'auto', allowSerial:bool}
   * @returns {ok, date, ambiguous, reason}
   */
  U.parseDate = function (v, opt) {
    opt = opt || {};
    var order = opt.order || 'auto';
    if (v instanceof Date) {
      return isNaN(v) ? { ok: false, reason: 'invalid' } : { ok: true, date: v, ambiguous: false };
    }
    if (typeof v === 'number') {
      if (opt.allowSerial === false) return { ok: false, reason: 'number' };
      var ds = U.serialToDate(v);
      return ds ? { ok: true, date: ds, ambiguous: false, fromSerial: true } : { ok: false, reason: 'range' };
    }
    var s = U.toStr(v).trim();
    if (s === '') return { ok: false, reason: 'empty' };
    // Chỉ lấy phần ngày nếu có kèm giờ
    s = s.split(/[ T]/)[0];
    var m = s.match(/^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})$/);
    if (!m) {
      // chuỗi toàn số dạng serial (ví dụ "45200")
      if (/^\d{4,5}$/.test(s) && opt.allowSerial !== false) {
        var d2 = U.serialToDate(Number(s));
        if (d2) return { ok: true, date: d2, ambiguous: false, fromSerial: true };
      }
      return { ok: false, reason: 'format' };
    }
    var a = Number(m[1]), b = Number(m[2]), c = Number(m[3]);
    var day, mon, year, ambiguous = false;
    if (m[1].length === 4) { year = a; mon = b; day = c; }          // yyyy-MM-dd
    else if (m[3].length === 4 || c > 31) {
      year = c;
      if (order === 'DMY') { day = a; mon = b; }
      else if (order === 'MDY') { mon = a; day = b; }
      else if (a > 12) { day = a; mon = b; }
      else if (b > 12) { mon = a; day = b; }
      else { day = a; mon = b; ambiguous = true; }                   // 01/02/2026 — không xác định được
    } else {
      year = c < 100 ? (c > 50 ? 1900 + c : 2000 + c) : c;
      if (a > 12) { day = a; mon = b; } else { day = a; mon = b; ambiguous = b <= 12; }
    }
    if (!U.isRealDate(day, mon, year)) return { ok: false, reason: 'invalid', day: day, mon: mon, year: year };
    return { ok: true, date: new Date(year, mon - 1, day), ambiguous: ambiguous };
  };

  /* ---------- Kiểu dữ liệu ---------- */
  U.cellType = function (v) {
    if (U.isBlank(v)) return 'empty';
    if (v instanceof Date) return 'date';
    if (typeof v === 'number') return 'number';
    if (typeof v === 'boolean') return 'boolean';
    var s = String(v).trim();
    if (/^[+-]?\d{1,3}(\.\d{3})*(,\d+)?$/.test(s) && s.indexOf('.') > 0) return 'number';
    if (/^[+-]?(\d+([.,]\d+)?)$/.test(s)) return 'number';
    if (/^\d{1,4}[\/\-.]\d{1,2}[\/\-.]\d{1,4}$/.test(s)) return 'date';
    return 'text';
  };
  U.toNumber = function (v) {
    if (typeof v === 'number') return v;
    var s = U.toStr(v).trim().replace(/\s/g, '');
    if (s === '') return NaN;
    if (/^[+-]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.'); // 1.234,5
    else s = s.replace(',', '.');
    var n = Number(s);
    return isNaN(n) ? NaN : n;
  };

  U.RE = {
    email: /^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/,
    phoneVN: /^(?:\+?84|0)\d{8,11}$/,
    code: /^[A-Za-z0-9._-]+$/
  };

  /* ---------- Trình duyệt ---------- */
  U.download = function (blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 800);
  };
  U.copyText = function (text) {
    if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text);
    return new Promise(function (res, rej) {
      var ta = document.createElement('textarea');
      ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy'); res(); } catch (e) { rej(e); } finally { ta.remove(); }
    });
  };
  U.toast = function (msg, kind) {
    var wrap = U.$('#toastWrap'); if (!wrap) return;
    var t = U.h('div', { class: 'toast ' + (kind || 'info'), text: msg });
    wrap.appendChild(t);
    setTimeout(function () { t.style.opacity = '0'; t.style.transition = 'opacity .25s'; setTimeout(function () { t.remove(); }, 260); }, 2800);
  };
  // Một số trình duyệt chặn localStorage khi mở tệp bằng file:// — khi đó dùng bộ nhớ phiên.
  var mem = {};
  U.store = {
    get: function (k, dflt) {
      try {
        var v = localStorage.getItem('ascdc.' + k);
        if (v !== null) return JSON.parse(v);
      } catch (e) { /* dùng bộ nhớ phiên */ }
      return Object.prototype.hasOwnProperty.call(mem, k) ? mem[k] : dflt;
    },
    set: function (k, v) {
      mem[k] = v;
      try { localStorage.setItem('ascdc.' + k, JSON.stringify(v)); } catch (e) { /* hết dung lượng hoặc bị chặn */ }
    },
    del: function (k) { delete mem[k]; try { localStorage.removeItem('ascdc.' + k); } catch (e) { } }
  };
  U.nextFrame = function () {
    return new Promise(function (r) { (root.requestAnimationFrame || setTimeout)(function () { r(); }, 0); });
  };

  root.ASC = root.ASC || {};
  root.ASC.util = U;
  if (typeof module !== 'undefined' && module.exports) module.exports = U;
})(typeof window !== 'undefined' ? window : globalThis);
