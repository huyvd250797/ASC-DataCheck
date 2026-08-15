/* ASC-DataCheck — Data Grid ảo hóa.
   Chỉ render những dòng đang nhìn thấy để cuộn mượt với hàng chục nghìn dòng. */
(function (root) {
  'use strict';
  var U = root.ASC.util;
  var ROW_H = 30, GUTTER = 62, BUFFER = 8;

  function Grid(host, opt) {
    this.host = host;
    this.opt = opt || {};
    this.rowH = this.opt.rowHeight || ROW_H;
    this.headers = [];
    this.rows = [];
    this.rowNo = [];
    this.order = [];       // chỉ số dòng đang hiển thị (sau lọc + sắp xếp)
    this.base = [];        // chỉ số dòng sau lọc, trước sắp xếp
    this.widths = [];
    this.issueIndex = { cell: new Map(), row: new Map() };
    this.sortCol = -1; this.sortDir = 0;
    this.sel = null;
    this._build();
  }

  Grid.prototype._build = function () {
    var self = this;
    this.host.innerHTML = '';
    this.head = U.h('div', { class: 'grid-head' });
    this.headInner = U.h('div', { class: 'grid-head-inner' });
    this.head.appendChild(this.headInner);
    // Ô "#" của cột số dòng nằm đè lên, không trôi khi cuộn ngang
    this.head.appendChild(U.h('div', { class: 'gh-fixed', text: '#' }));
    this.body = U.h('div', { class: 'grid-body', tabindex: '0' });
    this.canvas = U.h('div', { class: 'grid-canvas' });
    this.body.appendChild(this.canvas);
    this.host.appendChild(this.head);
    this.host.appendChild(this.body);

    this._ticking = false;
    this.body.addEventListener('scroll', function () {
      self.headInner.style.transform = 'translateX(' + -self.body.scrollLeft + 'px)';
      if (!self._ticking) {
        self._ticking = true;
        requestAnimationFrame(function () { self._ticking = false; self._render(); });
      }
    });

    this.body.addEventListener('click', function (e) {
      var cell = e.target.closest ? e.target.closest('.gc') : null;
      if (!cell || cell.classList.contains('gc-gutter')) return;
      var r = Number(cell.getAttribute('data-r')), c = Number(cell.getAttribute('data-c'));
      self.select(r, c);
      if (self.opt.onCellClick) self.opt.onCellClick(r, c, self.issueIndex.cell.get(r + '|' + c));
    });

    this.body.addEventListener('keydown', function (e) {
      if (!(e.ctrlKey || e.metaKey)) return;
      if (e.key === 'c' || e.key === 'C') {
        if (!self.sel) return;
        var text;
        if (e.shiftKey) {
          text = self.headers.map(function (h, i) { return U.toStr(self.rows[self.sel.r][i]); }).join('\t');
        } else {
          text = U.toStr(self.rows[self.sel.r][self.sel.c]);
        }
        U.copyText(text).then(function () { U.toast(e.shiftKey ? 'Đã sao chép dòng' : 'Đã sao chép ô', 'ok'); });
        e.preventDefault();
      }
    });

    window.addEventListener('resize', function () { self._render(); });
  };

  Grid.prototype.setData = function (ds) {
    this.headers = ds.headers || [];
    this.rows = ds.rows || [];
    this.rowNo = ds.rowNo || [];
    this.sortCol = -1; this.sortDir = 0; this.sel = null;
    this._measure();
    this.base = this.rows.map(function (_, i) { return i; });
    this.order = this.base.slice();
    this._renderHead();
    this._render(true);
  };

  Grid.prototype.setIssueIndex = function (idx) { this.issueIndex = idx || { cell: new Map(), row: new Map() }; this._render(true); };

  Grid.prototype.setVisibleRows = function (arr) {
    this.base = arr;
    this._applySort();
    this.body.scrollTop = 0;
    this._render(true);
  };

  Grid.prototype._measure = function () {
    var self = this;
    var sample = Math.min(this.rows.length, 200);
    this.widths = this.headers.map(function (h, c) {
      var max = String(h.name || '').length + 4;
      for (var i = 0; i < sample; i++) {
        var v = self.rows[i][c];
        var L = U.toStr(v).length;
        if (L > max) max = L;
      }
      return Math.max(84, Math.min(300, Math.round(max * 7.2) + 22));
    });
  };

  Grid.prototype.totalWidth = function () {
    return this.widths.reduce(function (a, b) { return a + b; }, GUTTER);
  };

  Grid.prototype._renderHead = function () {
    var self = this;
    this.headInner.innerHTML = '';
    var g = U.h('div', { class: 'gh-cell gutter' });
    g.style.width = GUTTER + 'px';
    this.headInner.appendChild(g);

    this.headers.forEach(function (h, c) {
      var cell = U.h('div', { class: 'gh-cell', title: h.name + '  (cột ' + h.letter + ')' });
      cell.style.width = self.widths[c] + 'px';
      cell.style.flexBasis = self.widths[c] + 'px';
      cell.appendChild(U.h('span', { class: 'gh-name', text: h.name }));
      var tag = U.h('span', { class: 'gh-tag', text: self.sortCol === c ? (self.sortDir > 0 ? '↑' : '↓') : h.letter });
      cell.appendChild(tag);
      cell.addEventListener('click', function (e) {
        if (e.altKey || e.shiftKey) { self.toggleSort(c); return; }
        if (self.opt.onHeaderClick) self.opt.onHeaderClick(c);
      });
      cell.addEventListener('dblclick', function () { self.toggleSort(c); });
      var rz = U.h('div', { class: 'gh-resize' });
      rz.addEventListener('mousedown', function (ev) { self._startResize(ev, c); });
      rz.addEventListener('click', function (ev) { ev.stopPropagation(); });
      cell.appendChild(rz);
      self.headInner.appendChild(cell);
    });
    this.headInner.style.width = this.totalWidth() + 'px';
  };

  Grid.prototype._startResize = function (ev, c) {
    var self = this, x0 = ev.clientX, w0 = this.widths[c];
    ev.preventDefault(); ev.stopPropagation();
    function move(e) {
      self.widths[c] = Math.max(60, w0 + (e.clientX - x0));
      self._renderHead(); self._render(true);
    }
    function up() { document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up); }
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  };

  Grid.prototype.toggleSort = function (c) {
    if (this.sortCol === c) this.sortDir = this.sortDir === 1 ? -1 : (this.sortDir === -1 ? 0 : 1);
    else { this.sortCol = c; this.sortDir = 1; }
    if (this.sortDir === 0) this.sortCol = -1;
    this._applySort();
    this._renderHead();
    this._render(true);
  };

  Grid.prototype._applySort = function () {
    var self = this;
    if (this.sortCol < 0) { this.order = this.base.slice(); return; }
    var c = this.sortCol, dir = this.sortDir;
    this.order = this.base.slice().sort(function (ra, rb) {
      var a = self.rows[ra][c], b = self.rows[rb][c];
      var ea = U.isBlank(a), eb = U.isBlank(b);
      if (ea && eb) return ra - rb;
      if (ea) return 1;
      if (eb) return -1;
      var na = U.toNumber(a), nb = U.toNumber(b);
      if (!isNaN(na) && !isNaN(nb)) return (na - nb) * dir;
      return U.toStr(a).localeCompare(U.toStr(b), 'vi') * dir;
    });
  };

  Grid.prototype.select = function (r, c) {
    this.sel = { r: r, c: c };
    this._render(true);
  };

  Grid.prototype.scrollToRow = function (rowIdx) {
    var pos = this.order.indexOf(rowIdx);
    if (pos < 0) return false;
    this.body.scrollTop = Math.max(0, pos * this.rowH - this.body.clientHeight / 2);
    this._render(true);
    return true;
  };

  Grid.prototype._render = function (force) {
    var self = this;
    var total = this.order.length;
    this.canvas.style.height = (total * this.rowH) + 'px';
    this.canvas.style.width = this.totalWidth() + 'px';

    if (total === 0) {
      this.canvas.innerHTML = '';
      if (!this._emptyEl) this._emptyEl = U.h('div', { class: 'grid-empty', text: 'Không có dòng nào khớp bộ lọc hiện tại.' });
      this.canvas.appendChild(this._emptyEl);
      return;
    }

    var st = this.body.scrollTop, h = this.body.clientHeight || 400;
    var start = Math.max(0, Math.floor(st / this.rowH) - BUFFER);
    var end = Math.min(total, Math.ceil((st + h) / this.rowH) + BUFFER);
    if (!force && this._start === start && this._end === end) return;
    this._start = start; this._end = end;

    var frag = document.createDocumentFragment();
    for (var p = start; p < end; p++) {
      var r = this.order[p];
      frag.appendChild(this._buildRow(r, p));
    }
    this.canvas.innerHTML = '';
    this.canvas.appendChild(frag);
  };

  Grid.prototype._buildRow = function (r, pos) {
    var self = this;
    var row = U.h('div', { class: 'grid-row' });
    row.style.transform = 'translateY(' + (pos * this.rowH) + 'px)';
    row.style.width = this.totalWidth() + 'px';

    var rstat = this.issueIndex.row.get(r);
    var gut = U.h('div', { class: 'gc gc-gutter' });
    gut.style.width = GUTTER + 'px';
    gut.style.flexBasis = GUTTER + 'px';
    if (rstat && rstat.error) gut.appendChild(U.h('span', { class: 'row-flag e', text: '!', title: rstat.error + ' lỗi' }));
    else if (rstat && rstat.warning) gut.appendChild(U.h('span', { class: 'row-flag w', text: '!', title: rstat.warning + ' cảnh báo' }));
    gut.appendChild(U.h('span', { text: String(this.rowNo[r] || (r + 1)) }));
    if (rstat) gut.title = (rstat.error ? rstat.error + ' lỗi' : '') + (rstat.error && rstat.warning ? ', ' : '') + (rstat.warning ? rstat.warning + ' cảnh báo' : '');
    row.appendChild(gut);

    for (var c = 0; c < this.headers.length; c++) {
      var v = this.rows[r][c];
      var cls = 'gc';
      var hit = this.issueIndex.cell.get(r + '|' + c);
      if (hit) cls += hit.severity === 'error' ? ' err' : (hit.severity === 'warning' ? ' warn' : '');
      if (this.sel && this.sel.r === r && this.sel.c === c) cls += ' sel';
      var text = U.toStr(v);
      var cell = U.h('div', { class: cls, 'data-r': r, 'data-c': c, text: text });
      if (hit) cell.title = hit.list.map(function (i) { return i.message; }).join('\n');
      else if (text.length > 18) cell.title = text;
      cell.style.width = this.widths[c] + 'px';
      cell.style.flexBasis = this.widths[c] + 'px';
      row.appendChild(cell);
    }
    return row;
  };

  root.ASC = root.ASC || {};
  root.ASC.Grid = Grid;
})(typeof window !== 'undefined' ? window : globalThis);
