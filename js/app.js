/* ASC-DataCheck — điều phối giao diện */
(function (root) {
  'use strict';
  var U = root.ASC.util, E = root.ASC.engine, R = root.ASC.rules, X = root.ASC.exporter;
  var $ = U.$, h = U.h;
  var APP_VERSION = '1.0.0';

  var S = {
    file: null, wb: null, sheets: [], aoaCache: {}, sheet: null, headerRowIdx: 0,
    dataset: null, mode: null, ruleSetId: null, mapping: {}, refs: {}, advancedRefs: [],
    query: { sheet: '', conditions: [], result: null, useAsReference: false },
    issues: [], summary: null, index: null, profiles: null,
    filter: { tab: 'all', rule: null, column: null, search: '', scope: '*' },
    byRuleRows: null, byColumnRows: null, grid: null, previewGrid: null, token: null
  };

  /* ================= Khởi động ================= */
  function init() {
    $('#versionLabel').textContent = 'v' + APP_VERSION;
    R.custom = U.store.get('customRules', []);
    applyTheme(U.store.get('theme', 'dark'));
    bindHeader();
    bindHome();
    bindSheetView();
    bindMappingView();
    bindResultView();
    bindShortcuts();
    $('#progCancel').addEventListener('click', function () { if (S.token) S.token.cancelled = true; });
    renderHistory();
    setTimeout(function () { $('#splash').classList.add('gone'); }, 700);
    setTimeout(function () {
      var el = $('#splash'); if (el) el.remove();
      if (!U.store.get('hasSeenQuickGuide', false)) openHelp();
    }, 1150);
  }

  /* ================= Theme ================= */
  function applyTheme(t) {
    document.documentElement.setAttribute('data-theme', t);
    $('#iconMoon').classList.toggle('hidden', t === 'light');
    $('#iconSun').classList.toggle('hidden', t !== 'light');
    U.store.set('theme', t);
  }

  /* ================= Header & modal ================= */
  function bindHeader() {
    $('#themeBtn').addEventListener('click', function () {
      applyTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
    });
    $('#helpBtn').addEventListener('click', openHelp);
    $('#helpClose').addEventListener('click', closeHelp);
    $('#helpCloseBtn').addEventListener('click', closeHelp);
    $('#brandBtn').addEventListener('click', function () { if (S.file) gotoView('sheet'); else gotoView('home'); });
    U.$$('#steps .step').forEach(function (b) {
      b.addEventListener('click', function () {
        if (!b.classList.contains('done')) return;
        var s = Number(b.getAttribute('data-step'));
        if (s === 1) gotoView('home');
        else if (s <= 3) gotoView('sheet');
        else if (s === 4 && S.mode === 'rule') gotoView('mapping');
        else if (s === 5 && S.summary) gotoView('result');
      });
    });
    // Modal KHÔNG đóng khi click ra ngoài — chỉ đóng bằng X hoặc nút Đóng
    $('#exportClose').addEventListener('click', closeExport);
    $('#exportCloseBtn').addEventListener('click', closeExport);
  }

  function lockScroll(on) { document.body.classList.toggle('modal-open', on); }
  function openHelp() {
    $('#dontShowAgain').checked = U.store.get('hasSeenQuickGuide', true);
    $('#helpModal').classList.remove('hidden'); lockScroll(true);
  }
  function closeHelp() {
    // Bỏ tick nếu muốn hướng dẫn tự mở lại ở lần truy cập sau
    U.store.set('hasSeenQuickGuide', $('#dontShowAgain').checked);
    $('#helpModal').classList.add('hidden'); lockScroll(false);
  }
  function openExport() { $('#exportModal').classList.remove('hidden'); lockScroll(true); renderExportOptions(); }
  function closeExport() { $('#exportModal').classList.add('hidden'); lockScroll(false); }

  function bindShortcuts() {
    document.addEventListener('keydown', function (e) {
      var typing = /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement.tagName);
      if (e.key === '?' && !typing) { openHelp(); e.preventDefault(); }
      if ((e.ctrlKey || e.metaKey) && e.key === 'f' && $('#viewResult').classList.contains('active')) {
        e.preventDefault(); $('#searchInput').focus();
      }
      if ((e.ctrlKey || e.metaKey) && e.key === 'e' && S.summary) { e.preventDefault(); openExport(); }
      if (e.key === 'Escape' && !$('#exportModal').classList.contains('hidden')) closeExport();
    });
  }

  /* ================= Điều hướng ================= */
  function gotoView(name) {
    var map = { home: 'viewHome', sheet: 'viewSheet', mapping: 'viewMapping', result: 'viewResult' };
    U.$$('.view').forEach(function (v) { v.classList.remove('active'); });
    $('#' + map[name]).classList.add('active');
    var step = { home: 1, sheet: 2, mapping: 4, result: 5 }[name];
    U.$$('#steps .step').forEach(function (b) {
      var s = Number(b.getAttribute('data-step'));
      b.classList.toggle('active', s === step || (name === 'sheet' && s === 3 && !!S.sheet));
      b.classList.toggle('done', s < step);
    });
    if (name === 'result' && S.grid) setTimeout(function () { S.grid._render(true); }, 0);
  }

  /* ================= HOME ================= */
  function bindHome() {
    var dz = $('#dropzone'), input = $('#fileInput');
    $('#pickBtn').addEventListener('click', function () { input.click(); });
    dz.addEventListener('click', function (e) { if (e.target === dz) input.click(); });
    input.addEventListener('change', function () { if (input.files[0]) loadFile(input.files[0]); });
    ['dragenter', 'dragover'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('over'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); if (ev === 'dragleave' && dz.contains(e.relatedTarget)) return; dz.classList.remove('over'); });
    });
    dz.addEventListener('drop', function (e) {
      var f = e.dataTransfer.files[0]; if (f) loadFile(f);
    });
    document.addEventListener('dragover', function (e) { e.preventDefault(); });
    document.addEventListener('drop', function (e) { e.preventDefault(); });
    $('#clearHistoryBtn').addEventListener('click', function () {
      U.store.set('history', []); renderHistory(); U.toast('Đã xóa lịch sử kiểm tra', 'ok');
    });
    $('#changeFileBtn').addEventListener('click', function () { gotoView('home'); });
  }

  function renderHistory() {
    var list = U.store.get('history', []);
    var host = $('#recentList');
    host.innerHTML = '';
    if (!list.length) {
      host.appendChild(h('div', { class: 'empty-state', text: 'Chưa có phiên kiểm tra dữ liệu.' }));
      return;
    }
    list.forEach(function (it) {
      host.appendChild(h('div', { class: 'recent-row' }, [
        h('span', { class: 'r-name', text: it.fileName + ' · ' + it.sheet }),
        h('span', { class: 'r-meta', text: U.fmtInt(it.rows) + ' dòng' }),
        h('span', { class: 'r-meta', style: 'color:var(--error)', text: U.fmtInt(it.errors) + ' lỗi' }),
        h('span', { class: 'r-meta', style: 'color:var(--warn)', text: U.fmtInt(it.warnings) + ' cảnh báo' }),
        h('span', { class: 'r-meta', text: it.quality + '%' }),
        h('span', { class: 'r-meta', text: it.at })
      ]));
    });
  }

  function saveHistory() {
    var list = U.store.get('history', []);
    list.unshift({
      fileName: S.file.name, sheet: S.sheet, rows: S.summary.totalRows,
      errors: S.summary.errorCount, warnings: S.summary.warningCount,
      quality: S.summary.qualityScore, at: U.fmtDateTime(new Date()),
      mode: S.mode, ruleSet: S.ruleSetId
    });
    U.store.set('history', list.slice(0, 12));
    renderHistory();
  }

  /* ================= Đọc tệp ================= */
  function loadFile(file) {
    if (!/\.(xlsx|xlsm|xls|csv)$/i.test(file.name)) {
      U.toast('ASC-DataCheck hiện hỗ trợ tệp Excel (.xlsx, .xls, .xlsm) và .csv', 'err');
      return;
    }
    if (file.size > 50 * 1024 * 1024) {
      U.toast('Tệp có dung lượng lớn. Quá trình kiểm tra có thể dùng nhiều bộ nhớ.', 'info');
    }
    showProgress('Đang đọc tệp Excel…', file.name, null);
    var reader = new FileReader();
    reader.onerror = function () { hideProgress(); U.toast('Không đọc được tệp. Kiểm tra lại tệp nguồn.', 'err'); };
    reader.onload = function (e) {
      setTimeout(function () {
        try {
          var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true, cellStyles: false });
          if (!wb.SheetNames.length) throw new Error('EMPTY');
          S.file = { name: file.name, size: file.size };
          S.wb = wb; S.aoaCache = {}; S.sheet = null; S.dataset = null;
          S.summary = null; S.issues = []; S.index = null;
          S.sheets = wb.SheetNames.map(function (n) {
            var ws = wb.Sheets[n];
            var ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
            return { name: n, rows: ref ? ref.e.r - ref.s.r : 0, cols: ref ? ref.e.c - ref.s.c + 1 : 0 };
          });
          hideProgress();
          renderSheetView();
          gotoView('sheet');
          if (S.sheets.length === 1) selectSheet(S.sheets[0].name);
        } catch (err) {
          hideProgress();
          U.toast(err.message === 'EMPTY' ? 'Không tìm thấy dữ liệu để kiểm tra.' : 'Tệp không đọc được hoặc đã hỏng.', 'err');
        }
      }, 30);
    };
    reader.readAsArrayBuffer(file);
  }

  function aoaOf(sheetName) {
    if (S.aoaCache[sheetName]) return S.aoaCache[sheetName];
    var ws = S.wb.Sheets[sheetName];
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    S.aoaCache[sheetName] = aoa;
    return aoa;
  }

  function detectHeaderRow(aoa) {
    var limit = Math.min(aoa.length, 12);
    var best = 0, bestScore = -1;
    for (var i = 0; i < limit; i++) {
      var row = aoa[i] || [], filled = 0, texts = 0;
      for (var c = 0; c < row.length; c++) {
        if (U.isBlank(row[c])) continue;
        filled++;
        if (typeof row[c] === 'string' && !/^\d+$/.test(row[c].trim())) texts++;
      }
      var next = aoa[i + 1] || [], nextFilled = 0;
      for (var c2 = 0; c2 < next.length; c2++) if (!U.isBlank(next[c2])) nextFilled++;
      var score = filled * 2 + texts * 2 + (nextFilled >= filled - 1 ? 4 : 0) - i;
      if (filled < 2) score -= 8;
      if (score > bestScore) { bestScore = score; best = i; }
    }
    return best;
  }

  /* ================= VIEW: SHEET ================= */
  function bindSheetView() {
    $('#headerRowSel').addEventListener('change', function () {
      S.headerRowIdx = Number(this.value);
      buildDataset();
    });
    $('#skipEmptyRows').addEventListener('change', buildDataset);
    $('#modeQuick').addEventListener('click', function () { S.mode = 'quick'; runQuickCheck(); });
    $('#modeRule').addEventListener('click', function () { S.mode = 'rule'; openMapping(); });
  }

  function renderSheetView() {
    $('#fbName').textContent = S.file.name;
    $('#fbMeta').textContent = U.fmtBytes(S.file.size) + ' · ' + S.sheets.length + ' sheet';
    var grid = $('#sheetGrid'); grid.innerHTML = '';
    S.sheets.forEach(function (s) {
      var card = h('button', { class: 'sheet-card', 'data-sheet': s.name }, [
        h('b', { text: s.name }),
        h('span', { text: U.fmtInt(s.rows) + ' dòng · ' + s.cols + ' cột' })
      ]);
      card.addEventListener('click', function () { selectSheet(s.name); });
      grid.appendChild(card);
    });
    $('#previewCard').classList.add('hidden');
    U.$$('.mode-grid .mode-card').forEach(function (b) { b.disabled = true; });
  }

  function selectSheet(name) {
    S.sheet = name;
    U.$$('#sheetGrid .sheet-card').forEach(function (c) {
      c.classList.toggle('sel', c.getAttribute('data-sheet') === name);
    });
    var aoa = aoaOf(name);
    S.headerRowIdx = detectHeaderRow(aoa);
    var sel = $('#headerRowSel'); sel.innerHTML = '';
    for (var i = 0; i < Math.min(aoa.length, 12); i++) {
      var preview = (aoa[i] || []).filter(function (v) { return !U.isBlank(v); }).slice(0, 3).map(U.toStr).join(' | ');
      sel.appendChild(h('option', { value: String(i), text: 'Dòng ' + (i + 1) + (preview ? ' — ' + preview.substring(0, 44) : ' — (trống)') }));
    }
    sel.value = String(S.headerRowIdx);
    $('#previewCard').classList.remove('hidden');
    U.$$('.mode-grid .mode-card').forEach(function (b) { b.disabled = false; });
    buildDataset();
    gotoView('sheet');
  }

  function buildDataset() {
    var aoa = aoaOf(S.sheet);
    S.dataset = E.buildDataset(aoa, S.headerRowIdx, { keepEmptyRows: !$('#skipEmptyRows').checked });
    S.dataset.sheetName = S.sheet;
    if (!S.previewGrid) S.previewGrid = new root.ASC.Grid($('#previewHost'), {});
    S.previewGrid.setData(S.dataset);
    renderStructureNotes();
  }

  function renderStructureNotes() {
    var host = $('#structureNotes'); host.innerHTML = '';
    var ws = S.wb.Sheets[S.sheet];
    var notes = S.dataset.notes.slice();

    notes.push({ kind: 'info', text: U.fmtInt(S.dataset.rows.length) + ' dòng dữ liệu · ' + S.dataset.headers.length + ' cột.' });
    if (ws['!merges'] && ws['!merges'].length) {
      notes.push({ kind: 'warn', text: 'Sheet có ' + ws['!merges'].length + ' vùng ô gộp. Ô gộp có thể làm lệch dữ liệu khi đọc.' });
    }
    var hiddenCols = (ws['!cols'] || []).filter(function (c) { return c && c.hidden; }).length;
    if (hiddenCols) notes.push({ kind: 'warn', text: 'Sheet có ' + hiddenCols + ' cột đang bị ẩn — dữ liệu ẩn vẫn được kiểm tra.' });
    var dupHeaders = S.dataset.headers.filter(function (x) { return x.duplicateHeader; });
    if (dupHeaders.length) {
      notes.push({ kind: 'warn', text: 'Trùng tên cột: ' + dupHeaders.map(function (x) { return x.name; }).join(', ') + '.' });
    }
    notes.forEach(function (n) {
      host.appendChild(h('div', {
        style: 'font-size:12.5px;color:' + (n.kind === 'warn' ? 'var(--warn)' : 'var(--muted)') + ';padding:2px 0',
        text: (n.kind === 'warn' ? '⚠ ' : '· ') + n.text
      }));
    });
  }

  /* ================= VIEW: MAPPING ================= */
  function bindMappingView() {
    $('#ruleSetSel').addEventListener('change', function () { S.ruleSetId = this.value; renderMapping(true); });
    $('#autoMapBtn').addEventListener('click', function () { autoMap(); renderMapping(false); U.toast('Đã ghép tự động theo tên cột và alias', 'ok'); });
    $('#backToSheetBtn').addEventListener('click', function () { gotoView('sheet'); });
    $('#runValidationBtn').addEventListener('click', runRuleValidation);
    $('#cloneRuleBtn').addEventListener('click', function () { openRuleEditor('clone'); });
    $('#editRuleBtn').addEventListener('click', function () { openRuleEditor('edit'); });
    $('#newTempRuleBtn').addEventListener('click', function () { openRuleEditor('new'); });
    $('#deleteRuleBtn').addEventListener('click', deleteCustomRule);
    $('#addAdvRefBtn').addEventListener('click', function () { addAdvancedRef(); renderMapping(false); });
    $('#addQueryCondBtn').addEventListener('click', function () { addQueryCondition(); renderQueryBuilder(); });
    $('#runQueryBtn').addEventListener('click', runQuery);
    $('#useQueryAsRefBtn').addEventListener('click', useQueryAsReference);
    $('#ruleEditorClose').addEventListener('click', closeRuleEditor);
    $('#ruleEditorCancel').addEventListener('click', closeRuleEditor);
    $('#formatRuleBtn').addEventListener('click', formatRuleJson);
    $('#saveRuleBtn').addEventListener('click', saveRuleFromEditor);
    $('#exportRuleBtn').addEventListener('click', function () {
      var rs = R.byId(S.ruleSetId);
      U.download(new Blob([JSON.stringify(rs, null, 2)], { type: 'application/json' }), rs.id + '.json');
      U.toast('Đã xuất bộ rule ra JSON', 'ok');
    });
    $('#importRuleBtn').addEventListener('click', function () { $('#ruleFileInput').click(); });
    $('#ruleFileInput').addEventListener('change', function () {
      var f = this.files[0]; if (!f) return;
      var fr = new FileReader();
      fr.onload = function (e) {
        try {
          var rs = JSON.parse(e.target.result);
          if (!rs.id || !Array.isArray(rs.fields)) throw new Error('bad');
          R.custom = R.custom.filter(function (x) { return x.id !== rs.id; });
          R.custom.unshift(rs);
          saveCustomRules();
          S.ruleSetId = rs.id;
          fillRuleSelect(); renderMapping(true);
          U.toast('Đã nhập bộ rule "' + rs.name + '"', 'ok');
        } catch (err) { U.toast('Tệp rule không hợp lệ.', 'err'); }
      };
      fr.readAsText(f);
      this.value = '';
    });
  }

  function fillRuleSelect() {
    var sel = $('#ruleSetSel'); sel.innerHTML = '';
    R.custom.concat(R.presets).forEach(function (rs) {
      sel.appendChild(h('option', { value: rs.id, text: rs.name + ' (v' + rs.version + ')' }));
    });
    if (!S.ruleSetId) S.ruleSetId = guessRuleSet();
    sel.value = S.ruleSetId;
    $('#deleteRuleBtn').disabled = !isCustomRule(S.ruleSetId);
  }

  function saveCustomRules() {
    U.store.set('customRules', R.custom || []);
  }

  function isCustomRule(id) {
    return (R.custom || []).some(function (x) { return x.id === id; });
  }

  function makeTempRule(rs, suffix) {
    var out = R.clone(rs);
    var base = U.normKey(rs.id || rs.name || 'RULE').toUpperCase() || 'RULE';
    out.id = base + '_' + suffix + '_' + Date.now();
    out.name = (rs.name || 'Bộ rule') + ' — bản tạm';
    out.version = 'temp';
    return out;
  }

  function openRuleEditor(mode) {
    var rs = R.byId(S.ruleSetId) || R.presets[0];
    var editRule = mode === 'new'
      ? makeTempRule({ id: 'RULE_TAM', name: 'Bộ rule tạm', version: 'temp', desc: 'Bộ rule tạm do người dùng tạo.', fields: [], uniqueGroups: [], crossFields: [] }, 'NEW')
      : (mode === 'clone' || !isCustomRule(rs.id) ? makeTempRule(rs, mode === 'clone' ? 'COPY' : 'EDIT') : R.clone(rs));
    $('#ruleEditorTitle').textContent = mode === 'new' ? 'Tạo rule tạm' : (mode === 'clone' ? 'Nhân bản rule' : 'Sửa bộ rule');
    $('#ruleEditorHint').textContent = isCustomRule(rs.id) && mode === 'edit'
      ? 'Đang sửa rule tạm đã lưu trong trình duyệt. Anh có thể xóa bằng nút "Xóa rule tạm".'
      : 'Rule gốc ASC không bị ghi đè. Khi lưu, hệ thống tạo một rule tạm và lưu trong trình duyệt.';
    $('#ruleJsonInput').value = JSON.stringify(editRule, null, 2);
    $('#ruleEditorModal').classList.remove('hidden'); lockScroll(true);
  }

  function closeRuleEditor() {
    $('#ruleEditorModal').classList.add('hidden'); lockScroll(false);
  }

  function formatRuleJson() {
    try {
      $('#ruleJsonInput').value = JSON.stringify(JSON.parse($('#ruleJsonInput').value), null, 2);
    } catch (e) { U.toast('JSON chưa hợp lệ, chưa format được.', 'err'); }
  }

  function saveRuleFromEditor() {
    try {
      var rs = JSON.parse($('#ruleJsonInput').value);
      if (!rs.id || !rs.name || !Array.isArray(rs.fields)) throw new Error('bad');
      rs.uniqueGroups = rs.uniqueGroups || [];
      rs.crossFields = rs.crossFields || [];
      R.custom = (R.custom || []).filter(function (x) { return x.id !== rs.id; });
      R.custom.unshift(rs);
      saveCustomRules();
      S.ruleSetId = rs.id;
      fillRuleSelect();
      renderMapping(true);
      closeRuleEditor();
      U.toast('Đã lưu bộ rule "' + rs.name + '" trong trình duyệt', 'ok');
    } catch (e) {
      U.toast('Rule JSON chưa hợp lệ. Cần có id, name và fields[].', 'err');
    }
  }

  function deleteCustomRule() {
    if (!isCustomRule(S.ruleSetId)) { U.toast('Chỉ xóa được rule tạm/custom.', 'info'); return; }
    var rs = R.byId(S.ruleSetId);
    if (!confirm('Xóa rule tạm "' + rs.name + '" khỏi trình duyệt?')) return;
    R.custom = R.custom.filter(function (x) { return x.id !== S.ruleSetId; });
    saveCustomRules();
    S.ruleSetId = null;
    fillRuleSelect();
    renderMapping(true);
    U.toast('Đã xóa rule tạm.', 'ok');
  }

  function guessRuleSet() {
    var k = U.normKey(S.sheet || '');
    var hit = R.presets.find(function (rs) { return k.indexOf(U.normKey(rs.id.split('_')[0])) >= 0; });
    return hit ? hit.id : R.presets[0].id;
  }

  function openMapping() {
    fillRuleSelect();
    renderMapping(true);
    gotoView('mapping');
  }

  function scoreMatch(field, header) {
    var hk = header.key;
    var names = [field.key, field.label].concat(field.aliases || []);
    for (var i = 0; i < names.length; i++) {
      var nk = U.normKey(names[i]);
      if (!nk) continue;
      if (nk === hk) return 100;
    }
    for (var j = 0; j < names.length; j++) {
      var nk2 = U.normKey(names[j]);
      if (!nk2 || nk2.length < 3) continue;
      if (hk.indexOf(nk2) >= 0 || nk2.indexOf(hk) >= 0) return 70 - Math.abs(hk.length - nk2.length);
    }
    return 0;
  }

  function autoMap() {
    var rs = R.byId(S.ruleSetId); if (!rs) return;
    var used = {};
    S.mapping = {};
    var pairs = [];
    rs.fields.forEach(function (f) {
      S.dataset.headers.forEach(function (hd) {
        var sc = scoreMatch(f, hd);
        if (sc > 0) pairs.push({ f: f.key, c: hd.index, s: sc });
      });
    });
    pairs.sort(function (a, b) { return b.s - a.s; });
    pairs.forEach(function (p) {
      if (S.mapping[p.f] !== undefined || used[p.c]) return;
      S.mapping[p.f] = p.c; used[p.c] = 1;
    });
    autoRefs();
  }

  function autoRefs() {
    var rs = R.byId(S.ruleSetId); if (!rs) return;
    S.refs = {};
    rs.fields.forEach(function (f) {
      if (!f.rules.reference) return;
      var hint = U.normKey(f.rules.reference.sheetHint || '');
      var sheet = S.sheets.map(function (s) { return s.name; }).find(function (n) {
        return n !== S.sheet && (U.normKey(n).indexOf(hint) >= 0 || hint.indexOf(U.normKey(n)) >= 0);
      });
      if (!sheet) { S.refs[f.key] = null; return; }
      var col = findRefColumn(sheet, f.rules.reference.columnHint, f);
      S.refs[f.key] = col === null ? null : { sheet: sheet, column: col.name, colIdx: col.index };
    });
  }

  function refDatasetOf(sheetName) {
    var aoa = aoaOf(sheetName);
    return E.buildDataset(aoa, detectHeaderRow(aoa), {});
  }

  function findRefColumn(sheetName, columnHint, field) {
    var ds = refDatasetOf(sheetName);
    var want = U.normKey(columnHint || '');
    var best = null, bestScore = 0;
    ds.headers.forEach(function (hd) {
      var sc = 0;
      if (hd.key === want) sc = 100;
      else if (want && (hd.key.indexOf(want) >= 0 || want.indexOf(hd.key) >= 0)) sc = 70;
      else if (field) sc = scoreMatch(field, hd);
      if (sc > bestScore) { bestScore = sc; best = hd; }
    });
    return bestScore > 0 ? best : null;
  }

  function refDatasetByName(name) {
    if (name === '__QUERY__') return S.query.result;
    return refDatasetOf(name);
  }

  function refSourceNames() {
    var arr = S.sheets.map(function (s) { return { value: s.name, text: s.name }; });
    if (S.query.useAsReference && S.query.result) arr.unshift({ value: '__QUERY__', text: 'Kết quả query tạm' });
    return arr;
  }

  function fillSelector(sel, value, rs) {
    sel.innerHTML = '';
    var fHead = h('option', { value: '', text: '— Chọn Field ASC hoặc cột Excel —' });
    sel.appendChild(fHead);
    (rs.fields || []).forEach(function (f) {
      sel.appendChild(h('option', { value: 'field:' + f.key, text: 'Field ASC: ' + f.label + ' (' + f.key + ')' }));
    });
    S.dataset.headers.forEach(function (hd) {
      sel.appendChild(h('option', { value: 'col:' + hd.index, text: 'Cột Excel: ' + hd.name + ' (' + hd.letter + ')' }));
    });
    sel.value = value || '';
  }

  function fillRefColumns(sel, sourceName, selected) {
    sel.innerHTML = '';
    var ds = sourceName ? refDatasetByName(sourceName) : null;
    if (!ds || !ds.headers.length) { sel.disabled = true; return; }
    sel.disabled = false;
    ds.headers.forEach(function (hd) {
      sel.appendChild(h('option', { value: String(hd.index), text: hd.name + (hd.letter ? ' (' + hd.letter + ')' : '') }));
    });
    if (selected !== undefined && selected !== null) sel.value = String(selected);
  }

  function selectorToCol(sel) {
    if (!sel) return -1;
    if (sel.indexOf('field:') === 0) return S.mapping[sel.slice(6)] === undefined ? -1 : S.mapping[sel.slice(6)];
    if (sel.indexOf('col:') === 0) return Number(sel.slice(4));
    return -1;
  }

  function selectorLabel(sel, rs) {
    if (!sel) return 'Chưa chọn';
    if (sel.indexOf('field:') === 0) {
      var k = sel.slice(6), f = (rs.fields || []).find(function (x) { return x.key === k; });
      return f ? f.label : k;
    }
    var c = selectorToCol(sel);
    return S.dataset.headers[c] ? S.dataset.headers[c].name : 'Cột ' + (c + 1);
  }

  function addAdvancedRef() {
    var rs = R.byId(S.ruleSetId);
    var firstField = (rs.fields || []).find(function (f) { return S.mapping[f.key] !== undefined; });
    var target = firstField ? 'field:' + firstField.key : (S.dataset.headers[0] ? 'col:0' : '');
    var firstSource = refSourceNames()[0];
    var firstSheet = firstSource ? firstSource.value : '';
    var refCol = 0;
    if (firstSheet && firstField && firstSheet !== '__QUERY__') {
      var hit = findRefColumn(firstSheet, firstField.key, firstField);
      if (hit) refCol = hit.index;
    }
    S.advancedRefs.push({
      id: 'adv_' + Date.now() + '_' + Math.random().toString(16).slice(2),
      enabled: true, target: target, sheet: firstSheet, logic: 'AND', severity: 'error',
      conditions: [{ source: target, op: 'eq', refColIdx: refCol }]
    });
  }

  function renderAdvancedRefs(rs) {
    var host = $('#advRefBody'); host.innerHTML = '';
    if (!S.advancedRefs.length) {
      host.appendChild(h('div', { class: 'empty-state', text: 'Chưa có tham chiếu nâng cao. Bấm "Thêm tham chiếu" để tự đối chiếu Field ASC hoặc bất kỳ cột Excel nào.' }));
      return;
    }
    S.advancedRefs.forEach(function (ar, idx) {
      var item = h('div', { class: 'builder-item' });
      var title = h('div', { class: 'builder-title' }, [
        h('b', { text: 'Tham chiếu ' + (idx + 1) }),
        h('small', { text: selectorLabel(ar.target, rs) }),
        h('div', { class: 'spacer' })
      ]);
      var enabled = h('label', { class: 'checkline' }, [h('input', { type: 'checkbox' }), document.createTextNode(' Bật')]);
      enabled.querySelector('input').checked = ar.enabled !== false;
      enabled.querySelector('input').addEventListener('change', function () { ar.enabled = this.checked; });
      var del = h('button', { class: 'btn ghost sm', text: 'Xóa' });
      del.addEventListener('click', function () { S.advancedRefs.splice(idx, 1); renderMapping(false); });
      title.appendChild(enabled); title.appendChild(del); item.appendChild(title);

      var top = h('div', { class: 'field-row compact' });
      var targetSel = h('select', { class: 'inp grow' }); fillSelector(targetSel, ar.target, rs);
      targetSel.addEventListener('change', function () { ar.target = this.value; if (ar.conditions[0]) ar.conditions[0].source = this.value; renderMapping(false); });
      var sheetSel = h('select', { class: 'inp' });
      refSourceNames().forEach(function (s) { sheetSel.appendChild(h('option', { value: s.value, text: s.text })); });
      sheetSel.value = ar.sheet || '';
      sheetSel.addEventListener('change', function () {
        ar.sheet = this.value;
        ar.conditions.forEach(function (c) { c.refColIdx = 0; });
        renderMapping(false);
      });
      var logicSel = h('select', { class: 'inp' }, [h('option', { value: 'AND', text: 'AND' }), h('option', { value: 'OR', text: 'OR' })]);
      logicSel.value = ar.logic || 'AND';
      logicSel.addEventListener('change', function () { ar.logic = this.value; renderAdvancedRefs(rs); });
      var sevSel = h('select', { class: 'inp' }, [
        h('option', { value: 'error', text: 'Lỗi' }), h('option', { value: 'warning', text: 'Cảnh báo' }), h('option', { value: 'info', text: 'Gợi ý' })
      ]);
      sevSel.value = ar.severity || 'error';
      sevSel.addEventListener('change', function () { ar.severity = this.value; });
      top.appendChild(h('label', { class: 'lbl', text: 'Báo lỗi tại' })); top.appendChild(targetSel);
      top.appendChild(h('label', { class: 'lbl', text: 'Danh mục' })); top.appendChild(sheetSel);
      top.appendChild(h('label', { class: 'lbl', text: 'Logic' })); top.appendChild(logicSel);
      top.appendChild(sevSel);
      item.appendChild(top);

      var condHost = h('div', {});
      (ar.conditions || []).forEach(function (c, ci) {
        var row = h('div', { class: 'cond-row' });
        row.appendChild(h('div', { class: 'logic-badge', text: ci === 0 ? 'WHERE' : (ar.logic || 'AND') }));
        var srcSel = h('select', { class: 'inp' }); fillSelector(srcSel, c.source, rs);
        srcSel.addEventListener('change', function () { c.source = this.value; });
        var opSel = h('select', { class: 'inp' }, [
          h('option', { value: 'eq', text: '=' }),
          h('option', { value: 'contains', text: 'LIKE' }),
          h('option', { value: 'refContains', text: 'REF LIKE' }),
          h('option', { value: 'starts', text: 'STARTS' }),
          h('option', { value: 'notEq', text: '<>' })
        ]);
        opSel.value = c.op || 'eq';
        opSel.addEventListener('change', function () { c.op = this.value; });
        var refColSel = h('select', { class: 'inp' }); fillRefColumns(refColSel, ar.sheet, c.refColIdx);
        refColSel.addEventListener('change', function () { c.refColIdx = Number(this.value); });
        var delCond = h('button', { class: 'btn ghost sm', text: 'Xóa' });
        delCond.disabled = ar.conditions.length <= 1;
        delCond.addEventListener('click', function () { ar.conditions.splice(ci, 1); renderMapping(false); });
        row.appendChild(srcSel); row.appendChild(opSel); row.appendChild(refColSel); row.appendChild(delCond);
        condHost.appendChild(row);
      });
      item.appendChild(condHost);
      var addCond = h('button', { class: 'btn ghost sm', style: 'margin-top:10px', text: 'Thêm điều kiện tham chiếu' });
      addCond.addEventListener('click', function () {
        ar.conditions.push({ source: ar.target, op: 'eq', refColIdx: 0 });
        renderMapping(false);
      });
      item.appendChild(addCond);
      host.appendChild(item);
    });
  }

  function addQueryCondition() {
    S.query.conditions.push({ logic: S.query.conditions.length ? 'OR' : 'WHERE', colIdx: 0, op: 'contains', value: '' });
  }

  function renderQueryBuilder() {
    var sheetSel = $('#querySheetSel');
    sheetSel.innerHTML = '';
    S.sheets.forEach(function (s) { sheetSel.appendChild(h('option', { value: s.name, text: s.name })); });
    S.query.sheet = S.query.sheet || S.sheet || (S.sheets[0] && S.sheets[0].name) || '';
    sheetSel.value = S.query.sheet;
    sheetSel.onchange = function () { S.query.sheet = this.value; renderQueryBuilder(); };
    if (!S.query.conditions.length) addQueryCondition();
    var ds = S.query.sheet ? refDatasetOf(S.query.sheet) : S.dataset;
    var body = $('#queryCondBody'); body.innerHTML = '';
    S.query.conditions.forEach(function (c, idx) {
      var row = h('div', { class: 'cond-row' });
      var logic = h('select', { class: 'inp' }, [
        h('option', { value: 'WHERE', text: 'WHERE' }), h('option', { value: 'AND', text: 'AND' }), h('option', { value: 'OR', text: 'OR' })
      ]);
      logic.value = idx === 0 ? 'WHERE' : (c.logic || 'OR');
      logic.disabled = idx === 0;
      logic.addEventListener('change', function () { c.logic = this.value; });
      var col = h('select', { class: 'inp' });
      (ds.headers || []).forEach(function (hd) { col.appendChild(h('option', { value: String(hd.index), text: hd.name + ' (' + hd.letter + ')' })); });
      col.value = String(c.colIdx || 0);
      col.addEventListener('change', function () { c.colIdx = Number(this.value); });
      var op = h('select', { class: 'inp' }, [
        h('option', { value: 'contains', text: 'LIKE' }), h('option', { value: 'eq', text: '=' }),
        h('option', { value: 'notEq', text: '<>' }), h('option', { value: 'starts', text: 'STARTS' }),
        h('option', { value: 'blank', text: 'BLANK' }), h('option', { value: 'notBlank', text: 'NOT BLANK' })
      ]);
      op.value = c.op || 'contains';
      op.addEventListener('change', function () { c.op = this.value; });
      var val = h('input', { class: 'inp', value: c.value || '', placeholder: 'Giá trị hoặc $F$3' });
      val.addEventListener('input', function () { c.value = this.value; });
      var del = h('button', { class: 'btn ghost sm', text: 'Xóa' });
      del.disabled = S.query.conditions.length <= 1;
      del.addEventListener('click', function () { S.query.conditions.splice(idx, 1); renderQueryBuilder(); });
      row.appendChild(logic); row.appendChild(col); row.appendChild(op); row.appendChild(val); row.appendChild(del);
      body.appendChild(row);
    });
  }

  function colIndexFromLetter(letter) {
    var n = 0;
    String(letter || '').toUpperCase().replace(/[^A-Z]/g, '').split('').forEach(function (ch) { n = n * 26 + ch.charCodeAt(0) - 64; });
    return n - 1;
  }

  function cellValue(sheetName, addr) {
    var m = String(addr || '').match(/\$?([A-Z]+)\$?(\d+)/i);
    if (!m) return '';
    var aoa = aoaOf(sheetName || S.sheet);
    var r = Number(m[2]) - 1, c = colIndexFromLetter(m[1]);
    return aoa[r] ? U.toStr(aoa[r][c]).trim() : '';
  }

  function resolveQueryValue(raw) {
    var s = String(raw || '').trim();
    if ((s[0] === '"' && s[s.length - 1] === '"') || (s[0] === "'" && s[s.length - 1] === "'")) s = s.slice(1, -1);
    s = s.replace(/"&/g, '').replace(/&"/g, '').replace(/'&/g, '').replace(/&'/g, '');
    s = s.replace(/\$?[A-Z]+\$?\d+/gi, function (m) { return cellValue(S.sheet, m); });
    return s.replace(/^%|%$/g, '');
  }

  function datasetFromRange(sheetName, startCol, startRow, endCol) {
    var aoa = aoaOf(sheetName);
    var c0 = colIndexFromLetter(startCol || 'A');
    var c1 = endCol ? colIndexFromLetter(endCol) : Math.max.apply(null, aoa.map(function (r) { return (r || []).length - 1; }));
    var r0 = Math.max(0, Number(startRow || 1) - 1);
    var sliced = aoa.slice(r0).map(function (r) { return (r || []).slice(c0, c1 + 1); });
    return E.buildDataset(sliced, 0, {});
  }

  function parseQueryText(txt) {
    txt = String(txt || '').trim();
    if (!txt) return null;
    var out = { sheet: S.query.sheet || S.sheet, dataset: null, conditions: [] };
    var q = txt.match(/^QUERY\('([^']+)'\!\$?([A-Z]+)\$?(\d+)\:\$?([A-Z]+)?[^;]*;\s*"select\s+\*\s+where\s+(.+)"\)$/i);
    if (q) {
      out.sheet = q[1];
      out.dataset = datasetFromRange(q[1], q[2], q[3], q[4]);
      txt = 'select * where ' + q[5];
    }
    var where = txt.match(/^select\s+\*\s+where\s+(.+)$/i);
    if (!where) throw new Error('QUERY_PARSE');
    var parts = where[1].split(/\s+(AND|OR)\s+/i);
    for (var i = 0; i < parts.length; i += 2) {
      var logic = i === 0 ? 'WHERE' : String(parts[i - 1]).toUpperCase();
      var m = parts[i].trim().match(/^([A-Z]+|`[^`]+`)\s*(LIKE|=|<>|!=|CONTAINS|STARTS)\s*(.+)$/i);
      if (!m) throw new Error('QUERY_PARSE');
      out.conditions.push({ logic: logic, columnRef: m[1].replace(/`/g, ''), op: m[2].toUpperCase(), value: resolveQueryValue(m[3]) });
    }
    return out;
  }

  function queryColIndex(ds, ref) {
    if (/^[A-Z]+$/i.test(ref)) return colIndexFromLetter(ref);
    var want = U.normKey(ref);
    var h = ds.headers.find(function (x) { return x.key === want || U.normKey(x.name) === want; });
    return h ? h.index : -1;
  }

  function matchQueryValue(v, op, rawValue) {
    var s = U.toStr(v).trim(), n = U.normKey(s), q = U.normKey(resolveQueryValue(rawValue));
    if (op === 'blank') return U.isBlank(v);
    if (op === 'notBlank') return !U.isBlank(v);
    if (op === 'eq' || op === '=') return n === q;
    if (op === 'notEq' || op === '<>' || op === '!=') return n !== q;
    if (op === 'starts' || op === 'STARTS') return n.indexOf(q) === 0;
    return n.indexOf(q) >= 0;
  }

  function runQuery() {
    try {
      var parsed = parseQueryText($('#querySqlInput').value);
      var ds = parsed && parsed.dataset ? parsed.dataset : refDatasetOf((parsed && parsed.sheet) || S.query.sheet || S.sheet);
      var conditions = parsed ? parsed.conditions.map(function (c) {
        return { logic: c.logic, colIdx: queryColIndex(ds, c.columnRef), op: c.op === 'LIKE' || c.op === 'CONTAINS' ? 'contains' : c.op, value: c.value };
      }) : S.query.conditions;
      conditions = conditions.filter(function (c) { return c.colIdx >= 0; });
      if (!conditions.length) throw new Error('NO_COND');
      var rows = ds.rows.filter(function (row) {
        var ok = false;
        conditions.forEach(function (c, idx) {
          var hit = matchQueryValue(row[c.colIdx], c.op, c.value);
          if (idx === 0) ok = hit;
          else if (c.logic === 'AND') ok = ok && hit;
          else ok = ok || hit;
        });
        return ok;
      });
      S.query.result = { headers: ds.headers, rows: rows, rowNo: rows.map(function (_, i) { return i + 1; }), headerRowIdx: 0 };
      S.query.useAsReference = false;
      renderQueryResult();
      $('#useQueryAsRefBtn').disabled = rows.length === 0;
      U.toast('Query trả về ' + U.fmtInt(rows.length) + ' dòng.', rows.length ? 'ok' : 'info');
    } catch (e) {
      U.toast(e.message === 'NO_COND' ? 'Chưa có điều kiện query hợp lệ.' : 'Chưa đọc được cú pháp query/SQL.', 'err');
    }
  }

  function renderQueryResult() {
    var host = $('#queryResultHost'); host.innerHTML = '';
    if (!S.query.result) {
      host.appendChild(h('div', { class: 'empty-state', text: 'Chạy query để xem nhanh dữ liệu hoặc dùng kết quả làm danh mục tham chiếu tạm.' }));
      return;
    }
    var table = h('table');
    var thead = h('thead'), trh = h('tr');
    S.query.result.headers.forEach(function (hd) { trh.appendChild(h('th', { text: hd.name })); });
    thead.appendChild(trh); table.appendChild(thead);
    var tb = h('tbody');
    S.query.result.rows.slice(0, 50).forEach(function (r) {
      var tr = h('tr');
      S.query.result.headers.forEach(function (hd) { tr.appendChild(h('td', { text: U.toStr(r[hd.index]) })); });
      tb.appendChild(tr);
    });
    table.appendChild(tb); host.appendChild(table);
    if (S.query.result.rows.length > 50) host.appendChild(h('div', { class: 'tip', text: 'Đang hiển thị 50 dòng đầu trên tổng ' + U.fmtInt(S.query.result.rows.length) + ' dòng.' }));
  }

  function useQueryAsReference() {
    if (!S.query.result || !S.query.result.rows.length) return;
    S.query.useAsReference = true;
    renderMapping(false);
    U.toast('Đã bật kết quả query làm danh mục tham chiếu tạm.', 'ok');
  }

  function renderMapping(withAuto) {
    var rs = R.byId(S.ruleSetId);
    $('#ruleSetDesc').textContent = rs.desc || '';
    if (withAuto) autoMap();

    var body = $('#mapBody'); body.innerHTML = '';
    rs.fields.forEach(function (f) {
      var sel = h('select', { class: 'inp', style: 'width:100%' });
      sel.appendChild(h('option', { value: '-1', text: '— Không dùng —' }));
      S.dataset.headers.forEach(function (hd) {
        sel.appendChild(h('option', { value: String(hd.index), text: hd.name + '  (' + hd.letter + ')' }));
      });
      sel.value = String(S.mapping[f.key] === undefined ? -1 : S.mapping[f.key]);
      sel.addEventListener('change', function () {
        var v = Number(this.value);
        if (v < 0) delete S.mapping[f.key]; else S.mapping[f.key] = v;
        renderMapping(false);
      });

      var mapped = S.mapping[f.key] !== undefined;
      var required = !!f.rules.required;
      var status = mapped
        ? h('span', { class: 'map-status ok', text: '✓ Đã ghép' })
        : (required ? h('span', { class: 'map-status miss', text: '✕ Chưa ghép' })
          : h('span', { class: 'map-status warn', text: '⚠ Bỏ qua' }));

      var chips = h('div', { class: 'rule-chips' });
      Object.keys(f.rules).forEach(function (k) {
        var label = R.labels[k];
        if (!label) return;
        var v = f.rules[k];
        var txt = label;
        if (k === 'minLength' || k === 'maxLength' || k === 'min' || k === 'max') txt = label + ' ' + v;
        if (k === 'type') txt = 'Kiểu: ' + v;
        if (k === 'allowedValues') txt = 'Cho phép: ' + v.slice(0, 3).join('/');
        chips.appendChild(h('span', { class: 'chip', text: txt }));
      });

      body.appendChild(h('tr', {}, [
        h('td', { class: 'map-field' }, [
          h('b', {}, [document.createTextNode(f.label), required ? h('span', { class: 'req-dot', text: '*' }) : null]),
          h('small', { text: f.key })
        ]),
        h('td', {}, [sel]),
        h('td', {}, [status]),
        h('td', {}, [chips])
      ]));
    });

    // Danh mục tham chiếu
    var refFields = rs.fields.filter(function (f) { return f.rules.reference; });
    var refCard = $('#refCard'), refBody = $('#refBody');
    refBody.innerHTML = '';
    refCard.classList.toggle('hidden', refFields.length === 0);
    refFields.forEach(function (f) {
      var cur = S.refs[f.key];
      var sheetSel = h('select', { class: 'inp' });
      sheetSel.appendChild(h('option', { value: '', text: '— Không đối chiếu —' }));
      S.sheets.forEach(function (s) {
        if (s.name === S.sheet) return;
        sheetSel.appendChild(h('option', { value: s.name, text: s.name }));
      });
      sheetSel.value = cur ? cur.sheet : '';

      var colSel = h('select', { class: 'inp' });
      function fillCols(sheetName, selected) {
        colSel.innerHTML = '';
        if (!sheetName) { colSel.disabled = true; return; }
        colSel.disabled = false;
        var ds = refDatasetOf(sheetName);
        ds.headers.forEach(function (hd) {
          colSel.appendChild(h('option', { value: String(hd.index), text: hd.name }));
        });
        if (selected !== undefined && selected !== null) colSel.value = String(selected);
      }
      fillCols(cur ? cur.sheet : '', cur ? cur.colIdx : null);

      sheetSel.addEventListener('change', function () {
        var sn = this.value;
        if (!sn) { S.refs[f.key] = null; fillCols(''); return; }
        var col = findRefColumn(sn, f.rules.reference.columnHint, f);
        S.refs[f.key] = { sheet: sn, column: col ? col.name : '', colIdx: col ? col.index : 0 };
        fillCols(sn, col ? col.index : 0);
      });
      colSel.addEventListener('change', function () {
        var sn = sheetSel.value; if (!sn) return;
        var ds = refDatasetOf(sn);
        var hd = ds.headers[Number(this.value)];
        S.refs[f.key] = { sheet: sn, column: hd.name, colIdx: hd.index };
      });

      refBody.appendChild(h('tr', {}, [
        h('td', { class: 'map-field', style: 'width:26%' }, [
          h('b', { text: f.label }), h('small', { text: f.rules.reference.label || 'Danh mục' })
        ]),
        h('td', { style: 'width:28%' }, [sheetSel]),
        h('td', {}, [colSel]),
        h('td', {}, [h('span', {
          class: 'map-status ' + (S.refs[f.key] ? 'ok' : 'warn'),
          text: S.refs[f.key] ? '✓ Sẽ đối chiếu' : '⚠ Bỏ qua đối chiếu'
        })])
      ]));
    });

    renderAdvancedRefs(rs);
    renderQueryBuilder();
    renderQueryResult();

    var missing = rs.fields.filter(function (f) { return f.rules.required && S.mapping[f.key] === undefined; });
    $('#mapStatusText').textContent = missing.length
      ? missing.length + ' field bắt buộc chưa ghép: ' + missing.map(function (f) { return f.label; }).join(', ')
      : 'Tất cả field bắt buộc đã được ghép.';
    $('#mapStatusText').style.color = missing.length ? 'var(--warn)' : 'var(--valid)';
    $('#runValidationBtn').disabled = missing.length > 0;
  }

  /* ================= Chạy kiểm tra ================= */
  function showProgress(title, detail, count) {
    $('#progTitle').textContent = title;
    $('#progDetail').textContent = detail || '';
    $('#progCount').textContent = count || '';
    $('#progBar').style.width = '0%';
    $('#progressOverlay').classList.remove('hidden');
  }
  function updateProgress(p) {
    var pct = p.total ? Math.round(p.done / p.total * 100) : 0;
    $('#progBar').style.width = pct + '%';
    $('#progDetail').textContent = p.phase || '';
    $('#progCount').textContent = 'Đã xử lý ' + U.fmtInt(p.done) + ' / ' + U.fmtInt(p.total);
  }
  function hideProgress() { $('#progressOverlay').classList.add('hidden'); }

  function runQuickCheck() {
    S.token = { cancelled: false };
    showProgress('Đang phân tích dữ liệu…', 'Chuẩn bị', '');
    var t0 = Date.now();
    E.quickCheck(S.dataset, { sheet: S.sheet }, updateProgress, S.token)
      .then(function (res) {
        S.issues = res.issues; S.profiles = res.profiles;
        finishRun(t0);
      })
      .catch(handleRunError);
  }

  function runRuleValidation() {
    var rs = R.byId(S.ruleSetId);
    S.token = { cancelled: false };
    showProgress('Đang kiểm tra dữ liệu…', 'Dựng danh mục tham chiếu', '');
    var t0 = Date.now();
    setTimeout(function () {
      var refs = {};
      try {
        Object.keys(S.refs).forEach(function (k) {
          var r = S.refs[k]; if (!r) return;
          var ds = refDatasetOf(r.sheet);
          refs[k] = { sheet: r.sheet, column: r.column, set: E.buildReferenceSet(ds, r.colIdx) };
        });
      } catch (e) { hideProgress(); U.toast('Không dựng được danh mục tham chiếu.', 'err'); return; }
      var advancedRefs = [];
      try {
        advancedRefs = (S.advancedRefs || []).filter(function (r) { return r && r.enabled !== false && r.sheet; }).map(function (r) {
          var ds = refDatasetByName(r.sheet);
          return Object.assign({}, r, { dataset: ds, sheet: r.sheet === '__QUERY__' ? 'Kết quả query tạm' : r.sheet });
        }).filter(function (r) { return r.dataset && r.dataset.rows && r.dataset.rows.length; });
      } catch (e2) { hideProgress(); U.toast('Không dựng được tham chiếu nâng cao.', 'err'); return; }

      E.validate(S.dataset, rs, S.mapping, refs, updateProgress, S.token, { sheet: S.sheet, advancedRefs: advancedRefs })
        .then(function (res) {
          S.issues = res.issues;
          S.profiles = null;
          finishRun(t0);
        })
        .catch(handleRunError);
    }, 30);
  }

  function handleRunError(err) {
    hideProgress();
    if (err && err.message === 'CANCELLED') U.toast('Đã hủy kiểm tra dữ liệu', 'info');
    else { console.error(err); U.toast('Có lỗi khi kiểm tra dữ liệu. Thử lại với sheet khác.', 'err'); }
  }

  function finishRun(t0) {
    S.summary = E.summarize(S.dataset, S.issues);
    S.index = E.indexIssues(S.issues);
    buildFilterMaps();
    hideProgress();
    S.filter = { tab: 'all', rule: null, column: null, search: '', scope: '*' };
    $('#searchInput').value = '';
    renderResult();
    gotoView('result');
    saveHistory();
    U.toast('Kiểm tra xong sau ' + ((Date.now() - t0) / 1000).toFixed(1) + 's — ' +
      U.fmtInt(S.summary.errorCount) + ' lỗi, ' + U.fmtInt(S.summary.warningCount) + ' cảnh báo', 'ok');
  }

  function buildFilterMaps() {
    S.byRuleRows = new Map(); S.byColumnRows = new Map();
    S.issues.forEach(function (is) {
      if (is.rowIndex < 0) return;
      if (!S.byRuleRows.has(is.ruleType)) S.byRuleRows.set(is.ruleType, new Set());
      S.byRuleRows.get(is.ruleType).add(is.rowIndex);
      var ck = is.column || '(bảng)';
      if (!S.byColumnRows.has(ck)) S.byColumnRows.set(ck, new Set());
      S.byColumnRows.get(ck).add(is.rowIndex);
    });
  }

  /* ================= VIEW: RESULT ================= */
  function bindResultView() {
    U.$$('#resultTabs .tab').forEach(function (t) {
      t.addEventListener('click', function () {
        U.$$('#resultTabs .tab').forEach(function (x) { x.classList.remove('on'); });
        t.classList.add('on');
        S.filter.tab = t.getAttribute('data-tab');
        applyFilter();
      });
    });
    var si = $('#searchInput'), timer = null;
    si.addEventListener('input', function () {
      clearTimeout(timer);
      timer = setTimeout(function () { S.filter.search = si.value.trim(); applyFilter(); }, 180);
    });
    $('#searchScope').addEventListener('change', function () { S.filter.scope = this.value; applyFilter(); });
    $('#exportBtn').addEventListener('click', openExport);
    $('#copyErrorsBtn').addEventListener('click', function () {
      U.copyText(X.errorsToText(exportCtx(), 500)).then(function () {
        U.toast('Đã sao chép danh sách lỗi', 'ok');
      }).catch(function () { U.toast('Không sao chép được, thử xuất tệp báo cáo.', 'err'); });
    });
    $('#railToggle').addEventListener('click', function () { $('#rail').classList.toggle('mobile-open'); });
    window.addEventListener('resize', function () {
      $('#railToggle').classList.toggle('hidden', window.innerWidth > 820);
      if (window.innerWidth > 820) $('#rail').classList.remove('mobile-open');
    });
  }

  function renderResult() {
    var s = S.summary;
    var cards = $('#summaryCards'); cards.innerHTML = '';
    [
      { n: U.fmtInt(s.totalRows), l: 'Bản ghi', c: '' },
      { n: U.fmtInt(s.validRows), l: 'Hợp lệ', c: 'v' },
      { n: U.fmtInt(s.errorCount), l: 'Lỗi', c: 'e' },
      { n: U.fmtInt(s.warningCount), l: 'Cảnh báo', c: 'w' },
      { n: s.qualityScore + '%', l: 'Chất lượng', c: '' }
    ].forEach(function (x) {
      cards.appendChild(h('div', { class: 'scard ' + x.c, title: x.l === 'Chất lượng' ? 'Tỷ lệ dòng không có lỗi trên tổng số dòng kiểm tra.' : '' }, [
        h('div', { class: 'sc-num', text: x.n }),
        h('div', { class: 'sc-lbl', text: x.l })
      ]));
    });

    $('#cntAll').textContent = U.fmtInt(s.totalRows);
    $('#cntErr').textContent = U.fmtInt(s.invalidRows);
    $('#cntWarn').textContent = U.fmtInt(s.warnRowSet.size);
    $('#cntValid').textContent = U.fmtInt(s.totalRows - s.invalidRows - countWarnOnly());

    var scope = $('#searchScope'); scope.innerHTML = '';
    scope.appendChild(h('option', { value: '*', text: 'Mọi cột' }));
    S.dataset.headers.forEach(function (hd) {
      scope.appendChild(h('option', { value: String(hd.index), text: hd.name }));
    });

    $('#sbSheet').textContent = S.sheet;
    $('#sbRows').textContent = U.fmtInt(s.totalRows);
    $('#sbErr').textContent = U.fmtInt(s.errorCount);
    $('#sbWarn').textContent = U.fmtInt(s.warningCount);
    $('#sbQuality').textContent = s.qualityScore + '%';

    var rd = $('#readiness');
    if (s.errorCount === 0 && s.warningCount === 0) { rd.className = 'readiness ready'; rd.textContent = '✓ Dữ liệu sẵn sàng import'; }
    else if (s.errorCount === 0) { rd.className = 'readiness warn'; rd.textContent = '✓ Sẵn sàng — còn cảnh báo'; }
    else { rd.className = 'readiness notready'; rd.textContent = '✕ Chưa import được'; }

    renderRail();

    if (!S.grid) {
      S.grid = new root.ASC.Grid($('#gridHost'), {
        onCellClick: function (r, c, hit) { showCellDetail(r, c, hit); },
        onHeaderClick: function (c) { showColumnProfile(c); }
      });
    }
    S.grid.setData(S.dataset);
    S.grid.setIssueIndex(S.index);
    applyFilter();
    $('#railToggle').classList.toggle('hidden', window.innerWidth > 820);
    showDefaultDetail();
  }

  function countWarnOnly() {
    var n = 0;
    S.summary.warnRowSet.forEach(function (r) { if (!S.summary.errorRowSet.has(r)) n++; });
    return n;
  }

  function renderRail() {
    var rail = $('#rail'); rail.innerHTML = '';
    var s = S.summary;

    var sec1 = h('div', { class: 'rail-sec' }, [h('span', { class: 'eyebrow', text: 'Lỗi theo loại' })]);
    if (!s.issuesByRule.length) {
      sec1.appendChild(h('div', { class: 'empty-state', style: 'padding:12px 0', text: '✓ Không phát hiện lỗi nào.' }));
    }
    var max = s.issuesByRule.length ? s.issuesByRule[0].total : 1;
    s.issuesByRule.forEach(function (r) {
      var label = R.typeLabels[r.rule] || r.rule;
      var item = h('button', { class: 'rail-item' + (S.filter.rule === r.rule ? ' on' : '') }, [
        h('div', { class: 'ri-top' }, [
          h('span', { text: label, title: label }),
          h('b', { text: U.fmtInt(r.total) })
        ]),
        h('div', { class: 'ri-bar' }, [
          h('i', { class: r.error ? 'e' : 'w', style: 'width:' + Math.max(4, Math.round(r.total / max * 100)) + '%' })
        ])
      ]);
      item.addEventListener('click', function () {
        S.filter.rule = S.filter.rule === r.rule ? null : r.rule;
        S.filter.column = null;
        renderRail(); applyFilter(); showRuleDetail(r.rule);
      });
      sec1.appendChild(item);
    });
    rail.appendChild(sec1);

    var sec2 = h('div', { class: 'rail-sec' }, [h('span', { class: 'eyebrow', text: 'Lỗi theo cột' })]);
    s.issuesByColumn.slice(0, 14).forEach(function (r) {
      var item = h('button', { class: 'rail-item' + (S.filter.column === r.column ? ' on' : '') }, [
        h('div', { class: 'ri-top' }, [
          h('span', { text: r.column, title: r.column }),
          h('b', { text: U.fmtInt(r.total) })
        ]),
        h('div', { class: 'ri-bar' }, [
          h('i', { class: r.error ? 'e' : 'w', style: 'width:' + Math.max(4, Math.round(r.total / max * 100)) + '%' })
        ])
      ]);
      item.addEventListener('click', function () {
        S.filter.column = S.filter.column === r.column ? null : r.column;
        S.filter.rule = null;
        renderRail(); applyFilter();
      });
      sec2.appendChild(item);
    });
    rail.appendChild(sec2);
  }

  function applyFilter() {
    var f = S.filter, idx = S.index, ds = S.dataset;
    var rows = [];
    var searchLower = f.search.toLowerCase();
    var scopeCol = f.scope === '*' ? -1 : Number(f.scope);

    for (var i = 0; i < ds.rows.length; i++) {
      var st = idx.row.get(i);
      if (f.tab === 'error' && !(st && st.error)) continue;
      if (f.tab === 'warning' && !(st && st.warning)) continue;
      if (f.tab === 'valid' && st) continue;
      if (f.rule && !(S.byRuleRows.get(f.rule) || new Set()).has(i)) continue;
      if (f.column && !(S.byColumnRows.get(f.column) || new Set()).has(i)) continue;
      if (searchLower) {
        var found = false;
        if (scopeCol >= 0) found = U.toStr(ds.rows[i][scopeCol]).toLowerCase().indexOf(searchLower) >= 0;
        else {
          for (var c = 0; c < ds.headers.length; c++) {
            if (U.toStr(ds.rows[i][c]).toLowerCase().indexOf(searchLower) >= 0) { found = true; break; }
          }
        }
        if (!found) continue;
      }
      rows.push(i);
    }
    S.grid.setVisibleRows(rows);
    $('#resultCount').textContent = U.fmtInt(rows.length) + ' / ' + U.fmtInt(ds.rows.length) + ' dòng';
    renderActiveChips();
  }

  function renderActiveChips() {
    var host = $('#activeFilterHost'); host.innerHTML = '';
    function chip(text, clear) {
      var c = h('span', { class: 'filter-chip' }, [
        document.createTextNode(text),
        h('button', { text: '✕', title: 'Bỏ lọc', onclick: clear })
      ]);
      host.appendChild(c);
    }
    if (S.filter.rule) chip(R.typeLabels[S.filter.rule] || S.filter.rule, function () { S.filter.rule = null; renderRail(); applyFilter(); });
    if (S.filter.column) chip('Cột: ' + S.filter.column, function () { S.filter.column = null; renderRail(); applyFilter(); });
  }

  /* ---------- Panel chi tiết ---------- */
  function detailShell() {
    var d = $('#detail');
    d.innerHTML = '';
    var close = h('button', { class: 'btn ghost sm detail-close', text: 'Đóng bảng chi tiết' });
    close.addEventListener('click', function () { d.classList.remove('mobile-open'); });
    d.appendChild(close);
    return d;
  }

  function showDefaultDetail() {
    var d = detailShell();
    if (S.summary.errorCount === 0 && S.summary.warningCount === 0) {
      d.appendChild(h('div', { class: 'empty-state', text: '✓ Không phát hiện lỗi theo bộ rule hiện tại.' }));
      return;
    }
    d.appendChild(h('div', { class: 'empty-state', text: 'Bấm vào một ô được tô màu, hoặc chọn một loại lỗi ở cột trái để xem chi tiết.' }));
  }

  function sevBadge(sev) {
    var label = { error: 'Lỗi', warning: 'Cảnh báo', info: 'Gợi ý' }[sev] || sev;
    return h('span', { class: 'sev-badge ' + sev, text: label });
  }

  function field(label, valueEl) {
    return h('div', { class: 'dfield' }, [h('span', { class: 'eyebrow', text: label }), valueEl]);
  }

  function showCellDetail(r, c, hit) {
    var d = detailShell();
    var hd = S.dataset.headers[c];
    var value = U.toStr(S.dataset.rows[r][c]);
    if (!hit) {
      d.appendChild(h('h4', {}, [document.createTextNode('Ô hợp lệ')]));
      d.appendChild(field('Cột', h('div', { class: 'dval', text: hd.name })));
      d.appendChild(field('Dòng', h('div', { class: 'dval mono', text: String(S.dataset.rowNo[r]) })));
      d.appendChild(field('Giá trị', h('div', { class: 'dval mono', text: value === '' ? '(trống)' : value })));
      d.classList.add('mobile-open');
      return;
    }
    hit.list.forEach(function (is, k) {
      if (k > 0) d.appendChild(h('hr', { style: 'border:0;border-top:1px solid var(--line);margin:14px 0' }));
      d.appendChild(h('h4', {}, [
        document.createTextNode(R.typeLabels[is.ruleType] || is.ruleType), sevBadge(is.severity)
      ]));
      d.appendChild(field('Sheet', h('div', { class: 'dval', text: is.sheet || S.sheet })));
      d.appendChild(field('Dòng · cột', h('div', { class: 'dval mono', text: is.rowNo + ' · ' + is.column + ' (' + hd.letter + ')' })));
      d.appendChild(field('Giá trị hiện tại', h('div', { class: 'dval mono', text: value === '' ? '(trống)' : '"' + value + '"' })));
      d.appendChild(field('Rule', h('div', { class: 'dval mono', text: is.ruleType + (is.field ? ' · ' + is.field : '') })));
      d.appendChild(field('Nội dung', h('div', { class: 'dval', text: is.message })));
      if (is.suggestion) d.appendChild(field('Gợi ý xử lý', h('div', { class: 'suggestion', text: is.suggestion })));
      if (is.relatedRows && is.relatedRows.length > 1) {
        var list = h('div', { class: 'dup-list' });
        is.relatedRows.slice(0, 100).forEach(function (rn) {
          var b = h('button', { text: 'Dòng ' + rn });
          b.addEventListener('click', function () { jumpToRowNo(rn); });
          list.appendChild(b);
        });
        d.appendChild(field('Các dòng liên quan (' + is.relatedRows.length + ')', list));
      }
      var copyBtn = h('button', { class: 'btn sm', text: 'Sao chép lỗi này' });
      copyBtn.addEventListener('click', function () {
        U.copyText('Dòng ' + is.rowNo + ' - ' + is.column + ' - ' + (value || '(trống)') + ' - ' + is.message)
          .then(function () { U.toast('Đã sao chép', 'ok'); });
      });
      d.appendChild(copyBtn);
    });
    d.classList.add('mobile-open');
  }

  function jumpToRowNo(rowNo) {
    var idx = S.dataset.rowNo.indexOf(rowNo);
    if (idx < 0) return;
    if (!S.grid.scrollToRow(idx)) {
      S.filter = { tab: 'all', rule: null, column: null, search: '', scope: '*' };
      $('#searchInput').value = '';
      U.$$('#resultTabs .tab').forEach(function (x) { x.classList.toggle('on', x.getAttribute('data-tab') === 'all'); });
      renderRail(); applyFilter();
      S.grid.scrollToRow(idx);
    }
    var firstIssue = S.issues.find(function (is) { return is.rowIndex === idx; });
    if (firstIssue) S.grid.select(idx, firstIssue.col);
  }

  function showRuleDetail(ruleType) {
    var d = detailShell();
    var list = S.issues.filter(function (is) { return is.ruleType === ruleType; });
    if (!list.length) return;
    d.appendChild(h('h4', {}, [
      document.createTextNode(R.typeLabels[ruleType] || ruleType), sevBadge(list[0].severity)
    ]));
    d.appendChild(field('Số lượng', h('div', { class: 'dval mono', text: U.fmtInt(list.length) + ' mục' })));
    d.appendChild(field('Nội dung', h('div', { class: 'dval', text: list[0].message })));
    if (list[0].suggestion) d.appendChild(field('Gợi ý xử lý', h('div', { class: 'suggestion', text: list[0].suggestion })));

    var wrap = h('div', { class: 'dup-list' });
    list.slice(0, 200).forEach(function (is) {
      var b = h('button', { text: 'Dòng ' + (is.rowIndex < 0 ? '(tiêu đề)' : is.rowNo) + ' · ' + is.column + ' · ' + (U.toStr(is.value) || '(trống)') });
      b.addEventListener('click', function () {
        if (is.rowIndex < 0) return;
        jumpToRowNo(is.rowNo);
        showCellDetail(is.rowIndex, is.col, S.index.cell.get(is.rowIndex + '|' + is.col));
      });
      wrap.appendChild(b);
    });
    d.appendChild(field('Danh sách' + (list.length > 200 ? ' (200 mục đầu)' : ''), wrap));
    d.classList.add('mobile-open');
  }

  function showColumnProfile(c) {
    var d = detailShell();
    var p = (S.profiles && S.profiles[c]) || E.profileColumn(S.dataset, c);
    var hd = S.dataset.headers[c];
    d.appendChild(h('h4', {}, [document.createTextNode('Thống kê cột')]));
    d.appendChild(field('Cột', h('div', { class: 'dval', text: hd.name + ' (' + hd.letter + ')' })));

    var rows = [
      ['Tổng số dòng', U.fmtInt(p.total)],
      ['Có dữ liệu', U.fmtInt(p.filled)],
      ['Ô trống', U.fmtInt(p.blank)],
      ['Giá trị khác nhau', U.fmtInt(p.unique)],
      ['Giá trị bị trùng', U.fmtInt(p.duplicateValues)],
      ['Độ dài min / max / TB', p.minLength + ' / ' + p.maxLength + ' / ' + p.avgLength],
      ['Kiểu chiếm đa số', p.dominantType]
    ];
    if (p.numeric) rows.push(['Số: min / max / TB', p.numeric.min + ' / ' + p.numeric.max + ' / ' + p.numeric.avg]);
    var box = h('div', {});
    rows.forEach(function (r) {
      box.appendChild(h('div', { class: 'profile-row' }, [h('span', { text: r[0] }), h('b', { text: String(r[1]) })]));
    });
    d.appendChild(box);

    if (p.top.length) {
      var maxc = p.top[0][1];
      var freq = h('div', {});
      p.top.slice(0, 15).forEach(function (t) {
        freq.appendChild(h('div', {}, [
          h('div', { class: 'freq-row' }, [
            h('span', { class: 'fr-val', text: t[0] === '' ? '(trống)' : t[0], title: t[0] }),
            h('span', { class: 'fr-cnt', text: U.fmtInt(t[1]) })
          ]),
          h('div', { class: 'freq-bar' }, [h('i', { style: 'width:' + Math.round(t[1] / maxc * 100) + '%' })])
        ]));
      });
      d.appendChild(field('Giá trị phổ biến (' + U.fmtInt(p.distinctCount) + ' giá trị khác nhau)', freq));
    }
    d.classList.add('mobile-open');
  }

  /* ================= Export ================= */
  function exportCtx() {
    var rs = S.ruleSetId ? R.byId(S.ruleSetId) : null;
    return {
      fileName: S.file.name, sheetName: S.sheet, dataset: S.dataset, issues: S.issues,
      summary: S.summary, index: S.index, mode: S.mode,
      ruleSetName: S.mode === 'rule' && rs ? rs.name : null, project: null
    };
  }

  function renderExportOptions() {
    var body = $('#exportBody'); body.innerHTML = '';
    var opts = [
      { t: 'Excel có đánh dấu lỗi', d: 'Giữ nguyên dữ liệu gốc, tô màu ô lỗi kèm ghi chú, thêm sheet DATACHECK_ERRORS và DATACHECK_SUMMARY.', fn: 'excelWithErrors' },
      { t: 'Chỉ các dòng có lỗi', d: 'Tệp gọn để gửi khách hàng chỉnh sửa.', fn: 'invalidRows' },
      { t: 'Chỉ các dòng hợp lệ', d: 'Phần dữ liệu không còn lỗi, có thể dùng để import.', fn: 'validRows' },
      { t: 'Báo cáo lỗi', d: 'Chỉ gồm danh sách lỗi và bảng tổng hợp.', fn: 'errorReport' }
    ];
    opts.forEach(function (o) {
      var btn = h('button', { class: 'export-opt' }, [
        h('span', { class: 'eo-ico', html: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M12 4v11m0 0l-4-4m4 4l4-4" stroke-linecap="round" stroke-linejoin="round"/><path d="M4 18.5h16" stroke-linecap="round"/></svg>' }),
        h('span', {}, [h('b', { text: o.t }), h('span', { text: o.d })])
      ]);
      btn.addEventListener('click', function () {
        btn.disabled = true;
        showProgress('Đang tạo tệp export…', o.t, '');
        setTimeout(function () {
          X[o.fn](exportCtx())
            .then(function (name) { hideProgress(); closeExport(); U.toast('Đã xuất ' + name, 'ok'); })
            .catch(function (err) { hideProgress(); U.toast(err.message || 'Không xuất được tệp.', 'err'); })
            .then(function () { btn.disabled = false; });
        }, 40);
      });
      body.appendChild(btn);
    });

    var copyBtn = h('button', { class: 'export-opt' }, [
      h('span', { class: 'eo-ico', html: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5.8C5 5.4 5.4 5 5.8 5H15" stroke-linecap="round"/></svg>' }),
      h('span', {}, [h('b', { text: 'Sao chép danh sách lỗi' }), h('span', { text: 'Dạng văn bản gọn để dán vào Teams, Zalo hoặc email.' })])
    ]);
    copyBtn.addEventListener('click', function () {
      U.copyText(X.errorsToText(exportCtx(), 500)).then(function () { closeExport(); U.toast('Đã sao chép danh sách lỗi', 'ok'); });
    });
    body.appendChild(copyBtn);
  }

  document.addEventListener('DOMContentLoaded', init);
})(typeof window !== 'undefined' ? window : globalThis);
