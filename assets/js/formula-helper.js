/*!
 * formula-helper.js — Workspace Formula Helper.
 *
 * Dựng UI theo nhu cầu nghiệp vụ ("Bạn muốn làm gì?"), sinh công thức qua
 * FormulaEngine, mô phỏng kết quả qua FormulaExplainer và lưu History/Favorite
 * qua FormulaStore. Không đụng tới generator.js.
 */
(function (global) {
  'use strict';

  var WB = null, ENGINE = null, PRESETS = null, STORE = null, EXPLAINER = null;

  var host = {
    toast: function () {},
    onImport: function () {},
    onWorkbookEdit: function () {},
    getSelectedColumn: function () { return null; },
    setActiveSheet: function () {}
  };

  var state = {
    view: 'home',
    toolId: null,
    config: {},
    settings: null,
    showAdvanced: false,
    gridOpen: false,
    reviewOpen: false,
    sidebarCollapsed: false,
    sidebarWidth: 420,
    query: '',
    lastResult: null,
    previewTimer: null
  };

  var root = null;
  var nodes = {};
  var dataGrid = null;
  var dataGridSheetIndex = -1;
  var dataGridSyncing = false;

  /* ============================================================
   * Tiện ích DOM
   * ========================================================= */

  function h(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }

  function svgIcon(paths, cls) {
    var wrap = document.createElement('span');
    wrap.className = 'fh-icon' + (cls ? ' ' + cls : '');
    wrap.setAttribute('aria-hidden', 'true');
    wrap.innerHTML = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" ' +
      'stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">' + paths + '</svg>';
    return wrap;
  }

  function button(label, cls, onClick) {
    var b = h('button', cls || 'btn btn-tiny', label);
    b.type = 'button';
    if (onClick) b.addEventListener('click', onClick);
    return b;
  }

  function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); }

  /* ============================================================
   * Cấu hình & tham chiếu
   * ========================================================= */

  function currentSheetName() {
    return WB.getSheetName(WB.getActiveIndex());
  }

  function makeRef(sheetIndex, colIndex) {
    var headers = WB.getHeaders(sheetIndex);
    var col = headers[colIndex];
    return {
      mode: 'column',
      sheetIndex: sheetIndex,
      sheetName: WB.getSheetName(sheetIndex),
      colIndex: colIndex,
      column: col ? col.name : WB.columnLetter(colIndex),
      type: col ? col.type : 'auto'
    };
  }

  /** Làm mới nhãn sheet/cột của ref theo workbook hiện tại (§167). */
  function refreshRef(ref) {
    if (!ref || ref.mode !== 'column') return ref;
    var name = WB.getSheetName(ref.sheetIndex);
    if (name) ref.sheetName = name;
    var headers = WB.getHeaders(ref.sheetIndex);
    if (headers[ref.colIndex]) {
      ref.column = headers[ref.colIndex].name;
      ref.type = headers[ref.colIndex].type;
    }
    return ref;
  }

  function selectedFormulaGridColumn() {
    if (dataGrid && state.gridOpen && dataGridSheetIndex >= 0) {
      var sel = dataGrid.getSelection();
      if (sel) return { sheetIndex: dataGridSheetIndex, colIndex: sel.column };
    }
    return host.getSelectedColumn();
  }

  function tool() { return state.toolId ? PRESETS.getTool(state.toolId) : null; }

  function buildConfig() {
    var t = tool();
    if (!t) return null;
    var cfg = {};
    Object.keys(state.config).forEach(function (k) { cfg[k] = state.config[k]; });
    cfg.type = t.type;
    cfg.platform = state.settings.platform;
    cfg.separator = state.settings.separator;
    cfg.excelVersion = state.settings.excelVersion;
    cfg.rangeMode = state.settings.rangeMode;
    cfg.referenceType = state.settings.referenceType;
    cfg.dataRow = state.settings.dataRow;
    cfg.currentSheet = currentSheetName();
    cfg.usedRows = Math.max(WB.getDataRowCount(WB.getActiveIndex()) + state.settings.dataRow - 1, 10);
    return cfg;
  }

  function saveSettings(patch) {
    state.settings = STORE.saveSettings(patch);
  }

  /* ============================================================
   * Khung workspace
   * ========================================================= */

  function mount(container) {
    root = container;
    clear(root);

    var bar = h('div', 'fh-bar');

    nodes.platform = selectControl('Nền tảng', [
      { value: 'excel', label: 'Microsoft Excel' },
      { value: 'sheets', label: 'Google Sheets' }
    ], state.settings.platform, function (v) {
      saveSettings({ platform: v });
      renderPanel();
    });

    nodes.separator = selectControl('Dấu phân cách', [
      { value: 'auto', label: 'Tự động' },
      { value: 'comma', label: 'Dấu phẩy  ,' },
      { value: 'semicolon', label: 'Chấm phẩy  ;' }
    ], state.settings.separator, function (v) {
      saveSettings({ separator: v });
      updateOutput();
    });

    nodes.dataRow = numberControl('Dòng dữ liệu bắt đầu', state.settings.dataRow, function (v) {
      saveSettings({ dataRow: Math.max(1, v || 2) });
      updateOutput();
    });

    bar.appendChild(nodes.platform);
    bar.appendChild(nodes.separator);
    bar.appendChild(nodes.dataRow);

    var barActions = h('div', 'fh-bar-actions');
    nodes.sidebarToggle = button('Ẩn thanh công cụ', 'btn btn-tiny', function () {
      setSidebarCollapsed(!state.sidebarCollapsed);
    });
    barActions.appendChild(nodes.sidebarToggle);
    nodes.resultToggle = button('Xem kết quả', 'btn btn-tiny', function () {
      setReviewOpen(true);
    });
    barActions.appendChild(nodes.resultToggle);
    barActions.appendChild(button('Thư viện hàm', 'btn btn-tiny', function () {
      state.view = state.view === 'library' ? 'home' : 'library';
      renderPanel();
    }));
    barActions.appendChild(button('Tuỳ chọn nâng cao', 'btn btn-tiny', function () {
      state.showAdvanced = !state.showAdvanced;
      renderPanel();
    }));
    bar.appendChild(barActions);
    root.appendChild(bar);

    nodes.workArea = h('div', 'fh-workarea');
    nodes.dataGridPanel = h('section', 'fh-data-grid-panel');
    nodes.workArea.appendChild(nodes.dataGridPanel);

    nodes.toolDock = h('aside', 'fh-tool-dock');
    nodes.resizeHandle = h('div', 'fh-tool-resizer');
    nodes.resizeHandle.title = 'Kéo để đổi độ rộng thanh công cụ';
    nodes.resizeHandle.addEventListener('mousedown', startSidebarResize);
    nodes.toolDock.appendChild(nodes.resizeHandle);
    nodes.panel = h('div', 'fh-panel');
    nodes.toolDock.appendChild(nodes.panel);
    nodes.workArea.appendChild(nodes.toolDock);
    root.appendChild(nodes.workArea);

    nodes.reviewOverlay = h('div', 'fh-review-overlay');
    nodes.reviewOverlay.hidden = true;
    nodes.reviewOverlay.addEventListener('click', function (e) {
      if (e.target === nodes.reviewOverlay) setReviewOpen(false);
    });
    var reviewShell = h('aside', 'fh-review-shell');
    var reviewHead = h('div', 'fh-review-head');
    var reviewTitle = h('div');
    reviewTitle.appendChild(h('h3', null, 'Kết quả & Workbook'));
    reviewTitle.appendChild(h('p', null, 'Xem preview, sheet hiện tại và kết quả mô phỏng khi cần đối chiếu.'));
    reviewHead.appendChild(reviewTitle);
    reviewHead.appendChild(button('×', 'fh-review-close', function () { setReviewOpen(false); }));
    reviewShell.appendChild(reviewHead);
    nodes.side = h('div', 'fh-review-content');
    reviewShell.appendChild(nodes.side);
    nodes.reviewOverlay.appendChild(reviewShell);
    root.appendChild(nodes.reviewOverlay);

    nodes.libraryDemoOverlay = h('div', 'fh-review-overlay fh-library-demo-overlay');
    nodes.libraryDemoOverlay.hidden = true;
    nodes.libraryDemoOverlay.addEventListener('click', function (e) {
      if (e.target === nodes.libraryDemoOverlay) closeLibraryDemo();
    });
    var demoShell = h('aside', 'fh-review-shell fh-library-demo-shell');
    var demoHead = h('div', 'fh-review-head');
    nodes.libraryDemoTitle = h('div');
    demoHead.appendChild(nodes.libraryDemoTitle);
    demoHead.appendChild(button('×', 'fh-review-close', closeLibraryDemo));
    demoShell.appendChild(demoHead);
    nodes.libraryDemoBody = h('div', 'fh-review-content fh-library-demo-content');
    demoShell.appendChild(nodes.libraryDemoBody);
    nodes.libraryDemoOverlay.appendChild(demoShell);
    root.appendChild(nodes.libraryDemoOverlay);

    renderPanel();
    renderSide();
  }

  function setSidebarCollapsed(collapsed) {
    state.sidebarCollapsed = !!collapsed;
    syncFormulaLayout();
  }

  function syncFormulaLayout() {
    var inTool = state.view === 'tool';
    var wasGridOpen = state.gridOpen;
    state.gridOpen = inTool;
    if (nodes.workArea) {
      nodes.workArea.classList.toggle('is-tool', inTool);
      nodes.workArea.classList.toggle('is-home', !inTool);
    }
    if (nodes.toolDock) {
      nodes.toolDock.classList.toggle('is-collapsed', inTool && state.sidebarCollapsed);
      nodes.toolDock.style.setProperty('--fh-toolbar-width', state.sidebarWidth + 'px');
    }
    if (nodes.sidebarToggle) {
      nodes.sidebarToggle.hidden = !inTool;
      nodes.sidebarToggle.textContent = state.sidebarCollapsed ? 'Hiện thanh công cụ' : 'Ẩn thanh công cụ';
      nodes.sidebarToggle.classList.toggle('is-on', inTool && !state.sidebarCollapsed);
    }
    if (wasGridOpen !== state.gridOpen || (state.gridOpen && !dataGrid) || (!state.gridOpen && dataGrid)) {
      renderDataGridPanel();
    } else if (nodes.dataGridPanel) {
      nodes.dataGridPanel.hidden = !state.gridOpen;
    }
  }

  function startSidebarResize(e) {
    if (state.view !== 'tool' || state.sidebarCollapsed || !nodes.toolDock) return;
    e.preventDefault();
    var startX = e.clientX;
    var startWidth = nodes.toolDock.getBoundingClientRect().width || state.sidebarWidth;
    var onMove = function (ev) {
      var next = startWidth - (ev.clientX - startX);
      state.sidebarWidth = Math.max(320, Math.min(680, Math.round(next)));
      nodes.toolDock.style.setProperty('--fh-toolbar-width', state.sidebarWidth + 'px');
    };
    var onUp = function () {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      document.body.classList.remove('is-resizing-formula');
    };
    document.body.classList.add('is-resizing-formula');
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }

  function setReviewOpen(open) {
    state.reviewOpen = !!open;
    if (nodes.reviewOverlay) nodes.reviewOverlay.hidden = !state.reviewOpen;
    if (nodes.resultToggle) {
      nodes.resultToggle.textContent = state.reviewOpen ? 'Đang xem kết quả' : 'Xem kết quả';
      nodes.resultToggle.classList.toggle('is-on', state.reviewOpen);
    }
    if (state.reviewOpen) renderSide();
  }

  function selectControl(label, options, value, onChange) {
    var wrap = h('div', 'fh-field fh-field-inline');
    var id = 'fh_' + label.replace(/\W+/g, '');
    var lab = h('label', null, label);
    lab.setAttribute('for', id);
    var sel = h('select');
    sel.id = id;
    options.forEach(function (o) {
      var opt = h('option', null, o.label);
      opt.value = o.value;
      if (String(o.value) === String(value)) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', function () { onChange(sel.value); });
    wrap.appendChild(lab);
    wrap.appendChild(sel);
    return wrap;
  }

  function numberControl(label, value, onChange) {
    var wrap = h('div', 'fh-field fh-field-inline fh-field-num');
    var id = 'fh_' + label.replace(/\W+/g, '');
    var lab = h('label', null, label);
    lab.setAttribute('for', id);
    var input = document.createElement('input');
    input.type = 'number';
    input.id = id;
    input.min = '1';
    input.value = value;
    input.addEventListener('input', function () { onChange(parseInt(input.value, 10)); });
    wrap.appendChild(lab);
    wrap.appendChild(input);
    return wrap;
  }

  /* ============================================================
   * Lưới dữ liệu trong Formula Helper
   * ========================================================= */

  function renderDataGridPanel() {
    if (!nodes.dataGridPanel) return;
    if (dataGrid && !dataGridSyncing) commitDataGridToWorkbook();
    dataGrid = null;
    dataGridSheetIndex = -1;
    clear(nodes.dataGridPanel);
    nodes.dataGridPanel.hidden = !state.gridOpen;
    if (!state.gridOpen) return;

    var head = h('div', 'fh-data-grid-head');
    var title = h('div', 'fh-data-grid-title');
    title.appendChild(h('h3', null, 'Lưới dữ liệu Formula Helper'));
    title.appendChild(h('p', null, 'Sửa, dán, fill, copy và xoá dữ liệu trực tiếp trên workbook riêng của Formula Helper.'));
    head.appendChild(title);
    head.appendChild(button('Nhập từ Excel', 'btn btn-tiny', function () { host.onImport(); }));
    nodes.dataGridPanel.appendChild(head);

    if (!WB.getSheets().length) {
      var empty = h('div', 'fh-data-grid-empty');
      empty.appendChild(h('p', null, 'Chưa có Workbook. Nhập file Excel hoặc tạo dữ liệu trên lưới Formula Helper.'));
      nodes.dataGridPanel.appendChild(empty);
      return;
    }

    var meta = h('div', 'fh-data-grid-meta');
    var active = WB.getActiveIndex();
    meta.appendChild(h('strong', null, WB.getSheetName(active) || 'Sheet'));
    meta.appendChild(h('span', null, WB.getHeaders(active).length + ' cột · ' + WB.getDataRowCount(active) + ' dòng dữ liệu'));
    nodes.dataGridPanel.appendChild(meta);

    nodes.dataGridHost = h('div', 'fh-data-grid-host grid-scroll');
    nodes.dataGridPanel.appendChild(nodes.dataGridHost);
    ensureDataGrid();
    syncDataGridFromWorkbook();
  }

  function ensureDataGrid() {
    if (dataGrid || !global.Grid || !nodes.dataGridHost) return;
    dataGrid = new global.Grid(nodes.dataGridHost, {
      onChange: commitDataGridToWorkbook,
      onStatus: host.toast,
      onHistoryChange: function () {},
      onColumnTypesChange: commitDataGridToWorkbook,
      mode: 'insert'
    });
  }

  function syncDataGridFromWorkbook() {
    if (!dataGrid || !state.gridOpen || !WB.getSheets().length) return;
    var index = WB.getActiveIndex();
    var sheet = WB.getSheet(index);
    if (!sheet) return;
    dataGridSyncing = true;
    dataGridSheetIndex = index;
    dataGrid.setColumnTypes(sheet.columnTypes || []);
    dataGrid.setMode('insert', 0);
    dataGrid.setData(sheet.data || [[]]);
    setTimeout(function () { dataGridSyncing = false; }, 0);
  }

  function commitDataGridToWorkbook() {
    if (dataGridSyncing || !dataGrid || dataGridSheetIndex < 0) return;
    var sheet = WB.getSheet(dataGridSheetIndex);
    if (!sheet) return;
    var oldHeader = ((sheet.data && sheet.data[0]) || []).join('\u001f');
    sheet.data = dataGrid.getData();
    sheet.columnTypes = dataGrid.getColumnTypes();
    var newHeader = ((sheet.data && sheet.data[0]) || []).join('\u001f');
    host.onWorkbookEdit(oldHeader === newHeader ? 'data' : 'structure');
  }

  /* ============================================================
   * Panel: Home / Tool / Library
   * ========================================================= */

  function renderPanel() {
    if (!nodes.panel) return;
    syncFormulaLayout();
    clear(nodes.panel);
    if (state.view === 'tool') renderTool();
    else if (state.view === 'library') renderLibrary();
    else renderHome();
  }

  function renderHome() {
    var p = nodes.panel;

    var head = h('div', 'fh-head');
    var title = h('div', 'fh-head-text');
    title.appendChild(h('h2', null, 'Bạn muốn làm gì?'));
    title.appendChild(h('p', null, 'Chọn nhu cầu, Formula Helper sẽ tự chọn công thức phù hợp.'));
    head.appendChild(title);
    p.appendChild(head);

    var searchWrap = h('div', 'fh-search');
    var search = document.createElement('input');
    search.type = 'search';
    search.placeholder = 'Tìm công cụ hoặc tên hàm… (trùng, xlookup, ngày)';
    search.value = state.query;
    search.setAttribute('aria-label', 'Tìm công cụ hoặc công thức');
    search.addEventListener('input', function () {
      state.query = search.value;
      renderToolList(listHost);
    });
    searchWrap.appendChild(search);
    p.appendChild(searchWrap);

    // Yêu thích §108
    var favorites = STORE.getFavorites();
    if (favorites.length) {
      var favSec = h('section', 'fh-section');
      favSec.appendChild(h('h3', 'fh-section-title', 'Yêu thích'));
      var favList = h('div', 'fh-fav-list');
      favorites.forEach(function (fav) {
        var chip = h('div', 'fh-fav');
        var open = button('★ ' + fav.name, 'fh-fav-open', function () { loadSaved(fav); });
        open.title = fav.subtitle || fav.formula;
        var del = button('×', 'fh-fav-del', function () {
          STORE.removeFavorite(fav.id);
          renderPanel();
        });
        del.title = 'Bỏ yêu thích';
        del.setAttribute('aria-label', 'Bỏ yêu thích ' + fav.name);
        chip.appendChild(open);
        chip.appendChild(del);
        favList.appendChild(chip);
      });
      favSec.appendChild(favList);
      p.appendChild(favSec);
    }

    var allSec = h('section', 'fh-section');
    allSec.appendChild(h('h3', 'fh-section-title', 'Tất cả công cụ'));
    var listHost = h('div', 'fh-cards');
    allSec.appendChild(listHost);
    p.appendChild(allSec);
    renderToolList(listHost);

    renderHistorySection(p);
  }

  function renderToolList(hostNode) {
    clear(hostNode);
    var platform = state.settings.platform;
    var found = PRESETS.search(state.query).filter(function (t) {
      return t.platforms.indexOf(platform) !== -1;
    });

    if (!found.length) {
      var empty = h('p', 'fh-empty', 'Không tìm thấy công cụ phù hợp. Thử từ khoá khác hoặc mở Thư viện hàm.');
      hostNode.appendChild(empty);
      return;
    }

    found.forEach(function (t) {
      var card = h('button', 'fh-card');
      card.type = 'button';
      card.appendChild(svgIcon(t.icon));
      var body = h('span', 'fh-card-body');
      body.appendChild(h('strong', null, t.name));
      body.appendChild(h('span', 'fh-card-desc', t.desc));
      body.appendChild(h('span', 'fh-card-tag', PRESETS.CATEGORIES[t.category] || t.category));
      card.appendChild(body);
      card.addEventListener('click', function () { openTool(t.id); });
      hostNode.appendChild(card);
    });
  }

  function renderHistorySection(p) {
    var history = STORE.getHistory();
    if (!history.length) return;
    var sec = h('section', 'fh-section');
    var head = h('div', 'fh-section-head');
    head.appendChild(h('h3', 'fh-section-title', 'Lịch sử'));
    head.appendChild(button('Xoá lịch sử', 'btn btn-tiny btn-danger', function () {
      STORE.clearHistory();
      renderPanel();
    }));
    sec.appendChild(head);

    var list = h('ul', 'fh-history');
    history.slice(0, 12).forEach(function (item) {
      var li = h('li');
      var open = h('button', 'fh-history-item');
      open.type = 'button';
      open.appendChild(h('span', 'fh-history-time', formatTime(item.timestamp)));
      var text = h('span', 'fh-history-text');
      text.appendChild(h('strong', null, item.title || item.type));
      if (item.subtitle) text.appendChild(h('em', null, item.subtitle));
      text.appendChild(h('code', null, item.formula));
      open.appendChild(text);
      open.addEventListener('click', function () { loadSaved(item); });
      li.appendChild(open);
      sec.appendChild(list);
      list.appendChild(li);
    });
    p.appendChild(sec);
  }

  function formatTime(ts) {
    var d = new Date(ts);
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    var today = new Date();
    var sameDay = d.toDateString() === today.toDateString();
    return sameDay
      ? pad(d.getHours()) + ':' + pad(d.getMinutes())
      : pad(d.getDate()) + '/' + pad(d.getMonth() + 1);
  }

  /** Nạp lại toàn bộ cấu hình từ History/Favorite (§102 §121 §122). */
  function loadSaved(item) {
    if (!item || !item.config) {
      host.toast('Mục này không còn cấu hình để nạp lại.');
      return;
    }
    var t = PRESETS.getTool(item.config.__toolId || item.type);
    if (!t) { host.toast('Công cụ không còn tồn tại.'); return; }

    state.toolId = t.id;
    state.config = JSON.parse(JSON.stringify(item.config));
    delete state.config.__toolId;
    state.view = 'tool';
    state.sidebarCollapsed = false;

    if (item.platform) saveSettings({ platform: item.platform });

    var issues = validateRefs(state.config);
    renderPanel();
    if (issues.length) host.toast(issues[0]);
  }

  /** Kiểm tra các tham chiếu còn hợp lệ sau khi workbook đổi. §121 §122 */
  function validateRefs(cfg) {
    var issues = [];
    var walk = function (value) {
      if (!value || typeof value !== 'object') return;
      if (value.mode === 'column') {
        refreshRef(value);
        var check = WB.validateField(value);
        if (!check.ok && issues.indexOf(check.message) === -1) issues.push(check.message);
        return;
      }
      if (Array.isArray(value)) { value.forEach(walk); return; }
      Object.keys(value).forEach(function (k) { walk(value[k]); });
    };
    walk(cfg);
    return issues;
  }

  function openTool(id) {
    var t = PRESETS.getTool(id);
    if (!t) return;
    state.toolId = id;
    state.view = 'tool';
    state.sidebarCollapsed = false;
    state.config = {};
    Object.keys(t.defaults || {}).forEach(function (k) { state.config[k] = t.defaults[k]; });
    applySmartDefaults(t);
    renderPanel();
  }

  /** Dùng cột đang chọn trên Grid làm giá trị mặc định (§28 §150 §151). */
  function applySmartDefaults(t) {
    var sel = selectedFormulaGridColumn();
    if (!sel) return;
    var ref = makeRef(sel.sheetIndex, sel.colIndex);
    (t.schema || []).forEach(function (field) {
      if (!field.useSelected) return;
      if (field.type === 'field' && !state.config[field.key]) state.config[field.key] = ref;
      if (field.type === 'fields' && !(state.config[field.key] || []).length) state.config[field.key] = [ref];
    });
    if (t.id === 'filter' && !state.config.sourceSheet) {
      state.config.sourceSheet = WB.getSheetName(sel.sheetIndex);
    }
  }

  /* ============================================================
   * Tool builder
   * ========================================================= */

  function renderTool() {
    var t = tool();
    if (!t) { state.view = 'home'; renderPanel(); return; }
    var p = nodes.panel;

    var head = h('div', 'fh-head');
    head.appendChild(button('‹ Quay lại', 'btn btn-tiny fh-back', function () {
      state.view = 'home';
      state.toolId = null;
      state.query = '';
      renderPanel();
    }));
    var title = h('div', 'fh-head-text');
    var titleRow = h('h2');
    titleRow.appendChild(document.createTextNode(t.name));
    var info = h('button', 'fh-info');
    info.type = 'button';
    info.textContent = 'i';
    info.title = t.primaryFunction + ' — ' + t.desc;
    info.setAttribute('aria-label', 'Thông tin về ' + t.primaryFunction);
    info.addEventListener('click', function () {
      host.toast(t.primaryFunction + ': ' + libraryNote(t.primaryFunction));
    });
    titleRow.appendChild(info);
    title.appendChild(titleRow);
    title.appendChild(h('p', null, t.desc));
    head.appendChild(title);

    var headActions = h('div', 'fh-head-actions');
    headActions.appendChild(button('Đặt lại', 'btn btn-tiny', function () { openTool(t.id); }));
    head.appendChild(headActions);
    p.appendChild(head);

    var form = h('div', 'fh-form');
    (t.schema || []).forEach(function (field) {
      if (field.platform && field.platform !== state.settings.platform) return;
      if (field.advanced && !state.showAdvanced) return;
      if (!matchShowIf(field)) return;
      var node = renderField(field, t);
      if (node) form.appendChild(node);
    });
    p.appendChild(form);

    if (state.showAdvanced) p.appendChild(renderAdvancedSettings());

    nodes.output = h('div', 'fh-output');
    p.appendChild(nodes.output);
    updateOutput();
  }

  function libraryNote(name) {
    for (var i = 0; i < PRESETS.LIBRARY.length; i++) {
      if (PRESETS.LIBRARY[i].name === name) return PRESETS.LIBRARY[i].use;
    }
    return 'Hàm được dùng cho công cụ này.';
  }

  function matchShowIf(field) {
    if (!field.showIf) return true;
    var ok = true;
    Object.keys(field.showIf).forEach(function (k) {
      var expect = field.showIf[k];
      var actual = state.config[k];
      if (Array.isArray(expect)) { if (expect.indexOf(actual) === -1) ok = false; }
      else if (String(expect) !== String(actual)) ok = false;
    });
    return ok;
  }

  function renderAdvancedSettings() {
    var sec = h('section', 'fh-advanced');
    sec.appendChild(h('h3', 'fh-section-title', 'Tuỳ chọn nâng cao'));
    var grid = h('div', 'fh-advanced-grid');

    grid.appendChild(selectControl('Phiên bản Excel', [
      { value: 'modern', label: 'Microsoft 365 / 2021' },
      { value: 'legacy', label: 'Excel đời cũ' }
    ], state.settings.excelVersion, function (v) {
      saveSettings({ excelVersion: v });
      updateOutput();
    }));

    grid.appendChild(selectControl('Vùng tham chiếu', [
      { value: 'full', label: 'Toàn cột  A:A' },
      { value: 'used', label: 'Vùng dữ liệu  A2:A1000' }
    ], state.settings.rangeMode, function (v) {
      saveSettings({ rangeMode: v });
      updateOutput();
    }));

    grid.appendChild(selectControl('Kiểu tham chiếu ô', [
      { value: 'relative', label: 'Tương đối  A2' },
      { value: 'absoluteColumn', label: 'Khoá cột  $A2' },
      { value: 'absoluteCell', label: 'Khoá ô  $A$2' }
    ], state.settings.referenceType, function (v) {
      saveSettings({ referenceType: v });
      updateOutput();
    }));

    sec.appendChild(grid);
    return sec;
  }

  /* -------------------------------------------------- field renderers */

  function fieldWrap(field) {
    var wrap = h('div', 'fh-field');
    var lab = h('label', null, field.label);
    if (field.required) lab.appendChild(h('span', 'fh-req', ' *'));
    if (field.technical) {
      var info = h('button', 'fh-info fh-info-sm');
      info.type = 'button';
      info.textContent = 'i';
      info.title = field.technical + ' trong ' + (tool() ? tool().primaryFunction : 'công thức');
      info.setAttribute('aria-label', 'Tên kỹ thuật: ' + field.technical);
      info.addEventListener('click', function () {
        host.toast(field.label + ' tương ứng với ' + field.technical + '.');
      });
      lab.appendChild(info);
    }
    wrap.appendChild(lab);
    return wrap;
  }

  function renderField(field, t) {
    switch (field.type) {
      case 'field': return renderRefField(field);
      case 'fields': return renderRefListField(field);
      case 'text': return renderTextField(field);
      case 'number': return renderNumberField(field);
      case 'select': return renderSelectField(field);
      case 'checkbox': return renderCheckboxField(field);
      case 'sheet': return renderSheetField(field);
      case 'conditions': return renderConditionsField(field);
      case 'branches': return renderBranchesField(field);
      default: return null;
    }
  }

  function refSelector(current, onChange, opts) {
    opts = opts || {};
    var box = h('div', 'fh-ref');
    var mode = (current && current.mode) || 'column';

    var modeSel = h('select', 'fh-ref-mode');
    [
      { value: 'column', label: 'Cột' },
      { value: 'cell', label: 'Ô / Vùng' },
      { value: 'custom', label: 'Tuỳ ý' }
    ].forEach(function (o) {
      var opt = h('option', null, o.label);
      opt.value = o.value;
      if (o.value === mode) opt.selected = true;
      modeSel.appendChild(opt);
    });
    modeSel.setAttribute('aria-label', 'Kiểu tham chiếu');
    modeSel.addEventListener('change', function () {
      onChange(modeSel.value === 'column'
        ? null
        : { mode: modeSel.value, raw: (current && current.raw) || '' });
      renderPanel();
    });
    box.appendChild(modeSel);

    if (mode === 'column') {
      var sel = h('select', 'fh-ref-select');
      var placeholder = h('option', null, '— Chọn cột —');
      placeholder.value = '';
      sel.appendChild(placeholder);

      var fields = WB.getAllFields();
      var groups = {};
      fields.forEach(function (f) {
        if (!groups[f.sheetName]) {
          groups[f.sheetName] = h('optgroup');
          groups[f.sheetName].label = f.sheetName;
          sel.appendChild(groups[f.sheetName]);
        }
        var opt = h('option', null, f.label + '  ·  Cột ' + f.letter);
        opt.value = f.value;
        if (current && current.mode === 'column' &&
            current.sheetIndex === f.sheetIndex && current.colIndex === f.colIndex) {
          opt.selected = true;
        }
        groups[f.sheetName].appendChild(opt);
      });

      if (!fields.length) {
        sel.disabled = true;
        placeholder.textContent = '— Chưa có dữ liệu, dùng chế độ Ô / Tuỳ ý —';
      }

      sel.addEventListener('change', function () {
        var parsed = WB.parseFieldValue(sel.value);
        onChange(parsed ? makeRef(parsed.sheetIndex, parsed.colIndex) : null);
        updateOutput();
      });
      box.appendChild(sel);

      if (opts.useSelected) {
        var useBtn = button('Dùng cột đang chọn', 'btn btn-tiny fh-use-selected', function () {
          var s = selectedFormulaGridColumn();
          if (!s) { host.toast('Chưa chọn cột nào trên lưới dữ liệu.'); return; }
          onChange(makeRef(s.sheetIndex, s.colIndex));
          renderPanel();
        });
        box.appendChild(useBtn);
      }
    } else {
      var input = document.createElement('input');
      input.type = 'text';
      input.className = 'fh-ref-input';
      input.placeholder = mode === 'cell' ? 'A2  ·  $A$2  ·  A2:A100' : "Database!B:B  ·  'CONFIG - Process'!A:A";
      input.value = (current && current.raw) || '';
      input.addEventListener('input', function () {
        onChange({ mode: mode, raw: input.value });
        updateOutput();
      });
      box.appendChild(input);
    }

    if (current && current.mode === 'column') {
      var hint = h('span', 'fh-ref-hint',
        current.sheetName + '.' + current.column + '  ·  Cột ' + ENGINE.columnLetter(current.colIndex));
      box.appendChild(hint);
    }
    return box;
  }

  function renderRefField(field) {
    var wrap = fieldWrap(field);
    wrap.appendChild(refSelector(state.config[field.key], function (ref) {
      state.config[field.key] = ref;
    }, { useSelected: field.useSelected }));
    return wrap;
  }

  function renderRefListField(field) {
    var wrap = fieldWrap(field);
    var list = state.config[field.key] || [];
    if (!list.length) list = [null];

    list.forEach(function (ref, index) {
      var row = h('div', 'fh-row');
      row.appendChild(refSelector(ref, function (next) {
        var arr = (state.config[field.key] || []).slice();
        while (arr.length <= index) arr.push(null);
        arr[index] = next;
        state.config[field.key] = arr;
      }, { useSelected: field.useSelected && index === 0 }));

      if (list.length > (field.min || 1)) {
        var del = button('×', 'fh-row-del', function () {
          var arr = (state.config[field.key] || []).slice();
          arr.splice(index, 1);
          state.config[field.key] = arr;
          renderPanel();
        });
        del.setAttribute('aria-label', 'Xoá dòng này');
        row.appendChild(del);
      }
      wrap.appendChild(row);
    });

    wrap.appendChild(button('+ Thêm cột', 'btn btn-tiny fh-add', function () {
      var arr = (state.config[field.key] || []).slice();
      arr.push(null);
      state.config[field.key] = arr;
      renderPanel();
    }));
    return wrap;
  }

  function renderTextField(field) {
    var wrap = fieldWrap(field);
    var input = document.createElement('input');
    input.type = 'text';
    input.placeholder = field.placeholder || '';
    input.value = state.config[field.key] == null ? '' : state.config[field.key];
    input.addEventListener('input', function () {
      state.config[field.key] = input.value;
      updateOutput();
    });
    wrap.appendChild(input);
    return wrap;
  }

  function renderNumberField(field) {
    var wrap = fieldWrap(field);
    var input = document.createElement('input');
    input.type = 'number';
    input.value = state.config[field.key] == null ? '' : state.config[field.key];
    input.addEventListener('input', function () {
      state.config[field.key] = input.value === '' ? '' : parseInt(input.value, 10);
      updateOutput();
    });
    wrap.appendChild(input);
    return wrap;
  }

  function renderSelectField(field) {
    var wrap = fieldWrap(field);
    var sel = h('select');
    field.options.forEach(function (o) {
      var opt = h('option', null, o.label);
      opt.value = o.value;
      if (String(o.value) === String(state.config[field.key])) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', function () {
      state.config[field.key] = sel.value;
      renderPanel();
    });
    wrap.appendChild(sel);
    return wrap;
  }

  function renderCheckboxField(field) {
    var wrap = h('div', 'fh-field fh-field-check');
    var lab = h('label', 'fh-check');
    var input = document.createElement('input');
    input.type = 'checkbox';
    input.checked = !!state.config[field.key];
    input.addEventListener('change', function () {
      state.config[field.key] = input.checked;
      updateOutput();
    });
    lab.appendChild(input);
    lab.appendChild(h('span', null, field.label));
    wrap.appendChild(lab);
    return wrap;
  }

  function renderSheetField(field) {
    var wrap = fieldWrap(field);
    var sel = h('select');
    var placeholder = h('option', null, '— Chọn Sheet —');
    placeholder.value = '';
    sel.appendChild(placeholder);
    WB.getSheets().forEach(function (s, i) {
      var opt = h('option', null, s.name);
      opt.value = s.name;
      if (s.name === state.config[field.key]) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', function () {
      state.config[field.key] = sel.value;
      updateOutput();
    });
    wrap.appendChild(sel);
    return wrap;
  }

  function operatorSelect(value, onChange) {
    var sel = h('select', 'fh-op');
    PRESETS.OPERATOR_OPTIONS.forEach(function (o) {
      var opt = h('option', null, o.label);
      opt.value = o.value;
      if (o.value === value) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.setAttribute('aria-label', 'Phép so sánh');
    sel.addEventListener('change', function () { onChange(sel.value); });
    return sel;
  }

  function conditionRow(cond, onChange, onDelete, opts) {
    opts = opts || {};
    var row = h('div', 'fh-cond');

    if (!opts.valueOnly) {
      row.appendChild(refSelector(cond.field, function (ref) {
        cond.field = ref;
        onChange();
      }, {}));
    } else {
      row.appendChild(refSelector(cond.field, function (ref) {
        cond.field = ref;
        onChange();
      }, {}));
    }

    row.appendChild(operatorSelect(cond.operator || '=', function (op) {
      cond.operator = op;
      onChange();
      renderPanel();
    }));

    if (cond.operator !== 'blank' && cond.operator !== 'notBlank') {
      var input = document.createElement('input');
      input.type = 'text';
      input.className = 'fh-cond-value';
      input.placeholder = 'Giá trị hoặc ô  (F3)';
      input.value = cond.value == null ? '' : cond.value;
      input.setAttribute('aria-label', 'Giá trị so sánh');
      input.addEventListener('input', function () {
        cond.value = input.value;
        updateOutput();
      });
      row.appendChild(input);
    }

    if (onDelete) {
      var del = button('×', 'fh-row-del', onDelete);
      del.setAttribute('aria-label', 'Xoá điều kiện');
      row.appendChild(del);
    }
    return row;
  }

  function renderConditionsField(field) {
    var wrap = fieldWrap(field);
    var list = state.config[field.key] || [];

    list.forEach(function (cond, i) {
      wrap.appendChild(conditionRow(cond, function () { updateOutput(); }, function () {
        list.splice(i, 1);
        state.config[field.key] = list;
        renderPanel();
      }, { valueOnly: field.valueOnly }));
    });

    wrap.appendChild(button('+ Thêm điều kiện', 'btn btn-tiny fh-add', function () {
      list.push({ field: null, operator: '=', value: '' });
      state.config[field.key] = list;
      renderPanel();
    }));
    return wrap;
  }

  /* -------------------------------------------------- điều kiện lồng §55 §159 */

  function renderGroup(group, onChange, depth) {
    var box = h('div', 'fh-group' + (depth ? ' fh-group-nested' : ''));

    var head = h('div', 'fh-group-head');
    var opSel = h('select', 'fh-group-op');
    [{ value: 'AND', label: 'Thoả TẤT CẢ (AND)' }, { value: 'OR', label: 'Thoả BẤT KỲ (OR)' }].forEach(function (o) {
      var opt = h('option', null, o.label);
      opt.value = o.value;
      if (o.value === group.operator) opt.selected = true;
      opSel.appendChild(opt);
    });
    opSel.setAttribute('aria-label', 'Cách nối các điều kiện trong nhóm');
    opSel.addEventListener('change', function () {
      group.operator = opSel.value;
      onChange();
      renderPanel();
    });
    head.appendChild(opSel);
    box.appendChild(head);

    (group.children || []).forEach(function (child, i) {
      var remove = function () {
        group.children.splice(i, 1);
        onChange();
        renderPanel();
      };
      if (child.kind === 'group') {
        var nested = renderGroup(child, onChange, (depth || 0) + 1);
        var del = button('× Xoá nhóm', 'btn btn-tiny fh-row-del', remove);
        nested.appendChild(del);
        box.appendChild(nested);
      } else {
        box.appendChild(conditionRow(child, onChange, remove, {}));
      }
    });

    var actions = h('div', 'fh-group-actions');
    actions.appendChild(button('+ Điều kiện', 'btn btn-tiny', function () {
      group.children = group.children || [];
      group.children.push({ kind: 'cond', field: null, operator: '=', value: '' });
      onChange();
      renderPanel();
    }));
    if ((depth || 0) < 2) {
      actions.appendChild(button('+ Nhóm con', 'btn btn-tiny', function () {
        group.children = group.children || [];
        group.children.push({ kind: 'group', operator: 'OR', children: [] });
        onChange();
        renderPanel();
      }));
    }
    box.appendChild(actions);
    return box;
  }

  function renderBranchesField(field) {
    var wrap = fieldWrap(field);
    var branches = state.config[field.key] || [];
    if (!branches.length) {
      branches.push({ condition: { kind: 'group', operator: 'AND', children: [{ kind: 'cond', field: null, operator: '=', value: '' }] }, result: '' });
      state.config[field.key] = branches;
    }

    branches.forEach(function (branch, i) {
      var box = h('div', 'fh-branch');
      var head = h('div', 'fh-branch-head');
      head.appendChild(h('strong', null, i === 0 ? 'NẾU' : 'NGƯỢC LẠI, NẾU'));
      if (branches.length > 1) {
        var del = button('×', 'fh-row-del', function () {
          branches.splice(i, 1);
          renderPanel();
        });
        del.setAttribute('aria-label', 'Xoá nhánh');
        head.appendChild(del);
      }
      box.appendChild(head);
      box.appendChild(renderGroup(branch.condition, function () { updateOutput(); }, 0));

      var resWrap = h('div', 'fh-field');
      resWrap.appendChild(h('label', null, 'thì ghi'));
      var input = document.createElement('input');
      input.type = 'text';
      input.value = branch.result == null ? '' : branch.result;
      input.addEventListener('input', function () {
        branch.result = input.value;
        updateOutput();
      });
      resWrap.appendChild(input);
      box.appendChild(resWrap);
      wrap.appendChild(box);
    });

    wrap.appendChild(button('+ Thêm nhánh điều kiện', 'btn btn-tiny fh-add', function () {
      branches.push({ condition: { kind: 'group', operator: 'AND', children: [{ kind: 'cond', field: null, operator: '=', value: '' }] }, result: '' });
      renderPanel();
    }));
    return wrap;
  }

  /* ============================================================
   * Formula preview / copy / explanation
   * ========================================================= */

  function updateOutput() {
    if (!nodes.output) return;
    var t = tool();
    if (!t) return;

    var cfg = buildConfig();
    var result = ENGINE.generate(cfg);
    state.lastResult = result;
    state.lastConfig = cfg;

    clear(nodes.output);

    if (result.error) {
      var err = h('div', 'fh-inline-error');
      err.appendChild(h('strong', null, 'Chưa sinh được công thức'));
      err.appendChild(h('p', null, result.error));
      nodes.output.appendChild(err);
      renderPreviewTable(null);
      return;
    }

    // ---- khối công thức
    var block = h('section', 'fh-formula');
    var head = h('div', 'fh-formula-head');
    head.appendChild(h('h3', null, 'Công thức'));

    var formatToggle = h('div', 'fh-format');
    [{ value: 'single', label: 'Một dòng' }, { value: 'formatted', label: 'Xuống dòng' }].forEach(function (o) {
      var b = button(o.label, 'fh-format-btn' + (state.settings.formatView === o.value ? ' is-on' : ''), function () {
        saveSettings({ formatView: o.value });
        updateOutput();
      });
      b.setAttribute('aria-pressed', state.settings.formatView === o.value ? 'true' : 'false');
      formatToggle.appendChild(b);
    });
    head.appendChild(formatToggle);
    block.appendChild(head);

    var pre = h('pre', 'fh-formula-code');
    pre.tabIndex = 0;
    pre.textContent = state.settings.formatView === 'formatted' ? result.formattedFormula : result.formula;
    pre.title = 'Bấm để sao chép công thức';
    pre.addEventListener('click', function () { copyFormula(result); });
    pre.addEventListener('keydown', function (e) {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
        e.preventDefault();
        copyFormula(result);
      }
    });
    block.appendChild(pre);

    var actions = h('div', 'fh-formula-actions');
    var copyBtn = button('Sao chép công thức', 'btn btn-primary fh-copy', function () { copyFormula(result); });
    actions.appendChild(copyBtn);
    actions.appendChild(button('★ Lưu yêu thích', 'btn btn-tiny', function () { saveFavorite(result); }));
    if (result.separator) {
      actions.appendChild(h('span', 'fh-sep-note',
        'Dấu phân cách đang dùng: ' + (result.separator === ',' ? 'dấu phẩy' : 'chấm phẩy')));
    }
    block.appendChild(actions);
    nodes.output.appendChild(block);

    // ---- cảnh báo §92 §162
    if (result.warnings && result.warnings.length) {
      var warnBox = h('div', 'fh-warnings');
      warnBox.appendChild(h('h4', null, 'Lưu ý'));
      result.warnings.forEach(function (w) { warnBox.appendChild(h('p', null, w)); });
      nodes.output.appendChild(warnBox);
    }

    // ---- giải thích §89 §161
    if (result.explanation && result.explanation.length) {
      var exp = h('details', 'fh-explain');
      exp.open = true;
      exp.appendChild(h('summary', null, 'Giải thích'));
      var dl = h('dl');
      result.explanation.forEach(function (item) {
        dl.appendChild(h('dt', null, item.code));
        dl.appendChild(h('dd', null, item.text));
      });
      exp.appendChild(dl);
      nodes.output.appendChild(exp);
    }

    schedulePreview(cfg);
    pushHistoryDebounced(t, cfg, result);
  }

  var historyTimer = null;
  function pushHistoryDebounced(t, cfg, result) {
    if (historyTimer) clearTimeout(historyTimer);
    historyTimer = setTimeout(function () {
      var stored = JSON.parse(JSON.stringify(state.config));
      stored.__toolId = t.id;
      STORE.pushHistory({
        type: t.id,
        title: t.name,
        subtitle: describeConfig(t, cfg),
        platform: state.settings.platform,
        separator: result.separator,
        formula: result.formula,
        config: stored
      });
    }, 1500);
  }

  function describeConfig(t, cfg) {
    var label = function (ref) { return ref ? ENGINE.fieldLabel(ref) : ''; };
    if (t.id === 'lookup') return label(cfg.lookupValue) + ' → ' + label(cfg.returnRange);
    if (t.id === 'duplicate' || t.id === 'blank') {
      return (cfg.fields || []).map(label).filter(Boolean).join(', ');
    }
    if (t.id === 'compare') return label(cfg.valueA) + ' ↔ ' + label(cfg.valueB);
    if (t.id === 'existence' || t.id === 'listCompare') return label(cfg.value) + ' ∈ ' + label(cfg.reference);
    if (t.id === 'filter') return cfg.sourceSheet || '';
    return label(cfg.field) || '';
  }

  function copyFormula(result) {
    // Mặc định copy bản một dòng (§88)
    var text = result.formula;
    var done = function () { host.toast('Đã sao chép công thức.'); markCopied(); };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done, fallback);
    } else fallback();

    function fallback() {
      var ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); done(); }
      catch (err) { host.toast('Trình duyệt chặn thao tác sao chép.'); }
      document.body.removeChild(ta);
    }
  }

  function markCopied() {
    var btn = nodes.output && nodes.output.querySelector('.fh-copy');
    if (!btn) return;
    var original = btn.textContent;
    btn.textContent = '✓ Đã sao chép';
    btn.classList.add('is-copied');
    setTimeout(function () {
      btn.textContent = original;
      btn.classList.remove('is-copied');
    }, 1400);
  }

  function saveFavorite(result) {
    var t = tool();
    if (!t || !result || !result.formula) return;
    var suggested = t.name + (describeConfig(t, state.lastConfig) ? ' — ' + describeConfig(t, state.lastConfig) : '');
    var name = prompt('Đặt tên cho mục yêu thích:', suggested);
    if (name == null) return;
    var stored = JSON.parse(JSON.stringify(state.config));
    stored.__toolId = t.id;
    STORE.addFavorite({
      name: name.trim() || suggested,
      type: t.id,
      platform: state.settings.platform,
      separator: result.separator,
      formula: result.formula,
      subtitle: describeConfig(t, state.lastConfig),
      config: stored
    });
    host.toast('Đã lưu vào yêu thích.');
  }

  /* ============================================================
   * Result preview (§95 §171 §172)
   * ========================================================= */

  function schedulePreview(cfg) {
    if (state.previewTimer) clearTimeout(state.previewTimer);
    state.previewTimer = setTimeout(function () {
      renderPreviewTable(EXPLAINER.preview(cfg, WB));
    }, 300);
  }

  function renderPreviewTable(data) {
    if (!nodes.previewHost) return;
    clear(nodes.previewHost);
    if (!data || !data.rows || !data.rows.length) {
      nodes.previewHost.appendChild(h('p', 'fh-empty',
        'Xem trước kết quả sẽ hiện ở đây khi công cụ có thể mô phỏng bằng dữ liệu hiện tại.'));
      return;
    }

    if (data.directOutput) {
      var direct = h('div', 'fh-direct');
      direct.appendChild(h('h4', null, 'Kết quả trực tiếp từ dữ liệu'));
      var code = h('pre', 'fh-direct-code');
      code.textContent = data.directOutput.whereIn;
      direct.appendChild(code);
      var acts = h('div', 'fh-direct-actions');
      acts.appendChild(button('Copy values', 'btn btn-tiny', function () {
        copyText(data.directOutput.values, 'Đã sao chép danh sách giá trị.');
      }));
      acts.appendChild(button('Copy WHERE IN', 'btn btn-tiny', function () {
        copyText(data.directOutput.whereIn, 'Đã sao chép mệnh đề WHERE IN.');
      }));
      direct.appendChild(acts);
      nodes.previewHost.appendChild(direct);
    }

    var table = h('table', 'fh-preview-table');
    var thead = h('thead');
    var hr = h('tr');
    data.columns.forEach(function (c) { hr.appendChild(h('th', null, c)); });
    thead.appendChild(hr);
    table.appendChild(thead);

    var tbody = h('tbody');
    data.rows.slice(0, EXPLAINER.PREVIEW_ROWS).forEach(function (row) {
      var tr = h('tr');
      row.forEach(function (cell, i) {
        var td = h('td', i === row.length - 1 ? 'fh-preview-result' : null,
          cell === '' ? '—' : String(cell));
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    nodes.previewHost.appendChild(table);

    if (data.note) nodes.previewHost.appendChild(h('p', 'fh-preview-note', data.note));
  }

  function copyText(text, message) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () { host.toast(message); });
    } else {
      var ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); host.toast(message); } catch (err) { /* ignore */ }
      document.body.removeChild(ta);
    }
  }

  /* ============================================================
   * Thư viện hàm (§187)
   * ========================================================= */

  function renderLibrary() {
    var p = nodes.panel;
    var head = h('div', 'fh-head');
    head.appendChild(button('‹ Quay lại', 'btn btn-tiny fh-back', function () {
      state.view = 'home';
      renderPanel();
    }));
    var title = h('div', 'fh-head-text');
    title.appendChild(h('h2', null, 'Thư viện hàm'));
    title.appendChild(h('p', null, 'Tra nhanh công dụng và công thức mẫu. Bấm vào hàm để xem mô phỏng bằng lưới Excel.'));
    head.appendChild(title);
    p.appendChild(head);

    var searchWrap = h('div', 'fh-search');
    var search = document.createElement('input');
    search.type = 'search';
    search.placeholder = 'Tìm hàm… (xlookup, countif, query)';
    search.setAttribute('aria-label', 'Tìm hàm trong thư viện');
    searchWrap.appendChild(search);
    p.appendChild(searchWrap);

    var listHost = h('div', 'fh-library');
    p.appendChild(listHost);

    var draw = function () {
      clear(listHost);
      PRESETS.searchLibrary(search.value)
        .filter(function (f) { return f.platforms.indexOf(state.settings.platform) !== -1; })
        .forEach(function (f) {
          var item = h('article', 'fh-lib-item');
          item.tabIndex = 0;
          item.setAttribute('role', 'button');
          item.setAttribute('aria-label', 'Xem mô phỏng hàm ' + f.name);
          item.appendChild(h('h4', null, f.name));
          item.appendChild(h('p', null, f.use));
          var code = h('code', null, f.example);
          item.appendChild(code);
          item.addEventListener('click', function () { openLibraryDemo(f); });
          item.addEventListener('keydown', function (e) {
            if (e.key === 'Enter' || e.key === ' ') {
              e.preventDefault();
              openLibraryDemo(f);
            }
          });
          listHost.appendChild(item);
        });
    };
    search.addEventListener('input', draw);
    draw();
  }

  function openLibraryDemo(f) {
    if (!nodes.libraryDemoOverlay || !nodes.libraryDemoTitle || !nodes.libraryDemoBody) return;
    clear(nodes.libraryDemoTitle);
    clear(nodes.libraryDemoBody);
    nodes.libraryDemoTitle.appendChild(h('h3', null, f.name + ' - mô phỏng hàm'));
    nodes.libraryDemoTitle.appendChild(h('p', null, f.use));

    var demo = f.demoSheet || null;
    var formulaBar = h('div', 'fh-demo-formula-bar');
    formulaBar.appendChild(h('span', null, 'fx'));
    formulaBar.appendChild(h('code', null, demo && demo.formula ? demo.formula : f.example));
    nodes.libraryDemoBody.appendChild(formulaBar);

    var guide = h('p', 'fh-demo-guide',
      demo && demo.guide ? demo.guide : 'Nhìn giống một sheet Excel nhỏ: cột A là dữ liệu mẫu, cột B là công thức, cột C là kết quả dự kiến.');
    nodes.libraryDemoBody.appendChild(guide);

    nodes.libraryDemoBody.appendChild(renderLibraryDemoGrid(f));

    nodes.libraryDemoBody.appendChild(renderLibraryExplanation(f));

    nodes.libraryDemoOverlay.hidden = false;
  }

  function closeLibraryDemo() {
    if (nodes.libraryDemoOverlay) nodes.libraryDemoOverlay.hidden = true;
  }

  function renderLibraryDemoGrid(f) {
    if (f.demoSheet && f.demoSheet.rows && f.demoSheet.rows.length) {
      return renderSpreadsheetGrid(f.demoSheet);
    }
    return renderExamplesGrid(f.examples || []);
  }

  function renderSpreadsheetGrid(sheet) {
    var table = h('table', 'fh-demo-grid');
    var thead = document.createElement('thead');
    var head = document.createElement('tr');
    [''].concat(sheet.letters || ['A', 'B', 'C']).forEach(function (label) {
      head.appendChild(h('th', null, label));
    });
    thead.appendChild(head);
    table.appendChild(thead);

    var body = document.createElement('tbody');
    sheet.rows.forEach(function (row, i) {
      var tr = document.createElement('tr');
      tr.appendChild(h('th', null, String(i + 1)));
      row.forEach(function (cell, colIndex) {
        var cls = i === 0 ? 'fh-demo-cell-head' : '';
        if (i > 0 && String(cell || '').charAt(0) === '=') cls = 'fh-demo-formula-cell';
        if (i > 0 && colIndex === 2 && String(cell || '') !== '') cls = 'fh-demo-result-cell';
        tr.appendChild(h('td', cls, cell));
      });
      body.appendChild(tr);
    });
    table.appendChild(body);
    return table;
  }

  function renderExamplesGrid(examples) {
    var table = h('table', 'fh-demo-grid');
    var thead = document.createElement('thead');
    var head = document.createElement('tr');
    ['', 'A', 'B', 'C'].forEach(function (label) { head.appendChild(h('th', null, label)); });
    thead.appendChild(head);
    table.appendChild(thead);

    var body = document.createElement('tbody');
    var title = document.createElement('tr');
    title.appendChild(h('th', null, '1'));
    title.appendChild(h('td', 'fh-demo-cell-head', 'Dữ liệu mẫu'));
    title.appendChild(h('td', 'fh-demo-cell-head', 'Công thức'));
    title.appendChild(h('td', 'fh-demo-cell-head', 'Kết quả'));
    body.appendChild(title);

    examples.forEach(function (ex, i) {
      var tr = document.createElement('tr');
      tr.appendChild(h('th', null, String(i + 2)));
      tr.appendChild(h('td', null, ex.input));
      tr.appendChild(h('td', 'fh-demo-formula-cell', ex.formula));
      tr.appendChild(h('td', 'fh-demo-result-cell', ex.result));
      body.appendChild(tr);
    });
    table.appendChild(body);
    return table;
  }

  function renderLibraryExplanation(f) {
    var note = h('div', 'fh-demo-note');
    note.appendChild(h('strong', null, 'Giải thích công thức'));
    var steps = f.demoSheet && f.demoSheet.explanation ? f.demoSheet.explanation : null;
    if (steps && steps.length) {
      var list = document.createElement('ol');
      steps.forEach(function (step) { list.appendChild(h('li', null, step)); });
      note.appendChild(list);
    } else {
      note.appendChild(h('p', null,
        'Nếu dữ liệu ở cột A giống tình huống của bạn, dùng công thức ở cột B và đối chiếu kết quả ở cột C để biết hàm đang xử lý gì.'));
    }
    return note;
  }

  /* ============================================================
   * Review ẩn: Workbook + Preview
   * ========================================================= */

  function renderSide() {
    if (!nodes.side) return;
    clear(nodes.side);

    var wbBox = h('section', 'fh-wb');
    wbBox.appendChild(h('h3', 'fh-section-title', 'Workbook hiện tại'));

    if (WB.isEmpty()) {
      var empty = h('div', 'fh-wb-empty');
      empty.appendChild(h('p', null, 'Chưa có Workbook. Bạn vẫn có thể tạo công thức thủ công bằng chế độ Ô / Tuỳ ý.'));
      empty.appendChild(button('Nhập từ Excel', 'btn btn-primary', function () { host.onImport(); }));
      wbBox.appendChild(empty);
    } else {
      var file = WB.getFileName();
      if (file) {
        var fileRow = h('p', 'fh-wb-file');
        fileRow.appendChild(h('span', 'fh-wb-label', 'File'));
        fileRow.appendChild(h('strong', null, file));
        wbBox.appendChild(fileRow);
      }
      var sheetList = h('div', 'fh-wb-sheets');
      WB.getSheets().forEach(function (s, i) {
        var b = button(s.name, 'fh-wb-sheet' + (i === WB.getActiveIndex() ? ' is-on' : ''), function () {
          host.setActiveSheet(i);
        });
        b.title = WB.getHeaders(i).length + ' cột · ' + WB.getDataRowCount(i) + ' dòng dữ liệu';
        sheetList.appendChild(b);
      });
      wbBox.appendChild(sheetList);

      var active = WB.getActiveIndex();
      var meta = h('p', 'fh-wb-meta',
        WB.getHeaders(active).length + ' cột · ' + WB.getDataRowCount(active) + ' dòng dữ liệu');
      wbBox.appendChild(meta);
    }
    nodes.side.appendChild(wbBox);

    var previewBox = h('section', 'fh-preview');
    previewBox.appendChild(h('h3', 'fh-section-title', 'Xem trước kết quả'));
    nodes.previewHost = h('div', 'fh-preview-host');
    previewBox.appendChild(nodes.previewHost);
    nodes.side.appendChild(previewBox);

    renderPreviewTable(null);
    if (state.view === 'tool' && state.lastConfig) schedulePreview(state.lastConfig);
  }

  /* ============================================================
   * API
   * ========================================================= */

  function init(container, options) {
    WB = global.Workbook;
    ENGINE = global.FormulaEngine;
    PRESETS = global.FormulaPresets;
    STORE = global.FormulaStore;
    EXPLAINER = global.FormulaExplainer;

    options = options || {};
    if (options.toast) host.toast = options.toast;
    if (options.onImport) host.onImport = options.onImport;
    if (options.onWorkbookEdit) host.onWorkbookEdit = options.onWorkbookEdit;
    if (options.getSelectedColumn) host.getSelectedColumn = options.getSelectedColumn;
    if (options.setActiveSheet) host.setActiveSheet = options.setActiveSheet;

    state.settings = STORE.getSettings();
    mount(container);

    // Chỉ dựng lại form khi workbook thay đổi về cấu trúc. Với thay đổi dữ liệu
    // thường ('data') chỉ làm mới panel phải, tránh mất focus khi người dùng đang gõ.
    var STRUCTURAL = ['import', 'clear', 'sheet', 'workspace', 'structure'];
    WB.subscribe(function (reason) {
      var structural = STRUCTURAL.indexOf(reason) !== -1;
      if (reason === 'import' || reason === 'clear') {
        // Workbook mới: làm mới danh sách sheet/cột, giữ History (§167)
        var issues = validateRefs(state.config);
        if (issues.length) host.toast(issues[0]);
      }
      renderSide();
      if (structural) renderDataGridPanel();
      if (structural) renderPanel();
      else if (state.view === 'tool') updateOutput();
    });
  }

  function refresh() {
    if (!root) return;
    renderSide();
    renderPanel();
  }

  function reset() {
    if (state.toolId) openTool(state.toolId);
  }

  global.FormulaHelper = {
    init: init,
    refresh: refresh,
    reset: reset,
    openTool: openTool,
    goHome: function () { state.view = 'home'; renderPanel(); }
  };
})(typeof window !== 'undefined' ? window : this);
