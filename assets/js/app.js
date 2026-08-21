/*!
 * app.js — App shell: ghép lưới nhập liệu, bộ sinh SQL và Formula Helper.
 *
 * V2.4.0: Formula Helper dùng workbench lưới + sidebar công cụ bên phải.
 */
(function () {
  'use strict';

  var LEGACY_STORE_KEY = 'genscript.v1';
  var DEVICE_ID_KEY = 'genscript.device.v1';
  var DATA_STORE_PREFIX = 'genscript.data.v2.';
  var FORMULA_DATA_STORE_PREFIX = 'genscript.formulaWorkbook.v1.';
  var THEME_STORE_PREFIX = 'genscript.theme.v1.';
  var WORKSPACE_STORE_PREFIX = 'genscript.workspace.v1.';
  var storage = getStorage();
  var deviceId = ensureDeviceId();
  var STORE_KEY = DATA_STORE_PREFIX + deviceId;
  var FORMULA_STORE_KEY = FORMULA_DATA_STORE_PREFIX + deviceId;
  var THEME_KEY = THEME_STORE_PREFIX + deviceId;
  var WORKSPACE_KEY = WORKSPACE_STORE_PREFIX + deviceId;
  var APP_LOADING_MS = 1500;
  var colorSchemeMeta = document.querySelector('meta[name="color-scheme"]');
  var $ = function (id) { return document.getElementById(id); };

  var el = {
    appLoader: $('appLoader'),
    tableName: $('tableName'),
    whereField: $('whereField'),
    whereColumns: $('whereColumns'),
    selectWhereField: $('selectWhereField'),
    selectWhereMode: $('selectWhereMode'),
    dialect: $('dialect'),
    optionsPanel: $('optionsPanel'),
    btnOptions: $('btnOptions'),
    gridHost: $('gridHost'),
    sheetTabs: $('sheetTabs'),
    sqlCode: $('sqlCode'),
    sqlOut: $('sqlOut'),
    stats: $('stats'),
    warnings: $('warnings'),
    toast: $('toast'),
    reviewToast: $('reviewToast'),
    fileInput: $('fileInput'),
    helpDialog: $('helpDialog'),
    formulaHelpDialog: $('formulaHelpDialog'),
    formulaRoot: $('formulaRoot'),
    wsGenScript: $('workspaceGenScript'),
    wsFormula: $('workspaceFormula'),
    scriptDialog: $('scriptDialog'),
    btnTheme: $('btnTheme'),
    sqlFormatInputs: document.querySelectorAll('input[name="sqlFormat"]')
  };

  var OPT_INPUTS = {
    identityInsert: $('optIdentity'),
    unicodePrefix: $('optUnicode'),
    nullKeyword: $('optNullKeyword'),
    trimValues: $('optTrim'),
    leadingZeroAsText: $('optLeadingZero'),
    skipEmptyRows: $('optSkipEmpty'),
    multiRowInsert: $('optMultiRow'),
    semicolon: $('optSemicolon')
  };

  var state = {
    mode: 'insert',
    activeSheet: 0,
    sheets: [],
    lastSql: '',
    lastResult: null,
    theme: 'dark',
    sqlFormat: 'pretty',
    workspace: 'genscript'
  };

  var formulaWorkbook = {
    activeSheet: 0,
    sheets: [],
    fileName: ''
  };

  var grid = null;
  var renderTimer = null;

  function normalizeMode(mode) {
    return mode === 'update' || mode === 'select' ? mode : 'insert';
  }

  function defaultWhereColumns(mode) {
    return mode === 'select' ? 0 : 1;
  }

  function normalizeWhereColumns(mode, value) {
    var n = parseInt(value, 10);
    if (mode === 'select') return Math.max(0, Math.min(20, isNaN(n) ? 0 : n));
    return Math.max(1, Math.min(20, isNaN(n) ? 1 : n));
  }

  function sheetWhereColumns(sheet, mode) {
    if (mode === 'select') return normalizeWhereColumns(mode, sheet && sheet.selectWhereColumns);
    return normalizeWhereColumns(mode, sheet && sheet.whereColumns);
  }

  function normalizeSelectWhereMode(mode) {
    return mode === 'inAnd' || mode === 'inOr' ? mode : 'matchRows';
  }

  function normalizeSqlFormat(format) {
    return format === 'compact' ? 'compact' : 'pretty';
  }

  function normalizeColumnType(type) {
    return type === 'text' || type === 'number' ? type : 'auto';
  }

  function normalizeColumnTypes(types) {
    return (types || []).map(normalizeColumnType);
  }

  function sheetSelectWhereMode(sheet) {
    return normalizeSelectWhereMode(sheet && sheet.selectWhereMode);
  }

  /* ------------------------------------------------------------ local device */

  function getStorage() {
    try {
      var s = window.localStorage;
      var test = '__genscript_storage_test__';
      s.setItem(test, '1');
      s.removeItem(test);
      return s;
    } catch (err) {
      return null;
    }
  }

  function randomId() {
    if (window.crypto && window.crypto.randomUUID) return window.crypto.randomUUID();
    var bytes = new Uint8Array(16);
    if (window.crypto && window.crypto.getRandomValues) {
      window.crypto.getRandomValues(bytes);
      return Array.prototype.map.call(bytes, function (b) {
        return ('0' + b.toString(16)).slice(-2);
      }).join('');
    }
    return String(Date.now()) + '-' + Math.random().toString(36).slice(2);
  }

  function ensureDeviceId() {
    if (!storage) return 'session';
    var id = null;
    try { id = storage.getItem(DEVICE_ID_KEY); } catch (err) { id = null; }
    if (!id) {
      id = randomId();
      try { storage.setItem(DEVICE_ID_KEY, id); } catch (err) { /* ignore */ }
    }
    return id;
  }

  function storageGet(key) {
    if (!storage) return null;
    try { return storage.getItem(key); } catch (err) { return null; }
  }

  function storageSet(key, value) {
    if (!storage) return false;
    try {
      storage.setItem(key, value);
      return true;
    } catch (err) {
      return false;
    }
  }

  function storageRemove(key) {
    if (!storage) return;
    try { storage.removeItem(key); } catch (err) { /* ignore */ }
  }

  /* ------------------------------------------------------------ sheets */

  function newSheet(name, data, owner) {
    var list = owner && owner.sheets ? owner.sheets : state.sheets;
    return {
      name: name || 'Bang_' + (list.length + 1),
      data: data && data.length ? data : [[]],
      whereColumns: 1,
      selectWhereColumns: 0,
      selectWhereMode: 'matchRows',
      columnTypes: []
    };
  }

  function currentSheet() {
    return state.sheets[state.activeSheet];
  }

  function renderTabs() {
    el.sheetTabs.innerHTML = '';
    state.sheets.forEach(function (sheet, i) {
      var tab = document.createElement('button');
      tab.type = 'button';
      tab.className = 'tab' + (i === state.activeSheet ? ' is-on' : '');
      tab.setAttribute('role', 'tab');
      tab.setAttribute('aria-selected', i === state.activeSheet ? 'true' : 'false');
      tab.title = sheet.name;

      var name = document.createElement('span');
      name.className = 'tab-name';
      name.textContent = sheet.name;
      tab.appendChild(name);

      if (state.sheets.length > 1) {
        var del = document.createElement('span');
        del.className = 'mini-del';
        del.textContent = '×';
        del.title = 'Xoá bảng ' + sheet.name;
        del.dataset.del = i;
        tab.appendChild(del);
      }

      tab.addEventListener('click', function (e) {
        if (e.target.dataset && e.target.dataset.del !== undefined) {
          removeSheet(i);
          return;
        }
        selectSheet(i);
      });
      el.sheetTabs.appendChild(tab);
    });

    var add = document.createElement('button');
    add.type = 'button';
    add.className = 'tab-add';
    add.textContent = '+';
    add.title = 'Thêm bảng';
    add.setAttribute('aria-label', 'Thêm bảng');
    add.addEventListener('click', function () {
      state.sheets.push(newSheet());
      selectSheet(state.sheets.length - 1);
    });
    el.sheetTabs.appendChild(add);
  }

  function commitActiveSheet() {
    var sheet = currentSheet();
    if (!sheet || !grid) return;
    var next = grid.getData();
    if (sheet.seed && JSON.stringify(next) !== JSON.stringify(sheet.data)) sheet.seed = false;
    sheet.data = next;
    sheet.columnTypes = grid.getColumnTypes();
    if (sheet.seed && sheet.name !== el.tableName.value) sheet.seed = false;
    sheet.name = el.tableName.value;
    if (state.mode === 'select') {
      sheet.selectWhereColumns = normalizeWhereColumns(state.mode, el.whereColumns.value);
      sheet.selectWhereMode = normalizeSelectWhereMode(el.selectWhereMode.value);
    } else if (state.mode === 'update') {
      sheet.whereColumns = normalizeWhereColumns(state.mode, el.whereColumns.value);
    }
  }

  function syncActiveSheetName() {
    var sheet = currentSheet();
    if (!sheet) return;
    var nextName = el.tableName.value;
    if (sheet.name === nextName) return;
    if (sheet.seed) sheet.seed = false;
    sheet.name = nextName;
    renderTabs();
  }

  function selectSheet(index, skipCommit) {
    if (!skipCommit) commitActiveSheet();
    state.activeSheet = Math.max(0, Math.min(index, state.sheets.length - 1));
    var sheet = currentSheet();
    el.tableName.value = sheet.name;
    el.whereColumns.value = sheetWhereColumns(sheet, state.mode);
    el.selectWhereMode.value = sheetSelectWhereMode(sheet);
    grid.setColumnTypes(normalizeColumnTypes(sheet.columnTypes));
    grid.setMode(state.mode, sheetWhereColumns(sheet, state.mode));
    grid.setData(sheet.data);
    renderTabs();
    scheduleRender();
  }

  function removeSheet(index) {
    if (state.sheets.length <= 1) return;
    state.sheets.splice(index, 1);
    if (state.activeSheet >= state.sheets.length) state.activeSheet = state.sheets.length - 1;
    var sheet = currentSheet();
    el.tableName.value = sheet.name;
    el.whereColumns.value = sheetWhereColumns(sheet, state.mode);
    el.selectWhereMode.value = sheetSelectWhereMode(sheet);
    grid.setColumnTypes(normalizeColumnTypes(sheet.columnTypes));
    grid.setMode(state.mode, sheetWhereColumns(sheet, state.mode));
    grid.setData(sheet.data);
    renderTabs();
    scheduleRender();
  }

  /* ------------------------------------------------------------ options */

  function readOptions() {
    var o = {
      dialect: el.dialect.value,
      whereColumns: normalizeWhereColumns(state.mode, el.whereColumns.value),
      selectWhereMode: normalizeSelectWhereMode(el.selectWhereMode.value),
      sqlFormat: state.sqlFormat,
      columnTypes: grid ? grid.getColumnTypes() : [],
      goEvery: Math.max(0, parseInt($('optGoEvery').value, 10) || 0)
    };
    Object.keys(OPT_INPUTS).forEach(function (k) { o[k] = OPT_INPUTS[k].checked; });
    return o;
  }

  function syncOptionAvailability() {
    var d = window.SqlGen.DIALECTS[el.dialect.value] || window.SqlGen.DIALECTS.mssql;
    OPT_INPUTS.identityInsert.closest('.opt')
      .classList.toggle('is-off', !d.supportsIdentityInsert || state.mode !== 'insert');
    OPT_INPUTS.unicodePrefix.closest('.opt').classList.toggle('is-off', !d.unicodePrefix);
    OPT_INPUTS.multiRowInsert.closest('.opt')
      .classList.toggle('is-off', state.mode !== 'insert' || d.maxRowsPerInsert < 2);
    $('optGoEvery').closest('.opt').classList.toggle('is-off', !d.supportsGo);
  }

  function setMode(mode) {
    mode = normalizeMode(mode);
    if (grid && state.sheets.length && mode !== state.mode) commitActiveSheet();
    state.mode = mode;
    document.body.dataset.mode = mode;
    Array.prototype.forEach.call(document.querySelectorAll('.mode-btn'), function (b) {
      var on = b.dataset.mode === mode;
      b.classList.toggle('is-on', on);
      b.setAttribute('aria-checked', on ? 'true' : 'false');
    });
    el.whereField.hidden = mode === 'insert';
    el.selectWhereField.hidden = mode !== 'select';
    el.whereColumns.min = mode === 'select' ? '0' : '1';
    el.whereColumns.value = state.sheets.length ? sheetWhereColumns(currentSheet(), mode) : defaultWhereColumns(mode);
    el.selectWhereMode.value = state.sheets.length ? sheetSelectWhereMode(currentSheet()) : 'matchRows';
    $('gridHint').textContent = mode === 'update'
      ? 'Cột tô hổ phách là WHERE · Arrow di chuyển · double click để sửa · Ctrl+C/V · Delete.'
      : mode === 'select'
        ? 'WHERE = 0 sẽ SELECT toàn bảng · chọn kiểu WHERE khi lọc theo dữ liệu.'
        : 'Arrow di chuyển · double click để sửa · Ctrl+C/V · Delete · kéo chấm vuông để fill.';
    grid.setMode(mode, normalizeWhereColumns(mode, el.whereColumns.value));
    syncOptionAvailability();
    scheduleRender();
  }

  function syncSqlFormatControls() {
    Array.prototype.forEach.call(el.sqlFormatInputs, function (input) {
      input.checked = input.value === state.sqlFormat;
    });
  }

  function setSqlFormat(format) {
    state.sqlFormat = normalizeSqlFormat(format);
    syncSqlFormatControls();
    if (el.scriptDialog && el.scriptDialog.open) render(true);
    else scheduleRender();
  }

  /* ------------------------------------------------------------ workspace */

  function normalizeWorkspace(name) {
    return name === 'formula' ? 'formula' : 'genscript';
  }

  /** Chuyển giữa hai workspace, mỗi workspace giữ workbook/lưới riêng. */
  function setWorkspace(name, skipPersist) {
    state.workspace = normalizeWorkspace(name);
    document.body.dataset.workspace = state.workspace;

    if (el.wsGenScript) el.wsGenScript.hidden = state.workspace !== 'genscript';
    if (el.wsFormula) el.wsFormula.hidden = state.workspace !== 'formula';

    Array.prototype.forEach.call(document.querySelectorAll('.ws-btn'), function (b) {
      var on = b.dataset.workspace === state.workspace;
      b.classList.toggle('is-on', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
    });

    if (state.workspace === 'formula') {
      commitActiveSheet();
      ensureFormulaHelper();
      notifyWorkbook('workspace');
    } else if (grid && state.sheets.length) {
      selectSheet(state.activeSheet, true);
    }
    if (!skipPersist) storageSet(WORKSPACE_KEY, state.workspace);
  }

  var formulaReady = false;

  function ensureFormulaHelper() {
    if (formulaReady || !window.FormulaHelper || !el.formulaRoot) return;
    formulaReady = true;
    window.FormulaHelper.init(el.formulaRoot, {
      toast: toast,
      onImport: function () { el.fileInput.click(); },
      onWorkbookEdit: function (reason) {
        persistFormulaWorkbook();
        notifyWorkbook(reason || 'data');
      },
      getSelectedColumn: function () {
        return null;
      },
      setActiveSheet: function (index) {
        var max = Math.max(0, formulaWorkbook.sheets.length - 1);
        formulaWorkbook.activeSheet = Math.max(0, Math.min(index | 0, max));
        persistFormulaWorkbook();
        notifyWorkbook('sheet');
      }
    });
  }

  function notifyWorkbook(reason) {
    if (window.Workbook) window.Workbook.notify(reason);
  }

  /* ------------------------------------------------------------ theme */

  function applyTheme(theme) {
    state.theme = theme === 'light' ? 'light' : 'dark';
    document.documentElement.dataset.theme = state.theme;
    document.body.dataset.theme = state.theme;
    if (colorSchemeMeta) {
      colorSchemeMeta.setAttribute('content', state.theme === 'dark' ? 'dark light' : 'light dark');
    }
    if (el.btnTheme) {
      el.btnTheme.setAttribute('aria-pressed', state.theme === 'dark' ? 'true' : 'false');
      el.btnTheme.dataset.theme = state.theme;
      el.btnTheme.title = state.theme === 'dark' ? 'Đang dùng giao diện tối' : 'Đang dùng giao diện sáng';
      el.btnTheme.setAttribute('aria-label', state.theme === 'dark'
        ? 'Đang dùng giao diện tối, bấm để chuyển sang giao diện sáng'
        : 'Đang dùng giao diện sáng, bấm để chuyển sang giao diện tối');
    }
  }

  function toggleTheme() {
    var current = document.documentElement.dataset.theme === 'light' ? 'light' : 'dark';
    applyTheme(current === 'dark' ? 'light' : 'dark');
    storageSet(THEME_KEY, state.theme);
  }

  /* ------------------------------------------------------------ highlight */

  var TOKEN_RE = /(N?'(?:[^']|'')*')|(\[(?:[^\]]|\]\])*\])|(`[^`]*`)|("(?:[^"]|"")*")|\b(NULL)\b|\b(GO)\b|\b(SELECT|FROM|IN|OR|INSERT|INTO|VALUES|UPDATE|SET|WHERE|AND|IS|ON|OFF|IDENTITY_INSERT)\b|(\b\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\b)/g;

  function escapeHtml(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function highlight(sql) {
    if (sql.length > 400000) return escapeHtml(sql);
    var lines = sql.split('\n');
    if (lines.length > 6000) return escapeHtml(sql);
    return lines.map(function (line) {
      return '<span class="ln">' + highlightLine(line) + '</span>';
    }).join('');
  }

  function highlightLine(sql) {
    var out = '', last = 0, m;
    TOKEN_RE.lastIndex = 0;
    while ((m = TOKEN_RE.exec(sql)) !== null) {
      out += escapeHtml(sql.slice(last, m.index));
      var cls = m[1] ? 'tok-str'
              : (m[2] || m[3] || m[4]) ? 'tok-id'
              : m[5] ? 'tok-null'
              : m[6] ? 'tok-go'
              : m[7] ? 'tok-kw'
              : 'tok-num';
      out += '<span class="' + cls + '">' + escapeHtml(m[0]) + '</span>';
      last = m.index + m[0].length;
    }
    out += escapeHtml(sql.slice(last));
    return out;
  }

  /* ------------------------------------------------------------ render */

  function scheduleRender() {
    if (renderTimer) clearTimeout(renderTimer);
    renderTimer = setTimeout(syncState, 120);
  }

  function syncState() {
    renderTimer = null;
    var sheet = currentSheet();
    var oldName = sheet && sheet.name;
    commitActiveSheet();
    if (sheet && sheet.name !== oldName) renderTabs();
    persist();
  }

  function render(showReview) {
    renderTimer = null;
    var matrix = grid.getData();
    var table = el.tableName.value;
    var result = window.SqlGen.build(state.mode, table, matrix, readOptions());

    state.lastSql = result.sql;
    state.lastResult = result;

    if (showReview || el.scriptDialog.open) renderReview(result);

    commitActiveSheet();
    persist();
    return result;
  }

  function renderReview(result) {
    if (result.sql) {
      el.sqlCode.innerHTML = highlight(result.sql);
    } else {
      el.sqlCode.innerHTML = '<span class="sql-empty">Chưa có gì để sinh. Kiểm tra tên bảng và dữ liệu trên lưới.</span>';
    }

    if (result.warnings.length) {
      el.warnings.innerHTML = result.warnings.slice(0, 20).map(function (w) {
        return '<p>' + escapeHtml(w) + '</p>';
      }).join('') + (result.warnings.length > 20
        ? '<p>… và ' + (result.warnings.length - 20) + ' cảnh báo khác</p>' : '');
      el.warnings.hidden = false;
    } else {
      el.warnings.hidden = true;
      el.warnings.innerHTML = '';
    }

    var s = result.stats;
    el.stats.textContent = s.statements
      ? s.rows + ' dòng · ' + s.columns + ' cột · ' + s.statements + ' câu lệnh · ' + formatBytes(result.sql.length)
      : 'Chưa có câu lệnh hợp lệ';
  }

  function openScriptReview() {
    commitActiveSheet();
    render(true);
    if (el.scriptDialog.showModal) el.scriptDialog.showModal();
    else el.scriptDialog.setAttribute('open', '');
  }

  function formatBytes(n) {
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
    return (n / 1048576).toFixed(1) + ' MB';
  }

  /* ------------------------------------------------------------ storage */

  function persist() {
    storageSet(STORE_KEY, JSON.stringify({
      mode: state.mode,
      activeSheet: state.activeSheet,
      sheets: state.sheets,
      dialect: el.dialect.value,
      sqlFormat: state.sqlFormat,
      workspace: state.workspace,
      goEvery: $('optGoEvery').value,
      opts: Object.keys(OPT_INPUTS).reduce(function (acc, k) {
        acc[k] = OPT_INPUTS[k].checked; return acc;
      }, {})
    }));
  }

  function restore() {
    var raw = storageGet(STORE_KEY);
    if (!raw) {
      raw = storageGet(LEGACY_STORE_KEY);
      if (raw) {
        storageSet(STORE_KEY, raw);
        storageRemove(LEGACY_STORE_KEY);
      }
    }
    if (!raw) return false;
    try {
      var saved = JSON.parse(raw);
      if (!saved || !saved.sheets || !saved.sheets.length) return false;
      if (saved.sheets.length === 1 && saved.sheets[0].seed) {
        storageRemove(STORE_KEY);
        storageRemove(LEGACY_STORE_KEY);
        return false;
      }
      state.sheets = saved.sheets;
      state.activeSheet = Math.min(saved.activeSheet || 0, saved.sheets.length - 1);
      state.mode = normalizeMode(saved.mode);
      state.sqlFormat = normalizeSqlFormat(saved.sqlFormat);
      state.workspace = normalizeWorkspace(saved.workspace);
      if (saved.dialect) el.dialect.value = saved.dialect;
      if (saved.goEvery != null) $('optGoEvery').value = saved.goEvery;
      if (saved.opts) {
        Object.keys(OPT_INPUTS).forEach(function (k) {
          if (saved.opts[k] !== undefined) OPT_INPUTS[k].checked = !!saved.opts[k];
        });
      }
      return true;
    } catch (err) { return false; }
  }

  function persistFormulaWorkbook() {
    storageSet(FORMULA_STORE_KEY, JSON.stringify(formulaWorkbook));
  }

  function restoreFormulaWorkbook() {
    var raw = storageGet(FORMULA_STORE_KEY);
    if (!raw) return false;
    try {
      var saved = JSON.parse(raw);
      if (!saved || !saved.sheets || !saved.sheets.length) return false;
      formulaWorkbook.sheets = saved.sheets;
      formulaWorkbook.activeSheet = Math.min(saved.activeSheet || 0, formulaWorkbook.sheets.length - 1);
      formulaWorkbook.fileName = saved.fileName || '';
      return true;
    } catch (err) { return false; }
  }

  function ensureFormulaWorkbook() {
    if (!formulaWorkbook.sheets.length) {
      formulaWorkbook.sheets = [newSheet('FormulaData', null, formulaWorkbook)];
      formulaWorkbook.activeSheet = 0;
      formulaWorkbook.fileName = '';
    }
  }

  /* ------------------------------------------------------------ import */

  function sanitizeName(name) {
    return String(name || '').replace(/[\r\n\t]/g, ' ').trim() || 'Bang';
  }

  function makeImportedSheet(name, data) {
    return {
      name: sanitizeName(name),
      data: data,
      whereColumns: 1,
      selectWhereColumns: 0,
      selectWhereMode: 'matchRows',
      columnTypes: []
    };
  }

  function sheetHasValues(sheet) {
    var data = (sheet && sheet.data) || [];
    for (var r = 0; r < data.length; r++) {
      var row = data[r] || [];
      for (var c = 0; c < row.length; c++) {
        if (String(row[c] == null ? '' : row[c]).trim() !== '') return true;
      }
    }
    return false;
  }

  function isInitialFormulaWorkbook(target) {
    return target && target.sheets && target.sheets.length === 1 &&
      target.sheets[0].name === 'FormulaData' && !target.fileName &&
      !sheetHasValues(target.sheets[0]);
  }

  function importWorkbook(file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      var wb;
      var xlsx = window.XLSX;
      if (!xlsx || !xlsx.read || !xlsx.utils) {
        toast('Không đọc được file: thư viện Excel chưa được tải. Hãy tải lại app rồi thử lại.');
        return;
      }
      try {
        wb = xlsx.read(new Uint8Array(e.target.result), { type: 'array', cellDates: false });
      } catch (err) {
        toast('Không đọc được file: ' + err.message);
        return;
      }
      var target = state.workspace === 'formula' ? formulaWorkbook : state;
      var targetIsFormula = target === formulaWorkbook;
      if (!targetIsFormula) commitActiveSheet();
      if (target.sheets.length === 1 && target.sheets[0].seed) target.sheets = [];
      if (targetIsFormula && isInitialFormulaWorkbook(target)) target.sheets = [];
      var added = 0;
      wb.SheetNames.forEach(function (name) {
        var ws = wb.Sheets[name];
        if (!ws) return;
        var rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: false });
        if (!rows.length) return;
        var clean = rows.map(function (r) {
          return r.map(function (v) { return v == null ? '' : String(v); });
        });
        target.sheets.push(makeImportedSheet(name, clean));
        added++;
      });
      if (!added) {
        if (!target.sheets.length) target.sheets = [newSheet(targetIsFormula ? 'FormulaData' : 'Bang_1', null, target)];
        if (!targetIsFormula) renderTabs();
        else {
          persistFormulaWorkbook();
          notifyWorkbook('import');
        }
        toast('File không có sheet nào chứa dữ liệu.');
        return;
      }
      target.activeSheet = target.sheets.length - added;
      if (targetIsFormula) {
        target.fileName = file.name;
        if (window.Workbook) window.Workbook.setFileName(file.name);
        persistFormulaWorkbook();
        notifyWorkbook('import');
      } else {
        selectSheet(target.activeSheet, true);
      }
      toast('Đã nạp ' + added + ' bảng từ ' + file.name);
    };
    reader.onerror = function () { toast('Không đọc được file.'); };
    reader.readAsArrayBuffer(file);
  }

  /* ------------------------------------------------------------ actions */

  function toast(msg) {
    var target = (el.scriptDialog && el.scriptDialog.open && el.reviewToast) ? el.reviewToast : el.toast;
    target.textContent = msg;
    target.hidden = false;
    requestAnimationFrame(function () { target.classList.add('is-on'); });
    clearTimeout(target._toastTimer);
    target._toastTimer = setTimeout(function () {
      target.classList.remove('is-on');
      setTimeout(function () { target.hidden = true; }, 220);
    }, 2200);
  }

  function copySql() {
    render(false);
    if (!state.lastSql) { toast('Chưa có script để sao chép.'); return; }
    var done = function () { toast('Đã sao chép script.'); };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(state.lastSql).then(done, fallbackCopy);
    } else fallbackCopy();

    function fallbackCopy() {
      var ta = document.createElement('textarea');
      ta.value = state.lastSql;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); done(); }
      catch (err) { toast('Trình duyệt chặn thao tác sao chép.'); }
      document.body.removeChild(ta);
    }
  }

  function downloadSql() {
    render(false);
    if (!state.lastSql) { toast('Chưa có script để tải.'); return; }
    var name = (el.tableName.value || 'script').replace(/[^\w.-]+/g, '_');
    var blob = new Blob(['\ufeff' + state.lastSql], { type: 'text/plain;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = name + '_' + state.mode + '.sql';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function finishAppLoading() {
    if (!el.appLoader) return;
    el.appLoader.style.setProperty('--loader-duration', APP_LOADING_MS + 'ms');
    setTimeout(function () {
      el.appLoader.classList.add('is-hiding');
      document.body.classList.remove('is-loading');
      el.appLoader.setAttribute('aria-hidden', 'true');
      setTimeout(function () { el.appLoader.hidden = true; }, 320);
    }, APP_LOADING_MS);
  }

  /* ------------------------------------------------------------ init */

  function init() {
    applyTheme(storageGet(THEME_KEY) || 'dark');

    restoreFormulaWorkbook();
    ensureFormulaWorkbook();
    if (window.Workbook) {
      window.Workbook.bind(formulaWorkbook);
      window.Workbook.setFileName(formulaWorkbook.fileName || '');
    }
    if (window.FormulaStore) window.FormulaStore.scopeTo(deviceId);

    grid = new Grid(el.gridHost, {
      onChange: scheduleRender,
      onStatus: toast,
      onHistoryChange: function (history) {
        $('btnUndo').disabled = !history.canUndo;
        $('btnRedo').disabled = !history.canRedo;
      },
      onColumnTypesChange: function (types) {
        var sheet = currentSheet();
        if (sheet) sheet.columnTypes = normalizeColumnTypes(types);
        scheduleRender();
      },
      mode: state.mode
    });

    if (!restore()) {
      state.sheets = [newSheet('Bang_1')];
      state.activeSheet = 0;
    }

    document.body.dataset.mode = state.mode;
    var sheet = currentSheet();
    el.tableName.value = sheet.name;
    el.whereColumns.value = sheetWhereColumns(sheet, state.mode);
    el.selectWhereMode.value = sheetSelectWhereMode(sheet);
    grid.setColumnTypes(normalizeColumnTypes(sheet.columnTypes));
    grid.setMode(state.mode, sheetWhereColumns(sheet, state.mode));
    grid.setData(sheet.data);
    renderTabs();
    setMode(state.mode);
    syncSqlFormatControls();

    Array.prototype.forEach.call(document.querySelectorAll('.ws-btn'), function (b) {
      b.addEventListener('click', function () { setWorkspace(b.dataset.workspace); });
    });
    setWorkspace(storageGet(WORKSPACE_KEY) || state.workspace, true);

    el.tableName.addEventListener('input', function () {
      syncActiveSheetName();
      scheduleRender();
    });
    el.whereColumns.addEventListener('input', function () {
      grid.setMode(state.mode, normalizeWhereColumns(state.mode, el.whereColumns.value));
      scheduleRender();
    });
    el.selectWhereMode.addEventListener('change', scheduleRender);
    el.dialect.addEventListener('change', function () {
      var d = window.SqlGen.DIALECTS[el.dialect.value];
      if (d) OPT_INPUTS.semicolon.checked = d.defaultSemicolon;
      syncOptionAvailability();
      scheduleRender();
    });

    Array.prototype.forEach.call(document.querySelectorAll('.mode-btn'), function (b) {
      b.addEventListener('click', function () { setMode(b.dataset.mode); });
    });

    Object.keys(OPT_INPUTS).forEach(function (k) {
      OPT_INPUTS[k].addEventListener('change', scheduleRender);
    });
    $('optGoEvery').addEventListener('input', scheduleRender);
    Array.prototype.forEach.call(el.sqlFormatInputs, function (input) {
      input.addEventListener('change', function () {
        if (input.checked) setSqlFormat(input.value);
      });
    });

    el.btnOptions.addEventListener('click', function () {
      var open = el.optionsPanel.hidden;
      el.optionsPanel.hidden = !open;
      el.btnOptions.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
    if (el.btnTheme) {
      el.btnTheme.addEventListener('click', function (e) {
        e.preventDefault();
        toggleTheme();
      });
    }

    $('btnAddRows').addEventListener('click', function () { grid.addRows(10); });
    $('btnAddCols').addEventListener('click', function () { grid.addCols(1); });
    $('btnUndo').addEventListener('click', function () { grid.undo(); });
    $('btnRedo').addEventListener('click', function () { grid.redo(); });
    $('btnClear').addEventListener('click', function () {
      if (confirm('Xoá toàn bộ dữ liệu của bảng "' + el.tableName.value + '"?')) {
        grid.clear();
      }
    });

    $('btnGenerate').addEventListener('click', openScriptReview);
    $('btnCopy').addEventListener('click', copySql);
    $('btnDownload').addEventListener('click', downloadSql);
    el.sqlOut.addEventListener('click', copySql);
    $('btnImport').addEventListener('click', function () { el.fileInput.click(); });
    el.fileInput.addEventListener('change', function () {
      if (this.files && this.files[0]) importWorkbook(this.files[0]);
      this.value = '';
    });

    $('btnHelp').addEventListener('click', function () {
      var dialog = state.workspace === 'formula' ? el.formulaHelpDialog : el.helpDialog;
      if (!dialog) return;
      if (dialog.showModal) dialog.showModal();
      else dialog.setAttribute('open', '');
    });

    // Modal hướng dẫn Formula Helper: không đóng khi bấm ra ngoài (§118).
    if (el.formulaHelpDialog) {
      var closeFormulaHelp = function () {
        if (el.formulaHelpDialog.close) el.formulaHelpDialog.close();
        else el.formulaHelpDialog.removeAttribute('open');
      };
      el.formulaHelpDialog.addEventListener('click', function (e) {
        if (e.target === el.formulaHelpDialog) e.stopPropagation();
      });
      var btnClose = $('btnFormulaHelpClose');
      var btnDone = $('btnFormulaHelpDone');
      if (btnClose) btnClose.addEventListener('click', closeFormulaHelp);
      if (btnDone) btnDone.addEventListener('click', closeFormulaHelp);
    }

    document.addEventListener('keydown', function (e) {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); downloadSql(); }
      if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'c') { e.preventDefault(); copySql(); }
      var target = e.target;
      var isEditor = target && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName));
      if (!isEditor && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) grid.redo(); else grid.undo();
      }
      if (!isEditor && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault(); grid.redo();
      }
    });

    // Thả file vào cửa sổ để nạp
    ['dragover', 'drop'].forEach(function (evt) {
      window.addEventListener(evt, function (e) {
        e.preventDefault();
        if (evt === 'drop' && e.dataTransfer.files && e.dataTransfer.files[0]) {
          importWorkbook(e.dataTransfer.files[0]);
        }
      });
    });

    syncState();
    finishAppLoading();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
