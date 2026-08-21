/*!
 * grid.js — Lưới nhập liệu kiểu bảng tính: chọn vùng, copy/cut/paste,
 * undo/redo, fill handle, fill down/right và thao tác dòng/cột.
 */
(function (global) {
  'use strict';

  var MIN_ROWS = 20;
  var MIN_COLS = 4;
  var HISTORY_LIMIT = 100;
  var COLUMN_TYPES = ['auto', 'text', 'number'];
  var COLUMN_TYPE_LABELS = {
    auto: 'Auto',
    text: 'Chuỗi',
    number: 'Số'
  };

  function colLabel(n) {
    var s = '';
    n += 1;
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function cloneMatrix(matrix) {
    return matrix.map(function (row) { return row.slice(); });
  }

  function parseClipboard(text) {
    var normalized = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    var rows = [], row = [], cell = '', inQuotes = false;
    for (var i = 0; i < normalized.length; i++) {
      var ch = normalized[i];
      if (inQuotes) {
        if (ch === '"') {
          if (normalized[i + 1] === '"') { cell += '"'; i++; }
          else inQuotes = false;
        } else cell += ch;
      } else if (ch === '"' && cell === '') {
        inQuotes = true;
      } else if (ch === '\t') {
        row.push(cell); cell = '';
      } else if (ch === '\n') {
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += ch;
    }
    row.push(cell);
    rows.push(row);
    while (rows.length > 1 && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') rows.pop();
    return rows;
  }

  function normalizeRange(range) {
    return {
      r1: Math.min(range.r1, range.r2),
      c1: Math.min(range.c1, range.c2),
      r2: Math.max(range.r1, range.r2),
      c2: Math.max(range.c1, range.c2)
    };
  }

  function isStrictNumber(value) {
    return /^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(String(value).trim());
  }

  function cleanNumber(value) {
    return String(Math.round(value * 1000000000000) / 1000000000000);
  }

  function normalizeColumnType(type) {
    return COLUMN_TYPES.indexOf(type) >= 0 ? type : 'auto';
  }

  function normalizeColumnTypes(types, width) {
    var out = [];
    for (var i = 0; i < width; i++) out.push(normalizeColumnType(types && types[i]));
    return out;
  }

  function Grid(root, options) {
    options = options || {};
    this.root = root;
    this.onChange = options.onChange || function () {};
    this.onStatus = options.onStatus || function () {};
    this.onHistoryChange = options.onHistoryChange || function () {};
    this.onColumnTypesChange = options.onColumnTypesChange || function () {};
    this.mode = options.mode || 'insert';
    this.whereColumns = 1;
    this.data = [];
    this.columnTypes = [];
    this.active = { r: 0, c: 0 };
    this.selection = { r1: 0, c1: 0, r2: 0, c2: 0 };
    this.history = [];
    this.future = [];
    this._lastRecordKey = '';
    this._lastRecordAt = 0;
    this._draggingSelection = false;
    this._filling = false;
    this._fillTarget = null;
    this.editing = false;
    this._editCell = null;
    this._editOriginal = '';
    this._keyboardActive = false;
    this._buildSkeleton();
    this._buildContextMenu();
    this._bind();
  }

  Grid.prototype._buildSkeleton = function () {
    this.root.innerHTML = '';
    this.root.tabIndex = 0;
    this.table = document.createElement('table');
    this.table.className = 'grid';
    this.thead = document.createElement('thead');
    this.tbody = document.createElement('tbody');
    this.table.appendChild(this.thead);
    this.table.appendChild(this.tbody);
    this.root.appendChild(this.table);
  };

  Grid.prototype._buildContextMenu = function () {
    var menu = document.createElement('div');
    menu.className = 'grid-menu';
    menu.hidden = true;
    menu.setAttribute('role', 'menu');
    menu.innerHTML = [
      '<button type="button" data-action="copy">Sao chép <kbd>Ctrl+C</kbd></button>',
      '<button type="button" data-action="cut">Cắt <kbd>Ctrl+X</kbd></button>',
      '<button type="button" data-action="paste">Dán <kbd>Ctrl+V</kbd></button>',
      '<span class="grid-menu-sep"></span>',
      '<button type="button" data-action="fill-down">Fill xuống <kbd>Ctrl+D</kbd></button>',
      '<button type="button" data-action="fill-right">Fill sang phải <kbd>Ctrl+R</kbd></button>',
      '<button type="button" data-action="clear">Xoá nội dung <kbd>Delete</kbd></button>',
      '<span class="grid-menu-sep"></span>',
      '<button type="button" data-action="delete-row">Xoá dòng</button>',
      '<button type="button" data-action="delete-col">Xoá cột</button>',
      '<span class="grid-menu-sep"></span>',
      '<button type="button" data-action="col-type-auto">Kiểu cột: Auto</button>',
      '<button type="button" data-action="col-type-text">Kiểu cột: Chuỗi</button>',
      '<button type="button" data-action="col-type-number">Kiểu cột: Số</button>'
    ].join('');
    document.body.appendChild(menu);
    this.menu = menu;
  };

  Grid.prototype.setMode = function (mode, whereColumns) {
    this.mode = mode;
    this.whereColumns = Math.max(0, whereColumns || 0);
    this._applyKeyHighlight();
  };

  Grid.prototype.setData = function (matrix) {
    var rows = (matrix && matrix.length) ? matrix : [[]];
    var cols = 0, i;
    for (i = 0; i < rows.length; i++) cols = Math.max(cols, (rows[i] || []).length);
    cols = Math.max(cols, MIN_COLS);
    var target = Math.max(rows.length, MIN_ROWS);

    this.data = [];
    for (i = 0; i < target; i++) {
      var src = rows[i] || [];
      var row = [];
      for (var j = 0; j < cols; j++) row.push(src[j] == null ? '' : String(src[j]));
      this.data.push(row);
    }
    this.columnTypes = normalizeColumnTypes(this.columnTypes, cols);
    this.active = { r: 0, c: 0 };
    this.selection = { r1: 0, c1: 0, r2: 0, c2: 0 };
    this.history = [];
    this.future = [];
    this._notifyHistory();
    this.render();
  };

  Grid.prototype.getData = function () {
    var rows = cloneMatrix(this.data);
    var lastRow = -1, lastCol = -1;
    for (var i = 0; i < rows.length; i++) {
      for (var j = 0; j < rows[i].length; j++) {
        if (String(rows[i][j]).trim() !== '') {
          lastRow = Math.max(lastRow, i);
          lastCol = Math.max(lastCol, j);
        }
      }
    }
    if (lastRow < 0) return [];
    var out = [];
    for (var r = 0; r <= lastRow; r++) out.push(rows[r].slice(0, lastCol + 1));
    return out;
  };

  Grid.prototype.getColumnTypes = function () {
    var width = this.data[0] ? this.data[0].length : MIN_COLS;
    this.columnTypes = normalizeColumnTypes(this.columnTypes, width);
    return this.columnTypes.slice();
  };

  Grid.prototype.setColumnTypes = function (types) {
    var width = this.data[0] ? this.data[0].length : MIN_COLS;
    this.columnTypes = normalizeColumnTypes(types, width);
    this._applyKeyHighlight();
  };

  Grid.prototype._notifyColumnTypes = function () {
    this.onColumnTypesChange(this.getColumnTypes());
  };

  Grid.prototype.setColumnType = function (col, type) {
    var width = this.data[0] ? this.data[0].length : MIN_COLS;
    if (col < 0 || col >= width) return;
    this.columnTypes = normalizeColumnTypes(this.columnTypes, width);
    this.columnTypes[col] = normalizeColumnType(type);
    this._applyKeyHighlight();
    this._notifyColumnTypes();
    this.onChange();
    this.onStatus('Đã đặt cột ' + colLabel(col) + ' là ' + COLUMN_TYPE_LABELS[this.columnTypes[col]] + '.');
  };

  Grid.prototype.toggleColumnType = function (col) {
    var current = normalizeColumnType(this.columnTypes[col]);
    var next = current === 'auto' ? 'text' : current === 'text' ? 'number' : 'auto';
    this.setColumnType(col, next);
  };

  Grid.prototype.render = function () {
    var self = this;
    var cols = this.data[0] ? this.data[0].length : MIN_COLS;
    var htr = document.createElement('tr');
    var corner = document.createElement('th');
    corner.className = 'gutter corner';
    corner.textContent = '';
    corner.title = 'Chọn toàn bộ bảng';
    htr.appendChild(corner);

    for (var j = 0; j < cols; j++) {
      var th = document.createElement('th');
      th.className = 'col-head';
      th.dataset.col = j;
      th.title = 'Click để chọn cột ' + colLabel(j);
      var label = document.createElement('span');
      label.className = 'col-letter';
      label.textContent = colLabel(j);
      var badge = document.createElement('button');
      badge.type = 'button';
      badge.className = 'col-badge';
      badge.title = 'Bấm để đổi kiểu dữ liệu cột: Auto / Chuỗi / Số';
      badge.setAttribute('aria-label', 'Đổi kiểu dữ liệu cột ' + colLabel(j));
      var del = document.createElement('button');
      del.type = 'button';
      del.className = 'mini-del';
      del.title = 'Xoá cột ' + colLabel(j);
      del.setAttribute('aria-label', 'Xoá cột ' + colLabel(j));
      del.textContent = '×';
      th.appendChild(label);
      th.appendChild(badge);
      th.appendChild(del);
      htr.appendChild(th);
    }
    var hfill = document.createElement('th');
    hfill.className = 'filler';
    htr.appendChild(hfill);
    this.thead.innerHTML = '';
    this.thead.appendChild(htr);

    var frag = document.createDocumentFragment();
    for (var i = 0; i < this.data.length; i++) {
      var tr = document.createElement('tr');
      if (i === 0) tr.className = 'header-row';
      var gut = document.createElement('td');
      gut.className = 'gutter';
      gut.dataset.row = i;
      gut.title = 'Click để chọn dòng ' + (i === 0 ? 'tên cột' : i);
      var num = document.createElement('span');
      num.textContent = i === 0 ? 'Cột' : String(i);
      gut.appendChild(num);
      if (i > 0) {
        var rdel = document.createElement('button');
        rdel.type = 'button';
        rdel.className = 'mini-del';
        rdel.title = 'Xoá dòng ' + i;
        rdel.setAttribute('aria-label', 'Xoá dòng ' + i);
        rdel.textContent = '×';
        gut.appendChild(rdel);
      }
      tr.appendChild(gut);

      for (var c = 0; c < cols; c++) {
        var td = document.createElement('td');
        td.className = 'cell';
        td.contentEditable = 'false';
        td.tabIndex = 0;
        td.spellcheck = false;
        td.dataset.r = i;
        td.dataset.c = c;
        td.textContent = this.data[i][c] || '';
        tr.appendChild(td);
      }
      var fill = document.createElement('td');
      fill.className = 'filler';
      tr.appendChild(fill);
      frag.appendChild(tr);
    }
    this.tbody.innerHTML = '';
    this.tbody.appendChild(frag);
    this._applyKeyHighlight();
    this._applySelection();
    window.requestAnimationFrame(function () { self.onChange(); });
  };

  Grid.prototype._applyKeyHighlight = function () {
    var keyCount = (this.mode === 'update' || this.mode === 'select') ? this.whereColumns : 0;
    var heads = this.thead.querySelectorAll('th.col-head');
    this.columnTypes = normalizeColumnTypes(this.columnTypes, heads.length);
    for (var i = 0; i < heads.length; i++) {
      var isKey = i < keyCount;
      heads[i].classList.toggle('is-key', isKey);
      var named = !!(this.data[0] && String(this.data[0][i] || '').trim());
      var badge = heads[i].querySelector('.col-badge');
      var role = !keyCount ? '' : (isKey ? 'WHERE' : (this.mode === 'update' && named ? 'SET' : ''));
      var type = normalizeColumnType(this.columnTypes[i]);
      if (badge) {
        badge.textContent = (role ? role + ' · ' : '') + COLUMN_TYPE_LABELS[type];
        badge.dataset.type = type;
      }
    }
    var cells = this.tbody.querySelectorAll('td.cell');
    for (var n = 0; n < cells.length; n++) {
      cells[n].classList.toggle('is-key', (+cells[n].dataset.c) < keyCount);
    }
  };

  Grid.prototype._cellAt = function (r, c) {
    return this.tbody.querySelector('td.cell[data-r="' + r + '"][data-c="' + c + '"]');
  };

  Grid.prototype._selectionRange = function () {
    return normalizeRange(this.selection || { r1: 0, c1: 0, r2: 0, c2: 0 });
  };

  Grid.prototype._contains = function (r, c) {
    var s = this._selectionRange();
    return r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2;
  };

  Grid.prototype._select = function (r, c, extend) {
    r = Math.max(0, Math.min(r, this.data.length - 1));
    c = Math.max(0, Math.min(c, this.data[0].length - 1));
    if (extend && this.selection) {
      this.selection.r2 = r;
      this.selection.c2 = c;
    } else {
      this.selection = { r1: r, c1: c, r2: r, c2: c };
    }
    this.active = { r: r, c: c };
    this._applySelection();
  };

  Grid.prototype._selectRange = function (r1, c1, r2, c2) {
    this.selection = { r1: r1, c1: c1, r2: r2, c2: c2 };
    this.active = { r: r2, c: c2 };
    this._applySelection();
  };

  Grid.prototype._applySelection = function () {
    if (!this.tbody || !this.data.length) return;
    var old = this.table.querySelectorAll('.is-selected, .is-active, .is-fill-anchor, .is-fill-preview, .is-range-head');
    for (var i = 0; i < old.length; i++) {
      old[i].classList.remove('is-selected', 'is-active', 'is-fill-anchor', 'is-fill-preview', 'is-range-head');
    }
    var s = this._selectionRange();
    for (var r = s.r1; r <= s.r2; r++) {
      for (var c = s.c1; c <= s.c2; c++) {
        var cell = this._cellAt(r, c);
        if (cell) cell.classList.add('is-selected');
      }
    }
    var active = this._cellAt(this.active.r, this.active.c);
    if (active) active.classList.add('is-active');
    var anchor = this._cellAt(s.r2, s.c2);
    if (anchor) anchor.classList.add('is-fill-anchor');
    var heads = this.thead.querySelectorAll('th.col-head');
    for (var h = s.c1; h <= s.c2 && heads[h]; h++) heads[h].classList.add('is-range-head');
    var rows = this.tbody.querySelectorAll('td.gutter[data-row]');
    for (var g = s.r1; g <= s.r2 && rows[g]; g++) rows[g].classList.add('is-range-head');
  };

  Grid.prototype.focusCell = function (r, c, extend) {
    this._endEdit(true);
    this._ensureSize(r + 1, c + 1);
    var cell = this._cellAt(r, c);
    if (!cell) return;
    this._keyboardActive = true;
    this._select(r, c, extend);
    cell.focus();
  };

  Grid.prototype._setCaret = function (cell, selectAll) {
    var range = document.createRange();
    range.selectNodeContents(cell);
    if (!selectAll) range.collapse(false);
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  };

  Grid.prototype._beginEdit = function (cell, options) {
    if (!cell) return;
    options = options || {};
    var r = +cell.dataset.r, c = +cell.dataset.c;
    this._endEdit(true);
    this._select(r, c, false);
    this.editing = true;
    this._editCell = cell;
    this._editOriginal = this.data[r][c] || '';
    cell.contentEditable = 'true';
    cell.classList.add('is-editing');
    cell.focus();
    if (options.replace != null) {
      this._record('edit-' + r + '-' + c, true);
      cell.textContent = options.replace;
      this.data[r][c] = options.replace;
      if (r === 0) this._applyKeyHighlight();
      this.onChange();
    }
    this._setCaret(cell, !!options.selectAll);
  };

  Grid.prototype._endEdit = function (commit) {
    var cell = this._editCell;
    if (!cell) return;
    var r = +cell.dataset.r, c = +cell.dataset.c;
    if (commit === false) {
      cell.textContent = this._editOriginal || '';
      this.data[r][c] = this._editOriginal || '';
    } else {
      this.data[r][c] = cell.textContent.replace(/\u00a0/g, ' ');
    }
    cell.contentEditable = 'false';
    cell.classList.remove('is-editing');
    this.editing = false;
    this._editCell = null;
    this._editOriginal = '';
    if (r === 0) this._applyKeyHighlight();
    this.onChange();
  };

  Grid.prototype._ensureSize = function (rows, cols) {
    var changed = false;
    var width = this.data[0] ? this.data[0].length : 0;
    if (cols > width) {
      for (var i = 0; i < this.data.length; i++) {
        while (this.data[i].length < cols) this.data[i].push('');
      }
      changed = true;
      width = cols;
    }
    while (this.data.length < rows) {
      var row = [];
      for (var j = 0; j < width; j++) row.push('');
      this.data.push(row);
      changed = true;
    }
    if (changed) this.render();
    return changed;
  };

  Grid.prototype._record = function (key, coalesce) {
    var now = Date.now();
    if (coalesce && key === this._lastRecordKey && now - this._lastRecordAt < 700) {
      this._lastRecordAt = now;
      return;
    }
    this.history.push(cloneMatrix(this.data));
    if (this.history.length > HISTORY_LIMIT) this.history.shift();
    this.future = [];
    this._lastRecordKey = key || '';
    this._lastRecordAt = now;
    this._notifyHistory();
  };

  Grid.prototype._notifyHistory = function () {
    this.onHistoryChange({ canUndo: this.history.length > 0, canRedo: this.future.length > 0 });
  };

  Grid.prototype.undo = function () {
    if (!this.history.length) return false;
    this.future.push(cloneMatrix(this.data));
    this.data = this.history.pop();
    this._lastRecordKey = '';
    this._clampSelection();
    this.render();
    this._notifyHistory();
    this.onStatus('Đã hoàn tác thao tác gần nhất.');
    return true;
  };

  Grid.prototype.redo = function () {
    if (!this.future.length) return false;
    this.history.push(cloneMatrix(this.data));
    this.data = this.future.pop();
    this._lastRecordKey = '';
    this._clampSelection();
    this.render();
    this._notifyHistory();
    this.onStatus('Đã làm lại thao tác.');
    return true;
  };

  Grid.prototype._clampSelection = function () {
    var maxR = Math.max(0, this.data.length - 1);
    var maxC = Math.max(0, this.data[0].length - 1);
    this.selection.r1 = Math.min(this.selection.r1, maxR);
    this.selection.r2 = Math.min(this.selection.r2, maxR);
    this.selection.c1 = Math.min(this.selection.c1, maxC);
    this.selection.c2 = Math.min(this.selection.c2, maxC);
    this.active.r = Math.min(this.active.r, maxR);
    this.active.c = Math.min(this.active.c, maxC);
  };

  Grid.prototype.addRows = function (n) {
    this._record('add-rows');
    var width = this.data[0] ? this.data[0].length : MIN_COLS;
    for (var i = 0; i < n; i++) {
      var row = [];
      for (var j = 0; j < width; j++) row.push('');
      this.data.push(row);
    }
    this.render();
  };

  Grid.prototype.addCols = function (n) {
    this._record('add-cols');
    for (var i = 0; i < this.data.length; i++) {
      for (var j = 0; j < n; j++) this.data[i].push('');
    }
    for (var k = 0; k < n; k++) this.columnTypes.push('auto');
    this.render();
    this._notifyColumnTypes();
  };

  Grid.prototype.deleteRow = function (r) {
    if (r <= 0 || this.data.length <= 1) return;
    this._record('delete-row');
    this.data.splice(r, 1);
    while (this.data.length < MIN_ROWS) {
      var row = [];
      for (var i = 0; i < this.data[0].length; i++) row.push('');
      this.data.push(row);
    }
    this._clampSelection();
    this.render();
  };

  Grid.prototype.deleteCol = function (c) {
    var width = this.data[0] ? this.data[0].length : 0;
    if (width <= 1) return;
    this._record('delete-col');
    for (var i = 0; i < this.data.length; i++) this.data[i].splice(c, 1);
    this.columnTypes.splice(c, 1);
    while (this.data[0].length < MIN_COLS) {
      for (var r = 0; r < this.data.length; r++) this.data[r].push('');
      this.columnTypes.push('auto');
    }
    this._clampSelection();
    this.render();
    this._notifyColumnTypes();
  };

  Grid.prototype.clear = function () {
    this._record('clear-all');
    var rows = Math.max(MIN_ROWS, this.data.length);
    var cols = Math.max(MIN_COLS, this.data[0] ? this.data[0].length : 0);
    this.data = [];
    for (var r = 0; r < rows; r++) {
      var row = [];
      for (var c = 0; c < cols; c++) row.push('');
      this.data.push(row);
    }
    this.selection = { r1: 0, c1: 0, r2: 0, c2: 0 };
    this.active = { r: 0, c: 0 };
    this.render();
  };

  Grid.prototype.clearSelection = function () {
    var s = this._selectionRange();
    var changed = false;
    for (var r = s.r1; r <= s.r2; r++) {
      for (var c = s.c1; c <= s.c2; c++) {
        if (this.data[r][c] !== '') changed = true;
      }
    }
    if (!changed) return false;
    this._record('clear-selection');
    for (var y = s.r1; y <= s.r2; y++) {
      for (var x = s.c1; x <= s.c2; x++) this.data[y][x] = '';
    }
    this.render();
    return true;
  };

  Grid.prototype.selectionText = function () {
    var s = this._selectionRange();
    var lines = [];
    for (var r = s.r1; r <= s.r2; r++) {
      var row = [];
      for (var c = s.c1; c <= s.c2; c++) row.push(this.data[r][c]);
      lines.push(row.join('\t'));
    }
    return lines.join('\n');
  };

  Grid.prototype.copySelection = function () {
    var self = this;
    var text = this.selectionText();
    var done = function () { self.onStatus('Đã sao chép vùng dữ liệu.'); };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done, function () { self._fallbackCopy(text, done); });
    } else this._fallbackCopy(text, done);
  };

  Grid.prototype._fallbackCopy = function (text, done) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); done(); }
    catch (err) { this.onStatus('Trình duyệt chặn thao tác sao chép.'); }
    document.body.removeChild(ta);
  };

  Grid.prototype.cutSelection = function () {
    this.copySelection();
    this.clearSelection();
  };

  Grid.prototype.pasteText = function (text, r0, c0) {
    var block = parseClipboard(text);
    if (!block.length) return;
    this._endEdit(true);
    r0 = r0 == null ? this.active.r : r0;
    c0 = c0 == null ? this.active.c : c0;
    var selected = this._selectionRange();
    var needRows = r0 + block.length;
    var needCols = c0;
    var blockCols = 1;
    for (var i = 0; i < block.length; i++) {
      needCols = Math.max(needCols, c0 + block[i].length);
      blockCols = Math.max(blockCols, block[i].length);
    }
    var fillSelected = block.length === 1 && blockCols === 1 &&
      (selected.r2 > selected.r1 || selected.c2 > selected.c1) &&
      this._contains(r0, c0);
    if (fillSelected) {
      r0 = selected.r1;
      c0 = selected.c1;
      needRows = selected.r2 + 1;
      needCols = selected.c2 + 1;
    }
    this._record('paste');
    this._ensureSize(needRows, needCols);
    if (fillSelected) {
      for (var fy = selected.r1; fy <= selected.r2; fy++) {
        for (var fx = selected.c1; fx <= selected.c2; fx++) this.data[fy][fx] = block[0][0];
      }
      this.selection = { r1: selected.r1, c1: selected.c1, r2: selected.r2, c2: selected.c2 };
    } else {
      for (var y = 0; y < block.length; y++) {
        for (var x = 0; x < block[y].length; x++) this.data[r0 + y][c0 + x] = block[y][x];
      }
      this.selection = { r1: r0, c1: c0, r2: r0 + block.length - 1, c2: c0 + blockCols - 1 };
    }
    this.active = { r: r0, c: c0 };
    this.render();
  };

  Grid.prototype.fillDown = function () {
    var s = this._selectionRange();
    if (s.r2 <= s.r1) { this.onStatus('Chọn từ 2 dòng trở lên để fill xuống.'); return false; }
    this._record('fill-down');
    for (var r = s.r1 + 1; r <= s.r2; r++) {
      for (var c = s.c1; c <= s.c2; c++) this.data[r][c] = this.data[s.r1][c];
    }
    this.render();
    this.onStatus('Đã fill dữ liệu xuống.');
    return true;
  };

  Grid.prototype.fillRight = function () {
    var s = this._selectionRange();
    if (s.c2 <= s.c1) { this.onStatus('Chọn từ 2 cột trở lên để fill sang phải.'); return false; }
    this._record('fill-right');
    for (var r = s.r1; r <= s.r2; r++) {
      for (var c = s.c1 + 1; c <= s.c2; c++) this.data[r][c] = this.data[r][s.c1];
    }
    this.render();
    this.onStatus('Đã fill dữ liệu sang phải.');
    return true;
  };

  Grid.prototype._applyDragFill = function (targetR, targetC) {
    var s = this._selectionRange();
    var down = Math.max(0, targetR - s.r2);
    var right = Math.max(0, targetC - s.c2);
    if (!down && !right) return;
    this._record('drag-fill');

    if (down >= right) {
      this._ensureSize(targetR + 1, this.data[0].length);
      var sourceRows = s.r2 - s.r1 + 1;
      for (var c = s.c1; c <= s.c2; c++) {
        var first = this.data[s.r1][c];
        var last = this.data[s.r2][c];
        var series = sourceRows > 1 && isStrictNumber(first) && isStrictNumber(last);
        var step = series ? (Number(last) - Number(first)) / (sourceRows - 1) : 0;
        for (var r = s.r2 + 1; r <= targetR; r++) {
          this.data[r][c] = series
            ? cleanNumber(Number(last) + step * (r - s.r2))
            : this.data[s.r1 + ((r - s.r2 - 1) % sourceRows)][c];
        }
      }
      this.selection = { r1: s.r1, c1: s.c1, r2: targetR, c2: s.c2 };
    } else {
      this._ensureSize(this.data.length, targetC + 1);
      var sourceCols = s.c2 - s.c1 + 1;
      for (var row = s.r1; row <= s.r2; row++) {
        var firstValue = this.data[row][s.c1];
        var lastValue = this.data[row][s.c2];
        var isSeries = sourceCols > 1 && isStrictNumber(firstValue) && isStrictNumber(lastValue);
        var colStep = isSeries ? (Number(lastValue) - Number(firstValue)) / (sourceCols - 1) : 0;
        for (var col = s.c2 + 1; col <= targetC; col++) {
          this.data[row][col] = isSeries
            ? cleanNumber(Number(lastValue) + colStep * (col - s.c2))
            : this.data[row][s.c1 + ((col - s.c2 - 1) % sourceCols)];
        }
      }
      this.selection = { r1: s.r1, c1: s.c1, r2: s.r2, c2: targetC };
    }
    this.active = { r: this.selection.r2, c: this.selection.c2 };
    this.render();
    this.onStatus('Đã fill vùng dữ liệu.');
  };

  Grid.prototype._showFillPreview = function (targetR, targetC) {
    var old = this.tbody.querySelectorAll('.is-fill-preview');
    for (var i = 0; i < old.length; i++) old[i].classList.remove('is-fill-preview');
    var s = this._selectionRange();
    var down = Math.max(0, targetR - s.r2);
    var right = Math.max(0, targetC - s.c2);
    if (!down && !right) return;
    var preview = down >= right
      ? { r1: s.r2 + 1, c1: s.c1, r2: targetR, c2: s.c2 }
      : { r1: s.r1, c1: s.c2 + 1, r2: s.r2, c2: targetC };
    for (var r = preview.r1; r <= preview.r2; r++) {
      for (var c = preview.c1; c <= preview.c2; c++) {
        var cell = this._cellAt(r, c);
        if (cell) cell.classList.add('is-fill-preview');
      }
    }
  };

  Grid.prototype._nearFillHandle = function (cell, event) {
    if (!cell.classList.contains('is-fill-anchor')) return false;
    var rect = cell.getBoundingClientRect();
    return event.clientX >= rect.right - 11 && event.clientY >= rect.bottom - 11;
  };

  Grid.prototype._hideMenu = function () {
    this.menu.hidden = true;
  };

  Grid.prototype._showMenu = function (event, target) {
    this.contextTarget = target;
    this.menu.hidden = false;
    this.menu.style.left = Math.min(event.clientX, window.innerWidth - 220) + 'px';
    this.menu.style.top = Math.min(event.clientY, window.innerHeight - 300) + 'px';
  };

  Grid.prototype._edgeTarget = function (r, c, key) {
    if (key === 'ArrowDown') return { r: this.data.length - 1, c: c };
    if (key === 'ArrowUp') return { r: 0, c: c };
    if (key === 'ArrowRight') return { r: r, c: this.data[0].length - 1 };
    if (key === 'ArrowLeft') return { r: r, c: 0 };
    return { r: r, c: c };
  };

  Grid.prototype._stepTarget = function (r, c, key) {
    if (key === 'ArrowDown') return { r: r + 1, c: c };
    if (key === 'ArrowUp') return { r: Math.max(0, r - 1), c: c };
    if (key === 'ArrowRight') return { r: r, c: c + 1 };
    if (key === 'ArrowLeft') return { r: r, c: Math.max(0, c - 1) };
    return { r: r, c: c };
  };

  Grid.prototype._isPrintableKey = function (e) {
    return e.key && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey;
  };

  Grid.prototype._handleKeyDown = function (e, td) {
    if (!td || !td.classList || !td.classList.contains('cell')) return false;
    var r = +td.dataset.r, c = +td.dataset.c;
    var width = this.data[0].length;
    var command = e.ctrlKey || e.metaKey;
    var key = e.key.toLowerCase();

    if (this.editing) {
      if (e.key === 'Escape') {
        e.preventDefault();
        this._endEdit(false);
        td.focus();
        return true;
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        this._endEdit(true);
        this.focusCell(e.shiftKey ? Math.max(0, r - 1) : r + 1, c);
        return true;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        this._endEdit(true);
        if (e.shiftKey) {
          if (c > 0) this.focusCell(r, c - 1);
          else if (r > 0) this.focusCell(r - 1, width - 1);
        } else {
          if (c < width - 1) this.focusCell(r, c + 1);
          else this.focusCell(r + 1, 0);
        }
        return true;
      }
      if (/^Arrow(Down|Up|Left|Right)$/.test(e.key) && !e.altKey) {
        e.preventDefault();
        this._endEdit(true);
        var editTarget = command ? this._edgeTarget(r, c, e.key) : this._stepTarget(r, c, e.key);
        this.focusCell(editTarget.r, editTarget.c, e.shiftKey);
        return true;
      }
      return false;
    }

    if (command && key === 'z') {
      e.preventDefault();
      if (e.shiftKey) this.redo(); else this.undo();
      return true;
    }
    if (command && key === 'y') {
      e.preventDefault(); this.redo(); return true;
    }
    if (command && key === 'c') {
      e.preventDefault(); this.copySelection(); return true;
    }
    if (command && key === 'x') {
      e.preventDefault(); this.cutSelection(); return true;
    }
    if (command && key === 'd') {
      e.preventDefault(); this.fillDown(); return true;
    }
    if (command && key === 'r') {
      e.preventDefault(); this.fillRight(); return true;
    }
    if (e.key === 'Delete') {
      e.preventDefault(); this.clearSelection(); return true;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      if (e.shiftKey) {
        if (c > 0) this.focusCell(r, c - 1);
        else if (r > 0) this.focusCell(r - 1, width - 1);
      } else {
        if (c < width - 1) this.focusCell(r, c + 1);
        else this.focusCell(r + 1, 0);
      }
      return true;
    }
    if (e.key === 'Enter') {
      e.preventDefault();
      this._beginEdit(td, { selectAll: true });
      return true;
    }
    if (/^Arrow(Down|Up|Left|Right)$/.test(e.key) && !e.altKey) {
      e.preventDefault();
      var target = command ? this._edgeTarget(r, c, e.key) : this._stepTarget(r, c, e.key);
      this.focusCell(target.r, target.c, e.shiftKey);
      return true;
    }
    if (e.key === 'F2') {
      e.preventDefault();
      this._beginEdit(td);
      return true;
    }
    if (this._isPrintableKey(e)) {
      e.preventDefault();
      this._beginEdit(td, { replace: e.key });
      return true;
    }
    if (e.key === 'Escape') {
      td.blur();
      return true;
    }
    return false;
  };

  Grid.prototype._runMenuAction = function (action) {
    var self = this;
    var target = this.contextTarget || this.active;
    this._hideMenu();
    if (action === 'copy') this.copySelection();
    else if (action === 'cut') this.cutSelection();
    else if (action === 'clear') this.clearSelection();
    else if (action === 'fill-down') this.fillDown();
    else if (action === 'fill-right') this.fillRight();
    else if (action === 'delete-row') this.deleteRow(target.r);
    else if (action === 'delete-col') this.deleteCol(target.c);
    else if (action === 'col-type-auto') this.setColumnType(target.c, 'auto');
    else if (action === 'col-type-text') this.setColumnType(target.c, 'text');
    else if (action === 'col-type-number') this.setColumnType(target.c, 'number');
    else if (action === 'paste') {
      if (!navigator.clipboard || !navigator.clipboard.readText) {
        this.onStatus('Hãy dùng Ctrl+V để dán dữ liệu.');
        return;
      }
      navigator.clipboard.readText().then(function (text) {
        self.pasteText(text, target.r, target.c);
      }).catch(function () { self.onStatus('Hãy dùng Ctrl+V để dán dữ liệu.'); });
    }
  };

  Grid.prototype._bind = function () {
    var self = this;

    this.tbody.addEventListener('beforeinput', function (e) {
      var td = e.target;
      if (!td.classList || !td.classList.contains('cell')) return;
      if (!self.editing) return;
      self._record('edit-' + td.dataset.r + '-' + td.dataset.c, true);
    });

    this.tbody.addEventListener('input', function (e) {
      var td = e.target;
      if (!td.classList || !td.classList.contains('cell')) return;
      var r = +td.dataset.r, c = +td.dataset.c;
      self.data[r][c] = td.textContent.replace(/\u00a0/g, ' ');
      if (r === 0) self._applyKeyHighlight();
      self.onChange();
    });

    this.tbody.addEventListener('focusin', function (e) {
      var td = e.target;
      if (!td.classList || !td.classList.contains('cell')) return;
      self._keyboardActive = true;
      var r = +td.dataset.r, c = +td.dataset.c;
      self.active = { r: r, c: c };
      if (!self.editing && !self._contains(r, c)) self._select(r, c, false);
      else self._applySelection();
    });

    this.tbody.addEventListener('focusout', function (e) {
      var td = e.target;
      if (td === self._editCell) self._endEdit(true);
    });

    this.tbody.addEventListener('mousedown', function (e) {
      var td = e.target.closest ? e.target.closest('td.cell') : null;
      if (!td || e.button !== 0) return;
      self._keyboardActive = true;
      var r = +td.dataset.r, c = +td.dataset.c;
      if (self._nearFillHandle(td, e)) {
        e.preventDefault();
        self._filling = true;
        self._fillTarget = { r: r, c: c };
        self.root.classList.add('is-filling');
        return;
      }
      if (e.detail > 1) {
        e.preventDefault();
        self._beginEdit(td);
        return;
      }
      self._endEdit(true);
      e.preventDefault();
      if (e.shiftKey) self._select(r, c, true);
      else self._select(r, c, false);
      td.focus();
      self._draggingSelection = true;
    });

    this.tbody.addEventListener('dblclick', function (e) {
      var td = e.target.closest ? e.target.closest('td.cell') : null;
      if (!td) return;
      e.preventDefault();
      self._beginEdit(td);
    });

    this.tbody.addEventListener('mouseover', function (e) {
      var td = e.target.closest ? e.target.closest('td.cell') : null;
      if (!td) return;
      var r = +td.dataset.r, c = +td.dataset.c;
      if (self._filling) {
        self._fillTarget = { r: r, c: c };
        self._showFillPreview(r, c);
      } else if (self._draggingSelection && e.buttons === 1) {
        self.selection.r2 = r;
        self.selection.c2 = c;
        self.active = { r: r, c: c };
        self._applySelection();
      }
    });

    document.addEventListener('mouseup', function () {
      self._draggingSelection = false;
      if (!self._filling) return;
      self._filling = false;
      self.root.classList.remove('is-filling');
      if (self._fillTarget) self._applyDragFill(self._fillTarget.r, self._fillTarget.c);
      self._fillTarget = null;
    });

    this.tbody.addEventListener('keydown', function (e) {
      self._handleKeyDown(e, e.target);
    });

    this.tbody.addEventListener('paste', function (e) {
      var td = e.target;
      if (!td.classList || !td.classList.contains('cell')) return;
      var text = (e.clipboardData || window.clipboardData).getData('text/plain');
      e.preventDefault();
      self.pasteText(text, +td.dataset.r, +td.dataset.c);
    });

    this.tbody.addEventListener('click', function (e) {
      var btn = e.target.closest ? e.target.closest('.mini-del') : null;
      if (btn) {
        var tr = btn.closest('tr');
        self.deleteRow(tr.sectionRowIndex);
        return;
      }
      var gut = e.target.closest ? e.target.closest('td.gutter[data-row]') : null;
      if (gut) {
        self._keyboardActive = true;
        var row = +gut.dataset.row;
        self._selectRange(row, 0, row, self.data[0].length - 1);
        var rowCell = self._cellAt(row, 0);
        if (rowCell) rowCell.focus();
      }
    });

    this.thead.addEventListener('click', function (e) {
      var btn = e.target.closest ? e.target.closest('.mini-del') : null;
      if (btn) {
        var th = btn.closest('th');
        self.deleteCol(+th.dataset.col);
        return;
      }
      var typeBadge = e.target.closest ? e.target.closest('.col-badge') : null;
      if (typeBadge) {
        e.preventDefault();
        self.toggleColumnType(+typeBadge.closest('th').dataset.col);
        return;
      }
      var head = e.target.closest ? e.target.closest('th.col-head') : null;
      if (head) {
        self._keyboardActive = true;
        var col = +head.dataset.col;
        self._selectRange(0, col, self.data.length - 1, col);
        var colCell = self._cellAt(0, col);
        if (colCell) colCell.focus();
      }
      else if (e.target.closest && e.target.closest('th.corner')) {
        self._keyboardActive = true;
        self._selectRange(0, 0, self.data.length - 1, self.data[0].length - 1);
        var firstCell = self._cellAt(0, 0);
        if (firstCell) firstCell.focus();
      }
    });

    this.table.addEventListener('contextmenu', function (e) {
      var cell = e.target.closest ? e.target.closest('td.cell') : null;
      var head = e.target.closest ? e.target.closest('th.col-head') : null;
      var gut = e.target.closest ? e.target.closest('td.gutter[data-row]') : null;
      if (!cell && !head && !gut) return;
      e.preventDefault();
      self._keyboardActive = true;
      var target;
      if (cell) {
        target = { r: +cell.dataset.r, c: +cell.dataset.c };
        if (!self._contains(target.r, target.c)) self._select(target.r, target.c, false);
      } else if (head) {
        target = { r: self.active.r, c: +head.dataset.col };
        self._selectRange(0, target.c, self.data.length - 1, target.c);
      } else {
        target = { r: +gut.dataset.row, c: self.active.c };
        self._selectRange(target.r, 0, target.r, self.data[0].length - 1);
      }
      self._showMenu(e, target);
    });

    this.menu.addEventListener('click', function (e) {
      var btn = e.target.closest ? e.target.closest('button[data-action]') : null;
      if (btn) self._runMenuAction(btn.dataset.action);
    });
    document.addEventListener('mousedown', function (e) {
      if (!self.menu.hidden && !self.menu.contains(e.target)) self._hideMenu();
      if (!self.root.contains(e.target) && !self.menu.contains(e.target)) {
        self._keyboardActive = false;
        self._endEdit(true);
      }
    });
    document.addEventListener('keydown', function (e) {
      if (e.defaultPrevented || !self._keyboardActive) return;
      var target = e.target;
      if (target && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT|BUTTON)$/.test(target.tagName))) return;
      var cell = self._cellAt(self.active.r, self.active.c);
      self._handleKeyDown(e, cell);
    });
    document.addEventListener('paste', function (e) {
      if (e.defaultPrevented || !self._keyboardActive || self.editing) return;
      var target = e.target;
      if (target && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName))) return;
      var text = (e.clipboardData || window.clipboardData).getData('text/plain');
      if (!text) return;
      e.preventDefault();
      self.pasteText(text, self.active.r, self.active.c);
    });
    window.addEventListener('blur', function () { self._hideMenu(); });
    window.addEventListener('resize', function () { self._hideMenu(); });
  };

  Grid.parseClipboard = parseClipboard;
  global.Grid = Grid;
})(window);
