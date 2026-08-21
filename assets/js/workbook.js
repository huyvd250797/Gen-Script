/*!
 * workbook.js — Kho dữ liệu Workbook dùng chung cho Gen Script và Formula Helper.
 *
 * Module này KHÔNG sở hữu bản sao dữ liệu. Nó gắn (bind) vào state đang có của
 * app.js để tránh nhân đôi dataset (yêu cầu hiệu năng §170), rồi cung cấp một
 * API đọc ổn định cho Formula Helper.
 */
(function (global) {
  'use strict';

  var host = null;          // { sheets, activeSheet } — tham chiếu sống tới state của app
  var fileName = '';
  var listeners = [];

  function noSheets() { return []; }

  function sheets() {
    return host && host.sheets ? host.sheets : noSheets();
  }

  function activeIndex() {
    if (!host) return 0;
    var n = sheets().length;
    if (!n) return 0;
    return Math.max(0, Math.min(host.activeSheet | 0, n - 1));
  }

  /* ------------------------------------------------------------ cột / header */

  function columnLetter(index) {
    var i = (index | 0) + 1;
    var out = '';
    while (i > 0) {
      var rem = (i - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      i = Math.floor((i - 1) / 26);
    }
    return out || 'A';
  }

  function letterToIndex(letter) {
    var s = String(letter || '').toUpperCase().replace(/[^A-Z]/g, '');
    if (!s) return -1;
    var n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }

  function sheetAt(index) {
    var list = sheets();
    if (index == null) index = activeIndex();
    return list[index] || null;
  }

  function sheetName(index) {
    var s = sheetAt(index);
    return s ? String(s.name || '') : '';
  }

  /** Danh sách tiêu đề cột của một sheet (dòng đầu tiên của lưới). */
  function headers(index) {
    var s = sheetAt(index);
    if (!s || !s.data || !s.data.length) return [];
    var row = s.data[0] || [];
    var width = 0;
    for (var r = 0; r < s.data.length; r++) {
      width = Math.max(width, (s.data[r] || []).length);
    }
    var out = [];
    for (var c = 0; c < width; c++) {
      var raw = row[c] == null ? '' : String(row[c]).trim();
      out.push({
        index: c,
        letter: columnLetter(c),
        name: raw || ('Cột ' + columnLetter(c)),
        named: !!raw,
        type: normalizeType(s.columnTypes && s.columnTypes[c])
      });
    }
    return out;
  }

  function normalizeType(t) {
    return t === 'text' || t === 'number' ? t : 'auto';
  }

  /** Số dòng dữ liệu thật (không tính dòng header, bỏ dòng trống cuối). */
  function dataRowCount(index) {
    var s = sheetAt(index);
    if (!s || !s.data) return 0;
    var last = 0;
    for (var r = 1; r < s.data.length; r++) {
      var row = s.data[r] || [];
      for (var c = 0; c < row.length; c++) {
        if (String(row[c] == null ? '' : row[c]).trim() !== '') { last = r; break; }
      }
    }
    return last; // last là chỉ số dòng cuối có dữ liệu, cũng chính là số dòng dữ liệu
  }

  /** Lấy giá trị một cột (không gồm header) để phục vụ Result Preview. */
  function columnValues(sheetIndex, colIndex, limit) {
    var s = sheetAt(sheetIndex);
    if (!s || !s.data) return [];
    var out = [];
    var max = limit ? limit + 1 : s.data.length;
    for (var r = 1; r < s.data.length && out.length < (limit || Infinity); r++) {
      var row = s.data[r] || [];
      out.push(row[colIndex] == null ? '' : String(row[colIndex]));
    }
    return out;
  }

  /** Toàn bộ cột của mọi sheet, dạng phẳng — dùng cho dropdown Sheet.Column. */
  function allFields() {
    var out = [];
    sheets().forEach(function (sheet, si) {
      headers(si).forEach(function (h) {
        out.push({
          sheetIndex: si,
          sheetName: String(sheet.name || ''),
          colIndex: h.index,
          letter: h.letter,
          column: h.name,
          type: h.type,
          label: String(sheet.name || '') + '.' + h.name,
          value: si + ':' + h.index
        });
      });
    });
    return out;
  }

  function parseFieldValue(value) {
    var parts = String(value || '').split(':');
    if (parts.length !== 2) return null;
    var si = parseInt(parts[0], 10);
    var ci = parseInt(parts[1], 10);
    if (isNaN(si) || isNaN(ci)) return null;
    return { sheetIndex: si, colIndex: ci };
  }

  /** Kiểm tra một tham chiếu (sheet/cột) còn hợp lệ trong workbook hiện tại. */
  function validateField(ref) {
    if (!ref) return { ok: false, message: 'Chưa chọn trường dữ liệu.' };
    if (ref.mode && ref.mode !== 'column') return { ok: true };
    var s = sheetAt(ref.sheetIndex);
    if (!s) {
      return {
        ok: false,
        message: 'Sheet ' + (ref.sheetName || '') + ' không còn tồn tại trong Workbook hiện tại.'
      };
    }
    var hs = headers(ref.sheetIndex);
    if (!hs[ref.colIndex]) {
      return {
        ok: false,
        message: 'Cột ' + (ref.column || '') + ' không tồn tại trong Sheet ' + s.name + '.'
      };
    }
    return { ok: true };
  }

  function isEmpty() {
    var list = sheets();
    if (!list.length) return true;
    if (list.length === 1) {
      var s = list[0];
      if (s.seed) return true;
      var hs = headers(0);
      var named = hs.some(function (h) { return h.named; });
      if (!named && dataRowCount(0) === 0) return true;
    }
    return false;
  }

  /* ------------------------------------------------------------ pub/sub */

  function subscribe(fn) {
    if (typeof fn === 'function') listeners.push(fn);
    return function () {
      listeners = listeners.filter(function (f) { return f !== fn; });
    };
  }

  function notify(reason) {
    for (var i = 0; i < listeners.length; i++) {
      try { listeners[i](reason || 'change'); } catch (err) { /* không chặn app */ }
    }
  }

  global.Workbook = {
    bind: function (state) { host = state; },
    setFileName: function (name) { fileName = String(name || ''); },
    getFileName: function () { return fileName; },
    getSheets: sheets,
    getSheet: sheetAt,
    getSheetName: sheetName,
    getActiveIndex: activeIndex,
    getHeaders: headers,
    getDataRowCount: dataRowCount,
    getColumnValues: columnValues,
    getAllFields: allFields,
    parseFieldValue: parseFieldValue,
    validateField: validateField,
    isEmpty: isEmpty,
    columnLetter: columnLetter,
    letterToIndex: letterToIndex,
    subscribe: subscribe,
    notify: notify
  };
})(typeof window !== 'undefined' ? window : this);
