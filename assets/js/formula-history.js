/*!
 * formula-history.js — Lịch sử và mục yêu thích của Formula Helper.
 * Lưu trên máy theo deviceId, giới hạn số bản ghi gần nhất. §99–§105
 */
(function (global) {
  'use strict';

  var HISTORY_LIMIT = 50;
  var historyKey = 'genscript.formulaHistory.v1';
  var favoriteKey = 'genscript.formulaFavorites.v1';
  var settingsKey = 'genscript.formulaSettings.v1';

  var store = null;
  try {
    var probe = '__fh_test__';
    global.localStorage.setItem(probe, '1');
    global.localStorage.removeItem(probe);
    store = global.localStorage;
  } catch (err) { store = null; }

  function read(key, fallback) {
    if (!store) return fallback;
    try {
      var raw = store.getItem(key);
      if (!raw) return fallback;
      var parsed = JSON.parse(raw);
      return parsed == null ? fallback : parsed;
    } catch (err) { return fallback; }
  }

  function write(key, value) {
    if (!store) return false;
    try { store.setItem(key, JSON.stringify(value)); return true; }
    catch (err) { return false; }
  }

  function scopeTo(deviceId) {
    if (!deviceId) return;
    historyKey = 'genscript.formulaHistory.v1.' + deviceId;
    favoriteKey = 'genscript.formulaFavorites.v1.' + deviceId;
    settingsKey = 'genscript.formulaSettings.v1.' + deviceId;
  }

  function newId() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  /* ------------------------------------------------------------ history */

  function getHistory() {
    var list = read(historyKey, []);
    return Array.isArray(list) ? list : [];
  }

  /**
   * Ghi một mục lịch sử. §101
   * item = { type, title, platform, formula, separator, config }
   */
  function pushHistory(item) {
    if (!item || !item.formula) return null;
    var list = getHistory();
    // Không ghi trùng liên tiếp cùng một công thức
    if (list.length && list[0].formula === item.formula) {
      list[0].timestamp = Date.now();
      write(historyKey, list);
      return list[0];
    }
    var entry = {
      id: newId(),
      timestamp: Date.now(),
      type: item.type || '',
      title: item.title || '',
      subtitle: item.subtitle || '',
      platform: item.platform || 'excel',
      separator: item.separator || ',',
      formula: item.formula,
      config: item.config || null
    };
    list.unshift(entry);
    if (list.length > HISTORY_LIMIT) list.length = HISTORY_LIMIT;
    write(historyKey, list);
    return entry;
  }

  function removeHistory(id) {
    write(historyKey, getHistory().filter(function (h) { return h.id !== id; }));
  }

  function clearHistory() { write(historyKey, []); }

  function findHistory(id) {
    var list = getHistory();
    for (var i = 0; i < list.length; i++) if (list[i].id === id) return list[i];
    return null;
  }

  /* ------------------------------------------------------------ favorites */

  function getFavorites() {
    var list = read(favoriteKey, []);
    return Array.isArray(list) ? list : [];
  }

  /** Lưu cấu hình thành mục yêu thích có tên riêng. §103 §105 */
  function addFavorite(item) {
    if (!item) return null;
    var list = getFavorites();
    var entry = {
      id: newId(),
      name: item.name || item.title || 'Công thức',
      type: item.type || '',
      platform: item.platform || 'excel',
      separator: item.separator || ',',
      formula: item.formula || '',
      subtitle: item.subtitle || '',
      config: item.config || null,
      createdAt: Date.now()
    };
    list.unshift(entry);
    write(favoriteKey, list);
    return entry;
  }

  function removeFavorite(id) {
    write(favoriteKey, getFavorites().filter(function (f) { return f.id !== id; }));
  }

  function renameFavorite(id, name) {
    var list = getFavorites();
    list.forEach(function (f) { if (f.id === id) f.name = name; });
    write(favoriteKey, list);
  }

  function findFavorite(id) {
    var list = getFavorites();
    for (var i = 0; i < list.length; i++) if (list[i].id === id) return list[i];
    return null;
  }

  /* ------------------------------------------------------------ settings §163 */

  var DEFAULT_SETTINGS = {
    platform: 'excel',
    separator: 'auto',
    excelVersion: 'modern',
    rangeMode: 'full',
    referenceType: 'relative',
    dataRow: 2,
    formatView: 'single'
  };

  function getSettings() {
    var saved = read(settingsKey, {});
    var out = {};
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
      out[k] = saved && saved[k] !== undefined ? saved[k] : DEFAULT_SETTINGS[k];
    });
    return out;
  }

  function saveSettings(patch) {
    var next = getSettings();
    Object.keys(patch || {}).forEach(function (k) { next[k] = patch[k]; });
    write(settingsKey, next);
    return next;
  }

  global.FormulaStore = {
    scopeTo: scopeTo,
    getHistory: getHistory,
    pushHistory: pushHistory,
    removeHistory: removeHistory,
    clearHistory: clearHistory,
    findHistory: findHistory,
    getFavorites: getFavorites,
    addFavorite: addFavorite,
    removeFavorite: removeFavorite,
    renameFavorite: renameFavorite,
    findFavorite: findFavorite,
    getSettings: getSettings,
    saveSettings: saveSettings,
    HISTORY_LIMIT: HISTORY_LIMIT
  };
})(typeof window !== 'undefined' ? window : this);
