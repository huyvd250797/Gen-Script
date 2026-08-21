/*!
 * formula-explainer.js — Giải thích công thức và mô phỏng kết quả trên vài dòng.
 *
 * KHÔNG phải Excel Calculation Engine (§97). Chỉ tính được các loại công thức
 * đơn giản mà JavaScript mô phỏng được: lookup, duplicate, existence, compare,
 * blank, text, concat, sqlIn. §95 §96
 */
(function (global) {
  'use strict';

  var PREVIEW_ROWS = 15;
  var SUPPORTED = ['lookup', 'duplicate', 'existence', 'compare', 'listCompare', 'blank', 'text', 'concat', 'sqlIn', 'sqlQuote'];

  function canPreview(type) {
    return SUPPORTED.indexOf(type) !== -1;
  }

  function norm(v) {
    return String(v == null ? '' : v).trim();
  }

  function colValues(wb, ref, limit) {
    if (!ref || ref.mode !== 'column') return null;
    return wb.getColumnValues(ref.sheetIndex, ref.colIndex, limit);
  }

  function labelOf(ref) {
    if (!ref) return '';
    if (ref.mode !== 'column') return String(ref.raw || '');
    return ref.column || '';
  }

  /**
   * Mô phỏng kết quả.
   * @returns {{columns:string[], rows:string[][], note:string}|null}
   */
  function preview(config, wb) {
    if (!config || !canPreview(config.type) || !wb) return null;
    try {
      switch (config.type) {
        case 'lookup': return previewLookup(config, wb);
        case 'duplicate': return previewDuplicate(config, wb);
        case 'existence':
        case 'listCompare': return previewExistence(config, wb);
        case 'compare': return previewCompare(config, wb);
        case 'blank': return previewBlank(config, wb);
        case 'text': return previewText(config, wb);
        case 'concat': return previewConcat(config, wb);
        case 'sqlQuote': return previewSqlQuote(config, wb);
        case 'sqlIn': return previewSqlIn(config, wb);
        default: return null;
      }
    } catch (err) {
      return null;
    }
  }

  function previewLookup(cfg, wb) {
    if (cfg.method === 'vlookup' || cfg.method === 'index-match' || (cfg.conditions || []).length) {
      // vẫn mô phỏng được vì bản chất là dò khớp chính xác
    }
    var values = colValues(wb, cfg.lookupValue, PREVIEW_ROWS);
    var keys = colValues(wb, cfg.lookupRange);
    var returns = colValues(wb, cfg.returnRange);
    if (!values || !keys || !returns) return null;

    var map = {};
    for (var i = 0; i < keys.length; i++) {
      var k = norm(keys[i]);
      if (k !== '' && !(k in map)) map[k] = returns[i] == null ? '' : returns[i];
    }
    var fallback = cfg.notFound || '';
    var rows = values.filter(function (v) { return norm(v) !== ''; }).map(function (v) {
      var k = norm(v);
      return [v, k in map ? String(map[k]) : (fallback || '(rỗng)')];
    });
    return {
      columns: [labelOf(cfg.lookupValue), labelOf(cfg.returnRange)],
      rows: rows,
      note: 'Mô phỏng dò khớp chính xác trên ' + rows.length + ' dòng đầu.'
    };
  }

  function previewDuplicate(cfg, wb) {
    var fields = (cfg.fields || []).filter(function (f) { return f && f.mode === 'column'; });
    if (!fields.length) return null;
    var cols = fields.map(function (f) { return colValues(wb, f); });
    if (cols.some(function (c) { return !c; })) return null;

    var total = cols[0].length;
    var counts = {};
    for (var r = 0; r < total; r++) {
      var key = cols.map(function (c) { return norm(c[r]); }).join('\u0001');
      counts[key] = (counts[key] || 0) + 1;
    }
    var rows = [];
    for (var i = 0; i < total && rows.length < PREVIEW_ROWS; i++) {
      var vals = cols.map(function (c) { return c[i]; });
      if (vals.every(function (v) { return norm(v) === ''; })) continue;
      var k = vals.map(norm).join('\u0001');
      var n = counts[k] || 0;
      rows.push(vals.concat([
        cfg.resultType === 'count' ? String(n)
          : (n > 1 ? (cfg.duplicateLabel || 'TRÙNG') : (cfg.uniqueLabel || '(rỗng)'))
      ]));
    }
    var dupCount = 0;
    Object.keys(counts).forEach(function (k) { if (counts[k] > 1) dupCount += counts[k]; });
    return {
      columns: fields.map(labelOf).concat(['Kết quả']),
      rows: rows,
      note: 'Toàn bộ dữ liệu có ' + dupCount + ' dòng bị trùng.'
    };
  }

  function previewExistence(cfg, wb) {
    var values = colValues(wb, cfg.value, PREVIEW_ROWS);
    var refs = colValues(wb, cfg.reference);
    if (!values || !refs) return null;
    var set = {};
    refs.forEach(function (v) { var k = norm(v); if (k) set[k] = true; });

    var foundLabel = cfg.type === 'listCompare'
      ? (cfg.matchLabel || 'CÓ') : (cfg.foundLabel || 'CÓ');
    var missLabel = cfg.type === 'listCompare'
      ? (cfg.missLabel || 'THIẾU') : (cfg.missingLabel || 'KHÔNG');

    var missing = 0;
    var rows = values.filter(function (v) { return norm(v) !== ''; }).map(function (v) {
      var ok = !!set[norm(v)];
      if (!ok) missing++;
      return [v, ok ? foundLabel : missLabel];
    });
    return {
      columns: [labelOf(cfg.value), 'Kết quả'],
      rows: rows,
      note: missing ? missing + ' giá trị trong nhóm xem trước không có ở danh mục đối chiếu.' : 'Các giá trị xem trước đều tồn tại.'
    };
  }

  function previewCompare(cfg, wb) {
    var a = colValues(wb, cfg.valueA, PREVIEW_ROWS);
    var b = colValues(wb, cfg.valueB, PREVIEW_ROWS);
    if (!a || !b) return null;
    var rows = [];
    for (var i = 0; i < a.length && rows.length < PREVIEW_ROWS; i++) {
      if (norm(a[i]) === '' && norm(b[i]) === '') continue;
      var x = cfg.trim ? norm(a[i]) : String(a[i] || '');
      var y = cfg.trim ? norm(b[i]) : String(b[i] || '');
      if (cfg.ignoreCase) { x = x.toLowerCase(); y = y.toLowerCase(); }
      rows.push([a[i], b[i], x === y ? (cfg.sameLabel || 'GIỐNG') : (cfg.diffLabel || 'KHÁC')]);
    }
    return {
      columns: [labelOf(cfg.valueA), labelOf(cfg.valueB), 'Kết quả'],
      rows: rows,
      note: ''
    };
  }

  function previewBlank(cfg, wb) {
    var fields = (cfg.fields || []).filter(function (f) { return f && f.mode === 'column'; });
    if (!fields.length) return null;
    var cols = fields.map(function (f) { return colValues(wb, f, PREVIEW_ROWS); });
    if (cols.some(function (c) { return !c; })) return null;
    var rows = [];
    for (var i = 0; i < cols[0].length && rows.length < PREVIEW_ROWS; i++) {
      var vals = cols.map(function (c) { return c[i]; });
      if (vals.every(function (v) { return norm(v) === ''; })) continue;
      var missing = vals.some(function (v) { return norm(v) === ''; });
      rows.push(vals.concat([missing ? (cfg.missingLabel || 'THIẾU DỮ LIỆU') : (cfg.okLabel || '(rỗng)')]));
    }
    return { columns: fields.map(labelOf).concat(['Kết quả']), rows: rows, note: '' };
  }

  function applyText(value, cfg) {
    var s = String(value == null ? '' : value);
    switch (cfg.preset) {
      case 'trim': return s.trim().replace(/\s+/g, ' ');
      case 'clean': return s.replace(/[\x00-\x1F\x7F]/g, '');
      case 'trimClean': return s.replace(/[\x00-\x1F\x7F]/g, '').trim().replace(/\s+/g, ' ');
      case 'upper': return s.trim().toUpperCase();
      case 'lower': return s.trim().toLowerCase();
      case 'proper': return s.trim().toLowerCase().replace(/(^|\s)\S/g, function (c) { return c.toUpperCase(); });
      case 'removeSpace': return s.replace(/ /g, '');
      case 'removeLineBreak': return s.replace(/\n/g, ' ');
      case 'substitute': return cfg.find ? s.split(cfg.find).join(cfg.replace == null ? '' : cfg.replace) : s;
      case 'left': return s.slice(0, parseInt(cfg.count, 10) || 1);
      case 'right': return s.slice(-(parseInt(cfg.count, 10) || 1));
      case 'mid': return s.substr((parseInt(cfg.start, 10) || 1) - 1, parseInt(cfg.count, 10) || 1);
      case 'len': return String(s.length);
      default: return s;
    }
  }

  function previewText(cfg, wb) {
    var values = colValues(wb, cfg.field, PREVIEW_ROWS);
    if (!values) return null;
    var rows = values.filter(function (v) { return norm(v) !== ''; }).map(function (v) {
      return [v, applyText(v, cfg)];
    });
    return { columns: [labelOf(cfg.field), 'Sau khi xử lý'], rows: rows, note: '' };
  }

  var SEP_PRESETS = { space: ' ', dash: '-', slash: '/', underscore: '_', comma: ', ', none: '' };

  function previewConcat(cfg, wb) {
    var fields = (cfg.fields || []).filter(function (f) { return f && f.mode === 'column'; });
    if (fields.length < 2) return null;
    var cols = fields.map(function (f) { return colValues(wb, f, PREVIEW_ROWS); });
    if (cols.some(function (c) { return !c; })) return null;
    var sep = cfg.separatorText != null && cfg.separatorText !== ''
      ? cfg.separatorText
      : (SEP_PRESETS[cfg.joinWith] != null ? SEP_PRESETS[cfg.joinWith] : ' ');
    var rows = [];
    for (var i = 0; i < cols[0].length && rows.length < PREVIEW_ROWS; i++) {
      var vals = cols.map(function (c) { return c[i]; });
      if (vals.every(function (v) { return norm(v) === ''; })) continue;
      var parts = cfg.ignoreBlank === false ? vals : vals.filter(function (v) { return norm(v) !== ''; });
      rows.push(vals.concat([parts.join(sep)]));
    }
    return { columns: fields.map(labelOf).concat(['Kết quả ghép']), rows: rows, note: '' };
  }

  function previewSqlQuote(cfg, wb) {
    var values = colValues(wb, cfg.field, PREVIEW_ROWS);
    if (!values) return null;
    var rows = values.filter(function (v) { return norm(v) !== ''; }).map(function (v) {
      return [v, cfg.dataType === 'number' ? norm(v) : "'" + norm(v).replace(/'/g, "''") + "'"];
    });
    return { columns: [labelOf(cfg.field), 'Giá trị SQL'], rows: rows, note: '' };
  }

  /** SQL IN có thể tạo trực tiếp từ dữ liệu, không cần dán công thức. §84 */
  function sqlInValues(cfg, wb) {
    var values = colValues(wb, cfg.field);
    if (!values) return null;
    var seen = {};
    var out = [];
    values.forEach(function (v) {
      var s = norm(v);
      if (s === '' || seen[s]) return;
      seen[s] = true;
      out.push(cfg.dataType === 'number' ? s : "'" + s.replace(/'/g, "''") + "'");
    });
    return out;
  }

  function previewSqlIn(cfg, wb) {
    var list = sqlInValues(cfg, wb);
    if (!list) return null;
    var column = cfg.sqlColumn || labelOf(cfg.field) || 'Column';
    var joined = list.join(',');
    return {
      columns: ['Giá trị'],
      rows: list.slice(0, PREVIEW_ROWS).map(function (v) { return [v]; }),
      note: list.length + ' giá trị duy nhất.',
      directOutput: {
        values: joined,
        whereIn: 'WHERE ' + column + ' IN (' + joined + ')'
      }
    };
  }

  global.FormulaExplainer = {
    canPreview: canPreview,
    preview: preview,
    sqlInValues: sqlInValues,
    PREVIEW_ROWS: PREVIEW_ROWS
  };
})(typeof window !== 'undefined' ? window : this);
