/*!
 * formula-engine.js — Sinh công thức Excel / Google Sheets từ FormulaConfig.
 *
 * Module thuần: không chạm DOM, không phụ thuộc generator.js.
 * API chính:  FormulaEngine.generate(config) -> { formula, formattedFormula, explanation, warnings }
 */
(function (global) {
  'use strict';

  var SEPARATORS = { comma: ',', semicolon: ';' };

  /* =============================================================
   * 1. Tham chiếu: cột, sheet, ô
   * ========================================================== */

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

  var CELL_LIKE = /^(\$?[A-Za-z]{1,3}\$?\d{1,7})$/;

  /**
   * Sheet name cần bọc nháy đơn khi có khoảng trắng, dấu -, ký tự đặc biệt,
   * bắt đầu bằng số, hoặc trùng dạng một địa chỉ ô.  §22 §23
   */
  function needsSheetQuote(name) {
    var s = String(name || '');
    if (!s) return false;
    if (/^\d/.test(s)) return true;
    if (CELL_LIKE.test(s)) return true;
    // Cho phép chữ cái (kể cả tiếng Việt có dấu), số, gạch dưới và dấu chấm
    return !/^[A-Za-z_\u00C0-\u024F\u1E00-\u1EFF][A-Za-z0-9_.\u00C0-\u024F\u1E00-\u1EFF]*$/.test(s);
  }

  function escapeSheetName(name) {
    var s = String(name || '');
    if (!needsSheetQuote(s)) return s;
    return "'" + s.replace(/'/g, "''") + "'";
  }

  /** Tiền tố sheet cho một tham chiếu; rỗng nếu cùng sheet đang thao tác. */
  function sheetPrefix(name, currentName) {
    var s = String(name || '');
    if (!s) return '';
    if (currentName != null && String(currentName) === s) return '';
    return escapeSheetName(s) + '!';
  }

  /** Chuỗi văn bản trong công thức: nháy kép, escape bằng cách nhân đôi. §199 */
  function quoteText(value) {
    return '"' + String(value == null ? '' : value).replace(/"/g, '""') + '"';
  }

  function looksNumeric(value) {
    var s = String(value == null ? '' : value).trim();
    if (s === '') return false;
    return /^-?\d+(?:[.,]\d+)?$/.test(s);
  }

  /**
   * Giá trị so sánh trong công thức. Nếu là tham chiếu ô (A2, $F$3) hoặc số
   * thì để trần; còn lại bọc nháy kép.
   */
  function literal(value, opts) {
    opts = opts || {};
    var s = String(value == null ? '' : value).trim();
    if (s === '') return opts.blankAsEmptyText === false ? '""' : '""';
    if (opts.forceText) return quoteText(s);
    if (opts.raw) return s;
    if (CELL_LIKE.test(s) || /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/.test(s)) return s;
    if (looksNumeric(s) && !opts.forceText) return s.replace(',', '.');
    if (/^(TRUE|FALSE)$/i.test(s)) return s.toUpperCase();
    return quoteText(s);
  }

  /* =============================================================
   * 2. Ref model
   *   { mode:'column'|'cell'|'custom', sheetName, colIndex, raw }
   * ========================================================== */

  function isRef(ref) {
    return !!(ref && (ref.mode === 'column' ? ref.colIndex != null && ref.colIndex >= 0 : String(ref.raw || '').trim() !== ''));
  }

  /** Địa chỉ ô của dòng dữ liệu hiện tại: A2 / $A2 / $A$2 */
  function cellRef(ref, ctx) {
    if (!ref) return '';
    if (ref.mode && ref.mode !== 'column') return String(ref.raw || '').trim();
    var letter = columnLetter(ref.colIndex);
    var row = ctx && ctx.dataRow ? ctx.dataRow : 2;
    var style = (ctx && ctx.referenceType) || 'relative';
    var prefix = sheetPrefix(ref.sheetName, ctx && ctx.currentSheet);
    if (style === 'absoluteColumn') return prefix + '$' + letter + row;
    if (style === 'absoluteCell') return prefix + '$' + letter + '$' + row;
    return prefix + letter + row;
  }

  /**
   * Vùng tham chiếu của một cột.
   * options.absolute  -> $B:$B
   * options.rangeMode -> 'full' | 'used' | 'custom'
   */
  function rangeRef(ref, ctx, options) {
    if (!ref) return '';
    options = options || {};
    if (ref.mode && ref.mode !== 'column') return String(ref.raw || '').trim();

    var letter = columnLetter(ref.colIndex);
    var abs = options.absolute ? '$' : '';
    var prefix = sheetPrefix(ref.sheetName, options.qualifyAlways ? null : (ctx && ctx.currentSheet));
    var mode = options.rangeMode || (ctx && ctx.rangeMode) || 'full';

    if (mode === 'used') {
      var start = (ctx && ctx.dataRow) || 2;
      var end = Math.max(start, (ref.usedRows || (ctx && ctx.usedRows) || 10000));
      // Khoá dòng để khi copy công thức xuống, vùng tham chiếu không trượt theo.
      // Công thức mảng dùng một lần (UNIQUE, SQL IN) không cần khoá.
      var rowLock = options.lockRows === false ? '' : '$';
      return prefix + abs + letter + rowLock + start + ':' + abs + letter + rowLock + end;
    }
    if (mode === 'custom' && ref.customRange) {
      return prefix + String(ref.customRange).trim();
    }
    return prefix + abs + letter + ':' + abs + letter;
  }

  /** Vùng nhiều cột: A:G  hoặc  A6:G */
  function multiRangeRef(sheetName, fromCol, toCol, ctx, options) {
    options = options || {};
    var prefix = sheetPrefix(sheetName, options.qualifyAlways ? null : (ctx && ctx.currentSheet));
    var a = columnLetter(fromCol);
    var b = columnLetter(toCol);
    if (options.startRow) return prefix + a + options.startRow + ':' + b;
    return prefix + a + ':' + b;
  }

  /* =============================================================
   * 3. Dựng lời gọi hàm theo separator
   * ========================================================== */

  function makeFn(sep) {
    return function (name, args) {
      return name + '(' + args.filter(function (a) { return a !== undefined && a !== null; }).join(sep) + ')';
    };
  }

  function resolveSeparator(setting, platform) {
    if (setting === 'comma') return SEPARATORS.comma;
    if (setting === 'semicolon') return SEPARATORS.semicolon;
    // AUTO: Google Sheets luôn dùng dấu phẩy; Excel theo locale máy người dùng.
    if (platform === 'sheets') return SEPARATORS.comma;
    return detectLocaleSeparator();
  }

  function detectLocaleSeparator() {
    try {
      var probe = (1234.5).toLocaleString();
      // Locale dùng dấu phẩy làm dấu thập phân (vi-VN, de-DE...) thường dùng ; trong Excel
      if (/1\.234,5/.test(probe) || /1234,5/.test(probe)) return SEPARATORS.semicolon;
    } catch (err) { /* bỏ qua */ }
    return SEPARATORS.comma;
  }

  /* =============================================================
   * 4. Pretty printer — tách argument cấp 1 xuống dòng  §87
   * ========================================================== */

  function prettyPrint(formula, sep) {
    if (!formula) return '';
    var out = '';
    var depth = 0;
    var inString = false;
    var indent = function (n) { return new Array(n + 1).join('  '); };
    var justOpened = false;

    for (var i = 0; i < formula.length; i++) {
      var ch = formula[i];

      if (inString) {
        out += ch;
        if (ch === '"') inString = false;
        continue;
      }
      if (ch === '"') { inString = true; out += ch; justOpened = false; continue; }

      if (ch === '(') {
        depth++;
        out += '(\n' + indent(depth);
        justOpened = true;
        continue;
      }
      if (ch === ')') {
        depth--;
        out = out.replace(/[ \t]+$/, '');
        out += '\n' + indent(depth) + ')';
        justOpened = false;
        continue;
      }
      if (ch === sep) {
        out = out.replace(/[ \t]+$/, '');
        out += sep + '\n' + indent(depth);
        justOpened = true;
        continue;
      }
      if (justOpened && ch === ' ') continue;
      justOpened = false;
      out += ch;
    }
    // Gộp lại dòng rỗng thừa
    return out.split('\n').filter(function (line, idx, arr) {
      return line.trim() !== '' || idx === 0 || idx === arr.length - 1;
    }).join('\n');
  }

  /* =============================================================
   * 5. Ngữ cảnh sinh công thức
   * ========================================================== */

  function buildContext(config) {
    var platform = config.platform === 'sheets' ? 'sheets' : 'excel';
    var sep = resolveSeparator(config.separator || 'auto', platform);
    return {
      platform: platform,
      sep: sep,
      fn: makeFn(sep),
      dataRow: Math.max(1, parseInt(config.dataRow, 10) || 2),
      currentSheet: config.currentSheet || null,
      referenceType: config.referenceType || 'relative',
      rangeMode: config.rangeMode || 'full',
      excelVersion: config.excelVersion === 'legacy' ? 'legacy' : 'modern',
      usedRows: config.usedRows || 10000,
      warnings: [],
      explanation: []
    };
  }

  function explain(ctx, code, text) {
    ctx.explanation.push({ code: code, text: text });
  }

  function warn(ctx, text) {
    if (ctx.warnings.indexOf(text) === -1) ctx.warnings.push(text);
  }

  function warnFullColumn(ctx, used) {
    if (used && ctx.rangeMode === 'full') {
      warn(ctx, 'Tham chiếu toàn cột có thể làm file Excel chậm nếu dữ liệu rất lớn. Cân nhắc chuyển sang Used Range.');
    }
  }

  function requireRef(ctx, ref, label) {
    if (!isRef(ref)) {
      throw new ConfigError(label);
    }
    return ref;
  }

  function ConfigError(message) {
    this.name = 'ConfigError';
    this.message = message;
  }
  ConfigError.prototype = Object.create(Error.prototype);

  /* =============================================================
   * 6. Các bộ sinh công thức theo nhu cầu nghiệp vụ
   * ========================================================== */

  var GENERATORS = {};

  /* ---- 6.1 Tra cứu dữ liệu (§31–36) ---- */
  GENERATORS.lookup = function (cfg, ctx) {
    var value = requireRef(ctx, cfg.lookupValue, 'Vui lòng chọn giá trị cần tìm.');
    var search = requireRef(ctx, cfg.lookupRange, 'Vui lòng chọn cột cần tìm.');
    var ret = requireRef(ctx, cfg.returnRange, 'Vui lòng chọn cột trả về.');

    var method = cfg.method || 'auto';
    if (method === 'auto') method = ctx.excelVersion === 'legacy' ? 'index-match' : 'xlookup';

    var extra = (cfg.conditions || []).filter(function (c) { return isRef(c.field) && String(c.value || '').trim() !== ''; });
    if (extra.length) return multiConditionLookup(cfg, ctx, value, search, ret, extra);

    var lookupCell = cellRef(value, ctx);
    var searchRange = rangeRef(search, ctx, { absolute: false });
    var returnRange = rangeRef(ret, ctx, { absolute: false });
    var fallback = quoteText(cfg.notFound || '');
    var formula;

    warnFullColumn(ctx, true);

    if (method === 'vlookup') {
      if (ret.sheetName !== search.sheetName) {
        warn(ctx, 'VLOOKUP yêu cầu cột tìm và cột trả về nằm trong cùng một Sheet. Nên dùng XLOOKUP.');
      } else if (ret.colIndex < search.colIndex) {
        warn(ctx, 'VLOOKUP không phù hợp vì cột trả về nằm bên trái cột tìm kiếm. Nên sử dụng XLOOKUP.');
      }
      var tableStart = Math.min(search.colIndex, ret.colIndex);
      var tableEnd = Math.max(search.colIndex, ret.colIndex);
      var table = multiRangeRef(search.sheetName, tableStart, tableEnd, ctx, {});
      var colOffset = (ret.colIndex - tableStart) + 1;
      var core = ctx.fn('VLOOKUP', [lookupCell, table, String(colOffset), 'FALSE']);
      formula = cfg.notFound !== undefined && cfg.notFound !== null
        ? ctx.fn('IFERROR', [core, fallback])
        : core;
      explain(ctx, lookupCell, 'Giá trị cần tìm ở dòng hiện tại.');
      explain(ctx, table, 'Vùng bảng tra cứu, cột tìm phải là cột đầu tiên.');
      explain(ctx, String(colOffset), 'Thứ tự cột trả về tính từ cột đầu của vùng.');
      explain(ctx, 'FALSE', 'Tìm chính xác tuyệt đối.');
    } else if (method === 'index-match') {
      var core2 = ctx.fn('INDEX', [
        returnRange,
        ctx.fn('MATCH', [lookupCell, searchRange, '0'])
      ]);
      formula = ctx.fn('IFERROR', [core2, fallback]);
      explain(ctx, lookupCell, 'Giá trị cần tìm ở dòng hiện tại.');
      explain(ctx, searchRange, 'MATCH xác định vị trí dòng khớp trong cột tìm.');
      explain(ctx, returnRange, 'INDEX lấy giá trị tại vị trí đó trong cột trả về.');
      explain(ctx, fallback, 'Giá trị trả về khi không tìm thấy.');
    } else {
      formula = ctx.fn('XLOOKUP', [lookupCell, searchRange, returnRange, fallback]);
      if (ctx.excelVersion === 'legacy') {
        warn(ctx, 'XLOOKUP yêu cầu phiên bản Excel hỗ trợ hàm này (Microsoft 365 / Excel 2021).');
      }
      explain(ctx, lookupCell, 'Giá trị cần tìm ở dòng hiện tại.');
      explain(ctx, searchRange, 'Cột dùng để tìm ' + fieldLabel(value) + '.');
      explain(ctx, returnRange, 'Cột trả về ' + fieldLabel(ret) + '.');
      explain(ctx, fallback, cfg.notFound ? 'Giá trị trả về khi không tìm thấy.' : 'Nếu không tìm thấy thì trả về rỗng.');
    }

    if (!cfg.notFound) {
      warn(ctx, 'Giá trị khi không tìm thấy đang để rỗng. Cân nhắc điền nhãn để dễ rà soát.');
    }
    return formula;
  };

  function multiConditionLookup(cfg, ctx, value, search, ret, extra) {
    // Nhiều điều kiện -> FILTER (§35 §36)
    var returnRange = rangeRef(ret, ctx, {});
    var parts = [];
    parts.push('(' + rangeRef(search, ctx, {}) + '=' + cellRef(value, ctx) + ')');
    extra.forEach(function (c) {
      parts.push('(' + rangeRef(c.field, ctx, {}) + '=' + literal(c.value) + ')');
    });
    var condition = parts.join('*');
    explain(ctx, condition, 'Các điều kiện nhân với nhau, tương đương phép AND.');
    explain(ctx, returnRange, 'Cột trả về ' + fieldLabel(ret) + '.');
    return ctx.fn('FILTER', [returnRange, condition, quoteText(cfg.notFound || 'Không tìm thấy')]);
  }

  /* ---- 6.2 Kiểm tra trùng (§37–39) ---- */
  GENERATORS.duplicate = function (cfg, ctx) {
    var fields = (cfg.fields || []).filter(isRef);
    if (!fields.length) throw new ConfigError('Vui lòng chọn ít nhất một cột để kiểm tra trùng.');

    var args = [];
    fields.forEach(function (f) {
      args.push(rangeRef(f, ctx, { absolute: true }));
      args.push(cellRef(f, ctx));
    });

    var counter = fields.length > 1
      ? ctx.fn('COUNTIFS', args)
      : ctx.fn('COUNTIF', args);

    explain(ctx, args[0], 'Vùng dữ liệu dùng để đếm số lần xuất hiện.');
    explain(ctx, args[1], 'Giá trị của dòng hiện tại đem đi đếm.');
    warnFullColumn(ctx, true);

    if (cfg.resultType === 'count') {
      explain(ctx, counter, 'Kết quả là số lần giá trị này xuất hiện.');
      return counter;
    }
    var dup = quoteText(cfg.duplicateLabel == null ? 'TRÙNG' : cfg.duplicateLabel);
    var uniq = quoteText(cfg.uniqueLabel == null ? '' : cfg.uniqueLabel);
    explain(ctx, '>1', 'Xuất hiện nhiều hơn một lần nghĩa là bị trùng.');
    return ctx.fn('IF', [counter + '>1', dup, uniq]);
  };

  /* ---- 6.3 Kiểm tra tồn tại (§40–41) ---- */
  GENERATORS.existence = function (cfg, ctx) {
    var value = requireRef(ctx, cfg.value, 'Vui lòng chọn giá trị cần kiểm tra.');
    var reference = requireRef(ctx, cfg.reference, 'Vui lòng chọn cột danh mục để đối chiếu.');

    var cell = cellRef(value, ctx);
    var range = rangeRef(reference, ctx, {});
    var found = quoteText(cfg.foundLabel == null ? 'CÓ' : cfg.foundLabel);
    var missing = quoteText(cfg.missingLabel == null ? 'KHÔNG' : cfg.missingLabel);
    var method = cfg.method || 'countif';

    explain(ctx, cell, 'Giá trị cần kiểm tra ở dòng hiện tại.');
    explain(ctx, range, 'Danh mục dùng để đối chiếu.');
    warnFullColumn(ctx, true);

    if (method === 'xmatch') {
      if (ctx.excelVersion === 'legacy') warn(ctx, 'XMATCH yêu cầu Excel Microsoft 365 / 2021.');
      return ctx.fn('IF', [
        ctx.fn('ISNUMBER', [ctx.fn('XMATCH', [cell, range])]),
        found, missing
      ]);
    }
    if (method === 'xlookup') {
      if (ctx.excelVersion === 'legacy') warn(ctx, 'XLOOKUP yêu cầu Excel Microsoft 365 / 2021.');
      return ctx.fn('XLOOKUP', [cell, range, found, missing]);
    }
    explain(ctx, '>0', 'Đếm được ít nhất một lần nghĩa là giá trị có tồn tại.');
    return ctx.fn('IF', [ctx.fn('COUNTIF', [range, cell]) + '>0', found, missing]);
  };

  /* ---- 6.4 So sánh hai giá trị (§42–45) ---- */
  GENERATORS.compare = function (cfg, ctx) {
    var a = requireRef(ctx, cfg.valueA, 'Vui lòng chọn giá trị thứ nhất.');
    var b = requireRef(ctx, cfg.valueB, 'Vui lòng chọn giá trị thứ hai.');

    var left = cellRef(a, ctx);
    var right = cellRef(b, ctx);
    if (cfg.trim) { left = ctx.fn('TRIM', [left]); right = ctx.fn('TRIM', [right]); }
    if (cfg.ignoreCase) { left = ctx.fn('LOWER', [left]); right = ctx.fn('LOWER', [right]); }

    explain(ctx, left, 'Giá trị bên trái' + (cfg.trim ? ', đã cắt khoảng trắng' : '') + (cfg.ignoreCase ? ', đã đưa về chữ thường' : '') + '.');
    explain(ctx, right, 'Giá trị bên phải để đối chiếu.');

    return ctx.fn('IF', [
      left + '=' + right,
      quoteText(cfg.sameLabel == null ? 'GIỐNG' : cfg.sameLabel),
      quoteText(cfg.diffLabel == null ? 'KHÁC' : cfg.diffLabel)
    ]);
  };

  /* ---- 6.5 So sánh hai danh sách (§46–47) ---- */
  GENERATORS.listCompare = function (cfg, ctx) {
    var value = requireRef(ctx, cfg.value, 'Vui lòng chọn cột của danh sách A.');
    var other = requireRef(ctx, cfg.reference, 'Vui lòng chọn cột của danh sách B.');

    var cell = cellRef(value, ctx);
    var range = rangeRef(other, ctx, {});
    var count = ctx.fn('COUNTIF', [range, cell]);

    explain(ctx, cell, 'Giá trị đang xét ở danh sách A.');
    explain(ctx, range, 'Toàn bộ danh sách B dùng để đối chiếu.');
    warnFullColumn(ctx, true);

    var mode = cfg.resultMode || 'both';
    if (mode === 'onlyA') {
      explain(ctx, '=0', 'Không xuất hiện ở danh sách B nghĩa là chỉ có ở A.');
      return ctx.fn('IF', [count + '=0', quoteText(cfg.matchLabel || 'CHỈ CÓ Ở A'), quoteText('')]);
    }
    return ctx.fn('IF', [
      count + '>0',
      quoteText(cfg.matchLabel == null ? 'CÓ' : cfg.matchLabel),
      quoteText(cfg.missLabel == null ? 'THIẾU' : cfg.missLabel)
    ]);
  };

  /* ---- 6.6 Lọc dữ liệu (§48–51) ---- */
  var OPERATORS = ['=', '<>', '>', '<', '>=', '<=', 'contains', 'startsWith', 'endsWith', 'blank', 'notBlank'];

  function conditionExpr(cond, ctx, asArray) {
    var ref = cond.field;
    var target = asArray
      ? rangeRef(ref, ctx, {})
      : cellRef(ref, ctx);
    var raw = cond.value == null ? '' : String(cond.value);
    var val = literal(raw, { forceText: cond.forceText });

    switch (cond.operator) {
      case 'contains':
        return ctx.fn('ISNUMBER', [ctx.fn('SEARCH', [val, target])]);
      case 'startsWith':
        return ctx.fn('LEFT', [target, ctx.fn('LEN', [val])]) + '=' + val;
      case 'endsWith':
        return ctx.fn('RIGHT', [target, ctx.fn('LEN', [val])]) + '=' + val;
      case 'blank':
        return target + '=""';
      case 'notBlank':
        return target + '<>""';
      default:
        return target + (cond.operator || '=') + val;
    }
  }

  GENERATORS.filter = function (cfg, ctx) {
    var sheetName = cfg.sourceSheet;
    if (!sheetName) throw new ConfigError('Vui lòng chọn Sheet nguồn.');
    var conds = (cfg.conditions || []).filter(function (c) { return isRef(c.field); });
    if (!conds.length) throw new ConfigError('Vui lòng thêm ít nhất một điều kiện lọc.');

    var returnRange = cfg.returnRange
      ? sheetPrefix(sheetName, null) + String(cfg.returnRange).trim()
      : multiRangeRef(sheetName, cfg.fromCol || 0, cfg.toCol || 0, ctx, {});

    if (ctx.platform === 'sheets' && cfg.sheetsMethod === 'query') {
      return buildQuery(cfg, ctx, sheetName, conds);
    }

    var joiner = cfg.logic === 'OR' ? '+' : '*';
    var parts = conds.map(function (c) { return '(' + conditionExpr(c, ctx, true) + ')'; });
    var condition = parts.join(joiner);

    explain(ctx, returnRange, 'Vùng dữ liệu sẽ được trả về.');
    explain(ctx, condition, joiner === '+'
      ? 'Các điều kiện cộng với nhau, tương đương phép OR.'
      : 'Các điều kiện nhân với nhau, tương đương phép AND.');
    warnFullColumn(ctx, true);
    if (ctx.platform === 'excel' && ctx.excelVersion === 'legacy') {
      warn(ctx, 'FILTER yêu cầu Excel Microsoft 365 / 2021. Excel đời cũ không có hàm này.');
    }

    return ctx.fn('FILTER', [returnRange, condition, quoteText(cfg.notFound || '')]);
  };

  function queryColumnLetter(cond) {
    return columnLetter(cond.field.colIndex);
  }

  function buildQuery(cfg, ctx, sheetName, conds) {
    var range = cfg.returnRange
      ? escapeSheetName(sheetName) + '!' + String(cfg.returnRange).trim()
      : escapeSheetName(sheetName) + '!' + multiRangeRef('', cfg.fromCol || 0, cfg.toCol || 0, ctx, {});
    var joiner = cfg.logic === 'OR' ? ' or ' : ' and ';

    var where = conds.map(function (c) {
      var col = queryColumnLetter(c);
      var v = String(c.value == null ? '' : c.value);
      var isCell = CELL_LIKE.test(v.trim());
      var quoted = isCell ? '\'"&' + v.trim() + '&"\'' : "'" + v.replace(/'/g, "\\'") + "'";
      if (c.operator === 'contains') return col + ' contains ' + quoted;
      if (c.operator === 'startsWith') return col + ' starts with ' + quoted;
      if (c.operator === 'endsWith') return col + ' ends with ' + quoted;
      if (c.operator === 'blank') return col + ' is null';
      if (c.operator === 'notBlank') return col + ' is not null';
      return col + ' ' + (c.operator === '<>' ? '!=' : (c.operator || '=')) + ' ' + quoted;
    }).join(joiner);

    var select = cfg.querySelect ? String(cfg.querySelect).trim() : '*';
    var query = 'select ' + select + ' where ' + where;
    if (cfg.orderBy) query += ' order by ' + cfg.orderBy + (cfg.orderDesc ? ' desc' : '');
    if (cfg.limit) query += ' limit ' + parseInt(cfg.limit, 10);

    explain(ctx, range, 'Vùng dữ liệu nguồn của QUERY.');
    explain(ctx, '"' + query + '"', 'Câu truy vấn dạng SQL rút gọn của Google Sheets.');
    explain(ctx, '0', 'Số dòng tiêu đề trong vùng nguồn; đặt 0 nếu vùng đã bỏ header.');

    var literalQuery = '"' + query.replace(/"/g, '""') + '"';
    // Ghép chuỗi khi có tham chiếu ô
    literalQuery = literalQuery.replace(/""&/g, '"&').replace(/&""/g, '&"');
    return ctx.fn('QUERY', [range, literalQuery, String(cfg.headerRows == null ? 0 : cfg.headerRows)]);
  }

  /* ---- 6.7 Kiểm tra điều kiện (§52–55) ---- */
  function groupExpr(group, ctx) {
    if (!group) return '';
    if (group.kind === 'group') {
      var children = (group.children || [])
        .map(function (c) { return groupExpr(c, ctx); })
        .filter(function (s) { return s; });
      if (!children.length) return '';
      if (children.length === 1) return children[0];
      return ctx.fn(group.operator === 'OR' ? 'OR' : 'AND', children);
    }
    if (!isRef(group.field)) return '';
    return conditionExpr(group, ctx, false);
  }

  GENERATORS.logic = function (cfg, ctx) {
    var branches = (cfg.branches || []).filter(function (b) {
      return b && b.condition && groupExpr(b.condition, ctx);
    });

    if (cfg.style === 'ifs' || branches.length > 1) {
      if (!branches.length) throw new ConfigError('Vui lòng thêm ít nhất một điều kiện.');
      var args = [];
      branches.forEach(function (b) {
        args.push(groupExpr(b.condition, ctx));
        args.push(quoteText(b.result == null ? '' : b.result));
      });
      args.push('TRUE');
      args.push(quoteText(cfg.elseResult == null ? '' : cfg.elseResult));
      explain(ctx, 'TRUE', 'Nhánh cuối luôn đúng, đóng vai trò ELSE.');
      if (ctx.platform === 'excel' && ctx.excelVersion === 'legacy') {
        warn(ctx, 'IFS yêu cầu Excel 2019 trở lên. Excel đời cũ cần lồng nhiều hàm IF.');
      }
      return ctx.fn('IFS', args);
    }

    if (!branches.length) throw new ConfigError('Vui lòng thêm ít nhất một điều kiện.');
    var expr = groupExpr(branches[0].condition, ctx);
    explain(ctx, expr, 'Điều kiện cần kiểm tra.');
    return ctx.fn('IF', [
      expr,
      quoteText(branches[0].result == null ? '' : branches[0].result),
      quoteText(cfg.elseResult == null ? '' : cfg.elseResult)
    ]);
  };

  /* ---- 6.8 Kiểm tra dữ liệu rỗng (§76) ---- */
  GENERATORS.blank = function (cfg, ctx) {
    var fields = (cfg.fields || []).filter(isRef);
    if (!fields.length) throw new ConfigError('Vui lòng chọn ít nhất một cột bắt buộc.');
    var checks = fields.map(function (f) { return cellRef(f, ctx) + '=""'; });
    var expr = checks.length === 1 ? checks[0] : ctx.fn('OR', checks);
    explain(ctx, expr, 'Chỉ cần một cột bắt buộc bị rỗng là báo thiếu.');
    return ctx.fn('IF', [
      expr,
      quoteText(cfg.missingLabel == null ? 'THIẾU DỮ LIỆU' : cfg.missingLabel),
      quoteText(cfg.okLabel == null ? '' : cfg.okLabel)
    ]);
  };

  /* ---- 6.9 Làm sạch Text (§56–67) ---- */
  var TEXT_PRESETS = {
    trim: function (cell, cfg, ctx) { return ctx.fn('TRIM', [cell]); },
    clean: function (cell, cfg, ctx) { return ctx.fn('CLEAN', [cell]); },
    trimClean: function (cell, cfg, ctx) { return ctx.fn('TRIM', [ctx.fn('CLEAN', [cell])]); },
    upper: function (cell, cfg, ctx) { return ctx.fn('UPPER', [ctx.fn('TRIM', [cell])]); },
    lower: function (cell, cfg, ctx) { return ctx.fn('LOWER', [ctx.fn('TRIM', [cell])]); },
    proper: function (cell, cfg, ctx) { return ctx.fn('PROPER', [ctx.fn('TRIM', [cell])]); },
    removeSpace: function (cell, cfg, ctx) { return ctx.fn('SUBSTITUTE', [cell, '" "', '""']); },
    removeLineBreak: function (cell, cfg, ctx) {
      return ctx.fn('SUBSTITUTE', [cell, ctx.fn('CHAR', ['10']), '" "']);
    },
    substitute: function (cell, cfg, ctx) {
      return ctx.fn('SUBSTITUTE', [cell, quoteText(cfg.find || ''), quoteText(cfg.replace || '')]);
    },
    left: function (cell, cfg, ctx) { return ctx.fn('LEFT', [cell, String(parseInt(cfg.count, 10) || 1)]); },
    right: function (cell, cfg, ctx) { return ctx.fn('RIGHT', [cell, String(parseInt(cfg.count, 10) || 1)]); },
    mid: function (cell, cfg, ctx) {
      return ctx.fn('MID', [cell, String(parseInt(cfg.start, 10) || 1), String(parseInt(cfg.count, 10) || 1)]);
    },
    len: function (cell, cfg, ctx) { return ctx.fn('LEN', [cell]); }
  };

  var TEXT_NOTES = {
    trim: 'TRIM cắt khoảng trắng đầu, cuối và gộp khoảng trắng kép ở giữa.',
    clean: 'CLEAN loại bỏ các ký tự điều khiển không in được.',
    trimClean: 'CLEAN bỏ ký tự lạ trước, TRIM cắt khoảng trắng sau.',
    upper: 'UPPER đưa toàn bộ về chữ hoa.',
    lower: 'LOWER đưa toàn bộ về chữ thường.',
    proper: 'PROPER viết hoa chữ cái đầu mỗi từ.',
    removeSpace: 'Thay mọi khoảng trắng bằng chuỗi rỗng.',
    removeLineBreak: 'CHAR(10) là ký tự xuống dòng, được thay bằng dấu cách.',
    substitute: 'SUBSTITUTE thay thế mọi lần xuất hiện của chuỗi cần tìm.',
    left: 'LEFT lấy số ký tự tính từ bên trái.',
    right: 'RIGHT lấy số ký tự tính từ bên phải.',
    mid: 'MID lấy ký tự bắt đầu từ vị trí chỉ định.',
    len: 'LEN đếm số ký tự của giá trị.'
  };

  GENERATORS.text = function (cfg, ctx) {
    var field = requireRef(ctx, cfg.field, 'Vui lòng chọn cột cần xử lý.');
    var preset = cfg.preset || 'trim';
    var builder = TEXT_PRESETS[preset] || TEXT_PRESETS.trim;
    var cell = cellRef(field, ctx);
    explain(ctx, cell, 'Ô dữ liệu của dòng hiện tại.');
    if (TEXT_NOTES[preset]) explain(ctx, preset.toUpperCase(), TEXT_NOTES[preset]);
    return builder(cell, cfg, ctx);
  };

  /* ---- 6.10 Ghép dữ liệu (§68–69) ---- */
  var SEP_PRESETS = { space: ' ', dash: '-', slash: '/', underscore: '_', comma: ', ', none: '' };

  GENERATORS.concat = function (cfg, ctx) {
    var fields = (cfg.fields || []).filter(isRef);
    if (fields.length < 2) throw new ConfigError('Vui lòng chọn ít nhất hai cột để ghép.');
    var sepValue = cfg.separatorText != null && cfg.separatorText !== ''
      ? cfg.separatorText
      : (SEP_PRESETS[cfg.joinWith] != null ? SEP_PRESETS[cfg.joinWith] : ' ');
    var cells = fields.map(function (f) { return cellRef(f, ctx); });

    if (cfg.method === 'ampersand') {
      explain(ctx, '&', 'Toán tử & nối các giá trị lại với nhau.');
      return cells.join(sepValue === '' ? '&' : '&' + quoteText(sepValue) + '&');
    }
    explain(ctx, quoteText(sepValue), 'Ký tự ngăn cách giữa các giá trị.');
    explain(ctx, cfg.ignoreBlank === false ? 'FALSE' : 'TRUE',
      cfg.ignoreBlank === false ? 'Vẫn giữ ô rỗng khi ghép.' : 'Bỏ qua các ô rỗng khi ghép.');
    return ctx.fn('TEXTJOIN', [quoteText(sepValue), cfg.ignoreBlank === false ? 'FALSE' : 'TRUE'].concat(cells));
  };

  /* ---- 6.11 Xử lý ngày (§70–75) ---- */
  GENERATORS.date = function (cfg, ctx) {
    var field = requireRef(ctx, cfg.field, 'Vui lòng chọn cột ngày.');
    var cell = cellRef(field, ctx);
    var preset = cfg.preset || 'format';
    explain(ctx, cell, 'Ô chứa giá trị ngày của dòng hiện tại.');

    switch (preset) {
      case 'format':
        var fmt = cfg.format || 'dd/mm/yyyy';
        explain(ctx, quoteText(fmt), 'Định dạng hiển thị ngày mong muốn.');
        return ctx.fn('TEXT', [cell, quoteText(fmt)]);
      case 'age':
        explain(ctx, 'TODAY()', 'Ngày hiện tại, dùng làm mốc tính tuổi.');
        explain(ctx, '"Y"', 'Đơn vị tính là số năm tròn.');
        return ctx.fn('DATEDIF', [cell, ctx.fn('TODAY', []), '"Y"']);
      case 'daysBetween':
      case 'monthsBetween':
      case 'yearsBetween':
        var other = requireRef(ctx, cfg.field2, 'Vui lòng chọn cột ngày thứ hai.');
        var unit = preset === 'daysBetween' ? '"D"' : (preset === 'monthsBetween' ? '"M"' : '"Y"');
        explain(ctx, unit, 'Đơn vị chênh lệch giữa hai mốc thời gian.');
        return ctx.fn('DATEDIF', [cell, cellRef(other, ctx), unit]);
      case 'year': return ctx.fn('YEAR', [cell]);
      case 'month': return ctx.fn('MONTH', [cell]);
      case 'day': return ctx.fn('DAY', [cell]);
      case 'academicYear':
        var start = Math.max(1, Math.min(12, parseInt(cfg.startMonth, 10) || 8));
        explain(ctx, String(start), 'Tháng bắt đầu năm học.');
        return ctx.fn('IF', [
          ctx.fn('MONTH', [cell]) + '>=' + start,
          ctx.fn('YEAR', [cell]) + '&"-"&' + ctx.fn('YEAR', [cell]) + '+1',
          ctx.fn('YEAR', [cell]) + '-1&"-"&' + ctx.fn('YEAR', [cell])
        ]);
      default:
        return ctx.fn('TEXT', [cell, quoteText('dd/mm/yyyy')]);
    }
  };

  /* ---- 6.12 Đếm / Tổng (§77–79) ---- */
  GENERATORS.aggregate = function (cfg, ctx) {
    var conds = (cfg.conditions || []).filter(function (c) { return isRef(c.field); });
    var kind = cfg.kind || 'sum';

    if (kind === 'count') {
      if (!conds.length) throw new ConfigError('Vui lòng thêm ít nhất một điều kiện đếm.');
      var cargs = [];
      conds.forEach(function (c) {
        cargs.push(rangeRef(c.field, ctx, { absolute: true }));
        cargs.push(criteriaValue(c, ctx));
      });
      warnFullColumn(ctx, true);
      explain(ctx, cargs[0], 'Vùng chứa dữ liệu cần xét điều kiện.');
      explain(ctx, cargs[1], 'Điều kiện áp dụng cho vùng đó.');
      return ctx.fn(conds.length > 1 ? 'COUNTIFS' : 'COUNTIF', cargs);
    }

    var sumField = requireRef(ctx, cfg.sumField, 'Vui lòng chọn cột cần tính tổng.');
    var sumRange = rangeRef(sumField, ctx, { absolute: true });
    if (!conds.length) {
      explain(ctx, sumRange, 'Vùng dữ liệu được cộng lại.');
      return ctx.fn('SUM', [sumRange]);
    }
    var args = [sumRange];
    conds.forEach(function (c) {
      args.push(rangeRef(c.field, ctx, { absolute: true }));
      args.push(criteriaValue(c, ctx));
    });
    explain(ctx, sumRange, 'Vùng dữ liệu được cộng lại.');
    explain(ctx, args[1], 'Vùng chứa điều kiện.');
    warnFullColumn(ctx, true);
    return ctx.fn(conds.length > 1 ? 'SUMIFS' : 'SUMIF', args);
  };

  function criteriaValue(cond, ctx) {
    var raw = cond.value == null ? '' : String(cond.value).trim();
    var op = cond.operator || '=';
    if (op === 'contains') return '"*"&' + literal(raw) + '&"*"';
    if (op === 'startsWith') return literal(raw) + '&"*"';
    if (op === 'endsWith') return '"*"&' + literal(raw);
    if (op === '=') return literal(raw);
    if (CELL_LIKE.test(raw)) return quoteText(op) + '&' + raw;
    return quoteText(op + raw);
  }

  /* ---- 6.13 Danh sách duy nhất (§80) ---- */
  GENERATORS.unique = function (cfg, ctx) {
    var field = requireRef(ctx, cfg.field, 'Vui lòng chọn cột cần lấy danh sách.');
    var range = rangeRef(field, ctx, { rangeMode: 'used', lockRows: false });
    explain(ctx, range, 'Vùng dữ liệu nguồn, không bao gồm dòng tiêu đề.');
    if (ctx.platform === 'excel' && ctx.excelVersion === 'legacy') {
      warn(ctx, 'UNIQUE yêu cầu Excel Microsoft 365 / 2021.');
    }
    var core = ctx.fn('UNIQUE', [range]);
    if (cfg.sort) {
      explain(ctx, 'SORT', 'Sắp xếp kết quả tăng dần.');
      return ctx.fn('SORT', [core]);
    }
    return core;
  };

  /* ---- 6.14 SQL Tools (§81–84) ---- */
  GENERATORS.sqlQuote = function (cfg, ctx) {
    var field = requireRef(ctx, cfg.field, 'Vui lòng chọn cột cần tạo giá trị SQL.');
    var cell = cellRef(field, ctx);
    if (cfg.dataType === 'number') {
      explain(ctx, cell, 'Kiểu số nên giữ nguyên, không đóng nháy.');
      return cell;
    }
    explain(ctx, '"\'"', 'Nháy đơn mở và đóng của chuỗi trong SQL.');
    explain(ctx, cell, 'Giá trị lấy từ ô dữ liệu.');
    return '"\'"&' + cell + '&"\'"';
  };

  GENERATORS.sqlIn = function (cfg, ctx) {
    var field = requireRef(ctx, cfg.field, 'Vui lòng chọn cột cần gom thành danh sách IN.');
    var range = rangeRef(field, ctx, { rangeMode: 'used', lockRows: false });
    explain(ctx, range, 'Vùng giá trị được gom lại thành một chuỗi.');
    explain(ctx, '","', 'Dấu phân cách giữa các phần tử trong mệnh đề IN.');
    explain(ctx, 'TRUE', 'Bỏ qua các ô rỗng.');

    if (cfg.dataType === 'number') {
      return ctx.fn('TEXTJOIN', ['","', 'TRUE', range]);
    }
    return ctx.fn('TEXTJOIN', ['","', 'TRUE', '"\'"&' + range + '&"\'"']);
  };

  /* =============================================================
   * 7. API công khai
   * ========================================================== */

  function fieldLabel(ref) {
    if (!ref) return 'giá trị';
    if (ref.mode && ref.mode !== 'column') return String(ref.raw || '');
    return (ref.sheetName ? ref.sheetName + '.' : '') + (ref.column || columnLetter(ref.colIndex));
  }

  /**
   * Sinh công thức từ FormulaConfig.  §160
   * @returns {{formula:string, formattedFormula:string, explanation:Array, warnings:Array, error:string|null}}
   */
  function generate(config) {
    config = config || {};
    var ctx = buildContext(config);
    var gen = GENERATORS[config.type];

    if (!gen) {
      return {
        formula: '', formattedFormula: '', explanation: [], warnings: [],
        error: 'Chưa hỗ trợ loại công thức: ' + config.type
      };
    }

    var body;
    try {
      body = gen(config, ctx);
    } catch (err) {
      return {
        formula: '', formattedFormula: '', explanation: [], warnings: ctx.warnings,
        error: err && err.message ? err.message : 'Cấu hình chưa hợp lệ.'
      };
    }

    if (!body) {
      return {
        formula: '', formattedFormula: '', explanation: [], warnings: ctx.warnings,
        error: 'Cấu hình chưa đủ để sinh công thức.'
      };
    }

    var formula = '=' + body;
    var check = validateParens(formula);
    if (!check.ok) {
      return {
        formula: '', formattedFormula: '', explanation: ctx.explanation, warnings: ctx.warnings,
        error: check.message
      };
    }

    return {
      formula: formula,
      formattedFormula: '=' + prettyPrint(body, ctx.sep),
      explanation: ctx.explanation,
      warnings: ctx.warnings,
      separator: ctx.sep,
      error: null
    };
  }

  /** Kiểm tra ngoặc và nháy kép cân đối. §128 §156 */
  function validateParens(formula) {
    var depth = 0, inString = false;
    for (var i = 0; i < formula.length; i++) {
      var ch = formula[i];
      if (ch === '"') { inString = !inString; continue; }
      if (inString) continue;
      if (ch === '(') depth++;
      if (ch === ')') {
        depth--;
        if (depth < 0) return { ok: false, message: 'Công thức bị dư dấu đóng ngoặc.' };
      }
    }
    if (inString) return { ok: false, message: 'Công thức có chuỗi chưa đóng nháy kép.' };
    if (depth > 0) return { ok: false, message: 'Công thức bị thiếu ' + depth + ' dấu đóng ngoặc.' };
    return { ok: true };
  }

  global.FormulaEngine = {
    generate: generate,
    columnLetter: columnLetter,
    escapeSheetName: escapeSheetName,
    needsSheetQuote: needsSheetQuote,
    sheetPrefix: sheetPrefix,
    quoteText: quoteText,
    literal: literal,
    cellRef: cellRef,
    rangeRef: rangeRef,
    prettyPrint: prettyPrint,
    resolveSeparator: resolveSeparator,
    validateParens: validateParens,
    fieldLabel: fieldLabel,
    OPERATORS: OPERATORS,
    SEPARATORS: SEPARATORS
  };
})(typeof window !== 'undefined' ? window : this);
