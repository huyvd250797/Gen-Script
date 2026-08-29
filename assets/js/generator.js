/*!
 * generator.js — Sinh script INSERT / UPDATE / SELECT từ dữ liệu dạng bảng.
 * Port từ VBA Module1.GetInsertSQL và Module2.UpdateSQL trong gen-script.xlsm.
 * Không phụ thuộc thư viện ngoài.
 */
(function (global) {
  'use strict';

  /* ---------------------------------------------------------------
   * Cấu hình theo hệ quản trị CSDL
   * ------------------------------------------------------------- */
  var DIALECTS = {
    mssql: {
      label: 'SQL Server (T-SQL)',
      open: '[', close: ']',
      escapeIdent: function (s) { return s.replace(/]/g, ']]'); },
      unicodePrefix: true,
      supportsIdentityInsert: true,
      supportsGo: true,
      maxRowsPerInsert: 1000,
      defaultSemicolon: false
    },
    mysql: {
      label: 'MySQL / MariaDB',
      open: '`', close: '`',
      escapeIdent: function (s) { return s.replace(/`/g, '``'); },
      unicodePrefix: false,
      supportsIdentityInsert: false,
      supportsGo: false,
      maxRowsPerInsert: 1000,
      defaultSemicolon: true
    },
    postgres: {
      label: 'PostgreSQL',
      open: '"', close: '"',
      escapeIdent: function (s) { return s.replace(/"/g, '""'); },
      unicodePrefix: false,
      supportsIdentityInsert: false,
      supportsGo: false,
      maxRowsPerInsert: 1000,
      defaultSemicolon: true
    },
    oracle: {
      label: 'Oracle',
      open: '"', close: '"',
      escapeIdent: function (s) { return s.replace(/"/g, '""'); },
      unicodePrefix: true,
      supportsIdentityInsert: false,
      supportsGo: false,
      maxRowsPerInsert: 1,
      defaultSemicolon: true
    }
  };

  var DEFAULT_OPTIONS = {
    dialect: 'mssql',
    identityInsert: true,      // SET IDENTITY_INSERT ... ON/OFF (chỉ SQL Server)
    unicodePrefix: true,       // N'...' cho chuỗi
    nullKeyword: true,         // ô ghi chữ NULL -> giá trị NULL
    trimValues: true,          // cắt khoảng trắng đầu/cuối
    leadingZeroAsText: true,   // 001129 giữ nguyên dạng chuỗi
    skipEmptyRows: true,       // bỏ dòng trống hoàn toàn
    multiRowInsert: false,     // gộp nhiều dòng vào một lệnh INSERT
    semicolon: false,          // kết thúc câu lệnh bằng ;
    goEvery: 0,                // chèn GO sau mỗi N câu lệnh (0 = tắt)
    whereColumns: 1,           // số cột đầu dùng làm điều kiện WHERE (UPDATE/SELECT)
    selectWhereMode: 'matchRows', // matchRows | inAnd | inOr
    sqlFormat: 'pretty',       // pretty = V1.3.0, compact = V1.2.2
    columnTypes: []            // auto | text | number theo index cột trên lưới
  };

  function mergeOptions(opts) {
    var o = {}, k;
    for (k in DEFAULT_OPTIONS) {
      if (Object.prototype.hasOwnProperty.call(DEFAULT_OPTIONS, k)) o[k] = DEFAULT_OPTIONS[k];
    }
    if (opts) {
      for (k in opts) {
        if (Object.prototype.hasOwnProperty.call(opts, k) && opts[k] !== undefined) o[k] = opts[k];
      }
    }
    return o;
  }

  /* ---------------------------------------------------------------
   * Tiện ích giá trị
   * ------------------------------------------------------------- */

  // Tương đương IsNumeric() của VBA, nhưng cố ý KHÔNG chấp nhận dấu phân
  // cách hàng nghìn ("1,000") vì VBA cho ra SQL sai ở trường hợp đó.
  function looksNumeric(s) {
    return /^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(s);
  }

  // Số bắt đầu bằng 0 hoặc quá dài (mã BN, CCCD, số thẻ...) nên là chuỗi.
  function numericLosesMeaning(s) {
    var body = s.replace(/^[+-]/, '');
    if (/^0\d/.test(body)) return true;
    if (/^\d+$/.test(body) && body.length > 15) return true;
    return false;
  }

  function quoteIdent(name, d) {
    return d.open + d.escapeIdent(String(name == null ? '' : name)) + d.close;
  }

  function quoteString(text, d, o) {
    var prefix = (d.unicodePrefix && o.unicodePrefix) ? 'N' : '';
    return prefix + "'" + text.replace(/'/g, "''") + "'";
  }

  function normalizeColumnType(type) {
    return type === 'text' || type === 'number' ? type : 'auto';
  }

  function columnTypeAt(o, index) {
    return normalizeColumnType(o.columnTypes && o.columnTypes[index]);
  }

  function formatValue(raw, d, o, columnType) {
    var text = raw == null ? '' : String(raw);
    if (o.trimValues) text = text.replace(/^\s+|\s+$/g, '');
    if (text.length === 0) return 'NULL';
    if (o.nullKeyword && text.toUpperCase() === 'NULL') return 'NULL';
    columnType = normalizeColumnType(columnType);
    if (columnType === 'text') return quoteString(text, d, o);
    if (columnType === 'number') return text;
    if (looksNumeric(text)) {
      if (o.leadingZeroAsText && numericLosesMeaning(text)) return quoteString(text, d, o);
      return text;
    }
    return quoteString(text, d, o);
  }

  function warnNumberType(raw, row, column, o, warnings) {
    if (!warnings || column.type !== 'number') return;
    var text = raw == null ? '' : String(raw);
    if (o.trimValues) text = text.replace(/^\s+|\s+$/g, '');
    if (!text || (o.nullKeyword && text.toUpperCase() === 'NULL')) return;
    if (!looksNumeric(text)) {
      warnings.push('Dòng ' + row.sourceRow + ', cột "' + column.name + '": đang chọn kiểu Số nhưng giá trị "' + text + '" không phải số.');
    }
  }

  function formatColumnValue(raw, row, column, d, o, warnings) {
    warnNumberType(raw, row, column, o, warnings);
    return formatValue(raw, d, o, column.type);
  }

  function isBlank(v) {
    return v == null || String(v).replace(/^\s+|\s+$/g, '') === '';
  }

  /* ---------------------------------------------------------------
   * Chuẩn hoá dữ liệu bảng: dòng đầu là tên cột
   * ------------------------------------------------------------- */
  function prepare(table, matrix, o) {
    var warnings = [];
    var rows = matrix || [];

    if (isBlank(table)) warnings.push('Chưa có tên bảng. Điền tên bảng ở ô "Tên bảng".');

    if (!rows.length) {
      return { ok: false, warnings: ['Bảng chưa có dữ liệu.'], columns: [], body: [] };
    }

    // Cột hợp lệ = cột có tên ở dòng 1
    var headerRow = rows[0] || [];
    var columns = [];
    for (var j = 0; j < headerRow.length; j++) {
      if (!isBlank(headerRow[j])) {
        columns.push({
          index: j,
          name: String(headerRow[j]).replace(/^\s+|\s+$/g, ''),
          type: columnTypeAt(o, j)
        });
      }
    }

    if (!columns.length) {
      return { ok: false, warnings: ['Dòng đầu tiên phải là tên các cột.'], columns: [], body: [] };
    }

    var seen = {}, dupes = [];
    for (var c = 0; c < columns.length; c++) {
      var key = columns[c].name.toLowerCase();
      if (seen[key]) { if (dupes.indexOf(columns[c].name) < 0) dupes.push(columns[c].name); }
      seen[key] = true;
    }
    if (dupes.length) warnings.push('Tên cột bị trùng: ' + dupes.join(', '));

    // Các dòng dữ liệu
    var body = [];
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i] || [];
      var values = [];
      var allBlank = true;
      for (var k = 0; k < columns.length; k++) {
        var cell = row[columns[k].index];
        if (!isBlank(cell)) allBlank = false;
        values.push(cell);
      }
      if (o.skipEmptyRows && allBlank) continue;
      body.push({ sourceRow: i + 1, values: values });
    }

    if (!body.length) warnings.push('Không có dòng dữ liệu nào (chỉ có dòng tiêu đề).');

    return { ok: true, warnings: warnings, columns: columns, body: body };
  }

  /* ---------------------------------------------------------------
   * Ghép các câu lệnh + chèn GO
   * ------------------------------------------------------------- */
  function assemble(statements, d, o) {
    var separator = isPrettyFormat(o) ? '\n\n' : '\n';
    if (!o.goEvery || !d.supportsGo) return statements.join(separator);
    var out = [];
    for (var i = 0; i < statements.length; i++) {
      out.push(statements[i]);
      if ((i + 1) % o.goEvery === 0 && i < statements.length - 1) out.push('GO');
    }
    return out.join(separator);
  }

  function term(o) { return o.semicolon ? ';' : ''; }

  function isPrettyFormat(o) {
    return !o || o.sqlFormat !== 'compact';
  }

  var INDENT = '    ';

  function formatWrappedList(items) {
    return '(\n' + items.map(function (item) {
      return INDENT + item;
    }).join(',\n') + '\n)';
  }

  function formatTupleLine(values) {
    return INDENT + '(' + values.join(', ') + ')';
  }

  function formatInlineTuple(values) {
    return '(' + values.join(', ') + ')';
  }

  function formatAssignments(assignments) {
    return assignments.map(function (item) {
      return INDENT + item;
    }).join(',\n');
  }

  function formatConditions(conditions, joiner) {
    return conditions.map(function (item, index) {
      return INDENT + (index ? joiner + ' ' : '') + item;
    }).join('\n');
  }

  function formatPredicate(predicate) {
    return predicate
      .replace(/\) OR \(/g, ')\nOR (')
      .replace(/ OR (?=(\(|\[|`|"))/g, '\nOR ')
      .replace(/ AND (?=(\[|`|"))/g, '\nAND ')
      .split('\n')
      .map(function (line) { return INDENT + line; })
      .join('\n');
  }

  function getInWrapEvery(o) {
    var n = parseInt(o && o.inWrapEvery, 10) || 0;
    if (n < 1) return 0;
    return Math.min(n, 1000);
  }

  function formatInList(values, o) {
    var every = getInWrapEvery(o);
    if (!every || values.length <= every) return '(' + values.join(', ') + ')';
    var lines = [];
    for (var i = 0; i < values.length; i += every) {
      var line = values.slice(i, i + every).join(', ');
      if (i + every < values.length) line += ',';
      lines.push(INDENT + line);
    }
    return '(\n' + lines.join('\n') + '\n)';
  }

  function formatInCondition(col, values, hasNull, o) {
    var condition = col + ' IN ' + formatInList(values, o);
    return hasNull ? '(' + condition + ' OR ' + col + ' IS NULL)' : condition;
  }

  function formatSelectSql(qualified, predicate, o) {
    if (!isPrettyFormat(o)) {
      return 'SELECT * FROM ' + qualified + (predicate ? ' WHERE ' + predicate : '') + term(o);
    }
    var sql = 'SELECT\n' + INDENT + '*\nFROM\n' + INDENT + qualified;
    if (predicate) sql += '\nWHERE\n' + formatPredicate(predicate);
    return sql + term(o);
  }

  /* ---------------------------------------------------------------
   * INSERT  (Module1.GetInsertSQL)
   * ------------------------------------------------------------- */
  function buildInsert(table, matrix, options) {
    var o = mergeOptions(options);
    var d = DIALECTS[o.dialect] || DIALECTS.mssql;
    var prep = prepare(table, matrix, o);

    if (!prep.ok) return { sql: '', warnings: prep.warnings, stats: emptyStats() };

    var tableName = String(table || 'TABLE_NAME').replace(/^\s+|\s+$/g, '');
    var qualified = quoteIdent(tableName, d);

    var colList = [];
    for (var c = 0; c < prep.columns.length; c++) colList.push(quoteIdent(prep.columns[c].name, d));
    var pretty = isPrettyFormat(o);
    var intoClause = pretty
      ? 'INSERT INTO ' + qualified + '\n' + formatWrappedList(colList)
      : 'INSERT INTO ' + qualified + ' (' + colList.join(', ') + ')';

    var warnings = prep.warnings.slice();
    var tuples = [];
    for (var r = 0; r < prep.body.length; r++) {
      var row = prep.body[r];
      var vals = [];
      for (var v = 0; v < prep.body[r].values.length; v++) {
        vals.push(formatColumnValue(row.values[v], row, prep.columns[v], d, o, warnings));
      }
      tuples.push(vals);
    }

    var statements = [];
    if (o.multiRowInsert && tuples.length > 1 && d.maxRowsPerInsert > 1) {
      var size = d.maxRowsPerInsert;
      for (var s = 0; s < tuples.length; s += size) {
        var chunk = tuples.slice(s, s + size);
        statements.push(pretty
          ? intoClause + '\nVALUES\n' + chunk.map(formatTupleLine).join(',\n') + term(o)
          : intoClause + ' VALUES\n' + chunk.map(formatInlineTuple).join(',\n') + term(o));
      }
    } else {
      for (var t = 0; t < tuples.length; t++) {
        statements.push(pretty
          ? intoClause + '\nVALUES\n' + formatWrappedList(tuples[t]) + term(o)
          : intoClause + ' VALUES ' + formatInlineTuple(tuples[t]) + term(o));
      }
    }

    var parts = [];
    var useIdentity = o.identityInsert && d.supportsIdentityInsert && statements.length > 0;
    if (useIdentity) parts.push('SET IDENTITY_INSERT ' + qualified + ' ON' + term(o));
    parts.push(assemble(statements, d, o));
    if (useIdentity) parts.push('SET IDENTITY_INSERT ' + qualified + ' OFF' + term(o));

    return {
      sql: parts.join(pretty ? '\n\n' : '\n'),
      warnings: warnings,
      stats: {
        rows: prep.body.length,
        columns: prep.columns.length,
        statements: statements.length + (useIdentity ? 2 : 0)
      }
    };
  }

  /* ---------------------------------------------------------------
   * UPDATE  (Module2.UpdateSQL)
   * Cột 1..N = điều kiện WHERE, các cột còn lại = SET
   * ------------------------------------------------------------- */
  function buildUpdate(table, matrix, options) {
    var o = mergeOptions(options);
    var d = DIALECTS[o.dialect] || DIALECTS.mssql;
    var prep = prepare(table, matrix, o);

    if (!prep.ok) return { sql: '', warnings: prep.warnings, stats: emptyStats() };

    var warnings = prep.warnings.slice();
    var keyCount = Math.max(1, Math.min(o.whereColumns | 0 || 1, prep.columns.length));

    if (keyCount >= prep.columns.length) {
      warnings.push('Cần ít nhất một cột sau nhóm cột điều kiện để đưa vào mệnh đề SET.');
      return { sql: '', warnings: warnings, stats: emptyStats() };
    }

    var tableName = String(table || 'TABLE_NAME').replace(/^\s+|\s+$/g, '');
    var qualified = quoteIdent(tableName, d);

    var statements = [];
    for (var r = 0; r < prep.body.length; r++) {
      var row = prep.body[r];
      var sets = [], wheres = [], blankKey = false;

      for (var k = 0; k < keyCount; k++) {
        if (isBlank(row.values[k])) blankKey = true;
        var kv = formatColumnValue(row.values[k], row, prep.columns[k], d, o, warnings);
        wheres.push(quoteIdent(prep.columns[k].name, d) + (kv === 'NULL' ? ' IS NULL' : ' = ' + kv));
      }
      for (var j = keyCount; j < prep.columns.length; j++) {
        sets.push(quoteIdent(prep.columns[j].name, d) + ' = ' + formatColumnValue(row.values[j], row, prep.columns[j], d, o, warnings));
      }

      if (blankKey) {
        warnings.push('Dòng ' + row.sourceRow + ': cột điều kiện đang trống — câu lệnh sẽ dùng IS NULL.');
      }

      if (isPrettyFormat(o)) {
        statements.push('UPDATE ' + qualified + '\nSET\n' + formatAssignments(sets) +
                        '\nWHERE\n' + formatConditions(wheres, 'AND') + term(o));
      } else {
        statements.push('UPDATE ' + qualified + ' SET ' + sets.join(', ') +
                        ' WHERE ' + wheres.join(' AND ') + term(o));
      }
    }

    return {
      sql: assemble(statements, d, o),
      warnings: warnings,
      stats: { rows: prep.body.length, columns: prep.columns.length, statements: statements.length }
    };
  }

  /* ---------------------------------------------------------------
   * SELECT
   * whereColumns = 0: SELECT * FROM Table
   * matchRows: một dòng dùng =, nhiều dòng một cột dùng IN,
   *            nhiều cột dùng nhóm AND nối OR theo từng dòng
   * inAnd: mỗi cột WHERE là một IN, các cột nối AND
   * inOr: mỗi cột WHERE là một IN, các cột nối OR
   * ------------------------------------------------------------- */
  function buildSelect(table, matrix, options) {
    var o = mergeOptions(options);
    var d = DIALECTS[o.dialect] || DIALECTS.mssql;
    var tableName = String(table || 'TABLE_NAME').replace(/^\s+|\s+$/g, '');
    var qualified = quoteIdent(tableName, d);
    var warnings = [];

    if (isBlank(table)) warnings.push('Chưa có tên bảng. Điền tên bảng ở ô "Tên bảng".');

    var requestedKeys = Math.max(0, o.whereColumns | 0 || 0);
    if (!requestedKeys) {
      return {
        sql: formatSelectSql(qualified, '', o),
        warnings: warnings,
        stats: { rows: 0, columns: 0, statements: 1 }
      };
    }

    var prep = prepare(table, matrix, o);
    if (!prep.ok) return { sql: '', warnings: warnings.concat(prep.warnings), stats: emptyStats() };

    warnings = warnings.concat(prep.warnings);
    var keyCount = Math.min(requestedKeys, prep.columns.length);
    if (requestedKeys > prep.columns.length) {
      warnings.push('Số cột làm WHERE lớn hơn số cột có tên. App chỉ dùng ' + keyCount + ' cột đầu.');
    }
    if (!prep.body.length) {
      warnings.push('Cần ít nhất một dòng dữ liệu để sinh điều kiện WHERE cho SELECT.');
      return { sql: '', warnings: warnings, stats: emptyStats() };
    }

    var selectWhereMode = normalizeSelectWhereMode(o.selectWhereMode);
    var predicate = selectWhereMode === 'inAnd'
      ? buildColumnInSelectWhere(prep, d, o, keyCount, 'AND', warnings)
      : selectWhereMode === 'inOr'
        ? buildColumnInSelectWhere(prep, d, o, keyCount, 'OR', warnings)
        : keyCount === 1
          ? buildSingleColumnSelectWhere(prep, d, o, warnings)
          : buildMultiColumnSelectWhere(prep, d, o, keyCount, warnings);

    if (!predicate) {
      warnings.push('Không có giá trị WHERE hợp lệ cho SELECT. App đã bỏ mệnh đề WHERE.');
      return {
        sql: formatSelectSql(qualified, '', o),
        warnings: warnings,
        stats: { rows: prep.body.length, columns: prep.columns.length, statements: 1 }
      };
    }

    return {
      sql: formatSelectSql(qualified, predicate, o),
      warnings: warnings,
      stats: { rows: prep.body.length, columns: prep.columns.length, statements: 1 }
    };
  }

  function normalizeSelectWhereMode(mode) {
    return mode === 'inAnd' || mode === 'inOr' ? mode : 'matchRows';
  }

  function formatSelectWhereCell(raw, row, column, d, o, warnings) {
    if (o.skipEmptyRows && isBlank(raw)) return null;
    return formatColumnValue(raw, row, column, d, o, warnings);
  }

  function buildSingleColumnSelectWhere(prep, d, o, warnings) {
    var col = quoteIdent(prep.columns[0].name, d);
    if (prep.body.length === 1) {
      var onlyValue = formatSelectWhereCell(prep.body[0].values[0], prep.body[0], prep.columns[0], d, o, warnings);
      if (onlyValue == null) return '';
      return col + (onlyValue === 'NULL' ? ' IS NULL' : ' = ' + onlyValue);
    }

    var values = [];
    var seen = {};
    var hasNull = false;
    for (var r = 0; r < prep.body.length; r++) {
      var value = formatSelectWhereCell(prep.body[r].values[0], prep.body[r], prep.columns[0], d, o, warnings);
      if (value == null) continue;
      if (value === 'NULL') {
        hasNull = true;
      } else if (!seen[value]) {
        seen[value] = true;
        values.push(value);
      }
    }

    if (!values.length) return hasNull ? col + ' IS NULL' : '';
    return formatInCondition(col, values, hasNull, o);
  }

  function buildColumnInSelectWhere(prep, d, o, keyCount, joiner, warnings) {
    var conditions = [];
    for (var k = 0; k < keyCount; k++) {
      var condition = buildOneColumnInCondition(prep, d, o, k, warnings);
      if (condition) conditions.push(condition);
    }
    return conditions.join(' ' + joiner + ' ');
  }

  function buildOneColumnInCondition(prep, d, o, columnIndex, warnings) {
    var col = quoteIdent(prep.columns[columnIndex].name, d);
    var values = [];
    var seen = {};
    var hasNull = false;

    for (var r = 0; r < prep.body.length; r++) {
      var value = formatSelectWhereCell(prep.body[r].values[columnIndex], prep.body[r], prep.columns[columnIndex], d, o, warnings);
      if (value == null) continue;
      if (value === 'NULL') {
        hasNull = true;
      } else if (!seen[value]) {
        seen[value] = true;
        values.push(value);
      }
    }

    if (!values.length) return hasNull ? col + ' IS NULL' : '';
    return formatInCondition(col, values, hasNull, o);
  }

  function buildMultiColumnSelectWhere(prep, d, o, keyCount, warnings) {
    var groups = [];
    for (var r = 0; r < prep.body.length; r++) {
      var row = prep.body[r];
      var wheres = [];
      var blankKey = false;
      for (var k = 0; k < keyCount; k++) {
        if (isBlank(row.values[k])) blankKey = true;
        var value = formatSelectWhereCell(row.values[k], row, prep.columns[k], d, o, warnings);
        if (value == null) continue;
        wheres.push(quoteIdent(prep.columns[k].name, d) + (value === 'NULL' ? ' IS NULL' : ' = ' + value));
      }
      if (blankKey && !o.skipEmptyRows) {
        warnings.push('Dòng ' + row.sourceRow + ': cột điều kiện đang trống — SELECT sẽ dùng IS NULL.');
      }
      if (wheres.length) groups.push('(' + wheres.join(' AND ') + ')');
    }
    return groups.join(' OR ');
  }

  function emptyStats() { return { rows: 0, columns: 0, statements: 0 }; }

  function build(mode, table, matrix, options) {
    if (mode === 'update') return buildUpdate(table, matrix, options);
    if (mode === 'select') return buildSelect(table, matrix, options);
    return buildInsert(table, matrix, options);
  }

  global.SqlGen = {
    DIALECTS: DIALECTS,
    DEFAULT_OPTIONS: DEFAULT_OPTIONS,
    build: build,
    buildInsert: buildInsert,
    buildUpdate: buildUpdate,
    buildSelect: buildSelect,
    formatValue: formatValue,
    looksNumeric: looksNumeric
  };
})(typeof window !== 'undefined' ? window : this);
