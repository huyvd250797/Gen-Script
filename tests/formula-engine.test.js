/*!
 * formula-engine.test.js — Kiểm thử bộ sinh công thức (§194–199).
 * Chạy: node tests/formula-engine.test.js
 */
'use strict';

var path = require('path');
var vm = require('vm');
var fs = require('fs');

var sandbox = { window: {}, console: console };
sandbox.window.window = sandbox.window;
vm.createContext(sandbox);
vm.runInContext(
  fs.readFileSync(path.join(__dirname, '..', 'assets', 'js', 'formula-engine.js'), 'utf8'),
  sandbox
);
var E = sandbox.window.FormulaEngine;

var pass = 0, fail = 0;

function is(label, actual, expected) {
  if (actual === expected) {
    pass++;
    console.log('  ok   ' + label);
  } else {
    fail++;
    console.log('  FAIL ' + label);
    console.log('       mong đợi : ' + expected);
    console.log('       nhận được: ' + actual);
  }
}

function col(sheet, index, name) {
  return { mode: 'column', sheetName: sheet, colIndex: index, column: name };
}

function gen(cfg) {
  var r = E.generate(cfg);
  if (r.error) return 'LỖI: ' + r.error;
  return r.formula;
}

console.log('\n— Sheet name escaping (§22 §23 §196)');
is('tên sheet thường', E.escapeSheetName('Database'), 'Database');
is('tên sheet có khoảng trắng và gạch', E.escapeSheetName('CONFIG - Process'), "'CONFIG - Process'");
is('tên sheet bắt đầu bằng số', E.escapeSheetName('2026Data'), "'2026Data'");
is('tên sheet có nháy đơn', E.escapeSheetName("Sinh'Vien"), "'Sinh''Vien'");
is('tên sheet tiếng Việt có dấu', E.escapeSheetName('SinhVien'), 'SinhVien');
is('tên sheet trùng dạng ô', E.escapeSheetName('A1'), "'A1'");

console.log('\n— XLOOKUP dấu phẩy (§195)');
is('XLOOKUP comma', gen({
  type: 'lookup', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 2,
  lookupValue: col('SinhVien', 0, 'MaSinhVien'),
  lookupRange: col('Database', 1, 'MaSinhVien'),
  returnRange: col('Database', 0, 'IDSinhVien'),
  notFound: ''
}), '=XLOOKUP(A2,Database!B:B,Database!A:A,"")');

console.log('\n— XLOOKUP dấu chấm phẩy (§197 §203)');
is('XLOOKUP semicolon', gen({
  type: 'lookup', platform: 'excel', separator: 'semicolon', currentSheet: 'SinhVien', dataRow: 2,
  lookupValue: col('SinhVien', 0, 'MaSinhVien'),
  lookupRange: col('Database', 1, 'MaSinhVien'),
  returnRange: col('Database', 0, 'IDSinhVien'),
  notFound: ''
}), '=XLOOKUP(A2;Database!B:B;Database!A:A;"")');

console.log('\n— Sheet có khoảng trắng trong công thức (§196 §204)');
is('lookup sheet có khoảng trắng', gen({
  type: 'lookup', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 2,
  lookupValue: col('SinhVien', 0, 'MaSinhVien'),
  lookupRange: col('CONFIG - Process', 0, 'Ma'),
  returnRange: col('CONFIG - Process', 3, 'ID'),
  notFound: ''
}), '=XLOOKUP(A2,\'CONFIG - Process\'!A:A,\'CONFIG - Process\'!D:D,"")');

console.log('\n— Dòng dữ liệu bắt đầu (§30)');
is('data start = 4', gen({
  type: 'lookup', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 4,
  lookupValue: col('SinhVien', 0, 'MaSinhVien'),
  lookupRange: col('Database', 1, 'MaSinhVien'),
  returnRange: col('Database', 0, 'IDSinhVien'),
  notFound: ''
}), '=XLOOKUP(A4,Database!B:B,Database!A:A,"")');

console.log('\n— Chuỗi đặc biệt trong kết quả (§199)');
is('fallback tiếng Việt', gen({
  type: 'lookup', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 2,
  lookupValue: col('SinhVien', 0, 'MaSinhVien'),
  lookupRange: col('Database', 1, 'Ma'),
  returnRange: col('Database', 0, 'ID'),
  notFound: 'Không tìm thấy'
}), '=XLOOKUP(A2,Database!B:B,Database!A:A,"Không tìm thấy")');
is('fallback có nháy kép', gen({
  type: 'lookup', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 2,
  lookupValue: col('SinhVien', 0, 'Ma'),
  lookupRange: col('Database', 1, 'Ma'),
  returnRange: col('Database', 0, 'ID'),
  notFound: 'không có "mã"'
}), '=XLOOKUP(A2,Database!B:B,Database!A:A,"không có ""mã""")');

console.log('\n— VLOOKUP cảnh báo cột trả về bên trái (§34)');
(function () {
  var r = E.generate({
    type: 'lookup', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 2,
    method: 'vlookup',
    lookupValue: col('SinhVien', 0, 'MaSinhVien'),
    lookupRange: col('Database', 1, 'MaSinhVien'),
    returnRange: col('Database', 0, 'IDSinhVien'),
    notFound: ''
  });
  is('có cảnh báo VLOOKUP', r.warnings.some(function (w) { return /bên trái/.test(w); }), true);
})();

console.log('\n— COUNTIF kiểm tra trùng (§37 §223)');
is('duplicate 1 cột', gen({
  type: 'duplicate', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 2,
  fields: [col('SinhVien', 0, 'MaSinhVien')],
  duplicateLabel: 'TRÙNG', uniqueLabel: ''
}), '=IF(COUNTIF($A:$A,A2)>1,"TRÙNG","")');

console.log('\n— COUNTIFS nhiều cột (§38 §198)');
is('duplicate 3 cột comma', gen({
  type: 'duplicate', platform: 'excel', separator: 'comma', currentSheet: 'KetQua', dataRow: 2,
  fields: [col('KetQua', 0, 'MaSinhVien'), col('KetQua', 1, 'MaMonHoc'), col('KetQua', 2, 'HocKy')],
  duplicateLabel: 'TRÙNG', uniqueLabel: ''
}), '=IF(COUNTIFS($A:$A,A2,$B:$B,B2,$C:$C,C2)>1,"TRÙNG","")');
is('duplicate 2 cột semicolon', gen({
  type: 'duplicate', platform: 'excel', separator: 'semicolon', currentSheet: 'KetQua', dataRow: 2,
  fields: [col('KetQua', 0, 'A'), col('KetQua', 1, 'B')],
  resultType: 'count'
}), '=COUNTIFS($A:$A;A2;$B:$B;B2)');

console.log('\n— Chỉ đếm số lần (§39)');
is('duplicate count', gen({
  type: 'duplicate', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 2,
  fields: [col('SinhVien', 0, 'MaSinhVien')], resultType: 'count'
}), '=COUNTIF($A:$A,A2)');

console.log('\n— Kiểm tra tồn tại (§40)');
is('existence COUNTIF', gen({
  type: 'existence', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 2,
  value: col('SinhVien', 0, 'MaSinhVien'),
  reference: col('Database', 1, 'MaSinhVien'),
  foundLabel: 'CÓ', missingLabel: 'KHÔNG'
}), '=IF(COUNTIF(Database!B:B,A2)>0,"CÓ","KHÔNG")');

console.log('\n— So sánh dữ liệu (§43 §44 §45)');
is('compare thường', gen({
  type: 'compare', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  valueA: col('S', 0, 'A'), valueB: col('S', 1, 'B')
}), '=IF(A2=B2,"GIỐNG","KHÁC")');
is('compare trim', gen({
  type: 'compare', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  valueA: col('S', 0, 'A'), valueB: col('S', 1, 'B'), trim: true
}), '=IF(TRIM(A2)=TRIM(B2),"GIỐNG","KHÁC")');
is('compare trim + lower', gen({
  type: 'compare', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  valueA: col('S', 0, 'A'), valueB: col('S', 1, 'B'), trim: true, ignoreCase: true
}), '=IF(LOWER(TRIM(A2))=LOWER(TRIM(B2)),"GIỐNG","KHÁC")');

console.log('\n— So sánh hai danh sách (§47)');
is('list compare', gen({
  type: 'listCompare', platform: 'excel', separator: 'comma', currentSheet: 'SheetA', dataRow: 2,
  value: col('SheetA', 0, 'Ma'), reference: col('SheetB', 0, 'Ma')
}), '=IF(COUNTIF(SheetB!A:A,A2)>0,"CÓ","THIẾU")');

console.log('\n— Kiểm tra rỗng (§76)');
is('blank check 4 cột', gen({
  type: 'blank', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 2,
  fields: [col('SinhVien', 0, 'A'), col('SinhVien', 1, 'B'), col('SinhVien', 2, 'C'), col('SinhVien', 3, 'D')],
  missingLabel: 'THIẾU DỮ LIỆU', okLabel: ''
}), '=IF(OR(A2="",B2="",C2="",D2=""),"THIẾU DỮ LIỆU","")');

console.log('\n— IF đơn và IFS nhiều bậc (§53 §54)');
is('IF đơn', gen({
  type: 'logic', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  branches: [{ condition: { kind: 'cond', field: col('S', 0, 'Diem'), operator: '>=', value: '5' }, result: 'Đạt' }],
  elseResult: 'Không đạt'
}), '=IF(A2>=5,"Đạt","Không đạt")');
is('IFS nhiều bậc', gen({
  type: 'logic', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  branches: [
    { condition: { kind: 'cond', field: col('S', 0, 'Diem'), operator: '>=', value: '8.5' }, result: 'Giỏi' },
    { condition: { kind: 'cond', field: col('S', 0, 'Diem'), operator: '>=', value: '7' }, result: 'Khá' },
    { condition: { kind: 'cond', field: col('S', 0, 'Diem'), operator: '>=', value: '5' }, result: 'Trung bình' }
  ],
  elseResult: 'Không đạt'
}), '=IFS(A2>=8.5,"Giỏi",A2>=7,"Khá",A2>=5,"Trung bình",TRUE,"Không đạt")');

console.log('\n— AND/OR lồng nhau (§17 §55)');
is('nested AND/OR', gen({
  type: 'logic', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  branches: [{
    condition: {
      kind: 'group', operator: 'AND', children: [
        { kind: 'cond', field: col('S', 0, 'MaSinhVien'), operator: 'notBlank' },
        { kind: 'cond', field: col('S', 1, 'TrangThai'), operator: '=', value: 'Đang học' },
        {
          kind: 'group', operator: 'OR', children: [
            { kind: 'cond', field: col('S', 2, 'MaLop'), operator: '=', value: 'CNTT' },
            { kind: 'cond', field: col('S', 2, 'MaLop'), operator: '=', value: 'KTPM' }
          ]
        }
      ]
    },
    result: 'Hợp lệ'
  }],
  elseResult: 'Không hợp lệ'
}), '=IF(AND(A2<>"",B2="Đang học",OR(C2="CNTT",C2="KTPM")),"Hợp lệ","Không hợp lệ")');

console.log('\n— Làm sạch Text (§57–64)');
var textCases = [
  ['trim', '=TRIM(A2)'],
  ['clean', '=CLEAN(A2)'],
  ['trimClean', '=TRIM(CLEAN(A2))'],
  ['upper', '=UPPER(TRIM(A2))'],
  ['lower', '=LOWER(TRIM(A2))'],
  ['proper', '=PROPER(TRIM(A2))'],
  ['removeSpace', '=SUBSTITUTE(A2," ","")'],
  ['removeLineBreak', '=SUBSTITUTE(A2,CHAR(10)," ")']
];
textCases.forEach(function (c) {
  is('text ' + c[0], gen({
    type: 'text', preset: c[0], platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
    field: col('S', 0, 'HoTen')
  }), c[1]);
});
is('substitute builder (§65)', gen({
  type: 'text', preset: 'substitute', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  field: col('S', 0, 'X'), find: '-', replace: '/'
}), '=SUBSTITUTE(A2,"-","/")');
is('right 5 ký tự (§67)', gen({
  type: 'text', preset: 'right', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  field: col('S', 0, 'X'), count: 5
}), '=RIGHT(A2,5)');

console.log('\n— Ghép dữ liệu (§68)');
is('TEXTJOIN', gen({
  type: 'concat', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  fields: [col('S', 0, 'HoDem'), col('S', 1, 'Ten')], joinWith: 'space', ignoreBlank: true
}), '=TEXTJOIN(" ",TRUE,A2,B2)');
is('nối bằng &', gen({
  type: 'concat', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  fields: [col('S', 0, 'A'), col('S', 1, 'B')], joinWith: 'dash', method: 'ampersand'
}), '=A2&"-"&B2');

console.log('\n— Xử lý ngày (§71–75)');
is('format date', gen({
  type: 'date', preset: 'format', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  field: col('S', 0, 'NgaySinh'), format: 'dd/mm/yyyy'
}), '=TEXT(A2,"dd/mm/yyyy")');
is('tính tuổi', gen({
  type: 'date', preset: 'age', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  field: col('S', 0, 'NgaySinh')
}), '=DATEDIF(A2,TODAY(),"Y")');
is('số ngày giữa hai mốc', gen({
  type: 'date', preset: 'daysBetween', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  field: col('S', 0, 'A'), field2: col('S', 1, 'B')
}), '=DATEDIF(A2,B2,"D")');
is('lấy năm', gen({
  type: 'date', preset: 'year', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  field: col('S', 0, 'A')
}), '=YEAR(A2)');
is('năm học', gen({
  type: 'date', preset: 'academicYear', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  field: col('S', 0, 'A'), startMonth: 8
}), '=IF(MONTH(A2)>=8,YEAR(A2)&"-"&YEAR(A2)+1,YEAR(A2)-1&"-"&YEAR(A2))');

console.log('\n— SUMIFS (§79)');
is('SUMIFS 2 điều kiện', gen({
  type: 'aggregate', kind: 'sum', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 2,
  sumField: col('Database', 3, 'HocPhi'),
  conditions: [
    { field: col('Database', 0, 'MaSinhVien'), operator: '=', value: 'A2' },
    { field: col('Database', 1, 'HocKy'), operator: '=', value: 'B2' }
  ]
}), '=SUMIFS(Database!$D:$D,Database!$A:$A,A2,Database!$B:$B,B2)');

console.log('\n— UNIQUE / SORT (§80)');
is('unique', gen({
  type: 'unique', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2, usedRows: 10000,
  field: col('S', 0, 'A')
}), '=UNIQUE(A2:A10000)');
is('sort unique', gen({
  type: 'unique', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2, usedRows: 10000,
  field: col('S', 0, 'A'), sort: true
}), '=SORT(UNIQUE(A2:A10000))');

console.log('\n— SQL Tools (§82 §83)');
is('SQL quote value', gen({
  type: 'sqlQuote', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  field: col('S', 0, 'Ma')
}), '="\'"&A2&"\'"');
is('SQL quote số giữ nguyên (§154)', gen({
  type: 'sqlQuote', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  field: col('S', 0, 'ID'), dataType: 'number'
}), '=A2');
is('SQL IN', gen({
  type: 'sqlIn', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2, usedRows: 100,
  field: col('S', 0, 'Ma')
}), '=TEXTJOIN(",",TRUE,"\'"&A2:A100&"\'")');

console.log('\n— FILTER nhiều điều kiện OR (§49)');
is('filter contains OR', gen({
  type: 'filter', platform: 'excel', separator: 'comma', currentSheet: 'KetQua', dataRow: 2,
  sourceSheet: 'SinhVien', fromCol: 0, toCol: 6, logic: 'OR',
  conditions: [
    { field: col('SinhVien', 5, 'MaLop'), operator: 'contains', value: 'F3' },
    { field: col('SinhVien', 5, 'MaLop'), operator: 'contains', value: 'F4' }
  ]
}), '=FILTER(SinhVien!A:G,(ISNUMBER(SEARCH(F3,SinhVien!F:F)))+(ISNUMBER(SEARCH(F4,SinhVien!F:F))),"")');

console.log('\n— Google Sheets QUERY (§50 §224)');
is('QUERY contains OR', gen({
  type: 'filter', platform: 'sheets', separator: 'auto', currentSheet: 'Main', dataRow: 2,
  sourceSheet: 'CONFIG - Process', returnRange: 'A6:G', logic: 'OR', sheetsMethod: 'query',
  conditions: [
    { field: col('CONFIG - Process', 5, 'F'), operator: 'contains', value: 'F3' },
    { field: col('CONFIG - Process', 5, 'F'), operator: 'contains', value: 'F4' }
  ]
}), '=QUERY(\'CONFIG - Process\'!A6:G,"select * where F contains \'"&F3&"\' or F contains \'"&F4&"\'",0)');

console.log('\n— Kiểm tra ngoặc (§128 §156)');
is('ngoặc cân đối', E.validateParens('=IF(A2>1,"x","y")').ok, true);
is('thiếu ngoặc', E.validateParens('=IF(A2>1,"x","y"').ok, false);
is('dư ngoặc', E.validateParens('=IF(A2>1,"x","y"))').ok, false);
is('nháy kép lẻ bị phát hiện', E.validateParens('=IF(A2>1,"x,"y")').ok, false);

console.log('\n— Formatted view (§87)');
(function () {
  var r = E.generate({
    type: 'lookup', platform: 'excel', separator: 'comma', currentSheet: 'SinhVien', dataRow: 2,
    lookupValue: col('SinhVien', 0, 'Ma'),
    lookupRange: col('Database', 1, 'Ma'),
    returnRange: col('Database', 0, 'ID'),
    notFound: ''
  });
  is('formatted xuống dòng', r.formattedFormula,
    '=XLOOKUP(\n  A2,\n  Database!B:B,\n  Database!A:A,\n  ""\n)');
})();

console.log('\n— Lỗi cấu hình (§120 §173)');
(function () {
  var r = E.generate({
    type: 'lookup', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
    lookupValue: col('S', 0, 'Ma'),
    lookupRange: col('Database', 1, 'Ma')
  });
  is('thiếu cột trả về', r.error, 'Vui lòng chọn cột trả về.');
  is('không sinh công thức lỗi', r.formula, '');
})();

console.log('\n— Manual mode (§123 §124 §175)');
is('nhập tay tham chiếu', gen({
  type: 'lookup', platform: 'excel', separator: 'comma', dataRow: 2,
  lookupValue: { mode: 'cell', raw: '$F$3' },
  lookupRange: { mode: 'custom', raw: 'Database!B:B' },
  returnRange: { mode: 'custom', raw: 'Database!A:A' },
  notFound: ''
}), '=XLOOKUP($F$3,Database!B:B,Database!A:A,"")');

console.log('\n— Reference type (§125)');
is('absolute cell', gen({
  type: 'compare', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  referenceType: 'absoluteCell',
  valueA: col('S', 0, 'A'), valueB: col('S', 1, 'B')
}), '=IF($A$2=$B$2,"GIỐNG","KHÁC")');

console.log('\n— Used range (§93)');
is('used range trong lookup', gen({
  type: 'lookup', platform: 'excel', separator: 'comma', currentSheet: 'S', dataRow: 2,
  rangeMode: 'used', usedRows: 500,
  lookupValue: col('S', 0, 'Ma'),
  lookupRange: col('Database', 1, 'Ma'),
  returnRange: col('Database', 0, 'ID'),
  notFound: ''
}), '=XLOOKUP(A2,Database!B$2:B$500,Database!A$2:A$500,"")');

console.log('\n' + (fail === 0 ? '✔ ' : '✘ ') + pass + ' đạt, ' + fail + ' hỏng.\n');
process.exit(fail ? 1 : 0);
