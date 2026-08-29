/*!
 * app-smoke.test.js — Kiểm thử tích hợp: nạp index.html thật, chuyển workspace,
 * mở công cụ và kiểm tra công thức sinh ra. Bao phủ §200 §201 §205.
 */
'use strict';

var fs = require('fs');
var path = require('path');
var { JSDOM } = require('jsdom');

var ROOT = path.join(__dirname, '..');
var pass = 0, fail = 0;

function is(label, actual, expected) {
  if (actual === expected) { pass++; console.log('  ok   ' + label); }
  else {
    fail++;
    console.log('  FAIL ' + label);
    console.log('       mong đợi : ' + expected);
    console.log('       nhận được: ' + actual);
  }
}
function ok(label, cond) { is(label, !!cond, true); }

var html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
// Bỏ SheetJS (không cần cho smoke test) và font ngoài
html = html.replace('<script src="assets/vendor/xlsx.core.min.js"></script>', '');

var dom = new JSDOM(html, {
  runScripts: 'outside-only',
  url: 'http://localhost/',
  pretendToBeVisual: true
});
var win = dom.window;

// dialog chưa được jsdom hỗ trợ đầy đủ
win.HTMLDialogElement = win.HTMLDialogElement || function () {};
['helpDialog', 'formulaHelpDialog', 'scriptDialog'].forEach(function (id) {
  var d = win.document.getElementById(id);
  if (d && !d.showModal) {
    d.showModal = function () { this.open = true; this.setAttribute('open', ''); };
    d.close = function () { this.open = false; this.removeAttribute('open'); };
  }
});

// Chờ jsdom parse xong rồi mới nạp script, để app.js chỉ khởi tạo đúng một lần.
win.addEventListener('load', function () { setTimeout(main, 0); });

function main() {
var scripts = [
  'assets/js/generator.js',
  'assets/js/grid.js',
  'assets/js/workbook.js',
  'assets/js/formula-engine.js',
  'assets/js/formula-presets.js',
  'assets/js/formula-explainer.js',
  'assets/js/formula-history.js',
  'assets/js/formula-helper.js',
  'assets/js/app.js'
];

console.log('\n— Nạp toàn bộ script');
try {
  scripts.forEach(function (f) {
    win.eval(fs.readFileSync(path.join(ROOT, f), 'utf8'));
  });
  console.log('  ok   không lỗi runtime khi khởi tạo');
  pass++;
} catch (err) {
  console.log('  FAIL lỗi khi nạp script: ' + err.message);
  console.log(err.stack);
  fail++;
  process.exit(1);
}

var doc = win.document;

/** jsdom không cho Object.keys(localStorage) nên duyệt bằng key(i). */
function lsKey(fragment) {
  for (var i = 0; i < win.localStorage.length; i++) {
    var k = win.localStorage.key(i);
    if (k && k.indexOf(fragment) !== -1) return k;
  }
  return null;
}

console.log('\n— Không regression Gen Script (§201)');
ok('SqlGen còn nguyên', typeof win.SqlGen.build === 'function');
ok('Grid khởi tạo được', !!doc.querySelector('.grid'));
ok('nút INSERT/UPDATE/SELECT còn', doc.querySelectorAll('.mode-btn').length === 3);
is('sinh INSERT vẫn chạy',
  win.SqlGen.build('insert', 'DM_Test', [['Ma', 'Ten'], ['1', 'A']], { dialect: 'mssql', identityInsert: false })
    .sql.indexOf('INSERT INTO [DM_Test]') === 0, true);
ok('tab bảng render', doc.querySelectorAll('.sheet-tabs .tab').length >= 1);
ok('phiên bản hiển thị V2.5.0', doc.querySelector('.app-version').textContent === 'V2.5.0');
var wrappedIn = win.SqlGen.build('select', 'DM_Test', [
  ['ID'], ['1'], ['2'], ['3'], ['4'], ['5'], ['6']
], {
  dialect: 'mssql',
  whereColumns: 1,
  selectWhereMode: 'inAnd',
  inWrapEvery: 5
}).sql;
ok('SELECT IN xuống dòng theo số lượng cấu hình', wrappedIn.indexOf('[ID] IN (') !== -1 && wrappedIn.indexOf('1, 2, 3, 4, 5,\n') !== -1);
ok('SELECT IN không có dấu phẩy dư trước dấu đóng ngoặc', wrappedIn.indexOf('6,\n    )') === -1);

console.log('\n— Workspace switch (§7 §200)');
var wsButtons = doc.querySelectorAll('.ws-btn');
is('có 2 workspace', wsButtons.length, 2);
ok('Gen Script đang active', wsButtons[0].classList.contains('is-on'));
ok('pane Formula đang ẩn', doc.getElementById('workspaceFormula').hidden);

// Nạp dữ liệu vào workbook riêng của Formula Helper giống như import Excel
win.Workbook.setFileName('SinhVien.xlsx');

var sheets = win.Workbook.getSheets();
sheets.push({
  name: 'SinhVien',
  data: [
    ['MaSinhVien', 'HoTen', 'MaLop', 'NgaySinh'],
    ['SV001', 'Nguyễn Văn A', 'CNTT', '2004-01-05'],
    ['SV002', 'Trần Thị B', 'KTPM', '2003-11-20'],
    ['SV003', 'Lê Văn C', 'CNTT', '2004-06-30'],
    ['SV001', 'Nguyễn Văn A', 'CNTT', '2004-01-05']
  ],
  whereColumns: 1, selectWhereColumns: 0, selectWhereMode: 'matchRows', columnTypes: []
});
sheets.push({
  name: 'CONFIG - Process',
  data: [
    ['IDSinhVien', 'MaSinhVien', 'TrangThai'],
    ['125', 'SV001', 'Đang học'],
    ['128', 'SV002', 'Đang học']
  ],
  whereColumns: 1, selectWhereColumns: 0, selectWhereMode: 'matchRows', columnTypes: []
});

console.log('\n— Formula Workbook riêng (§10 §200)');
is('Formula workbook đọc được 3 sheet', win.Workbook.getSheets().length, 3);
is('tên cột lấy từ dòng đầu', win.Workbook.getHeaders(1)[0].name, 'MaSinhVien');
is('số dòng dữ liệu', win.Workbook.getDataRowCount(1), 4);
var flat = win.Workbook.getAllFields().map(function (f) { return f.label; });
ok('danh sách field phẳng có SinhVien.MaSinhVien', flat.indexOf('SinhVien.MaSinhVien') !== -1);
ok('sheet có khoảng trắng nằm trong danh sách', flat.indexOf('CONFIG - Process.IDSinhVien') !== -1);
ok('Gen Script không tự nhận sheet Formula',
  Array.prototype.every.call(doc.querySelectorAll('.sheet-tabs .tab-name'), function (n) {
    return n.textContent !== 'SinhVien';
  }));

console.log('\n— Chuyển sang Formula Helper');
wsButtons[1].dispatchEvent(new win.Event('click', { bubbles: true }));
ok('pane Gen Script ẩn đi', doc.getElementById('workspaceGenScript').hidden);
ok('pane Formula hiện ra', !doc.getElementById('workspaceFormula').hidden);
ok('nút Formula Helper active', wsButtons[1].classList.contains('is-on'));
ok('Workbook không mất dữ liệu', win.Workbook.getSheets().length === 3);

console.log('\n— Formula Helper Home (§14 §134)');
var cards = doc.querySelectorAll('.fh-card');
ok('có thẻ công cụ', cards.length >= 12);
ok('panel review ẩn mặc định', doc.querySelector('.fh-review-overlay').hidden);
ok('có nút xem kết quả', Array.prototype.some.call(doc.querySelectorAll('.fh-bar-actions button'), function (b) {
  return b.textContent === 'Xem kết quả';
}));
ok('Home ẩn lưới dữ liệu', doc.querySelector('.fh-data-grid-panel').hidden);
ok('Home chưa hiện nút sidebar công cụ', Array.prototype.every.call(doc.querySelectorAll('.fh-bar-actions button'), function (b) {
  return b.hidden || b.textContent !== 'Ẩn thanh công cụ';
}));
ok('panel Workbook hiện tên file', doc.querySelector('.fh-wb-file strong').textContent === 'SinhVien.xlsx');
is('liệt kê đủ sheet', doc.querySelectorAll('.fh-wb-sheet').length, 3);

console.log('\n— Thư viện hàm có ví dụ mô phỏng');
Array.prototype.filter.call(doc.querySelectorAll('.fh-bar-actions button'), function (b) {
  return b.textContent === 'Thư viện hàm';
})[0].dispatchEvent(new win.Event('click', { bubbles: true }));
ok('mở được thư viện hàm', !!doc.querySelector('.fh-library'));
ok('thư viện có ví dụ dễ hiểu', doc.querySelectorAll('.fh-lib-demo').length >= 20);
ok('ví dụ có kết quả mô phỏng', Array.prototype.some.call(doc.querySelectorAll('.fh-lib-demo em'), function (n) {
  return n.textContent.indexOf('Kết quả:') === 0;
}));
var libSearch = doc.querySelector('.fh-search input');
libSearch.value = 'Không có';
libSearch.dispatchEvent(new win.Event('input', { bubbles: true }));
ok('tìm được theo nội dung mô phỏng', Array.prototype.some.call(doc.querySelectorAll('.fh-lib-item h4'), function (n) {
  return n.textContent === 'XLOOKUP';
}));
doc.querySelector('.fh-back').dispatchEvent(new win.Event('click', { bubbles: true }));

console.log('\n— Đổi Sheet ngay trong Formula Helper (§111)');
Array.prototype.filter.call(doc.querySelectorAll('.fh-bar-actions button'), function (b) {
  return b.textContent === 'Xem kết quả';
})[0].dispatchEvent(new win.Event('click', { bubbles: true }));
ok('mở được panel xem kết quả', !doc.querySelector('.fh-review-overlay').hidden);
var sheetBtns = doc.querySelectorAll('.fh-wb-sheet');
sheetBtns[1].dispatchEvent(new win.Event('click', { bubbles: true }));
is('sheet đang mở là SinhVien', win.Workbook.getSheetName(win.Workbook.getActiveIndex()), 'SinhVien');

console.log('\n— Tìm kiếm công cụ (§109)');
var search = doc.querySelector('.fh-search input');
search.value = 'trùng';
search.dispatchEvent(new win.Event('input', { bubbles: true }));
var results = doc.querySelectorAll('.fh-card strong');
ok('tìm "trùng" ra Kiểm tra trùng',
  Array.prototype.some.call(results, function (n) { return n.textContent === 'Kiểm tra trùng'; }));
search.value = 'xlookup';
search.dispatchEvent(new win.Event('input', { bubbles: true }));
ok('tìm "xlookup" ra Tra cứu dữ liệu',
  Array.prototype.some.call(doc.querySelectorAll('.fh-card strong'), function (n) {
    return n.textContent === 'Tra cứu dữ liệu';
  }));

console.log('\n— Dựng công thức tra cứu qua UI (§222)');
win.FormulaHelper.openTool('lookup');
ok('chọn công cụ thì hiện lưới', !doc.querySelector('.fh-data-grid-panel').hidden);
ok('lưới dữ liệu render được trong tool view', !!doc.querySelector('.fh-data-grid-host table.grid'));
ok('workbench ở chế độ tool', doc.querySelector('.fh-workarea').classList.contains('is-tool'));
var sidebarToggle = Array.prototype.filter.call(doc.querySelectorAll('.fh-bar-actions button'), function (b) {
  return b.textContent === 'Ẩn thanh công cụ';
})[0];
ok('có nút ẩn thanh công cụ', !!sidebarToggle);
sidebarToggle.dispatchEvent(new win.Event('click', { bubbles: true }));
ok('ẩn được sidebar công cụ', doc.querySelector('.fh-tool-dock').classList.contains('is-collapsed'));
sidebarToggle = Array.prototype.filter.call(doc.querySelectorAll('.fh-bar-actions button'), function (b) {
  return b.textContent === 'Hiện thanh công cụ';
})[0];
sidebarToggle.dispatchEvent(new win.Event('click', { bubbles: true }));
ok('hiện lại được sidebar công cụ', !doc.querySelector('.fh-tool-dock').classList.contains('is-collapsed'));
is('lưới đang mở đúng sheet active', doc.querySelector('.fh-data-grid-meta strong').textContent, win.Workbook.getSheetName(win.Workbook.getActiveIndex()));
var refSelects = doc.querySelectorAll('.fh-ref-select');
ok('có 3 ô chọn cột', refSelects.length >= 3);

function pick(select, label) {
  var options = select.querySelectorAll('option');
  for (var i = 0; i < options.length; i++) {
    if (options[i].textContent.indexOf(label) === 0) {
      select.value = options[i].value;
      select.dispatchEvent(new win.Event('change', { bubbles: true }));
      return true;
    }
  }
  return false;
}

ok('chọn được SinhVien.MaSinhVien', pick(refSelects[0], 'SinhVien.MaSinhVien'));
refSelects = doc.querySelectorAll('.fh-ref-select');
ok('chọn được CONFIG - Process.MaSinhVien', pick(refSelects[1], 'CONFIG - Process.MaSinhVien'));
refSelects = doc.querySelectorAll('.fh-ref-select');
ok('chọn được CONFIG - Process.IDSinhVien', pick(refSelects[2], 'CONFIG - Process.IDSinhVien'));

var code = doc.querySelector('.fh-formula-code');
is('công thức sinh đúng, sheet có khoảng trắng được bọc nháy',
  code.textContent,
  '=XLOOKUP(A2,\'CONFIG - Process\'!B:B,\'CONFIG - Process\'!A:A,"")');

console.log('\n— Đổi dấu phân cách (§203)');
var sepSelect = doc.querySelectorAll('.fh-bar select')[1];
sepSelect.value = 'semicolon';
sepSelect.dispatchEvent(new win.Event('change', { bubbles: true }));
is('công thức dùng chấm phẩy',
  doc.querySelector('.fh-formula-code').textContent,
  '=XLOOKUP(A2;\'CONFIG - Process\'!B:B;\'CONFIG - Process\'!A:A;"")');
is('lưu lại setting', JSON.parse(win.localStorage.getItem(lsKey('formulaSettings'))).separator, 'semicolon');
sepSelect.value = 'comma';
sepSelect.dispatchEvent(new win.Event('change', { bubbles: true }));

console.log('\n— Giải thích và xem trước (§89 §95)');
ok('có phần giải thích', doc.querySelectorAll('.fh-explain dt').length >= 3);
is('giải thích ô đầu tiên', doc.querySelector('.fh-explain dt').textContent, 'A2');

console.log('\n— Đổi dòng dữ liệu bắt đầu (§30)');
var rowInput = doc.querySelector('.fh-field-num input');
rowInput.value = '4';
rowInput.dispatchEvent(new win.Event('input', { bubbles: true }));
ok('công thức đổi sang A4', doc.querySelector('.fh-formula-code').textContent.indexOf('=XLOOKUP(A4,') === 0);
rowInput.value = '2';
rowInput.dispatchEvent(new win.Event('input', { bubbles: true }));

console.log('\n— Kiểm tra trùng qua UI (§223)');
win.FormulaHelper.openTool('duplicate');
ok('đổi công cụ vẫn giữ dữ liệu lưới', doc.querySelector('.fh-data-grid-meta strong').textContent === 'SinhVien');
pick(doc.querySelector('.fh-ref-select'), 'SinhVien.MaSinhVien');
is('công thức COUNTIF',
  doc.querySelector('.fh-formula-code').textContent,
  '=IF(COUNTIF($A:$A,A2)>1,"TRÙNG","")');

console.log('\n— Quay lại chọn công cụ khác');
doc.querySelector('.fh-back').dispatchEvent(new win.Event('click', { bubbles: true }));
ok('quay lại Home thì ẩn lưới', doc.querySelector('.fh-data-grid-panel').hidden);
ok('Home hiện lại danh sách công cụ', doc.querySelectorAll('.fh-card').length >= 12);
win.FormulaHelper.openTool('duplicate');
ok('chọn lại công cụ thì lưới hiện lại', !doc.querySelector('.fh-data-grid-panel').hidden);
is('dữ liệu workbook vẫn ở sheet SinhVien', doc.querySelector('.fh-data-grid-meta strong').textContent, 'SinhVien');
pick(doc.querySelector('.fh-ref-select'), 'SinhVien.MaSinhVien');

console.log('\n— Xem trước kết quả bằng JavaScript (§95 §96)');
// preview có debounce 300ms
var done = function () {
  var rows = doc.querySelectorAll('.fh-preview-table tbody tr');
  ok('bảng xem trước có dữ liệu', rows.length >= 3);
  var resultCells = doc.querySelectorAll('.fh-preview-result');
  ok('phát hiện được SV001 bị trùng',
    Array.prototype.some.call(resultCells, function (td) { return td.textContent === 'TRÙNG'; }));

  console.log('\n— Lỗi cấu hình hiển thị inline (§120 §173)');
  win.FormulaHelper.openTool('existence');
  ok('hiện lỗi thiếu cấu hình', !!doc.querySelector('.fh-inline-error'));
  ok('không dùng alert', doc.querySelector('.fh-inline-error p').textContent.length > 0);

  console.log('\n— Lịch sử lưu local (§206)');
  setTimeout(function () {
    var histKey = lsKey('genscript.formulaHistory');
    ok('có key lịch sử', !!histKey);
    var hist = JSON.parse(win.localStorage.getItem(histKey) || '[]');
    ok('lịch sử có bản ghi', hist.length >= 1);
    ok('bản ghi lưu cả cấu hình (§102)', !!(hist[0] && hist[0].config && hist[0].config.__toolId));

    console.log('\n— Help theo workspace (§116 §205)');
    doc.getElementById('btnHelp').dispatchEvent(new win.Event('click', { bubbles: true }));
    ok('mở đúng hướng dẫn Formula Helper', doc.getElementById('formulaHelpDialog').open === true);
    ok('không mở hướng dẫn Gen Script', !doc.getElementById('helpDialog').open);
    doc.getElementById('btnFormulaHelpDone').dispatchEvent(new win.Event('click', { bubbles: true }));
    ok('đóng được bằng nút Đóng', doc.getElementById('formulaHelpDialog').open === false);

    doc.querySelectorAll('.ws-btn')[0].dispatchEvent(new win.Event('click', { bubbles: true }));
    doc.getElementById('btnHelp').dispatchEvent(new win.Event('click', { bubbles: true }));
    ok('ở Gen Script mở hướng dẫn Gen Script', doc.getElementById('helpDialog').open === true);

    console.log('\n— Quay lại Gen Script, hai lưới vẫn tách riêng (§200)');
    ok('pane Gen Script hiện lại', !doc.getElementById('workspaceGenScript').hidden);
    is('Formula workbook vẫn đủ 3 sheet', win.Workbook.getSheets().length, 3);
    is('dữ liệu Formula SinhVien còn nguyên', win.Workbook.getSheets()[1].data.length, 5);
    ok('Gen Script vẫn không có tab SinhVien',
      Array.prototype.every.call(doc.querySelectorAll('.sheet-tabs .tab-name'), function (n) {
        return n.textContent !== 'SinhVien';
      }));

    console.log('\n' + (fail === 0 ? '✔ ' : '✘ ') + pass + ' đạt, ' + fail + ' hỏng.\n');
    process.exit(fail ? 1 : 0);
  }, 1800);
};
setTimeout(done, 500);

}
