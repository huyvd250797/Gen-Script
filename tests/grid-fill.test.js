/*!
 * grid-fill.test.js — Kiểm thử fill handle V2.8.2 mà không cần DOM/jsdom.
 * Chạy: node tests/grid-fill.test.js
 */
'use strict';

var fs = require('fs');
var path = require('path');
var vm = require('vm');

var sandbox = { window: {}, console: console };
sandbox.window.window = sandbox.window;
vm.createContext(sandbox);
vm.runInContext(
  fs.readFileSync(path.join(__dirname, '..', 'assets', 'js', 'grid.js'), 'utf8'),
  sandbox
);
var Grid = sandbox.window.Grid;
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

function makeGrid(data, selection) {
  var g = Object.create(Grid.prototype);
  g.data = data.map(function (r) { return r.slice(); });
  g.selection = Object.assign({}, selection);
  g.active = { r: selection.r2, c: selection.c2 };
  g.history = [];
  g.future = [];
  g._lastRecordKey = '';
  g._lastRecordAt = 0;
  g.onHistoryChange = function () {};
  g.onStatus = function () {};
  g.onChange = function () {};
  g.render = function () {};
  g._endEdit = function () {};
  g._ensureSize = function (rows, cols) {
    var width = g.data[0].length;
    while (width < cols) {
      g.data.forEach(function (row) { row.push(''); });
      width++;
    }
    while (g.data.length < rows) {
      g.data.push(Array(width).fill(''));
    }
    return true;
  };
  return g;
}

console.log('\n— Fill một ô số tăng +1');
var g1 = makeGrid([
  ['Id', 'Ten'],
  ['137', 'A'],
  ['', 'B'],
  ['', 'C'],
  ['', 'D']
], { r1: 1, c1: 0, r2: 1, c2: 0 });
g1._applyDragFill(4, 0);
is('137 kéo xuống thành 137,138,139,140', g1.data.slice(1, 5).map(function (r) { return r[0]; }).join(','), '137,138,139,140');

console.log('\n— Fill chuỗi hai số giữ bước chênh lệch');
var g2 = makeGrid([
  ['Id'], ['10'], ['20'], [''], ['']
], { r1: 1, c1: 0, r2: 2, c2: 0 });
g2._applyDragFill(4, 0);
is('10,20 tiếp tục 30,40', g2.data.slice(1, 5).map(function (r) { return r[0]; }).join(','), '10,20,30,40');

console.log('\n— Double-click xác định dòng cuối theo dữ liệu bên cạnh');
var g3 = makeGrid([
  ['Id', 'Ten', 'GhiChu'],
  ['137', 'A', ''],
  ['', 'B', ''],
  ['', 'C', ''],
  ['', 'D', ''],
  ['', '', '']
], { r1: 1, c1: 0, r2: 1, c2: 0 });
is('dòng cuối vùng liền kề là 4', g3._autoFillEndRow(), 4);
g3._autoFillDown();
is('double-click tự điền đến dòng cuối', g3.data.slice(1, 5).map(function (r) { return r[0]; }).join(','), '137,138,139,140');

console.log('\n— Fallback khi cột sát cạnh trống');
var g4 = makeGrid([
  ['Id', 'Phu', 'Ten'],
  ['5', '', 'A'],
  ['', '', 'B'],
  ['', '', 'C'],
  ['', '', '']
], { r1: 1, c1: 0, r2: 1, c2: 0 });
is('vẫn tìm được dòng cuối ở cột xa hơn', g4._autoFillEndRow(), 3);

console.log('\nKết quả: ' + pass + ' passed, ' + fail + ' failed.');
if (fail) process.exit(1);
