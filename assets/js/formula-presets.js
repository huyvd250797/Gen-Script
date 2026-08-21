/*!
 * formula-presets.js — Danh mục công cụ Formula Helper.
 *
 * Mỗi công cụ mô tả bằng metadata + builderSchema để UI render tự động (§190 §191),
 * hạn chế hardcode giao diện cho từng công thức (§192).
 */
(function (global) {
  'use strict';

  var BOTH = ['excel', 'sheets'];

  var OPERATOR_OPTIONS = [
    { value: '=', label: 'bằng (=)' },
    { value: '<>', label: 'khác (<>)' },
    { value: '>', label: 'lớn hơn (>)' },
    { value: '>=', label: 'lớn hơn hoặc bằng (>=)' },
    { value: '<', label: 'nhỏ hơn (<)' },
    { value: '<=', label: 'nhỏ hơn hoặc bằng (<=)' },
    { value: 'contains', label: 'chứa' },
    { value: 'startsWith', label: 'bắt đầu bằng' },
    { value: 'endsWith', label: 'kết thúc bằng' },
    { value: 'blank', label: 'đang trống' },
    { value: 'notBlank', label: 'không trống' }
  ];

  /* ------------------------------------------------------------ icons */

  var ICONS = {
    lookup: '<circle cx="11" cy="11" r="6"></circle><path d="M20 20l-4.5-4.5"></path>',
    duplicate: '<rect x="8" y="8" width="12" height="12" rx="2"></rect><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2"></path>',
    check: '<path d="M20 6L9 17l-5-5"></path>',
    compare: '<path d="M4 7h7M4 17h7M17 4v16M14 8l3-3 3 3M14 16l3 3 3-3"></path>',
    filter: '<path d="M3 5h18l-7 8v6l-4 2v-8z"></path>',
    logic: '<circle cx="6" cy="6" r="2.5"></circle><circle cx="6" cy="18" r="2.5"></circle><circle cx="18" cy="12" r="2.5"></circle><path d="M8.5 6H12a3 3 0 0 1 3 3v0M8.5 18H12a3 3 0 0 0 3-3v0"></path>',
    text: '<path d="M4 6h16M9 6v14M15 6v6"></path>',
    date: '<rect x="3" y="5" width="18" height="16" rx="2"></rect><path d="M3 10h18M8 3v4M16 3v4"></path>',
    concat: '<path d="M9 7H6a5 5 0 0 0 0 10h3M15 7h3a5 5 0 0 1 0 10h-3M9 12h6"></path>',
    sql: '<ellipse cx="12" cy="6" rx="8" ry="3"></ellipse><path d="M4 6v12c0 1.7 3.6 3 8 3s8-1.3 8-3V6M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"></path>',
    sigma: '<path d="M18 5H7l6 7-6 7h11"></path>',
    unique: '<path d="M5 12h6M5 6h14M5 18h9"></path><circle cx="17" cy="12" r="2.5"></circle>'
  };

  /* ------------------------------------------------------------ công cụ */

  var TOOLS = [
    {
      id: 'lookup',
      type: 'lookup',
      category: 'LOOKUP',
      name: 'Tra cứu dữ liệu',
      desc: 'Tìm dữ liệu ở Sheet khác và trả về giá trị tương ứng.',
      icon: ICONS.lookup,
      keywords: ['tra cuu', 'lookup', 'xlookup', 'vlookup', 'index', 'match', 'do chieu', 'lay id', 'match id'],
      platforms: BOTH,
      primaryFunction: 'XLOOKUP',
      schema: [
        { key: 'lookupValue', type: 'field', label: 'Giá trị cần tìm', technical: 'lookup_value', useSelected: true, required: true },
        { key: 'lookupRange', type: 'field', label: 'Cột dùng để tìm', technical: 'lookup_array', required: true },
        { key: 'returnRange', type: 'field', label: 'Cột trả về', technical: 'return_array', required: true },
        { key: 'notFound', type: 'text', label: 'Nếu không tìm thấy', placeholder: 'để trống hoặc "Không tìm thấy"' },
        {
          key: 'method', type: 'select', label: 'Cách tra cứu', advanced: true,
          options: [
            { value: 'auto', label: 'Tự chọn (ưu tiên XLOOKUP)' },
            { value: 'xlookup', label: 'XLOOKUP' },
            { value: 'vlookup', label: 'VLOOKUP' },
            { value: 'index-match', label: 'INDEX + MATCH' }
          ]
        },
        { key: 'conditions', type: 'conditions', label: 'Điều kiện thêm (tra cứu nhiều điều kiện)', advanced: true, valueOnly: true }
      ],
      defaults: { method: 'auto', notFound: '' }
    },
    {
      id: 'duplicate',
      type: 'duplicate',
      category: 'CHECK',
      name: 'Kiểm tra trùng',
      desc: 'Phát hiện giá trị xuất hiện nhiều lần trong cùng một cột.',
      icon: ICONS.duplicate,
      keywords: ['trung', 'duplicate', 'countif', 'countifs', 'lap', 'mssv trung'],
      platforms: BOTH,
      primaryFunction: 'COUNTIF',
      schema: [
        { key: 'fields', type: 'fields', label: 'Cột kiểm tra', min: 1, useSelected: true, required: true },
        {
          key: 'resultType', type: 'select', label: 'Kết quả trả về',
          options: [
            { value: 'label', label: 'Nhãn (TRÙNG / rỗng)' },
            { value: 'count', label: 'Số lần xuất hiện' }
          ]
        },
        { key: 'duplicateLabel', type: 'text', label: 'Khi trùng ghi', showIf: { resultType: 'label' } },
        { key: 'uniqueLabel', type: 'text', label: 'Khi không trùng ghi', showIf: { resultType: 'label' } }
      ],
      defaults: { resultType: 'label', duplicateLabel: 'TRÙNG', uniqueLabel: '' }
    },
    {
      id: 'existence',
      type: 'existence',
      category: 'CHECK',
      name: 'Kiểm tra tồn tại',
      desc: 'Kiểm tra dữ liệu có nằm trong danh mục khác hay không.',
      icon: ICONS.check,
      keywords: ['ton tai', 'existence', 'co khong', 'countif', 'xmatch', 'danh muc'],
      platforms: BOTH,
      primaryFunction: 'COUNTIF',
      schema: [
        { key: 'value', type: 'field', label: 'Giá trị cần kiểm tra', useSelected: true, required: true },
        { key: 'reference', type: 'field', label: 'Đối chiếu với cột', required: true },
        { key: 'foundLabel', type: 'text', label: 'Nếu có ghi' },
        { key: 'missingLabel', type: 'text', label: 'Nếu không có ghi' },
        {
          key: 'method', type: 'select', label: 'Cách kiểm tra', advanced: true,
          options: [
            { value: 'countif', label: 'COUNTIF (mặc định)' },
            { value: 'xmatch', label: 'XMATCH' },
            { value: 'xlookup', label: 'XLOOKUP' }
          ]
        }
      ],
      defaults: { foundLabel: 'CÓ', missingLabel: 'KHÔNG', method: 'countif' }
    },
    {
      id: 'compare',
      type: 'compare',
      category: 'COMPARE',
      name: 'So sánh dữ liệu',
      desc: 'So sánh hai giá trị trên cùng một dòng, ví dụ OLD và NEW.',
      icon: ICONS.compare,
      keywords: ['so sanh', 'compare', 'giong khac', 'old new', 'doi chieu'],
      platforms: BOTH,
      primaryFunction: 'IF',
      schema: [
        { key: 'valueA', type: 'field', label: 'Giá trị A', useSelected: true, required: true },
        { key: 'valueB', type: 'field', label: 'Giá trị B', required: true },
        { key: 'trim', type: 'checkbox', label: 'Cắt khoảng trắng trước khi so sánh' },
        { key: 'ignoreCase', type: 'checkbox', label: 'Không phân biệt hoa thường' },
        { key: 'sameLabel', type: 'text', label: 'Nếu giống ghi' },
        { key: 'diffLabel', type: 'text', label: 'Nếu khác ghi' }
      ],
      defaults: { sameLabel: 'GIỐNG', diffLabel: 'KHÁC', trim: true }
    },
    {
      id: 'listCompare',
      type: 'listCompare',
      category: 'COMPARE',
      name: 'So sánh hai danh sách',
      desc: 'Đối chiếu danh sách A với danh sách B để tìm phần thiếu.',
      icon: ICONS.compare,
      keywords: ['so sanh danh sach', 'list', 'thieu', 'doi soat', 'countif'],
      platforms: BOTH,
      primaryFunction: 'COUNTIF',
      schema: [
        { key: 'value', type: 'field', label: 'Giá trị ở danh sách A', useSelected: true, required: true },
        { key: 'reference', type: 'field', label: 'Danh sách B', required: true },
        {
          key: 'resultMode', type: 'select', label: 'Muốn biết',
          options: [
            { value: 'both', label: 'Có ở B hay không' },
            { value: 'onlyA', label: 'Chỉ có ở A' }
          ]
        },
        { key: 'matchLabel', type: 'text', label: 'Nhãn khi khớp' },
        { key: 'missLabel', type: 'text', label: 'Nhãn khi thiếu', showIf: { resultMode: 'both' } }
      ],
      defaults: { resultMode: 'both', matchLabel: 'CÓ', missLabel: 'THIẾU' }
    },
    {
      id: 'filter',
      type: 'filter',
      category: 'FILTER',
      name: 'Lọc dữ liệu',
      desc: 'Trích các dòng thoả điều kiện sang một vùng khác.',
      icon: ICONS.filter,
      keywords: ['loc', 'filter', 'query', 'contains', 'trich'],
      platforms: BOTH,
      primaryFunction: 'FILTER',
      schema: [
        { key: 'sourceSheet', type: 'sheet', label: 'Sheet nguồn', required: true },
        { key: 'returnRange', type: 'text', label: 'Vùng trả về', placeholder: 'A:G hoặc A6:G' },
        {
          key: 'sheetsMethod', type: 'select', label: 'Hàm sử dụng', platform: 'sheets',
          options: [
            { value: 'filter', label: 'FILTER' },
            { value: 'query', label: 'QUERY' }
          ]
        },
        {
          key: 'logic', type: 'select', label: 'Nối các điều kiện bằng',
          options: [
            { value: 'AND', label: 'AND — thoả tất cả' },
            { value: 'OR', label: 'OR — thoả bất kỳ' }
          ]
        },
        { key: 'conditions', type: 'conditions', label: 'Điều kiện lọc', required: true },
        { key: 'notFound', type: 'text', label: 'Nếu không có dòng nào', advanced: true },
        { key: 'orderBy', type: 'text', label: 'Sắp xếp theo cột', advanced: true, platform: 'sheets', placeholder: 'A' },
        { key: 'limit', type: 'number', label: 'Giới hạn số dòng', advanced: true, platform: 'sheets' }
      ],
      defaults: { logic: 'AND', sheetsMethod: 'filter', returnRange: '' }
    },
    {
      id: 'logic',
      type: 'logic',
      category: 'LOGIC',
      name: 'Kiểm tra điều kiện',
      desc: 'Dựng IF, IFS, AND, OR bằng giao diện, không cần nhớ cú pháp lồng.',
      icon: ICONS.logic,
      keywords: ['dieu kien', 'if', 'ifs', 'and', 'or', 'xep loai', 'phan loai'],
      platforms: BOTH,
      primaryFunction: 'IF',
      schema: [
        { key: 'branches', type: 'branches', label: 'Các nhánh điều kiện', required: true },
        { key: 'elseResult', type: 'text', label: 'Còn lại thì ghi' }
      ],
      defaults: { elseResult: '' }
    },
    {
      id: 'blank',
      type: 'blank',
      category: 'CHECK',
      name: 'Kiểm tra dữ liệu rỗng',
      desc: 'Cảnh báo khi thiếu dữ liệu ở các cột bắt buộc.',
      icon: ICONS.check,
      keywords: ['rong', 'blank', 'thieu du lieu', 'bat buoc', 'trong'],
      platforms: BOTH,
      primaryFunction: 'IF',
      schema: [
        { key: 'fields', type: 'fields', label: 'Các cột bắt buộc', min: 1, useSelected: true, required: true },
        { key: 'missingLabel', type: 'text', label: 'Nếu thiếu ghi' },
        { key: 'okLabel', type: 'text', label: 'Nếu đủ ghi' }
      ],
      defaults: { missingLabel: 'THIẾU DỮ LIỆU', okLabel: '' }
    },
    {
      id: 'text',
      type: 'text',
      category: 'TEXT',
      name: 'Làm sạch Text',
      desc: 'Trim, Clean, đổi hoa thường, thay thế và cắt chuỗi.',
      icon: ICONS.text,
      keywords: ['text', 'trim', 'clean', 'chuoi', 'lam sach', 'hoa thuong', 'substitute', 'left', 'right', 'mid', 'len'],
      platforms: BOTH,
      primaryFunction: 'TRIM',
      schema: [
        { key: 'field', type: 'field', label: 'Cột cần xử lý', useSelected: true, required: true },
        {
          key: 'preset', type: 'select', label: 'Việc cần làm',
          options: [
            { value: 'trim', label: 'Cắt khoảng trắng thừa (TRIM)' },
            { value: 'clean', label: 'Xoá ký tự không in được (CLEAN)' },
            { value: 'trimClean', label: 'Trim + Clean' },
            { value: 'upper', label: 'Viết hoa toàn bộ' },
            { value: 'lower', label: 'Viết thường toàn bộ' },
            { value: 'proper', label: 'Viết hoa đầu mỗi từ' },
            { value: 'removeSpace', label: 'Xoá toàn bộ khoảng trắng' },
            { value: 'removeLineBreak', label: 'Xoá ký tự xuống dòng' },
            { value: 'substitute', label: 'Thay thế ký tự' },
            { value: 'left', label: 'Lấy N ký tự đầu' },
            { value: 'right', label: 'Lấy N ký tự cuối' },
            { value: 'mid', label: 'Lấy ký tự từ vị trí' },
            { value: 'len', label: 'Đếm số ký tự' }
          ]
        },
        { key: 'find', type: 'text', label: 'Tìm ký tự', showIf: { preset: 'substitute' } },
        { key: 'replace', type: 'text', label: 'Thay bằng', showIf: { preset: 'substitute' } },
        { key: 'start', type: 'number', label: 'Bắt đầu từ vị trí', showIf: { preset: 'mid' } },
        { key: 'count', type: 'number', label: 'Số ký tự', showIf: { preset: ['left', 'right', 'mid'] } }
      ],
      defaults: { preset: 'trim', start: 1, count: 5 }
    },
    {
      id: 'date',
      type: 'date',
      category: 'DATE',
      name: 'Xử lý ngày',
      desc: 'Định dạng ngày, tính tuổi, chênh lệch và năm học.',
      icon: ICONS.date,
      keywords: ['ngay', 'date', 'tuoi', 'nam hoc', 'thang', 'nam', 'datedif', 'text'],
      platforms: BOTH,
      primaryFunction: 'TEXT',
      schema: [
        { key: 'field', type: 'field', label: 'Cột ngày', useSelected: true, required: true },
        {
          key: 'preset', type: 'select', label: 'Việc cần làm',
          options: [
            { value: 'format', label: 'Định dạng ngày' },
            { value: 'age', label: 'Tính tuổi' },
            { value: 'daysBetween', label: 'Số ngày giữa hai mốc' },
            { value: 'monthsBetween', label: 'Số tháng giữa hai mốc' },
            { value: 'yearsBetween', label: 'Số năm giữa hai mốc' },
            { value: 'year', label: 'Lấy năm' },
            { value: 'month', label: 'Lấy tháng' },
            { value: 'day', label: 'Lấy ngày' },
            { value: 'academicYear', label: 'Tạo năm học' }
          ]
        },
        { key: 'format', type: 'text', label: 'Định dạng', showIf: { preset: 'format' }, placeholder: 'dd/mm/yyyy' },
        { key: 'field2', type: 'field', label: 'Cột ngày thứ hai', showIf: { preset: ['daysBetween', 'monthsBetween', 'yearsBetween'] } },
        { key: 'startMonth', type: 'number', label: 'Tháng bắt đầu năm học', showIf: { preset: 'academicYear' } }
      ],
      defaults: { preset: 'format', format: 'dd/mm/yyyy', startMonth: 8 }
    },
    {
      id: 'concat',
      type: 'concat',
      category: 'TEXT',
      name: 'Ghép dữ liệu',
      desc: 'Nối nhiều cột thành một chuỗi, bỏ qua ô rỗng.',
      icon: ICONS.concat,
      keywords: ['ghep', 'concat', 'textjoin', 'noi chuoi', 'ho ten'],
      platforms: BOTH,
      primaryFunction: 'TEXTJOIN',
      schema: [
        { key: 'fields', type: 'fields', label: 'Các cột cần ghép', min: 2, useSelected: true, required: true },
        {
          key: 'joinWith', type: 'select', label: 'Ngăn cách bằng',
          options: [
            { value: 'space', label: 'Khoảng trắng' },
            { value: 'dash', label: 'Dấu gạch ngang' },
            { value: 'slash', label: 'Dấu gạch chéo' },
            { value: 'underscore', label: 'Gạch dưới' },
            { value: 'comma', label: 'Dấu phẩy' },
            { value: 'none', label: 'Không ngăn cách' }
          ]
        },
        { key: 'separatorText', type: 'text', label: 'Hoặc ký tự tuỳ ý', advanced: true },
        { key: 'ignoreBlank', type: 'checkbox', label: 'Bỏ qua ô rỗng' },
        {
          key: 'method', type: 'select', label: 'Hàm sử dụng', advanced: true,
          options: [
            { value: 'textjoin', label: 'TEXTJOIN' },
            { value: 'ampersand', label: 'Toán tử &' }
          ]
        }
      ],
      defaults: { joinWith: 'space', ignoreBlank: true, method: 'textjoin' }
    },
    {
      id: 'aggregate',
      type: 'aggregate',
      category: 'ADVANCED',
      name: 'Đếm & Tính tổng',
      desc: 'COUNTIF, COUNTIFS, SUMIF, SUMIFS theo điều kiện.',
      icon: ICONS.sigma,
      keywords: ['dem', 'tong', 'count', 'sum', 'countif', 'sumifs', 'thong ke'],
      platforms: BOTH,
      primaryFunction: 'SUMIFS',
      schema: [
        {
          key: 'kind', type: 'select', label: 'Việc cần làm',
          options: [
            { value: 'sum', label: 'Tính tổng' },
            { value: 'count', label: 'Đếm số dòng' }
          ]
        },
        { key: 'sumField', type: 'field', label: 'Cột cần tính tổng', showIf: { kind: 'sum' }, required: true },
        { key: 'conditions', type: 'conditions', label: 'Điều kiện' }
      ],
      defaults: { kind: 'sum' }
    },
    {
      id: 'unique',
      type: 'unique',
      category: 'ADVANCED',
      name: 'Lấy danh sách duy nhất',
      desc: 'Loại bỏ giá trị lặp để có danh mục sạch.',
      icon: ICONS.unique,
      keywords: ['duy nhat', 'unique', 'distinct', 'sort', 'danh muc'],
      platforms: BOTH,
      primaryFunction: 'UNIQUE',
      schema: [
        { key: 'field', type: 'field', label: 'Cột nguồn', useSelected: true, required: true },
        { key: 'sort', type: 'checkbox', label: 'Sắp xếp kết quả' }
      ],
      defaults: {}
    },
    {
      id: 'sqlQuote',
      type: 'sqlQuote',
      category: 'SQL',
      name: 'Tạo giá trị SQL',
      desc: "Bọc giá trị thành 'SV001' để dán vào câu lệnh SQL.",
      icon: ICONS.sql,
      keywords: ['sql', 'quote', 'nhay don', 'gia tri'],
      platforms: BOTH,
      primaryFunction: 'CONCAT',
      schema: [
        { key: 'field', type: 'field', label: 'Cột nguồn', useSelected: true, required: true },
        {
          key: 'dataType', type: 'select', label: 'Kiểu dữ liệu',
          options: [
            { value: 'text', label: 'Chuỗi — có nháy đơn' },
            { value: 'number', label: 'Số — không nháy' }
          ]
        }
      ],
      defaults: { dataType: 'text' }
    },
    {
      id: 'sqlIn',
      type: 'sqlIn',
      category: 'SQL',
      name: 'Tạo danh sách SQL IN',
      desc: "Gom cả cột thành 'SV001','SV002' cho mệnh đề IN.",
      icon: ICONS.sql,
      keywords: ['sql in', 'where in', 'textjoin', 'danh sach', 'gom'],
      platforms: BOTH,
      primaryFunction: 'TEXTJOIN',
      supportsDirectOutput: true,
      schema: [
        { key: 'field', type: 'field', label: 'Cột nguồn', useSelected: true, required: true },
        {
          key: 'dataType', type: 'select', label: 'Kiểu dữ liệu',
          options: [
            { value: 'text', label: 'Chuỗi — có nháy đơn' },
            { value: 'number', label: 'Số — không nháy' }
          ]
        },
        { key: 'sqlColumn', type: 'text', label: 'Tên cột trong SQL', placeholder: 'MaSinhVien', advanced: true }
      ],
      defaults: { dataType: 'text' }
    }
  ];

  /* ------------------------------------------------------------ nhóm */

  var CATEGORIES = {
    LOOKUP: 'Tra cứu',
    CHECK: 'Kiểm tra',
    COMPARE: 'So sánh',
    FILTER: 'Lọc dữ liệu',
    LOGIC: 'Điều kiện',
    TEXT: 'Xử lý Text',
    DATE: 'Xử lý ngày',
    SQL: 'SQL Tools',
    ADVANCED: 'Nâng cao'
  };

  /* ------------------------------------------------------------ thư viện hàm (§187 §188) */

  var LIBRARY = [
    { name: 'XLOOKUP', use: 'Tìm giá trị trong một vùng và trả về giá trị tương ứng từ vùng khác.', example: '=XLOOKUP(A2,Database!B:B,Database!A:A,"")', platforms: BOTH },
    { name: 'VLOOKUP', use: 'Tra cứu theo cột đầu tiên của bảng, trả về cột thứ N bên phải.', example: '=VLOOKUP(A2,Data!A:D,4,FALSE)', platforms: BOTH },
    { name: 'INDEX', use: 'Lấy giá trị tại một vị trí trong vùng.', example: '=INDEX(B:B,MATCH(A2,A:A,0))', platforms: BOTH },
    { name: 'MATCH', use: 'Trả về vị trí của giá trị trong một vùng.', example: '=MATCH(A2,A:A,0)', platforms: BOTH },
    { name: 'COUNTIF', use: 'Đếm số ô thoả một điều kiện.', example: '=COUNTIF($A:$A,A2)', platforms: BOTH },
    { name: 'COUNTIFS', use: 'Đếm số ô thoả nhiều điều kiện cùng lúc.', example: '=COUNTIFS($A:$A,A2,$B:$B,B2)', platforms: BOTH },
    { name: 'SUMIF', use: 'Cộng các giá trị thoả một điều kiện.', example: '=SUMIF(A:A,A2,C:C)', platforms: BOTH },
    { name: 'SUMIFS', use: 'Cộng các giá trị thoả nhiều điều kiện.', example: '=SUMIFS(D:D,A:A,A2,B:B,B2)', platforms: BOTH },
    { name: 'IF', use: 'Trả về giá trị khác nhau tuỳ điều kiện đúng hay sai.', example: '=IF(A2>=5,"Đạt","Không đạt")', platforms: BOTH },
    { name: 'IFS', use: 'Xét nhiều điều kiện theo thứ tự, thay cho IF lồng nhau.', example: '=IFS(A2>=8.5,"Giỏi",A2>=7,"Khá",TRUE,"Đạt")', platforms: BOTH },
    { name: 'AND', use: 'Đúng khi tất cả điều kiện đều đúng.', example: '=AND(A2<>"",B2="Đang học")', platforms: BOTH },
    { name: 'OR', use: 'Đúng khi có ít nhất một điều kiện đúng.', example: '=OR(A2="CNTT",A2="KTPM")', platforms: BOTH },
    { name: 'FILTER', use: 'Trả về các dòng thoả điều kiện.', example: '=FILTER(A:G,(F:F="CNTT"),"")', platforms: BOTH },
    { name: 'UNIQUE', use: 'Loại bỏ giá trị trùng lặp trong một vùng.', example: '=UNIQUE(A2:A10000)', platforms: BOTH },
    { name: 'SORT', use: 'Sắp xếp một vùng dữ liệu.', example: '=SORT(UNIQUE(A2:A10000))', platforms: BOTH },
    { name: 'TEXTJOIN', use: 'Nối nhiều giá trị bằng một ký tự ngăn cách.', example: '=TEXTJOIN(" ",TRUE,A2,B2)', platforms: BOTH },
    { name: 'TRIM', use: 'Cắt khoảng trắng đầu, cuối và gộp khoảng trắng kép.', example: '=TRIM(A2)', platforms: BOTH },
    { name: 'CLEAN', use: 'Loại bỏ ký tự điều khiển không in được.', example: '=CLEAN(A2)', platforms: BOTH },
    { name: 'SUBSTITUTE', use: 'Thay thế một chuỗi con bằng chuỗi khác.', example: '=SUBSTITUTE(A2,"-","/")', platforms: BOTH },
    { name: 'LEFT', use: 'Lấy số ký tự tính từ bên trái.', example: '=LEFT(A2,3)', platforms: BOTH },
    { name: 'RIGHT', use: 'Lấy số ký tự tính từ bên phải.', example: '=RIGHT(A2,5)', platforms: BOTH },
    { name: 'MID', use: 'Lấy ký tự bắt đầu từ một vị trí.', example: '=MID(A2,4,4)', platforms: BOTH },
    { name: 'LEN', use: 'Đếm số ký tự của giá trị.', example: '=LEN(A2)', platforms: BOTH },
    { name: 'SEARCH', use: 'Tìm vị trí chuỗi con, không phân biệt hoa thường.', example: '=SEARCH("CNTT",A2)', platforms: BOTH },
    { name: 'FIND', use: 'Tìm vị trí chuỗi con, có phân biệt hoa thường.', example: '=FIND("CNTT",A2)', platforms: BOTH },
    { name: 'TEXT', use: 'Định dạng số hoặc ngày thành chuỗi.', example: '=TEXT(A2,"dd/mm/yyyy")', platforms: BOTH },
    { name: 'DATEDIF', use: 'Tính chênh lệch giữa hai mốc thời gian.', example: '=DATEDIF(A2,TODAY(),"Y")', platforms: BOTH },
    { name: 'YEAR', use: 'Lấy phần năm của một ngày.', example: '=YEAR(A2)', platforms: BOTH },
    { name: 'MONTH', use: 'Lấy phần tháng của một ngày.', example: '=MONTH(A2)', platforms: BOTH },
    { name: 'DAY', use: 'Lấy phần ngày của một ngày.', example: '=DAY(A2)', platforms: BOTH },
    { name: 'QUERY', use: 'Truy vấn dữ liệu bằng cú pháp giống SQL.', example: '=QUERY(A6:G,"select * where F contains \'F3\'",0)', platforms: ['sheets'] },
    { name: 'ARRAYFORMULA', use: 'Áp dụng công thức cho cả một vùng cùng lúc.', example: '=ARRAYFORMULA(TRIM(A2:A))', platforms: ['sheets'] },
    { name: 'REGEXMATCH', use: 'Kiểm tra chuỗi có khớp biểu thức chính quy không.', example: '=REGEXMATCH(A2,"^SV\\d+$")', platforms: ['sheets'] },
    { name: 'REGEXEXTRACT', use: 'Trích phần chuỗi khớp biểu thức chính quy.', example: '=REGEXEXTRACT(A2,"\\d+")', platforms: ['sheets'] }
  ];

  function getTool(id) {
    for (var i = 0; i < TOOLS.length; i++) if (TOOLS[i].id === id) return TOOLS[i];
    return null;
  }

  /** Bỏ dấu tiếng Việt để tìm kiếm không dấu (§109 §186). */
  function deaccent(s) {
    return String(s || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd');
  }

  function search(query) {
    var q = deaccent(query).trim();
    if (!q) return TOOLS.slice();
    return TOOLS.filter(function (t) {
      var hay = deaccent(t.name + ' ' + t.desc + ' ' + t.category + ' ' + t.primaryFunction + ' ' + t.keywords.join(' '));
      return hay.indexOf(q) !== -1;
    });
  }

  function searchLibrary(query) {
    var q = deaccent(query).trim();
    if (!q) return LIBRARY.slice();
    return LIBRARY.filter(function (f) {
      return deaccent(f.name + ' ' + f.use).indexOf(q) !== -1;
    });
  }

  global.FormulaPresets = {
    TOOLS: TOOLS,
    CATEGORIES: CATEGORIES,
    LIBRARY: LIBRARY,
    OPERATOR_OPTIONS: OPERATOR_OPTIONS,
    getTool: getTool,
    search: search,
    searchLibrary: searchLibrary,
    deaccent: deaccent
  };
})(typeof window !== 'undefined' ? window : this);
