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

  function demo(input, formula, result) {
    return { input: input, formula: formula, result: result };
  }

  var LIBRARY = [
    { name: 'XLOOKUP', use: 'Tìm giá trị trong một vùng và trả về giá trị tương ứng từ vùng khác.', example: '=XLOOKUP(A2,Database!B:B,Database!A:A,"")', platforms: BOTH, examples: [
      demo('A2 = SV001, Database cột B có SV001, cột A là 125', '=XLOOKUP(A2,Database!B:B,Database!A:A,"")', '125'),
      demo('A2 = SV999 không có trong Database', '=XLOOKUP(A2,Database!B:B,Database!A:A,"Không có")', 'Không có')
    ] },
    { name: 'VLOOKUP', use: 'Tra cứu theo cột đầu tiên của bảng, trả về cột thứ N bên phải.', example: '=VLOOKUP(A2,Data!A:D,4,FALSE)', platforms: BOTH, examples: [
      demo('A2 = M01, Data cột A có M01, cột D là Đang học', '=VLOOKUP(A2,Data!A:D,4,FALSE)', 'Đang học'),
      demo('A2 = L01, Data cột A có L01, cột B là CNTT', '=VLOOKUP(A2,Data!A:D,2,FALSE)', 'CNTT')
    ] },
    { name: 'INDEX', use: 'Lấy giá trị tại một vị trí trong vùng.', example: '=INDEX(B:B,MATCH(A2,A:A,0))', platforms: BOTH, examples: [
      demo('MATCH tìm A2 ở dòng 5, cột B dòng 5 là Nguyễn An', '=INDEX(B:B,MATCH(A2,A:A,0))', 'Nguyễn An'),
      demo('Lấy phần tử thứ 3 trong vùng B2:B10', '=INDEX(B2:B10,3)', 'Giá trị ở ô B4')
    ] },
    { name: 'MATCH', use: 'Trả về vị trí của giá trị trong một vùng.', example: '=MATCH(A2,A:A,0)', platforms: BOTH, examples: [
      demo('A2 = SV003, trong A:A đứng ở dòng thứ 4 của vùng tìm', '=MATCH(A2,A:A,0)', '4'),
      demo('Tìm chữ CNTT trong C2:C20, khớp ở dòng thứ 2 của vùng', '=MATCH("CNTT",C2:C20,0)', '2')
    ] },
    { name: 'COUNTIF', use: 'Đếm số ô thoả một điều kiện.', example: '=COUNTIF($A:$A,A2)', platforms: BOTH, demoSheet: {
      formula: '=COUNTIF($B$2:$B$11,C2)',
      guide: 'COUNTIF sẽ quét vùng dữ liệu ban đầu rồi đếm xem giá trị cần đếm xuất hiện bao nhiêu lần.',
      letters: ['B', 'C', 'D', 'E'],
      rows: [
        ['Dữ liệu ban đầu', 'Dữ liệu cần đếm', 'Kết quả đếm', 'Công thức'],
        ['1', '1', '2', '=COUNTIF($B$2:$B$11,C2)'],
        ['1', '2', '2', '=COUNTIF($B$2:$B$11,C3)'],
        ['3', '3', '3', '=COUNTIF($B$2:$B$11,C4)'],
        ['3', '4', '2', '=COUNTIF($B$2:$B$11,C5)'],
        ['2', '5', '1', '=COUNTIF($B$2:$B$11,C6)'],
        ['3', '', '', ''],
        ['4', '', '', ''],
        ['2', '', '', ''],
        ['4', '', '', ''],
        ['5', '', '', '']
      ],
      explanation: [
        '$B$2:$B$11 là vùng dữ liệu ban đầu cần quét. Dấu $ giúp cố định vùng này khi kéo công thức xuống.',
        'C2 là giá trị cần đếm ở dòng đầu tiên. Trong ví dụ này C2 = 1.',
        'D2 trả về 2 vì số 1 xuất hiện 2 lần trong vùng $B$2:$B$11.',
        'Khi kéo công thức xuống, vùng $B$2:$B$11 giữ nguyên, còn C2 đổi thành C3, C4, C5... để đếm số 2, số 3, số 4, số 5.'
      ]
    }, examples: [
      demo('Cột A có SV001 xuất hiện 2 lần, A2 = SV001', '=COUNTIF($A:$A,A2)', '2'),
      demo('Cột C có 5 dòng là CNTT', '=COUNTIF(C:C,"CNTT")', '5')
    ] },
    { name: 'COUNTIFS', use: 'Đếm số ô thoả nhiều điều kiện cùng lúc.', example: '=COUNTIFS($A:$A,A2,$B:$B,B2)', platforms: BOTH, examples: [
      demo('Cột A = SV001 và cột B = HK1 khớp 1 dòng', '=COUNTIFS(A:A,"SV001",B:B,"HK1")', '1'),
      demo('Lớp CNTT và trạng thái Đang học có 32 dòng', '=COUNTIFS(C:C,"CNTT",D:D,"Đang học")', '32')
    ] },
    { name: 'SUMIF', use: 'Cộng các giá trị thoả một điều kiện.', example: '=SUMIF(A:A,A2,C:C)', platforms: BOTH, examples: [
      demo('Cộng học phí cột C của các dòng có lớp CNTT ở cột A', '=SUMIF(A:A,"CNTT",C:C)', 'Tổng học phí lớp CNTT'),
      demo('Cộng số lượng cột D của mã SP01 trong cột A', '=SUMIF(A:A,"SP01",D:D)', 'Tổng số lượng SP01')
    ] },
    { name: 'SUMIFS', use: 'Cộng các giá trị thoả nhiều điều kiện.', example: '=SUMIFS(D:D,A:A,A2,B:B,B2)', platforms: BOTH, examples: [
      demo('Cộng tiền cột D khi lớp = CNTT và học kỳ = HK1', '=SUMIFS(D:D,A:A,"CNTT",B:B,"HK1")', 'Tổng tiền CNTT HK1'),
      demo('Cộng số tín chỉ đã đạt của SV001 trong năm 2026', '=SUMIFS(E:E,A:A,"SV001",C:C,2026)', 'Tổng tín chỉ')
    ] },
    { name: 'IF', use: 'Trả về giá trị khác nhau tuỳ điều kiện đúng hay sai.', example: '=IF(A2>=5,"Đạt","Không đạt")', platforms: BOTH, examples: [
      demo('A2 = 7', '=IF(A2>=5,"Đạt","Không đạt")', 'Đạt'),
      demo('A2 đang trống', '=IF(A2="","Thiếu dữ liệu","OK")', 'Thiếu dữ liệu')
    ] },
    { name: 'IFS', use: 'Xét nhiều điều kiện theo thứ tự, thay cho IF lồng nhau.', example: '=IFS(A2>=8.5,"Giỏi",A2>=7,"Khá",TRUE,"Đạt")', platforms: BOTH, examples: [
      demo('A2 = 8', '=IFS(A2>=8.5,"Giỏi",A2>=7,"Khá",TRUE,"Đạt")', 'Khá'),
      demo('A2 = 4', '=IFS(A2>=8.5,"Giỏi",A2>=7,"Khá",TRUE,"Chưa đạt")', 'Chưa đạt')
    ] },
    { name: 'AND', use: 'Đúng khi tất cả điều kiện đều đúng.', example: '=AND(A2<>"",B2="Đang học")', platforms: BOTH, examples: [
      demo('A2 có mã SV, B2 = Đang học', '=AND(A2<>"",B2="Đang học")', 'TRUE'),
      demo('Điểm A2 = 6, chuyên cần B2 = 9', '=AND(A2>=5,B2>=8)', 'TRUE')
    ] },
    { name: 'OR', use: 'Đúng khi có ít nhất một điều kiện đúng.', example: '=OR(A2="CNTT",A2="KTPM")', platforms: BOTH, examples: [
      demo('A2 = CNTT', '=OR(A2="CNTT",A2="KTPM")', 'TRUE'),
      demo('Trạng thái là Đang học hoặc Bảo lưu', '=OR(B2="Đang học",B2="Bảo lưu")', 'TRUE nếu khớp một trong hai')
    ] },
    { name: 'FILTER', use: 'Trả về các dòng thoả điều kiện.', example: '=FILTER(A:G,(F:F="CNTT"),"")', platforms: BOTH, examples: [
      demo('Bảng A:G, cột F là ngành, cần lấy ngành CNTT', '=FILTER(A:G,F:F="CNTT","")', 'Các dòng ngành CNTT'),
      demo('Lọc sinh viên Đang học ở cột D', '=FILTER(A:D,D:D="Đang học","Không có")', 'Danh sách đang học')
    ] },
    { name: 'UNIQUE', use: 'Loại bỏ giá trị trùng lặp trong một vùng.', example: '=UNIQUE(A2:A10000)', platforms: BOTH, examples: [
      demo('A2:A6 = CNTT, CNTT, KTPM, QTKD', '=UNIQUE(A2:A6)', 'CNTT, KTPM, QTKD'),
      demo('Danh sách mã lớp có nhiều dòng trùng', '=UNIQUE(C2:C1000)', 'Mỗi mã lớp chỉ còn 1 lần')
    ] },
    { name: 'SORT', use: 'Sắp xếp một vùng dữ liệu.', example: '=SORT(UNIQUE(A2:A10000))', platforms: BOTH, examples: [
      demo('Danh sách lớp chưa theo thứ tự', '=SORT(UNIQUE(A2:A10000))', 'Danh sách lớp không trùng và đã sắp xếp'),
      demo('Sắp xếp bảng A2:D100 theo cột 2 tăng dần', '=SORT(A2:D100,2,TRUE)', 'Bảng đã xếp theo cột B')
    ] },
    { name: 'TEXTJOIN', use: 'Nối nhiều giá trị bằng một ký tự ngăn cách.', example: '=TEXTJOIN(" ",TRUE,A2,B2)', platforms: BOTH, examples: [
      demo('A2 = Nguyễn Văn, B2 = An', '=TEXTJOIN(" ",TRUE,A2,B2)', 'Nguyễn Văn An'),
      demo('A2:A4 = SV001, SV002, SV003', '=TEXTJOIN(",",TRUE,A2:A4)', 'SV001,SV002,SV003')
    ] },
    { name: 'TRIM', use: 'Cắt khoảng trắng đầu, cuối và gộp khoảng trắng kép.', example: '=TRIM(A2)', platforms: BOTH, examples: [
      demo('A2 = "  Nguyễn   Văn   An  "', '=TRIM(A2)', 'Nguyễn Văn An'),
      demo('A2 có dư khoảng trắng cuối mã', '=TRIM(A2)', 'Mã sạch để đối chiếu')
    ] },
    { name: 'CLEAN', use: 'Loại bỏ ký tự điều khiển không in được.', example: '=CLEAN(A2)', platforms: BOTH, examples: [
      demo('A2 có ký tự xuống dòng/ẩn từ copy web', '=CLEAN(A2)', 'Chuỗi sạch hơn'),
      demo('Mã sinh viên bị dính ký tự không nhìn thấy', '=CLEAN(A2)', 'Mã có thể so sánh lại')
    ] },
    { name: 'SUBSTITUTE', use: 'Thay thế một chuỗi con bằng chuỗi khác.', example: '=SUBSTITUTE(A2,"-","/")', platforms: BOTH, examples: [
      demo('A2 = 01-09-2026', '=SUBSTITUTE(A2,"-","/")', '01/09/2026'),
      demo('A2 = SV.001.2026', '=SUBSTITUTE(A2,".","")', 'SV0012026')
    ] },
    { name: 'LEFT', use: 'Lấy số ký tự tính từ bên trái.', example: '=LEFT(A2,3)', platforms: BOTH, examples: [
      demo('A2 = CNTT-K49', '=LEFT(A2,4)', 'CNTT'),
      demo('A2 = SV001', '=LEFT(A2,2)', 'SV')
    ] },
    { name: 'RIGHT', use: 'Lấy số ký tự tính từ bên phải.', example: '=RIGHT(A2,5)', platforms: BOTH, examples: [
      demo('A2 = CNTT-K49', '=RIGHT(A2,3)', 'K49'),
      demo('A2 = SV000125', '=RIGHT(A2,3)', '125')
    ] },
    { name: 'MID', use: 'Lấy ký tự bắt đầu từ một vị trí.', example: '=MID(A2,4,4)', platforms: BOTH, examples: [
      demo('A2 = SV-2026-001', '=MID(A2,4,4)', '2026'),
      demo('A2 = ABCDEF', '=MID(A2,2,3)', 'BCD')
    ] },
    { name: 'LEN', use: 'Đếm số ký tự của giá trị.', example: '=LEN(A2)', platforms: BOTH, examples: [
      demo('A2 = SV001', '=LEN(A2)', '5'),
      demo('A2 = 001234567890', '=LEN(A2)', '12')
    ] },
    { name: 'SEARCH', use: 'Tìm vị trí chuỗi con, không phân biệt hoa thường.', example: '=SEARCH("CNTT",A2)', platforms: BOTH, examples: [
      demo('A2 = Khoa CNTT, tìm CNTT', '=SEARCH("CNTT",A2)', '6'),
      demo('A2 = khoa cntt, tìm CNTT', '=SEARCH("CNTT",A2)', '6')
    ] },
    { name: 'FIND', use: 'Tìm vị trí chuỗi con, có phân biệt hoa thường.', example: '=FIND("CNTT",A2)', platforms: BOTH, examples: [
      demo('A2 = Khoa CNTT, tìm CNTT', '=FIND("CNTT",A2)', '6'),
      demo('A2 = khoa cntt, tìm CNTT', '=FIND("CNTT",A2)', '#VALUE! vì khác hoa thường')
    ] },
    { name: 'TEXT', use: 'Định dạng số hoặc ngày thành chuỗi.', example: '=TEXT(A2,"dd/mm/yyyy")', platforms: BOTH, examples: [
      demo('A2 là ngày 2026-09-01', '=TEXT(A2,"dd/mm/yyyy")', '01/09/2026'),
      demo('A2 = 1250000', '=TEXT(A2,"#,##0")', '1,250,000')
    ] },
    { name: 'DATEDIF', use: 'Tính chênh lệch giữa hai mốc thời gian.', example: '=DATEDIF(A2,TODAY(),"Y")', platforms: BOTH, examples: [
      demo('A2 = 01/01/2000, hôm nay sau sinh nhật năm 2026', '=DATEDIF(A2,TODAY(),"Y")', '26'),
      demo('Tính số tháng giữa A2 và B2', '=DATEDIF(A2,B2,"M")', 'Số tháng chênh lệch')
    ] },
    { name: 'YEAR', use: 'Lấy phần năm của một ngày.', example: '=YEAR(A2)', platforms: BOTH, examples: [
      demo('A2 = 15/08/2026', '=YEAR(A2)', '2026'),
      demo('A2 là ngày sinh 20/11/2003', '=YEAR(A2)', '2003')
    ] },
    { name: 'MONTH', use: 'Lấy phần tháng của một ngày.', example: '=MONTH(A2)', platforms: BOTH, examples: [
      demo('A2 = 15/08/2026', '=MONTH(A2)', '8'),
      demo('A2 = 01/12/2026', '=MONTH(A2)', '12')
    ] },
    { name: 'DAY', use: 'Lấy phần ngày của một ngày.', example: '=DAY(A2)', platforms: BOTH, examples: [
      demo('A2 = 15/08/2026', '=DAY(A2)', '15'),
      demo('A2 = 01/12/2026', '=DAY(A2)', '1')
    ] },
    { name: 'QUERY', use: 'Truy vấn dữ liệu bằng cú pháp giống SQL.', example: '=QUERY(A6:G,"select * where F contains \'F3\'",0)', platforms: ['sheets'], examples: [
      demo('Bảng A6:G, cột F có ngành CNTT', '=QUERY(A6:G,"select * where F contains \'CNTT\'",0)', 'Các dòng có F chứa CNTT'),
      demo('Lấy cột A,B nơi cột C = Đang học', '=QUERY(A:C,"select A,B where C = \'Đang học\'",1)', 'Danh sách A,B đang học')
    ] },
    { name: 'ARRAYFORMULA', use: 'Áp dụng công thức cho cả một vùng cùng lúc.', example: '=ARRAYFORMULA(TRIM(A2:A))', platforms: ['sheets'], examples: [
      demo('A2:A có nhiều tên bị dư khoảng trắng', '=ARRAYFORMULA(TRIM(A2:A))', 'Cả cột tên được làm sạch'),
      demo('A2:A là họ, B2:B là tên', '=ARRAYFORMULA(A2:A&" "&B2:B)', 'Ghép họ tên cho cả cột')
    ] },
    { name: 'REGEXMATCH', use: 'Kiểm tra chuỗi có khớp biểu thức chính quy không.', example: '=REGEXMATCH(A2,"^SV\\d+$")', platforms: ['sheets'], examples: [
      demo('A2 = SV001', '=REGEXMATCH(A2,"^SV\\d+$")', 'TRUE'),
      demo('A2 = ABC001', '=REGEXMATCH(A2,"^SV\\d+$")', 'FALSE')
    ] },
    { name: 'REGEXEXTRACT', use: 'Trích phần chuỗi khớp biểu thức chính quy.', example: '=REGEXEXTRACT(A2,"\\d+")', platforms: ['sheets'], examples: [
      demo('A2 = SV001', '=REGEXEXTRACT(A2,"\\d+")', '001'),
      demo('A2 = LOP-CNTT-K49', '=REGEXEXTRACT(A2,"K\\d+")', 'K49')
    ] }
  ];

  function sheet(formula, guide, letters, rows, explanation) {
    return {
      formula: formula,
      guide: guide,
      letters: letters || ['A', 'B', 'C', 'D'],
      rows: rows,
      explanation: explanation
    };
  }

  var DEMO_SHEETS = {
    XLOOKUP: sheet('=XLOOKUP(B2,$F$2:$F$5,$G$2:$G$5,"Không tìm thấy")',
      'XLOOKUP tìm mã sinh viên ở cột F và trả về ID tương ứng ở cột G.',
      ['B', 'C', 'D', 'F', 'G'],
      [
        ['Mã cần tìm', 'Kết quả ID', 'Công thức', 'Mã SV gốc', 'ID'],
        ['SV001', '125', '=XLOOKUP(B2,$F$2:$F$5,$G$2:$G$5,"Không tìm thấy")', 'SV001', '125'],
        ['SV002', '128', '=XLOOKUP(B3,$F$2:$F$5,$G$2:$G$5,"Không tìm thấy")', 'SV002', '128'],
        ['SV999', 'Không tìm thấy', '=XLOOKUP(B4,$F$2:$F$5,$G$2:$G$5,"Không tìm thấy")', 'SV003', '130'],
        ['', '', '', 'SV004', '145']
      ],
      ['$F$2:$F$5 là vùng chứa mã sinh viên để tìm.', '$G$2:$G$5 là vùng trả về ID tương ứng cùng dòng.', 'B2 là mã cần tìm; nếu không có, công thức trả về "Không tìm thấy".']),
    VLOOKUP: sheet('=VLOOKUP(B2,$F$2:$H$5,3,FALSE)',
      'VLOOKUP tìm giá trị ở cột đầu tiên của bảng và trả về cột thứ N bên phải.',
      ['B', 'C', 'D', 'F', 'G', 'H'],
      [
        ['Mã cần tìm', 'Trạng thái', 'Công thức', 'Mã SV', 'Họ tên', 'Trạng thái'],
        ['SV001', 'Đang học', '=VLOOKUP(B2,$F$2:$H$5,3,FALSE)', 'SV001', 'An', 'Đang học'],
        ['SV002', 'Bảo lưu', '=VLOOKUP(B3,$F$2:$H$5,3,FALSE)', 'SV002', 'Bình', 'Bảo lưu'],
        ['SV004', 'Tốt nghiệp', '=VLOOKUP(B4,$F$2:$H$5,3,FALSE)', 'SV003', 'Chi', 'Đang học'],
        ['', '', '', 'SV004', 'Dũng', 'Tốt nghiệp']
      ],
      ['$F$2:$H$5 là bảng tra cứu, cột F bắt buộc là cột chứa mã cần tìm.', 'Số 3 nghĩa là trả về cột thứ 3 trong bảng F:H, tức cột H.', 'FALSE yêu cầu khớp chính xác mã sinh viên.']),
    INDEX: sheet('=INDEX($G$2:$G$5,MATCH(B2,$F$2:$F$5,0))',
      'INDEX lấy giá trị ở một vị trí; MATCH tìm vị trí đó.',
      ['B', 'C', 'D', 'F', 'G'],
      [
        ['Mã cần tìm', 'Tên trả về', 'Công thức', 'Mã SV', 'Họ tên'],
        ['SV001', 'An', '=INDEX($G$2:$G$5,MATCH(B2,$F$2:$F$5,0))', 'SV001', 'An'],
        ['SV003', 'Chi', '=INDEX($G$2:$G$5,MATCH(B3,$F$2:$F$5,0))', 'SV002', 'Bình'],
        ['SV004', 'Dũng', '=INDEX($G$2:$G$5,MATCH(B4,$F$2:$F$5,0))', 'SV003', 'Chi'],
        ['', '', '', 'SV004', 'Dũng']
      ],
      ['MATCH(B2,$F$2:$F$5,0) tìm vị trí của mã SV001 trong danh sách mã.', 'INDEX($G$2:$G$5, vị trí) lấy họ tên ở đúng vị trí vừa tìm.', 'Cặp INDEX+MATCH dùng tốt khi cột trả về nằm bên trái hoặc muốn tra cứu linh hoạt.']),
    MATCH: sheet('=MATCH(B2,$F$2:$F$6,0)',
      'MATCH trả về vị trí tương đối của giá trị trong một vùng.',
      ['B', 'C', 'D', 'F'],
      [
        ['Mã cần tìm', 'Vị trí', 'Công thức', 'Danh sách mã'],
        ['SV003', '3', '=MATCH(B2,$F$2:$F$6,0)', 'SV001'],
        ['SV005', '5', '=MATCH(B3,$F$2:$F$6,0)', 'SV002'],
        ['SV001', '1', '=MATCH(B4,$F$2:$F$6,0)', 'SV003'],
        ['', '', '', 'SV004'],
        ['', '', '', 'SV005']
      ],
      ['$F$2:$F$6 là danh sách cần dò.', 'B2 là mã cần tìm; SV003 nằm ở vị trí thứ 3 trong vùng F2:F6.', 'Số 0 nghĩa là tìm khớp chính xác.']),
    COUNTIF: sheet('=COUNTIF($B$2:$B$11,C2)',
      'COUNTIF quét vùng dữ liệu ban đầu rồi đếm xem giá trị cần đếm xuất hiện bao nhiêu lần.',
      ['B', 'C', 'D', 'E'],
      [
        ['Dữ liệu ban đầu', 'Dữ liệu cần đếm', 'Kết quả đếm', 'Công thức'],
        ['1', '1', '2', '=COUNTIF($B$2:$B$11,C2)'],
        ['1', '2', '2', '=COUNTIF($B$2:$B$11,C3)'],
        ['3', '3', '3', '=COUNTIF($B$2:$B$11,C4)'],
        ['3', '4', '2', '=COUNTIF($B$2:$B$11,C5)'],
        ['2', '5', '1', '=COUNTIF($B$2:$B$11,C6)'],
        ['3', '', '', ''],
        ['4', '', '', ''],
        ['2', '', '', ''],
        ['4', '', '', ''],
        ['5', '', '', '']
      ],
      ['$B$2:$B$11 là vùng dữ liệu ban đầu cần quét. Dấu $ giúp cố định vùng này khi kéo công thức xuống.', 'C2 là giá trị cần đếm ở dòng đầu tiên. Trong ví dụ này C2 = 1.', 'D2 trả về 2 vì số 1 xuất hiện 2 lần trong vùng $B$2:$B$11.', 'Khi kéo công thức xuống, vùng $B$2:$B$11 giữ nguyên, còn C2 đổi thành C3, C4, C5... để đếm số 2, số 3, số 4, số 5.']),
    COUNTIFS: sheet('=COUNTIFS($B$2:$B$8,F2,$C$2:$C$8,G2)',
      'COUNTIFS đếm số dòng thoả nhiều điều kiện cùng lúc.',
      ['B', 'C', 'F', 'G', 'H', 'I'],
      [
        ['Lớp', 'Trạng thái', 'Lớp cần đếm', 'Trạng thái cần đếm', 'Kết quả', 'Công thức'],
        ['CNTT', 'Đang học', 'CNTT', 'Đang học', '3', '=COUNTIFS($B$2:$B$8,F2,$C$2:$C$8,G2)'],
        ['CNTT', 'Đang học', 'KTPM', 'Đang học', '1', '=COUNTIFS($B$2:$B$8,F3,$C$2:$C$8,G3)'],
        ['KTPM', 'Đang học', 'CNTT', 'Bảo lưu', '1', '=COUNTIFS($B$2:$B$8,F4,$C$2:$C$8,G4)'],
        ['CNTT', 'Bảo lưu', '', '', '', ''],
        ['QTKD', 'Đang học', '', '', '', ''],
        ['CNTT', 'Đang học', '', '', '', ''],
        ['KTPM', 'Nghỉ học', '', '', '', '']
      ],
      ['$B$2:$B$8 là vùng lớp, F2 là lớp cần đếm.', '$C$2:$C$8 là vùng trạng thái, G2 là trạng thái cần đếm.', 'H2 trả về 3 vì có 3 dòng vừa là CNTT vừa là Đang học.']),
    SUMIF: sheet('=SUMIF($B$2:$B$7,E2,$C$2:$C$7)',
      'SUMIF cộng các số ở vùng tính tổng khi vùng điều kiện khớp một giá trị.',
      ['B', 'C', 'E', 'F', 'G'],
      [
        ['Lớp', 'Học phí', 'Lớp cần cộng', 'Tổng học phí', 'Công thức'],
        ['CNTT', '100', 'CNTT', '420', '=SUMIF($B$2:$B$7,E2,$C$2:$C$7)'],
        ['KTPM', '150', 'KTPM', '350', '=SUMIF($B$2:$B$7,E3,$C$2:$C$7)'],
        ['CNTT', '120', 'QTKD', '80', '=SUMIF($B$2:$B$7,E4,$C$2:$C$7)'],
        ['QTKD', '80', '', '', ''],
        ['KTPM', '200', '', '', ''],
        ['CNTT', '200', '', '', '']
      ],
      ['$B$2:$B$7 là vùng điều kiện lớp.', 'E2 là lớp cần cộng, ví dụ CNTT.', '$C$2:$C$7 là vùng số tiền được cộng; CNTT có 100 + 120 + 200 = 420.']),
    SUMIFS: sheet('=SUMIFS($D$2:$D$8,$B$2:$B$8,F2,$C$2:$C$8,G2)',
      'SUMIFS cộng số tiền khi nhiều điều kiện cùng đúng.',
      ['B', 'C', 'D', 'F', 'G', 'H', 'I'],
      [
        ['Lớp', 'Học kỳ', 'Số tiền', 'Lớp cần cộng', 'HK cần cộng', 'Kết quả', 'Công thức'],
        ['CNTT', 'HK1', '100', 'CNTT', 'HK1', '300', '=SUMIFS($D$2:$D$8,$B$2:$B$8,F2,$C$2:$C$8,G2)'],
        ['CNTT', 'HK2', '120', 'CNTT', 'HK2', '120', '=SUMIFS($D$2:$D$8,$B$2:$B$8,F3,$C$2:$C$8,G3)'],
        ['KTPM', 'HK1', '150', 'KTPM', 'HK1', '350', '=SUMIFS($D$2:$D$8,$B$2:$B$8,F4,$C$2:$C$8,G4)'],
        ['CNTT', 'HK1', '200', '', '', '', ''],
        ['QTKD', 'HK1', '80', '', '', '', ''],
        ['KTPM', 'HK1', '200', '', '', '', ''],
        ['CNTT', 'HK3', '90', '', '', '', '']
      ],
      ['$D$2:$D$8 là vùng số tiền cần cộng.', '$B$2:$B$8/F2 là điều kiện lớp.', '$C$2:$C$8/G2 là điều kiện học kỳ; chỉ dòng thoả cả hai điều kiện mới được cộng.']),
    IF: sheet('=IF(B2>=5,"Đạt","Không đạt")',
      'IF kiểm tra điều kiện; đúng thì trả kết quả thứ nhất, sai thì trả kết quả thứ hai.',
      ['B', 'C', 'D'],
      [
        ['Điểm', 'Kết quả', 'Công thức'],
        ['7', 'Đạt', '=IF(B2>=5,"Đạt","Không đạt")'],
        ['4', 'Không đạt', '=IF(B3>=5,"Đạt","Không đạt")'],
        ['5', 'Đạt', '=IF(B4>=5,"Đạt","Không đạt")']
      ],
      ['B2>=5 là điều kiện cần kiểm tra.', 'Nếu điều kiện đúng, C2 trả về "Đạt".', 'Nếu điều kiện sai, C2 trả về "Không đạt".']),
    IFS: sheet('=IFS(B2>=8.5,"Giỏi",B2>=7,"Khá",TRUE,"Đạt")',
      'IFS xét nhiều điều kiện theo thứ tự từ trên xuống.',
      ['B', 'C', 'D'],
      [
        ['Điểm', 'Xếp loại', 'Công thức'],
        ['9', 'Giỏi', '=IFS(B2>=8.5,"Giỏi",B2>=7,"Khá",TRUE,"Đạt")'],
        ['7.5', 'Khá', '=IFS(B3>=8.5,"Giỏi",B3>=7,"Khá",TRUE,"Đạt")'],
        ['6', 'Đạt', '=IFS(B4>=8.5,"Giỏi",B4>=7,"Khá",TRUE,"Đạt")']
      ],
      ['Excel kiểm tra điều kiện đầu tiên trước.', 'Nếu điểm >= 8.5 thì trả về Giỏi và dừng.', 'TRUE ở cuối là nhánh còn lại khi các điều kiện trên không đúng.']),
    AND: sheet('=AND(B2<>"",C2="Đang học")',
      'AND chỉ trả TRUE khi tất cả điều kiện đều đúng.',
      ['B', 'C', 'D', 'E'],
      [
        ['Mã SV', 'Trạng thái', 'Kết quả', 'Công thức'],
        ['SV001', 'Đang học', 'TRUE', '=AND(B2<>"",C2="Đang học")'],
        ['', 'Đang học', 'FALSE', '=AND(B3<>"",C3="Đang học")'],
        ['SV003', 'Bảo lưu', 'FALSE', '=AND(B4<>"",C4="Đang học")']
      ],
      ['B2<>"" kiểm tra mã sinh viên không trống.', 'C2="Đang học" kiểm tra trạng thái.', 'Cả hai điều kiện cùng đúng thì kết quả mới là TRUE.']),
    OR: sheet('=OR(B2="CNTT",B2="KTPM")',
      'OR trả TRUE nếu có ít nhất một điều kiện đúng.',
      ['B', 'C', 'D'],
      [
        ['Ngành', 'Kết quả', 'Công thức'],
        ['CNTT', 'TRUE', '=OR(B2="CNTT",B2="KTPM")'],
        ['QTKD', 'FALSE', '=OR(B3="CNTT",B3="KTPM")'],
        ['KTPM', 'TRUE', '=OR(B4="CNTT",B4="KTPM")']
      ],
      ['Công thức kiểm tra B2 có là CNTT hoặc KTPM không.', 'Chỉ cần một điều kiện đúng là trả TRUE.', 'Nếu không khớp điều kiện nào, kết quả là FALSE.']),
    FILTER: sheet('=FILTER(B2:D7,D2:D7=G2,"Không có")',
      'FILTER trả về các dòng thoả điều kiện.',
      ['B', 'C', 'D', 'G', 'H'],
      [
        ['Mã SV', 'Họ tên', 'Ngành', 'Ngành cần lọc', 'Kết quả'],
        ['SV001', 'An', 'CNTT', 'CNTT', 'SV001 - An - CNTT'],
        ['SV002', 'Bình', 'KTPM', '', 'SV003 - Chi - CNTT'],
        ['SV003', 'Chi', 'CNTT', '', 'SV005 - Hạnh - CNTT'],
        ['SV004', 'Dũng', 'QTKD', '', ''],
        ['SV005', 'Hạnh', 'CNTT', '', ''],
        ['SV006', 'Lan', 'KTPM', '', '']
      ],
      ['B2:D7 là bảng cần trả về.', 'D2:D7=G2 là điều kiện lọc ngành bằng CNTT.', 'Các dòng không thoả điều kiện sẽ bị bỏ qua.']),
    UNIQUE: sheet('=UNIQUE(B2:B8)',
      'UNIQUE lấy danh sách không trùng từ một vùng dữ liệu.',
      ['B', 'D', 'E'],
      [
        ['Ngành ban đầu', 'Kết quả không trùng', 'Công thức'],
        ['CNTT', 'CNTT', '=UNIQUE(B2:B8)'],
        ['CNTT', 'KTPM', ''],
        ['KTPM', 'QTKD', ''],
        ['QTKD', '', ''],
        ['CNTT', '', ''],
        ['KTPM', '', ''],
        ['CNTT', '', '']
      ],
      ['B2:B8 là danh sách có dữ liệu trùng.', 'UNIQUE chỉ giữ lại lần xuất hiện đầu tiên của mỗi giá trị.', 'Kết quả đổ xuống nhiều dòng ở cột D.']),
    SORT: sheet('=SORT(B2:B8)',
      'SORT sắp xếp dữ liệu theo thứ tự tăng hoặc giảm.',
      ['B', 'D', 'E'],
      [
        ['Tên ban đầu', 'Tên đã sắp xếp', 'Công thức'],
        ['Lan', 'An', '=SORT(B2:B8)'],
        ['An', 'Bình', ''],
        ['Dũng', 'Chi', ''],
        ['Bình', 'Dũng', ''],
        ['Chi', 'Lan', ''],
        ['Hạnh', 'Hạnh', ''],
        ['Minh', 'Minh', '']
      ],
      ['B2:B8 là vùng dữ liệu cần sắp xếp.', 'SORT mặc định sắp xếp tăng dần.', 'Kết quả đổ xuống cột D theo thứ tự mới.']),
    TEXTJOIN: sheet('=TEXTJOIN(" ",TRUE,B2:D2)',
      'TEXTJOIN nối nhiều ô thành một chuỗi, có thể bỏ qua ô trống.',
      ['B', 'C', 'D', 'E', 'F'],
      [
        ['Họ', 'Tên đệm', 'Tên', 'Họ tên', 'Công thức'],
        ['Nguyễn', 'Văn', 'An', 'Nguyễn Văn An', '=TEXTJOIN(" ",TRUE,B2:D2)'],
        ['Trần', '', 'Bình', 'Trần Bình', '=TEXTJOIN(" ",TRUE,B3:D3)'],
        ['Lê', 'Thị', 'Chi', 'Lê Thị Chi', '=TEXTJOIN(" ",TRUE,B4:D4)']
      ],
      ['Dấu cách " " là ký tự ngăn giữa các phần.', 'TRUE nghĩa là bỏ qua ô trống.', 'B2:D2 là các ô cần nối thành họ tên.']),
    TRIM: sheet('=TRIM(B2)',
      'TRIM xoá khoảng trắng đầu/cuối và gộp nhiều khoảng trắng giữa chữ.',
      ['B', 'C', 'D'],
      [
        ['Dữ liệu gốc', 'Sau TRIM', 'Công thức'],
        ['  Nguyễn   Văn   An  ', 'Nguyễn Văn An', '=TRIM(B2)'],
        [' SV001  ', 'SV001', '=TRIM(B3)'],
        ['CNTT   K49', 'CNTT K49', '=TRIM(B4)']
      ],
      ['B2 là chuỗi đang dư khoảng trắng.', 'TRIM bỏ khoảng trắng ở đầu/cuối.', 'Các khoảng trắng lặp ở giữa được gộp còn một.']),
    CLEAN: sheet('=CLEAN(B2)',
      'CLEAN xoá ký tự điều khiển không nhìn thấy khi copy từ web/hệ thống khác.',
      ['B', 'C', 'D'],
      [
        ['Dữ liệu gốc', 'Sau CLEAN', 'Công thức'],
        ['SV001↵', 'SV001', '=CLEAN(B2)'],
        ['An□', 'An', '=CLEAN(B3)'],
        ['CNTT', 'CNTT', '=CLEAN(B4)']
      ],
      ['B2 có thể chứa ký tự xuống dòng hoặc ký tự ẩn.', 'CLEAN loại bỏ các ký tự không in được.', 'Thường kết hợp CLEAN với TRIM để làm sạch dữ liệu nhập từ nhiều nguồn.']),
    SUBSTITUTE: sheet('=SUBSTITUTE(B2,"-","/")',
      'SUBSTITUTE thay một chuỗi con bằng chuỗi khác.',
      ['B', 'C', 'D'],
      [
        ['Dữ liệu gốc', 'Sau thay thế', 'Công thức'],
        ['01-09-2026', '01/09/2026', '=SUBSTITUTE(B2,"-","/")'],
        ['SV.001.2026', 'SV0012026', '=SUBSTITUTE(B3,".","")'],
        ['CNTT K49', 'CNTT-K49', '=SUBSTITUTE(B4," ","-")']
      ],
      ['B2 là chuỗi cần xử lý.', 'Tham số thứ hai là phần cần tìm.', 'Tham số thứ ba là phần thay vào.']),
    LEFT: sheet('=LEFT(B2,4)',
      'LEFT lấy N ký tự tính từ bên trái.',
      ['B', 'C', 'D'],
      [
        ['Dữ liệu gốc', 'Kết quả', 'Công thức'],
        ['CNTT-K49', 'CNTT', '=LEFT(B2,4)'],
        ['SV001', 'SV', '=LEFT(B3,2)'],
        ['2026-HK1', '2026', '=LEFT(B4,4)']
      ],
      ['B2 là chuỗi gốc.', 'Số 4 nghĩa là lấy 4 ký tự đầu tiên.', 'Dùng khi mã có cấu trúc cố định ở đầu chuỗi.']),
    RIGHT: sheet('=RIGHT(B2,3)',
      'RIGHT lấy N ký tự tính từ bên phải.',
      ['B', 'C', 'D'],
      [
        ['Dữ liệu gốc', 'Kết quả', 'Công thức'],
        ['CNTT-K49', 'K49', '=RIGHT(B2,3)'],
        ['SV000125', '125', '=RIGHT(B3,3)'],
        ['2026-HK1', 'HK1', '=RIGHT(B4,3)']
      ],
      ['B2 là chuỗi gốc.', 'Số 3 nghĩa là lấy 3 ký tự cuối.', 'Dùng để lấy hậu tố như lớp, học kỳ, mã cuối.']),
    MID: sheet('=MID(B2,4,4)',
      'MID lấy một đoạn ký tự từ giữa chuỗi.',
      ['B', 'C', 'D'],
      [
        ['Dữ liệu gốc', 'Kết quả', 'Công thức'],
        ['SV-2026-001', '2026', '=MID(B2,4,4)'],
        ['ABCDEF', 'BCD', '=MID(B3,2,3)'],
        ['LOP-CNTT-K49', 'CNTT', '=MID(B4,5,4)']
      ],
      ['B2 là chuỗi gốc.', 'Số 4 là vị trí bắt đầu lấy.', 'Số 4 tiếp theo là số ký tự cần lấy.']),
    LEN: sheet('=LEN(B2)',
      'LEN đếm tổng số ký tự trong một ô.',
      ['B', 'C', 'D'],
      [
        ['Dữ liệu', 'Số ký tự', 'Công thức'],
        ['SV001', '5', '=LEN(B2)'],
        ['001234567890', '12', '=LEN(B3)'],
        ['CNTT K49', '8', '=LEN(B4)']
      ],
      ['B2 là ô cần đếm ký tự.', 'LEN tính cả số, chữ và khoảng trắng.', 'Dùng để kiểm tra mã có đủ độ dài hay không.']),
    SEARCH: sheet('=SEARCH("CNTT",B2)',
      'SEARCH tìm vị trí chuỗi con, không phân biệt hoa thường.',
      ['B', 'C', 'D'],
      [
        ['Dữ liệu', 'Vị trí tìm thấy', 'Công thức'],
        ['Khoa CNTT', '6', '=SEARCH("CNTT",B2)'],
        ['khoa cntt', '6', '=SEARCH("CNTT",B3)'],
        ['Phòng QLĐT', '#VALUE!', '=SEARCH("CNTT",B4)']
      ],
      ['SEARCH tìm chữ CNTT trong B2.', 'Kết quả 6 nghĩa là CNTT bắt đầu từ ký tự thứ 6.', 'SEARCH không phân biệt hoa thường nên "cntt" vẫn tìm được.']),
    FIND: sheet('=FIND("CNTT",B2)',
      'FIND tìm vị trí chuỗi con và có phân biệt hoa thường.',
      ['B', 'C', 'D'],
      [
        ['Dữ liệu', 'Vị trí tìm thấy', 'Công thức'],
        ['Khoa CNTT', '6', '=FIND("CNTT",B2)'],
        ['khoa cntt', '#VALUE!', '=FIND("CNTT",B3)'],
        ['Phòng CNTT', '7', '=FIND("CNTT",B4)']
      ],
      ['FIND tìm đúng chữ CNTT trong B2.', 'Nếu dữ liệu là cntt thường, FIND không xem là khớp.', 'Dùng FIND khi cần phân biệt hoa/thường.']),
    TEXT: sheet('=TEXT(B2,"dd/mm/yyyy")',
      'TEXT định dạng số hoặc ngày thành chuỗi hiển thị theo mẫu.',
      ['B', 'C', 'D'],
      [
        ['Dữ liệu gốc', 'Sau TEXT', 'Công thức'],
        ['2026-09-01', '01/09/2026', '=TEXT(B2,"dd/mm/yyyy")'],
        ['1250000', '1,250,000', '=TEXT(B3,"#,##0")'],
        ['0.85', '85%', '=TEXT(B4,"0%")']
      ],
      ['B2 là ngày/số cần định dạng.', '"dd/mm/yyyy" là mẫu hiển thị ngày.', 'Kết quả của TEXT là chuỗi, phù hợp để làm báo cáo hoặc ghép text.']),
    DATEDIF: sheet('=DATEDIF(B2,C2,"Y")',
      'DATEDIF tính chênh lệch giữa hai mốc thời gian.',
      ['B', 'C', 'D', 'E'],
      [
        ['Ngày bắt đầu', 'Ngày kết thúc', 'Kết quả', 'Công thức'],
        ['01/01/2000', '01/01/2026', '26', '=DATEDIF(B2,C2,"Y")'],
        ['01/01/2026', '15/03/2026', '2', '=DATEDIF(B3,C3,"M")'],
        ['01/08/2026', '10/08/2026', '9', '=DATEDIF(B4,C4,"D")']
      ],
      ['B2 là ngày bắt đầu, C2 là ngày kết thúc.', '"Y" trả về số năm tròn giữa hai ngày.', 'Có thể đổi "Y" thành "M" để tính tháng hoặc "D" để tính ngày.']),
    YEAR: sheet('=YEAR(B2)',
      'YEAR lấy phần năm từ một ngày.',
      ['B', 'C', 'D'],
      [
        ['Ngày', 'Năm', 'Công thức'],
        ['15/08/2026', '2026', '=YEAR(B2)'],
        ['20/11/2003', '2003', '=YEAR(B3)'],
        ['01/01/2025', '2025', '=YEAR(B4)']
      ],
      ['B2 là ô chứa ngày.', 'YEAR chỉ lấy phần năm.', 'Dùng để lọc theo năm sinh, năm học, năm chứng từ.']),
    MONTH: sheet('=MONTH(B2)',
      'MONTH lấy phần tháng từ một ngày.',
      ['B', 'C', 'D'],
      [
        ['Ngày', 'Tháng', 'Công thức'],
        ['15/08/2026', '8', '=MONTH(B2)'],
        ['01/12/2026', '12', '=MONTH(B3)'],
        ['05/01/2025', '1', '=MONTH(B4)']
      ],
      ['B2 là ô chứa ngày.', 'MONTH trả về số tháng từ 1 đến 12.', 'Dùng để nhóm dữ liệu theo tháng.']),
    DAY: sheet('=DAY(B2)',
      'DAY lấy phần ngày trong tháng từ một ngày.',
      ['B', 'C', 'D'],
      [
        ['Ngày', 'Ngày trong tháng', 'Công thức'],
        ['15/08/2026', '15', '=DAY(B2)'],
        ['01/12/2026', '1', '=DAY(B3)'],
        ['29/02/2024', '29', '=DAY(B4)']
      ],
      ['B2 là ô chứa ngày.', 'DAY chỉ lấy số ngày trong tháng.', 'Dùng khi cần tách ngày/tháng/năm ra các cột riêng.']),
    QUERY: sheet('=QUERY(B1:D7,"select B,C where D = \'CNTT\'",1)',
      'QUERY trong Google Sheets lọc/truy vấn dữ liệu bằng cú pháp gần giống SQL.',
      ['B', 'C', 'D', 'F', 'G'],
      [
        ['Mã SV', 'Họ tên', 'Ngành', 'Kết quả mã', 'Kết quả tên'],
        ['SV001', 'An', 'CNTT', 'SV001', 'An'],
        ['SV002', 'Bình', 'KTPM', 'SV003', 'Chi'],
        ['SV003', 'Chi', 'CNTT', 'SV005', 'Hạnh'],
        ['SV004', 'Dũng', 'QTKD', '', ''],
        ['SV005', 'Hạnh', 'CNTT', '', ''],
        ['SV006', 'Lan', 'KTPM', '', '']
      ],
      ['B1:D7 là bảng nguồn có dòng tiêu đề.', 'select B,C nghĩa là chỉ lấy cột mã và họ tên.', 'where D = \'CNTT\' chỉ giữ các dòng ngành CNTT.']),
    ARRAYFORMULA: sheet('=ARRAYFORMULA(TRIM(B2:B6))',
      'ARRAYFORMULA áp dụng công thức cho cả vùng thay vì nhập từng dòng.',
      ['B', 'C', 'D'],
      [
        ['Tên gốc', 'Tên đã làm sạch', 'Công thức'],
        ['  An  ', 'An', '=ARRAYFORMULA(TRIM(B2:B6))'],
        [' Bình', 'Bình', ''],
        ['Chi  ', 'Chi', ''],
        ['  Dũng  ', 'Dũng', ''],
        ['Hạnh', 'Hạnh', '']
      ],
      ['B2:B6 là cả vùng cần xử lý.', 'TRIM làm sạch từng ô trong vùng.', 'ARRAYFORMULA giúp công thức đổ kết quả cho nhiều dòng cùng lúc.']),
    REGEXMATCH: sheet('=REGEXMATCH(B2,"^SV\\d+$")',
      'REGEXMATCH kiểm tra chuỗi có đúng mẫu biểu thức chính quy hay không.',
      ['B', 'C', 'D'],
      [
        ['Mã', 'Hợp lệ?', 'Công thức'],
        ['SV001', 'TRUE', '=REGEXMATCH(B2,"^SV\\d+$")'],
        ['ABC001', 'FALSE', '=REGEXMATCH(B3,"^SV\\d+$")'],
        ['SVABC', 'FALSE', '=REGEXMATCH(B4,"^SV\\d+$")']
      ],
      ['^SV nghĩa là chuỗi phải bắt đầu bằng SV.', '\\d+ nghĩa là phía sau phải có một hoặc nhiều chữ số.', '$ nghĩa là kết thúc chuỗi, không được dư ký tự khác.']),
    REGEXEXTRACT: sheet('=REGEXEXTRACT(B2,"\\d+")',
      'REGEXEXTRACT trích phần chuỗi khớp mẫu regex.',
      ['B', 'C', 'D'],
      [
        ['Dữ liệu', 'Phần trích ra', 'Công thức'],
        ['SV001', '001', '=REGEXEXTRACT(B2,"\\d+")'],
        ['LOP-CNTT-K49', '49', '=REGEXEXTRACT(B3,"\\d+")'],
        ['HK1-2026', '1', '=REGEXEXTRACT(B4,"\\d+")']
      ],
      ['B2 là chuỗi gốc.', '\\d+ là mẫu tìm một hoặc nhiều chữ số.', 'REGEXEXTRACT trả về phần đầu tiên khớp mẫu.'])
  };

  LIBRARY.forEach(function (item) {
    if (DEMO_SHEETS[item.name]) item.demoSheet = DEMO_SHEETS[item.name];
  });

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
      var demos = (f.examples || []).map(function (ex) {
        return ex.input + ' ' + ex.formula + ' ' + ex.result;
      }).join(' ');
      return deaccent(f.name + ' ' + f.use + ' ' + f.example + ' ' + demos).indexOf(q) !== -1;
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
