# ASC-GenScript

Công cụ xử lý dữ liệu cho nhân sự triển khai phần mềm. Chạy hoàn toàn trong
trình duyệt: không cần Excel, không cần bật macro, không có backend.

Từ `V2.2.0`, app có **hai workspace** với hai lưới/workbook riêng:

| Workspace | Làm gì |
| --- | --- |
| **Gen Script** | Bảng dữ liệu → câu lệnh `INSERT` / `UPDATE` / `SELECT` |
| **Formula Helper** | Nhu cầu nghiệp vụ → công thức Excel / Google Sheets |

Import Excel ở workspace nào thì dữ liệu chỉ nạp vào workspace đó. Chuyển qua
lại giữa Gen Script và Formula Helper không làm ghi đè lưới còn lại.

---

## Deploy lên Netlify

Trang tĩnh thuần, **không có bước build**.

**Cách 1 — kéo thả (nhanh nhất)**

1. Giải nén file zip.
2. Mở <https://app.netlify.com/drop>.
3. Kéo cả thư mục vừa giải nén thả vào trang đó.

**Cách 2 — qua Git**

Đẩy thư mục này lên GitHub/GitLab rồi nối vào Netlify. `netlify.toml` đã khai
báo sẵn:

| Mục | Giá trị |
| --- | --- |
| Build command | *(để trống)* |
| Publish directory | `.` |

**Cách 3 — Netlify CLI**

```bash
npm i -g netlify-cli
netlify deploy --prod --dir .
```

---

## Cách dùng

| Trong Excel cũ | Trong app |
| --- | --- |
| Tên sheet | Ô **Tên bảng** |
| Sheet | Tab ở đáy khung lưới |
| Dòng 1 | Dòng đầu của lưới (tên cột) |
| Nút *Insert* | Nút **INSERT** |
| Nút *Update* | Nút **UPDATE** |
| Tra cứu dữ liệu | Nút **SELECT** |

Ba cách đưa dữ liệu vào:

- **Dán từ Excel** — bôi đen vùng dữ liệu, `Ctrl+C`, bấm vào ô đầu của lưới rồi
  `Ctrl+V`. Lưới tự giãn vừa đủ.
- **Nhập từ Excel** — chọn file `.xlsx / .xlsm / .xls / .csv`, mỗi sheet thành
  một tab, tên sheet thành tên bảng. Kéo thả file vào cửa sổ cũng được.
- **Gõ tay** — di chuyển bằng `Tab`, `Enter`, phím mũi tên.

Sau khi hoàn thành dữ liệu, bấm **Gen script** để mở modal lớn review toàn bộ
script. Có thể click trực tiếp vào vùng script hoặc bấm **Sao chép script** để
copy; nút **Tải .sql** nằm ngay trong modal.

### Thao tác lưới như Excel

- Click rồi kéo chuột để chọn một vùng ô; click tiêu đề cột hoặc số dòng để
  chọn toàn cột/toàn dòng.
- Click một ô chỉ chọn ô, double click hoặc `F2` mới vào chế độ sửa và hiện
  cursor typing.
- `Arrow` để di chuyển qua lại giữa các ô; `Shift+Arrow` để mở rộng vùng chọn.
- `Ctrl+Shift+Arrow` chọn nhanh tới mép bảng theo hướng mũi tên.
- `Ctrl+C`, `Ctrl+X`, `Ctrl+V` để sao chép, cắt và dán vùng dữ liệu. Khi copy
  một ô rồi chọn nhiều ô để paste, giá trị sẽ được fill vào toàn bộ vùng chọn.
- `Delete` để xoá nội dung vùng đang chọn.
- `Ctrl+Z` hoàn tác tối đa 100 bước; `Ctrl+Y` hoặc `Ctrl+Shift+Z` để làm lại.
- Kéo chấm vuông ở góc phải dưới vùng chọn để fill xuống hoặc sang phải. Khi
  vùng nguồn có từ hai số, lưới tiếp tục chuỗi số theo bước chênh lệch.
- `Ctrl+D` fill xuống từ dòng đầu vùng chọn; `Ctrl+R` fill sang phải từ cột đầu.
- Click chuột phải để sao chép, dán, fill, xoá nội dung, xoá dòng hoặc xoá cột.
- Dấu `×` trên tiêu đề cột và đầu dòng vẫn cho phép xoá nhanh cột/dòng.

Phím tắt cho script: `Ctrl+Shift+C` sao chép script, `Ctrl+S` tải file `.sql`.

Dữ liệu được lưu riêng theo thiết bị/profile trình duyệt. Mỗi máy mở app sẽ có
kho dữ liệu riêng, không ghi đè lẫn nhau khi nhiều người cùng thao tác. Hai tab
trong cùng một profile trình duyệt vẫn nhìn thấy cùng dữ liệu của profile đó.
Không có máy chủ nào nhận dữ liệu — mọi thứ xử lý tại máy bạn.

Giao diện mặc định là dark mode. Bấm công tắc biểu tượng mặt trời/mặt trăng trên
thanh trên cùng để đổi theme; lựa chọn theme cũng được lưu riêng trên thiết bị
đang dùng.

Footer hiển thị bản quyền **© 2026 HuyVo. All rights reserved.** và phiên bản
hiện tại `V2.6.0` ở góc phải dưới cùng.

### INSERT

```sql
SET IDENTITY_INSERT [DM_TrangThai] ON
INSERT INTO [DM_TrangThai]
(
    [ID],
    [Loai],
    [TenTrangThai]
)
VALUES
(
    1,
    1,
    N'Đề xuất sửa chữa'
)
INSERT INTO [DM_TrangThai]
(
    [ID],
    [Loai],
    [TenTrangThai]
)
VALUES
(
    2,
    1,
    N'Đã phân công'
)
SET IDENTITY_INSERT [DM_TrangThai] OFF
```

### UPDATE

Cột đầu tiên là điều kiện `WHERE`, các cột còn lại vào `SET`:

```sql
UPDATE [DM_TrangThai]
SET
    [Loai] = 1,
    [TenTrangThai] = N'Đề xuất sửa chữa'
WHERE
    [ID] = 1
```

Khoá gồm nhiều cột thì tăng ô **Số cột làm WHERE** — các cột khoá được tô hổ
phách ngay trên lưới để nhìn là biết.

### SELECT

Mặc định **Số cột làm WHERE = 0** sẽ sinh câu tra cứu toàn bảng:

```sql
SELECT
    *
FROM
    [DM_TrangThai]
```

Nếu **Số cột làm WHERE = 1**, app lấy cột đầu tiên làm điều kiện. Một dòng dữ
liệu dùng `=`, nhiều dòng dữ liệu dùng `IN (...)`:

```sql
SELECT
    *
FROM
    [DM_TrangThai]
WHERE
    [ID] IN (1, 2)
```

Nếu **Số cột làm WHERE >= 2**, app lấy các cột đầu tiên làm nhóm điều kiện
`AND`; nhiều dòng dữ liệu được nối bằng `OR` để giữ đúng từng cặp giá trị:

```sql
SELECT
    *
FROM
    [DM_TrangThai]
WHERE
    ([ID] = 1
    AND [Loai] = 1)
    OR ([ID] = 2
    AND [Loai] = 1)
```

Có 3 kiểu WHERE SELECT:

| Kiểu | SQL sinh ra | Khi nên dùng |
| --- | --- | --- |
| Theo dòng (AND/OR) | `([C1] = ... AND [C2] = ...) OR (...)` | Cần giữ đúng từng cặp/từng dòng dữ liệu |
| IN từng cột + AND | `[C1] IN (...) AND [C2] IN (...)` | Cần lọc giao nhau giữa nhiều danh sách giá trị |
| IN từng cột + OR | `[C1] IN (...) OR [C2] IN (...)` | Cần tìm rộng, khớp bất kỳ danh sách nào |

Các kiểu có thể bổ sung thêm về sau nếu cần: `NOT IN` để loại trừ danh sách,
`LIKE` để tìm gần đúng chuỗi, `BETWEEN` để lọc khoảng giá trị/ngày, hoặc
`EXISTS` để tra theo bảng phụ.

Khi bật **Bỏ qua dòng trống**, riêng chế độ SELECT sẽ bỏ qua cả các ô trống
trong vùng cột WHERE. Ô trống sẽ không sinh `IS NULL`; dòng/cột nào có dữ liệu
thì mới được đưa vào điều kiện.

Tuỳ chọn **Xuống dòng IN sau mỗi N giá trị** giúp danh sách `IN (...)` dễ đọc
hơn khi có nhiều giá trị. Ví dụ đặt `5`:

```sql
WHERE
    [ID] IN (
        1, 2, 3, 4, 5,
        6, 7, 8, 9, 10
    )
```

Để `0` nếu muốn giữ danh sách `IN` trên một dòng như các bản cũ.

---

## Formula Helper

Bấm **Formula Helper** trên thanh tiêu đề. Điểm khác biệt so với việc tra Google:
app không hỏi *"bạn muốn dùng hàm nào"* mà hỏi *"bạn muốn làm gì"*, rồi tự chọn
công thức phù hợp.

Formula Helper có cấu trúc workbench: ở màn hình danh sách công cụ, lưới được
ẩn để tập trung chọn đúng nhu cầu. Sau khi chọn một công cụ, lưới dữ liệu hiện
ở vùng chính và thanh cấu hình/công thức nằm ở bên phải. Thanh này có thể kéo
resize từ mép trái hoặc ẩn/hiện bằng nút **Ẩn/Hiện thanh công cụ** để mở rộng
lưới.

Lưới Formula Helper độc lập với Gen Script, hỗ trợ dán Excel, chọn vùng,
copy/cut/paste, xoá dữ liệu, kéo fill, `Ctrl+Z`/`Ctrl+Y`, phím mũi tên và sửa ô
trực tiếp như lưới Gen Script. Khi sửa trên lưới này, danh sách sheet/cột,
preview và công thức được cập nhật theo workbook Formula Helper hiện tại.

Panel **Workbook / Xem trước kết quả** được ẩn mặc định để lưới rộng nhất có
thể. Bấm **Xem kết quả** trên thanh Formula Helper khi cần mở lại phần review.

### Các công cụ

| Nhóm | Công cụ | Hàm sinh ra |
| --- | --- | --- |
| Tra cứu | Tra cứu dữ liệu | `XLOOKUP` · `VLOOKUP` · `INDEX+MATCH` · `FILTER` |
| Kiểm tra | Kiểm tra trùng | `COUNTIF` · `COUNTIFS` |
| Kiểm tra | Kiểm tra tồn tại | `COUNTIF` · `XMATCH` · `XLOOKUP` |
| Kiểm tra | Kiểm tra dữ liệu rỗng | `IF` + `OR` |
| So sánh | So sánh dữ liệu | `IF` (+ `TRIM`, `LOWER`) |
| So sánh | So sánh hai danh sách | `COUNTIF` |
| Lọc | Lọc dữ liệu | `FILTER` · `QUERY` (Google Sheets) |
| Điều kiện | Kiểm tra điều kiện | `IF` · `IFS` · `AND` · `OR` lồng nhau |
| Text | Làm sạch Text | `TRIM` `CLEAN` `UPPER` `LOWER` `PROPER` `SUBSTITUTE` `LEFT` `RIGHT` `MID` `LEN` |
| Text | Ghép dữ liệu | `TEXTJOIN` · toán tử `&` |
| Ngày | Xử lý ngày | `TEXT` `DATEDIF` `YEAR` `MONTH` `DAY` · năm học |
| Nâng cao | Đếm & Tính tổng | `COUNTIF` `COUNTIFS` `SUM` `SUMIF` `SUMIFS` |
| Nâng cao | Lấy danh sách duy nhất | `UNIQUE` · `SORT` |
| SQL | Tạo giá trị SQL | `="'"&A2&"'"` |
| SQL | Tạo danh sách SQL IN | `TEXTJOIN` + xuất thẳng `WHERE ... IN (...)` |

### Chọn cột theo tên, không theo A/B/C

Dropdown hiển thị dạng `Sheet.Cột` kèm chú thích cột Excel tương ứng:

```
SinhVien.MaSinhVien       ·  Cột A
Database.MaSinhVien       ·  Cột B
Database.IDSinhVien       ·  Cột A
```

Nếu đang chọn sẵn một cột trên lưới Formula Helper, bấm **Dùng cột đang chọn**
để điền nhanh.

### Dấu phân cách `,` và `;`

Excel trên máy cài vùng Việt Nam thường dùng `;`, máy cài vùng Mỹ dùng `,`.
Chọn ở thanh trên cùng — mặc định **Tự động** đoán theo locale trình duyệt.
Lựa chọn được ghi nhớ cho lần sau.

```
=XLOOKUP(A2,Database!B:B,Database!A:A,"")
=XLOOKUP(A2;Database!B:B;Database!A:A;"")
```

### Sheet có khoảng trắng

Tự bọc nháy đơn, không cần tự thêm:

```
'CONFIG - Process'!A:A
```

### Xem trước kết quả

Với các công cụ mà JavaScript mô phỏng được (tra cứu, trùng, tồn tại, so sánh,
làm sạch text, ghép chuỗi, SQL IN), panel bên phải chạy thử trên vài chục dòng
đầu để đối chiếu trước khi dán công thức vào Excel. Đây **không phải** bộ tính
toán Excel đầy đủ — app không cố hỗ trợ mọi hàm.

### Lịch sử và yêu thích

Mỗi công thức tạo ra được ghi vào Lịch sử ngay trên máy (50 mục gần nhất). Bấm
vào một mục sẽ nạp lại **toàn bộ cấu hình**, không chỉ dán lại chữ. Bấm
**★ Lưu yêu thích** để đặt tên riêng và ghim lên đầu trang.

Nếu Workbook đã đổi và sheet/cột cũ không còn, app báo rõ *"Sheet Database không
còn tồn tại trong Workbook hiện tại"* thay vì sinh công thức sai.

### Chưa nhập Excel vẫn dùng được

Ở mỗi ô chọn cột, đổi sang chế độ **Ô / Vùng** hoặc **Tuỳ ý** rồi gõ trực tiếp
`A2`, `$A:$A`, `Database!B:B`.

### Việc V2.0.0 không làm

Không áp công thức thẳng vào Workbook, không tính lại toàn sheet, không macro,
không Power Query, không Pivot, không AI, không backend. Luồng cố định là:
dựng công thức → xem trước → sao chép → dán vào Excel.

---

## Quy tắc chuyển giá trị

| Ô nhập | Kết quả |
| --- | --- |
| Trống | `NULL` |
| Chữ `NULL` | `NULL` *(tắt được trong Tuỳ chọn)* |
| `123`, `-2.5`, `1e5` | để trần, không đóng nháy |
| `001129` | `N'001129'` — giữ dạng chuỗi |
| Còn lại | `N'...'`, dấu `'` được nhân đôi thành `''` |

---

## Khác gì so với VBA gốc

Ba chỗ bản VBA sinh ra SQL sai, đã sửa:

1. **Mã có số 0 đứng đầu.** `IsNumeric("001129")` trả về True nên VBA sinh
   `VALUES (001129)`, vào CSDL thành số `1129`. Bản web giữ nguyên dạng chuỗi.
2. **Số có dấu phân cách hàng nghìn.** `1,000` bị VBA coi là số, sinh ra
   `VALUES (1,000)` — thừa một cột, câu lệnh hỏng. Bản web coi đây là chuỗi.
3. **Tên định danh chưa bọc ngoặc.** VBA viết `SET IDENTITY_INSERT Ten Bang ON`
   và `WHERE Ma Cot = 1` không có `[ ]`, gãy khi tên có dấu cách. Bản web bọc
   `[ ]` nhất quán ở mọi nơi.

Phần thêm mới:

- Khoá `WHERE` gồm nhiều cột (VBA cũ cố định 1 cột).
- Sinh câu `SELECT * FROM Table`, hỗ trợ `WHERE`, `IN`, nhiều cột điều kiện và
  tự nhận diện kiểu dữ liệu giống INSERT/UPDATE.
- Chọn kiểu WHERE cho SELECT: theo dòng (AND/OR), `IN` từng cột + `AND`, hoặc
  `IN` từng cột + `OR`.
- Cột khoá trống sinh `IS NULL` thay vì `= NULL` trong UPDATE, hoặc trong
  SELECT khi không bật **Bỏ qua dòng trống**.
- Chọn hệ CSDL: SQL Server, MySQL/MariaDB, PostgreSQL, Oracle — đổi cả kiểu bọc
  định danh lẫn tiền tố `N`.
- Gộp nhiều dòng vào một lệnh `INSERT` (tự cắt lô 1000 dòng cho SQL Server).
- Chèn `GO` sau mỗi N câu lệnh, tuỳ chọn kết thúc bằng `;`.
- Bỏ qua dòng trống, cắt khoảng trắng thừa, cảnh báo khi tên cột trùng hoặc
  thiếu tên bảng.
- Tô màu cú pháp, đếm số câu lệnh và dung lượng script.
- Trong modal review script, chọn `Format: Chuẩn` để dùng dạng clean code
  giống V1.3.0, hoặc `Format: Mặc định` để sao chép/tải script compact giống
  V1.2.2.
- Khi sửa tên bảng, tên tab sheet bên dưới đổi theo ngay để tránh nhầm bảng
  đang thao tác.
- Header cột có badge kiểu dữ liệu `Auto` / `Chuỗi` / `Số`. Bấm badge hoặc
  click chuột phải vào cột để đổi kiểu. Chọn `Chuỗi` cho các mã dạng số nhưng
  database lưu `varchar`, ví dụ `MaSinhVien`, để script sinh `N'24810860071'`
  thay vì `24810860071`.

Mã VBA gốc được giữ trong `docs/` để đối chiếu.

---

## Cấu trúc

```
.
├── index.html
├── netlify.toml
├── assets/
│   ├── css/styles.css
│   ├── js/
│   │   ├── generator.js          ← lõi sinh SQL, không phụ thuộc DOM
│   │   ├── grid.js               ← lưới nhập liệu, dán từ Excel
│   │   ├── workbook.js           ← API đọc workbook riêng của Formula Helper
│   │   ├── formula-engine.js     ← lõi sinh công thức, không phụ thuộc DOM
│   │   ├── formula-presets.js    ← danh mục công cụ + schema builder
│   │   ├── formula-explainer.js  ← giải thích + mô phỏng kết quả
│   │   ├── formula-history.js    ← lịch sử, yêu thích, thiết lập
│   │   ├── formula-helper.js     ← giao diện Formula Helper
│   │   └── app.js                ← app shell, ghép mọi thứ lại
│   └── vendor/
│       ├── xlsx.core.min.js      ← SheetJS 0.18.5, dùng để đọc file Excel
│       └── SheetJS-LICENSE.txt   ← Apache-2.0
├── tests/
│   ├── formula-engine.test.js    ← 61 test cho bộ sinh công thức
│   └── app-smoke.test.js         ← test tích hợp (cần jsdom)
└── docs/
    ├── Module1.bas        ← GetInsertSQL gốc
    └── Module2.bas        ← UpdateSQL gốc
```

Hai lõi `generator.js` và `formula-engine.js` tách bạch: Gen Script không biết gì
về công thức Excel, Formula Helper không biết gì về SQL. `workbook.js` chỉ là
lớp đọc ổn định cho workbook riêng của Formula Helper.

### Chạy test

```bash
node tests/formula-engine.test.js     # không cần cài gì thêm
npm install jsdom && node tests/app-smoke.test.js
```

`generator.js` không đụng tới DOM nên tái sử dụng được ở Node:

```js
global.window = global;
require('./assets/js/generator.js');

const out = window.SqlGen.buildInsert('DM_TrangThai', [
  ['ID', 'TenTrangThai'],
  ['1', 'Đề xuất sửa chữa']
], { dialect: 'mssql' });

console.log(out.sql, out.warnings, out.stats);
```

---

## Changelog

### V2.6.0

**Mới**

- Thư viện hàm trở lại dạng danh sách gọn: mỗi card chỉ hiển thị tên hàm, công
  dụng và công thức mẫu.
- Bấm vào một hàm sẽ mở modal mô phỏng riêng bằng lưới Excel nhỏ, gồm dữ liệu
  mẫu, công thức và kết quả dự kiến.

**Cải tiến**

- Phần ví dụ không còn hiển thị trực tiếp trên toàn bộ card, giúp màn hình
  Thư viện hàm thoáng và dễ lướt hơn.

### V2.5.0

**Mới**

- Thư viện hàm có thêm ví dụ dữ liệu đầu vào, công thức mẫu và kết quả mô phỏng
  cho từng hàm Excel/Google Sheets.
- Tìm kiếm trong Thư viện hàm có thể tìm theo cả nội dung ví dụ và kết quả.
- Gen Script có tuỳ chọn **Xuống dòng IN sau mỗi N giá trị** cho SELECT dùng
  `IN (...)`; nhập `0` để giữ một dòng.

**Cải tiến**

- Bổ sung dữ liệu mô phỏng phía sau để phục vụ màn hình học hàm.

### V2.4.0

**Mới**

- Formula Helper đổi sang workbench dễ thao tác hơn: chọn công cụ trước, sau đó
  mới hiện lưới dữ liệu.
- Khi đang dùng công cụ, lưới là vùng chính; cấu hình và công thức nằm trong
  sidebar bên phải.
- Sidebar công cụ có thể ẩn/hiện và kéo resize từ mép trái.

**Cải tiến**

- Quay lại danh sách công cụ sẽ ẩn lưới để chọn công cụ khác rõ ràng hơn.
- Chọn công cụ mới sẽ hiện lại lưới với đúng dữ liệu workbook Formula Helper đã
  nhập trước đó.

### V2.3.0

**Mới**

- Formula Helper chuyển sang layout chuyên nghiệp hơn: thanh tuỳ chọn trên cùng,
  cấu hình/chức năng ở phía trên và lưới dữ liệu full width ở phía dưới.
- Lưới Formula Helper mở mặc định để thao tác dữ liệu ngay.
- Panel Workbook / Xem trước kết quả được ẩn mặc định và mở bằng nút
  **Xem kết quả** khi cần review.

**Sửa lỗi**

- Đóng lưới rồi mở lại không còn mất table/grid.
- Khi sửa header hoặc dữ liệu trên lưới Formula Helper, dropdown chọn cột,
  công thức và preview được cập nhật theo workbook hiện tại.

### V2.2.0

**Mới**

- Gen Script và Formula Helper dùng hai lưới/workbook riêng, không còn chia sẻ
  cùng một bộ sheet.
- **Nhập từ Excel** nạp dữ liệu vào đúng workspace đang mở: Gen Script hoặc
  Formula Helper.
- Formula Helper có workbook mặc định `FormulaData` để mở lưới và nhập dữ liệu
  ngay, không cần tạo dữ liệu từ Gen Script trước.

**Cải tiến**

- Nút **Dùng cột đang chọn** trong Formula Helper lấy cột từ lưới Formula Helper
  đang mở.
- Hướng dẫn trong app và README ghi rõ dữ liệu hai workspace độc lập để tránh
  hiểu nhầm khi nhiều người thao tác cùng lúc.

### V2.1.0

**Mới**

- Formula Helper có nút **Mở lưới dữ liệu** để thao tác workbook dạng grid như
  Gen Script: dán dữ liệu, fill, copy/cut/paste, xoá ô/dòng/cột, undo/redo.
- Lưới trong Formula Helper sửa trực tiếp workbook nội bộ của Formula Helper.

**Sửa lỗi**

- Thay file `assets/vendor/xlsx.core.min.js` bị thiếu đoạn cuối, gây lỗi
  `XLSX is not defined` khi bấm **Nhập từ Excel**.
- `app.js` đọc SheetJS qua `window.XLSX` và báo lỗi rõ hơn nếu thư viện Excel
  chưa được tải.
- Khi đổi sheet từ Formula Helper, app không commit lại grid Gen Script đang ẩn
  để tránh ghi đè dữ liệu vừa sửa trong Formula Helper.

### V2.0.0

**Mới**

- Workspace **Formula Helper** với 15 công cụ theo nhu cầu nghiệp vụ.
- Chọn nền tảng **Microsoft Excel** hoặc **Google Sheets**.
- Dấu phân cách `,` / `;` — tự động hoặc chọn tay, có ghi nhớ.
- Tự escape tên sheet có khoảng trắng và ký tự đặc biệt.
- Xem trước kết quả bằng JavaScript trên vài chục dòng đầu.
- Giải thích từng thành phần công thức và cảnh báo các trường hợp dễ sai.
- Lịch sử công thức và mục yêu thích, nạp lại được cả cấu hình.
- Thư viện hàm tra nhanh công dụng; bấm từng hàm để mở modal mô phỏng kết quả.
- Hướng dẫn riêng cho từng workspace qua nút **Hướng dẫn**.

**Cải tiến**

- Tách `WorkbookState` khỏi bộ sinh SQL để Formula Helper đọc dữ liệu qua API
  workbook ổn định.
- Chia nhỏ JavaScript thành các module theo trách nhiệm.
- Bộ test tự động cho lõi sinh công thức và luồng tích hợp.

**Không đổi**

- Toàn bộ chức năng Gen Script của `V1.5.0` giữ nguyên: INSERT, UPDATE, SELECT,
  lưới dữ liệu, tab bảng, kiểu dữ liệu cột, nhập Excel, theme, phím tắt.

---

## Lưu ý

- Script sinh ra chỉ là văn bản — app **không kết nối** tới CSDL nào. Đọc lại
  trước khi chạy trên môi trường thật.
- Font tải từ Google Fonts. Mạng nội bộ chặn thì trang vẫn chạy, chỉ đổi sang
  font hệ thống. Muốn bỏ hẳn thì xoá hai thẻ `<link>` Google Fonts trong
  `index.html`.
- File Excel rất lớn (vài chục nghìn dòng) sẽ làm trình duyệt chậm — nên chia
  nhỏ trước khi nạp.
