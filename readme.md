# Gen Script

Bản web của file `gen-script.xlsm` — nhập dữ liệu dạng bảng, nhận ngay câu lệnh
`INSERT`, `UPDATE` hoặc `SELECT`. Không cần Excel, không cần bật macro, chạy hoàn toàn
trong trình duyệt.

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
hiện tại `V1.5.0` ở góc phải dưới cùng.

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
│   │   ├── generator.js   ← lõi sinh SQL, không phụ thuộc DOM
│   │   ├── grid.js        ← lưới nhập liệu, dán từ Excel
│   │   └── app.js         ← ghép giao diện với lõi
│   └── vendor/
│       ├── xlsx.core.min.js      ← SheetJS 0.18.5, dùng để đọc file Excel
│       └── SheetJS-LICENSE.txt   ← Apache-2.0
└── docs/
    ├── Module1.bas        ← GetInsertSQL gốc
    └── Module2.bas        ← UpdateSQL gốc
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

## Lưu ý

- Script sinh ra chỉ là văn bản — app **không kết nối** tới CSDL nào. Đọc lại
  trước khi chạy trên môi trường thật.
- Font tải từ Google Fonts. Mạng nội bộ chặn thì trang vẫn chạy, chỉ đổi sang
  font hệ thống. Muốn bỏ hẳn thì xoá hai thẻ `<link>` Google Fonts trong
  `index.html`.
- File Excel rất lớn (vài chục nghìn dòng) sẽ làm trình duyệt chậm — nên chia
  nhỏ trước khi nạp.
