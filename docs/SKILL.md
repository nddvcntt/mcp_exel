# SKILL: mcp_excel — Excel Generation & Analysis

> **Version**: 3.0 | **MCP Endpoint**: `http://localhost:5003/mcp` | **Tools**: 25 | **npm test**: 43/43 PASS

## Mô tả & Khi nào kích hoạt

Skill tạo, đọc, phân tích và format file Excel (.xlsx) **offline hoàn toàn** qua `uvx excel-mcp-server`.

**Kích hoạt skill này khi user dùng các từ khóa sau:**
> xuất Excel · tạo bảng tính · báo cáo Excel · file xlsx · pivot table · biểu đồ Excel · bảng số liệu · tổng hợp dữ liệu ra Excel · download Excel · phân tích dữ liệu Excel

**KHÔNG dùng skill này khi:**
- User chỉ muốn xem bảng trong chat (dùng Markdown table)
- User muốn slide/trình chiếu (dùng skill `web_ppt`)
- User chỉ hỏi về công thức Excel (trả lời trực tiếp, không cần tạo file)

---

## ⓵ Quy tắc Bắt buộc (KHÔNG được vi phạm)

1. **`create_workbook` PHẢI là bước đầu tiên** — không ghi dữ liệu vào file chưa tồn tại.
2. **`data` trong `write_data_to_excel` là `list of lists`** — hàng đầu = header, các hàng sau = data. **Không dùng list of dicts.**
3. **Format header NGAY SAU ghi data** — `format_range` bold + màu nền dòng 1 trước khi làm bất cứ việc gì khác.
4. **Tên file: chữ thường, gạch dưới, có ngày** — `bao_cao_doanh_thu_20260427.xlsx`. Không có dấu tiếng Việt trong tên file.
5. **Filepath là tên file tương đối** — KHÔNG dùng đường dẫn tuyệt đối như `D:\...`.
6. **Validate công thức trước khi apply** — gọi `validate_formula_syntax` với công thức phức tạp (có IF, VLOOKUP, nested functions).
7. **Mỗi chủ đề một sheet riêng** — nếu báo cáo có nhiều phần (doanh thu + chi phí + tổng hợp), tạo worksheet riêng cho mỗi phần.
8. **Sau khi hoàn tất, thông báo tên file** — user cần biết tên file để download: `GET /download/<tên_file>`.

---

## ⓶ Workflow Chuẩn

### A. Báo cáo đơn giản (1 sheet)

```
1. create_workbook       { filepath: "ten_bao_cao_YYYYMMDD.xlsx" }
2. write_data_to_excel   { filepath, sheet_name: "Sheet1", data: [[headers...], [row1...], ...] }
3. format_range          { ..., start_cell: "A1", end_cell: "X1", bold: true, bg_color: "FF4472C4", font_color: "FFFFFFFF" }
4. apply_formula         { ..., cell: "B{n+1}", formula: "=SUM(B2:B{n})" }  ← cho cột số
5. [Tùy] create_table    { ..., data_range: "A1:X{n}", table_name: "BangDuLieu" }
```

### B. Báo cáo đa sheet (chuyên nghiệp)

```
1. create_workbook
2. Sheet "DuLieu":   write_data_to_excel → format_range header → apply_formula tổng
3. Sheet "TongHop":  create_worksheet → write_data_to_excel → merge_cells tiêu đề → format_range
4. Sheet "BieuDo":   create_worksheet → create_chart (data_range trỏ về sheet DuLieu)
5. [Tùy] Sheet "PivotTable": create_worksheet → create_pivot_table
```

### C. Đọc & Phân tích file có sẵn

```
1. read_data_from_excel  { filepath: "file_co_san.xlsx", sheet_name: "Sheet1" }
2. get_workbook_metadata { filepath } ← biết có bao nhiêu sheets, bao nhiêu rows
3. Phân tích dữ liệu đọc được → viết kết quả phân tích vào sheet mới
4. create_worksheet → write_data_to_excel (kết quả phân tích)
```

### D. Zebra-row (bảng dữ liệu lớn, chuyên nghiệp)

```
1. format_range header   → bg_color: "FF4472C4", font_color: "FFFFFFFF"
2. format_range row chẵn → bg_color: "FFF2F2F2" (từng cặp hàng hoặc toàn bộ body)
3. create_table          → tự động thêm filter + auto-style zebra
   # Khuyến nghị: dùng create_table thay for loop format từng hàng — nhanh hơn
```

---

## ⓷ Danh sách 25 Tools — Params Chính Xác

> ⚠️ **Params sau đã được verify từ schema thực tế** — không tự ý đổi tên.

### Nhóm 1: Workbook & Worksheet
| Tool | Args bắt buộc | Args tùy chọn |
|------|--------------|--------------|
| `create_workbook` | `filepath` | — |
| `get_workbook_metadata` | `filepath` | — |
| `create_worksheet` | `filepath`, `sheet_name` | — |
| `copy_worksheet` | `filepath`, `source_sheet`, `target_sheet` | — |
| `delete_worksheet` | `filepath`, `sheet_name` | — |
| `rename_worksheet` | `filepath`, `old_name`, `new_name` | — |

### Nhóm 2: Đọc/Ghi dữ liệu
| Tool | Args bắt buộc | Ghi chú |
|------|--------------|---------|
| `write_data_to_excel` | `filepath`, `sheet_name`, `data` | `data` = `[[h1,h2],[v1,v2],...]` |
| `read_data_from_excel` | `filepath`, `sheet_name` | Trả về text mô tả dữ liệu |
| `copy_range` | `filepath`, `sheet_name`, `source_start`, `source_end`, `target_start` | `source_start`/`end` = cell addr (A1) |
| `delete_range` | `filepath`, `sheet_name`, `start_cell`, `end_cell` | — |

### Nhóm 3: Hàng & Cột
| Tool | Args bắt buộc | Ghi chú |
|------|--------------|---------|
| `insert_rows` | `filepath`, `sheet_name`, `start_row` | `count` mặc định 1 |
| `insert_columns` | `filepath`, `sheet_name`, `start_col` | `count` mặc định 1 |
| `delete_sheet_rows` | `filepath`, `sheet_name`, `start_row` | `count` mặc định 1 |
| `delete_sheet_columns` | `filepath`, `sheet_name`, `start_col` | `count` mặc định 1 |

### Nhóm 4: Formatting
| Tool | Args bắt buộc | Args tùy chọn |
|------|--------------|--------------|
| `format_range` | `filepath`, `sheet_name`, `start_cell`, `end_cell` | `bold`, `bg_color` (ARGB), `font_color` (ARGB) |
| `merge_cells` | `filepath`, `sheet_name`, `start_cell`, `end_cell` | — |
| `unmerge_cells` | `filepath`, `sheet_name`, `start_cell`, `end_cell` | — |
| `get_merged_cells` | `filepath`, `sheet_name` | — |

### Nhóm 5: Công thức & Validation
| Tool | Args bắt buộc | Ghi chú |
|------|--------------|---------|
| `apply_formula` | `filepath`, `sheet_name`, `cell`, `formula` | `formula` bắt đầu bằng `=` |
| `validate_formula_syntax` | `formula` | Gọi trước `apply_formula` nếu công thức phức tạp |
| `validate_excel_range` | `range_str` | Kiểm tra `"A1:Z100"` có hợp lệ không |
| `get_data_validation_info` | `filepath`, `sheet_name` | Xem validation rules của sheet |

### Nhóm 6: Bảng & Biểu đồ
| Tool | Args bắt buộc | Ghi chú quan trọng |
|------|--------------|-------------------|
| `create_table` | `filepath`, `sheet_name`, `data_range` | `data_range` = `"A1:D10"` (không phải start/end_cell riêng lẻ) |
| `create_chart` | `filepath`, `sheet_name`, `chart_type`, `data_range`, `title` | `chart_type`: `"bar"`, `"line"`, `"pie"`, `"column"` |
| `create_pivot_table` | `filepath`, `data_sheet`, `pivot_sheet`, `rows`, `cols`, `values` | `rows`/`cols`/`values` = tên cột |
| `get_data_validation_info` | `filepath`, `sheet_name` | — |

---

## ⓸ Ví dụ Payload Đầy đủ

### Báo cáo doanh thu Q1 (chuẩn production)

```json
// 1. Tạo workbook
{ "name": "create_workbook", "arguments": { "filepath": "doanh_thu_q1_2026.xlsx" } }

// 2. Ghi dữ liệu — data là list of lists
{ "name": "write_data_to_excel", "arguments": {
    "filepath": "doanh_thu_q1_2026.xlsx",
    "sheet_name": "DuLieu",
    "data": [
      ["Tháng", "Sản phẩm", "Doanh Thu", "Chi Phí", "Lợi Nhuận"],
      ["Tháng 1", "SP-A", 150000000, 80000000, 70000000],
      ["Tháng 2", "SP-A", 180000000, 90000000, 90000000],
      ["Tháng 3", "SP-A", 200000000, 100000000, 100000000]
    ]
}}

// 3. Format header (bold, xanh dương, chữ trắng)
{ "name": "format_range", "arguments": {
    "filepath": "doanh_thu_q1_2026.xlsx", "sheet_name": "DuLieu",
    "start_cell": "A1", "end_cell": "E1",
    "bold": true, "bg_color": "FF4472C4", "font_color": "FFFFFFFF"
}}

// 4. Tạo Table (tự có filter + zebra style)
{ "name": "create_table", "arguments": {
    "filepath": "doanh_thu_q1_2026.xlsx", "sheet_name": "DuLieu",
    "data_range": "A1:E4", "table_name": "DanhSachDoanhThu"
}}

// 5. Công thức tổng
{ "name": "apply_formula", "arguments": {
    "filepath": "doanh_thu_q1_2026.xlsx", "sheet_name": "DuLieu",
    "cell": "C5", "formula": "=SUM(C2:C4)"
}}
{ "name": "apply_formula", "arguments": { ..., "cell": "E5", "formula": "=SUM(E2:E4)" }}

// 6. Tạo biểu đồ cột doanh thu
{ "name": "create_chart", "arguments": {
    "filepath": "doanh_thu_q1_2026.xlsx", "sheet_name": "DuLieu",
    "chart_type": "column",
    "data_range": "A1:C4",
    "title": "Doanh Thu Q1 2026 Theo Tháng"
}}
```

### Pivot Table phân tích theo vùng

```json
// Cần data_sheet có sẵn dữ liệu với các cột: Vùng, Sản phẩm, Doanh Thu
{ "name": "create_pivot_table", "arguments": {
    "filepath": "phan_tich_2026.xlsx",
    "data_sheet": "DuLieu",
    "pivot_sheet": "PivotTongHop",
    "rows": ["Vùng"],
    "cols": ["Sản phẩm"],
    "values": ["Doanh Thu"]
}}
```

### Validate công thức phức tạp trước khi dùng

```json
{ "name": "validate_formula_syntax", "arguments": {
    "formula": "=IF(C2>0,C2/B2*100,0)"
}}
// Nếu trả về valid → mới gọi apply_formula
```

### Tiêu đề span nhiều cột (merge + format)

```json
{ "name": "merge_cells", "arguments": {
    "filepath": "bao_cao.xlsx", "sheet_name": "TongHop",
    "start_cell": "A1", "end_cell": "E1"
}}
{ "name": "format_range", "arguments": {
    "filepath": "bao_cao.xlsx", "sheet_name": "TongHop",
    "start_cell": "A1", "end_cell": "E1",
    "bold": true, "bg_color": "FF1F3864", "font_color": "FFFFFFFF"
}}
// Sau đó ghi text vào A1 bằng write_data_to_excel
```

---

## ⓹ Palette Màu Chuẩn (bg_color — ARGB hex, luôn bắt đầu bằng FF)

| Màu | ARGB Code | Dùng cho |
|-----|-----------|---------|
| Xanh dương đậm (corporate) | `FF1F3864` | Tiêu đề chính / merge header |
| Xanh dương vừa | `FF4472C4` | Header cột chính |
| Xanh lam nhạt | `FF9DC3E6` | Sub-header / header cấp 2 |
| Xanh lá | `FF70AD47` | Giá trị tích cực / đạt target |
| Vàng cam | `FFED7D31` | Cảnh báo / gần target |
| Đỏ nhạt | `FFFF9999` | Lỗi / không đạt (nền nhạt để chữ đọc được) |
| Xám nhạt (zebra) | `FFF2F2F2` | Hàng chẵn trong bảng dữ liệu lớn |
| Vàng nhạt (highlight) | `FFFFFF99` | Ô cần chú ý |
| Trắng | `FFFFFFFF` | Nền mặc định / chữ trên nền tối |
| Đen | `FF000000` | Chữ trên nền vàng/nhạt |

---

## ⓺ Quy tắc Đặt tên File

| Loại báo cáo | Pattern tên file |
|-------------|----------------|
| Báo cáo doanh thu | `doanh_thu_{period}_{YYYYMMDD}.xlsx` |
| Tổng hợp | `tong_hop_{chu_de}_{YYYYMMDD}.xlsx` |
| Phân tích | `phan_tich_{chu_de}_{YYYYMMDD}.xlsx` |
| Export dữ liệu | `export_{table}_{YYYYMMDD}.xlsx` |
| Template | `template_{loai}.xlsx` |

**Không được:** tên có dấu tiếng Việt (`báo cáo.xlsx` ❌), spaces (`bao cao.xlsx` ❌), hay ký tự đặc biệt.

---

## ⓻ Khi nào dùng tool nào?

| Tình huống | Tool ưu tiên | Lý do |
|-----------|-------------|-------|
| Tạo bảng có filter/sort | `create_table` | Tự zebra, tự filter — không cần format từng hàng |
| Tạo bảng tĩnh | `write_data_to_excel` + `format_range` | Kiểm soát hoàn toàn |
| Phân tích đa chiều | `create_pivot_table` | Tổng hợp tự động theo rows/cols |
| Trực quan hóa | `create_chart` | Dùng `"column"` cho dữ liệu theo thời gian, `"pie"` cho tỷ lệ |
| Tiêu đề span | `merge_cells` + `write_data_to_excel` | Gộp trước, ghi sau |
| Nhiều phần báo cáo | Nhiều `create_worksheet` | Mỗi sheet = 1 chủ đề |
| Template sao chép | `copy_worksheet` | Tạo sheet mẫu rồi copy |
| Chèn thêm hàng giữa | `insert_rows` (start_row = số thứ tự hàng) | Không phải row_index |
| Ghi kết quả phân tích | `read_data_from_excel` → xử lý → `write_data_to_excel` | Read-analyze-write pipeline |

---

## ⓼ Lỗi Thường Gặp & Cách Fix

| Lỗi | Nguyên nhân thường gặp | Fix |
|-----|----------------------|-----|
| `File not found` | Chưa `create_workbook` | Gọi `create_workbook` trước |
| `Sheet not found` | Sheet chưa tồn tại | Gọi `create_worksheet` trước |
| `validation error: start_row Field required` | Dùng `row_index` thay vì `start_row` | Dùng đúng: `start_row` |
| `validation error: start_col Field required` | Dùng `col_index` thay vì `start_col` | Dùng đúng: `start_col` |
| `validation error: data_range Field required` | Dùng `start_cell`/`end_cell` cho `create_table` | Dùng đúng: `data_range: "A1:D10"` |
| `source_start Field required` | Dùng `source_range` cho `copy_range` | Dùng đúng: `source_start`, `source_end`, `target_start` |
| `Extension không hỗ trợ` | File không phải `.xlsx` | Đổi tên sang `.xlsx` |
| `data phải là list of lists` | Gửi list of dicts | Chuyển về `[[h1,h2],[v1,v2]]` |
| `End row X out of bounds` | `delete_range` trỏ quá số hàng thực tế | Kiểm tra `get_workbook_metadata` trước |

---

## ⓽ Upload & Download

```
# Upload file Excel có sẵn vào server
POST http://localhost:5003/upload
  → field: "file" (multipart), query: ?filename=ten_file.xlsx&overwrite=true

# Download file đã tạo
GET http://localhost:5003/download/<ten_file.xlsx>
  → trả về binary stream với Content-Disposition: attachment

# Xem danh sách file
GET http://localhost:5003/files

# Xem danh sách 25 tools
GET http://localhost:5003/tools

# Xóa file
DELETE http://localhost:5003/files/<ten_file.xlsx>
```

---

## ⓾ Ví dụ Prompt → Pipeline đầy đủ

**User**: *"Tạo báo cáo tổng hợp nhân sự gồm danh sách nhân viên, bảng lương, và biểu đồ phân bổ theo phòng ban"*

```
Pipeline AI nên thực hiện:
1. create_workbook { filepath: "bao_cao_nhan_su_20260427.xlsx" }

2. [Sheet 1] Danh sách nhân viên
   create_worksheet { sheet_name: "DanhSach" }  ← Sheet1 mặc định đã có
   write_data_to_excel { sheet_name: "DanhSach", data: [[Mã NV, Họ Tên, Phòng Ban, Chức Vụ, Ngày Vào], ...] }
   format_range header → blue
   create_table { data_range: "A1:E{n}", table_name: "DanhSachNV" }

3. [Sheet 2] Bảng lương
   create_worksheet { sheet_name: "BangLuong" }
   write_data_to_excel { sheet_name: "BangLuong", data: [[Mã NV, Lương Cơ Bản, Phụ Cấp, Thưởng, Tổng], ...] }
   format_range header → green
   apply_formula tổng từng cột
   create_table { data_range: "A1:E{n}", table_name: "BangLuong" }

4. [Sheet 3] Biểu đồ phòng ban
   create_worksheet { sheet_name: "PhanBo" }
   write_data_to_excel { data: [[Phòng Ban, Số NV], [IT, 12], [HR, 5], ...] }
   create_chart { chart_type: "pie", data_range: "A1:B{n}", title: "Phân Bổ Nhân Sự Theo Phòng Ban" }

5. Thông báo: "Đã tạo file bao_cao_nhan_su_20260427.xlsx — download tại /download/bao_cao_nhan_su_20260427.xlsx"
```
