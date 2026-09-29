# 📘 Hướng Dẫn Sử Dụng MetadataCompiler V1

> Dành cho người mới — hướng dẫn từng bước từ A đến Z.

---

## 📌 Mục Lục

1. [Tổng Quan](#1-tổng-quan)
2. [Cấu Trúc Thư Mục](#2-cấu-trúc-thư-mục)
3. [Yêu Cầu Cài Đặt](#3-yêu-cầu-cài-đặt)
4. [Bước 1: Chuẩn Bị Input](#4-bước-1-chuẩn-bị-input)
5. [Bước 2: Chạy SCAN — Đọc File Excel](#5-bước-2-chạy-scan--đọc-file-excel)
6. [Bước 3: Chạy COMPILE — Sinh Mã VBA](#6-bước-3-chạy-compile--sinh-mã-vba)
7. [Bước 4 (Tùy Chọn): Dùng AI Agent](#7-bước-4-tùy-chọn-dùng-ai-agent)
8. [Kết Quả Output](#8-kết-quả-output)
9. [Xử Lý Lỗi Thường Gặp](#9-xử-lý-lỗi-thường-gặp)
10. [Quy Trình Đầy Đủ Tóm Tắt](#10-quy-trình-đầy-đủ-tóm-tắt)

---

## 1. Tổng Quan

**MetadataCompiler V1** là công cụ tự động:

1. **SCAN** — Đọc file Excel (`.xlsx`) chứa metadata, trích xuất dữ liệu ra file JSON.
2. **COMPILE** — Đọc file JSON vừa xuất, sinh ra mã nguồn VBA (`.bas`) tự động.

```
File Excel (.xlsx)
       ↓  [SCAN]
  sheet_raw.json
       ↓  [COMPILE]
  output.bas (VBA code)
```

---

## 2. Cấu Trúc Thư Mục

```
MetadataCompiler_V1/
├── input/                  ← Chứa các folder input của bạn
│   └── test7/              ← Ví dụ: folder tên "test7"
│       └── update.xlsx     ← File Excel đặt vào đây
│
├── output/                 ← Kết quả được sinh ra ở đây (tự động tạo)
│   └── test7/
│       └── <tên_file>/
│           └── <tên_sheet>/
│               ├── sheet_raw.json   ← Kết quả SCAN
│               └── output.bas       ← Kết quả COMPILE (nếu có)
│
├── config/
│   ├── scanner/
│   │   └── scan.py         ← Script chạy SCAN
│   └── compiler/
│       └── vba_compiler.py ← Script chạy COMPILE
│
├── .agents/                ← Workspace Customizations (skills, rules, workflows, AGENTS.md)
│   ├── skills/             ← Thư viện kỹ năng (vba-access-architect, senior-ba-metadata-compiler...)
│   ├── rules/              ← Quy tắc hệ thống (vba-rule.md)
│   └── workflows/          ← Tài liệu quy trình làm việc
├── agents                  ← Symlink tương thích ngược trỏ đến .agents
├── sample/                 ← Từ điển ngữ nghĩa (semantic_dictionary.json)
└── HUONG_DAN.md            ← File này
```

---

## 3. Yêu Cầu Cài Đặt

### 3.1 Python

Cần Python 3.10 trở lên. Kiểm tra:

```bash
python3 --version
```

### 3.2 Thư Viện Python

```bash
pip install openpyxl
```

> Nếu dùng tính năng AI Agent (Bước 4), cần thêm:
> ```bash
> pip install google-generativeai
> ```

### 3.3 Mở Terminal tại đúng thư mục

**Quan trọng:** Tất cả lệnh đều phải chạy từ thư mục gốc của project:

```bash
cd /home/huyhg/Documents/MetadataCompiler_V1
```

Kiểm tra bạn đang ở đúng chỗ:

```bash
ls
# Nên thấy: .agents/  config/  input/  output/  agents  sample/  HUONG_DAN.md
```

---

## 4. Bước 1: Chuẩn Bị Input

### 4.1 Tạo folder input

Đặt file Excel của bạn vào một folder con trong `input/`. Tên folder sẽ dùng làm tham số lệnh.

**Ví dụ:** Tạo folder `myproject` và đặt file Excel vào:

```
input/
└── myproject/
    └── data.xlsx
```

> 💡 **Lưu ý:** File Excel phải có ít nhất 1 sheet chứa đủ 4 cột tiêu đề:
> - `ファイル名` (Tên file)
> - `出力条件` (Điều kiện xuất)
> - `通常処理` (Xử lý thường)
> - `特殊処理` (Xử lý đặc biệt)
>
> Các sheet không có đủ 4 cột này sẽ bị bỏ qua.

---

## 5. Bước 2: Chạy SCAN — Đọc File Excel

### 5.1 Cú pháp lệnh

```bash
Usage:
    python3 main.py scan    --test2          # Scan xlsx → sheet_raw.json
    python3 main.py compile --test2          # Compile sheet_raw.json → .bas
    python3 main.py run     --test2          # scan + compile (all-in-one)
    python3 main.py learn   --test1          # Học snippet tự động từ file .bas
```

Thay `<tên_folder>` bằng tên folder bạn tạo trong `input/`.

### 5.2 Ví dụ thực tế

**Ví dụ 1:** Scan folder `test7` (đã có sẵn trong project):

```bash
python3  main.py scan --test7
```

### 5.3 Kết quả khi chạy thành công

**Lưu ý** : Thông báo sẽ xuất hiện những Sheet chỉ thiếu 1 hoặc 2 khóa bị bỏ qua. Hãy kiểm tra lại sheet đó và bổ sung các khóa còn thiếu rồi chạy lại compile nếu muốn sử dụng các thông tin đó.

```
════════════════════════════════════════════════════════════
  SCAN  →  input/test7
════════════════════════════════════════════════════════════
  📂  Input  : /path/to/MetadataCompiler_V1/input/test7
  📁  Output : /path/to/MetadataCompiler_V1/output/test7
  📊  Tìm thấy 1 file xlsx

────────────────────────────────────────────────────────────
📄  File: update.xlsx
  ✅  15 sheet hợp lệ sẽ được xuất:
        ✔  'アカウント（オーナー）'  →  output/test7/update/アカウント（オーナー）/sheet_raw.json
        ✔  '建物'                   →  output/test7/update/建物/sheet_raw.json
        ...

════════════════════════════════════════════════════════════
  KẾT QUẢ
════════════════════════════════════════════════════════════
  ✅  Đã xuất  : 15 sheet
  ⏭️   Bỏ qua  : 2 sheet
```

Sau khi SCAN xong, bạn sẽ thấy các file `sheet_raw.json` xuất hiện trong `output/<tên_folder>/`.

---

## 6. Bước 3: Chạy COMPILE — Sinh Mã VBA

Sau khi SCAN xong, dùng lệnh COMPILE để sinh mã VBA từ các file JSON.


## 7. Bước 4 (Tùy Chọn): Dùng AI Agent

Nếu compiler thông thường không xử lý được một số trường hợp phức tạp, bạn có thể dùng AI Agent (Google Gemini).

### 7.1 Cài đặt API Key

```bash
export GOOGLE_API_KEY="your-api-key-here"
```

> Lấy API Key tại: https://aistudio.google.com/app/apikey

### 7.2 Chọn model (tùy chọn)

Mặc định dùng `gemini-2.5-pro`. Có thể đổi:

```bash
export GEMINI_MODEL="gemini-2.5-flash"
```



---

## 8. Kết Quả Output

### 8.1 Cấu trúc file JSON (sheet_raw.json)

Sau khi SCAN, file JSON có dạng:

```json
{
  "_meta": {
    "generated_at": "2026-06-26T08:00:00",
    "source_file": "update.xlsx",
    "source_sheet": "建物",
    "folder": "test7"
  },
  "sections": {
    "ファイル名": [...],
    "出力条件": {...},
    "通常処理": {...},
    "特殊処理": {...}
  }
}
```

### 8.2 Vị trí file output

```
output/
└── test7/
    └── update/               ← Tên file Excel (không có .xlsx)
        ├── 建物/             ← Tên sheet
        │   ├── sheet_raw.json
        │   └── output.bas    ← File VBA được sinh ra
        └── 契約/
            ├── sheet_raw.json
            └── output.bas
```

---

### ⚠️ Sheet bị bỏ qua (skipped)

Nếu thấy output:
```
⏭️   Bỏ qua X sheet (thiếu keyword)
```

**Nguyên nhân:** Sheet trong Excel không có đủ 4 cột tiêu đề bắt buộc.

**Cách kiểm tra:** Mở file Excel, xem sheet đó có các hàng tiêu đề:
- `ファイル名`
- `出力条件`
- `通常処理`
- `特殊処理`

---

## 10. Quy Trình Đầy Đủ Tóm Tắt

```bash
# 0. Chuyển đến thư mục project
cd /home/huyhg/Documents/MetadataCompiler_V1

# 1. Tạo folder và đặt file Excel vào
mkdir -p input/myproject
cp /path/to/your/file.xlsx input/myproject/

# 2. Chạy SCAN
python3 config/scanner/scan.py --myproject

# 3. Xem kết quả JSON
ls output/myproject/

# 4. Chạy COMPILE cho từng sheet
python3 config/compiler/vba_compiler.py output/myproject/<tên_file>/<tên_sheet>/sheet_raw.json

# 5. (Tùy chọn) Compile toàn bộ
find output/myproject -name "sheet_raw.json" | while read f; do
    python3 config/compiler/vba_compiler.py "$f"
done
```

---

## 📎 Tham Khảo Thêm

| File | Mô tả |
|------|-------|
| `config/scanner/scan.py` | Script SCAN chính |
| `config/compiler/vba_compiler.py` | Script COMPILE chính |
| `config/compiler/agent_compiler.py` | Agent AI (dùng Gemini) |
| `.agents/workflows/vba_generation_workflow.md` | Workflow sinh VBA cho AI |
| `.agents/rules/vba-rule.md` | Quy tắc viết VBA |
| `.agents/skills/vba-access-architect/SKILL.md` | Kiến trúc Access VBA chi tiết |
| `sample/semantic_dictionary.json` | Từ điển ngữ nghĩa dùng khi compile |
| `verify_report.md` | Báo cáo kiểm tra kết quả scan |

---

> 📝 **Ghi chú:** Tool này chạy hoàn toàn offline (không cần API key) ở bước SCAN và COMPILE cơ bản. Chỉ cần API key khi dùng AI Agent (Bước 4) và sau khi Compile chạy (sẽ có sai sót). Ta sử dụng câu Prompt (Terminal sẽ sinh ra) --> chạy trên Antigravity để Agent chỉnh sửa online.

> 📝 **Quy Trình Chuẩn:**  

**Bước 1**: Tạo folder và thêm file xlsx vào input/<tên folder>/<file xlsx>

**Bước 2**: Chạy lệnh Scan - python3 main.py scan --<tên folder> 

**Bước 3**: Check kết quả terminal để xem các sheet nào thiếu 1/2 khóa để kiểm tra bổ sung.

**Bước 4**: Nếu cần thêm thông tin hoặc chỉnh sửa, chạy lệnh Compile - python3 main.py compile --<tên folder>

**Bước 5**: Copy Prompt sinh ra từ terminal -> paste vào Antigravity để chỉnh sửa online, 

**Bước 6**: Check lai file sinh ra trong output và so sánh với spec_raw.json.

