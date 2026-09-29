---
name: senior-ba-metadata-compiler
version: 1.0.0
last_updated: 2026-05-28
description: >
  Bộ kỹ năng toàn diện cho Senior BA làm việc trong pipeline MetadataCompiler.
  Trigger khi: review spec Excel JP, phân loại token semantic dictionary,
  khai báo job_metadata, phân tích điều kiện SQL phức hợp, hoặc bất kỳ
  tác vụ nào liên quan đến chuyển đổi nghiệp vụ JP → AST → VBA/SQL.
---

# Senior BA Skill — MetadataCompiler Pipeline
> Domain: Japanese Property Management · Stack: MS Access / VBA / SQL / Python DSL

---

## PHẦN I — TRIẾT LÝ LÀM VIỆC

### 1.1 BA trong Pipeline này là gì?

BA **không phải** người viết code. BA là **người biên dịch ngôn ngữ nghiệp vụ** (tiếng Nhật, Excel spec, quy tắc kinh doanh) sang **ngôn ngữ máy hiểu được** (JSON DSL, semantic token, AST config).

Một quyết định sai của BA ở `semantic_dictionary.json` → lỗi lan tất cả các file VBA sinh ra từ token đó. Ngược lại, một BA phân loại đúng → pipeline chạy trơn không cần sửa code.

### 1.2 Nguyên tắc vàng

| # | Nguyên tắc | Áp dụng thực tế |
|---|---|---|
| G1 | **Explicit hơn Implicit** | Đừng để điều kiện dạng văn xuôi trong Excel nếu có thể tách ra thành các ô riêng |
| G2 | **Phân rã trước, gom sau** | Một điều kiện phức hợp AND/OR phải được tách ra các cột `条件_1`, `条件_2` trước khi phân loại |
| G3 | **Config là source of truth** | Mọi thay đổi nghiệp vụ phải đi qua config JSON, không email, không verbal |
| G4 | **Audit mọi quyết định** | Mọi entry trong `semantic_dictionary.json` phải có `classified_by`, `notes` giải thích lý do |
| G5 | **Fail sớm, fail rõ** | Nếu một token chưa rõ → đánh `UNKNOWN`, không đoán mò → pipeline sẽ báo lỗi sớm |

---

## PHẦN II — NĂNG LỰC ĐỌC EXCEL SPEC TIẾNG NHẬT

### 2.1 Cấu trúc Sheet nghiệp vụ chuẩn

Một worksheet hợp lệ trong MetadataCompiler gồm **3 section** được nhận diện qua tên hàng tiêu đề:

```
┌─────────────────────────────────┐
│  出力条件 (Output Conditions)    │  ← Quy tắc lọc & dedup
├─────────────────────────────────┤
│  通常処理 (Normal Processing)    │  ← Mapping 1-1 field thông thường
├─────────────────────────────────┤
│  特殊処理 (Special Processing)  │  ← Xử lý có điều kiện phức tạp
└─────────────────────────────────┘
```

**Kỹ năng bắt buộc**: Phân biệt rõ 3 section này trong mọi sheet khách hàng cung cấp,
kể cả khi header bị merge cell, đặt sai vị trí, hoặc có tên biến thể (略称).

### 2.2 Bảng Cột Quan Trọng Theo Section

#### 出力条件 — Output Conditions
| Cột JP | Alias Python | Ý nghĩa BA cần nắm |
|--------|-------------|-------------------|
| 開発用チェック | `dev_check` | TRUE = hàng này được xử lý. FALSE/trống = bị bỏ qua HOÀN TOÀN → **luôn xác nhận với Dev trước khi để False** |
| 処理 | `action` | Loại xử lý: `BLANK_CHECK`, `EXISTS_JOIN`, `DEDUPLICATE` — BA phân loại thành `ast_type` |
| 条件 | `condition` | Điều kiện lọc. Phải là biểu thức SQL-like, **không được là văn xuôi** |
| 該当項目名_1..4 | `source_field_1..4` | Trường nguồn tương ứng |

#### 通常処理 — Normal Processing
| Cột JP | Alias Python | Ý nghĩa BA cần nắm |
|--------|-------------|-------------------|
| DB_Field | `db_field` | Tên trường trong Access DB (target) — phải khớp chính xác với schema DB |
| DB_Name | `db_name` | Tên hiển thị tiếng Nhật — chỉ để tham chiếu, không đưa vào SQL |
| 処理 | `action` | `固定値`, `直接マッピング`, `結合` — xác định `ast_type` |
| 項目名 / 出力内容_1..3 | `source_val_1..3` | Giá trị nguồn hoặc biểu thức |

#### 特殊処理 — Special Processing
| Cột JP | Alias Python | Ý nghĩa BA cần nắm |
|--------|-------------|-------------------|
| 項目名 | `target_field` | Trường đích. **Nhiều hàng cùng `項目名` = nhiều branch của 1 SWITCH** |
| 処理 | `action` | Phải là token đã có trong `semantic_dictionary.json` |
| 条件 / 項目マッピング | `condition_marker` | Điều kiện kích hoạt branch này — **PHẢI là biểu thức tường minh** |
| 該当項目名_1 | `source_field_1` | Trường nguồn cho branch này |

### 2.3 Những Bẫy Thường Gặp Khi Đọc Excel KH

**Bẫy 1 — Điều kiện văn xuôi trong `条件`:**
```
❌ BAD:  "個人区分が'個人'かつ入居有無が'入居有り'の場合"
✅ GOOD: "T1.[個人・法人区分] = '個人' AND T1.[入居有無] = '入居有り'"
```
→ BA phải yêu cầu KH cung cấp lại, hoặc tự tách ra trước khi phân loại.

**Bẫy 2 — Merge cell ẩn hàng:**
Header section đôi khi bị merge nhiều cột → `excel_parser` có thể miss section boundary.
→ BA phải kiểm tra file Excel gốc bằng mắt trước khi chạy scan.

**Bẫy 3 — Cùng `項目名` nhưng khác logic:**
Nếu `特殊処理` có 5 hàng cùng `項目名 = "end_date"` → đây là 5 branch của 1 SWITCH.
Không phải 5 trường khác nhau. BA phải đảm bảo mỗi hàng có `condition_marker` riêng biệt.

**Bẫy 4 — `固定値` xuất hiện ở 2 section khác nhau:**
- `固定値` trong `通常処理` → `ast_type: CONSTANT_MAPPING` (INSERT luôn giá trị cố định)
- `固定値` trong `特殊処理` → `ast_type: CONDITIONAL_BRANCH` (SET field = value NẾU điều kiện đúng)

---

## PHẦN III — PHÂN LOẠI TOKEN TRONG SEMANTIC DICTIONARY

### 3.1 Quy Trình Phân Loại Chuẩn (Standard Classification Flow)

```
Bước 1: Mở global_token_candidates.json
         → Lọc các entry có classified_by = "AUTO" hoặc count >= threshold

Bước 2: Với mỗi token JP chưa phân loại:
         a. Tra context: source_sheets + found_in_keys
         b. Xác định role (xem bảng 3.2)
         c. Xác định ast_type (xem bảng 3.3)
         d. Ghi notes giải thích lý do
         e. Đổi classified_by = "MANUAL"

Bước 3: Chạy excel_sync.py để sync về Excel cho BA khác review
Bước 4: Chạy validator.py để đảm bảo schema hợp lệ
```

### 3.2 Role Taxonomy — Bảng Phân Loại Vai Trò

| Role | Định nghĩa | Dấu hiệu nhận biết | Ví dụ token JP |
|------|-----------|-------------------|---------------|
| `OPERATION_VERB` | Động từ hành động chính | Xuất hiện trong cột `処理`, quyết định loại SQL | `更新する`, `削除`, `固定値`, `コピー` |
| `CONDITION_MARKER` | Từ/cụm từ đánh dấu điều kiện | Xuất hiện trong cột `条件`, `条件/項目マッピング` | `の場合`, `かつ`, `または`, `～以外` |
| `FILTER_TYPE` | Loại bộ lọc dữ liệu | Trong `出力条件`, xác định kiểu WHERE clause | `ブランクチェック`, `重複排除`, `存在チェック` |
| `DB_FIELD_NAME` | Tên trường DB thực tế | Trong cột `DB_Field` — dùng trực tiếp trong SQL | `メールアドレス`, `契約開始日` |
| `DB_DISPLAY_NAME` | Tên hiển thị JP | Trong cột `DB_Name` — KHÔNG dùng trong SQL | `メールアドレス表示名` |
| `SOURCE_FILE` | Tên file/bảng nguồn | Tham chiếu đến source CSV/table | `入居状況一覧`, `解約一覧` |
| `SOURCE_FIELD` | Tên trường trong nguồn | Trong `該当項目名`, `項目名/出力内容` | `入居者氏名`, `契約番号` |
| `CONSTANT_VALUE` | Giá trị literal cố định | Giá trị chuỗi/số không phải tên trường | `'個人'`, `0`, `NULL`, `'-'` |
| `IGNORE` | Token không có semantic | Từ formatting, header phụ, ký tự đặc biệt | `※`, `▼`, `---` |
| `UNKNOWN` | Chưa xác định | Khi BA không chắc chắn → **bắt buộc ghi notes** | — |

### 3.3 AST Type Taxonomy — Bảng Loại Node AST

| AST Type | Sinh ra VBA pattern gì | Điều kiện áp dụng |
|----------|----------------------|------------------|
| `DIRECT_MAPPING` | `INSERT INTO T ([tgt]) VALUES (T1.[src])` | 1 source → 1 target, không điều kiện |
| `CONSTANT_MAPPING` | `INSERT INTO T ([tgt]) VALUES ('value')` | Giá trị cố định |
| `CONDITIONAL_MAPPING` | `UPDATE T SET [tgt] = SWITCH(...)` | Nhiều branch theo điều kiện |
| `HYPHEN_TO_NULL` | `IIF(T1.[f]='-', NULL, T1.[f])` | Source có thể là dấu gạch ngang `'-'` |
| `CONCAT_WITH_HYPHEN` | `T1.[f1] & '-' & T1.[f2]` | Ghép 2 trường bằng dấu gạch |
| `DEDUPLICATE_RULE` | `SELECT MIN(ID) GROUP BY key` | Cuối section 出力条件, loại trùng |
| `DELETE_RECORD` | `DELETE FROM T WHERE condition` | Xóa record theo điều kiện |
| `FILTER_RULE` | `WHERE T1.[f] IS NOT NULL` / `DELETE` | Lọc đầu vào |
| `BRANCH_CONDITION` | Một nhánh trong SWITCH block | Một hàng trong khối 特殊処理 cùng 項目名 |
| `LIKE_PATTERN` | `WHERE T1.[f] LIKE 'pattern%'` | Điều kiện mờ |
| `NOT_LIKE_PATTERN` | `WHERE T1.[f] NOT LIKE 'pattern%'` | Loại trừ pattern |
| `DATE_FORMAT` | `Format(T1.[f], 'yyyy/mm/dd')` | Chuẩn hóa ngày tháng |
| `REMOVE_CHAR_OUTPUT` | `Replace(T1.[f], 'char', '')` | Xóa ký tự khỏi output |

### 3.4 Ma Trận Quyết Định: Section × Token → AST Type

```
                  │ OPERATION_VERB  │ FILTER_TYPE     │ CONDITION_MARKER
──────────────────┼─────────────────┼─────────────────┼─────────────────
出力条件          │ → FILTER_RULE   │ → FILTER_RULE   │ N/A
通常処理          │ → DIRECT/CONST  │ N/A             │ N/A
特殊処理          │ → CONDITIONAL   │ N/A             │ → BRANCH_CONDITION
                  │   _MAPPING      │                 │
```

---

## PHẦN IV — KHAI BÁO JOB METADATA

### 4.1 Schema job_metadata.json (đầy đủ với chú thích)

```jsonc
{
  "jobs": [
    {
      "job_id": "account",               // Unique ID, dùng làm tên folder output
      "job_name": "アカウント (Account)", // Tên hiển thị cho BA
      "target_table": "Account",         // Tên bảng đích trong Access DB

      // OPTIONAL — chỉ điền khi job cần bảng phụ
      "helper_template": "templates/contract_helpers.bas",
      "extra_tables": [
        {
          "alias": "T2",                 // Alias dùng trong SQL
          "table_name": "解約一覧",       // Tên bảng phụ
          "join_type": "INNER JOIN",     // INNER JOIN | LEFT JOIN
          "join_key_target": "契約番号",  // Khóa phía bảng đích
          "join_key_source": "契約番号"   // Khóa phía bảng phụ
        }
      ],

      "sheets": [
        {
          "sheet_name": "アカウント（オーナー）_O",  // Tên tab sheet CHÍNH XÁC
          "source_table": "入居状況一覧",            // Bảng nguồn chính (alias T1)
          "id_prefix": "O_",                        // Prefix cho ID sinh ra
          "dev_check_override": false               // true = ignore dev_check column
        }
      ]
    }
  ]
}
```

### 4.2 Checklist Trước Khi Submit job_metadata.json

```
□ job_id không có dấu cách, không có ký tự JP
□ target_table khớp chính xác tên bảng trong Access DB (case-sensitive)
□ sheet_name khớp chính xác tên tab trong file Excel (kể cả ký tự toàn/bán giác)
□ source_table tồn tại trong danh sách bảng nguồn được phê duyệt
□ id_prefix unique trong toàn project (không trùng giữa các sheets)
□ extra_tables chỉ khai báo khi sheet có JOIN nhiều bảng trong 特殊処理
□ join_key đã được xác nhận với DBA là Foreign Key đúng
```

---

## PHẦN V — PHÂN TÍCH ĐIỀU KIỆN SQL PHỨC HỢP

### 5.1 Kỹ Năng Tách Điều Kiện Văn Xuôi JP → SQL Tường Minh

Đây là kỹ năng quan trọng nhất để đưa pipeline từ 80% → 95% accuracy.

**Pattern nhận diện câu điều kiện JP:**

| Pattern JP | SQL tương đương |
|-----------|----------------|
| `AがBの場合` | `WHERE T1.[A] = 'B'` |
| `AがBかつCがDの場合` | `WHERE T1.[A] = 'B' AND T1.[C] = 'D'` |
| `AがBまたはCの場合` | `WHERE T1.[A] = 'B' OR T1.[A] = 'C'` |
| `AがB以外の場合` | `WHERE T1.[A] <> 'B'` |
| `Aが空白の場合` | `WHERE T1.[A] IS NULL OR T1.[A] = ''` |
| `Aが空白以外の場合` | `WHERE T1.[A] IS NOT NULL AND T1.[A] <> ''` |
| `AにBを含む場合` | `WHERE T1.[A] LIKE '*B*'` |
| `AがB〜Cの場合` | `WHERE T1.[A] >= 'B' AND T1.[A] <= 'C'` |

### 5.2 Quy Trình Tách Điều Kiện Phức Hợp

Với điều kiện như:
> `"個人・法人区分_33が'個人'で、かつ契約者1入居有無が'入居有り'の場合"`

Tách thành:
```
条件_1: T1.[個人・法人区分_33] = '個人'
条件_2: T1.[契約者1入居有無] = '入居有り'
結合:   AND
```

Sau đó điền vào Excel spec với 2 cột riêng biệt, không gộp văn xuôi.

### 5.3 Nhận Diện Cấu Trúc SWITCH vs IF-ELSE-IF

**Dùng SWITCH (CONDITIONAL_MAPPING)** khi:
- Cùng 1 `項目名` có N hàng với N điều kiện khác nhau
- Các điều kiện mutually exclusive (không overlap)
- Default value có thể là `NULL` hoặc giá trị cố định

```sql
-- Output mong đợi từ SWITCH block:
UPDATE T SET T.[end_date] = SWITCH(
    T1.[区分] = '個人' AND T1.[入居有無] = '入居有り', T1.[退去日_33],
    T1.[区分] = '法人', T1.[退去日_43],
    True, NULL
)
WHERE ...
```

**Dùng IF-ELSE logic (BRANCH_CONDITION riêng biệt)** khi:
- Các điều kiện có thể overlap
- Mỗi branch cần UPDATE một tập trường khác nhau
- Thứ tự xử lý quan trọng

### 5.4 Multi-Table JOIN — Cách BA Cung Cấp Thông Tin

Khi 特殊処理 cần JOIN nhiều bảng, BA phải ghi rõ trong sheet:

```
項目名:          end_date
処理:            更新 (T2参照)        ← ghi rõ tên alias bảng phụ
条件/項目マッピング: T1.契約番号 = T2.契約番号 AND T1.区分 = '個人'
該当項目名_1:    T2.解約日             ← prefix bảng phụ bằng alias
```

---

## PHẦN VI — QUY TRÌNH KIỂM TRA CHẤT LƯỢNG (QA)

### 6.1 BA Self-Review Checklist (trước khi push config)

```
SEMANTIC DICTIONARY
□ Tất cả token count >= 5 đã được classify MANUAL
□ Không có token nào role = UNKNOWN mà không có notes
□ Không có token nào vừa là DB_FIELD_NAME vừa là SOURCE_FIELD
□ Tất cả OPERATION_VERB đã có ast_type tương ứng
□ validator.py chạy không có error

JOB METADATA
□ Tất cả sheet_name trong job_metadata tồn tại trong file Excel tương ứng
□ extra_tables chỉ khai báo khi có bằng chứng trong 特殊処理
□ Tất cả join_key đã confirm với DBA
□ dev_check_override không để True trừ khi có approval từ Lead

EXCEL SPEC
□ Không có ô 条件 nào chứa văn xuôi JP thuần
□ Mọi 特殊処理 block cùng 項目名 đều có condition_marker riêng biệt
□ Tất cả cột 開発用チェック đã được review (không để trống ngầm hiểu là False)
```

### 6.2 Cross-Check Với Output VBA Sinh Ra

Sau khi pipeline chạy, BA phải kiểm tra xác suất lấy mẫu (ít nhất 20% file output):

```
Bước 1: Đọc generated_vba.bas tìm comment "MANUAL: No source_expr"
         → Đây là các trường pipeline không tìm được mapping → BA thiếu config

Bước 2: Đếm số lệnh UPDATE trong output vs số hàng 特殊処理 trong Excel
         → Phải bằng nhau (mỗi 項目名 unique = 1 UPDATE block)

Bước 3: Kiểm tra mọi INSERT INTO có đủ số cột vs schema bảng đích
         → Thiếu cột thường do dev_check = False hoặc token IGNORE sai

Bước 4: Với contract files — kiểm tra multi-join block
         → Nếu JOIN chỉ có T và T1 trong khi spec yêu cầu T2 → extra_tables thiếu
```

### 6.3 Audit Trail Khi Phát Hiện Lỗi

Khi phát hiện discrepancy giữa output VBA và spec Excel, BA phải ghi:

```jsonc
// Trong semantic_dictionary.json entry liên quan:
{
  "notes": "[2026-05-28] Sai: pipeline gen CONSTANT_MAPPING thay vì CONDITIONAL_MAPPING.
            Nguyên nhân: token '固定値' trong 特殊処理 bị classify cùng rule với 通常処理.
            Fix: tách thành 2 entry theo found_in_keys context.
            Fix verified bởi: [tên BA]"
}
```

---

## PHẦN VII — DOMAIN KNOWLEDGE: PROPERTY MANAGEMENT JP

### 7.1 Các Bảng Nguồn Cốt Lõi

| Tên Bảng JP | Ý nghĩa | Trường Key |
|------------|---------|-----------|
| 入居状況一覧 | Danh sách trạng thái cư trú | 契約番号, 入居者ID |
| 解約一覧 | Danh sách hủy hợp đồng | 契約番号, 解約日 |
| オーナー情報 | Thông tin chủ nhà | オーナーID, 物件ID |
| 物件情報 | Thông tin bất động sản | 物件ID, 部屋番号 |

### 7.2 Phân Biệt Các Loại Hợp Đồng (区分)

| 区分 | Ý nghĩa | Ảnh hưởng đến logic |
|-----|---------|-------------------|
| `個人` | Cá nhân thuê | Thường có trường 入居有無 |
| `法人` | Doanh nghiệp thuê | Không có 入居有無, có 担当者名 |
| `オーナー` (O_) | Chủ sở hữu | id_prefix = "O_", table khác |

### 7.3 Timestamp & Date Format Chuẩn

```
yyyy/mm/dd      ← Access standard date format
yyyymmdd        ← Source CSV format (thường gặp)
yyyy年mm月dd日  ← Display format JP (KHÔNG dùng trong SQL)
```

Khi thấy ngày JP trong source → ast_type phải là `DATE_FORMAT`.

---

## PHẦN VIII — ESCALATION & COMMUNICATION

### 8.1 Khi Nào BA Phải Escalate Lên Lead Dev

```
ESCALATE ngay khi:
- Token xuất hiện > 50 lần nhưng không thể xác định role rõ ràng
- extra_tables cần > 2 bảng phụ (Multi-join 3 chiều)
- Phát hiện điều kiện vòng lặp (A phụ thuộc B, B phụ thuộc A)
- Sheet có cấu trúc section không theo chuẩn (thiếu 出力条件 hoặc 通常処理)
- Output VBA có < 60% cột so với schema DB đích
```

### 8.2 Cách Viết Yêu Cầu Làm Rõ Cho KH

Khi cần KH làm rõ điều kiện văn xuôi, email template:

```
件名: [MetadataCompiler] シート「{sheet_name}」の条件確認依頼

{sheet_name}シートの特殊処理ブロック「{項目名}」の条件について
確認をお願いいたします。

現在の記載:
  条件: {văn xuôi hiện tại}

以下の形式での記載をお願いできますでしょうか:
  条件1: [フィールド名] = '値'
  条件2: [フィールド名] = '値'
  結合:   AND / OR
```

---

## PHẦN IX — QUICK REFERENCE CARDS

### 9.1 Token Phân Loại Nhanh (Top 20 Tokens JP Thường Gặp)

| Token JP | Role | AST Type | Ghi chú |
|---------|------|----------|---------|
| `固定値` | OPERATION_VERB | CONSTANT_MAPPING / BRANCH_CONDITION | Context-dependent! |
| `直接マッピング` | OPERATION_VERB | DIRECT_MAPPING | |
| `削除` | OPERATION_VERB | DELETE_RECORD | |
| `重複排除` | FILTER_TYPE | DEDUPLICATE_RULE | Luôn cuối 出力条件 |
| `ブランクチェック` | FILTER_TYPE | FILTER_RULE | |
| `存在チェック` | FILTER_TYPE | FILTER_RULE | JOIN-based check |
| `の場合` | CONDITION_MARKER | BRANCH_CONDITION | Phải có tên field trước |
| `かつ` | CONDITION_MARKER | — | Nối 2 điều kiện AND |
| `または` | CONDITION_MARKER | — | Nối 2 điều kiện OR |
| `以外` | CONDITION_MARKER | — | Toán tử NOT / <> |
| `～を含む` | CONDITION_MARKER | LIKE_PATTERN | |
| `結合` | OPERATION_VERB | CONCAT_WITH_HYPHEN | |
| `ハイフンをNull変換` | OPERATION_VERB | HYPHEN_TO_NULL | |
| `日付フォーマット` | OPERATION_VERB | DATE_FORMAT | |
| `文字削除` | OPERATION_VERB | REMOVE_CHAR_OUTPUT | |
| `更新` | OPERATION_VERB | CONDITIONAL_MAPPING | Trong 特殊処理 |
| `コピー` | OPERATION_VERB | DIRECT_MAPPING | Alias của 直接マッピング |
| `空白` | CONSTANT_VALUE | — | Ánh xạ tới NULL |
| `入居有り` | CONSTANT_VALUE | — | Giá trị so sánh |
| `個人` | CONSTANT_VALUE | — | Giá trị so sánh 区分 |

### 9.2 Decision Tree — Khi Nào Dùng AST Type Nào?

```
                  Trong section nào?
                  /               \
          通常処理              特殊処理
          /    \                /      \
    有条件?    無条件      N hàng       1 hàng
      |           |       cùng項目名?     |
  CONDITIONAL   DIRECT     |          CONDITIONAL
  _MAPPING    _MAPPING    YES          _MAPPING
  (SWITCH)   /CONSTANT   /  \         (single)
             _MAPPING  SWITCH BRANCH
                        _ALL _CONDITION
```

### 9.3 File Cần Verify Trước Mỗi Sprint

```
Priority 1 — Mỗi ngày:
  semantic_dictionary.json (count UNKNOWN entries)
  → Target: 0 UNKNOWN với count >= 5

Priority 2 — Trước mỗi compile run:
  job_metadata.json (validator.py)
  section_rules.json (nếu KH gửi sheet format mới)

Priority 3 — Sau mỗi compile run:
  validation_report.txt
  → Check "MANUAL: No source_expr" count
  → Target: < 5% tổng số trường
```

---

## PHỤ LỤC A — MAPPING AST TYPE → VBA SKELETON

```vba
' DIRECT_MAPPING
INSERT INTO [Target] (T.[db_field]) VALUES (T1.[source_field])

' CONSTANT_MAPPING
INSERT INTO [Target] (T.[db_field]) VALUES ('constant_value')

' HYPHEN_TO_NULL
INSERT INTO [Target] (T.[db_field]) VALUES (IIF(T1.[src]='-', Null, T1.[src]))

' CONDITIONAL_MAPPING (SWITCH block)
UPDATE [Target] AS T INNER JOIN [Source] AS T1 ON T.ID = T1.ID
SET T.[db_field] = SWITCH(
    T1.[cond_field] = 'val_1', T1.[src_1],
    T1.[cond_field] = 'val_2', T1.[src_2],
    True, Null
)
WHERE T.[dev_flag] = True

' DELETE_RECORD
DELETE FROM [Target] AS T
INNER JOIN [Source] AS T1 ON T.ID = T1.ID
WHERE T1.[filter_field] IS NULL

' DEDUPLICATE_RULE
DELETE FROM [Target]
WHERE ID NOT IN (
    SELECT MIN(ID) FROM [Target] GROUP BY [key_field]
)

' FILTER_RULE — BLANK_CHECK
DELETE FROM [Target] AS T
INNER JOIN [Source] AS T1 ON T.ID = T1.ID
WHERE T1.[required_field] IS NULL OR T1.[required_field] = ''
```

---

## PHỤ LỤC B — CHECKLIST ONBOARDING BA MỚI

Để một BA mới có thể làm việc độc lập trên MetadataCompiler pipeline:

```
Tuần 1 — Nền tảng:
□ Đọc architecture_summary.md từ đầu đến cuối
□ Chạy thử pipeline với 1 job nhỏ (account_N)
□ Review 50 token đầu trong global_token_candidates.json

Tuần 2 — Thực hành:
□ Tự classify 1 sheet 通常処理 mới từ đầu đến cuối
□ Phân tích 1 sheet 特殊処理 phức tạp (có SWITCH block)
□ Khai báo job_metadata.json cho 1 job mới và validate thành công

Tuần 3 — Nâng cao:
□ Tách được điều kiện văn xuôi JP thành SQL tường minh
□ Xử lý được trường hợp token disambiguation (cùng token, khác ast_type)
□ Review output VBA và phát hiện được ít nhất 1 lỗi thiếu config

Sign-off: Lead BA xác nhận BA mới đạt 3 tuần trên = có thể làm độc lập
```
