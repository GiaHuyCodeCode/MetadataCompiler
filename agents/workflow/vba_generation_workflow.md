# VBA Generation Workflow
# ワークフロー: JSON Spec → VBA Code → Review

> **Phạm vi áp dụng:** Mọi sheet module trong MetadataCompiler_V1  
> **Ví dụ thực thi:** `output/test1/sample/アカウント（入居者）_Y/`

---

## BƯỚC 1 — BẮT ĐẦU SCAN & CHUẨN BỊ WORKSPACE

1. Xác định thư mục test và sheet cần xử lý (VD: `output/test1/sample/アカウント（既存入居者）_P`).
2. Quét file `sheet_raw.json` và các tài liệu liên quan trong thư mục để nắm specs nghiệp vụ.

## BƯỚC 2 — KIỂM TRA SKILL & QUY TẮC

1. Đọc kỹ các file skill: `agents/sk-architect/SKILL.md`, `agents/rules/vba-rule.md`, `agents/sk-standard/SKILL.md`.
2. Đọc file `sample/dictionary_proccess.json` để lấy các SQL pattern mẫu.

## BƯỚC 3 — GIAO VIỆC CHO AGENT

Gửi câu lệnh prompt chuẩn cho Agent (bắt đầu tiến trình sinh code hoặc review):

**Cho 1 file (Cần Approval):**
> "Hãy generate/review code VBA cho sheet `[Tên Sheet]` ở thư mục `[Tên thư mục test]`. Yêu cầu tuân thủ nghiêm ngặt các quy tắc trong `SKILL.md`, đặc biệt là các quy tắc về tối ưu Production. Bắt buộc phải lập Implementation Plan chờ duyệt, và sau khi được duyệt, phải in ra Checklist & Chứng minh đầy đủ trước khi xuất khối code VBA chất lượng ra bên ngoài."

**Cho thực thi HÀNG LOẠT (Bỏ qua Approval):**
> "Hãy rà soát và chỉnh sửa code VBA hàng loạt cho các sheet trong thư mục `[Tên thư mục test]` so với spec `sheet_raw.json`. Yêu cầu: BỎ QUA việc lập Implementation Plan và chờ Approve. Agent hãy tự động phân tích, sửa code thẳng vào các file `.bas` nếu có sai lệch so với spec và `SKILL.md`. Bắt buộc vẫn phải tự rà soát Checklist (Bước 5.4) trước khi hoàn tất mỗi file."

---

## BƯỚC 4 — AGENT ĐỌC SPEC & TẠO `summary_sheet.json`

### 4.1 Mục tiêu

Đọc file `sheet_raw.json` (hoặc `sheet_spec.json`) trong thư mục output của sheet cần xử lý.  
Tổng hợp nội dung thành file `summary_sheet.json` — chỉ giữ **4 section chính** và các **câu sentences tổng hợp** (không liệt kê raw data row).

### 4.2 Cấu trúc `summary_sheet.json` bắt buộc

```json
{
  "_meta": {
    "sheet_name": "<tên sheet>",
    "target_table": "<tên bảng đích VBA>",
    "id_prefix": "<prefix ID, ví dụ: Y_>",
    "primary_source": "<bảng nguồn chính>"
  },
  "ファイル名": {
    "sentence": "<tên file CSV/xlsx dùng làm nguồn dữ liệu>"
  },
  "出力条件": {
    "sentences": [
      "<câu mô tả điều kiện lọc 1>",
      "<câu mô tả điều kiện lọc 2>"
    ]
  },
  "通常処理": {
    "sentences": [
      "<câu mô tả mapping field thông thường>"
    ]
  },
  "特殊処理": {
    "sentences": [
      "<câu mô tả nhóm xử lý đặc biệt 1>",
      "<câu mô tả nhóm xử lý đặc biệt 2>"
    ]
  }
}
```

### 4.3 Quy tắc tổng hợp sentences

| Section | Quy tắc |
|---------|---------|
| `出力条件` | Mỗi nhóm lọc → 1 câu ngắn mô tả mục đích + điều kiện (ví dụ: *"契約状況が「契約中」または「解約予定」のみを対象とする"*) |
| `通常処理` | Liệt kê các field có `処理: 固定値出力` hoặc `DIRECT_MAPPING` thành 1 câu INSERT tổng quát |
| `特殊処理` | Nhóm các rows có cùng `条件 / 項目マッピング` vào 1 câu. Bỏ qua `is_active: False` và `処理: 項目マッピング` |

### 4.4 Ví dụ — `アカウント（入居者）_Y`

```json
{
  "_meta": {
    "sheet_name": "アカウント（入居者）_Y",
    "target_table": "Account",
    "id_prefix": "Y_",
    "primary_source": "入居状況一覧"
  },
  "ファイル名": {
    "sentence": "GMO 入居状況一覧.csv を使用する"
  },
  "出力条件": {
    "sentences": [
      "入居状況一覧の「契約状況」が「契約中」または「解約予定」のレコードのみを出力対象とし、それ以外はFLGで除外する（IN_LIST_FILTER / FLG Logic Ngược）"
    ]
  },
  "通常処理": {
    "sentences": [
      "AccountテーブルにID='Y_'+T.ID、klass='Resident'、legacy_charge_user_id='gmo002'、sheet='アカウント（入居者）_Y' を固定値でINSERTする"
    ]
  },
  "特殊処理": {
    "sentences": [
      "【優先順位①: 契約者1が個人かつ2人とも入居有り】name_family, company_name, name_family_kana, company_name_kana, email, tel_fixed, tel_mobile, kind_id, birthday を契約者1（列番号_33/_35/_36/_38/_39/_40/_42）から更新する",
      "【優先順位②: 契約者3が入居有り（上記以外）】同フィールドを契約者3（列番号_53/_55/_56/_58/_59/_60/_62）から更新する",
      "【優先順位③: 契約者2が入居有り（上記以外）】同フィールドを契約者2（列番号_43/_45/_46/_48/_49/_50/_52）から更新する",
      "【優先順位④: いずれにも該当しない場合】契約者1のデータ（デフォルト）を使用する",
      "【個人・法人区分による後処理】個人の場合: company_name / company_name_kana をNULLクリア。法人の場合: name_family / name_family_kana をNULLクリア",
      "【kind_id変換】kind_idが'個人'の場合→'10'、'法人'の場合→'20' に変換する",
      "【tag】契約状況が'解約予定'の場合、tagフィールドに'解約予定'をセットする"
    ]
  }
}
```

---

## BƯỚC 5 — TRA CỨU DICTIONARY & SINH CODE VBA

### 5.1 Tra cứu `sample/dictionary_proccess.json`

1. Tìm entry có `sheet_pattern` khớp với tên sheet đang xử lý.
2. Nếu **tìm thấy**: sử dụng `output_conditions`, `special_processing` trong dictionary làm **tham chiếu SQL pattern**.
3. Nếu **không tìm thấy**: chuyển sang bước 5.2 để sinh code thuần từ spec.

> **Lưu ý:** Dictionary chứa SQL tham chiếu — có thể thiếu điều kiện SWITCH hoặc có placeholder rỗng.  
> Bắt buộc xác minh lại với `summary_sheet.json` và các skill rules trước khi dùng.

### 5.2 Áp dụng skill rules (bắt buộc kết hợp)

| File | Vai trò |
|------|---------|
| `agents/rules/vba-rule.md` | Rule nền tảng, SQL style, coding convention |
| `agents/sk-architect/SKILL.md` | Pattern chi tiết cho từng `ast_type`, FLG workflow, SWITCH, Dim con |
| `agents/sk-standard/SKILL.md` | Quality gate, tiêu chuẩn naming, logging, cleanup checklist |

### 5.3 Quy trình sinh code VBA — Step by step

```


Step 2 — PER SECTION (lặp cho từng section trong summary_sheet.json)
  ├─ '======================================================================
  ├─ '<sheet_name>
  ├─ '======================================================================
  │
  ├─ 出力条件
  │   ├─ AddNewFieldToTable "<SourceTable>", "FLG", "TEXT(1)"
  │   ├─ FLGリセット:
  │   │   ├─ Nếu IN_LIST_FILTER → db.Execute "UPDATE ... SET FLG='1'" (logic ngược)
  │   │   └─ Mặc định → db.Execute "UPDATE ... SET FLG='0'"
  │   ├─ UPDATE blocks theo ast_type (mỗi điều kiện 1 block riêng)
  │   │   ├─ BLANK_CHECK → WHERE Nz(T.[Field],'')=''
  │   │   ├─ IN_LIST_FILTER + NOT_IN → SET FLG='0' WHERE field IN (...)
  │   │   ├─ JOIN_FILTER → UPDATE INNER JOIN SET FLG='1' (không WHERE)
  │   │   ├─ DEDUP → WHERE ID NOT IN (SELECT MIN(ID)...) — luôn cuối cùng
  │   │   └─ FLG Multi-value → FLG='1','2','3' khi loại trừ nhiều tầng
  │   └─ log_write "set_<table>:<section> → 不要行を削除するため(FLG=1更新)"
  │
  ├─ 通常処理
  │   ├─ INSERT INTO <TargetTable> SELECT
  │   │   ├─ '<prefix>_' & T.ID as ID  ← bắt buộc
  │   │   ├─ 'FixedValue' as field     ← CONSTANT
  │   │   ├─ T.[SourceField] as field  ← DIRECT_MAPPING
  │   │   └─ '<sheet_name>' as sheet   ← bắt buộc
  │   ├─ FROM <SourceTable> AS T WHERE T.FLG='0'
  │   └─ log_write "set_<table>:<section> → 通常処理"
  │
  ├─ 特殊処理
  │   ├─ Dim con1/con2/con3 (khai báo 1 lần nếu dùng SWITCH nhiều block)
  │   ├─ UPDATE blocks:
  │   │   ├─ UPDATE_SWITCH → SWITCH(..., True, NULL)
  │   │   ├─ UPDATE_WHERE_EQUALS → WHERE cond_field='val'
  │   │   ├─ UPDATE_LIKE → WHERE field LIKE/NOT LIKE '%pattern%'
  │   │   ├─ UPDATE_SWITCH_MULTI_FIELD → gộp nhiều field cùng WHERE
  │   │   └─ Multi-join → (Table1 INNER JOIN Table2) INNER JOIN Table3
  │   └─ log_write cho mỗi block "set_<table>:<section> → 特殊処理: <field>"
  │
  └─ DeleteFieldInTable "<SourceTable>", "FLG"

Step 3 — FOOTER
  ├─ chk_required  ← GỌI 1 LẦN DUY NHẤT, cuối Sub
  ├─ Debug.Print "<table>:" & Timer - t
  └─ log_write "set_<table>:out"
```

### 5.4 Checklist bắt buộc trước khi xuất code

- [ ] Có section separator `'====...`
- [ ] SQL build bằng `SQL = SQL &` (không hardcode 1 dòng)
- [ ] Alias chỉ dùng T / T1 / T2 / T3
- [ ] Có log sau mỗi `db.Execute SQL`
- [ ] INSERT có ID (với prefix) và sheet
- [ ] SWITCH thay cho IIF lồng nhau (≥3 nhánh)
- [ ] Nz() xử lý NULL
- [ ] Nz(Replace(field,'-',''),'')='' trong UPDATE SET khi kiểm tra hyphen
- [ ] Gộp fields cùng WHERE vào 1 UPDATE
- [ ] Dim con khai báo 1 lần trước tất cả SWITCH block
- [ ] DeleteFieldInTable xóa FLG sau mỗi section
- [ ] chk_required 1 lần duy nhất cuối Sub
- [ ] Không tự suy đoán business logic ngoài spec
- [ ] Tên Sub: `set_account()` (lowercase)

### 5.5 Ví dụ output — `アカウント（入居者）_Y` (trích)

```vba
Attribute VB_Name = "アカウント（入居者）_Y"
Option Compare Database
Option Explicit

Sub set_account()
    log_write "set_account:in"

    Dim t As Single
    t = Timer

    Dim db As ADODB.Connection
    Dim SQL As String

    Set db = CurrentProject.Connection

    db.Execute "DELETE FROM Account WHERE ID LIKE 'Y_%';"
    log_write "account delete"

    '================================================================================
    'アカウント（入居者）_Y
    '================================================================================

    '出力条件
    AddNewFieldToTable "入居状況一覧", "FLG", "TEXT(1)"

    '■FLGリセット (IN_LIST_FILTER: logic ngược — khởi đầu tất cả FLG='1')
    db.Execute "UPDATE [入居状況一覧] SET [入居状況一覧].[FLG] = '1';"

    '■有効レコードのみFLG=0に戻す
    SQL = ""
    SQL = SQL & "UPDATE [入居状況一覧] AS T "
    SQL = SQL & "SET T.[FLG] = '0' "
    SQL = SQL & "WHERE T.[契約状況] = '契約中' OR T.[契約状況] = '解約予定'; "
    db.Execute SQL

    log_write "set_account:アカウント（入居者）_Y → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Account SELECT "
    SQL = SQL & "    'Y_' & T.ID as ID, "
    SQL = SQL & "    'Resident' as klass, "
    SQL = SQL & "    'gmo002' as legacy_charge_user_id, "
    SQL = SQL & "    'アカウント（入居者）_Y' as sheet "
    SQL = SQL & "FROM [入居状況一覧] AS T "
    SQL = SQL & "WHERE T.[FLG] = '0';"
    db.Execute SQL
    log_write "set_account:アカウント（入居者）_Y → 通常処理"

    '特殊処理
    ' Dim con — ưu tiên 契約者1→3→2 (khai báo 1 lần)
    Dim con1 As String
    Dim con3 As String
    Dim con2 As String
    con1 = "(T1.[個人・法人区分_33] = '個人' AND T1.[個人・法人区分_43] = '個人' AND T1.[契約者1入居有無] = '入居有り' AND T1.[契約者2入居有無] = '入居有り')"
    con3 = "(T1.[契約者3入居有無] = '入居有り')"
    con2 = "(T1.[契約者2入居有無] = '入居有り')"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN [入居状況一覧] AS T1 ON T.ID = ('Y_' & T1.ID) "
    SQL = SQL & "SET T.[name_family] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者名_35], "
    SQL = SQL & "    " & con3 & ", T1.[契約者名_55], "
    SQL = SQL & "    " & con2 & ", T1.[契約者名_45], "
    SQL = SQL & "    True, T1.[契約者名_35] ); "
    db.Execute SQL
    log_write "set_account:アカウント（入居者）_Y → 特殊処理: name_family"

    ' ... (company_name, name_family_kana, email, tel_fixed, tel_mobile, kind_id, birthday) ...

    ' 個人の場合: company_name / company_name_kana をNULLクリア
    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN [入居状況一覧] AS T1 ON T.ID = ('Y_' & T1.ID) "
    SQL = SQL & "SET T.[company_name] = NULL, "
    SQL = SQL & "    T.[company_name_kana] = NULL "
    SQL = SQL & "WHERE T1.[個人・法人区分_33] = '個人' "
    SQL = SQL & "  AND T1.[個人・法人区分_43] = '個人' "
    SQL = SQL & "  AND T1.[個人・法人区分_53] = '個人'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者）_Y → 特殊処理: 個人→company_name NULL"

    ' kind_id: '個人'→'10', '法人'→'20'
    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    T.[kind_id] = '個人', '10', "
    SQL = SQL & "    T.[kind_id] = '法人', '20', "
    SQL = SQL & "    True, NULL ) "
    SQL = SQL & "WHERE T.ID LIKE 'Y_%'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者）_Y → 特殊処理: kind_id変換"

    ' tag: 解約予定
    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN [入居状況一覧] AS T1 ON T.ID = ('Y_' & T1.ID) "
    SQL = SQL & "SET T.[tag] = '解約予定' "
    SQL = SQL & "WHERE T1.[契約状況] = '解約予定'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者）_Y → 特殊処理: tag"

    '■項目削除
    DeleteFieldInTable "[入居状況一覧]", "FLG"

    chk_required

    Debug.Print "Account:" & Timer - t
    log_write "set_account:out"

End Sub
```

---

## BƯỚC 6 — SO SÁNH VÀ TẠO FILE NHẬN XÉT `review_report.md`

### 6.1 Mục tiêu

So sánh code VBA đã sinh (`.bas`) với spec JSON (`summary_sheet.json`) và file dictionary.  
Xuất file `review_report.md` với đánh giá chi tiết theo từng hạng mục.

### 6.2 Cấu trúc `review_report.md`

```markdown
# Review Report: <sheet_name>

**File VBA:** `<đường dẫn file .bas>`  
**Spec JSON:** `<đường dẫn sheet_raw.json>`  
**Ngày review:** <YYYY-MM-DD>

---

## 1. TỔNG QUAN

| Hạng mục | Trạng thái | Ghi chú |
|----------|-----------|---------|
| FLG workflow | ✅ / ⚠️ / ❌ | ... |
| IN_LIST_FILTER logic | ... | ... |
| 通常処理 INSERT | ... | ... |
| 特殊処理 SWITCH | ... | ... |
| Dim con pattern | ... | ... |
| log_write đầy đủ | ... | ... |
| Cleanup FLG | ... | ... |
| chk_required | ... | ... |

---

## 2. CHI TIẾT TỪNG SECTION

### 2.1 出力条件

**Spec yêu cầu:**
> ...

**VBA hiện tại:**
```vba
...
```

**Nhận xét:** ✅ Đúng / ⚠️ Cần điều chỉnh / ❌ Sai

**Lý do / Gợi ý sửa:**
> ...

---

### 2.2 通常処理

...

### 2.3 特殊処理

...

---

## 3. CÁC VẤN ĐỀ PHÁT HIỆN

| # | Mức độ | Vị trí | Mô tả vấn đề | Gợi ý sửa |
|---|--------|--------|-------------|----------|
| 1 | 🔴 Critical | line XX | ... | ... |
| 2 | 🟡 Warning | ... | ... | ... |
| 3 | 🟢 Info | ... | ... | ... |

---

## 4. KẾT LUẬN

**Tổng điểm chất lượng:** X / 10  
**Kết luận:** Đạt chuẩn production / Cần sửa trước khi deploy  
**Các bước tiếp theo:** ...
```

### 6.3 Tiêu chí đánh giá

| Hạng mục | Mức độ | Điểm |
|----------|--------|------|
| FLG logic đúng (IN_LIST_FILTER vs BLANK_CHECK) | Critical | 2 |
| SWITCH condition đúng thứ tự ưu tiên (con1→con3→con2) | Critical | 2 |
| INSERT đủ field (ID prefix, sheet) | High | 1 |
| Dim con khai báo 1 lần, dùng lại đúng | High | 1 |
| Gộp fields cùng WHERE | Medium | 1 |
| log_write đầy đủ sau mọi db.Execute | Medium | 1 |
| Cleanup FLG + chk_required | Medium | 1 |
| Không suy đoán logic ngoài spec | Critical | 1 |

---

## BƯỚC 7 — CHỈNH SỬA CODE SAU REVIEW (AUTO-FIX)

### 7.1 Mục tiêu
Sử dụng Agent để rà soát code VBA (do pipeline sinh ra) với Spec JSON, sau đó tự động sửa các điểm chưa chính xác nhằm đảm bảo code khớp 100% với nghiệp vụ và `SKILL.md`.

### 7.2 Quy trình thực hiện
1. **Đọc Spec và Code:** Agent đọc file `sheet_raw.json` và file `.bas` hiện tại của sheet.
2. **Phân tích sai lệch (Gap Analysis):** Agent tự phát hiện các lỗi logic (thiếu JOIN, xử lý sai Hyphen, thiếu field, sai ưu tiên Dim con...). Có thể ghi chú nhanh hoặc tạo `review_report.md`.
3. **Chỉnh sửa Code (Auto-Fix):** Agent tiến hành sửa trực tiếp file `.bas`.
4. **Chạy Checklist (Bắt buộc):** Bất kể chế độ nào, Agent phải đảm bảo các mục trong Checklist (Bước 5.4) đều được pass.
5. **Chế độ Hàng loạt (Bulk Mode):** 
   - Nếu user gọi thực hiện hàng loạt nhiều file, Agent **không cần** dừng lại hỏi Approve / Implementation Plan.
   - Tiến hành lặp lại quy trình [Đọc $\rightarrow$ Phân tích $\rightarrow$ Sửa $\rightarrow$ Checklist] cho từng file một cách tự động và liên tục cho đến khi hoàn thành.

---

## PHỤ LỤC — MAPPING NHANH CÁC VẤN ĐỀ THƯỜNG GẶP

| Vấn đề | Triệu chứng | Cách sửa |
|--------|------------|---------|
| FLG reset sai | IN_LIST_FILTER nhưng reset về '0' | Đổi reset về '1', rồi SET='0' cho record tốt |
| SWITCH condition rỗng | `SWITCH( , T1.[field], ...)` | Điền đúng con1/con2/con3 từ spec |
| Thiếu con3 trong thứ tự | Bỏ qua 契約者3 | Thêm con3 giữa con1 và con2 |
| ID LIKE pattern sai | `T.ID LIKE '%Y_%'` (match bất kỳ Y) | Dùng `T.ID LIKE 'Y_%'` (chỉ match prefix Y_) |
| Hyphen check sai trong UPDATE SET | `T1.[field] = '-'` | Dùng `Nz(Replace(T1.[field],'-',''),'')=''` |
| Không gộp fields cùng WHERE | Nhiều UPDATE riêng với WHERE giống nhau | Gộp SET nhiều field vào 1 UPDATE |
| tag field rỗng | `SET T.[tag] = T1.[]` | Dùng `SET T.[tag] = '解約予定'` (固定値) |
| Set db sai | `Set db = CurrentDb` | `Set db = CurrentProject.Connection` |
