# VBA Generation Workflow
# ワークフロー: JSON Spec → VBA Code → Review

> **Phạm vi áp dụng:** Mọi sheet module trong MetadataCompiler_V1  
> **Ví dụ thực thi:** `output/test1/sample/アカウント（入居者）_Y/`

---

## BƯỚC 1 — BẮT ĐẦU SCAN & CHUẨN BỊ WORKSPACE

1. Xác định thư mục test và sheet cần xử lý (VD: `output/test1/sample/アカウント（既存入居者）_P`).
2. Quét file `sheet_raw.json` và các tài liệu liên quan trong thư mục để nắm specs nghiệp vụ.

## BƯỚC 2 — KIỂM TRA SKILL & QUY TẮC

1. Đọc kỹ các file skill: `.agents/skills/vba-access-architect/SKILL.md`, `.agents/rules/vba-rule.md`, `.agents/skills/vba-development-standard/SKILL.md`.
2. Đọc file `sample/dictionary_proccess.json` để lấy các SQL pattern mẫu.

## BƯỚC 3 — GIAO VIỆC CHO AGENT

Gửi câu lệnh prompt chuẩn cho Agent (bắt đầu tiến trình sinh code hoặc review):

**Cho 1 file (Cần Approval):**
> "Hãy generate/review code VBA cho sheet `[Tên Sheet]` ở thư mục `[Tên thư mục test]`. Yêu cầu tuân thủ nghiêm ngặt các quy tắc trong `.agents/skills/vba-access-architect/SKILL.md`, đặc biệt là các quy tắc về tối ưu Production. Bắt buộc phải lập Implementation Plan chờ duyệt, và sau khi được duyệt, phải in ra Checklist & Chứng minh đầy đủ trước khi xuất khối code VBA chất lượng ra bên ngoài."

**Cho thực thi HÀNG LOẠT (Bỏ qua Approval):**
> "Hãy rà soát và chỉnh sửa code VBA hàng loạt cho các sheet trong thư mục `[Tên thư mục test]` so với spec `sheet_raw.json`. Yêu cầu: BỎ QUA việc lập Implementation Plan và chờ Approve. Agent hãy tự động phân tích, sửa code thẳng vào các file `.bas` nếu có sai lệch so với spec và `.agents/skills/vba-access-architect/SKILL.md`. Bắt buộc vẫn phải tự rà soát Checklist (Bước 4.3) trước khi hoàn tất mỗi file."

> **⚠️ CẢNH BÁO CRITICAL:**
> **KHÔNG BAO GIỜ** được chạy lại lệnh `python3 main.py --compile` hoặc `run` bằng terminal SAU KHI bạn đã thực hiện Bước 3 (nhờ AI Agent review/sửa code). 
> Lệnh `compile` của Python chỉ dùng để tạo bộ khung ban đầu. Nếu chạy lại, nó sẽ **ghi đè và xóa sạch toàn bộ** những logic phức tạp mà AI Agent vừa sửa tay trực tiếp trong file `.bas`, khiến đoạn code trở về trạng thái lỗi cũ (rất vô nghĩa và mất công).

---

### 4.1 Áp dụng skill rules (bắt buộc kết hợp)

| File | Vai trò |
|------|---------|
| `.agents/rules/vba-rule.md` | Rule nền tảng, SQL style, coding convention |
| `.agents/skills/vba-access-architect/SKILL.md` | Pattern chi tiết cho từng `ast_type`, FLG workflow, SWITCH, Dim con |
| `.agents/skills/vba-development-standard/SKILL.md` | Quality gate, tiêu chuẩn naming, logging, cleanup checklist |

### 4.2 Quy trình sinh code VBA — Step by step

```


Step 2 — PER SECTION (lặp cho từng section trong sheet_raw.json)
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

### 4.3 Checklist bắt buộc trước khi xuất code

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

### 4.4 Ví dụ output — `アカウント（入居者）_Y` (trích)

```vba


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

```

---

## BƯỚC 5 — SO SÁNH VÀ TẠO FILE NHẬN XÉT `review_report.md`

### 5.1 Mục tiêu

So sánh code VBA đã sinh (`.bas`) với spec JSON (`sheet_raw.json`) và file dictionary.  
Xuất file `review_report.md` với đánh giá chi tiết theo từng hạng mục.

### 5.2 Cấu trúc `review_report.md`

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

### 5.3 Tiêu chí đánh giá

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

## BƯỚC 6 — CHỈNH SỬA CODE SAU REVIEW (AUTO-FIX)

### 6.1 Mục tiêu
Sử dụng Agent để rà soát code VBA (do pipeline sinh ra) với Spec JSON, sau đó tự động sửa các điểm chưa chính xác nhằm đảm bảo code khớp 100% với nghiệp vụ và `.agents/skills/vba-access-architect/SKILL.md`.

### 6.2 Quy trình thực hiện
1. **Đọc Spec và Code:** Agent đọc file `sheet_raw.json` và file `.bas` hiện tại của sheet.
2. **Phân tích sai lệch (Gap Analysis):** Agent tự phát hiện các lỗi logic (thiếu JOIN, xử lý sai Hyphen, thiếu field, sai ưu tiên Dim con...). Có thể ghi chú nhanh hoặc tạo `review_report.md`.
3. **Chỉnh sửa Code (Auto-Fix):** Agent tiến hành sửa trực tiếp file `.bas`.
4. **Chạy Checklist (Bắt buộc):** Bất kể chế độ nào, Agent phải đảm bảo các mục trong Checklist (Bước 4.3) đều được pass.
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
