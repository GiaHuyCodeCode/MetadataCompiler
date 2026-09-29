---
name: vba-access-architect
description: "Chuyên gia kiến trúc Microsoft Access VBA — chuyển đổi đặc tả nghiệp vụ có cấu trúc (JSON spec) thành mã VBA/SQL chất lượng production cho hệ thống quản lý bất động sản Nhật Bản (GMO). Trigger khi: user yêu cầu sinh code VBA từ JSON spec, viết Sub set_xxx(), tạo INSERT/UPDATE SQL cho Access, mapping field từ bảng nguồn JP sang bảng đích, xử lý 出力条件 / 通常処理 / 特殊処理, hoặc bất kỳ tác vụ nào liên quan đến data migration VBA Access với đặc tả tiếng Nhật. Luôn dùng skill này khi thấy từ khóa: set_account, set_contractor, FLG, AddNewFieldToTable, DeleteFieldInTable, log_write, chk_required, 出力条件, 通常処理, 特殊処理, JSON spec → VBA, IN_LIST_FILTER, UPDATE_SWITCH, Dim con pattern, 契約者優先順, FLG logic ."
---

# CHUYÊN GIA KIẾN TRÚC ACCESS VBA (TỪ SPEC SANG CODE) — v4.0

## VAI TRÒ

Chuyển đổi đặc tả nghiệp vụ có cấu trúc (JSON) thành mã VBA chất lượng cao, sẵn sàng production.
Tuân thủ nghiêm ngặt template và best practices. Chỉ dùng thông tin trong JSON + template, không suy đoán logic ngoài đặc tả.

> **Tham chiếu nhanh:** Mục 5 = chi tiết từng pattern · Mục 6 = workflow · Mục 7 = checklist chuẩn

---

## 1. CẤU TRÚC DỮ LIỆU ĐẦU VÀO

**Đặc tả JSON:** Chứa `sheet_name`, `出力条件` (điều kiện lọc), `通常処理` (mapping), `特殊処理` (logic phức tạp).

**Các `ast_type` trong JSON và VBA pattern tương ứng:**

| ast_type | VBA Pattern sinh ra |
|----------|-------------------|
| `BLANK_CHECK` | UPDATE SET FLG='1' WHERE Nz(field,'')='' |
| `IN_LIST_FILTER` | FLG Logic Ngược: reset FLG='1', set FLG='0' cho record hợp lệ |
| `JOIN_FILTER` | UPDATE INNER JOIN SET FLG='1' (không WHERE) |
| `DEDUP` | UPDATE WHERE ID NOT IN (SELECT MIN(ID)...) |
| `CONSTANT` | `'Value' as field` trong INSERT |
| `DIRECT_MAPPING` | `T.[field] as alias` trong INSERT |
| `UPDATE_SWITCH` | UPDATE INNER JOIN SET field = SWITCH(..., True, NULL) |
| `UPDATE_WHERE_EQUALS` | UPDATE INNER JOIN SET field='val' WHERE cond_field='cond_val' |
| `UPDATE_LIKE` | UPDATE INNER JOIN SET ... WHERE field LIKE/NOT LIKE '%pattern%' |
| `UPDATE_SWITCH_MULTI_FIELD` | Nhiều field=SWITCH() trong 1 UPDATE, dùng Dim con |

---

## 2. CẤU TRÚC SQL BẮT BUỘC

1. Khởi tạo: `SQL = ""`
2. Nối chuỗi: `SQL = SQL & "Dòng lệnh SQL... "` (mỗi phần logic một dòng)
3. Thực thi: `db.Execute SQL`
4. Log: `log_write "Nội dung log"`
5. **Tên bảng đích:** Bỏ phần chữ Nhật trong ngoặc, viết hoa chữ đầu. Ví dụ: `account（入居者）` → `Account`, `set_account()`.
6. **SWITCH nhiều điều kiện:** Mỗi điều kiện một dòng riêng, luôn kết thúc bằng `True, NULL`.
7. **NULL safety:** Dùng `Nz(T.[Field], '')` thay vì `IS NULL` khi field có thể là chuỗi rỗng.
8. Tránh subquery phức tạp trong `出力条件` — dùng UPDATE riêng với FLG.
9. Không được Self-Join.
10. Table Name nào đứng trước sẽ được ưu tiên hơn Table Name đứng sau khi Join.
11. Bảng GMO 入居状況一覧, 【GMO用】新規契約更新一覧  khi sử dụng thì chỉ giữ lại 入居状況一覧 , 新規契約更新一覧  (bắt buộc).
12. Sủ dụng % thay vì * khi tạo LIKE hoặc NOT LIKE trong SQL (vd: T.[FieldName] LIKE '%pattern%'). 
13. Đa phần sử dụng FLG Logic (99%), 1% còn lại mới sử dụng FLg Logic ngược. 
14. Nếu đã gắn field `ID` với prefix thì khi xết các điều kiện như dedup hoặc like thì phải đính kèm lọc các phần có gắn prefix chứ không được lọc toàn bộ ID.
15. Ở Section "特殊処理" nếu có cùng 条件 / 項目マッピング thì ta gom lại làm 1 lần, không làm riêng lẻ.
16. Các từ tiếng Nhật nên được đặt trong ngoặc vuông [].


---

## 3. TỪ ĐIỂN MAPPING

| Token JP | VBA/SQL Pattern |
|----------|----------------|
| **そのまま出力** | `T.[FieldName] as Alias` |
| **固定値出力** | `'Value' as Alias` |
| **yyyy-mm-dd形式** | `IIF(IsDate(T.[Field]), Format(T.[Field], 'yyyy-mm-dd'), NULL) as Alias` |
| **yyyymm形式** | `Format(T.[Field], "yyyymm") as Alias` |
| **ハイフン付出力** | `T.[F1] & '-' & T.[F2] as Alias` |
| **指定文字削除出力** | `Val(Replace(Replace(T.[Field], ',', ''), '円', '')) as Alias` |
| **ハイフンをNull変換** | `IIF(T.[Field] = '-', NULL, T.[Field]) as Alias` |
| **条件分岐** | `SWITCH(cond1, val1, ..., True, NULL)` |
| **削除** | `FLG = '1'` |
| **出力なし** | Loại trừ khỏi SQL |
| **マージ** | `T.[F1] & T.[F2] as Alias` |
| **重複チェック** | `ID NOT IN (SELECT MIN(ID) FROM Source WHERE FLG='0' GROUP BY Field)` |
| **ID Prefix** | `'XX_' & T.ID as ID` |

---

## 4. QUY TẮC KHÔNG SINH HEADER / FOOTER

> **⚠️ TUYỆT ĐỐI KHÔNG SINH CÁC THÀNH PHẦN SAU:**
> 1. KHÔNG sinh phần Header (ví dụ: `Attribute VB_Name`, `Option Explicit`, `Sub set_xxx()`, `Dim db`, `Set db`, `DELETE FROM` ban đầu).
> 2. KHÔNG sinh phần Footer (ví dụ: `chk_required`, `Debug.Print`, `End Sub`).
> 3. CHỈ SINH phần thân (Logic SQL cốt lõi: Tạo FLG, Điều kiện loại trừ, Insert/Update xử lý). Hệ thống Native Python Compiler đã tự động gộp phần boilerplate (Header/Footer) rồi.

## 5. CHI TIẾT TỪNG PHẦN XỬ LÝ

### A. KHỞI TẠO FLG (đầu mỗi section)

```vba
'出力条件
AddNewFieldToTable "[SourceTable]", "FLG", "TEXT(1)"

'■FLGリセット
db.Execute "UPDATE [SourceTable] SET [SourceTable].[FLG] = '0';"
```

- FLG thêm vào **bảng nguồn** (không phải bảng đích).
- Mặc định reset `'0'` (giữ). **Ngoại lệ:** `IN_LIST_FILTER` → dùng FLG Logic Ngược (xem 5.D).
- **Chỉ AddNewFieldToTable cho bảng cần track trạng thái** (bảng được INSERT hoặc cần lọc theo nhiều điều kiện). Bảng chỉ dùng để JOIN (lookup table) → **KHÔNG** thêm FLG.

**FLG Multi-value** — dùng khi logic loại trừ có nhiều tầng trong cùng 1 bảng:

> Thay vì binary `'0'/'1'`, dùng giá trị số để phân tầng — INSERT chỉ lấy `FLG='0'`.

**Bảng mapping giá trị FLG theo tầng loại trừ:**

| Giá trị FLG | Ý nghĩa | WHERE trong block tiếp theo |
|-------------|---------|----------------------------|
| `'0'` | Giữ lại (chưa bị loại) | `WHERE T.[FLG] = '0'` |
| `'1'` | Bị loại — tầng 1 (điều kiện chính, JOIN không có WHERE) | — |
| `'2'` | Bị loại — tầng 2 (cascading từ FLG='0' còn lại) | `WHERE T.[FLG] = '0' AND T2.[FLG] = '0'` |
| `'3'` | Bị loại — tầng 3 (cascading tiếp theo) | `WHERE T.[FLG] = '0'` |

> **Quy tắc WHERE theo tầng:**
> - Tầng 1 (JOIN ON SWITCH hoặc JOIN_FILTER đơn): **KHÔNG WHERE** — toàn bộ record khớp JOIN đều bị loại
> - Tầng 2 trở đi: **BẮT BUỘC có WHERE T.[FLG] = '0'** để chỉ xét record chưa bị loại ở tầng trước
> - Nếu tầng sau cần lọc thêm qua bảng phụ có FLG riêng: thêm `AND T2.[FLG] = '0'`

```vba
'■FLGリセット
db.Execute "UPDATE [SourceTable] SET [SourceTable].[FLG] = '0';"

'Tầng 1: Loại trừ theo điều kiện chính → FLG='1' (không cần WHERE)
SQL = ""
SQL = SQL & "UPDATE [SourceTable] AS T1 "
SQL = SQL & "INNER JOIN [Table2] AS T2 ON T1.[Key] = T2.[Key] "
SQL = SQL & "SET T1.[FLG] = '1' "
db.Execute SQL

'Tầng 2: Loại trừ thêm (chỉ xét FLG='0' còn lại, qua bảng phụ cũng chưa bị loại) → FLG='2'
SQL = ""
SQL = SQL & "UPDATE [SourceTable] AS T1 "
SQL = SQL & "INNER JOIN [Table3] AS T2 ON T1.[Key] = T2.[Key] "
SQL = SQL & "SET T1.[FLG] = '2' "
SQL = SQL & "WHERE T1.[FLG] = '0' AND T2.[FLG] = '0' "
db.Execute SQL

'Tầng 3: Loại trừ thêm (chỉ xét FLG='0' còn lại) → FLG='3'
SQL = ""
SQL = SQL & "UPDATE [SourceTable] AS T1 "
SQL = SQL & "INNER JOIN [Table4] AS T2 ON T1.[Key] = T2.[Key] "
SQL = SQL & "SET T1.[FLG] = '3' "
SQL = SQL & "WHERE T1.[FLG] = '0' "
db.Execute SQL

'→ INSERT chỉ lấy FLG='0' (không khớp bất kỳ điều kiện loại trừ nào)
```

**Ví dụ thực tế — アカウント（既存入居者）_P (3 tầng loại trừ):**
```vba
'Tầng 1: JOIN ON SWITCH (ưu tiên 契約者1→3→2) → FLG='1' (không WHERE)
SQL = ""
SQL = SQL & "UPDATE [入居者管理] AS T1 "
SQL = SQL & "INNER JOIN [入居状況一覧] AS T2 "
SQL = SQL & "ON T1.[基幹入居者ID] = Switch( "
SQL = SQL & "    ((T2.[個人・法人区分_33] = '個人') AND (T2.[個人・法人区分_43] = '個人') "
SQL = SQL & "        AND (T2.[契約者1入居有無] = '入居有り') AND (T2.[契約者2入居有無] = '入居有り')), T2.[契約者1No], "
SQL = SQL & "    (T2.[契約者3入居有無] = '入居有り'), T2.[契約者3No], "
SQL = SQL & "    (T2.[契約者2入居有無] = '入居有り'), T2.[契約者2No], "
SQL = SQL & "    True, T2.[契約者1No] "
SQL = SQL & ") "
SQL = SQL & "SET T1.[FLG] = '1' "
db.Execute SQL

'Tầng 1b: 新規契約更新一覧 phụ trợ → FLG='1' (không WHERE, bảng con tách riêng)
SQL = ""
SQL = SQL & "UPDATE [新規契約更新一覧] AS T1 "
SQL = SQL & "INNER JOIN [入居状況一覧] AS T2 "
SQL = SQL & "    ON T1.[物件No] = T2.[物件No] "
SQL = SQL & "    AND T1.[部屋No] = T2.[部屋No] "
SQL = SQL & "    AND T1.[契約者No] = T2.[契約者1No] "
SQL = SQL & "SET T1.[FLG] = '1' "
db.Execute SQL

'Tầng 2: cascading qua 新規契約更新一覧 (chỉ FLG='0' ở cả 2 bảng) → FLG='2'
SQL = ""
SQL = SQL & "UPDATE [入居者管理] AS T1 "
SQL = SQL & "INNER JOIN [新規契約更新一覧] AS T2 "
SQL = SQL & "    ON T1.[基幹入居者ID] = T2.[契約者No] "
SQL = SQL & "SET T1.[FLG] = '2' "
SQL = SQL & "WHERE T1.[FLG] = '0' AND T2.[FLG] = '0' "
db.Execute SQL

'Tầng 3: cascading qua 解約情報一覧 (chỉ xét FLG='0' còn lại) → FLG='3'
SQL = ""
SQL = SQL & "UPDATE [入居者管理] AS T1 "
SQL = SQL & "INNER JOIN [解約情報一覧] AS T2 "
SQL = SQL & "    ON T1.[基幹入居者ID] = T2.[契約者No] "
SQL = SQL & "SET T1.[FLG] = '3' "
SQL = SQL & "WHERE T1.[FLG] = '0' "
db.Execute SQL
```

**Khi nào dùng FLG multi-value:**
- Logic loại trừ có thứ tự ưu tiên rõ ràng (loại A trước, rồi loại B trong số còn lại, rồi loại C...)
- Cần JOIN qua nhiều bảng với điều kiện cascading
- Binary FLG không đủ để phân biệt "bị loại vì lý do gì" (hỗ trợ debug/trace)

---

### B. 出力条件 (Điều kiện lọc)

- Mỗi điều kiện = một khối `SQL = "" ... db.Execute SQL` riêng biệt.
- Chỉ 1 `log_write` DUY NHẤT sau TẤT CẢ các khối SQL của section.
- Format log: `"set_[table]:[SheetName] → 不要行を削除するため(FLG=1更新)"`

**① Blank check (`BLANK_CHECK`):**
```vba
SQL = ""
SQL = SQL & "UPDATE [SourceTable] AS T "
SQL = SQL & "SET T.FLG = '1' "
SQL = SQL & "WHERE Nz(T.[FieldName], '') = ''; "
db.Execute SQL
```

**② JOIN-based filter (`JOIN_FILTER`) — loại bỏ bản ghi KHỚP:**
```vba
SQL = ""
SQL = SQL & "UPDATE [SourceTable] AS T1 "
SQL = SQL & "INNER JOIN [OtherTable] AS T2 ON T1.[Key1] = T2.[Key1] AND T1.[Key2] = T2.[Key2] "
SQL = SQL & "SET T1.FLG = '1' "
db.Execute SQL
```
> ⚠️ KHÔNG thêm WHERE vào JOIN block này. Bản ghi khớp JOIN → FLG=1.

**③ Dedup (`DEDUP`) - Loại bỏ bản ghi trùng lặp**
```vba
SQL = ""
SQL = SQL & "UPDATE [SourceTable] "
SQL = SQL & "SET [SourceTable].FLG = '1' "
SQL = SQL & "WHERE [SourceTable].ID NOT IN ( "
SQL = SQL & "    SELECT MIN(ID) "
SQL = SQL & "    FROM [SourceTable] "
SQL = SQL & "    WHERE FLG = '0' "
SQL = SQL & "    GROUP BY [KeyField] "
SQL = SQL & ");"
db.Execute SQL
```

**④ IN list — GIỮ record hợp lệ (`IN_LIST_FILTER` + `filter_logic: NOT_IN`):**

> ⚠️ Đây là **FLG Logic Ngược** — phải reset FLG='1' TRƯỚC (thay vì '0' mặc định), rồi set FLG='0' cho record hợp lệ.

```vba
'■FLGリセット (logic ngược: khởi đầu tất cả là FLG=1 = bị loại)
db.Execute "UPDATE [SourceTable] SET [SourceTable].[FLG] = '1';"

'■有効レコードのみFLG=0に戻す
SQL = ""
SQL = SQL & "UPDATE [SourceTable] AS T "
SQL = SQL & "SET T.FLG = '0' "
SQL = SQL & "WHERE T.[Field] = 'Value1' OR T.[Field] = 'Value2'; "
db.Execute SQL
```

Nhận biết `IN_LIST_FILTER` khi JSON có:
- `ast_type: "IN_LIST_FILTER"` với `filter_logic: "NOT_IN"`, và `in_list_values: [...]`
- Hoặc: điều kiện văn xuôi dạng `"AまたはBの場合"` (giữ A hoặc B, loại bỏ tất cả còn lại)

**⑤ Kết hợp nhiều điều kiện (OR/AND):**
```vba
SQL = ""
SQL = SQL & "UPDATE [SourceTable] AS T "
SQL = SQL & "SET T.FLG = '1' "
SQL = SQL & "WHERE Nz(T.[Field], '') = '' OR T.[Field] = '0'; "
db.Execute SQL
```

**⑥ JOIN ON SWITCH() — JOIN key là kết quả SWITCH ưu tiên:**

> Dùng khi cần JOIN với key thay đổi tùy điều kiện (ví dụ: 契約者優先順 1→3→2→default).
> Thay vì 4 UPDATE riêng biệt, gộp thành 1 UPDATE với JOIN key là SWITCH.

```vba
SQL = ""
SQL = SQL & "UPDATE [SourceTable] AS T1 "
SQL = SQL & "INNER JOIN [LookupTable] AS T2 "
SQL = SQL & "ON T1.[KeyField] = Switch( "
SQL = SQL & "    (cond_priority1), T2.[Key1], "
SQL = SQL & "    (cond_priority3), T2.[Key3], "
SQL = SQL & "    (cond_priority2), T2.[Key2], "
SQL = SQL & "    True, T2.[Key1] "
SQL = SQL & ") "
SQL = SQL & "SET T1.FLG = '1' "
db.Execute SQL
```

Ví dụ thực tế — 契約者優先順 (1→3→2→default):
```vba
SQL = ""
SQL = SQL & "UPDATE [入居者管理] AS T1 "
SQL = SQL & "INNER JOIN [入居状況一覧] AS T2 "
SQL = SQL & "ON T1.[基幹入居者ID] = Switch( "
SQL = SQL & "    ((T2.[契約者1入居有無] = '入居有り') AND (T2.[契約者2入居有無] = '入居有り') AND (T2.[個人・法人区分_33] = '個人') AND (T2.[個人・法人区分_43] = '個人')), T2.[契約者1No], "
SQL = SQL & "    T2.[契約者3入居有無] = '入居有り', T2.[契約者3No], "
SQL = SQL & "    T2.[契約者2入居有無] = '入居有り', T2.[契約者2No], "
SQL = SQL & "    True, T2.[契約者1No] "
SQL = SQL & ") "
SQL = SQL & "SET T1.[FLG] = '1' "
db.Execute SQL
```

> ⚠️ **LookupTable** (ví dụ: 入居状況一覧) chỉ dùng để JOIN → **KHÔNG AddNewFieldToTable / KHÔNG DeleteFieldInTable** cho bảng này.

---

### C. 通常処理 (INSERT)

```vba
'通常処理
SQL = ""
SQL = SQL & "INSERT INTO [TargetTable] SELECT "
SQL = SQL & "    'Y_' & T.ID as ID, "
SQL = SQL & "    'FixedValue' as field1, "
SQL = SQL & "    T.[SourceField] as field2, "
SQL = SQL & "    'SheetName' as sheet "
SQL = SQL & "FROM [SourceTable] AS T "
SQL = SQL & "WHERE T.FLG = '0';"
db.Execute SQL
log_write "set_[table]:[SheetName] → 通常処理"
```

**Quy tắc INSERT:**
1. KHÔNG liệt kê cột sau tên bảng đích — dùng `as TargetColumn` trong SELECT.
2. BẮT BUỘC field `ID` với prefix: `'[Prefix]_' & T.ID as ID`.
3. BẮT BUỘC field `sheet` = tên sheet_name cố định.
4. Chỉ INSERT field có `ast_type: CONSTANT` hoặc `DIRECT_MAPPING`. 
5. BỎ QUA field `is_active = FALSE` hoặc action = `項目マッピング`hoặc không có giá trị ở cột 処理 và có is_active = TRUE (→ 特殊処理xử lý sau).
6. SWITCH inline trong INSERT:
```vba
SQL = SQL & "    SWITCH( "
SQL = SQL & "        (T.[区分] = '個人'), '10', "
SQL = SQL & "        (T.[区分] = '法人'), '20', "
SQL = SQL & "        True, NULL ) as kind_id, "
```

---

### D. FLG LOGIC NGƯỢC ← **Pattern quan trọng — áp dụng cho IN_LIST_FILTER**

Khi spec **giữ** một danh sách giá trị (không phải loại bỏ):
- Nhận diện: `IN_LIST_FILTER` + `filter_logic: NOT_IN`, hoặc điều kiện dạng `"AまたはBの場合"`
- Logic: reset toàn bộ FLG='1' → set FLG='0' cho record hợp lệ

```vba
'■FLGリセット (logic ngược)
db.Execute "UPDATE [SourceTable] SET [SourceTable].[FLG] = '1';"

SQL = ""
SQL = SQL & "UPDATE [SourceTable] AS T "
SQL = SQL & "SET T.FLG = '0' "
SQL = SQL & "WHERE T.[Field1] = 'Value1' OR T.[Field1] = 'Value2'; "
db.Execute SQL
```

**Phân biệt với `BLANK_CHECK` thông thường:**
| Pattern | FLGリセット ban đầu | Logic lọc |
|---------|-----------------|----------|
| BLANK_CHECK / JOIN_FILTER | Reset về `'0'` | UPDATE SET FLG='1' WHERE ... (xác định record XẤU) |
| IN_LIST_FILTER (NOT_IN) | Reset về `'1'` | UPDATE SET FLG='0' WHERE ... (xác định record TỐT) |

---

### E. 特殊処理 (UPDATE đặc biệt sau INSERT)

**⓪ Bắt buộc comment tóm tắt logic xử lý:**
Trước khi bắt đầu các khối lệnh SQL trong `特殊処理`, bắt buộc phải có các dòng comment (`'`) liệt kê tóm tắt các trường sẽ được update và điều kiện tương ứng (nếu có) dựa trên JSON spec.
Ví dụ:
```vba
    '特殊処理
    'name_family, kind_id, name_family_kana, name_first, name_first_kana -> "個人"の場合
    'company_name, kind_id, company_name_kana -> "法人"の場合
    'tag
    '解約情報にデータがある入居者は退去済
```

**① Chiến lược gom nhóm (Grouping Strategies):**
Mọi quyết định gom nhóm bằng `WHERE` hay `SWITCH` đều phải được cấu hình tại `sample/semantic_dictionary.json` dưới mục `grouping_strategies`:
- `condition_based_where`: Gom nhóm các trường có chung một điều kiện cụ thể (ví dụ: `"個人"の場合`) thành một lệnh `UPDATE ... WHERE`.
- `field_based_switch`: Gom tất cả các điều kiện của một trường (ví dụ: `tag`, `sublease`) thành một lệnh `UPDATE ... SWITCH`.
Tuyệt đối không tự ý quyết định thuật toán gom nhóm mà không tra cứu dictionary. Tương tự, Native Compiler cũng đọc rules này để sinh code tự động.

**① WHERE LIKE / NOT LIKE (`UPDATE_LIKE`):**
```vba
'field_name -> "会社"を含まない場合
SQL = ""
SQL = SQL & "UPDATE [TargetTable] AS T "
SQL = SQL & "INNER JOIN [SourceTable] AS T1 ON T.ID = ('[Prefix]_' & T1.ID) "
SQL = SQL & "SET T.name_family = T1.[貸主名], "
SQL = SQL & "    T.kind_id = '10' "
SQL = SQL & "WHERE T1.[貸主名] NOT LIKE '%会社%'; "
db.Execute SQL
log_write "set_[table]:[SheetName] → ""会社""を含まない場合"
```
> Nhiều field có **cùng WHERE** → gộp vào **1 UPDATE** duy nhất.

**① (mở rộng) Gộp fields cùng WHERE condition — bắt buộc:**

> Khi nhiều fields trong 特殊処理 có **cùng điều kiện WHERE**, PHẢI gộp vào 1 UPDATE SET nhiều field.
> Không được sinh UPDATE riêng cho từng field nếu WHERE condition giống nhau.

```vba
'Đúng: gộp tất cả fields cùng WHERE 個人 vào 1 UPDATE
SQL = ""
SQL = SQL & "UPDATE [TargetTable] AS T "
SQL = SQL & "INNER JOIN [SourceTable] AS T1 ON T.ID = ('[Prefix]_' & T1.ID) "
SQL = SQL & "SET T.name_family = T1.[姓], "
SQL = SQL & "    T.kind_id = '10', "
SQL = SQL & "    T.name_family_kana = T1.[姓 (カナ)], "
SQL = SQL & "    T.name_first = IIF(T1.[名] = '-', NULL, T1.[名]), "
SQL = SQL & "    T.name_first_kana = IIF(T1.[名 (カナ)] = '-', NULL, T1.[名 (カナ)]) "
SQL = SQL & "WHERE T1.[個人／法人] = '個人'; "
db.Execute SQL
log_write "set_[table]:[SheetName] → 特殊処理: 個人フィールド一括更新"

'Sai: sinh UPDATE riêng cho từng field (verbose, không cần thiết)
'SQL ... SET T.name_family = ...  WHERE 個人  → db.Execute SQL
'SQL ... SET T.kind_id = '10'    WHERE 個人  → db.Execute SQL  ← KHÔNG làm vậy
```

**① (mở rộng 2) Hyphen-to-NULL trong SET — dùng `Nz(Replace(...))` thay vì `= '-'`:**

> Khi field nguồn có thể chứa `'-'`, chuỗi rỗng `''`, hoặc `NULL` đều phải xử lý như nhau (→ NULL).
> So sánh trực tiếp `T1.[Field] = '-'` **không bắt được NULL hay chuỗi rỗng** — dùng `Nz(Replace(...))`.

**Quy tắc chọn pattern:**

| Ngữ cảnh | Pattern đúng | Không dùng |
|----------|-------------|-----------|
| Trong SELECT của INSERT (`通常処理`) | `IIF(T.[Field] = '-', NULL, T.[Field])` | `Nz(Replace(...))` — không cần thiết trong SELECT |
| Trong SET của UPDATE (`特殊処理`) | `IIF(Nz(Replace(T1.[Field], '-', ''), '') = '', NULL, T1.[Field])` | `IIF(T1.[Field] = '-', NULL, ...)` |
| Điều kiện SWITCH cho tag/状態フィールド | `Nz(Replace(T1.[タグ], '-', ''), '') = ''` | `T1.[タグ] = '-'` |

**Ví dụ thực tế — tag SWITCH trong 特殊処理:**
```vba
'Sai: chỉ bắt được '-', bỏ sót NULL và chuỗi rỗng
SQL = SQL & "SET T.[tag] = SWITCH( "
SQL = SQL & "    T1.[タグ] = '解約予定', '退去済', "
SQL = SQL & "    T1.[タグ] = '退去済',   '退去済', "
SQL = SQL & "    T1.[タグ] = '-',        '管理移行', "    ' ← SAI
SQL = SQL & "    True, NULL ); "

'Đúng: Nz(Replace(...)) bắt được cả '-', NULL, chuỗi rỗng
SQL = SQL & "SET T.[tag] = SWITCH( "
SQL = SQL & "    T1.[タグ] = '解約予定', '退去済', "
SQL = SQL & "    T1.[タグ] = '退去済',   '退去済', "
SQL = SQL & "    Nz(Replace(T1.[タグ], '-', ''), '') = '', '管理移行', "    ' ← ĐÚNG
SQL = SQL & "    True, NULL ); "
```

> ⚠️ Quy tắc này áp dụng cho **bất kỳ SWITCH/IIF nào trong UPDATE SET** khi điều kiện kiểm tra là giá trị hyphen `'-'` hoặc có thể rỗng. Trong INSERT SELECT, vẫn dùng `IIF(T.[Field] = '-', NULL, T.[Field])` như bình thường.

---

**① (mở rộng 3) end_date / 処理対象日 — Pattern chuẩn:**

> Khi spec đề cập lấy ngày từ "Mainフォームの処理対象日": nguồn thực tế là bảng buffer `T_BUF_DATE`, không phải `Forms!` reference.
> Pattern bắt buộc: Dim VBA Date variable → `DLookup` → `DateSerial(..., 0)` → `Format` → nối string vào SQL.

```vba
'end_date: 処理対象日(yyyymm) → 前月末の日付
Dim targetDate As Date
Dim endOfPreviousMonth As Date
Dim sqlDate As String

targetDate = DLookup("DATE_FROM", "T_BUF_DATE")
endOfPreviousMonth = DateSerial(Year(targetDate), Month(targetDate), 0)
sqlDate = Format(endOfPreviousMonth, "'yyyy-mm-dd'")

SQL = ""
SQL = SQL & "UPDATE [TargetTable] As T "
SQL = SQL & "SET T.[end_date] = " & sqlDate & " "
SQL = SQL & "WHERE T.ID LIKE '[Prefix]_%'; "
db.Execute SQL
log_write "set_[table]:[SheetName] → 特殊処理: end_date"
```

> ⚠️ **KHÔNG dùng** `Forms![Mainフォーム]![処理対象日]` — code chạy batch có thể không có Form mở.
> WHERE dùng `T.ID LIKE '[Prefix]_%'` để giới hạn scope thay vì `WHERE sheet=`.

---

**② SWITCH UPDATE đơn giản (`UPDATE_SWITCH`):**
```vba
SQL = ""
SQL = SQL & "UPDATE [TargetTable] AS T "
SQL = SQL & "INNER JOIN [SourceTable] AS T1 ON T.ID = ('[Prefix]_' & T1.ID) "
SQL = SQL & "SET T.[field] = SWITCH( "
SQL = SQL & "    (condition1), value1, "
SQL = SQL & "    (condition2), value2, "
SQL = SQL & "    True, NULL ); "
db.Execute SQL
log_write "set_[table]:[SheetName] → 特殊処理: [field_name]"
```

**③ Dim con pattern — SWITCH với nhiều branch tái sử dụng (`UPDATE_SWITCH` với nhiều block):**

> Áp dụng khi: nhiều field đều dùng **cùng tập điều kiện** (vd: 契約者優先順 1→3→2→default).
> Khai báo `Dim con` **một lần duy nhất** trước tất cả SWITCH block, dùng lại cho mọi field.

```vba
' Khai báo điều kiện dùng chung (khai báo 1 lần trước tất cả SWITCH block)
Dim con1 As String
Dim con3 As String
Dim con2 As String
con1 = "(T1.[契約者1入居有無] = '入居有り')"
con3 = "(T1.[契約者3入居有無] = '入居有り')"
con2 = "(T1.[契約者2入居有無] = '入居有り')"

' Dùng lại con1/con3/con2 cho từng field
SQL = ""
SQL = SQL & "UPDATE [TargetTable] AS T "
SQL = SQL & "INNER JOIN [SourceTable] AS T1 ON T.ID = ('[Prefix]_' & T1.ID) "
SQL = SQL & "SET T.field1 = SWITCH( "
SQL = SQL & "    " & con1 & ", T1.[SrcField_1], "
SQL = SQL & "    " & con3 & ", T1.[SrcField_3], "
SQL = SQL & "    " & con2 & ", T1.[SrcField_2], "
SQL = SQL & "    True, T1.[SrcField_1] ); "
db.Execute SQL
log_write "set_[table]:[SheetName] → 特殊処理: field1"

SQL = ""
SQL = SQL & "UPDATE [TargetTable] AS T "
SQL = SQL & "INNER JOIN [SourceTable] AS T1 ON T.ID = ('[Prefix]_' & T1.ID) "
SQL = SQL & "SET T.field2 = SWITCH( "
SQL = SQL & "    " & con1 & ", T1.[SrcField2_1], "
SQL = SQL & "    " & con3 & ", T1.[SrcField2_3], "
SQL = SQL & "    " & con2 & ", T1.[SrcField2_2], "
SQL = SQL & "    True, T1.[SrcField2_1] ); "
db.Execute SQL
log_write "set_[table]:[SheetName] → 特殊処理: field2"
```

**④ 契約者優先順 — Pattern chuẩn cho Account (入居者):**

Khi spec có nhiều `契約者N入居有無` với cùng điều kiện lặp → **bắt buộc** dùng Dim con pattern với thứ tự ưu tiên: **契約者1 → 契約者3 → 契約者2 → default**.

`con1` = khớp 契約者1
`con3` = khớp 契約者3
`con2` = khớp 契約者2

Logic xử lý `法人/個人` riêng (ví dụ `name_family` vs `company_name`):
- Field dành cho 個人 (name_family, name_family_kana): thêm branch `(T1.[個人・法人区分] = '法人'), NULL` **trước** True
- Field dành cho 法人 (company_name, company_name_kana): thêm branch `(T1.[個人・法人区分] = '個人'), NULL` **đầu tiên** (trước con1)

**⑤ IIF trong UPDATE SET:**
```vba
SQL = SQL & "SET T.[field] = IIF(Nz(T1.[F1], '') <> '' OR Nz(T1.[F2], '') <> '', '1', NULL); "
```

**⑥ UPDATE_WHERE_EQUALS (1 điều kiện đơn giản):**
```vba
'tag — 解約予定の場合のみセット
SQL = ""
SQL = SQL & "UPDATE [TargetTable] AS T "
SQL = SQL & "INNER JOIN [SourceTable] AS T1 ON T.ID = ('[Prefix]_' & T1.ID) "
SQL = SQL & "SET T.tag = '解約予定' "
SQL = SQL & "WHERE T1.[契約状況] = '解約予定'; "
db.Execute SQL
log_write "set_[table]:[SheetName] → 特殊処理: tag"
```

**⑦ Multi-join UPDATE (3 bảng):**
```vba
SQL = ""
SQL = SQL & "UPDATE ([TargetTable] As T "
SQL = SQL & "INNER JOIN [Table1] AS T1 ON T.ID = ('[Prefix]_' & T1.ID)) "
SQL = SQL & "INNER JOIN [Table2] AS T2 ON T1.[Key] = T2.[Key] "
SQL = SQL & "SET T.[field] = T2.[SourceField] "
SQL = SQL & "WHERE condition; "
db.Execute SQL
```

**⑧ DELETE dedup:**
```vba
SQL = ""
SQL = SQL & "DELETE FROM [TargetTable] "
SQL = SQL & "WHERE ID NOT IN ( "
SQL = SQL & "    SELECT MIN(ID) "
SQL = SQL & "    FROM [TargetTable] "
SQL = SQL & "    GROUP BY [key_field] "
SQL = SQL & ")"
db.Execute SQL
log_write "set_[table]:[SheetName] → 特殊処理"
```

**⑨ Nested SWITCH (chuyển đổi giá trị bên trong SWITCH ngoài):**

Dùng khi field nguồn cần map giá trị (vd: `個人/法人` → `'10'/'20'`), đồng thời bên ngoài có SWITCH chọn nguồn:
```vba
SQL = SQL & "SET T.kind_id = SWITCH( "
SQL = SQL & "    " & con1 & ", SWITCH(T1.[個人・法人区分_33]='個人','10',T1.[個人・法人区分_33]='法人','20',True,'20'), "
SQL = SQL & "    " & con3 & ", SWITCH(T1.[個人・法人区分_53]='個人','10',T1.[個人・法人区分_53]='法人','20',True,'20'), "
SQL = SQL & "    True, SWITCH(T1.[個人・法人区分_33]='個人','10',T1.[個人・法人区分_33]='法人','20',True,'20') ); "
```

**Giá trị kind_id chuẩn:** `'10'` = 個人, `'20'` = 法人 (default = `'20'`).

---

### F. DỌN DẸP FLG (cuối mỗi section)

```vba
'■項目削除
DeleteFieldInTable "[SourceTable]", "FLG"
```

- Tên đúng: `DeleteFieldInTable` (không phải `DeleteFieldFromTable`).
- Nhiều bảng → delete FLG từng bảng riêng.

---

## 6. QUY TRÌNH TỪNG BƯỚC (WORKFLOW)

```
Step 1 — HEADER:
  → log_write "...:in"
  → Dim variables (db, SQL, t, ERR, s_buf, ...)
  → Set db = CurrentProject.Connection
  → DELETE FROM [TargetTable]
  → log_write "[table] delete"
  → DCount guard → Exit Sub nếu rỗng

Step 2 — PER SECTION (repeat cho mỗi sheet nguồn):
  → Section separator comment '====...
  → AddNewFieldToTable + FLGリセット
    ↳ Nếu IN_LIST_FILTER: FLGリセット = '1' (logic ngược)
    ↳ Mặc định: FLGリセット = '0'
  → 出力条件: UPDATE blocks theo ast_type
    ↳ IN_LIST_FILTER → FLG logic ngược (SET FLG='0' cho record hợp lệ)
    ↳ BLANK_CHECK / JOIN_FILTER → SET FLG='1' cho record xấu
    ↳ DEDUP → luôn CUỐI CÙNG
  → log_write "section → 不要行を削除するため(FLG=1更新)"
  → 通常処理: INSERT INTO ... SELECT ... WHERE FLG='0'
  → log_write "section → 通常処理"
  → 特殊処理:
    ↳ Khai báo Dim con (nếu có nhiều SWITCH block cùng điều kiện)
    ↳ UPDATE blocks: LIKE / SWITCH / IIF / WHERE_EQUALS / multi-join / DELETE dedup
    ↳ Mỗi block → 1 log_write "section → 特殊処理: fieldname"
  → DeleteFieldInTable (cleanup FLG)

Step 3 — FOOTER:
  → Debug.Print "[table]:" & Timer - t
  → log_write "...:out"
```

---

## 7. TIÊU CHUẨN CODE (BẮT BUỘC)

| Quy tắc | Đúng | Sai |
|---------|------|-----|
| Kết nối DB | `Set db = CurrentProject.Connection` | `Set db = New ADODB.Connection` |
| Trường JP | `T.[フィールド名]` | `T.フィールド名` |
| Delete FLG | `DeleteFieldInTable "Table", "FLG"` | `DeleteFieldFromTable ...` |
| FLG type | `"TEXT(1)"` | `"TEXT"` |
| log_write format | `"set_[table]:[Section] → [説明]"` | `"set [table]: [section]"` |
| Timer | `Dim t As Single` + `t = Timer` + `Debug.Print` | Thiếu timer |
| DELETE target | Ngay sau `Set db`, trước DCount | Sau DCount hoặc thiếu |
| Sheet field | `'SheetName' as sheet` trong INSERT | Thiếu |
| IN_LIST_FILTER | FLG reset='1' rồi set='0' cho valid | Reset='0' rồi set='1' (ngược) |
| Dim con | Khai báo 1 lần trước tất cả SWITCH block | Khai báo lặp trong từng block |
| 契約者優先順 | con1=契約者1, con3=契約者3, con2=契約者2 | Thứ tự khác |
| Nested SWITCH | Bên trong SWITCH ngoài khi cần map giá trị | Subquery |
| **AddFLG scope** | **Chỉ bảng cần track trạng thái (INSERT/filter)** | **AddFLG cho cả lookup table** |
| **JOIN ON SWITCH** | **JOIN key = SWITCH(cond1→Key1, cond3→Key3, True→Key1)** | **4 UPDATE riêng biệt** |
| **FLG multi-value** | **FLG='1','2','3' cho loại trừ nhiều tầng, INSERT lấy FLG='0'; tầng 2+ BẮT BUỘC WHERE T.[FLG]='0'** | **Binary FLG hoặc thiếu WHERE khi cascading** |
| **Hyphen check trong UPDATE SET** | **`Nz(Replace(T1.[Field], '-', ''), '') = ''`** | **`T1.[Field] = '-'` trong SWITCH/IIF của UPDATE** |
| **Hyphen check trong INSERT SELECT** | **`IIF(T.[Field] = '-', NULL, T.[Field])`** | **`Nz(Replace(...))` trong SELECT** |
| **Gộp WHERE giống nhau** | **1 UPDATE SET nhiều field cùng WHERE condition** | **UPDATE riêng cho từng field** |

---

## 8. LẬP IMPLEMENTATION PLAN & CHỜ DUYỆT (BẮT BUỘC TRƯỚC KHI CODE)

Để tránh việc Agent bỏ sót logic hoặc vi phạm SKILL, **TRƯỚC KHI** sinh ra mã nguồn VBA, Agent BẮT BUỘC phải lập kế hoạch (Implementation Plan) và **CHỜ NGƯỜI DÙNG DUYỆT**.

1. Agent đọc `sheet_raw.json` và toàn bộ `SKILL.md`.
2. **Tự Đặt Câu Hỏi & Đánh Giá Tối Ưu (Production Optimization)**:
   - Agent **phải tự đặt câu hỏi và so sánh** xem phương pháp sinh SQL nào là **tối ưu nhất cho môi trường Production** (Access DB). 
   - *Ví dụ:* Dùng 1 lệnh `UPDATE` gộp chứa `SWITCH` sẽ tối ưu (chỉ quét bảng 1 lần) và tránh lỗi index (non-sargable) đối với các hàm phức tạp như `Nz(Replace(...))` so với việc dùng nhiều lệnh `UPDATE ... WHERE`.
   - Agent phải chủ động **tìm kiếm từ khóa (search/grep)** trong file SKILL để map đúng pattern (VD: search từ khóa "tag", "SWITCH", "FLG") để tránh nhầm lẫn.
3. Agent tạo bảng phân tích chi tiết (bằng tiếng Việt) trình bày:
   - File chính/phụ cần JOIN.
   - Các điều kiện loại trừ (Sử dụng tầng FLG nào: 1, 2, 3?).
   - Quyết định thiết kế (Tại sao dùng SWITCH thay vì WHERE? Tại sao dùng NOT EXISTS?).
   - Các khối UPDATE gộp (gom fields cùng WHERE hoặc gộp bằng SWITCH).
   - Mapping đặc biệt (SWITCH, Hyphen-to-NULL).
4. Agent dừng lại và hỏi: "Bản kế hoạch này đã đúng ý bạn chưa? Vui lòng duyệt (Approve) để tôi bắt đầu sinh code."
5. **TUYỆT ĐỐI KHÔNG** dùng tool `write_to_file` để sinh file `.bas` cho đến khi người dùng phản hồi "Duyệt" hoặc đồng ý.

---

## 9. CHECKLIST TRƯỚC KHI XUẤT OUTPUT (BẮT BUỘC)

Trước khi generate mã nguồn VBA cuối cùng, Agent (bạn) phải tự kiểm tra lại các lỗi phổ biến sau:

1. **[ ] Kiểm tra kết nối DB**: BẮT BUỘC khai báo `Dim db As ADODB.Connection` và `Set db = CurrentProject.Connection`. (TUYỆT ĐỐI KHÔNG dùng `DAO.Database` / `CurrentDb`).
2. **[ ] Kiểm tra SQL Syntax & Khoảng trắng**: Mọi chuỗi SQL nối nhau `SQL = SQL & "..."` phải chắc chắn có dư một khoảng trắng ở cuối (tránh dính chữ như `as aliasFROM`).
3. **[ ] Ngoặc vuông cho tiếng Nhật**: Tên bảng và tên cột chứa tiếng Nhật phải luôn được bọc trong `[...]`.
4. **[ ] Chuẩn hoá Tên Bảng (Ngoại lệ)**: Nếu spec nhắc đến bảng `GMO 入居状況一覧`, CHỈ sử dụng tên `[入居状況一覧]`.
5. **[ ] Dấu nháy trong điều kiện String**: Kiểm tra kỹ các lỗi typo dấu nháy do copy paste (ví dụ: `'入居有り"場合'` là sai cú pháp, phải là `'入居有り'`).
6. **[ ] Vị trí Dedup**: Block `重複チェック` phải luôn được đặt Ở CUỐI CÙNG của phần `出力条件`.
7. **[ ] Rà soát toàn bộ File**: Kiểm tra lại toàn bộ các file được nhắc đến trong JSON (các cột `使用ファイル`, `該当ファイル名`). Đảm bảo không bỏ sót bất kỳ file/bảng nào (phải được xử lý qua JOIN, xuất ra, hoặc comment rõ `' (No action required)` nếu không cần thao tác).
8. **[ ] TUYỆT ĐỐI BÁM SÁT JSON (Không tự suy diễn)**: Mã nguồn sinh ra đã map chính xác 100% với `sheet_raw.json` (từ điều kiện lọc, JOIN, cho đến cột tương ứng) hay chưa? Không được tự ý đoán logic nghiệp vụ nếu không có định nghĩa rõ ràng. Mọi điều kiện (như "紐づかない場合") phải được dịch sát nghĩa đen thành SQL (ví dụ `LEFT JOIN ... WHERE T2.ID IS NULL`).

---

## 10. ĐỊNH DẠNG ĐẦU RA (BẮT BUỘC THEO THỨ TỰ)

**TRƯỚC KHI** xuất mã nguồn VBA, bạn **BẮT BUỘC** phải xuất ra Checklist và Chứng minh theo định dạng sau:

```markdown
### ✅ CHECKLIST & PROOF OF JSON ADHERENCE

1. **[x] Kết nối DB**: Đã dùng ADODB.Connection và CurrentProject.Connection.
2. **[x] SQL Syntax**: Đã kiểm tra khoảng trắng cuối mỗi dòng SQL.
3. **[x] Ngoặc vuông tiếng Nhật**: Đã bọc `[...]` cho tất cả table/field tiếng Nhật.
4. **[x] Tên bảng GMO**: Đã loại bỏ chữ "GMO " khỏi `[入居状況一覧]`.
5. **[x] Dấu nháy String**: Không có lỗi typo `"場合'`.
6. **[x] Vị trí Dedup**: Block `重複チェック` nằm cuối `出力条件`.
7. **[x] Rà soát File**: [Liệt kê các bảng đã lấy từ JSON và cách xử lý tương ứng]
8. **[x] Bám sát JSON 100%**: 
   - *Chứng minh cụ thể bằng cách trích dẫn các phần đặc biệt hoặc khó trong JSON và cách bạn đã chuyển đổi chính xác sang SQL mà không tự suy diễn.* (Đặc biệt chứng minh các đoạn "紐づかない場合" hay các JOIN key gộp bằng dấu `-`).
```

**SAU ĐÓ**, mới được phép trả về khối mã VBA duy nhất:

- Cung cấp mã VBA trong một Markdown code block (` ```vba ... ``` `).
- Đặt comment `' -- NEEDS BA REVIEW` cho field có `needs_review = true`.
- Tuân thủ tuyệt đối phong cách template đã cung cấp.

---

## 11. CHANGELOG

| Version | Ngày | Thay đổi |
|---------|------|---------|
| v1.0 | — | Phiên bản gốc |
| v2.0 | 2026-05-29 | Thêm: IN_LIST_FILTER + FLG Logic Ngược, Dim con pattern, 契約者優先順 chuẩn, Nested SWITCH, UPDATE_WHERE_EQUALS, bảng ast_type→pattern, phân biệt FLG reset logic |
| v3.0 | 2026-06-05 | Thêm: JOIN ON SWITCH() pattern (join key là kết quả SWITCH ưu tiên), FLG Multi-value (phân tầng '1'/'2'/'3'), Rule AddFLG scope (chỉ bảng track trạng thái, không thêm cho lookup table), Rule gộp fields cùng WHERE condition vào 1 UPDATE |
| v4.1 | 2026-06-08 | Thêm: Pattern end_date/処理対象日 — dùng DLookup("DATE_FROM","T_BUF_DATE") thay vì Forms! reference; WHERE dùng ID LIKE prefix thay vì sheet= |
| v5.0 | 2026-06-08 | Thêm: Rule BẮT BUỘC lập Implementation Plan và chờ duyệt trước khi generate code |
| v5.1 | 2026-06-08 | Thêm: Yêu cầu Agent tự đặt câu hỏi đánh giá hiệu năng Production (VD: SWITCH vs WHERE) và chủ động search từ khóa trong SKILL.md khi lên Plan |