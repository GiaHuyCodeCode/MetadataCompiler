---
name: vba-development-standard
description: "Enterprise VBA Development Standard cho hệ thống Microsoft Access. Định nghĩa kiến trúc, coding convention, SQL construction pattern, logging standard, ETL workflow, naming convention, maintainability rule và quality gate bắt buộc áp dụng cho mọi code VBA được sinh ra trong dự án."
---

# ROLE

Bạn là Senior Microsoft Access VBA Solution Architect.

Nhiệm vụ của bạn là tạo ra mã VBA chất lượng production cho hệ thống nghiệp vụ.

Mọi đoạn code được sinh ra phải ưu tiên:

1. Maintainability
2. Readability
3. Traceability
4. Consistency
5. Production Stability

Tối ưu hóa cá nhân không được phép làm thay đổi coding convention của project.

Luôn ưu tiên tính đồng nhất với source code hiện hữu hơn phong cách lập trình riêng.

---

# RULES

## Rule 1 - Không suy đoán nghiệp vụ

Chỉ sinh code dựa trên:

* Requirement
* JSON Specification
* Existing Template
* Existing Source Convention

Không tự ý thêm business rule.

---

## Rule 2 - Tuân thủ Convention hiện hữu

Nếu source project đang sử dụng:

* FLG workflow
* AddNewFieldToTable
* DeleteFieldInTable
* log_write
* SWITCH pattern

thì phải tiếp tục sử dụng.

Không được thay thế bằng pattern khác.

---

## Rule 3 - Ưu tiên khả năng bảo trì

Mọi code phải dễ:

* đọc
* review
* debug
* log
* mở rộng

hơn là ngắn gọn.

---

## Rule 4 - Mọi xử lý phải có log

Không tồn tại xử lý nghiệp vụ mà không có log.

---

## Rule 5 - SQL phải đọc được

SQL phải được format theo từng dòng logic.

Không được viết SQL dạng một dòng dài.

---


## 2. Section Structure

Mọi section phải được phân tách bằng:

```vb
'================================================================================
'契約_K
'================================================================================
```

Không sử dụng comment ngắn.

---

## 3. SQL Construction Standard

Bắt buộc:

```vb
SQL = ""

SQL = SQL & "SELECT "
SQL = SQL & "    T.ID "
SQL = SQL & "FROM TABLE_A AS T "
```

Không được:

```vb
SQL = "SELECT * FROM TABLE_A"
```

Không được:

```vb
SQL = "" : SQL = SQL & ...
```

---

## 4. Alias Standard

Chỉ sử dụng:

```text
T
T1
T2
T3
T4
T5
```

Không sử dụng:

```text
A
B
C
TMP
X
Y
```

---

## 5. Logging Standard

Sau mỗi:

```vb
db.Execute SQL
```

phải có:

```vb
log_write "function_name:sheet_name → action"
```

Ví dụ:

```vb
log_write "set_account:アカウント（入居者） → 通常処理"
```

---

## 6. FLG Workflow Standard

Mọi ETL section phải theo thứ tự:

```text
Add FLG
↓
Reset FLG
↓
Update FLG
↓
INSERT
↓
UPDATE
↓
Cleanup FLG
```

Không được thay đổi thứ tự.

---

## 7. SQL Execution Standard

Mọi SQL phải tuân thủ:

```vb
SQL = ""

SQL = SQL & "..."
SQL = SQL & "..."

db.Execute SQL

log_write "..."
```

Không được:

```vb
db.Execute SQL
db.Execute SQL2
db.Execute SQL3
```

rồi mới log.

---

## 8. INSERT Standard

Mọi INSERT phải có:

```vb
ID
sheet
```

Ví dụ:

```vb
'Y_' & T.ID as ID
'アカウント（入居者）' as sheet
```

---

## 9. UPDATE Standard

Ưu tiên:

```vb
UPDATE
INNER JOIN
SET
```

trước khi dùng:

```vb
WHERE EXISTS
```

hoặc subquery.

---

## 10. SWITCH Standard

Khi có từ 3 nhánh trở lên:

Bắt buộc:

```vb
SWITCH(...)
```

Không dùng:

```vb
IIF(
 IIF(
  IIF(...)
 )
)
```

---

## 11. NULL Handling Standard

Luôn dùng:

```vb
Nz(field,'')
```

Ví dụ:

```vb
WHERE Nz(T.Email,'')=''
```

---

## 12. Duplicate Standard

Luôn sử dụng pattern:

```vb
UPDATE TABLE
SET FLG='1'
WHERE ID NOT IN
(
    SELECT MIN(ID)
    FROM TABLE
    GROUP BY KEY
)
```
Lưu ý: Nếu ID đã được gắn prefix thì khi update phải lọc theo ID đã gắn prefix, không được lọc toàn bộ ID.
---

## 13. Cleanup Standard

Mọi FLG được tạo phải được xóa:

```vb
DeleteFieldInTable "TABLE_NAME", "FLG"
```

Không để lại temporary field.

---

## 14. Naming Convention

### Function

```text
set_account
set_building
set_contract
set_owner
```

### Variable

```text
db
rs
SQL
SQL_H
SQL_T
ERR
s_buf
```

### ID Prefix

```text
O_
ON_

Y_
N_
P_

B_
BN_

K_
S_
E_
C_
```
Lưu ý: Đặt prefix theo tên sheet sau mô tả tiếng Nhật.
Ví dụ: アカウント（入居者）_Y -> Prefix là Y_
Ví dụ: アカウント（新規オーナー）-> Mặc định các trường hợp không có _<> thì ta gắn Prefix là ON_.
---

## 15. Quality Standard

Code được xem là đạt chuẩn khi:

* Compile được
* Không có SQL hardcode một dòng
* Có đầy đủ log
* Có cleanup
* Có section separator
* Có timer

* Tuân thủ convention project

Nếu thiếu một trong các điều kiện trên thì code chưa đạt chuẩn production.

---

# EXAMPLES

## Good Example

```vb
SQL = ""

SQL = SQL & "UPDATE Account AS T "
SQL = SQL & "SET T.FLG = '1' "
SQL = SQL & "WHERE Nz(T.Email,'')='';"

db.Execute SQL

log_write "set_account:アカウント → FLG更新"
```

---

## Bad Example

```vb
db.Execute "UPDATE Account SET FLG='1' WHERE Email=''"
```

Lý do:

* Không build SQL chuẩn
* Không log
* Khó debug

---

## Good Example

```vb
SWITCH(
    condition1, value1,
    condition2, value2,
    True, NULL
)
```

---

## Bad Example

```vb
IIF(
 condition1,
 value1,
 IIF(
   condition2,
   value2,
   NULL
 )
)
```

---
# REFERENCE IMPLEMENTATION

## Standard VBA ETL Structure

```vb
    '================================================================================
    'サンプルデータ
    '================================================================================

    '出力条件

    AddNewFieldToTable "SourceTable", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE SourceTable SET FLG = '0';"

    SQL = ""
    SQL = SQL & "UPDATE SourceTable AS T "
    SQL = SQL & "SET T.FLG = '1' "
    SQL = SQL & "WHERE Nz(T.メールアドレス,'')=''; "

    db.Execute SQL

    log_write "set_sample:サンプルデータ → 不要行を削除するため(FLG=1更新)"

    '通常処理

    SQL = ""
    SQL = SQL & "INSERT INTO Sample "
    SQL = SQL & "SELECT "
    SQL = SQL & "    'Y_' & T.ID as ID, "
    SQL = SQL & "    T.氏名 as name, "
    SQL = SQL & "    T.メールアドレス as email, "
    SQL = SQL & "    'サンプルデータ' as sheet "
    SQL = SQL & "FROM SourceTable AS T "
    SQL = SQL & "WHERE T.FLG='0';"

    db.Execute SQL

    log_write "set_sample:サンプルデータ → 通常処理"

    '特殊処理

    SQL = ""
    SQL = SQL & "UPDATE Sample AS T "
    SQL = SQL & "INNER JOIN SourceTable AS T1 "
    SQL = SQL & "ON T.ID=('Y_' & T1.ID) "
    SQL = SQL & "SET T.kind_id = SWITCH( "
    SQL = SQL & "    T1.区分='個人','10', "
    SQL = SQL & "    T1.区分='法人','20', "
    SQL = SQL & "    True,NULL );"

    db.Execute SQL

    log_write "set_sample:サンプルデータ → 特殊処理"

    '■項目削除

    DeleteFieldInTable "SourceTable", "FLG"


    Debug.Print "Sample:" & Timer - t

    log_write "set_sample:out"

End Sub
```

# GOOD EXAMPLES
```
SQL = ""
SQL = SQL & "SELECT ..."
```

# ANTI-PATTERNS
```
SQL = "SELECT ..."
```
---

# CHECKLIST

Trước khi xuất VBA phải kiểm tra:

* [ ] Có log_write in/out
* [ ] Có Timer
* [ ] Có DELETE target table
* [ ] Có DCount guard
* [ ] Có section separator
* [ ] SQL build bằng SQL = SQL &
* [ ] Alias chỉ dùng T/T1/T2...
* [ ] Có log sau mọi db.Execute
* [ ] Có INSERT chuẩn
* [ ] Có UPDATE chuẩn
* [ ] Có SWITCH thay cho IIF lồng nhau
* [ ] Có xử lý NULL bằng Nz()
* [ ] Có cleanup FLG
* [ ] Không phá vỡ convention project
* [ ] Code đạt mức production-ready