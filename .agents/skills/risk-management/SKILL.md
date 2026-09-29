---
name: risk-management
description: >
  Buộc agent phân tích rủi ro và bao phủ TOÀN BỘ edge case trước khi sinh code.
  Kích hoạt khi viết bất kỳ Sub/Function VBA nào có thể bị user khai thác bởi thao tác
  bất thường (input sai, click sai thứ tự, file thiếu, NULL, encoding lỗi, v.v.).
  Tham chiếu: OWASP Agentic Skills Top 10 (AST10), NIST AI RMF, ISO 31000.
metadata:
  version: "1.0"
  scope: VBA / Microsoft Access / Python ETL Pipeline
  trigger: "Bắt buộc kích hoạt MỌI LÚC trước khi sinh code trong workspace này"
---

# Risk Management — Code-Level Risk Analysis & Defensive VBA Engineering

> **Nguyên tắc tối thượng**: Một đoạn code chỉ "chạy được" KHÔNG ĐỦ.
> Code phải **không sụp đổ** khi user thực hiện bất kỳ thao tác hợp lệ lẫn bất thường nào.
> Agent **BẮT BUỘC** chạy qua Risk Matrix dưới đây cho TỪNG Sub/Function trước khi viết code.

---

## 1. Root Cause — Tại Sao Code Agent Sinh Code Thiếu An Toàn

Code agent thường sinh code theo **happy path** — tức là ngầm giả định:
- File luôn tồn tại và đúng format
- User luôn thực hiện đúng thứ tự bước
- Data không bao giờ NULL hoặc rỗng
- Encoding luôn là Shift-JIS hoặc UTF-8 sạch
- Connection đến CSDL luôn thành công

**Thực tế**: User bấm nút khi chưa chọn file → crash. User bấm "Chạy" 2 lần → duplicate data. File CSV thiếu dòng header → Type Mismatch. Connection timeout → unhandled error.

Đây là các rủi ro **có thể dự đoán được** và **BẮT BUỘC phải được cover** ngay trong lần sinh code đầu tiên.

---

## 2. The 5-Axis Risk Matrix — Bắt Buộc Trước Mỗi Code Block

Trước khi viết bất kỳ Sub/Function nào, agent phải chạy qua **5 trục phân tích rủi ro** và output kết quả vào `walkthrough.md`:

```
RISK ANALYSIS FOR: <TênSub/Function>
═══════════════════════════════════════════════════════

[AXIS 1] USER BEHAVIOR RISKS — Thao tác bất thường của người dùng
  □ Bấm nút khi pipeline chưa sẵn sàng (file chưa chọn, kết nối chưa mở)
  □ Bấm cùng một nút hai lần liên tiếp (double-click, rapid re-run)
  □ Thoát khỏi Form giữa chừng khi đang xử lý
  □ Nhập giá trị sai kiểu dữ liệu (chữ vào ô số, ngày sai format)
  □ Chọn file sai loại (.xlsx khi cần .csv, file đang mở bởi Excel)
  □ Không chọn gì cả và vẫn nhấn OK/Run

[AXIS 2] DATA INTEGRITY RISKS — Dữ liệu đầu vào bất thường
  □ File CSV/Excel hoàn toàn rỗng (0 dòng dữ liệu, chỉ có header)
  □ File thiếu cột bắt buộc hoặc cột bị đổi tên
  □ Ô dữ liệu là NULL / empty string / chỉ có khoảng trắng
  □ Ngày tháng sai format (dd/mm vs yyyy/mm/dd vs text "令和3年")
  □ Số âm hoặc số vượt giới hạn kiểu dữ liệu
  □ Ký tự đặc biệt trong tên file/thư mục (ngoặc, tiếng Nhật, space)
  □ File encoding sai (UTF-8 BOM khi cần Shift-JIS)
  □ Dữ liệu trùng lặp vi phạm ràng buộc unique

[AXIS 3] SYSTEM & ENVIRONMENT RISKS — Môi trường thực thi bất ổn
  □ Đường dẫn file/thư mục không tồn tại tại thời điểm thực thi
  □ File đang bị khóa bởi tiến trình khác (Excel, Access đang dùng)
  □ Không đủ quyền đọc/ghi (permission denied)
  □ Ổ đĩa đầy (disk full) trong khi đang ghi file/log
  □ CSDL Access bị khóa bởi user khác (multi-user lock)
  □ External COM object (Excel.Application) bị crash → ghost process
  □ Runtime timeout khi xử lý file lớn (>50,000 dòng)

[AXIS 4] PIPELINE STATE RISKS — Trạng thái pipeline không nhất quán
  □ Bảng buffer chưa được clear từ lần chạy trước (re-run scenario)
  □ FLG column chưa được thêm (AddNewFieldToTable chưa chạy)
  □ Sub trước đó bị crash → pipeline ở trạng thái nửa chừng
  □ Transaction chưa CommitTrans khi bị interrupt → dữ liệu không nhất quán
  □ Log table bị đầy hoặc bị khóa → log_write thất bại âm thầm

[AXIS 5] RESOURCE LEAK RISKS — Rò rỉ tài nguyên
  □ ADODB.Recordset chưa Close + Set = Nothing khi xảy ra lỗi
  □ Excel.Application/Workbook chưa Quit + Nothing → ghost EXCEL.EXE
  □ FileSystemObject/TextStream chưa Close → file lock
  □ DAO.Database chưa Close → ACE Engine lock
  □ DoCmd.Hourglass True không có Finally để reset về False
```

---

## 3. Defensive VBA Code Patterns — Bắt Buộc Áp Dụng

### Pattern 1: Guard Clause (Kiểm tra tiền điều kiện sớm)

```vba
' DUNG — Guard Clause: thoat som neu dieu kien chua du
Public Sub btn_run_Click()
    ' [GUARD-1] Kiem tra file da chon chua
    If Len(Trim(Me.txt_file_path)) = 0 Then
        MsgBox "ファイルを選択してください。", vbExclamation, "入力エラー"
        Me.txt_file_path.SetFocus
        Exit Sub
    End If
    
    ' [GUARD-2] Kiem tra file co ton tai khong
    If Not FileExists(Me.txt_file_path) Then
        MsgBox "指定されたファイルが見つかりません:" & vbCrLf & Me.txt_file_path, vbCritical, "ファイルエラー"
        Exit Sub
    End If
    
    ' [GUARD-3] Ngan double-run: disable nut trong khi dang chay
    Me.btn_run.Enabled = False
    DoCmd.Hourglass True
    
    On Error GoTo ErrHandler
    ' ... logic chinh ...
    
Cleanup:
    Me.btn_run.Enabled = True
    DoCmd.Hourglass False
    Exit Sub
ErrHandler:
    log_write "btn_run_Click ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    Resume Cleanup
End Sub
```

### Pattern 2: Safe NULL Handling

```vba
' An toan voi NULL va Empty
Function safe_string(val As Variant) As String
    If IsNull(val) Or IsEmpty(val) Then
        safe_string = ""
    Else
        safe_string = Trim(CStr(val))
    End If
End Function

Function safe_date(val As Variant) As String
    If IsNull(val) Or IsEmpty(val) Or Not IsDate(val) Then
        safe_date = ""
    Else
        safe_date = Format(val, "yyyy-mm-dd")
    End If
End Function

Function safe_long(val As Variant, defaultVal As Long) As Long
    If IsNull(val) Or IsEmpty(val) Or Not IsNumeric(val) Then
        safe_long = defaultVal
    Else
        safe_long = CLng(val)
    End If
End Function
```

### Pattern 3: Resource Lock & Release

```vba
' Moi COM Object deu duoc giai phong trong ca happy path lan error path
Public Sub import_excel_file(filePath As String)
    Dim xlApp As Object
    Dim xlWb  As Object
    Dim xlWs  As Object
    
    On Error GoTo ErrHandler
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    Set xlWb = xlApp.Workbooks.Open(filePath, ReadOnly:=True)
    Set xlWs = xlWb.Worksheets(1)
    
    ' ... xu ly du lieu ...
    
Cleanup:
    ' BAT BUOC: Giai phong theo thu tu nguoc lai (Ws -> Wb -> App)
    If Not xlWs Is Nothing Then Set xlWs = Nothing
    If Not xlWb Is Nothing Then
        xlWb.Close SaveChanges:=False
        Set xlWb = Nothing
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    Exit Sub

ErrHandler:
    log_write "import_excel_file ERROR [" & filePath & "]: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Sub
```

### Pattern 4: Pipeline State Check

```vba
' Kiem tra trang thai pipeline truoc khi chay de xu ly re-run
Public Sub set_account()
    Dim db       As DAO.Database
    Dim lngCount As Long
    
    On Error GoTo ErrHandler
    Set db = CurrentDb()
    
    ' [STATE-CHECK] Bang dich phai rong truoc khi INSERT
    lngCount = DCount("*", "Account")
    If lngCount > 0 Then
        Dim answer As Integer
        answer = MsgBox("Accountテーブルに" & lngCount & "件のデータが既に存在します。" & vbCrLf & _
                        "上書きしますか？", vbYesNo + vbQuestion, "確認")
        If answer = vbNo Then Exit Sub
        db.Execute "DELETE FROM Account", dbFailOnError
        log_write "set_account: Account cleared (re-run detected)"
    End If
    
    ' ... tiep tuc pipeline ...
```

### Pattern 5: Transaction Rollback

```vba
' Wrap toan bo batch trong transaction de rollback neu co loi giua chung
Public Sub batch_insert_safe()
    Dim db              As DAO.Database
    Dim blnInTransaction As Boolean
    
    On Error GoTo ErrHandler
    Set db = CurrentDb()
    
    db.BeginTrans
    blnInTransaction = True
    log_write "batch_insert_safe: Transaction BEGIN"
    
    db.Execute "INSERT INTO ...", dbFailOnError
    db.Execute "UPDATE ...", dbFailOnError
    
    db.CommitTrans
    blnInTransaction = False
    log_write "batch_insert_safe: Transaction COMMIT"
    
Cleanup:
    If blnInTransaction Then
        db.RollbackTrans
        log_write "batch_insert_safe: Transaction ROLLBACK (error recovery)"
    End If
    If Not db Is Nothing Then Set db = Nothing
    Exit Sub

ErrHandler:
    log_write "batch_insert_safe ERROR: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Sub
```

---

## 4. Risk-Weighted Priority Table

| Rủi ro | Xác suất | Tác động | Hành động bắt buộc |
|--------|----------|----------|--------------------|
| NULL data gây Type Mismatch | Rất cao | Crash | safe_* functions BẮT BUỘC |
| File lock → ghost EXCEL.EXE | Cao | Process leak | Cleanup pattern BẮT BUỘC |
| Double-click nút Run | Cao | Duplicate data | Disable button BẮT BUỘC |
| File không tồn tại | Cao | Crash | FileExists guard BẮT BUỘC |
| Encoding mismatch Shift-JIS | Trung bình | Corrupt data | ADODB.Stream charset |
| Transaction interrupt | Trung bình | Corrupt DB | BeginTrans/Rollback |
| Re-run trên data cũ | Trung bình | Duplicate data | State check trước INSERT |
| Disk full khi ghi log | Thấp | Log miss | On Error Resume Next cho log |

---

## 5. Self-Review Checklist — Sau Khi Sinh Code

```
CHO TUNG Sub hoac Function duoc sinh ra:

INPUT VALIDATION:
  [ ] Moi tham so dau vao deu duoc kiem tra NULL/Empty truoc khi dung?
  [ ] Moi duong dan file deu duoc kiem tra FileExists() truoc khi mo?
  [ ] Moi ten bang/field deu ton tai truoc khi query?

ERROR HANDLING:
  [ ] Co "On Error GoTo ErrHandler" o dau moi Sub/Function?
  [ ] ErrHandler co goi log_write voi Err.Number va Err.Description?
  [ ] ErrHandler co Resume ve Cleanup label de giai phong tai nguyen?
  [ ] KHONG co "On Error Resume Next" khong co ly do ro rang?

RESOURCE MANAGEMENT:
  [ ] Moi Set obj = CreateObject() deu co Set obj = Nothing trong Cleanup?
  [ ] Moi xlWb.Open deu co xlWb.Close trong Cleanup?
  [ ] xlApp.Quit duoc goi truoc Set xlApp = Nothing?
  [ ] DoCmd.Hourglass reset ve False trong Cleanup?
  [ ] Me.btn_xxx.Enabled reset ve True trong Cleanup?

STATE MANAGEMENT:
  [ ] Neu re-runnable: kiem tra va xu ly trang thai tu lan chay truoc?
  [ ] Buffer table duoc clear_table() truoc khi bat dau?
  [ ] FLG column duoc AddNewFieldToTable() truoc khi dung trong WHERE?

NULL/TYPE SAFETY:
  [ ] Moi Recordset field access deu wrap trong safe_string/safe_date/safe_long?
  [ ] CStr/CLng/CDate deu co IsNull/IsNumeric/IsDate guard truoc?
  [ ] DLookup() luon kiem tra IsNull() tren gia tri tra ve?
```

---

## 6. Mandatory Risk Statement — Bắt Kèm Theo Mỗi Code Block

```markdown
## RISK STATEMENT for `<TenSub>`

### Risks Covered (da phong ve):
- [AXIS-1] Double-click: Button disabled trong khi chay
- [AXIS-2] NULL date field: safe_date() tra ve empty string
- [AXIS-3] File bi khoa: FileExists + On Error handler
- [AXIS-5] Ghost Excel: xlApp.Quit trong Cleanup label

### Risks Accepted (chap nhan rui ro):
- [AXIS-3] Disk full: Log co the that bai, pipeline van tiep tuc (muc do thap)

### THINGS I DIDN'T TOUCH (intentionally):
- Sub clear_table() — khong sua, khong lien quan
- Sub log_write() — khong sua, da on dinh
```

---

## 7. Khi Nào Kích Hoạt Skill Này

Skill `risk-management` PHẢI được kích hoạt **trước Bước 4 (viết code)** trong pipeline 4 bước:

```
interview-me -> solution-architecture-council -> idea-refine
                                                      |
                                          [risk-management] <-- TAI DAY
                                          (Risk Analysis truoc khi code)
                                                      |
                                          using-agent-skills -> code
```

**Trigger tự động** khi prompt chứa bất kỳ từ khóa nào:
- `viết code`, `tạo Sub`, `tạo Function`, `implement`, `fix bug`
- `set_XXX`, `import_XXX`, `export_XXX`, `clear_XXX`
- `INSERT INTO`, `UPDATE`, `DELETE`
- Bất kỳ task nào thao tác với file I/O, ADODB, DAO, Excel COM

---

## 8. Nguồn Tham Khảo

| Framework | Ứng dụng |
|-----------|----------|
| OWASP Agentic Skills Top 10 (AST10) | Lethal Trifecta — data + execution + external access |
| OWASP Top 10 for Agentic Applications (2025) | Privilege abuse, goal hijacking → Guard Clause |
| NIST AI RMF | Lifecycle governance, robustness, transparency |
| ISO 31000 Risk Management | Risk Matrix: Probability × Impact → Action |
| Defensive Programming (NASA/JPL Standard) | "Code as if every assumption will be wrong" |
| Microsoft Access Best Practices | DAO transaction, COM object lifecycle |
