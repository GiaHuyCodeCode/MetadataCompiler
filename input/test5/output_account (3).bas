Attribute VB_Name = "output_account"
Option Compare Database
Option Explicit

Sub set_account()
    
    log_write "set_account:in"
    
    Dim t As Single
    t = Timer
    
    Dim db As ADODB.Connection
    Dim rs1 As ADODB.Recordset
    Dim SQL1 As String
    Dim SQL As String
    Dim i As Long
    
    Dim ERR As Long
    Dim s_buf As String
    Dim s_tel As String
    Dim s_cell As String
    Dim frmName As String
    Dim birth As String
    
    Dim FLG As Boolean
    
    frmName = "F_Main"
    
    Set db = CurrentProject.Connection
    
    db.Execute "DELETE FROM Account"
    log_write "set_account:Account delete"
    
    FLG = False
    
    'Owner
    If DCount("*", "家主基本情報") <= 0 Then
        log_write "set_account:Account_Owner_0件終了"
        
    Else
        
        '■家主基本情報設定=============================================================Y
        SQL = ""
        SQL = SQL & "INSERT INTO Account SELECT "
        SQL = SQL & "家主基本情報.ID , "
        SQL = SQL & """Y"" & ID as 識別ID, "
        SQL = SQL & """Owner"" as klass, "
        SQL = SQL & "iif(家主区分 = ""個人"",家主名,"""") as name_family, "
        SQL = SQL & "iif(家主区分 = ""個人"",家主名カナ,"""") as name_family_kana, "
        SQL = SQL & "優先メール as email, "
        SQL = SQL & "Replace(Nz(TEL1,""""),""-"","""",1,-1,1) as tel_fixed, "
        SQL = SQL & "Replace(Nz(携帯1,""""),""-"","""",1,-1,1) as tel_mobile, "
        SQL = SQL & "Replace(Nz(郵便番号,""""),""-"","""",1,-1,1) as zip_code, "
        SQL = SQL & "住所1 as prefecture_code, "
        SQL = SQL & "住所2 as address_2, "
        SQL = SQL & "iif(家主区分 = ""法人"",家主名,"""") as company_name, "
        SQL = SQL & "iif(家主区分 = ""法人"",家主名カナ,"""") as company_name_kana, "
        SQL = SQL & "家主区分 as kind_id, "
        SQL = SQL & "性別 as gender_id, "
        SQL = SQL & "format(生年月日,""yyyy-mm-dd"") as birthday, "
        SQL = SQL & "[家主No] as legacy_id "
        SQL = SQL & "FROM 家主基本情報;"
        db.Execute SQL
        log_write "set_account:家主基本情報設定"
        
        
        '■kind_idの設定(項目マッピング)
        SQL = ""
        SQL = SQL & "UPDATE Account "
        SQL = SQL & "LEFT JOIN T_MAP_ACC_KIND_ID "
        SQL = SQL & "ON Account.kind_id = "
        SQL = SQL & "T_MAP_ACC_KIND_ID.ORG "
        SQL = SQL & "SET Account.kind_id = "
        SQL = SQL & "[T_MAP_ACC_KIND_ID]![NEW];"
        db.Execute SQL
        log_write "set_account:kind_id_Oの設定"
        
        
        '■gender_idの設定(項目マッピング)
        SQL = ""
        SQL = SQL & "UPDATE Account "
        SQL = SQL & "LEFT JOIN T_MAP_ACC_GENDER_ID "
        SQL = SQL & "ON Account.gender_id = "
        SQL = SQL & "T_MAP_ACC_GENDER_ID.ORG "
        SQL = SQL & "SET Account.gender_id = "
        SQL = SQL & "[T_MAP_ACC_GENDER_ID]![NEW];"
        db.Execute SQL
        log_write "set_account:gender_id_Oの設定"
        
        
         FLG = True
    
    End If
    
    
    'Resident
    If DCount("*", "契約者情報") <= 0 Or DCount("*", "入居状況一覧") <= 0 Then
        log_write "set_account:Account_Resident_0件終了"
        
        If FLG = True Then
            '進捗状況の更新~~~~~~~~~~~~~~~~~~
            DoEvents
            Call UpdateProgressBar(frmName)
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        End If
        
        Exit Sub

    Else
        '■契約者情報設定==========================================================================================K
        '■エラーフラグリセット　2025/10/30追加
        db.Execute "UPDATE 契約者情報 SET エラーフラグ = '1';"
        log_write "set_account:契約者情報エラーフラグリセット1"
        
        '■出力レコード選定 2025/10/30追加
        SQL = ""
        SQL = SQL & "UPDATE 契約者情報 "
        SQL = SQL & "INNER JOIN 新規契約更新一覧 ON 契約者情報.[契約者No] = 新規契約更新一覧.[契約者2No] "
        SQL = SQL & "SET 契約者情報.エラーフラグ = '0' "
        SQL = SQL & "WHERE 契約者情報.エラーフラグ = '1' "
        SQL = SQL & "AND 新規契約更新一覧.[契約者2入居有無] = '入居有り';"
        db.Execute SQL
        
        SQL = ""
        SQL = SQL & "UPDATE 契約者情報 "
        SQL = SQL & "INNER JOIN 新規契約更新一覧 ON 契約者情報.[契約者No] = 新規契約更新一覧.[契約者1No] "
        SQL = SQL & "SET 契約者情報.エラーフラグ = '0' "
        SQL = SQL & "WHERE 契約者情報.エラーフラグ = '1' "
        SQL = SQL & "AND (新規契約更新一覧.[契約者2入居有無] IS NULL OR 新規契約更新一覧.[契約者2入居有無] = '');"
        db.Execute SQL
        
        SQL = ""
        SQL = SQL & "UPDATE 契約者情報 "
        SQL = SQL & "INNER JOIN 入居状況一覧 ON 契約者情報.[契約者No] = 入居状況一覧.[契約者2No] "
        SQL = SQL & "SET 契約者情報.エラーフラグ = '0' "
        SQL = SQL & "WHERE (入居状況一覧.[契約状況] = '契約中' OR 入居状況一覧.[契約状況] = '解約予定') "
        SQL = SQL & "AND Nz(入居状況一覧.[契約者2No],"""") <> """" "
        SQL = SQL & "AND ("
        SQL = SQL & "    EXISTS (SELECT 1 FROM 契約者情報 AS Tmp WHERE Tmp.[契約者No] = 入居状況一覧.[契約者No] AND Tmp.エラーフラグ = '1') "
        SQL = SQL & "    OR "
        SQL = SQL & "    契約者情報.エラーフラグ = '1' "
        SQL = SQL & ");"
        db.Execute SQL
        
        SQL = ""
        SQL = SQL & "UPDATE 契約者情報 "
        SQL = SQL & "INNER JOIN 入居状況一覧 ON 契約者情報.[契約者No] = 入居状況一覧.[契約者No] "
        SQL = SQL & "SET 契約者情報.エラーフラグ = '0' "
        SQL = SQL & "WHERE 契約者情報.エラーフラグ = '1' "
        SQL = SQL & "AND (入居状況一覧.[契約状況] = '契約中' OR 入居状況一覧.[契約状況] = '解約予定') "
        SQL = SQL & "AND Nz(入居状況一覧.[契約者2No],"""") = """";"
        db.Execute SQL
            
        log_write "set_account:契約者情報出力レコード抽出"
        
    
        '■契約者情報INSERT 2025/10/30改修
        SQL = ""
        SQL = SQL & "INSERT INTO Account SELECT "
        SQL = SQL & "契約者情報.ID , "
        SQL = SQL & """K"" & 契約者情報.ID as 識別ID, "
        SQL = SQL & """Resident"" as klass, "
        SQL = SQL & "iif(契約者情報.契約者区分 = ""個人"",契約者情報.契約者名,"""") as name_family, "
        SQL = SQL & "iif(契約者情報.契約者区分 = ""個人"",契約者情報.契約者名カナ,"""") as name_family_kana, "
        SQL = SQL & "契約者情報.優先メール as email, "
        SQL = SQL & "Replace(Nz(契約者情報.TEL1,""""),""-"","""",1,-1,1) as tel_fixed, "
        SQL = SQL & "Replace(Nz(契約者情報.携帯1,""""),""-"","""",1,-1,1) as tel_mobile, "
        SQL = SQL & "iif(契約者情報.契約者区分 = ""法人"",契約者情報.契約者名,"""") as company_name, "
        SQL = SQL & "iif(契約者情報.契約者区分 = ""法人"",契約者情報.契約者名カナ,"""") as company_name_kana, "
        SQL = SQL & "契約者情報.契約者区分 as kind_id, "
        SQL = SQL & "契約者情報.性別 as gender_id, "
        SQL = SQL & "format(契約者情報.生年月日,""yyyy-mm-dd"") as birthday, "
        SQL = SQL & """gmo002"" as legacy_charge_user_id, "
        SQL = SQL & "契約者情報.[契約者No] as legacy_id "
        SQL = SQL & "FROM 契約者情報 "
        SQL = SQL & "WHERE 契約者情報.エラーフラグ = ""0"";"
        db.Execute SQL
        log_write "set_account:契約者情報設定"
         
        
'        '■legacy_idの更新
'        SQL = ""
'        SQL = SQL & "UPDATE (Account "
'        SQL = SQL & "INNER JOIN 契約者情報 "
'        SQL = SQL & "ON Account.ID = 契約者情報.ID) "
'        SQL = SQL & "INNER JOIN 入居状況一覧 "
'        SQL = SQL & "ON (契約者情報.契約者No = 入居状況一覧.契約者No) "
'        SQL = SQL & "SET Account.legacy_id = " & _
'                    "Iif(Nz(入居状況一覧.契約者2No,"""") = """"," & _
'                    "入居状況一覧.契約者No,入居状況一覧.契約者2No)"
'        SQL = SQL & "Where Account.識別ID like ""K%"";"
'        db.Execute SQL
'        log_write "set_account:legacy_id_Rの設定"

        '■kind_idの設定(項目マッピング)
        SQL = ""
        SQL = SQL & "UPDATE Account "
        SQL = SQL & "LEFT JOIN T_MAP_ACC_KIND_ID "
        SQL = SQL & "ON Account.kind_id = "
        SQL = SQL & "T_MAP_ACC_KIND_ID.ORG "
        SQL = SQL & "SET Account.kind_id = "
        SQL = SQL & "[T_MAP_ACC_KIND_ID]![NEW] "
        SQL = SQL & "WHERE 識別ID like ""K%"";"
        db.Execute SQL
        log_write "set_account:kind_id_Rの設定"
        
        
        '■gender_idの設定(項目マッピング)
        SQL = ""
        SQL = SQL & "UPDATE Account "
        SQL = SQL & "LEFT JOIN T_MAP_ACC_GENDER_ID "
        SQL = SQL & "ON Account.gender_id = "
        SQL = SQL & "T_MAP_ACC_GENDER_ID.ORG "
        SQL = SQL & "SET Account.gender_id = "
        SQL = SQL & "[T_MAP_ACC_GENDER_ID]![NEW] "
        SQL = SQL & "WHERE 識別ID like ""K%"";"
        db.Execute SQL
        log_write "set_account:gender_id_Rの設定"
        
        
        
        '■エラーフラグリセット　2025/10/30追加
        db.Execute "UPDATE 契約者情報 SET エラーフラグ = '0';"
        log_write "set_account:契約者情報エラーフラグリセット2"
    
    End If
    
    
    '■エラーチェック
    chk_required
    
    
    '進捗状況の更新~~~~~~~~~~~~~~~~~~
    DoEvents
    Call UpdateProgressBar(frmName)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Set db = Nothing
            
    Debug.Print "account:" & Timer - t
    
    log_write "set_account:out"

End Sub

Private Sub chk_required()
    log_write "chk_required:in"
    
    Dim t As Single
    t = Timer
    
    Dim db As ADODB.Connection
    Dim SQL As String
    
    Set db = CurrentProject.Connection
    
    
    '■Kind_idが「10」かつname_familyがブランク
    SQL = ""
    SQL = SQL & "UPDATE Account "
    SQL = SQL & "SET Account.[エラー内容] = "
    SQL = SQL & "Account.[エラー内容] & ""/家主名または契約者名がブランク"", "
    SQL = SQL & "Account.[エラー項目] = "
    SQL = SQL & "Account.[エラー項目] & ""/name_family"" "
    SQL = SQL & "WHERE (((Account.kind_id)=""10"") "
    SQL = SQL & "AND ((Account.name_family) Is Null "
    SQL = SQL & "Or (Account.name_family)=""""));"
    db.Execute SQL
    log_write "chk_required:name_family→10"
    
    
    '■Kind_idが「20」かつcompany_nameがブランク
    SQL = ""
    SQL = SQL & "UPDATE Account "
    SQL = SQL & "SET Account.[エラー内容] = "
    SQL = SQL & "Account.[エラー内容] & ""/家主名または契約者名がブランク"", "
    SQL = SQL & "Account.[エラー項目] = "
    SQL = SQL & "Account.[エラー項目] & ""/company_name"" "
    SQL = SQL & "WHERE (((Account.kind_id)=""20"") "
    SQL = SQL & "AND ((Account.company_name) Is Null "
    SQL = SQL & "Or (Account.company_name)=""""));"
    db.Execute SQL
    log_write "chk_required:company_name→20"
        
    
'    '■kind_idの値が10 or 20以外
'    SQL = ""
'    SQL = SQL & "UPDATE Account "
'    SQL = SQL & "SET Account.[エラー内容] = "
'    SQL = SQL & "Account.[エラー内容] & ""/関係者種別の値が不正"", "
'    SQL = SQL & "Account.[エラー項目] = "
'    SQL = SQL & "Account.[エラー項目] & ""/kind_id"" "
'    SQL = SQL & "WHERE (((Account.kind_id)<>""10"") "
'    SQL = SQL & "AND ((Account.kind_id)<>""20"")) "
'    SQL = SQL & "OR (Nz(Account.kind_id,"""") = """");"
'    db.Execute SQL
'    log_write "chk_required:kind_id"
    
    
    
''    ■legacy_charge_user_idがブランク
'    SQL = ""
'    SQL = SQL & "UPDATE Account "
'    SQL = SQL & "SET Account.[エラー内容] = "
'    SQL = SQL & "Account.[エラー内容] & ""/担当者IDがブランク"", "
'    SQL = SQL & "Account.[エラー項目] = "
'    SQL = SQL & "Account.[エラー項目] & ""/legacy_charge_user_id"""
'    SQL = SQL & "WHERE (((Account.legacy_charge_user_id) Is Null "
'    SQL = SQL & "Or (Account.legacy_charge_user_id)=""""));"
'    db.Execute SQL
'    log_write "chk_required:legacy_charge_user_id"
    
    
    
    '■legacy_idがブランク
    SQL = ""
    SQL = SQL & "UPDATE Account "
    SQL = SQL & "SET Account.[エラー内容] = "
    SQL = SQL & "Account.[エラー内容] & ""/家主Noまたは契約者Noがブランク"", "
    SQL = SQL & "Account.[エラー項目] = "
    SQL = SQL & "Account.[エラー項目] & ""/legacy_id"""
    SQL = SQL & "WHERE (((Account.legacy_id) Is Null "
    SQL = SQL & "Or (Account.legacy_id)=""""));"
    db.Execute SQL
    log_write "chk_required:legacy_id"
    
    Set db = Nothing
    
    Debug.Print Timer - t
    log_write "chk_required:out"
    
End Sub



Sub out_account_csv(strDir As String)
    log_write "out_account_csv:in"
    
    Dim t As Single
    t = Timer
    
    Dim db As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim SQL As String
    
    Dim ST As ADODB.Stream
    Dim WS As ADODB.Stream
    Dim LINE As String
    Dim i As Long
    
    Set ST = New ADODB.Stream
    ST.Charset = "UTF-8"
    ST.LineSeparator = adCRLF
    ST.Open
    
    Set db = CurrentProject.Connection
    
    SQL = ""
    SQL = SQL & "SELECT * FROM Account WHERE (((Account.[エラー内容]) Is Null or (Account.[エラー内容])=''));"
    
    Set RS = db.Execute(SQL)
    log_write "out_account_csv:Account取得"
    
    LINE = ""
    For i = 5 To RS.Fields.Count - 1
        LINE = LINE & """" & RS.Fields(i).Name & ""","
    Next
    LINE = Left(LINE, Len(LINE) - 1)
    ST.WriteText LINE, stWriteLine
    log_write "out_account_csv:Accountヘッダ出力"
   
    Do Until RS.EOF
    
        LINE = ""
        
        For i = 5 To RS.Fields.Count - 1
            If Nz(RS.Fields(i).Value, "") = "" Then
                LINE = LINE & ","
            Else
                LINE = LINE & """" & Replace(Nz(RS.Fields(i).Value, ""), """", """""", , , vbTextCompare) & ""","
            End If

        Next
        LINE = Left(LINE, Len(LINE) - 1)
        ST.WriteText LINE, stWriteLine
        
        
        RS.MoveNext
        DoEvents
    Loop
    log_write "out_account_csv:Accountレコード出力"
    
    ST.SaveToFile strDir & "\account.csv", adSaveCreateOverWrite
    log_write "out_account_csv:Account.csv保存"
    
    Debug.Print Timer - t
    log_write "out_account_csv:out"
    
End Sub

Sub out_account_xls(strDir As String)
    log_write "out_account_xls:in"
    Dim t As Single
    t = Timer
    
    Dim FS As Object
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    If FS.FileExists(strDir & "\account.xlsx") = True Then
        FS.DeleteFile strDir & "\account.xlsx"
    End If
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "Q_Account", strDir & "\account.xlsx", True
    log_write "out_account_xls:Account.xlsx保存"
        
    Debug.Print Timer - t
    log_write "out_account_xls:out"
    
End Sub
