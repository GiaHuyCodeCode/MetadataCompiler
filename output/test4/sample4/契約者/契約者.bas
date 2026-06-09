Attribute VB_Name = "契約者"
Option Compare Database
Option Explicit

Sub set_contract()
    log_write "set_contract:in"

    Dim t As Single
    t = Timer

    Dim db As DAO.Database
    Dim SQL As String

    Set db = CurrentDb

    db.Execute "DELETE FROM Contract WHERE ID LIKE 'X_%';"

    If DCount("ID", "Contract", "ID LIKE 'X_*'") = 0 Then

    '================================================================================
    '契約者
    '================================================================================

    '出力条件
    AddNewFieldToTable "入居状況一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 入居状況一覧 SET 入居状況一覧.[FLG] = '0';"

    'JOIN filter
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T1 "
    SQL = SQL & "INNER JOIN 契約情報一覧 AS T2 "
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) AND (T1.[部屋No] = T2.[部屋No]) AND (T1.[初回契約始期] = T2.[初回契約始期]) "
    SQL = SQL & "SET T1.FLG = '1'; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "契約者情報.csv", "該当ファイル名": "契約情報一覧.csv", "該当項目名_1": "契約者No", "該当項目名_2": "契約者名(SJIS)", "処理": "そのまま出力"}]

    'JOIN filter
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T1 "
    SQL = SQL & "INNER JOIN 契約者情報 AS T2 "
    SQL = SQL & "ON (T1.[上記で契約情報一覧.csvからマージした契約者No] = T2.[上記で契約情報一覧.csvからマージした契約者No]) "
    SQL = SQL & "SET T1.FLG = '1'; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"該当ファイル名": "契約者情報.csv", "該当項目名_1": "契約者区分", "該当項目名_2": "携帯1", "該当項目名_3": "優先メール", "処理": "そのまま出力"}]

    log_write "set_contract:契約者 → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Contract SELECT "
    SQL = SQL & "    'X_' & T.ID as ID,"
    SQL = SQL & "    T.[契約者名(SJIS)] as name,"
    SQL = SQL & "    T.[優先メール] as email,"
    SQL = SQL & "    T.[携帯1] as tel_mobile,"
    SQL = SQL & "    T.[契約者No] as legacy_id,"
    SQL = SQL & "    '契約者' as sheet"
    SQL = SQL & "FROM 入居状況一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_contract:契約者 → 通常処理"


    '■項目削除
    DeleteFieldInTable "入居状況一覧", "FLG"

    End If

    chk_required

    Debug.Print "Contract:" & Timer - t
    log_write "set_contract:out"

End Sub
