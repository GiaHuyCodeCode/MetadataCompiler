Attribute VB_Name = "建物（新規）"
Option Compare Database
Option Explicit

Sub set_building()
    log_write "set_building:in"

    Dim t As Single
    t = Timer

    Dim db As DAO.Database
    Dim SQL As String

    Set db = CurrentDb

    db.Execute "DELETE FROM Building WHERE ID LIKE 'B_%';"

    If DCount("ID", "Building", "ID LIKE 'B_*'") = 0 Then

    '================================================================================
    '建物（新規）
    '================================================================================

    '出力条件
    AddNewFieldToTable "新規契約更新一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 新規契約更新一覧 SET 新規契約更新一覧.[FLG] = '1';"

    'Blank check — 物件No
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T "
    SQL = SQL & "SET T.FLG = '1' "
    SQL = SQL & "WHERE Nz(T.[物件No], '') = ''; "
    db.Execute SQL

    'JOIN filter → 紐づかないデータを抽出
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T1 "
    SQL = SQL & "LEFT JOIN 入居状況一覧 AS T2 "
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[物件No] IS NULL; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"該当ファイル名": "新規契約更新一覧.csv", "該当項目名_1": "物件No", "処理": "そのまま出力"}]

    log_write "set_building:建物（新規） → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Building SELECT "
    SQL = SQL & "    'B_' & T.ID as ID,"
    SQL = SQL & "    T.[物件名] as name,"
    SQL = SQL & "    T.[物件No] as legacy_id,"
    SQL = SQL & "    '建物（新規）' as sheet"
    SQL = SQL & "FROM 新規契約更新一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_building:建物（新規） → 通常処理"


    '■項目削除
    DeleteFieldInTable "新規契約更新一覧", "FLG"

    End If

    chk_required

    Debug.Print "Building:" & Timer - t
    log_write "set_building:out"

End Sub
