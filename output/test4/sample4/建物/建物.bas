Attribute VB_Name = "建物"
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
    '建物
    '================================================================================

    '出力条件
    AddNewFieldToTable "入居状況一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 入居状況一覧 SET 入居状況一覧.[FLG] = '0';"

    'Blank check — 物件No
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T "
    SQL = SQL & "SET T.FLG = '1' "
    SQL = SQL & "WHERE Nz(T.[物件No], '') = ''; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "GMO 入居状況一覧.csv"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "物件基本情報.csv"}]

    log_write "set_building:建物 → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Building SELECT "
    SQL = SQL & "    'B_' & T.ID as ID,"
    SQL = SQL & "    T.[物件名] as name,"
    SQL = SQL & "    T.[物件No] as legacy_id,"
    SQL = SQL & "    T.[都道府県名] as prefecture_code,"
    SQL = SQL & "    T.[市区町村名] as city_code,"
    SQL = SQL & "    T.[町地域名] & T.[丁番地名] & T.[住所その他] as address,"
    SQL = SQL & "    T.[郵便番号] as zip_code,"
    SQL = SQL & "    '建物' as sheet"
    SQL = SQL & "FROM 入居状況一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_building:建物 → 通常処理"

    '特殊処理

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"目的": "物件基本情報.csvをマージし各項目を出力", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "物件No", "処理": "紐づけ", "_target_field": ""}, {"該当ファイル名": "家主基本情報.csv", "該当項目名_1": "物件No", "処理": "紐づけ", "条件 / 項目マッピング": "紐づく場合", "_target_field": ""}]


    '■項目削除
    DeleteFieldInTable "入居状況一覧", "FLG"

    End If

    chk_required

    Debug.Print "Building:" & Timer - t
    log_write "set_building:out"

End Sub
