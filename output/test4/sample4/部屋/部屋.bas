Attribute VB_Name = "部屋"
Option Compare Database
Option Explicit

Sub set_property()
    log_write "set_property:in"

    Dim t As Single
    t = Timer

    Dim db As DAO.Database
    Dim SQL As String

    Set db = CurrentDb

    db.Execute "DELETE FROM Property WHERE ID LIKE 'BN_%';"

    If DCount("ID", "Property", "ID LIKE 'BN_*'") = 0 Then

    '================================================================================
    '部屋
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

    log_write "set_property:部屋 → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Property SELECT "
    SQL = SQL & "    'BN_' & T.ID as ID,"
    SQL = SQL & "    T.[家主No] as legacy_owner_id,"
    SQL = SQL & "    T.[物件No] as legacy_building_id,"
    SQL = SQL & "    T.[部屋No] as name,"
    SQL = SQL & "    T.[郵便番号] as zip_code,"
    SQL = SQL & "    T.[都道府県名] as prefecture_code,"
    SQL = SQL & "    T.[市区町村名] as city_code,"
    SQL = SQL & "    T.[町地域名] & T.[丁番地名] & T.[住所その他] as address_1,"
    SQL = SQL & "    T.[物件No] & ""-"" & T.[部屋No] as legacy_id,"
    SQL = SQL & "    '部屋' as sheet"
    SQL = SQL & "FROM 入居状況一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_property:部屋 → 通常処理"


    '■項目削除
    DeleteFieldInTable "入居状況一覧", "FLG"

    End If

    chk_required

    Debug.Print "Property:" & Timer - t
    log_write "set_property:out"

End Sub
