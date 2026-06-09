Attribute VB_Name = "部屋（新規）"
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
    '部屋（新規）
    '================================================================================

    '出力条件
    AddNewFieldToTable "新規契約更新一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 新規契約更新一覧 SET 新規契約更新一覧.[FLG] = '1';"

    'JOIN filter → 紐づかないデータを抽出
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T1 "
    SQL = SQL & "LEFT JOIN 入居状況一覧 AS T2 "
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) AND (T1.[部屋No] = T2.[部屋No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[物件No] IS NULL; "
    db.Execute SQL

    ' --- AI GENERATED (sk-architect) ---
    '出力 (No filter required)
    ' --- END AI GENERATED ---

    '■有効レコードのみFLG=0に戻す
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T "
    SQL = SQL & "SET T.FLG = '0' "
    SQL = SQL & "WHERE T.[家主No] = 'ブランク' OR T.[家主No] = '0'; "
    db.Execute SQL

    ' --- AI GENERATED (sk-architect) ---
    'Exclude check — 部屋No contains 駐車場 or 看板
        SQL = ""
        SQL = SQL & "UPDATE 新規契約更新一覧 AS T "
        SQL = SQL & "SET T.FLG = '1' "
        SQL = SQL & "WHERE (T.[部屋No] LIKE '*駐車場*' OR T.[部屋No] LIKE '*看板*'); "
        db.Execute SQL
    ' --- END AI GENERATED ---

    ' --- AI GENERATED (sk-architect) ---
    '出力なし (No action required)
    ' --- END AI GENERATED ---

    log_write "set_property:部屋（新規） → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Property SELECT "
    SQL = SQL & "    'BN_' & T.ID as ID,"
    SQL = SQL & "    T.[家主No] as legacy_owner_id,"
    SQL = SQL & "    T.[物件No] as legacy_building_id,"
    SQL = SQL & "    T.[部屋No] as name,"
    SQL = SQL & "    '10' as kind_id,"
    SQL = SQL & "    T.[物件No] & ""-"" & T.[部屋No] as legacy_id,"
    SQL = SQL & "    '部屋（新規）' as sheet"
    SQL = SQL & "FROM 新規契約更新一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_property:部屋（新規） → 通常処理"


    '■項目削除
    DeleteFieldInTable "新規契約更新一覧", "FLG"

    End If

    chk_required

    Debug.Print "Property:" & Timer - t
    log_write "set_property:out"

End Sub
