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

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"該当ファイル名": "新規契約更新一覧.csv", "処理": "そのまま出力", "条件": "重複チェックし、データ順で1番最初の1件のみ出力"}]

    log_write "set_property:部屋（新規） → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Property SELECT "
    SQL = SQL & "    'BN_' & T.ID as ID,"
    SQL = SQL & "    T.[貸主No] as legacy_owner_id,"
    SQL = SQL & "    T.[物件No] as legacy_building_id,"
    SQL = SQL & "    T.[部屋No] as name,"
    SQL = SQL & "    T.[物件No] & ""-"" & T.[部屋No] as legacy_id,"
    SQL = SQL & "    '部屋（新規）' as sheet"
    SQL = SQL & "FROM 新規契約更新一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_property:部屋（新規） → 通常処理"

    '特殊処理

    SQL = ""
    SQL = SQL & "UPDATE Property AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('BN_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    (T1.[部屋No] LIKE '%P%'), '20.0', "
    SQL = SQL & "    (T1.[部屋No] = ''), '10.0', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_property:部屋（新規） → 特殊処理: kind_id"


    '■項目削除
    DeleteFieldInTable "新規契約更新一覧", "FLG"

    End If

    chk_required

    Debug.Print "Property:" & Timer - t
    log_write "set_property:out"

End Sub
