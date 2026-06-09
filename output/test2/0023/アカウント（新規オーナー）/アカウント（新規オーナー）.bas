Attribute VB_Name = "アカウント（新規オーナー）"
Option Compare Database
Option Explicit

Sub set_account()
    log_write "set_account:in"

    Dim t As Single
    t = Timer

    Dim db As DAO.Database
    Dim SQL As String

    Set db = CurrentDb

    db.Execute "DELETE FROM Account WHERE ID LIKE 'ON_%';"

    If DCount("ID", "Account", "ID LIKE 'ON_*'") = 0 Then

    '================================================================================
    'アカウント（新規オーナー）
    '================================================================================

    '出力条件
    AddNewFieldToTable "新規契約更新一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 新規契約更新一覧 SET 新規契約更新一覧.[FLG] = '1';"

    ' --- AI GENERATED (sk-architect) ---
    'Blank check — 部屋No
        SQL = ""
        SQL = SQL & "UPDATE 新規契約更新一覧 AS T "
        SQL = SQL & "SET T.FLG = '1' "
        SQL = SQL & "WHERE Nz(T.[部屋No], '') = ''; "
        db.Execute SQL
    ' --- END AI GENERATED ---

    'JOIN filter → 紐づかないデータを抽出
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T1 "
    SQL = SQL & "LEFT JOIN 入居状況一覧 AS T2 "
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) AND (T1.[部屋No] = T2.[部屋No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[物件No] IS NULL; "
    db.Execute SQL

    'JOIN filter → 紐づかないデータを抽出
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T1 "
    SQL = SQL & "LEFT JOIN 新規契約更新一覧 AS T2 "
    SQL = SQL & "ON (T1.[家主No] = T2.[家主No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[家主No] IS NULL; "
    db.Execute SQL

    ' --- AI GENERATED (sk-architect) ---
    '重複チェック 家主Noをキーに重複を削除
        SQL = ""
        SQL = SQL & "UPDATE 新規契約更新一覧 "
        SQL = SQL & "SET 新規契約更新一覧.FLG = '1' "
        SQL = SQL & "WHERE 新規契約更新一覧.ID NOT IN ( "
        SQL = SQL & "    SELECT MIN(ID) "
        SQL = SQL & "    FROM 新規契約更新一覧 "
        SQL = SQL & "    WHERE FLG = '0' "
        SQL = SQL & "    GROUP BY [家主No] "
        SQL = SQL & "); "
        db.Execute SQL
    ' --- END AI GENERATED ---

    log_write "set_account:アカウント（新規オーナー） → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Account SELECT "
    SQL = SQL & "    'ON_' & T.ID as ID,"
    SQL = SQL & "    'Owner' as klass,"
    SQL = SQL & "    '12' as closing_month,"
    SQL = SQL & "    '31' as closing_day,"
    SQL = SQL & "    T.[家主No] as legacy_id,"
    SQL = SQL & "    'アカウント（新規オーナー）' as sheet"
    SQL = SQL & "FROM 新規契約更新一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_account:アカウント（新規オーナー） → 通常処理"

    '特殊処理

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('ON_' & T1.ID) "
    SQL = SQL & "SET T.[name_family] = SWITCH( "
    SQL = SQL & "    (T1.[貸主名] NOT LIKE '%会社%'), T1.[貸主名], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（新規オーナー） → 特殊処理: name_family"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('ON_' & T1.ID) "
    SQL = SQL & "SET T.[company_name] = SWITCH( "
    SQL = SQL & "    (T1.[貸主名] LIKE '%会社%'), T1.[貸主名], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（新規オーナー） → 特殊処理: company_name"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('ON_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    (T1.[貸主名] LIKE '%会社%'), '20', "
    SQL = SQL & "    (True), '10', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（新規オーナー） → 特殊処理: kind_id"


    '■項目削除
    DeleteFieldInTable "新規契約更新一覧", "FLG"

    End If

    chk_required

    Debug.Print "Account:" & Timer - t
    log_write "set_account:out"

End Sub
