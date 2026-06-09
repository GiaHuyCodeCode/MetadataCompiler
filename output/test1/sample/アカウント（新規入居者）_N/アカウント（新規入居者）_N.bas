Attribute VB_Name = "アカウント（新規入居者）_N"
Option Compare Database
Option Explicit

Sub set_account()
    log_write "set_account:in"

    Dim t As Single
    t = Timer

    Dim db As ADODB.Connection
    Dim SQL As String

    Set db = CurrentProject.Connection

    db.Execute "DELETE FROM Account WHERE ID LIKE 'N_%';"

    If DCount("ID", "Account", "ID LIKE 'N_*'") = 0 Then

    '================================================================================
    'アカウント（新規入居者）_N
    '================================================================================

    '出力条件
    AddNewFieldToTable "新規契約更新一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 新規契約更新一覧 SET 新規契約更新一覧.[FLG] = '1';"

    'JOIN filter → 紐づかないデータを抽出
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T1 "
    SQL = SQL & "LEFT JOIN 入居状況一覧 AS T2 "
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) AND (T1.[部屋No] = T2.[部屋No]) AND (T1.[契約者No] = T2.[契約者No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[物件No] IS NULL; "
    db.Execute SQL

    '重複チェック 契約者Noをキーに重複を削除
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 "
    SQL = SQL & "SET 新規契約更新一覧.FLG = '1' "
    SQL = SQL & "WHERE 新規契約更新一覧.ID NOT IN ( "
    SQL = SQL & "    SELECT MIN(ID) "
    SQL = SQL & "    FROM 新規契約更新一覧 "
    SQL = SQL & "    WHERE FLG = '0' "
    SQL = SQL & "    GROUP BY [契約者No] "
    SQL = SQL & "); "
    db.Execute SQL

    log_write "set_account:アカウント（新規入居者）_N → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Account SELECT "
    SQL = SQL & "    'N_' & T.ID as ID,"
    SQL = SQL & "    'Resident' as klass,"
    SQL = SQL & "    'gmo002' as legacy_charge_user_id,"
    SQL = SQL & "    T.[契約者No] as legacy_id,"
    SQL = SQL & "    'アカウント（新規入居者）_N' as sheet"
    SQL = SQL & "FROM 新規契約更新一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_account:アカウント（新規入居者）_N → 通常処理"

    '特殊処理

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID) "
    SQL = SQL & "SET T.[name_family] = SWITCH( "
    SQL = SQL & "    (T1.[契約者名] NOT LIKE '%会社%'), T1.[契約者名], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（新規入居者）_N → 特殊処理: name_family"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID) "
    SQL = SQL & "SET T.[company_name] = SWITCH( "
    SQL = SQL & "    (T1.[契約者名] LIKE '%会社%'), T1.[契約者名], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（新規入居者）_N → 特殊処理: company_name"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    T1.[契約者名] NOT LIKE '%会社%', '10', "
    SQL = SQL & "    T1.[契約者名] LIKE '%会社%', '20', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（新規入居者）_N → 特殊処理: kind_id"


    '■項目削除
    DeleteFieldInTable "新規契約更新一覧", "FLG"

    End If

    chk_required

    Debug.Print "Account:" & Timer - t
    log_write "set_account:out"

End Sub
