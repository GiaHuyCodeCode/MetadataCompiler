Attribute VB_Name = "アカウント（オーナー）_O"
Option Compare Database
Option Explicit

Sub set_account()
    log_write "set_account:in"

    Dim t As Single
    t = Timer

    Dim db As ADODB.Connection
    Dim SQL As String

    Set db = CurrentProject.Connection

    db.Execute "DELETE FROM Account WHERE ID LIKE 'ON_%';"

    If DCount("ID", "Account", "ID LIKE 'ON_*'") = 0 Then

    '================================================================================
    'アカウント（オーナー）_O
    '================================================================================

    '出力条件
    AddNewFieldToTable "入居状況一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 入居状況一覧 SET 入居状況一覧.[FLG] = '0';"

    'Blank check — 家主No
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T "
    SQL = SQL & "SET T.FLG = '1' "
    SQL = SQL & "WHERE Nz(T.[家主No], '') = ''; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"使用ファイル": "GMO 入居状況一覧.csv"}]

    log_write "set_account:アカウント（オーナー）_O → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Account SELECT "
    SQL = SQL & "    'ON_' & T.ID as ID,"
    SQL = SQL & "    'Owner' as klass,"
    SQL = SQL & "    T.[メールアドレス] as email,"
    SQL = SQL & "    T.[TEL1] as tel_fixed,"
    SQL = SQL & "    '12' as closing_month,"
    SQL = SQL & "    '31' as closing_day,"
    SQL = SQL & "    T.[家主No] as legacy_id,"
    SQL = SQL & "    'アカウント（オーナー）_O' as sheet"
    SQL = SQL & "FROM 入居状況一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_account:アカウント（オーナー）_O → 通常処理"

    '特殊処理

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('ON_' & T1.ID) "
    SQL = SQL & "SET T.[name_family] = SWITCH( "
    SQL = SQL & "    (T1.[個人・法人区分_0] = '1'), T1.[家主名], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（オーナー）_O → 特殊処理: name_family"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('ON_' & T1.ID) "
    SQL = SQL & "SET T.[name_family_kana] = SWITCH( "
    SQL = SQL & "    (T1.[個人・法人区分_0] = '1'), T1.[家主名カナ], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（オーナー）_O → 特殊処理: name_family_kana"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('ON_' & T1.ID) "
    SQL = SQL & "SET T.[company_name] = SWITCH( "
    SQL = SQL & "    (T1.[個人・法人区分_0] = '2'), T1.[家主名], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（オーナー）_O → 特殊処理: company_name"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('ON_' & T1.ID) "
    SQL = SQL & "SET T.[company_name_kana] = SWITCH( "
    SQL = SQL & "    (T1.[個人・法人区分_0] = '2'), T1.[家主名カナ], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（オーナー）_O → 特殊処理: company_name_kana"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('ON_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    T1.[個人・法人区分_0] = '1', '10', "
    SQL = SQL & "    T1.[個人・法人区分_0] = '2', '20', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（オーナー）_O → 特殊処理: kind_id"


    '■項目削除
    DeleteFieldInTable "入居状況一覧", "FLG"

    End If

    chk_required

    Debug.Print "Account:" & Timer - t
    log_write "set_account:out"

End Sub
