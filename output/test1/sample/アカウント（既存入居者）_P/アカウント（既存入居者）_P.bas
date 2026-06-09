Attribute VB_Name = "アカウント（既存入居者）_P"
Option Compare Database
Option Explicit

Sub set_account()
    log_write "set_account:in"

    Dim t As Single
    t = Timer

    Dim db As ADODB.Connection
    Dim SQL As String

    Set db = CurrentProject.Connection

    db.Execute "DELETE FROM Account WHERE ID LIKE 'P_%';"

    If DCount("ID", "Account", "ID LIKE 'P_*'") = 0 Then

    '================================================================================
    'アカウント（既存入居者）_P
    '================================================================================

    '出力条件
    AddNewFieldToTable "入居者管理", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 入居者管理 SET 入居者管理.[FLG] = '1';"

    'JOIN filter → 紐づかないデータを抽出
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T1 "
    SQL = SQL & "LEFT JOIN 入居状況一覧 AS T2 "
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) AND (T1.[部屋No] = T2.[部屋No]) AND (T1.[契約者No] = T2.[契約者No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[物件No] IS NULL; "
    db.Execute SQL

    'JOIN filter → 紐づかないデータを抽出
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T1 "
    SQL = SQL & "LEFT JOIN 入居者管理 AS T2 "
    SQL = SQL & "ON (T1.[契約者No] = T2.[契約者No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[契約者No] IS NULL; "
    db.Execute SQL

    log_write "set_account:アカウント（既存入居者）_P → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Account SELECT "
    SQL = SQL & "    'P_' & T.ID as ID,"
    SQL = SQL & "    'Resident' as klass,"
    SQL = SQL & "    IIF(T.[メールアドレス] = '-', NULL, T.[メールアドレス]) as email,"
    SQL = SQL & "    IIF(T.[電話番号] = '-', NULL, T.[電話番号]) as tel_fixed,"
    SQL = SQL & "    IIF(T.[携帯番号/SMS] = '-', NULL, T.[携帯番号/SMS]) as tel_mobile,"
    SQL = SQL & "    IIF(T.[郵便番号] = '-', NULL, T.[郵便番号]) as zip_code,"
    SQL = SQL & "    IIF(T.[都道府県] = '-', NULL, T.[都道府県]) as prefecture_code,"
    SQL = SQL & "    IIF(T.[市区町村] = '-', NULL, T.[市区町村]) as city_code,"
    SQL = SQL & "    IIF(T.[住所１] = '-', NULL, T.[住所１]) as address_1,"
    SQL = SQL & "    IIF(T.[住所2] = '-', NULL, T.[住所2]) as address_2,"
    SQL = SQL & "    IIF(T.[性別] = '-', NULL, T.[性別]) as gender_id,"
    SQL = SQL & "    IIF(T.[生年月日] = '-', NULL, T.[生年月日]) as birthday,"
    SQL = SQL & "    'gmo002' as legacy_charge_user_id,"
    SQL = SQL & "    T.[基幹入居者ID] as legacy_id,"
    SQL = SQL & "    'アカウント（既存入居者）_P' as sheet"
    SQL = SQL & "FROM 入居者管理 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_account:アカウント（既存入居者）_P → 通常処理"

    '特殊処理

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居者管理 AS T1 ON T.ID = ('P_' & T1.ID) "
    SQL = SQL & "SET T.[name_family] = SWITCH( "
    SQL = SQL & "    (T1.[個人／法人] = '個人'), T1.[姓], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（既存入居者）_P → 特殊処理: name_family"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居者管理 AS T1 ON T.ID = ('P_' & T1.ID) "
    SQL = SQL & "SET T.[company_name] = SWITCH( "
    SQL = SQL & "    (T1.[個人／法人] = '法人'), T1.[姓], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（既存入居者）_P → 特殊処理: company_name"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居者管理 AS T1 ON T.ID = ('P_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    T1.[個人／法人] = '個人', '10', "
    SQL = SQL & "    T1.[個人／法人] = '法人', '20', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（既存入居者）_P → 特殊処理: kind_id"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居者管理 AS T1 ON T.ID = ('P_' & T1.ID) "
    SQL = SQL & "SET T.[tag] = SWITCH( "
    SQL = SQL & "    T1.[タグ] = '解約予定', '退去済', "
    SQL = SQL & "    T1.[タグ] = '退去済', '退去済', "
    SQL = SQL & "    Nz(Replace(T1.[タグ], '-', ''), '') = '', '管理移行', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（既存入居者）_P → 特殊処理: tag"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居者管理 AS T1 ON T.ID = ('P_' & T1.ID) "
    SQL = SQL & "SET T.[name_family_kana] = SWITCH( "
    SQL = SQL & "    (T1.[個人／法人] = '個人'), T1.[姓 (カナ)], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（既存入居者）_P → 特殊処理: name_family_kana"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居者管理 AS T1 ON T.ID = ('P_' & T1.ID) "
    SQL = SQL & "SET T.[company_name_kana] = SWITCH( "
    SQL = SQL & "    T1.[個人／法人] = '法人', T1.[会社名 (カナ)], "
    SQL = SQL & "    Nz(Replace(T1.[会社名 (カナ)], '-', ''), '') = '', NULL, "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（既存入居者）_P → 特殊処理: company_name_kana"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居者管理 AS T1 ON T.ID = ('P_' & T1.ID) "
    SQL = SQL & "SET T.[name_first] = SWITCH( "
    SQL = SQL & "    T1.[個人／法人] = '個人', T1.[名], "
    SQL = SQL & "    Nz(Replace(T1.[名], '-', ''), '') = '', NULL, "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（既存入居者）_P → 特殊処理: name_first"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居者管理 AS T1 ON T.ID = ('P_' & T1.ID) "
    SQL = SQL & "SET T.[name_first_kana] = SWITCH( "
    SQL = SQL & "    T1.[個人／法人] = '個人', T1.[名 (カナ)], "
    SQL = SQL & "    Nz(Replace(T1.[名 (カナ)], '-', ''), '') = '', NULL, "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（既存入居者）_P → 特殊処理: name_first_kana"


    '■項目削除
    DeleteFieldInTable "入居者管理", "FLG"

    End If

    chk_required

    Debug.Print "Account:" & Timer - t
    log_write "set_account:out"

End Sub
