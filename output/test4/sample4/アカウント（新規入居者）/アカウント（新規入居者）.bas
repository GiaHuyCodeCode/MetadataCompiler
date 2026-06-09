Attribute VB_Name = "アカウント（新規入居者）"
Option Compare Database
Option Explicit

Sub set_account()
    log_write "set_account:in"

    Dim t As Single
    t = Timer

    Dim db As DAO.Database
    Dim SQL As String

    Set db = CurrentDb

    db.Execute "DELETE FROM Account WHERE ID LIKE 'X_%';"

    If DCount("ID", "Account", "ID LIKE 'X_*'") = 0 Then

    '================================================================================
    'アカウント（新規入居者）
    '================================================================================

    '出力条件
    AddNewFieldToTable "新規契約更新一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 新規契約更新一覧 SET 新規契約更新一覧.[FLG] = '1';"

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "ファイル名", "目的": "入居状況にない未来の契約の入居者を抽出", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約者3入居有無", "処理": "条件分岐", "条件": "\"入居有り\"の場合"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "新規契約更新一覧.csv", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約者3No", "処理": "そのまま出力"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "GMO 入居状況一覧.csv"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約者2入居有無", "処理": "条件分岐", "条件": "上記以外で\"入居有り\"の場合"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約者2No", "処理": "そのまま出力"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約者1入居有無", "処理": "条件分岐", "条件": "上記以外で\"入居有り\"の場合"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約者1No", "処理": "そのまま出力"}]

    'JOIN filter → 紐づかないデータを抽出
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T1 "
    SQL = SQL & "LEFT JOIN 新規契約更新一覧 AS T2 "
    SQL = SQL & "ON (T1.[上記の条件分岐で出力した契約者No] = T2.[上記の条件分岐で出力した契約者No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[上記の条件分岐で出力した契約者No] IS NULL; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"目的": "重複削除", "該当ファイル名": "新規契約更新一覧.csv", "該当項目名_1": "契約者No", "処理": "そのまま出力", "条件": "データ順で１番最初のデータを出力対象とする"}]

    log_write "set_account:アカウント（新規入居者） → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Account SELECT "
    SQL = SQL & "    'X_' & T.ID as ID,"
    SQL = SQL & "    'Resident' as klass,"
    SQL = SQL & "    'ohta-ff' as legacy_charge_user_id,"
    SQL = SQL & "    T.[契約者No] as legacy_id,"
    SQL = SQL & "    'アカウント（新規入居者）' as sheet"
    SQL = SQL & "FROM 新規契約更新一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_account:アカウント（新規入居者） → 通常処理"

    '特殊処理

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[name_family] = SWITCH( "
    SQL = SQL & "    (T1.[契約者名] NOT LIKE '%会社%'), T1.[契約者名], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（新規入居者） → 特殊処理: name_family"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[company_name] = SWITCH( "
    SQL = SQL & "    (T1.[契約者名] LIKE '%会社%'), T1.[契約者名], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（新規入居者） → 特殊処理: company_name"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    (T1.[契約者名] NOT LIKE '%会社%'), '10', "
    SQL = SQL & "    (T1.[契約者名] LIKE '%会社%'), '20', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（新規入居者） → 特殊処理: kind_id"


    '■項目削除
    DeleteFieldInTable "新規契約更新一覧", "FLG"

    End If

    chk_required

    Debug.Print "Account:" & Timer - t
    log_write "set_account:out"

End Sub
