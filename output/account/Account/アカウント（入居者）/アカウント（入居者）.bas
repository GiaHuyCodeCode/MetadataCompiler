Attribute VB_Name = "アカウント（入居者）"
Option Compare Database
Option Explicit

Sub set_account()
    log_write "set_account:in"

    Dim t As Single
    t = Timer

    Dim db As ADODB.Connection
    Dim SQL As String

    Set db = CurrentProject.Connection

    db.Execute "DELETE FROM [Account] WHERE ID LIKE 'X_%';"

    If DCount("ID", "Account", "ID LIKE 'X_*'") = 0 Then

    '================================================================================
    'アカウント（入居者）
    '================================================================================

    '出力条件
    AddNewFieldToTable "契約者情報テキスト.CSV", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE [契約者情報テキスト.CSV] SET [契約者情報テキスト.CSV].[FLG] = '0';"

    'Blank check — legacy_id（加工後）
    SQL = ""
    SQL = SQL & "UPDATE [契約者情報テキスト.CSV] AS T "
    SQL = SQL & "SET T.FLG = '1' "
    SQL = SQL & "WHERE Nz(T.[legacy_id（加工後）], '') = ''; "
    db.Execute SQL

    log_write "set_account:アカウント（入居者） → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Account SELECT "
    SQL = SQL & "    'X_' & T.ID as ID, "
    SQL = SQL & "    'Resident' as klass, "
    SQL = SQL & "    '10' as kind_id, "
    SQL = SQL & "    'chintai-app' as legacy_charge_user_id, "
    SQL = SQL & "    'アカウント（入居者）' as sheet "
    SQL = SQL & "FROM [契約者情報テキスト.CSV] AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 通常処理"

    ' TODO:AI_REVIEW — AI fallback failed: SKILL.md không tìm thấy
    ' Context: [{"目的": "契約形態のチェック用項目の作成", "該当ファイル名": "契約者情報テキスト.CSV", "該当項目名_1": "物件コード", "該当項目名_2": "契約者読み", "処理": "結合", "条件": "物件コードは下2桁削除した状態に加工してから結合\n契約テキスト.CSVに新たな項目として追加　項目名：判別キー", "開発用チェック": "True"}]

    ' TODO:AI_REVIEW — AI fallback failed: SKILL.md không tìm thấy
    ' Context: [{"該当ファイル名": "電話番号.CSV", "該当項目名_1": "管理契約No", "該当項目名_2": "読み(契)", "処理": "結合", "条件": "電話番号.CSVに新たな項目として追加　項目名：判別キー", "開発用チェック": "True"}]

    '特殊処理
    'tel_mobile -> 080、090、070から始まる場合
    SQL = ""
    SQL = SQL & "UPDATE [Account] AS T \nINNER JOIN [契約者情報テキスト.CSV] AS T1 ON T.[legacy_id] = T1.[入居者コード] "
    SQL = SQL & "SET T.[tel_mobile] = T1.[TEL1(入)] "
    SQL = SQL & "WHERE T1.[TEL1(入)] = '080、090、070から始まる場合'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 特殊処理: tel_mobile"

    '特殊処理
    'legacy_id, tel_mobile -> 080、090、070から始まらない、nullの場合
    SQL = ""
    SQL = SQL & "UPDATE [Account] AS T \nINNER JOIN [契約者情報テキスト.CSV] AS T1 ON T.[legacy_id] = T1.[入居者コード] "
    SQL = SQL & "SET T.[tel_mobile] = T1.[携帯(入)], "
    SQL = SQL & "    T.[legacy_id] = T1.[入居者コード] "
    SQL = SQL & "WHERE T1.[TEL1(入)] = '080、090、070から始まらない、null'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 特殊処理: legacy_id, tel_mobile"

    '特殊処理
    'tel_mobile -> 080、090、070から始まる場合
    SQL = ""
    SQL = SQL & "UPDATE [Account] AS T \nINNER JOIN [契約者情報テキスト.CSV] AS T1 ON T.[legacy_id] = T1.[契約者コード] "
    SQL = SQL & "SET T.[tel_mobile] = IIF(Nz(T.[tel_mobile], '') = '', T1.[TEL1(契)], T.[tel_mobile]), "
    SQL = SQL & "    T.[tel_mobile] = T1.[TEL1(契)] "
    SQL = SQL & "WHERE T1.[TEL1(契)] = '080、090、070から始まる場合'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 特殊処理: tel_mobile"

    '特殊処理
    'legacy_id, tel_mobile -> 080、090、070から始まらない、nullの場合
    SQL = ""
    SQL = SQL & "UPDATE [Account] AS T \nINNER JOIN [契約者情報テキスト.CSV] AS T1 ON T.[legacy_id] = T1.[契約者コード] "
    SQL = SQL & "SET T.[tel_mobile] = T1.[携帯(契)], "
    SQL = SQL & "    T.[legacy_id] = IIF(Nz(T.[legacy_id], '') = '', T1.[契約者コード], T.[legacy_id]) "
    SQL = SQL & "WHERE T1.[TEL1(契)] = '080、090、070から始まらない、null'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 特殊処理: legacy_id, tel_mobile"

    '特殊処理
    'tel_mobile -> 080、090、070から始まらない、nullの場合
    SQL = ""
    SQL = SQL & "UPDATE [Account] AS T \nINNER JOIN [契約者情報テキスト.CSV] AS T1 ON T.[legacy_id] = T1.[契約者コード] "
    SQL = SQL & "SET T.[tel_mobile] = T1.[携帯(契)] "
    SQL = SQL & "WHERE (T1.[TEL1(契)] = '080、090、070から始まらない、null' AND T1.[携帯(契)] = '080、090、070から始まる場合') AND NOT (T1.[TEL1(契)] = '080、090、070から始まる場合') AND NOT (T1.[TEL1(契)] = '080、090、070から始まらない、null'); "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 特殊処理: tel_mobile"

    '特殊処理
    'legacy_id, tel_mobile -> 080、090、070から始まらない、nullの場合
    SQL = ""
    SQL = SQL & "UPDATE [Account] AS T \nINNER JOIN [契約者情報テキスト.CSV] AS T1 ON T.[legacy_id] = T1.[契約者コード] "
    SQL = SQL & "SET T.[tel_mobile] = T1.[携帯(入)], "
    SQL = SQL & "    T.[legacy_id] = T1.[契約者コード] "
    SQL = SQL & "WHERE (T1.[携帯(契)] = '080、090、070から始まらない、null') AND NOT (T1.[TEL1(契)] = '080、090、070から始まらない、null'); "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 特殊処理: legacy_id, tel_mobile"

    '特殊処理
    'tel_fixed -> 紐づけできなかった場合
    SQL = ""
    SQL = SQL & "UPDATE [Account] AS T \nINNER JOIN [契約者メールアドレス.CSV] AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[tel_fixed] = IIF(Nz(T.[tel_fixed], '') = '', T1.[TEL1(契)], T.[tel_fixed]) "
    SQL = SQL & "WHERE T1.[メールアドレス] = '出力しない'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 特殊処理: tel_fixed"

    '特殊処理
    'fields: name_family_kana
    'conditions: ”法人”の場合, 契約者コード≠入居者コードの場合, 契約者コード＝入居者コードの場合
    SQL = ""
    SQL = SQL & "UPDATE [Account] AS T \nINNER JOIN [契約者情報テキスト.CSV] AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[name_family_kana] = SWITCH(  "
    SQL = SQL & "    T1.[契約者コード] = '契約者コード≠入居者コード' AND T1.[入居者コード] = '契約者コード≠入居者コード', T1.[入居者読み],  "
    SQL = SQL & "    T1.[契約形態] = '”法人”', T1.[入居者名],  "
    SQL = SQL & "    T1.[契約者コード] = '契約者コード＝入居者コード' AND T1.[入居者コード] = '契約者コード＝入居者コード', T1.[契約者名],  "
    SQL = SQL & "    True, NULL ) "
    SQL = SQL & "WHERE (T1.[契約者コード] = '契約者コード≠入居者コード' AND T1.[入居者コード] = '契約者コード≠入居者コード') OR (T1.[契約形態] = '”法人”') OR (T1.[契約者コード] = '契約者コード＝入居者コード' AND T1.[入居者コード] = '契約者コード＝入居者コード'); "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 特殊処理: name_family_kana"

    '特殊処理
    'fields: email
    'conditions: メールアドレスが重複している場合は一番最初の行の値を出力
    SQL = ""
    SQL = SQL & "UPDATE [Account] AS T \nINNER JOIN [契約者メールアドレス.CSV] AS T1 ON T.[legacy_id] = T1.[入居者コード] "
    SQL = SQL & "SET T.[email] = SWITCH(  "
    SQL = SQL & "    True, T1.[メールアドレス],  "
    SQL = SQL & "    True, T1.[メールアドレス],  "
    SQL = SQL & "    True, T1.[メールアドレス],  "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 特殊処理: email"

    '特殊処理
    'fields: tel_fixed
    'conditions: 紐づけできなかった場合
    SQL = ""
    SQL = SQL & "UPDATE [Account] AS T \nINNER JOIN [契約者メールアドレス②.CSV] AS T1 ON T.[legacy_id] = T1.[入居者コード] "
    SQL = SQL & "SET T.[tel_fixed] = SWITCH(  "
    SQL = SQL & "    True, T1.[TEL1(入)],  "
    SQL = SQL & "    True, T1.[TEL1(契)],  "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 特殊処理: tel_fixed"


    ' duplicate legacy_id
    SQL = ""
    SQL = SQL & "DELETE FROM [Account] "
    SQL = SQL & "WHERE ID LIKE 'X_%' "
    SQL = SQL & "AND ID NOT IN ( "
    SQL = SQL & "   SELECT MIN(ID) "
    SQL = SQL & "   FROM [Account] "
    SQL = SQL & "   WHERE ID LIKE 'X_%' "
    SQL = SQL & "   GROUP BY legacy_id "
    SQL = SQL & ") "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → delete duplicate"


    '■項目削除
    DeleteFieldInTable "契約者情報テキスト.CSV", "FLG"

    End If

    chk_required

    Debug.Print "Account:" & Timer - t
    log_write "set_account:out"

End Sub
