Attribute VB_Name = "アカウント（新規入居者）_N"
Option Compare Database
Option Explicit

Sub set_account()
    log_write "set_account:in"

    Dim t As Single
    t = Timer

    Dim db As DAO.Database
    Dim SQL As String

    Set db = CurrentDb

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
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) AND (T1.[部屋No] = T2.[部屋No]) AND (T1.[契約者1No] = T2.[契約者1No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[物件No] IS NULL; "
    db.Execute SQL

    '重複チェック legacy_idをキーに重複を削除
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 "
    SQL = SQL & "SET 新規契約更新一覧.FLG = '1' "
    SQL = SQL & "WHERE 新規契約更新一覧.ID NOT IN ( "
    SQL = SQL & "    SELECT MIN(ID) "
    SQL = SQL & "    FROM 新規契約更新一覧 "
    SQL = SQL & "    WHERE FLG = '0' "
    SQL = SQL & "    GROUP BY [legacy_id] "
    SQL = SQL & "); "
    db.Execute SQL

    ' --- AI GENERATED (sk-architect) ---
    ' sử dụng file 契約者情報.csv (No action required)
    ' --- END AI GENERATED ---

    log_write "set_account:アカウント（新規入居者）_N → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Account SELECT "
    SQL = SQL & "    'N_' & T.ID as ID,"
    SQL = SQL & "    'Resident' as klass,"
    SQL = SQL & "    'gmo002' as legacy_charge_user_id,"
    SQL = SQL & "    'アカウント（新規入居者）_N' as sheet"
    SQL = SQL & "FROM 新規契約更新一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_account:アカウント（新規入居者）_N → 通常処理"

    '特殊処理

    ' Dim con — điều kiện ưu tiên 契約者1→3→2 (khai báo 1 lần)
    Dim con3 As String
    con3 = "(T1.[契約者3入居有無] = '入居有り')"
    Dim con2 As String
    con2 = "(T1.[契約者3入居有無] <> '入居有り' AND T1.[契約者2入居有無] = '入居有り')"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID) "
    SQL = SQL & "SET T.[tel_mobile] = SWITCH( "
    SQL = SQL & "    (T1.[個人・法人区分_16] = '個人' AND T1.[個人・法人区分_19] = '個人'), T1.[携帯1], "
    SQL = SQL & "    (T1.[契約者3入居有無] = '入居有り'), T1.[優先メール], "
    SQL = SQL & "    (T1.[契約者2入居有無] = '入居有り'), T1.[携帯番号/SMS], "
    SQL = SQL & "    (True), T1.[メールアドレス], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（新規入居者）_N → 特殊処理: tel_mobile"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID) "
    SQL = SQL & "SET T.[legacy_id] = SWITCH( "
    SQL = SQL & "    T1.[個人・法人区分_16] = '個人' AND T1.[個人・法人区分_19] = '個人', T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ); "
    db.Execute SQL
    log_write "set_account:アカウント（新規入居者）_N → 特殊処理: legacy_id"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    (T1.[個人・法人区分_16] = '個人' AND T1.[個人・法人区分_19] = '個人'), T1.[契約者名カナ], "
    SQL = SQL & "    (T1.[契約者3入居有無] = '入居有り'), T1.[契約者名], "
    SQL = SQL & "    (T1.[契約者2入居有無] = '入居有り'), T1.[契約者名カナ], "
    SQL = SQL & "    (True), T1.[契約者名], "
    SQL = SQL & "    (T1.[契約者区分] = '上記で出力したデータの個人・法人区分が""個人'), '20', "
    SQL = SQL & "    (True), '10', "
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
