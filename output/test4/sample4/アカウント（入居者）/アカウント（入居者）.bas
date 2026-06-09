Attribute VB_Name = "アカウント（入居者）"
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
    'アカウント（入居者）
    '================================================================================

    '出力条件
    AddNewFieldToTable "入居状況一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 入居状況一覧 SET 入居状況一覧.[FLG] = '1';"

    '■有効レコードのみFLG=0に戻す
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T "
    SQL = SQL & "SET T.FLG = '0' "
    SQL = SQL & "WHERE T.[契約状況] = '契約中' OR T.[契約状況] = '契約中(他社)' OR T.[契約状況] = '解約予定""'; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "GMO 入居状況一覧.csv"}]

    'Blank check — legacy_id（加工後）
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T "
    SQL = SQL & "SET T.FLG = '1' "
    SQL = SQL & "WHERE Nz(T.[legacy_id（加工後）], '') = ''; "
    db.Execute SQL

    log_write "set_account:アカウント（入居者） → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Account SELECT "
    SQL = SQL & "    'X_' & T.ID as ID,"
    SQL = SQL & "    'Resident' as klass,"
    SQL = SQL & "    'ohta-ff' as legacy_charge_user_id,"
    SQL = SQL & "    'アカウント（入居者）' as sheet"
    SQL = SQL & "FROM 入居状況一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 通常処理"

    '特殊処理

    ' Dim con — điều kiện ưu tiên 契約者1→3→2 (khai báo 1 lần)
    Dim con1 As String
    con1 = "(T1.[契約者1入居有無] = '入居有り')"
    Dim con3 As String
    con3 = "(T1.[契約者1入居有無] <> '入居有り' AND T1.[契約者3入居有無] = '入居有り')"
    Dim con2 As String
    con2 = "(T1.[契約者1入居有無] <> '入居有り' AND T1.[契約者3入居有無] <> '入居有り' AND T1.[契約者2入居有無] = '入居有り')"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    (T1.[契約者3入居有無] = '入居有り'), T1.[契約者名_12], "
    SQL = SQL & "    (T1.[契約者2入居有無] = '入居有り'), T1.[契約者名_4], "
    SQL = SQL & "    (T1.[契約者1入居有無] = '入居有り'), T1.[契約者名], "
    SQL = SQL & "    (T1.[個人・法人区分_1] = '上記で出力したデータの個人・法人区分が""個人' AND T1.[個人・法人区分_3] = '上記で出力したデータの個人・法人区分が""個人' AND T1.[個人・法人区分_11] = '上記で出力したデータの個人・法人区分が""個人'), '20', "
    SQL = SQL & "    (T1.[個人・法人区分_1] = '上記で出力したデータの個人・法人区分が""法人' AND T1.[個人・法人区分_3] = '上記で出力したデータの個人・法人区分が""法人' AND T1.[個人・法人区分_11] = '上記で出力したデータの個人・法人区分が""法人'), '10', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（入居者） → 特殊処理: kind_id"


    '■項目削除
    DeleteFieldInTable "入居状況一覧", "FLG"

    End If

    chk_required

    Debug.Print "Account:" & Timer - t
    log_write "set_account:out"

End Sub
