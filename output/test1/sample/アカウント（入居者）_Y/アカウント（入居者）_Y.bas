Attribute VB_Name = "アカウント（入居者）_Y"
Option Compare Database
Option Explicit

Sub set_account()
    log_write "set_account:in"

    Dim t As Single
    t = Timer

    Dim db As ADODB.Connection
    Dim SQL As String

    Set db = CurrentProject.Connection

    db.Execute "DELETE FROM Account WHERE ID LIKE 'Y_%';"

    If DCount("ID", "Account", "ID LIKE 'Y_*'") = 0 Then

    '================================================================================
    'アカウント（入居者）_Y
    '================================================================================

    '出力条件
    AddNewFieldToTable "入居状況一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 入居状況一覧 SET 入居状況一覧.[FLG] = '1';"

    '■有効レコードのみFLG=0に戻す
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T "
    SQL = SQL & "SET T.FLG = '0' "
    SQL = SQL & "WHERE T.[契約状況] = '契約中' OR T.[契約状況] = '解約予定'; "
    db.Execute SQL

    ' 出力 (No filter required)
    log_write "set_account:アカウント（入居者）_Y → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Account SELECT "
    SQL = SQL & "    'Y_' & T.ID as ID,"
    SQL = SQL & "    'Resident' as klass,"
    SQL = SQL & "    'gmo002' as legacy_charge_user_id,"
    SQL = SQL & "    'アカウント（入居者）_Y' as sheet"
    SQL = SQL & "FROM 入居状況一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者）_Y → 通常処理"

    '特殊処理

    ' Dim con — điều kiện ưu tiên 契約者1→3→2 (khai báo 1 lần)
    Dim con1 As String
    con1 = "(T1.[個人・法人区分_33] = '個人' AND T1.[個人・法人区分_43] = '個人' AND T1.[契約者1入居有無] = '入居有り' AND T1.[契約者2入居有無] = '入居有り')"
    Dim con3 As String
    con3 = "(T1.[契約者3入居有無] = '入居有り')"
    Dim con2 As String
    con2 = "(T1.[契約者2入居有無] = '入居有り')"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('Y_' & T1.ID) "
    SQL = SQL & "SET T.[tag] = SWITCH( "
    SQL = SQL & "    (T1.[契約状況] = '解約予定'), T1.[契約状況], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_account:アカウント（入居者）_Y → 特殊処理: tag"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "SET T.[name_family] = '契約者名_35', T.[company_name] = '契約者名_35', T.[name_family_kana] = '契約者カナ_36', T.[company_name_kana] = '契約者カナ_36', T.[email] = '優先メールアドレス_38', T.[tel_fixed] = 'TEL1_39', T.[tel_mobile] = '携帯1_40', T.[kind_id] = '個人・法人区分_33' "
    SQL = SQL & "WHERE T.ID LIKE '%Y_%' AND T.kind_id = '個人'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者）_Y → 上記で出力したデータの個人・法人区分が'個人'の場合"

    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "SET T.[name_family] = '', T.[name_family_kana] = '' "
    SQL = SQL & "WHERE T.ID LIKE '%Y_%' AND T.kind_id = '法人'; "
    db.Execute SQL
    log_write "set_account:アカウント（入居者）_Y → 上記で出力したデータの個人・法人区分が'法人'の場合"


    '■項目削除
    DeleteFieldInTable "入居状況一覧", "FLG"

    End If

    chk_required

    Debug.Print "Account:" & Timer - t
    log_write "set_account:out"

End Sub
