Attribute VB_Name = "契約者"
Option Compare Database
Option Explicit

Sub set_contract()
    log_write "set_contract:in"

    Dim t As Single
    t = Timer

    Dim db As ADODB.Connection
    Dim SQL As String

    Set db = CurrentProject.Connection

    db.Execute "DELETE FROM Contract WHERE ID LIKE 'X_%';"

    If DCount("ID", "Contract", "ID LIKE 'X_*'") = 0 Then

    '================================================================================
    '契約者
    '================================================================================

    '出力条件
    AddNewFieldToTable "入居状況一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 入居状況一覧 SET 入居状況一覧.[FLG] = '1';"

    '■有効レコードのみFLG=0に戻す
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T "
    SQL = SQL & "SET T.FLG = '0' "
    SQL = SQL & "WHERE T.[契約状況] = '契約中' OR T.[契約状況] = '”解約予定”'; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"使用ファイル": "GMO 入居状況一覧.csv", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約者3入居有無", "該当項目名_2": "契約者2入居有無", "処理": "条件分岐", "条件": "「該当項目名_1」OR「該当項目名_2」が\"入居有り\"の場合", "備考": "AND", "開発用チェック": "True"}]

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "個人・法人区分_33", "処理": "条件分岐", "条件": "\"法人\"の場合", "備考": "AND", "開発用チェック": "True"}]

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約者1No", "処理": "出力", "条件": "条件分岐に当てはまる", "開発用チェック": "True"}]

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "処理": "条件分岐", "条件": "条件分岐に当てはまらない", "備考": "AND", "開発用チェック": "True"}]

    '■有効レコードのみFLG=0に戻す
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T "
    SQL = SQL & "SET T.FLG = '0' "
    SQL = SQL & "WHERE T.[契約状況] = '契約中' OR T.[契約状況] = '”解約予定”'; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約者1入居有無", "処理": "条件分岐", "条件": "ブランクの場合", "備考": "AND", "開発用チェック": "True"}]

    ' 出力 (No filter required)
    log_write "set_contract:契約者 → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Contract SELECT "
    SQL = SQL & "    'X_' & T.ID as ID,"
    SQL = SQL & "    T.[契約者名_35] as name,"
    SQL = SQL & "    T.[優先メールアドレス_38] as email,"
    SQL = SQL & "    T.[携帯1_40] as tel_mobile,"
    SQL = SQL & "    T.[契約者1No] as legacy_id,"
    SQL = SQL & "    '契約者' as sheet"
    SQL = SQL & "FROM 入居状況一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_contract:契約者 → 通常処理"

    '特殊処理

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    T1.[個人・法人区分_33] = '個人', '10', "
    SQL = SQL & "    T1.[個人・法人区分_33] = '法人', '20', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_contract:契約者 → 特殊処理: kind_id"

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"項目名": "legacy_id", "目的": "すべてのデータ出力後にcontractorファイル内で重複チェックして重複削除", "処理": "重複チェック", "開発用チェック": "True"}]


    ' duplicate legacy_id
    SQL = ""
    SQL = SQL & "DELETE FROM Contract "
    SQL = SQL & "WHERE ID NOT IN ( "
    SQL = SQL & "   SELECT MIN(ID) "
    SQL = SQL & "   FROM Contract "
    SQL = SQL & "   GROUP BY legacy_id "
    SQL = SQL & ") "
    db.Execute SQL
    log_write "set_contract:契約者 → delete duplicate"


    '■項目削除
    DeleteFieldInTable "入居状況一覧", "FLG"

    End If

    chk_required

    Debug.Print "Contract:" & Timer - t
    log_write "set_contract:out"

End Sub
