Attribute VB_Name = "契約（新規）"
Option Compare Database
Option Explicit

Sub set_contract()
    log_write "set_contract:in"

    Dim t As Single
    t = Timer

    Dim db As DAO.Database
    Dim SQL As String

    Set db = CurrentDb

    db.Execute "DELETE FROM Contract WHERE ID LIKE 'X_%';"

    If DCount("ID", "Contract", "ID LIKE 'X_*'") = 0 Then

    '================================================================================
    '契約（新規）
    '================================================================================

    '出力条件
    AddNewFieldToTable "新規契約更新一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 新規契約更新一覧 SET 新規契約更新一覧.[FLG] = '1';"

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "ファイル名"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "新規契約更新一覧.csv", "目的": "入居状況にない未来の契約を抽出", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約者3入居有無", "処理": "条件分岐", "条件": "\"入居有り\"の場合"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "GMO 入居状況一覧.csv", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約者3No", "処理": "そのまま出力"}]

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
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) AND (T1.[部屋No] = T2.[部屋No]) AND (T1.[上記の条件分岐で出力した契約者No] = T2.[上記の条件分岐で出力した契約者No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[物件No] IS NULL; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"目的": "重複削除", "該当ファイル名": "新規契約更新一覧.csv", "該当項目名_1": "物件No", "該当項目名_2": "部屋No", "該当項目名_3": "契約者No", "処理": "そのまま出力", "条件": "データ順で１番最初のデータを出力対象とする"}]

    log_write "set_contract:契約（新規） → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Contract SELECT "
    SQL = SQL & "    'X_' & T.ID as ID,"
    SQL = SQL & "    T.[物件No] & ""-"" & T.[部屋No] as legacy_property_id,"
    SQL = SQL & "    T.[契約者No] as legacy_resident_id,"
    SQL = SQL & "    T.[貸主No] as legacy_owner_id,"
    SQL = SQL & "    T.[物件No] as legacy_building_id,"
    SQL = SQL & "    T.[契約始期] as start_from,"
    SQL = SQL & "    T.[契約終期] as end_until,"
    SQL = SQL & "    T.[家賃保証会社名] as corporate_guarantor_name,"
    SQL = SQL & "    T.[物件No] & ""-"" & T.[部屋No] & ""-"" & T.[契約者No] as legacy_id,"
    SQL = SQL & "    '契約（新規）' as sheet"
    SQL = SQL & "FROM 新規契約更新一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_contract:契約（新規） → 通常処理"


    '■項目削除
    DeleteFieldInTable "新規契約更新一覧", "FLG"

    End If

    chk_required

    Debug.Print "Contract:" & Timer - t
    log_write "set_contract:out"

End Sub
