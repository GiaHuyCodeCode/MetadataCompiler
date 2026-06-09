Attribute VB_Name = "契約"
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
    '契約
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

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "物件管理情報一覧.csv"}]

    'Blank check — legacy_id（加工後）
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T "
    SQL = SQL & "SET T.FLG = '1' "
    SQL = SQL & "WHERE Nz(T.[legacy_id（加工後）], '') = ''; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"使用ファイル": "解約情報一覧.csv"}]

    log_write "set_contract:契約 → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Contract SELECT "
    SQL = SQL & "    'X_' & T.ID as ID,"
    SQL = SQL & "    T.[物件No] & ""-"" & T.[部屋No] as legacy_property_id,"
    SQL = SQL & "    T.[家主No] as legacy_owner_id,"
    SQL = SQL & "    T.[物件No] as legacy_building_id,"
    SQL = SQL & "    T.[家賃保証会社名] as corporate_guarantor_name,"
    SQL = SQL & "    T.[ガス業者名] as gas_contact_name,"
    SQL = SQL & "    T.[ガス業者TEL1] as gas_contact_tel,"
    SQL = SQL & "    '契約' as sheet"
    SQL = SQL & "FROM 入居状況一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_contract:契約 → 通常処理"

    '特殊処理

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[sublease] = SWITCH( "
    SQL = SQL & "    (T1.[管理形態] = '紐づく場合かつ管理形態が""一括借上'), '1.0', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_contract:契約 → 特殊処理: sublease"

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"項目名": "end_date", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "物件No", "該当項目名_2": "部屋No", "該当項目名_3": "legacy_resident_id", "処理": "紐づけ", "_target_field": "end_date"}, {"該当ファイル名": "解約情報一覧.csv", "該当項目名_1": "物件No", "該当項目名_2": "部　屋", "該当項目名_3": "契約者No", "処理": "紐づけ", "_target_field": "end_date"}]

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[legacy_contractor_id] = SWITCH( "
    SQL = SQL & "    (True), T1.[契約者No], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_contract:契約 → 特殊処理: legacy_contractor_id"

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"項目名": "legacy_id", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "物件No", "該当項目名_2": "部屋No", "該当項目名_3": "legacy_contractor_id", "処理": "ハイフン付出力", "条件 / 項目マッピング": "\"legacy_contractor_id\"がブランクではない場合", "_target_field": "legacy_id"}, {"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "物件No", "該当項目名_2": "部屋No", "該当項目

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "電気業者名", "処理": "そのまま出力", "条件 / 項目マッピング": "\"電気　家主一括契約\"または\"水道　家主一括契約\"を含まない場合", "_target_field": ""}, {"項目名": "electricity_contact_name", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "電気業者名", "_target_field": "electricity_contact_name"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "上水道業者名", "処理": "そのまま出力", "条件 / 項目マッピング": "\"水道　家主一括契約\"を含まない場合", "_target_field": ""}, {"項目名": "drink_water_contact_name", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "上水道業者名", "_target_field": "drink_water_contact_name"}]

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[charge_gratuity] = SWITCH( "
    SQL = SQL & "    (True), '0.0', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_contract:契約 → 特殊処理: charge_gratuity"

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    (T1.[契約者区分] = ''), '10.0', "
    SQL = SQL & "    (T1.[契約状況] = '契約中(他社)'), '10.0', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_contract:契約 → 特殊処理: kind_id"

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"項目名": "start_from", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "初回契約始期", "処理": "そのまま出力", "条件 / 項目マッピング": "データがある場合", "_target_field": "start_from"}, {"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約状況", "処理": "固定値出力", "条件 / 項目マッピング": "\"契約中(他社)\"の場合", "_target_field": "start_from"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"項目名": "end_until", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約終期", "処理": "そのまま出力", "条件 / 項目マッピング": "データがある場合", "_target_field": "end_until"}, {"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "契約状況", "処理": "固定値出力", "条件 / 項目マッピング": "\"契約中(他社)\"の場合", "_target_field": "end_until"}]


    '■項目削除
    DeleteFieldInTable "入居状況一覧", "FLG"

    End If

    chk_required

    Debug.Print "Contract:" & Timer - t
    log_write "set_contract:out"

End Sub
