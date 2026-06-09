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
    SQL = SQL & "    T.[初回契約始期] as start_from,"
    SQL = SQL & "    T.[契約終期] as end_until,"
    SQL = SQL & "    T.[家賃保証会社名] as corporate_guarantor_name,"
    SQL = SQL & "    T.[ガス業者名] as gas_contact_name,"
    SQL = SQL & "    T.[ガス業者TEL1] as gas_contact_tel,"
    SQL = SQL & "    '契約' as sheet"
    SQL = SQL & "FROM 入居状況一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_contract:契約 → 通常処理"

    '特殊処理

    ' Dim con — điều kiện ưu tiên 契約者1→3→2 (khai báo 1 lần)
    Dim con1 As String
    con1 = "(T1.[契約者1入居有無] = '入居有り')"
    Dim con3 As String
    con3 = "(T1.[契約者1入居有無] <> '入居有り' AND T1.[契約者3入居有無] = '入居有り')"
    Dim con2 As String
    con2 = "(T1.[契約者1入居有無] <> '入居有り' AND T1.[契約者3入居有無] <> '入居有り' AND T1.[契約者2入居有無] = '入居有り')"

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('X_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    (T1.[契約者3入居有無] = '入居有り'), T1.[契約者3No], "
    SQL = SQL & "    (T1.[契約者2入居有無] = '入居有り'), T1.[契約者2No], "
    SQL = SQL & "    (T1.[契約者1入居有無] = '入居有り'), T1.[契約者1No], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_contract:契約 → 特殊処理: kind_id"

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
    SQL = SQL & "    (T1.[上記で出力した契約者No] = '契約者No""≠""legacy_resident_id' AND T1.[legacy_resident_id] = '契約者No""≠""legacy_resident_id'), T1.[契約者No], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_contract:契約 → 特殊処理: legacy_contractor_id"

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"項目名": "legacy_id", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "物件No", "該当項目名_2": "部屋No", "該当項目名_3": "legacy_contractor_id", "処理": "ハイフン付出力", "条件 / 項目マッピング": "\"legacy_contractor_id\"がブランクではない場合", "_target_field": "legacy_id"}, {"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "物件No", "該当項目名_2": "部屋No", "該当項目

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "電気業者名", "処理": "そのまま出力", "条件 / 項目マッピング": "\"電気　家主一括契約\"または\"水道　家主一括契約\"を含まない場合", "_target_field": ""}, {"項目名": "electricity_contact_name", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "電気業者名", "_target_field": "electricity_contact_name"}]

    ' TODO:AI_REVIEW — AI fallback failed: GOOGLE_API_KEY chưa set
    ' Context: [{"該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "上水道業者名", "処理": "そのまま出力", "条件 / 項目マッピング": "\"水道　家主一括契約\"を含まない場合", "_target_field": ""}, {"項目名": "drink_water_contact_name", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "上水道業者名", "_target_field": "drink_water_contact_name"}]


    '■項目削除
    DeleteFieldInTable "入居状況一覧", "FLG"

    End If

    chk_required

    Debug.Print "Contract:" & Timer - t
    log_write "set_contract:out"

End Sub
