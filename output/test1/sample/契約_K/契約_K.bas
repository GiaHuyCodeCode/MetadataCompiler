Attribute VB_Name = "契約_K"
Option Compare Database
Option Explicit

Sub set_contract()
    log_write "set_contract:in"

    Dim t As Single
    t = Timer

    Dim db As ADODB.Connection
    Dim SQL As String

    Set db = CurrentProject.Connection

    db.Execute "DELETE FROM Contract WHERE ID LIKE 'K_%';"

    If DCount("ID", "Contract", "ID LIKE 'K_*'") = 0 Then

    '================================================================================
    '契約_K
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

    'Blank check — 物件No
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T "
    SQL = SQL & "SET T.FLG = '1' "
    SQL = SQL & "WHERE Nz(T.[物件No], '') = ''; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"使用ファイル": "解約情報一覧.csv"}]

    log_write "set_contract:契約_K → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Contract SELECT "
    SQL = SQL & "    'K_' & T.ID as ID,"
    SQL = SQL & "    T.[物件No] & ""-"" & T.[部屋No] as legacy_property_id,"
    SQL = SQL & "    T.[家主No] as legacy_owner_id,"
    SQL = SQL & "    T.[物件No] as legacy_building_id,"
    SQL = SQL & "    T.[初回契約始期] as start_from,"
    SQL = SQL & "    T.[契約終期] as end_until,"
    SQL = SQL & "    T.[家賃保証会社名] as corporate_guarantor_name,"
    SQL = SQL & "    T.[物件No] & ""-"" & T.[部屋No] & ""-"" & T.[契約者1No] as legacy_id,"
    SQL = SQL & "    T.[内　容(メールBOX)] as mail_box_no,"
    SQL = SQL & "    T.[電気業者TEL1] as electricity_contact_tel,"
    SQL = SQL & "    T.[上水道業者TEL1] as drink_water_contact_tel,"
    SQL = SQL & "    T.[排水業者TEL1] as drainage_contact_tel,"
    SQL = SQL & "    T.[ガス業者TEL1] as gas_contact_tel,"
    SQL = SQL & "    T.[その他業者1名] as internet_contact_name,"
    SQL = SQL & "    T.[その他業者1TEL1] as internet_contact_tel,"
    SQL = SQL & "    '契約_K' as sheet"
    SQL = SQL & "FROM 入居状況一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_contract:契約_K → 通常処理"

    '特殊処理

    ' Dim con — điều kiện ưu tiên 契約者1→3→2 (khai báo 1 lần)
    Dim con1 As String
    con1 = "(T1.[個人・法人区分_33] = '個人' AND T1.[個人・法人区分_43] = '個人' AND T1.[契約者1入居有無] = '入居有り' AND T1.[契約者2入居有無] = '入居有り')"
    Dim con3 As String
    con3 = "(T1.[契約者3入居有無] = '入居有り"場合')"
    Dim con2 As String
    con2 = "(T1.[契約者2入居有無] = '入居有り"場合')"

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('K_' & T1.ID) "
    SQL = SQL & "SET T.[legacy_resident_id] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No] ); "
    db.Execute SQL
    log_write "set_contract:契約_K → 特殊処理: legacy_resident_id"

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('K_' & T1.ID) "
    SQL = SQL & "SET T.[cohabitant_names] = SWITCH( "
    SQL = SQL & "    (T1.[契約者1入居有無] = '入居有り' AND T1.[契約者2入居有無] = '入居有り'), T1.[契約者名_45], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_contract:契約_K → 特殊処理: cohabitant_names"

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"項目名": "end_date", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "物件No", "該当項目名_2": "部屋No", "該当項目名_3": "契約者1No", "処理": "紐づけ", "開発用チェック": "True"}, {"該当ファイル名": "解約情報一覧.csv", "該当項目名_1": "物件No", "該当項目名_2": "部　屋", "該当項目名_3": "契約者No", "処理": "紐づけ", "開発用チェック": "True"}]

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('K_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    T1.[契約者3入居有無] = '入居有り', '20', "
    SQL = SQL & "    T1.[契約者2入居有無] = '入居有り', '20', "
    SQL = SQL & "    T1.[個人・法人区分_33] = '法人', '20', "
    SQL = SQL & "    True, '10' ); "
    db.Execute SQL
    log_write "set_contract:契約_K → 特殊処理: kind_id"

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('K_' & T1.ID) "
    SQL = SQL & "SET T.[legacy_contractor_id] = SWITCH( "
    SQL = SQL & "    T1.[契約者3入居有無] = '「該当項目名_1」OR「該当項目名_2」が""入居有り' AND T1.[契約者2入居有無] = '「該当項目名_1」OR「該当項目名_2」が""入居有り' AND T1.[個人・法人区分_33] = '法人', T1.[契約者1No], "
    SQL = SQL & "    T1.[契約者1入居有無] = 'ブランク', T1.[契約者1No], "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_contract:契約_K → 特殊処理: legacy_contractor_id"

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T1 ON T.ID = ('K_' & T1.ID) "
    SQL = SQL & "SET T.[enable_support_24h] = SWITCH( "
    SQL = SQL & "    T1.[金額(税込額)(安心サービス24)] = 'いずれかの項目がブランクではない場合' AND T1.[金額(税込額)(24時間コールセンター)] = 'いずれかの項目がブランクではない場合' AND T1.[金額(税込額)(安心サポート)] = 'いずれかの項目がブランクではない場合', '1', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_contract:契約_K → 特殊処理: enable_support_24h"

    ' --- AI GENERATED (sk-architect) ---
    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN [GMO 入居状況一覧] AS T1 ON T.ID = ('K_' & T1.ID) "
    SQL = SQL & "SET T.electricity_contact_name = T1.[電気業者名] "
    SQL = SQL & "WHERE T1.[電気業者名] NOT LIKE '%連絡不要%'; "
    db.Execute SQL
    log_write "set_contract:契約_K → 特殊処理: electricity_contact_name"
    ' --- END AI GENERATED ---

    ' --- AI GENERATED (sk-architect) ---
    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN [GMO 入居状況一覧] AS T1 ON T.ID = ('K_' & T1.ID) "
    SQL = SQL & "SET T.drink_water_contact_name = T1.[上水道業者名] "
    SQL = SQL & "WHERE T1.[上水道業者名] NOT LIKE '%連絡不要%'; "
    db.Execute SQL
    log_write "set_contract:契約_K → 特殊処理: drink_water_contact_name"
    ' --- END AI GENERATED ---

    ' --- AI GENERATED (sk-architect) ---
    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN [GMO 入居状況一覧] AS T1 ON T.ID = ('K_' & T1.ID) "
    SQL = SQL & "SET T.gas_contact_name = T1.[ガス業者名] "
    SQL = SQL & "WHERE T1.[ガス業者名] NOT LIKE '%連絡不要%' OR T1.[ガス業者名] NOT LIKE '%田中総合燃料%'; "
    db.Execute SQL
    log_write "set_contract:契約_K → 特殊処理: gas_contact_name"
    ' --- END AI GENERATED ---

    ' --- AI GENERATED (sk-architect) ---
    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN [GMO 入居状況一覧] AS T1 ON T.ID = ('K_' & T1.ID) "
    SQL = SQL & "SET T.drainage_contact_name = T1.[排水業者名] "
    SQL = SQL & "WHERE T1.[排水業者名] NOT LIKE '%連絡不要%' OR T1.[排水業者名] NOT LIKE '%汲み取り業者%'; "
    db.Execute SQL
    log_write "set_contract:契約_K → 特殊処理: drainage_contact_name"
    ' --- END AI GENERATED ---


    '■項目削除
    DeleteFieldInTable "入居状況一覧", "FLG"

    End If

    chk_required

    Debug.Print "Contract:" & Timer - t
    log_write "set_contract:out"

End Sub
