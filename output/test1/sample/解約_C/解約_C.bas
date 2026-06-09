Attribute VB_Name = "解約_C"
Option Compare Database
Option Explicit

Sub set_data()
    log_write "set_data:in"

    Dim t As Single
    t = Timer

    Dim db As ADODB.Connection
    Dim SQL As String

    Set db = CurrentProject.Connection

    db.Execute "DELETE FROM TargetTable WHERE ID LIKE 'C_%';"

    If DCount("ID", "TargetTable", "ID LIKE 'C_*'") = 0 Then

    '================================================================================
    '解約_C
    '================================================================================

    '出力条件
    AddNewFieldToTable "契約情報管理", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 契約情報管理 SET 契約情報管理.[FLG] = '1';"

    'JOIN filter → 紐づかないデータを抽出
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 AS T1 "
    SQL = SQL & "LEFT JOIN 解約情報一覧 AS T2 "
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) AND (T1.[部屋No] = T2.[部屋No]) AND (T1.[契約者1No] = T2.[契約者1No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[物件No] IS NULL; "
    db.Execute SQL

    'JOIN filter
    SQL = ""
    SQL = SQL & "UPDATE 解約情報一覧 AS T1 "
    SQL = SQL & "INNER JOIN 契約情報管理 AS T2 "
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) AND (T1.[部　屋] = T2.[部　屋]) AND (T1.[契約者No] = T2.[契約者No]) "
    SQL = SQL & "SET T1.FLG = '1'; "
    db.Execute SQL

    ' 出力 (No filter required)
    log_write "set_data:解約_C → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO TargetTable SELECT "
    SQL = SQL & "    'C_' & T.ID as ID,"
    SQL = SQL & "    T.[基幹部屋ID] as legacy_property_id,"
    SQL = SQL & "    T.[基幹入居者ID] as legacy_resident_id,"
    SQL = SQL & "    T.[基幹オーナーID] as legacy_owner_id,"
    SQL = SQL & "    T.[基幹建物ID] as legacy_building_id,"
    SQL = SQL & "    IIF(IsDate(T.[賃貸開始]), Format(T.[賃貸開始], 'yyyy-mm-dd'), NULL) as start_from,"
    SQL = SQL & "    IIF(IsDate(T.[賃貸終了]), Format(T.[賃貸終了], 'yyyy-mm-dd'), NULL) as end_until,"
    SQL = SQL & "    Val(Replace(T.[賃料], '円', '')) as charge_rent,"
    SQL = SQL & "    Val(Replace(T.[共益費], '円', '')) as charge_management,"
    SQL = SQL & "    Val(Replace(T.[駐車料], '円', '')) as charge_parking,"
    SQL = SQL & "    Val(Replace(T.[修繕積立金], '円', '')) as charge_accumulation,"
    SQL = SQL & "    Val(Replace(T.[その他], '円', '')) as charge_additional,"
    SQL = SQL & "    Val(Replace(T.[敷金], '円', '')) as charge_deposit,"
    SQL = SQL & "    Val(Replace(T.[礼金], '円', '')) as charge_gratuity,"
    SQL = SQL & "    Val(Replace(T.[保証料], '円', '')) as charge_bond,"
    SQL = SQL & "    Val(Replace(T.[更新料], '円', '')) as months_of_rent_for_renewal,"
    SQL = SQL & "    T.[連帯保証人] as individual_guarantor_name,"
    SQL = SQL & "    T.[保証会社] as corporate_guarantor_name,"
    SQL = SQL & "    T.[同居人] as cohabitant_names,"
    SQL = SQL & "    T.[備考] as note,"
    SQL = SQL & "    IIF(IsDate(T.[入居日]), Format(T.[入居日], 'yyyy-mm-dd'), NULL) as property_checkup_from,"
    SQL = SQL & "    NULL as end_date,"
    SQL = SQL & "    T.[基幹契約ID] as legacy_id,"
    SQL = SQL & "    T.[24時間サポート] as enable_support_24h,"
    SQL = SQL & "    T.[解約予告期間] as termination_by_day,"
    SQL = SQL & "    T.[メールボックスNo.] as mail_box_no,"
    SQL = SQL & "    T.[【電気】連絡先名] as electricity_contact_name,"
    SQL = SQL & "    T.[【電気】連絡先TEL] as electricity_contact_tel,"
    SQL = SQL & "    T.[【水道(飲料水)】連絡先名] as drink_water_contact_name,"
    SQL = SQL & "    T.[【水道(飲料水)】連絡先TEL] as drink_water_contact_tel,"
    SQL = SQL & "    T.[【水道(排水)】連絡先名] as drainage_contact_name,"
    SQL = SQL & "    T.[【水道(排水)】連絡先TEL] as drainage_contact_tel,"
    SQL = SQL & "    T.[【ガス】連絡先名] as gas_contact_name,"
    SQL = SQL & "    T.[【ガス】連絡先TEL] as gas_contact_tel,"
    SQL = SQL & "    T.[【給湯】連絡先名] as hot_water_contact_name,"
    SQL = SQL & "    T.[【給湯】連絡先TEL] as hot_water_contact_tel,"
    SQL = SQL & "    T.[【インターネット】連絡先名] as internet_contact_name,"
    SQL = SQL & "    T.[【インターネット】連絡先TEL] as internet_contact_tel,"
    SQL = SQL & "    T.[【ケーブルテレビ】連絡先名] as cable_tv_contact_name,"
    SQL = SQL & "    T.[【ケーブルテレビ】連絡先TEL] as cable_tv_contact_tel,"
    SQL = SQL & "    '解約_C' as sheet"
    SQL = SQL & "FROM 契約情報管理 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_data:解約_C → 通常処理"

    '特殊処理

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"項目名": "legacy_contractor_id", "該当ファイル名": "契約情報管理.csv", "該当項目名_1": "契約者", "処理": "紐づけ", "開発用チェック": "True"}, {"該当ファイル名": "契約者情報管理.csv", "該当項目名_1": "氏名/会社名", "処理": "紐づけ", "開発用チェック": "True"}]

    SQL = ""
    SQL = SQL & "UPDATE TargetTable AS T "
    SQL = SQL & "INNER JOIN 契約情報管理 AS T1 ON T.ID = ('C_' & T1.ID) "
    SQL = SQL & "SET T.[sublease] = SWITCH( "
    SQL = SQL & "    T1.[管理方法] = 'サブリース', '1', "
    SQL = SQL & "    T1.[管理方法] = '管理委託', '0', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_data:解約_C → 特殊処理: sublease"

    SQL = ""
    SQL = SQL & "UPDATE TargetTable AS T "
    SQL = SQL & "INNER JOIN 契約情報管理 AS T1 ON T.ID = ('C_' & T1.ID) "
    SQL = SQL & "SET T.[kind_id] = SWITCH( "
    SQL = SQL & "    T1.[個人/法人] = '個人', '10', "
    SQL = SQL & "    T1.[個人/法人] = '法人', '20', "
    SQL = SQL & "    True, NULL ); "
    db.Execute SQL
    log_write "set_data:解約_C → 特殊処理: kind_id"


    '■項目削除
    DeleteFieldInTable "契約情報管理", "FLG"

    End If

    chk_required

    Debug.Print "TargetTable:" & Timer - t
    log_write "set_data:out"

End Sub
