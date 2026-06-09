Attribute VB_Name = "output_contract"
Option Compare Database
Option Explicit

Sub set_contract()

    log_write "set_contract:in"
    
    Dim t As Single
    t = Timer
    
    Dim db As ADODB.Connection
    Dim rs1 As ADODB.Recordset
    Dim SQL As String
    Dim SQL1 As String
    Dim i As Long
    
    Dim ERR As Long
    Dim s_buf As String
    Dim v_sum As Variant
    
    Dim s_name As String
    Dim s_building As String
    Dim s_room As String
    Dim s_from As String
    Dim s_next As String
    Dim frmName As String
    Dim vOLD As String

    
    Set db = CurrentProject.Connection
    
    
    db.Execute "DELETE FROM Contract"
    log_write "set_contract:Contract_delete"
    
    If DCount("*", "家主基本情報") <= 0 Or _
        DCount("*", "契約者情報") <= 0 Or _
        DCount("*", "物件管理情報一覧") <= 0 Or _
        DCount("*", "契約情報一覧") <= 0 Or _
        DCount("*", "新規契約更新一覧") <= 0 Or _
        DCount("*", "解約情報一覧") <= 0 Or _
        DCount("*", "入居状況一覧") <= 0 Or _
        DCount("*", "契約情報管理") <= 0 Then
        log_write "set_contract:Contract 0件終了"
        Exit Sub
    End If
    
    
    '■契約情報データの設定====================================================================================K
    SQL = ""
    SQL = SQL & "INSERT INTO Contract SELECT "
    SQL = SQL & "入居状況一覧.ID, "
    SQL = SQL & """K"" & 入居状況一覧.ID as 識別ID, "
    SQL = SQL & "[物件No] & ""-"" & [部屋No] as legacy_property_id, "
    SQL = SQL & "iif(Nz([契約者2No],"""") = """",[契約者No],[契約者2No]) as legacy_resident_id, "
    SQL = SQL & "[物件No] as legacy_building_id, "
    SQL = SQL & "format(契約始期,""yyyy-mm-dd"") as start_from, "
    SQL = SQL & "format(契約終期,""yyyy-mm-dd"") as end_until, "
    SQL = SQL & "iif(入居状況一覧.契約者名 like ""%（株）%"" or " & _
                "入居状況一覧.契約者名 like ""%(株)%"" or " & _
                "入居状況一覧.契約者名 like ""%組合%"", " & _
                """20"",""10"") as kind_id, " '2025/11/5追加
    SQL = SQL & "[物件No] & ""-"" & [部屋No] & ""-"" & [契約者No] as legacy_id, "  '2025/11/12改修
    SQL = SQL & """1"" as enable_support_24h " '2025/11/5追加
    SQL = SQL & "FROM 入居状況一覧 "
    SQL = SQL & "WHERE 入居状況一覧.契約状況 IN (""契約中"", ""解約予定""); "
    db.Execute SQL
    log_write "set_contract:Insert_K"
    
    
    '■charge項目の更新
    set_charge
        
    
    '■legacy_owner_id,sublease更新
    SQL = ""
    SQL = SQL & "UPDATE Contract "
    SQL = SQL & "INNER JOIN 物件管理情報一覧 "
    SQL = SQL & "ON Contract.legacy_building_id = "
    SQL = SQL & "物件管理情報一覧.物件No "
    SQL = SQL & "SET Contract.legacy_owner_id = "
    SQL = SQL & "物件管理情報一覧.送金先1家主No, "
    SQL = SQL & "Contract.sublease = "
    SQL = SQL & "iif(物件管理情報一覧.管理形態 = ""管理委託"",""0"","""");"
    db.Execute SQL
    log_write "set_contract:legacy_owner_id,sublease"
    
    
    
    '■corporate_guarantor_nameの更新
    SQL = ""
    SQL = SQL & "UPDATE (Contract "
    SQL = SQL & "INNER JOIN 入居状況一覧 "
    SQL = SQL & "ON Contract.ID = 入居状況一覧.ID) "
    SQL = SQL & "INNER JOIN 契約情報一覧 "
    SQL = SQL & "ON (入居状況一覧.[物件No] = 契約情報一覧.[物件No]) "
    SQL = SQL & "AND (入居状況一覧.[部屋No] = 契約情報一覧.[部　屋]) "
    SQL = SQL & "AND (入居状況一覧.[契約者No] = 契約情報一覧.[契約者No]) "
    SQL = SQL & "SET Contract.corporate_guarantor_name = "
    SQL = SQL & "契約情報一覧.[家賃保証会社名(SJIS)];"
    db.Execute SQL
    log_write "set_contract:corporate_guarantor_name"
    
    
    '■end_dateの更新
    '契約者2Noに値がある場合
    SQL = ""
    SQL = SQL & "UPDATE (Contract "
    SQL = SQL & "INNER JOIN 入居状況一覧 "
    SQL = SQL & "ON Contract.ID = 入居状況一覧.ID) "
    SQL = SQL & "INNER JOIN 解約情報一覧 "
    SQL = SQL & "ON (入居状況一覧.[物件No] = 解約情報一覧.[物件No]) "
    SQL = SQL & "AND (入居状況一覧.[部屋No] = 解約情報一覧.[部　屋]) "
    SQL = SQL & "AND (入居状況一覧.[契約者2No] = 解約情報一覧.[契約者No]) "
    SQL = SQL & "SET Contract.end_date = "
    SQL = SQL & "解約情報一覧.[受付・立会情報 解約日] "
    SQL = SQL & "WHERE (Nz(入居状況一覧.[契約者2No],"""") <> """");"
    db.Execute SQL
    log_write "set_contract:end_date_1"
    
    '■legacy_contractor_idの更新
    SQL = ""
    SQL = SQL & "UPDATE Contract "
    SQL = SQL & "INNER JOIN 入居状況一覧 "
    SQL = SQL & "ON Contract.ID = 入居状況一覧.ID "
    SQL = SQL & "SET Contract.legacy_contractor_id = "
    SQL = SQL & "IIF(Nz(入居状況一覧.契約者2No,'') = '', NULL, 入居状況一覧.契約者No)"
    db.Execute SQL
    log_write "set_contract:legacy_contractor_id"
    
    
    '契約者2Noに値がない場合
    SQL = ""
    SQL = SQL & "UPDATE (Contract "
    SQL = SQL & "INNER JOIN 入居状況一覧 "
    SQL = SQL & "ON Contract.ID = 入居状況一覧.ID) "
    SQL = SQL & "INNER JOIN 解約情報一覧 "
    SQL = SQL & "ON (入居状況一覧.[物件No] = 解約情報一覧.[物件No]) "
    SQL = SQL & "AND (入居状況一覧.[部屋No] = 解約情報一覧.[部　屋]) "
    SQL = SQL & "AND (入居状況一覧.[契約者No] = 解約情報一覧.[契約者No]) "
    SQL = SQL & "SET Contract.end_date = "
    SQL = SQL & "解約情報一覧.[受付・立会情報 解約日] "
    SQL = SQL & "WHERE (Nz(入居状況一覧.[契約者2No],"""") = """");"
    db.Execute SQL
    log_write "set_contract:end_date_2"
       
   
    
    '■各種ライフラインカラム更新 2025/11/5改修
    If DCount("*", "ライフライン") > 0 Then  '2025/10/30条件追加
        SQL = ""
        SQL = SQL & "UPDATE (Contract "
        SQL = SQL & "INNER JOIN 入居状況一覧 "
        SQL = SQL & "ON Contract.ID = 入居状況一覧.ID) "
        SQL = SQL & "INNER JOIN ライフライン "
        SQL = SQL & "ON (入居状況一覧.[物件No] = ライフライン.物件No) "
        SQL = SQL & "AND (入居状況一覧.[部屋No] = ライフライン.部屋No) "
        SQL = SQL & "SET Contract.mail_box_no = ライフライン.メールボックス, "
        SQL = SQL & "Contract.drink_water_contact_name = ライフライン.飲料水会社名, "
        SQL = SQL & "Contract.drink_water_contact_tel = ライフライン.飲料水電話番号, "
        SQL = SQL & "Contract.drainage_contact_name = ライフライン.排水会社名, "
        SQL = SQL & "Contract.drainage_contact_tel = ライフライン.排水電話番号, "
        SQL = SQL & "Contract.gas_contact_name = ライフライン.ガス会社名, "
        SQL = SQL & "Contract.gas_contact_tel = ライフライン.ガス会社電話番号, "
        SQL = SQL & "Contract.internet_contact_name = ライフライン.インターネット会社名, "
        SQL = SQL & "Contract.internet_contact_tel = ライフライン.インターネット電話番号, "
        SQL = SQL & "Contract.cable_tv_contact_name = ライフライン.ケーブルテレビ会社名, "
        SQL = SQL & "Contract.cable_tv_contact_tel = ライフライン.ケーブルテレビ電話番号; "
        db.Execute SQL
        log_write "set_contract:lifeline_column"
    End If
      
    
    
    '■新規契約更新一覧データの設定===============================================================S
    SQL = ""
    SQL = SQL & "INSERT INTO Contract SELECT "
    SQL = SQL & "新規契約更新一覧.ID, "
    SQL = SQL & """S"" & 新規契約更新一覧.ID as 識別ID, "
    SQL = SQL & "新規契約更新一覧.[物件No] & ""-"" & 新規契約更新一覧.[部屋No] as legacy_property_id, "
    SQL = SQL & "IIf(新規契約更新一覧.[契約者2入居有無] = '入居有り', 新規契約更新一覧.[契約者2No], 新規契約更新一覧.[契約者1No]) as legacy_resident_id, "
    SQL = SQL & "新規契約更新一覧.[貸主No] as legacy_owner_id, "
    SQL = SQL & "新規契約更新一覧.[物件No] as legacy_building_id, "
    SQL = SQL & "format(新規契約更新一覧.契約始期,""yyyy-mm-dd"") as start_from, "
    SQL = SQL & "format(新規契約更新一覧.契約終期,""yyyy-mm-dd"") as end_until, "
    SQL = SQL & "新規契約更新一覧.家賃保証会社名 as [corporate_guarantor_name], "
    SQL = SQL & "iif(新規契約更新一覧.契約者名 like ""%（株）%"" or " & _
                "新規契約更新一覧.契約者名 like ""%(株)%"" or " & _
                "新規契約更新一覧.契約者名 like ""%組合%"", " & _
                """20"",""10"") as kind_id, " '2025/11/5追加
    SQL = SQL & "新規契約更新一覧.[契約者No_] as legacy_contractor_id, "
    SQL = SQL & "新規契約更新一覧.[物件No] & ""-"" & 新規契約更新一覧.[部屋No] & ""-"" & 新規契約更新一覧.[契約者No_] as legacy_id, "
    SQL = SQL & """1"" as enable_support_24h " '2025/11/5追加
    SQL = SQL & "FROM 新規契約更新一覧 "
    SQL = SQL & "LEFT JOIN 入居状況一覧 "
    SQL = SQL & "ON (新規契約更新一覧.[物件No] = 入居状況一覧.[物件No]) "
    SQL = SQL & "AND (新規契約更新一覧.[部屋No] = 入居状況一覧.[部屋No]) "
    SQL = SQL & "AND (DateValue(新規契約更新一覧.[契約始期]) = DateValue(入居状況一覧.[契約始期])) "
    SQL = SQL & "AND (新規契約更新一覧.[契約者No_] = 入居状況一覧.[契約者No]) "
    SQL = SQL & "Where 入居状況一覧.[ID] Is Null;"
    db.Execute SQL
    log_write "set_contract:Insert_S"
    
    
    
    '■legacy_owner_idの更新
'    SQL = ""
'    SQL = SQL & "UPDATE (Contract "
'    SQL = SQL & "INNER JOIN 新規契約更新一覧 "
'    SQL = SQL & "ON Contract.ID = 新規契約更新一覧.ID)"
'    SQL = SQL & "INNER JOIN 家主基本情報 "
'    SQL = SQL & "ON (新規契約更新一覧.貸主名 = 家主基本情報.家主名) "
'    SQL = SQL & "SET Contract.legacy_owner_id = "
'    SQL = SQL & "家主基本情報.[家主No] "
'    SQL = SQL & "Where Contract.[識別ID] like ""S%"";"
'    db.Execute SQL
'    log_write "set_contract:legacy_owner_id_S"
'
'
'    '■legacy_resident_id,legacy_idの更新
'    SQL = ""
'    SQL = SQL & "UPDATE (Contract "
'    SQL = SQL & "INNER JOIN 新規契約更新一覧 "
'    SQL = SQL & "ON Contract.ID = 新規契約更新一覧.ID)"
'    SQL = SQL & "INNER JOIN 契約者情報 "
'    SQL = SQL & "ON (新規契約更新一覧.契約者名 = 契約者情報.契約者名) "
'    SQL = SQL & "SET Contract.legacy_resident_id = "
'    SQL = SQL & "契約者情報.[契約者No], "
'    SQL = SQL & "Contract.legacy_id = "
'    SQL = SQL & "新規契約更新一覧.[物件No] & ""-"" & "
'    SQL = SQL & "新規契約更新一覧.[部屋No] & ""-"" & "
'    SQL = SQL & "契約者情報.[契約者No] "
'    SQL = SQL & "Where Contract.[識別ID] like ""S%"";"
'    db.Execute SQL
'    log_write "set_contract:legacy_resident_id_legacy_id_S"
    
    
    
    
    '■各種ライフラインカラム更新 2025/11/5追加
        If DCount("*", "ライフライン") > 0 Then
            SQL = ""
            SQL = SQL & "UPDATE (Contract "
            SQL = SQL & "INNER JOIN 新規契約更新一覧 "
            SQL = SQL & "ON Contract.ID = 新規契約更新一覧.ID) "
            SQL = SQL & "INNER JOIN ライフライン "
            SQL = SQL & "ON (新規契約更新一覧.[物件No] = ライフライン.物件No) "
            SQL = SQL & "AND (新規契約更新一覧.[部屋No] = ライフライン.部屋No) "
            SQL = SQL & "SET Contract.mail_box_no = ライフライン.メールボックス, "
            SQL = SQL & "Contract.drink_water_contact_name = ライフライン.飲料水会社名, "
            SQL = SQL & "Contract.drink_water_contact_tel = ライフライン.飲料水電話番号, "
            SQL = SQL & "Contract.drainage_contact_name = ライフライン.排水会社名, "
            SQL = SQL & "Contract.drainage_contact_tel = ライフライン.排水電話番号, "
            SQL = SQL & "Contract.gas_contact_name = ライフライン.ガス会社名, "
            SQL = SQL & "Contract.gas_contact_tel = ライフライン.ガス会社電話番号, "
            SQL = SQL & "Contract.internet_contact_name = ライフライン.インターネット会社名, "
            SQL = SQL & "Contract.internet_contact_tel = ライフライン.インターネット電話番号, "
            SQL = SQL & "Contract.cable_tv_contact_name = ライフライン.ケーブルテレビ会社名, "
            SQL = SQL & "Contract.cable_tv_contact_tel = ライフライン.ケーブルテレビ電話番号 "
            SQL = SQL & "WHERE Contract.[識別ID] like ""S%"";"
            db.Execute SQL
            log_write "set_contract:lifeline_column_S"
            
        End If
    
    
    '解約データ追加　2025/11/12追加
    '■契約情報データの設定====================================================================================CA
    AddNewFieldToTable "入居状況一覧", "マッチキー", "TEXT(255)"
    AddNewFieldToTable "解約情報一覧", "マッチキー", "TEXT(255)"
    log_write "set_contract:カラム追加_C"
    
    '■マッチキーリセット
    db.Execute "UPDATE 入居状況一覧 SET 入居状況一覧.マッチキー = '';"
    db.Execute "UPDATE 解約情報一覧 SET 解約情報一覧.マッチキー = '';"
    log_write "set_contract:マッチキーリセット_C"
    
    '■エラーフラグリセット
    db.Execute "UPDATE 契約情報管理 SET 契約情報管理.エラーフラグ = '0';"
    log_write "set_contract:エラーフラグリセット_C"
    
    '■契約終了ステータスの除外
    SQL = ""
    SQL = SQL & "UPDATE 契約情報管理 "
    SQL = SQL & "SET 契約情報管理.エラーフラグ = '1' "
    SQL = SQL & "WHERE 契約情報管理.契約更新状況 = '契約終了';"
    db.Execute SQL
    log_write "set_contract:契約終了ステータスの除外_C"
    
    
    '■マッチキーの作成
    '入居状況一覧
    SQL = ""
    SQL = SQL & "UPDATE 入居状況一覧 "
    SQL = SQL & "SET 入居状況一覧.[マッチキー] = "
    SQL = SQL & "入居状況一覧.[物件No] & ""-"" & 入居状況一覧.[部屋No] & ""-"" & 入居状況一覧.[契約者No];"
    db.Execute SQL
    log_write "set_contract:マッチキーの作成(入居状況一覧)_C"
    
    
    '解約情報一覧
    SQL = ""
    SQL = SQL & "UPDATE 解約情報一覧 "
    SQL = SQL & "SET 解約情報一覧.[マッチキー] = "
    SQL = SQL & "解約情報一覧.[物件No] & ""-"" & 解約情報一覧.[部　屋] & ""-"" & 解約情報一覧.[契約者No];"
    db.Execute SQL
    log_write "set_contract:マッチキーの作成(解約情報一覧)_C"
    
    
    '■入居状況一覧とのマッチング
    SQL = ""
    SQL = SQL & "UPDATE 契約情報管理 "
    SQL = SQL & "INNER JOIN 入居状況一覧 "
    SQL = SQL & "ON 契約情報管理.基幹契約ID = "
    SQL = SQL & "入居状況一覧.マッチキー "
    SQL = SQL & "SET 契約情報管理.エラーフラグ = ""1"" "
    SQL = SQL & "WHERE 契約情報管理.エラーフラグ = ""0"";"
    db.Execute SQL
    log_write "set_contract:入居状況一覧とのマッチング_C"
    
    
    
    '■契約情報管理のインサート
    SQL = ""
    SQL = SQL & "INSERT INTO Contract SELECT "
    SQL = SQL & "契約情報管理.ID, "
    SQL = SQL & """C"" & 契約情報管理.ID as 識別ID, "
    SQL = SQL & "契約情報管理.[基幹部屋ID] as legacy_property_id, "
    SQL = SQL & "契約情報管理.[基幹入居者ID] as legacy_resident_id, "
    SQL = SQL & "契約情報管理.[基幹オーナーID] as legacy_owner_id, "
    SQL = SQL & "契約情報管理.[基幹建物ID] as legacy_building_id, "
    SQL = SQL & "format(契約情報管理.賃貸開始,""yyyy-mm-dd"") as start_from, "
    SQL = SQL & "format(契約情報管理.賃貸終了,""yyyy-mm-dd"") as end_until, "
    SQL = SQL & "replace(replace(契約情報管理.賃料,""円"","""",1,-1,1),"","","""",1,-1,1) as charge_rent,"
    SQL = SQL & "replace(replace(契約情報管理.共益費,""円"","""",1,-1,1),"","","""",1,-1,1) as charge_management,"
    SQL = SQL & "replace(replace(契約情報管理.駐車場,""円"","""",1,-1,1),"","","""",1,-1,1) as charge_parking,"
    SQL = SQL & "replace(replace(契約情報管理.その他,""円"","""",1,-1,1),"","","""",1,-1,1) as charge_additional,"
    SQL = SQL & "IIF(契約情報管理.保証会社 = '-', NULL, 契約情報管理.保証会社) as corporate_guarantor_name,"
    SQL = SQL & "format(契約情報管理.入居日,""yyyy-mm-dd"") as property_checkup_from,"
    SQL = SQL & "format(解約情報一覧.[受付・立会情報 解約日],""yyyy-mm-dd"") as end_date,"
    SQL = SQL & "契約情報管理.[個人/法人] as kind_id,"
    SQL = SQL & "解約情報一覧.[契約者No] as legacy_contractor_id,"
    SQL = SQL & "契約情報管理.[基幹契約ID] as legacy_id,"
    SQL = SQL & "契約情報管理.[24時間サポート] as enable_support_24h,"
    SQL = SQL & "IIF(契約情報管理.[メールボックスNo] = '-', NULL, 契約情報管理.[メールボックスNo]) as mail_box_no,"
    SQL = SQL & "IIF(契約情報管理.[【電気】連絡先名] = '-', NULL, 契約情報管理.[【電気】連絡先名]) as electricity_contact_name,"
    SQL = SQL & "IIF(契約情報管理.[【電気】連絡先TEL] = '-', NULL, 契約情報管理.[【電気】連絡先TEL]) as electricity_contact_tel,"
    SQL = SQL & "IIF(契約情報管理.[【水道(飲料水)】連絡先名] = '-', NULL, 契約情報管理.[【水道(飲料水)】連絡先名]) as drink_water_contact_name,"
    SQL = SQL & "IIF(契約情報管理.[【水道(飲料水)】連絡先TEL] = '-', NULL, 契約情報管理.[【水道(飲料水)】連絡先TEL]) as drink_water_contact_tel,"
    SQL = SQL & "IIF(契約情報管理.[【水道(排水)】連絡先名] = '-', NULL, 契約情報管理.[【水道(排水)】連絡先名]) as drainage_contact_name,"
    SQL = SQL & "IIF(契約情報管理.[【水道(排水)】連絡先TEL] = '-', NULL, 契約情報管理.[【水道(排水)】連絡先TEL]) as drainage_contact_tel,"
    SQL = SQL & "IIF(契約情報管理.[【ガス】連絡先名] = '-', NULL, 契約情報管理.[【ガス】連絡先名]) as gas_contact_name,"
    SQL = SQL & "IIF(契約情報管理.[【ガス】連絡先TEL] = '-', NULL, 契約情報管理.[【ガス】連絡先TEL]) as gas_contact_tel,"
    SQL = SQL & "IIF(契約情報管理.[【インターネット】連絡先名] = '-', NULL, 契約情報管理.[【インターネット】連絡先名]) as internet_contact_name,"
    SQL = SQL & "IIF(契約情報管理.[【インターネット】連絡先TEL] = '-', NULL, 契約情報管理.[【インターネット】連絡先TEL]) as internet_contact_tel "
    SQL = SQL & "FROM 契約情報管理 "
    SQL = SQL & "INNER JOIN 解約情報一覧 "
    SQL = SQL & "ON 契約情報管理.[基幹契約ID] = 解約情報一覧.[マッチキー] "
    SQL = SQL & "Where 契約情報管理.[エラーフラグ] = ""0"";"
    db.Execute SQL
    log_write "set_contract:Insert_C"
    
    
    '■kind_idの設定(項目マッピング)
    SQL = ""
    SQL = SQL & "UPDATE Contract "
    SQL = SQL & "LEFT JOIN T_MAP_CONT_KIND_ID "
    SQL = SQL & "ON Contract.kind_id = "
    SQL = SQL & "T_MAP_CONT_KIND_ID.ORG "
    SQL = SQL & "SET Contract.kind_id = "
    SQL = SQL & "[T_MAP_CONT_KIND_ID]![NEW] "
    SQL = SQL & "Where Contract.[識別ID] like ""C%"";"
    db.Execute SQL
    log_write "set_contract:kind_idの設定(項目マッピング)_C"
    
    
    '■enable_support_24hの設定(項目マッピング)
    SQL = ""
    SQL = SQL & "UPDATE Contract "
    SQL = SQL & "LEFT JOIN T_MAP_CONT_SUPPORT24 "
    SQL = SQL & "ON Contract.enable_support_24h = "
    SQL = SQL & "T_MAP_CONT_SUPPORT24.ORG "
    SQL = SQL & "SET Contract.enable_support_24h = "
    SQL = SQL & "[T_MAP_CONT_SUPPORT24]![NEW] "
    SQL = SQL & "Where Contract.[識別ID] like ""C%"";"
    db.Execute SQL
    log_write "set_contract:enable_support_24hの設定(項目マッピング)_C"
    
    
    
    '■エラーチェック
    chk_required
    
    '進捗状況の更新~~~~~~~~~~~~~~~~~~
    frmName = "F_Main"
    
    DoEvents
    Call UpdateProgressBar(frmName)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    log_write "set_contract:out"
    

    Set db = Nothing
    
End Sub

Private Sub chk_required()

    log_write "chk_required:in"
    
    Dim t As Single
    t = Timer
    
    Dim db As ADODB.Connection
    Dim SQL As String
    
    Set db = CurrentProject.Connection
    
    
    '■legacy_property_idがブランク
    SQL = ""
    SQL = SQL & "UPDATE Contract "
    SQL = SQL & "SET Contract.[エラー内容] = "
    SQL = SQL & "[エラー内容] & ""/部屋Noがブランク"", "
    SQL = SQL & "Contract.[エラー項目] = "
    SQL = SQL & "[エラー項目] & ""/legacy_property_id"" "
    SQL = SQL & "WHERE (((Contract.legacy_building_id) Is Not Null "
    SQL = SQL & "And (Contract.legacy_building_id)<>"""") "
    SQL = SQL & "AND ((Contract.legacy_property_id) Is Null "
    SQL = SQL & "Or (Contract.legacy_property_id)=""""));"
    SQL = SQL & ""
    db.Execute SQL
    log_write "chk_required:legacy_property_id"
    
   
   
    '■legacy_resident_idがブランク
    SQL = ""
    SQL = SQL & "UPDATE Contract "
    SQL = SQL & "SET Contract.[エラー内容] = "
    SQL = SQL & "Contract.[エラー内容] & ""/契約者1No／契約者2Noどちらもブランク"", "
    SQL = SQL & "Contract.[エラー項目] = "
    SQL = SQL & "Contract.[エラー項目] & ""/legacy_resident_id"""
    SQL = SQL & "WHERE (((Contract.legacy_resident_id) Is Null "
    SQL = SQL & "Or (Contract.legacy_resident_id)=""""));"
    db.Execute SQL
    log_write "chk_required:legacy_resident_id"
    
    
    
     '■legacy_owner_idがブランク
    SQL = ""
    SQL = SQL & "UPDATE Contract "
    SQL = SQL & "SET Contract.[エラー内容] = "
    SQL = SQL & "Contract.[エラー内容] & ""/送金先1家主Noが取得できない"", "
    SQL = SQL & "Contract.[エラー項目] = "
    SQL = SQL & "Contract.[エラー項目] & ""/legacy_owner_id"""
    SQL = SQL & "WHERE (((Contract.legacy_owner_id) Is Null "
    SQL = SQL & "Or (Contract.legacy_owner_id)=""""));"
    db.Execute SQL
    log_write "chk_required:legacy_owner_id"
    
    
    
    '■legacy_building_idがブランク
    SQL = ""
    SQL = SQL & "UPDATE Contract "
    SQL = SQL & "SET Contract.[エラー内容] = "
    SQL = SQL & "Contract.[エラー内容] & ""/物件Noがブランク"", "
    SQL = SQL & "Contract.[エラー項目] = "
    SQL = SQL & "Contract.[エラー項目]  & ""/legacy_building_id"""
    SQL = SQL & "WHERE (((Contract.legacy_building_id) Is Null "
    SQL = SQL & "Or (Contract.legacy_building_id)=""""));"
    db.Execute SQL
    log_write "chk_required:legacy_building_id"
    
    
    
    '■start_fromがブランク
    SQL = ""
    SQL = SQL & "UPDATE Contract "
    SQL = SQL & "SET Contract.[エラー内容] = "
    SQL = SQL & "Contract.[エラー内容] & ""/契約始期がブランク"", "
    SQL = SQL & "Contract.[エラー項目] = "
    SQL = SQL & "Contract.[エラー項目] & ""/start_from"""
    SQL = SQL & "WHERE (((Contract.start_from) Is Null "
    SQL = SQL & "Or (Contract.start_from)=""""));"
    db.Execute SQL
    log_write "chk_required:start_from"
    
    
    
    '■end_untilがブランク
    SQL = ""
    SQL = SQL & "UPDATE Contract "
    SQL = SQL & "SET Contract.[エラー内容] = "
    SQL = SQL & "Contract.[エラー内容] & ""/契約終期がブランク"", "
    SQL = SQL & "Contract.[エラー項目] = "
    SQL = SQL & "Contract.[エラー項目]  & ""/end_until"""
    SQL = SQL & "WHERE (((Contract.end_until) Is Null "
    SQL = SQL & "Or (Contract.end_until)=""""));"
    db.Execute SQL
    log_write "chk_required:end_until"
    
    
    Set db = Nothing
     
    
    Debug.Print Timer - t
    
    log_write "chk_required:out"
    
End Sub

Sub out_contract_csv(strDir As String)
    log_write "out_contract_csv:in"
    Dim t As Single
    t = Timer
    
    Dim db As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim SQL As String
    
    Dim ST As ADODB.Stream
    Dim WS As ADODB.Stream
    Dim LINE As String
    Dim i As Long
    
    Set ST = New ADODB.Stream
    ST.Charset = "UTF-8"
    ST.LineSeparator = adCRLF
    ST.Open
    
    Set db = CurrentProject.Connection
    
    SQL = ""
    SQL = SQL & "SELECT * FROM Contract WHERE (((Contract.[エラー内容]) Is Null OR (Contract.[エラー内容])=''));"
    
    Set RS = db.Execute(SQL)
    log_write "out_contract_csv:Contract取得"
    
    LINE = ""
    For i = 4 To RS.Fields.Count - 1
        LINE = LINE & """" & RS.Fields(i).Name & ""","
    Next
    LINE = Left(LINE, Len(LINE) - 1)
    ST.WriteText LINE, stWriteLine
    log_write "out_contract_csv:Contractヘッダ出力"
   
    Do Until RS.EOF
    
        LINE = ""
        
        For i = 4 To RS.Fields.Count - 1
            If Nz(RS.Fields(i).Value, "") = "" Then
                LINE = LINE & ","
            Else
                LINE = LINE & """" & Replace(Nz(RS.Fields(i).Value, ""), """", """""", , , vbTextCompare) & ""","
            End If
        Next
        LINE = Left(LINE, Len(LINE) - 1)
        ST.WriteText LINE, stWriteLine
        
        
        RS.MoveNext
        DoEvents
    Loop
    log_write "out_contract_csv:Contractレコード出力"
    
    ST.SaveToFile strDir & "\contract.csv", adSaveCreateOverWrite
    log_write "out_contract_csv:Contract.csv保存"
    
    Debug.Print Timer - t
    log_write "out_contract_csv:out"
    
End Sub

Sub out_contract_xls(strDir As String)
    log_write "out_contract_xls:in"
    Dim t As Single
    t = Timer
    
    Dim FS As Object
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    If FS.FileExists(strDir & "\contract.xlsx") = True Then
        FS.DeleteFile strDir & "\contract.xlsx"
    End If
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "Q_Contract", strDir & "\contract.xlsx", True
    log_write "out_contract_csv:Contract.xlsx保存"
        
    Debug.Print Timer - t
    log_write "out_contract_xls:out"
    
End Sub

Sub set_charge()
            
    
    log_write "set_charge:in"
    
    Dim db As ADODB.Connection
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim SQL As String
    Dim SQL1 As String
    Dim SQL2 As String
    Dim vOLD As String
    Dim total As Long
    Dim firstRs1 As Boolean
    Dim fld As ADODB.Field
    
    Set db = CurrentProject.Connection
    
    Set rs1 = New ADODB.Recordset
    SQL1 = ""
    SQL1 = "SELECT * FROM 入居状況一覧  order by ID"
    rs1.Open SQL1, db, adOpenKeyset, adLockOptimistic
    
    
    Set rs2 = New ADODB.Recordset
    SQL2 = ""
    SQL2 = "SELECT * FROM T_MAP_CONTRACT WHERE (((出力項目) Is Not Null And (出力項目)<>"""")) order by 出力項目;"
    rs2.Open SQL2, db, adOpenKeyset, adLockOptimistic
    
    If rs2.BOF And rs2.EOF Then
        log_write "データが0件のため、set_charge処理スキップ"
    Else
    
        vOLD = ""
        total = 0
        rs1.MoveFirst
        Do Until rs1.EOF
            vOLD = ""
            firstRs1 = True  ' rs1のレコードが変わった直後のフラグ
            total = 0
            
            rs2.MoveFirst
            Do Until rs2.EOF
                If Not firstRs1 And rs2!出力項目 <> vOLD Then
                    
                    SQL = ""
                    SQL = SQL & "UPDATE Contract "
                    SQL = SQL & "SET Contract.[" & vOLD & "] = " & total
                    SQL = SQL & " WHERE (((Contract.ID)=" & rs1!Id & "));"
                    db.Execute SQL
                    
                    total = 0
                    
                End If
                
                If firstRs1 Then
                    vOLD = rs2!出力項目  ' rs1のレコードが変わった後、最初のrs2でvOLDを設定
                    firstRs1 = False
                Else
                    vOLD = rs2!出力項目
                End If
            
                
                For Each fld In rs1.Fields
                    If fld.Name = rs2!項目略名 Then    '項目略名と完全一致のカラム
                        If Nz(rs1.Fields(fld.Name).Value, "") <> "" Then
                            total = total + rs1.Fields(fld.Name).Value
                        End If
                    End If
                Next
                
                rs2.MoveNext
            Loop
            
            If vOLD <> "" Then   '最後の金額を更新する
                SQL = ""
                SQL = SQL & "UPDATE Contract "
                SQL = SQL & "SET Contract.[" & vOLD & "] = " & total
                SQL = SQL & " WHERE Contract.ID = " & rs1!Id & ";"
                db.Execute SQL
                
            End If
            
            rs1.MoveNext
        
        Loop
    End If

    rs1.Close
    rs2.Close
    Set rs1 = Nothing
    Set rs2 = Nothing
    Set db = Nothing

    log_write "set_charge:out"
    
End Sub
