Attribute VB_Name = "新規契約_S"
Option Compare Database
Option Explicit

Sub set_contract()
    log_write "set_contract:in"

    Dim t As Single
    t = Timer

    Dim db As DAO.Database
    Dim SQL As String

    Set db = CurrentDb

    db.Execute "DELETE FROM Contract WHERE ID LIKE 'S_%';"

    If DCount("ID", "Contract", "ID LIKE 'S_*'") = 0 Then

    '================================================================================
    '新規契約_S
    '================================================================================

    '出力条件
    AddNewFieldToTable "新規契約更新一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 新規契約更新一覧 SET 新規契約更新一覧.[FLG] = '1';"

    ' --- AI GENERATED (sk-architect) ---
    'Blank check — 部屋No
        SQL = ""
        SQL = SQL & "UPDATE 新規契約更新一覧 AS T "
        SQL = SQL & "SET T.FLG = '1' "
        SQL = SQL & "WHERE Nz(T.[部屋No], '') = ''; "
        db.Execute SQL
    ' --- END AI GENERATED ---

    'JOIN filter → 紐づかないデータを抽出
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T1 "
    SQL = SQL & "LEFT JOIN 入居状況一覧 AS T2 "
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) AND (T1.[部屋No] = T2.[部屋No]) AND (T1.[契約者1No] = T2.[契約者1No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[物件No] IS NULL; "
    db.Execute SQL

    log_write "set_contract:新規契約_S → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Contract SELECT "
    SQL = SQL & "    'S_' & T.ID as ID,"
    SQL = SQL & "    T.[物件No] & ""-"" & T.[部屋No] as legacy_property_id,"
    SQL = SQL & "    T.[家主No] as legacy_owner_id,"
    SQL = SQL & "    T.[物件No] as legacy_building_id,"
    SQL = SQL & "    T.[契約始期] as start_from,"
    SQL = SQL & "    T.[契約終期] as end_until,"
    SQL = SQL & "    '0' as charge_rent,"
    SQL = SQL & "    T.[家賃保証会社名] as corporate_guarantor_name,"
    SQL = SQL & "    T.[物件No] & ""-"" & T.[部屋No] & ""-"" & T.[契約者1No] as legacy_id,"
    SQL = SQL & "    '新規契約_S' as sheet"
    SQL = SQL & "FROM 新規契約更新一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_contract:新規契約_S → 通常処理"

    '特殊処理

    ' Dim con — điều kiện ưu tiên 契約者1→3→2 (khai báo 1 lần)
    Dim con3 As String
    con3 = "(T1.[契約者3入居有無] = '入居有り')"
    Dim con2 As String
    con2 = "(T1.[契約者3入居有無] <> '入居有り' AND T1.[契約者2入居有無] = '入居有り')"

    SQL = ""
    SQL = SQL & "UPDATE Contract AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('S_' & T1.ID) "
    SQL = SQL & "SET T.[legacy_resident_id] = SWITCH( "
    SQL = SQL & "    T1.[個人・法人区分_16] = '個人' AND T1.[個人・法人区分_19] = '個人', T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ); "
    db.Execute SQL
    log_write "set_contract:新規契約_S → 特殊処理: legacy_resident_id"


    '■項目削除
    DeleteFieldInTable "新規契約更新一覧", "FLG"

    End If

    chk_required

    Debug.Print "Contract:" & Timer - t
    log_write "set_contract:out"

End Sub
