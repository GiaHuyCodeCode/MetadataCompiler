Sub set_account()

    log_write "set_account:in"

    Dim db As ADODB.Connection
    Dim SQL As String
    Dim t As Single
    t = Timer

    Set db = CurrentProject.Connection

'■対象テーブルを全削除
    db.Execute "DELETE FROM Account;"
    log_write "account delete"

'■データ存在チェック
    If DCount("*", "新規契約更新一覧") = 0 Then
        log_write "set_account: 新規契約更新一覧 is empty → Exit"
        GoTo PROC_END
    End If

'================================================================================
'【GMO用】新規契約更新一覧 × 入居状況一覧（新規入居者を抽出）
'================================================================================

'出力条件
    AddNewFieldToTable "新規契約更新一覧", "FLG", "TEXT(1)"

'■FLGリセット
    db.Execute "UPDATE 新規契約更新一覧 SET 新規契約更新一覧.[FLG] = '0';"

'■入居状況一覧に紐づくレコードを除外（未来の契約のみ残す）
'  JOIN_FILTER: 物件No + 部屋No + 契約者1No が一致するものをFLG=1
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T1 "
    SQL = SQL & "INNER JOIN 入居状況一覧 AS T2 "
    SQL = SQL & "    ON T1.[物件No] = T2.[物件No] "
    SQL = SQL & "    AND T1.[部屋No] = T2.[部屋No] "
    SQL = SQL & "    AND T1.[契約者1No] = T2.[契約者1No] "
    SQL = SQL & "SET T1.FLG = '1'; "
    db.Execute SQL

    log_write "set_account:新規契約更新一覧 → 不要行を削除するため(FLG=1更新)"

'通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Account SELECT "
    SQL = SQL & "    'N_' & T.ID as ID, "
    SQL = SQL & "    'Resident' as klass, "
    SQL = SQL & "    'gmo002' as legacy_charge_user_id, "
    SQL = SQL & "    'アカウント（新規入居者）_N' as sheet "
    SQL = SQL & "FROM 新規契約更新一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0';"
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 通常処理"

'■FLG削除（出力条件用）
'  ※ DEDUPはINSERT後に実施するためFLG削除は後述

'================================================================================
'特殊処理 — 共通条件 Dim con（全UPDATE_SWITCHブロックで再利用）
'================================================================================
'  con1: 個人・法人区分_16 AND _19 がともに「個人」→ 契約者1No_17 (結果①)
'  con3: 上記以外 AND 契約者3入居有無=入居有り → 契約者3No (結果②)
'  con2: 上記以外 AND 契約者2入居有無=入居有り → 契約者2No (結果③)
'  True(fallback): 契約者1No_17 (結果④)
    Dim con1 As String
    Dim con3 As String
    Dim con2 As String
    con1 = "(T1.[個人・法人区分_16] = '個人' AND T1.[個人・法人区分_19] = '個人')"
    con3 = "(NOT (T1.[個人・法人区分_16] = '個人' AND T1.[個人・法人区分_19] = '個人') AND T1.[契約者3入居有無] = '入居有り')"
    con2 = "(NOT (T1.[個人・法人区分_16] = '個人' AND T1.[個人・法人区分_19] = '個人') AND T1.[契約者3入居有無] <> '入居有り' AND T1.[契約者2入居有無] = '入居有り')"

'================================================================================
'特殊処理 — legacy_id
'  SWITCH(契約者優先順) → 契約者No を legacy_id にセット
'================================================================================
    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID) "
    SQL = SQL & "SET T.legacy_id = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ); "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: legacy_id"

'================================================================================
'出力条件 — DEDUP（legacy_idをキーに重複削除、INSERT後に実施）
'  データ順で最初の1件のみ採用
'================================================================================
    SQL = ""
    SQL = SQL & "DELETE FROM Account "
    SQL = SQL & "WHERE Account.ID NOT IN ( "
    SQL = SQL & "    SELECT MIN(ID) "
    SQL = SQL & "    FROM Account "
    SQL = SQL & "    WHERE Account.ID LIKE 'N_%' "
    SQL = SQL & "    GROUP BY legacy_id "
    SQL = SQL & ")  AND Account.ID LIKE 'N_%';"
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → legacy_idキーで重複削除(DEDUP)"

'■FLG削除（新規契約更新一覧）
    DeleteFieldInTable "新規契約更新一覧", "FLG"

'================================================================================
'特殊処理 — email / tel_mobile
'  ルート①: 入居者管理.基幹入居者ID に紐づく場合（HYPHEN_TO_NULL適用）
'================================================================================

'--- email (入居者管理経由) ---
    SQL = ""
    SQL = SQL & "UPDATE (Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
    SQL = SQL & "INNER JOIN 入居者管理 AS T2 ON T2.[基幹入居者ID] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ) "
    SQL = SQL & "SET T.email = IIF(T2.[メールアドレス] = '-', NULL, T2.[メールアドレス]); "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: email (入居者管理経由)"

'--- tel_mobile (入居者管理経由) ---
    SQL = ""
    SQL = SQL & "UPDATE (Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
    SQL = SQL & "INNER JOIN 入居者管理 AS T2 ON T2.[基幹入居者ID] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ) "
    SQL = SQL & "SET T.tel_mobile = IIF(T2.[携帯番号/SMS] = '-', NULL, T2.[携帯番号/SMS]); "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: tel_mobile (入居者管理経由)"

'  ルート②: 入居者管理に紐づかない場合 → 契約者情報経由（NULLのレコードのみ更新）

'--- email (契約者情報経由) ---
    SQL = ""
    SQL = SQL & "UPDATE (Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
    SQL = SQL & "INNER JOIN 契約者情報 AS T2 ON T2.[契約者No] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ) "
    SQL = SQL & "SET T.email = T2.[優先メール] "
    SQL = SQL & "WHERE Nz(T.email, '') = ''; "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: email (契約者情報経由)"

'--- tel_mobile (契約者情報経由) ---
    SQL = ""
    SQL = SQL & "UPDATE (Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
    SQL = SQL & "INNER JOIN 契約者情報 AS T2 ON T2.[契約者No] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ) "
    SQL = SQL & "SET T.tel_mobile = T2.[携帯1] "
    SQL = SQL & "WHERE Nz(T.tel_mobile, '') = ''; "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: tel_mobile (契約者情報経由)"

'================================================================================
'特殊処理 — name_family / name_family_kana / company_name / company_name_kana / kind_id
'  契約者情報.契約者No に 契約者優先順(SWITCH) で紐づけてまとめて出力
'================================================================================

'--- name_family (契約者名) ---
    SQL = ""
    SQL = SQL & "UPDATE (Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
    SQL = SQL & "INNER JOIN 契約者情報 AS T2 ON T2.[契約者No] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ) "
    SQL = SQL & "SET T.name_family = T2.[契約者名]; "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: name_family"

'--- name_family_kana (契約者名カナ) ---
    SQL = ""
    SQL = SQL & "UPDATE (Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
    SQL = SQL & "INNER JOIN 契約者情報 AS T2 ON T2.[契約者No] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ) "
    SQL = SQL & "SET T.name_family_kana = T2.[契約者名カナ]; "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: name_family_kana"

'--- company_name_kana (契約者名 → company_name_kana) ---
    SQL = ""
    SQL = SQL & "UPDATE (Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
    SQL = SQL & "INNER JOIN 契約者情報 AS T2 ON T2.[契約者No] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ) "
    SQL = SQL & "SET T.company_name_kana = T2.[契約者名]; "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: company_name_kana"

'--- company_name (契約者名カナ → company_name) ---
    SQL = ""
    SQL = SQL & "UPDATE (Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
    SQL = SQL & "INNER JOIN 契約者情報 AS T2 ON T2.[契約者No] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ) "
    SQL = SQL & "SET T.company_name = T2.[契約者名カナ]; "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: company_name"

'--- kind_id (契約者区分 → 一時セット、後でコード変換) ---
    SQL = ""
    SQL = SQL & "UPDATE (Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
    SQL = SQL & "INNER JOIN 契約者情報 AS T2 ON T2.[契約者No] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ) "
    SQL = SQL & "SET T.kind_id = T2.[契約者区分]; "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: kind_id (仮セット)"

'================================================================================
'特殊処理 — 個人/法人区分によるフィールドクリア
'  個人の場合: company_name / company_name_kana を NULL
'  法人の場合: name_family / name_family_kana を NULL
'================================================================================

'--- 個人の場合 → company_name, company_name_kana をクリア ---
    SQL = ""
    SQL = SQL & "UPDATE (Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
    SQL = SQL & "INNER JOIN 契約者情報 AS T2 ON T2.[契約者No] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ) "
    SQL = SQL & "SET T.company_name = NULL, "
    SQL = SQL & "    T.company_name_kana = NULL "
    SQL = SQL & "WHERE T2.[契約者区分] = '個人'; "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: 個人の場合 company_name/company_name_kana クリア"

'--- 法人の場合 → name_family, name_family_kana をクリア ---
    SQL = ""
    SQL = SQL & "UPDATE (Account AS T "
    SQL = SQL & "INNER JOIN 新規契約更新一覧 AS T1 ON T.ID = ('N_' & T1.ID)) "
    SQL = SQL & "INNER JOIN 契約者情報 AS T2 ON T2.[契約者No] = SWITCH( "
    SQL = SQL & "    " & con1 & ", T1.[契約者1No_17], "
    SQL = SQL & "    " & con3 & ", T1.[契約者3No], "
    SQL = SQL & "    " & con2 & ", T1.[契約者2No], "
    SQL = SQL & "    True, T1.[契約者1No_17] ) "
    SQL = SQL & "SET T.name_family = NULL, "
    SQL = SQL & "    T.name_family_kana = NULL "
    SQL = SQL & "WHERE T2.[契約者区分] = '法人'; "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: 法人の場合 name_family/name_family_kana クリア"

'================================================================================
'特殊処理 — kind_id コード変換
'  '個人' → '10' / '法人' → '20'
'================================================================================
    SQL = ""
    SQL = SQL & "UPDATE Account AS T "
    SQL = SQL & "SET T.kind_id = SWITCH( "
    SQL = SQL & "    T.kind_id = '個人', '10', "
    SQL = SQL & "    T.kind_id = '法人', '20', "
    SQL = SQL & "    True, NULL ) "
    SQL = SQL & "WHERE T.ID LIKE 'N_%'; "
    db.Execute SQL
    log_write "set_account:新規契約更新一覧 → 特殊処理: kind_id コード変換(個人→10/法人→20)"

PROC_END:
    chk_required

    Debug.Print "account:" & Timer - t
    log_write "set_account:out"

End Sub