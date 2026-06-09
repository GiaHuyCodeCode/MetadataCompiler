Attribute VB_Name = "部屋（新規）"
Option Compare Database
Option Explicit

Sub set_property()
    log_write "set_property:in"

    Dim t As Single
    t = Timer

    Dim db As ADODB.Connection
    Dim SQL As String

    Set db = CurrentProject.Connection

    db.Execute "DELETE FROM Property WHERE ID LIKE 'BN_%';"

    If DCount("ID", "Property", "ID LIKE 'BN_*'") = 0 Then

    '================================================================================
    '部屋（新規）
    '================================================================================

    '出力条件
    AddNewFieldToTable "新規契約更新一覧", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE 新規契約更新一覧 SET 新規契約更新一覧.[FLG] = '1';"

    'JOIN filter → 紐づかないデータを抽出
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T1 "
    SQL = SQL & "LEFT JOIN 入居状況一覧 AS T2 "
    SQL = SQL & "ON (T1.[物件No] = T2.[物件No]) AND (T1.[部屋No] = T2.[部屋No]) "
    SQL = SQL & "SET T1.FLG = '0' "
    SQL = SQL & "WHERE T2.[物件No] IS NULL; "
    db.Execute SQL

    ' 出力 (No filter required)
    '■有効レコードのみFLG=0に戻す
    SQL = ""
    SQL = SQL & "UPDATE 新規契約更新一覧 AS T "
    SQL = SQL & "SET T.FLG = '0' "
    SQL = SQL & "WHERE T.[貸主No] = 'ブランク' OR T.[貸主No] = '0'; "
    db.Execute SQL

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"該当ファイル名": "新規契約更新一覧.csv", "該当項目名_1": "部屋No", "処理": "条件分岐", "条件": "\"駐車場\",\"看板\"を含む", "開発用チェック": "True"}]

    ' --- AI GENERATED (sk-architect) ---
    '出力なし (No action required)
    ' --- END AI GENERATED ---

    log_write "set_property:部屋（新規） → 不要行を削除するため(FLG=1更新)"

    '通常処理
    SQL = ""
    SQL = SQL & "INSERT INTO Property SELECT "
    SQL = SQL & "    'BN_' & T.ID as ID,"
    SQL = SQL & "    T.[部屋No] as name,"
    SQL = SQL & "    '部屋（新規）' as sheet"
    SQL = SQL & "FROM 新規契約更新一覧 AS T "
    SQL = SQL & "WHERE T.FLG = '0'; "
    db.Execute SQL
    log_write "set_property:部屋（新規） → 通常処理"


    '■項目削除
    DeleteFieldInTable "新規契約更新一覧", "FLG"

    End If

    chk_required

    Debug.Print "Property:" & Timer - t
    log_write "set_property:out"

End Sub
