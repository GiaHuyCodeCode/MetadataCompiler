Attribute VB_Name = "未入金"
Option Compare Database
Option Explicit

Sub set_data()
    log_write "set_data:in"

    Dim t As Single
    t = Timer

    Dim db As ADODB.Connection
    Dim SQL As String

    Set db = CurrentProject.Connection

    db.Execute "DELETE FROM TargetTable WHERE ID LIKE 'X_%';"

    If DCount("ID", "TargetTable", "ID LIKE 'X_*'") = 0 Then

    '================================================================================
    '未入金
    '================================================================================

    '出力条件
    AddNewFieldToTable "SourceTable", "FLG", "TEXT(1)"

    '■FLGリセット
    db.Execute "UPDATE SourceTable SET SourceTable.[FLG] = '0';"

    ' TODO:AI_REVIEW — AI fallback failed: 429 You exceeded your current quota, please check your plan and billing details. For more information on this error, hea
    ' Context: [{"使用ファイル": "ファイル名"}]

    log_write "set_data:未入金 → 不要行を削除するため(FLG=1更新)"


    '■項目削除
    DeleteFieldInTable "SourceTable", "FLG"

    End If

    chk_required

    Debug.Print "TargetTable:" & Timer - t
    log_write "set_data:out"

End Sub
