Attribute VB_Name = "M_Dlaw"
Option Explicit

' Tasksシートの列インデックス
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME = 2
Private Const COL_DURATION As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_END_DATE As Long = 5
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

' ID入力
Public Sub DrawLine()
    ' イベントの発生を一時的に無効にする
    Application.EnableEvents = False
    ' エラーが発生しても必ずイベントを有効に戻すためのエラーハンドリング
    On Error GoTo ErrHandler
    
    ' 変数宣言
    Dim i As Integer
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim lastRaw As Integer
    Dim dt As Date
    ' 代入
    Set ws = ThisWorkbook.Worksheets("Tasks")
    lastRaw = ws.Cells(ws.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    Set dataRange = ws.Range( _
        ws.Cells(1, COL_TASK_ID), _
        ws.Cells(lastRaw, COL_STATUS) _
    )
    ws.Cells.Borders.LineStyle = xlNone
    ' 入力をリセット
    ' 入力---------------------------------------
    For i = 2 To lastRaw
        ws.Cells(i, COL_TASK_ID).Value = i - 1
        dt = ws.Cells(i, COL_START_DATE).Value
        dt = dt + ws.Cells(i, COL_DURATION).Value - 1
        ws.Cells(i, COL_END_DATE).NumberFormat = "yyyy/mm/dd"
        ws.Cells(i, COL_END_DATE).Value = dt
    Next i
    '--------------------------------------------
    ' 見出し行に罫線を引く
    ws.Range(ws.Cells(1, COL_TASK_ID), ws.Cells(1, COL_STATUS)).Borders.LineStyle = xlContinuous
    ' フォーマット整える
    With dataRange
        ' 幅自動調整
        .EntireColumn.AutoFit
        ' 罫線を引く
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    
ExitProc:     ' 正常終了、またはエラー処理後のジャンプ先
    ' 開放
    Set ws = Nothing
    Set dataRange = Nothing
    ' イベントの発生を再度有効にする
    Application.EnableEvents = True
    Exit Sub

ErrHandler: ' エラー発生時のジャンプ先
    MsgBox "DrawLineプロシージャでエラーが発生しました: " & Err.Description, vbCritical
    ' エラーが発生しても必ずイベントを有効に戻すため、ExitProcにジャンプ
    Resume ExitProc
End Sub
