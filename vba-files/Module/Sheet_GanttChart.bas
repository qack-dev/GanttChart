Attribute VB_Name = "Sheet_GanttChart"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// Sheet_GanttChart シートモジュール
'// 「GanttChart」シートに固有のイベント処理を記述します。
'////////////////////////////////////////////////////////////////////////////////////////////////////

'--- Tasksシートの列インデックス ---
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME As Long = 2
Private Const COL_DURATION As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_END_DATE As Long = 5
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

'--- GanttChartシートの固定行・列 ---
Private Const GANTT_START_ROW As Long = 5 ' ガントチャートの開始行
Private Const GANTT_START_COL As Long = 2 ' ガントチャートの開始列 (タスク名表示エリア)
Private Const TIMELINE_ROW As Long = 4    ' タイムラインの表示行

'/**
' * @brief "更新"ボタンがクリックされたときに呼び出されます。
' */
Private Sub UpdateChartButton_Click()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Call M_GanttChart.UpdateGanttChart
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "チャート更新ボタンの処理中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

'/**
' * @brief ワークシートの選択範囲が変更されたときに発生するイベントです。
' *        ガントチャートのタスクバー（着色されたセル）が選択された場合、
' *        そのタスクの詳細情報をMsgBoxで表示します。
' * @param Target 選択されたセル範囲
' */
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo ErrHandler

    Dim selectedRow As Long
    Dim selectedCol As Long
    Dim taskRow As Long
    Dim wsTasks As Worksheet
    Dim msg As String

    ' --- 選択されたセルが単一セルでない場合は処理を抜ける ---
    If Target.Cells.CountLarge > 1 Then Exit Sub

    selectedRow = Target.Row
    selectedCol = Target.Column

    ' --- 選択されたセルがガントチャートのタスク描画エリア内か判定 ---
    If selectedRow >= GANTT_START_ROW And selectedCol > GANTT_START_COL Then
        
        ' --- 選択されたセルに着色があるか（タスクバーか）判定 ---
        If Target.Interior.Color <> xlNone Then
            
            ' --- 対応するタスク情報を取得 ---
            taskRow = selectedRow - GANTT_START_ROW + 2 ' Tasksシートの行番号に変換
            Set wsTasks = ThisWorkbook.Sheets("Tasks")
            
            ' --- タスク情報が存在するか確認 ---
            If wsTasks.Cells(taskRow, COL_TASK_NAME).Value <> "" Then
                
                ' --- メッセージボックスで詳細情報を表示 ---
                msg = "■ タスク詳細" & vbCrLf & vbCrLf & _
                      "タスク名: " & wsTasks.Cells(taskRow, COL_TASK_NAME).Value & vbCrLf & _
                      "担当者: " & "(未実装)" & vbCrLf & _
                      "期間: " & Format(wsTasks.Cells(taskRow, COL_START_DATE).Value, "yyyy/m/d") & " - " & Format(wsTasks.Cells(taskRow, COL_END_DATE).Value, "yyyy/m/d") & " (" & wsTasks.Cells(taskRow, COL_DURATION).Value & "日間)" & vbCrLf & _
                      "進捗率: " & Format(wsTasks.Cells(taskRow, COL_PROGRESS).Value, "0%") & vbCrLf & _
                      "ステータス: " & wsTasks.Cells(taskRow, COL_STATUS).Value
                      
                MsgBox msg, vbInformation, "タスク詳細"
            End If
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox "SelectionChangeイベントの処理中にエラーが発生しました: " & Err.Description, vbCritical
End Sub
