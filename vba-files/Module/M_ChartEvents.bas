Attribute VB_Name = "M_ChartEvents"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// M_ChartEvents モジュール
'// ガントチャート上のタスクバー（Shapeオブジェクト）クリックイベントを処理します。
'////////////////////////////////////////////////////////////////////////////////////////////////////

' Tasksシートの列インデックス
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME As Long = 2
Private Const COL_DURATION As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_END_DATE As Long = 5
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

' クリックされたタスクの詳細を表示する
Public Sub ShowTaskDetails()
    On Error GoTo ErrHandler

    Dim clickedShape As Shape
    Dim taskID As Long
    Dim wsTasks As Worksheet
    Dim lastTaskRow As Long
    Dim i As Long
    Dim taskFound As Boolean
    Dim msg As String

    ' クリックされたShapeオブジェクトを取得
    Set clickedShape = ActiveSheet.Shapes(Application.Caller)

    ' Shapeの名前からタスクIDを抽出
    If Left(clickedShape.Name, 8) = "TaskBar_" Then
        taskID = CLng(Mid(clickedShape.Name, 9))
    Else
        Exit Sub ' タスクバー以外のShapeがクリックされた場合は何もしない
    End If

    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    taskFound = False

    ' Tasksシートから該当タスクの情報を検索
    For i = 2 To lastTaskRow
        If wsTasks.Cells(i, COL_TASK_ID).value = taskID Then
            msg = "タスクID: " & wsTasks.Cells(i, COL_TASK_ID).value & vbCrLf & _
                  "タスク名: " & wsTasks.Cells(i, COL_TASK_NAME).value & vbCrLf & _
                  "期間: " & wsTasks.Cells(i, COL_DURATION).value & "日" & vbCrLf & _
                  "開始日: " & Format(wsTasks.Cells(i, COL_START_DATE).value, "yyyy/mm/dd") & vbCrLf & _
                  "終了日: " & Format(wsTasks.Cells(i, COL_END_DATE).value, "yyyy/mm/dd") & vbCrLf & _
                  "進捗: " & Format(wsTasks.Cells(i, COL_PROGRESS).value, "0%") & vbCrLf & _
                  "ステータス: " & wsTasks.Cells(i, COL_STATUS).value
            taskFound = True
            Exit For
        End If
    Next i

    If taskFound Then
        MsgBox msg, vbInformation, "タスク詳細"
    Else
        MsgBox "タスク情報が見つかりませんでした。", vbExclamation
    End If

    Exit Sub

ErrHandler:
    MsgBox "タスク詳細の表示中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

