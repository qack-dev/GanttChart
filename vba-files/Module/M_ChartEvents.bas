Attribute VB_Name = "M_ChartEvents"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// M_ChartEvents モジュール
'// ガントチャート上のタスクバー（Shapeオブジェクト）クリックイベントを処理します。
'////////////////////////////////////////////////////////////////////////////////////////////////////

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
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, R1C2).End(xlUp).Row
    taskFound = False

    ' Tasksシートから該当タスクの情報を検索
    For i = 2 To lastTaskRow
        If wsTasks.Cells(R1C1, i).Value = taskID Then
            msg = "タスクID: " & wsTasks.Cells(R1C1, i).Value & vbCrLf & _
                  "タスク名: " & wsTasks.Cells(R1C2, i).Value & vbCrLf & _
                  "期間: " & wsTasks.Cells(R1C3, i).Value & "日" & vbCrLf & _
                  "開始日: " & Format(wsTasks.Cells(R1C4, i).Value, "yyyy/mm/dd") & vbCrLf & _
                  "終了日: " & Format(wsTasks.Cells(R1C5, i).Value, "yyyy/mm/dd") & vbCrLf & _
                  "進捗: " & Format(wsTasks.Cells(R1C6, i).Value, "0%") & vbCrLf & _
                  "ステータス: " & wsTasks.Cells(R1C7, i).Value
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