VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// Sheet_Tasks シートモジュール
'// 「Tasks」シートのイベント（例: Worksheet_Change）を処理し、データ変更時にガントチャートの更新をトリガーします。
'////////////////////////////////////////////////////////////////////////////////////////////////////

' Tasksシートの列インデックス
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME As Long = 2
Private Const COL_DURATION As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_END_DATE As Long = 5
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim dataRange As Range
    Set ws = ThisWorkbook.Worksheets("Tasks")
    Set dataRange = ws.Range(ws.Cells(1, COL_TASK_NAME), ws.Cells(ws.Rows.Count, COL_STATUS)) ' B列からG列までを監視対象とする

    ' 変更されたセルが監視対象範囲内にあるか、かつ複数セルが変更されていないかを確認
    If Not Intersect(Target, dataRange) Is Nothing And Target.Cells.Count = 1 Then
        ' 変更された行がデータ行（ヘッダー行以外）であるかを確認
        If Target.Row >= 2 Then
            ' 終了日 (E列) が変更された場合は、期間 (C列) を再計算
            If Target.Column = COL_END_DATE Then ' E列 (終了日)
                ws.Cells(Target.Row, COL_DURATION).Value = ws.Cells(Target.Row, COL_END_DATE).Value - ws.Cells(Target.Row, COL_START_DATE).Value + 1
            End If

            ' 期間 (C列) または開始日 (D列) が変更された場合は、終了日 (E列) を再計算
            If Target.Column = COL_DURATION Or Target.Column = COL_START_DATE Then ' C列 (期間) または D列 (開始日)
                ws.Cells(Target.Row, COL_END_DATE).Value = ws.Cells(Target.Row, COL_START_DATE).Value + ws.Cells(Target.Row, COL_DURATION).Value - 1
            End If

            ' ガントチャートを更新
            Call M_GanttChart.UpdateGanttChart
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox "Tasksシートの変更イベント中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

