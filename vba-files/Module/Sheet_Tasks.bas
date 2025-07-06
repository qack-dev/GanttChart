Attribute VB_Name = "Sheet_Tasks"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// Sheet_Tasks シートモジュール
'// 「Tasks」シートのイベント（例: Worksheet_Change）を処理し、データ変更時にガントチャートの更新をトリガーします。
'////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler

    Dim dataRange As Range
    Set dataRange = Me.Range("B:G") ' B列からG列までを監視対象とする

    ' 変更されたセルが監視対象範囲内にあるか、かつ複数セルが変更されていないかを確認
    If Not Intersect(Target, dataRange) Is Nothing And Target.Cells.Count = 1 Then
        ' 変更された行がデータ行（ヘッダー行以外）であるかを確認
        If Target.Row >= 2 Then
            ' 終了日 (E列) が変更された場合は、期間 (C列) を再計算
            If Target.Column = 5 Then ' E列 (終了日)
                Me.Cells(Target.Row, R1C3).Value = Me.Cells(Target.Row, R1C5).Value - Me.Cells(Target.Row, R1C4).Value + 1
            End If

            ' 期間 (C列) または開始日 (D列) が変更された場合は、終了日 (E列) を再計算
            If Target.Column = 3 Or Target.Column = 4 Then ' C列 (期間) または D列 (開始日)
                Me.Cells(Target.Row, R1C5).Value = Me.Cells(Target.Row, R1C4).Value + Me.Cells(Target.Row, R1C3).Value - 1
            End If

            ' ガントチャートを更新
            Call M_GanttChart.UpdateGanttChart
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox "Tasksシートの変更イベント中にエラーが発生しました: " & Err.Description, vbCritical
End Sub
