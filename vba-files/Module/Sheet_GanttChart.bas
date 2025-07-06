Attribute VB_Name = "Sheet_GanttChart"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// Sheet_GanttChart シートモジュール
'// 「GanttChart」シートに配置されたボタンのアクションなどを記述します。
'////////////////////////////////////////////////////////////////////////////////////////////////////

' このシートにボタンを配置し、このマクロを登録してください。
' ボタンがクリックされたときにガントチャートを更新します。
Public Sub UpdateChartButton_Click()
    On Error GoTo ErrHandler
    Call M_GanttChart.UpdateGanttChart
    Exit Sub
ErrHandler:
    MsgBox "チャート更新ボタンクリック時にエラーが発生しました: " & Err.Description, vbCritical
End Sub
