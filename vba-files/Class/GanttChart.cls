VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GanttChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// Sheet_GanttChart Sheet Module
'// Describes actions for buttons placed on the "GanttChart" sheet, etc.
'////////////////////////////////////////////////////////////////////////////////////////////////////

' Updates the Gantt chart when the button is clicked.
Private Sub UpdateChartButton_Click()
    On Error GoTo ErrHandler
    Call M_GanttChart.UpdateGanttChart
    Exit Sub
ErrHandler:
    MsgBox "チャート更新ボタンクリック時にエラーが発生しました: " & Err.Description, vbCritical
End Sub