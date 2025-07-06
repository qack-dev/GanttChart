Attribute VB_Name = "Sheet_GanttChart"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
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