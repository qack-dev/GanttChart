VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// Settings Class Module
'// Manages application settings from the Settings sheet.
'////////////////////////////////////////////////////////////////////////////////////////////////////

Private m_ChartStartRow As Long
Private m_ChartStartCol As Long
Private m_BarHeight As Long
Private m_RowHeight As Long
Private m_ColWidth As Long
Private m_ColorUnstarted As Long
Private m_ColorInProgress As Long
Private m_ColorCompleted As Long
Private m_ColorDelayed As Long

Public Property Get chartStartRow() As Long
    chartStartRow = m_ChartStartRow
End Property

Public Property Get chartStartCol() As Long
    chartStartCol = m_ChartStartCol
End Property

Public Property Get BarHeight() As Long
    BarHeight = m_BarHeight
End Property

Public Property Get RowHeight() As Long
    RowHeight = m_RowHeight
End Property

Public Property Get colWidth() As Long
    colWidth = m_ColWidth
End Property

Public Property Get ColorUnstarted() As Long
    ColorUnstarted = m_ColorUnstarted
End Property

Public Property Get ColorInProgress() As Long
    ColorInProgress = m_ColorInProgress
End Property

Public Property Get ColorCompleted() As Long
    ColorCompleted = m_ColorCompleted
End Property

Public Property Get ColorDelayed() As Long
    ColorDelayed = m_ColorDelayed
End Property

Public Sub LoadFromSheet(wsSettings As Worksheet)
    m_ChartStartRow = wsSettings.Cells(1, 2).value
    m_ChartStartCol = wsSettings.Cells(1, 3).value
    m_BarHeight = wsSettings.Cells(2, 2).value
    m_RowHeight = wsSettings.Cells(3, 2).value
    m_ColWidth = wsSettings.Cells(4, 2).value
    m_ColorUnstarted = wsSettings.Cells(6, 2).value
    m_ColorInProgress = wsSettings.Cells(7, 2).value
    m_ColorCompleted = wsSettings.Cells(8, 2).value
    m_ColorDelayed = wsSettings.Cells(9, 2).value
End Sub
