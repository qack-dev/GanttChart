Attribute VB_Name = "Settings"
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

Public Property Get ChartStartRow() As Long
    ChartStartRow = m_ChartStartRow
End Property

Public Property Get ChartStartCol() As Long
    ChartStartCol = m_ChartStartCol
End Property

Public Property Get BarHeight() As Long
    BarHeight = m_BarHeight
End Property

Public Property Get RowHeight() As Long
    RowHeight = m_RowHeight
End Property

Public Property Get ColWidth() As Long
    ColWidth = m_ColWidth
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
    m_ChartStartRow = wsSettings.Cells(1, 2).Value
    m_ChartStartCol = wsSettings.Cells(1, 3).Value
    m_BarHeight = wsSettings.Cells(2, 2).Value
    m_RowHeight = wsSettings.Cells(3, 2).Value
    m_ColWidth = wsSettings.Cells(4, 2).Value
    m_ColorUnstarted = wsSettings.Cells(6, 2).Value
    m_ColorInProgress = wsSettings.Cells(7, 2).Value
    m_ColorCompleted = wsSettings.Cells(8, 2).Value
    m_ColorDelayed = wsSettings.Cells(9, 2).Value
End Sub