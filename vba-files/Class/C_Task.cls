Attribute VB_Name = "C_Task"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// C_Task Class Module
'// Represents a single task object with properties.
'////////////////////////////////////////////////////////////////////////////////////////////////////

Private m_TaskID As Long
Private m_TaskName As String
Private m_Duration As Long
Private m_StartDate As Date
Private m_EndDate As Date
Private m_Progress As Double
Private m_Status As String

Public Property Get TaskID() As Long
    TaskID = m_TaskID
End Property

Public Property Let TaskID(ByVal value As Long)
    m_TaskID = value
End Property

Public Property Get TaskName() As String
    TaskName = m_TaskName
End Property

Public Property Let TaskName(ByVal value As String)
    m_TaskName = value
End Property

Public Property Get Duration() As Long
    Duration = m_Duration
End Property

Public Property Let Duration(ByVal value As Long)
    m_Duration = value
End Property

Public Property Get StartDate() As Date
    StartDate = m_StartDate
End Property

Public Property Let StartDate(ByVal value As Date)
    m_StartDate = value
End Property

Public Property Get EndDate() As Date
    EndDate = m_EndDate
End Property

Public Property Let EndDate(ByVal value As Date)
    m_EndDate = value
End Property

Public Property Get Progress() As Double
    Progress = m_Progress
End Property

Public Property Let Progress(ByVal value As Double)
    m_Progress = value
End Property

Public Property Get Status() As String
    Status = m_Status
End Property

Public Property Let Status(ByVal value As String)
    m_Status = value
End Property