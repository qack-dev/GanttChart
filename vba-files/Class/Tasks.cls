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
'// Tasks Class Module
'// Manages a collection of task data using Scripting.Dictionary objects.
'////////////////////////////////////////////////////////////////////////////////////////////////////

Private m_Tasks As Collection

Private Sub class_initialize()
    Set m_Tasks = New Collection
End Sub

Public Function Item(index As Variant) As Object
    Set Item = m_Tasks.Item(index)
End Function

Public Function Count() As Long
    Count = m_Tasks.Count
End Function

Public Sub LoadFromSheet(wsTasks As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim taskData As Object ' Use Object for Scripting.Dictionary

    ' Clear existing tasks before loading new ones
    Set m_Tasks = New Collection

    lastRow = wsTasks.Cells(wsTasks.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For i = 2 To lastRow
        Set taskData = CreateObject("Scripting.Dictionary")
        With taskData
            .Add "TaskID", wsTasks.Cells(i, 1).value
            .Add "TaskName", wsTasks.Cells(i, 2).value
            .Add "Duration", wsTasks.Cells(i, 3).value
            .Add "StartDate", wsTasks.Cells(i, 4).value
            .Add "EndDate", wsTasks.Cells(i, 5).value
            .Add "Progress", wsTasks.Cells(i, 6).value
            .Add "Status", wsTasks.Cells(i, 7).value
        End With
        
        ' Add the task data to the collection, handling potential duplicate keys
        On Error Resume Next
        m_Tasks.Add taskData, CStr(taskData("TaskID"))
        If Err.Number <> 0 Then
            ' If a duplicate key error occurs, ignore it and continue.
            ' You could add logging here if needed.
            Err.Clear
        End If
        On Error GoTo 0
    Next i
End Sub

Public Function GetMinDate() As Date
    Dim minD As Date
    Dim taskData As Object
    minD = #12/31/9999#
    For Each taskData In m_Tasks
        If taskData("StartDate") < minD Then
            minD = taskData("StartDate")
        End If
    Next taskData
    GetMinDate = minD
End Function

Public Function GetMaxDate() As Date
    Dim maxD As Date
    Dim taskData As Object
    maxD = #1/1/1900#
    For Each taskData In m_Tasks
        If taskData("EndDate") > maxD Then
            maxD = taskData("EndDate")
        End If
    Next taskData
    GetMaxDate = maxD
End Function

' Allows "For Each" iteration over the collection
Public Function NewEnum() As IUnknown
    Set NewEnum = m_Tasks.[_NewEnum]
End Function

