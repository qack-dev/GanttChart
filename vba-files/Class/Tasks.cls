Attribute VB_Name = "Tasks"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// Tasks Class Module
'// Manages a collection of C_Task objects.
'////////////////////////////////////////////////////////////////////////////////////////////////////

Private m_Tasks As Collection

Private Sub Class_Initialize()
    Set m_Tasks = New Collection
End Sub

Public Sub Add(task As C_Task)
    m_Tasks.Add task, CStr(task.TaskID)
End Sub

Public Function Item(index As Variant) As C_Task
    Set Item = m_Tasks.Item(index)
End Function

Public Function Count() As Long
    Count = m_Tasks.Count
End Function

Public Sub LoadFromSheet(wsTasks As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim task As C_Task

    lastRow = wsTasks.Cells(wsTasks.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For i = 2 To lastRow
        Set task = New C_Task
        With task
            .TaskID = wsTasks.Cells(i, 1).Value
            .TaskName = wsTasks.Cells(i, 2).Value
            .Duration = wsTasks.Cells(i, 3).Value
            .StartDate = wsTasks.Cells(i, 4).Value
            .EndDate = wsTasks.Cells(i, 5).Value
            .Progress = wsTasks.Cells(i, 6).Value
            .Status = wsTasks.Cells(i, 7).Value
        End With
        Add task
    Next i
End Sub

Public Function GetMinDate() As Date
    Dim minD As Date
    Dim task As C_Task
    minD = #12/31/9999#
    For Each task In m_Tasks
        If task.StartDate < minD Then
            minD = task.StartDate
        End If
    Next task
    GetMinDate = minD
End Function

Public Function GetMaxDate() As Date
    Dim maxD As Date
    Dim task As C_Task
    maxD = #1/1/1900#
    For Each task In m_Tasks
        If task.EndDate > maxD Then
            maxD = task.EndDate
        End If
    Next task
    GetMaxDate = maxD
End Function