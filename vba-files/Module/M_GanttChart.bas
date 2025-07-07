Attribute VB_Name = "M_GanttChart"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// M_GanttChart Module
'// Contains the main logic for drawing, updating the Gantt chart, and updating the overall progress graph.
'////////////////////////////////////////////////////////////////////////////////////////////////////

' Chart Objects Name
Private Const OVERALL_PROGRESS_CHART_NAME As String = "OverallProgressChart"
Private Const TASK_BAR_NAME_PREFIX As String = "TaskBar_"
Private Const TIMELINE_NAME_PREFIX As String = "Timeline_"
Private Const PROGRESS_NAME_PREFIX As String = "Progress_"

' Task Status
Private Const STATUS_UNSTARTED As String = "Unstarted"
Private Const STATUS_IN_PROGRESS As String = "In Progress"
Private Const STATUS_COMPLETED As String = "Completed"
Private Const STATUS_DELAYED As String = "Delayed"


' Main procedure to update the Gantt chart
Public Sub UpdateGanttChart()
    On Error GoTo ErrHandler

    Dim wsGantt As Worksheet
    Dim wsTasks As Worksheet
    Dim wsSettings As Worksheet
    Dim appSettings As Settings
    Dim allTasks As Tasks
    Dim task As Object
    Dim minDate As Date
    Dim maxDate As Date
    Dim i As Long

    Set wsGantt = ThisWorkbook.Sheets("GanttChart")
    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    ' Load settings and tasks
    Set appSettings = New Settings
    appSettings.LoadFromSheet wsSettings
    
    Set allTasks = New Tasks
    allTasks.LoadFromSheet wsTasks

    If allTasks.Count = 0 Then
        MsgBox "No task data found.", vbInformation
        Exit Sub
    End If

    ' Clear the old chart elements
    Call ClearGanttChart(wsGantt)

    ' Determine the date range for the chart
    minDate = allTasks.GetMinDate
    maxDate = allTasks.GetMaxDate

    ' Draw the timeline
    Call DrawTimeline(wsGantt, minDate, maxDate, appSettings.ChartStartRow, appSettings.ChartStartCol, appSettings.ColWidth)

    ' Draw the bar for each task
    For i = 1 To allTasks.Count
        Set task = allTasks.Item(i)
        Call DrawTaskBar(wsGantt, task, appSettings, minDate, i - 1)
    Next i

    ' Update the overall progress chart
    Call UpdateOverallProgressChart(wsGantt, allTasks, appSettings)

    Exit Sub

ErrHandler:
    MsgBox "Error in UpdateGanttChart: " & Err.Description, vbCritical
End Sub

' Clears the Gantt chart of all shapes
Private Sub ClearGanttChart(wsGantt As Worksheet)
    On Error Resume Next ' Continue if a shape is not found

    Dim sh As Shape
    For Each sh In wsGantt.Shapes
        If Left(sh.Name, Len(TASK_BAR_NAME_PREFIX)) = TASK_BAR_NAME_PREFIX Or _
           Left(sh.Name, Len(TIMELINE_NAME_PREFIX)) = TIMELINE_NAME_PREFIX Or _
           Left(sh.Name, Len(PROGRESS_NAME_PREFIX)) = PROGRESS_NAME_PREFIX Or _
           sh.Type = msoChart Then
            sh.Delete
        End If
    Next sh

    On Error GoTo 0 ' Reset error handling
End Sub

' Draws a bar corresponding to a single task
Private Sub DrawTaskBar(wsGantt As Worksheet, task As Object, appSettings As Settings, minChartDate As Date, index As Long)
    On Error GoTo ErrHandler

    Dim barLeft As Double
    Dim barTop As Double
    Dim barWidth As Double
    Dim barColor As Long
    Dim taskShape As Shape
    Dim rowNum As Long

    rowNum = appSettings.ChartStartRow + index

    ' Calculate the starting position and width of the bar
    barLeft = wsGantt.Cells(rowNum, 1).Left + (task("StartDate") - minChartDate) * appSettings.ColWidth
    barTop = wsGantt.Cells(rowNum, 1).Top + (wsGantt.Cells(rowNum, 1).Height - appSettings.BarHeight) / 2
    barWidth = (task("EndDate") - task("StartDate") + 1) * appSettings.ColWidth

    ' Get the color based on the status
    barColor = GetColorByStatus(task("Status"), appSettings)

    ' Draw the bar
    Set taskShape = wsGantt.Shapes.AddShape(msoShapeRectangle, barLeft, barTop, barWidth, appSettings.BarHeight)
    With taskShape
        .Fill.ForeColor.RGB = barColor
        .Line.Visible = msoFalse
        .Name = TASK_BAR_NAME_PREFIX & task("TaskID")
        .OnAction = "M_ChartEvents.ShowTaskDetails"
        .TextFrame2.TextRange.Text = task("TaskName")
        With .TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
            .Solid
        End With
        .TextFrame2.TextRange.Font.Size = 8
        .TextFrame2.TextRange.Font.Bold = msoFalse
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.WordArtformat = msoTextEffect1
    End With

    Exit Sub

ErrHandler:
    MsgBox "Error in DrawTaskBar: " & Err.Description, vbCritical
End Sub

' Draws the timeline
Private Sub DrawTimeline(wsGantt As Worksheet, startDate As Date, endDate As Date, _
                         chartStartRow As Long, chartStartCol As Long, colWidth As Long)
    On Error GoTo ErrHandler

    Dim currentDate As Date
    Dim colOffset As Long
    Dim headerRow As Long

    headerRow = chartStartRow - 1

    ' Clear the timeline header
    wsGantt.Range(wsGantt.Cells(headerRow, chartStartCol), wsGantt.Cells(headerRow, chartStartCol + (endDate - startDate + 1))).Clear

    colOffset = 0
    For currentDate = startDate To endDate
        With wsGantt.Cells(headerRow, chartStartCol + colOffset)
            .Value = Format(currentDate, "m/d")
            .ColumnWidth = colWidth / 6
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Orientation = 90

            ' Change the background color of weekends
            If Weekday(currentDate) = vbSaturday Or Weekday(currentDate) = vbSunday Then
                .Interior.Color = RGB(220, 220, 220)
            Else
                .Interior.Pattern = xlNone
            End If
        End With
        colOffset = colOffset + 1
    Next currentDate

    Exit Sub

ErrHandler:
    MsgBox "Error in DrawTimeline: " & Err.Description, vbCritical
End Sub

' Updates the overall progress chart
Private Sub UpdateOverallProgressChart(wsGantt As Worksheet, allTasks As Tasks, appSettings As Settings)
    On Error GoTo ErrHandler

    Dim totalDuration As Double
    Dim completedDuration As Double
    Dim progressPercentage As Double
    Dim chartObj As ChartObject
    Dim chartData(1 To 2) As Double
    Dim task As Object
    Dim i As Long

    ' Delete the old chart
    On Error Resume Next
    wsGantt.ChartObjects(OVERALL_PROGRESS_CHART_NAME).Delete
    On Error GoTo ErrHandler

    totalDuration = 0
    completedDuration = 0

    For i = 1 To allTasks.Count
        Set task = allTasks.Item(i)
        totalDuration = totalDuration + task("Duration")
        If task("Status") = STATUS_COMPLETED Then
            completedDuration = completedDuration + task("Duration")
        Else
            completedDuration = completedDuration + (task("Duration") * task("Progress"))
        End If
    Next i

    If totalDuration > 0 Then
        progressPercentage = completedDuration / totalDuration
    Else
        progressPercentage = 0
    End If

    ' Store the chart data in an array
    chartData(1) = progressPercentage
    chartData(2) = 1 - progressPercentage

    ' Create the chart
    Set chartObj = wsGantt.ChartObjects.Add( _
        Left:=wsGantt.Cells(appSettings.ChartStartRow + allTasks.Count, 1).Left, _
        Top:=wsGantt.Cells(appSettings.ChartStartRow + allTasks.Count, 1).Top + 20, _
        Width:=200, _
        Height:=120)

    With chartObj
        .Name = OVERALL_PROGRESS_CHART_NAME
        With .Chart
            .ChartType = xlDoughnut
            .HasTitle = True
            .ChartTitle.Text = "Overall Progress"
            .ChartTitle.Font.Size = 10
            .HasLegend = False
            .ChartGroups(1).DoughnutHoleSize = 75

            ' Set the data from the array
            With .SeriesCollection.NewSeries
                .Values = chartData
                .Points(1).Format.Fill.ForeColor.RGB = RGB(0, 176, 80) ' Completed (Green)
                .Points(2).Format.Fill.ForeColor.RGB = RGB(220, 220, 220) ' Incomplete (Gray)
                .Points(1).Border.LineStyle = xlNone
                .Points(2).Border.LineStyle = xlNone

                ' Remove all data labels
                .ApplyDataLabels
                .DataLabels.Delete

                ' Add a data label for the center
                .Points(1).ApplyDataLabels
                With .DataLabels(1)
                    .Text = Format(progressPercentage, "0%")
                    .Font.Size = 12
                    .Font.Bold = True
                    .Position = xlLabelPositionCenter
                End With
            End With
        End With
    End With

    Exit Sub

ErrHandler:
    MsgBox "Error in UpdateOverallProgressChart: " & Err.Description, vbCritical
End Sub

' Returns a color based on the status
Private Function GetColorByStatus(status As String, appSettings As Settings) As Long
    Select Case status
        Case STATUS_UNSTARTED
            GetColorByStatus = appSettings.ColorUnstarted
        Case STATUS_IN_PROGRESS
            GetColorByStatus = appSettings.ColorInProgress
        Case STATUS_COMPLETED
            GetColorByStatus = appSettings.ColorCompleted
        Case STATUS_DELAYED
            GetColorByStatus = appSettings.ColorDelayed
        Case Else
            GetColorByStatus = RGB(192, 192, 192) ' Default Color (Gray)
    End Select
End Function
