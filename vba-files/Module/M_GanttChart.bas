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

    ' --- �`�揈�� ---
    Call DrawTimeline(wsGantt, minDate, maxDate)
    Call DrawAllTaskBars(wsGantt, wsTasks, lastTaskRow, minDate)
    ' Call UpdateLoadGraph(wsGantt, wsTasks, minDate, maxDate) ' �K�v�ɉ����ăR�����g����

    Exit Sub

ErrHandler:
    MsgBox "Error in DrawTaskBar: " & Err.Description, vbCritical
End Sub

' Draws the timeline
Private Sub DrawTimeline(wsGantt As Worksheet, startDate As Date, endDate As Date, _
                         chartStartRow As Long, chartStartCol As Long, colWidth As Long)
    On Error GoTo ErrHandler

'/**
' * @brief �K���g�`���[�g�̕`��G���A�i�^�C�����C���A�^�X�N�o�[�A�^�X�N���j���N���A���܂��B
' * @param wsGantt �Ώۂ�GanttChart�V�[�g
' * @param lastTaskRow Tasks�V�[�g�̍ŏI�s
' */
Private Sub ClearGanttArea(ByVal wsGantt As Worksheet, ByVal lastTaskRow As Long)
    On Error Resume Next ' �N���A�Ώۂ����݂��Ȃ��ꍇ���l��

    ' --- �^�C�����C���G���A�̃N���A ---
    wsGantt.Rows(TIMELINE_ROW).Clear

    ' --- �^�X�N�`��G���A�̃N���A ---
    ' �O��̕`��͈͂��s���Ȃ��߁A�\���Ȕ͈͂��N���A����
    Dim clearRange As Range
    Set clearRange = wsGantt.Range(wsGantt.Cells(GANTT_START_ROW, GANTT_START_COL), wsGantt.Cells(GANTT_START_ROW + lastTaskRow + 5, 256))
    
    With clearRange
        .ClearContents
        .Interior.Color = xlNone
        .Borders.LineStyle = xlNone
    End With
    
    On Error GoTo 0
End Sub

'/**
' * @brief �^�C�����C���i���t�w�b�_�[�j��`�悵�܂��B
' * @param wsGantt �Ώۂ�GanttChart�V�[�g
' * @param startDate �\������ŏ��̓��t
' * @param endDate �\������Ō�̓��t
' */
Private Sub DrawTimeline(ByVal wsGantt As Worksheet, ByVal startDate As Date, ByVal endDate As Date)
    Dim currentDate As Date
    Dim col As Long
    
    col = GANTT_START_COL + 1 ' �^�C�����C���̓^�X�N���̉E�̗񂩂�J�n

<<<<<<< HEAD
    ' --- ���t�̕`�� ---
    For currentDate = startDate To endDate
        With wsGantt.Cells(TIMELINE_ROW, col)
            .Value = Format(currentDate, "m/d")
            .ColumnWidth = 4
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 8
            
            ' --- �T���̃n�C���C�g ---
            If Weekday(currentDate) = vbSaturday Or Weekday(currentDate) = vbSunday Then
                .Interior.Color = RGB(240, 240, 240) ' �����D�F
            End If
            
            ' --- ���̋�؂�� ---
            If Day(currentDate) = 1 Then
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
            End If
        End With
        col = col + 1
    Next currentDate
End Sub

'/**
' * @brief ���ׂẴ^�X�N�o�[�i�Z���̒��F�j��`�悵�܂��B
' * @param wsGantt �Ώۂ�GanttChart�V�[�g
' * @param wsTasks Tasks�V�[�g
' * @param lastTaskRow Tasks�V�[�g�̍ŏI�s
' * @param minDate �^�C�����C���̊J�n��
' */
Private Sub DrawAllTaskBars(ByVal wsGantt As Worksheet, ByVal wsTasks As Worksheet, ByVal lastTaskRow As Long, ByVal minDate As Date)
    Dim i As Long
    For i = 2 To lastTaskRow
        Dim taskName As String
        Dim startDate As Date
        Dim duration As Long
        Dim status As String
        
        ' --- �^�X�N�����擾 ---
        With wsTasks.Rows(i)
            taskName = .Cells(COL_TASK_NAME).Value
            startDate = .Cells(COL_START_DATE).Value
            duration = .Cells(COL_DURATION).Value
            status = .Cells(COL_STATUS).Value
        End With
        
        ' --- �^�X�N�o�[��`�� ---
        Call HighlightTaskPeriod(wsGantt, i - 1, taskName, startDate, duration, status, minDate)
    Next i
End Sub

'/**
' * @brief �ʂ̃^�X�N�o�[�i�Z���̒��F�j��`�悵�܂��B
' * @param wsGantt �Ώۂ�GanttChart�V�[�g
' * @param taskRowIndex GanttChart�V�[�g��̃^�X�N�̍s�C���f�b�N�X (1����n�܂�)
' * @param taskName �^�X�N��
' * @param startDate �^�X�N�̊J�n��
' * @param duration �^�X�N�̊��ԁi�����j
' * @param status �^�X�N�̃X�e�[�^�X
' * @param minDate �^�C�����C���̊J�n��
' */
Private Sub HighlightTaskPeriod(ByVal wsGantt As Worksheet, ByVal taskRowIndex As Long, ByVal taskName As String, ByVal startDate As Date, ByVal duration As Long, ByVal status As String, ByVal minDate As Date)
    On Error GoTo ErrHandler

    Dim startCol As Long
    Dim endCol As Long
    Dim taskRow As Long
    Dim barColor As Long
    Dim taskRange As Range

    ' --- �`��ʒu�̌v�Z ---
    taskRow = GANTT_START_ROW + taskRowIndex - 1
    startCol = (startDate - minDate) + GANTT_START_COL + 1
    endCol = startCol + duration - 1

    ' --- �^�X�N���̕\�� ---
    wsGantt.Cells(taskRow, GANTT_START_COL).Value = taskName

    ' --- ���ԃZ���̓���ƒ��F ---
    If startCol <= endCol Then
        Set taskRange = wsGantt.Range(wsGantt.Cells(taskRow, startCol), wsGantt.Cells(taskRow, endCol))
        
        ' --- �X�e�[�^�X�ɉ������F���擾 ---
        barColor = GetColorByStatus(status)
        
        ' --- �Z���̏����ݒ� ---
        With taskRange.Interior
            .Color = barColor
        End With
        
        ' --- �^�X�N�o�[�ɘg����ǉ� ---
        With taskRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(150, 150, 150)
        End With
    End If

<<<<<<< HEAD
=======
    ' �O���t�̃f�[�^�͈͂�ݒ� (�ꎞ�I�ɃV�[�g�ɏ����o��)
    wsGantt.Cells(1, 1).Value = "�i��"
    wsGantt.Cells(1, 2).Value = progressPercentage
=======
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
>>>>>>> dev_tmp

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

>>>>>>> ebbfa819f61cab5c67d7badd685f80efd03e37fa
    Exit Sub

ErrHandler:
<<<<<<< HEAD
    MsgBox "�^�X�N�o�[�̕`�撆�ɃG���[���������܂���: " & vbCrLf & "�^�X�N��: " & taskName & vbCrLf & Err.Description, vbCritical
End Sub

'/**
' * @brief �X�e�[�^�X������ɑΉ�����F�萔��Ԃ��܂��B
' * @param status �^�X�N�̃X�e�[�^�X
' * @return �Ή�����F��Long�l
' */
Private Function GetColorByStatus(ByVal status As String) As Long
    Select Case status
        Case "������"
            GetColorByStatus = COLOR_UNSTARTED
        Case "�i�s��"
            GetColorByStatus = COLOR_IN_PROGRESS
        Case "����"
            GetColorByStatus = COLOR_COMPLETED
        Case "�x��"
            GetColorByStatus = COLOR_DELAYED
        Case Else
            GetColorByStatus = vbWhite ' �s���ȃX�e�[�^�X�͔�
    End Select
<<<<<<< HEAD
End Function

'/**
' * @brief �i�Q�l�j���׃O���t���X�V���܂��B����̉��C�͈͊O�ł����A�K�v�ɉ����ė��p���܂��B
' */
Private Sub UpdateLoadGraph(wsGantt As Worksheet, wsTasks As Worksheet, minDate As Date, maxDate As Date)
    ' ���̃v���V�[�W���͍���̉��C�v���ɂ͊܂܂�Ă��܂��񂪁A
    ' �K�v�ɉ����ăZ���x�[�X�̃f�[�^�ƘA�g����悤�ɉ��C�\�ł��B
    ' (���݂̎�����Shape�Ɉˑ����Ă���\�������邽�߁A���r���[���K�v�ł�)
    MsgBox "UpdateLoadGraph�͌��ݎ�������Ă��܂���B", vbInformation
End Sub
=======
End Function
>>>>>>> ebbfa819f61cab5c67d7badd685f80efd03e37fa
=======
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
>>>>>>> dev_tmp
