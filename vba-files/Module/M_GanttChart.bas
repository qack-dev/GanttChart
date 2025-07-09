Attribute VB_Name = "M_GanttChart"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// M_GanttChart ���W���[��
'// �K���g�`���[�g�̕`��A�X�V�A�S�̐i���O���t�̍X�V�ȂǁA��v�ȃ��W�b�N���i�[���܂��B
'// --- �ύX���� ---
'// [2025/07/07] ������̃f�o�b�O���o�āA���蓮�삷��ŏI�ŃR�[�h�ɏC���B
'//               �s�����WorksheetFunction�̎g�p�𒆎~���A���S�ȃ��[�v�����ɉ�A�B
'////////////////////////////////////////////////////////////////////////////////////////////////////

' --- Tasks�V�[�g�̗�C���f�b�N�X ---
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME As String = "B"
Private Const COL_DURATION As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_END_DATE As Long = 5
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

' --- Settings�V�[�g�̒�` ---
Private Const SETTINGS_VALUE_COL As Long = 2
Private Const SETTINGS_CHART_START_COL As Long = 3
Private Const SETTING_ROW_CHART_START As Long = 1
Private Const SETTING_ROW_BAR_HEIGHT As Long = 2
Private Const SETTING_ROW_ROW_HEIGHT As Long = 3
Private Const SETTING_ROW_COL_WIDTH As Long = 4
Private Const SETTING_ROW_COLOR_UNSTARTED As Long = 5
Private Const SETTING_ROW_COLOR_IN_PROGRESS As Long = 6
Private Const SETTING_ROW_COLOR_COMPLETED As Long = 7
Private Const SETTING_ROW_COLOR_DELAYED As Long = 8

' --- �V�F�C�v���̐ړ��� ---
Private Const SHAPE_PREFIX_TASK_BAR As String = "TaskBar_"
Private Const SHAPE_NAME_PROGRESS_CHART As String = "OverallProgressChart"

Public Sub UpdateGanttChart()
    Dim procName As String: procName = "UpdateGanttChart"
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim wsGantt As Worksheet, wsTasks As Worksheet, wsSettings As Worksheet
    Set wsGantt = ThisWorkbook.Sheets("GanttChart")
    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    Call ClearGanttChart(wsGantt)

    Dim chartStartRow As Long, chartStartCol As Long, barHeight As Long, rowHeight As Long
    Dim colWidth As Double
    chartStartRow = wsSettings.Cells(SETTING_ROW_CHART_START, SETTINGS_VALUE_COL).Value
    chartStartCol = wsSettings.Cells(SETTING_ROW_CHART_START, SETTINGS_CHART_START_COL).Value
    barHeight = wsSettings.Cells(SETTING_ROW_BAR_HEIGHT, SETTINGS_VALUE_COL).Value
    rowHeight = wsSettings.Cells(SETTING_ROW_ROW_HEIGHT, SETTINGS_VALUE_COL).Value
    colWidth = wsSettings.Cells(SETTING_ROW_COL_WIDTH, SETTINGS_VALUE_COL).Value

    Dim lastTaskRow As Long
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    If lastTaskRow < 2 Then GoTo CleanUp

    ' --- �v���W�F�N�g�̓��t�͈͂���� (���S�ȃ��[�v����) --- ������ �C���ӏ� ������
    Dim minDate As Date, maxDate As Date
    Dim i As Long
    ' �ŏ��̗L���ȓ��t���������ď����l�ɐݒ�
    For i = 2 To lastTaskRow
        If IsDate(wsTasks.Cells(i, COL_START_DATE).Value) Then
            minDate = wsTasks.Cells(i, COL_START_DATE).Value
            maxDate = wsTasks.Cells(i, COL_END_DATE).Value
            Exit For
        End If
    Next i
    
    ' �c��̃^�X�N�Ɣ�r���Đ^�̍ŏ����E�ő�������߂�
    For i = i + 1 To lastTaskRow
        If IsDate(wsTasks.Cells(i, COL_START_DATE).Value) And wsTasks.Cells(i, COL_START_DATE).Value < minDate Then
            minDate = wsTasks.Cells(i, COL_START_DATE).Value
        End If
        If IsDate(wsTasks.Cells(i, COL_END_DATE).Value) And wsTasks.Cells(i, COL_END_DATE).Value > maxDate Then
            maxDate = wsTasks.Cells(i, COL_END_DATE).Value
        End If
    Next i


    Call DrawTimeline(wsGantt, minDate, maxDate, chartStartRow, chartStartCol, colWidth, rowHeight)

    For i = 2 To lastTaskRow
        Call DrawTaskBar(wsGantt, wsTasks.Cells(i, COL_TASK_ID).Value, wsTasks.Cells(i, COL_TASK_NAME).Value, _
                         wsTasks.Cells(i, COL_START_DATE).Value, wsTasks.Cells(i, COL_END_DATE).Value, _
                         wsTasks.Cells(i, COL_STATUS).Value, chartStartRow + i - 1, chartStartCol, _
                         colWidth, barHeight, minDate, rowHeight)
    Next i

    Call UpdateOverallProgressChart(wsGantt, wsTasks, lastTaskRow, chartStartRow + lastTaskRow + 2, chartStartCol)

CleanUp:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "�G���[���������܂��� (" & procName & "): " & vbCrLf & Err.Description, vbCritical
    GoTo CleanUp
End Sub

Private Sub ClearGanttChart(wsGantt As Worksheet)
    On Error Resume Next
    Dim sh As Shape
    For Each sh In wsGantt.Shapes
        If sh.Name Like SHAPE_PREFIX_TASK_BAR & "*" Or sh.Name = SHAPE_NAME_PROGRESS_CHART Then sh.Delete
    Next sh
    wsGantt.Range("A4:ZZ100").Clear
    wsGantt.Range("A4:ZZ100").Interior.ColorIndex = xlNone
    On Error GoTo 0
End Sub

Private Sub DrawTaskBar(wsGantt As Worksheet, taskID As Long, taskName As String, _
                        startDate As Date, endDate As Date, status As String, _
                        rowNum As Long, chartStartCol As Long, colWidth As Double, barHeight As Long, _
                        minChartDate As Date, taskRowHeight As Long)
    ' �J�n���܂��͏I�����������ȏꍇ�͕`�悵�Ȃ�
    If Not IsDate(startDate) Or Not IsDate(endDate) Then Exit Sub

    wsGantt.Rows(rowNum).rowHeight = taskRowHeight

    Dim barLeft As Double, barTop As Double, barWidth As Double
    barLeft = wsGantt.Columns(chartStartCol).Left + (startDate - minChartDate) * colWidth
    barWidth = (CDbl(endDate) - CDbl(startDate) + 1) * colWidth
    barTop = wsGantt.Rows(rowNum).Top + (wsGantt.Rows(rowNum).Height - barHeight) / 2

    Dim sh As Shape
    Set sh = wsGantt.Shapes.AddShape(msoShapeRectangle, barLeft, barTop, barWidth, barHeight)
    With sh
        .Fill.ForeColor.RGB = GetColorByStatus(status)
        .Line.Visible = msoFalse
        .Name = SHAPE_PREFIX_TASK_BAR & taskID
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 5: .MarginRight = 5: .WordWrap = msoFalse
            With .TextRange
                .Text = taskName
                .ParagraphFormat.Alignment = msoAlignLeft
                With .Font
                    .Fill.ForeColor.RGB = IIf(.Parent.Parent.Parent.Fill.ForeColor.RGB = &HC0C0C0, RGB(0, 0, 0), RGB(255, 255, 255))
                    .Size = 9
                    .Bold = msoTrue
                End With
            End With
        End With
    End With
End Sub

Private Sub DrawTimeline(wsGantt As Worksheet, startDate As Date, endDate As Date, _
                         chartStartRow As Long, chartStartCol As Long, colWidth As Double, taskRowHeight As Long)
    Dim timelineHeaderRow As Long: timelineHeaderRow = chartStartRow - 2
    Dim timelineDayRow As Long: timelineDayRow = chartStartRow - 1
    
    With wsGantt.Rows(timelineHeaderRow & ":" & timelineDayRow)
        .Clear
        .rowHeight = 15
    End With
    wsGantt.Rows(timelineDayRow).rowHeight = 30
    
    Dim totalDays As Long: totalDays = endDate - startDate + 10
    If totalDays <= 0 Then totalDays = 100 ' �G���[���
    wsGantt.Columns(chartStartCol).Resize(, totalDays).ColumnWidth = colWidth / 7

    Dim currentDate As Date, colOffset As Long, currentColumn As Long
    Dim yearMonth As String: yearMonth = Format(startDate, "yyyy/mm")
    Dim mergeStartCol As Long: mergeStartCol = chartStartCol
    
    For colOffset = 0 To totalDays - 10
        currentDate = startDate + colOffset
        currentColumn = chartStartCol + colOffset

        If Format(currentDate, "yyyy/mm") <> yearMonth Then
            wsGantt.Range(wsGantt.Cells(timelineHeaderRow, mergeStartCol), wsGantt.Cells(timelineHeaderRow, currentColumn - 1)).Merge
            yearMonth = Format(currentDate, "yyyy/mm")
            mergeStartCol = currentColumn
        End If
        With wsGantt.Cells(timelineHeaderRow, mergeStartCol)
            .Value = yearMonth: .HorizontalAlignment = xlCenter
        End With
        With wsGantt.Cells(timelineDayRow, currentColumn)
            .Value = Format(currentDate, "d"): .HorizontalAlignment = xlCenter
        End With
        If Weekday(currentDate) = vbSaturday Or Weekday(currentDate) = vbSunday Then
            wsGantt.Range(wsGantt.Cells(timelineDayRow, currentColumn), wsGantt.Cells(100, currentColumn)).Interior.Color = RGB(242, 242, 242)
        End If
    Next colOffset
    
    wsGantt.Range(wsGantt.Cells(timelineHeaderRow, mergeStartCol), wsGantt.Cells(timelineHeaderRow, currentColumn)).Merge
End Sub

Private Sub UpdateOverallProgressChart(wsGantt As Worksheet, wsTasks As Worksheet, lastTaskRow As Long, chartTopRow As Long, chartStartCol As Long)
    Dim totalWorkload As Double, completedWorkload As Double, progressPercentage As Double
    Dim duration As Double, progress As Double, i As Long

    For i = 2 To lastTaskRow
        If IsNumeric(wsTasks.Cells(i, COL_DURATION).Value) And IsNumeric(wsTasks.Cells(i, COL_PROGRESS).Value) Then
            duration = wsTasks.Cells(i, COL_DURATION).Value
            progress = wsTasks.Cells(i, COL_PROGRESS).Value
            totalWorkload = totalWorkload + duration
            completedWorkload = completedWorkload + (duration * progress)
        End If
    Next i

    If totalWorkload > 0 Then progressPercentage = completedWorkload / totalWorkload Else progressPercentage = 0

    Dim dataRange As Range: Set dataRange = wsGantt.Range("Z1:Z2")
    dataRange.Cells(1, 1).Value = progressPercentage
    dataRange.Cells(2, 1).Value = 1 - progressPercentage

    Dim chObj As ChartObject
    Set chObj = wsGantt.ChartObjects.Add(Left:=wsGantt.Columns(2).Left, Top:=wsGantt.Rows(chartTopRow).Top, Width:=200, Height:=120)
    With chObj
        .Name = SHAPE_NAME_PROGRESS_CHART
        With .Chart
            .ChartType = xlDoughnut
            .SetSourceData Source:=dataRange
            .HasLegend = False
            .DoughnutGroups(1).DoughnutHoleSize = 75
            With .SeriesCollection(1)
                ' �|�C���g1 (��������) �̏����ݒ�
                With .Points(1).Format.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = GetColorByStatus("����")
                    .Solid
                End With
                
                ' �|�C���g2 (����������) �̏����ݒ�
                With .Points(2).Format.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(220, 220, 220)
                    .Solid
                End With

                ' �n��̘g���̐F�𔒂ɐݒ�
                .Border.Color = RGB(255, 255, 255)
            End With
            .HasTitle = True
            With .ChartTitle
                .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(89, 89, 89)
                .Format.TextFrame2.TextRange.Font.Size = 14
                .Format.TextFrame2.TextRange.Font.Bold = msoTrue
                .Text = Format(progressPercentage, "0.0%")
            End With
            .PlotArea.Format.Fill.Visible = msoFalse
            .PlotArea.Format.Fill.Visible = msoFalse '�v���b�g�G���A�͓����̂܂�

            ' �O���t�S�̂̔w�i�iChartArea�j��ݒ�
            With .ChartArea.Format.Fill
                .Visible = msoTrue ' �h��Ԃ���L���ɂ���
                .ForeColor.RGB = RGB(225, 235, 250) ' �w�i�F��ɐݒ�
                .Solid ' �P�F�œh��Ԃ�
            End With
            
            .ChartArea.Format.Line.Visible = msoFalse ' �O�g�̐��͔�\���̂܂�
        End With
    End With
    ' dataRange.ClearContents
End Sub

Private Function GetColorByStatus(status As String) As Long
    On Error Resume Next
    Dim wsSettings As Worksheet: Set wsSettings = ThisWorkbook.Sheets("Settings")
    Select Case status
        Case "������": GetColorByStatus = wsSettings.Cells(SETTING_ROW_COLOR_UNSTARTED, SETTINGS_VALUE_COL).Value
        Case "�i�s��": GetColorByStatus = wsSettings.Cells(SETTING_ROW_COLOR_IN_PROGRESS, SETTINGS_VALUE_COL).Value
        Case "����": GetColorByStatus = wsSettings.Cells(SETTING_ROW_COLOR_COMPLETED, SETTINGS_VALUE_COL).Value
        Case "�x��": GetColorByStatus = wsSettings.Cells(SETTING_ROW_COLOR_DELAYED, SETTINGS_VALUE_COL).Value
        Case Else: GetColorByStatus = RGB(192, 192, 192)
    End Select
End Function
