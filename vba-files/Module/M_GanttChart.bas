Attribute VB_Name = "M_GanttChart"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// M_GanttChart ���W���[��
'// �K���g�`���[�g�̕`��A�X�V�A�S�̐i���O���t�̍X�V�ȂǁA��v�ȃ��W�b�N���i�[���܂��B
'////////////////////////////////////////////////////////////////////////////////////////////////////

' Tasks�V�[�g�̗�C���f�b�N�X
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME As Long = 2
Private Const COL_DURATION As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_END_DATE As Long = 5
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

' Settings�V�[�g�̃Z���Q��
Private Const SETTING_CHART_START_ROW As Long = 1
Private Const SETTING_CHART_START_COL As Long = 2
Private Const SETTING_BAR_HEIGHT As Long = 3
Private Const SETTING_ROW_HEIGHT As Long = 4
Private Const SETTING_COL_WIDTH As Long = 5
Private Const SETTING_COLOR_UNSTARTED As Long = 6
Private Const SETTING_COLOR_IN_PROGRESS As Long = 7
Private Const SETTING_COLOR_COMPLETED As Long = 8
Private Const SETTING_COLOR_DELAYED As Long = 9

' �K���g�`���[�g���X�V���郁�C���v���V�[�W��
Public Sub UpdateGanttChart()
    On Error GoTo ErrHandler

    Dim wsGantt As Worksheet
    Dim wsTasks As Worksheet
    Dim wsSettings As Worksheet
    Dim lastTaskRow As Long
    Dim i As Long
    Dim taskID As Long
    Dim taskName As String
    Dim duration As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim progress As Double
    Dim status As String
    Dim minDate As Date
    Dim maxDate As Date
    Dim chartStartCol As Long
    Dim chartStartRow As Long
    Dim barHeight As Long
    Dim rowHeight As Long
    Dim colWidth As Long

    Set wsGantt = ThisWorkbook.Sheets("GanttChart")
    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    ' �����̃`���[�g���N���A
    Call ClearGanttChart(wsGantt)

    ' �ݒ�l�̓ǂݍ���
    chartStartRow = wsSettings.Cells(SETTING_CHART_START_ROW, SETTING_CHART_START_COL).Value ' ��: Settings!B1 �ɊJ�n�s
    chartStartCol = wsSettings.Cells(SETTING_CHART_START_ROW, SETTING_CHART_START_COL + 1).Value ' ��: Settings!C1 �ɊJ�n��
    barHeight = wsSettings.Cells(SETTING_CHART_START_ROW, SETTING_BAR_HEIGHT + 1).Value    ' ��: Settings!D1 �Ƀo�[�̍���
    rowHeight = wsSettings.Cells(SETTING_CHART_START_ROW, SETTING_ROW_HEIGHT + 1).Value    ' ��: Settings!E1 �ɍs�̍���
    colWidth = wsSettings.Cells(SETTING_CHART_START_ROW, SETTING_COL_WIDTH + 1).Value     ' ��: Settings!F1 �ɗ�̕�

    ' �^�X�N�f�[�^�̍ŏI�s���擾 (Tasks�V�[�g��B����)
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row

    If lastTaskRow < 2 Then ' �w�b�_�[�s�݂̂̏ꍇ
        MsgBox "�^�X�N�f�[�^������܂���B", vbInformation
        Exit Sub
    End If

    ' ���t�͈͂̓���
    minDate = wsTasks.Cells(2, COL_START_DATE).Value ' �J�n���̃w�b�_�[
    maxDate = wsTasks.Cells(2, COL_END_DATE).Value ' �I�����̃w�b�_�[

    For i = 2 To lastTaskRow
        If wsTasks.Cells(i, COL_START_DATE).Value < minDate Then minDate = wsTasks.Cells(i, COL_START_DATE).Value
        If wsTasks.Cells(i, COL_END_DATE).Value > maxDate Then maxDate = wsTasks.Cells(i, COL_END_DATE).Value
    Next i

    ' �^�C�����C���̕`��
    Call DrawTimeline(wsGantt, minDate, maxDate, chartStartRow, chartStartCol, colWidth)

    ' �e�^�X�N�̃o�[��`��
    For i = 2 To lastTaskRow
        taskID = wsTasks.Cells(i, COL_TASK_ID).Value
        taskName = wsTasks.Cells(i, COL_TASK_NAME).Value
        duration = wsTasks.Cells(i, COL_DURATION).Value
        startDate = wsTasks.Cells(i, COL_START_DATE).Value
        endDate = wsTasks.Cells(i, COL_END_DATE).Value
        progress = wsTasks.Cells(i, COL_PROGRESS).Value
        status = wsTasks.Cells(i, COL_STATUS).Value

        ' �^�X�N�o�[�̕`��
        Call DrawTaskBar(wsGantt, taskID, taskName, startDate, endDate, status, _
                         chartStartRow + i - 1, chartStartCol, colWidth, barHeight, minDate)
    Next i

    ' �S�̐i���O���t�̍X�V
    Call UpdateLoadGraph(wsGantt, wsTasks, chartStartRow, chartStartCol, colWidth, minDate, maxDate)

    Exit Sub

ErrHandler:
    MsgBox "�K���g�`���[�g�̍X�V���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' �����̃K���g�`���[�g���N���A����
Private Sub ClearGanttChart(wsGantt As Worksheet)
    On Error Resume Next ' �G���[���������Ă������𑱍s

    Dim sh As Shape
    For Each sh In wsGantt.Shapes
        If Left(sh.Name, 8) = "TaskBar_" Or Left(sh.Name, 9) = "Timeline_" Or Left(sh.Name, 10) = "Progress_" Then
            sh.Delete
        End If
    Next sh

    ' �O���t���N���A (���������)
    For Each sh In wsGantt.Shapes
        If sh.Type = msoChart Then
            sh.Delete
        End If
    Next sh

    On Error GoTo 0 ' �G���[�n���h�����O�����Z�b�g
End Sub

' 1�̃^�X�N�ɑΉ�����o�[��`�悷��
Private Sub DrawTaskBar(wsGantt As Worksheet, taskID As Long, taskName As String, _
                        startDate As Date, endDate As Date, status As String, _
                        rowNum As Long, chartStartCol As Long, colWidth As Long, barHeight As Long, _
                        minChartDate As Date)
    On Error GoTo ErrHandler

    Dim barLeft As Double
    Dim barTop As Double
    Dim barWidth As Double
    Dim barColor As Long
    Dim sh As Shape

    ' �o�[�̊J�n�ʒu�ƕ����v�Z
    barLeft = wsGantt.Cells(rowNum, 1).Left + (startDate - minChartDate) * colWidth
    barTop = wsGantt.Cells(rowNum, 1).Top + (wsGantt.Cells(rowNum, 1).Height - barHeight) / 2
    barWidth = (endDate - startDate + 1) * colWidth

    ' �X�e�[�^�X�ɉ������F���擾
    barColor = GetColorByStatus(status)

    ' �o�[��`��
    Set sh = wsGantt.Shapes.AddShape(msoShapeRectangle, barLeft, barTop, barWidth, barHeight)
    With sh
        .Fill.ForeColor.RGB = barColor
        .Line.Visible = msoFalse
        .Name = "TaskBar_" & taskID ' �^�X�NID�𖼑O�Ɋ܂߂�
        .OnAction = "M_ChartEvents.ShowTaskDetails" ' �N���b�N�C�x���g�̃}�N�������蓖��
        .TextFrame2.TextRange.Text = taskName
        With .TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0) ' �e�L�X�g�F�����ɐݒ�
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
    MsgBox "�^�X�N�o�[�̕`�撆�ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' �^�C�����C����`�悷��
Private Sub DrawTimeline(wsGantt As Worksheet, startDate As Date, endDate As Date, _
                         chartStartRow As Long, chartStartCol As Long, colWidth As Long)
    On Error GoTo ErrHandler

    Dim currentDate As Date
    Dim colOffset As Long
    Dim headerRow As Long

    headerRow = chartStartRow - 1 ' �^�C�����C���̃w�b�_�[�s

    ' ���t�w�b�_�[�̃N���A
    wsGantt.Range(wsGantt.Cells(headerRow, chartStartCol), wsGantt.Cells(headerRow, chartStartCol + (endDate - startDate + 1))).ClearContents

    colOffset = 0
    For currentDate = startDate To endDate
        wsGantt.Cells(headerRow, chartStartCol + colOffset).Value = Format(currentDate, "m/d")
        wsGantt.Cells(headerRow, chartStartCol + colOffset).ColumnWidth = colWidth / 6 ' ���t�\���ɍ��킹�Ē���
        wsGantt.Cells(headerRow, chartStartCol + colOffset).HorizontalAlignment = xlCenter
        wsGantt.Cells(headerRow, chartStartCol + colOffset).VerticalAlignment = xlCenter
        wsGantt.Cells(headerRow, chartStartCol + colOffset).Orientation = 90 ' �c����

        ' �T���̔w�i�F��ύX
        If Weekday(currentDate, vbSaturday) = vbSaturday Or Weekday(currentDate, vbSaturday) = vbSunday Then
            With wsGantt.Cells(headerRow, chartStartCol + colOffset).Interior
                .Color = RGB(220, 220, 220) ' �����D�F
            End With
        Else
            With wsGantt.Cells(headerRow, chartStartCol + colOffset).Interior
                .Pattern = xlNone
            End With
        End If

        colOffset = colOffset + 1
    Next currentDate

    Exit Sub

ErrHandler:
    MsgBox "�^�C�����C���̕`�撆�ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' �S�̐i���O���t���X�V����
Private Sub UpdateLoadGraph(wsGantt As Worksheet, wsTasks As Worksheet, _
                            chartStartRow As Long, chartStartCol As Long, colWidth As Long, _
                            minChartDate As Date, maxChartDate As Date)
    On Error GoTo ErrHandler

    Dim lastTaskRow As Long
    Dim i As Long
    Dim totalDuration As Double
    Dim completedDuration As Double
    Dim progressPercentage As Double
    Dim chartObj As ChartObject
    Dim chartName As String

    chartName = "OverallProgressChart"

    ' �����̃O���t���폜
    For Each chartObj In wsGantt.ChartObjects
        If chartObj.Name = chartName Then
            chartObj.Delete
            Exit For
        End If
    Next chartObj

    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    totalDuration = 0
    completedDuration = 0

    For i = 2 To lastTaskRow
        Dim duration As Long
        Dim progress As Double
        Dim status As String

        duration = wsTasks.Cells(i, COL_DURATION).Value ' ����
        progress = wsTasks.Cells(i, COL_PROGRESS).Value ' �i��
        status = wsTasks.Cells(i, COL_STATUS).Value   ' �X�e�[�^�X

        totalDuration = totalDuration + duration

        If status = "����" Then
            completedDuration = completedDuration + duration
        Else
            completedDuration = completedDuration + (duration * progress)
        End If
    Next i

    If totalDuration > 0 Then
        progressPercentage = completedDuration / totalDuration
    Else
        progressPercentage = 0
    End If

    ' �O���t�̃f�[�^�͈͂�ݒ� (�ꎞ�I�ɃV�[�g�ɏ����o��)
    wsGantt.Cells(1, 1).Value = "�i��"
    wsGantt.Cells(1, 2).Value = progressPercentage

    ' �O���t�̍쐬
    Set chartObj = wsGantt.ChartObjects.Add(Left:=wsGantt.Cells(chartStartRow, chartStartCol).Left, _
                                            Top:=wsGantt.Cells(chartStartRow, chartStartCol).Top + (maxChartDate - minChartDate + 2) * wsGantt.Cells(1, 1).Height, _
                                            Width:=300, Height:=150)
    With chartObj
        .Name = chartName
        With .Chart
            .ChartType = xlDoughnut
            .SetSourceData Source:=wsGantt.Range(wsGantt.Cells(1, 1), wsGantt.Cells(1, 2))
            .HasTitle = True
            .ChartTitle.Text = "�S�̐i����"
            .ChartTitle.Font.Size = 10
            .HasLegend = False
            .DoughnutHoleSize = 60

            ' �f�[�^�n��̐ݒ�
            With .SeriesCollection(1)
                .Points(1).Interior.Color = RGB(0, 176, 80) ' �������� (��)
                .Points(2).Interior.Color = RGB(200, 200, 200) ' ���������� (�D�F)
                .ApplyDataLabels
                .DataLabels.ShowPercentage = True
                .DataLabels.Font.Size = 10
                .DataLabels.Position = xlLabelPositionCenter
            End With
        End With
    End With

    Exit Sub

ErrHandler:
    MsgBox "�S�̐i���O���t�̍X�V���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' �X�e�[�^�X�ɉ������F��Ԃ��֐�
Private Function GetColorByStatus(status As String) As Long
    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    Select Case status
        Case "������"
            GetColorByStatus = wsSettings.Cells(SETTING_CHART_START_ROW + 1, SETTING_COL_WIDTH + 1).Value ' ��: Settings!G2 �ɖ�����̐F
        Case "�i�s��"
            GetColorByStatus = wsSettings.Cells(SETTING_CHART_START_ROW + 2, SETTING_COL_WIDTH + 1).Value ' ��: Settings!G3 �ɐi�s���̐F
        Case "����"
            GetColorByStatus = wsSettings.Cells(SETTING_CHART_START_ROW + 3, SETTING_COL_WIDTH + 1).Value ' ��: Settings!G4 �Ɋ����̐F
        Case "�x��"
            GetColorByStatus = wsSettings.Cells(SETTING_CHART_START_ROW + 4, SETTING_COL_WIDTH + 1).Value ' ��: Settings!G5 �ɒx���̐F
        Case Else
            GetColorByStatus = RGB(192, 192, 192) ' �f�t�H���g�F (�D�F)
    End Select
End Function
