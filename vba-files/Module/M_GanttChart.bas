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

' Settings�V�[�g�̃Z���Q�� (�s�ԍ�)
Private Const SETTING_CHART_START_ROW As Long = 1 ' �`���[�g�J�n�s/��
Private Const SETTING_BAR_HEIGHT As Long = 2      ' �o�[�̍���
Private Const SETTING_ROW_HEIGHT As Long = 3      ' �s�̍���
Private Const SETTING_COL_WIDTH As Long = 4       ' ��̕�
Private Const SETTING_COLOR_UNSTARTED As Long = 5 ' ������̐F
Private Const SETTING_COLOR_IN_PROGRESS As Long = 6 ' �i�s���̐F
Private Const SETTING_COLOR_COMPLETED As Long = 7 ' �����̐F
Private Const SETTING_COLOR_DELAYED As Long = 8   ' �x���̐F

' �K���g�`���[�g���X�V���郁�C���v���V�[�W��
Public Sub UpdateGanttChart()
    On Error GoTo ErrHandler

    Dim wsGantt As Worksheet
    Dim wsTasks As Worksheet
    Dim wsSettings As Worksheet
    Dim lastTaskRow As Long
    Dim i As Long
    Dim minDate As Date, maxDate As Date
    Dim chartStartRow As Long, chartStartCol As Long
    Dim barHeight As Double, rowHeight As Double, colWidth As Double
    
    '--- �ϐ��̏����� ---
    Dim vStartDate As Variant, vEndDate As Variant
    Dim taskID As Long, duration As Long
    Dim taskName As String, status As String
    Dim progress As Double

    Set wsGantt = ThisWorkbook.Sheets("GanttChart")
    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    ' �����̃`���[�g�Ɛ}�`���N���A (�X�V�{�^���͏���)
    Call ClearGanttChart(wsGantt)

    ' �ݒ�l�̓ǂݍ��� (Settings�V�[�g��B�񂩂�l��ǂݍ���)
    chartStartRow = wsSettings.Cells(SETTING_CHART_START_ROW, 2).Value
    chartStartCol = wsSettings.Cells(SETTING_CHART_START_ROW, 3).Value ' �J�n���C��
    barHeight = wsSettings.Cells(SETTING_BAR_HEIGHT, 2).Value
    rowHeight = wsSettings.Cells(SETTING_ROW_HEIGHT, 2).Value
    colWidth = wsSettings.Cells(SETTING_COL_WIDTH, 2).Value

    ' �^�X�N�f�[�^�̍ŏI�s���擾 (Tasks�V�[�g��B����)
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row

    If lastTaskRow < 2 Then ' �w�b�_�[�s�݂̂̏ꍇ
        MsgBox "�^�X�N�f�[�^������܂���B", vbInformation
        Exit Sub
    End If

    ' --- ���t�͈͂̓��� (�L���ȓ��t�݂̂�Ώ�) ---
    minDate = Date + 36500 ' �����̑傫�ȓ��t�ŏ�����
    maxDate = Date - 36500 ' �ߋ��̏����ȓ��t�ŏ�����
    
    For i = 2 To lastTaskRow
        vStartDate = wsTasks.Cells(i, COL_START_DATE).Value
        vEndDate = wsTasks.Cells(i, COL_END_DATE).Value
        If IsDate(vStartDate) And IsDate(vEndDate) Then
            If CDate(vStartDate) < minDate Then minDate = CDate(vStartDate)
            If CDate(vEndDate) > maxDate Then maxDate = CDate(vEndDate)
        End If
    Next i
    
    ' �L���ȃ^�X�N������Ȃ������ꍇ
    If minDate > maxDate Then
        MsgBox "�L���ȓ��t�f�[�^�����^�X�N������܂���B", vbInformation
        Exit Sub
    End If

    ' �^�C�����C���̕`��
    Call DrawTimeline(wsGantt, minDate, maxDate, chartStartRow, chartStartCol, colWidth)

    ' �e�^�X�N�̃o�[��`�悷��O�ɁA���t�f�[�^�̗L�������`�F�b�N
    For i = 2 To lastTaskRow
        vStartDate = wsTasks.Cells(i, COL_START_DATE).Value
        vEndDate = wsTasks.Cells(i, COL_END_DATE).Value

        ' ���t�f�[�^���L���ȏꍇ�̂ݕ`�揈�����s��
        If IsDate(vStartDate) And IsDate(vEndDate) Then
            ' ����ɏI�������J�n���ȍ~���`�F�b�N
            If CDate(vEndDate) >= CDate(vStartDate) Then
                ' �L���ȃf�[�^�݂̂�ϐ��Ɋi�[
                taskID = wsTasks.Cells(i, COL_TASK_ID).Value
                taskName = wsTasks.Cells(i, COL_TASK_NAME).Value
                duration = wsTasks.Cells(i, COL_DURATION).Value
                progress = wsTasks.Cells(i, COL_PROGRESS).Value
                status = wsTasks.Cells(i, COL_STATUS).Value

                ' �^�X�N�o�[�̕`��
                Call DrawTaskBar(wsGantt, taskID, taskName, CDate(vStartDate), CDate(vEndDate), status, _
                                 chartStartRow + i - 1, chartStartCol, colWidth, barHeight, minDate)
            Else
                Debug.Print "�s " & i & ": �I�������J�n�����O�̂��߃X�L�b�v"
            End If
        Else
            Debug.Print "�s " & i & ": ���t�f�[�^���s���̂��߃X�L�b�v"
        End If
    Next i

    ' �S�̐i���O���t�̍X�V
    Call UpdateLoadGraph(wsGantt, wsTasks, chartStartRow, chartStartCol, colWidth, minDate, maxDate)

    Exit Sub

ErrHandler:
    MsgBox "�K���g�`���[�g�̍X�V���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' ������ �C���ӏ� ������
' "UpdateChartButton"�Ƃ������O�̃{�^���ȊO�̐}�`�����ׂč폜����
Private Sub ClearGanttChart(wsGantt As Worksheet)
    On Error Resume Next ' �G���[���������Ă������𑱍s

    Dim sh As Shape
    ' For Each���[�v�ŃV�[�g��̑S�}�`���m�F
    For Each sh In wsGantt.Shapes
        ' �}�`�̖��O��"UpdateChartButton"�łȂ��ꍇ�Ɍ���A�폜����
        If sh.Name <> "UpdateChartButton" Then
            sh.Delete
        End If
    Next sh

    On Error GoTo 0 ' �G���[�n���h�����O�����Z�b�g
End Sub

' 1�̃^�X�N�ɑΉ�����o�[��`�悷��
Private Sub DrawTaskBar(wsGantt As Worksheet, taskID As Long, taskName As String, _
                        startDate As Date, endDate As Date, status As String, _
                        rowNum As Long, chartStartCol As Long, colWidth As Double, barHeight As Double, _
                        minChartDate As Date)
    On Error GoTo ErrHandler

    Dim barLeft As Double, barTop As Double, barWidth As Double
    Dim barColor As Long
    Dim sh As Shape

    ' �o�[�̊J�n�ʒu�ƕ����v�Z
    barLeft = wsGantt.Cells(rowNum, chartStartCol).Left + (startDate - minChartDate) * colWidth
    barTop = wsGantt.Cells(rowNum, 1).Top + (wsGantt.Cells(rowNum, 1).Height - barHeight) / 2
    barWidth = (endDate - startDate + 1) * colWidth

    ' �o�[�̕���0���傫���ꍇ�̂ݕ`��
    If barWidth <= 0 Then Exit Sub

    ' �X�e�[�^�X�ɉ������F���擾
    barColor = GetColorByStatus(status)

    ' �o�[��`��
    Set sh = wsGantt.Shapes.AddShape(msoShapeRectangle, barLeft, barTop, barWidth, barHeight)
    With sh
        .Fill.ForeColor.RGB = barColor
        .Line.Visible = msoFalse
        .Name = "TaskBar_" & taskID
        .TextFrame2.TextRange.Text = taskName
        With .TextFrame2.TextRange.Font
            .Fill.ForeColor.RGB = RGB(255, 255, 255) ' �e�L�X�g��
            .Size = 8
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
    End With

    Exit Sub

ErrHandler:
    MsgBox "�^�X�N�o�[�̕`�撆�ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' �^�C�����C����`�悷��
Private Sub DrawTimeline(wsGantt As Worksheet, startDate As Date, endDate As Date, _
                         chartStartRow As Long, chartStartCol As Long, colWidth As Double)
    On Error GoTo ErrHandler

    If chartStartRow <= 1 Then
        Err.Raise Number:=vbObjectError, Description:="�`���[�g�̊J�n�s��2�s�ڈȍ~�ɐݒ肵�Ă��������B"
    End If
    
    Dim currentDate As Date
    Dim colOffset As Long
    Dim headerRow As Long
    headerRow = chartStartRow - 1

    ' �^�C�����C���͈͂̏����ݒ�
    With wsGantt.Range(wsGantt.Cells(headerRow, chartStartCol), wsGantt.Cells(wsGantt.Rows.Count, chartStartCol + (endDate - startDate + 2)))
        .Clear
        .ColumnWidth = colWidth / 7
        .HorizontalAlignment = xlCenter
    End With

    colOffset = 0
    For currentDate = startDate To endDate
        With wsGantt.Cells(headerRow, chartStartCol + colOffset)
            .Value = Format(currentDate, "m/d")
            ' �T���̔w�i�F��ύX
            If Weekday(currentDate) = vbSaturday Or Weekday(currentDate) = vbSunday Then
                .Interior.Color = RGB(240, 240, 240)
            End If
        End With
        colOffset = colOffset + 1
    Next currentDate

    Exit Sub

ErrHandler:
    MsgBox "�^�C�����C���̕`�撆�ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' �S�̐i���O���t���X�V����
Private Sub UpdateLoadGraph(wsGantt As Worksheet, wsTasks As Worksheet, _
                            chartStartRow As Long, chartStartCol As Long, colWidth As Double, _
                            minChartDate As Date, maxChartDate As Date)
    On Error GoTo ErrHandler

    Dim lastTaskRow As Long
    Dim i As Long
    Dim totalDuration As Double
    Dim completedDuration As Double
    Dim progressPercentage As Double
    Dim chartObj As ChartObject
    Dim chartName As String
    Dim vProgress As Variant

    chartName = "OverallProgressChart"
    
    ' �O�̂��ߊ����̓����O���t���폜
    On Error Resume Next
    wsGantt.ChartObjects(chartName).Delete
    On Error GoTo ErrHandler

    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    totalDuration = 0
    completedDuration = 0

    For i = 2 To lastTaskRow
        If IsNumeric(wsTasks.Cells(i, COL_DURATION).Value) And IsNumeric(wsTasks.Cells(i, COL_PROGRESS).Value) Then
            totalDuration = totalDuration + wsTasks.Cells(i, COL_DURATION).Value
            completedDuration = completedDuration + (wsTasks.Cells(i, COL_DURATION).Value * wsTasks.Cells(i, COL_PROGRESS).Value)
        End If
    Next i

    If totalDuration > 0 Then
        progressPercentage = completedDuration / totalDuration
    Else
        progressPercentage = 0
    End If

    ' �O���t�̃f�[�^�Ƃ��āu�i�����v�Ɓu�c��v��2�̒l���ꎞ�Z���ɏ����o��
    wsGantt.Range("A1").Value = progressPercentage
    wsGantt.Range("B1").Value = 1 - progressPercentage

    ' �O���t�̍쐬
    Set chartObj = wsGantt.ChartObjects.Add(Left:=wsGantt.Cells(chartStartRow, chartStartCol).Left, _
                                            Top:=wsGantt.Cells(lastTaskRow + 2, 1).Top, _
                                            Width:=200, Height:=120)
    With chartObj
        .Name = chartName
        With .Chart
            .ChartType = xlDoughnut
            .SetSourceData Source:=wsGantt.Range("A1:B1")
            .HasTitle = True
            .ChartTitle.Text = "�S�̐i����"
            .ChartTitle.Font.Size = 10
            .HasLegend = False
            .ChartGroups(1).DoughnutHoleSize = 60

            With .SeriesCollection(1)
                .Points(1).Interior.Color = RGB(0, 176, 80)    ' �������� (��)
                .Points(2).Interior.Color = RGB(220, 220, 220) ' ���������� (�D�F)
                
                ' �i�����𒆉��ɕ\��
                .ApplyDataLabels
                With .DataLabels(1)
                    .ShowValue = True
                    .ShowCategoryName = False
                    .ShowSeriesName = False
                    .ShowPercentage = False
                    .NumberFormat = "0%"
                    .Font.Size = 12
                    .Font.Bold = True
                    .Position = xlLabelPositionCenter
                End With
            End With
        End With
    End With
    
    ' �ꎞ�I�Ɏg�p�����Z�����N���A
    wsGantt.Range("A1:B1").ClearContents

    Exit Sub

ErrHandler:
    MsgBox "�S�̐i���O���t�̍X�V���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' �X�e�[�^�X�ɉ������F��Ԃ��֐�
Private Function GetColorByStatus(status As String) As Long
    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    ' Settings�V�[�g��B��ɂ���Z���́u�w�i�F�v���擾
    Const VALUE_COL As Long = 2
    
    On Error Resume Next
    Select Case status
        Case "������"
            GetColorByStatus = wsSettings.Cells(SETTING_COLOR_UNSTARTED, VALUE_COL).Interior.Color
        Case "�i�s��"
            GetColorByStatus = wsSettings.Cells(SETTING_COLOR_IN_PROGRESS, VALUE_COL).Interior.Color
        Case "����"
            GetColorByStatus = wsSettings.Cells(SETTING_COLOR_COMPLETED, VALUE_COL).Interior.Color
        Case "�x��"
            GetColorByStatus = wsSettings.Cells(SETTING_COLOR_DELAYED, VALUE_COL).Interior.Color
        Case Else
            GetColorByStatus = RGB(192, 192, 192) ' �f�t�H���g�F (�D�F)
    End Select
    
    If Err.Number <> 0 Then
        GetColorByStatus = RGB(192, 192, 192)
        Err.Clear
    End If
    On Error GoTo 0
End Function

