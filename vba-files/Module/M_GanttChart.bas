Attribute VB_Name = "M_GanttChart"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// M_GanttChart ���W���[��
'// �K���g�`���[�g�̕`��A�X�V�A���׃O���t�̍X�V�ȂǁA��v�ȃ��W�b�N���i�[���܂��B
'// �]����Shape�I�u�W�F�N�g�ɂ��`���������A�Z���̔w�i�F�𒅐F��������ɑS�ʓI�ɉ��C�B
'////////////////////////////////////////////////////////////////////////////////////////////////////

'--- �F��` ---
' �e�^�X�N�X�e�[�^�X�̔w�i�F���`���܂��B
Public Const COLOR_UNSTARTED As Long = &HC0C0C0 ' ������ (�D�F)
Public Const COLOR_IN_PROGRESS As Long = &HFFFF00 ' �i�s�� (���F)
Public Const COLOR_COMPLETED As Long = &H92D050 ' ���� (�ΐF)
Public Const COLOR_DELAYED As Long = &H0000FF ' �x�� (�ԐF)

'--- Tasks�V�[�g�̗�C���f�b�N�X ---
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME As Long = 2
Private Const COL_START_DATE As Long = 4
Private Const COL_DURATION As Long = 3
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

'--- GanttChart�V�[�g�̌Œ�s�E�� ---
Private Const GANTT_START_ROW As Long = 5 ' �K���g�`���[�g�̊J�n�s
Private Const GANTT_START_COL As Long = 2 ' �K���g�`���[�g�̊J�n�� (�^�X�N���\���G���A)
Private Const TIMELINE_ROW As Long = 4    ' �^�C�����C���̕\���s

'====================================================================================================
'// Public Procedures
'====================================================================================================

'/**
' * @brief ���C���v���V�[�W���B�K���g�`���[�g�S�̂��X�V���܂��B
' */
Public Sub UpdateGanttChart()
    On Error GoTo ErrHandler

    Dim wsGantt As Worksheet
    Dim wsTasks As Worksheet
    Dim lastTaskRow As Long
    Dim i As Long
    Dim minDate As Date
    Dim maxDate As Date

    Set wsGantt = ThisWorkbook.Sheets("GanttChart")
    Set wsTasks = ThisWorkbook.Sheets("Tasks")

    ' --- �f�[�^�̗L�����`�F�b�N ---
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    If lastTaskRow < 2 Then
        MsgBox "�^�X�N�����͂���Ă��܂���B", vbInformation
        Exit Sub
    End If

    ' --- �`���[�g�`��G���A�̃N���A ---
    Call ClearGanttArea(wsGantt, lastTaskRow)

    ' --- �^�C���X�P�[���̌��� ---
    ' �^�X�N���X�g����ŏ��J�n���ƍő�I�������v�Z
    With Application.WorksheetFunction
        minDate = .Min(wsTasks.Range("D2:D" & lastTaskRow))
        maxDate = .Max(wsTasks.Range("E2:E" & lastTaskRow))
    End With

    ' --- �`�揈�� ---
    Call DrawTimeline(wsGantt, minDate, maxDate)
    Call DrawAllTaskBars(wsGantt, wsTasks, lastTaskRow, minDate)
    ' Call UpdateLoadGraph(wsGantt, wsTasks, minDate, maxDate) ' �K�v�ɉ����ăR�����g����

    Exit Sub

ErrHandler:
    MsgBox "�K���g�`���[�g�̍X�V���ɃG���[���������܂���: " & vbCrLf & Err.Description, vbCritical
End Sub

'====================================================================================================
'// Private Procedures
'====================================================================================================

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

    Exit Sub

ErrHandler:
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