Attribute VB_Name = "Sheet_GanttChart"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// Sheet_GanttChart �V�[�g���W���[��
'// �uGanttChart�v�V�[�g�ɌŗL�̃C�x���g�������L�q���܂��B
'////////////////////////////////////////////////////////////////////////////////////////////////////

'--- Tasks�V�[�g�̗�C���f�b�N�X ---
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME As Long = 2
Private Const COL_DURATION As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_END_DATE As Long = 5
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

'--- GanttChart�V�[�g�̌Œ�s�E�� ---
Private Const GANTT_START_ROW As Long = 5 ' �K���g�`���[�g�̊J�n�s
Private Const GANTT_START_COL As Long = 2 ' �K���g�`���[�g�̊J�n�� (�^�X�N���\���G���A)
Private Const TIMELINE_ROW As Long = 4    ' �^�C�����C���̕\���s

'/**
' * @brief "�X�V"�{�^�����N���b�N���ꂽ�Ƃ��ɌĂяo����܂��B
' */
Private Sub UpdateChartButton_Click()
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Call M_GanttChart.UpdateGanttChart
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "�`���[�g�X�V�{�^���̏������ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

'/**
' * @brief ���[�N�V�[�g�̑I��͈͂��ύX���ꂽ�Ƃ��ɔ�������C�x���g�ł��B
' *        �K���g�`���[�g�̃^�X�N�o�[�i���F���ꂽ�Z���j���I�����ꂽ�ꍇ�A
' *        ���̃^�X�N�̏ڍ׏���MsgBox�ŕ\�����܂��B
' * @param Target �I�����ꂽ�Z���͈�
' */
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo ErrHandler

    Dim selectedRow As Long
    Dim selectedCol As Long
    Dim taskRow As Long
    Dim wsTasks As Worksheet
    Dim msg As String

    ' --- �I�����ꂽ�Z�����P��Z���łȂ��ꍇ�͏����𔲂��� ---
    If Target.Cells.CountLarge > 1 Then Exit Sub

    selectedRow = Target.Row
    selectedCol = Target.Column

    ' --- �I�����ꂽ�Z�����K���g�`���[�g�̃^�X�N�`��G���A�������� ---
    If selectedRow >= GANTT_START_ROW And selectedCol > GANTT_START_COL Then
        
        ' --- �I�����ꂽ�Z���ɒ��F�����邩�i�^�X�N�o�[���j���� ---
        If Target.Interior.Color <> xlNone Then
            
            ' --- �Ή�����^�X�N�����擾 ---
            taskRow = selectedRow - GANTT_START_ROW + 2 ' Tasks�V�[�g�̍s�ԍ��ɕϊ�
            Set wsTasks = ThisWorkbook.Sheets("Tasks")
            
            ' --- �^�X�N��񂪑��݂��邩�m�F ---
            If wsTasks.Cells(taskRow, COL_TASK_NAME).Value <> "" Then
                
                ' --- ���b�Z�[�W�{�b�N�X�ŏڍ׏���\�� ---
                msg = "�� �^�X�N�ڍ�" & vbCrLf & vbCrLf & _
                      "�^�X�N��: " & wsTasks.Cells(taskRow, COL_TASK_NAME).Value & vbCrLf & _
                      "�S����: " & "(������)" & vbCrLf & _
                      "����: " & Format(wsTasks.Cells(taskRow, COL_START_DATE).Value, "yyyy/m/d") & " - " & Format(wsTasks.Cells(taskRow, COL_END_DATE).Value, "yyyy/m/d") & " (" & wsTasks.Cells(taskRow, COL_DURATION).Value & "����)" & vbCrLf & _
                      "�i����: " & Format(wsTasks.Cells(taskRow, COL_PROGRESS).Value, "0%") & vbCrLf & _
                      "�X�e�[�^�X: " & wsTasks.Cells(taskRow, COL_STATUS).Value
                      
                MsgBox msg, vbInformation, "�^�X�N�ڍ�"
            End If
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox "SelectionChange�C�x���g�̏������ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub
