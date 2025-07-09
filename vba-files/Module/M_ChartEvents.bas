Attribute VB_Name = "M_ChartEvents"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// M_ChartEvents ���W���[��
'// �K���g�`���[�g��̃^�X�N�o�[�iShape�I�u�W�F�N�g�j�N���b�N�C�x���g���������܂��B
'////////////////////////////////////////////////////////////////////////////////////////////////////

' Tasks�V�[�g�̗�C���f�b�N�X
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME As Long = 2
Private Const COL_DURATION As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_END_DATE As Long = 5
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

' �N���b�N���ꂽ�^�X�N�̏ڍׂ�\������
Public Sub ShowTaskDetails()
    On Error GoTo ErrHandler

    Dim clickedShape As Shape
    Dim taskID As Long
    Dim wsTasks As Worksheet
    Dim lastTaskRow As Long
    Dim i As Long
    Dim taskFound As Boolean
    Dim msg As String

    ' �N���b�N���ꂽShape�I�u�W�F�N�g���擾
    Set clickedShape = ActiveSheet.Shapes(Application.Caller)

    ' Shape�̖��O����^�X�NID�𒊏o
    If Left(clickedShape.Name, 8) = "TaskBar_" Then
        taskID = CLng(Mid(clickedShape.Name, 9))
    Else
        Exit Sub ' �^�X�N�o�[�ȊO��Shape���N���b�N���ꂽ�ꍇ�͉������Ȃ�
    End If

    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    taskFound = False

    ' Tasks�V�[�g����Y���^�X�N�̏�������
    For i = 2 To lastTaskRow
        If wsTasks.Cells(i, COL_TASK_ID).value = taskID Then
            msg = "�^�X�NID: " & wsTasks.Cells(i, COL_TASK_ID).value & vbCrLf & _
                  "�^�X�N��: " & wsTasks.Cells(i, COL_TASK_NAME).value & vbCrLf & _
                  "����: " & wsTasks.Cells(i, COL_DURATION).value & "��" & vbCrLf & _
                  "�J�n��: " & Format(wsTasks.Cells(i, COL_START_DATE).value, "yyyy/mm/dd") & vbCrLf & _
                  "�I����: " & Format(wsTasks.Cells(i, COL_END_DATE).value, "yyyy/mm/dd") & vbCrLf & _
                  "�i��: " & Format(wsTasks.Cells(i, COL_PROGRESS).value, "0%") & vbCrLf & _
                  "�X�e�[�^�X: " & wsTasks.Cells(i, COL_STATUS).value
            taskFound = True
            Exit For
        End If
    Next i

    If taskFound Then
        MsgBox msg, vbInformation, "�^�X�N�ڍ�"
    Else
        MsgBox "�^�X�N��񂪌�����܂���ł����B", vbExclamation
    End If

    Exit Sub

ErrHandler:
    MsgBox "�^�X�N�ڍׂ̕\�����ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

