Attribute VB_Name = "M_ChartEvents"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// M_ChartEvents ���W���[��
'// �K���g�`���[�g��̃^�X�N�o�[�iShape�I�u�W�F�N�g�j�N���b�N�C�x���g���������܂��B
'////////////////////////////////////////////////////////////////////////////////////////////////////

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
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, 2).End(xlUp).Row
    taskFound = False

    ' Tasks�V�[�g����Y���^�X�N�̏�������
    For i = 2 To lastTaskRow
        If wsTasks.Cells(i, 1).Value = taskID Then
            msg = "�^�X�NID: " & wsTasks.Cells(i, 1).Value & vbCrLf & _
                  "�^�X�N��: " & wsTasks.Cells(i, 2).Value & vbCrLf & _
                  "����: " & wsTasks.Cells(i, 3).Value & "��" & vbCrLf & _
                  "�J�n��: " & Format(wsTasks.Cells(i, 4).Value, "yyyy/mm/dd") & vbCrLf & _
                  "�I����: " & Format(wsTasks.Cells(i, 5).Value, "yyyy/mm/dd") & vbCrLf & _
                  "�i��: " & Format(wsTasks.Cells(i, 6).Value, "0%") & vbCrLf & _
                  "�X�e�[�^�X: " & wsTasks.Cells(i, 7).Value
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