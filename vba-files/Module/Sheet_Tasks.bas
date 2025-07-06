Attribute VB_Name = "Sheet_Tasks"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// Sheet_Tasks �V�[�g���W���[��
'// �uTasks�v�V�[�g�̃C�x���g�i��: Worksheet_Change�j���������A�f�[�^�ύX���ɃK���g�`���[�g�̍X�V���g���K�[���܂��B
'////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler

    Dim dataRange As Range
    Set dataRange = Me.Range("B:G") ' B�񂩂�G��܂ł��Ď��ΏۂƂ���

    ' �ύX���ꂽ�Z�����Ď��Ώ۔͈͓��ɂ��邩�A�������Z�����ύX����Ă��Ȃ������m�F
    If Not Intersect(Target, dataRange) Is Nothing And Target.Cells.Count = 1 Then
        ' �ύX���ꂽ�s���f�[�^�s�i�w�b�_�[�s�ȊO�j�ł��邩���m�F
        If Target.Row >= 2 Then
            ' �I���� (E��) ���ύX���ꂽ�ꍇ�́A���� (C��) ���Čv�Z
            If Target.Column = 5 Then ' E�� (�I����)
                Me.Cells(Target.Row, R1C3).Value = Me.Cells(Target.Row, R1C5).Value - Me.Cells(Target.Row, R1C4).Value + 1
            End If

            ' ���� (C��) �܂��͊J�n�� (D��) ���ύX���ꂽ�ꍇ�́A�I���� (E��) ���Čv�Z
            If Target.Column = 3 Or Target.Column = 4 Then ' C�� (����) �܂��� D�� (�J�n��)
                Me.Cells(Target.Row, R1C5).Value = Me.Cells(Target.Row, R1C4).Value + Me.Cells(Target.Row, R1C3).Value - 1
            End If

            ' �K���g�`���[�g���X�V
            Call M_GanttChart.UpdateGanttChart
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox "Tasks�V�[�g�̕ύX�C�x���g���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub
