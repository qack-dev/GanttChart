Attribute VB_Name = "Sheet_GanttChart"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// Sheet_GanttChart �V�[�g���W���[��
'// �uGanttChart�v�V�[�g�ɔz�u���ꂽ�{�^���̃A�N�V�����Ȃǂ��L�q���܂��B
'////////////////////////////////////////////////////////////////////////////////////////////////////

' ���̃V�[�g�Ƀ{�^����z�u���A���̃}�N����o�^���Ă��������B
' �{�^�����N���b�N���ꂽ�Ƃ��ɃK���g�`���[�g���X�V���܂��B
Public Sub UpdateChartButton_Click()
    On Error GoTo ErrHandler
    Call M_GanttChart.UpdateGanttChart
    Exit Sub
ErrHandler:
    MsgBox "�`���[�g�X�V�{�^���N���b�N���ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub
