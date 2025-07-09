Attribute VB_Name = "M_Dlaw"
Option Explicit

' Tasks�V�[�g�̗�C���f�b�N�X
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME = 2
Private Const COL_DURATION As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_END_DATE As Long = 5
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

' ID����
Public Sub DrawLine()
    ' �C�x���g�̔������ꎞ�I�ɖ����ɂ���
    Application.EnableEvents = False
    ' �G���[���������Ă��K���C�x���g��L���ɖ߂����߂̃G���[�n���h�����O
    On Error GoTo ErrHandler
    
    ' �ϐ��錾
    Dim i As Integer
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim lastRaw As Integer
    Dim dt As Date
    ' ���
    Set ws = ThisWorkbook.Worksheets("Tasks")
    lastRaw = ws.Cells(ws.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    Set dataRange = ws.Range( _
        ws.Cells(1, COL_TASK_ID), _
        ws.Cells(lastRaw, COL_STATUS) _
    )
    ws.Cells.Borders.LineStyle = xlNone
    ' ���͂����Z�b�g
    ' ����---------------------------------------
    For i = 2 To lastRaw
        ws.Cells(i, COL_TASK_ID).Value = i - 1
        dt = ws.Cells(i, COL_START_DATE).Value
        dt = dt + ws.Cells(i, COL_DURATION).Value - 1
        ws.Cells(i, COL_END_DATE).NumberFormat = "yyyy/mm/dd"
        ws.Cells(i, COL_END_DATE).Value = dt
    Next i
    '--------------------------------------------
    ' ���o���s�Ɍr��������
    ws.Range(ws.Cells(1, COL_TASK_ID), ws.Cells(1, COL_STATUS)).Borders.LineStyle = xlContinuous
    ' �t�H�[�}�b�g������
    With dataRange
        ' ����������
        .EntireColumn.AutoFit
        ' �r��������
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    
ExitProc:     ' ����I���A�܂��̓G���[������̃W�����v��
    ' �J��
    Set ws = Nothing
    Set dataRange = Nothing
    ' �C�x���g�̔������ēx�L���ɂ���
    Application.EnableEvents = True
    Exit Sub

ErrHandler: ' �G���[�������̃W�����v��
    MsgBox "DrawLine�v���V�[�W���ŃG���[���������܂���: " & Err.Description, vbCritical
    ' �G���[���������Ă��K���C�x���g��L���ɖ߂����߁AExitProc�ɃW�����v
    Resume ExitProc
End Sub
