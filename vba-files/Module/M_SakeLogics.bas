Attribute VB_Name = "M_SakeLogics"
Option Explicit

' SakeLogger Form (frmSakeLogger)
' @see https://github.com/tack3/SakeLogger
'
' Copyright (c) 2024 tack3
' Released under the MIT license
' https://github.com/tack3/SakeLogger/blob/main/LICENSE

' =================================================================================================
' ### Private Functions
' =================================================================================================

'''
' @brief �t�H�[���̓��͒l���猻�݂̏d�ʂ��擾���邱�Ƃ����݂�
' @param frm �t�H�[���I�u�W�F�N�g
' @param fullWeight ���^�����̏d��
' @param emptyWeight ��r�̏d��
' @param previousWeight �O��̏d�ʁi��d�ʁj
' @param outCurrentWeight [out] �v�Z���ꂽ���݂̏d��
' @param outErrorMessage [out] �G���[���b�Z�[�W
' @return Boolean ���������ꍇ�� True�A����ȊO�̏ꍇ�� False
'''
Private Function TryGetCurrentWeightFromInput( _
    ByVal frm As frmSakeLogger, _
    ByVal fullWeight As Double, _
    ByVal emptyWeight As Double, _
    ByVal previousWeight As Double, _
    ByRef outCurrentWeight As Double, _
    ByRef outErrorMessage As String _
) As Boolean
    Dim inputCount As Long
    inputCount = 0
    If frm.txtNowWeight.Value <> "" Then inputCount = inputCount + 1
    If frm.txtNowPercent.Value <> "" Then inputCount = inputCount + 1
    If frm.txtNowNum.Value <> "" Then inputCount = inputCount + 1

    ' ���͐��̌���
    If inputCount = 0 Then
        outErrorMessage = "���݂̏d�ʁA�c��(%)�A�܂��͔t������͂��Ă��������B"
        TryGetCurrentWeightFromInput = False
        Exit Function
    End If
    If inputCount > 1 Then
        outErrorMessage = "���͉ӏ���1�����ɂ��Ă��������B"
        TryGetCurrentWeightFromInput = False
        Exit Function
    End If

    ' ��������̌���
    If frm.txtNowNum.Value <> "" And frm.optContinued.Value = True Then
        outErrorMessage = "�t���ł̓��͂́A�p���L�^�Ɠ����Ɏg�p�ł��܂���B"
        TryGetCurrentWeightFromInput = False
        Exit Function
    End If

    ' ���݂̏d�ʂ��v�Z
    Dim netWeight As Double
    netWeight = fullWeight - emptyWeight

    If frm.txtNowWeight.Value <> "" Then
        outCurrentWeight = CDbl(frm.txtNowWeight.Value)
    ElseIf frm.txtNowPercent.Value <> "" Then
        outCurrentWeight = emptyWeight + (netWeight * (CDbl(frm.txtNowPercent.Value) / 100))
    ElseIf frm.txtNowNum.Value <> "" Then
        ' �J���҃���:
        ' ���̌v�Z���́u���e�ʑS�̂ɑ΂��銄���v�Ƃ��Ĕt�������߂��Ă��܂��B
        ' �Ⴆ�΁A���e��720ml�̕r�ɑ΂��āu1�t�v����͂���ƁA720ml���񂾂Ɖ��߂���܂��B
        ' ��ʓI�ȁu1�t = 180ml�v�̂悤�ȌŒ�ʂł͂Ȃ����߁A���[�U�[�̒����ƈقȂ�\��������܂��B
        ' �����I�ȉ��P�Ƃ��āA�ݒ��1�t������̗ʂ��`�ł���悤�ɂ���Ȃǂ̌������l�����܂��B
        Dim drunkAmount As Double
        drunkAmount = netWeight * CDbl(frm.txtNowNum.Value)
        outCurrentWeight = previousWeight - drunkAmount
    End If

    TryGetCurrentWeightFromInput = True
End Function


' =================================================================================================
' ### Public Functions
' =================================================================================================

'''
' @brief ���{���̃A���R�[�������v�Z����
' @param frm SakeLogger�t�H�[��
' @param currentWeight ���݂̏d��
' @param previousWeight �O��̏d��
' @param fullWeight ���^�����̏d��
' @param emptyWeight ��r�̏d��
' @param alcoholContent �A���R�[���x��
' @return Boolean ���������ꍇ�� True�A����ȊO�̏ꍇ�� False
'''
Public Function CalculateAlcoholInfo( _
    ByVal frm As frmSakeLogger, _
    ByVal currentWeight As Double, _
    ByVal previousWeight As Double, _
    ByVal fullWeight As Double, _
    ByVal emptyWeight As Double, _
    ByVal alcoholContent As Double _
) As Boolean
    On Error GoTo ErrorHandler

    Dim netWeight As Double
    netWeight = fullWeight - emptyWeight
    If netWeight <= 0 Then GoTo ErrorHandler

    ' --- ���݂̎c�ʂ��v�Z ---
    Dim currentAmount As Double
    currentAmount = currentWeight - emptyWeight
    frm.lblNowAmount.Caption = Round(currentAmount, 2) & " g"

    ' --- ����̈��ʂ��v�Z ---
    Dim drunkWeight As Double
    drunkWeight = previousWeight - currentWeight
    frm.lblDrunkWeight.Caption = Round(drunkWeight, 2) & " g"

    ' --- �A���R�[���ێ�ʂ��v�Z ---
    Dim alcoholVolume As Double
    alcoholVolume = (drunkWeight / DENSITY_SAKE) * (alcoholContent / 100)
    frm.lblDrunkAlcohol.Caption = Round(alcoholVolume, 2) & " ml"

    ' --- �c�ʂ��p�[�Z���g�ŕ\�� ---
    Dim remainingPercent As Double
    remainingPercent = (currentAmount / netWeight) * 100
    frm.lblNowPercent.Caption = Round(remainingPercent, 1) & " %"

    CalculateAlcoholInfo = True
    Exit Function

ErrorHandler:
    CalculateAlcoholInfo = False
End Function
