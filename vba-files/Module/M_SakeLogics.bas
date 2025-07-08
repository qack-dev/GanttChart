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
' @brief フォームの入力値から現在の重量を取得することを試みる
' @param frm フォームオブジェクト
' @param fullWeight 満タン時の重量
' @param emptyWeight 空瓶の重量
' @param previousWeight 前回の重量（基準重量）
' @param outCurrentWeight [out] 計算された現在の重量
' @param outErrorMessage [out] エラーメッセージ
' @return Boolean 成功した場合は True、それ以外の場合は False
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

    ' 入力数の検証
    If inputCount = 0 Then
        outErrorMessage = "現在の重量、残量(%)、または杯数を入力してください。"
        TryGetCurrentWeightFromInput = False
        Exit Function
    End If
    If inputCount > 1 Then
        outErrorMessage = "入力箇所は1つだけにしてください。"
        TryGetCurrentWeightFromInput = False
        Exit Function
    End If

    ' 特殊条件の検証
    If frm.txtNowNum.Value <> "" And frm.optContinued.Value = True Then
        outErrorMessage = "杯数での入力は、継続記録と同時に使用できません。"
        TryGetCurrentWeightFromInput = False
        Exit Function
    End If

    ' 現在の重量を計算
    Dim netWeight As Double
    netWeight = fullWeight - emptyWeight

    If frm.txtNowWeight.Value <> "" Then
        outCurrentWeight = CDbl(frm.txtNowWeight.Value)
    ElseIf frm.txtNowPercent.Value <> "" Then
        outCurrentWeight = emptyWeight + (netWeight * (CDbl(frm.txtNowPercent.Value) / 100))
    ElseIf frm.txtNowNum.Value <> "" Then
        ' 開発者メモ:
        ' この計算式は「内容量全体に対する割合」として杯数を解釈しています。
        ' 例えば、内容量720mlの瓶に対して「1杯」を入力すると、720ml飲んだと解釈されます。
        ' 一般的な「1杯 = 180ml」のような固定量ではないため、ユーザーの直感と異なる可能性があります。
        ' 将来的な改善として、設定で1杯あたりの量を定義できるようにするなどの検討が考えられます。
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
' @brief 日本酒のアルコール情報を計算する
' @param frm SakeLoggerフォーム
' @param currentWeight 現在の重量
' @param previousWeight 前回の重量
' @param fullWeight 満タン時の重量
' @param emptyWeight 空瓶の重量
' @param alcoholContent アルコール度数
' @return Boolean 成功した場合は True、それ以外の場合は False
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

    ' --- 現在の残量を計算 ---
    Dim currentAmount As Double
    currentAmount = currentWeight - emptyWeight
    frm.lblNowAmount.Caption = Round(currentAmount, 2) & " g"

    ' --- 今回の飲量を計算 ---
    Dim drunkWeight As Double
    drunkWeight = previousWeight - currentWeight
    frm.lblDrunkWeight.Caption = Round(drunkWeight, 2) & " g"

    ' --- アルコール摂取量を計算 ---
    Dim alcoholVolume As Double
    alcoholVolume = (drunkWeight / DENSITY_SAKE) * (alcoholContent / 100)
    frm.lblDrunkAlcohol.Caption = Round(alcoholVolume, 2) & " ml"

    ' --- 残量をパーセントで表示 ---
    Dim remainingPercent As Double
    remainingPercent = (currentAmount / netWeight) * 100
    frm.lblNowPercent.Caption = Round(remainingPercent, 1) & " %"

    CalculateAlcoholInfo = True
    Exit Function

ErrorHandler:
    CalculateAlcoholInfo = False
End Function
