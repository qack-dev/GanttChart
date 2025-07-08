VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSakeLogger 
   Caption         =   "飲酒記録フォーム"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   OleObjectBlob   =   "frmSakeLogger.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSakeLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub UserForm_Initialize()
    Dim masterSheet As Worksheet
    Set masterSheet = M_SakeForm.GetMasterSheet()
    
    Dim lastRow As Long
    lastRow = M_SakeForm.GetLastMasterDataCell().Row
    
    If lastRow >= 2 Then
        Dim i As Long
        For i = 2 To lastRow
            cmbSake.AddItem masterSheet.Cells(i, COL_MASTER_ID).Value & "." & masterSheet.Cells(i, COL_MASTER_NAME).Value
        Next i
    End If
    
    txtDate.Value = Format(Date, "yyyy/mm/dd")
End Sub

Private Sub cmbSake_Change()
    Dim masterSheet As Worksheet
    Set masterSheet = M_SakeForm.GetMasterSheet()
    
    Dim lastRow As Long
    lastRow = M_SakeForm.GetLastMasterDataCell().Row
    
    Dim i As Long
    For i = 2 To lastRow
        If masterSheet.Cells(i, COL_MASTER_ID).Value & "." & masterSheet.Cells(i, COL_MASTER_NAME).Value = cmbSake.Value Then
            lblABV.Caption = "度数: " & masterSheet.Cells(i, COL_MASTER_ALCOHOL).Value & " %"
            lblFullW.Caption = "満タン重量: " & masterSheet.Cells(i, COL_MASTER_FULL_WEIGHT).Value & " g"
            
            If IsEmpty(masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value) Then
                lblEmptyW.Caption = "空ボトル重量: 未登録"
                lblAlert.Caption = "!!! このお酒は空ボトル重量が未登録です !!!"
            Else
                lblEmptyW.Caption = "空ボトル重量: " & masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value & " g"
                lblAlert.Caption = ""
            End If
            
            Exit For
        End If
    Next i
End Sub

Private Sub btnCalc_Click()
    Dim sakeName As String
    Dim currentWeight As Double, drankWeight As Double, pureAlcohol As Double

    sakeName = cmbSake.Value

    If sakeName = "" Then
        MsgBox "お酒を選択してください。", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtNowWeight.Value) Then
        MsgBox "現在の重量を半角数字で入力してください。", vbExclamation
        Exit Sub
    End If

    currentWeight = CDbl(txtNowWeight.Value)

    If M_SakeLogics.CalculateAlcoholInfo(sakeName, currentWeight, drankWeight, pureAlcohol) Then
        lblResult.Caption = "純アルコール量: " & Format(pureAlcohol, "0.0") & " g"
    End If
End Sub

Private Sub btnSave_Click()
    Dim sakeName As String
    Dim drankWeight As Double, pureAlcohol As Double
    Dim currentWeight As Double
    Dim lastLogRow As Long
    Dim logSheet As Worksheet

    ' --- 入力チェック ---
    If cmbSake.Value = "" Then
        MsgBox "お酒を選択してください。", vbExclamation
        Exit Sub
    End If

    If Not M_SakeLogics.IsValidDateFormat(txtDate.Value) Then
        MsgBox "日付は'yyyy/mm/dd'形式で入力してください。", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtNowWeight.Value) Then
        MsgBox "現在の重量を入力してください。", vbExclamation
        Exit Sub
    End If

    sakeName = cmbSake.Value
    currentWeight = CDbl(txtNowWeight.Value)

    If Not M_SakeLogics.CalculateAlcoholInfo(sakeName, currentWeight, drankWeight, pureAlcohol) Then
        Exit Sub
    End If

    Set logSheet = M_SakeForm.GetLogSheet()
    lastLogRow = logSheet.Cells(logSheet.Rows.Count, COL_LOG_ID).End(xlUp).Row + 1

    ' --- 書式設定 ---
    logSheet.Cells(lastLogRow, COL_LOG_DATE).NumberFormat = "yyyy/mm/dd"
    logSheet.Cells(lastLogRow, COL_LOG_CURRENT_WEIGHT).NumberFormat = "0.0"
    logSheet.Cells(lastLogRow, COL_LOG_PURE_ALCOHOL).NumberFormat = "0.0"
    logSheet.Cells(lastLogRow, COL_LOG_DRANK_WEIGHT).NumberFormat = "0.0"

    ' --- ログに記録 ---
    logSheet.Cells(lastLogRow, COL_LOG_DATE).Value = CDate(txtDate.Value)
    logSheet.Cells(lastLogRow, COL_LOG_NAME).Value = sakeName
    logSheet.Cells(lastLogRow, COL_LOG_CURRENT_WEIGHT).Value = currentWeight
    logSheet.Cells(lastLogRow, COL_LOG_PURE_ALCOHOL).Value = Round(pureAlcohol, 1)
    logSheet.Cells(lastLogRow, COL_LOG_DRANK_WEIGHT).Value = Round(drankWeight, 1)
    logSheet.Cells(lastLogRow, COL_LOG_ID).Value = lastLogRow - 1
    
    MsgBox "記録を保存しました。", vbInformation

    ' --- 入力欄をリセット ---
    txtNowWeight.Value = ""
    lblResult.Caption = ""
End Sub
