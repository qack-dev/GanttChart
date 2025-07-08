Attribute VB_Name = "M_SakeForm"
Option Explicit

Private masterSheet As Worksheet
Private logSheet As Worksheet
Private lastMasterDataCell As Range

' ユーザーフォームを表示するメインプロシージャ
Public Sub ShowSakeLoggerForm()
    If Not InitializeObjects() Then Exit Sub
    
    ' フォーム表示前にマスタシートをアクティブ化
    masterSheet.Activate
    
    Load frmSakeLogger
    frmSakeLogger.Show
    Unload frmSakeLogger
    
    ' フォームが閉じた後、ログシートをアクティブ化し、整形
    logSheet.Activate
    Dim lastLogCell As Range
    Set lastLogCell = logSheet.Cells(logSheet.Rows.Count, COL_LOG_ID).End(xlUp)
    If lastLogCell.Row > 1 Then
        Call M_SheetUtils.FormatTable(logSheet.Range(logSheet.Cells(1, COL_LOG_ID), lastLogCell.Offset(0, COL_LOG_COMMENT - 1)), True)
    End If
    
    ReleaseObjects
End Sub

' オブジェクト変数を初期化
Private Function InitializeObjects() As Boolean
    On Error GoTo ErrorHandler
    Set masterSheet = ThisWorkbook.Worksheets(SHEET_MASTER)
    Set logSheet = ThisWorkbook.Worksheets(SHEET_LOG)
    Set lastMasterDataCell = masterSheet.Cells(masterSheet.Rows.Count, COL_MASTER_ID).End(xlUp)
    InitializeObjects = True
    Exit Function

ErrorHandler:
    MsgBox "初期化に失敗しました。シート名が変更されていないか確認してください。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical
    InitializeObjects = False
End Function

' オブジェクト変数を解放
Private Sub ReleaseObjects()
    Set masterSheet = Nothing
    Set logSheet = Nothing
    Set lastMasterDataCell = Nothing
End Sub

' frmSakeLoggerから呼び出される公開プロシージャ
Public Function GetMasterSheet() As Worksheet
    Set GetMasterSheet = masterSheet
End Function

Public Function GetLogSheet() As Worksheet
    Set GetLogSheet = logSheet
End Function

Public Function GetLastMasterDataCell() As Range
    Set GetLastMasterDataCell = lastMasterDataCell
End Function
