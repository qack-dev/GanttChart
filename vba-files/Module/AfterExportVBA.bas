Attribute VB_Name = "AfterExportVBA"
Option Explicit

Sub ForceRestoreWindow()
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Windows.Count > 0 Then
            With wb.Windows(1)
                .Visible = True
                .WindowState = xlNormal
                .Top = 100
                .Left = 100
                .Height = 800
                .Width = 1200
            End With
        End If
    Next
    MsgBox "すべてのウィンドウの表示とサイズを復元しました。"
End Sub

