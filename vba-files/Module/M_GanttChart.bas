Attribute VB_Name = "M_GanttChart"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// M_GanttChart モジュール
'// ガントチャートの描画、更新、全体進捗グラフの更新など、主要なロジックを格納します。
'////////////////////////////////////////////////////////////////////////////////////////////////////

' Tasksシートの列インデックス
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME As Long = 2
Private Const COL_DURATION As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_END_DATE As Long = 5
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

' Settingsシートのセル参照 (行番号)
Private Const SETTING_CHART_START_ROW As Long = 1 ' チャート開始行/列
Private Const SETTING_BAR_HEIGHT As Long = 2      ' バーの高さ
Private Const SETTING_ROW_HEIGHT As Long = 3      ' 行の高さ
Private Const SETTING_COL_WIDTH As Long = 4       ' 列の幅
Private Const SETTING_COLOR_UNSTARTED As Long = 5 ' 未着手の色
Private Const SETTING_COLOR_IN_PROGRESS As Long = 6 ' 進行中の色
Private Const SETTING_COLOR_COMPLETED As Long = 7 ' 完了の色
Private Const SETTING_COLOR_DELAYED As Long = 8   ' 遅延の色

' ガントチャートを更新するメインプロシージャ
Public Sub UpdateGanttChart()
    On Error GoTo ErrHandler

    Dim wsGantt As Worksheet, wsTasks As Worksheet, wsSettings As Worksheet
    Dim lastTaskRow As Long, i As Long
    Dim chartStartRow As Long, chartStartCol As Long
    Dim barHeight As Double, rowHeight As Double, colWidth As Double
    Dim minDate As Date, maxDate As Date
    Dim taskID As Long, taskName As String, status As String, startDate As Date, endDate As Date
    Dim vValue As Variant, vTaskID As Variant, vStartDate As Variant, vEndDate As Variant
    Dim vDuration As Variant, vProgress As Variant

    Application.ScreenUpdating = False
    Set wsGantt = ThisWorkbook.Sheets("GanttChart")
    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    Call ClearGanttChart(wsGantt)

    vValue = wsSettings.Cells(SETTING_CHART_START_ROW, 2).Value
    If IsNumeric(vValue) Then chartStartRow = CLng(vValue) Else GoTo SettingsErr
    vValue = wsSettings.Cells(SETTING_CHART_START_ROW, 3).Value
    If IsNumeric(vValue) Then chartStartCol = CLng(vValue) Else GoTo SettingsErr
    vValue = wsSettings.Cells(SETTING_BAR_HEIGHT, 2).Value
    If IsNumeric(vValue) Then barHeight = CDbl(vValue) Else GoTo SettingsErr
    vValue = wsSettings.Cells(SETTING_ROW_HEIGHT, 2).Value
    If IsNumeric(vValue) Then rowHeight = CDbl(vValue) Else GoTo SettingsErr
    vValue = wsSettings.Cells(SETTING_COL_WIDTH, 2).Value
    If IsNumeric(vValue) Then colWidth = CDbl(vValue) Else GoTo SettingsErr

    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    If lastTaskRow < 2 Then
        MsgBox "タスクデータがありません。", vbInformation
        GoTo ExitHandler
    End If

    minDate = Date + 36500
    maxDate = Date - 36500
    For i = 2 To lastTaskRow
        vStartDate = wsTasks.Cells(i, COL_START_DATE).Value
        vEndDate = wsTasks.Cells(i, COL_END_DATE).Value
        If IsDate(vStartDate) And IsDate(vEndDate) Then
            If CDate(vStartDate) < minDate Then minDate = CDate(vStartDate)
            If CDate(vEndDate) > maxDate Then maxDate = CDate(vEndDate)
        End If
    Next i

    If minDate > maxDate Then
        MsgBox "有効な日付データを持つタスクがありません。", vbInformation
        GoTo ExitHandler
    End If

    Call DrawTimeline(wsGantt, minDate, maxDate, chartStartRow, chartStartCol, colWidth)

    For i = 2 To lastTaskRow
        vTaskID = wsTasks.Cells(i, COL_TASK_ID).Value
        vStartDate = wsTasks.Cells(i, COL_START_DATE).Value
        vEndDate = wsTasks.Cells(i, COL_END_DATE).Value
        vDuration = wsTasks.Cells(i, COL_DURATION).Value
        vProgress = wsTasks.Cells(i, COL_PROGRESS).Value

        If IsNumeric(vTaskID) And IsDate(vStartDate) And IsDate(vEndDate) And IsNumeric(vDuration) And IsNumeric(vProgress) Then
            If CDate(vEndDate) >= CDate(vStartDate) Then
                taskID = CLng(vTaskID)
                startDate = CDate(vStartDate)
                endDate = CDate(vEndDate)
                taskName = CStr(wsTasks.Cells(i, COL_TASK_NAME).Value)
                status = CStr(wsTasks.Cells(i, COL_STATUS).Value)
                
                Call DrawTaskBar(wsGantt, taskID, taskName, startDate, endDate, status, _
                                 chartStartRow + i - 1, chartStartCol, colWidth, barHeight, minDate)
            Else
                Debug.Print "行 " & i & ": 終了日が開始日より前のためスキップ"
            End If
        Else
            Debug.Print "行 " & i & ": ID,日付,期間,進捗のいずれかのデータが不正のため描画をスキップ"
        End If
    Next i

    Call UpdateLoadGraph(wsGantt, wsTasks, chartStartRow, lastTaskRow, minDate, maxDate)

ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub
SettingsErr:
    MsgBox "Settingsシートの設定値が不正です。" & vbCrLf & "数値が入力されるべきセルが空白または文字列になっていないか確認してください。", vbCritical, "設定エラー"
    GoTo ExitHandler
ErrHandler:
    MsgBox "ガントチャートの更新中に予期せぬエラーが発生しました: " & Err.Description, vbCritical
    GoTo ExitHandler
End Sub

Private Sub ClearGanttChart(wsGantt As Worksheet)
    On Error Resume Next
    Dim sh As Shape
    For Each sh In wsGantt.Shapes
        If sh.Name <> "UpdateChartButton" Then sh.Delete
    Next sh
    On Error GoTo 0
End Sub

Private Sub DrawTaskBar(wsGantt As Worksheet, taskID As Long, taskName As String, _
                        startDate As Date, endDate As Date, status As String, _
                        rowNum As Long, chartStartCol As Long, colWidth As Double, barHeight As Double, _
                        minChartDate As Date)
    Dim barLeft As Double, barTop As Double, barWidth As Double, barColor As Long
    Dim sh As Shape
    barLeft = wsGantt.Cells(rowNum, chartStartCol).Left + (startDate - minChartDate) * colWidth
    barTop = wsGantt.Cells(rowNum, 1).Top + (wsGantt.Cells(rowNum, 1).Height - barHeight) / 2
    barWidth = (endDate - startDate + 1) * colWidth
    If barWidth <= 0 Then Exit Sub
    barColor = GetColorByStatus(status)
    Set sh = wsGantt.Shapes.AddShape(msoShapeRectangle, barLeft, barTop, barWidth, barHeight)
    With sh
        .Fill.ForeColor.RGB = barColor
        .Line.Visible = msoFalse
        .Name = "TaskBar_" & taskID
        .TextFrame2.TextRange.Text = taskName
        With .TextFrame2.TextRange.Font
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Size = 8
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
    End With
End Sub

Private Sub DrawTimeline(wsGantt As Worksheet, startDate As Date, endDate As Date, _
                         chartStartRow As Long, chartStartCol As Long, colWidth As Double)
    If chartStartRow <= 1 Then Err.Raise vbObjectError, , "チャートの開始行は2行目以降に設定してください。"
    Dim currentDate As Date, colOffset As Long, headerRow As Long
    headerRow = chartStartRow - 1
    colOffset = 0
    For currentDate = startDate To endDate
        With wsGantt.Cells(headerRow, chartStartCol + colOffset)
            .Value = Format(currentDate, "m/d")
            .ColumnWidth = colWidth / 7
            .HorizontalAlignment = xlCenter
            If Weekday(currentDate) = vbSaturday Or Weekday(currentDate) = vbSunday Then
                .Interior.Color = RGB(240, 240, 240)
            End If
        End With
        colOffset = colOffset + 1
    Next currentDate
End Sub

Private Sub UpdateLoadGraph(wsGantt As Worksheet, wsTasks As Worksheet, _
                            chartStartRow As Long, lastTaskRow As Long, _
                            minChartDate As Date, maxChartDate As Date)
    On Error GoTo ErrHandler
    Dim i As Long, totalDuration As Double, completedDuration As Double, progressPercentage As Double
    Dim chartObj As ChartObject, chartName As String
    chartName = "OverallProgressChart"
    
    On Error Resume Next
    wsGantt.ChartObjects(chartName).Delete
    On Error GoTo 0 ' エラーハンドリングを元に戻す

    totalDuration = 0
    completedDuration = 0
    For i = 2 To lastTaskRow
        If IsNumeric(wsTasks.Cells(i, COL_DURATION).Value) And IsNumeric(wsTasks.Cells(i, COL_PROGRESS).Value) Then
            totalDuration = totalDuration + CDbl(wsTasks.Cells(i, COL_DURATION).Value)
            completedDuration = completedDuration + (CDbl(wsTasks.Cells(i, COL_DURATION).Value) * CDbl(wsTasks.Cells(i, COL_PROGRESS).Value))
        End If
    Next i

    If totalDuration > 0 Then progressPercentage = completedDuration / totalDuration Else progressPercentage = 0
    wsGantt.Range("A1").Value = progressPercentage
    wsGantt.Range("B1").Value = 1 - progressPercentage

    Set chartObj = wsGantt.ChartObjects.Add(Left:=wsGantt.Cells(chartStartRow, 2).Left, _
                                            Top:=wsGantt.Cells(lastTaskRow + 3, 1).Top, _
                                            Width:=200, Height:=120)
    With chartObj
        .Name = chartName
        With .Chart
            .ChartType = xlDoughnut
            .SetSourceData Source:=wsGantt.Range("A1:B1")
            .HasTitle = True
            .ChartTitle.Text = "全体進捗率"
            .ChartTitle.Font.Size = 10
            .HasLegend = False
            .ChartGroups(1).DoughnutHoleSize = 60

            '--- ★★★ 修正箇所 1: データラベルの適用方法を変更 ★★★
            With .SeriesCollection(1)
                .Points(1).Interior.Color = RGB(0, 176, 80)
                .Points(2).Interior.Color = RGB(220, 220, 220)
                
                ' 進捗部分(Points(1))にのみデータラベルを適用し、書式設定する
                .Points(1).ApplyDataLabels
                With .Points(1).DataLabel
                    .ShowValue = True
                    .NumberFormat = "0%"
                    .Font.Size = 12
                    .Position = xlLabelPositionCenter
                End With
            End With
        End With
    End With
    
    wsGantt.Range("A1:B1").ClearContents
    Exit Sub
ErrHandler:
    MsgBox "全体進捗グラフの更新中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

Private Function GetColorByStatus(status As String) As Long
    '--- ★★★ 修正箇所 2: 色の取得方法を .Interior.Color から .Value に変更 ★★★
    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    Const VALUE_COL As Long = 2
    Dim vColor As Variant
    Dim defaultColor As Long
    defaultColor = RGB(192, 192, 192) ' デフォルト色 (灰色)
    
    Select Case status
        Case "未着手"
            vColor = wsSettings.Cells(SETTING_COLOR_UNSTARTED, VALUE_COL).Value
        Case "進行中"
            vColor = wsSettings.Cells(SETTING_COLOR_IN_PROGRESS, VALUE_COL).Value
        Case "完了"
            vColor = wsSettings.Cells(SETTING_COLOR_COMPLETED, VALUE_COL).Value
        Case "遅延"
            vColor = wsSettings.Cells(SETTING_COLOR_DELAYED, VALUE_COL).Value
        Case Else
            GetColorByStatus = defaultColor
            Exit Function
    End Select
    
    ' 値が数値なら色として返し、そうでなければデフォルト色を返す
    If IsNumeric(vColor) Then
        GetColorByStatus = CLng(vColor)
    Else
        GetColorByStatus = defaultColor
    End If
End Function
