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

' Settingsシートのセル参照
Private Const SETTING_CHART_START_ROW As Long = 1
Private Const SETTING_CHART_START_COL As Long = 2
Private Const SETTING_BAR_HEIGHT As Long = 3
Private Const SETTING_ROW_HEIGHT As Long = 4
Private Const SETTING_COL_WIDTH As Long = 5
Private Const SETTING_COLOR_UNSTARTED As Long = 6
Private Const SETTING_COLOR_IN_PROGRESS As Long = 7
Private Const SETTING_COLOR_COMPLETED As Long = 8
Private Const SETTING_COLOR_DELAYED As Long = 9

' ガントチャートを更新するメインプロシージャ
Public Sub UpdateGanttChart()
    On Error GoTo ErrHandler

    Dim wsGantt As Worksheet
    Dim wsTasks As Worksheet
    Dim wsSettings As Worksheet
    Dim lastTaskRow As Long
    Dim i As Long
    Dim taskID As Long
    Dim taskName As String
    Dim duration As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim progress As Double
    Dim status As String
    Dim minDate As Date
    Dim maxDate As Date
    Dim chartStartCol As Long
    Dim chartStartRow As Long
    Dim barHeight As Long
    Dim rowHeight As Long
    Dim colWidth As Long

    Set wsGantt = ThisWorkbook.Sheets("GanttChart")
    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    ' 既存のチャートをクリア
    Call ClearGanttChart(wsGantt)

    ' 設定値の読み込み
    chartStartRow = wsSettings.Cells(SETTING_CHART_START_ROW, SETTING_CHART_START_COL).Value ' 例: Settings!B1 に開始行
    chartStartCol = wsSettings.Cells(SETTING_CHART_START_ROW, SETTING_CHART_START_COL + 1).Value ' 例: Settings!C1 に開始列
    barHeight = wsSettings.Cells(SETTING_CHART_START_ROW, SETTING_BAR_HEIGHT + 1).Value    ' 例: Settings!D1 にバーの高さ
    rowHeight = wsSettings.Cells(SETTING_CHART_START_ROW, SETTING_ROW_HEIGHT + 1).Value    ' 例: Settings!E1 に行の高さ
    colWidth = wsSettings.Cells(SETTING_CHART_START_ROW, SETTING_COL_WIDTH + 1).Value     ' 例: Settings!F1 に列の幅

    ' タスクデータの最終行を取得 (TasksシートのB列を基準)
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row

    If lastTaskRow < 2 Then ' ヘッダー行のみの場合
        MsgBox "タスクデータがありません。", vbInformation
        Exit Sub
    End If

    ' 日付範囲の特定
    minDate = wsTasks.Cells(2, COL_START_DATE).Value ' 開始日のヘッダー
    maxDate = wsTasks.Cells(2, COL_END_DATE).Value ' 終了日のヘッダー

    For i = 2 To lastTaskRow
        If wsTasks.Cells(i, COL_START_DATE).Value < minDate Then minDate = wsTasks.Cells(i, COL_START_DATE).Value
        If wsTasks.Cells(i, COL_END_DATE).Value > maxDate Then maxDate = wsTasks.Cells(i, COL_END_DATE).Value
    Next i

    ' タイムラインの描画
    Call DrawTimeline(wsGantt, minDate, maxDate, chartStartRow, chartStartCol, colWidth)

    ' 各タスクのバーを描画
    For i = 2 To lastTaskRow
        taskID = wsTasks.Cells(i, COL_TASK_ID).Value
        taskName = wsTasks.Cells(i, COL_TASK_NAME).Value
        duration = wsTasks.Cells(i, COL_DURATION).Value
        startDate = wsTasks.Cells(i, COL_START_DATE).Value
        endDate = wsTasks.Cells(i, COL_END_DATE).Value
        progress = wsTasks.Cells(i, COL_PROGRESS).Value
        status = wsTasks.Cells(i, COL_STATUS).Value

        ' タスクバーの描画
        Call DrawTaskBar(wsGantt, taskID, taskName, startDate, endDate, status, _
                         chartStartRow + i - 1, chartStartCol, colWidth, barHeight, minDate)
    Next i

    ' 全体進捗グラフの更新
    Call UpdateLoadGraph(wsGantt, wsTasks, chartStartRow, chartStartCol, colWidth, minDate, maxDate)

    Exit Sub

ErrHandler:
    MsgBox "ガントチャートの更新中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 既存のガントチャートをクリアする
Private Sub ClearGanttChart(wsGantt As Worksheet)
    On Error Resume Next ' エラーが発生しても処理を続行

    Dim sh As Shape
    For Each sh In wsGantt.Shapes
        If Left(sh.Name, 8) = "TaskBar_" Or Left(sh.Name, 9) = "Timeline_" Or Left(sh.Name, 10) = "Progress_" Then
            sh.Delete
        End If
    Next sh

    ' グラフもクリア (もしあれば)
    For Each sh In wsGantt.Shapes
        If sh.Type = msoChart Then
            sh.Delete
        End If
    Next sh

    On Error GoTo 0 ' エラーハンドリングをリセット
End Sub

' 1つのタスクに対応するバーを描画する
Private Sub DrawTaskBar(wsGantt As Worksheet, taskID As Long, taskName As String, _
                        startDate As Date, endDate As Date, status As String, _
                        rowNum As Long, chartStartCol As Long, colWidth As Long, barHeight As Long, _
                        minChartDate As Date)
    On Error GoTo ErrHandler

    Dim barLeft As Double
    Dim barTop As Double
    Dim barWidth As Double
    Dim barColor As Long
    Dim sh As Shape

    ' バーの開始位置と幅を計算
    barLeft = wsGantt.Cells(rowNum, 1).Left + (startDate - minChartDate) * colWidth
    barTop = wsGantt.Cells(rowNum, 1).Top + (wsGantt.Cells(rowNum, 1).Height - barHeight) / 2
    barWidth = (endDate - startDate + 1) * colWidth

    ' ステータスに応じた色を取得
    barColor = GetColorByStatus(status)

    ' バーを描画
    Set sh = wsGantt.Shapes.AddShape(msoShapeRectangle, barLeft, barTop, barWidth, barHeight)
    With sh
        .Fill.ForeColor.RGB = barColor
        .Line.Visible = msoFalse
        .Name = "TaskBar_" & taskID ' タスクIDを名前に含める
        .OnAction = "M_ChartEvents.ShowTaskDetails" ' クリックイベントのマクロを割り当て
        .TextFrame2.TextRange.Text = taskName
        With .TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0) ' テキスト色を黒に設定
            .Transparency = 0
            .Solid
        End With
        .TextFrame2.TextRange.Font.Size = 8
        .TextFrame2.TextRange.Font.Bold = msoFalse
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.WordArtformat = msoTextEffect1
    End With

    Exit Sub

ErrHandler:
    MsgBox "タスクバーの描画中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' タイムラインを描画する
Private Sub DrawTimeline(wsGantt As Worksheet, startDate As Date, endDate As Date, _
                         chartStartRow As Long, chartStartCol As Long, colWidth As Long)
    On Error GoTo ErrHandler

    Dim currentDate As Date
    Dim colOffset As Long
    Dim headerRow As Long

    headerRow = chartStartRow - 1 ' タイムラインのヘッダー行

    ' 日付ヘッダーのクリア
    wsGantt.Range(wsGantt.Cells(headerRow, chartStartCol), wsGantt.Cells(headerRow, chartStartCol + (endDate - startDate + 1))).ClearContents

    colOffset = 0
    For currentDate = startDate To endDate
        wsGantt.Cells(headerRow, chartStartCol + colOffset).Value = Format(currentDate, "m/d")
        wsGantt.Cells(headerRow, chartStartCol + colOffset).ColumnWidth = colWidth / 6 ' 日付表示に合わせて調整
        wsGantt.Cells(headerRow, chartStartCol + colOffset).HorizontalAlignment = xlCenter
        wsGantt.Cells(headerRow, chartStartCol + colOffset).VerticalAlignment = xlCenter
        wsGantt.Cells(headerRow, chartStartCol + colOffset).Orientation = 90 ' 縦書き

        ' 週末の背景色を変更
        If Weekday(currentDate, vbSaturday) = vbSaturday Or Weekday(currentDate, vbSaturday) = vbSunday Then
            With wsGantt.Cells(headerRow, chartStartCol + colOffset).Interior
                .Color = RGB(220, 220, 220) ' 薄い灰色
            End With
        Else
            With wsGantt.Cells(headerRow, chartStartCol + colOffset).Interior
                .Pattern = xlNone
            End With
        End If

        colOffset = colOffset + 1
    Next currentDate

    Exit Sub

ErrHandler:
    MsgBox "タイムラインの描画中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 全体進捗グラフを更新する
Private Sub UpdateLoadGraph(wsGantt As Worksheet, wsTasks As Worksheet, _
                            chartStartRow As Long, chartStartCol As Long, colWidth As Long, _
                            minChartDate As Date, maxChartDate As Date)
    On Error GoTo ErrHandler

    Dim lastTaskRow As Long
    Dim i As Long
    Dim totalDuration As Double
    Dim completedDuration As Double
    Dim progressPercentage As Double
    Dim chartObj As ChartObject
    Dim chartName As String

    chartName = "OverallProgressChart"

    ' 既存のグラフを削除
    For Each chartObj In wsGantt.ChartObjects
        If chartObj.Name = chartName Then
            chartObj.Delete
            Exit For
        End If
    Next chartObj

    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    totalDuration = 0
    completedDuration = 0

    For i = 2 To lastTaskRow
        Dim duration As Long
        Dim progress As Double
        Dim status As String

        duration = wsTasks.Cells(i, COL_DURATION).Value ' 期間
        progress = wsTasks.Cells(i, COL_PROGRESS).Value ' 進捗
        status = wsTasks.Cells(i, COL_STATUS).Value   ' ステータス

        totalDuration = totalDuration + duration

        If status = "完了" Then
            completedDuration = completedDuration + duration
        Else
            completedDuration = completedDuration + (duration * progress)
        End If
    Next i

    If totalDuration > 0 Then
        progressPercentage = completedDuration / totalDuration
    Else
        progressPercentage = 0
    End If

    ' グラフのデータ範囲を設定 (一時的にシートに書き出す)
    wsGantt.Cells(1, 1).Value = "進捗"
    wsGantt.Cells(1, 2).Value = progressPercentage

    ' グラフの作成
    Set chartObj = wsGantt.ChartObjects.Add(Left:=wsGantt.Cells(chartStartRow, chartStartCol).Left, _
                                            Top:=wsGantt.Cells(chartStartRow, chartStartCol).Top + (maxChartDate - minChartDate + 2) * wsGantt.Cells(1, 1).Height, _
                                            Width:=300, Height:=150)
    With chartObj
        .Name = chartName
        With .Chart
            .ChartType = xlDoughnut
            .SetSourceData Source:=wsGantt.Range(wsGantt.Cells(1, 1), wsGantt.Cells(1, 2))
            .HasTitle = True
            .ChartTitle.Text = "全体進捗率"
            .ChartTitle.Font.Size = 10
            .HasLegend = False
            .DoughnutHoleSize = 60

            ' データ系列の設定
            With .SeriesCollection(1)
                .Points(1).Interior.Color = RGB(0, 176, 80) ' 完了部分 (緑)
                .Points(2).Interior.Color = RGB(200, 200, 200) ' 未完了部分 (灰色)
                .ApplyDataLabels
                .DataLabels.ShowPercentage = True
                .DataLabels.Font.Size = 10
                .DataLabels.Position = xlLabelPositionCenter
            End With
        End With
    End With

    Exit Sub

ErrHandler:
    MsgBox "全体進捗グラフの更新中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' ステータスに応じた色を返す関数
Private Function GetColorByStatus(status As String) As Long
    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    Select Case status
        Case "未着手"
            GetColorByStatus = wsSettings.Cells(SETTING_CHART_START_ROW + 1, SETTING_COL_WIDTH + 1).Value ' 例: Settings!G2 に未着手の色
        Case "進行中"
            GetColorByStatus = wsSettings.Cells(SETTING_CHART_START_ROW + 2, SETTING_COL_WIDTH + 1).Value ' 例: Settings!G3 に進行中の色
        Case "完了"
            GetColorByStatus = wsSettings.Cells(SETTING_CHART_START_ROW + 3, SETTING_COL_WIDTH + 1).Value ' 例: Settings!G4 に完了の色
        Case "遅延"
            GetColorByStatus = wsSettings.Cells(SETTING_CHART_START_ROW + 4, SETTING_COL_WIDTH + 1).Value ' 例: Settings!G5 に遅延の色
        Case Else
            GetColorByStatus = RGB(192, 192, 192) ' デフォルト色 (灰色)
    End Select
End Function
