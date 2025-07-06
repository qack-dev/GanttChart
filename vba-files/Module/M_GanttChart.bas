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

    Dim wsGantt As Worksheet
    Dim wsTasks As Worksheet
    Dim wsSettings As Worksheet
    Dim lastTaskRow As Long
    Dim i As Long
    Dim minDate As Date, maxDate As Date
    Dim chartStartRow As Long, chartStartCol As Long
    Dim barHeight As Double, rowHeight As Double, colWidth As Double
    
    '--- 変数の初期化 ---
    Dim vStartDate As Variant, vEndDate As Variant
    Dim taskID As Long, duration As Long
    Dim taskName As String, status As String
    Dim progress As Double

    Set wsGantt = ThisWorkbook.Sheets("GanttChart")
    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    ' 既存のチャートと図形をクリア (更新ボタンは除く)
    Call ClearGanttChart(wsGantt)

    ' 設定値の読み込み (SettingsシートのB列から値を読み込む)
    chartStartRow = wsSettings.Cells(SETTING_CHART_START_ROW, 2).Value
    chartStartCol = wsSettings.Cells(SETTING_CHART_START_ROW, 3).Value ' 開始列はC列
    barHeight = wsSettings.Cells(SETTING_BAR_HEIGHT, 2).Value
    rowHeight = wsSettings.Cells(SETTING_ROW_HEIGHT, 2).Value
    colWidth = wsSettings.Cells(SETTING_COL_WIDTH, 2).Value

    ' タスクデータの最終行を取得 (TasksシートのB列を基準)
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row

    If lastTaskRow < 2 Then ' ヘッダー行のみの場合
        MsgBox "タスクデータがありません。", vbInformation
        Exit Sub
    End If

    ' --- 日付範囲の特定 (有効な日付のみを対象) ---
    minDate = Date + 36500 ' 未来の大きな日付で初期化
    maxDate = Date - 36500 ' 過去の小さな日付で初期化
    
    For i = 2 To lastTaskRow
        vStartDate = wsTasks.Cells(i, COL_START_DATE).Value
        vEndDate = wsTasks.Cells(i, COL_END_DATE).Value
        If IsDate(vStartDate) And IsDate(vEndDate) Then
            If CDate(vStartDate) < minDate Then minDate = CDate(vStartDate)
            If CDate(vEndDate) > maxDate Then maxDate = CDate(vEndDate)
        End If
    Next i
    
    ' 有効なタスクが一つもなかった場合
    If minDate > maxDate Then
        MsgBox "有効な日付データを持つタスクがありません。", vbInformation
        Exit Sub
    End If

    ' タイムラインの描画
    Call DrawTimeline(wsGantt, minDate, maxDate, chartStartRow, chartStartCol, colWidth)

    ' 各タスクのバーを描画する前に、日付データの有効性をチェック
    For i = 2 To lastTaskRow
        vStartDate = wsTasks.Cells(i, COL_START_DATE).Value
        vEndDate = wsTasks.Cells(i, COL_END_DATE).Value

        ' 日付データが有効な場合のみ描画処理を行う
        If IsDate(vStartDate) And IsDate(vEndDate) Then
            ' さらに終了日が開始日以降かチェック
            If CDate(vEndDate) >= CDate(vStartDate) Then
                ' 有効なデータのみを変数に格納
                taskID = wsTasks.Cells(i, COL_TASK_ID).Value
                taskName = wsTasks.Cells(i, COL_TASK_NAME).Value
                duration = wsTasks.Cells(i, COL_DURATION).Value
                progress = wsTasks.Cells(i, COL_PROGRESS).Value
                status = wsTasks.Cells(i, COL_STATUS).Value

                ' タスクバーの描画
                Call DrawTaskBar(wsGantt, taskID, taskName, CDate(vStartDate), CDate(vEndDate), status, _
                                 chartStartRow + i - 1, chartStartCol, colWidth, barHeight, minDate)
            Else
                Debug.Print "行 " & i & ": 終了日が開始日より前のためスキップ"
            End If
        Else
            Debug.Print "行 " & i & ": 日付データが不正のためスキップ"
        End If
    Next i

    ' 全体進捗グラフの更新
    Call UpdateLoadGraph(wsGantt, wsTasks, chartStartRow, chartStartCol, colWidth, minDate, maxDate)

    Exit Sub

ErrHandler:
    MsgBox "ガントチャートの更新中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' ★★★ 修正箇所 ★★★
' "UpdateChartButton"という名前のボタン以外の図形をすべて削除する
Private Sub ClearGanttChart(wsGantt As Worksheet)
    On Error Resume Next ' エラーが発生しても処理を続行

    Dim sh As Shape
    ' For Eachループでシート上の全図形を確認
    For Each sh In wsGantt.Shapes
        ' 図形の名前が"UpdateChartButton"でない場合に限り、削除する
        If sh.Name <> "UpdateChartButton" Then
            sh.Delete
        End If
    Next sh

    On Error GoTo 0 ' エラーハンドリングをリセット
End Sub

' 1つのタスクに対応するバーを描画する
Private Sub DrawTaskBar(wsGantt As Worksheet, taskID As Long, taskName As String, _
                        startDate As Date, endDate As Date, status As String, _
                        rowNum As Long, chartStartCol As Long, colWidth As Double, barHeight As Double, _
                        minChartDate As Date)
    On Error GoTo ErrHandler

    Dim barLeft As Double, barTop As Double, barWidth As Double
    Dim barColor As Long
    Dim sh As Shape

    ' バーの開始位置と幅を計算
    barLeft = wsGantt.Cells(rowNum, chartStartCol).Left + (startDate - minChartDate) * colWidth
    barTop = wsGantt.Cells(rowNum, 1).Top + (wsGantt.Cells(rowNum, 1).Height - barHeight) / 2
    barWidth = (endDate - startDate + 1) * colWidth

    ' バーの幅が0より大きい場合のみ描画
    If barWidth <= 0 Then Exit Sub

    ' ステータスに応じた色を取得
    barColor = GetColorByStatus(status)

    ' バーを描画
    Set sh = wsGantt.Shapes.AddShape(msoShapeRectangle, barLeft, barTop, barWidth, barHeight)
    With sh
        .Fill.ForeColor.RGB = barColor
        .Line.Visible = msoFalse
        .Name = "TaskBar_" & taskID
        .TextFrame2.TextRange.Text = taskName
        With .TextFrame2.TextRange.Font
            .Fill.ForeColor.RGB = RGB(255, 255, 255) ' テキスト白
            .Size = 8
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
    End With

    Exit Sub

ErrHandler:
    MsgBox "タスクバーの描画中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' タイムラインを描画する
Private Sub DrawTimeline(wsGantt As Worksheet, startDate As Date, endDate As Date, _
                         chartStartRow As Long, chartStartCol As Long, colWidth As Double)
    On Error GoTo ErrHandler

    If chartStartRow <= 1 Then
        Err.Raise Number:=vbObjectError, Description:="チャートの開始行は2行目以降に設定してください。"
    End If
    
    Dim currentDate As Date
    Dim colOffset As Long
    Dim headerRow As Long
    headerRow = chartStartRow - 1

    ' タイムライン範囲の書式設定
    With wsGantt.Range(wsGantt.Cells(headerRow, chartStartCol), wsGantt.Cells(wsGantt.Rows.Count, chartStartCol + (endDate - startDate + 2)))
        .Clear
        .ColumnWidth = colWidth / 7
        .HorizontalAlignment = xlCenter
    End With

    colOffset = 0
    For currentDate = startDate To endDate
        With wsGantt.Cells(headerRow, chartStartCol + colOffset)
            .Value = Format(currentDate, "m/d")
            ' 週末の背景色を変更
            If Weekday(currentDate) = vbSaturday Or Weekday(currentDate) = vbSunday Then
                .Interior.Color = RGB(240, 240, 240)
            End If
        End With
        colOffset = colOffset + 1
    Next currentDate

    Exit Sub

ErrHandler:
    MsgBox "タイムラインの描画中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 全体進捗グラフを更新する
Private Sub UpdateLoadGraph(wsGantt As Worksheet, wsTasks As Worksheet, _
                            chartStartRow As Long, chartStartCol As Long, colWidth As Double, _
                            minChartDate As Date, maxChartDate As Date)
    On Error GoTo ErrHandler

    Dim lastTaskRow As Long
    Dim i As Long
    Dim totalDuration As Double
    Dim completedDuration As Double
    Dim progressPercentage As Double
    Dim chartObj As ChartObject
    Dim chartName As String
    Dim vProgress As Variant

    chartName = "OverallProgressChart"
    
    ' 念のため既存の同名グラフを削除
    On Error Resume Next
    wsGantt.ChartObjects(chartName).Delete
    On Error GoTo ErrHandler

    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    totalDuration = 0
    completedDuration = 0

    For i = 2 To lastTaskRow
        If IsNumeric(wsTasks.Cells(i, COL_DURATION).Value) And IsNumeric(wsTasks.Cells(i, COL_PROGRESS).Value) Then
            totalDuration = totalDuration + wsTasks.Cells(i, COL_DURATION).Value
            completedDuration = completedDuration + (wsTasks.Cells(i, COL_DURATION).Value * wsTasks.Cells(i, COL_PROGRESS).Value)
        End If
    Next i

    If totalDuration > 0 Then
        progressPercentage = completedDuration / totalDuration
    Else
        progressPercentage = 0
    End If

    ' グラフのデータとして「進捗率」と「残り」の2つの値を一時セルに書き出す
    wsGantt.Range("A1").Value = progressPercentage
    wsGantt.Range("B1").Value = 1 - progressPercentage

    ' グラフの作成
    Set chartObj = wsGantt.ChartObjects.Add(Left:=wsGantt.Cells(chartStartRow, chartStartCol).Left, _
                                            Top:=wsGantt.Cells(lastTaskRow + 2, 1).Top, _
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

            With .SeriesCollection(1)
                .Points(1).Interior.Color = RGB(0, 176, 80)    ' 完了部分 (緑)
                .Points(2).Interior.Color = RGB(220, 220, 220) ' 未完了部分 (灰色)
                
                ' 進捗率を中央に表示
                .ApplyDataLabels
                With .DataLabels(1)
                    .ShowValue = True
                    .ShowCategoryName = False
                    .ShowSeriesName = False
                    .ShowPercentage = False
                    .NumberFormat = "0%"
                    .Font.Size = 12
                    .Font.Bold = True
                    .Position = xlLabelPositionCenter
                End With
            End With
        End With
    End With
    
    ' 一時的に使用したセルをクリア
    wsGantt.Range("A1:B1").ClearContents

    Exit Sub

ErrHandler:
    MsgBox "全体進捗グラフの更新中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' ステータスに応じた色を返す関数
Private Function GetColorByStatus(status As String) As Long
    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    ' SettingsシートのB列にあるセルの「背景色」を取得
    Const VALUE_COL As Long = 2
    
    On Error Resume Next
    Select Case status
        Case "未着手"
            GetColorByStatus = wsSettings.Cells(SETTING_COLOR_UNSTARTED, VALUE_COL).Interior.Color
        Case "進行中"
            GetColorByStatus = wsSettings.Cells(SETTING_COLOR_IN_PROGRESS, VALUE_COL).Interior.Color
        Case "完了"
            GetColorByStatus = wsSettings.Cells(SETTING_COLOR_COMPLETED, VALUE_COL).Interior.Color
        Case "遅延"
            GetColorByStatus = wsSettings.Cells(SETTING_COLOR_DELAYED, VALUE_COL).Interior.Color
        Case Else
            GetColorByStatus = RGB(192, 192, 192) ' デフォルト色 (灰色)
    End Select
    
    If Err.Number <> 0 Then
        GetColorByStatus = RGB(192, 192, 192)
        Err.Clear
    End If
    On Error GoTo 0
End Function

