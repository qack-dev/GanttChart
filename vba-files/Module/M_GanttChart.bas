Attribute VB_Name = "M_GanttChart"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////
'// M_GanttChart モジュール
'// ガントチャートの描画、更新、負荷グラフの更新など、主要なロジックを格納します。
'// 従来のShapeオブジェクトによる描画方式から、セルの背景色を着色する方式に全面的に改修。
'////////////////////////////////////////////////////////////////////////////////////////////////////

'--- 色定義 ---
' 各タスクステータスの背景色を定義します。
Public Const COLOR_UNSTARTED As Long = &HC0C0C0 ' 未着手 (灰色)
Public Const COLOR_IN_PROGRESS As Long = &HFFFF00 ' 進行中 (黄色)
Public Const COLOR_COMPLETED As Long = &H92D050 ' 完了 (緑色)
Public Const COLOR_DELAYED As Long = &H0000FF ' 遅延 (赤色)

'--- Tasksシートの列インデックス ---
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME As Long = 2
Private Const COL_START_DATE As Long = 4
Private Const COL_DURATION As Long = 3
Private Const COL_PROGRESS As Long = 6
Private Const COL_STATUS As Long = 7

'--- GanttChartシートの固定行・列 ---
Private Const GANTT_START_ROW As Long = 5 ' ガントチャートの開始行
Private Const GANTT_START_COL As Long = 2 ' ガントチャートの開始列 (タスク名表示エリア)
Private Const TIMELINE_ROW As Long = 4    ' タイムラインの表示行

'====================================================================================================
'// Public Procedures
'====================================================================================================

'/**
' * @brief メインプロシージャ。ガントチャート全体を更新します。
' */
Public Sub UpdateGanttChart()
    On Error GoTo ErrHandler

    Dim wsGantt As Worksheet
    Dim wsTasks As Worksheet
    Dim lastTaskRow As Long
    Dim i As Long
    Dim minDate As Date
    Dim maxDate As Date

    Set wsGantt = ThisWorkbook.Sheets("GanttChart")
    Set wsTasks = ThisWorkbook.Sheets("Tasks")

    ' --- データの有効性チェック ---
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    If lastTaskRow < 2 Then
        MsgBox "タスクが入力されていません。", vbInformation
        Exit Sub
    End If

    ' --- チャート描画エリアのクリア ---
    Call ClearGanttArea(wsGantt, lastTaskRow)

    ' --- タイムスケールの決定 ---
    ' タスクリストから最小開始日と最大終了日を計算
    With Application.WorksheetFunction
        minDate = .Min(wsTasks.Range("D2:D" & lastTaskRow))
        maxDate = .Max(wsTasks.Range("E2:E" & lastTaskRow))
    End With

    ' --- 描画処理 ---
    Call DrawTimeline(wsGantt, minDate, maxDate)
    Call DrawAllTaskBars(wsGantt, wsTasks, lastTaskRow, minDate)
    ' Call UpdateLoadGraph(wsGantt, wsTasks, minDate, maxDate) ' 必要に応じてコメント解除

    Exit Sub

ErrHandler:
    MsgBox "ガントチャートの更新中にエラーが発生しました: " & vbCrLf & Err.Description, vbCritical
End Sub

'====================================================================================================
'// Private Procedures
'====================================================================================================

'/**
' * @brief ガントチャートの描画エリア（タイムライン、タスクバー、タスク名）をクリアします。
' * @param wsGantt 対象のGanttChartシート
' * @param lastTaskRow Tasksシートの最終行
' */
Private Sub ClearGanttArea(ByVal wsGantt As Worksheet, ByVal lastTaskRow As Long)
    On Error Resume Next ' クリア対象が存在しない場合も考慮

    ' --- タイムラインエリアのクリア ---
    wsGantt.Rows(TIMELINE_ROW).Clear

    ' --- タスク描画エリアのクリア ---
    ' 前回の描画範囲が不明なため、十分な範囲をクリアする
    Dim clearRange As Range
    Set clearRange = wsGantt.Range(wsGantt.Cells(GANTT_START_ROW, GANTT_START_COL), wsGantt.Cells(GANTT_START_ROW + lastTaskRow + 5, 256))
    
    With clearRange
        .ClearContents
        .Interior.Color = xlNone
        .Borders.LineStyle = xlNone
    End With
    
    On Error GoTo 0
End Sub

'/**
' * @brief タイムライン（日付ヘッダー）を描画します。
' * @param wsGantt 対象のGanttChartシート
' * @param startDate 表示する最初の日付
' * @param endDate 表示する最後の日付
' */
Private Sub DrawTimeline(ByVal wsGantt As Worksheet, ByVal startDate As Date, ByVal endDate As Date)
    Dim currentDate As Date
    Dim col As Long
    
    col = GANTT_START_COL + 1 ' タイムラインはタスク名の右の列から開始

    ' --- 日付の描画 ---
    For currentDate = startDate To endDate
        With wsGantt.Cells(TIMELINE_ROW, col)
            .Value = Format(currentDate, "m/d")
            .ColumnWidth = 4
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 8
            
            ' --- 週末のハイライト ---
            If Weekday(currentDate) = vbSaturday Or Weekday(currentDate) = vbSunday Then
                .Interior.Color = RGB(240, 240, 240) ' 薄い灰色
            End If
            
            ' --- 月の区切り線 ---
            If Day(currentDate) = 1 Then
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
            End If
        End With
        col = col + 1
    Next currentDate
End Sub

'/**
' * @brief すべてのタスクバー（セルの着色）を描画します。
' * @param wsGantt 対象のGanttChartシート
' * @param wsTasks Tasksシート
' * @param lastTaskRow Tasksシートの最終行
' * @param minDate タイムラインの開始日
' */
Private Sub DrawAllTaskBars(ByVal wsGantt As Worksheet, ByVal wsTasks As Worksheet, ByVal lastTaskRow As Long, ByVal minDate As Date)
    Dim i As Long
    For i = 2 To lastTaskRow
        Dim taskName As String
        Dim startDate As Date
        Dim duration As Long
        Dim status As String
        
        ' --- タスク情報を取得 ---
        With wsTasks.Rows(i)
            taskName = .Cells(COL_TASK_NAME).Value
            startDate = .Cells(COL_START_DATE).Value
            duration = .Cells(COL_DURATION).Value
            status = .Cells(COL_STATUS).Value
        End With
        
        ' --- タスクバーを描画 ---
        Call HighlightTaskPeriod(wsGantt, i - 1, taskName, startDate, duration, status, minDate)
    Next i
End Sub

'/**
' * @brief 個別のタスクバー（セルの着色）を描画します。
' * @param wsGantt 対象のGanttChartシート
' * @param taskRowIndex GanttChartシート上のタスクの行インデックス (1から始まる)
' * @param taskName タスク名
' * @param startDate タスクの開始日
' * @param duration タスクの期間（日数）
' * @param status タスクのステータス
' * @param minDate タイムラインの開始日
' */
Private Sub HighlightTaskPeriod(ByVal wsGantt As Worksheet, ByVal taskRowIndex As Long, ByVal taskName As String, ByVal startDate As Date, ByVal duration As Long, ByVal status As String, ByVal minDate As Date)
    On Error GoTo ErrHandler

    Dim startCol As Long
    Dim endCol As Long
    Dim taskRow As Long
    Dim barColor As Long
    Dim taskRange As Range

    ' --- 描画位置の計算 ---
    taskRow = GANTT_START_ROW + taskRowIndex - 1
    startCol = (startDate - minDate) + GANTT_START_COL + 1
    endCol = startCol + duration - 1

    ' --- タスク名の表示 ---
    wsGantt.Cells(taskRow, GANTT_START_COL).Value = taskName

    ' --- 期間セルの特定と着色 ---
    If startCol <= endCol Then
        Set taskRange = wsGantt.Range(wsGantt.Cells(taskRow, startCol), wsGantt.Cells(taskRow, endCol))
        
        ' --- ステータスに応じた色を取得 ---
        barColor = GetColorByStatus(status)
        
        ' --- セルの書式設定 ---
        With taskRange.Interior
            .Color = barColor
        End With
        
        ' --- タスクバーに枠線を追加 ---
        With taskRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(150, 150, 150)
        End With
    End If

    Exit Sub

ErrHandler:
    MsgBox "タスクバーの描画中にエラーが発生しました: " & vbCrLf & "タスク名: " & taskName & vbCrLf & Err.Description, vbCritical
End Sub

'/**
' * @brief ステータス文字列に対応する色定数を返します。
' * @param status タスクのステータス
' * @return 対応する色のLong値
' */
Private Function GetColorByStatus(ByVal status As String) As Long
    Select Case status
        Case "未着手"
            GetColorByStatus = COLOR_UNSTARTED
        Case "進行中"
            GetColorByStatus = COLOR_IN_PROGRESS
        Case "完了"
            GetColorByStatus = COLOR_COMPLETED
        Case "遅延"
            GetColorByStatus = COLOR_DELAYED
        Case Else
            GetColorByStatus = vbWhite ' 不明なステータスは白
    End Select
End Function

'/**
' * @brief （参考）負荷グラフを更新します。今回の改修範囲外ですが、必要に応じて利用します。
' */
Private Sub UpdateLoadGraph(wsGantt As Worksheet, wsTasks As Worksheet, minDate As Date, maxDate As Date)
    ' このプロシージャは今回の改修要件には含まれていませんが、
    ' 必要に応じてセルベースのデータと連携するように改修可能です。
    ' (現在の実装はShapeに依存している可能性があるため、レビューが必要です)
    MsgBox "UpdateLoadGraphは現在実装されていません。", vbInformation
End Sub