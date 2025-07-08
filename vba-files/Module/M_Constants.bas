Attribute VB_Name = "M_Constants"
Option Explicit

' --- シート名 ---
Public Const SHEET_MASTER As String = "お酒マスタ"
Public Const SHEET_LOG As String = "飲酒記録"
Public Const SHEET_SUMMARY As String = "集計"
Public Const SHEET_HOLIDAY As String = "祝日マスタ"

' --- お酒マスタシート列インデックス ---
Public Const COL_MASTER_ID As Long = 1
Public Const COL_MASTER_NAME As Long = 2
Public Const COL_MASTER_KIND As Long = 3
Public Const COL_MASTER_ALCOHOL As Long = 4
Public Const COL_MASTER_FULL_WEIGHT As Long = 5
Public Const COL_MASTER_EMPTY_WEIGHT As Long = 6

' --- 飲酒記録シート列インデックス ---
Public Const COL_LOG_ID As Long = 1
Public Const COL_LOG_DATE As Long = 2
Public Const COL_LOG_NAME As Long = 3
Public Const COL_LOG_CURRENT_WEIGHT As Long = 4
Public Const COL_LOG_PURE_ALCOHOL As Long = 5
Public Const COL_LOG_DRANK_WEIGHT As Long = 6
Public Const COL_LOG_COMMENT As Long = 7

' --- 集計シート列インデックス ---
Public Const COL_SUMMARY_DATE As Long = 1
Public Const COL_SUMMARY_PURE_ALCOHOL As Long = 2
Public Const COL_SUMMARY_TOTAL As Long = 6
Public Const COL_SUMMARY_MONTHLY_TOTAL As Long = 7
Public Const COL_SUMMARY_HELPER_START_DATE As Long = 8
Public Const COL_SUMMARY_HELPER_END_DATE As Long = 9
