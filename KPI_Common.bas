Attribute VB_Name = "KPI_Common"
Option Explicit

' ================== SHARED CONFIG ==================
Public Const SRC_SHEET      As String = "Stope Cadence"
Public Const TGT_SHEET      As String = "SchedulerData"
Public Const TBL_NAME       As String = "KPI"
Public Const TABLE_ANCHOR   As String = "A25"
Public Const STATE_SHEET    As String = "_StageStateCache"
Public Const SUMMARY_ANCHOR As String = "A1"

' Source column numbers (1 based)
Public Const COL_ID       As Long = 3    ' C
Public Const COL_USER     As Long = 8    ' H
Public Const COL_STAGE    As Long = 22   ' V
Public Const COL_SUB      As Long = 23   ' W
Public Const COL_COMMENTS As Long = 24   ' X
Public Const COL_ZONE     As Long = 35   ' AI

' Base header names
Public Const H_ID    As String = "ID Stope"
Public Const H_USER  As String = "User"
Public Const H_STAGE As String = "Design Stage"
Public Const H_SUB   As String = "Sub Process"
Public Const H_ZONE  As String = "Zone Status"

' Thresholds for a "yes"
Public Const STEPS_RED   As Long = 1
Public Const STEPS_BLACK As Long = 2

' ================== SHARED FUNCTIONS ==================
Public Function SheetOrNothing(nm As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Public Function CleanStr(v As Variant) As String
    CleanStr = Trim(CStr(v))
End Function

Public Function ReadColumn(ws As Worksheet, col As Long, r1 As Long, r2 As Long) As Variant
    Dim v As Variant
    v = ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)).Value
    If Not IsArray(v) Then
        Dim a(1 To 1, 1 To 1) As Variant: a(1, 1) = v: v = a
    End If
    ReadColumn = v
End Function
