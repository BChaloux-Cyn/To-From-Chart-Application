Attribute VB_Name = "modHarnessBuild"
Option Explicit

Public Const SAVED_CHART_HEADER_ROW As Long = 6
Public Const SAVED_CHART_FIRST_ROW As Long = 7
Public Const SAVED_CHART_LAST_ROW As Long = 1006
Public Const SAVED_COL_JOIN_FROM As Long = 12
Public Const SAVED_COL_JOIN_TO As Long = 13

Public Const XL_SHEET_VERY_HIDDEN As Long = 2

Public Function BuildHarnessSheets(destWb As Workbook) As Boolean
    If destWb.Worksheets.Count <> 1 Then Exit Function

    Dim original As Worksheet, wsHarness As Worksheet, wsSnapshot As Worksheet
    Set original = destWb.Worksheets(1)

    Set wsHarness = destWb.Worksheets.Add(After:=original)
    wsHarness.Name = "Harness"
    Set wsSnapshot = destWb.Worksheets.Add(After:=wsHarness)
    wsSnapshot.Name = "_Snapshot"

    Dim bPriorAlerts As Boolean
    bPriorAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    original.Delete
    Application.DisplayAlerts = bPriorAlerts

    wsSnapshot.Visible = XL_SHEET_VERY_HIDDEN

    BuildHarnessSheets = True
End Function
