Attribute VB_Name = "modHarnessActions"
Option Explicit

Public Function SaveHarness(destWb As Workbook) As Variant
    If Not modHarnessBuild.BuildHarnessSheets(destWb) Then
        SaveHarness = modContract.Failure("HARNESS_SAVE_FAILED", "destination workbook is not fresh")
        Exit Function
    End If

    Dim wsHarness As Worksheet, wsSnapshot As Worksheet
    Set wsHarness = destWb.Worksheets("Harness")
    Set wsSnapshot = destWb.Worksheets("_Snapshot")

    modHarnessBuild.CopyTitleBlock wsHarness
    Dim sHarnessNumber As String, sRevision As String
    sHarnessNumber = CStr(wsHarness.Range("E2").Value)
    sRevision = CStr(wsHarness.Range("H2").Value)

    Dim nUsedRows As Long
    nUsedRows = modHarnessBuild.CopyChartRows(wsHarness)
    modHarnessBuild.CopyChartValidation destWb
    modHarnessBuild.CopySnapshot wsSnapshot
    modHarnessBuild.BuildConnectorPages destWb, wsSnapshot, sHarnessNumber, sRevision
    modPageSetup.ApplyHarnessPageSetup wsHarness, sHarnessNumber, sRevision

    SaveHarness = modContract.Success("HARNESS_SAVED", nUsedRows)
End Function

Private Function SheetExists(wb As Workbook, ByVal sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Public Function LoadHarness(srcWb As Workbook) As Variant
    If Not SheetExists(srcWb, "Harness") Or Not SheetExists(srcWb, "_Snapshot") Then
        LoadHarness = modContract.Failure("HARNESS_LOAD_FAILED", "not a harness file")
        Exit Function
    End If

    modChart.NewHarness

    Dim wsSrcHarness As Worksheet, wsSrcSnapshot As Worksheet
    Set wsSrcHarness = srcWb.Worksheets("Harness")
    Set wsSrcSnapshot = srcWb.Worksheets("_Snapshot")

    Dim wsDestHarness As Worksheet, wsDestSnapshot As Worksheet, wsDestConnectors As Worksheet
    Set wsDestHarness = ThisWorkbook.Worksheets(modChart.CHART_SHEET)
    Set wsDestSnapshot = ThisWorkbook.Worksheets("_Snapshot")
    Set wsDestConnectors = ThisWorkbook.Worksheets(modConnectors.CONN_SHEET)

    modHarnessLoad.CopySnapshotInto wsSrcSnapshot, wsDestSnapshot
    modHarnessLoad.CopyTitleBlockValues wsSrcHarness, wsDestHarness
    modHarnessLoad.RebuildConnectorInstances srcWb, wsDestSnapshot, wsDestConnectors

    Dim nUsedRows As Long
    nUsedRows = modHarnessLoad.CopyChartValues(wsSrcHarness, wsDestHarness)

    Dim r As Long
    For r = modChart.CHART_FIRST_ROW To modChart.CHART_LAST_ROW
        modChart.RebuildPinValidation r, modChart.COL_FROM_CONN, False
        modChart.RebuildPinValidation r, modChart.COL_TO_CONN, False
    Next r

    modState.SetState "HarnessPath", srcWb.FullName
    modState.ClearDirty

    LoadHarness = modContract.Success("HARNESS_LOADED", nUsedRows)
End Function
