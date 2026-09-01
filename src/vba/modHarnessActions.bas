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
    modHarnessBuild.CopySnapshot wsSnapshot
    modHarnessBuild.BuildConnectorPages destWb, wsSnapshot, sHarnessNumber, sRevision
    modPageSetup.ApplyHarnessPageSetup wsHarness, sHarnessNumber, sRevision

    SaveHarness = modContract.Success("HARNESS_SAVED", nUsedRows)
End Function
