Attribute VB_Name = "modHarnessLoad"
Option Explicit

Private Const PHOTO_SHAPE_PREFIX As String = "PHOTO_"
Private Const TB_VALUE_CELLS_NO_UNITS As String = "B2,E2,H2,B3,E3,H3,B4"

' Replaces the Creator's own _Snapshot with the one carried by a saved
' harness file, so a Load leaves no trace of the prior session's data.
Public Sub CopySnapshotInto(wsSrcSnapshot As Worksheet, wsDestSnapshot As Worksheet)
    CopySnapshotValues wsSrcSnapshot, wsDestSnapshot
    CopySnapshotShapes wsSrcSnapshot, wsDestSnapshot
End Sub

Private Sub CopySnapshotValues(wsSrcSnapshot As Worksheet, wsDestSnapshot As Worksheet)
    Dim r As Long, c As Long

    For r = modSnapshot.SNAP_CONN_FIRST_ROW To modSnapshot.SNAP_CONN_LAST_ROW
        For c = 1 To modLibrary.LIB_FIELD_COUNT
            wsDestSnapshot.Cells(r, c).Value = wsSrcSnapshot.Cells(r, c).Value
        Next c
    Next r

    For r = modSnapshot.SNAP_PINS_FIRST_ROW To modSnapshot.SNAP_PINS_LAST_ROW
        For c = 1 To modLibrary.PIN_FIELD_COUNT
            wsDestSnapshot.Cells(r, c).Value = wsSrcSnapshot.Cells(r, c).Value
        Next c
    Next r
End Sub

' Worksheet.Paste needs its target sheet to be the active sheet of the
' active workbook. modHarnessBuild.CopySnapshot gets that for free - it
' pastes into the freshly-created destination workbook, which is already
' active. On Load the direction is reversed: the destination is the Creator
' while the just-opened source file is active, so the very hidden _Snapshot
' has to be unhidden and activated for the duration of the paste (the same
' thing modLibrary.ExportShapeToFile does for its own clipboard target) and
' restored afterwards. Only the activation is borrowed from
' ExportShapeToFile - not its Chart.Paste, which is confirmed unreliable on
' this machine (docs/superpowers/plans/phase-2-manual-verification.md, 2b/2c).
Private Sub CopySnapshotShapes(wsSrcSnapshot As Worksheet, wsDestSnapshot As Worksheet)
    Dim shp As Shape
    Dim nOriginalVisible As XlSheetVisibility
    Dim wbPriorActive As Workbook, wsPriorActive As Worksheet
    Dim nErr As Long, sErr As String

    Do While wsDestSnapshot.Shapes.Count > 0
        wsDestSnapshot.Shapes(1).Delete
    Loop

    If wsSrcSnapshot.Shapes.Count = 0 Then Exit Sub

    Set wbPriorActive = Application.ActiveWorkbook
    Set wsPriorActive = wsDestSnapshot.Parent.ActiveSheet
    nOriginalVisible = wsDestSnapshot.Visible

    wsDestSnapshot.Visible = xlSheetVisible
    wsDestSnapshot.Parent.Activate
    wsDestSnapshot.Activate

    On Error GoTo CleanUp
    For Each shp In wsSrcSnapshot.Shapes
        shp.Copy
        wsDestSnapshot.Paste
        wsDestSnapshot.Shapes(wsDestSnapshot.Shapes.Count).Name = shp.Name
    Next shp

CleanUp:
    nErr = Err.Number
    sErr = Err.Description

    On Error Resume Next
    wsPriorActive.Activate
    wsDestSnapshot.Visible = nOriginalVisible
    wbPriorActive.Activate
    On Error GoTo 0

    If nErr <> 0 Then Err.Raise nErr, "modHarnessLoad.CopySnapshotShapes", sErr
End Sub

Public Sub CopyTitleBlockValues(wsSrcHarness As Worksheet, wsDestHarness As Worksheet)
    Dim vCells As Variant, i As Long, sCell As String
    Dim bEvents As Boolean

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    vCells = Split(TB_VALUE_CELLS_NO_UNITS, ",")
    For i = LBound(vCells) To UBound(vCells)
        sCell = CStr(vCells(i))
        wsDestHarness.Range(sCell).Value = wsSrcHarness.Range(sCell).Value
    Next i

CleanUp:
    Application.EnableEvents = bEvents
    modChart.SetLengthUnits CStr(wsSrcHarness.Range("H4").Value)
End Sub
