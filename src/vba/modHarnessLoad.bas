Attribute VB_Name = "modHarnessLoad"
Option Explicit

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

Public Function CopyChartValues(wsSrcHarness As Worksheet, wsDestHarness As Worksheet) As Long
    Dim r As Long, c As Long, n As Long, sFrom As String, sTo As String
    Dim bEvents As Boolean

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    For r = modChart.CHART_FIRST_ROW To modChart.CHART_LAST_ROW
        For c = 1 To 11
            wsDestHarness.Cells(r, c).Value = wsSrcHarness.Cells(r, c).Value
        Next c
        sFrom = Trim$(CStr(wsSrcHarness.Cells(r, 1).Value))
        sTo = Trim$(CStr(wsSrcHarness.Cells(r, 9).Value))
        If Len(sFrom) > 0 Or Len(sTo) > 0 Then n = n + 1
    Next r

CleanUp:
    Application.EnableEvents = bEvents
    CopyChartValues = n
End Function

Public Function RebuildConnectorInstances(srcWb As Workbook, wsSnapshot As Worksheet, _
                                          wsConnectors As Worksheet) As Long
    Dim sh As Worksheet, sRefDes As String, sConnectorID As String, vFields As Variant
    Dim r As Long, n As Long, bEvents As Boolean

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    r = modConnectors.CONN_FIRST_ROW
    For Each sh In srcWb.Worksheets
        If Left$(sh.Name, 5) = "CONN_" Then
            sRefDes = Mid$(sh.Name, 6)
            sConnectorID = Trim$(CStr(sh.Cells(1, modConnectorPage.CONN_META_COL).Value))

            vFields = modLibrary.ReadConnector(wsSnapshot, modSnapshot.SNAP_CONN_FIRST_ROW, _
                modSnapshot.SNAP_CONN_LAST_ROW, sConnectorID)
            If Not IsEmpty(vFields) Then
                wsConnectors.Cells(r, 1).Value = sRefDes
                wsConnectors.Cells(r, 2).Value = sConnectorID
                wsConnectors.Cells(r, 3).Value = vFields(LBound(vFields) + modLibrary.LIB_COL_NAME - 1)
                wsConnectors.Cells(r, 4).Value = vFields(LBound(vFields) + modLibrary.LIB_COL_PARTNUM - 1)
                wsConnectors.Cells(r, 5).Value = vFields(LBound(vFields) + modLibrary.LIB_COL_TYPE - 1)
                wsConnectors.Cells(r, 6).Value = vFields(LBound(vFields) + modLibrary.LIB_COL_PINCOUNT - 1)
                r = r + 1
                n = n + 1
            End If
        End If
    Next sh

CleanUp:
    Application.EnableEvents = bEvents
    RebuildConnectorInstances = n
End Function
