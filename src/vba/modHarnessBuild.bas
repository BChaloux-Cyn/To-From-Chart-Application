Attribute VB_Name = "modHarnessBuild"
Option Explicit

Public Const SAVED_CHART_HEADER_ROW As Long = 6
Public Const SAVED_CHART_FIRST_ROW As Long = 7
Public Const SAVED_CHART_LAST_ROW As Long = 1006
Public Const SAVED_COL_JOIN_FROM As Long = 12
Public Const SAVED_COL_JOIN_TO As Long = 13

Public Const XL_SHEET_VERY_HIDDEN As Long = 2

Private Const TB_VALUE_CELLS As String = "B2,E2,H2,B3,E3,H3,B4,H4"

' Widens each free-text title-block value past its single narrow chart-grid
' column so the gray fill (and the text) doesn't stop short when a value -
' Harness Name and Description especially - is longer than one column.
' H4 (Length Units) is a short controlled value and is left unmerged.
Private Const TB_MERGE_SPANS As String = "B2:C2,E2:F2,H2:I2,B3:C3,E3:F3,H3:I3,B4:F4"

Private Const TB_LABEL_CELLS As String = "A2,D2,G2,A3,D3,G3,A4,G4"
Private Const TB_LABELS As String = _
    "Harness Name,Harness Number,Revision,Student,Class / Project,Date,Description,Length Units"

' Mirrors build/layout.py's CHART_COLUMN_WIDTHS - the saved harness is built
' from a blank workbook in VBA, so it needs its own copy of the same widths
' rather than inheriting them from a template. A (Harness Name/Description)
' and G (Length Units) are widened past the Creator's original chart-only
' sizing so the title-block labels in those columns don't clip.
Private Const CHART_COLUMN_WIDTHS As String = "14,9,15,18,13,7,13,15,11,9,30"

' Test-only accessor: modConnectorPage's formula builders hardcode 7/1006 as
' literal string fragments rather than referencing these constants (an
' intentional choice - see the module's header), so a pytest tripwire needs
' a way to read the live constant value to detect if the two ever drift.
Public Function SavedChartLastRow() As Long
    SavedChartLastRow = SAVED_CHART_LAST_ROW
End Function

Public Function BuildHarnessSheets(destWb As Workbook) As Boolean
    If destWb.Worksheets.Count <> 1 Then Exit Function

    Dim original As Worksheet, wsHarness As Worksheet, wsSnapshot As Worksheet, wsLists As Worksheet
    Set original = destWb.Worksheets(1)

    Set wsHarness = destWb.Worksheets.Add(After:=original)
    wsHarness.Name = "Harness"
    Set wsSnapshot = destWb.Worksheets.Add(After:=wsHarness)
    wsSnapshot.Name = "_Snapshot"
    Set wsLists = destWb.Worksheets.Add(After:=wsSnapshot)
    wsLists.Name = "_Lists"

    Dim bPriorAlerts As Boolean
    bPriorAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    original.Delete
    Application.DisplayAlerts = bPriorAlerts

    wsSnapshot.Visible = XL_SHEET_VERY_HIDDEN
    wsLists.Visible = XL_SHEET_VERY_HIDDEN

    BuildHarnessSheets = True
End Function

Public Sub CopyTitleBlock(wsDestHarness As Worksheet)
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(modChart.CHART_SHEET)

    wsDestHarness.Range("A1").Value = "WIRE HARNESS TO-FROM CHART"
    wsDestHarness.Range("A1").Font.Size = 16
    wsDestHarness.Range("A1").Font.Bold = True

    Dim vLabelCells As Variant, vLabels As Variant, k As Long
    vLabelCells = Split(TB_LABEL_CELLS, ",")
    vLabels = Split(TB_LABELS, ",")
    For k = LBound(vLabelCells) To UBound(vLabelCells)
        wsDestHarness.Range(CStr(vLabelCells(k))).Value = vLabels(k)
        wsDestHarness.Range(CStr(vLabelCells(k))).Font.Bold = True
    Next k

    Dim vCells As Variant, i As Long, sCell As String
    vCells = Split(TB_VALUE_CELLS, ",")
    For i = LBound(vCells) To UBound(vCells)
        sCell = CStr(vCells(i))
        wsDestHarness.Range(sCell).Value = wsSrc.Range(sCell).Value
        wsDestHarness.Range(sCell).Interior.Color = &HF2F2F2
    Next i

    Dim vSpans As Variant, j As Long
    vSpans = Split(TB_MERGE_SPANS, ",")
    For j = LBound(vSpans) To UBound(vSpans)
        wsDestHarness.Range(CStr(vSpans(j))).Merge
    Next j

    For i = LBound(vCells) To UBound(vCells)
        With wsDestHarness.Range(CStr(vCells(i))).MergeArea.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    Next i

    Dim vWidths As Variant, w As Long
    vWidths = Split(CHART_COLUMN_WIDTHS, ",")
    For w = LBound(vWidths) To UBound(vWidths)
        wsDestHarness.Columns(w + 1).ColumnWidth = CLng(vWidths(w))
    Next w

    Dim vHeaders As Variant, nUnitsIndex As Long
    vHeaders = Array("From Conn", "From Pin", "From Term", "Signal", "Color", "AWG", _
                      wsSrc.Cells(SAVED_CHART_HEADER_ROW, 7).Value, "To Term", "To Conn", "To Pin", "Notes")
    For i = LBound(vHeaders) To UBound(vHeaders)
        Dim cel As Range
        Set cel = wsDestHarness.Cells(SAVED_CHART_HEADER_ROW, i + 1)
        cel.Value = vHeaders(i)
        cel.Font.Bold = True
        cel.Interior.Color = &HD9D9D9
    Next i
End Sub

Public Function CopyChartRows(wsDestHarness As Worksheet) As Long
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(modChart.CHART_SHEET)

    Dim r As Long, c As Long, n As Long
    Dim sFrom As String, sTo As String

    For r = SAVED_CHART_FIRST_ROW To SAVED_CHART_LAST_ROW
        For c = 1 To 11
            wsDestHarness.Cells(r, c).Value = wsSrc.Cells(r, c).Value
        Next c

        sFrom = Trim$(CStr(wsSrc.Cells(r, 1).Value))
        sTo = Trim$(CStr(wsSrc.Cells(r, 9).Value))
        If Len(sFrom) > 0 Or Len(sTo) > 0 Then n = n + 1

        wsDestHarness.Cells(r, SAVED_COL_JOIN_FROM).Formula = _
            "=IF(A" & r & "="""",""""," & "A" & r & "&""|""&B" & r & ")"
        wsDestHarness.Cells(r, SAVED_COL_JOIN_TO).Formula = _
            "=IF(I" & r & "="""",""""," & "I" & r & "&""|""&J" & r & ")"
    Next r

    wsDestHarness.Columns(SAVED_COL_JOIN_FROM).Hidden = True
    wsDestHarness.Columns(SAVED_COL_JOIN_TO).Hidden = True

    CopyChartRows = n
End Function

' Carries the Creator's Color/AWG/Termination dropdowns into the saved
' harness so a hand-edit on the macro-free file is still constrained to
' valid values, not just live-recalculated - the Creator's own chart
' validation (test_validation.py's EXPECTED) has no automated coverage on
' the saved file, since CopyChartRows only ever copied cell values. From/To
' Conn and From/To Pin are intentionally excluded - those depend on which
' connectors exist in this specific harness (modChart.RebuildPinValidation),
' not a static list, so replicating them here would need dynamic per-row
' logic a macro-free workbook cannot run.
Public Sub CopyChartValidation(destWb As Workbook)
    Dim wsSrcLists As Worksheet, wsDestLists As Worksheet, wsDestHarness As Worksheet
    Set wsSrcLists = ThisWorkbook.Worksheets("_Lists")
    Set wsDestLists = destWb.Worksheets("_Lists")
    Set wsDestHarness = destWb.Worksheets("Harness")

    Dim nCol As Long, nLastRow As Long, r As Long
    For nCol = 1 To 3
        nLastRow = wsSrcLists.Cells(wsSrcLists.Rows.Count, nCol).End(xlUp).Row
        For r = 1 To nLastRow
            ' Text format keeps AWG sizes as strings rather than numbers,
            ' matching build_lists' own cell in the Creator's "_Lists" sheet -
            ' without it, "24" round-trips as the number 24.
            wsDestLists.Cells(r, nCol).NumberFormat = "@"
            wsDestLists.Cells(r, nCol).Value = wsSrcLists.Cells(r, nCol).Value
        Next r
    Next nCol

    destWb.Names.Add Name:="ListColor", RefersTo:= _
        "=OFFSET('_Lists'!$A$2,0,0,MAX(1,COUNTA('_Lists'!$A:$A)-1),1)"
    destWb.Names.Add Name:="ListAWG", RefersTo:= _
        "=OFFSET('_Lists'!$B$2,0,0,MAX(1,COUNTA('_Lists'!$B:$B)-1),1)"
    destWb.Names.Add Name:="ListTermination", RefersTo:= _
        "=OFFSET('_Lists'!$C$2,0,0,MAX(1,COUNTA('_Lists'!$C:$C)-1),1)"

    ApplyListValidation wsDestHarness, 3, "=ListTermination"  ' From Term
    ApplyListValidation wsDestHarness, 5, "=ListColor"        ' Color
    ApplyListValidation wsDestHarness, 6, "=ListAWG"          ' AWG
    ApplyListValidation wsDestHarness, 8, "=ListTermination"  ' To Term
End Sub

Private Sub ApplyListValidation(ws As Worksheet, ByVal nCol As Long, ByVal sFormula As String)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(SAVED_CHART_FIRST_ROW, nCol), ws.Cells(SAVED_CHART_LAST_ROW, nCol))
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, Formula1:=sFormula
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
End Sub

Public Sub CopySnapshot(wsDestSnapshot As Worksheet)
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets("_Snapshot")

    Dim r As Long, c As Long

    For r = modSnapshot.SNAP_CONN_FIRST_ROW To modSnapshot.SNAP_CONN_LAST_ROW
        For c = 1 To modLibrary.LIB_FIELD_COUNT
            wsDestSnapshot.Cells(r, c).Value = wsSrc.Cells(r, c).Value
        Next c
    Next r

    For r = modSnapshot.SNAP_PINS_FIRST_ROW To modSnapshot.SNAP_PINS_LAST_ROW
        For c = 1 To modLibrary.PIN_FIELD_COUNT
            wsDestSnapshot.Cells(r, c).Value = wsSrc.Cells(r, c).Value
        Next c
    Next r

    Dim shp As Shape
    For Each shp In wsSrc.Shapes
        shp.Copy
        wsDestSnapshot.Paste
        wsDestSnapshot.Shapes(wsDestSnapshot.Shapes.Count).Name = shp.Name
    Next shp
End Sub

Public Sub BuildConnectorPages(destWb As Workbook, wsSnapshot As Worksheet, _
                               ByVal sHarnessNumber As String, ByVal sRevision As String)
    Dim vInstances As Variant, i As Long
    Dim sRefDes As String, sConnectorID As String
    Dim wsPage As Worksheet, vPins As Variant, shpPhoto As Shape
    Dim sPhotoPath As String

    vInstances = modConnectors.AllInstances()
    If IsEmpty(vInstances) Then Exit Sub

    For i = LBound(vInstances, 1) To UBound(vInstances, 1)
        sRefDes = CStr(vInstances(i, 1))
        sConnectorID = CStr(vInstances(i, 2))

        Set wsPage = destWb.Worksheets.Add(After:=destWb.Worksheets(destWb.Worksheets.Count))
        wsPage.Name = "CONN_" & sRefDes

        vPins = modLibrary.ReadPinsForConnector(wsSnapshot, modSnapshot.SNAP_PINS_FIRST_ROW, _
            modSnapshot.SNAP_PINS_LAST_ROW, sConnectorID)

        sPhotoPath = modConnectorPage.PagePhotoPath(modSnapshot.LibraryFolder(), sConnectorID)
        If modConnectorPage.PlacePhoto(wsPage, sPhotoPath) Then
            Set shpPhoto = wsPage.Shapes("PAGE_PHOTO")
            modConnectorPage.PlaceCallouts wsPage, shpPhoto, vPins
            modConnectorPage.PlaceLeaderLines wsPage, shpPhoto, vPins
        End If

        modConnectorPage.WriteTableSkeleton wsPage, vPins
        modConnectorPage.WriteLiveFormulas wsPage, sRefDes, vPins
        modConnectorPage.WriteMetadata wsPage, sConnectorID
        modConnectorPage.WritePageTitleBlock wsPage, sHarnessNumber, sRevision, sRefDes, sConnectorID
        modPageSetup.ApplyConnectorPageSetup wsPage, sHarnessNumber, sRevision
    Next i
End Sub
