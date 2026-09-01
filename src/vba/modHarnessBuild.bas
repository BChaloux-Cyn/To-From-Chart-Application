Attribute VB_Name = "modHarnessBuild"
Option Explicit

Public Const SAVED_CHART_HEADER_ROW As Long = 6
Public Const SAVED_CHART_FIRST_ROW As Long = 7
Public Const SAVED_CHART_LAST_ROW As Long = 1006
Public Const SAVED_COL_JOIN_FROM As Long = 12
Public Const SAVED_COL_JOIN_TO As Long = 13

Public Const XL_SHEET_VERY_HIDDEN As Long = 2

Private Const TB_VALUE_CELLS As String = "B2,E2,H2,B3,E3,H3,B4,H4"

' Test-only accessor: modConnectorPage's formula builders hardcode 7/1006 as
' literal string fragments rather than referencing these constants (an
' intentional choice - see the module's header), so a pytest tripwire needs
' a way to read the live constant value to detect if the two ever drift.
Public Function SavedChartLastRow() As Long
    SavedChartLastRow = SAVED_CHART_LAST_ROW
End Function

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

Public Sub CopyTitleBlock(wsDestHarness As Worksheet)
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(modChart.CHART_SHEET)

    wsDestHarness.Range("A1").Value = "WIRE HARNESS TO-FROM CHART"
    wsDestHarness.Range("A1").Font.Size = 16
    wsDestHarness.Range("A1").Font.Bold = True

    Dim vCells As Variant, i As Long, sCell As String
    vCells = Split(TB_VALUE_CELLS, ",")
    For i = LBound(vCells) To UBound(vCells)
        sCell = CStr(vCells(i))
        wsDestHarness.Range(sCell).Value = wsSrc.Range(sCell).Value
        wsDestHarness.Range(sCell).Interior.Color = &HF2F2F2
    Next i

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
        modPageSetup.ApplyConnectorPageSetup wsPage, sHarnessNumber, sRevision
    Next i
End Sub
