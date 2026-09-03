Attribute VB_Name = "modPageSetup"
Option Explicit

Public Const CONN_PAGE_MIN_PRINT_ROW As Long = 30

Public Function LastUsedChartRow(wsHarness As Worksheet) As Long
    Dim r As Long

    For r = modHarnessBuild.SAVED_CHART_LAST_ROW To modHarnessBuild.SAVED_CHART_FIRST_ROW Step -1
        If Application.WorksheetFunction.CountA( _
                wsHarness.Range(wsHarness.Cells(r, 1), wsHarness.Cells(r, 11))) > 0 Then
            LastUsedChartRow = r
            Exit Function
        End If
    Next r

    LastUsedChartRow = modHarnessBuild.SAVED_CHART_HEADER_ROW
End Function

Private Function FooterText(ByVal sHarnessNumber As String, ByVal sRevision As String) As String
    FooterText = Trim$(sHarnessNumber & " Rev " & sRevision) & " - Page &P of &N"
End Function

Private Sub ApplyNarrowMargins(ps As PageSetup)
    ps.LeftMargin = Application.InchesToPoints(0.25)
    ps.RightMargin = Application.InchesToPoints(0.25)
    ps.TopMargin = Application.InchesToPoints(0.75)
    ps.BottomMargin = Application.InchesToPoints(0.75)
    ps.HeaderMargin = Application.InchesToPoints(0.3)
    ps.FooterMargin = Application.InchesToPoints(0.3)
End Sub

Public Sub ApplyHarnessPageSetup(wsHarness As Worksheet, ByVal sHarnessNumber As String, ByVal sRevision As String)
    Dim nLastRow As Long
    nLastRow = LastUsedChartRow(wsHarness)

    With wsHarness.PageSetup
        .PrintArea = "$A$1:$K$" & CStr(nLastRow)
        .PrintTitleRows = "$1:$" & modHarnessBuild.SAVED_CHART_HEADER_ROW
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .CenterFooter = FooterText(sHarnessNumber, sRevision)
        ApplyNarrowMargins wsHarness.PageSetup
    End With
End Sub

Private Function LastUsedTableRow(wsPage As Worksheet) As Long
    Dim nTableLast As Long
    nTableLast = wsPage.Cells(wsPage.Rows.Count, modConnectorPage.CONN_TABLE_FIRST_COL).End(xlUp).Row
    If nTableLast < CONN_PAGE_MIN_PRINT_ROW Then nTableLast = CONN_PAGE_MIN_PRINT_ROW
    LastUsedTableRow = nTableLast
End Function

Public Sub ApplyConnectorPageSetup(wsPage As Worksheet, ByVal sHarnessNumber As String, ByVal sRevision As String)
    With wsPage.PageSetup
        .PrintArea = "$A$1:$Q$" & CStr(LastUsedTableRow(wsPage))
        .PrintTitleRows = ""
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Zoom = False
        .CenterFooter = FooterText(sHarnessNumber, sRevision)
        ApplyNarrowMargins wsPage.PageSetup
    End With
End Sub
