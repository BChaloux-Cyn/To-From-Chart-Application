Attribute VB_Name = "modPageSetup"
Option Explicit

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

Public Sub ApplyHarnessPageSetup(wsHarness As Worksheet, ByVal sHarnessNumber As String, ByVal sRevision As String)
    Dim nLastRow As Long
    nLastRow = LastUsedChartRow(wsHarness)

    With wsHarness.PageSetup
        .PrintArea = "$A$1:$K$" & CStr(nLastRow)
        .PrintTitleRows = "$" & modHarnessBuild.SAVED_CHART_HEADER_ROW & ":$" & modHarnessBuild.SAVED_CHART_HEADER_ROW
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .CenterFooter = FooterText(sHarnessNumber, sRevision)
    End With
End Sub
