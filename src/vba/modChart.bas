Attribute VB_Name = "modChart"
Option Explicit

Public Const CHART_SHEET As String = "Harness"
Public Const CHART_HEADER_ROW As Long = 6
Public Const CHART_FIRST_ROW As Long = 7
Public Const CHART_LAST_ROW As Long = 1006

Public Const COL_FROM_CONN As Long = 1
Public Const COL_FROM_PIN As Long = 2
Public Const COL_LENGTH As Long = 7
Public Const COL_TO_CONN As Long = 9
Public Const COL_TO_PIN As Long = 10
Public Const COL_NOTES As Long = 11

Private Const MAX_FORMULA1 As Long = 255

' Module-level Const/Dim must sit in the declarations section, before any
' Sub/Function: VBA does not recognize one declared between procedures.
Private Const TB_CLEAR_NAMES As String = _
    "TB_Name,TB_Number,TB_Rev,TB_Student,TB_Class,TB_Date,TB_Desc"

Public Sub RebuildPinValidation(ByVal nRow As Long, ByVal nConnCol As Long, _
                                Optional ByVal bClearStale As Boolean = True)
    Dim ws As Worksheet, cel As Range
    Dim nPinCol As Long, nPins As Long, i As Long
    Dim sRef As String, sList As String

    Select Case nConnCol
        Case COL_FROM_CONN: nPinCol = COL_FROM_PIN
        Case COL_TO_CONN:   nPinCol = COL_TO_PIN
        Case Else:          Exit Sub
    End Select

    Set ws = ThisWorkbook.Worksheets(CHART_SHEET)
    Set cel = ws.Cells(nRow, nPinCol)

    sRef = Trim$(CStr(ws.Cells(nRow, nConnCol).Value))
    nPins = modConnectors.PinCountFor(sRef)

    cel.Validation.Delete
    ' A single-cell connector edit invalidates whatever pin was already
    ' there, so it must be cleared. A bulk paste or a pin-count edit
    ' elsewhere sets/keeps the pin value deliberately, so it must not be
    ' wiped out from under the user - only the dropdown list is refreshed.
    If bClearStale Then cel.ClearContents
    If nPins < 1 Then Exit Sub

    For i = 1 To nPins
        If Len(sList) > 0 Then sList = sList & ","
        sList = sList & CStr(i)
    Next i

    If Len(sList) <= MAX_FORMULA1 Then
        cel.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                           Operator:=xlBetween, Formula1:=sList
        cel.Validation.InCellDropdown = True
    Else
        cel.Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                           Operator:=xlBetween, Formula1:="1", Formula2:=CStr(nPins)
    End If
    cel.Validation.IgnoreBlank = True
End Sub

Public Sub SetLengthUnits(ByVal sUnit As String)
    Dim ws As Worksheet
    Dim s As String
    Dim bEvents As Boolean

    s = LCase$(Trim$(sUnit))
    If s <> "in" And s <> "mm" Then Exit Sub

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    Set ws = ThisWorkbook.Worksheets(CHART_SHEET)
    ws.Cells(CHART_HEADER_ROW, COL_LENGTH).Value = "Length (" & s & ")"
    ThisWorkbook.Names("TB_Units").RefersToRange.Value = s
    modState.SetState "LengthUnits", s

CleanUp:
    Application.EnableEvents = bEvents
End Sub

Public Sub RefreshChartRowsForConnector(ByVal sRefDes As String)
    Dim ws As Worksheet
    Dim r As Long
    Dim sRef As String

    If Len(Trim$(sRefDes)) = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(CHART_SHEET)
    For r = CHART_FIRST_ROW To CHART_LAST_ROW
        sRef = Trim$(CStr(ws.Cells(r, COL_FROM_CONN).Value))
        If StrComp(sRef, sRefDes, vbTextCompare) = 0 Then
            RebuildPinValidation r, COL_FROM_CONN, False
        End If
        sRef = Trim$(CStr(ws.Cells(r, COL_TO_CONN).Value))
        If StrComp(sRef, sRefDes, vbTextCompare) = 0 Then
            RebuildPinValidation r, COL_TO_CONN, False
        End If
    Next r
End Sub

Public Sub NewHarness()
    Dim wsHarness As Worksheet, wsConn As Worksheet, wsCheck As Worksheet
    Dim vNames As Variant, i As Long
    Dim bEvents As Boolean

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    Set wsHarness = ThisWorkbook.Worksheets(CHART_SHEET)
    Set wsConn = ThisWorkbook.Worksheets(modConnectors.CONN_SHEET)
    Set wsCheck = ThisWorkbook.Worksheets("Check")

    wsHarness.Range(wsHarness.Cells(CHART_FIRST_ROW, COL_FROM_CONN), _
                    wsHarness.Cells(CHART_LAST_ROW, COL_NOTES)).ClearContents

    ' Pin validation is built per row, so it must be torn down per row too.
    wsHarness.Range(wsHarness.Cells(CHART_FIRST_ROW, COL_FROM_PIN), _
                    wsHarness.Cells(CHART_LAST_ROW, COL_FROM_PIN)).Validation.Delete
    wsHarness.Range(wsHarness.Cells(CHART_FIRST_ROW, COL_TO_PIN), _
                    wsHarness.Cells(CHART_LAST_ROW, COL_TO_PIN)).Validation.Delete

    vNames = Split(TB_CLEAR_NAMES, ",")
    For i = LBound(vNames) To UBound(vNames)
        ThisWorkbook.Names(vNames(i)).RefersToRange.ClearContents
    Next i

    wsConn.Range(wsConn.Cells(modConnectors.CONN_FIRST_ROW, 1), _
                 wsConn.Cells(wsConn.Rows.Count, 6)).ClearContents

    wsCheck.Range(wsCheck.Cells(2, 1), _
                  wsCheck.Cells(wsCheck.Rows.Count, 3)).ClearContents

    modState.SetState "HarnessPath", ""
    SetLengthUnits "in"
    modState.ClearDirty

CleanUp:
    Application.EnableEvents = bEvents
End Sub
