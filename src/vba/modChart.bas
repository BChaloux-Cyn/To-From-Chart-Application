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

Public Sub RebuildPinValidation(ByVal nRow As Long, ByVal nConnCol As Long)
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
    cel.ClearContents
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
