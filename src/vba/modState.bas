Attribute VB_Name = "modState"
Option Explicit

Public Const STATE_SHEET As String = "_State"
Private Const STATE_FIRST_ROW As Long = 2

Private Function KeyRow(ws As Worksheet, ByVal sKey As String) As Long
    Dim r As Long, nLast As Long
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = STATE_FIRST_ROW To nLast
        If StrComp(CStr(ws.Cells(r, 1).Value), sKey, vbTextCompare) = 0 Then
            KeyRow = r
            Exit Function
        End If
    Next r
    KeyRow = 0
End Function

Public Function GetState(ByVal sKey As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(STATE_SHEET)
    r = KeyRow(ws, sKey)
    If r = 0 Then
        GetState = ""
    Else
        GetState = CStr(ws.Cells(r, 2).Value)
    End If
End Function

Public Sub SetState(ByVal sKey As String, ByVal sValue As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(STATE_SHEET)
    r = KeyRow(ws, sKey)
    If r = 0 Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < STATE_FIRST_ROW Then r = STATE_FIRST_ROW
        ws.Cells(r, 1).Value = sKey
    End If
    ' Text format stops Excel coercing "TRUE" into a Boolean.
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value = sValue
End Sub

Public Sub MarkDirty()
    SetState "Dirty", "TRUE"
End Sub

Public Sub ClearDirty()
    SetState "Dirty", "FALSE"
End Sub

Public Function IsDirty() As Boolean
    IsDirty = (UCase$(GetState("Dirty")) = "TRUE")
End Function
