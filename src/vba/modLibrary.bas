Attribute VB_Name = "modLibrary"
Option Explicit

' Column layout of the Connectors table, shared by the library workbook,
' an imported library file, and (in a later phase) the _Snapshot sheet's
' Connectors block.
Public Const LIB_COL_ID As Long = 1
Public Const LIB_COL_NAME As Long = 2
Public Const LIB_COL_MFG As Long = 3
Public Const LIB_COL_PARTNUM As Long = 4
Public Const LIB_COL_TYPE As Long = 5
Public Const LIB_COL_PINCOUNT As Long = 6
Public Const LIB_COL_NOTES As Long = 7
Public Const LIB_COL_PHOTOSHAPE As Long = 8
Public Const LIB_COL_CREATED As Long = 9
Public Const LIB_COL_MODIFIED As Long = 10
Public Const LIB_COL_ORIGIN As Long = 11
Public Const LIB_FIELD_COUNT As Long = 11

' A generous default window for callers addressing a dedicated whole sheet
' (the library workbook or an imported library file), where there is no
' second table sharing the same rows.
Public Const LIB_ROW_CAP As Long = 100000

Public Function LastUsedRowInWindow(ws As Worksheet, ByVal nCol As Long, ByVal nLastRow As Long) As Long
    ' Public: modPinEditor (sub-plan 2b) reuses this for the same bounded-
    ' window-safe single-row delete it needs for "Delete Pin."
    '
    ' Cells(nLastRow, nCol).End(xlUp) only finds the true last-used row when
    ' the starting cell itself is empty. If nLastRow already holds data -
    ' the window is full, or a caller probes a small window right at
    ' existing data - End(xlUp) instead walks UP through the contiguous
    ' non-blank run and overshoots past the real data (potentially into a
    ' header row above nFirstRow). Since no function here ever looks past
    ' nLastRow anyway, an occupied nLastRow already IS the answer.
    If Len(Trim$(CStr(ws.Cells(nLastRow, nCol).Value))) > 0 Then
        LastUsedRowInWindow = nLastRow
    Else
        LastUsedRowInWindow = ws.Cells(nLastRow, nCol).End(xlUp).Row
    End If
End Function

Public Function FindConnectorRow(wsConn As Worksheet, ByVal nFirstRow As Long, _
                                 ByVal nLastRow As Long, ByVal sConnectorID As String) As Long
    Dim r As Long, nLast As Long

    nLast = LastUsedRowInWindow(wsConn, LIB_COL_ID, nLastRow)
    If nLast < nFirstRow Then Exit Function

    For r = nFirstRow To nLast
        If StrComp(Trim$(CStr(wsConn.Cells(r, LIB_COL_ID).Value)), sConnectorID, vbTextCompare) = 0 Then
            FindConnectorRow = r
            Exit Function
        End If
    Next r
End Function

Public Function WriteConnector(wsConn As Worksheet, ByVal nFirstRow As Long, _
                               ByVal nLastRow As Long, ByVal vFields As Variant) As Boolean
    Dim r As Long, c As Long, nLast As Long

    If UBound(vFields) - LBound(vFields) + 1 <> LIB_FIELD_COUNT Then Exit Function

    r = FindConnectorRow(wsConn, nFirstRow, nLastRow, CStr(vFields(LBound(vFields))))
    If r = 0 Then
        nLast = LastUsedRowInWindow(wsConn, LIB_COL_ID, nLastRow)
        If nLast < nFirstRow Then
            r = nFirstRow
        Else
            r = nLast + 1
        End If
        If r > nLastRow Then Exit Function
    End If

    For c = 1 To LIB_FIELD_COUNT
        wsConn.Cells(r, c).Value = vFields(LBound(vFields) + c - 1)
    Next c

    WriteConnector = True
End Function

Public Function ReadConnector(wsConn As Worksheet, ByVal nFirstRow As Long, _
                              ByVal nLastRow As Long, ByVal sConnectorID As String) As Variant
    Dim r As Long, vResult(1 To 11) As Variant, c As Long

    r = FindConnectorRow(wsConn, nFirstRow, nLastRow, sConnectorID)
    If r = 0 Then Exit Function

    For c = 1 To LIB_FIELD_COUNT
        vResult(c) = wsConn.Cells(r, c).Value
    Next c

    ReadConnector = vResult
End Function

Public Function DeleteConnector(wsConn As Worksheet, ByVal nFirstRow As Long, _
                                ByVal nLastRow As Long, ByVal sConnectorID As String) As Boolean
    Dim r As Long, nLast As Long, c As Long

    r = FindConnectorRow(wsConn, nFirstRow, nLastRow, sConnectorID)
    If r = 0 Then Exit Function

    nLast = LastUsedRowInWindow(wsConn, LIB_COL_ID, nLastRow)

    ' Swap the last row's data into the deleted row's slot, then clear the
    ' (now-duplicate) last row. This keeps every write inside
    ' [nFirstRow, nLastRow] - a plain Range.Delete Shift:=xlUp would pull
    ' rows from below nLastRow upward too, corrupting whatever else shares
    ' the sheet (the _Snapshot Pins block, in a later phase).
    If r < nLast Then
        For c = 1 To LIB_FIELD_COUNT
            wsConn.Cells(r, c).Value = wsConn.Cells(nLast, c).Value
        Next c
    End If
    wsConn.Range(wsConn.Cells(nLast, 1), wsConn.Cells(nLast, LIB_FIELD_COUNT)).ClearContents

    DeleteConnector = True
End Function

Public Function SlugifyConnectorID(ByVal sPartNumber As String, ByVal sName As String) As String
    Dim sSource As String, sResult As String, i As Long, ch As String

    sSource = Trim$(sPartNumber)
    If Len(sSource) = 0 Then sSource = Trim$(sName)
    sSource = UCase$(sSource)

    For i = 1 To Len(sSource)
        ch = Mid$(sSource, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            sResult = sResult & ch
        Else
            sResult = sResult & "-"
        End If
    Next i

    SlugifyConnectorID = sResult
End Function

Public Function UniqueConnectorID(wsConn As Worksheet, ByVal nFirstRow As Long, _
                                  ByVal nLastRow As Long, ByVal sBaseID As String) As String
    Dim sCandidate As String, nSuffix As Long

    sCandidate = sBaseID
    nSuffix = 1
    Do While FindConnectorRow(wsConn, nFirstRow, nLastRow, sCandidate) > 0
        nSuffix = nSuffix + 1
        sCandidate = sBaseID & "-" & CStr(nSuffix)
    Loop

    UniqueConnectorID = sCandidate
End Function
