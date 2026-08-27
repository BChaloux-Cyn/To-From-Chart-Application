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

' Column layout of the Pins table. Declared here, at the top of the module
' alongside the Connectors columns, rather than next to WritePin/ReadPins -
' VBA requires a module-level Const to be declared before any Sub/Function
' that references it, or the compiler reports "Variable not defined" on the
' constant's own name.
Public Const PIN_COL_CONNID As Long = 1
Public Const PIN_COL_PINNUM As Long = 2
Public Const PIN_COL_LABEL As Long = 3
Public Const PIN_COL_NORMX As Long = 4
Public Const PIN_COL_NORMY As Long = 5
Public Const PIN_COL_LABELX As Long = 6
Public Const PIN_COL_LABELY As Long = 7
Public Const PIN_FIELD_COUNT As Long = 7

Public Const PHOTO_GRID_COLUMNS As Long = 4
Public Const PHOTO_GRID_CELL_WIDTH As Long = 120
Public Const PHOTO_GRID_CELL_HEIGHT As Long = 120
Public Const PHOTO_GRID_MARGIN As Long = 8

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

Public Function WritePin(wsPins As Worksheet, ByVal nFirstRow As Long, _
                         ByVal nLastRow As Long, ByVal vFields As Variant) As Boolean
    Dim r As Long, c As Long, nLast As Long

    If UBound(vFields) - LBound(vFields) + 1 <> PIN_FIELD_COUNT Then Exit Function

    nLast = LastUsedRowInWindow(wsPins, PIN_COL_CONNID, nLastRow)
    If nLast < nFirstRow Then
        r = nFirstRow
    Else
        r = nLast + 1
    End If
    If r > nLastRow Then Exit Function

    For c = 1 To PIN_FIELD_COUNT
        wsPins.Cells(r, c).Value = vFields(LBound(vFields) + c - 1)
    Next c

    WritePin = True
End Function

Private Sub SwapPinRows(vRows As Variant, ByVal a As Long, ByVal b As Long)
    Dim c As Long, vTmp As Variant
    For c = 1 To PIN_FIELD_COUNT
        vTmp = vRows(a, c)
        vRows(a, c) = vRows(b, c)
        vRows(b, c) = vTmp
    Next c
End Sub

Public Function ReadPinsForConnector(wsPins As Worksheet, ByVal nFirstRow As Long, _
                                     ByVal nLastRow As Long, ByVal sConnectorID As String) As Variant
    Dim nLast As Long, r As Long, n As Long, i As Long, j As Long
    Dim vRows() As Variant, vResult() As Variant

    nLast = LastUsedRowInWindow(wsPins, PIN_COL_CONNID, nLastRow)
    If nLast < nFirstRow Then Exit Function

    n = 0
    ReDim vRows(1 To nLast - nFirstRow + 1, 1 To PIN_FIELD_COUNT)
    For r = nFirstRow To nLast
        If StrComp(Trim$(CStr(wsPins.Cells(r, PIN_COL_CONNID).Value)), sConnectorID, vbTextCompare) = 0 Then
            n = n + 1
            For i = 1 To PIN_FIELD_COUNT
                vRows(n, i) = wsPins.Cells(r, i).Value
            Next i
        End If
    Next r
    If n = 0 Then Exit Function

    ' Insertion sort by PinNumber - pin counts per connector are small.
    For i = 2 To n
        For j = i To 2 Step -1
            If CDbl(vRows(j, PIN_COL_PINNUM)) < CDbl(vRows(j - 1, PIN_COL_PINNUM)) Then
                SwapPinRows vRows, j, j - 1
            Else
                Exit For
            End If
        Next j
    Next i

    ReDim vResult(1 To n, 1 To PIN_FIELD_COUNT)
    For i = 1 To n
        For j = 1 To PIN_FIELD_COUNT
            vResult(i, j) = vRows(i, j)
        Next j
    Next i

    ReadPinsForConnector = vResult
End Function

Public Function DeletePinsForConnector(wsPins As Worksheet, ByVal nFirstRow As Long, _
                                       ByVal nLastRow As Long, ByVal sConnectorID As String) As Long
    Dim nLast As Long, r As Long, w As Long, c As Long, nDeleted As Long

    nLast = LastUsedRowInWindow(wsPins, PIN_COL_CONNID, nLastRow)
    If nLast < nFirstRow Then Exit Function

    ' Single-pass compaction: copy every non-matching row down to a write
    ' cursor, then clear the leftover tail. Bounded to [nFirstRow, nLastRow]
    ' for the same reason DeleteConnector avoids Range.Delete Shift:=xlUp.
    w = nFirstRow
    For r = nFirstRow To nLast
        If StrComp(Trim$(CStr(wsPins.Cells(r, PIN_COL_CONNID).Value)), sConnectorID, vbTextCompare) = 0 Then
            nDeleted = nDeleted + 1
        Else
            If w <> r Then
                For c = 1 To PIN_FIELD_COUNT
                    wsPins.Cells(w, c).Value = wsPins.Cells(r, c).Value
                Next c
            End If
            w = w + 1
        End If
    Next r

    If w <= nLast Then
        wsPins.Range(wsPins.Cells(w, 1), wsPins.Cells(nLast, PIN_FIELD_COUNT)).ClearContents
    End If

    DeletePinsForConnector = nDeleted
End Function

Public Sub RemoveConnectorPhoto(wsPhotos As Worksheet, ByVal sConnectorID As String)
    On Error Resume Next
    wsPhotos.Shapes("PHOTO_" & sConnectorID).Delete
    On Error GoTo 0
End Sub

Public Function EmbedConnectorPhoto(wsPhotos As Worksheet, ByVal sConnectorID As String, _
                                    ByVal sImagePath As String) As String
    Dim sShapeName As String, nIndex As Long, nCol As Long, nRow As Long
    Dim shp As Shape

    If Len(Dir$(sImagePath)) = 0 Then Exit Function

    sShapeName = "PHOTO_" & sConnectorID
    RemoveConnectorPhoto wsPhotos, sConnectorID

    nIndex = wsPhotos.Shapes.Count
    nCol = nIndex Mod PHOTO_GRID_COLUMNS
    nRow = nIndex \ PHOTO_GRID_COLUMNS

    Set shp = wsPhotos.Shapes.AddPicture(sImagePath, False, True, _
        nCol * PHOTO_GRID_CELL_WIDTH, nRow * PHOTO_GRID_CELL_HEIGHT, _
        PHOTO_GRID_CELL_WIDTH - PHOTO_GRID_MARGIN, PHOTO_GRID_CELL_HEIGHT - PHOTO_GRID_MARGIN)
    shp.Name = sShapeName

    EmbedConnectorPhoto = sShapeName
End Function

Public Function CachePhotoPath(ByVal sWorkbookFolder As String, ByVal sConnectorID As String) As String
    Dim sFolder As String

    sFolder = sWorkbookFolder
    If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
    sFolder = sFolder & "Photos\"
    If Len(Dir$(sFolder, vbDirectory)) = 0 Then MkDir sFolder

    CachePhotoPath = sFolder & sConnectorID & ".png"
End Function

Public Function ExportShapeToFile(shp As Shape, ByVal sPath As String) As Boolean
    ' Excel has no direct "export a Shape to an image file" call. Pasting it
    ' into a throwaway ChartObject on the shape's own sheet and exporting
    ' the chart is the standard workaround - and, unlike Worksheet.Paste,
    ' Chart.Paste does not require the host sheet to be active, so this
    ' works even when shp's parent sheet is very hidden (_Snapshot, _Edit).
    Dim wsHost As Worksheet, cht As ChartObject

    Set wsHost = shp.Parent
    shp.Copy
    Set cht = wsHost.ChartObjects.Add(0, 0, shp.Width, shp.Height)
    cht.Chart.Paste
    cht.Chart.Export sPath, "PNG"
    cht.Delete

    ExportShapeToFile = (Len(Dir$(sPath)) > 0)
End Function
