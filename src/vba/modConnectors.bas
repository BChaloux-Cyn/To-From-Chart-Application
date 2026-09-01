Attribute VB_Name = "modConnectors"
Option Explicit

Public Const CONN_SHEET As String = "Connectors"
Public Const CONN_FIRST_ROW As Long = 2

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

Public Function PrefixForType(ByVal sType As String) As String
    Select Case LCase$(Trim$(sType))
        Case "connector": PrefixForType = "J"
        Case "stud":      PrefixForType = "ST"
        Case "splice":    PrefixForType = "SP"
        Case "tail":      PrefixForType = "TL"
        Case Else:        PrefixForType = ""
    End Select
End Function

Public Function NextRefDes(ByVal sPrefix As String) As String
    Dim ws As Worksheet
    Dim r As Long, nLast As Long, nMax As Long, nNum As Long
    Dim sVal As String, sTail As String, sUpper As String

    If Len(sPrefix) = 0 Then Exit Function

    sUpper = UCase$(sPrefix)
    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = CONN_FIRST_ROW To nLast
        sVal = UCase$(Trim$(CStr(ws.Cells(r, 1).Value)))
        If Left$(sVal, Len(sUpper)) = sUpper Then
            sTail = Mid$(sVal, Len(sUpper) + 1)
            If IsAllDigits(sTail) Then
                nNum = CLng(sTail)
                If nNum > nMax Then nMax = nNum
            End If
        End If
    Next r

    NextRefDes = sUpper & CStr(nMax + 1)
End Function

Public Function AddConnectorInstance(ByVal sConnectorID As String, _
                                     ByVal sName As String, _
                                     ByVal sPartNumber As String, _
                                     ByVal sType As String, _
                                     ByVal nPinCount As Long) As String
    Dim ws As Worksheet
    Dim sPrefix As String, sRef As String
    Dim r As Long

    sPrefix = PrefixForType(sType)
    If Len(sPrefix) = 0 Then Exit Function
    If nPinCount < 1 Then Exit Function

    sRef = NextRefDes(sPrefix)
    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < CONN_FIRST_ROW Then r = CONN_FIRST_ROW

    ws.Cells(r, 1).Value = sRef
    ws.Cells(r, 2).Value = sConnectorID
    ws.Cells(r, 3).Value = sName
    ws.Cells(r, 4).Value = sPartNumber
    ws.Cells(r, 5).Value = Trim$(sType)
    ws.Cells(r, 6).Value = nPinCount

    modState.MarkDirty
    AddConnectorInstance = sRef
End Function

Public Function PinCountFor(ByVal sRefDes As String) As Long
    Dim ws As Worksheet
    Dim r As Long, nLast As Long

    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = CONN_FIRST_ROW To nLast
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), Trim$(sRefDes), vbTextCompare) = 0 Then
            PinCountFor = CLng(Val(ws.Cells(r, 6).Value))
            Exit Function
        End If
    Next r

    PinCountFor = 0
End Function

Public Function RenameRefDes(ByVal sOldRefDes As String, ByVal sNewRefDes As String) As Boolean
    Dim ws As Worksheet, wsChart As Worksheet
    Dim r As Long, nLast As Long, nMatches As Long

    sOldRefDes = Trim$(sOldRefDes)
    sNewRefDes = Trim$(sNewRefDes)
    If Len(sOldRefDes) = 0 Or Len(sNewRefDes) = 0 Then Exit Function
    If StrComp(sOldRefDes, sNewRefDes, vbTextCompare) = 0 Then Exit Function

    ' The renamed row already carries sNewRefDes by the time this runs (the
    ' sheet edit happens before Worksheet_Change fires), so exactly one
    ' match is the non-colliding case; more than one means a different row
    ' already used that ref des.
    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = CONN_FIRST_ROW To nLast
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), sNewRefDes, vbTextCompare) = 0 Then
            nMatches = nMatches + 1
        End If
    Next r
    If nMatches <> 1 Then Exit Function

    Set wsChart = ThisWorkbook.Worksheets(modChart.CHART_SHEET)
    For r = modChart.CHART_FIRST_ROW To modChart.CHART_LAST_ROW
        If StrComp(Trim$(CStr(wsChart.Cells(r, modChart.COL_FROM_CONN).Value)), sOldRefDes, vbTextCompare) = 0 Then
            wsChart.Cells(r, modChart.COL_FROM_CONN).Value = sNewRefDes
        End If
        If StrComp(Trim$(CStr(wsChart.Cells(r, modChart.COL_TO_CONN).Value)), sOldRefDes, vbTextCompare) = 0 Then
            wsChart.Cells(r, modChart.COL_TO_CONN).Value = sNewRefDes
        End If
    Next r

    RenameRefDes = True
End Function

' Display strings for every placed connector instance, for the Remove
' Connector picker - "<RefDes> - <Name>" paired with the bare RefDes,
' matching modLibrary.ConnectorIndex's shape for the library picker.
Public Function InstanceIndex() As Variant
    Dim ws As Worksheet
    Dim r As Long, nLast As Long, n As Long
    Dim vRows() As Variant

    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If nLast < CONN_FIRST_ROW Then Exit Function

    ReDim vRows(1 To nLast - CONN_FIRST_ROW + 1, 1 To 2)
    For r = CONN_FIRST_ROW To nLast
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0 Then
            n = n + 1
            vRows(n, 1) = Trim$(CStr(ws.Cells(r, 1).Value)) & " - " & CStr(ws.Cells(r, 3).Value)
            vRows(n, 2) = Trim$(CStr(ws.Cells(r, 1).Value))
        End If
    Next r
    If n = 0 Then Exit Function

    Dim vResult() As Variant, i As Long
    ReDim vResult(1 To n, 1 To 2)
    For i = 1 To n
        vResult(i, 1) = vRows(i, 1)
        vResult(i, 2) = vRows(i, 2)
    Next i
    InstanceIndex = vResult
End Function

Public Function RemoveConnectorInstance(ByVal sRefDes As String) As Boolean
    Dim ws As Worksheet, wsChart As Worksheet
    Dim r As Long, nLast As Long, c As Long
    Dim bEvents As Boolean

    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    r = 0
    Dim i As Long
    For i = CONN_FIRST_ROW To nLast
        If StrComp(Trim$(CStr(ws.Cells(i, 1).Value)), sRefDes, vbTextCompare) = 0 Then
            r = i
            Exit For
        End If
    Next i
    If r = 0 Then Exit Function

    ' The row-compaction write below touches column A of the Connectors
    ' sheet, which shConnectors.evt watches for ref-des renames - without
    ' suppressing events here, that handler sees the swapped-in RefDes as a
    ' "rename" of the just-removed row, finds it collides with the row it
    ' was copied from (not yet cleared), and reverts the write mid-removal.
    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    If r < nLast Then
        For c = 1 To 6
            ws.Cells(r, c).Value = ws.Cells(nLast, c).Value
        Next c
    End If
    ws.Range(ws.Cells(nLast, 1), ws.Cells(nLast, 6)).ClearContents

    Set wsChart = ThisWorkbook.Worksheets(modChart.CHART_SHEET)
    For i = modChart.CHART_FIRST_ROW To modChart.CHART_LAST_ROW
        If StrComp(Trim$(CStr(wsChart.Cells(i, modChart.COL_FROM_CONN).Value)), sRefDes, vbTextCompare) = 0 Then
            wsChart.Cells(i, modChart.COL_FROM_CONN).ClearContents
            wsChart.Cells(i, modChart.COL_FROM_PIN).Validation.Delete
            wsChart.Cells(i, modChart.COL_FROM_PIN).ClearContents
        End If
        If StrComp(Trim$(CStr(wsChart.Cells(i, modChart.COL_TO_CONN).Value)), sRefDes, vbTextCompare) = 0 Then
            wsChart.Cells(i, modChart.COL_TO_CONN).ClearContents
            wsChart.Cells(i, modChart.COL_TO_PIN).Validation.Delete
            wsChart.Cells(i, modChart.COL_TO_PIN).ClearContents
        End If
    Next i

    RemoveConnectorInstance = True

CleanUp:
    Application.EnableEvents = bEvents
End Function

' Removes every placed instance of a connector definition, for when that
' definition is deleted from the library out from under them. Ref des are
' collected first: RemoveConnectorInstance compacts the sheet by row, so
' removing by value in a fresh scan each time (rather than by row index)
' is what stays correct as rows shift underneath the loop.
Public Function RemoveInstancesOfConnectorType(ByVal sConnectorID As String) As Variant
    Dim ws As Worksheet
    Dim r As Long, nLast As Long, n As Long
    Dim vMatches() As String

    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ReDim vMatches(1 To nLast)
    For r = CONN_FIRST_ROW To nLast
        If StrComp(Trim$(CStr(ws.Cells(r, 2).Value)), sConnectorID, vbTextCompare) = 0 Then
            n = n + 1
            vMatches(n) = Trim$(CStr(ws.Cells(r, 1).Value))
        End If
    Next r
    If n = 0 Then Exit Function

    Dim i As Long
    For i = 1 To n
        RemoveConnectorInstance vMatches(i)
    Next i

    Dim vResult() As String
    ReDim vResult(1 To n)
    For i = 1 To n
        vResult(i) = vMatches(i)
    Next i
    RemoveInstancesOfConnectorType = vResult
End Function

Public Function AllInstances() As Variant
    Dim ws As Worksheet
    Dim r As Long, nLast As Long, n As Long, c As Long
    Dim vResult() As Variant

    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If nLast < CONN_FIRST_ROW Then Exit Function

    n = nLast - CONN_FIRST_ROW + 1
    ReDim vResult(1 To n, 1 To 6)
    For r = CONN_FIRST_ROW To nLast
        For c = 1 To 6
            vResult(r - CONN_FIRST_ROW + 1, c) = ws.Cells(r, c).Value
        Next c
    Next r

    AllInstances = vResult
End Function

' Everything shConnectors's Worksheet_Change decides. sPriorRefDes and
' nPriorRow come from the sheet module's SelectionChange bookkeeping, which
' is genuine event-lifecycle state and stays there. On rejection the
' payload is the value the caller must write back into the cell.
Public Function ApplyConnectorEdit(rTarget As Range, ByVal sPriorRefDes As String, _
                                   ByVal nPriorRow As Long) As Variant
    Dim rw As Range, sRef As String, sNewRefDes As String

    modState.MarkDirty

    For Each rw In rTarget.Rows
        If rw.Row >= CONN_FIRST_ROW Then
            sRef = Trim$(CStr(rTarget.Worksheet.Cells(rw.Row, 1).Value))
            modChart.RefreshChartRowsForConnector sRef
        End If
    Next rw

    If Not (rTarget.Cells.Count = 1 And rTarget.Column = 1 _
            And rTarget.Row = nPriorRow And Len(sPriorRefDes) > 0) Then
        ApplyConnectorEdit = modContract.Failure("NO_RENAME")
        Exit Function
    End If

    sNewRefDes = Trim$(CStr(rTarget.Value))
    If StrComp(sNewRefDes, sPriorRefDes, vbTextCompare) = 0 Then
        ApplyConnectorEdit = modContract.Failure("NO_RENAME")
        Exit Function
    End If

    If RenameRefDes(sPriorRefDes, sNewRefDes) Then
        ApplyConnectorEdit = modContract.Success("RENAMED", sNewRefDes)
    Else
        ApplyConnectorEdit = modContract.Failure("RENAME_REJECTED", sPriorRefDes)
    End If
End Function
