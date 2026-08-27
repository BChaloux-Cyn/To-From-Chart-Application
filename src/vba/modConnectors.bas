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

Public Function RemoveConnectorInstance(ByVal sRefDes As String) As Boolean
    Dim ws As Worksheet, wsChart As Worksheet
    Dim r As Long, nLast As Long, c As Long

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
End Function
