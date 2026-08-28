Attribute VB_Name = "modEditorActions"
Option Explicit

' Restricted to JPG: LoadPicture's legacy OLE loader rejects valid PNGs
' with error 481 on some Windows/Office configurations, even though
' Shapes.AddPicture handles the same file. Offering only what LoadPicture
' opens is the fix, and keeping the string here makes it assertable.
Public Function PhotoFileFilter() As String
    PhotoFileFilter = "Pictures (*.jpg; *.jpeg), *.jpg;*.jpeg"
End Function

Public Function MarkerControlName(ByVal nPinNumber As Long) As String
    MarkerControlName = "lblMarker" & CStr(nPinNumber)
End Function

' Where a just-saved photo should be copied so the editor's preview can
' read it back without a clipboard round trip. Empty when nothing needs
' copying: no photo was chosen, or the chosen file already is the cache
' (FileCopy onto itself raises).
Public Function PhotoCacheRefreshTarget(ByVal sWorkbookPath As String, _
                                        ByVal sConnectorID As String, _
                                        ByVal sPhotoPath As String) As String
    Dim sCachePath As String

    If Len(Trim$(sPhotoPath)) = 0 Then Exit Function

    sCachePath = modLibrary.CachePhotoPath(sWorkbookPath, sConnectorID, "jpg")
    If StrComp(sPhotoPath, sCachePath, vbTextCompare) = 0 Then Exit Function

    PhotoCacheRefreshTarget = sCachePath
End Function

' The connector Type options, read from the sheet passed in rather than
' through an unqualified RowSource, which resolved against ActiveWorkbook
' and left the combo empty during frmManageLibrary's Edit flow.
Public Function TypeListItems(wsLists As Worksheet) As Variant
    Dim vItems() As String, r As Long, n As Long

    r = 2
    Do While Len(Trim$(CStr(wsLists.Cells(r, 4).Value))) > 0
        n = n + 1
        ReDim Preserve vItems(1 To n)
        vItems(n) = CStr(wsLists.Cells(r, 4).Value)
        r = r + 1
    Loop

    If n = 0 Then Exit Function
    TypeListItems = vItems
End Function

' One row per placed pin: the display string the list box shows, the pin
' number that row resolves to, and the marker's normalized position.
' Derived from the scratch sheet on every call, which is what lets the
' form drop the mListPinNumbers collection that used to desync from it.
Public Function PinListItems(wsScratch As Worksheet, ByVal sConnectorID As String) As Variant
    Dim vPins As Variant, i As Long, n As Long
    Dim vRows() As Variant

    vPins = modLibrary.ReadPinsForConnector(wsScratch, modPinEditor.SCRATCH_FIRST_ROW, _
                                            modPinEditor.SCRATCH_LAST_ROW, sConnectorID)
    If IsEmpty(vPins) Then Exit Function

    n = UBound(vPins, 1) - LBound(vPins, 1) + 1
    ReDim vRows(1 To n, 1 To 4)
    For i = 1 To n
        vRows(i, 1) = "Pin " & CStr(CLng(vPins(LBound(vPins, 1) + i - 1, 2)))
        vRows(i, 2) = CLng(vPins(LBound(vPins, 1) + i - 1, 2))
        vRows(i, 3) = CDbl(vPins(LBound(vPins, 1) + i - 1, 6))
        vRows(i, 4) = CDbl(vPins(LBound(vPins, 1) + i - 1, 7))
    Next i

    PinListItems = vRows
End Function

' One past the highest placed pin number, so a deletion never causes a
' reused number. Replaces the form's mNextPinNumber counter.
Public Function NextPinNumber(wsScratch As Worksheet, ByVal sConnectorID As String) As Long
    Dim vItems As Variant, i As Long, nMax As Long

    vItems = PinListItems(wsScratch, sConnectorID)
    If IsEmpty(vItems) Then
        NextPinNumber = 1
        Exit Function
    End If

    For i = LBound(vItems, 1) To UBound(vItems, 1)
        If CLng(vItems(i, 2)) > nMax Then nMax = CLng(vItems(i, 2))
    Next i

    NextPinNumber = nMax + 1
End Function
