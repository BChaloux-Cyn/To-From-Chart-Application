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

' mConnectorID is derived from Name and Part Number when the photo loads
' and never recomputed, so placing pins before both are filled in used to
' do nothing at all, silently. The guard runs before the file picker opens.
Public Function CanLoadPhoto(ByVal sName As String, ByVal sPartNumber As String) As Variant
    If Len(Trim$(sName)) = 0 Or Len(Trim$(sPartNumber)) = 0 Then
        CanLoadPhoto = modContract.Failure("MISSING_NAME_OR_PART")
        Exit Function
    End If
    CanLoadPhoto = modContract.Success("OK")
End Function

' Where the editor's photo preview should read from, and whether a one-time
' backfill from the embedded Shape is needed first. The on-disk cache is
' preferred because re-exporting the Shape goes through the clipboard,
' which is unreliable for VBA-triggered operations on this machine.
Public Function PhotoSourceForEdit(ByVal sWorkbookPath As String, _
                                   ByVal sConnectorID As String) As Variant
    Dim sCachePath As String
    sCachePath = modLibrary.CachePhotoPath(sWorkbookPath, sConnectorID, "jpg")

    If Len(Dir$(sCachePath)) = 0 Then
        PhotoSourceForEdit = modContract.Failure("NEEDS_BACKFILL", sCachePath)
        Exit Function
    End If

    PhotoSourceForEdit = modContract.Success("CACHE_READY", sCachePath)
End Function

Public Function DeletePinRequest(wsScratch As Worksheet, ByVal sConnectorID As String, _
                                 ByVal nPinNumber As Long) As Variant
    If modPinEditor.RemovePin(wsScratch, sConnectorID, nPinNumber) Then
        DeletePinRequest = modContract.Success("PIN_DELETED", nPinNumber)
    Else
        DeletePinRequest = modContract.Failure("PIN_NOT_FOUND", nPinNumber)
    End If
End Function

' Everything a click on the photo can mean. bPlaceMode is the Place Pins
' toggle; nSelectedPin is a PIN NUMBER (0 when nothing is selected), not a
' list index - the caller resolves that through PinListItems. The placed
' count and next pin number are derived from wsScratch rather than passed
' in, so no counter in the form can drift out of sync with the sheet.
Public Function PhotoClickAction(wsScratch As Worksheet, ByVal sConnectorID As String, _
                                 ByVal bPlaceMode As Boolean, ByVal nSelectedPin As Long, _
                                 ByVal sPinCountText As String, _
                                 ByVal dNormX As Double, ByVal dNormY As Double) As Variant
    If Len(Trim$(sConnectorID)) = 0 Then
        PhotoClickAction = modContract.Failure("NO_OP")
        Exit Function
    End If

    If bPlaceMode Then
        Dim nPinCount As Long, nPlaced As Long, nNext As Long

        If Not IsNumeric(Trim$(sPinCountText)) Then
            PhotoClickAction = modContract.Failure("BAD_PIN_COUNT")
            Exit Function
        End If
        nPinCount = CLng(Val(sPinCountText))
        If nPinCount <= 0 Then
            PhotoClickAction = modContract.Failure("BAD_PIN_COUNT")
            Exit Function
        End If

        nPlaced = modContract.TableRowCount(PinListItems(wsScratch, sConnectorID))
        If nPlaced >= nPinCount Then
            PhotoClickAction = modContract.Failure("PIN_LIMIT_REACHED", nPinCount)
            Exit Function
        End If

        nNext = NextPinNumber(wsScratch, sConnectorID)
        If modPinEditor.PlacePin(wsScratch, sConnectorID, nNext, _
                                 "Pin " & CStr(nNext), dNormX, dNormY) Then
            PhotoClickAction = modContract.Success("PLACED", nNext)
        Else
            PhotoClickAction = modContract.Failure("NO_OP")
        End If
        Exit Function
    End If

    ' Place Pins is off and a pin is selected: the click moves that pin's
    ' anchor. modPinEditor.MoveAnchor decides whether the marker travels
    ' with it (it does only if it was still sitting on the anchor).
    If nSelectedPin > 0 Then
        If modPinEditor.MoveAnchor(wsScratch, sConnectorID, nSelectedPin, dNormX, dNormY) Then
            PhotoClickAction = modContract.Success("MOVED_ANCHOR", nSelectedPin)
            Exit Function
        End If
    End If

    PhotoClickAction = modContract.Failure("NO_OP")
End Function

' The whole Save transaction bar the workbook open, close, and photo copy,
' which stay in the adapter. vFields is zero based: name, manufacturer,
' part number, type, pin count AS TEXT, notes - the pin count arrives as
' the text box supplies it so the coercion is tested here.
' sNowUtc is passed in rather than read from Now, which keeps this
' deterministic and lets a test assert the exact timestamp written.
Public Function SaveFromEditor(wsLibConn As Worksheet, wsLibPins As Worksheet, _
                               wsLibPhotos As Worksheet, wsScratch As Worksheet, _
                               ByVal sConnectorID As String, ByVal sOriginalID As String, _
                               ByVal vFields As Variant, ByVal sPhotoPath As String, _
                               ByVal sNowUtc As String) As Variant
    Dim nExistingRow As Long

    ' A collision only matters when the row it would overwrite belongs to a
    ' DIFFERENT connector than the one this session opened for editing -
    ' re-saving the connector you are editing must not flag itself.
    nExistingRow = modLibrary.FindConnectorRow(wsLibConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If nExistingRow > 0 And StrComp(sConnectorID, sOriginalID, vbTextCompare) <> 0 Then
        SaveFromEditor = modContract.Failure("ID_COLLISION", sConnectorID)
        Exit Function
    End If

    If modPinEditor.SaveConnector(wsLibConn, wsLibPins, wsLibPhotos, wsScratch, _
            sConnectorID, _
            CStr(vFields(LBound(vFields))), _
            CStr(vFields(LBound(vFields) + 1)), _
            CStr(vFields(LBound(vFields) + 2)), _
            CStr(vFields(LBound(vFields) + 3)), _
            CLng(Val(CStr(vFields(LBound(vFields) + 4)))), _
            CStr(vFields(LBound(vFields) + 5)), _
            sPhotoPath, sNowUtc, sNowUtc, "Local") Then
        SaveFromEditor = modContract.Success("SAVED", sConnectorID)
    Else
        SaveFromEditor = modContract.Failure("SAVE_FAILED", sConnectorID)
    End If
End Function
