Attribute VB_Name = "modPinEditor"
Option Explicit

Public Const SCRATCH_FIRST_ROW As Long = 2
Public Const SCRATCH_LAST_ROW As Long = 2000

' A normalized-distance tolerance for the editor's own live preview only -
' see MarkerSitsOnAnchor. Declared here, with the other module-level
' consts, not next to its first use: VBA requires every module-level Const
' to sit in the general-declarations section before the first Sub/Function,
' not merely before whichever procedure first references it.
Public Const PREVIEW_LEADER_THRESHOLD As Double = 0.01

Public Sub ClearScratchPins(wsScratch As Worksheet)
    wsScratch.Range(wsScratch.Cells(SCRATCH_FIRST_ROW, 1), _
                    wsScratch.Cells(SCRATCH_LAST_ROW, modLibrary.PIN_FIELD_COUNT)).ClearContents
End Sub

Public Function LoadScratchPins(wsScratch As Worksheet, wsLibPins As Worksheet, _
                                ByVal sConnectorID As String) As Long
    Dim vPins As Variant, i As Long, j As Long, vRow(1 To 7) As Variant

    ClearScratchPins wsScratch
    vPins = modLibrary.ReadPinsForConnector(wsLibPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vPins) Then Exit Function

    For i = LBound(vPins, 1) To UBound(vPins, 1)
        For j = 1 To modLibrary.PIN_FIELD_COUNT
            vRow(j) = vPins(i, j)
        Next j
        modLibrary.WritePin wsScratch, SCRATCH_FIRST_ROW, SCRATCH_LAST_ROW, vRow
    Next i

    LoadScratchPins = UBound(vPins, 1) - LBound(vPins, 1) + 1
End Function

Private Function FindPinRow(wsScratch As Worksheet, ByVal sConnectorID As String, _
                            ByVal nPinNumber As Long) As Long
    Dim r As Long, nLast As Long

    nLast = modLibrary.LastUsedRowInWindow(wsScratch, modLibrary.PIN_COL_CONNID, SCRATCH_LAST_ROW)
    If nLast < SCRATCH_FIRST_ROW Then Exit Function

    For r = SCRATCH_FIRST_ROW To nLast
        If StrComp(Trim$(CStr(wsScratch.Cells(r, modLibrary.PIN_COL_CONNID).Value)), sConnectorID, vbTextCompare) = 0 _
           And CLng(wsScratch.Cells(r, modLibrary.PIN_COL_PINNUM).Value) = nPinNumber Then
            FindPinRow = r
            Exit Function
        End If
    Next r
End Function

Public Function PlacePin(wsScratch As Worksheet, ByVal sConnectorID As String, _
                         ByVal nPinNumber As Long, ByVal sLabel As String, _
                         ByVal dNormX As Double, ByVal dNormY As Double) As Boolean
    ' A fresh placement: anchor and marker start identical - the marker
    ' sits directly on the point until a student drags it away.
    Dim vFields As Variant

    RemovePin wsScratch, sConnectorID, nPinNumber
    vFields = Array(sConnectorID, nPinNumber, sLabel, dNormX, dNormY, dNormX, dNormY)
    PlacePin = modLibrary.WritePin(wsScratch, SCRATCH_FIRST_ROW, SCRATCH_LAST_ROW, vFields)
End Function

Public Function RemovePin(wsScratch As Worksheet, ByVal sConnectorID As String, _
                          ByVal nPinNumber As Long) As Boolean
    Dim r As Long, nLast As Long, c As Long

    r = FindPinRow(wsScratch, sConnectorID, nPinNumber)
    If r = 0 Then Exit Function

    nLast = modLibrary.LastUsedRowInWindow(wsScratch, modLibrary.PIN_COL_CONNID, SCRATCH_LAST_ROW)
    If r < nLast Then
        For c = 1 To modLibrary.PIN_FIELD_COUNT
            wsScratch.Cells(r, c).Value = wsScratch.Cells(nLast, c).Value
        Next c
    End If
    wsScratch.Range(wsScratch.Cells(nLast, 1), wsScratch.Cells(nLast, modLibrary.PIN_FIELD_COUNT)).ClearContents

    RemovePin = True
End Function

Public Function MarkerSitsOnAnchor(ByVal dAnchorX As Double, ByVal dAnchorY As Double, _
                                   ByVal dLabelX As Double, ByVal dLabelY As Double) As Boolean
    Dim dDx As Double, dDy As Double
    dDx = dLabelX - dAnchorX
    dDy = dLabelY - dAnchorY
    MarkerSitsOnAnchor = (Sqr(dDx * dDx + dDy * dDy) <= PREVIEW_LEADER_THRESHOLD)
End Function

Public Function MoveAnchor(wsScratch As Worksheet, ByVal sConnectorID As String, _
                           ByVal nPinNumber As Long, ByVal dNormX As Double, ByVal dNormY As Double) As Boolean
    Dim r As Long, dOldAnchorX As Double, dOldAnchorY As Double, dLabelX As Double, dLabelY As Double

    r = FindPinRow(wsScratch, sConnectorID, nPinNumber)
    If r = 0 Then Exit Function

    dOldAnchorX = CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_NORMX).Value)
    dOldAnchorY = CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_NORMY).Value)
    dLabelX = CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_LABELX).Value)
    dLabelY = CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_LABELY).Value)

    If MarkerSitsOnAnchor(dOldAnchorX, dOldAnchorY, dLabelX, dLabelY) Then
        wsScratch.Cells(r, modLibrary.PIN_COL_LABELX).Value = dNormX
        wsScratch.Cells(r, modLibrary.PIN_COL_LABELY).Value = dNormY
    End If
    wsScratch.Cells(r, modLibrary.PIN_COL_NORMX).Value = dNormX
    wsScratch.Cells(r, modLibrary.PIN_COL_NORMY).Value = dNormY

    MoveAnchor = True
End Function

Public Function MoveMarker(wsScratch As Worksheet, ByVal sConnectorID As String, _
                           ByVal nPinNumber As Long, ByVal dNormX As Double, ByVal dNormY As Double) As Boolean
    Dim r As Long
    r = FindPinRow(wsScratch, sConnectorID, nPinNumber)
    If r = 0 Then Exit Function

    wsScratch.Cells(r, modLibrary.PIN_COL_LABELX).Value = dNormX
    wsScratch.Cells(r, modLibrary.PIN_COL_LABELY).Value = dNormY

    MoveMarker = True
End Function

Public Function SnapLabelToPin(wsScratch As Worksheet, ByVal sConnectorID As String, _
                               ByVal nPinNumber As Long) As Boolean
    Dim r As Long
    r = FindPinRow(wsScratch, sConnectorID, nPinNumber)
    If r = 0 Then Exit Function

    wsScratch.Cells(r, modLibrary.PIN_COL_LABELX).Value = wsScratch.Cells(r, modLibrary.PIN_COL_NORMX).Value
    wsScratch.Cells(r, modLibrary.PIN_COL_LABELY).Value = wsScratch.Cells(r, modLibrary.PIN_COL_NORMY).Value

    SnapLabelToPin = True
End Function

Public Function NeedsLeaderLine(wsScratch As Worksheet, ByVal sConnectorID As String, _
                                ByVal nPinNumber As Long) As Boolean
    Dim r As Long
    r = FindPinRow(wsScratch, sConnectorID, nPinNumber)
    If r = 0 Then Exit Function

    NeedsLeaderLine = Not MarkerSitsOnAnchor( _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_NORMX).Value), _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_NORMY).Value), _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_LABELX).Value), _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_LABELY).Value))
End Function

' Array(anchorNormX, anchorNormY, labelNormX, labelNormY), or Empty if the
' pin isn't found. Lets a caller (the leader-line drawing code) place both
' endpoints without duplicating FindPinRow/column lookups itself.
Public Function PinGeometry(wsScratch As Worksheet, ByVal sConnectorID As String, _
                            ByVal nPinNumber As Long) As Variant
    Dim r As Long
    r = FindPinRow(wsScratch, sConnectorID, nPinNumber)
    If r = 0 Then Exit Function

    PinGeometry = Array( _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_NORMX).Value), _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_NORMY).Value), _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_LABELX).Value), _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_LABELY).Value))
End Function

Public Function FitAspectRatio(ByVal dSourceWidth As Double, ByVal dSourceHeight As Double, _
                               ByVal dBoxWidth As Double, ByVal dBoxHeight As Double) As Variant
    Dim dSourceRatio As Double, dBoxRatio As Double, dOutWidth As Double, dOutHeight As Double

    If dSourceWidth <= 0 Or dSourceHeight <= 0 Or dBoxWidth <= 0 Or dBoxHeight <= 0 Then Exit Function

    dSourceRatio = dSourceWidth / dSourceHeight
    dBoxRatio = dBoxWidth / dBoxHeight

    If dSourceRatio > dBoxRatio Then
        dOutWidth = dBoxWidth
        dOutHeight = dBoxWidth / dSourceRatio
    Else
        dOutHeight = dBoxHeight
        dOutWidth = dBoxHeight * dSourceRatio
    End If

    FitAspectRatio = Array(dOutWidth, dOutHeight)
End Function

' The on-screen top-left for a marker badge whose centre sits on the
' normalized point (dNormX, dNormY) within the photo control's box.
' Exact inverse of NormFromMarker.
Public Function MarkerTopLeft(ByVal dNormX As Double, ByVal dNormY As Double, _
                              ByVal dPhotoLeft As Double, ByVal dPhotoTop As Double, _
                              ByVal dPhotoW As Double, ByVal dPhotoH As Double, _
                              ByVal dMarkerW As Double, ByVal dMarkerH As Double) As Variant
    MarkerTopLeft = Array( _
        dPhotoLeft + dNormX * dPhotoW - (dMarkerW / 2), _
        dPhotoTop + dNormY * dPhotoH - (dMarkerH / 2))
End Function

' The normalized point under a marker badge's centre. Exact inverse of
' MarkerTopLeft. Empty when the photo box has no area to normalize against.
Public Function NormFromMarker(ByVal dLeft As Double, ByVal dTop As Double, _
                               ByVal dMarkerW As Double, ByVal dMarkerH As Double, _
                               ByVal dPhotoLeft As Double, ByVal dPhotoTop As Double, _
                               ByVal dPhotoW As Double, ByVal dPhotoH As Double) As Variant
    If dPhotoW <= 0 Or dPhotoH <= 0 Then Exit Function

    NormFromMarker = Array( _
        (dLeft + dMarkerW / 2 - dPhotoLeft) / dPhotoW, _
        (dTop + dMarkerH / 2 - dPhotoTop) / dPhotoH)
End Function

Public Function SaveConnector(wsLibConn As Worksheet, wsLibPins As Worksheet, wsLibPhotos As Worksheet, _
                              wsScratch As Worksheet, ByVal sConnectorID As String, ByVal sName As String, _
                              ByVal sManufacturer As String, ByVal sPartNumber As String, ByVal sType As String, _
                              ByVal nPinCount As Long, ByVal sNotes As String, ByVal sPhotoPath As String, _
                              ByVal sCreatedUtc As String, ByVal sModifiedUtc As String, ByVal sOrigin As String) As Boolean
    Dim sShapeName As String, vPins As Variant, vFields As Variant, i As Long

    sShapeName = modLibrary.EmbedConnectorPhoto(wsLibPhotos, sConnectorID, sPhotoPath)
    If Len(sShapeName) = 0 Then Exit Function

    vFields = Array(sConnectorID, sName, sManufacturer, sPartNumber, sType, nPinCount, _
                     sNotes, sShapeName, sCreatedUtc, sModifiedUtc, sOrigin)
    If Not modLibrary.WriteConnector(wsLibConn, 2, modLibrary.LIB_ROW_CAP, vFields) Then Exit Function

    modLibrary.DeletePinsForConnector wsLibPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID
    vPins = modLibrary.ReadPinsForConnector(wsScratch, SCRATCH_FIRST_ROW, SCRATCH_LAST_ROW, sConnectorID)
    If Not IsEmpty(vPins) Then
        Dim vRow(1 To 7) As Variant, j As Long
        For i = LBound(vPins, 1) To UBound(vPins, 1)
            For j = 1 To modLibrary.PIN_FIELD_COUNT
                vRow(j) = vPins(i, j)
            Next j
            modLibrary.WritePin wsLibPins, 2, modLibrary.LIB_ROW_CAP, vRow
        Next i
    End If

    SaveConnector = True
End Function
