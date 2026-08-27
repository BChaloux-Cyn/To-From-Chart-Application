Attribute VB_Name = "modLibraryTransfer"
Option Explicit

Public Function CopyConnectorPhoto(wsSrcPhotos As Worksheet, wsDestPhotos As Worksheet, _
                                   ByVal sSrcConnectorID As String, ByVal sDestConnectorID As String) As Boolean
    On Error GoTo Failed
    modLibrary.RemoveConnectorPhoto wsDestPhotos, sDestConnectorID
    wsSrcPhotos.Shapes("PHOTO_" & sSrcConnectorID).Copy
    wsDestPhotos.Paste
    wsDestPhotos.Shapes(wsDestPhotos.Shapes.Count).Name = "PHOTO_" & sDestConnectorID
    CopyConnectorPhoto = True
    Exit Function
Failed:
    CopyConnectorPhoto = False
End Function

Public Function ExportConnector(wsSrcConn As Worksheet, wsSrcPins As Worksheet, wsSrcPhotos As Worksheet, _
                                wsDestConn As Worksheet, wsDestPins As Worksheet, wsDestPhotos As Worksheet, _
                                ByVal sConnectorID As String) As Boolean
    Dim vFields As Variant
    vFields = modLibrary.ReadConnector(wsSrcConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vFields) Then Exit Function
    If Not modLibrary.WriteConnector(wsDestConn, 2, modLibrary.LIB_ROW_CAP, vFields) Then Exit Function

    Dim vPins As Variant, i As Long, j As Long, vRow(1 To 7) As Variant
    vPins = modLibrary.ReadPinsForConnector(wsSrcPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If Not IsEmpty(vPins) Then
        For i = LBound(vPins, 1) To UBound(vPins, 1)
            For j = 1 To modLibrary.PIN_FIELD_COUNT
                vRow(j) = vPins(i, j)
            Next j
            modLibrary.WritePin wsDestPins, 2, modLibrary.LIB_ROW_CAP, vRow
        Next i
    End If

    ExportConnector = CopyConnectorPhoto(wsSrcPhotos, wsDestPhotos, sConnectorID, sConnectorID)
End Function
