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

Public Function ImportConnector(wsSrcConn As Worksheet, wsSrcPins As Worksheet, wsSrcPhotos As Worksheet, _
                                wsDestConn As Worksheet, wsDestPins As Worksheet, wsDestPhotos As Worksheet, _
                                ByVal sConnectorID As String, ByVal sOriginFileName As String) As String
    Dim vFields As Variant
    vFields = modLibrary.ReadConnector(wsSrcConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vFields) Then Exit Function

    Dim sDestID As String
    sDestID = modLibrary.UniqueConnectorID(wsDestConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)

    vFields(modLibrary.LIB_COL_ID) = sDestID
    vFields(modLibrary.LIB_COL_PHOTOSHAPE) = "PHOTO_" & sDestID
    vFields(modLibrary.LIB_COL_ORIGIN) = sOriginFileName
    If Not modLibrary.WriteConnector(wsDestConn, 2, modLibrary.LIB_ROW_CAP, vFields) Then Exit Function

    Dim vPins As Variant, i As Long, j As Long, vRow(1 To 7) As Variant
    vPins = modLibrary.ReadPinsForConnector(wsSrcPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If Not IsEmpty(vPins) Then
        For i = LBound(vPins, 1) To UBound(vPins, 1)
            For j = 1 To modLibrary.PIN_FIELD_COUNT
                vRow(j) = vPins(i, j)
            Next j
            vRow(1) = sDestID
            modLibrary.WritePin wsDestPins, 2, modLibrary.LIB_ROW_CAP, vRow
        Next i
    End If

    CopyConnectorPhoto wsSrcPhotos, wsDestPhotos, sConnectorID, sDestID

    ImportConnector = sDestID
End Function

Public Sub BuildExportSheets(wb As Workbook)
    Dim names As Variant, i As Long, sheet As Worksheet, original As Worksheet
    names = Array("Connectors", "Pins", "Photos")

    Set original = wb.Worksheets(1)
    For i = LBound(names) To UBound(names)
        Set sheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        sheet.Name = CStr(names(i))
    Next i

    ' Worksheet.Delete always shows a "permanently remove" confirmation
    ' unless DisplayAlerts is off - original is the new workbook's own
    ' empty default sheet, nothing the user asked to keep.
    Dim bPriorAlerts As Boolean
    bPriorAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    original.Delete
    Application.DisplayAlerts = bPriorAlerts

    Dim connHeaders As Variant, pinHeaders As Variant, c As Long
    connHeaders = Array("ConnectorID", "Name", "Manufacturer", "PartNumber", "Type", _
                         "PinCount", "Notes", "PhotoShapeName", "CreatedUtc", "ModifiedUtc", "Origin")
    pinHeaders = Array("ConnectorID", "PinNumber", "PinLabel", "NormX", "NormY", "LabelX", "LabelY")

    For c = LBound(connHeaders) To UBound(connHeaders)
        wb.Worksheets("Connectors").Cells(1, c + 1).Value = connHeaders(c)
    Next c
    For c = LBound(pinHeaders) To UBound(pinHeaders)
        wb.Worksheets("Pins").Cells(1, c + 1).Value = pinHeaders(c)
    Next c
End Sub
