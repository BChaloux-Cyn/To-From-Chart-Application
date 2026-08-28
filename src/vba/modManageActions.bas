Attribute VB_Name = "modManageActions"
Option Explicit

' Removing a connector everywhere it exists: the three library sheets plus
' the editor's on-disk preview cache, which would otherwise be orphaned -
' nothing ever reads a cache file whose connector is gone.
Public Function DeleteFromLibrary(wsLibConn As Worksheet, wsLibPins As Worksheet, _
                                  wsLibPhotos As Worksheet, ByVal sWorkbookPath As String, _
                                  ByVal sConnectorID As String) As Variant
    Dim sCachePath As String

    modLibrary.DeleteConnector wsLibConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID
    modLibrary.DeletePinsForConnector wsLibPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID
    modLibrary.RemoveConnectorPhoto wsLibPhotos, sConnectorID

    sCachePath = modLibrary.CachePhotoPath(sWorkbookPath, sConnectorID, "jpg")
    If Len(Dir$(sCachePath)) > 0 Then Kill sCachePath

    DeleteFromLibrary = modContract.Success("CONNECTOR_DELETED", sConnectorID)
End Function

' destWb is created by the adapter (Workbooks.Add) and saved by it
' afterwards; this shapes the sheets and copies the record into them.
Public Function ExportToWorkbook(wsLibConn As Worksheet, wsLibPins As Worksheet, _
                                 wsLibPhotos As Worksheet, destWb As Workbook, _
                                 ByVal sConnectorID As String) As Variant
    modLibraryTransfer.BuildExportSheets destWb

    If modLibraryTransfer.ExportConnector(wsLibConn, wsLibPins, wsLibPhotos, _
            destWb.Worksheets("Connectors"), destWb.Worksheets("Pins"), _
            destWb.Worksheets("Photos"), sConnectorID) Then
        ExportToWorkbook = modContract.Success("EXPORTED", sConnectorID)
    Else
        ExportToWorkbook = modContract.Failure("EXPORT_FAILED", sConnectorID)
    End If
End Function
