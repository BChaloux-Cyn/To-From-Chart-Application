Attribute VB_Name = "modManageActions"
Option Explicit

' Removing a connector everywhere it exists: the three library sheets, the
' editor's on-disk preview cache (which would otherwise be orphaned -
' nothing ever reads a cache file whose connector is gone), and any
' instance of it already placed on the currently open harness's chart,
' which would otherwise keep referencing a ConnectorID the library no
' longer has.
Public Function DeleteFromLibrary(wsLibConn As Worksheet, wsLibPins As Worksheet, _
                                  wsLibPhotos As Worksheet, ByVal sWorkbookPath As String, _
                                  ByVal sConnectorID As String) As Variant
    Dim sCachePath As String, vRemoved As Variant

    modLibrary.DeleteConnector wsLibConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID
    modLibrary.DeletePinsForConnector wsLibPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID
    modLibrary.RemoveConnectorPhoto wsLibPhotos, sConnectorID

    sCachePath = modLibrary.CachePhotoPath(sWorkbookPath, sConnectorID, "jpg")
    If Len(Dir$(sCachePath)) > 0 Then Kill sCachePath

    vRemoved = modConnectors.RemoveInstancesOfConnectorType(sConnectorID)
    If IsEmpty(vRemoved) Then
        DeleteFromLibrary = modContract.Success("CONNECTOR_DELETED", sConnectorID)
    Else
        DeleteFromLibrary = modContract.Success("CONNECTOR_DELETED_CASCADED", vRemoved)
    End If
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

' One connector from a shared export or library file, copied into this
' library under its own ID - the adapter has already resolved any Part
' Number conflict (Keep/Overwrite) before ever calling this, so an existing
' row of the same ID here means overwrite. If overwriting changes the pin
' count, any chart instances already placed against this ID would hold
' stale pin references, so they are removed the same way a library
' deletion removes them (modConnectors.RemoveInstancesOfConnectorType).
Public Function ImportOneConnector(srcWb As Workbook, wsLibConn As Worksheet, _
                                   wsLibPins As Worksheet, wsLibPhotos As Worksheet, _
                                   ByVal sConnectorID As String, ByVal sOriginFileName As String) As Variant
    Dim vExisting As Variant, bExisted As Boolean, nOldPinCount As Long, nNewPinCount As Long
    Dim vRemoved As Variant

    vExisting = modLibrary.ReadConnector(wsLibConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    bExisted = Not IsEmpty(vExisting)
    If bExisted Then nOldPinCount = CLng(vExisting(modLibrary.LIB_COL_PINCOUNT))

    modLibraryTransfer.ImportConnector srcWb.Worksheets("Connectors"), srcWb.Worksheets("Pins"), _
        srcWb.Worksheets("Photos"), wsLibConn, wsLibPins, wsLibPhotos, sConnectorID, sOriginFileName

    If bExisted Then
        vExisting = modLibrary.ReadConnector(wsLibConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
        nNewPinCount = CLng(vExisting(modLibrary.LIB_COL_PINCOUNT))
        If nNewPinCount <> nOldPinCount Then
            vRemoved = modConnectors.RemoveInstancesOfConnectorType(sConnectorID)
            If Not IsEmpty(vRemoved) Then
                ImportOneConnector = modContract.Success("CONNECTOR_IMPORTED_CASCADED", vRemoved)
                Exit Function
            End If
        End If
    End If

    ImportOneConnector = modContract.Success("CONNECTOR_IMPORTED", sConnectorID)
End Function

' Whether the just-imported connector's photo actually carried over. The
' photo copy goes through the clipboard and is not reliable, so rather than
' trusting the copy call's own return value, this checks whether the shape
' it should have produced exists. The adapter prompts for a replacement
' where it did not.
Public Function ImportedPhotoOk(wsLibPhotos As Worksheet, ByVal sConnectorID As String) As Boolean
    On Error Resume Next
    ImportedPhotoOk = Not (wsLibPhotos.Shapes("PHOTO_" & sConnectorID) Is Nothing)
    On Error GoTo 0
End Function

' Every connector in the library, exported into one shared workbook -
' mirrors ExportToWorkbook per connector rather than duplicating its logic.
Public Function ExportLibraryToWorkbook(wsLibConn As Worksheet, wsLibPins As Worksheet, _
                                        wsLibPhotos As Worksheet, destWb As Workbook) As Variant
    modLibraryTransfer.BuildExportSheets destWb

    Dim vIndex As Variant, i As Long, n As Long, nExported As Long
    vIndex = modLibrary.ConnectorIndex(wsLibConn)
    n = modContract.TableRowCount(vIndex)

    For i = 1 To n
        If modLibraryTransfer.ExportConnector(wsLibConn, wsLibPins, wsLibPhotos, _
                destWb.Worksheets("Connectors"), destWb.Worksheets("Pins"), _
                destWb.Worksheets("Photos"), CStr(vIndex(i, 2))) Then
            nExported = nExported + 1
        End If
    Next i

    ExportLibraryToWorkbook = modContract.Success("LIBRARY_EXPORTED", nExported)
End Function

Public Function AttachReplacementPhoto(wsLibPhotos As Worksheet, ByVal sDestID As String, _
                                       ByVal sPath As String) As Variant
    If Len(modLibrary.EmbedConnectorPhoto(wsLibPhotos, sDestID, sPath)) > 0 Then
        AttachReplacementPhoto = modContract.Success("PHOTO_ATTACHED", sDestID)
    Else
        AttachReplacementPhoto = modContract.Failure("PHOTO_FAILED", sDestID)
    End If
End Function
