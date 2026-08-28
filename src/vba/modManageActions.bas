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

' Every connector in a shared export file, copied into this library.
' modLibraryTransfer.ImportConnector renames on an ID collision and
' attempts the photo copy itself; the photo copy goes through the
' clipboard and is not reliable, so rather than redoing it here (which
' could give a different answer than the one that actually happened) this
' checks whether the shape it should have produced exists, and reports
' per connector. The adapter prompts for a replacement where it did not.
Public Function ImportAllFromWorkbook(srcWb As Workbook, wsLibConn As Worksheet, _
                                      wsLibPins As Worksheet, wsLibPhotos As Worksheet) As Variant
    Dim wsSrcConn As Worksheet, nLast As Long, r As Long, n As Long
    Dim sConnectorID As String, sDestID As String, sOriginName As String
    Dim vRows() As Variant, bPhotoOk As Boolean

    Set wsSrcConn = srcWb.Worksheets("Connectors")
    sOriginName = srcWb.Name
    nLast = wsSrcConn.Cells(wsSrcConn.Rows.Count, modLibrary.LIB_COL_ID).End(xlUp).Row
    If nLast < 2 Then
        ImportAllFromWorkbook = modContract.Success("IMPORTED", Empty)
        Exit Function
    End If

    ReDim vRows(1 To nLast - 1, 1 To 2)
    For r = 2 To nLast
        sConnectorID = Trim$(CStr(wsSrcConn.Cells(r, modLibrary.LIB_COL_ID).Value))
        If Len(sConnectorID) > 0 Then
            sDestID = modLibraryTransfer.ImportConnector(wsSrcConn, _
                srcWb.Worksheets("Pins"), srcWb.Worksheets("Photos"), _
                wsLibConn, wsLibPins, wsLibPhotos, sConnectorID, sOriginName)

            If Len(sDestID) > 0 Then
                bPhotoOk = False
                On Error Resume Next
                bPhotoOk = Not (wsLibPhotos.Shapes("PHOTO_" & sDestID) Is Nothing)
                On Error GoTo 0

                n = n + 1
                vRows(n, 1) = sDestID
                vRows(n, 2) = bPhotoOk
            End If
        End If
    Next r

    If n = 0 Then
        ImportAllFromWorkbook = modContract.Success("IMPORTED", Empty)
        Exit Function
    End If

    Dim vResult() As Variant, i As Long
    ReDim vResult(1 To n, 1 To 2)
    For i = 1 To n
        vResult(i, 1) = vRows(i, 1)
        vResult(i, 2) = vRows(i, 2)
    Next i
    ImportAllFromWorkbook = modContract.Success("IMPORTED", vResult)
End Function

Public Function AttachReplacementPhoto(wsLibPhotos As Worksheet, ByVal sDestID As String, _
                                       ByVal sPath As String) As Variant
    If Len(modLibrary.EmbedConnectorPhoto(wsLibPhotos, sDestID, sPath)) > 0 Then
        AttachReplacementPhoto = modContract.Success("PHOTO_ATTACHED", sDestID)
    Else
        AttachReplacementPhoto = modContract.Failure("PHOTO_FAILED", sDestID)
    End If
End Function
