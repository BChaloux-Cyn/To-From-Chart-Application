Attribute VB_Name = "modSnapshot"
Option Explicit

Public Const SNAP_CONN_FIRST_ROW As Long = 2
Public Const SNAP_CONN_LAST_ROW As Long = 201
Public Const SNAP_PINS_FIRST_ROW As Long = 211
Public Const SNAP_PINS_LAST_ROW As Long = 2210

Public Function LibraryFolder() As String
    LibraryFolder = ThisWorkbook.Path
End Function

Public Function SnapshotConnector(wsSnap As Worksheet, wsLibConn As Worksheet, wsLibPins As Worksheet, _
                                  wsLibPhotos As Worksheet, ByVal sConnectorID As String) As Boolean
    ' Frozen once per distinct ConnectorID - a second instance of the same
    ' library part (J1 and J2 from one DTM-04P) never duplicates the
    ' definition it shares.
    If modLibrary.FindConnectorRow(wsSnap, SNAP_CONN_FIRST_ROW, SNAP_CONN_LAST_ROW, sConnectorID) > 0 Then
        SnapshotConnector = True
        Exit Function
    End If

    Dim vFields As Variant
    vFields = modLibrary.ReadConnector(wsLibConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vFields) Then Exit Function
    If Not modLibrary.WriteConnector(wsSnap, SNAP_CONN_FIRST_ROW, SNAP_CONN_LAST_ROW, vFields) Then Exit Function

    Dim vPins As Variant, i As Long, j As Long, vRow(1 To 7) As Variant
    vPins = modLibrary.ReadPinsForConnector(wsLibPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If Not IsEmpty(vPins) Then
        For i = LBound(vPins, 1) To UBound(vPins, 1)
            For j = 1 To modLibrary.PIN_FIELD_COUNT
                vRow(j) = vPins(i, j)
            Next j
            modLibrary.WritePin wsSnap, SNAP_PINS_FIRST_ROW, SNAP_PINS_LAST_ROW, vRow
        Next i
    End If

    Dim sCachePath As String
    sCachePath = modLibrary.CachePhotoPath(LibraryFolder(), sConnectorID)
    If Len(Dir$(sCachePath)) > 0 Then
        modLibrary.EmbedConnectorPhoto wsSnap, sConnectorID, sCachePath
    Else
        ' No local cache file yet (e.g. the sample fixture in this plan's
        ' own tests never wrote one) - copy the shape straight off the
        ' library's own Photos sheet instead of failing the snapshot.
        modLibrary.RemoveConnectorPhoto wsSnap, sConnectorID
        On Error Resume Next
        wsLibPhotos.Shapes("PHOTO_" & sConnectorID).Copy
        wsSnap.Paste
        wsSnap.Shapes(wsSnap.Shapes.Count).Name = "PHOTO_" & sConnectorID
        On Error GoTo 0
    End If

    SnapshotConnector = True
End Function
