Attribute VB_Name = "modPickerActions"
Option Explicit

' Adding a library connector to this harness: allocate a ref des, write the
' instance row, and freeze the definition into _Snapshot. One transaction,
' called from both the picker's Add button and its New-then-save chain,
' which previously duplicated it.
Public Function AddFromLibrary(wsSnapshot As Worksheet, wsLibConn As Worksheet, _
                               wsLibPins As Worksheet, wsLibPhotos As Worksheet, _
                               ByVal sConnectorID As String) As Variant
    Dim vFields As Variant, sRefDes As String

    vFields = modLibrary.ReadConnector(wsLibConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vFields) Then
        AddFromLibrary = modContract.Failure("CONNECTOR_NOT_FOUND", sConnectorID)
        Exit Function
    End If

    sRefDes = modConnectors.AddConnectorInstance(CStr(vFields(1)), CStr(vFields(2)), _
        CStr(vFields(4)), CStr(vFields(5)), CLng(vFields(6)))
    If Len(sRefDes) = 0 Then
        AddFromLibrary = modContract.Failure("ADD_FAILED", sConnectorID)
        Exit Function
    End If

    modSnapshot.SnapshotConnector wsSnapshot, wsLibConn, wsLibPins, wsLibPhotos, sConnectorID

    AddFromLibrary = modContract.Success("ADDED", sRefDes)
End Function
