Attribute VB_Name = "modMessages"
Option Explicit

' Every string a student reads lives here rather than in an .evt file, so
' the exact rendered text is asserted by a test. An outcome with no case
' below is silent by design - the adapter shows nothing for it.
Public Function MessageFor(vResult As Variant) As String
    Dim sOutcome As String, vPayload As Variant
    sOutcome = modContract.Outcome(vResult)
    vPayload = modContract.Payload(vResult)

    Select Case sOutcome
        Case "ID_COLLISION"
            MessageFor = "Part Number already exists in the library (" & _
                CStr(vPayload) & "). Choose a different Part Number."
        Case "SAVE_FAILED"
            MessageFor = "Could not save " & CStr(vPayload) & ". Load a photo before saving."
        Case "BAD_PIN_COUNT"
            MessageFor = "Enter a valid Pin Count before placing pins."
        Case "PIN_LIMIT_REACHED"
            MessageFor = "All " & CStr(vPayload) & " pins have been placed."
        Case "MISSING_NAME_OR_PART"
            MessageFor = "Enter Name and Part Number before loading a photo."
        Case "CONNECTOR_NOT_FOUND"
            MessageFor = "No connector '" & CStr(vPayload) & "' found in the library."
        Case "ADD_FAILED"
            MessageFor = "Could not add an instance of " & CStr(vPayload) & "."
        Case "CONNECTOR_DELETED"
            MessageFor = "Deleted " & CStr(vPayload) & "."
        Case "CONNECTOR_DELETED_CASCADED"
            MessageFor = "Deleted from the library. " & RemovedInstancesClause(vPayload)
        Case "EXPORTED"
            MessageFor = "Exported " & CStr(vPayload) & "."
        Case "EXPORT_FAILED"
            MessageFor = "Could not export " & CStr(vPayload) & "."
        Case "LIBRARY_EXPORTED"
            MessageFor = "Exported " & CStr(vPayload) & " connector(s)."
        Case "PHOTO_FAILED"
            MessageFor = "Could not attach a photo for " & CStr(vPayload) & "."
        Case "INSTANCE_NOT_FOUND"
            MessageFor = "No connector instance '" & CStr(vPayload) & "' found."
        Case "HARNESS_SAVED"
            MessageFor = "Saved. " & CStr(vPayload) & " wire(s) written."
        Case "HARNESS_SAVE_FAILED"
            MessageFor = "Could not save the harness: " & CStr(vPayload) & "."
        Case "HARNESS_LOADED"
            MessageFor = "Loaded. " & CStr(vPayload) & " wire(s) read."
        Case "HARNESS_LOAD_FAILED"
            MessageFor = "Could not load the harness: " & CStr(vPayload) & "."
        Case Else
            MessageFor = ""
    End Select
End Function

' "Removed 2 connector instance(s) from the chart: J1, J2." - shared by
' every outcome whose payload is a flat list of ref des a cascade removed.
Private Function RemovedInstancesClause(vPayload As Variant) As String
    Dim i As Long, sList As String

    For i = LBound(vPayload) To UBound(vPayload)
        If Len(sList) > 0 Then sList = sList & ", "
        sList = sList & CStr(vPayload(i))
    Next i

    RemovedInstancesClause = "Removed " & CStr(modContract.TableRowCount(vPayload)) & _
        " connector instance(s) from the chart: " & sList & "."
End Function

' The one message an import loop shows, after every per-connector
' Keep/Overwrite prompt has already been answered - not an outcome-driven
' message, since the counts are tallied by the adapter across many calls
' rather than returned by any single action.
Public Function ImportSummaryMessage(ByVal nImported As Long, ByVal nKept As Long, _
                                     ByVal nOverwritten As Long, vRemovedRefDes As Variant) As String
    Dim sMsg As String
    sMsg = "Import complete. Imported: " & CStr(nImported) & _
        ". Kept: " & CStr(nKept) & ". Overwritten: " & CStr(nOverwritten) & "."

    If Not IsEmpty(vRemovedRefDes) Then
        sMsg = sMsg & " " & RemovedInstancesClause(vRemovedRefDes)
    End If

    ImportSummaryMessage = sMsg
End Function

Public Function MessageStyleFor(vResult As Variant) As Long
    If modContract.Ok(vResult) Then
        MessageStyleFor = vbInformation
    Else
        MessageStyleFor = vbExclamation
    End If
End Function
