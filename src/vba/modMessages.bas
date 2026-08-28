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
        Case "EXPORTED"
            MessageFor = "Exported " & CStr(vPayload) & "."
        Case "EXPORT_FAILED"
            MessageFor = "Could not export " & CStr(vPayload) & "."
        Case "IMPORTED"
            MessageFor = "Import complete. " & _
                CStr(modContract.TableRowCount(vPayload)) & " connector(s) imported."
        Case "PHOTO_FAILED"
            MessageFor = "Could not attach a photo for " & CStr(vPayload) & "."
        Case Else
            MessageFor = ""
    End Select
End Function

Public Function MessageStyleFor(vResult As Variant) As Long
    If modContract.Ok(vResult) Then
        MessageStyleFor = vbInformation
    Else
        MessageStyleFor = vbExclamation
    End If
End Function
