Attribute VB_Name = "modContract"
Option Explicit

Public Const KIND_NONE As String = "NONE"
Public Const KIND_STRING As String = "STRING"
Public Const KIND_LONG As String = "LONG"
Public Const KIND_DOUBLE As String = "DOUBLE"
Public Const KIND_TABLE As String = "TABLE"

' Every outcome code any layer 1 action may return. Success/Failure reject
' anything absent from this list, so a typo fails at construction rather
' than reaching an adapter that has no case for it.
Public Function OutcomeCodes() As Variant
    OutcomeCodes = Array( _
        "PLACED", "MOVED_ANCHOR", "BAD_PIN_COUNT", "PIN_LIMIT_REACHED", "NO_OP", _
        "SAVED", "ID_COLLISION", "SAVE_FAILED", _
        "CACHE_READY", "NEEDS_BACKFILL", _
        "OK", "MISSING_NAME_OR_PART", _
        "PIN_DELETED", "PIN_NOT_FOUND", _
        "ADDED", "ADD_FAILED", "CONNECTOR_NOT_FOUND", "CONNECTOR_DELETED", _
        "EXPORTED", "EXPORT_FAILED", "IMPORTED", _
        "PHOTO_ATTACHED", "PHOTO_FAILED", _
        "RENAMED", "RENAME_REJECTED", "NO_RENAME", _
        "BULK_REBUILT", "CELLS_REBUILT", "UNITS_SET")
End Function

' One code maps to exactly one payload kind. PIN_DELETED and
' CONNECTOR_DELETED are deliberately distinct codes rather than one
' DELETED: their payloads are a pin number and a connector ID.
Public Function PayloadKind(ByVal sOutcome As String) As String
    Select Case sOutcome
        Case "PLACED", "MOVED_ANCHOR", "PIN_LIMIT_REACHED", "PIN_DELETED", "PIN_NOT_FOUND"
            PayloadKind = KIND_LONG
        Case "BULK_REBUILT", "CELLS_REBUILT"
            PayloadKind = KIND_LONG
        Case "SAVED", "ID_COLLISION", "SAVE_FAILED", "CACHE_READY", "NEEDS_BACKFILL"
            PayloadKind = KIND_STRING
        Case "ADDED", "ADD_FAILED", "CONNECTOR_NOT_FOUND", "CONNECTOR_DELETED"
            PayloadKind = KIND_STRING
        Case "EXPORTED", "EXPORT_FAILED", "PHOTO_ATTACHED", "PHOTO_FAILED"
            PayloadKind = KIND_STRING
        Case "RENAMED", "RENAME_REJECTED", "UNITS_SET"
            PayloadKind = KIND_STRING
        Case "IMPORTED"
            PayloadKind = KIND_TABLE
        Case "BAD_PIN_COUNT", "NO_OP", "OK", "MISSING_NAME_OR_PART", "NO_RENAME"
            PayloadKind = KIND_NONE
        Case Else
            PayloadKind = ""
    End Select
End Function

Public Function Success(ByVal sOutcome As String, Optional ByVal vPayload As Variant) As Variant
    If IsMissing(vPayload) Then
        Success = Build(True, sOutcome, Empty)
    Else
        Success = Build(True, sOutcome, vPayload)
    End If
End Function

Public Function Failure(ByVal sOutcome As String, Optional ByVal vPayload As Variant) As Variant
    If IsMissing(vPayload) Then
        Failure = Build(False, sOutcome, Empty)
    Else
        Failure = Build(False, sOutcome, vPayload)
    End If
End Function

' Array() is zero based for an in-process VBA caller and for a COM caller
' alike, so vResult(0) here and result[0] in pytest are the same element.
Private Function Build(ByVal bOk As Boolean, ByVal sOutcome As String, _
                       ByVal vPayload As Variant) As Variant
    Dim sKind As String
    sKind = PayloadKind(sOutcome)
    If Len(sKind) = 0 Then
        Err.Raise vbObjectError + 1, "modContract", "Unknown outcome code: " & sOutcome
    End If
    If Not PayloadMatches(sKind, vPayload) Then
        Err.Raise vbObjectError + 2, "modContract", _
            "Payload for " & sOutcome & " must be " & sKind
    End If
    Build = Array(bOk, sOutcome, vPayload)
End Function

Private Function PayloadMatches(ByVal sKind As String, ByVal vPayload As Variant) As Boolean
    Select Case sKind
        Case KIND_NONE:   PayloadMatches = IsEmpty(vPayload)
        Case KIND_STRING: PayloadMatches = (VarType(vPayload) = vbString)
        Case KIND_LONG:   PayloadMatches = (Not IsArray(vPayload)) And IsNumeric(vPayload)
        Case KIND_DOUBLE: PayloadMatches = (Not IsArray(vPayload)) And IsNumeric(vPayload)
        Case KIND_TABLE:  PayloadMatches = IsEmpty(vPayload) Or IsArray(vPayload)
    End Select
End Function

Public Function Ok(vResult As Variant) As Boolean
    Ok = CBool(vResult(0))
End Function

Public Function Outcome(vResult As Variant) As String
    Outcome = CStr(vResult(1))
End Function

Public Function Payload(vResult As Variant) As Variant
    If IsObject(vResult(2)) Then
        Set Payload = vResult(2)
    Else
        Payload = vResult(2)
    End If
End Function

' Adapters call this rather than LBound/UBound, so a zero row table cannot
' produce a subscript error in a form.
Public Function TableRowCount(vPayload As Variant) As Long
    If IsEmpty(vPayload) Then Exit Function
    If Not IsArray(vPayload) Then Exit Function
    TableRowCount = UBound(vPayload, 1) - LBound(vPayload, 1) + 1
End Function
