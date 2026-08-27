Attribute VB_Name = "modConnectorUI"
Option Explicit

Public Sub ShowAddConnector()
    frmConnectorPicker.Show
End Sub

Public Sub ShowManageLibrary()
    frmManageLibrary.Show
End Sub

Public Sub ShowRemoveConnector()
    Dim sRefDes As String
    sRefDes = InputBox("Ref des to remove:", "Remove Connector")
    If Len(Trim$(sRefDes)) = 0 Then Exit Sub
    If Not modConnectors.RemoveConnectorInstance(Trim$(sRefDes)) Then
        MsgBox "No connector instance '" & sRefDes & "' found.", vbExclamation
    End If
End Sub

' Shared by frmConnectorPicker and frmManageLibrary, whose listboxes are
' both a plain "ConnectorID - Name" rendering of the library's Connectors
' sheet. Returns the row order's ConnectorIDs, 1-indexed to match
' lst.ListIndex + 1, so callers can resolve a click back to a ConnectorID.
Public Function RefreshConnectorList(wsConn As Worksheet, lst As MSForms.ListBox) As String()
    Dim nLast As Long, r As Long, n As Long
    Dim vIDs() As String

    lst.Clear
    nLast = wsConn.Cells(wsConn.Rows.Count, 1).End(xlUp).Row
    If nLast < 2 Then Exit Function

    ReDim vIDs(1 To nLast - 1)
    For r = 2 To nLast
        n = n + 1
        vIDs(n) = Trim$(CStr(wsConn.Cells(r, modLibrary.LIB_COL_ID).Value))
        lst.AddItem vIDs(n) & " - " & CStr(wsConn.Cells(r, modLibrary.LIB_COL_NAME).Value)
    Next r

    RefreshConnectorList = vIDs
End Function
