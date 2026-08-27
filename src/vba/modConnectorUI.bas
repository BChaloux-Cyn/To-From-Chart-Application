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
