Attribute VB_Name = "modConnectorActions"
Option Explicit

' The user-intent transaction behind Remove Connector: turns the layer 0
' boolean into the envelope frmRemoveConnector can render a message from.
Public Function RemoveInstance(ByVal sRefDes As String) As Variant
    If modConnectors.RemoveConnectorInstance(sRefDes) Then
        RemoveInstance = modContract.Success("INSTANCE_REMOVED", sRefDes)
    Else
        RemoveInstance = modContract.Failure("INSTANCE_NOT_FOUND", sRefDes)
    End If
End Function
