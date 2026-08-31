Attribute VB_Name = "modConnectorUI"
Option Explicit

' Set by frmConnectorEditor.cmdSave_Click on a successful save, "" otherwise
' (Cancel, or no save yet). Lets frmConnectorPicker's cmdNew_Click know
' whether - and which connector - to chain into adding an instance of,
' without frmConnectorEditor needing to know anything about that caller.
' A standard module's variable, not a property on the form itself: the form
' unloads itself right after setting this, and re-reading a property off a
' just-unloaded predeclared form instance risks re-triggering its
' Initialize (which would reset it) instead of returning the value just set.
Public LastSavedConnectorID As String

Public Sub ShowAddConnector()
    frmConnectorPicker.Show
End Sub

Public Sub ShowManageLibrary()
    frmManageLibrary.Show
End Sub

Public Sub ShowRemoveConnector()
    frmRemoveConnector.Show
End Sub
