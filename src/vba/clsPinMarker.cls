VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPinMarker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PinNumber As Long
Public ConnectorID As String
Public ScratchSheet As Worksheet
Public PhotoControl As MSForms.Image

Private WithEvents mLabel As MSForms.Label
Private mDragging As Boolean
Private mGrabDx As Single, mGrabDy As Single

Public Property Get LabelControl() As MSForms.Label
    Set LabelControl = mLabel
End Property

Public Property Set LabelControl(ByVal ctl As MSForms.Label)
    Set mLabel = ctl
End Property

Private Sub mLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        mDragging = True
        mGrabDx = X
        mGrabDy = Y
    End If
End Sub

Private Sub mLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    If mDragging Then
        mLabel.Left = mLabel.Left + (X - mGrabDx)
        mLabel.Top = mLabel.Top + (Y - mGrabDy)
    End If
End Sub

Private Sub mLabel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                           ByVal X As Single, ByVal Y As Single)
    If Not mDragging Then Exit Sub
    mDragging = False

    Dim vNorm As Variant
    vNorm = modPinEditor.NormFromMarker( _
        CDbl(mLabel.Left), CDbl(mLabel.Top), CDbl(mLabel.Width), CDbl(mLabel.Height), _
        CDbl(PhotoControl.Left), CDbl(PhotoControl.Top), _
        CDbl(PhotoControl.Width), CDbl(PhotoControl.Height))
    If IsEmpty(vNorm) Then Exit Sub

    modPinEditor.MoveMarker ScratchSheet, ConnectorID, PinNumber, vNorm(0), vNorm(1)
End Sub
