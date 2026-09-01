Attribute VB_Name = "modConnectorPage"
Option Explicit

Public Const CONN_PHOTO_LEFT As Double = 20
Public Const CONN_PHOTO_TOP As Double = 60
Public Const CONN_PHOTO_MAX_WIDTH As Double = 300
Public Const CONN_PHOTO_MAX_HEIGHT As Double = 300
Public Const CONN_META_COL As Long = 27
Public Const CONN_TABLE_FIRST_COL As Long = 10
Public Const CONN_TABLE_HEADER_ROW As Long = 1
Public Const CONN_TABLE_FIRST_ROW As Long = 2

Public Function PagePhotoPath(ByVal sLibraryFolder As String, ByVal sConnectorID As String) As String
    Dim sJpg As String, sPng As String
    sJpg = modLibrary.CachePhotoPath(sLibraryFolder, sConnectorID, "jpg")
    If Len(Dir$(sJpg)) > 0 Then
        PagePhotoPath = sJpg
        Exit Function
    End If

    sPng = modLibrary.CachePhotoPath(sLibraryFolder, sConnectorID)
    If Len(Dir$(sPng)) > 0 Then PagePhotoPath = sPng
End Function

Public Function PlacePhoto(wsPage As Worksheet, ByVal sPhotoPath As String) As Boolean
    If Len(sPhotoPath) = 0 Then Exit Function
    If Len(Dir$(sPhotoPath)) = 0 Then Exit Function

    Dim shpProbe As Shape
    Set shpProbe = wsPage.Shapes.AddPicture(sPhotoPath, False, True, 0, 0, -1, -1)
    Dim vFit As Variant
    vFit = modPinEditor.FitAspectRatio(shpProbe.Width, shpProbe.Height, CONN_PHOTO_MAX_WIDTH, CONN_PHOTO_MAX_HEIGHT)

    shpProbe.Left = CONN_PHOTO_LEFT
    shpProbe.Top = CONN_PHOTO_TOP
    shpProbe.LockAspectRatio = False
    shpProbe.Width = vFit(0)
    shpProbe.Height = vFit(1)
    shpProbe.Name = "PAGE_PHOTO"

    PlacePhoto = True
End Function
