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
Public Const CONN_OVAL_DIAMETER As Double = 14

Private Const MSO_SHAPE_OVAL As Long = 9

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

' modLibrary.PIN_COL_* are 1-based offsets into a pin row. A vPins array
' built inside VBA (modLibrary.ReadPinsForConnector) is itself 1-based, but
' one arriving as a COM argument from a Python test's tuple-of-tuples is
' 0-based - indexing by the raw 1-based constant would silently read the
' wrong column there. Anchoring to the array's own LBound makes every
' PIN_COL_* lookup correct regardless of which side constructed the array.
Private Function PinCol(vPins As Variant, ByVal nCol1Based As Long) As Long
    PinCol = LBound(vPins, 2) + (nCol1Based - 1)
End Function

Public Function PlaceCallouts(wsPage As Worksheet, shpPhoto As Shape, vPins As Variant) As Long
    Dim i As Long, nPinNumber As Long, dLabelX As Double, dLabelY As Double
    Dim vTopLeft As Variant, shp As Shape, n As Long

    If IsEmpty(vPins) Then Exit Function

    For i = LBound(vPins, 1) To UBound(vPins, 1)
        nPinNumber = CLng(vPins(i, PinCol(vPins, modLibrary.PIN_COL_PINNUM)))
        dLabelX = CDbl(vPins(i, PinCol(vPins, modLibrary.PIN_COL_LABELX)))
        dLabelY = CDbl(vPins(i, PinCol(vPins, modLibrary.PIN_COL_LABELY)))

        vTopLeft = modPinEditor.MarkerTopLeft(dLabelX, dLabelY, _
            shpPhoto.Left, shpPhoto.Top, shpPhoto.Width, shpPhoto.Height, _
            CONN_OVAL_DIAMETER, CONN_OVAL_DIAMETER)

        Set shp = wsPage.Shapes.AddShape(MSO_SHAPE_OVAL, vTopLeft(0), vTopLeft(1), _
            CONN_OVAL_DIAMETER, CONN_OVAL_DIAMETER)
        shp.Name = "PIN_" & CStr(nPinNumber)
        shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.Line.ForeColor.RGB = RGB(0, 0, 0)
        shp.TextFrame2.TextRange.Text = CStr(nPinNumber)
        shp.TextFrame2.TextRange.Font.Size = 8
        shp.TextFrame2.WordWrap = False

        n = n + 1
    Next i

    PlaceCallouts = n
End Function
