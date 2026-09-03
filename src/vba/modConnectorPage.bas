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
Private Const MSO_ALIGN_CENTER As Long = 2
Private Const MSO_ANCHOR_MIDDLE As Long = 3
Private Const MSO_TRUE As Long = -1
Private Const TABLE_HEADERS As String = "Pin,Label,Wire To,Signal,Color,AWG,Termination,Length"

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
    If IsEmpty(vFit) Then
        shpProbe.Delete
        Exit Function
    End If

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
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        shp.TextFrame2.TextRange.Font.Bold = MSO_TRUE
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = MSO_ALIGN_CENTER
        shp.TextFrame2.VerticalAnchor = MSO_ANCHOR_MIDDLE
        shp.TextFrame2.MarginLeft = 0
        shp.TextFrame2.MarginRight = 0
        shp.TextFrame2.MarginTop = 0
        shp.TextFrame2.MarginBottom = 0
        shp.TextFrame2.WordWrap = False

        n = n + 1
    Next i

    PlaceCallouts = n
End Function

Public Sub PlaceLeaderLines(wsPage As Worksheet, shpPhoto As Shape, vPins As Variant)
    Dim i As Long, nPinNumber As Long
    Dim dAnchorX As Double, dAnchorY As Double, dLabelX As Double, dLabelY As Double
    Dim vAnchorPt As Variant, vMarkerPt As Variant
    Dim dMarkerCx As Double, dMarkerCy As Double, dAnchorCx As Double, dAnchorCy As Double
    Dim dDx As Double, dDy As Double, dDist As Double, dStartX As Double, dStartY As Double
    Dim ln As Shape

    If IsEmpty(vPins) Then Exit Sub

    For i = LBound(vPins, 1) To UBound(vPins, 1)
        nPinNumber = CLng(vPins(i, PinCol(vPins, modLibrary.PIN_COL_PINNUM)))
        dAnchorX = CDbl(vPins(i, PinCol(vPins, modLibrary.PIN_COL_NORMX)))
        dAnchorY = CDbl(vPins(i, PinCol(vPins, modLibrary.PIN_COL_NORMY)))
        dLabelX = CDbl(vPins(i, PinCol(vPins, modLibrary.PIN_COL_LABELX)))
        dLabelY = CDbl(vPins(i, PinCol(vPins, modLibrary.PIN_COL_LABELY)))

        If modPinEditor.MarkerSitsOnAnchor(dAnchorX, dAnchorY, dLabelX, dLabelY) Then GoTo NextPin

        vAnchorPt = modPinEditor.MarkerTopLeft(dAnchorX, dAnchorY, _
            shpPhoto.Left, shpPhoto.Top, shpPhoto.Width, shpPhoto.Height, 0, 0)
        vMarkerPt = modPinEditor.MarkerTopLeft(dLabelX, dLabelY, _
            shpPhoto.Left, shpPhoto.Top, shpPhoto.Width, shpPhoto.Height, 0, 0)

        dAnchorCx = vAnchorPt(0): dAnchorCy = vAnchorPt(1)
        dMarkerCx = vMarkerPt(0): dMarkerCy = vMarkerPt(1)

        dDx = dAnchorCx - dMarkerCx
        dDy = dAnchorCy - dMarkerCy
        dDist = Sqr(dDx * dDx + dDy * dDy)
        If dDist > 0 Then
            dStartX = dMarkerCx + (dDx / dDist) * (CONN_OVAL_DIAMETER / 2)
            dStartY = dMarkerCy + (dDy / dDist) * (CONN_OVAL_DIAMETER / 2)
        Else
            dStartX = dMarkerCx
            dStartY = dMarkerCy
        End If

        Set ln = wsPage.Shapes.AddLine(dStartX, dStartY, dAnchorCx, dAnchorCy)
        ln.Name = "LEADER_" & CStr(nPinNumber)
        ln.Line.Weight = 0.75

NextPin:
    Next i
End Sub

Public Sub WriteTableSkeleton(wsPage As Worksheet, vPins As Variant)
    Dim vHeaders As Variant, i As Long, r As Long
    Dim cel As Range

    vHeaders = Split(TABLE_HEADERS, ",")
    For i = LBound(vHeaders) To UBound(vHeaders)
        Set cel = wsPage.Cells(CONN_TABLE_HEADER_ROW, CONN_TABLE_FIRST_COL + i)
        cel.Value = vHeaders(i)
        cel.Font.Bold = True
        cel.Interior.Color = &HD9D9D9
    Next i

    If IsEmpty(vPins) Then Exit Sub

    For i = LBound(vPins, 1) To UBound(vPins, 1)
        r = CONN_TABLE_FIRST_ROW + (i - LBound(vPins, 1))
        wsPage.Cells(r, CONN_TABLE_FIRST_COL).Value = CLng(vPins(i, PinCol(vPins, modLibrary.PIN_COL_PINNUM)))
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 1).Value = CStr(vPins(i, PinCol(vPins, modLibrary.PIN_COL_LABEL)))
    Next i
End Sub

Public Sub WriteMetadata(wsPage As Worksheet, ByVal sConnectorID As String)
    wsPage.Cells(1, CONN_META_COL).Value = sConnectorID
    wsPage.Columns(CONN_META_COL).Hidden = True
End Sub

' Row 1, columns A:I sit above/beside the photo (Top:=CONN_PHOTO_TOP is well
' below row 1's height) and left of the pin table (CONN_TABLE_FIRST_COL is
' column J) - free space for a compact header so the harness/connector this
' page belongs to is visible on every connector page, not just the Harness
' sheet's own title block.
Public Sub WritePageTitleBlock(wsPage As Worksheet, ByVal sHarnessNumber As String, _
                                ByVal sRevision As String, ByVal sRefDes As String, _
                                ByVal sConnectorID As String)
    Dim rng As Range
    Set rng = wsPage.Range("A1:I1")
    rng.Merge
    rng.Value = "Harness " & sHarnessNumber & "  Rev " & sRevision & _
                "   -   " & sRefDes & " (" & sConnectorID & ")"
    rng.Font.Bold = True
    rng.Interior.Color = &HF2F2F2
End Sub

Private Function KeyExpr(ByVal sRefDes As String, ByVal nTableRow As Long) As String
    Dim sEscaped As String
    sEscaped = Replace(sRefDes, """", """""")
    KeyExpr = """" & sEscaped & "|""&$J" & CStr(nTableRow)
End Function

' A matched-but-blank chart cell makes INDEX return an empty string, not an
' error - so the outer IFERROR alone can't turn it into a blank pin-table
' cell. Each branch below repeats its INDEX(...) sub-expression once for an
' IF("" ...) blank test and once for the return value (Excel 2016 has no
' LET() to name it once) so a successful-but-blank match still renders "".
Public Function LookupFormula(ByVal sFromCol As String, ByVal sToCol As String, _
                              ByVal sRefDes As String, ByVal nTableRow As Long) As String
    Dim sKey As String, Q As String
    Dim sFromIndex As String, sToIndex As String
    sKey = KeyExpr(sRefDes, nTableRow)
    Q = """"

    sFromIndex = "INDEX(Harness!$" & sFromCol & "$7:$" & sFromCol & "$1006," & _
        "MATCH(" & sKey & ",Harness!$L$7:$L$1006,0))"
    sToIndex = "INDEX(Harness!$" & sToCol & "$7:$" & sToCol & "$1006," & _
        "MATCH(" & sKey & ",Harness!$M$7:$M$1006,0))"

    LookupFormula = "=IFERROR(IF(" & sFromIndex & "=" & Q & Q & "," & Q & Q & "," & sFromIndex & ")," & _
        "IFERROR(IF(" & sToIndex & "=" & Q & Q & "," & Q & Q & "," & sToIndex & ")," & Q & Q & "))"
End Function

Public Function WireToFormula(ByVal sRefDes As String, ByVal nTableRow As Long) As String
    Dim sKey As String, Q As String
    Dim sFromConn As String, sFromPin As String, sToConn As String, sToPin As String
    sKey = KeyExpr(sRefDes, nTableRow)
    Q = """"

    sFromConn = "INDEX(Harness!$I$7:$I$1006,MATCH(" & sKey & ",Harness!$L$7:$L$1006,0))"
    sFromPin = "INDEX(Harness!$J$7:$J$1006,MATCH(" & sKey & ",Harness!$L$7:$L$1006,0))"
    sToConn = "INDEX(Harness!$A$7:$A$1006,MATCH(" & sKey & ",Harness!$M$7:$M$1006,0))"
    sToPin = "INDEX(Harness!$B$7:$B$1006,MATCH(" & sKey & ",Harness!$M$7:$M$1006,0))"

    WireToFormula = "=IFERROR(IF(OR(" & sFromConn & "=" & Q & Q & "," & sFromPin & "=" & Q & Q & ")," & Q & Q & "," & _
        sFromConn & "&" & Q & "-" & Q & "&" & sFromPin & ")," & _
        "IFERROR(IF(OR(" & sToConn & "=" & Q & Q & "," & sToPin & "=" & Q & Q & ")," & Q & Q & "," & _
        sToConn & "&" & Q & "-" & Q & "&" & sToPin & ")," & Q & Q & "))"
End Function

Public Sub WriteLiveFormulas(wsPage As Worksheet, ByVal sRefDes As String, vPins As Variant)
    Dim i As Long, r As Long

    If IsEmpty(vPins) Then Exit Sub

    For i = LBound(vPins, 1) To UBound(vPins, 1)
        r = CONN_TABLE_FIRST_ROW + (i - LBound(vPins, 1))

        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 2).Formula = WireToFormula(sRefDes, r)
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 3).Formula = LookupFormula("D", "D", sRefDes, r)
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 4).Formula = LookupFormula("E", "E", sRefDes, r)
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 5).Formula = LookupFormula("F", "F", sRefDes, r)
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 6).Formula = LookupFormula("C", "H", sRefDes, r)
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 7).Formula = LookupFormula("G", "G", sRefDes, r)
    Next i
End Sub
