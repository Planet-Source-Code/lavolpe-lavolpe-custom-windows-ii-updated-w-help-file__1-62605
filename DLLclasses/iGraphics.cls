VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' this class exposes some of the drawing functions used by the DLL
' Why should you have to reinvent the wheel if it is already done?


Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal uFlags As Long) As Long
Private Declare Function DrawFrameControl Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Const DSS_DISABLED As Long = &H20
Private Const DSS_MONO As Long = &H80
Private Const DSS_NORMAL As Long = &H0
Private Const DST_ICON As Long = &H3

Private mImplementation As Long ' functionality to get other class's properites


'---------------------------------------------------------------------------------------
' Procedure : cGraphics.ImageFillEx
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : stretch, smart stretch, or tileBlt a stdPicture across a DC
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub ImageFillEx(ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, _
                ByVal Right As Long, ByVal Bottom As Long, SourceImage As StdPicture, _
                ByVal Options As eBackStyles, Optional ByVal Vertical As Boolean, _
                Optional ByVal GrayScaled As Boolean)
        
    ' X,Y are the top/left coordinates to begin the graphics
    ' Width,Height are self-explanatory
    ' SourceImage must be a stdPicture object: picBox, Image, ImageList item, etc)
    ' Options: stretch uses simple StretchBlt
    '        : tile will tile blt the image, no scaling is used
    '        : smartStretch depends on the Vertical flag
    '        :: if Vertical=True. Only stretched vertically, not horizontally
    '        ::       otherwise   Only stretched horizontally, not vertically
    '        :: additionally, if image doesn't need to be stretched it won't
    '           so if image area is smaller than the image, no stretching occurs
    ' Vertical: vertical (top/bottom) smartStretch vs horizontal (left/right)
    
    On Error Resume Next
    
    If SourceImage Is Nothing Then Exit Sub
    If hDC = 0 Then Exit Sub
        
    ' pass the function off
    FillBarImage SourceImage, hDC, Left, Top, Right - Left, Bottom - Top, Options, Vertical, GrayScaled
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.GradientFillEx
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : fill gradient fill a DC either left to right or top to bottom
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub GradientFillEx(ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, _
                ByVal Right As Long, ByVal Bottom As Long, ByVal Vertical As Boolean, _
                ByVal Color1 As Long, ByVal Color2 As Long, Optional GrayScaled As Boolean)
                
    ' X,Y are the top/left coordinates to begin gradient
    ' Width,Height are self-explanatory
    ' Vertical will start gradient from Y and fill to Height,
    '       otherwise start from X and fill to Width
    ' Color1, Color2 are the far colors encompassing the gradient range
    ' bGrayScaled will grayscale Color1 & Color2 before drawing the gradient
    
    On Error Resume Next
    If hDC = 0 Then Exit Sub
    
    ' pass the function off
    If Vertical Then Bottom = -Bottom ' flag to indicate vertical vs horizontal
    GradientFill Color1, Color2, hDC, Left, Top, Right, Bottom, True, , , GrayScaled
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.TextDrawEx
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : draw text in up to 4 styles using up to 3 colors
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub TextDrawEx(ByVal Text As String, ByVal hDC As Long, ByVal Left As Long, _
            ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, _
            ByVal Style As eFX, ByVal DTFlags As eDTflags, _
            ByVal FontHandle As eGFont, gdcFont As StdFont, ParamArray Colors() As Variant)
    
    ' the Left,Top,Right,Bottom form the rectangle needed by DrawText API
    ' style: flat uses only 1st color in passed array
    '      : sunken uses 1st 2 colors in passed array
    '      : raised uses 1st 2 colors in passed array
    '      : engraved uses all 3 colors in passed array
    ' DTflags are standard DT_flag values needed by DrawText API
    ' FontHandle: gDC_Font, do not select a font, use the font in the DC
    '           : gMenu_Font, get copy of the actual menu font
    '           : gTitlebar_Font, get copy of the actual titlebar font
    '           Note. Copies used in case user decides to unselect the font from the DC
    ' Colors: an array of up to 3 colors containing the colors you want used
    '       this array can be byte, integer or long - nothing else
    '       1st Color is the primary text color
    '       2nd Color is the highlight used for raised,sunken, and engraved
    '       3rd Color is only used for the engraved style
    
    On Error Resume Next
    If Text = "" Then Exit Sub
    If hDC = 0 Then Exit Sub
    
    ' null ParamArray?
    If LBound(Colors) > UBound(Colors) Then Exit Sub
    
    Dim arrBase As Long, arrUBound As Long
    Dim C As Long, fColors() As Long
    
    ' start parsing out the colors and add to a Long vs Variant array
    arrBase = LBound(Colors)
    arrUBound = UBound(Colors)
    If arrUBound > arrBase + 2 Then arrUBound = arrBase + 2
    
    ReDim fColors(arrBase To arrUBound)
    For C = arrBase To arrUBound
        Select Case VarType(Colors(C))
        Case vbLong, vbInteger, vbByte ' don't accept any other varTypes
            fColors(C) = CLng(Colors(C))
        Case Else
            ' the color will be black, value of zero
        End Select
    Next
    
    Dim hFont As Long, dFont As Long
    Dim cClient As iImpProps
    
    ' get the requested font as needed
    If FontHandle = gMenu_Font Or FontHandle = gTitlebar_Font Then
        GetObjectFromPointer mImplementation, cClient
        If FontHandle = gMenu_Font Then
            dFont = cClient.GetPropValue([_MenuFont])
        Else
            dFont = cClient.GetPropValue([_TitleFont])
        End If
        Set cClient = Nothing
        ' select the font into the DC
        hFont = SelectObject(hDC, dFont)
    Else
        If Not gdcFont Is Nothing Then
            dFont = FontToHandle(gdcFont, False)
            If dFont Then hFont = SelectObject(hDC, dFont)
        End If
    End If
    
    ' not passed below are optional flags only passed from the internal drawing classes
    StyleText Text, Left, Top, Right, Bottom, Style, fColors(), hDC, DTFlags, True, arrBase + 0
    
    If hFont Then DeleteObject SelectObject(hDC, hFont)
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.DrawImageIcon
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : draws an icon as normal or disabled
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub DrawImageIcon(ByVal hImage As Long, ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, _
                    Optional ByVal Cx = 16, Optional ByVal Cy = 16, Optional asDisabled As Boolean, _
                    Optional ByVal PrimaryColor As Long = vbGrayText, Optional ByVal Highlight As Long = vb3DHighlight)

    ' parameters are self-explanatory I believe with following exception
    ' To draw a disabled icon using only one color, pass Highlight as -1

    On Error Resume Next

    Dim tImage As Long
    Dim brGray As Long, brLight As Long
    
    ' no image to drawn? nothing to do
    If hImage = 0 Then Exit Sub
    If hDC = 0 Then Exit Sub
    ' if drawing disabled icons and no primary brush provided (-1) then exit sub
    If asDisabled = True And PrimaryColor = -1 Then Exit Sub
    
    ' we'll make a copy of the passed image. Why?
    ' Because we don't know its source. Using the ImageList.ListImage(x).Picture.Handle
    ' will not work with DrawState (used for disabled icons here) only if DrawState
    ' will be drawing the image at a different size than the source size.
    ' But using a copy of the image, DrawState has no such issues
    
    tImage = CopyImage(hImage, IMAGE_ICON, Cx, Cy, LR_COPYFROMRESOURCE)
    If tImage = 0 Then Exit Sub
            
    If asDisabled Then
        ' get the primary brush (can be any color; don't just settle for gray)
        brGray = CreateSolidBrush(ConvertVBSysColor(PrimaryColor))
        ' get the outline/highlight brush if any. Again, any color can be used
        If Highlight <> -1 Then brLight = CreateSolidBrush(ConvertVBSysColor(Highlight))
            
        ' if the highlight brush is defined, draw it first, shifting it 1 pixel
        If brLight <> 0 Then
            DrawState hDC, brLight, 0, tImage, 0, Left + 1, Top + 1, 0, 0, DSS_MONO Or DST_ICON
        End If
        ' now draw the icon using the primary brush
        DrawState hDC, brGray, 0, tImage, 0, Left, Top, 0, 0, DSS_MONO Or DST_ICON
        
        ' clean up brushes
        DeleteObject brGray
        If brLight Then DeleteObject brLight
    Else
        ' no easier way to draw an icon than by using DrawIconEx
        DrawIconEx hDC, Left, Top, tImage, 0, 0, 0, 0, &H3
    End If
    
    ' delete the icon copy
    DestroyIcon tImage

End Sub
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.DrawLine
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Draws a single line within a DC using passed color
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub DrawLine(ByVal hDC As Long, ByVal Color As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal X1 As Long, ByVal Y1 As Long, _
        Optional ByVal PenWidth As Long = 1, Optional ByVal PenStyle As Long = 0)

    ' Parameters
    ' Color: the color for the pen
    ' X,Y: the line's starting coordinates
    ' X1,Y1: the line's ending coordinates
    ' PenWidth: width of the pen
    ' PenStyle: if PenWidth is anything but 1, APIs default to PenStyle=0 (solid)

    On Error Resume Next

    If hDC = 0 Then Exit Sub
    
    Dim hPen As Long
    
    hPen = SelectObject(hDC, CreatePen(PenStyle, PenWidth, ConvertVBSysColor(Color)))
    ' draw a simple line using APIs
    MoveToEx hDC, X, Y, ByVal 0&
    LineTo hDC, X1, Y1
    ' clean up
    DeleteObject SelectObject(hDC, hPen)

End Sub
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.DrawPolyLine
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Draws a multiple segment line using passed color
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub DrawPolyLine(ByVal hDC As Long, ByVal Color As Long, _
                            ParamArray XYcoords() As Variant)
    
    ' Parameters
    ' Color : the color of the line
    ' XYcoords are a complete set of paired X,Y coordinates for line points
    
    On Error Resume Next
    
    If hDC = 0 Then Exit Sub
    
    ' null ParamArray?
    If LBound(XYcoords) > UBound(XYcoords) Then Exit Sub
    
    Dim Pt As Long, nrPts As Long, arrBase As Long
    Dim hPen As Long, hOldPen As Long
    
    ' start parsing out the XY coordinates
    arrBase = LBound(XYcoords)
    
    For Pt = arrBase To UBound(XYcoords)
        Select Case VarType(XYcoords(Pt))
        Case vbLong, vbInteger, vbByte
        Case Else
            ' unusable, exit routine
            Exit Sub
        End Select
    Next
    
    ' calculate number of line points to draw and truncate if user passed odd number
    nrPts = UBound(XYcoords) - arrBase + 1
    If nrPts Mod 2 > 0 Then nrPts = nrPts - 1
    
    hPen = SelectObject(hDC, CreatePen(0, 1, ConvertVBSysColor(Color)))
    
    ' move to the 1st X,Y coordinates in the DC
    MoveToEx hDC, XYcoords(arrBase), XYcoords(arrBase + 1), ByVal 0&
    ' now loop thru & simply draw the polyLine
    For Pt = 1 To nrPts / 2 - 1
        LineTo hDC, XYcoords(arrBase + Pt * 2), XYcoords(arrBase + Pt * 2 + 1)
    Next
    
    ' clean up
    DeleteObject SelectObject(hDC, hPen)

End Sub
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.DrawRect
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Fills and/or Frames a rectangle
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub DrawRect(ByVal hDC As Long, ByVal Color As Long, _
        ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, _
        Optional ByVal FillColor As Long = -1)
    
    ' parameters are self explanatory with these exceptions
    ' To fill the rectangle with a color, pass the color as FillColor otherwise leave at -1
    ' To prevent drawing border around rectangle, pass -1 as the Color parameter
    
    On Error Resume Next
    
    If hDC = 0 Then Exit Sub
    
    Dim dRect As RECT, hBrush As Long
    SetRect dRect, Left, Top, Right, Bottom
    
    ' if user wants the rectangle filled, then do so
    If FillColor <> -1 Then
        hBrush = CreateSolidBrush(ConvertVBSysColor(FillColor))
        FillRect hDC, dRect, hBrush
        DeleteObject hBrush
    End If
    
    ' outline the rectangle if color <> -1
    If Color <> -1 Then
        hBrush = CreateSolidBrush(ConvertVBSysColor(Color))
        FrameRect hDC, dRect, hBrush
        DeleteObject hBrush
    End If
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.TextMeasureEx
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Returns the width & height of a passed string of text
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub TextMeasureEx(ByVal hDC As Long, ByRef Text As String, _
        ByVal FontHandle As eGFont, Optional textWidthRtn As Long, _
        Optional textHeightRtn As Long, Optional DTFlags As eDTflags, _
        Optional gdcFont As StdFont, Optional Rotation As Long = 0)
        
    ' Parameters.
    ' Text: the text you want drawn
    ' FontHandle: gDC_Font, do not select a font, use the font in the DC
    '           : gMenu_Font, get copy of the actual menu font
    '           : gTitlebar_Font, get copy of the actual titlebar font
    '           Note. Copies used in case user decides to unselect the font from the DC
    ' textWidthRtn: on exit, will contain the width needed to display the text
    ' textHeightRtn: on exit, will contain the height needed to display the text
    ' DTflags: optional and are standard DT_flag values needed by DrawText API
        
    On Error Resume Next
    
    If hDC = 0 Then Exit Sub
    
    Dim hFont As Long, dFont As Long
    Dim cClient As iImpProps, tRect As RECT
    
    ' get copy of the font if needed
    If FontHandle = gMenu_Font Or FontHandle = gTitlebar_Font Then
        GetObjectFromPointer mImplementation, cClient
        If FontHandle = gMenu_Font Then
            dFont = cClient.GetPropValue([_MenuFont])
        Else
            dFont = cClient.GetPropValue([_TitleFont])
        End If
        Set cClient = Nothing
        hFont = SelectObject(hDC, dFont)
    Else
        If Not gdcFont Is Nothing Then
            dFont = FontToHandle(gdcFont, False, Rotation)
            If dFont Then hFont = SelectObject(hDC, dFont)
        End If
    End If
    
    ' when using a null rectangle, draw text will fill in the right & bottom elements
    If (DTFlags And DT_MODIFIABLE) = DT_MODIFIABLE Then
        SetRect tRect, 0, 0, textWidthRtn, 0
    Else
        SetRect tRect, 0, 0, 0, 0
    End If
    DrawText hDC, Text, -1, tRect, DT_CALCRECT Or DTFlags
    
    ' return the results
    If Rotation = 2700 Or Rotation = 900 Then
        textHeightRtn = tRect.Right + 1
        textWidthRtn = tRect.Bottom + 1
    Else
        textWidthRtn = tRect.Right + 1
        textHeightRtn = tRect.Bottom + 1
    End If
    
    ' destroy the font if we created it
    If hFont Then DeleteObject SelectObject(hDC, hFont)
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.DrawButtonShape
' DateTime  : 9/14/2005
' Author    : LaVolpe
' Purpose   : Creates a blank button using standard system colors
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub DrawButtonShape(ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, _
                    ByVal Right As Long, ByVal Bottom As Long, ByVal btnType As eDFCtype, _
                    ByVal State As eDFCstate)
                    
    ' Parameters
    ' Top,Left,Right,Bottom comprise the boundaries of the button
    ' btnType is either a blank button or a caption button (captions are the State parameter)
    ' State are combinations of the captions and or up/down states
    '   Note: Up vs Down state simply excludes the DFCS_PUSHED flag
    
    Dim btnRect As RECT
    If hDC = 0 Then Exit Sub

    If btnType <> dfcBlankButton And btnType <> dfcCaptionButton Then Exit Sub

    SetRect btnRect, Left, Top, Right, Bottom
    DrawFrameControl hDC, btnRect, btnType, State

                    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.CreateBitmapAndDC
' DateTime  : 9/20/2005
' Author    : LaVolpe
' Purpose   : Returns a bitmap handle and or DC handle
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub CreateBitmapAndDC(ByVal Width As Long, ByVal Height As Long, _
                hBitmap As Long, hDC As Long, Optional bWantDC As Boolean = False)

    ' Parameters:
    ' Width & Height are the bitmap width & height needed.
    '           Any value < 1 will not return a bitmap handle
    ' hBitmap will be the returned handle of the bitmap if it was created
    ' hDC will be the returned handle of the DC if bWantDC=True
    ' bWantDC. True to return a DC handle in the hDC parameter, otherwise no DC is created
    
    ' To return a DC without a bitmap, simply set bWantDC=True & the width/height=0
    ' You are responsible for destroying both the bitmap and/or DC
    
    hBitmap = MakeBitmap(Width, Height, hDC, bWantDC)

End Sub
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.ConvHandleToStdPic
' DateTime  : 9/26/2005
' Author    : LaVolpe
' Purpose   : creates a stdPicture icon, cursor or bitmap from a memory handle
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function ConvHandleToStdPic(ByVal hHandle As Long) As StdPicture

    If hHandle = 0 Then Exit Function
    
    Dim bmpInfo As BITMAPINFOHEADER
    Dim iInfo As ICONINFO
    
    If GetGDIObject(hHandle, Len(bmpInfo), bmpInfo) = 0 Then
        
        ' not a bitmap, is it an icon or cursor?
        If GetIconInfo(hHandle, iInfo) = 0 Then Exit Function
        ' delete these bitmaps created by call to GetIconInfo
        If iInfo.hbmColor <> 0 Then DeleteObject iInfo.hbmColor
        If iInfo.hbmMask <> 0 Then DeleteObject iInfo.hbmMask
        
        Set ConvHandleToStdPic = HandleToPicture(hHandle, False)
    Else
        Set ConvHandleToStdPic = HandleToPicture(hHandle, True)
    
    End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.GrayScale_Image
' DateTime  : 10/04/2005
' Author    : LaVolpe
' Purpose   : Grayscale an image and draw it on a DC
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function GrayScale_Image(ByVal hHandle As Long, ByVal hDC As Long, _
                    Optional ByVal Left As Long = 0, Optional ByVal Top As Long = 0, _
                    Optional ByVal imgWidth As Long, Optional ByVal imgHeight As Long) As Boolean

    ' Parameters
    ' hHandle is the image handle (transparent GIFs will not remain transparent)
    ' hDC is the DC to draw the grayscaled image
    ' Left,Top are the left and top coordinates to draw the image
    ' imgWidth,imgHeight are the size of the drawn image. If zero passed, full image size used
    GrayScale_Image = GrayScaleImage(hHandle, hDC, Left, Top, imgWidth, imgHeight)

End Function
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.GrayScale_DC
' DateTime  : 10/04/2005
' Author    : LaVolpe
' Purpose   : Grayscale an entire DC or a portion of a DC
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub GrayScale_DC(ByVal hDC As Long, _
                    Optional ByVal Left As Long = 0, Optional ByVal Top As Long = 0, _
                    Optional ByVal Right As Long = -1, Optional ByVal Bottom As Long = -1)

    ' Parameters
    ' hDC is the DC to grayscal
    ' Left,Top are the left and top coordinates to begin grayscaling
    ' Right,Bottom are where grayscaling will stop
    '   passing -1 will use the DC's right and/or bottom coordinates
    If hDC = 0 Then Exit Sub
    GrayScaleDC hDC, 0, Left, Top, Right, Bottom

End Sub
'---------------------------------------------------------------------------------------
' Procedure : cGraphics.GrayScale_Color
' DateTime  : 10/04/2005
' Author    : LaVolpe
' Purpose   : Return the grayscale value of a non-grayscaled color
' Comments  : Passing a color already grayscaled returns the same color
'---------------------------------------------------------------------------------------
'
Public Function GrayScale_Color(ColorVal As Long) As Long
    Dim rtnColor As Long
    rtnColor = ConvertVBSysColor(ColorVal)
    GrayScaleColor VarPtr(rtnColor)
    GrayScale_Color = rtnColor
End Function


'---------------------------------------------------------------------------------------
' Procedure : cGraphics.pImplementation
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : CustomWindow class's pointer for the iImpProps interface
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Property Let pImplementation(cImplement As Long)
    mImplementation = cImplement
End Property

