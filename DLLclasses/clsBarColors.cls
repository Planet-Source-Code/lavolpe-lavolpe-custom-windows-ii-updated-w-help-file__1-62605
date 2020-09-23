VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Class is simply a warehouse for colors for a titlebar


' don't think there is a VB color constant for these
Private Const COLOR_GRADIENTINACTIVECAPTION As Long = 28
'   if it existed a VB equivalent would be -2147483620
Private Const COLOR_GRADIENTACTIVECAPTION As Long = 27
'   if it existed a VB equivalent would be -2147483621


Private cStyleFore As Byte
' possible values:
'   0=default titlebar colors for text
'   1=flat (1 color used)
'   2=sunken (2 colors used)
'   3=raised (2 colors used)
'   4=engraved (3 colors used)
Private cBackStyle As Byte
'0=solid, default gradient backcolors
'1=solid, custom backcolor
'2=gradient east-west fill
'4=gradient, north-south fill
'8=transparent
'32=grayscaled

Private cFrameStyle As Byte
'1=has flat frame (1 color)
'2=has 3D frame (2 colors)

Private cImgStyle As Byte
' one of the eBackStyles enumerators
' 32 = grayscaled image

Private cGradients(0 To 1) As Long   ' gradients element 0 is solid backcolor if no gradients used
Private cFore() As Long              ' 9 colors when used to display menu items in various styles in all 3 menu item states
Private cFrame() As Long             ' 2 colors used to draw a frame around the menubar
Private cImage As StdPicture         ' optional image for the menu bar
Private cUsage As Byte               ' 1=active state else inactive

Implements iPropertyBag

'---------------------------------------------------------------------------------------
' Procedure : clsBarColors.ResetToSystemDefaults
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Reset colors, images to system defaults
'---------------------------------------------------------------------------------------
'
Public Sub ResetToSystemDefaults(ByVal Options As eBarResets)

    Dim I As Integer, iOption As Integer
    If Options = rstAll Then Options = rstBackGroundColors Or rstNoImage Or rstTextColors
    
    iOption = 2
    ' loop thru possible options & reset as needed
    For I = 1 To 4
        If (Options And iOption) = iOption Then
            Select Case iOption
            Case rstBackGroundColors
                cBackStyle = 0
                cFrameStyle = 0
                Erase cFrame
            Case rstNoImage
                Set cImage = Nothing
                cImgStyle = 0
            Case rstTextColors
                Erase cFore()
                cStyleFore = 0
            End Select
        End If
        iOption = iOption * 2
    Next
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsBarColors.Get/SetImageBackground
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Sets or retrieves the title bar image
'---------------------------------------------------------------------------------------
'
Public Sub GetImageBackground(mImage As StdPicture, Optional Style As eBackStyles, Optional GrayScaled As Boolean)
    Set mImage = cImage
    Style = cImgStyle And Not bfGrayScaled
    GrayScaled = ((cImgStyle And bfGrayScaled) = bfGrayScaled)
'    If Not cImage Is Nothing Then Debug.Print "image "; cImage.Handle; Style; GrayScaled
End Sub
Public Sub SetImageBackground(mImage As StdPicture, ByVal Style As eBackStyles, Optional ByVal GrayScaled As Boolean)
    If Style < bsSmartStretch Or Style > bsTiled Then Exit Sub
    cImgStyle = Style
    Set cImage = mImage
    If GrayScaled Then cImgStyle = cImgStyle Or bfGrayScaled
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsBarColors.Get/SetBackColors
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Sets or retrieves the title bar back colors & style
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub SetBackColors(ByVal Style As eFillStyle, ByVal ColorLeftTop As Long, _
                            Optional ByVal ColorRightBottom As Long)
                        
    ' set the option flag
    cBackStyle = 1
    Me.BackStyle = Style
    
    ' all colors are converted to normal RGB here so any drawing routines
    ' don't need to check. VB system colors are negative, non RGB values.
    If Style <> bfTransparent Then
        cGradients(0) = ConvertVBSysColor(ColorLeftTop)
        If (Style And bfSolid) = 0 Then cGradients(1) = ConvertVBSysColor(ColorRightBottom)
    End If

End Sub
Public Sub GetBackColors(Style As eFillStyle, ColorLeftTop As Long, _
                    Optional ColorRightBottom As Long)
    
    If pIsDefaultGradient Then  ' default colors
        Style = bfGradientEW
        If cUsage Then
            ColorLeftTop = ConvertVBSysColor(vbActiveTitleBar)
            ColorRightBottom = GetSysColor(COLOR_GRADIENTACTIVECAPTION)
        Else
            ColorLeftTop = ConvertVBSysColor(vbInactiveTitleBar)
            ColorRightBottom = GetSysColor(COLOR_GRADIENTINACTIVECAPTION)
        End If
    Else
        Style = Me.BackStyle
        ColorLeftTop = cGradients(0)
        ColorRightBottom = cGradients(1)
    End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsBarColors.Get/SetTextColors
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Sets or retrieves the title bar text colors and style
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub SetTextColors(ByVal Effects As eFX, ByVal baseColor As Long, _
            Optional ByVal EffectColor1 As Long, Optional ByVal EffectColor2 As Long)

    If Effects < fxDefault Or Effects > fxEngraved Then Exit Sub
    
    If IsArrayEmpty(Not cFore) Then ReDim cFore(0 To 2)
    
    cStyleFore = Effects
    ' all colors are converted to normal RGB at here so any drawing routines
    ' don't need to check. VB system colors are negative, non RGB values.
    If Effects > fxDefault Then
        cFore(0) = ConvertVBSysColor(baseColor)
        cFore(1) = ConvertVBSysColor(EffectColor1)
        cFore(2) = ConvertVBSysColor(EffectColor2)
    End If
End Sub
Public Sub GetTextColors(Effects As eFX, baseColor As Long, _
        Optional EffectColor1 As Long, Optional EffectColor2 As Long)
    
    If IsArrayEmpty(Not cFore) Then
        Effects = fxDefault                 ' system defaults
    Else
        Effects = cStyleFore                ' custom or system defaults
    End If
    If Effects = fxDefault Then             ' system colors
        If cUsage = 1 Then                  ' active window
            baseColor = ConvertVBSysColor(vbActiveTitleBarText)
        Else
            baseColor = ConvertVBSysColor(vbInactiveTitleBarText)
        End If
        EffectColor1 = 0
        EffectColor1 = 0
    Else
        baseColor = cFore(0)
        EffectColor1 = cFore(1)
        EffectColor2 = cFore(2)
    End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsBarColors.pClassUsage
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Identifies whether this class is used for active/inactive windows
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Property Let pClassUsage(isActive As Boolean)
    '1=active state else inactive
    cUsage = CByte(Abs(isActive))
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsBarColors.pIsDefaultGradient
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Helper function used when drawing. Is the titlebar default gradient?
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Property Get pIsDefaultGradient() As Boolean
    pIsDefaultGradient = (cBackStyle = 0)
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsBarColors.BackStyle
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve the title bar back style
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Let BackStyle(Style As eFillStyle)
    If Style < bfSolid Then Exit Property
    Select Case (Style And Not bfGrayScaled)
        Case bfSolid: cBackStyle = Style + 1
        Case bfGradientEW:
            If cBackStyle <> 0 Then cBackStyle = Style + 1 Or (Style And bfGrayScaled)
        Case bfGradientNS:
            If cBackStyle <> 0 Then cBackStyle = 4 Or (Style And bfGrayScaled)
        Case bfTransparent: cBackStyle = 8
            '^^ transparent will not fill the titlebar
        Case Else: cBackStyle = 0
    End Select
    
End Property
Public Property Get BackStyle() As eFillStyle
    If cBackStyle = 0 Then
        BackStyle = bfGradientEW
    Else
        Select Case (cBackStyle And Not bfGrayScaled)
        Case 1: BackStyle = bfSolid Or (cBackStyle And bfGrayScaled)
        Case 2: BackStyle = bfGradientEW Or (cBackStyle And bfGrayScaled)
        Case 4: BackStyle = bfGradientNS Or (cBackStyle And bfGrayScaled)
        Case 8: BackStyle = bfTransparent
        End Select
    End If
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsBarColors.Get/SetFrame
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve an optional 1-pixel frame/border around the title bar
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub SetFrame(ByVal fType As e3Dframe, Optional ByVal Color1 As Long, Optional ByVal Color2 As Long)
    
    If fType < bxNone Or fType > bx3D Then Exit Sub

    cFrameStyle = fType
    If fType = 0 Then
        Erase cFrame
    Else
        ' all colors are converted to normal RGB at here so any drawing routines
        ' don't need to check. VB system colors are negative, non RGB values.
        ReDim cFrame(0 To 1)
        cFrame(0) = ConvertVBSysColor(Color1)
        cFrame(1) = ConvertVBSysColor(Color2)
    End If
End Sub
Public Sub GetFrame(fType As e3Dframe, Optional Color1 As Long, Optional Color2 As Long)
    
    Select Case cFrameStyle
    Case bx3D
        Color1 = cFrame(0)
        Color2 = cFrame(1)
    Case bxFlat
        Color1 = cFrame(0)
    End Select
    fType = cFrameStyle

End Sub


'---------------------------------------------------------------------------------------
' Procedure : clsBarColors.Read/WriteProperites
' DateTime  : 9/26/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve custom properites
' Comments  : Used by Export/ImportCustomSettings
'---------------------------------------------------------------------------------------
'
Private Sub iPropertyBag_ReadProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim sProp As String
    Dim sProps() As String
    Dim I As Integer
    
    With propBag
    
        sProp = .ReadProperty("TB" & cUsage & ".StyleMisc", "0000")
        cBackStyle = CByte(Mid$(sProp, 1, 1))
        cFrameStyle = CByte(Mid$(sProp, 2, 1))
        cStyleFore = CByte(Mid$(sProp, 3, 1))
        cImgStyle = CByte(Mid$(sProp, 4))
        
        sProps = Split(.ReadProperty("TB" & cUsage & ".ColorG", "0,0"), ",")
        cGradients(0) = CLng(sProps(0))
        cGradients(1) = CLng(sProps(1))
        
        sProp = .ReadProperty("TB" & cUsage & ".ColorE", "")
        If Len(sProp) > 0 Then
            sProps = Split(sProp, ",")
            ReDim cFore(0 To 2)
            For I = 0 To 2
                cFore(I) = CLng(sProps(I))
            Next
        Else
            Erase cFore()
        End If
        
        sProp = .ReadProperty("TB" & cUsage & ".ColorF", "")
        If Len(sProp) > 0 Then
            sProps = Split(sProp, ",")
            ReDim cFrame(0 To 2)
            For I = 0 To 2
                cFrame(I) = CLng(sProps(I))
            Next
        Else
            Erase cFrame()
        End If
        
        If (Limits And exBkgImages) = exBkgImages Then
            If cUsage = 1 Then ' active class, load the bkg image if any
                Set cImage = .ReadProperty("TB" & cUsage & ".Image", Nothing)
                lParam = .ReadProperty("TB" & cUsage & ".ImagePtr", 0)
            Else
                ' inactive class, load bkg image only if it is not the same as active class
                If .ReadProperty("TB" & cUsage & ".ImagePtr", 0) <> lParam Then
                    Set cImage = .ReadProperty("TB" & cUsage & ".Image", Nothing)
                    lParam = 0 ' prevents parent class from setting our image from the active class
                End If
            End If
        End If
        
    End With
    
End Sub

Private Sub iPropertyBag_WriteProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)
    ' update/add properties to a property bag for the inactive/active menubar

    Dim sProp As String, sColors As String
    Dim I As Integer, J As Integer
    
    ' the back, frame & image , text styles
    propBag.WriteProperty "TB" & cUsage & ".StyleMisc", CStr(cBackStyle) & CStr(cFrameStyle) & CStr(cStyleFore) & CStr(cImgStyle), "0000"
    ' the gradients used
    propBag.WriteProperty "TB" & cUsage & ".ColorG", CStr(cGradients(0)) & "," & CStr(cGradients(1)), "0,0"
    
    sColors = ""    ' the text colors
    If Not IsArrayEmpty(Not cFore) Then
        For I = 0 To 2
            sColors = sColors & "," & CStr(cFore(I))
        Next
    End If
    propBag.WriteProperty "TB" & cUsage & ".ColorE", Mid$(sColors, 2), ""
    
    sColors = ""    ' the frame colors
    If Not IsArrayEmpty(Not cFrame) Then
        For I = 0 To 1
            sColors = sColors & "," & CStr(cFrame(I))
        Next
    End If
    propBag.WriteProperty "TB" & cUsage & ".ColorF", Mid$(sColors, 2), ""
    ' finally the titlebar image
    
    If (Limits And exBkgImages) = exBkgImages Then
        If Not cImage Is Nothing Then
            propBag.WriteProperty "TB" & cUsage & ".ImagePtr", cImage.Handle, 0
            If cUsage = 1 Then ' active class
                propBag.WriteProperty "TB" & cUsage & ".Image", cImage
                lParam = cImage.Handle
            Else
                'if this inactive class is using the same picture, ref the picture
                'vs saving it again. Wasted resources when the settings are imported
                'since each class will create a separate instance of the picture
                If cImage.Handle <> lParam Then
                    'different images being used, if an image is used at all
                   propBag.WriteProperty "TB" & cUsage & ".Image", cImage
                End If
            End If
        End If
    End If

End Sub
