VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuBarColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Class is simply a warehouse for colors for a menubar

Private cStyleFore(0 To 2, 0 To 1) As Byte ' 6bytes
'1st array elements: one for each state 0=standard,1=hover,2=select
' possible values:
'   0=default menubar colors for text
'   1=flat (1 color used)
'   2=sunken (2 colors used)
'   3=raised (2 colors used)
'   4=engraved (3 colors used)
'2nd array elements
'   0=enabled items, 1=disabled items

Private cStyleBdr(0 To 1) As Byte   ' 2bytes & apply to both enabled & disabled items
'1st array elements: one for each state 1=hover,0=select
' possible values:
'   0=default borders, 3d raised for hover & 3d sunken for select
'   1=3d raised, custom colors
'   2=3d sunken, custom colors
'   3=flat & filled (similar to Office XP)
'   4=no borders
Private cBackStyle As Byte
'0=solid, default backcolor
'1=solid, custom backcolor
'2=gradient east-west fill
'4=gradient, north-south fill
'8=transparent
'16=menubar is ownerdrawn
'32=grayscaled colors
Private cFrameStyle As Byte
'0=none
'1=flat (1 color)
'2=3d (2 colors)
Private cImgStyle As Byte            ' one of the eBackStyles enumerators; 32=grayscaled
Private cGradients(0 To 1) As Long   ' gradients element 0 is solid backcolor if no gradients used
Private cEnabled() As Long           ' 9 colors when used to display menu items in various styles in all 3 menu item states
Private cDisabled() As Long          ' 9 colors when used to display menu items in various styles in all 3 menu item states
Private cBdrColors() As Long         ' 4 colors to define hover & select border colors
Private cFrame() As Long
Private cImage As StdPicture         ' optional image for the menu bar
Private cImgAll As StdPicture
Private cUsage As Byte               '1=inactive state else active

Implements iPropertyBag

'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.ResetToSystemDefaults
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Resets colors and images to system defaults
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub ResetToSystemDefaults(ByVal Options As eBarResets)

Dim I As Integer, iOption As Integer

If Options = rstAll Then Options = rstBackGroundColors Or rstMenuSelect3Dbox Or rstNoImage Or rstTextColors

' loop thru the possible properties to see if any are reset
iOption = 2
For I = 1 To 4
    If (Options And iOption) = iOption Then
        Select Case iOption
        Case rstBackGroundColors
            cBackStyle = cBackStyle And 16
            cFrameStyle = 0
            Erase cFrame
        Case rstMenuSelect3Dbox
            Erase cBdrColors
            Erase cStyleBdr()
        Case rstNoImage
            Set cImage = Nothing
            Set cImgAll = Nothing
        Case rstTextColors
            Erase cEnabled()
            Erase cDisabled()
            Erase cStyleFore
        End Select
    End If
    iOption = iOption * 2
Next
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.Get/SetImageBackground
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Get/Set image to be used for the menubar background
' Comments  : Pass Nothing as the picture to reset. Refresh the window afterwards
'---------------------------------------------------------------------------------------
'
Public Sub GetImageBackground(mImage As StdPicture, Optional Style As eBackStyles, Optional GrayScaled As Boolean)
    Set mImage = cImage
    Style = cImgStyle And Not bfGrayScaled
    GrayScaled = ((cImgStyle And bfGrayScaled) = bfGrayScaled)
End Sub
Public Sub SetImageBackground(mImage As StdPicture, ByVal Style As eBackStyles, Optional ByVal GrayScaled As Boolean)
    If Style < bsSmartStretch Or Style > bsTiled Then Exit Sub
    cImgStyle = Style
    Set cImage = mImage
    If GrayScaled Then cImgStyle = cImgStyle Or bfGrayScaled
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.Get/SetBackColors
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Get/Set the background colors for the menubar
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub SetBackColors(ByVal Style As eFillStyle, ByVal ColorLeftTop As Long, _
                        Optional ByVal ColorRightBottom As Long)
                        
    ' set the option flag
    Me.BackStyle = Style Or (cBackStyle And 16)
    ' all colors are converted to normal RGB at here so any drawing routines
    ' don't need to check. VB system colors are negative, non RGB values.
    If Style <> bfTransparent Then
        cGradients(0) = ConvertVBSysColor(ColorLeftTop)
        If Style <> bfSolid Then cGradients(1) = ConvertVBSysColor(ColorRightBottom)
    End If

End Sub
Public Sub GetBackColors(Style As eFillStyle, ColorLeftTop As Long, _
                        Optional ColorRightBottom As Long)
    
    If (cBackStyle And Not 16) = 0 Then ' default colors
        Style = bfSolid
        ColorLeftTop = ConvertVBSysColor(vbMenuBar)
        ColorRightBottom = 0
    Else
        Style = Me.BackStyle
        ColorLeftTop = cGradients(0)
        ColorRightBottom = cGradients(1)
    End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.Get/SetEnabledTextColors
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Get/Set the menu item fore color & styles
' Comments  : Different colors & styles can be set for normal, hover or selected
'---------------------------------------------------------------------------------------
'
Public Sub GetEnabledTextColors(Effects As eFX, ByVal Action As eMnuForeColors, _
                        baseColor As Long, Optional EffectColor1 As Long, _
                        Optional EffectColor2 As Long)
    
    Select Case Action
        Case mcStandard: Action = 0
        Case mcHover: Action = 1
        Case mcSelect: Action = 2
        Case Else: Exit Sub
    End Select
    
    If IsArrayEmpty(Not cEnabled) Then
        Effects = fxDefault                 ' system defaults
    Else
        Effects = cStyleFore(Action, 0)     ' custom or system defaults
    End If
    If Effects = fxDefault Then             ' system colors
        If cUsage = 1 Then                  ' active window
            baseColor = ConvertVBSysColor(vbMenuText)
        Else
            baseColor = ConvertVBSysColor(vbGrayText)
        End If
        EffectColor1 = 0
        EffectColor1 = 0
    Else
        baseColor = cEnabled(Action, 0)
        EffectColor1 = cEnabled(Action, 1)
        EffectColor2 = cEnabled(Action, 2)
    End If

End Sub
Public Sub SetEnabledTextColors(ByVal Effects As eFX, ByVal Action As eMnuForeColors, _
                        ByVal baseColor As Long, Optional ByVal EffectColor1 As Long, _
                        Optional ByVal EffectColor2 As Long)

    If Action < mcStandard Or Action > [_mcAllActions] Then Exit Sub
    If Effects < fxDefault Or Effects > fxEngraved Then Exit Sub
    
    If IsArrayEmpty(Not cEnabled) Then ReDim cEnabled(0 To 2, 0 To 2)
    
    Select Case Action
    Case [_mcAllActions]
        SetEnabledTextColors Effects, mcHover, baseColor, EffectColor1, EffectColor2
        SetEnabledTextColors Effects, mcSelect, baseColor, EffectColor1, EffectColor2
        Action = mcStandard
    Case mcHover Or mcSelect
        SetEnabledTextColors Effects, mcHover, baseColor, EffectColor1, EffectColor2
        Action = mcSelect
    Case mcHover Or mcStandard
        SetEnabledTextColors Effects, mcHover, baseColor, EffectColor1, EffectColor2
        Action = mcStandard
    Case mcSelect Or mcStandard
        SetEnabledTextColors Effects, mcSelect, baseColor, EffectColor1, EffectColor2
        Action = mcStandard
    End Select
    
    Select Case Action
        Case mcStandard: Action = 0
        Case mcHover: Action = 1
        Case mcSelect: Action = 2
    End Select
    
    cStyleFore(Action, 0) = Effects
    ' all colors are converted to normal RGB at here so any drawing routines
    ' don't need to check. VB system colors are negative, non RGB values.
    If Effects > fxDefault Then
        cEnabled(Action, 0) = ConvertVBSysColor(baseColor)
        cEnabled(Action, 1) = ConvertVBSysColor(EffectColor1)
        cEnabled(Action, 2) = ConvertVBSysColor(EffectColor2)
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.Get/SetDisabledTextColors
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Get/Set colors and styles for disabled menu items
' Comments  : Different colors/styles can be applied to normal, hover & selected states
'---------------------------------------------------------------------------------------
'
Public Sub GetDisabledTextColors(Effects As eFX, ByVal Action As eMnuForeColors, _
                        baseColor As Long, Optional EffectColor1 As Long, _
                        Optional EffectColor2 As Long)
    
    Select Case Action
        Case mcStandard: Action = 0
        Case mcHover: Action = 1
        Case mcSelect: Action = 2
        Case Else: Exit Sub
    End Select
    
    If IsArrayEmpty(Not cDisabled) Then
        Effects = fxDefault                 ' system defaults
    Else
        Effects = cStyleFore(Action, 1)     ' custom or system defaults
    End If
    If Effects = fxDefault Then             ' system colors
        If cUsage = 1 Then                  ' active window
            baseColor = ConvertVBSysColor(vbGrayText)
            EffectColor1 = ConvertVBSysColor(vb3DHighlight)
            Effects = fxSunken
        Else
            baseColor = ConvertVBSysColor(vbGrayText)
            EffectColor1 = 0
        End If
        EffectColor2 = 0
    Else
        baseColor = cDisabled(Action, 0)
        EffectColor1 = cDisabled(Action, 1)
        EffectColor2 = cDisabled(Action, 2)
    End If

End Sub
Public Sub SetDisabledTextColors(ByVal Effects As eFX, ByVal Action As eMnuForeColors, _
                        ByVal baseColor As Long, Optional ByVal EffectColor1 As Long, _
                        Optional ByVal EffectColor2 As Long)

    If Action < mcStandard Or Action > [_mcAllActions] Then Exit Sub
    If Effects < fxDefault Or Effects > fxEngraved Then Exit Sub
    
    If IsArrayEmpty(Not cDisabled) Then ReDim cDisabled(0 To 2, 0 To 2)
    
    Select Case Action
    Case [_mcAllActions]
        SetDisabledTextColors Effects, mcHover, baseColor, EffectColor1, EffectColor2
        SetDisabledTextColors Effects, mcSelect, baseColor, EffectColor1, EffectColor2
        Action = mcStandard
    Case mcHover Or mcSelect
        SetDisabledTextColors Effects, mcHover, baseColor, EffectColor1, EffectColor2
        Action = mcSelect
    Case mcHover Or mcStandard
        SetDisabledTextColors Effects, mcHover, baseColor, EffectColor1, EffectColor2
        Action = mcStandard
    Case mcSelect Or mcStandard
        SetDisabledTextColors Effects, mcSelect, baseColor, EffectColor1, EffectColor2
        Action = mcStandard
    End Select
    
    Select Case Action
        Case mcStandard: Action = 0
        Case mcHover: Action = 1
        Case mcSelect: Action = 2
    End Select
    cStyleFore(Action, 1) = Effects
    ' all colors are converted to normal RGB at here so any drawing routines
    ' don't need to check. VB system colors are negative, non RGB values.
    If Effects > fxDefault Then
        cDisabled(Action, 0) = ConvertVBSysColor(baseColor)
        cDisabled(Action, 1) = ConvertVBSysColor(EffectColor1)
        cDisabled(Action, 2) = ConvertVBSysColor(EffectColor2)
    End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.Get/SetMenuSelectionStyle
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Get/Set the borders used for hover & selected menu items
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub SetMenuSelectionStyle(ByVal Style As eMnuBorders, ByVal selectType As e3DBox, ByVal Color1 As Long, ByVal Color2 As Long)
    
    If Style < mbDefault Or Style > mbNoBorders Then Exit Sub
    If IsArrayEmpty(Not cBdrColors) Then ReDim cBdrColors(0 To 1, 0 To 1) As Long
    
    ' all colors are converted to normal RGB at here so any drawing routines
    ' don't need to check. VB system colors are negative, non RGB values.
    cStyleBdr(Abs(selectType = bxHover)) = Style
    cBdrColors(Abs(selectType = bxHover), 0) = ConvertVBSysColor(Color1)
    cBdrColors(Abs(selectType = bxHover), 1) = ConvertVBSysColor(Color2)
    
End Sub
Public Sub GetMenuSelectionStyle(ByVal asHover As Boolean, Color1 As Long, Color2 As Long, Optional Style As eMnuBorders)
    
    If IsArrayEmpty(Not cBdrColors) Then
        ' default colors
        Style = mbDefault
    Else
        If asHover Then
            Style = cStyleBdr(1)
        Else
            Style = cStyleBdr(0)
        End If
    End If
    If Style = mbDefault Then
        If asHover Then
            Style = mbRaisedCustom
            Color1 = ConvertVBSysColor(vb3DHighlight)
            Color2 = ConvertVBSysColor(vb3DShadow)
        Else
            Style = mbSunkenCustom
            Color2 = ConvertVBSysColor(vb3DHighlight)
            Color1 = ConvertVBSysColor(vb3DShadow)
        End If
    Else
        Color1 = cBdrColors(Abs(asHover), 0)
        Color2 = cBdrColors(Abs(asHover), 1)
    End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.Get/SetFrame
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : The menubar can have a 1 or 2 color border surrounding it
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
' Procedure : clsMenuBarColors.pClassUsage
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Determines whether this class is for active or inactive window states
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Property Let pClassUsage(isActive As Boolean)
    '1=inactive state else active
    cUsage = CByte(Abs(isActive))
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.pIsImgBkg
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Called when painting the menubar. Is an image background applied?
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Property Get pIsImgBkg() As StdPicture
    If Not cImgAll Is Nothing Then
        Set pIsImgBkg = cImgAll
    End If
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.pSized
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Called when the menubar is being measured. A copy of the menubar
'             background is maintained to allow menuitem effects
'---------------------------------------------------------------------------------------
'
Friend Sub pSized(ByVal tRectPtr As Long, farExtent As Long, isMeasuring As Boolean, mBarOD As Boolean)
    
    If isMeasuring Then Me.isMenbarOwnerDrawn = mBarOD
    
    If (cImage Is Nothing) And ((cBackStyle And Not bfGrayScaled) < 8) Then Exit Sub
    
        ' when we can't accurately draw the menubar, we will cache an image of
        ' it so the image can be used when updating the menubar with hover and/or
        ' selection rectangles/text styles. This is called each time the menubar
        ' is either being measured or being drawn, although routine may abort if unneeded
        ' The following settings trigger this routine:
        ' 1. Menubar is owner drawn
        ' 2. Menubar backstyle is transparent
        ' 3. An image is used for the menubar
        
        If isMeasuring Then ' measuring
            ' when measuring, something changed & we will need a new snap shot of the menubar
            Set cImgAll = Nothing
            Exit Sub
        Else                ' drawing
            ' if we already have a snapshot, nothing to do here
            If Not cImgAll Is Nothing Then Exit Sub
        End If
        
        Dim mbRect As RECT, tDC As Long, hDC As Long
        Dim hBmp As Long, hOldBmp As Long
        Dim Cx As Long, Cy As Long, hWnd As Long
        
        ' copy the menubar rectangle & determine its height/width
        CopyMemory mbRect, ByVal tRectPtr, &H10
        ' calculate the menubar size
        Cx = farExtent - mbRect.Left
        Cy = mbRect.Bottom - mbRect.Top
        ' create a bitmap to paint the menubar into
        hBmp = MakeBitmap(Cx, Cy)
        ' select the bitmap into the DC & get the primary global-use DC
        tDC = Canvas(False, True)
        hOldBmp = SelectObject(tDC, hBmp)
        hDC = Canvas(True)
        ' simply blt the DC contents
        BitBlt tDC, 0, 0, Cx, Cy, hDC, mbRect.Left, mbRect.Top, vbSrcCopy
        ' unselect the bitmap
        SelectObject tDC, hOldBmp
        ' convert the bitmap into a stdPicture object
        Set cImgAll = HandleToPicture(hBmp, True)
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.BackStyle
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Various backstyle settings
'---------------------------------------------------------------------------------------
'
Public Property Let BackStyle(Style As eFillStyle)
    If Style < bfSolid Then Exit Property
    Select Case (Style And Not bfGrayScaled)
    Case bfSolid, bfGradientEW: cBackStyle = Style + 1 Or (cBackStyle And 16) Or (Style And bfGrayScaled)
    Case bfGradientNS: cBackStyle = 4 Or (cBackStyle And 16) Or (Style And bfGrayScaled)
    Case bfTransparent: cBackStyle = 8 Or (cBackStyle And 16)
    End Select
    If Not cImgAll Is Nothing Then
        If Style < bfTransparent And cImage Is Nothing Then Set cImgAll = Nothing
    End If
End Property
Public Property Get BackStyle() As eFillStyle
    Select Case ((cBackStyle And Not 16) And Not bfGrayScaled)
    Case 1, 2: BackStyle = (cBackStyle And Not 16) - 1 Or (cBackStyle And bfGrayScaled)
    Case 4: BackStyle = bfGradientNS Or (cBackStyle And bfGrayScaled)
    Case 8: BackStyle = bfTransparent
    Case Else: BackStyle = bfGradientEW
    End Select
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.isMenbarOwnerDrawn
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Helper property to return if user is painting the menubar
'---------------------------------------------------------------------------------------
'
Friend Property Let isMenbarOwnerDrawn(isOD As Boolean)
    If isOD Then
        cBackStyle = cBackStyle Or 16
    Else
        cBackStyle = cBackStyle And Not 16
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : clsMenuBarColors.iPropertyBag_Read/WriteProperites
' DateTime  : 9/26/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve custom properites
' Comments  : Used by Export/ImportCustomSettings
'---------------------------------------------------------------------------------------
'
Private Sub iPropertyBag_ReadProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim sProp As String
    Dim sProps() As String
    Dim I As Integer, J As Integer
    
    With propBag
    
        sProp = .ReadProperty("MB" & cUsage & ".StyleFore", "000000")
        For I = 0 To 2
            cStyleFore(I, 0) = CByte(Mid$(sProp, I * 2 + 1, 1))
            cStyleFore(I, 1) = CByte(Mid$(sProp, I * 2 + 2, 1))
        Next
        
        sProp = .ReadProperty("MB" & cUsage & ".StyleBdr", "00")
        cStyleBdr(0) = CByte(Mid$(sProp, 1, 1))
        cStyleBdr(1) = CByte(Mid$(sProp, 2, 1))
        
        sProp = .ReadProperty("MB" & cUsage & ".StyleMisc", "000")
        cBackStyle = CByte(Mid$(sProp, 1, 1))
        cFrameStyle = CByte(Mid$(sProp, 2, 1))
        cImgStyle = CByte(Mid$(sProp, 3))
        
        sProps = Split(.ReadProperty("MB" & cUsage & ".ColorG", "0,0"), ",")
        cGradients(0) = CLng(sProps(0))
        cGradients(1) = CLng(sProps(1))
        
        sProp = .ReadProperty("MB" & cUsage & ".ColorE", "")
        If Len(sProp) Then
            sProps = Split(sProp, ",")
            ReDim cEnabled(0 To 2, 0 To 2)
            For I = 0 To 2
                For J = 0 To 2
                    cEnabled(I, J) = CLng(sProps(I * 3 + J))
                Next
            Next
        Else
            Erase cEnabled()
        End If
        
        sProp = .ReadProperty("MB" & cUsage & ".ColorD", "")
        If Len(sProp) > 0 Then
            sProps = Split(sProp, ",")
            ReDim cDisabled(0 To 2, 0 To 2)
            For I = 0 To 2
                For J = 0 To 2
                    cDisabled(I, J) = CLng(sProps(I * 3 + J))
                Next
            Next
        Else
            Erase cDisabled()
        End If
        
        sProp = .ReadProperty("MB" & cUsage & ".ColorB", "")
        If Len(sProp) > 0 Then
            sProps = Split(sProp, ",")
            ReDim cBdrColors(0 To 1, 0 To 1)
            For I = 0 To 1
                For J = 0 To 1
                    cBdrColors(I, J) = CLng(sProps(I * 2 + J))
                Next
            Next
        Else
            Erase cBdrColors()
        End If
        
        sProp = .ReadProperty("MB" & cUsage & ".ColorF", "")
        If Len(sProp) > 0 Then
            sProps = Split(sProp, ",")
            ReDim cFrame(0 To 1)
            For I = 0 To 1
                cFrame(I) = CLng(sProps(I))
            Next
        Else
            Erase cFrame()
        End If
        
        If (Limits And exBkgImages) = exBkgImages Then
            If cUsage = 1 Then ' active class, load the bkg image if any
                Set cImage = .ReadProperty("MB" & cUsage & ".Image", Nothing)
                lParam = .ReadProperty("MB" & cUsage & ".ImagePtr", 0)
            Else
                ' inactive class, load bkg image only if it is not the same as active class
                If .ReadProperty("MB" & cUsage & ".ImagePtr", 0) <> lParam Then
                    Set cImage = .ReadProperty("MB" & cUsage & ".Image", Nothing)
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
    
    ' the text styles for all 3 menu item states
    For I = 0 To 2
        sProp = sProp & CStr(cStyleFore(I, 0)) & CStr(cStyleFore(I, 1))
    Next
    propBag.WriteProperty "MB" & cUsage & ".StyleFore", sProp, "000000"
    ' the border, back, frame & image styles
    propBag.WriteProperty "MB" & cUsage & ".StyleBdr", CStr(cStyleBdr(0)) & CStr(cStyleBdr(1)), "00"
    propBag.WriteProperty "MB" & cUsage & ".StyleMisc", CStr(cBackStyle) & CStr(cFrameStyle) & CStr(cImgStyle), "000"
    ' the gradients used
    propBag.WriteProperty "MB" & cUsage & ".ColorG", CStr(cGradients(0)) & "," & CStr(cGradients(1)), "0,0"
    
    sColors = ""    ' the enabled text colors
    If Not IsArrayEmpty(Not cEnabled) Then
        For I = 0 To 2
            For J = 0 To 2
                sColors = sColors & "," & CStr(cEnabled(I, J))
            Next
        Next
    End If
    propBag.WriteProperty "MB" & cUsage & ".ColorE", Mid$(sColors, 2), ""
    
    sColors = ""    ' the disabled text colors
    If Not IsArrayEmpty(Not cDisabled) Then
        For I = 0 To 2
            For J = 0 To 2
                sColors = sColors & "," & CStr(cDisabled(I, J))
            Next
        Next
    End If
    propBag.WriteProperty "MB" & cUsage & ".ColorD", Mid$(sColors, 2), ""
    
    sColors = ""    ' the border colors
    If Not IsArrayEmpty(Not cBdrColors) Then
        For I = 0 To 1
            For J = 0 To 1
                sColors = sColors & "," & CStr(cBdrColors(I, J))
            Next
        Next
    End If
    propBag.WriteProperty "MB" & cUsage & ".ColorB", Mid$(sColors, 2), ""
    
    sColors = ""    ' the frame colors
    If Not IsArrayEmpty(Not cFrame) Then
        For I = 0 To 1
            sColors = sColors & "," & CStr(cFrame(I))
        Next
    End If
    propBag.WriteProperty "MB" & cUsage & ".ColorF", Mid$(sColors, 2), ""
    
    If (Limits And exBkgImages) = exBkgImages Then
        ' finally the menubar image
        If Not cImage Is Nothing Then
            propBag.WriteProperty "MB" & cUsage & ".ImagePtr", cImage.Handle, 0
            If cUsage = 1 Then ' active class
                propBag.WriteProperty "MB" & cUsage & ".Image", cImage
                lParam = cImage.Handle
            Else
                'if this inactive class is using the same picture, ref the picture
                'vs saving it again. Wasted resources when the settings are imported
                'since each class will create a separate instance of the picture
                If cImage.Handle <> lParam Then
                    'different images being used, if an image is used at all
                   propBag.WriteProperty "MB" & cUsage & ".Image", cImage
                End If
            End If
        End If
    End If

End Sub
