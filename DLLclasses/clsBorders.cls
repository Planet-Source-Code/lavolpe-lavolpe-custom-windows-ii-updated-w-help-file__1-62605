VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' This class bascially contains color settings for the frame borders and background
' and possibly an optional image for the background.

Private cOptions As Byte
'1=custom border colors, 2=custom NCinset color, 4=transparent inset, 8=grayscaled image
Private cImgStyle As Byte

Private cInsetColor As Long         ' inset color, overrides cBackcolor
Private cBackColor As Long          ' the nonclient backcolor
Private cImage As StdPicture        ' inset image
Private cCustomColors() As Long     ' border colors for the window frame
Private cUsage As Byte              ' 1=active else inactive

Implements iPropertyBag

'---------------------------------------------------------------------------------------
' Procedure : clsBorders.ResetToSystemDefaults
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Resets colors to system defaults
'---------------------------------------------------------------------------------------
'
Public Sub ResetToSystemDefaults(ByVal Options As eBarResets)

    If Options = rstAll Then Options = rstBorderColors Or rstNoImage
    If (Options And rstBorderColors) = rstBorderColors Then
        Erase cCustomColors()
        cOptions = cOptions And 8
        cBackColor = ConvertVBSysColor(vbButtonFace)
        cInsetColor = -1
    End If
    If (Options And rstNoImage) = rstNoImage Then
        cOptions = cOptions And Not 8
        Set cImage = Nothing
        cImgStyle = bsSmartStretch
    End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsBorders.Get/SetBorderColors
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Allows user to customize the 4 pixel, 3D border around the window
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub GetBorderColors(Face As Long, Highlight As Long, _
                            Shadow As Long, DarkShadow As Long)
                            
    If (cOptions Or 1) = cOptions Then
        Face = cCustomColors(0)
        Highlight = cCustomColors(1)
        Shadow = cCustomColors(2)
        DarkShadow = cCustomColors(3)
    Else
        Face = ConvertVBSysColor(vbButtonFace)
        Highlight = ConvertVBSysColor(vb3DHighlight)
        Shadow = ConvertVBSysColor(vb3DShadow)
        DarkShadow = ConvertVBSysColor(vb3DDKShadow)
    End If
End Sub
Public Sub SetBorderColors(ByVal Face As Long, ByVal Highlight As Long, _
                            ByVal Shadow As Long, ByVal DarkShadow As Long)

    ' all colors are converted to normal RGB at here so any drawing routines
    ' don't need to check. VB system colors are negative, non RGB values.
    ReDim cCustomColors(0 To 3)
    cCustomColors(0) = ConvertVBSysColor(Face)
    cCustomColors(1) = ConvertVBSysColor(Highlight)
    cCustomColors(2) = ConvertVBSysColor(Shadow)
    cCustomColors(3) = ConvertVBSysColor(DarkShadow)
    cOptions = cOptions Or 1

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsBorders.Get/SetNCBackColor
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Backcolor to be used to fill any inset/gaps between the titlebar/menubar
'             and the client rectangle. A minimum of 1 pixel is used to separate the
'             client rectangle from the menubar and/or titlebar edges
' Comments  : This overrides the frame's default nonclient backcolor
'---------------------------------------------------------------------------------------
'
Public Function GetInsetBackColor(Optional isTransparent As Boolean) As Long
    If cInsetColor = -1 Then
        GetInsetBackColor = cBackColor
    Else
        GetInsetBackColor = cInsetColor
    End If
    isTransparent = ((cOptions And 4) = 4)
End Function
Public Sub SetInsetBackColor(ByVal lColor As Long, Optional ByVal bReset As Boolean, _
                        Optional ByVal isTransparent As Boolean)
    
    ' all colors are converted to normal RGB at here so any drawing routines
    ' don't need to check. VB system colors are negative, non RGB values.
    If bReset Then
        cOptions = cOptions And Not 2
        cInsetColor = -1
    Else
        cOptions = cOptions Or 2
        cInsetColor = ConvertVBSysColor(lColor)
    End If
    If isTransparent Then
        cOptions = cOptions Or 4
    Else
        cOptions = cOptions And Not 4
    End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsBorders.BackColor
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve the nonclient background color
' Comments  : The Inset color overrides this color.
'             see Active/Inactive.SetInsetBackColor
'---------------------------------------------------------------------------------------
'
Public Property Get BackColor() As Long
    BackColor = cBackColor
End Property
Public Property Let BackColor(Color As Long)
    cBackColor = ConvertVBSysColor(Color)
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsBorders.Get/SetInsetImage
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Get/Set image to be used for optional inset offset
' Comments  : Pass Nothing as the picture to reset. Refresh the window afterwards
'---------------------------------------------------------------------------------------
'
Public Sub SetInsetImage(mImage As StdPicture, ByVal Style As eBackStyles, Optional ByVal GrayScaled As Boolean)
    If Style < bsSmartStretch Or Style > bsTiled Then Style = bsSmartStretch
    cImgStyle = Style
    Set cImage = mImage
    If GrayScaled Then
        cOptions = cOptions Or 8
    Else
        cOptions = cOptions And Not 8
    End If
End Sub
Public Sub GetInsetImage(mImage As StdPicture, Optional Style As eBackStyles, Optional GrayScaled As Boolean)
    Set mImage = cImage
    Style = cImgStyle
    GrayScaled = ((cOptions And 8) = 8)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : clsBorders.Class_Initialize, Class_Terminate
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Intialization & clean up
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    cBackColor = ConvertVBSysColor(vbButtonFace)
    cInsetColor = -1
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsBorders.pIsCustomBorder
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Helper function to determine if border is default or not
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Property Get pIsCustomBorder() As Boolean
    pIsCustomBorder = ((cOptions And 1) = 1)
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsBorders.pClassUsage
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Determines whether this class is for active or inactive window states
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Property Let pClassUsage(isActive As Boolean)
    '1=active state else active
    cUsage = CByte(Abs(isActive))
End Property


'---------------------------------------------------------------------------------------
' Procedure : clsBorders.Read/WriteProperites
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
        sProp = .ReadProperty("FM" & cUsage & ".ColorF", "")
        If sProp = "" Then
            Erase cCustomColors
        Else
            sProps = Split(sProp, ",")
            ReDim cCustomColors(0 To 3)
            For I = 0 To 3
                cCustomColors(I) = CLng(sProps(I))
            Next
        End If
        cImgStyle = .ReadProperty("FM" & cUsage & ".StyleMisc", 0)
        sProps = Split(.ReadProperty("FM" & cUsage & ".ColorB", "0,0"), ",")
        cBackColor = CLng(sProps(0))
        cInsetColor = CLng(sProps(1))
        cOptions = .ReadProperty("FM" & cUsage & ".Options", 0)
        
        If (Limits And exBkgImages) = exBkgImages Then
            If cUsage = 1 Then ' active class, load the bkg image if any
                Set cImage = .ReadProperty("FM" & cUsage & ".Image", Nothing)
                lParam = .ReadProperty("FM" & cUsage & ".ImagePtr", 0)
            Else
                ' inactive class, load bkg image only if it is not the same as active class
                If .ReadProperty("FM" & cUsage & ".ImagePtr", 0) <> lParam Then
                    Set cImage = .ReadProperty("FM" & cUsage & ".Image", Nothing)
                    lParam = 0 ' prevents parent class from setting our image from the active class
                End If
            End If
        End If
        
    End With
End Sub

Private Sub iPropertyBag_WriteProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)
    ' update/add properties to a property bag for the inactive/active menubar

    Dim sProp As String
    Dim I As Integer
    
    ' the frame colors
    If Not IsArrayEmpty(Not cCustomColors) Then
        For I = 0 To 3
            sProp = sProp & "," & CStr(cCustomColors(I))
        Next
    End If
    propBag.WriteProperty "FM" & cUsage & ".ColorF", Mid$(sProp, 2), ""
    ' the image styles
    propBag.WriteProperty "FM" & cUsage & ".StyleMisc", cImgStyle, 0
    ' the backcolors
    propBag.WriteProperty "FM" & cUsage & ".ColorB", CStr(cBackColor) & "," & CStr(cInsetColor), CStr(vbButtonFace) & ",-1"
    ' the options
    propBag.WriteProperty "FM" & cUsage & ".Options", cOptions, 0
    
    If (Limits And exBkgImages) = exBkgImages Then
        ' finally the inset image
        If Not cImage Is Nothing Then
            propBag.WriteProperty "FM" & cUsage & ".ImagePtr", cImage.Handle, 0
            If cUsage = 1 Then ' active class
                propBag.WriteProperty "FM" & cUsage & ".Image", cImage
                lParam = cImage.Handle
            Else
                'if this inactive class is using the same picture, ref the picture
                'vs saving it again. Wasted resources when the settings are imported
                'since each class will create a separate instance of the picture
                If cImage.Handle <> lParam Then
                    'different images being used, if an image is used at all
                   propBag.WriteProperty "FM" & cUsage & ".Image", cImage
                End If
            End If
        End If
    End If
End Sub
