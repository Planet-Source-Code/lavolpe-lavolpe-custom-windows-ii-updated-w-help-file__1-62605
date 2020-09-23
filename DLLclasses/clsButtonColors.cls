VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsButtonColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' This class bascially contains color settings for the system button frame borders
' background, and for colors

Private cCustomColors() As Long     ' button edge colors & background
Private cForeColor() As Long        ' button forecolor/pen
Private cUsage As Byte              ' 1=active else inactive

Implements iPropertyBag

'---------------------------------------------------------------------------------------
' Procedure : clsButtonColors.ResetToSystemDefaults
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Resets colors to system defaults
'---------------------------------------------------------------------------------------
'
Public Sub ResetToSystemDefaults(ByVal Options As eBarResets)

    If Options = rstAll Or Options = rstButtonColors Then
        Erase cCustomColors()
        Erase cForeColor()
    End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtonColors.Get/SetBorderColors
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Allows user to customize the 2 pixel, 3D border around the system button
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub GetBorderColors(Face As Long, Highlight As Long, _
                            Shadow As Long, DarkShadow As Long)
                            
    If IsArrayEmpty(Not cCustomColors) Then
        Face = ConvertVBSysColor(vbButtonFace)
        Highlight = ConvertVBSysColor(vb3DHighlight)
        Shadow = ConvertVBSysColor(vb3DShadow)
        DarkShadow = ConvertVBSysColor(vb3DDKShadow)
    Else
        Face = cCustomColors(0)
        Highlight = cCustomColors(1)
        Shadow = cCustomColors(2)
        DarkShadow = cCustomColors(3)
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

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtonColors.Get/SetForeColors
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Sets or retrieves the system button fore and border colors
' Comments  : if DisabledSingleColor=True, then button drawn with a single color
'---------------------------------------------------------------------------------------
'
Public Sub SetForeColors(ByVal EnabledColor As Long, ByVal DisabledColor As Long, _
                ByVal DisabledHighlight As Long, Optional ByVal DisabledSingleColor As Boolean = False, _
                Optional ByVal HoverColor As Long = -1)
                
    ReDim cForeColor(0 To 3)
    cForeColor(0) = ConvertVBSysColor(EnabledColor)
    cForeColor(1) = ConvertVBSysColor(DisabledColor)
    
    If DisabledSingleColor = True Then
        ' don't use 2 colors for disabled buttons
        cForeColor(2) = -1
    Else
        cForeColor(2) = ConvertVBSysColor(DisabledHighlight)
    End If
    If HoverColor <> -1 Then cForeColor(3) = ConvertVBSysColor(HoverColor)
    
End Sub
Public Sub GetForeColors(EnabledColor As Long, DisabledColor As Long, _
                DisabledHighlight As Long, Optional DisabledSingleColor As Boolean, _
                Optional HoverColor As Long)
                
    If IsArrayEmpty(Not cForeColor) Then
        EnabledColor = ConvertVBSysColor(vbButtonText)
        DisabledColor = ConvertVBSysColor(vbGrayText)
        DisabledHighlight = ConvertVBSysColor(vb3DHighlight)
        HoverColor = -1
        DisabledSingleColor = False
    Else
        EnabledColor = cForeColor(0)
        DisabledColor = cForeColor(1)
        DisabledHighlight = cForeColor(2)
        DisabledSingleColor = (cForeColor(2) = -1)
        HoverColor = cForeColor(3)
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtonColors.pUsage
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Identifies whether this class is used for active/inactive windows
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Property Let pUsage(isActive As Byte)
    cUsage = isActive
End Property


'---------------------------------------------------------------------------------------
' Procedure : clsButtonColors.Read/WriteProperites
' DateTime  : 9/26/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve custom properites
' Comments  : Used by Export/ImportCustomSettings
'---------------------------------------------------------------------------------------
'
Private Sub iPropertyBag_ReadProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim I As Integer
    Dim sProp As String, sProps() As String

    sProp = propBag.ReadProperty("BN" & cUsage & ".ColorB", "")
    If sProp = "" Then
        Erase cCustomColors
    Else
        sProps = Split(sProp, ",")
        ReDim cCustomColors(0 To 3)
        For I = 0 To 3
            cCustomColors(I) = CLng(sProps(I))
        Next
    End If
    
    sProp = propBag.ReadProperty("BN" & cUsage & ".ColorF", "")
    If sProp = "" Then
        Erase cForeColor
    Else
        sProps = Split(sProp, ",")
        ReDim cForeColor(0 To 3)
        For I = 0 To 3
            cForeColor(I) = CLng(sProps(I))
        Next
    End If
    
End Sub

Private Sub iPropertyBag_WriteProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim I As Integer, sProp As String
    
    If Not IsArrayEmpty(Not cCustomColors) Then
        For I = 0 To 3
            sProp = sProp & "," & CStr(cCustomColors(I))
        Next
    End If
    propBag.WriteProperty "BN" & cUsage & ".ColorB", Mid$(sProp, 2), ""
    
    sProp = ""
    If Not IsArrayEmpty(Not cForeColor) Then
        For I = 0 To 3
            sProp = sProp & "," & CStr(cForeColor(I))
        Next
    End If
    propBag.WriteProperty "BN" & cUsage & ".ColorF", Mid$(sProp, 2), ""

End Sub

