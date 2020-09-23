VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMinMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Module does a lot of work for you. Basically it allows you to dictate the
' sizing constraints of any window passed. And is 100% customizable.
' Although untested on virtual screens (multi-monitors), logically the
' sizing algorithms should work on those systems also.
' Will update when known for sure.

' You can restrict how small the window can be sized to, how large it can be sized to,
' what its maximized window size will be and where the maximized window should appear
' on the screen. Additionally, if your maximized window isn't going to be full screen,
' you can opt to have it centered on the screen when maximized.

' Every setting is optional & if not used, the class will use Windows default.

' APIs used in this class
Private Declare Function GetWindowPlacement Lib "user32.dll" (ByVal hWnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32.dll" (ByVal hWnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long

Private Const SM_CXMINTRACK As Long = 34
Private Const SM_CYMINTRACK As Long = 35

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Private userMinMaxInfo As MINMAXINFO    ' user-defined window size settings
Private bCentered(0 To 1) As Boolean    ' optional centering of non-full screen maximized windows
Private BorderWidth As POINTAPI         ' used to assist non-full screen Maximized window placement

Private styleAffected As Integer
'0= not in use
'2=MinSize controlled
'4=MinDrag controlled  ' change: always controlled
'8=MaxDrag controlled
Private cMinTrackSize As POINTAPI
Private cMinTrackDefault As POINTAPI

Implements iPropertyBag

Public Sub MinDragSize(bSet As Boolean, Width As Long, Height As Long)
' this is the value to restrict what the smallest window size can be
' Leave at zero and it will be the O/S default. The O/S may prevent
' resizing below absolute minimum values.
With userMinMaxInfo
    If bSet Then
        .ptMinTrackSize.X = Width
        .ptMinTrackSize.Y = Height
        styleAffected = styleAffected Or 4
    Else
        Width = .ptMinTrackSize.X
        Height = .ptMinTrackSize.Y
    End If
End With
End Sub

Public Sub MaxDragSize(bSet As Boolean, Width As Long, Height As Long)
' this is the value to restrict what the largest window size can be
' Leave at zero and routines will calculate it based on system metrics
With userMinMaxInfo
    If bSet Then
        .ptMaxTrackSize.X = Width
        .ptMaxTrackSize.Y = Height
        styleAffected = styleAffected Or 8
    Else
        Width = .ptMaxTrackSize.X
        Height = .ptMaxTrackSize.Y
    End If
End With
End Sub

Public Sub MaximizedSizePos(bSet As Boolean, Left As Long, Top As Long, _
    Width As Long, Height As Long, Optional centerHorizontal As Boolean, _
    Optional centerVertical As Boolean)
    
' Routine allows user to set the window size & position of a maximized window
' These values are interrelated. Routine will not allow you to set the
' position (Left/Top) without setting the size also (width/height).
' However you may set the size without setting the position.
' Think about it for a moment; makes sense.

' Parameters
' Left :: left edge of a maximized window. Use 0 for autocalculating
' Top :: top edge of a maximized window. Use 0 for autocalculating
' Width :: width of maximized window. Use 0 for autocalculating
' Height :: height of maximized window. Use 0 for autocalculating
' centerHorizontal :: only used if Width is non-zero & Top/Left is 0,0
'   will center the preset width of a window horizontally on screen
' centerVertical :: only used if Height is non-zero & Top/Left is 0,0
'   will center the preset height of a window vertically on screen

' Untested for multiMonitor scenarios (virtual desktops)

With userMinMaxInfo
    If bSet Then
        ' don't allow preset window position if window size isn't also preset
        If (Left <> 0 Or Top <> 0) And (Width < 1 Or Height < 1) Then Exit Sub
        .ptMaxSize.X = Width
        .ptMaxSize.Y = Height
        .ptMaxPosition.X = Left
        .ptMaxPosition.Y = Top
        ' center options only available if the maximized window size is preset
        ' and the maximized window position is not preset
        If Top = 0 And Left = 0 Then
            If Width <> 0 Then bCentered(0) = centerHorizontal
            If Height <> 0 Then bCentered(1) = centerVertical
        End If
        If Width = 0 And Height = 0 Then
            styleAffected = styleAffected And Not 8
        Else
            styleAffected = styleAffected Or 8
        End If
    Else
        Left = .ptMaxPosition.X
        Top = .ptMaxPosition.Y
        centerHorizontal = bCentered(0)
        centerVertical = bCentered(1)
        Width = .ptMaxSize.X
        Height = .ptMaxSize.Y
    End If
End With
End Sub

Friend Function SetMinMaxInfo(ByVal hWnd As Long, ByVal xParam As Long, ByVal bIsIconic As Boolean) As Boolean
' Function called from lvSubclasser to overwrite the minmaxinfo structure
' with user-defined settings

If styleAffected = 0 Then Exit Function
' when minimized always use window default settings
If bIsIconic Then Exit Function

Dim tMinMax As MINMAXINFO, tMinDrag As POINTAPI
Dim tPTpos As POINTAPI, tPtsize As POINTAPI

' get copy of window's minmaxinfo <> this is the suggested window size by Windows
CopyMemory tMinMax, ByVal xParam, Len(tMinMax)

    With userMinMaxInfo
    ' We'll be using tMinMax & our userMinMaxInfo structures to merge into one structure
    ' start comparing....
    ' 1. Maximized Size
        tPtsize = tMinMax.ptMaxSize ' cache to aid in centering if needed
        If .ptMaxSize.X > 0 Then tMinMax.ptMaxSize.X = .ptMaxSize.X
        If .ptMaxSize.Y > 0 Then tMinMax.ptMaxSize.Y = .ptMaxSize.Y
    ' 2. Maximized Position -- key property see further down
        tPTpos = tMinMax.ptMaxPosition ' cache to aid in centering if needed
        If .ptMaxPosition.X > 0 Then tMinMax.ptMaxPosition.X = .ptMaxPosition.X
        If .ptMaxPosition.Y > 0 Then tMinMax.ptMaxPosition.Y = .ptMaxPosition.Y
    ' 3. Minimum window size
        tMinDrag = .ptMaxTrackSize
        If (styleAffected Or 16) = styleAffected Then
            If tMinDrag.X < cMinTrackSize.X Then tMinDrag.X = cMinTrackSize.X
            If tMinDrag.Y < cMinTrackSize.Y Then tMinDrag.Y = cMinTrackSize.Y
        End If
        If tMinDrag.X > 0 Then tMinMax.ptMinTrackSize.X = tMinDrag.X
        If tMinDrag.Y > 0 Then tMinMax.ptMinTrackSize.Y = tMinDrag.Y
    ' 4. Maximum window size
        If IsZoomed(hWnd) <> 0 Then ' window is maximized
            ' change the max width/height of window to at least match the maximized size
            tMinMax.ptMaxTrackSize = tMinMax.ptMaxSize
        Else
            If .ptMaxTrackSize.X > 0 Then tMinMax.ptMaxTrackSize.X = .ptMaxTrackSize.X
            If .ptMaxTrackSize.Y > 0 Then tMinMax.ptMaxTrackSize.Y = .ptMaxTrackSize.Y
        End If
    ' 5. Centering as needed, should work on virtual screens (untested) since we are not
    '       using the physical screen to calculate centering, we are using Windows
    '       suggested maximized window position along with its suggested maximizd size
        If bCentered(0) Then
            tMinMax.ptMaxPosition.X = (tPtsize.X - tMinMax.ptMaxSize.X - BorderWidth.X * 2) / 2 + tPTpos.X + BorderWidth.X
            ' windows passed us the upper left corner, don't allow position less than that
            If tMinMax.ptMaxPosition.X < tPTpos.X Then tMinMax.ptMaxPosition.X = tPTpos.X
        Else
            ' if a maximized window preset size and no preset position, window will put in the
            ' top left corner, but the left&top borders are not shown while right&bottom are. Fix that.
            If .ptMaxPosition.X = 0 And .ptMaxSize.X > 0 Then tMinMax.ptMaxPosition.X = tMinMax.ptMaxPosition.X + BorderWidth.X
        End If
        If bCentered(1) Then
            tMinMax.ptMaxPosition.Y = (tPtsize.Y - tMinMax.ptMaxSize.Y - BorderWidth.Y * 2) / 2 + tPTpos.Y + BorderWidth.Y
            ' windows passed us the upper left corner, don't allow position less than that
            If tMinMax.ptMaxPosition.Y < tPTpos.Y Then tMinMax.ptMaxPosition.Y = tPTpos.Y
        Else
            ' if a maximized window preset size and no preset position, window will put in the
            ' top left corner, but the left&top borders are not shown while right&bottom are. Fix that.
            If .ptMaxPosition.Y = 0 And .ptMaxSize.Y > 0 Then tMinMax.ptMaxPosition.Y = tMinMax.ptMaxPosition.Y + BorderWidth.Y
        End If
    End With
    
' ok, done
CopyMemory ByVal xParam, tMinMax, Len(tMinMax)


' Following info applies only if you are setting the X,Y coordinates for a
' maximized window.  Once you change the maximized window position parameters,
' Windows will no longer auto-calculate the best placement for a maximized window.
' Bummer. It will use the setting you passed probably forever. But this is unsuitable
' since it will make me calculate the window dimensions (think virtual screen &
' multi-monitors here). I definitely don't want that hassle. O/S: Just tell me
' where you think a maximized window should go & I'll modify it from there...

' Well, the trick to this problem is to find the cached setting & reset it ourselves.
' The cached window settings are stored in a WindowPlacement structure. What we will
' do is reset that structure's value back to -1,-1 whenever we modify the
' ptMaxPosition manually. This way windows is forced to remeasure.

' So why care if user sets his/her own X,Y coordinates? 'Cause the user can
' change their preferences at a later time & once we ignore this patch &
' subclassing terminated, user is stuck with those settings even if unwanted

If (tPTpos.X <> -1) Or (tPTpos.Y <> -1) Then 'indication we should check

    Dim wP As WINDOWPLACEMENT
    wP.Length = Len(wP)
    GetWindowPlacement hWnd, wP
    ' only change if needed; otherwise we get a bit recursive
    If wP.ptMaxPosition.X <> -1 Or wP.ptMaxPosition.Y <> -1 Then
        wP.ptMaxPosition.X = -1
        wP.ptMaxPosition.Y = -1
        SetWindowPlacement hWnd, wP
    End If
End If

SetMinMaxInfo = True

End Function

Public Sub ResetToSystemDefaults()
    styleAffected = 16
    Dim tMinMax As MINMAXINFO
    userMinMaxInfo = tMinMax
End Sub

Friend Sub SetMinDragSize(ByVal tBAR As Long, ByVal BordersCx As Byte, ByVal BordersCy As Byte, ByVal tbarAlign As eBarAlignment, ByVal minCapLen As Long)

Dim tRect As RECT
Dim tBarBorder As POINTAPI

CopyMemory tRect, ByVal tBAR, &H10
CopyMemory tBarBorder, ByVal minCapLen, &H8

styleAffected = styleAffected Or 16
    Select Case tbarAlign
    Case barAlignLeft, barAlignRight ' vertical
        cMinTrackSize.Y = GetSystemMetrics(SM_CXMINTRACK)
        If cMinTrackSize.Y < BordersCy * 2 + tBarBorder.Y Then
            cMinTrackSize.Y = BordersCy * 2 + tBarBorder.Y
        End If
        cMinTrackSize.X = BordersCx * 2 + tRect.Right - tRect.Left + tBarBorder.X
    Case barAlignBottom, barAlignTop ' horizontal
        cMinTrackSize.X = GetSystemMetrics(SM_CXMINTRACK)
        If cMinTrackSize.X < BordersCx * 2 + tBarBorder.X Then
            cMinTrackSize.X = BordersCx * 2 + tBarBorder.X
        End If
        cMinTrackSize.Y = BordersCy * 2 + tRect.Bottom - tRect.Top + tBarBorder.Y
    End Select
    
    BorderWidth.X = BordersCx
    BorderWidth.Y = BordersCy
    
End Sub

Friend Sub GetMinDragSize(X As Long, Y As Long)
    X = cMinTrackSize.X
    Y = cMinTrackSize.Y
End Sub


'---------------------------------------------------------------------------------------
' Procedure : clsMinMax.Read/WriteProperites
' DateTime  : 9/26/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve custom properites
' Comments  : Used by Export/ImportCustomSettings
'---------------------------------------------------------------------------------------
'
Private Sub iPropertyBag_ReadProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)
    
    Dim sProps() As String
    With userMinMaxInfo
        sProps = Split(propBag.ReadProperty("MM.ptMaxP", "0,0"), ",")
        .ptMaxPosition.X = CLng(sProps(0))
        .ptMaxPosition.Y = CLng(sProps(1))
        sProps = Split(propBag.ReadProperty("MM.ptMaxS", "0,0"), ",")
        .ptMaxSize.X = CLng(sProps(0))
        .ptMaxSize.Y = CLng(sProps(1))
        sProps = Split(propBag.ReadProperty("MM.ptMaxT", "0,0"), ",")
        .ptMaxTrackSize.X = CLng(sProps(0))
        .ptMaxTrackSize.Y = CLng(sProps(1))
        sProps = Split(propBag.ReadProperty("MM.ptMinT", "0,0"), ",")
        .ptMinTrackSize.X = CLng(sProps(0))
        .ptMinTrackSize.Y = CLng(sProps(1))
    End With
    
End Sub

Private Sub iPropertyBag_WriteProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim I As Integer, sProp As String
    
    propBag.WriteProperty "MM.Options", styleAffected, 0
    For I = 0 To UBound(bCentered)
        sProp = sProp & Abs(bCentered(I))
    Next
    propBag.WriteProperty "MM.Center", Mid(sProp, 2), "000"
    
    With userMinMaxInfo
        propBag.WriteProperty "MM.ptMaxP", .ptMaxPosition.X & "," & .ptMaxPosition.Y, "0,0"
        propBag.WriteProperty "MM.ptMaxS", .ptMaxSize.X & "," & .ptMaxSize.Y, "0,0"
        propBag.WriteProperty "MM.ptMaxT", .ptMaxTrackSize.X & "," & .ptMaxTrackSize.Y, "0,0"
        propBag.WriteProperty "MM.ptMinT", .ptMinTrackSize.X & "," & .ptMinTrackSize.Y, "0,0"
    
    End With
    

End Sub
