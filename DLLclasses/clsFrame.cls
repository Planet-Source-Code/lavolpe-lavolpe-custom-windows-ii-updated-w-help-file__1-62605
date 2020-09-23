VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32.dll" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Type TrackingRectangle
    rcItem As RECT
    ID As Long
    State As Byte
    tTip As clsToolTip
End Type

Private cOptions As Byte
'2=allow dragging window on any nonborder inset
'4=inset used


Private cColors(0 To 1) As clsBorders      ' active/inactive window colors
Private cMinMax As clsMinMax                ' GetMinMaxInfo class
Private cSysTray As clsSysTrays             ' collection of sys tray icon(s)

Private cTrackers() As TrackingRectangle    ' collection of user-defined NC tracking rectangles
Private cTrackID As Long                    ' currently hovered tracking rectangle
Private cLastHit As Long                    ' last hit test of tracking rectangles
Private cActive As Byte                     ' cached window state

Private cInset() As Long                    ' optional extra NC inset borders
Private cBorderWidthCx As Byte              ' optional border size
Private cBorderWidthCy As Byte              ' default is 4
Private mNotifyUser As Long                 ' owner drawn flags
Private mMainImplementation As Long         ' primary class pointer
Private mODImplementation As Long           ' owner drawn Implementation pointer
Private pHwnd As Long                       ' cached hWnd
Private cAutoRedraw As Boolean              ' flag used to prevent redrawing during multiple property settings

Implements iTimer                           ' timer call back
Implements iPropertyBag

'---------------------------------------------------------------------------------------
' Procedure : clsFrame.ResetToSystemDefaults
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Resets colors to system defaults
'---------------------------------------------------------------------------------------
'
Public Sub ResetToSystemDefaults(ByVal Options As eBarResets)
    If (Options And rstBorderColors) = rstBorderColors Or Options = rstAll Then
        Me.Active.ResetToSystemDefaults Options
        Me.Inactive.ResetToSystemDefaults Options
    End If
    If (Options And rstBorderSizes) = rstBorderSizes Or Options = rstAll Then
        cBorderWidthCx = 4
        cBorderWidthCy = 4
        Erase cInset()
        cOptions = cOptions And Not 4
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.Active/Inactive
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Return color classes for active/inactive windows
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Get Active() As clsBorders
    If cColors(1) Is Nothing Then Set cColors(1) = New clsBorders
    Set Active = cColors(1)
End Property
Public Property Get Inactive() As clsBorders
    If cColors(0) Is Nothing Then Set cColors(0) = New clsBorders
    Set Inactive = cColors(0)
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.Get/SetInsetOffset
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Allows user to add additional space between the menubar/titlebar/border
'             edges and the client rectangle. Useful to custom patterns on the NC
'             or placing additional text (i.e., URL on the NC)
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub GetInsetOffset(BufferLeft As Byte, BufferTop As Byte, BufferRight As Byte, BufferBottom As Byte)
    If (cOptions Or 4) = cOptions Then
        BufferBottom = cInset(0)
        BufferTop = cInset(1)
        BufferRight = cInset(2)
        BufferLeft = cInset(3)
    End If
End Sub
Public Sub SetInsetOffset(ByVal BufferLeft As Byte, ByVal BufferTop As Byte, _
                        ByVal BufferRight As Byte, ByVal BufferBottom As Byte)
    
    ' all zeros will reset the inset to default
    ' The inset Offset cannot reduce the minimum 1 pixel border that may be applied
    ' when measuring the titlebar and menubar. This 1 pixel border will not exist between
    ' the client rectangle and the border edges if they are adjacent to each other
    
    ReDim cInset(0 To 3)
    cInset(0) = BufferLeft
    cInset(1) = BufferTop
    cInset(2) = BufferRight
    cInset(3) = BufferBottom
    If cInset(0) + cInset(1) + cInset(2) + cInset(3) = 0 Then
        Erase cInset
        cOptions = cOptions And Not 4
    Else
        cOptions = cOptions Or 4
    End If
    RefreshWindow
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.BorderWidth/BorderHeight Properties
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Sets or retrieves the border sizes for the window
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Get BorderWidth() As Byte
    BorderWidth = cBorderWidthCx
End Property
Public Property Let BorderWidth(Size As Byte)

    cBorderWidthCx = Size
    RefreshWindow
    
End Property
Public Property Get BorderHeight() As Byte
    BorderHeight = cBorderWidthCy
End Property
Public Property Let BorderHeight(Size As Byte)

    cBorderWidthCy = Size
    RefreshWindow

End Property
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.DrawFrame
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Called internally to fill the inset and draw borders
'---------------------------------------------------------------------------------------
'
Friend Function DrawFrame(ByVal winRectPtr As Long, ByVal clientRectPtr As Long, _
                ByVal currentInset As Long, ByVal isActive As Byte) As Long

    Dim cRect As RECT, wRect As RECT, ncRect As RECT
    Dim cClient As iOwnerDrawn, BKG As BkgAction_LV
    Dim bDefaultBkg As Boolean, bGrayScaled As Boolean
    Dim bkgImage As StdPicture, bkgStyle As eBackStyles
    Dim insetColor As Long, isTransparent As Boolean
    Dim hDC As Long, hBrush As Long
    Dim borderEdge(0 To 1) As Byte
    
    ' the nonClient inset coordinates/dimensions are not tracked. Need to calculate
    CopyMemory wRect, ByVal winRectPtr, &H10
    CopyMemory ncRect, ByVal clientRectPtr, &H10
        
    cActive = isActive
    
    OffsetRect ncRect, -wRect.Left, -wRect.Top
    OffsetRect wRect, -wRect.Left, -wRect.Top
    BKG.rcItem = wRect
    
    ' next line is done w/o clipping regions so I can tell if any of my calculations
    ' are off.  If they are off, I'll draw over the border(s)
    DrawBorders wRect       ' draw our outside borders first
    hDC = Canvas(True)
    cAutoRedraw = True
    
    ' calculate the non-client inset
    If (currentInset Or 1) = currentInset Then ncRect.Left = ncRect.Left - 1
    If (currentInset Or 2) = currentInset Then ncRect.Top = ncRect.Top - 1
    If (currentInset Or 4) = currentInset Then ncRect.Right = ncRect.Right + 1
    If (currentInset Or 8) = currentInset Then ncRect.Bottom = ncRect.Bottom + 1

    If (cOptions Or 4) = cOptions Then
        ' additional nc inset is used
        ncRect.Left = ncRect.Left - cInset(0)
        ncRect.Top = ncRect.Top - cInset(1)
        ncRect.Right = ncRect.Right + cInset(2)
        ncRect.Bottom = ncRect.Bottom + cInset(3)
        ' now adjust the inset if our window is being resized really small
        ValidateBorders borderEdge(0), borderEdge(1)
        If ncRect.Bottom > BKG.rcItem.Bottom - borderEdge(1) Then ncRect.Bottom = BKG.rcItem.Bottom - borderEdge(1)
        If ncRect.Right > BKG.rcItem.Right - borderEdge(0) Then ncRect.Right = BKG.rcItem.Right - borderEdge(0)
    End If
    
    ' see if the inset uses an image
    cColors(isActive).GetInsetImage bkgImage, bkgStyle, bGrayScaled
    
    ' set flag to draw background colors
    If bkgImage Is Nothing Then
        bDefaultBkg = True
    Else
        If (cOptions And 4) = 4 Then bDefaultBkg = True
        ' otherwise the bkg image would fill the inset; so don't color bkg
    End If
    
    ' color the inset background as needed
    If bDefaultBkg Then
        insetColor = cColors(isActive).GetInsetBackColor(isTransparent)
        If Not isTransparent Then
            hBrush = CreateSolidBrush(insetColor)
            FillRect hDC, ncRect, hBrush
            DeleteObject hBrush
        End If
    End If
    
    ' fill the bkg with the image as needed
    If Not bkgImage Is Nothing Then
        FillBarImage bkgImage, hDC, ncRect.Left, ncRect.Top, ncRect.Right - ncRect.Left, ncRect.Bottom - ncRect.Top, bkgStyle, False, bGrayScaled
        Set bkgImage = Nothing
    End If

    ' see if the user is requesting drawing borders.
    ' If so, the user has the capability of overdrawing everything we just drew
    If mODImplementation <> 0 Then
            
        If (mNotifyUser And odFrameBorders) = odFrameBorders Then

            ' prepare the UDT and pass it to the user
            On Error Resume Next ' always, to try & prevent user from crashing
            SetRect BKG.rcExtra, cBorderWidthCx, cBorderWidthCy, BKG.rcItem.Right - cBorderWidthCx, BKG.rcItem.Bottom - cBorderWidthCy
            BKG.hDC = hDC
            
            GetObjectFromPointer mODImplementation, cClient
            cClient.OwnerDrawMessage omDrawPreNC, isActive + 0, VarPtr(BKG), ""
            Set cClient = Nothing
            
            SelectClipRgn hDC, ByVal 0& ' always remove any clipping region client may have left behing
        
        End If
        
        ' return pointer to user's implementation so that a post NC message
        ' can also be sent after the menubar & titlebar are drawn
        If (mNotifyUser And odPostNCDrawing) = odPostNCDrawing Then DrawFrame = mODImplementation
        
    End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.DrawBorders
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Called internally to draw the standard or custom border edges
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub DrawBorders(winRect As RECT)

    Dim cRect As RECT, clipRgn As Long
    Dim hDC As Long, hPenBrush As Long
    Dim lColors(0 To 3) As Long
    Dim lBorders(0 To 1) As Byte
    
    Const BDR_RAISEDINNER As Long = &H4
    Const BDR_RAISEDOUTER As Long = &H1
    Const BF_LEFT As Long = &H1
    Const BF_TOP As Long = &H2
    Const BF_RIGHT As Long = &H4
    Const BF_BOTTOM As Long = &H8
    
    hDC = Canvas(True)
    hPenBrush = CreateSolidBrush(cColors(cActive).BackColor)
    FillRect hDC, winRect, hPenBrush
    DeleteObject hPenBrush
        
    ' by default the DLL will draw a 3pixel border for non-standard border sizes
    ValidateBorders lBorders(0), lBorders(1)
    If lBorders(0) > 4 Then lBorders(0) = 3
    If lBorders(1) > 4 Then lBorders(1) = 3
        
    If cColors(cActive).pIsCustomBorder Then  ' custom colors used
    
        ' get custom defined border colors
        cColors(cActive).GetBorderColors lColors(0), lColors(1), lColors(2), lColors(3)
        
        If lBorders(0) = 1 Then
            ' draw a single border
            hPenBrush = CreateSolidBrush(lColors(3))
        Else
            ' draw the 1st 3 pixels of the border
            hPenBrush = SelectObject(hDC, CreatePen(0, 1, lColors(0)))
            MoveToEx hDC, 0, winRect.Bottom - 2, ByVal 0&
            LineTo hDC, 0, 0
            LineTo hDC, winRect.Right - 1, 0
            DeleteObject SelectObject(hDC, CreatePen(0, 1, lColors(3)))
            LineTo hDC, winRect.Right - 1, winRect.Bottom - 1
            LineTo hDC, -1, winRect.Bottom - 1
            MoveToEx hDC, 1, winRect.Bottom - 2, ByVal 0
            DeleteObject SelectObject(hDC, CreatePen(0, 1, lColors(1)))
            LineTo hDC, 1, 1
            LineTo hDC, winRect.Right - 2, 1
            DeleteObject SelectObject(hDC, CreatePen(0, 1, lColors(2)))
            LineTo hDC, winRect.Right - 2, winRect.Bottom - 2
            LineTo hDC, 1, winRect.Bottom - 2
            DeleteObject SelectObject(hDC, hPenBrush)
            hPenBrush = CreateSolidBrush(lColors(0))
            InflateRect winRect, -2, -2
        End If
    Else
        ' use system settings
        If lBorders(0) = 1 Then
            hPenBrush = CreateSolidBrush(0)
        Else
            DrawEdge hDC, winRect, BDR_RAISEDINNER Or BDR_RAISEDOUTER, BF_BOTTOM Or BF_LEFT Or BF_RIGHT Or BF_TOP
            InflateRect winRect, -2, -2
            hPenBrush = CreateSolidBrush(ConvertVBSysColor(vbButtonFace))
        End If
    End If
    
    ' finish by filling in the 1 pixel inner rectangle
    FrameRect hDC, winRect, hPenBrush
    If lBorders(0) = 4 Then
        InflateRect winRect, -1, -1
        FrameRect hDC, winRect, hPenBrush
    End If
    DeleteObject hPenBrush
            
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.HitTest
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : the hit test for the borders, processed mainly by WM_NCHITTEST message
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function HitTest(ByVal x As Long, ByVal Y As Long, _
                    ByVal winRectPtr As Long, ByVal clientRectPtr As Long)
    
    ' Hit test functionality for the borders
    ' Called for the WM_NCHITTEST message
    
    Dim lRtn As Long
    Dim wRect As RECT, cRect As RECT
    Dim bdrCx As Byte, bdrCy As Byte
    
    CopyMemory wRect, ByVal winRectPtr, &H10
    
    ' ensure point in our window
    If PtInRect(wRect, x, Y) = 0 Then Exit Function
    
    ' quick test vs doing all the IFs below. Is cursor within inside edges of our borders?
    ValidateBorders bdrCx, bdrCy
    If bdrCx > 4 Then bdrCx = 4
    If bdrCy > 4 Then bdrCy = 4
    
    SetRect cRect, wRect.Left + bdrCx, wRect.Top + bdrCy, wRect.Right - bdrCx - 1, wRect.Bottom - bdrCy - 1
    If PtInRect(cRect, x, Y) Then
        
        ' not a border hit, but could it be a custom inset or the DLL defined 1 pixel border
        ' btwn the menubar and client rect or btwn the titlebar and client rect?
        ' We could also be tracking stuff on the non-client area for our user.
        ' Don't want to return NULL if the hit test is inside our window
    
        cLastHit = TrackerHitTest(x - wRect.Left, Y - wRect.Top)
        If cLastHit > 0 Then ' we have a hit on a tracked item
            If cTrackID > 0 Then
                If cLastHit <> cTrackID Then  'otherwise mouse still in tracker rectangle
                    ' return the mouse leave event for the user
                    clearTracker
                End If
            End If
            lRtn = HTNC_Tracker
        Else
            lRtn = HTNC_Custom
            If (cOptions And 2) = 2 Then 'expanded dragging is allowed
                SetRect cRect, wRect.Left + cBorderWidthCx, wRect.Top + cBorderWidthCy, wRect.Right - cBorderWidthCx, wRect.Bottom - cBorderWidthCy
                If PtInRect(cRect, x, Y) = 0 Then
                    lRtn = HTNC_Caption
                End If
            End If
        End If
        
    Else
    
        If cTrackID > 0 Then ' was tracking something
            ' return the mouse leave event for the user
            clearTracker
        End If
    
        With wRect
    
            ' check for bottom border
            If Y < .Bottom + 1 And Y > .Bottom - bdrCy - 1 Then
                ' check for a right & left hit
                If x < .Right + 1 And x > .Right - bdrCx - HTTOP Then
                    lRtn = HTBOTTOMRIGHT
                ElseIf x > .Left - 1 And x < .Left + bdrCx + HTTOP Then
                    lRtn = HTBOTTOMLEFT
                Else
                    lRtn = HTBOTTOM
                End If
    
                ' check for right border
            ElseIf x > .Right - bdrCx - 1 And x < .Right + 1 Then
                ' check for bottom & top borders
                If Y < .Bottom + 1 And Y > .Bottom - bdrCy - HTTOP Then
                    lRtn = HTBOTTOMRIGHT
                ElseIf Y > .Top - 1 And Y < .Top + bdrCy + HTTOP Then
                    lRtn = HTTOPRIGHT
                Else
                    lRtn = HTRIGHT
                End If
    
                ' check for left border
            ElseIf x > .Left - 1 And x < .Left + bdrCx + 1 Then
                ' check for bottom & top borders
                If Y < .Bottom + 1 And Y > .Bottom - bdrCy - HTTOP Then
                    lRtn = HTBOTTOMLEFT
                ElseIf Y > .Top - 1 And Y < .Top + bdrCy + HTTOP Then
                    lRtn = HTTOPLEFT
                Else
                    lRtn = HTLEFT
                End If
    
                ' check for top border
            ElseIf Y > .Top - 1 And Y < .Top + bdrCy + 1 Then
                ' check for a right & left hit
                If x < .Right + 1 And x > .Right - bdrCx - HTTOP Then
                    lRtn = HTTOPRIGHT
                ElseIf x > .Left - 1 And x < .Left + bdrCx + HTTOP Then
                    lRtn = HTTOPLEFT
                Else
                    lRtn = HTTOP
                End If
            End If
        End With
    
    End If
    
    HitTest = lRtn
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.BorderWidths
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Returns pointers to the standard border widths & any custom inset used
'---------------------------------------------------------------------------------------
'
Friend Sub BorderWidths(wDefaultCx As Byte, wDefaultCy As Byte, wCustom As Long)
    ValidateBorders wDefaultCx, wDefaultCy
    If (cOptions Or 4) = cOptions Then wCustom = VarPtr(cInset(0)) Else wCustom = 0
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.sharedProperty
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : A way to get other class properties without having to
'             worry about caching them here & keeping them updated
'---------------------------------------------------------------------------------------
'
Friend Property Let pMainImplementation(hWnd As Long, lImpObj As Long)
    pHwnd = hWnd
    mMainImplementation = lImpObj
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.OwnerDrawn
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Passed owner drawn flags
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Sub OwnerDrawn(cImplementation As Long, odItems As eOwnerDrawn, isClosing As Boolean)
    
    If isClosing Then
        mNotifyUser = 0
        mODImplementation = 0
    Else
        mNotifyUser = mNotifyUser And 2048  ' carry over tracking rectangle ownerdrawn flag
        If (odItems And odPostNCDrawing) = odPostNCDrawing Then mNotifyUser = mNotifyUser Or odPostNCDrawing
        If (odItems And odFrameBorders) = odFrameBorders Then mNotifyUser = mNotifyUser Or odFrameBorders
        If mNotifyUser = 0 Then
            mODImplementation = 0
            CreateTimer 0
        Else
            mODImplementation = cImplementation
        End If
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.MinMax
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Expose the GetMinMaxInfo class
'---------------------------------------------------------------------------------------
'
Public Property Get MinMax() As clsMinMax
    If cMinMax Is Nothing Then Set cMinMax = New clsMinMax
    Set MinMax = cMinMax
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.SysTray
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Exposes the System Tray class
'---------------------------------------------------------------------------------------
'
Public Function SystemTray(ByVal uniqueID As Long, iImplementation As iSysTrayCallback) As clsSysTray
    
' Unique ID: Application-defined identifier of the taskbar icon.
' Values from 0 to 12 are reserved and should not be used
    On Error Resume Next
    If Not iImplementation Is Nothing Then
        If cSysTray Is Nothing Then
            Dim cClient As iImpProps
            GetObjectFromPointer mMainImplementation, cClient
            cClient.SetPropValue [_sysTrayImp], ObjPtr(iImplementation)
            Set cClient = Nothing
            Set cSysTray = New clsSysTrays
            If WM_IECrashNotify = 0 Then WM_IECrashNotify = RegisterWindowMessage("TaskbarCreated")
        End If
        Set SystemTray = cSysTray.TrayIcon(pHwnd, uniqueID)
    End If
End Function
Public Sub RemoveSystemTray(ByVal uniqueID As Long, Optional RemoveAll As Boolean, Optional DeleteIcon As Boolean)
' Unique ID: Application-defined identifier of the taskbar icon.
' Values from 0 to 12 are reserved and should not be used
    If Not cSysTray Is Nothing Then
        cSysTray.RemoveTrayIcon uniqueID, RemoveAll, DeleteIcon
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.RestoreSystemTray
' DateTime  : 9/24/2005
' Author    : LaVolpe
' Purpose   : Restores tray icons after an Explorer crash
'---------------------------------------------------------------------------------------
'
Friend Sub RestoreSystemTray()
    If Not cSysTray Is Nothing Then cSysTray.RestoreSysTray
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.AllowExtendedMove
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : With non-standard borders, this option allows user to drag the window
'             within the extended borders and also display the system menu
'             The user cannot double click to maximize, that still must be done
'             on the titlebar proper, by design
'---------------------------------------------------------------------------------------
'
Public Property Let AllowExtendedMove(bAllow As Boolean)
    If bAllow Then
        cOptions = cOptions Or 2
    Else
        cOptions = cOptions And Not 2
    End If
End Property
Public Property Get AllowExtendedMove() As Boolean
    AllowExtendedMove = ((cOptions And 2) = 2)
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.Add/RemoveTrackingRect
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Allows user to add/remove a callback function for tracking customized NC space
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function AddTrackingRect(ByVal ID As Long, ByVal Left As Long, ByVal Top As Long, _
                                ByVal Right As Long, ByVal Bottom As Long, _
                                iImplementation As iOwnerDrawn, _
                                Optional ToolTip As String = "nc") As Boolean

    ' Tracking rectangles allow user to identify space on the nonclient area
    ' and automatically requests some subclassed message traffic to include
    ' the following messages. This allows the user to modify that NC space
    ' based on user mouse actions
    
    ' Messages trapped and interperted for user are:
    '        WM_SETCURSOR
    '        WM_NCMOUSEMOVE
    '        WM_NCLBUTTONUP , WM_NCLBUTTONDOWN
    '        WM_NCRBUTTONUP , WM_NCRBUTTONDOWN
    '   These messages are sent as
    '       omTrackClick, omTrackCursor, omDrawTrackRect messages to your
    '       OwnerDrawMessage implementation


    If iImplementation Is Nothing Then Exit Function
    If pHwnd = 0 Then Exit Function
    
    Dim T As Long, tRect As RECT, wRect As RECT, cRect As RECT
    
    SetRect tRect, Left, Top, Right, Bottom
    If IsRectEmpty(tRect) Then Exit Function
    
    CopyMemory cRect, ByVal sharedProperty([_clientRect]), &H10
    CopyMemory wRect, ByVal sharedProperty([_winRect]), &H10
    ' convert the client's screen coordinates to window coordinates
    OffsetRect cRect, -wRect.Left, -wRect.Top
    
    mODImplementation = ObjPtr(iImplementation)
    mNotifyUser = mNotifyUser Or 2048   ' flag for ownerdrawing tracking rectangles
    
    If IsArrayEmpty(Not cTrackers) Then
        ReDim cTrackers(0)  ' add new tracker
    Else
        For T = 0 To UBound(cTrackers)
            ' find the tracker if it exists
            If cTrackers(T).ID = ID Then
                ' it exists, simply update it
                With cTrackers(T)
                    .rcItem = tRect
                    If Len(ToolTip) = 0 Then
                        If Not .tTip Is Nothing Then Set .tTip = Nothing
                    ElseIf ToolTip <> "nc" Then
                        OffsetRect tRect, -tRect.Left + (tRect.Left - cRect.Left), -tRect.Top + (tRect.Top - cRect.Top)
                        If .tTip Is Nothing Then
                            Set .tTip = New clsToolTip
                            .tTip.CreateToolTip pHwnd, ToolTip, VarPtr(tRect)
                        Else
                            .tTip.TipText = ToolTip
                            .tTip.TipRect = VarPtr(tRect)
                        End If
                    End If
                End With
                AddTrackingRect = True
                Exit Function
            End If
        Next
        ' increase tracker array
        ReDim Preserve cTrackers(0 To UBound(cTrackers) + 1)
    End If
    ' add the tracker elements
    With cTrackers(UBound(cTrackers))
        .rcItem = tRect
        .ID = ID
        If Len(ToolTip) > 0 Then
            If ToolTip <> "nc" Then
                OffsetRect tRect, -tRect.Left + (tRect.Left - cRect.Left), -tRect.Top + (tRect.Top - cRect.Top)
                Set .tTip = New clsToolTip
                .tTip.CreateToolTip pHwnd, ToolTip, VarPtr(tRect)
            End If
        End If
    End With
    AddTrackingRect = True
    
End Function
Public Function RemoveTrackingRect(ByVal ID As Long, Optional RemoveAll As Boolean) As Boolean
    
    ' ^^ see AddTrackingRect for more information
    
    If IsArrayEmpty(Not cTrackers) Then Exit Function

    If RemoveAll Then
        clearTracker
        Erase cTrackers
    Else
        Dim T As Long
        For T = 0 To UBound(cTrackers)
            ' find the ID to remove
            If cTrackers(T).ID = ID Then Exit For
        Next
        If T > UBound(cTrackers) Then Exit Function
        
        clearTracker
        ' remove it & resize the array as needed
        If UBound(cTrackers) = 0 Then
            Erase cTrackers
        Else
            For T = T To UBound(cTrackers) - 1
                cTrackers(T) = cTrackers(T + 1)
            Next
            ReDim Preserve cTrackers(0 To UBound(cTrackers) - 1)
        End If
    End If
    If IsArrayEmpty(Not cTrackers) Then
        mNotifyUser = mNotifyUser And Not 2048 ' remove flag for ownerdrawn trackers
        If mNotifyUser = 0 Then mODImplementation = 0
    End If
    RemoveTrackingRect = True
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.GetTrackingRect
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Retrieves a tracking rectangle for the user
' Comments  : Not called by the DLL anywhere
'---------------------------------------------------------------------------------------
'
Public Function GetTrackingRect(ByVal ID As Long, Left As Long, Top As Long, _
                                    Right As Long, Bottom As Long) As Boolean

    If IsArrayEmpty(Not cTrackers) Then Exit Function

    Dim T As Long
    For T = 0 To UBound(cTrackers)
        If cTrackers(T).ID = ID Then Exit For
    Next
    If T > UBound(cTrackers) Then Exit Function
        
    With cTrackers(T).rcItem
        Left = .Left
        Right = .Right
        Top = .Top
        Bottom = .Bottom
    End With
    
    GetTrackingRect = True

End Function
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.TrackerHitTest
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Part of the HitTest function & returns non-zero if mouse is over a tracker
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function TrackerHitTest(ByVal x As Long, ByVal Y As Long) As Long
    
        Dim T As Long
        If mODImplementation = 0 Then Exit Function
        If IsArrayEmpty(Not cTrackers) Then
            T = 0
        Else
            For T = UBound(cTrackers) To 0 Step -1
                If PtInRect(cTrackers(T).rcItem, x, Y) Then Exit For
            Next
            T = T + 1
        End If
        TrackerHitTest = T
        
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.clearTracker
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Clears any timer and posts a redraw message to user when trackers exist
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub clearTracker()
    
    If cTrackID > 0 Then
        
        Dim cClient As iOwnerDrawn, CID As CustomItemDraw_LV
        Dim tID As Long
        
        tID = cTrackID - 1
        CreateTimer 0
        
        On Error Resume Next ' always, to try & prevent user from crashing
        GetObjectFromPointer mODImplementation, cClient
        With CID
            .hDC = Canvas(True)
            .itemID = cTrackers(tID).ID
            .itemState = mcStandard
            .rcItem = cTrackers(tID).rcItem
        End With
        SendTrackerInfo CID
        cTrackers(tID).State = 0
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.SetMouseAction
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Rerouted messages from the primary class for trackers only
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function SetMouseAction(ByVal wMsg As Long, ByVal isActive As Byte) As Boolean
    
    If cLastHit = 0 Then Exit Function
    If mODImplementation = 0 Then Exit Function
    cActive = isActive
    
    Dim cClient As iOwnerDrawn, CID As CustomItemDraw_LV
    Dim Overridden As Long
    
    On Error Resume Next ' always, to try & prevent user from crashing
    Select Case wMsg
        Case WM_SETCURSOR
            If cLastHit > 0 Then
                GetObjectFromPointer mODImplementation, cClient
                cClient.OwnerDrawMessage omTrackCursor, cTrackers(cLastHit - 1).ID, VarPtr(Overridden), ""
                Set cClient = Nothing
                SetMouseAction = (Overridden <> 0)
            End If
            
        Case WM_NCMOUSEMOVE
            If (cTrackers(cLastHit - 1).State And mcHover) = 0 Then
                ' mouse hasn't been over this
                clearTracker
                cTrackers(cLastHit - 1).State = cTrackers(cLastHit - 1).State Or mcHover
                With CID
                    .hDC = Canvas(True)
                    .itemID = cTrackers(cLastHit - 1).ID
                    .itemState = mcHover
                    .rcItem = cTrackers(cLastHit - 1).rcItem
                End With
                SendTrackerInfo CID
                CreateTimer cLastHit
            End If
            SetMouseAction = True
            
        Case WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN
            
            With CID
                .hDC = Canvas(True)
                .itemID = cTrackers(cLastHit - 1).ID
                .itemState = mcSelect
                If wMsg = WM_NCRBUTTONDOWN Then .itemState = .itemState Or mcRightButton
                .rcItem = cTrackers(cLastHit - 1).rcItem
            End With
            
            GetObjectFromPointer mODImplementation, cClient
            cClient.OwnerDrawMessage omDrawTrackRect, cActive + 0, VarPtr(CID), ""
            Set cClient = Nothing
            
            cTrackers(cLastHit - 1).State = mcSelect Or mcHover ' down, no click & mouse over
            SetMouseAction = True
            
        Case WM_NCRBUTTONUP, WM_NCLBUTTONUP
            
            If (cTrackers(cLastHit - 1).State And mcSelect) = mcSelect Then
                With CID
                    .hDC = Canvas(True)
                    .itemID = cTrackers(cLastHit - 1).ID
                    .itemState = mcStandard
                    .rcItem = cTrackers(cLastHit - 1).rcItem
                End With
            
                clearTracker
                GetObjectFromPointer mODImplementation, cClient
                If wMsg = WM_NCRBUTTONUP Then
                    cClient.OwnerDrawMessage omTrackClick, cTrackers(cLastHit - 1).ID, cActive Or mcRightButton, ""
                Else
                    cClient.OwnerDrawMessage omTrackClick, cTrackers(cLastHit - 1).ID, cActive + 0, ""
                End If
                Set cClient = Nothing
            End If
            clearTracker
            SetMouseAction = True
        
        Case Else
    
    End Select
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.SendTrackerInfo
' DateTime  : 9/20/2005
' Author    : LaVolpe
' Purpose   : Copies exsiting tracker image to offscreen bitmap so user can update if needed
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub SendTrackerInfo(CID As CustomItemDraw_LV)
    
    Dim cClient As iOwnerDrawn
    Dim wDC As Long
    If mODImplementation <> 0 Then
        wDC = GetWindowDC(pHwnd)
        With CID
            BitBlt .hDC, .rcItem.Left, .rcItem.Top, .rcItem.Right - .rcItem.Left + 1, _
                .rcItem.Bottom - .rcItem.Top + 1, wDC, .rcItem.Left, .rcItem.Top, vbSrcCopy
        End With
        ReleaseDC pHwnd, wDC
        On Error Resume Next
        GetObjectFromPointer mODImplementation, cClient
        cClient.OwnerDrawMessage omDrawTrackRect, cActive + 0, VarPtr(CID), ""
        Set cClient = Nothing
        
        'user must modify the .itemOD element in order to post any changes
        If CID.itemOD <> 0 Then
            wDC = GetWindowDC(pHwnd)
            With CID
                SelectClipRgn .hDC, ByVal 0&
                BitBlt wDC, .rcItem.Left, .rcItem.Top, .rcItem.Right - .rcItem.Left + 1, _
                    .rcItem.Bottom - .rcItem.Top + 1, .hDC, .rcItem.Left, .rcItem.Top, vbSrcCopy
            End With
            ReleaseDC pHwnd, wDC
        End If
    End If
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.ValidateBorders
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Helper function that returns the preferred border size for drawing a border
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Sub ValidateBorders(Cx As Byte, Cy As Byte)

    If mODImplementation = 0 Then
        If cBorderWidthCx = cBorderWidthCy Then
            Select Case cBorderWidthCx
            Case 2
                Cx = 3
                Cy = 3
            Case 1, 3, 4
                Cx = cBorderWidthCx
                Cy = cBorderWidthCy
            Case Else
                Cx = 4
                Cy = 4
            End Select
        Else
            Cx = 4
            Cy = 4
        End If
    Else
        Cx = cBorderWidthCx
        Cy = cBorderWidthCy
    End If
                    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.CreateTimer
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : Establishes a timer that firest 8x a second
'---------------------------------------------------------------------------------------
'
Private Sub CreateTimer(tID As Long)

    If cTrackID Then
        If tID = 0 Then  ' killing the timer
            SetStopTimer False, pHwnd, ObjPtr(Me), 0
            cTrackID = tID
        End If
    Else
        If tID <> 0 Then      ' creating the timer
            SetStopTimer True, pHwnd, ObjPtr(Me), 150
            cTrackID = tID
        End If
    End If
        
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.GetInset
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Returns the inset dimensions to the primary class
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function GetInset(ByVal clientRectPtr As Long, currentInset As Byte) As Long

    Dim ncRect As RECT
    
    CopyMemory ncRect, ByVal clientRectPtr, &H10
    
    If (currentInset Or 1) = currentInset Then ncRect.Left = ncRect.Left - 1
    If (currentInset Or 2) = currentInset Then ncRect.Top = ncRect.Top - 1
    If (currentInset Or 4) = currentInset Then ncRect.Right = ncRect.Right + 1
    If (currentInset Or 8) = currentInset Then ncRect.Bottom = ncRect.Bottom + 1

    If (cOptions Or 4) = cOptions Then
        ncRect.Left = ncRect.Left - cInset(0)
        ncRect.Top = ncRect.Top - cInset(1)
        ncRect.Right = ncRect.Right + cInset(2)
        ncRect.Bottom = ncRect.Bottom + cInset(3)
    End If

    GetInset = VarPtr(ncRect)

End Function

'---------------------------------------------------------------------------------------
' Procedure : clsFrame.Read/WriteProperites
' DateTime  : 9/26/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve custom properites
' Comments  : Used by Export/ImportCustomSettings
'---------------------------------------------------------------------------------------
'
Private Sub iPropertyBag_ReadProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim I As Integer, J As Integer
    Dim sProp As String, sProps() As String
    Dim iBag As iPropertyBag
    
    cOptions = propBag.ReadProperty("FM.Options", 0)
    mNotifyUser = propBag.ReadProperty("FM.OD", 0)
    
    sProp = propBag.ReadProperty("FM.Inset", "")
    If sProp = "" Then
        Erase cInset
    Else
        ReDim cInset(0 To 3)
        sProps = Split(sProp, ",")
        For I = 0 To 3
            cInset(I) = CLng(sProps(I))
        Next
    End If
    sProps = Split(propBag.ReadProperty("FM.Borders", "4,4"), ",")
    cBorderWidthCx = CByte(sProps(0))
    cBorderWidthCy = CByte(sProps(1))
    
    If (Limits And extTrackRects) = extTrackRects Then
        I = propBag.ReadProperty("FM.Trackers", 0)
        If I = 0 Then
            Erase cTrackers
        Else
            ReDim cTrackers(0 To I - 1)
            For I = 0 To UBound(cTrackers)
                sProps = Split(propBag.ReadProperty("FM.Trk" & I), ",")
                cTrackers(I).ID = CLng(sProps(0))
                SetRect cTrackers(I).rcItem, CLng(sProps(1)), CLng(sProps(2)), CLng(sProps(3)), CLng(sProps(4))
                sProp = propBag.ReadProperty("FM.TrkTip" & I, "")
                If Len(sProp) Then
                    Set cTrackers(I).tTip = New clsToolTip
                    cTrackers(I).tTip.CreateToolTip pHwnd, sProp, VarPtr(cTrackers(I).rcItem)
                End If
            Next
        End If
    Else
        mNotifyUser = mNotifyUser And Not 2048
    End If
    
    mODImplementation = lParam
    
    Dim sharedImgPtr As Long, iPic As StdPicture
    Dim iGrayScaled As Boolean, iStyle As eBackStyles
    '^^ above are used when trying to share images between active/inactive classes
    '   vs creating separate instances of the same image
    
    Set iBag = Me.Active
    ' if the active class has an image, the sharedImgPtr will be the value
    ' of the image handle at the time the property bag was saved
    iBag.ReadProperties propBag, Limits, sharedImgPtr, Index
    Set iBag = Me.Inactive
    ' The inactive class will also compare it's cached value of its image when
    ' the property bag was saved. If the active & inactive class's image values
    ' are different, the sharedImgPtr below will be zero
    iBag.ReadProperties propBag, Limits, sharedImgPtr, Index
    If sharedImgPtr <> 0 Then
        ' The active and inactive class will share the same image.
        ' Get the inactive class's current image properties
        Me.Inactive.GetInsetImage Nothing, iStyle, iGrayScaled
        ' Get the active class's image
        Me.Active.GetInsetImage iPic
        ' now reassign the inactive class's image properties & image
        Me.Inactive.SetInsetImage iPic, iStyle, iGrayScaled
    End If
    Set iBag = Me.MinMax
    iBag.ReadProperties propBag, Limits, lParam, Index
    
    If (Limits And exSysTray) = exSysTray Then
        If propBag.ReadProperty("ST.Count", 0) > 0 Then
            Set cSysTray = New clsSysTrays
            If Index <> 0 Then
                Set iBag = cSysTray
                iBag.ReadProperties propBag, Limits, pHwnd, 0
                If WM_IECrashNotify = 0 Then WM_IECrashNotify = RegisterWindowMessage("TaskbarCreated")
            End If
        End If
    End If
    Set iBag = Nothing

End Sub

Private Sub iPropertyBag_WriteProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim I As Integer, sProp As String
    Dim iBag As iPropertyBag
    Dim sharedImgPtr As Long
    
    propBag.WriteProperty "FM.Options", cOptions, 0
    
    If (Limits And extTrackRects) = extTrackRects Then
        If Not IsArrayEmpty(Not cTrackers) Then
            For I = 0 To UBound(cTrackers)
                With cTrackers(I)
                    sProp = .ID & "," & .rcItem.Left & "," & .rcItem.Top & "," & .rcItem.Right & "," & .rcItem.Bottom
                    propBag.WriteProperty "FM.Trk" & I, sProp
                    If Not cTrackers(I).tTip Is Nothing Then
                        propBag.WriteProperty "FM.TrkTip" & I, cTrackers(I).tTip.TipText, "nc"
                    End If
                End With
            Next
        End If
        propBag.WriteProperty "FM.Trackers", I, 0
    End If
    
    sProp = ""
    If Not IsArrayEmpty(Not cInset) Then
        For I = 0 To 3
            sProp = sProp & "," & CByte(cInset(I))
        Next
    End If
    propBag.WriteProperty "FM.Inset", Mid$(sProp, 2), ""
    propBag.WriteProperty "FM.Borders", CStr(cBorderWidthCx) & "," & CStr(cBorderWidthCy), "4,4"
    propBag.WriteProperty "FM.OD", mNotifyUser, 0
    
    Set iBag = Me.Active
    iBag.WriteProperties propBag, Limits, sharedImgPtr
    Set iBag = Me.Inactive
    iBag.WriteProperties propBag, Limits, sharedImgPtr
    Set iBag = Me.MinMax
    iBag.WriteProperties propBag, Limits
    
    If (Limits And exSysTray) = exSysTray Then
        If Not cSysTray Is Nothing Then
            Set iBag = cSysTray
            iBag.WriteProperties propBag, Limits, pHwnd
        Else
            propBag.WriteProperty "ST.Count", 0, 0
        End If
    End If
    Set iBag = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : clsFrame.iTimer_TimerEvent
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Timer callback function
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub iTimer_TimerEvent(ByVal hWnd As Long, ByVal TimerID As Long)

    Dim wPTs As POINTAPI, wRect As RECT
    Dim tID As Long
    
    ' if the button has already been painted by some other action, simply abort
    If cTrackID = 0 Then
        CreateTimer 0
        
    Else
        GetCursorPos wPTs
        GetWindowRect hWnd, wRect
        ' offset the points to our titlebar location
        wPTs.x = wPTs.x - wRect.Left
        wPTs.Y = wPTs.Y - wRect.Top
        
        ' get button ID from the hit test code
        tID = TrackerHitTest(wPTs.x, wPTs.Y)
        ' if the mouse is still over the button highlighted, no action needed
        If tID <> cTrackID Then clearTracker
    
    End If
    

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.SysTrayInitialized
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Helper function to determine if a system tray class is in use
'---------------------------------------------------------------------------------------
'
Friend Property Get SysTrayInitialized() As Boolean
    SysTrayInitialized = Not cSysTray Is Nothing
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.NoRedraw
' DateTime  : 9/15/2005
' Author    : LaVolpe
' Purpose   : Prevents redrawing while multiple properties are being changed
'---------------------------------------------------------------------------------------
'
Public Property Let noRedraw(bNoRedraw As Boolean)
    cAutoRedraw = Not bNoRedraw
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.RefreshWindow
' DateTime  : 9/15/2005
' Author    : LaVolpe
' Purpose   : Forces window to repaint
'---------------------------------------------------------------------------------------
'
Private Sub RefreshWindow()
    If pHwnd Then
        If cAutoRedraw Then ForceRefresh pHwnd
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.sharedProperty
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : A way to get other class properties without having to
'             worry about caching them here & keeping them updated
'---------------------------------------------------------------------------------------
'
Private Property Get sharedProperty(propID As eProperties) As Long

    If mMainImplementation = 0 Then Exit Property
    
    Dim cImp As iImpProps
        GetObjectFromPointer mMainImplementation, cImp
        sharedProperty = cImp.GetPropValue(propID)
        Set cImp = Nothing

End Property

'---------------------------------------------------------------------------------------
' Procedure : clsFrame.Class_Initialize/Class_Terminate
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Class initialization/destruction routine
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    cBorderWidthCx = 4
    cBorderWidthCy = 4
    Set cColors(1) = New clsBorders
    Set cColors(0) = New clsBorders
    cColors(1).pClassUsage = True
End Sub

Private Sub Class_Terminate()
    CreateTimer 0
    Set cColors(0) = Nothing
    Set cColors(1) = Nothing
    Set cMinMax = Nothing
End Sub

