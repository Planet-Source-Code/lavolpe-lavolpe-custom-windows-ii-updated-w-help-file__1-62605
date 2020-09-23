VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenubar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' menus designed around following standards:
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/resources/menus/menuprogrammingconsiderations.asp

' Probably the most complicated class in this project. As the menubar is the most
' painted and modified object in the window

' This is the menubar and associated drawing for that bar
' Major changes from previous version (Custom Windows Part I)
' - user can override and custom draw anything thru the iOwnerDrawn implementation
' - menu loop. The previous version attempted to use keybd & mouse hooks to intercept
'   every mouse movement, menu message and keyboard activity to control how the menus
'   would be displayed, selected, highlighted, etc.
'   This version uses a keyboard hook that traps very few keys and also a menu message
'   filter that does not intercept (only previews) messages. No mouse hook is used.
'   The main difference is that the menu messages are allowed to flow thru to the
'   application & are rerouted to the menubar as needed. This new approach
'   significantly reduces the amount of code; therefore reduces the number of possible
'   errors/actions that the class is responsible for. O/S handles most of the traffic
' - overall the routines are still difficult, but much easier to troubleshoot


Private Declare Function GetMenuItemInfo_String Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO_String) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function GetSubMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long


Private Type MENUITEMINFO_String        ' used to retrieve/store menu items
     cbSize As Long                      ' this structure is used with all O/S
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As String
     cch As Long
End Type
Private Type MenuDataStructure
    ID As Long
    flags As Long
    itemData As Long
    Position As RECT
    Caption As String
    odProp As Integer       '1=img OD, 2=Full OD, 4=Img RtAligned if Img OD
End Type

Private Type MenubarTracker
    TimerOn As Byte         'timer created/destroyed
    Active As Byte          '0=no hooks, 1=hooking; menubar active & in loop
    TriState As Byte        '0=not selected, see below for more info
    LastDrawn As Integer    'last menu item drawn on the window (1 LBound vs 0 LBound)
    Position As Integer     'current position; -1 = sysmenu
    HitTest As Integer      'current hit test (1 LBound vs 0 Lbound)
    flags As Long           '2=can move left,4=can move right,8=align right;32=align bottom
                            '16=menubar hittest but not in menuitem,64=not submenu-less
    hWnd As Long            'menubar owner
    ID As Long              'panel ID/menu ID
    CursorPt As POINTAPI    'current XY of cursor
End Type
'^^ Regarding the TriState flag. The menu bar can be in three states:
' Not active but highlighting menu items as mouse moves over them
'   -- .TriState=1 and .Active=0 identifies that scenario
' Active but not showing submenus. This can happen when just the Alt key is pressed
' or user hits Escape when the last submenu item is displayed
'   -- .TriState=1 and .Active=1 identifies that scenario
' Active and showing submenus. Occurs when user clicks or Alt+Acclerator to activate
' the menubar; whether or not the menu item actually has submenus. This can also
' occur when the user displays the system menu via the system icon & then
' uses the left/right arrow keys to navigate the menubar
'   -- .TriState=2 and .Active=1 identifies that scenario

Private Enum eODActions
    odNormal = 0    'normal state
    odSelect = 1    'Selected/Pushed state
    odDisable = 2   'Disabled
    odHover = 4     'MouseOver
    odActive = 8    ' window is active or not
    odMenuBar = 16  ' menubar being drawn
    odMenuItem = 32 ' menuitem being drawn/updated
End Enum
Private Const mbBottomAligned = 2
Private Const mbLeftAligned = 4
Private Const mbRightAligned = 8
Private Const mbHideDisabled = 16
Private Const mbHideMenu = 32
Private Const mbHideMenuCtrlKey = 64

Private Const kbCanExpand = 4
Private Const kbCanCollapse = 2
Private Const kbHasSubMenu = 1
Private Const kbNonMnuItem = 8

Private Const MIIM_ID As Long = &H2
Private Const MIIM_STATE As Long = &H1
Private Const MIIM_STRING As Long = &H40
Private Const MIIM_DATA As Long = &H20


Private Const cHoverBorderCx As Long = 4     ' left edge buffer btwn menu caption & selection border
Private Const cHoverBorderCy As Long = 3     ' top edge buffer btwn menu caption & selection border
Private Const IconSize As Long = 16


Private cBorderWidthCx As Byte
Private cBorderWidthCy As Byte

Public Enum eTrackingFlags
    mbActiveStatus = 0
    mbTriState = 1
    mbPosition = 2
    mbID = 3
    mbHwnd = 4
    mbLastDrawn = 5
    mbFlags = 6
    mbHitTest = 7
End Enum

Private cFont As Long                       ' menu bar font; must be destroyed
Private cMenu As Long                       ' handle to the window's menu bar
Private cSysMenu As Long                    ' handle to the system menu
Private cBar As RECT                        ' menu bar size in screen coords
Private cOptions As Integer                 ' collection of property flags
Private cAutoRedraw As Boolean              ' flag used to prevent redrawing during multiple property settings

Private cMenuOrder() As Integer
Private cClipRgn As Long
Private cActive As Byte

Private cMBarItem() As MenuDataStructure    ' collection of menu item info
Private cTracker As MenubarTracker
Private cColors(0 To 1) As clsMenuBarColors

Private frameOffset As Byte ' 0 or 1
Private mMainImplentation As Long
Private mODImplementation As Long
Private mNotifyUser As Long

Implements iTimer
Implements iPropertyBag

'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.Active/Inactive
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Exposes the active and inactive color classes
'---------------------------------------------------------------------------------------
'
Public Property Get Active() As clsMenuBarColors
    If cColors(1) Is Nothing Then Set cColors(1) = New clsMenuBarColors
    Set Active = cColors(1)
End Property
Public Property Get Inactive() As clsMenuBarColors
    If cColors(0) Is Nothing Then Set cColors(0) = New clsMenuBarColors
    Set Inactive = cColors(0)
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.ResetToSystemDefaults
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Resets menubar properties to system defaults
' Comments  : Can also be called from the main class's ResetToSystemDefaults sub
'---------------------------------------------------------------------------------------
'
Public Sub ResetToSystemDefaults(Options As eBarResets)
    If (Options And rstFont) = rstFont Or Options = rstAll Then
        If cFont Then DeleteObject cFont
        cFont = 0   ' not setting to zero won't force use of system font
        Set Me.Font = Nothing ' will now use system font
    End If
    If (Options And rstBorderSizes) = rstBorderSizes Or Options = rstAll Then
        Me.BorderWidth = 0
        Me.BorderHeight = 0
    End If
    Me.Active.ResetToSystemDefaults Options
    Me.Inactive.ResetToSystemDefaults Options
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.Alignment
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Set/Get the menubar alignment in relation to the client rectangle
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Let Alignment(mbAlign As eBarAlignment)
    
    ' vertical menubar options not coded at this time
    ' Note a lot of code is already added for vertical menubars.
    ' That may come out in a future update.
    
    Dim oldOption As Byte
    
    oldOption = cOptions    ' cache in case user passed something bogus
    cOptions = cOptions And Not mbBottomAligned And Not mbRightAligned And Not mbLeftAligned
    
    Select Case mbAlign
        'Case barAlignRight: cOptions  = cOptions Or mbRightAligned
        'Case barAlignLeft: cOptions = cOptions Or mbLeftAligned
        Case barAlignBottom: cOptions = cOptions Or mbBottomAligned
        Case barAlignTop: ' no flag needed
        Case Else
            cOptions = oldOption
            Exit Property
    End Select
    Set Me.Font = Nothing 'automatically refreshes window

End Property
Public Property Get Alignment() As eBarAlignment
    
    If (cOptions Or mbBottomAligned) = cOptions Then
        Alignment = barAlignBottom
    Else
        ' vertical menubar not coded at this time
        If (cOptions Or mbLeftAligned) = cOptions Then
            Alignment = barAlignLeft
        Else
            If (cOptions Or mbRightAligned) = cOptions Then
                Alignment = barAlignRight
            Else
                Alignment = barAlignTop
            End If
        End If
    End If
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.NoRedraw
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Prevents redrawing while multiple properties are being changed
'---------------------------------------------------------------------------------------
'
Public Property Let noRedraw(bNoRedraw As Boolean)
    cAutoRedraw = Not bNoRedraw
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.ShowMenuBar
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Toggle hiding or showing the menubar
'---------------------------------------------------------------------------------------
'
Public Property Get ShowMenuBar() As Boolean
    ShowMenuBar = ((cOptions And mbHideMenu) = 0)
End Property
Public Property Let ShowMenuBar(bShow As Boolean)
    If bShow Then
        cOptions = cOptions And Not mbHideMenu
        If pHasMenuBar Then RefreshWindow
    Else
        cOptions = cOptions Or mbHideMenu
        RefreshWindow
    End If
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.ShowMenuBar
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Internal return to indicate if a menubar is being displayed
'---------------------------------------------------------------------------------------
'
Friend Property Get pHasMenuBar() As Boolean
    If cMenu <> 0 Then
        If (cOptions And mbHideMenu) = 0 Then
            ' not hidden, but still may not have titlebar
            If Not IsArrayEmpty(Not cMenuOrder) Then pHasMenuBar = True
        End If
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.HideDisabledItems
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Option to hide disabled menu bar items
'---------------------------------------------------------------------------------------
'
Public Property Get HideDisabledItems() As Boolean
    HideDisabledItems = ((cOptions Or mbHideDisabled) = cOptions)
End Property
Public Property Let HideDisabledItems(bHide As Boolean)
    If bHide Then
        cOptions = cOptions Or mbHideDisabled
    Else
        cOptions = cOptions And Not mbHideDisabled
    End If
    RefreshWindow
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.Font
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : Let/Set/Get the menubar font
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Let Font(newFont As StdFont)
    ' prevents error when not using Set Font = Font Object
    Set Font = newFont
End Property
Public Property Set Font(newFont As StdFont)
' modifies the font used for the menubar
    Dim tFont As Long
    
    tFont = FontToHandle(newFont, True)
    
    If tFont <> 0 Then
        ' font created, replace the current font with the new one
        If cFont Then DeleteObject cFont
        cFont = tFont
        RefreshWindow
    'else the font failed, continue using the current font
    End If

End Property
Public Property Get Font() As StdFont
' function returns a stdFont object from a memory handle font
    Set Font = HandleToFont(cFont)
End Property


'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.MeasureMenubar
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : Returned is the adjusted height of a dynamically sized menubar
' Comments  : Also calculates position of each menuitem within the menubar
'---------------------------------------------------------------------------------------
'
Friend Function MeasureMenubar(ByVal cRectPtr As Long, ByVal isActive As Byte, _
            rtnAlign As eBarAlignment, ByVal HideSysMenu As Boolean) As Long

' Each time a WM_NCCalcSize message is sent, we must iterate thru the menubar items
' Why? Well several reasons....
'   1. The width of the form could have changed; menuitem rows may need to be shifted
'   2. When you use something like the following, VB will generate a WM_NCCalcSize message
'           mnuControl.Enabled = False/True or mnuControl.Visible = False/True
'   3. When you add or remove a menu control like the following, VB also sends the message
'           Controls.Add or Load Object
' Unfortunately, many other actions can generate a WM_NCCalcSize. Therefore we have
' no real way of knowing why the message was sent. Therefore we will remeasure
' and reposition items each time this routine is called.

' Note: checks could be made to prevent repositioning items. For example, if
' the form width is increased and a menubar is already a single line, then
' there is no need to reposition menu items, right? Right, but we also need
' to check to see if a menu item property changed: visible, enable, or caption
' Since we are forced to iterate thru the menubar items anyway, might as well
' just remeasure vs adding all the extra code to see if positioning is really
' needed. Any way, the menubar items, by nature, are generally small in number


    Dim hMenu As Long, nrMenuItems As Long  ' menubar & nr of items
    Dim cRect As RECT                       ' used for positioning/sizing
    Dim I As Long, J As Long                        ' generic variables
    Dim Yoffset As Long, itemHeight As Long         ' used to adjust menubar or items
    Dim bNoMenubar As Boolean                       ' whether or not a menubar is displayed
    Dim hFont As Long, hDC As Long, DTFlags As Long ' used for measuring captions
    Dim oldMenuItems() As MenuDataStructure         ' copy of existing menu items
    ' used for owner drawn menu items if applicable
    Dim cClient As iOwnerDrawn, CID As CustomItemDraw_LV, bOwnerDrawn As Integer
    
    cAutoRedraw = True ' reset internal flag
    cActive = isActive
    
    ' ensure the menu handle didn't change on us. Expected is actual handle or zero
    hMenu = GetMenu(cTracker.hWnd)
    If hMenu <> 0 Then
        ' cache the new menubar handle if needed
        If hMenu <> cMenu Then cMenu = hMenu
    End If
    If HideSysMenu Then
        cSysMenu = 0
    Else
        ' always check. It could have been removed by user
        hMenu = GetSystemMenu(cTracker.hWnd, 0)
        If hMenu Then
            If hMenu <> cSysMenu Then cSysMenu = hMenu
        End If
    End If
    
    ' no need to continue if one of the following apply
    If cMenu = 0 Then
        bNoMenubar = True ' window has no menu
    Else
        If ((cOptions Or mbHideMenu) = cOptions) Then bNoMenubar = True ' the menubar is hidden
    End If
    
    ' if we were tracking the menubar when something modified the menubar dimensions
    If cTracker.Active Then
        If cTracker.Position > -1 Then
            'if non-sysMenu being displayed send a WM_CancelMode to clear the menu only if necessary
            If (cTracker.flags Or kbHasSubMenu) = cTracker.flags Then
                ' above says menubar items has submenus & below checks to see if displayed
                If cTracker.TriState = 2 Then PostMessage cTracker.hWnd, WM_CANCELMODE, 0, ByVal 0&
                ' ^^ don't use SendMessage here... may interfere with the WM_NCCalc message
                '    being processed at this time.
            End If
        End If
        pTracking(mbActiveStatus) = 0 ' Remove the keyboard hook
    Else
        If pTracking(mbActiveStatus) <> 0 Then
            pTracking(-1) = 0       ' reset some key tracking variables
        End If
    End If
    
    ' get client rectangle; entire size is avaiable for use for the menubar
    CopyMemory cRect, ByVal cRectPtr, &H10
    
    nrMenuItems = MenuItemCount(0)
    If nrMenuItems < 1 Then
        Erase cMBarItem             ' all menu items are set visible = false
        Erase cMenuOrder            ' or no menu exists
        'Exit Function
        bNoMenubar = True
    End If
    
    If Not bNoMenubar Then
        ' depending on the orientation of our menubar and the titlebar, adjust
        ' the client rectangle to remove the single pixel border around the
        ' client rect that the titlebar placed
        Select Case rtnAlign
            Case barAlignBottom
                If Me.Alignment = barAlignBottom Then cRect.Bottom = cRect.Bottom + 1
            Case barAlignTop
                If Me.Alignment = barAlignTop Then cRect.Top = cRect.Top - 1
            Case barAlignLeft
                If Me.Alignment <> barAlignLeft Then cRect.Left = cRect.Left - 1
            Case barAlignRight
                If Me.Alignment <> barAlignRight Then cRect.Right = cRect.Right + 1
        End Select
        cBar = cRect                        ' let this be the menubar for now
        
        If IsArrayEmpty(Not cMBarItem) Then
            ReDim oldMenuItems(-1 To -1)    ' prevents error in For:Next below
        Else
            oldMenuItems() = cMBarItem()    ' copy the recent menu items
        End If
        ' erase the current menubar item collection
        ReDim cMBarItem(0 To nrMenuItems - 1)
    
        ' cache vs recalculating each loop
        DTFlags = DT_CALCRECT Or DT_SINGLELINE Or DT_LEFT
    
        ' place our font in a DC for the DrawText API
        hDC = Canvas(False)
        If hDC = 0 Then Exit Function
        hFont = SelectObject(hDC, cFont)
    
        If (mNotifyUser And odMenuItem_ImageOnly) = odMenuItem_ImageOnly Then
            bOwnerDrawn = odMenuItem_ImageOnly
        Else
           If (mNotifyUser And odMenuItem_Complete) = odMenuItem_Complete Then bOwnerDrawn = odMenuItem_Complete
        End If
        If bOwnerDrawn Then bOwnerDrawn = Abs((mODImplementation <> 0) * bOwnerDrawn)
        If bOwnerDrawn Then
            On Error Resume Next ' always, to try & prevent user from crashing
            
            ' user wants to control some aspects of drawing menu items
            GetObjectFromPointer mODImplementation, cClient
            CID.hDC = hDC   ' set, doesn't change in this routine
            
            For I = 0 To nrMenuItems - 1
                ' get the menu caption & any pertinent menu properties
                GetMenuItemCaption I
                CID.itemData = 0
                CID.itemOD = bOwnerDrawn
                For J = 0 To UBound(oldMenuItems)
                    ' see if we have it already, if so, transfer some properties
                    ' Note: really don't want to use a loop, but menubar items are
                    ' generally small in number & I don't think this is too time consuming
                    If oldMenuItems(I).ID = cMBarItem(J).ID Then
                        cMBarItem(I).odProp = oldMenuItems(J).odProp Or bOwnerDrawn
                        cMBarItem(I).itemData = oldMenuItems(J).itemData
                        CID.itemData = cMBarItem(I).itemData        ' user-supplied data
                        Exit For
                    End If
                Next
                If ((cOptions Or mbHideDisabled) = cOptions) And ((cMBarItem(I).flags And MF_DISABLED) = MF_DISABLED) Then
                    ' don't send user disabled items to measure if hiding disabled items is active
                    SetRect cMBarItem(I).Position, 0, 0, 0, 0
                Else
                    ' prepare the UDT to pass to the user
                    With CID
                        .itemID = cMBarItem(I).ID       ' menu id
                        .itemPos = I                    ' position on the menubar
                        .itemState = Abs((cMBarItem(I).flags And MF_DISABLED) = MF_DISABLED)
                        SetRect .rcItem, 0, 0, 0, 0
                    End With
                    ' in case user does something unwise like unloading while measuring items
                    If mODImplementation = 0 Then
                        Set cClient = Nothing
                        Exit Function
                    End If
                    ' send info to the user so the image can be measured
                    cClient.OwnerDrawMessage omMeasureMenuItem, 0, VarPtr(CID), cMBarItem(I).Caption
                
                    With cMBarItem(I)
                        ' cache the results
                        .itemData = CID.itemData        ' user supplied data
                        If bOwnerDrawn = odMenuItem_Complete Then
                            ' then completely owner drawn
                            .Position = CID.rcItem      ' size of menu item
                            .odProp = odMenuItem_Complete
                            OffsetRect .Position, -CID.rcItem.Left, -CID.rcItem.Top
                        Else
                            .odProp = (CID.itemOD And odImgAlignRight) Or (CID.itemOD And odImgAlignLeft) Or odMenuItem_ImageOnly
                            ' user not fully drawing item, we calculate the size
                            DrawText hDC, .Caption, -1, .Position, DTFlags
                            ' user wants to draw an image with the item otherwise wouldn't have requested this option
                            ' ensure enough space for the image, all items are
                            ' automatically buffered with additional space later
                            If .Position.Bottom < IconSize Then .Position.Bottom = IconSize
                            ' ensure enough width for the icon
                            If ((.odProp And odImgAlignLeft) = odImgAlignLeft) Or _
                                ((.odProp And odImgAlignRight) = odImgAlignRight) Then
                                    .Position.Right = .Position.Right + IconSize
                            End If
                            .Position.Right = .Position.Right + cHoverBorderCx
                        End If
                        ' keep track of tallest menu item. All menu items will have the
                        ' same size on the menubar
                        If .Position.Bottom > itemHeight Then itemHeight = .Position.Bottom
                        
                    End With
                End If
            Next
            Set cClient = Nothing           ' clean up
            Erase oldMenuItems
        
        Else
            ' no user intervention in drawing/measuring menu items
            For I = 0 To nrMenuItems - 1
                ' get the menu caption & any pertinent menu properties
                GetMenuItemCaption I
                ' measure the menuitem width & height
                DrawText hDC, cMBarItem(I).Caption, -1, cMBarItem(I).Position, DTFlags
            Next
            itemHeight = cMBarItem(0).Position.Bottom
        End If
        
        ' more clean up: remove & replace the font
        SelectObject hDC, hFont
        
        If (mNotifyUser And odMenuBarBkg) = 0 Or (mODImplementation = 0) Then
            ' get whether this menubar has a frame or not & cache the result as a byte
            cColors(isActive).GetFrame I, 0, 0
            If I > 0 Then frameOffset = 1 Else frameOffset = 0
        Else
            frameOffset = 0
        End If
        
        ' call function to position each menu item. That function will also hide
        ' any disabled menu items if the HideDisabledItems is True. Because of
        ' that, it is possible that no items will be displayed if all are disabled.
        ' The following function returns true if any items are displayed and also
        ' modifies the cBar rectangle to encompass all the displayed items
        
        If SetMenuItemPositions(itemHeight) Then
            
            If Me.Alignment = barAlignBottom Then ' bottom aligned
                ' special case. When the menubar is bottom aligned, the bottom
                ' of the cBar rect can extend over the bottom border an/or
                ' the titlebar if it is also bottom aligned: Adjust
                Yoffset = cRect.Bottom - (cBar.Bottom - cBar.Top) - cBar.Top
                OffsetRect cBar, 0, Yoffset
                
                ' more issues with bottom aligned menubars. If the window is
                ' resized vertically and the menubar has more than one line of items,
                ' the top of the menubar will be cut off if the window is sized
                ' very small vertically. I personally, think this looks cheap.
                ' So, I will shift the top of the menubar down so that that bottom
                ' is cutoff vs the top. This requires more checks
                If cBar.Top < cRect.Top Then
                    Yoffset = Yoffset + cRect.Top - cBar.Top
                    OffsetRect cBar, 0, cRect.Top - cBar.Top
                    cBar.Bottom = cRect.Bottom
                End If
                ' now we need to shift each of the menu items to their new relative position
                For I = 0 To UBound(cMenuOrder)
                    OffsetRect cMBarItem(cMenuOrder(I)).Position, 0, Yoffset
                Next
            Else
                ' ensure bar doesn't overlap borders or a bottom-aligned titlebar
                If cBar.Bottom > cRect.Bottom Then cBar.Bottom = cRect.Bottom
            End If
            
            ' a matter of taste here. I'm offsetting the client rectangle based
            ' on the relationship between the titlebar and menubar and their
            ' orientations. Basically, I want a one to two pixel border around
            ' the client rectangle
            Select Case rtnAlign
                Case barAlignTop, barAlignBottom, 0
                    Select Case Me.Alignment
                    Case barAlignLeft:
                        cRect.Left = cBar.Left + 1
                    Case barAlignRight
                        cRect.Right = cBar.Right - 1
                    Case barAlignBottom
                        cRect.Bottom = cBar.Top - 1
                    Case barAlignTop
                        cRect.Top = cBar.Bottom + 1
                    End Select
                    
                Case barAlignLeft, barAlignRight ' vertical titlebar
                    Select Case Me.Alignment
                    Case barAlignBottom
                        cRect.Bottom = cBar.Top - 1
                    Case barAlignTop
                        cRect.Top = cBar.Bottom + 1
                    Case barAlignLeft
                        cRect.Left = cBar.Right + 1
                    Case barAlignRight
                        cRect.Right = cBar.Left - 1
                    End Select
                    If rtnAlign = barAlignLeft Then
                        If Me.Alignment <> barAlignLeft Then cRect.Left = cRect.Left + 1
                    Else
                        If Me.Alignment <> barAlignRight Then cRect.Right = cRect.Right - 1
                    End If
            End Select
           
            ' create a clipping region for the menubar. This is needed 'cause the
            ' menubar can be partially hidden when sizing window to small measures.
            ' Without a clipping region, the menuitems can be drawn over just
            ' about anything since we paint to the WindowDC
            If cClipRgn Then DeleteObject cClipRgn
            cClipRgn = CreateRectRgn(cBar.Left + cBorderWidthCx, cBar.Top + cBorderWidthCy, _
                cBar.Right - cBorderWidthCx, cBar.Bottom - cBorderWidthCy)
            
            ' notify the colors class so it can modify the image background if used
            cColors(0).pSized 0, 0, True, (mODImplementation <> 0 And ((mNotifyUser And odMenuBarBkg) = odMenuBarBkg))
            cColors(1).pSized 0, 0, True, (mODImplementation <> 0 And ((mNotifyUser And odMenuBarBkg) = odMenuBarBkg))
            
            ' return the menubar rect via a pointer
            rtnAlign = Me.Alignment
            MeasureMenubar = VarPtr(cBar)
            
        Else
            rtnAlign = 0                ' no menu items
        End If
    End If
    
    ' update the client rectangle size
    CopyMemory ByVal cRectPtr, cRect, &H10
        
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.SetMenuItemPositions
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Positions the menubar items in relation to the menubar's top/left
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Function SetMenuItemPositions(itemHeight As Long) As Boolean

    Dim mnuItem As Long, mnuPos As Long
    Dim maxWidth As Long
    Dim mPos As POINTAPI
    
    ReDim cMenuOrder(0 To UBound(cMBarItem))
    '^^ keeps track of which menu items are visible when HideDisabledItems is applied
    
    ' calculate the position of the first menu item
    mPos.x = cBar.Left + 1 + frameOffset + cBorderWidthCx
    mPos.Y = cBar.Top + frameOffset + 1 + cBorderWidthCy
    
    ' now loop thru each menu item
    For mnuItem = 0 To UBound(cMBarItem)
    
        With cMBarItem(mnuItem)
            ' reset the item's position to 0,0
            OffsetRect .Position, -.Position.Left, -.Position.Top
            
            ' determine if this item is disabled & if HideDisabledItems is set
            If ((cOptions Or mbHideDisabled) = cOptions) Then ' hide disbled items
                If (.flags And MF_DISABLED) = MF_DISABLED Then SetRect .Position, 0, 0, 0, 0
            End If
            
            If .Position.Right > 0 Then ' otherwise zero-len caption or hidden disabled
                ' make all menuitems the same height
                .Position.Bottom = .Position.Top + itemHeight
                
                ' add the item to the visible menu item array & start counting items
                cMenuOrder(mnuPos) = mnuItem
                mnuPos = mnuPos + 1
                
                ' see if this item will fit on the existing row
                If .Position.Right + mPos.x + 3 + cHoverBorderCx * 2 > cBar.Right - frameOffset - cBorderWidthCx Then
                    ' unless this is the 1st item on the row, move it to next row
                    If mPos.x > cBar.Left + 1 + frameOffset + cBorderWidthCx Then
                        ' update the new row X & Y coordinates
                        If (.odProp And 2) = 2 Then
                            mPos.Y = .Position.Bottom - .Position.Top + mPos.Y + 1
                        Else
                            mPos.Y = .Position.Bottom - .Position.Top + mPos.Y + cHoverBorderCy * 2 + 1
                        End If
                        mPos.x = cBar.Left + 1 + frameOffset + cBorderWidthCx
                    End If
                End If
                ' resize dimensions to allow for drawing 3D boxes
                If (.odProp And 2) = 2 Then
                    OffsetRect .Position, mPos.x, mPos.Y
                    mPos.x = .Position.Right + 1 ' keep track of current X coord
                Else
                    InflateRect .Position, cHoverBorderCx, cHoverBorderCy
                    ' shift it back to position. Inflate will change the top/left coords
                    OffsetRect .Position, cHoverBorderCx + mPos.x, cHoverBorderCy + mPos.Y
                    mPos.x = .Position.Right + 1    ' keep track of current X coord
                End If
                ' keep track of the widest item to be used to size the
                ' common-use bitmap where all items, all windows, are drawn
                If (.Position.Right - .Position.Left) > maxWidth Then maxWidth = (.Position.Right - .Position.Left)
                
            End If
        End With
    Next
    
    If mnuPos = 0 Then ' no visible menu items or all are disabled
        cBar.Bottom = cBar.Top  ' make the menubar invalid
        Erase cMenuOrder
    Else
        ' ensure our common-use bitmap is wide enough for the widest menu item
        ExpandDC maxWidth + 17, (cMBarItem(cMenuOrder(mnuPos - 1)).Position.Bottom - cMBarItem(cMenuOrder(mnuPos - 1)).Position.Top) * 2 + 1, False, cTracker.hWnd
        ' redim the array for only those visible items
        ReDim Preserve cMenuOrder(0 To mnuPos - 1)
        ' adjust the menubar's bottom edge & return
        cBar.Bottom = cMBarItem(mnuPos - 1).Position.Bottom + 2 + frameOffset + cBorderWidthCy
        SetMenuItemPositions = True
    End If
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.DrawMenuBar
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : refreshes the menubar
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Sub DrawMenuBar(ByVal isActive As Byte)

    cActive = isActive
    
    ' ensure we have items to draw
    If IsArrayEmpty(Not cMenuOrder) Then Exit Sub
    If (cOptions Or mbHideMenu) = cOptions Then Exit Sub
    
    Dim Cy As Long, bkgImg As StdPicture, hDC As Long
    Dim hBrush As Long, bkStyle As eFillStyle, bGrayScaled As Boolean
    Dim fColor(0 To 1) As Long, fxStyle As Long
    
    
    hDC = Canvas(True)
    
    ' if user is owner-drawing the menubar background, get copy of their implementation
    If (mNotifyUser And odMenuBarBkg) = odMenuBarBkg And mODImplementation <> 0 Then
    
        On Error Resume Next ' always, to try & prevent user from crashing
        Dim cClient As iOwnerDrawn, BKG As BkgAction_LV
        GetObjectFromPointer mODImplementation, cClient
        BKG.hDC = hDC
        BKG.rcItem = cBar
        BKG.rcExtra = cBar
        InflateRect BKG.rcExtra, -cBorderWidthCx, -cBorderWidthCy
        cClient.OwnerDrawMessage omDrawMBarBkg, isActive + 0, VarPtr(BKG), ""
        Set cClient = Nothing
        SelectClipRgn hDC, ByVal 0& ' always remove any clipping region client may have left behing
        
    Else
    
        ' get the menubar background image if applicable
        Call cColors(isActive).GetImageBackground(bkgImg, fxStyle, bGrayScaled)
        
        ' unless we have a transparent background, don't continue
        If (bkStyle And Not bfGrayScaled) <> bfTransparent Then
        
            ' if image exists, then it will completely fill the menubar
            If Not bkgImg Is Nothing Then
                FillBarImage bkgImg, hDC, cBar.Left + frameOffset, cBar.Top + frameOffset, cBar.Right - cBar.Left - frameOffset * 2, cBar.Bottom - cBar.Top - frameOffset * 2, fxStyle, False, bGrayScaled
                Set bkgImg = Nothing
                
            Else
                ' get the menubar background colors & style/orientation of gradients
                Call cColors(isActive).GetBackColors(bkStyle, fColor(0), fColor(1))
                bGrayScaled = ((bkStyle And bfGrayScaled) = bfGrayScaled)
                bkStyle = bkStyle And Not bfGrayScaled
                
                If bkStyle = bfGradientEW Or bkStyle = bfGradientNS Then 'gradient used
                    ' the CY when negative means vertical gradient vs horizontal
                    Cy = cBar.Bottom - frameOffset
                    If bkStyle = bfGradientNS Then Cy = -Cy
                    ' fill the menubar
                    GradientFill fColor(0), fColor(1), hDC, cBar.Left + frameOffset, cBar.Top + frameOffset, cBar.Right - frameOffset, Cy, 0, , , bGrayScaled
                    
                ElseIf bkStyle = bfSolid Then ' no gradients, use solid color
                    If bGrayScaled Then GrayScaleColor VarPtr(fColor(0))
                    hBrush = CreateSolidBrush(fColor(0))
                    FillRect hDC, cBar, hBrush
                    DeleteObject hBrush
                End If
            
            End If
        End If
        
        ' if user wants a frame around the menubar, add that now
        If frameOffset Then
            ' get the frame color(s)
            cColors(isActive).GetFrame fxStyle, fColor(0), fColor(1)
            ' draw the frame
            Select Case fxStyle
            Case bxFlat ' single color
                hBrush = CreateSolidBrush(fColor(0))
                FrameRect hDC, cBar, hBrush
                DeleteObject hBrush
            Case bx3D   ' two colors
                hBrush = SelectObject(hDC, CreatePen(0, 1, fColor(0)))
                MoveToEx hDC, cBar.Left, cBar.Bottom - 1, ByVal 0&
                LineTo hDC, cBar.Left, cBar.Top
                LineTo hDC, cBar.Right - 1, cBar.Top
                DeleteObject SelectObject(hDC, CreatePen(0, 1, fColor(1)))
                LineTo hDC, cBar.Right - 1, cBar.Bottom - 1
                LineTo hDC, cBar.Left - 1, cBar.Bottom - 1
                DeleteObject SelectObject(hDC, hBrush)
            End Select
        End If
        
    End If
    
    ' call routine to create snapshot of the menubar if we don't control
    ' its drawing (owner drawn, or stretched image). Depending on the availability
    ' of the menubar to our DLL, we may need to cache a snapshot of it
    If cMBarItem(UBound(cMenuOrder)).Position.Top = cMBarItem(cMenuOrder(0)).Position.Top Then
        ' single line menu bar, we don't need to copy the entire menubar, do we?
        cColors(isActive).pSized VarPtr(cBar), cMBarItem(UBound(cMenuOrder)).Position.Right, False, False
    Else
        cColors(isActive).pSized VarPtr(cBar), cBar.Right, False, False
    End If
    
    SelectClipRgn hDC, cClipRgn
        DrawMenuItem mcStandard, hDC, 0, 0, isActive, -1
    SelectClipRgn hDC, ByVal 0&

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.DrawMenuItem
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : This is called to draw a single item or the entire menu
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub DrawMenuItem(Action As Long, hDC As Long, ByVal x As Long, ByVal Y As Long, _
                        ByVal isActive As Byte, menuItem As Long)
    
    Dim mnuFrom As Long, mnuTo As Long
    Dim I As Long
    
    If menuItem < 0 Then    ' when -1 is passed, all items are drawn
        mnuFrom = 0
        mnuTo = UBound(cMenuOrder)
    Else                    ' otherwise one item is being drawn (menu in loop or hovering)
        mnuFrom = menuItem
        mnuTo = mnuFrom
    End If
    
    ' used for owner-drawn menu items
    Dim cClient As iOwnerDrawn, CID As CustomItemDraw_LV
    Dim bOwnerDrawn As Integer, odYoffset As Long
    
    Dim DTFlags As Long, fxID As Integer
    Dim hOldPen As Long, hFont As Long
    Dim fColor(0 To 5) As Long, fx(0 To 2) As Long
    Dim fBorder(0 To 2) As Long
    ' 3 rects for 3 objects (t=Text, i=Image, m=menuitem total)
    Dim tRect As RECT, iRect As RECT, mRect As RECT
        
    ' select our font & colors
    hFont = SelectObject(hDC, cFont)
    ' cache vs recalculating each time
    DTFlags = DT_VCENTER Or DT_SINGLELINE Or DT_CENTER
    
    If (mNotifyUser And odMenuItem_ImageOnly) = odMenuItem_ImageOnly Then
        bOwnerDrawn = 1
    Else
        If (mNotifyUser And odMenuItem_Complete) = odMenuItem_Complete Then bOwnerDrawn = 2
    End If
    If bOwnerDrawn Then bOwnerDrawn = Abs((mODImplementation <> 0) * bOwnerDrawn)
    If bOwnerDrawn Then
        On Error Resume Next ' always, to try & prevent user from crashing
        GetObjectFromPointer mODImplementation, cClient
        CID.hDC = hDC
    End If
        
        
    ' get the fore colors & styles of enabled & disabled menu items
    Call cColors(isActive).GetEnabledTextColors(fx(0), Action, fColor(0), fColor(1), fColor(2))
    Call cColors(isActive).GetDisabledTextColors(fx(1), Action, fColor(3), fColor(4), fColor(5))
        
    ' if drawing a single item, get the colors for the menubar state
    ' otherwise the state is normal & no 3D boxes are needed to be drawn
    If menuItem > -1 And Action <> mcStandard Then Call cColors(isActive).GetMenuSelectionStyle((Action = mcHover), fBorder(0), fBorder(1), fx(2))
        
    If bOwnerDrawn = 2 Then
        
        ' loop thru each item & draw it
        For I = mnuFrom To mnuTo
        
            With cMBarItem(cMenuOrder(I))
            
                ' owner is drawing the complete menu item
                CID.itemData = .itemData
                CID.itemID = .ID
                CID.itemOD = odMenuItem_Complete
                CID.rcItem = .Position
                If menuItem > -1 Then   ' all items are being drawn (menubar refresh)
                    OffsetRect CID.rcItem, -.Position.Left + x, -.Position.Top + Y
                End If
                CID.itemPos = cMenuOrder(I)
                CID.itemState = Abs((.flags And MF_DISABLED) = MF_DISABLED) * 8 Or Action
               
                ' in case user does something unwise like unloading while measuring items
                If mODImplementation = 0 Then
                    Set cClient = Nothing
                    Exit Sub
                End If
                cClient.OwnerDrawMessage omDrawMenuItem, isActive + 0, VarPtr(CID), .Caption
                If menuItem < 0 Then
                    SelectClipRgn hDC, cClipRgn
                Else
                    SelectClipRgn hDC, ByVal 0&
                End If
            End With
        Next
    
    Else
        
        ' loop thru each item & draw it
        For I = mnuFrom To mnuTo
        
            With cMBarItem(cMenuOrder(I))
                
                If Len(.Caption) Then   ' a check, but shouldn't be in cMenuOrder anyway
                    mRect = .Position  ' cache current menu item position
                    
                    If menuItem > -1 Then   ' drawing single item
                        ' these are drawn to a different DC, so offset it for that DC
                        OffsetRect mRect, -mRect.Left + x, -mRect.Top + Y
                        If fx(2) = mbFilledCustom Then  ' solid backfill for item state
                            mRect.Bottom = mRect.Bottom + 1
                            mRect.Right = mRect.Right + 1
                            hOldPen = CreateSolidBrush(fBorder(0))
                            FillRect hDC, mRect, hOldPen
                            DeleteObject hOldPen
                            hOldPen = CreateSolidBrush(fBorder(1))
                            FrameRect hDC, mRect, hOldPen
                            DeleteObject hOldPen
                            mRect.Bottom = mRect.Bottom - 1
                            mRect.Right = mRect.Right - 1
                        End If
                    End If
                    
                    ' calculate the text portion and image portion of the mRect
                    tRect = mRect
                    If bOwnerDrawn = 1 Then
                        ' has an image to be drawn by owner
                        ' calculate vertical center
                        iRect.Top = (tRect.Bottom - tRect.Top - IconSize) \ 2 + tRect.Top + 1
                        ' now set the tRect & iRect
                        If (.odProp And odImgAlignLeft) = odImgAlignLeft Then ' img is right aligned
                            SetRect iRect, tRect.Left + cHoverBorderCx, iRect.Top, tRect.Left + cHoverBorderCx + IconSize, iRect.Top + IconSize
                            tRect.Left = iRect.Right + cHoverBorderCx
                        ElseIf (.odProp And odImgAlignRight) = odImgAlignRight Then
                            SetRect iRect, tRect.Right - IconSize - cHoverBorderCx, iRect.Top, tRect.Right - cHoverBorderCx, iRect.Top + IconSize
                            tRect.Right = iRect.Left - cHoverBorderCx
                            tRect.Left = mRect.Left + cHoverBorderCx
                        Else
                            SetRect iRect, 0, 0, 0, 0
                        End If
                        If iRect.Right > iRect.Left Then
                            ' prepare UDT for user
                            CID.itemData = .itemData
                            CID.itemID = .ID
                            CID.itemOD = odMenuItem_ImageOnly
                            CID.itemPos = cMenuOrder(I)
                            CID.itemState = Abs((.flags Or MF_DISABLED) = .flags) * 8 Or Action
                            CID.rcItem = iRect
                            ' in case user does something unwise like unloading while measuring items
                            If mODImplementation = 0 Then
                                Set cClient = Nothing
                                Exit Sub
                            End If
                            'BitBlt cOwnerDrawnDC, 0, 0, iRect.Right - iRect.Left + 1, iRect.Bottom - iRect.Top + 1, hDC, iRect.Left, iRect.Top, vbSrcCopy
                            ' pass message to user so the icon can be drawn
                            cClient.OwnerDrawMessage omDrawMenuItem, isActive + 0, VarPtr(CID), cMBarItem(cMenuOrder(I)).Caption
                            DTFlags = DTFlags And Not DT_CENTER
                        Else
                            DTFlags = DTFlags Or DT_CENTER
                        End If
                    Else
                        DTFlags = DTFlags Or DT_CENTER
                    End If
                
                    ' now to the caption, following handles all color & style options
                    
                    ' all the forecolors were loaded into a 6 element array,
                    ' flag which portion of that array will be used
                    If (.flags Or MF_DISABLED) = .flags Then fxID = 1 Else fxID = 0
                    
                    StyleText .Caption, tRect.Left, tRect.Top, tRect.Right, tRect.Bottom, _
                            fx(fxID), fColor(), hDC, DTFlags, , fxID * 3
                
                    ' if drawing single item, add the hover/selected 3D box if needed
                    Select Case fx(2)
                        Case mbRaisedCustom, mbSunkenCustom
                            hOldPen = SelectObject(hDC, CreatePen(0, 1, fBorder(0)))
                            MoveToEx hDC, mRect.Left, mRect.Bottom, ByVal 0&
                            LineTo hDC, mRect.Left, mRect.Top + 1
                            LineTo hDC, mRect.Right, mRect.Top + 1
                            DeleteObject SelectObject(hDC, CreatePen(0, 1, fBorder(1)))
                            LineTo hDC, mRect.Right, mRect.Bottom
                            LineTo hDC, mRect.Left - 1, mRect.Bottom
                            DeleteObject SelectObject(hDC, hOldPen)
                    End Select
                    ' wow! done with the menubar
                End If
            End With
        Next
    End If
    ' clean up
    Set cClient = Nothing
    SelectObject hDC, hFont
  
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.UpdateMenuBar
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : Draws an individual menu item in the proper tri-state
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Sub UpdateMenuBar(ByVal menuItem As Long, Optional wDC As Long, Optional isActive As Byte)

' MenuItem is the menubar item to change appearance/triState
' This parameter has 3 meanings:
'   if -2, then this routine's only purpose is to reset the menubar item to non-selected
'   if -1, same as -1 and also identifies that the next menu being displayed is the sysMenu
'   else, the menu position to change appearance

    If IsArrayEmpty(Not cMenuOrder) Then Exit Sub

    Dim tgtDC As Long, hDC As Long
    Dim mRect As RECT, mWidth As Long, mHeight As Long
    Dim hBrush As Long
    Dim gColors(0 To 1) As Long, bkStyle As eFillStyle
    ' ^^ background colors & styles of the menubar
    Dim ActionType As Long
    Dim bkgImg As StdPicture

    ' this may also be called while repainting the non-client area.
    ' If so, a DC will be provided
    If wDC Then tgtDC = wDC Else tgtDC = GetWindowDC(cTracker.hWnd)
    hDC = Canvas(False)
    If hDC = 0 Then Exit Sub

    If cTracker.LastDrawn > 0 Then
        ' did we draw something? If so reset it now
        With cMBarItem(cMenuOrder(cTracker.LastDrawn - 1)).Position
            BitBlt tgtDC, .Left, .Top, .Right - .Left + 1, .Bottom - .Top + 1, hDC, 16, 0, vbSrcCopy
        End With
        
    End If

    If menuItem > -1 Then ' actually drawing a hovered/selected item here
    
        With cMBarItem(cMenuOrder(menuItem)).Position
            
            ' now determine what action is needed: hover or selected
            If cTracker.TriState = 1 And cTracker.Active = 0 Then
                ActionType = mcHover
            Else
                If cTracker.TriState <> 0 Then ActionType = mcSelect Else ActionType = mcStandard
            End If
            
            ' calculate width/height of item (used a bit)
            mWidth = .Right - .Left + 1
            mHeight = .Bottom - .Top + 1
            
            ' if this isn't the current item in our bitmap, include it now
            If cTracker.LastDrawn - 1 <> menuItem And ActionType <> mcStandard Then
                BitBlt hDC, 16, 0, mWidth, mHeight, tgtDC, .Left, .Top, vbSrcCopy
                cTracker.LastDrawn = menuItem + 1
            End If
            
            SetRect mRect, 16, mHeight, mWidth + 16, mHeight * 2 + 1
            
            ' see if the menubar has an image background
            Set bkgImg = cColors(isActive).pIsImgBkg
            ' we are preparing a blank spot to redraw the text for this item
            
            If bkgImg Is Nothing Then
                ' the appropriate back colors for the menubar
                cColors(isActive).GetBackColors bkStyle, gColors(0), gColors(1)
                If bkStyle = bfGradientEW Or bkStyle = bfGradientNS Then
                    ' gradient fill only the portion that is the menu item
                    ' I could have created a clipping region, but decided to just grind
                    ' and make the gradient routine do the work by supplying some extra
                    ' optional variables (the last 2 parameters on the next couple lines)
                    If bkStyle = bfGradientNS Then
                        GradientFill gColors(0), gColors(1), hDC, mRect.Left, mRect.Top, mRect.Right, -mRect.Bottom, , _
                            cBar.Bottom - cBar.Top - frameOffset * 2, .Top - cBar.Top + frameOffset
                    Else
                        GradientFill gColors(0), gColors(1), hDC, mRect.Left, mRect.Top, mRect.Right, mRect.Bottom, , _
                            cBar.Right - cBar.Left - frameOffset * 2, .Left - cBar.Left + frameOffset
                    End If
                ElseIf bkStyle = bfSolid Then
                    ' simple solid color background
                    hBrush = CreateSolidBrush(gColors(0))
                    FillRect hDC, mRect, hBrush
                    DeleteObject hBrush
                End If
                
            Else
                ' ok we have an image background. When used, I convert full menubar
                ' background image to a stdPicture. I then simply use its Render method
                ' to draw just the portion of the image that I need
                
                ' If you want to learn more about Render, take a look at some other
                ' posts I put on PSC: Render & Animated GIF Control
                bkgImg.Render hDC + 0, 16&, mHeight + 0, mWidth + 0, mHeight + 0, _
                    RevertHimetrix(.Left - cBar.Left + 1, True), _
                    RevertHimetrix(cBar.Bottom - .Top - 1, False), _
                    RevertHimetrix(mWidth, True), -RevertHimetrix(mHeight, False), ByVal 0&
            End If
            
            ' draw the menu item in the requested colors and/or style
            DrawMenuItem ActionType, hDC, 16, mHeight + 0, isActive, menuItem
            ' clip our region & draw it to the visible window
            SelectClipRgn tgtDC, cClipRgn
            BitBlt tgtDC, .Left, .Top, mWidth, mHeight, hDC, 16, mHeight + 0, vbSrcCopy
            SelectClipRgn tgtDC, ByVal 0&
            
        End With
    
    Else
        ' -1 is system menu, -2 is simply to replace previous drawn item
        If menuItem = -1 Then cTracker.LastDrawn = 0
    
    End If
    ' release window dc if used
    If wDC = 0 Then ReleaseDC cTracker.hWnd, tgtDC

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.GetMenuItemCaption
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : Returns specific information about a menu item
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub GetMenuItemCaption(ByVal itemPos As Long)
' simply get a specific menu item caption

    Dim MIS As MENUITEMINFO_String
    ' annotation needed:  menu captions max length of 128 characters
    MIS.dwTypeData = String$(128, 0)
    MIS.fMask = MIIM_ID Or MIIM_STATE Or &H10 ' [MIIM_Type]
    MIS.cbSize = LenB(MIS)
    MIS.cch = 127
    GetMenuItemInfo_String cMenu, itemPos, True, MIS
    
    With cMBarItem(itemPos)
        .ID = MIS.wID
        .flags = MIS.fState
        'if this fails then the menu is either owner-drawn or it is a SC_Min/Max/Close button
        If MIS.cch > 0 Then
            .Caption = Left$(MIS.dwTypeData, MIS.cch)
        Else
            .Caption = ""
        End If
    End With

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.MenuItemCount
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : simply returns the number of menu items either in a submenu or on the menubar
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Function MenuItemCount(hSubMenu As Long) As Long

If cMenu Then
    If hSubMenu Then
        MenuItemCount = GetMenuItemCount(hSubMenu)
    Else
        MenuItemCount = GetMenuItemCount(cMenu)
    End If
End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.SetMouseAction
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : process messages related to just the menubar
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function SetMouseAction(wParam As Long, lParam As Long, msgSource As Long) As Boolean
' This function is called from both the message filter hook established whenever
' the menubar is activated via mouse clicks or keyboard action.
' Only 2 messages trapped from the menu filter; the other three come from the subclasser

Dim mnuMsg As MSG
Dim lValue As Long

' copy the message so we can look at it
Select Case msgSource

Case -1: ' result of msgFilter hook
    CopyMemory mnuMsg, ByVal lParam, Len(mnuMsg)
    ' only trapping wm_mousemove & wm_menuselect
    
Case WM_NCHITTEST   ' related to custom timer triggering
'^^ :) probably should use WM_TIMER; but didn't want to declare a variable for only that reason
    mnuMsg.Message = msgSource
    
Case WM_NCMOUSEMOVE, WM_NCLBUTTONDOWN
    ' ^^ the passed messages from the filter do not contain NC button actions
    ' all actions are client button actions (result of TrackPopupMenuEx capturing
    ' the mouse). Therefore, we let them fall thru which will trigger the
    ' appropriate NC button action that will come back here. Intercepting &
    ' converting the client button/mousemove actions requires more code than
    ' simply letting the natural process recurse back to this point
    mnuMsg.Message = msgSource
    mnuMsg.wParam = wParam
    mnuMsg.lParam = lParam
    
End Select

    Select Case mnuMsg.Message ' these are common WM_flags
    
    Case WM_NCHITTEST
    ' timer function call back to check if the menubar region has been exited
        Dim wRect As RECT
        If cTracker.Active = 0 Then ' menubar is not in loop
            ' remove the previous item that was highlighted
            If cTracker.LastDrawn > 0 Then
                ' but only if we are outside of the menubar
                GetWindowRect cTracker.hWnd, wRect
                GetCursorPos mnuMsg.Pt
                If PtInRect(wRect, mnuMsg.Pt.x, mnuMsg.Pt.Y) = 0 Then
                    CreateTimer False
                    mnuMsg.lParam = cTracker.LastDrawn
                    pTracking(-1) = 0
                    UpdateMenuBar mnuMsg.lParam - 1, , cActive
                    
                Else
                    ' offset the points to our menubar location
                    mnuMsg.Pt.x = mnuMsg.Pt.x - wRect.Left
                    mnuMsg.Pt.Y = mnuMsg.Pt.Y - wRect.Top
                    ' run the test
                    If PtInRect(cBar, mnuMsg.Pt.x, mnuMsg.Pt.Y) = 0 Then
                        CreateTimer False   ' kill the timer & reset flags
                        UpdateMenuBar -2
                        pTracking(-1) = 0
                    End If
                End If
            Else
                CreateTimer False   ' kill timer
            End If
        End If
        
    Case WM_MOUSEMOVE
        ' sysMenu will be at -1; don't intercept any messages for that menu
            If cTracker.Position > -1 Then
                ' ensure the point is somewhere in our window
                If (WindowFromPoint(mnuMsg.Pt.x, mnuMsg.Pt.Y) = cTracker.hWnd) Then
                    ' each popup usually is followed by 1 or more mouse moves
                    ' Since we cached the current coords, we can prevent misfiring
                    ' when cursor is over one item, but user keyboards to another
                    ' and when it opens, the mouse move message would trigger the
                    ' routines to go back to the item where the cursor is. Annoying.
                    If cTracker.CursorPt.x <> mnuMsg.Pt.x Or cTracker.CursorPt.Y <> mnuMsg.Pt.Y Then
                        If HitTest(mnuMsg.Pt.x, mnuMsg.Pt.Y, True) = HTMenu_Custom Then
                            cTracker.CursorPt = mnuMsg.Pt
                            If (cTracker.flags Or kbNonMnuItem) <> cTracker.flags Then
                                ShowMenu cTracker.HitTest - 1, cTracker.TriState, False
                            End If
                        End If
                    End If
                End If
            Else ' not in our window; maybe in the submenu or another window/child
                cTracker.CursorPt = mnuMsg.Pt
            End If
        
    Case WM_NCMOUSEMOVE
        ' not called from the message hook, but is routed from the subclasser when
        ' the mouse is over a submenu-less menu item
        If (cTracker.flags Or kbNonMnuItem) <> cTracker.flags Then
            If cTracker.HitTest <> cTracker.LastDrawn Then
                ' highlight the item that mouse is over
                If cTracker.Active = 0 Then
                    If cTracker.HitTest > 0 Then
                        cTracker.TriState = 1
                        UpdateMenuBar cTracker.HitTest - 1, , CByte(Abs(wParam <> 0))
                        CreateTimer True
                    End If
                Else    ' menubar is active, select different menu item
                    mnuMsg.Pt.x = LoWord(lParam)
                    mnuMsg.Pt.Y = HiWord(lParam)
                    If cTracker.CursorPt.x <> mnuMsg.Pt.x Or cTracker.CursorPt.Y <> mnuMsg.Pt.Y Then
                        cTracker.CursorPt = mnuMsg.Pt
                        ShowMenu cTracker.HitTest - 1, cTracker.TriState, False
                    End If
                End If
            End If
        End If
        
    Case WM_NCLBUTTONDOWN
        ' not called from the message hook, but is routed from the subclasser
        If mnuMsg.wParam = HTMenu_Custom Then
            If cTracker.TriState = 2 Then   '
                ' test to see if this a submenu-less menubar item
                If Not PostMenuCommand(cTracker.HitTest - 1, False) Then
                    ' if the user clicked the menuitem that was already clicked
                    ' then exit the loop
                    If cTracker.HitTest - 1 = cTracker.Position Then
                        UpdateMenuBar -2
                        pTracking(mbActiveStatus) = 0
                    Else    ' otherwise, show the new menu item
                        cTracker.CursorPt = mnuMsg.Pt
                        ShowMenu cTracker.HitTest - 1, 2, False
                    End If
                End If
            Else
                ' menubar not in loop, user clicked on a menu item -- enter loop
                If cTracker.HitTest > 0 Then
                    If Not PostMenuCommand(cTracker.HitTest - 1, False) Then
                        cTracker.CursorPt = mnuMsg.Pt
                        ShowMenu cTracker.HitTest - 1, 2, False
                    End If
                End If
            End If
        Else    ' non-menubar region clicked; exit loop if needed
            If cTracker.Active > 0 Then
                UpdateMenuBar -2
                pTracking(mbActiveStatus) = 0
            Else
                pTracking(-1) = 0
            End If
            
        End If
        
    Case WM_MENUSELECT
        lValue = HiWord(mnuMsg.wParam)
        If lValue <> -1 Or mnuMsg.lParam <> 0 Then
            ' set flag to show this panel can open another submenu panel
            If (lValue Or MF_POPUP) = lValue Then
                cTracker.flags = cTracker.flags Or kbCanExpand ' don't trap right arrow
            Else
                cTracker.flags = cTracker.flags And Not kbCanExpand ' trap right arrow
            End If
            ' set flag to show this panel cannot open another submenu
            If mnuMsg.lParam = cTracker.ID Then
                cTracker.flags = cTracker.flags And Not kbCanCollapse ' trap left arrow
            Else
                cTracker.flags = cTracker.flags Or kbCanCollapse ' don't trap left arrow
            End If
        Else
            ' menu closed; no action required
        End If
        
    Case Else
    
    End Select

End Function
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.SetKeyBdAction
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : process keyboard messages from the Keyboard Hook
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function SetKeyBdAction(wKeyCode As Long, lParam As Long) As Boolean

' The actions below are coded from experimenting with as many key combintations
' as I could think of. I may not have all of them accounted for; but...

Dim lMenuItem As Long
Dim bShowingPanel As Boolean, bHasSubMenu As Boolean
Dim bCanCollapse As Boolean, bCanExpand As Boolean

If Not ((lParam And &H80000000) = &H80000000) Then ' Key Press

    ' set some key attributes to be used below
    bHasSubMenu = ((cTracker.flags Or kbHasSubMenu) = cTracker.flags)
    bShowingPanel = (((cTracker.TriState = 2) And bHasSubMenu))
    bCanCollapse = ((cTracker.flags Or kbCanCollapse) = cTracker.flags)
    bCanExpand = ((cTracker.flags Or kbCanExpand) = cTracker.flags)
    
    Select Case wKeyCode
    
    Case vbKeyF10, vbKeyMenu   ' second alt keydown; end menu loop
        SendMessage cTracker.hWnd, WM_CANCELMODE, 0, ByVal 0&
        UpdateMenuBar -2
        pTracking(mbActiveStatus) = 0
        
    Case vbKeyTab, vbKeyBack
        ' exit out of menubar unless showing a submenu
        If Not bShowingPanel Then
            UpdateMenuBar -2
            pTracking(mbActiveStatus) = 0
            Exit Function   ' allow key to fall thru
        End If
        
    Case vbKeyEscape
        If bCanCollapse Then ' not much to do here
            Exit Function ' allow ESC to fall thru
        Else
            ' cannot collapse further
            If cTracker.TriState = 1 Then
                ' final level before exiting the loop, refresh menubar & clean up
                UpdateMenuBar -2
                pTracking(mbActiveStatus) = 0
            Else
                If (cTracker.flags Or kbHasSubMenu) = cTracker.flags Then
                    ' if this menu has submenus, close them
                    SendMessage cTracker.hWnd, WM_CANCELMODE, 0, 0
                    ' revert to TriState of 1
                    cTracker.TriState = 1
                    UpdateMenuBar cTracker.Position, , 1
                    ' remove any flags indicating menu item can accept right or left arrows
                    cTracker.flags = (cTracker.flags And Not kbCanCollapse) And Not kbCanExpand
                Else
                    ' submenu-less menuitem; simply exit menu loop
                    UpdateMenuBar -2
                    pTracking(mbActiveStatus) = 0
                End If
            End If
        End If
    
    Case vbKeyUp, vbKeyDown
        If bShowingPanel Then
            Exit Function       ' allow key to fall thru
        Else
            ' display submenu if appropriate
            If bHasSubMenu Or cTracker.Position < 0 Then
                ShowMenu cTracker.Position, 2, True
            End If
        End If
        
    Case vbKeyReturn
        If bShowingPanel Then
            ' user hit enter while submenu displayed or on sysmenu
            UpdateMenuBar -2
            cTracker.flags = cTracker.flags Or SC_KEYMENU Or SC_MOUSEMENU
            Exit Function      ' allow the ENTER to fall thru
        Else
            ' if this has a submenu or the sysMenu displayed then show the submenu
            If bHasSubMenu Or cTracker.Position < 0 Then
                ShowMenu cTracker.Position, 2, True
            Else
                'submenu-less menubar item
                PostMenuCommand cTracker.Position, True
            End If
        End If

    Case vbKeySpace     ' if a sysicon then show it
        ' if this menubar item doesn't have submenus then show system menu
        If Not bHasSubMenu Then
            lMenuItem = -1
            ' if the menubar is in highlight mode, show the system menu
        ElseIf bShowingPanel Then
            Exit Function
        ElseIf Not bShowingPanel Then
            lMenuItem = -1
        End If
        If lMenuItem < 0 And cSysMenu <> 0 Then
            ShowMenu lMenuItem, 2, True
        Else
            ' if not, exit the menubar loop
            UpdateMenuBar -2
            pTracking(mbActiveStatus) = 0
        End If
        ' either way, eat the keystroke
        
    Case vbKeyLeft
        ' if this menu bar can collapse, let it go
        If bCanCollapse Then
            Exit Function
        Else
            If Not IsArrayEmpty(Not cMenuOrder) Then
                ' calculate next menubar item to show & show it
                lMenuItem = cTracker.Position
                If lMenuItem < 0 Then
                    'lMenuItem = UBound(cMBarItem)
                    lMenuItem = cMenuOrder(UBound(cMenuOrder))
                Else
                    lMenuItem = cTracker.Position - 1
                End If
                ' test for case of no system menu
                If lMenuItem < 0 And cSysMenu = 0 Then
                    'lMenuItem = UBound(cMBarItem)
                    lMenuItem = cMenuOrder(UBound(cMenuOrder))
                End If
                If lMenuItem <> cTracker.Position Then
                    ShowMenu lMenuItem, cTracker.TriState, True
                End If
            End If
        End If
        
    Case vbKeyRight ' navigating at menubar
        ' if this menu panel can expand, then let it go
        If bCanExpand Then
            Exit Function
        Else
            If Not IsArrayEmpty(Not cMenuOrder) Then
                ' calculate next menubar item to show & show it
                lMenuItem = cTracker.Position
                'If lMenuItem = UBound(cMBarItem) Then
                If lMenuItem = cMenuOrder(UBound(cMenuOrder)) Then
                    lMenuItem = -1
                Else
                    lMenuItem = lMenuItem + 1
                End If
                ' test for case of no system menu
                If lMenuItem < 0 And cSysMenu = 0 Then
                    lMenuItem = 0
                End If
                If lMenuItem <> cTracker.Position Then
                    ShowMenu lMenuItem, cTracker.TriState, True
                End If
            End If
        End If
    
    Case Else ' some other key was pressed
        ' if this menubar item has no submneus, look for next match
        If Not bHasSubMenu Then
            lMenuItem = 1
            ' if menubar is in highlight mode, look for next match
        ElseIf bShowingPanel = False Then
            lMenuItem = 1
        End If
        If lMenuItem Then
            ' see if another menubar item matches
            If Not FindMenuItem(wKeyCode) Then
                ' if not, exit the menubar loop
                UpdateMenuBar -2
                pTracking(mbActiveStatus) = 0
            End If
        Else
            ' let the key go thru
            Exit Function
        End If
    End Select
    
End If

SetKeyBdAction = True

End Function
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.FindMenuItem
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : Called internally & externally to locate menu item via keystroke
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function FindMenuItem(ByVal KeyCode As Long) As Boolean
    
    Dim lMenuItem As Long, lItem As Long
    Dim sCriteria As String, sKey As String
    Dim bInDupLoop As Boolean
    
    Select Case KeyCode
    
    Case -1 ' system menu
    
        If cSysMenu <> 0 Then lMenuItem = -1
        
    Case -2 ' alt + no key
    
        If IsArrayEmpty(Not cMenuOrder) Then Exit Function
        If (cOptions Or mbHideMenu) = cOptions Then Exit Function
        lMenuItem = 1        ' set to first menu item (enabled or not)
        bInDupLoop = True       ' prevents showing submenu if it has one
        UpdateMenuBar -2        ' refresh menubar if needed
        pTracking(-1) = 0       ' reset key flags
        cTracker.HitTest = 1    ' prevents mouse from erasing selection if moved over menubar
                                ' in an area not occupied by a menuitem
    Case Else
    
        If IsArrayEmpty(Not cMenuOrder) Then Exit Function
        If cMenu = 0 Then Exit Function
        
        sKey = Chr$(KeyCode)
        'Debug.Print "looking for " & sKey, "cur pos: " & cTracker.Position
        
        For lItem = 0 To UBound(cMenuOrder)
            If (cMBarItem(cMenuOrder(lItem)).flags Or MF_DISABLED) = cMBarItem(cMenuOrder(lItem)).flags Then
                ' don't allow disabled items to be accessed via accelerator keys
                sCriteria = sCriteria & Chr$(0)
            Else
                lMenuItem = InStrRev(cMBarItem(cMenuOrder(lItem)).Caption, "&")
                If lMenuItem Then
                    'identify those with ampersands as upper case
                    sCriteria = sCriteria & UCase(Mid$(cMBarItem(cMenuOrder(lItem)).Caption, lMenuItem + 1, 1))
                Else ' and those without as lowercase
                    sCriteria = sCriteria & LCase(Left$(cMBarItem(cMenuOrder(lItem)).Caption, 1))
                End If
            End If
        Next
        
        ' see if the requested key exists in menubar (with & w/o &)
        lMenuItem = InStr(1, sCriteria, sKey, vbTextCompare)
        ' if it does exist, is it duplicated (with & w/o &)
        If lMenuItem Then bInDupLoop = (InStr(lMenuItem + 1, sCriteria, sKey, vbTextCompare) > 0)
        
        ' but depending on the menubar state we will differentiate a bit
        If (cTracker.TriState = 1) Then ' hilighting only
            If bInDupLoop Then
                ' multiple menubar items exist with the same accelerator
                ' Here's where I adapt my routines from standard routines & R is pressed
                ' Example: You have Run, Que&ry, & P&roject
                '   VB apps will only toggle btwn Query & Project 'cause they have proper accelerators
                '   The only way to get to "Run" is to use arrows. I will include the non-proper in the loop
            
                ' when in a dup key situation, we want the next key position
                lMenuItem = InStr(cTracker.Position + 2, sCriteria, sKey, vbTextCompare)
                ' when no key to the right, select first option
                If lMenuItem = 0 Then lMenuItem = InStr(1, sCriteria, sKey, vbTextCompare)
            End If
        Else        ' the Alt+key was pressed
        
            If bInDupLoop Then
                ' when more than one key exists, but only one has the Accelerator
                ' then we will select that one. If 2 or more have Accelerators then
                ' we push the 1st one
                lMenuItem = InStr(1, sCriteria, UCase(sKey))
                If lMenuItem Then bInDupLoop = (InStr(lMenuItem + 1, sCriteria, UCase(sKey)))
                
            End If
        End If
        
    End Select
    
    If lMenuItem Then ' we have something that fits the key being pressed
        ' adjust for non-sysMenu item
        If lMenuItem > 0 Then lMenuItem = lMenuItem - 1
        If Not bInDupLoop Then
            ' unique key, not duplicated in the menubar
            If lMenuItem > -1 Then ' exlcude sysmenu from this check
                If GetSubMenu(cMenu, lMenuItem) = 0 Then
                    ' if submenu-less item, send click event
                    FindMenuItem = PostMenuCommand(lMenuItem, True)
                    Exit Function
                End If
            End If
        End If
        ' show the requested menu item
        ShowMenu lMenuItem, 2 + bInDupLoop, True
        FindMenuItem = True
    End If
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.ShowMenu
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : Displays the selected menuitem if applicable
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Sub ShowMenu(ByVal menuItem As Long, ByVal TriState As Long, ByVal viaKeyboard As Boolean)
    
    ' ensure the requested menu item exists
    If menuItem > -1 Then
        If IsArrayEmpty(Not cMenuOrder) Then Exit Sub
        If (cOptions Or mbHideMenu) = cOptions Then Exit Sub
    Else
        If cSysMenu = 0 Then Exit Sub
    End If
    ' next test: if the same menu item as the one being displayed is selected
    ' and the TriState value hasn't changed; abort
    If menuItem = cTracker.Position Then
        If TriState = cTracker.TriState Then Exit Sub
    End If
    
    Dim bCanDisplay As Boolean, tbAlignment As eBarAlignment
    Dim x As Long, Y As Long, tRect As RECT
    
    If menuItem < 0 Then    'system menu

        cTracker.ID = GetSystemMenu(cTracker.hWnd, 0) ' cSysMenu
        cTracker.flags = kbHasSubMenu ' flag indicating can drop down (tweaked; forces showing if needed)
        ' get the titlebar rectangle and its alignment
        tbAlignment = sharedProperty([_tbarAlignment])
        x = sharedProperty([_titleBarRect])
        If x <> 0 Then CopyMemory tRect, ByVal x, &H10
        
        ' position the menu in relation to the titlebar, menubar or client rectangle
        If IsRectEmpty(tRect) Then      ' no titlebar?
            If (cOptions And mbHideMenu) = mbHideMenu Or IsRectEmpty(cBar) <> 0 Then
                ' hidden menubar or no menubar
                x = 1   ' use 1,1 as that will be the client coords +/- borders
                Y = 1
            Else
                x = cBar.Left   ' use the menubar's top/left
                Y = cBar.Top
            End If
        Else
            ' expected. We have a titlebar, place sysmenu popup near the sysmenu icon
            ' set flags to be used when this message is posted & received
            If tbAlignment < barAlignBottom Then ' vertical bar
                If tbAlignment = barAlignRight Then
                    x = tRect.Left - 1
                    Y = tRect.Top
                    cTracker.flags = cTracker.flags Or TPM_RIGHTALIGN
                Else
                    x = tRect.Right + 1
                    Y = tRect.Bottom
                    cTracker.flags = cTracker.flags Or TPM_BOTTOMALIGN
                End If
            Else
                If tbAlignment = barAlignBottom Then
                    x = tRect.Left
                    Y = tRect.Top
                    cTracker.flags = cTracker.flags Or TPM_BOTTOMALIGN
                Else    ' top aligned, default flags: top & left aligned
                    x = tRect.Left
                    Y = tRect.Bottom
                End If
            End If
        End If
    Else    ' some other menu item from the menubar
        With cMBarItem(cMenuOrder(menuItem))
            cTracker.ID = .ID
            x = .Position.Left
            Y = .Position.Bottom + 1
        End With
        cTracker.flags = 0
        ' don't show disabled menu items
        If (cMBarItem(cMenuOrder(menuItem)).flags Or MF_DISABLED) <> cMBarItem(cMenuOrder(menuItem)).flags Then
            ' can this menuitem be dropped? Does it have a submenu
            If GetSubMenu(cMenu, menuItem) Then cTracker.flags = kbHasSubMenu
        End If
        
    End If
    ' set some other tracking values
    cTracker.Position = menuItem
    cTracker.TriState = TriState
    
    ' determine whether or not the menu has submenus
    bCanDisplay = ((TriState = 2 And (cTracker.flags Or kbHasSubMenu) = cTracker.flags))
    
    ' remove any previously displayed menu now
    SendMessage cTracker.hWnd, WM_CANCELMODE, 0, 0
    ' cache the cursor location if this menu selected via keyboard
    If viaKeyboard Then
        GetCursorPos cTracker.CursorPt
        ' set type of selection (keyboard or mouse)
        cTracker.flags = cTracker.flags Or SC_KEYMENU
    End If

    ' begin hooking the keyboard and menu messages
    If cTracker.Active = 0 Then pTracking(mbActiveStatus) = 1
    
    ' send custom message to display the menu
    If bCanDisplay Then PostMessage cTracker.hWnd, WM_LVPopup, x, ByVal Y
    
    ' draw the menubar item
    UpdateMenuBar menuItem, , 1
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.PostMenuCommand
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : sends user a menu item click event
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Function PostMenuCommand(ByVal menuItem As Integer, viaKeyboard As Boolean) As Boolean

Dim bReset As Boolean

If (cTracker.flags Or kbNonMnuItem) = cTracker.flags And viaKeyboard = False Then
    ' when the mouse is clicked in the far right sections of a menubar where no
    ' menu item exists, the app may think the currently selected menu item is
    ' to be fired. This flag value prevents that from happening
    If menuItem < 0 Then
        bReset = True   ' if system menu was displayed then it closes the menu loop
    Else
        PostMenuCommand = True
        Exit Function
    End If

Else
    If menuItem > -1 Then
        If (cMBarItem(cMenuOrder(menuItem)).flags Or MF_DISABLED) = cMBarItem(cMenuOrder(menuItem)).flags Then
            ' don't send click event if the menu item is disabled
            bReset = True
        Else
            If GetSubMenu(cMenu, menuItem) = 0 Then
                ' no submenus send a command (submenu-less)
                bReset = True
                PostMessage cTracker.hWnd, WM_COMMAND, MakeLong(cMBarItem(cMenuOrder(menuItem)).ID + 0, 0), ByVal 0&
            End If
        End If
    Else
        '
    End If
End If

If bReset Then
    UpdateMenuBar -2
    pTracking(mbActiveStatus) = 0
    PostMenuCommand = True
End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.HitTest
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : Hit test functionality for the menubar
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function HitTest(ByVal x As Long, ByVal Y As Long, Optional bInternal As Boolean) As Long
    
' Called by the WM_NCHITTEST message action &
' also may be called by the SetMouseAction routine
    
    Dim wRect As RECT
    Dim lItem As Long
        
    If pHasMenuBar Then
    
        If bInternal Then
            ' called from SetMouseAction, client coords vs window coords
            GetWindowRect cTracker.hWnd, wRect
            ' offset the points to align with our menubar coords
            x = x - wRect.Left
            Y = Y - wRect.Top
        End If
        
        If PtInRect(cBar, x, Y) Then
            ' if the mouse is on our menubar, is it in a menu item?
            For lItem = UBound(cMenuOrder) To 0 Step -1
                If PtInRect(cMBarItem(cMenuOrder(lItem)).Position, x, Y) Then
                    cTracker.HitTest = lItem + 1
                    Exit For
                End If
            Next
            ' set flag to indicate that hit test includes a menuitem or not.
            ' Not everything in the menubar is a menuitem. There can be
            ' blank space also. Want to differentiate btwn the two
            If lItem < 0 Then
                cTracker.flags = cTracker.flags Or kbNonMnuItem
            Else
                cTracker.flags = cTracker.flags And Not kbNonMnuItem
            End If
            
            ' return result
            HitTest = HTMenu_Custom
            
        End If
    End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.pTracking
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : exposes different tracker properties that can be set outside this class
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Property Get pTracking(Flag As eTrackingFlags) As Long
    Select Case Flag
    Case mbActiveStatus: pTracking = cTracker.Active
    Case mbHwnd: pTracking = cTracker.hWnd
    Case mbPosition: pTracking = cTracker.Position
    Case mbID: pTracking = cTracker.ID
    Case mbLastDrawn: pTracking = cTracker.LastDrawn
    Case mbTriState: pTracking = cTracker.TriState
    Case mbFlags: pTracking = cTracker.flags
    Case mbHitTest: pTracking = cTracker.HitTest
    Case Else ' error
        pTracking = cTracker.Active
        Stop
    End Select
End Property
Friend Property Let pTracking(Flag As eTrackingFlags, fValue As Long)
    Select Case Flag
    Case -1: ' reset
        cTracker.Active = 0
        cTracker.LastDrawn = 0
        cTracker.flags = 0
        cTracker.TriState = 0
        cTracker.HitTest = 0
        cTracker.Position = 0
        cTracker.CursorPt.x = -100
        cTracker.CursorPt.Y = -100
    Case mbActiveStatus
        cTracker.Active = fValue
        SetInputHook CBool(fValue), ObjPtr(Me)
        If fValue = 0 Then pTracking(-1) = 0
    Case mbID: cTracker.ID = fValue
    Case mbPosition: cTracker.Position = fValue
    Case mbTriState: cTracker.TriState = fValue
    Case mbLastDrawn: cTracker.LastDrawn = fValue
    Case mbHwnd: cTracker.hWnd = fValue
    Case mbFlags: cTracker.flags = fValue
    Case mbHitTest: cTracker.HitTest = fValue
    Case Else
     Stop
    End Select
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.CreateTimer
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : Establishes a timer that firest 8x a second
' Comments  : When the mouse is over a menubar item & users scrolls so fast that
'               a NChittest isn't sent by the system, there is really no way of
'               knowing the mouse left the menubar & the menu item remains highlighted
'               The timer is created when a menu item is highlighted & will check
'               if the menu item should be reset. Timer can delete itself
'---------------------------------------------------------------------------------------
'
Private Sub CreateTimer(bNew As Boolean)

If cTracker.TimerOn Then
    If Not bNew Then    ' killing the timer
        SetStopTimer False, cTracker.hWnd, ObjPtr(Me), 0
        cTracker.TimerOn = 0
    End If
Else
    If bNew Then        ' creating the timer
        SetStopTimer True, cTracker.hWnd, ObjPtr(Me), 150
        cTracker.TimerOn = 1
    End If
End If
        
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.pMenuHandle
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : sets/returns the menu handle (may no longer be used?)
' Comments  : TODO item
'---------------------------------------------------------------------------------------
'
Friend Property Get pMenuHandle() As Long
    pMenuHandle = cMenu
End Property
Friend Property Let pMenuHandle(hWnd As Long)
    cMenu = GetMenu(hWnd)
    cSysMenu = GetSystemMenu(hWnd, 0)
    cTracker.hWnd = hWnd
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.pSysMenuHandle
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : returns the system menu handle (may no longer be used?)
' Comments  : TODO item
'---------------------------------------------------------------------------------------
'
Friend Property Get pSysMenuHandle() As Long
    pSysMenuHandle = cSysMenu
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.pMainImplementation
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : called by main class to pass its implementation pointer
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Property Let pMainImplementation(hWnd As Long, lImpObj As Long)
    mMainImplentation = lImpObj
    pMenuHandle = hWnd
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.sharedProperty
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : A way to get other class properties without having to
'             worry about caching them here & keeping them updated
'---------------------------------------------------------------------------------------
'
Private Property Get sharedProperty(propID As eProperties) As Long

    If mMainImplentation = 0 Then Exit Property
    
    Dim cImp As iImpProps
        GetObjectFromPointer mMainImplentation, cImp
        sharedProperty = cImp.GetPropValue(propID)
        Set cImp = Nothing

End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.RefreshWindow
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Forces window to repaint
' Comments  : Called when some key element changes menubar size or shape
'---------------------------------------------------------------------------------------
'
Private Sub RefreshWindow()
    If cTracker.hWnd Then
        If cAutoRedraw Then ForceRefresh cTracker.hWnd
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.pFont
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : returns a copy of the menubar font vs the actual font handle
'---------------------------------------------------------------------------------------
'
Friend Property Get pFont() As Long
    pFont = CopyFont(cFont)
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.Location
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : returns a pointer to the menubar rectangle, if any
'---------------------------------------------------------------------------------------
'
Friend Property Get Location() As Long
    If (cOptions And mbHideMenu) = 0 Then
        If IsRectEmpty(cBar) = 0 Then Location = VarPtr(cBar)
    End If
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.OwnerDrawn
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : accepts & forwards any owner drawn request sent to the main class
'---------------------------------------------------------------------------------------
'
Friend Sub OwnerDrawn(cImplementation As Long, odItems As eOwnerDrawn, isClosing As Boolean)

    If isClosing Then
        mNotifyUser = 0
        mODImplementation = 0
    Else
        Dim I As Integer
        If Not IsArrayEmpty(Not cMBarItem) Then
            ' clear any ownerdrawn properties from the menu item collection
            For I = 0 To UBound(cMBarItem)
                cMBarItem(I).odProp = 0
            Next
        End If
        mNotifyUser = 0
        For I = 1 To 4  ' check the 3 menu bar ownerdrawn flags
            If (odItems And I) = I Then mNotifyUser = mNotifyUser Or I
            I = I * 2 - 1
        Next
        If mNotifyUser = 0 Then
            mODImplementation = 0
            CreateTimer False
        Else
            mODImplementation = cImplementation      ' update the implementation pointer
        End If
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.BorderWidth/BorderHeight Property
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : sets optional menubar border sizes
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
' Procedure : clsMenubar.Class_Initialize, Class_Terminate
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Initialization & clean up code
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set cColors(1) = New clsMenuBarColors
    Set cColors(0) = New clsMenuBarColors
    cColors(1).pClassUsage = True
    ' initialize the class's menubar font
    Set Font = Nothing
End Sub

Private Sub Class_Terminate()
    ' clean up memory
    CreateTimer False
    If cTracker.Active Then pTracking(mbActiveStatus) = 0
    If cFont Then DeleteObject cFont
    Set cColors(1) = Nothing
    Set cColors(0) = Nothing
    If cClipRgn Then DeleteObject cClipRgn
End Sub

'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.Read/WriteProperites
' DateTime  : 9/26/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve custom properites
' Comments  : Used by Export/ImportCustomSettings
'---------------------------------------------------------------------------------------
'
Private Sub iPropertyBag_ReadProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)
    
    Dim iBag As iPropertyBag
    Dim sProps() As String
    
    cOptions = propBag.ReadProperty("MB.Options", 0)
    mNotifyUser = propBag.ReadProperty("MB.OD", 0)
    Set Me.Font = propBag.ReadProperty("MB.Font", Nothing)
    sProps = Split(propBag.ReadProperty("MB.Borders", "0,0"), ",")
    cBorderWidthCx = CByte(sProps(0))
    cBorderWidthCy = CByte(sProps(1))
    
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
        Me.Inactive.GetImageBackground Nothing, iStyle, iGrayScaled
        ' Get the active class's image
        Me.Active.GetImageBackground iPic
        ' now reassign the inactive class's image properties & image
        Me.Inactive.SetImageBackground iPic, iStyle, iGrayScaled
    End If
    Set iBag = Nothing
    
End Sub

Private Sub iPropertyBag_WriteProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim mFont As StdFont
    Dim sysNCM As NONCLIENTMETRICS
    Dim NCM As LOGFONT
    Dim iBag As iPropertyBag
    Dim sharedImgPtr As Long
    
    ' don't cache the font if it is the system font
    sysNCM.cbSize = Len(sysNCM)
    ' this will return the system font info
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, sysNCM, 0
    
    GetGDIObject cFont, Len(NCM), NCM
    With sysNCM.lfMenuFont
        If .lfFaceName = NCM.lfFaceName Then
            If .lfHeight = NCM.lfHeight Then
                If .lfWeight = NCM.lfWeight Then
                    If .lfWidth = NCM.lfWidth Then
                        If .lfItalic = NCM.lfItalic Then
                            If .lfUnderline = NCM.lfUnderline Then
                                If .lfStrikeOut = NCM.lfStrikeOut Then
                                    If .lfPitchAndFamily = NCM.lfPitchAndFamily Then .lfHeight = 0
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
    If sysNCM.lfMenuFont.lfHeight <> 0 Then Set mFont = HandleToFont(cFont)
    propBag.WriteProperty "MB.Font", mFont, Nothing
    propBag.WriteProperty "MB.Options", cOptions, 0
    propBag.WriteProperty "MB.OD", mNotifyUser, 0
    propBag.WriteProperty "MB.Borders", CStr(cBorderWidthCx) & "," & CStr(cBorderWidthCy), "0,0"
    
    
    Set iBag = Me.Active
    iBag.WriteProperties propBag, Limits, sharedImgPtr
    Set iBag = Me.Inactive
    iBag.WriteProperties propBag, Limits, sharedImgPtr
    Set iBag = Nothing

End Sub
'---------------------------------------------------------------------------------------
' Procedure : iTimer_TimerEvent
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : timer callback function
'---------------------------------------------------------------------------------------
'
Private Sub iTimer_TimerEvent(ByVal hWnd As Long, ByVal TimerID As Long)
    SetMouseAction 0, 0, WM_NCHITTEST
End Sub

