VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SendMessageTimeout Lib "user32.dll" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, ByRef lpdwResult As Long) As Long

Private Type MENUITEMINFO        ' used to retrieve/store menu items
     cbSize As Long              ' this structure is used with all O/S
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As Long 'String
     cch As Long
End Type

Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID    As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const SC_MOVE_Clone         As Long = &H7AA ' custom
Private Const SC_SIZE_Clone         As Long = &H7AB ' custom
Private Const SC_CLOSE_Clone        As Long = &H7AC ' custom
Private Const SC_MINIMIZE_Clone     As Long = &H7AD ' custom
Private Const SC_RESTORE_Clone      As Long = &H7AE ' custom
Private Const SC_MAXIMIZE_Clone     As Long = &H7AF ' custom

Private hSysMenu As Long        ' the window's system menu
Private cHwnd As Long           ' the widows's hWnd

Private Type tButtonInfo
    Loc As RECT         'position of the button relative to window's 0,0 coordinates
    Status As Integer   '0=normal up state,2=mouse over,4=down state
    Enabled As Byte     '0=enabled, 1=disabled/not applicable, 2=disabled by user, 4=maximized
End Type
Private Type tCustomButton
    ID As Long          'user-supplied button ID
    tTipClass As clsToolTip
End Type

Private cButtons() As tButtonInfo   ' collection of system buttons
' 0 to 3 are used for HTSYSMENU,HTCLOSE,HTMAXBUTTON,HTMINBUTTON -- in that order
' 4 to 8 are used for custom buttons
Private cUserBtn() As tCustomButton ' collection of user buttons
' 4 & above are for custom buttons with custom hitTest code of 101 thru xxx
' ^^ this array's LBound is always 4 and max UBound is 8

Private mMainImplentation As Long
Private mODImplementation As Long       ' implementation pointer when owner-drawn
Private mNotifyUser As Long             ' owner-drawn flags

Private cTracking As Long               ' button ID being tracked with timer
Private cTimer As Long                  ' active timer cross-ref to button ID
Private winActive As Byte               ' cached window state
Private cOptions As Integer
    '1=hide disabled buttons
    '2=no Restore button when maximized
    '4=customButtons displayed
    '8=no Move
    '16=no size
    '32 icon 90degrees
    '64=icon 270degrees
    '128=icon auto rotated with text alignment
    '256 ' reserved, no longer used
Private cSysIconOptions As Byte
    '0=sys icon shown and acts normally
    '1=sys icon not shown but system menu allowed to be displayed (right clicks)
    '2=system menu is not allowed to be displayed, icon visible unless value = 3
    
Private cBtnSize As Long                    ' system default button size
Private cColors(0 To 1) As clsButtonColors  ' active/inactive color options
Private hRotatedImage As Long               ' optional image when system icon is rotated
Private cAutoRedraw As Boolean
Implements iTimer                           ' timer call back
Implements iPropertyBag

'---------------------------------------------------------------------------------------
' Procedure : clsButtons.ResetToSystemDefaults
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Forward resets to our color classes
' Comments  : This class has no values to reset
'---------------------------------------------------------------------------------------
'
Public Sub ResetToSystemDefaults(ByVal Options As eBarResets)
        cColors(0).ResetToSystemDefaults Options
        cColors(1).ResetToSystemDefaults Options
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.HitTest
' DateTime  : 8/30/2005
' Author    : LaVolpe
' Purpose   : Hit test functionality for the titlebar buttons
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function HitTest(ByVal x As Long, ByVal Y As Long) As Long

    If IsArrayEmpty(Not cButtons) Then Exit Function

    ' loop thru the titlebar buttons to determine if mouse is in their area
    Dim B As Long, lRtn As Long
    For B = 0 To UBound(cButtons)
        ' disabled buttons or invisible buttons will have a zero-size Rect
        If PtInRect(cButtons(B).Loc, x, Y) Then
            Select Case B
            Case 0: lRtn = HTSYSMENU
            Case 1: lRtn = HTCLOSE
            Case 2: lRtn = HTMAXBUTTON
            Case 3: lRtn = HTMINBUTTON
            Case Else
                ' hit test = custom button flag
                ' the custom hitTest code is 101 for 1st custom button & +1 for each additional
                lRtn = B + 97
            End Select
            Exit For
        End If
    Next
    HitTest = lRtn

End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.CalcButtonOffsets
' DateTime  : 8/30/2005
' Author    : LaVolpe
' Purpose   : calculate space needed to display buttons
' Comments  :  adjusts the available space to display window caption
'---------------------------------------------------------------------------------------
'
Friend Function CalcButtonOffsets(ByVal cTBarPtr As Long, ByVal isActive As Byte, _
                    tbAlignment As eBarAlignment) As Long

' modifies the passed rectangle structure and returns the overall height/width
' needed to display the titlebar buttons (dependent on orientation of the titlebar)

' Also called when a window is toggled between minimized state and Restore/Maximized
' so the class can update the enabled state of the titlebar buttons

    Dim nonVizFlag As Byte
    Dim TbarRect As RECT, capRect As RECT
    Dim btnID As Long, isDisabled As Boolean
    Dim nrSysBtns As Integer, nrVisible As Integer
    Dim btnXY As POINTAPI, btnRect As RECT
    
    If cTBarPtr <> 0 Then nonVizFlag = 32
    winActive = isActive
    
    ' need to dermine how many buttons will be visible on the titlebar
    ' option to hide disabled buttons
    For btnID = 1 To UBound(cButtons)
        If (cButtons(btnID).Enabled And 2) = 2 Then
            ' user disable this button, don't allow system to override it
            isDisabled = True
        Else
            Select Case btnID
            Case 1: ' sc_close: should always be true but check
                isDisabled = (pEnabled(SC_CLOSE) = False)
            Case 2: 'sc_maximze, sc_restore
                If IsZoomed(cHwnd) = 0 Then
                    ' not maximized (can maximize)
                    cButtons(btnID).Enabled = cButtons(btnID).Enabled And Not 4
                    isDisabled = False
                Else ' maximized (can restore)
                    isDisabled = ((cOptions And 2) = 2)
                    cButtons(btnID).Enabled = cButtons(btnID).Enabled Or 4
                End If
            Case 3: ' can minimize
                isDisabled = (IsIconic(cHwnd) <> 0)
                If isDisabled Then Exit Function
            Case Else   ' all others will be enabled by default, otherwise,
                ' custom buttons would have had the Enable property set to 2
                isDisabled = False
            End Select
        End If
        ' keep track of how many system buttons are also visible
        ' this number compared to the visibility of user buttons will determine
        ' the height of the system buttons
        If isDisabled Then
            If (cOptions And 1) = 1 Then ' hiding disabled buttons
                cButtons(btnID).Enabled = cButtons(btnID).Enabled Or nonVizFlag
            Else
                If btnID < 4 Then nrSysBtns = nrSysBtns + 1
                nrVisible = nrVisible + 1
            End If
            cButtons(btnID).Enabled = cButtons(btnID).Enabled Or 1
        Else
            If btnID < 4 Then nrSysBtns = nrSysBtns + 1
            nrVisible = nrVisible + 1
            cButtons(btnID).Enabled = cButtons(btnID).Enabled And Not 1
        End If
    Next
    
    If cTBarPtr = 0 Then Exit Function      ' no titlebar pointer, no more work
    '^^ only passed from the ProcessMessage routine when window is restored/maximized
    
    cAutoRedraw = True
    If nrVisible = 0 Then
        ' no icon to display? If not, we don't have a titlebar per se
        If hSysMenu = 0 Or ((cSysIconOptions And eSysIconHidden) = eSysIconHidden) Then Exit Function
    End If
    
    ' get the available space to draw the buttons
    CopyMemory TbarRect, ByVal cTBarPtr, &H10
    
    ' the alignment of the titlebar requires significant number of IF statements
    ' when calculating the overall size of the titlebar and caption recangle; therefore
    ' I used 2 similar routines to improve overall speed & less confusing code
    If tbAlignment < barAlignBottom Then ' vertical titlebar
        If MeasureCaptionRect_V(tbAlignment, TbarRect, capRect, nrVisible, nrSysBtns) Then
            CalcButtonOffsets = VarPtr(capRect)
        End If
    Else
        If MeasureCaptionRect_H(tbAlignment, TbarRect, capRect, nrVisible, nrSysBtns) Then
            CalcButtonOffsets = VarPtr(capRect)
        End If
    End If
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.DrawButtons
' DateTime  : 8/30/2005
' Author    : LaVolpe
' Purpose   : Draws all buttons on a titlebar when titlebar is refreshed
'---------------------------------------------------------------------------------------
'
Friend Sub DrawButtons(ByVal hDC As Long, ByVal isActive As Byte)
        
    If IsArrayEmpty(Not cButtons) Then Exit Sub
    
    Dim btnID As Long, lState As Long
    Dim tDC As Long, hSysImg As Long, holdImg As Long
    
    winActive = isActive
    For btnID = 1 To UBound(cButtons)
        ' determine whether button is enabled/disabled
        If cButtons(btnID).Loc.Right > 0 Then
            If (cButtons(btnID).Enabled And 1) = 1 Then
                lState = DFCS_INACTIVE
            Else
                lState = 0
            End If
            DrawButton btnID, lState, hDC ' draw the button
        End If
    Next
    
    ' do the system icon
    If cButtons(0).Loc.Right > cButtons(0).Loc.Left Then
        If hRotatedImage = 0 Then
            ' no rotated image, get system icon from window & draw it
            hSysImg = SendMessage(cHwnd, WM_GETICON, 0, ByVal 0&)
            ' notice we don't destroy the icon we retrieved? This is 'cause
            ' we don't own it. It is a shared resource with the window
        Else
            hSysImg = hRotatedImage
        End If
        If hSysImg <> 0 Then DrawIconEx hDC, cButtons(0).Loc.Left, cButtons(0).Loc.Top, hSysImg, cBtnSize, cBtnSize, 0, 0, &H3
    End If
            

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.Active/Inactive
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Returns the color classes for active/inactive windows
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Get Active() As clsButtonColors
    If cColors(1) Is Nothing Then
        Set cColors(1) = New clsButtonColors
        cColors(1).pUsage = 1
    End If
    Set Active = cColors(1)
End Property
Public Property Get Inactive() As clsButtonColors
    If cColors(0) Is Nothing Then Set cColors(0) = New clsButtonColors
    Set Inactive = cColors(0)
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.HideDisabledButtons Property
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Option to not show disabled titlebar buttons
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Let HideDisabledButtons(bHide As Boolean)
    If bHide Then
        cOptions = cOptions Or 1
    Else
        cOptions = cOptions And Not 1
    End If
    RefreshWindow
End Property
Public Property Get HideDisabledButtons() As Boolean
    HideDisabledButtons = ((cOptions And 1) = 1)
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.SystemIconMenu Property
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Option to hide and or disable the system icon and/or menu
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Let SystemIconMenu(ByVal Settings As eSysIconOptions)
    If Settings < eSysIconNormal Or Settings > eSysMenuHidden Then Settings = eSysIconNormal
    cSysIconOptions = Settings
End Property
Public Property Get SystemIconMenu() As eSysIconOptions
    SystemIconMenu = cSysIconOptions
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.DrawButton
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Draws an individual titlebar button
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub DrawButton(ByVal btnID As Long, ByVal btnState As Long, Optional ByVal hDC As Long)
    
    Dim I As Integer, winDC As Long
    Dim btnRect As RECT, iXY As POINTAPI
    Dim sizeOffset As Long, xOffset As Long
    Dim hBrush As Long, hPen As Long, hOldPen As Long
    Dim disableOffset As Long, hHighlight As Long
    Dim bColors(0 To 3) As Long, fColors(0 To 3) As Long
    Dim bIsOwnerDrawn As Boolean, customTTip As String
    
    btnRect = cButtons(btnID).Loc  ' cache modifyable copy of the actual button coords
    
    ' set some offsets
   
    If (btnState And DFCS_PUSHED) = DFCS_PUSHED Then xOffset = 1  ' button is down, offset 1 down & right
    If (btnState And DFCS_INACTIVE) = DFCS_INACTIVE Then disableOffset = 1 ' button is disabled
    If btnRect.Bottom - btnRect.Top < cBtnSize - 1 Then sizeOffset = 1  ' close button should be reduced
    '^^ when no user buttons displayed on horizontal titlebar, the close button's
    '   dimensions are actually 18x14 to have gap btwn it and the maximize button
    '   On a vertical bar, the adjusment will be from 16x16 to 16x14
    
    If hDC = 0 Then
        ' we are updating the window directly, not our offscreen window
        winDC = 1   ' flag
        hDC = Canvas(False) ' get offscreen DC to do initial painting
        ' align button 0,0
        OffsetRect btnRect, -btnRect.Left, -btnRect.Top
        ' adjust the close button to be 16pixels wide, may currently be 18
        If sizeOffset = 1 And btnID = 1 Then btnRect.Right = btnRect.Right - 2
        
    Else ' updating our offscreen window bitmap
        ' different DC, different offset
        ' adjust the close button to be 16pixels wide, may currently be 18
        If sizeOffset = 1 And btnID = 1 Then btnRect.Left = btnRect.Left + 2
    End If
    
    If sizeOffset = 0 And btnID = 1 Then
        ' no custom buttons being displayed on vertical bar, adjust 18pixels to 16
        If (cOptions And 4) = 4 Then btnRect.Bottom = btnRect.Bottom - 2
    End If
    
        
    ' load the default or custom colors & draw the blank button
    cColors(winActive).GetBorderColors bColors(0), bColors(1), bColors(2), bColors(3)
    DrawBlankButton hDC, btnRect, bColors(0), bColors(1), bColors(2), bColors(3), xOffset
    
    If btnID > 3 Then   ' custom buttons
        bIsOwnerDrawn = True    ' always; mNotifyUser will contain 1024
    Else                ' system buttons, are these owner drawn?
        bIsOwnerDrawn = ((mNotifyUser And odSysButtons) = odSysButtons)
    End If
    If bIsOwnerDrawn Then bIsOwnerDrawn = (mODImplementation <> 0)
    
    If bIsOwnerDrawn Then
        On Error Resume Next ' always, to try & prevent user from crashing
        
        Dim cClient As iOwnerDrawn, CID As CustomItemDraw_LV
        ' set up the UDT to pass to the user
        CID.hDC = hDC
        CID.rcItem = btnRect
        Select Case btnID
            Case 1: CID.itemID = SC_CLOSE
            Case 3: CID.itemID = SC_MINIMIZE
            Case 2
                If ((cButtons(2).Enabled And 4) = 0) Then     ' not maximized
                    CID.itemID = SC_MAXIMIZE
                Else
                    CID.itemID = SC_RESTORE
                End If
            Case Else:
                CID.itemID = cUserBtn(btnID).ID
                CID.itemPos = btnID - 4
        End Select
        CID.itemState = (((btnState = DFCS_PUSHED) * -mcSelect)) Or (((btnState And DFCS_HOT) = DFCS_HOT) * -mcHover) Or ((disableOffset = 1) * -mcDisabled)
        
        GetObjectFromPointer mODImplementation, cClient
        If btnID > 3 Then   ' system button
            If Not cUserBtn(btnID).tTipClass Is Nothing Then customTTip = cUserBtn(btnID).tTipClass.TipText
            cClient.OwnerDrawMessage omDrawUserBtn, winActive + 0, VarPtr(CID), customTTip
        Else                ' user button, different message IDs
            cClient.OwnerDrawMessage omDrawSysBtn, winActive + 0, VarPtr(CID), ""
        End If
        Set cClient = Nothing
        SelectClipRgn hDC, ByVal 0& ' always remove any clipping region client may have left behing
        
    Else    ' not owner drawn item
    
        ' because we can't change the system button colors w/o affecting every window
        ' in the system, we are forced to draw our own system icons
        
        ' get the forecolors to use
        cColors(winActive).GetForeColors fColors(0), fColors(1), fColors(2), , fColors(3)
        
        If disableOffset Then   ' select appropriate color for disabled button
            If fColors(2) > -1 Then
                hPen = CreatePen(0, 1, fColors(2))  ' 2 color disabled button
            Else
                disableOffset = 0
                hPen = CreatePen(0, 1, fColors(1))  ' single color disabled button
            End If
        Else                    ' select color for non-disabled button
            ' if no hover color is used, use the default color instead
            If fColors(3) = -1 Then fColors(3) = fColors(0)
            If (btnState And DFCS_HOT) = DFCS_HOT Then
                hPen = CreatePen(0, 1, fColors(3))  ' hovered
            Else
                hPen = CreatePen(0, 1, fColors(0))  ' not hovered
            End If
        End If
        hOldPen = SelectObject(hDC, hPen)
        
        Select Case btnID
        Case 1 ' close button.
            
            ' adjust the offset for drawing the X, the X is always 8x8
            iXY.x = (btnRect.Right - btnRect.Left - 8) / 2 + btnRect.Left + xOffset
            iXY.Y = (btnRect.Bottom - btnRect.Top - 8) / 2 + btnRect.Top + xOffset
            
            If disableOffset > 0 Then
                ' draw the disabled highlight first if needed
                MoveToEx hDC, iXY.x + 2, iXY.Y, ByVal 0&
                LineTo hDC, iXY.x + 9, iXY.Y + 7
                MoveToEx hDC, iXY.x + 2, iXY.Y + 6, ByVal 0&
                LineTo hDC, iXY.x + 8, iXY.Y
                ' select the "gray" color into the DC
                DeleteObject SelectObject(hDC, CreatePen(0, 1, fColors(1)))
            End If
            ' draw the colored X
            For I = 0 To 1
                MoveToEx hDC, iXY.x + I, iXY.Y, ByVal 0&
                LineTo hDC, iXY.x + 8 - I, iXY.Y + 7
                MoveToEx hDC, iXY.x + I, iXY.Y + 6, ByVal 0&
                LineTo hDC, iXY.x + 8 - I, iXY.Y - 1
            Next
            DeleteObject SelectObject(hDC, hOldPen)
            
        Case 2  ' maximize/restore button
            
            ' offset the starting point for drawing the maximize button
            iXY.x = (btnRect.Right - btnRect.Left - 8) / 2 + btnRect.Left - 1 + xOffset
            iXY.Y = (btnRect.Bottom - btnRect.Top - 8) / 2 + btnRect.Top - 1 + xOffset
            
            If ((cButtons(2).Enabled And 4) = 0) Then     ' not maximized
                
                ' draw the maximized button
                
                If disableOffset > 0 Then
                    MoveToEx hDC, iXY.x + 7, iXY.Y + 2, ByVal 0&
                    LineTo hDC, iXY.x + 1, iXY.Y + 2
                    LineTo hDC, iXY.x + 1, iXY.Y + 9
                    LineTo hDC, iXY.x + 9, iXY.Y + 9
                    LineTo hDC, iXY.x + 9, iXY.Y
                    ' select the "gray" color intot the DC
                    DeleteObject SelectObject(hDC, CreatePen(0, 1, fColors(1)))
                End If
                ' draw the colored maximize button
                MoveToEx hDC, iXY.x, iXY.Y + 2, ByVal 0&
                LineTo hDC, iXY.x, iXY.Y + 8
                LineTo hDC, iXY.x + 8, iXY.Y + 8
                LineTo hDC, iXY.x + 8, iXY.Y + 1
                LineTo hDC, iXY.x, iXY.Y + 1
                LineTo hDC, iXY.x, iXY.Y
                LineTo hDC, iXY.x + 9, iXY.Y
                
            Else 'If winState = SIZE_MAXIMIZED Then
                ' draw the restore button
                
                If disableOffset > 0 Then
                    MoveToEx hDC, iXY.x + 4, iXY.Y + 5, ByVal 0&
                    LineTo hDC, iXY.x + 1, iXY.Y + 5
                    LineTo hDC, iXY.x + 1, iXY.Y + 9
                    LineTo hDC, iXY.x + 6, iXY.Y + 9
                    LineTo hDC, iXY.x + 6, iXY.Y + 6
                    LineTo hDC, iXY.x + 8, iXY.Y + 6
                    LineTo hDC, iXY.x + 8, iXY.Y
                    MoveToEx hDC, iXY.x + 3, iXY.Y + 2, ByVal 0&
                    LineTo hDC, iXY.x + 6, iXY.Y + 2
                    LineTo hDC, iXY.x + 6, iXY.Y + 5
                    ' select the "Gray" color into the DC
                    DeleteObject SelectObject(hDC, CreatePen(0, 1, fColors(1)))
                End If
                
                ' draw the colored Restore button
                MoveToEx hDC, iXY.x, iXY.Y + 4, ByVal 0&
                LineTo hDC, iXY.x, iXY.Y + 8
                LineTo hDC, iXY.x + 5, iXY.Y + 8
                LineTo hDC, iXY.x + 5, iXY.Y + 4
                LineTo hDC, iXY.x, iXY.Y + 4
                LineTo hDC, iXY.x, iXY.Y + 3
                LineTo hDC, iXY.x + 6, iXY.Y + 3
                
                MoveToEx hDC, iXY.x + 6, iXY.Y + 5, ByVal 0&
                LineTo hDC, iXY.x + 7, iXY.Y + 5
                LineTo hDC, iXY.x + 7, iXY.Y + 1
                LineTo hDC, iXY.x + 2, iXY.Y + 1
                LineTo hDC, iXY.x + 2, iXY.Y
                LineTo hDC, iXY.x + 8, iXY.Y
                MoveToEx hDC, iXY.x + 2, iXY.Y + 1, ByVal 0&
                LineTo hDC, iXY.x + 2, iXY.Y + 3
            
            End If
            DeleteObject SelectObject(hDC, hOldPen)
            
        Case 3
        
            ' offset the starting point for drawing the minmize button
            iXY.x = (btnRect.Right - btnRect.Left - 8) / 2 + btnRect.Left - 1 + xOffset
            iXY.Y = (btnRect.Bottom - btnRect.Top - 8) / 2 + btnRect.Top - 1 + xOffset
        
            If disableOffset > 0 Then
                MoveToEx hDC, iXY.x + 2, iXY.Y + 9, ByVal 0&
                LineTo hDC, iXY.x + 8, iXY.Y + 9
                LineTo hDC, iXY.x + 8, iXY.Y + 7
                ' select the "gray" color into the DC
                DeleteObject SelectObject(hDC, CreatePen(0, 1, fColors(1)))
            End If
            ' draw the colored maximize button
            MoveToEx hDC, iXY.x + 1, iXY.Y + 7, ByVal 0&
            LineTo hDC, iXY.x + 7, iXY.Y + 7
            LineTo hDC, iXY.x + 7, iXY.Y + 8
            LineTo hDC, iXY.x, iXY.Y + 8
            DeleteObject SelectObject(hDC, hOldPen)
        
        Case Else
            DeleteObject SelectObject(hDC, hOldPen)
            
        End Select
    
    End If
    
    If winDC = 1 Then
        ' we will be updating the window DC
        winDC = GetWindowDC(cHwnd)
        If winDC = 0 Then Exit Sub
        
        With cButtons(btnID).Loc
            BitBlt winDC, .Right - cBtnSize, .Top, cBtnSize, btnRect.Bottom - btnRect.Top, hDC, 0, 0, vbSrcCopy
        End With
        ReleaseDC cHwnd, winDC
        
    End If
        
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.pMainImplementation
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Used to set the pointer to the primary class
'---------------------------------------------------------------------------------------
'
Friend Property Let pMainImplementation(hWnd As Long, lImpObj As Long)
    mMainImplentation = lImpObj
    cHwnd = hWnd
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.SetMouseAction
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Handles some subclassed messages for the primary class
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function SetMouseAction(ByVal wMsg As Long, ByVal ncHitTest As Long, _
                                ByVal x As Long, ByVal Y As Long) As Long

    Dim btnID As Long
    Dim rtnValue As Long
    Dim cClient As iOwnerDrawn
    Dim customTTip As String
    
' Before a status is updated, any previous status needs to be removed
' Exceptions:
'   1. If a button is down & dragged, no other button can receive a status
'   2. If the mouse is not over the drag button, it is drawn as normal
'   3. If the mouse is over the drag button, it is drawn as down
'   4. When a button up occurs, then
'       If over the drag button - click
'       If not over the drag button, status is cleared and button drawn as normal
' Other states
' When button is moved over with mouse, status 2 is applied (mouse over)
    
    Select Case wMsg
    Case WM_LBUTTONUP
        ' if we weren't tracking a down position, don't fire a click event
        If cTracking = 0 Then Exit Function
        ReleaseCapture
        ' get the button ref & ensure it matches the one that got the button up
        btnID = ConvertHitTestToIndex(HitTest(x, Y))
        
        If btnID = cTracking Then
            rtnValue = 1
            Select Case btnID
            Case 1: btnID = SC_CLOSE
            Case 3: btnID = SC_MINIMIZE
            Case 2
                If ((cButtons(2).Enabled And 4) = 0) Then 'not maximized
                    btnID = SC_MAXIMIZE
                Else
                    btnID = SC_RESTORE
                End If
            Case Else   ' custom buttons
                If mODImplementation <> 0 Then
                    On Error Resume Next ' always, to try & prevent user from crashing
                    If Not cUserBtn(btnID).tTipClass Is Nothing Then customTTip = cUserBtn(btnID).tTipClass.TipText
                    GetObjectFromPointer mODImplementation, cClient
                    cClient.OwnerDrawMessage omUserBtnClick, cUserBtn(btnID).ID, 0, customTTip
                    Set cClient = Nothing
                    btnID = 0
                End If
            End Select
            ' for system buttons, post the appropriate message
            If btnID Then PostMessage cHwnd, WM_SYSCOMMAND, btnID, ByVal 0&
        End If
        ' redraw clicked button as unclicked
        ClearStates
        cTracking = 0
        Exit Function
        
    Case WM_MOUSEMOVE
        btnID = ConvertHitTestToIndex(HitTest(x, Y))
    Case Else
        btnID = ConvertHitTestToIndex(ncHitTest)

    End Select
    
    If btnID = 0 Then
        ' mouse moved off our buttons to somewhere else
        ClearStates
        Exit Function
    End If
    
    Select Case wMsg
        Case WM_NCLBUTTONDOWN
            ' track the button as down unless it is disabled
            ClearStates
            If (cButtons(btnID).Enabled And 1) = 1 Then Exit Function
            cButtons(btnID).Status = 4         ' down, mouse over, drag
            DrawButton btnID, DFCS_PUSHED       ' draw button as down
            cTracking = btnID
            SetCapture cHwnd
            
        Case WM_NCMOUSEMOVE, WM_MOUSEMOVE
            ' basically fire a hover message unless we are already over the button
            If cTracking Then
                If btnID = cTracking Then
                    ' tracking down state, it was up, draw it as down
                    If cButtons(btnID).Status = 0 Then
                        cButtons(btnID).Status = 4
                        DrawButton btnID, DFCS_PUSHED
                    End If
                Else
                    If cButtons(cTracking).Status = 4 Then
                        ' tracking down state, it was down, draw it as up
                        cButtons(cTracking).Status = 0
                        DrawButton cTracking, 0&
                    End If
                End If
            Else
                ' abort if button is disabled
                If (cButtons(btnID).Enabled And 1) = 1 Then
                    ClearStates
                    Exit Function
                Else
                    If cButtons(btnID).Status = 2 Then
                        ' already showing as mouse over
                        Exit Function
                    Else
                        ' show as mouse hover
                        ClearStates
                        cButtons(btnID).Status = 2
                        DrawButton btnID, DFCS_HOT
                        CreateTimer True, btnID
                    End If
                End If
            End If
            
    End Select
            
SetMouseAction = True

End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.ClearStates
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Draws a button in the up state
'---------------------------------------------------------------------------------------
'
Private Sub ClearStates()

Dim B As Long
    CreateTimer False, 0
    For B = 1 To UBound(cButtons)
        If cButtons(B).Status > 0 Then
            cButtons(B).Status = 0
            DrawButton B, 0&
            Exit For
        End If
    Next
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.ConvertHitTestToIndex
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Helper function converts SC_values into button IDs
'---------------------------------------------------------------------------------------
'
Private Function ConvertHitTestToIndex(ncHitTest As Long) As Long

Dim btnID As Long

    If ncHitTest < HTMenu_Custom Then
        Select Case ncHitTest
        Case HTMINBUTTON: btnID = 3
        Case HTMAXBUTTON: btnID = 2
        Case HTCLOSE: btnID = 1
        End Select
    Else
        btnID = ncHitTest - 97
        If btnID < 4 Or btnID > UBound(cButtons) Then btnID = 0
    End If
    ConvertHitTestToIndex = btnID
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.GetXRef
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Helper function to convert SC_values into clone values
'---------------------------------------------------------------------------------------
'
Private Function GetXRef(stdConst As Long) As Long

Select Case stdConst
    Case SC_CLOSE:      GetXRef = SC_CLOSE_Clone
    Case SC_MAXIMIZE:   GetXRef = SC_MAXIMIZE_Clone
    Case SC_MINIMIZE:   GetXRef = SC_MINIMIZE_Clone
    Case SC_MOVE:       GetXRef = SC_MOVE_Clone
    Case SC_SIZE:       GetXRef = SC_SIZE_Clone
    Case SC_RESTORE:    GetXRef = SC_RESTORE_Clone
    Case Else   ' custom menu item on system tray or bad parameter
        GetXRef = stdConst
End Select

End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.CreateTimer
' DateTime  : 9/12/2005
' Author    : LaVolpe
' Purpose   : Establishes a timer that fires 8x a second
' Comments  : used to determine a mouse out event for titlebar buttons
'---------------------------------------------------------------------------------------
'
Private Sub CreateTimer(bNew As Boolean, btnID As Long)

If cTimer > 0 Then
    If Not bNew Then    ' killing the timer
        SetStopTimer False, cHwnd, ObjPtr(Me), 0
        cTimer = 0
    End If
Else
    If bNew Then        ' creating the timer
        SetStopTimer True, cHwnd, ObjPtr(Me), 150
        cTimer = btnID
    End If
End If
        
End Sub

'---------------------------------------------------------------------------------------
' Procedure : clsButtons.Read/WriteProperites
' DateTime  : 9/26/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve custom properites
' Comments  : Used by Export/ImportCustomSettings
'---------------------------------------------------------------------------------------
'
Private Sub iPropertyBag_ReadProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim I As Integer, mnuRef As Long
    Dim sProp As String, sProps() As String
    Dim iBag As iPropertyBag
    
    cOptions = propBag.ReadProperty("BN.Options", 0)
    mNotifyUser = propBag.ReadProperty("BN.OD", 0)
    cSysIconOptions = propBag.ReadProperty("BN.SysIcon", 0)
    
    If (Limits And exCustomButtons) = exCustomButtons Then
        I = propBag.ReadProperty("BN.UsrBtns", 0)
        
        ReDim cButtons(0 To I + 3)
        If UBound(cButtons) > 3 Then
            ReDim cUserBtn(4 To UBound(cButtons))
        Else
            Erase cUserBtn
        End If
        
        For I = 4 To UBound(cButtons)
            sProps = Split(propBag.ReadProperty("BN.User" & I), Chr$(0))
            cUserBtn(I).ID = CLng(sProps(0))
            If Len(sProps(1)) > 0 Then
                Set cUserBtn(I).tTipClass = New clsToolTip
                cUserBtn(I).tTipClass.CreateToolTip cHwnd, sProps(1), 0
            End If
        Next
        sProps = Split(propBag.ReadProperty("BN.Enabled"), ",")
        For I = 4 To UBound(cButtons)
            EnableUserButton cUserBtn(I).ID, Abs(Val(sProps(I)) - 1)
        Next
    Else
        sProps = Split(propBag.ReadProperty("BN.Enabled"), ",")
        mNotifyUser = mNotifyUser And Not 1024
    End If
    
    sProps = Split(propBag.ReadProperty("BN.Enabled"), ",")
    For I = 1 To 3
        Select Case I
        Case 1: mnuRef = SC_CLOSE
        Case 2: mnuRef = SC_MAXIMIZE
        Case 3: mnuRef = SC_MINIMIZE
        End Select
        EnableSysMenuItem mnuRef, Abs(Val(sProps(I)) - 1)
    Next
   
    mODImplementation = lParam
    
    Set iBag = Me.Active
    iBag.ReadProperties propBag, Limits, lParam, Index
    Set iBag = Me.Inactive
    iBag.ReadProperties propBag, Limits, lParam, Index
    Set iBag = Nothing
    
    RefreshIcon
    If ((cOptions And 8) = 0) Then
        EnableSysMenuItem SC_MOVE, sysEnable
    Else
        EnableSysMenuItem SC_MOVE, sysDisable
    End If
    If ((cOptions And 16) = 0) Then
        EnableSysMenuItem SC_SIZE, sysEnable
    Else
        EnableSysMenuItem SC_SIZE, sysDisable
    End If
    
End Sub

Private Sub iPropertyBag_WriteProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim I As Integer, sProp As String
    Dim iBag As iPropertyBag
    
    propBag.WriteProperty "BN.Options", cOptions, 0
    propBag.WriteProperty "BN.OD", mNotifyUser, 0
    propBag.WriteProperty "BN.SysIcon", cSysIconOptions, 0
    propBag.WriteProperty "BN.UsrBtns", UBound(cButtons) - 3, 0
    
    For I = 0 To UBound(cButtons)
        sProp = sProp & "," & Abs((cButtons(I).Enabled And 2) = 2)
    Next
    propBag.WriteProperty "BN.Enabled", Mid$(sProp, 2)
    
    If (Limits And exCustomButtons) = exCustomButtons Then
        For I = 4 To UBound(cButtons)
            sProp = cUserBtn(I).ID & Chr$(0)
            If Not cUserBtn(I).tTipClass Is Nothing Then
                sProp = sProp & cUserBtn(I).tTipClass.TipText
            End If
            propBag.WriteProperty "BN.User" & I, sProp
        Next
    End If
    
    Set iBag = Me.Active
    iBag.WriteProperties propBag, Limits
    Set iBag = Me.Inactive
    iBag.WriteProperties propBag, Limits
    Set iBag = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : clsButtons.iTimer_TimerEvent
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Timer callback function
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub iTimer_TimerEvent(ByVal hWnd As Long, ByVal TimerID As Long)

    Dim wPTs As POINTAPI, wRect As RECT
    Dim btnID As Long
    
    ' if the button has already been painted by some other action, simply abort
    If cButtons(cTimer).Status = 0 Then
        CreateTimer False, 0
        
    Else
        GetCursorPos wPTs
        GetWindowRect hWnd, wRect
        ' offset the points to our titlebar location
        wPTs.x = wPTs.x - wRect.Left
        wPTs.Y = wPTs.Y - wRect.Top
        
        ' get button ID from the hit test code
        btnID = ConvertHitTestToIndex(HitTest(wPTs.x, wPTs.Y))
        ' if the mouse is still over the button highlighted, no action needed
        If btnID <> cTimer Then
            ' otherwise, let's kill the timer & repaint the button to normal
            btnID = cTimer          ' cache 'cause CreateTimer resets the cTimer value
            CreateTimer False, 0
            DrawButton btnID, 0&
            cButtons(btnID).Status = 0  ' reset the status of this button
        End If
    End If
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.MeasureCaptionRect_V
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Calculates positioning of buttons on titlebar & sets the caption rectangle
' Comments  : Vertical Titlebars only. See Below
'---------------------------------------------------------------------------------------
'
Private Function MeasureCaptionRect_V(tbAlignment As eBarAlignment, _
                    TbarRect As RECT, capRect As RECT, _
                    nrVisible As Integer, nrSysBtns As Integer) As Boolean

    ' there are significant amounts of IF statements needed to position the
    ' buttons and calculate the caption rectangle to calculate this data
    ' for vertical and horizontal bars. Therefore, 2 separate routines
    ' are used so we don't process all those IFs in the following loops
    
    Dim btnID As Long
    Dim isDisabled As Boolean
    Dim btnRowOffset As Long
    
    Dim btnXY As POINTAPI
    Dim vRowOffset As Long
    
    ' the close button is always 16x16 or whatever the system button size is
    ' but displaying that 3 pixel separation btwn close & max, depends on
    ' the visibility of custom buttons
    If nrSysBtns = nrVisible Then cOptions = cOptions Or 4 Else cOptions = cOptions And Not 4
    
    For btnID = 1 To UBound(cButtons)
        If (cButtons(btnID).Enabled And 32) = 32 Then
            ' won't be seen, create an invalid rect & remove the temporary flag
            SetRect cButtons(btnID).Loc, 0, 0, 0, 0
            cButtons(btnID).Enabled = cButtons(btnID).Enabled And Not 32
        Else
            If btnID = 1 Then ' close button
                SetRect cButtons(btnID).Loc, 0, 0, cBtnSize, cBtnSize
            Else
                If nrVisible > nrSysBtns Then   ' custom buttons being displayed
                    SetRect cButtons(btnID).Loc, 0, 0, cBtnSize, cBtnSize
                Else    ' otherwise subtract 2 pixels from the button sizes
                    SetRect cButtons(btnID).Loc, 0, 0, cBtnSize, cBtnSize - 2
                End If
            End If
            ' start measuring the caption rectangle's top offset
            capRect.Top = capRect.Top + cButtons(btnID).Loc.Bottom
        End If
    Next
    
    ' calcualte the button break when displaying more than
    Select Case nrVisible
    Case Is < 5
        vRowOffset = nrVisible + 1
    Case 5
        vRowOffset = 4
    Case Else
        vRowOffset = nrVisible \ 2 + 1
    End Select
    
    ' finish calculating the caption rectangle's top offset
    If nrVisible > 4 Then
        capRect.Top = capRect.Top - (nrVisible - ((nrVisible + 0.75) \ 2)) * cBtnSize
        btnRowOffset = cBtnSize + 2 ' adjust horizontal width to display more than 4 buttons
    End If
    capRect.Top = capRect.Top + 4 - ((nrVisible > 0) * 4)
    
    ' calculate the horizontal offset to center button(s) in titlebar
    ' There is no standard for multirow button placement on vertical titlebars, no duh
    ' So I am making a rule here...
    '       Left Vertical, the system buttons are left aligned
    '       Right Vertical, the system buttons are right aligned
    
    ' btnXY are offsets for position the buttons
    If tbAlignment = barAlignRight Then
        If TbarRect.Left + cBtnSize + 3 + btnRowOffset > TbarRect.Right Then
            TbarRect.Left = TbarRect.Right - cBtnSize - 3 - btnRowOffset
        End If
        btnXY.x = ((TbarRect.Right - TbarRect.Left) + 1 - cBtnSize) \ 2 + TbarRect.Left
        If nrVisible > 4 Then btnXY.x = btnXY.x + (cBtnSize \ 2) + 1
    Else
        If TbarRect.Left + cBtnSize + 3 + btnRowOffset > TbarRect.Right Then
            TbarRect.Right = TbarRect.Left + cBtnSize + 3 + btnRowOffset
        End If
        btnXY.x = ((TbarRect.Right - TbarRect.Left) + 1 - cBtnSize) \ 2 + TbarRect.Left
        If nrVisible > 4 Then btnXY.x = btnXY.x - (cBtnSize \ 2) - 1
    End If
    ' buffer from bottom edge
    capRect.Bottom = capRect.Bottom + 4
    btnXY.Y = TbarRect.Top + 2  ' start 1st button 2 pixels from top of titlebar
    btnRowOffset = btnXY.x      ' cache this starting position
    ' now finalize the caption rectangle
    SetRect capRect, TbarRect.Left, TbarRect.Top + capRect.Top, TbarRect.Right, TbarRect.Bottom - capRect.Bottom
    
    ' reset for use below
    nrVisible = 0
    ' position buttons
    For btnID = 1 To UBound(cButtons)
        
        If cButtons(btnID).Loc.Right > 0 Then   ' button will be displayed
            nrVisible = nrVisible + 1
            If nrVisible = vRowOffset Then      ' is this the row break?
                ' if so, set the new X,Y start position for the new row
                If tbAlignment = barAlignLeft Then
                    ' keeping rows left to right
                    btnXY.x = btnRowOffset + cBtnSize + 1
                    btnXY.Y = TbarRect.Top + 2
                Else ' keeping rows right to left
                    btnXY.x = btnRowOffset - cBtnSize - 1
                    btnXY.Y = TbarRect.Top + 2
                End If
            End If
            ' now position the button
            OffsetRect cButtons(btnID).Loc, btnXY.x, btnXY.Y
            btnXY.Y = cButtons(btnID).Loc.Bottom  ' set next button's Y coordinate
        
        End If
    Next
    
    ' position the system icon
    If hSysMenu = 0 Or ((cSysIconOptions And eSysIconHidden) = eSysIconHidden) Then
        SetRect cButtons(0).Loc, 0, 0, 0, 0
    Else
        btnXY.x = ((TbarRect.Right - TbarRect.Left) + 1 - cBtnSize) \ 2 + TbarRect.Left
        If tbAlignment = barAlignRight Then
            capRect.Top = capRect.Top + cBtnSize + 4
            SetRect cButtons(0).Loc, btnXY.x, capRect.Top - cBtnSize - 4, btnXY.x + cBtnSize, capRect.Top - 4
        Else
            capRect.Bottom = capRect.Bottom - cBtnSize - 4
            SetRect cButtons(0).Loc, btnXY.x, TbarRect.Bottom - cBtnSize - 4, btnXY.x + cBtnSize, TbarRect.Bottom - 4
        End If
    End If
    MeasureCaptionRect_V = True
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.MeasureCaptionRect_H
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Calculates positioning of buttons on titlebar & sets the caption rectangle
' Comments  : Horizontal Titlebars only. See Below
'---------------------------------------------------------------------------------------
'
Private Function MeasureCaptionRect_H(tbAlignment As eBarAlignment, _
                    TbarRect As RECT, capRect As RECT, _
                    nrVisible As Integer, nrSysBtns As Integer) As Boolean

    ' there are significant amounts of IF statements needed to position the
    ' buttons and calculate the caption rectangle to calculate this data
    ' for vertical and horizontal bars. Therefore, 2 separate routines
    ' are used so we don't process all those IFs in the following loops
    
    Dim btnID As Long
    Dim isDisabled As Boolean
    Dim btnRowOffset As Long
    
    Dim btnXY As POINTAPI
    Dim vRowOffset As Long
    
    ' the option to tweak height of the Close button does not apply to horizontal bars
    cOptions = cOptions And Not 4
    
    For btnID = 1 To UBound(cButtons)
        If (cButtons(btnID).Enabled And 32) = 32 Then
            ' won't be seen, create an invalid rect & remove the temporary flag
            SetRect cButtons(btnID).Loc, 0, 0, 0, 0
            cButtons(btnID).Enabled = cButtons(btnID).Enabled And Not 32
        Else
            If btnID = 1 Then ' close button
                SetRect cButtons(btnID).Loc, 0, 0, cBtnSize - ((nrSysBtns = nrVisible) * 2), cBtnSize + ((nrSysBtns = nrVisible) * 2)
            Else
                If nrVisible > nrSysBtns Then   ' custom buttons being displayed
                    SetRect cButtons(btnID).Loc, 0, 0, cBtnSize, cBtnSize
                Else    ' otherwise subtract 2 pixels from the button sizes
                    SetRect cButtons(btnID).Loc, 0, 0, cBtnSize, cBtnSize - 2
                End If
            End If
            ' start measuring the caption rectangle's right offset
            capRect.Right = capRect.Right + cButtons(btnID).Loc.Right
        End If
    Next
    
    ' calcualte the button break when displaying more than
    Select Case nrVisible
    Case Is < 5
        vRowOffset = nrVisible + 1
    Case 5
        vRowOffset = 4
    Case Else
        vRowOffset = nrVisible \ 2 + 1
    End Select
    
    ' finish calculating the caption rectangle's right offset
    If nrVisible > 4 Then
        capRect.Right = capRect.Right - (nrVisible - ((nrVisible + 0.75) \ 2)) * cBtnSize
        btnRowOffset = cBtnSize + 2 ' adjust horizontal width to display more than 4 buttons
    End If
    capRect.Right = capRect.Right + 4 - ((nrVisible > 0) * 4)
    capRect.Left = capRect.Left + 4
    
    ' calculate the vertical offset to center button(s) in titlebar
    
    ' btnXY are offsets for position the buttons
    If TbarRect.Top + cBtnSize + 3 + btnRowOffset > TbarRect.Bottom Then
        If tbAlignment = barAlignTop Then
            TbarRect.Bottom = TbarRect.Top + cBtnSize + 3 + btnRowOffset
        Else
            TbarRect.Top = TbarRect.Bottom - cBtnSize - 3 - btnRowOffset
        End If
    End If
    btnXY.Y = ((TbarRect.Bottom - TbarRect.Top) + 1 - cBtnSize) \ 2 + TbarRect.Top
    ' adjust for all buttons being forced at 16pixels height
    If nrVisible > 4 Then
        btnXY.Y = btnXY.Y - (cBtnSize \ 2) - 1
    Else 'adjust height if only system buttons are being displayed
        If nrVisible = nrSysBtns Then btnXY.Y = btnXY.Y + 1
    End If
    btnXY.x = TbarRect.Right - 2  ' start 1st button 2 pixels from right of titlebar
    btnRowOffset = btnXY.Y  ' cache this starting position
    ' now finalize the caption rectangle
    SetRect capRect, TbarRect.Left + capRect.Left, TbarRect.Top, TbarRect.Right - capRect.Right, TbarRect.Bottom
        
    ' reset for use below
    nrVisible = 0
    ' position buttons
    For btnID = 1 To UBound(cButtons)
        
        If cButtons(btnID).Loc.Right > 0 Then   ' button will be displayed
            nrVisible = nrVisible + 1
            If nrVisible = vRowOffset Then      ' is this the row break?
                ' if so, set the new X,Y start position for the new row
                btnXY.Y = btnRowOffset + cBtnSize + 1
                btnXY.x = TbarRect.Right - 2
            End If
            ' now position the button from right to left
            OffsetRect cButtons(btnID).Loc, btnXY.x - cButtons(btnID).Loc.Right, btnXY.Y
            btnXY.x = cButtons(btnID).Loc.Left  'set next button's X coordinate
            
        End If
    Next
    
    ' position the system icon
    If hSysMenu = 0 Or ((cSysIconOptions And eSysIconHidden) = eSysIconHidden) Then
        SetRect cButtons(0).Loc, 0, 0, 0, 0
    Else
        btnXY.Y = ((TbarRect.Bottom - TbarRect.Top) + 1 - cBtnSize) \ 2 + TbarRect.Top
        SetRect cButtons(0).Loc, TbarRect.Left + 4, btnXY.Y, TbarRect.Left + 4 + cBtnSize, btnXY.Y + cBtnSize
        capRect.Left = capRect.Left + cBtnSize + 4
    End If
    MeasureCaptionRect_H = True
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.GetLocalSysMenu
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Retrieves copy of the system menu and also initializes it if needed
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function GetLocalSysMenu() As Boolean

Dim bInitialize As Boolean, hMenu As Long
If cHwnd = 0 Then Exit Function

' see if a copy has already been made?
hMenu = GetSystemMenu(cHwnd, 0)
If hMenu = 0 Then
    bInitialize = True
    ' no copy, create one now & get the handle
    GetSystemMenu cHwnd, True
    hSysMenu = GetSystemMenu(cHwnd, False)
    If hSysMenu = 0 Then Exit Function
Else
    If hMenu <> hSysMenu Then bInitialize = True
    hSysMenu = hMenu
End If

' now, if a menu has never been activated yet, then all menu items
' may appear as "available" and enabled when returned by GetMenuItemInfo. However,
' immediately before system menu is first shown, VB will remove ones that you
' decided you didn't want (i.e., Movable=False, MinButton=False, MaxButton=False)

' So we don't get false readings when processing routines in the class, we will
' send your window a fake message indicating the system menu is about to be shown.
' That way VB will update the menus appropriately.
If hSysMenu <> 0 Then
    If bInitialize = True Then
        ' send the fake popup & then send a "never-mind" message.
        ' This api is used vs SendMessage to prevent DLL from being hung by
        ' sending a message to an already hung application
        ' FYI: 65536 = LoWord 0 & HiWord 1
        SendMessageTimeout cHwnd, WM_ENTERMENULOOP, 0, 0, 0, 50, 0
        SendMessageTimeout cHwnd, WM_INITMENUPOPUP, hSysMenu, 65536, 0, 50, 0
        SendMessageTimeout cHwnd, WM_CANCELMODE, hSysMenu, 0, 0, 10, 0
'        Debug.Print "initialized cSysmenu"
    End If
    ' the function's return result
    GetLocalSysMenu = True
End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.EnableSysMenuItem
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Allows user to enable/disable system buttons and/or menu items
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function EnableSysMenuItem(ByVal sysMenuItem As eSysButtons, _
                                ByVal Action As eEnableConstants) As Boolean

If Action < sysDisable Or Action > sysEnable Then Exit Function
If sysMenuItem = 0 Then Exit Function

' Pass only one value at a time as the sysMenuItem parameter

' Function will disable system menu items and re-enable those that it disabled.
' This is done by changing the menu item's ID to a non-system ID (cloning)
' which effectively disables the menu item & any related button (i.e., close).
' In order to re-enable the cloned button, the system menu ID will be replaced.

' This was written to allow toggling custom added menu items that exist on the
' system menu. To toggle those, simply pass that custom menu's ID as the
' sysMenuItem parameter (those custom menu items are not cloned)

' About VB Form's Movable property if set to True in design view...
' This function will not work for VB Form's Movable property. VB will internally
' reset that property back to False just before the menu is displayed. In addition,
' VB will trap the WM_NCLButtonDown message to prevent you from moving the form.
' So, never set Movable=False, instead use this class to do it as needed
' so you can undo it at will.


Dim MII As MENUITEMINFO
Dim cloneID As Long ' must be integer value in a long variable
Dim curID As Long   ' ID to be used to ref menu structure
Dim lRtn As Boolean, bUpdate As Boolean
Dim btnID As Long

' cross-reference passed menu ID to local button ID
Select Case sysMenuItem
    Case SC_CLOSE: btnID = 1
    Case SC_MAXIMIZE: btnID = 2
    Case SC_MINIMIZE: btnID = 3
    Case SC_RESTORE, SC_MOVE, SC_SIZE
        ' these don't have buttons
        btnID = -1
    Case SC_RESTORE_Clone
        ' in-class use only. Disables/enables Restore menu item when form is maximized
        sysMenuItem = SC_RESTORE
        btnID = 0
End Select

' call function to return the system menu handle
If GetLocalSysMenu = False Then
    ' no system menu, disable all buttons, but not non-button menu items
    cButtons(0).Enabled = 1
    cButtons(1).Enabled = 1
    cButtons(2).Enabled = 1
    cButtons(3).Enabled = 1
    Exit Function
End If

' return the clone ID for the passed sysMenuItem ID
cloneID = GetXRef(sysMenuItem)

MII.cbSize = Len(MII)
MII.fMask = MIIM_ID Or MIIM_STATE   ' only need to tweak ID & States for cloning

' see if the menu item exists
If GetMenuItemInfo(hSysMenu, sysMenuItem, False, MII) Then
    curID = sysMenuItem ' original menu item, not a clone
Else
    ' Cloned? If not, it simply doesn't exist in the system menu
    If GetMenuItemInfo(hSysMenu, cloneID, False, MII) <> 0 Then curID = cloneID
End If

If curID <> 0 Then
    ' When curID=0 then we don't have a menu item to enable/disable
    
    If Action = sysEnable Then ' want to enable menu item
        MII.fState = 0 ' reset to enabled
        If curID = cloneID Then ' cloned, need to unclone
            MII.wID = sysMenuItem   ' reset back to system ID
            bUpdate = True          ' flag to do the update
        Else    ' not cloned, update if needed
            ' primarily used for custom system menu items
            If ((MII.fState Or MFS_GRAYED) = MII.fState) Then lRtn = True Else bUpdate = True
        End If
        
    Else   ' want to disable menu item
        If curID = sysMenuItem Then     ' not cloned, clone & then change
            MII.fState = MFS_GRAYED ' disabled
            MII.wID = cloneID       ' set clone ID
            bUpdate = True          ' flag to do the update
        Else    ' already cloned which means already disabled
            lRtn = True
        End If
    End If
    ' next line returns zero if it fails... Enable/Disable menu item
    If bUpdate Then lRtn = (SetMenuItemInfo(hSysMenu, curID, False, MII) <> 0)

    Select Case btnID
        ' now we track the individual state of the button
        Case 1, 2, 3    ' SC_Close, SC_Maximize, SC_Minimize
            ' remove the Disabled flag & re-apply if needed
            cButtons(btnID).Enabled = (Action - 1) * -3
        Case Is < 0
            ' these don't have buttons, so we use the cOptions variable to track state
            Select Case sysMenuItem
            Case SC_MOVE
                cOptions = (cOptions And Not 8) Or (Action - 1) * -8
            Case SC_RESTORE
                cOptions = (cOptions And Not 2) Or (Action - 1) * -2
            Case SC_SIZE
                cOptions = (cOptions And Not 16) Or (Action - 1) * -16
        Case 0 ' called internally to toggle the Restore menu item's enable/disable state
        End Select
    End Select
End If

EnableSysMenuItem = lRtn

End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.EnableUserButton
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Allows user to enable/disable a custom button
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function EnableUserButton(ButtonID As Long, Action As eEnableConstants) As Boolean
' function that enables/disables a user button
' Disabled user buttons will only forward a paint request when the titlebar
' needs to be painted. It will not forward hover, down or click events

    Dim I As Integer
    For I = 4 To UBound(cButtons)
        If cUserBtn(I).ID = ButtonID Then
            cButtons(I).Enabled = (Action = sysDisable) * -2
            EnableUserButton = True
            Exit For
        End If
    Next

End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.pEnabled
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Helper function to quickly identify a button as being enabled/disabled
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Function pEnabled(ByVal sysMenuItem As eSysButtons) As Boolean
' Function returns whether or not a specific menu item is enabled
' Used internally only to update button drawing states or used when
' a command is recieved to minimize/maximize/move, etc a window

If sysMenuItem = 0 Then Exit Function

Dim bRtn As Boolean

If hSysMenu = 0 Then GetLocalSysMenu

Select Case sysMenuItem

    Case SC_CLOSE: bRtn = (cButtons(1).Enabled = 0)
    
    Case SC_MAXIMIZE
        bRtn = ((cButtons(2).Enabled And 4) = 0)  ' otherwise maximized already
        If bRtn Then bRtn = ((cButtons(2).Enabled And 1) = 0)
    
    Case SC_MINIMIZE: bRtn = (cButtons(3).Enabled = 0)
    
    Case SC_MOVE, (SC_MOVE Or HTCAPTION)
        bRtn = (((cOptions And 8) = 0) And ((cButtons(2).Enabled And 4) = 0)) ' not maximized
        sysMenuItem = SC_MOVE
    
    Case SC_RESTORE:
        If (cButtons(2).Enabled And 4) = 4 Then   ' maximized window
            bRtn = ((cButtons(2).Enabled And 1) = 0)
        Else
            bRtn = True
        End If
    
    Case SC_SIZE: bRtn = (((cOptions And 16) = 0) And ((cButtons(2).Enabled And 4) = 0))
    
    Case Else
        ' test for a custom SC_Size command
        Select Case (sysMenuItem And Not SC_SIZE) + 9&
            Case HTLEFT To HTBOTTOMRIGHT
                sysMenuItem = SC_SIZE
                bRtn = (((cOptions And 16) = 0) And ((cButtons(2).Enabled And 4) = 0))
        End Select
End Select
    
pEnabled = bRtn

End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.DrawBlankButton
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Draws a blank button in the up or down state, using custom colors
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub DrawBlankButton(hDC As Long, btnRect As RECT, Face As Long, Highlight As Long, Shadow As Long, DarkShadow As Long, xOffset As Long)
    
    Dim hBrushPen As Long
    
    hBrushPen = CreateSolidBrush(Face)
    FillRect hDC, btnRect, hBrushPen
    DeleteObject hBrushPen
    
    If xOffset Then ' button in down position
        hBrushPen = SelectObject(hDC, CreatePen(0, 1, DarkShadow))
        MoveToEx hDC, btnRect.Left, btnRect.Bottom - 1, ByVal 0&
        LineTo hDC, btnRect.Left, btnRect.Top
        LineTo hDC, btnRect.Right - 1, btnRect.Top
        DeleteObject SelectObject(hDC, CreatePen(0, 1, Highlight))
        LineTo hDC, btnRect.Right - 1, btnRect.Bottom - 1
        LineTo hDC, btnRect.Left - 1, btnRect.Bottom - 1
        DeleteObject SelectObject(hDC, CreatePen(0, 1, Shadow))
        MoveToEx hDC, btnRect.Left + 1, btnRect.Bottom - 2, ByVal 0&
        LineTo hDC, btnRect.Left + 1, btnRect.Top + 1
        LineTo hDC, btnRect.Right - 2, btnRect.Top + 1
    
    Else ' draw it in the up position
        hBrushPen = SelectObject(hDC, CreatePen(0, 1, Highlight))
        MoveToEx hDC, btnRect.Left, btnRect.Bottom - 1, ByVal 0&
        LineTo hDC, btnRect.Left, btnRect.Top
        LineTo hDC, btnRect.Right - 1, btnRect.Top
        DeleteObject SelectObject(hDC, CreatePen(0, 1, DarkShadow))
        LineTo hDC, btnRect.Right - 1, btnRect.Bottom - 1
        LineTo hDC, btnRect.Left - 1, btnRect.Bottom - 1
        DeleteObject SelectObject(hDC, CreatePen(0, 1, Shadow))
        MoveToEx hDC, btnRect.Right - 2, btnRect.Top + 1, ByVal 0&
        LineTo hDC, btnRect.Right - 2, btnRect.Bottom - 2
        LineTo hDC, btnRect.Left, btnRect.Bottom - 2
    End If
    
    DeleteObject SelectObject(hDC, hBrushPen)
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.AddButton
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Allows user to create a custom button for the titlebar
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function AddButton(iImplementation As iOwnerDrawn, ByVal ButtonID As Long, Optional ByVal ToolTipText As String) As Boolean

' allows user to add up to 5 custom buttons, to the titlebar, using their own
' button ID and tool tip

    If iImplementation Is Nothing Then Exit Function
    ' ^^ Required. The DLL does not draw user buttons with the following exception
    '    The DLL will provide a blank button in the proper button state to draw on or over
    If cHwnd = 0 Then Exit Function
    ' ^^ can't add a button until the form is subclassed
    
    Dim btnID As Long
    
    Select Case UBound(cButtons)
    Case 3: ' no user buttons added
        ReDim cUserBtn(4 To 4)
        btnID = 4
    Case 8: 'maxed out of user buttons
        Exit Function
    Case Else
        For btnID = 4 To UBound(cButtons)
            If cUserBtn(btnID).ID = ButtonID Then Exit Function
        Next
        ReDim Preserve cUserBtn(4 To btnID)
    End Select
    
    ' add the custom information for the user button
    With cUserBtn(btnID)
        ' add the userbutton to the class's button collection
         If btnID > UBound(cButtons) Then ReDim Preserve cButtons(0 To btnID)
        .ID = ButtonID
        If ToolTipText <> "" Then
            Set .tTipClass = New clsToolTip
            .tTipClass.CreateToolTip cHwnd, ToolTipText, VarPtr(cButtons(btnID).Loc)
        End If
    End With
    
    ' update the implementation pointer & return result
    mODImplementation = ObjPtr(iImplementation)
    mNotifyUser = mNotifyUser Or 1024   ' flag indicating user buttons in play
    AddButton = True
    RefreshWindow
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.RemoveButton
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Allows user to remove a custom button from the titlebar
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function RemoveButton(ByVal ButtonID As Long, Optional AllButtons As Boolean) As Boolean
    
' removes a user button, if it exists
    
    If UBound(cButtons) = 3 Then Exit Function ' no user buttons
    
    If AllButtons Then  ' remove all user buttons
        ReDim Preserve cButtons(0 To 3)
        Erase cUserBtn
        RemoveButton = True
    Else
        Dim Btn As Integer, I As Integer
        For Btn = 4 To UBound(cUserBtn)
            ' find the user button to delete
            If cUserBtn(Btn).ID = ButtonID Then
                ' found it
                If UBound(cButtons) = 4 Then
                    ' when number of userbuttons is 1, simply clear the userbutton array
                    ReDim Preserve cButtons(0 To 3)
                    Erase cUserBtn
                Else
                    For I = Btn + 1 To UBound(cUserBtn)
                        cUserBtn(I - 1) = cUserBtn(I)
                        cButtons(I - 1) = cButtons(I) ' carry over any disable states
                    Next
                    ReDim Preserve cUserBtn(4 To UBound(cUserBtn) - 1)
                    ReDim Preserve cButtons(0 To UBound(cUserBtn))
                End If
                RemoveButton = True
                Exit For
            End If
        Next
    End If
    If UBound(cButtons) = 3 Then ' all custom buttons were removed
        mNotifyUser = mNotifyUser And Not 1024  ' no user buttons in play, remove flag
        If mNotifyUser = 0 Then mODImplementation = 0 ' remove pointer to implmentation
    End If
    RefreshWindow
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.PositionButton
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Allows user to reposition/re-order custom buttons
' Comments  : Why?
'---------------------------------------------------------------------------------------
'
Public Function PositionButton(ByVal ButtonID As Long, ByVal NewPosition As Long) As Boolean
    
' re-Orders/re-positions a button at the passed NewPosition value (0-bound positions)
    
    ' check if any userbuttons exist or if only one exists
    If UBound(cButtons) < 5 Then Exit Function
    
    ' ensure provided parameter is within bounds (0 to 4)
    If NewPosition < 1 Then NewPosition = 0
    If NewPosition > UBound(cButtons) - 4 Then NewPosition = UBound(cButtons) - 4
    ' convert zero LBound position to a 4 LBound position
    NewPosition = NewPosition + 4
    
    Dim Btn As Integer, I As Integer
    Dim newUserBtns() As tCustomButton 'array to hold swapped buttons
    
    For Btn = 4 To UBound(cButtons)
        ' find the user button to position
        If cUserBtn(Btn).ID = ButtonID Then
            newUserBtns = cUserBtn
            ' found it
            Select Case Btn - NewPosition
            Case 0 ' if it's position is already set, nothing more to do
                PositionButton = True
                Exit Function
            Case Is < 0 ' btn is forward of the position
                ' slide the previous buttons backward
                For I = Btn + 1 To NewPosition
                    newUserBtns(I - 1) = cUserBtn(I)
                    cButtons(I - 1) = cButtons(I)
                Next
            Case Else ' btn is backward of the position
                ' slide the next buttons forward
                For I = NewPosition + 1 To Btn
                    newUserBtns(I) = cUserBtn(I - 1)
                    cButtons(I) = cButtons(I - 1)
                Next
            End Select
            
            ' now simply update the button's position
            newUserBtns(NewPosition) = cUserBtn(Btn)
            cButtons(NewPosition) = cButtons(Btn)
            cUserBtn = newUserBtns
            Erase newUserBtns
            PositionButton = True
            Exit For
        
        End If
    Next
    RefreshWindow
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.ModifyButtonTip
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Allows user to change the tooltip text on a custom button
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function ModifyButtonTip(ByVal ButtonID As Long, ByVal ToolTipText As String) As Boolean

    Dim btnID As Long, cRect As RECT, btnRect As RECT
    
    For btnID = 4 To UBound(cButtons)
        
        If cUserBtn(btnID).ID = ButtonID Then
            If ToolTipText = "" Then
                ' if this button had a toolwindow/class, destroy it
                Set cUserBtn(btnID).tTipClass = Nothing
                ModifyButtonTip = True
            
            Else
                With cUserBtn(btnID)
                    If .tTipClass Is Nothing Then
                        ' this button had no previous tooltip; therefore no previous class
                        ' to position the toolwindow, we need to convert the button
                        ' coodinates to clientrect coordinates
                        
                        ' Get our modified client & window rect coordinates
                        CopyMemory cRect, ByVal sharedProperty([_clientRect]), &H10
                        CopyMemory btnRect, ByVal sharedProperty([_winRect]), &H10
                        ' convert the client's screen coordinates to window coordinates
                        OffsetRect cRect, -btnRect.Left, -btnRect.Top
                        ' now convert the button's window coordinates to client coordinates
                        btnRect = cButtons(btnID).Loc
                        OffsetRect btnRect, -btnRect.Left + (btnRect.Left - cRect.Left), -btnRect.Top + (btnRect.Top - cRect.Top)
                        ' create a new class & create the toolwindow
                        Set .tTipClass = New clsToolTip
                        ModifyButtonTip = .tTipClass.CreateToolTip(cHwnd, ToolTipText, VarPtr(btnRect))
                        
                    Else ' easy. Already had a class, just update the tooltip text
                        .tTipClass.TipText = ToolTipText
                        ModifyButtonTip = True
                    End If
                End With
                
            End If
            Exit For
            
        End If
    Next

End Function
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.RotateSystemIcon
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Allows user to create a rotated/unrotated system icon
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Let RotateSystemIcon(ByVal Rotation As eImgRotation)
    
    ' the RotateAuto value will rotate the icon with same orientation as titlebar caption
    
    If Rotation < Rotate0 Or Rotation > RotateAuto Then Exit Property
    
    Dim hSysImg As Long, tBarAlignment As Long
    
    ' clear any previously rotated image & reset option flags
    If hRotatedImage <> 0 Then
        DestroyIcon hRotatedImage
        hRotatedImage = 0
    End If
    cOptions = (cOptions And Not 32) And Not 64
    ' remove the auto rotate flag if needed
    If Rotation < RotateAuto Then cOptions = cOptions And Not 128
    
    If Rotation > Rotate0 Then
        If (cOptions And 128) = 128 Or Rotation = RotateAuto Then
            ' then auto rotation requested
            cOptions = cOptions Or 128  ' add the auto rotate flag
            If cHwnd = 0 Then Exit Property
            tBarAlignment = sharedProperty([_tbarAlignment])
            If tBarAlignment < barAlignBottom Then
                ' titlebar is vertically aligned, set the required rotation
                If tBarAlignment = barAlignLeft Then Rotation = Rotate90 Else Rotation = Rotate270
            Else
                ' auto rotation but titlebar is not rotated. Nothing to do
                Exit Property
            End If
        Else    ' specified to rotate at a set degree (either 90 or 270)
            If Rotation = Rotate270 Then
                cOptions = cOptions Or 64
            Else
                cOptions = cOptions Or 32
            End If
        End If
        ' get the icon handle
        If cHwnd = 0 Then Exit Property
        hSysImg = SendMessage(cHwnd, WM_GETICON, 0, ByVal 0&)
        If hSysImg Then
            hSysImg = RotateImage(hSysImg, Rotation)
            If hSysImg <> 0 Then hRotatedImage = hSysImg
        End If
    End If
End Property
Public Property Get RotateSystemIcon() As eImgRotation
    If (cOptions And 128) = 128 Then
        RotateSystemIcon = RotateAuto
    ElseIf (cOptions And 64) = 64 Then
        RotateSystemIcon = Rotate270
    ElseIf (cOptions And 32) = 32 Then
        RotateSystemIcon = Rotate90
    Else
        RotateSystemIcon = Rotate0
    End If
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.RefreshIcon
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Called whenever a user Sets a new system icon
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Sub RefreshIcon()

    Dim tBarAlignment As Long
    Dim newRotation As eImgRotation
    
    If (cOptions And 128) = 128 Then ' auto rotation
        tBarAlignment = sharedProperty([_tbarAlignment])
        If tBarAlignment > barAlignBottom Then
            ' title bar is horizontally aligned
            If hRotatedImage = 0 Then Exit Sub  ' nothing to do here
        End If
        newRotation = RotateAuto            ' rotate icon appropriately
        '^^ Always. New icon, need new rotated icon
    Else
        ' if system icon is rotated, rotate the new system icon
        If (cOptions And 64) = 64 Then
            newRotation = Rotate270
        ElseIf (cOptions And 32) = 32 Then
            newRotation = Rotate90
        End If
    End If
    If newRotation > Rotate0 Then Me.RotateSystemIcon = newRotation
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.OwnerDrawn
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Pass owner drawn flags
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Sub OwnerDrawn(cImplementation As Long, odItems As Long, isClosing As Boolean)

    If isClosing Then
        mODImplementation = 0
        mNotifyUser = 0
        CreateTimer False, 0        ' remove any existing timer
    Else
        mNotifyUser = mNotifyUser And 1024 ' carry over custom button owner drawn flag
        If (odItems And odSysButtons) = odSysButtons Then mNotifyUser = mNotifyUser Or odSysButtons
        If mNotifyUser = 0 Then
            mODImplementation = 0
            CreateTimer False, 0        ' remove any existing timer
        Else
            mODImplementation = cImplementation
        End If
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.UpdateUserToolTips
' DateTime  : 9/20/2005
' Author    : LaVolpe
' Purpose   : Position tooltip windows so they overlap the user buttons
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Friend Sub UpdateUserToolTips(ByVal clientPtr As Long)

    ' Tooltip windows are in relation to the client rectangle's top left corner
    ' When we resize our window, the top/left client coordinates may not move
    ' but our buttons may move horizontally or vertically. The tooltip API is
    ' not truly designed for tracking non-client regions, so we need to fudge
    ' a bit and simply tell the tooltip API where our toolwindow should now be
    
    ' Note that calling this function when a window moves is not necessary, only
    ' when the window changes size
    
    If UBound(cButtons) = 3 Then Exit Sub   ' no user buttons
    
    Dim btnID As Long
    Dim btnRect As RECT, clientRect As RECT

    CopyMemory clientRect, ByVal clientPtr, &H10

    For btnID = 4 To UBound(cButtons)
        If Not cUserBtn(btnID).tTipClass Is Nothing Then
            btnRect = cButtons(btnID).Loc
            ' convert the nonclient coordinates to client coordinates
            OffsetRect btnRect, -btnRect.Left + (btnRect.Left - clientRect.Left), -btnRect.Top + (btnRect.Top - clientRect.Top)
            ' call class property to update the toolwindow's position
            cUserBtn(btnID).tTipClass.TipRect = VarPtr(btnRect)
        End If
    Next

End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.DisableRestore
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Enables/disables the Restore button when window is maximized
' Comments  : See property CustomWindow.MaximizeFullScreen
'---------------------------------------------------------------------------------------
'
Friend Sub DisableRestore(isMaximized As Boolean)
    If (cOptions And 2) = 2 Then
        EnableSysMenuItem SC_RESTORE_Clone, Abs(isMaximized = False)
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.NoRedraw
' DateTime  : 9/15/2005
' Author    : LaVolpe
' Purpose   : Prevents redrawing while multiple properties are being changed
'---------------------------------------------------------------------------------------
'
Public Property Let noRedraw(bNoRedraw As Boolean)
    cAutoRedraw = Not bNoRedraw
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.RefreshWindow
' DateTime  : 9/15/2005
' Author    : LaVolpe
' Purpose   : Forces window to repaint
'---------------------------------------------------------------------------------------
'
Private Sub RefreshWindow()
    If cHwnd Then
        If cAutoRedraw Then ForceRefresh cHwnd
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.Class_Initialize
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Initialization/Destruction routines
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    cBtnSize = GetSystemMetrics(SM_CYSMICON)
    ReDim cButtons(0 To 3)
    Set cColors(0) = New clsButtonColors
    Set cColors(1) = New clsButtonColors
    cColors(1).pUsage = 1
End Sub

Private Sub Class_Terminate()
    CreateTimer False, 0
    If hRotatedImage <> 0 Then DestroyIcon hRotatedImage
End Sub


'---------------------------------------------------------------------------------------
' Procedure : clsButtons.sharedProperty
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


