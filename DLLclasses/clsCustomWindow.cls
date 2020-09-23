VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' This is the interface for the custom window DLL that users will have access to
' Any other classes are not exposed directly.

' A summary of changes and known bugs/issues can be found in iCoreMessages

' You will notice there are 4 Implementations (interface classes) for you to
' pick and choose from? Why not combine them all into one interface? Speed.

' By breaking the interfaces out into logical operations, the DLL can check to
' see if forwarding a message or command is needed, vs always forwarding everything
' even though the user may of wanted just one little notification.

' Much more detail will be provided in a HowTo.RTF file I am writing/re-writing.
' Improvements over the 1st version:
' - Complete, on-demand, non-client modifications (very few restrictions)
' - Far less subclassing used & less "noisy" (very little extra generated messages)
'    v1 @ http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57406&lngWId=1
' - Built-in capability to owner-draw the non-client routines
'       Must implement iOwnerDrawn
' - Built-in message forwarding so you can use the subclassed messages for
'       other purposes like owner-drawn popup menus.
'       Must implement iCoreMessages
' - Built-in capability to subclass your window's child controls
'       Must implement iControlsCallback
' - Added built-in system tray icon support
'       Must implement iSysTrayCallback
' - Exposed common graphic routines useful for drawing/measuring captions and
'       backgrounds (including stretching,tiling,etc) & more.
'       See cGraphics. That class not used by the DLL, it only exists for the user
' - IMO, much better use of windows messages to produce the desired effect:
'       stop VB from repainting the non-client (NC) area, with minimal intervention
' - Finally broke the code to allow O/S to move/size the window so I don't have to.
'   This is huge! By moving it ourselves we needed to write much more code & the
'   downside was trails while dragging the size. Now the O/S does the proper
'   thing, assuming by use of copy bits, to repaint the uncovered desktop
'   Note that I cannot use SendMessage with WM_NCLButtonDown & HTCaption as this
'   will trigger VB to repaint the titlebar and other NC areas. In fact, letting
'   any WM_NCLButtonDown will allow VB to repaint, so need some workarounds
'   FYI: Here's what I discovered.....
'       Moving: sending WM_SYSCOMMAND with SC_Move will move a window but will
'           place the cursor in the expected titlebar even if it is somewhere else,
'           however, passing (SC_Move + HTCaption) will move the window
'           using the current cursor pos
'       Sizing: sending WM_SYSCOMMAND with SC_Size will send the cursor to the
'           middle of the client area and wait for you to move the cursor to a
'           border location to begin sizing. But, if you supply the border HTflag
'           (i.e., HTBottom, HTLeft, etc) with the SC_Size parameter then the O/S
'           will leave your curosr in place and size the window appropriately.
'           The trick is to subtract 9 from the HT flag. For example, HTleft=10.
'           To use SC_Size, we would supply SC_Size + (HTLeft - 9)
' Limitations
'   .. No MDI support. Do not rem out the checks for MDI forms
'       However, you can still sublcass MDI forms for your own need by using
'       the  Subclass_OtherWindow function
'   .. Do not change your caption via Me.Caption, either use the Caption property
'       of the Titlebar class, SendMessage with a WM_SETTEXT message, or SetWindowText.
'       Otherwise, VB will redraw internally & I cannot trap this message to
'       prevent it from happening

'///////////////////////////////////////////////////////////////////////////////////////

Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32.dll" (ByRef lpRect As RECT) As Long
Private Declare Function DrawAnimatedRects Lib "user32.dll" (ByVal hWnd As Long, ByVal idAni As Long, ByRef lprcFrom As RECT, ByRef lprcTo As RECT) As Long

Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Function GetWindowPlacement Lib "user32.dll" (ByVal hWnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetMenu Lib "user32.dll" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As Long, ByRef lpTPMParams As TPMPARAMS) As Long

Private Const SPI_GETWORKAREA = &H30
Private Const WS_MINIMIZE As Long = &H20000000
Private Const WS_MAXIMIZE As Long = &H1000000

Private Type NCCALCSIZE_PARAMS  ' used when client rect needs to be calculated
   rgrc(0 To 2) As RECT
   lppos As Long    'pointer to WINDOWPOS
End Type
Private Type TPMPARAMS           ' used with TrackpopupMenuEX
    cbSize As Long
    rcExclude As RECT
End Type
Private Type WINDOWPOS         ' used when client rect needs to be calcualted
    hWnd As Long
    hWndInsertAfter As Long
    x As Long
    Y As Long
    Cx As Long
    Cy As Long
    flags As Long
End Type
Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Public Enum eBarAlignment   ' used for titlebar/menubar orientation
    barAlignLeft = 10
    barAlignRight = 11
    barAlignBottom = 12
    barAlignTop = 13
End Enum
Public Enum eBarResets      ' used for system default resets for all classes
    rstTextColors = 2       ' applies to menus & titlebars
    rstMenuSelect3Dbox = 4  ' only applies to the menubar colors class
    rstBackGroundColors = 8 ' applies to menus and titlebars
    rstNoImage = 16         ' applies to menus, titlebars, and the NC inset
    rstFont = 32            ' applies to menus and titlebars
    rstBorderColors = 64    ' only applies to the borders class
    rstButtonColors = 128   ' only applies to the Buttons class
    rstBorderSizes = 256    ' applies to all: menubars,titlebars,frames,nc inset
    rstAll = -1
End Enum
Public Enum e3DBox          ' types of menu item selection
    bxHover = 1
    bxSelect = 2
End Enum
Public Enum e3Dframe        ' options for framing menubar & titlebar
    bxNone = 0
    bxFlat = 1
    bx3D = 2
End Enum
Public Enum eSysButtons
    SC_CLOSE = &HF060&        ' all are used to detect sysmenu actions
    SC_MAXIMIZE = &HF030&
    SC_MINIMIZE = &HF020&
    SC_MOVE = &HF010&
    SC_SIZE = &HF000&
    SC_RESTORE = &HF120&
End Enum
Public Enum eEnableConstants
    sysDisable = 0
    sysEnable = 1
End Enum
Public Enum eMnuForeColors  ' used for menu item styles
    mcStandard = 1
    mcHover = 2
    mcSelect = 4
    [_mcAllActions] = 7
End Enum
Public Enum eMnuBorders     ' used to assign colors to 3D menu item boxes
    mbDefault = 0
    mbRaisedCustom = 1
    mbSunkenCustom = 2
    mbFilledCustom = 3
    mbNoBorders = 4
End Enum
Public Enum eBackStyles     ' used for image painting options
    bsSmartStretch = 0
    bsStretch = 1
    bsTiled = 2
End Enum
Public Enum eFX         ' used for menu item styles
    fxDefault = 0
    fxFlat = 1
    fxSunken = 2
    fxRaised = 3
    fxEngraved = 4
End Enum
Public Enum eOwnerDrawn
    odMenuItem_ImageOnly = 1
    odMenuItem_Complete = 2
    odMenuBarBkg = 4
    odTitlebarBkg = 8
    odSysButtons = 16
    odPostNCDrawing = 32
    odFrameBorders = 64
    ' 128 & 512 reserved for future if needed; can't think of anything right now
    '1024 = custom buttons; set automatically when buttons are added
    '2048 = tracking rectangles; set automatically when those are set
End Enum
Public Enum eDTflags
    DT_CENTER = &H1
    DT_LEFT = &H0
    DT_NOCLIP = &H100
    DT_NOPREFIX = &H800
    DT_RIGHT = &H2
    DT_SINGLELINE = &H20
    DT_VCENTER = &H4
    DT_HIDEPREFIX = &H100000
    DT_WORDBREAK = &H10
    DT_WORD_ELLIPSIS = &H40000
    DT_MODIFIABLE = &H10000
End Enum
Public Enum eFillStyle
    bfSolid = 0
    bfGradientEW = 1
    bfGradientNS = 2
    bfTransparent = 3
    bfGrayScaled = 32
End Enum
Public Enum eGFont
    gDC_Font = 0
    gMenu_Font = 1
    gTitlebar_Font = 2
End Enum
Public Enum eDFCtype
    ' I'm using a custom name vs the DFCS proper name,
    ' only cause I think the proper names might be confusing to some
    dfcBlankButton = &H10 ' DFCS_BUTTONPUSH = &H10
    dfcCaptionButton = &H1 'DFC_CAPTION=1
End Enum
Public Enum eDFCstate
    DFCS_CAPTIONCLOSE = &H0
    DFCS_CAPTIONHELP = &H4
    DFCS_CAPTIONMAX = &H2
    DFCS_CAPTIONMIN = &H1
    DFCS_CAPTIONRESTORE = &H3
    DFCS_HOT = &H1000
    DFCS_INACTIVE = &H100
    DFCS_PUSHED = &H200
End Enum
Public Enum eImgRotation
    Rotate0 = 0
    Rotate90 = 1
    Rotate270 = 2
    RotateAuto = 3
End Enum
Public Enum eSysIconOptions
    eSysIconNormal = 0
    eSysIconHidden = 1
    eSysMenuHidden = 2
End Enum
Public Enum eODmessageConstants
    omMeasureMenuItem = 0
    omDrawMenuItem = 1
    omDrawUserBtn = 2
    omDrawSysBtn = 3
    omDrawMBarBkg = 4
    omDrawTBarBkg = 5
    omDrawPostNC = 6
    omDrawPreNC = 7
    omDrawTrackRect = 8
    omTrackCursor = 9
    omUserBtnClick = 10
    omTrackClick = 11
End Enum
Public Enum eExportConstants
    exStylesColors = 0  ' always exported/imported
    exBkgImages = 1     ' exclude if menubar/titlebar/inset images not to be exported
    exSysTray = 2       ' exclude if system tray icon info not to be exported
    exCustomButtons = 4 ' exclude if userbuttons not to be exported
    extTrackRects = 8
    exAll = exBkgImages Or exSysTray Or exCustomButtons Or extTrackRects
End Enum

Public Enum eDLLconstants
    ' a way/cheat to make other constants public
    ' since you can't use Public Const statements in a non-global class
    odImgNone = 0
    odImgAlignLeft = 8
    odImgAlignRight = 4
    mcDisabled = 8
    mcRightButton = 16
End Enum
Public Enum eMin2TrayOptions
    SIZE_MAXIMIZED = 2
    SIZE_MINIMIZED = 1
    SIZE_RESTORED = 0
End Enum

Private mClientImplementation As Long
Private mControlsImplementation As Long
Private mTrayImplementation As Long

'////////////// CUSTOM CONSTANTS/VARIABLES \\\\\\\\\\\\\\\
'====== Subclassing Instance
Private mWinRect As RECT        ' dimensions of current window
Private mClientRect As RECT     ' dimensions of client area in screen coords
Private mMovementFlag As Long   ' htCaption=moving;htLeft to htBottomRight=sizing;htMin=titlebar button;WM_LButtonDown
Private mWinMenu As Long        ' cached 'cause we remove & replace as needed
Private mHWnd As Long           ' form-level hWnd being subclassed
Private mAutoRedraw As Boolean  '
Private mNotifyUser As Long     ' option for user to control drawing/subclassing
Private mOptions As Byte        ' 2=minimize to systray,
Private mActiveState As Byte    ' option to show as active vs inactive (same thread)
    '0=really inactive
    '1=really active
    '2=override option set
    '4=appear always active in current thread only
    '8=appear active always
    '16=adjusted state is active else inactive
Private mClientInset As Byte
'1=left inset,2=top inset,4=right inset,8=bottom inset, 0=no inset

' messages to forward to user by request
Private cClientMsgs() As Long
Private cMsg_NChitTest As Byte  ' cached request vs lookup
Private cMsg_SetCursor As Byte  ' cached request vs lookup


Private mControls As Collection ' collection of control hWnds also being subclassed
Private mSelfModifying As Boolean ' flag to prevent excessive NC calculations (see ToggleWSCaption routine)
Private mWindowState As Long      ' same as VB.WindowState, except OR'd w/4 indicates minimized to system tray
' child classes
Private mMenuBar As clsMenubar
Private mTitleBar As clsTitlebarD
Private mFrame As clsFrame
Private mGraphics As cGraphics
' callback for child classes to reference this class' properties
Implements iImpProps

'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.BeginCustomize
' DateTime  : 8/27/05 22:25
' Author    : LaVolpe
' Purpose   : function is called to begin customizing the window
' Comments  : Any other initial properties should have been set, but can be set at any time
'---------------------------------------------------------------------------------------
'
Public Function BeginCustomize(ByVal hWnd As Long) As Boolean
    
    Dim bResult As Boolean, hIcon As Long
    
    ' start subclassing
    mHWnd = hWnd    ' cache the hWnd
    ' initialization of any uninitialized classes will occur
    ' while calling Me.Titlebar.[anything]
    Me.Titlebar.Buttons.pEnabled SC_CLOSE
    ' ensure we track the window's sizes for our global DC
    AddClient mHWnd, 0, 0
    ExpandDC 16, 16, False, mHWnd ' set absolute minimum size of the menuitem DC
    ' subclass the window
    bResult = SubclassWindow(hWnd, ObjPtr(Me), False)
    If bResult Then
        InitialiazeSkin
        mWinMenu = 0
        RemoveFormMenu
        SetSmallIcon mHWnd
    Else    ' something's wrong, probably being subclassed already or MDI
        mHWnd = 0
        RemoveClient hWnd
    End If

    BeginCustomize = bResult    ' return the result
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.Subclass_OtherWindow
' DateTime  : 8/27/05 22:25
' Author    : LaVolpe
' Purpose   : function to subclass individual controls
' Comments  : Pass any controls you want to receive messages for
'---------------------------------------------------------------------------------------
'
Public Function Subclass_OtherWindow(iImplementation As iControlsCallback, ByVal hWnd As Long) As Boolean
' function to subclass individual controls

    ' not allowed if the main-level form is not subclassed already
    If mHWnd = 0 Then Exit Function
    If iImplementation Is Nothing Then Exit Function
    
    Dim bResult As Boolean
    
    ' initialize the controls hWnd collection
    If mControls Is Nothing Then Set mControls = New Collection
    ' attempt to subclass
    mControlsImplementation = ObjPtr(iImplementation)
    bResult = SubclassWindow(hWnd, ObjPtr(Me), True)
    
    On Error Resume Next
    If bResult Then
        'add hWnd to the collection
        mControls.Add hWnd, "h" & hWnd
    Else
        mControls.Remove "h" & hWnd
    End If
    On Error GoTo 0
    
    ' return result
    Subclass_OtherWindow = bResult
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.Unsubclass_OtherWindow
' DateTime  : 8/27/05 22:26
' Author    : LaVolpe
' Purpose   : un-subclasses a control, not the main-level window
' Comments  : see below
'---------------------------------------------------------------------------------------
'
Public Sub Unsubclass_OtherWindow(ByVal hWnd As Long)
' The only way the main-level window can be unsubclassed is:
'   1. Setting this class to Nothing
'   2. Terminating the window normally (i.e., don't use END statement or hit the End button)
    
    If hWnd = 0 Then Exit Sub
    ' don't allow the user to pass the main-level hWnd to this routine
    If hWnd <> mHWnd Then SubclassTerminated hWnd

End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.ForwardMessage
' DateTime  : 8/27/05 22:26
' Author    : LaVolpe
' Purpose   : this is the user's CallWndProc function
' Comments  : see below
'---------------------------------------------------------------------------------------
'
Public Function ForwardMessage(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' The user won't know, the previous GWL_WNDPROC value
' for the subclassed window or control, required in the CallWindowProc.
' Therefore, we provide this function for that purpose if desired

Dim hWndproc As Long

On Error Resume Next
If hWnd Then
    hWndproc = GetProp(hWnd, lvWndProcName)
    If hWndproc Then
        ForwardMessage = CallWindowProc(hWndproc, hWnd, wMsg, wParam, lParam)
    End If
End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.RefreshWindow
' DateTime  : 8/27/05 22:27
' Author    : LaVolpe
' Purpose   : called by user to apply any new options/styles
' Comments  : Optional & called by user after changing multiple settings
'---------------------------------------------------------------------------------------
'
Public Sub RefreshWindow()
    ForceRefresh mHWnd
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.ResetToSystemDefaults
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Proxy routine to pass the request to the appropriate subclasses
' Comments  : Setting values at this level prevents having to call the individual levels
'---------------------------------------------------------------------------------------
'
Public Sub ResetToSystemDefaults(Options As eBarResets)
    Me.Titlebar.ResetToSystemDefaults Options
    Me.MenuBar.ResetToSystemDefaults Options
    Me.Frame.ResetToSystemDefaults Options
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.Titlebar
' DateTime  : 8/27/05 22:28
' Author    : LaVolpe
' Purpose   : Exposes the titlebar class
'---------------------------------------------------------------------------------------
'
Public Property Get Titlebar() As clsTitlebarD
    If mTitleBar Is Nothing Then
        ' when the titlebar is Nothing, re-instantiate it & the other key classes too
        Set mTitleBar = New clsTitlebarD
        mTitleBar.pMainImplementation(mHWnd) = ObjPtr(Me)
        Me.MenuBar.HideDisabledItems = Me.MenuBar.HideDisabledItems
    End If
    Set Titlebar = mTitleBar
End Property
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.Frame
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Exposes the Frame class
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Get Frame() As clsFrame
    If mFrame Is Nothing Then
        Set mFrame = New clsFrame
        mFrame.pMainImplementation(mHWnd) = ObjPtr(Me)
    End If
    Set Frame = mFrame
End Property
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.MenuBar
' DateTime  : 8/27/05 22:29
' Author    : LaVolpe
' Purpose   : Exposes the menubar class
'---------------------------------------------------------------------------------------
'
Public Property Get MenuBar() As clsMenubar
    If mMenuBar Is Nothing Then
        Set mMenuBar = New clsMenubar
        mMenuBar.pMainImplementation(mHWnd) = ObjPtr(Me)
    End If
    Set MenuBar = mMenuBar
End Property
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.AmIActive
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Returns 1 if the window is painted as Active, otherwise returns 0
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Get AmIActive() As Byte
    If (mActiveState Or 1) = mActiveState Then
        AmIActive = 1
    ElseIf (mActiveState Or 2) = mActiveState Then
        If (mActiveState Or 16) = mActiveState Then AmIActive = 1
    End If
End Property
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.Graphics
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Exposes the Graphics class
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Get Graphics() As cGraphics
    If mGraphics Is Nothing Then
        Set mGraphics = New cGraphics
        mGraphics.pImplementation = ObjPtr(Me)
    End If
    Set Graphics = mGraphics
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsMenubar.NoRedraw
' DateTime  : 9/5/2005
' Author    : LaVolpe
' Purpose   : Prevents redrawing while multiple properties are being changed
'---------------------------------------------------------------------------------------
'
Public Property Let noRedraw(bNoRedraw As Boolean)
    mAutoRedraw = Not bNoRedraw
    Me.Titlebar.noRedraw = bNoRedraw
    Me.MenuBar.noRedraw = bNoRedraw
    Me.Frame.noRedraw = bNoRedraw
End Property
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.isOwnerDrawn
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Returns whether or not the passed parameter indicates an owner-drawn object
'---------------------------------------------------------------------------------------
'
Public Function isOwnerDrawn(ByVal odItem As eOwnerDrawn) As Long

    If odItem < odMenuItem_ImageOnly Or odItem > odFrameBorders * 2 - 1 Then
        isOwnerDrawn = mNotifyUser
    Else
        isOwnerDrawn = (mNotifyUser And odItem)
    End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.OwnerDrawn
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Sets whether or not objects are owner-drawn
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub OwnerDrawn(cImplementation As iOwnerDrawn, odItems As eOwnerDrawn)
    
    If odItems > odFrameBorders * 2 - 1 Then
        ' prevent user from passing values > 1027 (custom user-defined titlebar buttons)
        ' 1024 & 2048 are hidden, DLL-use only flags
        Exit Sub
    Else
        If odItems < 0 Then Exit Sub
    End If
    
    Dim mODImplementation As Long
    If cImplementation Is Nothing Or odItems = 0 Then
        mODImplementation = 0
    Else
        mODImplementation = ObjPtr(cImplementation)
    End If
    mNotifyUser = odItems
    Me.Titlebar.OwnerDrawn mODImplementation, odItems, (cImplementation Is Nothing)
    Me.MenuBar.OwnerDrawn mODImplementation, odItems, (cImplementation Is Nothing)
    Me.Frame.OwnerDrawn mODImplementation, odItems, (cImplementation Is Nothing)

End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.ProcessMessage
' DateTime  : 8/27/05 22:29
' Author    : LaVolpe
' Purpose   : This is the subclasser, processing & passing windows messages
' Comments  : see below (the key message receipt point)
'---------------------------------------------------------------------------------------
'
Friend Function ProcessMessage(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal hOldWndProc As Long) As Long

' How do I know what each message does or how to use each one?
'   www.msdn.com for how, why & when
'   spy++ to get hints on which ones are firing
'   a lot of trial & error & crashing :)

' Note about the custom function used: ToggleWSCaption hwnd, True/False
'   The ToggleWSCaption function tricks VB into thinking it doesn't have a titlebar.
'   See that function in modSC for more info

' 20 Sep 05. Modified structure quite a bit. Broke messages into sections to be
'   executed when window is minimized and when it is not. I am handing off 99.9%
'   of the messages to windows, unmolested, when the window is minimized.

    Dim lOverrideReturn As Long, bOverridden As Boolean
    Dim wPTs As POINTAPI ' common use for converting lParam to X,Y coords
    ' values used for PostMessage API
    Dim postWparam As Long, postLparam As Long, postWmsg As Long
    Dim isMinimized As Boolean

    On Error Resume Next

    If hWnd <> mHWnd Then ' must be a child control that is subclassed
        ' pass the message to the proxy function
        ProcessMessage = ControlProxy(hWnd, wMsg, wParam, lParam, hOldWndProc)
        Exit Function
    End If
    
'        Debug.Print wMsg; wParam, WM_TrayNotify
    isMinimized = ((mWindowState = SIZE_MINIMIZED) Or ((mWindowState And 4) = 4))

Select Case wMsg

    ' these messages are processed first, so the user will know the
    ' proper hit test code. Then they are forwarded to the user as needed

    Case WM_NCHITTEST ' Returns the hit test the cursor is currently over
        
        If isMinimized Then
            ' would only get this if our window is minimized off the task bar
            lOverrideReturn = DefWindowProc(hWnd, wMsg, wParam, ByVal lParam)
        Else
            bOverridden = True
            wPTs.x = LoWord(lParam) ' convert lParam to PointAPI
            wPTs.Y = HiWord(lParam)
            If PtInRect(mClientRect, wPTs.x, wPTs.Y) Then
                ' if the cursor is in the client area, pass the message thru
                 lOverrideReturn = DefWindowProc(hWnd, wMsg, wParam, ByVal lParam)
            Else ' what part of the nonclient area is the cursor in?
                ' test the menubar
                lOverrideReturn = mMenuBar.HitTest(wPTs.x - mWinRect.Left, wPTs.Y - mWinRect.Top)
                If lOverrideReturn = 0 Then ' test titlebar & buttons
                    lOverrideReturn = mTitleBar.HitTest(wPTs.x - mWinRect.Left, wPTs.Y - mWinRect.Top)
                    If lOverrideReturn = 0 Then ' test frame borders & custom inset(s)
                        lOverrideReturn = mFrame.HitTest(wPTs.x, wPTs.Y, VarPtr(mWinRect), VarPtr(mClientRect))
                    End If
                End If
            End If
        End If
        ' inform user of the ncHitTest; user cannot override this message
        If cMsg_NChitTest Then
            Select Case lOverrideReturn
            Case HTNC_Caption, HTNC_Custom
                SendCallBack hWnd, wMsg, wParam, lParam, True, HTNOWHERE + 0
            Case HTMenu_Custom
                SendCallBack hWnd, wMsg, wParam, lParam, True, HTMENU + 0
            Case Else
                SendCallBack hWnd, wMsg, wParam, lParam, True, lOverrideReturn + 0
            End Select
        End If
        
    Case WM_SETCURSOR ' sets the cursor based on the our custom hit test values
        ' ensure we set the cursor for our custom hit tests. We don't want VB getting
        ' this message when in the NC 'cause VB will redraw the min/max button group
        If mMovementFlag = 1 Then
            ' when accessing Move from the sysMenu using the keyboard, we fudge
            ' a bit to ensure the 4-star icon is shown as expected.
            SetCursor commonCursor(5)
            bOverridden = True
        Else
            lOverrideReturn = LoWord(lParam)
            ' allow user to override the cursor setting
            If isMinimized Then
                If cMsg_SetCursor Then SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn + 0
            Else
                Select Case lOverrideReturn
                    Case HTMenu_Custom
                        If cMsg_SetCursor Then SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, HTMENU + 0
                        If Not bOverridden Then SetCursor commonCursor(0)
                        bOverridden = True
                    Case HTCLIENT, HTERROR, HTBORDER, HTNOWHERE
                        If cMsg_SetCursor Then SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn + 0
                    Case HTLEFT To HTBOTTOMRIGHT
                        If mTitleBar.Buttons.pEnabled(SC_SIZE) = False Then
                            If cMsg_SetCursor Then SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn + 0
                            If Not bOverridden Then SetCursor commonCursor(0)
                        Else
                            Select Case lOverrideReturn
                            Case HTBOTTOMLEFT, HTTOPRIGHT:
                                If cMsg_SetCursor Then SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn + 0
                                If Not bOverridden Then SetCursor commonCursor(1)
                            Case HTBOTTOM, HTTOP:
                                If cMsg_SetCursor Then SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn + 0
                                If Not bOverridden Then SetCursor commonCursor(2)
                            Case HTBOTTOMRIGHT, HTTOPLEFT:
                                If cMsg_SetCursor Then SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn + 0
                                If Not bOverridden Then SetCursor commonCursor(3)
                            Case HTRIGHT, HTLEFT:
                                If cMsg_SetCursor Then SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn + 0
                                If Not bOverridden Then SetCursor commonCursor(4)
                            End Select
                        End If
                        bOverridden = True
                    Case HTNC_Tracker
                        If mFrame.SetMouseAction(wMsg, AmIActive) = False Then SetCursor commonCursor(0)
                        bOverridden = True
                    Case Else: ' could be our HTNC_Custom flag or something else
                        ' regardless, tell the user the hit test is nowhere
                        If cMsg_SetCursor Then SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, HTNOWHERE + 0
                        If Not bOverridden Then SetCursor commonCursor(0)
                        bOverridden = True
                End Select
                If mMenuBar.pTracking(mbActiveStatus) Then
                    Select Case HiWord(lParam)
                    ' ensure a button down anywhere is handled if we are in a menu loop
                    Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_XBUTTONDOWN
                        mMenuBar.SetMouseAction HTNOWHERE, 0, WM_NCLBUTTONDOWN
                    End Select
                End If
            End If
        End If
        
    Case WM_LVPopup ' menu about to be shown. Allow user to abort
        
        Dim tpParam As TPMPARAMS
        
        If mClientImplementation <> 0 Then
            If mMenuBar.pTracking(mbPosition) < 0 Then 'system menu
                postLparam = MakeLong(0, 1)
            Else
                postLparam = mMenuBar.pTracking(mbPosition)
            End If
            SendCallBack hWnd, WM_INITMENUPOPUP, mMenuBar.pTracking(mbID), postLparam, bOverridden, lOverrideReturn
        End If
        If bOverridden Then
            ' we hooked menu messages and/or keyboard in preparation
            ' for this menu. User cancelled it, therefore we release the hooks
            ReleaseHooks
        Else
        
            tpParam.cbSize = Len(tpParam)
            bOverridden = True
        
            If mMenuBar.pTracking(mbPosition) < 0 Then ' sysMenu
                lOverrideReturn = 0
                ' the flags may have extra meaning when the titlebar is vertical. Check & remove the flags
                If (mMenuBar.pTracking(mbFlags) Or TPM_BOTTOMALIGN) = mMenuBar.pTracking(mbFlags) Then lOverrideReturn = TPM_BOTTOMALIGN
                If (mMenuBar.pTracking(mbFlags) Or TPM_RIGHTALIGN) = mMenuBar.pTracking(mbFlags) Then lOverrideReturn = lOverrideReturn Or TPM_RIGHTALIGN
                ' remove the extra flags now
                mMenuBar.pTracking(mbFlags) = mMenuBar.pTracking(mbFlags) And Not lOverrideReturn
            End If
        
            lOverrideReturn = TrackPopupMenuEx(mMenuBar.pTracking(mbID), lOverrideReturn Or TPM_RETURNCMD, _
                        wParam + mWinRect.Left, lParam + mWinRect.Top, hWnd, tpParam)
            
            If lOverrideReturn = 0 Then
                If (mMenuBar.pTracking(mbFlags) And (SC_KEYMENU Or SC_MOUSEMENU)) = (SC_KEYMENU Or SC_MOUSEMENU) Then
                    ' user used enter to close the menu, terminate the menu loop
                    ' This is only received is user did Enter on a disabled item
                    ReleaseHooks
                End If
            Else
                ' any selected menu item from the menbar will be returned here
                ' The posted messages will be trapable & forwardable to user
                If mMenuBar.pTracking(mbPosition) < 0 Then ' sys menu item
                   PostMessage hWnd, WM_SYSCOMMAND, lOverrideReturn, ByVal 0&
                Else
                    ' forward non-system the menu selection
                    PostMessage hWnd, WM_COMMAND, MakeLong(lOverrideReturn + 0, 0), ByVal 0&
                End If
                ReleaseHooks
            End If
            
        End If
        
        ' these are never sent to the user
        ' One is to restore the system tray & the other sent to the iSysTrayCallback interface
            
    Case WM_IECrashNotify   ' custom message
        ' Explorer crash & restored; reinitialize tray icons
        If Frame.SysTrayInitialized Then Frame.RestoreSystemTray
        bOverridden = True
        
    Case WM_TrayNotify
        If mTrayImplementation <> 0 Then
            On Error Resume Next
            Dim cClient As iSysTrayCallback
            GetObjectFromPointer mTrayImplementation, cClient
            cClient.ProcessTrayIcon hWnd, wParam, lParam, bOverridden, lOverrideReturn
            Set cClient = Nothing
        End If
        bOverridden = True
    
    Case WM_ENTERMENULOOP
        ToggleWSCaption True, False
        If mClientImplementation <> 0 Then ' forward message to user
            SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn
            If Not bOverridden Then lOverrideReturn = CallWindowProc(hOldWndProc, hWnd, wMsg, wParam, lParam)
        End If
        If Not bOverridden Then
            bOverridden = True
            ' if this was a keyboard activated menu, highlight the 1st menu item
            If (mMenuBar.pTracking(mbFlags) And SC_KEYMENU) = SC_KEYMENU Then
                ' is this a menu we are hooking?
                mMenuBar.pTracking(mbFlags) = mMenuBar.pTracking(mbFlags) And Not SC_KEYMENU
                If mMenuBar.pTracking(mbActiveStatus) > 0 Then PostMessage hWnd, &H100, vbKeyDown, ByVal &H1500001
            End If
        Else
            ToggleWSCaption False, False
        End If
        
    Case WM_GETSYSMENU
        If (mTitleBar.Buttons.SystemIconMenu And eSysMenuHidden) = 0 Then
            ToggleWSCaption True, False
            If mClientImplementation <> 0 Then ' forward message to user
                SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn
                If bOverridden Then ToggleWSCaption False, False
            End If
        Else
            bOverridden = True
        End If
        
    Case WM_EXITMENULOOP
        ' with wm_exitmenuloop: add the ws_caption style back
        ' before a system menu popup appeared, we removed the WS_Caption style
        ToggleWSCaption False, False
        If mClientImplementation <> 0 Then ' forward message to user
            SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn
        End If
        
    Case WM_INITMENUPOPUP
        ToggleWSCaption True, False
        If mClientImplementation <> 0 Then ' forward message to user
            SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn
            If Not bOverridden Then lOverrideReturn = CallWindowProc(hOldWndProc, hWnd, wMsg, wParam, lParam)
        End If
        ToggleWSCaption False, False
        bOverridden = True

    Case Else
    
        If mClientImplementation <> 0 Then
            ' forward message to user
            If hOldWndProc <> 0 Then SendCallBack hWnd, wMsg, wParam, lParam, bOverridden, lOverrideReturn
        End If
        If Not bOverridden Then
        
            If isMinimized Then
                ' only these 2 messages get specifically processed when minimized
                
                ' we don't draw & generally don't mess with the window when it is
                ' minimized. However if user doesn't want the system menu displayed
                ' then prevent it if that option was chosen
                If (mTitleBar.Buttons.SystemIconMenu And eSysMenuHidden) = eSysMenuHidden Then
                    Select Case wMsg
                    Case WM_GETSYSMENU, WM_CONTEXTMENU
                        bOverridden = True
                    'Case wParam = SC_KEYMENU, wPram= SC_MOUSEMENU
                    ' ^^ above WM_SYSCOMMAND wParam values handled too but later
                    End Select
                End If
            
            Else    ' only process these messages if the window is not minimized
                
                Select Case wMsg
            
                Case WM_MOUSEMOVE ' client area mousemove vs WM_NCMOUSEMOVE
                    ' is titlebar button being dragged
                    Select Case mMovementFlag
                        Case Is > HTMenu_Custom, HTCLOSE, HTMINBUTTON, HTMAXBUTTON
                            wPTs.x = LoWord(lParam)
                            wPTs.Y = HiWord(lParam)
                            ClientToScreen mHWnd, wPTs
                            mTitleBar.Buttons.SetMouseAction wMsg, mMovementFlag, wPTs.x - mWinRect.Left, wPTs.Y - mWinRect.Top
                    End Select
                
                Case WM_NCMOUSEMOVE ' NonClient mouse move
                    Select Case wParam
                    Case HTMenu_Custom
                        mMenuBar.SetMouseAction AmIActive, lParam, WM_NCMOUSEMOVE
                    Case HTCLOSE, HTMINBUTTON, HTMAXBUTTON
                        mTitleBar.Buttons.SetMouseAction wMsg, wParam, 0, 0
                    Case Is > HTMenu_Custom
                        mTitleBar.Buttons.SetMouseAction wMsg, wParam, 0, 0
                    Case HTNC_Tracker
                        mFrame.SetMouseAction wMsg, AmIActive
                    End Select
                    
                Case WM_NCPAINT ' never let VB paint after you paint
                    RemoveFormMenu ' with bug fix for wm_ncactivate, menu will
                            ' overdraw window unless checked & removed here
                    DoNCpaint
                    bOverridden = True
                
                Case WM_NCRBUTTONDOWN
                    ' allow right click notifications for NC tracking rectangles
                    If wParam = HTNC_Tracker Then mFrame.SetMouseAction wMsg, AmIActive
                
                Case WM_NCRBUTTONUP
                    ' test to display sysmenu. Needed for vertical titlebars
                    bOverridden = True
                    Select Case wParam
                    Case HTCAPTION, HTSYSMENU, HTNC_Caption
                        If (mTitleBar.Buttons.SystemIconMenu And eSysMenuHidden) = 0 Then
                            postLparam = lParam: postWmsg = WM_GETSYSMENU
                        End If
                    Case HTNC_Tracker
                        mFrame.SetMouseAction wMsg, AmIActive
                    End Select
                    
                Case WM_NCLBUTTONDOWN
                    If (mActiveState And 1) = 0 Then SetForegroundWindow hWnd
                    ' never let VB accept a Left Button click on NC. It will repaint
                    bOverridden = True
                    Select Case wParam
                        ' the postmessage values are cached so PostMessage can be called if the user
                        ' does not override the message
                        Case HTERROR, HTBORDER, HTNOWHERE ' let these go thru
                            bOverridden = False
                            mMovementFlag = 0
                        Case Is > HTMenu_Custom, HTCLOSE, HTMAXBUTTON, HTMINBUTTON
                            ' pass message to the button class
                            mMovementFlag = wParam
                            If mTitleBar.Buttons.SetMouseAction(wMsg, wParam, 0, 0) = 0 Then mMovementFlag = 0
                        Case HTCAPTION, HTMenu_Custom, HTNC_Caption
                            If wParam = HTMenu_Custom Then
                                lOverrideReturn = 2
                                If (mMenuBar.pTracking(mbFlags) And 8) = 8 Then
                                    If mTitleBar.Location = 0 Then lOverrideReturn = 1
                                End If
                            End If
                            If lOverrideReturn = 2 Then
                                mMovementFlag = 0
                                mMenuBar.SetMouseAction wParam, lParam, WM_NCLBUTTONDOWN
                            Else
                                'set a flag to indicate what NC area is being clicked
                                If isMinimized Then
                                    mMovementFlag = 0
                                Else
                                    mMovementFlag = HTCAPTION
                                    ' using undocumented parameter. See notes at top of module for more info
                                    postLparam = lParam: postWmsg = WM_SYSCOMMAND
                                    postWparam = SC_MOVE Or HTCAPTION ' undocumented
                                End If
                            End If
                            lOverrideReturn = 0
                        Case HTLEFT To HTBOTTOMRIGHT   ' borders
                            'set a flag to indicate what NC area is being clicked
                            mMovementFlag = wParam
                            ' for vertically aligned titlebars, remove some window styles so that
                            ' we can size horizontally below the system setting
                            If mTitleBar.Alignment < barAlignBottom Then ToggleWSCaption True, True
                            ' using undocumented parameter. See notes at top of module for more info
                            postLparam = lParam: postWmsg = WM_SYSCOMMAND
                            postWparam = SC_SIZE Or (wParam - 9) ' undocumented
                        Case HTSYSMENU
                            If (mTitleBar.Buttons.SystemIconMenu And eSysMenuHidden) = 0 Then
                                mMovementFlag = wParam
                                mMenuBar.ShowMenu -1, 2, False
                            End If
                        Case HTNC_Tracker
                            mMovementFlag = 0
                            mFrame.SetMouseAction wMsg, AmIActive
                        Case Else
                            mMovementFlag = 0
                    End Select
            
                Case WM_NCLBUTTONDBLCLK ' don't let VB have this one either ot it will repaint
                    bOverridden = True
                    If wParam = HTCAPTION Then  ' maximize or restore
                        If ((mWindowState And SIZE_MAXIMIZED) = 0) Then postWparam = SC_MAXIMIZE Else postWparam = SC_RESTORE
                        postWmsg = WM_SYSCOMMAND
                    End If
                    
                Case WM_NCLBUTTONUP
                    ' never let VB accept a Left Button click on NC. It will repaint
                    bOverridden = True
                    Select Case wParam
                        ' the postmessage values are cached so PostMessage can be called if the user
                        ' does not override the message
                        Case HTERROR, HTBORDER, HTNOWHERE ' let these go thru
                            bOverridden = False
                        ' for htclose,htiminbutton,htmaxbutton, only allow the click event to
                        ' occur if the button was in the down state. This prevents someone from dragging
                        ' off the menubar onto a button and firing a click on the up event. The
                        ' DrawButtons routine below returns false if the state didn't change
                        Case HTCLOSE, HTMINBUTTON, HTMAXBUTTON
                            ' these can only come in if someone drags a mouse from anywhere
                            ' and release it over a system button. Normal system button clicks
                            ' have been diverted via the WM_LButtonUp message due to SetCapture
                        Case HTSYSMENU ' this is only triggered on a double click; that is because
                                    ' when first displayed, SetCapture is used which removes the
                                    ' second click, but we catch it here. As a precautionary task,
                                    ' the mMovement flag was set when the 1st click happened.
                                    ' Kinda awkware IMO, but it works really well
                            If mMovementFlag = HTSYSMENU Then
                                If (mTitleBar.Buttons.SystemIconMenu And eSysMenuHidden) = 0 Then
                                    postWparam = SC_CLOSE: postWmsg = WM_SYSCOMMAND
                                End If
                            End If
                        Case HTNC_Tracker
                            mFrame.SetMouseAction wMsg, AmIActive
                        Case Else: ' ???
                    End Select
                    mMovementFlag = 0
                
                Case WM_LBUTTONUP
                    ' ensure we release mouse capture & trigger the ExitSizeMove message as needed
                    If mMovementFlag Then
                        ' something is in flux...
                        Select Case mMovementFlag
                        Case Is > HTMenu_Custom, HTCLOSE, HTMINBUTTON, HTMAXBUTTON
                            bOverridden = True
                            ReleaseCapture
                            wPTs.x = LoWord(lParam)
                            wPTs.Y = HiWord(lParam)
                            ClientToScreen mHWnd, wPTs
                            mTitleBar.Buttons.SetMouseAction wMsg, mMovementFlag, wPTs.x - mWinRect.Left, wPTs.Y - mWinRect.Top
                        Case Else
                            ' shouldn't be anything else where we captured the mouse
                        End Select
                        mMovementFlag = 0
                    End If
                
                Case WM_CONTEXTMENU ' popup menu.
                    ' If it is the system menu we will reroute the request
                    GetCursorPos wPTs
                    bOverridden = True
                    ' is this in within the caption or sysIcon
                    If mMenuBar.HitTest(wPTs.x - mWinRect.Left, wPTs.Y - mWinRect.Top) = 0 Then
                        Select Case mTitleBar.HitTest(wPTs.x - mWinRect.Left, wPTs.Y - mWinRect.Top)
                        Case HTCAPTION, HTSYSMENU, HTNC_Caption
                            If (mTitleBar.Buttons.SystemIconMenu And eSysMenuHidden) = 0 Then
                                postWmsg = WM_GETSYSMENU: postLparam = lParam
                            End If
                        Case Is < HTCLIENT + 1
                            ' not in the NC
                        Case Else
                            bOverridden = False
                        End Select
                    Else
                        bOverridden = True
                    End If
                
                Case WM_ENTERSIZEMOVE
                    ReleaseHooks
                    If mMovementFlag = 1 Then SetCursor commonCursor(5)
                
                Case WM_EXITSIZEMOVE
                    ToggleWSCaption False, False
                    mMovementFlag = 0
                    ReduceDC 0, 0, True, 0
                
                Case WM_MENUCHAR
                    If mMenuBar.pTracking(mbActiveStatus) = 0 Then
                        If mMenuBar.pMenuHandle <> 0 Then
                            ' not currently tracking
                            bOverridden = True
                            ' begin tracking the menu loop if not already started
                            mMenuBar.pTracking(mbHwnd) = hWnd
                            ' since we removed the menu, any Alt+Keys will generate a beep
                            ' if we don't prevent it. Per MSDN this should do it: loword 0 & hiword 1
                            If mMenuBar.FindMenuItem(LoWord(wParam)) Then lOverrideReturn = 65536
                        End If
                    End If
                    
                End Select
                
            End If
            
            Select Case wMsg
            
            Case WM_NCCALCSIZE
                If wParam Then  ' nonclient dimensions requested
                    ' send original along & then modify the NC area as needed
                    DefWindowProc hWnd, wMsg, wParam, ByVal lParam
                    CalcNCArea lParam
                    lOverrideReturn = WVR_VALIDRECTS ' special meaning
                    bOverridden = True ' don't forward this message
                End If
            
            Case WM_APPACTIVATE ' used when option to keep form looking active is applied
                If (mActiveState Or 2) = mActiveState Then  ' user wants this option
                    If wParam = 0 Then '  not our thread
                        If ((mActiveState Or 8) = mActiveState) Then
                            'want to appear active in all threads
                            If (mActiveState Or 16) <> mActiveState Then
                                mActiveState = mActiveState Or 16
                                DoNCpaint
                            End If
                        Else
                            ' only want to appear in our thread; so show as inactive
                            If (mActiveState Or 16) = mActiveState Then
                                mActiveState = mActiveState And Not 16
                                DoNCpaint
                            End If
                        End If
                    Else                    ' this is our thread
                        ' want to appear active in our thread else mActiveState wouldn't include 2
                        If (mActiveState Or 16) <> mActiveState Then
                            mActiveState = mActiveState Or 16
                            DoNCpaint
                        End If
                    End If
                End If
                
            Case WM_NCACTIVATE ' NC draws; if wParam<>0 then NC is active else inactive
                ' remove the ws_caption style & menu & flag actual Active/Inactive status
                If wParam = 0 Then
                    ReleaseHooks
                    mActiveState = mActiveState And Not 1 ' not active
                    ' going active from inactive, prevent menubar from using last drawn object
                    mMenuBar.pTracking(mbLastDrawn) = 0
                Else
                    mActiveState = mActiveState Or 1      ' active
                End If
                If Not isMinimized Then
                    ' remove the caption style & forward message to base class
                    ToggleWSCaption True, False
                    ' need to pass with CallWindowProc vs DefWindowProc else ill effects
                    ' when opening a form via usercontrol & that new form is also subclassed
                    lOverrideReturn = CallWindowProc(hOldWndProc, hWnd, wMsg, wParam, lParam)
                    ' replace the ws_caption style
                    RemoveFormMenu  ' remove the form's menu if any
                    ToggleWSCaption False, False
                
                    ' to update the menubar (which would be inactive colors)
                    bOverridden = True
                    DoNCpaint
                End If
                
            Case WM_GETMINMAXINFO   ' handle overriden window size settings
                bOverridden = mFrame.MinMax.SetMinMaxInfo(hWnd, lParam, isMinimized)
                
            Case WM_SYSCOMMAND
                ' remove ws_caption style for Alt+Space sysmenu
                Select Case wParam
                    Case SC_KEYMENU, SC_MOUSEMENU, SC_MOUSEMENU Or 3
                        If isMinimized Then
                            bOverridden = ((mTitleBar.Buttons.SystemIconMenu And eSysMenuHidden) = eSysMenuHidden)
                        Else
                            If lParam = 0 Then  ' alt+key
                                RemoveFormMenu
                                bOverridden = True
                                mMenuBar.FindMenuItem -2
                            Else
                                If lParam = 32 Then
                                    If (mTitleBar.Buttons.SystemIconMenu And eSysMenuHidden) = 0 Then
                                        mMenuBar.FindMenuItem -1
                                    End If
                                    bOverridden = True
                                Else
                                    ToggleWSCaption True, False
                                End If
                            End If
                        End If
                    Case SC_MOVE, (SC_MOVE Or HTCAPTION)
                        If mTitleBar.Buttons.pEnabled(SC_MOVE) Then
                            mMovementFlag = Abs(wParam = SC_MOVE)
                            ' flag triggers the 4-star cursor, but may be overwritten
                        Else
                            bOverridden = True
                        End If
                    Case SC_SIZE
                        If mTitleBar.Buttons.pEnabled(SC_SIZE) Then
                            If mTitleBar.Alignment < barAlignBottom Then ToggleWSCaption True, True
                        Else
                            bOverridden = True
                        End If
                    Case SC_CLOSE
                        bOverridden = Not mTitleBar.Buttons.pEnabled(wParam)
                    Case SC_MAXIMIZE, SC_RESTORE, SC_MINIMIZE
                        bOverridden = Not mTitleBar.Buttons.pEnabled(wParam)
                        If Not bOverridden Then bOverridden = MinimizeAnimated(wMsg, wParam, lParam)
                    Case (SC_SIZE Or HTBOTTOMRIGHT)
                            bOverridden = True
                            SetWindowPos hWnd, 0, 0, 0, _
                                LoWord(lParam), HiWord(lParam), SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_NOMOVE
                    Case Else
            '            Stop
                End Select
                
            Case WM_MOVE ' our copy of the window & client coords changed. Update them
                wPTs.x = LoWord(lParam): wPTs.Y = HiWord(lParam)
                OffsetRect mWinRect, wPTs.x - mClientRect.Left, wPTs.Y - mClientRect.Top
                OffsetRect mClientRect, wPTs.x - mClientRect.Left, wPTs.Y - mClientRect.Top
            
            Case WM_SETICON, WM_SETTEXT ' VB will redraw NC w/o sending NCpaint message
                ToggleWSCaption True, False
                lOverrideReturn = CallWindowProc(hOldWndProc, hWnd, wMsg, wParam, lParam)
                bOverridden = True
                ToggleWSCaption False, False
                If wMsg = WM_SETICON Then mTitleBar.Buttons.RefreshIcon
                DoNCpaint
                
        End Select
    End If

End Select  ' end of the post notifications

    If bOverridden Then   ' then return the override value
        ' all PostMessage actions are overridden as result of preventing mouse clicks
        ' on titlebar or to prevent VB redrawing titlebar as result of sysmenu action
        If postWmsg Then PostMessage hWnd, postWmsg, postWparam, ByVal postLparam
        ' we don't pass the message in case next subclasser un-does what we/you did
        ProcessMessage = lOverrideReturn
    Else                ' otherwise pass to next subclasser in the chain
        ProcessMessage = CallWindowProc(hOldWndProc, hWnd, wMsg, wParam, lParam)
    End If
    If wMsg = WM_SIZE Then
        Select Case wParam
            Case SIZE_RESTORED ' restoring from minimize or maximize position
                If mWindowState <> SIZE_RESTORED Then
                    mWindowState = SIZE_RESTORED
                    mTitleBar.Buttons.CalcButtonOffsets 0, 0, 0
                    DoNCpaint
                End If
            Case SIZE_MAXIMIZED
                mWindowState = SIZE_MAXIMIZED
                mTitleBar.Buttons.CalcButtonOffsets 0, 0, 0
                mTitleBar.Buttons.DisableRestore True
                DoNCpaint
            Case SIZE_MINIMIZED
                mWindowState = (mWindowState And 4) Or SIZE_MINIMIZED
                mTitleBar.Buttons.DisableRestore False
                mTitleBar.Buttons.CalcButtonOffsets 0, 0, 0
        End Select
    End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.SubclassTerminated
' DateTime  : 8/27/05 22:29
' Author    : LaVolpe
' Purpose   : clean up code called when a wm_destroy message is received
' Comments  : see below
'---------------------------------------------------------------------------------------
'
Friend Sub SubclassTerminated(ByVal hWnd As Long)

    ' Also called if this class is set to Nothing

    Dim H As Long, bIsInterface As Boolean
    
    If hWnd = 0 Then Exit Sub
    
    bIsInterface = (hWnd = mHWnd)   ' is this the main-level window?
    If bIsInterface Then            ' if so, unsubclass it
        mClientImplementation = 0
        Me.OwnerDrawn Nothing, 0
        ReleaseSubClassing hWnd
        mHWnd = 0
        mActiveState = 0
        Set mMenuBar = Nothing
        Set mTitleBar = Nothing
        Set mFrame = Nothing
        Set mGraphics = Nothing
    End If

    If Not mControls Is Nothing Then    ' do we have subclassed controls?
        On Error Resume Next
        If bIsInterface Then            ' if unsubclassing main-level form
            For H = 1 To mControls.Count    ' also unsubclass any controls
                ReleaseSubClassing mControls.Item(H)
            Next
            Set mControls = Nothing
        Else                            ' unsubclassing a control & not the main-level
            ReleaseSubClassing hWnd
            If mControls.Count = 1 Then
                Set mControls = Nothing
            Else
                mControls.Remove "h" & hWnd
            End If
        End If
        If Err Then Err.Clear
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.DoNCpaint
' DateTime  : 8/29/05
' Author    : LaVolpe
' Purpose   : Coordnates painting the non-client (NC) area of the window
' Comments  : see below
'---------------------------------------------------------------------------------------
'
Private Sub DoNCpaint()

    If mSelfModifying Then Exit Sub
        ' ^^ see ToggleWSCaption routine for usage of this variable
    
    If (mWindowState And SIZE_MINIMIZED) = SIZE_MINIMIZED Then Exit Sub
    
'    Debug.Print "painting NC "; mActiveState; mHWnd
    
    Dim cRgn As Long
    Dim cRect As RECT, ncRect As RECT
    Dim MbarRect As RECT, TbarRect As RECT, insetRect As RECT
    Dim hDC As Long, hRgn As Long, cDC As Long
    Dim Cx As Long, Cy As Long
    Dim hBrush As Long, lRtn As Long
    Dim cClient As iOwnerDrawn, cClientID As Long, BKG As BkgAction_LV
    
    Cx = mWinRect.Right - mWinRect.Left + 1
    Cy = mWinRect.Bottom - mWinRect.Top + 1
    
    cDC = Canvas(True)
    If cDC = 0 Then Exit Sub
    
    ncRect = mWinRect
    
    If IsRectEmpty(mClientRect) = 0 Then
        cRect = mClientRect
        ' shift the client area appropriately to act as a clipping region
        OffsetRect cRect, -mWinRect.Left, -mWinRect.Top
        ' set our window screen coords to 0,0
        OffsetRect ncRect, -ncRect.Left, -ncRect.Top
        ' create the overall region & subtract the client area
        hRgn = CreateRectRgnIndirect(ncRect)
        cRgn = CreateRectRgnIndirect(cRect)
        CombineRgn hRgn, hRgn, cRgn, &H4
        DeleteObject cRgn
    Else
        OffsetRect ncRect, -mWinRect.Left, -mWinRect.Top
        OffsetRect cRect, -mWinRect.Left, -mWinRect.Top
    End If
    
    cClientID = mFrame.DrawFrame(VarPtr(mWinRect), VarPtr(mClientRect), mClientInset, AmIActive)
    
    ' testing for crashing if user unloads while using callbacks,
    ' we can pretty much prevent a crash if we place checks in key
    ' spots in the classes... The following allow owner-drawn procedures
    mTitleBar.DrawTitleBar AmIActive, mWindowState
    mMenuBar.DrawMenuBar AmIActive
   
    If mHWnd = 0 Then Exit Sub  ' ensure user didn't unload b4 we got here
    
    cDC = Canvas(True)
    If cClientID Then
        ' for post NC painting, we will be providing the user the locations of
        ' all the key rectangles used for this window. Notice we do not pass
        ' them pointers to the actual rectangles. Don't want them changed
        BKG.hDC = cDC
        BKG.rcItem = ncRect                 ' pass the window rectangle
        BKG.rcExtra.Left = VarPtr(cRect)    ' pass the adjusted client rectangle
            ' calculate the titlebar rectangle
            BKG.rcExtra.Top = mTitleBar.Location
            If BKG.rcExtra.Top = 0 Then
                SetRect TbarRect, 0, 0, 0, 0
            Else
                CopyMemory TbarRect, ByVal BKG.rcExtra.Top, &H10
                InflateRect TbarRect, mTitleBar.BorderWidth, mTitleBar.BorderHeight
            End If
            BKG.rcExtra.Top = VarPtr(TbarRect) ' pass the titlebar rectangle
        ' calcualte the menubar rectangle
        BKG.rcExtra.Right = mMenuBar.Location
        If BKG.rcExtra.Right = 0 Then
            SetRect MbarRect, 0, 0, 0, 0
        Else
            CopyMemory MbarRect, ByVal BKG.rcExtra.Right, &H10
        End If
        BKG.rcExtra.Right = VarPtr(MbarRect) ' pass the menubar rectangle
        ' have frame calculate the nonclient inset & return pointer
        CopyMemory insetRect, ByVal Frame.GetInset(BKG.rcExtra.Left, mClientInset), &H10
        BKG.rcExtra.Right = VarPtr(insetRect) ' pass the nonclient inset rectangle
        
        ' now we are ready to pass the info to the client
        On Error Resume Next ' always, to try & prevent user from crashing
        GetObjectFromPointer cClientID, cClient
        cClient.OwnerDrawMessage omDrawPostNC, AmIActive + 0, VarPtr(BKG), ""
        Set cClient = Nothing
        On Error GoTo 0
    
    End If
    
    ' get the DC & copy the client area
    hDC = GetWindowDC(mHWnd)
    ' add the clipping region & copy the image
    If hRgn Then
        SelectClipRgn hDC, hRgn
        DeleteObject hRgn
    End If
    BitBlt hDC, 0, 0, Cx, Cy, cDC, 0, 0, vbSrcCopy
    
    ' Clean up
    If hRgn Then SelectClipRgn hDC, ByVal 0&
    mMenuBar.UpdateMenuBar mMenuBar.pTracking(mbLastDrawn) - 1, hDC, AmIActive
    ReleaseDC mHWnd, hDC

End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.CalcNCArea
' DateTime  : 8/29/05
' Author    : LaVolpe
' Purpose   : Calculates the client area in relation to the non-client area
' Comments  : see below
'---------------------------------------------------------------------------------------
'
Private Sub CalcNCArea(lParam As Long)


    Dim tNCR As NCCALCSIZE_PARAMS
    Dim tWP As WINDOWPOS
    Dim rBorderCx As Byte, rBorderCy As Byte, cBorder As RECT
    Dim lRtn As Long, cBorderPtr As Long
    Dim minCapLen As POINTAPI
    Dim tbAlignment As eBarAlignment

    ' Get recommended client rect info
    CopyMemory tNCR, ByVal lParam, Len(tNCR)
    

    If Not mSelfModifying Then
    ' ^^ see ToggleWSCaption routine for usage of this variable
    
        '        Debug.Print "calculating NonClient "; mSelfModifying
        mAutoRedraw = True
    
        ' Get the new window position/size
        CopyMemory tWP, ByVal tNCR.lppos, Len(tWP)
        
        ' extract the 4 border sizes & custom inset offset
        mFrame.BorderWidths rBorderCx, rBorderCy, cBorderPtr
        
        ' set the new window rect here
        SetRect mWinRect, tWP.x, tWP.Y, tWP.x + tWP.Cx, tWP.Y + tWP.Cy
        ' build the client rectangle using 0,0 as top left. Initially windowsize-borders
        SetRect mClientRect, rBorderCx, rBorderCy, tWP.Cx - rBorderCx, tWP.Cy - rBorderCy
        
        ExpandDC tWP.Cx, tWP.Cy, True, mHWnd
                
        ' have class calculate the titlebar. Return value is a pointer to the titlebar Rect
        lRtn = mTitleBar.MeasureTitlebar(VarPtr(mClientRect), AmIActive, tbAlignment, minCapLen.x, minCapLen.Y)
        
        If lRtn = 0 Then  ' otherwise no titlebar used
            mClientInset = 0
            mFrame.MinMax.ResetToSystemDefaults
            tbAlignment = 0
        Else
            ' custom drawn nc, custom minimum drag size needed too
            mFrame.MinMax.SetMinDragSize lRtn, rBorderCx, rBorderCy, tbAlignment, VarPtr(minCapLen)
            
            Select Case tbAlignment
                Case barAlignLeft: mClientInset = 1
                Case barAlignTop: mClientInset = 2
                Case barAlignRight: mClientInset = 4
                Case barAlignBottom: mClientInset = 8
            End Select
        End If
        
        ' have class calculate the menubar. Return value is a pointer to the menubar rect
        lRtn = mMenuBar.MeasureMenubar(VarPtr(mClientRect), AmIActive, tbAlignment, (mTitleBar.Buttons.SystemIconMenu And eSysMenuHidden) = eSysMenuHidden)
        If lRtn <> 0 Then
            Select Case tbAlignment
                Case barAlignLeft: mClientInset = mClientInset Or 1
                Case barAlignTop: mClientInset = mClientInset Or 2
                Case barAlignRight: mClientInset = mClientInset Or 4
                Case barAlignBottom: mClientInset = mClientInset Or 8
            End Select
        End If
        
        ' adjust the client rectangle for the non-client inset
        If cBorderPtr <> 0 Then
            CopyMemory cBorder, ByVal cBorderPtr, &H10
            With mClientRect
                .Left = .Left + cBorder.Left
                .Top = .Top + cBorder.Top
                .Right = .Right - cBorder.Right
                .Bottom = .Bottom - cBorder.Bottom
            End With
        End If
        
        ' ensure a valid rectangle after all adjustments are made
        If mClientRect.Bottom < mClientRect.Top Then mClientRect.Bottom = mClientRect.Top
        If mClientRect.Right <= mClientRect.Left Then mClientRect.Right = mClientRect.Left
        
        mTitleBar.Buttons.UpdateUserToolTips VarPtr(mClientRect)
        
        ' adjust the 0,0 client coordinates to window/screen coordinates
        OffsetRect mClientRect, mWinRect.Left, mWinRect.Top
               
    End If
    
    ' Return the new client area
    tNCR.rgrc(0) = mClientRect
    tNCR.rgrc(1) = mClientRect
    CopyMemory ByVal lParam, tNCR, Len(tNCR)
    
    ' in a real app, the following probably won't apply
    ' Toggling btwn horizontal and vertical on a super small window
    ' can allow titlebar items to paint over each other, so we
    ' will adjust the window size if needed.
    
    If mSelfModifying = False Then
        Me.Frame.MinMax.GetMinDragSize tWP.x, tWP.Y
        If tWP.Cx < tWP.x Then
            tWP.Cx = tWP.x
            tWP.x = -1
        End If
        If tWP.Cy < tWP.Y Then
            tWP.Cy = tWP.Y
            tWP.Y = -1
        End If
        If tWP.x < 0 Or tWP.Y < 0 Then PostMessage mHWnd, WM_SYSCOMMAND, SC_SIZE Or HTBOTTOMRIGHT, MakeLong(Abs(tWP.Cx) + 0, Abs(tWP.Cy) + 0)
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.ControlProxy
' DateTime  : 8/27/05 22:32
' Author    : LaVolpe
' Purpose   : Forwards messages sent to a subclassed control
' Comments  : see below
'---------------------------------------------------------------------------------------
'
Private Function ControlProxy(ByVal hWnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Long, ByVal hOldWndProc As Long) As Long
' This class also allows subclassing child controls once the parent has been subclassed
' Function simply passes the control's window message, allowing it to be modified/discarded

    Dim bConsume As Boolean
    Dim lRtnVal As Long
    Dim Client As iControlsCallback

    If hOldWndProc = 0 Then Exit Function
    
    On Error Resume Next
    ' send user the windows message
    
    If mControlsImplementation = 0 Then
        ' we shouldn't get here, but if we did some how, lets forward
        ' the message and unsubclass the window
        SubclassTerminated hWnd
        ControlProxy = CallWindowProc(hOldWndProc, hWnd, wMsg, wParam, lParam)
        
    Else
        GetObjectFromPointer mControlsImplementation, Client
        Client.WindowMessage hWnd, wMsg, wParam, lParam, bConsume, lRtnVal
        Set Client = Nothing
        
        If Not bConsume Then
            ' user is letting the message go thru
            lRtnVal = CallWindowProc(hOldWndProc, hWnd, wMsg, wParam, lParam)
        End If
    
        ' return the value
        ControlProxy = lRtnVal
        
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.SetKeepActive
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Forces a window to paint active even if it isn't the active window
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function SetKeepActive(bSet As Boolean, AlwaysActive As Boolean, AllThreads As Boolean) As Long
    
    ' bSet :: sets properties, otherwise, returns properties
    ' AlwaysActive :: False to show inactive when window does not have the focus
    '                 True to override the default behavior
    ' AllThreads :: Only applies if AlwaysActive is True
    '           if AllThreads is True, window will never paint as inactive
    '           if AllThreads is False, window will paint inactive
    '               Only when the thread does not have the focus.
    '               In other words, when any window in the same thread gets the
    '               focus the window will paint as active. This option is kinda
    '               neat to override the calling form inactive when a message box is shown
    
    If bSet Then
        
        If AlwaysActive Then
            ' clear settings, keeping those that are appropriate
            mActiveState = mActiveState And Not 8 And Not 4
            mActiveState = mActiveState Or (4 + Abs(4 * AllThreads)) Or 2
            ' apply the new settings. See this variable declaration at top
            ' of module for the meaning of the different flag values
            If AllThreads Then
                mActiveState = mActiveState Or 16
            Else
                If GetCurrentThreadId = GetWindowThreadProcessId(GetForegroundWindow(), 0) Then
                    mActiveState = mActiveState Or 16
                End If
            End If
        Else
            ' default behavior
            mActiveState = mActiveState And Not 2
        End If
        ' repaint the window if appropriate
        If mHWnd <> 0 And mAutoRedraw = True Then DoNCpaint
        SetKeepActive = True
    
    Else
        ' return settings, set the parameter values
        If (mActiveState And 2) = 2 Then
            If (mActiveState And 8) = 8 Then  ' always active
                SetKeepActive = 2
            Else
                If SetKeepActive = (mActiveState And 4) = 4 Then SetKeepActive = 1
            End If
        End If
    End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.ShowOnTaskbar
' DateTime  : 9/25/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve the window style needed to place/remove taskbar items
'---------------------------------------------------------------------------------------
'
Public Property Get ShowOnTaskbar() As Boolean
    
    ShowOnTaskbar = ((GetWindowLong(mHWnd, GWL_EXSTYLE) And WS_EX_APPWINDOW) = WS_EX_APPWINDOW)
    
End Property
Public Property Let ShowOnTaskbar(bShow As Boolean)

    If mHWnd = 0 Then Exit Property
    
    Dim wStyle As Long, bWasMinimized As Boolean
    
    ' don't allow toggling if the window is minimized to the system tray
    If (mWindowState And 4) = 4 Then Exit Property
    
    wStyle = GetWindowLong(mHWnd, GWL_EXSTYLE)
    If (wStyle And WS_EX_APPWINDOW) = 0 Then ' not on taskbar
        If bShow Then
            mSelfModifying = True
            wStyle = wStyle Or WS_EX_APPWINDOW
        End If
    Else                                    ' currently on taskbar
        If Not bShow Then
            mSelfModifying = True
            wStyle = wStyle And Not WS_EX_APPWINDOW
        End If
    End If
    If mSelfModifying Then
        ' using LockWindowUpdate, we can prevent the window from disappearing
        ' while the ShowWindow & SetWindowLong functions are executed
        
        ' when triggering this for minimized windows. The O/S screams that it
        ' can be done. We simply need to unminimize the window during the action
        bWasMinimized = ((mWindowState And SIZE_MINIMIZED) = SIZE_MINIMIZED)
        LockWindowUpdate mHWnd
        ShowWindow mHWnd, 0
        SetWindowLong mHWnd, GWL_EXSTYLE, wStyle    ' set the style
        ForceRefresh mHWnd                          ' force the style
        If bWasMinimized Then
            ShowWindow mHWnd, &H1   ' show normal
            ShowWindow mHWnd, &H6   ' show minimized
        Else
            ShowWindow mHWnd, &H5   ' show the window
        End If
        mSelfModifying = False      ' clean up
        LockWindowUpdate 0&
    End If
End Property
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.AlwaysOnTop
' DateTime  : 9/25/2005
' Author    : LaVolpe
' Purpose   : Sets/Retrieves the window style neede to make a form Always on Top
' Comment   : Always on top is only applicable to the thread. Other windows that
'             have the same style, different threads, may overlap your window
'---------------------------------------------------------------------------------------
'
Public Property Get AlwaysOnTop() As Boolean
    
    Const WS_EX_TOPMOST As Long = &H8&
    AlwaysOnTop = ((GetWindowLong(mHWnd, GWL_EXSTYLE) And WS_EX_TOPMOST) = WS_EX_TOPMOST)

End Property
Public Property Let AlwaysOnTop(bOntop As Boolean)
    If mHWnd <> 0 Then
        If Me.AlwaysOnTop Then
            If bOntop = True Then Exit Property
        Else
            If bOntop = False Then Exit Property
        End If
        SetWindowPos mHWnd, ((bOntop = False) * 1) - 1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Property
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.AppIcon
' DateTime  : 9/25/2005
' Author    : LaVolpe
' Purpose   : Set the icon to appear in the Alt+Tab window & Task Manager
' Comments  : You need to preserve the icon until your app closes
'---------------------------------------------------------------------------------------
'
Public Property Let AppIcon(ByVal hIcon As Long)
    If hIcon = 0 Then Exit Property
    If mHWnd <> 0 Then
        Dim iInfo As ICONINFO, bmpInfo As BITMAPINFOHEADER
        If GetIconInfo(hIcon, iInfo) <> 0 Then
            If iInfo.hbmColor = 0 Then
                If iInfo.hbmMask = 0 Then
                    hIcon = 0
                Else
                    If GetGDIObject(iInfo.hbmMask, Len(bmpInfo), bmpInfo) = 0 Then hIcon = 0
                End If
            Else
                If GetGDIObject(iInfo.hbmColor, Len(bmpInfo), bmpInfo) = 0 Then hIcon = 0
                DeleteObject iInfo.hbmColor
            End If
            If iInfo.hbmMask Then DeleteObject iInfo.hbmMask
            If hIcon Then
                If bmpInfo.biWidth <> 32 Then hIcon = 0
            End If
            If hIcon <> 0 Then SetApplicationIcon 0, mHWnd
        End If
    End If
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsFrame.MaximizeFullScreen
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Option to show maximized window over the entire desk top, taskbars and all
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Property Let MaximizeFullScreen(bMaxFullScreen As Boolean)
    
    If mHWnd <> 0 Then
        If bMaxFullScreen = True And Me.MaximizeFullScreen = False Then
            SetWindowLong mHWnd, GWL_STYLE, GetWindowLong(mHWnd, GWL_STYLE) And Not WS_MAXIMIZEBOX And Not WS_THICKFRAME
            If mAutoRedraw Then ForceRefresh mHWnd
        Else
            If bMaxFullScreen = False And Me.MaximizeFullScreen = True Then
                SetWindowLong mHWnd, GWL_STYLE, GetWindowLong(mHWnd, GWL_STYLE) Or WS_MAXIMIZEBOX Or WS_THICKFRAME
                If mAutoRedraw Then ForceRefresh mHWnd
            End If
        End If
    End If
    
End Property
Public Property Get MaximizeFullScreen() As Boolean
    
    If mHWnd <> 0 Then
        Dim lStyle As Long
        lStyle = GetWindowLong(mHWnd, GWL_STYLE)
        MaximizeFullScreen = ((lStyle And (WS_MAXIMIZEBOX Or WS_THICKFRAME)) = 0)
    End If
    
End Property
'---------------------------------------------------------------------------------------
' Procedure : clsButtons.MinimizeToSysTray
' DateTime  : 9/17/2005
' Author    : LaVolpe
' Purpose   : Option sets properties so window appears to minimize to the system tray
'             vs just minimizing to the taskbar
'---------------------------------------------------------------------------------------
'
Public Property Let MinimizeToSysTray(bToSysTray As Boolean)
    
    If mHWnd = 0 Then Exit Property
    
    If bToSysTray Then
        If (mOptions And 2) = 2 Then Exit Property
        mOptions = mOptions Or 2    ' set flag
        If mWindowState = SIZE_MINIMIZED Then
            ' if window is already minimized, then hide it
            mWindowState = mWindowState Or 4
            ShowWindow mHWnd, 0
        Else
            If mWindowState = SIZE_RESTORED Then
                ' if the window was shown and never moved, minimized, maximized
                ' Windows won't have a restore rect for it yet. We need that
                ' rect if user wants to animate from the system tray. Therefore
                ' we force a minimize and restore to get Windows to save the
                ' restore rect so we can get it.
                LockWindowUpdate mHWnd
                mSelfModifying = True
                ShowWindow mHWnd, 0
                ShowWindow mHWnd, &H6   ' minimize
                ShowWindow mHWnd, 0
                ShowWindow mHWnd, &H9   ' previous state
                mSelfModifying = False
                LockWindowUpdate 0
            End If
        End If
    Else
        If (mOptions And 2) = 0 Then Exit Property
        mOptions = mOptions And Not 2
        If (mWindowState And 4) = 4 Then ' minimized to system tray
            LockWindowUpdate mHWnd
            ' remove the minimize flag
            mWindowState = mWindowState And Not 4
            ' show & then hide the window
            ShowWindow mHWnd, &H1
            ShowWindow mHWnd, &H6
            ' refresh
            LockWindowUpdate 0&
        End If
    End If
End Property
Public Property Get MinimizeToSysTray() As Boolean
    MinimizeToSysTray = ((mOptions And 2) = 2)
End Property
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.ShowWindowFromTrayIcon
' DateTime  : 9/24/2005
' Author    : LaVolpe
' Purpose   : This function will help a window appear to be maximized, restored
'       from the system tray. Could be called when tray icon click displays the window
'---------------------------------------------------------------------------------------
'
Public Sub ShowWindowFromTrayIcon(State As eMin2TrayOptions)

    ' only called by the user, never the DLL
    If mHWnd = 0 Then Exit Sub

    Dim wParam As Long
    
    Select Case State
    Case SIZE_MINIMIZED
        wParam = SC_MINIMIZE  ' show window as maximized out of the tray
    Case SIZE_RESTORED
        wParam = SC_RESTORE
    Case SIZE_MAXIMIZED
        wParam = SC_MAXIMIZE
    Case Else
        Exit Sub
    End Select
    PostMessage mHWnd, WM_SYSCOMMAND, wParam, ByVal 0&

End Sub

'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.ReleaseHooks
' DateTime  : 8/27/05 22:32
' Author    : LaVolpe
' Purpose   : Releases the 2 hooks used to track menubar actions
'---------------------------------------------------------------------------------------
'
Private Sub ReleaseHooks()
    ' called only at key points in the window active state
    mMenuBar.UpdateMenuBar -2
    SetInputHook False, 0
    mMenuBar.pTracking(-1) = 0
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.RemoveFormMenu
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : removes the window's menu as needed
'---------------------------------------------------------------------------------------
'
Private Sub RemoveFormMenu()
    Dim fMenu As Long
    fMenu = GetMenu(mHWnd)
    If fMenu Then
        mWinMenu = fMenu
            mSelfModifying = True   ' see ToggleWScaption for details on this flag
        SetMenu mHWnd, 0&
            mSelfModifying = False
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.ToggleWSCaption
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : modifies specific window style
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Function ToggleWSCaption(bRemove As Boolean, bSizing As Boolean) As Boolean
' VB will draw the min/max/close buttons, sys icon, & caption on specific messages.
' This is not normally an issue as we can draw over them most of the time.
' However, VB always assumes the Non-Client (NC) area will have space for a titlebar
' and menu if used. If this space is removed or relocated via processing of the
' WM_NCCalcSize message, VB still refuses to recognize the space is no longer there.
' It will draw the tilebar, menu, etc in the expected space even if that space is
' now occupied by the client area. In this case the titlebar info is painted directly
' on the client -- dumb! If we are allowed to modify the NC,
' then VB should recognize the modification IMO

' By removing the WS_Caption style from the window, VB is tricked into thinking there
' is no space to draw the buttons/caption/icon. We don't want to permanently remove the
' style 'cause VB can treat the window as Borderless & maximize over entire desktop
' and other issues that would have to be addressed, like no longer knowing it has
' min/max/close buttons and/or system icon. However, with that hack comes a downside>

' About the mSelfModifying variable used below...
' This flag is set/reset to prevent 2 routines from processing when there is no need
' to process: wm_ncCalcSize & the wm_ncPaint. Since I am triggering those messages
' as a result of this routine, I know that the window won't be changing size or position
' as a result of the action. Therefore, we can avoid the wasted attempts to recalculate
' menubar & titlebar size/positions, and also repainting the nonClient area.
' Both the CalcNCArea & DoNCpaint routines are coded to recgonize this flag.

' Side Note: Each time a menubar item displays a submenu, this routine is called to
' prevent VB from redrawing either the menu or non-Client titlebar. So without this
' added prevention; the nonClient would repaint every time a submenu is opened & closed.

' When sizing window that has a vertical titlebar, windows will apply the
' minimium drag sizes (width & height). This allows the vertical titlebar to
' be dragged vertically to a few pixels high which won't replicate the same effect
' if the titlebar was horizontal. Additionally, the window won't be allowed to
' be dragged horizontally to a point of showing just the titlebar. By removing
' the WS_DLGFrame style, we can pretty much allow the window to
' be dragged to any minimum size we want. Therefore this style is removed
' before a vertical titlebar window is sized & replaced after it is sized

    Dim lStyle As Long, tgtStyle As Long
    Const WS_DLGFRAME As Long = &H400000
    Const WS_BORDER As Long = &H800000
    
    ' add or remove the style if needed
    lStyle = GetWindowLong(mHWnd, GWL_STYLE)
    If bRemove Then
        If bSizing Then
            tgtStyle = lStyle And Not WS_DLGFRAME   ' remove DLGFrame
        Else
            tgtStyle = (lStyle And Not WS_CAPTION) Or WS_BORDER ' remove caption
        End If
    Else
        ' still not sure why this works exactly. Tweaking other combinations to prevent
        ' VB from drawing works, but has other side effects like not being able to
        ' toggle Focus from the windows taskbar. This combination seems to work in all
        ' cases. We'll see. Of course, the window when maximized will maximize over the
        ' entire desktop and we have to account for that by replacing the WS_Caption
        ' after the window is maximized to prevent that from happening.
        tgtStyle = lStyle Or WS_CAPTION Or WS_DLGFRAME Or WS_BORDER
    End If
    If tgtStyle <> lStyle Then
'        RemoveFormMenu
        ' apply the new window style
        mSelfModifying = True
            SetWindowLong mHWnd, GWL_STYLE, tgtStyle
            ForceRefresh mHWnd
        mSelfModifying = False
        ToggleWSCaption = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.iImpProps_Get/SetPropValue
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Returns properties of this class and other child classes
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Function iImpProps_GetPropValue(propID As eProperties) As Long
    Select Case propID
        Case [_tbarAlignment]: iImpProps_GetPropValue = mTitleBar.Alignment
            '^^ called by the menu class when trying to display system menu
        Case [_titleBarRect]: iImpProps_GetPropValue = mTitleBar.Location
            '^^ called by the menu class when trying to display system menu
        Case [_MenuFont]: iImpProps_GetPropValue = mMenuBar.pFont
            '^^ called by the cGraphics class when user want to load menu font
        Case [_TitleFont]: iImpProps_GetPropValue = mTitleBar.pFont
            '^^ called by the cGraphics class when user want to load titlebar font
        Case [_clientRect]: iImpProps_GetPropValue = VarPtr(mClientRect)
            '^^ called when adding a tooltip to a toolwindow w/o a previous tooltip
        Case [_winRect]: iImpProps_GetPropValue = VarPtr(mWinRect)
            '^^ called when adding a tooltip to a toolwindow w/o a previous tooltip
        
        ' the following not used any longer; will remove when satisfied not needed
        Case [_mnuSysMenu]: iImpProps_GetPropValue = mMenuBar.pSysMenuHandle
        Case [_winHwnd]: iImpProps_GetPropValue = mHWnd
        Case [_winState]: iImpProps_GetPropValue = (mWindowState And Not 4)
        
    End Select
End Function
Private Sub iImpProps_SetPropValue(propID As eProperties, vValue As Long)
    Select Case propID
    Case [_sysTrayImp]  ' callback when user establishes a system tray icon
        mTrayImplementation = vValue
    End Select
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.SendCallBack
' DateTime  : 9/25/2005
' Author    : LaVolpe
' Purpose   : Pass subclassed messages to subclassed window for pre-processing
'---------------------------------------------------------------------------------------
'
Private Sub SendCallBack(hWnd As Long, wMsg As Long, wParam As Long, lParam As Long, _
                            bBlockMessage As Boolean, BlockValue As Long)
    
    If mClientImplementation <> 0 Then
        If SearchMessageArray(wMsg) > -1 Then
            On Error Resume Next
            Dim mClient As iCoreMessages
            GetObjectFromPointer mClientImplementation, mClient
            mClient.WindowMessage hWnd, wMsg, wParam, lParam, bBlockMessage, BlockValue
            Set mClient = Nothing
        End If
    End If
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.InitialiazeSkin
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Do what is needed to start the skinning
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub InitialiazeSkin()

' this routine has 3 purposes & only called when customizing first begins
' 1. Reset key variables/properties
' 2. Pre-initialize classes like the clsSysMenu & clsButtons
' 3. Size the new skinned window so that the client rectangle will be the
'       same size after being skinned

Dim rBorderCx As Byte, rBorderCy As Byte
Dim lRtn As Long, cBorderPtr As Long
Dim cRect As RECT, cBorder As RECT
Dim Cx As Long, Cy As Long
Dim activeHwnd As Long

    mMovementFlag = 0                   ' used when dragging NC objects
    Me.Titlebar.pMainImplementation(mHWnd) = ObjPtr(Me)
    Me.MenuBar.pMainImplementation(mHWnd) = ObjPtr(Me)
    Me.Frame.pMainImplementation(mHWnd) = ObjPtr(Me)
    
    activeHwnd = GetForegroundWindow()  ' see if this window is the foreground
    If activeHwnd = mHWnd Then          ' it is, update the active flag
        mActiveState = mActiveState Or 1
        ' if user is forcing appearance to be active, update the flag again
        If (mActiveState Or 2) = mActiveState Then mActiveState = mActiveState Or 16
    Else
        ' not foreground, but is user forcing it to appear active?
        mActiveState = mActiveState And Not 1 And Not 16
        If (mActiveState Or 2) = mActiveState Then
            If (mActiveState Or 8) = mActiveState Then
                ' user wants it to appear active all the time - why?
                mActiveState = mActiveState Or 16
            Else
                ' see if user wants it to appear active only in its thread
                If GetWindowThreadProcessId(activeHwnd, ByVal 0&) = GetCurrentThreadId() Then
                    mActiveState = mActiveState Or 16
                End If
            End If
        End If
    End If
    
    ' get the window & client rectangles
    GetWindowRect mHWnd, mWinRect
    GetClientRect mHWnd, cRect
    ExpandDC mWinRect.Right - mWinRect.Left, mWinRect.Bottom - mWinRect.Top, True, mHWnd
    
    ' extract the 4 border sizes
    mFrame.BorderWidths rBorderCx, rBorderCy, cBorderPtr
    
    ' build the client rectangle using 0,0 as top left. Initially windowsize-borders
    SetRect mClientRect, rBorderCx, rBorderCy, mWinRect.Right - mWinRect.Left - rBorderCx + 1, mWinRect.Bottom - mWinRect.Top - rBorderCy + 1
        
    ' have class calculate the titlebar. Return value is a pointer to the titlebar Rect
    lRtn = mTitleBar.MeasureTitlebar(VarPtr(mClientRect), AmIActive, 0, 0, 0)
    
    ' have class calculate the menubar. Return value is a pointer to the menubar rect
    lRtn = mMenuBar.MeasureMenubar(VarPtr(mClientRect), AmIActive, 0, False)
    
    'the mClientRect is the current size for the current settings
    ' adjust the window size appropriately & send the frame changed notification
    ' this when done, should apply the new custom settings & keep the client area same size
    Cx = cRect.Right - (mClientRect.Right - mClientRect.Left) + (mWinRect.Right - mWinRect.Left) + 1
    Cy = cRect.Bottom - (mClientRect.Bottom - mClientRect.Top) + (mWinRect.Bottom - mWinRect.Top) + 1
    If cBorderPtr <> 0 Then
        CopyMemory cBorder, ByVal cBorderPtr, &H10
        Cx = Cx + cBorder.Left + cBorder.Right
        Cy = Cy + cBorder.Top + cBorder.Bottom
    End If
    
    ' force window to redraw itself completely,
    ' triggering the needed WM_NCCalcSize message
    SetWindowPos mHWnd, 0, mWinRect.Left, mWinRect.Top, Cx, Cy, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOZORDER
    If mTitleBar.Buttons.RotateSystemIcon > Rotate0 Then mTitleBar.Buttons.RefreshIcon

End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.MessagesAdd/MessageRemove
' DateTime  : 9/18/2005
' Author    : LaVolpe
' Purpose   : Pass a comma-delimited listing of messages you want to view/modify
'---------------------------------------------------------------------------------------
'
Public Sub MessagesAdd(iImplementation As iCoreMessages, ParamArray Messages() As Variant)
    
    If iImplementation Is Nothing Then Exit Sub
    
    Dim M As Long, newArray() As Long, nrItems As Long
    
    If LBound(Messages) <= UBound(Messages) Then ' invalid array passed or null variant
        ReDim newArray(0 To UBound(Messages) - LBound(Messages))
        For M = LBound(Messages) To UBound(Messages)
            Select Case VarType(Messages(M))
            Case vbByte, vbInteger, vbLong
                If SearchMessageArray(CLng(Messages(M))) < 0 Then
                    newArray(nrItems) = CLng(Messages(M))
                    nrItems = nrItems + 1
                End If
            End Select
        Next
        If nrItems > 0 Then
            mClientImplementation = ObjPtr(iImplementation)
            If IsArrayEmpty(Not cClientMsgs) Then
                ReDim cClientMsgs(0 To nrItems - 1)
                CopyMemory cClientMsgs(0), newArray(0), nrItems * 4
            Else
                ReDim Preserve cClientMsgs(0 To UBound(cClientMsgs) + nrItems)
                CopyMemory cClientMsgs(UBound(cClientMsgs) - nrItems + 1), newArray(0), nrItems * 4
            End If
            ShellSortMessages
            cMsg_NChitTest = Abs(SearchMessageArray(WM_NCHITTEST) > -1)
            cMsg_SetCursor = Abs(SearchMessageArray(WM_SETCURSOR) > -1)
        End If
    End If
End Sub
Public Sub MessagesRemove(ParamArray Messages() As Variant)
    
    If IsArrayEmpty(Not cClientMsgs) Then Exit Sub
    
    Dim M As Long, Index As Long
    If LBound(Messages) <= UBound(Messages) Then ' invalid array passed or null variant
                
        For M = LBound(Messages) To UBound(Messages)
            Select Case VarType(Messages(M))
            Case vbByte, vbInteger, vbLong
                Index = SearchMessageArray(CLng(Messages(M)))
                If Index > -1 Then
                    If UBound(cClientMsgs) = 0 Then
                        Erase cClientMsgs
                        mClientImplementation = 0
                        Exit For
                    ElseIf Index = UBound(cClientMsgs) Then
                        ReDim Preserve cClientMsgs(0 To UBound(cClientMsgs) - 1)
                    Else
                        CopyMemory cClientMsgs(Index), cClientMsgs(Index + 1), (UBound(cClientMsgs) - Index) * 4
                        ReDim Preserve cClientMsgs(0 To UBound(cClientMsgs) - 1)
                    End If
                End If
            Case Else
            End Select
        Next
        If mClientImplementation = 0 Then
            cMsg_NChitTest = 0
            cMsg_SetCursor = 0
        Else
            cMsg_NChitTest = Abs(SearchMessageArray(WM_NCHITTEST) > -1)
            cMsg_SetCursor = Abs(SearchMessageArray(WM_SETCURSOR) > -1)
        End If
    End If
           
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.SearchMessageArray
' DateTime  : 9/25/2005
' Author    : LaVolpe
' Purpose   : Binary Search routine
'---------------------------------------------------------------------------------------
'
Private Function SearchMessageArray(Criteria As Long) As Long

    ' Binary search routine
    ' Very short, fast & effective routine for finding a value in a sorted array


    On Error GoTo ReturnResult
    Dim UB As Long, LB As Long, MB As Long, lRtn As Long
    
    lRtn = -1
    ' UB=upper bound of the array segment to search
    ' LB=lower bound of the array segment to search
    ' MB=middle point of the UB & LB values
    
    If Not IsArrayEmpty(Not cClientMsgs) Then
    
        UB = UBound(cClientMsgs)
        MB = UB \ 2 + LB
    
        ' the logic is real simple...
        ' start at the middle of the array & see if the the value is found
        ' if not then divide the next section in half & try again... loop until done
        ' Tip: In large arrays, it may be better to place an initial search for the
        '       1st or last record to see if they match since these are always the
        '       last to be checked; especially if the match is expected in that position
    
        Do
            If cClientMsgs(MB) = Criteria Then Exit Do
            If cClientMsgs(MB) > Criteria Then UB = MB - 1 Else LB = MB + 1
            MB = (UB - LB) \ 2 + LB
        Loop Until LB >= UB
        If cClientMsgs(MB) = Criteria Then lRtn = MB

    End If
    
ReturnResult:
    SearchMessageArray = lRtn
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.ShellSortMessages
' DateTime  : 9/25/2005
' Author    : LaVolpe
' Purpose   : Pretty quick sort routine
'---------------------------------------------------------------------------------------
'
Private Sub ShellSortMessages()
  
  Dim lLoop1 As Long
  Dim lHold As Long
  Dim lHValue As Long
  Dim lTemp As Long

  lHValue = LBound(cClientMsgs)
  Do
    lHValue = 3 * lHValue + 1
  Loop Until lHValue > UBound(cClientMsgs)
  Do
    lHValue = lHValue \ 3
    For lLoop1 = lHValue + LBound(cClientMsgs) To UBound(cClientMsgs)
      lTemp = cClientMsgs(lLoop1)
      lHold = lLoop1
      Do While cClientMsgs(lHold - lHValue) > lTemp
        cClientMsgs(lHold) = cClientMsgs(lHold - lHValue)
        lHold = lHold - lHValue
        If lHold < lHValue Then Exit Do
      Loop
      cClientMsgs(lHold) = lTemp
    Next lLoop1
  Loop Until lHValue = LBound(cClientMsgs)
  
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.MinimizeAnimated
' DateTime  : 9/24/2005
' Author    : LaVolpe
' Purpose   : Make minimized window minimize towards the system tray
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Function MinimizeAnimated(wMsg As Long, wParam As Long, lParam As Long) As Boolean

    Dim rcTray As RECT, rcAni As RECT, rcScreen As RECT, trayHwnd As Long
    Dim wP As WINDOWPLACEMENT, bAbort As Boolean
    
    ' if not minimizing window and the window is not already in the system tray,
    ' then we have nothing to do here
    bAbort = ((mWindowState And 4) = 0 And wParam <> SC_MINIMIZE)
    ' abort this routine if the window doesn't have the following property set
    If MinimizeToSysTray = False Then bAbort = True
    
    If Not bAbort Then
        ' find the system taskbar if it exists
        trayHwnd = FindWindow("Shell_TrayWnd", vbNullString)
        If trayHwnd <> 0 Then
            ' now find the system tray if it exists
            trayHwnd = FindWindowEx(trayHwnd, 0, "TrayNotifyWnd", vbNullString)
            If trayHwnd <> 0 Then
                SystemParametersInfo SPI_GETWORKAREA, 0, rcScreen, 0&
                Select Case wParam
                Case SC_MAXIMIZE
                    rcAni = rcScreen
                Case SC_MINIMIZE
                    rcAni = mWinRect
                Case Else
                    ' get the restore rectangle managed by windows
                    wP.Length = Len(wP)
                    GetWindowPlacement mHWnd, wP
                    rcAni = wP.rcNormalPosition
                    ' the restore window is in workspace coordinates unless the
                    ' ws_ex_toolwindow style is set (NA for this DLL).
                    ' So we convert to screen coordinates
                    OffsetRect rcAni, -rcScreen.Left, -rcScreen.Top
                End Select
                
                ' get the system tray coordiantes & size
                GetWindowRect trayHwnd, rcTray
                If rcTray.Right > rcTray.Left Then
                    ' use a 2x2 rectangle as the source or destination Rect
                    ' depending on if we are minimizing to tray or restoring from tray
                    rcTray.Left = (rcTray.Right - rcTray.Left - 4) \ 2 + rcTray.Left
                    rcTray.Right = rcTray.Left + 2
                    rcTray.Top = (rcTray.Bottom - rcTray.Top - 4) \ 2 + rcTray.Top
                    rcTray.Bottom = rcTray.Top + 2
                End If
                
                ' DrawAnimatedRects will fail to draw the animation if the window
                ' uses a WindowRgn and O/S is less than XP
                If wParam = SC_MINIMIZE Then
                    ' regardless if DrawAnimatedRects fails or not, hide the window
                    mWindowState = SIZE_MINIMIZED Or 4    ' flag as in the system tray
                    DrawAnimatedRects mHWnd, &H3, rcAni, rcTray
                    ShowWindow mHWnd, 0 ' hide the window
                    If ShowOnTaskbar = False Then
                        ' remove the minimized window from the desktop
                        LockWindowUpdate mHWnd
                        ShowWindow mHWnd, &H7   ' minimize no focus
                        ShowWindow mHWnd, 0&
                        LockWindowUpdate 0&
                    End If
                    wMsg = WM_SIZE
                    wParam = mWindowState
                Else
                    mWindowState = mWindowState And Not 4
                    DrawAnimatedRects mHWnd, &H3, rcTray, rcAni
                    wMsg = WM_SIZE
                    If wParam = SC_MAXIMIZE Then
                        ShowWindow mHWnd, &H3   ' show maximized
                        wParam = SIZE_MAXIMIZED
                    Else ' restoring or no change
                        ShowWindow mHWnd, &H1   ' show restored state
                        wParam = SIZE_RESTORED
                    End If
                End If
                RemoveFormMenu  ' VB can put a menu back when toggling visibility
                MinimizeAnimated = True ' abort the wm_syscommand message
            End If
        End If
        
    Else
                
        ' ensure this is set when the window is minimized outside of the sys tray
        If wParam = SC_MINIMIZE Then mWindowState = SIZE_MINIMIZED
        
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.Class_Initialize & CustomWindow.Class_Terminate
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : initialize & clean up when this class is created or terminates
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    ' this call will initialize any uninitialized classes
    Me.Titlebar.noRedraw = False
End Sub

Private Sub Class_Terminate()
    'clean up
    mClientImplementation = 0
    If mHWnd Then               ' class is being set to Nothing explicitly, otherwise
        Dim hWnd As Long        '   when the form is destroyed/closed, mHwnd would = 0
        hWnd = mHWnd            ' cache cause next routine will erase it
        Call SubclassTerminated(mHWnd)  ' unsubclass window & reset to pre-skinned GUI
        If mWinMenu Then SetMenu hWnd, mWinMenu
        SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
    Else
        Set mMenuBar = Nothing
        Set mTitleBar = Nothing
        Set mFrame = Nothing
        Set mGraphics = Nothing
        Set mControls = Nothing
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.ExportCustomSettings
' DateTime  : 9/25/2005
' Author    : LaVolpe
' Purpose   : Development aid, saving customization settings
' Comments  : Will save the custom window settings to a binary file that can be
'             uploaded and applied. Note that this cannot save callback code you
'             may have in your form. Just saves visual settings/styles/options
'             including any non-owner drawn images, system tray information,
'             user button properites (except images/captions), tracking rectangles, etc
'---------------------------------------------------------------------------------------
'
Public Function ExportCustomSettings(ByVal PathFile As String, _
                Optional ExportType As eExportConstants = exAll, _
                Optional propBag As PropertyBag) As Boolean
                        
    ' Parameters
    ' PathFile is a valid path & file name to save the settings to
    ' ExportType - you may not want the background images (titlebar, menubar, inset, etc)
    '   to be exported. You have the option to exclude them
    ' PropBag is optional. If desired, create your own property bag, add whatever you
    '       want to it (i.e., images, fonts, cursors, strings, etc) then pass it
    '       How do you create a property bag in VB6?
    '           Dim myPropBag As PropertyBag
    '           Set myPropBag = New PropertyBag
    '           Then use the .ReadProperty & .WriteProperty methods to save your data
    
    ' exit if minimum requirements not met
    If mHWnd = 0 Then Exit Function
    If PathFile = "" Then Exit Function
    
    Dim pBag As PropertyBag
    Dim iBag As iPropertyBag
    
    Dim fNr As Integer
    On Error Resume Next
    ' see if the passed path can be accessed
    fNr = FreeFile()
    Open PathFile For Binary As #fNr
    Put #fNr, , 0
    Close #fNr
    If Err Then
        ' nope, no can do - exit
        Err.Clear
        Exit Function
    End If
    Kill PathFile
    If Err Then
        ' this error shouldn't happen unless we don't have access to delete the file
        Err.Clear
        Exit Function
    End If
    ' so far so good, let's open the file for real
    fNr = FreeFile()
    Open PathFile For Binary Access Read Write Lock Read Write As #fNr
    If Err Then
        Err.Clear
        Exit Function
    End If
    
    
    ' here we go... Create a new property bag so we can abuse it ;)
    Set pBag = New PropertyBag
    
    ' cache this class's properties
    pBag.WriteProperty "Me.Ver", 1  ' expect future version to have some enhancements
    pBag.WriteProperty "Me.Type", ExportType, exAll
    pBag.WriteProperty "Me.Options", mOptions
    pBag.WriteProperty "Me.Alive", (mActiveState And 8) Or (mActiveState And 4) Or (mActiveState And 2), 0
    If Me.AlwaysOnTop Then pBag.WriteProperty "OnTop", CByte(1), 0
    If Me.ShowOnTaskbar Then pBag.WriteProperty "TaskBar", CByte(1), 0
    If Me.MaximizeFullScreen Then pBag.WriteProperty "FullScreen", CByte(1), 0
    
    ' load the titlebar, system/user button, disabled system button/menus properties
    Set iBag = Me.Titlebar
    iBag.WriteProperties pBag, ExportType
    
    ' load the menubar properties
    Set iBag = mMenuBar
    iBag.WriteProperties pBag, ExportType
    
    ' load the frame, minmax and tray icon properties
    Set iBag = mFrame
    iBag.WriteProperties pBag, ExportType
    
    Set iBag = Nothing
    If Not propBag Is Nothing Then pBag.WriteProperty "UserBag", propBag.Contents(), 0
    ' write the variant array of contents & close the file
    Put #fNr, , pBag.Contents
    Close #fNr

    ' we're done
    ExportCustomSettings = True
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : CustomWindow.ImportCustomSettings
' DateTime  : 8/27/05 22:25
' Author    : LaVolpe
' Purpose   : imports custom settings from a file or a .RES resource
' Comments  : See below
'---------------------------------------------------------------------------------------
'
Public Function ImportCustomSettings(Settings As Variant, _
                iODimplementation As iOwnerDrawn, _
                iTrayImplementation As iSysTrayCallback, _
                Optional propBag As PropertyBag) As Boolean

    ' parameters:
    ' Settings is the data that was created by a prior call to ExportCustomSettings
    '   and can be provided in one of two formats:
    '   1) A full path & filename of that file
    '   2) As a byte array of that file (i.e., data is in RES file & extracted via LoadResData()
    ' iOdimplementation is a valid iOwnerDrawn implementation, usually the form
    '   if the implementation is not used, supply NOTHING as the parameter value
    ' iTrayImplementation is a valid iSysTrayCallBack implementation, usually the form
    '   if the implementation is not used, supply NOTHING as the parameter value
    ' If the Settings has a user-supplied property bag, the parameter will contain
    '   that property bag
    
    Dim vBag As Variant
    Dim pBag As PropertyBag
    Dim iBag As iPropertyBag
    Dim iOD As Long, iSysTray As Long
    Dim vData() As Byte
    Dim Limits As eExportConstants
    
    Dim PathFile As String
    Dim fNr As Long
    Dim I As Integer
    
    Set propBag = Nothing
    If mHWnd = 0 Then Exit Function
    
    On Error Resume Next
    If VarType(Settings) = vbString Then
        ' user is passing string, should be a full path & file name
        ' use simple, unintelligent attempts to open & read the file
        ' relying on error trapping
        fNr = FreeFile()
        Open Settings For Input Access Read As #fNr
        Close #fNr
        If Err Then
            ' can't open the file; invalid string most likely the culprit
            Err.Clear
            Exit Function
        End If
        fNr = FreeFile()    ' test passed, open it for real
        Open Settings For Binary Access Read As #fNr
        Get #fNr, , vBag            ' assign file contents to variant
        Close #fNr
        Set pBag = New PropertyBag  ' assign variant to the property bag contents
        pBag.Contents = vBag
    Else
        ' if not a filename, should be a byte array from a RES file
        If (VarType(Settings) And (vbArray Or vbByte)) = 0 Then Exit Function
        
        vBag = Settings                 ' get the passed array into a variant
        Set pBag = New PropertyBag      ' attempt to assign to the property bag
        pBag.Contents = vBag
        If Err Then
            Err.Clear
            ' when the settings file was loaded into a RES file, it may have a 12byte
            ' header that needs to be removed....
            If UBound(vBag) > 12 Then
                vData = vBag    ' copy the variant into an array
                ' shift the bytes left 12
                CopyMemory vData(0), vData(12), UBound(vData) - 11
                ' redim the array and assign to the property bag
                ReDim Preserve vData(0 To UBound(vBag) - 12)
                pBag.Contents = vData
                Erase vData
            End If
        End If
    End If
    If Err Then
        ' made all valiant attempts to get the user's custom settings & failed
        Err.Clear
        Exit Function
    End If
    ' final check. Is it our property bag?
    If Val(pBag.ReadProperty("Me.Ver", 0)) = 0 Then Exit Function
    
    vBag = Empty
    Me.noRedraw = True
    
    ' see if user passed the required implementation references
    If Not iODimplementation Is Nothing Then iOD = ObjPtr(iODimplementation)
    If Not iTrayImplementation Is Nothing Then
        iSysTray = ObjPtr(iTrayImplementation)
        mTrayImplementation = iSysTray          ' update local pointer
    End If
    
    ' load the custom settings main option & any limitations on importing
    mOptions = pBag.ReadProperty("Me.Options", 0)
    Limits = pBag.ReadProperty("Me.Type", exAll)
    ' load the KeepAlive setting
    mActiveState = pBag.ReadProperty("Me.Alive", 0)
    If (mActiveState And 2) = 2 Then
        ' overriding is in play
        If (mActiveState And 8) = 8 Then
            Me.SetKeepActive True, True, True   ' always active 100% of the time
        Else
            If (mActiveState And 4) = 4 Then    ' always active in our thread only
                Me.SetKeepActive True, True, False
            Else                                ' shouldn't get here, but if we do...
                Me.SetKeepActive True, False, False
            End If
        End If
    Else                                        ' no overriding in play
        Me.SetKeepActive True, False, False
    End If
    ' pass the information to the other classes which will cascade to
    ' to their child classes, grandchildren, etc
    Set iBag = Me.Titlebar
    iBag.ReadProperties pBag, Limits, iOD, iSysTray
    Set iBag = mMenuBar
    iBag.ReadProperties pBag, Limits, iOD, iSysTray
    Set iBag = mFrame
    iBag.ReadProperties pBag, Limits, iOD, iSysTray
    Set iBag = Nothing

    ' logical assumption applied here
    If (mOptions And 2) = 2 Then ' minimize to system tray option
        ' as we are importing, don't allow this setting on import
        ' unless there is a system tray
        If mFrame.SysTrayInitialized = False Then mOptions = mOptions And Not 2
    End If
    
    ' set these properties if needed
    Me.AlwaysOnTop = (pBag.ReadProperty("OnTop", 0) > 0)
    Me.MaximizeFullScreen = (pBag.ReadProperty("FullScreen", 0) > 0)
    Me.ShowOnTaskbar = (pBag.ReadProperty("Taskbar", 0) > 0)
    
    ' return a user-supplied property bag if applicable
    If Not IsMissing(propBag) Then
        
        vBag = pBag.ReadProperty("UserBag", 0)
        If (VarType(vBag) And vbArray) = vbArray Then
            ' user propertybag supplied, return it
            Set propBag = New PropertyBag
            propBag.Contents = vBag
        End If
    
    End If
    
    Set pBag = Nothing
    ImportCustomSettings = True
    
End Function

