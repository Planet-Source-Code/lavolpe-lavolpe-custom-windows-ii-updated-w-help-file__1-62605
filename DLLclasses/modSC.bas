Attribute VB_Name = "modSC"
Option Explicit

' Common functions and declarations for classes in this project
' Read the remarks at top of each class module for comments related to this project
' Start with comments in the CustomWindow class & iCoreMessages interface class

'////////////// KERNEL32 DLL \\\\\\\\\\\\\\\
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) 'used everywhere
Public Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long 'also used in CustomWindow

'////////////// USER32 DLL \\\\\\\\\\\\\\\
Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long 'also used in CustomWindow
Public Declare Function CopyImage Lib "user32.dll" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long 'also used in cGraphics
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long 'also used clsButtons & cGraphics
Public Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long 'also used clsButtons & cGraphics
Public Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long 'used in several classes
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long 'used in several classes
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long ' used in several classes
Public Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long 'also used in CustomWindow
Public Declare Function GetMenu Lib "user32.dll" (ByVal hWnd As Long) As Long 'used in clsMenubar & CustomWindow
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long 'also used in CustomWindow
Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long 'also used in clsBarColors
Public Declare Function GetSystemMenu Lib "user32.dll" (ByVal hWnd As Long, ByVal bRevert As Long) As Long 'used in clsMenubar and clsButtons
Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long 'used in clsMinMax & clsButtons
Public Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long 'also used in several classes
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long 'also used in CustomWindow
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long 'also used in several classes
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long 'used in several classes
Public Declare Function IsRectEmpty Lib "user32.dll" (ByRef lpRect As RECT) As Long 'used in several classes
Public Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long 'used in clsButtons and CustomWindow
Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long 'used in several classes
Public Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long 'also used in other modules/classes
Public Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long 'used in several clasees
Public Declare Function ReleaseCapture Lib "user32.dll" () As Long 'used in clsButtons & CustomWindow
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long 'also used in several classes
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long 'also used in several classes
Public Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long 'used in clsButtons & CustomWindow
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long 'used in several classes
Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long 'also used in CustomWindow
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long 'also used in clsToolTip & CustomWindow
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long 'also used in CustomWindow
Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long

'////////////// GDI32 DLL \\\\\\\\\\\\\\\
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long 'used in several classes
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long 'also used in clsTitleBar
Public Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long 'used in several classes
Public Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long 'used in clsMenubar also
Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long 'used in several classes
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long 'used in several classes
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long 'used in several classes
Public Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long 'also used in clsTitlebar & CustomWindow
Public Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long 'used in several classes
Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As Any) As Long 'used in several classes
Public Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long 'used in several classes
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long 'used in several classes

'////////////// COMMON TYPES \\\\\\\\\\\\\\\
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type LOGFONT             ' used to create memory fonts
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type
Public Type MSG                 ' used when hooking menu messages (MSGF_Menu)
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    Pt As POINTAPI
End Type
Public Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 3) As Byte
End Type
Public Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Public Type NONCLIENTMETRICS     ' used to retrieve/set system settings
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type


Public Type CustomItemDraw_LV
    itemID As Long
    itemPos As Long
    itemData As Long
    itemState As Long
    itemOD As Long
    hdc As Long
    rcItem As RECT
End Type
Public Type BkgAction_LV
    hdc As Long
    rcItem As RECT
    rcExtra As RECT
End Type


'////////////// DLL CONSTANTS \\\\\\\\\\\\\\\
'------ DrawText API
Public Const DT_CALCRECT As Long = &H400   ' all used to calculate or draw text to DC
'------ SetWindowPos API
Public Const SWP_NOSIZE As Long = &H1          ' all are used to move/restyle/size windows
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_NOREDRAW As Long = &H8
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_FRAMECHANGED As Long = &H20
'------ Get/SetWindowLong API
Public Const GWL_WNDPROC As Long = -4          ' used to subclass
Public Const GWL_STYLE As Long = -16           ' used to set/get window styles
Public Const GWL_EXSTYLE As Long = -20         ' used to set/get window extended styles
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_CAPTION As Long = &HC00000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_DLGFRAME As Long = &H400000
Public Const WS_EX_APPWINDOW As Long = &H40000 ' used for toggling in/out of taskbar
Public Const WS_CHILD As Long = &H40000000     ' a specific style. to detect invalid object for subclassing
'------ SetWindowsHookEx/UnhookWindowsHookEx API
Private Const WH_KEYBOARD As Long = 2
Private Const WH_MSGFILTER As Long = -1
Private Const MSGF_MENU As Long = 2
'------ Send/PostMessage API
Public Const WM_USER As Long = &H400          ' used for custom window messages
Public Const WM_APPACTIVATE As Long = &H1C
Public Const WM_CANCELMODE As Long = &H1F      ' to detect when to release capture
Public Const WM_COMMAND As Long = &H111        ' used to post menu item clicks
Public Const WM_CONTEXTMENU As Long = &H7B     ' to detect sysmenu popups
Public Const WM_DESTROY As Long = &H2          ' to unsubclass
Public Const WM_ENTERMENULOOP As Long = &H211  ' to prevent VB from painting NC
Public Const WM_ENTERSIZEMOVE As Long = &H231
Public Const WM_EXITMENULOOP As Long = &H212   ' to clean up window style modifications
Public Const WM_EXITSIZEMOVE As Long = &H232
Public Const WM_GETICON As Long = &H7F
Public Const WM_GETMINMAXINFO As Long = &H24
Public Const WM_GETSYSMENU = &H313             ' to intercept Alt+Space
Public Const WM_INITMENUPOPUP As Long = &H117
Public Const WM_LBUTTONDOWN As Long = &H201    ' track dragging & menu selections
Public Const WM_LBUTTONUP As Long = &H202      ' to detect when to release capture
Public Const WM_MBUTTONDOWN As Long = &H207    ' track menu selections
Public Const WM_MENUCHAR As Long = &H120       ' track Alt+Key menu selections
Public Const WM_MENUSELECT As Long = &H11F     ' track menu item selection
Public Const WM_MOUSEMOVE As Long = &H200      ' to detect where to move/resize window
Public Const WM_MOVE As Long = &H3             ' to track window location
Public Const WM_NCACTIVATE As Long = &H86      ' to detect window active status
Public Const WM_NCCALCSIZE As Long = &H83      ' to customize NC area
Public Const WM_NCRBUTTONUP As Long = &HA5     ' track dragging & menu selections
Public Const WM_NCHITTEST As Long = &H84       ' to determine where mouse is on NC
Public Const WM_NCLBUTTONDBLCLK As Long = &HA3 ' to detect restore/maximize
Public Const WM_NCLBUTTONDOWN As Long = &HA1   ' to detect moving/sizing/menu selection
Public Const WM_NCMOUSEMOVE As Long = &HA0     ' track moving/dragging
Public Const WM_NCPAINT As Long = &H85         ' to paint the NC
Public Const WM_NCLBUTTONUP As Long = &HA2     ' to detect when to release capture
Public Const WM_NCRBUTTONDOWN As Long = &HA4   ' track system menu popups
Public Const WM_NCMBUTTONDOWN As Long = &HA7
Public Const WM_NCMBUTTONUP As Long = &HA8
Public Const WM_NCXBUTTONDOWN As Long = &HAB   ' track menu item highlighting reset
Public Const WM_RBUTTONDOWN As Long = &H204    ' track menu item highlighting reset
Public Const WM_RBUTTONUP As Long = &H205      ' track system menu popups
Public Const WM_SETCURSOR As Long = &H20       ' to set cursor for borders
Public Const WM_SETICON As Long = &H80         ' to detect window icon changes including Me.Icon
Public Const WM_SETTEXT As Long = &HC          ' to detect Caption changes except for Me.Caption
Public Const WM_SIZE As Long = &H5             ' used to track window state
Public Const WM_SYSCOMMAND As Long = &H112     ' to reroute sysmenu commands
Public Const WM_XBUTTONDOWN As Long = &H20B    ' track release of dragging
'----- used with wm_nccalcsize
Public Const WVR_VALIDRECTS As Long = &H400    ' rtn value for WM_NCCALCSIZE
'------ determines properties of a menu item
Public Const MF_MOUSESELECT As Long = &H8000&
Public Const MF_SYSMENU As Long = &H2000&
Public Const MF_POPUP As Long = &H10&
Public Const MF_DISABLED As Long = &H2&
'------ used with wm_syscommand
'Public Const SC_CLOSE As Long = &HF060&        ' all are used to detect sysmenu actions
Public Const SC_KEYMENU As Long = &HF100&
'Public Const SC_MAXIMIZE As Long = &HF030&
'Public Const SC_MINIMIZE As Long = &HF020&
Public Const SC_MOUSEMENU As Long = &HF090&
'Public Const SC_MOVE As Long = &HF010&
'Public Const SC_RESTORE As Long = &HF120&
'Public Const SC_SIZE As Long = &HF000&
'------ used with TrackPopumenuEx API
Public Const TPM_RETURNCMD As Long = &H100&
Public Const TPM_BOTTOMALIGN As Long = &H20&
Public Const TPM_RIGHTALIGN As Long = &H8&
'------ used with GetSystemMetrics API
Public Const SM_CYSMICON = 50
Public Const SM_CXSIZE As Long = 30
Public Const SM_CYSIZE As Long = 31
Public Const SM_CXSMSIZE As Long = 52
Public Const SM_CYSMSIZE As Long = 53
'------ used with wm_nchittest
Public Const HTBORDER As Long = 18             ' all are used to detect where cursor is
Public Const HTBOTTOM As Long = 15
Public Const HTBOTTOMLEFT As Long = 16
Public Const HTBOTTOMRIGHT As Long = 17
Public Const HTCAPTION As Long = 2
Public Const HTCLIENT As Long = 1
Public Const HTCLOSE As Long = 20
Public Const HTERROR As Long = -2
Public Const HTLEFT As Long = 10
Public Const HTMAXBUTTON As Long = 9
Public Const HTMENU As Long = 5
Public Const HTMINBUTTON As Long = 8
Public Const HTNOWHERE As Long = 0
Public Const HTRIGHT As Long = 11
Public Const HTSYSMENU As Long = 3
Public Const HTTOP As Long = 12
Public Const HTTOPLEFT As Long = 13
Public Const HTTOPRIGHT As Long = 14
'------ LoadCursor API
Private Const IDC_SIZEALL As Long = 32646&
Private Const IDC_ARROW As Long = 32512&        ' all are default system cursor indexes
Private Const IDC_SIZENESW As Long = 32643&
Private Const IDC_SIZENWSE As Long = 32642&
Private Const IDC_SIZENS As Long = 32645&
Private Const IDC_SIZEWE As Long = 32644&

'////////////// CUSTOM CONSTANTS/VARIABLES \\\\\\\\\\\\\\\
'------ Set/GetProp
Public Const lvWndProcName As String = "lvcwProc"
Private Const lvWndClient As String = "lvcwClient"
Public Const WM_LVPopup = WM_USER + 3762
'====== Subclassing Instance
' Custom, used to restore sys tray icons after Explorer crash & used w/above API
Public WM_IECrashNotify As Long
Public Const WM_TrayNotify As Long = WM_USER + &H1962
Public Const HTMenu_Custom As Long = 100 ' can't let VB have a HTMENU so we tweak a bit
Public Const HTNC_Custom As Long = 99    ' in our NC inset
Public Const HTNC_Tracker As Long = 98
Public Const HTNC_Caption As Long = 97
Public commonCursor(0 To 5) As Long
Private hHookClient As Long     ' pointer to class using a keyboard hook
Private hKybdHookOld As Long    ' old hook handle
Private hMouseHookOld As Long   ' old hook handle

Private cInit As clsInit

'---------------------------------------------------------------------------------------
' Procedure : modSC.Main
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : Used to help determine when this DLL is completely unloaded
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub Main()
    If cInit Is Nothing Then Set cInit = New clsInit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : modSC.SubclassWindow
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : subclasses a window or control
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function SubclassWindow(ByVal hWnd As Long, ByVal cInterface As Long, ByVal isCtrl As Boolean) As Boolean
' routine is called by the CustomWindow class to subclass a window or a control

Dim hWndproc As Long, I As Integer
Dim sClass As String, lRtn As Long

' get previous window procedure if it exists and unsubclass it first
hWndproc = GetProp(hWnd, lvWndProcName)
If hWndproc <> 0 Then
    SetWindowLong hWnd, GWL_WNDPROC, hWndproc
    RemoveProp hWnd, lvWndProcName
End If
    
If isCtrl = False Then
    ' Validity checks...this version is NOT MDI-friendly
        If FindWindowEx(hWnd, 0, "MDIClient", "") Then Exit Function
        '^^ MDI Parent form, don't subclass. Version not compatible
        sClass = String$(256, 0)
        lRtn = GetClassName(GetParent(hWnd), sClass, 255)
        If lRtn Then
            If InStr(1, sClass, "MDIClient", vbTextCompare) = 1 Then Exit Function
        '   ^^ MDI Child check, don't subclass. Version not compatible
        End If
        lRtn = GetWindowLong(hWnd, GWL_STYLE)
        If (lRtn Or WS_CHILD) = lRtn Then Exit Function
        '^^ don't subclass child windows/controls unless known about it
        ' You should be using the DLL's Subclass_OtherWindow function
    ' ^^End of Validity checks. If you persist in trying to do what this
    '   isn't designed to do; odds are that your application will crash
    
    ' force the following styles on the subclassed window
    lRtn = GetWindowLong(hWnd, GWL_STYLE)
    lRtn = lRtn Or WS_CAPTION Or WS_DLGFRAME Or WS_MAXIMIZEBOX _
            Or WS_MINIMIZEBOX Or WS_SYSMENU Or WS_THICKFRAME
    SetWindowLong hWnd, GWL_STYLE, lRtn
    ForceRefresh hWnd
End If


' Load system cursors; these are system shared resources, so we don't/can't destroy them
If commonCursor(0) = 0 Then
    ' the cursors are used for borders, menubar, system menu & titlebars (WM_SETCURSOR)
    commonCursor(0) = LoadCursor(0&, IDC_ARROW)
    commonCursor(1) = LoadCursor(0&, IDC_SIZENESW)
    commonCursor(2) = LoadCursor(0&, IDC_SIZENS)
    commonCursor(3) = LoadCursor(0&, IDC_SIZENWSE)
    commonCursor(4) = LoadCursor(0&, IDC_SIZEWE)
    commonCursor(5) = LoadCursor(0&, IDC_SIZEALL)
End If

' subclass the form
SetProp hWnd, lvWndClient, cInterface
SetProp hWnd, lvWndProcName, SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ProcessMessage)
SubclassWindow = True

End Function
Public Sub ForceRefresh(hWnd As Long, Optional flags As Long)
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or flags
End Sub
'---------------------------------------------------------------------------------------
' Procedure : modSC.SetInputHook
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : keyboard hook used only when the menubar becomes activated
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub SetInputHook(bSet As Boolean, callingClass As Long)
' Keyboard hook and menu hook when the menubar is being accessed
' Only one window will be allowed to hook the keyboard (the active window)

Dim hookAddr As Long
If hKybdHookOld Then ' currently existing hook; remove it
    UnhookWindowsHookEx hKybdHookOld
    hKybdHookOld = 0
End If
If hMouseHookOld Then
    Debug.Print "released hooks"
    UnhookWindowsHookEx hMouseHookOld
    hMouseHookOld = 0
End If
If bSet Then
    If callingClass <> 0 Then
        hHookClient = callingClass
        hookAddr = ReturnAddressOf(AddressOf KeybdFilterProc)
        hKybdHookOld = SetWindowsHookEx(WH_KEYBOARD, hookAddr, 0, GetCurrentThreadId())
        hookAddr = ReturnAddressOf(AddressOf MenuFilterProc)
        hMouseHookOld = SetWindowsHookEx(WH_MSGFILTER, hookAddr, 0, GetCurrentThreadId())
        Debug.Print "set hooks"
    End If
End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : modSC.SetStopTimer
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : simple timer creation/destruction stub
'---------------------------------------------------------------------------------------
'
Public Sub SetStopTimer(bSet As Boolean, hWnd As Long, TimerID As Long, Duration As Long)
    If bSet Then
        SetTimer hWnd, TimerID, Duration, AddressOf TimerCallback
    Else
        If TimerID Then KillTimer hWnd, TimerID
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : modSC.ProcessMessage
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : the receiving point for all subclassed windows
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Function ProcessMessage(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' all subclassed messages come thru this single point of failure

    Dim cClient As CustomWindow
    Dim cInst_Client As Long
    Dim hWndproc As Long
    Dim nRtnVal As Long
    
    ' get the previous window procedure handle & class instance associated for this hWnd
    hWndproc = GetProp(hWnd, lvWndProcName)
    cInst_Client = GetProp(hWnd, lvWndClient)
    
    ' sanity checks
    If hWndproc = 0 Or cInst_Client = 0 Then Exit Function
    ' return a reference to the associated class
    If GetObjectFromPointer(cInst_Client, cClient) = False Then Exit Function
    
    If wMsg = WM_DESTROY Then
        ' auto-unsubclassing as long as END button/statement not used.
        ' hide the window before destroying it. That way you won't see it revert back
        ' to non-skinned appearance. Optionally, you can minimize it instead. Whatever
        Call cClient.SubclassTerminated(hWnd)
        RemoveClient hWnd
        nRtnVal = CallWindowProc(hWndproc, hWnd, wMsg, wParam, lParam)
    Else
        ' send message to processor. Options exist in the class that allows
        ' users to modify the message before or after the class processes them
        nRtnVal = cClient.ProcessMessage(hWnd, wMsg, wParam, lParam, hWndproc)
    End If
    ' clean up & return then message value
    Set cClient = Nothing
    ProcessMessage = nRtnVal
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : modSC.TimerCallback
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : callback used for SetTimer
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Function TimerCallback(ByVal hWnd As Long, ByVal uMsg As Long, ByVal uID As Long, ByVal uElapsed As Long) As Long
' A timer is used to detect when the mouse leaves a menubar after a menu item
' has been highlighted by the mouse. Simply checking the WM_NCHITTEST message
' isn't good enough as the mouse can be scrolled so fast as to not register
' the message. If the timer isn't used a left-over highlighted menu item
' can be seen even though the mouse is nowhere near the menubar

    On Error Resume Next
    Dim cClient As iTimer
    ' return a reference to the associated class
    GetObjectFromPointer uID, cClient
    If Err.Number = 0 Then
        ' call the function that processes menubar mouse messages
        Call cClient.TimerEvent(hWnd, uID)
    End If
    If Err Then
        KillTimer hWnd, uID
        Err.Clear
    End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : modSC.KeybdFilterProc
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : callback used for the keyboard hook
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Function KeybdFilterProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

If hKybdHookOld = 0 Then
    ' hook was released; this should NEVER happen w/o my code purposely doing this
    KeybdFilterProc = 1

Else
    If ncode > -1 Then  ' per MSDN always forward & don't process if nCode=-1
    
        Dim tgtClass As clsMenubar
        If GetObjectFromPointer(hHookClient, tgtClass) Then
            ' pass the key to the menubar class
            If tgtClass.SetKeyBdAction(wParam, lParam) = True Then
                ' don't want the key forwarded
                KeybdFilterProc = 1
                Exit Function
            End If
        End If
    
    End If
    ' pass the key
    KeybdFilterProc = CallNextHookEx(hKybdHookOld, ncode, wParam, lParam)

End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : modSC.MenuFilterProc
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : callback for the message filter hook
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Function MenuFilterProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

If hMouseHookOld = 0 Then
    ' hook was released; this should NEVER happen w/o my code purposely doing this
    MenuFilterProc = 1

Else
    If ncode = MSGF_MENU Then ' only process menu messages
        
        Dim tgtClass As clsMenubar
        If GetObjectFromPointer(hHookClient, tgtClass) Then
            ' pass the key to the menubar class
            If tgtClass.SetMouseAction(wParam, lParam, -1) = True Then
                ' don't want the mouse message forwarded
                MenuFilterProc = 1
                Exit Function
            End If
        End If
    
    End If
    
    ' pass the message
    MenuFilterProc = CallNextHookEx(hMouseHookOld, ncode, wParam, lParam)
End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : modSC.GetObjectFromPointer
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : function to convert a valid pointer into the passed object type
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function GetObjectFromPointer(ByVal oPtr As Long, outObject As Object) As Boolean
    If oPtr Then
        On Error Resume Next
        Dim tgtObject As Object
        CopyMemory tgtObject, oPtr, &H4
        If Err.Number = 0 Then
            Set outObject = tgtObject
            CopyMemory tgtObject, 0&, &H4
        Else
            Set tgtObject = Nothing
        End If
        GetObjectFromPointer = True
    End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : modSC.ReleaseSubClassing
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : unsubclasses a subclassed window/control
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function ReleaseSubClassing(ByVal hWnd As Long)
' function called from within the CustomWindow class to ensure
' window properties are updated/removed. Windows can be unsubclassed
' 2 ways: 1) Setting subclassing Class to Nothing, 2) Closing the window

Dim hWndproc As Long
    If hWnd Then
        hWndproc = GetProp(hWnd, lvWndProcName)
        If hWndproc Then
            ' if we dragging when trying to release subclassing, it would
            ' proabbly fail & crash app. Ensure not dragging...
            ' Things I test for :) Had to be creative by firing a timer that
            ' unloads while I was dragging the subclassed window
            If GetCapture() = hWnd Then ReleaseCapture
            ' unsubclass now & remove key custom window properties
            SetWindowLong hWnd, GWL_WNDPROC, hWndproc
            RemoveProp hWnd, lvWndProcName
            RemoveProp hWnd, lvWndClient
        End If
    End If
    
End Function
'--------------------------------------------------------------------------------------
' Procedure : modSC.HiWord, modSC.LoWord, modSC.MakeLong
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : used to extract integers from longs or combine integers into longs
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function HiWord(LongIn As Long) As Integer
  Call CopyMemory(HiWord, ByVal VarPtr(LongIn) + 2, 2)
End Function
Public Function LoWord(LongIn As Long) As Integer
  Call CopyMemory(LoWord, ByVal VarPtr(LongIn), 2)
End Function
Public Function MakeLong(ByVal LoWordIn As Integer, ByVal HiWordIn As Integer) As Long
  MakeLong = CLng(LoWordIn)
  Call CopyMemory(ByVal VarPtr(MakeLong) + 2, HiWordIn, 2)
End Function
'---------------------------------------------------------------------------------------
' Procedure : modSC.IsArrayEmpty
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : quick & dirty way to determine if an array has been initialized
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function IsArrayEmpty(ByVal lArrayPointer As Long) As Boolean
  ' test to see if an array has been initialized
  ' Cannot be used on variants
  IsArrayEmpty = (lArrayPointer = -1)
End Function
'---------------------------------------------------------------------------------------
' Procedure : modSC.ReturnAddressOf
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : used to cache a value of an AddressOf pointer
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Private Function ReturnAddressOf(lAddress As Long) As Long
    ReturnAddressOf = lAddress
End Function

