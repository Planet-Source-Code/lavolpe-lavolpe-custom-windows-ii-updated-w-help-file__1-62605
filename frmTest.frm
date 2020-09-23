VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmTest 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Test Window"
   ClientHeight    =   3555
   ClientLeft      =   165
   ClientTop       =   1020
   ClientWidth     =   4530
   FillColor       =   &H0080FFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3915
      Top             =   375
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   120
      ScaleHeight     =   1020
      ScaleWidth      =   4215
      TabIndex        =   1
      Top             =   2355
      Visible         =   0   'False
      Width           =   4275
      Begin VB.Image imgBdr 
         Height          =   1650
         Index           =   1
         Left            =   0
         Picture         =   "frmTest.frx":0000
         Top             =   -675
         Width           =   1650
      End
      Begin VB.Image imgBdr 
         Height          =   1650
         Index           =   0
         Left            =   135
         Picture         =   "frmTest.frx":0CC7
         Top             =   75
         Width           =   1650
      End
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   3810
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":16A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":239E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":26B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":33AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3800
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":40A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":41FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4358
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   2025
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmTest.frx":68A6
      Top             =   0
      Width           =   3675
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   0
         Begin VB.Menu mnuOpen 
            Caption         =   "This File"
            Index           =   0
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "That File"
            Index           =   1
            Shortcut        =   ^B
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "C&lose"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Edit"
      Index           =   1
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&View"
      Index           =   2
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Project"
      Index           =   3
   End
   Begin VB.Menu mnuMain 
      Caption         =   "For&mat"
      Index           =   4
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Debu&g"
      Index           =   5
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug Sub"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "R&un"
      Index           =   6
      Begin VB.Menu mnuRun 
         Caption         =   "Run Sub"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Que&ry"
      Index           =   7
      Begin VB.Menu mnuQuery 
         Caption         =   "Query Sub"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Help"
      Index           =   8
      Begin VB.Menu mnuHelp 
         Caption         =   "&Contents"
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Inde&x"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Remote"
      Enabled         =   0   'False
      Index           =   9
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "na"
      Visible         =   0   'False
      Begin VB.Menu mnuPU 
         Caption         =   "Do something"
         Index           =   0
      End
      Begin VB.Menu mnuPU 
         Caption         =   "Do something else"
         Index           =   1
      End
      Begin VB.Menu mnuPU 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPU 
         Caption         =   "Close"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' used to activate PSC hyperlink
Private Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' general purpose
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long


' used for subclassing the text box example
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const WM_PRINT As Long = &H317
Private Const PRF_CHECKVISIBLE As Long = &H1&
Private Const PRF_CLIENT As Long = &H4&
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function BeginPaint Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetUpdateRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved As Byte
End Type
Private Const WM_PAINT As Long = &HF&
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_LBUTTONDOWN As Long = &H201

' used to create clipping regions
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByRef lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long

Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' used to set custom cursor
Private Const IDC_HAND As Long = 32649
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long


Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_DISABLED As Long = &H2&
Private Const MF_GRAYED As Long = &H1&
Private Const MF_DEFAULT As Long = &H1000&
Private Const WM_SYSCOMMAND As Long = &H112

' custom UDTs to owner draw the Nonclient
Private Type CustomItemDraw_LV
    itemID As Long
    itemPos As Long
    itemData As Long
    itemState As Long
    itemOD As Long
    hdc As Long
    rcItem As RECT
End Type
Private Type BkgAction_LV
    hdc As Long
    rcItem As RECT
    rcExtra As RECT
End Type

Private Const WM_USER As Long = &H400
'balloon tip notification messages, only received in XP & higher
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Const WM_RBUTTONUP As Long = &H205
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOW As Long = 5


Private polyPts(0 To 8) As POINTAPI     ' create custom menu hover border

Public lvCW As CustomWindow             ' our DLL reference


' all implementations are optional, however, they are expected in some functions and
' should you try to call a DLL function without adding the implementation, you
' most likely will get a generic error something like "... With Block not Set"

Implements iOwnerDrawn                  ' allows owner drawing non client
Implements iControlsCallback            ' allows subclassing child controls
Implements iSysTrayCallback             ' allows callbacks for the system tray
Implements iCoreMessages                ' allows you to subclass the subclassed window

Private Sub Form_Load()
    
    Set lvCW = New CustomWindow     ' create instance of the DLL
    
    ' add gee-whiz text to the textbox
    Me.Text1 = "As you click on the different samples in the other window, " & _
    "this text box will contain some additional information and/or tips." & _
    vbCrLf & vbCrLf & "The project is way too big and flexible to " & _
    "try to show all possible scenarios and options you have. " & _
    vbCrLf & vbCrLf & "Enjoy using this & as always, provide feedback :)"
    
    mnuMain(1).Enabled = False
    Timer1.Enabled = True           ' change caption & form icon after 10 seconds


    Dim hSysM As Long, poM As Long
    hSysM = GetSystemMenu(hwnd, 0)
    poM = CreatePopupMenu()
    AppendMenu hSysM, MF_SEPARATOR Or MF_DISABLED, 11, ""
    AppendMenu hSysM, &H10&, poM, "System Menu Add On"
    AppendMenu poM, MF_DEFAULT, 1, "Item #1"
    AppendMenu poM, MF_DISABLED Or MF_GRAYED, 2, "Item #2"

    lvCW.MessagesAdd Me, WM_SYSCOMMAND
    ' want to see if our custom system menu item will be clicked

End Sub


Private Sub iControlsCallback_WindowMessage(ByVal hwnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Long, bBlockMessage As Boolean, BlockValue As Long)
    
    ' only a real-short, non-perfect routine to show that the control is being subclassed
    ' This will draw a gradient in the text box.... But 'cause I'm lazy and this is
    ' not really part of my project (example only), I'll unsubclass when user clicks in
    ' the client area.
    
    ' since this callback can be used for multiple child controls, suggest
    ' select casing the hWnd...
    
    Select Case hwnd
    
    Case Text1.hwnd
        Select Case wMsg
        Case WM_PAINT
        
            Dim ps As PAINTSTRUCT
            Dim hdc As Long, tDC As Long, tBmp As Long, tOldBmp As Long
            
            ' per MSDN, always call this function before calling BeginPaint
            If GetUpdateRect(hwnd, ps.rcPaint, 0) = 0 Then Exit Sub
            
            ' have DLL create an offscreen bitmap and DC for us
            lvCW.Graphics.CreateBitmapAndDC ScaleWidth / Screen.TwipsPerPixelX, ScaleHeight / Screen.TwipsPerPixelY, tBmp, tDC, True
            ' select the bitmap into the DC
            tOldBmp = SelectObject(tDC, tBmp)
            ' gradient fill that DC now
            lvCW.Graphics.GradientFillEx tDC, 0, 0, ScaleWidth / Screen.TwipsPerPixelX, ScaleHeight / Screen.TwipsPerPixelY, False, vbGreen, vbCyan
            ' make the background transparent
            SetBkMode tDC, &H3
            ' tell text box to paint to our offscreen DC
            SendMessage hwnd, WM_PRINT, tDC, ByVal PRF_CLIENT Or PRF_CHECKVISIBLE
            
            ' now get the update rectangle for the text box & Blt the portion over
            hdc = BeginPaint(hwnd, ps)
            BitBlt hdc, ps.rcPaint.Left, ps.rcPaint.Top, _
                    ps.rcPaint.Right - ps.rcPaint.Left + 1, _
                    ps.rcPaint.Bottom - ps.rcPaint.Top + 1, tDC, _
                    ps.rcPaint.Left, ps.rcPaint.Top, vbSrcCopy
            ' terminate the paint function & clean up
            EndPaint hwnd, ps
            
            ' replace the original DC bitmap & delete the one we created: 1 step
            DeleteObject SelectObject(tDC, tOldBmp)
            DeleteDC tDC
            
            bBlockMessage = True    ' prevent the wm_paint from being forwarded
            BlockValue = 0  ' we don't need to set this since it is zero anyway
            '^^ always goto MSDN and look up the message you are intercepting
            '   to find out what the return value should be to override defaults
        
        Case WM_LBUTTONDBLCLK, WM_LBUTTONDOWN
            ' unsubclass the window
            lvCW.Unsubclass_OtherWindow hwnd
            bBlockMessage = True
            BlockValue = 0
            Text1.Refresh
        End Select
        
    Case Picture1.hwnd
        ' subclassing code here
    Case Else
        ' etc, etc
    End Select
End Sub


Private Sub iCoreMessages_WindowMessage(ByVal hwnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Long, bBlockMessage As Boolean, BlockValue As Long)

' these are forwarded message per our request using the MessageAdd function
' I only wanted one, so I could see if the system menuitem we added was clicked
    
' If I wanted to use fancy sizing cursors, for example, I would trap the
' WM_SETCURSOR message, so I could assign them as needed

    Select Case wMsg
    Case WM_SYSCOMMAND
        Select Case wParam
        Case 1: ' our enabled system menu addon menu item
            MsgBox "System menu item we added was clicked: Item #1"
        End Select
    End Select

End Sub

Private Sub iOwnerDrawn_OwnerDrawMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dllCaption As String)

' this is the procedure where you will get all your owner-drawn messages
' Depending on what you are owner-drawing, messages either must be handled
' or are optional to handle.

' I've rerouted each call to a separate procedure so that you can look at each
' one individually. Depending on the complexity of your code for owner-drawing
' you may wish to separate your messages into different procedures or just
' pile them all in this procedure:  personal preferences

' Always use an ON ERROR statement in a callback routine. If you don't want to
' "Resume Next", you can GO TO a error handling routine to log or debug.print
' the Err.Description. Errors in callback routines are a sure recipe for crashes
On Error Resume Next


' the iOwnerDrawn class has pretty specific information on how to use these messages

Select Case wMsg

Case omMeasureMenuItem
    ' if we are drawing the complete menubar items, we must tell DLL what size
    ' rectangle we need to display the menu item. If we are drawing menu item
    ' images only, this where we should internally determine which image
    ' will be displayed and tell DLL whether the image will be left or right
    ' aligned with the menu item caption
    DoMeasureItem lParam, dllCaption

Case omDrawMenuItem
    ' draw menu items and/or just the images
    DoDrawMenuItem wParam, lParam, dllCaption
    
Case omDrawMBarBkg, omDrawTBarBkg
    ' draw menubar & titlebar backgrounds
    DoDrawBarBackground wParam, lParam
    
Case omDrawPostNC
    ' add other graphics to the just before it is updated
    ' This is where we will add/update tracking rectangles
    DoDrawPostNonClient wParam, lParam
    
Case omDrawPreNC
    ' here we are drawing custom borders for the form, but we also have
    ' the opportunity to draw the complete background of the window
    DoDrawBorders wParam, lParam
    
Case omDrawSysBtn
    ' we want to custom draw the system buttons: min/max/restore/close
    DoDrawSystemButton wParam, lParam
    
Case omDrawUserBtn
    ' we added custom buttons to the titlebar, we must draw those
    DoDrawCustomButton wParam, lParam
    
Case omUserBtnClick
    ' user clicked one of our custom buttons
    ' Not much of an example. But you get the point
    MsgBox "Button ID # " & wParam & vbNewLine & dllCaption, vbInformation + vbOKOnly
    
Case omDrawTrackRect
    ' we added a tracking rectangle, we can monitor this message to
    ' update the graphics for that tracking rectangle if we want to
    DoDrawTracker wParam, lParam
    
Case omTrackClick
    ' user clicked on a tracking rectangle. What do we want to do?
    DoTrackerClick wParam, lParam
    
Case omTrackCursor
    ' cursor is over our tracking rectangle. Do we want to set a different cursor?
    ' when mouse is over our URL link, we will change cursor to a hand
    Select Case wParam ' id of our tracking rectangle(s)
    Case 62
        wParam = LoadCursor(0&, IDC_HAND)
        If wParam Then
            CopyMemory ByVal lParam, &H1, &H4
            SetCursor wParam
        End If
    End Select

End Select

End Sub

Private Function MarkMenuImage(mnuPos As Long, mnuCaption As String) As Long

' Purpose: Assign a ImageList reference to a specific menu item

Dim m As Long
Dim imgRef As Long

' the following two IF statements should never fire false if you
' use an array index for your top level menus. Obviously this is
' much easier in many levels, not just referencing.

    If mnuPos > mnuMain.LBound - 1 Then
        If mnuPos < mnuMain.UBound + 1 Then
        
            ' need loop 'cause menu items where .Visible=False are not
            ' known by the DLL, so the mnuItem position passed may not
            ' coincide with your array index
            For m = mnuMain.LBound To mnuMain.UBound
            
                If mnuMain(m).Caption = mnuCaption Then
                    imgRef = m + 1
                    MarkMenuImage = imgRef
                    Exit For
                End If
                
            Next
        End If
    End If

End Function

Private Sub iSysTrayCallback_ProcessTrayIcon(ByVal hwnd As Long, ByVal wTrayIconID As Long, wMsg As Long, bOverriden As Boolean, lOverrideReturn As Long)
' this is the event used & sent only for tray icon messages...

' only one hWnd being subclassed for tray icons; no need to Select Case for it

' note that I'm testing for right button up. You could test for right button down
' instead but when using multiple tray icons for same hWnd, right button down
' tends to flash the app's taskbar icon if right clicking on consecutive tray icons,
' whether or not SetForeGroundWindow used.

' if a single icon is used, wm_rbuttondown appears to be problem free

    Select Case wMsg
        Case WM_RBUTTONUP
        ' wParam is the unique id we assigned to the tray icon
        ' 1=1st one loaded, 2=2nd one loaded, 3=3rd one loaded
    
        SetForegroundWindow hwnd ' if we don't do this, the menu may hang if not clicked on
        
        PopupMenu mnuPopup
        
    Case WM_LBUTTONDOWN
        ' when left clicking on icon, show our form
        ' Here I'll use the DLL function which will show the form as
        ' maximized, restored state, or most recent state
        lvCW.ShowWindowFromTrayIcon SIZE_RESTORED
        SetForegroundWindow hwnd ' we shown ourself, let's set focus to ourself
        
    ' following NIN messages are only received in XP & higher
    Case NIN_BALLOONHIDE
        ' your balloon was closed
    Case NIN_BALLOONSHOW
        ' your balloon appeared
    Case NIN_BALLOONTIMEOUT
        ' your balloon closed due to time out
    Case NIN_BALLOONUSERCLICK
        ' your balloon was clicked
    
    End Select


End Sub

Private Sub mnuFile_Click(Index As Integer)
    If Index = 1 Then Unload Me     ' emergency exit should something go wrong
    ' also provides form will unload even if the Close system menu item is disabled
End Sub

Private Sub mnuMain_Click(Index As Integer)
   On Error GoTo mnuMain_Click_Error

    ' testing purposes
    Select Case Index
    Case 2: Debug.Print "View clicked"
    Case 1: Debug.Print "Edit clicked"
    Case 3: Debug.Print "Project clicked"
    Case 4: Debug.Print "Format clicked"
    'Case 6: MsgBox "Run clicked"
    'Case 7: MsgBox "Query clicked"
    Case 9: Debug.Print "Remote clicked"
    End Select

   On Error GoTo 0
   Exit Sub

mnuMain_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuMain_Click of Form frmTest"

End Sub


Private Sub mnuOpen_Click(Index As Integer)
    ' testing purposes
    If Index = 1 Then MsgBox "Got Ctrl Shortcut"
End Sub


Private Sub mnuPU_Click(Index As Integer)
Select Case Index
    Case 0 ' do something
        Me.WindowState = vbNormal
    Case 1 ' do something else
        Me.WindowState = vbNormal
    Case 2 ' separator
    Case 3 ' unload me
        Unload Me
End Select
End Sub

Private Sub Timer1_Timer()
    
    ' change caption and icon
    Timer1.Enabled = False
    lvCW.Titlebar.Caption = "LaVolpe Custom Windows"
    Set Me.Icon = imgLst.ListImages(4).Picture
    
    ' add the new icon to the Alt+Tab window
    lvCW.AppIcon = Me.Icon
    
    ' see if we have a system tray class established (one of the examples)
    If lvCW.Frame.SystemTray(Me.hwnd, Me).isActive Then
        lvCW.Frame.SystemTray(Me.hwnd, Me).Icon = Me.Icon
        lvCW.Frame.SystemTray(Me.hwnd, Me).Tip = Me.Caption
    Else
        ' calling above created a class for us, let's destroy it now
        lvCW.Frame.RemoveSystemTray 0, True
    End If
    
End Sub



Private Sub DoMeasureItem(lParam As Long, dllCaption As String)
'////////////////////// wMsg = omMeasureMenuItem \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To measure a menu item you are drawing as a result of passing
' the odMenuItem_Complete or the odMenuItem_ImageOnly flag to the
' OwnerDrawn function of the DLL

' wParam is not used
' lParam is a pointer to a CustomItemDraw_LV structure
' dllCaption is the menu item's caption

'CustomItemDraw_LV values
'   itemID. The menu item identifier returned by Windows
'   itemPos. The zero-bound position of the item on the menubar
'       ** Note that menu items where .Visible=False are not considered
'          part of the menubar. Tip. Use array for your top level menu items
'   itemData. Any value you provide to help uniquely identify this item
'       ** Tip. Good idea to assign the image reference for your menu item here
'   itemState. 0 if item is enabled or 1 if it is disabled
'   itemOD. either odMenuItem_Complete or odMenuItem_ImageOnly
'       ** When odMenuItem_ImageOnly, no measurement is needed. This is provided
'           so you can tell the DLL the image alignment and also assign the
'           itemData portion of the UDT as you see fit. The return value
'           must contain odImgAlignRight to align image to right of menu item
'           or odImgAlignLeft to algin image left of the menu item, otherwise
'           any other value will be assumed to mean no image
'   hDC. a DC with correct menu font already selected. Do not remove the font
'        or the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem.
'       If itemOD is odMenuItem_Complete then you must fill in the size of the
'           menu item you need to draw your text and any image and/or borders
'       If itemOD is odMenuItem_ImageOnly, not used.

' Final notes:
'   if itemOD = odMenuItem_ImageOnly then itemOD must be changed & updated if drawing an image
'   if itemOD = odMenuItem_Complete then rcItem must be changed & updated
'   for both, itemData should be set to a value that is meaningful to you


Dim CID As CustomItemDraw_LV
Dim bCompleteOwnerDrawn As Boolean

    ' First we will assign images to the menu items
    
    ' get pointer into our structure
    CopyMemory CID, ByVal lParam, Len(CID)
    
    ' caching this 'cause we will be changing it
    bCompleteOwnerDrawn = (CID.itemOD = odMenuItem_Complete)
    
    ' see if we already assigned an image to this menu item
    If CID.itemData = 0 Then
        '^^ this flag is for our use. I am using it to cache a reference to
        ' an image list item, so I can skip the following lines of code if
        ' I've already done it before. Additionally, the value I use here
        ' will be processed when the time comes to draw the item.
    
        ' if not, call a simple routine to associate an image with this item
        CID.itemData = MarkMenuImage(CID.itemPos, dllCaption)
    End If
    ' if the image was assigned, we want to update the CID & pass it back
    If CID.itemData <> 0 Then
        CID.itemOD = odImgAlignLeft
        ' here if we wanted right aligned images we would supply:
        'CID.itemOD = odImgAlignRight
        ' if we didn't want an image at all we would supply:
        'CID.itemOD = 0
        CopyMemory ByVal lParam, CID, Len(CID)
    End If
        
        
    If bCompleteOwnerDrawn Then
        
        ' remember, you are responsible for drawing everything when you passed
        ' the odMenuItem_Complete flag to the OwnerDrawn function of the DLL
        
        ' What does everything mean? Caption, borders, image. So ensure your
        ' measurements give you space for the drawings
        With CID
            ' call function to measure text for us, using the menu font for our window
            lvCW.Graphics.TextMeasureEx .hdc, dllCaption, gMenu_Font, .rcItem.Right, .rcItem.Bottom
            
            ' adjust for a border when items selected/hovered over & images
            .rcItem.Bottom = .rcItem.Bottom + 4         ' add space for top/bottom border
            .rcItem.Right = .rcItem.Right + 4          ' add space for left/right border
            
            ' if we are adding an image, add space for the image to
            If .itemData <> 0 Then .rcItem.Right = .rcItem.Right + 22
            '^^ 16 icon pixels + 4 pixel image/text separation + 2 for raised/sunken effect
            
            ' ensure big enough to display the image & border
            If .rcItem.Bottom < 24 Then .rcItem.Bottom = 24
            '^^ 16 icon pixels + 8 for top/bottom octagon slanted corners (4 pixels each)
            
            ' Note: that the .Left & .Top will always be zero,
            ' changing them has no effect, the Left & Top are eventually
            ' calculated by the DLL and cannot be overriden
        
        End With
        
        ' pass the updated rectangle
        
        CopyMemory ByVal lParam, CID, Len(CID)
    
    End If

End Sub

Private Sub DoDrawMenuItem(wParam As Long, lParam As Long, dllCaption As String)
'////////////////////// wMsg = omDrawMenuItem \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a menu item image or the entire menu item (image
' borders, caption, etc) as a result of passing the odMenuItem_Complete
' or the odMenuItem_ImageOnly flag to the OwnerDrawn function of the DLL

' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a CustomItemDraw_LV structure
' dllCaption is the menu item's caption

'CustomItemDraw_LV values
'   itemID. The menu item identifier returned by Windows
'   itemPos. The zero-bound position of the item on the menubar
'       ** Note that menu items where .Visible=False are not considered
'          part of the menubar. Tip. Use array for you top level menu items
'   itemData. Any value you provided during the omMeasureMenuItem message
'   itemState. will contain one or more of the following values using OR
'               mcStandard - item is not selected, not hovered over
'               mcHover - mouse is currently over the menu item
'               mcSelect - menu item is currently selected
'               mcDisabled - menu item is disabled
'   itemOD. either odMenuItem_Complete or odMenuItem_ImageOnly
'           odMenuItem_ImageOnly. draw your image in the rcItem provided
'           odMenuItem_Complete. draw the entire menu item in the rcItem provided
'   hDC. a DC with correct menu font already selected. Do not remove the font
'        or the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. The rectangle to draw. This rectangle is NOT clipped and you should
'           ensure your drawing stays within its bounds

' Final notes:
'   if itemOD = odMenuItem_ImageOnly then simply draw the image
'   if itemOD = odMenuItem_Complete then draw everything: text, image, borders

Dim CID As CustomItemDraw_LV
Dim imgOffset As POINTAPI
Dim clipRgn As Long

    CopyMemory CID, ByVal lParam, Len(CID) ' get the drawing info
        
    If CID.itemOD = odMenuItem_Complete Then
        ' If we are required to draw every little detail, then do exactly that
        ' Note that the following would be better organized as states,
        ' vs background, caption, & then border. I did it this way
        ' so that there is no confusion as to what steps are required.
        ' You can organize your code anyway you want obviously
        
        imgOffset.X = 24   ' caption offset so image not overdrawn
        
        With CID.rcItem
            
            ' when the form is inactive, we'll draw using standard colors
            If wParam = 0 Then ' inactive window
            
                ' draw the caption, no background will be used
                lvCW.Graphics.TextDrawEx dllCaption, CID.hdc, .Left + imgOffset.X, .Top, .Right, .Bottom, fxFlat, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, gMenu_Font, Nothing, vbGrayText
                
                ' draw the border when hovering over inactive window's menubar
                If (CID.itemState And mcHover) = mcHover Then
                    lvCW.Graphics.DrawPolyLine CID.hdc, vb3DHighlight, .Left + 4, .Bottom, .Left, .Bottom - 4, .Left, .Top + 4, .Left + 4, .Top, .Right - 4, .Top
                    lvCW.Graphics.DrawPolyLine CID.hdc, vb3DShadow, .Right - 4, .Top, .Right, .Top + 4, .Right, .Bottom - 4, .Right - 4, .Bottom, .Left + 3, .Bottom
                End If
                ' Note that you can get as fancy for the inactive window as you want
                    
            Else
            
                ' Note: the background will already be drawn, unless you identified the menubar
                ' to have a transparent background. You can overdraw it anyway at this point
                
                ' Always draw the background first, if needed.
                If (CID.itemState And mcHover) = mcHover Then ' draw hover
                    
                    ' Create the clipping region for an octagon. PolyRegions are tricky to get perfect
                    ' Keep in mind that you always need to extend your region 1 pixel
                    '   to the right & bottom to be able to draw on your right & bottom edges
                    '   Regions, by default, exclude those outer edges
                    polyPts(0).X = .Left + 3: polyPts(1).X = .Left: polyPts(2).X = .Left: polyPts(3).X = .Left + 4
                    polyPts(4).X = .Right - 3: polyPts(5).X = .Right + 1: polyPts(6).X = .Right + 1: polyPts(7).X = .Right - 4
                    polyPts(0).Y = .Bottom + 0: polyPts(1).Y = .Bottom - 4: polyPts(2).Y = .Top + 4: polyPts(3).Y = .Top
                    polyPts(4).Y = .Top: polyPts(5).Y = .Top + 4: polyPts(6).Y = .Bottom - 4: polyPts(7).Y = .Bottom + 1
                    polyPts(8) = polyPts(0)
                    clipRgn = CreatePolygonRgn(polyPts(0), 9, 2)
                    
                    ' now select it into the DC
                    SelectClipRgn CID.hdc, clipRgn
                    ' destroy the region or memory leaks occur
                    DeleteObject clipRgn
                    ' fill the region with gradients
                    lvCW.Graphics.GradientFillEx CID.hdc, .Left, .Top, .Right, .Bottom, True, vbCyan, RGB(128, 128, 128)
                
                Else
                    If (CID.itemState And mcSelect) = mcSelect Then
                        ' selected item. None used, be creative?
                    Else
                        ' normal state, not selected, not hovered over.
                    End If
                End If
                
                ' draw the caption
                If (CID.itemState And mcHover) = mcHover Then ' hover
                    ' different style for disabled items when hovering vs normal
                    If (CID.itemState And mcDisabled) = mcDisabled Then ' item is disabled
                        lvCW.Graphics.TextDrawEx dllCaption, CID.hdc, .Left + imgOffset.X, .Top, .Right, .Bottom, fxFlat, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, gMenu_Font, Nothing, vb3DDKShadow
                    Else
                        lvCW.Graphics.TextDrawEx dllCaption, CID.hdc, .Left + imgOffset.X, .Top, .Right, .Bottom, fxRaised, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, gMenu_Font, Nothing, vbBlue, vbWhite
                    End If
                
                Else
                    If (CID.itemState And mcDisabled) = mcDisabled Then
                        ' disabled item, draw same style,color for selected & standard items
                        lvCW.Graphics.TextDrawEx dllCaption, CID.hdc, .Left + imgOffset.X, .Top, .Right, .Bottom, fxSunken, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, gMenu_Font, Nothing, vb3DDKShadow, vbCyan
                    Else
                        If (CID.itemData And mcSelect) = mcSelect Then ' selected
                            lvCW.Graphics.TextDrawEx dllCaption, CID.hdc, .Left + imgOffset.X, .Top, .Right, .Bottom, fxSunken, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, gMenu_Font, Nothing, vbBlue, vbWhite
                        Else                                           ' normal
                            lvCW.Graphics.TextDrawEx dllCaption, CID.hdc, .Left + imgOffset.X, .Top, .Right, .Bottom, fxSunken, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, gMenu_Font, Nothing, vbBlue, vbWhite
                        End If
                    End If
                End If
                
                ' draw the border
                If (CID.itemState And mcHover) = mcHover Then ' hover
                    ' use DLL function to draw a polyline
                    lvCW.Graphics.DrawPolyLine CID.hdc, RGB(128, 128, 128), .Left + 4, .Bottom, .Left, .Bottom - 4, .Left, .Top + 4, .Left + 4, .Top, .Right - 4, .Top
                    lvCW.Graphics.DrawPolyLine CID.hdc, &HFF80FF, .Right - 4, .Top, .Right, .Top + 4, .Right, .Bottom - 4, .Right - 4, .Bottom, .Left + 3, .Bottom
                Else
                    If (CID.itemState And mcSelect) = mcSelect Then ' selected
                        ' here we are simply going to outline the octagon
                        lvCW.Graphics.DrawPolyLine CID.hdc, vb3DDKShadow, .Left + 4, .Bottom, .Left, .Bottom - 4, .Left, .Top + 4, .Left + 4, .Top, .Right - 4, .Top
                        lvCW.Graphics.DrawPolyLine CID.hdc, &HFF80FF, .Right - 4, .Top, .Right, .Top + 4, .Right, .Bottom - 4, .Right - 4, .Bottom, .Left + 3, .Bottom
                    Else           ' normal/standard
                        ' do nothing, up to you
                    End If
                End If
            
            End If
            
        End With
    End If
    
    ' now that the caption is drawn if needed, let's do the icons
    
    If CID.itemData > 0 Then    ' then we assigned an icon to this menu item
        
        ' calculate the offset for the example being displayed
        If CID.itemOD = odMenuItem_Complete Then
            imgOffset.X = 3 ' left edge of image
            imgOffset.Y = (CID.rcItem.Bottom - CID.rcItem.Top - 15) / 2
        Else
            ' example is only drawing images, not caption.
            ' Provided rectangle must be used as is
            imgOffset.X = 0
            imgOffset.Y = 0
        End If
            
        ' we will use the DLL's Graphics object to help draw the icons
        ' cause its there and pretty easy to use, even with an ImageList source
        
        ' when menu item is disabled or window is inactive,
        ' draw then icon with different shading vs the blah gray
        If (CID.itemState And mcDisabled) = mcDisabled Or wParam = 0 Then
        
            lvCW.Graphics.DrawImageIcon imgLst.ListImages(CID.itemData).Picture.Handle, _
                CID.hdc, CID.rcItem.Left + imgOffset.X, CID.rcItem.Top + imgOffset.Y, , , True, &H808000, vbCyan
                
        Else ' draw the normal icon
        
            lvCW.Graphics.DrawImageIcon imgLst.ListImages(CID.itemData).Picture.Handle, _
                CID.hdc, CID.rcItem.Left + imgOffset.X, CID.rcItem.Top + imgOffset.Y
        End If
    
    End If

End Sub

Private Sub DoDrawBarBackground(wParam As Long, lParam As Long)
'////////////////////// wMsg = omDrawMBarBkg \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw the menubar background as a result of passing
' the odMenuBarBkg flag to the OwnerDrawn function of the DLL

' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a BkgAction_LV structure
' dllCaption is not used

'BkgAction_LV values
'   hDC. a DC to draw in. Do not remove the font (without replacing it) or the
'           bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. the menubar boundaries. This is NOT clipped
'       ** This could be a zero size rectangle if window is dragged very small
'       ** Background will always be erased with the current Frame.Backcolor value
'           or filled with the current gradient colros/image unless the
'           BackStyle property is set to bfTRansparent
'   rcExtra. the rectangle where the menubar items will be drawn. This is NOT clipped
'           and you should ensure your drawing stays within its bounds


'////////////////////// wMsg = omDrawTBarBkg \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a titlebar background as a result of passing
' the odTitlebarBkg flag to the OwnerDrawn function of the DLL

' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a BkgAction_LV structure
' dllCaption is not used

'BkgAction_LV values
'   hDC. a DC with correct titlebar font already selected. Do not remove the font
'        or the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. the titlebar boundaries. This is NOT clipped
'       ** This could be a zero size rectangle if window is dragged very small
'       ** The caption rectangle portion of the titlebar will be drawn with the
'           current solid/gradient colors/image unless you set the
'           BackStyle property to bfTransparent
'   rcExtra. the rectangle where the menubar items will be drawn. This is NOT clipped
'           and you should ensure your drawing stays within its bounds


Dim BKG As BkgAction_LV

' example of filling the menubar or titlebar with an image
' I simply used the same drawing for both 'cause I'm lazy :)

' The DLL will fill a background with image, solids or gradients.
' But if you want to do something special, do it yourself. This example
' will tile an image around the space used for the titlebar caption or
' menubar menu items, depending on the example you clicked.

    
    CopyMemory BKG, ByVal lParam, Len(BKG) ' get drawing info into our UDT
    
'    If wParam = 1 Then  ' active window
    
        With BKG.rcItem
            
            ' use dll graphics class to tile, outline & gradient fill the area.
            ' Note that the menu/title bar was told to have a transparent background
            ' so it wouldn't draw over our background.
            
            lvCW.Graphics.ImageFillEx BKG.hdc, .Left, .Top, .Right, .Bottom, imgLst.ListImages(13).Picture, bsTiled, , (wParam = 0)
            lvCW.Graphics.DrawRect BKG.hdc, vbBlack, .Left + 6, .Top + 6, .Right - 6, .Bottom - 6
            lvCW.Graphics.GradientFillEx BKG.hdc, .Left + 7, .Top + 7, .Right - 7, .Bottom - 7, False, &H8080FF, vbWhite, (wParam = 0)
        
        End With
        
'    End If

End Sub

Private Sub DoDrawPostNonClient(wParam As Long, lParam As Long)
'////////////////////// wMsg = omDrawPostNC \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw anything you want on the nonclient and
' is received as a result of passing the odPostNCDrawing flag
' to the OwnerDrawn function of the DLL

' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a BkgAction_LV structure
' dllCaption is not used

'BkgAction_LV values
'   hDC. a DC to draw in. Do not remove the font (without replacing it) or
'           the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. the window boundaries where top/left is always 0,0
'   rcExtra. this is not a standard rectangle. The elements are pointers to other rectangles
'       rcExtra.Left is a pointer to the client rectangle
'       rcExtra.Top is a pointer to the complete titlebar rectangle
'       rcExtra.Right is a pointer to the complete menubar rectangle
'       rcExtra.Bottom is a pointer to the non-client inset
'       ** Using that information, you can calculate available space to add any
'           additional graphics or Tracking Rectangles (explained later)
'       ** to copy a pointer (rcExtra.Left) to your rectangle variable:
'               CopyMemory myRectVariable, ByVal rcExtra.Left, &H10



' This example adds a the PSC hyperlink and I told the DLL that I wanted
' to track mouse actions over it, so I'll provide a rectangle for the DLL
' to track

Dim testFont As StdFont
Dim BKG As BkgAction_LV
Dim trackRect As RECT
Dim trackCaption As String
    
    If wParam = 0 Then ' inactive window
    
        ' for simplicity, I won't show the tracking rectangle when the
        ' form is inactive. Remove it if we already had it set
        lvCW.Frame.RemoveTrackingRect 62
        
        ' otherwise, I could draw it and either request or not request
        ' mouse action messages for that rectangle
        Exit Sub
        
    End If
    
    Set testFont = New StdFont  ' font for our tracking recangle
    testFont = Me.Font
    testFont.Bold = True
    testFont.Underline = True
    testFont.Size = 10
    testFont.Name = "Tahoma"
    
    ' get the drawing information
    CopyMemory BKG, ByVal lParam, Len(BKG)
    
    ' the BKG.rcExtra contains 4 pointers to 4 different rectangles to help you
    ' position any text/graphics anywhere on the non-client area. For this example,
    ' I only want the the space immediately below the enlarged inset we created.
    
    ' The inset rectangle is at pointer rcExtra.Bottom and
    ' the client rectangle is at pointer rcExtra.Left
    ' I could use either one to calculate what I want, I'll use the client rectangle
    
    ' Note that the following will completely erase the other pointers, so should
    ' you need them for calculations, copy them into a separate temporary RECT
    ' structure  vs doing it this way...
    CopyMemory BKG.rcExtra, ByVal BKG.rcExtra.Left, &H10
    
    ' I want to center the caption under the client rectangle
    trackCaption = "Planet Source Code"
    
    ' calculate max width available for our caption (client rectangle width)
    trackRect.Right = BKG.rcExtra.Right - BKG.rcExtra.Left + 1
    
    ' send caption to function to measure & truncate the caption (add the ... if needed)
    ' The function also will fill in the right & bottom elements of the rectangle
    lvCW.Graphics.TextMeasureEx BKG.hdc, trackCaption, gDC_Font, trackRect.Right, trackRect.Bottom, DT_MODIFIABLE Or DT_WORD_ELLIPSIS, testFont
    
    ' now position the rectangle & draw the text
    With BKG.rcExtra
        OffsetRect trackRect, (.Right - .Left - trackRect.Right + 2) \ 2 + .Left, .Bottom + 5
    End With
    
    ' now that we have our tracking rectangle, use it to draw and pass to the DLL
    With trackRect
    
        ' one important note here. If I wanted to do something more fancy when I
        ' got my mouse notifications, I would want to cache the background to my
        ' own bitmap so I could replace it before I draw the caption. When I do
        ' draw the caption again as a result of mouse clicks, I am only changing
        ' the font color, so I don't need to cache the background. The DLL does
        ' not manage the background of the nonclient for tracking rectangles
    
        ' draw our caption
        lvCW.Graphics.TextDrawEx trackCaption, BKG.hdc, .Left, .Top, .Right, .Bottom, fxFlat, DT_SINGLELINE, gDC_Font, testFont, vbYellow
        
        ' tell DLL we want to get mouse notification for this tracking rectangle
        ' the 62 below is the ID we will reference the rectangle by.
        ' You can add as many as you need.
        
        lvCW.Frame.AddTrackingRect 62, .Left, .Top, .Right, .Bottom, Me, "PSC Top Ten List"
    
    End With
    

End Sub

Private Sub DoDrawBorders(wParam As Long, lParam As Long)
'////////////////////// wMsg = omDrawPreNC \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a window border as a result of passing
' the odFrameBorders flag to the OwnerDrawn function of the DLL.
' Here you have the opportunity to completely draw the entire nonclient

' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a BkgAction_LV structure
' dllCaption is not used

'BkgAction_LV values
'   hDC. a DC with to draw in. Do not remove the font or the bitmap from the DC.
'           Doing so will prevent your window from drawing
'   rcItem. the window boundaries where top/left is always 0,0
'   rcExtra. the nonclient area remaining after the border width/height are subtracted

' ** Note: If you are drawing the background of the nonclient area, be aware that
'   if the menubar and titlebar Backstyle properties are not set to bfTransparent,
'   then those objects will be filled in as normal.
    
Dim BKG As BkgAction_LV
Dim clipRgn As Long
Dim imgRect As RECT

    ' draw custom borders, background, etc
    CopyMemory BKG, ByVal lParam, Len(BKG)
    
    With BKG
        
        ' we'll leave the 3 pixel border created by the DLL, although we could
        ' draw over it if we wanted to.
        ' Fill in the space between the outer borders with a gradient
        
        ' If wParam = 1 Then window is painted as active
        ' Note last parameter in next line. We will use same colors for active/inactive
        ' but will grayscale the gradient colors when the window is inactive
        lvCW.Graphics.GradientFillEx .hdc, .rcItem.Left + 3, .rcItem.Top + 3, .rcItem.Right - 3, .rcItem.Bottom - 3, False, &H24CACE, &H80FFFF, (wParam = 0)
    
    End With
    
    ' add the custom borders. If I really wanted to get fancy, I would have separate
    ' borders for the inactive and active windows
    
    ' I will create a clipping region here. Because if the window is shrunk very
    ' small there is a possibility that our graphics will over draw the borders.
    ' Alternatively, I could calculate the minimum size the window should be
    ' in order to display the borders without clipping them, then call the
    ' .Frame.MinMax.MinDragSize function to set a minimum window size.
    
    ' clip to prevent overdrawing the 3 pixel border
    clipRgn = CreateRectRgn(2, 2, BKG.rcItem.Right - 3, BKG.rcItem.Bottom - 3)
    SelectClipRgn BKG.hdc, clipRgn
    DeleteObject clipRgn
    
    ' Our border images are in Image Controls. I prefer to use VB's .Render method
    ' of the stdPicture object. This is hazardous if not used correctly, and will
    ' crash. If you are uncomfortable with .Render, you could always select your
    ' image into a bitmap or picBox and BitBlt or .PaintPicture as you prefer
    
    ' don't modify the following one bit!
    With imgBdr(0).Picture
        .Render BKG.hdc + 0, 2, 2, ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), _
            0, .Height, .Width, -.Height, ByVal 0&
        ' if window is inactive we will grayscale the transparent image area after it
        ' is painted. This works well because we already grayscaled the gradient
        ' background & re-grayscaling over grayscaled colors doesn't change the color
        If wParam = 0 Then lvCW.Graphics.GrayScale_DC BKG.hdc, 2, 2, 2 + ScaleX(.Width, vbHimetric, vbPixels), 2 + ScaleY(.Height, vbHimetric, vbPixels)
    End With
    
    With imgBdr(1).Picture
        .Render BKG.hdc + 0, BKG.rcItem.Right - 3 - ScaleX(.Width, vbHimetric, vbPixels), _
            BKG.rcItem.Bottom - ScaleY(.Height, vbHimetric, vbPixels) - 3, _
            ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), _
            0, .Height, .Width, -.Height, ByVal 0&
            
        ' if window is inactive we will grayscale the transparent image area after it
        ' is painted. This works well because we already grayscaled the gradient
        ' background & re-grayscaling over grayscaled colors doesn't change the color
        
        ' Note: imgRect simply used to help make the function call more readable
        imgRect.Left = BKG.rcItem.Right - 3 - ScaleX(.Width, vbHimetric, vbPixels)
        imgRect.Top = BKG.rcItem.Bottom - ScaleY(.Height, vbHimetric, vbPixels) - 3
        imgRect.Right = imgRect.Left + ScaleX(.Width, vbHimetric, vbPixels)
        imgRect.Bottom = imgRect.Top + ScaleY(.Height, vbHimetric, vbPixels)
        If wParam = 0 Then lvCW.Graphics.GrayScale_DC BKG.hdc, imgRect.Left, imgRect.Top, imgRect.Right, imgRect.Bottom
    End With

End Sub

Private Sub DoDrawSystemButton(wParam As Long, lParam As Long)
'////////////////////// wMsg = omDrawSysBtn \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a system button as a result of passing
' the odSysButtons flag to the OwnerDrawn function of the DLL

' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a CustomItemDraw_LV structure
' dllCaption is not used

'CustomItemDraw_LV values
'   itemID. will be one of the following
'           SC_CLOSE. Draw the close button
'           SC_MINIMIZE. Draw the minimize button
'           SC_MAXIMIZE. Draw the maximize button
'           SC_RESTORE. Draw the Restore Down button (window is maximized)
'   itemPos. not used
'   itemData. not used
'   itemState. will contain one or more of the following values using OR
'               mcStandard - item is not selected, not hovered over
'               mcHover - mouse is currently over the menu item
'               mcSelect - menu item is currently selected
'               mcDisabled - menu item is disabled
'   itemOD. not used
'   hDC. a DC to daw in. Do not remove the font (without replacing it) or
'           the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. The rectangle to draw. This rectangle is NOT clipped and you should
'           ensure your drawing stays within its bounds.
'       ** Note. The DLL provides a blank button to draw on having a 2 pixel border.
'           You can draw within that button (rcItem dimensions) or draw over it.

Dim CID As CustomItemDraw_LV
Dim btnCaption As String

    ' no real example here. It's same as the custom button examples (better example)
    ' The only difference is that the .itemID parameter will be one of the following
    '   SC_CLOSE, SC_MINIMIZE, SC_MAXIMIZE or SC_RESTORE
    
    ' the following simple example I used to ensure the routine worked
    CopyMemory CID, ByVal lParam, Len(CID)
    Select Case CID.itemID
    Case SC_CLOSE: btnCaption = "x"
    Case SC_RESTORE: btnCaption = "r"
    Case SC_MAXIMIZE: btnCaption = "M"
    Case SC_MINIMIZE: btnCaption = "m"
    End Select
    With CID.rcItem
        lvCW.Graphics.TextDrawEx btnCaption, CID.hdc, .Left, .Top, .Right, .Bottom, fxFlat, DT_CENTER Or DT_VCENTER, gMenu_Font, Nothing, vbBlue
    End With

End Sub


Private Sub DoDrawCustomButton(wParam As Long, lParam As Long)
'////////////////////// wMsg = omDrawUserBtn \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a custom button. This message is
' automatically sent when you added your button thru the
' TitleBar.Buttons.AddButton function

' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a CustomItemDraw_LV structure
' dllCaption is the tool tip you provided when you added the button

'CustomItemDraw_LV values
'   itemID. The button ID you provided when you added the button
'   itemPos. The zero-bound position of the button (0 thru 4)
'   itemData. not used
'   itemState. will contain one or more of the following values using OR
'               mcStandard - item is not selected, not hovered over
'               mcHover - mouse is currently over the menu item
'               mcSelect - menu item is currently selected
'               mcDisabled - menu item is disabled
'   itemOD. not used
'   hDC. a DC to draw in. Do not remove the font (without replacing it) or
'       the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. The rectangle to draw. This rectangle is NOT clipped and you should
'           ensure your drawing stays within its bounds.
'       ** Note. The DLL provides a blank button to draw on having a 2 pixel border.
'           You can draw within that button (rcItem dimensions) or draw over it.

Dim CID As CustomItemDraw_LV
Dim btnCaption As String
Dim imgOffset As POINTAPI
Dim testFont As StdFont
Dim btnSize As POINTAPI

    ' FYI: You will not get a hover or down-state message if you disabled your button
    '      However, you will get a draw every time the window needs to be painted
    
    CopyMemory CID, ByVal lParam, Len(CID)  ' get the drawing information
    
    With CID
        If (.itemState And mcSelect) = mcSelect Then
            ' when button is in down/click state, offset graphics/text so
            ' the image appears to move down and right with a click movement
            imgOffset.X = 1
            imgOffset.Y = 1
        End If
        
        Select Case .itemID ' the button ids we assigned when we created the button
        Case 22, 33
            ' Example of using text on the buttons
            Set testFont = New StdFont
            testFont.Name = "Tahoma"
            testFont.Size = 9
            testFont.Bold = True
            
            If .itemID = 33 Then btnCaption = "!" Else btnCaption = "@"
            
            ' measure the text so we can center it on the button
            lvCW.Graphics.TextMeasureEx .hdc, btnCaption, gDC_Font, btnSize.X, btnSize.Y, DT_SINGLELINE, testFont
            
            ' now center the what will be the text rectangle
            .rcItem.Left = (.rcItem.Right - .rcItem.Left - btnSize.X) \ 2 + .rcItem.Left
            .rcItem.Right = .rcItem.Left + btnSize.X
            .rcItem.Top = (.rcItem.Bottom - .rcItem.Top - btnSize.Y) \ 2 + .rcItem.Top
            .rcItem.Bottom = .rcItem.Top + btnSize.Y
            
            ' offset the text rectangle if the button is in the click state
            OffsetRect .rcItem, imgOffset.X, imgOffset.Y
            ' depending on enabled or disabled state, draw the text
            If (.itemState And mcDisabled) = mcDisabled Then
                lvCW.Graphics.TextDrawEx btnCaption, .hdc, .rcItem.Left, .rcItem.Top, .rcItem.Right, .rcItem.Bottom, fxSunken, DT_SINGLELINE, gDC_Font, testFont, vbGrayText, vbWhite
            Else
                If (lvCW.Frame.Active.GetInsetBackColor = vbCyan And ((.itemState And mcHover) = mcHover)) And wParam = 1 Then
                    lvCW.Graphics.TextDrawEx btnCaption, .hdc, .rcItem.Left, .rcItem.Top, .rcItem.Right, .rcItem.Bottom, fxSunken, DT_SINGLELINE, gDC_Font, testFont, vbYellow, vbBlue
                Else
                    lvCW.Graphics.TextDrawEx btnCaption, .hdc, .rcItem.Left, .rcItem.Top, .rcItem.Right, .rcItem.Bottom, fxSunken, DT_SINGLELINE, gDC_Font, testFont, vbBlue, vbCyan
                End If
            End If
            
        Case 11
            ' example shows overriding the button with a standard caption done with API
            ' there are only a few standard caption: close, minimize, restore, etc
            lvCW.Graphics.DrawButtonShape .hdc, .rcItem.Left, .rcItem.Top, .rcItem.Right, .rcItem.Bottom, dfcCaptionButton, DFCS_CAPTIONHELP Or (imgOffset.X * DFCS_PUSHED)
        
        Case 0
            ' examples shows using an icon for the button image. Some notes here:
            ' 1. The button will always be 16x16 with a 1 pixel border, leaving 14x14 for the image
            ' 2. To make the yellow diamond & arrow, I simply used the VB-provided, free ImageEdit
            '    creating a 16x16, 16color icon.
            '   The trick is to not use the outer 2 pixels of the icon leaving it
            '   transparent and in effect, giving you a 14x14 icon
            
            ' offset the image if button in down position
            OffsetRect .rcItem, imgOffset.X, imgOffset.Y
            
            ' the custom buttons also receive a hover message also, so you could
            ' change the image based on the mouse over. When the mouse leaves the
            ' button you will get another message where .itemState=0 do draw normal
            
            ' One last Note. You don't have to accept the gray button the DLL provides.
            ' You can color that 16x16 space anyway you want: solids, gradients, etc
            If (.itemState And mcHover) = mcHover Then
                lvCW.Graphics.DrawImageIcon imgLst.ListImages(12).Picture.Handle, .hdc, .rcItem.Left - 1, .rcItem.Top
            Else
                lvCW.Graphics.DrawImageIcon imgLst.ListImages(11).Picture.Handle, .hdc, .rcItem.Left - 1, .rcItem.Top
            End If
            
        End Select
        
    End With

End Sub

Private Sub DoDrawTracker(wParam As Long, lParam As Long)
'////////////////////// wMsg = omDrawTrackRect \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a item you are tracking that exists within
' the nonclient area. This message is a result of calling
' the Frame.AddTrackingRect function. Processing this message is optional

' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a CustomItemDraw_LV structure
' dllCaption is not used

'CustomItemDraw_LV values
'   itemID. The tracking rectangle you provided when you called AddTrackingRect
'   itemPos. not used
'   itemData. not used
'   itemState. will contain one or more of the following values using OR
'               mcStandard - item is not selected, not hovered over
'               mcHover - mouse is currently over the menu item
'               mcSelect - menu item is currently selected
'   itemOD. You must provide a non-zero value if you drew the updated rectangle
'           into the DC and want it posted to the visible window
'   hDC. a DC with operating system's standard font selected.
'           Do not remove the font without replacing it. Memory leaks would occur
'   rcItem. The rectangle to draw in
'       ** The DLL does not graphically track or replace the space occupied by
'           this rectangle. You should have cached the background when processing
'           the omDrawPostNC message and copying the background to your own bitmap
'       ** Tip. By drawing text only or redrawing text, using the same text styles
'           and font, no need at all to cache the background image.
'
'   Final Notes: You are required to update the itemOD value if you redrew the
'           rectangle and want it posted to the visible window


Dim CID As CustomItemDraw_LV
Dim testFont As StdFont
Dim trackCaption As String

    
    CopyMemory CID, ByVal lParam, Len(CID)
    
    Select Case CID.itemID  ' the id of the tracking rectangle we assigned
    Case 62
        Set testFont = New StdFont
        testFont = Me.Font
        testFont.Bold = True
        testFont.Underline = True
        testFont.Size = 10
        trackCaption = "Planet Source Code"
    Case Else
        ' other tracking rectangles
        ' I used text as the graphics in this example, you could just as
        ' easily use images to. But with transparent images, you should
        ' cache the background before the image is drawn. See the
        ' DoDrawPostNonClient routine above for a bit more information
    End Select
    
    With CID.rcItem
        ' depending on the mouse action, change colors of the text
        
        If (CID.itemState And mcStandard) = mcStandard Then
            ' no mouse over, no click, nothing
            lvCW.Graphics.TextDrawEx trackCaption, CID.hdc, .Left, .Top, .Right, .Bottom, fxFlat, DT_SINGLELINE Or DT_WORD_ELLIPSIS, gDC_Font, testFont, vbYellow
            
        ElseIf (CID.itemState And mcSelect) = mcSelect Then
            ' mouse is down on the rectangle
            lvCW.Graphics.TextDrawEx trackCaption, CID.hdc, .Left, .Top, .Right, .Bottom, fxFlat, DT_SINGLELINE Or DT_WORD_ELLIPSIS, gDC_Font, testFont, vbRed
            
        Else
            ' mouse is over the rectangle, but no button is down
            lvCW.Graphics.TextDrawEx trackCaption, CID.hdc, .Left, .Top, .Right, .Bottom, fxFlat, DT_SINGLELINE Or DT_WORD_ELLIPSIS, gDC_Font, testFont, vbMagenta
        End If
    End With
    
    ' we updated the graphics, let DLL know so it can be updated to our window
    CID.itemOD = 1  ' any non-zero number will suffice
    CopyMemory ByVal lParam, CID, Len(CID)

End Sub

Private Sub DoTrackerClick(wParam As Long, lParam As Long)
'////////////////////// wMsg = omTrackClick \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To inform user that a tracking rectangle was clicked
' Note. Do not interpret the itemState value of mcSelect in the
'   omDrawTrackRect message as a button click. When a user clicks down
'   on a tracker, a omDrawTrackRect message is sent, but if user does not
'   release the mouse and drags mouse off the tracker, it is not a click
'   event even though another omDrawTrackRect message will be sent with
'   a itemState value of mcStandard to draw the omDrawTrackRect in the up state

' wParam is the tracking rectangle ID you assigned
' lParam is 0 if window is inactive, or 1 if active
'   ** This value may also contain mcRightButton if the click was a right button
' dllCaption is not used
    
    Select Case wParam ' id of our tracking rectangle(s)
    Case 62 ' url to psc
        If (lParam And mcRightButton) = 0 Then   ' left clicked
            apiShellExecute Me.hwnd, "OPEN", _
                "http://www.planetsourcecode.com/vb/contest/ContestAndLeaderBoard.asp?lngWid=1", _
                "", "", 1
                
        Else    ' right clicked the tracking rectangle
            MsgBox "When left clicked, this will jump to PSC's Top Ten", vbInformation + vbOKOnly
            
        End If
        
    End Select

End Sub

Private Sub Form_Resize()
' resize the text box to full window size
' Amazingly, this is the only thing that flickers when reszing the window :)
If Not Me.WindowState = vbMinimized Then
    On Error Resume Next
    Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End If
End Sub


Private Function HiWord(LongIn As Long) As Integer
  Call CopyMemory(HiWord, ByVal VarPtr(LongIn) + 2, 2)
End Function
Private Function LoWord(LongIn As Long) As Integer
  Call CopyMemory(LoWord, ByVal VarPtr(LongIn), 2)
End Function

