VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "LaVolpe Sample Menu"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3405
   FillColor       =   &H0080FFFF&
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   15
      TabIndex        =   0
      Top             =   480
      Width           =   3360
   End
   Begin VB.Frame frameAlign 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   15
      TabIndex        =   2
      Top             =   3585
      Width           =   3375
      Begin VB.CheckBox chkRotated 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Use Rotated Icons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   840
         Left            =   2250
         TabIndex        =   11
         ToolTipText     =   "Will rotate system icon in same direction as the titlebar caption"
         Top             =   975
         Width           =   960
      End
      Begin VB.CheckBox chkAlwaysActive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Allow Test Form to Go Inactive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   840
         Left            =   2250
         TabIndex        =   10
         ToolTipText     =   "Toggle between the test form being told to appear active while the thread has the focus"
         Top             =   135
         Width           =   1020
      End
      Begin VB.CheckBox optAlignMB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Menu bar Bottom Aligned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   1545
         Width           =   2130
      End
      Begin VB.CheckBox optAlignMB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Menu bar Top Aligned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   1275
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.OptionButton optAlignTB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Title bar Bottom Aligned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   6
         Top             =   960
         Width           =   2040
      End
      Begin VB.OptionButton optAlignTB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Title bar Right Aligned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   5
         Top             =   705
         Width           =   1935
      End
      Begin VB.OptionButton optAlignTB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Title bar Left Aligned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   4
         Top             =   450
         Width           =   1905
      End
      Begin VB.OptionButton optAlignTB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Title bar Top Aligned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   195
         Value           =   -1  'True
         Width           =   1950
      End
   End
   Begin VB.PictureBox Picture1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   150
      ScaleHeight     =   1605
      ScaleWidth      =   2865
      TabIndex        =   9
      Top             =   1065
      Visible         =   0   'False
      Width           =   2925
      Begin VB.Image imgInset 
         Height          =   720
         Left            =   465
         Picture         =   "frmMain.frx":0000
         Top             =   540
         Width           =   720
      End
      Begin VB.Image imgTBar 
         Height          =   3840
         Left            =   1035
         Picture         =   "frmMain.frx":0504
         Top             =   435
         Width           =   3840
      End
      Begin VB.Image imgMenubar 
         Height          =   1110
         Left            =   30
         Picture         =   "frmMain.frx":8566
         Top             =   60
         Visible         =   0   'False
         Width           =   1665
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Click to Show Example"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15
      TabIndex        =   1
      Top             =   75
      Width           =   3300
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Classy As CustomWindow
Private bInitialized As Boolean

Private Sub chkAlwaysActive_Click()
    
    If chkAlwaysActive.Value = 1 Then
        ' form is active only when it has the focus
        frmTest.lvcw.SetKeepActive True, False, False
    Else
        ' keep window active always while in thread, not always
        frmTest.lvcw.SetKeepActive True, True, False
    End If
    ' changing this setting will automatically force a refresh if needed
End Sub

Private Sub chkRotated_Click()
    ' option to display rotated system icon
    
    If chkRotated = 0 Then
        frmTest.lvcw.Titlebar.Buttons.RotateSystemIcon = Rotate0
    Else
        frmTest.lvcw.Titlebar.Buttons.RotateSystemIcon = RotateAuto
    End If
    frmTest.lvcw.RefreshWindow
End Sub

Private Sub Form_Load()
Debug.Print "main form hwnd "; Me.hwnd

    ' I ran 150 and the GDI object count obviously went up 'cause of VB usage
    ' but when all were closed (closing the main form closes all the others),
    ' the GDI count went back to where it started... therefore, no leaks yet :)
    ' When I tried 200, started having memory issues -- again VB, cause this
    ' DLL only uses a few resources per project, not per form
    
    ' un rem the following section if you want to see that
    ' this DLL can handle several windows without any problems AFAIK
    
'    Dim x(0 To 4) As frmTest
'    Dim y As Integer
'        For y = 0 To UBound(x)
'            Set x(y) = New frmTest
'            Load x(y)
'            x(y).lvCW.BeginCustomize x(y).hwnd
'            x(y).lvCW.Titlebar.Alignment = Int(Rnd * 4) + 10 ' barAlignLeft
'            x(y).Show
'            Set x(y) = Nothing
'        Next
    
    LoadListBox
    
    ' Don't know what font your system is using, but for this example
    ' project, I will use Tahoma for the titlebar
    Me.Font.Name = "Tahoma"
    Me.Font.Size = 9.5
    Me.Font.Bold = True
    
    Label1.Caption = Label1.Caption & vbNewLine & "(MB) Menubar (TB) Titlebar (BTN) Buttons"
    
    ' create new instance
    Set Classy = New CustomWindow
    ' tell it to appear active as long as we are in its thread
    Classy.SetKeepActive True, True, False
    
    ' let's jazz up our form a little
    ' no menubar, so we won't worry about that
    
    With Classy.Titlebar
        ' set up the titlebar alignment, back & fore colors/style
        .NoRedraw = True
        .Alignment = barAlignRight
        .Active.SetBackColors bfGradientNS, &HC0&, &H8080FF
        .Active.SetTextColors fxSunken, vbWhite, &H800000
        .Font = Me.Font
        .Buttons.Active.SetForeColors vbButtonText, vbGrayText, 0, True, vbRed
        ' hide the system icon 'cause I think it detracts from the titlebar
        ' This will not prevent user from right clicking on the titlebar without
        ' one of the other optional flags set too.
        .Buttons.SystemIconMenu = eSysIconHidden

    End With
    
    ' since all settings were done before we subclassed, no need to
    ' call the RefreshWindow method. As soon as we begin customization
    ' the window will refresh automatically...
    Classy.BeginCustomize hwnd
    
    ' the minmax class won't take effect unless we are subclassed
    With Classy.Frame
        .Active.SetBorderColors vbButtonFace, &H8080FF, &HC0&, &H80&
        ' prevent the window from being resized
        .MinMax.MaxDragSize True, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY
        .MinMax.MinDragSize True, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY
    End With
    
    ' prevent maximizing. Although I prevented sizing by setting the MinMax.DragSizes
    ' the Size system menu item will still appear and the cursors will change when moved
    ' over the window borders. By disabling SC_Size below the DLL will prevent the
    ' cursors from changing and will also disable the Size system menu item
    With Classy.Titlebar.Buttons
        .EnableSysMenuItem SC_MAXIMIZE, sysDisable
        .EnableSysMenuItem SC_SIZE, sysDisable
        .HideDisabledButtons = True
    End With
    '^^ the systembuttons won't take effect unless we are subclassed
    
    ' place window where I want it
    Move (Screen.Width - Me.Width) / 2 + Me.Width / 2, (Screen.Height - Me.Height) / 2
        
    
        
    ' now load our sample/test form
    Show
    DoEvents
    Load frmTest
    With frmTest

        .lvcw.SetKeepActive True, True, False
        .lvcw.BeginCustomize .hwnd
        bInitialized = True

        ' position it & show it
        .Move Me.Left - .Width - 15, (Screen.Height - .Height) / 2
        .Show
        DoEvents

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim f As Form
    For Each f In Forms
        If f.hwnd <> Me.hwnd Then Unload f
    Next
End Sub

Private Sub optAlignMB_Click(Index As Integer)
' option to change menubar alignment

If optAlignMB(Index) Then
    
    optAlignMB(Abs(Index - 1)) = Abs(optAlignMB(Index).Value - 1)
    With frmTest.lvcw.MenuBar
        .NoRedraw = True        ' don't refresh, still have some changes to do
        If Index = 0 Then .Alignment = barAlignTop Else .Alignment = barAlignBottom
    End With
    
    frmTest.lvcw.RefreshWindow ' show the changes

Else

    optAlignMB(Abs(Index - 1)) = 1

End If
    
End Sub

Private Sub optAlignTB_Click(Index As Integer)
' option to change titlebar alignment

If optAlignTB(Index) Then
    
    With frmTest.lvcw.Titlebar
        .NoRedraw = True   ' don't refresh until I tell it to
        
        Select Case Index
        Case 0: .Alignment = barAlignTop
        Case 1: .Alignment = barAlignLeft
        Case 2: .Alignment = barAlignRight
        Case 3: .Alignment = barAlignBottom
        End Select
        
        ' as noted in the ShowAdditionalInfo routine, vertical gradients do not
        ' automatically flip the gradient when the titlebar changes orientation.
        ' This is on purpose with only exception being use of default gradients.
        
        ' Therefore, we should flip the gradient if our titlebar orientation
        ' changes, unless we don't want to
        
        If List1.Text <> "TB. Non-Default Gradients" Then ' don't fip for this example
            ' not flipping the gradient may not be the visual effect you want
            If frmTest.lvcw.AmIActive Then
                If .Alignment < barAlignBottom Then ' vertical titlebar
                    If .Active.BackStyle = bfGradientEW Then .Active.BackStyle = bfGradientNS
                Else
                    If .Active.BackStyle = bfGradientNS Then .Active.BackStyle = bfGradientEW
                End If
            End If
            ' could test for inactive window state & flip the gradient too
            ' but all the samples use default color schemes for inactive titlebars
            ' and the DLL automatically flips gradients when default scheme is used
        End If
        
    End With
    
    frmTest.lvcw.RefreshWindow ' show the changes

End If
End Sub

Private Sub LoadListBox()
    List1.AddItem "MB. Gradient Horizontal"
    List1.AddItem "MB. Gradient Vertical"
    List1.AddItem "MB. Picture Background"
    List1.AddItem "MB. Custom Menu Item Look"
    List1.AddItem "MB. Hide Disabled Menu Items"
    List1.AddItem "MB. Toggle Menu Bar Visibility"
    List1.AddItem "MB. Owner Drawn (Image Responsibility)"
    List1.AddItem "MB. Owner Drawn (Full Responsibility)"
    List1.AddItem "MB. Custom Menu Bar Borders"
    List1.AddItem "TB. Non-Default Gradients"
    List1.AddItem "TB. Solid Back Colors"
    List1.AddItem "TB. Picture Background"
    List1.AddItem "TB. Custom Title Bar Borders"
    List1.AddItem "TB. Wrappable Caption"
    List1.AddItem "TB. Toggle Title Bar Visibility"
    List1.AddItem "TB. Add Custom Buttons"
    List1.AddItem "BTN. Disable Maximize & Minimize"
    List1.AddItem "BTN. Disable Move & Size"
    List1.AddItem "BTN. Disable Close"
    List1.AddItem "BTN. Hide Disabled Buttons"
    List1.AddItem "Frame. Custom Colors {Old Blue}"
    List1.AddItem "Frame. Custom Inset"
    List1.AddItem "Frame. OwnerDrawn Borders"
    List1.AddItem "Misc. Toggle Show On Taskbar"
    List1.AddItem "Misc. Subclass Child Window"
    List1.AddItem "Misc. Tray Icons"
    List1.AddItem "Misc. Minimize To System Tray"
    List1.AddItem "Misc. Clone via Saved Settings"
    List1.AddItem "Misc. Other Stuff"
    
End Sub

Private Sub List1_Click()

Dim tbFont As StdFont
Dim mbFont As StdFont
Dim sTitle As String

If List1.Text = "Misc. Clone via Saved Settings" Then
    CloneSample
    Exit Sub
End If

If List1.Text <> "Misc. Toggle Show On Taskbar" And _
    List1.Text <> "Misc. Other Stuff" Then Call ResetToBasics

With frmTest.lvcw
    
    Select Case List1.Text
    
    Case "Misc. Other Stuff"
        ' nothing to do here
    
    Case "TB. Add Custom Buttons"
        .Titlebar.Buttons.RemoveButton 0, True
        .Titlebar.Buttons.AddButton frmTest, 0, "Using Icon for the button image"
        .Titlebar.Buttons.AddButton frmTest, 11, "Using standard API caption: DFC_Help"
        .Titlebar.Buttons.AddButton frmTest, 22, "Using Text for the button image"
        .Titlebar.Buttons.AddButton frmTest, 33, "Using Text for the button image #2"
    
    Case "MB. Gradient Horizontal"
        ' add these gradients & change the menu item colors so they can be seen easily
        .MenuBar.Active.SetBackColors bfGradientEW, RGB(232, 232, 232), vb3DShadow
        .MenuBar.Active.SetEnabledTextColors fxFlat, mcSelect Or mcStandard Or mcHover, vbBlue
'        .MenuBar.Active.SetFrame bxFlat, vbYellow
        
    Case "MB. Gradient Vertical"
        ' add these gradients & change the menu item colors so they can be seen easily
        With .MenuBar.Active
            .SetBackColors bfGradientNS, vb3DShadow, RGB(232, 232, 232)
            .SetEnabledTextColors fxFlat, mcSelect Or mcStandard Or mcHover, vbBlue
        End With
    
    Case "MB. Picture Background"
        ' add the image & change the menu item colors so they can be seen easily
        With .MenuBar.Active
            ' spend some time to ensure your text colors look well enough
            .SetEnabledTextColors fxRaised, mcSelect Or mcStandard, RGB(213, 213, 0), vb3DDKShadow
            .SetEnabledTextColors fxFlat, mcHover, vbWhite
            .SetDisabledTextColors fxSunken, mcHover Or mcSelect Or mcStandard, vbBlack, RGB(162, 162, 0)
            .SetImageBackground imgMenubar, bsSmartStretch
        End With
        .MenuBar.Inactive.SetImageBackground imgMenubar, bsSmartStretch, True
        
    Case "MB. Custom Menu Item Look"
        With .MenuBar.Active
            .SetMenuSelectionStyle mbFilledCustom, bxHover, vbBlue, vbWhite
            .SetMenuSelectionStyle mbSunkenCustom, bxSelect, vbBlue, vbWhite
            .SetEnabledTextColors fxSunken, mcHover, vbYellow, vbBlack
            .SetDisabledTextColors fxFlat, mcHover, vbGrayText
            .SetDisabledTextColors fxSunken, mcSelect Or mcStandard, &H808000, vbWhite
            .SetEnabledTextColors fxRaised, mcStandard Or mcSelect, vbBlue, vbWhite
            .SetBackColors bfGradientEW, vbCyan, vb3DShadow
            .SetFrame bx3D, vbCyan, &H800000
        End With
        .MenuBar.Inactive.SetBackColors bfGradientEW Or bfGrayScaled, vbCyan, vbButtonFace
        .MenuBar.Inactive.SetFrame bx3D, .Graphics.GrayScale_Color(vbCyan), .Graphics.GrayScale_Color(&H800000)
        
    Case "MB. Hide Disabled Menu Items"
        .MenuBar.HideDisabledItems = True
        
    Case "MB. Toggle Menu Bar Visibility"
        .MenuBar.ShowMenuBar = Not .MenuBar.ShowMenuBar
        
    Case "MB. Owner Drawn (Image Responsibility)", "MB. Owner Drawn (Full Responsibility)"
        If List1.Text = "MB. Owner Drawn (Full Responsibility)" Then
            .OwnerDrawn frmTest, odMenuItem_Complete Or .isOwnerDrawn(0) And Not odMenuItem_ImageOnly
        Else
            .OwnerDrawn frmTest, odMenuItem_ImageOnly Or .isOwnerDrawn(0) And Not odMenuItem_Complete
        End If
        With .MenuBar.Active
            ' note the following 2 won't be effective when full responsibility is deferred
            .SetMenuSelectionStyle mbFilledCustom, bxHover, RGB(192, 192, 255), vbBlue
            .SetEnabledTextColors fxRaised, mcHover, &H8000000, vbWhite
        End With
        
    Case "MB. Custom Menu Bar Borders"
        .MenuBar.BorderWidth = 8
        .MenuBar.BorderHeight = 7
        .OwnerDrawn frmTest, odMenuBarBkg
        
    Case "TB. Toggle Title Bar Visibility"
        .Titlebar.ShowTitlebar = Not .Titlebar.ShowTitlebar
        
    Case "TB. Non-Default Gradients"
        With .Titlebar.Active
            .SetBackColors False, &H808000, vbCyan
            .SetTextColors fxSunken, &H57E3E0, vbBlue
            .SetFrame bx3D, vbGreen, vbBlue
        End With
        With .MenuBar.Active
            .SetMenuSelectionStyle mbNoBorders, bxHover, 0, 0
            .SetEnabledTextColors fxFlat, mcHover, vbBlue
        End With
        
    Case "TB. Solid Back Colors"
        With .Titlebar.Active
            .SetBackColors bfSolid, &H57E3E0
            .SetTextColors fxSunken, &H800000, vbCyan
            .SetFrame bxFlat, vbBlue
        End With
        .Titlebar.Inactive.SetBackColors bfSolid Or bfGrayScaled, &H57E3E0
        With .MenuBar.Active
            .SetMenuSelectionStyle mbNoBorders, bxHover, 0, 0
            .SetEnabledTextColors fxRaised, mcHover, &H800000, vbWhite
        End With
        
        Set mbFont = New StdFont
        With mbFont
            .Name = "Comic Sans"
            .Size = 9
        End With
        .MenuBar.Font = mbFont
        
    Case "TB. Wrappable Caption"
        sTitle = InputBox("Enter a longer caption if desired, then resize the test window to see it wrap", _
            "Wrappable Caption", frmTest.Caption)
        If Len(sTitle) Then .Titlebar.Caption = sTitle
        .Titlebar.Caption = sTitle
        .Titlebar.WrapCaption = True
        
    Case "BTN. Disable Maximize & Minimize"
        If .Titlebar.Buttons.EnableSysMenuItem(SC_MAXIMIZE, sysDisable) = False Then
            MsgBox "Sorry, that button could not be disabled -- try again later"
        Else
            .Titlebar.Buttons.EnableSysMenuItem SC_MINIMIZE, sysDisable
        End If
        
    Case "BTN. Disable Move & Size"
        If .Titlebar.Buttons.EnableSysMenuItem(SC_MOVE, sysDisable) = False Then
            MsgBox "Sorry, that button could not be disabled --- try again later"
        Else
            .Titlebar.Buttons.EnableSysMenuItem SC_SIZE, sysDisable
        End If
        
    Case "BTN. Disable Close"
        If .Titlebar.Buttons.EnableSysMenuItem(SC_CLOSE, sysDisable) = False Then
            MsgBox "Sorry, that button could not be disabled --- try again later"
        End If
        
    Case "BTN. Hide Disabled Buttons"
        With .Titlebar.Buttons
            .HideDisabledButtons = True
            .EnableSysMenuItem SC_MAXIMIZE, sysDisable
            .EnableSysMenuItem SC_MINIMIZE, sysDisable
        End With
        
    Case "TB. Picture Background"
        ' add the image & change the menu item colors so they can be seen easily
        With .Titlebar.Active
            ' spend some time to ensure your text colors look well enough
            .SetTextColors fxSunken, vbCyan, vbBlack
            .SetImageBackground imgTBar, bsSmartStretch
        End With
        .Titlebar.Inactive.SetImageBackground imgTBar, bsSmartStretch, True
        Set tbFont = New StdFont
        With tbFont
            .Name = "Tahoma"
            .Size = 11
            .Bold = True
            .Italic = True
        End With
        .Titlebar.Font = tbFont
        With .MenuBar.Active
            .SetDisabledTextColors fxSunken, mcSelect Or mcStandard, &H416996, vbWhite
            .SetEnabledTextColors fxRaised, mcStandard Or mcSelect, vbBlue, vbWhite
            .SetEnabledTextColors fxSunken, mcHover, vbYellow, vbBlack
            .SetBackColors bfGradientEW, vbCyan, vb3DShadow
        End With
    
    Case "TB. Custom Title Bar Borders"
        .Titlebar.BorderHeight = 8
        .Titlebar.BorderWidth = 7
        .Titlebar.Active.SetTextColors fxSunken, vbRed, 0
        .Titlebar.Inactive.SetTextColors fxFlat, vb3DShadow
        .OwnerDrawn frmTest, odTitlebarBkg
        
    Case "Misc. Toggle Show On Taskbar"
        .ShowOnTaskbar = Not .ShowOnTaskbar
    
    Case "Frame. Custom Colors {Old Blue}"
        .Frame.Active.SetBorderColors &HFFC0C0, vbCyan, vbBlue, &H400040
        .Frame.Active.SetInsetBackColor vbCyan
        With .MenuBar.Active
            .SetDisabledTextColors fxSunken, mcSelect Or mcStandard, &H416996, vbWhite
            .SetEnabledTextColors fxRaised, mcStandard Or mcSelect, vbBlue, vbWhite
            .SetEnabledTextColors fxSunken, mcHover, vbYellow, vbBlack
            .SetDisabledTextColors fxFlat, mcHover Or mcSelect Or mcStandard, &H808000
            .SetBackColors bfGradientEW, vbCyan, vb3DShadow
        End With
        With .Titlebar
            .Active.SetTextColors fxSunken, vbCyan, &H800000
            .Buttons.Active.SetBorderColors &HFFC0C0, vbCyan, vb3DShadow, vbBlue
            .Buttons.Active.SetForeColors vbBlue, vbGrayText, vbWhite, False, vbYellow
        End With
        .MenuBar.Inactive.SetBackColors bfGradientEW Or bfGrayScaled, vbCyan, vbButtonFace
        .MenuBar.Inactive.SetFrame bx3D, .Graphics.GrayScale_Color(vbCyan), .Graphics.GrayScale_Color(&H800000)
        .OwnerDrawn frmTest, odMenuItem_ImageOnly
        .Titlebar.Buttons.AddButton frmTest, 0, "Using Icon for the button image"
        .Titlebar.Buttons.AddButton frmTest, 11, "Using standard API caption: DFC_Help"
        .Titlebar.Buttons.AddButton frmTest, 22, "Using Text for the button image"
        .Titlebar.Buttons.AddButton frmTest, 33, "Using Text for the button image"
        
        
    Case "Frame. Custom Inset"
        With .Frame
            .SetInsetOffset 10, 10, 10, 30
            .Active.SetInsetImage imgInset, bsTiled
            .Inactive.SetInsetImage imgInset, bsTiled, True
        End With
        .OwnerDrawn frmTest, odPostNCDrawing
        
    Case "Frame. OwnerDrawn Borders"
        ' the DLL will automatically add a 3 pixel sizing border using current
        ' border color settings when a border size is > 4 pixels, otherwise draws
        ' a border = to the border sizes you set. You can overdraw this default
        ' border if needed. If not needed, the DLL saved you some drawing
        
        With .Frame
            .BorderWidth = 28         ' make room for our 26pixel wide image
            .BorderHeight = 28        ' & 2pixel border
            .Active.SetInsetBackColor 0, True, True   ' don't allow the nc inset
            .Inactive.SetInsetBackColor 0, True, True ' to be painted
        End With
        With .Titlebar
            .Active.BackStyle = bfTransparent  ' no fill option for both
            .Inactive.BackStyle = bfTransparent ' active and inactive states
            
            .Active.SetTextColors fxRaised, &H800000, vbWhite   ' set titlebar colors
            .Inactive.SetTextColors fxFlat, RGB(64, 64, 64)     ' active & inactive
            ' customize the system buttons
            .Buttons.Active.SetBorderColors &H80FFFF, vbWhite, &H24CACE, &H4979B
            .Buttons.Active.SetForeColors 0, vb3DShadow, 0, True, vbBlue
        End With
        With .MenuBar
            .Inactive.BackStyle = bfTransparent ' no fill option for both
            .Active.BackStyle = bfTransparent   ' active and inactive states
        End With
        ' tell DLL we want to paint the borders & non-client background
        .OwnerDrawn frmTest, odFrameBorders
        ' allow user to move window on extended border
        .Frame.AllowExtendedMove = True
        
        
    Case "Misc. Subclass Child Window"
        .Subclass_OtherWindow frmTest, frmTest.Text1.hwnd
        
    Case "Misc. Tray Icons"
        ' This function will not activate the system tray until you subclass
        ' your form and include an icon. Once you subclass your form, call ActivateTray again to
        ' activate it. Since we are already subclassed, we only call ActivateTray once
        .Frame.SystemTray(frmTest.hwnd, frmTest).ActivateTray frmTest.Icon, frmTest.Caption
        If .Frame.SystemTray(frmTest.hwnd, frmTest).isBalloonCapable Then
            .Frame.SystemTray(frmTest.hwnd, frmTest).ShowBalloon "You can have multiple icons for the same window", "FYI. Did you know?", icInfo
        Else
            MsgBox "You can have multiple icons for the same window." & vbCrLf & "One has been placed in your system tray", vbInformation + vbOKOnly, "FYI. Did you know?"
        End If
        
    Case "Misc. Minimize To System Tray"
        If .MinimizeToSysTray Then
            .MinimizeToSysTray = False
        Else
            .MinimizeToSysTray = True
            .Frame.SystemTray(frmTest.hwnd, frmTest).ActivateTray frmTest.Icon, frmTest.Caption
            .Frame.SystemTray(frmTest.hwnd, frmTest).ShowBalloon "Minimize the form and then restore by clicking on this icon.", "Minimize to/from system tray", icInfo
        End If
    End Select

End With

    If bInitialized Then
        frmTest.lvcw.RefreshWindow ' show the changes
    Else
        bInitialized = True
        ' start the customization (all defaults as no custom settings done yet)
        frmTest.lvcw.BeginCustomize frmTest.hwnd
    End If
    
    Call ShowAdditionalInfo
    
    On Error Resume Next    ' in case the test form is minimized
    If chkAlwaysActive = 1 Then frmTest.SetFocus
    
End Sub

Private Sub ShowAdditionalInfo()

Dim sTip As String

    Select Case List1.Text
        Case "MB. Gradient Horizontal", "MB. Gradient Vertical"
            sTip = "The gradient background can also be applied to an inactive window" & vbCrLf & _
                vbCrLf & "Obviously font colors on the menu bar would be dependent upon the colors used " & _
                "in the gradients. Therefore, you have full control over how the colors are set." & _
                vbCrLf & vbCrLf & "Use the SetEnabledTextColors & SetDisabledTextColors of the " & _
                "MenuBar.Active and the MenuBar.Inactive classes."
    
        Case "MB. Picture Background"
            sTip = "Any picture object can be used for the background of the menubar."
    
        Case "MB. Custom Menu Item Look"
            sTip = "This example shows some of the possible options. Hover over the menu items & click on " & _
                "one or two of them to see some custom options in use." & vbCrLf & vbCrLf & _
                "Each menu item state (normal, hover, & selected) can have different colors and " & _
                "different font styles (sunken, engraved, flat or raised)" & vbCrLf & vbCrLf & _
                "Use the SetEnabledTextColors & SetDisabledTextColors of the " & _
                "MenuBar.Active and the MenuBar.Inactive classes. " & _
                "To set different hover box styles or colors, use the SetMenuSelectionStyle class."
                
        Case "MB. Custom Menu Bar Borders"
            sTip = "You can add space around your menubar in order to draw the menubar in ways not " & _
                "supported by the DLL. For example, mixing gradients, tiling, etc." & vbCrLf & vbCrLf & _
                "When you add the space, you should also request to owner draw the menubar background, otherwise, " & _
                "the DLL will fill the extra space with the default non-client backcolor."
                
        Case "MB. Hide Disabled Menu Items"
            sTip = "The menu item titled ""Remote"" & ""Edit"" are a disabled items. Notice that you cannot see them now." & _
            vbCrLf & vbCrLf & "If you didn't notice they were part of the sample menu, click on another sample " & _
            "to see where they were."
            
        Case "MB. Toggle Menu Bar Visibility"
            sTip = "This neat little option simply toggles the visibility of the entire menu." & vbCrLf & vbCrLf & _
            "The system menu is still available." & vbCrLf & vbCrLf & "Also any Ctrl+keys assigned to the menu items as long as " & _
            "the menu items are visible & enabled in your application." & vbCrLf & vbCrLf & "        " & _
            "Press Ctrl+B to get a message box from one of the menu items."
            
        Case "TB. Non-Default Gradients"
            sTip = "Like the menu bar, the title bar can have any color gradients and in vertical or horizontal " & _
                "direction." & vbCrLf & vbCrLf & "To make the menu bar a little less plain, the hover/mouse over " & _
                "effect is border-less and changes the fore color of the menu item. Move the mouse over the menu items." & _
                vbCrLf & vbCrLf & "Also added an optional two-color border to the title bar just for some extra pizazz." & _
                vbCrLf & vbCrLf & "Pay attention to vertical title bars. Horizontal gradients on a horizontal title bar looks " & _
                "fine, but horizontal gradient on a vertical title bar may look poor. The gradient does not automatically " & _
                "reverse direction when the titlebar orientation changes. This feature is only automatic if the title bar " & _
                "gradient colors are system default colors. Toggle vertical title bar to see difference: You can use " & _
                "Titlebar.Active/Inactive.GradientNorthSouth to toggle the gradient direction"
        
        Case "TB. Picture Background"
            sTip = "Any picture object can be used for the background of the title bar."
                    
        Case "TB. Solid Back Colors"
            sTip = "Like the menu bar, the title bar can be any color." & vbCrLf & vbCrLf & "To make the menu bar a " & _
                "little less plain, the hover/mouse over effect is border-less and changes the fore color of the " & _
                "menu item. Move the mouse over the menu items." & vbCrLf & vbCrLf & "Also added an optional " & _
                "single-color border to the title bar just for fun"
                
        Case "BTN. Disable Maximize & Minimize"
            sTip = "This maximize, minimize, move, size & close buttons can be disabled to the point that the user cannot " & _
                "click on them and cannot use SendMessage to trigger the event. " & vbCrLf & vbCrLf & "This does not prevent you from using " & _
                "code like Me.WindowState = vbMaximized; because the maximize & minimize styles are not removed from " & _
                "the window; rather the click event on the window is controlled/restricted"
            
        Case "BTN. Disable Move & Size"
            sTip = "These system menu items can also be disabled and you can still use their associated commands in your application. " & _
                "For example, Me.Move still works as usual." & vbCrLf & vbCrLf & _
                "Notice too that the code will prevent the Sizing Cursors from displaying while the cursor is over " & _
                "any of the borders; try it -- move cursor over the borders and it shouldn't change" & vbCrLf & vbCrLf & _
                "Of course if you don't disable the maximize and/or minimize buttons too; kinda defeats the purpose of not sizing, huh?"
                
        Case "BTN. Disable Close"
            sTip = vbCrLf & "WARNING...... If you disable the Close button, ensure you give your user a way to close your form. " & vbCrLf & vbCrLf & _
                "The window can still be closed via code. For example you can use Unload Me, but a user cannot use the system menu " & _
                "or Alt+F4 to close the window. This application is not designed to trap & prevent Task Manager or Ctrl+Alt+Del combinations."
    
        Case "BTN. Hide Disabled Buttons"
            sTip = "Disabled buttons can be hidden so they don't detract from your titlebar." & vbCrLf & vbCrLf & "The example here has both the " & _
                "Minimize & Maximize buttons disabled and also not shown"
                
        Case "Frame. Custom Colors {Old Blue}"
            sTip = "This example uses custom colors for the borders generating a pleasant blue appearance." & vbCrLf & vbCrLf & _
                "Additionally, the buttons are shaded blue too, with a yellow mouse over effect." & vbCrLf & vbCrLf & _
                "Notice that the one button that uses the windows API caption" & _
                " is not shaded. Using those captions means that the API will fill the background of the button with the system " & _
                "default colors... The reason I had to draw the min/max/close & restore buttons manually so that we could override the colors."
                
        Case "Frame. Custom Inset"
            sTip = "Space can be added between the client rectangle and its adjacent non-client regions. This example added 10 pixels " & _
            "to top, left & right edges & 30 pixels to the bottom edge of the client rectangle. Keep in mind that your client rectangle is reduced to supply that space." & vbCrLf & vbCrLf & _
                "This space can be used to display text and can be painted as you wish. To create the space, you simply make " & _
                "a call to the Frame.SetInsetOffset method. Solid fill color can be set via Frame.SetNCBackColor." & vbCrLf & vbCrLf & _
                "Custom painting that additional space requires you to request an owner-drawn Inset which will notify you when to paint, the window active/inactive status, " & _
                "and also forwards a hit test when the cursor is over that new inset." & vbCrLf & vbCrLf & _
                "The hit test could be used by you to determine if a mouse is over a custom icon, text, whatever. This " & _
                "is an ideal location to add custom buttons/icons, hyperlinks, etc."
                
        Case "MB. Owner Drawn (Image Responsibility)"
            sTip = "Icons can add real flash to the menubar. But too many icons may look a bit childish IMO." & vbCrLf & vbCrLf & _
                "This example uses icons that are left aligned to the menu caption. The menu background is gradient filled. " & _
                "Icons can also be right aligned. The menubar can have mixed and matched settings, some icons can be right aligned, " & _
                "while others can be left aligned, and others may not have icons at all. In fact, you can take over 100% responsibility " & _
                "for drawing the menubar by calling a single method. P.S. Mouse over the icons :)" & vbCrLf & vbCrLf & _
                "And all this can be done with about a dozen lines of code, using an Image List control to store the images. " & _
                "The same thing can easily be done using an array of simple Image controls or memory handles... Your preference."
                
        Case "MB. Owner Drawn (Full Responsibility)"
            sTip = "This example expects you to completely measure and draw the menu items, their borders, and any background you desire." & vbCrLf & vbCrLf & _
                "The example shown here displays the images, some custom font styles/colors, and a unique hover border that can only be " & _
                "accomplished by taking full ownership of the menu item drawing. The DLL only draws rectangular shapes." & vbCrLf & vbCrLf & _
                "The inactive window drawing also needs to be done by you. You will see that I simply used default settings for " & _
                "most of the actions, except adding the octagon border to the inactive window's menubar." & vbCrLf & vbCrLf & _
                "Don't expect drawing everything will be easy. However, the code in the test form shows heavy use of some common " & _
                "drawing routines that are already provided to you from this DLL"
                
        Case "TB. Wrappable Caption"
            sTip = "The wrapple caption will wrap a caption on to no more than two lines when there is no space on the title bar to " & _
                "display the complete caption ." & vbCrLf & vbCrLf & "One stipulation is that the first line of the caption, must " & _
                "be able to display at least one word of the caption.  If not, the truncated caption will be displayed on a single line." & _
                vbCrLf & vbCrLf & "A neat option, drag the window to a smaller size to see the results."
        
        Case "TB. Toggle Title Bar Visibility"
            sTip = "The title bar can be toggled on and off. This has a couple of nice benefits... You can have a title bar that shows on the " & _
                "windows taskbar, but not have the title bar show on your window. Additionally, the form is completely sizable and can be " & _
                "minimized/maximized from the windows taskbar (right click and see)." & vbCrLf & vbCrLf & _
                "A small bonus is that if a menu is being displayed, the user can drag the window by clicking on the menu bar in an " & _
                "area that doesn't have a menu item in it. This is generally the right side of the menu bar."
    
        Case "TB. Custom Title Bar Borders"
            sTip = "You can add space around your titlebar in order to draw the titlebar in ways not " & _
                "supported by the DLL. For example, mixing gradients, tiling, etc." & vbCrLf & vbCrLf & _
                "When you add the space, you should also request to owner draw the titlebar background, otherwise, " & _
                "the DLL will fill the entire titlebar with the current color and/or image settings."
    
        Case "Frame. OwnerDrawn Borders"
            sTip = "This neat option can really enhance your form to make it completely different than anything else out there!" & vbCrLf & vbCrLf & _
                "The example requests to owner-draw the borders. When this is requested, the DLL fills the background with a default color and " & _
                "adds a 3 pixel border using current border color settings. Then when the sample gets the paint request, it gradient fills from " & _
                "border to border and places the images at the top & bottom corners." & vbCrLf & vbCrLf & _
                "Note. In order for the titlebar and menubar to appear to use the gradient color scheme, the sample simply told the DLL " & _
                "to fill the menubar and title bar as transaprent. Hint. Using transparent fills, you may also want them to be transparent " & _
                "when the window is inactive too -- depending on your inactive window color scheme."
    
        Case "TB. Add Custom Buttons"
            sTip = "You can add up to 5 custom buttons to your titlebar. The above example shows four." & vbCrLf & vbCrLf & _
                "The size of the buttons are defined by the system, normally 16x16. The DLL will provide a blank button in " & _
                "the appropriate state (down/up) for you to draw over. Additionally the buttons will use any custom color " & _
                "settings you chose." & vbCrLf & vbCrLf & "You are not required to draw inside the button, the 16x16 space " & _
                "occupied by the button is yours to do with what you want." & vbCrLf & vbCrLf & _
                "The buttons will receive a down, up and/or hover message so you can draw appropriately. " & _
                "A separate click event is sent when the button is clicked. Disabled buttons will only receive a message " & _
                "to re-draw the button as needed, no hover or click events will be forwarded."
    
        Case "Misc. Toggle Show On Taskbar"
            sTip = "The windows task bar will remove or add this form's caption and icon alternately as you click the option " & _
                "again and again. " & vbCrLf & vbCrLf & "Nothing special here, except windows as issues when trying to do this " & _
                "when the window is minimized. In this case the window will be restored, task bar toggled, and then re-minimized."
        
        Case "Misc. Subclass Child Window"
            sTip = "Just a short simple example showing that the DLL can subclass any of your child controls and forward you the " & _
                "messages to do with what you want. This form will be unSubclassed when you click on the text or close the window." & vbCrLf & vbCrLf & _
                "Although this DLL is not compatible with MDI forms or their MDI children, you can still subclass those windows and " & _
                "receive the messages."
    
        Case "Misc. Other Stuff"
            sTip = "Other nice properties are included or replicated:" & vbCrLf & vbCrLf & _
            "If you don't read over the short QuickLook.RTF file, you won't have a clue to all of the possible " & _
            "settings you can use to completely customize a standard window." & vbCrLf & vbCrLf & _
            "1. AlwaysOnTop sets the WS_EX_TOPMOST window style" & vbCrLf & _
            "2. AppIcon sets the icon associated with your app for the Alt + Tab window" & vbCrLf & _
            "3. MaximizeFullScreen property maximizes window over window taskbars" & vbCrLf & vbCrLf & _
            "And so many other properties and settings"
    
        Case "Misc. Tray Icons"
            sTip = "A class is provided that allows you to display and interact with the system tray via " & _
                "an icon you provide. You must implement the iSysTrayCallback class to use the system tray." & vbCrLf & vbCrLf & _
                "With the call back you can tell when some clicks on your icon, when someone clicks on a balloon tip, or " & _
                "whether the balloon tip timed out or closed some other way. Additionally, the class is designed with " & _
                "an Explorer crash, restoration routine. If Explorer crashes, your icons come back automatically."
        
        Case "Misc. Minimize To System Tray"
            sTip = "This handy property will show minimizing animation to the system tray area vs just minimizing " & _
                "to the window's task bar. " & vbCrLf & vbCrLf & "When restoring, it can animate from the system tray " & _
                "to Restore or Maximized if you use the optional method ShowWindowFromTrayIcon." & vbCrLf & vbCrLf & _
                "Try it. Minimize this window and the animated titlebar will terminate in the system tray, regardless " & _
                "where this window exists on the task bar. Then click on the Fox icon in the system tray and watch the " & _
                "animation start from the system tray to the restored location."
    
    End Select

    frmTest.Text1.Text = sTip
    frmTest.Text1.SelStart = 0
    
End Sub


Private Sub ResetToBasics()

' Note this would much easier to simply set the main class to nothing and then
' create a new instance. But that would make the sample form flicker between
' updates as titlebars, menubars, etc can be shifted from unique positions
' to the standard positions.... Therefore, I'll simply reset all the options
' that could have been applied in any previous example:

With frmTest.lvcw

    .NoRedraw = True           ' prevent updating in all classes for now
    
    ' note that the following can be called individually if needed
'    .MenuBar.NoRedraw = True    ' updating any changed properties
'    .Titlebar.NoRedraw = True   ' prevent refreshing/redrawing until done
    
    With .Titlebar.Buttons
        .EnableSysMenuItem SC_CLOSE, sysEnable
        .EnableSysMenuItem SC_MAXIMIZE, sysEnable
        .EnableSysMenuItem SC_MINIMIZE, sysEnable
        .EnableSysMenuItem SC_MOVE, sysEnable
        .EnableSysMenuItem SC_SIZE, sysEnable
        .RemoveButton 0, True
    End With
    .Frame.RemoveTrackingRect 0, True
    
    
    Select Case List1.Text
        ' don't reset window if these items are selected.
        ' This way you can see how the current window looks with the following set
        Case "MB. Hide Disabled Menu Items"
            .Titlebar.ShowTitlebar = True
            .MenuBar.ShowMenuBar = True
            .Titlebar.Buttons.HideDisabledButtons = False
            
        Case "MB. Toggle Menu Bar Visibility"
            .Titlebar.ShowTitlebar = True
            .Titlebar.Buttons.HideDisabledButtons = False
            .MenuBar.HideDisabledItems = False
            
        Case "TB. Toggle Title Bar Visibility"
            .MenuBar.ShowMenuBar = True
            .MenuBar.HideDisabledItems = False
            .Titlebar.Buttons.HideDisabledButtons = False
        
        Case "TB. Wrappable Caption", _
            "TB. Add Custom Buttons", _
            "MB. Owner Drawn (Image Responsibility)", _
            "MB. Owner Drawn (Full Responsibility)"
            
                .Titlebar.ShowTitlebar = True
                .MenuBar.ShowMenuBar = True
                .MenuBar.HideDisabledItems = False
                .Titlebar.Buttons.HideDisabledButtons = False
    
        Case Else
            ' reset to defaults so we don't carry over anything from previous example(s)
            
            ' note: each level has its own Reset function. So if you only wanted to
            ' reset the menubar font colors for active windows, you could use the following
            '   .MenuBar.Active.ResetToSystemDefaults(xxxx)
            ' or if you wanted the Menubar in its entirety, you could use
            '   .MenuBar.ResetToSystemDefaults(xxxx)
            ' the following is called from the root level: lvCW
            
            ' The .ResetToSystemDefaults resets colors and fonts, it does not reset
            '   other properties. The best way to completely remove all properties
            '   is to simply set the main class to nothing and re-instantiate it
            
            .ResetToSystemDefaults rstAll ' reset entire window colors/fonts to defaults
            '^^ resets the titlebar, menubar, borders, & system buttons
            
            .OwnerDrawn frmTest, 0     ' stop any ownerdrawn NC regions
            .Titlebar.Buttons.RemoveButton 0, True
            ' reset other properties/options that may have been set in other examples
            .MenuBar.HideDisabledItems = False
            .MenuBar.ShowMenuBar = True
            
            ' we reset the font to the system setting, but I want to use this form's font
            With .Titlebar
                .Font = Me.Font
                .WrapCaption = False
                .ShowTitlebar = True
                .Buttons.HideDisabledButtons = False
            End With
            
    End Select
    
   
End With

End Sub

Private Sub CloneSample()

    If frmTest.lvcw Is Nothing Then Exit Sub
    
    Dim I As Long, myPropBag As PropertyBag
    Dim X As Long, Y As Long
    
    If Forms.Count > 2 Then
        ' save the left/top coords if the form is already opened
        ' so we can display it again in the same location
        X = frmClone.Left
        Y = frmClone.Top
        Unload frmClone
    End If
    
    ' I'll use a property bag to save the ImageList items and custom borders
    ' Then I simply tell the DLL to include the property bag with the export
    Set myPropBag = New PropertyBag
    
    With frmTest
        
        ' add all the image list items to the property bag
        For I = 1 To .imgLst.ListImages.Count
            myPropBag.WriteProperty "Img" & I, .imgLst.ListImages(I).Picture, Nothing
        Next
        ' add the custom borders too
        myPropBag.WriteProperty "Brdr0", .imgBdr(0).Picture, Nothing
        myPropBag.WriteProperty "Brdr1", .imgBdr(1).Picture, Nothing
        ' add any other information you may need or want in the export
        myPropBag.WriteProperty "NrImages", .imgLst.ListImages.Count
        
        ' now call the DLL to create the Export for you
        .lvcw.ExportCustomSettings App.Path & "\_TestClone.pbg", exAll, myPropBag
        ' frmClone uses this export to import settings and assign menu item images
        
    End With
    
    ' if the form was shown & then reshown, show in same place
    If X > 0 Or Y > 0 Then
        Load frmClone
        frmClone.Move X, Y
    End If
    
    ' make the form visible
    frmClone.Show
    
End Sub
