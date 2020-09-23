VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' almost all of this class from http://vbnet.mvps.org/index.html?code/subclass/shellnotifystructinfo.htm
' added additional routines & modified a couple so they can be used generically
' and also to allow multiple icons per hWnd

' This class will allow you to modify system tray icons for your application

' Any hWnd can have a tray icon. All icons you assign to the tray icons are your
' responsibility for destroying when appropriate. Only top level windows should
' be installing tray icons. This is because should Explorer crash & then
' restart & new taskbar created, the message we look for to restore the icons
' on the new taskbar is only sent to top level windows.

' This class can be setup in your routines at any time and can be deactivated at
' any time, destroyed, or activated any time after the hWnd has begun subclassing.


Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, _
   ByVal lpSubBlock As String, lpBuffer As Any, nVerSize As Long) As Long

Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128      'shell 5+  <> 64 chars max for shell <5
   dwState As Long            'shell 5+  <> not used in this class
   dwStateMask As Long        'shell 5+
   szInfo As String * 256     'shell 5+
   uTimeoutAndVersion As Long 'shell 5+
   szInfoTitle As String * 64 'shell 5+
   dwInfoFlags As Long        'shell 5+
   guidItem(15) As Byte       ' GUID  shell 6+  <> reserved by Windows
End Type

' tray update commands
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
      
' tray update masks
Private Const NIF_MESSAGE = &H1
Private Const NIF_INFO As Long = &H10
Private Const NIF_ICON = &H2
Private Const NIF_STATE As Long = &H8
Private Const NIF_TIP = &H4
Private Const NIS_HIDDEN = &H1

'balloon tip notification messages <> not used here; for ref only
'Private Const NIN_BALLOONSHOW = (WM_USER + 2)
'Private Const NIN_BALLOONHIDE = (WM_USER + 3)
'Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
'Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

' tray version settings
Private Const NOTIFYICON_VERSION As Long = &H3
Private Const NIM_SETVERSION As Long = &H4
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long

Public Enum BalloonIcons
    icNone = &H0        'NIIF_NONE As Long = &H0
    icInfo = &H1        'NIIF_INFO As Long = &H1
    icWarning = &H2     'NIIF_WARNING As Long = &H2
    icError = &H3       'NIIF_ERROR As Long = &H3
    icTrayIcon = &H4    'NIIF_USER As Long = &H4 ???
    icNoSound = &H10    'NIIF_NOSOUND = &H10
    icNoChange = &H40   'class-custom value
End Enum

Private nid As NOTIFYICONDATA   ' cached UDT values
Private iconPic As StdPicture

Implements iPropertyBag

Public Property Get isActive() As Boolean
    ' returns whether or not tray icon is active.
    ' Active means it is sending your hWnd messages
    isActive = (nid.hWnd <> 0) And (nid.uCallbackMessage <> 0) And (nid.hIcon <> 0)
End Property

' property to set/get hWnd this class is assigned to.
' Called by collection class: cTrays
Friend Sub PrepTray(ByVal parentHwnd As Long, trayID As Long)
    nid.hWnd = parentHwnd
    nid.uID = trayID
End Sub

Public Sub ActivateTray(ByVal hIcon As Long, Optional newTip As String)
' This is a nice to have routine that fills in the icon & tip in one step
' vs calling those idividual properties.

' Parameters
' hIcon is a handle to an existing icon you want to display in the tray
'       -- IMPORTANT: you are responsible for destroying the icon if needed
' Tip is displayed when mouse hovers icon. 128 max unless < Shell32.v5 which is 64

    Dim lAction As Long
    
    If newTip <> "" Then Tip = newTip
    If hIcon <> 0 Then Icon = hIcon

    If NOTIFYICONDATA_SIZE = 0 Then GetTrayVersion
    If isActive = False Then
        If nid.hWnd <> 0 Then
            lAction = NIM_ADD
        Else
            lAction = NIM_MODIFY
        End If
        If nid.hIcon = 0 Then Exit Sub
    Else
        lAction = NIM_MODIFY
    End If
    
    With nid
        .cbSize = NOTIFYICONDATA_SIZE
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uTimeoutAndVersion = NOTIFYICON_VERSION
        .uCallbackMessage = WM_TrayNotify ' custom message
        ' The value of the message, should you also be using custom messages is
        ' WM_USER + &H1962
    End With
    
    Shell_NotifyIcon lAction, nid
    If lAction = NIM_ADD Then Shell_NotifyIcon NIM_SETVERSION, nid

    ' Note: The custom message above isn't really important to you.
    ' It will never be sent to your ProcessMessage routine.

End Sub

Public Property Let Icon(ByVal hIcon As Long)
' Used to change the icon displayed on the system tray
' IMPORTANT: You are responsible for destroying icons if needed

    If Not iconPic Is Nothing Then Set iconPic = Nothing
    
    nid.hIcon = hIcon
    If isActive Then
        nid.uFlags = NIF_ICON
        Shell_NotifyIcon NIM_MODIFY, nid
    End If
    
End Property
Public Property Get Icon() As Long
    Icon = nid.hIcon    ' returns handle to currently assigned icon
End Property

Public Property Let Tip(ByVal newTip As String)
    ' Used to change the tip displayed when mouse hovers over your icon
    nid.szTip = newTip & vbNullChar
    If isActive Then
        nid.uFlags = NIF_TIP
        Shell_NotifyIcon NIM_MODIFY, nid
    End If
End Property
Public Property Get Tip() As String
    Tip = Left$(nid.szTip, InStr(nid.szTip & Chr$(0), Chr$(0)) - 1)
End Property

' Read-Only properties return balloon title & message.
' Use ShowBalloon to set those values
Public Property Get BalloonTitle() As String
    BalloonTitle = Left$(nid.szInfoTitle, InStr(nid.szInfoTitle & Chr$(0), Chr$(0)) - 1)
End Property
Public Property Get BalloonMessage() As String
    BalloonMessage = Left$(nid.szInfo, InStr(nid.szInfo & Chr$(0), Chr$(0)) - 1)
End Property

Public Property Get isBalloonCapable() As Boolean
    ' property will let you know if you can use balloons on clients pc
    isBalloonCapable = (NOTIFYICONDATA_SIZE > NOTIFYICONDATA_V1_SIZE)
End Property

Public Sub ShowBalloon(Optional ByVal Message As String, _
    Optional ByVal Title As String, _
    Optional ByVal Icon As BalloonIcons = icNoChange)

' Routine shows a balloon tip only after this class has been assigned to your subclasser
' The class does not destroy your tip, title or icon settings; therefore you only need
' to pass the information that you want changed from the last time you called the tip

    If NOTIFYICONDATA_SIZE = 0 Then GetTrayVersion
    
    If NOTIFYICONDATA_SIZE > NOTIFYICONDATA_V1_SIZE Then
        ' V1 of Shell32.dll doesn't have balloon capability (pre W2K)
        
        With nid
            ' only make needed changes
            If Len(Message) Then .szInfo = Message & vbNullChar
            If Len(Title) Then .szInfoTitle = Title & vbNullChar
            If (Icon And icNoChange) = 0 Then
                If NOTIFYICONDATA_SIZE < NOTIFYICONDATA_V3_SIZE Then
                    ' filter out non-XP settings if not right version of tray
                    If (Icon Or icTrayIcon) = Icon Then Icon = Icon And Not icTrayIcon
                    If (Icon Or icNoSound) = Icon Then Icon = Icon And Not icNoSound
                End If
                .dwInfoFlags = Icon
            End If
            .uFlags = NIF_INFO
        End With
        ' if the class has been added to subclasser, send the change
        If isActive Then Shell_NotifyIcon NIM_MODIFY, nid
    End If
    
End Sub

Public Sub HideBalloon()

' function destroys a balloon if visible

' if the class has been added to subclasser, send the change
If isActive Then
    
    If NOTIFYICONDATA_SIZE > NOTIFYICONDATA_V1_SIZE Then
        Dim oldNid As NOTIFYICONDATA
        oldNid = nid
        With oldNid
             ' remove all tip info
            .szInfo = vbNullChar
            .szInfoTitle = vbNullChar
            .szTip = vbNullChar
            .uFlags = NIF_INFO Or NIF_TIP
        End With
        ' update
        Shell_NotifyIcon NIM_MODIFY, oldNid
        
        ' replace the original tip info
        nid.uFlags = NIF_TIP
        Shell_NotifyIcon NIM_MODIFY, nid
    End If
End If
End Sub

Private Sub GetTrayVersion()
' unmodified from source

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
        
         'IsShellVersion = nVerMajor >= version
         Select Case nVerMajor
         Case Is < 5
            NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE
        Case Is < 6
            NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE
        Case Else
            NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE
        End Select
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Sub

Public Sub DestroyTrayIcon(Optional bDestroyIcon As Boolean)
    ' remove the tray icon if activated
    If isActive Then
        nid.uFlags = 0
        Shell_NotifyIcon NIM_DELETE, nid
        nid.uCallbackMessage = 0
    End If
    If bDestroyIcon Then
        If nid.hIcon <> 0 Then DestroyIcon nid.hIcon
        nid.hIcon = 0
    End If
End Sub

Friend Property Get uniqueID() As Long
    uniqueID = nid.uID
End Property

Friend Sub Restore()
    ' function is called only from the DLL after a new Explorer re-starts after crash
    If isActive Then
        ' not all icons may be active, reactivate only previously active ones
        nid.uCallbackMessage = 0
        ActivateTray 0, ""
    End If
End Sub

Private Sub Class_Terminate()
    DestroyTrayIcon
End Sub

'---------------------------------------------------------------------------------------
' Procedure : clsSysTray.Read/WriteProperites
' DateTime  : 9/26/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve custom properites
' Comments  : Used by Export/ImportCustomSettings
'---------------------------------------------------------------------------------------
'
Private Sub iPropertyBag_ReadProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)
    
        Dim sProp As String
        
        sProp = propBag.ReadProperty("ST" & Index & ".ID", "hwnd")
        If sProp = "hwnd" Then
            nid.uID = lParam
        Else
            nid.uID = Val(sProp)
        End If
        
        nid.szInfo = propBag.ReadProperty("ST" & Index & ".Balloon", "") & Chr$(0)
        nid.szInfoTitle = propBag.ReadProperty("ST" & Index & ".Title", "") & Chr$(0)
        nid.szTip = propBag.ReadProperty("ST" & Index & ".Tip", "") & Chr$(0)
        nid.hWnd = lParam
        
        Set iconPic = propBag.ReadProperty("ST" & Index & ".Image", Nothing)
        nid.hIcon = iconPic.Handle
        Me.ActivateTray 0
        
End Sub

Private Sub iPropertyBag_WriteProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim trayImg As StdPicture, hHandle As Long
    
    propBag.WriteProperty "ST" & Index & ".Tip", Me.Tip, ""
    propBag.WriteProperty "ST" & Index & ".Balloon", Me.BalloonMessage, ""
    propBag.WriteProperty "ST" & Index & ".Title", Me.BalloonTitle, ""
    
    If nid.hIcon <> 0 Then
        hHandle = CopyImage(nid.hIcon, IMAGE_ICON, 0, 0, 0)
        If hHandle <> 0 Then
            Set trayImg = HandleToPicture(hHandle, False)
            propBag.WriteProperty "ST" & Index & ".Image", trayImg, Nothing
        End If
    End If
    If nid.uID = lParam Then
        propBag.WriteProperty "ST" & Index & ".ID", "hwnd"
    Else
        propBag.WriteProperty "ST" & Index & ".ID", nid.uID, 0
    End If
    
    Set trayImg = Nothing
    
End Sub
