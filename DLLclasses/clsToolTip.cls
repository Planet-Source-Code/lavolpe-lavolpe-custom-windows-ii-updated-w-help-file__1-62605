VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Most of this class came from PSC, but it was lacking just a bit for my needs
' and had a bunch of extra stuff like balloon styles. Not wanted for titlebar buttons
    'http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=28419&lngWId=1

''Windows API Functions
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

''Windows API Constants
Private Const CW_USEDEFAULT = &H80000000

''Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_NEWTOOLRECTA As Long = (WM_USER + 6)
Private Const TTM_UPDATETIPTEXTA As Long = (WM_USER + 12)
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

''Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    lHwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

'private data
Private ttHwnd As Long
Private ttInfo As TOOLINFO

Public Function CreateToolTip(ByVal hWnd As Long, ByVal tTip As String, ByVal LocPtr As Long) As Boolean
    
    Dim lWinStyle As Long
    
    If ttHwnd = 0 Then
        On Error Resume Next
        ' when running compiled, not calling this will prevent tooltips from
        ' displaying. From what I can find on the net, calling more than once
        ' has no ill side effects
        Call InitCommonControls
        On Error GoTo 0
    Else
        ' shouldn't happen, but just in case
        DestroyWindow ttHwnd
        ttHwnd = 0
    End If
    
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    
    ttHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, _
                CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, _
                hWnd, 0&, App.hInstance, 0&)
                
    If ttHwnd <> 0 Then
        ''make our tooltip window a topmost window
        SetWindowPos ttHwnd, -1, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
                    
        ''now set our tooltip info structure
        With ttInfo
            .lSize = 40 + Len(tTip) * 2
            .lpStr = tTip
            .lFlags = TTF_SUBCLASS
            .lHwnd = hWnd
            .lId = 0
            .hInstance = App.hInstance
            If LocPtr <> 0 Then CopyMemory .lpRect, ByVal LocPtr, &H10
        
        End With
        
        ''add the tooltip structure
        SendMessage ttHwnd, TTM_ADDTOOLA, 0&, ttInfo
    
    End If
End Function

Public Property Get TipText() As String
    ' retrieve the tip text, not cached in the button class
    TipText = ttInfo.lpStr
End Property
Public Property Let TipText(newTip As String)
    ttInfo.lpStr = newTip
    ttInfo.lSize = 40 + Len(newTip) * 2
    SendMessage ttHwnd, TTM_UPDATETIPTEXTA, 0&, ttInfo
End Property

Public Property Let TipRect(rectPtr As Long)
    ' each time the toolbar is measured, any custom buttons can change
    ' position, especially since we have settings like vertical, caption wrapping.
    
    ' Plus the rect structure passed to the tooltip window must be in client coords.
    ' Since the client rectangle is the last thing to be fully calcualted during
    ' the DLL's measurement process, this routine is called after all known
    ' calculations are completed
    
    If ttHwnd <> 0 Then
        CopyMemory ttInfo.lpRect, ByVal rectPtr, &H10
        SendMessage ttHwnd, TTM_NEWTOOLRECTA, 0&, ttInfo
    End If
    
End Property


Private Sub Class_Terminate()
    ' more of a safety net. The tooltip window was created with the
    ' subclassed window as the parent; therefore, the tooltip window
    ' should be destroyed automatically when the parent gets destroyed.
    
    ' However, the button could be removed, and we would want to destroy
    ' the tip when this class gets set to Nothing.  Hmmmm? Better check
    ' before destroying anything...
        
    If ttHwnd <> 0 Then
        If IsWindow(ttHwnd) Then DestroyWindow ttHwnd
    End If
    
End Sub




