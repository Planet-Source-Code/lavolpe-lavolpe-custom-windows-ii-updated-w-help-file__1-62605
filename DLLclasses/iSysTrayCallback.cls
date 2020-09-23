VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iSysTrayCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' separate event for tray icons help keep user projects cleaner. In addition,
' this prevents user from having to know what the custom message value is in order to process it
Public Sub ProcessTrayIcon(ByVal hWnd As Long, ByVal wTrayIconID As Long, _
     wMsg As Long, bOverriden As Boolean, lOverrideReturn As Long)

End Sub
' * users don't need the custom wMsg value anyway & wParam/lParam are less descriptive


