VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iControlsCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub WindowMessage(ByVal hWnd As Long, ByVal wMsg As Long, _
        ByRef wParam As Long, ByRef lParam As Long, _
        ByRef bBlockMessage As Boolean, ByRef BlockValue As Long)

End Sub

' This implementation must be used when subclassing a child control or a MDI form
'
'///////////////////////////////////////////////////////////////////////////////////////
' Using the bBlockMessage & BlockValue, you can override the return value expected
' by the WindowProcedure and usage of the class function ForwardMessage, you can either
' modify or completely discard the message being passed.
' To override the message, set bBlockMessage=True & change BlockValue as needed
'   Note: Most window messages require BlockValue to be zero, but not all
'         You should always look up the blocked return value for the message @ MSDN
'
'   Your WindowMessage event is similar to a standard WindowProcedure routine
'   the DLL 's public function ForwardMessage is identical usage of the CallWndProc API
'
' When you implement this class, ALL messages will be forwarded to your window
' for the MDI form or control that is subclassed.
' You can override these messages; except the WM_Destroy message.
'
' WM_Destroy can never be overridden. It is not even sent to the CustomWindow class
'   To prevent your form MDI from closing, you should process the form's QueryUnload event

