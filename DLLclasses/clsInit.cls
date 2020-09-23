VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Separate class to act as a Lib_Initialize & Lib_Terminate event
' when the DLL is first started/referenced. And IF app closes gracefully
' then this class will terminate & basically clean up for the entire DLL

' Easiest way to really know if your dll terminated cleanly is to
' set up some system that fires a terminate event. Therefore,
' this class is used mainly for that purpose.



Private Sub Class_Terminate()

 TerminateClients

End Sub




