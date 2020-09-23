VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iImpProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' not for public use.
' This implementation is used between classes and the CustomWindow class to retrieve
' properties of the window vs caching key properties in each class

' these enumerations are hidden & not made viewable to VB's intellisense
Public Enum eProperties
    [_tbarAlignment] = 1        ' is titlebar vertical & left aligned
    [_sysTrayImp] = 2
    [_mnuSysMenu] = 4           ' does window have a system menu
    [_winState] = 5             ' is window maximized (zoomed)
    [_clientRect] = 6           ' pointer to the windows size & position (screen coords)
    [_winHwnd] = 7              ' the window handle (not all classes need this all the time)
    [_winRect] = 8              ' pointer to the windows size & position (screen coords)
    [_titleBarRect] = 9         ' pointer to the client size & position (screen coords)
    [_MenuFont] = 10            ' copy of the menubar font
    [_TitleFont] = 11           ' copy of the titlebar font
End Enum

Public Function GetPropValue(propID As eProperties) As Long
'
End Function
Public Sub SetPropValue(propID As eProperties, vValue As Long)
'
End Sub
