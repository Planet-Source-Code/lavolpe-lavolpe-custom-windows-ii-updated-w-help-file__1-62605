VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysTrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' collection of tray icons.
' Each hWnd can have multiple tray icons.

' Note that I use routines to search the collection using simple For:Next
' Since it is highly unlikely that a single hWnd would
' contain more than 1 or 2 tray icons, the overhead of additional class collection
' levels or more advanced searching routines is wasteful IMO

Private colTray As Collection ' Collection of system Tray Icons for a single hWnd
Implements iPropertyBag

Private Sub Class_Initialize()
    Set colTray = New Collection ' initialize collection
End Sub

Private Sub Class_Terminate()
    Set colTray = Nothing ' destroy collection & automatically removes any tray icons
End Sub

Public Property Get TrayIcon(hWnd As Long, TrayIconID As Long) As clsSysTray
' Tray icon as referenced throughout is not the image, it is the placeholder
' that the image rests on; it is the tray area that recieves mouse actions.

' TrayIconId :: a unique identifier you assign for the tray icon. Used for multiple
'   icons from a single hWnd. Yes you can put more than 1 icon on the
'   system tray for the same hWnd
' IMPORTANT. This TrayIconID is used whenever you want to modify the system tray
'   like changing tips, showing balloons or changing icons, activating or deactivating
'   Additionally, this TrayIconID is the wParam value within your ProcessTrayIcon event

' Suggestion. If you are only going to have a single tray icon for your hWnd, then
' use the hWnd as the TrayIconID which prevents you from having to cache another
' variable within your project.

On Error Resume Next
Dim TrayNotifier As clsSysTray

    ' see if we have this tray icon in our collection
    If colTray("t" & TrayIconID) Is Nothing Then
        
        If Err Then Err.Clear   ' should always be an error if we get here
        
        If colTray Is Nothing Then Set colTray = New Collection
        
        ' create a new class
        Set TrayNotifier = New clsSysTray
        
        ' set its internal flag to know which hWnd it will be assigned to
        TrayNotifier.PrepTray hWnd, TrayIconID
        
        ' add the class to our collection & return
        colTray.Add TrayNotifier, "t" & TrayIconID
        Set TrayNotifier = Nothing
        
    Else
        
        If hWnd <> 0 Then
            ' set its internal flag to know which hWnd it will be assigned to
            TrayNotifier.PrepTray hWnd, TrayIconID
        
        Else
            Exit Property
            
        End If
    
    End If
    
    Set TrayIcon = colTray("t" & TrayIconID)

End Property

Public Sub RemoveTrayIcon(TrayIconID As Long, RemoveAll As Boolean, DeleteIcon As Boolean)

    ' used by DLL to remove all tray icons or a single icon
    ' Generally as a result of unsubclassing or user-decision
    
    Dim I As Integer, nrActive As Long
    Dim TrayNotifier As clsSysTray
    
    If RemoveAll = True Then
        If DeleteIcon Then
            For I = 1 To colTray.Count
                colTray.Item(I).DestroyTrayIcon True
            Next
        End If
        Set colTray = Nothing
    Else
        For I = colTray.Count To 1 Step -1
            Set TrayNotifier = colTray.Item(I)
            If TrayNotifier.uniqueID = TrayIconID Then
                If DeleteIcon Then TrayNotifier.DestroyTrayIcon True
                colTray.Remove I
                Exit For
            End If
        Next
    End If
End Sub

Public Sub RestoreSysTray()
' called by the DLL to restore tray icons after an Explorer crash
' also may be called to remove all tray icons in preparation for an IDE GPF crash
' Only those that were active will be restored

    Dim I As Integer
    Dim tTrayClass As clsSysTray
    
    For I = 1 To colTray.Count
        
        ' need to get hard reference for Friend functions to work
        Set tTrayClass = colTray(I)
        tTrayClass.Restore
        Set tTrayClass = Nothing
    
    Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : clsSysTrays.iPropertyBag_Read/WriteProperites
' DateTime  : 9/26/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve custom properites
' Comments  : Used by Export/ImportCustomSettings
'---------------------------------------------------------------------------------------
'
Private Sub iPropertyBag_ReadProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim I As Long
    Dim trayClass As clsSysTray
    Dim uniqueID As Long, sProp As String
    Dim iBag As iPropertyBag
    
    Set colTray = New Collection
    For I = 0 To propBag.ReadProperty("ST.Count", 0) - 1
        Set trayClass = New clsSysTray
        sProp = propBag.ReadProperty("ST" & I + 1 & ".ID", "hwnd")
        If sProp = "hwnd" Then
            uniqueID = lParam
        Else
            uniqueID = Val(sProp)
        End If
        
        Set iBag = trayClass
        iBag.ReadProperties propBag, Limits, lParam, I + 1
        Set iBag = Nothing
        
        colTray.Add trayClass, "t" & uniqueID
        Set trayClass = Nothing
    Next
    
End Sub

Private Sub iPropertyBag_WriteProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)
    
    If colTray Is Nothing Then
        propBag.WriteProperty "ST.Count", 0, 0
        Exit Sub
    End If
    
    Dim I As Integer
    Dim iBag As iPropertyBag
    
    For I = 1 To colTray.Count
        
        Set iBag = colTray.Item(I)
        iBag.WriteProperties propBag, Limits, lParam, I + 0
        Set iBag = Nothing
    Next
    propBag.WriteProperty "ST.Count", I - 1
    
End Sub
