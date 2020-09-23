VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTitlebarD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Private Enum eTBarOptions
    tbVertical = 1
    tbCaptionWrap = 2
    tbMinMaxBottom = 4      ' not coded yet & therefore not an option (flip btn loc on vertical bars)
    tbVerticalRight = 8
    tbBottomAligned = 16
    tbHidden = 32
End Enum

Private cBorderWidthCx As Byte  ' optional titlebar border sizes
Private cBorderWidthCy As Byte

Private pHwnd As Long           ' copy of hWnd this class is assigned to
Private cFont As Long           ' memory font for the titlebar
Private cAutoRedraw As Boolean  ' whether class can self-refresh

Private cTBarRect As RECT       ' titlebar rect relative to window's 0,0 XY
Private cCaptionRect As RECT    ' caption rect relative to window's 0,0 XY
Private cOptions As Integer ' one or more of the following:

Private cButtons As clsButtons  ' class of system and custom buttons
Private cTBarColors(0 To 1) As clsBarColors ' class of active & inactive color schemes

Private mODImplementation As Long
Private mNotifyUser As Long         ' owner-drawn flags

Implements iPropertyBag

Private Sub Class_Initialize()
    Set cTBarColors(1) = New clsBarColors
    Set cTBarColors(0) = New clsBarColors
    cTBarColors(1).pClassUsage = True
    Set Font = Nothing
    Set cButtons = New clsButtons
End Sub

Private Sub Class_Terminate()
    ' clean up
    If cFont Then DeleteObject cFont
    'If cFontVertical Then DeleteObject cFontVertical
    Set cButtons = Nothing
    Set cTBarColors(1) = Nothing
    Set cTBarColors(0) = Nothing
    
End Sub

'// r/w Change titlebar to vertical, left aligned
Public Property Let Alignment(tbAlign As eBarAlignment)
    
    Dim oldOption As Byte
    
    oldOption = cOptions    ' cache in case user passed something bogus
    cOptions = cOptions And Not tbVertical And Not tbVerticalRight And Not tbBottomAligned
    
    Select Case tbAlign
        Case barAlignBottom: cOptions = cOptions Or tbBottomAligned
        Case barAlignRight: cOptions = cOptions Or tbVerticalRight
        Case barAlignLeft: cOptions = cOptions Or tbVertical
        Case barAlignTop: ' no flag needed
        Case Else:
            cOptions = oldOption
            Exit Property
    End Select
    Me.Buttons.RefreshIcon
    Set Me.Font = Nothing 'automatically refreshes window

End Property

Public Property Get Alignment() As eBarAlignment
    If (cOptions Or tbBottomAligned) = cOptions Then
        Alignment = barAlignBottom
    Else
        If (cOptions Or tbVertical) = cOptions Then
            Alignment = barAlignLeft
        Else
            If (cOptions Or tbVerticalRight) = cOptions Then
                Alignment = barAlignRight
            Else
                Alignment = barAlignTop
            End If
        End If
    End If
End Property

'//  r/w Change whether or not Caption can continue on 2 lines
Public Property Let WrapCaption(isMultiline As Boolean)
    If isMultiline Then
        cOptions = cOptions Or tbCaptionWrap
    Else
        cOptions = cOptions And Not tbCaptionWrap
    End If
    RefreshWindow
End Property
Public Property Get WrapCaption() As Boolean
    WrapCaption = ((cOptions Or tbCaptionWrap) = cOptions)
End Property
'// r/w Titlebar font
Public Property Let Font(newFont As StdFont)
    ' prevents error when not using Set Font = Font Object
    Set Font = newFont
End Property

Public Property Set Font(newFont As StdFont)
    ' modifies the font used for the titlebar
    Dim tFont As Long, Orientation As Long
    Dim nFont As LOGFONT
    
    Select Case Me.Alignment
        Case barAlignLeft: Orientation = 900
        Case barAlignRight: Orientation = 2700
        Case Else: Orientation = 0
    End Select
    
    If newFont Is Nothing Then
        If cFont <> 0 Then
            ' the Alignment property will call this and in those cases, we use the
            ' same font, just different orientation. This scenario only happens when
            ' the newFont is Nothing and cFont is a valid handle
            If GetGDIObject(cFont, Len(nFont), nFont) Then
                nFont.lfEscapement = Orientation
                nFont.lfOrientation = Orientation
                nFont.lfCharSet = 1
                tFont = CreateFontIndirect(nFont)
            End If
        End If
    End If
    
    If tFont = 0 Then tFont = FontToHandle(newFont, False, Orientation)
    
    If tFont <> 0 Then
        ' font created, replace the current font with the new one
        If cFont Then DeleteObject cFont
        cFont = tFont
        RefreshWindow
    End If

End Property
Public Property Get Font() As StdFont
    Set Font = HandleToFont(cFont)
End Property



'//////////////// UNEXPOSED - CLASS USE ONLY \\\\\\\\\\\\\
Friend Function MeasureTitlebar(ByVal cRectPtr As Long, ByVal isActive As Byte, _
        rtnAlign As eBarAlignment, minCaptionCx As Long, minCaptionCy As Long) As Long

' called when the WM_NCcalcSize message is processed

'[In]
    'cRectPtr :: pointer to the current client rectangle
    'isActive :: window state
    'rtnAlign :: return value indicating titlebar orientation in relation to client area
    'hasSystemMenu :: true if the window has a system icon
    'newCaption :: the titlebar caption
'[out]
    'pointer to the calculated titlebar rectangle
    

    Dim hDC As Long, hFont As Long      ' DC and font used to measure the font height
    Dim cCaption As String              ' Current window caption
    Dim cRect As RECT                   ' used for measurements
    Dim btnRect As RECT
    Dim cCapRectPtr As Long             ' pointer for passing rectangle to other classes
    Dim wrapOffset As Long              ' titlebar adjustment if caption will wrap to 2nd line
    
    cAutoRedraw = True   ' reset
    
    If (cOptions Or tbHidden) = cOptions Then
        SetRect cTBarRect, 0, 0, 0, 0
        Exit Function
    End If
    
    ' get a DC & place font in it to measure the font height
    hDC = Canvas(False)
    If hDC = 0 Then Exit Function   ' error somewhere; this is huge but shouldn't happen
    hFont = SelectObject(hDC, cFont)

    ' get the current client rectangle.
    CopyMemory cTBarRect, ByVal cRectPtr, &H10
    
    ' because we can have window borders larger than 4pixels, and we  can also set an
    ' individual title bar border, we need to track the minimum size a scrunched window
    ' can be in order to display at least the 1st 4 characters of the title and also
    ' preserve the title bar borders.
    ' This value is added later to the space taken by the system/user buttons and also
    ' the space needed for the borders. All together we track the size a window is
    ' shrunk to, so it does not allow a caption to be completely hidden
    DrawText hDC, "W...", 4, btnRect, DT_CALCRECT Or DT_SINGLELINE
    
    ' get the titlebar caption & measure the height
    cCaption = GetCaption
    DrawText hDC, cCaption, -1, cRect, DT_CALCRECT Or DT_SINGLELINE
    SelectObject hDC, hFont
    
    rtnAlign = Me.Alignment
    
    If rtnAlign < barAlignBottom Then
        ' adjust the titlebar dimensions for a vertical titlebar
        If rtnAlign = barAlignRight Then
            cTBarRect.Left = cTBarRect.Right - cRect.Bottom - 5
        Else
            cTBarRect.Right = cTBarRect.Left + cRect.Bottom + 5
        End If
    Else
        ' adjust the titlebar dimensions for horizontal titlebar
        If rtnAlign = barAlignBottom Then
            cTBarRect.Top = cTBarRect.Bottom - cRect.Bottom - 5
        Else
            cTBarRect.Bottom = cTBarRect.Top + cRect.Bottom + 5
        End If
        
    End If
    InflateRect cTBarRect, -cBorderWidthCx, -cBorderWidthCy
    
    ' refresh the current client rectangle.
    CopyMemory cRect, ByVal cRectPtr, &H10
    
    ' call function to calculate the space needed for titlebar buttons
    ' return value is the available space for the window caption
    cCapRectPtr = cButtons.CalcButtonOffsets(VarPtr(cTBarRect), isActive, rtnAlign)
        
    If cCapRectPtr = 0 Then
        ' no buttons being displayed
        If Len(cCaption) = 0 Then
            SetRect cTBarRect, 0, 0, 0, 0
            cCaptionRect = cTBarRect
        Else
            cCaptionRect = cTBarRect
            If rtnAlign < barAlignBottom Then
                InflateRect cCaptionRect, 0, -4
            Else
                InflateRect cCaptionRect, -4, 0
            End If
        End If
    Else
        ' got the left over space after subtracting the space needed for titlebar buttons
        CopyMemory cCaptionRect, ByVal cCapRectPtr, &H10
        ' adjust the caption rectangle a bit so it isn't against the last button or border
'        If rtnAlign = barAlignRight Then cCaptionRect.Top = cCaptionRect.Top + 3
    End If
    
    ' test for valid rectangle
    If cCaptionRect.Right > cCaptionRect.Left Then
        
        ' see if caption wrapping is desired, whether it applies, and how much
        ' extra space to add to the titlebar if needed
        wrapOffset = WordWrapCaption(hDC, 0, True, cCaption)
        
        If wrapOffset Then
            ' 2-line titlebar is going to be displayed
            ' adjust the caption rectangle & readjust the titlebar buttons
            Select Case Me.Alignment
                Case barAlignBottom
                    cCaptionRect.Top = cCaptionRect.Top - wrapOffset
                Case barAlignTop
                    cCaptionRect.Bottom = cCaptionRect.Bottom + wrapOffset
                Case barAlignLeft
                    cCaptionRect.Right = cCaptionRect.Right + wrapOffset
                Case barAlignRight
                    cCaptionRect.Left = cCaptionRect.Left - wrapOffset
            End Select
        End If
    End If
    
    ' regardless of 2-line titlebar or not, ensure we have a valid rectangle remaining
    If Me.Alignment < barAlignBottom Then
        cTBarRect.Left = cCaptionRect.Left
        cTBarRect.Right = cCaptionRect.Right
        minCaptionCy = btnRect.Right + (cCaptionRect.Top - cTBarRect.Top) + (cTBarRect.Bottom - cCaptionRect.Bottom) + 1 + cBorderWidthCx * 2
        minCaptionCx = cBorderWidthCx * 2
    Else
        cTBarRect.Bottom = cCaptionRect.Bottom
        cTBarRect.Top = cCaptionRect.Top
        minCaptionCx = btnRect.Right + (cTBarRect.Right - cCaptionRect.Right) + 1 + (cCaptionRect.Left - cTBarRect.Left) + cBorderWidthCy * 2
        minCaptionCy = cBorderWidthCy * 2
    End If
    
    If cTBarRect.Right > 0 And cTBarRect.Bottom > 0 Then
        
        ' add the 1-pixel inset for the titlebar
        Select Case rtnAlign
            Case barAlignTop: cRect.Top = cTBarRect.Bottom + 1 + cBorderWidthCx
            Case barAlignLeft: cRect.Left = cTBarRect.Right + 1 + cBorderWidthCx
            Case barAlignRight: cRect.Right = cTBarRect.Left - 1 - cBorderWidthCx
            Case barAlignBottom: cRect.Bottom = cTBarRect.Top - 1 - cBorderWidthCx
        End Select
    
        ' return the new client rectangle dimensions
        CopyMemory ByVal cRectPtr, cRect, ByVal &H10
        ' return the pointer of the titlebar
        If cTBarRect.Right > cTBarRect.Left Then MeasureTitlebar = VarPtr(cTBarRect)

    Else
        '
        rtnAlign = 0
    End If

End Function

Friend Sub DrawTitleBar(ByVal isActive As Byte, ByVal wState As Long)
'[in]
    'hDC :: the DC to draw in
    'isActive :: window state
'[out] nothing

' Draws the titlebar and buttons
    
    If (cOptions Or tbHidden) = cOptions Then Exit Sub
    
    ' abort if we have not titlebar to draw
    If cTBarRect.Left = cTBarRect.Right Then Exit Sub
    
    Dim bkgImg As StdPicture, bkgStyle As eFillStyle, frmRect As RECT
    Dim isDefault As Boolean, bGrayScaled As Boolean
    Dim Cy As Long
    Dim hBrush As Long, hDC As Long
    
    Dim fColor(0 To 1) As Long
    Dim fxType As Long
    Dim cClient As iOwnerDrawn, BKG As BkgAction_LV
    
    hDC = Canvas(True)
    
    If (mNotifyUser And odTitlebarBkg) = odTitlebarBkg And mODImplementation <> 0 Then
        On Error Resume Next ' always, to try & prevent user from crashing
        BKG.hDC = hDC
        BKG.rcItem = cTBarRect
        BKG.rcExtra = cTBarRect
        InflateRect BKG.rcItem, cBorderWidthCx, cBorderWidthCy
        GetObjectFromPointer mODImplementation, cClient
        cClient.OwnerDrawMessage omDrawTBarBkg, isActive + 0, VarPtr(BKG), ""
        Set cClient = Nothing
        SelectClipRgn hDC, ByVal 0& ' always remove any clipping region client may have left behing
        
    Else
    
        ' get the background colors & style
        Call cTBarColors(isActive).GetBackColors(bkgStyle, fColor(0), fColor(1))
        
        ' unless we have a transparent background, don't continue
        If (bkgStyle And Not bfGrayScaled) <> bfTransparent Then
            
            ' get any background image
            Call cTBarColors(isActive).GetImageBackground(bkgImg, fxType, bGrayScaled)
            
            ' draw the background image if needed
            If Not bkgImg Is Nothing Then
                FillBarImage bkgImg, hDC, cTBarRect.Left, cTBarRect.Top, cTBarRect.Right - cTBarRect.Left, cTBarRect.Bottom - cTBarRect.Top, fxType, (Me.Alignment < barAlignBottom), bGrayScaled
                Set bkgImg = Nothing
                
            Else
            
                bGrayScaled = ((bkgStyle And bfGrayScaled) = bfGrayScaled)
                bkgStyle = bkgStyle And Not bfGrayScaled
                If bkgStyle = bfGradientEW Or bkgStyle = bfGradientNS Then
                    Cy = cTBarRect.Bottom
                    ' check for vertical titlebar options & create the titlebar rect
                    If Me.Alignment < barAlignBottom Then
                        isDefault = cTBarColors(isActive).pIsDefaultGradient
                        If Not Me.Alignment = barAlignRight Then
                            If isDefault Then
                                ' special case when using default colors & vertical titlebar
                                ' I like to see the dark to light shading reversed
                                fColor(0) = fColor(0) Xor fColor(1)
                                fColor(1) = fColor(0) Xor fColor(1)
                                fColor(0) = fColor(0) Xor fColor(1)
                            End If
                        End If
                        If isDefault Or bkgStyle = bfGradientNS Then Cy = -Cy
                    Else
                        If bkgStyle = bfGradientNS Then Cy = -Cy
                    End If
                    ' the Cy value is negative to draw gradients vertical else drawn horizontally
                    GradientFill fColor(0), fColor(1), hDC, cTBarRect.Left, cTBarRect.Top, cTBarRect.Right, Cy, 0, , , bGrayScaled
                    
                ElseIf bkgStyle = bfSolid Then
                    ' solid background, easy
                    If bGrayScaled Then GrayScaleColor VarPtr(fColor(0))
                    hBrush = CreateSolidBrush(fColor(0))
                    FillRect hDC, cTBarRect, hBrush
                    DeleteObject hBrush
                End If
            End If
        
        End If
        
        Call cTBarColors(isActive).GetFrame(fxType, fColor(0), fColor(1))
        Select Case fxType
        Case bxFlat
            hBrush = CreateSolidBrush(fColor(0))
            FrameRect hDC, cTBarRect, hBrush
            DeleteObject hBrush
        Case bx3D
            hBrush = SelectObject(hDC, CreatePen(0, 1, fColor(0)))
            MoveToEx hDC, cTBarRect.Left, cTBarRect.Bottom - 1, ByVal 0&
            LineTo hDC, cTBarRect.Left, cTBarRect.Top
            LineTo hDC, cTBarRect.Right - 1, cTBarRect.Top
            DeleteObject SelectObject(hDC, CreatePen(0, 1, fColor(1)))
            LineTo hDC, cTBarRect.Right - 1, cTBarRect.Bottom - 1
            LineTo hDC, cTBarRect.Left, cTBarRect.Bottom - 1
            DeleteObject SelectObject(hDC, hBrush)
        End Select
    
    End If
    
    ' now draw the caption
    WordWrapCaption hDC, isActive, False, ""
    
    ' call routine to draw the buttons
    cButtons.DrawButtons hDC, isActive
    
End Sub


Private Function WordWrapCaption(ByVal hDC As Long, isActive As Byte, _
        Optional isMeasuring As Boolean, Optional actualCaption As String) As Long

    If isMeasuring = True Then
        ' when measuring the menubar, and no caption wrapping wanted, no extra space needed
        If Me.WrapCaption = False Then Exit Function
    End If
    
    Dim bWrap As Boolean                    ' indicates if wrapping is needed
    Dim vRect As RECT, tmpRect As RECT      ' temporary RECT structures
    Dim I As Integer, iLen As Integer       ' loop & caption length variables
    Dim adjCaption As String                ' truncated/adjusted caption
    Dim DTFlags As Long                     ' DrawText flags
    Dim hFont As Long                       ' font handle
    Dim wrapOffset As Long                  ' positioning offset for 2-line captions

    ' the caption won't be passed from the DrawMenuBar routine, but will from MeasureMenuBar
    If actualCaption = "" Then actualCaption = GetCaption
    If Len(actualCaption) = 0 Then Exit Function

    ' set up flags for measuring the caption & cache a copy of the caption
    DTFlags = DT_SINGLELINE Or DT_MODIFIABLE Or DT_CALCRECT Or DT_NOPREFIX Or DT_WORD_ELLIPSIS
    adjCaption = actualCaption
    
    ' select our font into the DC
    hFont = SelectObject(hDC, cFont)

    ' calculate the needed rectangle dimension. DrawText will supply the .Bottom element
    If Me.Alignment < barAlignBottom Then
        ' vertical titlebar
        SetRect vRect, 0, 0, cCaptionRect.Bottom - cCaptionRect.Top, 0
    Else
        SetRect vRect, 0, 0, cCaptionRect.Right - cCaptionRect.Left, 0
    End If
    
    ' cache copy of the measured rectangle to possibly reuse later & measure text
    tmpRect = vRect
    DrawText hDC, adjCaption, -1, vRect, DTFlags
    ' at this point, adjCaption may or may not have ellipses (...)
    
    If Me.WrapCaption = True Then
        ' see if the titlebar was big enough to display full caption
        
        If adjCaption <> actualCaption Then
            ' caption is too long for the titlebar rectangle
            
            ' when DrawText adds ellipse to a caption it places a null character
            ' at the break point unless the ellipse overlaps the caption
            iLen = InStr(adjCaption & Chr$(0), Chr$(0)) - 1
            ' locate the 1st space before the break point
            For I = iLen To 1 Step -1
                If Mid$(adjCaption, I, 1) = " " Then Exit For
            Next
            
            ' if I=0 then the break point is on the 1st word of the caption
            ' In that case, we will not show a 2-line caption as it is pointless
            If I > 0 Then
                ' we have a space, now let's get the 2nd line for the caption
                bWrap = True    ' flag indicating 2-line caption
                ' now truncate the second line if needed
                adjCaption = Mid$(actualCaption, I + 1)
                vRect = tmpRect ' use our same rectangle we used in the 1st measurement
                DrawText hDC, adjCaption, -1, vRect, DTFlags
            End If
        End If
        
    End If
    
    If bWrap Then ' the caption needs to be wrapped
        
        If isMeasuring Then
            ' not painting, just return the added titlebar space needed
            If Me.Alignment < barAlignBottom Then
                wrapOffset = cCaptionRect.Right - (cCaptionRect.Left + (vRect.Bottom * 2) + 5)
            Else
                wrapOffset = cCaptionRect.Bottom - (cCaptionRect.Top + (vRect.Bottom * 2) + 5)
            End If
            If wrapOffset < 0 Then WordWrapCaption = Abs(wrapOffset)
        Else ' We are painting the 2-line titlebar
            
            ' cache the original cCaptionRect to replace it later
            tmpRect = cCaptionRect
            DTFlags = DT_NOCLIP Or DT_NOPREFIX
            
            Select Case Me.Alignment
                Case barAlignBottom, barAlignTop
                    ' painting horizontal text is soooo much easier
                    OffsetRect cCaptionRect, 0, (cCaptionRect.Bottom - cCaptionRect.Top - (vRect.Bottom * 2)) \ 2
                    DrawCaption hDC, isActive, DTFlags Or DT_WORDBREAK, Left$(actualCaption, I) & vbNewLine & adjCaption
                Case barAlignLeft
                    SetRect cCaptionRect, (tmpRect.Right - tmpRect.Left) \ 2 + tmpRect.Left, tmpRect.Bottom, 0, 0
                    DrawCaption hDC, isActive, DTFlags, adjCaption
                    OffsetRect cCaptionRect, -vRect.Bottom, 0
                    DrawCaption hDC, isActive, DTFlags, Left$(actualCaption, I)
                Case barAlignRight
                    SetRect cCaptionRect, tmpRect.Left + (tmpRect.Right - tmpRect.Left + 1) \ 2 + vRect.Bottom + 1, tmpRect.Top, 0, 0
                    DrawCaption hDC, isActive, DTFlags, Left$(actualCaption, I)
                    OffsetRect cCaptionRect, -vRect.Bottom, 0
                    DrawCaption hDC, isActive, DTFlags, adjCaption
            End Select
            ' restore the original caption rectangle
            cCaptionRect = tmpRect
        End If
        
    ElseIf isMeasuring = False Then
        ' drawing a non-wrapped titlebar
        
        ' cache the original cCaptionRect to replace it later
        tmpRect = cCaptionRect
        DTFlags = DT_NOCLIP Or DT_NOPREFIX
        
        Select Case Me.Alignment
            Case barAlignBottom, barAlignTop
                DTFlags = DTFlags Or DT_VCENTER Or DT_SINGLELINE
            Case barAlignLeft
                SetRect cCaptionRect, (tmpRect.Right - tmpRect.Left - 1 - vRect.Bottom) \ 2 + tmpRect.Left - 1, tmpRect.Bottom, 0, 0
            Case barAlignRight
                SetRect cCaptionRect, tmpRect.Left + (tmpRect.Right - tmpRect.Left + 1 + vRect.Bottom) \ 2 + 1, tmpRect.Top, 0, 0
        End Select
        DrawCaption hDC, isActive, DTFlags, adjCaption
        ' restore the original caption rectangle
        cCaptionRect = tmpRect
    End If
    
        ' clean up
    SelectObject hDC, hFont
    

End Function

Private Sub DrawCaption(hDC As Long, isActive As Byte, DTFlags As Long, adjCaption As String)
    
' function is the final result of the WordWrapCaption function
' Paint the caption on to the DC using custom colors & styles

    If Len(adjCaption) = 0 Then Exit Sub
    
    Dim fColor(0 To 2) As Long, fx As Long
    
    ' get user-defined text colors and the text style
    Call cTBarColors(isActive).GetTextColors(fx, fColor(0), fColor(1), fColor(2))
        
    StyleText adjCaption, cCaptionRect.Left, cCaptionRect.Top, _
        cCaptionRect.Right, cCaptionRect.Bottom, fx, fColor(), hDC, DTFlags
   
End Sub

Friend Function HitTest(ByVal x As Long, ByVal Y As Long) As Long
    
' Hit test functionality for the titlebar
' Called for the WM_NCHITTEST message
    
    If cTBarRect.Right = cTBarRect.Left Then Exit Function
        
    Dim lRtn As Long
    
    If PtInRect(cTBarRect, x, Y) Then
        lRtn = cButtons.HitTest(x, Y)
        If lRtn = 0 Then lRtn = HTCAPTION
    End If

    HitTest = lRtn

End Function


Public Property Get Buttons() As clsButtons
    If cButtons Is Nothing Then Set cButtons = New clsButtons
    Set Buttons = cButtons
End Property
Public Property Get Active() As clsBarColors
    Set Active = cTBarColors(1)
End Property
Public Property Get Inactive() As clsBarColors
    Set Inactive = cTBarColors(0)
End Property

Friend Property Let pMainImplementation(mHWnd As Long, lImpObj As Long)
    pHwnd = mHWnd
    cButtons.pMainImplementation(pHwnd) = lImpObj
End Property

Public Property Let noRedraw(bNoRedraw As Boolean)
    cAutoRedraw = Not bNoRedraw
    Me.Buttons.noRedraw = bNoRedraw
End Property

Private Sub RefreshWindow()
    If pHwnd Then
        If cAutoRedraw Then ForceRefresh pHwnd
    End If
End Sub
Public Sub ResetToSystemDefaults(Options As eBarResets)
    If (Options And rstFont) = rstFont Or Options = rstAll Then
        If cFont Then DeleteObject cFont
        cFont = 0
        Set Me.Font = Nothing
    End If
    If (Options And rstBorderSizes) = rstBorderSizes Or Options = rstAll Then
        Me.BorderHeight = 0
        Me.BorderWidth = 0
    End If
    cButtons.ResetToSystemDefaults Options
    cTBarColors(1).ResetToSystemDefaults Options
    cTBarColors(0).ResetToSystemDefaults Options
End Sub

'//write only property. Set caption; don't use Me.Caption
Public Property Let Caption(ByVal newCaption As String)
    If pHwnd Then SetWindowText pHwnd, newCaption
End Property

Public Property Get BorderWidth() As Byte
    BorderWidth = cBorderWidthCx
End Property
Public Property Let BorderWidth(Size As Byte)
    cBorderWidthCx = Size
    RefreshWindow
End Property
Public Property Get BorderHeight() As Byte
    BorderHeight = cBorderWidthCy
End Property
Public Property Let BorderHeight(Size As Byte)
    cBorderWidthCy = Size
    RefreshWindow
End Property
Friend Property Get pFont() As Long
    pFont = CopyFont(cFont)
End Property
Private Function GetCaption() As String
    If pHwnd Then
        Dim lRtn As Long, sCaption As String
        ' have class calculate the titlebar. Return value is a pointer to the titlebar Rect
        ' first get the caption; may have been changed
        sCaption = String$(256, 0)
        lRtn = GetWindowText(pHwnd, sCaption, 256)
        If lRtn Then GetCaption = Left$(sCaption, lRtn)
    End If
End Function

Public Property Let ShowTitlebar(bShow As Boolean)
    If bShow Then
        cOptions = cOptions And Not tbHidden
    Else
        cOptions = cOptions Or tbHidden
    End If
    RefreshWindow
End Property
Public Property Get ShowTitlebar() As Boolean
    ShowTitlebar = ((cOptions And tbHidden) = 0)
End Property

Friend Property Get Location() As Long
    If (cOptions And tbHidden) = 0 Then
        ' not hidden, but still may not have titlebar
        If IsRectEmpty(cTBarRect) = 0 Then Location = VarPtr(cTBarRect)
    End If

End Property

Friend Sub OwnerDrawn(cImplementation As Long, odItems As eOwnerDrawn, isClosing As Boolean)

    If (odItems And odTitlebarBkg) = 0 Or isClosing = True Then
        mNotifyUser = 0
        mODImplementation = 0
    Else
        mODImplementation = cImplementation      ' update the implementation pointer
        mNotifyUser = odTitlebarBkg
    End If
    cButtons.OwnerDrawn cImplementation, odItems, isClosing
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : clsTitlebarD.iPropertyBag_Read/WriteProperites
' DateTime  : 9/26/2005
' Author    : LaVolpe
' Purpose   : Set or retrieve custom properites
' Comments  : Used by Export/ImportCustomSettings
'---------------------------------------------------------------------------------------
'
Private Sub iPropertyBag_ReadProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)
    
    Dim sProps() As String
    Dim iBag As iPropertyBag
    
    mODImplementation = lParam
    
    mNotifyUser = propBag.ReadProperty("TB.OD", 0)
    cOptions = propBag.ReadProperty("TB.Options", 0)
    sProps = Split(propBag.ReadProperty("TB.Borders", "0,0"), ",")
    cBorderWidthCx = CByte(sProps(0))
    cBorderWidthCy = CByte(sProps(1))
    Set Me.Font = propBag.ReadProperty("TB.Font", Nothing)
    
    
    Dim sharedImgPtr As Long, iPic As StdPicture
    Dim iGrayScaled As Boolean, iStyle As eBackStyles
    '^^ above are used when trying to share images between active/inactive classes
    '   vs creating separate instances of the same image
    
    Set iBag = Me.Active
    ' if the active class has an image, the sharedImgPtr will be the value
    ' of the image handle at the time the property bag was saved
    iBag.ReadProperties propBag, Limits, sharedImgPtr, Index
    Set iBag = Me.Inactive
    ' The inactive class will also compare it's cached value of its image when
    ' the property bag was saved. If the active & inactive class's image values
    ' are different, the sharedImgPtr below will be zero
    iBag.ReadProperties propBag, Limits, sharedImgPtr, Index
    If sharedImgPtr <> 0 Then
        ' The active and inactive class will share the same image.
        ' Get the inactive class's current image properties
        Me.Inactive.GetImageBackground Nothing, iStyle, iGrayScaled
        ' Get the active class's image
        Me.Active.GetImageBackground iPic
        ' now reassign the inactive class's image properties & image
        Me.Inactive.SetImageBackground iPic, iStyle, iGrayScaled
    End If
    Set iBag = Me.Buttons
    iBag.ReadProperties propBag, Limits, lParam, Index
    Set iBag = Nothing
    
    
    
End Sub

Private Sub iPropertyBag_WriteProperties(propBag As PropertyBag, Limits As eExportConstants, Optional lParam As Long, Optional Index As Long)

    Dim mFont As StdFont
    Dim sysNCM As NONCLIENTMETRICS
    Dim NCM As LOGFONT
    Dim iBag As iPropertyBag
    
    ' don't cache the font if it is the system font
    sysNCM.cbSize = Len(sysNCM)
    ' this will return the system font info
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, sysNCM, 0
    
    GetGDIObject cFont, Len(NCM), NCM
    With sysNCM.lfCaptionFont
        If .lfFaceName = NCM.lfFaceName Then
            If .lfHeight = NCM.lfHeight Then
                If .lfWeight = NCM.lfWeight Then
                    If .lfWidth = NCM.lfWidth Then
                        If .lfItalic = NCM.lfItalic Then
                            If .lfUnderline = NCM.lfUnderline Then
                                If .lfStrikeOut = NCM.lfStrikeOut Then
                                    If .lfPitchAndFamily = NCM.lfPitchAndFamily Then .lfHeight = 0
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
    If sysNCM.lfMenuFont.lfHeight <> 0 Then Set mFont = HandleToFont(cFont)
    propBag.WriteProperty "TB.Font", mFont, Nothing
    propBag.WriteProperty "TB.Options", cOptions, 0
    propBag.WriteProperty "TB.Borders", CStr(cBorderWidthCx) & "," & CStr(cBorderWidthCy), "0,0"
    propBag.WriteProperty "TB.OD", mNotifyUser, 0
    
    Dim sharedImgPtr As Long, iPic As StdPicture
    Dim iGrayScaled As Boolean, iStyle As eBackStyles
    
    Set iBag = Me.Active
    iBag.WriteProperties propBag, Limits, sharedImgPtr
    Set iBag = Me.Inactive
    iBag.WriteProperties propBag, Limits, sharedImgPtr
    Set iBag = Me.Buttons
    iBag.WriteProperties propBag, Limits
    
    Set iBag = Nothing
    
End Sub
