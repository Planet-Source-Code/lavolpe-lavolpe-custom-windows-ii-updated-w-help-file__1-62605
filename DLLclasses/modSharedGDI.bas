Attribute VB_Name = "modSharedGDI"
Option Explicit
' This routine is used primarily for drawing.
' non-Public drawing related functions are encapsulated here
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Private Const OBJ_BITMAP As Long = 7

' GDI32 functions
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32.dll" (ByRef piconinfo As ICONINFO) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetClipRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
'Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByRef lpBits As Any, ByRef lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

' used to convert icons/bitmaps to stdPicture objects
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" _
    (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
    iPic As IPicture) As Long
Private Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type

' custom UDTs for tracking global-use DCs
Private Type DCclient
    hWnd As Long
    wSize As POINTAPI
End Type
Private Type DCclients
    DCsize(-1 To 0) As POINTAPI
    DC As Long
    oldBmp As Long
    primaryBmp(-1 To 0) As Long
    curBmp As Long
    Client() As DCclient
End Type

' Constants used
Public Const IMAGE_ICON As Long = 1
Public Const LR_COPYFROMRESOURCE As Long = &H4000
'------ SystemParametersInfo API
Public Const SPI_GETNONCLIENTMETRICS As Long = 41


' About bitmaps used. The fully skinnable (next version) classes will contain
' bitmaps that hold their borders, titlebars, buttons, etc. The DLL will share
' any skin if more than one window in the project is using the same one.

' That being said, for this version, the offscreen drawing of all windows will
' be done with just 2 shared bitmaps and one global-use DC....
' Think about it - your project has 10 windows subclassed, and the project
' only uses 2 bitmaps and one DC! Very resource friendly.

' The description of the bitmaps follow. But first: When you are owner-drawing
' any portion of the window, the global DC and bitmap are passed to you so you
' can draw on them. Along with the bitmap may be the menubar or titlebar font.
' Do not delete or unselect these objects. Doing so will cause memory leaks
' beyond belief. Most of us wouldn't think about destroying our window's bitmap,
' but some of us are just too curious. You've been forewarned.

' 1. A bitmap to be used for drawing the window frames & interior will
' increase with size as needed and can be as large as a full screen, but will
' also reduce size at key points in a window's life-cycle. Added overhead
' includes 5 functions to keep track of the minimal size needed to paint any
' window being subclassed by this DLL.

' 2. A separate bitmap is used to draw the menu states and simply blt over to the
' window as needed as the menu item's change states. This bitmap will only be as
' large as 2x the largest menu item's height being processed by the DLL, and as
' wide as the widest menu item being processed. This bitmap will never decrease
' in size. Menu items tend to be small & the added overhead to track each
' window's requirements to allow reducing size, or the option to create & destroy
' the bitmap every time it is needed, is not worth it IMO.

' There is one caveat. A menubar that has a picture background (vs solid or gradient)
' will maintain a separate stdPicture object for that background. This is necessary
' to keep an image of the menubar with the stretched/tiled background in order to
' clear space needed to draw a new menu item state. When a solid or gradient
' background is used for the menubar, this stdPicture object is not created.


' Local variables to this module
Private Type ClientList
    hWnd As Long                ' hWnd of the Client
    Bounds As POINTAPI          ' The width/height of the client
End Type
Private mClients() As ClientList    ' only initialized when hUsage=0

Private mhDC As Long                ' global-use DC handle
Private imgDC As Long               ' used for drawing non-stdPic objects
Private mOldBmp As Long             ' original bitmap in DC when DC was created
Private mBitmap(0 To 1) As Long     '0=window bitmap, 1=menu item bitmap
Private mIndex As Byte              'identifies which bitmap is in the global DC
Private mSize(0 To 1) As POINTAPI   'size of the 2 bitmaps
Private mGenUseClipRgn As Long      'used only to test passed DC if user supplied a clipping region
Private gScaleLookup() As Byte    'grayscale look up table

' The functions listed below are used to ensure the global-use DC
' is the proper size to draw a window and also reduces the DC when it can
' so that the bitmap doesn't remain overly large. It resizes in step values

' Canvas: Returns the DC handle and selects the appropriate bitmap
' AddClient: Called whenever a window is first subclassed
' RemoveClient: Called when a subclassed window is being destroyed
' ExpandDC: Called when a window is resizing
' ReduceDC: Called when a window is moved or resized or a window is destroyed
' TerminateClients: Called to release the global-use DC and memory objects

'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.Canvas
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Creates the global use DC
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function Canvas(bPrimary As Long, Optional ImageDC As Boolean) As Long
    
    If mhDC = 0 Then
        Dim gsc As Integer
        mGenUseClipRgn = CreateRectRgn(0, 0, 0, 0)
        ' not yet created, do that know
        mBitmap(0) = MakeBitmap(1, 1, mhDC, True)
        ' create the 2nd temporary global bitmap
        mBitmap(1) = MakeBitmap(1, 1, imgDC, True)
        ' select one into our new DC
        mOldBmp = SelectObject(mhDC, mBitmap(0))
        mIndex = 0
        
        ReDim gScaleLookup(0 To 255)
        For gsc = 1 To 255
            ' cache look up vs recalculating for each pixel when grayscaling
            ' adding 2 to offset integer division: 10,10,190 should =70 but would=69
            ' below, so we will soften it up a bit: 10,10,190 would now be 72 vs 69
            gScaleLookup(gsc) = (((gsc + 2) * 33) \ 100)
            ' note: do not increment 33. Any above total*3>100 may cause Overflow
        Next
    
    End If
        
    If ImageDC Then
    
        Canvas = imgDC
    Else
        
        Dim newIndex As Byte
        
        ' ensure the correct bitmap is selected into our DC
        newIndex = CByte(Abs(bPrimary))
        If newIndex <> mIndex Then SelectObject mhDC, mBitmap(newIndex)
        mIndex = newIndex   ' set the bitmap Index
        ' return the DC handle
        Canvas = mhDC
        
    End If
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.AddClient
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Adds a hWnd to the global-DC's client listing
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub AddClient(cHwnd As Long, Cx As Long, Cy As Long)
    
    ' called whenever a window is being subclassed
    If mhDC = 0 Then Call Canvas(True)
    
    Dim C As Long
    
    If IsArrayEmpty(Not mClients) Then
        ReDim mClients(0)   ' no client yet, add it
    Else
        ' see if the client already exists.
        ' Could only happen when the CustomWindow class is set to nothing
        ' but the window was not destroyed.
        For C = 0 To UBound(mClients)
            If mClients(C).hWnd = cHwnd Then
                ' yep, got this client. Ensure bitmap meets requirements
                ExpandDC Cx, Cy, True, cHwnd
                Exit Sub
            End If
        Next
        ' add the client to our list
        ReDim Preserve mClients(0 To UBound(mClients) + 1)
    End If
    ' update the client's handle
    mClients(UBound(mClients)).hWnd = cHwnd
    ' ensure bitmap meets requirements. Function also tracks client's size
    ExpandDC Cx, Cy, True, cHwnd
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.RemoveClient
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Removes a window from the global DC's client listing
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub RemoveClient(cHwnd As Long)
    
    ' called whenever a subclassed window is being destroyed
    
    Dim C As Long
    Dim Cx As Long, Cy As Long
    
    If Not IsArrayEmpty(Not mClients) Then
        ' find the window that is closing
        For C = 0 To UBound(mClients)
            'found. If only one client, easy enough
            If mClients(C).hWnd = cHwnd Then
                If UBound(mClients) = 0 Then
                    TerminateClients        ' clean up global memory objects
                    Exit Sub
                End If
                Exit For
            End If
        Next
    End If
    ' if not found, nothing else to do
    If C > UBound(mClients) Then Exit Sub
    
    ' resize our client array
    mClients(C) = mClients(UBound(mClients))
    ReDim Preserve mClients(0 To UBound(mClients) - 1)
    
    ' see if our DC can be reduced in size now
    ReduceDC 1, 1, 1, 0

End Sub
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.ExpandDC
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Enlarge a DC if needed
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub ExpandDC(ByVal Cx As Long, ByVal Cy As Long, bPrimary As Boolean, cHwnd As Long)
    
    ' called whenever a WM_CalcSize message is received & window has grown
    
    Dim C As Long
    Dim maxSize As POINTAPI
    Dim Client As Long
    Dim newIndex As Byte
    Dim bExpand As Boolean
    
    If Cx = 0 Or Cy = 0 Then Exit Sub
    If mhDC = 0 Then Call Canvas(True)
    
    ' cache current DC size for comparison
    newIndex = CByte(Abs(bPrimary))
    maxSize = mSize(newIndex)
    
    
    If cHwnd <> 0 Then  ' never zero, but left in for testing purposes
        ' only track the Window bitmap size
        If newIndex = 1 Then
            ' find the client so we can update its widht/height
            Client = -1
            If Not IsArrayEmpty(Not mClients) Then
                For C = 0 To UBound(mClients)
                    If mClients(C).hWnd = cHwnd Then Client = C
                    ' track the maximize size needed
                    If mClients(C).Bounds.x > maxSize.x Then maxSize.x = mClients(C).Bounds.x
                    If mClients(C).Bounds.Y > maxSize.Y Then maxSize.Y = mClients(C).Bounds.Y
                Next
            End If
        
            If Client < 0 Then  ' add the client if not done
                ' testing purposes. All clients are added now when window is first subclassed
                AddClient cHwnd, 0, 0
                Client = UBound(mClients)
            End If
            ' update the client's width/height
            mClients(Client).Bounds.x = Cx
            mClients(Client).Bounds.Y = Cy
            
        End If
    End If
    ' see if the DC needs to be expanded
    If Cx >= maxSize.x Then
        maxSize.x = Cx + 5      ' expand & add a little buffer to prevent expanding more often
        bExpand = True
    End If
    If Cy >= maxSize.Y Then
        maxSize.Y = Cy + 5      ' expand & add a little buffer to prevent expanding more often
        bExpand = True
    End If
    
    If bExpand = True Then
        ' update the DC size.
        mSize(newIndex) = maxSize
        ' if the bitmap is selected, unselect it
        If mIndex = newIndex Then SelectObject mhDC, mOldBmp
        ' delete the bitmap and create a new one
        DeleteObject mBitmap(newIndex)
        mBitmap(newIndex) = MakeBitmap(maxSize.x, maxSize.Y)
        ' select it into the DC & update the index
        SelectObject mhDC, mBitmap(newIndex)
        mIndex = newIndex
        
    End If
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.ReduceDC
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Reduces a DC in size if possible
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub ReduceDC(Cx As Long, Cy As Long, ByVal bPrimary As Boolean, cHwnd As Long)

    ' called when a WM_ExitMoveSize message is sent
    ' or when a subclassed window is destroyed
    
    Dim C As Long
    Dim maxSize As POINTAPI
    Dim Client As Long
    Dim newIndex As Byte
    Dim bReduce As Boolean
    
    If mhDC = 0 Then Exit Sub
    If IsArrayEmpty(Not mClients) Then Exit Sub
    
    ' set the new size as the maximum size
    maxSize.x = Cx
    maxSize.Y = Cy
    
    Client = -1
    For C = 0 To UBound(mClients)
        ' find the client
        If mClients(C).hWnd = cHwnd Then Client = C
        ' calculate the maximum size needed
        If mClients(C).Bounds.x > maxSize.x Then maxSize.x = mClients(C).Bounds.x
        If mClients(C).Bounds.Y > maxSize.Y Then maxSize.Y = mClients(C).Bounds.Y
    Next
        
    If Client > -1 Then ' if not, then we were called when a window was destroyed
        mClients(Client).Bounds.x = Cx
        mClients(Client).Bounds.Y = Cy
    End If
    
    ' see if our DC does need to be resized
    newIndex = CByte(Abs(bPrimary))
    If maxSize.x + 5 < mSize(newIndex).x Then
        bReduce = True
        maxSize.x = maxSize.x + 5   ' add a little buffer to prevent expanding more often
    Else
        maxSize.x = mSize(newIndex).x
    End If
    If maxSize.Y + 5 < mSize(newIndex).Y Then
        bReduce = True
        maxSize.Y = maxSize.Y + 5   ' add a little buffer to prevent expanding more often
    Else
        maxSize.Y = mSize(newIndex).Y
    End If
    If bReduce Then
        ' update our DC size & Index
        mSize(newIndex) = maxSize
        ' unselect the bitmap we are about to resize & delete it
        If newIndex = mIndex Then SelectObject mhDC, mOldBmp
        ' delete the bitmap and create a new one
        DeleteObject mBitmap(newIndex)
        mBitmap(newIndex) = MakeBitmap(maxSize.x, maxSize.Y)
        ' select the bitmap and upate the Index
        SelectObject mhDC, mBitmap(newIndex)
        mIndex = newIndex
    
        'Debug.Print "reduced "; maxSize.X; maxSize.Y
    
    End If
    
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.TerminateClients
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Removes memory objects associated with the global-use DC
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub TerminateClients()
    ' called when all clients have been destroyed
    ' or when the DLL is terminated
    If mhDC <> 0 Then
        SelectObject mhDC, mOldBmp
        If mBitmap(1) <> 0 Then DeleteObject mBitmap(1)
        If mBitmap(0) <> 0 Then DeleteObject mBitmap(0)
        DeleteDC mhDC
        DeleteDC imgDC
        If mGenUseClipRgn <> 0 Then DeleteObject mGenUseClipRgn
        Erase gScaleLookup
    End If
    Erase mBitmap
    Erase mSize
    Erase mClients
    mhDC = 0
    mIndex = 0
    imgDC = 0
    mGenUseClipRgn = 0
End Sub
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.GradientFill
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : fills a rectangle with gradient in multiple directions
' Comments  : TODO -- clean up; can be significantly reduced
'---------------------------------------------------------------------------------------
'
Public Sub GradientFill(ByVal FromColor As Long, ByVal ToColor As Long, _
    hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal X1 As Long, ByVal Y1 As Long, _
    Optional ValidateColors As Boolean, Optional xyOffset As Long, _
     Optional lptStep As Long, Optional bGrayScale As Boolean)

    ' FromColor :: any valid RGB color or system color (i.e vbActiveTitleBar)
    ' ToColor :: any valid RGB color or system color (i.e vbInactiveTitleBar)
    ' hDC :: the DC to draw gradient on
    ' X :: left edge of gradient rectangle
    ' Y :: top edge of gradient rectangle
    ' X1:: right edge of gradient
    ' Y1:: bottom edge of gradient
    '      if Y1 < 0 then to direction of gradient is vertical else horizontal
    ' ValidateColors:: Forces the ConvertVBSysColor routine to trigger
    ' .... Added: 12 Sep 05
    ' Needed to be able to create parital gradients for menubar items
    ' XYoffset is the width of the gradient range (width of the menubar)
    ' lPtStep is where in that width the menu item's left edge would start
    
    If X1 <= x Then Exit Sub
    If Abs(Y1) < Y Then Exit Sub
    
    Dim bColor(0 To 3) As Byte, eColor(0 To 3) As Byte
    Dim clipRect As RECT
    
    
    If ValidateColors Then  ' user calling routine from the cGraphics class
        FromColor = ConvertVBSysColor(FromColor)
        ToColor = ConvertVBSysColor(ToColor)
    End If
    ' quick easy way to convert long to RGB values
    CopyMemory bColor(0), FromColor, &H3
    CopyMemory eColor(0), ToColor, &H3
    
    If bGrayScale Then
        bColor(0) = gScaleLookup(bColor(0)) + gScaleLookup(bColor(1)) + gScaleLookup(bColor(2))
        bColor(1) = bColor(0)
        bColor(2) = bColor(0)
    
        eColor(0) = gScaleLookup(eColor(0)) + gScaleLookup(eColor(1)) + gScaleLookup(eColor(2))
        eColor(1) = eColor(0)
        eColor(2) = eColor(0)
    
    End If

    Dim lWxHx As Long   ' adjusted width/height of gradient rectangle
    Dim lPoint As Long  ' loop variable
    Dim lPtStart As Long ' loop varaible
    ' values to add/subtracted from RGB to show next gradient color
    Dim ratioRed As Single, ratioGreen As Single, ratioBlue As Single
    ' memory DC variables if needed
    Dim hPen As Long, hOldPen As Long

    If xyOffset Then
        lWxHx = xyOffset    ' menuitem highlighting
    Else
        ' calculate width of gradient region
        If Y1 < 0 Then      ' vertical gradient
            lWxHx = Abs(Y1) - Y
        Else                ' horizontal gradient
            lWxHx = X1 - x
        End If
    End If
    
    On Error Resume Next
    ' calculate color step value
    ratioRed = ((eColor(0) + 0 - bColor(0)) / lWxHx)
    ratioGreen = ((eColor(1) + 0 - bColor(1)) / lWxHx)
    ratioBlue = ((eColor(2) + 0 - bColor(2)) / lWxHx)
    On Error GoTo 0

    If xyOffset Then        ' menu item highlighting
        If Y1 < 0 Then      ' supply the actual width of the menu item
            lWxHx = Abs(Y1) - Y
        Else
            lWxHx = X1 - x
        End If
    End If
    
    ' two types of gradient routines used....
    ' this 1st one is never called by the DLL, but can be triggered as a
    ' result of the user employing the cGraphics class to gradient fill a
    ' clipped region
    
    If GetClipRgn(hDC, mGenUseClipRgn) > 0 Then
        ' when the passed DC has clipRegion we will use a row by row gradient
        ' (slower over next routine)
        
        ' For example when filling an octagon shape, the top/left corner of the
        ' octagon should be clipped which would prevent using the primary routine below
        
        ' get the 1st color pen into the DC
        hOldPen = SelectObject(hDC, CreatePen(0, 1, FromColor))
    
        If Y1 < 0 Then          ' vertical gradients
            Y1 = Abs(Y1)
            X1 = X1 + 1         ' LineTo requires going one more pixel than true X1
            For lPoint = Y To Y1 - 1
                ' move to the new row/column and draw line to end
                MoveToEx hDC, x, lPoint, ByVal 0&
                LineTo hDC, X1, lPoint
                ' increment color step value & load next color pen
                lptStep = lptStep + 1
                FromColor = RGB(bColor(0) + lptStep * ratioRed, _
                    bColor(1) + lptStep * ratioGreen, _
                    bColor(2) + lptStep * ratioBlue)
                DeleteObject SelectObject(hDC, CreatePen(0, 1, FromColor))
            Next
            
        Else ' horizontal gradients, same remarks as above
            Y1 = Y1 + 1
            For lPoint = x To X1 - 1
                MoveToEx hDC, lPoint, Y, ByVal 0&
                LineTo hDC, lPoint, Y1
                lptStep = lptStep + 1
                FromColor = RGB(bColor(0) + lptStep * ratioRed, _
                    bColor(1) + lptStep * ratioGreen, _
                    bColor(2) + lptStep * ratioBlue)
                DeleteObject SelectObject(hDC, CreatePen(0, 1, FromColor))
            Next
        End If
        DeleteObject SelectObject(hDC, hOldPen)
    
    Else
        ' this routine uses SetPixel & StretchBlt, but only colors one
        ' row or column, depending on gradient direction and is only
        ' valid if no clip region is used or it is a single, rectangular region
        ' The DLL uses only single, simple, rectangular clipping regions as needed
    
        If Y1 < 0 Then  ' vertical gradient
            Y1 = Abs(Y1)
            ' set the pixel color
            For lPtStart = Y To Y1 - 1
                SetPixelV hDC, x, lPtStart, _
                    RGB(bColor(0) + lptStep * ratioRed, _
                    bColor(1) + lptStep * ratioGreen, _
                    bColor(2) + lptStep * ratioBlue)
                lptStep = lptStep + 1
            Next
            ' now simply stretch the column across the width of the region
            ' On Win98, stretching a 1 pixel row more than 200x size fails. So step it
            For lPtStart = 20 To X1 - x Step 20
                StretchBlt hDC, lPtStart - 20 + x, Y, 20, Y1 - Y, hDC, x, Y, 1, lWxHx, vbSrcCopy
            Next
            StretchBlt hDC, lPtStart - 20 + x, Y, (X1 - x) - (lPtStart - 20), Y1 - Y, hDC, x, Y, 1, lWxHx, vbSrcCopy
            
        Else        ' horizontal gradients
            ' same remarks as above
            For lPtStart = x To X1 - 1
                SetPixelV hDC, lPtStart, Y, _
                    RGB(bColor(0) + lptStep * ratioRed, _
                    bColor(1) + lptStep * ratioGreen, _
                    bColor(2) + lptStep * ratioBlue)
                lptStep = lptStep + 1
            Next
            ' simply stretch the top row down the height of the region
            ' On Win98, stretching a 1 pixel row more than 200x size fails. So step it
            For lPtStart = 20 To Y1 - Y Step 20
                StretchBlt hDC, x, lPtStart - 20 + Y, X1 - x, 20, hDC, x, Y, lWxHx, 1, vbSrcCopy
            Next
            StretchBlt hDC, x, lPtStart - 20 + Y, X1 - x, (Y1 - Y) - (lPtStart - 20), hDC, x, Y, lWxHx, 1, vbSrcCopy
        End If
        
    End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.ConvertVBSysColor
' DateTime  : 8/28/2005
' Author    : LaVolpe
' Purpose   : converts a vbSystemColor variable to a long color variable
'---------------------------------------------------------------------------------------
'
Public Function ConvertVBSysColor(inColor As Long) As Long
' I've never seen the GetSysColor API return an error, but just in case...
    On Error GoTo ExitRoutine
    If inColor < 0 Then
        ' vb System colors (i.e, vbButton face) are negative & contain the &HFF& value
        ConvertVBSysColor = GetSysColor(inColor And &HFF&)
    Else
        ConvertVBSysColor = inColor
    End If
    Exit Function
    
ExitRoutine:
    ConvertVBSysColor = inColor
End Function
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.ConvertHimetrix
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Convert Himetrix values to Pixels
' Comments  : Useful when you don't have ability to use ScaleX/ScaleY
'---------------------------------------------------------------------------------------
'
Public Function ConvertHimetrix(vHiMetrix As Long, asWidth As Boolean) As Long
    'converts himetrix to pixels
    If asWidth Then
        ConvertHimetrix = vHiMetrix * 1440 / 2540 / Screen.TwipsPerPixelX
    Else
        ConvertHimetrix = vHiMetrix * 1440 / 2540 / Screen.TwipsPerPixelY
    End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.RevertHimetrix
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Convert pixels to himetrix
' Comments  : Useful when you don't have ability to use ScaleX/ScaleY
'---------------------------------------------------------------------------------------
'
Public Function RevertHimetrix(vPixels As Long, asWidth As Boolean) As Long
    'converts pixels to himetrix
    If asWidth Then
        RevertHimetrix = vPixels / 1440 * 2540 * Screen.TwipsPerPixelX
    Else
        RevertHimetrix = vPixels / 1440 * 2540 * Screen.TwipsPerPixelY
    End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.FillBarImage
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Fills a region with a stdPicture object
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Sub FillBarImage(barImg As StdPicture, hDC As Long, x As Long, Y As Long, _
            Cx As Long, Cy As Long, Style As eBackStyles, tbarVertical As Boolean, _
            Optional bGrayScaled As Boolean)

    ' this routine uses the stdPictur.Render method to draw
    ' .Render is very powerful, doesn't require a secondary DC as the Blt APIs do.
    '   it works on any stdPicture object (Image, PictureBox, ImageList, etc)
    ' But!!! gotta be careful. Render will crash if the slightest thing is off

    ' verify necessary info provided
    If barImg Is Nothing Then Exit Sub
    If Cx < 1 Or Cy < 1 Then Exit Sub

    Dim picWidth As Long, picHeight As Long
    Dim clipRgn As Long, bltArea As Long
    Dim clipRect As RECT

    'get the image dimensions in pixels
    picWidth = ConvertHimetrix(barImg.Width, True)
    picHeight = ConvertHimetrix(barImg.Height, False)

    ' the DLL won't pass a clipping region, but the user can
    ' via the cGraphics class. Check so we don't override the user's region
    If GetClipRgn(hDC, mGenUseClipRgn) < 1 Then
        ' create our own clipping region
        clipRgn = CreateRectRgn(x, Y, x + Cx, Y + Cy)
        SelectClipRgn hDC, clipRgn
        DeleteObject clipRgn
    End If

    With barImg
        Select Case Style
            Case bsSmartStretch
                ' this option will stretch only if needed, otherwise clipped as needed
                ' When stretched horizontally, it won't scrunch vertically
                ' When stretched vertically, it won't scrunch horizontally
                ' Just a different look and can appear better than simple StretchBlt
                
                If bGrayScaled Then
                    If tbarVertical Then
                        If picWidth > Cx Then
                            GrayScaleImage .Handle, hDC, x, Y, picWidth, Cy
                        Else
                            GrayScaleDC .Handle, hDC, x, Y, Cx, Cy
                        End If
                    Else
                        If picHeight > Cy Then
                            GrayScaleImage .Handle, hDC, x, Y, Cx, picHeight
                        Else
                            GrayScaleDC .Handle, hDC, x, Y, Cx, Cy
                        End If
                    End If
                Else
                
                    ' DO NOT remove any '+ 0' below. Doing so will allow .Render to crash
                    If tbarVertical Then
                        If picWidth > Cx Then
                            .Render hDC + 0, x + 0, Y + 0, picWidth + 0, Cy + 0, 0, .Height, .Width, -.Height, ByVal 0&
                        Else
                            .Render hDC + 0, x + 0, Y + 0, Cx + 0, Cy + 0, 0, .Height, .Width, -.Height, ByVal 0&
                        End If
                    Else
                        If picHeight > Cy Then
                            .Render hDC + 0, x + 0, Y + 0, Cx + 0, picHeight + 0, 0, .Height, .Width, -.Height, ByVal 0&
                        Else
                            .Render hDC + 0, x + 0, Y + 0, Cx + 0, Cy + 0, 0, .Height, .Width, -.Height, ByVal 0&
                        End If
                    End If
                End If
                
            Case bsStretch
                If bGrayScaled Then
                    GrayScaleImage .Handle, hDC, x, Y, Cx, Cy
                Else
                    ' simple stretchBlt using .Render
                    ' DO NOT remove any '+ 0' below. Doing so will allow .Render to crash
                    .Render hDC + 0, x + 0, Y + 0, Cx + 0, Cy + 0, 0, .Height, .Width, -.Height, ByVal 0&
                End If
                
            Case bsTiled
                ' custom tile blt
                ' DO NOT remove any '+ 0' below. Doing so will allow .Render to crash
                
                ' blt the first tile on our off-screen DC
                If picWidth > Cx Or picHeight > Cy Then
                    If picHeight > Cy Then picHeight = Cy
                    If picWidth > Cx Then picWidth = Cx
                    .Render hDC + 0, x + 0, Y + 0, picWidth + 0, picHeight + 0, 0, .Height, _
                        RevertHimetrix(picWidth, True), -(RevertHimetrix(picHeight, False)), ByVal 0&
                Else
                    .Render hDC + 0, x + 0, Y + 0, picWidth + 0, picHeight + 0, 0, .Height, .Width, -.Height, ByVal 0&
                End If
                ' when grayscaling, grayscale just this one tile now vs grayscaling
                ' the entire tiled area later
                If bGrayScaled Then GrayScaleDC hDC, 0, x, Y, x + picWidth - 1, Y + picHeight - 1
                
                ' now we will bitblt the rest using incremental steps
                bltArea = picWidth
                Do Until bltArea * 2 >= Cx
                    BitBlt hDC, bltArea + x, Y, bltArea, picHeight, hDC, x, Y, vbSrcCopy
                    bltArea = bltArea * 2
                Loop
                If bltArea < Cx Then ' need IF so below calc doesn't pass negative blt width
                    ' did as much as we could w/o overdrawing, now blt the remainder
                    BitBlt hDC, bltArea + x, Y, Cx - bltArea, picHeight, hDC, x, Y, vbSrcCopy
                End If
                
                ' do the same vertically, 1st row already done
                bltArea = picHeight
                Do Until bltArea * 2 >= Cy
                    BitBlt hDC, x, bltArea + Y, Cx, bltArea, hDC, x, Y, vbSrcCopy
                    bltArea = bltArea * 2
                Loop
                If bltArea < Cy Then ' need IF so below calc doesn't pass negative blt height
                    ' did as much as we could w/o overdrawing, now blt the remainder
                    BitBlt hDC, x, bltArea + Y, Cx, Cy - bltArea, hDC, x, Y, vbSrcCopy
                End If
                
        End Select
    End With

    If clipRgn Then SelectClipRgn hDC, ByVal 0&

End Sub
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.HandleToPicture
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Converts a memory handle bitmap/icon to a stdPicture object
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function HandleToPicture(ByVal hHandle As Long, isBitmap As Boolean) As IPicture

On Error GoTo ExitRoutine

    Dim pic As PICTDESC
    Dim guid(0 To 3) As Long
    
    ' initialize the PictDesc structure
    pic.cbSize = Len(pic)
    ' TODO: if I expose this function to users, use GetObject and/or GetIconInfo
    '       to self-determine whether passed handle is of bitmap or icon/cursor type
    If isBitmap Then pic.pictType = vbPicTypeBitmap Else pic.pictType = vbPicTypeIcon
    pic.hIcon = hHandle
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    guid(0) = &H7BF80980
    guid(1) = &H101ABF32
    guid(2) = &HAA00BB8B
    guid(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect pic, guid(0), True, HandleToPicture

ExitRoutine:
End Function
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.HandleToFont
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : returns a stdFont object from a memory handle font
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function HandleToFont(ByVal hFont As Long) As StdFont

    Dim tFont As StdFont, nFont As LOGFONT

    ' the memory font for our class (cFont) cannot be null
    ' it was created during class Initialize
    If GetGDIObject(hFont, Len(nFont), nFont) = 0 Then Exit Function
    
    Set tFont = New StdFont
    With tFont
        If InStr(nFont.lfFaceName, Chr$(0)) Then
            .Name = Left$(nFont.lfFaceName, InStr(nFont.lfFaceName, Chr$(0)) - 1)
        Else
            .Name = nFont.lfFaceName
        End If
        .Bold = nFont.lfWeight > 400
        .Italic = nFont.lfItalic <> 0
        .Underline = nFont.lfUnderline <> 0
        .Strikethrough = nFont.lfStrikeOut <> 0
        If nFont.lfHeight < 0 Then
            .Size = (nFont.lfHeight * Screen.TwipsPerPixelY) \ -20
        Else
            .Size = 8.25
        End If
    End With

    On Error Resume Next
    Set HandleToFont = tFont


End Function
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.FontToHandle
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : creates a memory font from a stdFont object
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function FontToHandle(newFont As StdFont, bForMenus As Boolean, Optional Orientation As Long) As Long
    
    Dim nFont As LOGFONT
    
    If newFont Is Nothing Then
        Dim NCM As NONCLIENTMETRICS
        NCM.cbSize = Len(NCM)
        ' this will return the system menu font info
        SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, NCM, 0
        If bForMenus Then
            nFont = NCM.lfMenuFont
        Else
            nFont = NCM.lfCaptionFont
        End If
    Else
        With newFont
            ' convert a stdFont object to a memory font handle
            nFont.lfFaceName = .Name & String$(32, 0)
            nFont.lfHeight = (.Size * -20) / Screen.TwipsPerPixelY
            nFont.lfItalic = Abs(.Italic)
            nFont.lfStrikeOut = Abs(.Strikethrough)
            nFont.lfUnderline = Abs(.Underline)
            nFont.lfWeight = Abs(.Bold) * 300 + 400
        End With
    End If
    nFont.lfEscapement = Orientation
    nFont.lfOrientation = Orientation
    ' attempt to create the font
    nFont.lfCharSet = 1
    FontToHandle = CreateFontIndirect(nFont)

End Function
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.CopyFont
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Copies a memory font
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function CopyFont(hFont As Long) As Long
    If hFont <> 0 Then
        Dim nFont As LOGFONT
        If GetGDIObject(hFont, Len(nFont), nFont) Then
            CopyFont = CreateFontIndirect(nFont)
        End If
    End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.MakeBitmap
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Creates a bitmap based off of desktop DC & returns a new DC handle
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function MakeBitmap(Cx As Long, Cy As Long, Optional rtnHdc As Long, Optional newDC As Boolean = False) As Long
    
    Dim dTop As Long, dDC As Long
    
    dTop = GetDesktopWindow()
    dDC = GetDC(dTop)
    
    ' create the desired bitmap
    If Cx > 0 And Cy > 0 Then MakeBitmap = CreateCompatibleBitmap(dDC, Cx, Cy)
    ' return the DC if desired
    If newDC Then
        rtnHdc = CreateCompatibleDC(dDC)
        SetBkMode rtnHdc, &H3
    End If
    ' release desktop DC & select the bitmap into our DC
    ReleaseDC dTop, dDC

End Function
'---------------------------------------------------------------------------------------
' Procedure : modSharedGDI.StyleText
' DateTime  : 9/11/2005
' Author    : LaVolpe
' Purpose   : Draws text in various styles using multiple color options
' Comments  : See Below
'---------------------------------------------------------------------------------------
'
Public Function StyleText(sText As String, Left As Long, Top As Long, _
            Width As Long, Height As Long, fx As eFX, _
            textColors() As Long, hDC As Long, Optional DTFlags As eDTflags, _
            Optional ValidateColors As Boolean, Optional arrayOffset As Integer)

    Dim tRect As RECT
    Dim I As Long
    Dim arrSize As Long, arrLB As Long
    Dim cRef(0 To 2) As Long
    
    ' determine the the LBound and UBound offsets of the passed array
    If IsArrayEmpty(Not textColors) Then
        arrSize = -1
    Else
        arrLB = LBound(textColors)
        arrSize = UBound(textColors) - arrLB
    End If
    
    ' if an invalid array was passed, use standard text color
    If arrSize < arrayOffset Then cRef(0) = ConvertVBSysColor(vbButtonText)
    
    ' fill the cRef() array with the passed values
    For I = 0 To 2
        If arrSize < I + arrayOffset + arrLB Then
            cRef(I) = cRef(0)
        Else
            cRef(I) = textColors(arrayOffset + I + arrLB)
        End If
        ' the DLL automatically stores colors in proper ColorRef values
        ' but user can pass vbSystemColors via the cGraphics class
        If ValidateColors Then cRef(I) = ConvertVBSysColor(cRef(I))
    Next
    
    ' create the DrawText rectangle
    SetRect tRect, Left, Top, Width, Height

    ' now depending on the text style, do a lot!
    Select Case fx
    Case 0, 1: 'default & flat styles - select primary color
        SetTextColor hDC, cRef(0)
    Case 2: ' sunken - select secondary color
        OffsetRect tRect, 1, 1
        SetTextColor hDC, cRef(1)
    Case 3:  ' raised - select secondary color
        OffsetRect tRect, -1, -1
        SetTextColor hDC, cRef(1)
    Case 4:  ' engraved styles - select tertiary color
        OffsetRect tRect, 1, 1
        SetTextColor hDC, cRef(2)
    End Select
    DrawText hDC, sText, -1, tRect, DTFlags  ' draw text
                
    ' now continue for 2 or 3 color styles
    If fx > 1 Then
        Select Case fx
        Case 2 ' sunken style - select primary color
            OffsetRect tRect, -1, -1
            SetTextColor hDC, cRef(0)
        Case 3 ' raised style - select primary color
            OffsetRect tRect, 1, 1
            SetTextColor hDC, cRef(0)
        Case 4 ' engraved - select secondary color
            OffsetRect tRect, -2, -2
            SetTextColor hDC, cRef(1)
        End Select
        DrawText hDC, sText, -1, tRect, DTFlags ' draw text
                    
        If fx = 4 Then    ' 3 color styles
            OffsetRect tRect, 1, 1
            SetTextColor hDC, cRef(0) ' primary color
            DrawText hDC, sText, -1, tRect, DTFlags ' draw text
        End If
    End If

    Erase cRef

End Function


Public Function SetSmallIcon(ByVal hWnd As Long) As Long

    ' routine sets the small icon displayed on the titlebar, if any.
    ' VB won't always create a small icon that can be returned by a
    ' call to SendMessage WM_GETICON, ICON_SMALL.
    ' So we will create one and set it ourselves. Other routines in
    ' this project call the small icon for painting on the titlebar

    Dim bigIcon As Long, smallIcon As Long

    ' see if a big icon exists. If not, the user has set Me.Icon=Nothing
    
    bigIcon = SendMessage(hWnd, WM_GETICON, 1, ByVal 0&)
    If bigIcon Then ' got a big icon
        ' testing seems to show that creating our own may be better
        ' quality than the one VB gives us; at worse it is the same
        ' quality. So we will always create a new one
        smallIcon = CopyImage(bigIcon, IMAGE_ICON, 16, 16, LR_COPYFROMRESOURCE)
        If smallIcon <> 0 Then PostMessage hWnd, WM_SETICON, 0, ByVal smallIcon
    End If
    
    ' notice we don't destroy the icon we retrieved? This is 'cause we don't
    ' own it. It is a shared resource with the window
    
End Function

Public Sub SetApplicationIcon(hIcon As Long, mainHwnd As Long)
    
    ' hIcon must be a 32x32 pixel icon
    ' mainHwnd can be any open form
    ' This will not work in IDE but
    ' works when compiled
    Dim tHwnd As Long, cParent As Long
    
    Const ICON_BIG As Long = 1
    Const GWL_HWNDPARENT = (-8)
    
    If hIcon = 0 Then
        hIcon = SendMessage(mainHwnd, WM_GETICON, ICON_BIG, ByVal 0&)
        If hIcon = 0 Then Exit Sub
    End If
    ' Get starting point
    tHwnd = GetWindowLong(mainHwnd, GWL_HWNDPARENT)
    ' Get EXE's wrapper class (all VB compiled apps have them)
    
    Do While tHwnd
        cParent = tHwnd
        tHwnd = GetWindowLong(cParent, GWL_HWNDPARENT)
    Loop
    On Error Resume Next
    ' tell the wrapper class what icon to use
    PostMessage cParent, WM_SETICON, ICON_BIG, ByVal hIcon
End Sub

Public Function RotateImage(ByVal hImage As Long, ByVal Rotation As Long) As Long

    ' designed for quick 90/270 degree rotation of a small icon/bitmap.

    Dim srcDC As Long, destDC As Long
    Dim destBMP As Long, oldBMPs(0 To 1) As Long
    Dim x As Long, Y As Long, iCount As Byte
    Dim bmpInfo As BITMAPINFOHEADER, iInfo As ICONINFO
    
    srcDC = Canvas(False)           ' bitmap to rotate
    destDC = Canvas(False, True)    ' rotated bitmap
    
    ' get some information about the bitmap being passed as hImage
    If GetIconInfo(hImage, iInfo) = 0 Then Exit Function
    ' see if we can get information on the mask
    If GetGDIObject(iInfo.hbmMask, Len(bmpInfo), bmpInfo) = 0 Then
        If iInfo.hbmColor <> 0 Then DeleteObject iInfo.hbmColor
        If iInfo.hbmMask <> 0 Then DeleteObject iInfo.hbmMask
        Exit Function
    End If
    
    If iInfo.hbmColor = 0 Then ' we have a black & white icon
        ' too small to rotate?
        If bmpInfo.biHeight < 3 Or bmpInfo.biWidth < 2 Then Exit Function
        
        ' black & white icons are stacked in a single image.
        ' The mask is on the bottom half, if I remember correctly;
        ' but it doesn't matter for this routine
        
        bmpInfo.biHeight = bmpInfo.biHeight \ 2 ' change to 1/2 height
        ' create a B&W bitmap with reversed dimensions
        destBMP = CreateBitmap(bmpInfo.biHeight, bmpInfo.biWidth * 2, 1, 1, 0&)
        ' select source & destinations bitmaps into DCs
        oldBMPs(0) = SelectObject(destDC, destBMP)
        oldBMPs(1) = SelectObject(srcDC, iInfo.hbmMask)
    
        ' simply loop thru pixels & transfer from source to destination
        ' for small images, this is quite fast enough
        If Rotation = 2 Then ' vertical right rotation
            For Y = 0 To bmpInfo.biHeight - 1
                For x = 0 To bmpInfo.biWidth - 1
                    ' draw both the rotated image & mask on same bitmap, offsetting as needed
                    BitBlt destDC, Y, x, 1, 1, srcDC, x, bmpInfo.biHeight - Y - 1, vbSrcCopy
                    BitBlt destDC, Y, x + bmpInfo.biWidth, 1, 1, srcDC, x, bmpInfo.biHeight + bmpInfo.biHeight - Y - 1, vbSrcCopy
                Next
            Next
        Else
            For Y = 0 To bmpInfo.biHeight - 1 ' vertical left rotation
                For x = 0 To bmpInfo.biWidth - 1
                    ' draw both the rotated image & mask on same bitmap, offsetting as needed
                    BitBlt destDC, Y, bmpInfo.biWidth - x - 1, 1, 1, srcDC, x, Y, vbSrcCopy
                    BitBlt destDC, Y, bmpInfo.biWidth + x, 1, 1, srcDC, x, Y + bmpInfo.biHeight, vbSrcCopy
                Next
            Next
        End If
        ' delete and remove the old mask
        DeleteObject SelectObject(srcDC, oldBMPs(1))
        ' now remove the new mask & assign it to the .hbmMask element
        iInfo.hbmMask = SelectObject(destDC, oldBMPs(0))
        
    Else
        ' too small to rotate?
        If bmpInfo.biHeight < 2 Or bmpInfo.biWidth < 2 Then Exit Function
    
        For iCount = 0 To 1
        
            If iCount = 0 Then ' image (color bitmap)
                destBMP = MakeBitmap(bmpInfo.biHeight, bmpInfo.biWidth)
                oldBMPs(1) = SelectObject(srcDC, iInfo.hbmColor)
            Else                ' mask (B&W bitmap)
                destBMP = CreateBitmap(bmpInfo.biHeight, bmpInfo.biWidth, 1, 1, ByVal 0&)
                oldBMPs(1) = SelectObject(srcDC, iInfo.hbmMask)
            End If
            
            oldBMPs(0) = SelectObject(destDC, destBMP)
            
            ' simply loop thru pixels & transfer from source to destination
            ' for small images, this is quite fast enough
            If Rotation = 2 Then ' vertical right rotation
                For Y = 0 To bmpInfo.biHeight - 1
                    For x = 0 To bmpInfo.biWidth - 1
            '            SetPixelV destDC, y, bmpInfo.biWidth - x - 1, GetPixel(srcDC, x, y)
                        BitBlt destDC, Y, x, 1, 1, srcDC, x, bmpInfo.biHeight - Y - 1, vbSrcCopy
                    Next
                Next
            Else
                For Y = 0 To bmpInfo.biHeight - 1 ' vertical left rotation
                    For x = 0 To bmpInfo.biWidth - 1
            '            SetPixelV destDC, y, bmpInfo.biWidth - x - 1, GetPixel(srcDC, x, y)
                        BitBlt destDC, Y, bmpInfo.biWidth - x - 1, 1, 1, srcDC, x, Y, vbSrcCopy
                    Next
                Next
            End If
            
            ' remove and delete the original icon's bitmap (image & then mask)
            DeleteObject SelectObject(srcDC, oldBMPs(1))
            If iCount = 0 Then  ' image
                iInfo.hbmColor = SelectObject(destDC, oldBMPs(0))
            Else                ' mask
                iInfo.hbmMask = SelectObject(destDC, oldBMPs(0))
            End If
        
        Next
    
    End If
    
    RotateImage = CreateIconIndirect(iInfo)
    ' next line won't be true if the icon is B&W, single image
    If iInfo.hbmColor <> 0 Then DeleteObject iInfo.hbmColor
    DeleteObject iInfo.hbmMask


End Function

Public Function GrayScaleDC(hDestDC As Long, Optional hBmp As Long, _
                        Optional ByVal x As Long, Optional ByVal Y As Long, _
                        Optional ByVal X1 As Long = -1, Optional ByVal Y1 As Long = -1) As Boolean

    ' Parameters
    ' hDestDC is an existing DC to do the painting
    ' X, Y, X1, & Y1 are rectangle coordinates for grayscaling
    ' ^^ if X1=-1 then the full DC width will be used
    ' ^^ if Y1=-1 then the full DC height will be used

    If hDestDC = 0 Then Exit Function
    
    Dim bmp As BITMAPINFOHEADER
    Dim bmi As BITMAPINFO
    Dim dibArray() As Byte
    Dim Xb As Long, Yb As Long
    Dim scanWidth As Long, xFrom As Long, xTo As Long
    
    ' get the handle to the bitmap in the DC.
    ' For picture boxes, this is equivalent to Picture1.Image.Handle
    If hBmp = 0 Then
        hBmp = GetCurrentObject(hDestDC, OBJ_BITMAP)
        If hBmp = 0 Then Exit Function
    End If
    
    ' if next line fails, we can't set up our DIB arrays
    If GetGDIObject(hBmp, Len(bmp), bmp) = 0 Then Exit Function
    ' ensure a non-empty bitmap
    If bmp.biHeight < 1 Or bmp.biWidth < 1 Then Exit Function
    
    With bmi.bmiHeader
        ' setup the UDT
        .biSize = Len(bmi.bmiHeader)
        .biBitCount = 24            ' automatically aligns bytes on DWord boundaries
        .biCompression = 0
        .biHeight = bmp.biHeight
        .biWidth = bmp.biWidth
        .biPlanes = 1
        scanWidth = (bmp.biWidth * 3 + 3) And &HFFFFFFFC
        
        ' because we will be setting up an array, we need to make absolutely
        ' sure, we will not go out of bounds. If we do, the APIs will crash VB
        If X1 > .biWidth - 1 Or X1 = -1 Then X1 = .biWidth - 1
        If Y1 > .biHeight - 1 Or Y1 = -1 Then Y1 = .biHeight - 1
        If x < 0 Then x = 0
        If Y < 0 Then Y = 0
        If X1 < x Or Y1 < Y Then Exit Function
        ' cache any adjusted height (Y<>0 and/or Y1<>.biHeight)
        bmp.biHeight = Y1 - Y + 1
        
        ' size byte array for the section to be grayscaled
        ReDim dibArray(0 To (scanWidth * bmp.biHeight) - 1)
        ' get only the section to be grayscaled.
        ' Note funky .biHeight-Y1-1 as the starting Y coordinate. Image is flipped.
        GetDIBits hDestDC, hBmp, .biHeight - Y1 - 1, bmp.biHeight, dibArray(0), bmi, 0&
        ' note that the dibArray array is BRG vs RGB
        
        ' loop thru the bytes, rows then columns
        xFrom = x * 3
        xTo = X1 * 3
        For Yb = 0 To UBound(dibArray) Step scanWidth
            ' since we needed to pull bytes starting at column 0, we need
            ' to adjust the startpoint for each row if user passed a
            ' parameter other than X = 0
            For Xb = Yb + xFrom To Yb + xTo Step 3
                ' add up the gray scale bytes and apply to all 3 source bytes
                dibArray(Xb) = gScaleLookup(dibArray(Xb)) + gScaleLookup(dibArray(Xb + 1)) + gScaleLookup(dibArray(Xb + 2))
                dibArray(Xb + 1) = dibArray(Xb)
                dibArray(Xb + 2) = dibArray(Xb)
                'dibArray(xb+3) not used; the 4th byte in a 32 byte color
            Next
        Next
        
        ' adjust the height in the bmi UDT
        .biHeight = bmp.biHeight
        ' simply paste the changes back to the destination dc
        GrayScaleDC = (StretchDIBits(hDestDC, 0, Y, .biWidth, .biHeight, 0, 0, .biWidth, .biHeight, dibArray(0), bmi, 0, vbSrcCopy) <> 0)
    
    End With
    
        
End Function



Public Function GrayScaleImage(hImage As Long, hDestDC As Long, x As Long, Y As Long, _
                            Optional ByVal Cx As Long, Optional ByVal Cy As Long) As Boolean
                            
    ' Parameters
    ' hImage is the handle to the memory image or [object].Picture.Handle
    ' hDestDC is an existing DC to do the painting
    ' X, Y are left & top coordinates to draw the image
    ' Cx and Cy are the width/height of the drawn image


If hImage = 0 Or hDestDC = 0 Then Exit Function
                            
Dim bmp As BITMAPINFOHEADER
Dim bmi As BITMAPINFO
Dim dibArray() As Byte, Xb As Long, Yb As Long
Dim dibMask() As Byte   ' used for icons only
Dim iInfo As ICONINFO
Dim hBmp As Long
Dim scanWidth As Long

If GetGDIObject(hImage, Len(bmp), bmp) = 0 Then

    ' handle passed is not a bitmap. If not an icon, then abort
    If GetIconInfo(hImage, iInfo) = 0 Then Exit Function
    
    ' it's an icon, but is it valid?
    If iInfo.hbmColor = 0 Then
        ' a black & white icon/cursor. nothing to do as it is already grayscaled
        If iInfo.hbmMask <> 0 Then DeleteObject iInfo.hbmMask
        DrawIconEx hDestDC, x, Y, hImage, Cx, Cy, 0, 0, &H3
        Exit Function
    End If
    ' ok, got a color bitmap
    hBmp = iInfo.hbmColor
    ' get bitmap information for the icon
    If GetGDIObject(hBmp, Len(bmp), bmp) = 0 Then
        ' If we have iInfo.hbmColor, then this test should never be triggered
        ' However, for robutsness, simply paint the icon & delete execess memory items
        DeleteObject hBmp
        If iInfo.hbmMask <> 0 Then DeleteObject iInfo.hbmMask
        DrawIconEx hDestDC, x, Y, hImage, Cx, Cy, 0, 0, &H3
        Exit Function
    End If

Else
    ' bitmap. Transparent GIFs won't work with this routine.
    hBmp = hImage
End If

' ensure we have a non-empty bitmap
If bmp.biHeight < 1 Or bmp.biWidth < 1 Then Exit Function

With bmi.bmiHeader
    ' setup the UDT
    .biSize = Len(bmi.bmiHeader)
    .biBitCount = 24
    .biCompression = 0
    .biHeight = bmp.biHeight
    .biWidth = bmp.biWidth
    .biPlanes = 1
    scanWidth = (bmp.biWidth * 3 + 3) And &HFFFFFFFC
    
    ReDim dibArray(0 To (scanWidth * .biHeight - 1))
    GetDIBits hDestDC, hBmp, 0, .biHeight, dibArray(0), bmi, 0&
    
    For Yb = 0 To UBound(dibArray) Step scanWidth
        For Xb = Yb To Yb + scanWidth - 1 Step 3
            ' add up the gray scale bytes and apply to all 3 source bytes
            dibArray(Xb) = gScaleLookup(dibArray(Xb)) + gScaleLookup(dibArray(Xb + 1)) + gScaleLookup(dibArray(Xb + 2))
            dibArray(Xb + 1) = dibArray(Xb)
            dibArray(Xb + 2) = dibArray(Xb)
        Next
    Next
    
    ' use the image width/height if user passed zeros
    If Cx < 1 Then Cx = .biWidth
    If Cy < 1 Then Cy = .biHeight

    If iInfo.hbmColor Then  ' icon was passed
        dibMask = dibArray
        GetDIBits hDestDC, iInfo.hbmMask, 0, .biHeight, dibMask(0), bmi, 0&
        StretchDIBits hDestDC, x, Y, Cx, Cy, 0, 0, .biWidth, .biHeight, dibMask(0), bmi, 0, vbSrcAnd
        GrayScaleImage = (StretchDIBits(hDestDC, x, Y, Cx, Cy, 0, 0, .biWidth, .biHeight, dibArray(0), bmi, 0, vbSrcPaint) <> 0)
        ' the bitmaps returned by a call to GetIconInfo must be deleted else memory leaks
        DeleteObject hBmp
        If iInfo.hbmMask <> 0 Then DeleteObject iInfo.hbmMask
    Else
        GrayScaleImage = (StretchDIBits(hDestDC, x, Y, Cx, Cy, 0, 0, .biWidth, .biHeight, dibArray(0), bmi, 0, vbSrcCopy) <> 0)
    End If
    
End With


End Function

Public Sub GrayScaleColor(ByVal inColor As Long)
    Dim bColor(0 To 3) As Byte
    CopyMemory bColor(0), ByVal inColor, &H4
    bColor(0) = gScaleLookup(bColor(0)) + gScaleLookup(bColor(1)) + gScaleLookup(bColor(2))
    bColor(1) = bColor(0)
    bColor(2) = bColor(0)
    CopyMemory ByVal inColor, bColor(0), &H4
End Sub
