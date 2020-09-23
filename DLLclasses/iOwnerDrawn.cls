VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iOwnerDrawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub OwnerDrawMessage(ByVal wMsg As eODmessageConstants, ByVal wParam As Long, _
                            ByVal lParam As Long, ByVal dllCaption As String)

' (*) below indicates the UDT at lParam must be updated on return of sub
' (&) below indicates the UDT at lParam can be updated to force your changes
' wMsg [in] will be one of the following Owner callback types
' (*) omMeasureMenuItem :: measure a menu item
'     omDrawMenuItem :: draw a menu item or menu image as applicable
'     omDrawUserBtn :: draw custom title bar button
'     omDrawSysBtn :: custom draw the min/max/restore/close buttons
'     omDrawMBarBkg :: fill the menubar
'     omDrawTBarBkg :: fill the titlebar
'     omDrawPostNC :: add any final touches to the NC after window is drawn
'     omDrawPreNC :: draw custom window background/borders
'     omUserBtnClick :: custom button clicked
' (&) omDrawTrackRect :: optional drawing of a tracking rectangle in the NC
'     omTrackClick :: a tracking rectangle was clicked
' (&) omTrackCursor :: cursor over a tracking rectangle, optionally set the cursor

' wParam, lParam and dllCaption meanings are dependent upon the value of wMsg
'   each of the messages are explained below, along with the values of the
'   appropriate UDT (BkgAction_LV or CustomItemDraw_LV), and any return values

End Sub

' Two structures you will need to declare in your application

' When measuring menu items and also drawing menu items, custom buttons, tracking rects
'Type CustomItemDraw_LV
'    itemID As Long
'    itemPos As Long
'    itemData As Long
'    itemState As Long
'    itemOD As Long
'    hDC As Long
'    rcItem As RECT
'End Type
'
'
' for drawing backgrounds on the window, menubars and titlebars
'Type BkgAction_LV
'    hDC As Long                 ' DC used for drawing
'    rcItem As RECT              ' Rect used for drawing
'    rcExtra As RECT             ' not used
'End Type



'////////////////////// wMsg = omMeasureMenuItem \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To measure a menu item you are drawing as a result of passing
' the odMenuItem_Complete or the odMenuItem_ImageOnly flag to the
' OwnerDrawn function of the DLL
'
' wParam is not used
' lParam is a pointer to a CustomItemDraw_LV structure
' dllCaption is the menu item's caption
'
'CustomItemDraw_LV Values
'   itemID. The menu item identifier returned by Windows
'   itemPos. The zero-bound position of the item on the menubar
'       ** Note that menu items where .Visible=False are not considered
'          part of the menubar. Tip. Use array for your top level menu items
'   itemData. Any value you provide to help uniquely identify this item
'       ** Tip. Good idea to assign the image reference for your menu item here
'   itemState. 0 if item is enabled or 1 if it is disabled
'   itemOD.either odMenuItem_Complete Or odMenuItem_ImageOnly
'       ** When odMenuItem_ImageOnly, no measurement is needed. This is provided
'           so you can tell the DLL the image alignment and also assign the
'           itemData portion of the UDT as you see fit. The return value
'           must contain odImgAlignRight to align image to right of menu item
'           or odImgAlignLeft to algin image left of the menu item, otherwise
'           any other value will be assumed to mean no image
'   hDC. a DC with correct menu font already selected. Do not remove the font
'        or the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem.
'       If itemOD is odMenuItem_Complete then you must fill in the size of the
'           menu item you need to draw your text and any image and/or borders
'       If itemOD is odMenuItem_ImageOnly, not used.

' Final notes:
'   if itemOD = odMenuItem_ImageOnly then itemOD must be changed & updated if drawing an image
'   if itemOD = odMenuItem_Complete then rcItem must be changed & updated
'   for both, itemData should be set to a value that is meaningful to you



'////////////////////// wMsg = omDrawMenuItem \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a menu item image or the entire menu item (image
' borders, caption, etc) as a result of passing the odMenuItem_Complete
' or the odMenuItem_ImageOnly flag to the OwnerDrawn function of the DLL
'
' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a CustomItemDraw_LV structure
' dllCaption is the menu item's caption
'
'CustomItemDraw_LV Values
'   itemID. The menu item identifier returned by Windows
'   itemPos. The zero-bound position of the item on the menubar
'       ** Note that menu items where .Visible=False are not considered
'          part of the menubar. Tip. Use array for you top level menu items
'   itemData. Any value you provided during the omMeasureMenuItem message
'   itemState. will contain one or more of the following values using OR
'               mcStandard - item is not selected, not hovered over
'               mcHover - mouse is currently over the menu item
'               mcSelect - menu item is currently selected
'               mcDisabled - menu item is disabled
'   itemOD.either odMenuItem_Complete Or odMenuItem_ImageOnly
'           odMenuItem_ImageOnly. draw your image in the rcItem provided
'           odMenuItem_Complete. draw the entire menu item in the rcItem provided
'   hDC. a DC with correct menu font already selected. Do not remove the font
'        or the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. The rectangle to draw. This rectangle is NOT clipped and you should
'           ensure your drawing stays within its bounds
'
' Final notes:
'   if itemOD = odMenuItem_ImageOnly then simply draw the image
'   If itemOD = odMenuItem_Complete Then draw everything: Text , Image, borders



'////////////////////// wMsg = omDrawUserBtn \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a custom button. This message is
' automatically sent when you added your button thru the
' TitleBar.Buttons.AddButton function
'
' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a CustomItemDraw_LV structure
' dllCaption is the tool tip you provided when you added the button
'
'CustomItemDraw_LV Values
'   itemID. The button ID you provided when you added the button
'   itemPos. The zero-bound position of the button (0 thru 4)
'   itemData.not used
'   itemState. will contain one or more of the following values using OR
'               mcStandard - item is not selected, not hovered over
'               mcHover - mouse is currently over the menu item
'               mcSelect - menu item is currently selected
'               mcDisabled - menu item is disabled
'   itemOD.not used
'   hDC. a DC to draw in. Do not remove the font (without replacing it) or
'       the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. The rectangle to draw. This rectangle is NOT clipped and you should
'           ensure your drawing stays within its bounds.
'       ** Note. The DLL provides a blank button to draw on having a 2 pixel border.
'           You can draw within that button (rcItem dimensions) or draw over it.


'////////////////////// wMsg = omDrawSysBtn \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a system button as a result of passing
' the odSysButtons flag to the OwnerDrawn function of the DLL
'
' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a CustomItemDraw_LV structure
' dllCaption is not used
'
'CustomItemDraw_LV Values
'   itemID. will be one of the following
'           SC_CLOSE. Draw the close button
'           SC_MINIMIZE. Draw the minimize button
'           SC_MAXIMIZE. Draw the maximize button
'           SC_RESTORE. Draw the Restore Down button (window is maximized)
'   itemPos.not used
'   itemData.not used
'   itemState. will contain one or more of the following values using OR
'               mcStandard - item is not selected, not hovered over
'               mcHover - mouse is currently over the menu item
'               mcSelect - menu item is currently selected
'               mcDisabled - menu item is disabled
'   itemOD.not used
'   hDC. a DC to daw in. Do not remove the font (without replacing it) or
'           the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. The rectangle to draw. This rectangle is NOT clipped and you should
'           ensure your drawing stays within its bounds.
'       ** Note. The DLL provides a blank button to draw on having a 2 pixel border.
'           You can draw within that button (rcItem dimensions) or draw over it.



'////////////////////// wMsg = omDrawMBarBkg \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw the menubar background as a result of passing
' the odMenuBarBkg flag to the OwnerDrawn function of the DLL
'
' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a BkgAction_LV structure
' dllCaption is not used
'
'BkgAction_LV Values
'   hDC. a DC to draw in. Do not remove the font (without replacing it) or the
'           bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. the menubar boundaries. This is NOT clipped
'       ** This could be a zero size rectangle if window is dragged very small
'       ** Background will always be erased with the current Frame.Backcolor value
'           or filled with the current gradient colros/image unless the
'           BackStyle property is set to bfTRansparent
'   rcExtra. the rectangle where the menubar items will be drawn. This is NOT clipped
'           and you should ensure your drawing stays within its bounds
'

'////////////////////// wMsg = omDrawTBarBkg \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a titlebar background as a result of passing
' the odTitlebarBkg flag to the OwnerDrawn function of the DLL
'
' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a BkgAction_LV structure
' dllCaption is not used
'
'BkgAction_LV Values
'   hDC. a DC with correct titlebar font already selected. Do not remove the font
'        or the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. the titlebar boundaries. This is NOT clipped
'       ** This could be a zero size rectangle if window is dragged very small
'       ** The caption rectangle portion of the titlebar will be drawn with the
'           current solid/gradient colors/image unless you set the
'           BackStyle property to bfTransparent
'   rcExtra. the rectangle where the menubar items will be drawn. This is NOT clipped
'           and you should ensure your drawing stays within its bounds



'////////////////////// wMsg = omDrawPostNC \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw anything you want on the nonclient and
' is received as a result of passing the odPostNCDrawing flag
' to the OwnerDrawn function of the DLL
'
' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a BkgAction_LV structure
' dllCaption is not used
'
'BkgAction_LV Values
'   hDC. a DC to draw in. Do not remove the font (without replacing it) or
'           the bitmap from the DC. Doing so will prevent your window from drawing
'   rcItem. the window boundaries where top/left is always 0,0
'   rcExtra. this is not a standard rectangle. The elements are pointers to other rectangles
'       rcExtra.Left is a pointer to the client rectangle
'       rcExtra.Top is a pointer to the complete titlebar rectangle
'       rcExtra.Right is a pointer to the complete menubar rectangle
'       rcExtra.Bottom is a pointer to the non-client inset
'       ** Using that information, you can calculate available space to add any
'           additional graphics or Tracking Rectangles (explained later)
'       ** to copy a pointer (rcExtra.Left) to your rectangle variable:
'               CopyMemory myRectVariable, ByVal rcExtra.Left, &H10



'////////////////////// wMsg = omDrawPreNC \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a window border as a result of passing
' the odFrameBorders flag to the OwnerDrawn function of the DLL.
' Here you have the opportunity to completely draw the entire nonclient
'
' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a BkgAction_LV structure
' dllCaption is not used
'
'BkgAction_LV Values
'   hDC. a DC with to draw in. Do not remove the font or the bitmap from the DC.
'           Doing so will prevent your window from drawing
'   rcItem. the window boundaries where top/left is always 0,0
'   rcExtra. the nonclient area remaining after the border width/height are subtracted
'
' ** Note: If you are drawing the background of the nonclient area, be aware that
'   if the menubar and titlebar Backstyle properties are not set to bfTransparent,
'   then those objects will be filled in as normal.



'////////////////////// wMsg = omUserBtnClick \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To inform user that a custom button was clicked
' Note. Do not interpret the itemState value of mcSelect in the
'   omDrawUserBtn message as a button click. When a user clicks down
'   on a button a omDrawUserBtn message is sent, but if user does not
'   release the mouse and drags mouse off the button, it is not a click
'   event even though another omDrawUserBtn message will be sent with
'   a itemState value of mcStandard to draw the button in the up state
'
' wParam is the button ID of the custom button you assigned
' lParam is not used, it is Null
' dllCaption is tooltip text of the button



'////////////////////// wMsg = omDrawTrackRect \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To draw a item you are tracking that exists within
' the nonclient area. This message is a result of calling
' the Frame.AddTrackingRect function. Processing this message is optional
'
' wParam is 0 if window is not active or 1 if window is active
' lParam is a pointer to a CustomItemDraw_LV structure
' dllCaption is not used
'
'CustomItemDraw_LV Values
'   itemID. The tracking rectangle you provided when you called AddTrackingRect
'   itemPos.not used
'   itemData.not used
'   itemState. will contain one or more of the following values using OR
'               mcStandard - item is not selected, not hovered over
'               mcHover - mouse is currently over the menu item
'               mcSelect - menu item is currently selected
'   itemOD. You must provide a non-zero value if you drew the updated rectangle
'           into the DC and want it posted to the visible window
'   hDC. a DC with operating system's standard font selected.
'           Do not remove the font without replacing it. Memory leaks would occur
'   rcItem. The rectangle to draw in
'       ** The DLL does not graphically track or replace the space occupied by
'           this rectangle. You should have cached the background when processing
'           the omDrawPostNC message and copying the background to your own bitmap
'       ** Tip. By drawing text only or redrawing text, using the same text styles
'           and font, no need at all to cache the background image.
'
'   Final Notes: You are required to update the itemOD value if you redrew the
'           rectangle and want it posted to the visible window



'////////////////////// wMsg = omTrackClick \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To inform user that a tracking rectangle was clicked
' Note. Do not interpret the itemState value of mcSelect in the
'   omDrawTrackRect message as a button click. When a user clicks down
'   on a tracker, a omDrawTrackRect message is sent, but if user does not
'   release the mouse and drags mouse off the tracker, it is not a click
'   event even though another omDrawTrackRect message will be sent with
'   a itemState value of mcStandard to draw the omDrawTrackRect in the up state
'
' wParam is the tracking rectangle ID you assigned
' lParam is 0 if window is inactive, or 1 if active
'   ** This value may also contain mcRightButton if the click was a right button
' dllCaption is not used



'////////////////////// wMsg = omTrackCursor \\\\\\\\\\\\\\\\\\\\
' PURPOSE: To inform user that the mouse is over a tracking rectangle
' When this message is received you can use the SetCursor API to
' assign a cursor specific to whatever the tracking rectangle applies to
'
'
' wParam is the tracking rectangle ID you assigned
' dllCaption is not used
'
' lParam is a pointer to a long value, 4 bytes.
'   ** If you are assigning your own cursor, you must set this parameter to a
'   non-zero value so the DLL does not replace your cursor with an standard pointer.
'   To replace the lParam value:  CopyMemory ByVal lParam, &H1, &H4

