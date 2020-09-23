VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iCoreMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Scroll further down to see the history of changes to this project

Public Sub WindowMessage(ByVal hWnd As Long, ByVal wMsg As Long, _
        ByRef wParam As Long, ByRef lParam As Long, _
        ByRef bBlockMessage As Boolean, ByRef BlockValue As Long)

End Sub



'///////////////////////////////////////////////////////////////////////////////////////
' Using the bBlockMessage & BlockValue, you can override the return value expected
' by the WindowProcedure and usage of the class function ForwardMessage, you can either
' modify or completely discard the message being passed.
' To override the message, set bBlockMessage=True & change BlockValue as needed
'   Note: Most window messages require BlockValue to be zero, but not all
'         You should always look up the blocked return value for the message @ MSDN
'
'   Your WindowMessage event is similar to a standard WindowProcedure routine
'   The DLL 's public function ForwardMessage is identical usage of the CallWndProc API
'
' When you implement this class, ALL messages you have identified via the
' MessagesAdd and/or MessagesRemove functions will be forwarded to your window
' for the form that is subclassed. You can override these messages with few exceptions.
'
' WM_Destroy can never be overridden. It is not even sent to the CustomWindow class
'   To prevent your form from closing, you should process the form's QueryUnload event
'
' WM_NCHitTest will be calculated, then forwarded to you for your info.
'   You cannot change the value of this message (BlockValue)
'   The BlockValue will be the calcualted HitTest value and its return value is ignored
'
' WM_SetCursor. The BlockValue will contain the HitTest code to use for the cursor
'   The return value of BlockValue is ignored.
'   This message is always overridden either by you or by the DLL


'///////////////////////////////////////////////////////////////////////////////////////
'   SUMMARY OF CHANGES/MODIFICATIONS FOR THE ENTIRE PROJECT
'///////////////////////////////////////////////////////////////////////////////////////
' Unresolved Bugs:
'  18 Sep 05 found by me.
'       Not really a bug. If your form has the ControlBox property set to false, the
'       DLL cannot override VB preventing closing the window when users clicks the
'       X button or selecting Close from the system menu. VB overrides the command.
'       I would have to call the DestroyWindow API to override VB & won't do that.
'       Additionally, setting the Form's Movable property to False, prevents the user
'       from moving the form by dragging its titlebar. VB overrides the Move command.
'   19 Sep 05 found by me.
'       Only when the window is minimized, right clicking on taskbar and selecting
'       the system menu items Restore or Maximize, you lose capability to simply toggle
'       minimizing/restoring the window by clicking on the taskbar. To get this
'       capability back, you must display the system menu by right clicking on the
'       taskbar button or titlebar whether or not you select a system menu item.
'       This is a bug in VB6 when a form is Owned (i.e., FormX.Show 0,Me). Still
'       tweaking to try to get past this bug, but am believing it is related to
'       changing window styles on the fly. If I cannot resolve the bug, it will
'       remain. Not toggling window styles is not an option -- it is the cleanest
'       way I can prevent VB from drawing the nonclient area when a menu pops up
'       or change in active state occurs. Truly annoying bug & has me somewhat peeved

' SUMMARY OF CHANGES (modifications and/or bug fixes)

' 20 Sep 05. Fixed bugs found 19/20 Sep - posted 20 Sep/2230
'   :: When providing tracking rectangle to user, the current rectangle was not copied to the passed DC
'   :: When updating owner-drawn flags, potential for the DLL to lose track of the flags
'   :: When DLL's client window was unsubclassed, potential for DLL to still fire callbacks
'   :: Changed the odImageLeftAlign value to 8
'       DLL could not know if no image was assigned to a menu item
'   :: Forced certain window styles on the window when it is first subclassed:
'       WS_SysMenu, WS_MaximizeBox, WS_MinimizeBox, WS_Caption, WS_ThickFrame, WS_DlgFrame
'       See unresolved bugs regarding the Close menu item when Form.ControlBox=False
'   :: Reorganized CustomWindow.ProcessMessage routine to only handle a few messages when
'       the window is minimized; all other messages are handled by O/S when minimized
'   :: cleaned up some comments, added some more, updated the QuickLook rtf file
' 21 Sep 05. Fixed bug created by last update :) - posted 21 Sep/1630
'   :: Used GetClipBox, expecting it to return the clip region without having to
'       create & use a GDI clipping region. Wrong! It returns much more than the
'       clipping region. Forced to use GDI region to test in FillBarImage and
'       GradientFill routines so that routines do not override a user supplied
'       clipping region. This is only possible because I exposed the cGraphics
'       class which can call those functions
'   :: Added CreateBitmapAndDC to cGraphics so user can have DLL create a DC
'       and/or Bitmap without having to do all the work themselves.
'   :: Win98 didn't stretch a pixel more than approx 200 times it size.
'       Modified the GraidentFill routine to step gradients @ 100 pixel increments
' 22 Sep 05: posted 23 Sep/1730
'   :: No longer using DIBs to rotate icons
'   :: Removed a few Rotated icon options & added a RotateAuto option
'        This new option will rotate the system icon in same orientation as titlebar text
'        Added the needed code to rotate an icon where image & mask are combined into a single bitmap
'   :: Neglected to add a Property Get for the rotate icon, it's there now
'   :: Added ModifyButtonTip to the clsButtons so user can change custom button tooltip text
'   :: Bug in PostionButton could cause Array Out of Bounds error
'   :: cleaned up a bunch of Public Declaration statements;
'        moved them to where they are used and now declared Private (organizational only)
' 23 Sep 05 - posted 23 Sep/2359
'   :: New ModifyButtonTip had potential to crash if adding a new tip to a button
'       that didn't already have a tip.
'   :: Modified GradientFill routine again! Stepping gradients by 100 was too much
'       for my Win98 to handle all the time. New stepping value is 20; hope that's it
'   :: minor tweak on system button vertical placement on a horizontal titlebar
'       when only system buttons shown (14 pixels tall), appears 1 pixel to high
'   :: Updated Rotate icon routine fixed issue with my Win98 displaying crappy
'       rotated icons. Now they are clear & crisp regardless of the orientation
' 24 Sep 05 - posted 25 Sep/0120
'   :: changed a few class Instancing propeties (were MultUse & didn't need to be)
'   :: created animation for a window being minimized to the system tray when the
'       MinimizeToSysTray option is set.
'   :: Added ShowWindowFromTrayIcon. Function animates a minimized window that was
'       minimized with MinimizeToSysTray option set to True, to one of
'       two states: Restored, Maximized
'   :: Modified code to use the new options & tweaked ShowOnTaskbar for a better,
'       cleaner toggle and also to be aware of the MinimizeToSysTray option
'   :: modified CalcNC to force minimum size on a window if the titlebar of that
'       window is toggled vertical<>horizontal and window is too small to display
'       the titlebar.
' 26 Sep 05 - posted 26 Sep/2205
'   :: Fixed bug when system menu "Size" selected & form would shrink to minimial size
'   :: Added ConvHandleToStdPic function in the cGraphics class
'   :: Major Mod: Added capability to export part or all of the custom settings,
'       background images and system tray data to a file <ExportCustomSettings>.
'       That generated file can be uploaded into a .RES file or used as is and
'       that file or the .RES file, via LoadResData() function, can be passed to
'       the DLL's <ImportCustomSettings> method to apply those settings. The
'       function has other options including adding a user-defined/supplied
'       property bag that can contain anything you want (i.e., menu images, etc)
'   :: Moved Frame.BackColor property into Active/Inactive classes vs the Frame class
'   :: Tidied up some code, updated the QuickLook RTF file
' 27 Sep 05 - not yet posted
'   :: Custom inset could draw over the right or bottom border
'   :: When user selects disabled menu item with keyboard, the menu loop did not exit
'   :: added option to exclude user buttons and/or tracking rectangles
'      when creating an export file. Tweaked the ImportCustomSettings routine a bit
' 3 & 4 Oct 05 - not yet posted
'   :: When space bar pressed on a sub menu when in menu loop, the loop would exit
'   :: added a bunch of gray scale options to active/inactive classes
'       Frame.Active/Inactive...  Set/GetImageBackground, Get/SetInsetImage and BackStyle property
'       Menubar.Active/Inactive... Set/GetImageBackground and BackStyle property
'       Titlebar.Active/Inactive... Set/GetImageBackground and BackStyle property
'   :: added 3 grayscale functions in the cGraphics class
'           GrayScale_Image, GrayScale_DC, GrayScale_Color
'   :: added grayscale optional parameters to following cGraphics functions
'           GradientFillEx and ImageFillEx
'   :: Finally figured out RedBird77's bug. Mouse driver was sending wm_initmenupopup
'       message to window. My routines assumed that message would be preceeded by a
'       wm_entermenuloop message and then ended with a wm_exitmenuloop message. This
'       is where the toggling of window styles took place. Needed to add the toggle
'       calls to the wm_initmenupopup also which is immediately where it is released.
'       The same calls remain in the wm_entermenuloop & wm_exitmenuloop routines.
'       This isn't an issue since the window style won't change if the change has been
'       already applied. Hopefully, this will catch any combination of those 3 messages
'       Bug was reported 20 Sep 05 & resolved with a lot of help from RedBird77
' 5 Oct - not yet posted
'   :: added functionality in the Export/ImportCustomSettings so that an image that is
'       used by both the Active/Inactive classes is not exported twice and imported twice.
'       The Active/Inactive class, in this case, will share the image; thereby reducing
'       resources (no longer double the images) when applying an exported settings file.
'   :: bug found by me. When multiple windows are being subclassed, the single-DC method
'       for drawing menubar item states could paint one window's menu item on another
'       window. SetMouseAction, UpdateMenuBar, DrawMenuItem tweaked to prevent bug.
'       - required caching the 1 byte window's active state in the clsMenubar routine
'   :: updated RTF file and also updated some examples to use grayscale options
' 6 Oct - posted 6 Oct 23:15
'   :: bug found. The AddButton routine was using an invalid reference into the
'       cButtons() array. This could cause an out of bounds. Fixed
'   :: added ToolTip capability to TrackingRectangles
'   :: finalizing the official .hlp file for this DLL
' 12 Oct - posted 12 Oct 20:40
'   :: bug found by Volte. SysTray class was crashing when the RemoveAll option was used.
'      Fixed. The internal collection of clsSysTray was not being reinitialized after set to Nothing
'   :: bug found by Richard Turnbull. Menus were not displaying with NT4
'      Fixed. Needed to call GetMenuItemInfo using &H10 (MIIM_Type) in the mask
' 22 Nov - posted 25 Nov
'   :: bug found by Eric P. When FormX.Load/Show was called from an active-x control on a
'       form that was being subclassed, and FormX was also being subclassed in the Load event,
'       FormX would not be accessible via the keyboard.
'   :: Fixed. Passing WM_NCActivate via CallWindowProc vs DefWindowProc. However that
'       allows VB to paint over the client area when the form has a menu and form
'       regains focus; no effect if form loses focus (new bug)
'       Fixed. By checking & removing the menu during the WM_NCPaint event.
'   :: bug found by me. During troubleshooting bugs above, noticed that if a form does not
'       have the focus but mouse is moved over a menu item & clicked, the menu item
'       does not react; a 2nd click is required.
'       Fixed: the menubar tracking hittest item was being reset unnecessarily.
'   :: bug found by me. Custom tooltips would not display when app is compiled
'       Fixed. Required a call to InitCommonControls at least once before a tooltip is displayed

