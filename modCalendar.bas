Attribute VB_Name = "modCalendar"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modCalendar"

'VGC 03/23/2017: CHANGES!

' ** DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 97 and Access2K
' **
' ** Copyright: Stephen Lebans - Lebans Holdings 1999 Ltd.
' **            Please feel free to use this code within your own
' **            projects whether they are private or commercial applications
' **            without obligation.
' **            This code may not be resold by itself or as part of a collection.
' **
' ** Name:      modCalendar
' **
' ** Version:   2.05
' **
' ** Purpose:
' **            Create a Window with a Window Procedure to house the
' **            Month Calendar control. Further provide a Menu interface
' **            to allow the user to modify the Calendar's properties.
' **            The Window procedure must reside in a standard Code module.
' **  
' ** Author:    Stephen Lebans
' **
' ** Email:     Stephen@lebans.com
' **
' ** Web Site:  www.lebans.com
' **
' ** Date:      Dec 09, 2004, 11:11:11 PM
' **
' ** Credits:   Based on code by:
' **            Ray Mercer - Window Creation & Messaging in VB
' **            Ken Getz & Michael Kaplan - AddrOf
' **            Charles Petzold - Window Creation and Message loops
' **            Dev Ashish - AddrOf implementation - Access Version checking
' **            Pedro Gil - Initial framework and props
' **            MSDN KB
' **
' ** BUGS:      Fixed the bug that appears as a result of  Access
' **            is caching the WinProc.
' **            Added a call to UnregisterClass to resolve the issue.
' **
' ** What's Missing:
' **            You tell Me.
' **
' **            Proper Error handling.
' **
' ** How it Works:
' **            The Month Calendar is created directly with the
' **            API's contained in the Common Controls DLL. In this manner we bypass
' **            the DatePicker ActiveX control, which is simply a wrapper for these
' **            calls anyway. This removes any problems from distribution and
' **            especially version issues of using the ActiveX control.
' **
' ** This is the 10th major release.

' ** To exit from the Window Procedure,
' ** thereby closing the MonthCalendar Control,
' ** you can either:
' **   1) Press the Escape Key
' **   2) Click on the Window's Close Button(x)
' **   3) Double Click or Single Click the Left Mouse Button on a Date
' **      It depends on your settings for the Calendar Properties

' ****************************************************
'               WARNING
' If you place a Breakpoint within the Window Procedure
' you will cause a GPF!
' ****************************************************

Private Type WNDCLASSEX
  cbSize As Long
  style As Long
  lpfnWndProc As Long
  cbClsExtra As Long
  cbWndExtra As Long
  hInstance As Long
  hIcon As Long
  hCursor As Long
  hbrBackground As Long
  lpszMenuName As String
  lpszClassName As String
  hIconSm As Long
End Type

Private Type PAINTSTRUCT
  hDC As Long
  fErase As Long
  rcPaint As RECT
  fRestore As Long
  fIncUpdate As Long
  rgbReserved(32) As Byte
End Type

' ** Bit-packed array of "bold" info for a month.
' ** If a bit is on, that day is drawn bold.
Private Type MONTHDAYSTATE
  lpMONTHDAYSTATE As Long
  ' ** Should really be array of 4 bytes because
  ' ** of VB's Signed datatypes.
End Type

' ** Control Message Header.
Private Type NMHDR
  hwndFrom As Long
  idfrom As Long
  Code As Long  'Integer
End Type

' ** MonthCalendar SelectChange.
Private Type NMSELCHANGE
  nm As NMHDR
  stSelStart As SYSTEMTIME
  stSelEnd As SYSTEMTIME
End Type

' ** DayState Header.
Private Type NMDAYSTATE
  nmhd As NMHDR  ' ** This must be first, so we don't break WM_NOTIFY.
  stStart As SYSTEMTIME
  cDayState As Long ' ** For ease of use always specify 12 months of data.
  prgDayState As Long  'MONTHDAYSTATE ' ** points to cDayState MONTHDAYSTATEs.
End Type

Private Type MCHITTESTINFO
  cbSize As Long
  pt As POINTAPI
  uHit As Long
  st As SYSTEMTIME
End Type

'**********************************************************************************************
' ** VB6 RUNTIMES must be present to resolve this call.
' ** Returns address of the address of the associated SafeArray descriptor.
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
'**********************************************************************************************

Private Declare Function GetDoubleClickTime Lib "user32.dll" () As Long

Private Declare Function GetMessageTime Lib "user32.dll" () As Long

Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" _
  (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Function InsertMenu Lib "user32.dll" Alias "InsertMenuA" _
  (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long

Private Declare Function CreateMenu Lib "user32.dll" () As Long

Private Declare Function CheckMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" _
  (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetMenu Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function LoadIcon Lib "user32.dll" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long

Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long

Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long

Private Declare Function RegisterClassEx Lib "user32.dll" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer

Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function BeginPaint Lib "user32.dll" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long

Private Declare Function EndPaint Lib "user32.dll" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" _
  (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Function UnregisterClass Lib "user32.dll" Alias "UnregisterClassA" _
  (ByVal lpClassname As String, ByVal hInstance As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long

Private Declare Function MessageBeep Lib "user32.dll" Alias "BeepA" (ByVal wType As Long) As Long

Private Declare Function Beep Lib "kernel32.dll" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

' ** Enable/Disable Main Access Window.
Private Declare Function EnableWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Private Declare Function IsWindowEnabled Lib "user32.dll" (ByVal hwnd As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hWndLock As Long) As Long

Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long

Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

' ** Button Control Styles.
'Private Const BS_PUSHBUTTON      As Integer = &H0&
'Private Const BS_DEFPUSHBUTTON   As Integer = &H1&
'Private Const BS_CHECKBOX        As Integer = &H2&
'Private Const BS_AUTOCHECKBOX    As Integer = &H3&
'Private Const BS_RADIOBUTTON     As Integer = &H4&
'Private Const BS_3STATE          As Integer = &H5&
'Private Const BS_AUTO3STATE      As Integer = &H6&
'Private Const BS_GROUPBOX        As Integer = &H7&
'Private Const BS_USERBUTTON      As Integer = &H8&
'Private Const BS_AUTORADIOBUTTON As Integer = &H9&
'Private Const BS_OWNERDRAW       As Integer = &HB&
'Private Const BS_LEFTTEXT        As Integer = &H20&

' ** User Button Notification Codes.
'Private Const BN_CLICKED       As Integer = 0
'Private Const BN_PAINT         As Integer = 1
'Private Const BN_HILITE        As Integer = 2
'Private Const BN_UNHILITE      As Integer = 3
'Private Const BN_DISABLE       As Integer = 4
'Private Const BN_DOUBLECLICKED As Integer = 5

' ** Button Control Messages.
'Private Const BM_GETCHECK As Integer = &HF0
'Private Const BM_SETCHECK As Integer = &HF1
'Private Const BM_GETSTATE As Integer = &HF2
'Private Const BM_SETSTATE As Integer = &HF3
'Private Const BM_SETSTYLE As Integer = &HF4

Private Const WS_VISIBLE           As Double = &H10000000
'Private Const WS_VSCROLL           As Double = &H200000
'Private Const WS_TABSTOP           As Double = &H10000
Private Const WS_THICKFRAME        As Double = &H40000
'Private Const WS_MAXIMIZE          As Double = &H1000000
Private Const WS_MAXIMIZEBOX       As Double = &H10000
Private Const WS_MINIMIZE          As Double = &H20000000
Private Const WS_MINIMIZEBOX       As Double = &H20000
Private Const WS_SYSMENU           As Double = &H80000
Private Const WS_BORDER            As Double = &H800000
Private Const WS_CAPTION           As Double = &HC00000
'Private Const WS_CHILD             As Double = &H40000000
'Private Const WS_CHILDWINDOW       As Double = (WS_CHILD)
'Private Const WS_CLIPCHILDREN      As Double = &H2000000
'Private Const WS_CLIPSIBLINGS      As Double = &H4000000
Private Const WS_DISABLED          As Double = &H8000000
'Private Const WS_DLGFRAME          As Double = &H400000
'Private Const WS_EX_ACCEPTFILES    As Double = &H10&
Private Const WS_EX_DLGMODALFRAME  As Double = &H1&
'Private Const WS_EX_NOPARENTNOTIFY As Double = &H4&
Private Const WS_EX_TOPMOST        As Double = &H8&
'Private Const WS_EX_TRANSPARENT    As Double = &H20&
'Private Const WS_GROUP             As Double = &H20000
'Private Const WS_HSCROLL           As Double = &H100000
'Private Const WS_ICONIC            As Double = WS_MINIMIZE
Private Const WS_OVERLAPPED        As Double = &H0&
Private Const WS_OVERLAPPEDWINDOW  As Double = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_POPUP             As Double = &H80000000
Private Const WS_POPUPWINDOW       As Double = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
'Private Const WS_SIZEBOX           As Double = WS_THICKFRAME
'Private Const WS_TILED             As Double = WS_OVERLAPPED
'Private Const WS_TILEDWINDOW       As Double = WS_OVERLAPPEDWINDOW

Private Const CW_USEDEFAULT As Long = &H80000000

Private Const CS_HREDRAW As Long = &H2
Private Const CS_VREDRAW As Long = &H1

Private Const IDI_APPLICATION As Long = 32512&

Private Const IDC_ARROW As Long = 32512&

Private Const WHITE_BRUSH As Integer = 0
'Private Const BLACK_BRUSH As Integer = 4

' ** CONSTANTS.
'Private Const WM_KEYFIRST      As Long = &H100
Private Const WM_KEYDOWN       As Long = &H100
Private Const WM_KEYUP         As Long = &H101
'Private Const WM_CHAR          As Long = &H102
'Private Const WM_DEADCHAR      As Long = &H103
'Private Const WM_SYSKEYDOWN    As Long = &H104
'Private Const WM_SYSKEYUP      As Long = &H105
'Private Const WM_SYSCHAR       As Long = &H106
Private Const WM_CLOSE         As Long = &H10
Private Const WM_DESTROY       As Long = &H2
Private Const WM_PAINT         As Long = &HF
Private Const WM_NOTIFY        As Long = &H4E
Private Const WM_PARENTNOTIFY  As Long = &H210
'Private Const WM_SETTEXT       As Long = &HC
'Private Const WM_INITMENU      As Long = &H116
'Private Const WM_INITMENUPOPUP As Long = &H117
'Private Const WM_MENUSELECT    As Long = &H11F
'Private Const WM_MENUCHAR      As Long = &H120
'Private Const WM_ENTERIDLE     As Long = &H121
'Private Const WM_MOUSEFIRST    As Long = &H200  ' ** Window Message.
'Private Const WM_MOUSEMOVE     As Long = &H200
Private Const WM_LBUTTONDOWN   As Long = &H201
'Private Const WM_LBUTTONUP     As Long = &H202
'Private Const WM_LBUTTONDBLCLK As Long = &H203
'Private Const WM_RBUTTONDOWN   As Long = &H204
'Private Const WM_RBUTTONUP     As Long = &H205
'Private Const WM_RBUTTONDBLCLK As Long = &H206
'Private Const WM_MBUTTONDOWN   As Long = &H207
'Private Const WM_MBUTTONUP     As Long = &H208
'Private Const WM_MBUTTONDBLCLK As Long = &H209
'Private Const WM_MOUSELAST     As Long = &H209
'Private Const WM_SETFOCUS      As Long = &H7
'Private Const WM_KILLFOCUS     As Long = &H8
Private Const WM_MOVE          As Long = &H3
'Private Const WM_SIZE          As Long = &H5
'Private Const WM_ENABLE        As Long = &HA
'Private Const WM_SETREDRAW     As Long = &HB
Private Const WM_COMMAND       As Long = &H111

' ** Virtual Keys, Standard Set.
'Private Const VK_LBUTTON As Integer = &H1
'Private Const VK_RBUTTON As Integer = &H2
'Private Const VK_CANCEL  As Integer = &H3
'Private Const VK_MBUTTON As Integer = &H4  ' ** NOT contiguous with L RBUTTON.
'Private Const VK_BACK    As Integer = &H8
'Private Const VK_TAB     As Integer = &H9
'Private Const VK_CLEAR   As Integer = &HC
Private Const VK_RETURN  As Integer = &HD
Private Const VK_SHIFT   As Integer = &H10
'Private Const VK_CONTROL As Integer = &H11
'Private Const VK_MENU    As Integer = &H12
'Private Const VK_PAUSE   As Integer = &H13
'Private Const VK_CAPITAL As Integer = &H14
Private Const VK_ESCAPE  As Integer = &H1B
'Private Const VK_SPACE   As Integer = &H20
'Private Const VK_PRIOR   As Integer = &H21
'Private Const VK_NEXT    As Integer = &H22
Private Const VK_END     As Integer = &H23
Private Const VK_HOME    As Integer = &H24
Private Const VK_LEFT    As Integer = &H25
Private Const VK_UP      As Integer = &H26
Private Const VK_RIGHT   As Integer = &H27
Private Const VK_DOWN    As Integer = &H28

'Private Const MB_ICONHAND        As Integer = &H10&
'Private Const MB_ICONQUESTION    As Integer = &H20&
'Private Const MB_ICONEXCLAMATION As Integer = &H30&
'Private Const MB_ICONASTERISK    As Integer = &H40&
'Private Const MB_ICONINFORMATION As Integer = MB_ICONASTERISK
'Private Const MB_ICONSTOP        As Integer = MB_ICONHAND

Private Const MF_ENABLED         As Integer = &H0&
Private Const MF_UNCHECKED       As Integer = &H0&
Private Const MF_CHECKED         As Integer = &H8&
'Private Const MF_USECHECKBITMAPS As Integer = &H200&
'Private Const MF_MENUBARBREAK    As Integer = &H20&
'Private Const MF_MENUBREAK       As Integer = &H40&
'Private Const MF_SEPARATOR       As Integer = &H800&
Private Const MF_BYPOSITION      As Integer = &H400&
Private Const MF_POPUP           As Integer = &H10&
Private Const MF_STRING          As Integer = &H0&

'Private Const NM_FIRST           As Long = 0  ' ** Generic to all controls.
'Private Const NM_LAST            As Long = -99
'Private Const NM_RELEASEDCAPTURE As Long = (NM_FIRST - 16)
'Private Const NM_KEYDOWN         As Long = (NM_FIRST - 15)
'Private Const NM_DBLCLK          As Long = (NM_FIRST - 3)

'Private Const DTN_FIRST As Long = -760
'Private Const DTN_LAST  As Long = -799

Private Const MCN_FIRST       As Long = -750
'Private Const MCN_LAST        As Long = -799
Private Const MCN_GETDAYSTATE As Long = (MCN_FIRST + 3)
Private Const MCN_SELECT      As Long = (MCN_FIRST + 4)
Private Const MCN_SELCHANGE   As Long = (MCN_FIRST + 1)

' ** Color part's of the Calendar.
Private Const MCSC_BACKGROUND   As Integer = 0  ' ** The background color (between months).
Private Const MCSC_TEXT         As Integer = 1  ' ** The dates.
Private Const MCSC_TITLEBK      As Integer = 2  ' ** Background of the title.
Private Const MCSC_TITLETEXT    As Integer = 3
Private Const MCSC_MONTHBK      As Integer = 4  ' ** Background within the month cal.
Private Const MCSC_TRAILINGTEXT As Integer = 5  ' ** The text color of header & trailing days.

Private Const MCM_FIRST   As Long = &H1000&
Private Const MCM_HITTEST As Long = (MCM_FIRST + &HE)

Private Const MCHT_TITLE            As Long = &H10000
Private Const MCHT_CALENDAR         As Long = &H20000
'Private Const MCHT_TODAYLINK        As Long = &H30000
Private Const MCHT_NEXT             As Long = &H1000000  ' ** These indicate that hitting
Private Const MCHT_PREV             As Long = &H2000000  ' ** here will go to the next/prev month.
'Private Const MCHT_NOWHERE          As Long = &H0
'Private Const MCHT_TITLEBK          As Long = (MCHT_TITLE)
'Private Const MCHT_TITLEMONTH       As Long = (MCHT_TITLE Or &H1)
'Private Const MCHT_TITLEYEAR        As Long = (MCHT_TITLE Or &H2)
'Private Const MCHT_TITLEBTNNEXT     As Long = (MCHT_TITLE Or MCHT_NEXT Or &H3)
'Private Const MCHT_TITLEBTNPREV     As Long = (MCHT_TITLE Or MCHT_PREV Or &H3)
'Private Const MCHT_CALENDARBK       As Long = (MCHT_CALENDAR)
Private Const MCHT_CALENDARDATE     As Long = (MCHT_CALENDAR Or &H1)
'Private Const MCHT_CALENDARDATENEXT As Long = (MCHT_CALENDARDATE Or MCHT_NEXT)
'Private Const MCHT_CALENDARDATEPREV As Long = (MCHT_CALENDARDATE Or MCHT_PREV)
'Private Const MCHT_CALENDARDAY      As Long = (MCHT_CALENDAR Or &H2)
'Private Const MCHT_CALENDARWEEKNUM  As Long = (MCHT_CALENDAR Or &H3)

' ** We'll translate above color indexes into Menu ID's
' ** by adding 1000 to the values.

' ** MISC Properties.

' ** Use Single Or Double Click to Select a Date.
'Private Const SingleOrDouble As Long = 720
Private Const SingleClick    As Long = 721
Private Const DoubleClick    As Long = 722

' ** Show Week Numbers.
'Private Const ShowWeekNum    As Long = 700
Private Const ShowWeekNumYES As Long = 701
Private Const ShowWeekNumNO  As Long = 702

' ** Show Today TodayNumbers.
'Private Const ShowToday    As Long = 705
Private Const ShowTodayYES As Long = 706
Private Const ShowTodayNO  As Long = 707

' ** Show CircleToday Numbers.
'Private Const ShowcircleToday    As Long = 708
Private Const ShowCircleTodayYES As Long = 709
Private Const ShowCircleTodayNO  As Long = 710

' ** Font Dialog Menu.
Private Const FontDialog As Long = 820

' ** Font Size Menu.
'Private Const Fontx5 As Long = 805

' ** Weeks Menu.
Private Const Monthx1  As Long = 901
Private Const Monthx2  As Long = 902
Private Const Monthx3  As Long = 903
Private Const Monthx4  As Long = 904
Private Const Monthx6  As Long = 906
Private Const Monthx8  As Long = 908
Private Const Monthx9  As Long = 909
Private Const Monthx12 As Long = 912

' ** WindowPosition menu.
Private Const Positionx0 As Long = 920
Private Const Positionx1 As Long = 921
Private Const Positionx2 As Long = 922
Private Const Positionx3 As Long = 923
Private Const Positionx4 As Long = 924
'Private Const Positionx5 As Long = 925
'Private Const Positionx6 As Long = 926
'Private Const Positionx7 As Long = 927
Private Const Positionx8 As Long = 928

Private Const CLASSNAME As String = "MonthCalendar"
Private Const Title As String = "Month Calendar"

' ** Variables to store our dynamic menu's item IDs.
'Private Menu1 As Long
'Private Menu2 As Long
'Private Menu3 As Long
'Private Menu4 As Long
'Private Menu5 As Long
'Private Menu6 As Long
'Private Menu7 As Long

' ** Junk Vars.
Private lngRetVal As Long
Private lngTmp01 As Long

' ** Module level var to hold handle to our Calendars hWnd.
'Private hWndCalendar As Long

' ** Module level var to hold reference to our Calendar object.
' ** We need this to access the Class from the WindowProc function.
Private clsMC As clsMonthCal

' ** Module level variable to hold the currently selected date.
Private SelectedDate As Date

' ** Module level variables to hold local copy of
' ** the currently selected Starting and Ending date Ranges.
Private localStartSelectedDate As Date
Private localEndSelectedDate As Date

' ** Module Var to track whether a Font or Color
' ** Dialog window is currently Open.
Private blnDialogOpen As Boolean

' ** Required to be Module level in order for WindowProc to have access to Menu handles.
Private hMenu As Long
Private hMenuPop As Long
Private hMenuPopMisc As Long
Private hMenuPopMiscShowWeekNumbers As Long
Private hMenuPopMiscFont  As Long
Private hMenuPopMiscColor As Long
Private hMenuPopMiscToday  As Long
Private hMenuPopMiscCircleToday  As Long
Private hMenuPopMiscWindowPosition  As Long
Private hMenuPopMiscOneClick  As Long

' ** To allow for Keyboard selection of Date(s).
Private SelChangeDateStart As Date
Private SelChangeDateEnd As Date
' **

Public Function GetFuncPtr(ByVal lngFnPtr As Long) As Long
' ** Wrapper function to allow AddressOf to be used within VB.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "GetFuncPtr"

110     GetFuncPtr = lngFnPtr

EXITP:
120     Exit Function

ERRH:
130     Select Case ERR.Number
        Case Else
140       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
150     End Select
160     Resume EXITP

End Function

Private Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long

200   On Error GoTo ERRH

        Const THIS_PROC As String = "MakeDWord"

210     MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)

EXITP:
220     Exit Function

ERRH:
230     Select Case ERR.Number
        Case Else
240       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
250     End Select
260     Resume EXITP

End Function

Public Function ShowMonthCalendar(ByRef clsMonthClass As clsMonthCal, ByRef datStartDate As Date, Optional ByRef datEndDate As Date = 0) As Boolean
' ************************************************************
' ** March 22, 2004
' ** Major modification to the function logic including calling Parameter order.
' ** Changed function to return Boolean FALSE and "datStartDate = 0"
' ** if user did not select a date from the MonthCalendar.
' ** The hWndForm param is no longer optional.
' ************************************************************
' ********* WARNING *************
' ** In order for this function to return Focus to the calling Form properly
' ** you must set the MonthCalendar class's hWndForm property BEFORE
' ** Calling this function!!!!!!!!!!
' *******************************
' **
' ** This function will always return the date selected by the
' ** user in the MonthCalendar as the return value for this function.
' ** If this function is called with the optional Date Range variables,
' ** then it will also return the starting and ending dates of the
' ** range of dates selected by the user.
' ** Finally if datStartDate and datEndDate <> 0, then their
' ** values will be used initialize the Calendar.

300   On Error GoTo ERRH  ' ** This will be Canceled, below.

        Const THIS_PROC As String = "ShowMonthCalendar"

        Dim typPT As POINTAPI
        Dim typWC As WNDCLASSEX
        Dim typMSG As MSG
        Dim lngHWnd As Long
        Dim lngClassAtom As Long  ' ** Class Atom.
        Dim lngHInstance As Long
        Dim blnFormIsPopup As Boolean
        Dim blnAppWindowIsModal As Boolean
        Dim lngEXStyle As Long
        Dim strTmp01 As String, lngTmp02 As Long
        Dim lngX As Long

        Const mcCLASSNAME As String = "MonthCalendar"
        Const mcTITLE As String = "Month Calendar"

310   On Error Resume Next

        ' ** Make sure the instance of MonthCalendar class is valid!
320     If clsMonthClass Is Nothing Then
330       strTmp01 = " The MonthCalendar class instance you passed to this function is INVALID!" & vbCrLf
340       strTmp01 = strTmp01 & " You must instantiate the MonthCalendar Class object before you call this function." & vbCrLf
350       strTmp01 = strTmp01 & " The code behind the sample Form shows you how to do this in the Form's Load event." & vbCrLf & vbCrLf
360       strTmp01 = strTmp01 & "' This must appear here!" & vbCrLf
370       strTmp01 = strTmp01 & "' Create an instance of our Class." & vbCrLf
380       strTmp01 = strTmp01 & "Private Sub Form_Load()" & vbCrLf
390       strTmp01 = strTmp01 & "Set clsMC = New clsMonthCal" & vbCrLf
400       strTmp01 = strTmp01 & "' You must set the class hWndForm prop!!!" & vbCrLf
410       strTmp01 = strTmp01 & "clsMC.hWndForm = Me.hWnd"
420       MsgBox strTmp01, vbInformation + vbOKOnly, "Invalid MonthCalendar Object!"
          ' ** Return nothing!
430       ShowMonthCalendar = 0
440       Exit Function
450     End If

        ' ** If this window already exists, then exit!
460     lngRetVal = FindWindow(mcCLASSNAME, mcTITLE)
470     If lngRetVal <> 0 Then
          ' ** strTmp01 = "The Calendar Window Already Exists!" & vbCrLf
          ' ** strTmp01 = strTmp01 & "Please Close and then Restart Access!"
          ' ** MsgBox strTmp01, vbCritical, "Critical Error. The MonthCalendar Window already exists"
          ' **  Return nothing!
          ' ** ShowMonthCalendar = 0
          ' **  We can just Return. The user has tried to open another instance of the Calendar.
          ' **  Up to this point, Version 98b, we only support one open instance at a time.
480       ShowMonthCalendar = 0
490       Exit Function
500     End If

        ' ** Create a local copy of the MonthCalendar class.
510     Set clsMC = clsMonthClass

        ' ** Update our init cursor props.
520     lngRetVal = GetCursorPos(typPT)
530     clsMC.CursorXinit = typPT.X
540     clsMC.CursorYinit = typPT.Y

        ' ** Ensure our SelChange vars are reset.
550     SelChangeDateStart = 0
560     SelChangeDateEnd = 0

        ' ** MENU creation tiMe.
570     hMenu = CreateMenu
580     hMenuPop = CreatePopupMenu
590     hMenuPopMisc = CreatePopupMenu
600     hMenuPopMiscShowWeekNumbers = CreatePopupMenu
610     hMenuPopMiscFont = CreatePopupMenu
620     hMenuPopMiscColor = CreatePopupMenu
630     hMenuPopMiscToday = CreatePopupMenu
640     hMenuPopMiscCircleToday = CreatePopupMenu
650     hMenuPopMiscWindowPosition = CreatePopupMenu
660     hMenuPopMiscOneClick = CreatePopupMenu

        ' ** Viewable Months Menu.
670     lngRetVal = InsertMenu(hMenuPopMisc, 1&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPop, "Viewable Months")
        ' ** Viewable Months SubMenus.
680     lngRetVal = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx1, "1 Month")
690     lngRetVal = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx2, "2 Months")
700     lngRetVal = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx3, "3 Months")
710     lngRetVal = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx4, "4 Months")
720     lngRetVal = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx6, "6 Months")
730     lngRetVal = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx8, "8 Months")
740     lngRetVal = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx9, "9 Months")
750     lngRetVal = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx12, "12 Months")

        ' ** Erase all check marks.
760     For lngX = 0 To 7
770       lngRetVal = CheckMenuItem(hMenuPop, 0, MF_UNCHECKED Or MF_BYPOSITION)
780     Next lngX

        ' ** Now set the Menu Check for the current number of months displayed.
790     lngTmp02 = (clsMC.MonthColumns * clsMC.MonthRows)
800     Select Case lngTmp02
        Case 1
810       lngX = 7
820     Case 2
830       lngX = 6
840     Case 3
850       lngX = 5
860     Case 4
870       lngX = 4
880     Case 6
890       lngX = 3
900     Case 8
910       lngX = 2
920     Case 9
930       lngX = 1
940     Case 12
950       lngX = 0
960     End Select

        ' ** VGC 06/11/2008: Removed menus as extraneous and confusing to user; per Rich.

        ' ** Now set the Menu Check.
970     lngRetVal = CheckMenuItem(hMenuPop, lngX, MF_CHECKED Or MF_BYPOSITION)

        ' ** Misc Properties Menu.
        'lngRetVal = InsertMenu(hMenu, 2&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMisc, "Properties")

        ' ** Let's add Top level Menu Item that does not contain any submen items.
        ' ** We will use it like a CommandButton to allow the users to Close the Calendar Window.
        'lngRetVal = InsertMenu(hMenu, 1&, MF_BYPOSITION Or MF_ENABLED, 998, "Close Window")

        ' ** Show WeekNumbers SubMenu
        'lngRetVal = InsertMenu(hMenuPopMisc, 1&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscShowWeekNumbers, "ShowWeek#'s")
        'lngRetVal = InsertMenu(hMenuPopMiscShowWeekNumbers, 0&, MF_STRING Or MF_BYPOSITION, ShowWeekNumYES, "YES")
        'lngRetVal = InsertMenu(hMenuPopMiscShowWeekNumbers, 0&, MF_STRING Or MF_BYPOSITION, ShowWeekNumNO, "NO")
        'If clsMC.ShowWeekNumbers = False Then
        '  lngRetVal = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 0, MF_CHECKED Or MF_BYPOSITION)
        '  lngRetVal = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 1, MF_UNCHECKED Or MF_BYPOSITION)
        'Else
        '  lngRetVal = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 1, MF_CHECKED Or MF_BYPOSITION)
        '  lngRetVal = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 0, MF_UNCHECKED Or MF_BYPOSITION)
        'End If

        ' ** Font stuff SubMenu
        'lngRetVal = InsertMenu(hMenuPopMisc, 2&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscFont, "Font")
        'lngRetVal = InsertMenu(hMenuPopMiscFont, 0&, MF_STRING Or MF_BYPOSITION, FontDialog, "Select Font")

        ' ** Color Props SubMenu
        'lngRetVal = InsertMenu(hMenuPopMisc, 3&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscColor, "Colors")
        'lngRetVal = InsertMenu(hMenuPopMiscColor, 0&, MF_STRING Or MF_BYPOSITION, MCSC_MONTHBK + 1000, "BackGround Color")
        'lngRetVal = InsertMenu(hMenuPopMiscColor, 0&, MF_STRING Or MF_BYPOSITION, MCSC_BACKGROUND + 1000, "Frame Color")
        'lngRetVal = InsertMenu(hMenuPopMiscColor, 0&, MF_STRING Or MF_BYPOSITION, MCSC_TEXT + 1000, "Dates Color")
        'lngRetVal = InsertMenu(hMenuPopMiscColor, 0&, MF_STRING Or MF_BYPOSITION, MCSC_TITLEBK + 1000, "Title BG Color")
        'lngRetVal = InsertMenu(hMenuPopMiscColor, 0&, MF_STRING Or MF_BYPOSITION, MCSC_TITLETEXT + 1000, "Title Text Color")
        'lngRetVal = InsertMenu(hMenuPopMiscColor, 0&, MF_STRING Or MF_BYPOSITION, MCSC_TRAILINGTEXT + 1000, "Trailing Text Color")
        'lngRetVal = InsertMenu(hMenuPopMiscColor, 0&, MF_STRING Or MF_BYPOSITION, MCSC_TRAILINGTEXT + 2000, "Reset All Colors")

        ' ** Show Today's Date
980     lngRetVal = InsertMenu(hMenuPopMisc, 4&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscToday, "Show Today")
990     lngRetVal = InsertMenu(hMenuPopMiscToday, 0&, MF_STRING Or MF_BYPOSITION, ShowTodayYES, "YES")
1000    lngRetVal = InsertMenu(hMenuPopMiscToday, 0&, MF_STRING Or MF_BYPOSITION, ShowTodayNO, "NO")

1010    If clsMC.NoToday = True Then
1020      lngRetVal = CheckMenuItem(hMenuPopMiscToday, 0, MF_CHECKED Or MF_BYPOSITION)
1030      lngRetVal = CheckMenuItem(hMenuPopMiscToday, 1, MF_UNCHECKED Or MF_BYPOSITION)
1040    Else
1050      lngRetVal = CheckMenuItem(hMenuPopMiscToday, 1, MF_CHECKED Or MF_BYPOSITION)
1060      lngRetVal = CheckMenuItem(hMenuPopMiscToday, 0, MF_UNCHECKED Or MF_BYPOSITION)
1070    End If

        ' ** Circle Today's Date
1080    lngRetVal = InsertMenu(hMenuPopMisc, 5&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscCircleToday, "Circle Today")
1090    lngRetVal = InsertMenu(hMenuPopMiscCircleToday, 0&, MF_STRING Or MF_BYPOSITION, ShowCircleTodayYES, "YES")
1100    lngRetVal = InsertMenu(hMenuPopMiscCircleToday, 0&, MF_STRING Or MF_BYPOSITION, ShowCircleTodayNO, "NO")
1110    If clsMC.NoTodayCircle = True Then
1120      lngRetVal = CheckMenuItem(hMenuPopMiscCircleToday, 0, MF_CHECKED Or MF_BYPOSITION)
1130      lngRetVal = CheckMenuItem(hMenuPopMiscCircleToday, 1, MF_UNCHECKED Or MF_BYPOSITION)
1140    Else
1150      lngRetVal = CheckMenuItem(hMenuPopMiscCircleToday, 1, MF_CHECKED Or MF_BYPOSITION)
1160      lngRetVal = CheckMenuItem(hMenuPopMiscCircleToday, 0, MF_UNCHECKED Or MF_BYPOSITION)
1170    End If

        ' ** Window Position
1180    lngRetVal = InsertMenu(hMenuPopMisc, 6&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscWindowPosition, "Calendar Location")
1190    lngRetVal = InsertMenu(hMenuPopMiscWindowPosition, 0&, MF_STRING Or MF_BYPOSITION, Positionx0, "Cursor Location when Calendar Opened")
1200    lngRetVal = InsertMenu(hMenuPopMiscWindowPosition, 0&, MF_STRING Or MF_BYPOSITION, Positionx1, "Where User Last Dragged")
1210    lngRetVal = InsertMenu(hMenuPopMiscWindowPosition, 0&, MF_STRING Or MF_BYPOSITION, Positionx2, "Center of Access App Window")
1220    lngRetVal = InsertMenu(hMenuPopMiscWindowPosition, 0&, MF_STRING Or MF_BYPOSITION, Positionx3, "Center of Screen")
1230    lngRetVal = InsertMenu(hMenuPopMiscWindowPosition, 0&, MF_STRING Or MF_BYPOSITION, Positionx4, "Top Left Corner")

1240    For lngX = 0 To 4
1250      lngRetVal = CheckMenuItem(hMenuPopMiscWindowPosition, lngX, MF_UNCHECKED Or MF_BYPOSITION)
1260    Next lngX

        ' ** Now set the Menu Check the current number of months displayed
1270    lngTmp02 = (clsMC.WindowLocation)
        ' ** Now set the Menu Check
1280    lngRetVal = CheckMenuItem(hMenuPopMiscWindowPosition, 4 - lngTmp02, MF_CHECKED Or MF_BYPOSITION)

        ' ** Single or Double Click to select Date
1290    lngRetVal = InsertMenu(hMenuPopMisc, 7&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscOneClick, "Single Or Double Click")
1300    lngRetVal = InsertMenu(hMenuPopMiscOneClick, 0&, MF_STRING Or MF_BYPOSITION, DoubleClick, "Double Click to Select Date")
1310    lngRetVal = InsertMenu(hMenuPopMiscOneClick, 0&, MF_STRING Or MF_BYPOSITION, SingleClick, "Single Click to Select Date")

1320    If clsMC.OneClick = True Then
1330      lngRetVal = CheckMenuItem(hMenuPopMiscOneClick, 0, MF_CHECKED Or MF_BYPOSITION)
1340      lngRetVal = CheckMenuItem(hMenuPopMiscOneClick, 1, MF_UNCHECKED Or MF_BYPOSITION)
1350    Else
1360      lngRetVal = CheckMenuItem(hMenuPopMiscOneClick, 1, MF_CHECKED Or MF_BYPOSITION)
1370      lngRetVal = CheckMenuItem(hMenuPopMiscOneClick, 0, MF_UNCHECKED Or MF_BYPOSITION)
1380    End If

        ' ** Get instance of this App
1390    lngHInstance = GetWindowLong(Application.hWndAccessApp, GWL_HINSTANCE)  ' ** API Function: modWindowFunctions.

        ' ** From code by Ray Mercer,
        ' ** Set up and register window class.
1400    typWC.cbSize = Len(typWC)
1410    typWC.style = CS_HREDRAW Or CS_VREDRAW

        ' ** Determine Access Version.
        ' *****************************
        ' ** For A97 MUST USE AddrOf.
        ' *****************************
        ' ** If Val(SysCmd(acSysCmdAccessVer)) < 8 Then
        ' ** typWC.lpfnWndProc = AddrOf("WindowProc")
        ' ** Else
1420    typWC.lpfnWndProc = GetFuncPtr(AddressOf WindowProc)
        ' ** End If

1430    typWC.cbClsExtra = 0&
1440    typWC.cbWndExtra = 0&
1450    typWC.hInstance = lngHInstance
1460    typWC.hIcon = LoadIcon(lngHInstance, IDI_APPLICATION)
1470    typWC.hCursor = LoadCursor(lngHInstance, IDC_ARROW)
1480    typWC.hbrBackground = GetStockObject(WHITE_BRUSH)
1490    typWC.lpszMenuName = 0&
1500    typWC.lpszClassName = CLASSNAME
1510    typWC.hIconSm = LoadIcon(lngHInstance, IDI_APPLICATION)

        ' ** Register this Class.
1520    lngClassAtom = RegisterClassEx(typWC)

        ' ** Force window to always stay on top.
1530    lngEXStyle = WS_EX_DLGMODALFRAME  ' ** April 6 trying to fix WIn98 Form in Popup view Or WS_EX_TOPMOST.

        ' ** Create a window Set to be NOT VISIBLE TO START Or WS_VISIBLE.
1540    lngHWnd = CreateWindowEx(lngEXStyle, CLASSNAME, Title, WS_POPUPWINDOW Or WS_CAPTION, CW_USEDEFAULT, _
          CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, clsMC.hWndForm, hMenu, lngHInstance, 0&)

        ' ** We have to allow for the following:
        ' ** 1) The calling Form's Modal prop is turned on.
1550    lngRetVal = GetWindowLong(Application.hWndAccessApp, GWL_STYLE)
1560    blnAppWindowIsModal = lngRetVal And WS_DISABLED

        ' ** 2) The calling Form's Popup prop is turned on.
1570    lngRetVal = GetWindowLong(clsMC.hWndForm, GWL_STYLE)
1580    blnFormIsPopup = lngRetVal And WS_POPUP

        ' ** We will actually create our MonthCal window by setting the
        ' ** Class hWnd property.
        ' ** Set the Control's Parent Window property.
1590    clsMC.hwnd = lngHWnd

        ' ** Init the Calendar to the date(s) supplied by the
        ' ** user in the calling function.
1600    If datStartDate <> 0 And datEndDate <> 0 Then
1610      clsMC.SetSelectedDateRange datStartDate, datEndDate
          ' ** Update our local copies of these vars.
          ' ** Need to redo the logic to get rid of these local vars.
          ' ** See the date select code in the WindProc.
1620      localStartSelectedDate = datStartDate
1630      localEndSelectedDate = datEndDate
          ' ** Clear our Return Date local Var.
1640      SelectedDate = 0
1650    Else
1660      If datStartDate <> 0 Then
1670        clsMC.SelectedDate = datStartDate
            ' ** Clear our Return Date local Var.
1680        SelectedDate = 0
1690      Else
1700        SelectedDate = 0
1710        localStartSelectedDate = 0
1720        localEndSelectedDate = 0
1730      End If
1740    End If

        ' ** The following logic is required to ensure our MonthCalendar window
        ' ** is MODAL (the user can only click in this window).
        ' ** If parent form's Popup prop is turned on then
        ' ** we have to Disable this Form ourselves.
1750    If blnFormIsPopup Then lngRetVal = EnableWindow(clsMC.hWndForm, 0)

        ' ** We only want to Disable the main app window if
        ' ** the Form's Modal prop is not true.
        ' ** Check and see if the main Access app window
        ' ** is disabled already - if not then disable it.
1760    If Not blnAppWindowIsModal Then
1770      lngRetVal = EnableWindow(Application.hWndAccessApp, 0)
1780    End If

        ' ** Show the Calendar's Parent window first then the MonthCal window.
1790    ShowWindow lngHWnd, SW_SHOWNORMAL  ' ** API Function: modWindowFunctions.
1800    ShowWindow clsMC.hWndCal, SW_SHOWNORMAL  ' ** API Function: modWindowFunctions.

        ' ** Enter message loop.
        ' ** (all window messages are handled in WindowProc().)
1810    Do While 0 <> GetMessage(typMSG, 0&, 0&, 0&)
1820      TranslateMessage typMSG
1830      DispatchMessage typMSG
1840    Loop

        ' ** User has closed the MonthCalendar window.
        ' ** Return the Selected Date.
        ' ** If the user has called this function with the optional
        ' ** date range vars, then fill them in.
1850    If SelectedDate <> 0 Then
          ' ** The Calendar Window is closed so we cannot
          ' ** use our Class methods that use SendMessage
          ' ** to get their current values.
1860      datStartDate = SelectedDate
1870      datEndDate = localEndSelectedDate
1880      ShowMonthCalendar = True
1890    Else
          ' ** User did not SELECT a Date.
1900      datStartDate = 0
1910      datEndDate = 0
1920      ShowMonthCalendar = False
1930    End If

        ' ** Unregister our Custom Window Class.
        ' ** If you don't then you will GPF on the next init of the class.
1940    lngRetVal = UnregisterClass(CLASSNAME, lngHInstance)

        ' ** If Form was Popup then Enable this window first.
1950    If blnFormIsPopup Then
1960      lngRetVal = EnableWindow(clsMC.hWndForm, 1)
1970    End If

        ' ** In order to prevent screen flashing upon closing
        ' ** our MonthCalendar window, we have to enable the
        ' ** main Access application window in the MonthCalendar's
        ' ** WinProc's WM_CLOSE message handler. From here now though,
        ' ** we have to Disable the main Access application window
        ' ** if the calling form's Modal prop was turned on.
1980    If blnAppWindowIsModal Then
          '** Disable Access App window.
1990      lngRetVal = EnableWindow(Application.hWndAccessApp, 0)
2000    End If

        ' ** Release Class reference required to be visible to our WindProc.
2010    Set clsMC = Nothing

        ' ** Ensure focus returns to calling form.
2020    SetFocus clsMonthClass.hWndForm

EXITP:
2030    Exit Function

ERRH:
2040    Select Case ERR.Number  ' ** This error handler will never be hit, but leave it in!
        Case Else
2050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2060    End Select
2070    Resume EXITP

End Function

Public Function WindowProc(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' ** Main message handler for the MonthCalendar window
' *** WARNING ***
' ** DO NOT PLACE DEBUG BREAKPOINTS IN THIS FUNCTION
' *** WARNING ***

2100  On Error GoTo ERRH  ' ** Will be Canceled, below.

        Const THIS_PROC As String = "WindowProc"

        ' ** Temp Window Handle for Dialogs
        'Dim lngHWndTmp As Long  NOT USED!
        ' ** To hold local copy of the current Message
        'Dim typMSG As MSG  NOT USED!
        'Dim typST As SYSTEMTIME  NOT USED!
        'Dim lngPtrArray As Long  NOT USED!
        'Dim arr_typST(0 To 1) As SYSTEMTIME  NOT USED!

        ' ** There is a bug or I am having alignment problems
        ' ** so we pass the second element of this array
        ' ** and leave the first zero'd out.
        Dim arr_typMDS(-1 To 13) As MONTHDAYSTATE
        Dim ps As PAINTSTRUCT
        'Dim typRECT As RECT  NOT USED!
        Dim nmsc As NMSELCHANGE
        Dim hdr As NMHDR
        Dim nmds As NMDAYSTATE
        Dim datStartDate As Date
        Dim datEndDate As Date
        Dim lngHdc As Long
        Dim lngCurMessagetime  As Long
        Dim lngWMmessage As Long
        Dim lngDoubleClickTime As Long
        Dim intStartMonth As Integer
        Dim intCurrentMonth As Integer
        Dim intCurrentYear As Integer
        Dim intTmp01 As Integer
        Dim intX As Integer
        Dim blnRetVal As Boolean

        ' ** Flag to make sure we have a MouseUp bewtween our
        ' ** MouseDown messages to signify a Double Click
        ' ** not just the Mouse Button held down
        Static blnMouseUp As Boolean
        Static lngLastMouseDown As Long

        ' ** You cannot have unhandled errors in a WinProc
        ' ** so we will just ingnore them all!! <grin>
        ' ** Really though, this is very heavily debugged code!
2110  On Error Resume Next

2120    Select Case Message

        Case WM_MOVE
          ' ** Update the MonthCalendar's current
2130      Call UpdateCursor(lParam, hwnd)

2140    Case WM_PAINT
          ' ** Must leave this in to ensure Window is Redrawn!!!
2150      lngHdc = BeginPaint(hwnd, ps)
2160      Call EndPaint(hwnd, ps)
2170      Exit Function

2180    Case WM_KEYDOWN, WM_KEYUP

          ' ** Select case on the Virtual Key Code
2190      Select Case wParam

          Case VK_ESCAPE
2200        Call PostMessage(hwnd, WM_CLOSE, 0, 0)
2210        Exit Function

2220      Case VK_SHIFT, VK_LEFT, VK_RIGHT, VK_DOWN, VK_UP, VK_HOME, VK_END, vbKeyPageDown, vbKeyPageUp
2230        KeysToMonthCal hwnd, Message, wParam, lParam
2240        Exit Function

2250      Case VK_RETURN
            ' ** If the SelChangeDateStart var != 0 then send our MCN_SELECT Message
2260        If SelChangeDateStart = 0 Then Exit Function

2270        If SelChangeDateEnd = SelChangeDateStart Then
2280          clsMC.SelectedDate = SelChangeDateStart
2290        Else
2300          clsMC.SetSelectedDateRange SelChangeDateStart, SelChangeDateEnd
2310        End If

            ' ** Update our local var
2320        SetSelectedDate SelChangeDateStart
            ' ** Update our Class starting and ending date range vars
2330        UpdateRangeVars SelChangeDateStart, SelChangeDateEnd

            ' ** Let's Close the Calendar
2340        Call PostMessage(hwnd, WM_CLOSE, 0, 0)
            'Debug
            ''debug.print "Used Enter key to select date!"
2350        Exit Function

2360      Case Else
2370        WindowProc = DefWindowProc(hwnd, Message, wParam, lParam)

2380        Exit Function
2390      End Select

2400    Case WM_CLOSE
          ' ** April 12, 2004
          ' ** FINALLY resolved issue of screen flickering with Win2K or higher!!
          ' ** We have to temporarily Enable the main Access application window
2410      lngRetVal = EnableWindow(Application.hWndAccessApp, 1)
2420      lngRetVal = ShowWindow(Application.hWndAccessApp, SW_SHOW)  ' ** API Function: modWindowFunctions.

2430      WindowProc = DefWindowProc(hwnd, Message, wParam, lParam)
2440      Exit Function

2450    Case WM_DESTROY
2460      PostQuitMessage 0&
2470      Exit Function

2480    Case WM_PARENTNOTIFY
          ' ** Grab the lower WORD
2490      lngWMmessage = (wParam And &HFFFF)
          ' ** Switch on Window Message
2500      Select Case lngWMmessage

          Case WM_LBUTTONDOWN

            ' ** Mod Nov 24 -2002
            ' ** Removed MouseButton logic to determine when to close
            ' ** calendar. Now we simply check it from the SELECT notification
            ' ** and close the window if CHeckOneClick property is TRUE.
            ' ** We do not use the DoubleCLick logic either.
            ' ** Get the current Double Click interval
2510        lngCurMessagetime = GetMessageTime
2520        lngDoubleClickTime = GetDoubleClickTime

            ' ** Make sure the Cursor is double clicking
            ' ** on an actual Date not on a Calendar control
2530        blnRetVal = LocationCursorOnCalendar(lParam)
2540        If Not blnRetVal Then
              ' ** Call the default WIndow proc
2550          WindowProc = DefWindowProc(hwnd, Message, wParam, lParam)
2560          Exit Function
2570        End If

            ' Debug. A2K closing date range on one click!
2580        If Abs((lngCurMessagetime - lngLastMouseDown)) < lngDoubleClickTime Then ' Or CheckOneClick = True Then
              ' ** Double CLicked-or CheckOneClick-Let's CLose the Calendar
2590          Call PostMessage(hwnd, WM_CLOSE, 0, 0)
2600          lngLastMouseDown = 0
2610          blnMouseUp = False
2620          Exit Function
2630        End If

            ' ** Always update our last left mouse button pressed var
2640        lngLastMouseDown = lngCurMessagetime

2650      Case Else
            ' ** Call the default Window proc
2660        WindowProc = DefWindowProc(hwnd, Message, wParam, lParam)
2670        Exit Function

            ' ** All Done!
2680      End Select

2690    Case WM_NOTIFY
          ' ** Update our class startdate, and range date props.
          ' ** Copy the NMRH structure to our local copy
2700      CopyMemory hdr, ByVal lParam, Len(hdr)  ' ** API Function: modWindowFunctions.

          ' ** Modified Nov 24 -2002
          ' ** SELECT is when the user explicitly clicks to select a date.
          ' ** SELCHANGE is when the user scrolls through the calendar automatically
          ' ** updating the selected date.
          ' ** Thanks to Blake Sell for catching this!

2710      Select Case hdr.Code

          Case MCN_SELECT
            ' ** This needs to be fixed up to have seperate routines
            ' ** for single vs range date selections.
            ' ** Drop local vars and use the MonthCalendar Class only
            ' ** Grab the struct info
2720        CopyMemory nmsc, ByVal lParam, Len(nmsc)  ' ** API Function: modWindowFunctions.

            ' ** Convert to our Date format
2730        With nmsc.stSelStart '(0)
2740          datStartDate = DateSerial(.wYear, .wMonth, .wDay)
2750        End With
2760        With nmsc.stSelEnd '(1)
2770          datEndDate = DateSerial(.wYear, .wMonth, .wDay)
2780        End With

            ' ** Update our local var
2790        SetSelectedDate datStartDate
            ' ** Update our Class starting and ending date range vars
2800        UpdateRangeVars datStartDate, datEndDate

            ' ** Modified Nov 24 -2002
            ' ** Removed MouseButton logic to determine when to close
            ' ** calendar. Now we simply check it from the SELECT notification
            ' ** and close the window if CHeckOneClick property is TRUE.
2810        If clsMC.OneClick = True Then
              ' ** Double CLicked-or CheckOneClick-Let's CLose the Calendar
2820          Call PostMessage(hwnd, WM_CLOSE, 0, 0)
2830          lngLastMouseDown = 0
2840          blnMouseUp = False
              'Exit Function
2850        End If

2860        Exit Function

            ' ** June 2 - 2004 - adding support for ENTER key to select currently highlighted date.
2870      Case MCN_SELCHANGE

            ' ** Grab the struct info
2880        CopyMemory nmsc, ByVal lParam, Len(nmsc)  ' ** API Function: modWindowFunctions.

            ' ** Convert to our Date format
2890        With nmsc.stSelStart '(0)
2900          SelChangeDateStart = DateSerial(.wYear, .wMonth, .wDay)
2910        End With
2920        With nmsc.stSelEnd '(1)
2930          SelChangeDateEnd = DateSerial(.wYear, .wMonth, .wDay)
2940        End With
            ' debug.print "DateStart:" & DateStart

2950      Case MCN_GETDAYSTATE
2960        For intX = -1 To UBound(arr_typMDS)
2970          arr_typMDS(intX).lpMONTHDAYSTATE = 0
2980        Next

2990        CopyMemory nmds, ByVal lParam, Len(nmds)  ' ** API Function: modWindowFunctions.
3000        intTmp01 = nmds.cDayState
            ''debug.print "Months requested:" & intTmp01
            ''debug.print time

            ' ** Have to allow for the fact that the month before and
            ' ** the month after are always requested. THis means the starting year
            ' ** can be one year before the year of the first fully displayed month.
3010        intStartMonth = nmds.stStart.wMonth
3020        intCurrentYear = nmds.stStart.wYear

3030        intCurrentMonth = intStartMonth '+ intX
3040        For intX = 0 To intTmp01 - 1

3050          If intCurrentMonth > 12 Then
3060            intCurrentMonth = intCurrentMonth - 12 '1
3070            intCurrentYear = intCurrentYear + 1
3080          End If
3090          arr_typMDS(intX).lpMONTHDAYSTATE = clsMC.GetDAYSTATE(intCurrentYear, intCurrentMonth)
3100          intCurrentMonth = intCurrentMonth + 1
3110        Next intX
            ' ** set the address of our array
3120        lngTmp01 = VarPtr(arr_typMDS(0))
3130        CopyMemory ByVal lParam + (Len(nmds) - 4), lngTmp01, 4  ' ** API Function: modWindowFunctions.

            ' ** Signal we want this message to be processed
3140        WindowProc = 0
3150        Exit Function

3160      Case Else
3170        WindowProc = DefWindowProc(hwnd, Message, wParam, lParam)

3180      End Select

3190    Case WM_COMMAND:
          ' ** WM_COMMAND is sent to the window
          ' ** whenever someone clicks a menu.
          ' ** The menu's item ID is stored in wParam.

3200      Select Case wParam
          Case Monthx1 To Monthx12
            'Call MsgBox("You clicked Dynamic Sub Menu 1!", vbExclamation)
3210        SetMonths (CInt(wParam) - 900)
3220        Exit Function

3230      Case ShowWeekNumYES
3240        ShowWeekNums True
3250        Exit Function

3260      Case ShowWeekNumNO
3270        ShowWeekNums False
3280        Exit Function

3290      Case FontDialog
3300        ShowFontDialog

3310        Exit Function

3320      Case MCSC_BACKGROUND + 1000
3330        SelectColor MCSC_BACKGROUND
3340        Exit Function

3350      Case MCSC_MONTHBK + 1000
3360        SelectColor MCSC_MONTHBK
3370        Exit Function

3380      Case MCSC_TEXT + 1000
3390        SelectColor MCSC_TEXT
3400        Exit Function

3410      Case MCSC_TITLEBK + 1000
3420        SelectColor MCSC_TITLEBK
3430        Exit Function

3440      Case MCSC_TITLETEXT + 1000
3450        SelectColor MCSC_TITLETEXT
3460        Exit Function

3470      Case MCSC_TRAILINGTEXT + 1000
3480        SelectColor MCSC_TRAILINGTEXT
3490        Exit Function

3500      Case MCSC_TRAILINGTEXT + 2000
3510        ResetColors
3520        Exit Function

            ' ** Show Today's Date at bottom of Calendar
3530      Case ShowTodayYES
3540        sShowToday False
3550        Exit Function

3560      Case ShowTodayNO
3570        sShowToday True
3580        Exit Function

            ' ** Circle Today's Date
3590      Case ShowCircleTodayYES
3600        sShowcircleToday False
3610        Exit Function

3620      Case ShowCircleTodayNO
3630        sShowcircleToday True
3640        Exit Function

            ' ** WindowPosition menu
3650      Case Positionx0 To Positionx8
3660        sWindowPosition wParam, hwnd

3670      Case SingleClick
3680        sClick True
3690        Exit Function

3700      Case DoubleClick
3710        sClick False
3720        Exit Function

3730      Case 998
3740        Call PostMessage(hwnd, WM_CLOSE, 0, 0)
3750        lngLastMouseDown = 0
3760        blnMouseUp = False
3770        Exit Function

3780      Case Else

            ' ** Call the Default Window Procedure for all other WM_COMMAND'
3790        WindowProc = DefWindowProc(hwnd, Message, wParam, lParam)
3800        Exit Function
3810      End Select

3820    Case Else
          ' ** pass all other messages to default window procedure
3830      WindowProc = DefWindowProc(hwnd, Message, wParam, lParam)

3840    End Select

EXITP:
3850    Exit Function

ERRH:
3860    Select Case ERR.Number  ' ** This error handler will never be hit, but leave it in!
        Case Else
3870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3880    End Select
3890    Resume EXITP

End Function

Private Function HiWord(ByVal DWord As Long) As Integer

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "HiWord"

3910    HiWord = (DWord And &HFFFF0000) \ &H10000

EXITP:
3920    Exit Function

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Function

Private Sub KeysToMonthCal(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "KeysToMonthCal"

4010    Call PostMessage(ByVal clsMC.hWndCal, ByVal MSG, ByVal wParam, ByVal lParam)

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Function LocationCursorOnCalendar(ByVal lParam As Long) As Boolean

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "LocationCursorOnCalendar"

        Dim ht As MCHITTESTINFO

        ' ** The x-coordinate of the cursor is the low-order word,
        ' ** and the y-coordinate of the cursor is the high-order word.
4110    ht.pt.X = LoWord(lParam)
4120    ht.pt.Y = HiWord(lParam)

        ' ** Set structure size.
4130    ht.cbSize = Len(ht)
4140    lngRetVal = SendMessage(ByVal clsMC.hWndCal, ByVal MCM_HITTEST, ByVal 0&, ht)  ' ** API Function: modWindowFunctions.
4150    If ht.uHit <> MCHT_CALENDARDATE Then
4160      LocationCursorOnCalendar = False
4170    Else
4180      LocationCursorOnCalendar = True
4190    End If

EXITP:
4200    Exit Function

ERRH:
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Function

Private Function LoWord(ByVal DWord As Long) As Integer

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "LoWord"

4310    If DWord And &H8000& Then  '&H8000& = &H00008000
4320      LoWord = DWord Or &HFFFF0000
4330    Else
4340      LoWord = DWord And &HFFFF&
4350    End If

EXITP:
4360    Exit Function

ERRH:
4370    Select Case ERR.Number
        Case Else
4380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4390    End Select
4400    Resume EXITP

End Function

Private Function ReleaseClass() As Boolean

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "ReleaseClass"

4510    Set clsMC = Nothing

EXITP:
4520    Exit Function

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Function

Private Sub ResetColors()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "ResetColors"

4610    clsMC.ResetCalendarColors

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub sClick(bl As Boolean)
' ** Sets the Class's OneClick property and the
' ** appropriate Menu Check Marks.

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "sClick"

4710    If bl Then
4720      clsMC.OneClick = True
4730      lngRetVal = CheckMenuItem(hMenuPopMiscOneClick, 0, MF_CHECKED Or MF_BYPOSITION)
4740      lngRetVal = CheckMenuItem(hMenuPopMiscOneClick, 1, MF_UNCHECKED Or MF_BYPOSITION)
4750    Else
4760      clsMC.OneClick = False
4770      lngRetVal = CheckMenuItem(hMenuPopMiscOneClick, 1, MF_CHECKED Or MF_BYPOSITION)
4780      lngRetVal = CheckMenuItem(hMenuPopMiscOneClick, 0, MF_UNCHECKED Or MF_BYPOSITION)
4790    End If

EXITP:
4800    Exit Sub

ERRH:
4810    Select Case ERR.Number
        Case Else
4820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4830    End Select
4840    Resume EXITP

End Sub

Private Sub sDayState()
' ** Pass Dummy value for now.

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "sDayState"

4910    clsMC.DAYSTATE = 0

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub SelectColor(ByVal index As Long)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "SelectColor"

5010    blnDialogOpen = True
5020    clsMC.ChooseColors index
5030    blnDialogOpen = False

EXITP:
5040    Exit Sub

ERRH:
5050    Select Case ERR.Number
        Case Else
5060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5070    End Select
5080    Resume EXITP

End Sub

Private Function SetMonths(ByVal mth As Integer) As Boolean

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "SetMonths"

        Dim lngX As Long

5110    clsMC.SetViewableMonths mth

        ' ** 7 Possible/Total Menu Items to uncheck.
5120    For lngX = 0 To 7
5130      lngRetVal = CheckMenuItem(hMenuPop, lngX, MF_UNCHECKED Or MF_BYPOSITION)
5140    Next lngX

        ' ** Now set the Menu Check the current number of months displayed.
5150    lngTmp01 = (clsMC.MonthColumns * clsMC.MonthRows)
5160    Select Case lngTmp01
        Case 1
5170      lngX = 7
5180    Case 2
5190      lngX = 6
5200    Case 3
5210      lngX = 5
5220    Case 4
5230      lngX = 4
5240    Case 6
5250      lngX = 3
5260    Case 8
5270      lngX = 2
5280    Case 9
5290      lngX = 1
5300    Case 12
5310      lngX = 0
5320    End Select

        ' ** Now set the Menu Check.
5330    lngRetVal = CheckMenuItem(hMenuPop, lngX, MF_CHECKED Or MF_BYPOSITION)

EXITP:
5340    Exit Function

ERRH:
5350    Select Case ERR.Number
        Case Else
5360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5370    End Select
5380    Resume EXITP

End Function

Private Function SetSelectedDate(ByVal dt As Date) As Boolean

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "SetSelectedDate"

5410    SelectedDate = dt

EXITP:
5420    Exit Function

ERRH:
5430    Select Case ERR.Number
        Case Else
5440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5450    End Select
5460    Resume EXITP

End Function

Private Sub ShowFontDialog()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowFontDialog"

5510    blnDialogOpen = True
5520    clsMC.SelectFont
5530    blnDialogOpen = False

EXITP:
5540    Exit Sub

ERRH:
5550    Select Case ERR.Number
        Case Else
5560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5570    End Select
5580    Resume EXITP

End Sub

Private Function ShowWeekNums(ByVal blnShow As Boolean) As Boolean

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowWeekNums"

5610    If blnShow = True Then
5620      clsMC.ShowWeekNumbers = True
5630      lngRetVal = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 1, MF_CHECKED Or MF_BYPOSITION)
5640      lngRetVal = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 0, MF_UNCHECKED Or MF_BYPOSITION)
5650    Else
5660      clsMC.ShowWeekNumbers = False
5670      lngRetVal = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 0, MF_CHECKED Or MF_BYPOSITION)
5680      lngRetVal = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 1, MF_UNCHECKED Or MF_BYPOSITION)
5690    End If

EXITP:
5700    Exit Function

ERRH:
5710    Select Case ERR.Number
        Case Else
5720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5730    End Select
5740    Resume EXITP

End Function

Private Sub sShowcircleToday(blnShow As Boolean)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "sShowcircleToday"

5810    If blnShow = True Then
5820      clsMC.NoTodayCircle = True
5830      lngRetVal = CheckMenuItem(hMenuPopMiscCircleToday, 0, MF_CHECKED Or MF_BYPOSITION)
5840      lngRetVal = CheckMenuItem(hMenuPopMiscCircleToday, 1, MF_UNCHECKED Or MF_BYPOSITION)
5850    Else
5860      lngRetVal = CheckMenuItem(hMenuPopMiscCircleToday, 1, MF_CHECKED Or MF_BYPOSITION)
5870      lngRetVal = CheckMenuItem(hMenuPopMiscCircleToday, 0, MF_UNCHECKED Or MF_BYPOSITION)
5880      clsMC.NoTodayCircle = False
5890    End If

EXITP:
5900    Exit Sub

ERRH:
5910    Select Case ERR.Number
        Case Else
5920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5930    End Select
5940    Resume EXITP

End Sub

Private Sub sShowToday(blnShow As Boolean)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "sShowToday"

6010    If blnShow Then
6020      clsMC.NoToday = True
6030      lngRetVal = CheckMenuItem(hMenuPopMiscToday, 0, MF_CHECKED Or MF_BYPOSITION)
6040      lngRetVal = CheckMenuItem(hMenuPopMiscToday, 1, MF_UNCHECKED Or MF_BYPOSITION)
6050    Else
6060      clsMC.NoToday = False
6070      lngRetVal = CheckMenuItem(hMenuPopMiscToday, 1, MF_CHECKED Or MF_BYPOSITION)
6080      lngRetVal = CheckMenuItem(hMenuPopMiscToday, 0, MF_UNCHECKED Or MF_BYPOSITION)
6090    End If

EXITP:
6100    Exit Sub

ERRH:
6110    Select Case ERR.Number
        Case Else
6120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6130    End Select
6140    Resume EXITP

End Sub

Private Sub sWindowPosition(lngWParam As Long, lngHWnd As Long)
' ** Position Window according to users Menu selections:
' ** a) 0 -Pop at cursor location when user activates Calendar.
' ** b) 1 -Where they manually move/leave it at.
' ** c) 2 -Centered in Access App Window.
' ** d) 3 -Centered on entire screen.
' ** d) 4 -Top Left Corner.

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "sWindowPosition"

        Dim typRECT As RECT
        'Dim typPT As POINTAPI  NOT USED!
        Dim lngX As Long

6210    Select Case lngWParam

        Case Positionx0
          ' ** Pop at Cursor.
6220      clsMC.PositionAtCursor = True

6230    Case Positionx1
6240      clsMC.PositionAtCursor = False
          ' ** Use current position of Calendar Window.
          ' ** Get rectangle for our Form.
          'Debug.Print "GetWindowRect- Me.hWnd:" & m_Form.hWnd
6250      lngRetVal = GetWindowRect(lngHWnd, typRECT)  ' ** API Function: modWindowFunctions.

6260      clsMC.CursorX = typRECT.Left  'typPT.x 'typRECT.Left
6270      clsMC.CursorY = typRECT.Top  'typPT.y

6280    Case Positionx2 To Positionx8
6290      clsMC.PositionAtCursor = False

6300    Case Else

6310    End Select

        ' ** Update Window Position property.
6320    clsMC.WindowLocation = lngWParam - 920
        'Debug.Print "modCalendar - clsMC.Windowlocation:" & lngWParam ' clsMC.WindowLocation
6330    For lngX = 0 To 4
6340      lngRetVal = CheckMenuItem(hMenuPopMiscWindowPosition, lngX, MF_UNCHECKED Or MF_BYPOSITION)
6350    Next lngX
        ' ** Now set the Menu Check the current number of months displayed.
6360    lngTmp01 = (clsMC.WindowLocation)
        ' ** Now set the Menu Check.
6370    lngRetVal = CheckMenuItem(hMenuPopMiscWindowPosition, 4 - lngTmp01, MF_CHECKED Or MF_BYPOSITION)

6380    clsMC.ReDraw

EXITP:
6390    Exit Sub

ERRH:
6400    Select Case ERR.Number
        Case Else
6410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6420    End Select
6430    Resume EXITP

End Sub

Private Sub UpdateCursor(ByVal lngParam As Long, ByVal lngHWnd As Long)
' ** xPos = (int)(short) LOWORD(lngParam);  ' ** Horizontal position.
' ** yPos = (int)(short) HIWORD(lngParam);  ' ** Vertical position.

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateCursor"

        Dim typRECT As RECT
        Dim typPT As POINTAPI

        ' ** Should not happen.
6510    If clsMC.hwnd = 0 Then Exit Sub
        ' ** Only update if the window is visible.
6520    lngRetVal = GetWindowLong(lngHWnd, GWL_STYLE)
6530    If Not (lngRetVal And WS_VISIBLE) Then Exit Sub

        ' ** If PositionAtCursor is True then
        ' ** DO NOT UPDATE!
6540    If clsMC.PositionAtCursor Then Exit Sub

6550    lngRetVal = GetWindowRect(lngHWnd, typRECT)  ' ** API Function: modWindowFunctions.
6560    typPT.X = typRECT.Left
6570    typPT.Y = typRECT.Top

        ''debug.print time & "  UpdateCursor -X:" & typRECT.Left & "  Y:" & typRECT.Top
6580    clsMC.CursorX = typPT.X
6590    clsMC.CursorY = typPT.Y

        ' ** UpdateCursor -X:" & clsMC.CursorX & "  Y:" & clsMC.CursorY

EXITP:
6600    Exit Sub

ERRH:
6610    Select Case ERR.Number
        Case Else
6620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6630    End Select
6640    Resume EXITP

End Sub

Private Sub UpdateRangeVars(ByVal datStartDate As Date, ByVal datEndDate As Date)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateRangeVars"

6710    localStartSelectedDate = datStartDate
6720    localEndSelectedDate = datEndDate

EXITP:
6730    Exit Sub

ERRH:
6740    Select Case ERR.Number
        Case Else
6750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6760    End Select
6770    Resume EXITP

End Sub
