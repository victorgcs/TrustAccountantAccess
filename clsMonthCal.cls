VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMonthCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "clsMonthCal"

'VGC 03/26/2017: CHANGES!

' ** DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 97 VBA
' **  There is a seperate A2K or higher version required because
' **  A97 does not support AddressOf.
' **  That is the only difference between the A97 and the A2K versions.

' ** Copyright: Stephen Lebans - Lebans Holdings 1999 Ltd.
' **            Pedro Gil
' **            Please feel free to use this code within your own
' **            projects whether they are private or commercial applications
' **            without obligation.
' **            This code may not be resold by itself or as part of a collection.

' ** Name:      clsMonthCal

' ** Version:   2.05

' ** Purpose:
' **            To allow for the selection of a Date or Dates using the standard
' **            Windows Month Calendar control without having to use the ActiveX control.
' ** а
' ** Authors:   Stephen Lebans
' **            Pedro Gil

' ** Email:     Stephen@lebans.com
' **            pmpg98@hotmail.com

' ** Web Site:  www.lebans.com
' **            http://www.geocities.com/pmpg98_pt/

' ** Date:      December 09, 2004, 1:11:11 PM

' ** Credits:   Pedro Gil for doing the background
' **            research necessary to code a working model.
' **            Stephen Lebans for the API non ActiveX solution,
' **            with autosizing, user selectable display.
' **            Michael Kaplan for the idea to create the
' **            Month Calendar directly from the Common Control DLL.
' **            jwolzvb@yahoo.de for a VB example of implementing
' **            a 1 month ActiveX MonthCalendar control.

' ** BUGS:      Please report any bugs to Stephen@Lebans.com

' ** What's Missing:
' **            Proper error handling.

' ** How it Works:
' **            The Month Calendar is created directly with the
' **            API's contained in the Common Controls DLL. In this manner we bypass
' **            the DatePicker ActiveX control, which is simply a wrapper for these
' **            calls anyway. This removes any problems from distribution and
' **            especially version issues of using the ActiveX control.

' **            DAYSTATE:
' **            Have a look at the Initialize sub for this Class to see how to
' **            set the DayStates for the months. You can also set the bits
' **            in the array yourself if you are converting data that you
' **            are storing in your own tables.

' ** This is the ?th release. We have exposed most of the properties of the
' ** Month Calendar. Simply add your own for any we have left out.

' ** Window Style Bits.
'Private Const WS_OVERLAPPED   As Double = &H0&
'Private Const WS_POPUP        As Double = &H80000000
Private Const WS_CHILD        As Double = &H40000000
'Private Const WS_MINIMIZE     As Double = &H20000000
Private Const WS_VISIBLE      As Double = &H10000000
'Private Const WS_DISABLED     As Double = &H8000000
'Private Const WS_CLIPSIBLINGS As Double = &H4000000
'Private Const WS_CLIPCHILDREN As Double = &H2000000
'Private Const WS_MAXIMIZE     As Double = &H1000000
'Private Const WS_CAPTION      As Double = &HC00000
Private Const WS_BORDER       As Double = &H800000
'Private Const WS_DLGFRAME     As Double = &H400000
'Private Const WS_VSCROLL      As Double = &H200000
'Private Const WS_HSCROLL      As Double = &H100000
'Private Const WS_SYSMENU      As Double = &H80000
'Private Const WS_THICKFRAME   As Double = &H40000

' ** Windows Style Bits Specific to DTP.
'Private Const DTS_UPDOWN          As Long = &H1   '0x0001  ' ** Use UPDOWN instead of MONTHCAL.
'Private Const DTS_SHOWNONE        As Long = &H2   '0x0002  ' ** Allow a NONE selection.
'Private Const DTS_SHORTDATEFORMAT As Long = &H0   '0x0000  ' ** Use the short date format (app must forward WM_WININICHANGE messages).
'Private Const DTS_LONGDATEFORMAT  As Long = &H4   '0x0004  ' ** Use the long date format (app must forward WM_WININICHANGE messages).
'Private Const DTS_TIMEFORMAT      As Long = &H9   '0x0009  ' ** Use the time format (app must forward WM_WININICHANGE messages).
'Private Const DTS_APPCANPARSE     As Long = &H10  '0x0010  ' ** Allow user entered strings (app MUST respond to DTN_USERSTRING).
'Private Const DTS_RIGHTALIGN      As Long = &H20  '0x0020  ' ** Right-align popup instead of left-align it.

' ** Used to change Font for Month Calendar common control.
Private Const WM_SETFONT As Long = &H30

' ** Bit-packed array of "bold" info for a month.
' ** If a bit is on, that day is drawn bold.
Private Type MONTHDAYSTATE
  lpMONTHDAYSTATE As Long
End Type

Private Type NMHDR
  hwndFrom As Long
  idfrom As Long
  Code As Long  'Integer
End Type

'Private Type NMDAYSTATE
'  nmhd As NMHDR  ' ** This must be first, so we don't break WM_NOTIFY.
'  stStart As SYSTEMTIME
'  cDayState As Long
'  ' ** For ease of use always specify 12 months of data.
'  prgDayState(11) As MONTHDAYSTATE  ' ** Points to cDayState MONTHDAYSTATEs.
'End Type

Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" _
  (ByVal h As Long, ByVal W As Long, ByVal e As Long, ByVal o As Long, ByVal W As Long, ByVal i As Long, ByVal U As Long, _
  ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal F As String) As Long

Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

' ** To register the class.
Private Type INITCOMMONCONTROLSEXSTRUCT
  dwsize As Long
  dwICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As INITCOMMONCONTROLSEXSTRUCT) As Long

Private Const ICC_DATE_CLASSES As Integer = &H100

Private Declare Function AdjustWindowRect Lib "user32.dll" (lpRectl As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long

' ** This one isn't used.
Private Declare Function AdjustWindowRectEx Lib "user32.dll" _
  (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

Private Declare Function CreateIC Lib "gdi32.dll" Alias "CreateICA" _
  (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long

Private Declare Function Beep Lib "kernel32.dll" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

' ** SysMonthCal32 messages.
Private Const MCM_FIRST             As Long = &H1000&
Private Const MCM_GETCURSEL         As Long = &H1001&   '(MCM_FIRST + 1)
Private Const MCM_SETCURSEL         As Long = (MCM_FIRST + &H2)
Private Const MCM_GETMINREQRECT     As Long = (MCM_FIRST + &H9)
Private Const MCM_SETMAXSELCOUNT    As Long = (MCM_FIRST + &H4)
'Private Const MCM_GETRANGE          As Long = (MCM_FIRST + &H11)
'Private Const MCM_SETRANGE          As Long = (MCM_FIRST + &H12)
Private Const MCM_GETSELRANGE       As Long = (MCM_FIRST + &H5)
'Private Const MCM_GETMONTHRANGE     As Long = (MCM_FIRST + &H7)
Private Const MCM_SETCOLOR          As Long = (MCM_FIRST + &HA)
Private Const MCM_GETCOLOR          As Long = (MCM_FIRST + &HB)
'Private Const MCM_SETTODAY          As Long = (MCM_FIRST + &HC)
'Private Const MCM_GETTODAY          As Long = (MCM_FIRST + &HD)
'Private Const MCM_SETFIRSTDAYOFWEEK As Long = (MCM_FIRST + &HF)
'Private Const MCM_GETFIRSTDAYOFWEEK As Long = (MCM_FIRST + &H10)
'Private Const MCM_SETDAYSTATE       As Long = (MCM_FIRST + &H8)
Private Const MCM_SETSELRANGE       As Long = (MCM_FIRST + &H6)
Private Const MCM_GETMAXTODAYWIDTH  As Long = (MCM_FIRST + &H15)

' ** Color parts of the Calendar.
Private Const MCSC_BACKGROUND   As Integer = 0  ' ** The background color (between months).
Private Const MCSC_TEXT         As Integer = 1  ' ** The dates.
Private Const MCSC_TITLEBK      As Integer = 2  ' ** Background of the title.
Private Const MCSC_TITLETEXT    As Integer = 3
Private Const MCSC_MONTHBK      As Integer = 4  ' ** Background within the month cal.
Private Const MCSC_TRAILINGTEXT As Integer = 5  ' ** The text color of header & trailing days.

Private Const MCS_DAYSTATE      As Integer = &H1
Private Const MCS_MULTISELECT   As Integer = &H2
Private Const MCS_WEEKNUMBERS   As Integer = &H4
Private Const MCS_NOTODAYCIRCLE As Integer = &H8
Private Const MCS_NOTODAY       As Integer = &H10

'Private Const GMR_VISIBLE  As Integer = 0  ' ** Visible portion of display.
'Private Const GMR_DAYSTATE As Integer = 1  ' ** Above plus the grayed out parts of partially displayed months.

Private Const POINTSPERINCH As Long = 72

'Private Const cerrPropertyNotFound As Integer = 3270

' ** Generally use Hex values but
' ** Decimal is clearer for educational purposes.
'Private Const Bit10 As Integer = 128
'Private Const Bit9  As Integer = 128
'Private Const Bit8  As Integer = 128
'Private Const Bit7  As Integer = 64
'Private Const Bit6  As Integer = 32
'Private Const Bit5  As Integer = 16
'Private Const Bit4  As Integer = 8
'Private Const Bit3  As Integer = 4
'Private Const Bit2  As Integer = 2
'Private Const Bit1  As Integer = 1

' ** Not used anymore
Private Const NotAColor As Double = &HEF000000

' ** Leftover DAO props so that Reference to DAO is no longer required.
Private Const dbLongDAO As Integer = 4
Private Const dbTextDAO As Integer = 10

' ** Module Level Variables.

' ** Junk vars.
Private lngRetVal As Long
Private blnRetVal As Boolean
Private varTmp00 As Variant

' ** Current Y axis Resolution in Dots per Inch.
Private lngYdpi As Long

' ** General purpose RECT structure.
Private typRECT As RECT

' ** There is a bug in Access when using the SetWindowPos API
' ** on a Form from an external Class module.
' ** Specifically if the Form contains a Header or a
' ** Command Button they are resized along with the Form.
' ** To get around this we need to store the original
' ** height of the CommandButton.
'Private originalCmdButHeight As Integer

' ** hWnd for the Month Calendar control.
Private m_hWndDTP As Long
' ** hWnd for the Form used to display the Calendar.
'Private m_Form As Access.Form

' ** Properties used to create the Month Calendar.
Private m_MonthColumns  As Integer
Private m_MonthRows  As Integer
Private m_SelectedDate As Date

' ** To init SetSelectedDateRange.
Private m_StartDate As Date
Private m_EndDate As Date

Private m_FontName As String
Private m_FontSize As Integer
Private m_ShowWeekNumbers As Long

' ** Allow the user to select a Range of Dates.
Private m_MultiSelect As Long

' Max number of days in a user selectable range
Private m_MaxSelectRangeofDays As Integer

' Offset from the Top of our Form where we draw the
' Month Calendar. Allows any controls we have placed at the
' top of the Form to be visible.
Private m_CalendarYOffset As Integer

' ** Month Calendar Color props
Private m_BackColor As Long
Private m_ForeColor As Long
Private m_MonthBackColor As Long
Private m_TitleBackColor As Long
Private m_TitleForeColor As Long
Private m_TrailingForeColor As Long
' ** Original values of above color props
' ** that exist at time of Calendar creation.
' ** These color values are as per the Users Window setting.
Private m_BackColorOrig As Long
Private m_ForeColorOrig As Long
Private m_MonthBackColorOrig As Long
Private m_TitleBackColorOrig As Long
Private m_TitleForeColorOrig As Long
Private m_TrailingForeColorOrig As Long

Private m_NoToday As Long
Private m_NoTodayCircle As Long

' ** Store current Cursor location for this session.
Private m_cursorX As Long
Private m_cursorY As Long
Private m_PositionAtCursor As Boolean
' ** Cursor coords when Calendar is first loaded.
Private m_cursorXinitpos As Long
Private m_cursorYinitpos As Long
' ** Positioning of Calendar Window.
Private m_WindowLocation As Long

' ** Requried to calculate Calendar Height.
Private m_TodayHeight As Long

' ** Temp vars to hold the size of our Calendar.
Private lngTmpRight As Long
Private lngTmpBottom As Long
Private hFontPrevious As Long

Public m_Hwnd As Long
Private m_hWndForm As Long

' ** Click once to close the Calendar?
Private m_oneClick As Boolean

' ** Max 12 Months displayed at once
' ** plus one month before and after = 14 months
' ** Allow for years 1000 to 3000.
Private arr_typMDS(1000 To 3000, 1 To 14) As MONTHDAYSTATE
' **

Private Sub Class_Initialize()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Class_Initialize"

        'Dim pt As POINTAPI

        ' ** Number of pixels to leave at the Top of the form before we
        ' ** begin to draw our Calendar. This leaves room for any
        ' ** controls we have at the top of the form.
110     Me.CalendarYOffset = 20

        ' ** Default. Number of Columns and Rows.
        ' ** 6 month display
        'Me.MonthColumns = 3
        'Me.MonthRows = 2

        ' ** Check and see if Custom Prop exists.
120     blnRetVal = GetProperty("MonthColumns", varTmp00)
130     If blnRetVal Then
140       m_MonthColumns = varTmp00
150     Else
160       m_MonthColumns = 3
170     End If

180     blnRetVal = GetProperty("MonthRows", varTmp00)
190     If blnRetVal Then
200       m_MonthRows = varTmp00
210     Else
220       m_MonthRows = 2
230     End If

        ' ** Default. Font stuff.
        ' ** Check and see if Custom Prop exists.
240     blnRetVal = GetProperty("FontName", varTmp00)
250     If blnRetVal Then
260       Me.FontName = varTmp00
270     Else
280       Me.FontName = "Arial"
290     End If

300     blnRetVal = GetProperty("FontSize", varTmp00)
310     If blnRetVal Then
320       Me.FontSize = varTmp00
330     Else
340       Me.FontSize = 8
350     End If

360     hFontPrevious = 0

370     blnRetVal = GetProperty("MCS_MULTISELECT", varTmp00)
380     If blnRetVal Then
390       MultiSelect = varTmp00
400     Else
          ' ** MultiSelect to start.
410       MultiSelect = MCS_MULTISELECT  ' ** 0 to Turn Off.
420     End If

430     blnRetVal = GetProperty("MCS_WEEKNUMBERS", varTmp00)
440     If blnRetVal Then
450       ShowWeekNumbers = varTmp00
460     Else
          ' ** MultiSelect to start.
470       ShowWeekNumbers = MCS_WEEKNUMBERS  ' ** 0 to Turn Off.
480     End If

490     blnRetVal = GetProperty("NoToday", varTmp00)
500     If blnRetVal Then
510       m_NoToday = varTmp00
520     Else
          ' ** Show Today is turned off.
530       m_NoToday = MCS_NOTODAY  ' ** 0 to Turn Off.
540     End If

550     blnRetVal = GetProperty("NoTodayCircle", varTmp00)
560     If blnRetVal Then
570       m_NoTodayCircle = varTmp00
580     Else
          ' ** Show Today is turned off.
590       m_NoToday = MCS_NOTODAYCIRCLE  ' ** 0 to Turn Off.
600     End If

610     blnRetVal = GetProperty("WindowLocation", varTmp00)
620     If blnRetVal Then
630       m_WindowLocation = varTmp00
640     Else
          ' ** Window Location.
650       m_WindowLocation = 0
660     End If

        ' ** June 7, 2003
        ' ** Fixing PositionAtCursor logic.
670     If varTmp00 = 0 Then
          ' ** Init.
680       Me.PositionAtCursor = True
690     Else
700       Me.PositionAtCursor = False
710     End If

        ' ** Only load if PositionAtCursor is False.
720     If Not Me.PositionAtCursor Then
730       blnRetVal = GetProperty("Cursorx", varTmp00)
740       If blnRetVal Then
750         m_cursorX = varTmp00
760       Else
            ' ** Leave Alone!
770         m_cursorX = 0
780       End If
790     End If

        ' ** Only load if PositionAtCursor is False.
800     If Not Me.PositionAtCursor Then
810       blnRetVal = GetProperty("CursorY", varTmp00)
820       If blnRetVal Then
830         m_cursorY = varTmp00
840       Else
            ' ** Get current Cursor Position!
            'lngRetVal = GetCursorPos(pt)
            'm_cursorXinit = pt.x
            'm_cursorYinit = pt.y
            'm_cursorX = pt.x
            'm_cursorY = pt.y
850         m_cursorY = 0
860       End If
870     End If

        ' ** Check our Color props.
        ' ** Set var to NotAColor constant if prop doesn't exist.
880     blnRetVal = GetProperty("BackColor", varTmp00)
890     If blnRetVal Then
900       m_BackColor = varTmp00
910     Else
          ' ** Set var to nonexistent color.
920       m_BackColor = NotAColor
930     End If

940     blnRetVal = GetProperty("MonthBackColor", varTmp00)
950     If blnRetVal Then
960       m_MonthBackColor = varTmp00
970     Else
          ' ** Set var to nonexistent color.
980       m_MonthBackColor = NotAColor
990     End If

1000    blnRetVal = GetProperty("ForeColor", varTmp00)
1010    If blnRetVal Then
1020      m_ForeColor = varTmp00
1030    Else
          ' ** Set var to nonexistent color.
1040      m_ForeColor = NotAColor
1050    End If

1060    blnRetVal = GetProperty("TitleBackColor", varTmp00)
1070    If blnRetVal Then
1080      m_TitleBackColor = varTmp00
1090    Else
          ' ** Set var to nonexistent color.
1100      m_TitleBackColor = NotAColor
1110    End If

1120    blnRetVal = GetProperty("TitleForeColor", varTmp00)
1130    If blnRetVal Then
1140      m_TitleForeColor = varTmp00
1150    Else
          ' ** Set var to nonexistent color.
1160      m_TitleForeColor = NotAColor
1170    End If

1180    blnRetVal = GetProperty("TrailingForeColor", varTmp00)
1190    If blnRetVal Then
1200      m_TrailingForeColor = varTmp00
1210    Else
          ' ** Set var to nonexistent color.
1220      m_TrailingForeColor = NotAColor
1230    End If

1240    blnRetVal = GetProperty("OneClick", varTmp00)
1250    If blnRetVal Then
1260      m_oneClick = varTmp00
1270    Else
          ' ** Set var to nonexistent color.
1280      m_oneClick = True
1290    End If

        ' ** Max number of days user can select in a continuous range.
1300    m_MaxSelectRangeofDays = 31
        ' ** To select a Range of Dates add code like below.
        'm_EndDate = #2/27/00#
        'm_StartDate = #2/21/00#

        ' ** Set DayState.
        ' ** For this example we will set one day per month
        ' ** corresponding to the Month number.
1310    Erase arr_typMDS
        ' ** Here is an example of how to Set the BoldDayState prop for Jan 19,2004.
        ' ** i.e. SetBoldDayState numYear 2004, numMonth:=1, day = 19

EXITP:
1320    Exit Sub

ERRH:
1330    Select Case ERR.Number
        Case Else
1340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1350    End Select
1360    Resume EXITP

End Sub

Public Sub Class_Terminate()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Class_Terminate"

        Dim hWndTemp As Long

        Const CLASSNAME As String = "MonthCalendar"
        Const Title As String = "Month Calendar"

        ' ** Cleanup after ourselves.
1410  On Error Resume Next

        ''debug.print "start ClsMonthCalendat Terminated:" & Now
        ' ** March 05 Debug.
        ' ** Not getting a clean exit when user Closes
        ' ** Calling Form with Calendar still Open.
1420    If m_hWndDTP <> 0 Then
1430      lngRetVal = DestroyWindow(m_hWndDTP)
          ''debug.print "Destroyed m_hWndDTP:" & lngRetVal
1440    End If

        ' ** Wait until Window is definately gone!
1450    lngRetVal = DestroyWindow(Me.m_Hwnd)
        ''debug.print "Destroyed Me.m_Hwnd:" & lngRetVal
1460    DoEvents

        ' ** If this window already exists then exit!
1470    hWndTemp = 1
1480    Do While hWndTemp <> 0
1490      hWndTemp = FindWindow(CLASSNAME, Title)
          ''debug.print "MonthCal Hwnd -hWndTemp:" & hWndTemp
1500    Loop

        ' ** See if we have a new Font handle hanging around.
1510    If hFontPrevious <> 0 Then
          ''debug.print "deleting FOnt"
1520      DeleteObject hFontPrevious
1530    End If

        ''debug.print "End ClsMonthCalendar Terminated:" & Now

EXITP:
1540    Exit Sub

ERRH:
1550    Select Case ERR.Number
        Case Else
1560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1570    End Select
1580    Resume EXITP

End Sub

Public Sub CreateDTPControl()
' ** Create the Month Calendar.

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "CreateDTPControl"

        ' ** Temp vars.
        Dim typICC As INITCOMMONCONTROLSEXSTRUCT
        Dim lngStyle As Long
        Dim hInstance As Long

        ' ** Temp handle to Information context.
        Dim lngIC As Long

        ' ** Only create MonthCalendar if there is a host window.
1610    If m_Hwnd = 0 Then Exit Sub

1620    If m_hWndDTP <> 0 Then
1630      Call DestroyWindow(m_hWndDTP)
1640    End If

        ' ** Courtesy of Ken Getz.
1650    hInstance = GetWindowLong(Application.hWndAccessApp, GWL_HINSTANCE)  ' ** API Function: modWindowFunctions.
1660    typICC.dwsize = Len(typICC)
1670    typICC.dwICC = ICC_DATE_CLASSES
1680    Call InitCommonControlsEx(typICC)

        ' ** The Calendar MultiSelect property determines whether the
        ' ** user can select a range of dates ' WS_BORDER OrOr WS_OVERLAPPED.
1690    lngStyle = WS_BORDER Or WS_CHILD Or m_MultiSelect _
          Or m_ShowWeekNumbers Or m_NoToday Or m_NoTodayCircle Or m_NoToday Or WS_VISIBLE Or MCS_DAYSTATE  ' ** DayState not ready yet! Or MCS_DAYSTATE

        ' ** Create our Month Calendar
1700    m_hWndDTP = CreateWindowEx(0&, "SysMonthCal32", "DateTimePicker", lngStyle, 0, 0, 250, 250, m_Hwnd, 0&, hInstance, ByVal 0&)

        ' ** Modified to allow for different screen resolutions
        ' ** and printer output. Needed to Calculate Font size.
1710    lngIC = CreateIC("DISPLAY", vbNullString, vbNullString, vbNullString)
1720    If lngIC <> 0 Then
1730      lngYdpi = GetDeviceCaps(lngIC, GSR_LOGPIXELSY)
1740      DeleteDC (lngIC)
1750    Else
1760      lngYdpi = 120  ' ** Default average value.
1770    End If

        ' ** Grab all of the Color props before we mangle them!
1780    m_BackColorOrig = BackColor
1790    m_ForeColorOrig = ForeColor
1800    m_MonthBackColorOrig = MonthBackColor
1810    m_TitleBackColorOrig = TitleBackColor
1820    m_TitleForeColorOrig = TitleForeColor
1830    m_TrailingForeColorOrig = TrailingForeColor

        ' ** Call our Sub to redraw the Form holding our Calendar
        ' ** and position the Calendar appropriately.
1840    ReDraw

EXITP:
1850    Exit Sub

ERRH:
1860    Select Case ERR.Number
        Case Else
1870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1880    End Select
1890    Resume EXITP

End Sub

Public Sub ReDraw()
' ** Redraw and Position the Parent Form and the Calendar.
' ** This is the function to tell us what size the window would be
' ** to hold the entire CLient Rectangle area. In other words it adds
' ** on the height for the Title Bar, Bottom Border etc.

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "ReDraw"

        Dim typRECT1 As RECT
        Dim typRECT2 As RECT
        Dim typRECT3 As RECT
        Dim typST As SYSTEMTIME
        Dim intBarBordersY As Integer
        Dim intBarBordersX As Integer
        Dim lngXtraPixels As Long
        Dim lngHWndDeskTop As Long

        ' ** Only create MonthCalendar if there is a host window.
1910    If m_Hwnd = 0 Then Exit Sub

        ' ** Make sure the Calendar has time to be initially rendered.
1920    DoEvents

        ' ** Re-render using our props for FontName and FontSize.
1930    Me.ChangeFont

        ' ****************************
        ' NOT READY YET!!!!
        ' TEST FOR DAYSTATE
        'Me.DayState = 1

        ' ** Set max range of days user can select.
1940    Call SendMessage(m_hWndDTP, MCM_SETMAXSELCOUNT, m_MaxSelectRangeofDays, ByVal 0&)  ' ** API Function: modWindowFunctions.

        ' ** If EndDate is Set and MultiSelect then call SetSelectedDateRange.
1950    If m_EndDate <> 0 And m_EndDate <> 0 And m_MultiSelect = MCS_MULTISELECT Then
1960      SetSelectedDateRange m_StartDate, m_EndDate
1970    Else
1980      With typST
1990        .wYear = year(m_StartDate)
2000        .wMonth = month(m_StartDate)
2010        .wDay = day(m_StartDate)
2020      End With
2030      Call SendMessage(m_hWndDTP, MCM_SETCURSEL, ByVal 0&, typST)  ' ** API Function: modWindowFunctions.
2040    End If

        ' ** Get the size required to display 1 complete month of the calendar.
2050    Call SendMessage(m_hWndDTP, MCM_GETMINREQRECT, ByVal 0&, typRECT)  ' ** API Function: modWindowFunctions.

        ' ** Check is TODAY is showing at bottom of Calendar
2060    If m_NoToday = 0 Then
2070      lngXtraPixels = SendMessage(m_hWndDTP, MCM_GETMAXTODAYWIDTH, ByVal 0&, ByVal 0&)  ' ** API Function: modWindowFunctions.
2080    Else
2090      lngXtraPixels = typRECT.Right ' = lngXtraPixels
2100    End If

        ' ** Set color props
2110    If m_BackColor <> NotAColor Then
2120      Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_BACKGROUND, ByVal m_BackColor)  ' ** API Function: modWindowFunctions.
2130    End If

2140    If m_ForeColor <> NotAColor Then
2150      Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TEXT, ByVal m_ForeColor)  ' ** API Function: modWindowFunctions.
2160    End If

2170    If m_MonthBackColor <> NotAColor Then
2180      Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_MONTHBK, ByVal m_MonthBackColor)  ' ** API Function: modWindowFunctions.
2190    End If

2200    If m_TitleForeColor <> NotAColor Then
2210      Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TITLETEXT, ByVal m_TitleForeColor)  ' ** API Function: modWindowFunctions.
2220    End If

2230    If m_TitleBackColor <> NotAColor Then
2240      Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TITLEBK, ByVal m_TitleBackColor)  ' ** API Function: modWindowFunctions.
2250    End If

2260    If m_TrailingForeColor <> NotAColor Then
2270      Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TRAILINGTEXT, ByVal m_TrailingForeColor)  ' ** API Function: modWindowFunctions.
2280    End If

        ' ** Get rectangle for our Form
        'Debug.Print "GetWindowRect- Me.hWnd:" & m_Form.hWnd
2290    lngRetVal = GetWindowRect(m_hWndDTP, typRECT1)  ' ** API Function: modWindowFunctions.
2300    lngRetVal = GetClientRect(m_hWndDTP, typRECT2)  ' ** API Function: modWindowFunctions.
        ' ** Calculate Borders and Title Bar area.
2310    intBarBordersY = ((typRECT1.Bottom - typRECT1.Top) - (typRECT2.Bottom - typRECT2.Top)) + 10
2320    intBarBordersX = ((typRECT1.Right - typRECT1.Left) - (typRECT2.Right - typRECT2.Left)) + 10

2330    With typRECT
          ' ** Handle possible extra Today string Width
2340      typRECT.Right = lngXtraPixels

          ' ** An OS bug/feature requires us to round up
2350      If m_NoToday <> 0 Then
2360        lngTmpRight = (.Right * m_MonthColumns) + (m_MonthColumns - 1) * 6 'intBarBordersX 'm_MonthColumns * 2 '5
2370        lngTmpBottom = (.Bottom * m_MonthRows) + (m_MonthRows * (m_TodayHeight * 1.5)) ' (m_TodayHeight * 4)  '(m_MonthRows - 1) * 24 '12 'intBarBordersY 'm_MonthRows * m_TodayHeight ')) ' + 10 ''m_MonthRows * 15
            ' ** ADD CODE TO HELP WITH NUMBER OF ROWS.
            ' ** If only 1 run then only need *2 or *3
2380      Else
2390        lngTmpRight = (.Right * m_MonthColumns) + (m_MonthColumns - 1) * 6 'intBarBordersX '+ m_MonthColumns * 2 '5
2400        lngTmpBottom = (.Bottom * m_MonthRows) ' + (m_MonthRows - 1) * 1 '12 'intBarBordersY '+ m_MonthRows * 10
2410      End If

2420    End With
        ' ** Resize the Month Calendar to display the user selected
        ' ** number of months. The CalendarYOffset is used to allow
        ' ** any controls we have placed at the Top of our Form
        ' ** to be visible.
        ' ** ***DEBUG - BUG FIX ******
        ' ** Try to fix visual display bug
        ' ** when only 1 month is selected.
        ' ** The left most column dissappears when
        ' ** when selecting a range of dates
        ' ** Add 4 pixels to the COntrol's Width

2430    If m_MonthRows = 1 And m_MonthColumns = 1 Then
2440      Call SetWindowPos(m_hWndDTP, 0&, 0&, 0&, lngTmpRight + 4, lngTmpBottom, 0&)  ' ** API Function: modWindowFunctions.
2450    Else
2460      Call SetWindowPos(m_hWndDTP, 0&, 0&, 0&, lngTmpRight, lngTmpBottom, 0&)  ' ** API Function: modWindowFunctions.
2470    End If

        ' ** Get rectangle for our Form
        'Debug.Print "GetWindowRect- Me.hWnd:" & m_Form.hWnd
        'lngRetVal = GetWindowRect(m_Hwnd, typRECT1)  ' ** API Function: modWindowFunctions.
        'lngRetVal = GetClientRect(m_Hwnd, typRECT2)
        ' ** Get rectangle for our Calendar
2480    lngRetVal = GetWindowRect(m_hWndDTP, typRECT3)  ' ** API Function: modWindowFunctions.

2490    lngRetVal = GetWindowLong(m_Hwnd, GWL_STYLE)
2500    lngRetVal = AdjustWindowRect(typRECT3, lngRetVal, -1)

        ' ** Resize our Form to display the entire Calendar
        ' ** Position Window at current cursor location.
        ' ** We will use the Cursor location when the Window was first
        ' ** created, not the current location. User could be
        ' ** selecting a Menu when this Redraw Sub is called.
        ' ** We also need to allow the user to manually move
        ' ** and place this window and respect those
        ' ** coordinates for this session.
        ' ** So if the current XY WIndow Coords are not the same
        ' ** as the stored Mouse Coords then let the current
        ' ** Calendars Parent Window coords stand.
        ' ** We do this directly in the Windows Procedure
        ' ** by resetting the CursorX and CursorY props
        ' ** on the Window Move message.

2510    If m_PositionAtCursor Then
          ' ** Position Window at Cursor Location when
          ' ** Calendar was created.

          ' ** Also add an option to position cursor itself in middle of calendar window
          ' ** May 05 - update m_cursorXinitpos and m_cursorYinitpos to reflect current cursor location when this calendar is opened
          ' ** This is done in the modCalendar module when the parent window for the Calendar is created.

          ' ** June 7, 2003 fixing position at cursor logic
          ' ** PositionAtCursor = True

          ' ** First check to ensure entire calendar window will fit.
          ' ** If not MAKE IT!
          ' ** COMING IN NEXT RELEASE!

2520      With typRECT
2530        Call SetWindowPos(m_Hwnd, 0&, m_cursorX, m_cursorY, typRECT3.Right - typRECT3.Left, typRECT3.Bottom - typRECT3.Top, 0&)  ' ** API Function: modWindowFunctions.
2540      End With

2550    Else
          ' ** Position Window according to users Menu selections
          ' ** a) 0 -Pop at cursor location when user activates Calendar
          ' ** b) 1 -Where they manually move/leave it at
          ' ** c) 2 -Centered in Access App Window
          ' ** d) 3 -Centered on entire screen
          ' ** d) 4 -Top Left Corner
          'Debug.Print "m_WindowLocation:" & m_WindowLocation

2560      Select Case m_WindowLocation

          Case 1
2570        With typRECT
2580          Call SetWindowPos(m_Hwnd, 0&, m_cursorX, m_cursorY, typRECT3.Right - typRECT3.Left, typRECT3.Bottom - typRECT3.Top, 0&)  ' ** API Function: modWindowFunctions.
2590        End With

2600      Case 2
            ' ** Centered within main Access App Window
            ' ** Get rectangle for our Calendar
2610        lngRetVal = GetWindowRect(Application.hWndAccessApp, typRECT2)  ' ** API Function: modWindowFunctions.
2620        With typRECT
2630          Call SetWindowPos(m_Hwnd, 0&, ((typRECT2.Right - typRECT2.Left) - (typRECT3.Right - typRECT3.Left)) / 2, _
                ((typRECT2.Bottom - typRECT2.Top) - (typRECT3.Bottom - typRECT3.Top)) / 2, typRECT3.Right - typRECT3.Left, _
                typRECT3.Bottom - typRECT3.Top, 0&)  ' ** API Function: modWindowFunctions.
2640        End With

2650      Case 3
            ' ** Centered within Screen
            ' ** Get rectangle for our Calendar
2660        lngHWndDeskTop = GetDesktopWindow()
2670        lngRetVal = GetWindowRect(lngHWndDeskTop, typRECT2)  ' ** API Function: modWindowFunctions.
2680        With typRECT
2690          Call SetWindowPos(m_Hwnd, 0&, ((typRECT2.Right - typRECT2.Left) - (typRECT3.Right - typRECT3.Left)) / 2, _
                ((typRECT2.Bottom - typRECT2.Top) - (typRECT3.Bottom - typRECT3.Top)) / 2, typRECT3.Right - typRECT3.Left, _
                typRECT3.Bottom - typRECT3.Top, 0&)  ' ** API Function: modWindowFunctions.
2700        End With

2710      Case 4
            ' ** Upper Left Hand Corner of Screen
            ' ** Get rectangle for our Calendar
2720        lngRetVal = GetWindowRect(lngHWndDeskTop, typRECT2)  ' ** API Function: modWindowFunctions.
2730        With typRECT
2740          Call SetWindowPos(m_Hwnd, 0&, 2, 2, typRECT3.Right - typRECT3.Left, _
                typRECT3.Bottom - typRECT3.Top, 0&)  ' ** API Function: modWindowFunctions.
2750        End With

2760      Case Else
2770        With typRECT
2780          Call SetWindowPos(m_Hwnd, 0&, m_cursorX, m_cursorY, typRECT3.Right - typRECT3.Left, typRECT3.Bottom - typRECT3.Top, 0&)  ' ** API Function: modWindowFunctions.
2790        End With

2800      End Select
2810    End If

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Public Sub ResetCalendarColors()
' ** Reset all of the color props to their original setting
' ** as per the users system settings for the GUI.

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "ResetCalendarColors"

2910    BackColor = m_BackColorOrig
2920    ForeColor = m_ForeColorOrig
2930    MonthBackColor = m_MonthBackColorOrig
2940    TitleBackColor = m_TitleBackColorOrig
2950    TitleForeColor = m_TitleForeColorOrig
2960    TrailingForeColor = m_TrailingForeColorOrig

EXITP:
2970    Exit Sub

ERRH:
2980    Select Case ERR.Number
        Case Else
2990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3000    End Select
3010    Resume EXITP

End Sub

Public Property Let MultiSelect(ms As Boolean)
' ** Allow user to select a Range of dates instead of just one.

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "MultiSelect"

3110    If ms Then
3120      m_MultiSelect = MCS_MULTISELECT
3130    Else
3140      m_MultiSelect = 0
3150    End If

        ' ** Update custom property.
3160    SetProperty "MultiSelect", dbLongDAO, m_MultiSelect

EXITP:
3170    Exit Property

ERRH:
3180    Select Case ERR.Number
        Case Else
3190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3200    End Select
3210    Resume EXITP

End Property

Public Property Get MultiSelect() As Boolean

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "MultiSelect"

3310    If m_MultiSelect = MCS_MULTISELECT Then
3320      MultiSelect = True
3330    Else
3340      MultiSelect = False
3350    End If

EXITP:
3360    Exit Property

ERRH:
3370    Select Case ERR.Number
        Case Else
3380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3390    End Select
3400    Resume EXITP

End Property

Public Property Let MaxSelectRangeofDays(maxdays As Integer)
' ** Maximum days user can select in a range of dates.

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "MaxSelectRangeofDays"

3510    If maxdays > 365 Or maxdays < 0 Then
3520      m_MaxSelectRangeofDays = 7
3530    Else
3540      m_MaxSelectRangeofDays = maxdays
3550    End If
3560    CreateDTPControl

EXITP:
3570    Exit Property

ERRH:
3580    Select Case ERR.Number
        Case Else
3590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3600    End Select
3610    Resume EXITP

End Property

Public Property Get MaxSelectRangeofDays() As Integer

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "MaxSelectRangeofDays"

3710    MaxSelectRangeofDays = m_MaxSelectRangeofDays

EXITP:
3720    Exit Property

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Property

Public Property Let CalendarYOffset(YOffset As Integer)
' ** Offset from the Top of our Form where we will
' ** draw our Calendat. This will leave us room at the
' ** the top of our Form for our controls.
' ** Some boundary checking. We'll arbitrarily use 3 inches for now.
' ** This prop is expressed in TWIPS!

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "CalendarYOffset"

3810    If YOffset > 3 * RZ_TWIPSPERINCH Or YOffset < 0 Then
3820      m_CalendarYOffset = 0
3830    Else
3840      m_CalendarYOffset = YOffset
3850    End If
        'CreateDTPControl

EXITP:
3860    Exit Property

ERRH:
3870    Select Case ERR.Number
        Case Else
3880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3890    End Select
3900    Resume EXITP

End Property

Public Property Get CalendarYOffset() As Integer

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "CalendarYOffset"

4010    CalendarYOffset = m_CalendarYOffset

EXITP:
4020    Exit Property

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Property

Public Property Let MonthRows(ByVal newMonthRows As Byte)
' ** The Month Calendar allows for a maximum of 12 months
' ** to be displayed at once.

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "MonthRows"

4110    If newMonthRows > 12 Or newMonthRows < 0 Then
4120      newMonthRows = 1
4130    End If
4140    If newMonthRows * m_MonthColumns > 12 Then
4150      m_MonthRows = 12 / m_MonthColumns
4160    Else
4170      m_MonthRows = newMonthRows
4180    End If

        ' ** Update custom property.
4190    Call SetProperty("MonthRows", dbLongDAO, m_MonthRows)

EXITP:
4200    Exit Property

ERRH:
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Property

Public Property Get MonthRows() As Byte

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "MonthRows"

4310    MonthRows = m_MonthRows

EXITP:
4320    Exit Property

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Property

Public Property Let MonthColumns(ByVal newMonthColumns As Byte)
' ** The Month Calendar allows for a maximum of 12 months
' ** to be displayed at once.

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "MonthColumns"

4410    If newMonthColumns > 12 Or newMonthColumns < 0 Then
4420      newMonthColumns = 1
4430    End If

4440    If newMonthColumns * m_MonthRows > 12 Then
4450      m_MonthColumns = 12 / m_MonthRows
4460    Else
4470      m_MonthColumns = newMonthColumns
4480    End If

        ' ** Update custom property.
4490    Call SetProperty("MonthColumns", dbLongDAO, m_MonthColumns)

EXITP:
4500    Exit Property

ERRH:
4510    Select Case ERR.Number
        Case Else
4520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4530    End Select
4540    Resume EXITP

End Property

Public Property Get MonthColumns() As Byte

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "MonthColumns"

4610    MonthColumns = m_MonthColumns

EXITP:
4620    Exit Property

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Property

Public Property Let WindowLocation(loc As Long)
' ** Allow user to select position of Calendar Window

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "WindowLocation"

        'Debug.Print "Prop Let:" & loc

4710    If loc < 0 Or loc > 4 Then
4720      m_WindowLocation = 0
4730    Else
4740      m_WindowLocation = loc
4750    End If

        ' ** Update custom property.
4760    SetProperty "WindowLocation", dbLongDAO, m_WindowLocation
        'Debug.Print "Saved Prop :" & m_WindowLocation

EXITP:
4770    Exit Property

ERRH:
4780    Select Case ERR.Number
        Case Else
4790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4800    End Select
4810    Resume EXITP

End Property

Public Property Get WindowLocation() As Long

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "WindowLocation"

4910    WindowLocation = m_WindowLocation

EXITP:
4920    Exit Property

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Property

Public Property Let PositionAtCursor(bl As Boolean)
' ** Does user want to popup Calendar at Current Cursor location
' ** when they invoke the Calendar?

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "PositionAtCursor"

        'Dim pt As PointAPI

5010    m_PositionAtCursor = bl
5020    If bl Then
          'lngRetVal = GetCursorPos(pt)
5030      m_cursorX = m_cursorXinitpos
5040      m_cursorY = m_cursorYinitpos
5050    End If

EXITP:
5060    Exit Property

ERRH:
5070    Select Case ERR.Number
        Case Else
5080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5090    End Select
5100    Resume EXITP

End Property

Public Property Get PositionAtCursor() As Boolean

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "PositionAtCursor"

5210    PositionAtCursor = m_PositionAtCursor

EXITP:
5220    Exit Property

ERRH:
5230    Select Case ERR.Number
        Case Else
5240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5250    End Select
5260    Resume EXITP

End Property

Public Property Let CursorXinit(X As Long)
' ** Store Cursor position.

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "CursorXinit"

5310    m_cursorXinitpos = X
5320    If Me.PositionAtCursor = True Then
5330      m_cursorX = X
5340    End If

        ' ** Update custom property.
        'Call SetProperty("CursorX", dbLongDAO, m_cursorX)

EXITP:
5350    Exit Property

ERRH:
5360    Select Case ERR.Number
        Case Else
5370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5380    End Select
5390    Resume EXITP

End Property

Public Property Let CursorYinit(Y As Long)
' ** Store Cursor position.

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "CursorYinit"

5410    m_cursorYinitpos = Y
5420    If Me.PositionAtCursor = True Then
5430      m_cursorY = Y
5440    End If

        ' ** Redraw the control.
5450    CreateDTPControl

        ' ** Update custom property.
        'Call SetProperty("CursorY", dbLongDAO, m_cursorY)

EXITP:
5460    Exit Property

ERRH:
5470    Select Case ERR.Number
        Case Else
5480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5490    End Select
5500    Resume EXITP

End Property

Public Property Let CursorX(X As Long)
' ** X pos.

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "CursorX"

5610    m_cursorX = X

        ' ** Update custom property.
5620    Call SetProperty("CursorX", dbLongDAO, m_cursorX)

EXITP:
5630    Exit Property

ERRH:
5640    Select Case ERR.Number
        Case Else
5650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5660    End Select
5670    Resume EXITP

End Property

Public Property Let CursorY(Y As Long)
' ** Y pos.

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "CursorY"

5710    m_cursorY = Y

        ' ** Update custom property.
5720    Call SetProperty("CursorY", dbLongDAO, m_cursorY)

EXITP:
5730    Exit Property

ERRH:
5740    Select Case ERR.Number
        Case Else
5750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5760    End Select
5770    Resume EXITP

End Property

Public Property Get CursorX() As Long
' ** X pos.

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "CursorX"

5810    CursorX = m_cursorX

EXITP:
5820    Exit Property

ERRH:
5830    Select Case ERR.Number
        Case Else
5840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5850    End Select
5860    Resume EXITP

End Property

Public Property Get CursorY() As Long
' ** Y Pos.

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "CursorY"

5910    CursorY = m_cursorY

EXITP:
5920    Exit Property

ERRH:
5930    Select Case ERR.Number
        Case Else
5940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5950    End Select
5960    Resume EXITP

End Property

Public Property Let FontSize(fs As Integer)
' ** Font Size to render the Calendar with.

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "FontSize"

6010    If fs = 0 Or fs < 0 Then fs = 8
6020    If fs > 24 Then fs = 8
6030    m_FontSize = fs

        ' ** Update custom property.
6040    SetProperty "FontSize", dbLongDAO, m_FontSize

EXITP:
6050    Exit Property

ERRH:
6060    Select Case ERR.Number
        Case Else
6070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6080    End Select
6090    Resume EXITP

End Property

Public Property Get FontSize() As Integer

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "FontSize"

6110    FontSize = m_FontSize

EXITP:
6120    Exit Property

ERRH:
6130    Select Case ERR.Number
        Case Else
6140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6150    End Select
6160    Resume EXITP

End Property

Public Property Let FontName(fn As String)
' ** Font Name to render the Calendar with.

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "FontName"

6210    m_FontName = fn

        ' ** Update custom property.
6220    Call SetProperty("FontName", dbTextDAO, m_FontName)

EXITP:
6230    Exit Property

ERRH:
6240    Select Case ERR.Number
        Case Else
6250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6260    End Select
6270    Resume EXITP

End Property

Public Property Get FontName() As String

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "FontName"

6310    FontName = m_FontName

EXITP:
6320    Exit Property

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Property

Public Property Let ShowWeekNumbers(bl As Boolean)
' ** Show Week Numbers.

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowWeekNumbers"

6410    If bl Then
6420      m_ShowWeekNumbers = MCS_WEEKNUMBERS
6430    Else
6440      m_ShowWeekNumbers = 0
6450    End If

        ' ** Update custom property.
6460    Call SetProperty("MCS_WEEKNUMBERS", dbLongDAO, m_ShowWeekNumbers)

        ' ** Redraw the control.
6470    CreateDTPControl

EXITP:
6480    Exit Property

ERRH:
6490    Select Case ERR.Number
        Case Else
6500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6510    End Select
6520    Resume EXITP

End Property

Public Property Get ShowWeekNumbers() As Boolean

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowWeekNumbers"

6610    If m_ShowWeekNumbers = MCS_WEEKNUMBERS Then
6620      ShowWeekNumbers = True
6630    Else
6640      ShowWeekNumbers = False
6650    End If

EXITP:
6660    Exit Property

ERRH:
6670    Select Case ERR.Number
        Case Else
6680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6690    End Select
6700    Resume EXITP

End Property

Public Property Get NoTodayCircle() As Boolean

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "NoTodayCircle"

6810    If m_NoTodayCircle = 0 Then
6820      NoTodayCircle = False
6830    Else
6840      NoTodayCircle = True
6850    End If

EXITP:
6860    Exit Property

ERRH:
6870    Select Case ERR.Number
        Case Else
6880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6890    End Select
6900    Resume EXITP

End Property

Public Property Let NoTodayCircle(bl As Boolean)
' ** Show Circle around Today's Date.

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "NoTodayCircle"

7010    If bl Then
7020      m_NoTodayCircle = MCS_NOTODAYCIRCLE
7030    Else
7040      m_NoTodayCircle = 0
7050    End If

        ' ** Update custom property.
7060    Call SetProperty("NoTodayCircle", dbLongDAO, m_NoTodayCircle)

        ' ** Redraw the control.
7070    CreateDTPControl

EXITP:
7080    Exit Property

ERRH:
7090    Select Case ERR.Number
        Case Else
7100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7110    End Select
7120    Resume EXITP

End Property

Public Property Get NoToday() As Boolean

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "NoToday"

7210    If m_NoToday = 0 Then
7220      NoToday = False
7230    Else
7240      NoToday = True
7250    End If

EXITP:
7260    Exit Property

ERRH:
7270    Select Case ERR.Number
        Case Else
7280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7290    End Select
7300    Resume EXITP

End Property

Public Property Let NoToday(bl As Boolean)
' ** Show Today's date at bottom of Calendar.

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "NoToday"

7410    If bl Then
7420      m_NoToday = MCS_NOTODAY
7430    Else
7440      m_NoToday = 0
7450    End If

        ' ** Update custom property.
7460    Call SetProperty("NoToday", dbLongDAO, m_NoToday)

        ' ** Redraw the control.
7470    CreateDTPControl

EXITP:
7480    Exit Property

ERRH:
7490    Select Case ERR.Number
        Case Else
7500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7510    End Select
7520    Resume EXITP

End Property

Public Property Let hwnd(h As Long)
' ** Host Window.

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "Hwnd"

7610    m_Hwnd = h

        ' ** ReCreate the control.
7620    CreateDTPControl

EXITP:
7630    Exit Property

ERRH:
7640    Select Case ERR.Number
        Case Else
7650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7660    End Select
7670    Resume EXITP

End Property

Public Property Get hwnd() As Long
' ** MonthCalendar Parent Window.

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "Hwnd"

7710    hwnd = m_Hwnd

EXITP:
7720    Exit Property

ERRH:
7730    Select Case ERR.Number
        Case Else
7740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7750    End Select
7760    Resume EXITP

End Property

Public Property Let hWndForm(h As Long)
' ** Calling Form's Window Handle.

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "hWndForm"

7810    m_hWndForm = h

EXITP:
7820    Exit Property

ERRH:
7830    Select Case ERR.Number
        Case Else
7840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7850    End Select
7860    Resume EXITP

End Property

Public Property Get hWndForm() As Long
' ** Calling Form's Window Handle.

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "hWndForm"

7910    hWndForm = m_hWndForm

EXITP:
7920    Exit Property

ERRH:
7930    Select Case ERR.Number
        Case Else
7940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7950    End Select
7960    Resume EXITP

End Property

Public Property Get hWndCal() As Long
' ** MonthCalendar Window.

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "hWndCal"

8010    hWndCal = m_hWndDTP

EXITP:
8020    Exit Property

ERRH:
8030    Select Case ERR.Number
        Case Else
8040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8050    End Select
8060    Resume EXITP

End Property

Public Property Let OneClick(bl As Boolean)
' ** Click Once to Close our Calendar.

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "OneClick"

8110    m_oneClick = bl

        ' ** Update custom property.
8120    Call SetProperty("OneClick", dbLongDAO, m_oneClick)

EXITP:
8130    Exit Property

ERRH:
8140    Select Case ERR.Number
        Case Else
8150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8160    End Select
8170    Resume EXITP

End Property

Public Property Get OneClick() As Boolean

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "OneClick"

8210    OneClick = m_oneClick

EXITP:
8220    Exit Property

ERRH:
8230    Select Case ERR.Number
        Case Else
8240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8250    End Select
8260    Resume EXITP

End Property

Public Property Let SelectedDate(ByVal newSelectedDate As Date)
' ** Date pre-selected in the Calendar.
' ** Default is Today's date.

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "SelectedDate"

        Dim typST As SYSTEMTIME

8310    m_SelectedDate = newSelectedDate
8320    With typST
8330      .wYear = year(newSelectedDate)
8340      .wMonth = month(newSelectedDate)
8350      .wDay = day(newSelectedDate)
8360    End With

8370    m_StartDate = newSelectedDate
8380    If m_MultiSelect <> MCS_MULTISELECT Then
8390      Call SendMessage(m_hWndDTP, MCM_SETCURSEL, ByVal 0&, typST)  ' ** API Function: modWindowFunctions.
8400    Else
8410      m_StartDate = newSelectedDate
8420      m_EndDate = m_StartDate
8430      Call SetSelectedDateRange(m_StartDate, m_EndDate)
          ' ** Sets the ...
8440    End If

EXITP:
8450    Exit Property

ERRH:
8460    Select Case ERR.Number
        Case Else
8470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8480    End Select
8490    Resume EXITP

End Property

Public Sub SetSelectedDateRange(ByVal StartDate As Date, ByVal EndDate As Date)
' ** Sets the Range of Dates visible in the Calendar.

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "SetSelectedDateRange"

        Dim arr_typST(1) As SYSTEMTIME

8510    If m_MultiSelect <> MCS_MULTISELECT Then Exit Sub

        ' ** Fill in the start and ending SYSTEMTIME structures.
8520    With arr_typST(0)
8530      .wDay = day(StartDate)
8540      .wMonth = month(StartDate)
8550      .wYear = year(StartDate)
8560    End With

8570    With arr_typST(1)
8580      .wDay = day(EndDate)
8590      .wMonth = month(EndDate)
8600      .wYear = year(EndDate)
8610    End With
8620    lngRetVal = SendMessage(m_hWndDTP, MCM_SETSELRANGE, ByVal 0&, arr_typST(0))  ' ** API Function: modWindowFunctions.
8630    With arr_typST(0)
8640      m_SelectedDate = DateSerial(.wYear, .wMonth, .wDay)
8650    End With
        'With arr_typST(1)
        '  m_endSelectedDate = DateSerial(.wYear, .wMonth, .wDay)
        'End With

EXITP:
8660    Exit Sub

ERRH:
8670    Select Case ERR.Number
        Case Else
8680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8690    End Select
8700    Resume EXITP

End Sub

Public Property Get StartSelectedDate() As Date
' ** Either the current single date selected by the user
' ** or the First date in a user selected range of dates.

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "StartSelectedDate"

        Dim typST As SYSTEMTIME

8810    If m_MultiSelect = MCS_MULTISELECT Then
8820      StartSelectedDate = GetSelectedDates(True)
8830    Else
8840      Call SendMessage(m_hWndDTP, MCM_GETCURSEL, ByVal 0&, typST)  ' ** API Function: modWindowFunctions.
8850      With typST
8860        m_SelectedDate = DateSerial(.wYear, .wMonth, .wDay)
8870      End With
8880      StartSelectedDate = m_SelectedDate
8890    End If

EXITP:
8900    Exit Property

ERRH:
8910    Select Case ERR.Number
        Case Else
8920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8930    End Select
8940    Resume EXITP

End Property

Public Property Get EndSelectedDate() As Date
' ** The End date in a user selected range of dates.

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "EndSelectedDate"

        Dim typST As SYSTEMTIME

9010    If m_MultiSelect = MCS_MULTISELECT Then
9020      EndSelectedDate = GetSelectedDates(False)
9030    Else
9040      Call SendMessage(m_hWndDTP, MCM_GETCURSEL, ByVal 0&, typST)  ' ** API Function: modWindowFunctions.
9050      With typST
9060        m_SelectedDate = DateSerial(.wYear, .wMonth, .wDay)
9070      End With
9080      EndSelectedDate = m_SelectedDate
9090    End If

EXITP:
9100    Exit Property

ERRH:
9110    Select Case ERR.Number
        Case Else
9120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9130    End Select
9140    Resume EXITP

End Property

Private Function GetSelectedDates(StartEnd As Boolean) As Date
' ** Return Start or END date of user selected Range.
' ** If StartEnd = TRUE then return Start Date.
' ** If StartEnd = FALSE then return END Date.

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "GetSelectedDates"

        Dim arr_typST(0 To 1) As SYSTEMTIME
        Dim datStartDate As Date
        Dim datEndDate As Date

        ' ** Get the currently selected range.
9210    lngRetVal = SendMessage(m_hWndDTP, MCM_GETSELRANGE&, ByVal 0&, arr_typST(0))  ' ** API Function: modWindowFunctions.

        ' ** Convert to our Date format.
9220    With arr_typST(0)
9230      datStartDate = DateSerial(.wYear, .wMonth, .wDay)
9240    End With
9250    With arr_typST(1)
9260      datEndDate = DateSerial(.wYear, .wMonth, .wDay)
9270    End With

        ' ** Return specified Start or End date.
9280    If StartEnd Then
9290      GetSelectedDates = datStartDate
9300    Else
9310      GetSelectedDates = datEndDate
9320    End If

EXITP:
9330    Exit Function

ERRH:
9340    Select Case ERR.Number
        Case Else
9350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9360    End Select
9370    Resume EXITP

End Function

Public Sub SetViewableMonths(numMonths As Integer)
' ** How many months are viewable in the Calendar.
' ** Options are 1,2,3,4,6,9,12.

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "SetViewableMonths"

9410    Select Case numMonths
        Case 1
9420      MonthColumns = 1
9430      MonthRows = 1
9440    Case 2
9450      MonthColumns = 2
9460      MonthRows = 1
9470    Case 3
9480      MonthColumns = 3
9490      MonthRows = 1
9500    Case 4
9510      MonthColumns = 2
9520      MonthRows = 2
9530    Case 6
9540      MonthColumns = 3
9550      MonthRows = 2
9560    Case 8
9570      MonthColumns = 4
9580      MonthRows = 2
9590    Case 9
9600      MonthColumns = 3
9610      MonthRows = 3
9620    Case 12
9630      MonthColumns = 4
9640      MonthRows = 3
9650    Case Else
9660      MonthColumns = 1
9670      MonthRows = 1
9680    End Select

        ' ** Create the new Month Calender.
9690    If m_hWndDTP = 0 Then
9700      CreateDTPControl
9710    Else
9720      ReDraw
9730    End If

EXITP:
9740    Exit Sub

ERRH:
9750    Select Case ERR.Number
        Case Else
9760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9770    End Select
9780    Resume EXITP

End Sub

Public Sub SelectFont()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "SelectFont"

        Dim typFFI As FORMFONTINFO

        ' ** Set some Defaults for the Font Dialog.
9810    With typFFI
9820      .Color = 0
9830      .Height = m_FontSize  ' ** 12.
9840      .Weight = 400
9850      .Italic = False
9860      .UnderLine = False
9870      .Name = m_FontName  ' ** "Arial".
9880    End With

        ' ** Call the Font Dialog.
9890    blnRetVal = DialogFont(typFFI, m_hWndDTP)
9900    If blnRetVal Then
          ' ** Copy users selections over to our class vars.
9910      With typFFI
9920        FontName = .Name
9930        FontSize = .Height
9940      End With
          ' ** Store selection to Custom Property.
9950      ReDraw
9960    End If

EXITP:
9970    Exit Sub

ERRH:
9980    Select Case ERR.Number
        Case Else
9990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10000   End Select
10010   Resume EXITP

End Sub

Public Sub ChangeFont()
' ** User selectable Font size for the Calendar.
' ** From limited testing it seems the Calendar
' ** deletes the Font we pass to it. Needs further
' ** testing though!

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "ChangeFont"

        ' ******************
        ' ** DEBUG
        ' ******************
        'Exit Sub

        ' ** Font Height in Device Pixels.
        Dim FHeight As Integer
        ' ** Declared at module level so that I can DeleteObject
        ' ** on all fonts I create.
        'Dim hFontPrevious As Long
        ' ** Newly created Font.
        Dim hFont As Long
        Dim sngTmp01 As Single

        ' ** Calculate/Convert requested Font Height
        ' ** into Font's Device Coordinate space.
10110   FHeight = MulDiv(m_FontSize, lngYdpi, 72)

        ' ** We use a negative value to signify
        ' ** to the CreateFont function that we want a Glyph
        ' ** outline of this size not a bounding box.

10120   hFont = CreateFont(-FHeight, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, m_FontName) '"Verdana")

10130   lngRetVal = SendMessage(m_hWndDTP, WM_SETFONT, hFont, ByVal -1)  ' ** API Function: modWindowFunctions.

10140   If hFontPrevious <> 0 Then
10150     DeleteObject hFontPrevious
10160   End If

        ' ** Store current Font into Previous var.
10170   hFontPrevious = hFont

10180   DoEvents

        ' ** Let's compute the height of 1 line of Text with this Font
        ' ** This figure willbe added to the desired window height
        ' ** of the Calendar when the NoToday flag is TRUE.
        ' ** Store the results in the module level var m_TodayHeight.
10190   sngTmp01 = CSng(m_FontSize) / CSng(POINTSPERINCH)
10200   sngTmp01 = sngTmp01 * CSng(lngYdpi)
10210   m_TodayHeight = CLng(sngTmp01)
        'ReDraw

EXITP:
10220   Exit Sub

ERRH:
10230   Select Case ERR.Number
        Case Else
10240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10250   End Select
10260   Resume EXITP

End Sub

Public Sub ChooseColors(index As Long)
' ** Index is the color index.

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "ChooseColors"

        Dim lngNewColor As Long
        ' ** Call the Color Dialog Window.

10310   lngNewColor = aDialogColor(m_hWndDTP)  ' ** Apr 02/2004 Dialog was behind CalendarApplication.hWndAccessApp).
        ' ** User did not pick a color.
10320   If lngNewColor = -1 Then Exit Sub

        ' ** Save users selection into the corresponding property.
10330   Select Case index
        Case MCSC_BACKGROUND
10340     BackColor = lngNewColor
10350   Case MCSC_MONTHBK
10360     MonthBackColor = lngNewColor
10370   Case MCSC_TEXT
10380     ForeColor = lngNewColor
10390   Case MCSC_TITLEBK
10400     TitleBackColor = lngNewColor
10410   Case MCSC_TITLETEXT
10420     TitleForeColor = lngNewColor
10430   Case MCSC_TRAILINGTEXT
10440     TrailingForeColor = lngNewColor
10450   Case Else
10460   End Select

EXITP:
10470   Exit Sub

ERRH:
10480   Select Case ERR.Number
        Case Else
10490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10500   End Select
10510   Resume EXITP

End Sub

Public Property Let BackColor(ByVal newColor As Long)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "BackColor"

10610   m_BackColor = newColor
        ' ** If the control is allready created set the Back Color.
        'debug.print Hex(newColor)
10620   If m_hWndDTP <> 0 Then
10630     Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_BACKGROUND, ByVal newColor)  ' ** API Function: modWindowFunctions.
10640   End If

        ' ** Update custom property.
10650   Call SetProperty("BackColor", dbLongDAO, m_BackColor)

EXITP:
10660   Exit Property

ERRH:
10670   Select Case ERR.Number
        Case Else
10680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10690   End Select
10700   Resume EXITP

End Property

Public Property Get BackColor() As Long

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "BackColor"

10810   lngRetVal = SendMessage(m_hWndDTP, MCM_GETCOLOR, MCSC_BACKGROUND, 0&)  ' ** API Function: modWindowFunctions.
10820   BackColor = lngRetVal

EXITP:
10830   Exit Property

ERRH:
10840   Select Case ERR.Number
        Case Else
10850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10860   End Select
10870   Resume EXITP

End Property

Public Property Let ForeColor(ByVal newColor As Long)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "ForeColor"

10910   m_ForeColor = newColor
        ' ** If the control is allready created set the Fore Color.
10920   If m_hWndDTP <> 0 Then
10930     Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TEXT, ByVal newColor)  ' ** API Function: modWindowFunctions.
10940   End If

        ' ** Update custom property.
10950   Call SetProperty("ForeColor", dbLongDAO, m_ForeColor)

EXITP:
10960   Exit Property

ERRH:
10970   Select Case ERR.Number
        Case Else
10980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10990   End Select
11000   Resume EXITP

End Property

Public Property Get ForeColor() As Long

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "ForeColor"

11110   lngRetVal = SendMessage(m_hWndDTP, MCM_GETCOLOR, MCSC_TEXT, 0&)  ' ** API Function: modWindowFunctions.
11120   ForeColor = lngRetVal

EXITP:
11130   Exit Property

ERRH:
11140   Select Case ERR.Number
        Case Else
11150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11160   End Select
11170   Resume EXITP

End Property

Public Property Let MonthBackColor(ByVal newColor As Long)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "MonthBackColor"

11210   m_MonthBackColor = newColor
        ' ** If the control is allready created set the Month Back Color.
11220   If m_hWndDTP <> 0 Then
11230     Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_MONTHBK, ByVal newColor)  ' ** API Function: modWindowFunctions.
11240   End If

        ' ** Update custom property.
11250   Call SetProperty("MonthBackColor", dbLongDAO, m_MonthBackColor)

EXITP:
11260   Exit Property

ERRH:
11270   Select Case ERR.Number
        Case Else
11280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11290   End Select
11300   Resume EXITP

End Property

Public Property Get MonthBackColor() As Long

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "MonthBackColor"

11410   lngRetVal = SendMessage(m_hWndDTP, MCM_GETCOLOR, MCSC_MONTHBK, 0&)  ' ** API Function: modWindowFunctions.
11420   MonthBackColor = lngRetVal

EXITP:
11430   Exit Property

ERRH:
11440   Select Case ERR.Number
        Case Else
11450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11460   End Select
11470   Resume EXITP

End Property

Public Property Let TitleBackColor(ByVal newColor As Long)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "TitleBackColor"

11510   m_TitleBackColor = newColor
        ' ** If the control is allready created set the Title Back Color.
11520   If m_hWndDTP <> 0 Then
11530     Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TITLEBK, ByVal newColor)  ' ** API Function: modWindowFunctions.
11540   End If

        ' ** Update custom property.
11550   Call SetProperty("TitleBackColor", dbLongDAO, m_TitleBackColor)

EXITP:
11560   Exit Property

ERRH:
11570   Select Case ERR.Number
        Case Else
11580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11590   End Select
11600   Resume EXITP

End Property

Public Property Get TitleBackColor() As Long

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "TitleBackColor"

11710   lngRetVal = SendMessage(m_hWndDTP, MCM_GETCOLOR, MCSC_TITLEBK, 0&)  ' ** API Function: modWindowFunctions.
11720   TitleBackColor = lngRetVal

EXITP:
11730   Exit Property

ERRH:
11740   Select Case ERR.Number
        Case Else
11750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11760   End Select
11770   Resume EXITP

End Property

Public Property Let TitleForeColor(ByVal newColor As Long)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "TitleForeColor"

11810   m_TitleForeColor = newColor
        ' ** If the control is allready created set the Title Fore Color.
11820   If m_hWndDTP <> 0 Then
11830     Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TITLETEXT, ByVal newColor)  ' ** API Function: modWindowFunctions.
11840   End If

        ' ** Update custom property.
11850   Call SetProperty("TitleForeColor", dbLongDAO, m_TitleForeColor)

EXITP:
11860   Exit Property

ERRH:
11870   Select Case ERR.Number
        Case Else
11880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11890   End Select
11900   Resume EXITP

End Property

Public Property Get TitleForeColor() As Long

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "TitleForeColor"

12010   lngRetVal = SendMessage(m_hWndDTP, MCM_GETCOLOR, MCSC_TITLETEXT, 0&)  ' ** API Function: modWindowFunctions.
12020   TitleForeColor = lngRetVal

EXITP:
12030   Exit Property

ERRH:
12040   Select Case ERR.Number
        Case Else
12050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12060   End Select
12070   Resume EXITP

End Property

Public Property Let TrailingForeColor(ByVal newColor As Long)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "TrailingForeColor"

12110   m_TrailingForeColor = newColor
        ' ** If the control is allready created set the Title Fore Color.
12120   If m_hWndDTP <> 0 Then
12130     Call SendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TRAILINGTEXT, ByVal newColor)  ' ** API Function: modWindowFunctions.
12140   End If

        ' ** Update custom property.
12150   Call SetProperty("TrailingForeColor", dbLongDAO, m_TrailingForeColor)

EXITP:
12160   Exit Property

ERRH:
12170   Select Case ERR.Number
        Case Else
12180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12190   End Select
12200   Resume EXITP

End Property

Public Property Get TrailingForeColor() As Long

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "TrailingForeColor"

12310   lngRetVal = SendMessage(m_hWndDTP, MCM_GETCOLOR, MCSC_TRAILINGTEXT, 0&)  ' ** API Function: modWindowFunctions.
12320   TrailingForeColor = lngRetVal

EXITP:
12330   Exit Property

ERRH:
12340   Select Case ERR.Number
        Case Else
12350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12360   End Select
12370   Resume EXITP

End Property

Public Property Let DAYSTATE(ByVal month As Long)
' *****************************
' ** Use the SetBoldDayState function instead.
' *****************************

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "DAYSTATE"

        ' ** If MonthCal does not exist then exit.
        'ннннннннннннннннннннннннннннннннннннннннннннннннна
        'If m_hWndDTP = 0 Then Exit Property
        'а
        ''(DWORD)SNDMSG(hmc, MCM_GETMONTHRANGE, (WPARAM)(gmr), (LPARAM)(rgst))
        'Dim arrayTime(0 To 1) As SYSTEMTIME
        'Dim arrayDayStates() As MONTHDAYSTATE
        'а
        'Dim DateStart As Date
        'Dim DateEnd As Date
        'Dim VisibleMonths As Long
        'Dim StartMonth As Long
        'Dim EndMonth As Long
        'Dim ctr As Long
        '' ** Get the currently selected range.
        'VisibleMonths = SendMessage(m_hWndDTP, MCM_GETMONTHRANGE, ByVal GMR_DAYSTATE, arrayTime(0))  ' ** API Function: modWindowFunctions.
        '  If VisibleMonths = 0 Then Exit Property
        'ReDim arrayDayStates(VisibleMonths - 1)
        '' ** What is First visible month.
        'StartMonth = arrayTime(0).wMonth
        'а
        'For ctr = 0 To VisibleMonths - 1
        '  arrayDayStates(ctr).lpMONTHDAYSTATE = arr_typMDS(StartMonth).lpMONTHDAYSTATE
        'а StartMonth = StartMonth + 1
        'а ' ** Rollover after 12 months.
        'а If StartMonth = 13 Then StartMonth = 1
        'Next ctr
        'а
        'lngRetVal = SendMessage(m_hWndDTP, MCM_SETDAYSTATE, UBound(arrayDayStates) + 1, arrayDayStates(0))  ' ** API Function: modWindowFunctions.
        'а
        'ннннннннннннннннннннннннннннннннннннннннннннннннна

EXITP:
12410   Exit Property

ERRH:
12420   Select Case ERR.Number
        Case Else
12430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12440   End Select
12450   Resume EXITP

End Property

Public Function GetDAYSTATE(ByVal year As Long, ByVal month As Long) As Long

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "GetDAYSTATE"

12510   GetDAYSTATE = arr_typMDS(year, month).lpMONTHDAYSTATE

EXITP:
12520   Exit Function

ERRH:
12530   Select Case ERR.Number
        Case Else
12540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12550   End Select
12560   Resume EXITP

End Function

Public Function SetBoldDayState(lngNumYear As Long, lngNumMonth As Long, day As Long, Optional varResetMonth As Variant = False) As Boolean
' ** Turn on Bolding for specified days of the month.
' ** Use named arguments when calling this function to
' ** cut down on your typing.
' ** i.e. SetBoldDayState lngNumYear 2004, lngNumMonth:=1, day = 19.

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "SetBoldDayState"

        'ннннннннннннннннннннннннннннннннннннннннннннннннна
        'Optional One As Boolean, Optional Two As Boolean, _
        'Optional Three As Boolean, Optional Four As Boolean, _
        'Optional Five As Boolean, Optional Six As Boolean, _
        'Optional Seven As Boolean, Optional Eight As Boolean, _
        'Optional Nine As Boolean, Optional Ten As Boolean, _
        'Optional Eleven As Boolean, Optional Twelve As Boolean, _
        'Optional Thirteen As Boolean, Optional Fourteen As Boolean, _
        'Optional Fifthteen As Boolean, Optional Sixteen As Boolean, _
        'Optional Seventeen As Boolean, Optional Eighteen As Boolean, _
        'Optional NineTeen As Boolean, Optional Twenty As Boolean, _
        'Optional TwentyOne As Boolean, Optional TwentyTwo As Boolean, _
        'Optional TwentyThree As Boolean, Optional TwentyFour As Boolean, _
        'Optional TwentyFive As Boolean, Optional TwentySix As Boolean, _
        'Optional TwentySeven As Boolean, Optional TwentyEight As Boolean, _
        'Optional TwentyNine As Boolean, Optional Thirty As Boolean, _
        'Optional ThirtyOne As Boolean, Optional ThirtyTwo As Boolean)
        'ннннннннннннннннннннннннннннннннннннннннннннннннна

        ' ** Init the month we are working with.
        ' ** Only clear this var if varResetMonth = True.
12610   If varResetMonth Then
12620     arr_typMDS(lngNumYear, lngNumMonth).lpMONTHDAYSTATE = 0
12630   End If

12640   If day > 0 And day < 31 Then day = day - 1
        'arr_typMDS(lngNumMonth) = BoldDays
12650   arr_typMDS(lngNumYear, lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumYear, lngNumMonth).lpMONTHDAYSTATE Or 2 ^ day
        'ннннннннннннннннннннннннннннннннннннннннннннннннна
        'If One Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = 1 'arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 1
        'If Two Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 1
        'If Three Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 2
        'If Four Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 3
        'If Five Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 4
        'If Six Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 5
        'If Seven Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 6
        'If Eight Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 7
        'If Nine Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 8
        'If Ten Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 9
        'If Eleven Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 10
        'If Twelve Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 11
        'If Thirteen Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 12
        'If Fourteen Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 13
        'If Fifthteen Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 14
        'If Sixteen Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 15
        'If Seventeen Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 16
        'If Eighteen Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 17
        'If NineTeen Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 18
        'а
        'If Twenty Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 19
        'If TwentyOne Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 20
        'If TwentyTwo Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 21
        'If TwentyThree Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 22
        'If TwentyFour Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 23
        'If TwentyFive Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 24
        'If TwentySix Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 25
        'If TwentySeven Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 26
        'If TwentyEight Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 27
        'If TwentyNine Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 28
        'а
        'If Thirty Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 29
        'If ThirtyOne Then arr_typMDS(lngNumMonth).lpMONTHDAYSTATE = arr_typMDS(lngNumMonth).lpMONTHDAYSTATE Or 2 ^ 30
        'ннннннннннннннннннннннннннннннннннннннннннннннннна

EXITP:
12660   Exit Function

ERRH:
12670   Select Case ERR.Number
        Case Else
12680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12690   End Select
12700   Resume EXITP

End Function

Private Function GetProperty(ByVal strPropName As String, ByRef strPropValue As Variant) As Boolean
' ** Changed strPropValue as String to as Variant.

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "GetProperty"

        Dim strDefault As String

12810   Select Case strPropName
        Case "BackColor"
12820     strDefault = vbWhite
12830   Case "FontName"
12840     strDefault = "Arial"
12850   Case "FontSize"
12860     strDefault = 8
12870   Case "ForeColor"
12880     strDefault = GetSysColor(COLOR_WINDOWTEXT)
12890   Case "MonthBackColor"
12900     strDefault = vbWhite
12910   Case "MonthColumns"
12920     strDefault = 1
12930   Case MultiSelect
12940     strDefault = "0" 'false
12950   Case "NoToday"
12960     strDefault = MCS_NOTODAY
12970   Case "NoTodayCircle"
12980     strDefault = MCS_NOTODAYCIRCLE
12990   Case "OneClick"
13000     strDefault = "1" ' true
13010   Case "MCS_WEEKNUMBERS"
13020     strDefault = MCS_WEEKNUMBERS
13030   Case "MonthRows"
13040     strDefault = 1
13050   Case "TitleBackColor"
13060     strDefault = GetSysColor(COLOR_ACTIVECAPTION)
13070   Case "TitleForeColor"
13080     strDefault = vbWhite  'GetSysColor()  'vbTitleBarText
13090   Case "TrailingForeColor"
13100     strDefault = GetSysColor(COLOR_GRAYTEXT)  'SendMessage(m_hWndDTP, MCM_GETCOLOR, MCSC_TRAILINGTEXT, 0&)  ' ** API Function: modWindowFunctions.
13110   Case "WindowLocation"
          ' ** Position Window according to users Menu selections:
          ' **   a) 0 - Pop at cursor location when user activates Calendar.
          ' **   b) 1 - Where they manually move/leave it at.
          ' **   c) 2 - Centered in Access App Window.
          ' **   d) 3 - Centered on entire screen.
          ' **   e) 4 - Top Left Corner.
13120     strDefault = 0
13130   Case Else
          ' ** Not possible...I hope!
13140     strDefault = ""
13150     GetProperty = False
13160     Exit Function
13170   End Select

13180   strPropValue = GetSetting("MonthCalendarSL", "CalendarProperties", strPropName, strDefault)
13190   GetProperty = True

EXITP:
13200   Exit Function

ERRH:
13210   Select Case ERR.Number
        Case Else
13220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13230   End Select
13240   Resume EXITP

End Function

Private Sub SetProperty(ByVal strPropName As String, ByVal varPropType As Variant, ByVal varPropValue As Variant)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "SetProperty"

13310   SaveSetting "MonthCalendarSL", "CalendarProperties", strPropName, varPropValue

EXITP:
13320   Exit Sub

ERRH:
13330   Select Case ERR.Number
        Case Else
13340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13350   End Select
13360   Resume EXITP

End Sub

Public Function IsCalendar() As Boolean

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "IsCalendar"

        Dim hWndTemp As Long

        Const CLASSNAME As String = "MonthCalendar"
        Const Title As String = "Month Calendar"

        ' ** If this window already exists then exit!
13410   hWndTemp = 0
13420   hWndTemp = FindWindow(CLASSNAME, Title)
13430   If hWndTemp <> 0 Then
13440     IsCalendar = True
13450     Call Beep(12000, 1000)
13460   Else
13470     IsCalendar = False
13480   End If

EXITP:
13490   Exit Function

ERRH:
13500   Select Case ERR.Number
        Case Else
13510     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13520   End Select
13530   Resume EXITP

End Function
