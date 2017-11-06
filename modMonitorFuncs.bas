Attribute VB_Name = "modMonitorFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modMonitorFuncs"

'VGC 03/23/2017: CHANGES!

'oxX frmAccountComments
'oxX frmAccountContacts
'oxX frmAccountHideTrans2_One
'oxX frmAccountIncExpCodes
'oxX frmAccountIncExpCodes_BlockAssign
'oxX frmAccountTaxCodes
'oxX frmAccountTaxCodes_BlockAssign
'oxX frmAccountTransactions
'oxX frmAccruedIncome
'oxX frmAssetPricing
'oxX frmAssetPricing_History
'oxX frmAssets
'oxX frmCountry_Currency
'oxX frmCountryCode
'oxX frmCurrency_Country
'oxX frmFeeCalculations
'oxX frmJournal
'oxX frmJournal_Columns_TaxLot
'oxX frmLocations
'oxX frmMap_Div_Detail
'oxX frmMap_Int_Detail
'oxX frmMap_Misc_LTCL_Detail
'oxX frmMap_Misc_STCGL_Detail
'oxX frmMap_Rec_Detail
'oxX frmMap_Reinvest_DivInt_Detail
'oxX frmMap_Reinvest_Rec_Detail
'oxX frmMap_Split_Detail
'oxX frmReinvest_Dividend
'oxX frmReinvest_Interest
'oxX frmReinvest_Received
'oxX frmReinvest_Sold
'oxX frmMasterBalance
'oxX frmMenu_Maintenance
'oxX frmPortfolioModeling_Select
'oxX frmRecurringItems
'oxX frmRpt_Checks
'oxX frmRpt_IncomeExpense
'oxX frmRpt_TaxLot
'oxX frmSiteMap
'oxX frmSiteMap_Journal
'oxX frmSweeper
'oxX frmTaxLot
'oxX frmTransaction_Audit
'oxX frmUser_Add
'oxX frmUser_Password
'oxX frmXAdmin_Shortcut
'oxX modIncExpFuncs
'xX modMonitorFuncs

Private Declare Function EnumDisplayMonitors Lib "user32" _
  (ByVal hDC As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwData As Long) As Long
' ** Parameters:
' **   hdc [in]
' **     A handle to a display device context that defines the visible region of interest.
' **     If this parameter is NULL, the hdcMonitor parameter passed to the callback function will be NULL, and
' **     the visible region of interest is the virtual screen that encompasses all the displays on the desktop.
' **   lprcClip [in]
' **     A pointer to a RECT structure that specifies a clipping rectangle. The region of interest
' **     is the intersection of the clipping rectangle with the visible region specified by hdc.
' **     If hdc is non-NULL, the coordinates of the clipping rectangle are relative to the origin of the hdc.
' **     If hdc is NULL, the coordinates are virtual-screen coordinates.
' **     This parameter can be NULL if you don't want to clip the region specified by hdc.
' **   lpfnEnum [in]
' **     A pointer to a MonitorEnumProc application-defined callback function.
' **   dwData [in]
' **     Application-defined data that EnumDisplayMonitors passes directly to the MonitorEnumProc function.

Private Declare Function MonitorFromRect Lib "user32" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
' ** The MonitorFromRect function retrieves a handle to the display monitor that has the largest area of intersection with a specified rectangle
' ** Parameters:
' **   lprc [in]
' **     A pointer to a RECT structure that specifies the rectangle of interest in virtual-screen coordinates.
' **   dwFlags [in]
' **     Determines the function's return value if the rectangle does not intersect any display monitor.

Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
' ** The GetMonitorInfo function retrieves information about a display monitor.
' ** Parameters:
' **   hMonitor [in]
' **     A handle to the display monitor of interest.
' **   lpmi [out]
' **     A pointer to a MONITORINFO or MONITORINFOEX structure that receives information about the specified display monitor.
' **     You must set the cbSize member of the structure to sizeof(MONITORINFO) or sizeof(MONITORINFOEX) before calling
' **       the GetMonitorInfo function. Doing so lets the function determine the type of structure you are passing to it.
' **     The MONITORINFOEX structure is a superset of the MONITORINFO structure. It has one additional member: a string that
' **       contains a name for the display monitor. Most applications have no use for a display monitor name, and so can
' **       save some bytes by using a MONITORINFO structure.

'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long  'ALREADY ELSEWHERE!

Private Declare Function UnionRect Lib "user32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
' ** The OffsetRect function moves the specified rectangle by the specified offsets.
' ** Parameters:
' **   lprc [in,out]
' **     Pointer to a RECT structure that contains the logical coordinates of the rectangle to be moved.
' **   dx [in]
' **     Specifies the amount to move the rectangle left or right. This parameter must be a negative value to move the rectangle to the left.
' **   dy [in]
' **     Specifies the amount to move the rectangle up or down. This parameter must be a negative value to move the rectangle up.

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
' ** Changes the position and dimensions of the specified window. For a top-level window, the position and dimensions are relative to
' ** the upper-left corner of the screen. For a child window, they are relative to the upper-left corner of the parent window's client area.
' ** Parameters:
' **   hwnd [in]
' **     Type: HWND
' **     A handle to the window.
' **   x [in]
' **     Type: int
' **     The new position of the left side of the window.
' **   y [in]
' **     Type: int
' **     The new position of the top of the window.
' **   nWidth [in]
' **     Type: int
' **     The new width of the window.
' **   nHeight [in]
' **     Type: int
' **     The new height of the window.
' **   bRepaint [in]
' **     Type: BOOL
' **     Indicates whether the window is to be repainted. If this parameter is TRUE, the window receives a message.
' **     If the parameter is FALSE, no repainting of any kind occurs. This applies to the client area, the nonclient area
' **     (including the title bar and scroll bars), and any part of the parent window uncovered as a result of moving a child window.

Private Const DD_Desktop = &H1
'Private Const DD_MultiDriver = &H2
Private Const DD_Primary = &H4
Private Const DD_Mirror = &H8
'Private Const DD_VGA = &H10
'Private Const DD_Removable = &H20
'Private Const DD_ModeSpruned = &H8000000
'Private Const DD_Remote = &H4000000
'Private Const DD_Disconnect = &H2000000

Private Const DD_Active = &H1
'Private Const DD_Attached = &H2

Private Const CCHDEVICENAME As Integer = 32
Private Const CCHFORMNAME As Integer = 32

'Private Type RECT  'ALREADY ELSEWHERE!
'  Left As Long
'  Top As Long
'  Right As Long
'  Bottom As Long
'End Type

Private Type MONITORINFO
  cbSize As Long
  rcMonitor As RECT
  rcWork As RECT
  dwFlags As Long
End Type

Private Type DisplayDevice
  cb As Long
  DeviceName As String * 32
  DeviceString As String * 128
  StateFlags As Long
  DeviceID As String * 128
  DeviceKey As String * 128
End Type

Private Type POINTL
  X As Long
  Y As Long
End Type

Private Type DEVMODE
  DeviceName As String * CCHDEVICENAME
  SpecVersion As Integer
  DriverVersion As Integer
  Size As Integer
  DriverExtra As Integer
  Fields As Long
  Position As POINTL
  Scale As Integer
  Copies As Integer
  DefaultSource As Integer
  PrintQuality As Integer
  Color As Integer
  Duplex As Integer
  YResolution As Integer
  TTOption As Integer
  Collate As Integer
  FormName As String * CCHFORMNAME
  LogPixels As Integer
  BitsPerPel As Long
  PelsWidth As Long
  PelsHeight As Long
  DisplayFlags As Long
  DisplayFrequency As Long
End Type

Private Const ENUM_CURRENT_SETTINGS = -1
Private Const ENUM_REGISTRY_SETTINGS = -2

Private Declare Function EnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesA" _
  (ByVal lpDevice As String, ByVal iDevNum As Long, lpDisplayDevice As DisplayDevice, dwFlags As Long) As Long

Private Declare Function EnumDisplaySettingsEx Lib "user32" Alias "EnumDisplaySettingsExA" _
  (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As DEVMODE, dwFlags As Long) As Long

Private Declare Function EnumDisplaySettings Lib "user32.dll" Alias "EnumDisplaySettingsA" _
  (ByVal lpszDeviceName As String, ByVal iModeNum As Long, ByRef lpDevMode As DEVMODE) As Long
' ** The EnumDisplaySettings function retrieves information about one of the graphics modes for a display device.
' ** To retrieve information for all the graphics modes of a display device, make a series of calls to this function.
' ** Parameters:
' **   lpszDeviceName [in]
' **     A pointer to a null-terminated string that specifies the display device
' **     about whose graphics mode the function will obtain information.
' **     This parameter is either NULL or a DISPLAY_DEVICE.DeviceName returned from EnumDisplayDevices.
' **     A NULL value specifies the current display device on the computer on which the calling thread is running.
' **   iModeNum [in]
' **     The type of information to be retrieved. This value can be a graphics mode index or one of the following values.
' **     Graphics mode indexes start at zero. To obtain information for all of a display device's graphics modes,
' **     make a series of calls to EnumDisplaySettings, as follows: Set iModeNum to zero for the first call, and
' **     increment iModeNum by one for each subsequent call. Continue calling the function until the return value is zero.
' **     When you call EnumDisplaySettings with iModeNum set to zero, the operating system initializes and caches
' **     information about the display device. When you call EnumDisplaySettings with iModeNum set to a nonzero value,
' **     the function returns the information that was cached the last time the function was called with iModeNum set to zero.
' **   lpDevMode [out]
' **     A pointer to a DEVMODE structure into which the function stores information about the specified graphics mode.
' **     Before calling EnumDisplaySettings, set the dmSize member to sizeof(DEVMODE), and set the dmDriverExtra member
' **     to indicate the size, in bytes, of the additional space available to receive private driver data.
' **     The EnumDisplaySettings function sets values for the following five DEVMODE members:
' **       • dmBitsPerPel
' **       • dmPelsWidth
' **       • dmPelsHeight
' **       • dmDisplayFlags
' **       • dmDisplayFrequency

Private Declare Function ChangeDisplaySettingsEx Lib "user32" Alias "ChangeDisplaySettingsExA" _
  (lpszDeviceName As Any, lpDevMode As Any, ByVal hwnd As Long, ByVal dwFlags As Long, lParam As Any) As Long

Private Declare Function MonitorFromPoint Lib "user32" (ByVal ptY As Long, ByVal ptX As Long, ByVal dwFlags As Long) As Long

Private Declare Function IsMaximised Lib "user32" Alias "IsZoomed" (ByVal hwnd As Long) As Boolean

Private Declare Function IsMinimised Lib "user32" Alias "IsIconic" (ByVal hwnd As Long) As Boolean

' ** GetSystemMetrics (modWindowFunctions):
' ** The GetSystemMetrics function returns values for the primary monitor, except for SM_CXMAXTRACK
' **   and SM_CYMAXTRACK, which refer to the entire desktop. The following metrics are the same for
' **   all device drivers: SM_CXCURSOR, SM_CYCURSOR, SM_CXICON, SMCYICON. The following display
' **   capabilities are the same for all monitors: LOGPIXELSX, LOGPIXELSY, DESTOPHORZRES, DESKTOPVERTRES.
' **
' ** GetSystemMetrics also has constants that refer only to a Multiple Monitor system. SM_XVIRTUALSCREEN
' **   and SM_YVIRTUALSCREEN identify the upper-left corner of the virtual screen, SM_CXVIRTUALSCREEN
' **   and SM_CYVIRTUALSCREEN are the vertical and horizontal measurements of the virtual screen,
' **   SM_CMONITORS is the number of monitors attached to the desktop, and SM_SAMEDISPLAYFORMAT
' **   indicates whether all the monitors on the desktop have the same color format.
' **
' ** To get information about a single display monitor or all of the display monitors in a desktop,
' **   use EnumDisplayMonitors. The rectangle of the desktop window returned by GetWindowRect or
' **   GetClientRect is always equal to the rectangle of the primary monitor, for compatibility
' **   with existing applications.
' **
' ** To change the work area of a monitor, call SystemParametersInfo with SPI_SETWORKAREA and pvParam
' **   pointing to a RECT structure that is on the desired monitor. If pvParam is NULL, the work area of
' **   the primary monitor is modified. Using SPI_GETWORKAREA always returns the work area of the primary
' **   monitor. To get the work area of a monitor other than the primary monitor, call GetMonitorInfo.

'GetTpp():
'SINCE THIS GETS THE hwnd FOR WHERE ACCESS IS, IT SHOULD GIVE US THE CORRECT MONITOR.
'lngHdc = GetDC(hWndAccessApp)
'GetDeviceCaps(lngHdc, GSR_HORZRES)
'GetDeviceCaps(lngHdc, GSR_LOGPIXELSX)
'FROM DeviceCaps: (GetScreenRes(GSR_HORZRES) / GetScreenRes(GSR_LOGPIXELSX))

Private Type Monitors
  Name As String
  Handle As Long
  X As Long
  Y As Long
  Width As Long
  Height As Long
  DevString As String
  Detected As Boolean
End Type

Private PrimaryMon As Monitors
Private SecondaryMon As Monitors

' ** dwFlags:
' ** Value                 Meaning
' ** ====================  ======================================
' ** MONITORINFOF_PRIMARY  This is the primary display monitor.

Private Const MONITOR_DEFAULTTONULL    As Long = &H0  ' ** If the monitor is not found, return 0.
'Private Const MONITOR_DEFAULTTOPRIMARY As Long = &H1  ' ** If the monitor is not found, return the primary monitor.
Private Const MONITOR_DEFAULTTONEAREST As Long = &H2  ' ** If the monitor is not found, return the nearest monitor.

Private rcMonitors() As RECT  ' ** Coordinate array for all monitors.
Private rcVS         As RECT  ' ** Coordinates for Virtual Screen.

Private rc As RECT, mi As MONITORINFO
Private lngThisTpp As Long
Private lngCnt As Long, hMonitor As Long ', lngLeft As Long ', lngTop As Long
' **

Public Function LoadPosition(lngHWnd As Long, Optional varCallingForm As Variant) As Boolean
'I HAVE NO IDEA HOW THIS WILL HANDLE SIZE CHANGES ONCE THE FORM IS OPEN!
'WE NEED TO DIFFERENTIATE BETWEEN MONITOR #1 AND MONITOR #2; IT SHOULDN'T BE NEEDED ON MONITOR#1!
'NEED FUNCTION TO QUICKLY TELL WHICH MONITOR WE'RE ON!
'BE ON THE LOOKOUT FOR FORMS FLIPPING FROM #2 TO #1!
'SIGNAL WHETHER ACCESS IS MAXIMIZED!
'KEEP IN MIND THERE MAY BE MORE THAN 2 MONITORS!

100   On Error GoTo ERRH

        Const THIS_PROC As String = "LoadPosition"

        Dim lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
        Dim lngWin_Left As Long, lngWin_Top As Long, lngWin_Width As Long, lngWin_Height As Long
        Dim strCallingForm As String
        Dim varTmp00 As Variant, lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long, lngTmp05 As Long, lngTmp06 As Long
        Dim blnRetVal As Boolean

110     blnRetVal = True

        ' ** I need to make this adaptable.
        ' ** ICP's appears to be 15 even though GetTpp() returns 20.
        ' ** It may be that I'm getting only the laptop's Tpp, rather than the external monitor's.
120     lngThisTpp = 15&

130     Select Case IsMissing(varCallingForm)
        Case True
140       strCallingForm = vbNullString
150     Case False
160       strCallingForm = varCallingForm
170     End Select

180     If strCallingForm <> vbNullString Then
          ' ** Variables are fed empty, then populated ByRef.
190       GetFormDimensions Forms(strCallingForm), lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.
          'Debug.Print "'lngFrm_Left: " & CStr(lngFrm_Left) & "  lngFrm_Top: " & CStr(lngFrm_Top) & "  lngFrm_Width: " & CStr(lngFrm_Width) & "  lngFrm_Height: " & CStr(lngFrm_Height) & "  " & THIS_PROC & "()"
200     End If

        ' ** Obtain the window rectangle.
210     GetWindowRect lngHWnd, rc  ' ** API Function: Above.

        ' ** Find the monitor closest to window rectangle.
220     hMonitor = MonitorFromRect(rc, MONITOR_DEFAULTTONEAREST)  ' ** API Function: Above.

        ' ** Get info about monitor coordinates and working area.
230     mi.cbSize = Len(mi)
240     GetMonitorInfo hMonitor, mi  ' ** API Function: Above.

        ' ** Variables are fed empty, then populated ByRef.
250     GetAppDimensions lngWin_Left, lngWin_Top, lngWin_Width, lngWin_Height  ' ** Module Function: modWindowFunctions.
        ' ** My monitor, maximized:
        ' **   -120,  -120,  21840,  13290

        ' ** Calculate Left.
260     lngTmp01 = rcVS.Left  ' ** Left edge of monitor #2.
270     lngTmp02 = ((lngWin_Left / lngThisTpp) - rcVS.Left)  ' ** Distance between Access window and left edge of monitor #2.
280     lngTmp03 = (((lngWin_Width / lngThisTpp) - (lngFrm_Width / lngThisTpp)) / 2&)  ' ** Access window width minus form width divided by 2, to center form in Access window.
290     lngTmp01 = (lngTmp01 + lngTmp02 + lngTmp03)  '-935
        'Debug.Print "'lngTmp01: " & CStr(lngTmp01)

        ' ** Calculate Top.
        ' ** Here, use my values to figure out where the Top should really be, since
        ' ** forms are seldom absolutely centered, and usually ride high on the screen.
        ' ** Most are just 66.66& percent of centered top.
300     varTmp00 = DLookup("[real_top]", "qryVBComponent_Monitor_08", "[frm_name] = '" & strCallingForm & "'")
310     If IsNull(varTmp00) = True Then varTmp00 = 1
320     lngTmp02 = rcVS.Top  ' ** Top edge of monitor #2
330     lngTmp03 = ((lngWin_Top / lngThisTpp) - rcVS.Top)  ' ** Distance between Access window and top edge of monitor #2.
340     lngTmp04 = (((lngWin_Height / lngThisTpp) - (lngFrm_Height / lngThisTpp)) / 2&)  ' ** Access Window height minus form height divided by 2, to center form in Access window.
350     lngTmp05 = (lngTmp02 + lngTmp03 + lngTmp04)
        ' ** Only apply percentage to that portion between form and Access top edge.
360     lngTmp02 = (lngTmp02 + lngTmp03 + (lngTmp04 * varTmp00))  '-1080
        'Debug.Print "'lngTmp02: " & CStr(lngTmp02)

        ' ** I don't understand what makes these forms,below, different!
370     If strCallingForm = "frmMasterBalance" Then
380       lngTmp01 = (lngFrm_Left / lngThisTpp)
390       lngTmp02 = (8& * lngThisTpp)
400     End If

410     If strCallingForm = "frmRpt_IncomeExpense" Then
420       lngTmp01 = (lngFrm_Left / lngThisTpp)
430       lngTmp02 = (8& * lngThisTpp)
440     End If

450     If strCallingForm = "frmRpt_TaxLot" Then
460       lngTmp01 = (lngFrm_Left / lngThisTpp)
470       lngTmp02 = (8& * lngThisTpp)
480     End If

490     If strCallingForm = "frmUser_Add" Then
500       If Forms(strCallingForm).frm_top = 0 Then
510         lngTmp02 = (lngTmp02 - 39&)  ' ** Just a shade higher.
520         lngTmp06 = lngTmp02
530         Forms(strCallingForm).frm_top = lngTmp06
540       Else
550         lngTmp06 = Forms(strCallingForm).frm_top
560         lngTmp02 = lngTmp06
570       End If
580     End If

590     If strCallingForm = "frmUser_Password" Then
          ' ** Ah! If it's Zero, it's the first time in,
          ' ** and Help isn't showing, so save this Top value.
600       If Forms(strCallingForm).frm_top = 0 Then
610         lngTmp02 = (lngTmp02 - 37&)  ' ** Just a shade higher.
620         lngTmp06 = lngTmp02
630         Forms(strCallingForm).frm_top = lngTmp06
640       Else
            ' ** Thereafter, use the same Top value.
650         lngTmp06 = Forms(strCallingForm).frm_top
660         lngTmp02 = lngTmp06
670       End If
680     End If

        ' ** Calculate Width.
690     lngTmp03 = (lngFrm_Width / lngThisTpp)  ' ** Form width.

        ' ** Calculate Height.
700     lngTmp04 = (lngFrm_Height / lngThisTpp)  ' ** Form height.

710     MoveWindow lngHWnd, lngTmp01, lngTmp02, lngTmp03, lngTmp04, 0     ' ** API Function: Above.
        'Debug.Print "'lngTmp01: " & CStr(lngTmp01) & "  lngTmp02: " & CStr(lngTmp02) & "  lngTmp03: " & CStr(lngTmp03) & "  lngTmp04: " & CStr(lngTmp04) & "  " & THIS_PROC & "()"

EXITP:
720     LoadPosition = blnRetVal
730     Exit Function

ERRH:
740     blnRetVal = False
750     Select Case ERR.Number
        Case Else
760       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
770     End Select
780     Resume EXITP

End Function

Public Sub SavePosition(hwnd As Long)
' ** This isn't used anywhere.

800   On Error GoTo ERRH

        Const THIS_PROC As String = "SavePosition"

        Dim rc As RECT

        ' ** Save position in pixel units.
810     GetWindowRect hwnd, rc  ' ** API Function: Above.

820     Interaction.SaveSetting "Multi Monitor Demo", "Position", "Left", rc.Left
830     Interaction.SaveSetting "Multi Monitor Demo", "Position", "Top", rc.Top

EXITP:
840     Exit Sub

ERRH:
850     Select Case ERR.Number
        Case Else
860       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
870     End Select
880     Resume EXITP

End Sub

Public Function GetMonitorCount() As Long

900   On Error GoTo ERRH

        Const THIS_PROC As String = "GetMonitorCount"

        Dim lngRetVal As Long

910     EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc2, lngRetVal  ' ** API Function: Above.

EXITP:
920     GetMonitorCount = lngRetVal
930     Exit Function

ERRH:
940     lngRetVal = 0&
950     Select Case ERR.Number
        Case Else
960       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
970     End Select
980     Resume EXITP

End Function

Public Function IsAppMax() As Boolean

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "IsAppMax"

        Dim strTmp01 As String
        Dim blnRetVal As Boolean

1010    blnRetVal = False

1020    strTmp01 = AppWindowState(Application)  ' ** Function: Below.
1030    Select Case strTmp01
        Case "Restore"
1040      blnRetVal = False
1050    Case "Maximize"
1060      blnRetVal = True
1070    Case "Minimize"
1080      blnRetVal = False
1090    End Select

EXITP:
1100    IsAppMax = blnRetVal
1110    Exit Function

ERRH:
1120    blnRetVal = False
1130    Select Case ERR.Number
        Case Else
1140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1150    End Select
1160    Resume EXITP

End Function

Public Function GetMonitorNum() As Long

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "GetMonitorNum"

        Dim lngWin_Left As Long, lngWin_Top As Long, lngWin_Width As Long, lngWin_Height As Long
        Dim dmPrimary As DEVMODE
        Dim dmSecondary As DEVMODE
        Dim dmTemp As DEVMODE
        Dim blnIsMaximized As Boolean
        Dim lngRetVal As Long

        'Const CDS_UPDATEREGISTRY As Long = &H1
        'Const CDS_NORESET         As Long = &H10000000
        'Const CDS_RESET           As Long = &H40000000
        Const DM_BITSPERPEL       As Long = &H40000
        Const DM_PELSWIDTH        As Long = &H80000
        Const DM_PELSHEIGHT       As Long = &H100000
        Const DM_POSITION         As Long = &H20
        Const DM_DISPLAYFREQUENCY As Long = &H400000

        Const szPrimaryDisplay As String = "\\.\DISPLAY1" '"[URL='file://\\.\DISPLAY1']\\.\DISPLAY1[/URL]"
        Const szSecondaryDisplay As String = "\\.\DISPLAY2" '"[URL='file://\\.\DISPLAY2']\\.\DISPLAY2[/URL]"
        ' ** These are my device names:
        ' **   DispDev.DeviceName 0: \\.\DISPLAY1
        ' **   DispDev.DeviceName 1: \\.\DISPLAY2
        ' **   DispDev.DeviceName 2: \\.\DISPLAY3
        ' **   DispDev.DeviceName 3: \\.\DISPLAY4
        ' **   DispDev.DeviceName 4: \\.\DISPLAYV1
        ' **   DispDev.DeviceName 5: \\.\DISPLAYV2
        ' **   DispDev.DeviceName 6: \\.\DISPLAYV3
        ' **   DispDev.DeviceName 7: \\.\DISPLAYV4

1210    lngRetVal = 0&

        ' ** I need to make this adaptable.
        ' ** ICP's appears to be 15 even though GetTpp() returns 20.
        ' ** It may be that I'm getting only the laptop's Tpp, rather than the external monitor's.
        ' ** Perhaps a checkbox on Options as to whether to use 15 or the GetTpp() number.
        ' ** Ctrl-Shift-T shows it to anyone on frmMenu_Maintenance.
1220    lngThisTpp = 15&

1230    dmPrimary.Size = Len(dmPrimary)
1240    dmTemp.Size = Len(dmTemp)
1250    dmSecondary.Size = Len(dmSecondary)

1260    If EnumDisplaySettings(szPrimaryDisplay, ENUM_CURRENT_SETTINGS, dmTemp) = False Then
1270      lngRetVal = 0&
1280      MsgBox "Primary Settings couldn't be enumerated.", vbCritical + vbOKOnly, "Unable To Retrieve Monitor Info"
1290    Else

1300      With dmPrimary
1310        .BitsPerPel = dmTemp.BitsPerPel
1320        .PelsHeight = dmTemp.PelsHeight
1330        .PelsWidth = dmTemp.PelsWidth
1340        .DisplayFrequency = dmTemp.DisplayFrequency
1350        .Fields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
1360        With .Position
1370          .X = dmTemp.Position.X
1380          .Y = dmTemp.Position.Y
1390        End With
1400        .Fields = .Fields Or DM_POSITION
1410      End With
1420      If dmPrimary.DisplayFrequency <> 0 Then dmPrimary.Fields = dmPrimary.Fields Or DM_DISPLAYFREQUENCY

1430      If EnumDisplaySettings(szSecondaryDisplay, ENUM_CURRENT_SETTINGS, dmTemp) = False Then
1440        lngRetVal = 0&
1450        MsgBox "Secondary Settings couldn't be enumerated.", vbCritical + vbOKOnly, "Unable To Retrieve Monitor Info"
1460      Else

1470        With dmSecondary
1480          .BitsPerPel = dmTemp.BitsPerPel
1490          .PelsHeight = dmTemp.PelsHeight
1500          .PelsWidth = dmTemp.PelsWidth
1510          .DisplayFrequency = dmTemp.DisplayFrequency
1520          .Fields = (DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT)
1530          If .DisplayFrequency <> 0 Then .Fields = .Fields Or DM_DISPLAYFREQUENCY
1540          With .Position
1550            .X = dmTemp.Position.X
1560            .Y = dmTemp.Position.Y
1570          End With
1580          .Fields = .Fields Or DM_POSITION
1590        End With

            ' ** Variables are fed empty, then populated ByRef.
1600        GetAppDimensions lngWin_Left, lngWin_Top, lngWin_Width, lngWin_Height  ' ** Module Function: modWindowFunctions.

1610        blnIsMaximized = IsAppMax  ' ** Function: Above.

1620        Select Case blnIsMaximized
            Case True
              ' ** Monitor #1:
              ' **   lngWin_Left: -120,  lngWin_Top: -120,  lngWin_Width: 29040,  lngWin_Height: 15840  MAXIMIZED!
              ' ** Monitor #2:
              ' **   lngWin_Left: -14145,  lngWin_Top: -16320,  lngWin_Width: 29040,  lngWin_Height: 16440  MAXIMIZED!
1630          If ((lngWin_Left / lngThisTpp) < dmPrimary.Position.X) And _
                  ((lngWin_Left / lngThisTpp) >= (dmSecondary.Position.X - (240& / lngThisTpp))) And _
                  ((lngWin_Top / lngThisTpp) < dmPrimary.Position.Y) And _
                  ((lngWin_Top / lngThisTpp) >= (dmSecondary.Position.Y - (240& / lngThisTpp))) Then
                'IF IT'S #2, WIN_LEFT WILL ALWAYS BE LESS THAN #1'S X.
                'BUT IT WON'T NECESSARILY BE LESS THAN #2'S X!
                'I GET -120 BOTH HERE AND ON MY OWN MONITOR, SO LET'S GO
                'WITH THAT AND ADD IT TO THE CRITERIA. (WITH A LITTLE EXTRA!)
1640            lngRetVal = 2&
1650          ElseIf (((lngWin_Left / lngThisTpp) >= (dmPrimary.Position.X - (240& / lngThisTpp))) And _
                  (((lngWin_Left / lngThisTpp) + (lngWin_Width / lngThisTpp)) <= ((dmPrimary.Position.X + dmPrimary.PelsWidth) + (240& / lngThisTpp)))) And _
                  (((lngWin_Top / lngThisTpp) >= (dmPrimary.Position.Y - (240& / lngThisTpp))) And _
                  (((lngWin_Top / lngThisTpp) + (lngWin_Height / lngThisTpp)) <= ((dmPrimary.Position.Y + dmPrimary.PelsHeight) + (240& / lngThisTpp)))) Then
1660            lngRetVal = 1&
1670          End If
1680        Case False
              ' ** Monitor #1:
              ' **   lngWin_Left: 2175,  lngWin_Top: 2265,  lngWin_Width: 25110,  lngWin_Height: 11400
              ' ** Monitor #2:
              ' **   lngWin_Left: -11850,  lngWin_Top: -13935,  lngWin_Width: 25110,  lngWin_Height: 11400
1690          If ((lngWin_Left / lngThisTpp) < dmPrimary.Position.X) And ((lngWin_Left / lngThisTpp) > dmSecondary.Position.X) And _
                  ((lngWin_Top / lngThisTpp) < dmPrimary.Position.Y) And ((lngWin_Top / lngThisTpp) > dmSecondary.Position.Y) Then
1700            lngRetVal = 2&
1710          ElseIf (((lngWin_Left / lngThisTpp) >= dmPrimary.Position.X) And _
                  (((lngWin_Left / lngThisTpp) + (lngWin_Width / lngThisTpp)) <= (dmPrimary.Position.X + dmPrimary.PelsWidth))) And _
                  (((lngWin_Top / lngThisTpp) >= dmPrimary.Position.Y) And _
                  (((lngWin_Top / lngThisTpp) + (lngWin_Height / lngThisTpp)) <= (dmPrimary.Position.Y + dmPrimary.PelsHeight))) Then
1720            lngRetVal = 1&
1730          End If
1740        End Select

            ' ** Monitor #1:
            'lngWin_Left: 2175,  lngWin_Top: 2265,  lngWin_Width: 25110,  lngWin_Height: 11400
            'lngWin_Left: -120,  lngWin_Top: -120,  lngWin_Width: 29040,  lngWin_Height: 15840  MAXIMIZED!
            'Position.x: 0
            'Position.y: 0
            'PelsHeight: 1080
            'PelsWidth: 1920

            ' ** Monitor #2:
            'lngWin_Left: -11850,  lngWin_Top: -13935,  lngWin_Width: 25110,  lngWin_Height: 11400
            'lngWin_Left: -14145,  lngWin_Top: -16320,  lngWin_Width: 29040,  lngWin_Height: 16440  MAXIMIZED!
            'Position.x: -935
            'Position.y: -1080
            'PelsHeight: 1080
            'PelsWidth: 1920

1750      End If
1760    End If

        'With dmTemp
        '  Debug.Print "'BitsPerPel: " & .BitsPerPel
        '  Debug.Print "'Collate: " & .Collate
        '  Debug.Print "'Color: " & .Color
        '  Debug.Print "'Copies: " & .Copies
        '  Debug.Print "'DefaultSource: " & .DefaultSource
        '  Debug.Print "'DeviceName: " & .DeviceName
        '  Debug.Print "'DisplayFlags: " & .DisplayFlags
        '  Debug.Print "'DisplayFrequency: " & .DisplayFrequency
        '  Debug.Print "'DriverExtra: " & .DriverExtra
        '  Debug.Print "'DriverVersion: " & .DriverVersion
        '  Debug.Print "'Duplex: " & .Duplex
        '  Debug.Print "'Fields: " & .Fields
        '  Debug.Print "'FormName: " & .FormName
        '  Debug.Print "'LogPixels: " & .LogPixels
        '  Debug.Print "'PelsHeight: " & .PelsHeight
        '  Debug.Print "'PelsWidth: " & .PelsWidth
        '  Debug.Print "'Position.x: " & .Position.x
        '  Debug.Print "'Position.y: " & .Position.y
        '  Debug.Print "'PrintQuality: " & .PrintQuality
        '  Debug.Print "'Scale: " & .Scale
        '  Debug.Print "'Size: " & .Size
        '  Debug.Print "'SpecVersion: " & .SpecVersion
        '  Debug.Print "'TTOption: " & .TTOption
        '  Debug.Print "'YResolution: " & .YResolution
        'End With

        'With dmTemp
        '  Debug.Print "'BitsPerPel: " & .BitsPerPel
        '  Debug.Print "'Collate: " & .Collate
        '  Debug.Print "'Color: " & .Color
        '  Debug.Print "'Copies: " & .Copies
        '  Debug.Print "'DefaultSource: " & .DefaultSource
        '  Debug.Print "'DeviceName: " & .DeviceName
        '  Debug.Print "'DisplayFlags: " & .DisplayFlags
        '  Debug.Print "'DisplayFrequency: " & .DisplayFrequency
        '  Debug.Print "'DriverExtra: " & .DriverExtra
        '  Debug.Print "'DriverVersion: " & .DriverVersion
        '  Debug.Print "'Duplex: " & .Duplex
        '  Debug.Print "'Fields: " & .Fields
        '  Debug.Print "'FormName: " & .FormName
        '  Debug.Print "'LogPixels: " & .LogPixels
        '  Debug.Print "'PelsHeight: " & .PelsHeight
        '  Debug.Print "'PelsWidth: " & .PelsWidth
        '  Debug.Print "'Position.x: " & .Position.x
        '  Debug.Print "'Position.y: " & .Position.y
        '  Debug.Print "'PrintQuality: " & .PrintQuality
        '  Debug.Print "'Scale: " & .Scale
        '  Debug.Print "'Size: " & .Size
        '  Debug.Print "'SpecVersion: " & .SpecVersion
        '  Debug.Print "'TTOption: " & .TTOption
        '  Debug.Print "'YResolution: " & .YResolution
        'End With

        ' ** Monitor #1:
        'BitsPerPel: 32
        'Collate: 0
        'Color: 0
        'Copies: 0
        'DefaultSource: 0
        'DeviceName: cdd
        'DisplayFlags: 0
        'DisplayFrequency: 60
        'DriverExtra: 0
        'DriverVersion: 1025
        'Duplex: 0
        'Fields: 8257696
        'FormName:
        'LogPixels: 0
        'PelsHeight: 1080
        'PelsWidth: 1920
        'Position.x: 0
        'Position.y: 0
        'PrintQuality: 0
        'Scale: 0
        'Size: 124
        'SpecVersion: 1025
        'TTOption: 0
        'YResolution: 0

        ' ** Monitor #2:
        'BitsPerPel: 32
        'Collate: 0
        'Color: 0
        'Copies: 0
        'DefaultSource: 0
        'DeviceName: cdd
        'DisplayFlags: 0
        'DisplayFrequency: 60
        'DriverExtra: 0
        'DriverVersion: 1025
        'Duplex: 0
        'Fields: 8257696
        'FormName:
        'LogPixels: 0
        'PelsHeight: 1080
        'PelsWidth: 1920
        'Position.x: -935
        'Position.y: -1080
        'PrintQuality: 0
        'Scale: 0
        'Size: 124
        'SpecVersion: 1025
        'TTOption: 0
        'YResolution: 0

EXITP:
1770    GetMonitorNum = lngRetVal
1780    Exit Function

ERRH:
1790    lngRetVal = 0&
1800    Select Case ERR.Number
        Case Else
1810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1820    End Select
1830    Resume EXITP

End Function

Public Function GetAppSpecs() As Boolean

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "GetAppSpecs"

        Dim lngWin_Left As Long, lngWin_Top As Long, lngWin_Width As Long, lngWin_Height As Long
        Dim blnRetVal As Boolean

        ' ** Variables are fed empty, then populated ByRef.
1910    GetAppDimensions lngWin_Left, lngWin_Top, lngWin_Width, lngWin_Height  ' ** Module Function: modWindowFunctions.

        'Debug.Print "'lngWin_Left: " & CStr(lngWin_Left) & ",  lngWin_Top: " & CStr(lngWin_Top) & ",  lngWin_Width: " & CStr(lngWin_Width) & ",  lngWin_Height: " & CStr(lngWin_Height)

        'WHEN MAXIMIZED, WINDOW EXTENDS BEYOND EDGES!
        'If ((lngWin_Left / 15) < mon1.Position.x) And ((lngWin_Left / 15) > mon2.Position.x) And _
        '    ((lngWin_Top / 15) < mon1.Position.y) And  ((lngWin_Top / 15) > mon2.Position.y) Then
        '  'THIS IS RUNNING ON MONITOR #2.
        'ElseIf (((lngWin_Left / 15) >= mon1.Position.x) And (((lngWin_Left / 15) + (lngWin_Width / 15)) <= (mon1.Position.x + mon1.PelsWidth))) And _
        '    ((lngWin_Top / 15) >= mon1.Position.y) And  (((lngWin_Top / 15) + (lngWin_Height / 15)) <= (mon1.Position.y + mon1.PelsHeight))) Then
        '  'THIS IS RUNN8ING ON MONITOR #1.
        'End If

        ' ** Monitor #1:
        'lngWin_Left: 2175,  lngWin_Top: 2265,  lngWin_Width: 25110,  lngWin_Height: 11400
        'lngWin_Left: -120,  lngWin_Top: -120,  lngWin_Width: 29040,  lngWin_Height: 15840  MAXIMIZED!
        'Position.x: 0
        'Position.y: 0
        'PelsHeight: 1080
        'PelsWidth: 1920

        ' ** Monitor #2:
        'lngWin_Left: -11850,  lngWin_Top: -13935,  lngWin_Width: 25110,  lngWin_Height: 11400
        'lngWin_Left: -14145,  lngWin_Top: -16320,  lngWin_Width: 29040,  lngWin_Height: 16440  MAXIMIZED!
        'Position.x: -935
        'Position.y: -1080
        'PelsHeight: 1080
        'PelsWidth: 1920

EXITP:
1920    GetAppSpecs = blnRetVal
1930    Exit Function

ERRH:
1940    blnRetVal = False
1950    Select Case ERR.Number
        Case Else
1960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1970    End Select
1980    Resume EXITP

End Function

Public Function GetDisplays() As Boolean

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "GetDisplays"

        Dim DispDev As DisplayDevice, MonDev As DisplayDevice  ' ** Holds device info/monitor info.
        Dim DispDevInd As Long, MonDevInd As Long              ' ** Index for the devices.
        Dim MonMode As DEVMODE                                 ' ** Holds mode information for each monitor.
        Dim hMonitor As Long                                   ' ** Holds handle to the correct monitor context.
        Dim MonInfo As MONITORINFO
        Dim blnRetVal As Boolean, blnMsg As Boolean

2010    blnRetVal = True
2020    blnMsg = False

        ' ** Initializations.
2030    DispDev.cb = Len(DispDev)
2040    MonDev.cb = Len(MonDev)
2050    DispDevInd = 0: MonDevInd = 0
2060    PrimaryMon.Detected = False
2070    SecondaryMon.Detected = False

2080    Do While EnumDisplayDevices(vbNullString, DispDevInd, DispDev, 0) <> 0  ' ** Enumerate the graphics cards.

2090      If Not CBool(DispDev.StateFlags And DD_Mirror) Then

            ' ** If it is real.
2100        Do While EnumDisplayDevices(DispDev.DeviceName, MonDevInd, MonDev, 0) <> 0  ' ** Iterate to the correct MonDev.
2110          If CBool(MonDev.StateFlags And DD_Active) Then Exit Do
2120          MonDevInd = MonDevInd + 1
2130        Loop

            ' ** If the device string is empty then its a default monitor.
2140        If cCstr(MonDev.DeviceString) = "" Then
2150          EnumDisplayDevices DispDev.DeviceName, 0, MonDev, 0
2160          If cCstr(MonDev.DeviceString) = "" Then MonDev.DeviceString = "Default Monitor"
2170        End If

            ' ** Get information about the display's position and the current display mode.
2180        MonMode.Size = Len(MonMode)
2190        If EnumDisplaySettingsEx(DispDev.DeviceName, ENUM_CURRENT_SETTINGS, MonMode, 0) = 0 Then
2200          EnumDisplaySettingsEx DispDev.DeviceName, ENUM_REGISTRY_SETTINGS, MonMode, 0
2210        End If

            ' ** Get the monitor handle and workspace.
2220        MonInfo.cbSize = Len(MonInfo)
2230        If CBool(DispDev.StateFlags And DD_Desktop) Then
              ' ** Display is enabled. Only enabled displays have a monitor handle.
2240          hMonitor = MonitorFromPoint(MonMode.Position.X, MonMode.Position.Y, MONITOR_DEFAULTTONULL)
2250          If hMonitor <> 0 Then
2260            GetMonitorInfo hMonitor, MonInfo
2270          End If
2280        End If

2290      End If

2300      If CBool(DispDev.StateFlags And DD_Desktop) Then    ' ** If it is an active monitor.
2310        If CBool(DispDev.StateFlags And DD_Primary) Then  ' ** If it is the primary.
2320          With PrimaryMon
2330            If MonDev.DeviceName <> "" Then .Name = cCstr(MonDev.DeviceName) Else .Name = cCstr(DispDev.DeviceName)
2340            Select Case blnMsg
                Case True
2350              MsgBox "1: " & .Name
2360            Case False
2370              Debug.Print "'MON #1: " & .Name
2380            End Select
2390            .Detected = True
2400            .Handle = hMonitor
2410            .DevString = cCstr(MonDev.DeviceString) & " on " & cCstr(DispDev.DeviceString)
2420            Select Case blnMsg
                Case True
2430              MsgBox .DevString
2440            Case False
2450              Debug.Print "'  " & .DevString
2460              Debug.Print "'  " & CStr(.Handle)
2470            End Select
2480            .X = MonMode.Position.X
2490            .Y = MonMode.Position.Y
2500            .Width = MonMode.PelsWidth
2510            .Height = MonMode.PelsHeight
2520            Debug.Print "'Left: " & CStr(.X)
2530            Debug.Print "'Top: " & CStr(.Y)
2540            Debug.Print "'Width: " & CStr(.Width)
2550            Debug.Print "'Height: " & CStr(.Height)
2560          End With
2570        Else
2580          If Not SecondaryMon.Detected Then  ' ** If it is a secondary (only do one).
2590            With SecondaryMon
2600              If MonDev.DeviceName <> "" Then .Name = cCstr(MonDev.DeviceName) Else .Name = cCstr(DispDev.DeviceName)
2610              Select Case blnMsg
                  Case True
2620                MsgBox "2: " & .Name
2630              Case False
2640                Debug.Print "'MON #2: " & .Name
2650              End Select
2660              .Detected = True
2670              .Handle = hMonitor
2680              .DevString = cCstr(MonDev.DeviceString) & " on " & cCstr(DispDev.DeviceString)
2690              Select Case blnMsg
                  Case True
2700                MsgBox .DevString
2710              Case False
2720                Debug.Print "'  " & .DevString
2730                Debug.Print "'  " & CStr(.Handle)
2740              End Select
2750              .X = MonMode.Position.X
2760              .Y = MonMode.Position.Y
2770              .Width = MonMode.PelsWidth
2780              .Height = MonMode.PelsHeight
2790              Debug.Print "'Left: " & CStr(.X)
2800              Debug.Print "'Top: " & CStr(.Y)
2810              Debug.Print "'Width: " & CStr(.Width)
2820              Debug.Print "'Height: " & CStr(.Height)
2830            End With
2840          End If
2850        End If
2860      End If

          'Debug.Print "'DispDev.DeviceName " & CStr(DispDevInd) & ": " & DispDev.DeviceName

2870      DispDevInd = DispDevInd + 1  ' ** Next graphics card.

2880    Loop

        'MON #1: \\.\DISPLAY1\Monitor0
        '  LCD 1920x1080 on NVIDIA Quadro FX 880M
        '  65537
        'Left: 0
        'Top: 0
        'Width: 1920
        'Height: 1080

        'MON #2: \\.\DISPLAY2\Monitor0
        '  Generic PnP Monitor on NVIDIA Quadro FX 880M
        '  65539
        'Left: -935
        'Top: -1080
        'Width: 1920
        'Height: 1080

        'GetDC(hWndAccessApp)
        '1879118415
        'hWndAccessApp
        '4589840

        'DispDev.DeviceName 0: \\.\DISPLAY1
        'DispDev.DeviceName 1: \\.\DISPLAY2
        'DispDev.DeviceName 2: \\.\DISPLAYV1
        'DispDev.DeviceName 3: \\.\DISPLAYV2
        'DispDev.DeviceName 4: \\.\DISPLAYV3

EXITP:
2890    GetDisplays = blnRetVal
2900    Exit Function

ERRH:
2910    blnRetVal = False
2920    Select Case ERR.Number
        Case Else
2930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2940    End Select
2950    Resume EXITP

End Function

Private Function cCstr(strInput As String) As String

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cCstr"

        Dim strChar As String
        Dim lngX As Long
        Dim strRetVal As String

3010    strRetVal = vbNullString

3020    For lngX = 1 To Len(strInput)
3030      strChar = Mid(strInput, lngX, 1)
3040      If strChar <> vbNullChar And strChar <> vbNullString Then
3050        strRetVal = strRetVal & strChar
3060      End If
3070    Next

EXITP:
3080    cCstr = strRetVal
3090    Exit Function

ERRH:
3100    strRetVal = vbNullString
3110    Select Case ERR.Number
        Case Else
3120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3130    End Select
3140    Resume EXITP

End Function

Private Function AppWindowState(appVar As Application, Optional varSet As Variant) As String
' ** AppWindowState() returns and/or sets the app's window state.

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "AppWindowState"

        Dim blnVisible As Boolean, strSet As String
        Dim strRetVal As String

3210    Select Case IsMissing(varSet)
        Case True
3220      strSet = "Read"
3230    Case False
3240      Select Case IsNull(varSet)
          Case True
3250        strSet = "Read"
3260      Case False
3270        strSet = varSet
3280      End Select
3290    End Select

3300    With appVar
3310      strRetVal = "Restore"
3320      If IsMaximised(.hWndAccessApp) Then strRetVal = "Maximize"
3330      If IsMinimised(.hWndAccessApp) Then strRetVal = "Minimize"
3340      If strSet <> "Read" Then
3350        If strSet <> strRetVal Then
3360          blnVisible = .Visible
3370          If blnVisible = False Then .Visible = True
3380          Select Case strSet
              Case "Maximize"
3390            Call .RunCommand(acCmdAppMaximize)
3400          Case "Minimize"
3410            Call .RunCommand(acCmdAppMinimize)
3420          Case "Restore"
3430            Call .RunCommand(acCmdAppRestore)
3440          End Select
3450          If blnVisible = False Then .Visible = False
3460        End If
3470      End If
3480    End With

EXITP:
3490    AppWindowState = strRetVal
3500    Exit Function

ERRH:
3510    strRetVal = RET_ERR
3520    Select Case ERR.Number
        Case Else
3530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3540    End Select
3550    Resume EXITP

End Function

Public Function FormLoad(frm As Access.Form)
' ** In Form_Open() or Form_Load() of any form using DoCmd.MoveSize.

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "FormLoad"

3610    EnumMonitors frm
3620    LoadPosition frm.hwnd

EXITP:
3630    Exit Function

ERRH:
3640    Select Case ERR.Number
        Case Else
3650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3660    End Select
3670    Resume EXITP

End Function

Private Sub FormUnload(Cancel As Integer, frm As Access.Form)
' ** This saves to the Registry. Let's not!

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormUnload"

        'SavePosition frm.hwnd

EXITP:
3710    Exit Sub

ERRH:
3720    Select Case ERR.Number
        Case Else
3730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3740    End Select
3750    Resume EXITP

End Sub

Public Function EnumMonitors(frm As Access.Form) As Long

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "EnumMonitors"

        Dim lngX As Long
        Dim lngRetVal As Long

3810    lngRetVal = 1&

3820    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, lngX  ' ** Function: Below.

        ' ** lngX gives the number of monitors.
        ' ** rcVS is the size of the Virtual Screen, which I think is all screens combined into one big one.

3830    lngCnt = lngX

EXITP:
3840    EnumMonitors = lngRetVal
3850    Exit Function

ERRH:
3860    lngRetVal = 0&
3870    Select Case ERR.Number
        Case Else
3880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3890    End Select
3900    Resume EXITP

End Function

Private Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, dwData As Long) As Long

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "MonitorEnumProc"

        Dim lngRetVal As Long

4010    lngRetVal = 1&  ' ** Continue.

4020    ReDim Preserve rcMonitors(dwData)

4030    rcMonitors(dwData) = lprcMonitor

        ' ** Merge all monitors together to get the virtual screen coordinates.
4040    UnionRect rcVS, rcVS, lprcMonitor  ' ** API Function: Above.

4050    dwData = dwData + 1&  ' ** Increase monitor count.

EXITP:
4060    MonitorEnumProc = lngRetVal
4070    Exit Function

ERRH:
4080    lngRetVal = 0&
4090    Select Case ERR.Number
        Case Else
4100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4110    End Select
4120    Resume EXITP

End Function

Private Function MonitorEnumProc2(ByVal hMonitor As Long, ByVal hdcMonitor As Long, ByVal lprcMonitor As Long, dwData As Long) As Long
' ** Parameters:
' **   hMonitor [in]
' **     A handle to the display monitor. This value will always be non-NULL.
' **   hdcMonitor [in]
' **     A handle to a device context.
' **     The device context has color attributes that are appropriate for the display
' **     monitor identified by hMonitor. The clipping area of the device context is set
' **     to the intersection of the visible region of the device context identified by
' **     the hdc parameter of EnumDisplayMonitors, the rectangle pointed to by the
' **     lprcClip parameter of EnumDisplayMonitors, and the display monitor rectangle.
' **     This value is NULL if the hdc parameter of EnumDisplayMonitors was NULL.
' **   lprcMonitor [in]
' **     A pointer to a RECT structure.
' **     If hdcMonitor is non-NULL, this rectangle is the intersection of the clipping area
' **     of the device context identified by hdcMonitor and the display monitor rectangle.
' **     The rectangle coordinates are device-context coordinates.
' **     If hdcMonitor is NULL, this rectangle is the display monitor rectangle.
' **     The rectangle coordinates are virtual-screen coordinates.
' **   dwData [in]
' **     Application-defined data that EnumDisplayMonitors passes directly to the enumeration function.

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "MonitorEnumProc2"

        Dim lngRetVal As Long

4210    lngRetVal = 1&  ' ** True: Continue enumeration; False: Stop enumeration.

4220    dwData = dwData + 1&  ' ** Increase monitor count.

EXITP:
4230    MonitorEnumProc2 = lngRetVal
4240    Exit Function

ERRH:
4250    lngRetVal = 0&
4260    Select Case ERR.Number
        Case Else
4270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4280    End Select
4290    Resume EXITP

End Function

Public Function TestMonitors() As Boolean

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "TestMonitors"

        Dim blnRetVal As Boolean

4310    blnRetVal = True

4320    MsgBox GetMonitorCount, vbInformation + vbOKOnly, "Monitor Count"

EXITP:
4330    TestMonitors = blnRetVal
4340    Exit Function

ERRH:
4350    blnRetVal = False
4360    Select Case ERR.Number
        Case Else
4370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4380    End Select
4390    Resume EXITP

End Function

Public Function FrmDimSave(strFrmName As String, lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long) As Boolean
' ** This saves the Access borders for every form affected by these monitor funcs.
' ** It should remain off (skipped) unless on the developer's single monitor system.

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "FrmDimSave"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngRecs As Long, lngThisDbsID As Long, lngFrmID As Long
        Dim blnAdd As Boolean, blnFound As Boolean, blnSkip As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

4410    blnRetVal = True

4420    blnSkip = True
4430    If blnSkip = False Then

4440      lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

4450      Set dbs = CurrentDb
4460      With dbs

4470        Set rst = .OpenRecordset("tblVBComponent_Monitor_Form", dbOpenDynaset, dbConsistent)
4480        With rst
4490          blnAdd = False: blnFound = False
4500          If .BOF = True And .EOF = True Then
4510            blnAdd = True
4520            varTmp00 = DLookup("[frm_id]", "tblForm", "[dbs_id] = " & CStr(lngThisDbsID) & " And " & _
                  "[frm_name] = '" & strFrmName & "'")
4530            If IsNull(varTmp00) = False Then
4540              lngFrmID = varTmp00
4550            Else
4560              Stop
4570            End If
4580          Else
4590            .MoveLast
4600            lngRecs = .RecordCount
4610            .MoveFirst
4620            For lngX = 1& To lngRecs
4630              If ![frm_name] = strFrmName Then
4640                blnFound = True
4650                Exit For
4660              End If
4670              If lngX < lngRecs Then .MoveNext
4680            Next
4690            If blnFound = False Then
4700              blnAdd = True
4710              varTmp00 = DLookup("[frm_id]", "tblForm", "[dbs_id] = " & CStr(lngThisDbsID) & " And " & _
                    "[frm_name] = '" & strFrmName & "'")
4720              If IsNull(varTmp00) = False Then
4730                lngFrmID = varTmp00
4740              Else
4750                Stop
4760              End If
4770            End If
4780          End If
4790          Select Case blnAdd
              Case True
4800            .AddNew
4810            ![dbs_id] = lngThisDbsID
4820            ![frm_id] = lngFrmID
                ' ** ![vbmonfrm_id] : AutoNumber.
4830            ![frm_name] = strFrmName
4840            ![vbmonfrm_top] = lngFrm_Top
4850            ![vbmonfrm_left] = lngFrm_Left
4860            ![vbmonfrm_width] = lngFrm_Width
4870            ![vbmonfrm_height] = lngFrm_Height
4880            ![vbmonfrm_datemodified] = Now()
4890            .Update
4900          Case False
4910            If ![vbmonfrm_top] <> lngFrm_Top Then
4920              .Edit
4930              ![vbmonfrm_top] = lngFrm_Top
4940              ![vbmonfrm_datemodified] = Now()
4950              .Update
4960            End If
4970            If ![vbmonfrm_left] <> lngFrm_Left Then
4980              .Edit
4990              ![vbmonfrm_left] = lngFrm_Left
5000              ![vbmonfrm_datemodified] = Now()
5010              .Update
5020            End If
5030            If ![vbmonfrm_width] <> lngFrm_Width Then
5040              .Edit
5050              ![vbmonfrm_width] = lngFrm_Width
5060              ![vbmonfrm_datemodified] = Now()
5070              .Update
5080            End If
5090            If ![vbmonfrm_height] <> lngFrm_Height Then
5100              .Edit
5110              ![vbmonfrm_height] = lngFrm_Height
5120              ![vbmonfrm_datemodified] = Now()
5130              .Update
5140            End If
5150          End Select
5160          .Close
5170        End With

5180        .Close
5190      End With

5200    End If  ' ** blnSkip.

EXITP:
5210    Set rst = Nothing
5220    Set dbs = Nothing
5230    FrmDimSave = blnRetVal
5240    Exit Function

ERRH:
5250    blnRetVal = False
5260    Select Case ERR.Number
        Case Else
5270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5280    End Select
5290    Resume EXITP

End Function

Public Function GetFormDim_Doc() As Boolean

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFormDim_Doc"

        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngHits As Long, arr_varHit() As Variant
        Dim strModName As String, strProcName As String, strLine As String, strCodeLine As String
        Dim lngLines As Long, lngDecLines As Long
        Dim lngThisDbsID As Long
        Dim blnAddAll As Boolean, blnAdd As Boolean, blnEdit As Boolean
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer
        Dim strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varHit().
        Const H_ELEMS As Integer = 10  ' ** Array's first-element UBound().
        Const H_DID  As Integer = 0
        Const H_VID  As Integer = 1
        Const H_VNAM As Integer = 2
        Const H_FID  As Integer = 3
        Const H_FNAM As Integer = 4
        Const H_PID  As Integer = 5
        Const H_PNAM As Integer = 6
        Const H_LIN  As Integer = 7
        Const H_COD  As Integer = 8
        Const H_TYP  As Integer = 9
        Const H_RAW  As Integer = 10

        Const H_FIND1 As String = "GetFormDimensions"
        Const H_FIND2 As String = "DoCmd.MoveSize"

5310    blnRetVal = True

5320    lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

5330    lngHits = 0&
5340    ReDim arr_varHit(H_ELEMS, 0)

5350    Set vbp = Application.VBE.ActiveVBProject
5360    With vbp
5370      For Each vbc In .VBComponents
5380        With vbc
5390          strModName = .Name
5400          Set cod = .CodeModule
5410          With cod
5420            lngLines = .CountOfLines
5430            lngDecLines = .CountOfDeclarationLines
5440            For lngX = lngDecLines To lngLines
5450              strLine = Trim(.Lines(lngX, 1))
5460              strProcName = vbNullString: strCodeLine = vbNullString
5470              If strLine <> vbNullString Then
5480                If Left(strLine, 1) <> "'" Then
5490                  If Left(strLine, 6) <> "Public" And Left(strLine, 7) <> "Private" Then
                        ' ** Skip procedure opening line.
5500                    intPos01 = InStr(strLine, H_FIND1)
5510                    intPos02 = InStr(strLine, H_FIND2)
5520                    If intPos01 > 0 Or intPos02 > 0 Then
5530                      strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
5540                      If strProcName <> THIS_PROC Then
5550                        intPos03 = InStr(strLine, " ")
5560                        strTmp01 = Trim(Left(strLine, intPos03))
5570                        If IsNumeric(strTmp01) = True Then
5580                          strCodeLine = strTmp01
5590                        End If
5600                        If intPos01 > 0 Then
                              ' ** GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height.
5610                          lngHits = lngHits + 1&
5620                          lngE = lngHits - 1&
5630                          ReDim Preserve arr_varHit(H_ELEMS, lngE)
5640                          arr_varHit(H_DID, lngE) = lngThisDbsID
5650                          arr_varHit(H_VID, lngE) = Null
5660                          arr_varHit(H_VNAM, lngE) = strModName
5670                          arr_varHit(H_FID, lngE) = Null
5680                          If Left(strModName, 5) = "Form_" Then
5690                            arr_varHit(H_FNAM, lngE) = Mid(strModName, 6)
5700                          Else
5710                            arr_varHit(H_FNAM, lngE) = Null
5720                          End If
5730                          arr_varHit(H_PID, lngE) = Null
5740                          arr_varHit(H_PNAM, lngE) = strProcName
5750                          arr_varHit(H_LIN, lngE) = lngX
5760                          If strCodeLine <> vbNullString Then
5770                            arr_varHit(H_COD, lngE) = CLng(strCodeLine)
5780                          Else
5790                            arr_varHit(H_COD, lngE) = Null
5800                          End If
5810                          arr_varHit(H_TYP, lngE) = H_FIND1
5820                          arr_varHit(H_RAW, lngE) = strLine
5830                        ElseIf intPos02 > 0 Then  ' ** They are mutually exclusive.
                              ' ** DoCmd.MoveSize lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height.
5840                          lngHits = lngHits + 1&
5850                          lngE = lngHits - 1&
5860                          ReDim Preserve arr_varHit(H_ELEMS, lngE)
5870                          arr_varHit(H_DID, lngE) = lngThisDbsID
5880                          arr_varHit(H_VID, lngE) = Null
5890                          arr_varHit(H_VNAM, lngE) = strModName
5900                          arr_varHit(H_FID, lngE) = Null
5910                          If Left(strModName, 5) = "Form_" Then
5920                            arr_varHit(H_FNAM, lngE) = Mid(strModName, 6)
5930                          Else
5940                            arr_varHit(H_FNAM, lngE) = Null
5950                          End If
5960                          arr_varHit(H_PID, lngE) = Null
5970                          arr_varHit(H_PNAM, lngE) = strProcName
5980                          arr_varHit(H_LIN, lngE) = lngX
5990                          If strCodeLine <> vbNullString Then
6000                            arr_varHit(H_COD, lngE) = CLng(strCodeLine)
6010                          Else
6020                            arr_varHit(H_COD, lngE) = Null
6030                          End If
6040                          arr_varHit(H_TYP, lngE) = H_FIND2
6050                          arr_varHit(H_RAW, lngE) = strLine
6060                        End If
6070                      End If  ' ** THIS_PROC.
6080                    End If  ' ** intPos01, intPos02.
6090                  End If  ' ** Public, Private.
6100                End If  ' ** Remark.
6110              End If  ' ** vbNullString.
6120            Next  ' ** lngX

6130          End With  ' ** cod.
6140        End With  ' ** vbc.
6150      Next  ' ** vbc.

6160      lngTmp02 = 0&: lngTmp03 = 0&
6170      For lngX = 0& To (lngHits - 1&)
6180        If arr_varHit(H_TYP, lngX) = H_FIND1 Then
6190          lngTmp02 = lngTmp02 + 1&
6200        ElseIf arr_varHit(H_TYP, lngX) = H_FIND2 Then
6210          lngTmp03 = lngTmp03 + 1&
6220        End If
6230      Next  ' ** lngX.

6240      Debug.Print "'GETFRMDIMS: " & CStr(lngTmp02)
6250      DoEvents
6260      Debug.Print "'MOVESIZES:  " & CStr(lngTmp03)
6270      DoEvents

6280    End With  ' ** vbp.

6290    If lngHits > 0& Then

6300      Set dbs = CurrentDb
6310      With dbs

6320        Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
6330        With rst
6340          For lngX = 0& To (lngHits - 1)
6350            .MoveFirst
6360            .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [vbcom_name] = '" & arr_varHit(H_VNAM, lngX) & "'"
6370            If .NoMatch = False Then
6380              arr_varHit(H_VID, lngX) = ![vbcom_id]
6390            Else
6400              Stop
6410            End If
6420          Next  ' ** lngX.
6430          .Close
6440        End With  ' ** rst.
6450        Set rst = Nothing

6460        Set rst = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
6470        With rst
6480          For lngX = 0& To (lngHits - 1)
6490            .MoveFirst
6500            .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [vbcom_id] = " & arr_varHit(H_VID, lngX) & " And " & _
                  "[vbcomproc_name] = '" & arr_varHit(H_PNAM, lngX) & "'"
6510            If .NoMatch = False Then
6520              arr_varHit(H_PID, lngX) = ![vbcomproc_id]
6530            Else
6540              Stop
6550            End If
6560          Next  ' ** lngX.
6570          .Close
6580        End With  ' ** rst.
6590        Set rst = Nothing

6600        Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
6610        With rst
6620          For lngX = 0& To (lngHits - 1)
6630            If IsNull(arr_varHit(H_FNAM, lngX)) = False Then
6640              .MoveFirst
6650              .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [frm_name] = '" & arr_varHit(H_FNAM, lngX) & "'"
6660              If .NoMatch = False Then
6670                arr_varHit(H_FID, lngX) = ![frm_id]
6680              Else
6690                Stop
6700              End If
6710            End If
6720          Next  ' ** lngX.
6730          .Close
6740        End With  ' ** rst.
6750        Set rst = Nothing

6760        blnAddAll = False: blnAdd = False
6770        Set rst = .OpenRecordset("tblVBComponent_Monitor", dbOpenDynaset, dbConsistent)
6780        With rst
6790          If .BOF = True And .EOF = True Then
6800            blnAddAll = True
6810          End If
6820          For lngX = 0& To (lngHits - 1&)
6830            blnAdd = False: blnEdit = False
6840            Select Case blnAddAll
                Case True
6850              blnAdd = True
6860            Case False
6870              .MoveFirst
6880              .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varHit(H_VID, lngX)) & " And " & _
                    "[vbcomproc_id] = " & CStr(arr_varHit(H_PID, lngX)) & " And [vbmon_type] = '" & arr_varHit(H_TYP, lngX) & "'"
6890              Select Case .NoMatch
                  Case True
6900                blnAdd = True
6910              Case False
6920                blnEdit = True
6930              End Select
6940            End Select
6950            If blnAdd = True Then
6960              .AddNew
6970              ![dbs_id] = arr_varHit(H_DID, lngX)
6980              ![vbcom_id] = arr_varHit(H_VID, lngX)
6990              ![vbcomproc_id] = arr_varHit(H_PID, lngX)
                  ' ** ![vbmon_id] : AutoNumber.
7000              ![frm_id] = arr_varHit(H_FID, lngX)
7010              ![vbcom_name] = arr_varHit(H_VNAM, lngX)
7020              ![frm_name] = arr_varHit(H_FNAM, lngX)
7030              ![vbcomproc_name] = arr_varHit(H_PNAM, lngX)
7040              ![vbmon_type] = arr_varHit(H_TYP, lngX)
7050              ![vbmon_line] = arr_varHit(H_LIN, lngX)
7060              ![vbmon_code] = arr_varHit(H_COD, lngX)
7070              ![vbmon_raw] = arr_varHit(H_RAW, lngX)
7080              ![vbmon_datemodified] = Now()
7090              .Update
7100            End If  ' ** blnAdd.
7110            If blnEdit = True Then
                  'ADD CODE!
7120              Stop
7130            End If  ' ** blnEdit.
7140          Next  ' ** lngX.

7150          .Close
7160        End With
7170        Set rst = Nothing
7180        .Close
7190      End With  ' ** dbs.
7200      Set dbs = Nothing

7210    End If  ' ** lngHits.

7220    Beep
7230    Debug.Print "'DONE!"

EXITP:
7240    Set rst = Nothing
7250    Set dbs = Nothing
7260    Set cod = Nothing
7270    Set vbc = Nothing
7280    Set vbp = Nothing
7290    GetFormDim_Doc = blnRetVal
7300    Exit Function

ERRH:
7310    blnRetVal = False
7320    Select Case ERR.Number
        Case Else
7330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7340    End Select
7350    Resume EXITP

End Function
