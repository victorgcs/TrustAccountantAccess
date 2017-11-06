Attribute VB_Name = "modShellFuncs"
Option Compare Text
Option Explicit

Private Const THIS_NAME As String = "modShellFuncs"

'VGC 07/05/2012: CHANGES!

' ** See modProcessFuncs:
' **   EXE_IsRunning()
' **   EXE_Terminate()

' ** VBWindowstyle enumeration:
' **    0  vbHide              Window is hidden and focus is passed to the hidden window.
' **                           The vbHide constant is not applicable on Macintosh platforms.
' **    1  vbNormalFocus       Window has focus and is restored to its original size and position.
' **    2  vbMinimizedFocus    Window is displayed as an icon with focus.
' **    3  vbMaximizedFocus    Window is maximized with focus.
' **    4  vbNormalNoFocus     Window is restored to its most recent size and position.
' **                           The currently active window remains active.
' **    6  vbMinimizedNoFocus  Window is displayed as an icon. The currently active window remains active.

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal lngHWnd As Long, ByVal strOperation As String, ByVal strFile As String, ByVal strParameters As String, _
  ByVal strDirectory As String, ByVal lngShow As Long) As Long

Private Const conHide As Long = 0
Private Const conShow As Long = 1

Private blnVisible As Boolean
' **

Public Function OpenExe(varPathFile As Variant) As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "OpenExe"

        Dim blnRetVal As Boolean

110     blnRetVal = False

120     If IsNull(varPathFile) = False Then
130       If Trim(varPathFile) <> vbNullString Then
140         ShellExecute 0, "Open", varPathFile, "", "", conShow  ' ** API Function: Above.
150         blnRetVal = True
160       End If
170     End If

EXITP:
180     OpenExe = blnRetVal
190     Exit Function

ERRH:
200     blnRetVal = False
210     Select Case ERR.Number
        Case Else
220       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
230     End Select
240     Resume EXITP

End Function

Public Function OpenHelp(varPathFile As Variant) As Boolean

300   On Error GoTo ERRH

        Const THIS_PROC As String = "OpenHelp"

        Dim blnRetVal As Boolean

310     blnRetVal = False

320     If IsNull(varPathFile) = False Then
330       If Trim(varPathFile) <> vbNullString Then
340         ShellExecute 0, "Open", varPathFile, "", "", conShow  ' ** API Function: Above.
350         blnRetVal = True
360       End If
370     End If

EXITP:
380     OpenHelp = blnRetVal
390     Exit Function

ERRH:
400     blnRetVal = False
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Function

Public Function OpenDAO(varPathFile As Variant) As Boolean

500   On Error GoTo ERRH

        Const THIS_PROC As String = "OpenDAO"

        Dim blnRetVal As Boolean

510     blnRetVal = False

520     If IsNull(varPathFile) = False Then
530       If Trim(varPathFile) <> vbNullString Then
540         ShellExecute 0, "Open", varPathFile, "", "", conShow  ' ** API Function: Above.
550         blnRetVal = True
560       End If
570     End If

EXITP:
580     OpenDAO = blnRetVal
590     Exit Function

ERRH:
600     blnRetVal = False
610     Select Case ERR.Number
        Case Else
620       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
630     End Select
640     Resume EXITP

End Function

Public Function OpenExplorer(Optional varPath As Variant, Optional varFile As Variant) As Boolean

700   On Error GoTo ERRH

        Const THIS_PROC As String = "OpenExplorer"

        Dim lngExploreHwnd As Long
        Dim strPath As String, strFile As String, strPathFile As String, strParams As String
        Dim blnRetVal As Boolean

710     blnRetVal = False

720     DoEvents
730     blnRetVal = IsExploreOpen  ' ** Function: Below.
740     DoEvents

750     Select Case IsMissing(varPath)
        Case True
760       strPath = "c:\"
770       strFile = vbNullString
780       strParams = "/e,/root"
790       strPathFile = strPath
800     Case False
810       strPath = Trim(varPath)
820       Select Case IsMissing(varFile)
          Case True
830         strFile = vbNullString
840         strParams = "/e,/root,"
850         strPathFile = strPath
860       Case False
870         strFile = Trim(varFile)
880         strParams = "/e,/select,"
890         strPathFile = strPath & LNK_SEP & strFile
900       End Select
910     End Select

920     If blnRetVal = False Then
          ' ** varPath includes final backslash.
930       ShellExecute 0, "Open", "explorer.exe", strParams & strPathFile, strPath, conShow  ' ** API Function: Above.
          ' ** ShellExecute (lngHwnd, strOperation, strFile, strParameters, strDirectory, lngShow)
940       blnRetVal = True
950     Else
960       lngExploreHwnd = FindWindow("ExploreWClass", "Explorer")  ' ** API Function: Above.
970       SetForegroundWindow lngExploreHwnd  ' ** API Function: Above.
980     End If

        'Explorer.exe Command-Line Options for Windows XP:
        '   Option            Function
        '   ----------------------------------------------------------------------
        '   /n                Opens a new single-pane window for the default
        '                     selection. This is usually the root of the drive that
        '                     Windows is installed on. If the window is already
        '                     open, a duplicate opens.
        '   /e                Opens Windows Explorer in its default view.
        '   /root,<object>    Opens a window view of the specified object.
        '   /select,<object>  Opens a window view with the specified folder, file,
        '                     or program selected.

        ' "%SystemRoot%\explorer.exe /e,c:\".

        'Parameters are separated by commas. Many combinations are allowed, but only a few examples are given.
        'Explorer.exe c:\                Open directory as a single pain of icons
        'Explorer.exe /e,c:\             Explore drive as 2 lists -
        '                                  directories on left & files on right
        'Explorer.exe /e,/root,c:\       Explore drive without showing other drives
        'Explorer.exe /n,/e,/select      Opens showing only drives
        'Explorer.exe /e,/idlist,%I,%L   From Folder\..\Explore in the registry
        '                                  %I - ID number
        '                                  %L - Long filename
        'Explorer.exe  /e,DriveOrDirectory
        'Explorer.exe  /e,/root,directory,sub-directory
        'Explorer.exe  /e,/root,directory,/select,sub-directory
        '/e  List (explorer) view, Show large icons if missing (Open view)
        '/root  Sets the top level folder.
        '/select  Specifies that the directory should be selected without displaying its contents.
        '/s  ????
        '/n  Do not open the selected directory, no effect on NT
        '/idlist,%I  Expects an ID/handle. May help with cacheing. By itself, opens the desktop as icons.
        '/inproc  Stops display of window (I don't know why this is useful)

EXITP:
990     OpenExplorer = blnRetVal
1000    Exit Function

ERRH:
1010    blnRetVal = False
1020    Select Case ERR.Number
        Case Else
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1040    End Select
1050    Resume EXITP

End Function

Public Function IsExploreOpen() As Boolean
' ** Check for another instance of the Windows Explorer.
' **   OpenExplore(), above.

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "IsExploreOpen"

        Dim lngRet As Long
        Dim lngParam As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

1110    blnRetVal = False  ' ** Default not open.

1120    glngClasses = 0&
1130    ReDim garr_varClass(CLS_ELEMS, 0)

1140    blnVisible = True  ' ** Only check visible windows.

        ' ** List all open windows into garr_varClass() array.
1150    lngRet = EnumWindows(AddressOf EnumWindowsProc, lngParam)  ' ** API Function: modWindowFunctions, Function: Below.

1160    For lngX = 0& To (glngClasses - 1&)
          ' ** Positive hit must be Explorer (ExploreWClass).
1170      If garr_varClass(CLS_CLASS, lngX) = "ExploreWClass" Then
1180        blnRetVal = True
1190        Exit For
1200      End If
1210    Next

EXITP:
1220    IsExploreOpen = blnRetVal
1230    Exit Function

ERRH:
1240    Select Case ERR.Number
        Case Else
1250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1260    End Select
1270    Resume EXITP

End Function

Private Function EnumWindowsProc(ByVal lngHWnd As Long, ByVal lngParam As Long) As Boolean
' ** Called by:
' **   IsCalcOpen(), above.

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "EnumWindowsProc"

        Dim strClass As String, strTitle1 As String
        Dim strClassBuf As String * 255, strTitle1Buf As String * 255
        Dim blnFound As Boolean
        Dim intVis As Integer
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

1310    blnRetVal = True

1320    strClass = GetClassName(lngHWnd, strClassBuf, 255)  ' ** API Function: modWindowFunctions.
1330    strClass = StripNulls(strClassBuf)  ' ** Function: Below.
1340    strTitle1 = GetWindowText(lngHWnd, strTitle1Buf, 255)  ' ** API Function: modVariables.
1350    strTitle1 = StripNulls(strTitle1Buf)  ' ** Function: Below.

1360    If blnVisible = False Then
          ' ** Check both visible and hidden windows.
1370      intVis = 1
1380    Else
          ' ** Check only visible windows.
1390      intVis = IsWindowVisible(lngHWnd)  ' ** API Function: modVariables.
1400    End If

        ' ** Check if Window is a parent and visible.
1410    If GetParent(lngHWnd) = 0 And intVis = 1 Then  ' ** API Function: modCalendar.
1420      blnFound = False
1430      For lngX = 0& To (glngClasses - 1)
1440        If garr_varClass(CLS_CLASS, lngX) = strClass And _
                garr_varClass(CLS_TITLE, lngX) = strTitle1 And _
                garr_varClass(CLS_HWND, lngX) = lngHWnd Then
1450          blnFound = True
1460          Exit For
1470        End If
1480      Next
1490      If blnFound = False Then
1500        glngClasses = glngClasses + 1&
1510        lngE = glngClasses - 1&
1520        ReDim Preserve garr_varClass(CLS_ELEMS, lngE)
1530        garr_varClass(CLS_CLASS, lngE) = strClass
1540        garr_varClass(CLS_TITLE, lngE) = strTitle1
1550        garr_varClass(CLS_HWND, lngE) = lngHWnd
1560      End If
1570    End If

EXITP:
1580    EnumWindowsProc = blnRetVal
1590    Exit Function

ERRH:
1600    blnRetVal = False
1610    Select Case ERR.Number
        Case Else
1620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1630    End Select
1640    Resume EXITP

End Function

Public Function OpenCalculator() As Boolean

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "OpenCalculator"

        Dim lngCalcHwnd As Long
        Dim blnRetVal As Boolean

1710    blnRetVal = False

1720    DoEvents
1730    blnRetVal = IsCalcOpen  ' ** Function: Below.
1740    DoEvents

1750    If blnRetVal = False Then
1760      ShellExecute 0, "Open", "calc.exe", "", "", conShow  ' ** API Function: Above.
1770      blnRetVal = True
1780    Else
1790      lngCalcHwnd = FindWindow("SciCalc", "Calculator")  ' ** API Function: Above.
1800      SetForegroundWindow lngCalcHwnd  ' ** API Function: Above.
1810    End If

EXITP:
1820    OpenCalculator = blnRetVal
1830    Exit Function

ERRH:
1840    blnRetVal = False
1850    Select Case ERR.Number
        Case Else
1860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1870    End Select
1880    Resume EXITP

End Function

Public Function IsCalcOpen() As Boolean
' ** Check for another instance of the Calculator.
' **   OpenCalculator(), above.

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "IsCalcOpen"

        Dim lngRet As Long
        Dim lngParam As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

1910    blnRetVal = False  ' ** Default not open.

1920    glngClasses = 0&
1930    ReDim garr_varClass(CLS_ELEMS, 0)

1940    blnVisible = True  ' ** Only check visible windows.

        ' ** List all open windows into garr_varClass() array.
1950    lngRet = EnumWindows(AddressOf EnumWindowsProc, lngParam)  ' ** API Function: modWindowFuncs, Function: Below.

1960    For lngX = 0& To (glngClasses - 1&)
          ' ** Positive hit must be Calculator (SciCalc).
1970      If garr_varClass(CLS_CLASS, lngX) = "SciCalc" Then
1980        blnRetVal = True
1990        Exit For
2000      End If
2010    Next

        ' ** Some application class names:
        ' **   Application          Class Name
        ' **   ===================  ===========================
        ' **   Access               OMain
        ' **   Excel                XLMAIN
        ' **   FrontPage            FrontPageExplorerWindow40
        ' **   Outlook              rctrl_renwnd32
        ' **   PowerPoint 95        PP7FrameClass
        ' **   PowerPoint 97        PP97FrameClass
        ' **   PowerPoint 2000      PP9FrameClass
        ' **   PowerPoint XP        PP10FrameClass
        ' **   Project              JWinproj-WhimperMainClass
        ' **   Visual Basic Editor  wndclass_desked_gsk
        ' **   Word                 OpusApp
        ' **   Calculator           SciCalc
        ' **   Windows Explorer     ExploreWClass
        ' **   Windows Explorer     CabinetWClass
        ' **   Internet Explorer    IEFrame
        ' **   Windows Media Player WMPlayerApp
        ' **   Solitaire            Solitaire
        ' **   System Tray          Shell_TrayWnd
        ' **   Program Manager      Progman

EXITP:
2020    IsCalcOpen = blnRetVal
2030    Exit Function

ERRH:
2040    Select Case ERR.Number
        Case Else
2050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2060    End Select
2070    Resume EXITP

End Function

Public Sub ShellTest()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "ShellTest"

        ' ** To open the file...
2110    ShellExecute 0, "Open", "C:\brakes.pdf", "", "", conShow  ' ** API Function: Above.

        ' ** To print the file to the default output device...
2120    ShellExecute 0, "Print", "C:\brakes.pdf", "", "", conHide  ' ** API Function: Above.

        ' ** To execute an executable file...
2130    ShellExecute 0, "Open", "C:\RunCompact.bat", "", "", conHide  ' ** API Function: Above.

        ' ** Open file,
2140    ShellExecute 0, "Open", "C:\Temp\123.txt", "", "", 1  ' ** API Function: Above.

        ' ** Print file.
2150    ShellExecute 0, "Print", "C:\Temp\123.txt", "", "", 0  ' ** API Function: Above.

        ' ** Execute an executable,
2160    ShellExecute 0, "Open", "C:\Temp\123.bat", "", "", 0  ' ** API Function: Above.

EXITP:
2170    Exit Sub

ERRH:
2180    Select Case ERR.Number
        Case Else
2190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2200    End Select
2210    Resume EXITP

End Sub
