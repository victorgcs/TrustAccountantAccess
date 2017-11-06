Attribute VB_Name = "modProcessFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modProcessFuncs"

'VGC 07/19/2016: CHANGES!

' ***************************************************************
' ***************************************************************
' ** These functions are used to check if EXCEL.EXE is still   **
' ** running after an import, and if so, to kill the process.  **
' ** 'Process' refers to the Process ID, and the list of       **
' ** 'Processes' found in the Windows Task Manager.            **
' ***************************************************************
' ***************************************************************

Private Type PROCESSENTRY32
  dwsize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * 260
End Type

Private Declare Function OpenProcess Lib "kernel32" _
  (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function EnumProcesses Lib "psapi.dll" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long

Private Declare Function EnumProcessModules Lib "psapi.dll" _
  (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long

Private Declare Function GetModuleBaseName Lib "psapi.dll" Alias "GetModuleBaseNameA" _
  (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, _
  ByVal ModuleName As String, ByVal nSize As Long) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

'Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Const PROCESS_VM_READ           As Integer = &H10
Private Const PROCESS_QUERY_INFORMATION As Integer = &H400
Private Const PROCESS_TERMINATE         As Integer = &H1

Private Const VER_PLATFORM_WIN32_WINDOWS As Integer = 1
Private Const VER_PLATFORM_WIN32_NT      As Integer = 2

Private Const TH32CS_SNAPPROCESS As Integer = &H2
' **

Public Function EXE_TestProcess() As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "EXE_TestProcess"

        Dim blnRetVal As Boolean

110     blnRetVal = True

'120     Debug.Print "'EXCEL: " & EXE_IsRunning("EXCEL.EXE")
'130     Debug.Print "'ACCESS: " & EXE_IsRunning("MSACCESS.EXE")
140     Debug.Print "'ACROBAT: " & EXE_IsRunning("Acrobat.exe")  ' ** Module Function: modProcessFuncs.

150     Beep

EXITP:
160     EXE_TestProcess = blnRetVal
170     Exit Function

ERRH:
180     blnRetVal = False
190     Select Case ERR.Number
        Case Else
200       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
210     End Select
220     Resume EXITP

End Function

Public Function EXE_IsRunning(ByVal strProcessEXE As String) As Boolean

300   On Error GoTo ERRH

        Const THIS_PROC As String = "EXE_IsRunning"

        Dim arr_lngProcess() As Long, arr_lngModule() As Long
        Dim lngRetVal As Long, lngHProcess As Long
        Dim lngX As Long
        Dim strName As String
        Dim blnRetVal As Boolean

        Const MAX_PATH As Long = 260

310     blnRetVal = False

320     strProcessEXE = UCase$(strProcessEXE)
330     ReDim arr_lngProcess(1023) As Long

340     If EnumProcesses(arr_lngProcess(0), 1024 * 4, lngRetVal) Then  ' ** API Function: Above.
350       For lngX = 0& To ((lngRetVal \ 4&) - 1&)
360         lngHProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, arr_lngProcess(lngX))  ' ** API Function: Above.
370         If lngHProcess Then
380           ReDim arr_lngModule(1023)
390           If EnumProcessModules(lngHProcess, arr_lngModule(0), 1024 * 4, lngRetVal) Then  ' ** API Function: Above.
400             strName = String$(MAX_PATH, vbNullChar)
410             GetModuleBaseName lngHProcess, arr_lngModule(0), strName, MAX_PATH  ' ** API Function: Above.
420             strName = Left(strName, InStr(strName, vbNullChar) - 1)
430             If Len(strName) = Len(strProcessEXE) Then
440               If strProcessEXE = UCase$(strName) Then
450                 blnRetVal = True
460                 Exit For
470               End If
480             End If
490           End If
500         End If
510         CloseHandle lngHProcess  ' ** API Function: modBackup.
520       Next
530     End If

EXITP:
540     EXE_IsRunning = blnRetVal
550     Exit Function

ERRH:
560     blnRetVal = False
570     Select Case ERR.Number
        Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Function

Public Function EXE_Terminate(ByVal strProcessEXE As String) As Boolean

700   On Error GoTo ERRH

        Const THIS_PROC As String = "EXE_Terminate"

        Dim lngPID As Long
        Dim lngProcess As Long
        Dim blnRetVal As Boolean

710     blnRetVal = False

720     lngPID = EXE_GetProcessID(strProcessEXE)  ' ** Function: Below.

730     If lngPID <> 0 Then

740       lngProcess = OpenProcess(PROCESS_TERMINATE, 0, lngPID)  ' ** API Function: Above.

750       TerminateProcess lngProcess, 0&  ' ** API Function: Above.
760       CloseHandle lngProcess  ' ** API Function: modBackup.

770       blnRetVal = True

780     End If

EXITP:
790     EXE_Terminate = blnRetVal
800     Exit Function

ERRH:
810     blnRetVal = False
820     Select Case ERR.Number
        Case Else
830       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
840     End Select
850     Resume EXITP

End Function

Private Function EXE_GetProcessID(ByVal strProcessEXE As String) As Long

900   On Error GoTo ERRH

        Const THIS_PROC As String = "EXE_GetProcessID"

        Dim arr_lngPID() As Long
        Dim lngProcesses As Long
        Dim lngProcess As Long
        Dim lngModule As Long
        Dim strName As String
        Dim intIndex As Integer
        Dim lngCopied As Long
        Dim lngSnapShot As Long
        Dim typPE As PROCESSENTRY32
        Dim lngVer As Long
        Dim blnDone As Boolean
        Dim lngTmp01 As Long
        Dim lngRetVal As Long

910     lngRetVal = -1&

920     lngVer = OS_CheckVersion  ' ** Function: Below.

930     If lngVer = VER_PLATFORM_WIN32_WINDOWS Then
          ' ** Windows 9x.

          ' ** Create a SnapShot of the Currently Running Processes.
940       lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)  ' ** API Function: Above.
950       If lngSnapShot >= 0 Then

960         typPE.dwsize = Len(typPE)

            ' ** Buffer the First Processes Info.
970         lngCopied = Process32First(lngSnapShot, typPE)  ' ** API Function: Above.

980         Do While lngCopied

              ' ** While there are Processes List them.
990           strName = Left(typPE.szExeFile, InStr(typPE.szExeFile, Chr(0)) - 1)
1000          strName = Mid(strName, InStrRev(strName, "\") + 1)

1010          If InStr(strName, Chr(0)) Then
1020            strName = Left(strName, InStr(strName, Chr(0)) - 1)
1030          End If

1040          lngCopied = Process32Next(lngSnapShot, typPE)  ' ** API Function: Above.
1050          If StrComp(strProcessEXE, strName, vbTextCompare) = 0 Then
1060            lngRetVal = typPE.th32ProcessID
1070            Exit Do
1080          End If

1090        Loop

1100      End If

1110    ElseIf lngVer = VER_PLATFORM_WIN32_NT Then
          ' ** Windows NT.

          ' ** The EnumProcesses Function doesn't indicate how many Process there are,
          ' ** so you need to pass a large array and trim off the empty elements
          ' ** as cbNeeded will return the number of Processes copied.

1120      ReDim arr_lngPID(255)
1130      EnumProcesses arr_lngPID(0), 1024, lngProcesses  ' ** API Function: Above.
1140      lngProcesses = lngProcesses / 4
1150      ReDim Preserve arr_lngPID(lngProcesses)

1160      For intIndex = 0 To (lngProcesses - 1&)

            ' ** Get the Process Handle, by Opening the Process.
1170        lngProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, arr_lngPID(intIndex))  ' ** API Function: Above.

1180        If lngProcess Then
              ' ** Just get the First Module, all we need is the Handle to get the Filename.

1190          If EnumProcessModules(lngProcess, lngModule, 4, 0&) Then  ' ** API Function: Above.
1200            strName = Space(260)
1210            GetModuleFileNameExA lngProcess, lngModule, strName, Len(strName)  ' ** API Function: Above.
1220            If InStr(strName, "\") > 0 Then
1230              strName = Mid(strName, InStrRev(strName, "\") + 1)
1240            End If
1250            If InStr(strName, Chr(0)) Then
1260              strName = Left(strName, InStr(strName, Chr(0)) - 1)
1270            End If
1280            If StrComp(strProcessEXE, strName, vbTextCompare) = 0 Then
1290              lngRetVal = arr_lngPID(intIndex)
1300              blnDone = True
1310            End If
1320          End If

              ' ** Close the Process Handle.
1330          lngTmp01 = CloseHandle(lngProcess)  ' ** API Function: modBackup.
1340          If blnDone Then Exit For

1350        End If

1360      Next

1370    End If

EXITP:
1380    EXE_GetProcessID = lngRetVal
1390    Exit Function

ERRH:
1400    lngRetVal = -1&
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Function

Private Function OS_CheckVersion() As Long

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "OS_CheckVersion"

        Dim typOS As OSVERSIONINFO
        Dim lngRetVal As Long

1510    lngRetVal = -1&

1520    typOS.dwOSVersionInfoSize = Len(typOS)
1530    GetVersionEx typOS  ' ** API Function: modWindowFunctions.
1540    lngRetVal = typOS.dwPlatformId

EXITP:
1550    OS_CheckVersion = lngRetVal
1560    Exit Function

ERRH:
1570    lngRetVal = -1&
1580    Select Case ERR.Number
        Case Else
1590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1600    End Select
1610    Resume EXITP

End Function

Public Function OS_GetName() As String
' ** This code was originally written by Dev Ashish.
' ** It is not to be altered or distributed,
' ** except as part of an application.
' ** You are free to use it in any application,
' ** provided the copyright notice is left unchanged.
' **
' ** Code Courtesy of
' ** Dev Ashish

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "OS_GetName"

        Dim typOS As OSVERSIONINFO
        Dim strRetVal As String

        ' ** To summarize, the base version info returned by GetVersionEx for various Windows versions is:
        ' ** ---------------------------------------------------------------------------------------------------------------------------------------------
        ' **                Win 95  Win 95 OSR2  Win 98  Win SE  Win Me  Win NT 4  Win 2000  Win XP  Win XP SP2  Win 2003 Server  Win Vista  Win Vista SP1
        ' **                ======  ===========  ======  ======  ======  ========  ========  ======  ==========  ===============  =========  =============
        ' ** PlatformID        1         1          1       1       1       2          2        2         2              2             2           2
        ' ** Major Version     4         4          4       4       4       4          5        5         5              5             6           6
        ' ** Minor Version     0         0         10      10      90       0          0        1         1              2             0           0
        ' ** Build            950*     1111       1998    2222    3000    1381       2195     2600      2600           3790          6000        6001
        ' ** ---------------------------------------------------------------------------------------------------------------------------------------------

1710    typOS.dwOSVersionInfoSize = Len(typOS)
1720    If CBool(GetVersionEx(typOS)) Then  ' ** API Function: modWindowFunctions.
1730      With typOS
            ' ** Vista.
1740        If .dwPlatformId = VER_PLATFORM_WIN32_NT And .dwMajorVersion = 6 Then
1750          strRetVal = "Windows Vista (Version " & .dwMajorVersion & "." & .dwMinorVersion & ") Build " & .dwBuildNumber
1760          If (Len(.szCSDVersion)) Then
1770            strRetVal = strRetVal & " (" & StripNullsX(.szCSDVersion) & ")"  ' ** Function: Below.
1780          End If
1790        End If
            ' ** Win 2000.
1800        If .dwPlatformId = VER_PLATFORM_WIN32_NT And .dwMajorVersion = 5 Then
1810          strRetVal = "Windows 2000 (Version " & .dwMajorVersion & "." & .dwMinorVersion & ") Build " & .dwBuildNumber
1820          If (Len(.szCSDVersion)) Then
1830            strRetVal = strRetVal & " (" & StripNullsX(.szCSDVersion) & ")"  ' ** Function: Below.
1840          End If
1850        End If
            ' ** XP.
1860        If .dwPlatformId = VER_PLATFORM_WIN32_NT And .dwMajorVersion = 5 And .dwMinorVersion = 1 Then
1870          strRetVal = "Windows XP (Version " & .dwMajorVersion & "." & .dwMinorVersion & ") Build " & .dwBuildNumber
1880          If (Len(.szCSDVersion)) Then
1890            strRetVal = strRetVal & " (" & StripNullsX(.szCSDVersion) & ")"  ' ** Function: Below.
1900          End If
1910        End If
            ' ** .Net Server.
1920        If .dwPlatformId = VER_PLATFORM_WIN32_NT And .dwMajorVersion = 5 And .dwMinorVersion = 2 Then
1930          strRetVal = "Windows .NET Server (Version " & .dwMajorVersion & "." & .dwMinorVersion & ") Build " & .dwBuildNumber
1940          If (Len(.szCSDVersion)) Then
1950            strRetVal = strRetVal & " (" & StripNullsX(.szCSDVersion) & ")"  ' ** Function: Below.
1960          End If
1970        End If
            ' ** Win ME.
1980        If (.dwMajorVersion = 4 And (.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS And .dwMinorVersion = 90)) Then
1990          strRetVal = "Windows Millenium"
2000        End If
            ' ** Win 98.
2010        If (.dwMajorVersion = 4 And (.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS And .dwMinorVersion = 10)) Then
2020          strRetVal = "Windows 98"
2030        End If
            ' ** Win 95.
2040        If (.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS And .dwMinorVersion = 0) Then
2050          strRetVal = "Windows 95"
2060        End If
            ' ** Win NT.
2070        If (.dwPlatformId = VER_PLATFORM_WIN32_NT And .dwMajorVersion <= 4) Then
2080          strRetVal = "Windows NT " & .dwMajorVersion & "." & .dwMinorVersion & " Build " & .dwBuildNumber
2090          If (Len(.szCSDVersion)) Then
2100            strRetVal = strRetVal & " (" & StripNullsX(.szCSDVersion) & ")"  ' ** Function: Below.
2110          End If
2120        End If
2130      End With
2140    End If

EXITP:
2150    OS_GetName = strRetVal
2160    Exit Function

ERRH:
2170    strRetVal = RET_ERR
2180    Select Case ERR.Number
        Case Else
2190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2200    End Select
2210    Resume EXITP

End Function

Private Function StripNullsX(strOriginal As String) As String
' ** Remove extra Nulls so String comparisons will work.
'THIS ONE'S DIFFERENT THAN THE ONE FROM TRUST ACCOUNTANT
'AND NOW I DON'T KNOW IF THIS ONE'S RIGHT.

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "StripNullsX"

        Dim intPos01 As Integer, intPos02 As Integer
        Dim intLen As Integer
        Dim intX As Integer
        Dim strRetVal As String

2310    strRetVal = vbNullString

        ' ** First, check for Nulls, Chr(0).
2320    intPos01 = InStr(1, strOriginal, vbNullChar)
2330    intPos02 = intPos01
2340    Do While intPos01 > 0
2350      If intPos01 = Len(strOriginal) Then
2360        strOriginal = Left(strOriginal, (intPos01 - 1))
2370      Else
2380        strOriginal = Left(strOriginal, (intPos01 - 1)) & Mid(strOriginal, (intPos01 + 1))
2390      End If
2400      intPos01 = InStr(strOriginal, vbNullChar)
2410    Loop

        ' ** Then check for trailing spaces.
2420    intLen = Len(strOriginal)
2430    If Mid(strOriginal, intLen) = Chr(32) Then
2440      intPos01 = 0
2450      For intX = intLen To 1 Step -1
2460        If Mid(strOriginal, intX, 1) = Chr(32) Then
2470          intPos01 = intX
2480        Else
2490          Exit For
2500        End If
2510      Next
2520      If intPos01 > 0 Then
2530        strOriginal = Left(strOriginal, (intPos01 - 1))
2540      End If
2550    End If

2560    strRetVal = strOriginal

EXITP:
2570    StripNullsX = strRetVal
2580    Exit Function

ERRH:
2590    strRetVal = vbNullString
2600    Select Case ERR.Number
        Case Else
2610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2620    End Select
2630    Resume EXITP

End Function
