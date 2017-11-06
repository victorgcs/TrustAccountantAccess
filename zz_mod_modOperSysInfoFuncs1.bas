Attribute VB_Name = "modOperSysInfoFuncs1"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modOperSysInfoFuncs1"

'VGC 06/09/2011: CHANGES!

' ****************************************************************
' ** Copyright ©1996-2009 VBnet, Randy Birch, All Rights Reserved.
' ** Some pages may also contain other copyrights by the author.
' ****************************************************************
' ** Distribution: You can freely use this code in your own
' **               applications, but you may not reproduce
' **               or publish this code on any web site,
' **               online service, or distribute as source
' **               on any media without express permission.
' ****************************************************************

' ** CONDITIONAL COMPILER CONSTANTS ARE ALWAYS PRIVATE!
' ** SO WHERE USED BELOW, THEIR CRITERIA WAS NEVER MET, SINCE THE CONSTANT NEVER EXISTED!

' ** dwPlatformId.
Private Const VER_PLATFORM_WIN32s As Long = 0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32_NT As Long = 2

' ** OS product type values.
Private Const VER_NT_WORKSTATION As Long = &H1
Private Const VER_NT_DOMAIN_CONTROLLER As Long = &H2
Private Const VER_NT_SERVER As Long = &H3

' ** Product types.
Private Const VER_SERVER_NT As Long = &H80000000
Private Const VER_WORKSTATION_NT As Long = &H40000000

Private Const VER_SUITE_SMALLBUSINESS As Long = &H1
Private Const VER_SUITE_ENTERPRISE As Long = &H2
Private Const VER_SUITE_BACKOFFICE As Long = &H4
Private Const VER_SUITE_COMMUNICATIONS As Long = &H8
Private Const VER_SUITE_TERMINAL As Long = &H10
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Long = &H20
Private Const VER_SUITE_EMBEDDEDNT = &H40
Private Const VER_SUITE_DATACENTER As Long = &H80
Private Const VER_SUITE_SINGLEUSERTS As Long = &H100
Private Const VER_SUITE_PERSONAL As Long = &H200
Private Const VER_SUITE_BLADE As Long = &H400
Private Const VER_SUITE_STORAGE_SERVER As Long = &H2000
Private Const VER_SUITE_COMPUTE_SERVER As Long = &H4000
Private Const VER_SUITE_WH_SERVER As Long = &H8000

Private Const OSV_LENGTH As Long = 148
Private Const OSVEX_LENGTH As Long = 156

Private Const PROCESSOR_ARCHITECTURE_INTEL As Long = 0           ' ** x86
Private Const PROCESSOR_ARCHITECTURE_IA64 As Long = 6            ' ** Intel Itanium Processor Family (IPF)
Private Const PROCESSOR_ARCHITECTURE_AMD64 As Long = 9           ' ** x64 (AMD or Intel)
Private Const PROCESSOR_ARCHITECTURE_IA32_ON_WIN64 As Long = 10  ' ** WOW64
Private Const PROCESSOR_ARCHITECTURE_UNKNOWN As Long = &HFFFF    ' ** Unknown architecture

Private Type SYSTEM_INFO
  dwOemID As Long
  'union {
  '  DWORD  dwOemId;
  '  struct {
  '    WORD wProcessorArchitecture;
  '    WORD wReserved;
  '  };
  dwPageSize As Long
  lpMinimumApplicationAddress As Long
  lpMaximumApplicationAddress As Long
  dwActiveProcessorMask As Long
  dwNumberOfProcessors As Long
  dwProcessorType As Long
  dwAllocationGranularity As Long
  wProcessorLevel As Integer
  wProcessorRevision As Integer
End Type

Private Declare Sub GetSystemInfo Lib "kernel32.dll" (lpSystemInfo As SYSTEM_INFO)

Private Declare Function GetProductInfo Lib "kernel32.dll" (ByVal dwOSMajorVersion As Long, ByVal dwOSMinorVersion As Long, _
  ByVal dwSpMajorVersion As Long, ByVal dwSpMinorVersion As Long, pdwReturnedProductType As Long) As Long

' ** GetProductInfo possible values.
Private Const PRODUCT_UNDEFINED = &H0                          ' ** An unknown product
Private Const PRODUCT_ULTIMATE = &H1                           ' ** Ultimate Edition
Private Const PRODUCT_HOME_BASIC = &H2                         ' ** Home Basic Edition
Private Const PRODUCT_HOME_PREMIUM = &H3                       ' ** Home Premium Edition
Private Const PRODUCT_ENTERPRISE = &H4                         ' ** Enterprise Edition
Private Const PRODUCT_HOME_BASIC_N = &H5                       ' ** Home Basic Edition
Private Const PRODUCT_BUSINESS = &H6                           ' ** Business Edition
Private Const PRODUCT_STANDARD_SERVER = &H7                    ' ** Server Standard Edition (full installation)
Private Const PRODUCT_DATACENTER_SERVER = &H8                  ' ** Server Datacenter Edition (full installation)
Private Const PRODUCT_SMALLBUSINESS_SERVER = &H9               ' ** Small Business Server
Private Const PRODUCT_ENTERPRISE_SERVER = &HA                  ' ** Server Enterprise Edition (full installation)
Private Const PRODUCT_STARTER = &HB                            ' ** Starter Edition
Private Const PRODUCT_DATACENTER_SERVER_CORE = &HC             ' ** Server Datacenter Edition (core installation)
Private Const PRODUCT_STANDARD_SERVER_CORE = &HD               ' ** Server Standard Edition (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_CORE = &HE             ' ** Server Enterprise Edition (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_IA64 = &HF             ' ** Server Enterprise Edition for Itanium-based Systems
Private Const PRODUCT_BUSINESS_N = &H10                        ' ** Business Edition
Private Const PRODUCT_WEB_SERVER = &H11                        ' ** Web Server Edition (full installation)
Private Const PRODUCT_CLUSTER_SERVER = &H12                    ' ** Cluster Server Edition
Private Const PRODUCT_HOME_SERVER = &H13                       ' ** Home Server Edition
Private Const PRODUCT_STORAGE_EXPRESS_SERVER = &H14            ' ** Storage Server Express Edition
Private Const PRODUCT_STORAGE_STANDARD_SERVER = &H15           ' ** Storage Server Standard Edition
Private Const PRODUCT_STORAGE_WORKGROUP_SERVER = &H16          ' ** Storage Server Workgroup Edition
Private Const PRODUCT_STORAGE_ENTERPRISE_SERVER = &H17         ' ** Storage Server Enterprise Edition
Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS = &H18          ' ** Server for Small Business Edition
Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM = &H19      ' ** Small Business Server Premium Edition
Private Const PRODUCT_HOME_PREMIUM_N = &H1A                    ' ** Home Premium Edition
Private Const PRODUCT_ENTERPRISE_N = &H1B                      ' ** Enterprise Edition
Private Const PRODUCT_ULTIMATE_N = &H1C                        ' ** Ultimate Edition
Private Const PRODUCT_WEB_SERVER_CORE = &H1D                   ' ** Web Server Edition (core installation)
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT = &H1E  ' ** Windows Essential Business Server Management Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY = &H1F    ' ** Windows Essential Business Server Security Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING = &H20   ' ** Windows Essential Business Server Messaging Server
Private Const PRODUCT_STANDARD_SERVER_V = &H24                 ' ** Server Standard Edition without Hyper-V (full installation)
Private Const PRODUCT_DATACENTER_SERVER_V = &H25               ' ** Server Datacenter Edition without Hyper-V (full installation)
Private Const PRODUCT_ENTERPRISE_SERVER_V = &H26               ' ** Server Enterprise Edition without Hyper-V (full installation)
Private Const PRODUCT_DATACENTER_SERVER_CORE_V = &H27          ' ** Server Datacenter Edition without Hyper-V (core installation)
Private Const PRODUCT_STANDARD_SERVER_CORE_V = &H28            ' ** Server Standard Edition without Hyper-V (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_CORE_V = &H29          ' ** Server Enterprise Edition without Hyper-V (core installation)
' **

Public Function IsWin_Load() As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "IsWin_Load"

        Dim frm As Access.Form, ctl1 As Control, ctl2 As Control   'Form_frmXAdmin_SysInfo
        Dim strDocName As String
        Dim blnRetVal As Boolean

110     blnRetVal = True

120     strDocName = "frmXAdmin_SysInfo"
130     If IsLoaded("frmXAdmin_SysInfo", acForm) = True Then  ' ** Module Function: modFileUtilities.
140       Set frm = Forms(strDocName)
150       With frm
160         For Each ctl1 In .Controls
170           With ctl1
180             Select Case .ControlType
                Case acTextBox
190               Select Case .Name
                  Case "IsBackOffServ_yn"
200                 .Value = IsBackOfficeServer  ' ** Function: Below.
210               Case "IsBladeServ_yn"
220                 .Value = IsBladeServer  ' ** Function: Below.
230               Case "IsDomainCont_yn"
240                 .Value = IsDomainController  ' ** Function: Below.
250               Case "IsEntServ_yn"
260                 .Value = IsEnterpriseServer  ' ** Function: Below.
270               Case "IsSmBusRestServ_yn"
280                 .Value = IsSmallBusinessRestrictedServer  ' ** Function: Below.
290               Case "IsSmBusServ_yn"
300                 .Value = IsSmallBusinessServer  ' ** Function: Below.
310               Case "IsTermServ_yn"
320                 .Value = IsTerminalServer  ' ** Function: Below.
330               Case "IsWin2000AdvServ_yn"
340                 .Value = IsWin2000AdvancedServer  ' ** Function: Below.
350               Case "IsWin2000Plus_yn"
360                 .Value = IsWin2000Plus  ' ** Function: Below.
370               Case "IsWin2000Serv_yn"
380                 .Value = IsWin2000Server  ' ** Function: Below.
390               Case "IsWin2000Work_yn"
400                 .Value = IsWin2000Workstation  ' ** Function: Below.
410               Case "IsWin2000_yn"
420                 .Value = IsWin2000  ' ** Function: Below.
430               Case "IsWin2003ServRC2_yn"
440                 .Value = IsWin2003ServerR2  ' ** Function: Below.
450               Case "IsWin2003Serv_yn"
460                 .Value = IsWin2003Server  ' ** Function: Below.
470               Case "IsWin95OSR2_yn"
480                 .Value = IsWin95OSR2  ' ** Function: Below.
490               Case "IsWin95_yn"
500                 .Value = IsWin95  ' ** Function: Below.
510               Case "IsWin98_yn"
520                 .Value = IsWin98  ' ** Function: Below.
530               Case "IsWinLongServ_yn"
540                 .Value = IsWinLonghornServer  ' ** Function: Below.
550               Case "IsWinME_yn"
560                 .Value = IsWinME  ' ** Function: Below.
570               Case "IsWinNT4Plus_yn"
580                 .Value = IsWinNT4Plus  ' ** Function: Below.
590               Case "IsWinNT4Serv_yn"
600                 .Value = IsWinNT4Server  ' ** Function: Below.
610               Case "IsWinNT4Work_yn"
620                 .Value = IsWinNT4Workstation  ' ** Function: Below.
630               Case "IsWinNT4_yn"
640                 .Value = IsWinNT4  ' ** Function: Below.
650               Case "IsWinVistaBus_yn"
660                 .Value = IsWinVistaBusiness  ' ** Function: Below.
670               Case "IsWinVistaEnt_yn"
680                 .Value = IsWinVistaEnterprise  ' ** Function: Below.
690               Case "IsWinVistaHomeBasic_yn"
700                 .Value = IsWinVistaHomeBasic  ' ** Function: Below.
710               Case "IsWinVistaHomePrem_yn"
720                 .Value = IsWinVistaHomePremium  ' ** Function: Below.
730               Case "IsWinVistaHomeServ_yn"
740                 .Value = IsWinVistaHomeServer  ' ** Function: Below.
750               Case "IsWinVistaPlus_yn"
760                 .Value = IsWinVistaPlus  ' ** Function: Below.
770               Case "IsWinVistaSP1_yn"
780                 .Value = IsWinVistaSP1  ' ** Function: Below.
790               Case "IsWinVistaUlt_yn"
800                 .Value = IsWinVistaUltimate  ' ** Function: Below.
810               Case "IsWinVista_yn"
820                 .Value = IsWinVista  ' ** Function: Below.
830               Case "IsWinXP64_yn"
840                 .Value = IsWinXP64  ' ** Function: Below.
850               Case "IsWinXPEmbed_yn"
860                 .Value = IsWinXPEmbedded  ' ** Function: Below.
870               Case "IsWinXPHome_yn"
880                 .Value = IsWinXPHomeEdition  ' ** Function: Below.
890               Case "IsWinXPMediaCtrEd_yn"
900                 .Value = IsWinXPMediaCenter  ' ** Function: Below.
910               Case "IsWinXPPlus_yn"
920                 .Value = IsWinXPPlus  ' ** Function: Below.
930               Case "IsWinXPPro_yn"
940                 .Value = IsWinXPProEdition  ' ** Function: Below.
950               Case "IsWinXPSP2_yn"
960                 .Value = IsWinXPSP2  ' ** Function: Below.
970               Case "IsWinXPStartEd_yn"
980                 .Value = IsWinXPStarter  ' ** Function: Below.
990               Case "IsWinXPTabPCEd_yn"
1000                .Value = IsWinXPTabletPc  ' ** Function: Below.
1010              Case "IsWinXP_yn"
1020                .Value = IsWinXP  ' ** Function: Below.
1030              Case Else
                    ' ** Nothing else right now.
1040              End Select
1050              If .Value = True Then
1060                .BackColor = CLR_LTBLU
1070                .FontBold = True
1080                .ForeColor = CLR_BLU
1090                For Each ctl2 In frm.Detail.Controls
1100                  Select Case ctl2.ControlType
                      Case acLabel
1110                    If ctl2.Parent.Name = .Name Then
1120                      ctl2.FontBold = True
1130                      ctl2.ForeColor = CLR_BLU
1140                      Exit For
1150                    End If
1160                  End Select
1170                Next
1180              End If
1190            Case Else
                  ' ** Don't care right now.
1200            End Select
1210          End With
1220        Next
1230      End With
1240    End If

EXITP:
1250    Set ctl1 = Nothing
1260    Set ctl2 = Nothing
1270    Set frm = Nothing
1280    IsWin_Load = blnRetVal
1290    Exit Function

ERRH:
1300    blnRetVal = False
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

End Function

Public Function IsBackOfficeServer() As Boolean
' ** Returns True if Microsoft BackOffice components are installed.

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "IsBackOfficeServer"

        Dim osv As OSVERSIONINFOEX

        ' ** OSVERSIONINFOEX supported on NT4 or later only, so a test is required before using.
1410    If IsWinNT4Plus() Then
1420      osv.dwOSVersionInfoSize = Len(osv)
1430      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
1440        IsBackOfficeServer = (osv.wSuiteMask And VER_SUITE_BACKOFFICE)
1450      End If
1460    End If

EXITP:
1470    Exit Function

ERRH:
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Function

Public Function IsBladeServer() As Boolean
' ** Returns True if Windows Server 2003 Web Edition is installed.

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "IsBladeServer"

        Dim osv As OSVERSIONINFOEX

        ' ** OSVERSIONINFOEX supported on NT4 or later only, so a test is required before using.
1610    If IsWin2003Server() Then
1620      osv.dwOSVersionInfoSize = Len(osv)
1630      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
1640        IsBladeServer = (osv.wSuiteMask And VER_SUITE_BLADE)
1650      End If
1660    End If

EXITP:
1670    Exit Function

ERRH:
1680    Select Case ERR.Number
        Case Else
1690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1700    End Select
1710    Resume EXITP

End Function

Public Function IsDomainController() As Boolean
' ** Returns True if the server is a domain controller (Win 2000 or later), including under active directory.

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "IsDomainController"

        Dim osv As OSVERSIONINFOEX

        ' ** OSVERSIONINFOEX supported on NT4 or later only, so a test is required before using.
1810    If IsWin2000Server() Then
1820      osv.dwOSVersionInfoSize = Len(osv)
1830      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
1840        IsDomainController = (osv.wProductType = VER_NT_SERVER) And (osv.wProductType = VER_NT_DOMAIN_CONTROLLER)
1850      End If
1860    End If

EXITP:
1870    Exit Function

ERRH:
1880    Select Case ERR.Number
        Case Else
1890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1900    End Select
1910    Resume EXITP

End Function

Public Function IsEnterpriseServer() As Boolean
' ** Returns True if Windows NT 4.0 Enterprise Edition, Windows 2000 Advanced Server, or Windows Server 2003 Enterprise Edition is installed.

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "IsEnterpriseServer"

        Dim osv As OSVERSIONINFOEX

        ' ** OSVERSIONINFOEX supported on NT4 or later only, so a test is required before using.
2010    If IsWinNT4Plus() Then
2020      osv.dwOSVersionInfoSize = Len(osv)
2030      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
2040        IsEnterpriseServer = (osv.wProductType = VER_NT_SERVER) And (osv.wSuiteMask And VER_SUITE_ENTERPRISE)
2050      End If
2060    End If

EXITP:
2070    Exit Function

ERRH:
2080    Select Case ERR.Number
        Case Else
2090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2100    End Select
2110    Resume EXITP

End Function

Public Function IsSmallBusinessServer() As Boolean
' ** Returns True if Microsoft Small Business Server is installed.

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "IsSmallBusinessServer"

        Dim osv As OSVERSIONINFOEX

        ' ** OSVERSIONINFOEX supported on NT4 or later only, so a test is required before using.
2210    If IsWinNT4Plus() Then
2220      osv.dwOSVersionInfoSize = Len(osv)
2230      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
2240        IsSmallBusinessServer = (osv.wSuiteMask And VER_SUITE_SMALLBUSINESS)
2250      End If
2260    End If

EXITP:
2270    Exit Function

ERRH:
2280    Select Case ERR.Number
        Case Else
2290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2300    End Select
2310    Resume EXITP

End Function

Public Function IsSmallBusinessRestrictedServer() As Boolean
' ** Returns True if Microsoft Small Business Server is installed with the restrictive client license in force.

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "IsSmallBusinessRestrictedServer"

        Dim osv As OSVERSIONINFOEX

        ' ** OSVERSIONINFOEX supported on NT4 or later only, so a test is required before using.
2410    If IsWinNT4Plus() Then
2420      osv.dwOSVersionInfoSize = Len(osv)
2430      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
2440        IsSmallBusinessRestrictedServer = (osv.wSuiteMask And VER_SUITE_SMALLBUSINESS_RESTRICTED)
2450      End If
2460    End If

EXITP:
2470    Exit Function

ERRH:
2480    Select Case ERR.Number
        Case Else
2490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2500    End Select
2510    Resume EXITP

End Function

Public Function IsTerminalServer() As Boolean
' ** Returns True if Terminal Services is installed.

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "IsTerminalServer"

        Dim osv As OSVERSIONINFOEX

        ' ** OSVERSIONINFOEX supported on NT4 or later only, so a test is required before using.
2610    If IsWinNT4Plus() Then
2620      osv.dwOSVersionInfoSize = Len(osv)
2630      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
2640        IsTerminalServer = (osv.wSuiteMask And VER_SUITE_TERMINAL)
2650      End If
2660    End If

EXITP:
2670    Exit Function

ERRH:
2680    Select Case ERR.Number
        Case Else
2690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2700    End Select
2710    Resume EXITP

End Function

Public Function IsWin95() As Boolean
' ** Returns True if running Windows 95.

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWin95"

        Dim osv As OSVERSIONINFO

2810    osv.dwOSVersionInfoSize = Len(osv)
2820    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
2830      IsWin95 = (osv.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) And (osv.dwMajorVersion = 4 And osv.dwMinorVersion = 0) And (osv.dwBuildNumber = 950)
2840    End If

EXITP:
2850    Exit Function

ERRH:
2860    Select Case ERR.Number
        Case Else
2870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2880    End Select
2890    Resume EXITP

End Function

Public Function IsWin95OSR2() As Boolean
' ** Returns True if running Windows 95 OSR2 (OEM Service Release 2).

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWin95OSR2"

        Dim osv As OSVERSIONINFO

2910    osv.dwOSVersionInfoSize = Len(osv)
2920    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
2930      IsWin95OSR2 = (osv.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) And (osv.dwMajorVersion = 4 And osv.dwMinorVersion = 0) And (osv.dwBuildNumber = 1111)
2940    End If

EXITP:
2950    Exit Function

ERRH:
2960    Select Case ERR.Number
        Case Else
2970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2980    End Select
2990    Resume EXITP

End Function

Public Function IsWin98() As Boolean
' ** Returns True if running Windows 98.

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWin98"

        Dim osv As OSVERSIONINFO

3010    osv.dwOSVersionInfoSize = Len(osv)
3020    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
3030      IsWin98 = (osv.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) And (osv.dwMajorVersion = 4 And osv.dwMinorVersion = 10) And (osv.dwBuildNumber >= 1998)
3040    End If

EXITP:
3050    Exit Function

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Function

Public Function IsWinME() As Boolean
' ** Returns True if running Windows ME.

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinME"

        Dim osv As OSVERSIONINFO

3110    osv.dwOSVersionInfoSize = Len(osv)
3120    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
3130      IsWinME = (osv.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) And (osv.dwMajorVersion = 4 And osv.dwMinorVersion = 90) And (osv.dwBuildNumber >= 3000)
3140    End If

EXITP:
3150    Exit Function

ERRH:
3160    Select Case ERR.Number
        Case Else
3170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3180    End Select
3190    Resume EXITP

End Function

Public Function IsWinNT4() As Boolean
' ** Returns True if running Windows NT4.

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinNT4"

        Dim osv As OSVERSIONINFO

3210    osv.dwOSVersionInfoSize = Len(osv)
3220    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
3230      IsWinNT4 = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion = 4 And osv.dwMinorVersion = 0) And (osv.dwBuildNumber >= 1381)
3240    End If

EXITP:
3250    Exit Function

ERRH:
3260    Select Case ERR.Number
        Case Else
3270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3280    End Select
3290    Resume EXITP

End Function

Public Function IsWinNT4Plus() As Boolean
' ** Returns True if running Windows NT4 or later.

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinNT4Plus"

        Dim osv As OSVERSIONINFO

3310    osv.dwOSVersionInfoSize = Len(osv)
3320    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
3330      IsWinNT4Plus = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion >= 4)
3340    End If

EXITP:
3350    Exit Function

ERRH:
3360    Select Case ERR.Number
        Case Else
3370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3380    End Select
3390    Resume EXITP

End Function

Public Function IsWinNT4Server() As Boolean
' ** Returns True if running Windows NT4 Server.

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinNT4Server"

        Dim osv As OSVERSIONINFOEX

3410    If IsWinNT4() Then
3420      osv.dwOSVersionInfoSize = Len(osv)
3430      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
3440        IsWinNT4Server = (osv.wProductType And VER_NT_SERVER)
3450      End If
3460    End If

EXITP:
3470    Exit Function

ERRH:
3480    Select Case ERR.Number
        Case Else
3490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3500    End Select
3510    Resume EXITP

End Function

Public Function IsWinNT4Workstation() As Boolean
' ** Returns True if running Windows NT4 Workstation.

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinNT4Workstation"

        Dim osv As OSVERSIONINFOEX

3610    If IsWinNT4() Then
3620      osv.dwOSVersionInfoSize = Len(osv)
3630      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
3640        IsWinNT4Workstation = (osv.wProductType And VER_NT_WORKSTATION)
3650      End If
3660    End If

EXITP:
3670    Exit Function

ERRH:
3680    Select Case ERR.Number
        Case Else
3690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3700    End Select
3710    Resume EXITP

End Function

Public Function IsWin2000() As Boolean
' ** Returns True if running Windows 2000 (NT5).

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWin2000"

        Dim osv As OSVERSIONINFO

3810    osv.dwOSVersionInfoSize = Len(osv)
3820    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
3830      IsWin2000 = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion = 5 And osv.dwMinorVersion = 0) And (osv.dwBuildNumber >= 2195)
3840    End If

EXITP:
3850    Exit Function

ERRH:
3860    Select Case ERR.Number
        Case Else
3870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3880    End Select
3890    Resume EXITP

End Function

Public Function IsWin2000Plus() As Boolean
' ** Returns True if running Windows 2000 or later.

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWin2000Plus"

        Dim osv As OSVERSIONINFO

3910    osv.dwOSVersionInfoSize = Len(osv)
3920    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
3930      IsWin2000Plus = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion >= 5 And osv.dwMinorVersion >= 0)
3940    End If

EXITP:
3950    Exit Function

ERRH:
3960    Select Case ERR.Number
        Case Else
3970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3980    End Select
3990    Resume EXITP

End Function

Public Function IsWin2000AdvancedServer() As Boolean
' ** Returns True if Windows 2000 Advanced Server.

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWin2000AdvancedServer"

        Dim osv As OSVERSIONINFOEX

        ' ** OSVERSIONINFOEX supported on NT4 or later only, so a test is required before using.
4010    If IsWin2000Plus() Then
4020      osv.dwOSVersionInfoSize = Len(osv)
4030      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
4040        IsWin2000AdvancedServer = ((osv.wProductType = VER_NT_SERVER) Or (osv.wProductType = VER_NT_DOMAIN_CONTROLLER)) And _
              (osv.wSuiteMask And VER_SUITE_ENTERPRISE)
4050      End If
4060    End If

EXITP:
4070    Exit Function

ERRH:
4080    Select Case ERR.Number
        Case Else
4090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4100    End Select
4110    Resume EXITP

End Function

Public Function IsWin2000Server() As Boolean
' ** Returns True if Windows 2000 Server.

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWin2000Server"

        Dim osv As OSVERSIONINFOEX

        ' ** OSVERSIONINFOEX supported on NT4 or later only, so a test is required before using.
4210    If IsWin2000() Then
4220      osv.dwOSVersionInfoSize = Len(osv)
4230      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
4240        IsWin2000Server = (osv.wProductType = VER_NT_SERVER)
4250      End If
4260    End If

EXITP:
4270    Exit Function

ERRH:
4280    Select Case ERR.Number
        Case Else
4290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4300    End Select
4310    Resume EXITP

End Function

Public Function IsWin2000Workstation() As Boolean
' ** Returns True if running Windows NT4 Workstation.

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWin2000Workstation"

        Dim osv As OSVERSIONINFOEX

4410    If IsWin2000() Then
4420      osv.dwOSVersionInfoSize = Len(osv)
4430      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
4440        IsWin2000Workstation = (osv.wProductType And VER_NT_WORKSTATION)
4450      End If
4460    End If

EXITP:
4470    Exit Function

ERRH:
4480    Select Case ERR.Number
        Case Else
4490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4500    End Select
4510    Resume EXITP

End Function

Public Function IsWin2003Server() As Boolean
' ** Returns True if running Windows 2003 (.NET) Server.

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWin2003Server"

        Dim osv As OSVERSIONINFO

4610    osv.dwOSVersionInfoSize = Len(osv)
4620    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
4630      IsWin2003Server = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion = 5 And osv.dwMinorVersion = 2) And (osv.dwBuildNumber = 3790)
4640    End If

EXITP:
4650    Exit Function

ERRH:
4660    Select Case ERR.Number
        Case Else
4670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4680    End Select
4690    Resume EXITP

End Function

Public Function IsWin2003ServerR2() As Boolean
' ** Returns True if running Windows 2003 (.NET) Server Release 2.

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWin2003ServerR2"

4710    If IsWin2003Server() Then
4720      IsWin2003ServerR2 = GetSystemMetrics(SM_SERVERR2)  ' ** API Function: modWindowFunctions.
4730    End If

EXITP:
4740    Exit Function

ERRH:
4750    Select Case ERR.Number
        Case Else
4760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4770    End Select
4780    Resume EXITP

End Function

Public Function IsWinXP() As Boolean
' ** Returns True if running Windows XP.

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinXP"

        Dim osv As OSVERSIONINFO

4810    osv.dwOSVersionInfoSize = Len(osv)
4820    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
4830      IsWinXP = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion = 5 And osv.dwMinorVersion = 1) And (osv.dwBuildNumber >= 2600)
4840    End If

EXITP:
4850    Exit Function

ERRH:
4860    Select Case ERR.Number
        Case Else
4870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4880    End Select
4890    Resume EXITP

End Function

Public Function IsWinXPSP2() As Boolean
' ** Returns True if running Windows XP SP2 (Service Pack 2).

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinXPSP2"

        Dim osv As OSVERSIONINFOEX

4910    If IsWinXP() Then
4920      osv.dwOSVersionInfoSize = Len(osv)
4930      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
4940        IsWinXPSP2 = InStr(osv.szCSDVersion, "Service Pack 2") > 0
4950      End If
4960    End If

EXITP:
4970    Exit Function

ERRH:
4980    Select Case ERR.Number
        Case Else
4990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5000    End Select
5010    Resume EXITP

End Function

Public Function IsWinXPPlus() As Boolean
' ** Returns True if running Windows XP or later.

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinXPPlus"

        Dim osv As OSVERSIONINFO

5110    osv.dwOSVersionInfoSize = Len(osv)
5120    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
5130      IsWinXPPlus = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And ((osv.dwMajorVersion >= 5 And osv.dwMinorVersion >= 1) Or _
            (osv.dwMajorVersion >= 6 And osv.dwMinorVersion >= 0))
5140    End If

EXITP:
5150    Exit Function

ERRH:
5160    Select Case ERR.Number
        Case Else
5170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5180    End Select
5190    Resume EXITP

End Function

Public Function IsWinXPHomeEdition() As Boolean
' ** Returns True if running Windows XP Home Edition.

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinXPHomeEdition"

        Dim osv As OSVERSIONINFOEX

5210    If IsWinXP() Then
5220      osv.dwOSVersionInfoSize = Len(osv)
5230      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
5240        IsWinXPHomeEdition = ((osv.wSuiteMask And VER_SUITE_PERSONAL) = VER_SUITE_PERSONAL)
5250      End If
5260    End If

EXITP:
5270    Exit Function

ERRH:
5280    Select Case ERR.Number
        Case Else
5290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5300    End Select
5310    Resume EXITP

End Function

Public Function IsWinXPProEdition() As Boolean
' ** Returns True if running Windows XP Pro.

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinXPProEdition"

        Dim osv As OSVERSIONINFOEX

5410    If IsWinXP() Then
5420      osv.dwOSVersionInfoSize = Len(osv)
5430      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
5440        IsWinXPProEdition = Not ((osv.wSuiteMask And VER_SUITE_PERSONAL) = VER_SUITE_PERSONAL)
5450      End If
5460    End If

EXITP:
5470    Exit Function

ERRH:
5480    Select Case ERR.Number
        Case Else
5490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5500    End Select
5510    Resume EXITP

End Function

Public Function IsWinXPMediaCenter() As Boolean
' ** Returns True if running Windows XP Media Center.

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinXPMediaCenter"

5610    If IsWinXP() Then
5620      IsWinXPMediaCenter = GetSystemMetrics(SM_MEDIACENTER)  ' ** API Function: modWindowFunctions.
5630    End If

EXITP:
5640    Exit Function

ERRH:
5650    Select Case ERR.Number
        Case Else
5660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5670    End Select
5680    Resume EXITP

End Function

Public Function IsWinXPStarter() As Boolean
' ** Returns True if running Windows XP Starter.

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinXPStarter"

5710    If IsWinXP() Then
5720      IsWinXPStarter = GetSystemMetrics(SM_STARTER)  ' ** API Function: modWindowFunctions.
5730    End If

EXITP:
5740    Exit Function

ERRH:
5750    Select Case ERR.Number
        Case Else
5760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5770    End Select
5780    Resume EXITP

End Function

Public Function IsWinXPTabletPc() As Boolean
' ** Returns True if running Windows XP Tablet PC.

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinXPTabletPc"

5810    If IsWinXP() Then
5820      IsWinXPTabletPc = GetSystemMetrics(SM_TABLETPC)  ' ** API Function: modWindowFunctions.
5830    End If

EXITP:
5840    Exit Function

ERRH:
5850    Select Case ERR.Number
        Case Else
5860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5870    End Select
5880    Resume EXITP

End Function

Public Function IsWinXPEmbedded() As Boolean
' ** Returns True if OS is Windows XP Embedded.

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinXPEmbedded"

        Dim osv As OSVERSIONINFOEX

        ' ** OSVERSIONINFOEX supported on NT4 or later only, so a test is required before using.
5910    If IsWinXP() Then
5920      osv.dwOSVersionInfoSize = Len(osv)
5930      If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
5940        IsWinXPEmbedded = (osv.wSuiteMask And VER_SUITE_EMBEDDEDNT)
5950      End If
5960    End If

EXITP:
5970    Exit Function

ERRH:
5980    Select Case ERR.Number
        Case Else
5990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6000    End Select
6010    Resume EXITP

End Function

Public Function IsWinXP64() As Boolean
' ** Returns True if running Windows XP 64-bit.

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinXP64"

        Dim osv As OSVERSIONINFOEX
        Dim si As SYSTEM_INFO

6110    osv.dwOSVersionInfoSize = Len(osv)
6120    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
6130      GetSystemInfo si
          ' ** PLEASE SEE THE COMMENTS SECTION AT THE BOTTOM
          ' ** OF THIS PAGE IF YOU ARE RUNNING WINDOWS 64-BIT.
          ' **   See: http://vbnet.mvps.org/index.html?code/system/getversionex.htm
          'IsWinXP64 = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion = 5 And osv.dwMinorVersion = 2) And _
          '  (osv.wProductType <> VER_NT_WORKSTATION) And (si.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_IA64)
6140    End If

EXITP:
6150    Exit Function

ERRH:
6160    Select Case ERR.Number
        Case Else
6170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6180    End Select
6190    Resume EXITP

End Function

Public Function IsWinVista() As Boolean
' ** Returns True if running Windows Vista.

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinVista"

        Dim osv As OSVERSIONINFO

6210    osv.dwOSVersionInfoSize = Len(osv)
6220    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
6230      IsWinVista = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion = 6)
6240    End If

EXITP:
6250    Exit Function

ERRH:
6260    Select Case ERR.Number
        Case Else
6270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6280    End Select
6290    Resume EXITP

End Function

Public Function IsWinVistaPlus() As Boolean
' ** Returns True if running Windows Vista or later.

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinVistaPlus"

        Dim osv As OSVERSIONINFO

6310    osv.dwOSVersionInfoSize = Len(osv)
6320    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
6330      IsWinVistaPlus = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion >= 6)
6340    End If

EXITP:
6350    Exit Function

ERRH:
6360    Select Case ERR.Number
        Case Else
6370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6380    End Select
6390    Resume EXITP

End Function

Public Function IsWinVistaSP1() As Boolean
' ** Returns True if running any Windows Vista version with SP1 applied.

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinVistaSP1"

        Dim osv As OSVERSIONINFO

6410    osv.dwOSVersionInfoSize = Len(osv)
6420    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
6430      IsWinVistaSP1 = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion = 6) And InStr(osv.szCSDVersion, "Service Pack 1") > 0
6440    End If

EXITP:
6450    Exit Function

ERRH:
6460    Select Case ERR.Number
        Case Else
6470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6480    End Select
6490    Resume EXITP

End Function

Public Function IsWinVistaBusiness() As Boolean
' ** Returns True if running Windows Vista Business.

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinVistaBusiness"

        Dim osv As OSVERSIONINFO
        Dim dwProduct As Long

6510    If IsWinVista() Then
6520      osv.dwOSVersionInfoSize = Len(osv)
6530      GetVersionEx osv  ' ** API Function: modWindowFunctions.
6540      If GetProductInfo(osv.dwMajorVersion, osv.dwMinorVersion, 0&, 0&, dwProduct) <> 0 Then
6550        IsWinVistaBusiness = (dwProduct = PRODUCT_BUSINESS) Or (dwProduct = PRODUCT_BUSINESS_N)
6560      End If
6570    End If

EXITP:
6580    Exit Function

ERRH:
6590    Select Case ERR.Number
        Case Else
6600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6610    End Select
6620    Resume EXITP

End Function

Public Function IsWinVistaUltimate() As Boolean
' ** Returns True if running Windows Vista Ultimate.

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinVistaUltimate"

        Dim osv As OSVERSIONINFO
        Dim dwProduct As Long

6710    If IsWinVista() Then
6720      osv.dwOSVersionInfoSize = Len(osv)
6730      GetVersionEx osv  ' ** API Function: modWindowFunctions.
6740      If GetProductInfo(osv.dwMajorVersion, osv.dwMinorVersion, 0&, 0&, dwProduct) <> 0 Then
6750        IsWinVistaUltimate = (dwProduct = PRODUCT_ULTIMATE) Or (dwProduct = PRODUCT_ULTIMATE_N)
6760      End If
6770    End If

EXITP:
6780    Exit Function

ERRH:
6790    Select Case ERR.Number
        Case Else
6800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6810    End Select
6820    Resume EXITP

End Function

Public Function IsWinVistaHomeServer() As Boolean
' ** Returns True if running Windows Vista Home Server.

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinVistaHomeServer"

        Dim osv As OSVERSIONINFO
        Dim dwProduct As Long

6910    If IsWinVista() Then
6920      osv.dwOSVersionInfoSize = Len(osv)
6930      GetVersionEx osv  ' ** API Function: modWindowFunctions.
6940      If GetProductInfo(osv.dwMajorVersion, osv.dwMinorVersion, 0&, 0&, dwProduct) <> 0 Then
6950        IsWinVistaHomeServer = (dwProduct = PRODUCT_HOME_SERVER)
6960      End If
6970    End If

EXITP:
6980    Exit Function

ERRH:
6990    Select Case ERR.Number
        Case Else
7000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7010    End Select
7020    Resume EXITP

End Function

Public Function IsWinVistaHomeBasic() As Boolean
' ** Returns True if running Windows Vista Home Basic.

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinVistaHomeBasic"

        Dim osv As OSVERSIONINFO
        Dim dwProduct As Long

7110    If IsWinVista() Then
7120      osv.dwOSVersionInfoSize = Len(osv)
7130      GetVersionEx osv  ' ** API Function: modWindowFunctions.
7140      If GetProductInfo(osv.dwMajorVersion, osv.dwMinorVersion, 0&, 0&, dwProduct) <> 0 Then
7150        IsWinVistaHomeBasic = (dwProduct = PRODUCT_HOME_BASIC) Or (dwProduct = PRODUCT_HOME_BASIC_N)
7160      End If
7170    End If

EXITP:
7180    Exit Function

ERRH:
7190    Select Case ERR.Number
        Case Else
7200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7210    End Select
7220    Resume EXITP

End Function

Public Function IsWinVistaHomePremium() As Boolean
' ** Returns True if running Windows Vista Home Premium.

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinVistaHomePremium"

        Dim osv As OSVERSIONINFO
        Dim dwProduct As Long

7310    If IsWinVista() Then
7320      osv.dwOSVersionInfoSize = Len(osv)
7330      GetVersionEx osv  ' ** API Function: modWindowFunctions.
7340      If GetProductInfo(osv.dwMajorVersion, osv.dwMinorVersion, 0&, 0&, dwProduct) <> 0 Then
7350        IsWinVistaHomePremium = (dwProduct = PRODUCT_HOME_PREMIUM) Or (dwProduct = PRODUCT_HOME_PREMIUM_N)
7360      End If
7370    End If

EXITP:
7380    Exit Function

ERRH:
7390    Select Case ERR.Number
        Case Else
7400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7410    End Select
7420    Resume EXITP

End Function

Public Function IsWinVistaEnterprise() As Boolean
' ** Returns True if running Windows Vista Enterprise.

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinVistaEnterprise"

        Dim osv As OSVERSIONINFO
        Dim dwProduct As Long

7510    If IsWinVista() Then
7520      osv.dwOSVersionInfoSize = Len(osv)
7530      GetVersionEx osv  ' ** API Function: modWindowFunctions.
7540      If GetProductInfo(osv.dwMajorVersion, osv.dwMinorVersion, 0&, 0&, dwProduct) <> 0 Then
7550        IsWinVistaEnterprise = (dwProduct = PRODUCT_ENTERPRISE) Or (dwProduct = PRODUCT_ENTERPRISE_N)
7560      End If
7570    End If

EXITP:
7580    Exit Function

ERRH:
7590    Select Case ERR.Number
        Case Else
7600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7610    End Select
7620    Resume EXITP

End Function

Public Function IsWinLonghornServer() As Boolean
' ** Returns True if running Windows Longhorn Server.

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "IsWinLonghornServer"

        Dim osv As OSVERSIONINFOEX

7710    osv.dwOSVersionInfoSize = Len(osv)
7720    If GetVersionEx(osv) = 1 Then  ' ** API Function: modWindowFunctions.
7730      IsWinLonghornServer = (osv.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion = 6 And osv.dwMinorVersion = 0) And _
            (osv.wProductType <> VER_NT_WORKSTATION)
7740    End If

EXITP:
7750    Exit Function

ERRH:
7760    Select Case ERR.Number
        Case Else
7770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7780    End Select
7790    Resume EXITP

End Function
