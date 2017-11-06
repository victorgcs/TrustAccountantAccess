Attribute VB_Name = "modOperSysInfoFuncs2"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modOperSysInfoFuncs2"

'VGC 11/19/2010: CHANGES!

' ** Windows-defined type SYSTEM_INFO.
Public Type SYSTEM_INFO
  wProcessorArchitecture As Integer
  wReserved As Integer
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

' ** Processor Archetecture type values.
Private Const PROCESSOR_ARCHITECTURE_INTEL         As Long = 0       ' ** x86
Private Const PROCESSOR_ARCHITECTURE_MIPS          As Long = 1
Private Const PROCESSOR_ARCHITECTURE_ALPHA         As Long = 2
Private Const PROCESSOR_ARCHITECTURE_PPC           As Long = 3
Private Const PROCESSOR_ARCHITECTURE_SHX           As Long = 4
Private Const PROCESSOR_ARCHITECTURE_ARM           As Long = 5
Private Const PROCESSOR_ARCHITECTURE_IA64          As Long = 6       ' ** Intel Itanium Processor Family (IPF)
Private Const PROCESSOR_ARCHITECTURE_ALPHA64       As Long = 7
Private Const PROCESSOR_ARCHITECTURE_MSIL          As Long = 8
Private Const PROCESSOR_ARCHITECTURE_AMD64         As Long = 9       ' ** x64 (AMD or Intel)
Private Const PROCESSOR_ARCHITECTURE_IA32_ON_WIN64 As Long = 10      ' ** WOW64
Private Const PROCESSOR_ARCHITECTURE_UNKNOWN       As Long = &HFFFF  ' ** Unknown architecture

' ** Platform ID values.
Private Const VER_PLATFORM_WIN32s        As Long = 0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32_NT      As Long = 2

' ** OS product type values.
Private Const VER_NT_WORKSTATION       As Long = &H1
Private Const VER_NT_DOMAIN_CONTROLLER As Long = &H2
Private Const VER_NT_SERVER            As Long = &H3

' ** GetProductInfo possible values.
Private Const PRODUCT_UNDEFINED                          As Long = &H0   ' ** An unknown product
Private Const PRODUCT_ULTIMATE                           As Long = &H1   ' ** Ultimate Edition
Private Const PRODUCT_HOME_BASIC                         As Long = &H2   ' ** Home Basic Edition
Private Const PRODUCT_HOME_PREMIUM                       As Long = &H3   ' ** Home Premium Edition
Private Const PRODUCT_ENTERPRISE                         As Long = &H4   ' ** Enterprise Edition
Private Const PRODUCT_HOME_BASIC_N                       As Long = &H5   ' ** Home Basic Edition
Private Const PRODUCT_BUSINESS                           As Long = &H6   ' ** Business Edition
Private Const PRODUCT_STANDARD_SERVER                    As Long = &H7   ' ** Server Standard Edition (full installation)
Private Const PRODUCT_DATACENTER_SERVER                  As Long = &H8   ' ** Server Datacenter Edition (full installation)
Private Const PRODUCT_SMALLBUSINESS_SERVER               As Long = &H9   ' ** Small Business Server
Private Const PRODUCT_ENTERPRISE_SERVER                  As Long = &HA   ' ** Server Enterprise Edition (full installation)
Private Const PRODUCT_STARTER                            As Long = &HB   ' ** Starter Edition
Private Const PRODUCT_DATACENTER_SERVER_CORE             As Long = &HC   ' ** Server Datacenter Edition (core installation)
Private Const PRODUCT_STANDARD_SERVER_CORE               As Long = &HD   ' ** Server Standard Edition (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_CORE             As Long = &HE   ' ** Server Enterprise Edition (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_IA64             As Long = &HF   ' ** Server Enterprise Edition for Itanium-based Systems
Private Const PRODUCT_BUSINESS_N                         As Long = &H10  ' ** Business Edition
Private Const PRODUCT_WEB_SERVER                         As Long = &H11  ' ** Web Server Edition (full installation)
Private Const PRODUCT_CLUSTER_SERVER                     As Long = &H12  ' ** Cluster Server Edition
Private Const PRODUCT_HOME_SERVER                        As Long = &H13  ' ** Home Server Edition
Private Const PRODUCT_STORAGE_EXPRESS_SERVER             As Long = &H14  ' ** Storage Server Express Edition
Private Const PRODUCT_STORAGE_STANDARD_SERVER            As Long = &H15  ' ** Storage Server Standard Edition
Private Const PRODUCT_STORAGE_WORKGROUP_SERVER           As Long = &H16  ' ** Storage Server Workgroup Edition
Private Const PRODUCT_STORAGE_ENTERPRISE_SERVER          As Long = &H17  ' ** Storage Server Enterprise Edition
Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS           As Long = &H18  ' ** Server for Small Business Edition
Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM       As Long = &H19  ' ** Small Business Server Premium Edition
Private Const PRODUCT_HOME_PREMIUM_N                     As Long = &H1A  ' ** Home Premium Edition
Private Const PRODUCT_ENTERPRISE_N                       As Long = &H1B  ' ** Enterprise Edition
Private Const PRODUCT_ULTIMATE_N                         As Long = &H1C  ' ** Ultimate Edition
Private Const PRODUCT_WEB_SERVER_CORE                    As Long = &H1D  ' ** Web Server Edition (core installation)
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT   As Long = &H1E  ' ** Windows Essential Business Server Management Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY     As Long = &H1F  ' ** Windows Essential Business Server Security Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING    As Long = &H20  ' ** Windows Essential Business Server Messaging Server
Private Const PRODUCT_SERVER_FOUNDATION                  As Long = &H21  ' ** Server Foundation
Private Const PRODUCT_HOME_PREMIUM_SERVER                As Long = &H22  ' **
Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS_V         As Long = &H23  ' ** Windows Server 2008 without Hyper-V for Windows Essential Server Solutions
Private Const PRODUCT_STANDARD_SERVER_V                  As Long = &H24  ' ** Server Standard Edition without Hyper-V (full installation)
Private Const PRODUCT_DATACENTER_SERVER_V                As Long = &H25  ' ** Server Datacenter Edition without Hyper-V (full installation)
Private Const PRODUCT_ENTERPRISE_SERVER_V                As Long = &H26  ' ** Server Enterprise Edition without Hyper-V (full installation)
Private Const PRODUCT_DATACENTER_SERVER_CORE_V           As Long = &H27  ' ** Server Datacenter Edition without Hyper-V (core installation)
Private Const PRODUCT_STANDARD_SERVER_CORE_V             As Long = &H28  ' ** Server Standard Edition without Hyper-V (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_CORE_V           As Long = &H29  ' ** Server Enterprise Edition without Hyper-V (core installation)
Private Const PRODUCT_HYPERV                             As Long = &H2A  ' ** Microsoft Hyper-V Server
Private Const PRODUCT_STORAGE_EXPRESS_SERVER_CORE        As Long = &H2B  ' ** (core installation)
Private Const PRODUCT_STORAGE_STANDARD_SERVER_CORE       As Long = &H2C  ' ** (core installation)
Private Const PRODUCT_STORAGE_WORKGROUP_SERVER_CORE      As Long = &H2D  ' ** (core installation)
Private Const PRODUCT_STORAGE_ENTERPRISE_SERVER_CORE     As Long = &H2E  ' ** (core installation)
Private Const PRODUCT_STARTER_N                          As Long = &H2F  ' ** Starter N
Private Const PRODUCT_PROFESSIONAL                       As Long = &H30  ' ** Professional Edition
Private Const PRODUCT_PROFESSIONAL_N                     As Long = &H31  ' ** Professional Edition
Private Const PRODUCT_SB_SOLUTION_SERVER                 As Long = &H32  ' ** Small Business
Private Const PRODUCT_SERVER_FOR_SB_SOLUTIONS            As Long = &H33  ' ** Small Business
Private Const PRODUCT_STANDARD_SERVER_SOLUTIONS          As Long = &H34  ' ** (full installation)
Private Const PRODUCT_STANDARD_SERVER_SOLUTIONS_CORE     As Long = &H35  ' ** (core installation)
Private Const PRODUCT_SB_SOLUTION_SERVER_EM              As Long = &H36  ' ** Small Business
Private Const PRODUCT_SERVER_FOR_SB_SOLUTIONS_EM         As Long = &H37  ' ** Small Business
Private Const PRODUCT_SOLUTION_EMBEDDEDSERVER            As Long = &H38  ' ** Windows MultiPoint Server (full installation)
Private Const PRODUCT_SOLUTION_EMBEDDEDSERVER_CORE       As Long = &H39  ' ** Windows MultiPoint Server (core installation)
Private Const PRODUCT_ESSENTIALBUSINESS_SERVER_MGMT      As Long = &H3B  ' **
Private Const PRODUCT_ESSENTIALBUSINESS_SERVER_ADDL      As Long = &H3C  ' **
Private Const PRODUCT_ESSENTIALBUSINESS_SERVER_MGMTSVC   As Long = &H3D  ' **
Private Const PRODUCT_ESSENTIALBUSINESS_SERVER_ADDLSVC   As Long = &H3E  ' **
Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM_CORE  As Long = &H3F  ' ** (core installation)
Private Const PRODUCT_CLUSTER_SERVER_V                   As Long = &H40  ' **
Private Const PRODUCT_EMBEDDED                           As Long = &H41  ' **
Private Const PRODUCT_STARTER_E                          As Long = &H42  ' ** Not supported
Private Const PRODUCT_HOME_BASIC_E                       As Long = &H43  ' ** Not supported
Private Const PRODUCT_HOME_PREMIUM_E                     As Long = &H44  ' ** Not supported
Private Const PRODUCT_PROFESSIONAL_E                     As Long = &H45  ' ** Not supported
Private Const PRODUCT_ENTERPRISE_E                       As Long = &H46  ' ** Not supported
Private Const PRODUCT_ULTIMATE_E                         As Long = &H47  ' ** Not supported

' ** Augmented Microsoft List:
' **   Constant                                   Value       Description
' **   =========================================  ==========  ====================================================================
' **   PRODUCT_UNDEFINED                          0x00000000  An unknown product
' **   PRODUCT_ULTIMATE                           0x00000001  Ultimate
' **   PRODUCT_HOME_BASIC                         0x00000002  Home Basic
' **   PRODUCT_HOME_PREMIUM                       0x00000003  Home Premium
' **   PRODUCT_ENTERPRISE                         0x00000004  Enterprise
' **   PRODUCT_HOME_BASIC_N                       0x00000005  Home Basic N
' **   PRODUCT_BUSINESS                           0x00000006  Business
' **   PRODUCT_STANDARD_SERVER                    0x00000007  Server Standard (full installation)
' **   PRODUCT_DATACENTER_SERVER                  0x00000008  Server Datacenter (full installation)
' **   PRODUCT_SMALLBUSINESS_SERVER               0x00000009  Windows Small Business Server
' **   PRODUCT_ENTERPRISE_SERVER                  0x0000000A  Server Enterprise (full installation)
' **   PRODUCT_STARTER                            0x0000000B  Starter
' **   PRODUCT_DATACENTER_SERVER_CORE             0x0000000C  Server Datacenter (core installation)
' **   PRODUCT_STANDARD_SERVER_CORE               0x0000000D  Server Standard (core installation)
' **   PRODUCT_ENTERPRISE_SERVER_CORE             0x0000000E  Server Enterprise (core installation)
' **   PRODUCT_ENTERPRISE_SERVER_IA64             0x0000000F  Server Enterprise for Itanium-based Systems
' **   PRODUCT_BUSINESS_N                         0x00000010  Business N
' **   PRODUCT_WEB_SERVER                         0x00000011  Web Server (full installation)
' **   PRODUCT_CLUSTER_SERVER                     0x00000012  HPC Edition
' **   PRODUCT_HOME_SERVER                        0x00000013  Home Server
' **   PRODUCT_STORAGE_EXPRESS_SERVER             0x00000014  Storage Server Express
' **   PRODUCT_STORAGE_STANDARD_SERVER            0x00000015  Storage Server Standard
' **   PRODUCT_STORAGE_WORKGROUP_SERVER           0x00000016  Storage Server Workgroup
' **   PRODUCT_STORAGE_ENTERPRISE_SERVER          0x00000017  Storage Server Enterprise
' **   PRODUCT_SERVER_FOR_SMALLBUSINESS           0x00000018  Windows Server 2008 for Windows Essential Server Solutions
' **   PRODUCT_SMALLBUSINESS_SERVER_PREMIUM       0x00000019
' **   PRODUCT_HOME_PREMIUM_N                     0x0000001A  Home Premium N
' **   PRODUCT_ENTERPRISE_N                       0x0000001B  Enterprise N
' **   PRODUCT_ULTIMATE_N                         0x0000001C  Ultimate N
' **   PRODUCT_WEB_SERVER_CORE                    0x0000001D  Web Server (core installation)
' **   PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT   0x0000001E  Windows Essential Business Server Management Server
' **   PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY     0x0000001F  Windows Essential Business Server Security Server
' **   PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING    0x00000020  Windows Essential Business Server Messaging Server
' **   PRODUCT_SERVER_FOUNDATION                  0x00000021  Server Foundation
' **   PRODUCT_HOME_PREMIUM_SERVER                0x00000022
' **   PRODUCT_SERVER_FOR_SMALLBUSINESS_V         0x00000023  Windows Server 2008 without Hyper-V for Windows Essential Server Solutions
' **   PRODUCT_STANDARD_SERVER_V                  0x00000024  Server Standard without Hyper-V (full installation)
' **   PRODUCT_DATACENTER_SERVER_V                0x00000025  Server Datacenter without Hyper-V (full installation)
' **   PRODUCT_ENTERPRISE_SERVER_V                0x00000026  Server Enterprise without Hyper-V (full installation)
' **   PRODUCT_DATACENTER_SERVER_CORE_V           0x00000027  Server Datacenter without Hyper-V (core installation)
' **   PRODUCT_STANDARD_SERVER_CORE_V             0x00000028  Server Standard without Hyper-V (core installation)
' **   PRODUCT_ENTERPRISE_SERVER_CORE_V           0x00000029  Server Enterprise without Hyper-V (core installation)
' **   PRODUCT_HYPERV                             0x0000002A  Microsoft Hyper-V Server
' **   PRODUCT_STORAGE_EXPRESS_SERVER_CORE        0x0000002B  (core installation)
' **   PRODUCT_STORAGE_STANDARD_SERVER_CORE       0x0000002C  (core installation)
' **   PRODUCT_STORAGE_WORKGROUP_SERVER_CORE      0x0000002D  (core installation)
' **   PRODUCT_STORAGE_ENTERPRISE_SERVER_CORE     0x0000002E  (core installation)
' **   PRODUCT_STARTER_N                          0x0000002F  Starter N
' **   PRODUCT_PROFESSIONAL                       0x00000030  Professional
' **   PRODUCT_PROFESSIONAL_N                     0x00000031  Professional N
' **   PRODUCT_SB_SOLUTION_SERVER                 0x00000032  Small Business
' **   PRODUCT_SERVER_FOR_SB_SOLUTIONS            0x00000033  Small Business
' **   PRODUCT_STANDARD_SERVER_SOLUTIONS          0x00000034  (full installation)
' **   PRODUCT_STANDARD_SERVER_SOLUTIONS_CORE     0x00000035  (core installation)
' **   PRODUCT_SB_SOLUTION_SERVER_EM              0x00000036  Small Business
' **   PRODUCT_SERVER_FOR_SB_SOLUTIONS_EM         0x00000037  Small Business
' **   PRODUCT_SOLUTION_EMBEDDEDSERVER            0x00000038  Windows MultiPoint Server (full installation)
' **   PRODUCT_SOLUTION_EMBEDDEDSERVER_CORE       0x00000039  Windows MultiPoint Server (core installation)
' **   PRODUCT_ESSENTIALBUSINESS_SERVER_MGMT      0x0000003B
' **   PRODUCT_ESSENTIALBUSINESS_SERVER_ADDL      0x0000003C
' **   PRODUCT_ESSENTIALBUSINESS_SERVER_MGMTSVC   0x0000003D
' **   PRODUCT_ESSENTIALBUSINESS_SERVER_ADDLSVC   0x0000003E
' **   PRODUCT_SMALLBUSINESS_SERVER_PREMIUM_CORE  0x0000003F  (core installation)
' **   PRODUCT_CLUSTER_SERVER_V                   0x00000040
' **   PRODUCT_EMBEDDED                           0x00000041
' **   PRODUCT_STARTER_E                          0x00000042  Not supported
' **   PRODUCT_HOME_BASIC_E                       0x00000043  Not supported
' **   PRODUCT_HOME_PREMIUM_E                     0x00000044  Not supported
' **   PRODUCT_PROFESSIONAL_E                     0x00000045  Not supported
' **   PRODUCT_ENTERPRISE_E                       0x00000046  Not supported
' **   PRODUCT_ULTIMATE_E                         0x00000047  Not supported

' ** New for 6.1.0.0         Value returned with 6.0.0.0
' ** ======================  ===========================
' ** PRODUCT_PROFESSIONAL    PRODUCT_BUSINESS
' ** PRODUCT_PROFESSIONAL_N  PRODUCT_BUSINESS_N
' ** PRODUCT_STARTER_N       PRODUCT_STARTER

Private Const VER_SUITE_SMALLBUSINESS            As Long = &H1
Private Const VER_SUITE_ENTERPRISE               As Long = &H2
Private Const VER_SUITE_BACKOFFICE               As Long = &H4
Private Const VER_SUITE_COMMUNICATIONS           As Long = &H8
Private Const VER_SUITE_TERMINAL                 As Long = &H10
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Long = &H20
Private Const VER_SUITE_EMBEDDEDNT               As Long = &H40
Private Const VER_SUITE_DATACENTER               As Long = &H80
Private Const VER_SUITE_SINGLEUSERTS             As Long = &H100
Private Const VER_SUITE_PERSONAL                 As Long = &H200
Private Const VER_SUITE_BLADE                    As Long = &H400
Private Const VER_SUITE_STORAGE_SERVER           As Long = &H2000
Private Const VER_SUITE_COMPUTE_SERVER           As Long = &H4000
Private Const VER_SUITE_WH_SERVER                As Long = &H8000

' ** Microsoft List:
' **   Constant                            Value       Description
' **   ==================================  ==========  ====================================================================
' **   VER_SUITE_SMALLBUSINESS             0x00000001  Microsoft Small Business Server was once installed on the system, but may have been upgraded to another version of Windows. Refer to the Remarks section for more information about this bit flag.
' **   VER_SUITE_ENTERPRISE                0x00000002  Windows Server 2008 Enterprise, Windows Server 2003, Enterprise Edition, or Windows 2000 Advanced Server is installed. Refer to the Remarks section for more information about this bit flag.
' **   VER_SUITE_BACKOFFICE                0x00000004  Microsoft BackOffice components are installed.
' **   VER_SUITE_TERMINAL                  0x00000010  Terminal Services is installed. This value is always set. If VER_SUITE_TERMINAL is set but VER_SUITE_SINGLEUSERTS is not set, the system is running in application server mode.
' **   VER_SUITE_SMALLBUSINESS_RESTRICTED  0x00000020  Microsoft Small Business Server is installed with the restrictive client license in force. Refer to the Remarks section for more information about this bit flag.
' **   VER_SUITE_EMBEDDEDNT                0x00000040  Windows XP Embedded is installed.
' **   VER_SUITE_DATACENTER                0x00000080  Windows Server 2008 Datacenter, Windows Server 2003, Datacenter Edition, or Windows 2000 Datacenter Server is installed.
' **   VER_SUITE_SINGLEUSERTS              0x00000100  Remote Desktop is supported, but only one interactive session is supported. This value is set unless the system is running in application server mode.
' **   VER_SUITE_PERSONAL                  0x00000200  Windows Vista Home Premium, Windows Vista Home Basic, or Windows XP Home Edition is installed.
' **   VER_SUITE_BLADE                     0x00000400  Windows Server 2003, Web Edition is installed.
' **   VER_SUITE_STORAGE_SERVER            0x00002000  Windows Storage Server 2003 R2 or Windows Storage Server 2003is installed.
' **   VER_SUITE_COMPUTE_SERVER            0x00004000  Windows Server 2003, Compute Cluster Edition is installed.
' **   VER_SUITE_WH_SERVER                 0x00008000  Windows Home Server is installed.

' ** From the web:
' ** A process may run as native 32 bit process, native 64 bit process, or 32 bit process on 64 bit OS (WOW).
' ** Some times you need to figure out the bitness of your process.
' ** For example, fusion needs to know this information when you install an assembly to GAC.
' ** It is illegal to install a 64 bit assembly in a 32 bit OS. But if fusion is running under WOW,
' ** then it becomes legal as long as the 64 bit assembly has the same processorArchitecture as the machine.
' ** You can use GetSystemInfo API to retrieve information about the system where the process is running.
' ** GetSystemInfo returns a SYSTEM_INFO struct.
' ** SYSTEM_INFO.wProcessorArchitecture is the system's processor architecture.
' ** If wProcessorArchitecture is PROCESSOR_ARCHITECTURE_IA64 or PROCESSOR_ARCHITECTURE_AMD64,
' ** then you know the process is running as a native 64 bit process.
' ** If wProcessorArchitecture is PROCESSOR_ARCHITECTURE_INTEL, the process is a 32 bit process.
' ** But it could be a native 32 bit process on 64 bit OS, or a WoW process.
' ** To figure out if the process is running under WoW, you can use the API IsWow64Process.
' ** If the process is running under WoW, you need to call GetNativeSystemInfo to retrieve
' ** the real system processor architecture.

Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long

Private Declare Function IsWow64Process Lib "kernel32.dll" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long

Private Declare Sub GetNativeSystemInfo Lib "kernel32.dll" (lpSystemInfo As SYSTEM_INFO)

Private Declare Sub GetSystemInfo Lib "kernel32.dll" (lpSystemInfo As SYSTEM_INFO)

Private Declare Function GetProductInfo Lib "kernel32.dll" (ByVal dwOSMajorVersion As Long, ByVal dwOSMinorVersion As Long, _
  ByVal dwSpMajorVersion As Long, ByVal dwSpMinorVersion As Long, pdwReturnedProductType As Long) As Long
' **

Public Function GetOSDisplayString() As String
' ** Display complete description of operating system environment.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "GetOSDisplayString"

        Dim osvi As OSVERSIONINFOEX  ' ** Public Type: modWindowFunctions
        Dim si As SYSTEM_INFO  ' ** Private Type: Above. {also modOperSysInfoFuncs1}
        Dim lngGNSI As Long  ' ** Pointer Get Native System Info.
        Dim lngGPI As Long   ' ** Pointer Get Product Info.
        Dim lngOsVersionInfoEx As Long
        Dim lngType As Long
        Dim strRetVal As String

110     strRetVal = vbNullString

120     osvi.dwOSVersionInfoSize = Len(osvi)

130     lngOsVersionInfoEx = GetVersionEx(osvi)  ' ** API Function: modWindowFunctions.

140     If lngOsVersionInfoEx = 1& Then

          ' ** Use GetNativeSystemInfo if supported, otherwise GetSystemInfo.
150       lngGNSI = GetProcAddress(GetModuleHandle("kernel32.dll"), "IsWow64Process")  ' ** API Functions: Above.
160       If lngGNSI > 0 Then
            ' ** IsWow64Process function exists. Now use the function to determine if this is running under Wow64.
170         IsWow64Process GetCurrentProcess(), lngGNSI  ' ** API Function: Above.
180       End If
190       If lngGNSI = 1 Then
200         GetNativeSystemInfo si  ' ** API Function: Above.
210       Else
220         GetSystemInfo si  ' ** API Function: Above.
230       End If

240       Select Case osvi.dwPlatformId
          Case VER_PLATFORM_WIN32s
            ' ** Response limited since it's likely never to be hit.
250         strRetVal = "Microsoft "
260         strRetVal = strRetVal & "Windows 32s"

270       Case VER_PLATFORM_WIN32_WINDOWS
            ' ** Response limited since it's likely never to be hit.
280         strRetVal = "Microsoft "
290         Select Case osvi.dwMinorVersion
            Case 0
300           strRetVal = strRetVal & "Windows 95"
310         Case 90
320           strRetVal = strRetVal & "Windows ME"
330         Case Else
340           strRetVal = strRetVal & "Windows 98"
350         End Select
360         strRetVal = strRetVal & osvi.dwMajorVersion & "." & osvi.dwMinorVersion
370         strRetVal = strRetVal & " (build " & osvi.dwBuildNumber & ")"
380         If InStr(osvi.szCSDVersion, Chr$(0)) > 0 Then
390           strRetVal = strRetVal & " " & Left$(osvi.szCSDVersion, (InStr(osvi.szCSDVersion, Chr$(0)) - 1))
400         End If

410       Case VER_PLATFORM_WIN32_NT

420         Select Case osvi.dwMajorVersion
            Case 3
              ' ** Response limited since it's likely never to be hit.
430           strRetVal = "Microsoft "
440           Select Case osvi.dwMinorVersion
              Case 0
450             strRetVal = strRetVal & "Windows NT 3"
460           Case 1
470             strRetVal = strRetVal & "Windows NT 3.1"
480           Case 5
490             strRetVal = strRetVal & "Windows NT 3.5"
500           Case 51
510             strRetVal = strRetVal & "Windows NT 3.51"
520           End Select

530         Case 4
              ' ** Response limited since it's likely never to be hit.
540           strRetVal = "Microsoft "
550           strRetVal = strRetVal & "Windows NT 4"

560         Case 5

570           Select Case osvi.dwMinorVersion
              Case 0
580             strRetVal = strRetVal & "Windows 2000 "
590             Select Case osvi.wProductType
                Case VER_NT_WORKSTATION
600               strRetVal = strRetVal & "Professional"
610             Case Else
620               If (osvi.wSuiteMask And VER_SUITE_DATACENTER) Then
630                 strRetVal = strRetVal & "Datacenter Server"
640               ElseIf (osvi.wSuiteMask And VER_SUITE_ENTERPRISE) Then
650                 strRetVal = strRetVal & "Advanced Server"
660               Else
670                 strRetVal = strRetVal & "Server"
680               End If
690             End Select

700           Case 1
710             strRetVal = strRetVal & "Windows XP "
720             If (osvi.wSuiteMask And VER_SUITE_PERSONAL) Then
730               strRetVal = strRetVal & "Home Edition"
740             Else
750               strRetVal = strRetVal & "Professional"
760             End If

770           Case 2

780             If GetSystemMetrics(SM_SERVERR2) Then  ' ** API Function: modWindowFunctions.
790               strRetVal = strRetVal & "Windows Server 2003 R2, "
800             ElseIf (osvi.wSuiteMask And VER_SUITE_STORAGE_SERVER) Then
810               strRetVal = strRetVal & "Windows Storage Server 2003"
820             ElseIf (osvi.wSuiteMask And VER_SUITE_WH_SERVER) Then
830               strRetVal = strRetVal & "Windows Home Server"
840             ElseIf (osvi.wProductType = VER_NT_WORKSTATION) And (si.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64) Then
850               strRetVal = strRetVal & "Windows XP Professional x64 Edition"
860             Else
870               strRetVal = strRetVal & "Windows Server 2003, "
880             End If  ' ** lngSystemMetrics.

                ' ** Test for the server type.
890             If osvi.wProductType <> VER_NT_WORKSTATION Then
900               Select Case si.wProcessorArchitecture
                  Case PROCESSOR_ARCHITECTURE_IA64
910                 If (osvi.wSuiteMask And VER_SUITE_DATACENTER) Then
920                   strRetVal = strRetVal & "Datacenter Edition for Itanium-based Systems"
930                 ElseIf (osvi.wSuiteMask And VER_SUITE_ENTERPRISE) Then
940                   strRetVal = strRetVal & "Enterprise Edition for Itanium-based Systems"
950                 End If
960               Case PROCESSOR_ARCHITECTURE_AMD64
970                 If (osvi.wSuiteMask And VER_SUITE_DATACENTER) Then
980                   strRetVal = strRetVal & "Datacenter x64 Edition"
990                 ElseIf (osvi.wSuiteMask And VER_SUITE_ENTERPRISE) Then
1000                  strRetVal = strRetVal & "Enterprise x64 Edition"
1010                Else
1020                  strRetVal = strRetVal & "Standard x64 Edition"
1030                End If
1040              Case Else
1050                If (osvi.wSuiteMask And VER_SUITE_COMPUTE_SERVER) Then
1060                  strRetVal = strRetVal & "Compute Cluster Edition"
1070                ElseIf (osvi.wSuiteMask And VER_SUITE_DATACENTER) Then
1080                  strRetVal = strRetVal & "Datacenter Edition"
1090                ElseIf (osvi.wSuiteMask And VER_SUITE_ENTERPRISE) Then
1100                  strRetVal = strRetVal & "Enterprise Edition"
1110                ElseIf (osvi.wSuiteMask And VER_SUITE_BLADE) Then
1120                  strRetVal = strRetVal & "Web Edition"
1130                Else
1140                  strRetVal = strRetVal & "Standard Edition"
1150                End If
1160              End Select  ' ** wProcessorArchitecture.
1170            End If  ' ** wProductType.

1180          End Select  ' ** dwMinorVersion.

1190        Case 6
1200          strRetVal = "Microsoft "

              ' ** Test for the specific product.
1210          Select Case osvi.dwMinorVersion
              Case 0
1220            Select Case osvi.wProductType
                Case VER_NT_WORKSTATION
1230              strRetVal = strRetVal & "Windows Vista "
1240            Case Else
1250              strRetVal = strRetVal & "Windows Server 2008 "
1260            End Select
1270          Case 1
1280            Select Case osvi.wProductType
                Case VER_NT_WORKSTATION
1290              strRetVal = strRetVal & "Windows 7 "
1300            Case Else
1310              strRetVal = strRetVal & "Windows Server 2008 R2 "
1320            End Select
1330          End Select  ' ** dwMinorVersion.

1340          lngGPI = GetProductInfo(osvi.dwMajorVersion, osvi.dwMinorVersion, 0&, 0&, lngType)

1350          If lngGPI <> 0 Then

1360            Select Case lngType
                Case PRODUCT_ULTIMATE
1370              strRetVal = strRetVal & "Ultimate Edition"
1380            Case PRODUCT_PROFESSIONAL
1390              strRetVal = strRetVal & "Professional"
1400            Case PRODUCT_HOME_PREMIUM
1410              strRetVal = strRetVal & "Home Premium Edition"
1420            Case PRODUCT_HOME_BASIC
1430              strRetVal = strRetVal & "Home Basic Edition"
1440            Case PRODUCT_ENTERPRISE
1450              strRetVal = strRetVal & "Enterprise Edition"
1460            Case PRODUCT_BUSINESS
1470              strRetVal = strRetVal & "Business Edition"
1480            Case PRODUCT_STARTER
1490              strRetVal = strRetVal & "Starter Edition"
1500            Case PRODUCT_CLUSTER_SERVER
1510              strRetVal = strRetVal & "Cluster Server Edition"
1520            Case PRODUCT_DATACENTER_SERVER
1530              strRetVal = strRetVal & "Datacenter Edition"
1540            Case PRODUCT_DATACENTER_SERVER_CORE
1550              strRetVal = strRetVal & "Datacenter Edition (core installation)"
1560            Case PRODUCT_ENTERPRISE_SERVER
1570              strRetVal = strRetVal & "Enterprise Edition"
1580            Case PRODUCT_ENTERPRISE_SERVER_CORE
1590              strRetVal = strRetVal & "Enterprise Edition (core installation)"
1600            Case PRODUCT_ENTERPRISE_SERVER_IA64
1610              strRetVal = strRetVal & "Enterprise Edition for Itanium-based Systems"
1620            Case PRODUCT_SMALLBUSINESS_SERVER
1630              strRetVal = strRetVal & "Small Business Server"
1640            Case PRODUCT_SMALLBUSINESS_SERVER_PREMIUM
1650              strRetVal = strRetVal & "Small Business Server Premium Edition"
1660            Case PRODUCT_STANDARD_SERVER
1670              strRetVal = strRetVal & "Standard Edition"
1680            Case PRODUCT_STANDARD_SERVER_CORE
1690              strRetVal = strRetVal & "Standard Edition (core installation)"
1700            Case PRODUCT_WEB_SERVER
1710              strRetVal = strRetVal & "Web Server Edition"
1720            End Select
1730          End If  ' ** lngGPI.

1740        End Select  ' ** dwMajorVersion.

            ' ** Include service pack (if any) and build number.
1750        If InStr(osvi.szCSDVersion, Chr$(0)) > 0 Then
1760          strRetVal = strRetVal & " " & Left$(osvi.szCSDVersion, (InStr(osvi.szCSDVersion, Chr$(0)) - 1))
1770        End If

1780        strRetVal = strRetVal & " (build " & osvi.dwBuildNumber & ")"

1790        If osvi.dwMajorVersion >= 6 Then
1800          Select Case si.wProcessorArchitecture
              Case PROCESSOR_ARCHITECTURE_AMD64
1810            strRetVal = strRetVal & ", 64-bit"
1820          Case PROCESSOR_ARCHITECTURE_INTEL
1830            strRetVal = strRetVal & ", 32-bit"
1840          End Select  ' ** wProcessorArchitecture.
1850        End If  ' ** dwMajorVersion.

1860      End Select  ' ** dwPlatformId.

1870    End If  ' ** lngOsVersionInfoEx.

EXITP:
1880    GetOSDisplayString = strRetVal
1890    Exit Function

ERRH:
1900    strRetVal = RET_ERR
1910    Select Case ERR.Number
        Case Else
1920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1930    End Select
1940    Resume EXITP

End Function
