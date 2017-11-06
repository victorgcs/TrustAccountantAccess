Attribute VB_Name = "modSecurityFunctions"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modSecurityFunctions"

'VGC 03/27/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const IsDemo = 0  ' ** 0 = new/upgrade; -1 = demo.
' ** Also in:
' **   modGlobConst
' **   zz_mod_MDEPrepFuncs
' **   zz_mod_DatabaseDocFuncs

'VGC 04/03/2015:
'WHEN A USER WITHOUT ADMINS PRIVILEGES TRIES TO DELETE
'AN EXTRANEOUS USER, IT ERRORS BECAUSE THEY DON'T HAVE
'THE NECESSARY PERMISSIONS!
'MAYBE 'TAAdmin' CAN DO IT, AND THE OTHER USER WILL
'THEN BE FINE, BUT WE'LL HAVE TO SEE!
'ANY OTHER IDEAS?
'WAIT! DID MY TRAPPING THE ERROR JUST FIX IT?!!!

' *******************************************************************************************************
' *******************************************************************************************************
' ** NOTE: superuser password different for Sale and Demo versions.
#If IsDemo Then
  Public Const TA_SEC As String = "TADemoAdmin1"   ' ** NEW 'Superuser' DEMO password.
  Public Const TA_SEC2 As String = "TAAdmin1"      ' ** NEW 'Superuser' password.
#Else
  Public Const TA_SEC As String = "TAAdmin1"       ' ** NEW 'Superuser' password.
  Public Const TA_SEC2 As String = "TADemoAdmin1"  ' ** NEW 'Superuser' DEMO password.
#End If
Public Const TA_SEC3 As String = "DeltaData1"      ' ** NEW 'TAAdmin' password.
Public Const TA_SEC4 As String = "TADemo1"         ' ** NEW 'TADemo' password.
Public Const TA_SEC5 As String = "taadmin"         ' ** OLD 'Superuser' password.
Public Const TA_SEC6 As String = "tademoadmin"     ' ** OLD 'Superuser' DEMO password.
Public Const TA_SEC7 As String = "deltadata"       ' ** OLD 'Admin' password.
Public Const TA_SEC8 As String = "demo"            ' ** OLD 'Demo' password.
' ** REMEMBER: Lots of code depends on the Demo password being longer
' ** than the New and Update passwords (which are always the same).
' ** NOTE: Though Admin will always remain (it's required by Access),
' ** I've set its permissions to have no access to anything. It still
' ** has a password, however, and it's the same as TAAdmin's.
' *******************************************************************************************************
' *******************************************************************************************************

' ** These 3 come with Access, and must always remain:
' **   admin
' **   Creator  : Technically, I don't think these two have passwords, though I've listed them in the security tables.
' **   Engine   : There appears to be no way to find out, either! (They can't be used for logging-in; I've tried.)
' ** This is the development user, and must always remain:
' **   superuser
' ** This is the new Admin user, and must always remain:
' **   TAAdmin
' ** This is the demo user, and must always remain in the demo version:
' **   TADemo
' ** These 2 are the sample demo users, and may be altered by users in the demo version:
' **   bill
' **   mary
' ** This is the Trust Import user, and must always remain in TrustImport.mdb/.mde.
' **   TAImport

'IF WE CHANGE TrustSec.mdw, WE'VE GOT TO ACCOMMODATE EXISTING USERS!!!

'OK, what are the implications of the various options.
'  1. Use new TrustSec.mdw.
'     a. Change the permissions in TrustDta.mdb and TrstArch.mdb on the first startup after installation.
'        This would require a copy be made before the security changes, so that recovery was possible.
'          i. Can such changes be made via VBA? Yes, I believe so.
'     b. If we leave the backend permissions as-is, the new groups I've created won't have access.
'        It looks like it would still be necessary to change their permissions.
'     c. The backends would still be owned by the Admin user, leaving them vulnerable.
'     d. If we also replace the backends, that would require transferring any existing data to the new MDBs.
'        Currently, that install option is only being considered for future upgrades where changes in the
'        backends would require it.
'  2. Use old TrustSec.mdw.
'     a. No new groups.
'     b. Data in the backends (TrustDta.mdb, TrstArch.mdb) can be viewed, exported, or manipulated
'        by even an inadvertant double-clicking of the MDBs directly.
'     c. To some extent, this possibility requires an intent on the part of the that viewer,
'        and it has not been brought up as a problem by end-users.
'     d. Though the main backend, TrustDta.mdb, does initially hide the database window when
'        opened, it can be viewed easily by either hitting the F11 key on the apparently-empty
'        window, or opening it with the shift key down. Both of these options are known to a
'        sizable portion of Access users.
'     e. All-in-all, the status quo is the best choice for now.

'Me.Username
'Me.[Primary Group]

' ** Function prototypes, constants, and type definitions for Windows 32-bit Registry API.

' ** Predefined Value Types.
'Public Const REG_NONE                        As Long = 0   ' No value type.
Public Const REG_SZ                          As Long = 1   ' Specifies a null-terminated Unicode string.
'Public Const REG_EXPAND_SZ                   As Long = 2   ' Specifies a null-terminated string that contains
                                                           ' unexpanded references to environment variables (for
                                                           ' example, %PATH%). It will be a Unicode or ANSI string
                                                           ' depending on whether you use the Unicode or ANSI functions.
Public Const REG_BINARY                      As Long = 3   ' Specifies binary data in any form.
Public Const REG_DWORD                       As Long = 4   ' Specifies a 32-bit number.
'Public Const REG_DWORD_LITTLE_ENDIAN         As Long = 4   ' Specifies a 32-bit number in little-endian format.
                                                           ' Windows is designed to run on little-endian computer
                                                           ' architectures. Therefore, this value is defined as
                                                           ' REG_DWORD in the Windows header files.
Public Const REG_DWORD_BIG_ENDIAN            As Long = 5   ' Specifies a 32-bit number in big-endian format.
                                                           ' Some UNIX systems support big-endian architectures.
'Public Const REG_LINK                        As Long = 6   ' Specifies a Unicode symbolic link.
                                                           ' Used internally; applications should not use this type.
'Public Const REG_MULTI_SZ                    As Long = 7   ' Specifies an array of null-terminated strings,
                                                           ' terminated by two null characters.
'Public Const REG_RESOURCE_LIST               As Long = 8   ' A device driver's list of hardware resources, used by
                                                           ' the driver or one of the physical devices it controls,
                                                           ' in the \ResourceMap tree.
'Public Const REG_FULL_RESOURCE_DESCRIPTOR    As Long = 9   ' List of hardware resources that a physical
                                                           ' device is using, detected and written into the
                                                           ' \HardwareDescription tree by the system.
'Public Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10   ' A device driver's list of possible hardware
                                                           ' resources it or one of the physical devices
                                                           ' it controls can use, from which the system
                                                           ' writes a subset into the \ResourceMap tree.
'Public Const REG_QWORD                       As Long = 11  ' 64-bit number
'Public Const REG_QWORD_LITTLE_ENDIAN         As Long = 11  ' Windows is designed to run on little-endian computer
                                                           ' architectures. Therefore, this value is defined as
                                                           ' REG_QWORD in the Windows header files.
'Public Const REG_QWORD_BIG_ENDIAN            As Long = 11  ' This is similar to REG_DWORD_BIG_ENDIAN except that it
                                                           ' contains the big-endian form of a 64-bit quadruple-word.

' ** Predefined reserved handle values:
Public Enum RegRoot  ' ** Variable: lRoot.
  HKEY_CLASSES_ROOT = &H80000000         ' ** -2147483648
  HKEY_CURRENT_USER = &H80000001         ' ** -2147483647
  HKEY_LOCAL_MACHINE = &H80000002        ' ** -2147483646
  HKEY_USERS = &H80000003                ' ** -2147483645
  HKEY_PERFORMANCE_DATA = &H80000004     ' ** -2147483644  Windows NT. This key is not supported in Windows Me/98/95.
  HKEY_CURRENT_CONFIG = &H80000005       ' ** -2147483643  This key does not exist for Windows NT 3.51 and earlier.
  HKEY_DYN_DATA = &H80000006             ' ** -2147483642  Windows Me/98/95 (not CE).
  HKEY_PERFORMANCE_NLSTEXT = &H80000050  ' ** -2147483568  Windows NT. This key is not supported in Windows Me/98/95.
  HKEY_PERFORMANCE_TEXT = &H80000060     ' ** -2147483552  Windows NT. This key is not supported in Windows Me/98/95.
End Enum

'Folder/predefined key     Description
'========================  ==============================================================================================
'HKEY_CURRENT_USER         Contains the root of the configuration information for the user who is currently logged on.
'                          The user's folders, screen colors, and Control Panel settings are stored here.
'                          This information is associated with the user's profile.
'                          This key is sometimes abbreviated as "HKCU."
'HKEY_USERS                Contains all the actively loaded user profiles on the computer.
'                          HKEY_CURRENT_USER is a subkey of HKEY_USERS. HKEY_USERS is sometimes abbreviated as "HKU."
'HKEY_LOCAL_MACHINE        Contains configuration information particular to the computer (for any user).
'                          This key is sometimes abbreviated as "HKLM."
'HKEY_CLASSES_ROOT         Is a subkey of HKEY_LOCAL_MACHINE\Software. The information that is stored here makes sure
'                          that the correct program opens when you open a file by using Windows Explorer.
'                          This key is sometimes abbreviated as "HKCR."
'                          Starting with Windows 2000, this information is stored under both the HKEY_LOCAL_MACHINE
'                          and HKEY_CURRENT_USER keys. The HKEY_LOCAL_MACHINE\Software\Classes key contains default
'                          settings that can apply to all users on the local computer.
'                          The HKEY_CURRENT_USER\Software\Classes key contains settings that override the default
'                          settings and apply only to the interactive user. The HKEY_CLASSES_ROOT key provides a view
'                          of the registry that merges the information from these two sources. HKEY_CLASSES_ROOT also
'                          provides this merged view for programs that are designed for earlier versions of Windows.
'                          To change the settings for the interactive user, changes must be made under
'                          HKEY_CURRENT_USER\Software\Classes instead of under HKEY_CLASSES_ROOT. To change the default
'                          settings, changes must be made under HKEY_LOCAL_MACHINE\Software\Classes. If you write keys
'                          to a key under HKEY_CLASSES_ROOT, the system stores the information under
'                          HKEY_LOCAL_MACHINE\Software\Classes. If you write values to a key under HKEY_CLASSES_ROOT,
'                          and the key already exists under HKEY_CURRENT_USER\Software\Classes, the system will store
'                          the information there instead of under HKEY_LOCAL_MACHINE\Software\Classes.
'HKEY_CURRENT_CONFIG       Contains information about the hardware profile that is used by the local computer
'                          at system startup.
'HKEY_PERFORMANCE_DATA     Registry entries subordinate to this key allow you to access performance data.
'                          The data is not actually stored in the registry; the registry functions cause the
'                          system to collect the data from its source.
'HKEY_PERFORMANCE_NLSTEXT  Registry entries subordinate to this key reference the text strings that describe
'                          counters in the local language of the area in which the computer system is running.
'                          These entries are not available to Regedit.exe and Regedt32.exe.
'HKEY_PERFORMANCE_TEXT     Registry entries subordinate to this key reference the text strings that describe
'                          counters in US English. These entries are not available to Regedit.exe and Regedt32.exe.

'Name                     Data type                       Description
'=======================  ==============================  =============================================================
'Binary Value             REG_BINARY                      Raw binary data. Most hardware component information is
'                                                         stored as binary data and is displayed in Registry Editor
'                                                         in hexadecimal format.
'DWORD Value              REG_DWORD                       Data represented by a number that is 4 bytes long (a 32-bit
'                                                         integer). Many parameters for device drivers and services
'                                                         are this type and are displayed in Registry Editor in
'                                                         binary, hexadecimal, or decimal format. Related values are
'                                                         DWORD_LITTLE_ENDIAN (least significant byte is at the lowest
'                                                         address) and REG_DWORD_BIG_ENDIAN (least significant byte is
'                                                         at the highest address).
'Expandable String Value  REG_EXPAND_SZ                   A variable-length data string. This data type includes
'                                                         variables that are resolved when a program or service uses
'                                                         the data.
'Multi-String Value       REG_MULTI_SZ                    A multiple string. Values that contain lists or multiple
'                                                         values in a form that people can read are generally this
'                                                         type. Entries are separated by spaces, commas, or other
'                                                         marks.
'String Value             REG_SZ                          A fixed-length text string.
'Binary Value             REG_RESOURCE_LIST               A series of nested arrays that is designed to store a
'                                                         resource list that is used by a hardware device driver or
'                                                         one of the physical devices it controls. This data is
'                                                         detected and written in the \ResourceMap tree by the system
'                                                         and is displayed in Registry Editor in hexadecimal format as
'                                                         a Binary Value.
'Binary Value             REG_RESOURCE_REQUIREMENTS_LIST  A series of nested arrays that is designed to store a device
'                                                         driver's list of possible hardware resources the driver or
'                                                         one of the physical devices it controls can use. The system
'                                                         writes a subset of this list in the \ResourceMap tree. This
'                                                         data is detected by the system and is displayed in Registry
'                                                         Editor in hexadecimal format as a Binary Value.
'Binary Value             REG_FULL_RESOURCE_DESCRIPTOR    A series of nested arrays that is designed to store a
'                                                         resource list that is used by a physical hardware device.
'                                                         This data is detected and written in the \HardwareDescription
'                                                         tree by the system and is displayed in Registry Editor in
'                                                         hexadecimal format as a Binary Value.
'None                     REG_NONE                        Data without any particular type. This data is written to
'                                                         the registry by the system or applications and is displayed
'                                                         in Registry Editor in hexadecimal format as a Binary Value.
'Link                     REG_LINK                        A Unicode string naming a symbolic link.
'QWORD Value              REG_QWORD                       Data represented by a number that is a 64-bit integer. This
'                                                         data is displayed in Registry Editor as a Binary Value and
'                                                         was introduced in Windows 2000.

' ** Error code constants. The full list is found in:
' **   C:\Program Files\Microsoft SDKs\Windows\v6.0\Include\WinErro.h
' ** Can use FormatMessage() function for readable description.
Public Const ERROR_SUCCESS              As Long = 0&     ' ** The operation completed successfully.
Public Const ERROR_INVALID_FUNCTION     As Long = 1&     ' ** Incorrect function.
'Public Const ERROR_FILE_NOT_FOUND       As Long = 2&     ' ** The system cannot find the file specified.
'Public Const ERROR_PATH_NOT_FOUND       As Long = 3&     ' ** The system cannot find the path specified.
'Public Const ERROR_TOO_MANY_OPEN_FILES  As Long = 4&     ' ** The system cannot open the file.
'Public Const ERROR_ACCESS_DENIED        As Long = 5&     ' ** Access is denied.
'Public Const ERROR_INVALID_HANDLE       As Long = 6&     ' ** The handle is invalid.
'Public Const ERROR_ARENA_TRASHED        As Long = 7&     ' ** The storage control blocks were destroyed.
'Public Const ERROR_NOT_ENOUGH_MEMORY    As Long = 8&     ' ** Not enough storage is available to process this command. (dderror)
Public Const ERROR_NOT_SUPPORTED        As Long = 50&
'Public Const ERROR_BAD_NETPATH          As Long = 53&    ' ** The network path was not found.
'Public Const ERROR_INVALID_PARAMETER    As Long = 87&    ' ** The parameter is incorrect. (dderror)
'Public Const ERROR_CALL_NOT_IMPLEMENTED As Long = 120&   ' ** This function is not supported on this system.
'Public Const ERROR_INSUFFICIENT_BUFFER  As Long = 122&   ' ** The data area passed to a system call is too small.
'Public Const ERROR_BAD_PATHNAME         As Long = 161&   ' ** The specified path is invalid.
Public Const ERROR_MORE_DATA            As Long = 234&   ' ** More data is available. (dderror)
'Public Const ERROR_NO_MORE_ITEMS        As Long = 259&   ' ** No more data is available.
'Public Const ERROR_BADDB                As Long = 1009&  ' ** The configuration registry database is corrupt.
'Public Const ERROR_BADKEY               As Long = 1010&  ' ** The configuration registry key is invalid.
'Public Const ERROR_CANTOPEN             As Long = 1011&  ' ** The configuration registry key could not be opened.
'Public Const ERROR_CANTREAD             As Long = 1012&  ' ** The configuration registry key could not be read.
'Public Const ERROR_CANTWRITE            As Long = 1013&  ' ** The configuration registry key could not be written.
'Public Const ERROR_REGISTRY_RECOVERED   As Long = 1014&  ' ** One of the files in the registry database had to be recovered by use of a log or alternate copy. The recovery was successful.
'Public Const ERROR_REGISTRY_CORRUPT     As Long = 1015&  ' ** The registry is corrupted. The structure of one of the files containing registry data is corrupted, or the system's memory image of the file is corrupted, or the file could not be recovered because the alternate copy or log was absent or corrupted.
'Public Const ERROR_REGISTRY_IO_FAILED   As Long = 1016&  ' ** An I/O operation initiated by the registry failed unrecoverably. The registry could not read in, or write out, or flush, one of the files that contain the system's image of the registry.
'Public Const ERROR_NOT_REGISTRY_FILE    As Long = 1017&  ' ** The system has attempted to load or restore a file into the registry, but the specified file is not in a registry file format.
'Public Const ERROR_KEY_DELETED          As Long = 1018&  ' ** Illegal operation attempted on a registry key that has been marked for deletion.
'Public Const ERROR_NO_LOG_SPACE         As Long = 1019&  ' ** System could not allocate the required space in a registry log.
'Public Const ERROR_KEY_HAS_CHILDREN     As Long = 1020&  ' ** Cannot create a symbolic link in a registry key that already has subkeys or values.
Public Const ERROR_BAD_DEVICE           As Long = 1200&
Public Const ERROR_CONNECTION_UNAVAIL   As Long = 1201&
Public Const ERROR_NO_NET_OR_BAD_PATH   As Long = 1203&
Public Const ERROR_EXTENDED_ERROR       As Long = 1208&
Public Const ERROR_NO_NETWORK           As Long = 1222&
'Public Const ERROR_NOT_ALL_ASSIGNED     As Long = 1300&  ' ** Not all privileges or groups referenced are assigned to the caller.
Public Const ERROR_NOT_CONNECTED        As Long = 2250&

Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" _
  (ByVal dwFlags As Long, ByVal lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
  ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long

' ** Windows Access Mask (WinRights) partial enumeration:
'Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const STANDARD_RIGHTS_READ     As Long = &H20000
Public Const STANDARD_RIGHTS_WRITE    As Long = &H20000
'Public Const STANDARD_RIGHTS_EXECUTE  As Long = &H20000
Public Const STANDARD_RIGHTS_ALL      As Long = &H1F0000
'Public Const SPECIFIC_RIGHTS_ALL      As Long = &HFFFF

' ** Identifiers for the IADsAccessControlEntry.AccessMask property for file and file share objects.
' ** NOTE: The 'ACE_' prefix is non-standard, and added by me to better identify them within this program.
'Public Const ACE_DELETE                 As Long = &H10000
'Public Const ACE_READ_CONTROL           As Long = &H20000
'Public Const ACE_WRITE_DAC              As Long = &H40000
'Public Const ACE_WRITE_OWNER            As Long = &H80000
Public Const ACE_SYNCHRONIZE            As Long = &H100000
'Public Const ACE_ACCESS_SYSTEM_SECURITY As Long = &H1000000  '(0x01000000)

' ** Standard access rights are those rights corresponding to operations common to most types of securable objects.
' **   Access Right              Value         Description
' **   ========================  ============  =============
' **   DELETE                    0x00010000    The right to delete the object.
' **   READ_CONTROL              0x00020000    The right to read the information in the file or directory object's security descriptor. This does not include the information in the SACL.
' **   WRITE_DAC                 0x00040000    The right to modify the DACL in the object's security descriptor.
' **   WRITE_OWNER               0x00080000    The right to change the owner in the object's security descriptor.
' **   SYNCHRONIZE               0x00100000    The right to use the object for synchronization. This enables a thread to wait until the object is in the signaled state. Some object types do not support this access right.
' **   STANDARD_RIGHTS_REQUIRED  0x000F0000    Combines DELETE, READ_CONTROL, WRITE_DAC, and WRITE_OWNER access.
' **   STANDARD_RIGHTS_READ      READ_CONTROL  Currently defined to equal READ_CONTROL.
' **   STANDARD_RIGHTS_WRITE     READ_CONTROL  Currently defined to equal READ_CONTROL.
' **   STANDARD_RIGHTS_EXECUTE   READ_CONTROL  Currently defined to equal READ_CONTROL.
' **   STANDARD_RIGHTS_ALL       0x001F0000    Combines DELETE, READ_CONTROL, WRITE_DAC, WRITE_OWNER, and SYNCHRONIZE access.

' ** Identifiers for the IADsAccessControlEntry.AccessMask property for registry objects.
Public Const KEY_QUERY_VALUE        As Long = &H1
Public Const KEY_SET_VALUE          As Long = &H2
Public Const KEY_CREATE_SUB_KEY     As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_NOTIFY             As Long = &H10
Public Const KEY_CREATE_LINK        As Long = &H20
'Public Const KEY_WOW64_32KEY        As Long = &H200
'Public Const KEY_WOW64_64KEY        As Long = &H100
'Public Const KEY_WOW64_RES          As Long = &H300
Public Const KEY_READ               As Long = _
  ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not ACE_SYNCHRONIZE))
'Public Const KEY_WRITE              As Long = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not ACE_SYNCHRONIZE))
'Public Const KEY_EXECUTE            As Long = ((KEY_READ) And (Not ACE_SYNCHRONIZE))
Public Const KEY_ALL_ACCESS         As Long = _
  ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or _
  KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not ACE_SYNCHRONIZE))

' ** Obsolete, replaced by RegOpenKeyEx().
'Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
'  (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, _
  phkResult As Long) As Long
  ' ** hKey          : Handle to a currently open key or any of the predefined reserved handle values.
  ' ** lpSubKey      : Pointer to a null-terminated string containing the name of the subkey to open.
  ' **               : If this parameter is NULL or a pointer to an empty string, the function will open a new
  ' **               : handle to the key identified by the hKey parameter. In this case, the function will not
  ' **               : close the handles previously opened.
  ' ** ulOptions     : Reserved; set to 0.
  ' ** samDesired    : Not supported; set to 0.  FALSE! WE'RE USING IT BELOW!
  ' ** phkResult     : Pointer to a variable that receives a handle to the opened key.
  ' **               : When you no longer need the returned handle, call the RegCloseKey function to close it.
  ' ** Return Value  : If the function succeeds, the return value is ERROR_SUCCESS.
  ' **
  ' ** When you call the RegOpenKeyEx function, the system checks the requested access rights against
  ' ** the key's security descriptor. If the user does not have the correct access to the registry key,
  ' ** the open operation fails. If an administrator needs access to the key, the solution is to enable
  ' ** the SE_TAKE_OWNERSHIP_NAME privilege and open the registry key with ACE_WRITE_OWNER access.
  ' ** For more information, see Enabling and Disabling Privileges.
  ' ** You can request the ACCESS_SYSTEM_SECURITY access right to a registry key if you want to read or
  ' ** write the key's SACL. For more information, see Access-Control Lists (ACLs) and SACL Access Right.

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

' ** Obsolete, replaced by RegQueryValueEx().
'Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" _
'  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long

' ** RegQueryValueEx retrieves the type, content and data for a specified value name.
' ** Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, _
  lpData As Any, lpcbData As Long) As Long

Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, _
  ByVal lpData As Long, lpcbData As Long) As Long

Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, _
  lpData As Long, lpcbData As Long) As Long

Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, _
  ByVal lpData As String, lpcbData As Long) As Long
  ' ** hKey          : A Handle to a currently open key or any of the predefined reserved handle values.
  ' ** lpValueName   : The name of the registry value.
  ' **               : If lpValueName is NULL or an empty string, "", the function retrieves the type and data
  ' **               : for the key's unnamed or default value, if any. For more information, see Registry Element
  ' **               : Size Limits. Keys do not automatically have an unnamed or default value.
  ' **               : Unnamed values can be of any type.
  ' ** lpReserved    : This parameter is reserved and must be NULL.
  ' ** lpType        : A pointer to a variable that receives a code indicating the type of data stored in the
  ' **               : specified value. For a list of the possible type codes, see Registry Value Types.
  ' **               : The lpType parameter can be NULL if the type code is not required.
  ' ** lpData        : A pointer to a buffer that receives the value's data. This parameter can be NULL if the
  ' **               : data is not required.
  ' ** lpcbData      : A pointer to a variable that specifies the size of the buffer pointed to by the lpData
  ' **               : parameter, in bytes. When the function returns, this variable contains the size of the
  ' **               : data copied to lpData. The lpcbData parameter can be NULL only if lpData is NULL.
  ' **               : If the data has the REG_SZ, REG_MULTI_SZ or REG_EXPAND_SZ type, this size includes any
  ' **               : terminating null character or characters. For more information, see Remarks.
  ' **               : If the buffer specified by lpData parameter is not large enough to hold the data, the
  ' **               : function returns ERROR_MORE_DATA and stores the required buffer size in the variable
  ' **               : pointed to by lpcbData. In this case, the contents of the lpData buffer are undefined.
  ' **               : If lpData is NULL, and lpcbData is non-NULL, the function returns ERROR_SUCCESS and
  ' **               : stores the size of the data, in bytes, in the variable pointed to by lpcbData.
  ' **               : This enables an application to determine the best way to allocate a buffer for the
  ' **               : value's data. If hKey specifies HKEY_PERFORMANCE_DATA and the lpData buffer is not
  ' **               : large enough to contain all of the returned data, RegQueryValueEx returns
  ' **               : ERROR_MORE_DATA and the value returned through the lpcbData parameter is undefined.
  ' **               : This is because the size of the performance data can change from one call to the next.
  ' **               : In this case, you must increase the buffer size and call RegQueryValueEx again passing
  ' **               : the updated buffer size in the lpcbData parameter. Repeat this until the function succeeds.
  ' **               : You need to maintain a separate variable to keep track of the buffer size, because the
  ' **               : value returned by lpcbData is unpredictable.
  ' ** Return Value  : If the function succeeds, the return value is ERROR_SUCCESS.
  ' **               : If the function fails, the return value is a system error code.
  ' **               : If the lpData buffer is too small to receive the data, the function returns ERROR_MORE_DATA.

' ** Obsolete, replaced by RegSetValueEx().
'Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" _
'  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, _
'  ByVal cbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, _
  lpData As Any, ByVal cbData As Long) As Long
  ' ** hKey           : A handle to an open registry key. The key must have been opened with the KEY_SET_VALUE
  ' **                : access right. For more information, see Registry Key Security and Access Rights.
  ' ** lpValueName    : The name of the value to be set. If a value with this name is not already present in
  ' **                : the key, the function adds it to the key. If lpValueName is NULL or an empty string, "",
  ' **                : the function sets the type and data for the key's unnamed or default value.
  ' **                : For more information, see Registry Element Size Limits. Registry keys do not have
  ' **                : default values, but they can have one unnamed value, which can be of any type.
  ' ** Reserved       : This parameter is reserved and must be zero.
  ' ** dwType         : The type of data pointed to by the lpData parameter. For a list of the possible types,
  ' **                : see Registry Value Types.
  ' ** lpData         : The data to be stored. For string-based types, such as REG_SZ, the string must be
  ' **                : null-terminated. With the REG_MULTI_SZ data type, the string must be terminated with two
  ' **                : null characters.
  ' ** cbData         : The size of the information pointed to by the lpData parameter, in bytes. If the data is
  ' **                : of type REG_SZ, REG_EXPAND_SZ, or REG_MULTI_SZ, cbData must include the size of the
  ' **                : terminating null character or characters.
  ' ** Return Value   : If the function succeeds, the return value is ERROR_SUCCESS.
  ' **                : If the function fails, the return value is a nonzero error code defined in Winerror.h.
  ' **                : You can use the FormatMessage function with the FORMAT_MESSAGE_FROM_SYSTEM flag to get
  ' **                : a generic description of the error.

Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, _
  ByVal lpValue As String, ByVal cbData As Long) As Long

Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, _
  lpValue As Long, ByVal cbData As Long) As Long

' ** PasswordDefault enumeration:
Public Const DEF_CYCLE As Long = 365&
Public Const DEF_SCREEN As Long = 30&
Public Const DEF_MSGBOX As Long = 7&

' ** My arr_varResult() constants:
Public Const RES_ELEMS As Integer = 4  ' ** Array's first-element UBound().
Public Const RES_ERR As Integer = 0
Public Const RES_YN  As Integer = 1
Public Const RES_SUB As Integer = 2
Public Const RES_LNG As Integer = 3
Public Const RES_VAR As Integer = 4

' ** Security group names currently available from TrustSec.mdw:
Public Const SGRP_ADMINS    As String = "Admins"
'Public Const SGRP_USERS     As String = "TAUsers"  ' ** NOT YET!
Public Const SGRP_VIEWONLY  As String = "ViewOnly"
'Public Const SGRP_DEMO      As String = "TADemos"  ' ** NOT YET!
'Public Const SGRP_OLD_USERS As String = "Users"
Public Const SGRP_DATAENTRY As String = "DataEntry"

' ** Array: garr_varHid().
Public glngHids As Long, garr_varHid As Variant
Public Const H_TYP As Integer = 0
Public Const H_ID  As Integer = 1
Public Const H_NAM As Integer = 2
'Public Const H_HID As Integer = 3

Public gstrAccept As String
Public gdatAccept As Date

Private strPath As String
' **

Public Function CurrentGroup(Optional varGroupName As Variant) As String
' ** Currently, Trust Accountant does not set different permissions for its
' ** groups and users using Tools -> Security -> User and Group Permissions...,
' ** though all passwords and broad permission are controlled that way.
' ** Our own levels of permission are just handled using code behind the forms.
' ** Trust Accountant group levels:
' **   Admins     High    All permissions; includes 'superuser'; highest level.
' **   DataEntry  Medium  Restricted from some areas, e.g., Utility Menu, some Posting features.
' **   ViewOnly   Low     No data entry allowed; lowest level.
' **   Users              Not specifically used by us, but required by Access.
' ** A specified group the current user does not belong to will return an empty string.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentGroup"

        Dim grp As DAO.Group, usr As DAO.User
        Dim strGroupName As String
        Dim lngX As Long
        Dim strRetVal As String

110     strRetVal = vbNullString

120     For lngX = 1& To 4&
130       If IsMissing(varGroupName) = True Then
            ' ** Highest level current user belongs to.
140         Select Case lngX
            Case 1&
150           strGroupName = "Admins"
160         Case 2&
170           strGroupName = "DataEntry"
180         Case 3&
190           strGroupName = "ViewOnly"
200         Case 4&
210           strGroupName = "Users"
220         End Select
230       Else
            ' ** Specified group.
240         strGroupName = varGroupName
250       End If

260       For Each grp In DBEngine.Workspaces(0).Groups
270         For Each usr In grp.Users
280           If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
290             If grp.Name = strGroupName Then
300               strRetVal = grp.Name
310               Exit For
320             End If
330           End If
340         Next
350         If strRetVal <> vbNullString Then Exit For
360       Next
370       If strRetVal <> vbNullString Then Exit For

380     Next

EXITP:
390     Set usr = Nothing
400     Set grp = Nothing
410     CurrentGroup = strRetVal
420     Exit Function

ERRH:
430     strRetVal = vbNullString
440     Select Case ERR.Number
        Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Function

Public Function Permissions_List() As Boolean
' ** Permissions property:
' ** The setting or return value is a Long constant that establishes permissions.
' ** The following tables list the valid constants for the Permissions property of various DAO objects.
' ** Unless otherwise noted, all constants shown in all tables are valid for Document objects.

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Permissions_List"

        Dim dbs As DAO.Database
        Dim cntr As DAO.Container
        Dim blnRetVal As Boolean

510     blnRetVal = True

        ' ** Enumerate Containers collection and display the current
        ' ** user and the permissions set for that user.
520     Set dbs = CurrentDb
530     With dbs
540       For Each cntr In .Containers
550         With cntr
560           Debug.Print "'Container: " & Left(.Name & Space(15), 15) & "  " & "User: " & .Username
570           Debug.Print "'  Permissions:    " & Permissions_Ref(.Permissions, .Name)     ' ** User-specific permissions.
580           Debug.Print "'  AllPermissions: " & Permissions_Ref(.AllPermissions, .Name)  ' ** Includes inherited Group permissions.
590         End With
600       Next
610       .Close
620     End With

        'Container: DataAccessPages  User: superuser
        '  Permissions:    dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Databases        User: superuser
        '  Permissions:    dbSecReadSec, dbSecWriteSec, dbSecFullAccess
        '  AllPermissions: dbSecReadSec, dbSecWriteSec, dbSecFullAccess
        'Container: Forms            User: superuser
        '  Permissions:    dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Modules          User: superuser
        '  Permissions:    dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Relationships    User: superuser
        '  Permissions:    dbSecReadSec, dbSecWriteSec, dbSecFullAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Reports          User: superuser
        '  Permissions:    dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Scripts          User: superuser
        '  Permissions:    dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: SysRel           User: superuser
        '  Permissions:    dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Tables           User: superuser
        '  Permissions:    dbSecCreate, dbSecReadDef, dbSecRetrieveData, dbSecInsertData, dbSecReplaceData, dbSecDeleteData, dbSecWriteDef
        '  AllPermissions: dbSecCreate, dbSecReadDef, dbSecRetrieveData, dbSecInsertData, dbSecReplaceData, dbSecDeleteData, dbSecWriteDef

        'Container: DataAccessPages  User: admin
        '  Permissions:    dbSecNoAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Databases        User: admin
        '  Permissions:    dbSecNoAccess
        '  AllPermissions: dbSecNoAccess
        'Container: Forms            User: admin
        '  Permissions:    dbSecNoAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Modules          User: admin
        '  Permissions:    dbSecNoAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Relationships    User: admin
        '  Permissions:    dbSecNoAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Reports          User: admin
        '  Permissions:    dbSecNoAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Scripts          User: admin
        '  Permissions:    dbSecNoAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: SysRel           User: admin
        '  Permissions:    dbSecNoAccess
        '  AllPermissions: dbSecDelete, dbSecReadSec, dbSecWriteSec, dbSecWriteOwner, dbSecFullAccess
        'Container: Tables           User: admin
        '  Permissions:
        '  AllPermissions: dbSecCreate, dbSecReadDef, dbSecRetrieveData, dbSecInsertData, dbSecReplaceData, dbSecDeleteData, dbSecWriteDef

EXITP:
630     Set cntr = Nothing
640     Set dbs = Nothing
650     Permissions_List = blnRetVal
660     Exit Function

ERRH:
670     blnRetVal = False
680     Select Case ERR.Number
        Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Function

Private Function Permissions_Ref(lngPerm As Long, strContainer As String) As String

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Permissions_Ref"

        Dim strRetVal As String

810     strRetVal = vbNullString

820     Select Case strContainer
        Case "Tables"
          ' ** The following table lists the possible settings and return values for the Tables container.
          ' **       1  dbSecCreate        The user can create new documents (not valid for Document objects).
          ' **       4  dbSecReadDef       The user can read the table definition, including column and index information.
          ' **      20  dbSecRetrieveData  The user can retrieve data from the Document object.
          ' **      32  dbSecInsertData    The user can add records.
          ' **      64  dbSecReplaceData   The user can modify records.
          ' **     128  dbSecDeleteData    The user can delete records.
          ' **   65548  dbSecWriteDef      The user can modify or delete the table definition, including column and index information.
830       If (lngPerm And dbSecCreate) > 0 Then strRetVal = strRetVal & "dbSecCreate, "
840       If (lngPerm And dbSecReadDef) > 0 Then strRetVal = strRetVal & "dbSecReadDef, "
850       If (lngPerm And dbSecRetrieveData) > 0 Then strRetVal = strRetVal & "dbSecRetrieveData, "
860       If (lngPerm And dbSecInsertData) > 0 Then strRetVal = strRetVal & "dbSecInsertData, "
870       If (lngPerm And dbSecReplaceData) > 0 Then strRetVal = strRetVal & "dbSecReplaceData, "
880       If (lngPerm And dbSecDeleteData) > 0 Then strRetVal = strRetVal & "dbSecDeleteData, "
890       If (lngPerm And dbSecWriteDef) > 0 Then strRetVal = strRetVal & "dbSecWriteDef, "
          'Case "Databases"  'THESE VALUES SEEM TO BE WRONG!!!
          '  ' ** The following table lists the possible settings and return values for the Databases container.
          '  ' **   1  dbSecDBCreate     The user can create new databases. This option is valid only on the Databases container in the workgroup information file (Systen.mdw). This constant isn't valid for Document objects.
          '  ' **   2  dbSecDBOpen       The user can open the database.
          '  ' **   4  dbSecDBExclusive  The user has exclusive access to the database.
          '  ' **   8  dbSecDBAdmin      The user can replicate a database and change the database password (not valid for Document objects).
          '  If (lngPerm And dbSecDBCreate) > 0 Then strRetVal = strRetVal & "dbSecDBCreate, "
          '  If (lngPerm And dbSecDBOpen) > 0 Then strRetVal = strRetVal & "dbSecDBOpen, "
          '  If (lngPerm And dbSecDBExclusive) > 0 Then strRetVal = strRetVal & "dbSecDBExclusive, "
          '  If (lngPerm And dbSecDBAdmin) > 0 Then strRetVal = strRetVal & "dbSecDBAdmin, "
900     Case Else
          ' ** The following table lists possible values for Container objects other than Tables and Databases containers.
          ' **         0  dbSecNoAccess    The user doesn't have access to the object (not valid for Document objects).
          ' **     65536  dbSecDelete      The user can delete the object.
          ' **    131072  dbSecReadSec     The user can read the object's security-related information.
          ' **    262144  dbSecWriteSec    The user can alter access permissions.
          ' **    524288  dbSecWriteOwner  The user can change the Owner property setting.
          ' **   1048575  dbSecFullAccess  The user has full access to the object.
910       If lngPerm = 0 Then strRetVal = strRetVal & "dbSecNoAccess, "
920       If (lngPerm And dbSecDelete) > 0 Then strRetVal = strRetVal & "dbSecDelete, "
930       If (lngPerm And dbSecReadSec) > 0 Then strRetVal = strRetVal & "dbSecReadSec, "
940       If (lngPerm And dbSecWriteSec) > 0 Then strRetVal = strRetVal & "dbSecWriteSec, "
950       If (lngPerm And dbSecWriteOwner) > 0 Then strRetVal = strRetVal & "dbSecWriteOwner, "
960       If (lngPerm And dbSecFullAccess) > 0 Then strRetVal = strRetVal & "dbSecFullAccess, "
970     End Select

980     strRetVal = Trim(strRetVal)
990     If Right(strRetVal, 1) = "," Then strRetVal = Left(strRetVal, (Len(strRetVal) - 1))

EXITP:
1000    Permissions_Ref = strRetVal
1010    Exit Function

ERRH:
1020    Select Case ERR.Number
        Case Else
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1040    End Select
1050    Resume EXITP

End Function

Public Function Security_SyncChk() As Boolean
' ** Cross-Checks all security settings among the MDW file,
' ** Users table, the tblSecurity_.. trio, and the _~xusr table.

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Security_SyncChk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim wrkLnk As DAO.Workspace, dbsLnk As DAO.Database, tdf As DAO.TableDef
        Dim grp As DAO.Group, usr As DAO.User
        Dim lngGrps As Long, arr_varGrp As Variant, lngCurrGrps As Long
        Dim lngUsrs As Long, arr_varUsr As Variant, lngUserID As Long
        Dim lngTAUsrs As Long, arr_varTAUsr As Variant
        Dim lngExtUsers As Long, arr_varExtUser() As Variant
        Dim blnGrpMismatch As Boolean, blnUsrMismatch As Boolean, blnExtraMDW As Boolean
        Dim strUsername As String, strUserPIN As String, strUserPW As String
        Dim lngUsersGrpID As Long, lngAdminsGrpID As Long
        Dim blnFound As Boolean, lngRecs1 As Long, lngRecs2 As Long
        Dim lngUserCreates As Long, arr_varUserCreate() As Variant
        Dim lngZetas As Long, arr_varZeta() As Variant
        Dim blnOneOfOurs As Boolean
        Dim strErrMsg As String
        Dim intPos01 As Integer
        Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varGrp().
        Const GRP_ID    As Integer = 0
        Const GRP_NAME  As Integer = 1
        Const GRP_FOUND As Integer = 2

        ' ** Array: arr_varUsr().
        Const USR_ID    As Integer = 0
        Const USR_NAME  As Integer = 1
        Const USR_DEF   As Integer = 2
        Const USR_GRPS  As Integer = 3
        Const USR_GRP1  As Integer = 4
        Const USR_GRP1D As Integer = 5
        Const USR_GRP2  As Integer = 6
        Const USR_GRP2D As Integer = 7
        Const USR_FND1  As Integer = 8
        Const USR_FND2  As Integer = 9
        Const USR_PW    As Integer = 10

        ' ** Array: arr_varTAUsr().
        Const TAU_GUID  As Integer = 0
        Const TAU_NAME  As Integer = 1
        Const TAU_EMP   As Integer = 2
        'Const TAU_DGRP  As Integer = 3
        'Const TAU_GRP1  As Integer = 4
        Const TAU_PGRP  As Integer = 5
        Const TAU_GRP2  As Integer = 6
        Const TAU_FND1  As Integer = 7
        Const TAU_FND2  As Integer = 8
        Const TAU_FND3  As Integer = 9
        Const TAU_PW    As Integer = 10
        Const TAU_OLDN  As Integer = 11

        ' ** Array: arr_varExtUser().
        Const XUSR_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const XUSR_NAME As Integer = 0

        ' ** Array: arr_varUserCreate().
        Const UC_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const UC_NAME As Integer = 0
        Const UC_PW   As Integer = 1
        Const UC_GUID As Integer = 2
        Const UC_OURS As Integer = 3

        ' ** Array: arr_varZeta().
        Const Z_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const Z_ID  As Integer = 0
        Const Z_FND As Integer = 1

1110    blnRetVal = True

        ' ** Groups():
        ' **   0 Admins
        ' **   1 DataEntry
        ' **   2 Users
        ' **   3 ViewOnly

        ' ** Empty MDW Users():
        ' **   0 admin
        ' **   1 Creator
        ' **   2 Engine
        ' **   3 superuser

        ' ** Test MDW Users():
        ' **   0 admin
        ' **   1 Creator
        ' **   2 Engine
        ' **   3 Bill
        ' **   4 Mary
        ' **   5 superuser

1120    blnGrpMismatch = False: blnUsrMismatch = False: blnExtraMDW = False

1130    lngUserCreates = 0&
1140    ReDim arr_varUserCreate(UC_ELEMS, 0)

1150    Set dbs = CurrentDb

        ' ** Update the Users table to make sure all have a Default Group of 'Users', as well as a Primary Group entry.
1160    Set qdf = dbs.QueryDefs("qrySecurity_Sync_06a")  ' ** Null or vbNullString.
1170    qdf.Execute dbFailOnError
1180    Set qdf = dbs.QueryDefs("qrySecurity_Sync_06b")  ' ** Other than 'Users'.
1190    qdf.Execute dbFailOnError
1200    Set qdf = dbs.QueryDefs("qrySecurity_Sync_06c")  ' ** No entry to 'ViewOnly'.
1210    qdf.Execute dbFailOnError

        ' ** Collect the Groups from tblSecurity_Group into arr_varGrp().
1220    Set qdf = dbs.QueryDefs("qrySecurity_Sync_01")
1230    Set rst1 = qdf.OpenRecordset
1240    With rst1
1250      .MoveLast
1260      lngGrps = .RecordCount
1270      .MoveFirst
1280      arr_varGrp = .GetRows(lngGrps)
          ' ***************************************************
          ' ** Array: arr_varGrp()
          ' **
          ' **   Field  Element  Name             Constant
          ' **   =====  =======  ===============  ===========
          ' **     1       0     secgrp_id        GRP_ID
          ' **     2       1     secgrp_name      GRP_NAME
          ' **     3       2     Found (False)    GRP_FOUND
          ' **
          ' ***************************************************
1290      .Close
1300    End With

        ' ** Save the 2 primary group IDs.
1310    For lngX = 0& To (lngGrps - 1&)
1320      If arr_varGrp(GRP_NAME, lngX) = "Users" Then
1330        lngUsersGrpID = arr_varGrp(GRP_ID, lngX)
1340      ElseIf arr_varGrp(GRP_NAME, lngX) = "Admins" Then
1350        lngAdminsGrpID = arr_varGrp(GRP_ID, lngX)
1360      End If
1370    Next

        ' ** Collect the Users from tblSecurity_User into arr_varUsr().
1380    Set qdf = dbs.QueryDefs("qrySecurity_Sync_02")
1390    Set rst1 = qdf.OpenRecordset
1400    With rst1
1410      .MoveLast
1420      lngUsrs = .RecordCount
1430      .MoveFirst
1440      arr_varUsr = .GetRows(lngUsrs)
          ' ****************************************************
          ' ** Array: arr_varUsr()
          ' **
          ' **   Field  Element  Name              Constant
          ' **   =====  =======  ================  ===========
          ' **     1       0     secusr_id         USR_ID
          ' **     2       1     secusr_name       USR_NAME
          ' **     3       2     secusr_default    USR_DEF
          ' **     4       3     grps              USR_GRPS
          ' **     5       4     grp1              USR_GRP1
          ' **     6       5     grp1_done         USR_GRP1D
          ' **     7       6     grp2              USR_GRP2
          ' **     8       7     grp2_done         USR_GRP2D
          ' **     9       8     Found1 (False)    USR_FND1  {tblSecurity_User not in TrustSec.mdw}
          ' **    10       9     Found2 (False)    USR_FND2  {tblSecurity_User not in Users}
          ' **    11      10     sec_password      USR_PW
          ' **
          ' ****************************************************
1450      .Close
1460    End With

        ' ** Add their Group membership.
1470    Set qdf = dbs.QueryDefs("qrySecurity_Sync_03")
1480    Set rst1 = qdf.OpenRecordset
1490    With rst1
1500      .MoveLast
1510      lngRecs1 = .RecordCount
1520      .MoveFirst
1530      For lngX = 1& To lngRecs1
1540        For lngY = 0& To (lngUsrs - 1&)
              ' ** Trust Accountant only allows for a max of 2 Groups.
1550          If arr_varUsr(USR_ID, lngY) = ![secusr_id] Then
1560            If arr_varUsr(USR_GRP1, lngY) = 0& Then
1570              arr_varUsr(USR_GRP1, lngY) = ![secgrp_id]
1580              arr_varUsr(USR_GRPS, lngY) = 1&
1590            ElseIf arr_varUsr(USR_GRP2, lngY) = 0& Then
1600              arr_varUsr(USR_GRP2, lngY) = ![secgrp_id]
1610              arr_varUsr(USR_GRPS, lngY) = 2&
1620            End If
1630            If arr_varUsr(USR_DEF, lngY) = True And ![secgrpusr_default] = False Then
1640              .Edit
1650              ![secgrpusr_default] = True
1660              .Update
1670            End If
1680            Exit For
1690          End If
1700        Next
1710        If lngX < lngRecs1 Then .MoveNext
1720      Next
1730      .Close
1740    End With

        ' ** Collect the Trust Accountant Users into arr_varTAUser().
1750    Set qdf = dbs.QueryDefs("qrySecurity_Sync_04")
1760    Set rst1 = qdf.OpenRecordset
1770    With rst1
1780      .MoveLast
1790      lngTAUsrs = .RecordCount
1800      .MoveFirst
1810      arr_varTAUsr = .GetRows(lngTAUsrs)
          ' ***************************************************
          ' ** Array: arr_varTAUsr()
          ' **
          ' **   Field  Element  Name             Constant
          ' **   =====  =======  ===============  ===========
          ' **     1       0     s_GUID           TAU_GUID
          ' **     2       1     Username         TAU_NAME
          ' **     3       2     Employee Name    TAU_EMP
          ' **     4       3     Default Group    TAU_DGRP
          ' **     5       4     secgrp_id1       TAU_GRP1
          ' **     6       5     Primary Group    TAU_PGRP
          ' **     7       6     secgrp_id2       TAU_GRP2
          ' **     8       7     Found1 (False)   TAU_FND1  {Trust Accountant User not in TrustSec.mdw}
          ' **     9       8     Found2 (False)   TAU_FND2  {Trust Accountant User not in tblSecurity_User}
          ' **    10       9     Found3 (False)   TAU_FND3  {Users not in _~xusr}
          ' **    11      10     sec_password     TAU_PW
          ' **    12      11     sec_oldname      TAU_OLDN
          ' **
          ' ***************************************************
1820      .Close
1830    End With

        ' ** If superuser isn't in the Trust Accountant Users table, there's something really wrong!
1840    blnFound = False
1850    For lngX = 0& To (lngTAUsrs - 1&)
1860      If arr_varTAUsr(TAU_NAME, lngX) = "superuser" Then
1870        blnFound = True
1880        Exit For
1890      End If
1900    Next
1910    If blnFound = False Then
1920      DoCmd.Hourglass False
1930      blnRetVal = False
1940      Beep
1950      MsgBox "The Users table is invalid!" & vbCrLf & vbCrLf & "Trust Accountant cannot continue." & vbCrLf & _
            "Contact Delta Data, Inc., for assistance.", vbCritical + vbOKOnly, "Invalid User Table"
1960      DoCmd.Hourglass True
1970    Else

          ' ** Make sure 'Creator' and 'Engine' are now also in the Users table.
1980      blnFound = False
1990      For lngX = 0& To (lngTAUsrs - 1&)
2000        If arr_varTAUsr(TAU_NAME, lngX) = "Creator" Then
2010          blnFound = True
2020          Exit For
2030        End If
2040      Next
2050      If blnFound = False Then
            ' ** Append 'Creator' to Users table.
2060        Set qdf = dbs.QueryDefs("qrySecurity_Sync_08")
2070        qdf.Execute dbFailOnError
2080      End If
2090      blnFound = False
2100      For lngX = 0& To (lngTAUsrs - 1&)
2110        If arr_varTAUsr(TAU_NAME, lngX) = "Engine" Then
2120          blnFound = True
2130          Exit For
2140        End If
2150      Next
2160      If blnFound = False Then
            ' ** Append 'Engine' to Users table.
2170        Set qdf = dbs.QueryDefs("qrySecurity_Sync_09")
2180        qdf.Execute dbFailOnError
2190      End If

          ' ** First, check if the tblSecurity_Group matches the connected TrustSet.mdw.
2200      lngCurrGrps = DBEngine.Workspaces(0).Groups.Count
          ' ** There's confusion involving the Trust Import group 'Importer'.
2210      blnFound = False
2220      If (lngCurrGrps = lngGrps) Or (lngCurrGrps = (lngGrps - 1&)) Or (lngCurrGrps = (lngGrps + 1&)) Then  ' ** Cover all the bases.
2230        blnFound = True
2240      End If
2250      Select Case blnFound
          Case True

2260        For lngX = 0& To (lngGrps - 1&)
2270          blnFound = False
2280          For Each grp In DBEngine.Workspaces(0).Groups
2290            With grp
2300              If .Name = arr_varGrp(GRP_NAME, lngX) Then
2310                blnFound = True
2320                arr_varGrp(GRP_FOUND, lngX) = True
2330                Exit For
2340              End If
2350            End With
2360          Next
2370          If arr_varGrp(GRP_NAME, lngX) = "Importer" And blnFound = False Then  ' ** Let's see if this takes care of it.
2380            blnFound = True
2390          End If
2400          If blnFound = False Then Exit For  ' ** A default group wasn't found!
2410        Next

2420        Select Case blnFound
            Case True

              ' ** Second, check users in tblSecurity_User against those in TrustSec.mdw.
2430          For lngX = 0& To (lngUsrs - 1&)
2440            blnFound = False
2450            For Each usr In DBEngine.Workspaces(0).Users
2460              With usr
2470                If .Name = arr_varUsr(USR_NAME, lngX) Then
2480                  arr_varUsr(USR_FND1, lngX) = True
2490                  blnFound = True
2500                  Exit For
2510                End If
2520              End With
2530            Next
2540          Next

              ' ** Third, check users in tblSecurity_User against those in the Trust Accountant Users table.
2550          For lngX = 0& To (lngUsrs - 1&)
2560            blnFound = False
2570            For lngY = 0& To (lngTAUsrs - 1&)
2580              If arr_varTAUsr(TAU_NAME, lngY) = arr_varUsr(USR_NAME, lngX) Then
2590                arr_varUsr(USR_FND2, lngX) = True
2600                blnFound = True
2610                Exit For
2620              End If
2630            Next
2640          Next

              ' ** Fourth, check users in the Trust Accountant Users table against those in TrustSec.mdw.
2650          For lngX = 0& To (lngTAUsrs - 1&)
2660            blnFound = False
2670            For Each usr In DBEngine.Workspaces(0).Users
2680              With usr
2690                If .Name = arr_varTAUsr(TAU_NAME, lngX) Then
2700                  arr_varTAUsr(TAU_FND1, lngX) = True
2710                  blnFound = True
2720                  Exit For
2730                End If
2740              End With
2750            Next
2760          Next

              ' ** Fifth, check users in the Trust Accountant Users table against those in tblSecurity_User.
2770          For lngX = 0& To (lngTAUsrs - 1&)
2780            blnFound = False
2790            For lngY = 0& To (lngUsrs - 1&)
2800              If arr_varUsr(USR_NAME, lngY) = arr_varTAUsr(TAU_NAME, lngX) Then
2810                arr_varTAUsr(TAU_FND2, lngX) = True
2820                blnFound = True
2830                Exit For
2840              End If
2850            Next
2860          Next

              ' ** Sixth, check users in tblSecurity_User that weren't found elsewhere.
2870          For lngX = 0& To (lngUsrs - 1&)
2880            If arr_varUsr(USR_FND1, lngX) = False Or arr_varUsr(USR_FND2, lngX) = False Then
2890              If arr_varUsr(USR_NAME, lngX) <> "admin" And arr_varUsr(USR_NAME, lngX) <> "Engine" And _
                      arr_varUsr(USR_NAME, lngX) <> "Creator" Then
                    ' ** Only need one hit to check it later.
2900                blnUsrMismatch = True
2910                Exit For
2920              End If
2930            End If
2940          Next

              ' ** Seventh, check users in the Trust Accountant Users table that weren't found elsewhere.
2950          For lngX = 0& To (lngTAUsrs - 1&)
2960            If arr_varTAUsr(TAU_FND1, lngX) = False Or arr_varTAUsr(TAU_FND2, lngX) = False Then
                  ' ** Only need one hit to check it later.
2970              blnUsrMismatch = True
2980              Exit For
2990            End If
3000          Next

              ' ** Finally, check for additional users in the connected TrustSet.mdw against those in tblSecurity_User.
3010          blnFound = False
3020          For Each usr In DBEngine.Workspaces(0).Users
3030            blnFound = False
3040            With usr
3050              For lngX = 0& To (lngUsrs - 1&)
3060                If .Name = arr_varUsr(USR_NAME, lngX) Then
3070                  blnFound = True
3080                  Exit For
3090                End If
3100              Next
3110            End With
3120            If blnFound = False Then
3130              blnExtraMDW = True
3140              Exit For
3150            End If
3160          Next

              ' ** And after finally, check the new _~xusr table.
3170          blnFound = False
3180          For Each tdf In dbs.TableDefs
3190            With tdf
3200              If .Name = "_~xusr" Then
3210                blnFound = True
3220                Exit For
3230              End If
3240            End With
3250          Next

              ' ** Though it should always be there, and InitializeTables() should make
              ' ** sure of that, create it now anyway, because it's such a new feature.
              ' ** REMEMBER: The s_GUID field is an AutoNumber, and this GUID is the
              ' ** link between the _~xusr and Users table, so the Users table should
              ' ** never be regenerated, or default users deleted, without corresponding
              ' ** changes in the _~xusr table!
3260          If blnFound = False Then

                ' ** Open TrustDta.mdb directly.
3270  On Error Resume Next
3280            Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
3290            If ERR.Number <> 0 Then
3300  On Error GoTo ERRH
'3370  On Error GoTo 0
3310  On Error Resume Next
3320              Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
3330              If ERR.Number <> 0 Then
3340  On Error GoTo ERRH
'3410  On Error GoTo 0
3350  On Error Resume Next
3360                Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
3370                If ERR.Number <> 0 Then
3380  On Error GoTo ERRH
'3450  On Error GoTo 0
3390  On Error Resume Next
3400                  Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
3410                  If ERR.Number <> 0 Then
3420  On Error GoTo ERRH
'3490  On Error GoTo 0
3430  On Error Resume Next
3440                    Set wrkLnk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
3450                    If ERR.Number <> 0 Then
3460  On Error GoTo ERRH
'3530  On Error GoTo 0
3470  On Error Resume Next
3480                      Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
3490                      If ERR.Number <> 0 Then
3500  On Error GoTo ERRH
'3570  On Error GoTo 0
3510  On Error Resume Next
3520                        Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
3530  On Error GoTo ERRH
'3600  On Error GoTo 0
3540                      Else
3550  On Error GoTo ERRH
'3620  On Error GoTo 0
3560                      End If
3570                    Else
3580  On Error GoTo ERRH
'3650  On Error GoTo 0
3590                    End If
3600                  Else
3610  On Error GoTo ERRH
'3680  On Error GoTo 0
3620                  End If
3630                Else
3640  On Error GoTo ERRH
'3710  On Error GoTo 0
3650                End If
3660              Else
3670  On Error GoTo ERRH
'3740  On Error GoTo 0
3680              End If
3690            Else
3700  On Error GoTo ERRH
'3770  On Error GoTo 0
3710            End If

                ' ** Check to see if the table is there, just not linked.
3720            With wrkLnk
3730              Set dbsLnk = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
3740              With dbsLnk
3750                For Each tdf In .TableDefs
3760                  With tdf
3770                    If .Name = "_~xusr" Then
3780                      blnFound = True
3790                      Exit For
3800                    End If
3810                  End With
3820                Next
3830              End With
3840            End With

                ' ** Copy the new _~xusr table to their TrustDta.mdb.
3850            If blnFound = False Then
3860              If Len(TA_SEC) < Len(TA_SEC2) Then
                    ' ** tblTemplate_Zeta1 : _~xusr for New and Upgrade versions.
                    ' ** tblTemplate_Zeta2 : Users for New and Upgrade versions.
3870                DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblTemplate_Zeta1", "_~xusr", False  ' ** Copy with data.
3880              Else
                    ' ** tblTemplate_Zeta3 : _~xusr for Demo version.
                    ' ** tblTemplate_Zeta4 : Users for Demo version.
3890                DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblTemplate_Zeta3", "_~xusr", False  ' ** Copy with data.
3900              End If
3910              dbsLnk.TableDefs.Refresh
3920            End If

3930            dbsLnk.Close
3940            wrkLnk.Close

                ' ** Link _~xusr to here.
3950            DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "_~xusr", "_~xusr"

3960            CurrentDb.TableDefs.Refresh
3970            CurrentDb.TableDefs("_~xusr").RefreshLink

3980          End If  ' ** blnFound.

3990          lngZetas = 0&
4000          ReDim arr_varZeta(Z_ELEMS, 0)

              ' ** First, make sure all the defaults are there.
4010          Set rst1 = dbs.OpenRecordset("_~xusr", dbOpenDynaset, dbConsistent)
4020          If rst1.BOF = True And rst1.EOF = True Then
                ' ** Could be one of my early setups, or... Oh, just do it!
4030            rst1.Close
4040            If Len(TA_SEC) < Len(TA_SEC2) Then
                  ' ** Append all defaults from tblTemplate_Zeta1 to _~xusr.
4050              Set qdf = dbs.QueryDefs("qryXUsr_02a")
4060              qdf.Execute dbFailOnError
4070            Else
                  ' ** Append all defaults from tblTemplate_Zeta1 to _~xusr, Demo.
4080              Set qdf = dbs.QueryDefs("qryXUsr_02a")
4090              qdf.Execute dbFailOnError
4100            End If
4110          Else
4120            rst1.MoveLast
4130            lngRecs1 = rst1.RecordCount
4140            If Len(TA_SEC) < Len(TA_SEC2) Then
4150              Set rst2 = dbs.OpenRecordset("tblTemplate_Zeta1", dbOpenDynaset, dbConsistent)
4160            Else
4170              Set rst2 = dbs.OpenRecordset("tblTemplate_Zeta3", dbOpenDynaset, dbConsistent)
4180            End If
4190            rst2.MoveLast
4200            lngRecs2 = rst2.RecordCount
4210            rst2.MoveFirst
4220            For lngX = 1& To lngRecs2
4230              varTmp00 = FilterGUIDString(StringFromGUID(rst2![s_GUID]))  ' ** Module Function: modCodeUtilities.
4240              blnFound = False
4250              rst1.MoveFirst
4260              For lngY = 1& To lngRecs1
4270                If FilterGUIDString(StringFromGUID(rst1![s_GUID])) = varTmp00 Then  ' ** Module Function: modCodeUtilities.
4280                  blnFound = True
4290                  Exit For
4300                End If
4310                If lngY < lngRecs1 Then rst1.MoveNext
4320              Next
4330              If blnFound = False Then
4340                lngZetas = lngZetas + 1&
4350                lngE = lngZetas - 1&
4360                ReDim Preserve arr_varZeta(Z_ELEMS, lngE)
4370                arr_varZeta(Z_ID, lngE) = rst2![xusr_id]
4380                arr_varZeta(Z_FND, lngE) = CBool(False)
4390              End If
4400              If lngX < lngRecs2 Then rst2.MoveNext
4410            Next
4420            rst1.Close
4430            rst2.Close
4440            If lngZetas > 0& Then
4450              For lngX = 0& To (lngZetas - 1&)
4460                If Len(TA_SEC) < Len(TA_SEC2) Then
                      ' ** Append backup to live, by specified [xusr].
4470                  Set qdf = dbs.QueryDefs("qryXUsr_03a")
4480                Else
                      ' ** Append backup to live, by specified [xusr], Demo.
4490                  Set qdf = dbs.QueryDefs("qryXUsr_03b")
4500                End If
4510                With qdf.Parameters
4520                  ![xusr] = arr_varZeta(Z_ID, lngX)
4530                End With
4540                qdf.Execute dbFailOnError
4550                Set qdf = Nothing
4560              Next
4570            End If
4580          End If

              ' ** Coordination between their Users table and _~xusr happens last, below.
4590          blnFound = True

4600        Case False
              ' ** A default Group wasn't found!
              ' ** Since they should always be identical, any not found, either way, means a problem!
4610          DoCmd.Hourglass False
4620          blnRetVal = False
4630          blnGrpMismatch = True
4640          Beep
4650          MsgBox "The Security file is invalid!" & vbCrLf & vbCrLf & "Trust Accountant cannot continue." & vbCrLf & _
                "Contact Delta Data, Inc., for assistance." & vbCrLf & "1", vbCritical + vbOKOnly, "Invalid Security File"
4660          DoCmd.Hourglass False
4670        End Select  ' ** blnFound, blnGrpMismatch.

            ' ** Make sure the new security table is hidden.
4680        If blnRetVal = True Then
4690          If Application.GetHiddenAttribute(acTable, "_~xusr") = False Then
4700            Application.SetHiddenAttribute acTable, "_~xusr", True
4710          End If
4720        End If

4730        If blnGrpMismatch = True Or blnUsrMismatch = True Or blnExtraMDW = True Then

              'Beep
              'MsgBox "Trust Accountant Users do not match the security file!" & vbCrLf & vbCrLf & _
              "Please wait while the two are synchronized.", vbCritical + vbOKOnly, "Security Out Of Sync"

              ' ** Make sure the Access MDW being used is TrustSec.mdw.
4740          strTmp01 = DBEngine.SystemDB
4750          If Right(strTmp01, Len(gstrFile_SecurityName)) <> gstrFile_SecurityName Then
4760            DoCmd.Hourglass False
4770            blnRetVal = False
4780            strTmp02 = Dir(gstrTrustDataLocation & gstrFile_SecurityName)
4790            If strTmp02 = vbNullString Then strTmp02 = gstrFile_SecurityName & " NOT FOUND!"
4800            Beep
4810            MsgBox "An invalid security file is being used!" & vbCrLf & vbCrLf & _
                  "Current, Invalid File:" & vbCrLf & "  " & strTmp01 & vbCrLf & vbCrLf & _
                  "Valid File:" & vbCrLf & "  " & strTmp02 & vbCrLf & vbCrLf & _
                  "Trust Accountant cannot continue." & vbCrLf & _
                  "Contact Delta Data, Inc., for assistance.", vbCritical + vbOKOnly, "Invalid Security File"
4820            DoCmd.Hourglass True
4830          Else

4840            If blnGrpMismatch = False And (blnUsrMismatch = True Or blnExtraMDW = True) Then
                  ' ** If only the users are out of sync, just fix those.

4850              If blnUsrMismatch = True Then
                    ' ** The Trust Accountant Users table should take precedence.

                    ' ** Check for default users first.
4860                For lngX = 0& To (lngUsrs - 1&)
4870                  If arr_varUsr(USR_DEF, lngX) = True And arr_varUsr(USR_FND1, lngX) = False Then
4880                    strUsername = arr_varUsr(USR_NAME, lngX)
4890                    If strUsername <> "Creator" And strUsername <> "Engine" Then
                          ' ** I don't know anything about these, so better not mess with them.

4900                      strUserPIN = arr_varUsr(USR_NAME, lngX)
                          ' ** Create new user account.
4910                      Set usr = DBEngine.Workspaces(0).CreateUser(strUsername, strUserPIN)
                          ' ** Save user account definition by appending it to Users collection.
4920                      DBEngine.Workspaces(0).Users.Append usr

                          ' ** Add default user to predefined, Users group.
4930                      blnFound = False
4940                      If arr_varUsr(USR_GRP1, lngX) = lngUsersGrpID Then
4950                        blnFound = True
4960                        arr_varUsr(USR_GRP1D, lngX) = True
4970                      ElseIf arr_varUsr(USR_GRP2, lngX) = lngUsersGrpID Then
4980                        blnFound = True
4990                        arr_varUsr(USR_GRP2D, lngX) = True
5000                      End If
5010                      If blnFound = True Then
5020                        usr.Groups.Append usr.CreateGroup("Users")
5030                      Else
                            ' ** A default user isn't in tblSecurity_GroupUser for the Users group!
5040                        Set qdf = dbs.QueryDefs("qrySecurity_Sync_05")
5050                        With qdf.Parameters
5060                          ![grpid] = lngUsersGrpID
5070                          ![usrid] = arr_varUsr(USR_ID, lngX)
5080                          ![IsDef] = True
5090                        End With
5100                        qdf.Execute dbFailOnError
5110                      End If

                          ' ** Add default user to predefined, Admins group.
5120                      blnFound = False
5130                      If arr_varUsr(USR_GRP1, lngX) = lngAdminsGrpID Then
5140                        blnFound = True
5150                        arr_varUsr(USR_GRP1D, lngX) = True
5160                      ElseIf arr_varUsr(USR_GRP2, lngX) = lngAdminsGrpID Then
5170                        blnFound = True
5180                        arr_varUsr(USR_GRP2D, lngX) = True
5190                      End If
5200                      If blnFound = True Then
5210                        usr.Groups.Append usr.CreateGroup("Admins")
5220                      Else
                            ' ** A default user isn't in tblSecurity_GroupUser for the Admins group!
5230                        Set qdf = dbs.QueryDefs("qrySecurity_Sync_05")
5240                        With qdf.Parameters
5250                          ![grpid] = lngAdminsGrpID
5260                          ![usrid] = arr_varUsr(USR_ID, lngX)
5270                          ![IsDef] = True
5280                        End With
5290                        qdf.Execute dbFailOnError
5300                      End If

5310                      usr.Groups.Refresh
5320                      If strUsername = "admin" Then
                            ' ** WARNING!: This User shouldn't have any rights at all, but
                            ' ** currently it's as open as superuser. For now, don't change that.
5330                        strUserPW = vbNullString
5340                      ElseIf strUsername = "superuser" Then
5350                        strUserPW = TA_SEC
5360                      Else
                            ' ** This shouldn't be hit for default Users!
5370                        intPos01 = InStr(strUserPIN, " ")
5380                        Do While intPos01 > 0
                              ' ** Just remove the space.
5390                          strUserPIN = Left(strUserPIN, (intPos01 - 1)) & Mid(strUserPIN, (intPos01 + 1))
5400                          intPos01 = InStr(strUserPIN, " ")
5410                        Loop
                            ' ** New password requirements.
                            ' **   5-14 characters.
                            ' **   1 Uppercase.
                            ' **   1 Lowercase.
                            ' **   1 Numeral.
                            ' ** Construct a valid password from the user's PIN.
5420                        strUserPW = Security_PIN_To_PW(strUserPIN)  ' ** Function: Below.
5430                      End If
5440                      arr_varUsr(USR_PW, lngX) = strUserPW
5450                      DBEngine.Workspaces(0).Users(strUsername).NewPassword vbNullString, strUserPW

                          ' ** Currently, the only default user required for the Trust Accountant
                          ' ** Users table is superuser, which has already been checked.

5460                    End If
5470                  End If
5480                Next

                    ' ** Check the Username length.
5490                For lngX = 0& To (lngTAUsrs - 1&)
5500                  If Len(arr_varTAUsr(TAU_NAME, lngX)) < 4 Then
5510                    arr_varTAUsr(TAU_OLDN, lngX) = arr_varTAUsr(TAU_NAME, lngX)  ' ** Save the old Username.
5520                    Select Case IsNull(arr_varTAUsr(TAU_EMP, lngX))
                        Case True
5530                      arr_varTAUsr(TAU_NAME, lngX) = Left(Trim(arr_varTAUsr(TAU_NAME, lngX)) & "XXXX", 4)
5540                    Case False
5550                      intPos01 = InStr(Trim(arr_varTAUsr(TAU_EMP, lngX)), " ")
5560                      If intPos01 > 0 Then
5570                        arr_varTAUsr(TAU_NAME, lngX) = Left(arr_varTAUsr(TAU_NAME, lngX) & _
                              Trim(Mid(Trim(arr_varTAUsr(TAU_EMP, lngX)), (intPos01 + 1))), 4)
5580                      Else
5590                        arr_varTAUsr(TAU_NAME, lngX) = Left(Trim(arr_varTAUsr(TAU_NAME, lngX)) & "XXXX", 4)
5600                      End If
5610                    End Select
5620                  End If
5630                Next
                    'MAKE SURE THIS PROPAGATES TO Users AND tblSecurity_User!!

                    ' ** Then check for Trust Accountant Users.
5640                For lngX = 0& To (lngTAUsrs - 1&)

                      ' ** Check against TrustSec.mdw.
5650                  If arr_varTAUsr(TAU_FND1, lngX) = False Then
5660                    strUsername = arr_varTAUsr(TAU_NAME, lngX)
5670                    strUserPIN = arr_varTAUsr(TAU_NAME, lngX)
                        ' ** Create new user account.
5680                    Set usr = DBEngine.Workspaces(0).CreateUser(strUsername, strUserPIN)
                        ' ** Save user account definition by appending it to Users collection.
5690  On Error Resume Next
5700                    DBEngine.Workspaces(0).Users.Append usr
                        '####  WRITE THESE ERRORS!  ####
5710                    If ERR.Number = 0 Then
5720  On Error GoTo ERRH
'5720  On Error GoTo 0
                          ' ** Add user to predefined, default Users group.
5730                      usr.Groups.Append usr.CreateGroup("Users")
                          ' ** Add user to primary group.
                          ' ** Error: 3030  '{no entry}' is not a valid account name.
5740                      If arr_varTAUsr(TAU_PGRP, lngX) = "{no entry}" Then
5750                        If IsNull(arr_varTAUsr(TAU_GRP2, lngX)) = True Then
5760                          arr_varTAUsr(TAU_GRP2, lngX) = CLng(4)
5770                          arr_varTAUsr(TAU_PGRP, lngX) = "ViewOnly"
5780                        Else
5790                          If arr_varTAUsr(TAU_GRP2, lngX) = 0& Then
5800                            arr_varTAUsr(TAU_GRP2, lngX) = CLng(4)
5810                            arr_varTAUsr(TAU_PGRP, lngX) = "ViewOnly"
5820                          Else
5830                            varTmp00 = DLookup("[secgrp_name]", "tblSecurity_Group", "[secgrp_id] = " & CStr(arr_varTAUsr(TAU_GRP2, lngX)))
5840                            If IsNull(varTmp00) = False Then
5850                              arr_varTAUsr(TAU_PGRP, lngX) = varTmp00
5860                            Else
5870                              arr_varTAUsr(TAU_GRP2, lngX) = CLng(4)
5880                              arr_varTAUsr(TAU_PGRP, lngX) = "ViewOnly"
5890                            End If
5900                          End If
5910                        End If
5920                      End If
5930                      usr.Groups.Append usr.CreateGroup(arr_varTAUsr(TAU_PGRP, lngX))
5940                      strUserPIN = LCase$(strUserPIN)
5950                      intPos01 = InStr(strUserPIN, " ")
5960                      Do While intPos01 > 0
                            ' ** Remove the space.
5970                        strUserPIN = Left(strUserPIN, (intPos01 - 1)) & Mid(strUserPIN, (intPos01 + 1))
5980                        intPos01 = InStr(strUserPIN, " ")
5990                      Loop
                          ' ** If this is one of our own users, use the same password.
                          ' **   Admin, Superuser, TADemo, TAImport, mary, bill
6000                      blnOneOfOurs = False
6010                      Select Case strUsername
                          Case "Admin", "Superuser", "TADemo", "TAImport"
6020                        If Len(TA_SEC) < Len(TA_SEC2) Then
                              ' ** New and Upgrade versions.
6030                          Select Case strUsername
                              Case "Admin", "Superuser"
                                ' ** tblTemplate_Zeta2, linked to tblTemplate_Zeta1, for New, Upgrade.
6040                            Set qdf = dbs.QueryDefs("qryXUsr_10")
6050                            Set rst1 = qdf.OpenRecordset
6060                            With rst1
6070                              If .BOF = True And .EOF = True Then
                                    ' ** Unlikely, but seriously wrong!
6080                              Else
6090                                .FindFirst "[Username] = '" & strUsername & "'"
6100                                If .NoMatch = False Then
                                      ' ** Admin, Superuser.
6110                                  blnOneOfOurs = True
6120                                  strUserPW = DecodeString(![xusr_extant])  ' ** Module Function: modCodeUtilities.
6130                                End If
6140                              End If
6150                              .Close
6160                            End With
6170                          Case "TADemo"
                                ' ** tblTemplate_Zeta4, linked to tblTemplate_Zeta3, for Demo.
6180                            Set qdf = dbs.QueryDefs("qryXUsr_11")
6190                            Set rst1 = qdf.OpenRecordset
6200                            With rst1
6210                              If .BOF = True And .EOF = True Then
                                    ' ** Unlikely, but seriously wrong!
6220                              Else
6230                                .FindFirst "[Username] = '" & strUsername & "'"
6240                                If .NoMatch = False Then
                                      ' ** TADemo (most likely my own machine, or Rich's).
6250                                  blnOneOfOurs = True
6260                                  strUserPW = DecodeString(![xusr_extant])  ' ** Module Function: modCodeUtilities.
6270                                End If
6280                              End If
6290                              .Close
6300                            End With
6310                          Case "TAImport"
6320                            Beep
6330                            MsgBox "TAImport not yet set up!", vbInformation + vbOKOnly, "TA Import Unavailable"
6340                          End Select
6350                        Else
                              ' ** Demo version.
6360                          Select Case strUsername
                              Case "Admin", "Superuser", "TADemo"
                                ' ** tblTemplate_Zeta4, linked to tblTemplate_Zeta3, for Demo.
6370                            Set qdf = dbs.QueryDefs("qryXUsr_11")
6380                            Set rst1 = qdf.OpenRecordset
6390                            With rst1
6400                              If .BOF = True And .EOF = True Then
                                    ' ** Unlikely, but seriously wrong!
6410                              Else
6420                                .FindFirst "[Username] = '" & strUsername & "'"
6430                                If .NoMatch = False Then
                                      ' ** Admin, Superuser, TADemo.
6440                                  blnOneOfOurs = True
6450                                  strUserPW = DecodeString(![xusr_extant])  ' ** Module Function: modCodeUtilities.
6460                                End If
6470                              End If
6480                              .Close
6490                            End With
6500                          End Select
6510                        End If
6520                        If blnOneOfOurs = False Then
                              ' ** Construct a valid password from the user's PIN.
6530                          strUserPW = Security_PIN_To_PW(strUserPIN)  ' ** Function: Below.
6540                        End If
6550                      Case "mary", "bill"
                            ' ** Check to see if this is our demo data.
                            'CHECK IF THIS IS BEFORE OR AFTER InitializeTables()!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
6560                        varTmp00 = DLookup("[CoInfo_Name]", "CompanyInformation")
6570                        If IsNull(varTmp00) = False Then
6580                          If Trim(varTmp00) = "North Fork Bank" Then
6590                            varTmp00 = DLookup("[shortname]", "account", "[accountno] = '11'")
6600                            If IsNull(varTmp00) = False Then
6610                              If Trim(varTmp00) = "William B. Johnson Trust" Then
                                    ' ** tblTemplate_Zeta4, linked to tblTemplate_Zeta3, for Demo.
6620                                Set qdf = dbs.QueryDefs("qryXUsr_11")
6630                                Set rst1 = qdf.OpenRecordset
6640                                With rst1
6650                                  If .BOF = True And .EOF = True Then
                                        ' ** Unlikely, but seriously wrong!
6660                                  Else
6670                                    .FindFirst "[Username] = '" & strUsername & "'"
6680                                    If .NoMatch = False Then
                                          ' ** mary, bill.
6690                                      blnOneOfOurs = True
6700                                      strUserPW = DecodeString(![xusr_extant])  ' ** Module Function: modCodeUtilities.
6710                                    End If
6720                                  End If
6730                                  .Close
6740                                End With
                                    ' ** Could also check Users table:
                                    ' **   [Employee Name] = "Mary Smith"
                                    ' **   [Employee Name] = "Bill Larson"
6750                              End If
6760                            End If
6770                          End If
6780                        End If
6790                        If blnOneOfOurs = False Then
                              ' ** Construct a valid password from the user's PIN.
6800                          strUserPW = Security_PIN_To_PW(strUserPIN)  ' ** Function: Below.
6810                        End If
6820                      Case Else
                            ' ** Construct a valid password from the user's PIN.
6830                        strUserPW = Security_PIN_To_PW(strUserPIN)  ' ** Function: Below.
6840                      End Select
6850                      arr_varTAUsr(TAU_PW, lngX) = strUserPW
6860                      DBEngine.Workspaces(0).Users(strUsername).NewPassword vbNullString, strUserPW
6870                      lngUserCreates = lngUserCreates + 1&
6880                      lngE = lngUserCreates - 1&
6890                      ReDim Preserve arr_varUserCreate(UC_ELEMS, lngE)
                          ' *****************************************
                          ' ** Array: arr_varUserCreate()
                          ' **
                          ' **   Element  Name           Constant
                          ' **   =======  =============  ==========
                          ' **      0     Username       UC_NAME
                          ' **      1     Password       UC_PW
                          ' **      2     s_GUID         UC_GUID
                          ' **      3     One Of Ours    UC_OURS
                          ' **
                          ' *****************************************
6900                      arr_varUserCreate(UC_OURS, lngE) = blnOneOfOurs
6910                      arr_varUserCreate(UC_NAME, lngE) = strUsername
6920                      arr_varUserCreate(UC_PW, lngE) = strUserPW
6930                      arr_varUserCreate(UC_GUID, lngE) = FilterGUIDString(StringFromGUID(arr_varTAUsr(TAU_GUID, lngX)))  ' ** Module Function: modCodeUtilities.
6940                    Else
                          ' ** Error: 3390  Account name already exists.
                          ' ** Error: 3304  You must enter a personal identifier (PID) consisting of at least 4 and no more than 20 characters and digits.
                          '####  WRITE THSE ERRORS!  ####
6950  On Error GoTo ERRH
'6950  On Error GoTo 0
6960                    End If
6970                  End If

                      ' ** Check against tblSecurity_User and tblSecurity_GroupUser.
6980                  If arr_varTAUsr(TAU_FND2, lngX) = False Then

                        ' ** Add user to tblSecurity_User.
6990                    Set rst1 = dbs.OpenRecordset("tblSecurity_User", dbOpenDynaset, dbConsistent)
7000                    With rst1
7010                      .AddNew
7020                      ![secusr_name] = arr_varTAUsr(TAU_NAME, lngX)
7030                      ![secusr_default] = False
7040                      ![secusr_user] = GetUserName  ' ** Module Function: modFileUtilities.
7050                      ![secusr_datecreated] = Now()
7060                      ![secusr_datemodified] = Now()
7070                      .Update
7080                      .Bookmark = .LastModified
7090                      lngUserID = ![secusr_id]
7100                      .Close
7110                    End With

                        ' ** Add user to tblSecurity_GroupUser.
7120                    Set qdf = dbs.QueryDefs("qrySecurity_Sync_05")
7130                    With qdf.Parameters
7140                      ![grpid] = lngUsersGrpID
7150                      ![usrid] = lngUserID
7160                      ![IsDef] = False
7170                    End With
7180                    qdf.Execute dbFailOnError
7190                    Set qdf = dbs.QueryDefs("qrySecurity_Sync_05")
7200                    With qdf.Parameters
7210                      ![grpid] = arr_varTAUsr(TAU_GRP2, lngX)
7220                      ![usrid] = lngUserID
7230                      ![IsDef] = False
7240                    End With
7250                    qdf.Execute dbFailOnError

7260                  End If

7270                Next

7280              End If  ' ** blnUsrMismatch.

7290              lngExtUsers = 0&
7300              ReDim arr_varExtUser(XUSR_ELEMS, 0)

                  ' ** Check for extraneous users in TrustSec.mdw.
7310              If blnExtraMDW = True Then
7320                For Each usr In DBEngine.Workspaces(0).Users
7330                  blnFound = False
7340                  With usr
7350                    For lngX = 0& To (lngTAUsrs - 1&)
7360                      If .Name = "Creator" Or .Name = "Engine" Or .Name = "Admin" Then
                            ' ** I don't know anything about these, so better not mess with them.
7370                        blnFound = True
7380                      Else
7390                        If arr_varTAUsr(TAU_NAME, lngX) = .Name Then
7400                          blnFound = True
7410                          Exit For
7420                        End If
7430                      End If
7440                    Next
7450                    If blnFound = False Then
                          ' ** Delete this user.
7460                      lngExtUsers = lngExtUsers + 1&
7470                      ReDim Preserve arr_varExtUser(XUSR_ELEMS, (lngExtUsers - 1&))
7480                      arr_varExtUser(XUSR_NAME, (lngExtUsers - 1&)) = .Name
7490                    End If
7500                  End With
7510                Next
7520                If lngExtUsers > 0& Then
7530                  For lngX = 0& To (lngExtUsers - 1&)
7540  On Error Resume Next
7550                    DBEngine.Workspaces(0).Users.Delete arr_varExtUser(XUSR_NAME, lngX)
                        '####  WRITE THESE ERRORS!  ####
7560                    Select Case ERR.Number
                        Case 3030  ' ** '|' is not a valid account name.
                          ' ** Since we're deleting it anyway, the error is moot; ignore it.
7570  On Error GoTo ERRH
'7570  On Error GoTo 0
7580                    Case 3032  ' ** Can't perform this operation.
                          ' ** Is this extra user the one that's logged in?
7590  On Error GoTo ERRH
'7590  On Error GoTo 0
7600                      If arr_varExtUser(XUSR_NAME, lngX) = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
7610                        If CurrentUser <> "superuser" And CurrentUser <> "Admin" Then  ' ** Internal Access Function: Trust Accountant login.
7620                          DoCmd.Hourglass False
7630                          blnRetVal = False
7640                          Beep
7650                          MsgBox "Security synchronization cannot be completed." & vbCrLf & _
                                "Please log in as either Admin or superuser.", vbInformation + vbOKOnly, "Security Sync Failed"
7660                          DoCmd.Hourglass True
7670                          Exit For
7680                        Else
                              ' ** Obviously, this user shouldn't be on a delete list!
7690                        End If
7700                      Else
                            ' ** The problem isn't that it's trying to delete
                            ' ** the logged in user, it's something else.
7710                        DoCmd.Hourglass False
7720                        blnRetVal = False
7730                        Beep
7740                        MsgBox "Security synchronization cannot be completed." & vbCrLf & _
                              "Please try again." & vbCrLf & vbCrLf & _
                              "If this message persists, please contact Delta Data, Inc.", vbInformation + vbOKOnly, "Security Sync Failed"
7750                        DoCmd.Hourglass True
7760                        Exit For
7770                      End If
7780                    Case 3033  ' ** You do not have the necessary permissions to use the 'MSysAccounts' object.
                          ' ** A user with less than 'Admins' permissions hit this code. Nothing we can do.
7790  On Error GoTo ERRH
'7790  On Error GoTo 0
7800                    Case 0
                          ' ** Everything's fine.
7810  On Error GoTo ERRH
'7790  On Error GoTo 0
7820                    Case Else
                          ' ** Some other error!
7830                      DoCmd.Hourglass False
7840                      blnRetVal = False
7850                      Beep
7860                      strErrMsg = vbNullString
7870                      strErrMsg = strErrMsg + "Error:" + vbTab + vbTab & CStr(ERR.Number) + vbCrLf
7880                      strErrMsg = strErrMsg + "Description:" + vbTab + ERR.description + vbCrLf + vbCrLf
7890                      strErrMsg = strErrMsg + "Module:" + vbTab + vbTab + THIS_NAME + vbCrLf
7900                      strErrMsg = strErrMsg + "Sub/Function:" + vbTab + THIS_PROC + vbCrLf
7910                      strErrMsg = strErrMsg + "Line:" + vbTab + vbTab + CStr(Erl)
7920                      MsgBox "Security synchronization cannot be completed." & vbCrLf & vbCrLf & _
                            strErrMsg, vbCritical + vbOKOnly, "Security Sync Failed"
7930  On Error GoTo ERRH
'7910  On Error GoTo 0
7940                      DoCmd.Hourglass True
7950                    End Select
7960                  Next
7970                  DBEngine.Workspaces(0).Users.Refresh
7980                End If
7990              End If

8000              If blnRetVal = True Then

                    ' ** Reload arr_varUsr().
8010                Set qdf = dbs.QueryDefs("qrySecurity_Sync_02")
8020                Set rst1 = qdf.OpenRecordset
8030                With rst1
8040                  .MoveLast
8050                  lngUsrs = .RecordCount
8060                  .MoveFirst
8070                  arr_varUsr = .GetRows(lngUsrs)
8080                  .Close
8090                End With

                    ' ** Now check for extraneous users in tblSecurity_User.
8100                For lngX = 0& To (lngUsrs - 1&)
8110                  If arr_varUsr(USR_DEF, lngX) = False Then
8120                    For lngY = 0& To (lngTAUsrs - 1&)
8130                      If arr_varTAUsr(TAU_NAME, lngY) = arr_varUsr(USR_NAME, lngX) Then
8140                        arr_varUsr(USR_FND1, lngX) = True
8150                        Exit For
8160                      End If
8170                    Next
8180                  End If
8190                Next
8200                For lngX = 0& To (lngUsrs - 1&)
8210                  If arr_varUsr(USR_DEF, lngX) = False And arr_varUsr(USR_FND1, lngX) = False Then
                        ' ** Delete user from tblSecurity_User, by specified [usrid].
8220                    Set qdf = dbs.QueryDefs("qrySecurity_Sync_07")
8230                    With qdf.Parameters
8240                      ![usrid] = arr_varUsr(USR_ID, lngX)
8250                    End With
8260                    qdf.Execute dbFailOnError
8270                  End If
8280                Next

8290              End If

8300            Else
                  ' ** If both are bolloxed, just tell the user.
8310              DoCmd.Hourglass False
8320              blnRetVal = False
8330              Beep
8340              MsgBox "The Security file is invalid!" & vbCrLf & vbCrLf & "Trust Accountant cannot continue." & vbCrLf & _
                    "Contact Delta Data, Inc., for assistance." & vbCrLf & "2", vbCritical + vbOKOnly, "Invalid Security File"
8350              DoCmd.Hourglass True
8360            End If

8370          End If  ' ** Wrong MDW.

8380        End If  ' ** There is a mismatch.

            ' ** Lastly, but not leastly, coordinate with _~xusr.
8390        If blnRetVal = True Then

8400          lngUsrs = 0&
8410          arr_varUsr = Empty

              ' ** Reload the arr_varUsr() array from tblSecurity_User.
8420          Set qdf = dbs.QueryDefs("qrySecurity_Sync_02")
8430          Set rst1 = qdf.OpenRecordset
8440          With rst1
8450            .MoveLast
8460            lngUsrs = .RecordCount
8470            .MoveFirst
8480            arr_varUsr = .GetRows(lngUsrs)
8490            .Close
8500          End With

              ' ****************************************************
              ' ** Array: arr_varUsr()
              ' **
              ' **   Field  Element  Name              Constant
              ' **   =====  =======  ================  ===========
              ' **     1       0     secusr_id         USR_ID
              ' **     2       1     secusr_name       USR_NAME
              ' **     3       2     secusr_default    USR_DEF
              ' **     4       3     grps              USR_GRPS
              ' **     5       4     grp1              USR_GRP1
              ' **     6       5     grp1_done         USR_GRP1D
              ' **     7       6     grp2              USR_GRP2
              ' **     8       7     grp2_done         USR_GRP2D
              ' **     9       8     Found1 (False)    USR_FND1  {tblSecurity_User not in TrustSec.mdw}
              ' **    10       9     Found2 (False)    USR_FND2  {tblSecurity_User not in Users}
              ' **    11      10     sec_password      USR_PW
              ' **
              ' ****************************************************

              ' ***************************************************
              ' ** Array: arr_varTAUsr()
              ' **
              ' **   Field  Element  Name             Constant
              ' **   =====  =======  ===============  ===========
              ' **     1       0     s_GUID           TAU_GUID
              ' **     2       1     Username         TAU_NAME
              ' **     3       2     Employee Name    TAU_EMP
              ' **     4       3     Default Group    TAU_DGRP
              ' **     5       4     secgrp_id1       TAU_GRP1
              ' **     6       5     Primary Group    TAU_PGRP
              ' **     7       6     secgrp_id2       TAU_GRP2
              ' **     8       7     Found1 (False)   TAU_FND1  {Trust Accountant User not in TrustSec.mdw}
              ' **     9       8     Found2 (False)   TAU_FND2  {Trust Accountant User not in tblSecurity_User}
              ' **    10       9     Found3 (False)   TAU_FND3  {Users not in _~xusr}
              ' **    11      10     sec_password     TAU_PW
              ' **    12      11     sec_oldname      TAU_OLDN
              ' **
              ' ***************************************************
              'Debug.Print "'s_GUID1: " & FilterGUIDString(StringFromGUID(arr_varTAUsr(TAU_GUID, 0)))  ' ** Module Function: modCodeUtilities.
              '  s_GUID1: {D856EA19-0B8B-11D2-8298-680D08C10000}
              'Debug.Print "'s_GUID2: " & arr_varTAUsr(TAU_GUID, 0)
              '  s_GUID2: {guid {D856EA19-0B8B-11D2-8298-680D08C10000}}

              ' ** Check _~xusr against the arr_varTAUsr() array.
8510          Set rst1 = dbs.OpenRecordset("_~xusr", dbOpenDynaset, dbConsistent)
8520          rst1.MoveLast
8530          lngRecs1 = rst1.RecordCount
8540          For lngX = 0& To (lngTAUsrs - 1&)
8550            rst1.MoveFirst
8560            For lngY = 1& To lngRecs1
8570              If FilterGUIDString(StringFromGUID(rst1![s_GUID])) = FilterGUIDString(StringFromGUID(arr_varTAUsr(TAU_GUID, lngX))) Then
                    ' ** Module Function: modCodeUtilities.
8580                arr_varTAUsr(TAU_FND3, lngX) = CBool(True)
8590                Exit For
8600              End If
8610              If lngY < lngRecs1 Then rst1.MoveNext
8620            Next
8630          Next

              ' ** Append any missing users to _~xusr.
8640          For lngX = 0& To (lngTAUsrs - 1&)
                ' ** Ignore the default users.
8650            blnFound = False
8660            For lngY = 0& To (lngUsrs - 1&)
8670              If arr_varUsr(USR_NAME, lngY) = arr_varTAUsr(TAU_NAME, lngX) And arr_varUsr(USR_DEF, lngY) = True Then
8680                blnFound = True
8690                Exit For
8700              End If
8710            Next
8720            If blnFound = False Then
                  ' ** Not a default user.
8730              If arr_varTAUsr(TAU_FND3, lngX) = False Then
                    ' ** Append User to live, by specified [usr], [ext], [org].
8740                Set qdf = dbs.QueryDefs("qryXUsr_04")
8750                With qdf.Parameters
8760                  ![usr] = arr_varTAUsr(TAU_NAME, lngX)
                      ' ** These new users should only happen once, the first time they log into a new installation.
8770                  If arr_varTAUsr(TAU_PW, lngX) <> vbNullString Then
8780                    ![ext] = EncodeString(CStr(arr_varTAUsr(TAU_PW, lngX)))  ' ** Module Function: modCodeUtilities.
8790                  Else
8800                    arr_varTAUsr(TAU_PW, lngX) = Security_PIN_To_PW(CStr(arr_varTAUsr(TAU_NAME, lngX)))  ' ** Function: Below.
8810                    ![ext] = EncodeString(CStr(arr_varTAUsr(TAU_PW, lngX)))  ' ** Module Function: modCodeUtilities.
8820                  End If
8830                  ![org] = EncodeString(Format(Date, "mm/dd/yyyy"))  ' ** Module Function: modCodeUtilities.
8840                End With
8850                qdf.Execute dbFailOnError
8860              End If
8870            End If
8880          Next

              ' ** Make sure arr_varUserCreate() has the current password.
8890          If lngUserCreates > 0& Then
8900            For lngX = 0& To (lngUserCreates - 1&)
8910              For lngY = 0& To (lngTAUsrs - 1&)
8920                If arr_varTAUsr(TAU_NAME, lngY) = arr_varUserCreate(UC_NAME, lngX) Then
8930                  If arr_varUserCreate(UC_PW, lngX) = vbNullString Then
8940                    If arr_varTAUsr(TAU_PW, lngY) = vbNullString Then
8950                      arr_varUserCreate(UC_PW, lngX) = Security_PIN_To_PW(CStr(arr_varTAUsr(TAU_NAME, lngY)))  ' ** Function: Below.
8960                    Else
8970                      arr_varUserCreate(UC_PW, lngX) = arr_varTAUsr(TAU_PW, lngY)
8980                    End If
8990                  Else
9000                    If arr_varTAUsr(TAU_PW, lngY) <> vbNullString Then
9010                      If arr_varUserCreate(UC_PW, lngX) <> arr_varTAUsr(TAU_PW, lngY) Then
9020                        Select Case Pass_Check(arr_varUserCreate(UC_PW, lngX), arr_varUserCreate(UC_NAME, lngX), False)  ' ** Module Function: modCodeUtilities.
                            Case True
                              ' ** The password in arr_varUserCreate() will take precedence.
9030                        Case False
9040                          Select Case Pass_Check(arr_varTAUsr(TAU_PW, lngY), arr_varTAUsr(TAU_NAME, lngY), False)  ' ** Module Function: modCodeUtilities.
                              Case True
9050                            arr_varUserCreate(UC_PW, lngX) = arr_varTAUsr(TAU_PW, lngY)
9060                          Case False
9070                            arr_varUserCreate(UC_PW, lngX) = Security_PIN_To_PW(CStr(arr_varUserCreate(UC_NAME, lngX)))  ' ** Function: Below.
9080                          End Select
9090                        End Select
                            ' ** Pass_Check(varInput As Variant, varUserName As Variant, Optional varShowMsg As Variant) As Boolean
9100                      End If
9110                    End If
9120                  End If
9130                  Exit For
9140                End If
9150              Next  ' ** lngY.
9160            Next  ' ** lngX.
9170          End If  ' ** lngUserCreates.

              ' **************************************
              ' ** Array: arr_varUserCreate()
              ' **
              ' **   Element  Name        Constant
              ' **   =======  ==========  ==========
              ' **      0     Username    UC_NAME
              ' **      1     Password    UC_PW
              ' **      2     s_GUID      UC_GUID
              ' **
              ' **************************************

              ' ** If users are in arr_varUserCreate(), we've got their current password.
9180          If lngUserCreates > 0& Then
9190            Set rst1 = dbs.OpenRecordset("_~xusr", dbOpenDynaset, dbConsistent)
9200            rst1.MoveLast
9210            lngRecs1 = rst1.RecordCount
9220            For lngX = 0& To (lngUserCreates - 1&)
9230              rst1.MoveFirst
9240              For lngY = 1& To lngRecs1
9250                If FilterGUIDString(StringFromGUID(rst1![s_GUID])) = arr_varUserCreate(UC_GUID, lngX) Then  ' ** Module Function: modCodeUtilities.
9260                  If rst1![xusr_extant] <> EncodeString(CStr(arr_varUserCreate(UC_PW, lngX))) Then  ' ** Module Function: modCodeUtilities.
9270                    rst1.Edit
9280                    rst1![xusr_antecedent] = rst1![xusr_extant]  ' ** May be our default 'NewXUsr' from above.
9290                    rst1![xusr_extant] = EncodeString(CStr(arr_varUserCreate(UC_PW, lngX)))  ' ** Module Function: modCodeUtilities.
9300                    rst1![xusr_origin] = EncodeString(Format(Date, "mm/dd/yyyy"))  ' ** Module Function: modCodeUtilities.
9310                    rst1![xusr_user] = GetUserName  ' ** Module Function: modFileUtilities.
9320                    rst1![xusr_datemodified] = Now()
9330                    rst1.Update
9340                  End If
9350                  Exit For
9360                End If
9370                If lngY < lngRecs1 Then rst1.MoveNext
9380              Next
9390            Next
9400            rst1.Close
9410          End If

              ' ** These 3 come with Access, and must always remain:
              ' **   Admin
              ' **   Creator
              ' **   Engine
              ' ** This is our development user, and must always remain:
              ' **   Superuser  - goes out in Users table.
              ' ** This is the demo user, and must always remain in the demo version:
              ' **   TADemo     - goes out in Users table.
              ' ** These are the demo users, and may be altered by users in the demo version:
              ' **   Bill       - goes out in Users table.
              ' **   Mary       - goes out in Users table.

              ' ** Backup tables for _~xusr.
              ' **   tblTemplate_Zeta1 : _~xusr for New and Upgrade versions.
              ' **   tblTemplate_Zeta2 : Users for New and Upgrade versions.
              ' **   tblTemplate_Zeta3 : _~xusr for Demo version.
              ' **   tblTemplate_Zeta4 : Users for Demo version.

9420        End If  ' ** blnRetVal.

9430      Case False
            ' ** Only count is invalid. Groups that are present may also be invalid.
9440        DoCmd.Hourglass False
9450        blnRetVal = False
9460        blnGrpMismatch = True
9470        Beep
9480        MsgBox "The Security file is invalid!" & vbCrLf & vbCrLf & "Trust Accountant cannot continue." & vbCrLf & _
              "Contact Delta Data, Inc., for assistance." & vbCrLf & "3", vbCritical + vbOKOnly, "Invalid Security File"
9490        DoCmd.Hourglass True
9500      End Select  ' ** blnFound.

9510    End If  ' ** blnGrpMismatch.

9520    If blnRetVal = True Then
          ' ** Make sure any name changes are reflected in Users and tblSecurity_User.
9530      For lngX = 0& To (lngTAUsrs - 1&)
9540        If arr_varTAUsr(TAU_OLDN, lngX) <> vbNullString Then
              ' ** Update Users, by specified [usr], [usrnew].
9550          Set qdf = dbs.QueryDefs("qrySecurity_Sync_10")
9560          With qdf.Parameters
9570            ![usr] = arr_varTAUsr(TAU_OLDN, lngX)
9580            ![usrnew] = arr_varTAUsr(TAU_NAME, lngX)
9590          End With
9600          qdf.Execute dbFailOnError
              ' ** Update tblSecurity_User, by specified [usr], [usrnew].
9610          Set qdf = dbs.QueryDefs("qrySecurity_Sync_11")
9620          With qdf.Parameters
9630            ![usr] = arr_varTAUsr(TAU_OLDN, lngX)
9640            ![usrnew] = arr_varTAUsr(TAU_NAME, lngX)
9650          End With
9660          qdf.Execute dbFailOnError
9670        End If
9680      Next

9690      dbs.Close

9700      If lngUserCreates > 0& Then
            ' **************************************
            ' ** Array: arr_varUserCreate()
            ' **
            ' **   Element  Name        Constant
            ' **   =======  ==========  ==========
            ' **      0     Username    UC_NAME
            ' **      1     Password    UC_PW
            ' **
            ' **************************************
9710        strErrMsg = vbNullString
9720        For lngX = 0& To (lngUserCreates - 1&)
9730          If arr_varUserCreate(UC_OURS, lngX) = False Then
9740            strErrMsg = strErrMsg & Left(arr_varUserCreate(UC_NAME, lngX) & Space(16), 16) & _
                  Left(arr_varUserCreate(UC_PW, lngX) & Space(15), 15) & vbCrLf
9750          End If
9760        Next
9770        If strErrMsg <> vbNullString Then
9780          DoCmd.Hourglass False
9790          DoCmd.OpenForm "frmUser_SecurityNotice", acNormal, , , , acDialog, strErrMsg
9800          DoCmd.Hourglass True
9810        End If
9820      End If

9830    End If  ' ** blnRetVal.

        ' ** It appears that, at a glance, the above code should already
        ' ** be doing what I've now put into Security_SyncChk2().
        ' ** So why doesn't it? The use (or lack thereof) of dbFailOnError?
9840    Security_SyncChk2  ' ** Function: Below.

EXITP:
9850    Set tdf = Nothing
9860    Set rst2 = Nothing
9870    Set rst1 = Nothing
9880    Set qdf = Nothing
9890    Set dbs = Nothing
9900    Set dbsLnk = Nothing
9910    Set wrkLnk = Nothing
9920    Set usr = Nothing
9930    Set grp = Nothing
9940    Security_SyncChk = blnRetVal
9950    Exit Function

ERRH:
9960    blnRetVal = False
9970    Select Case ERR.Number
        Case Else
9980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9990    End Select
10000   Resume EXITP

End Function

Public Function Security_SyncChk2() As Boolean

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_SyncChk2"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, qdf3 As DAO.QueryDef
        Dim rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset
        Dim lngDels As Long, arr_varDel() As Variant
        Dim lngXUsrID As Long, lngSecUsrID As Long
        Dim blnAddAll As Boolean, blnFound As Boolean ', blnAdd As Boolean
        Dim varTmp00 As Variant, lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varDel().
        Const D_ELEMS As Integer = 5  ' ** Array's first-element UBound().
        Const D_ID   As Integer = 0
        Const D_NAM  As Integer = 1
        Const D_GUID As Integer = 2
        Const D_DEL  As Integer = 3
        Const D_FND  As Integer = 4
        Const D_ADD  As Integer = 5

10110   blnRetVal = True

10120   Set dbs = CurrentDb
10130   With dbs

10140     lngTmp01 = 0&: lngTmp02 = 0&: lngTmp03 = 0&

          ' ** Users, just non-default users.
10150     Set qdf1 = .QueryDefs("qrySecurity_User_20_01")
10160     Set rst1 = qdf1.OpenRecordset
10170     With rst1
10180       If .BOF = True And .EOF = True Then
              ' ** No non-default users.
10190       Else
10200         .MoveLast
10210         lngTmp01 = .RecordCount
10220       End If
10230       .Close
10240     End With
10250     Set rst1 = Nothing
10260     Set qdf1 = Nothing

          ' ** _~xusr, just non-default users.
10270     Set qdf2 = .QueryDefs("qrySecurity_User_20_02")
10280     Set rst2 = qdf2.OpenRecordset
10290     With rst2
10300       If .BOF = True And .EOF = True Then
              ' ** No non-default users.
10310       Else
10320         .MoveLast
10330         lngTmp02 = .RecordCount
10340       End If
10350       .Close
10360     End With
10370     Set rst2 = Nothing
10380     Set qdf2 = Nothing

          ' ** tblSecurity_User, just non-default users.
10390     Set qdf3 = .QueryDefs("qrySecurity_User_20_03")
10400     Set rst3 = qdf3.OpenRecordset
10410     With rst3
10420       If .BOF = True And .EOF = True Then
              ' ** No non-default users.
10430       Else
10440         .MoveLast
10450         lngTmp03 = .RecordCount
10460       End If
10470       .Close
10480     End With
10490     Set rst3 = Nothing
10500     Set qdf3 = Nothing

10510     If lngTmp01 > 0& Or lngTmp02 > 0& Or lngTmp03 > 0& Then

            ' ** Users, just non-default users.
10520       Set qdf1 = .QueryDefs("qrySecurity_User_20_01")
10530       Set rst1 = qdf1.OpenRecordset
            ' ** _~xusr, just non-default users.
10540       Set qdf2 = .QueryDefs("qrySecurity_User_20_02")
10550       Set rst2 = qdf2.OpenRecordset

10560       If lngTmp01 <> lngTmp02 Then
              ' ** Users <> _~xusr.

10570         If lngTmp01 = 0& Then
                ' ** Users will be the primary table, so delete all non-default users in _~xusr.

10580           lngDels = 0&
10590           ReDim arr_varDel(D_ELEMS, 0)

                ' ** Collect the ID's.
10600           With rst2
10610             .MoveFirst
10620             For lngX = 1& To lngTmp02
10630               lngDels = lngDels + 1&
10640               lngE = lngDels - 1&
10650               ReDim Preserve arr_varDel(D_ELEMS, lngE)
10660               arr_varDel(D_ID, lngE) = ![xusr_id]
10670               arr_varDel(D_NAM, lngE) = Null
10680               arr_varDel(D_GUID, lngE) = ![s_GUID]
10690               arr_varDel(D_DEL, lngE) = CBool(True)
10700               arr_varDel(D_FND, lngE) = CBool(False)
10710               arr_varDel(D_ADD, lngE) = CBool(False)
10720               If lngX < lngTmp02 Then .MoveNext
10730             Next
10740           End With  ' ** rst2.
                ' ** Leave it open.

                ' ** Now delete them.
10750           For lngX = 0& To (lngDels - 1&)
10760             If arr_varDel(D_DEL, lngX) = True Then
                    ' ** Delete _~xusr, by specified [xusrid].
10770               Set qdf3 = .QueryDefs("qrySecurity_User_21_01")
10780               With qdf3.Parameters
10790                 ![xusrid] = arr_varDel(D_ID, lngX)
10800               End With
10810               qdf3.Execute
10820             End If
10830           Next
10840           Set qdf3 = Nothing

10850         Else

10860           lngDels = 0&
10870           ReDim arr_varDel(D_ELEMS, 0)

10880           blnAddAll = False
10890           With rst1
10900             .MoveFirst
10910             If lngTmp02 = 0& Then
10920               blnAddAll = True
10930             Else
10940               rst2.MoveFirst
10950               For lngX = 1& To lngTmp01
10960                 blnFound = False
10970                 For lngY = 1& To lngTmp02
10980                   If rst2![s_GUID] = ![s_GUID] Then
10990                     blnFound = True
11000                     Exit For
11010                   End If
11020                   If lngY < lngTmp02 Then rst2.MoveNext
11030                 Next  ' ** lngY.
11040                 lngDels = lngDels + 1&
11050                 lngE = lngDels - 1&
11060                 ReDim Preserve arr_varDel(D_ELEMS, lngE)
11070                 arr_varDel(D_ID, lngE) = Null
11080                 arr_varDel(D_NAM, lngE) = ![Username]
11090                 arr_varDel(D_GUID, lngE) = ![s_GUID]
11100                 arr_varDel(D_DEL, lngE) = CBool(False)
11110                 Select Case blnFound
                      Case True
11120                   arr_varDel(D_FND, lngE) = CBool(True)
11130                   arr_varDel(D_ADD, lngE) = CBool(False)
11140                 Case False
11150                   arr_varDel(D_FND, lngE) = CBool(False)
11160                   arr_varDel(D_ADD, lngE) = CBool(True)
11170                 End Select
11180                 If lngX < lngTmp01 Then .MoveNext
11190               Next  ' ** lngX.
11200             End If  ' ** lngTmp02.
11210             Select Case blnAddAll
                  Case True
11220               .MoveFirst
11230               For lngX = 1& To lngTmp01
11240                 varTmp00 = DLookup("[secusr_id]", "tblSecurity_User", "[secusr_name] = '" & ![Username] & "'")
11250                 Select Case IsNull(varTmp00)
                      Case True
                        ' ** It's not in tblSecurity_User either.
                        ' ** Append qrySecurity_User_22_01 (Users, as new _~xusr record, by specified [usr]) to _~xusr.
11260                   Set qdf3 = dbs.QueryDefs("qrySecurity_User_22_02")
11270                   With qdf3.Parameters
11280                     ![usr] = ![Username]
11290                   End With
11300                   qdf3.Execute dbFailOnError
11310                 Case False
11320                   lngXUsrID = varTmp00
                        ' ** Append qrySecurity_User_22_03 (Users, as new _~xusr record, by specified [xusrid], [usr])to _~xusr.
11330                   Set qdf3 = dbs.QueryDefs("qrySecurity_User_22_04")
11340                   With qdf3.Parameters
11350                     ![xusrid] = lngXUsrID
11360                     ![usr] = ![Username]
11370                   End With
11380                   qdf3.Execute dbFailOnError
11390                 End Select
11400                 If lngX < lngTmp01 Then .MoveNext
11410               Next  ' ** lngX.
11420               Set qdf3 = Nothing
11430             Case False
                    ' ** Look for extraneous extras in _~xusr.
11440               With rst2
11450                 .MoveFirst
11460                 For lngX = 0& To (lngTmp02 - 1&)
11470                   blnFound = False
11480                   For lngY = 0& To (lngDels - 1&)
11490                     If arr_varDel(D_GUID, lngY) = ![s_GUID] Then
11500                       blnFound = True
11510                       Exit For
11520                     End If
11530                   Next  ' ** lngY.
11540                   If blnFound = False Then
                          ' ** Dead extra.
11550                     lngDels = lngDels + 1&
11560                     lngE = lngDels - 1&
11570                     ReDim Preserve arr_varDel(D_ELEMS, lngE)
11580                     arr_varDel(D_ID, lngE) = ![xusr_id]
11590                     arr_varDel(D_NAM, lngE) = Null
11600                     arr_varDel(D_GUID, lngE) = ![s_GUID]
11610                     arr_varDel(D_DEL, lngE) = CBool(True)
11620                     arr_varDel(D_FND, lngE) = CBool(False)
11630                     arr_varDel(D_ADD, lngE) = CBool(False)
11640                   End If
11650                   If lngX < lngTmp02 Then .MoveNext
11660                 Next  ' ** lngX.
11670               End With  ' ** rst2.
11680               If lngDels > 0& Then
11690                 For lngX = 0& To (lngDels - 1&)
11700                   If arr_varDel(D_ADD, lngX) = True Then
11710                     varTmp00 = DLookup("[secusr_id]", "tblSecurity_User", "[secusr_name] = '" & arr_varDel(D_NAM, lngX) & "'")
11720                     Select Case IsNull(varTmp00)
                          Case True
                            ' ** It's not in tblSecurity_User either.
                            ' ** Append qrySecurity_User_22_01 (Users, as new _~xusr record, by specified [usr]) to _~xusr.
11730                       Set qdf3 = dbs.QueryDefs("qrySecurity_User_22_02")
11740                       With qdf3.Parameters
11750                         ![usr] = arr_varDel(D_NAM, lngX)
11760                       End With
11770                       qdf3.Execute dbFailOnError
11780                     Case False
11790                       lngXUsrID = varTmp00
                            ' ** Append qrySecurity_User_22_03 (Users, as new _~xusr record, by specified [xusrid], [usr])to _~xusr.
11800                       Set qdf3 = dbs.QueryDefs("qrySecurity_User_22_04")
11810                       With qdf3.Parameters
11820                         ![xusrid] = lngXUsrID
11830                         ![usr] = arr_varDel(D_NAM, lngX)
11840                       End With
11850                       qdf3.Execute dbFailOnError
11860                     End Select
11870                   ElseIf arr_varDel(D_DEL, lngX) = True Then
                          ' ** Delete _~xusr, by specified [xusrid].
11880                     Set qdf3 = dbs.QueryDefs("qrySecurity_User_21_01")
11890                     With qdf3.Parameters
11900                       ![xusrid] = arr_varDel(D_ID, lngX)
11910                     End With
11920                     qdf3.Execute dbFailOnError
11930                   End If
11940                 Next  ' ** lngX.
11950                 Set qdf3 = Nothing
11960               End If  ' ** lngDels.
11970             End Select  ' ** blnAddAll.
11980           End With  ' ** rst1.

11990         End If  ' ** lngTmp01.

12000       Else
              ' ** They're equal, so just compare them.

12010         lngDels = 0&
12020         ReDim arr_varDel(D_ELEMS, 0)

12030         If lngTmp01 > 0& Then
12040           With rst2
12050             .MoveFirst
12060             For lngX = 1& To lngTmp02
12070               rst1.MoveFirst
12080               blnFound = False
12090               For lngY = 1& To lngTmp01
12100                 If rst1![s_GUID] = ![s_GUID] Then
12110                   blnFound = True
12120                   Exit For
12130                 End If
12140                 If lngY < lngTmp01 Then rst1.MoveNext
12150               Next  ' ** lngY.
12160               If blnFound = False Then
                      ' ** Oh-oh! Same count, but not same users!
12170                 lngDels = lngDels + 1&
12180                 lngE = lngDels - 1&
12190                 ReDim Preserve arr_varDel(D_ELEMS, lngE)
12200                 arr_varDel(D_ID, lngE) = ![xusr_id]
12210                 arr_varDel(D_NAM, lngE) = Null
12220                 arr_varDel(D_GUID, lngE) = ![s_GUID]
12230                 arr_varDel(D_DEL, lngE) = CBool(True)
12240                 arr_varDel(D_FND, lngE) = CBool(False)
12250                 arr_varDel(D_ADD, lngE) = CBool(False)
12260               End If  ' ** blnFound.
12270               If lngX < lngTmp02 Then .MoveNext
12280             Next  ' ** lngX.
12290             If lngDels > 0& Then
                    ' ** And that means there's a User not in this table!
12300               With rst1
12310                 .MoveFirst
12320                 For lngX = 1& To lngTmp01
12330                   blnFound = False
12340                   rst2.MoveFirst
12350                   For lngY = 1& To lngTmp02
12360                     If rst2![s_GUID] = ![s_GUID] Then
12370                       blnFound = True
12380                       Exit For
12390                     End If
12400                     If lngY < lngTmp02 Then rst2.MoveNext
12410                   Next  ' ** lngY.
12420                   If blnFound = False Then
12430                     lngDels = lngDels + 1&
12440                     lngE = lngDels - 1&
12450                     ReDim Preserve arr_varDel(D_ELEMS, lngE)
12460                     arr_varDel(D_ID, lngE) = Null
12470                     arr_varDel(D_NAM, lngE) = ![Username]
12480                     arr_varDel(D_GUID, lngE) = ![s_GUID]
12490                     arr_varDel(D_DEL, lngE) = CBool(False)
12500                     arr_varDel(D_FND, lngE) = CBool(False)
12510                     arr_varDel(D_ADD, lngE) = CBool(True)
12520                   End If
12530                   If lngX < lngTmp01 Then .MoveNext
12540                 Next  ' ** lngX.
12550               End With  ' ** rst1.
12560             End If  ' ** lngDels.
12570           End With  ' ** rst2.
12580           If lngDels > 0& Then
12590             For lngX = 0& To (lngDels - 1&)
12600               If arr_varDel(D_DEL, lngX) = True Then
                      ' ** Delete _~xusr, by specified [xusrid].
12610                 Set qdf3 = dbs.QueryDefs("qrySecurity_User_21_01")
12620                 With qdf3.Parameters
12630                   ![xusrid] = arr_varDel(D_ID, lngX)
12640                 End With
12650                 qdf3.Execute dbFailOnError
12660               ElseIf arr_varDel(D_ADD, lngX) = True Then
12670                 varTmp00 = DLookup("[secusr_id]", "tblSecurity_User", "[secusr_name] = '" & arr_varDel(D_NAM, lngX) & "'")
12680                 Select Case IsNull(varTmp00)
                      Case True
                        ' ** It's not in tblSecurity_User either.
                        ' ** Append qrySecurity_User_22_01 (Users, as new _~xusr record, by specified [usr]) to _~xusr.
12690                   Set qdf3 = dbs.QueryDefs("qrySecurity_User_22_02")
12700                   With qdf3.Parameters
12710                     ![usr] = arr_varDel(D_NAM, lngX)
12720                   End With
12730                   qdf3.Execute dbFailOnError
12740                 Case False
12750                   lngXUsrID = varTmp00
                        ' ** Append qrySecurity_User_22_03 (Users, as new _~xusr record, by specified [xusrid], [usr])to _~xusr.
12760                   Set qdf3 = dbs.QueryDefs("qrySecurity_User_22_04")
12770                   With qdf3.Parameters
12780                     ![xusrid] = lngXUsrID
12790                     ![usr] = arr_varDel(D_NAM, lngX)
12800                   End With
12810                   qdf3.Execute dbFailOnError
12820                 End Select
12830               End If
12840             Next  ' ** lngX.
12850             Set qdf3 = Nothing
12860           End If  ' ** lngDels.
12870         End If  ' ** lngTmp01.

12880       End If  ' ** lngTmp01, lngTmp02.

12890       rst1.Close
12900       rst2.Close
12910       Set rst1 = Nothing
12920       Set rst2 = Nothing
12930       Set qdf1 = Nothing
12940       Set qdf2 = Nothing

            ' ** Users, just non-default users.
12950       Set qdf1 = .QueryDefs("qrySecurity_User_20_01")
12960       Set rst1 = qdf1.OpenRecordset
            ' ** tblSecurity_User, just non-default users.
12970       Set qdf2 = .QueryDefs("qrySecurity_User_20_03")
12980       Set rst2 = qdf2.OpenRecordset

12990       If lngTmp01 <> lngTmp03 Then
              ' ** Users <> tblSecurity_User.

13000         If lngTmp01 = 0& Then
                ' ** Users will be the primary table, so delete all non-default users in tblSecurity_User.

13010           lngDels = 0&
13020           ReDim arr_varDel(D_ELEMS, 0)

                ' ** Collect the ID's.
13030           With rst2
13040             .MoveFirst
13050             For lngX = 1& To lngTmp03
13060               lngDels = lngDels + 1&
13070               lngE = lngDels - 1&
13080               ReDim Preserve arr_varDel(D_ELEMS, lngE)
13090               arr_varDel(D_ID, lngE) = ![secusr_id]
13100               arr_varDel(D_NAM, lngE) = Null
13110               arr_varDel(D_GUID, lngE) = Null
13120               arr_varDel(D_DEL, lngE) = CBool(True)
13130               arr_varDel(D_FND, lngE) = CBool(False)
13140               arr_varDel(D_ADD, lngE) = CBool(False)
13150               If lngX < lngTmp03 Then .MoveNext
13160             Next
13170           End With  ' ** rst2.
                ' ** Leave it open.

                ' ** Now delete them.
13180           For lngX = 0& To (lngDels - 1&)
13190             If arr_varDel(D_DEL, lngX) = True Then
                    ' ** Delete tblSecurity_User, by specified [secusrid].
13200               Set qdf3 = .QueryDefs("qrySecurity_User_21_02")
13210               With qdf3.Parameters
13220                 ![secusrid] = arr_varDel(D_ID, lngX)
13230               End With
13240               qdf3.Execute
13250             End If
13260           Next  ' ** lngX.
13270           Set qdf3 = Nothing

13280         Else

13290           lngDels = 0&
13300           ReDim arr_varDel(D_ELEMS, 0)

13310           blnAddAll = False
13320           With rst1
13330             .MoveFirst
13340             If lngTmp03 = 0& Then
13350               blnAddAll = True
13360             Else
13370               rst2.MoveFirst
13380               For lngX = 1& To lngTmp01
13390                 blnFound = False
13400                 For lngY = 1& To lngTmp03
13410                   If rst2![secusr_name] = ![Username] Then
13420                     blnFound = True
13430                     Exit For
13440                   End If
13450                   If lngY < lngTmp03 Then rst2.MoveNext
13460                 Next  ' ** lngY.
13470                 If blnFound = False Then
13480                   lngDels = lngDels + 1&
13490                   lngE = lngDels - 1&
13500                   ReDim Preserve arr_varDel(D_ELEMS, lngE)
13510                   arr_varDel(D_ID, lngE) = Null
13520                   arr_varDel(D_NAM, lngE) = ![Username]
13530                   arr_varDel(D_GUID, lngE) = ![s_GUID]
13540                   arr_varDel(D_DEL, lngE) = CBool(False)
13550                   arr_varDel(D_FND, lngE) = CBool(False)
13560                   arr_varDel(D_ADD, lngE) = CBool(True)
13570                 End If
13580                 If lngX < lngTmp01 Then .MoveNext
13590               Next  ' ** lngX.
13600             End If  ' ** lngTmp03.
13610             Select Case blnAddAll
                  Case True
13620               .MoveFirst
13630               For lngX = 1& To lngTmp01
13640                 varTmp00 = DLookup("[xusr_id]", "_~xusr", "[s_GUID] = '" & ![s_GUID] & "'")
13650                 Select Case IsNull(varTmp00)
                      Case True
                        ' ** It's not in _~xusr either.
                        ' ** Append qrySecurity_User_23_01(Users, as new tblSecurity_User
                        ' ** record, by specified [usr]) to tblSecurity_User.
13660                   Set qdf3 = dbs.QueryDefs("qrySecurity_User_23_02")
13670                   With qdf3.Parameters
13680                     ![usr] = ![Username]
13690                   End With
13700                   qdf3.Execute dbFailOnError
13710                 Case False
13720                   lngSecUsrID = varTmp00
                        ' ** Append qrySecurity_User_23_03 (Users, as new tblSecurity_User
                        ' ** record, by specified [secusrid], [usr]) to tblSecurity_User.
13730                   Set qdf3 = dbs.QueryDefs("qrySecurity_User_23_04")
13740                   With qdf3.Parameters
13750                     ![secusrid] = lngSecUsrID
13760                     ![usr] = ![Username]
13770                   End With
13780                   qdf3.Execute dbFailOnError
13790                 End Select
13800                 If lngX < lngTmp01 Then .MoveNext
13810               Next  ' ** lngX.
13820               Set qdf3 = Nothing
13830             Case False
                    ' ** Look for extraneous extras in tblSecurity_User.
13840               With rst2
13850                 .MoveFirst
13860                 For lngX = 0& To (lngTmp03 - 1&)
13870                   blnFound = False
13880                   For lngY = 0& To (lngDels - 1&)
13890                     If arr_varDel(D_NAM, lngY) = ![secusr_name] Then
13900                       blnFound = True
13910                       Exit For
13920                     End If
13930                   Next  ' ** lngY.
13940                   If blnFound = False Then
                          ' ** Dead extra.
13950                     lngDels = lngDels + 1&
13960                     lngE = lngDels - 1&
13970                     ReDim Preserve arr_varDel(D_ELEMS, lngE)
13980                     arr_varDel(D_ID, lngE) = ![secusr_id]
13990                     arr_varDel(D_NAM, lngE) = ![secusr_name]
14000                     arr_varDel(D_GUID, lngE) = Null
14010                     arr_varDel(D_DEL, lngE) = CBool(True)
14020                     arr_varDel(D_FND, lngE) = CBool(False)
14030                     arr_varDel(D_ADD, lngE) = CBool(False)
14040                   End If
14050                   If lngX < lngTmp03 Then .MoveNext
14060                 Next  ' ** lngX.
14070               End With  ' ** rst2.
14080               If lngDels > 0& Then
14090                 For lngX = 0& To (lngDels - 1&)
14100                   If arr_varDel(D_ADD, lngX) = True Then
14110                     varTmp00 = DLookup("[xusr_id]", "_~xusr", "[s_GUID] = '" & arr_varDel(D_GUID, lngX) & "'")
14120                     Select Case IsNull(varTmp00)
                          Case True
                            ' ** It's not in _~xusr either.
                            ' ** Append qrySecurity_User_23_01(Users, as new tblSecurity_User
                            ' ** record, by specified [usr]) to tblSecurity_User.
14130                       Set qdf3 = dbs.QueryDefs("qrySecurity_User_23_02")
14140                       With qdf3.Parameters
14150                         ![usr] = ![Username]
14160                       End With
14170                       qdf3.Execute dbFailOnError
14180                     Case False
14190                       lngSecUsrID = varTmp00
                            ' ** Append qrySecurity_User_23_03 (Users, as new tblSecurity_User
                            ' ** record, by specified [secusrid], [usr]) to tblSecurity_User.
14200                       Set qdf3 = dbs.QueryDefs("qrySecurity_User_23_04")
14210                       With qdf3.Parameters
14220                         ![secusrid] = lngSecUsrID
14230                         ![usr] = ![Username]
14240                       End With
14250                       qdf3.Execute dbFailOnError
14260                     Case False
14270                       lngSecUsrID = varTmp00
                            ' ** Append qrySecurity_User_23_03 (Users, as new tblSecurity_User
                            ' ** record, by specified [secusrid], [usr]) to tblSecurity_User.
14280                       Set qdf3 = dbs.QueryDefs("qrySecurity_User_23_04")
14290                       With qdf3.Parameters
14300                         ![secusrid] = lngSecUsrID
14310                         ![usr] = ![Username]
14320                       End With
14330                       qdf3.Execute dbFailOnError
14340                     End Select
14350                   ElseIf arr_varDel(D_DEL, lngX) = True Then
                          ' ** Delete tblSecurity_User, by specified [secusrid].
14360                     Set qdf3 = dbs.QueryDefs("qrySecurity_User_21_02")
14370                     With qdf3.Parameters
14380                       ![secusrid] = arr_varDel(D_ID, lngX)
14390                     End With
14400                     qdf3.Execute dbFailOnError
14410                   End If
14420                   If lngX < lngTmp01 Then .MoveNext
14430                 Next  ' ** lngX.
14440                 Set qdf3 = Nothing
14450               End If  ' ** lngDels.
14460             End Select  ' ** blnAddAll.
14470           End With  ' ** rst1.

14480         End If  ' ** lngTmp01.

14490       Else
              ' ** They're equal, so just compare them.

14500         lngDels = 0&
14510         ReDim arr_varDel(D_ELEMS, 0)

14520         If lngTmp01 > 0& Then

14530           With rst2
14540             .MoveFirst
14550             For lngX = 1& To lngTmp03
14560               rst1.MoveFirst
14570               blnFound = False
14580               For lngY = 1& To lngTmp01
14590                 If rst1![Username] = ![secusr_name] Then
14600                   blnFound = True
14610                   Exit For
14620                 End If
14630                 If lngY < lngTmp01 Then rst1.MoveNext
14640               Next  ' ** lngY.
14650               If blnFound = False Then
                      ' ** Oh-oh! Same count, but not same users!
14660                 lngDels = lngDels + 1&
14670                 lngE = lngDels - 1&
14680                 ReDim Preserve arr_varDel(D_ELEMS, lngE)
14690                 arr_varDel(D_ID, lngE) = ![secusr_id]
14700                 arr_varDel(D_NAM, lngE) = ![secusr_name]
14710                 arr_varDel(D_GUID, lngE) = Null
14720                 arr_varDel(D_DEL, lngE) = CBool(True)
14730                 arr_varDel(D_FND, lngE) = CBool(False)
14740                 arr_varDel(D_ADD, lngE) = CBool(False)
14750               End If  ' ** blnFound.
14760               If lngX < lngTmp03 Then .MoveNext
14770             Next  ' ** lngX.
14780             If lngDels > 0& Then
                    ' ** And that means there's a User not in this table!
14790               With rst1
14800                 .MoveFirst
14810                 For lngX = 1& To lngTmp01
14820                   blnFound = False
14830                   rst2.MoveFirst
14840                   For lngY = 1& To lngTmp03
14850                     If rst2![secusr_name] = ![Username] Then
14860                       blnFound = True
14870                       Exit For
14880                     End If
14890                     If lngY < lngTmp03 Then rst2.MoveNext
14900                   Next  ' ** lngY.
14910                   If blnFound = False Then
14920                     lngDels = lngDels + 1&
14930                     lngE = lngDels - 1&
14940                     ReDim Preserve arr_varDel(D_ELEMS, lngE)
14950                     arr_varDel(D_ID, lngE) = Null
14960                     arr_varDel(D_NAM, lngE) = ![Username]
14970                     arr_varDel(D_GUID, lngE) = ![s_GUID]
14980                     arr_varDel(D_DEL, lngE) = CBool(False)
14990                     arr_varDel(D_FND, lngE) = CBool(False)
15000                     arr_varDel(D_ADD, lngE) = CBool(True)
15010                   End If
15020                   If lngX < lngTmp01 Then .MoveNext
15030                 Next  ' ** lngX.
15040               End With  ' ** rst1.
15050             End If  ' ** lngDels.
15060           End With  ' ** rst2.
15070           If lngDels > 0& Then
15080             For lngX = 0& To (lngDels - 1&)
15090               If arr_varDel(D_DEL, lngX) = True Then
                      ' ** Delete tblSecurity_User, by specified [secusrid].
15100                 Set qdf3 = dbs.QueryDefs("qrySecurity_User_21_02")
15110                 With qdf3.Parameters
15120                   ![secusrid] = arr_varDel(D_ID, lngX)
15130                 End With
15140                 qdf3.Execute dbFailOnError
15150               ElseIf arr_varDel(D_ADD, lngX) = True Then
15160                 varTmp00 = DLookup("[xusr_id]", "_~xusr", "[s_GUID] = '" & arr_varDel(D_GUID, lngX) & "'")
15170                 Select Case IsNull(varTmp00)
                      Case True
                        ' ** It's not in _~xusr either.
                        ' ** Append qrySecurity_User_23_01 (Users, as new tblSecurity_User
                        ' ** record, by specified [usr]) to tblSecurity_User.
15180                   Set qdf3 = dbs.QueryDefs("qrySecurity_User_23_02")
15190                   With qdf3.Parameters
15200                     ![usr] = arr_varDel(D_NAM, lngX)
15210                   End With
15220                   qdf3.Execute dbFailOnError
15230                 Case False
15240                   lngSecUsrID = varTmp00
                        ' ** Append qrySecurity_User_23_03 (Users, as new tblSecurity_User
                        ' ** record, by specified [secusrid], [usr]) to tblSecurity_User.
15250                   Set qdf3 = dbs.QueryDefs("qrySecurity_User_23_04")
15260                   With qdf3.Parameters
15270                     ![secusrid] = lngSecUsrID
15280                     ![usr] = arr_varDel(D_NAM, lngX)
15290                   End With
15300                   qdf3.Execute dbFailOnError
15310                 End Select
15320               End If
15330             Next  ' ** lngX.
15340             Set qdf3 = Nothing
15350           End If  ' ** lngDels.

15360         End If  ' ** lngTmp01.

15370       End If  ' ** lngTmp01, lngTmp03.

15380       rst1.Close
15390       rst2.Close
15400       Set rst1 = Nothing
15410       Set rst2 = Nothing
15420       Set qdf1 = Nothing
15430       Set qdf2 = Nothing

            ' ** OK, now check tblSecurity_GroupUser.

            ' ** tblSecurity_GroupUser, not in tblSecurity_User.
15440       Set qdf3 = .QueryDefs("qrySecurity_User_24_01")
15450       Set rst3 = qdf3.OpenRecordset
15460       If rst3.BOF = True And rst3.EOF = True Then
              ' ** All's well.
15470         rst3.Close
15480         Set rst3 = Nothing
15490         Set qdf3 = Nothing
15500       Else
15510         rst3.Close
15520         Set rst3 = Nothing
15530         Set qdf3 = Nothing
              ' ** Delete qrySecurity_User_24_01 (tblSecurity_GroupUser, not in tblSecurity_User).
15540         Set qdf3 = .QueryDefs("qrySecurity_User_24_02")
15550         qdf3.Execute dbFailOnError
15560         Set qdf3 = Nothing
15570       End If

15580       If lngTmp01 > 0& Then

              ' ** Group: Users.
              ' ** qrySecurity_User_25_01 (tblSecurity_User, and tblSecurity_Group,
              ' ** as new tblSecurity_GroupUser records, for 'Users'; Cartesian),
              ' ** not in tblSecurity_GroupUser.
15590         Set qdf3 = .QueryDefs("qrySecurity_User_25_02")
15600         Set rst3 = qdf3.OpenRecordset
15610         If rst3.BOF = True And rst3.EOF = True Then
                ' ** All's well.
15620           rst3.Close
15630           Set rst3 = Nothing
15640           Set qdf3 = Nothing
15650         Else
15660           rst3.Close
15670           Set rst3 = Nothing
15680           Set qdf3 = Nothing
                ' ** Append qrySecurity_User_25_02 (qrySecurity_User_25_01 (tblSecurity_User,
                ' ** and tblSecurity_Group, as new tblSecurity_GroupUser records, for 'Users';
                ' ** Cartesian), not in tblSecurity_GroupUser) to tblSecurity_GroupUser.
15690           Set qdf3 = .QueryDefs("qrySecurity_User_25_03")
15700           qdf3.Execute dbFailOnError
15710           Set qdf3 = Nothing
15720         End If

              ' ** Group: Admins.
              ' ** qrySecurity_User_26_01 (tblSecurity_User, linked to Users, tblSecurity_Group,
              ' ** as new tblSecurity_GroupUser records, for 'Admins'), not in tblSecurity_GroupUser.
15730         Set qdf3 = .QueryDefs("qrySecurity_User_26_02")
15740         Set rst3 = qdf3.OpenRecordset
15750         If rst3.BOF = True And rst3.EOF = True Then
                ' ** All's well.
15760           rst3.Close
15770           Set rst3 = Nothing
15780           Set qdf3 = Nothing
15790         Else
15800           rst3.Close
15810           Set rst3 = Nothing
15820           Set qdf3 = Nothing
                ' ** Append qrySecurity_User_26_02 (qrySecurity_User_26_01 (tblSecurity_User,
                ' ** linked to Users, tblSecurity_Group, as new tblSecurity_GroupUser records,
                ' ** for 'Admins'), not in tblSecurity_GroupUser) to tblSecurity_GroupUser.
15830           Set qdf3 = .QueryDefs("qrySecurity_User_26_03")
15840           qdf3.Execute dbFailOnError
15850           Set qdf3 = Nothing
15860         End If

              ' ** Group: DataEntry.
              ' ** qrySecurity_User_27_01 (tblSecurity_User, linked to Users, tblSecurity_Group,
              ' ** as new tblSecurity_GroupUser records, for 'DataEntry'), not in tblSecurity_GroupUser.
15870         Set qdf3 = .QueryDefs("qrySecurity_User_27_02")
15880         Set rst3 = qdf3.OpenRecordset
15890         If rst3.BOF = True And rst3.EOF = True Then
                ' ** All's well.
15900           rst3.Close
15910           Set rst3 = Nothing
15920           Set qdf3 = Nothing
15930         Else
15940           rst3.Close
15950           Set rst3 = Nothing
15960           Set qdf3 = Nothing
                ' ** Append qrySecurity_User_27_02 (qrySecurity_User_27_01 (tblSecurity_User,
                ' ** linked to Users, tblSecurity_Group, as new tblSecurity_GroupUser records,
                ' ** for 'DataEntry'), not in tblSecurity_GroupUser) to tblSecurity_GroupUser.
15970           Set qdf3 = .QueryDefs("qrySecurity_User_27_03")
15980           qdf3.Execute dbFailOnError
15990           Set qdf3 = Nothing
16000         End If

              ' ** Group: ViewOnly.
              ' ** qrySecurity_User_28_01 (tblSecurity_User, linked to Users, tblSecurity_Group,
              ' ** as new tblSecurity_GroupUser records, for 'ViewOnly'), not in tblSecurity_GroupUser.
16010         Set qdf3 = .QueryDefs("qrySecurity_User_28_02")
16020         Set rst3 = qdf3.OpenRecordset
16030         If rst3.BOF = True And rst3.EOF = True Then
                ' ** All's well.
16040           rst3.Close
16050           Set rst3 = Nothing
16060           Set qdf3 = Nothing
16070         Else
16080           rst3.Close
16090           Set rst3 = Nothing
16100           Set qdf3 = Nothing
                ' ** Append qrySecurity_User_28_02 (qrySecurity_User_28_01 (tblSecurity_User,
                ' ** linked to Users, tblSecurity_Group, as new tblSecurity_GroupUser records,
                ' ** for 'ViewOnly'), not in tblSecurity_GroupUser) to tblSecurity_GroupUser.
16110           Set qdf3 = .QueryDefs("qrySecurity_User_28_03")
16120           qdf3.Execute dbFailOnError
16130           Set qdf3 = Nothing
16140         End If

16150       End If  ' ** lngTmp01.

            ' ** Check TADemo.
            ' ** tblSecurity_User, linked to tblSecurity_GroupUser, just
            ' ** 'TADemo', with secusr_default_new, secgrpusr_default_new.
16160       Set qdf1 = .QueryDefs("qrySecurity_User_29_01")
16170       Set rst1 = qdf1.OpenRecordset
16180       If rst1.BOF = True And rst1.EOF = True Then
              ' ** Exactly as I'd hoped.
16190         rst1.Close
16200         Set rst1 = Nothing
16210         Set qdf1 = Nothing
16220       Else
16230         rst1.Close
16240         Set rst1 = Nothing
16250         Set qdf1 = Nothing
              ' ** Update qrySecurity_User_29_01 (tblSecurity_User, linked to tblSecurity_GroupUser,
              ' ** just'TADemo', with secusr_default_new, secgrpusr_default_new.).
16260         Set qdf1 = .QueryDefs("qrySecurity_User_29_02")
16270         qdf1.Execute
16280         Set qdf1 = Nothing
16290       End If

16300     Else
            ' ** Nothing to do.
16310     End If  ' ** Non-Default Users.

16320     .Close
16330   End With  ' ** dbs.
16340   Set dbs = Nothing

EXITP:
16350   Set rst1 = Nothing
16360   Set rst2 = Nothing
16370   Set rst3 = Nothing
16380   Set qdf1 = Nothing
16390   Set qdf2 = Nothing
16400   Set qdf3 = Nothing
16410   Set dbs = Nothing
16420   Security_SyncChk2 = blnRetVal
16430   Exit Function

ERRH:
16440   blnRetVal = False
16450   Select Case ERR.Number
        Case Else
16460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16470   End Select
16480   Resume EXITP

End Function

Public Function Security_SyncVer() As Boolean
' ** Make sure the version listed in Database Properties
' ** is the same as that in License Name.

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_SyncVer"

        Dim dbs As DAO.Database, rst As DAO.Recordset, qdf As DAO.QueryDef
        Dim strVer_DBProp As String, strVer_LicName As String, sngVer As Single
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

16510   blnRetVal = True

16520   strVer_DBProp = AppVersion_Get2  ' ** Module Function: modAppVersionFuncs.

16530   Set dbs = CurrentDb
16540   With dbs
16550     Set rst = .OpenRecordset("License Name", dbOpenDynaset, dbConsistent)
16560     With rst
16570       .MoveFirst
16580       strVer_LicName = CStr(![Version])  ' ** A Long Integer, so Revisions are incorporated into the Minor version number.
16590       intPos01 = InStr(strVer_LicName, ".")
16600       If intPos01 > 0 Then
              ' ** Yes, there's a Minor version number.
16610         If (intPos01 + 2) <= Len(strVer_LicName) Then
                ' ** Yes, there's a Revision version number.
16620           strVer_LicName = Left(strVer_LicName, (intPos01 + 1)) & "." & Mid(strVer_LicName, (intPos01 + 2))
16630         End If
16640       Else
              ' ** There's only a Major version number.
16650       End If
            'Debug.Print strVer_LicName
16660       If strVer_LicName <> strVer_DBProp Then
              ' ** Update the License Name table Version with the DB Properties Version.
16670         intPos01 = InStr(strVer_DBProp, ".")
16680         If intPos01 > 0 Then
                ' ** There's a Minor version number.
16690           sngVer = CSng(CLng(Val(Left(strVer_DBProp, (intPos01 - 1)))))
16700           strTmp01 = Mid(strVer_DBProp, (intPos01 + 1))
16710           intPos01 = InStr(strTmp01, ".")
16720           If intPos01 > 0 Then
                  ' ** There's a Revision version number.
16730             strTmp01 = Left(strTmp01, (intPos01 - 1)) & Mid(strTmp01, (intPos01 + 1))
16740             strTmp01 = CStr(sngVer) & "." & strTmp01
16750             sngVer = CSng(Round(Val(strTmp01), 3))
                  'Debug.Print CStr(sngVer)
16760           Else
                  ' ** There's no Revision version number.
16770             strTmp01 = CStr(sngVer) & "." & strTmp01
16780             sngVer = Val(strTmp01)
16790           End If
16800         Else
                ' ** There's no Minor version number.
16810           sngVer = Val(strVer_DBProp)
16820         End If
16830         .Edit
16840         ![Version] = sngVer
16850         .Update
16860       End If
16870       .Close
16880     End With
          ' ** For some reason, the Version, once updated, shows a whole slew of decimal places: 2.1.6000008583069!
          ' ** Damn floating point precision bug!
16890     Set qdf = .QueryDefs("qrySecurity_License_12")
16900     qdf.Execute
16910     .Close
16920   End With

EXITP:
16930   Set qdf = Nothing
16940   Set rst = Nothing
16950   Set dbs = Nothing
16960   Security_SyncVer = blnRetVal
16970   Exit Function

ERRH:
16980   blnRetVal = False
16990   Select Case ERR.Number
        Case Else
17000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17010   End Select
17020   Resume EXITP

End Function

Public Function Security_LicenseChk() As Boolean
' ** Check and confirm all license info.
' ** RESETS gstrTrustDataLocation TO THE LOCATION OF TA.LIC IF DIFFERENT!
' ** The Application.FileSearch was removed in Office 2007.
' ** If accessed, this property will return an error. To work around this issue, use
' ** the FileSystemObject to recursively search directories to find specific files.

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_LicenseChk"

        Dim fso As Scripting.FileSystemObject, fsfds As Scripting.Folders, fsfd As Scripting.Folder
        Dim strCurAppPath As String, strCurDataPath As String
        Dim intCurrentAccounts As Integer
        Dim strFirm As String
        Dim strExpires As String
        Dim strPricingExpires As String
        Dim strLimit As String
        Dim blnRetVal As Boolean

17110   blnRetVal = True   ' ** Unless proven otherwise.

17120   Set fso = CreateObject("Scripting.FileSystemObject")
17130   With fso
17140     If .FileExists(gstrTrustDataLocation & gstrFile_LIC) = False Then
            ' ** Search for Ta.lic in the current db directory and its subfolders.

17150       strCurAppPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
17160       strCurDataPath = CurrentBackendPath  ' ** Module Function: modFileUtilities.
17170       strPath = vbNullString

17180       Set fsfd = .GetFolder(strCurDataPath)
17190       Set fsfds = fsfd.SubFolders
17200       blnRetVal = FSO_Folders(fsfds)  ' ** Function: Below.

17210       If blnRetVal = True And strPath <> vbNullString Then
17220         gstrTrustDataLocation = strPath
17230       Else
17240         blnRetVal = False
17250       End If

17260     End If
17270   End With  ' ** fso.

17280   If blnRetVal = True Then

17290     strFirm = DecodeString(IniFile_Get("License", "Firm", EncodeString("Call Delta Data, Inc., for Licensing info."), _
            gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Functions: modCodeUtilities, Below.
17300     strExpires = DecodeString(IniFile_Get("License", "Expires", "", gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Functions: modCodeUtilities, Below.
17310     strPricingExpires = DecodeString(IniFile_Get("License", "Pricing", "", gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Functions: modCodeUtilities, Below.
17320     strLimit = DecodeString(IniFile_Get("License", "Limit", "", gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Functions: modCodeUtilities, Below.
17330     If Len(strLimit) = 0 Then strLimit = "0"

17340     gblnDemo = False
17350     intCurrentAccounts = DCount("*", "account")

17360     If Val(strLimit) < intCurrentAccounts Then
17370       blnRetVal = False
17380     End If

17390     If Len(strExpires) = 0 Then
17400       blnRetVal = False
17410     ElseIf CVDate(strExpires) < Now Then
17420       blnRetVal = False
17430     End If

          ' ** Set the global pricing allowed variable.
17440     If Len(strPricingExpires) = 0 Then
17450       gblnPricingAllowed = False
17460     Else
17470       gblnPricingAllowed = IIf(CVDate(strPricingExpires) >= Now, True, False)
17480     End If

17490   End If  ' ** blnRetVal.

EXITP:
17500   Set fso = Nothing
17510   Security_LicenseChk = blnRetVal
17520   Exit Function

ERRH:
17530   blnRetVal = False
17540   Select Case ERR.Number
        Case Else
17550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17560   End Select
17570   Resume EXITP

End Function

Public Function Security_PricingChk() As Boolean
' ** The Application.FileSearch was removed in Office 2007.
' ** If accessed, this property will return an error. To work around this issue, use
' ** the FileSystemObject to recursively search directories to find specific files.

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_PricingChk"

        Dim fso As Scripting.FileSystemObject, fsfds As Scripting.Folders, fsfd As Scripting.Folder
        Dim strPricingExpires As String, strCurAppPath As String, strCurDataPath As String
        Dim strConnect As String
        Dim blnRetVal As Boolean

17610   blnRetVal = False
17620   gblnPricingAllowed = False

17630   Set fso = CreateObject("Scripting.FileSystemObject")
17640   With fso
17650     If .FileExists(gstrTrustDataLocation & gstrFile_LIC) = False Then
            ' ** Search for Ta.lic in the current db directory and its subfolders.

17660       strCurAppPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
17670       strCurDataPath = CurrentBackendPath  ' ** Module Function: modFileUtilities.
17680       strPath = vbNullString

17690       Set fsfd = .GetFolder(strCurDataPath)
17700       Set fsfds = fsfd.SubFolders
17710       blnRetVal = FSO_Folders(fsfds)  ' ** Function: Below.

17720       If blnRetVal = True And strPath <> vbNullString Then
17730         blnRetVal = False
17740         strConnect = CurrentDb.TableDefs("ledger").Connect
17750         strConnect = Mid(strConnect, (InStr(strConnect, LNK_IDENT) + Len(LNK_IDENT)))
17760         strConnect = Parse_Path(strConnect)  ' ** Module Function: modFileUtilities.
17770         If GetUserName <> gstrDevUserName Then  ' ** Module Function: modFileUtilities.
17780           gstrTrustDataLocation = strPath
17790           blnRetVal = True
17800         Else
17810           If strConnect = strPath Then
17820             gstrTrustDataLocation = strPath
17830             blnRetVal = True
17840           End If
17850         End If
17860       End If

17870     Else
17880       blnRetVal = True
17890     End If
17900   End With  ' ** fso.

17910   If blnRetVal = True Then
17920     strPricingExpires = DecodeString(IniFile_Get("License", "Pricing", "", gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Functions: modCodeUtilities, Below.
          ' ** Set the global pricing allowed variable.
17930     If Trim(strPricingExpires) = vbNullString Then
17940       gblnPricingAllowed = False
17950     Else
17960       gblnPricingAllowed = IIf(CVDate(strPricingExpires) >= Now, True, False)
17970     End If
17980     blnRetVal = gblnPricingAllowed
17990   Else
18000     Beep
18010     MsgBox "A problem was encountered accessing your license file." & vbCrLf & vbCrLf & _
            "Please contact Delta Data, Inc., for assistance.", vbCritical + vbOKOnly, "License Error"
18020   End If  ' ** blnRetVal.

EXITP:
18030   Set fsfd = Nothing
18040   Set fsfds = Nothing
18050   Set fso = Nothing
18060   Security_PricingChk = blnRetVal
18070   Exit Function

ERRH:
18080   blnRetVal = False
18090   Select Case ERR.Number
        Case Else
18100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18110   End Select
18120   Resume EXITP

End Function

Public Function FSO_Folders(fsfds As Scripting.Folders) As Boolean
' ** Given folders, walk through each folder.

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "FSO_Folders"

        Dim fsfd As Scripting.Folder, fsfls As Scripting.FILES
        Dim blnRetVal As Boolean

18210   blnRetVal = True

18220   For Each fsfd In fsfds
18230     With fsfd
18240       Set fsfls = .FILES
18250       blnRetVal = FSO_Files(fsfd, fsfls)  ' ** Function: Below.
18260       If blnRetVal = True Then
18270         Exit For
18280       End If
18290     End With  ' ** fsfd.
18300   Next  ' ** fsfd.

EXITP:
18310   Set fsfls = Nothing
18320   Set fsfd = Nothing
18330   FSO_Folders = blnRetVal
18340   Exit Function

ERRH:
18350   blnRetVal = False
18360   Select Case ERR.Number
        Case Else
18370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18380   End Select
18390   Resume EXITP

End Function

Public Function FSO_Files(fsfd As Scripting.Folder, fsfls As Scripting.FILES) As Boolean
' ** Given files, walk through each file.

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "FSO_Files"

        Dim fsfds As Scripting.Folders, fsfl As Scripting.File
        Dim blnRetVal As Boolean

18410   blnRetVal = False

18420   For Each fsfl In fsfls
18430     With fsfl
18440       If Compare_StringA_StringB(.Name, "=", gstrFile_LIC) = True Then  ' ** Module Function: modStringFuncs.
18450         blnRetVal = True
18460         strPath = .Path  'THIS NEEDS TO BE AVAILABLE ELSEWHERE!
18470         gstrReportCallingForm = strPath  ' ** Just borrowing this.
18480         Exit For
18490       End If
18500     End With  ' ** fsfl.
18510   Next  ' ** fsfl.

18520   If blnRetVal = False Then
18530     Set fsfds = fsfd.SubFolders
18540     blnRetVal = FSO_Folders(fsfds)  ' ** Function: Above.
18550   End If

EXITP:
18560   Set fsfl = Nothing
18570   Set fsfds = Nothing
18580   FSO_Files = blnRetVal
18590   Exit Function

ERRH:
18600   blnRetVal = False
18610   Select Case ERR.Number
        Case Else
18620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18630   End Select
18640   Resume EXITP

End Function

Public Function Security_PW_Exp_Chk(frm As Access.Form) As Boolean
' ** Check password expiration.
' ** DEFAULT USER PASSWORDS DON'T EXPIRE!

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_PW_Exp_Chk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strCurUser As String
        Dim lngUsrs As Long, arr_varUsr As Variant
        Dim strOrigin As String, datOrigin As Date, lngOrigin As Long, lngToday As Long
        Dim lngCycleDays As Long, lngWarnScreen As Long, lngWarnMsg As Long
        Dim blnFound As Boolean
        Dim lngX As Long
        Dim blnRetVal As Boolean

        'Const U_ID     As Integer = 0
        'Const U_GUID   As Integer = 1
        Const U_ORIGIN As Integer = 2
        Const U_NAM1   As Integer = 3
        Const U_DEF1   As Integer = 4
        Const U_NAM2   As Integer = 5
        Const U_DEF2   As Integer = 6
        Const U_NAM3   As Integer = 7
        Const U_DEF3   As Integer = 8

18710   blnRetVal = True

18720   strCurUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
18730   lngUsrs = 0&

18740   Set dbs = CurrentDb

        ' ** Password Cycle:
        ' **  ![seclic_cycle]
        ' ** Warning Screen:
        ' **  30 days for cycle > 30
        ' **  7 days for cycle <= 30
        ' **  ![seclic_cycle_screen]
        ' ** Warning Message:
        ' **  7 days for cycle > 30
        ' **  3 days for cycle <= 30
        ' **  ![seclic_cycle_message]

18750   With dbs
18760     Set rst = .OpenRecordset("tblSecurity_License", dbOpenDynaset, dbConsistent)
18770     With rst
18780       If .BOF = True And .EOF = True Then
              ' ** Horrible! Horrible!
18790         blnRetVal = False
18800       Else
18810         .MoveFirst
18820         If IsNull(![seclic_cycle]) = False Then
18830           lngCycleDays = CLng(DecodeString(![seclic_cycle]))  ' ** Module Function: modCodeUtilities.
18840         Else
                ' ** Shouldn't have been allowed to happen.
18850           .Edit
18860           ![seclic_cycle] = EncodeString(CStr(DEF_CYCLE))  ' ** Module Function: modCodeUtilities.
18870           lngCycleDays = DEF_CYCLE
18880           ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
18890           ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
18900           ![seclic_datemodified] = Now()
18910           .Update
18920         End If
18930         If IsNull(![seclic_cycle_screen]) = False Then
18940           lngWarnScreen = CLng(DecodeString(![seclic_cycle_screen]))  ' ** Module Function: modCodeUtilities.
18950         Else
                ' ** Ditto.
18960           .Edit
18970           ![seclic_cycle_screen] = EncodeString(CStr(DEF_SCREEN))  ' ** Module Function: modCodeUtilities.
18980           lngWarnScreen = DEF_SCREEN
18990           ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
19000           ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
19010           ![seclic_datemodified] = Now()
19020           .Update
19030         End If
19040         If IsNull(![seclic_cycle_message]) = False Then
19050           lngWarnMsg = CLng(DecodeString(![seclic_cycle_message]))  ' ** Module Function: modCodeUtilities.
19060         Else
                ' ** Ditto.
19070           .Edit
19080           ![seclic_cycle_message] = EncodeString(CStr(DEF_MSGBOX))  ' ** Module Function: modCodeUtilities.
19090           lngWarnMsg = DEF_MSGBOX
19100           ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
19110           ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
19120           ![seclic_datemodified] = Now()
19130           .Update
19140         End If
19150       End If
19160       .Close
19170     End With  ' ** rst.
19180   End With  ' ** dbs.

19190   If blnRetVal = True Then

          ' ** Collect the Users from tblSecurity_User and _~xusr into arr_varUsr().
19200     Set qdf = dbs.QueryDefs("qryXUsr_06")
19210     Set rst = qdf.OpenRecordset
19220     With rst
19230       .MoveLast
19240       lngUsrs = .RecordCount
19250       .MoveFirst
19260       arr_varUsr = .GetRows(lngUsrs)
            ' ****************************************************
            ' ** Array: arr_varUsr()
            ' **
            ' **   Field  Element  Name               Constant
            ' **   =====  =======  =================  ==========
            ' **     1       0     xusr_id            U_ID
            ' **     2       1     s_GUID             U_GUID
            ' **     3       2     xusr_origin        U_ORIGIN
            ' **     4       3     secusr_name1       U_NAM1
            ' **     5       4     secusr_default1    U_DEF1
            ' **     6       5     secusr_name2       U_NAM2
            ' **     7       6     secusr_default2    U_DEF2
            ' **     8       7     secusr_name3       U_NAM3
            ' **     9       8     secusr_default3    U_DEF3
            ' **
            ' ****************************************************
19270       .Close
19280     End With

          ' ** Check if the current logged-in user is a default user.
          ' ** DEFAULT USER PASSWORDS DON'T EXPIRE!
          ' ** This is messy because I don't want to have separate Demo and non-Demo queries.
19290     blnFound = False: strOrigin = vbNullString
19300     For lngX = 0& To (lngUsrs - 1&)
19310       If IsNull(arr_varUsr(U_NAM1, lngX)) = False Then
19320         If arr_varUsr(U_NAM1, lngX) = strCurUser And arr_varUsr(U_DEF1, lngX) = True Then
19330           blnFound = True
19340           Exit For
19350         ElseIf arr_varUsr(U_NAM1, lngX) = strCurUser And arr_varUsr(U_DEF1, lngX) = False Then
19360           strOrigin = arr_varUsr(U_ORIGIN, lngX)
19370           Exit For
19380         End If
19390       End If
19400       If blnFound = False And strOrigin = vbNullString Then
19410         If IsNull(arr_varUsr(U_NAM2, lngX)) = False Then
19420           If arr_varUsr(U_NAM2, lngX) = strCurUser And arr_varUsr(U_DEF2, lngX) = True Then
19430             blnFound = True
19440             Exit For
19450           ElseIf arr_varUsr(U_NAM2, lngX) = strCurUser And arr_varUsr(U_DEF2, lngX) = False Then
19460             strOrigin = arr_varUsr(U_ORIGIN, lngX)
19470             Exit For
19480           End If
19490         End If
19500       End If
19510       If blnFound = False And strOrigin = vbNullString Then
19520         If IsNull(arr_varUsr(U_NAM3, lngX)) = False And blnFound = False Then
19530           If arr_varUsr(U_NAM3, lngX) = strCurUser And arr_varUsr(U_DEF3, lngX) = True Then
19540             blnFound = True
19550             Exit For
19560           ElseIf arr_varUsr(U_NAM3, lngX) = strCurUser And arr_varUsr(U_DEF3, lngX) = False Then
19570             strOrigin = arr_varUsr(U_ORIGIN, lngX)
19580             Exit For
19590           End If
19600         End If
19610       End If
19620     Next

          ' ** Password Cycle:
          ' **  ![seclic_cycle]
          ' ** Warning Screen:
          ' **  30 days for cycle > 30
          ' **  7 days for cycle <= 30
          ' **  ![seclic_cycle_screen]
          ' ** Warning Message:
          ' **  7 days for cycle > 30
          ' **  3 days for cycle <= 30
          ' **  ![seclic_cycle_message]

19630     If blnFound = False Then
            ' ** User is not a Default User.
19640       If lngCycleDays > 0& Then
              ' ** Zero seclic_cycle means passwords don't expire.
19650         datOrigin = CDate(DecodeString(strOrigin))  ' ** Module Function: modCodeUtilities.
19660         lngOrigin = CLng(CDbl(datOrigin))
19670         lngToday = CLng(CDbl(Date))
19680         If lngOrigin + lngCycleDays <= lngToday Then  ' ** Origin plus 1 year must be less than Today to have expired.
                ' ** Their password has expired!
19690           blnRetVal = False
                ' ** Give notice and exit!
19700           MsgBox "Your password has expired." & vbCrLf & vbCrLf & _
                  "Contact your administrator.", vbCritical + vbOKOnly, "Password Expired"
19710         Else
19720           If (((lngOrigin + lngCycleDays) - lngToday) <= lngWarnScreen) Then
                  ' ** Password has 30 days or less to expiration.
                  ' ** Give notice on frmMenu_Title,
19730             If (lngCycleDays > lngWarnScreen) Or ((lngCycleDays <= lngWarnScreen) And _
                      (((lngOrigin + lngCycleDays) - lngToday) <= lngWarnMsg)) Then
                    ' ** On-screen 30-day notice for cycles greater than 30 days,
                    ' ** and 7-day notice for cycles less than or equal to 30 days.
19740               With frm
19750                 If lngWarnScreen > 0& Then
19760                   If .Eval_lbl.Visible = True Then
19770                     .PW_Expiration_lbl.Top = .Eval_lbl.Top - .PW_Expiration_lbl.Height
19780                   End If
19790                   .PW_Expiration_lbl.Caption = "Your password will expire in: " & CStr((lngOrigin + lngCycleDays) - lngToday) & " days"
19800                   .PW_Expiration_lbl.Visible = True
19810                 End If
19820                 If ((lngCycleDays > lngWarnScreen) And (((lngOrigin + lngCycleDays) - lngToday) <= lngWarnMsg)) Or _
                          ((lngCycleDays <= lngWarnScreen) And (((lngOrigin + lngCycleDays) - lngToday) <= lngWarnMsg)) Then
                        ' ** MsgBox 7-day warning for cycles greater than 30 days,
                        ' ** and 3-day warning for cycles less than or equal to 30 days.
19830                   If lngWarnMsg > 0& Then
19840                     Beep
19850                     MsgBox "Your password will expire in: " & CStr((lngOrigin + lngCycleDays) - lngToday) & " days", _
                            vbExclamation + vbOKOnly, "Password Renewal Imminent"
19860                   End If

19870                 End If
19880               End With
19890             End If
19900           End If
19910         End If
19920       End If  ' ** lngCycleDays.
19930     End If  ' ** blnFound.

19940   End If  ' ** blnRetVal.

19950   dbs.Close

EXITP:
19960   Set rst = Nothing
19970   Set qdf = Nothing
19980   Set dbs = Nothing
19990   Security_PW_Exp_Chk = blnRetVal
20000   Exit Function

ERRH:
20010   blnRetVal = False
20020   Select Case ERR.Number
        Case Else
20030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20040     MsgBox "An error occurred within a password security routine." & vbCrLf & vbCrLf & _
            "Contact Delta Data, Inc., for assistance.", vbCritical + vbOKOnly, "Password Security Error"
20050   End Select
20060   Resume EXITP

End Function

Public Function Security_PW_Exp_Get(varItem As Variant) As Long

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_PW_Exp_Get"

        Dim strCycle As String, strCurUser As String
        Dim blnDefUser As Boolean, datOrigin As Date
        Dim varTmp00 As Variant, lngTmp01 As Long
        Dim lngRetVal As Long

20110   lngRetVal = -1&

20120   If IsNull(varItem) = False Then

20130     varTmp00 = DLookup("[seclic_cycle]", "tblSecurity_License")
20140     Select Case IsNull(varTmp00)
          Case True
            ' ** No setting. Shouldn't happen.
20150       strCycle = "0"
20160     Case False
20170       strCycle = DecodeString(CStr(varTmp00))  ' ** Module Function: modCodeUtilities.
20180     End Select

20190     Select Case varItem
          Case "Cycle"
            ' ** 0 = No Expiration.
20200       lngRetVal = CLng(strCycle)
20210     Case "RemDays"

20220       lngTmp01 = CLng(strCycle)
20230       If lngTmp01 = 0 Then
              ' ** No Expiration.
20240         lngRetVal = 99999
20250       Else

20260         strCurUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.

20270         varTmp00 = DLookup("[secusr_default]", "tblSecurity_User", "[secusr_name] = '" & strCurUser & "'")
20280         Select Case IsNull(varTmp00)
              Case True
                ' ** User doesn't exist!
20290           blnDefUser = False
20300           lngRetVal = 0&
20310         Case False
20320           blnDefUser = CBool(varTmp00)
20330         End Select

20340         If lngRetVal <> 0& Then
20350           Select Case blnDefUser
                Case True
                  ' ** Remaining days are infinite. User is immortal.
20360             lngRetVal = 99999  ' ** Approx. 274 years.
20370           Case False
                  ' ** Users, with xusr_origin.
20380             varTmp00 = DLookup("[xusr_origin]", "qrySecurity_User_20", "[Username] = '" & strCurUser & "'")
20390             Select Case IsNull(varTmp00)
                  Case True
                    ' ** No password inception date. Shouldn't happen.
20400               lngRetVal = 0&
20410             Case False
20420               Select Case IsDate(DecodeString(CStr(varTmp00)))  ' ** Module Function: modCodeUtilities.
                    Case True
20430                 datOrigin = CDate(DecodeString(CStr(varTmp00)))  ' ** Module Function: modCodeUtilities.
20440                 lngTmp01 = CLng(CDbl(Date) - CDbl(datOrigin))  ' ** Days since password change.
20450                 If lngTmp01 >= CLng(strCycle) Then
                        ' ** Expired.
20460                   lngRetVal = 0
20470                 Else
20480                   lngRetVal = (CLng(strCycle) - lngTmp01)  ' ** Days remaining.
20490                 End If
20500               Case False
                      ' ** Invalid password inception date. Shouldn't happen.
20510                 lngRetVal = 0
20520               End Select
20530             End Select
20540           End Select
20550         End If
20560       End If
20570     End Select
20580   End If

EXITP:
20590   Security_PW_Exp_Get = lngRetVal
20600   Exit Function

ERRH:
20610   lngRetVal = -9&
20620   Select Case ERR.Number
        Case Else
20630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20640   End Select
20650   Resume EXITP

End Function

Private Function Security_PIN_To_PW(strPIN As String) As String

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_PIN_To_PW"

        Dim intLen As Integer
        Dim blnFound As Boolean
        Dim strTmp01 As String
        Dim intX As Integer
        Dim strRetVal As String

20710   strRetVal = vbNullString

20720   intLen = Len(strPIN)
20730   strTmp01 = vbNullString
20740   If intLen >= 5 Then
20750     If intLen <= 14 Then
20760       strTmp01 = strPIN
20770     Else
20780       strTmp01 = Left(strPIN, 14)
20790     End If
20800   Else
20810     strTmp01 = Left(strPIN & String(5, "x"), 5)  ' ** Pad with X's.
20820   End If
20830   intLen = Len(strTmp01)
20840   blnFound = False
20850   For intX = 1 To intLen
20860     If Asc(Mid(strTmp01, intX, 1)) >= 97 And Asc(Mid(strTmp01, intX, 1)) <= 122 Then
            ' ** Has lower case letter.
20870       blnFound = True
20880       Exit For
20890     End If
20900   Next
20910   Select Case blnFound
        Case True
          ' ** Keep checking.
20920   Case False
20930     strTmp01 = Left(strTmp01, 1) & LCase$(Mid(strTmp01, 2))
20940   End Select
20950   blnFound = False
20960   For intX = 1 To intLen
20970     If Asc(Mid(strTmp01, intX, 1)) >= 65 And Asc(Mid(strTmp01, intX, 1)) <= 90 Then
            ' ** Has upper case.
20980       blnFound = True
20990       Exit For
21000     End If
21010   Next
21020   Select Case blnFound
        Case True
          ' ** Keep checking.
21030   Case False
21040     strTmp01 = UCase$(Left(strTmp01, 1)) & Mid(strTmp01, 2)
21050   End Select
21060   blnFound = False
21070   For intX = 1 To intLen
21080     If Asc(Mid(strTmp01, intX, 1)) >= 48 And Asc(Mid(strTmp01, intX, 1)) <= 57 Then
21090       blnFound = True
21100       Exit For
21110     End If
21120   Next
21130   Select Case blnFound
        Case True
          ' ** Good to go.
21140   Case False
21150     If intLen = 14 Then
21160       strTmp01 = Left(strTmp01, 13) & "1"
21170     Else
21180       strTmp01 = strTmp01 & "1"
21190     End If
21200   End Select
21210   strRetVal = strTmp01

EXITP:
21220   Security_PIN_To_PW = strRetVal
21230   Exit Function

ERRH:
21240   strRetVal = RET_ERR
21250   Select Case ERR.Number
        Case Else
21260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21270   End Select
21280   Resume EXITP

End Function

Public Function IsSingleUser(Optional varInitVars As Variant) As Boolean

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "IsSingleUser"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCnt1 As Long, lngCnt2 As Long, lngCnt3 As Long
        Dim blnInitVars As Boolean
        Dim blnRetVal As Boolean

21310   lngCnt1 = 0&: lngCnt2 = 0&: lngCnt3 = 0&

21320   If IsMissing(varInitVars) = True Then
21330     blnInitVars = False
21340   Else
21350     blnInitVars = varInitVars
21360   End If

21370   Set dbs = CurrentDb
21380   With dbs
          ' ** Count the number of non-default users in tblSecurity_User.
21390     Set qdf = .QueryDefs("qryJournal_01")
21400     Set rst = qdf.OpenRecordset
21410     With rst
21420       .MoveFirst
21430       lngCnt1 = ![cnt]
21440       .Close
21450     End With
21460     If lngCnt1 <= 1& Then
            ' ** Absolutely.
21470       blnRetVal = True
21480       If blnInitVars = True Then
21490         gblnSingleUser = True
21500         glngUserCntLedger = 1&
21510       End If
21520     Else
            ' ** Count the number of non-System users in Journal.
21530       Set qdf = .QueryDefs("qryJournal_02b")  ' ** Returns 0 if Journal empty (I tested it!).
21540       Set rst = qdf.OpenRecordset
21550       With rst
21560         .MoveFirst
21570         lngCnt2 = ![cnt]
21580         .Close
21590       End With
21600       If blnInitVars = True Or (blnInitVars = False And glngUserCntLedger <= 0&) Then
              ' ** Count the number of non-System users in Ledger.
21610         Set qdf = .QueryDefs("qryJournal_03b")  ' ** Returns 0 if Ledger empty (I tested it!).
21620         Set rst = qdf.OpenRecordset
21630         With rst
21640           .MoveFirst
21650           lngCnt3 = ![cnt]
21660           .Close
21670         End With
21680         glngUserCntLedger = lngCnt3
21690       Else
              ' ** This will cut down the querying time on large Ledgers.
21700         lngCnt3 = glngUserCntLedger
21710       End If
21720       If lngCnt2 > 1& Or lngCnt3 > 1& Then
21730         blnRetVal = False
21740         If blnInitVars = True Then gblnSingleUser = True
21750       Else
21760         blnRetVal = True
21770         If blnInitVars = True Then gblnSingleUser = True
21780       End If
21790     End If
21800     .Close
21810   End With

EXITP:
21820   Set rst = Nothing
21830   Set qdf = Nothing
21840   Set dbs = Nothing
21850   IsSingleUser = blnRetVal
21860   Exit Function

ERRH:
21870   blnRetVal = False
21880   Select Case ERR.Number
        Case Else
21890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21900   End Select
21910   Resume EXITP

End Function

Public Function Security_SetGroups(frm As Access.Form, ctl As Access.Control) As Boolean
' ** Currently not used; replaced by tblSecurity_Group query.

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_SetGroups"

        Dim blnRetVal As Boolean

22010   blnRetVal = True

22020   With ctl
22030     If .RowSourceType = "Value List" Then
            ' ** DataEntry;ViewOnly;Admins
22040       .RowSource = SGRP_DATAENTRY & ";" & SGRP_VIEWONLY & ";" & SGRP_ADMINS
22050     End If
22060   End With

EXITP:
22070   Security_SetGroups = blnRetVal
22080   Exit Function

ERRH:
22090   Select Case ERR.Number
        Case Else
22100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22110   End Select
22120   Resume EXITP

End Function

Public Function Security_Update() As Boolean
' ** In the future, this may be used for updating security
' ** settings when the program security is made more robust.

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_Update"

        Dim blnRetVal As Boolean

22210   blnRetVal = True

        'If the login is someone from Admins, check the Users table,
        'and automatically switch "Default Group", "Primary Group"
        'This should be done the first time the new update is opened.
        'Create a backup copy of the Users table in one of the backends,
        'in case they have to revert, or the backends are shared by
        'new frontends and old frontends.

        'We'll also have to change the security settings on the backends.
        'That means the MDW will have to be switched!!!!!!!!!!!!!!!!!!!!!

EXITP:
22220   Security_Update = blnRetVal
22230   Exit Function

ERRH:
22240   blnRetVal = False
22250   Select Case ERR.Number
        Case Else
22260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22270   End Select
22280   Resume EXITP

End Function

Public Function Security_Hidden_Get(varObjectName As Variant, Optional varObjectType As Variant) As Boolean
' ** Return True if the specified object is to be hidden.
' ** Called by:  NO LONGER USED!
' **   qrySecurity_Stat_01 - Tables
' **   qrySecurity_Stat_02 - Queries
' **   qrySecurity_Stat_03 - Forms
' **   qrySecurity_Stat_04 - Modules

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_Hidden_Get"

        Dim intObjType As Integer
        Dim lngX As Long
        Dim blnRetVal As Boolean

22310   blnRetVal = False

22320   If IsNull(varObjectName) = False Then
22330     If IsMissing(varObjectType) = True Then
22340       intObjType = acTable
22350     Else
22360       intObjType = varObjectType
22370     End If
22380     If Security_Hidden_Load = True Then  ' ** Function: Below.
22390       For lngX = 0& To (glngHids - 1&)
22400         If garr_varHid(H_TYP, lngX) = intObjType And garr_varHid(H_NAM, lngX) = varObjectName Then
22410           blnRetVal = True
22420           Exit For
22430         End If
22440       Next
22450     End If
22460   End If

EXITP:
22470   Security_Hidden_Get = blnRetVal
22480   Exit Function

ERRH:
22490   blnRetVal = False
22500   Select Case ERR.Number
        Case Else
22510     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22520   End Select
22530   Resume EXITP

End Function

Public Function Security_Hidden_Set() As Boolean

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_Hidden_Set"

        Dim lngX As Long
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

22610   blnRetVal = True

22620   Security_Hidden_Load  ' ** Function: Below.

22630   For lngX = 0& To (glngHids - 1&)
22640     varTmp00 = Null
22650 On Error Resume Next
22660     varTmp00 = GetHiddenAttribute(garr_varHid(H_TYP, lngX), garr_varHid(H_NAM, lngX))
22670     If IsNull(varTmp00) = False Then
22680 On Error GoTo ERRH
22690       If GetHiddenAttribute(garr_varHid(H_TYP, lngX), garr_varHid(H_NAM, lngX)) = False Then
22700         SetHiddenAttribute garr_varHid(H_TYP, lngX), garr_varHid(H_NAM, lngX), True
22710       End If
22720     Else
22730 On Error GoTo ERRH
22740     End If
22750   Next

22760   Beep

EXITP:
22770   Security_Hidden_Set = blnRetVal
22780   Exit Function

ERRH:
22790   blnRetVal = False
22800   Select Case ERR.Number
        Case Else
22810     Beep
22820     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()", _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
22830   End Select
22840   Resume EXITP

End Function

Public Function Security_Hidden_Load() As Boolean
' ** Return an array of Trust Accountant objects on which to set the Hidden attribute.

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_Hidden_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

22910   blnRetVal = True

22920   If glngHids = 0& Or IsEmpty(garr_varHid) = True Then
22930     Set dbs = CurrentDb
22940     With dbs
            ' ** Update qrySecurity_Stat_10d (tblDatabase_Table, 'tblImportExport_..' for sec_hidden = False, by specified CurrentAppName()).
22950       Set qdf = .QueryDefs("qrySecurity_Stat_10e_Table")
22960       qdf.Execute
22970       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_10f (tblDatabase_Table, 'tblRegistry..' for sec_hidden = False, by specified CurrentAppName()).
22980       Set qdf = .QueryDefs("qrySecurity_Stat_10g_Table")
22990       qdf.Execute
23000       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_10h (tblDatabase_Table, 'tblRelation_View_..' for sec_hidden = False, by specified CurrentAppName()).
23010       Set qdf = .QueryDefs("qrySecurity_Stat_10i_Table")
23020       qdf.Execute
23030       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_10j (tblDatabase_Table, 'tblSecurity_..' for sec_hidden = False, by specified CurrentAppName()).
23040       Set qdf = .QueryDefs("qrySecurity_Stat_10k_Table")
23050       qdf.Execute
23060       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_10l (tblDatabase_Table, 'tblTemplate_Zeta_..' for sec_hidden = False, by specified CurrentAppName()).
23070       Set qdf = .QueryDefs("qrySecurity_Stat_10m_Table")
23080       qdf.Execute
23090       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_10n (tblDatabase_Table, 'tblXAdmin_..' for sec_hidden = False, by specified CurrentAppName()).
23100       Set qdf = .QueryDefs("qrySecurity_Stat_10o_Table")
23110       qdf.Execute
23120       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_10p (tblDatabase_Table, 'tmpXAdmin_..' for sec_hidden = False, by specified CurrentAppName()).
23130       Set qdf = .QueryDefs("qrySecurity_Stat_10q_Table")
23140       qdf.Execute
23150       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_12c (tblQuery, 'zz_qry_System_nn' for sec_hidden = False, by specified CurrentAppName()).
23160       Set qdf = .QueryDefs("qrySecurity_Stat_12d_Query")
23170       qdf.Execute
23180       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_12e (tblQuery, 'qrySecurity_..' for sec_hidden = False, by specified CurrentAppName()).
23190       Set qdf = .QueryDefs("qrySecurity_Stat_12f_Query")
23200       qdf.Execute
23210       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_12g (tblQuery, 'qryXAdmin_..' for sec_hidden = False, by specified CurrentAppName()).
23220       Set qdf = .QueryDefs("qrySecurity_Stat_12h_Query")
23230       qdf.Execute
23240       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_12i (tblQuery, 'qryRelationView_..' for sec_hidden = False, by specified CurrentAppName()).
23250       Set qdf = .QueryDefs("qrySecurity_Stat_12j_Query")
23260       qdf.Execute
23270       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_12k (tblQuery, 'qryXUsr_..' for sec_hidden = False, by specified CurrentAppName()).
23280       Set qdf = .QueryDefs("qrySecurity_Stat_12l_Query")
23290       qdf.Execute
23300       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_13c (tblVBComponent, 'modCode..', 'modSecurity..', 'zz_mod_..', for sec_hidden = False, by specified CurrentAppName()).
23310       Set qdf = .QueryDefs("qrySecurity_Stat_13d_Module")
23320       qdf.Execute
23330       Set qdf = Nothing
            ' ** Update qrySecurity_Stat_14c (    tblMacro, 'mcrXAdmin_..', 'mcrAutoNumber..', etc., for sec_hidden = False, by CurrentAppName()).
23340       Set qdf = .QueryDefs("qrySecurity_Stat_14d_Macro")
23350       qdf.Execute
23360       Set qdf = Nothing
            ' ** qrySecurity_Stat_15 (Union of qrySecurity_Stat_10b_Table (qrySecurity_Stat_10a_Table (tblDatabase_Table,
            ' ** linked to qrySecurity_Stat_10c_Table (tblDatabase_Table_Link, by specified CurrentAppName()), just tables
            ' ** available at runtime), just sec_hidden = True), qrySecurity_Stat_11b_Form (qrySecurity_Stat_11a_Form
            ' ** (tblForm, just forms available at runtime, by specified CurrentAppName()), just sec_hidden = True),
            ' ** qrySecurity_Stat_12b_Query (qrySecurity_Stat_12a_Query (tblQuery, just queries available at runtime, by
            ' ** specified CurrentAppName()), just sec_hidden = True), qrySecurity_Stat_13b_Module (qrySecurity_Stat_13a_Module
            ' ** (tblVBComponent, just modules available at runtime, by specified CurrentAppName()), just sec_hidden = True),
            ' ** qrySecurity_Stat_14b_Macro (qrySecurity_Stat_14a_Macro (tblMacro, just macros available at runtime, by
            ' ** specified CurrentAppName()), just sec_hidden = True)), just needed fields.
23370       Set qdf = .QueryDefs("qrySecurity_Stat_16")
23380       Set rst = qdf.OpenRecordset
23390       With rst
23400         .MoveLast
23410         glngHids = .RecordCount
23420         .MoveFirst
23430         garr_varHid = .GetRows(glngHids)
              ' *************************************************
              ' ** Array: garr_varHidden()
              ' **
              ' **   Field  Element  Name            Constant
              ' **   =====  =======  ==============  ==========
              ' **     1       0     objtype_type    H_TYP
              ' **     2       1     obj_id          H_ID
              ' **     3       2     obj_name        H_NAM
              ' **     4       3     sec_hidden      H_HID
              ' **
              ' *************************************************
23440         .Close
23450       End With
23460       .Close
23470     End With
23480   End If

EXITP:
23490   Set rst = Nothing
23500   Set qdf = Nothing
23510   Set dbs = Nothing
23520   Security_Hidden_Load = blnRetVal
23530   Exit Function

ERRH:
23540   blnRetVal = False
23550   Select Case ERR.Number
        Case Else
23560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23570   End Select
23580   Resume EXITP

End Function

Public Function Security_Hidden_List() As Boolean
' ** List what's currently hidden.

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_Hidden_List"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, tdf As DAO.TableDef, cntr As DAO.Container, doc As DAO.Document, rst As DAO.Recordset
        Dim lngHids As Long, arr_varHid() As Variant
        Dim lngTbls As Long, lngQrys As Long, lngFrms As Long, lngMcrs As Long, lngRpts As Long, lngMods As Long
        Dim lngLastType As Long, strLastType As String, strLastField As String
        Dim lngDbsID As Long, strLastDbsName As String, strThisDbsName As String
        Dim blnSkip As Boolean
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varHid().
        Const H_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const H_TYP  As Integer = 0
        Const H_IDNT As Integer = 1
        Const H_NAM  As Integer = 2
        Const H_HIDE As Integer = 3
        Const H_CONN As Integer = 4

23610   blnRetVal = True

23620   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
23630   DoEvents

23640   lngTbls = 0&: lngQrys = 0&: lngFrms = 0&: lngMcrs = 0&: lngRpts = 0&: lngMods = 0&

23650   Set dbs = CurrentDb
23660   With dbs
          'Debug.Print "'HIDDEN TABLES:"
23670     For Each tdf In .TableDefs
23680       With tdf
23690         If Left(.Name, 4) <> "MSys" And Left(.Name, 4) <> "USys" Then
23700           If GetHiddenAttribute(acTable, .Name) = True Then
23710             lngHids = lngHids + 1&
23720             lngE = lngHids - 1&
23730             ReDim Preserve arr_varHid(H_ELEMS, lngE)
                  ' *************************************************
                  ' ** Array: arr_varHid()
                  ' **
                  ' **   Field  Element  Name            Constant
                  ' **   =====  =======  ==============  ==========
                  ' **     1       0     objtype_type    H_TYP
                  ' **     2       1     ID Number       H_IDNT
                  ' **     3       2     Object Name     H_NAM
                  ' **     4       3     Is Hid          H_HIDE
                  ' **     5       4     dbs_connect     H_CONN
                  ' **
                  ' *************************************************
23740             arr_varHid(H_TYP, lngE) = acTable
23750             arr_varHid(H_IDNT, lngE) = CLng(0)
23760             arr_varHid(H_NAM, lngE) = .Name
23770             arr_varHid(H_HIDE, lngE) = CBool(True)
23780             If .Connect <> vbNullString Then
23790               arr_varHid(H_CONN, lngE) = .Connect
23800             Else
23810               arr_varHid(H_CONN, lngE) = Null
23820             End If
23830             lngTbls = lngTbls + 1&
                  'Debug.Print "'  " & .Name
23840           End If
23850         End If
23860       End With
23870       Set tdf = Nothing
23880     Next  ' ** tdf.
          'Debug.Print "'HIDDEN QUERIES:"
23890     For Each qdf In .QueryDefs
23900       With qdf
23910         If Left(.Name, 4) <> "~sq_" Then
23920           If GetHiddenAttribute(acQuery, .Name) = True Then
23930             lngHids = lngHids + 1&
23940             lngE = lngHids - 1&
23950             ReDim Preserve arr_varHid(H_ELEMS, lngE)
23960             arr_varHid(H_TYP, lngE) = acQuery
23970             arr_varHid(H_IDNT, lngE) = CLng(0)
23980             arr_varHid(H_NAM, lngE) = .Name
23990             arr_varHid(H_HIDE, lngE) = CBool(True)
24000             arr_varHid(H_CONN, lngE) = Null
24010             lngQrys = lngQrys + 1&
                  'Debug.Print "'  " & .Name
24020           End If
24030         End If
24040       End With
24050       Set qdf = Nothing
24060     Next  ' ** tdf.
          'Debug.Print "'HIDDEN FORMS:"
24070     Set cntr = .Containers("Forms")
24080     With cntr
24090       For Each doc In .Documents
24100         With doc
24110           If GetHiddenAttribute(acForm, .Name) = True Then
24120             lngHids = lngHids + 1&
24130             lngE = lngHids - 1&
24140             ReDim Preserve arr_varHid(H_ELEMS, lngE)
24150             arr_varHid(H_TYP, lngE) = acForm
24160             arr_varHid(H_IDNT, lngE) = CLng(0)
24170             arr_varHid(H_NAM, lngE) = .Name
24180             arr_varHid(H_HIDE, lngE) = CBool(True)
24190             arr_varHid(H_CONN, lngE) = Null
24200             lngFrms = lngFrms + 1&
                  'Debug.Print "'  " & .Name
24210           End If
24220         End With
24230         Set doc = Nothing
24240       Next
24250     End With
24260     Set cntr = Nothing
          'Debug.Print "'HIDDEN MACROS:"
24270     Set cntr = .Containers("Scripts")
24280     With cntr
24290       For Each doc In .Documents
24300         With doc
24310           If GetHiddenAttribute(acMacro, .Name) = True Then
24320             lngHids = lngHids + 1&
24330             lngE = lngHids - 1&
24340             ReDim Preserve arr_varHid(H_ELEMS, lngE)
24350             arr_varHid(H_TYP, lngE) = acMacro
24360             arr_varHid(H_IDNT, lngE) = CLng(0)
24370             arr_varHid(H_NAM, lngE) = .Name
24380             arr_varHid(H_HIDE, lngE) = CBool(True)
24390             arr_varHid(H_CONN, lngE) = Null
24400             lngMcrs = lngMcrs + 1&
                  'Debug.Print "'  " & .Name
24410           End If
24420         End With
24430         Set doc = Nothing
24440       Next
24450     End With
24460     Set cntr = Nothing
          'Debug.Print "'HIDDEN REPORTS:"
24470     Set cntr = .Containers("Reports")
24480     With cntr
24490       For Each doc In .Documents
24500         With doc
24510           If GetHiddenAttribute(acReport, .Name) = True Then
24520             lngHids = lngHids + 1&
24530             lngE = lngHids - 1&
24540             ReDim Preserve arr_varHid(H_ELEMS, lngE)
24550             arr_varHid(H_TYP, lngE) = acMacro
24560             arr_varHid(H_IDNT, lngE) = CLng(0)
24570             arr_varHid(H_NAM, lngE) = .Name
24580             arr_varHid(H_HIDE, lngE) = CBool(True)
24590             arr_varHid(H_CONN, lngE) = Null
24600             lngRpts = lngRpts + 1&
                  'Debug.Print "'  " & .Name
24610           End If
24620         End With
24630         Set doc = Nothing
24640       Next
24650     End With
24660     Set cntr = Nothing
          'Debug.Print "'HIDDEN MODULES:"
24670     Set cntr = .Containers("Modules")
24680     With cntr
24690       For Each doc In .Documents
24700         With doc
24710           If GetHiddenAttribute(acModule, .Name) = True Then
24720             lngHids = lngHids + 1&
24730             lngE = lngHids - 1&
24740             ReDim Preserve arr_varHid(H_ELEMS, lngE)
24750             arr_varHid(H_TYP, lngE) = acMacro
24760             arr_varHid(H_IDNT, lngE) = CLng(0)
24770             arr_varHid(H_NAM, lngE) = .Name
24780             arr_varHid(H_HIDE, lngE) = CBool(True)
24790             arr_varHid(H_CONN, lngE) = Null
24800             lngMods = lngMods + 1&
                  'Debug.Print "'  " & .Name
24810           End If
24820         End With
24830         Set doc = Nothing
24840       Next
24850     End With
24860     Set cntr = Nothing
24870     .Close
24880   End With  ' ** dbs.
24890   Set dbs = Nothing

24900   blnSkip = True
24910   If blnSkip = False Then
24920     lngLastType = 0&: strLastType = vbNullString
24930     For lngX = 0& To (lngHids - 1&)
24940       If arr_varHid(H_TYP, lngX) <> lngLastType Then
24950         Select Case arr_varHid(H_TYP, lngX)
              Case acTable
24960           lngLastType = acTable
24970           strLastType = "'HIDDEN TABLES:"
24980         Case acQuery
24990           lngLastType = acQuery
25000           strLastType = "'HIDDEN QUERIES:"
25010         Case acForm
25020           lngLastType = acForm
25030           strLastType = "'HIDDEN FORMS:"
25040         Case acMacro
25050           lngLastType = acMacro
25060           strLastType = "'HIDDEN MACROS:"
25070         Case acReport
25080           lngLastType = acReport
25090           strLastType = "'HIDDEN REPORTS:"
25100         Case acModule
25110           lngLastType = acModule
25120           strLastType = "'HIDDEN MODULES:"
25130         End Select
25140         Debug.Print strLastType
25150       End If
25160       Debug.Print "'  " & arr_varHid(H_NAM, lngX)
25170       If (lngX + 1&) Mod 100 = 0 Then
25180         Stop
25190       End If
25200     Next  ' ** lngX.
25210   End If  ' ** blnSkip.

25220   Set dbs = CurrentDb
25230   With dbs
25240     lngLastType = -1&: strLastType = vbNullString: strLastField = vbNullString
25250     lngDbsID = 0&: strLastDbsName = vbNullString: strThisDbsName = vbNullString
25260     For lngX = 0& To (lngHids - 1&)
25270       If arr_varHid(H_TYP, lngX) <> lngLastType Then
25280         If strLastType <> vbNullString Then
25290           rst.Close
25300           Set rst = Nothing
25310         End If
25320         lngLastType = arr_varHid(H_TYP, lngX)
25330         Select Case lngLastType
              Case acTable
25340           strLastType = "tblDatabase_Table"
25350           strLastField = "tbl_name"
25360         Case acQuery
25370           strLastType = "tblQuery"
25380           strLastField = "qry_name"
25390         Case acForm
25400           strLastType = "tblForm"
25410           strLastField = "frm_name"
25420         Case acMacro
25430           strLastType = "tblMacro"
25440           strLastField = "mcrname"
25450         Case acReport
25460           strLastType = "tblReport"
25470           strLastField = "rpt_name"
25480         Case acModule
25490           strLastType = "tblVBComponent"
25500           strLastField = "vbcom_name"
25510         End Select
25520         Set rst = .OpenRecordset(strLastType, dbOpenDynaset, dbConsistent)
25530         rst.MoveFirst
25540       End If
25550       With rst
25560         If lngLastType = acTable Then
25570           Select Case IsNull(arr_varHid(H_CONN, lngX))
                Case True
25580             strThisDbsName = CurrentAppName  ' ** Module Function: modFileUtilities.
25590           Case False
25600             strThisDbsName = Parse_File(arr_varHid(H_CONN, lngX))  ' ** Module Function: modFileUtilities.
25610           End Select
25620           If strThisDbsName <> strLastDbsName Then
25630             varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strThisDbsName & "'")
25640             If IsNull(varTmp00) = True Then
                    ' ** tblDatabase and/or tblTemplate_Database haven't been reset!
25650               strTmp01 = Parse_Ext(strThisDbsName)  ' ** Module Function: modFileUtilities.
25660               Select Case strTmp01
                    Case "mde"
                      ' ** Update tblDatabase, for '.mdb' to '.mde'.
25670                 Set qdf = dbs.QueryDefs("qryVersion_Convert_09a")
25680                 qdf.Execute
                      ' ** Update tblTemplate_Database, for '.mdb' to '.mde'.
25690                 Set qdf = dbs.QueryDefs("qryVersion_Convert_10a")
25700                 qdf.Execute
25710                 DoEvents
25720                 varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strThisDbsName & "'")
25730               Case "mdb"
                      ' ** Update tblDatabase, for '.mde' to '.mdb'.
25740                 Set qdf = dbs.QueryDefs("qryVersion_Convert_09b")
25750                 qdf.Execute
                      ' ** Update tblTemplate_Database, for '.mde' to '.mdb'.
25760                 Set qdf = dbs.QueryDefs("qryVersion_Convert_10b")
25770                 qdf.Execute
25780                 varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strThisDbsName & "'")
25790               End Select
25800             End If
25810             lngDbsID = varTmp00
25820             strLastDbsName = strThisDbsName
25830           End If
25840           .FindFirst "[" & strLastField & "] = '" & arr_varHid(H_NAM, lngX) & "' And [dbs_id] = " & CStr(lngDbsID)
25850         Else
25860           .FindFirst "[" & strLastField & "] = '" & arr_varHid(H_NAM, lngX) & "'"
25870         End If
25880         If .NoMatch = False Then
25890           If ![sec_hidden] <> True Then  'table,query,form,macro,report,module
25900             .Edit
25910             ![sec_hidden] = True
25920             .Fields(Left(strLastField, (InStr(strLastField, "_") - 1)) & "_datemodified") = Now()
25930             .Update
25940           End If
25950         Else
25960           Stop
25970         End If
25980       End With
25990     Next  ' ** lngX.
26000     rst.Close
26010     .Close
26020   End With

26030   Beep

EXITP:
26040   Set rst = Nothing
26050   Set doc = Nothing
26060   Set cntr = Nothing
26070   Set tdf = Nothing
26080   Set qdf = Nothing
26090   Set dbs = Nothing
26100   Security_Hidden_List = blnRetVal
26110   Exit Function

ERRH:
26120   blnRetVal = False
26130   Select Case ERR.Number
        Case Else
26140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26150   End Select
26160   Resume EXITP

End Function

Public Function Security_List() As Boolean

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "Security_List"

        Dim wrk As DAO.Workspace, grp As DAO.Group, usr As DAO.User
        Dim lngGrps As Long, arr_varGrp() As Variant
        Dim lngUsrs As Long, arr_varUsr() As Variant
        Dim blnFound As Boolean
        Dim arr_varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varGrp().
        Const G_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const G_GNAM As Integer = 0
        Const G_USRS As Integer = 1

        ' ** Array: arr_varUsr().
        Const U_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const U_UNAM As Integer = 0
        Const U_GRPS As Integer = 1

26210   blnRetVal = True

        'Group       User
        '==========  =========
        'Admins
        '            admin
        '            superuser
        'DataEntry
        '            Mary
        'Users
        '            admin
        '            superuser
        '            Mary
        '            John
        'ViewOnly
        '            John

26220   lngGrps = 0&
26230   ReDim arr_varGrp(G_ELEMS, 0)

26240   lngUsrs = 0&
26250   ReDim arr_varUsrs(U_ELEMS, 0)

26260   Set wrk = DBEngine.Workspaces(0)
26270   With wrk
26280     lngGrps = .Groups.Count
26290     ReDim arr_varGrp(G_ELEMS, (lngGrps - 1&))
26300     For lngX = 0& To (lngGrps - 1&)
26310       Set grp = .Groups(lngX)
26320       With grp
26330         arr_varGrp(G_GNAM, lngX) = .Name
26340         arr_varGrp(G_USRS, lngX) = .Users.Count
26350         For Each usr In .Users
26360           With usr
26370             blnFound = False
26380             For lngY = 0& To (lngUsrs - 1&)
26390               If arr_varUsr(U_UNAM, lngY) = .Name Then
26400                 blnFound = True
26410                 arr_varUsr(U_GRPS, lngY) = arr_varUsr(U_GRPS, lngY) & ";" & CStr(lngX)
26420                 Exit For
26430               End If
26440             Next
26450             If blnFound = False Then
26460               lngUsrs = lngUsrs + 1&
26470               lngE = lngUsrs - 1&
26480               ReDim Preserve arr_varUsr(U_ELEMS, lngE)
26490               arr_varUsr(U_UNAM, lngE) = .Name
26500               arr_varUsr(U_GRPS, lngE) = CStr(lngX)
26510             End If
26520           End With
26530         Next
26540       End With  ' ** This Group: grp.
26550     Next  ' ** For each Group: lngX.
26560     .Close
26570   End With  ' ** This Workspace: wrk.

26580   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

26590   If lngGrps > 0& Then
26600     Debug.Print "'Group       User"
26610     Debug.Print "'==========  ========="
26620     For lngX = 0& To (lngGrps - 1&)
26630       Debug.Print "'" & arr_varGrp(G_GNAM, lngX)
26640       For lngY = 0& To (lngUsrs - 1&)
26650         If InStr(arr_varUsr(U_GRPS, lngY), ";") > 0 Then
26660           arr_varTmp00 = Split(arr_varUsr(U_GRPS, lngY), ";")
26670           For lngZ = 0& To UBound(arr_varTmp00)
26680             If CLng(arr_varTmp00(lngZ)) = lngX Then
26690               Debug.Print "'            " & arr_varUsr(U_UNAM, lngY)
26700               Exit For
26710             End If
26720           Next
26730         Else
26740           If CLng(arr_varUsr(U_GRPS, lngY)) = lngX Then
26750             Debug.Print "'            " & arr_varUsr(U_UNAM, lngY)
26760           End If
26770         End If
26780       Next
26790     Next
26800   Else
26810     Beep
26820     Debug.Print "'NO GROUPS?!"
26830   End If

EXITP:
26840   Set usr = Nothing
26850   Set grp = Nothing
26860   Set wrk = Nothing
26870   Security_List = blnRetVal
26880   Exit Function

ERRH:
26890   Select Case ERR.Number
        Case Else
26900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26910   End Select
26920   Resume EXITP

End Function

Public Function AutoCorrect_Get(strOption As String) As Variant
' ** Return Registry values within the Access 2000 Common section.

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "AutoCorrect_Get"

        Dim arr_varRetVal As Variant
        Dim varRetVal As Variant

        Const KEY_AUTO As String = "Software\Microsoft\Office\9.0\Common\AutoCorrect"

        ' *****************************************************************************
        ' ** AutoCorrect Options:
        ' **   Key Name (Option)            Tools->Options Entry
        ' **   ===========================  =========================================
        ' **   CorrectTwoInitialCapitals    Correct two initial capitals
        ' **   CapitalizeSentence           Capitalize first letter of sentence
        ' **   CapitalizeNamesOfDays        Capitalize names of days
        ' **   ToggleCapsLock               Correct accidental use of CAPS LOCK key
        ' **   ReplaceText                  Replace text as you type
        ' *****************************************************************************

27010   arr_varRetVal = QueryValue(HKEY_CURRENT_USER, KEY_AUTO, strOption)  ' ** Function: Below.
        ' **************************************************************
        ' ** Array: arr_arr_varRetVal()
        ' **
        ' **   Element  Description              Type       Constant
        ' **   =======  =======================  =========  ==========
        ' **      0     Error Code               Long       RES_ERR
        ' **      1     Exists                   Boolean    RES_YN
        ' **      2     SubKey Count             Long       RES_SUB
        ' **      3     Longest SubKey Length    Long       RES_LNG
        ' **      4     Value                    Variant    RES_VAR
        ' **
        ' **************************************************************

27020   If arr_varRetVal(RES_ERR, 0) = ERROR_SUCCESS Then
27030     varRetVal = CBool(arr_varRetVal(RES_VAR, 0))
27040   Else
27050     varRetVal = "#ERROR " & CStr(arr_varRetVal(RES_ERR, 0))
27060   End If

EXITP:
27070   AutoCorrect_Get = varRetVal
27080   Exit Function

ERRH:
27090   varRetVal = "#ERROR " & CStr(ERR.Number)
27100   Select Case ERR.Number
        Case Else
27110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27120   End Select
27130   Resume EXITP

End Function

Public Function AutoCorrect_Set(strOption As String, blnSetting As Boolean) As Variant
' ** Set Registry values within the Access 2000 Common section.

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "AutoCorrect_Set"

        Dim arr_varRetVal As Variant
        Dim varValue As Variant
        Dim varRetVal As Variant

        Const KEY_AUTO As String = "Software\Microsoft\Office\9.0\Common\AutoCorrect"

        ' *****************************************************************************
        ' ** AutoCorrect Options:
        ' **   Key Name (Option)            Tools->Options Entry
        ' **   ===========================  =========================================
        ' **   CorrectTwoInitialCapitals    Correct two initial capitals
        ' **   CapitalizeSentence           Capitalize first letter of sentence
        ' **   CapitalizeNamesOfDays        Capitalize names of days
        ' **   ToggleCapsLock               Correct accidental use of CAPS LOCK key
        ' **   ReplaceText                  Replace text as you type
        ' *****************************************************************************

        ' ** In the Registry, True = 1 and False = 0.
27210   If blnSetting = False Then varValue = 0 Else varValue = 1

27220   arr_varRetVal = SetKeyValue(HKEY_CURRENT_USER, KEY_AUTO, strOption, varValue, REG_DWORD)  ' ** Function: Below.
        ' **************************************************************
        ' ** Array: arr_arr_varRetVal()
        ' **
        ' **   Element  Description              Type       Constant
        ' **   =======  =======================  =========  ==========
        ' **      0     Error Code               Long       RES_ERR
        ' **      1     Exists                   Boolean    RES_YN
        ' **      2     SubKey Count             Long       RES_SUB
        ' **      3     Longest SubKey Length    Long       RES_LNG
        ' **      4     Value                    Variant    RES_VAR
        ' **
        ' **************************************************************

27230   If arr_varRetVal(RES_ERR, 0) = ERROR_SUCCESS Then
27240     varRetVal = CBool(True)
27250   Else
27260     varRetVal = CBool(False)
27270   End If

EXITP:
27280   AutoCorrect_Set = varRetVal
27290   Exit Function

ERRH:
27300   varRetVal = CBool(False)
27310   Select Case ERR.Number
        Case Else
27320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27330   End Select
27340   Resume EXITP

End Function

Public Function QueryValue(lngPredefinedKey As Long, strKeyName As String, strValueName As String) As Variant
' ** Description:
' **   This Function will return the data field of a value
' ** Syntax:
' **   Variable = QueryValue(Location, KeyName, ValueName)
' **   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS
' **   KeyName is the key that the value is under (example: "Software\Microsoft\Windows\CurrentVersion\Explorer")
' **   ValueName is the name of the value you want to access (example: "link")
' ** Return: Array

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "QueryValue"

        Dim lngResult As Long    ' ** Result of the API functions.
        Dim lngHKey As Long      ' ** Handle of opened key.
        Dim varValue As Variant  ' ** Setting of queried value.
        Dim blnExists As Boolean
        Dim arr_varRetVal() As Variant

27410   blnExists = False

27420   lngResult = RegOpenKeyEx(lngPredefinedKey, strKeyName, 0, KEY_QUERY_VALUE, lngHKey)  ' ** API Function: Above.

27430   lngResult = QueryValueEx(lngHKey, strValueName, varValue)  ' ** Function: Below.

27440   If lngResult = ERROR_SUCCESS Then blnExists = True

27450   ReDim arr_varRetVal(RES_ELEMS, 0)
27460   arr_varRetVal(RES_ERR, 0) = lngResult
27470   arr_varRetVal(RES_YN, 0) = blnExists
27480   arr_varRetVal(RES_SUB, 0) = Null
27490   arr_varRetVal(RES_LNG, 0) = Null
27500   arr_varRetVal(RES_VAR, 0) = varValue
        ' **************************************************************
        ' ** Array: arr_varRetVal()
        ' **
        ' **   Element  Description              Type       Constant
        ' **   =======  =======================  =========  ==========
        ' **      0     Error Code               Long       RES_ERR
        ' **      1     Exists                   Boolean    RES_YN
        ' **      2     SubKey Count             Long       RES_SUB
        ' **      3     Longest SubKey Length    Long       RES_LNG
        ' **      4     Value                    Variant    RES_LNG
        ' **
        ' **************************************************************

27510   RegCloseKey (lngHKey) ' ** Must close the key when finished.

EXITP:
27520   QueryValue = arr_varRetVal
27530   Exit Function

ERRH:
27540   Select Case ERR.Number
        Case Else
27550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27560   End Select
27570   Resume EXITP

End Function

Public Function QueryValueEx(ByVal lngHKeyX As Long, ByVal strValueName As String, varValue As Variant) As Long
' ** Determine the size and type of data to be read.

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "QueryValueEx"

        Dim lngResult As Long
        Dim lngHKey As Long
        Dim lngType As Long
        Dim lngValue As Long
        Dim strValue As String

27610   lngResult = RegQueryValueExNULL(lngHKeyX, strValueName, 0&, lngType, 0&, lngHKey)  ' ** API Function: Above.
27620   If lngResult <> ERROR_SUCCESS Then Error 5

        'RegGetValue()
        'Note for Visual Basic:
        'http://guidepc.altervista.org/windows/api/functions/regqueryvalueex.html
        'For XP SP2, where RegGetValue is not available, use SHRegGetValue in shlwapi.dll in place of RegGetValue.
        'LONG WINAPI RegGetValue(
        '  __in         lngHKey lngHKey,
        '  __in_opt     LPCTSTR lpSubKey,
        '  __in_opt     LPCTSTR lpValue,
        '  __in_opt     DWORD dwFlags,
        '  __out_opt    LPDWORD pdwType,
        '  __out_opt    PVOID pvData,
        '  __inout_opt  LPDWORD pcbData
        ');

        ' ** RegQueryValueEx retrieves the type, content and data for a specified value name.
        ' ** Note that if you declare the lpData parameter as String, you must pass it By Value.
        'RegQueryValueEx(ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long)

27630   Select Case lngType
        Case REG_SZ
          ' ** For strings.
27640     strValue = String(lngHKey, 0)
27650     lngResult = RegQueryValueExString(lngHKeyX, strValueName, 0&, lngType, strValue, lngHKey)  ' ** API Function: Above.
27660     If lngResult = ERROR_SUCCESS Then
27670       varValue = Left(strValue, lngHKey)
27680     Else
27690       varValue = Empty
27700     End If
27710   Case REG_DWORD
          ' ** For DWORDS.
27720     lngResult = RegQueryValueExLong(lngHKeyX, strValueName, 0&, lngType, lngValue, lngHKey)  ' ** API Function: Above.
27730     If lngResult = ERROR_SUCCESS Then varValue = lngValue
27740   Case Else
          ' ** All other data types not supported.
27750     lngResult = -1
27760   End Select

EXITP:
27770   QueryValueEx = lngResult
27780   Exit Function

ERRH:
27790   Select Case ERR.Number
        Case Else
27800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27810   End Select
27820   Resume EXITP

End Function

Private Function SetKeyValue(lngPredefinedKey As Long, strKeyName As String, strValueName As String, varValue As Variant, lngValueType As Long) As Variant
' ** Description:
' **   This Function will set the data field of a value
' ** Syntax:
' **   QueryValue Location, KeyName, ValueName, ValueSetting, ValueType
' **   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS
' **   KeyName is the key that the value is under (example: "Key1\SubKey1")
' **   ValueName is the name of the value you want create, or set the value of (example: "ValueTest")
' **   ValueSetting is what you want the value to equal
' **   ValueType must equal either REG_SZ (a string) Or REG_DWORD (an integer)

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "SetKeyValue"

        Dim lngResult As Long    'result of the API functions
        Dim hKey As Long         'handle of open key
        Dim blnExists As Boolean
        Dim arr_varRetVal() As Variant

27910   blnExists = False

27920   lngResult = RegOpenKeyEx(lngPredefinedKey, strKeyName, 0, KEY_ALL_ACCESS, hKey)  ' ** API Function: Above.

27930   lngResult = SetValueEx(hKey, strValueName, lngValueType, varValue)  ' ** Function: Below.

27940   If lngResult = ERROR_SUCCESS Then blnExists = True

27950   ReDim arr_varRetVal(RES_ELEMS, 0)
27960   arr_varRetVal(RES_ERR, 0) = lngResult
27970   arr_varRetVal(RES_YN, 0) = blnExists
27980   arr_varRetVal(RES_SUB, 0) = Null
27990   arr_varRetVal(RES_LNG, 0) = Null
28000   arr_varRetVal(RES_VAR, 0) = varValue
        ' **************************************************************
        ' ** Array: arr_varRetVal()
        ' **
        ' **   Element  Description              Type       Constant
        ' **   =======  =======================  =========  ==========
        ' **      0     Error Code               Long       RES_ERR
        ' **      1     Exists                   Boolean    RES_YN
        ' **      2     SubKey Count             Long       RES_SUB
        ' **      3     Longest SubKey Length    Long       RES_LNG
        ' **      4     Value                    Variant    RES_LNG
        ' **
        ' **************************************************************

28010   RegCloseKey (hKey)

EXITP:
28020   SetKeyValue = arr_varRetVal
28030   Exit Function

ERRH:
28040   Select Case ERR.Number
        Case Else
28050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28060   End Select
28070   Resume EXITP

End Function

Private Function SetValueEx(ByVal lngHKey As Long, strValueName As String, lngType As Long, varValue As Variant) As Long

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "SetValueEx"

        Dim lngValue As Long
        Dim strValue As String
        Dim lngRetVal As Long

28110   Select Case lngType
        Case REG_SZ
28120     strValue = varValue
28130     lngRetVal = RegSetValueExString(lngHKey, strValueName, 0&, lngType, strValue, Len(strValue))  ' ** API Function: Above.
28140   Case REG_DWORD
28150     lngValue = varValue
28160     lngRetVal = RegSetValueExLong(lngHKey, strValueName, 0&, lngType, lngValue, 4)  ' ** API Function: Above.
28170   End Select

EXITP:
28180   SetValueEx = lngRetVal
28190   Exit Function

ERRH:
28200   Select Case ERR.Number
        Case Else
28210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28220   End Select
28230   Resume EXITP

End Function

Public Function DemoLicense_Set() As Boolean
' ** Save the demo license acceptance and timestamp.

28300 On Error GoTo ERRH

        Const THIS_PROC As String = "DemoLicense_Set"

        Dim wrk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database, rst As DAO.Recordset, tdf As DAO.TableDef
        Dim strTimestamp As String, strCaption As String, strKey As String
        Dim strPiece1 As String, strPiece2 As String, strPiece3 As String, strPiece4 As String, strPiece5 As String, strPiece6 As String
        Dim intWrkType As Integer
        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim intX As Integer
        Dim blnRetVal As Boolean

        'When this first opens, it finds nothing in this 'cookie'.
        'Data and Archive also find nothing in their 'cookies'.
        'If frontends get mixed with backends, they won't match.
        'Each of the 3 pieces has a hidden table.
        'A code goes into vp_DE1, vd_DE1, and va_DE1.
        'vp_DE1 is used with the Archive hidden data,
        'va_DE1 is used with the Data hidden data,
        'and vd_DE1 is used with this hidden data.
        'If they aren't coordinated, the EULA comes up again.

        'Here's the Timestamp: 39907.7957638889  [CDbl(Now()]
        'Here's the Caption:   "I &Accept"

28310   blnRetVal = False

28320   If gdatAccept <> 1 And gstrAccept <> vbNullString Then

          ' ** gstrTrustDataLocation needs to be set earlier than normal.
28330     If IniFile_GetDataLoc = True Then  ' ** Module Procedure: modStartupFuncs.
28340       If Dir(gstrTrustDataLocation, vbDirectory) <> vbNullString Then

              ' ** 21 Characters.
              ' **   39907.795763888900000       (5.15)  21
28350         strTimestamp = CStr(CDbl(gdatAccept))
28360         If InStr(strTimestamp, ".") = 0 Then strTimestamp = strTimestamp & "."
28370         strTimestamp = Left(strTimestamp & String(16, "0"), 21)

              ' ** Should build to 27 characters.
              ' **   073032038065099099101112116         27
28380         intLen = Len(gstrAccept)
28390         strCaption = vbNullString
28400         For intX = 1 To intLen
28410           strCaption = strCaption & Right("000" & CStr(Asc(Mid(gstrAccept, intX, 1))), 3)
                ' ** 073 032 038 065 099 099 101 112 116
                ' ** Asc("I") = 73
                ' ** Asc(" ") = 32
                ' ** Asc("&") = 38
                ' ** Asc("A") = 65
                ' ** Asc("c") = 99
                ' ** Asc("c") = 99
                ' ** Asc("e") = 101
                ' ** Asc("p") = 112
                ' ** Asc("t") = 116
28420         Next

              ' ** Should be 48 characters.
              ' ** 07303203806509909910111211639907.795763888900000
28430         strKey = strCaption & strTimestamp

28440         Set dbsLoc = CurrentDb

28450         intWrkType = 0
28460 On Error Resume Next
28470         Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
28480         If ERR.Number <> 0 Then
28490 On Error GoTo ERRH
28500 On Error Resume Next
28510           Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
28520           If ERR.Number <> 0 Then
28530 On Error GoTo ERRH
28540 On Error Resume Next
28550             Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
28560             If ERR.Number <> 0 Then
28570 On Error GoTo ERRH
28580 On Error Resume Next
28590               Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
28600               If ERR.Number <> 0 Then
28610 On Error GoTo ERRH
28620 On Error Resume Next
28630                 Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
28640                 If ERR.Number <> 0 Then
28650 On Error GoTo ERRH
28660 On Error Resume Next
28670                   Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
28680                   If ERR.Number <> 0 Then
28690 On Error GoTo ERRH
28700 On Error Resume Next
28710                     Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
28720 On Error GoTo ERRH
28730                     intWrkType = 7
28740                   Else
28750 On Error GoTo ERRH
28760                     intWrkType = 6
28770                   End If
28780                 Else
28790 On Error GoTo ERRH
28800                   intWrkType = 5
28810                 End If
28820               Else
28830 On Error GoTo ERRH
28840                 intWrkType = 4
28850               End If
28860             Else
28870 On Error GoTo ERRH
28880               intWrkType = 3
28890             End If
28900           Else
28910 On Error GoTo ERRH
28920             intWrkType = 2
28930           End If
28940         Else
28950 On Error GoTo ERRH
28960           intWrkType = 1
28970         End If

28980         With wrk

                ' ** gstrTrustDataLocation:
                ' **   C:\VictorGCS_Clients\TrustAccountant\NewWorking\DemoDatabase\  '## OK
                ' ** With the final backslash (as gstrTrustDataLocation has),
                ' ** returns '.', indicating 'This Directory'. Without the final
                ' ** backslash, returns 'DemoDatabase', the directory name.
                ' ** Old DOS conventions: One dot, '.', 'This Directory',
                ' ** two dots, '..', 'The Directory One Level Above'
                ' ** (which can be concatinated, e.g., '..\..',
                ' ** 2 levels above, '..\..\..', 3 levels, etc.).

                ' ** Each gets 56 characters: 1 whole string plus 1/6th of the whole.
                ' ** So each whole has to match the assembled whole.
                ' ** I think I like that!
                ' ** (Of course, this isn't sophisticated at all! But it is meant
                ' ** to look like a code, and it does have all the info in it.)

                ' ** The hidden table gets the whole key, plus its piece,
                ' ** the visible table gets only its piece.
                ' **  _~rmcp  vgc_1  1  07303203806509909910111211639907.79576388890000007303203
                ' **  _~rmca  vgc_2  3  07303203806509909910111211639907.79576388890000099101112
                ' **  _~rmcd  vgc_3  5  07303203806509909910111211639907.795763888900000.7957638
                ' **  m_VP  vp_DE1   6  88900000
                ' **  m_VA  va_DE1   2  80650990
                ' **  m_VD  vd_DE1   4  11639907

                ' ** 07303203 80650990 99101112 11639907 .7957638 88900000
28990           strPiece1 = Left(strKey, 8)
29000           strPiece2 = Mid(strKey, 9, 8)
29010           strPiece3 = Mid(strKey, 17, 8)
29020           strPiece4 = Mid(strKey, 25, 8)
29030           strPiece5 = Mid(strKey, 33, 8)
29040           strPiece6 = Mid(strKey, 41, 8)
29050           blnFound = False

                ' ** _~rmcp  vgc_1  1  07303203806509909910111211639907.79576388890000007303203
29060           Set rst = dbsLoc.OpenRecordset("_~rmcp", dbOpenDynaset, dbConsistent)
29070           With rst
29080             If .BOF = True And .EOF = True Then
29090               .AddNew
29100             Else
29110               .Edit
29120             End If
29130             ![vgc_1] = strKey & strPiece1
29140             .Update
29150             .Close
29160           End With
                ' ** m_VP  vp_DE1   6  88900000
29170           Set rst = dbsLoc.OpenRecordset("m_VP", dbOpenDynaset, dbConsistent)
29180           With rst
29190             .Edit
29200             ![vp_DE1] = strPiece6
29210             .Update
29220             .Close
29230           End With

                ' ** gstrFile_ArchDataName : TrstArch.mdb
29240           Set dbsLnk = .OpenDatabase(gstrTrustDataLocation & gstrFile_ArchDataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
29250           With dbsLnk
                  ' ** Check if the hidden table is there.
29260             blnFound = False
29270             For Each tdf In .TableDefs
29280               With tdf
29290                 If .Name = "_~rmca" Then
29300                   blnFound = True
29310                   Exit For
29320                 End If
29330               End With
29340             Next
29350             If blnFound = True Then
                    ' ** _~rmca  vgc_2  3  07303203806509909910111211639907.79576388890000099101112
29360               Set rst = .OpenRecordset("_~rmca", dbOpenDynaset, dbConsistent)
29370               With rst
29380                 If .BOF = True And .EOF = True Then
29390                   .AddNew
29400                 Else
29410                   .Edit
29420                 End If
29430                 ![vgc_2] = strKey & strPiece3
29440                 .Update
29450                 .Close
29460               End With
                    ' ** m_VA  va_DE1   2  80650990
29470               Set rst = .OpenRecordset("m_VA", dbOpenDynaset, dbConsistent)
29480               With rst
29490                 .Edit
29500                 ![va_DE1] = strPiece2
29510                 .Update
29520                 .Close
29530               End With
29540             End If
29550             .Close
29560           End With

29570           If blnFound = True Then
                  ' ** gstrFile_DataName     : TrustDta.mdb
29580             Set dbsLnk = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
29590             With dbsLnk
                    ' ** Check if the hidden table is there.
29600               blnFound = False
29610               For Each tdf In .TableDefs
29620                 With tdf
29630                   If .Name = "_~rmcd" Then
29640                     blnFound = True
29650                     Exit For
29660                   End If
29670                 End With
29680               Next
29690               If blnFound = True Then
                      ' ** _~rmcd  vgc_3  5  07303203806509909910111211639907.795763888900000.7957638
29700                 Set rst = .OpenRecordset("_~rmcd", dbOpenDynaset, dbConsistent)
29710                 With rst
29720                   If .BOF = True And .EOF = True Then
29730                     .AddNew
29740                   Else
29750                     .Edit
29760                   End If
29770                   ![vgc_3] = strKey & strPiece5
29780                   .Update
29790                   .Close
29800                 End With
                      ' ** m_VD  vd_DE1   4  11639907
29810                 Set rst = .OpenRecordset("m_VD", dbOpenDynaset, dbConsistent)
29820                 With rst
29830                   .Edit
29840                   ![vd_DE1] = strPiece4
29850                   .Update
29860                   .Close
29870                 End With
29880               End If
29890               .Close
29900             End With

29910             If blnFound = True Then
29920               blnRetVal = True
29930             End If

29940           End If

29950           .Close
29960         End With
29970         dbsLoc.Close

29980         If blnFound = False Then
29990           Beep
30000           MsgBox "The data files for this Trust Accountant Demo are invalid." & vbCrLf & vbCrLf & _
                  "Please contact Delta Data, Inc.", vbCritical + vbOKOnly, "Invalid Data Files"
30010         End If

30020       Else
              ' ** gstrTrustDataLocation directory doesn't exist!
30030         Beep
30040         MsgBox "The data directory could not be found." & vbCrLf & "  " & gstrTrustDataLocation & vbCrLf & vbCrLf & _
                "Please contact Delta Data, Inc.", vbCritical + vbOKOnly, "Directory Not Found"
30050       End If
30060     Else
            ' ** Unable to get gstrTrustDataLocation.
30070       Beep
30080       MsgBox "There was a problem retrieving your Trust Accountant data location." & vbCrLf & vbCrLf & _
              "Please contact Delta Data, Inc.", vbCritical + vbOKOnly, "Problem With DDTrust.ini"
30090     End If  ' ** IniFile_GetDataLoc().

30100   End If

EXITP:
30110   Set tdf = Nothing
30120   Set rst = Nothing
30130   Set dbsLoc = Nothing
30140   Set dbsLnk = Nothing
30150   Set wrk = Nothing
30160   DemoLicense_Set = blnRetVal
30170   Exit Function

ERRH:
30180   blnRetVal = False
30190   Select Case ERR.Number
        Case Else
30200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30210   End Select
30220   Resume EXITP

End Function

Public Function DemoLicense_Get() As String
' ** Retrieve the status of the demo license.

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "DemoLicense_Get"

        Dim wrk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database, rst As DAO.Recordset
        Dim tdf As DAO.TableDef
        Dim strTimestamp As String, strCaption As String, strKey As String
        Dim strPiece1 As String, strPiece2 As String, strPiece3 As String, strPiece4 As String, strPiece5 As String, strPiece6 As String
        Dim strDB1 As String, strDB2 As String, strDB3 As String
        Dim strRet1 As String, strRet2 As String, strRet3 As String, strRet4 As String, strRet5 As String, strRet6 As String
        Dim intWrkType As Integer
        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String, datTmp02 As Date
        Dim intX As Integer
        Dim strRetVal As String

30310   strRetVal = vbNullString

30320   strTimestamp = vbNullString: strCaption = vbNullString: strKey = vbNullString
30330   strPiece1 = vbNullString: strPiece2 = vbNullString: strPiece3 = vbNullString
30340   strPiece4 = vbNullString: strPiece5 = vbNullString: strPiece6 = vbNullString
30350   strDB1 = vbNullString: strDB2 = vbNullString: strDB3 = vbNullString

30360   If IniFile_GetDataLoc = True Then  ' ** Module Procedure: modStartupFuncs.
30370     If Dir(gstrTrustDataLocation, vbDirectory) <> vbNullString Then

30380       Set dbsLoc = CurrentDb

30390       intWrkType = 0
30400 On Error Resume Next
30410       Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
30420       If ERR.Number <> 0 Then
30430 On Error GoTo ERRH
30440 On Error Resume Next
30450         Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
30460         If ERR.Number <> 0 Then
30470 On Error GoTo ERRH
30480 On Error Resume Next
30490           Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
30500           If ERR.Number <> 0 Then
30510 On Error GoTo ERRH
30520 On Error Resume Next
30530             Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
30540             If ERR.Number <> 0 Then
30550 On Error GoTo ERRH
30560 On Error Resume Next
30570               Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
30580               If ERR.Number <> 0 Then
30590 On Error GoTo ERRH
30600 On Error Resume Next
30610                 Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
30620                 If ERR.Number <> 0 Then
30630 On Error GoTo ERRH
30640 On Error Resume Next
30650                   Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
30660 On Error GoTo ERRH
30670                   intWrkType = 7
30680                 Else
30690 On Error GoTo ERRH
30700                   intWrkType = 6
30710                 End If
30720               Else
30730 On Error GoTo ERRH
30740                 intWrkType = 5
30750               End If
30760             Else
30770 On Error GoTo ERRH
30780               intWrkType = 4
30790             End If
30800           Else
30810 On Error GoTo ERRH
30820             intWrkType = 3
30830           End If
30840         Else
30850 On Error GoTo ERRH
30860           intWrkType = 2
30870         End If
30880       Else
30890 On Error GoTo ERRH
30900         intWrkType = 1
30910       End If

30920       With wrk

              ' ** _~rmcp  vgc_1  1  07303203806509909910111211639907.79576388890000007303203
30930         Set rst = dbsLoc.OpenRecordset("_~rmcp", dbOpenDynaset, dbConsistent)
30940         With rst
30950           If .BOF = True And .EOF = True Then
30960             strRetVal = "#EMPTY"
30970             strDB1 = "#EMPTY"
30980           Else
30990             .MoveFirst
31000             strRet1 = ![vgc_1]
31010           End If
31020           .Close
31030         End With

31040         If strRetVal = vbNullString Then

                ' ** m_VP  vp_DE1   6  88900000
31050           Set rst = dbsLoc.OpenRecordset("m_VP", dbOpenDynaset, dbConsistent)
31060           With rst
31070             .MoveFirst
31080             strRet2 = ![vp_DE1]
31090             .Close
31100           End With

                ' ** gstrFile_ArchDataName : TrstArch.mdb
31110           Set dbsLnk = .OpenDatabase(gstrTrustDataLocation & gstrFile_ArchDataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
31120           With dbsLnk
                  ' ** Check if the hidden table is there.
31130             blnFound = False
31140             For Each tdf In .TableDefs
31150               With tdf
31160                 If .Name = "_~rmca" Then
31170                   blnFound = True
31180                   Exit For
31190                 End If
31200               End With
31210             Next
31220             If blnFound = True Then
                    ' ** _~rmca  vgc_2  3  07303203806509909910111211639907.79576388890000099101112
31230               Set rst = .OpenRecordset("_~rmca", dbOpenDynaset, dbConsistent)
31240               With rst
31250                 If .BOF = True And .EOF = True Then
31260                   strRetVal = "#EMPTY"
31270                   strDB2 = "#EMPTY"
31280                 Else
31290                   .MoveFirst
31300                   strRet3 = ![vgc_2]
31310                 End If
31320                 .Close
31330               End With
31340             Else
31350               strRetVal = "#MISSING"
31360               strDB2 = "#MISSING"
31370             End If
31380           End With
31390           If strRetVal <> vbNullString Then
31400             dbsLnk.Close
31410           End If

31420         End If

31430         If strRetVal = vbNullString Then

31440           With dbsLnk
                  ' ** m_VA  va_DE1   2  80650990
31450             Set rst = .OpenRecordset("m_VA", dbOpenDynaset, dbConsistent)
31460             With rst
31470               .MoveFirst
31480               strRet4 = ![va_DE1]
31490               .Close
31500             End With
31510             .Close
31520           End With

                ' ** gstrFile_DataName     : TrustDta.mdb
31530           Set dbsLnk = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
31540           With dbsLnk
                  ' ** Check if the hidden table is there.
31550             blnFound = False
31560             For Each tdf In .TableDefs
31570               With tdf
31580                 If .Name = "_~rmcd" Then
31590                   blnFound = True
31600                   Exit For
31610                 End If
31620               End With
31630             Next
31640             If blnFound = True Then
                    ' ** _~rmcd  vgc_3  5  07303203806509909910111211639907.795763888900000.7957638
31650               Set rst = .OpenRecordset("_~rmcd", dbOpenDynaset, dbConsistent)
31660               With rst
31670                 If .BOF = True And .EOF = True Then
31680                   strRetVal = "#EMPTY"
31690                   strDB3 = "#EMPTY"
31700                 Else
31710                   .MoveFirst
31720                   strRet5 = ![vgc_3]
31730                 End If
31740                 .Close
31750               End With
31760             Else
31770               strRetVal = "#MISSING"
31780               strDB2 = "#MISSING"
31790             End If
31800           End With
31810           If strRetVal <> vbNullString Then
31820             dbsLnk.Close
31830           End If

31840         End If

31850         If strRetVal = vbNullString Then

31860           With dbsLnk
                  ' ** m_VD  vd_DE1   4  11639907
31870             Set rst = .OpenRecordset("m_VD", dbOpenDynaset, dbConsistent)
31880             With rst
31890               .MoveFirst
31900               strRet6 = ![vd_DE1]
31910               .Close
31920             End With
31930             .Close
31940           End With

31950         End If

31960         .Close
31970       End With
31980       dbsLoc.Close

31990       If strRetVal = vbNullString Then
              ' ** Now assemble the pieces and see if they match!

              ' ** Caption:
              ' ** Should build to 27 characters.
              ' **   073032038065099099101112116
              ' ** strCaption = strCaption & Right("000" & CStr(Asc(Mid(gstrAccept, intX, 1))), 3)
              ' ** 073 032 038 065 099 099 101 112 116
              ' ** Asc("I") = 73, Asc(" ") = 32, Asc("&") = 38, Asc("A") = 65, Asc("c") = 99,
              ' ** Asc("c") = 99, Asc("e") = 101, Asc("p") = 112, Asc("t") = 116

              ' ** Timestamp:
              ' ** 21 Characters.
              ' **   39907.795763888900000
              ' ** strTimestamp = CStr(CDbl(gdatAccept))
              ' ** strTimestamp = Left(strTimestamp & String(16, "0"), 21)

              ' ** Should be 48 characters.
              ' ** 07303203806509909910111211639907.795763888900000
              ' ** strKey = strCaption & strTimestamp

              ' ** 07303203 80650990 99101112 11639907 .7957638 88900000
              ' ** strPiece1 = Left(strKey, 8)
              ' ** strPiece2 = Mid(strKey, 9, 8)
              ' ** strPiece3 = Mid(strKey, 17, 8)
              ' ** strPiece4 = Mid(strKey, 25, 8)
              ' ** strPiece5 = Mid(strKey, 33, 8)
              ' ** strPiece6 = Mid(strKey, 41, 8)

              ' ** strRet1 :  _~rmcp  vgc_1  1  07303203806509909910111211639907.79576388890000007303203
              ' ** strRet2 :  m_VP  vp_DE1   6  88900000
              ' ** strRet3 :  _~rmca  vgc_2  3  07303203806509909910111211639907.79576388890000099101112
              ' ** strRet4 :  m_VA  va_DE1   2  80650990
              ' ** strRet5 :  _~rmcd  vgc_3  5  07303203806509909910111211639907.795763888900000.7957638
              ' ** strRet6 :  m_VD  vd_DE1   4  11639907
32000         strKey = Left(strRet1, 48)  ' ** This Trust.mde's key.
32010         If strKey = Left(strRet3, 48) And strKey = Left(strRet5, 48) Then
32020           strPiece1 = Left(strKey, 8)
32030           strPiece2 = Mid(strKey, 9, 8)
32040           strPiece3 = Mid(strKey, 17, 8)
32050           strPiece4 = Mid(strKey, 25, 8)
32060           strPiece5 = Mid(strKey, 33, 8)
32070           strPiece6 = Mid(strKey, 41, 8)
32080           If strPiece1 = Right(strRet1, 8) And strPiece2 = strRet4 And _
                    strPiece3 = Right(strRet3, 8) And strPiece4 = strRet6 And _
                    strPiece5 = Right(strRet5, 8) And strPiece6 = strRet2 Then
                  ' ** All pieces match. Proceed!
32090             strCaption = Left(strKey, 27)
32100             strTimestamp = Mid(strKey, 28)  ' ** 21 characters.
32110 On Error Resume Next
                  ' ** CDate("39907.795763888900000") = 4/4/2009 7:05:54 PM
32120             datTmp02 = CDate(strTimestamp)
32130             If ERR.Number = 0 Then
32140 On Error GoTo ERRH
                    ' ** A-OK!
                    ' ** Now let them into the program, frmLicense first.
32150               intLen = Len(strCaption)
32160               strTmp01 = vbNullString
32170               For intX = 1 To (intLen / 3)
32180                 strTmp01 = strTmp01 & Chr(Val(Mid(strCaption, (((intX - 1) * 3) + 1), 3)))
32190               Next
32200               strCaption = strTmp01
32210               strRetVal = Format(datTmp02, "mm/dd/yyyy hh:nn:ss AM/PM") & "~" & strCaption
                    ' ** Test: DemoLicense_Get() = #EULA;DB1=#EMPTY
32220             Else
32230 On Error GoTo ERRH
                    ' ** Timestamp invalid.
32240               strRetVal = "#EULA"
32250             End If
32260           Else
                  ' ** Keys match, but pieces don't
32270             strRetVal = "#EULA"
32280           End If
32290         Else
                ' ** Keys don't match.
32300           strRetVal = "#EULA"
32310         End If

32320       Else
              ' ** At least one hidden table was empty,
              ' ** so bring up the EULA.
32330         strRetVal = "#EULA"
32340         If strDB1 = "#EMPTY" Then
                ' ** This Trust.mde hasn't been opened yet.
32350           strRetVal = strRetVal & ";DB1=" & strDB1
32360         Else
32370           If strDB2 = "#EMPTY" Then
                  ' ** Hmm. Frontend OK, but Archive not. ?!
32380             strRetVal = strRetVal & ";DB2=" & strDB2
32390           Else
32400             If strDB3 = "#EMPTY" Then
                    ' ** Must've really bolloxed up the databases!
32410               strRetVal = strRetVal & ";DB3=" & strDB3
32420             Else
                    ' ** Nothing left! Shouldn't get here.
32430               strRetVal = RET_ERR
32440             End If
32450           End If
32460         End If
32470       End If

32480     Else
            ' ** gstrTrustDataLocation directory doesn't exist!
            ' ** This should have already been vetted, so it shouldn't get here.
32490       strRetVal = RET_ERR
32500     End If
32510   Else
          ' ** Unable to get gstrTrustDataLocation.
          ' ** This should have already been vetted, so it shouldn't get here.
32520     strRetVal = RET_ERR
32530   End If  ' ** IniFile_GetDataLoc().

EXITP:
32540   Set tdf = Nothing
32550   Set rst = Nothing
32560   Set dbsLoc = Nothing
32570   Set dbsLnk = Nothing
32580   Set wrk = Nothing
32590   DemoLicense_Get = strRetVal
32600   Exit Function

ERRH:
32610   strRetVal = RET_ERR
32620   Select Case ERR.Number
        Case Else
32630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32640   End Select
32650   Resume EXITP

End Function

Public Function ShowUserRosterMultipleUsers() As Boolean

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowUserRosterMultipleUsers"

        'Dim cnxn1 As New ADODB.Connection, cnxn2 As New ADODB.Connection, rsx1 As New ADODB.Recordset  ' ** Early binding.
        Dim cnxn1 As Object, cnxn2 As Object, rsx1 As Object                                            ' ** Late binding.
        Dim blnRetVal As Boolean

32710   blnRetVal = True

        'Set cnxn1 = New ADODB.Connection             ' ** Early binding.
32720   Set cnxn1 = CreateObject("ADODB.Connection")  ' ** Late binding.
        'Set cnxn2 = New ADODB.Connection             ' ** Early binding.
32730   Set cnxn2 = CreateObject("ADODB.Connection")  ' ** Late binding.

32740   cnxn1.Provider = "Microsoft.Jet.OLEDB.4.0"
32750   cnxn1.Open "Data Source=C:\VictorGCS_Clients\TrustAccountant\NewWorking\TestDatabase\TrustDta.mdb"  '## OK

32760   cnxn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=C:\VictorGCS_Clients\TrustAccountant\NewWorking\TestDatabase\TrustDta.mdb"  '## OK

32770   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

        ' ** The user roster is exposed as a provider-specific schema rowset in the Jet 4 OLE DB provider.
        ' ** You have to use a GUID to reference the schema, as provider-specific schemas are not
        ' ** listed in ADO's type library for schema rowsets.

32780   Set rsx1 = cnxn1.OpenSchema(adSchemaProviderSpecific, , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")
32790   With rsx1

          ' ** Output the list of all users in the current database.
32800     Debug.Print "'" & .Fields(0).Name & "  " & .Fields(1).Name & "  " & .Fields(2).Name & "  " & .Fields(3).Name

32810     While Not .EOF
32820       Debug.Print "'" & Trim(.Fields(0)) & "  " & Trim(.Fields(1)) & "  " & Trim(.Fields(2)) & "  " & Trim(Nz(.Fields(3), vbNullString))
32830       .MoveNext
32840     Wend

32850     .Close
32860   End With

        'COMPUTER_NAME  LOGIN_NAME  CONNECTED  SUSPECT_STATE
        'DELTADATA1   Admin   True
        'DELTADATA1   Admin   True

EXITP:
32870   Set cnxn1 = Nothing
32880   Set cnxn2 = Nothing
32890   Set rsx1 = Nothing
32900   ShowUserRosterMultipleUsers = blnRetVal
32910   Exit Function

ERRH:
32920   blnRetVal = False
32930   Select Case ERR.Number
        Case Else
32940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32950   End Select
32960   Resume EXITP

End Function

Public Sub ListUsersDAO()
' ** List all users in workgroup to debug window using DAO.
' ** From Access 2000 Developer's Handbook, Volume II
' ** by Litwin, Getz, and Gilbert. (Sybex)
' ** Copyright 1999. All Rights Reserved.

33000 On Error GoTo ERRH

        Const THIS_PROC As String = "ListUsersDAO"

        Dim wrk As DAO.Workspace
        Dim usr As DAO.User

33010   Set wrk = DBEngine.Workspaces(0)
33020   With wrk
33030     Debug.Print "The Users collection has " & _
            .Users.Count & " members:"
33040     For Each usr In .Users
33050       Debug.Print usr.Name
33060     Next usr
33070     .Close
33080   End With

EXITP:
33090   Set usr = Nothing
33100   Set wrk = Nothing
33110   Exit Sub

ERRH:
33120   Select Case ERR.Number
        Case Else
33130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
33140   End Select
33150   Resume EXITP

End Sub
