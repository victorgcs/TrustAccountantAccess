Attribute VB_Name = "modFileInfoFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modFileInfoFuncs"

'VGC 01/13/2013: CHANGES!

' ******************************************************************
' ** Copyright ©1996-2009 VBnet, Randy Birch, All Rights Reserved.
' ** Some pages may also contain other copyrights by the author.
' ******************************************************************
' ** Distribution: You can freely use this code in your own
' **               applications, but you may not reproduce
' **               or publish this code on any web site,
' **               online service, or distribute as source
' **               on any media without express permission.
' ******************************************************************

' ** Filter strings for Windows functions (tblFilterString).
Public FLTR_ALL1 As String
Public FLTR_ALL2 As String
Public FLTR_DLL1 As String
Public FLTR_DLL2 As String
Public FLTR_EXE1 As String
Public FLTR_EXE2 As String
Public FLTR_INI1 As String
Public FLTR_INI2 As String
Public FLTR_LDB1 As String
Public FLTR_LDB2 As String
Public FLTR_MDB1 As String
Public FLTR_MDB2 As String
Public FLTR_MDE1 As String
Public FLTR_MDE2 As String
Public FLTR_MDW1 As String
Public FLTR_MDW2 As String
Public FLTR_MD_1 As String
Public FLTR_MD_2 As String
Public FLTR_OCX1 As String
Public FLTR_OCX2 As String
Public FLTR_TLB1 As String
Public FLTR_TLB2 As String
Public FLTR_VXD1 As String
Public FLTR_VXD2 As String
Public FLTR_XLS1 As String
Public FLTR_XLS2 As String

Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersion As Long      ' ** e.g. 0x00000042 = "0.42"
  dwFileVersionMS As Long     ' ** e.g. 0x00030075 = "3.75"
  dwFileVersionLS As Long     ' ** e.g. 0x00000031 = "0.31"
  dwProductVersionMS As Long  ' ** e.g. 0x00030010 = "3.10"
  dwProductVersionLS As Long  ' ** e.g. 0x00000031 = "0.31"
  dwFileFlagsMask As Long     ' ** e.g. 0x3F for version "0.42"
  dwFileFlags As Long         ' ** e.g. VFF_DEBUG Or VFF_PRERELEASE
  dwFileOS As Long            ' ** e.g. VOS_DOS_WINDOWS16
  dwFileType As Long          ' ** e.g. VFT_DRIVER
  dwFileSubtype As Long       ' ** e.g. VFT2_DRV_KEYBOARD
  dwFileDateMS As Long        ' ** e.g. 0
  dwFileDateLS As Long        ' ** e.g. 0
End Type

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
'lptstrFilename [in]
'  Type: LPCTSTR
'  The name of the file of interest. The function uses the search sequence specified by the LoadLibrary function.
'lpdwHandle [out, optional]
'  Type: LPDWORD
'  A pointer to a variable that the function sets to zero.
'Return value
'  Type: DWORD
'  If the function succeeds, the return value is the size, in bytes, of the file's version information.
'  If the function fails, the return value is zero. To get extended error information, call GetLastError().

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
'lptstrFilename [in]
'  Type: LPCTSTR
'  The name of the file. If a full path is not specified, the function uses the search sequence specified by the LoadLibrary function.
'dwHandle
'  Type: DWORD
'  This parameter is ignored.
'dwLen [in]
'  Type: DWORD
'  The size, in bytes, of the buffer pointed to by the lpData parameter.
'  Call the GetFileVersionInfoSize function first to determine the size, in bytes, of a file's version information. The dwLen member should be equal to or greater than that value.
'  If the buffer pointed to by lpData is not large enough, the function truncates the file's version information to the size of the buffer.
'lpData [out]
'  Type: LPVOID
'  Pointer to a buffer that receives the file-version information.
'  You can use this value in a subsequent call to the VerQueryValue function to retrieve data from the buffer.
'Return value
'  Type: BOOL
'  If the function succeeds, the return value is nonzero.
'  If the function fails, the return value is zero. To get extended error information, call GetLastError().

Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" _
  (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, nVerSize As Long) As Long
'pBlock [in]
'  Type: LPCVOID
'  The version-information resource returned by the GetFileVersionInfo function.
'lpSubBlock [in]
'  Type: LPCTSTR
'  The version-information value to be retrieved. The string must consist of names separated by backslashes (\) and it must have one of the following forms.
'  \
'  The root block. The function retrieves a pointer to the VS_FIXEDFILEINFO structure for the version-information resource.
'   \VarFileInfo\Translation
'  The translation array in a Var variable information structure—the Value member of this structure. The function retrieves a pointer to this array of language and code page identifiers. An application can use these identifiers to access a language-specific StringTable structure (using the szKey member) in the version-information resource.
'   \StringFileInfo\lang-codepage\string-name
'  A value in a language-specific StringTable structure. The lang-codepage name is a concatenation of a language and code page identifier pair found as a DWORD in the translation array for the resource. Here the lang-codepage name must be specified as a hexadecimal string. The string-name name must be one of the predefined strings described in the following Remarks section. The function retrieves a string value specific to the language and code page indicated.
'lplpBuffer [out]
'  Type: LPVOID*
'  When this method returns, contains the address of a pointer to the requested version information in the buffer pointed to by pBlock. The memory pointed to by lplpBuffer is freed when the associated pBlock memory is freed.
'puLen [out]
'  Type: PUINT
'  When this method returns, contains a pointer to the size of the requested data pointed to by lplpBuffer: for version information values, the length in characters of the string stored at lplpBuffer; for translation array values, the size in bytes of the array stored at lplpBuffer; and for root block, the size in bytes of the structure.
'Return value
'  Type: BOOL
'  If the specified version-information structure exists, and version information is available, the return value is nonzero. If the address of the length buffer is zero, no value is available for the specified version-information name.
'  If the specified name does not exist or the specified resource is not valid, the return value is zero.
'  This function works on 16-, 32-, and 64-bit file images.
'The following are predefined version information Unicode strings.
'  Comments         InternalName      ProductName
'  CompanyName      LegalCopyright    ProductVersion
'  FileDescription  LegalTrademarks   PrivateBuild
'  FileVersion      OriginalFileName  SpecialBuild

Private Const MAXDWORD As Long = &HFFFFFFFF
'Private Const BIF_MAXPATH As Long = 260
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, _
  lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
'hFile
'  Type:
'  Identifies the files for which to get dates and times. The file handle must have been created with GENERIC_READ access to the file.
'lpCreationTime
'  Points to a FILETIME structure to receive the date and time the file was created. This parameter can be NULL if the application does not require this information.
'lpLastAccessTime
'  Points to a FILETIME structure to receive the date and time the file was last accessed. The last access time includes the last time the file was written to, read from, or, in the case of executable files, run. This parameter can be NULL if the application does not require this information.
'lpLastWriteTime
'  Points to a FILETIME structure to receive the date and time the file was last written to. This parameter can be NULL if the application does not require this information.
'Return Values
'  If the function succeeds, the return value is nonzero.
'  If the function fails, the return value is zero. To get extended error information, call GetLastError().

Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
'lpFileTime
'  The source time and date, which are in UTC time.
'lpLocalFileTime
'  Receives the time and date stored in lpFileTime converted into the computer's current time zone time.

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" _
  (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, _
  ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'lpFileName
'  The name of the file to create or open.
'dwDesiredAccess
'  Zero or more of the following flags specifying the amounts of read/write access to the file.
'dwShareMode
'  Zero or more of the following flags specifying the amounts of read/write access to grant other programs attempting to access the file while your program still has it open:
'lpSecurityAppributes
'  The security attributes to give the created or opened file. In Windows 95, a value of 0 must be passed instead or else the
'  function will fail (VB users must use the alternate Declare).
'dwCreationDisposition
'  Exactly one of the following flags specifying how and when to create or open the file depending if it already does or does
'  not exist:
'dwFlagsAndAttributes
'  The combination of the following flags specifying both the file attributes of a newly created file and other options for creating or opening the file. One flag specifying the file attributes must be included. (The file attributes that can only be set by the operating system are not listed here.)
'hTemplateFile
'  The handle of an open file to copy the attributes of, or 0 to not copy the attributes of any file.

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * BIF_MAXPATH
  cAlternate As String * 14
End Type

Public Type FILE_PARAMS   ' ** Writer's custom type for passing info.
  bRecurse As Boolean     ' ** Var not used in this demo.
  bList As Boolean
  bFound As Boolean       ' ** Var not used in this demo.
  sFileRoot As String
  sFileNameExt As String
  sResult As String       ' ** Var not used in this demo.
  nFileCount As Long      ' ** Var not used in this demo.
  nFileSize As Double     ' ** Var not used in this demo.
End Type

Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long

' ** May include wildcard characters.
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" _
  (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

' ** May include wildcard characters.
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" _
  (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function lstrcpyA Lib "kernel32.dll" (ByVal RetVal As String, ByVal Ptr As Long) As Long

Private Declare Function lstrlenA Lib "kernel32.dll" (ByVal Ptr As Any) As Long

' ** Array: arr_varFldr().
Private lngFldrs As Long, arr_varFldr() As Variant
Private Const FLDR_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const FLDR_PATH As Integer = 0
Private Const FLDR_SUBS As Integer = 1

' ** Array: arr_varFileInfo().
Private lngFileInfos As Long, arr_varFileInfo() As Variant
Private Const FI_ELEMS As Integer = 8  ' ** Array's first_element UBound().
Private Const FI_IDX As Integer = 0
Private Const FI_VER As Integer = 1
Private Const FI_NAM As Integer = 2
Private Const FI_SIZ As Integer = 3
Private Const FI_DSC As Integer = 4
Private Const FI_CDT As Integer = 5
Private Const FI_ADT As Integer = 6
Private Const FI_WDT As Integer = 7
Private Const FI_DIR As Integer = 8
' **

Public Function GetFileSpec1(FP As FILE_PARAMS) As Variant

100   On Error GoTo ERRH

        Const THIS_PROC As String = "GetFileSpec1"

        Dim arr_varRetVal As Variant

110     arr_varRetVal = SearchForFiles(FP, True)  ' ** Function: Below.

120     If lngFileInfos = 0& Then
130       gblnBadSec = False
140     End If

EXITP:
150     GetFileSpec1 = arr_varRetVal
160     Exit Function

ERRH:
170     Select Case ERR.Number
        Case Else
180       MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
190     End Select
200     Resume EXITP

End Function

Public Function GetFileSpec2(blnRecurse As Boolean, blnList As Boolean, strFileRoot As String, strFileExt As String) As Variant
' ** If the user-defined type is not available in the calling
' ** module, it may be instantiated here via simple parameters.

300   On Error GoTo ERRH

        Const THIS_PROC As String = "GetFileSpec2"

        Dim FP As FILE_PARAMS
        Dim intPos1 As Integer
        Dim strTmp00 As String
        Dim arr_varRetVal() As Variant

310     With FP
320       .bRecurse = blnRecurse
330       .bList = blnList
          '.bFound As Boolean    ' ** Returned.
340       .sFileRoot = strFileRoot
350       .sFileNameExt = strFileExt  ' ** May contain wildcards or a specific file.
          '.sResult As String    ' ** Returned.
          '.nFileCount As Long   ' ** Returned.
          '.nFileSize As Double  ' ** Returned.
360     End With

370     SearchForFiles FP  ' ** Function: Below.

380     ReDim arr_varRetVal(FI_ELEMS, 0)

390     With FP
400       If .bFound = True Then
410         strTmp00 = .sResult
420         intPos1 = InStr(strTmp00, "~")
430         If intPos1 > 0 Then
440           arr_varRetVal(FI_IDX, 0) = CLng(Left$(strTmp00, (intPos1 - 1)))
450           strTmp00 = Mid$(strTmp00, (intPos1 + 1))
460           intPos1 = InStr(strTmp00, "~")
470           arr_varRetVal(FI_VER, 0) = Left$(strTmp00, (intPos1 - 1))
480           strTmp00 = Mid$(strTmp00, (intPos1 + 1))
490           intPos1 = InStr(strTmp00, "~")
500           arr_varRetVal(FI_NAM, 0) = Left$(strTmp00, (intPos1 - 1))
510           strTmp00 = Mid$(strTmp00, (intPos1 + 1))
520           intPos1 = InStr(strTmp00, "~")
530           arr_varRetVal(FI_SIZ, 0) = CDbl(Left$(strTmp00, (intPos1 - 1)))
540           strTmp00 = Mid$(strTmp00, (intPos1 + 1))
550           intPos1 = InStr(strTmp00, "~")
560           arr_varRetVal(FI_DSC, 0) = Left$(strTmp00, (intPos1 - 1))
570           arr_varRetVal(FI_DIR, 0) = Mid$(strTmp00, (intPos1 + 1))
580         End If
590       Else
600         arr_varRetVal(FI_IDX, 0) = CLng(0)
610         arr_varRetVal(FI_NAM, 0) = RET_ERR
620       End If
630     End With

EXITP:
640     GetFileSpec2 = arr_varRetVal
650     Exit Function

ERRH:
660     Select Case ERR.Number
        Case Else
670       MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
680     End Select
690     Resume EXITP

End Function

Public Function SearchForFiles(FP As FILE_PARAMS, Optional varInit As Variant) As Double
' ** Returns the size of matching files.

700   On Error GoTo ERRH

        Const THIS_PROC As String = "SearchForFiles"

        Dim WFD As WIN32_FIND_DATA
        Dim lngFile As Long
        Dim dblSize As Double
        Dim strPath As String
        Dim strRoot As String
        Dim strTmp00 As String

710     If IsMissing(varInit) = False Then
720       If varInit = True Then
            ' ** Initialize the array.
730         lngFileInfos = 0&
740         ReDim arr_varFileInfo(FI_ELEMS, 0)
750       End If
760     End If

770     strRoot = QualifyPath(FP.sFileRoot)  ' ** Function: Below.
780     strPath = strRoot & "*.*"  '& FP.sFileNameExt

        ' ** Obtain handle to the first match. BUT THIS ISN'T A MATCH! IT'S *.*!!!!!!!
790     lngFile = FindFirstFile(strPath, WFD)  ' ** API Function: Above.

        ' ** If valid ...
800     If lngFile <> INVALID_HANDLE_VALUE Then

          ' ** This is where the method obtains the file list and data for the folder passed.

          ' ** GetFileInformation function returns the size,
          ' ** in bytes, of the files found matching the
          ' ** filespec in the passed folder, so its
          ' ** assigned to dblSize. It is not directly assigned
          ' ** to FP.nFileSize because dblSize is incremented
          ' ** below if a recursive search was specified.
810       dblSize = GetFileInformation(FP)  ' ** Function: Below.
820       FP.nFileSize = dblSize

830       Do

            ' ** If the returned item is a folder...
840         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then

              ' ** ..and the Recurse flag was specified.
850           If FP.bRecurse Then

                ' ** Remove trailing nulls.
860             strTmp00 = TrimNull(WFD.cFileName)  ' ** Function: Below.

                ' ** And if the folder is not the default self and parent folders...
870             If strTmp00 <> "." And strTmp00 <> ".." Then

                  ' ** ..then the item is a real folder, which
                  ' ** may contain other sub folders, so assign
                  ' ** the new folder name to FP.sFileRoot and
                  ' ** recursively call this function again with
                  ' ** the amended information.

                  ' ** Since dblSize is a local variable whose value
                  ' ** is both set above as well as returned as the
                  ' ** function call value, dblSize needs to be added
                  ' ** to previous calls in order to maintain accuracy.

                  ' ** However, because the nFileSize member of
                  ' ** FILE_PARAMS is passed back and forth through
                  ' ** the calls, dblSize is simply assigned to it
                  ' ** after the recursive call finishes.
880               FP.sFileRoot = strRoot & strTmp00
890               dblSize = dblSize + SearchForFiles(FP)  ' ** This, recursive.
900               FP.nFileSize = dblSize

910             End If  ' ** Not self or parent.
920           End If  ' ** Do recurse.
930         End If  ' ** Is directory.

            ' ** Continue looping until FindNextFile returns 0 (no more matches).
940       Loop While FindNextFile(lngFile, WFD)  ' ** API Function: Above.

          ' ** Close the find handle.
950       lngFile = FindClose(lngFile)  ' ** API Function: Above.

960     End If

EXITP:
        ' ** Because this routine is recursive, return the size of matching files.
970     SearchForFiles = dblSize  ' ** Why is the Function defined as Double, but we always return Long?
980     Exit Function             ' ** To handle very large file sizes, which would overflow a Long Integer!

ERRH:
990     Select Case ERR.Number
        Case Else
1000      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1010    End Select
1020    Resume EXITP

End Function

Private Function GetFileInformation(FP As FILE_PARAMS) As Long

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFileInformation"

        Dim WFD As WIN32_FIND_DATA
        Dim lngFile As Long, lngSize As Long
        Dim strPath As String, strRoot As String, strName As String
        Dim datDateCreated As Date, datDateAccessed As Date, datDateWritten As Date
        Dim strTmp00 As String
        Dim lngE As Long
        Dim blnRetVal As Boolean

        ' ** FP.sFileRoot (assigned to strRoot) contains the path to search.

        ' ** FP.sFileNameExt (assigned to strPath) contains the full path and filespec.
1110    strRoot = QualifyPath(FP.sFileRoot)  ' ** Function: Below.
1120    strPath = strRoot & FP.sFileNameExt

        ' ** Obtain handle to the first filespec match.
1130    lngFile = FindFirstFile(strPath, WFD)  ' ** API Function: Above.

        ' ** If valid ...
1140    If lngFile <> INVALID_HANDLE_VALUE Then

1150      strTmp00 = vbNullString

1160      Do

            ' ** Remove trailing nulls.
1170        strName = TrimNull(WFD.cFileName)  ' ** Function: Below.

            ' ** Even though this routine uses filespecs,
            ' ** *.* is still valid and will cause the search
            ' ** to return folders as well as files, so a
            ' ** check against folders is still required.
1180        If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then

              ' ** File found, so increase the file count.
1190          FP.nFileCount = FP.nFileCount + 1

              ' ** Retrieve the size and assign to lngSize to
              ' ** be returned at the end of this function call.
1200          lngSize = lngSize + (WFD.nFileSizeHigh * (MAXDWORD + 1)) + WFD.nFileSizeLow

              ' ** Add to the list if the flag indicates.
1210          Select Case FP.bList
              Case True

                ' **********************************************
                ' ** Array: arr_varFileInfo()
                ' **
                ' **   Element  Name                Constant
                ' **   =======  ==================  ==========
                ' **      0     Index               FI_IDX
                ' **      1     Version             FI_VER
                ' **      2     File Name           FI_NAM
                ' **      3     File Size           FI_SIZ
                ' **      4     File Description    FI_DSC
                ' **      5     Created Date        FI_CDT
                ' **      6     Accessed Date       FI_ADT
                ' **      7     Written Date        FI_WDT
                ' **      8     File Root Folder    FI_DIR
                ' **
                ' **********************************************

1220            lngFileInfos = lngFileInfos + 1&
1230            lngE = lngFileInfos - 1&
1240            ReDim Preserve arr_varFileInfo(FI_ELEMS, lngE)

                ' ** Got the data, so add it to the array.
1250            arr_varFileInfo(FI_IDX, lngE) = lngE
1260            arr_varFileInfo(FI_VER, lngE) = GetFileVersion(strRoot & strName)  ' ** Function: Below.
1270            arr_varFileInfo(FI_NAM, lngE) = strName
1280            arr_varFileInfo(FI_SIZ, lngE) = GetFileSizeStr(WFD.nFileSizeHigh + WFD.nFileSizeLow)  ' ** Function: Below.
1290            arr_varFileInfo(FI_DSC, lngE) = GetFileDescription(strRoot & strName)  ' ** Function: Below.
1300            arr_varFileInfo(FI_DIR, lngE) = strRoot
1310            blnRetVal = GetFileTimes(strRoot & strName, datDateCreated, datDateAccessed, datDateWritten, True)  ' ** Function: Below.
1320            Select Case blnRetVal
                Case True
1330              arr_varFileInfo(FI_CDT, lngE) = datDateCreated
1340              arr_varFileInfo(FI_ADT, lngE) = datDateAccessed
1350              arr_varFileInfo(FI_WDT, lngE) = datDateWritten
1360            Case False
1370              arr_varFileInfo(FI_CDT, lngE) = 0
1380              arr_varFileInfo(FI_ADT, lngE) = 0
1390              arr_varFileInfo(FI_WDT, lngE) = 0
1400            End Select
1410          Case False
                ' ** Not a list, just a single file.
1420            strTmp00 = "1~"                                                                   ' ** FI_IDX
1430            strTmp00 = strTmp00 & GetFileVersion(strRoot & strName) & "~"                     ' ** FI_VER; Function: Below.
1440            strTmp00 = strTmp00 & strName & "~"                                               ' ** FI_NAM.
1450            strTmp00 = strTmp00 & GetFileSizeStr(WFD.nFileSizeHigh + WFD.nFileSizeLow) & "~"  ' ** FI_SIZ; Function: Below.
1460            strTmp00 = strTmp00 & GetFileDescription(strRoot & strName) & "~"                 ' ** FI_DSC; Function: Below.
1470            strTmp00 = strTmp00 & strRoot                                                     ' ** FI_DIR.
1480            FP.sResult = strTmp00
1490            FP.bFound = True
1500          End Select

1510        End If

1520      Loop While FindNextFile(lngFile, WFD)  ' ** API Function: Above.

          ' ** Close the handle.
1530      lngFile = FindClose(lngFile)  ' ** API Function: Above.

1540    End If

EXITP:
        ' ** Return the size of files found.
1550    GetFileInformation = lngSize
1560    Exit Function

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1590    End Select
1600    Resume EXITP

End Function

Private Function GetFileSizeStr(fsize As Long) As String

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFileSizeStr"

        Dim strRetVal As String

1710    strRetVal = Format((fsize), "###,###,###")  '& " kb"

EXITP:
1720    GetFileSizeStr = strRetVal
1730    Exit Function

ERRH:
1740    Select Case ERR.Number
        Case Else
1750      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1760    End Select
1770    Resume EXITP

End Function

Private Function QualifyPath(sPath As String) As String
' ** Assures that a passed path ends in a slash.

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "QualifyPath"

        Dim strRetVal As String

1810    If Right$(sPath, 1) <> "\" Then
1820      strRetVal = sPath & "\"
1830    Else
1840      strRetVal = sPath
1850    End If

EXITP:
1860    QualifyPath = strRetVal
1870    Exit Function

ERRH:
1880    Select Case ERR.Number
        Case Else
1890      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1900    End Select
1910    Resume EXITP

End Function

Private Function TrimNull(strInput As String) As String
' ** Returns the string up to the first null, if present, or the passed string

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "TrimNull"

        Dim intPos1 As Integer
        Dim strRetVal As String

2010    intPos1 = InStr(strInput, Chr$(0))

2020    If intPos1 > 0 Then
2030      strRetVal = Left$(strInput, intPos1 - 1)
2040    Else
2050      strRetVal = strInput
2060    End If

EXITP:
2070    TrimNull = strRetVal
2080    Exit Function

ERRH:
2090    Select Case ERR.Number
        Case Else
2100      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2110    End Select
2120    Resume EXITP

End Function

Private Function HiWord(lngDW As Long) As Long

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "HiWord"

        Dim lngRetVal As Long

2210    If lngDW And &H80000000 Then
2220      lngRetVal = (lngDW \ 65535) - 1
2230    Else
2240      lngRetVal = lngDW \ 65535
2250    End If

EXITP:
2260    HiWord = lngRetVal
2270    Exit Function

ERRH:
2280    Select Case ERR.Number
        Case Else
2290      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2300    End Select
2310    Resume EXITP

End Function

Private Function LoWord(lngDW As Long) As Long

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "LoWord"

        Dim lngRetVal As Long

2410    If lngDW And &H8000& Then
2420      lngRetVal = &H8000& Or (lngDW And &H7FFF&)
2430    Else
2440      lngRetVal = lngDW And &HFFFF&
2450    End If

EXITP:
2460    LoWord = lngRetVal
2470    Exit Function

ERRH:
2480    Select Case ERR.Number
        Case Else
2490      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2500    End Select
2510    Resume EXITP

End Function

Private Function GetFileDescription(strSourceFile As String) As String

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFileDescription"

        Dim FI As VS_FIXEDFILEINFO
        Dim arr_bytBuffer() As Byte
        Dim lngBufferSize As Long
        Dim lngBuffer As Long
        Dim lngVerSize As Long
        Dim lngUnused As Long
        Dim strTmpVer As String
        Dim strBlock As String

2610    If Len(strSourceFile) > 0 Then

          ' ** Set file that has the encryption level info and call to get required size.
2620      lngBufferSize = GetFileVersionInfoSize(strSourceFile, lngUnused)  ' ** API Function: Above.

2630      ReDim arr_bytBuffer(lngBufferSize)

2640      If lngBufferSize > 0 Then

            ' ** Get the version info.
2650        GetFileVersionInfo strSourceFile, 0&, lngBufferSize, arr_bytBuffer(0)   ' ** API Function: Above.
2660        VerQueryValue arr_bytBuffer(0), "\", lngBuffer, lngVerSize  ' ** API Function: Above.
2670        CopyMemory FI, ByVal lngBuffer, Len(FI)  ' ** API Function: Above.

2680        If VerQueryValue(arr_bytBuffer(0), "\VarFileInfo\Translation", lngBuffer, lngVerSize) Then  ' ** API Function: Above.

2690          If lngVerSize Then

2700            strTmpVer = GetPointerToString(lngBuffer, lngVerSize)  ' ** Function: Below.
2710            strTmpVer = Right("0" & Hex(Asc(Mid(strTmpVer, 2, 1))), 2) & _
                  Right("0" & Hex(Asc(Mid(strTmpVer, 1, 1))), 2) & _
                  Right("0" & Hex(Asc(Mid(strTmpVer, 4, 1))), 2) & _
                  Right("0" & Hex(Asc(Mid(strTmpVer, 3, 1))), 2)
2720            strBlock = "\StringFileInfo\" & strTmpVer & "\FileDescription"

                ' ** Get predefined version resources.
2730            If VerQueryValue(arr_bytBuffer(0), strBlock, lngBuffer, lngVerSize) Then  ' ** API Function: Above.

2740              If lngVerSize Then

                    ' ** Get the file description.
2750                GetFileDescription = GetStrFromPtrA(lngBuffer)  ' ** Function: Below.

2760              End If   ' ** If lngVerSize.
2770            End If   ' ** If VerQueryValue.
2780          End If   ' ** If lngVerSize.
2790        End If   ' ** If VerQueryValue.
2800      End If   ' ** If lngBufferSize.
2810    End If   ' ** If strSourceFile.

EXITP:
2820    Exit Function

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2850    End Select
2860    Resume EXITP

End Function

Private Function GetStrFromPtrA(ByVal lngSzA As Long) As String

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "GetStrFromPtrA"

        Dim strRetVal As String

2910    strRetVal = String$(lstrlenA(ByVal lngSzA), 0)
2920    lstrcpyA ByVal strRetVal, ByVal lngSzA   ' ** API Function: Above.

EXITP:
2930    GetStrFromPtrA = strRetVal
2940    Exit Function

ERRH:
2950    Select Case ERR.Number
        Case Else
2960      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2970    End Select
2980    Resume EXITP

End Function

Private Function GetPointerToString(lngString As Long, lngBytes As Long) As String

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "GetPointerToString"

        Dim strBuffer As String
        Dim strRetVal As String

3010    If lngBytes <> 0& Then
3020      strBuffer = Space$(lngBytes)
3030      CopyMemory ByVal strBuffer, ByVal lngString, lngBytes  ' ** API Function: Above.
3040      strRetVal = strBuffer
3050    Else
3060      strRetVal = vbNullString
3070    End If

EXITP:
3080    GetPointerToString = strRetVal
3090    Exit Function

ERRH:
3100    Select Case ERR.Number
        Case Else
3110      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3120    End Select
3130    Resume EXITP

End Function

Private Function GetFileVersion(strDriverFile As String) As String

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFileVersion"

        Dim FI As VS_FIXEDFILEINFO
        Dim arr_bytBuffer() As Byte
        Dim lngBufferSize As Long, lngBuffer As Long
        Dim lngVerSize As Long, lngUnused As Long
        Dim strRetVal As String

        ' ** GetFileVersionInfoSize determines whether the operating
        ' ** system can obtain version information about a specified
        ' ** file. If version information is available, it returns
        ' ** the size in bytes of that information. As with other
        ' ** file installation functions, GetFileVersionInfoSize
        ' ** works only with Win32 file images.

        ' ** A empty variable must be passed as the second
        ' ** parameter, which the call returns 0 in.
3210    lngBufferSize = GetFileVersionInfoSize(strDriverFile, lngUnused)  ' ** API Function: Above.

3220    If lngBufferSize > 0 Then

          ' ** create a buffer to receive file-version
          ' ** (FI) information.
3230      ReDim arr_bytBuffer(lngBufferSize)
3240      GetFileVersionInfo strDriverFile, 0&, lngBufferSize, arr_bytBuffer(0)  ' ** API Function: Above.

          ' ** VerQueryValue function returns selected version info
          ' ** from the specified version-information resource. Grab
          ' ** the file info and copy it into the  VS_FIXEDFILEINFO structure.
3250      VerQueryValue arr_bytBuffer(0), "\", lngBuffer, lngVerSize  ' ** API Function: Above.
3260      CopyMemory FI, ByVal lngBuffer, Len(FI)  ' ** API Function: Above.

          ' ** Extract the file version from the FI structure
3270      strRetVal = Format(HiWord(FI.dwFileVersionMS)) & "." & Format(LoWord(FI.dwFileVersionMS), "00") & "."  ' ** Functions: Above.

3280      If FI.dwFileVersionLS > 0 Then
3290        strRetVal = strRetVal & Format(HiWord(FI.dwFileVersionLS), "00") & "." & Format(LoWord(FI.dwFileVersionLS), "00")  ' ** Functions: Above.
3300      Else
3310        strRetVal = strRetVal & Format(FI.dwFileVersionLS, "0000")
3320      End If

3330    End If

EXITP:
3340    GetFileVersion = strRetVal
3350    Exit Function

ERRH:
3360    Select Case ERR.Number
        Case Else
3370      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3380    End Select
3390    Resume EXITP

End Function

Public Function GetFileInfoArray() As Variant

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFileInfoArray"

        Dim varRetVal As Variant

3410    varRetVal = arr_varFileInfo

EXITP:
3420    GetFileInfoArray = varRetVal
3430    Exit Function

ERRH:
3440    Select Case ERR.Number
        Case Else
3450      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3460    End Select
3470    Resume EXITP

End Function

Public Function GetFileTimes(ByVal strFileName As String, ByRef datDateCreated As Date, ByRef datDateAccessed As Date, ByRef datDateWritten As Date, ByVal blnLocalTime As Boolean) As Boolean
' ** Returns False if there is an error.

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFileTimes"

        Dim lngFileHandle As Long
        Dim FT_CreationTime As FILETIME
        Dim FT_AccessTime As FILETIME
        Dim FT_WriteTime As FILETIME
        Dim FT_FileTime As FILETIME
        Dim SA As SECURITY_ATTRIBUTES  ' ** Not using this caused INVALID_HANDLE_VALUE error!
        Dim blnRetVal As Boolean

3510    blnRetVal = True

        ' ** Use the GetFileTime API function. This gives the file times in UTC (Universal Time
        ' ** Coordinates which is the same as GMT). If you want times in your local system time,
        ' ** use the FileTimeToLocalFileTime API function to convert the values.
        ' ** Then translate the results (which are in 100s of nanoseconds since January 1, 1601)
        ' ** into the date and time.

        ' ** FILETIME user-defined type:
        ' **   dwLowDateTime As Long
        ' **   dwHighDateTime As Long

        'GetFileTimes("C:\VictorGCS_Clients\TrustAccountant\NewWorking\accessrt.txt", gdatStartDate, gdatEndDate, gdatAccept, True)  '## OK
        'True
        'gdatStartDate = 06/30/2009 12:26:20 PM
        'gdatEndDate = 06/06/2011 09:46:09 PM
        'gdatAccept = 09/23/2011 12:07:00 PM

        ' ** Open the file.
3520    lngFileHandle = CreateFile(strFileName, GENERIC_READ, FILE_SHARE_READ, SA, OPEN_EXISTING, _
          FILE_ATTRIBUTE_ARCHIVE, 0&)  ' ** API Function: Above.
3530    If lngFileHandle = 0& Then
3540      blnRetVal = False
3550    ElseIf lngFileHandle = INVALID_HANDLE_VALUE Then
3560      blnRetVal = False
          'Debug.Print "'The handle is invalid."
3570    Else
          ' ** Get the times.
3580      If GetFileTime(lngFileHandle, FT_CreationTime, FT_AccessTime, FT_WriteTime) = 0& Then  ' ** API Function: Above.
3590        blnRetVal = False
            'Debug.Print "'" & FormatMessageTxt(GetLastError)
3600      Else
            ' ** Close the file.
3610        If CloseHandle(lngFileHandle) = 0 Then  ' ** API Function: modBackup.
3620          blnRetVal = False
3630        Else

              ' ** See if we should convert to the local file system time.
3640          If blnLocalTime = True Then

                ' ** Convert to local file system time.
3650            FileTimeToLocalFileTime FT_CreationTime, FT_FileTime  ' ** API Function: Above.
3660            FT_CreationTime = FT_FileTime

3670            FileTimeToLocalFileTime FT_AccessTime, FT_FileTime  ' ** API Function: Above.
3680            FT_AccessTime = FT_FileTime

3690            FileTimeToLocalFileTime FT_WriteTime, FT_FileTime  ' ** API Function: Above.
3700            FT_WriteTime = FT_FileTime

3710          End If

              ' ** Convert into dates.
3720          datDateCreated = FileTimeToDate(FT_CreationTime)  ' ** Function: Below.
3730          datDateAccessed = FileTimeToDate(FT_AccessTime)  ' ** Function: Below.
3740          datDateWritten = FileTimeToDate(FT_WriteTime)  ' ** Function: Below.

3750        End If
3760      End If
3770    End If

EXITP:
3780    GetFileTimes = blnRetVal
3790    Exit Function

ERRH:
3800    blnRetVal = False
3810    Select Case ERR.Number
        Case Else
3820      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3830    End Select
3840    Resume EXITP

End Function

Private Function FileTimeToDate(FT As FILETIME) As Date
' ** Convert the FILETIME structure into a Date.
' ** FILETIME units are 100s of nanoseconds.

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "FileTimeToDate"

        Dim cblLo_Time As Double, dblHi_Time As Double
        Dim dblSeconds As Double, dblHours As Double
        Dim datTheDate As Date
        Dim datRetVal As Date

        Const TICKS_PER_SECOND As Long = 10000000

3910    datRetVal = 0

        ' ** Get the low order data.
3920    If FT.dwLowDateTime < 0 Then
3930      cblLo_Time = 2 ^ 31 + (FT.dwLowDateTime And &H7FFFFFFF)
3940    Else
3950      cblLo_Time = FT.dwLowDateTime
3960    End If

        ' ** Get the high order data.
3970    If FT.dwHighDateTime < 0 Then
3980      dblHi_Time = 2 ^ 31 + (FT.dwHighDateTime And &H7FFFFFFF)
3990    Else
4000      dblHi_Time = FT.dwHighDateTime
4010    End If

        ' ** Combine them and turn the result into hours.
4020    dblSeconds = (cblLo_Time + 2 ^ 32 * dblHi_Time) / TICKS_PER_SECOND
4030    dblHours = CLng(dblSeconds / 3600)
4040    dblSeconds = dblSeconds - dblHours * 3600

        ' ** Make the date.
4050    datTheDate = DateAdd("h", dblHours, "1/1/1601 0:00 AM")
4060    datTheDate = DateAdd("s", dblSeconds, datTheDate)
4070    datRetVal = datTheDate

EXITP:
4080    FileTimeToDate = datRetVal
4090    Exit Function

ERRH:
4100    datRetVal = 0
4110    Select Case ERR.Number
        Case Else
4120      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4130    End Select
4140    Resume EXITP

End Function

Public Function Find_File(Optional varFile As Variant, Optional varMode As Variant) As Variant

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Find_File"

        Dim fso As Scripting.FileSystemObject, fsd As Scripting.Drive, fsfds As Scripting.Folders, fsfls As Scripting.Files
        Dim fsfd1 As Scripting.Folder, fsfd2 As Scripting.Folder, fsfl As Scripting.File
        Dim strFind As String
        Dim lngFiles_Root As Long, lngFolders_Root As Long
        Dim blnSpecDir As Boolean, strSpecDir1 As String, strSpecDir2 As String
        Dim blnContinue As Boolean, blnCalled As Boolean, blnFound As Boolean
        Dim intPos1 As Integer
        Dim lngX As Long, lngE As Long
        Dim varRetVal As Variant, lngRetVals As Long, arr_varRetVal() As Variant

        ' ** Array: arr_varRetVal().
        Const RV_ELEMS As Integer = 5  ' ** Array's first-element UBound().
        Const RV_NAME As Integer = 0
        Const RV_PATH As Integer = 1
        Const RV_SIZE As Integer = 2
        Const RV_CDAT As Integer = 3
        Const RV_ADAT As Integer = 4
        Const RV_WDAT As Integer = 5

4210    varRetVal = Empty
4220    lngRetVals = 0&
4230    ReDim arr_varRetVal(RV_ELEMS, 0)
4240    blnContinue = True

4250    Select Case IsMissing(varFile)
        Case True
4260      strFind = "TrustSec.mdw"
4270      If Len(TA_SEC) > Len(TA_SEC2) Then
4280        strSpecDir1 = "C:\VictorGCS_Clients\TrustAccountant\NewDemo\DemoDatabase"
4290      Else
4300        strSpecDir1 = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\EmptyDatabase"
4310      End If
4320      strSpecDir2 = "VictorGCS_Clients"
4330      blnSpecDir = True  ' ** True: specify directory; False: entire drive.
4340      blnCalled = False
4350    Case False
4360      blnCalled = True
4370      Select Case IsNull(varFile)
          Case True
4380        blnContinue = False
4390        arr_varRetVal(RV_PATH, 0) = RET_ERR
4400      Case False
4410        strFind = Trim(varFile)
4420        If strFind <> vbNullString Then
4430          intPos1 = InStr(strFind, LNK_SEP)
4440          If intPos1 > 0 Then
4450            blnSpecDir = True
4460            strSpecDir1 = Parse_Path(strFind)  ' ** Module Function: modFileUtilities.
4470            strFind = Parse_File(strFind)  ' ** Module Function: modFileUtilities.
4480            intPos1 = InStr(strSpecDir1, LNK_SEP)
4490            If intPos1 > 0 Then
4500              strSpecDir2 = Mid$(strSpecDir1, (intPos1 + 1))
4510              If Left$(strSpecDir2, 1) = LNK_SEP Then strSpecDir2 = Mid$(strSpecDir2, 2)  ' ** \\DELTADATA2.
4520              intPos1 = InStr(strSpecDir2, LNK_SEP)
4530              If intPos1 > 0 Then
4540                strSpecDir2 = Left$(strSpecDir2, (intPos1 - 1))
4550              End If
4560            Else
4570              strSpecDir2 = strSpecDir1
4580            End If
4590          Else
4600            blnSpecDir = False
4610          End If
4620        Else
4630          blnContinue = False
4640          arr_varRetVal(RV_PATH, 0) = RET_ERR
4650        End If  ' ** vbNullString.
4660      End Select  ' ** IsNull().
4670    End Select  ' ** IsMissing().

4680    If blnContinue = True Then
4690      Set fso = CreateObject("Scripting.FileSystemObject")
4700      With fso
4710        Set fsd = .GetDrive("C")  ' ** Or 'C:' or 'C:\'.
4720        With fsd
4730          Set fsfd1 = .RootFolder
4740          With fsfd1

                ' ** Check root files.
4750            lngFiles_Root = .Files.Count
4760            If lngFiles_Root > 0 Then
4770              Set fsfls = .Files
4780              For Each fsfl In fsfls
4790                With fsfl
4800                  If InStr(.Name, strFind) > 0 Then
4810                    Select Case blnCalled
                        Case True
4820                      lngRetVals = lngRetVals + 1&
4830                      lngE = lngRetVals - 1&
4840                      ReDim Preserve arr_varRetVal(RV_ELEMS, lngE)
4850                      arr_varRetVal(RV_NAME, lngE) = .Name
4860                      arr_varRetVal(RV_PATH, lngE) = .Path
4870                      Select Case varMode
                          Case 1, 2  ' ** Return all pieces of information.
4880                        arr_varRetVal(RV_SIZE, lngE) = .Size
4890                        arr_varRetVal(RV_CDAT, lngE) = .DateCreated
4900                        arr_varRetVal(RV_ADAT, lngE) = .DateLastAccessed
4910                        arr_varRetVal(RV_WDAT, lngE) = .DateLastModified
4920                      Case Else
                            ' ** Nothing at the moment.
4930                      End Select
4940                      If varMode = 1 Then  ' ** Only return 1 file.
4950                        blnContinue = False
4960                      End If
4970                    Case False
4980                      Debug.Print "'" & .Name & " : " & .Path
4990                    End Select
5000                  End If
5010                End With
5020                If blnContinue = False Then
5030                  Exit For
5040                End If
5050              Next
5060            Else
5070              If blnCalled = False Then
5080                Debug.Print "'" & .Path & " has no root files."
5090              End If
5100            End If

5110            If blnContinue = True Then

5120              lngFldrs = 0&
5130              ReDim arr_varFldr(FLDR_ELEMS, 0)

                  ' ** Check each subfolder.
5140              lngFolders_Root = .SubFolders.Count
5150              If lngFolders_Root > 0& Then
                    ' ** Fill the array with all folders on the drive: arr_varFldr().
5160                Set fsfds = .SubFolders
5170                For Each fsfd2 In fsfds
5180                  With fsfd2
5190                    blnFound = False
5200                    Select Case blnSpecDir
                        Case True
5210                      If .Name = strSpecDir2 Then
5220                        blnFound = True
5230                      End If
5240                    Case False
5250                      blnFound = True
5260                    End Select
5270                    If blnFound = True Then
5280                      Find_File_Subs fsfd2  ' ** Procedure: Below.  RECURRSIVE!
5290                    End If
5300                    If blnSpecDir = True And blnFound = True Then
5310                      Exit For
5320                    End If
5330                  End With
5340                Next
5350              Else
5360                If blnCalled = False Then
5370                  Debug.Print "'" & .Path & " has no folders."
5380                End If
5390              End If

5400            End If  ' ** blnContinue.

5410          End With  ' ** Root folder: fsfd1.
5420        End With  ' ** CD drive: fsd.

5430        If blnContinue = True Then
              ' ** Now search for the files.
5440          If lngFldrs > 0& Then
5450            Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
5460            For lngX = 0& To (lngFldrs - 1&)
5470              blnFound = False
5480              Select Case blnSpecDir
                  Case True
5490                If Left$(arr_varFldr(FLDR_PATH, lngX), Len(strSpecDir1)) = strSpecDir1 Then
5500                  blnFound = True
5510                End If
5520              Case False
5530                blnFound = True
5540              End Select
5550              If blnFound = True Then
5560                Set fsfd2 = fso.GetFolder(arr_varFldr(FLDR_PATH, lngX))
5570                With fsfd2
5580                  Set fsfls = .Files
5590                  For Each fsfl In fsfls
5600                    With fsfl
5610                      If InStr(.Name, strFind) > 0 Then
5620                        Select Case blnCalled
                            Case True
5630                          lngRetVals = lngRetVals + 1&
5640                          lngE = lngRetVals - 1&
5650                          ReDim Preserve arr_varRetVal(RV_ELEMS, lngE)
5660                          arr_varRetVal(RV_NAME, lngE) = .Name
5670                          arr_varRetVal(RV_PATH, lngE) = .Path
5680                          Select Case varMode
                              Case 1, 2  ' ** Return all pieces of information.
5690                            arr_varRetVal(RV_SIZE, lngE) = .Size
5700                            arr_varRetVal(RV_CDAT, lngE) = .DateCreated
5710                            arr_varRetVal(RV_ADAT, lngE) = .DateLastAccessed
5720                            arr_varRetVal(RV_WDAT, lngE) = .DateLastModified
5730                          Case Else
                                ' ** Nothing at the moment.
5740                          End Select
5750                          If varMode = 1 Then  ' ** Only return 1 file.
5760                            blnContinue = False
5770                          End If
5780                        Case False
5790                          Debug.Print "'" & .Name & " : " & .Path
5800                          Debug.Print "'  CREATED: " & .DateCreated & "  LAST ACCESSED: " & .DateLastAccessed & "  LAST MODIFIED: " & .DateLastModified
5810                        End Select
5820                      End If
5830                    End With
5840                  Next
5850                End With
5860              End If  ' ** blnFound.
5870            Next
5880          End If  ' ** lngFldrs.
5890        End If  ' ** blnContinue.

5900        If blnCalled = True And lngRetVals = 0& Then
5910          arr_varRetVal(RV_NAME, 0) = "#NOT FOUND"
5920          arr_varRetVal(RV_PATH, 0) = vbNullString
5930          arr_varRetVal(RV_SIZE, 0) = CLng(0)
5940          arr_varRetVal(RV_CDAT, 0) = Null
5950          arr_varRetVal(RV_ADAT, 0) = Null
5960          arr_varRetVal(RV_WDAT, 0) = Null
5970        End If

5980      End With  ' ** File System Object: fso.
5990    End If  ' ** blnContinue.

6000    Select Case blnCalled
        Case True
6010      varRetVal = arr_varRetVal
6020    Case False
6030      Beep
6040      varRetVal = CBool(True)
6050    End Select

EXITP:
6060    Set fsfl = Nothing
6070    Set fsfd1 = Nothing
6080    Set fsfd2 = Nothing
6090    Set fsfls = Nothing
6100    Set fsfds = Nothing
6110    Set fsd = Nothing
6120    Set fso = Nothing
6130    Find_File = varRetVal
6140    Exit Function

ERRH:
6150    Select Case blnCalled
        Case True
6160      arr_varRetVal(RV_NAME, 0) = RET_ERR
6170      varRetVal = arr_varRetVal
6180    Case False
6190      varRetVal = CBool(False)
6200    End Select
6210    Select Case ERR.Number
        Case Else
6220      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
6230    End Select
6240    Resume EXITP

End Function

Private Sub Find_File_Subs(fsfd As Scripting.Folder)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "Find_File_Subs"

        Dim fsfdx As Scripting.Folder
        Dim lngE As Long

6310    lngFldrs = lngFldrs + 1&
6320    lngE = lngFldrs - 1&
6330    ReDim Preserve arr_varFldr(FLDR_ELEMS, lngE)
6340    arr_varFldr(FLDR_PATH, lngE) = fsfd.Path
6350    arr_varFldr(FLDR_SUBS, lngE) = fsfd.SubFolders.Count
6360    If fsfd.SubFolders.Count > 0& Then
6370      For Each fsfdx In fsfd.SubFolders
6380        Find_File_Subs fsfdx  ' ** RECURSIVE!
6390      Next
6400    End If

EXITP:
6410    Set fsfdx = Nothing  ' ** DO I STILL DO THIS ON RECURSIVE PROCEDURES?
6420    Exit Sub

ERRH:
6430    Select Case ERR.Number
        Case 70  ' ** Permission denied.
          ' ** Why?
6440    Case Else
6450      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
6460    End Select
6470    Resume EXITP

End Sub
