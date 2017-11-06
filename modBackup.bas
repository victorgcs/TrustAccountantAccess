Attribute VB_Name = "modBackup"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modBackup"

'VGC 03/23/2017: CHANGES!

'Private Const fsAttrNormal     As Integer = 0     ' ** Normal file. No attributes are set.
Private Const fsAttrReadOnly   As Integer = 1     ' ** Read-only file. Attribute is read/write.
'Private Const fsAttrHidden     As Integer = 2     ' ** Hidden file. Attribute is read/write.
'Private Const fsAttrSystem     As Integer = 4     ' ** System file. Attribute is read/write.
'Private Const fsAttrVolume     As Integer = 8     ' ** Disk drive volume label. Attribute is read-only.
'Private Const fsAttrDirectory  As Integer = 16    ' ** Folder or directory. Attribute is read-only.
'Private Const fsAttrArchive    As Integer = 32    ' ** File has changed since last backup. Attribute is read/write.
'Private Const fsAttrAlias      As Integer = 1024  ' ** Link or shortcut. Attribute is read-only.
'Private Const fsAttrCompressed As Integer = 2048  ' ** Compressed file. Attribute is read-only.

'Public Const CREATE_NEW        As Long = 1  ' ** Create a new file. The function fails if the file already exists.
Public Const CREATE_ALWAYS     As Long = 2  ' ** Create a new file. Overwrite the file (i.e., delete the old one first) if it already exists.
Public Const OPEN_EXISTING     As Long = 3  ' ** Open an existing file. The function fails if the file does not exist.
'Public Const OPEN_ALWAYS       As Long = 4  ' ** Open an existing file. If the file does not exist, it will be created.
'Public Const TRUNCATE_EXISTING As Long = 5  ' ** Open an existing file and delete its contents. The function fails if the file does not exist.

Public Const FILE_SHARE_READ   As Long = &H1  ' ** Allow other programs to read data from the file.
Public Const FILE_SHARE_WRITE  As Long = &H2  ' ** Allow other programs to write data to the file.

Public Const GENERIC_WRITE     As Single = &H40000000  ' ** Allow the program to write data to the file.
Public Const GENERIC_READ      As Single = &H80000000  ' ** Allow the program to read data from the file.

' ** Family not used.
'Public Const FILE_ATTRIBUTE_READONLY As Long = &H1   ' ** A read-only file.
'Public Const FILE_ATTRIBUTE_HIDDEN   As Long = &H2   ' ** A hidden file, not normally visible to the user.
'Public Const FILE_ATTRIBUTE_SYSTEM   As Long = &H4   ' ** A system file, used exclusively by the operating system.
'Public Const FILE_ATTRIBUTE_ARCHIVE  As Long = &H20  ' ** An archive file (which most files are).
'Public Const FILE_ATTRIBUTE_NORMAL   As Long = &H80  ' ** An attribute-less file (cannot be combined with other attributes).

' ** Family not used.
'Public Const FILE_FLAG_POSIX_SEMANTICS As Single = &H1000000   ' ** Allow file names to be case-sensitive.
'Public Const FILE_FLAG_SEQUENTIAL_SCAN As Single = &H8000000   ' ** Optimize the file cache for sequential access (starting at the beginning and continuing to the end of the file).
'Public Const FILE_FLAG_RANDOM_ACCESS   As Single = &H10000000  ' ** Optimize the file cache for random access (skipping around to various parts of the file).
'Public Const FILE_FLAG_NO_BUFFERING    As Single = &H20000000  ' ** Do not use any buffers or caches. If used, the following things must be done: access to the file must begin at whole number multiples of the disk's sector size; the amounts of data accessed must be a whole number multiple of the disk's sector size; and buffer addresses for I/O operations must be aligned on whole number multiples of the disk's sector size.
'Public Const FILE_FLAG_OVERLAPPED      As Single = &H40000000  ' ** Allow asynchronous I/O; i.e., allow the file to be read from and written to simultaneously. If used, functions that read and write to the file must specify the OVERLAPPED structure identifying the file pointer. Windows 95 does not support overlapped disk files, although Windows NT does.
'Public Const FILE_FLAG_WRITE_THROUGH   As Single = &H80000000  ' ** Bypass any disk cache and instead read and write directly to the file.

'Private Const INVALID_HANDLE_VALUE As Long = -1&
'Private Const ERROR_HANDLE_EOF     As Long = 38&

Private Type COMMTIMEOUTS
  ReadIntervalTimeout As Long
  ReadTotalTimeoutMultiplier As Long
  ReadTotalTimeoutConstant As Long
  WriteTotalTimeoutMultiplier As Long
  WriteTotalTimeoutConstant As Long
End Type

Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long  ' ** Also used by modRegistryFuncs.

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" _
  (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" _
  (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
  ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, _
  ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Declare Function ReadFile Lib "kernel32.dll" _
  (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
  lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

Private Declare Function WriteFile Lib "kernel32.dll" _
  (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
  lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Private Declare Function SetEndOfFile Lib "kernel32.dll" (ByVal hFile As Long) As Long

Private Declare Function GetCommTimeouts Lib "kernel32.dll" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long

' ** This is covered by the Public declaration in modSecurityFunctions.
'Private Declare Function GetLastError Lib "kernel32.dll" () As Long

Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
' **

Public Function BackupToFloppy(strSourcefilename As String, strCallForm As String) As Boolean
' ** Called by:
' **   frmBackupOptions.cmdBackup_Click()
' **
' ** VGC 05/27/08: Strctly speaking, the QuitNow()'s aren't really necessary.
' ** However, I'd just as soon they start fresh in any case.
'NO LONGER USED!

100   On Error GoTo ERRH

        Const THIS_PROC As String = "BackupToFloppy"

        Dim lngSourceFile As Long, lngTargetFile As Long
        Dim bytBuff(4096) As Byte
        Dim lngBytesRead As Long, lngBytesWritten As Long, lngCumulativeBytesWritten As Long
        Dim blnHeaderWritten As Boolean
        Dim strTargetFileName As String
        Dim lngFileSize As Long, lngAverageFileSize As Long
        Dim intNumberDisks As Integer
        Dim strMsg As String
        Dim lngResult As Long
        Dim intX As Integer
        Dim blnContinue As Boolean, blnLoop As Boolean
        Dim blnRetVal As Boolean

        Const FLOPPY_CAPACITY As Long = 1400000

110     blnRetVal = True
120     blnContinue = True

130     gblnBeenToBackup = True
140     lngCumulativeBytesWritten = 0&
150     blnHeaderWritten = False

160   On Error Resume Next
170     gdbsDBLock.Close
180     Set gdbsDBLock = Nothing
190   On Error GoTo ERRH

200     DoEvents
210     lngSourceFile = CreateFile(strSourcefilename, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
        ' ** API Function: Above.
220     DoEvents
230     lngFileSize = GetFileSize(lngSourceFile, 0)  ' ** API Function: Above.

240     If lngFileSize = -1 Then
250       blnContinue = False
260       MsgBox "Error opening data file. Backup aborted!", vbCritical + vbOKOnly, "Backup Failed"
270     Else

280       If lngFileSize < FLOPPY_CAPACITY Then
290         intNumberDisks = 1
300         strMsg = "You will need 1 disk to backup this data." & vbCrLf & vbCrLf & "Do you wish to continue?"
310       Else
320         intNumberDisks = IIf(CInt(lngFileSize / FLOPPY_CAPACITY) < (lngFileSize / FLOPPY_CAPACITY), 1 + _
              CInt(lngFileSize / FLOPPY_CAPACITY), CInt(lngFileSize / FLOPPY_CAPACITY))
330         strMsg = "You will need " & intNumberDisks & " Disks to back up this data." & vbCrLf & vbCrLf & "Do you wish to continue?"
340       End If

350       If MsgBox(strMsg, vbQuestion + vbYesNo, ("Backup Trust Data" & Space(40))) = vbNo Then
360         blnContinue = False
370         MsgBox "Backup aborted!", vbInformation + vbOKOnly, "Backup Failed"
380       Else

390         For intX = 1 To intNumberDisks

400           If MsgBox("Insert a disk labeled:" & vbCrLf & vbCrLf & "  Trust Accountant Data Backup Disk #" & CStr(intX) & _
                  vbCrLf & vbCrLf & "into the A drive and click OK to proceed," & vbCrLf & "or click Cancel to exit.", _
                  vbInformation + vbOKCancel, ("Backup Trust Data" & Space(40))) = vbCancel Then
410             blnContinue = False
420             DoCmd.Hourglass False
430             MsgBox "Backup aborted!", vbInformation + vbOKOnly, "Backup Failed"
440             Exit For
450           Else

460             DoCmd.Hourglass True
470             DoCmd.OpenForm "frmPleaseWait", , , , , , strCallForm & "~" & "Backing Up Data ..."
480             SysCmd acSysCmdSetStatus, "Backup in progress. Please wait . . ."
490             DoEvents

500             strTargetFileName = "A:\Trst" & intX & ".md_"
510             blnLoop = True

520             Do While blnLoop = True And blnContinue = True
530               lngTargetFile = CreateFile(strTargetFileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, 0, 0)
                  ' ** API Function: Above.
540               If lngTargetFile = -1& Then
550                 If MsgBox("Insert a disk labeled:" & vbCrLf & vbCrLf & "  Trust Accountant Data Backup Disk #" & CStr(intX) & _
                        vbCrLf & vbCrLf & "into the A drive and click OK to proceed," & vbCrLf & "or click Cancel to exit.", _
                        vbInformation + vbOKCancel, ("Backup Trust Data" & Space(40))) = vbCancel Then
560                   blnContinue = False: blnLoop = False
570                   DoCmd.Hourglass False
580                   MsgBox "Backup aborted!", vbInformation + vbOKOnly, "Backup Failed"
590                 Else
                      ' ** Will continue to loop as long as lngTargetFile = -1 and they don't click Cancel.
600                 End If
610               Else
620                 blnLoop = False
630               End If
640             Loop

650             If blnContinue = True Then

660               lngCumulativeBytesWritten = 0
670               lngAverageFileSize = CLng(lngFileSize / intNumberDisks)
                  ' ** Initialize the variable to get into the while loop.
                  ' ** After all, it will be replaced by the calls to ReadFile
680               lngBytesRead = 4096
690               Do While (lngCumulativeBytesWritten <= lngAverageFileSize) And (lngBytesRead >= 4096)
700                 If intX = 1 And blnHeaderWritten = False Then
                      ' ** Write the number of disks to the first 2 bytes of
                      ' ** the first file so we know how many disks to restore.
710                   blnHeaderWritten = True
720                   bytBuff(0) = CByte(intNumberDisks)
730                   lngResult = ReadFile(lngSourceFile, bytBuff(1), 4096, lngBytesRead, 0)  ' ** API Function: Above.
740                   lngResult = WriteFile(lngTargetFile, bytBuff(0), lngBytesRead + 1, lngBytesWritten, 0)  ' ** API Function: Above.
750                 Else
760                   lngResult = ReadFile(lngSourceFile, bytBuff(1), 4096, lngBytesRead, 0)  ' ** API Function: Above.
770                   lngResult = WriteFile(lngTargetFile, bytBuff(1), lngBytesRead, lngBytesWritten, 0)  ' ** API Function: Above.
780                 End If
790                 DoEvents
800                 lngCumulativeBytesWritten = lngCumulativeBytesWritten + lngBytesWritten
810               Loop

820               CloseHandle lngTargetFile  ' ** API Function: Above.
830               DoCmd.Hourglass False
840               DoCmd.Close acForm, "frmPleaseWait"

850             Else
860               Exit For
870             End If

880           End If  ' ** blnContinue.

890         Next

900         If blnContinue = True Then
910           CloseHandle lngSourceFile  ' ** API Function: Above.
920           MsgBox "Trust Accountant Data Backup Successfully Completed!", vbInformation + vbOKOnly, "Backup Successful"
930           MsgBox "Connection to Data file closed." & vbCrLf & vbCrLf & _
                "In order to provide a clean link to the data you will need to restart Trust Accountant." & vbCrLf & vbCrLf & _
                "Trust Accountant will now exit.", vbInformation + vbOKOnly, "Trust Accountant Must Exit"
940           QuitNow  ' ** Module Procedure: modStartupFuncs.
950         End If

960       End If  ' ** blnContinue.

970     End If  ' ** lngFileSize.

980     If blnContinue = False Then
990       If IsLoaded("frmBackupRestore", acForm) = True Then  ' ** Module Functions: modFileUtilities.
1000        gblnSetFocus = True
1010        Forms("frmBackupRestore").TimerInterval = 100&
1020      End If
1030    End If

EXITP:
1040    SysCmd acSysCmdClearStatus
1050    CloseHandle lngSourceFile  ' ** API Function: Above.
        ' ** Make sure it's closed; if it is, just returns non-zero value.
1060    BackupToFloppy = blnRetVal
1070    Exit Function

ERRH:
1080    blnRetVal = False
1090    Select Case ERR.Number
        Case 3420  ' ** Object invalid or no longer set.
1100      MsgBox "To backup your data, please restart the Trust Accountant system." & vbCrLf & vbCrLf & _
            "You should not get this message again unless you reinstall the system." & vbCrLf & vbCrLf & _
            "Should you get this message again without reinstalling, please contact Delta Data, Inc.  (" & str(Erl) & ")", _
            vbExclamation + vbOKOnly, "First Use"
1110      QuitNow  ' ** Module Procedure: modStartupFuncs.
1120    Case Else
1130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1140    End Select
1150    Resume EXITP

End Function

Public Function BackupToDrive(strBackupPathFile As String) As Integer
' ** Return Codes:
' **    0  All's well.
' **    1  All's well CD.
' **   -3  Canceled.
' **   -9  Error.
' ** Called by:
' **   frmBackup.cmdBackup_Click()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "BackupToDrive"

        Dim fso As Scripting.FileSystemObject, fsfData As Scripting.File, fsfArchive As Scripting.File
        Dim drv As Scripting.Drive
        Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf As DAO.QueryDef, tdf As DAO.TableDef
        Dim strDataPathFile As String
        Dim strArchivePathFile As String
        Dim lngFileSize As Long, lngRecsArchive As Long
        Dim dblFreeSpace As Double, lngDriveType As Long
        Dim lngX As Long
        Dim intRetVal As Integer, blnContinue As Boolean

1210    intRetVal = -3  ' ** Canceled.
1220    blnContinue = True

        ' ** Set path independent of globals to be sure of accuracy.
1230    strDataPathFile = CurrentBackendPathFile("Ledger")  ' ** Module Function: modFileUtilities.
1240    strArchivePathFile = CurrentBackendPathFile("LedgerArchive")  ' ** Module Function: modFileUtilities.

        ' * Test for space on drive.
1250    Set fso = New FileSystemObject
1260    Set fsfData = fso.GetFile(strDataPathFile)

1270    Set fsfArchive = fso.GetFile(strArchivePathFile)
1280    lngFileSize = fsfData.Size + fsfArchive.Size

1290    Set drv = fso.GetDrive(Left(strBackupPathFile, 1))
1300    lngDriveType = GetTypeDrive(drv.DriveLetter)  ' ** Module Function: modFileUtilities.

1310    If lngDriveType = vbxDriveCDROM Then
          ' ** This will invoke the built-in CD Writing Wizard, which must be manually prompted to finish the copy process.
'1320      blnContinue = CopyToCD(strDataPathFile, strArchivePathFile, strBackupPathFile, lngFileSize)  ' ** Module Function: modCDBurnFuncs.
1320      Select Case blnContinue
          Case True
1330        intRetVal = 1  ' ** All's well CD.
1340      Case False
1350        intRetVal = -9  ' ** Error.
1360      End Select
1370    Else

1380  On Error Resume Next
1390      dblFreeSpace = drv.FreeSpace
1400      If ERR.Number <> 0 Then
            ' ** -2147024895  Method 'FreeSpace' of object 'IDrive' failed.
1410        If lngDriveType = vbxDriveCDROM Or lngDriveType = vbxDriveRAMDisk Or lngDriveType = vbxDriveRemovable Then
1420  On Error GoTo ERRH
1430          DoCmd.Hourglass False
1440          MsgBox "The specified drive media has not been formatted," & vbCrLf & _
                "and cannot be written to.", vbInformation + vbOKOnly, "Drive Not Formatted"
1450          blnContinue = False
1460          intRetVal = -9
1470        Else
1480          DoCmd.Hourglass False
1490          MsgBox "There is a problem with the specified drive, and its free space cannot be read." & vbCrLf & vbCrLf & _
                "Error: " & CStr(ERR.Number) & vbCrLf & _
                "Description: " & ERR.description, vbInformation + vbOKOnly, "Error Reading Drive"
1500  On Error GoTo ERRH
1510          blnContinue = False
1520          intRetVal = -9
1530        End If
1540      Else
1550  On Error GoTo ERRH
1560      End If

1570      If blnContinue = True Then

1580        If lngFileSize > drv.FreeSpace Then
1590          DoCmd.Hourglass False
1600          MsgBox "There isn't enough free space on the drive to save the backup file.", vbCritical + vbOKOnly, "Not Enough Space"
1610          blnContinue = False
1620          intRetVal = -9
1630          If IsLoaded("frmBackupRestore", acForm) = True Then  ' ** Module Functions: modFileUtilities.
1640            gblnSetFocus = True
1650            Forms("frmBackupRestore").TimerInterval = 100&
1660          End If
1670        Else

1680          lngRecsArchive = DCount("*", "LedgerArchive")

1690          If TableExists(gstrTable_LedgerArchive) = True Then  ' ** Module Function: modFileUtilities.
1700            DoCmd.DeleteObject acTable, gstrTable_LedgerArchive
1710          End If
1720          CurrentDb.TableDefs.Refresh
1730          CurrentDb.TableDefs.Refresh

              ' ** Copy file.
1740          fsfData.Copy strBackupPathFile

              ' ** Add the LedgerArchive.
1750  On Error Resume Next
1760          Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
1770          If ERR.Number <> 0 Then
1780  On Error GoTo ERRH
1790  On Error Resume Next
1800            Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
1810            If ERR.Number <> 0 Then
1820  On Error GoTo ERRH
1830  On Error Resume Next
1840              Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
1850              If ERR.Number <> 0 Then
1860  On Error GoTo ERRH
1870  On Error Resume Next
1880                Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
1890                If ERR.Number <> 0 Then
1900  On Error GoTo ERRH
1910  On Error Resume Next
1920                  Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
1930                  If ERR.Number <> 0 Then
1940  On Error GoTo ERRH
1950  On Error Resume Next
1960                    Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
1970                    If ERR.Number <> 0 Then
1980  On Error GoTo ERRH
1990  On Error Resume Next
2000                      Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
2010  On Error GoTo ERRH
2020                    Else
2030  On Error GoTo ERRH
2040                    End If
2050                  Else
2060  On Error GoTo ERRH
2070                  End If
2080                Else
2090  On Error GoTo ERRH
2100                End If
2110              Else
2120  On Error GoTo ERRH
2130              End If
2140            Else
2150  On Error GoTo ERRH
2160            End If
2170          Else
2180  On Error GoTo ERRH
2190          End If

2200          With wrk
2210            Set dbs = .OpenDatabase(strBackupPathFile, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
2220            With dbs

2230              For Each tdf In .TableDefs
2240                With tdf
2250                  If .Name = gstrTable_LedgerArchive Then
2260                    dbs.TableDefs.Delete gstrTable_LedgerArchive
2270                    Exit For
2280                  End If
2290                End With
2300              Next
2310              .TableDefs.Refresh
2320              .TableDefs.Refresh

                  ' ** Data-Definition: Create table LedgerArchive_Backup.
                  ' ** qryBackupRestore_03_01 - qryBackupRestore_03_13.
2330              For lngX = 1& To 13&
2340                .Execute CurrentDb.QueryDefs("qryBackupRestore_03_" & Right("00" & CStr(lngX), 2)).SQL
2350              Next
2360              .TableDefs.Refresh
2370              .TableDefs.Refresh

2380              .Close
2390            End With  ' ** dbs.
2400            .Close
2410          End With  ' ** wrk.

2420          If lngRecsArchive > 0& Then
2430            DoCmd.TransferDatabase acLink, "Microsoft Access", strBackupPathFile, acTable, _
                  gstrTable_LedgerArchive, gstrTable_LedgerArchive
2440            CurrentDb.TableDefs.Refresh
2450            CurrentDb.TableDefs.Refresh
                ' ** Append LedgerArchive to LedgerArchive_Backup.
2460            Set qdf = CurrentDb.QueryDefs("qryBackupRestore_02")
2470            qdf.Execute
2480            DoEvents
2490            DoCmd.DeleteObject acTable, gstrTable_LedgerArchive
2500            CurrentDb.TableDefs.Refresh
2510            CurrentDb.TableDefs.Refresh
2520          End If

2530          intRetVal = 0

2540        End If  ' ** lngFileSize.

2550      End If  ' ** blnContinue.

2560    End If  ' ** lngDriveType.

EXITP:
2570    Set tdf = Nothing
2580    Set qdf = Nothing
2590    Set dbs = Nothing
2600    Set wrk = Nothing
2610    Set drv = Nothing
2620    Set fsfData = Nothing
2630    Set fsfArchive = Nothing
2640    Set fso = Nothing
2650    BackupToDrive = intRetVal
2660    Exit Function

ERRH:
2670    DoCmd.Hourglass False
2680    intRetVal = -9
2690    Select Case ERR.Number
        Case Else
2700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2710    End Select
2720    Resume EXITP

End Function

Public Function RestoreFromFloppy(strCallForm As String) As Boolean
'THIS IS NO LONGER USED!
'WILL BE REPLACED BY A CD BURN ROUTINE!
' ** THIS COMPLETELY REPLACES THE CURRENT FILE WITH THE BACKED-UP FILE!
' ** If the backed-up file is from an earlier version, all holy-heck could break out.
' ** Called by:
' **   frmRestoreOptions.cmdRestore_Click()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "RestoreFromFloppy"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field, rst As DAO.Recordset
        Dim lngSourceFile As Long, lngTargetFile As Long
        Dim bytBuff(4096) As Byte, lngBytesRead As Long, lngBytesWritten As Long
        Dim lngCumulativeBytesWritten As Long
        Dim lngResult As Long
        Dim strSourcefilename As String, strDataPathFile_Current As String, strArchPathFile_Current As String
        Dim lngMaxFileSize As Long
        Dim intX As Integer
        Dim intNumberDisks As Integer
        Dim blnHeaderRead As Boolean
        Dim strVer_Current As String, strVer_Backup As String
        Dim blnFound As Boolean, blnLoop As Boolean, blnQuit As Boolean
        Dim arr_varTmp00() As Variant, strTmp01 As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

2810    blnRetVal = True
2820    blnLoop = False
2830    blnQuit = False

2840    DoCmd.Hourglass True

2850    lngCumulativeBytesWritten = 0&
2860    blnHeaderRead = False

2870    strDataPathFile_Current = gstrTrustDataLocation & gstrFile_DataName
2880    strArchPathFile_Current = gstrTrustDataLocation & gstrFile_ArchDataName

2890    Application.SysCmd acSysCmdSetStatus, "Restore in progress.  Please wait..."

        ' ** This process will create a backup of the existing file,
        ' ** then delete the file if they choose to proceed.
        ' ** Otherwise it will just exit.
2900    If FileExists_Local(strDataPathFile_Current) = True Then  ' ** Function: Below.

2910      DoCmd.Hourglass False
2920      If MsgBox("The Trust Accountant data file already exists." & vbCrLf & vbCrLf & _
              "If you choose to proceed, this file will be replaced." & vbCrLf & vbCrLf & _
              "Do you wish to continue?", vbQuestion + vbYesNo, ("Restore Trust Data" & Space(40))) = vbYes Then

2930        DoCmd.Hourglass True

2940  On Error Resume Next
2950        gdbsDBLock.Close
2960        Set gdbsDBLock = Nothing
2970  On Error GoTo ERRH

            ' ** Check for an existing TrustDta.BAK
2980        strTmp01 = Left(strDataPathFile_Current, Len(strDataPathFile_Current) - 3) & "BAK"
2990        If FileExists(strTmp01) = True Then
3000          Kill strTmp01
3010        End If

            ' ** Copy the current TrustDta.mdb to TrustDta.BAK.
3020        lngResult = CopyFile(strDataPathFile_Current, strTmp01, False)
            ' ** API Function: Above.

            ' ** Create the new one, overwriting the current TrustDta.mdb in the process.
3030        lngTargetFile = CreateFile(strDataPathFile_Current, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
              ByVal 0&, CREATE_ALWAYS, 0, 0)  ' ** API Function: Above.

3040      Else
3050        blnRetVal = False
3060      End If  ' ** Continue restore.
3070    Else
3080      blnRetVal = False
3090      DoCmd.Hourglass False
3100      MsgBox "Trust Accountant data file not found!", vbExclamation + vbOKOnly, "File Not Found"
3110    End If  ' ** Current data found.

3120    If blnRetVal = True Then

3130      blnLoop = True

          ' ** This loop just makes sure Disk 1 is in the drive and they're ready to proceed.
3140      Do While blnLoop = True And blnRetVal = True

3150        If MsgBox("Insert the disk labeled:" & vbCrLf & vbCrLf & "  Trust Accountant Data Backup Disk #1" & _
                vbCrLf & vbCrLf & "into the A drive and click OK to proceed," & vbCrLf & "or click Cancel to exit.", _
                vbInformation + vbOKCancel, ("Restore Trust Data" & Space(40))) = vbCancel Then
              ' ** Cancel is used to say both all disks have finished and Cancel the restore.
              ' ** A blnQuit = False should indicate that a restore hasn't taken place.
              'THIS MSGBOX() ONLY LOOKS FOR DISK #1, BUT DOESN'T YET DO ANYTHING WITH IT.
              'ANY CANCEL AT THIS POINT SEEMS TO MEAN NO RESTORE TOOK PLACE!
              'IF THEY CANCEL, AND THE .BAK FILE IS THERE, RENAME IT BACK TO WHERE IT WAS.
3160          If FileExists_Local(Left(strDataPathFile_Current, Len(strDataPathFile_Current) - 3) & "BAK") Then
3170            blnRetVal = False: blnLoop = False
3180            CloseHandle lngTargetFile  ' ** API Function: Above.
3190            lngResult = CopyFile(Left(strDataPathFile_Current, Len(strDataPathFile_Current) - 3) & "BAK", _
                  strDataPathFile_Current, False)
                ' ** API Function: Above.
3200            Kill Left(strDataPathFile_Current, Len(strDataPathFile_Current) - 3) & "BAK"
3210            MsgBox "Trust Accountant Data not restored." & vbCrLf & vbCrLf & _
                  "Trust Accountant will now exit." & vbCrLf & vbCrLf & _
                  "In order to provide a clean link to the data," & vbCrLf & "you will need to restart Trust Accountant.", _
                  vbInformation + vbOKOnly, ("Restore Canceled / Trust Accountant Must Exit" & Space(10))
3220            blnQuit = True
3230          Else
3240            blnRetVal = False: blnLoop = False
3250          End If
3260        End If
3270        DoEvents

            'WHEN FIRST BEGINNING, WE'LL CONTINUE AFTER THE MSGBOX(), ABOVE.
3280        If blnRetVal = True Then
3290          strSourcefilename = "A:\Trst1.md_"
3300          If FileExists_Local(strSourcefilename) = False Then
                ' ** File doesn't exist.
3310            blnLoop = True
                'IF THE FILE WASN'T FOUND, WE LOOP BACK TO THE MSGBOX(), ABOVE.
3320          Else
3330            blnLoop = False
                'IF THE FIRST FILE IS ON THE FLOPPY, WE EXIT THIS LOOP AND CONTINUE BELOW.
3340          End If
3350          DoEvents
3360        End If

3370      Loop  ' ** While blnRetVal = True and blnLoop = True.

3380    End If  ' ** blnRetVal.

3390    If blnRetVal = True Then

3400      DoCmd.Hourglass True
3410      DoCmd.OpenForm "frmPleaseWait", , , , , , strCallForm & "~" & "Restoring Data ..."
3420      DoEvents

3430      lngSourceFile = CreateFile(strSourcefilename, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
            ByVal 0&, OPEN_EXISTING, 0, 0)  ' ** API Function: Above.

          ' ** Doesn't this just get the size of the file on Disk 1 alone?
3440      lngResult = ReadFile(lngSourceFile, bytBuff(0), 1, lngBytesRead, 0)  ' ** API Function: Above.
3450      intNumberDisks = CInt(bytBuff(0))
3460      lngMaxFileSize = intNumberDisks * 1457664  ' ** Capacity of a 3.5 inch diskette.

3470      lngCumulativeBytesWritten = 0
3480      Do While lngBytesRead > 0
3490        lngResult = ReadFile(lngSourceFile, bytBuff(1), 4096, lngBytesRead, 0)  ' ** API Function: Above.
3500        lngResult = WriteFile(lngTargetFile, bytBuff(1), lngBytesRead, lngBytesWritten, 0)  ' ** API Function: Above.
3510        DoEvents
3520        lngCumulativeBytesWritten = lngCumulativeBytesWritten + lngBytesWritten
3530      Loop
3540      DoCmd.Hourglass False
3550      DoCmd.Close acForm, "frmPleaseWait"
3560      CloseHandle lngSourceFile  ' ** API Function: Above.

3570      If intNumberDisks > 1 Then
            ' ** At this point, Disk 1 is already in and copied.

3580        For intX = 2 To intNumberDisks

3590          blnLoop = True

3600          Do While blnLoop = True And blnRetVal = True
3610            DoEvents
3620            If MsgBox("Insert the disk labeled:" & vbCrLf & vbCrLf & "  Trust Accountant Data Backup Disk #" & CStr(intX) & _
                    vbCrLf & vbCrLf & "into the A drive and click OK to proceed," & vbCrLf & "or click Cancel to exit.", _
                    vbInformation + vbOKCancel, ("Restore Trust Data" & Space(40))) = vbCancel Then
                  'AGAIN, ANY CANCEL AT THIS POINT SEEMS TO MEAN NO RESTORE TOOK PLACE!
                  'IF THEY CANCEL, AND THE .BAK FILE IS THERE, RENAME IT BACK TO WHERE IT WAS.
3630              If FileExists_Local(Left(strDataPathFile_Current, Len(strDataPathFile_Current) - 3) & "BAK") Then
3640                blnRetVal = False: blnLoop = False
3650                CloseHandle lngTargetFile  ' ** API Function: Above.
3660                lngResult = CopyFile(Left(strDataPathFile_Current, Len(strDataPathFile_Current) - 3) & "BAK", _
                      strDataPathFile_Current, False)
                    ' ** API Function: Above.
3670                Kill Left(strDataPathFile_Current, Len(strDataPathFile_Current) - 3) & "BAK"
3680                MsgBox "Trust Accountant Data not restored." & vbCrLf & vbCrLf & _
                      "Trust Accountant will now exit." & vbCrLf & vbCrLf & _
                      "In order to provide a clean link to the data," & vbCrLf & "you will need to restart Trust Accountant.", _
                      vbInformation + vbOKOnly, ("Restore Canceled / Trust Accountant Must Exit" & Space(10))
3690                blnQuit = True
3700              End If
3710            End If

3720            If blnRetVal = True Then
3730              strSourcefilename = "A:\Trst" & intX & ".md_"
3740              If FileExists_Local(strSourcefilename) = False Then
                    ' ** File doesn't exist.
3750                blnLoop = True
3760              Else
3770                blnLoop = False
3780              End If
3790            End If

3800          Loop  ' ** While blnLoop = True And blnRetVal = True.

3810          If blnRetVal = True Then

3820            DoCmd.Hourglass True
3830            DoEvents
3840            DoCmd.OpenForm "frmPleaseWait", , , , , , strCallForm & "~" & "Restoring Data ..."

3850            lngSourceFile = CreateFile(strSourcefilename, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
                  OPEN_EXISTING, 0, 0)  ' ** API Function: Above.

                ' ** Initial Read.
3860            lngResult = ReadFile(lngSourceFile, bytBuff(1), 4096, lngBytesRead, 0)  ' ** API Function: Above.
3870            lngResult = WriteFile(lngTargetFile, bytBuff(1), lngBytesRead, lngBytesWritten, 0)  ' ** API Function: Above.
3880            Do While lngBytesRead > 0
3890              DoEvents
3900              lngResult = ReadFile(lngSourceFile, bytBuff(1), 4096, lngBytesRead, 0)  ' ** API Function: Above.
3910              lngResult = WriteFile(lngTargetFile, bytBuff(1), lngBytesRead, lngBytesWritten, 0)  ' ** API Function: Above.
3920              lngCumulativeBytesWritten = lngCumulativeBytesWritten + lngBytesWritten
3930            Loop
3940            CloseHandle lngSourceFile  ' ** API Function: Above.
3950            DoCmd.Hourglass False
3960            DoCmd.Close acForm, "frmPleaseWait"
3970            DoEvents

3980          Else
3990            Exit For
4000          End If  ' ** blnRetVal.

4010        Next  ' ** intX.

4020      End If  ' ** intNumberDisks.

4030    End If  ' ** blnRetVal.

4040    If blnRetVal = True Then

4050      lngResult = SetEndOfFile(lngTargetFile)  ' ** API Function: Above.
4060      CloseHandle lngTargetFile  ' ** API Function: Above.

          'INTERCEPT HERE, BEFORE LedgerArchive IS RESTORED, AND CHECK THE VERSION.
4070      blnFound = False

          ' ** Open Backed-up file.
4080      Set dbs = DBEngine.Workspaces(0).OpenDatabase(gstrTrustDataLocation & gstrFile_DataName)
4090      With dbs
4100        For Each tdf In .TableDefs
4110          With tdf
4120            If .Name = "m_VD" Then
4130              blnFound = True
4140              Exit For
4150            End If
4160          End With
4170        Next
4180        If blnFound = True Then
4190          Set rst = .OpenRecordset("m_VD", dbOpenDynaset, dbReadOnly)
4200          With rst
4210            If .BOF = True And .EOF = True Then
4220              blnFound = False
4230              blnRetVal = False
4240            Else
4250              .MoveFirst
4260              strVer_Backup = CStr(![vd_MAIN]) & "." & CStr(![vd_MINOR]) & "." & CStr(![vd_REVISION])
4270            End If
4280            .Close
4290          End With
4300        End If
4310        .Close
4320      End With
4330      Set tdf = Nothing
4340      Set rst = Nothing
4350      Set dbs = Nothing

4360      If blnFound = False Then
4370        blnRetVal = False
4380        DoCmd.Hourglass False
4390        MsgBox "The version of Trust Accountant that created the backup could not be ascertained!", _
              vbExclamation + vbOKOnly, "Backup Version Unknown"
4400      End If
4410    End If  ' ** Continue restore.

4420    If blnRetVal = True Then

          ' ** If it's like 2.1.6, make it 2.1.60.
4430      If Left(Right(strVer_Backup, 2), 1) = "." Then strVer_Backup = strVer_Backup & "0"
          ' ** Get this version.
4440      strVer_Current = AppVersion_Get2  ' ** Module Function: modAppVersionFuncs.
          ' ** If it's a Demo (Rich, for instance), trim the 'd' (backup won't have a 'd' because it comes from numeric table fields).
4450      If Right(strVer_Current, 1) = "d" Then strVer_Current = Left(strVer_Current, (Len(strVer_Current) - 1))
          ' ** If it's like 2.1.7, make it 2.1.70.
4460      If Left(Right(strVer_Current, 2), 1) = "." Then strVer_Current = strVer_Current & "0"

          ' ** Compare the backup version with this, current version.
4470      If strVer_Backup = strVer_Current Then
            ' ** Backed-up file is same version as current frontend.

            ' ** Before we close the whole procedure, check if there is a
            ' ** LedgerArchive_Backup table and restore that to its own mdb.
4480        If RestoreLedgerArchive(strDataPathFile_Current) = "fail" Then  ' ** Function: Below.
              ' ** If Err.Number = 3010 Then  ' ** Table '|' already exists.
4490          MsgBox "Trust Accountant Archive Data Restore has failed." & vbCrLf & vbCrLf & _
                "Consult Delta Data, Inc., for further information.", vbCritical + vbOKOnly, "Restore Failed"
4500        Else
              'THIS APPEARS TO BE THE ONLY SUCCESSFUL RESTORE!
              ' ** It looks like restoring from floppy creates TrustDta.mdb from scratch!
              ' ** If that's true, it should include all relationships and user-defined properties.

              ' ** Given recent changes to the PostingDate and CompanyInformation
              ' ** tables, check and replace if necessary.
4510          Set dbs = DBEngine.Workspaces(0).OpenDatabase(gstrTrustDataLocation & gstrFile_DataName)
4520          With dbs
4530            blnFound = False
4540            For Each tdf In .TableDefs
4550              With tdf
4560                If .Name = "PostingDate" Then
                      ' ** The original PostingDate table had only 1 field.
4570                  If .Fields.Count = 3 Then
4580                    blnFound = True
4590                  End If
4600                  Exit For
4610                End If
4620              End With
4630            Next
4640            If blnFound = False Then
4650              strTmp01 = vbNullString
4660              Set rst = .OpenRecordset("PostingDate", dbOpenDynaset, dbReadOnly)
4670              With rst
4680                If .BOF = True And .EOF = True Then
                      ' ** No need to copy the date.
4690                Else
4700                  .MoveFirst
4710                  If IsNull(.Fields(0)) = False Then
4720                    strTmp01 = Format(.Fields(0), "mm/dd/yyyy")
4730                  Else
                        ' ** No need to copy the date.
4740                  End If
4750                End If
4760                .Close
4770              End With
                  ' ** Update the template with their existing Posting Date.
4780              If strTmp01 <> vbNullString Then
4790                Set rst = CurrentDb.OpenRecordset("tblTemplate_PostingDate", dbOpenDynaset, dbConsistent)
4800                With rst
4810                  .MoveFirst
4820                  If ![Posting_Date] <> CDate(strTmp01) Then
4830                    .Edit
4840                    ![Posting_Date] = CDate(strTmp01)
4850                    .Update
4860                  End If
4870                  .Close
4880                End With
4890              End If
                  ' ** Delete their existing PostingDate table.
4900              .TableDefs.Delete "PostingDate"
4910              .TableDefs.Refresh
                  ' ** Transfer the new PostingDate template.
4920              DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                    acTable, "tblTemplate_PostingDate", "PostingDate", False  ' ** Copy with data.
4930              .TableDefs.Refresh
4940            End If
4950            blnFound = False
4960            For Each tdf In .TableDefs
4970              With tdf
4980                If .Name = "CompanyInformation" Then
4990                  If .Fields.Count = 16 Then
                        ' ** Old CompanyInformation tables had 7, 12, or 13 fields.
5000                    blnFound = True
5010                  End If
5020                  Exit For
5030                End If
5040              End With
5050            Next
5060            If blnFound = False Then
5070              ReDim arr_varTmp00(1, .TableDefs("CompanyInformation").Fields.Count)  ' ** I think we're only using this up to 11.
5080              For lngX = 0& To UBound(arr_varTmp00, 2)
5090                arr_varTmp00(0, lngX) = vbNullString
5100              Next
5110              Set rst = .OpenRecordset("CompanyInformation", dbOpenDynaset, dbReadOnly)
5120              With rst
5130                If .BOF = True And .EOF = True Then
                      ' ** No need to copy the data.
5140                Else
5150                  .MoveFirst
5160                  For lngX = 0& To (.Fields.Count - 1&)
5170                    arr_varTmp00(0, lngX) = .Fields(lngX).Name
5180                    Select Case .Fields(lngX).Type
                        Case dbText
5190                      If IsNull(.Fields(lngX)) = False Then
5200                        arr_varTmp00(1, lngX) = .Fields(lngX)
5210                      Else
5220                        arr_varTmp00(1, lngX) = vbNullString
5230                      End If
5240                    Case dbBoolean
5250                      arr_varTmp00(1, lngX) = CBool(.Fields(lngX))
5260                    End Select
5270                  Next
5280                End If
5290                .Close
5300              End With
                  ' ** Update the template with their existing data.
5310              If arr_varTmp00(0, 0) <> vbNullString Then
5320                Set rst = CurrentDb.OpenRecordset("tblTemplate_CompanyInformation", dbOpenDynaset, dbConsistent)
5330                With rst
5340                  .AddNew
5350                  For lngX = 0& To UBound(arr_varTmp00, 2)
5360                    If arr_varTmp00(0, lngX) <> vbNullString Then
5370                      Select Case .Fields(arr_varTmp00(0, lngX)).Type
                          Case dbText
5380                        If arr_varTmp00(1, lngX) <> vbNullString Then
5390                          .Fields(arr_varTmp00(0, lngX)) = arr_varTmp00(1, lngX)
5400                        End If
5410                      Case dbBoolean
5420                        .Fields(arr_varTmp00(0, lngX)) = arr_varTmp00(1, lngX)
5430                      End Select
5440                    End If
5450                  Next
5460                  .Update
5470                End With
5480              End If
                  ' ** Delete their existing CompanyInformation table.
5490              .TableDefs.Delete "CompanyInformation"
5500              .TableDefs.Refresh
                  ' ** Transfer the new CompanyInformation template.
5510              DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                    acTable, "tblTemplate_CompanyInformation", "CompanyInformation", False  ' ** Copy with data.
5520              .TableDefs.Refresh
5530            End If
5540            .Close
5550          End With

5560          MsgBox "Trust Accountant Data Restore Successfully Completed!", vbInformation + vbOKOnly, "Restore Successful"

5570        End If

5580        MsgBox "Trust Accountant will now exit." & vbCrLf & vbCrLf & _
              "In order to provide a clean link to the data," & vbCrLf & _
              "you will need to restart Trust Accountant.", vbInformation + vbOKOnly, "Trust Accountant Must Exit"
5590        blnQuit = True

5600      Else
            ' ** Backed-up file is different version from current frontend.

            ' ** Rename the currently restored file to gstrFile_RestoreDataName (TrstRest.mdb).
5610        Name strDataPathFile_Current As (gstrTrustDataLocation & gstrFile_RestoreDataName)

            ' ** Rename the original data file, now BAK, to gstrFile_DataName.
5620        strTmp01 = (Left(strDataPathFile_Current, (Len(strDataPathFile_Current) - 3)) & "BAK")
5630        Name strTmp01 As strDataPathFile_Current

            ' ** The original archive wasn't renamed, so it's still intact.

5640        blnRetVal = RestoreEarlierVersion(strDataPathFile_Current, strArchPathFile_Current, _
              strVer_Current, strVer_Backup, blnQuit)  ' ** Function: Below.

5650      End If  ' ** Current vs. earlier version.

5660    End If  ' ** blnRetVal.

5670    If blnRetVal = False And blnQuit = False Then
5680      MsgBox "Trust Accountant data not restored.", vbInformation + vbOKOnly, "Restore Canceled"
5690      If IsLoaded("frmBackupRestore", acForm) = True Then  ' ** Module Functions: modFileUtilities.
5700        gblnSetFocus = True
5710        Forms(strCallForm).TimerInterval = 100&
5720      End If
5730    End If

EXITP:
5740    Set fld = Nothing
5750    Set tdf = Nothing
5760    Set rst = Nothing
5770    Set dbs = Nothing
5780    Application.SysCmd acSysCmdClearStatus
5790    If blnQuit = True Then QuitNow  ' ** Module Procedure: modStartupFuncs.
5800    RestoreFromFloppy = blnRetVal
5810    Exit Function

ERRH:
5820    blnRetVal = False
5830    DoCmd.Hourglass False
5840    Select Case ERR.Number
        Case Else
5850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5860    End Select
5870    Resume EXITP

End Function

Public Function RestoreFromDrive() As Boolean
' ** THIS COMPLETELY REPLACES THE CURRENT FILE WITH THE BACKED-UP FILE!
' ** If the backed-up file is from an earlier version, all holy-heck could break out.
' ** Called by:
' **   frmRestoreOptions.cmdRestore_Click()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "RestoreFromDrive"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim fso As Scripting.FileSystemObject, fsf As Scripting.File
        Dim lngResult As Long
        Dim strDataPathFile_Current As String, strDataPathFile_Backup As String, strArchPathFile_Current As String
        Dim strVer_Current As String, strVer_Backup As String
        Dim strFilter As String
        Dim blnFound As Boolean, blnQuit As Boolean
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

        ' ** Array: arr_varLink().
        'Const L_DID  As Integer = 0
        'Const L_DNAM As Integer = 1
        'Const L_TID  As Integer = 2
        'Const L_TNAM As Integer = 3
        'Const L_ID   As Integer = 4
        'Const L_NAM  As Integer = 5
        'Const L_SNAM As Integer = 6

5910    blnRetVal = True
5920    blnQuit = False

5930    DoCmd.Hourglass True

5940    strDataPathFile_Current = gstrTrustDataLocation & gstrFile_DataName
5950    strArchPathFile_Current = gstrTrustDataLocation & gstrFile_ArchDataName

5960    Application.SysCmd acSysCmdSetStatus, "Restore in progress.  Please wait..."

        ' ** Get name of backed up file.
5970    strFilter = CreateWindowsFilterString("Backup Files (*.md_)", "*.md_")  ' ** Module Function: modBrowseFilesAndFolders.

5980    DoCmd.Hourglass False
5990    strDataPathFile_Backup = GetOpenFileSIS(strFilter, gstrTrustDataLocation, "Find Backed Up Data File")  ' ** Module Function: modBrowseFilesAndFolders.
6000    If strDataPathFile_Backup = vbNullString Then
6010      blnRetVal = False
6020      MsgBox "You didn't select a backed up file.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
6030    Else

6040      DoCmd.Hourglass True

          ' ** This process will create a backup of the existing file,
          ' ** then delete the file if they choose to proceed.
          ' ** Otherwise it will just exit.
6050      If FileExists_Local(strDataPathFile_Current) Then
6060        DoCmd.Hourglass False
6070        If MsgBox("The Trust Accountant data file already exists." & vbCrLf & vbCrLf & _
                "If you choose to proceed, this file will be replaced." & vbCrLf & vbCrLf & _
                "Do you wish to continue?", vbQuestion + vbYesNo, "Restore Trust Data") = vbYes Then

6080          DoCmd.Hourglass True

6090  On Error Resume Next
6100          gdbsDBLock.Close
6110          Set gdbsDBLock = Nothing
6120  On Error GoTo ERRH

              ' ** Copy it to the data directory, renaming on the way.
6130          Set fso = New FileSystemObject
6140          Set fsf = fso.GetFile(strDataPathFile_Backup)
6150          fsf.Copy (gstrTrustDataLocation & gstrFile_RestoreDataName)
6160          Set fsf = Nothing

              ' ** Unset read only attribute, if set.
6170          Set fso = New FileSystemObject
6180          Set fsf = fso.GetFile(gstrTrustDataLocation & gstrFile_RestoreDataName)
6190          If fsf.Attributes And fsAttrReadOnly Then
6200            fsf.Attributes = fsf.Attributes - fsAttrReadOnly
6210          End If
6220          Set fsf = Nothing

              ' ** FsAttributes enumeration:
              ' **      0  fsAttrNormal      Normal file. No attributes are set.
              ' **      1  fsAttrReadOnly    Read-only file. Attribute is read/write.
              ' **      2  fsAttrHidden      Hidden file. Attribute is read/write.
              ' **      4  fsAttrSystem      System file. Attribute is read/write.
              ' **      8  fsAttrVolume      Disk drive volume label. Attribute is read-only.
              ' **     16  fsAttrDirectory   Folder or directory. Attribute is read-only.
              ' **     32  fsAttrArchive     File has changed since last backup. Attribute is read/write.
              ' **   1024  fsAttrAlias       Link or shortcut. Attribute is read-only.
              ' **   2048  fsAttrCompressed  Compressed file. Attribute is read-only.

6230          blnFound = False: strVer_Backup = vbNullString

              ' ** Open Backed-up file.
6240          Set dbs = DBEngine.Workspaces(0).OpenDatabase(gstrTrustDataLocation & gstrFile_RestoreDataName)
6250          With dbs
6260            For Each tdf In .TableDefs
6270              With tdf
6280                If .Name = "m_VD" Then
6290                  blnFound = True
6300                  Exit For
6310                End If
6320              End With
6330            Next
6340            If blnFound = True Then
6350              Set rst = .OpenRecordset("m_VD", dbOpenDynaset, dbReadOnly)
6360              With rst
6370                If .BOF = True And .EOF = True Then
6380                  blnFound = False
6390                  blnRetVal = False
6400                Else
6410                  .MoveFirst
6420                  strVer_Backup = ![vd_MAIN] & "." & ![vd_MINOR] & "." & ![vd_REVISION]
6430                End If
6440                .Close
6450              End With
6460            Else
6470              blnRetVal = False
6480            End If
6490            .Close
6500          End With
6510          Set tdf = Nothing
6520          Set rst = Nothing
6530          Set dbs = Nothing

6540          If blnFound = False Then
6550            blnRetVal = False
6560            DoCmd.Hourglass False
6570            MsgBox "The version of Trust Accountant that created the backup could not be ascertained!" & vbCrLf & _
                  "The restore cannot continue.", vbExclamation + vbOKOnly, "Backup Version Unknown"
6580          End If
6590        End If  ' ** Continue restore.
6600      Else
            ' ** Do we want to make this an option at startup if their data is lost, but they still have a backup?
6610        blnRetVal = False
6620        DoCmd.Hourglass False
6630        MsgBox "Trust Accountant data file not found!", vbExclamation + vbOKOnly, "File Not Found"
6640      End If  ' ** Current data found.
6650    End If  ' ** Path to backup selected.

6660    If blnRetVal = True Then

          ' ** If it's like 2.1.6, make it 2.1.60.
6670      If Left(Right(strVer_Backup, 2), 1) = "." Then strVer_Backup = strVer_Backup & "0"
          ' ** Get this version.
6680      strVer_Current = AppVersion_Get2  ' ** Module Function: modAppVersionFuncs.
          ' ** If it's a Demo (Rich, for instance), trim the 'd' (backup won't have a 'd' because it comes from numeric table fields).
6690      If Right(strVer_Current, 1) = "d" Then strVer_Current = Left(strVer_Current, (Len(strVer_Current) - 1))
          ' ** If it's like 2.1.7, make it 2.1.70.
6700      If Left(Right(strVer_Current, 2), 1) = "." Then strVer_Current = strVer_Current & "0"

          ' ** Compare the backup version with this, current version.
6710      If strVer_Backup = strVer_Current Then
            ' ** Backed-up file is same version as current frontend.

            ' ** This changes the extension of the current TrustDta.mdb to .BAK.
6720        strTmp01 = (Left(strDataPathFile_Current, (Len(strDataPathFile_Current) - 3)) & "BAK")
6730        If FileExists(strTmp01) = True Then  ' ** Module Function: modFileUtilities.
              ' ** If there's another .BAK there, delete it!
6740          Kill strTmp01
6750        End If
6760        lngResult = CopyFile(strDataPathFile_Current, strTmp01, False)  ' ** API Function: Above.

            ' ** Also delete any TrstArch.BAK found.
6770        strTmp01 = Parse_Path(strTmp01) & LNK_SEP & gstrFile_ArchDataName  ' ** Module Function: modFileUtilities.
6780        strTmp01 = (Left(strTmp01, (Len(strTmp01) - 3)) & "BAK")
6790        If FileExists(strTmp01) = True Then  ' ** Module Function: modFileUtilities.
6800          Kill strTmp01
6810        End If
6820        strTmp01 = vbNullString

            ' ** Delete current data file.
6830        If FileExists(strDataPathFile_Current) = True Then  ' ** Module Function: modFileUtilities.
6840          blnRetVal = Tbl_DelAllLinks(gstrFile_DataName)  ' ** Function: Below.
6850          DoEvents
6860          Select Case blnRetVal
              Case True
6870  On Error Resume Next
6880            Kill strDataPathFile_Current
6890  On Error GoTo ERRH
6900          Case False
6910            Beep
6920            MsgBox "Problem deleting links."
6930          End Select
6940        End If

6950        If blnRetVal = True Then

              ' ** Copy file.
6960          Set fso = New FileSystemObject
6970          Set fsf = fso.GetFile(gstrTrustDataLocation & gstrFile_RestoreDataName)
6980          fsf.Copy strDataPathFile_Current
6990          Set fsf = Nothing

              ' ** Unset read only attribute, if set.
7000          Set fso = New FileSystemObject
7010          Set fsf = fso.GetFile(strDataPathFile_Current)
7020          If fsf.Attributes And fsAttrReadOnly Then
7030            fsf.Attributes = fsf.Attributes - fsAttrReadOnly
7040          End If

              ' ** Before we close the whole procedure, check if there is a LedgerArchive_Backup table and restore that to its own mdb.
7050          If RestoreLedgerArchive(strDataPathFile_Current) = "fail" Then  ' ** Function: Below.
                ' ** If Err.Number = 3010 Then  ' ** Table '|' already exists.
7060            blnRetVal = False
7070            DoCmd.Hourglass False
7080            MsgBox "Trust Accountant Archive Data Restore has failed." & vbCrLf & vbCrLf & _
                  "Consult Delta Data, Inc., for further information.", vbCritical + vbOKOnly, "Restore Failed"
7090          Else

                ' ** Delete intermediate files.
7100            If FileExists(gstrTrustDataLocation & "TrustDta.BAK") = True Then  ' ** Module Function: modFileUtilities.
7110              Kill gstrTrustDataLocation & "TrustDta.BAK"
7120            End If
7130            If FileExists(gstrTrustDataLocation & "TrstArch.BAK") = True Then  ' ** Module Function: modFileUtilities.
7140              Kill gstrTrustDataLocation & "TrstArch.BAK"
7150            End If
7160            If FileExists(gstrTrustDataLocation & "TrstRest.mdb") = True Then  ' ** Module Function: modFileUtilities.
7170              Kill gstrTrustDataLocation & "TrstRest.mdb"
7180            End If

7190            Tbl_AddAllLinks strDataPathFile_Current  ' ** Function: Below.

7200            DoCmd.Hourglass False
7210            MsgBox "Trust Accountant Data Restore Successfully Completed!", vbInformation + vbOKOnly, "Restore Successful"
7220            MsgBox "Trust Accountant will now exit." & vbCrLf & vbCrLf & _
                  "In order to provide a clean link to the data," & vbCrLf & _
                  "you will need to restart Trust Accountant.", vbInformation + vbOKOnly, "Trust Accountant Must Exit"
7230            blnQuit = True

7240          End If

7250        End If  ' ** blnRetVal.

7260      Else
            ' ** Backed-up file is different version from current frontend.

            ' ** Now, from here, can it be used by both Other and Floppy?
7270        blnRetVal = RestoreEarlierVersion(strDataPathFile_Current, strArchPathFile_Current, strVer_Current, strVer_Backup, blnQuit)  ' ** Function: Below.

7280      End If  ' ** Current vs. earlier version.
7290    End If  ' ** blnRetVal.

7300    If blnRetVal = False And blnQuit = False Then
7310      MsgBox "Trust Accountant data not restored.", vbInformation + vbOKOnly, "Restore Canceled"
7320      If IsLoaded("frmBackupRestore", acForm) = True Then  ' ** Module Functions: modFileUtilities.
7330        gblnSetFocus = True
7340        Forms("frmBackupRestore").TimerInterval = 100&
7350      End If
7360    End If

EXITP:
7370    If blnQuit = True Then QuitNow  ' ** Module Procedure: modStartupFuncs.
7380    Set rst = Nothing
7390    Set qdf = Nothing
7400    Set tdf = Nothing
7410    Set dbs = Nothing
7420    Set fsf = Nothing
7430    Set fso = Nothing
7440    RestoreFromDrive = blnRetVal
7450    Exit Function

ERRH:
7460    blnRetVal = False
7470    DoCmd.Hourglass False
7480    Select Case ERR.Number
        Case Else
7490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7500    End Select
7510    Resume EXITP

End Function

Private Function RestoreEarlierVersion(strDataPathFile_Current As String, strArchPathFile_Current As String, strVer_Current As String, strVer_Backup As String, blnQuit As Boolean) As Boolean

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "RestoreEarlierVersion"

        Dim dbs1 As DAO.Database, dbs2 As DAO.Database, tdf As DAO.TableDef, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim fso As Scripting.FileSystemObject, fsfd As Scripting.Folder, fsfls As Scripting.FILES, fsfl As Scripting.File
        Dim lngResult As Long
        Dim lngDrives As Long, arr_varDrive As Variant
        Dim strDrivePath As String, strDriveFile As String, strConvertPath As String
        Dim lngOldFiles As Long, arr_varOldFile() As Variant
        Dim strMsg As String, strSQL As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnFound As Boolean, blnCDROM As Boolean
        Dim strTmp01 As String, strTmp02 As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varOldFile().
        Const F_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const F_NAM    As Integer = 0
        Const F_PTHFIL As Integer = 1

7610    blnRetVal = True
7620    blnCDROM = False

7630    If gstrTrustDataLocation = vbNullString Then
7640      IniFile_GetDataLoc  ' ** Module Function: modStartupFuncs.
7650    End If

7660    If CurrentAppPath = gstrDir_Dev Then  ' ** Module Function: modFileUtilities.
7670      strConvertPath = gstrDir_Dev & LNK_SEP & gstrDir_Convert
7680    Else
7690      strConvertPath = gstrTrustDataLocation & gstrDir_Convert
7700    End If

        ' ** Check for \Convert_New, \Convert_Empty, TrustDta.md_, and TrstArch.md_.
7710    blnFound = False
7720    If DirExists(strConvertPath) = True Then  ' ** Module Function: modFileUtilities.
7730      If DirExists(strConvertPath & LNK_SEP & gstrDir_ConvertEmpty) = True Then  ' ** Module Function: modFileUtilities.
7740        strDrivePath = strConvertPath & LNK_SEP & gstrDir_ConvertEmpty
7750        strDriveFile = Left(gstrFile_DataName, (Len(gstrFile_DataName) - 1)) & "_"
7760        strTmp01 = (strConvertPath & LNK_SEP & gstrDir_ConvertEmpty & LNK_SEP & strDriveFile)
7770        If FileExists(strTmp01) = True Then  ' ** Module Function: modFileUtilities.
7780          strDriveFile = Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 1)) & "_"
7790          strTmp01 = (strConvertPath & LNK_SEP & gstrDir_ConvertEmpty & LNK_SEP & strDriveFile)
7800          If FileExists(strTmp01) = True Then  ' ** Module Function: modFileUtilities.
7810            blnFound = True
7820          End If
7830        End If
7840      End If
7850    End If

7860    If blnFound = False Then

7870      If DirExists(strConvertPath) = False Then
7880        MkDir (strConvertPath)
7890      End If
7900      If DirExists(strConvertPath & LNK_SEP & gstrDir_ConvertEmpty) = False Then
7910        MkDir (strConvertPath & LNK_SEP & gstrDir_ConvertEmpty)
7920      End If

7930      DoCmd.Hourglass False

7940      Do While blnRetVal = True

            ' ** Get empty copies of current databases.
7950        strMsg = "Your backup is from an earlier version of Trust Accountant." & vbCrLf & _
              "It will need to be converted before the restore process can be completed." & vbCrLf & vbCrLf & _
              "To do this, Trust Accountant needs additional files found on the installation disk." & vbCrLf & _
              "If you wish to continue, leave this popup open while you complete the Next Steps, below." & vbCrLf & _
              "If you wish to return to the Restore Options screen, click Cancel." & vbCrLf & vbCrLf & _
              "Next Steps:" & vbCrLf & _
              "1. Insert the Trust Accountant installation disk into the drive." & vbCrLf & _
              "2. If the installation program begins automatically, Cancel it at the first opportunity." & vbCrLf & _
              "3. Once the installation program has been canceled, return here and click OK to continue."

7960        msgResponse = MsgBox(strMsg, vbQuestion + vbOKCancel, "Restore From Earlier Version")
7970        If msgResponse = vbOK Then

7980          DoCmd.Hourglass True

7990          arr_varDrive = GetDriveStrings  ' ** Module function: modFileUtilities.
8000          lngDrives = UBound(arr_varDrive) + 1&

              ' ** Find the CD drive with our Setup.exe and Files directory.
8010          blnFound = False: strTmp02 = vbNullString
8020          For lngX = 0& To (lngDrives - 1&)
8030            strDrivePath = arr_varDrive(lngX)
8040            If Right(strDrivePath, 1) = LNK_SEP Then strDrivePath = Left(strDrivePath, (Len(strDrivePath) - 1))
8050            If Right(strDrivePath, 1) = ":" Then strDrivePath = Left(strDrivePath, (Len(strDrivePath) - 1))
8060            If GetTypeDrive(strDrivePath) = vbxDriveCDROM Then  ' ** Module function: modFileUtilities.
8070              gblnBeenToBackup = True  ' ** Used to signal back here if a drive is empty.
8080              If DirExists(strDrivePath & ":" & LNK_SEP & "Files") = True Then  ' ** Module function: modFileUtilities.
8090                strDrivePath = strDrivePath & ":" & LNK_SEP & "Files"
8100                strDriveFile = Left(gstrFile_DataName, (Len(gstrFile_DataName) - 1)) & "_"
8110                If FileExists(strDrivePath & LNK_SEP & strDriveFile) = True Then
8120                  strDriveFile = Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 1)) & "_"
8130                  If FileExists(strDrivePath & LNK_SEP & strDriveFile) = True Then
                        ' ** We'll check the version below.
8140                    blnFound = True
8150                    blnCDROM = True
8160                    Exit For
8170                  End If  ' ** It's got TrstArch.md_.
8180                End If  ' ** It's got TrustDta.md_.
8190              Else
8200                If gblnBeenToBackup = False Then
                      ' ** The drive is empty.
                      ' ** Continue through the loop, and if it's not found
                      ' ** by the end, let them know this drive was empty.
8210                  strTmp02 = strDrivePath & ":"
8220                Else
                      ' ** Not found, keep looking.
8230                  gblnBeenToBackup = False
8240                End If
8250              End If  ' ** It's got a Files directory.
8260            End If  ' ** It's a CDROM drive.
8270          Next

8280        Else
8290          blnRetVal = False
8300          DoCmd.Hourglass False
8310          MsgBox "Trust Accountant data not restored.", vbInformation + vbOKOnly, "Restore Canceled"
8320        End If  ' ** msgResponse.

8330      Loop  ' ** While blnRetVal = True.
8340    Else
8350      strMsg = "Your backup is from an earlier version of Trust Accountant, and" & vbCrLf & _
            "must be converted before the restore process can be completed." & vbCrLf & vbCrLf & _
            "Click OK to prepare your backup file for conversion." & vbCrLf & _
            "Click Cancel to return to the Restore Options screen."
8360      msgResponse = MsgBox(strMsg, vbExclamation + vbOKCancel, "Restore From Earlier Version")
8370      If msgResponse <> vbOK Then
8380        blnRetVal = False
8390        MsgBox "Trust Accountant data not restored.", vbInformation + vbOKOnly, "Restore Canceled"
8400      End If
8410    End If  ' ** blnFound.

8420    If blnRetVal = True Then

          ' ** Copy TrustDta.md_ and TrstArch.md_ from the Files directory.
8430      If blnFound = True Then

            ' ** This changes the extension of the current TrustDta.mdb to .BAK.
8440        strTmp01 = (Left(strDataPathFile_Current, Len(strDataPathFile_Current) - 3) & "BAK")
8450        If FileExists(strTmp01) = True Then  ' ** Module Function: modFileUtilities.
              ' ** If there's another .BAK there, delete it!
8460          Kill strTmp01
8470        End If
8480        lngResult = CopyFile(strDataPathFile_Current, strTmp01, False)  ' ** API Function: Above.

            ' ** Also delete any TrstArch.BAK found.
8490        strTmp01 = Parse_Path(strTmp01) & LNK_SEP & gstrFile_ArchDataName  ' ** Module Function: modFileUtilities.
8500        strTmp01 = (Left(strTmp01, (Len(strTmp01) - 3)) & "BAK")
8510        If FileExists(strTmp01) = True Then  ' ** Module Function: modFileUtilities.
8520          Kill strTmp01
8530        End If
8540        strTmp01 = vbNullString

            ' ** Delete current data file.
8550        If FileExists(strDataPathFile_Current) = True Then  ' ** Module Function: modFileUtilities.
8560  On Error Resume Next
8570          Kill strDataPathFile_Current
8580  On Error GoTo ERRH
8590        End If

8600        If blnCDROM = True Then
              ' ** Copy TrustDta.md_ from the CDROM to the Convert_Empty directory.
8610          Set fso = New FileSystemObject
8620          strDriveFile = Left(gstrFile_DataName, (Len(gstrFile_DataName) - 1)) & "_"
8630          Set fsfl = fso.GetFile(strDrivePath & LNK_SEP & strDriveFile)
8640          strTmp01 = (strConvertPath & LNK_SEP & gstrDir_ConvertEmpty & LNK_SEP & strDriveFile)
8650          fsfl.Copy strTmp01
8660          Set fsfl = Nothing
8670        Else
              ' ** Use existing Convert_Empty files.
8680          strDrivePath = strConvertPath & LNK_SEP & gstrDir_ConvertEmpty
8690          strDriveFile = Left(gstrFile_DataName, (Len(gstrFile_DataName) - 1)) & "_"
8700          strTmp01 = (strConvertPath & LNK_SEP & gstrDir_ConvertEmpty & LNK_SEP & strDriveFile)
8710        End If
8720        Set fso = Nothing
8730        Set fso = New FileSystemObject

            ' ** Unset read only attribute, if set.
8740        Set fso = New FileSystemObject
8750        Set fsfl = fso.GetFile(strTmp01)
8760        If fsfl.Attributes And fsAttrReadOnly Then
8770          fsfl.Attributes = fsfl.Attributes - fsAttrReadOnly
8780        End If
8790        Set fsfl = Nothing

8800        If blnCDROM = True Then
              ' ** Copy TrustDta.md_ from the CDROM to the data directory.
8810          Set fso = New FileSystemObject
8820          strDriveFile = Left(gstrFile_DataName, (Len(gstrFile_DataName) - 1)) & "_"
8830          Set fsfl = fso.GetFile(strDrivePath & LNK_SEP & strDriveFile)
8840          fsfl.Copy strDataPathFile_Current
8850          Set fsfl = Nothing
8860        Else
              ' ** Copy TrustDta.md_ from the Convert_Empty directory to the data directory.
8870          Set fso = New FileSystemObject
8880          Set fsfl = fso.GetFile(strDrivePath & LNK_SEP & strDriveFile)
8890          fsfl.Copy strDataPathFile_Current
8900          Set fsfl = Nothing
8910        End If

            ' ** Unset read only attribute, if set.
8920        Set fso = New FileSystemObject
8930        Set fsfl = fso.GetFile(strDataPathFile_Current)
8940        If fsfl.Attributes And fsAttrReadOnly Then
8950          fsfl.Attributes = fsfl.Attributes - fsAttrReadOnly
8960        End If
8970        Set fsfl = Nothing

            ' ** Check the version of the installation disk files.
8980        strTmp01 = vbNullString
8990        Set dbs1 = CurrentDb
9000        With dbs1
9010          Set rst = .OpenRecordset("m_VD", dbOpenDynaset, dbReadOnly)
9020          With rst
9030            .MoveFirst
9040            strTmp01 = ![vd_MAIN] & "." & ![vd_MINOR] & "." & ![vd_REVISION]
9050            If Right(strTmp01, 2) = ".0" Then strTmp01 = strTmp01 & "0"  ' ** Make v2.2.0 = v2.2.00.
9060            .Close
9070          End With
9080          .Close
9090        End With
9100        Set rst = Nothing
9110        Set dbs1 = Nothing
9120        If strTmp01 <> strVer_Current Then
              ' ** Delete the wrong version file.
9130          Kill strDataPathFile_Current
              ' ** Delete its copy in Convert_Empty.
9140          strTmp01 = (strConvertPath & LNK_SEP & gstrDir_ConvertEmpty & LNK_SEP & strDriveFile)
9150          Kill strTmp01
              ' ** Copy the BAK back to MDB.
9160          strTmp01 = (Left(strDataPathFile_Current, Len(strDataPathFile_Current) - 3) & "BAK")
9170          lngResult = CopyFile(strTmp01, strDataPathFile_Current, False)  ' ** API Function: Above.
              ' ** Delete the BAK.
9180          Kill strTmp01
9190          DoCmd.Hourglass False
9200          If blnCDROM = True Then
9210            blnRetVal = False
9220            MsgBox "The CD is not the current version.", vbInformation + vbOKOnly, ("File Not Found" & Space(40))
9230          Else
9240            blnRetVal = False
9250            MsgBox "Files necessary for proper restoration were not found.", vbInformation + vbOKOnly, ("File Not Found" & Space(40))
9260          End If
9270          MsgBox "Trust Accountant data not restored.", vbInformation + vbOKOnly, "Restore Canceled"
9280        Else

              ' ** This changes the extension of the current TrstArch.mdb to .BAK.
9290          strTmp01 = (Left(strArchPathFile_Current, (Len(strArchPathFile_Current) - 3)) & "BAK")
9300          lngResult = CopyFile(strArchPathFile_Current, strTmp01, False)  ' ** API Function: Above.

              ' ** Delete current archive file.
9310          If FileExists(strArchPathFile_Current) = True Then  ' ** Module Function: modFileUtilities.
9320            Kill strArchPathFile_Current
9330          End If

9340          If blnCDROM = True Then
                ' ** Copy TrstArch.md_ from the CDROM to the Convert_Empty directory.
9350            Set fso = New FileSystemObject
9360            strDriveFile = Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 1)) & "_"
9370            Set fsfl = fso.GetFile(strDrivePath & LNK_SEP & strDriveFile)
9380            strTmp01 = (strConvertPath & LNK_SEP & gstrDir_ConvertEmpty & LNK_SEP & strDriveFile)
9390            fsfl.Copy strTmp01
9400            Set fsfl = Nothing
9410          Else
                ' ** Use existing Convert_Empty files.
9420            strDrivePath = strConvertPath & LNK_SEP & gstrDir_ConvertEmpty
9430            strDriveFile = Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 1)) & "_"
9440            strTmp01 = (strConvertPath & LNK_SEP & gstrDir_ConvertEmpty & LNK_SEP & strDriveFile)
9450          End If

              ' ** Unset read only attribute, if set.
9460          Set fso = New FileSystemObject
9470          Set fsfl = fso.GetFile(strTmp01)
9480          If fsfl.Attributes And fsAttrReadOnly Then
9490            fsfl.Attributes = fsfl.Attributes - fsAttrReadOnly
9500          End If
9510          Set fsfl = Nothing

9520          If blnCDROM = True Then
                ' ** Copy TrstArch.md_ from the CDROM to the data directory.
9530            Set fso = New FileSystemObject
9540            strDriveFile = Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 1)) & "_"
9550            Set fsfl = fso.GetFile(strDrivePath & LNK_SEP & strDriveFile)
9560            fsfl.Copy strArchPathFile_Current
9570            Set fsfl = Nothing
9580          Else
                ' ** Copy TrstArch.md_ from the Convert_Empty directory to the data directory.
9590            Set fso = New FileSystemObject
9600            Set fsfl = fso.GetFile(strDrivePath & LNK_SEP & strDriveFile)
9610            fsfl.Copy strArchPathFile_Current
9620            Set fsfl = Nothing
9630          End If

              ' ** Unset read only attribute, if set.
9640          Set fso = New FileSystemObject
9650          Set fsfl = fso.GetFile(strArchPathFile_Current)
9660          If fsfl.Attributes And fsAttrReadOnly Then
9670            fsfl.Attributes = fsfl.Attributes - fsAttrReadOnly
9680          End If
9690          Set fsfl = Nothing

9700        End If  ' ** CD files are current version.

9710      Else
            ' ** Empty data files not found on a CD.
9720        If strTmp02 <> vbNullString Then
9730          DoCmd.Hourglass False
9740          MsgBox "The " & strTmp02 & " drive is empty.", vbInformation + vbOKCancel, ("File Not Found" & Space(40))
9750          blnRetVal = False
9760          MsgBox "Trust Accountant data not restored.", vbInformation + vbOKOnly, "Restore Canceled"
9770        Else
9780          blnRetVal = False
9790          DoCmd.Hourglass False
9800          MsgBox "Needed files not found on CD." & vbCrLf & _
                "Contact Delta Data, Inc., for assistance.", vbInformation + vbOKOnly, "File Not Found"
9810        End If
9820      End If

9830    End If  ' ** blnRetVal.

9840    If gstrTrustDataLocation = vbNullString Then
9850      IniFile_GetDataLoc  ' ** Module Function: modStartupFuncs.
9860    End If

        ' ** Delete anything we've already done if the restore was canceled above.
9870    If blnRetVal = False Then
          ' ** Delete the local copy of the backed-up file.
9880      If FileExists(gstrTrustDataLocation & gstrFile_RestoreDataName) = True Then  ' ** Module Function: modFileUtilities.
9890        Kill (gstrTrustDataLocation & gstrFile_RestoreDataName)
9900      End If
          ' ** Leave the Convert_Empty and Convert_New directories alone.
9910    End If

        ' ** Now move on to conversion setup.
9920    If blnRetVal = True Then

9930      strTmp01 = strConvertPath

9940      lngOldFiles = 0&
9950      ReDim arr_varOldFile(F_ELEMS, 0)

9960      Set fso = CreateObject("Scripting.FileSystemObject")
9970      With fso
9980        Set fsfd = .GetFolder(strTmp01)
9990        Set fsfls = fsfd.FILES
10000       lngOldFiles = fsfls.Count
10010       If lngOldFiles > 0& Then
10020         ReDim arr_varOldFile(F_ELEMS, (lngOldFiles - 1&))
              ' ** Get all the files in the Convert_New folder.
10030         lngX = -1&
10040         For Each fsfl In fsfls
10050           With fsfl
10060             lngX = lngX + 1&
10070             ReDim Preserve arr_varOldFile(F_ELEMS, lngX)
                  ' *******************************************************
                  ' ** Array: arr_varOldFile()
                  ' **
                  ' **   Element  Name                        Constant
                  ' **   =======  ==========================  ===========
                  ' **      0     Name                        F_NAM
                  ' **      1     Path and File               F_PTHFIL
                  ' **
                  ' *******************************************************
10080             arr_varOldFile(F_NAM, lngX) = .Name
10090             arr_varOldFile(F_PTHFIL, lngX) = .Path
10100           End With  ' ** This file: fsfl.
10110         Next  ' ** For each file: fsfl.
10120       End If
10130     End With  ' ** fso.

10140     If lngOldFiles > 0& Then
10150       If lngOldFiles = 1& And arr_varOldFile(F_NAM, 0) = "ReadMeCN.txt" Then
              ' ** Leave it be.
10160       Else
              ' ** Create a backup folder for the files found, and move them into it.
10170         strTmp02 = DirExists2(strTmp01, "Backup")  ' ** Module Function: modFileUtilities.
10180         For lngX = 0& To (lngOldFiles - 1&)
10190           If arr_varOldFile(F_NAM, lngX) <> "ReadMeCN.txt" Then
10200             Name arr_varOldFile(F_PTHFIL, lngX) As strTmp02 & LNK_SEP & arr_varOldFile(F_NAM, lngX)
10210           End If
10220         Next
10230       End If
10240     End If

          ' ** Copy backed-up file to Convert_new directory.
10250     Set fso = New FileSystemObject
10260     Set fsfl = fso.GetFile(gstrTrustDataLocation & gstrFile_RestoreDataName)
10270     strTmp01 = strConvertPath & LNK_SEP & gstrFile_DataName
10280     If FileExists(strTmp01) = True Then  ' ** Module Function: modFileUtilities.
10290       Kill strTmp01
10300     End If
10310     fsfl.Copy strTmp01
10320     Set fsfl = Nothing

          ' ** Unset read only attribute, if set.
10330     Set fso = New FileSystemObject
10340     Set fsfl = fso.GetFile(strTmp01)
10350     If fsfl.Attributes And fsAttrReadOnly Then
10360       fsfl.Attributes = fsfl.Attributes - fsAttrReadOnly
10370     End If
10380     Set fsfl = Nothing

          ' ** Delete first copy of backed-up file.
10390     If FileExists(gstrTrustDataLocation & gstrFile_RestoreDataName) = True Then  ' ** Module Function: modFileUtilities.
10400       Kill gstrTrustDataLocation & gstrFile_RestoreDataName
10410     End If

          ' ** Create an empty MDB for the 'old' archive.
10420     strTmp01 = (strConvertPath & LNK_SEP & gstrFile_ArchDataName)
10430     If FileExists(strTmp01) = True Then
10440       Kill strTmp01
10450     End If

10460     Set dbs1 = DBEngine.Workspaces(0).CreateDatabase(strTmp01, dbLangGeneral)  ' ** Use default Jet Engine.

          ' ** Create an m_VA table in the 'old' archive.
10470     strSQL = "SELECT m_VA.* INTO m_VA IN '" & strTmp01 & "' FROM m_VA;"
10480     CurrentDb.Execute strSQL, dbFailOnError

          ' ** Update m_VA with the backup version number.
10490     With dbs1
10500       .TableDefs.Refresh
10510       Set rst = .OpenRecordset("m_VA", dbOpenDynaset, dbConsistent)
10520       With rst
10530         .MoveFirst
10540         .Edit
10550         strTmp01 = strVer_Backup
10560         ![va_MAIN] = CInt(Left(strTmp01, (InStr(strTmp01, ".") - 1)))
10570         strTmp01 = Mid(strTmp01, (InStr(strTmp01, ".") + 1))
10580         If InStr(strTmp01, ".") > 0 Then
10590           ![va_MINOR] = CInt(Left(strTmp01, (InStr(strTmp01, ".") - 1)))
10600           strTmp01 = Mid(strTmp01, (InStr(strTmp01, ".") + 1))
10610           ![va_REVISION] = CInt(strTmp01)
10620         Else
10630           ![va_MINOR] = CInt(strTmp01)
10640           ![va_REVISION] = CInt(0)
10650         End If
10660         ![va_DE1] = Null
10670         ![va_DE2] = Null
10680         .Update
10690         .Close
10700       End With
10710     End With

          ' ** Copy in the LedgerArchive from the backup.
10720     strTmp01 = (strConvertPath & LNK_SEP & gstrFile_DataName)
10730     Set dbs2 = DBEngine.Workspaces(0).OpenDatabase(strTmp01)
10740     strSQL = "SELECT LedgerArchive_backup.* INTO Ledger IN '" & dbs1.Name & "' FROM LedgerArchive_backup;"
10750     dbs2.Execute strSQL

          ' ** Delete the original from the backup.
10760     dbs2.TableDefs.Delete "LedgerArchive_Backup"
10770     dbs1.Close
10780     dbs2.Close
10790     Set rst = Nothing
10800     Set dbs1 = Nothing
10810     Set dbs2 = Nothing

          ' ** Now we're ready for the conversion!
10820     Forms("frmBackupRestore").cmdClose.SetFocus
10830     Forms("frmBackupRestore").cmdRestore.Enabled = False
10840     Forms("frmBackupRestore").opgLocRestore.Enabled = False

10850     Set dbs1 = CurrentDb
10860     With dbs1
            ' ** tblPreference_User, for 'chkConversionCheck', by specified [usr].
10870       Set qdf = .QueryDefs("qryPreferences_06_01")  '##dbs_id
10880       With qdf.Parameters
10890         ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10900       End With
10910       Set rst = qdf.OpenRecordset
10920       With rst
10930         If .BOF = True And .EOF = True Then
                ' ** No preference for ConversionCheck().
10940         Else
10950           .MoveFirst
10960           If ![prefuser_boolean] = True Then
10970             .Edit
10980             ![prefuser_boolean] = False
10990             ![DateModified] = Now()
11000             .Update
11010           End If
11020         End If
11030         .Close
11040       End With
11050       .Close
11060     End With

11070     DoCmd.Hourglass False
11080     MsgBox "Your backup data is now ready for conversion." & vbCrLf & vbCrLf & _
            "When you reopen Trust Accountant, your backup " & vbCrLf & _
            "data will be converted to the current version.", vbExclamation + vbOKOnly, "Files Copied Successfully"
11090     MsgBox "Trust Accountant will now exit." & vbCrLf & vbCrLf & _
            "In order to provide a clean link to the data," & vbCrLf & _
            "you will need to restart Trust Accountant.", vbInformation + vbOKOnly, "Trust Accountant Must Exit"
11100     blnQuit = True

11110   End If  ' ** blnRetVal.

11120   If blnRetVal = False Then
11130     If IsLoaded("frmBackupRestore", acForm) = True Then  ' ** Module Functions: modFileUtilities.
11140       gblnSetFocus = True
11150       Forms("frmBackupRestore").TimerInterval = 100&
11160     End If
11170   End If

EXITP:
11180   Set rst = Nothing
11190   Set qdf = Nothing
11200   Set tdf = Nothing
11210   Set dbs1 = Nothing
11220   Set dbs2 = Nothing
11230   Set fsfl = Nothing
11240   Set fsfls = Nothing
11250   Set fsfd = Nothing
11260   Set fso = Nothing
11270   RestoreEarlierVersion = blnRetVal
11280   Exit Function

ERRH:
11290   blnRetVal = False
11300   DoCmd.Hourglass False
11310   Select Case ERR.Number
        Case Else
11320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11330   End Select
11340   Resume EXITP

End Function

Private Function RestoreLedgerArchive(strDataPathFile_Current As String) As String
' ** If the back up we just restored from has a LedgerArchive_Backup included,
' ** then we need to make TrstArch.mdb out of it

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "RestoreLedgerArchive"

        Dim fso As Scripting.FileSystemObject, fsfArchiveBackup As Scripting.File, fsfArchive As Scripting.File
        Dim dbsTarget As DAO.Database, dbsArch As DAO.Database, qdf As DAO.QueryDef, tdf As DAO.TableDef
        Dim strArchive As String
        Dim strArchiveBackup As String
        Dim lngRecsArchive As Long
        Dim blnDelete As Boolean
        Dim lngX As Long
        Dim strRetVal As String

        ' ** Preset to 'Fail'.
11410   strRetVal = "Fail"

11420   strArchive = gstrTrustDataLocation & gstrFile_ArchDataName

        ' ** If the file exists, copy it.
11430   If FileExists_Local(strArchive) = True Then
          ' ** This copies TrstArch.mdb to TrstArch.BAK, leaving the original intact.
11440     Set fso = New FileSystemObject
11450     strArchiveBackup = Left(strArchive, Len(strArchive) - 3) & "BAK"
11460     Set fsfArchive = fso.GetFile(strArchive)
11470     fsfArchive.Copy strArchiveBackup
11480   End If

11490   Set dbsArch = DBEngine.Workspaces(0).OpenDatabase(strArchive)
11500   With dbsArch
11510     blnDelete = False
11520     For Each tdf In .TableDefs
11530       If tdf.Name = "ledger" Then  ' ** Of course it'll be there!
11540         blnDelete = True
11550         Exit For
11560       End If
11570     Next
11580     Set tdf = Nothing
11590     If blnDelete = True Then
11600       .TableDefs.Delete "ledger"
11610     End If
11620     .TableDefs.Refresh
11630     .TableDefs.Refresh
          ' ** Data-Definition: Create table ledger.
          ' ** qryBackupRestore_04_01 - qryBackupRestore_04_13.
11640     For lngX = 1& To 13&
11650       .Execute CurrentDb.QueryDefs("qryBackupRestore_04_" & Right("00" & CStr(lngX), 2)).SQL
11660     Next
11670     .TableDefs.Refresh
11680     .TableDefs.Refresh
11690     .Close
11700   End With
11710   Set dbsArch = Nothing

11720   If TableExists(gstrTable_LedgerArchive, True, gstrFile_DataName) = True Then  ' ** Module Function: modFileUtilities.

11730     DoCmd.TransferDatabase acLink, "Microsoft Access", strDataPathFile_Current, acTable, _
            gstrTable_LedgerArchive, gstrTable_LedgerArchive
11740     CurrentDb.TableDefs.Refresh
11750     CurrentDb.TableDefs.Refresh

11760     lngRecsArchive = DCount("*", gstrTable_LedgerArchive)
11770     If lngRecsArchive > 0& Then

11780       blnDelete = False
11790       Select Case TableExists("LedgerArchive")  ' ** Module Function: modFileUtilities.
            Case True
11800         DoCmd.DeleteObject acTable, "LedgerArchive"
11810         CurrentDb.TableDefs.Refresh
11820         CurrentDb.TableDefs.Refresh
11830       Case False
11840         blnDelete = True
11850       End Select

11860       DoCmd.TransferDatabase acLink, "Microsoft Access", strArchive, acTable, "ledger", "LedgerArchive"
11870       CurrentDb.TableDefs.Refresh
11880       CurrentDb.TableDefs.Refresh

            ' ** Append LedgerArchive_Backup to LedgerArchive.
11890       Set qdf = CurrentDb.QueryDefs("qryBackupRestore_05")
11900       qdf.Execute
11910       Set qdf = Nothing

11920       DoCmd.DeleteObject acTable, gstrTable_LedgerArchive
11930       CurrentDb.TableDefs.Refresh
11940       CurrentDb.TableDefs.Refresh

11950       If blnDelete = True Then
11960         DoCmd.DeleteObject acTable, "LedgerArchive"
11970         CurrentDb.TableDefs.Refresh
11980         CurrentDb.TableDefs.Refresh
11990       End If

12000     End If

          ' ** We already know LedgerArchive_Backup exists in TrustDta.mdb.
12010     Set dbsTarget = DBEngine.Workspaces(0).OpenDatabase(strDataPathFile_Current)
12020     With dbsTarget
            ' ** Delete the original from the backup.
12030       .TableDefs.Delete gstrTable_LedgerArchive
12040       .TableDefs.Refresh
12050       .TableDefs.Refresh
12060       .Close
12070     End With
12080     Set dbsTarget = Nothing

12090   End If

12100   strRetVal = "Pass"

EXITP:
12110   Set tdf = Nothing
12120   Set qdf = Nothing
12130   Set fsfArchiveBackup = Nothing
12140   Set fsfArchive = Nothing
12150   Set fso = Nothing
12160   Set dbsTarget = Nothing
12170   Set dbsArch = Nothing
12180   RestoreLedgerArchive = strRetVal
12190   Exit Function

ERRH:
12200   strRetVal = "Fail"
12210   Select Case ERR.Number
        Case Else
12220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12230   End Select
12240   Resume EXITP

End Function

Private Function FileExists_Local(strFileName As String) As Boolean

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "FileExists_Local"

        Dim strMsg As String
        Dim blnRetVal As Boolean

        ' ** Define constants to represent intrinsic Visual Basic error codes.
        Const mnErrDiskNotReady      As Long = 71&
        Const mnErrDeviceUnavailable As Long = 68&

12310   blnRetVal = FileExists(strFileName)  ' ** Module Function: modFileUtilities.

EXITP:
12320   FileExists_Local = blnRetVal
12330   Exit Function

ERRH:
12340   Select Case ERR.Number
        Case mnErrDiskNotReady
12350     strMsg = "Put a CD/DVD disk in the drive, close the drive, and click OK to continue," & vbCrLf & _
            "or click Cancel to exit."
12360     If MsgBox(strMsg, vbInformation + vbOKCancel, "Drive Not Ready") = vbOK Then
12370       Resume       ' ** Resumes with statement causing the error.
12380     Else
12390       Resume Next  ' ** Resumes with statement after statement causing the error.
12400     End If
12410   Case mnErrDeviceUnavailable
12420     strMsg = "This drive or path does not exist:" & vbCrLf & vbCrLf & strFileName
12430     MsgBox strMsg, vbExclamation + vbOKOnly, "Path Not Found"
12440     Resume Next
12450   Case Else
12460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12470   End Select
12480   Resume EXITP

End Function

Public Function Tbl_DelAllLinks(Optional varSpecificFile As Variant) As Boolean

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "Tbl_DelAllLinks"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, Rel As DAO.Relation
        Dim lngLinks As Long, arr_varLink() As Variant
        Dim lngRels As Long, arr_varRel() As Variant
        Dim blnAll As Boolean, strSpecificFile As String
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

12510   blnRetVal = True

12520   Select Case IsMissing(varSpecificFile)
        Case True
12530     blnAll = True
12540     strSpecificFile = vbNullString
12550   Case False
12560     blnAll = False
12570     strSpecificFile = varSpecificFile
12580   End Select

12590   lngLinks = 0&
12600   ReDim arr_varLink(0)

12610   Set dbs = CurrentDb
12620   With dbs
12630     For Each tdf In .TableDefs
12640       With tdf
12650         If IsNull(.Connect) = False Then
12660           If Trim(.Connect) <> vbNullString Then
12670             Select Case blnAll
                  Case True
12680               lngLinks = lngLinks + 1&
12690               ReDim Preserve arr_varLink(lngLinks - 1&)
12700               arr_varLink(lngLinks - 1&) = .Name
12710             Case False
12720               If InStr(.Connect, strSpecificFile) > 0 Then
12730                 lngLinks = lngLinks + 1&
12740                 ReDim Preserve arr_varLink(lngLinks - 1&)
12750                 arr_varLink(lngLinks - 1&) = .Name
12760               End If
12770             End Select
12780           End If
12790         End If
12800       End With
12810     Next
12820   End With

12830   If lngLinks > 0& Then
12840     For lngX = (lngLinks - 1&) To 0& Step -1&
12850       lngRels = 0&
12860       ReDim arr_varRel(0)
12870 On Error Resume Next
12880       DoCmd.DeleteObject acTable, arr_varLink(lngX)
12890       If ERR.Number <> 0 Then
12900 On Error GoTo ERRH
12910         For Each Rel In dbs.Relations
12920           With Rel
12930             If .Table = arr_varLink(lngX) Or .ForeignTable = arr_varLink(lngX) Then
12940               strTmp01 = dbs.TableDefs(.Table).Connect
12950               strTmp02 = dbs.TableDefs(.ForeignTable).Connect
12960               strTmp01 = Parse_File(strTmp01)  ' ** Module Function: modFileUtilities.
12970               strTmp02 = Parse_File(strTmp02)  ' ** Module Function: modFileUtilities.
12980               If strTmp01 <> strTmp02 Then
                      ' ** Non-Contiguous relationship.
12990                 lngRels = lngRels + 1&
13000                 ReDim Preserve arr_varRel(lngRels - 1&)
13010                 arr_varRel(lngRels - 1&) = .Name
13020               End If
13030             End If
13040           End With
13050         Next
13060         If lngRels > 0& Then
13070           For lngY = 0& To (lngRels - 1&)
13080             dbs.Relations.Delete arr_varRel(lngY)
13090           Next
13100         End If
13110         DoCmd.DeleteObject acTable, arr_varLink(lngX)
13120       Else
13130 On Error GoTo ERRH
13140       End If
13150     Next
13160     If blnAll = True Or (blnAll = False And strSpecificFile = "TrustDta.mdb") Then
13170       strTmp03 = "_~xusr"
13180       If TableExists(strTmp03) = True Then  ' ** Module Function: modFileUtilities.
13190         DoCmd.DeleteObject acTable, strTmp03
13200       End If
13210     End If
13220   End If
13230   dbs.Close

13240   CurrentDb.TableDefs.Refresh
13250   CurrentDb.TableDefs.Refresh

13260   Beep

EXITP:
13270   Set Rel = Nothing
13280   Set tdf = Nothing
13290   Set dbs = Nothing
13300   Tbl_DelAllLinks = blnRetVal
13310   Exit Function

ERRH:
13320   blnRetVal = False
13330   Select Case ERR.Number
        Case Else
13340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13350   End Select
13360   Resume EXITP

End Function

Public Function Tbl_AddAllLinks(strPathFile As String) As Boolean
' ** Add links one mdb at a time.

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "Tbl_AddAllLinks"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngLinks As Long, arr_varLink As Variant
        Dim blnAuxLoc As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varLink().
        'Const L_DID  As Integer = 0
        'Const L_DNAM As Integer = 1
        'Const L_TID  As Integer = 2
        'Const L_TNAM As Integer = 3
        'Const L_ID   As Integer = 4
        Const L_NAM  As Integer = 5
        Const L_SNAM As Integer = 6

13410   blnRetVal = True

13420   varTmp00 = DLookup("[seclic_auxloc]", "tblSecurity_License")
13430   Select Case IsNull(varTmp00)
        Case True
13440     blnAuxLoc = False
13450   Case False
13460     Select Case varTmp00
          Case True
13470       blnAuxLoc = True
13480     Case False
13490       blnAuxLoc = False
13500     End Select
13510   End Select

13520   Set dbs = CurrentDb
13530   With dbs
          ' ** tblTemplate_Database_Table_Link, by specified [dbnam], CurrentAppName().
13540     Set qdf = .QueryDefs("qryBackupRestore_01")
13550     With qdf.Parameters
13560       If InStr(strPathFile, LNK_SEP) > 0 Then
13570         ![dbnam] = Parse_File(strPathFile)  ' ** Module Function: modFileUtilities.
13580       Else
13590         ![dbnam] = strPathFile
13600         If gstrTrustDataLocation = vbNullString Then
13610           blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
13620         End If
13630         If InStr(strPathFile, "TrustAux") > 0 Then
13640           Select Case blnAuxLoc
                Case True
13650             gstrTrustAuxLocation = CurrentAppPath & LNK_SEP  ' ** Module Function: modFileUtilities.
13660           Case False
13670             gstrTrustAuxLocation = gstrTrustDataLocation
13680           End Select  ' ** blnAuxLoc.
13690           strPathFile = gstrTrustAuxLocation & strPathFile
13700         Else
13710           strPathFile = gstrTrustDataLocation & strPathFile
13720         End If
13730       End If
13740     End With
13750     Set rst = qdf.OpenRecordset
13760     With rst
13770       .MoveLast
13780       lngLinks = .RecordCount
13790       .MoveFirst
13800       arr_varLink = .GetRows(lngLinks)
            ' ***********************************************************
            ' ** Array: arr_varLink()
            ' **
            ' **   Field  Element  Name                      Constant
            ' **   =====  =======  ========================  ==========
            ' **     1       0     dbs_id                    L_DID
            ' **     2       1     dbs_name                  L_DNAM
            ' **     3       2     tbl_id                    L_TID
            ' **     4       3     tbl_name                  L_TNAM
            ' **     5       4     tbllnk_id                 L_ID
            ' **     6       5     tbllnk_name               L_NAM
            ' **     7       6     tbllnk_sourcetablename    L_SNAM
            ' **
            ' ***********************************************************
13810     End With
13820     .Close
13830   End With

13840   For lngX = 0& To (lngLinks - 1&)
13850     If TableExists(CStr(arr_varLink(L_NAM, lngX))) = False Then  ' ** Module Function: modFileUtilities.
13860       DoCmd.TransferDatabase acLink, "Microsoft Access", strPathFile, acTable, _
              arr_varLink(L_SNAM, lngX), arr_varLink(L_NAM, lngX)
13870     End If
13880   Next

        'Rel_Missing True  ' ** Module Function: modXAdminFuncs.

EXITP:
13890   Tbl_AddAllLinks = blnRetVal
13900   Exit Function

ERRH:
13910   blnRetVal = False
13920   Select Case ERR.Number
        Case Else
13930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13940   End Select
13950   Resume EXITP

End Function
