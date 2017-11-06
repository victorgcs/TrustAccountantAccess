Attribute VB_Name = "zz_mod_VersionDocFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_VersionDocFuncs"

'VGC 10/27/2017: CHANGES!
'REM'D BACKEND LINK!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const IsDev = 0  ' ** 0 = release, -1 = development.
' ** Also in:
' **   frmXAdmin_Misc
' **   modAutonumberFieldFuncs
' **   modExcelFuncs
' **   zz_mod_MDEPrepFuncs

' *********************************************************************************************************************
' ** Microsoft Access Versions:
' **
' **   Number  Version Name            Date  Office Suite Version
' **   ======  =====================  ======  =======================================================================
' **    1.0    Access 1.1              1992
' **    2.0    Access 2.0              1993   Office 4.3 Pro
' **    7.0    Access for Windows 95   1995   Office 95 Professional
' **    8.0    Access 97               1997   Office 97 Professional and Developer
' **    9.0    Access 2000             1999   Office 2000 Professional, Premium and Developer
' **    10.0   Access 2002             2001   Office XP Professional and Developer
' **    11.0   Access 2003             2003   Office 2003 Professional and Professional Enterprise
' **    12.0   Access 2007             2007   Office 2007 Professional, Professional Plus, Ultimate and Enterprise
' **    14.0   Access 2010             2010   Office 2010 Professional, Professional Academic and Professional Plus
' **    15.0   Access 2013             2013   Office 2013 Professional and Professional Plus
' **    16.0   Access 2016             2016   Office 2016 Professional and Professional Plus
' ** There are no Access versions between 2.0 and 7.0 because the Windows 95 version was launched with Word 7.
' ** All of the Office 95 products have OLE 2 capabilities, and Access 7 shows that it was compatible with Word 7.
' ** Version 13.0 was skipped for marketing considerations.
' **
' *********************************************************************************************************************

' ** AcFileFormatAccess enumeration:
' **    2  acFileFormatAccess2     Microsoft Access 2.0 format
' **    7  acFileFormatAccess95    Microsoft Access 95 format
' **    8  acFileFormatAccess97    Microsoft Access 97 format
' **    9  acFileFormatAccess2000  Microsoft Access 2000 format
' **   10  acFileFormatAccess2002  Microsoft Access 2002 format
' **   11                          Microsoft Access 2003 format
' **   12  acFileFormatAccess12    Microsoft Access 2007 format
' ** Note that this is the file format, not the Access version.

' ** Error: 3112, Record(s) cannot be read; no read permission on 'm_VP'.
' ** Returned on a db1.mdb. These 'Compact & Repair' intermediaries
' ** I've found to be unreadable. In an attemt to open them directly,
' ** the following message is shown:
' **   This database is in an unexpected state; Microsoft Access can't open it.
' **   This database has been converted from a prior version of Microsoft
' **     Access by using the Tools menu DAO CompactDatabase method instead
' **     of the Convert Database command on the (Database Utilities submenu).
' **     This has left the database in a partially converted state.
' **   If you have a copy of the database in its original format, use the
' **     Convert Database command on the Tools menu (Database Utilities submenu)
' **     to convert it. If the original database is no longer available, create a
' **     new database and import your tables and queries to preserve your data.
' **     Your other database objects can't be recovered.
' ** I tried importing from the menu, but that just produced the above message again.
' ** I haven't tried via VBA.

Public glngInstance As Long
' **

Public Function Version_Convert_RunAllStart(Optional varInst As Variant) As Boolean
' ** Remarked out.

      #If IsDev Then

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Convert_RunAllStart"

        Dim blnRetVal As Boolean

110     blnRetVal = True

        ' ** Borrowing this for Directory Version numbers.
120     gstrCrtRpt_Version = vbNullString

130     If IsMissing(varInst) = True Then
140       glngInstance = 0&
150     Else
160       glngInstance = glngInstance + 1&
170     End If
180     If glngInstance >= 0& And glngInstance <= 20& Then
190       If IsLoaded("frmPleaseWait") = True Then  ' ** Module Function: modFileUtilities.
200         DoCmd.Close acForm, "frmPleaseWait"
210       End If
220       DoEvents
230       Version_Convert_RunAll glngInstance  ' ** Function: Below.
240     End If

EXITP:
250     Version_Convert_RunAllStart = blnRetVal
260     Exit Function

ERRH:
270     blnRetVal = False
280     Select Case ERR.Number
        Case Else
290       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
300     End Select
310     Resume EXITP

      #End If

End Function

Public Function Version_Convert_RunAll(lngInstance As Long) As Boolean
' ** Only above.

      #If IsDev Then

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Convert_RunAll"

        Dim fso As Scripting.FileSystemObject
        Dim fsfds1 As Scripting.Folders
        Dim fsfd1 As Scripting.Folder, fsfd2 As Scripting.Folder
        Dim strThisPath As String, strDataPath As String, strVerPath As String, strConvertPath As String
        Dim lngVerDirs As Long, arr_varVerDir() As Variant
        Dim intConversionCheck As Integer
        Dim blnContinue As Boolean
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varVerDir().
        Const D_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const D_DIR      As Integer = 0
        Const D_PATH     As Integer = 1
        Const D_FILES    As Integer = 2
        Const D_FILE_ARR As Integer = 3
        ' ******************************************
        ' ** Array: arr_varVerDir()
        ' **
        ' **   Element  Name          Constant
        ' **   =======  ============  ============
        ' **      0     Name          D_DIR
        ' **      1     Path          D_PATH
        ' **      2     Files         D_FILES
        ' **      3     File Array    D_FILE_ARR
        ' **
        ' ******************************************

410     blnRetVal = True
420     blnContinue = True

430     strThisPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
440     strVerPath = strThisPath & LNK_SEP & gstrDir_DevVer
450     strConvertPath = strThisPath & LNK_SEP & gstrDir_Convert

460     lngVerDirs = 0&
470     ReDim arr_varVerDir(D_ELEMS, 0)

480     Set fso = CreateObject("Scripting.FileSystemObject")
490     With fso

500       Set fsfd1 = .GetFolder(strVerPath)
510       Set fsfds1 = fsfd1.SubFolders

520       lngVerDirs = 0&
530       ReDim arr_varVerDir(D_ELEMS, 0)

540       For Each fsfd2 In fsfds1
550         With fsfd2
560           If Left(.Name, 4) = "Ver_" Then
570             lngVerDirs = lngVerDirs + 1&
580             lngE = lngVerDirs - 1&
590             ReDim Preserve arr_varVerDir(D_ELEMS, lngE)
600             arr_varVerDir(D_DIR, lngE) = .Name
610             arr_varVerDir(D_PATH, lngE) = .Path
620             arr_varVerDir(D_FILES, lngE) = CLng(0)
630             arr_varVerDir(D_FILE_ARR, lngE) = Empty
640           End If
650         End With  ' ** fsfd2.
660       Next

670     End With  ' ** fso.
680     Set fsfd2 = Nothing
690     Set fsfd1 = Nothing
700     Set fsfds1 = Nothing
710     Set fso = Nothing

        'For lngX = 0& To (lngVerDirs - 1&)
        '  Debug.Print "'" & arr_varVerDir(D_DIR, lngX)
        'Next
        ' ** Ver_1_7_00    Ver_2_1_00    Ver_2_1_41    Ver_2_1_47
        ' ** Ver_1_7_10    Ver_2_1_10    Ver_2_1_42    Ver_2_1_50
        ' ** Ver_1_7_20    Ver_2_1_20    Ver_2_1_44    Ver_2_1_55
        ' ** Ver_1_7_40    Ver_2_1_30    Ver_2_1_45    Ver_2_1_56
        ' ** Ver_2_0_00    Ver_2_1_40    Ver_2_1_46    Ver_2_1_57

720     If blnContinue = True Then

          'For lngX = 0& To (lngVerDirs - 1&)
730       lngX = lngInstance

          ' ** Temporarily link to the \DemoDatabase data files.
          'Backend_Link "Demo", False  ' ** Module Function: zz_mod_Backend_Compare.

740       strThisPath = CurrentAppPath  ' ** Module Function: modFileUtilities.

          ' ** Replace the \EmptyDatabase set.
750       If FileExists(strThisPath & LNK_SEP & gstrDir_DevEmpty & LNK_SEP & gstrFile_DataName) = True Then  ' ** Module Function: modFileUtilities.
760         Kill (strThisPath & LNK_SEP & gstrDir_DevEmpty & LNK_SEP & gstrFile_DataName)
770       End If
780       If FileExists(strThisPath & LNK_SEP & gstrDir_DevEmpty & LNK_SEP & gstrFile_ArchDataName) = True Then  ' ** Module Function: modFileUtilities.
790         Kill (strThisPath & LNK_SEP & gstrDir_DevEmpty & LNK_SEP & gstrFile_ArchDataName)
800       End If
810       FileCopy (strThisPath & LNK_SEP & gstrDir_DevEmpty & LNK_SEP & "bak" & LNK_SEP & gstrFile_DataName), _
            (strThisPath & LNK_SEP & gstrDir_DevEmpty & LNK_SEP & gstrFile_DataName)
820       FileCopy (strThisPath & LNK_SEP & gstrDir_DevEmpty & LNK_SEP & "bak" & LNK_SEP & gstrFile_ArchDataName), _
            (strThisPath & LNK_SEP & gstrDir_DevEmpty & LNK_SEP & gstrFile_ArchDataName)

          ' ** Relink to the \EmptyDatabase data files.
          'Backend_Link "Empty", False  ' ** Module Function: zz_mod_Backend_Compare.

          ' ** Initialize gstrTrustDataLocation.
830       IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
840       strDataPath = Left(gstrTrustDataLocation, (Len(gstrTrustDataLocation) - 1))  ' ** Remove final backslash.

850       If FileExists(strConvertPath & LNK_SEP & gstrFile_DataName) = True Then  ' ** Module Function: modFileUtilities.
860         Kill (strConvertPath & LNK_SEP & gstrFile_DataName)
870       End If
880       If FileExists(strConvertPath & LNK_SEP & gstrFile_ArchDataName) = True Then  ' ** Module Function: modFileUtilities.
890         Kill (strConvertPath & LNK_SEP & gstrFile_ArchDataName)
900       End If
910       If FileExists(arr_varVerDir(D_PATH, lngX) & LNK_SEP & gstrFile_DataName) = False Then  ' ** Module Function: modFileUtilities.
920         blnContinue = False
930         Stop
940       End If
950       If FileExists(arr_varVerDir(D_PATH, lngX) & LNK_SEP & gstrFile_ArchDataName) = False Then  ' ** Module Function: modFileUtilities.
960         blnContinue = False
970         Stop
980       End If

990       If blnContinue = True Then

            ' ** Borrowing this for Directory Version numbers.
            ' ** Example: Ver_2_1_47
1000        gstrCrtRpt_Version = Mid(arr_varVerDir(D_DIR, lngX), 5, 1) & "." & _
              Mid(arr_varVerDir(D_DIR, lngX), 7, 1) & "." & Mid(arr_varVerDir(D_DIR, lngX), 9, 2)
            'Debug.Print "'" & gstrCrtRpt_Version
            'Stop
            ' ** Copy this test Version set to \Convert_New.
1010        FileCopy (arr_varVerDir(D_PATH, lngX) & LNK_SEP & gstrFile_DataName), _
              (strConvertPath & LNK_SEP & gstrFile_DataName)
1020        FileCopy (arr_varVerDir(D_PATH, lngX) & LNK_SEP & gstrFile_ArchDataName), _
              (strConvertPath & LNK_SEP & gstrFile_ArchDataName)

            ' ** Start the conversion routine.
1030        intConversionCheck = ConversionCheck  ' ** Module Function: modVersionConvertFuncs2.
            ' ** Return values:
            ' **    1  Unnecessary
            ' **    0  OK
            ' **   -1  Can't Connect
            ' **   -2  Can't Open {TrustDta.mdb}
            ' **   -3  Canceled Status
            ' **   -4  Acount Empty
            ' **   -5  Canceled CoInfo
            ' **   -6  Index/Key
            ' **   -7  Can't Open {TrstArch.mdb}
            ' **   -9  Error

1040        If intConversionCheck = 0 Then

              ' ** Rename the 2 BAK's with their version suffix.
              ' ** Example: Ver_2_1_47
1050          strTmp01 = Mid(arr_varVerDir(D_DIR, lngX), 5, 1) & Mid(arr_varVerDir(D_DIR, lngX), 7, 1) & Mid(arr_varVerDir(D_DIR, lngX), 9, 2)

1060          strTmp02 = Left(gstrFile_DataName, (Len(gstrFile_DataName) - 4)) & ".BAK"  ' ** Should be renamed by conversion routine.
1070          strTmp03 = Left(gstrFile_DataName, (Len(gstrFile_DataName) - 4)) & "_" & strTmp01 & ".mdb"
1080          If FileExists(strConvertPath & LNK_SEP & strTmp02) = False Then  ' ** Module Function: modFileUtilities.
1090            blnContinue = False
1100            Stop
1110          Else
1120            If FileExists(strConvertPath & LNK_SEP & strTmp03) = True Then  ' ** Module Function: modFileUtilities.
1130              Kill (strConvertPath & LNK_SEP & strTmp03)
1140            End If
1150            Name (strConvertPath & LNK_SEP & strTmp02) As (strConvertPath & LNK_SEP & strTmp03)
1160            strTmp02 = Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 4)) & ".BAK"  ' ** Should be renamed by conversion routine.
1170            strTmp03 = Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 4)) & "_" & strTmp01 & ".mdb"
1180            If FileExists(strConvertPath & LNK_SEP & strTmp02) = False Then  ' ** Module Function: modFileUtilities.
1190              blnContinue = False
1200              Stop
1210            Else
1220              If FileExists(strConvertPath & LNK_SEP & strTmp03) = True Then  ' ** Module Function: modFileUtilities.
1230                Kill (strConvertPath & LNK_SEP & strTmp03)
1240              End If
1250              Name (strConvertPath & LNK_SEP & strTmp02) As (strConvertPath & LNK_SEP & strTmp03)
1260            End If
1270          End If

1280        End If  ' ** intConversionCheck.

1290      End If  ' ** blnContinue.

1300      If blnContinue = False Then
            'Exit For
1310      End If
          'Next  ' ** For each version directory: lngX.
1320    End If  ' ** blnContinue.

1330    Beep

EXITP:
1340    Set fsfd2 = Nothing
1350    Set fsfd1 = Nothing
1360    Set fsfds1 = Nothing
1370    Set fso = Nothing
1380    Version_Convert_RunAll = blnRetVal
1390    Exit Function

ERRH:
1400    blnRetVal = False
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

      #End If

End Function

Public Function Version_Client_Chk() As Boolean
' ** Not called.

      #If IsDev Then

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Client_Chk"

        Dim wrkLoc As DAO.Workspace, wrkLnk As DAO.Workspace, wrkLnk2 As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database
        Dim qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim tdf As DAO.TableDef, fld As DAO.Field, doc As DAO.Document, prp As DAO.Property
        Dim lngDbs As Long, arr_varDb As Variant
        Dim intWrkType As Integer
        Dim strPath As String, strFile As String, strVersion As String
        Dim lngTbls As Long
        Dim blnFound As Boolean
        Dim strTmp01 As String, strTmp02 As String, lngTmp03 As Long, lngTmp04 As Long, lngTmp05 As Long
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varDb().
        Const DB_DIRID   As Integer = 0
        Const DB_PATH    As Integer = 1
        Const DB_ID      As Integer = 2
        Const DB_NAME    As Integer = 3
        Const DB_EXT     As Integer = 5
        Const DB_ISMDB   As Integer = 6
        Const DB_CANOPEN As Integer = 7
        Const DB_ACCVER  As Integer = 8
        Const DB_MVER    As Integer = 9
        Const DB_APPVER  As Integer = 10
        Const DB_APPDATE As Integer = 11
        Const DB_TBLCNT  As Integer = 12
        Const DB_NOTE    As Integer = 13

1510    blnRetVal = True

1520    Set wrkLoc = DBEngine.Workspaces(0)
1530    Set dbsLoc = wrkLoc.Databases(0)
1540    Set qdf = dbsLoc.QueryDefs("zz_qry_Client_02")
1550    Set rst = qdf.OpenRecordset
1560    With rst
1570      .MoveLast
1580      lngDbs = .RecordCount
1590      .MoveFirst
1600      arr_varDb = .GetRows(lngDbs)
          ' *********************************************************
          ' ** Array: arr_varDb()
          ' **
          ' **   Field  Element  Name                  Constant
          ' **   =====  =======  ====================  ============
          ' **     1       0     clidir_id             DB_DIRID
          ' **     2       1     clidir_path           DB_PATH
          ' **     3       2     clifile_id            DB_ID
          ' **     4       3     clifile_name          DB_NAME
          ' **     5       4     Ux
          ' **     6       5     clifile_ext           DB_EXT
          ' **     7       6     clifile_ismdb         DB_ISMDB
          ' **     8       7     clifile_canopen       DB_CANOPEN
          ' **     9       8     clifile_accver        DB_ACCVER
          ' **    10       9     clifile_m_v_ver       DB_MVER
          ' **    11      10     clifile_appversion    DB_APPVER
          ' **    12      11     clifile_appdate       DB_APPDATE
          ' **    13      12     clifile_tblcnt        DB_TBLCNT
          ' **    14      13     clifile_note          DB_NOTE
          ' **
          ' *********************************************************
1610      .Close
1620    End With

        ' ** Empty the Note in case it finds different results.
1630    For lngX = 0& To (lngDbs - 1&)
1640      arr_varDb(DB_NOTE, lngX) = Null
1650    Next

1660    intWrkType = 0
1670  On Error Resume Next
1680    Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
1690    If ERR.Number <> 0 Then
1700  On Error GoTo ERRH
1710  On Error Resume Next
1720      Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
1730      If ERR.Number <> 0 Then
1740  On Error GoTo ERRH
1750  On Error Resume Next
1760        Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
1770        If ERR.Number <> 0 Then
1780  On Error GoTo ERRH
1790  On Error Resume Next
1800          Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
1810          If ERR.Number <> 0 Then
1820  On Error GoTo ERRH
1830  On Error Resume Next
1840            Set wrkLnk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
1850            If ERR.Number <> 0 Then
1860  On Error GoTo ERRH
1870  On Error Resume Next
1880              Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
1890              If ERR.Number <> 0 Then
1900  On Error GoTo ERRH
1910  On Error Resume Next
1920                Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
1930  On Error GoTo ERRH
1940                intWrkType = 7
1950              Else
1960  On Error GoTo ERRH
1970                intWrkType = 6
1980              End If
1990            Else
2000  On Error GoTo ERRH
2010              intWrkType = 5
2020            End If
2030          Else
2040  On Error GoTo ERRH
2050            intWrkType = 4
2060          End If
2070        Else
2080  On Error GoTo ERRH
2090          intWrkType = 3
2100        End If
2110      Else
2120  On Error GoTo ERRH
2130        intWrkType = 2
2140      End If
2150    Else
2160  On Error GoTo ERRH
2170      intWrkType = 1
2180    End If

2190    Set wrkLnk2 = CreateWorkspace("TmpAdmin", "Admin", "deltadata", dbUseJet)

2200    With wrkLnk

2210      lngTmp03 = 0&
2220      For lngX = 0& To (lngDbs - 1&)

2230        arr_varDb(DB_ISMDB, lngX) = CBool(True)  ' ** Making an assumption here.
2240        strFile = arr_varDb(DB_NAME, lngX)
2250        strPath = arr_varDb(DB_PATH, lngX)
2260        strVersion = vbNullString
2270        strTmp01 = vbNullString: strTmp02 = vbNullString

2280        If arr_varDb(DB_EXT, lngX) <> "mdb" Then
2290          strTmp01 = Left(strFile, (Len(strFile) - (Len(arr_varDb(DB_EXT, lngX)) + 1)))
2300          strTmp02 = strTmp01 & ".mdb"
2310          If Dir(strPath & LNK_SEP & strTmp02) = vbNullString Then
                ' ** Name oldpathname As newpathname
2320            Name (strPath & LNK_SEP & strFile) As (strPath & LNK_SEP & strTmp02)
2330          Else
2340            lngTmp03 = lngTmp03 + 1&
2350            strTmp02 = strTmp01 & CStr(lngTmp03) & ".mdb"
2360            If Dir(strPath & LNK_SEP & strTmp02) = vbNullString Then
2370              Name (strPath & LNK_SEP & strFile) As (strPath & LNK_SEP & strTmp02)
2380            Else
2390              lngTmp03 = lngTmp03 + 1&
2400              strTmp02 = strTmp01 & CStr(lngTmp03) & ".mdb"
2410              If Dir(strPath & LNK_SEP & strTmp02) = vbNullString Then
2420                Name (strPath & LNK_SEP & strFile) As (strPath & LNK_SEP & strTmp02)
2430              Else
2440                Stop
2450              End If
2460            End If
2470          End If
2480        Else
2490          strTmp02 = strFile
2500        End If

2510  On Error Resume Next
2520        Set dbsLnk = .OpenDatabase(strPath & LNK_SEP & strTmp02, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
2530        If ERR.Number <> 0 Then
2540          arr_varDb(DB_NOTE, lngX) = ("Error: " & CStr(ERR.Number)) & ", " & ERR.description
2550  On Error GoTo 0
              ' ** Leave DB_CANOPEN = False.
2560        Else
2570  On Error GoTo 0

              ' ** OK, we got it open. Let's see what we've got.
2580          arr_varDb(DB_CANOPEN, lngX) = CBool(True)
2590          With dbsLnk
2600            lngTbls = 0&

2610            For Each tdf In .TableDefs
2620              With tdf
2630                If Left(.Name, 4) <> "MSys" And Left(.Name, 4) <> "~TMP" And .Connect = vbNullString Then  ' ** Skip those pesky system tables.
2640                  lngTbls = lngTbls + 1&
2650                  Select Case .Name
                      Case "m_VA"
2660                    Set rst = dbsLnk.OpenRecordset("m_VA", dbOpenDynaset, dbReadOnly)
2670                    With rst
2680                      If .BOF = True And .EOF = True Then
                            ' ** That's a bust!
2690                      Else
2700                        .MoveFirst
2710                        strVersion = Nz(![va_MAIN], "0") & "."
2720                        strVersion = strVersion & Nz(![va_MINOR], "0") & "."
2730                        strVersion = strVersion & Nz(![va_REVISION], "0")
2740                        arr_varDb(DB_MVER, lngX) = strVersion
2750                      End If
2760                      .Close
2770                    End With
2780                  Case "m_VD"
2790                    Set rst = dbsLnk.OpenRecordset("m_VD", dbOpenDynaset, dbReadOnly)
2800                    With rst
2810                      If .BOF = True And .EOF = True Then
                            ' ** That's a bust!
2820                      Else
2830                        .MoveFirst
2840                        strVersion = Nz(![vd_MAIN], "0") & "."
2850                        strVersion = strVersion & Nz(![vd_MINOR], "0") & "."
2860                        strVersion = strVersion & Nz(![vd_REVISION], "0")
2870                        arr_varDb(DB_MVER, lngX) = strVersion
2880                      End If
2890                      .Close
2900                    End With
2910                  Case "m_VP"  ' ** Though I'm not expecting to find one of these.
2920                    Set rst = dbsLnk.OpenRecordset("m_VP", dbOpenDynaset, dbReadOnly)
2930                    With rst
2940                      If .BOF = True And .EOF = True Then
                            ' ** That's a bust!
2950                      Else
2960                        .MoveFirst
2970                        strVersion = Nz(![vp_MAIN], "0") & "."
2980                        strVersion = strVersion & Nz(![vp_MINOR], "0") & "."
2990                        strVersion = strVersion & Nz(![vp_REVISION], "0")
3000                        arr_varDb(DB_MVER, lngX) = strVersion
3010                      End If
3020                      .Close
3030                    End With
3040                  Case Else
                        ' ** Check other tables that might tell the version?
3050                  End Select
3060                End If
3070              End With  ' ** This TableDef: tdf.
3080            Next  ' ** For each TableDef: tdf.
3090            arr_varDb(DB_TBLCNT, lngX) = lngTbls

3100          End With  ' ** dbsLnk.

3110          If lngTbls = 0& Then
                ' ** Try the alternate workspace.
3120            dbsLnk.Close
3130            Set dbsLnk = wrkLnk2.OpenDatabase(strPath & LNK_SEP & strTmp02, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
3140            With dbsLnk
3150              For Each tdf In .TableDefs
3160                With tdf
3170                  If Left(.Name, 4) <> "MSys" And Left(.Name, 4) <> "~TMP" And .Connect = vbNullString Then  ' ** Skip those pesky system tables.
3180                    lngTbls = lngTbls + 1&
3190                  End If
3200                End With
3210              Next
3220              If lngTbls = 0& Then
                    ' ** It may be a truly empty database.
3230                If .TableDefs.Count = 5 Then
3240                  lngTmp04 = 0&
3250                  For Each tdf In .TableDefs
3260                    Select Case tdf.Name
                        Case "MSysRelationships"
3270                      lngTmp04 = lngTmp04 + 1&
3280                    Case "MSysQueries"
3290                      lngTmp04 = lngTmp04 + 1&
3300                    Case "MSysObjects"
3310                      lngTmp04 = lngTmp04 + 1&
3320                    Case "MSysACEs"
3330                      lngTmp04 = lngTmp04 + 1&
3340                    Case "MSysAccessObjects"
3350                      lngTmp04 = lngTmp04 + 1&
3360                    End Select
3370                  Next
3380                  If lngTmp04 = 5 Then
3390                    arr_varDb(DB_NOTE, lngX) = "Empty database"
3400                  End If
3410                Else
3420                  arr_varDb(DB_NOTE, lngX) = "No user tables found; TableDefs.Count = " & CStr(.TableDefs.Count)
3430                End If
3440              End If
3450            End With
                ' ** Leave this one open for the rest of the loop.
3460          End If

3470          With dbsLnk

3480            arr_varDb(DB_ACCVER, lngX) = .Containers("Databases").Documents("MSysDb").Properties("AccessVersion")
                ' ** CurrentDb.Containers("Databases").Documents("MSysDb").Properties("AccessVersion") = 08.50
3490            Select Case arr_varDb(DB_ACCVER, lngX)
                Case "02.00"
3500              arr_varDb(DB_ACCVER, lngX) = "Access 2.0"
3510            Case "06.68"
3520              arr_varDb(DB_ACCVER, lngX) = "Access 95"
3530            Case "07.53"
3540              arr_varDb(DB_ACCVER, lngX) = "Access 97"
3550            Case "08.50"
3560              arr_varDb(DB_ACCVER, lngX) = "Access 2000"
3570            Case "09.50"
3580              arr_varDb(DB_ACCVER, lngX) = "Access 2002/2003"
3590            Case Else
                  ' ** The Jet MDW.
3600            End Select

3610            For Each doc In .Containers("Databases").Documents
3620              With doc
3630                If .Name = "UserDefined" Then
3640                  For Each prp In .Properties
3650                    With prp
3660                      If .Name = "AppVersion" Then
3670                        arr_varDb(DB_APPVER, lngX) = .Value
3680                      ElseIf .Name = "AppDate" Then
3690                        arr_varDb(DB_APPDATE, lngX) = .Value
3700                      End If
3710                    End With  ' ** prp.
3720                  Next
3730                End If
3740              End With  ' ** doc.
3750            Next

3760            .Close
3770          End With  ' ** dbsLnk.

3780        End If  ' ** Can open.

3790        If strTmp02 <> strFile Then
              ' ** Rename it back to what it was.
3800          Name (strPath & LNK_SEP & strTmp02) As (strPath & LNK_SEP & strFile)
3810        End If

3820        Set dbsLnk = Nothing
3830      Next  ' ** For each possible MDB: lngX.

          ' ** Check for missing version numbers.
3840      For lngX = 0& To (lngDbs - 1&)
3850        If arr_varDb(DB_CANOPEN, lngX) = True And IsNull(arr_varDb(DB_NOTE, lngX)) = True Then
              ' ** If it's an empty database, there'll be a note about it, from above.
3860          blnFound = False
3870          If IsNull(arr_varDb(DB_MVER, lngX)) = True Then
3880            blnFound = True
3890          Else
3900            If arr_varDb(DB_MVER, lngX) = vbNullString Then
3910              blnFound = True
3920            End If
3930          End If
3940          If blnFound = True Then
3950            If InStr(arr_varDb(DB_NAME, lngX), "Arch") > 0 Then
                  ' ** Get a list of all samples with the same table count.
3960              Select Case arr_varDb(DB_TBLCNT, lngX)
                  Case 1&
3970                Set qdf = dbsLoc.QueryDefs("qryVersion_88b")
3980              Case 2&
3990                Stop
4000              End Select
4010              Set rst = qdf.OpenRecordset
4020              With rst
                    ' ** This is a universe of possible versions.
4030                .MoveLast
4040                lngTmp04 = .RecordCount
4050                .MoveFirst
4060                If lngTmp04 = 1& Then
                      ' ** Obviously only one choice (though I know it won't happen here!).
                      ' ** Example: Ver_1_7_20
4070                  strTmp01 = ![verdir_name]
4080                  strTmp01 = Mid(strTmp01, (InStr(strTmp01, "_") + 1))
4090                  strTmp01 = Left(strTmp01, 1) & "." & Mid(strTmp01, 3, 1) & "." & Mid(strTmp01, 5)
4100                  arr_varDb(DB_MVER, lngX) = strTmp01
4110                Else
                      ' ** Try narrowing it down by the number of fields in the tables present.
4120                  Select Case arr_varDb(DB_TBLCNT, lngX)
                      Case 1&
                        ' ** Ledger only present.
4130                    strFile = arr_varDb(DB_NAME, lngX)
4140                    strPath = arr_varDb(DB_PATH, lngX)
4150                    strVersion = vbNullString
4160                    strTmp01 = vbNullString: strTmp02 = vbNullString
4170                    If arr_varDb(DB_EXT, lngX) <> "mdb" Then
4180                      strTmp01 = Left(strFile, (Len(strFile) - (Len(arr_varDb(DB_EXT, lngX)) + 1)))
4190                      strTmp02 = strTmp01 & ".mdb"
4200                      If Dir(strPath & LNK_SEP & strTmp02) = vbNullString Then
                            ' ** Name oldpathname As newpathname
4210                        Name (strPath & LNK_SEP & strFile) As (strPath & LNK_SEP & strTmp02)
4220                      Else
4230                        lngTmp03 = lngTmp03 + 1&
4240                        strTmp02 = strTmp01 & CStr(lngTmp03) & ".mdb"
4250                        If Dir(strPath & LNK_SEP & strTmp02) = vbNullString Then
4260                          Name (strPath & LNK_SEP & strFile) As (strPath & LNK_SEP & strTmp02)
4270                        Else
4280                          lngTmp03 = lngTmp03 + 1&
4290                          strTmp02 = strTmp01 & CStr(lngTmp03) & ".mdb"
4300                          If Dir(strPath & LNK_SEP & strTmp02) = vbNullString Then
4310                            Name (strPath & LNK_SEP & strFile) As (strPath & LNK_SEP & strTmp02)
4320                          Else
4330                            Stop
4340                          End If
4350                        End If
4360                      End If
4370                    Else
4380                      strTmp02 = strFile
4390                    End If
4400                    lngTmp05 = 0&
4410                    Set dbsLnk = wrkLnk.OpenDatabase(strPath & LNK_SEP & strTmp02, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
4420                    Set tdf = dbsLnk.TableDefs("ledger")
4430                    lngTmp03 = tdf.Fields.Count  ' ** (I think lngTmp03 is now available.)
                        ' ** See if that matches only 1, or still several alternatives.
4440                    For lngY = 1& To lngTmp04
4450                      If ![FldCnt] = lngTmp03 Then
4460                        lngTmp05 = lngTmp05 + 1&
4470                      End If
4480                      If lngY < lngTmp04 Then .MoveNext
4490                    Next
4500                    .MoveFirst
4510                    If lngTmp05 = 1& Then
                          ' ** Yay! We got a hit!
4520                      strTmp01 = ![verdir_name]
4530                      strTmp01 = Mid(strTmp01, (InStr(strTmp01, "_") + 1))
4540                      strTmp01 = Left(strTmp01, 1) & "." & Mid(strTmp01, 3, 1) & "." & Mid(strTmp01, 5)
4550                      arr_varDb(DB_MVER, lngX) = strTmp01
4560                    Else
                          ' ** Will field names or properties help home-in?
                          'For 19 fields, these all match for properties so far documented.
                          'Ver_1_7_00
                          'Ver_1_7_10
                          'Ver_1_7_20
                          'Ver_1_7_40
                          'Ver_2_0_00
4570                      strTmp01 = "1.7x-2.0.0"  ' ** 10 chars max.
4580                      arr_varDb(DB_MVER, lngX) = strTmp01
4590                    End If
4600                    dbsLnk.Close
4610                    If strTmp02 <> strFile Then
                          ' ** Rename it back to what it was.
4620                      Name (strPath & LNK_SEP & strTmp02) As (strPath & LNK_SEP & strFile)
4630                    End If
4640                    Set dbsLnk = Nothing
4650                  Case 2&
                        ' ** Ledger and m_VA present.
4660                    Stop
4670                  Case Else
                        ' ** There shouldn't be any other possibilities.
4680                  End Select
4690                End If
4700                .Close
4710              End With
4720            Else
4730              Stop
4740            End If
4750          End If
4760        End If
4770      Next

4780      .Close
4790    End With  ' ** wrkLnk.

        ' *********************************************************
        ' ** Array: arr_varDb()
        ' **
        ' **   Field  Element  Name                  Constant
        ' **   =====  =======  ====================  ============
        ' **     1       0     clidir_id             DB_DIRID
        ' **     2       1     clidir_path           DB_PATH
        ' **     3       2     clifile_id            DB_ID
        ' **     4       3     clifile_name          DB_NAME
        ' **     5       4     Ux
        ' **     6       5     clifile_ext           DB_EXT
        ' **     7       6     clifile_ismdb         DB_ISMDB
        ' **     8       7     clifile_canopen       DB_CANOPEN
        ' **     9       8     clifile_accver        DB_ACCVER
        ' **    10       9     clifile_m_v_ver       DB_MVER
        ' **    11      10     clifile_appversion    DB_APPVER
        ' **    12      11     clifile_appdate       DB_APPDATE
        ' **    13      12     clifile_tblcnt        DB_TBLCNT
        ' **    14      13     clifile_note          DB_NOTE
        ' **
        ' *********************************************************
4800    With dbsLoc
4810      Set rst = .OpenRecordset("zz_tbl_Client_File", dbOpenDynaset, dbConsistent)
4820      With rst
4830        .MoveFirst
4840        For lngX = 0& To (lngDbs - 1&)
4850          .FindFirst "[clidir_id] = " & arr_varDb(DB_DIRID, lngX) & " And " & _
                "[clifile_id] = " & arr_varDb(DB_ID, lngX)
4860          If .NoMatch = False Then
4870            .Edit
4880            ![clifile_ismdb] = arr_varDb(DB_ISMDB, lngX)
4890            ![clifile_canopen] = arr_varDb(DB_CANOPEN, lngX)
4900            ![clifile_accver] = arr_varDb(DB_ACCVER, lngX)
4910            If IsNull(arr_varDb(DB_MVER, lngX)) = False Then
4920              ![clifile_m_v_ver] = arr_varDb(DB_MVER, lngX)
4930            End If
4940            If IsNull(arr_varDb(DB_APPVER, lngX)) = False Then
4950              ![clifile_appversion] = arr_varDb(DB_APPVER, lngX)
4960            End If
4970            If IsNull(arr_varDb(DB_APPDATE, lngX)) = False Then
4980              If CLng(arr_varDb(DB_APPDATE, lngX)) <> 0 Then
4990                ![clifile_appdate] = arr_varDb(DB_APPDATE, lngX)
5000              End If
5010            End If
5020            If IsNull(arr_varDb(DB_TBLCNT, lngX)) = False Then
5030              ![clifile_tblcnt] = arr_varDb(DB_TBLCNT, lngX)
5040            End If
5050            If IsNull(arr_varDb(DB_NOTE, lngX)) = False Then
5060              ![clifile_note] = arr_varDb(DB_NOTE, lngX)
5070            Else
5080              If IsNull(![clifile_note]) = False Then
5090                ![clifile_note] = Null
5100              End If
5110            End If
5120            ![clifile_datemodified] = Now()
5130            .Update
5140          Else
5150            Stop
5160          End If
5170        Next
5180      End With
5190      .Close
5200    End With

5210    wrkLnk2.Close
5220    wrkLoc.Close

5230    Beep

EXITP:
5240    Set rst = Nothing
5250    Set qdf = Nothing
5260    Set prp = Nothing
5270    Set doc = Nothing
5280    Set fld = Nothing
5290    Set tdf = Nothing
5300    Set dbsLnk = Nothing
5310    Set dbsLoc = Nothing
5320    Set wrkLnk2 = Nothing
5330    Set wrkLnk = Nothing
5340    Set wrkLoc = Nothing
5350    Version_Client_Chk = blnRetVal
5360    Exit Function

ERRH:
5370    blnRetVal = False
5380    Select Case ERR.Number
        Case Else
5390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5400    End Select
5410    Resume EXITP

      #End If

End Function

Public Function Version_Client_Doc() As Boolean
' ** Not called.

      #If IsDev Then

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Client_Doc"

        Dim wrkLnk As DAO.Workspace, dbsLnk As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field, doc As DAO.Document, prp As DAO.Property
        Dim wrkLoc As DAO.Workspace, dbsLoc As DAO.Database
        Dim rst1 As DAO.Recordset
        Dim fso As Scripting.FileSystemObject
        Dim fsfds1 As Scripting.Folders, fsfds2 As Scripting.Folders, fsfds3 As Scripting.Folders
        Dim fsfds4 As Scripting.Folders, fsfds5 As Scripting.Folders, fsfds6 As Scripting.Folders
        Dim fsfds7 As Scripting.Folders
        Dim fsfd1 As Scripting.Folder, fsfd2 As Scripting.Folder, fsfd3 As Scripting.Folder
        Dim fsfd4 As Scripting.Folder, fsfd5 As Scripting.Folder, fsfd6 As Scripting.Folder
        Dim fsfd7 As Scripting.Folder
        Dim fsfls As Scripting.Files, fsfl As Scripting.File
        Dim lngCliDirs As Long, arr_varCliDir() As Variant
        Dim lngCliFiles As Long, arr_varCliFile() As Variant, arr_varTmpCliFile As Variant
        Dim strBasePath As String, strThisDir As String
        Dim lngCliDirID As Long, lngCliFileID As Long, lngCliTblID As Long
        Dim lngSubCnt1 As Long, lngSubCnt2 As Long, lngSubCnt3 As Long, lngSubCnt4 As Long, lngSubCnt5 As Long, lngSubCnt6 As Long
        Dim blnDocThroughFilesOnly As Boolean, blnAddAll As Boolean, blnAddThis As Boolean, blnEditThis As Boolean
        Dim blnFound As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String, strTmp05 As String
        Dim lngTmp18 As Long, lngTmp19 As Long, lngTmp20 As Long, lngTmp21 As Long, lngTmp22 As Long
        Dim lngX As Long, lngY As Long, lngE As Long, lngF As Long, lngG As Long, lngH As Long, lngI As Long, lngJ As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varCliDir().
        Const C_ELEMS As Integer = 7  ' ** Array's first-element UBound().
        Const C_ID     As Integer = 0
        Const C_NAM    As Integer = 1
        Const C_PARID  As Integer = 2
        Const C_PARNAM As Integer = 3
        Const C_PATH   As Integer = 4
        Const C_FILES  As Integer = 5
        Const C_SUBS   As Integer = 6
        Const C_F_ARR  As Integer = 7

        ' ** Array: arr_varCliFile().
        Const F_ELEMS As Integer = 13  ' ** Array's first-element UBound().
        Const F_DIRID   As Integer = 0
        Const F_ID      As Integer = 1
        Const F_NAM     As Integer = 2
        Const F_CLIENT  As Integer = 3
        Const F_ISMDB   As Integer = 4
        Const F_CANOPEN As Integer = 5
        Const F_FILSIZ  As Integer = 6
        Const F_TBLS    As Integer = 7
        Const F_T_ARR   As Integer = 8
        Const F_ACC_VER As Integer = 9
        Const F_M_VER   As Integer = 10
        Const F_APPVER  As Integer = 11
        Const F_APPDATE As Integer = 12
        Const F_NOTE    As Integer = 13

5510    blnRetVal = True
5520    blnDocThroughFilesOnly = False

5530    Set wrkLoc = DBEngine.Workspaces(0)
5540    Set dbsLoc = wrkLoc.Databases(0)

5550    strBasePath = "C:\VictorGCS_Clients\TrustAccountant\Clients"

5560    lngCliDirs = 0&
5570    ReDim arr_varCliDir(C_ELEMS, 0)

5580    Set fso = CreateObject("Scripting.FileSystemObject")
5590    With fso

5600      Set fsfd1 = .GetFolder(strBasePath)
5610      Set fsfds1 = fsfd1.SubFolders

5620      For Each fsfd2 In fsfds1
5630        With fsfd2
5640          lngCliDirs = lngCliDirs + 1&
5650          lngE = lngCliDirs - 1&
5660          ReDim Preserve arr_varCliDir(C_ELEMS, lngE)
              ' **********************************************
              ' ** Array: arr_varCliDir()
              ' **
              ' **   Element  Name              Constant
              ' **   =======  ================  ============
              ' **      0     clidir_id         C_ID
              ' **      1     clidir_name       C_NAM
              ' **      2     clidir_parid      C_PARID
              ' **      3     clidir_parname    C_PARNAM
              ' **      4     clidir_path       C_PATH
              ' **      5     clidir_filecnt    C_FILES
              ' **      6     clidir_subcnt     C_SUBS
              ' **      7     File Array        C_F_ARR
              ' **
              ' **********************************************
5670          arr_varCliDir(C_ID, lngE) = CLng(0)
5680          arr_varCliDir(C_NAM, lngE) = .Name
5690          arr_varCliDir(C_PARID, lngE) = CLng(0)
5700          arr_varCliDir(C_PARNAM, lngE) = vbNullString
5710          arr_varCliDir(C_PATH, lngE) = .Path
5720          arr_varCliDir(C_FILES, lngE) = CLng(0)
5730          arr_varCliDir(C_F_ARR, lngE) = Empty
              ' ** Yes, we could do some sort of iterative loop for the
              ' ** subdirectories, but I just didn't feel like it today.
5740          Set fsfds2 = fsfd2.SubFolders
5750          lngSubCnt1 = fsfds2.Count
5760          arr_varCliDir(C_SUBS, lngE) = lngSubCnt1
5770          If lngSubCnt1 > 0& Then
5780            For Each fsfd3 In fsfds2
5790              With fsfd3
5800                lngCliDirs = lngCliDirs + 1&
5810                lngF = lngCliDirs - 1&
5820                ReDim Preserve arr_varCliDir(C_ELEMS, lngF)
5830                arr_varCliDir(C_ID, lngF) = CLng(0)
5840                arr_varCliDir(C_NAM, lngF) = .Name
5850                arr_varCliDir(C_PARID, lngF) = CLng(0)
5860                arr_varCliDir(C_PARNAM, lngF) = arr_varCliDir(C_NAM, lngE)
5870                arr_varCliDir(C_PATH, lngF) = .Path
5880                arr_varCliDir(C_FILES, lngF) = CLng(0)
5890                arr_varCliDir(C_F_ARR, lngF) = Empty
5900                Set fsfds3 = fsfd3.SubFolders
5910                lngSubCnt2 = fsfds3.Count
5920                arr_varCliDir(C_SUBS, lngF) = lngSubCnt2
5930                If lngSubCnt2 > 0& Then
5940                  For Each fsfd4 In fsfds3
5950                    With fsfd4
5960                      lngCliDirs = lngCliDirs + 1&
5970                      lngG = lngCliDirs - 1&
5980                      ReDim Preserve arr_varCliDir(C_ELEMS, lngG)
5990                      arr_varCliDir(C_ID, lngG) = CLng(0)
6000                      arr_varCliDir(C_NAM, lngG) = .Name
6010                      arr_varCliDir(C_PARID, lngG) = CLng(0)
6020                      arr_varCliDir(C_PARNAM, lngG) = arr_varCliDir(C_NAM, lngE) & ";" & arr_varCliDir(C_NAM, lngF)
6030                      arr_varCliDir(C_PATH, lngG) = .Path
6040                      arr_varCliDir(C_FILES, lngG) = CLng(0)
6050                      arr_varCliDir(C_F_ARR, lngG) = Empty
6060                      Set fsfds4 = fsfd4.SubFolders
6070                      lngSubCnt3 = fsfds4.Count
6080                      arr_varCliDir(C_SUBS, lngG) = lngSubCnt3
6090                      If lngSubCnt3 > 0& Then
6100                        For Each fsfd5 In fsfds4
6110                          With fsfd5
6120                            lngCliDirs = lngCliDirs + 1&
6130                            lngH = lngCliDirs - 1&
6140                            ReDim Preserve arr_varCliDir(C_ELEMS, lngH)
6150                            arr_varCliDir(C_ID, lngH) = CLng(0)
6160                            arr_varCliDir(C_NAM, lngH) = .Name
6170                            arr_varCliDir(C_PARID, lngH) = CLng(0)
6180                            arr_varCliDir(C_PARNAM, lngH) = arr_varCliDir(C_NAM, lngE) & ";" & arr_varCliDir(C_NAM, lngF) & _
                                  ";" & arr_varCliDir(C_NAM, lngG)
6190                            arr_varCliDir(C_PATH, lngH) = .Path
6200                            arr_varCliDir(C_FILES, lngH) = CLng(0)
6210                            arr_varCliDir(C_F_ARR, lngH) = Empty
6220                            Set fsfds5 = fsfd5.SubFolders
6230                            lngSubCnt4 = fsfds5.Count
6240                            arr_varCliDir(C_SUBS, lngH) = lngSubCnt4
6250                            If lngSubCnt4 > 0& Then
6260                              For Each fsfd6 In fsfds5
6270                                With fsfd6
6280                                  lngCliDirs = lngCliDirs + 1&
6290                                  lngI = lngCliDirs - 1&
6300                                  ReDim Preserve arr_varCliDir(C_ELEMS, lngI)
6310                                  arr_varCliDir(C_ID, lngI) = CLng(0)
6320                                  arr_varCliDir(C_NAM, lngI) = .Name
6330                                  arr_varCliDir(C_PARID, lngI) = CLng(0)
6340                                  arr_varCliDir(C_PARNAM, lngI) = arr_varCliDir(C_NAM, lngE) & ";" & arr_varCliDir(C_NAM, lngF) & _
                                        ";" & arr_varCliDir(C_NAM, lngG) & ";" & arr_varCliDir(C_NAM, lngH)
6350                                  arr_varCliDir(C_PATH, lngI) = .Path
6360                                  arr_varCliDir(C_FILES, lngI) = CLng(0)
6370                                  arr_varCliDir(C_F_ARR, lngI) = Empty
6380                                  Set fsfds6 = fsfd6.SubFolders
6390                                  lngSubCnt5 = fsfds6.Count
6400                                  arr_varCliDir(C_SUBS, lngI) = lngSubCnt5
6410                                  If lngSubCnt5 > 0& Then
6420                                    For Each fsfd7 In fsfds6
6430                                      With fsfd7
6440                                        lngCliDirs = lngCliDirs + 1&
6450                                        lngJ = lngCliDirs - 1&
6460                                        ReDim Preserve arr_varCliDir(C_ELEMS, lngJ)
6470                                        arr_varCliDir(C_ID, lngJ) = CLng(0)
6480                                        arr_varCliDir(C_NAM, lngJ) = .Name
6490                                        arr_varCliDir(C_PARID, lngJ) = CLng(0)
6500                                        arr_varCliDir(C_PARNAM, lngJ) = arr_varCliDir(C_NAM, lngE) & ";" & arr_varCliDir(C_NAM, lngF) & _
                                              ";" & arr_varCliDir(C_NAM, lngG) & ";" & arr_varCliDir(C_NAM, lngH) & _
                                              ";" & arr_varCliDir(C_NAM, lngI)
6510                                        arr_varCliDir(C_PATH, lngJ) = .Path
6520                                        arr_varCliDir(C_FILES, lngJ) = CLng(0)
6530                                        arr_varCliDir(C_F_ARR, lngJ) = Empty
6540                                        Set fsfds7 = fsfd7.SubFolders
6550                                        lngSubCnt6 = fsfds7.Count
6560                                        arr_varCliDir(C_SUBS, lngJ) = lngSubCnt6
6570                                        If lngSubCnt6 > 0& Then
6580                                          Stop

6590                                        End If
6600                                      End With
6610                                    Next
6620                                  End If
6630                                End With
6640                              Next
6650                            End If
6660                          End With
6670                        Next
6680                      End If
6690                    End With
6700                  Next
6710                End If
6720              End With  ' ** This subfolder: fsfd3.
6730            Next  ' ** For each subfolder: fsfd3.
6740          End If  ' ** Has subfolders: lngSubCnt1.
6750          Set fsfds2 = Nothing
6760        End With  ' ** This main client subfolder: fsfd2.
6770      Next  ' ** For each main client subfolder: fsfd2.

6780      For lngX = 0& To (lngCliDirs - 1&)

6790        strThisDir = arr_varCliDir(C_NAM, lngX)
6800        Set fsfd2 = .GetFolder(arr_varCliDir(C_PATH, lngX))
6810        Set fsfls = fsfd2.Files

6820        lngCliFiles = 0&
6830        ReDim arr_varCliFile(F_ELEMS, 0)

6840        For Each fsfl In fsfls
6850          With fsfl
6860            lngCliFiles = lngCliFiles + 1&
6870            lngE = lngCliFiles - 1&
6880            ReDim Preserve arr_varCliFile(F_ELEMS, lngE)
                ' *************************************************
                ' ** Array: arr_varCliFile()
                ' **
                ' **   Element  Name                  Constant
                ' **   =======  ====================  ===========
                ' **      0     clidir_id             F_DIRID
                ' **      1     clifile_id            F_ID
                ' **      2     clifile_name          F_NAM
                ' **      3     clifile_client        F_CLIENT
                ' **      4     clifile_ismdb         F_ISMDB
                ' **      5     clifile_canopen       F_CANOPEN
                ' **      6     clifile_filesize      F_FILSIZ
                ' **      7     clifile_tblcnt        F_TBLS
                ' **      8     Table Array           F_T_ARR
                ' **      9     clifile_accver        F_ACC_VER
                ' **     10     clifile_m_v_ver       F_M_VER
                ' **     11     clifile_appversion    F_APPVER
                ' **     12     clifile_appdate       F_APPDATE
                ' **     13     clifile_note          F_NOTE
                ' **
                ' *************************************************
6890            arr_varCliFile(F_DIRID, lngE) = CLng(0)
6900            arr_varCliFile(F_ID, lngE) = CLng(0)
6910            arr_varCliFile(F_NAM, lngE) = .Name
6920            arr_varCliFile(F_CLIENT, lngE) = strThisDir
6930            arr_varCliFile(F_ISMDB, lngE) = CBool(False)
6940            arr_varCliFile(F_CANOPEN, lngE) = CBool(False)
6950            arr_varCliFile(F_FILSIZ, lngE) = .Size
6960            arr_varCliFile(F_TBLS, lngE) = CLng(0)
6970            arr_varCliFile(F_T_ARR, lngE) = Empty
6980            arr_varCliFile(F_ACC_VER, lngE) = vbNullString
6990            arr_varCliFile(F_M_VER, lngE) = vbNullString
7000            arr_varCliFile(F_APPVER, lngE) = vbNullString
7010            arr_varCliFile(F_APPDATE, lngE) = Null
7020            arr_varCliFile(F_NOTE, lngE) = vbNullString
7030          End With  ' ** This file: fsfl.
7040        Next  ' ** For each file: fsfl.

7050        arr_varCliDir(C_FILES, lngX) = lngCliFiles
7060        arr_varCliDir(C_F_ARR, lngX) = arr_varCliFile

7070      Next  ' ** For each client directory: lngX.

7080      With dbsLoc

7090        blnAddAll = False: blnAddThis = False: blnEditThis = False
7100        Set rst1 = .OpenRecordset("zz_tbl_Client_Directory", dbOpenDynaset, dbConsistent)
7110        With rst1
7120          If .BOF = True And .EOF = True Then
                ' ** Nothing there yet.
7130            blnAddAll = True
7140          End If
7150          lngCliDirID = 0&: lngCliFileID = 0&: lngCliTblID = 0&
7160          For lngX = 0& To (lngCliDirs - 1&)
7170            blnAddThis = False: blnEditThis = False
7180            If blnAddAll = False Then
7190              .FindFirst "[clidir_path] = '" & arr_varCliDir(C_PATH, lngX) & "'"
7200              If .NoMatch = True Then
7210                blnAddThis = True
7220              Else
7230                blnEditThis = True
7240                lngCliDirID = ![clidir_id]
7250              End If
7260            Else
7270              blnAddThis = True
7280            End If
7290            If blnAddThis = True Or blnEditThis = True Then
7300              strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString: strTmp05 = vbNullString
7310              lngTmp18 = 0&: lngTmp19 = 0&: lngTmp20 = 0&: lngTmp21 = 0&: lngTmp22 = 0&
7320              If blnAddThis = True Then
7330                .AddNew
7340              Else
7350                .Edit
7360              End If
                  ' **********************************************
                  ' ** Array: arr_varCliDir()
                  ' **
                  ' **   Element  Name              Constant
                  ' **   =======  ================  ============
                  ' **      0     clidir_id         C_ID
                  ' **      1     clidir_name       C_NAM
                  ' **      2     clidir_parid      C_PARID
                  ' **      3     clidir_parname    C_PARNAM
                  ' **      4     clidir_path       C_PATH
                  ' **      5     clidir_filecnt    C_FILES
                  ' **      6     clidir_subcnt     C_SUBS
                  ' **      7     File Array        C_F_ARR
                  ' **
                  ' **********************************************
7370              ![clidir_name] = arr_varCliDir(C_NAM, lngX)
7380              If arr_varCliDir(C_PARNAM, lngX) <> vbNullString Then
7390                strTmp01 = arr_varCliDir(C_PARNAM, lngX)
7400                ![clidir_parname] = strTmp01
7410                intPos01 = InStr(strTmp01, ";")
7420                If intPos01 = 0 Then
                      ' ** 1st level subfolder.
7430                  blnFound = False
7440                  For lngY = 0& To (lngCliDirs - 1&)
7450                    If arr_varCliDir(C_ID, lngY) > 0& And arr_varCliDir(C_PARID, lngY) = 0& And _
                            arr_varCliDir(C_NAM, lngY) = strTmp01 Then
7460                      blnFound = True
7470                      ![clidir_parid] = arr_varCliDir(C_ID, lngY)
7480                      lngTmp18 = arr_varCliDir(C_ID, lngY)
7490                      arr_varCliDir(C_PARID, lngX) = lngTmp18
7500                      Exit For
7510                    End If
7520                  Next
7530                  If blnFound = False Then
7540                    Stop
7550                  End If
7560                Else
                      ' ** 2nd or deeper subfolder.
7570                  strTmp02 = Mid(strTmp01, (intPos01 + 1))
7580                  strTmp01 = Left(strTmp01, (intPos01 - 1))
                      'lngTmp22, lngTmp23, lngTmp24
                      ' ** Find 1st level clidir_id.
7590                  blnFound = False
7600                  For lngY = 0& To (lngCliDirs - 1&)
7610                    If arr_varCliDir(C_ID, lngY) > 0& And arr_varCliDir(C_PARID, lngY) = 0& And _
                            arr_varCliDir(C_NAM, lngY) = strTmp01 Then
7620                      blnFound = True
7630                      lngTmp18 = arr_varCliDir(C_ID, lngY)
7640                      Exit For
7650                    End If
7660                  Next
7670                  If blnFound = False Then
7680                    Stop
7690                  End If
7700                  intPos01 = InStr(strTmp02, ";")
7710                  If intPos01 = 0 Then
                        ' ** 2nd level subfolder.
7720                    blnFound = False
7730                    For lngY = 0& To (lngCliDirs - 1&)
7740                      If arr_varCliDir(C_ID, lngY) > 0& And arr_varCliDir(C_PARID, lngY) = lngTmp18 And _
                              arr_varCliDir(C_NAM, lngY) = strTmp02 Then
7750                        blnFound = True
7760                        ![clidir_parid] = arr_varCliDir(C_ID, lngY)
7770                        lngTmp19 = arr_varCliDir(C_ID, lngY)
7780                        arr_varCliDir(C_PARID, lngX) = lngTmp19
7790                        Exit For
7800                      End If
7810                    Next
7820                    If blnFound = False Then
7830                      Stop
7840                    End If
7850                  Else
                        ' ** 3rd or deeper subfolder.
7860                    strTmp03 = Mid(strTmp02, (intPos01 + 1))
7870                    strTmp02 = Left(strTmp02, (intPos01 - 1))
                        ' ** Find 2nd level clidir_id.
7880                    blnFound = False
7890                    For lngY = 0& To (lngCliDirs - 1&)
7900                      If arr_varCliDir(C_ID, lngY) > 0& And arr_varCliDir(C_PARID, lngY) = lngTmp18 And _
                              arr_varCliDir(C_NAM, lngY) = strTmp02 Then
7910                        blnFound = True
7920                        lngTmp19 = arr_varCliDir(C_ID, lngY)
7930                        Exit For
7940                      End If
7950                    Next
7960                    If blnFound = False Then
7970                      Stop
7980                    End If
7990                    intPos01 = InStr(strTmp03, ";")
8000                    If intPos01 = 0 Then
                          ' ** 3rd level subfolder.
8010                      blnFound = False
8020                      For lngY = 0& To (lngCliDirs - 1&)
8030                        If arr_varCliDir(C_ID, lngY) > 0& And arr_varCliDir(C_PARID, lngY) = lngTmp19 And _
                                arr_varCliDir(C_NAM, lngY) = strTmp03 Then
8040                          blnFound = True
8050                          ![clidir_parid] = arr_varCliDir(C_ID, lngY)
8060                          lngTmp20 = arr_varCliDir(C_ID, lngY)
8070                          arr_varCliDir(C_PARID, lngX) = lngTmp20
8080                          Exit For
8090                        End If
8100                      Next
8110                      If blnFound = False Then
8120                        Stop
8130                      End If
8140                    Else
                          ' ** 4th or deeper subfolder.
8150                      strTmp04 = Mid(strTmp03, (intPos01 + 1))
8160                      strTmp03 = Left(strTmp03, (intPos01 - 1))
                          ' ** Find 3rd level clidir_id.
8170                      blnFound = False
8180                      For lngY = 0& To (lngCliDirs - 1&)
8190                        If arr_varCliDir(C_ID, lngY) > 0& And arr_varCliDir(C_PARID, lngY) = lngTmp19 And _
                                arr_varCliDir(C_NAM, lngY) = strTmp03 Then
8200                          blnFound = True
8210                          lngTmp20 = arr_varCliDir(C_ID, lngY)
8220                          Exit For
8230                        End If
8240                      Next
8250                      If blnFound = False Then
8260                        Stop
8270                      End If
8280                      intPos01 = InStr(strTmp04, ";")
8290                      If intPos01 = 0 Then
                            ' ** 4th level subfolder.
8300                        blnFound = False
8310                        For lngY = 0& To (lngCliDirs - 1&)
8320                          If arr_varCliDir(C_ID, lngY) > 0& And arr_varCliDir(C_PARID, lngY) = lngTmp20 And _
                                  arr_varCliDir(C_NAM, lngY) = strTmp04 Then
8330                            blnFound = True
8340                            ![clidir_parid] = arr_varCliDir(C_ID, lngY)
8350                            lngTmp21 = arr_varCliDir(C_ID, lngY)
8360                            arr_varCliDir(C_PARID, lngX) = lngTmp21
8370                            Exit For
8380                          End If
8390                        Next
8400                        If blnFound = False Then
8410                          Stop
8420                        End If
8430                      Else
                            ' ** 5th or deeper subfolder.
8440                        strTmp05 = Mid(strTmp04, (intPos01 + 1))
8450                        strTmp04 = Left(strTmp04, (intPos01 - 1))
                            ' ** Find 4th level clidir_id.
8460                        blnFound = False
8470                        For lngY = 0& To (lngCliDirs - 1&)
8480                          If arr_varCliDir(C_ID, lngY) > 0& And arr_varCliDir(C_PARID, lngY) = lngTmp20 And _
                                  arr_varCliDir(C_NAM, lngY) = strTmp04 Then
8490                            blnFound = True
8500                            lngTmp21 = arr_varCliDir(C_ID, lngY)
8510                            Exit For
8520                          End If
8530                        Next
8540                        If blnFound = False Then
8550                          Stop
8560                        End If
8570                        intPos01 = InStr(strTmp05, ";")
8580                        If intPos01 = 0 Then
                              ' ** 5th level subfolder.
8590                          blnFound = False
8600                          For lngY = 0& To (lngCliDirs - 1&)
8610                            If arr_varCliDir(C_ID, lngY) > 0& And arr_varCliDir(C_PARID, lngY) = lngTmp21 And _
                                    arr_varCliDir(C_NAM, lngY) = strTmp05 Then
8620                              blnFound = True
8630                              ![clidir_parid] = arr_varCliDir(C_ID, lngY)
8640                              lngTmp22 = arr_varCliDir(C_ID, lngY)
8650                              arr_varCliDir(C_PARID, lngX) = lngTmp22
8660                              Exit For
8670                            End If
8680                          Next
8690                          If blnFound = False Then
8700                            Stop
8710                          End If
8720                        Else
                              ' ** 6th or deeper subfolder.
8730                          Stop

8740                        End If
8750                      End If
8760                    End If
8770                  End If
8780                End If
8790              End If
8800              ![clidir_path] = arr_varCliDir(C_PATH, lngX)
8810              ![clidir_filecnt] = arr_varCliDir(C_FILES, lngX)
8820              ![clidir_subcnt] = arr_varCliDir(C_SUBS, lngX)
8830              ![clidir_datemodified] = Now()
8840              .Update
8850              .Bookmark = .LastModified
8860              lngCliDirID = ![clidir_id]
8870            End If  ' ** blnAddThis.
8880            arr_varCliDir(C_ID, lngX) = lngCliDirID
8890          Next  ' ** For each client directory: lngX
8900          .Close
8910        End With  ' ** zz_tbl_Client_Directory: rst1.

8920        blnAddAll = False: blnAddThis = False: blnEditThis = False
8930        Set rst1 = .OpenRecordset("zz_tbl_Client_File", dbOpenDynaset, dbConsistent)
8940        With rst1
8950          If .BOF = True And .EOF = True Then
                ' ** Nothing there yet.
8960            blnAddAll = True
8970          End If
8980          lngCliDirID = 0&: lngCliFileID = 0&: lngCliTblID = 0&
8990          For lngX = 0& To (lngCliDirs - 1&)
9000            lngCliDirID = arr_varCliDir(C_ID, lngX)
9010            If arr_varCliDir(C_FILES, lngX) > 0& Then
9020              arr_varTmpCliFile = arr_varCliDir(C_F_ARR, lngX)
9030              lngCliFiles = UBound(arr_varTmpCliFile, 2) + 1&
9040              For lngY = 0& To (lngCliFiles - 1&)
9050                blnAddThis = False: blnEditThis = False
9060                If blnAddAll = False Then
9070                  .FindFirst "[clidir_id] = " & CStr(lngCliDirID) & " And " & _
                        "[clifile_name] = " & Chr(34) & arr_varTmpCliFile(F_NAM, lngY) & Chr(34)
                      ' ** There are file names with an apostrophe: '
9080                  If .NoMatch = True Then
9090                    blnAddThis = True
9100                  Else
9110                    blnEditThis = True
9120                    lngCliFileID = ![clifile_id]
9130                  End If
9140                Else
9150                  blnAddThis = True
9160                End If
9170                If blnAddThis = True Or blnEditThis = True Then
9180                  If blnAddThis = True Then
9190                    .AddNew
9200                  Else
9210                    .Edit
9220                  End If
                      ' *************************************************
                      ' ** Array: arr_varCliFile()/arr_varTmpCliFile()
                      ' **
                      ' **   Element  Name                  Constant
                      ' **   =======  ====================  ===========
                      ' **      0     clidir_id             F_DIRID
                      ' **      1     clifile_id            F_ID
                      ' **      2     clifile_name          F_NAM
                      ' **      3     clifile_client        F_CLIENT
                      ' **      4     clifile_ismdb         F_ISMDB
                      ' **      5     clifile_canopen       F_CANOPEN
                      ' **      6     clifile_filesize      F_FILSIZ
                      ' **      7     clifile_tblcnt        F_TBLS
                      ' **      8     Table Array           F_T_ARR
                      ' **      9     clifile_accver        F_ACC_VER
                      ' **     10     clifile_m_v_ver       F_M_VER
                      ' **     11     clifile_appversion    F_APPVER
                      ' **     12     clifile_appdate       F_APPDATE
                      ' **     13     clifile_note          F_NOTE
                      ' **
                      ' *************************************************
9230                  ![clidir_id] = lngCliDirID
9240                  ![clifile_name] = arr_varTmpCliFile(F_NAM, lngY)
9250                  ![clifile_client] = arr_varTmpCliFile(F_CLIENT, lngY)
9260                  ![clifile_ismdb] = arr_varTmpCliFile(F_ISMDB, lngY)
9270                  ![clifile_canopen] = arr_varTmpCliFile(F_CANOPEN, lngY)
9280                  ![clifile_filesize] = arr_varTmpCliFile(F_FILSIZ, lngY)
9290                  ![clifile_tblcnt] = arr_varTmpCliFile(F_TBLS, lngY)
9300                  If arr_varTmpCliFile(F_ACC_VER, lngY) <> vbNullString Then
9310                    ![clifile_accver] = arr_varTmpCliFile(F_ACC_VER, lngY)
9320                  End If
9330                  If arr_varTmpCliFile(F_M_VER, lngY) <> vbNullString Then
9340                    ![clifile_m_v_ver] = arr_varTmpCliFile(F_M_VER, lngY)
9350                  End If
9360                  If arr_varTmpCliFile(F_APPVER, lngY) <> vbNullString Then
9370                    ![clifile_appversion] = arr_varTmpCliFile(F_APPVER, lngY)
9380                  End If
9390                  ![clifile_appdate] = arr_varTmpCliFile(F_APPDATE, lngY)
9400                  If arr_varTmpCliFile(F_NOTE, lngY) <> vbNullString Then
9410                    ![clifile_note] = arr_varTmpCliFile(F_NOTE, lngY)
9420                  End If
9430                  ![clifile_datemodified] = Now()
9440                  .Update
9450                  .Bookmark = .LastModified
9460                  lngCliFileID = ![clifile_id]
9470                End If  ' ** blnAddThis.
9480                arr_varTmpCliFile(F_DIRID, lngY) = lngCliDirID
9490                arr_varTmpCliFile(F_ID, lngY) = lngCliFileID
9500              Next  ' ** For each client file: lngY
9510              arr_varCliDir(C_F_ARR, lngX) = arr_varTmpCliFile
9520            End If  ' ** Has files.
9530          Next  ' ** For each client directory: lngX.
9540          .Close
9550        End With

9560      End With  ' ** dbsLoc.

9570      If blnDocThroughFilesOnly = False Then

9580      End If

9590    End With  ' ** fso.

9600    dbsLoc.Close
9610    wrkLoc.Close

9620    Beep

EXITP:
9630    Set fsfl = Nothing
9640    Set fsfls = Nothing
9650    Set fsfd7 = Nothing
9660    Set fsfd6 = Nothing
9670    Set fsfd5 = Nothing
9680    Set fsfd4 = Nothing
9690    Set fsfd3 = Nothing
9700    Set fsfd2 = Nothing
9710    Set fsfd1 = Nothing
9720    Set fsfds7 = Nothing
9730    Set fsfds6 = Nothing
9740    Set fsfds5 = Nothing
9750    Set fsfds4 = Nothing
9760    Set fsfds3 = Nothing
9770    Set fsfds2 = Nothing
9780    Set fsfds1 = Nothing
9790    Set fso = Nothing
9800    Set prp = Nothing
9810    Set doc = Nothing
9820    Set fld = Nothing
9830    Set tdf = Nothing
9840    Set rst1 = Nothing
9850    Set dbsLnk = Nothing
9860    Set wrkLnk = Nothing
9870    Version_Client_Doc = blnRetVal
9880    Exit Function

ERRH:
9890    blnRetVal = False
9900    Select Case ERR.Number
        Case Else
9910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9920    End Select
9930    Resume EXITP

      #End If

End Function

Public Function Version_NF_Ex_Doc() As Boolean
' ** Document version differences using tables absent and extra.
' ** Not called.

      #If IsDev Then

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_NF_Ex_Doc"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngDTAs As Long, arr_varDTA As Variant
        Dim lngPrevs As Long, arr_varPrev As Variant
        Dim lngPrvTbls As Long, arr_varPrvTbl As Variant
        Dim lngNotFounds As Long, arr_varNotFound() As Variant
        Dim lngExtras As Long, arr_varExtra() As Variant
        Dim blnFound As Boolean, blnAddAll As Boolean, blnAddThis As Boolean
        Dim lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: lngDtas().
        'Const D_DID  As Integer = 0
        'Const D_DNAM As Integer = 1
        'Const D_TID  As Integer = 2
        Const D_TNAM As Integer = 3

        ' ** Array: arr_varPrev().
        Const P_PID  As Integer = 0
        Const P_PNAM As Integer = 1
        Const P_FID  As Integer = 2
        'Const P_FNAM As Integer = 3
        'Const P_TCNT As Integer = 4

        ' ** Array: arr_varPrvTbl().
        Const T_DID  As Integer = 0
        'Const T_DNAM As Integer = 1
        Const T_FID  As Integer = 2
        'Const T_FNAM As Integer = 3
        'Const T_TID  As Integer = 4
        Const T_TNAM As Integer = 5

        ' ** Array: arr_varNotFound().
        Const N_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const N_DID  As Integer = 0
        Const N_FID  As Integer = 1
        Const N_TNAM As Integer = 2

        ' ** Array: arr_varExtra().
        Const E_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const E_DID  As Integer = 0
        Const E_FID  As Integer = 1
        Const E_TNAM As Integer = 2

10010   blnRetVal = True

10020   Set dbs = CurrentDb
10030   With dbs

10040     For lngW = 1& To 2&

            ' ** tblDatabase_Table, tables currently in TrustDta.mdb/TrstArch.mdb; 47/3.
10050       Set qdf = .QueryDefs("qryVersion_80" & Chr(96& + lngW))  ' ** 80a, 80b.
10060       Set rst = qdf.OpenRecordset
10070       With rst
10080         .MoveLast
10090         lngDTAs = .RecordCount
10100         .MoveFirst
10110         arr_varDTA = .GetRows(lngDTAs)
              ' ***********************************************
              ' ** Array: arr_varDta()
              ' **
              ' **   Field  Element  Name        Constant
              ' **   =====  =======  ==========  ============
              ' **     1       0     dbs_id      D_DID
              ' **     2       1     dbs_name    D_DNAM
              ' **     3       2     tbl_id      D_TID
              ' **     4       3     tbl_name    D_TNAM
              ' **
              ' ***********************************************
10120         .Close
10130       End With

            ' ** .._81a, grouped by verdir_name, with cnt of tables in TrustDta.mdb/TrstArch.mdb; 16/16.
10140       Set qdf = .QueryDefs("qryVersion_82" & Chr(96& + lngW))  ' ** 82a, 82b.
10150       Set rst = qdf.OpenRecordset
10160       With rst
10170         .MoveLast
10180         lngPrevs = .RecordCount
10190         .MoveFirst
10200         arr_varPrev = .GetRows(lngPrevs)
              ' ***************************************************
              ' ** Array: arr_varPrev()
              ' **
              ' **   Field  Element  Name            Constant
              ' **   =====  =======  ==============  ============
              ' **     1       0     verdir_id       P_PID
              ' **     2       1     verdir_name     P_PNAM
              ' **     3       2     verfile_id      P_FID
              ' **     4       3     verfile_name    P_FNAM
              ' **     5       4     tblcnt          P_TCNT
              ' **
              ' ***************************************************
10210         .Close
10220       End With

10230       lngNotFounds = 0&
10240       ReDim arr_varNotFound(N_ELEMS, 0)

10250       lngExtras = 0&
10260       ReDim arr_varExtra(E_ELEMS, 0)

            ' ** Cross-reference each previous database with the current list of tables.
10270       For lngX = 0& To (lngPrevs - 1&)

              ' ** qryVersion_81a (tables found in previous TrustDta.mdb/TrstArch.mdb), by specified [verdir].
10280         Set qdf = .QueryDefs("qryVersion_83" & Chr(96& + lngW))  ' ** 83a, 83b.
10290         With qdf.Parameters
10300           ![verdir] = arr_varPrev(P_PNAM, lngX)
10310         End With
10320         Set rst = qdf.OpenRecordset
10330         With rst
10340           .MoveLast
10350           lngPrvTbls = .RecordCount
10360           .MoveFirst
10370           arr_varPrvTbl = .GetRows(lngPrvTbls)
                ' ****************************************************
                ' ** Array: arr_varPrvTbl()
                ' **
                ' **   Field  Element  Name            Constant
                ' **   =====  =======  ==============  =============
                ' **     1       0     verdir_id       T_DID
                ' **     2       1     verdir_name     T_DNAM
                ' **     3       2     verfile_id      T_FID
                ' **     4       3     verfile_name    T_FNAM
                ' **     5       4     vertbl_id       T_TID
                ' **     6       5     vertbl_name     T_TNAM
                ' **
                ' ****************************************************
10380           .Close
10390         End With

              ' ** Check for tables not found.
10400         For lngY = 0& To (lngDTAs - 1&)
10410           blnFound = False
10420           For lngZ = 0& To (lngPrvTbls - 1&)
10430             If arr_varPrvTbl(T_TNAM, lngZ) = arr_varDTA(D_TNAM, lngY) Then
10440               blnFound = True
10450               Exit For
10460             End If
10470           Next
10480           If blnFound = False Then
10490             lngNotFounds = lngNotFounds + 1&
10500             lngE = lngNotFounds - 1&
10510             ReDim Preserve arr_varNotFound(N_ELEMS, lngE)
                  ' ************************************************
                  ' ** Array: arr_varNotFound()
                  ' **
                  ' **   Field  Element  Name          Constant
                  ' **   =====  =======  ============  ===========
                  ' **     1       0     verdir_id     N_DID
                  ' **     2       1     verfile_id    N_FID
                  ' **     3       2     tbl_name      N_TNAM
                  ' **
                  ' ************************************************
10520             arr_varNotFound(N_DID, lngE) = arr_varPrev(P_PID, lngX)
10530             arr_varNotFound(N_FID, lngE) = arr_varPrev(P_FID, lngX)
10540             arr_varNotFound(N_TNAM, lngE) = arr_varDTA(D_TNAM, lngY)
10550           End If
10560         Next  ' ** lngY.

              ' ** Check for extra tables.
10570         For lngY = 0& To (lngPrvTbls - 1&)
10580           blnFound = False
10590           For lngZ = 0& To (lngDTAs - 1&)
10600             If arr_varDTA(D_TNAM, lngZ) = arr_varPrvTbl(T_TNAM, lngY) Then
10610               blnFound = True
10620               Exit For
10630             End If
10640           Next
10650           If blnFound = False Then
10660             lngExtras = lngExtras + 1&
10670             lngE = lngExtras - 1&
10680             ReDim Preserve arr_varExtra(E_ELEMS, lngE)
                  ' *************************************************
                  ' ** Array: arr_varExtra()
                  ' **
                  ' **   Field  Element  Name           Constant
                  ' **   =====  =======  =============  ===========
                  ' **     1       0     verdir_id      E_DID
                  ' **     2       1     verfile_id     E_FID
                  ' **     3       2     vertbl_name    E_TNAM
                  ' **
                  ' *************************************************
10690             arr_varExtra(E_DID, lngE) = arr_varPrvTbl(T_DID, lngX)
10700             arr_varExtra(E_FID, lngE) = arr_varPrvTbl(T_FID, lngX)
10710             arr_varExtra(E_TNAM, lngE) = arr_varPrvTbl(T_TNAM, lngY)
10720           End If
10730         Next  ' ** lngY.

10740       Next  ' ** lngx

            ' ** Save those not found.
10750       Set rst = .OpenRecordset("tblVersion_NotFound", dbOpenDynaset, dbConsistent)
10760       With rst
10770         blnAddAll = False
10780         If .BOF = True And .EOF = True Then
10790           blnAddAll = True
10800         End If
10810         For lngX = 0& To (lngNotFounds - 1&)
10820           blnAddThis = False
10830           If blnAddAll = False Then
10840             .FindFirst "[verdir_id] = " & CStr(arr_varNotFound(N_DID, lngX)) & " And " & _
                    "[verfile_id] = " & CStr(arr_varNotFound(N_FID, lngX)) & " And " & _
                    "[tbl_name] = '" & arr_varNotFound(N_TNAM, lngX) & "'"
10850             If .NoMatch = True Then
10860               blnAddThis = True
10870             End If
10880           Else
10890             blnAddThis = True
10900           End If
10910           If blnAddThis = True Then
10920             .AddNew
10930             ![verdir_id] = arr_varNotFound(N_DID, lngX)
10940             ![verfile_id] = arr_varNotFound(N_FID, lngX)
10950             ![tbl_name] = arr_varNotFound(N_TNAM, lngX)
10960             ![vernf_datemodified] = Now()
10970             .Update
10980           End If

10990         Next
11000         .Close
11010       End With

            ' ** Save extras.
11020       Set rst = .OpenRecordset("tblVersion_Extra", dbOpenDynaset, dbConsistent)
11030       With rst
11040         blnAddAll = False
11050         If .BOF = True And .EOF = True Then
11060           blnAddAll = True
11070         End If
11080         For lngX = 0& To (lngExtras - 1&)
11090           blnAddThis = False
11100           If blnAddAll = False Then
11110             .FindFirst "[verdir_id] = " & CStr(arr_varExtra(E_DID, lngX)) & " And " & _
                    "[verfile_id] = " & CStr(arr_varExtra(E_FID, lngX)) & " And " & _
                    "[vertbl_name] = '" & arr_varExtra(E_TNAM, lngX) & "'"
11120             If .NoMatch = True Then
11130               blnAddThis = True
11140             End If
11150           Else
11160             blnAddThis = True
11170           End If
11180           If blnAddThis = True Then
11190             .AddNew
11200             ![verdir_id] = arr_varExtra(E_DID, lngX)
11210             ![verfile_id] = arr_varExtra(E_FID, lngX)
11220             ![vertbl_name] = arr_varExtra(E_TNAM, lngX)
11230             ![verex_datemodified] = Now()
11240             .Update
11250           End If

11260         Next
11270         .Close
11280       End With

11290     Next  ' ** lngW

11300     .Close
11310   End With  ' ** dbs.

11320   Beep

EXITP:
11330   Set rst = Nothing
11340   Set qdf = Nothing
11350   Set dbs = Nothing
11360   Version_NF_Ex_Doc = blnRetVal
11370   Exit Function

ERRH:
11380   blnRetVal = False
11390   Select Case ERR.Number
        Case Else
11400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11410   End Select
11420   Resume EXITP

      #End If

End Function

Public Function Version_Doc() As Boolean
' ** Not called.

      #If IsDev Then

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Doc"

        Dim wrkLnk As DAO.Workspace, dbsLnk As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field, doc As DAO.Document, prp As DAO.Property
        Dim wrkLoc As DAO.Workspace, dbsLoc As DAO.Database
        Dim rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset, rst4 As DAO.Recordset
        Dim fso As Scripting.FileSystemObject, fsfds As Scripting.Folders, fsfls As Scripting.Files
        Dim fsfd1 As Scripting.Folder, fsfd2 As Scripting.Folder, fsfl As Scripting.File
        Dim intWrkType As Integer
        Dim strDocDatabase As String
        Dim lngVerDirs As Long, arr_varVerDir() As Variant
        Dim lngVerFiles As Long, arr_varVerFile() As Variant
        Dim lngVerTbls As Long, arr_varVerTbl() As Variant
        Dim lngVerFlds As Long, arr_varVerFld() As Variant
        Dim lngDirID As Long, lngFileID As Long, lngTblID As Long
        Dim lngDirsAdd As Long, lngFilesAdd As Long, lngTblsAdd As Long, lngFldsAdd As Long
        Dim lngAccVerUpdate As Long, lngTblCntUpdate As Long, lngNewVerUpdate As Long, lngNoteUpdate As Long
        Dim strVer As String
        Dim blnAdd As Boolean
        Dim arr_varTmp01 As Variant, arr_varTmp02 As Variant, arr_varTmp03 As Variant, strTmp04 As String, lngTmp05 As Long
        Dim lngV As Long, lngW As Long, lngX As Long, lngY As Long, lngZ As Long
        Dim blnRetVal As Boolean

        Const strBasePath As String = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\PreviousVersionDBs"  '## OK

        ' ** Array: arr_varVerDir().
        Const D_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const D_DIR   As Integer = 0
        Const D_PATH  As Integer = 1
        Const D_FILES As Integer = 2
        Const D_F_ARR As Integer = 3
        ' ******************************************
        ' ** Array: arr_varVerDir()
        ' **
        ' **   Element  Name          Constant
        ' **   =======  ============  ============
        ' **      0     Name          D_DIR
        ' **      1     Path          D_PATH
        ' **      2     Files         D_FILES
        ' **      3     File Array    D_F_ARR
        ' **
        ' ******************************************

        ' ** Array: arr_varVerFile().
        Const F_ELEMS As Integer = 9  ' ** Array's first-element UBound().
        Const F_FNAM    As Integer = 0
        Const F_TA_VER  As Integer = 1
        Const F_ACC_VER As Integer = 2
        Const F_OPEN    As Integer = 3
        Const F_TBLS    As Integer = 4
        Const F_T_ARR   As Integer = 5
        Const F_NOTE    As Integer = 6
        Const F_M_VER   As Integer = 7
        Const F_APPVER  As Integer = 8
        Const F_APPDATE As Integer = 9
        ' *******************************************************
        ' ** Array: arr_varVerFile()
        ' **
        ' **   Element  Name                        Constant
        ' **   =======  ==========================  ===========
        ' **      0     Name                        F_FNAM
        ' **      1     Trust Accountant Version    F_TA_VER
        ' **      2     Access Version              F_ACC_VER
        ' **      3     Can Open                    F_OPEN
        ' **      4     Tables                      F_TBLS
        ' **      5     Table Array                 F_T_ARR
        ' **      6     Note                        F_NOTE
        ' **      7     m_Vx Version                F_M_VER
        ' **      8     AppVersion                  F_APPVER
        ' **      9     AppDate                     F_APPDATE
        ' **
        ' *******************************************************

        ' ** Array: arr_varVerTbl().
        Const T_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const T_TNAM  As Integer = 0
        Const T_FLDS  As Integer = 1
        Const T_F_ARR As Integer = 2
        ' ******************************************
        ' ** Array: arr_varVerTbl()
        ' **
        ' **   Element  Name           Constant
        ' **   =======  =============  ===========
        ' **      0     Name           T_TNAM
        ' **      1     Fields         T_FLDS
        ' **      2     Field Array    T_F_ARR
        ' **
        ' ******************************************

        ' ** Array: arr_varVerFld().
        Const FD_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const FD_FNAM As Integer = 0
        Const FD_TYP  As Integer = 1
        Const FD_SIZ  As Integer = 2
        ' **********************************
        ' ** Array: arr_varVerFld()
        ' **
        ' **   Element  Name    Constant
        ' **   =======  ======  ==========
        ' **      0     Name    FD_FNAM
        ' **      1     Type    FD_TYP
        ' **      2     Size    FD_SIZ
        ' **
        ' **********************************

11510   blnRetVal = True

11520   lngDirsAdd = 0&: lngFilesAdd = 0&: lngTblsAdd = 0&: lngFldsAdd = 0&
11530   lngAccVerUpdate = 0&: lngTblCntUpdate = 0&: lngNewVerUpdate = 0&: lngNoteUpdate = 0&

11540   Set wrkLoc = DBEngine.Workspaces(0)
11550   Set dbsLoc = wrkLoc.Databases(0)

        ' ** It appears that accessing earlier versions of Access databases in this way
        ' ** is not a problem. When I attempted to open the 97's on this list directly,
        ' ** a message indicated they were from an earlier version and could only be
        ' ** viewed; that it would have to be converted in order to do anything with it.
        ' ** No such problem shows up here.

11560   For lngV = 1& To 2&  '21&

11570     Select Case lngV
          Case 1&
11580       strDocDatabase = gstrFile_DataName
11590     Case 2&
11600       strDocDatabase = gstrFile_ArchDataName
            'Case 3&
            '  strDocDatabase = gstrFile_App & "." & gstrExt_AppDev
            'Case 4&
            '  strDocDatabase = gstrFile_App & "." & gstrExt_AppRun
            'Case 5&
            '  strDocDatabase = "TRUSTDTA_empty.MDB"
            'Case 6&
            '  strDocDatabase = "TRUSTDTA_empty97.MDB"
            'Case 7&
            '  strDocDatabase = "TRSTARCH_empty.MDB"
            'Case 8&
            '  strDocDatabase = "TRSTARCH_empty97.MDB"
            'Case 9&
            '  strDocDatabase = "TrstAU.mdb"
            'Case 10&
            '  strDocDatabase = "TrustDU.mdb"
            'Case 11&
            '  strDocDatabase = "TrustUpd.mde"
            'Case 12&
            '  strDocDatabase = "TrustDta.BAK"
            'Case 13&
            '  strDocDatabase = "TrstArch.BAK"
            'Case 14&
            '  strDocDatabase = "Master.mdw"
            'Case 15&
            '  strDocDatabase = "TrustSec.mdw"
            'Case 16&
            '  strDocDatabase = "TA-CourtRptData.mdb"
            'Case 17&
            '  strDocDatabase = "TA-CourtRptArchive.mdb"
            'Case 18&
            '  strDocDatabase = "TA-CourtRpt.mdb"
            'Case 19&
            '  strDocDatabase = "TrustDta_bak.mdb"
            'Case 20&
            '  strDocDatabase = "TrstArch_bak.mdb"
            'Case 21&
            '  strDocDatabase = "TrustDta_20070214.mdb"
11610     End Select

          ' ** Is TrustSec.mdw (gstrFile_SecurityName) vis-à-vis Master.mdw handled entirely during installation?
          ' ** What about Permissions and Owners within their MDBs?
          ' ** I'm thinking it's better to wholly create new MDBs!

11620     lngVerDirs = 0&
11630     ReDim arr_varVerDir(D_ELEMS, 0)

11640     Set fso = CreateObject("Scripting.FileSystemObject")
11650     With fso

11660       Set fsfd1 = .GetFolder(strBasePath)
11670       Set fsfds = fsfd1.SubFolders

11680       lngVerDirs = 0&
11690       ReDim arr_varVerDir(D_ELEMS, 0)

11700       For Each fsfd2 In fsfds
11710         With fsfd2
11720           If Left(.Name, 4) = "Ver_" Then  'Or .Name = "CourtReports" Then
11730             lngVerDirs = lngVerDirs + 1&
11740             lngW = lngVerDirs - 1&
11750             ReDim Preserve arr_varVerDir(D_ELEMS, lngW)
11760             arr_varVerDir(D_DIR, lngW) = .Name
11770             arr_varVerDir(D_PATH, lngW) = .Path
11780             arr_varVerDir(D_FILES, lngW) = CLng(0)
11790             arr_varVerDir(D_F_ARR, lngW) = Empty
11800           End If
11810         End With  ' ** fsfd2.

11820       Next

11830       For lngW = 0& To (lngVerDirs - 1&)

11840         strVer = arr_varVerDir(D_DIR, lngW)
              'If strVer = "CourtReports" Then
              '  strVer = "2.0.00"
              'Else
11850         strVer = Mid(strVer, 5, 1) & "." & Mid(strVer, 7, 1) & "." & Mid(strVer, 9)  'Ver_2_1_45
              'End If

11860         If Right(strVer, 1) = "0" Then strVer = Left(strVer, (Len(strVer) - 1))

11870         Set fsfd2 = .GetFolder(arr_varVerDir(D_PATH, lngW))
11880         Set fsfls = fsfd2.Files

11890         lngVerFiles = fsfls.Count
11900         arr_varVerDir(D_FILES, lngW) = lngVerFiles
11910         ReDim arr_varVerFile(F_ELEMS, (lngVerFiles - 1&))

11920         lngX = -1&
11930         For Each fsfl In fsfls
11940           With fsfl
11950             Select Case .Name
                  Case "License.txt", "Trust.ico", "TA.lic"
                    ' ** Skip
11960               lngVerFiles = lngVerFiles - 1&
11970               arr_varVerDir(D_FILES, lngW) = lngVerFiles
11980             Case Else
11990               lngX = lngX + 1&
12000               ReDim Preserve arr_varVerFile(F_ELEMS, lngX)
12010               arr_varVerFile(F_FNAM, lngX) = .Name
12020               arr_varVerFile(F_TA_VER, lngX) = strVer
12030               arr_varVerFile(F_ACC_VER, lngX) = vbNullString
12040               arr_varVerFile(F_OPEN, lngX) = CBool(False)
12050               arr_varVerFile(F_TBLS, lngX) = CLng(0)
12060               arr_varVerFile(F_T_ARR, lngX) = Empty
12070               arr_varVerFile(F_NOTE, lngX) = vbNullString
12080               arr_varVerFile(F_M_VER, lngX) = vbNullString
12090               arr_varVerFile(F_APPVER, lngX) = vbNullString
12100               arr_varVerFile(F_APPDATE, lngX) = Null
12110             End Select
12120           End With  ' ** fsfl.
12130         Next

12140         arr_varVerDir(D_F_ARR, lngW) = arr_varVerFile

12150       Next

12160     End With  ' ** fso.

12170     Set fsfl = Nothing
12180     Set fsfd2 = Nothing
12190     Set fsfd1 = Nothing
12200     Set fsfls = Nothing
12210     Set fsfds = Nothing
12220     Set fso = Nothing

12230     intWrkType = 0
12240 On Error Resume Next
12250     Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
12260     If ERR.Number <> 0 Then
12270 On Error GoTo ERRH
12280 On Error Resume Next
12290       Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
12300       If ERR.Number <> 0 Then
12310 On Error GoTo ERRH
12320 On Error Resume Next
12330         Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
12340         If ERR.Number <> 0 Then
12350 On Error GoTo ERRH
12360 On Error Resume Next
12370           Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
12380           If ERR.Number <> 0 Then
12390 On Error GoTo ERRH
12400 On Error Resume Next
12410             Set wrkLnk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
12420             If ERR.Number <> 0 Then
12430 On Error GoTo ERRH
12440 On Error Resume Next
12450               Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
12460               If ERR.Number <> 0 Then
12470 On Error GoTo ERRH
12480 On Error Resume Next
12490                 Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
12500 On Error GoTo ERRH
12510                 intWrkType = 7
12520               Else
12530 On Error GoTo ERRH
12540                 intWrkType = 6
12550               End If
12560             Else
12570 On Error GoTo ERRH
12580               intWrkType = 5
12590             End If
12600           Else
12610 On Error GoTo ERRH
12620             intWrkType = 4
12630           End If
12640         Else
12650 On Error GoTo ERRH
12660           intWrkType = 3
12670         End If
12680       Else
12690 On Error GoTo ERRH
12700         intWrkType = 2
12710       End If
12720     Else
12730 On Error GoTo ERRH
12740       intWrkType = 1
12750     End If

12760     With wrkLnk

12770       For lngW = 0& To (lngVerDirs - 1&)
12780         If arr_varVerDir(D_FILES, lngW) > 0& Then

12790           lngVerFiles = arr_varVerDir(D_FILES, lngW)
12800           arr_varTmp01 = arr_varVerDir(D_F_ARR, lngW)  ' ** arr_varVerFile().

12810           For lngX = 0& To (lngVerFiles - 1&)

                  ' ** Only open specified file.
12820             If arr_varTmp01(F_FNAM, lngX) = strDocDatabase Then

12830 On Error Resume Next
12840               Set dbsLnk = .OpenDatabase(arr_varVerDir(D_PATH, lngW) & LNK_SEP & arr_varTmp01(F_FNAM, lngX), False, True)  ' ** {pathfile}, {exclusive}, {read-only}
12850               If ERR.Number = 0 Then
12860 On Error GoTo 0
12870                 arr_varTmp01(F_OPEN, lngX) = CBool(True)
12880                 With dbsLnk

12890                   If Right(arr_varTmp01(F_FNAM, lngX), 3) <> "mdw" Then
12900                     arr_varTmp01(F_ACC_VER, lngX) = .Containers("Databases").Documents("MSysDb").Properties("AccessVersion")
                          ' ** CurrentDb.Containers("Databases").Documents("MSysDb").Properties("AccessVersion") = 08.50
12910                   Else
12920 On Error Resume Next
12930                     arr_varTmp01(F_ACC_VER, lngX) = .Containers("Databases").Documents("MSysDb").Properties("AccessVersion")
12940                     If ERR.Number <> 0 Then
12950 On Error GoTo 0
12960                       arr_varTmp01(F_ACC_VER, lngX) = "Jet " & .Properties("Version")
12970                     Else
12980 On Error GoTo 0
12990                     End If

13000                   End If

13010                   Select Case arr_varTmp01(F_ACC_VER, lngX)
                        Case "02.00"
13020                     arr_varTmp01(F_ACC_VER, lngX) = "Access 2.0"
13030                   Case "06.68"
13040                     arr_varTmp01(F_ACC_VER, lngX) = "Access 95"
13050                   Case "07.53"
13060                     arr_varTmp01(F_ACC_VER, lngX) = "Access 97"
13070                   Case "08.50"
13080                     arr_varTmp01(F_ACC_VER, lngX) = "Access 2000"
13090                   Case "09.50"
13100                     arr_varTmp01(F_ACC_VER, lngX) = "Access 2002/2003"
13110                   Case Else
                          ' ** The Jet MDW.
13120                   End Select

13130                   For Each doc In .Containers("Databases").Documents
13140                     With doc
13150                       If .Name = "UserDefined" Then
13160                         For Each prp In .Properties
13170                           With prp
13180                             If .Name = "AppVersion" Then
13190                               arr_varTmp01(F_APPVER, lngX) = .Value
13200                             ElseIf .Name = "AppDate" Then
13210                               arr_varTmp01(F_APPDATE, lngX) = .Value
13220                             End If
13230                           End With  ' ** prp.
13240                         Next
13250                       End If
13260                     End With  ' ** doc.
13270                   Next

13280                   lngVerTbls = 0&
13290                   ReDim arr_varVerTbl(T_ELEMS, 0)

13300                   For Each tdf In .TableDefs
13310                     With tdf
13320                       If Left(.Name, 4) <> "MSys" And Left(.Name, 4) <> "~TMP" And _
                                .Connect = vbNullString Then  ' ** Skip those pesky system tables.

13330                         If Left(.Name, 3) = "m_V" Then
13340 On Error Resume Next
13350                           Set rst1 = dbsLnk.OpenRecordset(.Name, dbOpenDynaset, dbReadOnly)
13360                           If ERR.Number = 0 Then
13370 On Error GoTo 0
13380                             With rst1
13390                               .MoveFirst
13400                               strTmp04 = CStr(Nz(.Fields(0).Value, 0)) & "." & CStr(Nz(.Fields(1).Value, 0)) & "." & CStr(Nz(.Fields(2).Value, 0))
13410                               arr_varTmp01(F_M_VER, lngX) = strTmp04
13420                               .Close
13430                             End With  ' ** rst1.

13440                           Else
13450                             arr_varTmp01(F_NOTE, lngX) = arr_varTmp01(F_NOTE, lngX) & _
                                    " TBL: " & .Name & "  ERR: " & CStr(ERR.Number) & "  " & ERR.description
13460                             arr_varTmp01(F_NOTE, lngX) = Trim(arr_varTmp01(F_NOTE, lngX))
13470 On Error GoTo 0
13480                           End If

13490                           Set rst1 = Nothing
13500                         End If

                              'm_VP, m_VD, m_VA
                              'vp_MAIN, vd_MAIN, va_MAIN
                              'vp_MINOR, vd_MINOR, va_MINOR
                              'vp_REVISION, vd_REVISION, va_REVISION

13510                         If .Name = "journaltype" Then
                                'Debug.Print "'" & arr_varVerDir(D_DIR, lngW)
13520 On Error Resume Next
13530                           Set rst1 = dbsLnk.OpenRecordset(.Name, dbOpenDynaset, dbReadOnly)
13540                           If ERR.Number = 0 Then
13550 On Error GoTo 0
13560                             With rst1
13570                               If .BOF = True And .EOF = True Then
13580                                 strTmp04 = "NO RECS!"
13590                               Else
13600                                 .MoveLast
13610                                 lngTmp05 = .RecordCount
13620                                 .MoveFirst
13630                                 strTmp04 = CStr(lngTmp05)
13640                                 For lngZ = 1& To lngTmp05
13650                                   strTmp04 = strTmp04 & ", " & ![journaltype]
13660                                   If lngZ < lngTmp05 Then .MoveNext
13670                                 Next
13680                               End If
                                    'Debug.Print "'  " & strTmp04
13690                               .Close
13700                             End With  ' ** rst1.
13710                           Else
13720 On Error GoTo 0
                                  'Debug.Print "'  COULDN'T OPEN!"
13730                           End If
13740                           Set rst1 = Nothing
13750                         End If

13760                         lngVerTbls = lngVerTbls + 1&
13770                         lngY = lngVerTbls - 1&
13780                         ReDim Preserve arr_varVerTbl(T_ELEMS, lngY)
13790                         arr_varVerTbl(T_TNAM, lngY) = .Name
13800                         lngVerFlds = .Fields.Count
13810                         If lngVerFlds = 0& Then
13820                           arr_varTmp01(F_NOTE, lngX) = arr_varTmp01(F_NOTE, lngX) & " TBL: " & .Name & "  FLDS: " & CStr(lngVerFlds)
13830                           arr_varTmp01(F_NOTE, lngX) = Trim(arr_varTmp01(F_NOTE, lngX))
13840                         Else
13850                           arr_varVerTbl(T_FLDS, lngY) = lngVerFlds
13860                           arr_varVerTbl(T_F_ARR, lngY) = Empty
13870                           ReDim arr_varVerFld(FD_ELEMS, (lngVerFlds - 1&))
13880                         End If

13890                         lngZ = -1&
13900                         For Each fld In .Fields
13910                           With fld
13920                             lngZ = lngZ + 1&
13930                             arr_varVerFld(FD_FNAM, lngZ) = .Name
13940                             arr_varVerFld(FD_TYP, lngZ) = .Type
13950                             arr_varVerFld(FD_SIZ, lngZ) = .Size
13960                           End With  ' ** fld.
13970                         Next

13980                         arr_varTmp01(F_TBLS, lngX) = lngVerTbls
13990                         arr_varVerTbl(T_F_ARR, lngY) = arr_varVerFld

14000                       End If

14010                     End With  ' ** This Table: tdf.

14020                   Next  ' ** For each Table in TableDefs: tdf.

14030                   arr_varTmp01(F_T_ARR, lngX) = arr_varVerTbl  ' ** arr_varVerFile().

14040                   .Close
14050                 End With  ' ** This Database: dbsLnk.

14060               Else
                      ' ** Can't open database.
14070                 arr_varVerFile(F_NOTE, lngX) = CStr(ERR.Number) & "  " & ERR.description
14080 On Error GoTo 0
14090               End If  ' ** Err.Number.

14100             End If  ' ** strDocDatabase only.

14110           Next  ' ** For each File in arr_varVerFile(): lngX.

14120           arr_varVerDir(D_F_ARR, lngW) = arr_varTmp01  ' ** arr_varVerFile().

14130         End If  ' ** Directory has files.

14140       Next  ' ** For each Directory in arr_varVerDir(): lngW.

14150       .Close
14160     End With  ' ** wrkLnk.

14170     Set fld = Nothing
14180     Set tdf = Nothing
14190     Set dbsLnk = Nothing
14200     Set wrkLnk = Nothing

14210     Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
14220     DoEvents

14230     With dbsLoc

14240       Set rst1 = .OpenRecordset("tblVersion_Directory", dbOpenDynaset, dbConsistent)
14250       Set rst2 = .OpenRecordset("tblVersion_File", dbOpenDynaset, dbConsistent)
14260       Set rst3 = .OpenRecordset("tblVersion_Table", dbOpenDynaset, dbConsistent)
14270       Set rst4 = .OpenRecordset("tblVersion_Field", dbOpenDynaset, dbConsistent)

14280       For lngW = 0& To (lngVerDirs - 1&)
              ' ******************************************
              ' ** Array: arr_varVerDir()
              ' **
              ' **   Element  Name          Constant
              ' **   =======  ============  ============
              ' **      0     Name          D_DIR
              ' **      1     Path          D_PATH
              ' **      2     Files         D_FILES
              ' **      3     File Array    D_F_ARR
              ' **
              ' ******************************************
14290         lngVerFiles = arr_varVerDir(D_FILES, lngW)
14300         lngDirID = 0&
14310         blnAdd = False
14320         With rst1
14330           If .BOF = True And .EOF = True Then
14340             blnAdd = True
14350           Else
14360             .FindFirst "[verdir_name] = '" & arr_varVerDir(D_DIR, lngW) & "' And [verdir_path] = '" & arr_varVerDir(D_PATH, lngW) & "'"
14370             If .NoMatch = True Then blnAdd = True
14380           End If

14390           If blnAdd = True Then
14400             lngDirsAdd = lngDirsAdd + 1&
14410             .AddNew
14420             ![verdir_name] = arr_varVerDir(D_DIR, lngW)
14430             ![verdir_path] = arr_varVerDir(D_PATH, lngW)
14440             ![verdir_filecnt] = lngVerFiles
14450             ![verdir_datemodified] = Now()
14460             .Update
14470             .Bookmark = .LastModified
14480             lngDirID = ![verdir_id]
14490           Else
14500             lngDirID = ![verdir_id]
14510           End If

14520         End With  ' ** rst1.

14530         arr_varTmp01 = arr_varVerDir(D_F_ARR, lngW)
14540         For lngX = 0& To (lngVerFiles - 1&)
                ' *******************************************************
                ' ** Array: arr_varVerFile()
                ' **
                ' **   Element  Name                        Constant
                ' **   =======  ==========================  ===========
                ' **      0     Name                        F_FNAM
                ' **      1     Trust Accountant Version    F_TA_VER
                ' **      2     Access Version              F_ACC_VER
                ' **      3     Can Open                    F_OPEN
                ' **      4     Tables                      F_TBLS
                ' **      5     Table Array                 F_T_ARR
                ' **      6     Note                        F_NOTE
                ' **      7     m_Vx Version                F_M_VER
                ' **      8     AppVersion                  F_APPVER
                ' **      9     AppDate                     F_APPDATE
                ' **
                ' *******************************************************
14550           lngVerTbls = arr_varTmp01(F_TBLS, lngX)
14560           lngFileID = 0&
14570           blnAdd = False
14580           With rst2
14590             If .BOF = True And .EOF = True Then
14600               blnAdd = True
14610             Else
14620               .FindFirst "[verdir_id] = " & CStr(lngDirID) & _
                      " And [verfile_taver] = '" & arr_varTmp01(F_TA_VER, lngX) & "' And [verfile_name] = '" & arr_varTmp01(F_FNAM, lngX) & "'"
14630               If .NoMatch = True Then blnAdd = True
14640             End If

14650             If blnAdd = True Then
14660               lngFilesAdd = lngFilesAdd + 1&
14670               .AddNew
14680               ![verdir_id] = lngDirID
14690               ![verfile_name] = arr_varTmp01(F_FNAM, lngX)
14700               ![verfile_taver] = arr_varTmp01(F_TA_VER, lngX)
14710               If arr_varTmp01(F_ACC_VER, lngX) <> vbNullString Then
14720                 ![verfile_accver] = arr_varTmp01(F_ACC_VER, lngX)
14730               End If

14740               ![verfile_tblcnt] = lngVerTbls
14750               If arr_varTmp01(F_NOTE, lngX) <> vbNullString Then
14760                 ![verfile_note] = arr_varTmp01(F_NOTE, lngX)
14770               End If

14780               If arr_varTmp01(F_M_VER, lngX) <> vbNullString Then
14790                 ![verfile_m_v_ver] = arr_varTmp01(F_M_VER, lngX)
14800               End If

14810               If arr_varTmp01(F_APPVER, lngX) <> vbNullString Then
14820                 ![verfile_appversion] = arr_varTmp01(F_APPVER, lngX)
14830                 ![verfile_appdate] = arr_varTmp01(F_APPDATE, lngX)
14840               End If

14850               ![verfile_datemodified] = Now()
14860               .Update
14870               .Bookmark = .LastModified
14880               lngFileID = ![verfile_id]
14890             Else
14900               lngFileID = ![verfile_id]
14910               If arr_varTmp01(F_ACC_VER, lngX) <> vbNullString Then
14920                 If IsNull(![verfile_accver]) = True Then
14930                   lngAccVerUpdate = lngAccVerUpdate + 1&
14940                   .Edit
14950                   ![verfile_accver] = arr_varTmp01(F_ACC_VER, lngX)
14960                   ![verfile_datemodified] = Now()
14970                   .Update
14980                 Else
14990                   If ![verfile_accver] <> arr_varTmp01(F_ACC_VER, lngX) Then
15000                     lngAccVerUpdate = lngAccVerUpdate + 1&
15010                     .Edit
15020                     ![verfile_accver] = arr_varTmp01(F_ACC_VER, lngX)
15030                     ![verfile_datemodified] = Now()
15040                     .Update
15050                   End If

15060                 End If

15070               End If

15080               If lngVerTbls > 0& Then
15090                 If ![verfile_tblcnt] = 0& Then
15100                   lngTblCntUpdate = lngTblCntUpdate + 1&
15110                   .Edit
15120                   ![verfile_tblcnt] = lngVerTbls
15130                   ![verfile_datemodified] = Now()
15140                   .Update
15150                 End If

15160               End If

15170               If arr_varTmp01(F_M_VER, lngX) <> vbNullString Then
15180                 If IsNull(![verfile_m_v_ver]) = True Then
15190                   lngNewVerUpdate = lngNewVerUpdate + 1&
15200                   .Edit
15210                   ![verfile_m_v_ver] = arr_varTmp01(F_M_VER, lngX)
15220                   ![verfile_datemodified] = Now()
15230                   .Update
15240                 Else
15250                   If ![verfile_m_v_ver] <> arr_varTmp01(F_M_VER, lngX) Then
15260                     lngNewVerUpdate = lngNewVerUpdate + 1&
15270                     .Edit
15280                     ![verfile_m_v_ver] = arr_varTmp01(F_M_VER, lngX)
15290                     ![verfile_datemodified] = Now()
15300                     .Update
15310                   End If

15320                 End If

15330               End If

15340               If arr_varTmp01(F_APPVER, lngX) <> vbNullString Then
15350                 If IsNull(![verfile_appversion]) = True Then
15360                   lngNewVerUpdate = lngNewVerUpdate + 1&
15370                   .Edit
15380                   ![verfile_appversion] = arr_varTmp01(F_APPVER, lngX)
15390                   ![verfile_appdate] = arr_varTmp01(F_APPDATE, lngX)
15400                   ![verfile_datemodified] = Now()
15410                   .Update
15420                 Else
15430                   If ![verfile_appversion] <> arr_varTmp01(F_APPVER, lngX) Then
15440                     lngNewVerUpdate = lngNewVerUpdate + 1&
15450                     .Edit
15460                     ![verfile_appversion] = arr_varTmp01(F_APPVER, lngX)
15470                     ![verfile_appdate] = arr_varTmp01(F_APPDATE, lngX)
15480                     ![verfile_datemodified] = Now()
15490                     .Update
15500                   End If

15510                 End If

15520               End If

15530               If arr_varTmp01(F_NOTE, lngX) <> vbNullString Then
15540                 If IsNull(![verfile_note]) = True Then
15550                   lngNoteUpdate = lngNoteUpdate + 1&
15560                   .Edit
15570                   ![verfile_note] = arr_varTmp01(F_NOTE, lngX)
15580                   ![verfile_datemodified] = Now()
15590                   .Update
15600                 Else
15610                   If ![verfile_note] <> arr_varTmp01(F_NOTE, lngX) Then
15620                     lngNoteUpdate = lngNoteUpdate + 1&
15630                     .Edit
15640                     ![verfile_note] = arr_varTmp01(F_NOTE, lngX)
15650                     ![verfile_datemodified] = Now()
15660                     .Update
15670                   End If

15680                 End If

15690               End If

15700             End If

15710           End With  ' ** rst2.

15720           If lngVerTbls > 0& Then
15730             arr_varTmp02 = arr_varTmp01(F_T_ARR, lngX)
15740             For lngY = 0& To (lngVerTbls - 1&)
                    ' ******************************************
                    ' ** Array: arr_varVerTbl()
                    ' **
                    ' **   Element  Name           Constant
                    ' **   =======  =============  ===========
                    ' **      0     Name           T_TNAM
                    ' **      1     Fields         T_FLDS
                    ' **      2     Field Array    T_F_ARR
                    ' **
                    ' ******************************************
15750               lngVerFlds = arr_varTmp02(T_FLDS, lngY)
15760               lngTblID = 0&
15770               blnAdd = False
15780               With rst3
15790                 If .BOF = True And .EOF = True Then
15800                   blnAdd = True
15810                 Else
15820                   .FindFirst "[verdir_id] = " & CStr(lngDirID) & " And [verfile_id] = " & CStr(lngFileID) & _
                          " And [vertbl_name] = '" & arr_varTmp02(T_TNAM, lngY) & "'"
15830                   If .NoMatch = True Then blnAdd = True
15840                 End If

15850                 If blnAdd = True Then
15860                   lngTblsAdd = lngTblsAdd + 1&
15870                   .AddNew
15880                   ![verdir_id] = lngDirID
15890                   ![verfile_id] = lngFileID
15900                   ![vertbl_name] = arr_varTmp02(T_TNAM, lngY)
15910                   ![vertbl_fldcnt] = lngVerFlds
15920                   ![vertbl_datemodified] = Now()
15930                   .Update
15940                   .Bookmark = .LastModified
15950                   lngTblID = ![vertbl_id]
15960                 Else
15970                   lngTblID = ![vertbl_id]
15980                 End If

15990               End With  ' ** rst3.

16000               arr_varTmp03 = arr_varTmp02(T_F_ARR, lngY)
16010               For lngZ = 0& To (lngVerFlds - 1&)
                      ' **********************************
                      ' ** Array: arr_varVerFld()
                      ' **
                      ' **   Element  Name    Constant
                      ' **   =======  ======  ==========
                      ' **      0     Name    FD_FNAM
                      ' **      1     Type    FD_TYP
                      ' **      2     Size    FD_SIZ
                      ' **
                      ' **********************************
16020                 With rst4
16030                   blnAdd = False
16040                   If .BOF = True And .EOF = True Then
16050                     blnAdd = True
16060                   Else
16070                     .FindFirst "[verdir_id] = " & CStr(lngDirID) & " And [verfile_id] = " & CStr(lngFileID) & _
                            " And [vertbl_id] = " & CStr(lngTblID) & " And [verfld_name] = '" & arr_varTmp03(FD_FNAM, lngZ) & "'"
16080                     If .NoMatch = True Then blnAdd = True
16090                   End If

16100                   If blnAdd = True Then
16110                     lngFldsAdd = lngFldsAdd + 1&
16120                     .AddNew
16130                     ![verdir_id] = lngDirID
16140                     ![verfile_id] = lngFileID
16150                     ![vertbl_id] = lngTblID
16160                     ![verfld_name] = arr_varTmp03(FD_FNAM, lngZ)
16170                     ![datatype_db_type] = arr_varTmp03(FD_TYP, lngZ)
16180                     ![verfld_size] = arr_varTmp03(FD_SIZ, lngZ)
16190                     ![verfld_datemodified] = Now()
16200                     .Update
16210                   End If

16220                 End With  ' ** rst4.

16230               Next
16240             Next
16250           End If

16260         Next
16270       Next

16280       rst1.Close
16290       rst2.Close
16300       rst3.Close
16310       rst4.Close

16320     End With  ' ** dbsLoc.

16330   Next  ' ** For each different file name: lngV.

16340   dbsLoc.Close
16350   wrkLoc.Close

16360   If lngDirsAdd = 0& And lngFilesAdd = 0& And lngTblsAdd = 0& And lngFldsAdd = 0& Then
16370     Debug.Print "'NO NEW RECORDS ADDED!"
16380   Else
16390     Debug.Print "'ADDED DIRS: " & CStr(lngDirsAdd) & "  FILES: " & CStr(lngFilesAdd) & "  TBLS: " & CStr(lngTblsAdd) & "  FLDS: " & CStr(lngFldsAdd)
16400   End If

16410   If lngAccVerUpdate > 0& Then
16420     Debug.Print "'ACC VERS UPDATED: " & CStr(lngAccVerUpdate)
16430   End If

16440   If lngTblCntUpdate > 0& Then
16450     Debug.Print "'TBL CNTS UPDATED: " & CStr(lngTblCntUpdate)
16460   End If

16470   If lngNewVerUpdate > 0& Then
16480     Debug.Print "'NEW VERS UPDATED: " & CStr(lngNewVerUpdate)
16490   End If

16500   If lngNoteUpdate > 0& Then
16510     Debug.Print "'NOTES UPDATED: " & CStr(lngNoteUpdate)
16520   End If

16530   Beep

EXITP:
16540   Set fsfl = Nothing
16550   Set fsfd2 = Nothing
16560   Set fsfd1 = Nothing
16570   Set fsfls = Nothing
16580   Set fsfds = Nothing
16590   Set fso = Nothing
16600   Set prp = Nothing
16610   Set doc = Nothing
16620   Set fld = Nothing
16630   Set tdf = Nothing
16640   Set rst1 = Nothing
16650   Set rst2 = Nothing
16660   Set rst3 = Nothing
16670   Set rst4 = Nothing
16680   Set dbsLnk = Nothing
16690   Set wrkLnk = Nothing
16700   Version_Doc = blnRetVal
16710   Exit Function

ERRH:
16720   blnRetVal = False
16730   Select Case ERR.Number
        Case Else
16740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16750   End Select
16760   Resume EXITP

      #End If

End Function
