Attribute VB_Name = "modFileUtilities"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modFileUtilities"

'VGC 09/29/2017: CHANGES!

' ** Various file statements/functions:
' **   Name          Statement  Renames a disk file, directory, or folder.
' **                            Syntax: Name oldpathname As newpathname
' **   Kill          Statement  Deletes files from a disk.
' **                            Syntax: Kill pathname
' **   FileCopy      Statement  Copies a file.
' **                            Syntax: FileCopy Source, Destination
' **   RmDir         Statement  Removes an existing directory or folder.
' **                            Syntax: RmDir Path
' **   ChDir         Statement  Changes the current directory or folder.
' **                            Syntax: ChDir Path
' **   MkDir         Statement  Creates a new directory or folder.
' **                            Syntax: MkDir Path
' **   ChDrive       Statement  Changes the current drive.
' **                            Syntax: ChDrive Drive
' **   Dir           Function   Returns a String representing the name of a file, directory, or folder that matches
' **                            a specified pattern or file attribute, or the volume label of a drive.
' **                            Syntax: Dir [(pathname[, attributes])]
' **   CurDir        Function   Returns a Variant (String) representing the current path.
' **                            Syntax: CurDir [(drive)]
' **                            The optional drive argument is a string expression that specifies an existing drive.
' **                            If no drive is specified or if drive is a zero-length string (""), CurDir returns
' **                            the path for the current drive. On the Macintosh, CurDir ignores any drive specified
' **                            and simply returns the path for the current drive.
' **   FileDateTime  Function   Returns a Variant (Date) that indicates the date and time when a file was created or last modified.
' **                            Syntax: FileDateTime (pathname)
' **   FileLen       Function   Returns a Long specifying the length of a file in bytes.
' **                            Syntax: FileLen (pathname)
' **   GetAttr       Function   Returns an Integer representing the attributes of a file, directory, or folder.
' **                            Syntax: GetAttr(pathname)
' **                            The value returned by GetAttr is the sum of the following attribute values:
' **                               0  vbNormal     Normal.
' **                               1  vbReadOnly   Read-only.
' **                               2  vbHidden     Hidden.
' **                               4  vbSystem     System file. Not available on the Macintosh.
' **                              16  vbDirectory  Directory or folder.
' **                              32  vbArchive    File has changed since last backup. Not available on the Macintosh.
' **                              64  vbAlias      Specified file name is an alias. Available only on the Macintosh.

' ** FilterString family variables; filter strings for Windows functions (tblFilterString).
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

' ** VbxDriveType enumeration:
'Public Const vbxDriveUnknown   As Long = 0  ' ** DRIVE_UNKNOWN      The drive type cannot be determined.
'Public Const vbxDriveNoRootDir As Long = 1  ' ** DRIVE_NO_ROOT_DIR  The root path is invalid; for example, there is no volume mounted at the specified path.
Public Const vbxDriveRemovable As Long = 2  ' ** DRIVE_REMOVABLE    The drive has removable media; for example, a floppy drive, thumb drive, or flash card reader.
'Public Const vbxDriveFixed     As Long = 3  ' ** DRIVE_FIXED        The drive has fixed media; for example, a hard drive or flash drive.
'Public Const vbxDriveRemote    As Long = 4  ' ** DRIVE_REMOTE       The drive is a remote (network) drive.
Public Const vbxDriveCDROM     As Long = 5  ' ** DRIVE_CDROM        The drive is a CD-ROM drive.
Public Const vbxDriveRAMDisk   As Long = 6  ' ** DRIVE_RAMDISK      The drive is a RAM disk.

' ** This API declaration is used to return the type of drive from a drive letter.
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" _
  (ByVal nDrive As String) As Long

Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" _
  (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' ** This API declaration is used to return the UNC path from a drive letter.
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" _
  (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

' ** This API declaration is used to return the name of the user that logged into this computer.
Private Declare Function GetUserNameAPI Lib "advapi32.dll" Alias "GetUserNameA" _
  (ByVal lpBuffer As String, nSize As Long) As Long

' ** This API declaration is used to return the name of this computer.
Private Declare Function GetComputerNameAPI Lib "kernel32.dll" Alias "GetComputerNameA" _
  (ByVal lpBuffer As String, nSize As Long) As Long

' ** This API declaration is used to return the name of the default printer for this computer.
Private Declare Function GetDefaultPrinterAPI Lib "Winspool.drv" Alias "GetDefaultPrinterA" _
  (ByVal lpBuffer As String, nSize As Long) As Long
' **

Public Function IsLoaded(ByVal strObjectName As String, Optional intObjectType As Integer = acForm, Optional varIgnoreDesign As Variant) As Boolean
' ** Returns True if the specified form or report is open.
' ** For forms, it only evaluates as true if in Form view or Datasheet view.
' ** Arguments:
' **   strObjectName:   The name of the form or report.
' **   intObjectType:   Type of object. acForm (the default) or acReport.
' **   varIgnoreDesign: For dev, return true status irrespective of view.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "IsLoaded"

        Dim blnRetVal As Boolean

110     blnRetVal = False

        ' ** Validate object type.
120     If (intObjectType <> acForm) And (intObjectType <> acReport) Then
130       blnRetVal = False
140     Else
          ' ** Evaluate status.
150       Select Case intObjectType
          Case acForm
160         If SysCmd(acSysCmdGetObjectState, intObjectType, strObjectName) <> acObjStateClosed Then
170           If IsMissing(varIgnoreDesign) = True Then
180             If Forms(strObjectName).CurrentView <> acCurViewDesign Then
190               blnRetVal = True
200             End If
210           Else
220             If Forms(strObjectName).CurrentView <> acCurViewDesign Or varIgnoreDesign = True Then
230               blnRetVal = True
240             End If
250           End If
260         End If
270       Case acReport
280         If SysCmd(acSysCmdGetObjectState, intObjectType, strObjectName) <> acObjStateClosed Then
290           blnRetVal = True
300         End If
310       Case Else

            ' ** AcObjState enumeration:
            ' **   0  acObjStateClosed  Closed (my own).
            ' **   1  acObjStateOpen    Open.
            ' **   2  acObjStateDirty   Changed but not saved.
            ' **   3  acObjStateNew     New.

            ' ** AcCurrentView enumeration:
            ' **   0  acCurViewDesign        The object is in Design view.
            ' **   1  acCurViewFormBrowse    The object is in Form view.
            ' **   2  acCurViewDatasheet     The object is in Datasheet view.
            ' **   3  acCurViewPivotTable    The object is in PivotTable view.
            ' **   4  acCurViewPivotChart    The object is in PivotChart view.
            ' **   5  acCurViewPreview       The object is in Print Preview.
            ' **   6  acCurViewReportBrowse  The object is in Report view.
            ' **   7  acCurViewLayout        The object is in Layout view.

            ' ** AcObjectType enumeration:
            ' **  -1  acDefault
            ' **   0  acTable            Table
            ' **   1  acQuery            Query
            ' **   2  acForm             Form
            ' **   3  acReport           Report
            ' **   4  acMacro            Macro
            ' **   5  acModule           Module
            ' **   7  acServerView       Server View
            ' **   8  acDiagram          Database Diagram
            ' **   9  acStoredProcedure  Stored Procedure
            ' **  10  acFunction         Function

320       End Select
330     End If

EXITP:
340     IsLoaded = blnRetVal
350     Exit Function

ERRH:
360     blnRetVal = False
370     Select Case ERR.Number
        Case Else
380       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
390     End Select
400     Resume EXITP

End Function

Public Sub DispFrmName(frm As Access.Form, Optional varSub As Variant)
' ** Display a form's name on the title bar, for developer only.

500   On Error GoTo ERRH

        Const THIS_PROC As String = "DispFrmName"

        Dim ctl As Access.Control

        Const JRNL_CAP As String = "Journal Posting Form"

510     If GetUserName = gstrDevUserName Then  ' ** Function: Below.
520       With frm
530         Select Case gblnDev_NoDispName
            Case True
540           If InStr(.Caption, " (") > 0 Then
550             .Caption = Left(.Caption, (InStr(.Caption, " (") - 1))
560           End If
570         Case False
              ' ** Defaults to False.
              ' ** Ctrl+Shift+D on frmMenu_Title cycles On/Off.
580           .Caption = .Caption & " (" & .Name & ")"
590           If .Name = "frmJournal" Then
600             For Each ctl In .Controls
610               With ctl
620                 If Left(.Name, 3) = "frm" Then
630                   If .Visible = True Then
640                     frm.Caption = JRNL_CAP & " (" & frm.Name & ") (" & .Name & ")"
650                     Exit For
660                   End If
670                 End If
680               End With
690             Next
700           ElseIf IsMissing(varSub) = False Then
710             .Caption = Left(.Caption, Len(.Caption) - 1)
720             .Caption = .Caption & ") (" & varSub & ")"  ' ** varSub is a String.
730           End If
740         End Select
750       End With
760     End If

EXITP:
770     Set ctl = Nothing
780     Exit Sub

ERRH:
790     Select Case ERR.Number
        Case Else
800       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
810     End Select
820     Resume EXITP

End Sub

Public Sub DispRptName(rpt As Access.Report, Optional varSub As Variant)
' ** Display a report's name on the title bar, for developer only.

900   On Error GoTo ERRH

        Const THIS_PROC As String = "DispRptName"

910     If GetUserName = gstrDevUserName Then  ' ** Function: Below.
920       With rpt
930         Select Case gblnDev_NoDispName
            Case True
940           If InStr(.Caption, " (") > 0 Then
950             .Caption = Left(.Caption, (InStr(.Caption, " (") - 1))
960           End If
970         Case False
980           .Caption = .Caption & " (" & .Name & ")"
990           If IsMissing(varSub) = False Then
1000            .Caption = Left(.Caption, Len(.Caption) - 1)
1010            .Caption = .Caption & ") (" & varSub & ")"  ' ** varSub is a String.
1020          End If
1030        End Select
1040      End With
1050    End If

EXITP:
1060    Exit Sub

ERRH:
1070    Select Case ERR.Number
        Case Else
1080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1090    End Select
1100    Resume EXITP

End Sub

Public Function CurrentAppName() As String
' ** Return the current application's name.

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentAppName"

        Dim strRetVal As String

1210    strRetVal = vbNullString
1220    strRetVal = CurrentDb.Name
1230    strRetVal = Parse_File(strRetVal)  ' ** Function: Below.

EXITP:
1240    CurrentAppName = strRetVal
1250    Exit Function

ERRH:
1260    strRetVal = vbNullString
1270    Select Case ERR.Number
        Case Else
1280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1290    End Select
1300    Resume EXITP

End Function

Public Function CurrentAppExt() As String
' ** Return the current application's file extension.

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentAppExt"

        Dim strRetVal As String

1410    strRetVal = vbNullString
1420    strRetVal = CurrentAppName  ' ** Function: Above.
1430    strRetVal = Parse_Ext(strRetVal)  ' ** Function: Below.

EXITP:
1440    CurrentAppExt = strRetVal
1450    Exit Function

ERRH:
1460    strRetVal = vbNullString
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Function

Public Function CurrentAppID() As Long
' ** Return the current application's dbs_id.

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentAppID"

        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngRetVal As Long

1610    lngRetVal = 1&  ' ** Just default to Trust.
1620    strTmp01 = CurrentDb.Name
1630    strTmp01 = Parse_File(strTmp01)  ' ** Function: Below.
1640    varTmp00 = DLookup("[dbs_id]", "tblTemplate_Database", "[dbs_name] = '" & strTmp01 & "'")
1650    If IsNull(varTmp00) = True Then
1660      strTmp01 = Left(strTmp01, (Len(strTmp01) - 1)) & IIf(Right(strTmp01, 1) = "b", "e", "b")
1670      varTmp00 = DLookup("[dbs_id]", "tblTemplate_Database", "[dbs_name] = '" & strTmp01 & "'")
1680    End If
1690    lngRetVal = varTmp00

EXITP:
1700    CurrentAppID = lngRetVal
1710    Exit Function

ERRH:
1720    lngRetVal = 1&
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Function

Public Function CurrentAppPath() As String
' ** Return the current application's path, WITHOUT FINAL BACKSLASH!

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentAppPath"

        Dim strRetVal As String

1810    strRetVal = vbNullString
1820    strRetVal = CurrentDb.Name
1830    strRetVal = Parse_Path(strRetVal)  ' ** Function: Below.

EXITP:
1840    CurrentAppPath = strRetVal
1850    Exit Function

ERRH:
1860    strRetVal = vbNullString
1870    Select Case ERR.Number
        Case Else
1880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1890    End Select
1900    Resume EXITP

End Function

Public Function CurrentBackendPath() As String
' ** Returns the full path to TrustDta.mdb, WITHOUT FINAL BACKSLASH!

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentBackendPath"

        Dim dbs As DAO.Database, tdf As DAO.TableDef
        Dim strConnect As String
        Dim strRetVal As String

2010    strRetVal = vbNullString

2020    Set dbs = CurrentDb
2030    With dbs
2040      Set tdf = .TableDefs("m_VD")
2050      With tdf
2060        If .Connect <> vbNullString Then
2070          strConnect = .Connect
2080          If InStr(strConnect, gstrFile_DataName) > 0 Then
2090            strConnect = Mid(strConnect, (InStr(strConnect, LNK_IDENT) + Len(LNK_IDENT)))
2100            strConnect = Parse_Path(strConnect)  ' ** Function: Below.
2110            strRetVal = strConnect
2120          End If
2130        End If
2140      End With
2150      .Close
2160    End With

EXITP:
2170    Set tdf = Nothing
2180    Set dbs = Nothing
2190    CurrentBackendPath = strRetVal
2200    Exit Function

ERRH:
2210    strRetVal = vbNullString
2220    Select Case ERR.Number
        Case Else
2230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2240    End Select
2250    Resume EXITP

End Function

Public Function CurrentBackendPathFile(strTableName As String) As String
' ** Returns path and file name of the backend database for a specified table.

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentBackendPathFile"

        Dim dbs As DAO.Database, tbl As DAO.TableDef
        Dim strConnect As String  ' ** The connect string stored in a linked table.
        Dim strRetVal As String

        ' ** Get CONNECT string from one data table on backend file.
2310    Set dbs = CurrentDb
2320    With dbs
2330      Set tbl = dbs.TableDefs(strTableName)
2340      strConnect = tbl.Connect
2350      If strConnect <> vbNullString Then
            ' ** In case it's a local table.
2360        strRetVal = Mid(strConnect, (InStr(strConnect, LNK_IDENT) + Len(LNK_IDENT)))
2370      End If
2380      .Close
2390    End With

EXITP:
2400    Set tbl = Nothing
2410    Set dbs = Nothing
2420    CurrentBackendPathFile = strRetVal
2430    Exit Function

ERRH:
2440    strRetVal = vbNullString
2450    Select Case ERR.Number
        Case 3265  ' ** Item not found in this collection.
          ' ** Ignore.
2460    Case Else
2470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2480    End Select
2490    Resume EXITP

End Function

Public Function CurrentBackendCompact(blnSkipUI As Boolean, Optional varCallingForm As Variant) As Boolean

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentBackendCompact"

        Dim dbs As DAO.Database, tdf As DAO.TableDef
        Dim strDatabase As String, strArchive As String, strAuxiliary As String
        Dim strDatabaseCompacted As String, strArchiveCompacted As String, strAuxiliaryCompacted As String
        Dim blnRelink_Dta As Boolean, blnRelink_Arch As Boolean, blnRelink_Aux As Boolean
        Dim strNow As String, strCallingForm As String
        Dim blnAuxLoc As Boolean
        Dim varTmp00 As Variant
        Dim msgResponse As VbMsgBoxResult
        Dim blnRetVal As Boolean

2510    blnRetVal = True
2520    blnRelink_Dta = False: blnRelink_Arch = False: blnRelink_Aux = False

2530    varTmp00 = DLookup("[seclic_auxloc]", "tblSecurity_License")
2540    Select Case IsNull(varTmp00)
        Case True
2550      blnAuxLoc = False
2560    Case False
2570      Select Case varTmp00
          Case True
2580        blnAuxLoc = True
2590      Case False
2600        blnAuxLoc = False
2610      End Select
2620    End Select

2630    If blnSkipUI = False Then
2640      DoCmd.Hourglass False
2650      Beep
2660      msgResponse = MsgBox("Compacting the data files will require that Trust Accountant be closed." & vbCrLf & vbCrLf & _
            "Do you want to continue?", vbQuestion + vbYesNo, "Compact Trust Data")
2670    Else
2680      msgResponse = vbYes
2690    End If
2700    DoEvents

2710    If msgResponse = vbYes Then

2720      DoCmd.Hourglass True
2730      DoEvents

2740      Select Case IsMissing(varCallingForm)
          Case True
2750        strCallingForm = vbNullString
2760      Case False
2770        strCallingForm = varCallingForm
2780      End Select

2790      strNow = year(Now) & month(Now) & day(Now) & hour(Now) & Minute(Now)
2800      strDatabase = gstrTrustDataLocation & gstrFile_DataName
2810      strArchive = gstrTrustDataLocation & gstrFile_ArchDataName
2820      Select Case blnAuxLoc
          Case True
2830        gstrTrustAuxLocation = CurrentAppPath & LNK_SEP  ' ** Module Function: modFileUtilities.
2840      Case False
2850        gstrTrustAuxLocation = gstrTrustDataLocation
2860      End Select  ' ** blnAuxLoc.
2870      strAuxiliary = gstrTrustAuxLocation & gstrFile_AuxDataName
2880      strDatabaseCompacted = gstrTrustDataLocation & "TrustDta" & strNow & ".mdb"
2890      strArchiveCompacted = gstrTrustDataLocation & "TrstArch" & strNow & ".mdb"
2900      strAuxiliaryCompacted = gstrTrustAuxLocation & "TrstAux" & strNow & ".mdb"
2910      DoEvents

2920      If Dir(gstrTrustDataLocation & gstrFile_DataLockfile) <> vbNullString Then
2930        blnRetVal = Tbl_DelAllLinks(gstrFile_DataName)  ' ** Module Function: modBackup.
2940        If blnRetVal = True Then
2950          blnRelink_Dta = True
2960        End If
2970      End If
2980      DoEvents
2990      If blnRetVal = True Then
3000        If Dir(gstrTrustDataLocation & gstrFile_ArchDataLockfile) <> vbNullString Then
3010          blnRetVal = Tbl_DelAllLinks(gstrFile_ArchDataName)  ' ** Module Function: modBackup.
3020          If blnRetVal = True Then
3030            blnRelink_Arch = True
3040          End If
3050        End If
3060      End If
3070      DoEvents
3080      If blnRetVal = True Then
3090        If Dir(gstrTrustAuxLocation & gstrFile_AuxDataLockfile) <> vbNullString Then
3100          blnRetVal = Tbl_DelAllLinks(gstrFile_AuxDataName)  ' ** Module Function: modBackup.
3110          If blnRetVal = True Then
3120            blnRelink_Aux = True
3130          End If
3140        End If
3150      End If
3160      DoEvents

3170      CurrentDb.TableDefs.Refresh
3180      DoEvents
3190      CurrentDb.TableDefs.Refresh
3200      DoEvents

3210      ForcePause 5  ' ** Module Function: modCodeUtilities.

3220      If blnRetVal = True Then

            ' ** Close the BE first, then test for REMAINING LDBs so that we aren't seeing ourselves.
3230  On Error Resume Next
3240        gdbsDBLock.Close
3250  On Error GoTo ERRH

3260        Set gdbsDBLock = Nothing

            ' ** Check for open MDBs.
            'If Dir(gstrTrustDataLocation & gstrFile_DataLockfile) <> vbNullString Then
            '  blnRetVal = False
            '  DoCmd.Hourglass False
            '  MsgBox "There is someone in the data file." & vbCrLf & vbCrLf & _
            '    "You can't compact at this time.", vbInformation + vbOKOnly, "Compact Denied"
            'Else

            '  If Dir(gstrTrustDataLocation & gstrFile_ArchDataLockfile) <> vbNullString Then
            '    blnRetVal = False
            '    DoCmd.Hourglass False
            '    MsgBox "There is someone in the archive file." & vbCrLf & vbCrLf & _
            '      "You can't compact at this time.", vbInformation + vbOKOnly, "Compact Denied"
            '  Else

            '    If Dir(gstrTrustDataLocation & gstrFile_AuxDataLockfile) <> vbNullString Then
            '      blnRetVal = False
            '      DoCmd.Hourglass False
            '      MsgBox "There is someone in the auxiliary file." & vbCrLf & vbCrLf & _
            '        "You can't compact at this time.", vbInformation + vbOKOnly, "Compact Denied"
            '    Else

3270        gblnMessage = True ' ** I'm going to use this to signal frmLinkData that it's going to QuitNow().
3280        DoEvents

            ' ** Compact MDBs.
3290        DAO.CompactDatabase strDatabase, strDatabaseCompacted
3300        DoEvents
3310        DAO.CompactDatabase strArchive, strArchiveCompacted
3320        DoEvents
3330        DAO.CompactDatabase strAuxiliary, strAuxiliaryCompacted
3340        DoEvents

            ' ** Rename MDBs to _OLD.
3350        Name strDatabase As gstrTrustDataLocation & "TrustDta_OLD" & ".mdb"
3360        Name strArchive As gstrTrustDataLocation & "TrstArch_OLD" & ".mdb"
3370        Name strAuxiliary As gstrTrustAuxLocation & "TrstAux_OLD" & ".mdb"
3380        DoEvents

            ' ** Rename new to MDB.
3390        Name strDatabaseCompacted As strDatabase
3400        Name strArchiveCompacted As strArchive
3410        Name strAuxiliaryCompacted As strAuxiliary
3420        DoEvents

            ' ** Delete _OLD.
3430        Kill gstrTrustDataLocation & "TrustDta_OLD" & ".mdb"
3440        Kill gstrTrustDataLocation & "TrstArch_OLD" & ".mdb"
3450        Kill gstrTrustAuxLocation & "TrstAux_OLD" & ".mdb"
3460        DoEvents

3470        If blnRelink_Dta = True Then
3480          blnRetVal = Tbl_AddAllLinks(gstrTrustDataLocation & gstrFile_DataName)  ' ** Module Function: modBackup.
3490        End If
3500        DoEvents
3510        If blnRelink_Arch = True Then
3520          blnRetVal = Tbl_AddAllLinks(gstrTrustDataLocation & gstrFile_ArchDataName)  ' ** Module Function: modBackup.
3530        End If
3540        DoEvents
3550        If blnRelink_Aux = True Then
3560          blnRetVal = Tbl_AddAllLinks(gstrTrustAuxLocation & gstrFile_AuxDataName)  ' ** Module Function: modBackup.
3570        End If
3580        DoEvents

3590        DoCmd.Hourglass False

3600        If blnRetVal = True And (blnRelink_Dta = True And blnRelink_Arch = True And blnRelink_Aux = True) Then
3610          MsgBox "Successfully Compacted.", vbInformation + vbOKOnly, "Compact Successful"
3620          Select Case strCallingForm
              Case "frmLinkData"
3630            gblnMessage = True
3640            Forms(strCallingForm).cmdClose.SetFocus
3650          Case Else
                ' ** Nothing at the moment.
3660          End Select
3670        Else
3680          Set dbs = CurrentDb
3690          With dbs
3700            For Each tdf In .TableDefs
3710              With tdf
3720                If .Connect <> vbNullString Then
3730                  .RefreshLink
3740                End If
3750              End With
3760            Next
3770            .Close
3780          End With
3790          If CurrentUser = "Superuser" Or GetUserName = gstrDevUserName Then   ' ** Internal Access Function: Trust Accountant login.
3800            gblnMessage = True
3810            MsgBox "Successfully Compacted.", vbInformation + vbOKOnly, "Compact Successful" '  ' ** Module Function: modFileUtilities.
3820            Select Case strCallingForm
                Case "frmLinkData"
3830              Forms(strCallingForm).cmdClose.SetFocus
3840            Case Else
                  ' ** Nothing at the moment.
3850            End Select
3860          Else
3870            MsgBox "Successfully Compacted." & vbCrLf & vbCrLf & _
                  "Trust Accountant will now exit.", vbInformation + vbOKOnly, "Compact Successful"
3880            QuitNow  ' ** Module Procedure: modFileUtilities.
3890          End If
3900        End If

            '    End If
            '  End If
            'End If

3910      End If  ' ** blnRetVal.
3920    Else
3930      blnRetVal = False
3940    End If

EXITP:
3950    Set tdf = Nothing
3960    Set dbs = Nothing
3970    CurrentBackendCompact = blnRetVal
3980    Exit Function

ERRH:
3990    blnRetVal = False
4000    Select Case ERR.Number
        Case Else
4010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4020    End Select
4030    Resume EXITP

End Function

Public Function CurrentDBSize() As Variant

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrentDBSize"

        Dim fso As Scripting.FileSystemObject, fsf As Scripting.File
        Dim varRetVal As Variant

4110    Set fso = CreateObject("Scripting.FileSystemObject")
4120    Set fsf = fso.GetFile(CurrentDb.Name)

4130    varRetVal = fsf.Size

EXITP:
4140    Set fsf = Nothing
4150    Set fso = Nothing
4160    CurrentDBSize = varRetVal
4170    Exit Function

ERRH:
4180    varRetVal = 0#
4190    Select Case ERR.Number
        Case Else
4200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4210    End Select
4220    Resume EXITP

End Function

Public Function DirExists(strDir As String) As Boolean
' ** Determines if the named directory exists.

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "DirExists"

        Dim strTmp01 As String
        Dim blnRetVal As Boolean

4310    If Right(strDir, 1) = LNK_SEP Then
4320      strTmp01 = strDir & "."
4330    Else
4340      strTmp01 = strDir & LNK_SEP & "."
4350    End If

4360    blnRetVal = Len(Dir$(strTmp01, vbDirectory)) > 0

EXITP:
4370    DirExists = blnRetVal
4380    Exit Function

ERRH:
4390    blnRetVal = False
4400    Select Case ERR.Number
        Case 52  ' ** Bad file name or number.
          ' ** If there's nothing in the drive, it'll error,
          ' ** like me clicking OK without putting anything in it!
4410      If gblnBeenToBackup = False Then
4420        MsgBox "The " & strDir & " is empty.", vbInformation + vbOKOnly, "Drive Is Empty"
4430      Else
            ' ** This is coming from RestoreFromDrive() in modBackup.
4440        gblnBeenToBackup = False
4450      End If
4460    Case Else
4470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4480    End Select
4490    Resume EXITP

End Function

Public Function DirExists2(strPath As String, strDirNameBase As String) As String
' ** Create a directory, incrementing its suffix
' ** if one is already present, and return its name.

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "DirExists2"

        Dim fso As Scripting.FileSystemObject, fsfd1 As Scripting.Folder, fsfd2 As Scripting.Folder, fsfds As Scripting.Folders
        Dim lngDirs As Long, arr_varDir() As Variant
        Dim intDirNum As Integer
        Dim lngX As Long, lngE As Long
        Dim strRetVal As String

        Const D_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const D_NAM As Integer = 0
        Const D_SFX As Integer = 1

4510    strRetVal = vbNullString

4520    lngDirs = 0&
4530    ReDim arr_varDir(D_ELEMS, 0)

4540    Set fso = CreateObject("Scripting.FileSystemObject")
4550    With fso
4560      Set fsfd1 = .GetFolder(strPath)
4570      With fsfd1
4580        Set fsfds = .SubFolders
4590        For Each fsfd2 In fsfds
4600          With fsfd2
4610            If Left(.Name, Len(strDirNameBase)) = strDirNameBase Then
4620              lngDirs = lngDirs + 1&
4630              lngE = lngDirs - 1&
4640              ReDim Preserve arr_varDir(D_ELEMS, lngE)
4650              arr_varDir(D_NAM, lngE) = .Name
4660              If Len(.Name) > Len(strDirNameBase) Then
4670                arr_varDir(D_SFX, lngE) = Mid(.Name, (Len(strDirNameBase) + 1))
4680              Else
4690                arr_varDir(D_SFX, lngE) = vbNullString
4700              End If
4710            End If
4720          End With
4730        Next
4740        If lngDirs = 0& Then
4750          strRetVal = strPath & LNK_SEP & strDirNameBase & "1"
4760          fso.CreateFolder strRetVal
4770        Else
4780          intDirNum = 0
4790          For lngX = 0& To (lngDirs - 1&)
4800            If IsNumeric(arr_varDir(D_SFX, lngX)) Then
4810              If Val(arr_varDir(D_SFX, lngX)) > intDirNum Then
4820                intDirNum = Val(arr_varDir(D_SFX, lngX))
4830              End If
4840            End If
4850          Next
4860          strRetVal = strPath & LNK_SEP & strDirNameBase & CStr(intDirNum + 1)
4870          fso.CreateFolder strRetVal
4880        End If
4890      End With
4900    End With

EXITP:
4910    Set fsfd2 = Nothing
4920    Set fsfds = Nothing
4930    Set fsfd1 = Nothing
4940    Set fso = Nothing
4950    DirExists2 = strRetVal
4960    Exit Function

ERRH:
4970    strRetVal = vbNullString
4980    Select Case ERR.Number
        Case Else
4990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5000    End Select
5010    Resume EXITP

End Function

Public Function FileExists(strFile As String) As Boolean
' ** Determines if the named file exists; must include full path.

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "FileExists"

        Dim blnRetVal As Boolean

5110    blnRetVal = False

5120    If strFile <> vbNullString Then
5130      If Dir(strFile) <> vbNullString Then
5140        blnRetVal = True
5150      End If
5160    End If

EXITP:
5170    FileExists = blnRetVal
5180    Exit Function

ERRH:
5190    blnRetVal = False
5200    Select Case ERR.Number
        Case Else
5210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5220    End Select
5230    Resume EXITP

End Function

Public Function FormClose(strFormName As String) As Boolean
' ** If the specified form is open, close it.

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormClose"

        Dim blnRetVal As Boolean

5310    blnRetVal = False

5320    If IsLoaded(strFormName, acForm) = True Then  ' ** Function: Above.
5330      blnRetVal = True
5340      DoCmd.Close acForm, strFormName, acSaveNo
5350    End If

EXITP:
5360    FormClose = blnRetVal
5370    Exit Function

ERRH:
5380    blnRetVal = False
5390    Select Case ERR.Number
        Case Else
5400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5410    End Select
5420    Resume EXITP

End Function

Public Function ControlExists(strCtlName As String, ctls As Object) As Boolean  'Access.Controls

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "ControlExists"

        Dim ctl As Access.Control
        Dim blnRetVal As Boolean

5510    blnRetVal = False

5520    If ctls.Count > 0 Then
5530      For Each ctl In ctls
5540        With ctl
5550          If .Name = strCtlName Then
5560            blnRetVal = True
5570            Exit For
5580          End If
5590        End With
5600      Next
5610    End If

EXITP:
5620    Set ctl = Nothing
5630    ControlExists = blnRetVal
5640    Exit Function

ERRH:
5650    blnRetVal = False
5660    Select Case ERR.Number
        Case Else
5670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5680    End Select
5690    Resume EXITP

End Function

Public Function TableEmpty(strTblName As String) As Boolean
' ** Because my numbering of the purge-table queries seems to change so often...

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "TableEmpty"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

        Const TBL_EMPTY As String = "qryTmp_Table_Empty_"

5710    blnRetVal = False

5720    If strTblName <> vbNullString Then
5730      Set dbs = CurrentDb
5740      With dbs
5750        For Each qdf In .QueryDefs
5760          With qdf
5770            If Left(.Name, Len(TBL_EMPTY)) = TBL_EMPTY Then
                  ' ** Example: qryTmp_Table_Empty_30_tblTemplate_MasterAsset
5780              If Right(.Name, Len(strTblName)) = strTblName Then
5790                blnRetVal = True
5800                strTmp01 = .Name
5810                Exit For
5820              End If
5830            End If
5840          End With
5850        Next
5860        If blnRetVal = True Then
5870          If TableExists(strTblName) = True Then  ' ** Function: Above.
5880            Set qdf = .QueryDefs(strTmp01)
5890            qdf.Execute
5900          End If
5910        End If
5920        .Close
5930      End With

5940    End If

EXITP:
5950    Set qdf = Nothing
5960    Set dbs = Nothing
5970    TableEmpty = blnRetVal
5980    Exit Function

ERRH:
5990    blnRetVal = False
6000    Select Case ERR.Number
        Case Else
6010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6020    End Select
6030    Resume EXITP

End Function

Public Function TableExists(strTable As String, Optional varInBackend As Variant, Optional varDb As Variant, Optional varBadLink As Variant) As Boolean
' ** Determine if the table exists.

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "TableExists"

        Dim dbs As DAO.Database, tdf As DAO.TableDef
        Dim blnInBackend As Boolean
        Dim strBackendDB As String, strBackendPath As String
        Dim lngTdfs As Long
        Dim blnAuxLoc As Boolean, blnBadLink As Boolean, blnIsBadLink As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

6110    blnRetVal = False

6120    strBackendPath = vbNullString
6130    Select Case IsMissing(varInBackend)
        Case True
6140      blnInBackend = False
6150      strBackendDB = vbNullString
6160    Case False
6170      blnInBackend = CBool(varInBackend)
6180      Select Case IsMissing(varDb)
          Case True
6190        strBackendDB = gstrFile_DataName
6200      Case False
6210        strBackendDB = varDb
6220        If InStr(strBackendDB, "\") > 0 Then
              ' ** Both file and path were passed.
6230          strBackendPath = Parse_Path(strBackendDB)  ' ** Function: Below.
6240          strBackendDB = Parse_File(strBackendDB)  ' ** Function: Below.
6250        End If
6260      End Select
6270    End Select

6280    blnIsBadLink = False
6290    Select Case IsMissing(varBadLink)
        Case True
6300      blnBadLink = False
6310    Case False
6320      blnBadLink = CBool(varBadLink)
6330    End Select

6340    varTmp00 = DLookup("[seclic_auxloc]", "tblSecurity_License")
6350    Select Case IsNull(varTmp00)
        Case True
6360      blnAuxLoc = False
6370    Case False
6380      Select Case varTmp00
          Case True
6390        blnAuxLoc = True
6400      Case False
6410        blnAuxLoc = False
6420      End Select
6430    End Select

6440    If blnBadLink = True Then
6450      Select Case blnInBackend
          Case True
            ' ** Haven't programmed this...
6460      Case False
6470        Set dbs = CurrentDb
6480  On Error Resume Next
6490        lngTdfs = dbs.TableDefs.Count
6500        If ERR.Number <> 0 Then
6510  On Error GoTo ERRH
6520          blnIsBadLink = True
6530        Else
6540  On Error GoTo ERRH
6550        End If
6560        If blnIsBadLink = False Then
6570          For lngX = 0& To (lngTdfs - 1&)
6580  On Error Resume Next
6590            Set tdf = dbs.TableDefs(lngX)
6600            If ERR.Number <> 0 Then
6610  On Error GoTo ERRH
6620              blnIsBadLink = True
6630              Exit For
6640            Else
6650  On Error GoTo ERRH
6660            End If
6670            With tdf
6680              If .Name = strTable Then
6690  On Error Resume Next
6700                varTmp00 = .Fields.Count
6710                If ERR.Number <> 0 Then
6720  On Error GoTo ERRH
6730                  blnIsBadLink = True
6740                Else
6750  On Error GoTo ERRH
6760                  If varTmp00 = 0 Then
6770                    blnIsBadLink = True
6780                  End If
6790                End If
6800                Exit For
6810              End If
6820            End With
6830          Next
6840        End If
6850        dbs.Close
6860        Set dbs = Nothing
6870      End Select  ' ** blnInBackend.
6880    End If  ' ** blnBadLink.

6890    If blnBadLink = False Or (blnBadLink = True And blnIsBadLink = False) Then
6900      Select Case blnInBackend
          Case True
            ' ** gstrTrustDataLocation INCLUDES FINAL BACKSLASH!
6910        If gstrTrustDataLocation = vbNullString Then
6920          blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
6930        End If
6940        Select Case blnAuxLoc
            Case True
6950          gstrTrustAuxLocation = CurrentAppPath & LNK_SEP  ' ** Module Function: modFileUtilities.
6960        Case False
6970          gstrTrustAuxLocation = gstrTrustDataLocation
6980        End Select
6990        If InStr(strBackendDB, "TrustAux") > 0 Then
7000          If strBackendPath = vbNullString Then
7010            Set dbs = DBEngine.OpenDatabase(gstrTrustAuxLocation & strBackendDB, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
7020          Else
7030            Set dbs = DBEngine.OpenDatabase(strBackendPath & LNK_SEP & strBackendDB, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
7040          End If
7050        Else
7060          If strBackendPath = vbNullString Then
7070            Set dbs = DBEngine.OpenDatabase(gstrTrustDataLocation & strBackendDB, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
7080          Else
7090            Set dbs = DBEngine.OpenDatabase(strBackendPath & LNK_SEP & strBackendDB, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
7100          End If
7110        End If
7120      Case False
7130        Set dbs = CurrentDb
7140      End Select
7150      With dbs
7160  On Error Resume Next
7170        lngTdfs = .TableDefs.Count
7180  On Error GoTo ERRH
7190        For lngX = 0& To (lngTdfs - 1&)  'For Each tdf In .TableDefs
7200          Set tdf = .TableDefs(lngX)
7210          With tdf
7220            If .Name = strTable Then
7230              blnRetVal = True
7240              Exit For
7250            End If
7260          End With
7270        Next
7280        .Close
7290      End With
7300    End If

7310    If blnBadLink = True Then
7320      gblnBadLink = blnIsBadLink
7330    End If

EXITP:
7340    Set tdf = Nothing
7350    Set dbs = Nothing
7360    TableExists = blnRetVal
7370    Exit Function

ERRH:
7380    blnRetVal = False
7390    Select Case ERR.Number
        Case Else
7400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7410    End Select
7420    Resume EXITP

End Function

Public Function TableDelete(strTableName As String) As Boolean
' ** Delete a table if it exists.

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "TableDelete"

        Dim dbs As DAO.Database, tdf As DAO.TableDef
        Dim blnRetVal As Boolean

7510    blnRetVal = True  ' ** Unless proven otherwise.

7520    Set dbs = CurrentDb
7530    For Each tdf In dbs.TableDefs
7540      If tdf.Name = strTableName Then
7550  On Error Resume Next
7560        dbs.TableDefs.Delete strTableName
7570        If ERR.Number <> 0 Then
7580          If ERR.Number = 3211 Then    ' ** The database engine couldn't lock table '|' because
7590            If Reports.Count > 0 Then  ' ** it's already in use by another person or process.
7600  On Error GoTo ERRH
7610              Do While Reports.Count > 0
7620                DoCmd.Close acReport, Reports(0).Name
7630              Loop
7640              dbs.TableDefs.Delete strTableName
7650            Else
7660              zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7670  On Error GoTo ERRH
7680              blnRetVal = False
7690            End If
7700          Else
7710            zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7720  On Error GoTo ERRH
7730            blnRetVal = False
7740          End If
7750        Else
7760  On Error GoTo ERRH
7770        End If
7780        dbs.TableDefs.Refresh
7790        Exit For
7800      End If
7810    Next tdf

7820    dbs.Close

EXITP:
7830    Set tdf = Nothing
7840    Set dbs = Nothing
7850    TableDelete = blnRetVal
7860    Exit Function

ERRH:
7870    blnRetVal = False
7880    Select Case ERR.Number
        Case 3211  ' ** The database engine couldn't lock table '|' because it's already in use by another person or process.
7890      Beep
7900      MsgBox "You may have a report open that is preventing a temporary table from being re-created.", _
            vbCritical + vbOKOnly, "Error Creating Temporary Table"
7910    Case Else
7920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7930    End Select
7940    Resume EXITP

End Function

Public Function TableDescriptionUpdate(Optional varPathFile As Variant) As Boolean

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "TableDescriptionUpdate"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, tdf As DAO.TableDef, prp As Object, rst As DAO.Recordset
        Dim strConn As String
        Dim lngDbs As Long, arr_varDb As Variant
        Dim lngTbls As Long, arr_varTbl() As Variant
        Dim blnFound As Boolean, blnOptional As Boolean, blnIsDemo As Boolean
        Dim strTmp01 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varDb().
        Const D_DNAM As Integer = 1
        Const D_PATH As Integer = 2

        ' ** Array: arr_varTbl().
        Const T_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const T_MDB  As Integer = 0
        Const T_TNAM As Integer = 1
        Const T_DSC  As Integer = 2

8010    blnRetVal = True
8020    blnOptional = False

8030    Set dbs = CurrentDb
8040    With dbs
          ' ** ;DATABASE=C:\VictorGCS_Clients\TrustAccountant\NewWorking\EmptyDatabase\TrustDta.mdb  '## OK
8050      If IsMissing(varPathFile) = True Then
8060        Set rst = .OpenRecordset("tblDatabase", dbOpenDynaset, dbReadOnly)
8070        With rst
8080          If .BOF = True And .EOF = True Then
8090            blnRetVal = False
8100          Else
8110            .MoveLast
8120            lngDbs = .RecordCount
8130            .MoveFirst
8140            arr_varDb = .GetRows(lngDbs)
                ' *****************************************************
                ' ** Array: arr_varDb()
                ' **
                ' **   Field  Element  Name                Constant
                ' **   =====  =======  ==================  ==========
                ' **     1       0     dbs_id
                ' **     2       1     dbs_name            D_DNAM
                ' **     3       2     dbs_path            D_PATH
                ' **     4       3     dbs_tbl_cnt
                ' **     5       4     dbs_rel_cnt
                ' **     6       5     dbs_idx_cnt
                ' **     7       6     dbs_datemodified
                ' **
                ' *****************************************************
8150          End If
8160          .Close
8170        End With
8180        strConn = vbNullString
            'strConn = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\EmptyDatabase" & LNK_SEP & gstrFile_DataName  '## OK
8190      Else
8200        blnOptional = True
8210        strConn = varPathFile 'gstrTrustDataLocation & gstrFile_DataName
8220        lngDbs = 1&  ' ** Rather odd, but I'm borrowing this array for a moment.
8230        ReDim arr_varTbl(6, 0)
8240        arr_varTbl(D_DNAM, 0) = Parse_File(strConn)  ' ** Function: Above.
8250        arr_varTbl(D_PATH, 0) = Parse_Path(strConn)  ' ** Function: Above.
8260        arr_varDb = arr_varTbl
8270      End If
8280      .Close
8290    End With

8300    lngTbls = 0&
8310    ReDim arr_varTbl(T_ELEMS, 0)

8320    blnIsDemo = False

8330    If Len(TA_SEC) < Len(TA_SEC2) Then  ' ** This is not the Demo.
8340  On Error Resume Next
8350      Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
8360      If ERR.Number <> 0 Then
8370  On Error GoTo ERRH
8380  On Error Resume Next
8390        Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
8400        If ERR.Number <> 0 Then
8410  On Error GoTo ERRH
8420  On Error Resume Next
8430          Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
8440          If ERR.Number <> 0 Then
8450  On Error GoTo ERRH
8460  On Error Resume Next
8470            Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
8480            If ERR.Number <> 0 Then
8490  On Error GoTo ERRH
8500  On Error Resume Next
8510              Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
8520              If ERR.Number <> 0 Then
8530  On Error GoTo ERRH
8540  On Error Resume Next
8550                Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
8560                If ERR.Number <> 0 Then
8570  On Error GoTo ERRH
8580  On Error Resume Next
8590                  Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
8600  On Error GoTo ERRH
8610                Else
8620  On Error GoTo ERRH
8630                End If
8640              Else
8650  On Error GoTo ERRH
8660              End If
8670            Else
8680  On Error GoTo ERRH
8690            End If
8700          Else
8710  On Error GoTo ERRH
8720          End If
8730        Else
8740  On Error GoTo ERRH
8750        End If
8760      Else
8770  On Error GoTo ERRH
8780      End If
8790    Else  ' ** This is the Demo.
8800      blnIsDemo = True
8810  On Error Resume Next
8820      Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New Demo.
8830      If ERR.Number <> 0 Then
8840  On Error GoTo ERRH
8850  On Error Resume Next
8860        Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New non-Demo.
8870        If ERR.Number <> 0 Then
8880  On Error GoTo ERRH
8890  On Error Resume Next
8900          Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
8910          If ERR.Number <> 0 Then
8920  On Error GoTo ERRH
8930  On Error Resume Next
8940            Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old non-Demo.
8950            If ERR.Number <> 0 Then
8960  On Error GoTo ERRH
8970  On Error Resume Next
8980              Set wrk = CreateWorkspace("tmpDB", "TADemo", TA_SEC4, dbUseJet)  ' ** New Demo Admin.
8990              If ERR.Number <> 0 Then
9000  On Error GoTo ERRH
9010  On Error Resume Next
9020                Set wrk = CreateWorkspace("tmpDB", "Demo", "TA_SEC8", dbUseJet)  ' ** Old Demo Admin.
9030                If ERR.Number <> 0 Then
9040  On Error GoTo ERRH
9050  On Error Resume Next
9060                  Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
9070  On Error GoTo ERRH
9080                Else
9090  On Error GoTo ERRH
9100                End If
9110              Else
9120  On Error GoTo ERRH
9130              End If
9140            Else
9150  On Error GoTo ERRH
9160            End If
9170          Else
9180  On Error GoTo ERRH
9190          End If
9200        Else
9210  On Error GoTo ERRH
9220        End If
9230      Else
9240  On Error GoTo ERRH
9250      End If
9260    End If

9270    With wrk

9280      For lngX = 0& To (lngDbs - 1&)
9290        blnOptional = False
9300        If arr_varDb(D_DNAM, lngX) = "TAJrnTmp.mdb" Then blnOptional = True
9310        If blnOptional = True Or arr_varDb(D_DNAM, lngX) <> Parse_File(CurrentDb.Name) Then   ' ** Function: Above.
              ' ** Skip this database.

9320          If blnIsDemo = False Or (blnIsDemo = True And arr_varDb(D_DNAM, lngX) <> gstrFile_RePostDataName) Then
                ' ** If it's the Demo, skip the Journal reposting database.

9330            If blnOptional = False Then
9340              strConn = arr_varDb(D_PATH, lngX) & LNK_SEP & arr_varDb(D_DNAM, lngX)
9350            End If

9360            Set dbs = .OpenDatabase(strConn, False, True)  ' ** {pathfile}, {exclusive}, {read-only}

9370            With dbs
9380              For Each tdf In .TableDefs
9390                With tdf
9400                  For Each prp In .Properties
9410                    With prp
9420                      If .Name = "Description" Then
9430                        If .Value <> vbNullString Then
9440                          lngTbls = lngTbls + 1&
9450                          lngE = lngTbls - 1&
9460                          ReDim Preserve arr_varTbl(T_ELEMS, lngE)
9470                          arr_varTbl(T_MDB, lngE) = Parse_File(strConn)  ' ** Function: Above.
9480                          arr_varTbl(T_TNAM, lngE) = tdf.Name
9490                          arr_varTbl(T_DSC, lngE) = .Value
9500                          Exit For
9510                        End If
9520                      End If
9530                    End With
9540                  Next
9550                End With
9560              Next
9570              .Close
9580            End With

9590          End If  ' ** blnIsDemo.

9600        End If
9610        If blnOptional = True Then
9620          Exit For
9630        End If
9640      Next

9650      .Close
9660    End With

9670    If lngTbls > 0& Then
9680      Set dbs = CurrentDb
9690      With dbs
9700        For Each tdf In .TableDefs
9710          With tdf
9720            If .Connect <> vbNullString Then
9730              For lngX = 0& To (lngTbls - 1&)
9740                If InStr(.Connect, arr_varTbl(T_MDB, lngX)) > 0 Then
9750                  If .Name <> "LedgerArchive" And .Name <> "tblDataTypeDb1" Then
9760                    If arr_varTbl(T_TNAM, lngX) = .Name Then
9770                      blnFound = False
9780                      For Each prp In .Properties
9790                        With prp
9800                          If .Name = "Description" Then
9810                            blnFound = True
9820                            If tdf.Name = "tblDataTypeDb" Then
9830                              strTmp01 = arr_varTbl(T_DSC, lngX)
9840                              strTmp01 = Left(strTmp01, (Len(strTmp01) - 1)) & " (TrustAux.mdb)."
9850                            Else
9860                              strTmp01 = arr_varTbl(T_DSC, lngX)
9870                            End If
9880                            .Value = strTmp01
9890                            Exit For
9900                          End If
9910                        End With
9920                      Next
9930                      If blnFound = False Then
9940                        Set prp = .CreateProperty("Description", dbText, arr_varTbl(T_DSC, lngX))
9950                        .Properties.Append prp
9960                      End If
9970                    End If
9980                  Else
9990                    If (arr_varTbl(T_TNAM, lngX) = "ledger" And arr_varTbl(T_MDB, lngX) = "TrstArch.mdb") Then
10000                     blnFound = False
10010                     For Each prp In .Properties
10020                       With prp
10030                         If .Name = "Description" Then
10040                           blnFound = True
10050                           .Value = arr_varTbl(T_DSC, lngX)
10060                           Exit For
10070                         End If
10080                       End With
10090                     Next
10100                     If blnFound = False Then
10110                       Set prp = .CreateProperty("Description", dbText, arr_varTbl(T_DSC, lngX))
10120                       .Properties.Append prp
10130                     End If
10140                   ElseIf (arr_varTbl(T_TNAM, lngX) = "tblDataTypeDb" And arr_varTbl(T_MDB, lngX) = "TrustDta.mdb") Then
10150                     blnFound = False
10160                     For Each prp In .Properties
10170                       With prp
10180                         If .Name = "Description" Then
10190                           blnFound = True
10200                           strTmp01 = arr_varTbl(T_DSC, lngX)
10210                           strTmp01 = Left(strTmp01, (Len(strTmp01) - 1)) & " (TrustDta.mdb)."
10220                           .Value = strTmp01
10230                           Exit For
10240                         End If
10250                       End With
10260                     Next
10270                     If blnFound = False Then
10280                       Set prp = .CreateProperty("Description", dbText, arr_varTbl(T_DSC, lngX))
10290                       .Properties.Append prp
10300                     End If
10310                   End If
10320                 End If
10330               End If
10340             Next
10350           End If
10360         End With
10370       Next
10380       .TableDefs.Refresh
10390       .Close
10400     End With
10410   End If

EXITP:
10420   Set prp = Nothing
10430   Set tdf = Nothing
10440   Set dbs = Nothing
10450   Set wrk = Nothing
10460   TableDescriptionUpdate = blnRetVal
10470   Exit Function

ERRH:
10480   blnRetVal = False
10490   Select Case ERR.Number
        Case Else
10500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10510   End Select
10520   Resume EXITP

End Function

Public Function TableDescriptionUpdate2() As Boolean
' ** The purpose of this one is to copy the record counts I put
' ** in the Description here to their backend Description,
' ** otherwise everytime I run above, those numbers get wiped out.

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "TableDescriptionUpdate2"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, tdf As DAO.TableDef, prp As Object ', rst As DAO.Recordset
        Dim lngTbls As Long, arr_varTbl() As Variant
        Dim blnFound As Boolean, blnIrregular As Boolean
        Dim intLen As Integer
        Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String
        Dim lngX As Long, intY As Integer, lngE As Long
        Dim blnRetVal As Boolean

        Const T_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const T_NAM   As Integer = 0
        Const T_DBNAM As Integer = 1
        Const T_PATH  As Integer = 2
        Const T_CNT   As Integer = 3
        Const T_IRR   As Integer = 4

10610   blnRetVal = True

10620   lngTbls = 0&
10630   ReDim arr_varTbl(T_ELEMS, 0)

10640   Set dbs = CurrentDb
10650   With dbs
10660     For Each tdf In .TableDefs
10670       With tdf
10680         If Left(.Name, 4) <> "MSys" Then  ' ** Skip those pesky system tables.
10690           If .Connect <> vbNullString Then
10700             For Each prp In .Properties
10710               With prp
10720                 If .Name = "Description" Then
10730                   varTmp00 = .Value
10740                   If IsNull(varTmp00) = False Then
10750                     If Trim(varTmp00) <> vbNullString Then
10760                       strTmp01 = varTmp00
10770                       strTmp02 = vbNullString
10780                       intLen = Len(strTmp01)
10790                       For intY = intLen To 1 Step -1
10800                         If Mid(strTmp01, intY, 1) = ";" Then
10810                           strTmp02 = Trim(Mid(strTmp01, (intY + 1)))
10820                           Exit For
10830                         End If
10840                       Next
10850                       If strTmp02 <> vbNullString Then
                              ' ** System Doc: Table Relationships; 119.
10860                         blnIrregular = False
10870                         If Asc(Left(strTmp02, 1)) >= 48 And Asc(Left(strTmp02, 1)) <= 57 Then  ' ** 0 - 9.
10880                           If InStr(strTmp02, " ") > 0 Then
10890                             blnIrregular = True
10900                             strTmp02 = Trim(Left(strTmp02, InStr(strTmp02, " ")))
10910                           End If
10920                           If Right(strTmp02, 1) = "." Then
10930                             strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
10940                           End If
                                ' ** Are they all digits?
10950                           intLen = Len(strTmp02)
10960                           blnFound = False  ' ** No non-numeric characters.
10970                           For intY = 1 To intLen
10980                             If Asc(Mid(strTmp02, intY, 1)) < 48 Or Asc(Mid(strTmp02, intY, 1)) > 57 Then
10990                               blnFound = True
11000                               Exit For
11010                             End If
11020                           Next
11030                           If blnFound = False Then
                                  ' ** OK, it's a number!
11040                             lngTbls = lngTbls + 1&
11050                             lngE = lngTbls - 1&
11060                             ReDim Preserve arr_varTbl(T_ELEMS, lngE)
11070                             arr_varTbl(T_NAM, lngE) = tdf.Name
11080                             strTmp01 = tdf.Connect
11090                             intY = InStr(strTmp01, LNK_IDENT)
11100                             strTmp01 = Mid(strTmp01, (intY + Len(LNK_IDENT)))
11110                             arr_varTbl(T_DBNAM, lngE) = Parse_File(strTmp01)  ' ** Function: Above.
11120                             arr_varTbl(T_PATH, lngE) = Parse_Path(strTmp01)  ' ** Function: Above.
11130                             arr_varTbl(T_CNT, lngE) = strTmp02
11140                             arr_varTbl(T_IRR, lngE) = blnIrregular
11150                           End If
11160                         End If
11170                       End If
11180                     End If
11190                   End If
11200                 End If
11210               End With
11220             Next
11230           End If
11240         End If
11250       End With
11260     Next
11270     .Close
11280   End With

11290   If lngTbls > 0& Then

11300 On Error Resume Next
11310     Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
11320     If ERR.Number <> 0 Then
11330 On Error GoTo ERRH
11340 On Error Resume Next
11350       Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
11360       If ERR.Number <> 0 Then
11370 On Error GoTo ERRH
11380 On Error Resume Next
11390         Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
11400         If ERR.Number <> 0 Then
11410 On Error GoTo ERRH
11420 On Error Resume Next
11430           Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
11440           If ERR.Number <> 0 Then
11450 On Error GoTo ERRH
11460 On Error Resume Next
11470             Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
11480             If ERR.Number <> 0 Then
11490 On Error GoTo ERRH
11500 On Error Resume Next
11510               Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
11520               If ERR.Number <> 0 Then
11530 On Error GoTo ERRH
11540 On Error Resume Next
11550                 Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
11560 On Error GoTo ERRH
11570               Else
11580 On Error GoTo ERRH
11590               End If
11600             Else
11610 On Error GoTo ERRH
11620             End If
11630           Else
11640 On Error GoTo ERRH
11650           End If
11660         Else
11670 On Error GoTo ERRH
11680         End If
11690       Else
11700 On Error GoTo ERRH
11710       End If
11720     Else
11730 On Error GoTo ERRH
11740     End If

11750     With wrk
11760       For lngX = 0& To (lngTbls - 1&)
11770         strTmp01 = arr_varTbl(T_PATH, lngX) & LNK_SEP & arr_varTbl(T_DBNAM, lngX)
11780         Set dbs = .OpenDatabase(strTmp01, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
11790         With dbs
11800           Set tdf = .TableDefs(arr_varTbl(T_NAM, lngX))
11810           With tdf
11820             strTmp02 = Trim(.Properties("Description"))  ' ** If it's got a Description here, it should have one there.
11830             intLen = Len(strTmp02)
11840             For intY = intLen To 1 Step -1
11850               If Mid(strTmp02, intY, 1) = ";" Then
11860                 If arr_varTbl(T_IRR, lngX) = False Then
11870                   strTmp02 = Left(strTmp02, intY) & " " & arr_varTbl(T_CNT, lngX) & "."
11880                 Else
11890                   strTmp02 = Left(strTmp02, intY) & " " & arr_varTbl(T_CNT, lngX) & Mid(strTmp02, InStr((intY + 2), strTmp02, " "))
11900                 End If
11910                 Exit For
11920               End If
11930             Next
11940             .Properties("Description") = strTmp02
11950           End With
11960           .Close
11970         End With
11980       Next
11990       .Close
12000     End With

12010   End If

EXITP:
12020   Set prp = Nothing
12030   Set tdf = Nothing
        'Set rst = Nothing
12040   Set dbs = Nothing
12050   Set wrk = Nothing
12060   TableDescriptionUpdate2 = blnRetVal
12070   Exit Function

ERRH:
12080   blnRetVal = False
12090   Select Case ERR.Number
        Case Else
12100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12110   End Select
12120   Resume EXITP

End Function

Public Function MakeTempTable(rstTemplate As DAO.Recordset, strTableName As String) As Boolean
' ** Tables created, with procedures called by:
'frmRpt_CourtReports_CA.BuildAssetListInfo()
'  MakeTempTable(rst, "tmpAssetList2") : qryCourtReport_CA_07z
'  MakeTempTable(rst, "tmpAssetList2") : qryCourtReport_CA_07z
'  MakeTempTable(rst, "tmpAccountInfo") : qryCourtReport_CA_07z (Yes, they're the same!)
'frmRpt_CourtReports_FL.BuildAssetListInfo()
'  MakeTempTable(rst, "tmpAssetList2") : qryCourtReport_FL_07z
'  MakeTempTable(rst, "tmpAssetList2") : qryCourtReport_FL_07z
'  MakeTempTable(rst, "tmpAccountInfo") : qryCourtReport_FL_07z
'frmRpt_CourtReports_NS.BuildAssetListInfo()
'  MakeTempTable(rst, "tmpAssetList2") : qryCourtReport_NS_07z
'  MakeTempTable(rst, "tmpAssetList2") : qryCourtReport_NS_07z
'  MakeTempTable(rst, "tmpAccountInfo") : qryCourtReport_NS_07z
'modIncExpFuncs.IncomeExpense_BuildTable()
'  MakeTempTable(rst, "tmpIncomeExpenseReports") : qryIncomeExpenseReports_02/qryIncomeExpenseReports_04

        'modIncExpFuncs.IncomeExpense_BuildTable()
        '  MakeTempTable(rst, "tmpIncomeExpenseReports")

        'Make-Table queries use SELECT ... INTO {table name}
        'Append queries use INSERT INTO ...

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "MakeTempTable"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim tdf As DAO.TableDef
        Dim lngX As Long
        Dim blnRetVal As Boolean

12210   blnRetVal = True

12220   If TableExists(strTableName) = True Then  ' ** Function: Above.
12230     Set dbs = CurrentDb
12240     With dbs
12250       Select Case strTableName
            Case "tmpAssetList2"
              ' ** Empty tmpAssetList2.
12260         Set qdf = .QueryDefs("qryFileUtilities_01")
12270       Case "tmpAssetList5"
              ' ** Empty tmpAssetList5.
12280         Set qdf = .QueryDefs("qryFileUtilities_04")
12290       Case "tmpAccountInfo"
              ' ** Empty tmpAccountInfo.
12300         Set qdf = .QueryDefs("qryFileUtilities_02")
12310       Case "rstTmpAccountInfo2"
              ' ** Empty tmpAccountInfo2.
12320         Set qdf = .QueryDefs("qryFileUtilities_05")
12330       Case "tmpIncomeExpenseReports"
              ' ** Empty tmpIncomeExpenseReports.
12340         Set qdf = .QueryDefs("qryFileUtilities_03")
12350       End Select
12360       qdf.Execute
12370       Set qdf = Nothing
12380       .Close
12390     End With
12400     Set dbs = Nothing
12410   Else

12420 On Error Resume Next
12430     Set tdf = CurrentDb.CreateTableDef(strTableName)
12440     If ERR.Number <> 0 Then
12450       If ERR.Number = 3010 Then  ' ** Table '|' already exists.
12460         If Reports.Count > 0 Then
12470           Do While Reports.Count > 0
12480 On Error GoTo ERRH
12490             DoCmd.Close acReport, Reports(0).Name
12500           Loop
12510           Set tdf = CurrentDb.CreateTableDef(strTableName)
12520         Else
12530           zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12540 On Error GoTo ERRH
12550           blnRetVal = False
12560         End If
12570       Else
12580         zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12590 On Error GoTo ERRH
12600         blnRetVal = False
12610       End If
12620     Else
12630 On Error GoTo ERRH
12640     End If
12650     For lngX = 0& To (rstTemplate.Fields.Count - 1&)
12660       With tdf
12670         .Fields.Append .CreateField(rstTemplate.Fields(lngX).Name, rstTemplate.Fields(lngX).Type, rstTemplate.Fields(lngX).Size)
12680         If .Fields(lngX).Type = dbText Then
12690           .Fields(lngX).AllowZeroLength = True
12700         End If
12710       End With
12720     Next

12730     CurrentDb.TableDefs.Append tdf

12740   End If

EXITP:
12750   Set tdf = Nothing
12760   Set qdf = Nothing
12770   Set dbs = Nothing
12780   MakeTempTable = blnRetVal
12790   Exit Function

ERRH:
12800   blnRetVal = False
12810   Select Case ERR.Number
        Case 3010  ' ** Table '|' already exists.
12820     Beep
12830     MsgBox "You may have a report open that is preventing a temporary table from being re-created.", _
            vbCritical + vbOKOnly, "Error Creating Temporary Table"
12840   Case Else
12850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12860   End Select
12870   Resume EXITP

End Function

Public Function CopyToTempTable(rstFrom As DAO.Recordset, strTo As String) As Boolean
' ** NOTE: assumes recordset and table have congruent structures.

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "CopyToTempTable"

        Dim rstTo As DAO.Recordset
        Dim lngRecs As Long, lngFlds As Long
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

12910   blnRetVal = True  ' ** Unless proven otherwise.

12920   Set rstTo = CurrentDb.OpenRecordset(strTo, dbOpenDynaset)

12930   With rstFrom
12940     If .BOF = True And .EOF = True Then
            ' ** Do nothing.
12950     Else
12960       .MoveLast
12970       lngRecs = .RecordCount
12980       .MoveFirst
12990       lngFlds = .Fields.Count
13000       For lngX = 1& To lngRecs
13010         rstTo.AddNew
13020         For lngY = 0& To (lngFlds - 1&)
13030           rstTo.Fields(lngY) = .Fields(lngY)
13040         Next
13050         rstTo.Update
13060         If lngX < lngRecs Then .MoveNext
13070       Next
13080     End If
13090   End With
13100   rstTo.Close

EXITP:
13110   Set rstTo = Nothing
13120   CopyToTempTable = blnRetVal
13130   Exit Function

ERRH:
13140   blnRetVal = False
13150   Select Case ERR.Number
        Case Else
13160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13170   End Select
13180   Resume EXITP

End Function

Public Function CompareTableStructure(strTableAName As String, strTableBName As String) As String
' ** Given 2 table names, compare their structures.
' ** Returns: "SAME" if same, "DIFFERENT" if not, RET_ERR if errors out.
' ** Compares table structures based on field names instead of position in the table.
' ** Also, it logs the archive process to the \Database directory.  Process resets the log everytime.
' ** Called by:
' **   frmArchiveTransactions
' **     cmdArchive_Click()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "CompareTableStructure"

        Dim rstTableA As DAO.Recordset
        Dim rstTableB As DAO.Recordset
        Dim strSameStructure As String
        Dim strFieldName As String
        Dim strTableName As String
        Dim blnContinue As Boolean
        Dim intLoop As Integer
        Dim lngX As Long, lngFlds As Long, blnFound As Boolean

13210   zErrorLogWriter ("")
13220   zErrorLogWriter ("===================================================================================")
13230   zErrorLogWriter ("Start Archive Process " & Format(Now(), "mm/dd/yy hh:mm:ss:dd"))
13240   zErrorLogWriter ("Table A = " & strTableAName)
13250   zErrorLogWriter ("Table B = " & strTableBName)

13260   strSameStructure = "SAME"  ' ** Unless proven otherwise.
13270   blnContinue = True

13280   Set rstTableA = CurrentDb.OpenRecordset("SELECT TOP 1 * FROM [" & strTableAName & "]", dbOpenSnapshot)
13290   Set rstTableB = CurrentDb.OpenRecordset("SELECT TOP 1 * FROM [" & strTableBName & "]", dbOpenSnapshot)

        ' ** Both Records should have the same number of fields.
13300   If rstTableA.Fields.Count <> rstTableB.Fields.Count Then
13310     blnContinue = False
13320     strSameStructure = "DIFFERENT"
13330     zErrorLogWriter (strTableAName _
            & " FieldCount(" & rstTableA.Fields.Count & ")" _
            & " / " & strTableBName _
            & " FieldCount(" & rstTableB.Fields.Count & ")")
13340   Else

          ' ** Compare table A to Table B on a field by field basis.
          ' ** Use the name of the field instead of the index to allow for tables
          ' ** defined the same BUT the fields are in a different order.

          ' ** There may be TrstArch.mdb's out there with the Ledger field revcode_KB!
          ' ** Check for that and advise the user.
13350     blnFound = False
13360     If strTableAName = "ledger" And strTableBName = "LedgerArchive" Then
13370       lngFlds = rstTableB.Fields.Count
13380       For lngX = 0& To (lngFlds - 1&)
13390         If rstTableB.Fields(lngX).Name = "revcode_KD" Then
13400           blnFound = True
13410           Exit For
13420         End If
13430       Next
13440       If blnFound = True Then
13450         blnContinue = False
13460         strSameStructure = "DIFFERENT: revcode_KD"
13470         zErrorLogWriter (strTableAName & " FieldName(revcode_ID)" & " / " & strTableBName & " FieldName(revcode_KD)")
13480       End If
13490     End If

13500     If blnContinue = True Then

13510       intLoop = 0
13520       strTableName = strTableBName  ' ** Save name for error messages.
13530       Do While (intLoop < rstTableA.Fields.Count) And blnContinue = True

13540         strFieldName = rstTableA.Fields(intLoop).Name

13550         If rstTableA.Fields(strFieldName).Type <> rstTableB.Fields(strFieldName).Type Then
13560           blnContinue = False
13570           strSameStructure = "DIFFERENT"
13580           zErrorLogWriter ("Index " & intLoop & rstTableA.Fields(strFieldName).Name _
                  & " Type(" & rstTableA.Fields(strFieldName).Type & ")" _
                  & " / " & rstTableB.Fields(strFieldName).Name _
                  & " Type(" & rstTableB.Fields(strFieldName).Type & ")")
13590         End If

13600         If rstTableA.Fields(strFieldName).Size <> rstTableB.Fields(strFieldName).Size Then
13610           blnContinue = False
13620           strSameStructure = "DIFFERENT"
13630           zErrorLogWriter ("Index " & intLoop & rstTableA.Fields(strFieldName).Name _
                  & " Size(" & rstTableA.Fields(strFieldName).Size & ")" _
                  & " / " & rstTableB.Fields(strFieldName).Name _
                  & " Size(" & rstTableB.Fields(strFieldName).Size & ")")
13640         End If
13650         intLoop = intLoop + 1

13660       Loop

            ' ** Compare table B to Table A on a field by field basis to make
            ' ** ABSOLUTELY sure all the fields names match.
13670       strTableName = strTableAName  ' ** Save name for error messages.
13680       intLoop = 0

13690       Do While (intLoop < rstTableA.Fields.Count) And blnContinue = True

13700         strFieldName = rstTableA.Fields(intLoop).Name

13710         If rstTableA.Fields(strFieldName).Type <> rstTableB.Fields(strFieldName).Type Then
13720           blnContinue = False
13730           strSameStructure = "DIFFERENT"
13740           zErrorLogWriter ("Index " & intLoop & rstTableA.Fields(strFieldName).Name _
                  & " Type(" & rstTableA.Fields(strFieldName).Type & ")" _
                  & " / " & rstTableB.Fields(strFieldName).Name _
                  & " Type(" & rstTableB.Fields(strFieldName).Type & ")")
13750         End If
13760         If rstTableA.Fields(strFieldName).Size <> rstTableB.Fields(strFieldName).Size Then
13770           blnContinue = False
13780           strSameStructure = "DIFFERENT"
13790           zErrorLogWriter ("Index " & intLoop & rstTableA.Fields(strFieldName).Name _
                  & " Size(" & rstTableA.Fields(strFieldName).Size & ")" _
                  & " / " & rstTableB.Fields(strFieldName).Name _
                  & " Size(" & rstTableB.Fields(strFieldName).Size & ")")
13800         End If
13810         intLoop = intLoop + 1
13820       Loop

13830       If strSameStructure = "DIFFERENT" Then
13840         blnContinue = False
13850       End If

13860     End If

13870   End If

13880   rstTableA.Close
13890   rstTableB.Close

EXITP:
13900   Set rstTableA = Nothing
13910   Set rstTableB = Nothing
13920   CompareTableStructure = strSameStructure
13930   Exit Function

ERRH:
13940   strSameStructure = RET_ERR
13950   Select Case ERR.Number
        Case 3265 ' ** Item not found in this collection.
13960     zErrorLogWriter ("Index " & intLoop & " " & strFieldName & " not in table " & strTableName)
13970   Case Else
13980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13990   End Select
14000   Resume EXITP

End Function

Public Function GetUNCPath(strDriveLetter As String) As String

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "GetUNCPath"

        Dim strMsg As String, lngRetVal As Long
        Dim strLocalName As String
        Dim strRemoteName As String
        Dim lngRemoteName As Long
        Dim strRetVal As String

14110   strMsg = vbNullString: strRetVal = vbNullString

14120   strLocalName = strDriveLetter
14130   strRemoteName = String$(255, Chr$(32))
14140   lngRemoteName = Len(strRemoteName)
14150   lngRetVal = WNetGetConnection(strLocalName, strRemoteName, lngRemoteName)

14160   Select Case lngRetVal
        Case ERROR_BAD_DEVICE
14170     strMsg = "Error: Bad Device"
14180   Case ERROR_CONNECTION_UNAVAIL
14190     strMsg = "Error: Connection Unavailable"
14200   Case ERROR_EXTENDED_ERROR
14210     strMsg = "Error: Extended Error"
14220   Case ERROR_MORE_DATA
14230     strMsg = "Error: More Data"
14240   Case ERROR_NOT_SUPPORTED
14250     strMsg = "Error: Feature not Supported"
14260   Case ERROR_NO_NET_OR_BAD_PATH
14270     strMsg = "Error: No Network Available or Bad Path"
14280   Case ERROR_NO_NETWORK
14290     strMsg = "Error: No Network Available"
14300   Case ERROR_NOT_CONNECTED
14310     strMsg = "Error: Not Connected"
14320   Case ERROR_SUCCESS
          ' ** All is successful.
14330   End Select

14340   If strMsg <> vbNullString Then
14350     MsgBox strMsg, vbExclamation + vbOKOnly, "UNC Path Not Found"
14360   Else
          ' ** Display the path in a Message box or return the UNC through the function.
14370     MsgBox Left(strRemoteName, lngRemoteName), vbInformation + vbOKOnly, "UNC Path"
14380     strRetVal = Left(strRemoteName, lngRemoteName)
14390   End If

EXITP:
14400   GetUNCPath = strRetVal
14410   Exit Function

ERRH:
14420   strRetVal = "ERROR#"
14430   Select Case ERR.Number
        Case Else
14440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14450   End Select
14460   Resume EXITP

End Function

Public Function GetUserName() As String
' ** Set or retrieve the name of the User who logged into the computer.

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "GetUserName"

        Dim strBuffer As String
        Dim lngLen As Long
        Dim strRetVal As String

14510   strRetVal = vbNullString

14520   strBuffer = Space(255 + 1)
14530   lngLen = Len(strBuffer)
14540   If CBool(GetUserNameAPI(strBuffer, lngLen)) Then
14550     strRetVal = Left(strBuffer, lngLen - 1)
14560     If strRetVal = "Administrator" Then
            ' ** Most likely one of the developers.
14570     End If
14580   End If

EXITP:
14590   GetUserName = strRetVal
14600   Exit Function

ERRH:
14610   strRetVal = vbNullString
14620   Select Case ERR.Number
        Case Else
14630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14640   End Select
14650   Resume EXITP

End Function

Public Function GetComputerName() As String
' ** Set or retrieve the name of the computer.

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "GetComputerName"

        Dim strBuffer As String
        Dim lngLen As Long
        Dim strRetVal As String

14710   strRetVal = vbNullString
14720   strBuffer = Space(255 + 1)
14730   lngLen = Len(strBuffer)
14740   If CBool(GetComputerNameAPI(strBuffer, lngLen)) Then
14750     strRetVal = Left(strBuffer, lngLen)
14760   End If

EXITP:
14770   GetComputerName = strRetVal
14780   Exit Function

ERRH:
14790   strRetVal = vbNullString
14800   Select Case ERR.Number
        Case Else
14810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14820   End Select
14830   Resume EXITP

End Function

Public Function GetDefaultPrinter() As String
' ** Retrieve the name of the default printer.

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "GetDefaultPrinter"

        Dim strBuffer As String
        Dim lngLen As Long
        Dim strRetVal As String

14910   strRetVal = vbNullString

14920   strBuffer = Space(255 + 1)
14930   lngLen = Len(strBuffer)
14940   If CBool(GetDefaultPrinterAPI(strBuffer, lngLen)) Then
14950     strRetVal = Left(strBuffer, lngLen - 1)
14960   End If

EXITP:
14970   GetDefaultPrinter = strRetVal
14980   Exit Function

ERRH:
14990   strRetVal = vbNullString
15000   Select Case ERR.Number
        Case Else
15010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15020   End Select
15030   Resume EXITP

End Function

Public Function GetDefaultUser() As String
' ** Return the default administrative user name for this version.

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "GetDefaultUser"

        Dim strCurrApp As String
        Dim strRetVal As String

15110   strRetVal = vbNullString

        ' ** Should this make use of tblSecurity_User somehow?
15120   strCurrApp = CurrentAppName  ' ** Function: Above.
15130   strCurrApp = Left(strCurrApp, (InStr(strCurrApp, ".") - 1))  ' ** Just the name, no extension.
15140   Select Case strCurrApp
        Case gstrFile_App
15150     If Len(TA_SEC2) > Len(TA_SEC) Then
15160       strRetVal = "TAAdmin"
15170     Else
15180       strRetVal = "TADemo"
15190     End If
15200   Case "TrustImport"
15210     strRetVal = "TAAdmin"
15220   Case Else
15230     strRetVal = "TAAdmin"
15240   End Select

EXITP:
15250   GetDefaultUser = strRetVal
15260   Exit Function

ERRH:
15270   strRetVal = vbNullString
15280   Select Case ERR.Number
        Case Else
15290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15300   End Select
15310   Resume EXITP

End Function

Public Function Parse_Path(varInput As Variant) As Variant
' ** Return the path portion of a complete file-with-path string, WITHOUT FINAL BACKSLASH!
' ** Returns everything to the left of final backslash.

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "Parse_Path"

        Dim intPos01 As Integer, intLen As Integer
        Dim intX As Integer
        Dim varRetVal As Variant

15410   varRetVal = Null

15420   If IsNull(varInput) = False Then
15430     If varInput <> vbNullString Then
            ' ** C:\VictorGCS_Clients\TrustAccountant\NewWorking\Trust.mdb  '## OK
15440       intPos01 = InStr(varInput, LNK_SEP)
15450       If intPos01 > 0 Then
15460         intLen = Len(varInput)
15470         For intX = intLen To 1 Step -1
15480           If Mid(varInput, intX, 1) = LNK_SEP Then
15490             varRetVal = Left(varInput, (intX - 1))
15500             Exit For
15510           End If
15520         Next
15530       End If
15540     End If
15550   End If

EXITP:
15560   Parse_Path = varRetVal
15570   Exit Function

ERRH:
15580   varRetVal = Null
15590   Select Case ERR.Number
        Case Else
15600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15610   End Select
15620   Resume EXITP

End Function

Public Function Parse_File(varInput As Variant) As String
' ** Return the file name from a path and file name string.
' ** Returns everything to the right of final backslash.

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "Parse_File"

        Dim intLen As Integer
        Dim strInput As String
        Dim intX As Integer
        Dim strRetVal As String

15710   strRetVal = vbNullString

15720   If IsNull(varInput) = False Then
15730     strInput = CStr(varInput)
15740     If strInput <> vbNullString Then
15750       intLen = Len(strInput)
15760       If InStr(strInput, LNK_SEP) > 0 Then
              ' ** For now, assume the separator is a backslash.
15770         If Right(strInput, 1) <> LNK_SEP Then
                ' ** If it ends in a backslash, it's a directory with no file name.
15780           For intX = intLen To 1 Step -1
                  ' ** Find the last backslash in the string.
15790             If Mid(strInput, intX, 1) = LNK_SEP Then
15800               strRetVal = Mid(strInput, (intX + 1))
15810               Exit For
15820             End If
15830           Next
15840         End If
15850       End If
15860     End If
15870   End If

EXITP:
15880   Parse_File = strRetVal
15890   Exit Function

ERRH:
15900   strRetVal = vbNullString
15910   Select Case ERR.Number
        Case Else
15920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15930   End Select
15940   Resume EXITP

End Function

Public Function Parse_Ext(varInput As Variant, Optional varFirst As Variant) As String
' ** Return the file extension from a path and file name string.
' ** By default, returns everything after FIRST period.
' ** Sending varFirst = False returns everything after LAST period.

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "Parse_Ext"

        Dim strInput As String
        Dim blnFirst As Boolean
        Dim intPos01 As Integer, intLen As Integer
        Dim intX As Integer
        Dim strRetVal As String

16010   strRetVal = vbNullString

16020   If IsMissing(varFirst) = True Then
16030     blnFirst = True
16040   Else
16050     blnFirst = varFirst
16060   End If

16070   If IsNull(varInput) = False Then
16080     strInput = CStr(varInput)
16090     If strInput <> vbNullString Then
16100       intPos01 = InStr(strInput, LNK_SEP)
16110       If intPos01 > 0 Then
16120         strInput = Parse_File(strInput)  ' ** Function: Above.
16130       End If
16140       If strInput <> vbNullString Then
16150         If blnFirst = True Then
16160           intPos01 = InStr(strInput, ".")
16170           If intPos01 > 0 And intPos01 <> Len(strInput) Then
                  ' ** Take it from the first period, rather than the last (in case there's more than one).
                  ' ** The last character shouldn't be a period, but you never know.
16180             strRetVal = Mid(strInput, (intPos01 + 1))
16190           End If
16200         Else
16210           intLen = Len(strInput)
16220           For intX = intLen To 1 Step -1
16230             If Mid(strInput, intX, 1) = "." Then
16240               If intX < intLen Then
                      ' ** The last character shouldn't be a period, but you never know.
16250                 strRetVal = Mid(strInput, (intX + 1))
16260               End If
16270               Exit For
16280             End If
16290           Next
16300         End If
16310       End If
16320     End If
16330   End If

EXITP:
16340   Parse_Ext = strRetVal
16350   Exit Function

ERRH:
16360   strRetVal = vbNullString
16370   Select Case ERR.Number
        Case Else
16380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16390   End Select
16400   Resume EXITP

End Function

Public Function Rem_Ext(varInput As Variant) As String
' ** Remove the file extension from a full file name.

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Ext"

        Dim intLen As Integer
        Dim intPos01 As Integer
        Dim strRetVal As String

16510   strRetVal = vbNullString

16520   If IsNull(varInput) = False Then
16530     If Trim(varInput) <> vbNullString Then
16540       strRetVal = Trim(varInput)
16550       intLen = Len(strRetVal)
16560       intPos01 = InStr(strRetVal, ".")
16570       If intPos01 > 1 And intPos01 <> Len(strRetVal) Then
16580         strRetVal = Left(strRetVal, (intPos01 - 1))
16590       Else
16600         strRetVal = vbNullString
16610       End If
16620     End If
16630   End If

EXITP:
16640   Rem_Ext = strRetVal
16650   Exit Function

ERRH:
16660   strRetVal = vbNullString
16670   Select Case ERR.Number
        Case Else
16680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16690   End Select
16700   Resume EXITP

End Function

Public Function ConvertRTFtoTXT(strPathFile As String) As String
' ** Convert an EstateVal RTF (Rich Text Format) file to a plain text file.
' ** This removes all the formatting characters.

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "ConvertRTFtoTXT"

        Dim strFileNoExt As String, strFileTxt As String
        Dim strInput As String, strOutput As String
        Dim intPos01 As Integer
        Dim strRetVal As String

16810   strRetVal = vbNullString
16820   If strPathFile <> vbNullString Then
16830     If Dir(strPathFile) <> vbNullString Then
16840       If Right(strPathFile, 4) = ".rtf" Then

16850         intPos01 = InStr(strPathFile, ".rtf")
16860         strFileNoExt = Left(strPathFile, (intPos01 - 1))
16870         strFileTxt = strFileNoExt & "_conv.txt"
16880         If Dir(strFileTxt) <> vbNullString Then
16890           Kill (strFileTxt)
16900         End If
16910         strRetVal = strFileTxt

16920         Open strPathFile For Input As #1
16930         Open strFileTxt For Binary As #2

16940         Do While Not EOF(1)

16950           Line Input #1, strInput
16960           If Left(strInput, 1) <> "{" And Left(strInput, 1) <> "}" Then

16970             If Left(strInput, 7) = "\par{}" Then
                    ' ** Blank line.
16980               Put #2, , vbCrLf
16990             Else
17000               intPos01 = InStr(strInput, "{")  ' ** \par{
17010               If intPos01 > 0 Then
17020                 strOutput = Mid(strInput, (intPos01 + 1))
17030               Else
17040                 strOutput = strInput
17050               End If
17060               intPos01 = InStr(strOutput, "}")
17070               If intPos01 > 0 Then
17080                 strOutput = Left(strOutput, (intPos01 - 1))
17090               End If
17100               Put #2, , strOutput & vbCrLf
17110             End If
17120           End If

17130         Loop

17140         Close #1
17150         Close #2

17160       End If
17170     End If
17180   End If

EXITP:
17190   ConvertRTFtoTXT = strRetVal
17200   Exit Function

ERRH:
17210   strRetVal = RET_ERR
17220   If FreeFile > 1 Then Close #2
17230   If FreeFile > 0 Then Close #1
17240   Select Case ERR.Number
        Case Else
17250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17260   End Select
17270   Resume EXITP

End Function

Public Function GetTypeDrive(varDriveLetter As Variant) As Long
' ** Named so as not to confuse it with the actual API function, GetDriveType(), above.
' ** varDriveLetter should be LETTER ONLY, NO COLON OR BACKSLASH.

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "GetTypeDrive"

        Dim lngRetVal As Long

17310   lngRetVal = 0&

        ' ** Returns a vbxDriveType long integer (Note: These Public constants are mine and not part of VB).
17320   lngRetVal = GetDriveType(varDriveLetter & ":\")

        ' ** VbxDriveType enumeration:
        ' **   0  vbxDriveUnknown    DRIVE_UNKNOWN      The drive type cannot be determined.
        ' **   1  vbxDriveNoRootDir  DRIVE_NO_ROOT_DIR  The root path is invalid; for example, there is no volume mounted at the specified path.
        ' **   2  vbxDriveRemovable  DRIVE_REMOVABLE    The drive has removable media; for example, a floppy drive, thumb drive, or flash card reader.
        ' **   3  vbxDriveFixed      DRIVE_FIXED        The drive has fixed media; for example, a hard drive or flash drive.
        ' **   4  vbxDriveRemote     DRIVE_REMOTE       The drive is a remote (network) drive.
        ' **   5  vbxDriveCDROM      DRIVE_CDROM        The drive is a CD-ROM drive.
        ' **   6  vbxDriveRAMDisk    DRIVE_RAMDISK      The drive is a RAM disk.

EXITP:
17330   GetTypeDrive = lngRetVal
17340   Exit Function

ERRH:
17350   lngRetVal = 0&
17360   Select Case ERR.Number
        Case Else
17370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17380   End Select
17390   Resume EXITP

End Function

Public Function GetDriveStrings() As Variant
' ** Wrapper for calling the GetLogicalDriveStrings API.

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "GetDriveStrings"

        Dim lngResult As Long        ' ** Result of our API calls.
        Dim strDrives As String      ' ** String to pass to API call.
        Dim lngLenStrDrives As Long  ' ** Length of the above string.
        Dim lngDrives As Long
        Dim arr_varDrive() As Variant
        Dim lngPos01 As Long
        Dim strTmp01 As String

        ' ** Call GetLogicalDriveStrings with a buffer size of zero
        ' ** to find out how large our stringbuffer needs to be.
17410   lngResult = GetLogicalDriveStrings(0, strDrives)  ' ** API

17420   strDrives = String(lngResult, 0)
17430   lngLenStrDrives = lngResult

        ' ** Call again with our new buffer.
17440   lngResult = GetLogicalDriveStrings(lngLenStrDrives, strDrives)  ' ** API

17450   lngDrives = 0&
17460   ReDim arr_varDrive(0)

17470   If lngResult = 0 Then
          ' ** There was some error calling the API.
          ' ** Pass back an error.
17480     arr_varDrive(0) = RET_ERR
17490   Else
17500     lngPos01 = 1&
17510     Do While Not Mid(strDrives, lngPos01, 1) = Chr(0)
17520       strTmp01 = Mid(strDrives, lngPos01, 3)
17530       lngPos01 = lngPos01 + 4
17540       lngDrives = lngDrives + 1&
17550       ReDim Preserve arr_varDrive(lngDrives - 1&)
17560       arr_varDrive(lngDrives - 1&) = strTmp01
17570     Loop
17580   End If

EXITP:
17590   GetDriveStrings = arr_varDrive
17600   Exit Function

ERRH:
17610   arr_varDrive(0) = RET_ERR
17620   Select Case ERR.Number
        Case Else
17630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17640   End Select
17650   Resume EXITP

End Function

Public Function HexX(varInput As Variant) As Variant
' ** Convert Hex to a Long Integer; reverse of Hex() function.

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "HexX"

        Dim strInput As String
        Dim intPos01 As Integer, intLen As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
        Dim dblTmp05 As Double, dblTmp06 As Double, dblTmp07 As Double, dblTmp08 As Double
        Dim intX As Integer
        Dim varRetVal As Variant

17710   If IsNull(varInput) = False Then

17720     intLen = Len(varInput)
17730     dblTmp05 = 0&: dblTmp06 = 0&

17740     If intLen > 0 Then
17750       intPos01 = InStr(varInput, " Or ")
17760       If intPos01 > 0 Then
17770         varInput = Rem_Parens(varInput)  ' ** Module Function: modStringFuncs.
17780         intPos01 = InStr(varInput, " Or ")
17790         strTmp01 = Trim(Left(varInput, intPos01))
17800         strTmp02 = Trim(Mid(varInput, (intPos01 + 3)))
17810         intPos01 = InStr(strTmp02, " Or ")
17820         If intPos01 > 0 Then
17830           strTmp03 = Trim(Mid(strTmp02, (intPos01 + 3)))
17840           strTmp02 = Trim(Left(strTmp02, intPos01))
17850           intPos01 = InStr(strTmp03, " Or ")
17860           If intPos01 > 0 Then
17870             strTmp04 = Trim(Mid(strTmp03, (intPos01 + 3)))
17880             strTmp03 = Trim(Left(strTmp03, intPos01))
17890             dblTmp05 = Val(HexX(strTmp01))  ' ** Recursive.
17900             dblTmp06 = Val(HexX(strTmp02))  ' ** Recursive.
17910             dblTmp07 = Val(HexX(strTmp03))  ' ** Recursive.
17920             dblTmp08 = Val(HexX(strTmp04))  ' ** Recursive.
17930             varRetVal = (dblTmp05 Or dblTmp06 Or dblTmp07 Or dblTmp08)
17940           Else
17950             dblTmp05 = Val(HexX(strTmp01))  ' ** Recursive.
17960             dblTmp06 = Val(HexX(strTmp02))  ' ** Recursive.
17970             dblTmp07 = Val(HexX(strTmp03))  ' ** Recursive.
17980             varRetVal = (dblTmp05 Or dblTmp06 Or dblTmp07)
17990           End If
18000         Else
18010           dblTmp05 = Val(HexX(strTmp01))  ' ** Recursive.
18020           dblTmp06 = Val(HexX(strTmp02))  ' ** Recursive.
18030           varRetVal = (dblTmp05 Or dblTmp06)
18040         End If
18050       Else
18060         If Left(varInput, 2) = "&H" Or Left(varInput, 2) = "0x" Then
18070           strInput = Mid(varInput, 3)
18080           If Right(strInput, 1) = "&" Then strInput = Left(strInput, (Len(strInput) - 1))
18090           intLen = Len(strInput)
18100         Else
18110           strInput = varInput
18120         End If
18130         For intX = 1 To intLen
18140           strTmp01 = Mid(strInput, intX, 1)
18150           If Asc(strTmp01) >= 65 And Asc(strTmp01) <= 70 Then       ' ** A - F.
18160             dblTmp05 = (Asc(strTmp01) - 64) + 9
18170           ElseIf Asc(strTmp01) >= 97 And Asc(strTmp01) <= 102 Then  ' ** a - f.
18180             dblTmp05 = (Asc(strTmp01) - 96) + 9
18190           ElseIf Asc(strTmp01) >= 48 And Asc(strTmp01) <= 57 Then   ' ** 0 - 9
18200             dblTmp05 = Val(strTmp01)
18210           Else
18220             Beep
18230             MsgBox "Invalid Character.", vbInformation + vbOKOnly, "Invalid Characters"
18240             Exit For
18250           End If
                'FFFFFFFF
                'FF FF FF FF
18260           dblTmp05 = dblTmp05 * (16 ^ (intLen - intX))
18270           dblTmp06 = dblTmp06 + dblTmp05
18280         Next
18290         varRetVal = dblTmp06
18300       End If
18310     End If
18320   End If

EXITP:
18330   HexX = varRetVal
18340   Exit Function

ERRH:
18350   Select Case ERR.Number
        Case Else
18360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18370   End Select
18380   Resume EXITP

End Function

Public Function QueryExists(strQryName As String) As Boolean
' ** Determine if the query exists.

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "QueryExists"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnRetVal As Boolean

18410   blnRetVal = False

18420   If strQryName <> vbNullString Then
18430     Set dbs = CurrentDb
18440     With dbs
18450       For Each qdf In .QueryDefs
18460         If qdf.Name = strQryName Then
18470           blnRetVal = True
18480           Exit For
18490         End If
18500       Next
18510       .Close
18520     End With
18530   End If

EXITP:
18540   Set qdf = Nothing
18550   Set dbs = Nothing
18560   QueryExists = blnRetVal
18570   Exit Function

ERRH:
18580   blnRetVal = False
18590   Select Case ERR.Number
        Case Else
18600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18610   End Select
18620   Resume EXITP

End Function

Public Function FilterString_Load() As Boolean
' ** Populate the public filter string variables declared in modFileInfoFuncs.

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterString_Load"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

18710   blnRetVal = True

18720   Set dbs = CurrentDb
18730   With dbs
18740     Set rst = .OpenRecordset("tblFilterString", dbOpenDynaset, dbReadOnly)
18750     With rst
18760       .MoveLast
18770       lngRecs = .RecordCount
18780       .MoveFirst
18790       For lngX = 1& To lngRecs

18800         Select Case ![fltstr_constant_n]
              Case "FLTR_ALL1"  ' ** "All Files (*.*)"
18810           FLTR_ALL1 = ![fltstr_name]
18820         Case "FLTR_DLL1"  ' ** "Dynamic Link Libraries (*.dll)"
18830           FLTR_DLL1 = ![fltstr_name]
18840         Case "FLTR_EXE1"  ' ** "Executables (*.exe)"
18850           FLTR_EXE1 = ![fltstr_name]
18860         Case "FLTR_INI1"  ' ** "Initialization Files (*.ini)"
18870           FLTR_INI1 = ![fltstr_name]
18880         Case "FLTR_LDB1"  ' ** "Microsoft Access Lock (*.ldb)"
18890           FLTR_LDB1 = ![fltstr_name]
18900         Case "FLTR_MDB1"  ' ** "Microsoft Access (*.mdb)"
18910           FLTR_MDB1 = ![fltstr_name]
18920         Case "FLTR_MDE1"  ' ** "Microsoft Access (*.mde)"
18930           FLTR_MDE1 = ![fltstr_name]
18940         Case "FLTR_MDW1"  ' ** "Microsoft Access Workgroup (*.mdw)"
18950           FLTR_MDW1 = ![fltstr_name]
18960         Case "FLTR_MD_1"  ' ** "Backup Files (*.md_)"
18970           FLTR_MD_1 = ![fltstr_name]
18980         Case "FLTR_OCX1"  ' ** "OLE Control Extensions (*.ocx)"
18990           FLTR_OCX1 = ![fltstr_name]
19000         Case "FLTR_TLB1"  ' ** "OLE Type Libraries (*.tlb)"
19010           FLTR_TLB1 = ![fltstr_name]
19020         Case "FLTR_VXD1"  ' ** "Virtual Device Drivers (*.vxd)"
19030           FLTR_VXD1 = ![fltstr_name]
19040         Case "FLTR_XLS1"  ' ** "Microsoft Excel (*.xls)"
19050           FLTR_XLS1 = ![fltstr_name]
19060         End Select

19070         Select Case ![fltstr_constant_e]
              Case "FLTR_ALL2"  ' ** "*.*"
19080           FLTR_ALL2 = ![fltstr_extension]
19090         Case "FLTR_DLL2"  ' ** "*.dll"
19100           FLTR_DLL2 = ![fltstr_extension]
19110         Case "FLTR_EXE2"  ' ** "*.exe"
19120           FLTR_EXE2 = ![fltstr_extension]
19130         Case "FLTR_INI2"  ' ** "*.ini"
19140           FLTR_INI2 = ![fltstr_extension]
19150         Case "FLTR_LDB2"  ' ** "*.ldb"
19160           FLTR_LDB2 = ![fltstr_extension]
19170         Case "FLTR_MDB2"  ' ** "*.mdb"
19180           FLTR_MDB2 = ![fltstr_extension]
19190         Case "FLTR_MDE2"  ' ** "*.mde"
19200           FLTR_MDE2 = ![fltstr_extension]
19210         Case "FLTR_MDW2"  ' ** "*.mdw"
19220           FLTR_MDW2 = ![fltstr_extension]
19230         Case "FLTR_MD_2"  ' ** "*.md_"
19240           FLTR_MD_2 = ![fltstr_extension]
19250         Case "FLTR_OCX2"  ' ** "*.ocx"
19260           FLTR_OCX2 = ![fltstr_extension]
19270         Case "FLTR_TLB2"  ' ** "*.tlb"
19280           FLTR_TLB2 = ![fltstr_extension]
19290         Case "FLTR_VXD2"  ' ** "*.vxd"
19300           FLTR_VXD2 = ![fltstr_extension]
19310         Case "FLTR_XLS2"  ' ** "*.xls"
19320           FLTR_XLS2 = ![fltstr_extension]
19330         End Select

19340         If lngX < lngRecs Then .MoveNext
19350       Next
19360       .Close
19370     End With
19380     .Close
19390   End With

EXITP:
19400   Set rst = Nothing
19410   Set dbs = Nothing
19420   FilterString_Load = blnRetVal
19430   Exit Function

ERRH:
19440   blnRetVal = False
19450   Select Case ERR.Number
        Case Else
19460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19470   End Select
19480   Resume EXITP

End Function
