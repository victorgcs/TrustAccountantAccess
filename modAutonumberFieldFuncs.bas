Attribute VB_Name = "modAutonumberFieldFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modAutonumberFieldFuncs"

'VGC 11/23/2016: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const IsDev = 0  ' ** 0 = release, -1 = development.
' ** Also in:
' **   frmXAdmin_Misc
' **   modExcelFuncs
' **   modVersionDocFuncs
' **   zz_mod_MDEPrepFuncs

'THIS MODULE REQUIRES REFERENCE TO msadox.dll!
'Microsoft ADO Ext. 6.0 for DDL and Security (or similar, earlier version)

' ** Note: For this code to run correctly, you must reference both the
' ** Microsoft ActiveX Data Objects 2.x and the Microsoft ADO Ext 2.x for
' ** DDL and Security Libraries (where 2.x is 2.1 or later.)

Private strPathFile As String
Private lngRetCode As Long
Private strCalledPath As String, strCalledDB As String
Private lngDatabases1 As Long, lngDatabases2 As Long, lngDatabases3 As Long
' **

Public Function ChangeSeed_All(Optional varDbs As Variant) As Boolean
' ** Reset all AutoNumber fields in each linked database.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "ChangeSeed_All"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngAutos As Long, arr_varAuto() As Variant
        Dim strThisDB As String, strThisDBPath As String
        Dim strQuery As String, strDbs As String
        Dim lngRecs As Long
        Dim blnRetCode As Boolean, blnMeterOn As Boolean
        Dim strTmp01 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        Const A_ELEMS As Integer = 10  ' ** Array's first-element UBound().
        Const A_AID  As Integer = 0
        Const A_DID  As Integer = 1
        Const A_DNAM As Integer = 2
        Const A_PATH As Integer = 3
        Const A_TID  As Integer = 4
        Const A_TNAM As Integer = 5
        Const A_FID  As Integer = 6
        Const A_FNAM As Integer = 7
        Const A_LAST As Integer = 8
        Const A_NEXT As Integer = 9
        Const A_DONE As Integer = 10

110     blnRetVal = True
120     blnMeterOn = False: lngDatabases1 = 0&: lngDatabases2 = 0&: lngDatabases3 = 0&
130     strCalledPath = vbNullString: strCalledDB = vbNullString

140     Select Case IsMissing(varDbs)
        Case True
          ' ** tblDatabase_AutoNumber, just dbs_id = 1, 2, 4, 6
150       strQuery = "qryDatabase_AutoNumber_01"
160       strDbs = "All"
170     Case False
180       Select Case varDbs
          Case "All"
            ' ** tblDatabase_AutoNumber, just dbs_id = 1, 2, 4, 6
190         strQuery = "qryDatabase_AutoNumber_01"
200         strDbs = varDbs
210       Case "Trust"
            ' ** tblDatabase_AutoNumber, just dbs_id = 1, Trust.mdb.
220         strQuery = "qryDatabase_AutoNumber_02_01"
230         strDbs = varDbs
240       Case "TrustDta"
            ' ** tblDatabase_AutoNumber, just dbs_id = 2, TrustDta.mdb.
250         strQuery = "qryDatabase_AutoNumber_02_02"
260         strDbs = varDbs
270       Case "TrustAux"
            ' ** tblDatabase_AutoNumber, just dbs_id = 4, TrustAux.mdb.
280         strQuery = "qryDatabase_AutoNumber_02_03"
290         strDbs = varDbs
300       Case "TrstXAdm"
            ' ** tblDatabase_AutoNumber, just dbs_id = 6, TrstXAdm.mdb.
310         strQuery = "qryDatabase_AutoNumber_02_04"
320         strDbs = varDbs
330       Case "NonLocal"
            ' ** tblDatabase_AutoNumber, just dbs_id = 2, 4, 6.
340         strQuery = "qryDatabase_AutoNumber_02_05"
350         strDbs = varDbs
360       End Select
370     End Select
        ' ** Switch to the database window.
380     DoCmd.SelectObject acMacro, "mcrAutonumberFieldReset_All", True

390     DoCmd.Hourglass True
400     Application.SysCmd acSysCmdSetStatus, "Resetting AutoNumber fields . . ."

410     DoEvents

420     Set dbs = CurrentDb
430     With dbs

440       strThisDB = Parse_File(dbs.Name)  ' ** Module Function: modFileUtilities.
450       strThisDBPath = Parse_Path(dbs.Name)  ' ** Module Function: modFileUtilities.

460       lngAutos = 0&
470       ReDim arr_varAuto(A_ELEMS, 0)
          ' ****************************************************
          ' ** Array: arr_varAuto()
          ' **
          ' **   Field  Element  Name              Constant
          ' **   =====  =======  ================  ===========
          ' **     1       0     autonum_id        A_AID
          ' **     2       1     dbs_id            A_DID
          ' **     3       2     dbs_name          A_DNAM
          ' **     4       3     dbs_path          A_PATH
          ' **     5       4     tbl_id            A_TID
          ' **     6       5     tbl_name          A_TNAM
          ' **     7       6     fld_id            A_FID
          ' **     8       7     fld_name          A_FNAM
          ' **     9       8     autonum_lastid    A_LAST
          ' **    10       9     Next Number       A_NEXT
          ' **    11      10     Done              A_DONE
          ' **
          ' ****************************************************

          ' ** tblDatabase_AutoNumber, just . . .
480       Set qdf = .QueryDefs(strQuery)
490       Set rst = qdf.OpenRecordset
500       With rst
510         .MoveLast
520         lngRecs = .RecordCount
530         .MoveFirst
540         For lngX = 1& To lngRecs
550           lngAutos = lngAutos + 1&
560           lngE = lngAutos - 1&
570           ReDim Preserve arr_varAuto(A_ELEMS, lngE)
580           arr_varAuto(A_AID, lngE) = ![autonum_id]
590           arr_varAuto(A_DID, lngE) = ![dbs_id]
600           arr_varAuto(A_DNAM, lngE) = ![dbs_name]
610           arr_varAuto(A_PATH, lngE) = ![dbs_path]
620           If Right(arr_varAuto(A_PATH, lngE), 1) = LNK_SEP Then
630             arr_varAuto(A_PATH, lngE) = Left(arr_varAuto(A_PATH, lngE), (Len(arr_varAuto(A_PATH, lngE)) - 1))
640           End If
650           arr_varAuto(A_TID, lngE) = ![tbl_id]
660           arr_varAuto(A_TNAM, lngE) = ![tbl_name]
670           arr_varAuto(A_FID, lngE) = ![fld_id]
680           arr_varAuto(A_FNAM, lngE) = ![fld_name]
690           arr_varAuto(A_LAST, lngE) = ![autonum_lastid]
700           arr_varAuto(A_NEXT, lngE) = CLng(0)
710           arr_varAuto(A_DONE, lngE) = CBool(False)
720           If lngX < lngRecs Then .MoveNext
730         Next
740         .Close
750       End With
760       .Close
770     End With
780     Set dbs = Nothing

790     Application.SysCmd acSysCmdInitMeter, "Resetting AutoNumber Fields . . .", lngAutos
800     blnMeterOn = True
810     DoEvents

820     For lngX = 0 To (lngAutos - 1&)
830       strCalledDB = arr_varAuto(A_DNAM, lngX)
840       strCalledPath = arr_varAuto(A_PATH, lngX)
850       blnRetCode = ChangeSeed_Ext(arr_varAuto(A_TNAM, lngX), arr_varAuto(A_PATH, lngX) & LNK_SEP & arr_varAuto(A_DNAM, lngX))
          ' ** Function: Below.
860       arr_varAuto(A_DONE, lngX) = CBool(True)
870       If blnRetCode = False Then
880         Select Case lngRetCode
            Case -1&
890           strTmp01 = "NOT FND"
900         Case -2&
910           strTmp01 = Left("NO AN" & Space(7), 7)
920         Case Else
930           strTmp01 = Left(CStr(lngRetCode) & Space(7), 7)
940         End Select
            'Debug.Print "'CODE: " & strTmp01 & "  " & arr_varAuto(A_TNAM, lngX)
950       Else
960         arr_varAuto(A_NEXT, lngX) = lngRetCode
970       End If
980       Application.SysCmd acSysCmdUpdateMeter, (lngX + 1&)
990       lngDatabases1 = lngDatabases1 + 1&
1000      DoEvents
1010    Next  ' ** lngX.

1020    Application.SysCmd acSysCmdRemoveMeter
1030    DoCmd.Hourglass False

1040    Beep
1050    MsgBox "Tables for " & strDbs & " renumbered: " & CStr(lngDatabases1), vbInformation + vbOKOnly, ("Finished" & Space(40))

EXITP:
1060    Application.SysCmd acSysCmdClearStatus
1070    Set rst = Nothing
1080    Set qdf = Nothing
1090    Set dbs = Nothing
1100    ChangeSeed_All = blnRetVal
1110    Exit Function

ERRH:
1120    blnRetVal = False
1130    If blnMeterOn = True Then
1140      Application.SysCmd acSysCmdRemoveMeter
1150    End If
1160    DoCmd.Hourglass False
1170    Select Case ERR.Number
        Case Else
1180      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "File: " & CurrentDb.Name & vbCrLf & "Module: " & THIS_NAME & vbCrLf & _
            "Proc: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1190    End Select
1200    Resume EXITP

End Function

Public Function ChangeSeed_Ext(Optional varTblName1 As Variant, Optional varPathFile1 As Variant) As Boolean
' ** ChangeSeed Extended.

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "ChangeSeed_Ext"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnCalled As Boolean
        Dim strPath As String, strFile As String, strTable As String, strTable2 As String, strField As String, strField2 As String
        Dim lngRecs As Long, lngLastNum As Long, lngNextNum As Long, strNextNum As String
        Dim blnSpace_tdf As Boolean, blnSpace_fld As Boolean
        Dim strSQL As String
        Dim blnFound As String, blnIsLocal As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim strTmpTbl As String, strTmpFld As String
        Dim blnRetVal As Boolean

1310    blnRetVal = True

1320    If IsMissing(varTblName1) = True Then
1330      blnCalled = False
1340    Else
1350      blnCalled = True
1360      If IsMissing(varPathFile1) = True Then
1370        strPath = vbNullString
1380        strFile = vbNullString
1390      Else
1400        strFile = Parse_File(varPathFile1)  ' ** Module Function: modFileUtilities.
1410        strPath = Parse_Path(varPathFile1)  ' ** Module Function: modFileUtilities.
1420      End If
1430    End If

1440    If blnCalled = False Then
1450      Application.SysCmd acSysCmdSetStatus, "Resetting Autonumber Field . . ."
1460      DoEvents
1470    End If

1480    strTable = vbNullString: strTable2 = vbNullString: strField = vbNullString: strField2 = vbNullString
1490    blnSpace_tdf = False: blnSpace_fld = False: blnIsLocal = False

1500    If blnCalled = False Then
1510      strTable = InputBox("Table Name:", "The table whose Autonumber field you wish to reset")
1520    Else
1530      strTable = varTblName1
1540    End If

1550    If strTable <> vbNullString Then
1560      strTable = Trim(strTable)
1570      If Len(strTable) > 0 Then

1580        If InStr(strTable, " ") > 0 Then
1590          blnSpace_tdf = True
1600          strTable2 = "[" & strTable & "]"
1610        End If

1620        If strFile = vbNullString Then
              ' ** Check to see if it's the name of a linked table or a local table.
1630          Set dbs = CurrentDb
1640          With dbs
1650            blnFound = False
1660            For Each tdf In .TableDefs
1670              With tdf
1680                If .Name = strTable Then
1690                  If .Connect <> vbNullString Then
1700                    strFile = Parse_File(.Connect)  ' ** Module Function: modFileUtilities.
1710                    If strFile = "TrustDta.mdb" Or strFile = "TrstArch.mdb" Or strFile = "TrustAux.mdb" Then
1720                      blnFound = True
1730                      strPath = Mid(.Connect, (InStr(.Connect, LNK_IDENT) + Len(LNK_IDENT)))
1740                      strPath = Parse_Path(strPath)  ' ** Module Function: modFileUtilities.
1750                    Else
1760                      blnRetVal = False
1770                      MsgBox "Currently, this function only works with TrustDta.mdb or TrstArch.mdb." & vbCrLf & vbCrLf & _
                            "The table you specified is found in: " & strFile, vbInformation + vbOKOnly, "Table Not In TrustDta.mdb Or TrstArch.mdb"
1780                    End If
1790                  Else
1800                    blnFound = True
1810                    blnIsLocal = True
1820                    strPath = vbNullString
1830                  End If
1840                  Exit For
1850                End If
1860              End With
1870            Next
1880            If blnFound = False And blnRetVal = True Then
1890              blnRetVal = False
1900              MsgBox "The specified table could not be found.", vbInformation + vbOKOnly, "Table Not Found"
1910            End If
1920            .Close
1930          End With
1940        Else
1950          If strPath = Parse_Path(CurrentDb.Name) And strFile = Parse_File(CurrentDb.Name) Then  ' ** Module Functions: modFileUtilities.
1960            blnIsLocal = True
1970          Else
1980            blnIsLocal = False
1990          End If
2000        End If

2010        If blnRetVal = True Then
2020          strPathFile = strPath & LNK_SEP & strFile
2030          If blnIsLocal = True Then
2040            Set wrk = DBEngine.Workspaces(0)
2050          Else
2060  On Error Resume Next
2070            Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
2080            If ERR.Number <> 0 Then
2090  On Error GoTo ERRH
2100  On Error Resume Next
2110              Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
2120              If ERR.Number <> 0 Then
2130  On Error GoTo ERRH
2140  On Error Resume Next
2150                Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
2160                If ERR.Number <> 0 Then
2170  On Error GoTo ERRH
2180  On Error Resume Next
2190                  Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
2200                  If ERR.Number <> 0 Then
2210  On Error GoTo ERRH
2220  On Error Resume Next
2230                    Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
2240                    If ERR.Number <> 0 Then
2250  On Error GoTo ERRH
2260  On Error Resume Next
2270                      Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
2280                      If ERR.Number <> 0 Then
2290  On Error GoTo ERRH
2300  On Error Resume Next
2310                        Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
2320  On Error GoTo ERRH
2330                      Else
2340  On Error GoTo ERRH
2350                      End If
2360                    Else
2370  On Error GoTo ERRH
2380                    End If
2390                  Else
2400  On Error GoTo ERRH
2410                  End If
2420                Else
2430  On Error GoTo ERRH
2440                End If
2450              Else
2460  On Error GoTo ERRH
2470              End If
2480            Else
2490  On Error GoTo ERRH
2500            End If
2510          End If
2520          With wrk
2530            If blnIsLocal = True Then
2540              Set dbs = .Databases(0)
2550            Else
2560              Set dbs = .OpenDatabase(strPathFile, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
2570            End If
2580            With dbs

2590              Set tdf = .TableDefs(strTable)
2600              With tdf
2610                blnFound = False
2620                For Each fld In .Fields
2630                  With fld
2640                    If CBool(.Attributes And dbAutoIncrField) = True Then
                        ' **      1  dbDescending
                        ' **      1  dbFixedField
                        ' **      2  dbVariableField
                        ' **     16  dbAutoIncrField
                        ' **     32  dbUpdatableField
                        ' **   8192  dbSystemField
                        ' **  32768  dbHyperlinkField
2650                      blnFound = True
2660                      strField = .Name
2670                      If InStr(strField, " ") > 0 Then
2680                        blnSpace_fld = True
2690                        strField2 = "[" & strField & "]"
2700                      End If
2710                      Exit For
2720                    End If
                        ' ** Field Attribute constant enumeration:
                        ' **      1  dbDescending      The field is sorted in descending (Z to A or 100 to 0) order;
                        ' **                           this option applies only to a Field object in a Fields collection of an Index object.
                        ' **                           If you omit this constant, the field is sorted in ascending (A to Z or 0 to 100) order.
                        ' **                           This is the default value for Index and TableDef fields (Microsoft Jet workspaces only).
                        ' **      1  dbFixedField      The field size is fixed (default for Numeric fields).
                        ' **      2  dbVariableField   The field size is variable (Text fields only).
                        ' **     16  dbAutoIncrField   The field value for new records is automatically incremented to a unique Long integer
                        ' **                           that can't be changed (in a Microsoft Jet workspace, supported only for
                        ' **                           Microsoft Jet database(.mdb) tables).
                        ' **     32  dbUpdatableField  The field value can be changed.
                        ' **   8192  dbSystemField     The field stores replication information for replicas;
                        ' **                           you can't delete this type of field (Microsoft Jet workspaces only).
                        ' **  32768  dbHyperlinkField  The field contains hyperlink information (Memo fields only).
2730                  End With
2740                Next
2750                If blnFound = False Then
2760                  blnRetVal = False
2770                  Debug.Print "'There is no Autonumber field in the specified table: " & strTable & "."
2780                  MsgBox "There is no Autonumber field in the specified table.", vbInformation + vbOKOnly, "Autonumber Field Not Found"
2790                End If
2800              End With

2810              If blnRetVal = True Then
2820                lngRecs = -1&
2830                lngLastNum = -1&
2840                lngNextNum = 0&
2850                Set rst = .OpenRecordset(strTable, dbOpenDynaset, dbReadOnly)
2860                With rst
2870                  If .BOF = True And .EOF = True Then
2880                    lngRecs = 0&
2890                    lngLastNum = 0&
2900                  Else
2910                    .MoveLast
2920                    lngRecs = .RecordCount
2930                  End If
2940                  .Close
2950                End With
2960              End If

2970              If lngRecs > 0& Then
2980                If blnSpace_tdf = True Then strTmpTbl = strTable2 Else strTmpTbl = strTable
2990                If blnSpace_fld = True Then strTmpFld = strField2 Else strTmpFld = strField
3000                If Left(strTable, 2) = "_~" Then strTmpTbl = "[" & strTable & "]"
3010                strSQL = "SELECT Max(" & strTmpTbl & "." & strTmpFld & ") AS " & strTmpFld & " " & _
                      "FROM " & strTmpTbl & ";"
3020                Set rst = .OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
3030                With rst
3040                  .MoveFirst
3050                  lngLastNum = .Fields(strField)
3060                  .Close
3070                End With
3080              End If

3090              .Close
3100            End With
3110            .Close
3120          End With
3130        End If

3140        If blnRetVal = True Then
3150          lngNextNum = lngLastNum + 1&
3160          If blnCalled = False Then
3170            msgResponse = MsgBox("The field '" & strField & "' in the table '" & strTable & "'" & vbCrLf & _
                  "has a current high value of: " & CStr(lngLastNum) & vbCrLf & vbCrLf & _
                  "Click Yes to set the next value to: " & CStr(lngNextNum) & vbCrLf & _
                  "Click No to choose your own next value." & vbCrLf & _
                  "Click Cancel to do nothing.", vbQuestion + vbYesNoCancel, "Reset Autonumber")
3180          Else
3190            msgResponse = vbYes
3200          End If
3210          If msgResponse <> vbCancel Then

3220            If msgResponse = vbNo Then
3230              strNextNum = InputBox("Next Value:", "Choose Next Autonumber Value", CStr(lngNextNum))
3240              If Trim(strNextNum) <> vbNullString Then
3250                If IsNumeric(strNextNum) = True Then
3260                  lngNextNum = Val(strNextNum)
3270                  If lngNextNum = 0& Then
3280                    blnRetVal = False
3290                    MsgBox "You can't set an Autonumber field to 0.", vbInformation + vbOKOnly, "Input Evaluated To Zero"
3300                  Else
3310                    If lngNextNum = lngLastNum Then
3320                      blnRetVal = False
3330                      MsgBox "You can't set the next value to the current highest value.", vbInformation + vbOKOnly, "Invalid Entry"
3340                    Else
3350                      If lngNextNum < lngLastNum Then
3360                        If MsgBox("The value is less than the current highest value!" & vbCrLf & _
                                "If you proceed, Access will eventually assign an existing value." & vbCrLf & vbCrLf & _
                                "Do you wish to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Ill-Advised Assignment") = vbNo Then
3370                          blnRetVal = False
3380                        End If
3390                      End If
3400                    End If
3410                  End If
3420                Else
3430                  blnRetVal = False
3440                  MsgBox "An invalid character was found.", vbInformation + vbOKOnly, ("Invalid Characters" & Space(30))
3450                End If
3460              End If
3470            End If

3480            If blnRetVal = True Then
3490              If blnIsLocal = True Then
3500                Set wrk = DBEngine.Workspaces(0)
3510              Else
3520  On Error Resume Next
3530                Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
3540                If ERR.Number <> 0 Then
3550  On Error GoTo ERRH
3560  On Error Resume Next
3570                  Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
3580                  If ERR.Number <> 0 Then
3590  On Error GoTo ERRH
3600  On Error Resume Next
3610                    Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
3620                    If ERR.Number <> 0 Then
3630  On Error GoTo ERRH
3640  On Error Resume Next
3650                      Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
3660                      If ERR.Number <> 0 Then
3670  On Error GoTo ERRH
3680  On Error Resume Next
3690                        Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
3700                        If ERR.Number <> 0 Then
3710  On Error GoTo ERRH
3720  On Error Resume Next
3730                          Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
3740                          If ERR.Number <> 0 Then
3750  On Error GoTo ERRH
3760  On Error Resume Next
3770                            Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
3780  On Error GoTo ERRH
3790                          Else
3800  On Error GoTo ERRH
3810                          End If
3820                        Else
3830  On Error GoTo ERRH
3840                        End If
3850                      Else
3860  On Error GoTo ERRH
3870                      End If
3880                    Else
3890  On Error GoTo ERRH
3900                    End If
3910                  Else
3920  On Error GoTo ERRH
3930                  End If
3940                Else
3950  On Error GoTo ERRH
3960                End If
3970              End If
3980              With wrk
3990                If blnIsLocal = True Then
4000                  Set dbs = .Databases(0)
4010                Else
4020                  Set dbs = .OpenDatabase(strPathFile, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
4030                End If
4040                With dbs
                      ' ** This query calls the ChangeSeed() function as the field RetVal.
4050                  Set qdf = .QueryDefs("qryChangeSeed_01")
4060                  With qdf.Parameters
4070                    ![tblnam] = strTable
4080                    ![fldnam] = strField
4090                    ![nexnum] = lngNextNum
4100                    ![isloc] = blnIsLocal
4110                  End With
4120                  Set rst = qdf.OpenRecordset
4130                  With rst
4140                    .MoveFirst
4150                    blnRetVal = ![RetVal]
4160                    .Close
4170                  End With
4180                  .Close
4190                End With
4200                .Close
4210              End With
4220              If blnRetVal = True Then
4230                If blnCalled = False Then
4240                  Beep
4250                  MsgBox "The Autonumber field '" & strField & "' in the " & IIf(blnIsLocal = True, "local", "linked") & " table " & _
                        "'" & strTable & "' " & vbCrLf & "has been successfully reset to: " & CStr(lngNextNum), _
                        vbExclamation + vbOKOnly, "Assignment Successful"
4260                End If
4270              Else
4280                If blnCalled = False Then
4290                  Beep
4300                  MsgBox "The assignment failed.", vbCritical + vbOKOnly, "Autonumber Field Not Reset"
4310                End If
4320              End If
4330            End If

4340          Else
4350            blnRetVal = False
4360          End If
4370        End If

4380      Else
4390        If blnCalled = False Then
4400          Beep
4410        End If
4420      End If
4430    Else
4440      If blnCalled = False Then
4450        Beep
4460      End If
4470      blnRetVal = False
4480    End If

4490    If blnCalled = False Then
4500      Application.SysCmd acSysCmdClearStatus
4510    End If

EXITP:
4520    Set rst = Nothing
4530    Set qdf = Nothing
4540    Set fld = Nothing
4550    Set tdf = Nothing
4560    Set dbs = Nothing
4570    Set wrk = Nothing
4580    ChangeSeed_Ext = blnRetVal
4590    Exit Function

ERRH:
4600    blnRetVal = False
4610    Application.SysCmd acSysCmdClearStatus
4620    Select Case ERR.Number
        Case Else
4630      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "File: " & CurrentDb.Name & vbCrLf & "Module: " & THIS_NAME & vbCrLf & _
            "Proc: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4640    End Select
4650    Resume EXITP

End Function

Public Function ChangeSeed(strTbl As String, strCol As String, lngSeed As Long, blnIsLocal As Boolean, Optional varRetSeedOnly As Variant, Optional varPathFile As Variant) As Long
' ** You must pass the following variables to this function.
' **   strTbl = Table containing autonumber field
' **   strCol = Name of the autonumber field
' **   lngSeed = Long integer value you want to use for next AutoNumber.

        'ChangeSeed("tblVBComponent","vbcom_id",1&)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "ChangeSeed"

      #If IsDev Then
        Dim cnxn As ADODB.Connection, catx As New ADOX.Catalog, colx As ADOX.Column  ' ** Early binding.
      #Else
        Dim cnxn As Object, catx As Object, colx As Object                           ' ** Late binding.
      #End If
        Dim strUser As String, strPassword As String
        Dim strSecPath As String
        Dim strCnxn As String
        Dim blnSeedOnly As Boolean
        Dim lngRetVal As Long

4710    lngRetVal = 0&

4720    If IsMissing(varRetSeedOnly) = True Then
4730      blnSeedOnly = False
          ' ** strPathFile, strUser, and strPassword are set within the calling function.
          'NO THEY'RE NOT!!
4740      strUser = "superuser": strPassword = TA_SEC
4750    Else
4760      blnSeedOnly = varRetSeedOnly
4770      If IsMissing(varPathFile) = False Then
4780        strPathFile = varPathFile
4790      Else
4800        If blnIsLocal = True Then
4810          strPathFile = vbNullString  ' ** Not needed.
4820        Else
4830          lngRetVal = -9&
4840        End If
4850      End If
4860      strUser = "superuser": strPassword = TA_SEC
4870    End If

4880    strSecPath = CurrentBackendPath  ' ** Module Function: modFileUtilities.

4890    If lngRetVal = 0& Then

4900      If blnIsLocal = True Then
            ' ** Set connection and catalog to current database.
4910        Set cnxn = CurrentProject.Connection
4920      Else

      #If IsDev Then
4930        Set cnxn = New ADODB.Connection              ' ** Early binding.
      #Else
4940        Set cnxn = CreateObject("ADODB.Connection")  ' ** Late binding.
      #End If
            ' ** Open connection.

            'strCnxn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            '  "Data Source=" & strPathFile & ";" & _
            '  "Jet OLEDB:System database=" & Parse_Path(strPathFile) & LNK_SEP & "TrustSec.mdw"  ' ** Module Function: modFileUtilities.
4950        strCnxn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
              "Data Source=" & strPathFile & ";" & _
              "Jet OLEDB:System database=" & strSecPath & LNK_SEP & "TrustSec.mdw"  ' ** Module Function: modFileUtilities.
4960  On Error Resume Next
4970        cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword  ' ** Superuser, TA_SEC.
4980        If ERR.Number <> 0 Then
4990  On Error GoTo ERRH
5000          If Len(TA_SEC) < Len(TA_SEC2) Then  ' ** This is not the Demo.
5010            If strPassword = TA_SEC Then strPassword = TA_SEC5  ' ** Superuser, old password.
5020  On Error Resume Next
5030            cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5040            If ERR.Number <> 0 Then
5050  On Error GoTo ERRH
5060              strUser = "TAAdmin"
5070              strPassword = TA_SEC3  ' ** New TAAdmin password.
5080  On Error Resume Next
5090              cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5100              If ERR.Number <> 0 Then
5110  On Error GoTo ERRH
5120                strUser = "Admin"
5130                strPassword = TA_SEC7  ' ** Old Admin password.
5140  On Error Resume Next
5150                cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5160                If ERR.Number <> 0 Then
5170  On Error GoTo ERRH
5180                  strPassword = vbNullString  ' ** Generic Admin, no password.
5190  On Error Resume Next
5200                  cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5210                  If ERR.Number <> 0 Then
5220  On Error GoTo ERRH
                        ' ** Try Demo.
5230                    strUser = "Superuser"
5240                    strPassword = TA_SEC2  ' ** Superuser, new Demo password.
5250  On Error Resume Next
5260                    cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5270                    If ERR.Number <> 0 Then
5280  On Error GoTo ERRH
5290                      strPassword = TA_SEC6  ' ** Superuser, old Demo password.
5300  On Error Resume Next
5310                      cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5320                      If ERR.Number <> 0 Then
5330  On Error GoTo ERRH
5340                        strUser = "TADemo"
5350                        strPassword = TA_SEC4  ' ** New TADemo password.
5360  On Error Resume Next
5370                        cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5380                        If ERR.Number <> 0 Then
5390  On Error GoTo ERRH
5400                          strUser = "Demo"
5410                          strPassword = TA_SEC8  ' ** Old Demo password
5420  On Error Resume Next
5430                          cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5440                          If ERR.Number <> 0 Then
5450  On Error GoTo ERRH
                                ' ** I GIVE UP!
5460                          Else
5470  On Error GoTo ERRH
5480                          End If
5490                        Else
5500  On Error GoTo ERRH
5510                        End If
5520                      Else
5530  On Error GoTo ERRH
5540                      End If
5550                    Else
5560  On Error GoTo ERRH
5570                    End If
5580                  Else
5590  On Error GoTo ERRH
5600                  End If
5610                Else
5620  On Error GoTo ERRH
5630                End If
5640              Else
5650  On Error GoTo ERRH
5660              End If
5670            Else
5680  On Error GoTo ERRH
5690            End If
5700          Else  ' ** This is the Demo.
5710            If strPassword = TA_SEC Then strPassword = TA_SEC6  ' ** Superuser, old Demo password.
5720  On Error Resume Next
5730            cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5740            If ERR.Number <> 0 Then
5750  On Error GoTo ERRH
5760              strUser = "TADemo"
5770              strPassword = TA_SEC4  ' ** New TADemo password.
5780  On Error Resume Next
5790              cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5800              If ERR.Number <> 0 Then
5810  On Error GoTo ERRH
5820                strUser = "Demo"
5830                strPassword = TA_SEC8  ' ** Old Demo password
5840  On Error Resume Next
5850                cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5860                If ERR.Number <> 0 Then
5870  On Error GoTo ERRH
5880                  strUser = "Admin"
5890                  strPassword = vbNullString  ' ** Generic Admin, no password.
5900  On Error Resume Next
5910                  cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5920                  If ERR.Number <> 0 Then
5930  On Error GoTo ERRH
                        ' ** Try non-Demo.
5940                    strUser = "Superuser"
5950                    strPassword = TA_SEC2  ' ** Superuser, new non-Demo password.
5960  On Error Resume Next
5970                    cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
5980                    If ERR.Number <> 0 Then
5990  On Error GoTo ERRH
6000                      strPassword = TA_SEC5  ' ** Superuser, old non-Demo password.
6010  On Error Resume Next
6020                      cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
6030                      If ERR.Number <> 0 Then
6040  On Error GoTo ERRH
6050                        strUser = "TAAdmin"
6060                        strPassword = TA_SEC3  ' ** New TAAdmin password.
6070  On Error Resume Next
6080                        cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
6090                        If ERR.Number <> 0 Then
6100  On Error GoTo ERRH
6110                          strUser = "Admin"
6120                          strPassword = TA_SEC7  ' ** Old Admin password.
6130  On Error Resume Next
6140                          cnxn.Open ConnectionString:=strCnxn, UserId:=strUser, Password:=strPassword
6150                          If ERR.Number <> 0 Then
6160  On Error GoTo ERRH
                                ' ** I GIVE UP!
6170                          Else
6180  On Error GoTo ERRH
6190                          End If
6200                        Else
6210  On Error GoTo ERRH
6220                        End If
6230                      Else
6240  On Error GoTo ERRH
6250                      End If
6260                    Else
6270  On Error GoTo ERRH
6280                    End If
6290                  Else
6300  On Error GoTo ERRH
6310                  End If
6320                Else
6330  On Error GoTo ERRH
6340                End If
6350              Else
6360  On Error GoTo ERRH
6370              End If
6380            Else
6390  On Error GoTo ERRH
6400            End If
6410          End If
6420        Else
6430  On Error GoTo ERRH
6440        End If
6450      End If

      #If IsDev Then
6460      Set catx = New ADOX.Catalog              ' ** Early binding.
      #Else
6470      Set catx = CreateObject("ADOX.Catalog")  ' ** Late binding.
      #End If

6480      catx.ActiveConnection = cnxn

6490      Set colx = catx.Tables(strTbl).Columns(strCol)

6500      If blnSeedOnly = False Then
6510        colx.Properties("Seed") = lngSeed
6520        catx.Tables(strTbl).Columns.Refresh
6530        If colx.Properties("Seed") = lngSeed Then
6540          lngRetVal = -1&
6550        Else
6560          lngRetVal = 0&
6570        End If
6580      Else
            'Debug.Print "'SEED: " & CStr(colx.Properties("Seed"))  ' ** Tells us what the current Seed is.
6590        lngRetVal = colx.Properties("Seed")
6600      End If

6610    End If

EXITP:
6620    Set colx = Nothing
6630    Set catx = Nothing
6640    Set cnxn = Nothing
6650    ChangeSeed = lngRetVal
6660    Exit Function

ERRH:
6670    lngRetVal = -9&
6680    Select Case ERR.Number
        Case Else
6690      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "File: " & CurrentDb.Name & vbCrLf & "Module: " & THIS_NAME & vbCrLf & _
            "Proc: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
6700    End Select
6710    Resume EXITP

End Function
