Attribute VB_Name = "modBackendUpdate"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modBackendUpdate"

'VGC 10/27/2017: CHANGES!

'MAKE SURE LEDGER TRANSFERS DON'T DROP PERTINANT ENTRIES!

'Also do:
'  Title:
'  Application Title:
'  Author:
'  Manager:
'  Company:
'And:
'  Display Database Window = False
'  Allow Built-in Toolbars = False
'  Allow Toolbar/Menu Changes = False
'  Allow Full Menus = False
'  Allow Default Shortcut Menus = False
'  Use Access Special Keys = False

Private Const FRM_WAIT As String = "frmPleaseWait"
Private Const UPD_MSG  As String = "One-Time Data File Check... "

Private strUpdates As String, lngUpdates As Long, arr_varUpdate() As Variant, blnRunUpdates As Boolean
Private strSteps As String
Private strOldDtaVer As String, strOldArchVer As String
' **

Public Function Backend_Update(strCallForm As String) As Boolean
' ** Use this function to update the backend databases as necessary.
' ** Called by:
' **   frmMenu_Title.Form_Open()
' ** ADD ANY NEW 'tmp_' TBLES TO TmpTblList(), BELOW!

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Backend_Update"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, tdf As DAO.TableDef, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim idx As DAO.index, fld As DAO.Field, Rel As DAO.Relation
        Dim frm As Access.Form 'Form_frmPleaseWait
        Dim intWrkType As Integer
        Dim lngRecs As Long, arr_varRec() As Variant
        Dim intVDMaj As Integer, intVDMin As Integer, intVDRev As Integer
        Dim lngRevCodes As Long
        Dim blnAddOtherCharges As Boolean, blnAddOtherCredits As Boolean, blnAddOrdinaryDividend As Boolean, blnAddInterestIncome As Boolean
        Dim blnEdited As Boolean, blnMoveAll As Boolean, blnSkip As Boolean
        Dim blnFound As Boolean, blnFound2 As Boolean, blnSkip2200 As Boolean
        Dim blnFoundErrLineNum As Boolean, blnFoundHidType As Boolean, blnFoundHidMatch As Boolean
        Dim strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long, blnTmp05 As Boolean
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

110     blnRetVal = True
120     blnSkip2200 = True

130     DoCmd.Hourglass True
140     DoEvents
150     blnFoundErrLineNum = False: blnFoundHidType = False: blnFoundHidMatch = False

        ' ** Add each successive change here (unless it gets too much, then put it in a subroutine and/or save it another way).
160     lngUpdates = 23& ' ** (zero-based)  '####  NEW UPDATE!  ####
170     ReDim arr_varUpdate(lngUpdates - 1&)
180     arr_varUpdate(0) = 39870  ' ** The date I last made changes to this procedure: 02/26/2009.
190     arr_varUpdate(1) = 39874  ' ** The date I last made changes to this procedure: 03/02/2009.
200     arr_varUpdate(2) = 39875  ' ** The date I last made changes to this procedure: 03/03/2009.
210     arr_varUpdate(3) = 39880  ' ** The date I last made changes to this procedure: 03/08/2009.
220     arr_varUpdate(4) = 39881  ' ** The date I last made changes to this procedure: 03/09/2009.
230     arr_varUpdate(5) = 39889  ' ** The date I last made changes to this procedure: 03/17/2009.
240     arr_varUpdate(6) = 39892  ' ** The date I last made changes to this procedure: 03/20/2009.
250     arr_varUpdate(7) = 39895  ' ** The date I last made changes to this procedure: 03/23/2009.
260     arr_varUpdate(8) = 39896  ' ** The date I last made changes to this procedure: 03/24/2009.
270     arr_varUpdate(9) = 39944  ' ** The date I last made changes to this procedure: 05/11/2009.
280     arr_varUpdate(10) = 39949  ' ** The date I last made changes to this procedure: 05/16/2009.
290     arr_varUpdate(11) = 39951  ' ** The date I last made changes to this procedure: 05/18/2009.
300     arr_varUpdate(12) = 39955  ' ** The date I last made changes to this procedure: 05/22/2009.  NEW CONVERSION ADDED!
310     arr_varUpdate(13) = 39957  ' ** The date I last made changes to this procedure: 05/24/2009.  AssetType/TaxCode
320     arr_varUpdate(14) = 39984  ' ** The date I last made changes to this procedure: 06/20/2009.  Missed relationships, bad field names
330     arr_varUpdate(15) = 39989  ' ** The date I last made changes to this procedure: 06/25/2009.  tblPreference_User
340     arr_varUpdate(16) = 39993  ' ** The date I last made changes to this procedure: 06/25/2009.  tblPreference_User NO LINK TO USERS!
350     arr_varUpdate(17) = 39998  ' ** The date I last made changes to this procedure: 07/04/2009.  LedgerHidden update
360     arr_varUpdate(18) = 40002  ' ** The date I last made changes to this procedure: 07/04/2009.  LedgerHidden: hid_newmatch
370     arr_varUpdate(19) = 40010  ' ** The date I last made changes to this procedure: 07/16/2009.  dbLong: m_TBL, LedgerHidden
380     arr_varUpdate(20) = 40012  ' ** The date I last made changes to this procedure: 07/16/2009.  TrustAux.mdb
390     arr_varUpdate(21) = 40516  ' ** The date I last made changes to this procedure: 11/04/2010.  v2.2.00
400     arr_varUpdate(22) = 40887  ' ** The date I last made changes to this procedure: 12/10/2011.

410     lngTmp02 = arr_varUpdate(lngUpdates - 1&)  ' ** The date I last made changes.
420     lngTmp03 = CLng(Date)  ' ** Today's date.
430     strUpdates = Right("000" & CStr(lngUpdates), 3) & CStr(lngTmp03)
440     strTmp01 = strUpdates
        ' ** ? Len("00139870D1D2D3A3D40D41D42D43D5D6D7D8D9D10")
        ' **  41 characters  'If this grows too much, drop some, or use a different algorithm.
        ' ** vd_DE2, vd_DE2, vd_DE2: 50 characters.

450     If gstrTrustDataLocation <> vbNullString Then

          ' ** Check to see whether this update has already been run on this frontend and backend.
460       blnRunUpdates = False
470       blnRunUpdates = ChkUpdate(True, lngTmp02)  ' ** Function: Below.
480       blnTmp05 = blnRunUpdates  ' ** If it comes back True, then save the results whether or not strUpdates <> strTmp01.
          '1.1. strOldDtaVer = ''
          '1.2. strOldDtaVer = '2.1.41.'
          '2.1. strOldDtaVer = '2.1.41.'
          '2.2. strOldDtaVer = ''

490       DoCmd.Hourglass True  ' ** Make sure it's still running.
500       DoEvents

510       If blnRunUpdates = True Then

            'FIND A WAY TO CHANGE THE MSG IF IT'S JUST A ROUTINE OPENING!
            'BUT WAIT! IF IT'S A ROUTINE OPENING, IT SHOULDN'T GET HERE, SHOULD IT?

            ' *****************************************************************************
            ' ** Step 1. Initialization.
            ' *****************************************************************************

520         strSteps = " of 6"  '####  NEW UPDATE!  ####
530         DoCmd.OpenForm FRM_WAIT, , , , , , strCallForm & "~" & UPD_MSG & "1" & strSteps
540         DoEvents
550         SysCmd acSysCmdSetStatus, "Checking Data Files. Please wait . . ."
560         Set frm = Forms(FRM_WAIT)

570         intWrkType = 0
580   On Error Resume Next
590         Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
600         If ERR.Number <> 0 Then
610   On Error GoTo ERRH
620   On Error Resume Next
630           Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
640           If ERR.Number <> 0 Then
650   On Error GoTo ERRH
660   On Error Resume Next
670             Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
680             If ERR.Number <> 0 Then
690   On Error GoTo ERRH
700   On Error Resume Next
710               Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
720               If ERR.Number <> 0 Then
730   On Error GoTo ERRH
740   On Error Resume Next
750                 Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
760                 If ERR.Number <> 0 Then
770   On Error GoTo ERRH
780   On Error Resume Next
790                   Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
800                   If ERR.Number <> 0 Then
810   On Error GoTo ERRH
820   On Error Resume Next
830                     Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
840   On Error GoTo ERRH
850                     intWrkType = 7
860                   Else
870   On Error GoTo ERRH
880                     intWrkType = 6
890                   End If
900                 Else
910   On Error GoTo ERRH
920                   intWrkType = 5
930                 End If
940               Else
950   On Error GoTo ERRH
960                 intWrkType = 4
970               End If
980             Else
990   On Error GoTo ERRH
1000              intWrkType = 3
1010            End If
1020          Else
1030  On Error GoTo ERRH
1040            intWrkType = 2
1050          End If
1060        Else
1070  On Error GoTo ERRH
1080          intWrkType = 1
1090        End If

1100        With wrk

              ' *****************************************************************************
              ' ** Step 2. Update tblErrorLog.
              ' *****************************************************************************

1110          If blnSkip2200 = False Then

1120            frm.WaitMsg_lbl.Caption = UPD_MSG & "2" & strSteps
1130            DoEvents

1140            Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
1150            With dbs

1160              Set tdf = .TableDefs("tblErrorLog")
1170              With tdf
1180                blnFound = False
1190                For Each fld In .Fields
1200                  With fld
1210                    If .Name = "ErrLog_LineNum" Then
1220                      blnFoundErrLineNum = False
1230                      blnFound = True
1240                      Exit For
1250                    End If
1260                  End With
1270                Next
1280              End With

1290              If blnFound = False Then

1300                strUpdates = strUpdates & "D2"

                    ' ** Empty tblErrorLog.
1310                Set qdf = CurrentDb.QueryDefs("qryErrLog_05")
1320                qdf.Execute

                    ' ** Delete the current frontend's link.
1330                TableDelete "tblErrorLog"  ' ** Module Function: modFileUtilities.

                    ' ** Copy the new tblErrorLog table to their TrustDta.mdb.
1340                DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblTemplate_ErrorLog", "tblErrorLog", True  ' ** Structure Only.

1350                .TableDefs.Refresh

                    ' ** Relink tblErrorLog back to here.
1360                DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblErrorLog", "tblErrorLog"

1370                CurrentDb.TableDefs.Refresh
1380                blnFoundErrLineNum = True

1390              End If

1400              .Close
1410            End With

1420          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 3. Add the new m_REVCODE_TYPE table and update the m_REVCODE table.
              ' **         Make sure revcode_TYPE in m_REVCODE is dbLong,
              ' **         and it's linked to the m_REVCODE_TYPE table.
              ' *****************************************************************************

1430          If blnSkip2200 = False Then

1440            frm.WaitMsg_lbl.Caption = UPD_MSG & "3" & strSteps
1450            DoEvents

1460            Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
1470            With dbs

1480              blnFound = False: blnFound2 = False
1490              For Each tdf In .TableDefs
1500                With tdf
1510                  If .Name = "m_REVCODE_TYPE" = True Then
1520                    blnFound = True
1530                    Exit For
1540                  End If
1550                End With
1560              Next
1570              If blnFound = True Then
1580                Set tdf = .TableDefs("m_REVCODE")
1590                With tdf
1600                  Set fld = .Fields("revcode_TYPE")
1610                  If fld.Type = dbLong Then
1620                    blnFound2 = True
1630                  End If
1640                End With
1650              End If

1660              If blnFound = False Or blnFound2 = False Then

1670                strUpdates = strUpdates & "D3"

                    ' ** Delete the current frontend's link.
1680                TableDelete "m_REVCODE_TYPE"  ' ** Module Function: modFileUtilities.

                    ' ** Delete the backend copy if it's hangin' around.
1690                TableDelete "tmp_m_REVCODE"  ' ** Module Function: modFileUtilities.

                    ' ** Bring over a copy of their m_REVCODE table as a backup.
1700                DoCmd.TransferDatabase acImport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "m_REVCODE", "tmp_m_REVCODE", False

                    ' ** Empty tblTemplate_m_REVCODE, in case there's anything hangin' around.
1710                Set qdf = CurrentDb.QueryDefs("qryRevCodes_20")  ' ** Empties completely!
1720                qdf.Execute

                    ' ** Append tmp_m_REVCODE to tblTemplate_m_REVCODE, with all their original data.
1730                Set qdf = CurrentDb.QueryDefs("qryRevCodes_21")  ' ** Still has their original revcode_ID's.
1740                qdf.Execute
                    ' ** WHY IS IT THAT I'M REPLACING THEIR m_REVCODE TABLE?
                    ' ** 'Cause it's got a new index!

                    ' ** Copy the new m_REVCODE_TYPE table to their TrustDta.mdb.
1750                If blnFound = False Then
1760                  DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                        acTable, "tblTemplate_m_REVCODE_TYPE", "m_REVCODE_TYPE", False  ' ** Copy with data.
1770                  .TableDefs.Refresh
1780                End If

                    ' ** Link m_REVCODE_TYPE to here.
1790                DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "m_REVCODE_TYPE", "m_REVCODE_TYPE"

1800                CurrentDb.TableDefs.Refresh
1810                CurrentDb.TableDefs("m_REVCODE_TYPE").RefreshLink

1820                lngRecs = 0&
1830                ReDim arr_varRec(0)

                    ' ** Delete the existing relationships to Ledger and Journal.
1840                For Each Rel In .Relations
1850                  With Rel
1860                    If .Table = "m_REVCODE" Or .ForeignTable = "m_REVCODE" Then
1870                      lngRecs = lngRecs + 1&
1880                      ReDim Preserve arr_varRec(lngRecs - 1&)
1890                      arr_varRec(lngRecs - 1&) = .Name
1900                    End If
1910                  End With
1920                Next
1930                If lngRecs > 0& Then
1940                  For lngX = 0& To (lngRecs - 1&)
1950                    .Relations.Delete arr_varRec(lngX)
1960                  Next
1970                End If

                    ' ** Delete their existing m_REVCODE table.
1980                .TableDefs.Delete "m_REVCODE"
1990                .TableDefs.Refresh

                    ' ** Transfer the new m_REVCODE template.
2000                DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblTemplate_m_REVCODE", "m_REVCODE", False  ' ** Copy with data.
2010                .TableDefs.Refresh

                    ' ** Create all the relationships.
2020                Set Rel = .CreateRelation("m_REVCODE_TYPEm_REVCODE", "m_REVCODE_TYPE", "m_REVCODE", dbRelationUpdateCascade)
2030                With Rel
2040                  .Fields.Append .CreateField("revcode_TYPE", dbLong)
2050                  .Fields![revcode_TYPE].ForeignName = "revcode_TYPE"
2060                End With
2070                .Relations.Append Rel
2080                Set Rel = .CreateRelation("m_REVCODEjournal", "m_REVCODE", "journal", dbRelationUpdateCascade)
2090                With Rel
2100                  .Fields.Append .CreateField("revcode_ID", dbLong)
2110                  .Fields![revcode_ID].ForeignName = "revcode_ID"
2120                End With
2130                .Relations.Append Rel
2140                Set Rel = .CreateRelation("m_REVCODEledger", "m_REVCODE", "ledger", dbRelationUpdateCascade)
2150                With Rel
2160                  .Fields.Append .CreateField("revcode_ID", dbLong)
2170                  .Fields![revcode_ID].ForeignName = "revcode_ID"
2180                End With
2190                .Relations.Append Rel
2200                .Relations.Refresh

2210                CurrentDb.TableDefs.Refresh
2220                CurrentDb.TableDefs("m_REVCODE").RefreshLink

2230              End If

2240              .Close
2250            End With  ' ** dbs.

2260          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 4. Add the new InvestmentObjective table.
              ' *****************************************************************************

2270          If blnSkip2200 = False Then

2280            frm.WaitMsg_lbl.Caption = UPD_MSG & "4" & strSteps
2290            DoEvents

2300            Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
2310            With dbs

2320              blnFound = False
2330              For Each tdf In .TableDefs
2340                With tdf
2350                  If .Name = "InvestmentObjective" Then
2360                    blnFound = True
2370                    Exit For
2380                  End If
2390                End With
2400              Next

2410              If blnFound = False Then

2420                strUpdates = strUpdates & "D4"

                    ' ** Copy the new InvestmentObjective table to their TrustDta.mdb.
2430                DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblTemplate_InvestmentObjective", "InvestmentObjective", False  ' ** Copy with data.

2440                .TableDefs.Refresh

                    ' ** Relink InvestmentObjective back to here.
2450                DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "InvestmentObjective", "InvestmentObjective"

2460                CurrentDb.TableDefs.Refresh

                    ' ** Update Account, where investmentobj = '', set  investmentobj = Null.
2470                Set qdf = CurrentDb.QueryDefs("qryInvestmentObjective_03")
2480                qdf.Execute

                    ' ** qryInvestmentObjective_03_01 (Account, grouped), not in InvestmentObjective.
2490                Set qdf = CurrentDb.QueryDefs("qryInvestmentObjective_04")
2500                Set rst = qdf.OpenRecordset
2510                If rst.BOF = True And rst.EOF = True Then
                      ' ** All's well.
2520                  rst.Close
2530                Else
                      ' ** Even though this has always been a Value dropdown (I believe),
                      ' ** I don't want to depend on that.
                      ' ** Append qryInvestmentObjective_04 (those not in new table) to InvestmentObjective.
2540                  Set qdf = CurrentDb.QueryDefs("qryInvestmentObjective_05")
2550                  qdf.Execute
2560                End If

                    ' ** Create new relationship.
2570                Set Rel = .CreateRelation("InvestmentObjectiveaccount", "InvestmentObjective", "account", dbRelationUpdateCascade)
2580                With Rel
2590                  .Fields.Append .CreateField("invobj_name", dbText)
2600                  .Fields![invobj_name].ForeignName = "investmentobj"
2610                End With
2620                .Relations.Append Rel
2630                .Relations.Refresh

2640                CurrentDb.TableDefs.Refresh
2650                CurrentDb.TableDefs("InvestmentObjective").RefreshLink

2660              Else
                    ' ** Make sure it's linked.
2670                blnFound = False
2680                For Each tdf In CurrentDb.TableDefs
2690                  With tdf
2700                    If .Name = "InvestmentObjective" Then
2710                      blnFound = True
2720                      Exit For
2730                    End If
2740                  End With
2750                Next
2760                If blnFound = False Then
                      ' ** Relink InvestmentObjective back to here.
2770                  DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                        acTable, "InvestmentObjective", "InvestmentObjective"
2780                  CurrentDb.TableDefs.Refresh
2790                End If
2800              End If

2810              .Close
2820            End With

2830          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 5. Update RecurringItems table, with new AutoNumber field, RecurringItem_ID
              ' *****************************************************************************

2840          If blnSkip2200 = False Then

2850            frm.WaitMsg_lbl.Caption = UPD_MSG & "5" & strSteps
2860            DoEvents

2870            Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
2880            With dbs

2890              blnFound = False
2900              Set tdf = .TableDefs("RecurringItems")  ' ** Existing table name misspelled!
2910              With tdf
2920                For Each fld In .Fields
2930                  With fld
2940                    If .Name = "RecurringItem_ID" Then
2950                      blnFound = True
2960                      Exit For
2970                    End If
2980                  End With
2990                Next
3000              End With

3010              blnFound2 = False
3020              For Each tdf In .TableDefs
3030                With tdf
3040                  If .Name = "RecurringType" Then
3050                    blnFound2 = True
3060                    Exit For
3070                  End If
3080                End With
3090              Next

3100              If blnFound = False Or blnFound2 = False Then

3110                If blnFound = False Then

3120                  strUpdates = strUpdates & "D5"

                      ' ** Update RecurringItems, where Type = Null, set Type = 'Misc'.
                      ' ** Not likely to ever happen, but just in case...
3130                  Set qdf = CurrentDb.QueryDefs("qryRecurringItems_06")
3140                  qdf.Execute

                      ' ** Delete the current frontend's link.
3150                  TableDelete "RecurringItems"  ' ** Module Function: modFileUtilities.

                      ' ** Delete the backend copy if it's hangin' around.
3160                  TableDelete "tmp_RecurringItems"  ' ** Module Function: modFileUtilities.

                      ' ** Bring over a copy of their RecurringItems table as a backup.
3170                  DoCmd.TransferDatabase acImport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                        acTable, "RecurringItems", "tmp_RecurringItems", False

                      ' ** Empty tblTemplate_RecurringItems, for all but 2 standard entries.
3180                  Set qdf = CurrentDb.QueryDefs("qryRecurringItems_02")
3190                  qdf.Execute

3200                  If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
                        ' ** Reset the Autonumber field.
3210                    ChangeSeed_Ext "tblTemplate_RecurringItems"  ' ** Module Function: modAutonumberFieldFuncs.
3220                  End If

                      ' ** Append tmp_RecurringItems to tblTemplate_RecurringItems, with all their original data.
3230                  Set qdf = CurrentDb.QueryDefs("qryRecurringItems_03")
3240                  qdf.Execute

                      ' ** Update tblTemplate_RecurringItems, set RecurringItem_State = UCase([RecurringItem_State]).
3250                  Set qdf = CurrentDb.QueryDefs("qryRecurringItems_05")
3260                  qdf.Execute

                      ' ** Delete their copy of RecurringItems table.
3270                  .TableDefs.Delete "RecurringItems"

                      ' ** Copy the new RecurringItems table to their TrustDta.mdb.
3280                  DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                        acTable, "tblTemplate_RecurringItems", "RecurringItems", False  ' ** Copy with data.

3290                  .TableDefs.Refresh

                      ' ** Relink RecurringItems back to here.
3300                  DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                        acTable, "RecurringItems", "RecurringItems"

3310                End If

                    ' *****************************************************************************
                    ' ** Step 6. Add the new RecurringType table.
                    ' *****************************************************************************

3320                frm.WaitMsg_lbl.Caption = UPD_MSG & "6" & strSteps
3330                strUpdates = strUpdates & "D6"

                    ' ** Copy the new RecurringType table to their TrustDta.mdb.
3340                DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblTemplate_RecurringType", "RecurringType", False  ' ** Copy with data.

3350                .TableDefs.Refresh

                    ' ** Create new relationship.
3360                Set Rel = .CreateRelation("RecurringTypeRecurringItems", "RecurringType", "RecurringItems", dbRelationUpdateCascade)
3370                With Rel
3380                  .Fields.Append .CreateField("RecurringType", dbText)
3390                  .Fields![RecurringType].ForeignName = "Type"
3400                End With
3410                .Relations.Append Rel
3420                .Relations.Refresh

                    ' ** Link RecurringType to here.
3430                DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "RecurringType", "RecurringType"

3440                CurrentDb.TableDefs.Refresh

                    ' ** An enforceable relationship will not be created since the user may put in whatever they like.
                    ' ** The table merely contains often-used items.

                    ' ** Create new relationships.
3450                If blnFound = False Then
3460                  Set Rel = .CreateRelation("RecurringItemsledger", "RecurringItems", "ledger", dbRelationDontEnforce)
3470                  With Rel
3480                    .Fields.Append .CreateField("RecurringItem", dbText)
3490                    .Fields![RecurringItem].ForeignName = "RecurringItem"
3500                  End With
3510                  .Relations.Append Rel
3520                  Set Rel = .CreateRelation("RecurringItemsjournal", "RecurringItems", "journal", dbRelationDontEnforce)
3530                  With Rel
3540                    .Fields.Append .CreateField("RecurringItem", dbText)
3550                    .Fields![RecurringItem].ForeignName = "RecurringItem"
3560                  End With
3570                  .Relations.Append Rel
3580                  .Relations.Refresh

3590                  CurrentDb.TableDefs.Refresh
3600                  CurrentDb.TableDefs("RecurringItems").RefreshLink
3610                End If

3620                If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
                      ' ** Empty tblTemplate_RecurringItems, for all but 2 standard entries.
3630                  Set qdf = CurrentDb.QueryDefs("qryRecurringItems_02")
3640                  qdf.Execute
                      ' ** Reset the Autonumber field.
3650                  ChangeSeed_Ext "tblTemplate_RecurringItems"  ' ** Module Function: modAutonumberFieldFuncs.
3660                End If

3670              Else
                    ' ** Make sure it's linked.
3680                blnFound = False
3690                For Each tdf In CurrentDb.TableDefs
3700                  With tdf
3710                    If .Name = "RecurringType" Then
3720                      blnFound = True
3730                      Exit For
3740                    End If
3750                  End With
3760                Next
3770                If blnFound = False Then
                      ' ** Relink RecurringType back to here.
3780                  If TableExists("RecurringType", True, gstrFile_DataName) Then  ' ** Module Function: modFileUtilities.
3790                    DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                          acTable, "RecurringType", "RecurringType"
3800                    CurrentDb.TableDefs.Refresh
3810                  End If
3820                End If
3830              End If

3840              .Close
3850            End With

3860          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 7. Update Location table, with new '{Unassigned}' record, and new
              ' **         Username, DateCreated, and DateModified fields, then update
              ' **         Ledger, Journal, ActiveAssets, and LedgerArchive for new Location_ID.
              ' *****************************************************************************

3870          If blnSkip2200 = False Then

3880            DoCmd.Hourglass True
3890            frm.WaitMsg_lbl.Caption = UPD_MSG & "7" & strSteps
3900            DoEvents

3910            Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
3920            With dbs

3930              blnFound = False
3940              Set tdf = .TableDefs("Location")
3950              With tdf
3960                For Each fld In .Fields
3970                  With fld
3980                    If .Name = "Username" Then
3990                      blnFound = True
4000                      Exit For
4010                    End If
4020                  End With
4030                Next
4040              End With

4050              If blnFound = False Then

4060                strUpdates = strUpdates & "D7"

                    ' ** Empty the backend copy, tmpLocation.
4070                Set qdf = CurrentDb.QueryDefs("qryLocation_21")
4080                qdf.Execute
                    ' ** Empty the 2nd backend copy, tmpLocation2.
4090                Set qdf = CurrentDb.QueryDefs("qryLocation_22")
4100                qdf.Execute

4110                blnSkip = True
4120                If blnSkip = False Then
                      ' ** Make sure there's a journal_USER for every Ledger record.
                      ' ** Also clean up, as appropriate, those saying 'TADemo' or superuser.
4130                  If InStr(gstrTrustDataLocation, gstrDir_DevDemo) = 0 Then
4140                    strTmp01 = AppVersion_GetDta  ' ** Module Function: modAppVersionFuncs.
4150                    If Left(strTmp01, 1) = "#" Then
                          ' ** Doesn't have AppVersion property.
                          ' ** Citizens: Update Ledger, for journal_USER = Null, 'TADemo', 'superuser', set = 'System'.
                          'Set qdf = CurrentDb.QueryDefs("qryLocation_25k")
4160                    Else
                          ' ** Does have it.
4170                      If Right(strTmp01, 1) <> "d" Then
                            ' ** It's not a current demo version.
                            ' ** Citizens: Update Ledger, for journal_USER = Null, 'TADemo', 'superuser', set = 'System'.
                            'Set qdf = CurrentDb.QueryDefs("qryLocation_25k")
4180                      Else
                            ' ** It is a demo version.
                            ' ** Citizens: Update Ledger, for journal_USER = Null,  set = 'TADemo'.
                            'Set qdf = CurrentDb.QueryDefs("qryLocation_25l")
4190                      End If
4200                    End If
4210                  Else
                        ' ** It's a demo version.
                        ' ** Citizens: Update Ledger, for journal_USER = Null,  set = 'TADemo'.
                        'Set qdf = CurrentDb.QueryDefs("qryLocation_25l")
4220                  End If
4230                  qdf.Execute
4240                End If  ' ** blnSkip.

                    ' ** Append Location to tmpLocation.
4260                Set qdf = CurrentDb.QueryDefs("qryLocation_23")
4270                qdf.Execute

4280                DoEvents

                    ' ** Update tmpLocation with best guess for Username, DateCreated, DateModified.
4290                Set qdf = CurrentDb.QueryDefs("qryLocation_27")
4300                qdf.Execute

                    ' ** If on demo data (has accountno = 11, 'William B. Johnson Trust'),
                    ' ** check and change Ledger entry with Location_ID = 2,to 7 before moving on.
                    ' ** On user's machines, this will do nothing.
4310                Set qdf = CurrentDb.QueryDefs("qryLocation_28b")
4320                qdf.Execute

                    ' ** Grouped Location_ID's in Ledger, not in tmpLocation; Locations no longer present.
4330                Set qdf = CurrentDb.QueryDefs("qryLocation_28")
4340                Set rst = qdf.OpenRecordset
4350                If rst.BOF = True And rst.EOF = True Then
                      ' ** All present and accounted for.
4360                  rst.Close
4370                Else
4380                  rst.Close
                      ' ** Append above missing locations to tmpLocation, with unique names.
4390                  Set qdf = CurrentDb.QueryDefs("qryLocation_29")
4400                  qdf.Execute
4410                End If

4420                DoEvents

                    ' ** Append new record for Location_ID = 0 to tmpLocation; 1.
4430                Set qdf = CurrentDb.QueryDefs("qryLocation_30")
4440                qdf.Execute

                    ' ** Delete the current frontend's link to Location in TrustDta.mdb.
4450                TableDelete "Location"  ' ** Module Function: modFileUtilities.

                    ' ** Delete their copy of Location table.
4460  On Error Resume Next
4470                .TableDefs.Delete "Location"
                    ' ** If it's already got relationships, it shouldn't have gotten here!
4480                If ERR.Number <> 0 Then
                      ' ** Append new record to tblErrorLog, by specified [frmnam], [fnc], [errnum], [errmsg].
4490                  If blnFoundErrLineNum = True Then
4500                    Set qdf = dbs.QueryDefs("qryErrLog_03a")
4510                  Else
4520                    Set qdf = dbs.QueryDefs("qryErrLog_03")
4530                  End If
4540                  With qdf.Parameters
4550                    ![frmnam] = THIS_NAME  'OK!
4560                    ![fnc] = THIS_PROC
4570                    ![errnum] = ERR.Number
4580                    If blnFoundErrLineNum = True Then
4590                      ![linnum] = Erl
4600                    End If
4610                    ![errmsg] = ERR.description
4620                  End With
4630                  qdf.Execute
4640  On Error GoTo ERRH
                      ' ** Check for relationships.
4650                  lngTmp04 = .Relations.Count
4660                  blnFound = False
4670                  For lngX = (lngTmp04 - 1&) To 0& Step -1&
4680                    Set Rel = .Relations(lngX)
4690                    If Rel.Table = "Location" Or Rel.ForeignTable = "Location" Then
4700                      blnFound = True
4710                      .Relations.Delete Rel.Name
4720                    End If
4730                  Next
4740                  If blnFound = True Then
4750                    .Relations.Refresh
4760                    .TableDefs.Delete "Location"
4770                    .TableDefs.Refresh
4780                  Else
                        ' ** Don't know what the problem is!
                        ' ** I'll just have to let it crash.
4790                  End If
4800                Else
4810  On Error GoTo ERRH
4820                  .TableDefs.Refresh
4830                End If

4840                DoEvents

                    ' ** Copy the new Location table to their TrustDta.mdb.
4850                DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblTemplate_Location", "Location", True  ' ** Structure Only.
4860                .TableDefs.Refresh

                    ' ** Relink Location back to here.
4870                DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "Location", "Location"
4880                CurrentDb.TableDefs.Refresh
4890                DoEvents

4900                DoCmd.Hourglass True  ' ** Make sure it's still running.

                    ' ** Append tmpLocation back to Location (now new).
4910                Set qdf = CurrentDb.QueryDefs("qryLocation_31")
4920                qdf.Execute

                    ' ** Update tmpLocation with new Location_ID.
4930                Set qdf = CurrentDb.QueryDefs("qryLocation_33")
4940                qdf.Execute

                    ' ** Append tmpLocation to tmpLocation2 (no key to primary key).
4950                Set qdf = CurrentDb.QueryDefs("qryLocation_34")
4960                qdf.Execute

4970                DoEvents

                    ' ** Change DefaultValue in Ledger, Journal, and ActiveAssets.
                    ' ** LedgerArchive doesn't have a DefaultValue.

4980                .TableDefs("ledger").Fields("Location_ID").DefaultValue = "1"
4990                .TableDefs("journal").Fields("Location_ID").DefaultValue = "1"
5000                .TableDefs("ActiveAssets").Fields("Location_ID").DefaultValue = "1"
                    'tblTemplate_ActiveAssets
                    'tblTemplate_Journal
                    'tblTemplate_Ledger

                    ' ** Update Ledger, for new Location_ID.
5010                Set qdf = CurrentDb.QueryDefs("qryLocation_35")
5020                qdf.Execute

                    ' ** Update Journal, for new Location_ID.
5030                Set qdf = CurrentDb.QueryDefs("qryLocation_36")
5040                qdf.Execute

                    ' ** Update ActiveAssets, for new Location_ID.
5050                Set qdf = CurrentDb.QueryDefs("qryLocation_37")
5060                qdf.Execute

                    'LEDGER ARCHIVE!
                    'If there are records in LedgerArchive Then
5070                strUpdates = strUpdates & "A7"
                    ' ** Update LedgerArchive, for new Location_ID.
5080                Set qdf = CurrentDb.QueryDefs("qryLocation_38")
5090                qdf.Execute

5100                DoEvents

                    ' ** Create new relationships.
                    ' ** If I've missed a Location_ID cross-check, it'll show up here!
5110                Set Rel = .CreateRelation("Locationjournal", "Location", "journal", dbRelationUpdateCascade)
5120                With Rel
5130                  .Fields.Append .CreateField("Location_ID", dbLong)
5140                  .Fields![Location_ID].ForeignName = "Location_ID"  'Hate this inconsistency!
5150                End With
5160                .Relations.Append Rel
5170                Set Rel = .CreateRelation("Locationledger", "Location", "ledger", dbRelationUpdateCascade)
5180                With Rel
5190                  .Fields.Append .CreateField("Location_ID", dbLong)
5200                  .Fields![Location_ID].ForeignName = "Location_ID"  'Hate this inconsistency!
5210                End With
5220                .Relations.Append Rel
5230                .Relations.Refresh

5240                CurrentDb.TableDefs.Refresh
5250                CurrentDb.TableDefs("Location").RefreshLink

5260              End If

5270              .Close
5280            End With

5290          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 8. Add the new LedgerHidden and HiddenType tables.
              ' *****************************************************************************

5300          If blnSkip2200 = False Then

5310            DoCmd.Hourglass True  ' ** Make sure it's still running.
5320            frm.WaitMsg_lbl.Caption = UPD_MSG & "8" & strSteps
5330            DoEvents

5340            Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
5350            With dbs

5360              blnFound = False: blnFoundHidType = False: blnFoundHidMatch = False
5370              For Each tdf In .TableDefs
5380                With tdf
5390                  If .Name = "LedgerHidden" Then
5400                    blnFound = True
5410                    For Each fld In .Fields
5420                      With fld
5430                        If .Name = "hidtype" Then
5440                          blnFoundHidType = True
5450                        ElseIf .Name = "hid_newmatch" Then
5460                          blnFoundHidMatch = True
5470                          blnFoundHidMatch = False
5480                          For lngX = 0& To (tdf.Fields.Count - 1&)
5490                            If tdf.Fields(lngX).Name = "hid_order" Then
5500                              blnFoundHidMatch = True
5510                              Exit For
5520                            End If
5530                          Next
5540                        End If
5550                      End With
5560                    Next
5570                    Exit For
5580                  End If
5590                End With
5600              Next

5610              blnFound2 = False
5620              For Each tdf In .TableDefs
5630                With tdf
5640                  If .Name = "HiddenType" Then
5650                    blnFound2 = True
5660                    Exit For
5670                  End If
5680                End With
5690              Next

5700              If blnFound = False Or blnFound2 = False Or (blnFound = True And (blnFoundHidType = False Or blnFoundHidMatch = False)) Then

5710                strUpdates = strUpdates & "D8"

                    ' ** Check in case the user had one of our interim versions (before I changed the field name).
5720                If blnFound = True And (blnFoundHidType = False Or blnFoundHidMatch = False) Then

                      ' ** If LedgerHidden is linked, delete the link.
5730                  If TableExists("LedgerHidden") = True Then
5740                    DoCmd.DeleteObject acTable, "LedgerHidden"
5750                  End If

                      ' ** Check for the HiddenType relationship, and delete it before deleting LedgerHidden.
5760                  blnFound = False: strTmp01 = vbNullString
5770                  For Each Rel In .Relations
5780                    With Rel
5790                      If .Table = "HiddenType" And .ForeignTable = "LedgerHidden" Then
5800                        strTmp01 = .Name
5810                        blnFound = True
5820                      End If
5830                    End With
5840                  Next
5850                  If blnFound = True Then
5860                    .Relations.Delete strTmp01
5870                  End If

                      ' ** Delete their existing LedgerHidden table.
5880                  .TableDefs.Delete "LedgerHidden"
5890                  .TableDefs.Refresh

                      ' ** Now reset blnFound to indicate LedgerHidden wasn't found.
5900                  blnFound = False

5910                End If

                    ' ** Copy the new LedgerHidden table to their TrustDta.mdb.
5920                If blnFound = False Then
5930                  DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                        acTable, "tblTemplate_LedgerHidden", "LedgerHidden", True  ' ** Structure Only.
5940                End If

                    ' ** Copy the new HiddenType table to their TrustDta.mdb.
5950                If blnFound2 = False Then
5960                  DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                        acTable, "tblTemplate_HiddenType", "HiddenType", False  ' ** Copy with data.
5970                End If

5980                .TableDefs.Refresh

                    ' ** Relink LedgerHidden back to here.
5990                If blnFound = False Then
6000                  DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                        acTable, "LedgerHidden", "LedgerHidden"
6010                End If

                    ' ** In case someone's got a version before I moved it to TrustDta.mdb.
6020                For Each tdf In CurrentDb.TableDefs
6030                  With tdf
6040                    If .Name = "HiddenType" And .Connect = vbNullString Then
6050                      DoCmd.DeleteObject acTable, "HiddenType"
6060                      Exit For
6070                    End If
6080                  End With
6090                Next

                    ' ** Relink HiddenType back to here.
6100                If blnFound2 = False Then
6110                  DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                        acTable, "HiddenType", "HiddenType"
6120                End If

6130                CurrentDb.TableDefs.Refresh

                    ' ** Create the new relationship.
6140                Set Rel = .CreateRelation("HiddenTypeLedgerHidden", "HiddenType", "LedgerHidden", dbRelationUpdateCascade)
6150                With Rel
6160                  .Fields.Append .CreateField("hidtype", dbText)
6170                  .Fields![hidtype].ForeignName = "hidtype"
6180                End With
6190                .Relations.Append Rel
6200                .Relations.Refresh

6210              End If

6220              .Close
6230            End With

6240          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 9. Add the new tblPreference_User table.
              ' *****************************************************************************

6250          If blnSkip2200 = False Then

6260            frm.WaitMsg_lbl.Caption = UPD_MSG & "9" & strSteps
6270            DoEvents

6280            Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
6290            With dbs

                  'blnFound = False
                  'For Each tdf In .TableDefs
                  '  With tdf
                  '    If .Name = "tblPreference_User" Then
6300              blnFound = True
                  '      Exit For
                  '    End If
                  '  End With
                  'Next

6310              If blnFound = False Then

6320                strUpdates = strUpdates & "D9"

                    ' ** Delete the local link before adding the table.
6330                For Each tdf In CurrentDb.TableDefs
6340                  With tdf
6350                    If .Name = "tblPreference_User" Then
6360                      DoCmd.DeleteObject acTable, "tblPreference_User"
6370                      Exit For
6380                    End If
6390                  End With
6400                Next

6410                CurrentDb.TableDefs.Refresh

                    ' ** Copy the new tblPreference_User table to their TrustDta.mdb.
6420                DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblTemplate_Preference_User", "tblPreference_User", True  ' ** Structure Only.

                    ' ** Relink tblPreference_User back to here.
6430                DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblPreference_User", "tblPreference_User"

                    ' ** Check for the new LOCAL tblPreference_Control/tblPreference_User relationship.
6440                blnFound = False
6450                For Each Rel In CurrentDb.Relations
6460                  With Rel
6470                    If (.Table = "tblPreference_Control" And .ForeignTable = "tblPreference_User") Or _
                            (.Table = "tblPreference_User" And .ForeignTable = "tblPreference_Control") Then
6480                      blnFound = True
6490                    End If
6500                  End With
6510                Next

                    ' ** Create the new relationship.
6520                If blnFound = False Then
6530                  Set Rel = CurrentDb.CreateRelation("tblPreference_ControltblPreference_User", _
                        "tblPreference_Control", "tblPreference_User", dbRelationDontEnforce)
6540                  With Rel
6550                    .Fields.Append .CreateField("frm_name", dbText)
6560                    .Fields![frm_name].ForeignName = "frm_name"
6570                    .Fields.Append .CreateField("ctl_name", dbText)
6580                    .Fields![ctl_name].ForeignName = "ctl_name"
6590                  End With
6600                  CurrentDb.Relations.Append Rel
6610                  CurrentDb.Relations.Refresh
6620                End If

6630              End If
6640              .Close
6650            End With

6660          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 10. Make sure m_TBL has the new fields and is up-to-date.
              ' *****************************************************************************

6670          If blnSkip2200 = False Then

6680            frm.WaitMsg_lbl.Caption = UPD_MSG & "10" & strSteps
6690            DoEvents

6700            Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
6710            With dbs

6720              blnFound = False: blnFound2 = False
6730              For Each tdf In .TableDefs
6740                With tdf
6750                  If .Name = "m_TBL" Then
6760                    For Each fld In .Fields
6770                      With fld
6780                        If .Name = "mtbl_AUX" Then  ' ** Newest field.
6790                          blnFound = True
6800                        ElseIf .Name = "mtbl_ORDER" Then
6810                          If .Type = dbLong Then
6820                            blnFound2 = True
6830                          End If
6840                        End If
6850                      End With
6860                    Next
6870                    Exit For
6880                  End If
6890                End With
6900              Next

6910              If blnFound = False Or blnFound2 = False Then

6920                strUpdates = strUpdates & "D10"

                    ' ** Delete the local link before adding the table.
6930                For Each tdf In CurrentDb.TableDefs
6940                  With tdf
6950                    If .Name = "m_TBL" Then
6960                      DoCmd.DeleteObject acTable, "m_TBL"
6970                      Exit For
6980                    End If
6990                  End With
7000                Next

7010                CurrentDb.TableDefs.Refresh

                    ' ** Delete their existing m_TBL table.
7020                .TableDefs.Delete "m_TBL"
7030                .TableDefs.Refresh

                    ' ** Copy the new m_TBL table to their TrustDta.mdb.
7040                DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblTemplate_m_TBL", "m_TBL", False  ' ** Copy with data.

                    ' ** Relink m_TBL back to here.
7050                DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "m_TBL", "m_TBL"

7060              Else
                    ' ** Make sure their list is up-to-date.

                    ' ** Delete qrySystemUpdate_11a_m_TBL (m_TBL, not in tblTemplate_m_TBL), from m_TBL.
                    ' ** ####  tblTemplate_m_TBL TAKES PRECEDENCE  ####
7070                Set qdf = CurrentDb.QueryDefs("qrySystemUpdate_11f_m_TBL")
7080                qdf.Execute

                    ' ** Append qrySystemUpdate_11b_m_TBL (tblTemplate_m_TBL, not in m_TBL) to m_TBL.
                    ' ** ####  tblTemplate_m_TBL TAKES PRECEDENCE  ####
7090                Set qdf = CurrentDb.QueryDefs("qrySystemUpdate_11g_m_TBL")
7100                qdf.Execute

                    ' ** Update m_TBL from tblTemplate_m_TBL.
                    ' ** ####  tblTemplate_m_TBL TAKES PRECEDENCE  ####
7110                Set qdf = CurrentDb.QueryDefs("qrySystemUpdate_11h_m_TBL")
7120                qdf.Execute

7130              End If
7140              .Close
7150            End With

7160          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 11. Make sure taxcode in AssetType is dbInteger,
              ' **          and it's linked to the TaxCode table.
              ' *****************************************************************************

7170          If blnSkip2200 = False Then

7180            frm.WaitMsg_lbl.Caption = UPD_MSG & "11" & strSteps
7190            DoEvents

7200            Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
7210            With dbs

7220              blnFound = False
7230              Set tdf = .TableDefs("AssetType")
7240              With tdf
7250                Set fld = .Fields("taxcode")
7260                If fld.Type = dbInteger Then
7270                  blnFound = True
7280                End If
7290              End With
7300              If blnFound = False Then

7310                strUpdates = strUpdates & "D11"

7320                lngRecs = 0&
7330                ReDim arr_varRec(0)

                    ' ** Delete the existing relationship to MasterAsset.
7340                For Each Rel In .Relations
7350                  With Rel
7360                    If .Table = "assettype" Or .ForeignTable = "assettype" Then
7370                      lngRecs = lngRecs + 1&
7380                      ReDim Preserve arr_varRec(lngRecs - 1&)
7390                      arr_varRec(lngRecs - 1&) = .Name
7400                    End If
7410                  End With
7420                Next
7430                If lngRecs > 0& Then
7440                  For lngX = 0& To (lngRecs - 1&)
7450                    .Relations.Delete arr_varRec(lngX)
7460                  Next
7470                End If

                    ' ** Delete their existing assettype table.
7480                .TableDefs.Delete "assettype"
7490                .TableDefs.Refresh

                    ' ** Transfer the new AssetType template.
7500                DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblTemplate_AssetType", "assettype", False  ' ** Copy with data.
7510                .TableDefs.Refresh

                    ' ** Create all the relationships.
7520                Set Rel = .CreateRelation("assettypemasterasset", "assettype", "masterasset", dbRelationUpdateCascade)
7530                With Rel
7540                  .Fields.Append .CreateField("assettype", dbText)  ' ** NOTE:
7550                  .Fields![assettype].ForeignName = "assettype"     ' ** Different name!
7560                End With
7570                .Relations.Append Rel
7580                .Relations.Refresh

                    ' ** Delete their existing TaxCode table.
7590                .TableDefs.Delete "taxcode"
7600                .TableDefs.Refresh

                    ' ** Transfer the new TaxCode template.
7610                DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                      acTable, "tblTemplate_TaxCode", "taxcode", False  ' ** Copy with data.
7620                .TableDefs.Refresh

                    ' ** Create all the relationships.
7630                Set Rel = .CreateRelation("taxcodeassettype", "taxcode", "assettype", dbRelationUpdateCascade)
7640                With Rel
7650                  .Fields.Append .CreateField("taxcode", dbInteger)
7660                  .Fields![taxcode].ForeignName = "taxcode"
7670                End With
7680                .Relations.Append Rel
7690                .Relations.Refresh

7700                CurrentDb.TableDefs.Refresh
7710                CurrentDb.TableDefs("taxcode").RefreshLink

7720              End If

7730              .Close
7740            End With

7750          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 12. Add 2 new Income/Expense codes to the m_REVCODE table.
              ' *****************************************************************************

7760          If blnSkip2200 = False Then

7770            frm.WaitMsg_lbl.Caption = UPD_MSG & "12" & strSteps
7780            DoEvents

7790            Set dbs = CurrentDb
7800            With dbs

7810              lngRevCodes = DCount("[revcode_ID]", "m_REVCODE")
                  'IS THIS MESSING UP THEIR REVCODES?

                  ' ** Check for presence of codes.
7820              blnAddOtherCharges = True: blnAddOtherCredits = True: blnAddOrdinaryDividend = True: blnAddInterestIncome = True
                  ' ** m_REVCODE, for 'OC Other Charges', 'OC Other Credits', 'Ordinary Dividend', 'Interest Income'.
7830              Set qdf = .QueryDefs("qryRevCodes_22")
7840              Set rst = qdf.OpenRecordset
7850              If rst.BOF = True And rst.EOF = True Then
                    ' ** New codes not there; add them.
7860                rst.Close
7870              Else
7880                rst.MoveLast
7890                lngTmp02 = rst.RecordCount
7900                If lngTmp02 = 4 Then
                      ' ** All 4 records already present.
7910                  rst.Close
7920                  blnAddOtherCharges = False: blnAddOtherCredits = False: blnAddOrdinaryDividend = False: blnAddInterestIncome = False
7930                Else
7940                  rst.MoveFirst
7950                  For lngX = 1& To lngTmp02
7960                    Select Case rst![revcode_DESC]
                        Case "OC Other Charges"
7970                      blnAddOtherCharges = False
7980                    Case "OC Other Credits"
7990                      blnAddOtherCredits = False
8000                    Case "Ordinary Dividend"
8010                      blnAddOrdinaryDividend = False
8020                    Case "Interest Income"
8030                      blnAddInterestIncome = False
8040                    End Select
8050                    If lngX < lngTmp02 Then rst.MoveNext
8060                  Next
8070                  rst.Close
8080                End If
8090              End If

8100              .Close
8110            End With

8120            If blnAddOtherCharges = True Or blnAddOtherCredits = True Or blnAddOrdinaryDividend = True Or blnAddInterestIncome = True Then

                  ' ** Make tmpRevCodeEdit table.
8130              blnRetVal = RevCode_Setup(THIS_NAME)  ' ** Module Function: modRevCodeFuncs.  'OK!

                  ' ** Renumber if they've added some of their own Income/Expense codes.
8140              If lngRevCodes > 2& Then  ' ** Leave this at 2 for now.

8150                strUpdates = strUpdates & "D120"

                    ' ** Renumber the m_REVCODE table via tmpRevCodeEdit.
8160                If blnRetVal = True Then
8170                  blnRetVal = RevCode_Renum(THIS_NAME, blnEdited, blnMoveAll)  ' ** Module Function: modRevCodeFuncs.  'OK!
8180                End If

                    ' ** Update the m_REVCODE table from tmpRevCodeEdit.
8190                If blnRetVal = True Then
8200                  blnRetVal = RevCode_Update(THIS_NAME, blnEdited, wrk)  ' ** Module Function: modRevCodeFuncs.  'OK!
8210                End If

8220              End If

8230              If blnRetVal = True Then

8240                Set dbs = CurrentDb
8250                With dbs

8260                  strUpdates = strUpdates & "D121"

                      ' ** Append m_REVCODE_TYPE, linked to qryRevCodes_23a (Maximum
                      ' ** revcode_SORTORDER, Income), just 'OC Other Charges' record, to m_REVCODE.
8270                  If blnAddOtherCharges = True Then
8280                    Set qdf = .QueryDefs("qryRevCodes_24a")
8290                    qdf.Execute
8300                  End If

                      ' ** Append m_REVCODE_TYPE, linked to qryRevCodes_23b (Maximum
                      ' ** revcode_SORTORDER, Expense), just 'OC Other Credits' record, to m_REVCODE.
8310                  If blnAddOtherCredits = True Then
8320                    Set qdf = .QueryDefs("qryRevCodes_24b")
8330                    qdf.Execute
8340                  End If

                      ' ** Append m_REVCODE_TYPE, linked to qryRevCodes_23a (Maximum
                      ' ** revcode_SORTORDER, Income), just 'Ordinary Dividend' record, to m_REVCODE.
8350                  If blnAddOrdinaryDividend = True Then
8360                    Set qdf = .QueryDefs("qryRevCodes_24c")
8370                    qdf.Execute
8380                  End If

                      ' ** Append m_REVCODE_TYPE, linked to qryRevCodes_23a (Maximum
                      ' ** revcode_SORTORDER, Income), just 'Interest Income' record, to m_REVCODE.
8390                  If blnAddInterestIncome = True Then
8400                    Set qdf = .QueryDefs("qryRevCodes_24d")
8410                    qdf.Execute
8420                  End If

8430                  If lngRevCodes > 2& Then

8440                    strUpdates = strUpdates & "D122"

                        ' ** Delete tmpRevCodeEdit.
                        'TableDelete ("tmpRevCodeEdit")  ' ** Module Function: modFileUtilities.

                        ' ** Make tmpRevCodeEdit table (now contains new codes).
8450                    blnRetVal = RevCode_Setup(THIS_NAME)  'OK!

                        ' ** Now, set each OC to revcode_SORTORDER = 2.
                        ' ** tmpRevCodeEdit, for 'OC Other Charges', 'OC Other Credits'.
8460                    Set qdf = .QueryDefs("qryRevCodes_25a")
8470                    Set rst = qdf.OpenRecordset
8480                    With rst
8490                      .MoveFirst
8500                      For lngX = 1& To 2&
8510                        For lngY = 0& To (glngRevOs - 1&)
8520                          If garr_varRevO(REVO_ID, lngY) = ![revcode_ID] Then
8530                            garr_varRevO(REVO_CHANGED, lngY) = True
8540                            Exit For
8550                          End If
8560                        Next
8570                        .Edit
8580                        ![revcode_SORTORDER] = 2
8590                        ![revcode_CHANGED] = True
8600                        .Update
8610                        If lngX < 2& Then .MoveNext
8620                      Next
8630                    End With

                        ' ** tmpRevCodeEdit, for 'Ordinary Dividend', 'Interest Income'.
8640                    Set qdf = .QueryDefs("qryRevCodes_25b")
8650                    Set rst = qdf.OpenRecordset
8660                    With rst
8670                      .MoveFirst
8680                      For lngX = 1& To 2&
8690                        For lngY = 0& To (glngRevOs - 1&)
8700                          If garr_varRevO(REVO_ID, lngY) = ![revcode_ID] Then
8710                            garr_varRevO(REVO_CHANGED, lngY) = True
8720                            Exit For
8730                          End If
8740                        Next
8750                        .Edit
8760                        Select Case ![revcode_DESC]
                            Case "Ordinary Dividend"
8770                          ![revcode_SORTORDER] = 3
8780                        Case "Interest Income"
8790                          ![revcode_SORTORDER] = 4
8800                        End Select
8810                        ![revcode_CHANGED] = True
8820                        .Update
8830                        If lngX < 2& Then .MoveNext
8840                      Next
8850                    End With

8860                  End If

8870                  .Close
8880                End With

8890                If lngRevCodes > 2& Then

8900                  strUpdates = strUpdates & "D123"

                      ' ** Renumber the m_REVCODE table via tmpRevCodeEdit.
8910                  blnRetVal = RevCode_Renum(THIS_NAME, blnEdited, blnMoveAll)  ' ** Module Function: modRevCodeFuncs.  'OK!
                      ' ** Update the m_REVCODE table from tmpRevCodeEdit.
8920                  blnRetVal = RevCode_Update(THIS_NAME, blnEdited, wrk)  ' ** Module Function: modRevCodeFuncs.  'OK!

8930                End If

8940              End If

                  ' ** Delete tmpRevCodeEdit.
                  'TableDelete ("tmpRevCodeEdit")  ' ** Module Function: modFileUtilities.

8950            End If

8960          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 13. Add new 'Personal Property' assettype to AssetType table.
              ' *****************************************************************************

8970          If blnSkip2200 = False Then

8980            frm.WaitMsg_lbl.Caption = UPD_MSG & "13" & strSteps
8990            DoEvents

9000            Set dbs = CurrentDb
9010            With dbs

                  ' ** Check AssetType table.
9020              Set qdf = .QueryDefs("qryAssetType_07e")
9030              Set rst = qdf.OpenRecordset
9040              If rst.BOF = True And rst.EOF = True Then
                    ' ** Entry doesn't exist; append it.
9050                rst.Close
9060                strUpdates = strUpdates & "D13"
9070                Set qdf = .QueryDefs("qryAssetType_08e")
9080                qdf.Execute
9090              Else
                    ' ** Record present.
9100                rst.Close
9110              End If

9120            End With  ' ** dbs

9130          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 14. Add new 'Personal Property' assettype to AssetTypeGrouping table.
              ' *****************************************************************************

9140          If blnSkip2200 = False Then

9150            frm.WaitMsg_lbl.Caption = UPD_MSG & "14" & strSteps
9160            DoEvents

9170            With dbs

                  ' ** Check AssetTypeGrouping table.
9180              Set qdf = .QueryDefs("qryAssetType_09e")
9190              Set rst = qdf.OpenRecordset
9200              If rst.BOF = True And rst.EOF = True Then
                    ' ** Entry doesn't exist; append it.
9210                rst.Close
9220                strUpdates = strUpdates & "D14"
9230                Set qdf = .QueryDefs("qryAssetType_10e")
9240                qdf.Execute
9250              Else
                    ' ** Record present.
9260                rst.Close
9270              End If

9280              .Close
9290            End With  ' ** dbs.

9300          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 15. Check several other miscellaneous relationships.
              ' *****************************************************************************

9310          If blnSkip2200 = False Then

9320            frm.WaitMsg_lbl.Caption = UPD_MSG & "15" & strSteps
9330            DoEvents

9340            Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
9350            With dbs

                  ' ** Check for the new RecurringItems relationships.
9360              blnFound = False: blnFound2 = False
9370              For Each Rel In .Relations
9380                With Rel
9390                  If .Table = "RecurringItems" And .ForeignTable = "ledger" Then
9400                    blnFound = True
9410                  ElseIf .Table = "RecurringItems" And .ForeignTable = "journal" Then
9420                    blnFound2 = True
9430                  End If
9440                End With
9450              Next
9460              If blnFound = False Then
9470                Set Rel = .CreateRelation("RecurringItemsledger", "RecurringItems", "ledger", dbRelationDontEnforce)
9480                With Rel
9490                  .Fields.Append .CreateField("RecurringItem", dbText)
9500                  .Fields![RecurringItem].ForeignName = "RecurringItem"
9510                End With
9520                .Relations.Append Rel
9530                .Relations.Refresh
9540              End If
9550              If blnFound2 = False Then
9560                Set Rel = .CreateRelation("RecurringItemsjournal", "RecurringItems", "journal", dbRelationDontEnforce)
9570                With Rel
9580                  .Fields.Append .CreateField("RecurringItem", dbText)
9590                  .Fields![RecurringItem].ForeignName = "RecurringItem"
9600                End With
9610                .Relations.Append Rel
9620                .Relations.Refresh
9630              End If

9640              CurrentDb.TableDefs.Refresh

9650              .Close
9660            End With

9670          End If  ' ** blnSkip2200.

              ' *****************************************************************************
              ' ** Step 16. Check for presence of m_VA table.
              ' *****************************************************************************

9680          If blnSkip2200 = False Then

9690            frm.WaitMsg_lbl.Caption = UPD_MSG & "16" & strSteps
9700            DoEvents

9710            Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_ArchDataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
9720            With dbs

9730              blnFound = False
9740              For Each tdf In .TableDefs
9750                With tdf
9760                  If .Name = "m_VA" Then
9770                    blnFound = True
9780                    Exit For
9790                  End If
9800                End With
9810              Next

9820              If blnFound = False Then
9830                strUpdates = strUpdates & "A16"
9840                Set tdf = .CreateTableDef("m_VA")
9850                With tdf
9860                  Set fld = .CreateField("va_MAIN", dbInteger)
9870                  fld.Required = True
9880                  tdf.Fields.Append fld
9890                  Set fld = .CreateField("va_MINOR", dbInteger)
9900                  fld.Required = True
9910                  tdf.Fields.Append fld
9920                  Set fld = .CreateField("va_REVISION", dbInteger)
9930                  fld.Required = True
9940                  tdf.Fields.Append fld
9950                  Set fld = .CreateField("va_DE1", dbText, 50)
9960                  fld.AllowZeroLength = False
9970                  tdf.Fields.Append fld
9980                  Set fld = .CreateField("va_DE2", dbText, 50)
9990                  fld.AllowZeroLength = False
10000                 tdf.Fields.Append fld
10010               End With
10020               .TableDefs.Append tdf
10030               .TableDefs.Refresh
10040             End If

10050             .Close
10060           End With  ' ** dbs.

10070         End If  ' ** blnSkip2200.

10080       End With  ' ** wrk.

10090       If blnSkip2200 = False Then

              ' ** Make sure m_VA is linked to here.
10100         Set dbs = CurrentDb
10110         With dbs
10120           blnFound = False
10130           For Each tdf In .TableDefs
10140             With tdf
10150               If .Name = "m_VA" Then
10160                 blnFound = True
10170                 Exit For
10180               End If
10190             End With
10200           Next
10210           If blnFound = False Then
                  ' ** Link m_VA to here.
10220             DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), acTable, "m_VA", "m_VA"
10230             CurrentDb.TableDefs.Refresh
10240             DoEvents
10250           End If
10260           .Close
10270         End With

10280       End If  ' ** blnSkip2200.

            ' *****************************************************************************
            ' ** Step 17. Check other miscellaneous items.
            ' *****************************************************************************

10290       If blnSkip2200 = False Then

10300         frm.WaitMsg_lbl.Caption = UPD_MSG & "17" & strSteps
10310         DoEvents

10320         Set dbs = CurrentDb
10330         With dbs

                ' ** Check if m_VA is linked.
10340           blnFound = False
10350           For Each tdf In .TableDefs
10360             With tdf
10370               If .Name = "m_VA" Then
10380                 blnFound = True
10390                 Exit For
10400               End If
10410             End With
10420           Next
10430           If blnFound = False Then
10440             strUpdates = strUpdates & "A17"
10450             DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_ArchDataName), _
                    acTable, "m_VA", "m_VA"
10460             .TableDefs.Refresh
10470           Else
10480             .TableDefs("m_VA").RefreshLink
10490           End If

                ' ** Get the current TrustDta.mdb version number.
10500           Set rst = .OpenRecordset("m_VD", dbOpenDynaset, dbReadOnly)
10510           With rst
10520             .MoveFirst
10530             intVDMaj = ![vd_MAIN]
10540             intVDMin = ![vd_MINOR]
10550             intVDRev = ![vd_REVISION]
10560             .Close
10570           End With

10580         End With  ' ** dbs.

10590       End If  ' ** blnSkip2200.

            ' *****************************************************************************
            ' ** Step 18. Check if m_VA has its record.
            ' *****************************************************************************

10600       If blnSkip2200 = False Then

10610         frm.WaitMsg_lbl.Caption = UPD_MSG & "18" & strSteps
10620         DoEvents

10630         With dbs

10640           Set rst = .OpenRecordset("m_VA", dbOpenDynaset, dbConsistent)
10650           With rst
10660             If .BOF = True And .EOF = True Then
10670               strUpdates = strUpdates & "A18"
10680               .AddNew
10690               ![va_MAIN] = intVDMaj
10700               ![va_MINOR] = intVDMin
10710               ![va_REVISION] = intVDRev
10720               .Update
10730             End If
10740             .Close
10750           End With

10760         End With  ' ** dbs.

10770       End If  ' ** blnSkip2200.

            ' *****************************************************************************
            ' ** Step 19. Check if Journaltype has 'Cost Adj.'.
            ' *****************************************************************************

10780       If blnSkip2200 = False Then

10790         frm.WaitMsg_lbl.Caption = UPD_MSG & "19" & strSteps
10800         DoEvents

10810         With dbs

                ' ** Check for 'Cost Adj.' journaltype (this check used to be in InitializeTables()).
10820           Set qdf = .QueryDefs("qrySystemUpdate_13_JournalType")
10830           Set rst = qdf.OpenRecordset
10840           With rst
10850             If .EOF And .BOF Then
10860               .AddNew
10870               ![journaltype] = "Cost Adj."
10880               ![description] = "Cost Adjustment"
10890               ![sortOrder] = 12
10900               .Update
10910             End If
10920             .Close
10930           End With

10940           .Close
10950         End With  ' ** dbs.

10960       End If  ' ** blnSkip2200.

            ' *****************************************************************************
            ' ** Step 20. Additional relationship checks. (Don't check its return value.)
            ' *****************************************************************************

10970       If blnSkip2200 = False Then

10980         frm.WaitMsg_lbl.Caption = UPD_MSG & "20" & strSteps
10990         DoEvents

11000         Check_BE_Rels  ' ** Function: Below.

11010       End If  ' ** blnSkip2200.

11020     End If  ' ** blnRunUpdates

          ' ****************************************************************************************************
          ' ****************************************************************************************************
          ' ** Step 2. RUN THESE EVERY TIME!
          ' ** X Step 21. RUN THESE EVERY TIME!
          ' ****************************************************************************************************
          ' ****************************************************************************************************

11030     If IsLoaded(FRM_WAIT, acForm) = True Then  ' ** Module Function: modFileUtilities.
11040       frm.WaitMsg_lbl.Caption = UPD_MSG & "2" & strSteps
11050       DoEvents
11060     End If

          ' *********************************************************************
          ' ** Step 21.1. Make sure the [Location_ID] DefaultValue is "1".
          ' **            Even if gintConvertResponse is negative, I don't
          ' **            think the checks below will cause a problem.
          ' *********************************************************************

11070     If blnSkip2200 = False Then

11080       If intWrkType = 0 Then
11090 On Error Resume Next
11100         Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
11110         If ERR.Number <> 0 Then
11120 On Error GoTo ERRH
11130 On Error Resume Next
11140           Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
11150           If ERR.Number <> 0 Then
11160 On Error GoTo ERRH
11170 On Error Resume Next
11180             Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
11190             If ERR.Number <> 0 Then
11200 On Error GoTo ERRH
11210 On Error Resume Next
11220               Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
11230               If ERR.Number <> 0 Then
11240 On Error GoTo ERRH
11250 On Error Resume Next
11260                 Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
11270                 If ERR.Number <> 0 Then
11280 On Error GoTo ERRH
11290 On Error Resume Next
11300                   Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
11310                   If ERR.Number <> 0 Then
11320 On Error GoTo ERRH
11330 On Error Resume Next
11340                     Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
11350 On Error GoTo ERRH
11360                     intWrkType = 7
11370                   Else
11380 On Error GoTo ERRH
11390                     intWrkType = 6
11400                   End If
11410                 Else
11420 On Error GoTo ERRH
11430                   intWrkType = 5
11440                 End If
11450               Else
11460 On Error GoTo ERRH
11470                 intWrkType = 4
11480               End If
11490             Else
11500 On Error GoTo ERRH
11510               intWrkType = 3
11520             End If
11530           Else
11540 On Error GoTo ERRH
11550             intWrkType = 2
11560           End If
11570         Else
11580 On Error GoTo ERRH
11590           intWrkType = 1
11600         End If
11610       Else
11620         Select Case intWrkType
              Case 1
11630           Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
11640         Case 2
11650           Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
11660         Case 3
11670           Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
11680         Case 4
11690           Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
11700         Case 5
11710           Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
11720         Case 6
11730           Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
11740         Case 7
11750           Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
11760         End Select
11770       End If
11780       Set dbs = wrk.OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
11790       With dbs
11800         If .TableDefs("journal").Fields("Location_ID").DefaultValue <> "1" Then
11810           .TableDefs("journal").Fields("Location_ID").DefaultValue = "1"
11820         End If
11830         If .TableDefs("ledger").Fields("Location_ID").DefaultValue <> "1" Then
11840           .TableDefs("ledger").Fields("Location_ID").DefaultValue = "1"
11850         End If
11860         If .TableDefs("ActiveAssets").Fields("Location_ID").DefaultValue <> "1" Then
11870           .TableDefs("ActiveAssets").Fields("Location_ID").DefaultValue = "1"
11880         End If
11890         .Close
11900       End With
11910       wrk.Close

11920     End If  ' ** blnSkip2200.

11930     Set dbs = CurrentDb
11940     With dbs

            ' *********************************************************************
            ' ** Step 21.2. Make sure InvestmentObjective is linked.
            ' *********************************************************************

11950       If blnSkip2200 = False Then

11960         blnFound = False
11970         For Each tdf In CurrentDb.TableDefs
11980           With tdf
11990             If .Name = "InvestmentObjective" Then
12000               blnFound = True
12010               Exit For
12020             End If
12030           End With
12040         Next
12050         If blnFound = False Then
                ' ** Relink InvestmentObjective back to here.
12060           DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "InvestmentObjective", "InvestmentObjective"
12070           CurrentDb.TableDefs.Refresh
12080         End If

12090       End If  ' ** blnSkip2200.

            ' *********************************************************************
            ' ** Step 21.3. Make sure RecurringType is linked.
            ' *********************************************************************

12100       If blnSkip2200 = False Then

12110         blnFound = False
12120         CurrentDb.TableDefs.Refresh
12130         For Each tdf In CurrentDb.TableDefs
12140           With tdf
12150             If .Name = "RecurringType" Then
12160               blnFound = True
12170               Exit For
12180             End If
12190           End With
12200         Next
12210         If blnFound = False Then
                ' ** Relink RecurringType back to here.
12220           If TableExists("RecurringType", True, gstrFile_DataName) Then  ' ** Module Function: modFileUtilities.
12230             DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                    acTable, "RecurringType", "RecurringType"
12240             CurrentDb.TableDefs.Refresh
12250           End If
12260         End If

              ' ** These seem to be getting linked multiple times!
12270         For lngX = 1& To 10&
12280           TableDelete ("RecurringType" & CStr(lngX))  ' ** Module Function: modFileUtilities.
12290         Next

12300       End If  ' ** blnSkip2200.

            ' *********************************************************************
            ' ** Step 21.4. Make sure m_VA is linked.
            ' *********************************************************************

12310       If blnSkip2200 = False Then

12320         blnFound = False
12330         For Each tdf In CurrentDb.TableDefs
12340           With tdf
12350             If .Name = "m_VA" Then
12360               blnFound = True
12370               Exit For
12380             End If
12390           End With
12400         Next
12410         If blnFound = False Then
                ' ** Relink m_VA to here.
12420           DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_ArchDataName), _
                  acTable, "m_VA", "m_VA"
12430           CurrentDb.TableDefs.Refresh
12440         End If

12450       End If  ' ** blnSkip2200.

            ' ** Delete any errant db1.mdb's in this directory, and/or the \Database directory if it's local.
12460       Del_DB1s  ' ** Function: Below.

            ' *********************************************************************
            ' ** Step 3. Check for Null icash, pcash, or cost in Account table.
            ' ** X Step 21.5. Check for Null icash, pcash, or cost in Account table.
            ' *********************************************************************

12470       If IsLoaded(FRM_WAIT, acForm) = True Then  ' ** Module Function: modFileUtilities.
12480         frm.WaitMsg_lbl.Caption = UPD_MSG & "3" & strSteps
12490         DoEvents
12500       End If

            ' ** Account, with Null icash, pcash, or cost.
12510       Set qdf = .QueryDefs("qryAccount_01")
12520       Set rst = qdf.OpenRecordset
12530       If rst.BOF = True And rst.EOF = True Then
              ' ** Everything's fine.
12540         rst.Close
12550       Else
12560         rst.Close
              ' ** Update Account, with DLookups() to qryAccount_05, just discrepancies.
12570         Set qdf = .QueryDefs("qryAccount_07")
12580         qdf.Execute
12590       End If

            ' *********************************************************************
            ' ** Step 4. Check ActiveAssets for priceperunit = 0.
            ' ** X Step 21.6. Check ActiveAssets for priceperunit = 0.
            ' *********************************************************************

12600       If IsLoaded(FRM_WAIT, acForm) = True Then  ' ** Module Function: modFileUtilities.
12610         frm.WaitMsg_lbl.Caption = UPD_MSG & "4" & strSteps
12620         DoEvents
12630       End If

            ' ** Update qryActiveAssets_01 (ActiveAssets, where priceperunit = 0, with priceperunit_new).
12640       Set qdf = .QueryDefs("qryActiveAssets_02")
12650       qdf.Execute

            ' *********************************************************************
            ' ** Step 5. Give AssetType's '80' and '81' both Dividend and Interest.
            ' ** X Step {NEW}  NO!!
            ' *********************************************************************

12660       If IsLoaded(FRM_WAIT, acForm) = True Then  ' ** Module Function: modFileUtilities.
12670         frm.WaitMsg_lbl.Caption = UPD_MSG & "5" & strSteps
12680         DoEvents
12690       End If

            ' ** Update AssetType, just '80', '81', for Interest/Dividend = True.
            'Set qdf = .QueryDefs("qryAssetType_11")
            'qdf.Execute

            ' ** ActiveAssets, just extra spaces in accountno, with accountno_new
12700       Set qdf = .QueryDefs("qryActiveAssets_08_01")
12710       Set rst = qdf.OpenRecordset
12720       If rst.BOF = True And rst.EOF = True Then
              ' ** No problems.
12730         rst.Close
12740         Set rst = Nothing
12750         Set qdf = Nothing
12760       Else
12770         rst.Close
12780         Set rst = Nothing
12790         Set qdf = Nothing
              ' ** Update qryActiveAssets_08_01 (ActiveAssets, just
              ' ** extra spaces in accountno, with accountno_new).
12800         Set qdf = .QueryDefs("qryActiveAssets_08_02")
12810         qdf.Execute
12820         Set qdf = Nothing
12830       End If

12840       .Close
12850     End With

          ' *****************************************************************************
          ' ** Step 6. Save update string if changes were made.
          ' ** X Step 22. Save update string if changes were made.
          ' *****************************************************************************

12860     If IsLoaded(FRM_WAIT, acForm) = True Then  ' ** Module Function: modFileUtilities.
12870       frm.WaitMsg_lbl.Caption = UPD_MSG & "6" & strSteps
12880       DoEvents
12890     End If

12900     If strUpdates <> strTmp01 Or blnTmp05 = True Then
12910       blnRetVal = ChkUpdate(False, lngTmp02)  ' ** Function: Below.
12920     End If
          '1.1. strOldDtaVer = ''
          '1.2. strOldDtaVer = '2.1.41.'
          '2.1. strOldDtaVer = '2.1.41.'
          '2.2. strOldDtaVer = ''

12930     If IsLoaded(FRM_WAIT, acForm) = True Then DoCmd.Close acForm, FRM_WAIT  ' ** Module Function: modFileUtilities.
12940     SysCmd acSysCmdClearStatus
12950     DoEvents

12960   End If  ' ** gstrTrustDataLocation <> vbNullString.

EXITP:
        'DoCmd.Hourglass False
12970   Application.Echo True  ' ** Though it shouldn't be otherwise.
12980   If IsLoaded(FRM_WAIT, acForm) = True Then DoCmd.Close acForm, FRM_WAIT  ' ** Module Function: modFileUtilities.
12990   SysCmd acSysCmdClearStatus
13000   Set Rel = Nothing
13010   Set idx = Nothing
13020   Set fld = Nothing
13030   Set qdf = Nothing
13040   Set tdf = Nothing
13050   Set rst = Nothing
13060   Set dbs = Nothing
13070   Set wrk = Nothing
13080   Set frm = Nothing
13090   Backend_Update = blnRetVal
13100   Exit Function

ERRH:
13110   blnRetVal = False
13120   DoCmd.Hourglass False
13130   Select Case ERR.Number
        Case Else
13140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13150   End Select
13160   Resume EXITP

End Function

Public Function Check_BE_Rels() As Boolean
' ** Backend Relationships.

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "Check_BE_Rels"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst As DAO.Recordset
        Dim Rel As DAO.Relation, tdf As DAO.TableDef, fld As DAO.Field
        Dim doc As DAO.Document, prp As Object
        Dim lngRels As Long, arr_varRel() As Variant
        Dim strLastTable As String, strLastFTable As String
        Dim strLastField As String, strLastFField As String
        Dim lngErrs As Long, arr_varErr() As Variant
        Dim lngRecs As Long
        Dim blnFound As Boolean
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varRel().
        Const R_ELEMS As Integer = 6  ' ** Array's first-element UBound().
        Const R_NAM  As Integer = 0
        Const R_TBL  As Integer = 1
        Const R_FTBL As Integer = 2
        Const R_FLD  As Integer = 3
        Const R_FFLD As Integer = 4
        Const R_DEL  As Integer = 5
        Const R_ATTR As Integer = 6

        ' ** Array: arr_varErr().
        Const E_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const E_OBJ  As Integer = 0
        Const E_DESC As Integer = 1
        Const E_MISC As Integer = 2
        Const E_SKIP As Integer = 3

13210   blnRetVal = True

13220   DoCmd.Hourglass True  ' ** Make sure it's still running.
13230   DoEvents

13240   lngErrs = 0&
13250   ReDim arr_varErr(E_ELEMS, 0)

13260 On Error Resume Next
13270   Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
13280   If ERR.Number <> 0 Then
13290 On Error GoTo ERRH
13300 On Error Resume Next
13310     Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
13320     If ERR.Number <> 0 Then
13330 On Error GoTo ERRH
13340 On Error Resume Next
13350       Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
13360       If ERR.Number <> 0 Then
13370 On Error GoTo ERRH
13380 On Error Resume Next
13390         Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
13400         If ERR.Number <> 0 Then
13410 On Error GoTo ERRH
13420 On Error Resume Next
13430           Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
13440           If ERR.Number <> 0 Then
13450 On Error GoTo ERRH
13460 On Error Resume Next
13470             Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
13480             If ERR.Number <> 0 Then
13490 On Error GoTo ERRH
13500 On Error Resume Next
13510               Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
13520 On Error GoTo ERRH
13530             Else
13540 On Error GoTo ERRH
13550             End If
13560           Else
13570 On Error GoTo ERRH
13580           End If
13590         Else
13600 On Error GoTo ERRH
13610         End If
13620       Else
13630 On Error GoTo ERRH
13640       End If
13650     Else
13660 On Error GoTo ERRH
13670     End If
13680   Else
13690 On Error GoTo ERRH
13700   End If

13710   With wrk
13720     Set dbs = .OpenDatabase(gstrTrustDataLocation & gstrFile_DataName, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
13730     With dbs

13740       lngRels = 0&
13750       ReDim arr_varRel(R_ELEMS, 0)

            ' ** Get a list of all the relationships in TrustDta.mdb.
13760       For Each Rel In .Relations
13770         With Rel
13780           lngRels = lngRels + 1&
13790           lngE = lngRels - 1&
13800           ReDim Preserve arr_varRel(R_ELEMS, lngE)
13810           arr_varRel(R_NAM, lngE) = .Name
13820           arr_varRel(R_TBL, lngE) = .Table
13830           arr_varRel(R_FTBL, lngE) = .ForeignTable
13840           arr_varRel(R_FLD, lngE) = .Fields(0).Name  ' ** I know all these relationships only involve a single field.
13850           arr_varRel(R_FFLD, lngE) = .Fields(0).ForeignName
13860           arr_varRel(R_DEL, lngE) = CBool(False)
13870           arr_varRel(R_ATTR, lngE) = .Attributes
13880         End With
13890       Next

            ' ** Check for duplicates.
13900       strLastTable = vbNullString: strLastFTable = vbNullString
13910       strLastField = vbNullString: strLastFField = vbNullString
13920       For lngX = 0& To (lngRels - 1&)
13930         If arr_varRel(R_DEL, lngX) = False Then
                ' ** Skip if it's already marked to delete.
13940           strLastTable = arr_varRel(R_TBL, lngX)
13950           strLastFTable = arr_varRel(R_FTBL, lngX)
13960           strLastField = arr_varRel(R_FLD, lngX)
13970           strLastFField = arr_varRel(R_FFLD, lngX)
13980           For lngY = (lngX + 1&) To (lngRels - 1&)
13990             If arr_varRel(R_TBL, lngY) = strLastTable And arr_varRel(R_FTBL, lngY) = strLastFTable And _
                      arr_varRel(R_FLD, lngY) = strLastField And arr_varRel(R_FFLD, lngY) = strLastFField Then
14000               arr_varRel(R_DEL, lngY) = CBool(True)
14010             End If
14020           Next
14030         End If
14040       Next

            ' ** Delete any duplicates found.
14050       For lngX = 0& To (lngRels - 1&)
14060         If arr_varRel(R_DEL, lngX) = True Then
14070           .Relations.Delete arr_varRel(R_NAM, lngX)
14080         End If
14090       Next

            ' ** Delete any remaining feefreq relationships.
14100       For lngX = 0& To (lngRels - 1&)
14110         If (arr_varRel(R_TBL, lngX) = "feefreq" Or arr_varRel(R_FTBL, lngX) = "feefreq") And _
                  arr_varRel(R_DEL, lngX) = False Then
14120           .Relations.Delete arr_varRel(R_NAM, lngX)
14130         End If
14140       Next

            ' ** DbRelation enumeration:
            ' **          0  dbRelationEnforce        The relationship is enforced (referential integrity). {my constant}
            ' **          1  dbRelationUnique         The relationship is one-to-one.
            ' **          2  dbRelationDontEnforce    The relationship isn't enforced (no referential integrity).
            ' **          4  dbRelationInherited      The relationship exists in a non-current database that contains the two linked tables.
            ' **        256  dbRelationUpdateCascade  Updates will cascade.
            ' **       4096  dbRelationDeleteCascade  Deletions will cascade.
            ' **   16777216  dbRelationLeft           In Design view, display a LEFT JOIN as the default join type. Microsoft Access only.
            ' **   33554432  dbRelationRight          In Design view, display a RIGHT JOIN as the default join type. Microsoft Access only.

            ' ** accountActiveAssets
            ' **   dbRelationUpdateCascade + dbRelationDeleteCascade
            ' ** accountasset
            ' **   dbRelationUpdateCascade + dbRelationDeleteCascade
            ' ** accountBalance
            ' **   dbRelationUpdateCascade + dbRelationDeleteCascade
            ' ** accountjournal
            ' **   dbRelationUpdateCascade + dbRelationDeleteCascade
            ' ** accountledger
            ' **   dbRelationUpdateCascade + dbRelationDeleteCascade
            ' ** accountPortfolioModel
            ' **   dbRelationUpdateCascade + dbRelationDeleteCascade

            ' *********************************************************************
            ' ** Step 20.1. MasterAsset to ActiveAssets, [assetno].
            ' *********************************************************************

            ' ** Link MasterAsset to ActiveAssets.
14150       blnFound = False
14160       For lngX = 0& To (lngRels - 1&)
14170         If arr_varRel(R_TBL, lngX) = "masterasset" And arr_varRel(R_FTBL, lngX) = "ActiveAssets" Then
14180           blnFound = True
14190           Exit For
14200         End If
14210       Next
14220       If blnFound = False Then

              ' ** Check ActiveAssets against MasterAsset.
14230         Set qdf1 = CurrentDb.QueryDefs("qrySystemUpdate_01_ActiveAssets")
14240         Set qdf2 = .CreateQueryDef("", qdf1.SQL)
14250         Set rst = qdf2.OpenRecordset
14260         With rst
14270           If .BOF = True And .EOF = True Then
                  ' ** All's well.
14280           Else
                  ' ** This is so unlikely, I'd rather just note it and move on.
14290             .MoveLast
14300             lngRecs = .RecordCount
14310             lngErrs = lngErrs + 1&
14320             lngE = lngErrs - 1&
14330             ReDim Preserve arr_varErr(E_ELEMS, lngE)
14340             arr_varErr(E_OBJ, lngE) = "ActiveAssets"
14350             arr_varErr(E_DESC, lngE) = "There are ActiveAssets records without a corresponding record in MasterAsset."
14360             arr_varErr(E_MISC, lngE) = CStr(lngRecs)
14370             arr_varErr(E_SKIP, lngE) = CBool(True)
14380           End If
14390           .Close
14400         End With

              ' ** Check if ActiveAssets has Double or Long assetno.
14410         If .TableDefs("ActiveAssets").Fields("assetno").Type <> dbLong Then

                ' ** Delete the current ActiveAssets link.
14420           TableDelete "ActiveAssets"  ' ** Module Function: modFileUtilities.

                ' ** Delete the backend copy if it's hangin' around.
14430           TableDelete "tmp_ActiveAssets"  ' ** Module Function: modFileUtilities.

                ' ** Bring over a copy of their ActiveAssets table as a backup.
14440           DoCmd.TransferDatabase acImport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "ActiveAssets", "tmp_ActiveAssets", False

                ' ** Empty tblTemplate_ActiveAssets.
14450           Set qdf1 = CurrentDb.QueryDefs("qrySystemUpdate_02_Template_ActiveAssets")
14460           qdf1.Execute

                ' ** Append tmp_ActiveAssets to tblTemplate_ActiveAssets, with all their original data.
14470           Set qdf1 = CurrentDb.QueryDefs("qrySystemUpdate_03_tmp_ActiveAssets")
14480           qdf1.Execute

                ' ** Delete ActiveAssets' relationship to Account.
14490           For lngX = 0& To (lngRels - 1&)
14500             If arr_varRel(R_TBL, lngX) = "account" And arr_varRel(R_FTBL, lngX) = "ActiveAssets" Then
14510               For Each Rel In .Relations
14520                 With Rel
14530                   If .Name = arr_varRel(R_NAM, lngX) Then
14540                     dbs.Relations.Delete arr_varRel(R_NAM, lngX)
14550                     dbs.Relations.Refresh
14560                     Exit For
14570                   End If
14580                 End With
14590               Next
14600             End If
14610           Next

                ' ** Delete their copy of ActiveAssets table.
14620           .TableDefs.Delete "ActiveAssets"

                ' ** Copy the new ActiveAssets table to their TrustDta.mdb.
14630           DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "tblTemplate_ActiveAssets", "ActiveAssets", False  ' ** Don't just copy structure, copy with data.

14640           .TableDefs.Refresh

                ' ** NEW CASCADE DELETES!!!
                ' ** Re-create the relationship.
14650           Set Rel = .CreateRelation("accountActiveAssets", "account", "ActiveAssets", dbRelationUpdateCascade + dbRelationDeleteCascade)
14660           With Rel
14670             .Fields.Append .CreateField("accountno", dbLong)
14680             .Fields![accountno].ForeignName = "accountno"
14690           End With
14700           .Relations.Append Rel
14710           .Relations.Refresh

                ' ** Relink ActiveAssets back to here.
14720           DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "ActiveAssets", "ActiveAssets"

                ' *********************************************************************
                ' ** Step 20.2. MasterAsset to Asset, [assetno].
                ' *********************************************************************

                ' ** Delete the current Asset link.
14730           TableDelete "asset"  ' ** Module Function: modFileUtilities.

                ' ** Delete Asset's relationship to Account.
14740           For lngX = 0& To (lngRels - 1&)
14750             If arr_varRel(R_TBL, lngX) = "account" And arr_varRel(R_FTBL, lngX) = "asset" Then
14760               For Each Rel In .Relations
14770                 With Rel
14780                   If .Name = arr_varRel(R_NAM, lngX) Then
14790                     dbs.Relations.Delete arr_varRel(R_NAM, lngX)
14800                     dbs.Relations.Refresh
14810                     Exit For
14820                   End If
14830                 End With
14840               Next
14850             End If
14860           Next

                ' ** Delete their copy of Asset table.
14870           .TableDefs.Delete "asset"

                ' ** Copy the new Asset table to their TrustDta.mdb.
14880           DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "tblTemplate_Asset", "Asset", True  ' ** Only copy structure.

14890           .TableDefs.Refresh

                ' ** NEW CASCADE DELETES!!!
                ' ** Re-create the relationship.
14900           Set Rel = .CreateRelation("accountasset", "account", "asset", dbRelationUpdateCascade + dbRelationDeleteCascade)
14910           With Rel
14920             .Fields.Append .CreateField("accountno", dbLong)
14930             .Fields![accountno].ForeignName = "accountno"
14940           End With
14950           .Relations.Append Rel
14960           .Relations.Refresh

                ' ** Relink Asset back to here.
14970           DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "asset", "asset"

14980         End If

              ' *********************************************************************
              ' ** Step 20.3. Journal [assetno].
              ' *********************************************************************

              ' ** Check if Journal has Double or Long assetno.
14990         If .TableDefs("journal").Fields("assetno").Type <> dbLong Then

                ' ** Delete the current Journal link.
15000           TableDelete "journal"  ' ** Module Function: modFileUtilities.

                ' ** Delete the backend copy if it's hangin' around.
15010           TableDelete "tmp_Journal"  ' ** Module Function: modFileUtilities.

                ' ** Bring over a copy of their Journal table as a backup.
15020           DoCmd.TransferDatabase acImport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "journal", "tmp_Journal", False

                ' ** Empty tblTemplate_Journal.
15030           Set qdf1 = CurrentDb.QueryDefs("qrySystemUpdate_04_Template_Journal")
15040           qdf1.Execute

                ' ** Append tmp_Journal to tblTemplate_Journal, with all their original data.
15050           Set qdf1 = CurrentDb.QueryDefs("qrySystemUpdate_05_tmp_Journal")
15060           qdf1.Execute

                ' ** Delete Journal's relationship to Account, m_REVCODE, Location, RecurringItems.
15070           For lngX = 0& To (lngRels - 1&)
15080             If arr_varRel(R_TBL, lngX) = "account" And arr_varRel(R_FTBL, lngX) = "journal" Then
15090               For Each Rel In .Relations
15100                 With Rel
15110                   If .Name = arr_varRel(R_NAM, lngX) Then
15120                     dbs.Relations.Delete arr_varRel(R_NAM, lngX)
15130                     Exit For
15140                   End If
15150                 End With
15160               Next
15170             ElseIf arr_varRel(R_TBL, lngX) = "m_REVCODE" And arr_varRel(R_FTBL, lngX) = "journal" Then
15180               For Each Rel In .Relations
15190                 With Rel
15200                   If .Name = arr_varRel(R_NAM, lngX) Then
15210                     dbs.Relations.Delete arr_varRel(R_NAM, lngX)
15220                     Exit For
15230                   End If
15240                 End With
15250               Next
15260             ElseIf arr_varRel(R_TBL, lngX) = "Location" And arr_varRel(R_FTBL, lngX) = "journal" Then
15270               For Each Rel In .Relations
15280                 With Rel
15290                   If .Name = arr_varRel(R_NAM, lngX) Then
15300                     dbs.Relations.Delete arr_varRel(R_NAM, lngX)
15310                     Exit For
15320                   End If
15330                 End With
15340               Next
15350             ElseIf arr_varRel(R_TBL, lngX) = "RecurringItems" And arr_varRel(R_FTBL, lngX) = "journal" Then
15360               For Each Rel In .Relations
15370                 With Rel
15380                   If .Name = arr_varRel(R_NAM, lngX) Then
15390                     dbs.Relations.Delete arr_varRel(R_NAM, lngX)
15400                     Exit For
15410                   End If
15420                 End With
15430               Next
15440             End If
15450           Next
15460           .Relations.Refresh

                ' ** Delete their copy of Journal table.
15470           .TableDefs.Delete "journal"

                ' ** Copy the new Journal table to their TrustDta.mdb.
15480           DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "tblTemplate_Journal", "journal", False  ' ** Don't just copy structure, copy with data.

15490           .TableDefs.Refresh

                ' ** NEW CASCADE DELETES!!!
                ' ** Re-create the relationships.
15500           Set Rel = .CreateRelation("accountjournal", "account", "journal", dbRelationUpdateCascade + dbRelationDeleteCascade)
15510           With Rel
15520             .Fields.Append .CreateField("accountno", dbLong)
15530             .Fields![accountno].ForeignName = "accountno"
15540           End With
15550           .Relations.Append Rel
15560           Set Rel = .CreateRelation("m_REVCODEjournal", "m_REVCODE", "journal", dbRelationUpdateCascade)
15570           With Rel
15580             .Fields.Append .CreateField("revcode_ID", dbLong)
15590             .Fields![revcode_ID].ForeignName = "revcode_ID"
15600           End With
15610           .Relations.Append Rel
15620           Set Rel = .CreateRelation("Locationjournal", "Location", "journal", dbRelationUpdateCascade)
15630           With Rel
15640             .Fields.Append .CreateField("Location_ID", dbLong)
15650             .Fields![Location_ID].ForeignName = "Location_ID"
15660           End With
15670           .Relations.Append Rel
15680           Set Rel = .CreateRelation("RecurringItemsjournal", "RecurringItems", "journal", dbRelationDontEnforce)
15690           With Rel
15700             .Fields.Append .CreateField("RecurringItem", dbLong)
15710             .Fields![RecurringItem].ForeignName = "RecurringItem"
15720           End With
15730           .Relations.Append Rel
15740           .Relations.Refresh

                ' ** Relink Journal back to here.
15750           DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "journal", "journal"

15760         End If

              ' *********************************************************************
              ' ** Step 20.4. Ledger [assetno].
              ' *********************************************************************

              ' ** Check if Ledger has Double or Long assetno.
15770         If .TableDefs("ledger").Fields("assetno").Type <> dbLong Then

                ' ** Delete the current Ledger link.
15780           TableDelete "ledger"  ' ** Module Function: modFileUtilities.

                ' ** Delete the backend copy if it's hangin' around.
15790           TableDelete "tmp_Ledger"  ' ** Module Function: modFileUtilities.

                ' ** Bring over a copy of their Ledger table as a backup.
15800           DoCmd.TransferDatabase acImport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "ledger", "tmp_Ledger", False

                ' ** Empty tblTemplate_Ledger.
15810           Set qdf1 = CurrentDb.QueryDefs("qrySystemUpdate_06_Template_Ledger")
15820           qdf1.Execute

                ' ** Append tmp_Ledger to tblTemplate_Ledger, with all their original data.
15830           Set qdf1 = CurrentDb.QueryDefs("qrySystemUpdate_07_tmp_Ledger")
15840           qdf1.Execute

                ' ** Delete Ledger's relationship to Account, m_REVCODE, Location, RecurringItems.
15850           For lngX = 0& To (lngRels - 1&)
15860             If arr_varRel(R_TBL, lngX) = "account" And arr_varRel(R_FTBL, lngX) = "ledger" Then
15870               For Each Rel In .Relations
15880                 With Rel
15890                   If .Name = arr_varRel(R_NAM, lngX) Then
15900                     dbs.Relations.Delete arr_varRel(R_NAM, lngX)
15910                     Exit For
15920                   End If
15930                 End With
15940               Next
15950             ElseIf arr_varRel(R_TBL, lngX) = "m_REVCODE" And arr_varRel(R_FTBL, lngX) = "ledger" Then
15960               For Each Rel In .Relations
15970                 With Rel
15980                   If .Name = arr_varRel(R_NAM, lngX) Then
15990                     dbs.Relations.Delete arr_varRel(R_NAM, lngX)
16000                     Exit For
16010                   End If
16020                 End With
16030               Next
16040             ElseIf arr_varRel(R_TBL, lngX) = "Location" And arr_varRel(R_FTBL, lngX) = "ledger" Then
16050               For Each Rel In .Relations
16060                 With Rel
16070                   If .Name = arr_varRel(R_NAM, lngX) Then
16080                     dbs.Relations.Delete arr_varRel(R_NAM, lngX)
16090                     Exit For
16100                   End If
16110                 End With
16120               Next
16130             ElseIf arr_varRel(R_TBL, lngX) = "RecurringItems" And arr_varRel(R_FTBL, lngX) = "ledger" Then
16140               For Each Rel In .Relations
16150                 With Rel
16160                   If .Name = arr_varRel(R_NAM, lngX) Then
16170                     dbs.Relations.Delete arr_varRel(R_NAM, lngX)
16180                     Exit For
16190                   End If
16200                 End With
16210               Next
16220             End If
16230           Next
16240           .Relations.Refresh

                ' ** Delete their copy of Ledger table.
16250           .TableDefs.Delete "ledger"

                ' ** Copy the new Ledger table to their TrustDta.mdb.
16260           DoCmd.TransferDatabase acExport, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "tblTemplate_Ledger", "ledger", False  ' ** Don't just copy structure, copy with data.

16270           .TableDefs.Refresh

                ' ** NEW CASCADE DELETES!!!
                ' ** Re-create the relationships to Account, m_REVCODE, Location, RecurringItems.
16280           Set Rel = .CreateRelation("accountledger", "account", "ledger", dbRelationUpdateCascade + dbRelationDeleteCascade)
16290           With Rel
16300             .Fields.Append .CreateField("accountno", dbLong)
16310             .Fields![accountno].ForeignName = "accountno"
16320           End With
16330 On Error Resume Next
16340           .Relations.Append Rel
16350           If ERR.Number <> 0 Then
                  ' ** For now, just let this relationship go.
                  ' ** It's probably a client's file on my own computer.
16360 On Error GoTo ERRH
16370           Else
16380 On Error GoTo ERRH
16390           End If
16400           Set Rel = .CreateRelation("m_REVCODEledger", "m_REVCODE", "ledger", dbRelationUpdateCascade)
16410           With Rel
16420             .Fields.Append .CreateField("revcode_ID", dbLong)
16430             .Fields![revcode_ID].ForeignName = "revcode_ID"
16440           End With
16450 On Error Resume Next
16460           .Relations.Append Rel
16470           If ERR.Number <> 0 Then
                  ' ** For now, just let this relationship go.
                  ' ** It's probably a client's file on my own computer.
16480 On Error GoTo ERRH
16490           Else
16500 On Error GoTo ERRH
16510           End If
16520           Set Rel = .CreateRelation("Locationledger", "Location", "ledger", dbRelationUpdateCascade)
16530           With Rel
16540             .Fields.Append .CreateField("Location_ID", dbLong)
16550             .Fields![Location_ID].ForeignName = "Location_ID"
16560           End With
16570 On Error Resume Next
16580           .Relations.Append Rel
16590           If ERR.Number <> 0 Then
                  ' ** For now, just let this relationship go.
                  ' ** It's probably a client's file on my own computer.
16600 On Error GoTo ERRH
16610           Else
16620 On Error GoTo ERRH
16630           End If
16640           Set Rel = .CreateRelation("RecurringItemsledger", "RecurringItems", "ledger", dbRelationDontEnforce)
16650           With Rel
16660             .Fields.Append .CreateField("RecurringItem", dbLong)
16670             .Fields![RecurringItem].ForeignName = "RecurringItem"
16680           End With
16690 On Error Resume Next
16700           .Relations.Append Rel
16710           If ERR.Number <> 0 Then
                  ' ** For now, just let this relationship go.
                  ' ** It's probably a client's file on my own computer.
16720 On Error GoTo ERRH
16730           Else
16740 On Error GoTo ERRH
16750           End If
16760           .Relations.Refresh

                ' ** Relink Ledger back to here.
16770           DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), _
                  acTable, "ledger", "ledger"

16780         End If

              ' ** Create the relationships twixt MasterAsset and ActiveAssets, MasterAsset and Asset.
16790         If lngErrs = 0& Then
16800           Set Rel = .CreateRelation("masterassetActiveAssets", "masterasset", "ActiveAssets", dbRelationUpdateCascade)
16810           With Rel
16820             .Fields.Append .CreateField("assetno", dbLong)
16830             .Fields![assetno].ForeignName = "assetno"
16840           End With
16850           .Relations.Append Rel
16860           Set Rel = .CreateRelation("masterassetasset", "masterasset", "asset", dbRelationUpdateCascade)
16870           With Rel
16880             .Fields.Append .CreateField("assetno", dbLong)
16890             .Fields![assetno].ForeignName = "assetno"
16900           End With
16910           .Relations.Append Rel
16920           .Relations.Refresh
16930         End If

16940       End If

            ' *********************************************************************
            ' ** Step 20.5. JournalType to Journal.
            ' *********************************************************************

            ' ** Link journaltype to journal.
16950       blnFound = False
16960       For lngX = 0& To (lngRels - 1&)
16970         If arr_varRel(R_TBL, lngX) = "journaltype" And arr_varRel(R_FTBL, lngX) = "journal" Then
16980           blnFound = True
16990           Exit For
17000         End If
17010       Next
17020       If blnFound = False Then
17030         Set Rel = .CreateRelation("journaltypejournal", "journaltype", "journal", dbRelationUpdateCascade)
17040         With Rel
17050           .Fields.Append .CreateField("journaltype", dbText)
17060           .Fields![journaltype].ForeignName = "journaltype"
17070         End With
17080         .Relations.Append Rel
17090         .Relations.Refresh
17100       End If

            ' *********************************************************************
            ' ** Step 20.6. JournalType to Ledger.
            ' *********************************************************************

            ' ** Link journaltype to ledger.
17110       blnFound = False
17120       For lngX = 0& To (lngRels - 1&)
17130         If arr_varRel(R_TBL, lngX) = "journaltype" And arr_varRel(R_FTBL, lngX) = "ledger" Then
17140           blnFound = True
17150           Exit For
17160         End If
17170       Next
17180       If blnFound = False Then
17190         Set Rel = .CreateRelation("journaltypeledger", "journaltype", "ledger", dbRelationUpdateCascade)
17200         With Rel
17210           .Fields.Append .CreateField("journaltype", dbText)
17220           .Fields![journaltype].ForeignName = "journaltype"
17230         End With
17240         .Relations.Append Rel
17250         .Relations.Refresh
17260       End If

            ' *********************************************************************
            ' ** Step 20.7. Account to AccountType.
            ' *********************************************************************

            ' ** Check Account against AccountType.
17270       Set qdf1 = CurrentDb.QueryDefs("qrySystemUpdate_08_Account")
17280       Set qdf2 = .CreateQueryDef("", qdf1.SQL)
17290       Set rst = qdf2.OpenRecordset
17300       With rst
17310         If .BOF = True And .EOF = True Then
                ' ** All's well.
17320         Else
                ' ** Give them a default AccountType.
17330           .MoveLast
17340           lngRecs = .RecordCount
17350           .MoveFirst
17360           For lngX = 1& To lngRecs
17370             .Edit
17380             ![accounttype] = "85"  ' ** Other.
17390             .Update
17400             If lngX < lngRecs Then .MoveNext
17410           Next
17420           lngErrs = lngErrs + 1&
17430           lngE = lngErrs - 1&
17440           ReDim Preserve arr_varErr(E_ELEMS, lngE)
17450           arr_varErr(E_OBJ, lngE) = "ActiveAssets"
17460           arr_varErr(E_DESC, lngE) = "There were Account records without a valid AccountType. Corrected."
17470           arr_varErr(E_MISC, lngE) = CStr(lngRecs)
17480           arr_varErr(E_SKIP, lngE) = CBool(False)
17490         End If
17500         .Close
17510       End With

            ' ** Link accounttype to account.
17520       blnFound = False
17530       For lngX = 0& To (lngRels - 1&)
17540         If arr_varRel(R_TBL, lngX) = "accounttype" And arr_varRel(R_FTBL, lngX) = "account" Then
17550           blnFound = True
17560           Exit For
17570         End If
17580       Next
17590       If blnFound = False Then
17600         Set Rel = .CreateRelation("accounttypeaccount", "accounttype", "account", dbRelationUpdateCascade)
17610         With Rel
17620           .Fields.Append .CreateField("accounttype", dbText)
17630           .Fields![accounttype].ForeignName = "accounttype"
17640         End With
17650         .Relations.Append Rel
17660         .Relations.Refresh
17670       End If

            ' *********************************************************************
            ' ** Step 20.8. AssetType to MasterAsset.
            ' *********************************************************************

            ' ** Check MasterAsset against AssetType.
17680       Set qdf1 = CurrentDb.QueryDefs("qrySystemUpdate_09_MasterAsset")
17690       Set qdf2 = .CreateQueryDef("", qdf1.SQL)
17700       Set rst = qdf2.OpenRecordset
17710       With rst
17720         If .BOF = True And .EOF = True Then
                ' ** All's well.
17730         Else
                ' ** Give them a default AssetType.
17740           .MoveLast
17750           lngRecs = .RecordCount
17760           .MoveFirst
17770           For lngX = 1& To lngRecs
17780             .Edit
17790             ![assettype] = "75"  ' ** Other.
17800             .Update
17810             If lngX < lngRecs Then .MoveNext
17820           Next
17830           lngErrs = lngErrs + 1&
17840           lngE = lngErrs - 1&
17850           ReDim Preserve arr_varErr(E_ELEMS, lngE)
17860           arr_varErr(E_OBJ, lngE) = "ActiveAssets"
17870           arr_varErr(E_DESC, lngE) = "There were MasterAsset records without a valid AssetType. Corrected."
17880           arr_varErr(E_MISC, lngE) = CStr(lngRecs)
17890           arr_varErr(E_SKIP, lngE) = CBool(False)
17900         End If
17910         .Close
17920       End With

            ' ** Link assettype to masterasset.
17930       blnFound = False
17940       For lngX = 0& To (lngRels - 1&)
17950         If arr_varRel(R_TBL, lngX) = "assettype" And arr_varRel(R_FTBL, lngX) = "masterasset" Then
17960           blnFound = True
17970           Exit For
17980         End If
17990       Next
18000       If blnFound = False Then
18010         Set Rel = .CreateRelation("assettypemasterasset", "assettype", "masterasset", dbRelationUpdateCascade)
18020         With Rel
18030           .Fields.Append .CreateField("assettype", dbText)
18040           .Fields![assettype].ForeignName = "assettype"
18050         End With
18060         .Relations.Append Rel
18070         .Relations.Refresh
18080       End If

            ' *********************************************************************
            ' ** Step 20.9. TaxCode to AssetType.
            ' *********************************************************************

            ' ** Link taxcode to assettype.
18090       blnFound = False
18100       For lngX = 0& To (lngRels - 1&)
18110         If arr_varRel(R_TBL, lngX) = "taxcode" And arr_varRel(R_FTBL, lngX) = "assettype" Then
18120           blnFound = True
18130           Exit For
18140         End If
18150       Next
18160       If blnFound = False Then
18170         Set Rel = .CreateRelation("taxcodeassettype", "taxcode", "assettype", dbRelationUpdateCascade)
18180         With Rel
18190           .Fields.Append .CreateField("taxcode", dbInteger)
18200           .Fields![taxcode].ForeignName = "taxcode"
18210         End With
18220         .Relations.Append Rel
18230         .Relations.Refresh
18240       End If

            ' ** Delete dead tables.
18250       lngRecs = .TableDefs.Count
18260       For lngX = (lngRecs - 1&) To 0 Step -1&
18270         Set tdf = .TableDefs(lngX)
18280         Select Case tdf.Name
              Case "assetsub", "Lock", "masterasset temp", "tblAveragePrice"
18290           .TableDefs.Delete tdf.Name
18300         Case Else
                ' ** OK.
18310         End Select
18320       Next
18330       .TableDefs.Refresh

            ' *********************************************************************
            ' ** Step 20.10. Account to Journal CascadeDelete
            ' *********************************************************************

18340       blnFound = False
18350       For Each Rel In .Relations
18360         With Rel
18370           If ((.Table = "account" And .ForeignTable = "journal") Or _
                    (.Table = "journal" And .ForeignTable = "account")) Then
                  ' ** "accountjournal", "account", "journal"
18380             If ((.Attributes And dbRelationUpdateCascade) > 0) And _
                      ((.Attributes And dbRelationDeleteCascade) > 0) Then
18390               blnFound = True
18400               Exit For
18410             End If
18420           End If
18430         End With
18440       Next
18450       If blnFound = False Then
              ' ** Delete Journal's relationship to Account.
18460         For Each Rel In .Relations
18470           With Rel
18480             If ((.Table = "account" And .ForeignTable = "journal") Or _
                      (.Table = "journal" And .ForeignTable = "account")) Then
18490               dbs.Relations.Delete .Name
18500               dbs.Relations.Refresh
18510             End If
18520           End With
18530         Next
              ' ** NEW CASCADE DELETES!!!
              ' ** Re-create the relationship.
18540         Set Rel = .CreateRelation("accountjournal", "account", "journal", dbRelationUpdateCascade + dbRelationDeleteCascade)
18550         With Rel
18560           .Fields.Append .CreateField("accountno", dbLong)
18570           .Fields![accountno].ForeignName = "accountno"
18580         End With
18590 On Error Resume Next
18600         .Relations.Append Rel  ' ** If they didn't have one to begin with, it might not go now.
18610 On Error GoTo ERRH
18620         .Relations.Refresh
18630       End If

            ' *********************************************************************
            ' ** Step 20.11. Account to Ledger CascadeDelete
            ' *********************************************************************

18640       blnFound = False
18650       For Each Rel In .Relations
18660         With Rel
18670           If ((.Table = "account" And .ForeignTable = "ledger") Or _
                    (.Table = "ledger" And .ForeignTable = "account")) Then
                  ' ** "accountledger", "account", "ledger"
18680             If ((.Attributes And dbRelationUpdateCascade) > 0) And _
                      ((.Attributes And dbRelationDeleteCascade) > 0) Then
18690               blnFound = True
18700               Exit For
18710             End If
18720           End If
18730         End With
18740       Next
18750       If blnFound = False Then
              ' ** Delete Ledger's relationship to Account.
18760         For Each Rel In .Relations
18770           With Rel
18780             If ((.Table = "account" And .ForeignTable = "ledger") Or _
                      (.Table = "ledger" And .ForeignTable = "account")) Then
18790               dbs.Relations.Delete .Name
18800               dbs.Relations.Refresh
18810             End If
18820           End With
18830         Next
              ' ** NEW CASCADE DELETES!!!
              ' ** Re-create the relationship.
18840         Set Rel = .CreateRelation("accountledger", "account", "ledger", dbRelationUpdateCascade + dbRelationDeleteCascade)
18850         With Rel
18860           .Fields.Append .CreateField("accountno", dbLong)
18870           .Fields![accountno].ForeignName = "accountno"
18880         End With
18890 On Error Resume Next
18900         .Relations.Append Rel  ' ** If they didn't have one to begin with, it might not go now.
18910 On Error GoTo ERRH
18920         .Relations.Refresh
18930       End If

            ' *********************************************************************
            ' ** Step 20.12. Account to ActiveAssets CascadeDelete
            ' *********************************************************************

18940       blnFound = False
18950       For Each Rel In .Relations
18960         With Rel
18970           If ((.Table = "account" And .ForeignTable = "ActiveAssets") Or _
                    (.Table = "ActiveAssets" And .ForeignTable = "account")) Then
                  ' ** "accountActiveAssets", "account", "ActiveAssets"
18980             If ((.Attributes And dbRelationUpdateCascade) > 0) And _
                      ((.Attributes And dbRelationDeleteCascade) > 0) Then
18990               blnFound = True
19000               Exit For
19010             End If
19020           End If
19030         End With
19040       Next
19050       If blnFound = False Then
              ' ** Delete ActiveAssets' relationship to Account.
19060         For Each Rel In .Relations
19070           With Rel
19080             If ((.Table = "account" And .ForeignTable = "ActiveAssets") Or _
                      (.Table = "ActiveAssets" And .ForeignTable = "account")) Then
19090               dbs.Relations.Delete .Name
19100               dbs.Relations.Refresh
19110             End If
19120           End With
19130         Next
              ' ** NEW CASCADE DELETES!!!
              ' ** Re-create the relationship.
19140         Set Rel = .CreateRelation("accountActiveAssets", "account", "ActiveAssets", dbRelationUpdateCascade + dbRelationDeleteCascade)
19150         With Rel
19160           .Fields.Append .CreateField("accountno", dbLong)
19170           .Fields![accountno].ForeignName = "accountno"
19180         End With
19190 On Error Resume Next
19200         .Relations.Append Rel  ' ** If they didn't have one to begin with, it might not go now.
19210 On Error GoTo ERRH
19220         .Relations.Refresh
19230       End If

            ' *********************************************************************
            ' ** Step 20.13. Account to Asset CascadeDelete
            ' *********************************************************************

19240       blnFound = False
19250       For Each Rel In .Relations
19260         With Rel
19270           If ((.Table = "account" And .ForeignTable = "asset") Or _
                    (.Table = "asset" And .ForeignTable = "account")) Then
                  ' ** "accountasset", "account", "asset"
19280             If ((.Attributes And dbRelationUpdateCascade) > 0) And _
                      ((.Attributes And dbRelationDeleteCascade) > 0) Then
19290               blnFound = True
19300               Exit For
19310             End If
19320           End If
19330         End With
19340       Next
19350       If blnFound = False Then
              ' ** Delete Asset's relationship to Account.
19360         For Each Rel In .Relations
19370           With Rel
19380             If ((.Table = "account" And .ForeignTable = "asset") Or _
                      (.Table = "asset" And .ForeignTable = "account")) Then
19390               dbs.Relations.Delete .Name
19400               dbs.Relations.Refresh
19410             End If
19420           End With
19430         Next
              ' ** NEW CASCADE DELETES!!!
              ' ** Re-create the relationship.
19440         Set Rel = .CreateRelation("accountasset", "account", "asset", dbRelationUpdateCascade + dbRelationDeleteCascade)
19450         With Rel
19460           .Fields.Append .CreateField("accountno", dbLong)
19470           .Fields![accountno].ForeignName = "accountno"
19480         End With
19490 On Error Resume Next
19500         .Relations.Append Rel  ' ** If they didn't have one to begin with, it might not go now.
19510 On Error GoTo ERRH
19520         .Relations.Refresh
19530       End If

            ' *********************************************************************
            ' ** Step 20.14. Account to Balance CascadeDelete
            ' *********************************************************************

19540       blnFound = False
19550       For Each Rel In .Relations
19560         With Rel
19570           If ((.Table = "account" And .ForeignTable = "Balance") Or _
                    (.Table = "Balance" And .ForeignTable = "account")) Then
                  ' ** "accountBalance", "account", "Balance"
19580             If ((.Attributes And dbRelationUpdateCascade) > 0) And _
                      ((.Attributes And dbRelationDeleteCascade) > 0) Then
19590               blnFound = True
19600               Exit For
19610             End If
19620           End If
19630         End With
19640       Next
19650       If blnFound = False Then
              ' ** Delete Balance's relationship to Account.
19660         For Each Rel In .Relations
19670           With Rel
19680             If ((.Table = "account" And .ForeignTable = "Balance") Or _
                      (.Table = "Balance" And .ForeignTable = "account")) Then
19690               dbs.Relations.Delete .Name
19700               dbs.Relations.Refresh
19710             End If
19720           End With
19730         Next
              ' ** NEW CASCADE DELETES!!!
              ' ** Re-create the relationship.
19740         Set Rel = .CreateRelation("accountBalance", "account", "Balance", dbRelationUpdateCascade + dbRelationDeleteCascade)
19750         With Rel
19760           .Fields.Append .CreateField("accountno", dbLong)
19770           .Fields![accountno].ForeignName = "accountno"
19780         End With
19790 On Error Resume Next
19800         .Relations.Append Rel  ' ** If they didn't have one to begin with, it might not go now.
19810 On Error GoTo ERRH
19820         .Relations.Refresh
19830       End If

            ' *********************************************************************
            ' ** Step 20.15. Account to PortfolioModel CascadeDelete
            ' *********************************************************************

19840       blnFound = False
19850       For Each Rel In .Relations
19860         With Rel
19870           If ((.Table = "account" And .ForeignTable = "PortfolioModel") Or _
                    (.Table = "PortfolioModel" And .ForeignTable = "account")) Then
                  ' ** "accountPortfolioModel", "account", "PortfolioModel"
19880             If ((.Attributes And dbRelationUpdateCascade) > 0) And _
                      ((.Attributes And dbRelationDeleteCascade) > 0) Then
19890               blnFound = True
19900               Exit For
19910             End If
19920           End If
19930         End With
19940       Next
19950       If blnFound = False Then
              ' ** Delete PortfolioModel's relationship to Account.
19960         For Each Rel In .Relations
19970           With Rel
19980             If ((.Table = "account" And .ForeignTable = "PortfolioModel") Or _
                      (.Table = "PortfolioModel" And .ForeignTable = "account")) Then
19990               dbs.Relations.Delete .Name
20000               dbs.Relations.Refresh
20010             End If
20020           End With
20030         Next
              ' ** NEW CASCADE DELETES!!!
              ' ** Re-create the relationship.
20040         Set Rel = .CreateRelation("accountPortfolioModel", "account", "PortfolioModel", dbRelationUpdateCascade + dbRelationDeleteCascade)
20050         With Rel
20060           .Fields.Append .CreateField("accountno", dbLong)
20070           .Fields![accountno].ForeignName = "accountnum"  ' ***************************!
20080         End With
20090 On Error Resume Next
20100         .Relations.Append Rel  ' ** If they didn't have one to begin with, it might not go now.
20110 On Error GoTo ERRH
20120         .Relations.Refresh
20130       End If

            ' *********************************************************************
            ' ** Step 20.16. m_REVCODE_TYPE to m_REVCODE
            ' *********************************************************************
20140       blnFound = False
20150       For Each Rel In .Relations
20160         With Rel
20170           If ((.Table = "m_REVCODE_TYPE" And .ForeignTable = "m_REVCODE") Or _
                    (.Table = "m_REVCODE" And .ForeignTable = "m_REVCODE_TYPE")) Then
                  ' ** "m_REVCODE_TYPEm_REVCODE", "m_REVCODE_TYPE", "m_REVCODE"
20180             blnFound = True
20190             Exit For
20200           End If
20210         End With
20220       Next
20230       If blnFound = False Then
              ' ** Create the relationship.
20240         Set Rel = .CreateRelation("m_REVCODE_TYPEm_REVCODE", "m_REVCODE_TYPE", "m_REVCODE", dbRelationUpdateCascade)
20250         With Rel
20260           .Fields.Append .CreateField("revcode_TYPE", dbLong)
20270           .Fields![revcode_TYPE].ForeignName = "revcode_TYPE"
20280         End With
20290 On Error Resume Next
20300         .Relations.Append Rel
20310 On Error GoTo ERRH
20320         .Relations.Refresh
20330       End If

            ' *********************************************************************
            ' ** Step 20.17. InvestmentObjective to Account
            ' *********************************************************************
20340       blnFound = False
20350       For Each Rel In .Relations
20360         With Rel
20370           If ((.Table = "InvestmentObjective" And .ForeignTable = "account") Or _
                    (.Table = "account" And .ForeignTable = "InvestmentObjective")) Then
                  ' ** "InvestmentObjectiveaccount", "InvestmentObjective", "account"
20380             blnFound = True
20390             Exit For
20400           End If
20410         End With
20420       Next
20430       If blnFound = False Then
              ' ** Create the relationship.
20440         Set Rel = .CreateRelation("InvestmentObjectiveaccount", "InvestmentObjective", "account", dbRelationUpdateCascade)
20450         With Rel
20460           .Fields.Append .CreateField("invobj_name", dbText)
20470           .Fields![invobj_name].ForeignName = "investmentobj"
20480         End With
20490 On Error Resume Next
20500         .Relations.Append Rel
20510 On Error GoTo ERRH
20520         .Relations.Refresh
20530       End If

            ' *********************************************************************
            ' ** Step 20.18. Schedule to Account
            ' *********************************************************************
20540       blnFound = False
20550       For Each Rel In .Relations
20560         With Rel
20570           If ((.Table = "Schedule" And .ForeignTable = "account") Or _
                    (.Table = "account" And .ForeignTable = "Schedule")) Then
                  ' ** "Scheduleaccount", "Schedule", "account"
20580             blnFound = True
20590             Exit For
20600           End If
20610         End With
20620       Next
20630       If blnFound = False Then
              ' ** Create the relationship.
20640         Set Rel = .CreateRelation("Scheduleaccount", "Schedule", "account", dbRelationUpdateCascade)
20650         With Rel
20660           .Fields.Append .CreateField("Schedule_ID", dbText)
20670           .Fields![Schedule_ID].ForeignName = "Schedule_ID"
20680         End With
20690 On Error Resume Next
20700         .Relations.Append Rel
20710 On Error GoTo ERRH
20720         .Relations.Refresh
20730       End If

            ' *********************************************************************
            ' ** Step 20.19. Schedule to ScheduleDetail
            ' *********************************************************************
20740       blnFound = False
20750       For Each Rel In .Relations
20760         With Rel
20770           If ((.Table = "Schedule" And .ForeignTable = "ScheduleDetail") Or _
                    (.Table = "ScheduleDetail" And .ForeignTable = "Schedule")) Then
                  ' ** "ScheduleScheduleDetail", "Schedule", "ScheduleDetail"
20780             blnFound = True
20790             Exit For
20800           End If
20810         End With
20820       Next
20830       If blnFound = False Then
              ' ** Create the relationship.
20840         Set Rel = .CreateRelation("ScheduleScheduleDetail", "Schedule", "ScheduleDetail", dbRelationUpdateCascade + dbRelationDeleteCascade)
20850         With Rel
20860           .Fields.Append .CreateField("Schedule_ID", dbText)
20870           .Fields![Schedule_ID].ForeignName = "Schedule_ID"
20880         End With
20890 On Error Resume Next
20900         .Relations.Append Rel
20910 On Error GoTo ERRH
20920         .Relations.Refresh
20930       End If

            ' *********************************************************************
            ' ** Step 20.20. HiddenType to LedgerHidden
            ' *********************************************************************
20940       blnFound = False
20950       For Each Rel In .Relations
20960         With Rel
20970           If ((.Table = "HiddenType" And .ForeignTable = "LedgerHidden") Or _
                    (.Table = "LedgerHidden" And .ForeignTable = "HiddenType")) Then
                  ' ** "HiddenTypeLedgerHidden", "HiddenType", "LedgerHidden"
20980             blnFound = True
20990             Exit For
21000           End If
21010         End With
21020       Next
21030       If blnFound = False Then
21040         Set Rel = .CreateRelation("HiddenTypeLedgerHidden", "HiddenType", "LedgerHidden", dbRelationUpdateCascade)
21050         With Rel
21060           .Fields.Append .CreateField("hidtype", dbText)
21070           .Fields![hidtype].ForeignName = "hidtype"
21080         End With
21090 On Error Resume Next
21100         .Relations.Append Rel
21110 On Error GoTo ERRH
21120         .Relations.Refresh
21130       End If

            ' *********************************************************************
            ' ** Step 20.21. tblPreference_Control to tblPreference_User  LOCAL!
            ' *********************************************************************
            'blnFound = False
            'For Each rel In CurrentDb.Relations
            '  With rel
            '    If ((.Table = "tblPreference_Control" And .ForeignTable = "tblPreference_User") Or _
            '        (.Table = "tblPreference_User" And .ForeignTable = "tblPreference_Control")) Then
            '      ' ** "tblPreference_ControltblPreference_User", "tblPreference_Control", "tblPreference_User"
21140       blnFound = True
            '      Exit For
            '    End If
            '  End With
            'Next
21150       If blnFound = False Then
21160         Set Rel = CurrentDb.CreateRelation("tblPreference_ControltblPreference_User", _
                "tblPreference_Control", "tblPreference_User", dbRelationDontEnforce)
21170         With Rel
21180           .Fields.Append .CreateField("frm_name", dbText)
21190           .Fields![frm_name].ForeignName = "frm_name"
21200           .Fields.Append .CreateField("ctl_name", dbText)
21210           .Fields![ctl_name].ForeignName = "ctl_name"
21220         End With
21230 On Error Resume Next
21240         CurrentDb.Relations.Append Rel
21250         If ERR.Number <> 0 Then
21260           Debug.Print "'Check_BE_Rels(): tblPreference_ControltblPreference_User FAILED  Error: " & CStr(ERR.Number)
21270         End If
21280 On Error GoTo ERRH
21290         CurrentDb.Relations.Refresh
21300       End If
            ' ** DO NOT link it in the backend, because 'Admin' is never in the
            ' ** Users table, and that's the user most often used by customers!
21310       blnFound = False: strTmp01 = vbNullString
            'For Each rel In .Relations
            '  With rel
            '    If ((.Table = "Users" And .ForeignTable = "tblPreference_User") Or _
            '        (.Table = "tblPreference_User" And .ForeignTable = "Users")) Then
            '      ' ** "UserstblPreference_User", "Users", "tblPreference_User"
            '      strTmp01 = .Name
            '      blnFound = True
            '      Exit For
            '    End If
            '  End With
            'Next
21320       If blnFound = True Then
21330         .Relations.Delete strTmp01
21340         .Relations.Refresh
21350         strTmp01 = vbNullString
21360       End If

            ' *********************************************************************
            ' ** Step 20.22. Create/Update AppDate.
            ' *********************************************************************
21370       blnFound = False
21380       lngRecs = .Containers![Databases].Documents![UserDefined].Properties.Count
21390       For lngX = 0& To (lngRecs - 1&)
21400         Set prp = .Containers![Databases].Documents![UserDefined].Properties(lngX)
21410         With prp
21420           If .Name = "AppDate" Then
21430             blnFound = True
21440             Exit For
21450           End If
21460         End With
21470       Next
21480       If blnFound = False Then
21490         Set doc = .Containers![Databases].Documents![UserDefined]
21500         With doc
21510           Set prp = .CreateProperty("AppDate", dbDate, Now(), True)  ' ** True prevents easy Edit or Delete.
21520           .Properties.Append prp
21530         End With
21540       Else
21550         Set doc = .Containers![Databases].Documents![UserDefined]
21560         doc.Properties![AppDate].Value = Now()
21570       End If

            ' *********************************************************************
            ' ** Step 20.23. Create/Update AppVersion.
            ' *********************************************************************
21580       blnFound = False
21590       lngRecs = .Containers![Databases].Documents![UserDefined].Properties.Count
21600       For lngX = 0& To (lngRecs - 1&)
21610         Set prp = .Containers![Databases].Documents![UserDefined].Properties(lngX)
21620         With prp
21630           If .Name = "AppVersion" Then
21640             blnFound = True
21650             Exit For
21660           End If
21670         End With
21680       Next
21690       If blnFound = False Then
21700         Set doc = dbs.Containers![Databases].Documents![UserDefined]
21710         With doc
                'strOldArchVer
21720           If Right(strOldDtaVer, 1) = "." Then strOldDtaVer = Left(strOldDtaVer, (Len(strOldDtaVer) - 1))
21730           Set prp = .CreateProperty("AppVersion", dbText, strOldDtaVer, True)  ' ** True prevents easy Edit or Delete.
21740           .Properties.Append prp
21750         End With
21760       Else
              'Set doc = dbs.Containers![Databases].Documents![UserDefined]
              'doc.Properties![AppVersion].Value = "2.1.61"
21770       End If

            ' *********************************************************************
            'ADD OTHER PROPS.
            ' *********************************************************************

21780       .Close
21790     End With
21800     .Close
21810   End With

EXITP:
21820   DoCmd.Hourglass False
21830   Set prp = Nothing
21840   Set doc = Nothing
21850   Set fld = Nothing
21860   Set tdf = Nothing
21870   Set Rel = Nothing
21880   Set rst = Nothing
21890   Set qdf1 = Nothing
21900   Set qdf2 = Nothing
21910   Set dbs = Nothing
21920   Set wrk = Nothing
21930   Check_BE_Rels = blnRetVal
21940   Exit Function

ERRH:
21950   blnRetVal = False
21960   DoCmd.Hourglass False
21970   Select Case ERR.Number
        Case 3201  ' ** You cannot add or change a record because a related record is required in table '|'.
          ' ** This may be just me linking to an old client's database. Let it go...
21980   Case Else
21990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22000   End Select
22010   Resume EXITP

End Function

Private Function ChkUpdate(blnRead As Boolean, lngTmp02 As Long) As Boolean

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkUpdate"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, rst As DAO.Recordset
        Dim strTmp01 As String, lngTmp03 As Long
        Dim blnRetVal As Boolean

22110   blnRetVal = False  ' ** Default to not running the full update.

22120   If blnRead = True Then
22130     strOldDtaVer = vbNullString: strOldArchVer = vbNullString
22140   End If

22150   Select Case blnRead
        Case True
          ' ** Read the codes.
22160     Set dbs = CurrentDb
22170     With dbs
22180       Set rst = .OpenRecordset("m_VP", dbOpenDynaset, dbConsistent)
22190       With rst
22200         .MoveFirst
22210         If IsNull(![vp_DE2]) = True Then
                ' ** If m_VP![vp_DE2] is empty, run everything.
22220           blnRetVal = True
22230         Else
22240           lngTmp03 = CLng(Val(![vp_DE2]))
22250           If lngTmp03 <> lngTmp02 Then
22260             If lngUpdates = 1& Then
                    ' ** Something's screwy! If I've only made one, and this isn't it, run everything!
22270               blnRetVal = True
22280             Else
                    ' ** If m_VP![vp_DE2] has earlier change date, run everything.
22290               blnRetVal = True
22300             End If
22310           Else
                  ' ** If m_VP![vp_DE2] has latest change date, check m_VD![vd_DE2].
22320           End If
22330         End If
22340         .Close
22350       End With
22360       If blnRetVal = False Then
22370         Set rst = .OpenRecordset("m_VD", dbOpenDynaset, dbConsistent)
22380         rst.MoveFirst
22390         If IsNull(rst![vd_MAIN]) = False Then
22400           strOldDtaVer = CStr(rst![vd_MAIN]) & "."
22410           If IsNull(rst![vd_MINOR]) = False Then
22420             strTmp01 = CStr(rst![vd_MINOR])
22430             If Len(strTmp01) = 2 Then
22440               If Right(strTmp01, 1) = "0" Then
22450                 strTmp01 = Left(strTmp01, 1) & "."
22460                 If IsNull(rst![vd_REVISION]) = False Then
22470                   strTmp01 = strTmp01 & CStr(rst![vb_REVISION])
22480                 Else
22490                   strTmp01 = strTmp01 & "0"
22500                 End If
22510               Else
22520                 strTmp01 = Left(strTmp01, 1) & "." & Mid(strTmp01, 2)
22530                 If IsNull(rst![vd_REVISION]) = False Then
22540                   strTmp01 = strTmp01 & CStr(rst![vb_REVISION])
22550                 End If
22560               End If
22570               strOldDtaVer = strOldDtaVer & strTmp01
22580             Else
22590               strOldDtaVer = strOldDtaVer & strTmp01 & "."
22600               If IsNull(rst![vd_REVISION]) = False Then
22610                 strOldDtaVer = strOldDtaVer & CStr(rst![vd_REVISION]) & "."
22620               Else
22630                 strOldDtaVer = strOldDtaVer & "0."
22640               End If
22650             End If
22660           Else
22670             strOldDtaVer = strOldDtaVer & "0."
22680             If IsNull(rst![vd_REVISION]) = False Then
22690               strOldDtaVer = strOldDtaVer & CStr(rst![vd_REVISION]) & "."
22700             Else
22710               strOldDtaVer = strOldDtaVer & "0."
22720             End If
22730           End If
22740           strTmp01 = vbNullString
22750         Else
22760           strOldDtaVer = "2.x.x"
22770         End If
22780         If IsNull(rst![vd_DE2]) = True Then
                ' ** If m_VD![vd_DE2] is empty, run everyting.
22790           blnRetVal = True
22800           rst.Close
22810         Else
22820           strTmp01 = rst![vd_DE2]
22830           lngTmp03 = CLng(Val(Left(strTmp01, 3)))
22840           If lngTmp03 = lngUpdates Then
22850             If InStr(strTmp01, "A") > 0 Then
22860               rst.Close
22870               For Each tdf In .TableDefs
22880                 If tdf.Name = "m_VA" Then
22890                   Set rst = .OpenRecordset("m_VA", dbOpenDynaset, dbConsistent)
22900                   With rst
22910                     .MoveFirst
22920                     If IsNull(rst![va_MAIN]) = False Then
22930                       strOldArchVer = CStr(rst![va_MAIN]) & "."
22940                       If IsNull(rst![va_MINOR]) = False Then
22950                         strTmp01 = CStr(rst![va_MINOR])
22960                         If Len(strTmp01) = 2 Then
22970                           If Right(strTmp01, 1) = "0" Then
22980                             strTmp01 = Left(strTmp01, 1) & "."
22990                             If IsNull(rst![va_REVISION]) = False Then
23000                               strTmp01 = strTmp01 & CStr(rst![vb_REVISION])
23010                             Else
23020                               strTmp01 = strTmp01 & "0"
23030                             End If
23040                           Else
23050                             strTmp01 = Left(strTmp01, 1) & "." & Mid(strTmp01, 2)
23060                             If IsNull(rst![va_REVISION]) = False Then
23070                               strTmp01 = strTmp01 & CStr(rst![vb_REVISION])
23080                             End If
23090                           End If
23100                           strOldArchVer = strOldArchVer & strTmp01
23110                         Else
23120                           strOldArchVer = strOldArchVer & strTmp01 & "."
23130                           If IsNull(rst![va_REVISION]) = False Then
23140                             strOldArchVer = strOldArchVer & CStr(rst![va_REVISION]) & "."
23150                           Else
23160                             strOldArchVer = strOldArchVer & "0."
23170                           End If
23180                         End If
23190                       Else
23200                         strOldArchVer = strOldArchVer & "0."
23210                         If IsNull(rst![va_REVISION]) = False Then
23220                           strOldArchVer = strOldArchVer & CStr(rst![va_REVISION]) & "."
23230                         Else
23240                           strOldArchVer = strOldArchVer & "0."
23250                         End If
23260                       End If
23270                       strTmp01 = vbNullString
23280                     Else
23290                       strOldArchVer = "2.x.x"
23300                     End If
23310                     If IsNull(![va_DE2]) = True Then
                            ' ** If m_VD![vd_DE2] has latest change date, and there's
                            ' ** an 'A' in it, but m_VA![va_DE2] is empty, run everything.
23320                       blnRetVal = True
23330                     Else
23340                       If CLng(Val(Left(![va_DE2], 3))) <> lngUpdates Then
                              ' ** If m_VD![vd_DE2] has latest change date, and there's an 'A' in it,
                              ' ** but m_VA![va_DE2] is different, run everything.
23350                         blnRetVal = True
23360                       Else
                              ' ** Everything seems to be fine, so only check things marked RUN EVERY TIME!
23370                       End If
23380                     End If
23390                     .Close
23400                   End With
23410                   Exit For
23420                 End If
23430               Next
                    ' ** If it never finds m_VA, just leave blnRetVal = False.
23440             Else
                    ' ** If m_VD![vd_DE2] has latest change date, only check things marked RUN EVERY TIME!
23450               rst.Close
23460             End If
23470           Else
                  ' ** If m_VD![vd_DE2] has an earlier change date, run everything.
23480             blnRetVal = True
23490             rst.Close
23500           End If
23510         End If
23520       End If
23530       .Close
23540     End With
23550   Case False
          ' ** Save the codes.
23560     Set dbs = CurrentDb
23570     With dbs
23580       Set rst = .OpenRecordset("m_VP", dbOpenDynaset, dbConsistent)
23590       With rst
23600         .MoveFirst
23610         .Edit
23620         ![vp_DE2] = CStr(lngTmp02)  ' ** Just put in the change date.
23630         .Update
23640         .Close
23650       End With
23660       Set rst = .OpenRecordset("m_VD", dbOpenDynaset, dbConsistent)
23670       With rst
23680         .MoveFirst
23690         .Edit
23700         ![vd_DE2] = strUpdates  ' ** This will come over via the Private module-level variable.
23710         .Update
23720         .Close
23730       End With
23740       If InStr(strUpdates, "A") > 0 Then
23750         For Each tdf In .TableDefs
23760           If tdf.Name = "m_VA" Then
23770             Set rst = .OpenRecordset("m_VA", dbOpenDynaset, dbConsistent)
23780             With rst
23790               .MoveFirst
23800               .Edit
23810               ![va_DE2] = strUpdates
23820               .Update
23830               .Close
23840             End With
23850             Exit For
23860           End If
23870         Next
23880       End If
23890       .Close
23900     End With
23910     blnRetVal = True  ' ** Here, it means the save was successful.
23920   End Select

EXITP:
23930   ChkUpdate = blnRetVal
23940   Exit Function

ERRH:
23950   blnRetVal = True  ' ** If it encounters an error here, let it try to run the full update.
23960   Select Case ERR.Number
        Case Else
23970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23980   End Select
23990   Resume EXITP

End Function

Public Function TmpTblList() As Variant

24000 On Error GoTo ERRH

        Const THIS_PROC As String = "TmpTblList"

        Dim lngTmpTbls As Long, arr_varTmpTbl() As Variant
        Dim lngE As Long

24010   lngTmpTbls = 0&
24020   ReDim arr_varTmpTbl(0)

24030   lngTmpTbls = lngTmpTbls + 1&
24040   lngE = lngTmpTbls - 1&
24050   ReDim Preserve arr_varTmpTbl(lngE)
24060   arr_varTmpTbl(lngE) = "tmp_ActiveAssets"

24070   lngTmpTbls = lngTmpTbls + 1&
24080   lngE = lngTmpTbls - 1&
24090   ReDim Preserve arr_varTmpTbl(lngE)
24100   arr_varTmpTbl(lngE) = "tmp_Journal"

24110   lngTmpTbls = lngTmpTbls + 1&
24120   lngE = lngTmpTbls - 1&
24130   ReDim Preserve arr_varTmpTbl(lngE)
24140   arr_varTmpTbl(lngE) = "tmp_Ledger"

24150   lngTmpTbls = lngTmpTbls + 1&
24160   lngE = lngTmpTbls - 1&
24170   ReDim Preserve arr_varTmpTbl(lngE)
24180   arr_varTmpTbl(lngE) = "tmp_m_REVCODE"

24190   lngTmpTbls = lngTmpTbls + 1&
24200   lngE = lngTmpTbls - 1&
24210   ReDim Preserve arr_varTmpTbl(lngE)
24220   arr_varTmpTbl(lngE) = "tmp_RecurringItems"

EXITP:
24230   TmpTblList = arr_varTmpTbl
24240   Exit Function

ERRH:
24250   Select Case ERR.Number
        Case Else
24260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24270   End Select
24280   Resume EXITP

End Function

Public Function Del_DB1s() As Boolean

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "Del_DB1s"

        Dim strFile1 As String, strFile2 As String
        Dim strPath As String
        Dim lngDels As Long, arr_varDel() As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

24310   blnRetVal = True

24320   lngDels = 0&
24330   ReDim arr_varDel(0)

24340   For lngX = 1& To 2&
24350     Select Case lngX
          Case 1&
24360       strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
24370     Case 2&
24380       If Left(CurrentBackendPath, Len(strPath)) = strPath Then  ' ** Module Function: modFileUtilities.
              ' ** Only check the backend directory if it's local.
24390         strPath = CurrentBackendPath  ' ** Module Function: modFileUtilities.
24400       Else
24410         strPath = vbNullString
24420       End If
24430     End Select
24440     If strPath <> vbNullString Then
24450       strFile1 = Dir(strPath & LNK_SEP & "db*.mdb")
24460       Do While strFile1 <> vbNullString
24470         strFile2 = Left(strFile1, (Len(strFile1) - 4))
24480         If Len(strFile2) = 3 Or Len(strFile2) = 4 Then
24490           If IsNumeric(Mid(strFile2, 3)) = True Then
                  ' ** Make sure it's one of those generic ones, not one coincidently prefixed with 'db'.
24500             lngDels = lngDels + 1&
24510             ReDim Preserve arr_varDel(lngDels - 1&)
24520             arr_varDel(lngDels - 1&) = strPath & LNK_SEP & strFile1
24530           End If
24540         End If
24550         strFile1 = Dir()
24560       Loop
24570     End If
24580   Next

24590   If lngDels > 0& Then
24600     For lngX = 0& To (lngDels - 1&)
24610 On Error Resume Next
24620       Kill arr_varDel(lngX)
24630 On Error GoTo ERRH
24640     Next
24650   End If

EXITP:
24660   Del_DB1s = blnRetVal
24670   Exit Function

ERRH:
24680   blnRetVal = False
24690   Select Case ERR.Number
        Case Else
24700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24710   End Select
24720   Resume EXITP

End Function
