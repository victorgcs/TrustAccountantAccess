Attribute VB_Name = "modVersionConvertFuncs2"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modVersionConvertFuncs2"

'VGC 10/27/2017: CHANGES!

' ** Array: arr_varOldFile().
'Private Const F_ELEMS As Integer = 11  ' ** Array's first-element UBound().
Private Const F_FNAM    As Integer = 0
Private Const F_PTHFIL  As Integer = 1
'Private Const F_DATA    As Integer = 2
'Private Const F_CONV    As Integer = 3
'Private Const F_TA_VER  As Integer = 4
Private Const F_ACC_VER As Integer = 5
Private Const F_TBLS    As Integer = 6
Private Const F_T_ARR   As Integer = 7
Private Const F_M_VER   As Integer = 8
Private Const F_APPVER  As Integer = 9
Private Const F_APPDATE As Integer = 10
Private Const F_NOTE    As Integer = 11

' ** Array: arr_varOldTbl().
Private Const T_ELEMS As Integer = 3  ' ** Array's first-element UBound().
Private Const T_TNAM  As Integer = 0
'Private Const T_TNAMN As Integer = 1
Private Const T_FLDS  As Integer = 2
Private Const T_F_ARR As Integer = 3

' ** Array: arr_varAcct().
Private Const A_ELEMS As Integer = 10  ' ** Array's first-element UBound().
Private Const A_NUM     As Integer = 0
Private Const A_NUM_N   As Integer = 1
Private Const A_NAM     As Integer = 2
Private Const A_TYP     As Integer = 3
Private Const A_ADMIN   As Integer = 4
Private Const A_ADMIN_N As Integer = 5
Private Const A_SCHED   As Integer = 6
Private Const A_SCHED_N As Integer = 7
Private Const A_DROPPED As Integer = 8
Private Const A_ACCT99  As Integer = 9
Private Const A_DASTNO  As Integer = 10

Private strTmp_Name As String, strTmp_Address1 As String, strTmp_Address2 As String, strTmp_City As String
Private strTmp_State As String, strTmp_Zip As String, strTmp_Country As String, strTmp_PostalCode As String, strTmp_Phone As String
Private blnTmp_IncomeTaxCoding As Boolean, blnTmp_RevenueExpenseTracking As Boolean, blnTmp_AccountNoWithType As Boolean
Private blnTmp_SeparateCheckingAccounts As Boolean, blnTmp_TabCopyAccount As Boolean, blnTmp_LinkRevTaxCodes As Boolean
Private blnTmp_SpecialCapGainLoss As Boolean, intTmp_SpecialCapGainLossOpt As Integer

Private strPathFile_Data As String, strPathFile_Archive As String, strOldVersion As String, strReleaseDate As String
Private lngErrNum As Long, lngErrLine As Long, strErrDesc As String
' **

Public Function ConversionCheck() As Integer
' ** This is the gatekeeper, passing it on to Version_Upgrade_01(), and receiving its response.
' ** Version_Upgrade_01() is the main distributor of the Version_Upgrade_{nn}() functions.
' **
' ** Called by:
' **   frmMenu_Title:
' **     Form_Open()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "ConversionCheck"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intRetVal As Integer

110     intRetVal = 0

120     If gblnDev_NoErrHandle = True Then
130   On Error GoTo 0
140     End If

150     If Len(TA_SEC) > Len(TA_SEC2) Then
          ' ** If this is a Demo, don't even check.
          ' ** See note about these in zz_mod_MDEPrepFuncs and modSecurityFunctions.
160       If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
170         gintConvertResponse = Version_Upgrade_01  ' ** Module Function: modVersionConvertFuncs1.
180       Else
190         gintConvertResponse = 1
200       End If
210     Else
          ' ** Check if it's a conversion.
220       gintConvertResponse = Version_Upgrade_01  ' ** Module Function: modVersionConvertFuncs1.
          ' ** Return values:
          ' **    1  Unnecessary
          ' **    0  OK
          ' **   -1  Can't Connect
          ' **   -2  Can't Open {TrustDta.mdb}
          ' **   -3  Canceled Status
          ' **   -4  Acount Empty
          ' **   -5  Canceled CoInfo
          ' **   -6  Index/Key Error
          ' **   -7  Can't Open {TrstArch.mdb}
          ' **   -8  Tables Not Empty
          ' **   -9  Error
230     End If

        ' ** Negative:
        ' **   All negative responses have already been dealt with and displayed.
        ' **   Should it continue with below's update?
        ' **   I would say NO, because the data may be in an unstable place!
        'If gintConvertResponse < 0 Then blnRunUpdates = False
        ' ** Positive:
        ' **   No conversion present or necessary. Continue normally.
        ' ** Zero:
        ' **   A 0 means a conversion took place, and technically won't require
        ' **   the update below. However, I think there are additional checks below
        ' **   that aren't done during convert. Also, newer fixes will show up below
        ' **   that I probably won't go back and incorporate into the conversion.
        ' **   Anyway, a convert should pass through it quickly.

        ' ** At the end of a conversion, Version_Upgrade_01() passes "End"
        ' ** to Version_Status(), changing cmdCancel to "Continue".
        ' ** It then returns here, and shows the summary with Version_Status 5, modVersionConvertFuncs1.

240     intRetVal = gintConvertResponse

250     If intRetVal <= 0 Then
260       If intRetVal = 0 Then
270         Set dbs = CurrentDb
280         With dbs
              ' ** Update tblPreference_User, for 'chkConversionCheck' = True, by specified [usr].
290           Set qdf = .QueryDefs("qryPreferences_06_03")  '##dbs_id
300           With qdf.Parameters
310             ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
320           End With
330           qdf.Execute dbFailOnError
340           .Close
350         End With
360       End If
          ' ** Pass the results on to the status window, where we sit when this is finished.
370       If IsLoaded(FRM_CNV_STATUS, acForm) = True Then  ' ** Module Function: modFileUtilities.
380         Forms("frmVersion_Main").ConversionCheck_Response gintConvertResponse  ' ** Form Procedure: frmVersion_Main.
390         Version_Status 5  ' ** Module Function: modVersionConvertFuncs1.
400         DoEvents
410         DoCmd.SelectObject acForm, FRM_CNV_STATUS, False
420   On Error Resume Next
430         Forms(FRM_CNV_STATUS).Status3.SetFocus
440   On Error GoTo ERRH
450       End If
460     End If

EXITP:
470     Set qdf = Nothing
480     Set dbs = Nothing
490     ConversionCheck = intRetVal
500     Exit Function

ERRH:
510     intRetVal = -9
520     Select Case ERR.Number
        Case 2489  ' ** The object 'frmVersion_Main' isn't open.
          ' ** They may have, like I did, hit Continue before this process could finish, so ignore it.
530       intRetVal = 0
540     Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
560     End Select
570     Resume EXITP

End Function

Public Function Version_Upgrade_05(blnContinue As Boolean, blnConvert_TrustDta As Boolean, lngTrustDtaDbsID As Long, lngLedgerEmptyDels As Long, lngDupeNum As Long, lngAccts As Long, arr_varTmp05 As Variant, lngRevCodes As Long, arr_varRevCode As Variant, lngTaxDefCodes As Long, arr_varTaxDefCode As Variant, lngDupeUnks As Long, arr_varTmp04 As Variant, lngTmp15 As Long, arr_varTmp01 As Variant, lngOldFiles As Long, arr_varOldFile As Variant, lngOldTbls As Long, arr_varOldTbl As Variant, lngStats As Long, arr_varTmp03 As Variant, dblPB_ThisStep As Double, strKeyTbl As String, wrkLoc As DAO.Workspace, wrkLnk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database) As Integer
' ** This continues the conversion process with the main data tables.
' ** Tables converted here:
' **   ledger
' **   journal
' **   LedgerHidden
' **
' ** Return values:
' **    0  OK
' **   -6  Index/Key
' **   -7  Can't Open
' **   -9  Error
' **
' ** Called by:
' **   modVersionConvertFuncs1:
' **     Version_Upgrade_01()

        ' ** Version_Upgrade_05(
        ' **   blnContinue As Boolean, blnConvert_TrustDta As Boolean, lngTrustDtaDbsID As Long,
        ' **   lngLedgerEmptyDels As Long, lngDupeNum As Long, lngAccts As Long, arr_varAcct As Variant,
        ' **   lngRevCodes As Long, arr_varRevCode As Variant, lngTaxDefCodes As Long, arr_varTaxDefCode As Variant,
        ' **   lngDupeUnks As Long, arr_varTmp04 As Variant, lngTmp15 As Long, arr_varTmp01 As Variant,
        ' **   lngOldFiles As Long, arr_varOldFile As Variant, lngOldTbls As Long, arr_varOldTbl As Variant,
        ' **   lngStats As Long, arr_varTmp03 As Variant, dblPB_ThisStep As Double, strKeyTbl As String,
        ' **   wrkLoc As DAO.Workspace, wrkLnk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database
        ' ** ) As Integer

        ' ** arr_varTmp01() = MasterAsset with no Ledger activity whatsoever  {defined in Version_Upgrade_04()}
        ' ** arr_varTmp02() = arr_varOldFile()  {defined and only used here}
        ' ** arr_varTmp03() = arr_varStat()     {altered and passed back}
        ' ** arr_varTmp04() = arr_varDupeUnk()  {altered and passed back}
        ' ** arr_varTmp05() = arr_varAcct()     {altered and passed back}

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Upgrade_05"

        Dim qdf As DAO.QueryDef, rstLoc1 As DAO.Recordset, rstLoc2 As DAO.Recordset, rstLoc3 As DAO.Recordset, rstLnk As DAO.Recordset
        Dim fld As DAO.Field
        Dim lngJTypes As Long, arr_varJType As Variant
        Dim arr_varAcct() As Variant, arr_varStat() As Variant, arr_varDupeUnk() As Variant
        Dim strCurrTblName As String, lngCurrTblID As Long, strCurrKeyFldName As String, lngCurrKeyFldID As Long
        Dim lngRecs As Long, lngFlds As Long
        Dim blnFound As Boolean, blnFound2 As Boolean
        Dim varTmp00 As Variant, arr_varTmp02 As Variant
        Dim strTmp05 As String, strTmp06 As String, strTmp07 As String, strTmp08 As String
        Dim strTmp09 As String, strTmp10 As String, strTmp11 As String
        Dim lngTmp13 As Long, lngTmp14 As Long, lngTmp16 As Long, lngTmp17 As Long, lngTmp18 As Long, blnTmp27 As Boolean
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim intRetVal As Integer

        ' ** Array: arr_varJType().
        Const JT_TYP  As Integer = 0
        'Const JT_DSC  As Integer = 1
        'Const JT_SORT As Integer = 2

        ' ** Array: arr_varDupeUnk().
        Const DU_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const DU_TYP As Integer = 0
        Const DU_TBL As Integer = 1

        ' ** Array: arr_varTaxDefCode().
        Const TD_ID_OLD As Integer = 0
        Const TD_ID_NEW As Integer = 1
        'Const TD_DSC    As Integer = 2
        'Const TD_TYP    As Integer = 3

        ' ** Array: arr_varRevCode().
        'Const R_ELEMS As Integer = 10  ' ** Array's first-element UBound().
        'Const R_REC As Integer = 0
        Const R_ID  As Integer = 1
        'Const R_DSC As Integer = 2
        'Const R_TYP As Integer = 3
        'Const R_ORD As Integer = 4
        'Const R_ACT As Integer = 5
        'Const R_NSO As Integer = 6  ' ** New Sort Order.
        Const R_NID As Integer = 7  ' ** New ID.
        'Const R_EIM As Integer = 8  ' ** Element# It Matches.
        'Const R_DEL As Integer = 9
        'Const R_FND As Integer = 10

        ' ** Array: arr_varStat().
        Const STAT_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const STAT_ORD As Integer = 0
        Const STAT_NAM As Integer = 1
        Const STAT_CNT As Integer = 2
        Const STAT_DSC As Integer = 3

610     If gblnDev_NoErrHandle = True Then
620   On Error GoTo 0
630     End If

640     intRetVal = 0
650     lngRecs = 0&

660     If blnContinue = True Then  ' ** Is a conversion.

670       If blnContinue = True Then  ' ** Conversion not already done.

680         lngTmp16 = 0&
690         ReDim arr_varStat(STAT_ELEMS, 0)

700         If lngStats > 0& Then
710           For lngX = 0& To (lngStats - 1&)
720             lngTmp16 = lngTmp16 + 1&
730             lngE = lngTmp16 - 1&
740             ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
750             For lngZ = 0& To STAT_ELEMS
760               arr_varStat(lngZ, lngE) = arr_varTmp03(lngZ, lngX)
770             Next  ' ** lngZ.
780           Next  ' ** lngX.
790         End If  ' ** lngStats.

800         lngTmp17 = 0&
810         ReDim arr_varDupeUnk(DU_ELEMS, 0)

820         If lngDupeUnks > 0& Then
830           For lngX = 0& To (lngDupeUnks - 1&)
840             lngTmp17 = lngTmp17 + 1&
850             lngE = lngTmp17 - 1&
860             ReDim Preserve arr_varDupeUnk(DU_ELEMS, lngE)
870             For lngZ = 0& To DU_ELEMS
880               arr_varDupeUnk(lngZ, lngE) = arr_varTmp04(lngZ, lngX)
890             Next  ' ** lngZ.
900           Next  ' ** lngX.
910         End If  ' ** lngDupeUnks.

920         lngTmp18 = 0&
930         ReDim arr_varAcct(A_ELEMS, 0)

940         If lngAccts > 0& Then
950           For lngX = 0& To (lngAccts - 1&)
960             lngTmp18 = lngTmp18 + 1&
970             lngE = lngTmp18 - 1&
980             ReDim Preserve arr_varAcct(A_ELEMS, lngE)
990             For lngZ = 0& To A_ELEMS
1000              arr_varAcct(lngZ, lngE) = arr_varTmp05(lngZ, lngX)
1010            Next  ' ** lngZ.
1020          Next  ' ** lngX.
1030        End If  ' ** lngAccts.

1040        If blnConvert_TrustDta = True Then

1050          If blnContinue = True Then  ' ** Workspace opens.

1060            With wrkLnk

1070              If blnContinue = True Then  ' ** Open dbsLnk.

1080                With dbsLnk

1090                  If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ** Get a list of Journal Types.
1100                    Set qdf = dbsLoc.QueryDefs("qryJournalType_02")
1110                    Set rstLoc1 = qdf.OpenRecordset
1120                    With rstLoc1
1130                      .MoveLast
1140                      lngJTypes = .RecordCount
1150                      .MoveFirst
1160                      arr_varJType = .GetRows(lngJTypes)
                          ' ************************************************
                          ' ** Array: arr_varJType()
                          ' **
                          ' **   Field  Element  Name           Constant
                          ' **   =====  =======  =============  ==========
                          ' **     1       0     journaltype    JT_TYP
                          ' **     2       1     description    JT_DSC
                          ' **     3       2     sortOrder      JT_SORT
                          ' **
                          ' ************************************************
1170                      .Close
1180                    End With  ' ** rstLoc1.

                        ' ******************************
                        ' ** Table: ledger.
                        ' ******************************

                        'THIS SHOULD DEAL WITH NEW REV CODE ID'S BELOW!
                        'THIS IS THE 1ST UPDATE OF THE REV CODE ID FOR ITS USAGE BEYOND THE REV CODE TABLE!
                        ' ** Step 16: ledger.
1190                    dblPB_ThisStep = 16#
1200                    Version_Status 3, dblPB_ThisStep, "Ledger"  ' ** Module Function: modVersionConvertFuncs1.

1210                    strCurrTblName = "ledger"
1220                    lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

1230                    blnFound = False: blnFound2 = False: lngRecs = 0&: strTmp05 = vbNullString
1240                    strTmp06 = vbNullString: strTmp07 = vbNullString: strTmp08 = vbNullString
1250                    strTmp09 = vbNullString: strTmp10 = vbNullString: strTmp11 = vbNullString
1260                    For lngX = 0& To (lngOldTbls - 1&)
1270                      If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
1280                        blnFound = True
1290                        Exit For
1300                      End If
1310                    Next

1320                    If blnFound = True Then
1330                      Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
1340                      With rstLnk
1350                        If .BOF = True And .EOF = True Then
                              ' ** This really really should have records!
1360                        Else
1370                          strCurrKeyFldName = "journalno"
1380                          lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
1390                          Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
1400                          Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Of 16 example TrustDta.mdb's, one has 20, some have 21 fields, and some have 25.
                              ' ** Current field count is 26 fields.
                              ' ** Table: ledger
                              ' **   ![journalno]        dbLong
                              ' **   ![journaltype]      dbText      'Check
                              ' **   ![assetno]          dbLong      'Check
                              ' **   ![transdate]        dbDate
                              ' **   ![postdate]         dbDate
                              ' **   ![accountno]        dbText      'Check
                              ' **   ![shareface]        dbDouble
                              ' **   ![due]              dbDate
                              ' **   ![rate]             dbDouble
                              ' **   ![pershare]         dbDouble
                              ' **   ![icash]            dbCurrency
                              ' **   ![pcash]            dbCurrency
                              ' **   ![cost]             dbCurrency
                              ' **   ![assetdate]        dbDate
                              ' **   ![description]      dbText
                              ' **   ![posted]           dbDate
                              ' **   ![taxcode]          dbInteger   'Check
                              ' **   ![Location_ID]      dbLong      'Check
                              ' **   ![RecurringItem]   dbText
                              ' **   ![purchaseDate]     dbDate
                              ' **   ![ledger_HIDDEN]    dbBoolean
                              ' **   ![revcode_ID]       dbLong      'Check
                              ' **   ![journal_USER]     dbText      'Check
                              ' **   ![CheckNum]         dbLong
                              ' **   ![CheckPaid]        dbBoolean
                              ' **   ![curr_id]  Defaults to 150.
                              ' ** Two previous versions have only 10 JournalTypes:
                              ' **              Deposit, Dividend, Interest, Liability, Misc., Paid, Purchase, Received, Sold, Withdrawn
                              ' ** All the rest have 11, the current number:
                              ' **   Cost Adj., Deposit, Dividend, Interest, Liability, Misc., Paid, Purchase, Received, Sold, Withdrawn
                              ' ** Missing fields in various of the 16 previous documented versions.
                              ' **   Field                Versions Missing It
                              ' **   ===================  ===================
                              ' **   CheckNum             5
                              ' **   CheckPaid            5
                              ' **   journal_USER         5
                              ' **   revcode_ID           5
                              ' ** [journal_USER] and, especially, [revcode_ID] must be updated!
                              ' ** Though this does have its own unique key, no other table refers to the journalno.
1410                          .MoveLast
1420                          lngRecs = .RecordCount
1430                          Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
1440                          .MoveFirst
1450                          lngFlds = 0&
1460                          For lngX = 0& To (lngOldFiles - 1&)
1470                            If arr_varOldFile(F_FNAM, lngX) = gstrFile_DataName Then
1480                              arr_varTmp02 = arr_varOldFile(F_T_ARR, lngX)
1490                              lngOldTbls = (UBound(arr_varTmp02, 2) + 1)
1500                              For lngY = 0& To (lngOldTbls - 1&)
1510                                If arr_varTmp02(T_TNAM, lngY) = strCurrTblName Then
1520                                  lngFlds = arr_varTmp02(T_FLDS, lngY)
1530                                  Exit For
1540                                End If
1550                              Next
1560                              Exit For
1570                            End If
1580                          Next
1590                          For Each fld In .Fields
1600                            With fld
1610                              If .Name = "Location Id" Or .Name = "ReoccurringItem" Then
                                    ' ** Has old field names.
1620                                blnFound2 = True
1630                                Exit For
1640                              End If
1650                            End With
1660                          Next
1670                          strTmp07 = "Location_ID": strTmp08 = "Location Id"
1680                          If blnFound2 = False Then strTmp06 = strTmp07 Else strTmp06 = strTmp08
1690                          strTmp10 = "RecurringItem": strTmp11 = "ReoccurringItem"
1700                          If blnFound2 = False Then strTmp09 = strTmp10 Else strTmp09 = strTmp11
1710                          For lngX = 1& To lngRecs
1720                            If Nz(![shareface], 0) = 0 And Nz(![ICash], 0) = 0 And Nz(![PCash], 0) = 0 And Nz(![Cost], 0) = 0 Then
                                  ' ** A really, really empty record! Skip.
1730                              lngLedgerEmptyDels = lngLedgerEmptyDels + 1&
1740                            Else
1750                              Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
                                  ' ** Add the record to the new table.
1760                              rstLoc1.AddNew
1770                              rstLoc1![journalno] = ![journalno]  ' ** We DO NOT runumber journalno!
1780                              blnFound = False
1790                              For lngY = 0& To (lngJTypes - 1&)
1800                                If arr_varJType(JT_TYP, lngY) = ![journaltype] Then
1810                                  blnFound = True
1820                                  Exit For
1830                                End If
1840                              Next
1850                              If blnFound = False Then
1860                                rstLoc1![journaltype] = "Misc."
1870                                blnFound = True  ' ** Reset.
1880                              Else
1890                                rstLoc1![journaltype] = ![journaltype]
1900                              End If
1910                              If IsNull(![assetno]) = False Then
1920                                If ![assetno] > 0& Then
1930                                  rstLoc2.MoveFirst
1940                                  rstLoc2.FindFirst "[tbl_name] = 'masterasset' And [fld_name] = 'assetno' And [key_lng_id1] = " & CStr(![assetno])
1950                                  If rstLoc2.NoMatch = True And ![assetno] = 1& Then
                                        ' ** It's the 'Accrued Interest Asset', masterasset_TYPE = 'IA',
                                        ' ** which doesn't get moved.
1960                                    rstLoc1![assetno] = 1&
1970                                  ElseIf rstLoc2.NoMatch = False Then
1980                                    rstLoc1![assetno] = rstLoc2![key_lng_id2]
1990                                  Else
                                        ' ** We'll have to assume that this is an orphan.
                                        ' ** Add the record to the new table.
2000                                    lngTmp13 = ![assetno]
2010                                    lngTmp14 = 0&
2020                                    Set rstLoc3 = dbsLoc.OpenRecordset("masterasset", dbOpenDynaset, dbConsistent)
2030                                    rstLoc3.AddNew
2040                                    lngDupeNum = lngDupeNum + 1&  ' ** Though not really a dupe, we need as unique number.
2050                                    rstLoc3![cusip] = "UNK" & CStr(lngDupeNum)
2060                                    rstLoc3![description] = "UNKNOWN {Ledger, Date " & _
                                          Format(![transdate], "mm/dd/yyyy") & ", Asset " & CStr(lngTmp13) & "}"
2070                                    rstLoc3![shareface] = ![shareface]  ' ** Do we want to update this when we're finished?
2080                                    rstLoc3![assettype] = "75"  ' ** Other.
2090                                    rstLoc3![rate] = ![rate]
2100                                    rstLoc3![due] = ![due]
2110                                    rstLoc3![marketvalue] = Null  'CCur(0)
2120                                    rstLoc3![marketvaluecurrent] = (![shareface] * Nz(![pershare], 0))
2130                                    rstLoc3![yield] = CDbl(0)
                                        'Null [assetdate]
                                        'Default to today.
2140                                    rstLoc3![currentDate] = CDate(Format(Nz(![assetdate], Date), "mm/dd/yyyy"))
2150                                    rstLoc3![masterasset_TYPE] = "RA"
2160                                    If lngFlds = 26& Then  ' ** The 26th field is curr_id.
2170  On Error Resume Next
2180                                      rstLoc3![curr_id] = ![curr_id]
2190                                      If ERR.Number <> 0 Then
2200  On Error GoTo ERRH
2210                                        rstLoc3![curr_id] = 150&  ' ** Default to USD.
2220                                      Else
2230  On Error GoTo ERRH
2240                                      End If
2250                                      If IsNull(rstLoc3![curr_id]) = True Then
2260                                        rstLoc3![curr_id] = 150&  ' ** Default to USD.
2270                                      Else
2280                                        If rstLoc3![curr_id] = 0& Then
2290                                          rstLoc3![curr_id] = 150&  ' ** Default to USD.
2300                                        End If
2310                                      End If
2320                                    Else
2330                                      rstLoc3![curr_id] = 150&  ' ** Default to USD.
2340                                    End If
2350                                    rstLoc3.Update
2360                                    lngTmp17 = lngTmp17 + 1&
2370                                    ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngTmp17 - 1&))
2380                                    arr_varDupeUnk(DU_TYP, (lngTmp17 - 1&)) = "UNK"
2390                                    arr_varDupeUnk(DU_TBL, (lngTmp17 - 1&)) = "masterasset"
2400                                    rstLoc3.Bookmark = rstLoc3.LastModified
2410                                    lngTmp14 = rstLoc3![assetno]
2420                                    rstLoc3.Close
2430                                    Set rstLoc3 = Nothing
2440                                    rstLoc1![assetno] = lngTmp14
                                        ' ** Add the ID cross-reference to tblVersion_Key.
2450                                    rstLoc2.AddNew
2460                                    rstLoc2![tbl_id] = lngCurrTblID
2470                                    rstLoc2![tbl_name] = "masterasset"  'strCurrTblName  'ledger
2480                                    rstLoc2![fld_id] = lngCurrKeyFldID
2490                                    rstLoc2![fld_name] = "assetno"  'strCurrKeyFldName
2500                                    rstLoc2![key_lng_id1] = lngTmp13
                                        'rstLoc2![key_txt_id1] =
2510                                    rstLoc2![key_lng_id2] = lngTmp14
                                        'rstLoc2![key_txt_id2] =
2520                                    rstLoc2.Update
2530                                    lngTmp13 = 0&: lngTmp14 = 0&
2540                                  End If
2550                                Else
2560                                  rstLoc1![assetno] = CLng(0)
2570                                End If
2580                              Else
2590                                rstLoc1![assetno] = CLng(0)
2600                              End If
2610                              rstLoc1![transdate] = ![transdate]
2620                              rstLoc1![postdate] = ![postdate]
                                  ' ** Since accountno's only get dropped if they're one of our 99 accounts
                                  ' ** or a dupe, this would still match one of the original ones.
                                  ' ** An orphan, however, requires serious attention.
2630                              strTmp05 = Trim(![accountno])
                                  ' ** 11/05/2009: New Stadelli check.
2640                              If Left(strTmp05, 3) = "99-" Then
2650                                strTmp05 = Mid(strTmp05, 4)
2660                                blnFound = False
2670                                For lngY = 0& To (lngTmp18 - 1&)
2680                                  If arr_varAcct(A_NUM, lngY) = strTmp05 Or arr_varAcct(A_NUM, lngY) = ("99-" & strTmp05) Then
2690                                    blnFound = True
2700                                    Exit For
2710                                  End If
2720                                Next
2730                              Else
2740                                blnFound = False
2750                                For lngY = 0& To (lngTmp18 - 1&)
2760                                  If arr_varAcct(A_NUM, lngY) = strTmp05 Then
2770                                    blnFound = True
2780                                    Exit For
2790                                  End If
2800                                Next
2810                              End If
2820                              If blnFound = False Then
                                    ' ** This could only mean it's an orphan.
2830                                Set rstLoc3 = dbsLoc.OpenRecordset("account", dbOpenDynaset, dbConsistent)
2840                                rstLoc3.AddNew
2850                                rstLoc3![accountno] = strTmp05  ' ** dbText 15
2860                                rstLoc3![shortname] = "UNKNOWN" & CStr(lngX)  ' ** dbText 30
2870                                rstLoc3![legalname] = "UNKNOWN {Ledger, " & _
                                      "Posting Date " & Format(![transdate], "mm/dd/yyyy") & "}" ' ** dbText 100
2880                                rstLoc3![accounttype] = "85"  ' ** Other.
2890                                rstLoc3![cotrustee] = "No"
2900                                rstLoc3![amendments] = "No"
2910                                rstLoc3![courtsupervised] = "No"
2920                                rstLoc3![discretion] = "No"
2930                                rstLoc3![ICash] = ![ICash]
2940                                rstLoc3![PCash] = ![PCash]
2950                                rstLoc3![Cost] = ![Cost]
2960                                rstLoc3![predate] = (![transdate] - 1)
2970                                rstLoc3![investmentobj] = "Other"
2980                                rstLoc3![numCopies] = CInt(1)
2990                                rstLoc3![account_SWEEP] = CBool(False)
3000                                rstLoc3![taxlot] = "0"
3010                                If lngFlds = 26& Then  ' ** The 26th field is curr_id.
3020  On Error Resume Next
3030                                  rstLoc3![curr_id] = ![curr_id]
3040                                  If ERR.Number <> 0 Then
3050  On Error GoTo ERRH
3060                                    rstLoc3![curr_id] = 150&  ' ** Default to USD.
3070                                  Else
3080  On Error GoTo ERRH
3090                                  End If
3100                                  If IsNull(rstLoc3![curr_id]) = True Then
3110                                    rstLoc3![curr_id] = 150&  ' ** Default to USD.
3120                                  Else
3130                                    If rstLoc3![curr_id] = 0& Then
3140                                      rstLoc3![curr_id] = 150&  ' ** Default to USD.
3150                                    End If
3160                                  End If
3170                                Else
3180                                  rstLoc3![curr_id] = 150&  ' ** Default to USD.
3190                                End If
3200                                rstLoc3.Update
3210                                lngTmp17 = lngTmp17 + 1&
3220                                ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngTmp17 - 1&))
3230                                arr_varDupeUnk(DU_TYP, (lngTmp17 - 1&)) = "UNK"
3240                                arr_varDupeUnk(DU_TBL, (lngTmp17 - 1&)) = "account"
3250                                rstLoc3.Close
3260                                Set rstLoc3 = Nothing
3270                                lngTmp18 = lngTmp18 + 1&
3280                                lngE = lngTmp18 - 1&
3290                                ReDim Preserve arr_varAcct(A_ELEMS, lngE)
3300                                arr_varAcct(A_NUM, lngE) = strTmp05
3310                                arr_varAcct(A_NUM_N, lngE) = "#ORPHAN_LG"
3320                                arr_varAcct(A_NAM, lngE) = "UNKNOWN" & CStr(lngX)
3330                                arr_varAcct(A_TYP, lngE) = "85"
3340                                arr_varAcct(A_ADMIN, lngE) = Null
3350                                arr_varAcct(A_ADMIN_N, lngE) = CLng(0)
3360                                arr_varAcct(A_SCHED, lngE) = Null
3370                                arr_varAcct(A_SCHED_N, lngE) = CLng(0)
3380                                arr_varAcct(A_DROPPED, lngE) = CBool(False)
3390                                arr_varAcct(A_ACCT99, lngE) = vbNullString
3400                                arr_varAcct(A_DASTNO, lngE) = "0"
3410                                rstLoc1![accountno] = strTmp05
3420                              Else
3430                                rstLoc1![accountno] = strTmp05
3440                              End If
3450                              rstLoc1![shareface] = Nz(![shareface], 0)
3460                              rstLoc1![due] = ![due]
3470                              rstLoc1![rate] = Nz(![rate], 0)
3480                              rstLoc1![pershare] = Nz(![pershare], 0)
3490                              rstLoc1![ICash] = Nz(![ICash], 0)
3500                              rstLoc1![PCash] = Nz(![PCash], 0)
3510                              rstLoc1![Cost] = Nz(![Cost], 0)
                                  ' ** Check lngTmp15 and arr_varTmp01()!
3520                              If lngTmp15 > 0& And rstLoc1![assetno] > 0& And IsNull(![assetdate]) = False Then
3530                                blnTmp27 = False
3540                                For lngZ = 0& To (lngTmp15 - 1&)
3550                                  If arr_varTmp01(0, lngZ) = rstLoc1![accountno] And _
                                          arr_varTmp01(1, lngZ) = rstLoc1![assetno] And _
                                          (arr_varTmp01(2, lngZ) = ![assetdate] Or arr_varTmp01(3, lngZ) = CDbl(![assetdate])) Then
3560                                    blnTmp27 = True
3570                                    rstLoc1![assetdate] = arr_varTmp01(4, lngZ)
3580                                    arr_varTmp01(5, lngZ) = CBool(True)
3590                                    Exit For
3600                                  End If
3610                                Next  ' ** lngZ.
3620                                If blnTmp27 = False Then
3630                                  rstLoc1![assetdate] = ![assetdate]
3640                                End If
3650                              Else
3660                                rstLoc1![assetdate] = ![assetdate]
3670                              End If
3680                              rstLoc1![description] = ![description]
3690                              If ![posted] = CDate("12/30/1899 12:00:00 AM") Then
3700                                rstLoc1![posted] = CDate(Format(![transdate], "mm/dd/yyyy") & " 9:00:00 AM")
3710                              Else
3720                                rstLoc1![posted] = ![posted]
3730                              End If
3740                              If IsNull(![taxcode]) = True Then
3750                                Select Case ![journaltype]
                                    Case "Dividend", "Interest", "Deposit", "Purchase", "Withdrawn", "Sold", "Received"
                                      ' ** Dividends are always INCOME for Tax Codes.
                                      ' ** Interest is always INCOME for Tax Codes.
                                      ' ** Purchase, Deposit are always INCOME for Tax Codes.
                                      ' ** Sold, Withdrawn are always INCOME for Tax Codes.
                                      ' ** Received is always INCOME for Tax Codes.
3760                                  rstLoc1![taxcode] = TAXID_INC
3770                                Case "Liability", "Paid"
                                      ' ** Liability is always EXPENSE for Tax Codes.
                                      ' ** Paid is always EXPENSE for Tax Codes.
3780                                  rstLoc1![taxcode] = TAXID_DED
3790                                Case "Cost Adj."
                                      ' ** Cost Adj. is INCOME if negative, EXPENSE if positive.
3800                                  If ![Cost] > 0 Then
3810                                    rstLoc1![taxcode] = TAXID_DED
3820                                  Else
3830                                    rstLoc1![taxcode] = TAXID_INC
3840                                  End If
3850                                Case "Misc."
                                      ' ** Misc. can be either INCOME or EXPENSE for Tax Codes.
3860                                  rstLoc1![taxcode] = TAXID_INC
3870                                End Select
3880                              Else
                                    ' **************************************************
                                    ' ** Array: arr_varTaxDefCode()
                                    ' **
                                    ' **   Field  Element  Name            Constant
                                    ' **   =====  =======  ==============  ===========
                                    ' **     1       0     taxcode_old     TD_ID_OLD
                                    ' **     3       2     taxcode_new     TD_ID_NEW
                                    ' **     2       1     discription     TD_DSC
                                    ' **     4       3     taxcode_type    TD_TYP
                                    ' **
                                    ' **************************************************
                                    ' ** If Ledger has old field names, then TaxCode is old, too.
3890                                If blnFound2 = True Then
3900                                  blnFound = False
3910                                  For lngY = 0& To (lngTaxDefCodes - 1&)
3920                                    If arr_varTaxDefCode(TD_ID_OLD, lngY) = ![taxcode] Then
3930                                      blnFound = True
3940                                      rstLoc1![taxcode] = arr_varTaxDefCode(TD_ID_NEW, lngY)
3950                                      Exit For
3960                                    End If
3970                                  Next
3980                                  If blnFound = False Then
                                        ' ** Most likely, their TaxCode is a Zero.
3990                                    Select Case ![journaltype]
                                        Case "Dividend", "Interest", "Deposit", "Purchase", "Withdrawn", "Sold", "Received"
                                          ' ** Dividends are always INCOME for Tax Codes.
                                          ' ** Interest is always INCOME for Tax Codes.
                                          ' ** Purchase, Deposit are always INCOME for Tax Codes.
                                          ' ** Sold, Withdrawn are always INCOME for Tax Codes.
                                          ' ** Received is always INCOME for Tax Codes.
4000                                      rstLoc1![taxcode] = TAXID_INC
4010                                    Case "Liability", "Paid"
                                          ' ** Liability is always EXPENSE for Tax Codes.
                                          ' ** Paid is always EXPENSE for Tax Codes.
4020                                      rstLoc1![taxcode] = TAXID_DED
4030                                    Case "Cost Adj."
                                          ' ** Cost Adj. is INCOME if negative, EXPENSE if positive.
4040                                      If ![Cost] > 0 Then
4050                                        rstLoc1![taxcode] = TAXID_DED
4060                                      Else
4070                                        rstLoc1![taxcode] = TAXID_INC
4080                                      End If
4090                                    Case "Misc."
                                          ' ** Misc. can be either INCOME or EXPENSE for Tax Codes.
4100                                      rstLoc1![taxcode] = TAXID_INC
4110                                    End Select
4120                                  End If
4130                                Else
4140                                  rstLoc1![taxcode] = ![taxcode]
4150                                End If  ' ** blnFound2.
4160                              End If  ' ** IsNull.
4170                              If IsNull(.Fields(strTmp06)) = False Then
4180                                If .Fields(strTmp06) > 0& Then
4190                                  rstLoc2.MoveFirst
4200                                  rstLoc2.FindFirst "[tbl_name] = 'Location' And [fld_name] = 'Location_ID' And " & _
                                        "[key_lng_id1] = " & CStr(Nz(.Fields(strTmp06), 0&))
4210                                  If rstLoc2.NoMatch = False Then
4220                                    rstLoc1![Location_ID] = rstLoc2![key_lng_id2]
4230                                  Else
                                        ' ** It's an orphan; don't bother creating one.
4240                                    rstLoc1![Location_ID] = CLng(1)  ' ** {Unassigned}.
4250                                  End If
4260                                Else
4270                                  rstLoc1![Location_ID] = CLng(1)  ' ** {Unassigned}.
4280                                End If
4290                              Else
4300                                rstLoc1![Location_ID] = CLng(1)  ' ** {Unassigned}.
4310                              End If
4320                              If IsNull(.Fields(strTmp09)) = False Then
4330                                If Trim(.Fields(strTmp09)) <> vbNullString Then
4340                                  rstLoc1![RecurringItem] = .Fields(strTmp09)
4350                                Else
4360                                  rstLoc1![RecurringItem] = Null
4370                                End If
4380                              Else
4390                                rstLoc1![RecurringItem] = Null
4400                              End If
                                  ' ** Check lngTmp15 and arr_varTmp01()!
4410                              If lngTmp15 > 0& And rstLoc1![assetno] > 0& And IsNull(![PurchaseDate]) = False Then
4420                                blnTmp27 = False
4430                                For lngZ = 0& To (lngTmp15 - 1&)
4440                                  If arr_varTmp01(0, lngZ) = rstLoc1![accountno] And _
                                          arr_varTmp01(1, lngZ) = rstLoc1![assetno] And _
                                          (arr_varTmp01(2, lngZ) = ![PurchaseDate] Or arr_varTmp01(3, lngZ) = CDbl(![PurchaseDate])) Then
4450                                    blnTmp27 = True
4460                                    rstLoc1![PurchaseDate] = arr_varTmp01(4, lngZ)
4470                                    arr_varTmp01(5, lngZ) = CBool(True)
4480                                    Exit For
4490                                  End If
4500                                Next  ' ** lngZ.
4510                                If blnTmp27 = False Then
4520                                  rstLoc1![PurchaseDate] = ![PurchaseDate]
4530                                End If
4540                              Else
4550                                rstLoc1![PurchaseDate] = ![PurchaseDate]
4560                              End If
4570                              If lngFlds = 25& Or lngFlds = 26& Then  ' ** The 26th field is curr_id.
                                    ' ** All current Ledger fields are present.
4580                                rstLoc2.MoveFirst
                                    ' ** Ledger: Cross-reference the revcode_ID.
4590                                lngTmp13 = 0&
4600                                For lngY = 0& To (lngRevCodes - 1&)
                                      'GO THROUGH THE ARRAY TO SEE IF IT'S GOT A NEW ID!
4610                                  If arr_varRevCode(R_ID, lngY) = ![revcode_ID] Then
                                        ' ** This'll find it whether it's been moved or was a dupe.
4620                                    If arr_varRevCode(R_NID, lngY) > 0& Then
                                          'R_NID SHOULD BE ZERO IF IT DIDN'T NEED TO BE MOVED!
4630                                      lngTmp13 = arr_varRevCode(R_NID, lngY)
4640                                    ElseIf ![journaltype] = "Dividend" And arr_varRevCode(R_ID, lngY) = REVID_INC Then
                                          ' ** If it's an unspecified Dividend, give it the new one.
4650                                      lngTmp13 = REVID_ORDDIV
4660                                    ElseIf ![journaltype] = "Interest" And arr_varRevCode(R_ID, lngY) = REVID_INC Then
                                          ' ** If it's an unspecified Interest, give it the new one.
4670                                      lngTmp13 = REVID_INTINC
4680                                    End If
4690                                    Exit For
4700                                  End If
4710                                Next
4720                                If lngTmp13 > 0& Then
4730                                  rstLoc1![revcode_ID] = lngTmp13
4740                                Else
                                      'IF IT'S HERE, IT WASN'T GIVEN A NEW ID!
                                      'ONLY GIVE IT ONE IF IT DOESN'T HAVE A GOOD ONE TO BEGIN WITH!
4750                                  blnFound = False
4760                                  For lngY = 0& To (lngRevCodes - 1&)
4770                                    If arr_varRevCode(R_ID, lngY) = ![revcode_ID] Then
                                          'WE ALREADY KNOW IT DIDN'T HAVE A NEW ONE ASSIGNED TO IT!
4780                                      blnFound = True
4790                                      Exit For
4800                                    End If
4810                                  Next
4820                                  Select Case blnFound
                                      Case True
                                        'IT'S KEEPING ITS ORIGINAL REV CODE ID!
4830                                    rstLoc1![revcode_ID] = ![revcode_ID]
4840                                  Case False
                                        ' ** Is this Income or Expense?
4850                                    If ![Cost] > 0@ Then
4860                                      lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
4870                                    ElseIf ![Cost] < 0@ Then
4880                                      lngTmp13 = REVID_INC  ' ** Unspecified Income.
4890                                    Else
4900                                      If ![ICash] > 0@ And ![PCash] = 0@ Then
4910                                        lngTmp13 = REVID_INC  ' ** Unspecified Income.
4920                                      Else
4930                                        If ![ICash] > 0@ Or ![PCash] > 0@ Then
4940                                          lngTmp13 = REVID_INC  ' ** Unspecified Income.
4950                                        ElseIf ![ICash] < 0@ Or ![PCash] < 0@ Then
4960                                          lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
4970                                        Else
4980                                          lngTmp13 = REVID_INC  ' ** Unspecified Income.
4990                                        End If
5000                                      End If
5010                                    End If
                                        '###############################################################################
                                        '###############################################################################
                                        'AHA! HERE'S WHERE IT WAS ARBITRARILY GIVEN A NEW REV CODE ID,
                                        'EVEN THOUGH IT'S ORIGINAL SHOULD STILL BE GOOD!!!
5020                                    rstLoc1![revcode_ID] = lngTmp13
                                        '###############################################################################
                                        '###############################################################################
5030                                  End Select
5040                                End If
5050                                If IsNull(![journal_USER]) = True Then
5060                                  rstLoc1![journal_USER] = "TAAdmin"
5070                                Else
5080                                  If Trim(![journal_USER]) = vbNullString Then
5090                                    rstLoc1![journal_USER] = "TAAdmin"
5100                                  Else
5110                                    If ![journal_USER] <> "TADemo" Then
5120                                      If ![journal_USER] = "Admin" Then
                                            ' ** Give it the new administrative user name.
5130                                        rstLoc1![journal_USER] = "TAAdmin"
5140                                      Else
5150                                        rstLoc1![journal_USER] = ![journal_USER]
5160                                      End If
5170                                    Else
                                          ' ** If this isn't one of our Demo's, and
                                          ' ** 'TADemo' shows up, change it to 'Admin'.
5180                                      blnFound = False
5190                                      For lngY = 0& To (lngTmp18 - 1&)
5200                                        If arr_varAcct(A_NUM, lngY) = "11" And _
                                                arr_varAcct(A_NAM, lngY) = "William B. Johnson Trust" Then
5210                                          blnFound = True
5220                                          Exit For
5230                                        End If
5240                                      Next
5250                                      If blnFound = True Then
                                            ' ** Leave it 'TADemo'.
5260                                        rstLoc1![journal_USER] = ![journal_USER]
5270                                      Else
                                            ' ** I might have inadvertantly done it when working on a client's data,
                                            ' ** or an early demo was converted to real with some existing user entries.
                                            ' ** Who knows...
5280                                        rstLoc1![journal_USER] = "TAAdmin"
5290                                      End If
5300                                      blnFound = True  ' ** Reset.
5310                                    End If
5320                                  End If
5330                                End If
5340                                rstLoc1![CheckNum] = ![CheckNum]
5350                                rstLoc1![CheckPaid] = ![CheckPaid]
5360                                rstLoc1![ledger_HIDDEN] = ![ledger_HIDDEN]  ' ** Moved down here when v1.6.3 discovered.
5370                                blnFound = False
5380                                For Each fld In .Fields
5390                                  With fld
5400                                    If .Name = "curr_id" Then
5410                                      blnFound = True
5420                                      rstLoc1![curr_id] = rstLnk![curr_id]
5430                                      Exit For
5440                                    End If
5450                                  End With
5460                                Next
5470                                If blnFound = False Then
5480                                  rstLoc1![curr_id] = 150&
                                      'WE'LL HAVE TO CHECK ALL THESE ONCE tblCurrency_History IS CONVERTED!
5490                                End If
5500                              Else
                                    ' ** This version is missing at least 4 fields, including 2 required ones.
                                    ' ** Is this Income or Expense?
5510                                lngTmp13 = 0&
5520                                Select Case ![journaltype]
                                    Case "Dividend"
5530                                  lngTmp13 = REVID_ORDDIV
5540                                Case "Interest"
5550                                  lngTmp13 = REVID_INTINC
5560                                Case Else
5570                                  If ![Cost] > 0@ Then
5580                                    lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
5590                                  ElseIf ![Cost] < 0@ Then
5600                                    lngTmp13 = REVID_INC  ' ** Unspecified Income.
5610                                  Else
5620                                    If ![ICash] > 0@ And ![PCash] = 0@ Then
5630                                      lngTmp13 = REVID_INC  ' ** Unspecified Income.
5640                                    Else
5650                                      If ![ICash] > 0@ Or ![PCash] > 0@ Then
5660                                        lngTmp13 = REVID_INC  ' ** Unspecified Income.
5670                                      ElseIf ![ICash] < 0@ Or ![PCash] < 0@ Then
5680                                        lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
5690                                      Else
5700                                        lngTmp13 = REVID_INC  ' ** Unspecified Income.
5710                                      End If
5720                                    End If
5730                                  End If
5740                                End Select
5750                                rstLoc1![revcode_ID] = lngTmp13
5760                                rstLoc1![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
5770                                rstLoc1![CheckNum] = Null
5780                                rstLoc1![CheckPaid] = False
5790                                If lngFlds = 21& Then
5800                                  rstLoc1![ledger_HIDDEN] = ![ledger_HIDDEN]  ' ** Moved down here when v1.6.3 discovered.
5810                                Else
                                      ' ** Version 1.6.3, or thereabouts.
5820                                  rstLoc1![ledger_HIDDEN] = False
5830                                End If
5840                              End If
                                  ' ** I will trust that journalno's are always unique,
                                  ' ** and that I've covered every required field!
5850                              rstLoc1.Update
5860                              If blnContinue = True Then
                                    ' ** The key field doesn't change, so no need to put it in tblVersion_Key.
5870                              Else
5880                                Exit For
5890                              End If
5900                            End If  ' ** Non-empty record.
5910                            strTmp05 = vbNullString: lngTmp13 = 0&: lngTmp14 = 0&
5920                            If lngX < lngRecs Then .MoveNext
5930                          Next
5940                          rstLoc1.Close
5950                          rstLoc2.Close
5960                        End If  ' ** Records present.
5970                        .Close
5980                      End With  ' ** rstLnk.
5990                    End If  ' ** blnFound.

6000                    lngTmp16 = lngTmp16 + 1&
6010                    lngE = lngTmp16 - 1&
6020                    ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
6030                    arr_varStat(STAT_ORD, lngE) = CInt(3)
6040                    arr_varStat(STAT_NAM, lngE) = "Ledger Entries: "
6050                    arr_varStat(STAT_CNT, lngE) = CLng(lngRecs)
6060                    arr_varStat(STAT_DSC, lngE) = vbNullString

6070                  End If  ' ** blnContinue.
6080                  strTmp05 = vbNullString: lngTmp13 = 0&: lngTmp14 = 0&
6090                  strTmp06 = vbNullString: strTmp07 = vbNullString: strTmp08 = vbNullString
6100                  strTmp09 = vbNullString: strTmp10 = vbNullString: strTmp11 = vbNullString

6110                  If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ******************************
                        ' ** Table: journal.
                        ' ******************************

                        'THIS SHOULD DEAL WITH NEW REV CODE ID'S BELOW!
                        'THIS IS THE 2ND UPDATE OF THE REV CODE ID FOR ITS USAGE BEYOND THE REV CODE TABLE!
                        ' ** Step 17: journal.
6120                    dblPB_ThisStep = 17#
6130                    Version_Status 3, dblPB_ThisStep, "Journal"  ' ** Module Function: modVersionConvertFuncs1.

                        ' ** We certainly advise them to Post before upgrading, but that may not follow our advice.

6140                    strCurrTblName = "journal"
6150                    lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

6160                    blnFound = False: blnFound2 = False: lngRecs = 0&: strTmp05 = vbNullString
6170                    strTmp06 = vbNullString: strTmp07 = vbNullString: strTmp08 = vbNullString
6180                    strTmp09 = vbNullString: strTmp10 = vbNullString: strTmp11 = vbNullString
6190                    For lngX = 0& To (lngOldTbls - 1&)
6200                      If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
6210                        blnFound = True
6220                        Exit For
6230                      End If
6240                    Next

6250                    If blnFound = True Then
6260                      Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
6270                      With rstLnk
6280                        If .BOF = True And .EOF = True Then
                              ' ** Good, they followed our advice!
6290                        Else
6300                          strCurrKeyFldName = "ID"
6310                          lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
6320                          Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
6330                          Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Of 16 example TrustDta.mdb's, some have 24 fields and some have 27.
                              ' ** Current field count is 28 fields.
                              ' ** Table: journal
                              ' **   ![ID]              dbLong
                              ' **   ![assetno]         dbLong
                              ' **   ![accountno]       dbText
                              ' **   ![shareface]       dbDouble
                              ' **   ![rate]            dbDouble
                              ' **   ![pershare]        dbDouble
                              ' **   ![due]             dbDate
                              ' **   ![assetdate]       dbDate
                              ' **   ![assettype]       dbText
                              ' **   ![journaltype]     dbText
                              ' **   ![journalSubtype]  dbText
                              ' **   ![Location_ID]     dbLong
                              ' **   ![transdate]       dbDate
                              ' **   ![icash]           dbCurrency
                              ' **   ![pcash]           dbCurrency
                              ' **   ![cost]            dbCurrency
                              ' **   ![description]     dbText
                              ' **   ![posted]          dbBoolean
                              ' **   ![purchaseDate]    dbDate
                              ' **   ![taxcode]         dbInteger
                              ' **   ![IsAverage]       dbBoolean
                              ' **   ![RecurringItem]  dbText
                              ' **   ![Reinvested]      dbBoolean
                              ' **   ![PrintCheck]      dbBoolean
                              ' **   ![revcode_ID]      dbLong
                              ' **   ![journal_USER]    dbText
                              ' **   ![CheckNum]        dbLong
                              ' **   ![curr_id]  Defaults to 150.
                              ' ** Missing fields in various of the 16 previous documented versions.
                              ' **   Field                Versions Missing It
                              ' **   ===================  ===================
                              ' **   CheckNum             5
                              ' **   journal_USER         5
                              ' **   revcode_ID           5
                              ' ** [journal_USER] and, especially, [revcode_ID] must be updated!
                              ' ** Though this does have its own unique key, no other table refers to it.
6340                          .MoveLast
6350                          lngRecs = .RecordCount
6360                          Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
6370                          .MoveFirst
6380                          lngFlds = 0&
6390                          For lngX = 0& To (lngOldFiles - 1&)
6400                            If arr_varOldFile(F_FNAM, lngX) = gstrFile_DataName Then
6410                              arr_varTmp02 = arr_varOldFile(F_T_ARR, lngX)
6420                              lngOldTbls = (UBound(arr_varTmp02, 2) + 1)
6430                              For lngY = 0& To (lngOldTbls - 1&)
6440                                If arr_varTmp02(T_TNAM, lngY) = strCurrTblName Then
6450                                  lngFlds = arr_varTmp02(T_FLDS, lngY)
6460                                  Exit For
6470                                End If
6480                              Next
6490                              Exit For
6500                            End If
6510                          Next
6520                          For Each fld In .Fields
6530                            With fld
6540                              If .Name = "location id" Or .Name = "ReoccurringItem" Then
                                    ' ** Has old field names.
6550                                blnFound2 = True
6560                                Exit For
6570                              End If
6580                            End With
6590                          Next
6600                          strTmp07 = "Location_ID": strTmp08 = "location id"
6610                          If blnFound2 = False Then strTmp06 = strTmp07 Else strTmp06 = strTmp08
6620                          strTmp10 = "RecurringItem": strTmp11 = "ReoccurringItem"
6630                          If blnFound2 = False Then strTmp09 = strTmp10 Else strTmp09 = strTmp11
6640                          For lngX = 1& To lngRecs
6650                            Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
                                ' ** Add the record to the new table.
6660                            rstLoc1.AddNew
6670                            If IsNull(![assetno]) = False Then
6680                              If ![assetno] > 0& Then
6690                                rstLoc2.MoveFirst
6700                                rstLoc2.FindFirst "[tbl_name] = 'masterasset' And " & _
                                      "[fld_name] = 'assetno' And [key_lng_id1] = " & CStr(![assetno])
6710                                If rstLoc2.NoMatch = True And ![assetno] = 1& Then
                                      ' ** It's the 'Accrued Interest Asset', masterasset_TYPE = 'IA',
                                      ' ** which doesn't get moved.
6720                                  rstLoc1![assetno] = 1&
6730                                ElseIf rstLoc2.NoMatch = False Then
6740                                  rstLoc1![assetno] = rstLoc2![key_lng_id2]
6750                                Else
                                      ' ** We'll have to assume that this is an orphan.
                                      ' ** Add the record to the new table.
6760                                  lngTmp13 = ![assetno]
6770                                  lngTmp14 = 0&
6780                                  Set rstLoc3 = dbsLoc.OpenRecordset("masterasset", dbOpenDynaset, dbConsistent)
6790                                  rstLoc3.AddNew
6800                                  lngDupeNum = lngDupeNum + 1&  ' ** Though not really a dupe, we need as unique number.
6810                                  rstLoc3![cusip] = "UNK" & CStr(lngDupeNum)
6820                                  rstLoc3![description] = "UNKNOWN {Journal, Date " & _
                                        Format(![transdate], "mm/dd/yyyy") & ", Asset " & CStr(lngTmp13) & "}"
6830                                  rstLoc3![shareface] = ![shareface]  ' ** Do we want to update this when we're finished?
6840                                  rstLoc3![assettype] = "75"  ' ** Other.
6850                                  rstLoc3![rate] = ![rate]
6860                                  rstLoc3![due] = ![due]
6870                                  rstLoc3![marketvalue] = Null  'CCur(0)
6880                                  rstLoc3![marketvaluecurrent] = (![shareface] * Nz(![pershare], 0))
6890                                  rstLoc3![yield] = CDbl(0)
6900                                  rstLoc3![currentDate] = CDate(Format(Nz(![assetdate], Date), "mm/dd/yyyy"))
6910                                  rstLoc3![masterasset_TYPE] = "RA"
6920                                  If lngFlds = 28& Then  ' ** The 28th field is curr_id.
6930  On Error Resume Next
6940                                    rstLoc3![curr_id] = ![curr_id]
6950                                    If ERR.Number <> 0 Then
6960  On Error GoTo ERRH
6970                                      rstLoc3![curr_id] = 150&  ' ** Default to USD.
6980                                    Else
6990  On Error GoTo ERRH
7000                                    End If
7010                                    If IsNull(rstLoc3![curr_id]) = True Then
7020                                      rstLoc3![curr_id] = 150&  ' ** Default to USD.
7030                                    Else
7040                                      If rstLoc3![curr_id] = 0& Then
7050                                        rstLoc3![curr_id] = 150&  ' ** Default to USD.
7060                                      End If
7070                                    End If
7080                                  Else
7090                                    rstLoc3![curr_id] = 150&  ' ** Default to USD.
7100                                  End If
7110                                  rstLoc3.Update
7120                                  lngTmp17 = lngTmp17 + 1&
7130                                  ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngTmp17 - 1&))
7140                                  arr_varDupeUnk(DU_TYP, (lngTmp17 - 1&)) = "UNK"
7150                                  arr_varDupeUnk(DU_TBL, (lngTmp17 - 1&)) = "masterasset"
7160                                  rstLoc3.Bookmark = rstLoc3.LastModified
7170                                  lngTmp14 = rstLoc3![assetno]
7180                                  rstLoc3.Close
7190                                  Set rstLoc3 = Nothing
7200                                  rstLoc1![assetno] = lngTmp14
                                      ' ** Add the ID cross-reference to tblVersion_Key.
7210                                  rstLoc2.AddNew
7220                                  rstLoc2![tbl_id] = lngCurrTblID
7230                                  rstLoc2![tbl_name] = "masterasset"  'strCurrTblName  'journal
7240                                  rstLoc2![fld_id] = lngCurrKeyFldID
7250                                  rstLoc2![fld_name] = "assetno"  'strCurrKeyFldName
7260                                  rstLoc2![key_lng_id1] = lngTmp13
                                      'rstLoc2![key_txt_id1] =
7270                                  rstLoc2![key_lng_id2] = lngTmp14
                                      'rstLoc2![key_txt_id2] =
7280                                  rstLoc2.Update
7290                                  lngTmp13 = 0&: lngTmp14 = 0&
7300                                End If
7310                              Else
7320                                rstLoc1![assetno] = 0&
7330                              End If
7340                            Else
7350                              rstLoc1![assetno] = 0&
7360                            End If
                                ' ** Since accountno's only get dropped if they're one of our 99 accounts
                                ' ** or a dupe, this would still match one of the original ones.
                                ' ** An orphan, however, requires serious attention.
7370                            strTmp05 = Trim(![accountno])
7380                            blnFound = False
7390                            For lngY = 0& To (lngTmp18 - 1&)
7400                              If arr_varAcct(A_NUM, lngY) = strTmp05 Then
7410                                blnFound = True
7420                                Exit For
7430                              End If
7440                            Next
7450                            If blnFound = False Then
                                  ' ** This could only mean it's an orphan.
7460                              Set rstLoc3 = dbsLoc.OpenRecordset("account", dbOpenDynaset, dbConsistent)
7470                              rstLoc3.AddNew
7480                              rstLoc3![accountno] = strTmp05  ' ** dbText 15
7490                              rstLoc3![shortname] = "UNKNOWN" & CStr(lngX)  ' ** dbText 30
7500                              rstLoc3![legalname] = "UNKNOWN {Journal " & _
                                    "Posting Date " & Format(![transdate], "mm/dd/yyyy") & "}" ' ** dbText 100
7510                              rstLoc3![accounttype] = "85"  ' ** Other.
7520                              rstLoc3![cotrustee] = "No"
7530                              rstLoc3![amendments] = "No"
7540                              rstLoc3![courtsupervised] = "No"
7550                              rstLoc3![discretion] = "No"
7560                              rstLoc3![ICash] = ![ICash]
7570                              rstLoc3![PCash] = ![PCash]
7580                              rstLoc3![Cost] = ![Cost]
7590                              rstLoc3![predate] = (![transdate] - 1)
7600                              rstLoc3![investmentobj] = "Other"
7610                              rstLoc3![numCopies] = CInt(1)
7620                              rstLoc3![account_SWEEP] = CBool(False)
7630                              rstLoc3![taxlot] = "0"
7640                              If lngFlds = 28& Then  ' ** The 28th field is curr_id.
7650  On Error Resume Next
7660                                rstLoc3![curr_id] = ![curr_id]
7670                                If ERR.Number <> 0 Then
7680  On Error GoTo ERRH
7690                                  rstLoc3![curr_id] = 150&  ' ** Default to USD.
7700                                Else
7710  On Error GoTo ERRH
7720                                End If
7730                                If IsNull(rstLoc3![curr_id]) = True Then
7740                                  rstLoc3![curr_id] = 150&  ' ** Default to USD.
7750                                Else
7760                                  If rstLoc3![curr_id] = 0& Then
7770                                    rstLoc3![curr_id] = 150&  ' ** Default to USD.
7780                                  End If
7790                                End If
7800                              Else
7810                                rstLoc3![curr_id] = 150&  ' ** Default to USD.
7820                              End If
7830                              rstLoc3.Update
7840                              lngTmp17 = lngTmp17 + 1&
7850                              ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngTmp17 - 1&))
7860                              arr_varDupeUnk(DU_TYP, (lngTmp17 - 1&)) = "UNK"
7870                              arr_varDupeUnk(DU_TBL, (lngTmp17 - 1&)) = "account"
7880                              rstLoc3.Close
7890                              Set rstLoc3 = Nothing
7900                              lngTmp18 = lngTmp18 + 1&
7910                              lngE = lngTmp18 - 1&
7920                              ReDim Preserve arr_varAcct(A_ELEMS, lngE)
7930                              arr_varAcct(A_NUM, lngE) = strTmp05
7940                              arr_varAcct(A_NUM_N, lngE) = "#ORPHAN_JR"
7950                              arr_varAcct(A_NAM, lngE) = "UNKNOWN" & CStr(lngX)
7960                              arr_varAcct(A_TYP, lngE) = "85"
7970                              arr_varAcct(A_ADMIN, lngE) = Null
7980                              arr_varAcct(A_ADMIN_N, lngE) = CLng(0)
7990                              arr_varAcct(A_SCHED, lngE) = Null
8000                              arr_varAcct(A_SCHED_N, lngE) = CLng(0)
8010                              arr_varAcct(A_DROPPED, lngE) = CBool(False)
8020                              arr_varAcct(A_ACCT99, lngE) = vbNullString
8030                              arr_varAcct(A_DASTNO, lngE) = "0"
8040                              rstLoc1![accountno] = strTmp05
8050                            Else
8060                              rstLoc1![accountno] = strTmp05
8070                            End If
8080                            rstLoc1![shareface] = Nz(![shareface], 0)
8090                            rstLoc1![rate] = Nz(![rate], 0)
8100                            rstLoc1![pershare] = Nz(![pershare], 0)
8110                            rstLoc1![due] = ![due]
8120                            rstLoc1![assetdate] = ![assetdate]
                                ' ** You know, I'm not sure if this ever gets populated?
8130                            rstLoc1![assettype] = ![assettype]
8140                            blnFound = False
8150                            For lngY = 0& To (lngJTypes - 1&)
8160                              If arr_varJType(JT_TYP, lngY) = ![journaltype] Then
8170                                blnFound = True
8180                                Exit For
8190                              End If
8200                            Next
8210                            If blnFound = False Then
8220                              rstLoc1![journaltype] = "Misc."
8230                              blnFound = True  ' ** Reset.
8240                            Else
8250                              rstLoc1![journaltype] = ![journaltype]
8260                            End If
8270                            If IsNull(![journalSubtype]) = False Then
8280                              If Trim(![journalSubtype]) <> vbNullString Then
                                    ' ** For the recent versions, this could only be 'Reinvest'.
                                    ' ** Should we check?
8290                                rstLoc1![journalSubtype] = ![journalSubtype]
8300                              Else
8310                                rstLoc1![journalSubtype] = Null
8320                              End If
8330                            Else
8340                              rstLoc1![journalSubtype] = Null
8350                            End If
8360                            rstLoc2.MoveFirst
8370                            rstLoc2.FindFirst "[tbl_name] = 'Location' And [fld_name] = 'Location_ID' And " & _
                                  "[key_lng_id1] = " & CStr(Nz(.Fields(strTmp06), 0&))
8380                            If rstLoc2.NoMatch = False Then
8390                              rstLoc1![Location_ID] = rstLoc2![key_lng_id2]
8400                            Else
                                  ' ** It's an orphan; don't bother creating one.
8410                              rstLoc1![Location_ID] = CLng(1)  ' ** {Unassigned}.
8420                            End If
8430                            rstLoc1![transdate] = ![transdate]
8440                            rstLoc1![ICash] = Nz(![ICash], 0)
8450                            rstLoc1![PCash] = Nz(![PCash], 0)
8460                            rstLoc1![Cost] = Nz(![Cost], 0)
8470                            rstLoc1![description] = ![description]
                                ' ** This field never gets used.
8480                            rstLoc1![posted] = False
                                ' ** Check lngTmp15 and arr_varTmp01()!
8490                            If lngTmp15 > 0& And rstLoc1![assetno] > 0& And IsNull(![PurchaseDate]) = False Then
8500                              blnTmp27 = False
8510                              For lngZ = 0& To (lngTmp15 - 1&)
8520                                If arr_varTmp01(0, lngZ) = rstLoc1![accountno] And _
                                        arr_varTmp01(1, lngZ) = rstLoc1![assetno] And _
                                        (arr_varTmp01(2, lngZ) = ![PurchaseDate] Or arr_varTmp01(3, lngZ) = CDbl(![PurchaseDate])) Then
8530                                  blnTmp27 = True
8540                                  rstLoc1![PurchaseDate] = arr_varTmp01(4, lngZ)
8550                                  arr_varTmp01(5, lngZ) = CBool(True)
8560                                  Exit For
8570                                End If
8580                              Next  ' ** lngZ.
8590                              If blnTmp27 = False Then
8600                                rstLoc1![PurchaseDate] = ![PurchaseDate]
8610                              End If
8620                            Else
8630                              rstLoc1![PurchaseDate] = ![PurchaseDate]
8640                            End If
8650                            If IsNull(![taxcode]) = True Then
8660                              Select Case ![journaltype]
                                  Case "Dividend", "Interest", "Deposit", "Purchase", "Withdrawn", "Sold", "Received"
                                    ' ** Dividends are always INCOME for Tax Codes.
                                    ' ** Interest is always INCOME for Tax Codes.
                                    ' ** Purchase, Deposit are always INCOME for Tax Codes.
                                    ' ** Sold, Withdrawn are always INCOME for Tax Codes.
                                    ' ** Received is always INCOME for Tax Codes.
8670                                rstLoc1![taxcode] = TAXID_INC
8680                              Case "Liability", "Paid"
                                    ' ** Liability is always EXPENSE for Tax Codes.
                                    ' ** Paid is always EXPENSE for Tax Codes.
8690                                rstLoc1![taxcode] = TAXID_DED
8700                              Case "Cost Adj."
                                    ' ** Cost Adj. is INCOME if negative, EXPENSE if positive.
8710                                If ![Cost] > 0 Then
8720                                  rstLoc1![taxcode] = TAXID_DED
8730                                Else
8740                                  rstLoc1![taxcode] = TAXID_INC
8750                                End If
8760                              Case "Misc."
                                    ' ** Misc. can be either INCOME or EXPENSE for Tax Codes.
8770                                rstLoc1![taxcode] = TAXID_INC
8780                              End Select
8790                            Else
                                  ' **************************************************
                                  ' ** Array: arr_varTaxDefCode()
                                  ' **
                                  ' **   Field  Element  Name            Constant
                                  ' **   =====  =======  ==============  ===========
                                  ' **     1       0     taxcode_old     TD_ID_OLD
                                  ' **     3       2     taxcode_new     TD_ID_NEW
                                  ' **     2       1     discription     TD_DSC
                                  ' **     4       3     taxcode_type    TD_TYP
                                  ' **
                                  ' **************************************************
                                  ' ** If Journal has old field names, then TaxCode is old, too.
8800                              If blnFound2 = True Then
8810                                blnFound = False
8820                                For lngY = 0& To (lngTaxDefCodes - 1&)
8830                                  If arr_varTaxDefCode(TD_ID_OLD, lngY) = ![taxcode] Then
8840                                    blnFound = True
8850                                    rstLoc1![taxcode] = arr_varTaxDefCode(TD_ID_NEW, lngY)
8860                                    Exit For
8870                                  End If
8880                                Next
8890                                If blnFound = False Then
                                      ' ** Most likely, their TaxCode is a Zero.
8900                                  Select Case ![journaltype]
                                      Case "Dividend", "Interest", "Deposit", "Purchase", "Withdrawn", "Sold", "Received"
                                        ' ** Dividends are always INCOME for Tax Codes.
                                        ' ** Interest is always INCOME for Tax Codes.
                                        ' ** Purchase, Deposit are always INCOME for Tax Codes.
                                        ' ** Sold, Withdrawn are always INCOME for Tax Codes.
                                        ' ** Received is always INCOME for Tax Codes.
8910                                    rstLoc1![taxcode] = TAXID_INC
8920                                  Case "Liability", "Paid"
                                        ' ** Liability is always EXPENSE for Tax Codes.
                                        ' ** Paid is always EXPENSE for Tax Codes.
8930                                    rstLoc1![taxcode] = TAXID_DED
8940                                  Case "Cost Adj."
                                        ' ** Cost Adj. is INCOME if negative, EXPENSE if positive.
8950                                    If ![Cost] > 0 Then
8960                                      rstLoc1![taxcode] = TAXID_DED
8970                                    Else
8980                                      rstLoc1![taxcode] = TAXID_INC
8990                                    End If
9000                                  Case "Misc."
                                        ' ** Misc. can be either INCOME or EXPENSE for Tax Codes.
9010                                    rstLoc1![taxcode] = TAXID_INC
9020                                  End Select
9030                                End If
9040                              Else
9050                                rstLoc1![taxcode] = ![taxcode]
9060                              End If  ' ** blnFound2.
9070                            End If  ' ** IsNull.
9080                            rstLoc1![IsAverage] = ![IsAverage]
9090                            If IsNull(.Fields(strTmp09)) = False Then
9100                              If Trim(.Fields(strTmp09)) <> vbNullString Then
9110                                rstLoc1![RecurringItem] = .Fields(strTmp09)
9120                              Else
9130                                rstLoc1![RecurringItem] = Null
9140                              End If
9150                            Else
9160                              rstLoc1![RecurringItem] = Null
9170                            End If
9180                            rstLoc1![Reinvested] = ![Reinvested]
9190                            rstLoc1![PrintCheck] = ![PrintCheck]
9200                            If lngFlds = 27& Or lngFlds = 28& Then
                                  ' ** All current Journal fields are present.
9210                              rstLoc2.MoveFirst
                                  ' ** Journal: Cross-reference the revcode_ID.
9220                              lngTmp13 = 0&
9230                              For lngY = 0& To (lngRevCodes - 1&)
                                    'GO THROUGH THE ARRAY TO SEE IF IT'S GOT A NEW ID!
9240                                If arr_varRevCode(R_ID, lngY) = ![revcode_ID] Then
                                      ' ** This'll find it whether it's been moved or was a dupe.
9250                                  If arr_varRevCode(R_NID, lngY) > 0& Then
                                        'R_NID SHOULD BE ZERO IF IT DIDN'T NEED TO BE MOVED!
9260                                    lngTmp13 = arr_varRevCode(R_NID, lngY)
9270                                  ElseIf ![journaltype] = "Dividend" And arr_varRevCode(R_ID, lngY) = REVID_INC Then
                                        ' ** If it's an unspecified Dividend, give it the new one.
9280                                    lngTmp13 = REVID_ORDDIV
9290                                  ElseIf ![journaltype] = "Interest" And arr_varRevCode(R_ID, lngY) = REVID_INC Then
                                        ' ** If it's an unspecified Interest, give it the new one.
9300                                    lngTmp13 = REVID_INTINC
9310                                  End If
9320                                  Exit For
9330                                End If
9340                              Next
9350                              If lngTmp13 > 0& Then
9360                                rstLoc1![revcode_ID] = lngTmp13
9370                              Else
                                    'IF IT'S HERE, IT WASN'T GIVEN A NEW ID!
                                    'ONLY GIVE IT ONE IF IT DOESN'T HAVE A GOOD ONE TO BEGIN WITH!
9380                                blnFound = False
9390                                For lngY = 0& To (lngRevCodes - 1&)
9400                                  If arr_varRevCode(R_ID, lngY) = ![revcode_ID] Then
                                        'WE ALREADY KNOW IT DIDN'T HAVE A NEW ONE ASSIGNED TO IT!
9410                                    blnFound = True
9420                                    Exit For
9430                                  End If
9440                                Next
9450                                Select Case blnFound
                                    Case True
                                      'IT'S KEEPING ITS ORIGINAL REV CODE ID!
9460                                  rstLoc1![revcode_ID] = ![revcode_ID]
9470                                Case False
9480                                  Select Case ![journaltype]
                                      Case "Dividend"
9490                                    lngTmp13 = REVID_ORDDIV
9500                                  Case "Interest"
9510                                    lngTmp13 = REVID_INTINC
9520                                  Case Else
                                        ' ** Is this Income or Expense?
9530                                    If ![Cost] > 0@ Then
9540                                      lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
9550                                    ElseIf ![Cost] < 0@ Then
9560                                      lngTmp13 = REVID_INC  ' ** Unspecified Income.
9570                                    Else
9580                                      If ![ICash] > 0@ And ![PCash] = 0@ Then
9590                                        lngTmp13 = REVID_INC  ' ** Unspecified Income.
9600                                      Else
9610                                        If ![ICash] > 0@ Or ![PCash] > 0@ Then
9620                                          lngTmp13 = REVID_INC  ' ** Unspecified Income.
9630                                        ElseIf ![ICash] < 0@ Or ![PCash] < 0@ Then
9640                                          lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
9650                                        Else
9660                                          lngTmp13 = REVID_INC  ' ** Unspecified Income.
9670                                        End If
9680                                      End If
9690                                    End If
9700                                  End Select
                                      '###############################################################################
                                      '###############################################################################
                                      'AHA! HERE'S WHERE IT WAS ARBITRARILY GIVEN A NEW REV CODE ID,
                                      'EVEN THOUGH IT'S ORIGINAL SHOULD STILL BE GOOD!!!
9710                                  rstLoc1![revcode_ID] = lngTmp13
                                      '###############################################################################
                                      '###############################################################################
9720                                End Select
9730                              End If
9740                              If IsNull(![journal_USER]) = True Then
9750                                rstLoc1![journal_USER] = "TAAdmin"
9760                              Else
9770                                If Trim(![journal_USER]) = vbNullString Then
9780                                  rstLoc1![journal_USER] = "TAAdmin"
9790                                Else
9800                                  If ![journal_USER] <> "TADemo" Then
9810                                    If ![journal_USER] = "Admin" Then
                                          ' ** Give it the new administrative user name.
9820                                      rstLoc1![journal_USER] = "TAAdmin"
9830                                    Else
9840                                      rstLoc1![journal_USER] = ![journal_USER]
9850                                    End If
9860                                  Else
                                        ' ** If this isn't one of our Demo's, and
                                        ' ** 'TADemo' shows up, change it to 'Admin'.
9870                                    blnFound = False
9880                                    For lngY = 0& To (lngTmp18 - 1&)
9890                                      If arr_varAcct(A_NUM, lngY) = "11" And _
                                              arr_varAcct(A_NAM, lngY) = "William B. Johnson Trust" Then
9900                                        blnFound = True
9910                                        Exit For
9920                                      End If
9930                                    Next
9940                                    If blnFound = True Then
                                          ' ** Leave it 'TADemo'.
9950                                      rstLoc1![journal_USER] = ![journal_USER]
9960                                    Else
                                          ' ** I might have inadvertantly done it when working on a client's data,
                                          ' ** or an early demo was converted to real with some existing user entries.
                                          ' ** Who knows...
9970                                      rstLoc1![journal_USER] = "TAAdmin"
9980                                    End If
9990                                    blnFound = True  ' ** Reset.
10000                                 End If
10010                               End If
10020                             End If
10030                             rstLoc1![CheckNum] = ![CheckNum]
10040                             blnFound = False
10050                             For Each fld In .Fields
10060                               With fld
10070                                 If .Name = "curr_id" Then
10080                                   blnFound = True
10090                                   rstLoc1![curr_id] = rstLnk![curr_id]
10100                                   Exit For
10110                                 End If
10120                               End With
10130                             Next
10140                             If blnFound = False Then
10150                               rstLoc1![curr_id] = 150&
                                    'WE'LL HAVE TO CHECK ALL THESE ONCE tblCurrency_History IS CONVERTED!
10160                             End If
10170                           Else
                                  ' ** This version is missing 4 fields, including 2 required ones.
10180                             lngTmp13 = 0&
10190                             Select Case ![journaltype]
                                  Case "Dividend"
10200                               lngTmp13 = REVID_ORDDIV
10210                             Case "Interest"
10220                               lngTmp13 = REVID_INTINC
10230                             Case Else
                                    ' ** Is this Income or Expense?
10240                               If ![Cost] > 0@ Then
10250                                 lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
10260                               ElseIf ![Cost] < 0@ Then
10270                                 lngTmp13 = REVID_INC  ' ** Unspecified Income.
10280                               Else
10290                                 If ![ICash] > 0@ And ![PCash] = 0@ Then
10300                                   lngTmp13 = REVID_INC  ' ** Unspecified Income.
10310                                 Else
10320                                   If ![ICash] > 0@ Or ![PCash] > 0@ Then
10330                                     lngTmp13 = REVID_INC  ' ** Unspecified Income.
10340                                   ElseIf ![ICash] < 0@ Or ![PCash] < 0@ Then
10350                                     lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
10360                                   Else
10370                                     lngTmp13 = REVID_INC  ' ** Unspecified Income.
10380                                   End If
10390                                 End If
10400                               End If
10410                             End Select
10420                             rstLoc1![revcode_ID] = lngTmp13
10430                             rstLoc1![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10440                             rstLoc1![CheckNum] = Null
10450                           End If
                                ' ** I will trust that journalno's are always unique,
                                ' ** and that I've covered every required field!
10460                           rstLoc1.Update
10470                           If blnContinue = True Then
                                  ' ** The key field doesn't change, so no need to put it in tblVersion_Key.
10480                           Else
10490                             Exit For
10500                           End If
10510                           strTmp05 = vbNullString: lngTmp13 = 0&: lngTmp14 = 0&
                                ' ** I think I've covered every required or linked field.
10520                           If blnContinue = True Then
                                  ' ** The key field , [ID], isn't referenced anywhere, so no need to put it in tblVersion_Key.
10530                           Else
10540                             Exit For
10550                           End If
10560                           strTmp05 = vbNullString: lngTmp13 = 0&: lngTmp14 = 0&
10570                           If lngX < lngRecs Then .MoveNext
10580                         Next
10590                         rstLoc1.Close
10600                         rstLoc2.Close
10610                       End If  ' ** Records present.
10620                       .Close
10630                     End With  ' ** rstLnk.
10640                   End If  ' ** blnFound.

10650                 End If  ' ** blnContinue.
10660                 strTmp05 = vbNullString: lngTmp13 = 0&: lngTmp14 = 0&
10670                 strTmp06 = vbNullString: strTmp07 = vbNullString: strTmp08 = vbNullString
10680                 strTmp09 = vbNullString: strTmp10 = vbNullString: strTmp11 = vbNullString

10690                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ******************************
                        ' ** Table: LedgerHidden.
                        ' ******************************

                        ' ** Step 18: LedgerHidden.
10700                   dblPB_ThisStep = 18#
10710                   Version_Status 3, dblPB_ThisStep, "Ledger Hidden"  ' ** Module Function: modVersionConvertFuncs1.

                        ' ** This would have to be a very new version indeed.
                        ' ** The only issue would be the accountno.

10720                   strCurrTblName = "LedgerHidden"
10730                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

10740                   blnFound = False: lngRecs = 0&
10750                   For lngX = 0& To (lngOldTbls - 1&)
10760                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
10770                       blnFound = True
10780                       Exit For
10790                     End If
10800                   Next

10810                   If blnFound = True Then
10820                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
10830                     With rstLnk
10840                       If .BOF = True And .EOF = True Then
                              ' ** Not used yet.
10850                       Else
10860                         strCurrKeyFldName = "accountno"
10870                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
10880                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
10890                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** No earlier versions have this table.
                              ' ** It would only be here for someone upgrading a v2.1.6
                              ' ** Table: LedgerHidden
                              ' **   ![hid_id]           dbLong
                              ' **   ![accountno]        dbText
                              ' **   ![journalno]        dbLong
                              ' **   ![hidtype]          dbText
                              ' **   ![hid_grpnum]       dbLong
                              ' **   ![hid_sort]         dbLong
                              ' **   ![hid_sortdate]     dbDate
                              ' **   ![uniqueid]         dbText
                              ' **   ![hid_order]        dbLong
                              ' **   ![Username]         dbText
                              ' **   ![hid_datecreated]  dbDate
                              ' **   ![hid_datemodified] dbDate
10900                         .MoveLast
10910                         lngRecs = .RecordCount
10920                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
10930                         .MoveFirst
10940                         For lngX = 1& To lngRecs
10950                           Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
                                ' ** Add the record to the new table.
10960                           strTmp05 = ![accountno]
10970                           blnFound = False
10980                           For lngY = 0& To (lngTmp18 - 1&)
10990                             If arr_varAcct(A_NUM, lngY) = strTmp05 Then
11000                               blnFound = True
11010                               Exit For
11020                             End If
11030                           Next
11040                           If blnFound = False Then
11050                             varTmp00 = DLookup("[accountno]", "ledger", "[journalno] = " & CStr(![journalno]))
11060                             If IsNull(varTmp00) = False Then
11070                               blnFound = True
11080                               strTmp05 = varTmp00
11090                             Else
                                    ' ** I don't think it'll ever get here, but if it does, just ditch 'em.
11100                             End If
11110                           End If
11120                           If blnFound = True Then
11130                             rstLoc1.AddNew
11140                             rstLoc1![accountno] = strTmp05
11150                             rstLoc1![journalno] = ![journalno]
11160                             rstLoc1![hidtype] = ![hidtype]
11170                             rstLoc1![hid_grpnum] = ![hid_grpnum]
11180                             rstLoc1![hid_sort] = ![hid_sort]
11190                             rstLoc1![hid_sortdate] = ![hid_sortdate]
11200                             rstLoc1![uniqueid] = ![uniqueid]
11210                             For Each fld In .Fields  ' ** rstLnk
11220                               If fld.Name = "hid_ord" Then
11230                                 rstLoc1![hid_order] = ![hid_ord]  ' ** v2.1.68: hid_ord, v2.1.69: hid_order.
11240                                 Exit For
11250                               ElseIf fld.Name = "hid_order" Then
11260                                 rstLoc1![hid_order] = ![hid_order]  ' ** v2.1.68: hid_ord, v2.1.69: hid_order.
11270                                 Exit For
11280                               End If
11290                             Next
11300                             rstLoc1![Username] = ![Username]
11310                             rstLoc1![hid_datecreated] = ![hid_datecreated]
11320                             rstLoc1![hid_datemodified] = ![hid_datemodified]
11330                             rstLoc1.Update
11340                           End If
11350                           If blnContinue = True Then
                                  ' ** The key field , [ID], isn't referenced anywhere, so no need to put it in tblVersion_Key.
11360                           Else
11370                             Exit For
11380                           End If
11390                           strTmp05 = vbNullString
11400                           If lngX < lngRecs Then .MoveNext
11410                         Next
11420                         rstLoc1.Close
11430                         rstLoc2.Close
11440                       End If  ' ** Records present.
11450                       .Close
11460                     End With  ' ** rstLnk.
11470                   End If  ' ** blnFound.

11480                 End If  ' ** blnContinue.

                      ' ** Also check for a mixture of prefixed and non-prefixed
                      ' ** numbers, depending on the value of gblnAccountNoWithType.
                      ' **   99-INCOME O/U -->  INCOME O/U
                      ' **   99-SUSPENSE   -->  SUSPENSE
                      ' ** Do this check at the end of the conversion.
                      'strAcct99_IncomeOU, strAcct99_Suspense, blnAcct99_Both

11490               End With  ' ** TrustDta.mdb: dbsLnk.

11500             End If  ' ** dbsLnk opens.

11510           End With  ' ** wrkLnk.

11520         End If  ' ** Workspace opens: blnContinue.

11530       End If  ' ** blnConvert_TrustDta.

11540       If blnContinue = False Then
11550         dbsLoc.Close
11560         wrkLoc.Close
11570       End If

11580       If lngTmp16 > lngStats Then
11590         lngStats = lngTmp16
11600         arr_varTmp03 = arr_varStat
11610       End If

11620       If lngTmp17 > lngDupeUnks Then
11630         lngDupeUnks = lngTmp17
11640         arr_varTmp04 = arr_varDupeUnk
11650       End If

11660       If lngTmp18 > lngAccts Then
11670         lngAccts = lngTmp18
11680         arr_varTmp05 = arr_varAcct
11690       End If

11700     End If  ' ** Conversion not already done.

11710   End If  ' ** Is a conversion.

EXITP:
11720   Version_Upgrade_05 = intRetVal
11730   Exit Function

ERRH:
11740   intRetVal = -9
11750   DoCmd.Hourglass False
11760   lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
11770   Select Case ERR.Number
        Case Else
11780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11790   End Select
11800   Resume EXITP

End Function

Public Function Version_Upgrade_06(blnContinue As Boolean, blnConvert_TrstArch As Boolean, intWrkType As Integer, lngDtaElem As Long, lngArchElem As Long, lngTrustDtaDbsID As Long, lngTrstArchDbsID As Long, lngAccts As Long, arr_varAcct As Variant, lngRevCodes As Long, arr_varRevCode As Variant, lngTaxDefCodes As Long, arr_varTaxDefCode As Variant, lngBadNames As Long, arr_varBadName As Variant, lngOldFiles As Long, arr_varOldFile As Variant, lngStats As Long, arr_varTmp03 As Variant, dblPB_ThisStep As Double, strKeyTbl As String, wrkLoc As DAO.Workspace, dbsLoc As DAO.Database) As Integer
' ** This handles just LedgerArchive.
' ** Tables converted here:
' **   LedgerArchive
' **   tblPricing_MasterAsset_History (tblAssetPricing)
' **   tblJournal_Memo
' **
' ** Return values:
' **    0  OK
' **   -6  Index/Key
' **   -7  Can't Open
' **   -9  Error
' **
' ** Called by:
' **   modVersionConvertFuncs1:
' **     Version_Upgrade_01()

        ' ** Version_Upgrade_06(
        ' **   blnContinue As Boolean, blnConvert_TrstArch As Boolean, intWrkType As Integer,
        ' **   lngDtaElem As Long, lngArchElem As Long, lngTrustDtaDbsID As Long, lngTrstArchDbsID As Long,
        ' **   lngAccts As Long, arr_varAcct As Variant,
        ' **   lngRevCodes As Long, arr_varRevCode As Variant, lngTaxDefCodes As Long, arr_varTaxDefCode As Variant,
        ' **   lngBadNames As Long, arr_varBadName As Variant,
        ' **   lngOldFiles As Long, arr_varOldFile As Variant, lngStats As Long, arr_varTmp03 As Variant,
        ' **   dblPB_ThisStep As Double, strKeyTbl As String, wrkLoc As DAO.Workspace, dbsLoc As DAO.Database
        ' ** ) As Integer

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Upgrade_06"

        Dim wrkLnk As DAO.Workspace, dbsLnk As DAO.Database
        Dim rstLnk As DAO.Recordset, rstLoc1 As DAO.Recordset, rstLoc2 As DAO.Recordset
        Dim tdf As DAO.TableDef, fld As DAO.Field
        Dim lngOldTbls As Long, arr_varOldTbl() As Variant, lngOldFlds As Long
        Dim blnOldHistoryTable As Boolean, blnBadName As Boolean, blnBadNameThis As Boolean
        Dim strCurrTblName As String, strCurrTblNameLocal As String, strCurrKeyFldName As String
        Dim lngLedgerArchEmptyDels As Long, arr_varStat() As Variant
        Dim lngRecs As Long, lngHistElem As Long, lngBadElem As Long, lngCurrTblID As Long, lngCurrKeyFldID As Long, lngFlds As Long
        Dim blnFound As Boolean, blnFound2 As Boolean
        Dim varTmp00 As Variant, arr_varTmp01 As Variant, arr_varTmp02 As Variant, strTmp04 As String, strTmp05 As String
        Dim strTmp06 As String, strTmp07 As String, strTmp08 As String, strTmp09 As String, strTmp10 As String
        Dim lngTmp13 As Long, lngTmp14 As Long, lngTmp15 As Long, blnTmp22 As Boolean, blnTmp27 As Boolean
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim intRetVal As Integer

        ' ** Array: arr_varBadName().
        'Const BN_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const BN_BAD   As Integer = 0
        Const BN_GOOD  As Integer = 1
        Const BN_FILE  As Integer = 2
        Const BN_TABLE As Integer = 3

        ' ** Array: arr_varOldFld().
        Const FD_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const FD_FNAM As Integer = 0
        Const FD_TYP  As Integer = 1
        Const FD_SIZ  As Integer = 2

        ' ** Array: arr_varTaxDefCode().
        Const TD_ID_OLD As Integer = 0
        Const TD_ID_NEW As Integer = 1
        'Const TD_DSC    As Integer = 2
        'Const TD_TYP    As Integer = 3

        ' ** Array: arr_varRevCode().
        'Const R_ELEMS As Integer = 10  ' ** Array's first-element UBound().
        'Const R_REC As Integer = 0
        Const R_ID  As Integer = 1
        'Const R_DSC As Integer = 2
        'Const R_TYP As Integer = 3
        'Const R_ORD As Integer = 4
        'Const R_ACT As Integer = 5
        'Const R_NSO As Integer = 6  ' ** New Sort Order.
        Const R_NID As Integer = 7  ' ** New ID.
        'Const R_EIM As Integer = 8  ' ** Element# It Matches.
        'Const R_DEL As Integer = 9
        'Const R_FND As Integer = 10

        ' ** Array: arr_varStat().
        Const STAT_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const STAT_ORD As Integer = 0
        Const STAT_NAM As Integer = 1
        Const STAT_CNT As Integer = 2
        Const STAT_DSC As Integer = 3

11910   If gblnDev_NoErrHandle = True Then
11920 On Error GoTo 0
11930   End If

11940   intRetVal = 0
11950   lngRecs = 0&

11960   If blnContinue = True Then  ' ** Is a conversion.

11970     If blnContinue = True Then  ' ** Conversion not already done.

11980       lngTmp14 = 0&
11990       ReDim arr_varStat(STAT_ELEMS, 0)

12000       For lngX = 0& To (lngStats - 1&)
12010         lngTmp14 = lngTmp14 + 1&
12020         lngE = lngTmp14 - 1&
12030         ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
12040         For lngZ = 0& To STAT_ELEMS
12050           arr_varStat(lngZ, lngE) = arr_varTmp03(lngZ, lngX)
12060         Next  ' ** lngZ.
12070       Next  ' ** lngX.

12080       If blnConvert_TrstArch = True And blnContinue = True Then

              ' ** Open the workspace with type found in Version_GetOldVer(), above.
12090         Select Case intWrkType
              Case 1
12100           Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
12110         Case 2
12120           Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
12130         Case 3
12140           Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
12150         Case 4
12160           Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
12170         Case 5
12180           Set wrkLnk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
12190         Case 6
12200           Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
12210         Case 7
12220           Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
12230         End Select

12240         If blnContinue = True Then  ' ** Workspace opens.

12250           With wrkLnk

12260 On Error Resume Next
12270             Set dbsLnk = .OpenDatabase(arr_varOldFile(F_PTHFIL, lngArchElem), False, True)  ' ** {pathfile}, {exclusive}, {read-only}
                  ' ** Not Exclusive, Read-Only.
12280             If ERR.Number <> 0 Then
                    ' ** Error opening old database.
12290               If gblnDev_NoErrHandle = True Then
12300 On Error GoTo 0
12310               Else
12320 On Error GoTo ERRH
12330               End If
12340               intRetVal = -7
12350               blnContinue = False
12360             Else
12370               If gblnDev_NoErrHandle = True Then
12380 On Error GoTo 0
12390               Else
12400 On Error GoTo ERRH
12410               End If

12420               With dbsLnk

12430                 lngOldTbls = 0&
12440                 ReDim arr_varOldTbl(T_ELEMS, 0)

12450                 blnBadName = False: blnBadNameThis = False: lngBadElem = -1&
12460                 For lngX = 0& To (lngBadNames - 1&)
12470                   If arr_varBadName(BN_FILE, lngX) = Parse_File(.Name) Then  ' ** Module Function: modFileUtilities.
                          ' ** This database may have a bad field name.
12480                     blnBadName = True
12490                     Exit For
12500                   End If
12510                 Next

12520                 For Each tdf In .TableDefs
12530                   With tdf
12540                     If Left(.Name, 4) <> "MSys" And Left(.Name, 4) <> "~TMP" And _
                              .Connect = vbNullString Then  ' ** Skip those pesky system tables.

12550                       lngOldTbls = lngOldTbls + 1&
12560                       lngY = lngOldTbls - 1&
12570                       ReDim Preserve arr_varOldTbl(T_ELEMS, lngY)
                            ' ******************************************
                            ' ** Array: arr_varOldTbl()
                            ' **
                            ' **   Element  Name           Constant
                            ' **   =======  =============  ===========
                            ' **      0     Name           T_TNAM
                            ' **      1     Fields         T_FLDS
                            ' **      2     Field Array    T_F_ARR
                            ' **
                            ' ******************************************
12580                       arr_varOldTbl(T_TNAM, lngY) = .Name
12590                       lngOldFlds = .Fields.Count
12600                       If lngOldFlds = 0& Then
12610                         arr_varOldFile(F_NOTE, lngArchElem) = arr_varOldFile(F_NOTE, lngArchElem) & _
                                " TBL: " & .Name & "  FLDS: " & CStr(lngOldFlds)
12620                         arr_varOldFile(F_NOTE, lngArchElem) = Trim(arr_varOldFile(F_NOTE, lngArchElem))
12630                       Else
12640                         arr_varOldTbl(T_FLDS, lngY) = lngOldFlds
12650                         arr_varOldTbl(T_F_ARR, lngY) = Empty
12660                         ReDim arr_varOldFld(FD_ELEMS, (lngOldFlds - 1&))
                              ' **********************************
                              ' ** Array: arr_varOldFld()
                              ' **
                              ' **   Element  Name    Constant
                              ' **   =======  ======  ==========
                              ' **      0     Name    FD_FNAM
                              ' **      1     Type    FD_TYP
                              ' **      2     Size    FD_SIZ
                              ' **
                              ' **********************************
12670                       End If

12680                       lngZ = -1&
12690                       For Each fld In .Fields
12700                         With fld
12710                           lngZ = lngZ + 1&
12720                           arr_varOldFld(FD_FNAM, lngZ) = .Name
12730                           arr_varOldFld(FD_TYP, lngZ) = .Type
12740                           arr_varOldFld(FD_SIZ, lngZ) = .Size
12750                         End With  ' ** fld.
12760                       Next

12770                       arr_varOldFile(F_TBLS, lngArchElem) = lngOldTbls
12780                       arr_varOldTbl(T_F_ARR, lngY) = arr_varOldFld

12790                     End If  ' ** Not a system table.
12800                   End With  ' ** This table: tdf.
12810                 Next  ' ** For each table: tdf.

12820                 arr_varOldFile(F_T_ARR, lngArchElem) = arr_varOldTbl

12830               End With  ' ** dbsLnk.

12840               With dbsLnk

12850                 If blnContinue = True Then

                        ' ******************************
                        ' ** Table: LedgerArchive.
                        ' ******************************

                        'THIS SHOULD DEAL WITH NEW REV CODE ID'S BELOW!
                        'THIS IS THE 2ND UPDATE OF THE REV CODE ID FOR ITS USAGE BEYOND THE REV CODE TABLE!
                        ' ** Step 19: LedgerArchive.
12860                   dblPB_ThisStep = 19#
12870                   Version_Status 3, dblPB_ThisStep, "Ledger Archive"  ' ** Module Function: modVersionConvertFuncs1.

12880                   strCurrTblName = "Ledger"  ' ** Its name in TrstArch.mdb
12890                   strCurrTblNameLocal = "LedgerArchive"  ' ** Its name here in Trust.mdb.
12900                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrstArchDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

12910                   blnFound = False: blnFound2 = False: strTmp04 = vbNullString: varTmp00 = vbNullString
12920                   strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString
12930                   strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString
12940                   For lngX = 0& To (lngOldTbls - 1&)
12950                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
12960                       blnFound = True
12970                       Exit For
12980                     End If
12990                   Next

13000                   If blnFound = True Then

13010                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
13020                     With rstLnk
13030                       If .BOF = True And .EOF = True Then
                              ' ** Not used yet.
13040                       Else

                              ' ************************************************
                              ' ** Array: arr_varBadName()
                              ' **
                              ' **   Element  Name                  Constant
                              ' **   =======  ====================  ==========
                              ' **      0     Wrong Field Name      BN_BAD
                              ' **      1     Correct Field name    BN_GOOD
                              ' **      2     Database Name         BN_FILE
                              ' **      3     Table Name            BN_TABLE
                              ' **
                              ' ************************************************
13050                         blnBadNameThis = False: lngBadElem = -1&: strTmp04 = vbNullString: varTmp00 = vbNullString
13060                         If blnBadName = True Then
                                ' ** This database may have a bad field name.
13070                           For lngX = 0& To (lngBadNames - 1&)
13080                             If arr_varBadName(BN_TABLE, lngX) = strCurrTblName Then
                                    ' ** This table may have a bad field name.
13090                               For Each fld In .Fields
13100                                 If fld.Name = arr_varBadName(BN_BAD, lngX) Then
                                        ' ** Yes, this table does have the bad name.
13110                                   blnBadNameThis = True
13120                                   lngBadElem = lngX
13130                                   strTmp04 = fld.Name                       ' ** Wrong name.
13140                                   varTmp00 = arr_varBadName(BN_GOOD, lngX)  ' ** Right name.
13150                                   Exit For
13160                                 End If
13170                               Next
13180                               If blnBadNameThis = True Then
13190                                 Exit For
13200                               End If
13210                             End If
13220                           Next
13230                         End If

13240                         strCurrKeyFldName = "journalno"
13250                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrstArchDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
13260                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblNameLocal, dbOpenDynaset, dbConsistent)
13270                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Of 16 example TrstArch.mdb's, some have 19 fields, some 21, some 24, and some have 25.
                              ' ** Current field count is 26 fields.
                              ' ** Table: ledger (LedgerArchive)
                              ' **   ![journalno]       dbLong      Req
                              ' **   ![journaltype]     dbText      Req
                              ' **   ![assetno]         dbLong      Req  0
                              ' **   ![transdate]       dbDate
                              ' **   ![postdate]        dbDate
                              ' **   ![accountno]       dbText      Req
                              ' **   ![shareface]       dbDouble         0
                              ' **   ![due]             dbDate
                              ' **   ![rate]            dbDouble         0
                              ' **   ![pershare]        dbDouble         0
                              ' **   ![icash]           dbCurrency       0
                              ' **   ![pcash]           dbCurrency       0
                              ' **   ![cost]            dbCurrency       0
                              ' **   ![assetdate]       dbDate
                              ' **   ![description]     dbText
                              ' **   ![posted]          dbDate      Req
                              ' **   ![taxcode]         dbInteger   Req  0
                              ' **   ![Location_ID]     dbLong      Req  1
                              ' **   ![RecurringItem]  dbText
                              ' **   ![revcode_ID]      dbLong      Req  1
                              ' **   ![journal_USER]    dbText
                              ' **   ![purchaseDate]    dbDate
                              ' **   ![CheckNum]        dbLong
                              ' **   ![CheckPaid]       dbBoolean   Req  False
                              ' **   ![ledger_HIDDEN]   dbBoolean   Req  False
                              ' **   ![curr_id]
                              ' ** Missing fields in various of the 16 previous documented versions.
                              ' **   Field            Versions Missing It
                              ' **   ===============  ===================
                              ' **   CheckNum         6
                              ' **   CheckPaid        6
                              ' **   journal_USER     5
                              ' **   ledger_HIDDEN    7
                              ' **   purchaseDate     6
                              ' **   revcode_ID       10  revcode_KD!
                              ' ** No tables directly reference this table.
13280                         .MoveLast
13290                         lngRecs = .RecordCount
13300                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
13310                         .MoveFirst
13320                         lngFlds = 0&
13330                         For lngX = 0& To (lngOldFiles - 1&)
13340                           If arr_varOldFile(F_FNAM, lngX) = gstrFile_ArchDataName Then
13350                             arr_varTmp02 = arr_varOldFile(F_T_ARR, lngX)
13360                             lngOldTbls = (UBound(arr_varTmp02, 2) + 1)
13370                             For lngY = 0& To (lngOldTbls - 1&)
13380                               If arr_varTmp02(T_TNAM, lngY) = strCurrTblName Then
13390                                 lngFlds = arr_varTmp02(T_FLDS, lngY)
13400                                 Exit For
13410                               End If
13420                             Next
13430                             Exit For
13440                           End If
13450                         Next
13460                         For Each fld In .Fields
13470                           With fld
13480                             If .Name = "Location Id" Or .Name = "ReoccurringItem" Then
                                    ' ** Has old field names.
13490                               blnFound2 = True
13500                               Exit For
13510                             End If
13520                           End With
13530                         Next
13540                         strTmp06 = "Location_ID": strTmp07 = "Location Id"
13550                         If blnFound2 = False Then strTmp05 = strTmp06 Else strTmp05 = strTmp07
13560                         strTmp09 = "RecurringItem": strTmp10 = "ReoccurringItem"
13570                         If blnFound2 = False Then strTmp08 = strTmp09 Else strTmp08 = strTmp10
13580                         For lngX = 1& To lngRecs
13590                           Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
13600                           If IsNull(![journalno]) = True Or IsNull(![journaltype]) = True Or IsNull(![accountno]) = True Then
                                  ' ** Skip.
13610                             lngLedgerArchEmptyDels = lngLedgerArchEmptyDels + 1&
13620                           ElseIf Nz(![shareface], 0) = 0 And Nz(![ICash], 0) = 0 And Nz(![PCash], 0) = 0 And Nz(![Cost], 0) = 0 Then
                                  ' ** A really, really empty record! Skip.
13630                             lngLedgerArchEmptyDels = lngLedgerArchEmptyDels + 1&
13640                           Else
                                  ' ** Add the record to the new table.
13650                             rstLoc1.AddNew
13660                             rstLoc1![journalno] = ![journalno]              'Req
13670                             rstLoc1![journaltype] = ![journaltype]          'Req
13680                             If IsNull(![assetno]) = False Then
13690                               If ![assetno] > 0& Then
13700                                 rstLoc2.MoveFirst
13710                                 rstLoc2.FindFirst "[tbl_name] = 'masterasset' And [fld_name] = 'assetno' And " & _
                                        "[key_lng_id1] = " & CStr(![assetno])
13720                                 If rstLoc2.NoMatch = True And ![assetno] = 1& Then
                                        ' ** It's the 'Accrued Interest Asset', masterasset_TYPE = 'IA',
                                        ' ** which doesn't get moved.
13730                                   rstLoc1![assetno] = 1&
13740                                 ElseIf rstLoc2.NoMatch = False Then
13750                                   rstLoc1![assetno] = rstLoc2![key_lng_id2] 'Req  0
13760                                 Else
                                        ' ** It may be an orphan, but it'll have to remain one!
13770                                   rstLoc1![assetno] = ![assetno]
13780                                 End If
13790                               Else
13800                                 rstLoc1![assetno] = 0&
13810                               End If
13820                             Else
13830                               rstLoc1![assetno] = 0&
13840                             End If
13850                             rstLoc1![transdate] = ![transdate]
13860                             rstLoc1![postdate] = ![postdate]
                                  ' ** We don't care about accountno in the archive.
13870                             rstLoc1![accountno] = ![accountno]              'Req
13880                             rstLoc1![shareface] = Nz(![shareface], 0)       '     0
13890                             rstLoc1![due] = ![due]
13900                             rstLoc1![rate] = Nz(![rate], 0)                 '     0
13910                             rstLoc1![pershare] = Nz(![pershare], 0)         '     0
13920                             rstLoc1![ICash] = Nz(![ICash], 0)               '     0
13930                             rstLoc1![PCash] = Nz(![PCash], 0)               '     0
13940                             rstLoc1![Cost] = Nz(![Cost], 0)                 '     0
                                  ' ** Check lngTmp15 and arr_varTmp01()!
13950                             If lngTmp15 > 0& And rstLoc1![assetno] > 0& And IsNull(![assetdate]) = False Then
13960                               blnTmp27 = False
13970                               For lngZ = 0& To (lngTmp15 - 1&)
13980                                 If arr_varTmp01(0, lngZ) = rstLoc1![accountno] And _
                                          arr_varTmp01(1, lngZ) = rstLoc1![assetno] And _
                                          (arr_varTmp01(2, lngZ) = ![assetdate] Or arr_varTmp01(3, lngZ) = CDbl(![assetdate])) Then
13990                                   blnTmp27 = True
14000                                   rstLoc1![assetdate] = arr_varTmp01(4, lngZ)
14010                                   arr_varTmp01(5, lngZ) = CBool(True)
14020                                   Exit For
14030                                 End If
14040                               Next  ' ** lngZ.
14050                               If blnTmp27 = False Then
14060                                 rstLoc1![assetdate] = ![assetdate]
14070                               End If
14080                             Else
14090                               rstLoc1![assetdate] = ![assetdate]
14100                             End If
14110                             If IsNull(![description]) = False Then
14120                               If Trim(![description]) <> vbNullString Then
14130                                 rstLoc1![description] = ![description]
14140                               Else
14150                                 rstLoc1![description] = Null
14160                               End If
14170                             Else
14180                               rstLoc1![description] = Null
14190                             End If
14200                             If IsNull(![posted]) = False Then
14210                               If ![posted] = CDate("12/30/1899 12:00:00 AM") Then
14220                                 rstLoc1![posted] = Date  ' ** Use today.
14230                               Else
14240                                 rstLoc1![posted] = ![posted]                'Req
14250                               End If
14260                             Else
14270                               rstLoc1![posted] = Date  ' ** Use today.
14280                             End If
14290                             If IsNull(![taxcode]) = True Then
14300                               Select Case ![journaltype]
                                    Case "Dividend", "Interest", "Deposit", "Purchase", "Withdrawn", "Sold", "Received"
                                      ' ** Dividends are always INCOME for Tax Codes.
                                      ' ** Interest is always INCOME for Tax Codes.
                                      ' ** Purchase, Deposit are always INCOME for Tax Codes.
                                      ' ** Sold, Withdrawn are always INCOME for Tax Codes.
                                      ' ** Received is always INCOME for Tax Codes.
14310                                 rstLoc1![taxcode] = TAXID_INC
14320                               Case "Liability", "Paid"
                                      ' ** Liability is always EXPENSE for Tax Codes.
                                      ' ** Paid is always EXPENSE for Tax Codes.
14330                                 rstLoc1![taxcode] = TAXID_DED
14340                               Case "Cost Adj."
                                      ' ** Cost Adj. is INCOME if negative, EXPENSE if positive.
14350                                 If ![Cost] > 0 Then
14360                                   rstLoc1![taxcode] = TAXID_DED
14370                                 Else
14380                                   rstLoc1![taxcode] = TAXID_INC
14390                                 End If
14400                               Case "Misc."
                                      ' ** Misc. can be either INCOME or EXPENSE for Tax Codes.
14410                                 rstLoc1![taxcode] = TAXID_INC
14420                               End Select
14430                             Else
                                    ' **************************************************
                                    ' ** Array: arr_varTaxDefCode()
                                    ' **
                                    ' **   Field  Element  Name            Constant
                                    ' **   =====  =======  ==============  ===========
                                    ' **     1       0     taxcode_old     TD_ID_OLD
                                    ' **     3       2     taxcode_new     TD_ID_NEW
                                    ' **     2       1     discription     TD_DSC
                                    ' **     4       3     taxcode_type    TD_TYP
                                    ' **
                                    ' **************************************************
                                    ' ** If Journal has old field names, then TaxCode is old, too.
14440                               If blnFound2 = True Then
14450                                 blnFound = False
14460                                 For lngY = 0& To (lngTaxDefCodes - 1&)
14470                                   If arr_varTaxDefCode(TD_ID_OLD, lngY) = ![taxcode] Then
14480                                     blnFound = True
14490                                     rstLoc1![taxcode] = arr_varTaxDefCode(TD_ID_NEW, lngY)
14500                                     Exit For
14510                                   End If
14520                                 Next
14530                                 If blnFound = False Then
                                        ' ** Most likely, their TaxCode is a Zero.
14540                                   Select Case ![journaltype]
                                        Case "Dividend", "Interest", "Deposit", "Purchase", "Withdrawn", "Sold", "Received"
                                          ' ** Dividends are always INCOME for Tax Codes.
                                          ' ** Interest is always INCOME for Tax Codes.
                                          ' ** Purchase, Deposit are always INCOME for Tax Codes.
                                          ' ** Sold, Withdrawn are always INCOME for Tax Codes.
                                          ' ** Received is always INCOME for Tax Codes.
14550                                     rstLoc1![taxcode] = TAXID_INC
14560                                   Case "Liability", "Paid"
                                          ' ** Liability is always EXPENSE for Tax Codes.
                                          ' ** Paid is always EXPENSE for Tax Codes.
14570                                     rstLoc1![taxcode] = TAXID_DED
14580                                   Case "Cost Adj."
                                          ' ** Cost Adj. is INCOME if negative, EXPENSE if positive.
14590                                     If ![Cost] > 0 Then
14600                                       rstLoc1![taxcode] = TAXID_DED
14610                                     Else
14620                                       rstLoc1![taxcode] = TAXID_INC
14630                                     End If
14640                                   Case "Misc."
                                          ' ** Misc. can be either INCOME or EXPENSE for Tax Codes.
14650                                     rstLoc1![taxcode] = TAXID_INC
14660                                   End Select
14670                                 End If
14680                               Else
14690                                 rstLoc1![taxcode] = ![taxcode]
14700                               End If  ' ** blnFound2.
14710                             End If
14720                             If IsNull(.Fields(strTmp05)) = False Then
14730                               If .Fields(strTmp05) > 0& Then
14740                                 rstLoc2.MoveFirst
14750                                 rstLoc2.FindFirst "[tbl_name] = 'Location' And [fld_name] = 'Location_ID' And " & _
                                        "[key_lng_id1] = " & CStr(Nz(.Fields(strTmp05), 0&))
14760                                 If rstLoc2.NoMatch = False Then
14770                                   rstLoc1![Location_ID] = rstLoc2![key_lng_id2]  'Req  1
14780                                 Else
                                        ' ** It may be an orphan, but it'll have to remain one!
14790                                   rstLoc1![Location_ID] = .Fields(strTmp05)
14800                                 End If
14810                               Else
14820                                 rstLoc1![Location_ID] = CLng(1)  ' ** {Unassigned}.
14830                               End If
14840                             Else
14850                               rstLoc1![Location_ID] = CLng(1)  ' ** {Unassigned}.
14860                             End If
14870                             If IsNull(.Fields(strTmp08)) = False Then
14880                               If Trim(.Fields(strTmp08)) <> vbNullString Then
14890                                 rstLoc1![RecurringItem] = .Fields(strTmp08)
14900                               Else
14910                                 rstLoc1![RecurringItem] = Null
14920                               End If
14930                             Else
14940                               rstLoc1![RecurringItem] = Null
14950                             End If
14960                             If lngFlds = 25& Or lngFlds = 26& Then
                                    ' ** All current LedgerArchive fields are present.
14970                               If blnBadNameThis = True Then
                                      ' ** strTmp04 and varTmp00 have already been set, above.
                                      'revcode_KD       !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
14980                               Else
14990                                 strTmp04 = "revcode_ID"
15000                                 varTmp00 = "revcode_ID"
15010                               End If
15020                               rstLoc2.MoveFirst
                                    ' ** LedgerArchive1: Cross-reference the revcode_ID.
15030                               lngTmp13 = 0&
15040                               For lngY = 0& To (lngRevCodes - 1&)
                                      'GO THROUGH THE ARRAY TO SEE IF IT'S GOT A NEW ID!
15050                                 If arr_varRevCode(R_ID, lngY) = .Fields(strTmp04) Then
                                        ' ** This'll find it whether it's been moved or was a dupe.
15060                                   If arr_varRevCode(R_NID, lngY) > 0& Then
                                          'R_NID SHOULD BE ZERO IF IT DIDN'T NEED TO BE MOVED!
15070                                     lngTmp13 = arr_varRevCode(R_NID, lngY)
15080                                   ElseIf ![journaltype] = "Dividend" And arr_varRevCode(R_ID, lngY) = REVID_INC Then
                                          ' ** If it's an unspecified Dividend, give it the new one.
15090                                     lngTmp13 = REVID_ORDDIV
15100                                   ElseIf ![journaltype] = "Interest" And arr_varRevCode(R_ID, lngY) = REVID_INC Then
                                          ' ** If it's an unspecified Interest, give it the new one.
15110                                     lngTmp13 = REVID_INTINC
15120                                   End If
15130                                   Exit For
15140                                 End If
15150                               Next
15160                               If lngTmp13 > 0& Then
15170                                 rstLoc1![revcode_ID] = lngTmp13               'Req  1
15180                               Else
                                      'IF IT'S HERE, IT WASN'T GIVEN A NEW ID!
                                      'ONLY GIVE IT ONE IF IT DOESN'T HAVE A GOOD ONE TO BEGIN WITH!
15190                                 blnFound = False
15200                                 For lngY = 0& To (lngRevCodes - 1&)
15210                                   If arr_varRevCode(R_ID, lngY) = ![revcode_ID] Then
                                          'WE ALREADY KNOW IT DIDN'T HAVE A NEW ONE ASSIGNED TO IT!
15220                                     blnFound = True
15230                                     Exit For
15240                                   End If
15250                                 Next
15260                                 Select Case blnFound
                                      Case True
                                        'IT'S KEEPING ITS ORIGINAL REV CODE ID!
15270                                   rstLoc1![revcode_ID] = ![revcode_ID]
15280                                 Case False
15290                                   Select Case ![journaltype]
                                        Case "Dividend"
15300                                     lngTmp13 = REVID_ORDDIV
15310                                   Case "Interest"
15320                                     lngTmp13 = REVID_INTINC
15330                                   Case Else
                                          ' ** Is this Income or Expense?
15340                                     If ![Cost] > 0@ Then
15350                                       lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
15360                                     ElseIf ![Cost] < 0@ Then
15370                                       lngTmp13 = REVID_INC  ' ** Unspecified Income.
15380                                     Else
15390                                       If ![ICash] > 0@ And ![PCash] = 0@ Then
15400                                         lngTmp13 = REVID_INC  ' ** Unspecified Income.
15410                                       Else
15420                                         If ![ICash] > 0@ Or ![PCash] > 0@ Then
15430                                           lngTmp13 = REVID_INC  ' ** Unspecified Income.
15440                                         ElseIf ![ICash] < 0@ Or ![PCash] < 0@ Then
15450                                           lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
15460                                         Else
15470                                           lngTmp13 = REVID_INC  ' ** Unspecified Income.
15480                                         End If
15490                                       End If
15500                                     End If
15510                                   End Select
                                        '###############################################################################
                                        '###############################################################################
                                        'AHA! HERE'S WHERE IT WAS ARBITRARILY GIVEN A NEW REV CODE ID,
                                        'EVEN THOUGH IT'S ORIGINAL SHOULD STILL BE GOOD!!!
15520                                   rstLoc1![revcode_ID] = lngTmp13
                                        '###############################################################################
                                        '###############################################################################
15530                                 End Select
15540                               End If
15550                               If IsNull(![journal_USER]) = False Then
15560                                 If Trim(![journal_USER]) <> vbNullString Then
15570                                   rstLoc1![journal_USER] = ![journal_USER]
15580                                 Else
15590                                   rstLoc1![journal_USER] = "System"
15600                                 End If
15610                               Else
15620                                 rstLoc1![journal_USER] = "System"
15630                               End If
                                    ' ** Check lngTmp15 and arr_varTmp01()!
15640                               If lngTmp15 > 0& And rstLoc1![assetno] > 0& And IsNull(![PurchaseDate]) = False Then
15650                                 blnTmp27 = False
15660                                 For lngZ = 0& To (lngTmp15 - 1&)
15670                                   If arr_varTmp01(0, lngZ) = rstLoc1![accountno] And _
                                            arr_varTmp01(1, lngZ) = rstLoc1![assetno] And _
                                            (arr_varTmp01(2, lngZ) = ![PurchaseDate] Or arr_varTmp01(3, lngZ) = CDbl(![PurchaseDate])) Then
15680                                     blnTmp27 = True
15690                                     rstLoc1![PurchaseDate] = arr_varTmp01(4, lngZ)
15700                                     arr_varTmp01(5, lngZ) = CBool(True)
15710                                     Exit For
15720                                   End If
15730                                 Next  ' ** lngZ.
15740                                 If blnTmp27 = False Then
15750                                   rstLoc1![PurchaseDate] = ![PurchaseDate]
15760                                 End If
15770                               Else
15780                                 rstLoc1![PurchaseDate] = ![PurchaseDate]
15790                               End If
15800                               rstLoc1![CheckNum] = ![CheckNum]
15810                               rstLoc1![CheckPaid] = ![CheckPaid]             'Req  False
15820                               rstLoc1![ledger_HIDDEN] = ![ledger_HIDDEN]     'Req  False
15830                               blnFound = False
15840                               For Each fld In .Fields
15850                                 With fld
15860                                   If .Name = "curr_id" Then
15870                                     blnFound = True
15880                                     rstLoc1![curr_id] = rstLnk![curr_id]
15890                                     Exit For
15900                                   End If
15910                                 End With
15920                               Next
15930                               If blnFound = False Then
15940                                 rstLoc1![curr_id] = 150&
                                      'WE'LL HAVE TO CHECK ALL THESE ONCE tblCurrency_History IS CONVERTED!
15950                               End If
15960                             Else
                                    ' ** Some fields are missing.
15970                               For Each fld In .Fields
15980                                 Select Case fld.Name
                                      Case "CheckNum"
15990                                   rstLoc1![CheckNum] = ![CheckNum]
16000                                 Case "CheckPaid"
16010                                   rstLoc1![CheckPaid] = ![CheckPaid]
16020                                 Case "journal_USER"
16030                                   If IsNull(![journal_USER]) = False Then
16040                                     If Trim(![journal_USER]) <> vbNullString Then
16050                                       If ![journal_USER] = "Admin" Then
                                              ' ** Give it the new administrative user name.
16060                                         rstLoc1![journal_USER] = "TAAdmin"
16070                                       Else
16080                                         rstLoc1![journal_USER] = ![journal_USER]
16090                                       End If
16100                                     Else
16110                                       rstLoc1![journal_USER] = "TAAdmin"
16120                                     End If
16130                                   Else
16140                                     rstLoc1![journal_USER] = "TAAdmin"
16150                                   End If
16160                                 Case "ledger_HIDDEN"
16170                                   rstLoc1![ledger_HIDDEN] = ![ledger_HIDDEN]
16180                                 Case "purchaseDate"
16190                                   rstLoc1![PurchaseDate] = ![PurchaseDate]
16200                                 Case strTmp04  ' ** revcode_KD / revcode_ID
16210                                   rstLoc2.MoveFirst
                                        ' ** LedgerArchive2: Cross-reference the revcode_ID.
16220                                   lngTmp13 = 0&
16230                                   For lngY = 0& To (lngRevCodes - 1&)
                                          'GO THROUGH THE ARRAY TO SEE IF IT'S GOT A NEW ID!
16240                                     If arr_varRevCode(R_ID, lngY) = .Fields(strTmp04) Then
                                            ' ** This'll find it whether it's been moved or was a dupe.
16250                                       If arr_varRevCode(R_NID, lngY) > 0& Then
                                              'R_NID SHOULD BE ZERO IF IT DIDN'T NEED TO BE MOVED!
16260                                         lngTmp13 = arr_varRevCode(R_NID, lngY)
16270                                       ElseIf ![journaltype] = "Dividend" And arr_varRevCode(R_ID, lngY) = REVID_INC Then
                                              ' ** If it's an unspecified Dividend, give it the new one.
16280                                         lngTmp13 = REVID_ORDDIV
16290                                       ElseIf ![journaltype] = "Interest" And arr_varRevCode(R_ID, lngY) = REVID_INC Then
                                              ' ** If it's an unspecified Interest, give it the new one.
16300                                         lngTmp13 = REVID_INTINC
16310                                       End If
16320                                       Exit For
16330                                     End If
16340                                   Next
16350                                   If lngTmp13 > 0& Then
16360                                     rstLoc1![revcode_ID] = lngTmp13               'Req  1
16370                                   Else
                                          'IF IT'S HERE, IT WASN'T GIVEN A NEW ID!
                                          'ONLY GIVE IT ONE IF IT DOESN'T HAVE A GOOD ONE TO BEGIN WITH!
16380                                     blnFound = False
16390                                     For lngY = 0& To (lngRevCodes - 1&)
16400                                       If arr_varRevCode(R_ID, lngY) = ![revcode_ID] Then
                                              'WE ALREADY KNOW IT DIDN'T HAVE A NEW ONE ASSIGNED TO IT!
16410                                         blnFound = True
16420                                         Exit For
16430                                       End If
16440                                     Next
16450                                     Select Case blnFound
                                          Case True
                                            'IT'S KEEPING ITS ORIGINAL REV CODE ID!
16460                                       rstLoc1![revcode_ID] = ![revcode_ID]
16470                                     Case False
16480                                       Select Case ![journaltype]
                                            Case "Dividend"
16490                                         lngTmp13 = REVID_ORDDIV
16500                                       Case "Interest"
16510                                         lngTmp13 = REVID_INTINC
16520                                       Case Else
                                              ' ** Is this Income or Expense?
16530                                         If ![Cost] > 0@ Then
16540                                           lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
16550                                         ElseIf ![Cost] < 0@ Then
16560                                           lngTmp13 = REVID_INC  ' ** Unspecified Income.
16570                                         Else
16580                                           If ![ICash] > 0@ And ![PCash] = 0@ Then
16590                                             lngTmp13 = REVID_INC  ' ** Unspecified Income.
16600                                           Else
16610                                             If ![ICash] > 0@ Or ![PCash] > 0@ Then
16620                                               lngTmp13 = REVID_INC  ' ** Unspecified Income.
16630                                             ElseIf ![ICash] < 0@ Or ![PCash] < 0@ Then
16640                                               lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
16650                                             Else
16660                                               lngTmp13 = REVID_INC  ' ** Unspecified Income.
16670                                             End If
16680                                           End If
16690                                         End If
16700                                       End Select
                                            '###############################################################################
                                            '###############################################################################
                                            'AHA! HERE'S WHERE IT WAS ARBITRARILY GIVEN A NEW REV CODE ID,
                                            'EVEN THOUGH IT'S ORIGINAL SHOULD STILL BE GOOD!!!
16710                                       rstLoc1![revcode_ID] = lngTmp13
                                            '###############################################################################
                                            '###############################################################################
16720                                     End Select
16730                                   End If
16740                                 End Select
16750                               Next
                                    ' ** journal_USER and revcode_ID are the only ones really requiring an entry.
                                    ' ** Any of the others that didn't hit above just don't matter.
16760                               If IsNull(rstLoc1![journal_USER]) = True Then
16770                                 rstLoc1![journal_USER] = "TAAdmin"
16780                               End If
16790                               If IsNull(rstLoc1![revcode_ID]) = True Then
                                      'IF IT'S STILL NULL, THEN YES, GIVE IT ONE!!
16800                                 Select Case rstLoc1![journaltype]
                                      Case "Dividend"
16810                                   lngTmp13 = REVID_ORDDIV
16820                                 Case "Interest"
16830                                   lngTmp13 = REVID_INTINC
16840                                 Case Else
                                        ' ** Is this Income or Expense?
16850                                   lngTmp13 = 0&
16860                                   If ![Cost] > 0@ Then
16870                                     lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
16880                                   ElseIf ![Cost] < 0@ Then
16890                                     lngTmp13 = REVID_INC  ' ** Unspecified Income.
16900                                   Else
16910                                     If ![ICash] > 0@ And ![PCash] = 0@ Then
16920                                       lngTmp13 = REVID_INC  ' ** Unspecified Income.
16930                                     Else
16940                                       If ![ICash] > 0@ Or ![PCash] > 0@ Then
16950                                         lngTmp13 = REVID_INC  ' ** Unspecified Income.
16960                                       ElseIf ![ICash] < 0@ Or ![PCash] < 0@ Then
16970                                         lngTmp13 = REVID_EXP  ' ** Unspecified Expense.
16980                                       Else
16990                                         lngTmp13 = REVID_INC  ' ** Unspecified Income.
17000                                       End If
17010                                     End If
17020                                   End If
17030                                 End Select
17040                                 rstLoc1![revcode_ID] = lngTmp13
17050                               End If
17060                             End If  ' ** lngFlds.
17070                             If IsNull(rstLoc1![curr_id]) = True Then
17080                               rstLoc1![curr_id] = 150&  ' ** Default to USD.
17090                             End If
17100 On Error Resume Next
17110                             rstLoc1.Update
17120                             If ERR.Number <> 0 Then
17130                               If gblnDev_NoErrHandle = True Then
17140 On Error GoTo 0
17150                               Else
17160 On Error GoTo ERRH
17170                               End If
                                    ' ** At this stage, I'm just going to let it go.
                                    ' ** The record will still be around in the BAK copy.
17180                               rstLoc1.CancelUpdate
17190                             Else
17200                               If gblnDev_NoErrHandle = True Then
17210 On Error GoTo 0
17220                               Else
17230 On Error GoTo ERRH
17240                               End If
17250                             End If
17260                           End If  ' ** Required fields present.
17270                           If lngX < lngRecs Then .MoveNext
17280                         Next  ' ** lngX.
17290                       End If  ' ** Records present.
17300                       .Close
17310                     End With  ' ** rstLnk.

17320                   End If  ' ** blnFound.

17330                 End If  ' ** blnContinue.
17340                 strTmp04 = vbNullString: varTmp00 = Empty
17350                 strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString
17360                 strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString

17370                 .Close
17380               End With  ' ** TrstArch.mdb: dbsLnk.

17390             End If  ' ** dbsLnk opens.

17400             .Close
17410           End With  ' ** wrkLnk.

17420         End If  ' ** Workspace opens: blnContinue.

17430       End If  ' ** blnConvert_TrstArch, blnContinue.

17440       lngTmp14 = lngTmp14 + 1& 'lngStats = lngStats + 1&
17450       lngE = lngTmp14 - 1&     'lngE = lngStats - 1&
17460       ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
17470       arr_varStat(STAT_ORD, lngE) = CInt(4)
17480       arr_varStat(STAT_NAM, lngE) = "Archived Ledger Entries: "
17490       arr_varStat(STAT_CNT, lngE) = CLng(lngRecs)
17500       arr_varStat(STAT_DSC, lngE) = vbNullString

17510       If blnContinue = True Then

              ' ** Open the workspace with type found in Version_GetOldVer(), above.
17520         Select Case intWrkType
              Case 1
17530           Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
17540         Case 2
17550           Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
17560         Case 3
17570           Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
17580         Case 4
17590           Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
17600         Case 5
17610           Set wrkLnk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
17620         Case 6
17630           Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
17640         Case 7
17650           Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
17660         End Select

17670         With wrkLnk

                ' *******************************************
                ' ** Table: tblPricing_MasterAsset_History.
                ' *******************************************

                ' ** Step 20: tblPricing_MasterAsset_History (tblAssetPricing).
17680           dblPB_ThisStep = 20#
17690           Version_Status 3, dblPB_ThisStep, "tblPricing_MasterAsset_History"  ' ** Module Function: modVersionConvertFuncs1.

17700           blnTmp22 = True: strTmp04 = vbNullString: lngTmp13 = 0&: lngHistElem = -1&: blnOldHistoryTable = False

                ' ** THIS MUST HANDLE ALL 3 POSSIBILITIES:
                ' **  1. THE OLDER tblAssetPricing TABLE.
                ' **  2. tblPricing_MasterAsset_History IN TRUST.MDB (v1.7.0 only).
                ' **  3. tblPricing_MasterAsset_History NOW IN TRUSTDTA.MDB (v1.7.1+).

17710           strCurrTblName = "tblPricing_MasterAsset_History"
17720           strCurrTblNameLocal = "tblPricing_MasterAsset_History"  'OR tblAssetPricing
17730           lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                  "[tbl_name] = '" & strCurrTblName & "'")

                ' ** See if we've got the old Trust.mde.
17740           For lngX = 0& To (lngOldFiles - 1&)
17750             If Rem_Ext(arr_varOldFile(F_FNAM, lngX)) = gstrFile_App Then
17760               lngHistElem = lngX
17770               Exit For
17780             End If
17790           Next

                ' ** Decide where to look based on strOldVersion.
17800           If Left(strOldVersion, 1) = "1" Or Left(strOldVersion, 3) = "2.0" Or Left(strOldVersion, 4) = "2.10" Then
                  ' ** If it's there, could only be tblAssetPricing in Trust.mde.
17810             blnOldHistoryTable = True
17820             If lngHistElem < 0& Then
17830               blnTmp22 = False
17840             End If
17850           Else
                  ' ** I believe it first appeared in 2.1.61.
17860             If Val(Mid(strOldVersion, 3, 3)) < 1.7 Then
                    ' ** tblAssetPricing in Trust.mde.
17870               blnOldHistoryTable = True
17880               If lngHistElem < 0& Then
17890                 blnTmp22 = False
17900               End If
17910             ElseIf strOldVersion = "2.1.70" Then
                    ' ** tblPricing_MasterAsset_History in Trust.mde.
17920               If lngHistElem < 0& Then
17930                 blnTmp22 = False
17940               End If
17950             Else
                    ' ** tblPricing_MasterAsset_History in TrustDta.mdb
17960               lngHistElem = lngDtaElem
17970             End If
17980           End If

17990           If blnTmp22 = True Then

18000 On Error Resume Next
18010             Set dbsLnk = .OpenDatabase(arr_varOldFile(F_PTHFIL, lngHistElem), False, True)  ' ** {pathfile}, {exclusive}, {read-only}
                  ' ** Not Exclusive, Read-Only.
18020             If ERR.Number <> 0 Then
                    ' ** Error opening old database.
18030               If gblnDev_NoErrHandle = True Then
18040 On Error GoTo 0
18050               Else
18060 On Error GoTo ERRH
18070               End If
18080               intRetVal = 0  ' ** An error here is inconsequential.
18090               blnTmp22 = False
18100             Else
18110               If gblnDev_NoErrHandle = True Then
18120 On Error GoTo 0
18130               Else
18140 On Error GoTo ERRH
18150               End If

18160               With dbsLnk

18170                 blnFound = False: blnFound2 = False
18180                 For Each tdf In .TableDefs
18190                   With tdf
18200                     Select Case blnOldHistoryTable
                          Case True
18210                       If .Name = "tblAssetPricing" Then
18220                         If .Connect = vbNullString Then  ' ** Make sure we're not seeing a linked table.
18230                           blnFound = True
18240                           strTmp04 = .Name
18250                         End If
18260                       End If
18270                     Case False
18280                       If .Name = "tblPricing_MasterAsset_History" Then
18290                         If .Connect = vbNullString Then  ' ** Make sure we're not seeing a linked table.
18300                           blnFound = True
18310                           strTmp04 = .Name
18320                         End If
18330                       End If
18340                     End Select
18350                   End With
18360                   If blnFound = True Then Exit For
18370                 Next

18380                 If blnFound = True Then
                        ' ** Now see if it's got any records to transfer.
18390                   Set rstLnk = .OpenRecordset(strTmp04, dbOpenDynaset, dbReadOnly)
18400                   With rstLnk
18410                     If .BOF = True And .EOF = True Then
                            ' ** Nope!
18420                       blnTmp22 = False
18430                     Else
18440                       .MoveLast
18450                       lngRecs = .RecordCount
18460                       Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
18470                       .MoveFirst
                            ' ** All versions of the table are identical to the current one.  NEW FIELDS!
18480                       Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblNameLocal, dbOpenDynaset, dbConsistent)
18490                       Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                            ' ** Table: tblPricing_MasterAsset_History (tblAssetPricing)
                            ' **   ![AssetPricing_ID]     dbLong      Req
                            ' **   ![assetno]             dbLong      Req
                            ' **   ![currentDate]         dbDate      Req
                            ' **   ![cusip]               dbText      Req
                            ' **   ![description]         dbText
                            ' **   ![totdesc]             dbText
                            ' **   ![shareface]           dbDouble
                            ' **   ![assettype]           dbText      Req
                            ' **   ![rate]                dbDouble
                            ' **   ![due]                 dbDate
                            ' **   ![marketvalue]         dbDouble
                            ' **   ![marketvaluecurrent]  dbDouble
                            ' **   ![yield]               dbDouble
                            ' **   ![masterasset_TYPE]    dbText      Req
                            ' **   ![curr_id]             dbLong           New
                            ' **   ![curr_code]           dbText           New
                            ' **   ![currsym_symbol]      dbText           New
                            ' **   ![curr_date]           dbDate           New
                            ' **   ![curr_rate1]          dbDouble         New
                            ' **   ![curr_rate2]          dbDouble         New
                            ' **   ![CurrentJournalUser]  dbText      Req
                            ' **   ![DateModified]        dbDate      Req
18500                       For lngX = 1& To lngRecs
18510                         Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
18520                         lngTmp13 = 0&
18530                         If IsNull(![assetno]) = False Then
18540                           If ![assetno] > 0& Then
18550                             rstLoc2.MoveFirst
18560                             rstLoc2.FindFirst "[tbl_name] = 'masterasset' And [fld_name] = 'assetno' And " & _
                                    "[key_lng_id1] = " & CStr(![assetno])
18570                             If rstLoc2.NoMatch = True And ![assetno] = 1& Then
                                    ' ** It's the 'Accrued Interest Asset', masterasset_TYPE = 'IA',
                                    ' ** which doesn't get moved.
18580                               lngTmp13 = 1&
18590                             ElseIf rstLoc2.NoMatch = False Then
18600                               lngTmp13 = rstLoc2![key_lng_id2]
18610                             Else
                                    ' ** Skip it.
18620                             End If
18630                           Else
                                  ' ** Skip it.
18640                           End If
18650                         Else
                                ' ** Skip it.
18660                         End If
18670                         If lngTmp13 > 0& Then
                                ' ** Add the record to the new table.
18680                           rstLoc1.AddNew
18690                           rstLoc1![assetno] = lngTmp13
18700                           rstLoc1![currentDate] = Nz(![currentDate], Date)
18710                           If IsNull(![cusip]) = True Then
18720                             rstLoc1![cusip] = DLookup("[cusip]", "masterasset", "[assetno] = " & CStr(lngTmp13))
18730                           Else
18740                             rstLoc1![cusip] = ![cusip]
18750                           End If
18760                           rstLoc1![description] = NullIfNullStr(![description])  ' ** Module Function: modStringFuncs.
18770                           rstLoc1![totdesc] = NullIfNullStr(![totdesc])  ' ** Module Function: modStringFuncs.
18780                           rstLoc1![shareface] = Nz(![shareface], 0)
18790                           If IsNull(![assettype]) = True Then
18800                             rstLoc1![assettype] = DLookup("[assettype]", "masterasset", "[assetno] = " & CStr(lngTmp13))
18810                           Else
18820                             rstLoc1![assettype] = ![assettype]
18830                           End If
18840                           rstLoc1![rate] = Nz(![rate], 0)
18850                           rstLoc1![due] = ![due]
18860                           rstLoc1![marketvalue] = ![marketvalue]
18870                           rstLoc1![marketvaluecurrent] = ![marketvaluecurrent]
18880                           rstLoc1![yield] = ![yield]
18890                           If IsNull(![masterasset_TYPE]) = True Then
18900                             rstLoc1![masterasset_TYPE] = DLookup("[masterasset_TYPE]", "masterasset", "[assetno] = " & CStr(lngTmp13))
18910                           Else
18920                             rstLoc1![masterasset_TYPE] = ![masterasset_TYPE]
18930                           End If
                                ' ** ![curr_id]
                                ' ** ![curr_code]
                                ' ** ![currsym_symbol]
                                ' ** ![curr_date]
                                ' ** ![curr_rate1]
                                ' ** ![curr_rate2]
18940                           blnFound2 = False
18950                           For Each fld In .Fields
18960                             If fld.Name = "curr_id" Then
18970                               blnFound2 = True
18980                               Exit For
18990                             End If
19000                           Next
19010                           Select Case blnFound2
                                Case True
19020                             Select Case IsNull(![curr_id])
                                  Case True
19030                               rstLoc1![curr_id] = 150&  ' ** Default to USD.
19040                               rstLoc1![curr_code] = "USD"
19050                               rstLoc1![currsym_symbol] = "$"
19060                               rstLoc1![curr_date] = Date
19070                               rstLoc1![curr_rate1] = 1#
19080                               rstLoc1![curr_rate2] = 0#
19090                             Case False
19100                               If ![curr_id] = 0& Then
19110                                 rstLoc1![curr_id] = 150&  ' ** Default to USD.
19120                                 rstLoc1![curr_code] = "USD"
19130                                 rstLoc1![currsym_symbol] = "$"
19140                                 rstLoc1![curr_date] = Date
19150                                 rstLoc1![curr_rate1] = 1#
19160                                 rstLoc1![curr_rate2] = 0#
19170                               Else
19180                                 rstLoc1![curr_id] = ![curr_id]
19190                                 rstLoc1![curr_code] = ![curr_code]
19200                                 rstLoc1![currsym_symbol] = ![currsym_symbol]
19210                                 rstLoc1![curr_date] = ![curr_date]
19220                                 rstLoc1![curr_rate1] = ![curr_rate1]
19230                                 rstLoc1![curr_rate2] = ![curr_rate2]
19240                               End If
19250                             End Select
19260                           Case False
19270                             rstLoc1![curr_id] = 150&  ' ** Default to USD.
19280                             rstLoc1![curr_code] = "USD"
19290                             rstLoc1![currsym_symbol] = "$"
19300                             rstLoc1![curr_date] = Date
19310                             rstLoc1![curr_rate1] = 1#
19320                             rstLoc1![curr_rate2] = 0#
19330                           End Select
19340                           rstLoc1![CurrentJournalUser] = Nz(![CurrentJournalUser], CurrentUser)  ' ** Internal Access Function: Trust Accountant login.
19350                           rstLoc1![DateModified] = Nz(![DateModified], Now())
19360 On Error Resume Next
19370                           rstLoc1.Update
19380                           If ERR.Number <> 0 Then
19390                             If gblnDev_NoErrHandle = True Then
19400 On Error GoTo 0
19410                             Else
19420 On Error GoTo ERRH
19430                             End If
                                  ' ** At this stage, I'm just going to let it go.
                                  ' ** The record will still be around in the BAK copy.
19440                             rstLoc1.CancelUpdate
19450                           Else
19460                             If gblnDev_NoErrHandle = True Then
19470 On Error GoTo 0
19480                             Else
19490 On Error GoTo ERRH
19500                             End If
19510                           End If
19520                         End If  ' ** Required fields present.
19530                         If lngX < lngRecs Then .MoveNext
19540                       Next  ' ** lngX.
19550                       rstLoc1.Close
19560                       rstLoc2.Close
19570                     End If  ' ** Records present.
19580                     .Close
19590                   End With  ' ** rstLnk.
19600                 Else
                        ' ** No big deal.
19610                   blnTmp22 = False
19620                 End If  ' ** blnFound.

19630               End With  ' ** dbsLnk.

19640             End If  ' ** dbsLnk opens.

19650           Else
                  ' ** If their Trust.mde wasn't found, it's too old to matter.
19660           End If  ' ** blnTmp22.

19670           If blnContinue = True Then

19680             Set dbsLnk = .OpenDatabase(arr_varOldFile(F_PTHFIL, lngDtaElem), False, True)  ' ** {pathfile}, {exclusive}, {read-only}
                  ' ** Not Exclusive, Read-Only.

19690             If blnContinue = True Then  ' ** Database opens.

19700               With dbsLnk

                      ' *******************************************
                      ' ** Table: tblJournal_Memo.
                      ' *******************************************

                      'THIS SHOULD ONLY HAVE ENTRIES FOR THE CURRENT JOURNAL!
                      'IF THE JOURNAL IS EMPTY, SKIP THIS ENTIRELY!

                      ' ** Step 21: tblJournal_Memo.
19710                 dblPB_ThisStep = 21#
19720                 Version_Status 3, dblPB_ThisStep, "tblJournal_Memo"  ' ** Module Function: modVersionConvertFuncs1.

19730                 varTmp00 = DCount("*", "journal")
19740                 If varTmp00 = 0 Then
                        ' ** Skip this whole section.
19750                 Else

19760                   strTmp04 = vbNullString

19770                   strCurrTblName = "tblJournal_Memo"
19780                   strCurrTblNameLocal = "tblJournal_Memo"
19790                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

19800                   blnFound = False: lngRecs = 0&
19810                   For lngX = 0& To (lngOldTbls - 1&)
19820                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
19830                       blnFound = True
19840                       Exit For
19850                     End If
19860                   Next

19870                   If blnFound = True Then
19880                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
19890                     With rstLnk
19900                       If .BOF = True And .EOF = True Then
                              ' ** Not used yet.
19910                       Else
19920                         strCurrKeyFldName = "JrnlMemo_ID"
19930                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
19940                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
19950                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** No earlier versions have this table.
                              ' ** Table: tblJournal_Memo
                              ' **   ![Journal_ID]             dbLong    This corresponds to the Journal's ID.
                              ' **   ![JrnlMemo_ID]            dbLong    AutoNumber.
                              ' **   ![journaltype]            dbText
                              ' **   ![accountno]              dbText
                              ' **   ![transdate]              dbDate
                              ' **   ![JrnlMemo_Memo]          dbText
                              ' **   ![JrnlMemo_DateModified]  dbDate
19960                         .MoveLast
19970                         lngRecs = .RecordCount
19980                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
19990                         .MoveFirst
20000                         For lngX = 1& To lngRecs
20010                           varTmp00 = DCount("*", "journal", "[ID] = " & CStr(![Journal_ID]))
20020                           If varTmp00 = 1 Then
20030                             Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
                                  ' ** Add the record to the new table.
20040                             strTmp04 = ![accountno]
20050                             blnFound = False
20060                             For lngY = 0& To (lngAccts - 1&)
20070                               If arr_varAcct(A_NUM, lngY) = strTmp04 Then
20080                                 blnFound = True
20090                                 Exit For
20100                               End If
20110                             Next
20120                             If blnFound = True Then
20130                               rstLoc1.AddNew
20140                               rstLoc1![Journal_ID] = ![Journal_ID]
20150                               rstLoc1![journaltype] = ![journaltype]
20160                               rstLoc1![accountno] = ![accountno]
20170                               rstLoc1![transdate] = ![transdate]
20180                               rstLoc1![JrnlMemo_Memo] = ![JrnlMemo_Memo]
20190                               rstLoc1![JrnlMemo_DateModified] = ![JrnlMemo_DateModified]
20200                               rstLoc1.Update
20210                             Else
                                    ' ** No big deal, just throw it out.
20220                             End If
20230                             If blnContinue = True Then  ' ** Nothing sets this False, anyway!
                                    ' ** The key field isn't referenced anywhere, so no need to put it in tblVersion_Key.
20240                             Else
20250                               Exit For
20260                             End If
20270                           End If  ' ** varTmp00.
20280                           strTmp04 = vbNullString
20290                           If lngX < lngRecs Then .MoveNext
20300                         Next
20310                         rstLoc1.Close
20320                         rstLoc2.Close
20330                       End If  ' ** Records present.
20340                       .Close
20350                     End With  ' ** rstLnk.
20360                   End If  ' ** blnFound.

20370                 End If  ' ** varTmp00.

20380               End With  ' ** TrustDta.mdb: dbsLnk.

20390             End If  ' ** dbsLnk opens.

20400           End If  ' ** blnContinue.

20410         End With  ' ** wrkLnk.

20420       End If  ' ** Workspace opens: blnContinue.

20430       If blnContinue = False Then
20440         dbsLoc.Close
20450         wrkLoc.Close
20460       End If

20470       If lngTmp14 > lngStats Then
20480         lngStats = lngTmp14
20490         arr_varTmp03 = arr_varStat
20500       End If

20510     End If  ' ** Conversion not already done.

20520   End If  ' ** Is a conversion.

20530   DoCmd.Hourglass False

EXITP:
20540   Set fld = Nothing
20550   Set tdf = Nothing
20560   Set rstLoc1 = Nothing
20570   Set rstLoc2 = Nothing
20580   Set rstLnk = Nothing
20590   Set dbsLnk = Nothing
20600   Set wrkLnk = Nothing
20610   Version_Upgrade_06 = intRetVal
20620   Exit Function

ERRH:
20630   intRetVal = -9
20640   DoCmd.Hourglass False
20650   lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
20660   Select Case ERR.Number
        Case Else
20670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20680   End Select
20690   Resume EXITP

End Function

Public Function Version_Upgrade_07(blnContinue As Boolean, blnConvert_TrustDta As Boolean, lngTrustDtaDbsID As Long, strKeyTbl As String, dblPB_ThisStep As Double, lngOldTbls As Long, arr_varOldTbl As Variant, lngAccts As Long, arr_varAcct As Variant, lngStats As Long, arr_varTmp03 As Variant, wrkLoc As DAO.Workspace, wrkLnk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database) As Integer
' ** This handles the new currency tables.
' ** Tables converted here:
' **   tblCurrency
' **   tblCurrency_History
' **   tblCurrency_Account
' **   tblLedgerHidden
' **
' ** Return values:
' **    0  OK
' **   -6  Index/Key
' **   -7  Can't Open
' **   -9  Error
' **
' ** Called by:
' **   modVersionConvertFuncs1:
' **     Version_Upgrade_01()

        ' ** Version_Upgrade_07(
        ' **   blnContinue As Boolean, blnConvert_TrustDta As Boolean, lngTrustDtaDbsID As Long,
        ' **   strKeyTbl As String, dblPB_ThisStep As Double, lngOldTbls As Long, arr_varOldTbl As Variant,
        ' **   lngAccts As Long, arr_varAcct As Variant, lngStats As Long, arr_varTmp03 As Variant,
        ' **   wrkLoc As DAO.Workspace, wrkLnk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database
        ' ** ) As Integer

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Upgrade_07"

        Dim rstLnk As DAO.Recordset, rstLoc1 As DAO.Recordset, rstLoc2 As DAO.Recordset
        Dim dbsLnkX As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field
        Dim strCurrTblName As String, strCurrKeyFldName As String
        Dim lngCurrTblID As Long, lngCurrKeyFldID As Long
        Dim arr_varStat() As Variant
        Dim strPath As String, strFile As String, strPathFile As String
        Dim lngRecs As Long
        Dim blnFound As Boolean, blnFound2 As Boolean
        Dim intPos01 As Integer
        Dim varTmp00 As Variant, strTmp01 As String, blnTmp02 As Boolean, lngTmp14 As Long
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim intRetVal As Integer

        ' ** Array: arr_varStat().
        Const STAT_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const STAT_ORD As Integer = 0
        Const STAT_NAM As Integer = 1
        Const STAT_CNT As Integer = 2
        Const STAT_DSC As Integer = 3

20710   If gblnDev_NoErrHandle = True Then
20720 On Error GoTo 0
20730   End If

20740   intRetVal = 0
20750   lngRecs = 0&

20760   If blnContinue = True Then  ' ** Is a conversion.

20770     If blnContinue = True Then  ' ** Conversion not already done.

20780       lngTmp14 = 0&
20790       ReDim arr_varStat(STAT_ELEMS, 0)

20800       For lngX = 0& To (lngStats - 1&)
20810         lngTmp14 = lngTmp14 + 1&
20820         lngE = lngTmp14 - 1&
20830         ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
20840         For lngZ = 0& To STAT_ELEMS
20850           arr_varStat(lngZ, lngE) = arr_varTmp03(lngZ, lngX)
20860         Next  ' ** lngZ.
20870       Next  ' ** lngX.

20880       If blnConvert_TrustDta = True Then

20890         If blnContinue = True Then  ' ** Workspace opens.

20900           With wrkLnk

20910             If blnContinue = True Then  ' ** Open dbsLnk.

20920               With dbsLnk

20930                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' *******************************************
                        ' ** Table: tblCurrency.
                        ' *******************************************

                        ' ** Step 22: tblCurrency.
20940                   dblPB_ThisStep = 22#
20950                   Version_Status 3, dblPB_ThisStep, "tblCurrency"  ' ** Module Function: modVersionConvertFuncs1.

20960                   strCurrTblName = "tblCurrency"
20970                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

20980                   blnFound = False: lngRecs = 0&
20990                   For lngX = 0& To (lngOldTbls - 1&)
21000                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
21010                       blnFound = True
21020                       Exit For
21030                     End If
21040                   Next

21050                   If blnFound = True Then
21060                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
21070                     With rstLnk
21080                       If .BOF = True And .EOF = True Then
                              ' ** If they've got the table, it should have records!
21090                       Else
21100                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
                              ' ** No earlier versions have this table.
                              ' ** Table: tblCurrency
                              ' **   ![curr_id]            AutoNumber
                              ' **   ![curr_code]
                              ' **   ![curr_name]
                              ' **   ![curr_word1]
                              ' **   ![curr_word2]
                              ' **   ![curr_rate1]
                              ' **   ![curr_rate2]
                              ' **   ![curr_date]
                              ' **   ![curr_iso]
                              ' **   ![curr_decimal]
                              ' **   ![curr_active]
                              ' **   ![curr_fund]
                              ' **   ![curr_bmu]
                              ' **   ![curr_metal]
                              ' **   ![curr_alt]
                              ' **   ![curr_notes]
                              ' **   ![curr_username]
                              ' **   ![curr_datemodified]
                              ' ** We're only interested in active, rates, date, and any notes they may have added.
21110                         .MoveLast
21120                         lngRecs = .RecordCount
21130                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
21140                         .MoveFirst
21150                         For lngX = 1& To lngRecs
                                ' ** Though they should match ID-for-ID, I don't want to depend on that,
                                ' ** plus, an earlier version may have fewer overall entries.
21160                           rstLoc1.FindFirst "[curr_id] = " & CStr(![curr_id])
21170                           If rstLoc1.NoMatch = False Then
21180                             If ![curr_active] <> rstLoc1![curr_active] Then
21190                               rstLoc1.Edit
21200                               rstLoc1![curr_active] = ![curr_active]
21210                               rstLoc1![curr_username] = ![curr_username]
21220                               rstLoc1![curr_datemodified] = ![curr_datemodified]
21230                               rstLoc1.Update
21240                             End If
21250                             If ![curr_rate1] > 0 Then
21260                               If ![curr_rate1] <> rstLoc1![curr_rate1] Then
21270                                 rstLoc1.Edit
21280                                 rstLoc1![curr_rate1] = ![curr_rate1]
21290                                 rstLoc1![curr_username] = ![curr_username]
21300                                 rstLoc1![curr_datemodified] = ![curr_datemodified]
21310                                 rstLoc1.Update
21320                               End If
21330                             End If
21340                             If ![curr_rate2] > 0 Then
21350                               If ![curr_rate2] <> rstLoc1![curr_rate2] Then
21360                                 rstLoc1.Edit
21370                                 rstLoc1![curr_rate2] = ![curr_rate2]
21380                                 rstLoc1![curr_username] = ![curr_username]
21390                                 rstLoc1![curr_datemodified] = ![curr_datemodified]
21400                                 rstLoc1.Update
21410                               End If
21420                             End If
21430                             If ![curr_date] <> rstLoc1![curr_date] Then
21440                               rstLoc1.Edit
21450                               rstLoc1![curr_date] = ![curr_date]
21460                               rstLoc1![curr_username] = ![curr_username]
21470                               rstLoc1![curr_datemodified] = ![curr_datemodified]
21480                               rstLoc1.Update
21490                             End If
21500                             If IsNull(![curr_notes]) = False Then
21510                               Select Case IsNull(rstLoc1![curr_notes])
                                    Case True
                                      ' ** Ours has no notes, so use theirs.
21520                                 rstLoc1.Edit
21530                                 rstLoc1![curr_notes] = ![curr_notes]
21540                                 rstLoc1![curr_username] = ![curr_username]
21550                                 rstLoc1![curr_datemodified] = ![curr_datemodified]
21560                                 rstLoc1.Update
21570                               Case False
                                      ' ** Ours has notes.
21580                                 If ![curr_notes] <> rstLoc1![curr_notes] Then
                                        ' ** Our notes may have additional info to theirs, as well as their own notes.
21590                                   intPos01 = InStr(rstLoc1![curr_notes], ![curr_notes])
21600                                   If intPos01 > 0 Then
                                          ' ** Their notes are wholly contained within ours, so we added info.
                                          ' ** No change.
21610                                   Else
21620                                     intPos01 = InStr(![curr_notes], rstLoc1![curr_notes])
21630                                     If intPos01 > 0 Then
                                            ' ** Our notes are wholly contained within theirs, so they added notes.
21640                                       rstLoc1.Edit
21650                                       rstLoc1![curr_notes] = ![curr_notes]
21660                                       rstLoc1![curr_username] = ![curr_username]
21670                                       rstLoc1![curr_datemodified] = ![curr_datemodified]
21680                                       rstLoc1.Update
21690                                     Else
                                            ' ** Just use theirs.
21700                                       rstLoc1.Edit
21710                                       rstLoc1![curr_notes] = ![curr_notes]
21720                                       rstLoc1![curr_username] = ![curr_username]
21730                                       rstLoc1![curr_datemodified] = ![curr_datemodified]
21740                                       rstLoc1.Update
21750                                     End If
21760                                   End If
21770                                 End If
21780                               End Select
21790                             End If
21800                           End If  ' ** NoMatch.
21810                           If lngX < lngRecs Then .MoveNext
21820                         Next  ' ** lngX.
21830                         rstLoc1.Close
21840                       End If  ' ** BOF, EOF.
21850                       .Close
21860                     End With  ' ** rstLnk.
21870                   End If  ' ** blnFound.

                        ' *******************************************
                        ' ** Table: tblCurrency_History.
                        ' *******************************************

                        'DON'T FORGET ALL THE NOTES ABOVE TO CHECK ONCE CURRENCY HISTORY IS CONVERTED!
                        ' ** Step 23: tblCurrency_History.
21880                   dblPB_ThisStep = 23#
21890                   Version_Status 3, dblPB_ThisStep, "tblCurrency_History"  ' ** Module Function: modVersionConvertFuncs1.

21900                   strCurrTblName = "tblCurrency_History"
21910                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

21920                   blnFound = False: lngRecs = 0&
21930                   For lngX = 0& To (lngOldTbls - 1&)
21940                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
21950                       blnFound = True
21960                       Exit For
21970                     End If
21980                   Next

21990                   If blnFound = True Then
22000                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
22010                     With rstLnk
22020                       If .BOF = True And .EOF = True Then
                              ' ** If they've got the table, it should have records!
22030                       Else
22040                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
                              ' ** No earlier versions have this table.
                              ' ** Table: tblCurrency_History
                              ' **   ![curr_id]
                              ' **   ![currhist_id]            AutoNumber
                              ' **   ![curr_date]
                              ' **   ![curr_rate1]
                              ' **   ![curr_rate2]
                              ' **   ![currhist_datemodified]
22050                         .MoveLast
22060                         lngRecs = .RecordCount
22070                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
22080                         .MoveFirst
22090                         For lngX = 1& To lngRecs
22100                           rstLoc1.FindFirst "[curr_id] = " & CStr(![curr_id]) & " And " & _
                                  "[curr_date] = #" & Format(![curr_date], "mm/dd/yyyy") & "#"
22110                           Select Case .NoMatch
                                Case True
                                  ' ** Add their record.
22120                             rstLoc1.AddNew
22130                             rstLoc1![curr_id] = ![curr_id]
                                  ' ** ![currhist_id] : AutoNumber.
22140                             rstLoc1![curr_date] = ![curr_date]
22150                             rstLoc1![curr_rate1] = ![curr_rate1]
22160                             rstLoc1![curr_rate2] = ![curr_rate2]
22170                             rstLoc1![currhist_datemodified] = ![currhist_datemodified]
22180                             rstLoc1.Update
22190                           Case False
22200                             If ![curr_rate1] <> rstLoc1![curr_rate1] Then
22210                               rstLoc1.Edit
22220                               rstLoc1![curr_rate1] = ![curr_rate1]
22230                               rstLoc1![currhist_datemodified] = ![currhist_datemodified]
22240                               rstLoc1.Update
22250                             End If
22260                             If ![curr_rate2] <> rstLoc1![curr_rate2] Then
22270                               rstLoc1.Edit
22280                               rstLoc1![curr_rate2] = ![curr_rate2]
22290                               rstLoc1![currhist_datemodified] = ![currhist_datemodified]
22300                               rstLoc1.Update
22310                             End If
22320                           End Select
22330                           If lngX < lngRecs Then .MoveNext
22340                         Next  ' ** lngX.
22350                         rstLoc1.Close
22360                       End If  ' ** BOF, EOF.
22370                       .Close
22380                     End With  ' ** rstLnk.
22390                   End If  ' ** blnFound.

                        ' *******************************************
                        ' ** Table: tblCurrency_Account.
                        ' *******************************************

                        ' ** Step 24: tblCurrency_Account.
22400                   dblPB_ThisStep = 24#
22410                   Version_Status 3, dblPB_ThisStep, "tblCurrency_Account"  ' ** Module Function: modVersionConvertFuncs1.

22420                   strCurrTblName = "tblCurrency_Account"
22430                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

22440                   blnFound = False: blnFound2 = False: lngRecs = 0&
22450                   For lngX = 0& To (lngOldTbls - 1&)
22460                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
22470                       blnFound = True
22480                       Exit For
22490                     End If
22500                   Next

22510                   If blnFound = True Then

22520                     blnFound2 = False: blnTmp02 = False
22530                     For Each tdf In .TableDefs
22540                       With tdf
22550                         If .Name = strCurrTblName Then
22560                           blnFound2 = True
22570                           Exit For
22580                         End If
22590                       End With
22600                     Next

22610                     If blnFound2 = False Then
                            ' ** I think it was in Trust.mde for a while.
                            ' ** Do I look for it?  Is it worth it for just Georgetown?
                            ' ** The old Trust.mde may be in \Convert_new, or it may be \Trust Accountant\Trust_bak.mde.
22620                       strPath = Parse_Path(.Name)  ' ** Module Function: modFileUtilities.
22630                       strFile = gstrFile_App & "." & Parse_Ext(CurrentAppName)  ' ** Module Function: modFileUtilities.
22640                       strPathFile = strPath & LNK_SEP & strFile
22650                       blnFound2 = FileExists(strPathFile)  ' ** Module Function: modFileUtilities.
22660                       If blnFound2 = False Then
22670                         strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
22680                         strFile = gstrFile_App & "_bak." & Parse_Ext(CurrentAppName)  ' ** Module Function: modFileUtilities.
22690                         strPathFile = strPath & LNK_SEP & strFile
22700                         blnFound2 = FileExists(strPathFile)  ' ** Module Function: modFileUtilities.
22710                       End If  ' ** blnFound2.
22720                       If blnFound2 = True Then
22730                         Set dbsLnkX = wrkLnk.OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
                              ' ** Not Exclusive, Read-Only.
22740                         blnFound2 = False
22750                         With dbsLnkX
22760                           For Each tdf In .TableDefs
22770                             With tdf
22780                               If .Name = strCurrTblName Then
22790                                 blnFound2 = True
22800                                 blnTmp02 = True
22810                                 Exit For
22820                               End If
22830                             End With
22840                           Next
22850                         End With  ' ** dbsLnkX.
22860                       Else
                              ' ** Skip it!
22870                       End If  ' ** blnFound2.
22880                     End If  ' ** blnFound2.

22890                     If blnFound2 = True Then

22900                       Select Case blnTmp02
                            Case True
22910                         Set rstLnk = dbsLnkX.OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
22920                       Case False
22930                         Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
22940                       End Select

22950                       With rstLnk
22960                         If .BOF = True And .EOF = True Then
                                ' ** Not used yet.
22970                         Else
22980                           strCurrKeyFldName = "accountno"
22990                           lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                  "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                  "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
23000                           Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
23010                           Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                                ' ** No earlier versions have this table.
                                ' ** Table: tblCurrency_Account
                                ' **   ![curracct_id]         AutoNumber
                                ' **   ![accountno]
                                ' **   ![curracct_jno]
                                ' **   ![curracct_aa]
                                ' **   ![curracct_suppress]
                                ' **   ![curracct_sort]
                                ' **   ![curracct_datemodified]
23020                           .MoveLast
23030                           lngRecs = .RecordCount
23040                           Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
23050                           .MoveFirst
23060                           For lngX = 1& To lngRecs
                                  ' ** Add the record to the new table.
23070                             strTmp01 = ![accountno]
23080                             blnFound = False
23090                             For lngY = 0& To (lngAccts - 1&)
23100                               If arr_varAcct(A_NUM, lngY) = strTmp01 Then
23110                                 blnFound = True
23120                                 Exit For
23130                               End If
23140                             Next
23150                             If blnFound = True Then
23160                               rstLoc1.AddNew
                                    ' ** ![curracct_id] : AutoNumber.
23170                               rstLoc1![accountno] = strTmp01
23180                               rstLoc1![curracct_jno] = ![curracct_jno]
23190                               rstLoc1![curracct_aa] = ![curracct_aa]
23200                               rstLoc1![curracct_suppress] = ![curracct_suppress]
23210                               rstLoc1![curracct_sort] = ![curracct_sort]
23220                               rstLoc1![curracct_datemodified] = ![curracct_datemodified]
23230 On Error Resume Next
23240                               rstLoc1.Update
23250                               If ERR.Number <> 0 Then
                                      ' ** ERR.Number = 3022
                                      ' ** The changes you requested to the table were not successful because they
                                      ' ** would create duplicate values in the index, primary key, or relationship.
23260                                 If strTmp01 = "INCOME O/U" Or strTmp01 = "SUSPENSE" Then
                                        ' ** Sometimes these hit, sometimes they don't.
23270                                   rstLoc1.Cancel
23280                                 Else
23290                                   lngStats = lngStats + 1&
23300                                   lngE = lngStats - 1&
23310                                   ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
23320                                   arr_varStat(STAT_ORD, lngE) = CInt(24)
23330                                   arr_varStat(STAT_NAM, lngE) = "Currency: "
23340                                   arr_varStat(STAT_CNT, lngE) = lngX
23350                                   arr_varStat(STAT_DSC, lngE) = "Account " & strTmp01 & " created a conflict in tblCurrency_Account."
23360                                   rstLoc1.Cancel
23370                                 End If
23380 On Error GoTo ERRH
23390                               Else
23400 On Error GoTo ERRH
23410                               End If
23420                             End If
23430                             strTmp01 = vbNullString
23440                             If lngX < lngRecs Then .MoveNext
23450                           Next
23460                           rstLoc1.Close
23470                           rstLoc2.Close
23480                         End If  ' ** Records present.
23490                         .Close
23500                       End With  ' ** rstLnk.

23510                     End If  ' ** blnFound2.
23520                     If blnTmp02 = True Then
23530                       dbsLnkX.Close
23540                       Set dbsLnkX = Nothing
23550                     End If  ' ** blnTmp02.

23560                   End If  ' ** blnFound.

                        ' *******************************************
                        ' ** Table: tblLedgerHidden.
                        ' *******************************************

                        ' ** Step 25: tblLedgerHidden.
23570                   dblPB_ThisStep = 25#
23580                   Version_Status 3, dblPB_ThisStep, "tblLedgerHidden"  ' ** Module Function: modVersionConvertFuncs1.

23590                   strCurrTblName = "tblLedgerHidden"
23600                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

23610                   blnFound = False: blnFound2 = False: lngRecs = 0&
23620                   For lngX = 0& To (lngOldTbls - 1&)
23630                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
23640                       blnFound = True
23650                       Exit For
23660                     End If
23670                   Next

23680                   If blnFound = True Then
23690                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
23700                     With rstLnk
23710                       If .BOF = True And .EOF = True Then
                              ' ** Not used yet.
23720                       Else
23730                         strCurrKeyFldName = "assetno"
23740                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
23750                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
23760                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** No earlier versions have this table.
                              ' ** Table: tblLedgerHidden
                              ' **   ![ledghid_id]          AutoNumber
                              ' **   ![journalno]
                              ' **   ![accountno]
                              ' **   ![assetno]
                              ' **   ![transdate]
                              ' **   ![ledghid_cnt]
                              ' **   ![ledghid_grpnum]
                              ' **   ![ledghid_ord]
                              ' **   ![ledghidtype_type]
                              ' **   ![ledghid_uniqueid]
                              ' **   ![ledghid_username]
                              ' **   ![ledghid_datemodified]
23770                         .MoveLast
23780                         lngRecs = .RecordCount
23790                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
23800                         .MoveFirst
23810                         For lngX = 1& To lngRecs
                                ' ** Add the record to the new table.
23820                           strTmp01 = ![accountno]
23830                           blnFound = False
23840                           For lngY = 0& To (lngAccts - 1&)
23850                             If arr_varAcct(A_NUM, lngY) = strTmp01 Then
23860                               blnFound = True
23870                               Exit For
23880                             End If
23890                           Next
23900                           If blnFound = False Then
23910                             varTmp00 = DLookup("[accountno]", "ledger", "[journalno] = " & CStr(![journalno]))
23920                             If IsNull(varTmp00) = False Then
23930                               blnFound = True
23940                               strTmp01 = varTmp00
23950                             Else
                                    ' ** I don't think it'll ever get here, but if it does, just ditch 'em.
23960                             End If
23970                           End If
23980                           If blnFound = True Then
23990                             rstLoc1.AddNew
                                  ' ** ![ledghid_id] : AutoNumber.
24000                             rstLoc1![journalno] = ![journalno]
24010                             rstLoc1![accountno] = strTmp01
24020                             If IsNull(![assetno]) = False Then
24030                               If ![assetno] > 0& Then
24040                                 rstLoc2.MoveFirst
24050                                 rstLoc2.FindFirst "[tbl_name] = 'masterasset' And [fld_name] = 'assetno' And " & _
                                        "[key_lng_id1] = " & CStr(![assetno])
24060                                 If rstLoc2.NoMatch = True And ![assetno] = 1& Then
                                        ' ** It's the 'Accrued Interest Asset', masterasset_TYPE = 'IA',
                                        ' ** which doesn't get moved.
24070                                   rstLoc1![assetno] = 1&
24080                                 ElseIf rstLoc2.NoMatch = False Then
24090                                   rstLoc1![assetno] = rstLoc2![key_lng_id2] 'Req  0
24100                                 Else
                                        ' ** It may be an orphan, but it'll have to remain one!
24110                                   rstLoc1![assetno] = ![assetno]
24120                                 End If
24130                               Else
24140                                 rstLoc1![assetno] = 0&
24150                               End If
24160                             Else
24170                               rstLoc1![assetno] = 0&
24180                             End If
24190                             rstLoc1![transdate] = ![transdate]
24200                             rstLoc1![ledghid_cnt] = ![ledghid_cnt]
24210                             rstLoc1![ledghid_grpnum] = ![ledghid_grpnum]
24220                             rstLoc1![ledghid_ord] = ![ledghid_ord]
24230                             rstLoc1![ledghidtype_type] = ![ledghidtype_type]
24240                             rstLoc1![ledghid_uniqueid] = ![ledghid_uniqueid]
24250                             rstLoc1![ledghid_username] = ![ledghid_username]
24260                             rstLoc1![ledghid_datemodified] = ![ledghid_datemodified]
24270                             rstLoc1.Update
24280                           End If
24290                           strTmp01 = vbNullString
24300                           If lngX < lngRecs Then .MoveNext
24310                         Next
24320                         rstLoc1.Close
24330                         rstLoc2.Close
24340                       End If  ' ** Records present.
24350                       .Close
24360                     End With  ' ** rstLnk.
24370                   End If  ' ** blnFound.

24380                 End If  ' ** dbsLoc is still open.

                      '.Close  'TO VERSION_UPGRADE_08!
24390               End With  ' ** TrustDta.mdb: dbsLnk.

24400             End If  ' ** dbsLnk opens.

                  '.Close  'TO VERSION_UPGRADE_08!
24410           End With  ' ** wrkLnk.

24420         End If  ' ** Workspace opens: blnContinue.

24430       End If  ' ** blnConvert_TrustDta.

24440       If blnContinue = False Then
24450         dbsLoc.Close
24460         wrkLoc.Close
24470       End If

24480       If lngTmp14 > lngStats Then
24490         lngStats = lngTmp14
24500         arr_varTmp03 = arr_varStat
24510       End If

24520     End If  ' ** Conversion not already done.

24530   End If  ' ** Is a conversion.

EXITP:
24540   Set fld = Nothing
24550   Set tdf = Nothing
24560   Set dbsLnkX = Nothing
24570   Set rstLnk = Nothing
24580   Set rstLoc1 = Nothing
24590   Set rstLoc2 = Nothing
24600   Version_Upgrade_07 = intRetVal
24610   Exit Function

ERRH:
24620   intRetVal = -9
24630   DoCmd.Hourglass False
24640   lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
24650   ErrInfo_Set lngErrNum, lngErrLine, strErrDesc  ' ** Module Procedure: modVersionConvertFuncs1.
24660   Select Case ERR.Number
        Case Else
24670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24680   End Select
24690   Resume EXITP

End Function

Public Function Version_DataXFer(strType As String, strOpt As String, Optional arr_varTmp00 As Variant) As Variant
' ** This handles some data transfer between these functions and the conversion form.
' ** The Set must just precede the Get.
' **
' ** Called by:
' **   modVersionConvertFuncs1:
' **     Version_Upgrade_03()
' **     Version_Status()
' **   frmVersion_Input:
' **     Form_Load()
' **     cmdCancel_Click()
' **     cmdNext_Click()
' **   frmVersion_Main:
' **     Form_Open()

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_DataXFer"

        Dim arr_varRetVal As Variant

24710   If gblnDev_NoErrHandle = True Then
24720 On Error GoTo 0
24730   End If

24740   Select Case strType
        Case "Set"
24750     Select Case strOpt
          Case "PathFile"
            ' ** These are sent here via Version_DataXFer_PathVars(), below: strPathFile_Data, strPathFile_Archive, strOldVersion, strReleaseDate.
24760       ReDim arr_varDataXFer(3, 0)  ' ** 1 record of 4 fields.
24770       arr_varDataXFer(0, 0) = strPathFile_Data  ' ** arr_varOldFile(F_PTHFIL, lngDtaElem)
24780       arr_varDataXFer(1, 0) = strPathFile_Archive  ' ** arr_varOldFile(F_PTHFIL, lngArchElem)
24790       arr_varDataXFer(2, 0) = strOldVersion
24800       arr_varDataXFer(3, 0) = strReleaseDate  ' ** arr_varOldFile(F_APPDATE, lngDtaElem)
24810       arr_varRetVal = CBool(True)
24820     Case "CoInfo"
24830       ReDim arr_varDataXFer(16, 0)  ' ** 1 record of 17 fields.
24840       arr_varDataXFer(0, 0) = strTmp_Name
24850       arr_varDataXFer(1, 0) = strTmp_Address1
24860       arr_varDataXFer(2, 0) = strTmp_Address2
24870       arr_varDataXFer(3, 0) = strTmp_City
24880       arr_varDataXFer(4, 0) = strTmp_State
24890       arr_varDataXFer(5, 0) = strTmp_Zip
24900       arr_varDataXFer(6, 0) = strTmp_Phone
24910       arr_varDataXFer(7, 0) = blnTmp_IncomeTaxCoding
24920       arr_varDataXFer(8, 0) = blnTmp_RevenueExpenseTracking
24930       arr_varDataXFer(9, 0) = blnTmp_AccountNoWithType
24940       arr_varDataXFer(10, 0) = blnTmp_SeparateCheckingAccounts
24950       arr_varDataXFer(11, 0) = blnTmp_TabCopyAccount
24960       arr_varDataXFer(12, 0) = blnTmp_LinkRevTaxCodes
24970       arr_varDataXFer(13, 0) = blnTmp_SpecialCapGainLoss
24980       arr_varDataXFer(14, 0) = intTmp_SpecialCapGainLossOpt
24990       arr_varDataXFer(15, 0) = strTmp_Country
25000       arr_varDataXFer(16, 0) = strTmp_PostalCode
            'not saved: ![CoInfo_ID]
            'strTmp04 = ![CoInfo_Name]
            'strTmp05 = ![CoInfo_Address1]
            'strTmp06 = ![CoInfo_Address2]
            'strTmp07 = ![CoInfo_City]
            'strTmp08 = ![CoInfo_State]
            'strTmp09 = ![CoInfo_Zip]
            'strTmp11 = ![CoInfo_Country]
            'strTmp12 = ![CoInfo_PostalCode]
            'strTmp10 = ![CoInfo_Phone]
            'blnTmp22 = ![IncomeTaxCoding]
            'blnTmp23 = ![RevenueExpenseTracking]
            'blnTmp24 = ![AccountNoWithType]
            'blnTmp25 = ![SeparateCheckingAccounts]
            'blnTmp26 = ![TabCopyAccount]
            'blnTmp27 = ![LinkRevTaxCodes]
            'varTmp00 = ![SpecialCapGainLoss]
            'lngTmp13 = ![SpecialCapGainLossOpt]
            'not saved: ![Username]
            'not saved: ![CoInfo_DateModified]
25010       arr_varRetVal = CBool(True)
25020     Case Else
25030       ReDim arr_varDataXFer(0, 0)
25040       arr_varDataXFer(0, 0) = "#EMPTY"
25050     End Select
25060   Case "Get"
25070     Select Case strOpt
          Case "PathFile"
25080       arr_varRetVal = arr_varDataXFer
25090     Case "CoInfo"
25100       arr_varRetVal = arr_varDataXFer
25110     End Select
25120   Case "Ret"
25130     Select Case strOpt
          Case "CoInfo"
25140       If IsMissing(arr_varTmp00) = False Then
25150         arr_varDataXFer(0, 0) = arr_varTmp00(0, 0)
25160         If Left(arr_varTmp00(0, 0), 1) <> "#" Then
                ' ** CoInfo_ID, Username, and CoInfo_DateModified will use current values.
25170           arr_varDataXFer(1, 0) = arr_varTmp00(1, 0)
25180           arr_varDataXFer(2, 0) = arr_varTmp00(2, 0)
25190           arr_varDataXFer(3, 0) = arr_varTmp00(3, 0)
25200           arr_varDataXFer(4, 0) = arr_varTmp00(4, 0)
25210           arr_varDataXFer(5, 0) = arr_varTmp00(5, 0)
25220           arr_varDataXFer(6, 0) = arr_varTmp00(6, 0)
25230           arr_varDataXFer(7, 0) = arr_varTmp00(7, 0)
25240           arr_varDataXFer(8, 0) = arr_varTmp00(8, 0)
25250           arr_varDataXFer(9, 0) = arr_varTmp00(9, 0)
25260           arr_varDataXFer(10, 0) = arr_varTmp00(10, 0)
25270           arr_varDataXFer(11, 0) = arr_varTmp00(11, 0)
25280           arr_varDataXFer(12, 0) = arr_varTmp00(12, 0)
25290           arr_varDataXFer(13, 0) = arr_varTmp00(13, 0)
25300           arr_varDataXFer(14, 0) = arr_varTmp00(14, 0)
25310           arr_varDataXFer(15, 0) = arr_varTmp00(15, 0)
25320           arr_varDataXFer(16, 0) = arr_varTmp00(16, 0)
25330         End If
25340         arr_varRetVal = CBool(True)
25350       Else
25360         arr_varRetVal = CBool(False)
25370       End If
25380     End Select
25390   End Select

EXITP:
25400   Version_DataXFer = arr_varRetVal
25410   Exit Function

ERRH:
25420   Select Case strType
        Case "Set", "Ret"
25430     arr_varRetVal = CBool(False)
25440   Case "Get"
25450     arr_varRetVal(0, 0) = RET_ERR
25460   End Select
25470   lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
25480   ErrInfo_Set lngErrNum, lngErrLine, strErrDesc  ' ** Module Procedure: modVersionConvertFuncs1.
25490   Select Case ERR.Number
        Case Else
25500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25510   End Select
25520   Resume EXITP

End Function

Public Sub Version_DataXFer_PathVars(strTmp04 As String, strTmp05 As String, strTmp06 As String, strTmp07 As String)
' **
' ** Called by:
' **   modVersionConvertFuncs1:
' **     Version_Status()

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_DataXFer_PathVars"

25610   strPathFile_Data = strTmp04
25620   strPathFile_Archive = strTmp05
25630   If strTmp06 <> vbNullString Then
25640     strOldVersion = strTmp06
25650   End If
25660   strReleaseDate = strTmp07

        'ConversionCheck
        '  modVersionConvertFuncs2
        'Version_Upgrade_01
        '  modVersionConvertFuncs1
        'Version_Upgrade_02
        '  modVersionConvertFuncs1
        'Version_Status
        '  modVersionConvertFuncs1
        'Version_DataXFer
        '  modVersionConvertFuncs2
        '  THIS CALL EXPECTS strTmp04, strTmp05, strOldVersion FROM Version_Status,
        '  WHICH CAME FROM arr_varOldFile()!

EXITP:
25670   Exit Sub

ERRH:
25680   Select Case ERR.Number
        Case Else
25690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25700   End Select
25710   Resume EXITP

End Sub

Public Sub Version_DataXFer_CoInfoVars(strTmp04 As String, strTmp05 As String, strTmp06 As String, strTmp07 As String, strTmp08 As String, strTmp09 As String, strTmp10 As String, strTmp11 As String, strTmp12 As String, lngTmp13 As Long, blnTmp22 As Boolean, blnTmp23 As Boolean, blnTmp24 As Boolean, blnTmp25 As Boolean, blnTmp26 As Boolean, blnTmp27 As Boolean, blnTmp28 As Boolean)
' **
' ** Called by:
' **   modVersionConvertFuncs1:
' **     Version_Upgrade_03()

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_DataXFer_CoInfoVars"

25810   strTmp_Name = strTmp04
25820   strTmp_Address1 = strTmp05
25830   strTmp_Address2 = strTmp06
25840   strTmp_City = strTmp07
25850   strTmp_State = strTmp08
25860   strTmp_Zip = strTmp09
25870   strTmp_Country = strTmp11
25880   strTmp_PostalCode = strTmp12
25890   strTmp_Phone = strTmp10
25900   blnTmp_IncomeTaxCoding = blnTmp22
25910   blnTmp_RevenueExpenseTracking = blnTmp23
25920   blnTmp_AccountNoWithType = blnTmp24
25930   blnTmp_SeparateCheckingAccounts = blnTmp25
25940   blnTmp_TabCopyAccount = blnTmp26
25950   blnTmp_LinkRevTaxCodes = blnTmp27
25960   blnTmp_SpecialCapGainLoss = blnTmp28
25970   intTmp_SpecialCapGainLossOpt = lngTmp13

        'not saved: ![CoInfo_ID]
        'strTmp04 = ![CoInfo_Name]                 'strTmp_Name
        'strTmp05 = ![CoInfo_Address1]             'strTmp_Address1
        'strTmp06 = ![CoInfo_Address2]             'strTmp_Address2
        'strTmp07 = ![CoInfo_City]                 'strTmp_City
        'strTmp08 = ![CoInfo_State]                'strTmp_State
        'strTmp09 = ![CoInfo_Zip]                  'strTmp_Zip
        'strTmp11 = ![CoInfo_Country]              'strTmp_Country
        'strTmp12 = ![CoInfo_PostalCode]           'strTmp_PostalCode
        'strTmp10 = ![CoInfo_Phone]                'strTmp_Phone
        'blnTmp22 = ![IncomeTaxCoding]             'blnTmp_IncomeTaxCoding
        'blnTmp23 = ![RevenueExpenseTracking]      'blnTmp_RevenueExpenseTracking
        'blnTmp24 = ![AccountNoWithType]           'blnTmp_AccountNoWithType
        'blnTmp25 = ![SeparateCheckingAccounts]    'blnTmp_SeparateCheckingAccounts
        'blnTmp26 = ![TabCopyAccount]              'blnTmp_TabCopyAccount
        'blnTmp27 = ![LinkRevTaxCodes]             'blnTmp_LinkRevTaxCodes
        'varTmp00 = ![SpecialCapGainLoss]          'blnTmp_SpecialCapGainLoss
        'lngTmp13 = ![SpecialCapGainLossOpt]       'intTmp_SpecialCapGainLossOpt
        'not saved: ![Username]
        'not saved: ![CoInfo_DateModified]

EXITP:
25980   Exit Sub

ERRH:
25990   Select Case ERR.Number
        Case Else
26000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26010   End Select
26020   Resume EXITP

End Sub

Public Function Version_IsEmpty(strTmp04 As String) As Boolean
' ** This checks the current tables to make sure they're empty before starting a conversion!
' ** If an error occurred, get empty ones from \Convert_Empty.
' ** DON'T CHECK IF THIS IS A DEMO!
' **
' ** Called by:
' **   modVersionConvertFuncs1:
' **     Version_Upgrade_02()

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_IsEmpty"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

26110   blnRetVal = True

26120   Set dbs = CurrentDb
26130   With dbs
          ' ** Account, other than 'INCOME O/U', 'SUSPENSE'.
26140     Set qdf = .QueryDefs("qryVersion_Convert_20")
26150     Set rst = qdf.OpenRecordset
26160     With rst
26170       If .BOF = True And .EOF = True Then
              ' ** Empty!
26180       Else
26190         strTmp04 = strTmp04 & "Account not empty;"
26200         blnRetVal = False
26210       End If
26220       .Close
26230     End With
          ' ** MasterAsset, other than 'Accrued Interest Asset', 'IA'.
26240     Set qdf = .QueryDefs("qryVersion_Convert_21")
26250     Set rst = qdf.OpenRecordset
26260     With rst
26270       If .BOF = True And .EOF = True Then
              ' ** Empty!
26280       Else
26290         strTmp04 = strTmp04 & "MasterAsset not empty;"
26300         blnRetVal = False
26310       End If
26320       .Close
26330     End With
          ' ** m_REVCODE, other revcode_ID's 1-6 (leave criteria at 4).
26340     Set qdf = .QueryDefs("qryVersion_Convert_22")
26350     Set rst = qdf.OpenRecordset
26360     With rst
26370       If .BOF = True And .EOF = True Then
              ' ** Empty!
26380       Else
26390         strTmp04 = strTmp04 & "m_REVCODE not empty;"
26400         blnRetVal = False
26410       End If
26420       .Close
26430     End With
          ' ** RecurringItems, other than RecurringItem_ID's 1-2.
26440     Set qdf = .QueryDefs("qryVersion_Convert_23")
26450     Set rst = qdf.OpenRecordset
26460     With rst
26470       If .BOF = True And .EOF = True Then
              ' ** Empty!
26480       Else
26490         strTmp04 = strTmp04 & "RecurringItems not empty;"
26500         blnRetVal = False
26510       End If
26520       .Close
26530     End With
          ' ** Location, other than '{Unassigned}'.
26540     Set qdf = .QueryDefs("qryVersion_Convert_24")
26550     Set rst = qdf.OpenRecordset
26560     With rst
26570       If .BOF = True And .EOF = True Then
              ' ** Empty!
26580       Else
26590         strTmp04 = strTmp04 & "Location not empty;"
26600         blnRetVal = False
26610       End If
26620       .Close
26630     End With
          ' ** AdminOfficer, other than '{Unassigned}'.
26640     Set qdf = .QueryDefs("qryVersion_Convert_25")
26650     Set rst = qdf.OpenRecordset
26660     With rst
26670       If .BOF = True And .EOF = True Then
              ' ** Empty!
26680       Else
26690         strTmp04 = strTmp04 & "AdminOfficer not empty;"
26700         blnRetVal = False
26710       End If
26720       .Close
26730     End With
          ' ** Users, other than our default users.
          ' ** Whidh are: Admin, Creator, Engine, Superuser, TAAdmin.
26740     Set qdf = .QueryDefs("qryVersion_Convert_26")
26750     Set rst = qdf.OpenRecordset
26760     With rst
26770       If .BOF = True And .EOF = True Then
              ' ** Empty!
26780       Else
26790         strTmp04 = strTmp04 & "Users not empty;"
26800         blnRetVal = False
26810       End If
26820       .Close
26830     End With
26840     Set rst = .OpenRecordset("ActiveAssets", dbOpenDynaset, dbReadOnly)
26850     With rst
26860       If .BOF = True And .EOF = True Then
              ' ** Empty!
26870       Else
26880         strTmp04 = strTmp04 & "ActiveAssets not empty;"
26890         blnRetVal = False
26900       End If
26910       .Close
26920     End With
26930     Set rst = .OpenRecordset("Balance", dbOpenDynaset, dbReadOnly)
26940     With rst
26950       If .BOF = True And .EOF = True Then
              ' ** Empty!
26960       Else
              ' ** If it's just 'INCOME O/U' and 'SUSPENSE', that's OK!
26970         .MoveLast
26980         lngTmp01 = .RecordCount
26990         .MoveFirst
27000         If lngTmp01 = 2& Then
27010           lngTmp02 = 0&
27020           For lngX = 1& To 2&
27030             Select Case ![accountno]
                  Case "INCOME O/U"
27040               lngTmp02 = lngTmp02 + 1&
27050             Case "SUSPENSE"
27060               lngTmp02 = lngTmp02 + 1&
27070             End Select
27080           Next
27090           If lngTmp02 = 2& Then
                  ' ** OK!
27100           Else
27110             strTmp04 = strTmp04 & "Balance not empty;"
27120             blnRetVal = False
27130           End If
27140         Else
27150           strTmp04 = strTmp04 & "Balance not empty;"
27160           blnRetVal = False
27170         End If
27180         lngTmp01 = 0&: lngTmp02 = 0&
27190       End If
27200       .Close
27210     End With
27220     Set rst = .OpenRecordset("Journal", dbOpenDynaset, dbReadOnly)
27230     With rst
27240       If .BOF = True And .EOF = True Then
              ' ** Empty!
27250       Else
27260         strTmp04 = strTmp04 & "Journal not empty;"
27270         blnRetVal = False
27280       End If
27290       .Close
27300     End With
27310     Set rst = .OpenRecordset("Ledger", dbOpenDynaset, dbReadOnly)
27320     With rst
27330       If .BOF = True And .EOF = True Then
              ' ** Empty!
27340       Else
27350         strTmp04 = strTmp04 & "Ledger not empty;"
27360         blnRetVal = False
27370       End If
27380       .Close
27390     End With
27400     Set rst = .OpenRecordset("LedgerHidden", dbOpenDynaset, dbReadOnly)
27410     With rst
27420       If .BOF = True And .EOF = True Then
              ' ** Empty!
27430       Else
27440         strTmp04 = strTmp04 & "LedgerHidden not empty;"
27450         blnRetVal = False
27460       End If
27470       .Close
27480     End With
27490     Set rst = .OpenRecordset("Schedule", dbOpenDynaset, dbReadOnly)
27500     With rst
27510       If .BOF = True And .EOF = True Then
              ' ** Empty!
27520       Else
27530         strTmp04 = strTmp04 & "Schedule not empty;"
27540         blnRetVal = False
27550       End If
27560       .Close
27570     End With
27580     .Close
27590   End With

EXITP:
27600   Set rst = Nothing
27610   Set qdf = Nothing
27620   Set dbs = Nothing
27630   Version_IsEmpty = blnRetVal
27640   Exit Function

ERRH:
27650   blnRetVal = False
27660   Select Case ERR.Number
        Case Else
27670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27680   End Select
27690   Resume EXITP

End Function

Public Function Version_ArchCheck(blnContinue As Boolean, blnConvert_TrstArch As Boolean, lngArchiveRecs As Long, intWrkType As Integer, lngArchElem As Long, arr_varOldFile As Variant) As Integer
' **
' ** Return values:
' **    0  OK
' **   -7  Can't Open {TrstArch.mdb}
' **   -9  Error
' **
' ** Called by:
' **   modVersionConvertFuncs1:
' **     Version_Upgrade_02()

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_ArchCheck"

        Dim wrkLnk As DAO.Workspace, dbsLnk As DAO.Database, rstLnk As DAO.Recordset, tdf As DAO.TableDef, fld As DAO.Field
        Dim strCurrTblName As String
        Dim blnFound As Boolean
        Dim lngTmp01 As Long, blnTmp02 As Boolean, blnTmp03 As Boolean
        Dim intRetVal As Integer

27710   intRetVal = 0

        ' ** Open the workspace with type found in Version_GetOldVer(), below.
27720   Select Case intWrkType
        Case 1
27730     Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
27740   Case 2
27750     Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
27760   Case 3
27770     Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
27780   Case 4
27790     Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
27800   Case 5
27810     Set wrkLnk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
27820   Case 6
27830     Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
27840   Case 7
27850     Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
27860   End Select

27870   If blnContinue = True Then  ' ** Workspace opens.

27880     With wrkLnk

27890 On Error Resume Next
27900       Set dbsLnk = .OpenDatabase(arr_varOldFile(F_PTHFIL, lngArchElem), False, True)  ' ** {pathfile}, {exclusive}, {read-only}
            ' ** Not Exclusive, Read-Only.
27910       If ERR.Number <> 0 Then
              ' ** Error opening old database.
27920         If gblnDev_NoErrHandle = True Then
27930 On Error GoTo 0
27940         Else
27950 On Error GoTo ERRH
27960         End If
27970         intRetVal = -7
27980         blnContinue = False
27990       Else
28000         If gblnDev_NoErrHandle = True Then
28010 On Error GoTo 0
28020         Else
28030 On Error GoTo ERRH
28040         End If

28050         lngTmp01 = 0&: blnTmp02 = False: blnTmp03 = False
28060         With dbsLnk
28070           blnFound = False
28080           For Each tdf In .TableDefs
28090             If tdf.Name = "ledger" Then
28100               blnFound = True
28110               Exit For
28120             End If
28130           Next
28140           If blnFound = True Then
28150             strCurrTblName = "ledger"
28160             Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
28170             With rstLnk
28180               lngTmp01 = .Fields.Count
28190               If lngTmp01 > 0& Then
28200                 If .BOF = True And .EOF = True Then
                        ' ** Not used yet.
28210                   blnConvert_TrstArch = False
28220                   lngArchiveRecs = 0&
28230                 Else
28240                   .MoveLast
28250                   lngArchiveRecs = .RecordCount
                        ' ** If there's only 1 or 2, check to see if they're TADemo records (found in at least 1 customer's data).
28260                   If lngArchiveRecs <= 2 Then
28270                     blnFound = False
28280                     For Each fld In .Fields
28290                       If fld.Name = "journal_USER" Then
28300                         blnFound = True
28310                         Exit For
28320                       End If
28330                     Next
28340                     If blnFound = True Then
28350                       .MoveFirst
28360                       If IsNull(![journal_USER]) = False Then
28370                         If ![journal_USER] = "TADemo" Then
28380                           blnTmp02 = True
28390                         End If
28400                       End If
28410                       If lngArchiveRecs > 1& Then
28420                         .MoveNext
28430                         If IsNull(![journal_USER]) = False Then
28440                           If ![journal_USER] = "TADemo" Then
28450                             blnTmp03 = True
28460                           End If
28470                         End If
28480                       End If
28490                     End If
28500                     If (lngArchiveRecs = 1& And blnTmp02 = True) Or (lngArchiveRecs = 2& And blnTmp02 = True And blnTmp03 = True) Then
28510                       blnConvert_TrstArch = False
28520                       lngArchiveRecs = 0&
28530                     End If
28540                   End If
28550                 End If  ' ** Has records.
28560               End If  ' ** lngTmp01 (field count).
28570               .Close
28580             End With  ' ** rstLnk.
28590           Else
28600             blnConvert_TrstArch = False
28610             lngArchiveRecs = 0&
28620           End If  ' ** blnFound.
28630           .Close
28640         End With  ' ** dbsLnk.
28650         lngTmp01 = 0&: blnTmp02 = False: blnTmp03 = False

28660       End If  ' ** Database opens.

28670       .Close
28680     End With  ' ** wrkLnk.

28690   End If  ' ** Workspace opens.

EXITP:
28700   Set wrkLnk = Nothing
28710   Set dbsLnk = Nothing
28720   Set rstLnk = Nothing
28730   Set tdf = Nothing
28740   Set fld = Nothing
28750   Version_ArchCheck = intRetVal
28760   Exit Function

ERRH:
28770   intRetVal = -9
28780   blnConvert_TrstArch = False
28790   DoCmd.Hourglass False
28800   lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
28810   ErrInfo_Set lngErrNum, lngErrLine, strErrDesc  ' ** Module Procedure: modVersionConvertFuncs1.
28820   Select Case ERR.Number
        Case Else
28830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28840   End Select
28850   Resume EXITP

End Function

Public Function Version_GetOldVer(blnContinue As Boolean, blnConvert_TrustDta As Boolean, blnConvert_TrstArch As Boolean, intWrkType As Integer, strOldVersion As String, lngVerCnvID As Long, lngDtaElem As Long, arr_varOldFile As Variant) As Integer
' ** This determines the version being converted.
' ** Return values:
' **    0  OK
' **   -1  Can't Connect
' **   -2  Can't Open
' **   -9  Error
' **
' ** Called by:
' **   modVersionConvertFuncs1:
' **     Version_Upgrade_02()

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_GetOldVer"

        Dim wrkLnk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database, qdf As DAO.QueryDef, rstLnk As DAO.Recordset
        Dim tdf As DAO.TableDef, doc As DAO.Document, prp As DAO.Property
        'Dim cnxn As ADODB.Connection, rsx1 As ADODB.Recordset  ' ** Early binding.
        Dim cnxn As Object, rsx1 As Object                      ' ** Late binding.
        Dim strCnxn As String
        Dim lngRecs As Long
        Dim blnFound As Boolean
        Dim intPos01 As Integer
        Dim strTmp04 As String, strTmp05 As String, strTmp06 As String, strTmp07 As String
        Dim intRetVal As Integer

28910   If gblnDev_NoErrHandle = True Then
28920 On Error GoTo 0
28930   End If

28940   intRetVal = 0
28950   lngRecs = 0&

28960   If blnContinue = True Then  ' ** Is a conversion.

28970     If blnContinue = True And blnConvert_TrustDta = True Or blnConvert_TrstArch = True Then

28980       If blnConvert_TrustDta = True Then

28990         intWrkType = 0
29000 On Error Resume Next
29010         Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
29020         If ERR.Number <> 0 Then
29030           Select Case gblnDev_NoErrHandle
                Case True
29040 On Error GoTo 0
29050           Case False
29060 On Error GoTo ERRH
29070           End Select
29080 On Error Resume Next
29090           Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
29100           If ERR.Number <> 0 Then
29110             Select Case gblnDev_NoErrHandle
                  Case True
29120 On Error GoTo 0
29130             Case False
29140 On Error GoTo ERRH
29150             End Select
29160 On Error Resume Next
29170             Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
29180             If ERR.Number <> 0 Then
29190               Select Case gblnDev_NoErrHandle
                    Case True
29200 On Error GoTo 0
29210               Case False
29220 On Error GoTo ERRH
29230               End Select
29240 On Error Resume Next
29250               Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
29260               If ERR.Number <> 0 Then
29270                 Select Case gblnDev_NoErrHandle
                      Case True
29280 On Error GoTo 0
29290                 Case False
29300 On Error GoTo ERRH
29310                 End Select
29320 On Error Resume Next
29330                 Set wrkLnk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
29340                 If ERR.Number <> 0 Then
29350                   Select Case gblnDev_NoErrHandle
                        Case True
29360 On Error GoTo 0
29370                   Case False
29380 On Error GoTo ERRH
29390                   End Select
29400 On Error Resume Next
29410                   Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
29420                   If ERR.Number <> 0 Then
29430                     Select Case gblnDev_NoErrHandle
                          Case True
29440 On Error GoTo 0
29450                     Case False
29460 On Error GoTo ERRH
29470                     End Select
29480 On Error Resume Next
29490                     Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
29500                     If ERR.Number <> 0 Then
29510                       Select Case gblnDev_NoErrHandle
                            Case True
29520 On Error GoTo 0
29530                       Case False
29540 On Error GoTo ERRH
29550                       End Select
29560                       intRetVal = -1
29570                       blnContinue = False
29580                     Else
29590                       Select Case gblnDev_NoErrHandle
                            Case True
29600 On Error GoTo 0
29610                       Case False
29620 On Error GoTo ERRH
29630                       End Select
29640                       intWrkType = 7
29650                     End If
29660                   Else
29670                     Select Case gblnDev_NoErrHandle
                          Case True
29680 On Error GoTo 0
29690                     Case False
29700 On Error GoTo ERRH
29710                     End Select
29720                     intWrkType = 6
29730                   End If
29740                 Else
29750                   Select Case gblnDev_NoErrHandle
                        Case True
29760 On Error GoTo 0
29770                   Case False
29780 On Error GoTo ERRH
29790                   End Select
29800                   intWrkType = 5
29810                 End If
29820               Else
29830                 Select Case gblnDev_NoErrHandle
                      Case True
29840 On Error GoTo 0
29850                 Case False
29860 On Error GoTo ERRH
29870                 End Select
29880                 intWrkType = 4
29890               End If
29900             Else
29910               Select Case gblnDev_NoErrHandle
                    Case True
29920 On Error GoTo 0
29930               Case False
29940 On Error GoTo ERRH
29950               End Select
29960               intWrkType = 3
29970             End If
29980           Else
29990             Select Case gblnDev_NoErrHandle
                  Case True
30000 On Error GoTo 0
30010             Case False
30020 On Error GoTo ERRH
30030             End Select
30040             intWrkType = 2
30050           End If
30060         Else
30070           Select Case gblnDev_NoErrHandle
                Case True
30080 On Error GoTo 0
30090           Case False
30100 On Error GoTo ERRH
30110           End Select
30120           intWrkType = 1
30130         End If

30140         If blnContinue = True Then  ' ** Workspace opens.

30150           With wrkLnk

30160 On Error Resume Next
30170             Set dbsLnk = .OpenDatabase(arr_varOldFile(F_PTHFIL, lngDtaElem), False, True)  ' ** {pathfile}, {exclusive}, {read-only}
                  ' ** Not Exclusive, Read-Only.
30180             If ERR.Number <> 0 Then
                    ' ** Error opening old database.
30190               If gblnDev_NoErrHandle = True Then
30200 On Error GoTo 0
30210               Else
30220 On Error GoTo ERRH
30230               End If
30240               intRetVal = -2
30250               blnContinue = False
30260             Else
30270               If gblnDev_NoErrHandle = True Then
30280 On Error GoTo 0
30290               Else
30300 On Error GoTo ERRH
30310               End If
30320               With dbsLnk

30330                 arr_varOldFile(F_ACC_VER, lngDtaElem) = .Containers("Databases").Documents("MSysDb").Properties("AccessVersion")
                      ' ** CurrentDb.Containers("Databases").Documents("MSysDb").Properties("AccessVersion") = 08.50

30340                 Select Case arr_varOldFile(F_ACC_VER, lngDtaElem)
                      Case "02.00"
30350                   arr_varOldFile(F_ACC_VER, lngDtaElem) = "Access 2.0"
30360                 Case "06.68"
30370                   arr_varOldFile(F_ACC_VER, lngDtaElem) = "Access 95"
30380                 Case "07.53"
30390                   arr_varOldFile(F_ACC_VER, lngDtaElem) = "Access 97"
30400                 Case "08.50"
30410                   arr_varOldFile(F_ACC_VER, lngDtaElem) = "Access 2000"
30420                 Case "09.50"
30430                   arr_varOldFile(F_ACC_VER, lngDtaElem) = "Access 2002/2003"
                        '###############################
                        'THIS NEEDS TO BE UPDATED WITH
'2007, 2010, 2013, 2016!
                        '###############################
30440                 Case Else
                        ' ** The Jet MDW.
30450                 End Select

30460                 For Each doc In .Containers("Databases").Documents
30470                   With doc
30480                     If .Name = "UserDefined" Then
30490                       For Each prp In .Properties
30500                         With prp
30510                           If .Name = "AppVersion" Then
30520                             arr_varOldFile(F_APPVER, lngDtaElem) = .Value
30530                           ElseIf .Name = "AppDate" Then
30540                             arr_varOldFile(F_APPDATE, lngDtaElem) = .Value
30550                           End If
30560                         End With  ' ** prp.
30570                       Next
30580                     End If
30590                   End With  ' ** doc.
30600                 Next

                      'lngOldTbls = 0&
                      'ReDim arr_varOldTbl(T_ELEMS, 0)

30610                 blnFound = False
30620                 For Each tdf In .TableDefs

                        'lngOldFlds = 0&
                        'ReDim arr_varOldFld(F_ELEMS, 0)

30630                   With tdf
30640                     If Left(.Name, 4) <> "MSys" And Left(.Name, 4) <> "~TMP" And _
                              .Connect = vbNullString Then  ' ** Skip those pesky system tables.

30650                       If Left(.Name, 3) = "m_V" Then
30660 On Error Resume Next
30670                         Set rstLnk = dbsLnk.OpenRecordset(.Name, dbOpenDynaset, dbReadOnly)
30680                         If ERR.Number = 0 Then
30690                           If gblnDev_NoErrHandle = True Then
30700 On Error GoTo 0
30710                           Else
30720 On Error GoTo ERRH
30730                           End If
30740                           blnFound = True
30750                           With rstLnk
30760                             .MoveFirst
30770                             strTmp04 = CStr(Nz(.Fields(0).Value, 0)) & "." & _
                                    CStr(Nz(.Fields(1).Value, 0)) & "." & CStr(Nz(.Fields(2).Value, 0))
30780                             arr_varOldFile(F_M_VER, lngDtaElem) = strTmp04
30790                             If gblnDev_Debug = False Then
30800                               strOldVersion = strTmp04
30810                             Else
30820                               strOldVersion = gstrCrtRpt_Version  ' ** Borrowed for conversion loop testing.
30830                             End If
30840                             .Close
30850                           End With  ' ** rstLnk.
30860                         Else
30870                           arr_varOldFile(F_NOTE, lngDtaElem) = arr_varOldFile(F_NOTE, lngDtaElem) & _
                                  " TBL: " & .Name & "  ERR: " & CStr(ERR.Number) & "  " & ERR.description
30880                           arr_varOldFile(F_NOTE, lngDtaElem) = Trim(arr_varOldFile(F_NOTE, lngDtaElem))
30890                           If gblnDev_NoErrHandle = True Then
30900 On Error GoTo 0
30910                           Else
30920 On Error GoTo ERRH
30930                           End If
30940                         End If
30950                         Set rstLnk = Nothing
30960                       End If
                            ' ** Table: m_VP         m_VD         m_VA
                            ' ** 0.     vp_MAIN      vd_MAIN      va_MAIN
                            ' ** 0.0.   vp_MINOR     vd_MINOR     va_MINOR
                            ' ** 0.0.0  vp_REVISION  vd_REVISION  va_REVISION

30970                     End If  ' ** Not a system table.

30980                   End With  ' ** This TableDef: tdf.

30990                 Next  ' ** For each TableDef: tdf.

31000                 If Left(strOldVersion, 3) = "1.7" Then
                        ' ** Throughout the 1.7.x series, all the m_VD, m_VP, and m_VA tables say 1.7.0,
                        ' ** so it's not very accurate. The one accurate table seems to be License Name.
31010                   blnFound = False
31020                 End If

                      ' ** Here, a False blnFound means we're going to keep looking to see if a more definitive version can be found.
31030                 If blnFound = False Then
                        ' ** Assuming this only looks at TrustDta.mdb (which I could trace if I felt like it),
                        ' ** not finding m_VD means this must be the one earlier example I have: v1.1.63.
                        ' ** If I change the Wise installation to also copy the Trust.mde, I could then link
                        ' ** the License Name table and get the version from there: ![Version].
                        ' ** Version 1.6.3 TrustDta.mdb is missing 6 tables:
                        ' **   m_REVCODE
                        ' **   m_REVCODE_TYPE
                        ' **   m_TBL
                        ' **   m_VD
                        ' **   RecurringType
                        ' **   Statement date
                        ' ** As well as these new ones:
                        ' **   _~rmcd
                        ' **   HiddenType
                        ' **   InvestmentObjective
                        ' **   LedgerHidden
                        ' **   tblControlType
                        ' **   tblDataTypeDb
                        ' **   tblDataTypeVb
                        ' **   tblDecimalPlaceDb
                        ' **   tblYesNo

                        ' ** See if Trust.mde is there, and try to open it.
31040                   strTmp04 = Parse_Path(arr_varOldFile(F_PTHFIL, lngDtaElem))  ' ** Module Function: modFileUtilities.
31050                   strTmp05 = strTmp04
31060                   strTmp06 = strTmp04
31070                   strTmp05 = strTmp05 & LNK_SEP & gstrFile_App & "." & gstrExt_AppRun
                        ' ** Look for Trust.mde.
31080                   If FileExists(strTmp05) = False Then
                          ' ** If blnContinue is set to False here, it aborts the whole conversion.
                          ' ** I don't think we need that strong a response!
                          'blnContinue = False
31090                   End If

31100                 End If

31110                 strTmp07 = vbNullString

31120                 If blnFound = False And blnContinue = True Then
                        ' ** Look for an MDW file.
31130                   strTmp06 = strTmp06 & LNK_SEP & gstrFile_SecurityName
31140                   If FileExists(strTmp06) = False Then
31150                     strTmp06 = strTmp04 & LNK_SEP & "Master.mdw"
31160                     If FileExists(strTmp06) = False Then
31170                       blnContinue = False
31180                     End If
31190                   End If
31200                   If blnContinue = False Then
                          ' ** Neither a TrustSec.mdw, nor a Master.mdw was found!
                          ' ** Do we have any idea which of TrustSec or a Master is called for?
                          ' ** Try TrustSec.mdw first.
31210                     strTmp07 = (gstrTrustDataLocation & gstrDir_Convert & LNK_SEP & "Convert_Empty" & LNK_SEP & "TrustSec.md_")
31220                     If FileExists(strTmp07) = True Then
31230                       strTmp06 = (gstrTrustDataLocation & gstrDir_Convert & LNK_SEP & gstrFile_SecurityName)
31240                       FileCopy strTmp07, strTmp06
31250                     Else
                            ' ** All we can do is just use the current one.
31260                       strTmp06 = (gstrTrustDataLocation & "TrustSec.mdw")
31270                     End If
31280                   End If
31290                 End If

31300                 If blnFound = False And blnContinue = True Then

                        'Set cnxn = New ADODB.Connection              ' ** Early binding.
31310                   Set cnxn = CreateObject("ADODB.Connection")  ' ** Late binding.

                        ' ** Open connection.  'VGC 09/24/2010: CHANGES!
31320                   strCnxn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                          "Data Source=" & strTmp05 & ";" & "Jet OLEDB:System database=" & strTmp06
                        'MsgBox "System Database for 1.7.1: " & strTmp06
                        ' ** NOTE: Jet 3.51 OLEDB provider is designed to open Access 97 databases only.
                        ' ** Jet 4.0 OLEDB provider is designed to open Access 2000 or Access 97 databases.
                        ' ** If you must use the Jet 3.51 Provider in the above examples,
                        ' ** change the provider name to "Microsoft.Jet.OLEDB.3.51."
                        ' ** MSJTER40.DLL
                        ' ** MSJT4JLT.DLL
                        ' ** MSJTER35.DLL
31330 On Error Resume Next
31340                   cnxn.Open ConnectionString:=strCnxn, UserId:="superuser", Password:=TA_SEC
31350                   If ERR.Number <> 0 Then
31360 On Error GoTo ERRH
                          ' ** If we tried TrustSec.mdw, try Master.mdw
31370                     If strTmp07 <> vbNullString Then
31380                       strTmp07 = (gstrTrustDataLocation & gstrDir_Convert & LNK_SEP & "Convert_Empty" & LNK_SEP & "Master.md_")
31390                       If FileExists(strTmp07) = True Then
31400                         strTmp06 = (gstrTrustDataLocation & gstrDir_Convert & LNK_SEP & "Master.mdw")
31410                         FileCopy strTmp07, strTmp06
31420                       Else
                              ' ** All we can do is just use the current one.
31430                         strTmp06 = (gstrTrustDataLocation & "TrustSec.mdw")
31440                       End If
                            ' ** Open connection.
31450                       strCnxn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                              "Data Source=" & strTmp05 & ";" & "Jet OLEDB:System database=" & strTmp06
31460 On Error Resume Next
31470                       cnxn.Open ConnectionString:=strCnxn, UserId:="superuser", Password:=TA_SEC
31480                       If ERR.Number <> 0 Then
                              ' ** Out of options!
31490 On Error GoTo ERRH
31500                         blnContinue = False
31510                         strTmp04 = "0"
31520                       Else
31530 On Error GoTo ERRH
31540                       End If
31550                     Else
                            ' ** Don't know what else to do.
31560                       blnContinue = False
31570                       strTmp04 = "0"
31580                     End If
31590                   Else
31600 On Error GoTo ERRH
31610                   End If

31620                   If blnContinue = True Then
31630                     strTmp05 = "SELECT [License name], [Version] FROM [License Name]"
                          'Set rsx1 = New ADODB.Recordset             ' ** Early binding.
31640                     Set rsx1 = CreateObject("ADODB.Recordset")  ' ** Late binding.
31650 On Error Resume Next
31660                     rsx1.Open strTmp05, cnxn, adOpenStatic, adLockReadOnly, adCmdText
31670                     If ERR.Number <> 0 Then
31680                       cnxn.Close
31690                       strTmp04 = "0"
                            'MsgBox "Error: " & CStr(ERR.Number) & vbCrLf & "Description: " & ERR.description
31700 On Error GoTo ERRH
31710                     Else
31720 On Error GoTo ERRH
31730                       If rsx1.EOF Then
                              ' ** Outta luck.
31740                         strTmp04 = "0"
31750                       Else
31760                         rsx1.MoveFirst
31770                         strTmp04 = CStr(Nz(rsx1.Fields("Version").Value, 0))
31780                       End If
31790                       rsx1.Close
31800                       cnxn.Close
31810                     End If
31820                   End If
31830                   Set rsx1 = Nothing
31840                   Set cnxn = Nothing

31850                   If strTmp04 <> "0" Then
31860                     intPos01 = InStr(strTmp04, ".")
31870                     If intPos01 > 0 Then
31880                       If Len(Mid(strTmp04, (intPos01 + 1))) > 1 Then
31890                         strTmp04 = Left(strTmp04, (intPos01 + 1)) & "." & Left(Mid(strTmp04, (intPos01 + 2)) & "00", 2)
31900                       Else
31910                         strTmp04 = strTmp04 & ".00"
31920                       End If
31930                     End If
31940                   Else
31950                     If Left(strOldVersion, 3) = "1.7" Then
                            ' ** Don't lose the previously retrieved version if it can't be more accurate.
31960                       strTmp04 = strOldVersion
31970                     Else
31980                       strTmp04 = "1.6.00?"
31990                     End If
32000                   End If
32010                   arr_varOldFile(F_M_VER, lngDtaElem) = strTmp04
32020                   arr_varOldFile(F_NOTE, lngDtaElem) = arr_varOldFile(F_NOTE, lngDtaElem) & _
                          " PRE 1.7.0 VERSION: " & strTmp04
32030                   arr_varOldFile(F_NOTE, lngDtaElem) = Trim(arr_varOldFile(F_NOTE, lngDtaElem))
32040                   strOldVersion = strTmp04

                        ' ** If we didn't find, or couldn't get into, Trust.mde,
                        ' ** don't let that stop the rest of the conversion.
32050                   blnContinue = True
32060                   intRetVal = 0

32070                 End If

32080                 .Close
32090               End With  ' ** dbsLnk.

32100               Set dbsLoc = CurrentDb

                    ' ** Update tblVersion_Conversion, by specified [vercid], [verold].
32110               Set qdf = dbsLoc.QueryDefs("qryVersion_Convert_03")
32120               With qdf.Parameters
32130                 ![vercid] = lngVerCnvID
32140                 ![verold] = strOldVersion
32150               End With  ' ** Parameters.
32160               qdf.Execute
32170               dbsLoc.Close

32180             End If  ' ** dbsLnk opens.

32190             .Close
32200           End With  ' ** wrkLnk.

32210         End If  ' ** Workspace opens: blnContinue.

32220       End If  ' ** blnConvert_TrustDta.

32230     End If  ' ** Conversion not already done.

32240   End If  ' ** Is a conversion.

EXITP:
32250   Set rsx1 = Nothing
32260   Set cnxn = Nothing
32270   Set qdf = Nothing
32280   Set tdf = Nothing
32290   Set doc = Nothing
32300   Set prp = Nothing
32310   Set rstLnk = Nothing
32320   Set dbsLoc = Nothing
32330   Set dbsLnk = Nothing
32340   Set wrkLnk = Nothing
32350   Version_GetOldVer = intRetVal
32360   Exit Function

ERRH:
32370   intRetVal = -9
32380   DoCmd.Hourglass False
32390   Select Case ERR.Number
        Case Else
32400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32410   End Select
32420   Resume EXITP

End Function

Public Sub ProgBar_Width_Conv(frm As Access.Form, dblWidth As Double, intMode As Integer)

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "ProgBar_Width_Conv"

        Dim strCtlName As String, blnVis As Boolean
        Dim lngX As Long

32510   With frm
32520     Select Case intMode
          Case 1
32530       blnVis = CBool(dblWidth)
32540       For lngX = 1& To 11&
32550         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
32560         .Controls(strCtlName).Visible = blnVis
32570       Next
32580     Case 2
32590       For lngX = 1& To 11&
32600         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
32610         .Controls(strCtlName).Width = dblWidth
32620       Next
32630     End Select
32640   End With

EXITP:
32650   Exit Sub

ERRH:
32660   Select Case ERR.Number
        Case Else
32670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32680   End Select
32690   Resume EXITP

End Sub

Public Sub Version_Up1_Etc(strPath As String, dblPB_Steps As Double, dblPB_ThisStep As Double, intMode As Integer)

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Up1_Etc"

        Dim strTmp04 As String, strTmp05 As String, blnTmp22 As Boolean

32710   Select Case intMode
        Case 1

          ' ** Step 25: Rename files.
32720     dblPB_ThisStep = 25#
32730     Version_Status 3, dblPB_ThisStep, "Renaming Files"  ' ** Module Function: modVersionConvertFuncs1.

          ' ** Change data files' extensions to BAK.
32740     If FileExists(strPath & LNK_SEP & Left(gstrFile_DataName, (Len(gstrFile_DataName) - 3)) & "BAK") = True Then  ' ** Module Function: modFileUtilities.
32750       Kill (strPath & LNK_SEP & Left(gstrFile_DataName, (Len(gstrFile_DataName) - 3)) & "BAK")
32760     End If
32770     Name (strPath & LNK_SEP & gstrFile_DataName) As (strPath & LNK_SEP & Left(gstrFile_DataName, (Len(gstrFile_DataName) - 3)) & "BAK")
32780     If FileExists(strPath & LNK_SEP & Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 3)) & "BAK") = True Then  ' ** Module Function: modFileUtilities.
32790       Kill (strPath & LNK_SEP & Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 3)) & "BAK")
32800     End If
32810     Name (strPath & LNK_SEP & gstrFile_ArchDataName) As (strPath & LNK_SEP & Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 3)) & "BAK")
32820     If FileExists(strPath & LNK_SEP & gstrFile_App & "." & gstrExt_AppRun) = True Then
            ' ** Trust.mde is there also.
32830       If FileExists(strPath & LNK_SEP & gstrFile_App & "_v" & Rem_Period(strOldVersion) & ".BAK") = True Then
32840         Kill (strPath & LNK_SEP & gstrFile_App & "_v" & Rem_Period(strOldVersion) & ".BAK")
32850       End If
32860       strTmp05 = strOldVersion
32870       If Right(strTmp05, 1) = "?" Then strTmp05 = Left(strTmp05, (Len(strTmp05) - 1))
32880       Name (strPath & LNK_SEP & gstrFile_App & "." & gstrExt_AppRun) As _
              (strPath & LNK_SEP & gstrFile_App & "_v" & Rem_Period(strTmp05) & ".BAK")
32890     End If
32900     blnTmp22 = False
32910     strTmp04 = gstrFile_SecurityName
32920     If FileExists(strPath & LNK_SEP & strTmp04) = True Then
            ' ** TrustSec.mdw is there also.
32930       blnTmp22 = True
32940     Else
32950       strTmp04 = "Master.mdw"
32960       If FileExists(strPath & LNK_SEP & strTmp04) = True Then
              ' ** Or rather, Master.mdw is there also.
32970         blnTmp22 = True
32980       End If
32990     End If
33000     If blnTmp22 = True Then
33010       If FileExists(strPath & LNK_SEP & Left(strTmp04, (Len(strTmp04) - 4)) & "_v" & Rem_Period(strOldVersion) & ".BAK") = True Then
33020         Kill (strPath & LNK_SEP & Left(strTmp04, (Len(strTmp04) - 4)) & "_v" & Rem_Period(strOldVersion) & ".BAK")
33030       End If
33040       strTmp05 = strOldVersion
33050       If Right(strTmp05, 1) = "?" Then strTmp05 = Left(strTmp05, (Len(strTmp05) - 1))
33060       Name (strPath & LNK_SEP & strTmp04) As _
              (strPath & LNK_SEP & Left(strTmp04, (Len(strTmp04) - 4)) & "_v" & Rem_Period(strTmp05) & ".BAK")
33070     End If
          ' ** Step {none}: End.
33080     dblPB_ThisStep = dblPB_Steps + 1#
33090     Version_Status 3, dblPB_ThisStep, "End"  ' ** Module Function: modVersionConvertFuncs1.

33100     Beep

33110   End Select

EXITP:
33120   Exit Sub

ERRH:
33130   Select Case ERR.Number
        Case Else
33140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
33150   End Select
33160   Resume EXITP

End Sub

Public Sub Version_Up2_Etc(strPath As String, msgResponse As VbMsgBoxResult, lngOldFiles As Long, arr_varOldFile As Variant, blnConvert_TrustDta As Boolean, blnConvert_TrstArch As Boolean, blnCheckForBoth_TrustDta As Boolean, lngDtaElem As Long, lngArchElem As Long, intRetVal1 As Integer, lngTmp13 As Long, lngTmp14 As Long, intMode As Integer)

33200 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Up2_Etc"

        Dim strTmp06 As String, strTmp09 As String, strTmp10 As String
        Dim lngX As Long

33210   Select Case intMode
        Case 1

33220     Select Case msgResponse
          Case vbYes
            ' ** Create a backup folder for the other backups found.
33230       strTmp06 = DirExists2(strPath, "Backup")  ' ** Module Function: modFileUtilities.
            ' ** Move the backups to the new folder and run the conversion.
            ' ** All BAK's will be moved, so if Trust.mde and TrustSec.mdw were also there, but
            ' ** their extensions have not been changed, we'll just have to proceed without them.
            ' ** NOTE: If TrustDta.mdb isn't empty, it'll let the user know after Version_IsEmpty() is run.
33240       For lngX = 0& To (lngOldFiles - 1&)
33250         If Rem_Ext(arr_varOldFile(F_FNAM, lngX)) = Rem_Ext(gstrFile_DataName) Then  ' ** Module Functions: modFileUtilities.
33260           If Parse_Ext(arr_varOldFile(F_FNAM, lngX)) = "BAK" Then  ' ** Module Functions: modFileUtilities.
33270             Name arr_varOldFile(F_PTHFIL, lngX) As strTmp06 & LNK_SEP & arr_varOldFile(F_FNAM, lngX)
33280           End If
33290         End If
33300         If Rem_Ext(arr_varOldFile(F_FNAM, lngX)) = Rem_Ext(gstrFile_ArchDataName) Then  ' ** Module Functions: modFileUtilities.
33310           If Parse_Ext(arr_varOldFile(F_FNAM, lngX)) = "BAK" Then  ' ** Module Functions: modFileUtilities.
33320             Name arr_varOldFile(F_PTHFIL, lngX) As strTmp06 & LNK_SEP & arr_varOldFile(F_FNAM, lngX)
33330           End If
33340         End If
33350       Next

            ' ** Continue the conversion.
            ' ** If they did an upgrade, then discovered it converted the wrong data,
            ' ** or there was something wrong with the data it did convert,
            ' ** there will will be BAK's in Convert_New. Answering yes must assume
            ' ** that they really do want to convert what's in the MDB's, so we'll
            ' ** have to replace TrustDta.mdb and TrstArch.mdb with empties,
            ' ** as well as moving the old BAK's out of the way.
            ' ** My test give these results to the above variables.
            ' **   blnCheckForBoth_TrustDta = True
            ' **   blnConvert_TrustDta = False
            ' **   blnConvert_TrstArch = False
            ' **   blnArchiveNotPresent = False

            ' ** Rename the current data files in the data location to BAK, then copy over empties.
33360       strTmp09 = vbNullString: strTmp10 = vbNullString
33370       strTmp09 = gstrTrustDataLocation & Left(gstrFile_DataName, (Len(gstrFile_DataName) - 3)) & "BAK"
33380       If FileExists(strTmp09) = True Then  ' ** Module Function: modFileUtilities.
33390         Kill strTmp09
33400       End If
33410       Name gstrTrustDataLocation & gstrFile_DataName As strTmp09
33420       strTmp10 = gstrTrustDataLocation & Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 3)) & "BAK"
33430       If FileExists(strTmp10) = True Then  ' ** Module Function: modFileUtilities.
33440         Kill strTmp10
33450       End If
33460       Name gstrTrustDataLocation & gstrFile_ArchDataName As strTmp10
33470       strTmp09 = gstrTrustDataLocation & gstrDir_Convert & LNK_SEP & gstrDir_ConvertEmpty
33480       If DirExists(strTmp09) = True Then  ' ** Module Function: modFileUtilities.
33490         strTmp09 = gstrTrustDataLocation & gstrDir_Convert & LNK_SEP & gstrDir_ConvertEmpty & LNK_SEP & _
                Left(gstrFile_DataName, (Len(gstrFile_DataName) - 1)) & "_"
33500         If FileExists(strTmp09) = True Then  ' ** Module Function: modFileUtilities.
33510           FileCopy strTmp09, gstrTrustDataLocation & gstrFile_DataName
33520           strTmp10 = gstrTrustDataLocation & gstrDir_Convert & LNK_SEP & gstrDir_ConvertEmpty & LNK_SEP & _
                  Left(gstrFile_ArchDataName, (Len(gstrFile_ArchDataName) - 1)) & "_"
33530           If FileExists(strTmp10) = True Then  ' ** Module Function: modFileUtilities.
33540             FileCopy strTmp10, gstrTrustDataLocation & gstrFile_ArchDataName
33550             blnConvert_TrustDta = True
33560             blnConvert_TrstArch = True
33570             blnCheckForBoth_TrustDta = False
33580             If lngTmp13 <> 0& Then lngDtaElem = lngTmp13
33590             If lngTmp14 <> 0& Then lngArchElem = lngTmp14
33600           Else
                  ' ** No TrstArch.md_ in Convert_Empty.
33610             intRetVal1 = -9
33620             MsgBox "A file needed by Trust Accountant to complete the conversion is missing." & vbCrLf & _
                    "Contact Delta Data, Inc., for assistance.", vbCritical + vbOKOnly, "File Not Found: TrstArch.md_"
33630           End If
33640         Else
                ' ** No TrustDta.md_ in Convert_Empty.
33650           intRetVal1 = -9
33660           MsgBox "A file needed by Trust Accountant to complete the conversion is missing." & vbCrLf & _
                  "Contact Delta Data, Inc., for assistance.", vbCritical + vbOKOnly, "File Not Found: TrustDta.md_"
33670         End If
33680       Else
              ' ** No Convert_Empty directory.
33690         intRetVal1 = -9
33700         MsgBox "Files needed by Trust Accountant to complete the conversion are missing." & vbCrLf & _
                "Contact Delta Data, Inc., for assistance.", vbCritical + vbOKOnly, "Folder Not Found"
33710       End If
33720     Case vbNo
            ' ** Create a backup folder for the originals found.
33730       strTmp06 = DirExists2(strPath, "Backup")  ' ** Module Function: modFileUtilities.
            ' ** Move the originals to the new folder and don't run the conversion.
            ' ** If there are still originals of Trust.mde and TrustSec.mdw, just leave them alone.
33740       For lngX = 0& To (lngOldFiles - 1&)
33750         If arr_varOldFile(F_FNAM, lngX) = gstrFile_DataName Then
33760           Name arr_varOldFile(F_PTHFIL, lngX) As strTmp06 & LNK_SEP & arr_varOldFile(F_FNAM, lngX)
33770         End If
33780       Next
33790       For lngX = 0& To (lngOldFiles - 1&)
33800         If arr_varOldFile(F_FNAM, lngX) = gstrFile_ArchDataName Then
33810           Name arr_varOldFile(F_PTHFIL, lngX) As strTmp06 & LNK_SEP & arr_varOldFile(F_FNAM, lngX)
33820         End If
33830       Next
33840     End Select

33850   End Select

EXITP:
33860   Exit Sub

ERRH:
33870   Select Case ERR.Number
        Case Else
33880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
33890   End Select
33900   Resume EXITP

End Sub

Public Sub Version_Stat2_Etc(intMode As Integer, lngOff1 As Long, lngOff2 As Long, lngStat1Orig_Top As Long, lngStat2Orig_Top As Long, arr_dblPB_ThisIncr As Variant, dblPB_Width As Double, frm As Access.Form)

34000 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Stat2_Etc"

        Dim lngTpp As Long

34010   With frm

          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
34020     lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

34030     Select Case intMode
          Case 1

            ' ** Move status stuff to a better looking arrangement.
34040       lngOff1 = .TAVer_Old_lbl.Top - lngStat1Orig_Top '.Status1_lbl.Top
34050       lngOff2 = lngStat2Orig_Top - lngStat1Orig_Top '.Status1_lbl.Top
34060       .TAVer_Old_lbl.Top = .TAVer_Old_lbl.Top - lngOff1
34070       .TAVer_Old.Top = .TAVer_Old.Top - lngOff1
34080       .TAVer_Old_RelDate.Top = .TAVer_Old_RelDate.Top - lngOff1
34090       .TAVer_New_lbl.Top = .TAVer_New_lbl.Top - lngOff1
34100       .TAVer_New.Top = .TAVer_New.Top - lngOff1
34110       .TAVer_New_RelDate.Top = .TAVer_New_RelDate.Top - lngOff1
34120       .TAVer_Arrow.Top = .TAVer_Arrow.Top - lngOff1
34130       .TAVer_box.Top = .TAVer_box.Top - lngOff1
34140       .PathFile_TrustData.Top = .PathFile_TrustData.Top - lngOff1
34150       .PathFile_TrustData_lbl.Top = .PathFile_TrustData_lbl.Top - lngOff1
34160       .PathFile_TrustData.ForeColor = CLR_DKGRY
34170       .PathFile_TrustData.BackColor = CLR_LTTEAL
34180       .PathFile_TrustData_box.Top = .PathFile_TrustData_box.Top - lngOff1
34190       .PathFile_TrustData_box2.Top = .PathFile_TrustData_box2.Top - lngOff1
34200       .PathFile_TrustData_hline01.Top = .PathFile_TrustData_hline01.Top - lngOff1
34210       .PathFile_TrustData_hline02.Top = .PathFile_TrustData_hline02.Top - lngOff1
34220       .PathFile_TrustData_hline03.Top = .PathFile_TrustData_hline03.Top - lngOff1
34230       .PathFile_TrustData_vline01.Top = .PathFile_TrustData_vline01.Top - lngOff1
34240       .PathFile_TrustData_vline02.Top = .PathFile_TrustData_vline02.Top - lngOff1
34250       .PathFile_TrustData_vline03.Top = .PathFile_TrustData_vline03.Top - lngOff1
34260       .PathFile_TrustData_vline04.Top = .PathFile_TrustData_vline04.Top - lngOff1
34270       .PathFile_TrustArchive.ForeColor = CLR_DKGRY
34280       .PathFile_TrustArchive.BackColor = CLR_LTTEAL
34290       .PathFile_TrustArchive.Top = .PathFile_TrustArchive.Top - lngOff1
34300       .PathFile_TrustArchive_lbl.Top = .PathFile_TrustArchive_lbl.Top - lngOff1
34310       .PathFile_TrustArchive_box.Top = .PathFile_TrustArchive_box.Top - lngOff1
34320       .PathFile_TrustArchive_box2.Top = .PathFile_TrustArchive_box2.Top - lngOff1
34330       .PathFile_TrustArchive_hline01.Top = .PathFile_TrustArchive_hline01.Top - lngOff1
34340       .PathFile_TrustArchive_hline02.Top = .PathFile_TrustArchive_hline02.Top - lngOff1
34350       .PathFile_TrustArchive_hline03.Top = .PathFile_TrustArchive_hline03.Top - lngOff1
34360       .PathFile_TrustArchive_vline01.Top = .PathFile_TrustArchive_vline01.Top - lngOff1
34370       .PathFile_TrustArchive_vline02.Top = .PathFile_TrustArchive_vline02.Top - lngOff1
34380       .PathFile_TrustArchive_vline03.Top = .PathFile_TrustArchive_vline03.Top - lngOff1
34390       .PathFile_TrustArchive_vline04.Top = .PathFile_TrustArchive_vline04.Top - lngOff1
34400       .Status1_lbl.Top = (.PathFile_TrustArchive.Top + lngOff1) - (12& * lngTpp)
34410       .Status2_lbl.Top = (.Status1_lbl.Top + lngOff2) + (4& * lngTpp)

34420     Case 2

            ' ** Weight the steps.
34430       arr_dblPB_ThisIncr(1) = CDbl((dblPB_Width / 100#) * 3#)   '
34440       arr_dblPB_ThisIncr(2) = CDbl((dblPB_Width / 100#) * 2#)   '
34450       arr_dblPB_ThisIncr(3) = CDbl((dblPB_Width / 100#) * 2#)   '
34460       arr_dblPB_ThisIncr(4) = CDbl((dblPB_Width / 100#) * 2#)   '
34470       arr_dblPB_ThisIncr(5) = CDbl((dblPB_Width / 100#) * 2#)   '
34480       arr_dblPB_ThisIncr(6) = CDbl((dblPB_Width / 100#) * 2#)   '
34490       arr_dblPB_ThisIncr(7) = CDbl((dblPB_Width / 100#) * 2#)   '
34500       arr_dblPB_ThisIncr(8) = CDbl((dblPB_Width / 100#) * 2#)   '
34510       arr_dblPB_ThisIncr(9) = CDbl((dblPB_Width / 100#) * 2#)   '
34520       arr_dblPB_ThisIncr(10) = CDbl((dblPB_Width / 100#) * 2#)  '
34530       arr_dblPB_ThisIncr(11) = CDbl((dblPB_Width / 100#) * 2#)  '         8  1's (8)
34540       arr_dblPB_ThisIncr(12) = CDbl((dblPB_Width / 100#) * 8#)  '        34  2's (17)
34550       arr_dblPB_ThisIncr(13) = CDbl((dblPB_Width / 100#) * 8#)  '        12  3's (4)
34560       arr_dblPB_ThisIncr(14) = CDbl((dblPB_Width / 100#) * 10#) '        16  8's (2)
34570       arr_dblPB_ThisIncr(15) = CDbl((dblPB_Width / 100#) * 10#) '        30 10's (3)
34580       arr_dblPB_ThisIncr(16) = CDbl((dblPB_Width / 100#) * 10#) '       ===
34590       arr_dblPB_ThisIncr(17) = CDbl((dblPB_Width / 100#) * 2#)  '       100
34600       arr_dblPB_ThisIncr(18) = CDbl((dblPB_Width / 100#) * 2#)  '
34610       arr_dblPB_ThisIncr(19) = CDbl((dblPB_Width / 100#) * 2#)  '
34620       arr_dblPB_ThisIncr(20) = CDbl((dblPB_Width / 100#) * 2#)  '
34630       arr_dblPB_ThisIncr(21) = CDbl((dblPB_Width / 100#) * 2#)  '
34640       arr_dblPB_ThisIncr(22) = CDbl((dblPB_Width / 100#) * 2#)  '
34650       arr_dblPB_ThisIncr(23) = CDbl((dblPB_Width / 100#) * 3#)  '
34660       arr_dblPB_ThisIncr(24) = CDbl((dblPB_Width / 100#) * 3#)  '
34670       arr_dblPB_ThisIncr(25) = CDbl((dblPB_Width / 100#) * 3#)  '
34680       arr_dblPB_ThisIncr(26) = CDbl((dblPB_Width / 100#) * 1#)  '        'tblCheckMemo
34690       arr_dblPB_ThisIncr(27) = CDbl((dblPB_Width / 100#) * 1#)  '        'tblCheckReconcile_Account
34700       arr_dblPB_ThisIncr(28) = CDbl((dblPB_Width / 100#) * 1#)  '        'tblCheckReconcile_Item
34710       arr_dblPB_ThisIncr(29) = CDbl((dblPB_Width / 100#) * 1#)  '        'tblCheckPOSPay
34720       arr_dblPB_ThisIncr(30) = CDbl((dblPB_Width / 100#) * 1#)  '        'tblCheckPOSPay_Detail
34730       arr_dblPB_ThisIncr(31) = CDbl((dblPB_Width / 100#) * 1#)  '        'tblCheckBank
34740       arr_dblPB_ThisIncr(32) = CDbl((dblPB_Width / 100#) * 1#)  '        'tblCheckVoid
34750       arr_dblPB_ThisIncr(33) = CDbl((dblPB_Width / 100#) * 1#)  '        'tblRecurringAux1099
34760       arr_dblPB_ThisIncr(34) = CDbl((dblPB_Width / 100#) * 2#)  '        'Rename files

34770     Case 3

            ' ** Arrange the paths.
34780       lngOff1 = (6& * lngTpp) '90&
34790       .PathFile_TrustData.Top = ((.TAVer_Old.Top + .TAVer_Old.Height) + lngOff1) + (2& * lngTpp)
34800       .PathFile_TrustData_lbl.Top = ((.TAVer_Old.Top + .TAVer_Old.Height) + lngOff1) + (3& * lngTpp)
34810       .PathFile_TrustData.Width = ((.PathFile_TrustData.Width - .PathFile_TrustArchive_lbl.Width) - (4& * lngTpp))  ' ** The wider of the two.
34820       .PathFile_TrustData.Left = ((.PathFile_TrustData.Left + .PathFile_TrustArchive_lbl.Width) + (4& * lngTpp))
34830       .PathFile_TrustData_box.Visible = False
34840       .PathFile_TrustData_box2.Visible = False
34850       .PathFile_TrustData_hline01.Visible = False
34860       .PathFile_TrustData_hline02.Visible = False
34870       .PathFile_TrustData_hline03.Visible = False
34880       .PathFile_TrustData_vline01.Visible = False
34890       .PathFile_TrustData_vline02.Visible = False
34900       .PathFile_TrustData_vline03.Visible = False
34910       .PathFile_TrustData_vline04.Visible = False
34920       .PathFile_TrustArchive.Top = ((.PathFile_TrustData.Top + .PathFile_TrustData.Height) + lngOff1) + (2& * lngTpp)
34930       .PathFile_TrustArchive_lbl.Top = ((.PathFile_TrustData.Top + .PathFile_TrustData.Height) + lngOff1) + (3& * lngTpp)
34940       .PathFile_TrustArchive.Width = ((.PathFile_TrustArchive.Width - .PathFile_TrustArchive_lbl.Width) - (4& * lngTpp))  ' ** The wider of the two.
34950       .PathFile_TrustArchive.Left = ((.PathFile_TrustArchive.Left + .PathFile_TrustArchive_lbl.Width) + (4& * lngTpp))
34960       .PathFile_TrustArchive_box.Visible = False
34970       .PathFile_TrustArchive_box2.Visible = False
34980       .PathFile_TrustArchive_hline01.Visible = False
34990       .PathFile_TrustArchive_hline02.Visible = False
35000       .PathFile_TrustArchive_hline03.Visible = False
35010       .PathFile_TrustArchive_vline01.Visible = False
35020       .PathFile_TrustArchive_vline02.Visible = False
35030       .PathFile_TrustArchive_vline03.Visible = False
35040       .PathFile_TrustArchive_vline04.Visible = False
35050       lngOff1 = 120&
35060       lngOff2 = .Status3.Left - .Status3_lbl.Left
35070       .Status3.Top = ((.PathFile_TrustArchive.Top + .PathFile_TrustArchive.Height) + lngOff1)
35080       .Status3_lbl.Top = (.Status3.Top + (2& * lngTpp))
35090       .Status3.Left = .ProgBar_box.Left
35100       .Status3_lbl.Left = ((.Status3.Left - .Status3_lbl.Width) - (4& * lngTpp))
35110       .Status3.Width = 7200&  ' ** 5"
35120       .Status3.Height = ((.ProgBar_box.Top - .Status3.Top) - lngOff1) - lngTpp
35130       .Status1_lbl.Visible = False
35140       .Status2_lbl.Visible = False

35150       .cmdPrintReport.Top = (.Status3_lbl.Top + .Status3_lbl.Height) + lngOff1
35160       .cmdPrintReport.Left = (.Status3.Left - .cmdPrintReport.Width) - (4& * lngTpp)
35170       .cmdPrintReport_raised_img.Top = .cmdPrintReport.Top
35180       .cmdPrintReport_raised_img.Left = .cmdPrintReport.Left
35190       .cmdPrintReport_raised_semifocus_dots_img.Top = .cmdPrintReport.Top
35200       .cmdPrintReport_raised_semifocus_dots_img.Left = .cmdPrintReport.Left
35210       .cmdPrintReport_raised_focus_img.Top = .cmdPrintReport.Top
35220       .cmdPrintReport_raised_focus_img.Left = .cmdPrintReport.Left
35230       .cmdPrintReport_raised_focus_dots_img.Top = .cmdPrintReport.Top
35240       .cmdPrintReport_raised_focus_dots_img.Left = .cmdPrintReport.Left
35250       .cmdPrintReport_sunken_focus_dots_img.Top = .cmdPrintReport.Top
35260       .cmdPrintReport_sunken_focus_dots_img.Left = .cmdPrintReport.Left
35270       .cmdPrintReport_raised_img_dis.Top = .cmdPrintReport.Top
35280       .cmdPrintReport_raised_img_dis.Left = .cmdPrintReport.Left
35290       .cmdPrintReport.Visible = True
35300       .cmdPrintReport_raised_img.Visible = True

35310     End Select
35320   End With

EXITP:
35330   Exit Sub

ERRH:
35340   Select Case ERR.Number
        Case Else
35350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
35360   End Select
35370   Resume EXITP

End Sub
