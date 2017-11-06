Attribute VB_Name = "modVersionConvertFuncs3"
Option Compare Database
Option Explicit

'VGC 10/27/2017: CHANGES!

Private Const THIS_NAME As String = "modVersionConvertFuncs3"

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

' ** Array: arr_varMasterAsset().
Private Const MA_ELEMS As Integer = 6  ' ** Array's first-element UBound().
Private Const MA_OLD_ANO As Integer = 0
Private Const MA_NEW_ANO As Integer = 1
Private Const MA_NAM     As Integer = 2
Private Const MA_OLD_MVC As Integer = 3
Private Const MA_NEW_MVC As Integer = 4
Private Const MA_ERR     As Integer = 5
Private Const MA_ERRDESC As Integer = 6

' ** Array: arr_varAcctType(), arr_varAssetType().
Private Const AT_TYP As Integer = 0
'Private Const AT_DSC As Integer = 1

' ** Array: arr_varInvestObj().
Private lngInvestObjs As Long, arr_varInvestObj() As Variant
Private Const IO_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const IO_ID  As Integer = 0
Private Const IO_NAM As Integer = 1
Private Const IO_NEW As Integer = 2

' ** Array: arr_varStat().
Private lngStats As Long, arr_varStat() As Variant
Private Const STAT_ELEMS As Integer = 3  ' ** Array's first-element UBound().
Private Const STAT_ORD As Integer = 0
Private Const STAT_NAM As Integer = 1
Private Const STAT_CNT As Integer = 2
Private Const STAT_DSC As Integer = 3

' ** Array: arr_varOldFile().
Private Const F_ELEMS As Integer = 11  ' ** Array's first-element UBound().
Private Const F_FNAM    As Integer = 0
Private Const F_PTHFIL  As Integer = 1
Private Const F_DATA    As Integer = 2
Private Const F_CONV    As Integer = 3
Private Const F_TA_VER  As Integer = 4
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
Private Const T_TNAMN As Integer = 1
Private Const T_FLDS  As Integer = 2
Private Const T_F_ARR As Integer = 3

' ** Array: arr_varDupeUnk().
Private lngDupeUnks As Long, arr_varDupeUnk() As Variant
Private Const DU_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const DU_TYP As Integer = 0
Private Const DU_TBL As Integer = 1

Private Const A99_INC As String = "INCOME O/U"
Private Const A99_SUS As String = "SUSPENSE"

Private lngErrNum As Long, lngErrLine As Long, strErrDesc As String
' **

Public Function Version_Upgrade_04(blnContinue As Boolean, blnConvert_TrustDta As Boolean, lngAccts As Long, arr_varAcct As Variant, lngAcctTypes As Long, arr_varAcctType As Variant, lngAssetTypes As Long, arr_varAssetType As Variant, lngMasterAssets As Long, lngTmp01 As Long, arr_varTmp02 As Variant, lngTmp03 As Long, arr_varTmp04 As Variant, lngTmp05 As Long, arr_varTmp06 As Variant, lngOldFiles As Long, arr_varOldFile As Variant, lngOldTbls As Long, arr_varOldTbl As Variant, dblPB_ThisStep As Double, lngTrustDtaDbsID As Long, strKeyTbl As String, strTruncatedFields As String, strAcct99_IncomeOU As String, strAcct99_Suspense As String, lngArchElem As Long, lngDupeNum As Long, wrkLnk As DAO.Workspace, dbsLnk As DAO.Database, wrkLoc As DAO.Workspace, dbsLoc As DAO.Database) As Integer
' ** This continues the conversion process with the most complex tables.
' ** Tables converted here:
' **   Account       '04/27/2016: curr_id ADDED!
' **   Balance       '04/27/2016: curr_id ADDED!
' **   masterasset   '04/27/2016: curr_id ADDED!
' **   ActiveAssets  '04/27/2016: curr_id CHECKED!
' **
' ** Return values:
' **    0  OK
' **   -6  Index/Key
' **   -9  Error

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Upgrade_04"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstLoc1 As DAO.Recordset, rstLoc2 As DAO.Recordset, rstLoc3 As DAO.Recordset, rstLnk As DAO.Recordset
        Dim fld As DAO.Field
        Dim strCurrTblName As String, lngCurrTblID As Long, strCurrKeyFldName As String, lngCurrKeyFldID As Long
        Dim lngRecs As Long, lngFlds As Long
        Dim lngErrNum As Long, lngErrLine As Long, strErrDesc As String
        Dim blnFound As Boolean, blnFound2 As Boolean
        Dim varTmp00 As Variant, strTmp04 As String, strTmp05 As String, strTmp06 As String, strTmp07 As String, strTmp08 As String
        Dim lngTmp13 As Long, lngTmp14 As Long, lngTmp15 As Long
        Dim blnTmp22 As Boolean, blnTmp23 As Boolean, blnTmp24 As Boolean, datTmp28 As Date
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim intRetVal As Integer

110     If gblnDev_NoErrHandle = True Then
120   On Error GoTo 0
130     End If

140     intRetVal = 0
150     lngRecs = 0&

160     lngInvestObjs = 0&
170     ReDim arr_varInvestObj(IO_ELEMS, 0)

180     lngStats = 0&
190     ReDim arr_varStat(STAT_ELEMS, 0)

200     If lngTmp03 > 0 Then
210       For lngX = 0& To (lngTmp03 - 1&)
220         lngStats = lngStats + 1&
230         lngE = lngStats - 1&
240         ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
250         For lngY = 0& To STAT_ELEMS
260           arr_varStat(lngY, lngE) = arr_varTmp04(lngY, lngX)
270         Next
280       Next
290     End If

300     lngDupeUnks = 0&
310     ReDim arr_varDupeUnk(DU_ELEMS, 0)

320     If lngTmp05 > 0 Then
330       For lngX = 0& To (lngTmp05 - 1&)
340         lngDupeUnks = lngDupeUnks + 1&
350         lngE = lngDupeUnks - 1&
360         ReDim Preserve arr_varDupeUnk(STAT_ELEMS, lngE)
370         For lngY = 0& To DU_ELEMS
380           arr_varDupeUnk(lngY, lngE) = arr_varTmp06(lngY, lngX)
390         Next
400       Next
410     End If

420     If blnContinue = True Then  ' ** Is a conversion.

430       If blnContinue = True Then  ' ** Conversion not already done.

440         If blnConvert_TrustDta = True Then

450           If blnContinue = True Then  ' ** Workspace opens.

460             With wrkLnk

470               If blnContinue = True Then  ' ** Open dbsLnk.

480                 With dbsLnk

490                   If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ** Get a list of Account Types.
500                     Set qdf = dbsLoc.QueryDefs("qryAccountType_01")
510                     Set rstLoc1 = qdf.OpenRecordset
520                     With rstLoc1
530                       .MoveLast
540                       lngAcctTypes = .RecordCount
550                       .MoveFirst
560                       arr_varAcctType = .GetRows(lngAcctTypes)
                          ' ************************************************
                          ' ** Array: arr_varAcctType()
                          ' **
                          ' **   Field  Element  Name           Constant
                          ' **   =====  =======  =============  ==========
                          ' **     1       0     accounttype    AT_TYP
                          ' **     2       1     description    AT_DSC
                          ' **
                          ' ************************************************
                          ' ** Layout the same as arr_varAssetType().
570                       .Close
580                     End With  ' ** rstLoc1.

                        ' ** Get a list of the Investment Objective options.
590                     Set qdf = dbsLoc.QueryDefs("qryInvestmentObjective_01")
600                     Set rstLoc1 = qdf.OpenRecordset
610                     With rstLoc1
620                       .MoveLast
630                       lngRecs = .RecordCount
640                       .MoveFirst
650                       For lngX = 1& To lngRecs
660                         lngInvestObjs = lngInvestObjs + 1&
670                         lngE = lngInvestObjs - 1&
680                         ReDim Preserve arr_varInvestObj(IO_ELEMS, lngE)
                            ' ************************************************
                            ' ** Array: arr_varInvestObj()
                            ' **
                            ' **   Field  Element  Name           Constant
                            ' **   =====  =======  =============  ==========
                            ' **     1       0     invobj_id      IO_ID
                            ' **     2       1     invobj_name    IO_NAM
                            ' **     3       2     New (Y/N)      IO_NEW
                            ' **
                            ' ************************************************
690                         arr_varInvestObj(IO_ID, lngE) = ![invobj_id]
700                         arr_varInvestObj(IO_NAM, lngE) = ![invobj_name]
710                         arr_varInvestObj(IO_NEW, lngE) = CBool(False)
                            '![Username]
                            '![DateCreated]
                            '![DateModified]
720                         If lngX < lngRecs Then .MoveNext
730                       Next
740                       .Close
750                     End With  ' ** rstLoc1.

                        ' ******************************
                        ' ** Table: Account.
                        ' ******************************

                        ' ** Step 12: Account.
760                     dblPB_ThisStep = 12#
770                     Version_Status 3, dblPB_ThisStep, "Account"  ' ** Function: Below.

780                     strCurrTblName = "account"
790                     lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

800                     blnFound = False: lngRecs = 0&
810                     For lngX = 0& To (lngOldTbls - 1&)
820                       If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
830                         blnFound = True
840                         Exit For
850                       End If
860                     Next

870                     If blnFound = True Then
880                       Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
890                       With rstLnk
900                         If .BOF = True And .EOF = True Then
                              ' ** This has to have records!
910                         Else
920                           strCurrKeyFldName = "accountno"
930                           lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
940                           Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
950                           Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Of 16 example TrustDta.mdb's, some have 82 fields, one 87, and some 89.
                              ' ** Current field count is 95 fields.
                              ' ** Fields needing special checks.
                              ' **   ![accountno]         : dbText 15
                              ' **   ![related_accountno] : Let's wait for the form to clean this up; just copy.
                              ' **   ![accounttype]
                              ' **   ![admin]             : Long Integer from adminofficer; new one in arr_varAcct().
                              ' **   ![icash]             : Can't be Null.
                              ' **   ![pcash]             : Can't be Null.
                              ' **   ![cost]              : Can't be Null.
                              ' **   ![Schedule_ID]       : Long Integer from Schedule; new one in arr_varAcct().
                              ' **   ![cotrustee]         : Yes/No
                              ' **   ![amendments]        : Yes/No
                              ' **   ![courtsupervised]   : Yes/No
                              ' **   ![discretion]        : Yes/No
                              ' **   ![investmentobj]
                              ' **   ![taxlot]            : dbText, default AssetNo (using temporarily).
                              ' ** Table: account
                              ' **   89 fields
                              ' ** Referenced by:
                              ' **   Table: ActiveAssets                  {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: asset                         {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: Balance                       {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: FeeCalculations               {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: journal                       {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: Journal Map                   {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: ledger                        {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: LedgerArchive                 {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: LedgerHidden                  {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: reviewfreq                    {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: statementfreq                 {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: tblMasterBalance              {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tblPortfolioModeling          {Linked}
                              ' **     Field: [accountno]
                              ' **   Table: tblPortfolioModeling2         {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tblTemplate_ActiveAssets      {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tblTemplate_Asset             {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tblTemplate_Journal           {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tblTemplate_Ledger            {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tblTemplate_LedgerHidden      {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpAccountInfo                {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpAssetList2                 {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpAssetList3                 {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpAveragePrice               {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpCapitalGainsAndLosses      {Local}
                              ' **     Field: [accountno]                   dbText 255
                              ' **   Table: tmpCourtReportData            {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpCourtReportData2           {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpCourtReportData3           {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpCourtReportData4           {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpEdit03                      {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpTaxReports                 {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpTrx1                       {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpTrx2                       {Local}
                              ' **     Field: [accountno]
                              ' **   Table: tmpTrx3                       {Local}
                              ' **     Field: [accountno]
960                           .MoveLast
970                           lngRecs = .RecordCount
980                           Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Function: Below.
990                           .MoveFirst
1000                          lngFlds = 0&
1010                          For lngX = 0& To (lngOldFiles - 1&)
1020                            If arr_varOldFile(F_FNAM, lngX) = gstrFile_DataName Then
1030                              arr_varTmp02 = arr_varOldFile(F_T_ARR, lngX)
1040                              lngOldTbls = (UBound(arr_varTmp02, 2) + 1)
1050                              For lngY = 0& To (lngOldTbls - 1&)
1060                                If arr_varTmp02(T_TNAM, lngY) = strCurrTblName Then
1070                                  lngFlds = arr_varTmp02(T_FLDS, lngY)
1080                                  Exit For
1090                                End If
1100                              Next
1110                              Exit For
1120                            End If
1130                          Next
1140                          For lngX = 1& To lngRecs
1150                            blnTmp23 = False  ' ** Used for curr_id.
1160                            Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Function: Below.
1170                            If ![accountno] = A99_INC Or ![accountno] = A99_SUS Or _
                                    ![accountno] = ("99-" & A99_INC) Or ![accountno] = ("99-" & A99_SUS) Then
                                  ' ** These are distributed with the install.
                                  ' ** Also check for a mixture of prefixed and non-prefixed
                                  ' ** numbers, depending on the value of gblnAccountNoWithType.
                                  ' **   99-INCOME O/U -->  INCOME O/U
                                  ' **   99-SUSPENSE   -->  SUSPENSE
                                  ' ** Do this check at the end of the conversion.
1180                              For lngY = 0& To (lngAccts - 1&)
1190                                If arr_varAcct(A_NUM, lngY) = ![accountno] Then
1200                                  arr_varAcct(A_DROPPED, lngY) = CBool(True)
1210                                  Exit For
1220                                End If
1230                              Next
1240                            Else
                                  ' ** Add the record to the new table.
1250                              rstLoc1.AddNew
1260                              For Each fld In .Fields
1270                                Select Case fld.Name
                                    Case "accounttype"
1280                                  If IsNull(![accounttype]) = False Then
1290                                    strTmp04 = Trim(fld.Value)
1300                                    blnFound = False
1310                                    For lngY = 0& To (lngAcctTypes - 1&)
1320                                      If arr_varAcctType(AT_TYP, lngY) = strTmp04 Then
1330                                        blnFound = True
1340                                        Exit For
1350                                      End If
1360                                    Next
1370                                    If blnFound = False Then
1380                                      rstLoc1.Fields(fld.Name) = "85"  ' ** Other.
1390                                    Else
1400                                      rstLoc1.Fields(fld.Name) = strTmp04
1410                                    End If
1420                                    blnFound = True  ' ** Reset.
1430                                  Else
1440                                    rstLoc1.Fields(fld.Name) = "85"  ' ** Other.
1450                                  End If
1460                                Case "shortname"
                                      ' ** Check for quotes.
1470                                  strTmp04 = Trim(fld.Value)
1480                                  strTmp05 = FixQuotes(strTmp04)  ' ** Module Function: modStringFuncs.
1490                                  If strTmp05 <> strTmp04 Then
1500                                    rstLoc1.Fields(fld.Name) = strTmp05
1510                                  Else
1520                                    rstLoc1.Fields(fld.Name) = fld.Value
1530                                  End If
1540                                Case "legalname"
                                      ' ** Check for quotes.
1550                                  If IsNull(fld.Value) = False Then
1560                                    If Trim(fld.Value) <> vbNullString Then
1570                                      strTmp04 = Trim(fld.Value)
1580                                      strTmp05 = FixQuotes(strTmp04)  ' ** Module Function: modStringFuncs.
1590                                      If strTmp05 <> strTmp04 Then
1600                                        rstLoc1.Fields(fld.Name) = strTmp05
1610                                      Else
1620                                        rstLoc1.Fields(fld.Name) = fld.Value
1630                                      End If
1640                                    End If
1650                                  End If
1660                                Case "adminno", "admin"
1670                                  If IsNull(fld.Value) = False Then
1680                                    For lngY = 0& To (lngAccts - 1&)
1690                                      If arr_varAcct(A_NUM, lngY) = ![accountno] Then
1700  On Error Resume Next
1710                                        rstLoc1.Fields(fld.Name) = IIf(arr_varAcct(A_ADMIN_N, lngY) = 0&, 1&, arr_varAcct(A_ADMIN_N, lngY))  ' ** Here's where it SHOULD get
1720                                        If ERR.Number <> 0 Then                                  ' ** the new, CORRECT, adminno!
1730                                          If gblnDev_NoErrHandle = True Then
1740  On Error GoTo 0
1750                                          Else
1760  On Error GoTo ERRH
1770                                          End If
1780                                          Select Case fld.Name
                                              Case "admin"
1790                                            rstLoc1.Fields("adminno") = IIf(arr_varAcct(A_ADMIN_N, lngY) = 0&, 1&, arr_varAcct(A_ADMIN_N, lngY))
1800                                          Case "adminno"
1810                                            rstLoc1.Fields("admin") = IIf(arr_varAcct(A_ADMIN_N, lngY) = 0&, 1&, arr_varAcct(A_ADMIN_N, lngY))
1820                                          End Select
1830                                        Else
1840                                          If gblnDev_NoErrHandle = True Then
1850  On Error GoTo 0
1860                                          Else
1870  On Error GoTo ERRH
1880                                          End If
1890                                        End If
1900                                        Exit For
1910                                      End If
1920                                    Next
1930                                  Else
1940  On Error Resume Next
1950                                    rstLoc1.Fields(fld.Name) = Null
1960                                    If ERR.Number <> 0 Then
1970                                      If gblnDev_NoErrHandle = True Then
1980  On Error GoTo 0
1990                                      Else
2000  On Error GoTo ERRH
2010                                      End If
2020                                      Select Case fld.Name
                                          Case "admin"
2030                                        rstLoc1.Fields("adminno") = 1&
2040                                      Case "adminno"
2050                                        rstLoc1.Fields("admin") = 1&
2060                                      End Select
2070                                    Else
2080                                      If gblnDev_NoErrHandle = True Then
2090  On Error GoTo 0
2100                                      Else
2110  On Error GoTo ERRH
2120                                      End If
2130                                    End If
2140                                  End If
2150                                Case "icash", "pcash", "cost"
2160                                  If IsNull(fld.Value) = True Then
2170                                    rstLoc1.Fields(fld.Name) = CCur(0)
2180                                  Else
2190                                    rstLoc1.Fields(fld.Name) = fld.Value
2200                                  End If
2210                                Case "Schedule_ID", "Schedule ID"
                                      ' ** v2.2.00: Schedule_ID
                                      ' ** v2.1.71: Schedule ID
2220                                  If IsNull(fld.Value) = False Then
2230                                    For lngY = 0& To (lngAccts - 1&)
2240                                      If arr_varAcct(A_NUM, lngY) = ![accountno] Then
2250                                        rstLoc1![Schedule_ID] = arr_varAcct(A_SCHED_N, lngY)
2260                                        Exit For
2270                                      End If
2280                                    Next
2290                                  Else
2300                                    rstLoc1.Fields("Schedule_ID") = Null
2310                                  End If
2320                                Case "cotrustee", "amendments", "courtsupervised", "discretion"
2330                                  If IsNull(fld.Value) = True Then
2340                                    rstLoc1.Fields(fld.Name) = "No"
2350                                  Else
2360                                    rstLoc1.Fields(fld.Name) = fld.Value
2370                                  End If
2380                                Case "investmentobj"
2390                                  If IsNull(fld.Value) = False Then
2400                                    strTmp04 = Trim(fld.Value)
2410                                    blnFound = False
2420                                    For lngY = 0& To (lngInvestObjs - 1&)
2430                                      If arr_varInvestObj(IO_NAM, lngY) = strTmp04 Then
2440                                        blnFound = True
2450                                        Exit For
2460                                      End If
2470                                    Next
2480                                    If blnFound = False Then
2490                                      lngInvestObjs = lngInvestObjs + 1&
2500                                      lngE = lngInvestObjs - 1&
2510                                      ReDim Preserve arr_varInvestObj(IO_ELEMS, lngE)
2520                                      arr_varInvestObj(IO_ID, lngE) = CLng(0)
2530                                      arr_varInvestObj(IO_NAM, lngE) = strTmp04
2540                                      arr_varInvestObj(IO_NEW, lngE) = CBool(True)
2550                                      Set rstLoc2 = dbsLoc.OpenRecordset("InvestmentOptions", dbOpenDynaset, dbConsistent)
2560                                      With rstLoc2
2570                                        .AddNew
2580                                        ![invobj_name] = strTmp04
2590                                        ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2600                                        ![DateCreated] = Now()
2610                                        ![DateModified] = Now()
2620                                        .Update
2630                                        .Bookmark = .LastModified
2640                                        arr_varInvestObj(IO_ID, lngE) = ![invobj_id]
2650                                        .Close
2660                                      End With  ' ** rstLoc2.
2670                                    End If
2680                                    rstLoc1.Fields(fld.Name) = strTmp04
2690                                    blnFound = True  ' ** Reset.
2700                                  Else
2710                                    rstLoc1.Fields(fld.Name) = Null
2720                                  End If
2730                                Case "taxlot"  ' ** The default AssetNo, using this field temporarily.
2740                                  For lngY = 0& To (lngAccts - 1&)
2750                                    If arr_varAcct(A_NUM, lngY) = ![accountno] Then
2760                                      rstLoc1![taxlot] = CStr(arr_varAcct(A_DASTNO, lngY))
2770                                      Exit For
2780                                    End If
2790                                  Next
2800                                Case "CaseNum"
                                      ' ** In v2.1.71, at least, I've seen a width of 50 (default),
                                      ' ** and a user used 19 chars, but the width is now 16!  {might change to 20!}
2810                                  If IsNull(fld.Value) = False Then
2820                                    If Len(Trim(fld.Value)) > rstLoc1.Fields(fld.Name).Size Then
2830                                      If strTruncatedFields <> vbNullString Then strTruncatedFields = strTruncatedFields & vbCrLf
2840                                      strTruncatedFields = strTruncatedFields & "Account: " & .Fields("accountno") & "; " & _
                                            "CaseNum truncated: " & Trim(fld.Value)
2850                                      rstLoc1.Fields(fld.Name) = Left(Trim(fld.Value), rstLoc1.Fields(fld.Name).Size)
2860                                    Else
2870                                      rstLoc1.Fields(fld.Name) = Trim(fld.Value)
2880                                    End If
2890                                  End If
2900                                Case "FedIFNum1", "FedIFNum2"
                                      ' ** These were 35 (I think), and now they're 10.
2910                                  If IsNull(fld.Value) = False Then
2920                                    If Len(Trim(fld.Value)) > rstLoc1.Fields(fld.Name).Size Then
2930                                      If strTruncatedFields <> vbNullString Then strTruncatedFields = strTruncatedFields & vbCrLf
2940                                      strTruncatedFields = strTruncatedFields & "Account: " & .Fields("accountno") & "; " & _
                                            fld.Name & " truncated: " & Trim(fld.Value)
2950                                      rstLoc1.Fields(fld.Name) = Left(Trim(fld.Value), rstLoc1.Fields(fld.Name).Size)
2960                                    Else
2970                                      rstLoc1.Fields(fld.Name) = Trim(fld.Value)
2980                                    End If
2990                                  End If
3000                                Case "BankName", "Bank_Name"
                                      ' ** New field name in v2.2.20.
3010                                  If IsNull(fld.Value) = False Then
3020                                    If Trim(fld.Value) <> vbNullString Then
3030                                      rstLoc1.Fields("Bank_Name") = Trim(fld.Value)
3040                                    End If
3050                                  End If
3060                                Case "BankCity", "Bank_City"
                                      ' ** New field name in v2.2.20.
3070                                  If IsNull(fld.Value) = False Then
3080                                    If Trim(fld.Value) <> vbNullString Then
3090                                      rstLoc1.Fields("Bank_City") = Trim(fld.Value)
3100                                    End If
3110                                  End If
3120                                Case "BankState", "Bank_State"
                                      ' ** New field name in v2.2.20.
3130                                  If IsNull(fld.Value) = False Then
3140                                    If Trim(fld.Value) <> vbNullString Then
3150                                      If Len(Trim(fld.Value)) = 2 Then
3160                                        rstLoc1.Fields("Bank_State") = Trim(fld.Value)
3170                                      End If
3180                                    End If
3190                                  End If
3200                                Case "BankAccountNumber", "Bank_AccountNumber"
                                      ' ** New field name in v2.2.20.
3210                                  If IsNull(fld.Value) = False Then
3220                                    If Trim(fld.Value) <> vbNullString Then
3230                                      rstLoc1.Fields("Bank_AccountNumber") = Trim(fld.Value)
3240                                    End If
3250                                  End If
3260                                Case "BankRoutingNumber", "Bank_RoutingNumber"
                                      ' ** New field name in v2.2.20.
3270                                  If IsNull(fld.Value) = False Then
3280                                    If Trim(fld.Value) <> vbNullString Then
3290                                      rstLoc1.Fields("Bank_RoutingNumber") = Trim(fld.Value)
3300                                    End If
3310                                  End If
3320                                Case "Contact1", "Contact1_Name"
                                      ' ** New field name in v2.2.20.
3330                                  If IsNull(fld.Value) = False Then
3340                                    If Trim(fld.Value) <> vbNullString Then
3350                                      rstLoc1.Fields("Contact1_Name") = Trim(fld.Value)
3360                                    End If
3370                                  End If
3380                                Case "add1", "Contact1_Address1"
                                      ' ** New field name in v2.2.20.
3390                                  If IsNull(fld.Value) = False Then
3400                                    If Trim(fld.Value) <> vbNullString Then
3410                                      rstLoc1.Fields("Contact1_Address1") = Trim(fld.Value)
3420                                    End If
3430                                  End If
3440                                Case "add2", "Contact1_Address2"
                                      ' ** New field name in v2.2.20.
3450                                  If IsNull(fld.Value) = False Then
3460                                    If Trim(fld.Value) <> vbNullString Then
3470                                      rstLoc1.Fields("Contact1_Address2") = Trim(fld.Value)
3480                                    End If
3490                                  End If
3500                                Case "city", "Contact1_City"
                                      ' ** New field name in v2.2.20.
3510                                  If IsNull(fld.Value) = False Then
3520                                    If Trim(fld.Value) <> vbNullString Then
3530                                      rstLoc1.Fields("Contact1_City") = Trim(fld.Value)
3540                                    End If
3550                                  End If
3560                                Case "state", "Contact1_State"
                                      ' ** New field name in v2.2.20.
3570                                  If IsNull(fld.Value) = False Then
3580                                    If Trim(fld.Value) <> vbNullString Then
3590                                      rstLoc1.Fields("Contact1_State") = Trim(fld.Value)
3600                                    End If
3610                                  End If
3620                                Case "zip", "Contact1_Zip"
                                      ' ** New field name in v2.2.20.
3630                                  If IsNull(fld.Value) = False Then
3640                                    If Trim(fld.Value) <> vbNullString Then
3650                                      rstLoc1.Fields("Contact1_Zip") = Trim(fld.Value)
3660                                    End If
3670                                  End If
3680                                Case "phone", "Contact1_Phone1"
                                      ' ** New field name in v2.2.20.
3690                                  If IsNull(fld.Value) = False Then
3700                                    If Trim(fld.Value) <> vbNullString Then
3710                                      rstLoc1.Fields("Contact1_Phone1") = Trim(fld.Value)
3720                                    End If
3730                                  End If
3740                                Case "OtherPhone", "Contact1_Phone2"
                                      ' ** New field name in v2.2.20.
3750                                  If IsNull(fld.Value) = False Then
3760                                    If Trim(fld.Value) <> vbNullString Then
3770                                      rstLoc1.Fields("Contact1_Phone2") = Trim(fld.Value)
3780                                    End If
3790                                  End If
3800                                Case "fax", "Contact1_Fax"
                                      ' ** New field name in v2.2.20.
3810                                  If IsNull(fld.Value) = False Then
3820                                    If Trim(fld.Value) <> vbNullString Then
3830                                      rstLoc1.Fields("Contact1_Fax") = Trim(fld.Value)
3840                                    End If
3850                                  End If
3860                                Case "email", "Contact1_Email"
                                      ' ** New field name in v2.2.20.
3870                                  If IsNull(fld.Value) = False Then
3880                                    If Trim(fld.Value) <> vbNullString Then
3890                                      rstLoc1.Fields("Contact1_Email") = Trim(fld.Value)
3900                                    End If
3910                                  End If
3920                                Case "Contact2", "Contact2_Name"
                                      ' ** New field name in v2.2.20.
3930                                  If IsNull(fld.Value) = False Then
3940                                    If Trim(fld.Value) <> vbNullString Then
3950                                      rstLoc1.Fields("Contact2_Name") = Trim(fld.Value)
3960                                    End If
3970                                  End If
3980                                Case "add12", "Contact2_Address1"
                                      ' ** New field name in v2.2.20.
3990                                  If IsNull(fld.Value) = False Then
4000                                    If Trim(fld.Value) <> vbNullString Then
4010                                      rstLoc1.Fields("Contact2_Address1") = Trim(fld.Value)
4020                                    End If
4030                                  End If
4040                                Case "add22", "Contact2_Address2"
                                      ' ** New field name in v2.2.20.
4050                                  If IsNull(fld.Value) = False Then
4060                                    If Trim(fld.Value) <> vbNullString Then
4070                                      rstLoc1.Fields("Contact2_Address2") = Trim(fld.Value)
4080                                    End If
4090                                  End If
4100                                Case "city2", "Contact2_City"
                                      ' ** New field name in v2.2.20.
4110                                  If IsNull(fld.Value) = False Then
4120                                    If Trim(fld.Value) <> vbNullString Then
4130                                      rstLoc1.Fields("Contact2_City") = Trim(fld.Value)
4140                                    End If
4150                                  End If
4160                                Case "state2", "Contact2_State"
                                      ' ** New field name in v2.2.20.
4170                                  If IsNull(fld.Value) = False Then
4180                                    If Trim(fld.Value) <> vbNullString Then
4190                                      rstLoc1.Fields("Contact2_State") = Trim(fld.Value)
4200                                    End If
4210                                  End If
4220                                Case "zip2", "Contact2_Zip"
                                      ' ** New field name in v2.2.20.
4230                                  If IsNull(fld.Value) = False Then
4240                                    If Trim(fld.Value) <> vbNullString Then
4250                                      rstLoc1.Fields("Contact2_Zip") = Trim(fld.Value)
4260                                    End If
4270                                  End If
4280                                Case "phone2", "Contact2_Phone1"
                                      ' ** New field name in v2.2.20.
4290                                  If IsNull(fld.Value) = False Then
4300                                    If Trim(fld.Value) <> vbNullString Then
4310                                      rstLoc1.Fields("Contact2_Phone1") = Trim(fld.Value)
4320                                    End If
4330                                  End If
4340                                Case "OtherPhone2", "Contact2_Phone2"
                                      ' ** New field name in v2.2.20.
4350                                  If IsNull(fld.Value) = False Then
4360                                    If Trim(fld.Value) <> vbNullString Then
4370                                      rstLoc1.Fields("Contact2_Phone2") = Trim(fld.Value)
4380                                    End If
4390                                  End If
4400                                Case "fax2", "Contact2_Fax"
                                      ' ** New field name in v2.2.20.
4410                                  If IsNull(fld.Value) = False Then
4420                                    If Trim(fld.Value) <> vbNullString Then
4430                                      rstLoc1.Fields("Contact2_Fax") = Trim(fld.Value)
4440                                    End If
4450                                  End If
4460                                Case "email2", "Contact2_Email"
                                      ' ** New field name in v2.2.20.
4470                                  If IsNull(fld.Value) = False Then
4480                                    If Trim(fld.Value) <> vbNullString Then
4490                                      rstLoc1.Fields("Contact2_Email") = Trim(fld.Value)
4500                                    End If
4510                                  End If
4520                                Case "curr_id"
                                      ' ** If found, copy it.
4530                                  blnTmp23 = True
4540                                  rstLoc1![curr_id] = ![curr_id]
4550                                Case Else
4560                                  rstLoc1.Fields(fld.Name) = fld.Value
4570                                End Select
4580                                If blnTmp23 = False Then  ' ** curr_id not present.
4590                                  rstLoc1![curr_id] = 150&  ' ** Default to USD.
4600                                End If
4610                              Next  ' ** For each Field in Fields: fld.
4620                              If lngFlds <> 95& Then
                                    ' ** Missing fields in various of the 16 previous documented versions.
                                    ' **   Field                 Versions Missing It
                                    ' **   ====================  ===================
                                    ' **   Bank_City             5
                                    ' **   Bank_RoutingNumber    5
                                    ' **   Bank_State            5
                                    ' **   CaseNum               5
                                    ' **   FedIFNum1             6
                                    ' **   FedIfNum2             6
                                    ' **   LastCheckNum          5
                                    ' **   curr_id               All prior to 2.2.?
                                    ' **   Bank_Country
                                    ' **   Contact1_Country
                                    ' **   Contact1_PostalCode
                                    ' **   Contact2_Country
                                    ' **   Contact2_PostalCode
                                    ' ** None of these fields are required, so just let them be Null, with curr_id defaulting to 150.
4630                              End If
4640  On Error Resume Next
4650                              rstLoc1.Update
4660                              If ERR.Number <> 0 Then
4670                                If ERR.Number = 3022 Then
                                      ' ** Error 3022: The changes you requested to the table were not successful because they
                                      ' **             would create duplicate values in the index, primary key, or relationship.
4680                                  If gblnDev_NoErrHandle = True Then
4690  On Error GoTo 0
4700                                  Else
4710  On Error GoTo ERRH
4720                                  End If
                                      ' ** The Primary Key, [accountno], is the only field with a unique index.
                                      ' ** Though dupes shouldn't happen, if they do, just drop the dupe.
4730                                  For lngY = 0& To (lngAccts - 1&)
4740                                    If arr_varAcct(A_NUM, lngY) = ![accountno] Then
4750                                      If arr_varAcct(A_NUM_N, lngY) = vbNullString Then
4760                                        arr_varAcct(A_NUM_N, lngY) = "#DUPE: ~" & CStr(lngY) & "^" & arr_varAcct(A_NAM, lngY) & "^~"
4770                                        arr_varAcct(A_DROPPED, lngY) = CBool(True)
4780                                      Else
4790                                        arr_varAcct(A_NUM_N, lngY) = arr_varAcct(A_NUM_N, lngY) & _
                                              CStr(lngY) & "^" & arr_varAcct(A_NAM, lngY) & "^~"
4800                                        arr_varAcct(A_DROPPED, lngY) = CBool(True)
4810                                      End If
4820                                    End If
4830                                  Next
4840                                  rstLoc1.CancelUpdate
4850                                Else
4860                                  intRetVal = -6
4870                                  blnContinue = False
4880                                  lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
4890                                  MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
                                        "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
                                        vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4900                                  rstLoc1.CancelUpdate
4910                                  If gblnDev_NoErrHandle = True Then
4920  On Error GoTo 0
4930                                  Else
4940  On Error GoTo ERRH
4950                                  End If
4960                                End If
4970                              Else
4980                                If gblnDev_NoErrHandle = True Then
4990  On Error GoTo 0
5000                                Else
5010  On Error GoTo ERRH
5020                                End If
5030                              End If
5040                              If blnContinue = True Then
                                    ' ** The key field doesn't change, so no need to put it in tblVersion_Key.
5050                              Else
5060                                Exit For
5070                              End If
5080                            End If
5090                            If lngX < lngRecs Then .MoveNext
5100                          Next
5110                          rstLoc1.Close
5120                          rstLoc2.Close
5130                        End If  ' ** Records present.
5140                        .Close
5150                      End With  ' ** rstLnk.
5160                    End If  ' ** blnFound.

5170                  End If  ' ** blnContinue.

5180                  If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ******************************
                        ' ** Table: Balance.
                        ' ******************************

                        ' ** Step 13: Balance.
5190                    dblPB_ThisStep = 13#
5200                    Version_Status 3, dblPB_ThisStep, "Balance"  ' ** Function: Below.

5210                    strCurrTblName = "Balance"
5220                    lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

5230                    blnFound = False: lngRecs = 0&
5240                    For lngX = 0& To (lngOldTbls - 1&)
5250                      If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
5260                        blnFound = True
5270                        Exit For
5280                      End If
5290                    Next

5300                    If blnFound = True Then
5310                      Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
5320                      With rstLnk
5330                        If .BOF = True And .EOF = True Then
                              ' ** This really really should have records!
5340                        Else
5350                          strCurrKeyFldName = "accountno"
5360                          lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
5370                          Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
5380                          Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Of 16 example TrustDta.mdb's, all have 7 fields.
                              ' ** Current field count is 8 fields.
                              ' ** Table: Balance
                              ' **   ![accountno]
                              ' **   ![balance date]
                              ' **   ![icash]
                              ' **   ![pcash]
                              ' **   ![cost]
                              ' **   ![TotalMarketValue]
                              ' **   ![AccountValue]
                              ' **   ![curr_id}  Defaults to 150.
                              ' ** No tables reference this directly.
5390                          .MoveLast
5400                          lngRecs = .RecordCount
5410                          Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Function: Below.
5420                          .MoveFirst
5430                          lngFlds = .Fields.Count
5440                          For lngX = 1& To lngRecs
5450                            Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Function: Below.
                                ' ** Check to make sure it's an existing account.
5460                            strTmp04 = ![accountno]
5470                            blnFound = False
5480                            For lngY = 0& To (lngAccts - 1&)
5490                              If arr_varAcct(A_NUM, lngY) = strTmp04 Then
5500                                blnFound = True
5510                                Exit For
5520                              End If
5530                            Next
5540                            If blnFound = False Then
                                  ' ** See if it's a variation related to an Account Type prefix.
5550                              If InStr(strTmp04, A99_INC) > 0 Then
5560                                If Left(strAcct99_IncomeOU, 3) = "99-" And Left(strTmp04, 3) <> "99-" Then
5570                                  strTmp04 = "99-" & strTmp04
5580                                  blnFound = True
5590                                ElseIf Left(strAcct99_IncomeOU, 3) <> "99-" And Left(strTmp04, 3) = "99-" Then
5600                                  strTmp04 = Mid(strTmp04, 4)
5610                                  blnFound = True
5620                                End If
5630                              ElseIf InStr(strTmp04, A99_SUS) > 0 Then
5640                                If Left(strAcct99_Suspense, 3) = "99-" And Left(strTmp04, 3) <> "99-" Then
5650                                  strTmp04 = "99-" & strTmp04
5660                                  blnFound = True
5670                                ElseIf Left(strAcct99_Suspense, 3) <> "99-" And Left(strTmp04, 3) = "99-" Then
5680                                  strTmp04 = Mid(strTmp04, 4)
5690                                  blnFound = True
5700                                End If
5710                              Else
                                    ' ** Check to see if it's the other one in case they changed
                                    ' ** gblnAccountNoWithType at some point in the past.
5720                                If Mid(strTmp04, 3, 1) = "-" Then
                                      ' ** This might be a number with an Account Type prefix.
5730                                  For lngY = 0& To (lngAcctTypes - 1&)
5740                                    If arr_varAcctType(AT_TYP, lngY) = Left(strTmp04, 2) Then
                                          ' ** The prefix matches an Account Type.
5750                                      For lngZ = 0& To (lngAccts - 1&)
5760                                        If arr_varAcct(A_NUM, lngZ) = Mid(strTmp04, 4) Then
                                              ' ** And the rest of it matches an existing account.
5770                                          blnFound = True
5780                                          strTmp04 = Mid(strTmp04, 4)
5790                                          Exit For
5800                                        End If
5810                                      Next
5820                                      Exit For
5830                                    End If
5840                                  Next
5850                                  If blnFound = False Then
                                        ' ** A dead account; just delete it.
5860                                    strTmp04 = vbNullString
5870                                  End If
5880                                Else
                                      ' ** Check the other way around.
5890                                  For lngY = 0& To (lngAccts - 1&)
5900                                    If Mid(arr_varAcct(A_NUM, lngY), 3, 1) = "-" Then
                                          ' ** An existing number might have an Account Type prefix.
5910                                      For lngZ = 0& To (lngAcctTypes - 1&)
5920                                        If arr_varAcctType(AT_TYP, lngZ) = Left(arr_varAcct(A_NUM, lngY), 2) Then
                                              ' ** That prefix matches an Account Type.
5930                                          If strTmp04 = Mid(arr_varAcct(A_NUM, lngY), 4) Then
                                                ' ** And the rest of it matches our Balance record.
5940                                            blnFound = True
5950                                            strTmp04 = arr_varAcct(A_NUM, lngY)
5960                                            Exit For
5970                                          End If
5980                                        End If
5990                                      Next
6000                                    End If
6010                                    If blnFound = True Then Exit For
6020                                  Next
6030                                  If blnFound = False Then
                                        ' ** A dead account; just delete it.
6040                                    strTmp04 = vbNullString
6050                                  End If
6060                                End If
6070                              End If
6080                            End If
6090                            If blnFound = True Then
                                  '08/19/2009: ERRORED HERE ON OHANA DATA BECAUSE ACCOUNTNO WAS '99-INCOME O/U'/'99-SUSPENSE'!!!!
6100                              If strTmp04 = "99-INCOME O/U" Or strTmp04 = "99-SUSPENSE" Then
6110                                strTmp04 = Mid(strTmp04, 4)
6120                              End If
                                  ' ** Add the record to the new table.
6130                              rstLoc1.AddNew
6140                              rstLoc1![accountno] = strTmp04
6150                              rstLoc1![balance date] = ![balance date]
6160                              rstLoc1![ICash] = Nz(![ICash], 0)
6170                              rstLoc1![PCash] = Nz(![PCash], 0)
6180                              rstLoc1![Cost] = Nz(![Cost], 0)
6190                              rstLoc1![TotalMarketValue] = Nz(![TotalMarketValue], 0)
6200                              rstLoc1![AccountValue] = Nz(![AccountValue], 0)
                                  'ADD CURR_ID!
6210                              If lngFlds = 8& Then  ' ** The 8th field is curr_id.
6220  On Error Resume Next
6230                                rstLoc1![curr_id] = ![curr_id]
6240                                If ERR.Number <> 0 Then
6250  On Error GoTo ERRH
6260                                  rstLoc1![curr_id] = 150&  ' ** Default to USD.
6270                                Else
6280  On Error GoTo ERRH
6290                                End If
6300                                If IsNull(rstLoc1![curr_id]) = True Then
6310                                  rstLoc1![curr_id] = 150&  ' ** Default to USD.
6320                                Else
6330                                  If rstLoc1![curr_id] = 0& Then
6340                                    rstLoc1![curr_id] = 150&  ' ** Default to USD.
6350                                  End If
6360                                End If
6370                              Else
6380                                rstLoc1![curr_id] = 150&  ' ** Default to USD.
6390                              End If
6400  On Error Resume Next
6410                              rstLoc1.Update
6420                              If ERR.Number <> 0 Then
6430                                If ERR.Number = 3022 Then
                                      ' ** Error 3022: The changes you requested to the table were not successful because they
                                      ' **             would create duplicate values in the index, primary key, or relationship.
6440                                  If gblnDev_NoErrHandle = True Then
6450  On Error GoTo 0
6460                                  Else
6470  On Error GoTo ERRH
6480                                  End If
6490                                  rstLoc1.CancelUpdate
                                      ' ** If there are 2 records with the same date and accountno,
                                      ' ** make sure we keep the one with values (if either of them do).
                                      ' ** If the one already entered has values and this one doesn't,
                                      ' ** throw this one away. If this has values and the other
                                      ' ** doesn't, edit that one with these values.
6500                                  If ![ICash] <> 0 Or ![PCash] <> 0 Or ![Cost] <> 0 Or ![TotalMarketValue] <> 0 Or ![AccountValue] <> 0 Then
6510                                    rstLoc1.MoveFirst
6520                                    rstLoc1.FindFirst "[accountno] = '" & strTmp04 & "' And " & _
                                          "[balance date] = #" & Format(![balance date], "mm/dd/yyyy") & "#"
6530                                    If rstLoc1.NoMatch = False Then
6540                                      If rstLoc1![ICash] = 0 And rstLoc1![PCash] = 0 And rstLoc1![Cost] = 0 And _
                                              rstLoc1![TotalMarketValue] = 0 And rstLoc1![AccountValue] = 0 Then
6550                                        rstLoc1.Edit
6560                                        rstLoc1![ICash] = Nz(![ICash], 0)
6570                                        rstLoc1![PCash] = Nz(![PCash], 0)
6580                                        rstLoc1![Cost] = Nz(![Cost], 0)
6590                                        rstLoc1![TotalMarketValue] = Nz(![TotalMarketValue], 0)
6600                                        rstLoc1![AccountValue] = Nz(![AccountValue], 0)
6610                                        rstLoc1.Update
6620                                      Else
                                            ' ** Throw this one away.
6630                                      End If
6640                                    Else
                                          ' ** Oh well. Just throw it away.
6650                                    End If
6660                                  Else
                                        ' ** Throw this one away.
6670                                  End If
6680                                Else
6690                                  intRetVal = -6
6700                                  blnContinue = False
6710                                  lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
6720                                  MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
                                        "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
                                        vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
6730                                  rstLoc1.CancelUpdate
6740                                  If gblnDev_NoErrHandle = True Then
6750  On Error GoTo 0
6760                                  Else
6770  On Error GoTo ERRH
6780                                  End If
6790                                End If
6800                              Else
6810                                If gblnDev_NoErrHandle = True Then
6820  On Error GoTo 0
6830                                Else
6840  On Error GoTo ERRH
6850                                End If
6860                              End If
6870                              If blnContinue = True Then
                                    ' ** The key field doesn't change, so no need to put it in tblVersion_Key.
6880                              Else
6890                                Exit For
6900                              End If
6910                            End If
6920                            If lngX < lngRecs Then .MoveNext
6930                          Next
6940                          rstLoc1.Close
6950                          rstLoc2.Close
6960                        End If  ' ** Records present.
6970                        .Close
6980                      End With  ' ** rstLnk.
6990                    End If  ' ** blnFound.

                        ' ** Balance, grouped, with Min(balance date); earliest with non-zero values.
7000                    varTmp00 = DLookup("[Balance_Date]", "qryStatementParameters_15")
7010                    lngStats = lngStats + 1&
7020                    lngE = lngStats - 1&
7030                    ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
7040                    arr_varStat(STAT_ORD, lngE) = CInt(8)
7050                    arr_varStat(STAT_NAM, lngE) = "Statement Dates: "
7060                    If IsNull(varTmp00) = False Then
7070                      arr_varStat(STAT_CNT, lngE) = CLng(varTmp00)
7080                    Else
7090                      arr_varStat(STAT_CNT, lngE) = CLng(0)
7100                    End If
7110                    arr_varStat(STAT_DSC, lngE) = "Earliest: "

                        ' ** Balance, grouped, with Max(balance date); latest with non-zero values.
7120                    varTmp00 = DLookup("[Balance_Date]", "qryStatementParameters_16")
7130                    lngStats = lngStats + 1&
7140                    lngE = lngStats - 1&
7150                    ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
7160                    arr_varStat(STAT_ORD, lngE) = CInt(9)
7170                    arr_varStat(STAT_NAM, lngE) = "Statement Dates: "
7180                    If IsNull(varTmp00) = False Then
7190                      arr_varStat(STAT_CNT, lngE) = CLng(varTmp00)
7200                    Else
7210                      arr_varStat(STAT_CNT, lngE) = CLng(0)
7220                    End If
7230                    arr_varStat(STAT_DSC, lngE) = "Latest: "

7240                  End If  ' ** blnContinue.

7250                  If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ** Get a list of Asset Types.
7260                    Set qdf = dbsLoc.QueryDefs("qryAssetType_03")
7270                    Set rstLoc1 = qdf.OpenRecordset
7280                    With rstLoc1
7290                      .MoveLast
7300                      lngAssetTypes = .RecordCount
7310                      .MoveFirst
7320                      arr_varAssetType = .GetRows(lngAssetTypes)
                          ' ************************************************
                          ' ** Array: arr_varAssetType()
                          ' **
                          ' **   Field  Element  Name           Constant
                          ' **   =====  =======  =============  ==========
                          ' **     1       0     assettype      AT_TYP
                          ' **     2       1     description    AT_DSC
                          ' **
                          ' ************************************************
                          ' ** Layout the same as arr_varAcctType().
7330                      .Close
7340                    End With  ' ** rstLoc1.

                        ' ******************************
                        ' ** Table: masterasset.
                        ' ******************************

                        ' ** Step 14: masterasset.
7350                    dblPB_ThisStep = 14#
7360                    Version_Status 3, dblPB_ThisStep, "Master Asset"  ' ** Function: Below.

7370                    strCurrTblName = "masterasset"
7380                    lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

7390                    gdblCrtRpt_CostTot = 0#  ' ** Borrowing this for the sweep MarketValueCurrent.
7400                    gstrCrtRpt_Version = vbNullString  ' ** Borrowing this for the sweep asset description.

7410                    blnFound = False: lngRecs = 0&
7420                    For lngX = 0& To (lngOldTbls - 1&)
7430                      If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
7440                        blnFound = True
7450                        Exit For
7460                      End If
7470                    Next

7480                    If blnFound = True Then

7490                      lngMasterAssets = 0&
7500                      ReDim arr_varMasterAsset(MA_ELEMS, 0)

7510                      lngTmp13 = 0&
7520                      ReDim arr_varTmp01(3, 0)

                          ' ** Get a list of MasterAssets with no Ledger activity whatsoever.
7530                      strTmp04 = "SELECT masterasset.assetno, masterasset.description " & _
                            "FROM masterasset LEFT JOIN ledger ON masterasset.assetno = ledger.assetno " & _
                            "WHERE (((ledger.assetno) Is Null));"
                          ' ** Also check LedgerArchive.
7540                      blnTmp22 = False
7550                      For lngX = 0& To (lngOldFiles - 1&)
7560                        If arr_varOldFile(F_FNAM, lngX) = gstrFile_ArchDataName Then
7570                          Set dbs = wrkLnk.OpenDatabase(arr_varOldFile(F_PTHFIL, lngArchElem), False, True)  ' ** {pathfile}, {exclusive}, {read-only}
7580                          Exit For
7590                        End If
7600                      Next  ' ** lngX.
7610                      Set rstLoc3 = dbs.OpenRecordset("ledger", dbOpenDynaset, dbReadOnly)
7620                      If rstLoc3.BOF = True And rstLoc3.EOF = True Then
                            ' ** No archive.
7630                      Else
7640                        rstLoc3.MoveFirst
7650                        blnTmp22 = True
7660                      End If
7670                      Set qdf = dbsLnk.CreateQueryDef("", strTmp04)
7680                      Set rstLnk = qdf.OpenRecordset
7690                      With rstLnk
7700                        If .BOF = True And .EOF = True Then
                              ' ** All assets are in use.
7710                        Else
7720                          .MoveLast
7730                          lngRecs = .RecordCount
7740                          .MoveFirst
7750                          For lngX = 1& To lngRecs
7760                            blnFound2 = False
7770                            If blnTmp22 = True Then
7780                              rstLoc3.FindFirst "[assetno] = " & CStr(![assetno])
7790                              If rstLoc3.NoMatch = False Then
7800                                blnFound2 = True  ' ** Found in LedgerArchive.
7810                              End If
7820                            End If
7830                            If blnFound2 = False Then
7840                              lngTmp13 = lngTmp13 + 1&
7850                              lngE = lngTmp13 - 1&
7860                              ReDim Preserve arr_varTmp01(3, lngE)
7870                              arr_varTmp01(0, lngE) = ![assetno]
7880                              arr_varTmp01(1, lngE) = ![description]
7890                              arr_varTmp01(2, lngE) = CBool(True)   ' ** Default to Yes, this one's unused.
7900                              arr_varTmp01(3, lngE) = CBool(False)  ' ** Default to No, don't skip them.
7910                            End If  ' ** blnFound2.
7920                            If lngX < lngRecs Then .MoveNext
7930                          Next
7940                        End If
7950                        .Close
7960                      End With
7970                      rstLoc3.Close
7980                      Set rstLoc3 = Nothing
7990                      dbs.Close
8000                      Set dbs = Nothing
8010                      blnFound2 = False: blnTmp22 = False
8020                      DoEvents

8030                      Set rstLnk = Nothing
8040                      Set qdf = Nothing
8050                      lngRecs = 0&

8060                      If lngTmp13 > 0& Then

                            ' ** Collect the assetno's into a comma-delimited IN() string.
8070                        strTmp05 = vbNullString
8080                        For lngX = 0& To (lngTmp13 - 1&)
8090                          If lngX = 0& Then
8100                            strTmp05 = CStr(arr_varTmp01(0, lngX))
8110                          Else
8120                            strTmp05 = strTmp05 & "," & CStr(arr_varTmp01(0, lngX))
8130                          End If
8140                        Next

                            ' ** Now cross-check that list with ActiveAssets.
8150                        strTmp04 = "SELECT ActiveAssets.assetno, ActiveAssets.accountno, ActiveAssets.assetdate " & _
                              "FROM ActiveAssets " & _
                              "WHERE (((ActiveAssets.assetno) In (" & strTmp05 & ")));"
8160                        Set qdf = dbsLnk.CreateQueryDef("", strTmp04)
8170                        Set rstLnk = qdf.OpenRecordset
8180                        With rstLnk
8190                          If .BOF = True And .EOF = True Then
                                ' ** No orphan assets.
8200                          Else
                                ' ** We'll have to wait for the ActiveAssets table processing
                                ' ** to deal with these. For now, just remove them from the list.
8210                            .MoveLast
8220                            lngRecs = .RecordCount
8230                            .MoveFirst
8240                            For lngX = 1& To lngRecs
8250                              For lngY = 0& To (lngTmp13 - 1&)
8260                                If arr_varTmp01(0, lngY) = ![assetno] Then
8270                                  arr_varTmp01(2, lngY) = CBool(False)  ' ** No, this isn't unused!
8280                                  Exit For
8290                                End If
8300                              Next
8310                              If lngX < lngRecs Then .MoveNext
8320                            Next
8330                          End If
8340                          .Close
8350                        End With

                            ' ** Next, check their names for possible intentional non-use.
                            ' ** (Since Trust Accountant does not permit the deletion of
                            ' ** assets from the MasterAssets table, users sometimes just
                            ' ** sort of blank them out, for possible re-use in the future.)
8360                        For lngX = 0& To (lngTmp13 - 1&)
8370                          If arr_varTmp01(2, lngX) = True Then
                                ' ** Yes, it's unused.
8380                            strTmp04 = Trim(arr_varTmp01(1, lngX))
8390                            If strTmp04 <> "Accrued Interest Asset" Then
8400                              If (Left(strTmp04, 3) = "123") Or (Left(strTmp04, 3) = "xyz") Or _
                                      (Left(strTmp04, 5) = "Blank") Or (Left(strTmp04, 5) = "Error") Or _
                                      (Left(strTmp04, 5) = "Empty") Or (Left(strTmp04, 6) = "Delete") Or _
                                      (Left(strTmp04, 6) = "#Error") Or (Left(strTmp04, 7) = "Nothing") Or _
                                      (Left(strTmp04, 8) = "Not Used") Or (Left(strTmp04, 9) = "Available") Or _
                                      (Left(strTmp04, 10) = "Do Not Use") Then
                                    ' ** Remember, this is only applied to unused assets.
8410                                arr_varTmp01(3, lngX) = CBool(True)  ' ** Yes, skip this one.
8420                              End If
8430                            End If
8440                          End If
8450                        Next

8460                        Set rstLnk = Nothing
8470                        Set qdf = Nothing
8480                        strTmp04 = vbNullString: strTmp05 = vbNullString

8490                      End If  ' ** lngTmp13.  WHICH WE'RE STILL USING!!!

                          'OK, how do I want to do this?
                          'Since the tables aren't linked, I can't use a simple query.
                          'Should I write a query in dbsLnk?
                          'NO, I'd rather not disturb that database at all.
                          'However, I can create and run a query with no name
                          'that's automatically not saved -- for use on-the-fly.
                          'So, the first thing would be to find out if there
                          'are MasterAssets with no links to the Ledger or ActiveAssets.
                          'If results are zero, do nothing.
                          'If I do get hits, then check their names against the above list.
                          'Hits to the list get skipped.

8500                      lngRecs = 0&
8510                      Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
8520                      With rstLnk
8530                        If .BOF = True And .EOF = True Then
                              ' ** This really really should have records!
8540                        Else
8550                          strCurrKeyFldName = "assetno"
8560                          lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
8570                          Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
8580                          Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Of 21 example TrustDta.mdb's, all but 1 have 12 fields; v1.6.3 has 11.
                              ' ** Current field count is 13 fields.
                              ' ** Table: masterasset
                              ' **   ![assetno]            dbLong
                              ' **     Tables with mixed assetno data types:
                              ' **       ActiveAssets      dbLong
                              ' **       ActiveAssets        dbDouble
                              ' **       asset             dbLong
                              ' **       asset               dbDouble
                              ' **       journal           dbLong
                              ' **       journal             dbDouble
                              ' **       journal map       dbLong
                              ' **       journal map         dbDouble
                              ' **       ledger            dbLong
                              ' **       ledger              dbDouble
                              ' **   ![cusip]              dbText (unique)
                              ' **   ![description]        dbText 50
                              ' **   ![shareface]          dbDouble
                              ' **   ![assettype]          dbText 2
                              ' **   ![rate]               dbDouble
                              ' **   ![due]                dbDate
                              ' **   ![marketvalue]        dbDouble
                              ' **   ![marketvaluecurrent] dbDouble
                              ' **   ![yield]              dbDouble
                              ' **   ![currentDate]        dbDate
                              ' **   ![masterasset_TYPE]   dbText  MISSING IN v1.6.3!
                              ' **     RA = Regular Asset
                              ' **     IA = Interest Asset, cusip: 999999999  There should be only 1!
                              ' **     CONFUSION WITH SWEEP ASSET IN DEMO!
                              ' **   ![curr_id]  Defaults to 150.
                              ' ** Referenced by:
                              ' **   Table: ActiveAssets                  {Linked}
                              ' **     Field: [assetno]
                              ' **   Table: asset                         {Linked}
                              ' **     Field: [assetno]
                              ' **   Table: journal                       {Linked}
                              ' **     Field: [assetno]
                              ' **   Table: Journal Map                   {Linked}
                              ' **     Field: [assetno]
                              ' **   Table: ledger                        {Linked}
                              ' **     Field: [assetno]
                              ' **   Table: LedgerArchive                 {Linked}
                              ' **     Field: [assetno]
                              ' **   Table: masterasset                   {Linked}
                              ' **     Field: [assetno]
                              ' **   Table: tblAssetPricing               {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tblTemplate_ActiveAssets      {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tblTemplate_Asset             {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tblTemplate_Journal           {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tblTemplate_Ledger            {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpAccountInfo                {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpAssetList2                 {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpAssetList3                 {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpAveragePrice               {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpCourtReportData2           {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpEdit10                       {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpEdit02                      {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpEstateVal2                 {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpPricingUpdatedCusips       {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpTaxReports                 {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpTrx1                       {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpTrx2                       {Local}
                              ' **     Field: [assetno]
                              ' **   Table: tmpTrx3                       {Local}
                              ' **     Field: [assetno]
8590                          .MoveLast
8600                          lngRecs = .RecordCount
8610                          Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Function: Below.
8620                          .MoveFirst
8630                          lngFlds = 0&
8640                          For lngX = 0& To (lngOldFiles - 1&)
8650                            If arr_varOldFile(F_FNAM, lngX) = gstrFile_DataName Then
8660                              arr_varTmp02 = arr_varOldFile(F_T_ARR, lngX)
8670                              lngOldTbls = (UBound(arr_varTmp02, 2) + 1)
8680                              For lngY = 0& To (lngOldTbls - 1&)
8690                                If arr_varTmp02(T_TNAM, lngY) = strCurrTblName Then
8700                                  lngFlds = arr_varTmp02(T_FLDS, lngY)
8710                                  Exit For
8720                                End If
8730                              Next
8740                              Exit For
8750                            End If
8760                          Next
8770                          For lngX = 1& To lngRecs
8780                            Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Function: Below.
8790                            blnTmp22 = False: blnTmp23 = False: blnTmp24 = False
                                ' ** blnTmp22 is used twice:
                                ' **   1st, whether to skip it; if True, the other blnTmp's don't matter.
                                ' **   2nd, if False, reused as possible 'Accrued Interest Asset'.
                                ' ** blnTmp23 indicates not a perfect match for 'Accrued Interest Asset'.
                                ' ** blnTmp24 indicates masterasset_TYPE 'IA'.
8800                            For lngY = 0& To (lngTmp13 - 1&)
8810                              If arr_varTmp01(0, lngY) = ![assetno] Then
8820                                blnTmp22 = arr_varTmp01(3, lngY)  ' ** Should this be skipped: True/False.
8830                                Exit For
8840                              End If
8850                            Next
8860                            If blnTmp22 = False Then  ' ** No, it shouldn't be skipped.
8870                              lngMasterAssets = lngMasterAssets + 1&
8880                              lngE = lngMasterAssets - 1&
8890                              ReDim Preserve arr_varMasterAsset(MA_ELEMS, lngE)
8900                              arr_varMasterAsset(MA_OLD_ANO, lngE) = CLng(![assetno])
8910                              arr_varMasterAsset(MA_NEW_ANO, lngE) = CLng(0)
8920                              arr_varMasterAsset(MA_NAM, lngE) = ![description]
8930                              arr_varMasterAsset(MA_OLD_MVC, lngE) = CDbl(Nz(![marketvaluecurrent], 0))
8940                              arr_varMasterAsset(MA_NEW_MVC, lngE) = CDbl(0)
8950                              arr_varMasterAsset(MA_ERR, lngE) = CBool(False)
8960                              arr_varMasterAsset(MA_ERRDESC, lngE) = vbNullString
8970                              If lngFlds = 12& Or lngFlds = 13& Then  ' ** The 13th field is curr_id.
8980                                If ![masterasset_TYPE] = "IA" Then
8990                                  blnTmp24 = True
9000                                End If
9010                              Else
                                    ' ** An older 1.6.x version, without [masterasset_TYPE].
9020                              End If
9030                              If ![assetno] = 1& Or ![cusip] = "999999999" Or _
                                      ![description] = "Accrued Interest Asset" Or blnTmp24 = True Then
                                    ' ** This is the only one that's distributed with a new install.
9040                                blnTmp22 = True  ' ** It's a possible Accrued.
9050                                If lngFlds <> 12& And lngFlds <> 13& Then blnTmp24 = True   ' ** Let v1.6.3 pass this next test for that field.
9060                                If ![assetno] <> 1& Or ![cusip] <> "999999999" Or _
                                        ![description] <> "Accrued Interest Asset" Or blnTmp24 <> True Then
                                      ' ** This might have shown up elsewhere in a previous
                                      ' ** version, or it may not have been present at all.
                                      ' ** Regardless, whatever's here now has to be moved!
9070                                  blnTmp23 = True  ' ** But it's not a perfect Accrued.
9080                                End If
                                    ' ** These are used below.
9090                                gdblCrtRpt_CostTot = ![marketvaluecurrent]  ' ** Borrowing this for the sweep MarketValueCurrent.
9100                                gstrCrtRpt_Version = ![description]  ' ** Borrowing this for the sweep asset description.
9110                              End If
9120                              If blnTmp22 = False Or (blnTmp22 = True And blnTmp23 = True) Then
                                    ' ** If {it shouldn't be skipped} Or ({it's a possible accrued} And {that's not a perfect accrued}) Then ...
                                    ' ** {Note: blnTmp23 would never be True if this entry were truely marked to be skipped}
                                    ' ** An Accrued that matches all 4 criteria WILL BE SKIPPED here,
                                    ' ** because it's identical to the record already in the MasterAsset table!
                                    ' ** (The good, identical Accrued will have blnTmp22 = True, as set above
                                    ' ** at Line 26110, but it's blnTmp23 = False because it has no imperfections.)
                                    ' ** 07/17/09: I'm still confused about what happens to a perfect Accrued!?
9130                                If blnTmp23 = True Then
                                      ' ** First, find out whether this is the real IA, but with different settings.
9140                                  blnTmp24 = False
9150                                  If lngFlds = 12& Or lngFlds = 13& Then  ' ** The 13th field is curr_id.
                                        ' ** It's got the right number of fields, so proceed.
9160                                    If ![masterasset_TYPE] = "IA" Then
9170                                      blnTmp24 = True
9180                                    End If
9190                                  Else
                                        ' ** No [masterasset_TYPE], so this time let the criteria fail.
9200                                  End If
                                      ' ** And YES, an Accrued Interest Asset can also be a SWEEP!
                                      'THERE ISN'T ENOUGH CONFIRMING EVIDENCE YET!
                                      'IF ITS ONLY INDICATION IS assetno = 1, THEN THAT'S NOT NEARLY ENOUGH!
9210                                  If blnTmp24 = True Then
                                        ' ** Now, it only gets here if it's an IA!
                                        ' ** I would say this pretty much guarantees it's the real IA.
                                        ' ** It's possible the [assetno] or one of the other fields is off.
9220                                    If ![assetno] <> 1& Then
                                          ' ** 3-out-of-4 match.
                                          ' ** Add the ID cross-reference to tblVersion_Key.
9230                                      rstLoc2.AddNew
9240                                      rstLoc2![tbl_id] = lngCurrTblID
9250                                      rstLoc2![tbl_name] = strCurrTblName  'masterasset
9260                                      rstLoc2![fld_id] = lngCurrKeyFldID
9270                                      rstLoc2![fld_name] = strCurrKeyFldName
9280                                      rstLoc2![key_lng_id1] = ![assetno]
                                          'rstLoc2![key_txt_id1] =
9290                                      rstLoc2![key_lng_id2] = CLng(1)
                                          'rstLoc2![key_txt_id2] =
9300                                      rstLoc2.Update
9310                                    End If
                                        ' ** Save its values.
9320                                    rstLoc1.MoveFirst
9330                                    rstLoc1.Edit
9340                                    rstLoc1![description] = Nz(![description], "Accrued Interest Asset")  ' ** Let it bring over it's own
9350                                    rstLoc1![shareface] = Nz(![shareface], 0)                             ' **  description if it has one.
9360                                    rstLoc1![rate] = Nz(![rate], 0)  ' ** Default Rate must be Zero!
9370                                    rstLoc1![due] = ![due]
9380                                    rstLoc1![marketvaluecurrent] = ![marketvaluecurrent]  'Null  ' ** Why is this Null?
9390                                    rstLoc1![yield] = Nz(![yield], 0)
9400                                    If IsNull(![currentDate]) = False Then
9410                                      rstLoc1![currentDate] = ![currentDate]
9420                                    Else
9430                                      rstLoc1![currentDate] = Date
9440                                    End If
                                        'ADD CURR_ID!
9450                                    If lngFlds = 13& Then  ' ** The 13th field is curr_id.
9460  On Error Resume Next
9470                                      rstLoc1![curr_id] = ![curr_id]
9480                                      If ERR.Number <> 0 Then
9490  On Error GoTo ERRH
9500                                        rstLoc1![curr_id] = 150&  ' ** Default to USD.
9510                                      Else
9520  On Error GoTo ERRH
9530                                      End If
9540                                      If IsNull(rstLoc1![curr_id]) = True Then
9550                                        rstLoc1![curr_id] = 150&  ' ** Default to USD.
9560                                      Else
9570                                        If rstLoc1![curr_id] = 0 Then
9580                                          rstLoc1![curr_id] = 150&  ' ** Default to USD.
9590                                        End If
9600                                      End If
9610                                    Else
9620                                      rstLoc1![curr_id] = 150&  ' ** Default to USD.
9630                                    End If
9640                                    rstLoc1.Update
9650                                  Else  ' ** masterasset_TYPE.
9660                                    If ![cusip] = "999999999" And ![description] = "Accrued Interest Asset" Then
                                          ' ** Is it possible earlier versions had this without the IA?
                                          ' ** Treat it as real, and update the [assetno] if necessary.
9670                                      If ![assetno] <> 1& Then
                                            ' ** Add the ID cross-reference to tblVersion_Key.
9680                                        rstLoc2.AddNew
9690                                        rstLoc2![tbl_id] = lngCurrTblID
9700                                        rstLoc2![tbl_name] = strCurrTblName  'masterasset
9710                                        rstLoc2![fld_id] = lngCurrKeyFldID
9720                                        rstLoc2![fld_name] = strCurrKeyFldName
9730                                        rstLoc2![key_lng_id1] = ![assetno]
                                            'rstLoc2![key_txt_id1] =
9740                                        rstLoc2![key_lng_id2] = CLng(1)
                                            'rstLoc2![key_txt_id2] =
9750                                        rstLoc2.Update
9760                                      End If
                                          ' ** Save its values.
9770                                      rstLoc1.MoveFirst
9780                                      rstLoc1.Edit
9790                                      rstLoc1![description] = Nz(![description], "Accrued Interest Asset")
9800                                      rstLoc1![shareface] = Nz(![shareface], 0)
9810                                      rstLoc1![rate] = Nz(![rate], 0)  ' ** Default Rate must be Zero!
9820                                      rstLoc1![due] = ![due]
9830                                      rstLoc1![marketvaluecurrent] = ![marketvaluecurrent]  'Null  ' ** Why is this Null?
9840                                      rstLoc1![yield] = Nz(![yield], 0)
9850                                      If IsNull(![currentDate]) = False Then
9860                                        rstLoc1![currentDate] = ![currentDate]
9870                                      Else
9880                                        rstLoc1![currentDate] = Date
9890                                      End If
                                          'ADD CURR_ID!
9900                                      If lngFlds = 13& Then  ' ** The 13th field is curr_id.
9910  On Error Resume Next
9920                                        rstLoc1![curr_id] = ![curr_id]
9930                                        If ERR.Number <> 0 Then
9940  On Error GoTo ERRH
9950                                          rstLoc1![curr_id] = 150&  ' ** Default to USD.
9960                                        Else
9970  On Error GoTo ERRH
9980                                        End If
9990                                        If IsNull(rstLoc1![curr_id]) = True Then
10000                                         rstLoc1![curr_id] = 150&  ' ** Default to USD.
10010                                       Else
10020                                         If rstLoc1![curr_id] = 0 Then
10030                                           rstLoc1![curr_id] = 150&  ' ** Default to USD.
10040                                         End If
10050                                       End If
10060                                     Else
10070                                       rstLoc1![curr_id] = 150&  ' ** Default to USD.
10080                                     End If
10090                                     rstLoc1.Update
10100                                   ElseIf ![description] = "Accrued Interest Asset" And ![cusip] <> "999999999" Then
                                          ' ** Again, I'm treating it as the real IA.
10110                                     If ![assetno] <> 1& Then
                                            ' ** Add the ID cross-reference to tblVersion_Key.
10120                                       rstLoc2.AddNew
10130                                       rstLoc2![tbl_id] = lngCurrTblID
10140                                       rstLoc2![tbl_name] = strCurrTblName  'masterasset
10150                                       rstLoc2![fld_id] = lngCurrKeyFldID
10160                                       rstLoc2![fld_name] = strCurrKeyFldName
10170                                       rstLoc2![key_lng_id1] = ![assetno]
                                            'rstLoc2![key_txt_id1] =
10180                                       rstLoc2![key_lng_id2] = CLng(1)
                                            'rstLoc2![key_txt_id2] =
10190                                       rstLoc2.Update
10200                                     End If
                                          ' ** Save its values.
10210                                     rstLoc1.MoveFirst
10220                                     rstLoc1.Edit
10230                                     rstLoc1![description] = Nz(![description], "Accrued Interest Asset")
10240                                     rstLoc1![shareface] = Nz(![shareface], 0)
10250                                     rstLoc1![rate] = Nz(![rate], 0)  ' ** Default Rate must be Zero!
10260                                     rstLoc1![due] = ![due]
10270                                     rstLoc1![marketvaluecurrent] = ![marketvaluecurrent]  'Null  ' ** Why is this Null?
10280                                     rstLoc1![yield] = Nz(![yield], 0)
10290                                     If IsNull(![currentDate]) = False Then
10300                                       rstLoc1![currentDate] = ![currentDate]
10310                                     Else
10320                                       rstLoc1![currentDate] = Date
10330                                     End If
                                          'ADD CURR_ID!
10340                                     If lngFlds = 13& Then  ' ** The 13th field is curr_id.
10350 On Error Resume Next
10360                                       rstLoc1![curr_id] = ![curr_id]
10370                                       If ERR.Number <> 0 Then
10380 On Error GoTo ERRH
10390                                         rstLoc1![curr_id] = 150&  ' ** Default to USD.
10400                                       Else
10410 On Error GoTo ERRH
10420                                       End If
10430                                       If IsNull(rstLoc1![curr_id]) = True Then
10440                                         rstLoc1![curr_id] = 150&  ' ** Default to USD.
10450                                       Else
10460                                         If rstLoc1![curr_id] = 0 Then
10470                                           rstLoc1![curr_id] = 150&  ' ** Default to USD.
10480                                         End If
10490                                       End If
10500                                     Else
10510                                       rstLoc1![curr_id] = 150&  ' ** Default to USD.
10520                                     End If
10530                                     rstLoc1.Update
10540                                   ElseIf ![cusip] = "999999999" And ![description] <> "Accrued Interest Asset" Then
                                          ' ** Looks like a coincidence, and it's not the IA.
                                          ' ** Add the record to the new table.
10550                                     rstLoc1.AddNew
                                          ' ** Change the [cusip].
10560                                     rstLoc1![cusip] = "999999998"          ' ** It wasn't an IA, it's assetno wasn't 1, it's
10570                                     If IsNull(![description]) = True Then  ' ** description wasn't Accrued Interst Asset'.
10580                                       rstLoc1![description] = "UNKNOWN {may have been IA}"
10590                                       lngDupeUnks = lngDupeUnks + 1&
10600                                       ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
10610                                       arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
10620                                       arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
10630                                     Else
10640                                       If Trim(![description]) = vbNullString Then
10650                                         rstLoc1![description] = "UNKNOWN {may have been IA}"
10660                                         lngDupeUnks = lngDupeUnks + 1&
10670                                         ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
10680                                         arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
10690                                         arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
10700                                       Else
10710                                         rstLoc1![description] = Trim(![description])  ' ** Demo one retains it's [description]
10720                                       End If
10730                                     End If
10740                                     rstLoc1![shareface] = Nz(![shareface], 0)
10750                                     If IsNull(![assettype]) = False Then
10760                                       blnFound = False
10770                                       For lngY = 0& To (lngAssetTypes - 1&)
10780                                         If arr_varAssetType(AT_TYP, lngY) = ![assettype] Then
10790                                           blnFound = True
10800                                           Exit For
10810                                         End If
10820                                       Next
10830                                       If blnFound = False Then
10840                                         rstLoc1![assettype] = "75"  ' ** Other.
10850                                       Else
10860                                         rstLoc1![assettype] = ![assettype]
10870                                       End If
10880                                       blnFound = True  ' ** Reset.
10890                                     Else
10900                                       rstLoc1![assettype] = "75"  ' ** Other.
10910                                     End If
10920                                     If IsNull(![rate]) = False Then
10930                                       If ![rate] > 1# Then
10940                                         rstLoc1![rate] = ![rate] / 100#
10950                                       Else
10960                                         rstLoc1![rate] = ![rate]
10970                                       End If
10980                                     Else
10990                                       rstLoc1![rate] = CDbl(0)  ' ** Default Rate must be Zero!
11000                                     End If
11010                                     rstLoc1![due] = ![due]
11020                                     rstLoc1![marketvalue] = Null  'Nz(![marketvalue], 0)
11030                                     rstLoc1![marketvaluecurrent] = Nz(![marketvaluecurrent], 0)
11040                                     rstLoc1![yield] = Nz(![yield], 0)
11050                                     If IsNull(![currentDate]) = False Then
11060                                       rstLoc1![currentDate] = ![currentDate]
11070                                     Else
11080                                       rstLoc1![currentDate] = Date
11090                                     End If
11100                                     rstLoc1![masterasset_TYPE] = "RA"
                                          'ADD CURR_ID!
11110                                     If lngFlds = 13& Then  ' ** The 13th field is curr_id.
11120 On Error Resume Next
11130                                       rstLoc1![curr_id] = ![curr_id]
11140                                       If ERR.Number <> 0 Then
11150 On Error GoTo ERRH
11160                                         rstLoc1![curr_id] = 150&  ' ** Default to USD.
11170                                       Else
11180 On Error GoTo ERRH
11190                                       End If
11200                                       If IsNull(rstLoc1![curr_id]) = True Then
11210                                         rstLoc1![curr_id] = 150&  ' ** Default to USD.
11220                                       Else
11230                                         If rstLoc1![curr_id] = 0 Then
11240                                           rstLoc1![curr_id] = 150&  ' ** Default to USD.
11250                                         End If
11260                                       End If
11270                                     Else
11280                                       rstLoc1![curr_id] = 150&  ' ** Default to USD.
11290                                     End If
11300                                     rstLoc1.Update  ' ** I'm banking on no conflict with 999999998!
11310                                     rstLoc1.Bookmark = rstLoc1.LastModified
                                          ' ** Add the ID cross-reference to tblVersion_Key.
11320                                     rstLoc2.AddNew
11330                                     rstLoc2![tbl_id] = lngCurrTblID
11340                                     rstLoc2![tbl_name] = strCurrTblName  'masterasset
11350                                     rstLoc2![fld_id] = lngCurrKeyFldID
11360                                     rstLoc2![fld_name] = strCurrKeyFldName
11370                                     rstLoc2![key_lng_id1] = ![assetno]
                                          'rstLoc2![key_txt_id1] =
11380                                     rstLoc2![key_lng_id2] = rstLoc1![assetno]
                                          'rstLoc2![key_txt_id2] =
11390                                     rstLoc2.Update
11400                                   Else
                                          ' ** Legitimate asset; move it!
                                          ' ** Add the record to the new table.
11410                                     rstLoc1.AddNew
11420                                     If IsNull(![cusip]) = True Then
11430                                       strTmp04 = ("XX_" & CStr(![assetno]))
11440                                       rstLoc1![cusip] = strTmp04  ' ** Max 9 chars.
11450                                     Else
11460                                       If Trim(![cusip]) = vbNullString Then
11470                                         rstLoc1![cusip] = ("XX_" & CStr(![assetno]))
11480                                       Else
11490                                         rstLoc1![cusip] = Trim(![cusip])
11500                                       End If
11510                                     End If
11520                                     If IsNull(![description]) = True Then
11530                                       rstLoc1![description] = ("UNKNOWN" & CStr(![assetno]))
11540                                       lngDupeUnks = lngDupeUnks + 1&
11550                                       ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
11560                                       arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
11570                                       arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
11580                                     Else
11590                                       If Trim(![description]) = vbNullString Then
11600                                         rstLoc1![description] = ("UNKNOWN" & CStr(![assetno]))
11610                                         lngDupeUnks = lngDupeUnks + 1&
11620                                         ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
11630                                         arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
11640                                         arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
11650                                       Else
11660                                         rstLoc1![description] = Trim(![description])
11670                                       End If
11680                                     End If
11690                                     rstLoc1![shareface] = Nz(![shareface], 0)
11700                                     If IsNull(![assettype]) = False Then
11710                                       blnFound = False
11720                                       For lngY = 0& To (lngAssetTypes - 1&)
11730                                         If arr_varAssetType(AT_TYP, lngY) = ![assettype] Then
11740                                           blnFound = True
11750                                           Exit For
11760                                         End If
11770                                       Next
11780                                       If blnFound = False Then
11790                                         rstLoc1![assettype] = "75"  ' ** Other.
11800                                       Else
11810                                         rstLoc1![assettype] = ![assettype]
11820                                       End If
11830                                       blnFound = True  ' ** Reset.
11840                                     Else
11850                                       rstLoc1![assettype] = "75"  ' ** Other.
11860                                     End If
11870                                     If IsNull(![rate]) = False Then
11880                                       If ![rate] > 1# Then
11890                                         rstLoc1![rate] = ![rate] / 100#
11900                                       Else
11910                                         rstLoc1![rate] = ![rate]
11920                                       End If
11930                                     Else
11940                                       rstLoc1![rate] = CDbl(0)  ' ** Default Rate must be Zero!
11950                                     End If
11960                                     rstLoc1![due] = ![due]
11970                                     rstLoc1![marketvalue] = Null  'Nz(![marketvalue], 0)
11980                                     rstLoc1![marketvaluecurrent] = Nz(![marketvaluecurrent], 0)
11990                                     rstLoc1![yield] = Nz(![yield], 0)
12000                                     If IsNull(![currentDate]) = False Then
12010                                       rstLoc1![currentDate] = ![currentDate]
12020                                     Else
12030                                       rstLoc1![currentDate] = Date
12040                                     End If
12050                                     rstLoc1![masterasset_TYPE] = "RA"
                                          'ADD CURR_ID!
12060                                     If lngFlds = 13& Then  ' ** The 13th field is curr_id.
12070 On Error Resume Next
12080                                       rstLoc1![curr_id] = ![curr_id]
12090                                       If ERR.Number <> 0 Then
12100 On Error GoTo ERRH
12110                                         rstLoc1![curr_id] = 150&  ' ** Default to USD.
12120                                       Else
12130 On Error GoTo ERRH
12140                                       End If
12150                                       If IsNull(rstLoc1![curr_id]) = True Then
12160                                         rstLoc1![curr_id] = 150&  ' ** Default to USD.
12170                                       Else
12180                                         If rstLoc1![curr_id] = 0 Then
12190                                           rstLoc1![curr_id] = 150&  ' ** Default to USD.
12200                                         End If
12210                                       End If
12220                                     Else
12230                                       rstLoc1![curr_id] = 150&  ' ** Default to USD.
12240                                     End If
12250 On Error Resume Next
12260                                     rstLoc1.Update
12270                                     If ERR.Number <> 0 Then
12280                                       If ERR.Number = 3022 Then
                                              ' ** Error 3022: The changes you requested to the table were not successful because they
                                              ' **             would create duplicate values in the index, primary key, or relationship.
12290                                         If gblnDev_NoErrHandle = True Then
12300 On Error GoTo 0
12310                                         Else
12320 On Error GoTo ERRH
12330                                         End If
12340                                         lngDupeNum = lngDupeNum + 1&
12350                                         strTmp04 = Trim(![cusip])  ' ** Max 9 chars.
12360                                         If Len(strTmp04) > 7 Then
12370                                           If lngDupeNum < 10& Then
12380                                             strTmp04 = Left(strTmp04, 7) & "_" & CStr(lngDupeNum)
12390                                           Else
12400                                             strTmp04 = Left(strTmp04, 6) & "_" & CStr(lngDupeNum)
12410                                           End If
12420                                         Else
12430                                           strTmp04 = strTmp04 & "_" & CStr(lngDupeNum)
12440                                         End If
12450                                         rstLoc1![cusip] = strTmp04
12460                                         rstLoc1.Update
12470                                       Else
12480                                         intRetVal = -6
12490                                         blnContinue = False
12500                                         lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
12510                                         MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
                                                "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
                                                vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
12520                                         rstLoc1.CancelUpdate
12530                                         If gblnDev_NoErrHandle = True Then
12540 On Error GoTo 0
12550                                         Else
12560 On Error GoTo ERRH
12570                                         End If
12580                                       End If
12590                                     Else
12600                                       If gblnDev_NoErrHandle = True Then
12610 On Error GoTo 0
12620                                       Else
12630 On Error GoTo ERRH
12640                                       End If
12650                                     End If
12660                                     If blnContinue = True Then
12670                                       rstLoc1.Bookmark = rstLoc1.LastModified
                                            ' ** Add the ID cross-reference to tblVersion_Key.
12680                                       rstLoc2.AddNew
12690                                       rstLoc2![tbl_id] = lngCurrTblID
12700                                       rstLoc2![tbl_name] = strCurrTblName  'masterasset
12710                                       rstLoc2![fld_id] = lngCurrKeyFldID
12720                                       rstLoc2![fld_name] = strCurrKeyFldName
12730                                       rstLoc2![key_lng_id1] = ![assetno]
                                            'rstLoc2![key_txt_id1] =
12740                                       rstLoc2![key_lng_id2] = rstLoc1![assetno]
                                            'rstLoc2![key_txt_id2] =
12750                                       rstLoc2.Update
12760                                     End If
12770                                   End If  ' ** Confused with IA.
12780                                 End If  ' ** masterasset_TYPE.
12790                               Else  ' ** blnTmp23.
                                      ' ** Add the record to the new table.
12800                                 rstLoc1.AddNew
12810                                 If IsNull(![cusip]) = True Then
12820                                   rstLoc1![cusip] = ("XX_" & CStr(![assetno]))  ' ** Max 9 chars.
12830                                 Else
12840                                   If Trim(![cusip]) = vbNullString Then
12850                                     rstLoc1![cusip] = ("XX_" & CStr(![assetno]))
12860                                   Else
12870                                     rstLoc1![cusip] = Trim(![cusip])
12880                                   End If
12890                                 End If
12900                                 If IsNull(![description]) = True Then
12910                                   rstLoc1![description] = ("UNKNOWN" & CStr(![assetno]))
12920                                   lngDupeUnks = lngDupeUnks + 1&
12930                                   ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
12940                                   arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
12950                                   arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
12960                                 Else
12970                                   If Trim(![description]) = vbNullString Then
12980                                     rstLoc1![description] = ("UNKNOWN" & CStr(![assetno]))
12990                                     lngDupeUnks = lngDupeUnks + 1&
13000                                     ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
13010                                     arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
13020                                     arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
13030                                   Else
13040                                     rstLoc1![description] = Trim(![description])
13050                                   End If
13060                                 End If
13070                                 rstLoc1![shareface] = Nz(![shareface], 0)
13080                                 If IsNull(![assettype]) = False Then
13090                                   blnFound = False
13100                                   For lngY = 0& To (lngAssetTypes - 1&)
13110                                     If arr_varAssetType(AT_TYP, lngY) = ![assettype] Then
13120                                       blnFound = True
13130                                       Exit For
13140                                     End If
13150                                   Next
13160                                   If blnFound = False Then
13170                                     rstLoc1![assettype] = "75"  ' ** Other.
13180                                   Else
13190                                     rstLoc1![assettype] = ![assettype]
13200                                   End If
13210                                   blnFound = True  ' ** Reset.
13220                                 Else
13230                                   rstLoc1![assettype] = "75"  ' ** Other.
13240                                 End If
13250                                 If IsNull(![rate]) = False Then
13260                                   If ![rate] > 1# Then
13270                                     rstLoc1![rate] = ![rate] / 100#
13280                                   Else
13290                                     rstLoc1![rate] = ![rate]
13300                                   End If
13310                                 Else
13320                                   rstLoc1![rate] = CDbl(0)  ' ** Default Rate must be Zero!
13330                                 End If
13340                                 rstLoc1![due] = ![due]
13350                                 rstLoc1![marketvalue] = Null  'Nz(![marketvalue], 0)
13360                                 rstLoc1![marketvaluecurrent] = Nz(![marketvaluecurrent], 0)
13370                                 rstLoc1![yield] = Nz(![yield], 0)
13380                                 If IsNull(![currentDate]) = False Then
13390                                   rstLoc1![currentDate] = ![currentDate]
13400                                 Else
13410                                   rstLoc1![currentDate] = Date
13420                                 End If
13430                                 rstLoc1![masterasset_TYPE] = "RA"
                                      'ADD CURR_ID!
13440                                 If lngFlds = 13& Then  ' ** The 13th field is curr_id.
13450 On Error Resume Next
13460                                   rstLoc1![curr_id] = ![curr_id]
13470                                   If ERR.Number <> 0 Then
13480 On Error GoTo ERRH
13490                                     rstLoc1![curr_id] = 150&  ' ** Default to USD.
13500                                   Else
13510 On Error GoTo ERRH
13520                                   End If
13530                                   If IsNull(rstLoc1![curr_id]) = True Then
13540                                     rstLoc1![curr_id] = 150&  ' ** Default to USD.
13550                                   Else
13560                                     If rstLoc1![curr_id] = 0 Then
13570                                       rstLoc1![curr_id] = 150&  ' ** Default to USD.
13580                                     End If
13590                                   End If
13600                                 Else
13610                                   rstLoc1![curr_id] = 150&  ' ** Default to USD.
13620                                 End If
13630 On Error Resume Next
13640                                 rstLoc1.Update
13650                                 If ERR.Number <> 0 Then
13660                                   If ERR.Number = 3022 Then
                                          ' ** Error 3022: The changes you requested to the table were not successful because they
                                          ' **             would create duplicate values in the index, primary key, or relationship.
13670                                     If gblnDev_NoErrHandle = True Then
13680 On Error GoTo 0
13690                                     Else
13700 On Error GoTo ERRH
13710                                     End If
13720                                     lngDupeNum = lngDupeNum + 1&
13730                                     strTmp04 = Trim(![cusip])  ' ** Max 9 chars.
13740                                     If Len(strTmp04) > 7 Then
13750                                       If lngDupeNum < 10& Then
13760                                         strTmp04 = Left(strTmp04, 7) & "_" & CStr(lngDupeNum)
13770                                       Else
13780                                         strTmp04 = Left(strTmp04, 6) & "_" & CStr(lngDupeNum)
13790                                       End If
13800                                     Else
13810                                       strTmp04 = strTmp04 & "_" & CStr(lngDupeNum)
13820                                     End If
13830                                     rstLoc1![cusip] = strTmp04
13840                                     rstLoc1.Update
13850                                   Else
13860                                     intRetVal = -6
13870                                     blnContinue = False
13880                                     lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
13890                                     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
                                            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
                                            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
13900                                     rstLoc1.CancelUpdate
13910                                     If gblnDev_NoErrHandle = True Then
13920 On Error GoTo 0
13930                                     Else
13940 On Error GoTo ERRH
13950                                     End If
13960                                   End If
13970                                 Else
13980                                   If gblnDev_NoErrHandle = True Then
13990 On Error GoTo 0
14000                                   Else
14010 On Error GoTo ERRH
14020                                   End If
14030                                 End If
14040                                 If blnContinue = True Then
14050                                   rstLoc1.Bookmark = rstLoc1.LastModified
                                        ' ** Add the ID cross-reference to tblVersion_Key.
14060                                   arr_varMasterAsset(MA_NEW_ANO, lngE) = rstLoc1![assetno]
14070                                   arr_varMasterAsset(MA_NEW_MVC, lngE) = rstLoc1![marketvaluecurrent]
14080                                   rstLoc2.AddNew
14090                                   rstLoc2![tbl_id] = lngCurrTblID
14100                                   rstLoc2![tbl_name] = strCurrTblName  'masterasset
14110                                   rstLoc2![fld_id] = lngCurrKeyFldID
14120                                   rstLoc2![fld_name] = strCurrKeyFldName
14130                                   rstLoc2![key_lng_id1] = ![assetno]
                                        'rstLoc2![key_txt_id1] =
14140                                   rstLoc2![key_lng_id2] = rstLoc1![assetno]
                                        'rstLoc2![key_txt_id2] =
14150                                   rstLoc2.Update
14160                                 Else
14170                                   Exit For
14180                                 End If
14190                               End If  ' ** blnTmp23.
14200                             End If  ' ** blnTmp22.

14210                           End If  ' ** blnTmp22.
14220                           If lngX < lngRecs Then .MoveNext
14230                         Next
14240                         rstLoc2.Close
                              ' ** Copy the original Accrued Interest Asset values to tblTemplate_MasterAsset.
14250                         With rstLoc1
14260                           If .BOF = True And .EOF = True Then
                                  ' ** Shouldn't ever happen!
14270                           Else
14280                             .MoveFirst
14290                             If ![masterasset_TYPE] = "IA" Then
14300                               Set rstLoc2 = dbsLoc.OpenRecordset("tblTemplate_MasterAsset", dbOpenDynaset, dbConsistent)
14310                               blnTmp22 = False
14320                               If rstLoc2.BOF = True And rstLoc2.EOF = True Then
                                      ' ** Good, it's empty.
14330                               Else
14340                                 blnTmp22 = True
14350                               End If
14360                               If blnTmp22 = True Then
14370                                 rstLoc2.Close
                                      ' ** Because the numbering of my delete queries changes so often, get the query's name this way.
14380                                 strTmp04 = vbNullString
14390                                 For Each qdf In dbsLoc.QueryDefs
14400                                   With qdf
14410                                     If Left(.Name, 19) = "qryTmp_Table_Empty_" Then
14420                                       If Right(.Name, 24) = "_tblTemplate_MasterAsset" Then
14430                                         strTmp04 = .Name
14440                                         Exit For
14450                                       End If
14460                                     End If
14470                                   End With
14480                                 Next
14490                                 If strTmp04 <> vbNullString Then
14500                                   Set qdf = dbsLoc.QueryDefs(strTmp04)
14510                                   qdf.Execute
14520                                 End If
14530                                 Set rstLoc2 = dbsLoc.OpenRecordset("tblTemplate_MasterAsset", dbOpenDynaset, dbConsistent)
14540                                 blnTmp22 = False
14550                               End If
14560                               rstLoc2.AddNew
14570                               For Each fld In rstLoc2.Fields
14580                                 fld.Value = .Fields(fld.Name)
                                      'ADD CURR_ID!
                                      ' ** Under these conditions, curr_id is accommodated.
14590                               Next
14600                               rstLoc2.Update
14610                               rstLoc2.Close
14620                             End If
14630                           End If
14640                         End With
14650                         rstLoc1.Close
14660                       End If  ' ** Records present.
14670                       .Close
14680                     End With  ' ** rstLnk.
14690                   End If  ' ** blnFound.

14700                   lngStats = lngStats + 1&
14710                   lngE = lngStats - 1&
14720                   ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
14730                   arr_varStat(STAT_ORD, lngE) = CInt(6)
14740                   arr_varStat(STAT_NAM, lngE) = "Master Assets: "
14750                   arr_varStat(STAT_CNT, lngE) = CLng(lngRecs)
14760                   arr_varStat(STAT_DSC, lngE) = vbNullString

                        ' ** Make sure the current market value of the IA asset is copied!
14770                   If gdblCrtRpt_CostTot > 0# Then
14780                     Set rstLoc1 = dbsLoc.OpenRecordset("masterasset", dbOpenDynaset, dbConsistent)
14790                     With rstLoc1
14800                       .MoveFirst
14810                       blnFound = True
14820                       If ![assetno] <> 1& Then
14830                         .FindFirst "[assetno] = 1"
14840                         If .NoMatch = True Then
                                ' ** Skip the whole process. Shouldn't ever happen.
14850                           blnFound = False
14860                         End If
14870                       End If
14880                       If blnFound = True Then
14890                         .Edit
14900                         ![marketvaluecurrent] = gdblCrtRpt_CostTot  ' ** Borrowing this for the sweep MarketValueCurrent.
14910                         If gstrCrtRpt_Version <> vbNullString Then
14920                           If ![description] <> gstrCrtRpt_Version Then
14930                             ![description] = gstrCrtRpt_Version  ' ** Borrowing this for the sweep asset description.
14940                           End If
14950                         End If
14960                         .Update
14970                       End If
14980                       .Close
14990                     End With
15000                     Set rstLoc1 = Nothing
15010                   End If
15020                   gdblCrtRpt_CostTot = 0#: gstrCrtRpt_Version = vbNullString
15030                   DoEvents

                        ' ** Check the Default AssetNo from the Account table's TaxLot field against AssetNo's.
15040                   Set rstLoc1 = dbsLoc.OpenRecordset("account", dbOpenDynaset, dbConsistent)
15050                   For lngX = 0 To (lngAccts - 1&)
15060                     If Val(arr_varAcct(A_DASTNO, lngX)) > 0& Then
15070                       For lngY = 0& To (lngMasterAssets - 1&)
15080                         If arr_varMasterAsset(MA_OLD_ANO, lngY) = CLng(arr_varAcct(A_DASTNO, lngX)) Then
15090                           If arr_varMasterAsset(MA_NEW_ANO, lngY) <> arr_varMasterAsset(MA_OLD_ANO, lngY) Then
15100                             With rstLoc1
15110                               If arr_varAcct(A_NUM_N, lngX) = vbNullString Then
15120                                 .FindFirst "[accountno] = '" & arr_varAcct(A_NUM, lngX) & "'"
15130                               Else
15140                                 .FindFirst "[accountno] = '" & arr_varAcct(A_NUM_N, lngX) & "'"
15150                               End If
15160                               If .NoMatch = False Then
15170                                 .Edit
15180                                 ![taxlot] = CStr(arr_varMasterAsset(MA_NEW_ANO, lngY))
15190                                 .Update
15200                               End If
15210                             End With
15220                           End If
15230                         End If
15240                       Next
15250                     End If
15260                   Next
15270                   rstLoc1.Close

15280                 End If  ' ** blnContinue.

15290                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ** These must remain available for Journal and LedgerArchive, below!
15300                   lngTmp15 = 0&
15310                   ReDim arr_varTmp01(5, 0)
15320                   datTmp28 = 0

                        ' ******************************
                        ' ** Table: ActiveAssets.
                        ' ******************************

                        ' ** Step 15: ActiveAssets.
15330                   dblPB_ThisStep = 15#
15340                   Version_Status 3, dblPB_ThisStep, "Active Assets"  ' ** Function: Below.

15350                   strCurrTblName = "ActiveAssets"
15360                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

15370                   blnFound = False: blnFound2 = False: lngRecs = 0&
15380                   strTmp04 = vbNullString: strTmp08 = vbNullString
15390                   strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString
15400                   For lngX = 0& To (lngOldTbls - 1&)
15410                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
15420                       blnFound = True
15430                       Exit For
15440                     End If
15450                   Next

15460                   If blnFound = True Then
15470                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
15480                     With rstLnk
15490                       If .BOF = True And .EOF = True Then
                              ' ** I would have expected this to have records!
15500                       Else
15510                         strCurrKeyFldName = "assetno"
15520                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
15530                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
15540                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Of 16 example TrustDta.mdb's, all have 17 fields
                              ' ** Current field count is 20 fields.
                              ' ** Table: ActiveAssets
                              ' **   ![assetno]             dbLong
                              ' **   ![accountno]           dbText
                              ' **   ![assetdate]           dbDate
                              ' **   ![transdate]           dbDate
                              ' **   ![postdate]            dbDate
                              ' **   ![shareface]           dbDouble
                              ' **   ![due]                 dbDate
                              ' **   ![rate]                dbDouble
                              ' **   ![averagepriceperunit] dbDouble
                              ' **   ![priceperunit]        dbDouble
                              ' **   ![icash]               dbCurrency
                              ' **   ![pcash]               dbCurrency
                              ' **   ![cost]                dbCurrency
                              ' **   ![description]         dbText
                              ' **   ![posted]              dbDate
                              ' **   ![IsAverage]           dbBoolean
                              ' **   ![Location_ID]         dbLong
                              ' **   ![curr_id]  Defaults to 150.
                              ' **   ![cost_usd]
                              ' **   ![market_usd]
                              ' ** No table references it directly, and it does not have its own unique key.
                              'SO, ARE WE POPULATING COST_USD AND MARKET_USD ON NEW CONVERTS?
15550                         .MoveLast
15560                         lngRecs = .RecordCount
15570                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Function: Below.
15580                         .MoveFirst
15590                         lngFlds = .Fields.Count
15600                         For Each fld In .Fields
15610                           With fld
15620                             If .Name = "Location Id" Then
                                    ' ** Has old field name.
15630                               blnFound2 = True
15640                               Exit For
15650                             End If
15660                           End With
15670                         Next
15680                         strTmp06 = "Location_ID": strTmp07 = "Location Id"
15690                         If blnFound2 = False Then strTmp05 = strTmp06 Else strTmp05 = strTmp07
15700                         For lngX = 1& To lngRecs
15710                           Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Function: Below.
                                ' ** Add the record to the new table.
15720                           rstLoc1.AddNew
15730                           rstLoc2.MoveFirst
15740                           rstLoc2.FindFirst "[tbl_name] = 'masterasset' And [fld_name] = 'assetno' And [key_lng_id1] = " & CStr(![assetno])
15750                           If rstLoc2.NoMatch = True And ![assetno] = 1& Then
                                  ' ** It's the 'Accrued Interest Asset', masterasset_TYPE = 'IA',
                                  ' ** which doesn't get moved.
15760                             rstLoc1![assetno] = 1&
15770                           ElseIf rstLoc2.NoMatch = False Then
15780                             rstLoc1![assetno] = rstLoc2![key_lng_id2]
15790                           Else
                                  ' ** We'll have to assume that this is an orphan.
                                  ' ** Add the record to the new table.
15800                             lngTmp13 = ![assetno]
15810                             lngTmp14 = 0&
15820                             Set rstLoc3 = dbsLoc.OpenRecordset("masterasset", dbOpenDynaset, dbConsistent)
15830                             rstLoc3.AddNew
15840                             lngDupeNum = lngDupeNum + 1&  ' ** Though not really a dupe, we need as unique number.
15850                             rstLoc3![cusip] = "UNK" & CStr(lngDupeNum)
15860                             rstLoc3![description] = "UNKNOWN {ActiveAssets, Date " & _
                                    Format(![transdate], "mm/dd/yyyy") & ", Asset " & CStr(lngTmp13) & "}"
15870                             rstLoc3![shareface] = ![shareface]  ' ** Do we want to update this when we're finished?
15880                             rstLoc3![assettype] = "75"  ' ** Other.
15890                             rstLoc3![rate] = ![rate]
15900                             rstLoc3![due] = ![due]
15910                             rstLoc3![marketvalue] = Null  'CCur(0)
15920                             rstLoc3![marketvaluecurrent] = (![shareface] * Nz(![priceperunit], 0))
15930                             rstLoc3![yield] = CDbl(0)
15940                             rstLoc3![currentDate] = CDate(Format(Nz(![assetdate], Date), "mm/dd/yyyy"))
15950                             rstLoc3![masterasset_TYPE] = "RA"
                                  'ADD CURR_ID!
15960                             If lngFlds = 20& Then  ' ** The 18th field is curr_id.
15970 On Error Resume Next
15980                               rstLoc3![curr_id] = ![curr_id]
15990                               If ERR.Number <> 0 Then
16000 On Error GoTo ERRH
16010                                 rstLoc3![curr_id] = 150&  ' ** Default to USD.
16020                               Else
16030 On Error GoTo ERRH
16040                               End If
16050                               If IsNull(rstLoc3![curr_id]) = True Then
16060                                 rstLoc3![curr_id] = 150&  ' ** Default to USD.
16070                               Else
16080                                 If rstLoc3![curr_id] = 0& Then
16090                                   rstLoc3![curr_id] = 150&  ' ** Default to USD.
16100                                 End If
16110                               End If
16120                             Else
16130                               rstLoc3![curr_id] = 150&  ' ** Default to USD.
16140                             End If
16150                             rstLoc3.Update
16160                             lngDupeUnks = lngDupeUnks + 1&
16170                             ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
16180                             arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
16190                             arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = "masterasset"
16200                             rstLoc3.Bookmark = rstLoc3.LastModified
16210                             lngTmp14 = rstLoc3![assetno]
16220                             rstLoc3.Close
16230                             Set rstLoc3 = Nothing
16240                             rstLoc1![assetno] = lngTmp14
                                  ' ** Add the ID cross-reference to tblVersion_Key.
16250                             rstLoc2.AddNew
16260                             rstLoc2![tbl_id] = lngCurrTblID
16270                             rstLoc2![tbl_name] = "masterasset"  'strCurrTblName  'ActiveAssets
16280                             rstLoc2![fld_id] = lngCurrKeyFldID
16290                             rstLoc2![fld_name] = "assetno"  'strCurrKeyFldName
16300                             rstLoc2![key_lng_id1] = lngTmp13
                                  'rstLoc2![key_txt_id1] =
16310                             rstLoc2![key_lng_id2] = lngTmp14
                                  'rstLoc2![key_txt_id2] =
16320                             rstLoc2.Update
16330                             lngTmp13 = 0&: lngTmp14 = 0&
16340                           End If
                                ' ** Since accountno's only get dropped if they're one of our 99 accounts
                                ' ** or a dupe, this would still match one of the original ones.
                                ' ** An orphan, however, requires serious attention.
16350                           strTmp04 = Trim(![accountno])
16360                           blnFound = False
16370                           For lngY = 0& To (lngAccts - 1&)
16380                             If arr_varAcct(A_NUM, lngY) = strTmp04 Then
16390                               blnFound = True
16400                               Exit For
16410                             End If
16420                           Next
16430                           If blnFound = False Then
                                  ' ** This could only mean it's an orphan.
16440                             Set rstLoc3 = dbsLoc.OpenRecordset("account", dbOpenDynaset, dbConsistent)
16450                             rstLoc3.AddNew
16460                             rstLoc3![accountno] = strTmp04  ' ** dbText 15
16470                             rstLoc3![shortname] = "UNKNOWN" & CStr(lngX)  ' ** dbText 30
16480                             rstLoc3![legalname] = "UNKNOWN {ActiveAssets, " & _
                                    "Posting Date " & Format(![transdate], "mm/dd/yyyy") & "}" ' ** dbText 100
16490                             rstLoc3![accounttype] = "85"  ' ** Other.
16500                             rstLoc3![cotrustee] = "No"
16510                             rstLoc3![amendments] = "No"
16520                             rstLoc3![courtsupervised] = "No"
16530                             rstLoc3![discretion] = "No"
16540                             rstLoc3![ICash] = ![ICash]
16550                             rstLoc3![PCash] = ![PCash]
16560                             rstLoc3![Cost] = ![Cost]
16570                             rstLoc3![predate] = (![transdate] - 1)
16580                             rstLoc3![investmentobj] = "Other"
16590                             rstLoc3![numCopies] = CInt(1)
16600                             rstLoc3![account_SWEEP] = CBool(False)
16610                             rstLoc3![taxlot] = "0"
                                  'ADD CURR_ID!
16620                             If lngFlds = 20& Then  ' ** The 18th field is curr_id.
16630 On Error Resume Next
16640                               rstLoc3![curr_id] = ![curr_id]
16650                               If ERR.Number <> 0 Then
16660 On Error GoTo ERRH
16670                                 rstLoc3![curr_id] = 150&  ' ** Default to USD.
16680                               Else
16690 On Error GoTo ERRH
16700                               End If
16710                               If IsNull(rstLoc3![curr_id]) = True Then
16720                                 rstLoc3![curr_id] = 150&  ' ** Default to USD.
16730                               Else
16740                                 If rstLoc3![curr_id] = 0& Then
16750                                   rstLoc3![curr_id] = 150&  ' ** Default to USD.
16760                                 End If
16770                               End If
16780                             Else
16790                               rstLoc3![curr_id] = 150&  ' ** Default to USD.
16800                             End If
16810                             rstLoc3.Update
16820                             lngDupeUnks = lngDupeUnks + 1&
16830                             ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
16840                             arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
16850                             arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = "account"
16860                             rstLoc3.Close
16870                             Set rstLoc3 = Nothing
16880                             lngAccts = lngAccts + 1&
16890                             lngE = lngAccts - 1&
16900                             ReDim Preserve arr_varAcct(A_ELEMS, lngE)
16910                             arr_varAcct(A_NUM, lngE) = strTmp04
16920                             arr_varAcct(A_NUM_N, lngE) = "#ORPHAN_AA"
16930                             arr_varAcct(A_NAM, lngE) = "UNKNOWN" & CStr(lngX)
16940                             arr_varAcct(A_TYP, lngE) = "85"
16950                             arr_varAcct(A_ADMIN, lngE) = Null
16960                             arr_varAcct(A_ADMIN_N, lngE) = CLng(0)
16970                             arr_varAcct(A_SCHED, lngE) = Null
16980                             arr_varAcct(A_SCHED_N, lngE) = CLng(0)
16990                             arr_varAcct(A_DROPPED, lngE) = CBool(False)
17000                             arr_varAcct(A_ACCT99, lngE) = vbNullString
17010                             arr_varAcct(A_DASTNO, lngE) = "0"
17020                             rstLoc1![accountno] = strTmp04
17030                           Else
17040                             rstLoc1![accountno] = strTmp04
17050                           End If
17060                           If CLng(![assetdate]) = CDbl(![assetdate]) Then
17070                             strTmp08 = (Format(![assetdate], "mm/dd/yyyy") & " 09:00:00 AM")
17080                           Else
17090                             strTmp08 = Format(![assetdate], "mm/dd/yyyy hh:nn:ss AM/PM")
17100                           End If
17110                           rstLoc1![assetdate] = CDate(strTmp08)
17120                           rstLoc1![transdate] = ![transdate]
17130                           rstLoc1![postdate] = ![postdate]
17140                           rstLoc1![shareface] = Nz(![shareface], 0)
17150                           rstLoc1![due] = ![due]
17160                           rstLoc1![rate] = Nz(![rate], 0)
17170                           rstLoc1![averagepriceperunit] = Nz(![averagepriceperunit], 0)
17180                           rstLoc1![priceperunit] = Nz(![priceperunit], 0)
17190                           rstLoc1![ICash] = Nz(![ICash], 0)
17200                           rstLoc1![PCash] = Nz(![PCash], 0)
17210                           rstLoc1![Cost] = Nz(![Cost], 0)
17220                           If IsNull(![description]) = False Then
17230                             If Trim(![description]) <> vbNullString Then
17240                               rstLoc1![description] = ![description]
17250                             End If
17260                           End If
17270                           If ![posted] = CDate("12/30/1899 12:00:00 AM") Then
                                  ' ** All versions up to v2.1.55 or v2.1.56 put the Journal's [posted] field here,
                                  ' ** which is a Boolean field, rather than Now() at the time of posting.
                                  ' ** This brings over the Ledger's [posted] field for the initial entry.
17280                             varTmp00 = DLookup("[posted_L]", "qryVersion_Convert_02", _
                                    "[uniqueid] = '" & Left(![accountno] & String(15, "_"), 15) & Right(String(6, "0") & _
                                    CStr(![assetno]), 6) & Format(![assetdate], "mmddyyhhnnss") & "'")
17290                             If IsNull(varTmp00) = False Then
17300                               rstLoc1![posted] = CDate(varTmp00)
17310                             Else
17320                               rstLoc1![posted] = ![posted]
17330                             End If
17340                           Else
17350                             rstLoc1![posted] = ![posted]
17360                           End If
17370                           rstLoc1![IsAverage] = ![IsAverage]
17380                           rstLoc2.MoveFirst
                                'Null [Location_ID]
                                'Should that be changed to 1, for {Unassigned}?
                                'Not in the search; default to 0.  WHAT DOES THIS MEAN?
17390                           rstLoc2.FindFirst "[tbl_name] = 'Location' And [fld_name] = 'Location_ID' And " & _
                                  "[key_lng_id1] = " & CStr(Nz(.Fields(strTmp05), 0&))
17400                           If rstLoc2.NoMatch = False Then
17410                             rstLoc1![Location_ID] = rstLoc2![key_lng_id2]
17420                           Else
                                  ' ** It's an orphan; don't bother creating one.
17430                             rstLoc1![Location_ID] = CLng(1)  ' ** {Unassigned}.
17440                           End If
17450                           blnFound = False
17460                           For Each fld In .Fields
17470                             With fld
17480                               If .Name = "curr_id" Then
17490                                 blnFound = True
17500                                 rstLoc1![curr_id] = rstLnk![curr_id]
17510                                 Exit For
17520                               End If
17530                             End With
17540                           Next
17550                           If blnFound = False Then
17560                             rstLoc1![curr_id] = 150&  ' ** Default to USD.
                                  'CHECK AGAINST MASTER ASSET?
17570                           End If
17580                           blnFound = False
17590                           For Each fld In .Fields
17600                             With fld
17610                               If .Name = "cost_usd" Then
17620                                 blnFound = True
17630                                 rstLoc1![cost_usd] = rstLnk![cost_usd]
17640                                 Exit For
17650                               End If
17660                             End With
17670                           Next
17680                           If blnFound = False Then
17690                             If rstLoc1![curr_id] = 150& Then
17700                               rstLoc1![cost_usd] = ![Cost]
17710                             Else
                                    'WE'LL HAVE TO CHECK ALL THESE ONCE tblCurrency_History IS CONVERTED!
17720                             End If
17730                           End If
17740                           blnFound = False
17750                           For Each fld In .Fields
17760                             With fld
17770                               If .Name = "market_usd" Then
17780                                 blnFound = True
17790                                 rstLoc1![market_usd] = rstLnk![market_usd]
17800                                 Exit For
17810                               End If
17820                             End With
17830                           Next
17840                           If blnFound = False Then
                                  'WE'LL HAVE TO CHECK ALL THESE ONCE tblCurrency_History IS CONVERTED!
17850                           End If
                                ' ** I'm hoping the various contortions above will prevent a problem!
                                'BECAUSE OF THE ASSETDATE'S UNDERLYING PRECISION, 2 DATES WHICH
                                'MAY BE DIFFERENT IN THEIR DOUBLE VALUE COULD DISPLAY THE SAME!
                                'THEN, BECAUSE WE FORMAT, AND THEN CONVERT BACK TO DATE,
                                'THEY WILL COME OUT THE SAME HERE, CAUSING AN ERROR!
                                'HOW CAN WE HANDLE THAT?
                                'WE CAN'T SIMPLY CHANGE THEIR TIMESTAMPS HERE!
                                'WE WOULD HAVE TO FIND THE TRANSACTIONS ASSOCIATED WITH EACH,
                                'THEN CHANGE ONE SET BOTH HERE AND IN THE LEDGER!
                                'IS THIS PIECE OF CODE BEFORE OR AFTER THE LEDGER?
                                'BEFORE!
                                'SO, WE COULD SAVE THE 2 DIFFERENT DOUBLE VALUES HERE,
                                'THEN CHECK FOR THEM WHEN CONVERTING THE LEDGER, BELOW!
                                'HOW DIFFICULT WILL THAT BE?
                                'THE VARIABLE ARRAY, arr_varTmp01(), IS USED ABOVE FOR MASTER ASSET,
                                'BUT WE'RE DONE WITH IT BY THE TIME WE GET HERE!
                                'SO, WE COULD USE IT HERE, AND IT'D STILL BE AROUND FOR THE LEDGER, BELOW!
                                'CHECK LEDGER ARCHIVE AS WELL!
17860 On Error Resume Next
17870                           rstLoc1.Update
17880                           If ERR.Number <> 0 Then
17890                             If ERR.Number = 3022 Then  ' ** The changes you requested...
                                    ' ** We'll always assume it's the assetdate.
17900 On Error GoTo ERRH
17910                               lngTmp15 = lngTmp15 + 1&
17920                               lngE = lngTmp15 - 1&
17930                               ReDim Preserve arr_varTmp01(5, lngE)
17940                               arr_varTmp01(0, lngE) = rstLoc1![accountno]
17950                               arr_varTmp01(1, lngE) = rstLoc1![assetno]
17960                               arr_varTmp01(2, lngE) = rstLoc1![assetdate]        ' ** Old assetdate.
17970                               arr_varTmp01(3, lngE) = CDbl(rstLoc1![assetdate])  ' ** Old assetdate as Double.
17980                               datTmp28 = CDate(Format(rstLoc1![assetdate], "mm/dd/yyyy hh:nn:ss"))
17990                               datTmp28 = DateAdd("s", 17, datTmp28)  ' ** Arbitrary.
18000                               arr_varTmp01(4, lngE) = datTmp28                   ' ** New assetdate.
18010                               arr_varTmp01(5, lngE) = CBool(False)
18020                               rstLoc1![assetdate] = datTmp28
18030                               rstLoc1.Update  ' ** Let it error normally if it wasn't fixed.
18040                             Else
18050                               Beep
18060                               MsgBox "ERROR: " & CStr(ERR.Number) & vbCrLf & ERR.description, _
                                      vbCritical + vbOKOnly, "Error: " & CStr(ERR.Number)
18070 On Error GoTo ERRH
18080                               blnContinue = False
18090                               intRetVal = -9
18100                             End If
18110                           Else
18120 On Error GoTo ERRH
18130                           End If
18140                           If blnContinue = True Then
                                  ' ** The key field doesn't change, so no need to put it in tblVersion_Key.
18150                           Else
18160                             Exit For
18170                           End If
18180                           strTmp04 = vbNullString: strTmp08 = vbNullString
18190                           If lngX < lngRecs Then .MoveNext
18200                         Next
18210                         rstLoc1.Close
18220                         rstLoc2.Close
18230                       End If  ' ** Records present.
18240                       .Close
18250                     End With  ' ** rstLnk.
18260                   End If  ' ** blnFound.

18270                   lngStats = lngStats + 1&
18280                   lngE = lngStats - 1&
18290                   ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
18300                   arr_varStat(STAT_ORD, lngE) = CInt(5)
18310                   arr_varStat(STAT_NAM, lngE) = "Tax Lots: "
18320                   arr_varStat(STAT_CNT, lngE) = CLng(lngRecs)
18330                   arr_varStat(STAT_DSC, lngE) = vbNullString

18340                 End If  ' ** blnContinue.
18350                 strTmp04 = vbNullString: strTmp08 = vbNullString
18360                 strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString

18370               End With  ' ** TrustDta.mdb: dbsLnk.

18380             End If  ' ** dbsLnk opens.

18390           End With  ' ** wrkLnk.

18400         End If  ' ** Workspace opens: blnContinue.

18410       End If  ' ** blnConvert_TrustDta.

18420       If blnContinue = False Then
18430         dbsLoc.Close
18440         wrkLoc.Close
18450       End If

18460     End If  ' ** Conversion not already done.

18470   End If  ' ** Is a conversion.

18480   lngTmp01 = lngInvestObjs
18490   arr_varTmp02 = arr_varInvestObj

18500   lngTmp03 = lngStats
18510   arr_varTmp04 = arr_varStat

18520   lngTmp05 = lngDupeUnks
18530   arr_varTmp06 = arr_varDupeUnk

EXITP:
18540   Set fld = Nothing
18550   Set rstLnk = Nothing
18560   Set rstLoc1 = Nothing
18570   Set rstLoc2 = Nothing
18580   Set rstLoc3 = Nothing
18590   Set qdf = Nothing
18600   Set dbs = Nothing
18610   Version_Upgrade_04 = intRetVal
18620   Exit Function

ERRH:
18630   intRetVal = -9
18640   DoCmd.Hourglass False
18650   lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
18660   Select Case ERR.Number
        Case Else
18670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18680   End Select
18690   Resume EXITP

End Function

Public Function Version_Upgrade_08(blnContinue As Boolean, blnConvert_TrustDta As Boolean, lngTrustDtaDbsID As Long, strKeyTbl As String, dblPB_ThisStep As Double, lngOldTbls As Long, arr_varOldTbl As Variant, lngAccts As Long, arr_varAcct As Variant, lngStats As Long, arr_varTmp03 As Variant, wrkLoc As DAO.Workspace, wrkLnk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database) As Integer
' ** This handles the new miscellaneous tables.
' ** Tables converted here:
' **   tblCheckMemo
' **   tblCheckReconcile_Account
' **   tblCheckReconcile_Item
' **   tblCheckPOSPay
' **   tblCheckPOSPay_Detail
' **   tblCheckVoid
' **   tblCheckBank
' **   tblRecurringAux1099  (not used yet)
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

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Upgrade_08"

        Dim rstLnk As DAO.Recordset, rstLoc1 As DAO.Recordset, rstLoc2 As DAO.Recordset
        Dim arr_varStat() As Variant
        Dim lngItems As Long, arr_varItem() As Variant
        Dim strCurrTblName As String, lngCurrTblID As Long, lngCurrKeyFldID As Long
        Dim lngRecs As Long
        Dim blnFound As Boolean, blnAdd As Boolean
        Dim strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long, lngTmp14 As Long
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim intRetVal As Integer

        ' ** Array: arr_varStat().
        Const STAT_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const STAT_ORD As Integer = 0
        Const STAT_NAM As Integer = 1
        Const STAT_CNT As Integer = 2
        Const STAT_DSC As Integer = 3

        ' ** Array: arr_varItem().
        Const I_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const I_FNAM  As Integer = 0
        Const I_LNKID As Integer = 1
        Const I_LOCID As Integer = 2

18710   If gblnDev_NoErrHandle = True Then
18720 On Error GoTo 0
18730   End If

18740   intRetVal = 0
18750   lngRecs = 0&

18760   If blnContinue = True Then  ' ** Is a conversion.

18770     If blnContinue = True Then  ' ** Conversion not already done.

18780       lngTmp14 = 0&
18790       ReDim arr_varStat(STAT_ELEMS, 0)

            ' ** Stats are for anomalies, the stuff appearing after Statement Date on log file.
18800       For lngX = 0& To (lngStats - 1&)
18810         lngTmp14 = lngTmp14 + 1&
18820         lngE = lngTmp14 - 1&
18830         ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
18840         For lngZ = 0& To STAT_ELEMS
18850           arr_varStat(lngZ, lngE) = arr_varTmp03(lngZ, lngX)
18860         Next  ' ** lngZ.
18870       Next  ' ** lngX.

18880       If blnConvert_TrustDta = True Then

18890         If blnContinue = True Then  ' ** Workspace opens.

18900           With wrkLnk

18910             If blnContinue = True Then  ' ** Open dbsLnk.

18920               With dbsLnk

18930                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' *******************************************
                        ' ** Table: tblCheckMemo.
                        ' *******************************************

                        ' ** Step 26: tblCheckMemo.
18940                   dblPB_ThisStep = 26#
18950                   Version_Status 3, dblPB_ThisStep, "tblCheckMemo"  ' ** Module Function: modVersionConvertFuncs1.

18960                   strCurrTblName = "tblCheckMemo"
18970                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")  ' ** This doesn't seem to be used anywhere.

                        ' ** See if this new table is in the to-be-converted TrustDta.mdb.
18980                   blnFound = False: lngRecs = 0&
18990                   For lngX = 0& To (lngOldTbls - 1&)
19000                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
19010                       blnFound = True
19020                       Exit For
19030                     End If
19040                   Next

19050                   If blnFound = True Then
19060                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
19070                     With rstLnk
19080                       If .BOF = True And .EOF = True Then
                              ' ** Is anyone using this?
19090                       Else

19100                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
                              ' ** No earlier versions have this table.
                              ' ** Table: tblCheckMemo
                              ' **   ![ChkMemo_ID]            AutoNumber
                              ' **   ![ChkMemoType_Type]
                              ' **   ![ChkMemo_Memo]
                              ' **   ![Username]
                              ' **   ![ChkMemo_DateModified]
                              ' ** This comes with 1 record, "This is a Check Memo".
19110                         .MoveLast
19120                         lngRecs = .RecordCount
19130                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
19140                         .MoveFirst
19150                         If lngRecs = 1& And ![ChkMemo_Memo] = "This is a Check Memo" Then
                                ' ** Skip it!
19160                         Else
                                'rstLoc1.MoveFirst
19170                           For lngX = 1& To lngRecs
19180                             blnAdd = False
19190                             If lngX = 1& Then
19200                               If rstLoc1.BOF = True And rstLoc1.EOF = True Then
19210                                 blnAdd = True
19220                               Else
19230                                 If ![ChkMemo_ID] = 1& Then
19240                                   If ![ChkMemo_Memo] <> "This is a Check Memo" Then
19250                                     rstLoc1.Edit
19260                                     rstLoc1![ChkMemoType_Type] = ![ChkMemoType_Type]
19270                                     rstLoc1![ChkMemo_Memo] = ![ChkMemo_Memo]
19280                                     rstLoc1![Username] = ![Username]
19290                                     rstLoc1![ChkMemo_DateModified] = ![ChkMemo_DateModified]
19300                                     rstLoc1.Update
19310                                   Else
                                          ' ** Skip it!
19320                                   End If
19330                                 Else
19340                                   blnAdd = True
19350                                 End If
19360                               End If
19370                             End If
19380                             If blnAdd = True Then
                                    ' ** Add this record to the new table.
19390                               rstLoc1.AddNew
                                    ' ** rstLoc1![ChkMemo_ID] : AutoNumber
19400                               rstLoc1![ChkMemoType_Type] = ![ChkMemoType_Type]
19410                               rstLoc1![ChkMemo_Memo] = ![ChkMemo_Memo]
19420                               rstLoc1![Username] = ![Username]
19430                               rstLoc1![ChkMemo_DateModified] = ![ChkMemo_DateModified]
19440                               rstLoc1.Update
19450                             End If
19460                             If lngX < lngRecs Then .MoveNext
19470                           Next  ' ** lngX.
19480                         End If
19490                         rstLoc1.Close

19500                       End If  ' ** BOF, EOF.
19510                       .Close
19520                     End With  ' ** rstLnk.
19530                     Set rstLnk = Nothing
19540                   End If  ' ** blnFound.
19550                   DoEvents

                        ' *******************************************
                        ' ** Table: tblCheckReconcile_Account.
                        ' *******************************************

                        ' ** Step 27: tblCheckReconcile_Account.
19560                   dblPB_ThisStep = 27#
19570                   Version_Status 3, dblPB_ThisStep, "tblCheckReconcile_Account"  ' ** Module Function: modVersionConvertFuncs1.

19580                   strCurrTblName = "tblCheckReconcile_Account"
19590                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

                        ' ** See if this new table is in the to-be-converted TrustDta.mdb.
19600                   blnFound = False: lngRecs = 0&: strTmp01 = vbNullString
19610                   For lngX = 0& To (lngOldTbls - 1&)
19620                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
19630                       blnFound = True
19640                       Exit For
19650                     End If
19660                   Next

19670                   If blnFound = True Then
19680                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
19690                     With rstLnk
19700                       If .BOF = True And .EOF = True Then
                              ' ** Haven't used it.
19710                       Else

19720                         lngItems = 0&
19730                         ReDim arr_varItem(I_ELEMS, 0)

19740                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
19750                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** This has been in since v2.0.
                              ' ** Table: tblCheckReconcile_Account
                              ' **   ![cracct_id]            AutoNumber
                              ' **   ![accountno]
                              ' **   ![assetno]
                              ' **   ![cracct_date]
                              ' **   ![cracct_bsbalance]
                              ' **   ![cracct_datemodified]
                              ' ** This makes use of a phantom accountno of 'CRTC01'.
19760                         .MoveLast
19770                         lngRecs = .RecordCount
19780                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
19790                         .MoveFirst
                              'rstLoc1.MoveFirst  'THIS ONE'S EMPTY! OF COURSE IT ERRORED!!
19800                         For lngX = 1& To lngRecs
19810                           strTmp01 = ![accountno]
19820                           blnFound = False
                                ' ** Find the accountno in the master list.
19830                           For lngY = 0& To (lngAccts - 1&)
19840                             If arr_varAcct(A_NUM, lngY) = strTmp01 Then
19850                               blnFound = True
19860                               Exit For
19870                             End If
19880                           Next
19890                           If blnFound = True Then
19900                             lngItems = lngItems + 1&
19910                             lngE = lngItems - 1&
19920                             ReDim Preserve arr_varItem(I_ELEMS, lngE)
19930                             arr_varItem(I_FNAM, lngE) = "cracct_id"
19940                             arr_varItem(I_LNKID, lngE) = ![cracct_id]
19950                             arr_varItem(I_LOCID, lngE) = Null
                                  ' ** Add this record to the new table.
19960                             rstLoc1.AddNew
                                  ' ** rstLoc1![cracct_id] : AutoNumber
19970                             rstLoc1![accountno] = ![accountno]
                                  ' ** Find the assetno in the key table.
19980                             If IsNull(![assetno]) = False Then
19990                               If ![assetno] > 0& Then
20000                                 rstLoc2.MoveFirst
20010                                 rstLoc2.FindFirst "[tbl_name] = 'masterasset' And [fld_name] = 'assetno' And " & _
                                        "[key_lng_id1] = " & CStr(![assetno])
20020                                 If rstLoc2.NoMatch = True And ![assetno] = 1& Then
                                        ' ** It's the 'Accrued Interest Asset', masterasset_TYPE = 'IA',
                                        ' ** which doesn't get moved.
20030                                   rstLoc1![assetno] = 1&
20040                                 ElseIf rstLoc2.NoMatch = False Then
20050                                   rstLoc1![assetno] = rstLoc2![key_lng_id2] 'Req  0
20060                                 Else
                                        ' ** It may be an orphan, but it'll have to remain one!
20070                                   rstLoc1![assetno] = ![assetno]
20080                                 End If
20090                               Else
20100                                 rstLoc1![assetno] = 0&
20110                               End If
20120                             Else
20130                               rstLoc1![assetno] = 0&
20140                             End If
20150                             rstLoc1![cracct_date] = ![cracct_date]
20160                             rstLoc1![cracct_bsbalance] = ![cracct_bsbalance]
20170                             rstLoc1![cracct_datemodified] = ![cracct_datemodified]
20180                             rstLoc1.Update
20190                             rstLoc1.Bookmark = rstLoc1.LastModified
20200                             arr_varItem(I_LOCID, lngE) = rstLoc1![cracct_id]
20210                           End If
20220                           If lngX < lngRecs Then .MoveNext
20230                         Next
20240                         rstLoc1.Close
20250                         rstLoc2.Close

20260                       End If  ' ** BOF, EOF.
20270                       .Close
20280                     End With  ' ** rstLnk.
20290                     Set rstLnk = Nothing
20300                   End If  ' ** blnFound.
20310                   DoEvents

                        ' *******************************************
                        ' ** Table: tblCheckReconcile_Item.
                        ' *******************************************

                        ' ** Step 28: tblCheckReconcile_Item.
20320                   dblPB_ThisStep = 28#
20330                   Version_Status 3, dblPB_ThisStep, "tblCheckReconcile_Item"  ' ** Module Function: modVersionConvertFuncs1.

20340                   strCurrTblName = "tblCheckReconcile_Item"
20350                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

                        ' ** See if this new table is in the to-be-converted TrustDta.mdb.
20360                   blnFound = False: lngRecs = 0&: lngTmp02 = 0&: lngTmp03 = 0&
20370                   For lngX = 0& To (lngOldTbls - 1&)
20380                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
20390                       blnFound = True
20400                       Exit For
20410                     End If
20420                   Next

20430                   If blnFound = True Then
20440                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
20450                     With rstLnk
20460                       If .BOF = True And .EOF = True Then
                              ' ** Haven't used it.
20470                       Else
20480                         If lngItems > 0& Then  ' ** We need the ID cross-reference.

20490                           Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
                                ' ** This has been in since v2.0.
                                ' ** Table: tblCheckReconcile_Item
                                ' **   ![cracct_id]
                                ' **   ![critem_id]            AutoNumber
                                ' **   ![accountno]
                                ' **   ![assetno]
                                ' **   ![crsource_type]
                                ' **   ![crentry_type]
                                ' **   ![critem_description]
                                ' **   ![critem_amount]
                                ' **   ![critem_datemodified]
                                ' ** arr_varItem() is used to coordinate cracct_id.
20500                           .MoveLast
20510                           lngRecs = .RecordCount
20520                           Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
20530                           .MoveFirst
                                'rstLoc1.MoveFirst
20540                           For lngX = 1& To lngRecs
20550                             lngTmp02 = ![cracct_id]
20560                             lngTmp03 = 0&
20570                             blnFound = False
20580                             For lngY = 0& To (lngItems - 1&)
20590                               If arr_varItem(I_LNKID, lngY) = lngTmp02 Then
20600                                 blnFound = True
20610                                 lngTmp03 = arr_varItem(I_LOCID, lngY)
20620                                 Exit For
20630                               End If
20640                             Next
20650                             If blnFound = True Then
20660                               rstLoc1.AddNew
20670                               rstLoc1![cracct_id] = lngTmp03
                                    ' ** rstLoc1![critem_id] : AutoNumber
20680                               rstLoc1![accountno] = ![accountno]  ' ** Vetted with parent record.
20690                               rstLoc1![assetno] = ![assetno]  ' ** Vetted with parent record.
20700                               rstLoc1![crsource_type] = ![crsource_type]  ' ** These types aren't likely to ever change.
20710                               rstLoc1![crentry_type] = ![crentry_type]
20720                               rstLoc1![critem_description] = ![critem_description]
20730                               rstLoc1![critem_amount] = ![critem_amount]
20740                               rstLoc1![critem_datemodified] = ![critem_datemodified]
20750                               rstLoc1.Update
20760                             End If  ' ** blnFound
20770                             If lngX < lngRecs Then .MoveNext
20780                           Next  ' ** lngX.
20790                           rstLoc1.Close

20800                         End If  ' ** lngItems.
20810                       End If  ' ** BOF, EOF.
20820                       .Close
20830                     End With  ' ** rstLnk.
20840                     Set rstLnk = Nothing
20850                   End If  ' ** blnFound.
20860                   DoEvents

                        ' *******************************************
                        ' ** Table: tblCheckPOSPay.
                        ' *******************************************

                        ' ** Step 29: tblCheckPOSPay.
20870                   dblPB_ThisStep = 29#
20880                   Version_Status 3, dblPB_ThisStep, "tblCheckPOSPay"  ' ** Module Function: modVersionConvertFuncs1.

20890                   strCurrTblName = "tblCheckPOSPay"
20900                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

                        ' ** See if this new table is in the to-be-converted TrustDta.mdb.
20910                   blnFound = False: lngRecs = 0&
20920                   For lngX = 0& To (lngOldTbls - 1&)
20930                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
20940                       blnFound = True
20950                       Exit For
20960                     End If
20970                   Next

20980                   If blnFound = True Then
20990                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
21000                     With rstLnk
21010                       If .BOF = True And .EOF = True Then
                              ' ** Haven't used it.
21020                       Else

21030                         lngItems = 0&
21040                         ReDim arr_varItem(I_ELEMS, 0)

21050                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
                              ' ** This is new to v2.2.24.
                              ' ** Table: tblCheckPOSPay
                              ' **   ![pp_id]            AutoNumber
                              ' **   ![pp_date]
                              ' **   ![pp_description]
                              ' **   ![pp_pathfile]
                              ' **   ![pp_checks]
                              ' **   ![Username]
                              ' **   ![pp_user]
                              ' **   ![pp_datemodified]
                              ' ** arr_varItem() weill be used to coordinate pp_id.
21060                         .MoveLast
21070                         lngRecs = .RecordCount
21080                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
21090                         .MoveFirst
                              'rstLoc1.MoveFirst
21100                         For lngX = 1& To lngRecs
21110                           lngItems = lngItems + 1&
21120                           lngE = lngItems - 1&
21130                           ReDim Preserve arr_varItem(I_ELEMS, lngE)
21140                           arr_varItem(I_FNAM, lngE) = "pp_id"
21150                           arr_varItem(I_LNKID, lngE) = ![pp_id]
21160                           arr_varItem(I_LOCID, lngE) = Null
21170                           rstLoc1.AddNew
                                ' ** ![pp_id] : AutoNumber
21180                           rstLoc1![pp_date] = ![pp_date]
21190                           rstLoc1![pp_description] = ![pp_description]
21200                           rstLoc1![pp_pathfile] = ![pp_pathfile]
21210                           rstLoc1![pp_checks] = ![pp_checks]
21220                           rstLoc1![Username] = ![Username]
21230                           rstLoc1![pp_user] = ![pp_user]
21240                           rstLoc1![pp_datemodified] = ![pp_datemodified]
21250                           rstLoc1.Update
21260                           rstLoc1.Bookmark = rstLoc1.LastModified
21270                           arr_varItem(I_LOCID, lngE) = rstLoc1![pp_id]
21280                           If lngX < lngRecs Then .MoveNext
21290                         Next  ' ** lngX.
21300                         rstLoc1.Close

21310                       End If  ' ** BOF, EOF.
21320                       .Close
21330                     End With  ' ** rstLnk.
21340                     Set rstLnk = Nothing
21350                   End If  ' ** blnFound.
21360                   DoEvents

                        ' *******************************************
                        ' ** Table: tblCheckPOSPay_Detail.
                        ' *******************************************

                        ' ** Step 30: tblCheckPOSPay_Detail.
21370                   dblPB_ThisStep = 30#
21380                   Version_Status 3, dblPB_ThisStep, "tblCheckPOSPay_Detail"  ' ** Module Function: modVersionConvertFuncs1.

21390                   strCurrTblName = "tblCheckPOSPay_Detail"
21400                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

                        ' ** See if this new table is in the to-be-converted TrustDta.mdb.
21410                   blnFound = False: lngRecs = 0&: lngTmp02 = 0&: lngTmp03 = 0&
21420                   For lngX = 0& To (lngOldTbls - 1&)
21430                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
21440                       blnFound = True
21450                       Exit For
21460                     End If
21470                   Next

21480                   If blnFound = True Then
21490                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
21500                     With rstLnk
21510                       If .BOF = True And .EOF = True Then
                              ' ** Haven't used it.
21520                       Else
21530                         If lngItems > 0& Then  ' ** We need the ID cross-reference.

21540                           Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
21550                           Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                                ' ** This is new to v2.2.24.
                                ' ** Table: tblCheckPOSPay_Detail
                                ' **   ![pp_id]
                                ' **   ![ppd_id]            AutoNumber
                                ' **   ![Journal_ID]
                                ' **   ![journalno]
                                ' **   ![ppd_checknum]
                                ' **   ![ppd_issuedate]
                                ' **   ![accountno]
                                ' **   ![ppd_amount]
                                ' **   ![ppd_payee]
                                ' **   ![RecurringItem_ID]
                                ' **   ![ppd_bank_name]
                                ' **   ![ppd_aba_trc]
                                ' **   ![ppd_bank_account]
                                ' **   ![curr_id]
                                ' **   ![ppd_void]
                                ' **   ![Username]
                                ' **   ![ppd_user]
                                ' **   ![ppd_datemodified]
                                ' ** arr_varItem() is used to coordinate pp_id.
21560                           .MoveLast
21570                           lngRecs = .RecordCount
21580                           Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
21590                           .MoveFirst
                                'rstLoc1.MoveFirst
21600                           For lngX = 1& To lngRecs
21610                             lngTmp02 = ![pp_id]
21620                             lngTmp03 = 0&
21630                             blnFound = False
21640                             For lngY = 0& To (lngItems - 1&)
21650                               If arr_varItem(I_LNKID, lngY) = lngTmp02 Then
21660                                 blnFound = True
21670                                 lngTmp03 = arr_varItem(I_LOCID, lngY)
21680                                 Exit For
21690                               End If
21700                             Next
21710                             If blnFound = True Then
21720                               rstLoc1.AddNew
21730                               rstLoc1![pp_id] = lngTmp03
                                    ' ** rstLoc1![ppd_id] : AutoNumber
21740                               rstLoc1![Journal_ID] = ![Journal_ID]
21750                               rstLoc1![journalno] = ![journalno]  ' ** These don't get renumbered, so it should be there.
21760                               rstLoc1![ppd_checknum] = ![ppd_checknum]
21770                               rstLoc1![ppd_issuedate] = ![ppd_issuedate]
21780                               rstLoc1![accountno] = ![accountno]  ' ** I don't think it really matters if it's here or not.
21790                               rstLoc1![ppd_amount] = ![ppd_amount]
21800                               rstLoc1![ppd_payee] = ![ppd_payee]
21810                               If IsNull(![RecurringItem_ID]) = False Then
21820                                 If ![RecurringItem_ID] > 0& Then
21830                                   rstLoc2.MoveFirst
21840                                   rstLoc2.FindFirst "[tbl_name] = 'RecurringItems' And [fld_name] = 'RecurringItem_ID' And " & _
                                          "[key_lng_id1] = " & CStr(![RecurringItem_ID])
21850                                   If rstLoc2.NoMatch = True And (![RecurringItem_ID] = 1& Or ![RecurringItem_ID] = 2&) Then
                                          ' ** It's one of the 2 defaults, which don't get moved.
21860                                     rstLoc1![RecurringItem_ID] = ![RecurringItem_ID]
21870                                   ElseIf rstLoc2.NoMatch = False Then
21880                                     rstLoc1![RecurringItem_ID] = rstLoc2![key_lng_id2]
21890                                   Else
                                          ' ** It's an orphan, nix it!
21900                                     rstLoc1![RecurringItem_ID] = Null
21910                                   End If
21920                                 Else
21930                                   rstLoc1![RecurringItem_ID] = Null
21940                                 End If
21950                               Else
21960                                 rstLoc1![RecurringItem_ID] = Null
21970                               End If
21980                               rstLoc1![ppd_bank_name] = ![ppd_bank_name]
21990                               rstLoc1![ppd_aba_trc] = ![ppd_aba_trc]
22000                               rstLoc1![ppd_bank_account] = ![ppd_bank_account]
22010                               rstLoc1![curr_id] = ![curr_id]  ' ** Unlikely to be gone.
22020                               rstLoc1![ppd_void] = ![ppd_void]
22030                               rstLoc1![Username] = ![Username]
22040                               rstLoc1![ppd_user] = ![ppd_user]
22050                               rstLoc1![ppd_datemodified] = ![ppd_datemodified]
22060                               rstLoc1.Update
22070                             End If  ' ** blnFound
22080                             If lngX < lngRecs Then .MoveNext
22090                           Next  ' ** lngX.
22100                           rstLoc1.Close
22110                           rstLoc2.Close

22120                         End If  ' ** lngItems.
22130                       End If  ' ** BOF, EOF.
22140                       .Close
22150                     End With  ' ** rstLnk.
22160                     Set rstLnk = Nothing
22170                   End If  ' ** blnFound.
22180                   DoEvents

                        ' *******************************************
                        ' ** Table: tblCheckBank.
                        ' *******************************************

                        ' ** Step 31: tblCheckBank.
22190                   dblPB_ThisStep = 31#
22200                   Version_Status 3, dblPB_ThisStep, "tblCheckBank"  ' ** Module Function: modVersionConvertFuncs1.

22210                   strCurrTblName = "tblCheckBank"
22220                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

                        ' ** See if this new table is in the to-be-converted TrustDta.mdb.
22230                   blnFound = False: lngRecs = 0&: lngTmp02 = 0&: lngTmp03 = 0&
22240                   For lngX = 0& To (lngOldTbls - 1&)
22250                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
22260                       blnFound = True
22270                       Exit For
22280                     End If
22290                   Next

22300                   If blnFound = True Then
22310                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
22320                     With rstLnk
22330                       If .BOF = True And .EOF = True Then
                              ' ** Haven't used it.
22340                       Else

22350                         lngItems = 0&
22360                         ReDim arr_varItem(I_ELEMS, 0)

22370                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
                              ' ** New to this release of v2.2.24.
                              ' ** Table: tblCheckBank
                              ' **   ![chkbank_id]            AutoNumber
                              ' **   ![chkbank_name]
                              ' **   ![chkbank_acctnum]
                              ' **   ![accountno]
                              ' **   ![chkbank_active]
                              ' **   ![chkbank_datemodified]
                              ' ** arr_varItem() weill be used to coordinate chkbank_id.
22380                         .MoveLast
22390                         lngRecs = .RecordCount
22400                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
22410                         .MoveFirst
                              'rstLoc1.MoveFirst
22420                         For lngX = 1& To lngRecs
22430                           lngItems = lngItems + 1&
22440                           lngE = lngItems - 1&
22450                           ReDim Preserve arr_varItem(I_ELEMS, lngE)
22460                           arr_varItem(I_FNAM, lngE) = "chkbank_id"
22470                           arr_varItem(I_LNKID, lngE) = ![chkbank_id]
22480                           arr_varItem(I_LOCID, lngE) = Null
22490                           rstLoc1.AddNew
                                ' ** rstLoc1![chkbank_id] : AutoNumber
22500                           rstLoc1![chkbank_name] = ![chkbank_name]
22510                           rstLoc1![chkbank_acctnum] = ![chkbank_acctnum]
22520                           rstLoc1![accountno] = ![accountno]  ' ** As note, above.
22530                           rstLoc1![chkbank_active] = ![chkbank_active]
22540                           rstLoc1![chkbank_datemodified] = ![chkbank_datemodified]
22550                           rstLoc1.Update
22560                           rstLoc1.Bookmark = rstLoc1.LastModified
22570                           arr_varItem(I_LOCID, lngE) = rstLoc1![chkbank_id]
22580                           If lngX < lngRecs Then .MoveNext
22590                         Next  ' ** lngX.
22600                         rstLoc1.Close

22610                       End If  ' ** BOF, EOF.
22620                       .Close
22630                     End With  ' ** rstLnk.
22640                     Set rstLnk = Nothing
22650                   End If  ' ** blnFound.
22660                   DoEvents

                        ' *******************************************
                        ' ** Table: tblCheckVoid.
                        ' *******************************************

                        ' ** Step 32: tblCheckVoid.
22670                   dblPB_ThisStep = 32#
22680                   Version_Status 3, dblPB_ThisStep, "tblCheckVoid"  ' ** Module Function: modVersionConvertFuncs1.

22690                   strCurrTblName = "tblCheckVoid"
22700                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

                        ' ** See if this new table is in the to-be-converted TrustDta.mdb.
22710                   blnFound = False: lngRecs = 0&: lngTmp02 = 0&: lngTmp03 = 0&
22720                   For lngX = 0& To (lngOldTbls - 1&)
22730                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
22740                       blnFound = True
22750                       Exit For
22760                     End If
22770                   Next

22780                   If blnFound = True Then
22790                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
22800                     With rstLnk
22810                       If .BOF = True And .EOF = True Then
                              ' ** Haven't used it.
22820                       Else
22830                         If lngItems > 0& Then

22840                           Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
                                ' ** New to this release of v2.2.24.
                                ' ** Table: tblCheckVoid
                                ' **   ![chkvoid_id]            AutoNumber
                                ' **   ![chkbank_id]
                                ' **   ![chkbank_name]
                                ' **   ![chkbank_acctnum]
                                ' **   ![chkvoid_chknum]
                                ' **   ![chkvoid_date]
                                ' **   ![accountno]
                                ' **   ![transdate]
                                ' **   ![chkvoid_payee]
                                ' **   ![chkvoid_amount]
                                ' **   ![curr_id]
                                ' **   ![chkvoid_set]
                                ' **   ![journal_id]
                                ' **   ![chkvoid_datemodified]
                                ' ** arr_varItem() is used to coordinate pp_id.
22850                           .MoveLast
22860                           lngRecs = .RecordCount
22870                           Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Module Function: modVersionConvertFuncs1.
22880                           .MoveFirst
                                'rstLoc1.MoveFirst
22890                           For lngX = 1& To lngRecs
22900                             lngTmp02 = ![chkbank_id]
22910                             lngTmp03 = 0&
22920                             blnFound = False
22930                             For lngY = 0& To (lngItems - 1&)
22940                               If arr_varItem(I_LNKID, lngY) = lngTmp02 Then
22950                                 blnFound = True
22960                                 lngTmp03 = arr_varItem(I_LOCID, lngY)
22970                                 Exit For
22980                               End If
22990                             Next
23000                             If blnFound = True Then
23010                               rstLoc1.AddNew
                                    ' ** rstLoc1~![chkvoid_id] : AutoNumber
23020                               rstLoc1![chkbank_id] = lngTmp03
23030                               rstLoc1![chkbank_name] = ![chkbank_name]
23040                               rstLoc1![chkbank_acctnum] = ![chkbank_acctnum]
23050                               rstLoc1![chkvoid_chknum] = ![chkvoid_chknum]
23060                               rstLoc1![chkvoid_date] = ![chkvoid_date]
23070                               rstLoc1![accountno] = ![accountno]
23080                               rstLoc1![transdate] = ![transdate]
23090                               rstLoc1![chkvoid_payee] = ![chkvoid_payee]
23100                               rstLoc1![chkvoid_amount] = ![chkvoid_amount]
23110                               rstLoc1![curr_id] = ![curr_id]
23120                               rstLoc1![chkvoid_set] = ![chkvoid_set]
23130                               rstLoc1![Journal_ID] = ![Journal_ID]
23140                               rstLoc1![chkvoid_datemodified] = ![chkvoid_datemodified]
23150                               rstLoc1.Update
23160                             End If  ' ** blnFound
23170                             If lngX < lngRecs Then .MoveNext
23180                           Next  ' ** lngX.
23190                           rstLoc1.Close

23200                         End If  ' ** lngItems.
23210                       End If  ' ** BOF, EOF.
23220                       .Close
23230                     End With  ' ** rstLnk.
23240                     Set rstLnk = Nothing
23250                   End If  ' ** blnFound.
23260                   DoEvents

                        ' *******************************************
                        ' ** Table: tblRecurringAux1099.
                        ' *******************************************

                        ' ** Step 33: tblRecurringAux1099.
23270                   dblPB_ThisStep = 33#
23280                   Version_Status 3, dblPB_ThisStep, "tblRecurringAux1099"  ' ** Module Function: modVersionConvertFuncs1.

                        ' ** Not in use yet.

23290                 End If  ' ** dbsLoc is still open.

23300                 .Close
23310               End With  ' ** TrustDta.mdb: dbsLnk.

23320             End If  ' ** Open dbsLnk.

23330             .Close
23340           End With  ' ** wrkLnk.

23350         End If  ' ** Workspace opens.

23360       End If  ' ** blnConvert_TrustDta.

23370       dbsLoc.Close
23380       wrkLoc.Close

23390       If lngTmp14 > lngStats Then
23400         lngStats = lngTmp14
23410         arr_varTmp03 = arr_varStat
23420       End If

23430     End If  ' ** Conversion not already done.

23440   End If  ' ** Is a conversion.

23450   DoCmd.Hourglass False

EXITP:
23460   Set rstLnk = Nothing
23470   Set rstLoc1 = Nothing
23480   Set rstLoc2 = Nothing
23490   Version_Upgrade_08 = intRetVal
23500   Exit Function

ERRH:
23510   intRetVal = -9
23520   DoCmd.Hourglass False
23530   lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
23540   ErrInfo_Set lngErrNum, lngErrLine, strErrDesc  ' ** Module Procedure: modVersionConvertFuncs1.
23550   Select Case ERR.Number
        Case Else
23560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23570   End Select
23580   Resume EXITP

End Function

Public Function ResetConvertNew() As Boolean

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "ResetConvertNew"

        Dim fso As Scripting.FileSystemObject, fsfd As Scripting.Folder, fsfls As Scripting.FILES, fsfl As Scripting.File
        Dim strPath As String, strFile As String, strPathFile1 As String, strPathFile2 As String
        Dim lngFiles As Long, arr_varFile() As Variant
        Dim lngFilesRenamed As Long, blnHasLog As Boolean
        Dim intPos01 As Integer
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varFile().
        Const F_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const F_PATH As Integer = 0
        Const F_FIL1 As Integer = 1
        Const F_EXT1 As Integer = 2
        Const F_FIL2 As Integer = 3
        Const F_EXT2 As Integer = 4

23610   blnRetVal = True
23620   blnHasLog = False

23630   strPath = CurrentBackendPath  ' ** Module Function: modFileUtilities.
23640   strPath = strPath & LNK_SEP & gstrDir_Convert

23650   Set fso = CreateObject("Scripting.FileSystemObject")
23660   With fso

23670     Set fsfd = .GetFolder(strPath)
23680     Set fsfls = fsfd.FILES

23690     lngFiles = 0&
23700     ReDim arr_varFile(F_ELEMS, 0)

23710     For Each fsfl In fsfls
23720       With fsfl
23730         intPos01 = InStr(.Name, ".")
23740         If Mid(.Name, intPos01) = ".BAK" Then
                ' ** A converted TA data file.
23750           lngFiles = lngFiles + 1&
23760           lngE = lngFiles - 1&
23770           ReDim Preserve arr_varFile(F_ELEMS, lngE)
23780           arr_varFile(F_PATH, lngE) = .Path
23790           arr_varFile(F_FIL1, lngE) = Rem_Ext(.Name)  ' ** Module Function: modStringFuncs.
23800           arr_varFile(F_EXT1, lngE) = Parse_Ext(.Name)  ' ** Module Function: modStringFuncs.
23810           arr_varFile(F_FIL2, lngE) = Null
23820           arr_varFile(F_EXT2, lngE) = Null
23830         ElseIf .Name = gstrFile_ConvertLog Then
23840           blnHasLog = True
23850         End If
23860       End With
23870     Next
23880   End With

23890   Debug.Print "'FILES: " & CStr(lngFiles)
23900   DoEvents

23910   If lngFiles > 0& Then

23920     For lngX = 0& To (lngFiles - 1&)
23930       intPos01 = InStr(arr_varFile(F_FIL1, lngX), "_")
23940       If intPos01 > 0 Then
              ' ** TrustSec_v2224.BAK.
23950         arr_varFile(F_FIL2, lngX) = Left(arr_varFile(F_FIL1, lngX), (intPos01 - 1))
23960       Else
23970         arr_varFile(F_FIL2, lngX) = arr_varFile(F_FIL1, lngX)
23980       End If
23990       Select Case arr_varFile(F_FIL2, lngX)
            Case gstrFile_App  ' ** Does not include extension.
              ' ** This will assume they're the same.
24000         arr_varFile(F_EXT2, lngX) = CurrentAppExt  ' ** Module Function: modFileUtilities.
24010       Case Rem_Ext(gstrFile_DataName), Rem_Ext(gstrFile_ArchDataName)  ' ** Module Function: modStringFuncs.
24020         arr_varFile(F_EXT2, lngX) = gstrExt_AppDev
24030       Case Rem_Ext(gstrFile_SecurityName)  ' ** Module Function: modStringFuncs.
24040         arr_varFile(F_EXT2, lngX) = gstrExt_AppSec
24050       End Select
24060     Next  ' **  lngX.

24070     lngFilesRenamed = 0&
24080     For lngX = 0& To (lngFiles - 1&)
24090       If IsNull(arr_varFile(F_FIL2, lngX)) = False And IsNull(arr_varFile(F_EXT2, lngX)) = False Then
24100         strPathFile1 = strPath & LNK_SEP & arr_varFile(F_FIL1, lngX) & "." & arr_varFile(F_EXT1, lngX)
24110         strPathFile2 = strPath & LNK_SEP & arr_varFile(F_FIL2, lngX) & "." & arr_varFile(F_EXT2, lngX)
24120         Name strPathFile1 As strPathFile2
24130         lngFilesRenamed = lngFilesRenamed + 1&
24140       End If
24150     Next  ' **  lngX.

24160     If blnHasLog = True Then
24170       strPathFile1 = strPath & LNK_SEP & gstrFile_ConvertLog
24180       Kill strPathFile1
24190     End If

24200   End If  ' ** lngFiles.

24210   Debug.Print "'FILES RENAMEED: " & CStr(lngFilesRenamed)
24220   DoEvents

24230   Beep

24240   Debug.Print "'DONE!"
24250   DoEvents

EXITP:
24260   Set fsfl = Nothing
24270   Set fsfls = Nothing
24280   Set fsfd = Nothing
24290   Set fso = Nothing
24300   ResetConvertNew = blnRetVal
24310   Exit Function

ERRH:
24320   blnRetVal = False
24330   Select Case ERR.Number
        Case Else
24340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24350   End Select
24360   Resume EXITP

End Function

Public Function Version_Input1_EnterData() As Boolean

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Input1_EnterData"

        Dim frm As Access.Form
        Dim blnRetVal As Boolean

24410   blnRetVal = True

24420   Set frm = Forms("frmVersion_Input")
24430   With frm
24440     .CoInfo_Name = "North Fork Bank"
24450     .CoInfo_Address1 = "Oak Plaza"
24460     .CoInfo_Address2 = "100 Oak Street"
24470     .CoInfo_City = "North Fork"
24480     .CoInfo_State = "MN"
24490     .CoInfo_Zip = "551145123"
24500     .CoInfo_Country = Null
24510     .CoInfo_PostalCode = Null
24520     .CoInfo_Phone = "612-334-7800"
24530     .chkIncomeTaxCoding = True
24540     .chkRevenueExpenseTracking = True
24550     .chkSeparateCheckingAccounts = True
24560     .chkTabCopy = True
24570   End With

24580   Beep

EXITP:
24590   Set frm = Nothing
24600   Version_Input1_EnterData = blnRetVal
24610   Exit Function

ERRH:
24620   blnRetVal = False
24630   Select Case ERR.Number
        Case Else
24640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24650   End Select
24660   Resume EXITP

End Function

Public Function AccessVer(varInput As Variant) As Variant
' ** This returns a string for a given version number.
' ** Application.Version

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "AccessVer"

        Dim lngTmp01 As Long
        Dim varRetVal As Variant

24710   varRetVal = Null

24720   If IsNull(varInput) = False Then
24730     If IsNumeric(varInput) = True Then
24740       lngTmp01 = Val(varInput)
24750       Select Case lngTmp01
            Case 1#
24760         varRetVal = "Access 1.1"
24770       Case 2#
24780         varRetVal = "Access 2.0"
24790       Case 7#
24800         varRetVal = "Access for Windows 95"
24810       Case 8#
24820         varRetVal = "Access 97"
24830       Case 9#
24840         varRetVal = "Access 2000"
24850       Case 10#
24860         varRetVal = "Access 2002"
24870       Case 11#
24880         varRetVal = "Access 2003"
24890       Case 12#
24900         varRetVal = "Access 2007"
24910       Case 14#
24920         varRetVal = "Access 2010"
24930       Case 15#
24940         varRetVal = "Access 2013"
24950       Case 16#
24960         varRetVal = "Access 2016"
24970       End Select
24980     End If
24990   End If

EXITP:
25000   AccessVer = varRetVal
25010   Exit Function

ERRH:
25020   varRetVal = RET_ERR
25030   Select Case ERR.Number
        Case Else
25040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25050   End Select
25060   Resume EXITP

End Function
