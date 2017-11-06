Attribute VB_Name = "modQueryFunctions1"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modQueryFunctions1"

'VGC 09/05/2017: CHANGES!

'gstrFormQuerySpec
'gstrReportQuerySpec

'totdesc = ([description] & IIf([rate]>0,"  " & Format([rate],"#,##0.000%"),"") & IIf(IsNull([due])=True,"","  Due " & Format([due],"mm/dd/yyyy")))
'totdesc: ([description] & IIf([rate]>0,"  " & Format([rate],"#,##0.000%"),"") & IIf(IsNull([due])=True,"","  Due " & Format([due],"mm/dd/yyyy")))

' ** lbxShortAccountName list box constants:
'Private Const LBX_CHK_ID     As Integer = 0  ' ** check_id
Private Const LBX_CHK_ACTNO  As Integer = 1  ' ** Account_Number
'Private Const LBX_CHK_SNAME  As Integer = 2  ' ** Short_Name

' ** cmbMonth combo box constants:
Private Const CBX_MON_ID    As Integer = 0  ' ** month_id (same as month number)
'Private Const CBX_MON_NAME  As Integer = 1  ' ** month_name
Private Const CBX_MON_SHORT As Integer = 2  ' ** month_short
' **

Public Function FormRef(Optional varItem As Variant) As Variant
' ** This function was created so that queries need
' ** not specify direct references to specific forms.
' ** If a form name changes or a field name changes, the
' ** mismatch should show up here when the application is compiled.
' ** NO! CHANGED TO THE Forms(strFormName) SYNTAX!!
' **
' ** NOTE: The Map parameters are way over-covered! Only 1 for each of the 2 reports should actually hit.

        ' **************************************************************************************
        ' ** NOTE:
        ' ** The public variable 'gstrReportQuerySpec' needs to be initialized prior to calling.
        ' **************************************************************************************

100   On Error GoTo ERRH

        Const THIS_PROC As String = "FormRef"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim intPos01 As Integer
        Dim varTmp00 As Variant, strTmp01 As String, lngTmp02 As Long, blnTmp03 As Boolean
        Dim lngX As Long
        Dim varRetVal As Variant

        Static lngOddUps As Long, arr_varOddUp As Variant

110     varRetVal = Null

        ' ***********************************************************
        ' ** Reports.
        ' ***********************************************************
120     If gstrReportQuerySpec <> vbNullString Then

130       Select Case gstrReportQuerySpec

          Case "rptAccountSummary", "rptAccountSummary_ForEx"
            ' ** Called from frmStatementParameters.
140         strTmp01 = vbNullString
150         If IsNull(varItem) = True Then
              ' ** Stay here.
160         Else
170           Select Case varItem
              Case "frmStatementParameters", "rptAccountSummary", "rptAccountSummary_ForEx"
                ' ** Stay here.
180           Case Else
                ' ** Meant for gstrFormQuerySpec!
190             strTmp01 = gstrReportQuerySpec
200             varTmp00 = varItem
210             gstrReportQuerySpec = vbNullString
220             varRetVal = FormRef(varTmp00) ' ** Recursive calling of this function.
230             gstrReportQuerySpec = strTmp01
240           End Select
250         End If
260         If strTmp01 = vbNullString Then
270           If IsLoaded(gstrReportQuerySpec, acReport) = True Then  ' ** Module Function: modFileUtilities.
280             varRetVal = Forms("frmStatementParameters").cmbAccounts  'Reports![rptAccountSummary].accountno
290           Else
300             varRetVal = "SUSPENSE"  ' ** Arbitrary.
310           End If
320         End If
330       Case "rptPricingUpdateCusips_txt", "rptPricingUpdateCusips_prn"
340         If IsMissing(varItem) = False Then
350           Select Case varItem
              Case "OldDate"
360             varRetVal = CDate(Forms("frmAssetPricing").Date_Current)
370           Case "NewDate"
380             varRetVal = CDate(Forms("frmAssetPricing").Date_New)
390           End Select
400         End If
410       Case "rptMap_Reinvest_Div", "rptMap_Reinvest_Int"
            ' ** total_shareface: Sum([journal map].[icash]/[Forms]![frmMap_Reinvest_DivInt_Price]![txtprice])
420         If IsLoaded(gstrReportCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
430           If Forms(gstrReportCallingForm).CurrentView <> acCurViewDesign Then
                ' ** Borrowing this variable from the Court Reports.
440             If gdblCrtRpt_PrinTot <> 0# Then
450               varRetVal = gdblCrtRpt_PrinTot
460             Else
470               varRetVal = CCur(Forms("frmMap_Reinvest_DivInt_Detail").txtPrice)
480             End If
490           Else
500             varRetVal = 1@
510           End If
520         Else
530           varRetVal = 1@
540         End If
550       Case "rptMap_Reinvest_Rec"
            ' ** 'total_shareface: [total_pcash]/[Forms]![frmMap_Reinvest_Rec_Price]![txtprice]*-1
560         If IsLoaded(gstrReportCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
570           If Forms(gstrReportCallingForm).CurrentView <> acCurViewDesign Then
                ' ** Borrowing this variable from the Court Reports.
580             If gdblCrtRpt_PrinTot <> 0# Then
590               varRetVal = gdblCrtRpt_PrinTot
600             Else
610               varRetVal = CCur(Forms("frmMap_Reinvest_Rec_Detail").txtPrice)
620             End If
630           Else
640             varRetVal = 1@
650           End If
660         Else
670           varRetVal = 1@
680         End If
690       End Select

          ' ***********************************************************
          ' ** Forms.
          ' ***********************************************************
700     ElseIf gstrFormQuerySpec <> vbNullString Then

710       Select Case gstrFormQuerySpec

          Case "frmAccountHideTrans2"
720         If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
730           If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
740             If IsMissing(varItem) = False Then
750               Select Case varItem
                  Case "PriorPeriod"
760                 varRetVal = gdatEndDate
770               Case "AcctNum"
780                 varRetVal = gstrAccountNo
790               Case "EndDate"
800                 varRetVal = gdatEndDate
810               Case Else
                    ' ** A UniqueID is sent, to return its group number.
820                 varRetVal = Hide_Group(varItem)  ' ** Module Function: modHideTransactions1.
830               End Select
840             End If
850           Else
860             Select Case varItem
                Case "PriorPeriod"
870               varRetVal = gdatEndDate
880             Case "AcctNum"
890               varRetVal = gstrAccountNo
900             Case "EndDate"
910               varRetVal = gdatEndDate
920             Case Else
                  ' ** A UniqueID is sent, to return its group number.
930               varRetVal = Hide_Group(varItem)  ' ** Module Function: modHideTransactions1.
940             End Select
950           End If
960         Else
970           Select Case varItem
              Case "PriorPeriod"
980             varRetVal = gdatEndDate
990           Case "AcctNum"
1000            varRetVal = gstrAccountNo
1010          Case "EndDate"
1020            varRetVal = gdatEndDate
1030          Case Else
                ' ** A UniqueID is sent, to return its group number.
1040            varRetVal = Hide_Group(varItem)  ' ** Module Function: modHideTransactions1.
1050          End Select
1060        End If
1070      Case "frmAccountHideTrans2_Hidden"
1080        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
1090          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
1100            If IsMissing(varItem) = False Then
1110              Select Case varItem
                  Case "PriorPeriod"
1120                varRetVal = gdatEndDate
1130              Case "AcctNum", "frmAccountHideTrans", "frmAccountHideTrans2"
1140                varRetVal = gstrAccountNo
1150              Case "EndDate"
1160                varRetVal = gdatEndDate
1170              Case Else
                    ' ** A UniqueID is sent, to return its group number.
1180                varRetVal = Hide_Group(varItem)  ' ** Module Function: modHideTransactions1.
1190              End Select
1200            End If
1210          Else
1220            If IsMissing(varItem) = False Then
1230              Select Case varItem
                  Case "PriorPeriod"
1240                varRetVal = gdatEndDate
1250              Case "AcctNum"
1260                varRetVal = gstrAccountNo
1270              Case "EndDate"
1280                varRetVal = gdatEndDate
1290              Case Else
                    ' ** A UniqueID is sent, to return its group number.
1300                varRetVal = Hide_Group(varItem)  ' ** Module Function: modHideTransactions1.
1310              End Select
1320            End If
1330          End If
1340        ElseIf IsLoaded("frmAccountHideTrans2", acForm) = True Then  ' ** Module Function: modFileUtilities.
1350          If IsMissing(varItem) = False Then
1360            Select Case varItem
                Case "PriorPeriod"
1370              varRetVal = gdatEndDate
1380            Case "AcctNum"
1390              varRetVal = gstrAccountNo
1400            Case "EndDate"
1410              varRetVal = gdatEndDate
1420            Case Else
                  ' ** A UniqueID is sent, to return its group number.
1430              varRetVal = Hide_Group(varItem)  ' ** Module Function: modHideTransactions1.
1440            End Select
1450          End If
1460        End If
1470      Case "frmAccountHideTrans2_One"
1480        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
1490          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
1500            If IsMissing(varItem) = False Then
1510              Select Case varItem
                  Case "AcctNum"
1520                varRetVal = gstrAccountNo
1530              End Select
1540            End If
1550          Else
1560            Select Case varItem
                Case "AcctNum"
1570              varRetVal = gstrAccountNo
1580            End Select
1590          End If
1600        Else
1610          Select Case varItem
              Case "AcctNum"
1620            varRetVal = gstrAccountNo
1630          End Select
1640        End If
1650      Case "frmAccountHideTrans2_Select"
1660        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
1670          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
1680            If IsMissing(varItem) = False Then
1690              Select Case varItem
                  Case "AcctNum"
1700                varRetVal = gstrAccountNo
1710              End Select
1720            End If
1730          Else
1740            Select Case varItem
                Case "AcctNum"
1750              varRetVal = gstrAccountNo
1760            End Select
1770          End If
1780        Else
1790          Select Case varItem
              Case "AcctNum"
1800            varRetVal = gstrAccountNo
1810          End Select
1820        End If
1830      Case "frmAccountProfile"
1840        If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
1850          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
1860            varRetVal = Forms(gstrFormQuerySpec).frmAccountProfile_Sub.Form.accountno
1870          Else
1880            varRetVal = "00001"
1890          End If
1900        Else
1910          varRetVal = "00001"
1920        End If
1930      Case "frmCheckPOSPay"
1940        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
1950          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
1960            Select Case varItem
                Case "ppid"
1970              varRetVal = Forms(gstrFormQuerySpec).pp_id
1980            End Select
1990          Else
2000            Select Case varItem
                Case "ppid"
2010              varRetVal = 8&
2020            End Select
2030          End If
2040        End If
2050      Case "frmCheckReconcile"
2060        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
2070          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
2080            Select Case varItem
                Case "AccountNo"
2090              If IsNull(Forms(gstrFormQuerySpec).cmbAccounts) = False Then
2100                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
2110              Else
2120                varRetVal = vbNullString
2130              End If
2140            Case "AssetNo"
2150              If IsNull(Forms(gstrFormQuerySpec).cmbAssets) = False Then
2160                varRetVal = Forms(gstrFormQuerySpec).cmbAssets
2170              Else
2180                varRetVal = CLng(0)
2190              End If
2200            Case "PStartDate"
2210              If IsNull(Forms(gstrFormQuerySpec).DateStart) = False Then
2220                varRetVal = Forms(gstrFormQuerySpec).DateStart
2230              Else
2240                varRetVal = CDate("01/01/1900")
2250              End If
2260            Case "PEndDate"
2270              If IsNull(Forms(gstrFormQuerySpec).DateEnd) = False Then
2280                varRetVal = Forms(gstrFormQuerySpec).DateEnd
2290              Else
2300                varRetVal = Date
2310              End If
2320            Case "CRAcctID"
2330              If IsNull(Forms(gstrFormQuerySpec).cracct_id_tmp) = False Then
2340                varRetVal = Forms(gstrFormQuerySpec).cracct_id_tmp
2350              Else
2360                varRetVal = CLng(0)
2370              End If
2380            End Select
2390          Else
2400            Select Case varItem
                Case "AccountNo"
2410              varRetVal = "SUSPENSE"
2420            Case "AssetNo"
2430              varRetVal = CLng(0)
2440            Case "PStartDate"
2450              varRetVal = CDate("01/01/1900")
2460            Case "PEndDate"
2470              varRetVal = Date
2480            Case "CRAcctID"
2490              varRetVal = CLng(0)
2500            End Select
2510          End If
2520        Else
2530          Select Case varItem
              Case "AccountNo"
2540            varRetVal = "CRTC01"
2550          Case "AssetNo"
2560            varRetVal = CLng(0)
2570          Case "PStartDate"
2580            varRetVal = CDate("01/01/1900")
2590          Case "PEndDate"
2600            varRetVal = Date
2610          Case "CRAcctID"
2620            varRetVal = CLng(0)
2630          End Select
2640        End If
2650      Case "frmFeeCalculations"
2660        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
2670          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
2680            If IsMissing(varItem) = False Then
2690              Select Case varItem
                  Case "FeeFreq"
2700                varRetVal = Forms(gstrFormQuerySpec).FeeFreq
2710              End Select
2720            End If
2730          End If
2740        End If
2750      Case "frmFeeSchedules_Detail_Sub"
2760        If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
2770          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
2780            varRetVal = Forms("frmFeeSchedules").Schedule_ID
2790          Else
2800            varRetVal = CLng(1)
2810          End If
2820        Else
2830          varRetVal = CLng(1)
2840        End If
2850      Case "frmJournal"
2860        If IsMissing(varItem) = False Then
2870          Select Case varItem
              Case "frmJournal_Sub1_Dividend"
                ' ** [Forms]![frmJournal]![frmJournal_Sub1_Dividend]![dividendAccountNo]
2880            If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
2890              If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
2900                varRetVal = Forms(gstrFormQuerySpec).frmJournal_Sub1_Dividend.Form.dividendAccountNo
2910              Else
2920                varRetVal = "11" 'DLookup("[accountno]", "account")  ' ** Default to first one.
2930              End If
2940            Else
2950              varRetVal = "11" 'DLookup("[accountno]", "account")  ' ** Default to first one.
2960            End If
2970          Case "frmJournal_Sub2_Interest"
                ' ** [Forms]![frmJournal]![frmJournal_Sub2_Interest]![interestAccountNo]
2980            If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
2990              If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
3000                varRetVal = Forms(gstrFormQuerySpec).frmJournal_Sub2_Interest.Form.interestAccountNo
3010              Else
3020                varRetVal = DLookup("[accountno]", "account")  ' ** Default to first one.
3030              End If
3040            Else
3050              varRetVal = DLookup("[accountno]", "account")  ' ** Default to first one.
3060            End If
3070          Case "frmJournal_Sub3_Purchase"
3080            If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
3090              If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
3100                Select Case IsNull(Forms(gstrFormQuerySpec).Controls(varItem).Form.purchaseType)
                    Case True
3110                  varRetVal = CBool(False)
3120                Case False
3130                  If Trim(Forms(gstrFormQuerySpec).Controls(varItem).Form.purchaseType) <> vbNullString Then
3140                    strTmp01 = Forms(gstrFormQuerySpec).Controls(varItem).Form.purchaseType
3150                    Select Case strTmp01
                        Case "Liability", "Liability (+)"
3160                      varRetVal = CBool(True)
3170                    Case "Deposit", "Purchase"
3180                      varRetVal = CBool(False)
3190                    End Select
3200                  Else
3210                    varRetVal = CBool(False)
3220                  End If
3230                End Select
3240              Else
3250                varRetVal = CBool(False)
3260              End If
3270            End If
3280          Case "AccountNo"
                ' ** [Forms]![frmJournal]![frmJournal_Sub4_Sold]![saleAccountNo]
3290            If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
3300              If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
3310                Select Case IsNull(Forms(gstrFormQuerySpec).frmJournal_Sub4_Sold.Form.saleAccountNo)
                    Case True
3320                  varRetVal = DLookup("[accountno]", "account")  ' ** Default to first one.
3330                Case False
3340                  If Trim(Forms(gstrFormQuerySpec).frmJournal_Sub4_Sold.Form.saleAccountNo) <> vbNullString Then
3350                    varRetVal = Forms(gstrFormQuerySpec).frmJournal_Sub4_Sold.Form.saleAccountNo
3360                  Else
3370                    varRetVal = DLookup("[accountno]", "account")  ' ** Default to first one.
3380                  End If
3390                End Select
3400              Else
3410                varRetVal = DLookup("[accountno]", "account")  ' ** Default to first one.
3420              End If
3430            End If
3440          Case "AssetJType"
3450            If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
3460              If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
3470                Select Case IsNull(Forms(gstrFormQuerySpec).frmJournal_Sub4_Sold.Form.saleType)
                    Case True
3480                  varRetVal = 0  ' ** All Asset Types.
3490                Case False
3500                  If Trim(Forms(gstrFormQuerySpec).frmJournal_Sub4_Sold.Form.saleType) <> vbNullString Then
3510                    strTmp01 = Forms(gstrFormQuerySpec).frmJournal_Sub4_Sold.Form.saleType
3520                    Select Case strTmp01
                        Case "Cost Adj."
3530                      varRetVal = 0  ' ** All Asset Types.
3540                    Case "Liability", "Liability (-)"
3550                      varRetVal = 1
3560                    Case "Withdrawn", "Sold"
3570                      varRetVal = 2
3580                    End Select
3590                  Else
3600                    varRetVal = 0  ' ** All Asset Types.
3610                  End If
3620                End Select
3630              Else
3640                varRetVal = 0  ' ** All Asset Types.
3650              End If
3660            End If
3670          Case "Price"
3680            If IsLoaded("frmMap_Reinvest_DivInt_Price", acForm) = True Then  ' ** Module Function: modFileUtilities.
3690              If Forms("frmMap_Reinvest_DivInt_Price").CurrentView <> acCurViewDesign Then
                    ' ** Borrowing this variable from the Court Reports.
3700                If gdblCrtRpt_PrinTot <> 0# Then
3710                  varRetVal = gdblCrtRpt_PrinTot
3720                Else
3730                  varRetVal = CCur(Forms("frmMap_Reinvest_DivInt_Price").txtPrice)
3740                End If
3750              Else
3760                varRetVal = CCur(1)
3770              End If
3780            ElseIf IsLoaded("frmMap_Reinvest_DivInt_Detail", acForm) = True Then  ' ** Module Function: modFileUtilities.
3790              If Forms("frmMap_Reinvest_DivInt_Detail").CurrentView <> acCurViewDesign Then
                    ' ** Borrowing this variable from the Court Reports.
3800                varRetVal = gdblCrtRpt_PrinTot
3810              Else
3820                varRetVal = CCur(1)
3830              End If
3840            ElseIf IsLoaded("frmMap_Reinvest_Rec_Price", acForm) = True Then  ' ** Module Function: modFileUtilities.
3850              If Forms("frmMap_Reinvest_Rec_Price").CurrentView <> acCurViewDesign Then
                    ' ** Borrowing this variable from the Court Reports.
3860                If gdblCrtRpt_PrinTot <> 0# Then
3870                  varRetVal = gdblCrtRpt_PrinTot
3880                Else
3890                  varRetVal = CCur(Forms("frmMap_Reinvest_Rec_Price").txtPrice)
3900                End If
3910              Else
3920                varRetVal = CCur(1)
3930              End If
3940            ElseIf IsLoaded("frmMap_Reinvest_Rec_Detail", acForm) = True Then  ' ** Module Function: modFileUtilities.
3950              If Forms("frmMap_Reinvest_Rec_Detail").CurrentView <> acCurViewDesign Then
                    ' ** Borrowing this variable from the Court Reports.
3960                varRetVal = gdblCrtRpt_PrinTot
3970              Else
3980                varRetVal = CCur(1)
3990              End If
4000            Else
4010              varRetVal = CCur(1)
4020            End If
4030          Case "SingleUser"
4040            varRetVal = gblnSingleUser
4050          End Select
4060        Else
              ' ** This source requires an item string.
4070        End If
4080      Case "frmJournal_Columns", "frmJournal_Columns_Sub"
4090        If IsMissing(varItem) = False Then
4100          If IsLoaded("frmJournal_Columns", acForm) = True Then  ' ** Module Functions: modFileUtilities.
4110            Select Case varItem
                Case "accountno"
4120              varRetVal = Forms("frmJournal_Columns").frmJournal_Columns_Sub.Form.accountno
4130            Case "Recur_Type"
                  ' ** Misc, Payee, Payor
4140              varRetVal = Forms("frmJournal_Columns").frmJournal_Columns_Sub.Form.Recur_Type
4150            Case "SingleUser"
4160              varRetVal = gblnSingleUser
4170            Case "JUser"
4180              varRetVal = gstrJournalUser
4190            Case "Price"
4200              If IsLoaded("frmMap_Reinvest_DivInt_Price", acForm) = True Then  ' ** Module Function: modFileUtilities.
4210                If Forms("frmMap_Reinvest_DivInt_Price").CurrentView <> acCurViewDesign Then
                      ' ** Borrowing this variable from the Court Reports.
4220                  If gdblCrtRpt_PrinTot <> 0# Then
4230                    varRetVal = gdblCrtRpt_PrinTot
4240                  Else
4250                    varRetVal = CCur(Forms("frmMap_Reinvest_DivInt_Price").txtPrice)
4260                  End If
4270                Else
4280                  varRetVal = CCur(1)
4290                End If
4300              ElseIf IsLoaded("frmMap_Reinvest_DivInt_Detail", acForm) = True Then  ' ** Module Function: modFileUtilities.
4310                If Forms("frmMap_Reinvest_DivInt_Detail").CurrentView <> acCurViewDesign Then
                      ' ** Borrowing this variable from the Court Reports.
4320                  varRetVal = gdblCrtRpt_PrinTot
4330                Else
4340                  varRetVal = CCur(1)
4350                End If
4360              ElseIf IsLoaded("frmMap_Reinvest_Rec_Price", acForm) = True Then  ' ** Module Function: modFileUtilities.
4370                If Forms("frmMap_Reinvest_Rec_Price").CurrentView <> acCurViewDesign Then
                      ' ** Borrowing this variable from the Court Reports.
4380                  If gdblCrtRpt_PrinTot <> 0# Then
4390                    varRetVal = gdblCrtRpt_PrinTot
4400                  Else
4410                    varRetVal = CCur(Forms("frmMap_Reinvest_Rec_Price").txtPrice)
4420                  End If
4430                Else
4440                  varRetVal = CCur(1)
4450                End If
4460              ElseIf IsLoaded("frmMap_Reinvest_Rec_Detail", acForm) = True Then  ' ** Module Function: modFileUtilities.
4470                If Forms("frmMap_Reinvest_Rec_Detail").CurrentView <> acCurViewDesign Then
                      ' ** Borrowing this variable from the Court Reports.
4480                  varRetVal = gdblCrtRpt_PrinTot
4490                Else
4500                  varRetVal = CCur(1)
4510                End If
4520              Else
4530                varRetVal = CCur(1)
4540              End If
4550            End Select
4560          Else
                ' ** Dev only.
4570            Select Case varItem
                Case "accountno"
4580              varRetVal = "11"
4590            Case "Recur_Type"
4600              varRetVal = "Misc"
4610            End Select
4620          End If
4630        End If
4640      Case "frmJournal_Columns_TaxLot"
4650        If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Functions: modFileUtilities.
4660          If IsMissing(varItem) = False Then
4670            Select Case varItem
                Case "accountno"
4680              varRetVal = Forms(gstrFormQuerySpec).accountno
4690            Case "assetno"
4700              varRetVal = Forms(gstrFormQuerySpec).assetno
4710            Case "AssetDateNull"
4720              varRetVal = Forms(gstrFormQuerySpec).AssetDateNull
4730            Case "AssetDateSale"
4740              varRetVal = Forms(gstrFormQuerySpec).AssetDateSale
4750            End Select
4760          End If
4770        Else
              ' ** Dev only.
4780        End If
4790      Case "frmMap_Reinvest_DivInt_Price", "frmMap_Reinvest_DivInt_Detail"
4800        If IsLoaded("frmMap_Reinvest_DivInt_Price", acForm) = True Then  ' ** Module Function: modFileUtilities.
4810          If Forms("frmMap_Reinvest_DivInt_Price").CurrentView <> acCurViewDesign Then
                ' ** Borrowing this variable from the Court Reports.
4820            If gdblCrtRpt_PrinTot <> 0# Then
4830              varRetVal = gdblCrtRpt_PrinTot
4840            Else
4850              varRetVal = CCur(Forms("frmMap_Reinvest_DivInt_Price").txtPrice)
4860            End If
4870          Else
4880            varRetVal = CCur(1)
4890          End If
4900        ElseIf IsLoaded("frmMap_Reinvest_DivInt_Detail", acForm) = True Then  ' ** Module Function: modFileUtilities.
4910          If Forms("frmMap_Reinvest_DivInt_Detail").CurrentView <> acCurViewDesign Then
                ' ** Borrowing this variable from the Court Reports.
4920            varRetVal = gdblCrtRpt_PrinTot
4930          Else
4940            varRetVal = CCur(1)
4950          End If
4960        Else
4970          varRetVal = CCur(1)
4980        End If
4990      Case "frmMap_Reinvest_Rec_Price", "frmMap_Reinvest_Rec_Detail"
5000        If IsLoaded("frmMap_Reinvest_Rec_Price", acForm) = True Then  ' ** Module Function: modFileUtilities.
5010          If Forms("frmMap_Reinvest_Rec_Price").CurrentView <> acCurViewDesign Then
                ' ** Borrowing this variable from the Court Reports.
5020            If gdblCrtRpt_PrinTot <> 0# Then
5030              varRetVal = gdblCrtRpt_PrinTot
5040            Else
5050              varRetVal = CCur(Forms("frmMap_Reinvest_Rec_Price").txtPrice)
5060            End If
5070          Else
5080            varRetVal = CCur(1)
5090          End If
5100        ElseIf IsLoaded("frmMap_Reinvest_Rec_Detail", acForm) = True Then  ' ** Module Function: modFileUtilities.
5110          If Forms("frmMap_Reinvest_Rec_Detail").CurrentView <> acCurViewDesign Then
                ' ** Borrowing this variable from the Court Reports.
5120            varRetVal = gdblCrtRpt_PrinTot
5130          Else
5140            varRetVal = CCur(1)
5150          End If
5160        Else
5170          varRetVal = CCur(1)
5180        End If
5190      Case "frmMasterBalance"
5200        If IsMissing(varItem) = False Then
5210          Select Case varItem
              Case "DiscrepanciesOnly"
5220            varRetVal = Forms(gstrFormQuerySpec).chkDiscrepancies
5230          Case "gblnAccountNoWithType"
5240            varRetVal = gblnAccountNoWithType
5250          Case "DatBeg"
5260            varRetVal = "01/01/1900"
5270          Case "DatEnd"
5280            varRetVal = "01/01/2200"
5290          End Select
5300        Else
5310          varRetVal = "01/01/2200"
5320        End If
5330      Case "frmMenu_Post"
5340        If IsMissing(varItem) = False Then
5350          Select Case varItem
              Case "PrePost"
5360            varRetVal = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
5370          Case "SingleUser"
5380            varRetVal = gblnSingleUser
5390          Case "TransCnt"
5400            varRetVal = Forms(gstrFormQuerySpec).TransCnt
5410          Case "UncommitCnt"
5420            varRetVal = Forms(gstrFormQuerySpec).UncommitCnt
5430          Case "CheckCnt"
5440            varRetVal = Forms(gstrFormQuerySpec).CheckCnt
5450          End Select
5460        Else
5470          varRetVal = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
5480        End If
5490      Case "frmRpt_AccountProfile"
5500        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
5510          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
5520            If IsMissing(varItem) = False Then
5530              Select Case varItem
                  Case "accountno"
5540                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
5550              End Select
5560            End If
5570          Else
5580            If IsMissing(varItem) = False Then
5590              Select Case varItem
                  Case "accountno"
5600                varRetVal = "11"
5610              End Select
5620            End If
5630          End If
5640        End If
5650      Case "frmRpt_AccountReviews"
5660        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
5670          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
5680            If IsMissing(varItem) = False Then
5690              Select Case varItem
                  Case "AccountNo"
5700                varRetVal = Forms(gstrFormQuerySpec).accountno
5710              Case Else
5720                varRetVal = Forms(gstrFormQuerySpec).cmbMonth.Column(CBX_MON_SHORT)  ' ** 3-letter month.
5730              End Select
5740            Else
5750              varRetVal = Forms(gstrFormQuerySpec).cmbMonth.Column(CBX_MON_SHORT)  ' ** 3-letter month.
5760            End If
5770          Else
5780            If IsMissing(varItem) = False Then
5790              Select Case varItem
                  Case "AccountNo"
5800                varRetVal = "11"
5810              Case Else
5820                varRetVal = "jun"
5830              End Select
5840            Else
5850              varRetVal = "jun"
5860            End If
5870          End If
5880        End If
5890      Case "frmRpt_ArchivedTransactions"
5900        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
5910          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
5920            If IsMissing(varItem) = False Then
5930              Select Case varItem
                  Case "AccountNo"
5940                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
5950              Case "StartDate"
5960                varRetVal = CDate(Forms(gstrFormQuerySpec).DateStart)
5970              Case "EndDate"
5980                varRetVal = CDate(Forms(gstrFormQuerySpec).DateEnd)
5990              End Select
6000            End If
6010          Else
6020            If IsMissing(varItem) = False Then
6030              Select Case varItem
                  Case "AccountNo"
6040                varRetVal = "11"
6050              Case "StartDate"
6060                varRetVal = CDate("01/01/2004")
6070              Case "EndDate"
6080                varRetVal = Date
6090              End Select
6100            End If
6110          End If
6120        End If
6130      Case "frmRpt_AssetHistory"
6140        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
6150          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
6160            If IsMissing(varItem) = False Then
6170              Select Case varItem
                  Case "AccountNo"
6180                varRetVal = gstrAccountNo
6190              Case "AssetNo"
6200                varRetVal = Forms(gstrFormQuerySpec).cmbAssets
6210              Case "StartDate"
6220                varRetVal = CDate(Forms(gstrFormQuerySpec).DateStart)
6230              Case "PStartDate"
6240                varRetVal = CDate("01/01/1900")
6250              Case "EndDate"
6260                varRetVal = CDate(Forms(gstrFormQuerySpec).DateEnd)
6270              Case "PEndDate"
6280                If CDbl(CDate(Forms(gstrFormQuerySpec).DateStart)) <= 1# Then
                      ' ** Date given is earlier than 01/02/1900.
                      ' ** Microsoft and/or Intel cannot process dates earler than 01/01/100 (i.e., the year 100 CE/AD),
                      ' ** so, the DateAdd(), below, gets very confused, and errors with an Error 6, Overflow.
                      ' ** A date entered with only a 2-digit year is interpreted as within the 20th century,
                      ' ** e.g., 03/05/57 = 03/05/1957, NOT the year 57 CE/AD.
6290                  varRetVal = CDate(#1/1/100#)    ' ** Just return the year 100 CE/AD.
6300                Else
6310                  varRetVal = DateAdd("d", -1#, CDate(Forms(gstrFormQuerySpec).DateStart))
6320                End If
6330              Case "Archive"
6340                varRetVal = CBool(Forms(gstrFormQuerySpec).chkIncludeArchive)
6350              Case "PArchive"
6360                varRetVal = CBool(True)
6370              End Select
6380            Else
6390              varRetVal = CCur(0)
6400            End If
6410          Else
6420            If IsMissing(varItem) = False Then
6430              Select Case varItem
                  Case "AccountNo"
6440                varRetVal = "AI12080701"
6450              Case "AssetNo"
6460                varRetVal = 1155&
6470              Case "StartDate"
6480                varRetVal = CDate("01/01/1900")
6490              Case "PStartDate"
6500                varRetVal = CDate("01/01/1900")
6510              Case "EndDate"
6520                varRetVal = CDate("06/21/2016")
6530              Case "PEndDate"
6540                varRetVal = CDate("06/21/2016")
6550              Case "Archive"
6560                varRetVal = CBool(True)
6570              Case "PArchive"
6580                varRetVal = CBool(True)
6590              End Select
6600            Else
6610              varRetVal = CCur(0)
6620            End If
6630          End If
6640        Else
6650          If IsMissing(varItem) = False Then
6660            Select Case varItem
                Case "AccountNo"
6670              varRetVal = "00119"
6680            Case "AssetNo"
6690              varRetVal = 1&
6700            Case "StartDate"
6710              varRetVal = CDate("01/01/1900")
6720            Case "PStartDate"
6730              varRetVal = CDate("01/01/1900")
6740            Case "EndDate"
6750              varRetVal = CDate("10/31/2011")
6760            Case "PEndDate"
6770              varRetVal = CDate("10/31/2011")
6780            Case "Archive"
6790              varRetVal = CBool(False)
6800            Case "PArchive"
6810              varRetVal = CBool(False)
6820            End Select
6830          Else
6840            varRetVal = CCur(0)
6850          End If
6860        End If
6870      Case "frmRpt_CapitalGainAndLoss"
6880        If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
6890          If IsMissing(varItem) = False Then
6900            Select Case varItem
                Case "FromTo"
6910              If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
6920                varRetVal = "From " & Format(CDate(Forms(gstrFormQuerySpec).DateStart), "mm/dd/yyyy") & _
                      " To " & Format(CDate(Forms(gstrFormQuerySpec).DateEnd), "mm/dd/yyyy")
6930              Else
6940                varRetVal = "From 01/01/2000 To 12/31/2008"
6950              End If
6960            Case "AccountNo"
6970              varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
6980            Case "StartDate"
6990              varRetVal = CDate(Format(Forms(gstrFormQuerySpec).DateStart, "mm/dd/yyyy") & " 00:00:01")
7000            Case "StartDateLong"
7010              strTmp01 = CStr(CDbl(CDate(Forms(gstrFormQuerySpec).DateStart)))
7020              If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
7030              lngTmp02 = CLng(strTmp01)
7040              varRetVal = lngTmp02
7050            Case "EndDate"
7060              varRetVal = CDate(Format(Forms(gstrFormQuerySpec).DateEnd, "mm/dd/yyyy") & " 23:59:59")
7070            Case "EndDateLong"
7080              strTmp01 = CStr(CDbl(CDate(Forms(gstrFormQuerySpec).DateEnd)))
7090              If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
7100              lngTmp02 = CLng(strTmp01)
7110              varRetVal = lngTmp02
7120            End Select
7130          Else
7140            varRetVal = "From 01/01/2000 To 12/31/2008"
7150          End If
7160        Else
7170          Select Case varItem
              Case "FromTo"
7180            varRetVal = "From 01/01/2000 To 12/31/2008"
7190          Case "AccountNo"
7200            varRetVal = "11"
7210          Case "StartDate"
7220            varRetVal = CDate("01/01/2000 00:00:01")
7230          Case "StartDateLong"
7240            varRetVal = CLng(CDate("01/01/2000"))
7250          Case "EndDate"
7260            varRetVal = CDate("01/01/2008 23:59:59")
7270          Case "EndDateLong"
7280            varRetVal = CLng(CDate("01/01/2008"))
7290          Case Else
7300            varRetVal = "From 01/01/2000 To 12/31/2008"
7310          End Select
7320        End If
7330      Case "frmRpt_CashControl"
7340        If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
7350          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
7360            varRetVal = Format(Forms(gstrFormQuerySpec).DateAsOf, "mm/dd/yyyy")
7370          End If
7380        End If
7390      Case "frmRpt_Checks_Bank2"
7400        If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
7410          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
7420            If IsMissing(varItem) = False Then
7430              Select Case varItem
                  Case "chkvoid_set"
7440                varRetVal = Forms(gstrFormQuerySpec).chkvoid_set
7450              End Select
7460            End If
7470          End If
7480        End If
7490      Case "frmRpt_CourtReports_CA"
7500        If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
7510          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
7520            If IsMissing(varItem) = False Then
7530              Select Case varItem
                  Case "CRPT_ON_HAND_BEG"
7540                varRetVal = CCur(gdblCrtRpt_CA_POHBeg)
7550              Case "CRPT_ON_HAND_END"
7560                varRetVal = CCur(gdblCrtRpt_CA_POHEnd)
7570              Case "CRPT_Cash_BEG"
7580                varRetVal = CCur(gdblCrtRpt_CA_COHBeg)
7590              Case "CRPT_Cash_END"
7600                varRetVal = CCur(gdblCrtRpt_CA_COHEnd)
7610              Case "Period"
7620                Select Case Forms(gstrFormQuerySpec).chkAssetList_Start
                    Case False
7630                  varRetVal = "From " & Format(CDate(Forms(gstrFormQuerySpec).DateStart), "mm/dd/yyyy") & _
                        " To " & Format(CDate(Forms(gstrFormQuerySpec).DateEnd), "mm/dd/yyyy")
7640                Case True
7650                  varRetVal = "As of " & Format((CDate(Forms(gstrFormQuerySpec).DateStart) - 1), "mm/dd/yyyy")
7660                End Select
7670              Case "StartDate"
7680                varRetVal = CDate(Forms(gstrFormQuerySpec).DateStart)
7690              Case "StartDateLong"
7700                strTmp01 = CStr(CDbl(CDate(Forms(gstrFormQuerySpec).DateStart)))
7710                If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
7720                lngTmp02 = CLng(strTmp01)
7730                varRetVal = lngTmp02
7740              Case "EndDate"
7750                varRetVal = CDate(Forms(gstrFormQuerySpec).DateEnd)
7760              Case "EndDateLong"
7770                strTmp01 = CStr(CDbl(CDate(Forms(gstrFormQuerySpec).DateEnd)))
7780                If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
7790                lngTmp02 = CLng(strTmp01)
7800                varRetVal = lngTmp02
7810              Case "AccountNo"
7820                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
7830              Case "ShortName"
7840                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts.Column(3)
7850              Case "CaseNum"
7860                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts.Column(7)
7870              Case "OrdVer"
7880                varRetVal = Forms(gstrFormQuerySpec).Ordinal & " And " & Forms(gstrFormQuerySpec).Version
7890              Case "Title7"
7900                Select Case Forms(gstrFormQuerySpec).chkAssetList_Start
                    Case True
7910                  varRetVal = "Property on Hand at Beginning of Accounting Period"
7920                Case False
7930                  varRetVal = "Property on Hand at Close of Accounting Period - Schedule E"
7940                End Select
7950              Case "InvestInfo"
7960                varRetVal = CCur(gdblCrtRpt_CA_InvestInfo)
7970              Case "InvestChange"
7980                varRetVal = CCur(gdblCrtRpt_CA_InvestChange)
7990              End Select
8000            Else
8010              varRetVal = CCur(0)
8020            End If
8030          Else
8040            varRetVal = CCur(0)
8050          End If
8060        Else
8070          If GetUserName = gstrDevUserName Then
8080            blnTmp03 = False
8090            Select Case varItem
                Case "CRPT_ON_HAND_BEG"
8100              varRetVal = CCur(0)
8110            Case "CRPT_ON_HAND_END"
8120              varRetVal = CCur(0)
8130            Case "CRPT_Cash_BEG"
8140              varRetVal = CCur(0)
8150            Case "CRPT_Cash_END"
8160              varRetVal = CCur(0)
8170            Case "Period"
8180              Select Case blnTmp03
                  Case False
8190                varRetVal = "From " & Format(CDate("01/01/2008"), "mm/dd/yyyy") & _
                      " To " & Format(CDate("12/31/2008"), "mm/dd/yyyy")
8200              Case True
8210                varRetVal = "As of " & Format((CDate("01/01/2008") - 1), "mm/dd/yyyy")
8220              End Select
8230            Case "StartDate"
8240              varRetVal = CDate("01/01/2000")
8250            Case "StartDateLong"
8260              strTmp01 = CStr(CDbl(CDate("01/01/2000")))
8270              If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
8280              lngTmp02 = CLng(strTmp01)
8290              varRetVal = lngTmp02
8300            Case "EndDate"
8310              varRetVal = CDate("12/31/2014")
8320            Case "EndDateLong"
8330              strTmp01 = CStr(CDbl(CDate("12/31/2014")))
8340              If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
8350              lngTmp02 = CLng(strTmp01)
8360              varRetVal = lngTmp02
8370            Case "AccountNo"
8380              varRetVal = "00001"
8390            Case "Title7"
8400              Select Case blnTmp03
                  Case False
8410                varRetVal = "Property on Hand at Close of Account - Schedule E"
8420              Case True
8430                varRetVal = "Property on Hand at Beginning of Account"
8440              End Select
8450            Case "InvestInfo"
8460              varRetVal = CCur(0)
8470            Case "InvestChange"
8480              varRetVal = CCur(0)
8490            End Select
8500          Else
8510            varRetVal = CCur(0)
8520          End If
8530        End If
8540      Case "frmRpt_CourtReports_FL"
8550        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
8560          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
8570            If IsMissing(varItem) = False Then
8580              Select Case varItem
                  Case "CRPT_ON_HAND_BEG"
8590                varRetVal = CCur(Nz(Forms(gstrFormQuerySpec).PropOnHand_Beg, 0))
8600              Case "CRPT_ON_HAND_END"
8610                varRetVal = CCur(Nz(Forms(gstrFormQuerySpec).PropOnHand_End, 0))
8620              Case "CRPT_Cash_BEG"
8630                varRetVal = CCur(Forms(gstrFormQuerySpec).CashAssets_Beg)
8640              Case "CRPT_Cash_END"
8650                varRetVal = CCur(Forms(gstrFormQuerySpec).CashAssets_End)
8660              Case "Inc_Tot"
8670                varRetVal = CCur(gdblCrtRpt_IncTot)
8680              Case "Prin_Tot"
8690                varRetVal = CCur(gdblCrtRpt_PrinTot)
8700              Case "Cost_Tot"
8710                varRetVal = CCur(gdblCrtRpt_CostTot)
8720              Case "Period"
8730                Select Case Forms(gstrFormQuerySpec).chkAssetList_Start
                    Case True
8740                  varRetVal = "As of " & Format((CDate(Forms(gstrFormQuerySpec).DateStart) - 1), "mm/dd/yyyy")
8750                Case False
8760                  varRetVal = "From " & Format(CDate(Forms(gstrFormQuerySpec).DateStart), "mm/dd/yyyy") & _
                        " To " & Format(CDate(Forms(gstrFormQuerySpec).DateEnd), "mm/dd/yyyy")
8770                End Select
8780              Case "StartDate"
8790                varRetVal = CDate(Forms(gstrFormQuerySpec).DateStart)
8800              Case "PStartDate"
8810                varRetVal = CDate("01/01/1900")
8820              Case "EndDate"
8830                varRetVal = CDate(Forms(gstrFormQuerySpec).DateEnd)
8840              Case "PEndDate"
8850                varRetVal = DateAdd("d", -1#, Forms(gstrFormQuerySpec).DateStart)
8860              Case "AccountNo"
8870                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
8880              Case "ShortName"
8890                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts.Column(3)
8900              Case "Author"
8910                Select Case Forms(gstrFormQuerySpec).opgType
                    Case Forms(gstrFormQuerySpec).opgType_optRep.OptionValue
8920                  varRetVal = "Personal Representative"
8930                Case Forms(gstrFormQuerySpec).opgType_optGuard.OptionValue
8940                  varRetVal = "Guardian of Property"
8950                End Select
8960              Case "InvestInfo"  'NOT USED WITH FLORIDA.
8970                varRetVal = CCur(Nz(Forms(gstrFormQuerySpec).InvestInfo, 0))
8980              Case "InvestChange"  'NOT USED WITH FLORIDA.
8990                varRetVal = CCur(Nz(Forms(gstrFormQuerySpec).InvestChange, 0))
9000              Case "OrdVer"
9010                varRetVal = Forms(gstrFormQuerySpec).Ordinal & " And " & Forms(gstrFormQuerySpec).Version
9020              Case "Type"
9030                varRetVal = "Account of " & _
                      IIf(Forms(gstrFormQuerySpec).opgType = Forms(gstrFormQuerySpec).opgType_optGuard.OptionValue, _
                      "Guardian of Property", "Personal Representative")
9040              Case "TypeShort"
9050                varRetVal = IIf(Forms(gstrFormQuerySpec).opgType = Forms(gstrFormQuerySpec).opgType_optGuard.OptionValue, _
                      "Grdn", "Rep")
9060              Case "Schd_CapTrans"
9070                varRetVal = "Capital Transactions and Adjustments - " & Forms(gstrFormQuerySpec).cmdPreview04_Sch_lbl.Caption
9080              Case "Schd_Disburse"
9090                varRetVal = IIf(Forms(gstrFormQuerySpec).opgType = Forms(gstrFormQuerySpec).opgType_optGuard.OptionValue, _
                      "Disbursements and Distributions - ", "Disbursements - ") & _
                      IIf(Forms(gstrFormQuerySpec).chkGroupBy_IncExpCode = True, "Grouped - ", vbNullString) & "Schedule B"
9100              Case "Title7"
9110                Select Case Forms(gstrFormQuerySpec).chkAssetList_Start
                    Case True
9120                  varRetVal = "Assets on Hand at Beginning of Accounting Period"
9130                Case False
9140                  varRetVal = "Assets on Hand at Close of Accounting Period - " & Forms(gstrFormQuerySpec).cmdPreview05_Sch_lbl.Caption
9150                End Select
9160              End Select
9170            Else
9180              varRetVal = CCur(0)
9190            End If
9200          Else
9210            If IsMissing(varItem) = False Then
9220              Select Case varItem
                  Case "AccountNo"
9230                gstrAccountNo = "11"
9240                varRetVal = gstrAccountNo
9250              Case "StartDate"
9260                gdatStartDate = CDate("01/01/2009")
9270                varRetVal = gdatStartDate
9280              Case "EndDate"
9290                gdatEndDate = CDate("12/31/2009")
9300                varRetVal = gdatEndDate
9310              Case "TypeShort"
9320                varRetVal = "Grdn"
9330              End Select
9340            Else
9350              varRetVal = CCur(0)
9360            End If
9370          End If
9380        Else
9390          varRetVal = CCur(0)
9400        End If
9410      Case "frmRpt_CourtReports_NS"
9420        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
9430          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
9440            If IsMissing(varItem) = False Then
9450              Select Case varItem
                  Case "Period"
9460                varRetVal = "From " & Format(CDate(Forms(gstrFormQuerySpec).DateStart), "mm/dd/yyyy") & _
                      " To " & Format(CDate(Forms(gstrFormQuerySpec).DateEnd), "mm/dd/yyyy")
9470              Case "StartDate"
9480                varRetVal = CDate(Forms(gstrFormQuerySpec).DateStart)
9490              Case "StartDateLong"
9500                varRetVal = CLng(CDate(Forms(gstrFormQuerySpec).DateStart))
9510              Case "EndDate"
9520                varRetVal = CDate(Forms(gstrFormQuerySpec).DateEnd)
9530              Case "EndDateLong"
9540                varRetVal = CLng(CDate(Forms(gstrFormQuerySpec).DateEnd))
9550              Case "AccountNo"
9560                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
9570              Case "ShortName"
9580                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts.Column(3)
9590              Case "CaseNum"
9600                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts.Column(7)
9610              Case "OrdVer"
9620                varRetVal = Forms(gstrFormQuerySpec).Ordinal & " And " & Forms(gstrFormQuerySpec).Version
9630              End Select
9640            End If
9650          Else
9660            If IsMissing(varItem) = False Then
9670              Select Case varItem
                  Case "Period"
9680                varRetVal = "From 01/01/2009 To 12/31/2009"
9690              Case "StartDate"
9700                varRetVal = CDate("01/01/2009")
9710              Case "EndDate"
9720                varRetVal = CDate("12/31/2009")
9730              Case "AccountNo"
9740                varRetVal = "11"
9750              Case "ShortName"
9760                varRetVal = "William B. Johnson Trust"
9770              Case "CaseNum"
9780                varRetVal = "07-4-02640-9 SEA"
9790              End Select
9800            End If
9810          End If
9820        End If
9830      Case "frmRpt_CourtReports_NY"
9840        If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
9850          If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
9860            If IsMissing(varItem) = False Then
9870              Select Case varItem
                  Case "Period"
9880                varRetVal = "From " & Format(CDate(Forms(gstrFormQuerySpec).DateStart), "mm/dd/yyyy") & _
                      " To " & Format(CDate(Forms(gstrFormQuerySpec).DateEnd), "mm/dd/yyyy")
9890              Case "StartDate"
9900                varRetVal = gdatCrtRpt_NY_DateStart
9910              Case "EndDate"
9920                varRetVal = gdatCrtRpt_NY_DateEnd
9930              Case "AccountNo"
9940                varRetVal = gstrCrtRpt_NY_AccountNo
9950              Case "ShortName"
9960                varRetVal = Forms(gstrFormQuerySpec).cmbAccounts.Column(3)
9970              Case "OrdVer"
9980                varRetVal = Forms(gstrFormQuerySpec).Ordinal & " And " & Forms(gstrFormQuerySpec).Version & " Account"
9990              Case "NewInput"
10000               varRetVal = gcurCrtRpt_NY_InputNew  ' ** From input popup.
10010             Case "gcurCrtRpt_NY_IncomeBeg"
10020               gcurCrtRpt_NY_IncomeBeg = Nz(DLookup("icash", "qryCourtReport_NY_00_B_01"), 0)
10030               varRetVal = gcurCrtRpt_NY_IncomeBeg
10040             Case "IncomeCash"
10050               gcurCrtRpt_NY_ICash = Nz(DLookup("[icash]", "qryCourtReport_NY_00_B_01"), 0)
10060               varRetVal = gcurCrtRpt_NY_ICash
10070             Case "InvestedIncome"
10080               varRetVal = Nz(DLookup("[tcost]", "qryCourtReport_NY_InvestedIncome_b"), 0)
10090             Case "IncomeOnHand"
10100               varRetVal = Nz(DLookup("[tAmount]", "qryCourtReport_NY_InvestedIncome_h"), 0)
10110             Case "gstrCrtRpt_CashAssets_Beg"
10120               varRetVal = gstrCrtRpt_CashAssets_Beg  ' ** Same as gcurCrtRpt_NY_InputNew.
10130             End Select
10140           End If
10150         Else
10160           If IsMissing(varItem) = False Then
10170             Select Case varItem
                  Case "Period"
10180               varRetVal = "From 01/01/2015 To 12/31/2015"
10190             Case "StartDate"
10200               varRetVal = CDate("01/01/2015")
10210             Case "EndDate"
10220               varRetVal = CDate("12/31/2015")
10230             Case "AccountNo"
10240               varRetVal = "11"
10250             Case "ShortName"
10260               varRetVal = "William B. Johnson Trust"
10270             Case "OrdVer"
10280               varRetVal = "First and Final Account"
10290             Case "NewInput"
10300               varRetVal = 0
10310             Case "gcurCrtRpt_NY_IncomeBeg"
10320               varRetVal = 0
10330             Case "IncomeCash"
10340               varRetVal = 0
10350             Case "InvestedIncome"
10360               varRetVal = 0
10370             Case "IncomeOnHand"
10380               varRetVal = 0
10390             Case "gstrCrtRpt_CashAssets_Beg"
10400               varRetVal = 0
10410             End Select
10420           End If
10430         End If
10440       End If
10450     Case "frmRpt_Checks_Multi"
10460       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
10470         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
10480           If IsMissing(varItem) = False Then
10490             Select Case varItem
                  Case "CheckCnt"  ' ** Total number of checks for all users.
                    ' ** Journal, grouped, just PrintCheck = True, with cnt_chks,
                    ' ** by specified FormRef('AccountNo').
10500               varTmp00 = DLookup("[cnt_chks]", "qryRpt_Checks_Multi_03_03")
10510               Select Case IsNull(varTmp00)
                    Case True
10520                 varRetVal = 0&
10530               Case False
10540                 varRetVal = varTmp00
10550               End Select
10560             Case "AccountNo"
10570               varRetVal = Forms("frmRpt_Checks").lbxShortAccountName.Column(LBX_CHK_ACTNO)
10580             End Select
10590           End If
10600         Else
10610           If IsMissing(varItem) = False Then
10620             Select Case varItem
                  Case "CheckCnt"
                    ' ** Journal, grouped, just PrintCheck = True, with cnt_chks,
                    ' ** by specified FormRef('AccountNo').
10630               varTmp00 = DLookup("[cnt_chks]", "qryRpt_Checks_Multi_03_03")
10640               Select Case IsNull(varTmp00)
                    Case True
10650                 varRetVal = 0&
10660               Case False
10670                 varRetVal = varTmp00
10680               End Select
10690             Case "AccountNo"
10700               varRetVal = "215"
10710             End Select
10720           End If
10730         End If
10740       End If
10750     Case "frmRpt_Holdings"
10760       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
10770         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
10780           varRetVal = Forms(gstrFormQuerySpec).cmbAssets
10790         Else
10800           varRetVal = 2&
10810         End If
10820       End If
10830     Case "frmRpt_IncomeExpense"
10840       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
10850         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
10860           If IsMissing(varItem) = False Then
10870             Select Case varItem
                  Case "StartDate"
10880               varRetVal = CDate(Forms(gstrFormQuerySpec).DateStart)
10890             Case "EndDate"
10900               varRetVal = CDate(Forms(gstrFormQuerySpec).DateEnd)
10910             Case "AccountNo"
10920               varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
10930             Case "CriteriaMsg"
10940               varRetVal = gstrCrtRpt_Ordinal  ' ** Borrowing this variable from Court Reports.
10950             Case "AcctEveryLine"
10960               varRetVal = Forms(gstrFormQuerySpec).chkAcctEveryLine
10970             End Select
10980           Else
10990             varRetVal = CCur(0)
11000           End If
11010         Else
                ' ** For developement.
11020           If IsMissing(varItem) = False Then
11030             Select Case varItem
                  Case "StartDate"
11040               varRetVal = CDate("07/01/2010")
11050             Case "EndDate"
11060               varRetVal = CDate("06/30/2011")
11070             Case "AccountNo"
11080               varRetVal = "00119"
11090             Case "CriteriaMsg"
11100               varRetVal = gstrCrtRpt_Ordinal  ' ** Borrowing this variable from Court Reports.
11110             Case "AcctEveryLine"
11120               varRetVal = CBool(False)
11130             End Select
11140           End If
11150         End If
11160       Else
              ' ** For developement.
11170         If IsMissing(varItem) = False Then
11180           Select Case varItem
                Case "StartDate"
11190             varRetVal = CDate("07/01/2010")
11200           Case "EndDate"
11210             varRetVal = CDate("06/30/2011")
11220           Case "AccountNo"
11230             varRetVal = "00119"
11240           Case "CriteriaMsg"
11250             varRetVal = gstrCrtRpt_Ordinal  ' ** Borrowing this variable from Court Reports.
11260           Case "AcctEveryLine"
11270             varRetVal = CBool(False)
11280           End Select
11290         Else
11300           varRetVal = CCur(0)
11310         End If
11320       End If
11330     Case "frmRpt_IncomeStatement"
11340       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
11350         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
11360           Select Case varItem
                Case "StartDate"
11370             varRetVal = Forms("frmRpt_IncomeStatement").DateStart
11380           Case "EndDate"
11390             varRetVal = Forms("frmRpt_IncomeStatement").DateEnd
11400           Case "AccountNo"
11410             varRetVal = Forms("frmRpt_IncomeStatement").cmbAccounts
11420           Case "AccountName"
11430             varRetVal = gstrAccountName
11440           End Select
11450         End If
11460       Else
              ' ** For developement.
11470         If IsMissing(varItem) = False Then
11480           Select Case varItem
                Case "StartDate"
11490             varRetVal = CDate("01/01/2009")
11500           Case "EndDate"
11510             varRetVal = CDate("12/31/2009")
11520           Case "AccountNo"
11530             varRetVal = "11"
11540           Case "AccountName"
11550             varRetVal = "William B. Johnson Trust"
11560           End Select
11570         End If
11580       End If
11590     Case "frmRpt_Locations"
11600       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
11610         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
11620           varRetVal = Forms(gstrFormQuerySpec).cmbLocations
11630         End If
11640       End If
11650     Case "frmRpt_NewClosedAccounts"
11660       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
11670         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
11680           If IsMissing(varItem) = False Then
11690             Select Case varItem
                  Case "StartDate"
11700               varRetVal = Forms(gstrFormQuerySpec).DateStart
11710             Case "EndDate"
11720               varRetVal = Forms(gstrFormQuerySpec).DateEnd
11730             End Select
11740           End If
11750         End If
11760       End If
11770     Case "frmRpt_PurchasedSold"
11780       If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
11790         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
11800           If IsMissing(varItem) = False Then
11810             Select Case varItem
                  Case "DateStart"
11820               varRetVal = Format(CDate(Forms("frmRpt_PurchasedSold").DateStart), "mm/dd/yyyy")
11830             Case "DateEnd"
11840               varRetVal = Format(CDate(Forms("frmRpt_PurchasedSold").DateEnd), "mm/dd/yyyy")
11850             End Select
11860           End If
11870         End If
11880       End If
11890     Case "frmRpt_StatementOfCondition"
11900       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
11910         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
11920           Select Case varItem
                Case "Period"
11930             varRetVal = CDate(Forms(gstrFormQuerySpec).DateAsOf)
11940           Case "AcctType"
11950             Select Case Forms(gstrFormQuerySpec).opgAccountType
                  Case Forms(gstrFormQuerySpec).opgAccountType_optAll.OptionValue
11960               varRetVal = "All Accounts"
11970             Case Forms(gstrFormQuerySpec).opgAccountType_optDisc.OptionValue
11980               varRetVal = "Discretionary Accounts"
11990             Case Forms(gstrFormQuerySpec).opgAccountType_optNonDisc.OptionValue
12000               varRetVal = "Non-Discretionary Accounts"
12010             End Select
12020           End Select
12030         Else
12040           Select Case varItem
                Case "Period"
12050             varRetVal = CDate("07/15/2012")
12060           Case "AcctType"
12070             varRetVal = "All Accounts"
12080           End Select
12090         End If
12100       End If
12110     Case "frmRpt_TaxLot"
12120       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
12130         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
12140           Select Case varItem
                Case "AccountNo"
12150             varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
12160           Case Else
12170             varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
12180           End Select
12190         Else
12200           varRetVal = "11"
12210         End If
12220       End If
12230     Case "frmRpt_TransactionsByType"
12240       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
12250         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
12260           If IsMissing(varItem) = False Then
12270             Select Case varItem
                  Case "accountno"
12280               varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
12290             Case "StartDate"
12300               varRetVal = gdatStartDate
12310             Case "EndDate"
12320               varRetVal = gdatEndDate
12330             Case Else
12340               If IsNumeric(varItem) = True Then
12350                 For lngX = 1& To 12&
                        ' ** JType_01_chk - JType_12_chk.
12360                   If CStr(lngX) = varItem Then
12370                     If Forms(gstrFormQuerySpec).Controls("JType_" & Right(("00" & varItem), 2) & "_chk") = True Then
12380                       varRetVal = Forms(gstrFormQuerySpec).Controls("JType_" & Right(("00" & varItem), 2) & "_chk").Tag
12390                       If varRetVal = "Liability" Then
12400                         If InStr(Forms(gstrFormQuerySpec).Controls("JType_" & Right(("00" & varItem), 2) & "_chk").StatusBarText, "+") > 0 Then
12410                           varRetVal = varRetVal & " (+)"
12420                         Else
12430                           varRetVal = varRetVal & " (-)"
12440                         End If
12450                       End If
12460                     Else
12470                       varRetVal = vbNullString
12480                     End If
12490                     Exit For
12500                   End If
12510                 Next
12520               Else
12530                 varRetVal = vbNullString
12540               End If
12550             End Select
12560           End If
12570         End If
12580       End If
12590     Case "frmRpt_UnrealizedGainAndLoss"
12600       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
12610         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
12620           If IsMissing(varItem) = False Then
12630             Select Case varItem
                  Case "accountno"
12640               varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
12650             Case "DateAsOf"
12660               varRetVal = CDate(Forms(gstrFormQuerySpec).DateAsOf)
12670             Case "DateAsOfTime"
12680               varRetVal = CDate(Format(Forms(gstrFormQuerySpec).DateAsOf, "mm/dd/yyyy") & " 11:59:59 PM")
12690             End Select
12700           End If
12710         Else
12720           If IsMissing(varItem) = False Then
12730             Select Case varItem
                  Case "accountno"
12740               varRetVal = "11"
12750             Case "DateAsOf"
12760               varRetVal = CDate("04/26/2015")
12770             Case "DateAsOfTime"
12780               varRetVal = CDate("04/26/2015 11:59:59 PM")
12790             End Select
12800           End If
12810         End If
12820       End If
12830     Case "frmStatementBalance"
12840       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
12850         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
12860           If IsMissing(varItem) = False Then
12870             Select Case varItem
                  Case "LastBal"
12880               If IsNull(Forms(gstrFormQuerySpec).LastDate) = False Then
12890                 varRetVal = CDate(Forms(gstrFormQuerySpec).LastDate)
12900               Else
12910                 varRetVal = CDate("01/01/1900")
12920               End If
12930             Case "accountno"
12940               varRetVal = Forms(gstrFormQuerySpec).accountno
12950             End Select
12960           End If
12970         End If
12980       End If
12990     Case "frmStatementParameters"
13000       If IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
13010         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
13020           If IsMissing(varItem) = False Then
13030             Select Case varItem
                  Case "accountno"
13040               varRetVal = Forms(gstrFormQuerySpec).cmbAccounts
13050             Case "CurrentDate"
13060               varRetVal = CDate(Forms(gstrFormQuerySpec).currentDate)  '#2/15/2010#
13070             Case "EndDate"
13080               If IsNull(Forms(gstrFormQuerySpec).DateEnd) = False Then
13090                 If IsDate(Forms(gstrFormQuerySpec).DateEnd) = True Then
13100                   varRetVal = CDate(Forms(gstrFormQuerySpec).DateEnd)  '#2/28/2010#
13110                 End If
13120               Else
13130                 varRetVal = gdatEndDate
13140               End If
13150             Case "AsOf"
13160               varRetVal = CDate(Forms(gstrFormQuerySpec).AssetListDate)
13170             Case "Period"
13180               varRetVal = "As of " & Format(CDate(Forms(gstrFormQuerySpec).AssetListDate), "mmmm dd, yyyy")
13190             Case "StartDateTrans"
13200               If IsNull(Forms(gstrFormQuerySpec).TransDateStart) = False Then
13210                 varRetVal = CDate(Forms(gstrFormQuerySpec).TransDateStart)
13220               Else
13230                 If IsNull(Forms(gstrFormQuerySpec).DateStart) = False Then
13240                   varRetVal = CDate(Forms(gstrFormQuerySpec).DateStart)
13250                 Else
13260                   varRetVal = Date
13270                 End If
13280               End If
13290             Case "EndDateTrans"
13300               If IsNull(Forms(gstrFormQuerySpec).TransDateEnd) = False Then
13310                 varRetVal = CDate(Forms(gstrFormQuerySpec).TransDateEnd)
13320               Else
13330                 If IsNull(Forms(gstrFormQuerySpec).DateEnd) = False Then
13340                   varRetVal = CDate(Forms(gstrFormQuerySpec).DateEnd)
13350                 Else
13360                   varRetVal = Date
13370                 End If
13380               End If
13390             Case "MonthNum"
13400               varRetVal = CLng(Nz(Forms(gstrFormQuerySpec).cmbMonth.Column(CBX_MON_ID), 12))
13410             Case "Month"
13420               varRetVal = Forms(gstrFormQuerySpec).cmbMonth.Column(2)  ' ** 3-Ltr Month.
13430             Case "MaxBalDate"
                    ' ** 1st of statement month, plus 1 month, minus 1 day.
13440               varRetVal = CDate(DateAdd("y", -1, (DateAdd("m", 1, _
                      DateSerial(Forms(gstrFormQuerySpec).StatementsYear, Forms(gstrFormQuerySpec).cmbMonth.Column(CBX_MON_ID), 1)))))
13450             End Select
13460           End If
13470         Else
13480           If IsMissing(varItem) = False Then
13490             Select Case varItem
                  Case "accountno"
13500               varRetVal = "10-002"  'Forms("frmStatementParameters").cmbAccounts
13510             Case "CurrentDate"
13520               varRetVal = CDate("06/30/2016")  'CDate(Forms("frmStatementParameters").currentDate)  '#2/15/2010#
13530             Case "EndDate"
13540               varRetVal = CDate("06/30/2016")  'CDate(Forms("frmStatementParameters").DateEnd)  '#2/28/2010#
13550             Case "AsOf"
13560               varRetVal = CDate("06/30/2016")
13570             Case "Period"
13580               varRetVal = "As of " & Format(CDate("06/30/2016"), "mmmm dd, yyyy")  'CDate(Forms("frmStatementParameters").AssetListDate)
13590             Case "StartDateTrans"
13600               varRetVal = CDate("01/01/2016")
13610             Case "EndDateTrans"
13620               varRetVal = CDate("06/30/2016")
13630             Case "MonthNum"
13640               varRetVal = CInt(12)
13650             Case "Month"
13660               varRetVal = "dec"
13670             Case "MaxBalDate"
13680               varRetVal = DateAdd("y", -1, (DateAdd("m", 1, DateSerial(2006, 5, 1))))
13690             End Select
13700           End If
13710         End If
13720       End If
13730     Case "frmTaxLot"
            ' ** [actno] can use "frmJournal_Sub4_Sold"
            ' ** [astno] can use "frmTaxLot"
13740       If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
13750         If Forms(gstrFormQuerySpec).CurrentView <> acCurViewDesign Then
13760           Select Case varItem
                Case "accountno"
13770             varRetVal = Forms(gstrFormQuerySpec).accountno
13780           Case "assetno"
13790             varRetVal = CLng(Forms(gstrFormQuerySpec).frmJournal_Sub4_Sold.Form.saleAssetno.Column(2))
13800           Case "AssetDateNull"
13810             varRetVal = Forms(gstrFormQuerySpec).AssetDateNull
13820           Case "AssetDateSale"
13830             varRetVal = Forms(gstrFormQuerySpec).AssetDateSale
13840           End Select
13850         End If
13860       End If
13870     Case "frmXAdmin"
13880       If IsLoaded(gstrFormQuerySpec, acForm, False) = True Then  ' ** Module Function: modFileUtilities.
13890         If IsMissing(varItem) = False Then
13900           Select Case varItem
                Case "xadcust_id"
13910             varRetVal = CLng(Nz(Forms(gstrFormQuerySpec).xadcust_id, 0))
13920           End Select
13930         End If
13940       ElseIf IsLoaded(gstrFormQuerySpec, acForm, True) = True Then  ' ** Module Function: modFileUtilities.
13950         varRetVal = CLng(1)
13960       Else
13970         varRetVal = CLng(1)
13980       End If
13990     Case "frmXAdmin_Graphics"
14000       If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
14010         If IsMissing(varItem) = False Then
14020           Select Case varItem
                Case "Type"
14030             Select Case IsNull(Forms(gstrFormQuerySpec).lbxGraphicsType)
                  Case True
14040               varRetVal = CLng(0)
14050             Case False
14060               varRetVal = CLng(Forms(gstrFormQuerySpec).lbxGraphicsType)
14070             End Select
14080           Case "Group"
14090             Select Case IsNull(Forms(gstrFormQuerySpec).lbxGraphicsGroup)
                  Case True
14100               varRetVal = "All"
14110             Case False
14120               varRetVal = Forms(gstrFormQuerySpec).lbxGraphicsGroup
14130             End Select
14140           End Select
14150         End If
14160       End If
14170     Case "frmXAdmin_Version"
14180       If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
14190         If IsMissing(varItem) = False Then
14200           Select Case varItem
                Case "Most"
14210             varRetVal = Forms(gstrFormQuerySpec).UpgradeCode
14220           Case Else
14230             If lngOddUps = 0& Then
14240               Set dbs = CurrentDb
14250               With dbs
14260                 Set qdf = .QueryDefs("qryVersion_24")
14270                 Set rst = qdf.OpenRecordset
14280                 With rst
14290                   .MoveLast
14300                   lngOddUps = .RecordCount
14310                   .MoveFirst
14320                   arr_varOddUp = .GetRows(lngOddUps)
14330                   .Close
14340                 End With
14350                 .Close
14360               End With
14370             End If
14380             varRetVal = False
14390             For lngX = 0& To (lngOddUps - 1&)
14400               If arr_varOddUp(0, lngX) = varItem Then
14410                 varRetVal = True
14420                 Exit For
14430               End If
14440             Next
14450           End Select
14460         Else
14470           varRetVal = "{0DD4D247-4513-48CC-AA44-29C9F80B9EB2}"
14480         End If
14490       Else
14500         If IsMissing(varItem) = False Then
14510           Select Case varItem
                Case "Most"
14520             varRetVal = "{0DD4D247-4513-48CC-AA44-29C9F80B9EB2}"
14530           Case Else
14540             If lngOddUps = 0& Then
14550               Set dbs = CurrentDb
14560               With dbs
14570                 Set qdf = .QueryDefs("qryVersion_24")
14580                 Set rst = qdf.OpenRecordset
14590                 With rst
14600                   .MoveLast
14610                   lngOddUps = .RecordCount
14620                   .MoveFirst
14630                   arr_varOddUp = .GetRows(lngOddUps)
14640                   .Close
14650                 End With
14660                 .Close
14670               End With
14680             End If
14690             varRetVal = False
14700             For lngX = 0& To (lngOddUps - 1&)
14710               If arr_varOddUp(0, lngX) = varItem Then
14720                 varRetVal = True
14730                 Exit For
14740               End If
14750             Next
14760           End Select
14770         Else
14780           varRetVal = "{0DD4D247-4513-48CC-AA44-29C9F80B9EB2}"
14790         End If
14800       End If
14810     Case "frmXAdmin_Form_Graphics"
14820       If IsLoaded(gstrFormQuerySpec, acForm) = True Then  ' ** Module Function: modFileUtilities.
14830         If IsMissing(varItem) = False Then
14840           Select Case varItem
                Case "frm_id"
14850             varRetVal = Forms(gstrFormQuerySpec).frm_id
14860           Case "frm_id_new"
14870             varRetVal = CLng(Forms(gstrFormQuerySpec & "_Add").cmbForms)   '325
14880           End Select
14890         End If
14900       End If
14910     Case Else  ' ** Frm_Rebuild_01() in zz_mod_FormRebuildFuncs.
14920       intPos01 = InStr(gstrFormQuerySpec, "~")
14930       If intPos01 > 0 Then
14940         strTmp01 = Left(gstrFormQuerySpec, (intPos01 - 1))
14950         varTmp00 = Mid(gstrFormQuerySpec, (intPos01 + 1))
14960         Select Case strTmp01
              Case "Frm_Rebuild_01"
14970           Select Case varTmp00
                Case "frmRpt_Checks", "frmAccountContacts", "frmAccountContacts_Sub", "frmOptions", "frmMenu_Utility"
14980             varRetVal = glngUserCntLedger  ' ** Borrowing this from the Post Menu.
14990           Case Else
                  ' ** None at the moment.
15000           End Select
15010         Case Else
                ' ** Nothing else at the moment.
15020         End Select
15030       Else
15040         varRetVal = glngUserCntLedger  ' ** Borrowing this from the Post Menu.
15050       End If
15060     End Select
15070   Else
15080     If varItem = "accountno" Then
15090       varRetVal = "147" '"11"
15100     ElseIf varItem = "datbeg" Then
15110       varRetVal = "1/1/1900"
15120     ElseIf varItem = "datend" Then
15130       varRetVal = Format(Date, "mm/dd/yyyy")
15140     End If
15150   End If

EXITP:
15160   Set rst = Nothing
15170   Set qdf = Nothing
15180   Set dbs = Nothing
15190   FormRef = varRetVal
15200   Exit Function

ERRH:
15210   varRetVal = Null
15220   Select Case ERR.Number
        Case Else
15230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15240   End Select
15250   Resume EXITP

End Function

Public Function Qry_CheckBox(Optional varQry As Variant, Optional varFld As Variant, Optional varChkBox As Variant, Optional varType As Variant) As Boolean
' ** Change DisplayControl of Boolean field in a query to acCheckBox type.

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_CheckBox"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, fld As DAO.Field, prp As DAO.Property
        Dim strQdfName As String, strFldName As String
        Dim blnCheckBox As Boolean, intType As Integer, intCtlType As Integer, intMode As Integer
        Dim blnFound As Boolean
        Dim strTmp01 As String, blnTmp02 As Boolean, intTmp03 As Integer
        Dim lngW As Long, lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

        Const DT_S As Integer = 1
        Const DT_SHORT As String = "mm/dd/yyyy"
        Const DT_L As Integer = 2
        Const DT_LONG As String = "mm/dd/yyyy hh:nn:ss"

15310   blnRetVal = True

        'For lngZ = 1& To 5&

15320   For lngW = 1& To 1&

15330     intTmp03 = 0&: intMode = 0
15340     Select Case lngW
          Case 1&
15350       strTmp01 = "objtype_type" '"IsBlank_new" '"curr_active"  '"ledger_HIDDEN"  '"revcode_ACTIVE"
15360       blnTmp02 = True
15370       intTmp03 = acComboBox
15380       intMode = 8
15390     'Case 2&
15400     '  strTmp01 = "assetdate" '"IsHdr_new" '"IsData01_new" '"CheckPaid" '"ledghidtype_type_new"
15410     '  blnTmp02 = False
15420     '  intTmp03 = acCheckBox
15430     '  intMode = 0
15440     'Case 3&
15450     '  strTmp01 = "PurchaseDate" ' "IsFtr_new" '"IsSubHdr_new" '"assetdate"  '"IsAverage"
15460     '  blnTmp02 = False
15470     '  intTmp03 = acCheckBox
15480     '  intMode = 0
15490       'Case 4&
15500       '  strTmp01 = "has_vline1"  '"IsSubTot_new" '"posted" '"HasPay_new", '"HasPayHdr_new"
15510       '  blnTmp02 = True
15520       '  intTmp03 = acCheckBox
15530       '  intMode = 0
15540       'Case 5&
15550       '  strTmp01 = "has_vline2" '"IsColHdr_new" '"IsCatHdr_new" '"PurchaseDate" '"IsSubHdr_new"
15560       '  blnTmp02 = True
15570       '  intTmp03 = acCheckBox
15580       '  intMode = 0
15590       'Case 6&
15600       '  strTmp01 = "has_vline3" '"IsTot_new" '"HasShrsHdr_new" '"HasSym01_new" '"HasDate01_new"
15610       '  blnTmp02 = True
15620       '  intTmp03 = acCheckBox
15630       '  intMode = 0
15640       'Case 7&
15650       '  strTmp01 = "has_vline4" '"HasVal01_new" '"HasShrs_new" '"HasDsc01_new" '"HasDsc02_new"
15660       '  blnTmp02 = True
15670       '  intTmp03 = acCheckBox
15680       '  intMode = 0
15690       'Case 8&
15700       '  strTmp01 = "has_box2"  '"HasData_new" '"IsAst_new"
15710       '  blnTmp02 = True
15720       '  intTmp03 = acCheckBox
15730       '  intMode = 0
15740       'Case 9&
15750       '  strTmp01 = "fsp_alt_shift"  '"IsCatHdr_new"
15760       '  blnTmp02 = True
15770       '  intTmp03 = acCheckBox
15780       '  intMode = 0
15790       'Case 10&
15800       '  strTmp01 = "fsp_ctrl_alt_shift"
15810       '  blnTmp02 = True
15820       '  intTmp03 = acCheckBox
15830       '  intMode = 0
15840       'Case 11&
15850       '  strTmp01 = "vbcsc08_datemodified"
15860       '  blnTmp02 = False
15870       '  intTmp03 = acCheckBox
15880       '  intMode = 0
15890       'Case 12&
15900       '  strTmp01 = "HasSym01_new"
15910       '  blnTmp02 = True
15920       'Case 13&
15930       '  strTmp01 = "HasVal01_new"
15940       '  blnTmp02 = True
15950       'Case 14&
15960       '  strTmp01 = "HasVal04_new"
15970       '  blnTmp02 = True
15980     End Select

15990     For lngY = 1& To 3&

16000       Select Case IsMissing(varQry)
            Case True
              'Select Case lngZ
              'Case 1&
16010         strQdfName = "zzz_qry_zVBComponent_Query_11_03"
              'Case 2&
              '  strQdfName = "zzz_qry_zOhana_29_150_50_02"
              'Case 3&
              '  strQdfName = "zzz_qry_zOhana_29_150_50_03"
              'Case 4&
              '  strQdfName = "zzz_qry_zOhana_29_150_50_04"
              'Case 5&
              '  strQdfName = "zzz_qry_zOhana_29_150_50_05"
              'End Select
              'strFldName = "dp_closed_new"  '"ledger_HIDDEN", "CheckPaid", "transdate", "assetdate", "PurchaseDate", "posted"
16020         strFldName = strTmp01
              'blnCheckBox = True  ' ** True, Add CheckBox; False, Add Format.
16030         blnCheckBox = blnTmp02
16040         intType = DT_L  ' ** DT_S, DT_L
16050       Case False
16060         strQdfName = varQry
16070         strFldName = varFld
16080         Select Case IsMissing(varChkBox)
              Case True
16090           blnCheckBox = True
16100         Case False
16110           blnCheckBox = CBool(varChkBox)
16120         End Select
16130         Select Case IsMissing(varType)
              Case True
16140           intType = DT_L
16150         Case False
16160           intType = varType
16170         End Select
16180       End Select

16190       intCtlType = intTmp03  'acComboBox  'acCheckBox  'acComboBox

16200       Set dbs = CurrentDb
16210       With dbs
16220         Set qdf = .QueryDefs(strQdfName)
16230         With qdf
                'For Each fld In .Fields
16240           Set fld = .Fields(strFldName)
16250           With fld
16260             If .Name = strFldName Then

16270               Select Case blnCheckBox
                    Case True

                      ' ** Check if the DisplayControl property is already present.
16280                 blnFound = False
16290 On Error Resume Next
16300                 Set prp = .Properties("DisplayControl")
16310                 If ERR.Number = 0 Then
16320 On Error GoTo ERRH
16330                   blnFound = True
                        ' ** Change the DisplayControl property to acCheckBox.
16340                   If prp.Value <> intCtlType Then
16350                     prp.Value = intCtlType
16360                   End If
16370                 Else
16380 On Error GoTo ERRH
16390                 End If

16400                 If blnFound = True And intCtlType = acComboBox Then
16410                   For lngX = 1& To 9&
16420                     blnFound = False
16430                     For Each prp In .Properties
16440                       Select Case lngX
                            Case 1&
16450                         If prp.Name = "RowSourceType" Then
16460                           blnFound = True
16470                           If prp.Value <> "Table/Query" Then
16480                             prp.Value = "Table/Query"
16490                           End If
16500                           Exit For
16510                         End If
16520                       Case 2&
16530                         If prp.Name = "RowSource" Then
16540                           blnFound = True
16550                           Select Case intMode
                                Case 1
16560                             If prp.Value <> "qryCompilerDirective_01" Then
16570                               prp.Value = "qryCompilerDirective_01" '"qryDataTypeVb_01"
16580                             End If
16590                           Case 2
16600                             If prp.Value <> "qryCompilerDirectiveOption_01" Then
16610                               prp.Value = "qryCompilerDirectiveOption_01"
16620                             End If
16630                           Case 3
16640                             If prp.Value <> "qrySectionType_01" Then
16650                               prp.Value = "qrySectionType_01"
16660                             End If
16670                           Case 4
16680                             If prp.Value <> "qryControlType_01" Then
16690                               prp.Value = "qryControlType_01"
16700                             End If
16710                           Case 5
16720                             If prp.Value <> "qryKeyDownType_01" Then
16730                               prp.Value = "qryKeyDownType_01"
16740                             End If
16750                           Case 6
16760                             If prp.Value <> "qryLedgerHiddenType_01" Then
16770                               prp.Value = "qryLedgerHiddenType_01"
16780                             End If
16790                           Case 7
16800                             If prp.Value <> "qryDataTypeDb_01" Then
16810                               prp.Value = "qryDataTypeDb_01"
16820                             End If
100                             Case 8
110                               If prp.Value <> "qryObjectType_01" Then
120                                 prp.Value = "qryObjectType_01"
130                               End If
16830                           End Select
16840                           Exit For
16850                         End If
16860                       Case 3&
16870                         If prp.Name = "BoundColumn" Then
16880                           blnFound = True
16890                           If prp.Value <> 1 Then
16900                             prp.Value = 1
16910                           End If
16920                           Exit For
16930                         End If
16940                       Case 4&
16950                         If prp.Name = "ColumnCount" Then
16960                           blnFound = True
16970                           Select Case intMode
                                Case 1, 3, 4, 6, 7, 8
16980                             If prp.Value <> 3 Then
16990                               prp.Value = 3
17000                             End If
17010                           Case 2, 5
17020                             If prp.Value <> 4 Then
17030                               prp.Value = 4
17040                             End If
17050                           End Select
17060                           Exit For
17070                         End If
17080                       Case 5&
17090                         If prp.Name = "ColumnHeads" Then
17100                           blnFound = True
17110                           If prp.Value <> False Then
17120                             prp.Value = False
17130                           End If
17140                           Exit For
17150                         End If
17160                       Case 6&
17170                         If prp.Name = "ColumnWidths" Then
17180                           blnFound = True
17190                           Select Case intMode
                                Case 1, 3, 4, 6, 7, 8
17200                             If prp.Value <> "0;1440;0" Then
17210                               prp.Value = "0;1440;0"
17220                             End If
17230                           Case 2
17240                             If prp.Value <> "0;1830;1185;0" Then
17250                               prp.Value = "0;1830;1185;0"
17260                             End If
17270                           Case 5
17280                             If prp.Value <> "0;1440;0;0" Then
17290                               prp.Value = "0;1440;0;0"
17300                             End If
17310                           End Select
17320                           Exit For
17330                         End If
17340                       Case 7&
17350                         If prp.Name = "ListRows" Then
17360                           blnFound = True
17370                           If prp.Value <> 8 Then
17380                             prp.Value = 10
17390                           End If
17400                           Exit For
17410                         End If
17420                       Case 8&
17430                         If prp.Name = "ListWidth" Then
17440                           blnFound = True
17450                           Select Case intMode
                                Case 1, 3, 4, 5, 6, 7, 8
17460                             If prp.Value <> "Auto" Then
17470                               prp.Value = "Auto"
17480                             End If
17490                           Case 2
17500                             If prp.Value <> "3270twip" Then
17510                               prp.Value = "3270twip"
17520                             End If
17530                           End Select
17540                           Exit For
17550                         End If
17560                       Case 9&
17570                         If prp.Name = "LimitToList" Then
17580                           blnFound = True
17590                           If prp.Value <> True Then
17600                             prp.Value = True
17610                           End If
17620                           Exit For
17630                         End If
17640                       End Select
17650                     Next  ' ** prp
17660                     If blnFound = False Then
17670                       Select Case lngX
                            Case 1&
17680                         Set prp = .CreateProperty("RowSourceType", dbText, "Table/Query")
17690                         .Properties.Append prp
17700                         .Properties.Refresh
17710                       Case 2&
17720                         Select Case intMode
                              Case 1
17730                           Set prp = .CreateProperty("RowSource", dbText, "qryCompilerDirective_01")
17740                         Case 2
17750                           Set prp = .CreateProperty("RowSource", dbText, "qryCompilerDirectiveOption_01")
17760                         Case 3
17770                           Set prp = .CreateProperty("RowSource", dbText, "qrySectionType_01")
17780                         Case 4
17790                           Set prp = .CreateProperty("RowSource", dbText, "qryControlType_01")
17800                         Case 5
17810                           Set prp = .CreateProperty("RowSource", dbText, "qryKeyDownType_01")
17820                         Case 6
17830                           Set prp = .CreateProperty("RowSource", dbText, "qryLedgerHiddenType_01")
17840                         Case 7
17850                           Set prp = .CreateProperty("RowSource", dbText, "qryDataTypeDb_01")
140                           Case 8
150                             Set prp = .CreateProperty("RowSource", dbText, "qryObjectType_01")
17860                         End Select
17870                         .Properties.Append prp
17880                         .Properties.Refresh
17890                       Case 3&
17900                         Set prp = .CreateProperty("BoundColumn", dbInteger, 1)
17910                         .Properties.Append prp
17920                         .Properties.Refresh
17930                       Case 4&
17940                         Select Case intMode
                              Case 1, 3, 4, 6, 7, 8
17950                           Set prp = .CreateProperty("ColumnCount", dbInteger, 3)
17960                         Case 2, 5
17970                           Set prp = .CreateProperty("ColumnCount", dbInteger, 4)
17980                         End Select
17990                         .Properties.Append prp
18000                         .Properties.Refresh
18010                       Case 5&
18020                         Set prp = .CreateProperty("ColumnHeads", dbBoolean, False)
18030                         .Properties.Append prp
18040                         .Properties.Refresh
18050                       Case 6&
18060                         Select Case intMode
                              Case 1, 3, 4, 6, 7, 8
18070                           Set prp = .CreateProperty("ColumnWidths", dbText, "0;1440;0")
18080                         Case 2
18090                           Set prp = .CreateProperty("ColumnWidths", dbText, "0;1830;1185;0")
18100                         Case 5
18110                           Set prp = .CreateProperty("ColumnWidths", dbText, "0;1440;0;0")
18120                         End Select
18130                         .Properties.Append prp
18140                         .Properties.Refresh
18150                       Case 7&
18160                         Set prp = .CreateProperty("ListRows", dbInteger, 10)
18170                         .Properties.Append prp
18180                         .Properties.Refresh
18190                       Case 8&
18200                         Select Case intMode
                              Case 1, 3, 4, 5, 6, 7, 8
18210                           Set prp = .CreateProperty("ListWidth", dbText, "Auto")
18220                         Case 2
18230                           Set prp = .CreateProperty("ListWidth", dbText, "3270twip")
18240                         End Select
18250                         .Properties.Append prp
18260                         .Properties.Refresh
18270                       Case 9&
18280                         Set prp = .CreateProperty("LimitToList", dbBoolean, True)
18290                         .Properties.Append prp
18300                         .Properties.Refresh
18310                       End Select
18320                     End If
18330                   Next  ' ** lngX.
18340                 End If

                      ' ** Add the DisplayControl property, and set it to acCheckBox.
18350                 If blnCheckBox = True And blnFound = False Then
18360                   Set prp = .CreateProperty("DisplayControl", dbInteger, intCtlType)
18370 On Error Resume Next
18380                   .Properties.Append prp
18390                   If ERR.Number <> 0 Then
18400                     If .Properties("DisplayControl") <> intCtlType Then
18410                       .Properties("DisplayControl") = intCtlType
18420                     End If
18430 On Error GoTo ERRH
18440                   Else
18450 On Error GoTo ERRH
18460                   End If
18470                   .Properties.Refresh
18480                   If intCtlType = acComboBox Then
18490 On Error Resume Next
18500                     Set prp = .CreateProperty("RowSourceType", dbText, "Table/Query")
18510                     .Properties.Append prp
18520                     .Properties.Refresh
18530                     Select Case intMode
                          Case 1
18540                       Set prp = .CreateProperty("RowSource", dbText, "qryCompilerDirective_01")
18550                     Case 2
18560                       Set prp = .CreateProperty("RowSource", dbText, "qryCompilerDirectiveOption_01")
18570                     Case 3
18580                       Set prp = .CreateProperty("RowSource", dbText, "qrySectionType_01")
18590                     Case 4
18600                       Set prp = .CreateProperty("RowSource", dbText, "qryControlType_01")
18610                     Case 5
18620                       Set prp = .CreateProperty("RowSource", dbText, "qryKeyDownType_01")
18630                     Case 6
18640                       Set prp = .CreateProperty("RowSource", dbText, "qryLedgerHiddenType_01")
18650                     Case 7
18660                       Set prp = .CreateProperty("RowSource", dbText, "qryDataTypeDb_01")
160                       Case 8
170                         Set prp = .CreateProperty("RowSource", dbText, "qryObjectType_01")
18670                     End Select
18680                     .Properties.Append prp
18690                     .Properties.Refresh
18700                     Set prp = .CreateProperty("BoundColumn", dbInteger, 1)
18710                     .Properties.Append prp
18720                     .Properties.Refresh
18730                     Select Case intMode
                          Case 1, 3, 4, 6, 7, 8
18740                       Set prp = .CreateProperty("ColumnCount", dbInteger, 3)
18750                     Case 2, 5
18760                       Set prp = .CreateProperty("ColumnCount", dbInteger, 4)
18770                     End Select
18780                     .Properties.Append prp
18790                     .Properties.Refresh
18800                     Set prp = .CreateProperty("ColumnHeads", dbBoolean, False)
18810                     .Properties.Append prp
18820                     .Properties.Refresh
18830                     Select Case intMode
                          Case 1, 3, 4, 6, 7, 8
18840                       Set prp = .CreateProperty("ColumnWidths", dbText, "0;1440;0")
18850                     Case 2
18860                       Set prp = .CreateProperty("ColumnWidths", dbText, "0;1830;1185;0")
18870                     Case 5
18880                       Set prp = .CreateProperty("ColumnWidths", dbText, "0;1440;0;0")
18890                     End Select
18900                     .Properties.Append prp
18910                     .Properties.Refresh
18920                     Set prp = .CreateProperty("ListRows", dbInteger, 8)
18930                     .Properties.Append prp
18940                     .Properties.Refresh
18950                     Select Case intMode
                          Case 1, 3, 4, 5, 6, 7, 8
18960                       Set prp = .CreateProperty("ListWidth", dbText, "Auto")
18970                     Case 2
18980                       Set prp = .CreateProperty("ListWidth", dbText, "3270twip")
18990                     End Select
19000                     .Properties.Append prp
19010                     .Properties.Refresh
19020                     Set prp = .CreateProperty("LimitToList", dbBoolean, True)
19030                     .Properties.Append prp
19040                     .Properties.Refresh
19050 On Error GoTo ERRH
19060                   End If
19070                 End If

19080               Case False

                      ' ** See if the Format property already exists.
19090                 blnFound = False
                      'For Each prp In .Properties
19100 On Error Resume Next
19110                 Set prp = .Properties("Format")
19120                 If ERR.Number = 0 Then
19130 On Error GoTo ERRH
19140                   blnFound = True
19150                 Else
19160 On Error GoTo ERRH
19170                 End If
19180                 If blnFound = False Then
                        ' ** Add the Format property.
19190                   Set prp = .CreateProperty("Format", dbText, Choose(intType, DT_SHORT, DT_LONG))
19200                   .Properties.Append prp
19210                 Else
                        ' ** Update the Format property.
19220                   .Properties("Format") = Choose(intType, DT_SHORT, DT_LONG)
19230                 End If

19240               End Select
19250             End If
19260           End With
19270         End With

19280         .Close
19290       End With

19300       Beep

19310     Next  ' ** lngY.
19320   Next  ' ** lngW.

        'Next  ' ** lngZ

EXITP:
19330   Set prp = Nothing
19340   Set fld = Nothing
19350   Set qdf = Nothing
19360   Set dbs = Nothing
19370   Qry_CheckBox = blnRetVal
19380   Exit Function

ERRH:
19390   Select Case ERR.Number
        Case Else
19400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19410   End Select
19420   Resume EXITP

End Function

Public Function Qry_Copy() As Boolean

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_Copy"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object, rst As DAO.Recordset
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim strQryBase1 As String, strQryBase2 As String, strQryNum As String
        Dim lngRecs As Long, lngQrysCreated As Long
        Dim blnSkip As Boolean
        Dim intPos01 As Integer, intLen As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 6  ' ** Array's first-element UBound().
        Const Q_QNAM1 As Integer = 0
        Const Q_DSC1  As Integer = 1
        Const Q_SQL1  As Integer = 2
        Const Q_QNAM2 As Integer = 3
        Const Q_DSC2  As Integer = 4
        Const Q_SQL2  As Integer = 5
        Const Q_TYPE  As Integer = 6

19510 On Error GoTo 0

19520   blnRetVal = True

19530   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
19540   DoEvents

19550   strQryBase1 = "zzz_qry_yFrontier_"
19560   strQryBase2 = "zzz_qry_zFiduciary_"

19570   lngQrys = 0&
19580   ReDim arr_varQry(Q_ELEMS, 0)

19590   Set dbs = CurrentDb
19600   With dbs

19610     intLen = Len(strQryBase1)

19620     For Each qdf In .QueryDefs
19630       With qdf
19640         If Left(.Name, intLen) = strQryBase1 Then
19650           strQryNum = Mid(.Name, (intLen + 1), 2)
19660           If Val(strQryNum) < 20 Then
19670             lngQrys = lngQrys + 1&
19680             lngE = lngQrys - 1&
19690             ReDim Preserve arr_varQry(Q_ELEMS, lngE)
19700             arr_varQry(Q_QNAM1, lngE) = .Name
19710             arr_varQry(Q_DSC1, lngE) = .Properties("Description")
19720             arr_varQry(Q_SQL1, lngE) = .SQL
19730             arr_varQry(Q_QNAM2, lngE) = Null
19740             arr_varQry(Q_DSC2, lngE) = Null
19750             arr_varQry(Q_SQL2, lngE) = Null
19760             arr_varQry(Q_TYPE, lngE) = .Type
19770           End If
19780         End If
19790       End With  ' ** qdf.
19800     Next  ' ** qdf.
19810     Set qdf = Nothing

19820     Debug.Print "'QRYS: " & CStr(lngQrys)
19830     DoEvents

19840     If lngQrys > 0& Then

19850       For lngX = 0& To (lngQrys - 1&)
19860         strTmp01 = arr_varQry(Q_QNAM1, lngX)
19870         strTmp01 = strQryBase2 & Mid(strTmp01, (intLen + 1))
19880         arr_varQry(Q_QNAM2, lngX) = strTmp01
19890         strTmp01 = arr_varQry(Q_SQL1, lngX)
19900         strTmp01 = StringReplace(strTmp01, strQryBase1, strQryBase2)  ' ** Module Function: modStringFuncs.
19910         arr_varQry(Q_SQL2, lngX) = strTmp01
19920         strTmp01 = arr_varQry(Q_DSC1, lngX)
19930         strTmp02 = Right(strQryBase1, 6)
19940         strTmp03 = Right(strQryBase2, 6)
19950         strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.
19960         intPos01 = InStr(strTmp01, ";")
19970         If intPos01 > 0 Then
19980           strTmp01 = Left(strTmp01, (intPos01 + 1))
19990         End If
20000         arr_varQry(Q_DSC2, lngX) = strTmp01
20010       Next  ' ** lngX

20020       strTmp02 = Left(Right(strQryBase1, 5), 4)
20030       strTmp03 = Left(Right(strQryBase2, 5), 4)

20040       For lngX = 0& To (lngQrys - 1&)
20050         strTmp01 = arr_varQry(Q_SQL2, lngX)
              ' ** <=2004
20060         strTmp01 = StringReplace(strTmp01, "<=" & strTmp02, "<=" & strTmp03)  ' ** Module Function: modStringFuncs.
              ' ** <=#12/31/2004#
20070         strTmp01 = StringReplace(strTmp01, "<=#12/31/" & strTmp02 & "#", "<=#12/31/" & strTmp03 & "#")  ' ** Module Function: modStringFuncs.
              ' ** 2004 AS transyear
20080         strTmp01 = StringReplace(strTmp01, strTmp02 & " AS transyear", strTmp03 & " AS transyear")  ' ** Module Function: modStringFuncs.
              ' ** =#12/31/2004#
20090         strTmp01 = StringReplace(strTmp01, "=#12/31/" & strTmp02 & "#", "=#12/31/" & strTmp03 & "#")  ' ** Module Function: modStringFuncs.
              ' ** >=#01/01/2004# And <=#12/31/2004#
20100         strTmp01 = StringReplace(strTmp01, ">=#01/01/" & strTmp02 & "# And <=#12/31/" & strTmp02 & "#", _
                ">=#01/01/" & strTmp03 & "# And <=#12/31/" & strTmp03 & "#")  ' ** Module Function: modStringFuncs.
20110         strTmp01 = StringReplace(strTmp01, ">=#1/1/" & strTmp02 & "# And", ">=#1/1/" & strTmp03 & "# And")  ' ** Module Function: modStringFuncs.
20120         arr_varQry(Q_SQL2, lngX) = strTmp01
20130         strTmp01 = arr_varQry(Q_DSC2, lngX)
              ' ** <= 2004
20140         strTmp01 = StringReplace(strTmp01, "<= " & strTmp02, "<= " & strTmp03)  ' ** Module Function: modStringFuncs.
              ' ** <= 12/31/2004
20150         strTmp01 = StringReplace(strTmp01, "<= 12/31/" & strTmp02, "<= 12/31/" & strTmp03)  ' ** Module Function: modStringFuncs.
              ' ** just 12/31/2008
20160         strTmp01 = StringReplace(strTmp01, "just 12/31/" & strTmp02, "just 12/31/" & strTmp03)  ' ** Module Function: modStringFuncs.
20170         arr_varQry(Q_DSC2, lngX) = strTmp01
20180       Next  ' ** lngX.

20190       blnSkip = False
20200       If blnSkip = False Then

20210         lngQrysCreated = 0&
20220         For lngX = 0& To (lngQrys - 1&)
20230           Set qdf = .CreateQueryDef(arr_varQry(Q_QNAM2, lngX), arr_varQry(Q_SQL2, lngX))
20240           With qdf
20250             Set prp = .CreateProperty("Description", dbText, arr_varQry(Q_DSC2, lngX))
20260 On Error Resume Next
20270             qdf.Properties.Append prp
20280             If ERR.Number <> 0 Then
20290 On Error GoTo 0
20300               .Properties("Description") = arr_varQry(Q_DSC2, lngX)
20310             Else
20320 On Error GoTo 0
20330             End If
20340           End With
20350           DoEvents
20360           lngQrysCreated = lngQrysCreated + 1&
20370         Next  ' ** lngX.
20380         Set prp = Nothing
20390         Set qdf = Nothing

20400         .QueryDefs.Refresh

20410         blnSkip = True
20420         If blnSkip = False Then

20430           For lngX = 0& To (lngQrys - 1&)
20440             Set qdf = .QueryDefs(arr_varQry(Q_QNAM2, lngX))
20450             With qdf
20460               If .Type = dbQSelect Or .Type = dbQSetOperation Then
20470                 strTmp01 = .Properties("Description")
20480                 Set rst = .OpenRecordset
20490                 With rst
20500                   If .BOF = True And .EOF = True Then
20510                     lngRecs = 0&
20520                   Else
20530                     .MoveLast
20540                     lngRecs = .RecordCount
20550                   End If
20560                   strTmp01 = strTmp01 & CStr(lngRecs) & IIf(lngRecs = 0, "!", ".")
20570                   .Close
20580                 End With  ' ** rst.
20590                 .Properties("Description") = strTmp01
20600               End If
20610             End With  ' ** qdf.
20620             Set rst = Nothing
20630             Set qdf = Nothing
20640           Next  ' ** lngX.

20650         End If  ' ** blnSkip.

20660       End If  ' ** lngQrys.

20670     End If  ' ** blnSkip.

20680     .Close
20690   End With  ' ** dbs.
20700   Set dbs = Nothing

20710   Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
20720   DoEvents

20730   Beep

20740   Debug.Print "'DONE!"
20750   DoEvents

EXITP:
20760   Set rst = Nothing
20770   Set prp = Nothing
20780   Set qdf = Nothing
20790   Set dbs = Nothing
20800   Qry_Copy = blnRetVal
20810   Exit Function

ERRH:
20820   blnRetVal = False
20830   Select Case ERR.Number
        Case Else
20840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20850   End Select
20860   Resume EXITP

End Function

Public Function Qry_DDL_From_Tbl() As Boolean

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_DDL_From_Tbl"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst As DAO.Recordset, prp As Object
        Dim strQryName As String, strQryNumBase_Old As String, strQryNumBase_New As String, strQryNum As String
        Dim lngRecs As Long, lngQrysCreated As Long
        Dim strTmp01 As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

        Const QRY_BASE As String = "zz_qry_System_"

20910 On Error GoTo 0

20920   blnRetVal = True

20930   Set dbs = CurrentDb
20940   With dbs

20950     strQryNumBase_New = "59"

          ' ** zz_qry_System_50_01 (tblQuery, just 'dbQDDL'), just zz_tbl_Form_Property,
          ' ** zz_tbl_Form_Property_Value, zz_tbl_VBComponent_KeyDown.
20960     Set qdf1 = .QueryDefs("zz_qry_System_50_02")
20970     Set rst = qdf1.OpenRecordset
20980     rst.MoveLast
20990     lngRecs = rst.RecordCount
21000     rst.MoveFirst
21010     strQryNumBase_Old = vbNullString: lngQrysCreated = 0&
21020     For lngX = 1& To lngRecs
21030       strQryName = rst![qry_name]  ' ** qryXAdmin_DDL_52_01.
21040       strQryNum = Right(strQryName, 2)
21050       strTmp01 = Mid(strQryName, 15, 2)  ' ** '52'
21060       If strTmp01 <> strQryNumBase_Old Then
21070         strQryNumBase_Old = strTmp01
21080         strQryNumBase_New = CStr(Val(strQryNumBase_New) + 1)
21090       End If
21100       strQryName = QRY_BASE & strQryNumBase_New & "_" & strQryNum
21110       Set qdf2 = .CreateQueryDef(strQryName, rst![qry_sql])
21120       With qdf2
21130         Set prp = .CreateProperty("Description", dbText, rst![qry_description])
21140 On Error Resume Next
21150         .Properties.Append prp
21160         If ERR.Number <> 0 Then
21170 On Error GoTo 0
21180           .Properties("Description") = rst![qry_description]
21190         Else
21200 On Error GoTo 0
21210         End If
21220       End With
21230       Set qdf2 = Nothing
21240       lngQrysCreated = lngQrysCreated + 1&
21250       If lngX < lngRecs Then rst.MoveNext
21260     Next
21270     rst.Close

21280     .Close
21290   End With

21300   Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
21310   DoEvents

21320   Beep
21330   Debug.Print "'DONE!"

EXITP:
21340   Set prp = Nothing
21350   Set rst = Nothing
21360   Set qdf1 = Nothing
21370   Set qdf2 = Nothing
21380   Set dbs = Nothing
21390   Qry_DDL_From_Tbl = blnRetVal
21400   Exit Function

ERRH:
21410   blnRetVal = False
21420   Select Case ERR.Number
        Case Else
21430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21440   End Select
21450   Resume EXITP

End Function

Public Function Qry_Del_rel(Optional varQry As Variant) As Boolean
' ** Called by:
' **   mcrDelete_Query

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_Del_rel"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strFind As String
        Dim intQrys As Integer, arr_varQry() As Variant
        Dim msgResponse As VbMsgBoxResult
        Dim blnDeleteAll As Boolean
        Dim intX As Integer
        Dim blnRetVal As Boolean

21510   blnRetVal = True

21520   If IsMissing(varQry) = True Then
21530     strFind = InputBox("Enter Query Name Partial:", "Delete Query", vbNullString)
21540   Else
21550     strFind = varQry
21560   End If

21570   If strFind <> vbNullString Then

21580     If strFind = "zz" Or strFind = "zz_" Or strFind = "zz_qry" Or strFind = "zz_qry_" Then
21590       msgResponse = MsgBox("Delete all zz_qry's except for zz_qry_System_nn?", vbQuestion + vbYesNoCancel, "Leave System Queries")
21600       Select Case msgResponse
            Case vbYes
21610         blnDeleteAll = False
21620       Case vbNo
21630         blnDeleteAll = True
21640       Case Else
21650         blnRetVal = False
21660       End Select
21670     Else
21680       blnDeleteAll = True
21690     End If

21700     If blnRetVal = True Then

21710       intX = 0
21720       intQrys = 0
21730       ReDim arr_varQry(1, 0)

21740       Set dbs = CurrentDb

21750       For Each qdf In dbs.QueryDefs
21760         With qdf
21770           intX = intX + 1
21780           If Left(.Name, 1) <> "~" Then
21790             If InStr(.Name, strFind) > 0 Then
21800               If blnDeleteAll = False And Left(.Name, 14) = "zz_qry_System_" Then
                      ' ** Skip the administrative queries.
21810               Else
21820                 intQrys = intQrys + 1
21830                 ReDim Preserve arr_varQry(1, (intQrys - 1))
21840                 arr_varQry(0, (intQrys - 1)) = .Name
21850                 arr_varQry(1, (intQrys - 1)) = intX
21860               End If
21870             End If
21880           End If
21890         End With
21900       Next

21910       dbs.Close

21920       If intQrys > 0 Then
21930         Beep
21940         If MsgBox(CStr(intQrys) & " queries were found containing string:" & vbCrLf & vbCrLf & _
                  "    " & strFind & vbCrLf & vbCrLf & "Delete them?", _
                  vbQuestion + vbYesNo + vbDefaultButton1, "Queries Found") = vbYes Then
21950           SysCmd acSysCmdInitMeter, "Deleting... ", intQrys
21960           For intX = 0 To (intQrys - 1)
21970             SysCmd acSysCmdUpdateMeter, (intX + 1)
21980             DoCmd.DeleteObject acQuery, arr_varQry(0, intX)
21990           Next
22000           Beep
22010           MsgBox "Finished", vbExclamation + vbOKOnly, ("Finished" & Space(40))
22020           SysCmd acSysCmdClearStatus
22030         End If

22040       Else
22050         Beep
22060         MsgBox "None were found containing string:" & vbCrLf & vbCrLf & _
                "    " & strFind, vbInformation + vbOKOnly, "Query Not Found"
22070       End If

22080     End If

22090   End If

EXITP:
22100   Beep
22110   Set qdf = Nothing
22120   Set dbs = Nothing
22130   Qry_Del_rel = blnRetVal
22140   Exit Function

ERRH:
22150   blnRetVal = False
22160   Select Case ERR.Number
        Case Else
22170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22180   End Select
22190   Resume EXITP

End Function

Public Function Qry_FldList_rel(Optional varQryName As Variant) As Variant

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_FldList_rel"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, fld As DAO.Field
        Dim strQryName As String
        Dim blnRetArray As Boolean
        Dim lngFlds As Long, arr_varFld() As Variant
        Dim lngE As Long
        Dim varRetVal As Variant

        ' ** Array: arr_varFld().
        Const F_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const F_FNAM As Integer = 0

22210   varRetVal = Null

22220   Select Case IsMissing(varQryName)
        Case True
22230     strQryName = "qryPrintChecks_05_10"
22240     blnRetArray = False
22250   Case False
22260     strQryName = varQryName
22270     blnRetArray = True
22280   End Select

22290   lngFlds = 0&
22300   ReDim arr_varflds(F_ELEMS, 0&)

22310   If blnRetArray = False Then
22320     Debug.Print "'QRY: " & strQryName
22330   End If

22340   Set dbs = CurrentDb
22350   With dbs
22360     Set qdf = .QueryDefs(strQryName)
22370     With qdf
22380       For Each fld In .Fields
22390         With fld
22400           Select Case blnRetArray
                Case True
22410             lngFlds = lngFlds + 1&
22420             lngE = lngFlds - 1&
22430             ReDim Preserve arr_varFld(F_ELEMS, lngE)
22440             arr_varFld(F_FNAM, lngE) = .Name
22450           Case False
22460             Debug.Print "'" & .Name
22470           End Select
22480         End With
22490       Next
22500     End With
22510     .Close
22520   End With

22530   Select Case blnRetArray
        Case True
22540     varRetVal = arr_varFld
22550   Case False
22560     Debug.Print "'DONE!"
22570     varRetVal = "True"
22580     Beep
22590   End Select

EXITP:
22600   Set fld = Nothing
22610   Set qdf = Nothing
22620   Set dbs = Nothing
22630   Qry_FldList_rel = varRetVal
22640   Exit Function

ERRH:
22650   varRetVal = RET_ERR
22660   Select Case ERR.Number
        Case Else
22670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22680   End Select
22690   Resume EXITP

End Function

Public Function Qry_FindDesc_rel() As Boolean
' ** Find a string within a query's description.

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_FindDesc_rel"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object
        Dim strFind As String
        Dim strDesc As String

        Dim blnRetVal As Boolean

22710   blnRetVal = True

22720   strFind = "rptCourtRptNS_00D"

22730   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

22740   Set dbs = CurrentDb
22750   With dbs
22760     For Each qdf In .QueryDefs
22770       With qdf
22780 On Error Resume Next
22790         Set prp = .Properties("Description")
22800         If ERR.Number = 0 Then
22810 On Error GoTo ERRH
22820           strDesc = .Properties("Description")  ' ** DON'T TRIM!! I DON'T WANT TO LOSE THE INDENT!
22830           If strDesc <> vbNullString Then
22840             If InStr(strDesc, strFind) > 0 Then
22850               Debug.Print "'QRY: " & qdf.Name & "  '" & strDesc & "'"
22860               DoEvents
22870             End If
22880           End If
22890         Else
22900 On Error GoTo ERRH
22910         End If
22920       End With
22930     Next
22940     .Close
22950   End With

22960   Beep

EXITP:
22970   Set prp = Nothing
22980   Set qdf = Nothing
22990   Set dbs = Nothing
23000   Qry_FindDesc_rel = blnRetVal
23010   Exit Function

ERRH:
23020   blnRetVal = False
23030   Select Case ERR.Number
        Case Else
23040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23050   End Select
23060   Resume EXITP

End Function

Public Function Qry_FindInMod() As Boolean

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_FindInMod"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim lngLines As Long
        Dim strFind As String, strLine As String, strModName As String
        Dim blnFound As Boolean
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim intAsc As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

23110 On Error GoTo 0

23120   blnRetVal = True

23130   strFind = "zz_mod_QuerySQLDocFuncs"

23140   If Left(strFind, 3) = "frm" Then
23150     strFind = "Form_" & strFind
23160   ElseIf Left(strFind, 3) = "rpt" Then
23170     strFind = "Report_" & strFind
23180   End If

23190   Set vbp = Application.VBE.ActiveVBProject
23200   With vbp
23210     Set vbc = .VBComponents(strFind)
23220     With vbc
23230       strModName = .Name
23240       Debug.Print "'MOD: " & strModName
23250       DoEvents
23260       Set cod = .CodeModule
23270       With cod
23280         lngLines = .CountOfLines
23290         For lngX = 1& To lngLines
23300           strLine = Trim(.Lines(lngX, 1))
23310           strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
23320           intPos01 = 0: intPos02 = 0: intPos03 = 0
23330           If strLine <> vbNullString Then
23340             If Left(strLine, 1) <> "'" Then
23350               intPos01 = InStr(strLine, "zzz_qry")
23360               If intPos01 > 0 Then
23370                 strTmp01 = Mid(strLine, intPos01)
23380               End If
23390               If intPos01 > 0 Then
23400                 intPos02 = InStr((intPos01 + 5), strLine, "zz_qry")  ' ** Could be before, but unlikely.
23410               Else
23420                 intPos02 = InStr(strLine, "zz_qry")
23430               End If
23440               If intPos02 > 0 Then
23450                 strTmp02 = Mid(strLine, intPos02)
23460               End If
23470               If intPos01 = 0 And intPos02 = 0 Then
                      ' ** If one of those was found, don't search for more.
23480                 intPos03 = InStr(strLine, "qry")
23490                 If intPos03 > 1 Then  ' ** It wouldn't be at the beginning of the line.
23500                   strTmp03 = Mid(strLine, intPos03)
23510                 End If
23520               End If
23530               If strTmp01 <> vbNullString Then
23540                 intPos04 = InStr(strTmp01, " ")
23550                 If intPos04 > 0 Then strTmp01 = Trim(Left(strTmp01, intPos04))
23560                 If Right(strTmp01, 2) = Chr(34) & ")" Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 2))
23570                 If Right(strTmp01, 1) = Chr(34) Then strTmp01 = Trim(Left(strTmp01, (Len(strTmp01) - 1)))
23580                 If Right(strTmp01, 1) = ":" Then strTmp01 = Trim(Left(strTmp01, (Len(strTmp01) - 1)))
23590                 blnFound = True
23600                 Select Case strTmp01
                      Case "zzz_qry"
23610                   blnFound = False
23620                 End Select
23630                 If blnFound = True Then
23640                   Debug.Print "'" & strTmp01
23650                 End If
23660               End If
23670               If strTmp02 <> vbNullString Then
23680                 intPos04 = InStr(strTmp02, " ")
23690                 If intPos04 > 0 Then strTmp02 = Trim(Left(strTmp02, intPos04))
23700                 If Right(strTmp02, 2) = Chr(34) & ")" Then strTmp02 = Left(strTmp02, (Len(strTmp02) - 2))
23710                 If Right(strTmp02, 1) = Chr(34) Then strTmp02 = Trim(Left(strTmp02, (Len(strTmp02) - 1)))
23720                 If Right(strTmp02, 1) = ":" Then strTmp02 = Trim(Left(strTmp02, (Len(strTmp02) - 1)))
23730                 blnFound = True
23740                 Select Case strTmp02
                      Case "zz_qry"
23750                   blnFound = False
23760                 End Select
23770                 If blnFound = True Then
23780                   Debug.Print "'" & strTmp02
23790                 End If
23800               End If
23810               If strTmp03 <> vbNullString Then
23820                 intPos04 = InStr(strTmp03, " ")
23830                 If intPos04 > 0 Then strTmp03 = Trim(Left(strTmp03, intPos04))
23840                 If Right(strTmp03, 2) = Chr(34) & ")" Then strTmp03 = Trim(Left(strTmp03, (Len(strTmp03) - 2)))
23850                 If Right(strTmp03, 1) = ")" Then strTmp03 = Trim(Left(strTmp03, (Len(strTmp03) - 1)))
23860                 If Right(strTmp03, 1) = Chr(34) Then strTmp03 = Trim(Left(strTmp03, (Len(strTmp03) - 1)))
23870                 If Right(strTmp03, 1) = "]" Then strTmp03 = Trim(Left(strTmp03, (Len(strTmp03) - 1)))
23880                 If Right(strTmp03, 1) = ":" Then strTmp03 = Trim(Left(strTmp03, (Len(strTmp03) - 1)))
23890                 If Right(strTmp03, 1) = "," Then strTmp03 = Trim(Left(strTmp03, (Len(strTmp03) - 1)))
23900                 If Right(strTmp03, 1) = Chr(34) Then strTmp03 = Trim(Left(strTmp03, (Len(strTmp03) - 1)))  ' ** Again.
23910                 If Right(strTmp03, 1) = "]" Then strTmp03 = Trim(Left(strTmp03, (Len(strTmp03) - 1)))  ' ** Again.
23920                 intPos04 = InStr(strTmp03, "(")
23930                 If intPos04 > 0 Then strTmp03 = Trim(Left(strTmp03, (intPos04 - 1)))
23940                 strTmp04 = Mid(strLine, (intPos03 - 1), 1)  ' ** Character before 'qry'.
23950                 blnFound = True
                      ' ** Numbers: 48 - 57; Upper-Case Letters: 65 - 90; Lower-Case Letters: 97 - 122.
23960                 intAsc = Asc(strTmp04)
23970                 If (intAsc >= 48 And intAsc <= 57) Or (intAsc >= 65 And intAsc <= 90) Or (intAsc >= 97 And intAsc <= 122) Then
                        ' ** This isn't a query.
23980                   blnFound = False
23990                 Else
24000                   If strTmp04 = Chr(34) Then  ' ** Quotes.
                          ' ** OK.
24010                   ElseIf strTmp04 = "[" Then
                          ' ** No.
24020                     blnFound = False
24030                   End If
24040                   If Right(strTmp03, 1) = "(" Or Right(strTmp03, 1) = "." Then
                          ' ** No.
24050                     blnFound = False
24060                   End If
24070                 End If
24080                 If blnFound = True Then
24090                   Select Case strTmp03
                        Case "qry_id", "qry_name", "qry_sql", "qry_param", "qry_param_clause", "qry_paramcnt", _
                            "qry_formref", "qry_formrefcnt", "qry_fldcnt", "qry_datemodified", "qrytype_type"
                          ' ** No.
24100                     blnFound = False
24110                   Case "qryparam_id", "qryparam_name", "qryparam_order", "qryparam_sql", "qryparam_clause", _
                            "qryparam_datemodified", "qry_id_recsrc", "qryrecsrc_order", "qryrecsrc_name", _
                            "qryrecsrc_datemodified", "qryfld_id", "qryfld_name", "qryfld_format", "qryfld_datemodified"
                          ' ** No.
24120                     blnFound = False
24130                   Case "QRY", "QRYS", "QRYX", "zz_qry", "zzz_qry"
                          ' ** No.
24140                     blnFound = False
24150                   Case "Qry_Doc", "Qry_Parm_Doc", "Qry_Tbl_Doc", "Qry_ParseFormRef", "Qry_FldDoc", _
                            "Qry_CurrentAppName", "Qry_ChkDocQrys", "Qry_TblChk1", "Qry_TblChk2"
                          ' ** No.
24160                     blnFound = False
24170                   End Select
24180                 End If  ' ** blnFound.
24190                 If blnFound = True Then
24200                   Debug.Print "'" & strTmp03
24210                 End If  ' ** blnFound.
24220               End If  ' ** strTmp03
24230             End If  ' ** Remark.
24240           End If  ' ** vbNullString.
24250         Next  ' ** lngX.
24260       End With  ' ** cod.
24270     End With  ' ** vbc.
24280   End With  ' ** vbp

24290   Debug.Print "'DONE!"

24300   Beep

EXITP:
24310   Set cod = Nothing
24320   Set vbc = Nothing
24330   Set vbp = Nothing
24340   Set rst = Nothing
24350   Set qdf = Nothing
24360   Set dbs = Nothing
24370   Qry_FindInMod = blnRetVal
24380   Exit Function

ERRH:
24390   blnRetVal = False
24400   Select Case ERR.Number
        Case Else
24410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24420   End Select
24430   Resume EXITP

End Function

Public Function Qry_FindStr_rel(Optional varRenFind1 As Variant) As Variant
' ** Find a specified string within a query's SQL.

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_FindStr_rel"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngRefs As Long, arr_varRef() As Variant
        Dim strFind1 As String, strQryType As String
        Dim strSQL As String
        Dim blnFound1 As Boolean, blnFound2 As Boolean
        'Dim strType1 As String, strType2 As String
        'Dim strFld1 As String, strFld2 As String
        Dim blnFindAnd As Boolean, blnFindOr As Boolean, blnFindOne As Boolean
        Dim blnCalled As Boolean
        Dim intPos01 As Integer, intPos02 As Integer
        Dim strTmp01 As String
        Dim lngE As Long
        Dim varRetVal As Variant

        ' ** Array: arr_varRef().
        Const RF_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const RF_QRY As Integer = 0
        Const RF_SQL As Integer = 1

24510 On Error GoTo 0

24520   If IsMissing(varRenFind1) = True Then
24530     strFind1 = "tblTransaction_Audit_Filter"
'QRY: 'qryTransaction_Audit_08_01' tblTransaction_Audit_Filter  dbQSelect
'QRY: 'qryTransaction_Audit_08_02' tblTransaction_Audit_Filter  dbQSelect
'DONE!
          'strFind2 = "Invoice#"
24540     blnCalled = False
24550     varRetVal = False
24560   Else
24570     strFind1 = varRenFind1
24580     blnCalled = True
24590   End If

24600   Set dbs = CurrentDb

24610   blnFindAnd = True
24620   blnFindOr = False
24630   blnFindOne = False

24640   lngRefs = 0&
24650   ReDim arr_varRef(RF_ELEMS, 0)
24660   arr_varRef(RF_QRY, 0) = "0"

24670   With dbs
24680     For Each qdf In .QueryDefs
24690       blnFound1 = False: blnFound2 = False
            'strType1 = vbNullString: strType2 = vbNullString
            'strFld1 = vbNullString: strFld2 = vbNullString
24700       With qdf
24710         If Left(.Name, 4) <> "~TMP" Then  ' ** Skip those pesky system queries!
24720           If strFind1 = "tblForm_Graphics" And Left(.Name, 26) = "qryXAdmin_Form_Graphics_11" Then
                  ' ** Skip these!
24730           Else
24740             strSQL = .SQL
24750             If Left(.Name, 1) <> "~" Then
24760               intPos01 = InStr(strSQL, vbCrLf)
24770               Do While intPos01 > 0
24780                 strSQL = Left(strSQL, (intPos01 - 1)) & " " & Mid(strSQL, (intPos01 + 2))
24790                 intPos01 = InStr(strSQL, vbCrLf)
24800               Loop
24810               intPos01 = InStr(strSQL, "  ")
24820               Do While intPos01 > 0
24830                 strSQL = Left(strSQL, intPos01) & Mid(strSQL, (intPos01 + 2))
24840                 intPos01 = InStr(strSQL, "  ")
24850               Loop
24860               intPos01 = InStr(strSQL, strFind1)
24870               If intPos01 > 0 Then
24880                 blnFound1 = True
24890                 strQryType = DLookup("[qrytype_constant]", "tblQueryType", "[qrytype_type] = " & CStr(.Type))
                      'strType1 = DLookup("[datatype_db_constant]", "tblDataTypeDb", "[datatype_db_type]=" & CStr(.Type))
                      'strFld1 = .Name
24900                 intPos02 = InStr((intPos01 + 1), strSQL, " ")
24910                 If intPos02 > 0 Then
24920                   strTmp01 = Mid(strSQL, intPos01, (InStr(intPos01 + 1, strSQL, " ") - intPos01))
24930                 Else
24940                   strTmp01 = Mid(strSQL, intPos01)
24950                 End If

24960                 intPos02 = InStr(strTmp01, "]")
24970                 If intPos02 > 0 Then strTmp01 = Left(strTmp01, (intPos02 - 1))
24980                 intPos02 = InStr(strTmp01, "!")
24990                 If intPos02 > 0 Then strTmp01 = Left(strTmp01, (intPos02 - 1))
25000                 intPos02 = InStr(strTmp01, ".")
25010                 If intPos02 > 0 Then strTmp01 = Left(strTmp01, (intPos02 - 1))
25020                 intPos02 = InStr(strTmp01, ",")
25030                 If intPos02 > 0 Then strTmp01 = Left(strTmp01, (intPos02 - 1))
25040                 intPos02 = InStr(strTmp01, "+")
25050                 If intPos02 > 0 Then strTmp01 = Left(strTmp01, (intPos02 - 1))
25060                 intPos02 = InStr(strTmp01, "/")
25070                 If intPos02 > 0 Then strTmp01 = Left(strTmp01, (intPos02 - 1))
25080                 If blnCalled = False Then
25090                   Debug.Print "'QRY: '" & qdf.Name & "' " & strTmp01 & "  " & strQryType
25100                   varRetVal = True
25110                 End If

                      'ElseIf InStr(strSQL, strFind2) > 0 Then
                      '  blnFound2 = True
                      '  strType2 = DLookup("[datatype_db_constant]", "tblDataTypeDb", "[datatype_db_type]=" & CStr(.Type))
                      '  strFld2 = .Name
25120               End If

                    'Debug.Print "'QRY: '" & .Name & "'"
25130               If blnFound1 = True Then  'And blnFound2 = True Then
25140                 lngRefs = lngRefs + 1&
25150                 lngE = lngRefs - 1&
25160                 ReDim Preserve arr_varRef(RF_ELEMS, lngE)
25170                 arr_varRef(RF_QRY, lngE) = qdf.Name
25180                 arr_varRef(RF_SQL, lngE) = qdf.SQL
                      'arr_varRef(1, lngE) = strFld1 & " " & strType1
                      'arr_varRef(2, lngE) = strFld2 & " " & strType2
25190               End If

25200             End If

25210           End If  ' ** Skip 11's.
25220         End If
25230       End With

25240     Next
25250     .Close
25260   End With

25270   If blnCalled = True Then
25280     varRetVal = arr_varRef
25290   End If

25300   If blnCalled = False Then
25310     Debug.Print "'DONE!"
25320     Beep
25330   End If

EXITP:
25340   Set qdf = Nothing
25350   Set dbs = Nothing
25360   Qry_FindStr_rel = varRetVal
25370   Exit Function

ERRH:
25380   Select Case ERR.Number
        Case Else
25390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25400   End Select
25410   Resume EXITP

End Function

Public Function Qry_Export_rel() As Boolean
' ** Copies a group of queries from here to another MDB.

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_Export_rel"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim strMDB As String
        Dim lngHits As Long ', lngQrysCreated As Long
        Dim blnSkip As Boolean
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const Q_QNAM As Integer = 0
        Const Q_SQL  As Integer = 1
        Const Q_ERR1 As Integer = 2
        Const Q_ERR2 As Integer = 3
        Const Q_DESC As Integer = 4

25510 On Error GoTo 0

25520   blnRetVal = True

25530   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
25540   DoEvents

25550   DBEngine.SystemDB = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"

25560   strMDB = "C:\Program Files\Delta Data\Trust Accountant\TrustAux.mdb"

25570   blnSkip = False
25580   If blnSkip = False Then

25590     lngQrys = 0&
25600     ReDim arr_varQry(Q_ELEMS, 0)

25610     Set dbs = CurrentDb
25620     With dbs
25630       For Each qdf In .QueryDefs
25640         With qdf
25650           blnSkip = True
25660           If Left(.Name, 1) = "~" Then  ' ** Skip those pesky system queries.
25670             blnSkip = True
25680           Else
25690             If Left(.Name, 11) = "qryVersion_" And Left(.Name, 18) <> "qryVersion_Convert" Then
25700               blnSkip = False
25710             End If
25720           End If
25730           If blnSkip = False Then
25740             lngQrys = lngQrys + 1&
25750             lngE = lngQrys - 1&
25760             ReDim Preserve arr_varQry(Q_ELEMS, lngE)
25770             arr_varQry(Q_QNAM, lngE) = .Name
25780           End If
25790         End With
25800       Next
25810       .Close
25820     End With
25830     Set dbs = Nothing

25840     Debug.Print "'QRYS: " & CStr(lngQrys)
25850     DoEvents

25860     If lngQrys > 0& Then

25870       Set dbs = CurrentDb
25880       Debug.Print "'|";
25890       DoEvents
25900       lngHits = 0&

25910       blnSkip = True
25920       If blnSkip = False Then
25930         For lngX = 0& To (lngQrys - 1&)
25940           If arr_varQry(Q_QNAM, lngX) <> "qryVersion_20" Then
25950 On Error Resume Next
25960             DoCmd.TransferDatabase acExport, "Microsoft Access", strMDB, acQuery, arr_varQry(Q_QNAM, lngX), arr_varQry(Q_QNAM, lngX)
25970             If ERR.Number <> 0 Then
25980 On Error GoTo 0
25990               Debug.Print "'ERR: " & arr_varQry(Q_QNAM, lngX)
26000             Else
26010 On Error GoTo 0
26020             End If
26030             lngHits = lngHits + 1&
26040             If (lngHits Mod 1000) = 0 Then
26050               Debug.Print "|  " & CStr(lngHits) & " of " & CStr(lngQrys)
26060               Debug.Print "'|";
26070             ElseIf (lngHits Mod 100) = 0 Then
26080               Debug.Print "|";
26090             ElseIf (lngHits Mod 10) = 0 Then
26100               Debug.Print ".";
26110             End If
26120             DoEvents
26130           End If
26140         Next
26150         Debug.Print
26160       End If  ' ** blnSkip.

26170       Debug.Print "'QEYS COPIED: " & CStr(lngHits)

26180       lngHits = 0&
26190       For lngX = (lngQrys - 1&) To 0& Step -1&
26200         DoCmd.DeleteObject acQuery, arr_varQry(Q_QNAM, lngX)
26210         DoEvents
26220         lngHits = lngHits + 1&
26230       Next

26240       Debug.Print "'QRYS DELETED: " & CStr(lngHits)
26250       DoEvents

26260     End If
26270   End If  ' ** blnSkip.

        'DBEngine.SystemDB = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"

26280   Debug.Print "'DONE!"
26290   DoEvents

26300   Beep

EXITP:
26310   Set qdf = Nothing
26320   Set dbs = Nothing
26330   Qry_Export_rel = blnRetVal
26340   Exit Function

ERRH:
26350   blnRetVal = False
26360   DoCmd.SetWarnings True
26370   Select Case ERR.Number
        Case Else
26380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26390   End Select
26400   Resume EXITP

End Function

Public Function Qry_Import_rel() As Boolean
' ** Copies a group of queries from another MDB to here.

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_Import_rel"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object
        Dim strMDB As String, strDesc As String
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim lngHits As Long, lngQrysCreated As Long
        Dim blnSkip As Boolean
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const Q_QNAM As Integer = 0
        Const Q_SQL  As Integer = 1
        Const Q_ERR1 As Integer = 2
        Const Q_ERR2 As Integer = 3
        Const Q_DESC As Integer = 4

26510 On Error GoTo 0

26520   blnRetVal = True

26530   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
26540   DoEvents

26550   DBEngine.SystemDB = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"

'32360   strMDB = "C:\Program Files\Delta Data\Trust Accountant\Client_Frontends\Trust_c_wRptLstTransAudQrys.mdb"
'32360   strMDB = "C:\Program Files\Delta Data\Trust Accountant\Trust - Copy (7).mdb"
'32360   strMDB = "C:\Program Files\Delta Data\Trust Accountant\Trust_bak53.mdb"
'32360   strMDB = "C:\Program Files\Delta Data\Trust Accountant\Trust_mastertrust9.mdb"
26560   strMDB = "C:\Program Files\Delta Data\Trust Accountant\Trust_a_newXX.mdb"

26570   blnSkip = False
26580   If blnSkip = False Then

26590     lngQrys = 0&
26600     ReDim arr_varQry(Q_ELEMS, 0)

26610     Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)
26620     Set dbs = wrk.OpenDatabase(strMDB, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
26630     With dbs
26640       For Each qdf In .QueryDefs
26650         With qdf
26660           blnSkip = True
26670           If Left(.Name, 1) = "~" Then  ' ** Skip those pesky system queries.
26680             blnSkip = True
26690           Else
                  'blnSkip = True
                  'Select Case .Name
                  'Case "qryAccountProfile_TaxCodes_Block_03_02", "qryAccountProfile_TaxCodes_Block_03_03", "qryAccountProfile_TaxCodes_Block_03_04", "qryAccountProfile_TaxCodes_Block_05_01", _
                  '    "qryAccountProfile_TaxCodes_Block_05_02", "qryAccountProfile_TaxCodes_Block_05_03", "qryAccountProfile_TaxCodes_Block_05_04", "qryAccountProfile_TaxCodes_Block_06"
                  '  blnSkip = False
                  'Case "qryAccountProfile_TaxCodes_Block_07", "qryAccountProfile_TaxCodes_Block_08_01", "qryAccountProfile_TaxCodes_Block_08_02", "qryAccountProfile_TaxCodes_Block_08_03", _
                  '    "qryAccountProfile_TaxCodes_Block_08_04", "qryAccountProfile_TaxCodes_Block_08_05", "qryAccountProfile_TaxCodes_Block_08_06", "qryAccountProfile_Transactions_01_01"
                  '  blnSkip = False
                  'Case "qryAccountProfile_Transactions_01_02", "qryAccountProfile_Transactions_02_01", "qryAccountProfile_Transactions_02_02", "qryAccountProfile_Transactions_03", _
                  '    "qryAccountProfile_Transactions_03_01", "qryAccountProfile_Transactions_03_02", "qryAccountProfile_Transactions_03_03", "qryAccountProfile_Transactions_04_01"
                  '  blnSkip = False
                  'Case "qryAccountProfile_Transactions_04_02", "qryAccountProfile_Transactions_05_01", "qryAccountProfile_Transactions_05_02", "qryAccountProfile_Transactions_06_01", _
                  '    "qryAccountProfile_Transactions_06_02", "qryAccountProfile_Transactions_06_03", "qryAccountProfile_Transactions_06_04", "qryAccountProfile_Transactions_07"
                  '  blnSkip = False
                  'Case "qryAccountProfile_Transactions_08", "qryAccountReviews_01", "qryAccountSearch_01", "qryAccountSummary", "qryAccountSummary_01", "qryAccountSummary_02", _
                  '    "qryAccountSummary_02a", "qryAccountSummary_02b", "qryAccountSummary_03", "qryAccountSummary_03_01"
                  '  blnSkip = False
                  'End Select
                  'If Left(.Name, 23) = "zz_qry_VBComponent_Var_" Then
                  '  blnSkip = False
                  'End If
                  'If Left(.Name, 15) = "qryPrintChecks_" Then
                  '  blnSkip = False
                  'End If
                  'If .Name = "zzz_qry_zForm_Control_03_04_13_06" Then
                  '  blnSkip = False
                  'End If
                  'If Left(.Name, 2) = "zz" Then
                  '  If InStr(.Name, "MasterTrust") > 0 Then
26700             blnSkip = False
                  '  End If
                  'End If
                  'If InStr(.Name, "Germantown") > 0 Or InStr(.Name, "MasterTrust") > 0 Or InStr(.Name, "Fiduciary") > 0 Or _
                  '    InStr(.Name, "Bluffs") > 0 Or InStr(.Name, "DemoData") > 0 Then
                  '  blnSkip = True
                  'End If
                  'If Left(.Name, 13) <> "zz_qry_System" Then
                  '  blnSkip = True
                  'End If
                  'If InStr(.Name, "Ohana") > 0 Then
                  '  blnSkip = True
                  'End If
26710           End If
26720           If blnSkip = False Then
26730             lngQrys = lngQrys + 1&
26740             lngE = lngQrys - 1&
26750             ReDim Preserve arr_varQry(Q_ELEMS, lngE)
26760             arr_varQry(Q_QNAM, lngE) = .Name
      'On Error Resume Next
                  'arr_varQry(Q_SQL, lngE) = .SQL
                  'If ERR.Number <> 0 Then
      'On Error GoTo 0
                  '  arr_varQry(Q_ERR1, lngE) = CBool(True)
                  'Else
      'On Error GoTo 0
                  '  arr_varQry(Q_ERR1, lngE) = CBool(False)
                  'End If
                  'arr_varQry(Q_ERR2, lngE) = CBool(False)
26770             strDesc = vbNullString
26780 On Error Resume Next
26790             strDesc = .Properties("Description")
26800             If ERR.Number = 0 Then
26810 On Error GoTo 0
26820               If strDesc <> vbNullString Then
26830                 arr_varQry(Q_DESC, lngE) = strDesc
26840               Else
26850                 arr_varQry(Q_DESC, lngE) = Null
26860               End If
26870             Else
26880 On Error GoTo 0
26890               arr_varQry(Q_DESC, lngE) = Null
26900             End If
26910           End If
26920         End With
26930       Next
26940       .Close
26950     End With
26960     Set dbs = Nothing
26970     wrk.Close
26980     Set wrk = Nothing

26990     Debug.Print "'QRYS: " & CStr(lngQrys)
27000     DoEvents

27010     If lngQrys > 0& Then
27020       Set dbs = CurrentDb
27030       Debug.Print "'|";
27040       DoEvents
27050       lngHits = 0&: lngQrysCreated = 0&
27060       For lngX = 0& To (lngQrys - 1&)
27070         blnSkip = False
27080         If IsNull(arr_varQry(Q_DESC, lngX)) = True Then
27090           blnSkip = True
27100         End If
27110         If blnSkip = False Then
27120 On Error Resume Next
27130           Set qdf = dbs.QueryDefs(arr_varQry(Q_QNAM, lngX))
27140           If ERR.Number = 0 Then
27150 On Error GoTo 0
27160             With qdf
27170               Set prp = .CreateProperty("Description", dbText, arr_varQry(Q_DESC, lngX))
27180 On Error Resume Next
27190               .Properties.Append prp
27200               If ERR.Number <> 0 Then
27210 On Error GoTo 0
27220                 .Properties("Description") = arr_varQry(Q_DESC, lngX)
27230               Else
27240 On Error GoTo 0
27250               End If
27260               lngQrysCreated = lngQrysCreated + 1&
27270             End With
27280           Else
27290 On Error GoTo 0
27300           End If
27310           Set qdf = Nothing
27320         End If
              'If QueryExists(CStr(arr_varQry(Q_QNAM, lngX))) = True Then  ' ** Module Function: modFileUtilities.
              'Set qdf = dbs.QueryDefs(arr_varQry(Q_QNAM, lngX))
              'If qdf.SQL <> arr_varQry(Q_SQL, lngX) Then
              '  qdf.SQL = arr_varQry(Q_SQL, lngX)
              '  lngQrysCreated = lngQrysCreated + 1&
              'End If
              'Set qdf = Nothing
              'DoCmd.DeleteObject acQuery, arr_varQry(Q_QNAM, lngX)
              'DoEvents
              'CurrentDb.QueryDefs.Refresh
              'End If
27330         blnSkip = True
27340         If blnSkip = False Then
27350 On Error Resume Next
27360           DoCmd.TransferDatabase acImport, "Microsoft Access", strMDB, acQuery, arr_varQry(Q_QNAM, lngX), arr_varQry(Q_QNAM, lngX)
27370           If ERR.Number <> 0 Then
27380 On Error GoTo 0
27390             arr_varQry(Q_ERR2, lngX) = CBool(True)
27400           Else
27410 On Error GoTo 0
27420           End If
27430         End If  ' ** blnSkip.
27440         lngHits = lngHits + 1&
27450         If (lngHits Mod 1000) = 0 Then
27460           Debug.Print "|  " & CStr(lngHits) & " of " & CStr(lngQrys)
27470           Debug.Print "'|";
27480         ElseIf (lngHits Mod 100) = 0 Then
27490           Debug.Print "|";
27500         ElseIf (lngHits Mod 10) = 0 Then
27510           Debug.Print ".";
27520         End If
27530         DoEvents
27540       Next
27550       Debug.Print
27560       CurrentDb.QueryDefs.Refresh
            'For lngX = 0& To (lngQrys - 1&)
            '  If arr_varQry(Q_ERR1, lngX) = True Then
            '    Debug.Print "'PERM ERR!  '" & arr_varQry(Q_QNAM, lngX) & "'"
            '  End If
            'Next
27570     End If

'33160     Debug.Print "'QEYS COPIED: " & CStr(lngHits)
27580     Debug.Print "'QRYS EDITED: " & CStr(lngQrysCreated)

27590   End If  ' ** blnSkip.

        'DBEngine.SystemDB = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"

27600   Debug.Print "'DONE!"
27610   DoEvents

27620   Beep

EXITP:
27630   Set prp = Nothing
27640   Set qdf = Nothing
27650   Set dbs = Nothing
27660   Set wrk = Nothing
27670   Qry_Import_rel = blnRetVal
27680   Exit Function

ERRH:
27690   blnRetVal = False
27700   DoCmd.SetWarnings True
27710   Select Case ERR.Number
        Case Else
27720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27730   End Select
27740   Resume EXITP

End Function

Public Function Qry_List() As Boolean

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_List"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngQrys As Long
        Dim blnRetVal As Boolean

27810   blnRetVal = True

27820   Set dbs = CurrentDb
27830   With dbs
27840     lngQrys = 0&
27850     For Each qdf In .QueryDefs
27860       With qdf
              ' ** qryStatementOfCondition_37e.
27870         If Left(.Name, 24) = "qryStatementOfCondition_" Then
27880           lngQrys = lngQrys + 1&
27890         End If
27900       End With
27910     Next
27920     .Close
27930   End With

27940   Debug.Print "'QRY CNT: " & CStr(lngQrys)
        'QRY CNT: 83 !!
        'QRY CNT: 130

27950   Beep

EXITP:
27960   Set qdf = Nothing
27970   Set dbs = Nothing
27980   Qry_List = blnRetVal
27990   Exit Function

ERRH:
28000   blnRetVal = False
28010   Select Case ERR.Number
        Case Else
28020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28030   End Select
28040   Resume EXITP

End Function

Public Function Qry_LoadDDL() As Variant

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_LoadDDL"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef
        Dim lngQryDdls As Long, arr_varQryDdl() As Variant
        Dim lngQryIdxs As Long, arr_varQryIdx() As Variant
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long, lngE As Long, lngF As Long
        Dim varRetVal As Variant

        Const QRY_SYS As String = "zz_qry_System_"
        Const QRY_DDL As String = "CREATE TABLE "

        ' ** Array: arr_varQryDdl().
        Const QD_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const QD_QNAM As Integer = 0
        Const QD_TNAM As Integer = 1
        Const QD_IDXS As Integer = 2
        Const QD_IARR As Integer = 3

28110   varRetVal = Empty

28120   lngQryDdls = 0&
28130   ReDim arr_varQryDdl(QD_ELEMS, 0)

28140   Set dbs = CurrentDb
28150   With dbs
28160     For Each qdf1 In .QueryDefs
28170       With qdf1
28180         If Left(.Name, Len(QRY_SYS)) = QRY_SYS And .Type = dbQDDL Then
28190           If Right(.Name, 3) = "_01" And Left(.SQL, Len(QRY_DDL)) = QRY_DDL Then
28200             lngQryDdls = lngQryDdls + 1&
28210             lngE = lngQryDdls - 1&
28220             ReDim Preserve arr_varQryDdl(QD_ELEMS, lngE)
28230             arr_varQryDdl(QD_QNAM, lngE) = .Name
28240             arr_varQryDdl(QD_TNAM, lngE) = Left(Mid(.SQL, (Len(QRY_DDL) + 1)), (InStr(Mid(.SQL, (Len(QRY_DDL) + 1)), " ") - 1))
28250             strTmp01 = Left(.Name, (Len(.Name) - 2))
28260             lngQryIdxs = 0&
28270             ReDim arr_varQryIdx(0)
28280             For Each qdf2 In dbs.QueryDefs
28290               With qdf2
28300                 If Left(.Name, Len(strTmp01)) = strTmp01 And .Name <> qdf1.Name Then
28310                   lngQryIdxs = lngQryIdxs + 1&
28320                   lngF = lngQryIdxs - 1&
28330                   ReDim Preserve arr_varQryIdx(lngF)
28340                   arr_varQryIdx(lngF) = .Name
28350                 End If
28360               End With
28370             Next
28380             arr_varQryDdl(QD_IDXS, lngE) = lngQryIdxs
28390             If lngQryIdxs > 0& Then
                    ' ** Sort the arr_varQryIdx() array.
28400               For lngX = UBound(arr_varQryIdx) To 1& Step -1
28410                 For lngY = 0 To (lngX - 1)
28420                   If arr_varQryIdx(lngY) > arr_varQryIdx((lngY + 1)) Then
28430                     strTmp01 = arr_varQryIdx(lngY)
28440                     arr_varQryIdx(lngY) = arr_varQryIdx((lngY + 1))
28450                     arr_varQryIdx((lngY + 1)) = strTmp01
28460                   End If
28470                 Next
28480               Next
28490               arr_varQryDdl(QD_IARR, lngE) = arr_varQryIdx
28500             End If
28510           End If
28520         End If
28530       End With
28540     Next
28550     .Close
28560   End With

28570   varRetVal = arr_varQryDdl

EXITP:
28580   Set qdf1 = Nothing
28590   Set qdf2 = Nothing
28600   Set dbs = Nothing
28610   Qry_LoadDDL = varRetVal
28620   Exit Function

ERRH:
28630   varRetVal = Empty
28640   Select Case ERR.Number
        Case Else
28650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28660   End Select
28670   Resume EXITP

End Function

Public Function Qry_RemExpr_rel(Optional varSearch As Variant, Optional varRepair As Variant, Optional varFind As Variant) As Boolean
' ** Removes the 'Expr1:' name assignments created when source table wasn't available.

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_RemExpr_rel"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strFind As String
        Dim blnSearch As Boolean, blnRepair As Boolean, blnSkip As Boolean, blnIsCalled As Boolean
        Dim lngExprsFound As Long
        Dim intPos01 As Integer, intPos02 As Integer, intLen As Integer
        Dim strTmp01 As String
        Dim intX As Integer
        Dim blnRetVal As Boolean

28710   blnRetVal = True
28720   blnIsCalled = False

        'X QRY: qryAbout_03
        'X QRY: qryAccountHide_04a
        'X QRY: qryAccountHide_23a
        'X QRY: qryAccountHide_23b
        'X QRY: qryAccountHide_24a
        'X QRY: qryReport_List_67
        'EXPRS FOUND: 6
        'NONE FOUND!

28730   strFind = "xxx"
28740   If IsMissing(varSearch) = True Then
28750     blnSearch = True  ' ** True: look for them; False: look for strFind.
28760   Else
28770     blnIsCalled = True
28780     blnSearch = varSearch
28790   End If
28800   If IsMissing(varRepair) = True Then
28810     blnRepair = False  ' ** True: fix them; False: only list them.
28820   Else
28830     blnIsCalled = True
28840     blnRepair = varRepair
28850   End If
28860   If IsMissing(varFind) = False Then
28870     blnIsCalled = True
28880     If IsNull(varFind) = False Then
28890       If Trim(varFind) <> vbNullString Then
28900         strFind = varFind
28910       End If
28920     End If
28930   End If
28940   lngExprsFound = 0&

28950   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

28960   Set dbs = CurrentDb
28970   With dbs
28980     For Each qdf In .QueryDefs
28990       With qdf
29000         If Left(.Name, 4) <> "~TMP" Then  ' ** Skip those pesky system queries.
29010           blnSkip = False
29020           If blnSearch = False Then
29030             If .Name <> strFind Then
29040               blnSkip = True
29050             End If
29060           End If
29070           If blnSkip = False Then
29080             strTmp01 = .SQL
29090             intPos01 = InStr(strTmp01, "As Expr")
29100             If intPos01 > 0 Then
                    ' ** May be 'As Expr1' or 'As Expr10'
29110               lngExprsFound = lngExprsFound + 1&
29120               If blnSearch = True And blnIsCalled = False Then
29130                 Debug.Print "'QRY: " & qdf.Name
                      'Stop
29140               End If
29150               If blnRepair = True Then
29160                 Do While intPos01 > 0
29170                   If Mid(strTmp01, (intPos01 - 1), 1) = " " Then
29180                     intPos01 = intPos01 - 1  ' ** Move intPos01 to the space before 'As'.
29190                     intPos02 = 0
29200                     intLen = Len(strTmp01)
29210                     For intX = (intPos01 + 1) To intLen
29220                       Select Case intX
                            Case (intPos01 + 1)
29230                         If Mid(strTmp01, intX, 1) <> "A" Then
29240                           blnRetVal = False
29250                           Stop
29260                         End If
29270                       Case (intPos01 + 2)
29280                         If Mid(strTmp01, intX, 1) <> "s" Then
29290                           blnRetVal = False
29300                           Stop
29310                         End If
29320                       Case (intPos01 + 3)
29330                         If Mid(strTmp01, intX, 1) <> " " Then
29340                           blnRetVal = False
29350                           Stop
29360                         End If
29370                       Case (intPos01 + 4)
29380                         If Mid(strTmp01, intX, 1) <> "E" Then
29390                           blnRetVal = False
29400                           Stop
29410                         End If
29420                       Case (intPos01 + 5)
29430                         If Mid(strTmp01, intX, 1) <> "x" Then
29440                           blnRetVal = False
29450                           Stop
29460                         End If
29470                       Case (intPos01 + 6)
29480                         If Mid(strTmp01, intX, 1) <> "p" Then
29490                           blnRetVal = False
29500                           Stop
29510                         End If
29520                       Case (intPos01 + 7)
29530                         If Mid(strTmp01, intX, 1) <> "r" Then
29540                           blnRetVal = False
29550                           Stop
29560                         End If
29570                       Case Else
29580                         If Asc(Mid(strTmp01, intX, 1)) >= 48 And Asc(Mid(strTmp01, intX, 1)) <= 57 Then
                                ' ** Numeral, keep checking.
29590                         Else
                                ' ** 'As Expr' phrase finished.
29600                           intPos02 = intX
29610                           Exit For
29620                         End If
29630                       End Select
29640                     Next
29650                     If intPos02 > 0 Then
29660                       strTmp01 = Left(strTmp01, (intPos01 - 1)) & Mid(strTmp01, intPos02)  ' ** intPos01 is on the space before 'As'.
29670                       intPos01 = InStr(strTmp01, "As Expr")  ' ** Continue checking for more 'Expr'.
29680                     Else
29690                       blnRetVal = False
29700                       If blnIsCalled = False Then
29710                         Debug.Print "'EXPR END NOT FOUND: " & strFind
29720                         Stop
29730                       End If
29740                       Exit Do
29750                     End If
29760                   Else
29770                     blnRetVal = False
29780                     If blnIsCalled = False Then
29790                       Debug.Print "'SPACE BEFORE 'AS' NOT FOUND: " & strFind
29800                       Stop
29810                     End If
29820                     Exit Do
29830                   End If
29840                 Loop
29850                 If blnRetVal = True Then
29860                   .SQL = strTmp01
29870                   dbs.QueryDefs.Refresh
29880                 End If
29890               End If  ' ** blnRepair.
29900             Else
29910               If blnSearch = False Then
29920                 blnRetVal = False
29930                 If blnIsCalled = False Then
29940                   Debug.Print "'EXPR NOT FOUND: " & strFind
29950                 End If
29960               End If
29970             End If  ' ** intPos01.
29980           End If  ' ** blnSkip.
29990         End If
30000       End With  ' ** qdf.
30010     Next
30020     .Close
30030   End With  ' ** dbs.

30040   If blnIsCalled = False Then
30050     If blnSearch = True Then
30060       If lngExprsFound > 0& Then
30070         Debug.Print "'EXPRS FOUND: " & CStr(lngExprsFound)
30080       Else
30090         Debug.Print "'NONE FOUND!"
30100       End If
30110     Else
30120       If lngExprsFound > 0& Then
30130         Debug.Print "'EXPRS FOUND: " & CStr(lngExprsFound)
30140       End If
30150     End If
30160   End If

30170   Beep

EXITP:
30180   Set qdf = Nothing
30190   Set dbs = Nothing
30200   Qry_RemExpr_rel = blnRetVal
30210   Exit Function

ERRH:
30220   blnRetVal = False
30230   Select Case ERR.Number
        Case Else
30240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30250   End Select
30260   Resume EXITP

End Function

Public Function Qry_Rename() As Boolean

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_Rename"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim strFind1 As String, strFind2 As String
        Dim strQryName As String, strSQL As String, strDesc1 As String, strDesc2 As String
        Dim lngQrysCreated As Long, lngQrysDeleted As Long
        Dim blnSkip As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 8  ' ** Array's first-element UBound().
        Const Q_QNAM1 As Integer = 0
        Const Q_SQL1  As Integer = 1
        Const Q_DSC1  As Integer = 2
        Const Q_TYP   As Integer = 3
        Const Q_QNAM2 As Integer = 4
        Const Q_SQL2  As Integer = 5
        Const Q_DSC2  As Integer = 6
        Const Q_COPY  As Integer = 7
        Const Q_DEL   As Integer = 8

30310 On Error GoTo 0

30320   blnRetVal = True

30330   strFind1 = "zzz_Report_List_"
30340   strFind2 = "zzz_qry_Report_List_"
30350   strDesc1 = vbNullString
30360   strDesc2 = vbNullString

30370   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
30380   DoEvents

30390   lngQrys = 0&
30400   ReDim arr_varQry(Q_ELEMS, 0)

30410   Set dbs = CurrentDb
30420   With dbs

30430     intLen = Len(strFind1)
30440     For Each qdf In .QueryDefs
30450       With qdf
30460         If Left(.Name, intLen) = strFind1 Then
30470           lngQrys = lngQrys + 1&
30480           lngE = lngQrys - 1&
30490           ReDim Preserve arr_varQry(Q_ELEMS, lngE)
30500           arr_varQry(Q_QNAM1, lngE) = .Name
30510           arr_varQry(Q_SQL1, lngE) = .SQL
30520 On Error Resume Next
30530           strTmp01 = .Properties("Description")
30540 On Error GoTo 0
30550           If strTmp01 <> vbNullString Then
30560             arr_varQry(Q_DSC1, lngE) = strTmp01
30570           Else
30580             arr_varQry(Q_DSC1, lngE) = Null
30590             Stop
30600           End If
30610           arr_varQry(Q_TYP, lngE) = .Type
30620           arr_varQry(Q_QNAM2, lngE) = Null
30630           arr_varQry(Q_SQL2, lngE) = Null
30640           arr_varQry(Q_DSC2, lngE) = Null
30650           arr_varQry(Q_COPY, lngE) = CBool(False)
30660           arr_varQry(Q_DEL, lngE) = CBool(False)
30670         End If
30680       End With  ' ** qdf.
30690     Next  ' ** qdf.

30700     Debug.Print "'QRYS: " & CStr(lngQrys)
30710     DoEvents

30720     If lngQrys > 0& Then

30730       For lngX = 0& To (lngQrys - 1&)
30740         strQryName = arr_varQry(Q_QNAM1, lngX)
30750         strQryName = StringReplace(strQryName, strFind1, strFind2)  ' ** Module Function: modStringFuncs.
30760         arr_varQry(Q_QNAM2, lngX) = strQryName
30770         strSQL = arr_varQry(Q_SQL1, lngX)
30780         strSQL = StringReplace(strSQL, strFind1, strFind2)  ' ** Module Function: modStringFuncs.
30790         arr_varQry(Q_SQL2, lngX) = strSQL
30800         If IsNull(arr_varQry(Q_DSC1, lngX)) = False Then
30810           strTmp01 = arr_varQry(Q_DSC1, lngX)
                'strTmp01 = StringReplace(strTmp01, strDesc1, strDesc2)  ' ** Module Function: modStringFuncs.
30820           arr_varQry(Q_DSC2, lngX) = strTmp01
30830         Else
30840           arr_varQry(Q_DSC2, lngX) = Null
30850         End If
30860       Next  ' ** lngX.

30870       blnSkip = False
30880       If blnSkip = False Then
30890         lngQrysCreated = 0&
30900         For lngX = 0& To (lngQrys - 1&)
30910           Set qdf = .CreateQueryDef(arr_varQry(Q_QNAM2, lngX), arr_varQry(Q_SQL2, lngX))
30920           With qdf
30930             If IsNull(arr_varQry(Q_DSC2, lngX)) = False Then
30940               strTmp01 = arr_varQry(Q_DSC2, lngX)
30950               Set prp = .CreateProperty("Description", dbText, strTmp01)
30960 On Error Resume Next
30970               .Properties.Append prp
30980               If ERR.Number <> 0 Then
30990 On Error GoTo 0
31000                 .Properties("Description") = strTmp01
31010               Else
31020 On Error GoTo 0
31030               End If
31040             End If
31050           End With  ' ** qdf.
31060           Set qdf = Nothing
31070           arr_varQry(Q_COPY, lngX) = CBool(True)
31080           lngQrysCreated = lngQrysCreated + 1&
31090         Next  ' ** lngX.
31100       End If  ' ** blnSkip.

31110       Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
31120       DoEvents

            'HASN'T THIS ALREADY BEEN DONE?
31130       blnSkip = True
31140       If blnSkip = False Then
31150         For lngX = 0& To (lngQrys - 1)
31160           If IsNull(arr_varQry(Q_DSC2, lngX)) = False Then
31170             Set qdf = .QueryDefs(arr_varQry(Q_QNAM2, lngX))
31180             With qdf
31190               Set prp = .CreateProperty("Description", dbText, arr_varQry(Q_DSC2, lngX))
31200 On Error Resume Next
31210               .Properties.Append prp
31220               If ERR.Number <> 0 Then
31230 On Error GoTo 0
31240                 .Properties("Description") = arr_varQry(Q_DSC2, lngX)
31250               Else
31260 On Error GoTo 0
31270               End If
31280             End With  ' ** qdf.
31290             Set qdf = Nothing
31300           Else
31310             Stop
31320           End If
31330         Next  ' ** lngX.
31340       End If  ' ** blnSkip.

            'THIS DOESN'T SEARCH FOR ANY QRYS OUTSIDE THE GROUP!
            'NEED TO RUN QRY_REF!
31350       For lngX = 0& To (lngQrys - 1&)
31360         Qry_UpdateRef_rel arr_varQry(Q_QNAM1, lngX), arr_varQry(Q_QNAM2, lngX)  ' ** Function: Below.
31370         DoEvents
31380       Next  ' ** lngX.

31390       blnSkip = False
31400       If blnSkip = False Then
31410         lngQrysDeleted = 0&
31420         For lngX = (lngQrys - 1&) To 0& Step -1&
31430           DoCmd.DeleteObject acQuery, arr_varQry(Q_QNAM1, lngX)
31440           DoEvents
31450           arr_varQry(Q_DEL, lngX) = CBool(True)
31460           lngQrysDeleted = lngQrysDeleted + 1&
31470         Next  ' ** lngX.
31480       End If  ' ** blnSkip.

31490       Debug.Print "'QRYS DELETED: " & CStr(lngQrysDeleted)
31500       DoEvents

31510     Else
31520       Debug.Print "'NONE FOUND!"
31530       DoEvents
31540     End If  ' ** lngQrys.

31550   End With  ' ** dbs.
31560   Set dbs = Nothing

31570   Beep

31580   Debug.Print "'DONE!"
31590   DoEvents

EXITP:
31600   Set prp = Nothing
31610   Set qdf = Nothing
31620   Set dbs = Nothing
31630   Qry_Rename = blnRetVal
31640   Exit Function

ERRH:
31650   blnRetVal = False
31660   Select Case ERR.Number
        Case Else
31670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31680   End Select
31690   Resume EXITP

End Function

Public Function Qry_Transfer_rel() As Boolean
' ** Copies a group of queries to another MDB.

31700 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_Transfer_rel"

        Dim dbs1 As DAO.Database, qdf As DAO.QueryDef
        Dim strMDB As String
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

31710   blnRetVal = True

31720   lngQrys = 0&
31730   ReDim arr_varQry(0)

31740   Set dbs1 = CurrentDb
31750   With dbs1
31760     For Each qdf In .QueryDefs
31770       With qdf
31780         If Left(.Name, 6) = "z_qry_" Or Left(.Name, 7) = "zz_qry_" Then
31790           lngQrys = lngQrys + 1&
31800           lngE = lngQrys - 1&
31810           ReDim Preserve arr_varQry(lngE)
31820           arr_varQry(lngE) = .Name
31830         End If
31840       End With
31850     Next
31860     .Close
31870   End With

31880   If lngQrys > 0& Then
31890     DoCmd.SetWarnings False  ' ** Copies over any pre-existing query with same name.
31900     strMDB = "C:\VictorGCS_Clients\TrustAccountant\Ver2-14-2 Working\TrustRegData.mdb"
31910     For lngX = 0& To 4& '(lngQrys - 1&)
31920       DoCmd.CopyObject strMDB, , acQuery, arr_varQry(lngX)
31930     Next
31940     DoCmd.SetWarnings True
31950   End If

31960   Debug.Print "'QUERIES COPIED: " & lngQrys

31970   Beep

EXITP:
31980   Set qdf = Nothing
31990   Set dbs1 = Nothing
32000   Qry_Transfer_rel = blnRetVal
32010   Exit Function

ERRH:
32020   blnRetVal = False
32030   DoCmd.SetWarnings True
32040   Select Case ERR.Number
        Case Else
32050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32060   End Select
32070   Resume EXITP

End Function

Public Function Qry_UpdateDesc_rel() As Boolean
' ** Within a query's Description, change all references from one source to another.

32100 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_UpdateDesc_rel"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object
        Dim strFind As String, strFind2 As String
        Dim strDesc As String
        Dim lngQrys As Long
        Dim intPos01 As Integer, intLen As Integer
        Dim blnRetVal As Boolean

32110   blnRetVal = True

32120   strFind = "zz_tbl_VBComponent_Procedure_01"
32130   strFind2 = "zz_tbl_VBComponent_Procedure"

32140   Set dbs = CurrentDb
32150   With dbs
32160     lngQrys = 0&
32170     For Each qdf In .QueryDefs
32180       With qdf
32190         For Each prp In .Properties
32200           If prp.Name = "Description" Then
32210             strDesc = .Properties("Description")  ' ** DON'T TRIM!! I DON'T WANT TO LOSE THE INDENT!
32220             If strDesc <> vbNullString Then
32230               intPos01 = InStr(strDesc, strFind)
32240               If intPos01 > 0 Then
32250                 lngQrys = lngQrys + 1&
32260                 Do While intPos01 > 0
32270                   intLen = Len(strDesc)
32280                   If intPos01 = 1 Then
32290                     strDesc = strFind2 & Mid(strDesc, (Len(strFind) + 1))
32300                   Else
32310                     If intPos01 + Len(strFind) > intLen Then
32320                       strDesc = Left(strDesc, (intPos01 - 1)) & strFind2
32330                     Else
32340                       strDesc = Left(strDesc, (intPos01 - 1)) & strFind2 & Mid(strDesc, (intPos01 + Len(strFind)))
32350                     End If
32360                   End If
32370                   intPos01 = InStr(strDesc, strFind)
32380                 Loop
32390                 .Properties("Description") = strDesc
32400               End If
32410             End If
32420             Exit For
32430           End If
32440         Next
32450       End With
32460     Next
32470     .Close
32480   End With

32490   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

32500   If lngQrys > 0& Then
32510     Debug.Print "'QRYS CHANGED: " & CStr(lngQrys)
32520   Else
32530     Debug.Print "'NONE FOUND!"
32540   End If

32550   Beep

EXITP:
32560   Set prp = Nothing
32570   Set qdf = Nothing
32580   Set dbs = Nothing
32590   Qry_UpdateDesc_rel = blnRetVal
32600   Exit Function

ERRH:
32610   blnRetVal = False
32620   Select Case ERR.Number
        Case Else
32630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32640   End Select
32650   Resume EXITP

End Function

Public Function Qry_UpdateRef_rel(Optional varOldRef As Variant, Optional varNewRef As Variant, Optional varCase As Variant, Optional varQry As Variant) As Boolean
' ** Change all references from one source to another.

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_UpdateRef_rel"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strFind As String, strFind2 As String
        Dim lngQrys As Long
        Dim strSQL As String, strQryName As String
        Dim blnCase As Boolean, blnContinue As Boolean, blnIsCalled As Boolean, blnSpecified As Boolean
        Dim intPos01 As Integer, intLen As Integer
        Dim intX As Integer
        Dim blnRetVal As Boolean

32710   blnRetVal = True

32720   Select Case IsMissing(varOldRef)
        Case True
32730     strFind = "chkvoid_bank"
32740     strFind2 = "chkbank_name"
          '"chkvoid_bank" -> "chkbank_name"
          '"chkvoid_bankacctnum" -> "chkbank_acctnum"
32750     blnCase = False  ' ** Match case.
32760     blnIsCalled = False
32770   Case False
32780     strFind = CStr(varOldRef)
32790     strFind2 = CStr(varNewRef)
32800     blnIsCalled = True
32810     Select Case IsMissing(varCase)
          Case True
32820       blnCase = False
32830     Case False
32840       blnCase = CBool(varCase)
32850     End Select
32860     Select Case IsMissing(varQry)
          Case True
32870       blnSpecified = False
32880       strQryName = vbNullString
32890     Case False
32900       Select Case IsNull(varQry)
            Case True
32910         blnSpecified = False
32920         strQryName = vbNullString
32930       Case False
32940         If Trim(varQry) = vbNullString Then
32950           blnSpecified = False
32960           strQryName = vbNullString
32970         Else
32980           blnSpecified = True
32990           strQryName = varQry
33000         End If
33010       End Select
33020     End Select
33030   End Select

33040   Set dbs = CurrentDb
33050   With dbs
33060     lngQrys = 0&
33070     For Each qdf In .QueryDefs
33080       With qdf
33090         blnContinue = True
33100         If blnSpecified = True Then
33110           If .Name <> strQryName Then
33120             blnContinue = False
33130           End If
33140         End If
33150         If blnContinue = True Then
33160           If Left(.Name, 1) <> "~" Then  ' ** Skip those pesky system queries.
33170             strSQL = .SQL
33180             If blnCase = False Then
33190               intPos01 = InStr(strSQL, strFind)
33200             Else
33210               intPos01 = InStr(strSQL, strFind)  'ERRORS WITH Type Mismatch! WHY? : , vbTextCompare)
33220             End If
33230             If intPos01 > 0 Then
33240               lngQrys = lngQrys + 1&
33250               Do While intPos01 > 0
33260                 strSQL = Left(strSQL, (intPos01 - 1)) & strFind2 & Mid(strSQL, (intPos01 + Len(strFind)))
33270                 intPos01 = InStr((intPos01 + 1), strSQL, strFind)
                      ' ** VbCompare enumeration.
                      ' **    0  vbBinaryCompare     Performs a binary comparison.
                      ' **    1  vbTextCompare       Performs a textual comparison.
                      ' **    2  vbDatabaseCompare   Microsoft Access only. Performs a comparison based on information in your database.
                      ' **    3  vbUseCompareOption  Performs a comparison using the setting of the Option Compare statement. (Stated value, -1, is wrong!)
33280                 If blnCase = True And intPos01 > 0 And strFind = strFind2 Then
33290                   intLen = Len(strFind): blnContinue = False
33300                   For intX = 1 To intLen
33310                     If Asc(Mid(strFind, intX, 1)) <> Asc(Mid(strSQL, ((intPos01 + intX) - 1), 1)) Then
                            ' ** If they're not equal, then continue the loop
33320                       blnContinue = True
33330                       Exit For
33340                     End If
33350                   Next
33360                   If blnContinue = False Then
                          ' ** All characters were identical, so this isn't a match.
33370                     intPos01 = 0
33380                   End If
33390                 End If
33400               Loop
33410               .SQL = strSQL
33420             End If
33430           End If
33440         End If  ' ** blnContinue.
33450       End With
33460     Next
33470     .Close
33480   End With

33490   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

33500   If blnIsCalled = False Then
33510     Debug.Print "'QRYS CHANGED: " & CStr(lngQrys)
33520   End If

33530   Beep

EXITP:
33540   Set qdf = Nothing
33550   Set dbs = Nothing
33560   Qry_UpdateRef_rel = blnRetVal
33570   Exit Function

ERRH:
33580   blnRetVal = False
33590   Select Case ERR.Number
        Case Else
33600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
33610   End Select
33620   Resume EXITP

End Function

Public Function Qry_ZZ_Tbl() As Boolean

33700 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_ZZ_Tbl"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim strSQL As String
        Dim blnFound As Boolean, lngNotFounds As Long
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const Q_QNAM As Integer = 0
        Const Q_SQL  As Integer = 1
        Const Q_TNAM As Integer = 2

        Const TBL_BASE As String = "zz_tbl_"

33710 On Error GoTo 0

33720   blnRetVal = True

33730   lngQrys = 0&
33740   ReDim arr_varQry(Q_ELEMS, 0)

33750   Set dbs = CurrentDb
33760   With dbs
33770     For Each qdf In .QueryDefs
33780       With qdf
33790         strSQL = .SQL
33800         intPos01 = InStr(strSQL, TBL_BASE)
33810         If intPos01 > 0 Then
33820           strTmp01 = Mid(strSQL, intPos01)
33830           intPos01 = InStr(strTmp01, " ")
33840           If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, intPos01))
33850           intPos01 = InStr(strTmp01, ".")
33860           If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
33870           intPos01 = InStr(strTmp01, ",")
33880           If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
33890           intPos01 = InStr(strTmp01, ";")
33900           If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
33910           intPos01 = InStr(strTmp01, "]")
33920           If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
33930           intPos01 = InStr(strTmp01, ")")
33940           If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
33950           intPos01 = InStr(strTmp01, "'")
33960           If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
33970           strTmp01 = Rem_CRLF(strTmp01)  ' ** Module Function: modStringFuncs.
33980           blnFound = False
33990           For lngX = 0& To (lngQrys - 1&)
34000             If arr_varQry(Q_TNAM, lngX) = strTmp01 Then
34010               blnFound = True
34020               Exit For
34030             End If
34040           Next
34050           If blnFound = False Then
34060             lngQrys = lngQrys + 1&
34070             lngE = lngQrys - 1&
34080             ReDim Preserve arr_varQry(Q_ELEMS, lngE)
34090             arr_varQry(Q_QNAM, lngE) = .Name
34100             arr_varQry(Q_SQL, lngE) = .SQL
34110             arr_varQry(Q_TNAM, lngE) = strTmp01
34120           End If
34130         End If
34140       End With
34150     Next
34160     .Close
34170   End With

34180   Debug.Print "'ZZ_TBLS: " & CStr(lngQrys)
34190   DoEvents

34200   If lngQrys > 0& Then
34210     lngNotFounds = 0&
34220     For lngX = 0& To (lngQrys - 1&)
34230       If Right(arr_varQry(Q_TNAM, lngX), 1) <> "_" Then  ' ** Something about 'zz_tbl_RePost_'.
34240         If TableExists(CStr(arr_varQry(Q_TNAM, lngX))) = False Then  ' ** Module Function: modFileUtilities.
34250           Debug.Print "'TBL NOT FOUND: " & arr_varQry(Q_TNAM, lngX)
34260           DoEvents
34270           lngNotFounds = lngNotFounds + 1&
34280         End If
34290       End If
34300     Next
34310   End If

34320   If lngNotFounds > 0& Then
34330     Debug.Print "'TBLS NOT FOUND: " & CStr(lngNotFounds)
34340   Else
34350     Debug.Print "'ALL FOUND!"
34360   End If
34370   DoEvents

34380   Beep
34390   Debug.Print "'DONE!"

EXITP:
34400   Set qdf = Nothing
34410   Set dbs = Nothing
34420   Qry_ZZ_Tbl = blnRetVal
34430   Exit Function

ERRH:
34440   blnRetVal = False
34450   Select Case ERR.Number
        Case Else
34460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34470   End Select
34480   Resume EXITP

End Function

Public Function Qry_Doc_Simple() As Boolean

34500 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_Doc_Simple"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim lngDels As Long, arr_varDel() As Variant
        Dim lngUpdates As Long, lngNews As Long
        Dim lngThisDbsID As Long, lngRecs As Long
        Dim blnFound As Boolean
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 6  ' ** Array's first-element UBound().
        Const Q_QID  As Integer = 0
        Const Q_QNAM As Integer = 1
        Const Q_TYP  As Integer = 2
        Const Q_FLDS As Integer = 3
        Const Q_DSC  As Integer = 4
        Const Q_UPD  As Integer = 5
        Const Q_FND  As Integer = 6

        ' ** Array: arr_varDel().
        Const D_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const D_QID  As Integer = 0
        Const D_QNAM As Integer = 1

34510 On Error GoTo 0

34520   blnRetVal = True

34530   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
34540   DoEvents

34550   lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

34560   lngQrys = 0&
34570   ReDim arr_varQry(Q_ELEMS, 0)

34580   Set dbs = CurrentDb
34590   With dbs

34600     Debug.Print "'ASSEMBLING LIST:"
34610     Debug.Print "'|";
34620     lngRecs = .QueryDefs.Count
34630     lngX = 0&
34640     For Each qdf In .QueryDefs
34650       With qdf
34660         strTmp01 = vbNullString
34670         lngX = lngX + 1&
34680         If Left(.Name, 4) <> "~TMP" And Left(.Name, 3) <> "~sq" Then  ' ** Skip those pesky system queries!
34690           lngQrys = lngQrys + 1&
34700           lngE = lngQrys - 1&
34710           ReDim Preserve arr_varQry(Q_ELEMS, lngE)
34720           arr_varQry(Q_QID, lngE) = Null
34730           arr_varQry(Q_QNAM, lngE) = .Name
34740           arr_varQry(Q_TYP, lngE) = .Type
34750           arr_varQry(Q_FLDS, lngE) = .Fields.Count
34760 On Error Resume Next
34770           strTmp01 = .Properties("Description")
34780 On Error GoTo 0
34790           If strTmp01 <> vbNullString Then
34800             arr_varQry(Q_DSC, lngE) = strTmp01
34810           Else
34820             arr_varQry(Q_DSC, lngE) = Null
34830           End If
34840           arr_varQry(Q_UPD, lngE) = CBool(False)
34850           arr_varQry(Q_FND, lngE) = CBool(False)
34860         End If
34870       End With  ' ** qdf.
34880       If (lngX + 1) Mod 1000 = 0 Then
34890         Debug.Print "|  " & CStr(lngX + 1&) & " OF " & CStr(lngRecs)
34900         Debug.Print "'|";
34910       ElseIf (lngX + 1) Mod 100 = 0 Then
34920         Debug.Print "|";
34930       ElseIf (lngX + 1) Mod 10 = 0 Then
34940         Debug.Print ".";
34950       End If
34960       DoEvents
34970     Next  ' ** qdf.
34980     Set qdf = Nothing
34990     Debug.Print
35000     DoEvents

35010     Debug.Print "'QRYS: " & CStr(lngQrys)
35020     DoEvents

35030     If lngQrys > 0& Then

35040       lngDels = 0&
35050       ReDim arr_varDel(D_ELEMS, 0)

35060       Debug.Print "'CHECKING TABLE:"
35070       DoEvents

35080       Set rst = .OpenRecordset("tblQuery", dbOpenDynaset, dbConsistent)
35090       With rst
35100         .MoveLast
35110         lngRecs = .RecordCount
35120         .MoveFirst
35130         Debug.Print "'|";
35140         For lngX = 1& To lngRecs
35150           If ![dbs_id] = 1& Then
35160             blnFound = False
35170             For lngY = 0& To (lngQrys - 1&)
35180               If arr_varQry(Q_QNAM, lngY) = ![qry_name] Then
35190                 blnFound = True
35200                 If arr_varQry(Q_TYP, lngY) <> ![qrytype_type] Then
35210                   arr_varQry(Q_UPD, lngY) = CBool(True)
35220                 End If
35230                 If IsNull(![qry_description]) = True And IsNull(arr_varQry(Q_DSC, lngY)) = True Then
                        ' ** Fine.
35240                 Else
35250                   If IsNull(![qry_description]) = True Or IsNull(arr_varQry(Q_DSC, lngY)) = True Then
35260                     arr_varQry(Q_UPD, lngY) = CBool(True)
35270                   Else
35280                     If ![qry_description] <> arr_varQry(Q_DSC, lngY) Then
35290                       arr_varQry(Q_UPD, lngY) = CBool(True)
35300                     End If
35310                   End If
35320                 End If
35330                 If ![qry_fldcnt] <> arr_varQry(Q_FLDS, lngY) Then
35340                   arr_varQry(Q_UPD, lngY) = CBool(True)
35350                 End If
35360                 arr_varQry(Q_FND, lngY) = CBool(True)
35370                 Exit For
35380               End If
35390             Next  ' ** lngY.
35400             If blnFound = False Then
35410               lngDels = lngDels + 1&
35420               lngE = lngDels - 1&
35430               ReDim Preserve arr_varDel(D_ELEMS, lngE)
35440               arr_varDel(D_QID, lngE) = ![qry_id]
35450               arr_varDel(D_QNAM, lngE) = ![qry_name]
35460             End If  ' ** blnFound.
35470           End If  ' ** dbs_id.
35480           If (lngX + 1) Mod 1000 = 0 Then
35490             Debug.Print "|  " & CStr(lngX + 1&) & " OF " & CStr(lngRecs)
35500             Debug.Print "'|";
35510           ElseIf (lngX + 1) Mod 100 = 0 Then
35520             Debug.Print "|";
35530           ElseIf (lngX + 1) Mod 10 = 0 Then
35540             Debug.Print ".";
35550           End If
35560           DoEvents
35570           If lngX < lngRecs Then .MoveNext
35580         Next  ' ** lngX
35590         .Close
35600       End With  ' ** rst.
35610       Set rst = Nothing
35620       Debug.Print
35630       DoEvents

35640       lngNews = 0&: lngUpdates = 0&
35650       For lngX = 0& To (lngQrys - 1&)
35660         If arr_varQry(Q_FND, lngX) = False Then
35670           lngNews = lngNews + 1&
35680         ElseIf arr_varQry(Q_UPD, lngX) = True Then
35690           lngUpdates = lngUpdates + 1&
35700         End If
35710       Next  ' ** lngX.

35720       Debug.Print "'NEW QRYS: " & CStr(lngNews)
35730       DoEvents
35740       Debug.Print "'DELS:     " & CStr(lngDels)
35750       DoEvents
35760       Debug.Print "'UPDATES:  " & CStr(lngUpdates)
35770       DoEvents

35780       If lngDels > 0& Then
35790         Debug.Print "'DELETING OBSOLETE QRYS:"
35800         Debug.Print "'|";
35810         DoEvents
35820         For lngX = 0& To (lngDels - 1&)
                ' ** Delete tblQuery, by specified [qid].
35830           Set qdf = .QueryDefs("zz_qry_Query_01b")
35840           With qdf.Parameters
35850             ![qid] = arr_varDel(D_QID, lngX)
35860           End With
35870           qdf.Execute
35880           If (lngX + 1&) Mod 100 = 0 Then
35890             Debug.Print "|  " & CStr(lngX + 1&) & " OF " & CStr(lngDels)
35900             Debug.Print "'|";
35910           ElseIf (lngX + 1&) Mod 10 = 0 Then
35920             Debug.Print "|";
35930           Else
35940             Debug.Print ".";
35950           End If
35960           DoEvents
35970         Next  ' ** lngX.
35980         Debug.Print
35990         DoEvents
36000         .QueryDefs.Refresh
36010       End If  ' ** lngDels.
36020       DoEvents

36030       If lngNews > 0& Then
36040         Debug.Print "'ADDING NEW QRYS:"
36050         Debug.Print "'|";
36060         DoEvents
36070         lngY = 0&
36080         Set rst = .OpenRecordset("tblQuery", dbOpenDynaset, dbConsistent)
36090         For lngX = 0& To (lngQrys - 1&)
36100           If arr_varQry(Q_FND, lngX) = False Then
36110             lngY = lngY + 1&
36120             Set qdf = .QueryDefs(arr_varQry(Q_QNAM, lngX))
36130             With rst
36140               .AddNew
36150               ![dbs_id] = lngThisDbsID
                    ' ** ![qry_id] : AutoNumber.
36160               ![qry_name] = arr_varQry(Q_QNAM, lngX)
36170               ![qrytype_type] = arr_varQry(Q_TYP, lngX)
36180               If IsNull(arr_varQry(Q_DSC, lngX)) = False Then
36190                 ![qry_description] = arr_varQry(Q_DSC, lngX)
36200               Else
36210                 ![qry_description] = Null
36220               End If
36230               ![qry_sql] = qdf.SQL
36240               ![qry_param] = False  ' ** Not going to check for others.
36250               If Left(qdf.SQL, 10) = "PARAMETERS" Then
36260                 ![qry_param_clause] = True
36270               Else
36280                 ![qry_param_clause] = False
36290               End If
36300               ![qry_formref] = False
36310               ![sec_hidden] = False  ' ** Not going to check
36320               ![qry_tblcnt] = 0  ' ** Not going to check
36330               ![qry_fldcnt] = arr_varQry(Q_FLDS, lngX)
36340               ![qry_paramcnt] = 0  ' ** Not going to check
36350               ![qry_formrefcnt] = 0
36360               ![qry_datemodified] = Now()
36370               .Update
36380             End With  ' ** rst.
36390             If lngY Mod 100 = 0 Then
36400               Debug.Print "|  " & CStr(lngY) & " OF " & CStr(lngNews)
36410               Debug.Print "'|";
36420             ElseIf lngY Mod 10 = 0 Then
36430               Debug.Print "|";
36440             Else
36450               Debug.Print ".";
36460             End If
36470             DoEvents
36480           End If
36490         Next  ' ** lngX.
36500         rst.Close
36510         Set rst = Nothing
36520         Set qdf = Nothing
36530         Debug.Print
36540         DoEvents
36550       End If  ' ** lngNews.

36560       If lngUpdates > 0& Then
36570         Debug.Print "'UPDATING EXISTING QRYS:"
36580         Debug.Print "'|";
36590         DoEvents
36600         lngY = 0&
36610         Set rst = .OpenRecordset("tblQuery", dbOpenDynaset, dbConsistent)
36620         For lngX = 0& To (lngQrys - 1&)
36630           If arr_varQry(Q_UPD, lngX) = True Then
36640             lngY = lngY + 1&
36650             Set qdf = .QueryDefs(arr_varQry(Q_QNAM, lngX))
36660             With rst
36670               .MoveFirst
36680               .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [qry_name] = '" & arr_varQry(Q_QNAM, lngX) & "'"
36690               If .NoMatch = False Then
36700                 .Edit
36710                 ![qrytype_type] = arr_varQry(Q_TYP, lngX)
36720                 If IsNull(arr_varQry(Q_DSC, lngX)) = False Then
36730                   ![qry_description] = arr_varQry(Q_DSC, lngX)
36740                 Else
36750                   ![qry_description] = Null
36760                 End If
36770                 ![qry_sql] = qdf.SQL
36780                 If Left(qdf.SQL, 10) = "PARAMETERS" Then
36790                   ![qry_param_clause] = True
36800                 Else
36810                   ![qry_param_clause] = False
36820                 End If
36830                 ![qry_fldcnt] = arr_varQry(Q_FLDS, lngX)
36840                 ![qry_datemodified] = Now()
36850                 .Update
36860               Else
36870                 Debug.Print "'QRY NOT FOUND!  " & arr_varQry(Q_QNAM, lngX)
36880                 DoEvents
36890               End If
36900             End With  ' ** rst.
36910             If lngY Mod 100 = 0 Then
36920               Debug.Print "|  " & CStr(lngY) & " OF " & CStr(lngUpdates)
36930               Debug.Print "'|";
36940             ElseIf lngY Mod 10 = 0 Then
36950               Debug.Print "|";
36960             Else
36970               Debug.Print ".";
36980             End If
36990             DoEvents
37000           End If
37010         Next  ' ** lngX
37020         rst.Close
37030         Set rst = Nothing
37040         Set qdf = Nothing
37050         Debug.Print
37060         DoEvents
37070       End If  ' ** lngUpdates.

37080     End If  ' ** lngQrys.

37090     .Close
37100   End With  ' ** dbs.
37110   Set dbs = Nothing

37120   Beep

37130   Debug.Print "'DONE!"
37140   DoEvents

EXITP:
37150   Set rst = Nothing
37160   Set qdf = Nothing
37170   Set dbs = Nothing
37180   Qry_Doc_Simple = blnRetVal
37190   Exit Function

ERRH:
37200   blnRetVal = False
37210   Select Case ERR.Number
        Case Else
37220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
37230   End Select
37240   Resume EXITP

End Function

Public Function Qry_PropDoc() As Boolean

37300 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_PropDoc"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object, rst As DAO.Recordset, cntr As DAO.Container, doc As DAO.Document
        Dim lngProps As Long, arr_varProp() As Variant
        Dim lngThisDbsID As Long, lngQryType As Long
        Dim blnFound As Boolean, blnAddAll As Boolean, blnAdd As Boolean
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varProp().
        Const P_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const P_PNAM As Integer = 0
        Const P_PTYP As Integer = 1
        Const P_QTYP As Integer = 2

37310 On Error GoTo 0

37320   blnRetVal = True

37330   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
37340   DoEvents

37350   lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

37360   lngProps = 0&
37370   ReDim arr_varProp(P_ELEMS, 0)

37380   Set dbs = CurrentDb
37390   With dbs

37400     For Each qdf In .QueryDefs
37410       With qdf
37420         lngQryType = .Type
37430         For Each prp In .Properties
37440           With prp
37450             blnFound = False
37460             For lngY = 0& To (lngProps - 1&)
37470               If arr_varProp(P_PNAM, lngY) = .Name And arr_varProp(P_QTYP, lngY) = lngQryType Then
37480                 blnFound = True
37490                 Exit For
37500               End If
37510             Next  ' ** lngY.
37520             If blnFound = False Then
37530               lngProps = lngProps + 1&
37540               lngE = lngProps - 1&
37550               ReDim Preserve arr_varProp(P_ELEMS, lngE)
37560               arr_varProp(P_PNAM, lngE) = .Name
37570               arr_varProp(P_PTYP, lngE) = .Type
37580               arr_varProp(P_QTYP, lngE) = lngQryType
37590             End If
37600           End With  ' ** prp.
37610         Next  ' ** prp.
37620       End With  ' ** qdf.
37630     Next  ' ** qdf.
37640     Set prp = Nothing
37650     Set qdf = Nothing

37660     Set cntr = .Containers("Tables")
37670     With cntr
37680       For Each doc In .Documents
37690         With doc
37700           If Left(.Name, 3) = "qry" Or Left(.Name, 6) = "zz_qry" Or Left(.Name, 7) = "zzz_qry" Then
37710             lngQryType = dbs.QueryDefs(.Name).Type
37720             For Each prp In .Properties
37730               With prp
37740                 blnFound = False
37750                 For lngY = 0& To (lngProps - 1&)
37760                   If arr_varProp(P_PNAM, lngY) = .Name And arr_varProp(P_QTYP, lngY) = lngQryType Then
37770                     blnFound = True
37780                     Exit For
37790                   End If
37800                 Next  ' ** lngY.
37810                 If blnFound = False Then
37820                   lngProps = lngProps + 1&
37830                   lngE = lngProps - 1&
37840                   ReDim Preserve arr_varProp(P_ELEMS, lngE)
37850                   arr_varProp(P_PNAM, lngE) = .Name
37860                   arr_varProp(P_PTYP, lngE) = .Type
37870                   arr_varProp(P_QTYP, lngE) = lngQryType
37880                 End If
37890               End With  ' ** prp.
37900             Next  ' ** prp.
37910           End If
37920         End With  ' ** doc.
37930       Next  ' ** doc.
37940       Set prp = Nothing
37950       Set doc = Nothing
37960     End With  ' ** cntr.
37970     Set cntr = Nothing

37980     Debug.Print "'PROPS: " & CStr(lngProps)
37990     DoEvents

38000     If lngProps > 0& Then
38010       Set rst = .OpenRecordset("tblQuery_Properties", dbOpenDynaset, dbConsistent)
38020       With rst
38030         blnAddAll = False: blnAdd = False
38040         If .BOF = True And .EOF = True Then
38050           blnAddAll = True
38060         End If
38070         For lngX = 0& To (lngProps - 1&)
38080           blnAdd = False
38090           If blnAddAll = True Then
38100             blnAdd = True
38110           Else
38120             .MoveFirst
38130             If ![qrytype_type] = CStr(arr_varProp(P_QTYP, lngX)) And ![qryprop_name] = arr_varProp(P_PNAM, lngX) Then
                    ' ** It was the first record, so leave blnAdd = False.
38140             Else
38150               .FindFirst "[qrytype_type] = " & CStr(arr_varProp(P_QTYP, lngX)) & " And [qryprop_name] = '" & arr_varProp(P_PNAM, lngX) & "'"
38160               If .NoMatch = True Then
38170                 blnAdd = True
38180               End If
38190             End If
38200           End If
38210           If blnAdd = True Then
38220             .AddNew
                  ' ** ![qryprop_id] : AutoNumber.
38230             ![qrytype_type] = arr_varProp(P_QTYP, lngX)
38240             ![qryprop_name] = arr_varProp(P_PNAM, lngX)
38250             ![datatype_db_type] = arr_varProp(P_PTYP, lngX)
38260             ![qryprop_datemodified] = Now()
38270             .Update
38280           End If  ' ** blnAdd.
38290         Next  ' ** lngX.
38300       End With  ' ** rst.
38310       Set rst = Nothing
38320     End If  ' ** lngProps.

38330     .Close
38340   End With  ' ** dbs.
38350   Set dbs = Nothing

38360   Beep

38370   Debug.Print "'DONE!"
38380   DoEvents

EXITP:
38390   Set rst = Nothing
38400   Set doc = Nothing
38410   Set prp = Nothing
38420   Set qdf = Nothing
38430   Set cntr = Nothing
38440   Set dbs = Nothing
38450   Qry_PropDoc = blnRetVal
38460   Exit Function

ERRH:
38470   blnRetVal = False
38480   Select Case ERR.Number
        Case Else
38490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
38500   End Select
38510   Resume EXITP

End Function
