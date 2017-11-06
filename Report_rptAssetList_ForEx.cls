VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAssetList_ForEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAssetList_ForEx"

'VGC 11/27/2016: CHANGES!

' ** Possible RecordSources:
' **   qryStatementParameters_AssetList_06a -> qryStatementParameters_AssetList_70_01
'PRICING HISTORY!
'qryStatementParameters_AssetList_70_21
' **   qryStatementParameters_AssetList_06b -> qryStatementParameters_AssetList_70_02  ***
'PRICING HISTORY!
'qryStatementParameters_AssetList_70_22
' **   qryStatementParameters_AssetList_06c -> qryStatementParameters_AssetList_70_03
'PRICING HISTORY!
'qryStatementParameters_AssetList_70_23
' **   qryStatementParameters_AssetList_06e -> qryStatementParameters_AssetList_70_04
'PRICING HISTORY!
'qryStatementParameters_AssetList_70_24
' **   qryStatementParameters_AssetList_06f -> qryStatementParameters_AssetList_70_05
'PRICING HISTORY!
'qryStatementParameters_AssetList_70_25
' **   qryStatementParameters_AssetList_06g -> qryStatementParameters_AssetList_70_06
'PRICING HISTORY!
'qryStatementParameters_AssetList_70_26
' **   qryStatementParameters_AssetList_15  -> qryStatementParameters_AssetList_70_08
'PRICING HISTORY!
'qryStatementParameters_AssetList_70_28
' **   qryStatementParameters_AssetList_15b -> qryStatementParameters_AssetList_70_09
'PRICING HISTORY!
'qryStatementParameters_AssetList_70_29
' **   qryStatementParameters_AssetList_28  -> qryStatementParameters_AssetList_70_13
'PRICING HISTORY!
'qryStatementParameters_AssetList_70_33
' **   qryStatementParameters_AssetList_28a -> qryStatementParameters_AssetList_70_14
'PRICING HISTORY!
'qryStatementParameters_AssetList_70_34

'MarketValueCurrentX IS CURRENT LOCAL VALUE!
'IF THEY ROLLBACK, DO WE WANT TO RETRIEVE PREVIOUS VALUE?
'COST USD AND MARKET USD ARE AS OF PURCHASE DATE, AS IS COST,
'BUT MARKET COLUMN (NOT USD) IS TODAY'S MARKET VALUE IN LOCAL CURRENCY!
'SO WE'LL HAVE TO RETRIEVE THE PRICE AS OF THE REPORT DATE IF IT ROLLS BACK!

Private Const strCallingForm As String = "frmStatementParameters"

' ** Array: arr_varFld().
Private lngFlds As Long, arr_varFld() As Variant
Private Const F_ELEMS As Integer = 9  ' ** Array's first-element UBound().
Private Const F_ACTNO As Integer = 0
Private Const F_DATE  As Integer = 1
Private Const F_FOREX As Integer = 2
Private Const F_CST   As Integer = 3
Private Const F_MKT   As Integer = 4
Private Const F_ICSH  As Integer = 5
Private Const F_PCSH  As Integer = 6
Private Const F_TCST  As Integer = 7
Private Const F_TMKT  As Integer = 8
Private Const F_RECS  As Integer = 9

Private strAccountNo As String
Private datStartDate As Date, datEndDate As Date
Private blnPriceHistory As Boolean, blnStatements As Boolean, blnSaveData As Boolean, blnNoAssets As Boolean, blnNoAssets_All As Boolean
Private frm As Access.Form
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strFilter As String
        Dim blnSkip As Boolean
        Dim varTmp00 As Variant
        Dim lngZ As Long

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
            ' ** Data taken directly from form.
            ' ** .AssetListDate
            ' ** .DateEnd
            ' ** .chkStatements
            ' ** .shortname
            ' ** .legalname

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
140         blnSkip = False: blnSaveData = False
150         blnPriceHistory = False

160         Set frm = Forms(strCallingForm)

170         If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
180           If frm.PopUp = True Then
190             frm.Visible = False
200           End If
210         End If

220         Select Case frm.chkPageOf
            Case True
230           .PageNumOf.Visible = True
240           .PageNum.Visible = False
250         Case False
260           .PageNum.Visible = True
270           .PageNumOf.Visible = False
280           .PageNumOf.ControlSource = vbNullString
290         End Select

300         .Name_lbl.Caption = vbNullString
310         .group_footer1_height_line.Top = 0&  ' ** Just get it out of the way.

320         varTmp00 = frm.cmbAccounts
330         If IsNull(varTmp00) = False Then
340           varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
350           varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
360           .Caption = "Asset_List_" & varTmp00 & "_" & Format(CDate(frm.AssetListDate), "yymmdd")
370           strAccountNo = varTmp00
380         Else
390           .Caption = "Asset_List_All_" & Format(CDate(frm.AssetListDate), "yymmdd")
400         End If

410         VerifyDates  ' ** Procedure: Below.

420         datStartDate = frm.DateStart
430         datEndDate = frm.AssetListDate
440         blnNoAssets = frm.chkNoAssets
450         blnNoAssets_All = frm.chkNoAssets_All

            ' ** Change accountno grouping to KeepTogether, which prevents blank page at end.
            'Select Case frm.opgAccountNumber
            'Case frm.opgAccountNumber_optSpecified.OptionValue
            '  .Section("GroupFooter1").ForceNewPage = acForceNewPageNone
            'Case frm.opgAccountNumber_optAll.OptionValue
            '  .Section("GroupFooter1").ForceNewPage = acForceNewPageAfter
            'End Select

460         blnPriceHistory = frm.UsePriceHistory

470         Select Case blnPriceHistory
            Case True
              ' ** Detail:
480           .PercentMarket.Visible = False
490           .AnnualIncome.Visible = False
500           .PercentMarket_usd.Visible = True
510           .AnnualIncome_usd.Visible = True
              ' ** GroupFooter2:
520           .TotalPctMarket_Asset_sum.Visible = False
530           .TotalPctMarket_Asset_sum_usd.Visible = True
              ' ** GroupFooter1:
540           .TotalPctMarket_sum.Visible = False
550           .TotalPctMarket_Grand_sum.Visible = False
560           .marketvaluecurrent.Visible = False
570           .TotalPctMarket_sum_usd.Visible = True
580           .TotalPctMarket_Grand_sum_usd.Visible = True
590           .MarketValueCurrent_usd.Visible = True
600         Case False
              ' ** Detail:
610           .PercentMarket_usd.Visible = False
620           .AnnualIncome_usd.Visible = False
630           .PercentMarket.Visible = True
640           .AnnualIncome.Visible = True
              ' ** GroupFooter2:
650           .TotalPctMarket_Asset_sum_usd.Visible = False
660           .TotalPctMarket_Asset_sum.Visible = True
              ' ** GroupFooter1:
670           .TotalPctMarket_sum_usd.Visible = False
680           .TotalPctMarket_Grand_sum_usd.Visible = False
690           .MarketValueCurrent_usd.Visible = False
700           .TotalPctMarket_sum.Visible = True
710           .TotalPctMarket_Grand_sum.Visible = True
720           .marketvaluecurrent.Visible = True
730         End Select

740         Select Case frm.chkStatements
            Case True
750           blnStatements = True
760         Case False
770           blnStatements = False
780         End Select

790         Select Case gblnMessage
            Case True
              ' ** Rollbacks were needed.
800           Select Case frm.chkStatements
              Case True
810             Select Case blnNoAssets
                Case True
820               Select Case blnPriceHistory
                  Case True
                    ' ** qryStatementParameters_AssetList_74_40 (xx), linked to Account, for 'NoAssets', ForEx.
830                 .RecordSource = "qryStatementParameters_AssetList_70_36_01"
840               Case False
                    ' ** I don't think these no price history will ever hit!
                    ' ** qryStatementParameters_AssetList_74_41 (xx), linked to Account, for 'NoAssets'.
850                 .RecordSource = "qryStatementParameters_AssetList_70_36_02"
860               End Select
870             Case False
880               Select Case blnNoAssets_All
                  Case True
                    ' ** Union of qryStatementParameters_AssetList_70_29 (qryStatementParameters_AssetList_70_29_01 (xx),
                    ' ** linked to qryStatementParameters_AssetList_70_29_02 (xx)), qryStatementParameters_AssetList_70_29_03
                    ' ** (qryStatementParameters_AssetList_74_65 (xx), linked to Account, with add'l fields).
890                 .RecordSource = "qryStatementParameters_AssetList_70_29_04"
900               Case False
910                 Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** qryStatementParameters_AssetList_70_28 (xx), linked to qryStatementParameters_AssetList_19b
                      ' ** (xx), just scheduled accounts; From .._15b.
920                   .RecordSource = "qryStatementParameters_AssetList_70_29"
                      '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_29_02"
930                 Case False
                      ' ** I don't think these no price history will ever hit!
                      ' ** qryStatementParameters_AssetList_70_08 (qryStatementParameters_AssetList_70_07
                      ' ** (tmpAssetList5, all fields, with rollback, with TotalCost_usd, TotalMarket_usd),
                      ' ** with Liability adjustment), linked to qryStatementParameters_AssetList_19b
                      ' ** (qryStatementParameters_AssetList_19a (Account, with MonthNum = FormRef('MonthNum')),
                      ' ** just those matching MonthNum), just scheduled accounts; From .._15b.
940                   .RecordSource = "qryStatementParameters_AssetList_70_09"
                      '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_09_02"
950                 End Select
960               End Select
970             End Select
980           Case False
                ' ** chkAssetList.
                'NO ASSETS!
990             Select Case blnPriceHistory
                Case True
                  ' ** PRICING HISTORY!
                  ' ** qryStatementParameters_AssetList_70_27 (xx), with Liability adjustment; From .._15.
1000              .RecordSource = "qryStatementParameters_AssetList_70_28"
                  '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_28_02"
1010            Case False
                  ' ** I don't think these no price history will ever hit!
                  ' ** qryStatementParameters_AssetList_70_07 (tmpAssetList5, all fields, with rollback,
                  ' ** with TotalCost_usd, TotalMarket_usd), with Liability adjustment; From .._15.
1020              .RecordSource = "qryStatementParameters_AssetList_70_08"
                  '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_08_02"
1030            End Select
1040          End Select
1050          If frm.opgAccountNumber = frm.opgAccountNumber_optSpecified.OptionValue Then
1060            strFilter = "[accountno] = '" & frm.cmbAccounts & "'"
1070            .Filter = strFilter
1080            .FilterOn = True
1090          End If
1100        Case False
              ' ** No Rollbacks needed.
1110          Select Case frm.chkRelatedAccounts
              Case True
                ' ** With Related Accounts.
1120            Select Case frm.chkStatements
                Case True
                  'NO ASSETS!
1130              Select Case blnPriceHistory
                  Case True
                    ' ** PRICING HISTORY!
                    ' ** qryStatementParameters_AssetList_70_33 (xx), linked to qryStatementParameters_AssetList_19b
                    ' ** (xx), just scheduled accounts; From .._28a.
1140                .RecordSource = "qryStatementParameters_AssetList_70_34"
                    '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_34_02"
1150              Case False
                    ' ** qryStatementParameters_AssetList_70_13 (qryStatementParameters_AssetList_70_12
                    ' ** (tmpRelatedAccount_03, with qryStatementParameters_AssetList_70_11
                    ' ** (qryStatementParameters_AssetList_70_10 (tmpRelatedAccount_03, with Foreign Exchange,
                    ' ** linked to Account, grouped and summed, by accountno), grouped and summed), by specified
                    ' ** [ractnos], with TotalCost_usd, TotalMarket_usd; Cartesian), with qryStatementParameters_AssetList_70_11
                    ' ** (qryStatementParameters_AssetList_70_10 (tmpRelatedAccount_03, with Foreign Exchange, linked
                    ' ** to Account, grouped and summed, by accountno), grouped and summed), from code, with [ractnos]
                    ' ** replaced with actual accountno's, with TotalCost_usd, TotalMarket_usd; Cartesian), linked to
                    ' ** qryStatementParameters_AssetList_19b (qryStatementParameters_AssetList_19a (Account, with
                    ' ** MonthNum = FormRef('MonthNum')), just those matching MonthNum), just scheduled accounts; From .._28a.
1160                .RecordSource = "qryStatementParameters_AssetList_70_14"
                    '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_14_02"
1170              End Select
1180            Case False
                  'NO ASSETS!
1190              Select Case blnPriceHistory
                  Case True
                    ' ** PRICING HISTORY!
                    ' ** qryStatementParameters_AssetList_70_33 (xx), grouped and summed, by accountno; From .._28s.
1200                .RecordSource = "qryStatementParameters_AssetList_70_33"
                    '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_33_02"
1210              Case False
                    ' ** qryStatementParameters_AssetList_70_12 (tmpRelatedAccount_03, with
                    ' ** qryStatementParameters_AssetList_70_11 (qryStatementParameters_AssetList_70_10
                    ' ** (tmpRelatedAccount_03, with Foreign Exchange, linked to Account, grouped and summed,
                    ' ** by accountno), grouped and summed), by specified [ractnos], with TotalCost_usd,
                    ' ** TotalMarket_usd; Cartesian), with qryStatementParameters_AssetList_70_11
                    ' ** (qryStatementParameters_AssetList_70_10 (tmpRelatedAccount_03, with Foreign Exchange, linked
                    ' ** to Account, grouped and summed, by accountno), grouped and summed), from code, with [ractnos]
                    ' ** replaced with actual accountno's, with TotalCost_usd, TotalMarket_usd; Cartesian; From .._28.
1220                .RecordSource = "qryStatementParameters_AssetList_70_13"
                    '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_13_02"
1230              End Select
1240            End Select
1250          Case False
                ' ** Without Related Accounts.
1260            Select Case frm.opgAccountNumber
                Case frm.opgAccountNumber_optSpecified.OptionValue
                  ' ** One Account.
1270              Select Case frm.chkNoAssets
                  Case True
1280                Select Case frm.chkStatements
                    Case True
1290                  Select Case blnPriceHistory
                      Case True
                        ' ** PRICING HISTORY!
                        ' ** qryStatementParameters_AssetList_74_40 (xx), linked to Account, for 'NoAssets', ForEx.
1300                    .RecordSource = "qryStatementParameters_AssetList_70_36_01"
                        '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_26_02"  #########
1310                  Case False
                        ' ** qryStatementParameters_AssetList_74_41 (xx), linked to Account, for 'NoAssets'.
1320                    .RecordSource = "qryStatementParameters_AssetList_70_36_02"
                        '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_06_02"  #########
1330                  End Select
1340                Case False
                      'NO ASSETS!
1350                  Select Case blnPriceHistory
                      Case True
                        ' ** PRICING HISTORY!
                        ' ** Account, linked to ActiveAssets, grouped, with add'l fields; specified
                        ' ** FormRef('accountno'), w/o assetno<>Null, w/ TotalCost_usd, TotalMarket_usd; From .._06c.
1360                    .RecordSource = "qryStatementParameters_AssetList_70_23"
                        '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_23_02"
1370                  Case False
                        ' ** Account, linked to ActiveAssets, grouped, with add'l fields; specified
                        ' ** FormRef('accountno'), w/o assetno<>Null, w/ TotalCost_usd, TotalMarket_usd; From .._06c.
1380                    .RecordSource = "qryStatementParameters_AssetList_70_03"
                        '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_03_02"
1390                  End Select
1400                End Select
1410              Case False
1420                Select Case frm.chkStatements
                    Case True
1430                  Select Case blnPriceHistory
                      Case True
                        ' ** PRICING HISTORY!
                        ' ** qryStatementParameters_AssetList_70_22 (xx), linked to qryStatementParameters_AssetList_19b
                        ' ** (xx), just scheduled accounts; From .._06f.
1440                    .RecordSource = "qryStatementParameters_AssetList_70_25"
                        '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_25_02"
1450                  Case False
                        ' ** qryStatementParameters_AssetList_70_02 (Account, linked to ActiveAssets, grouped, with
                        ' ** add'l fields; specified FormRef('accountno'), w/ TotalCost_usd, TotalMarket_usd), linked
                        ' ** to qryStatementParameters_AssetList_19b (qryStatementParameters_AssetList_19a (Account, with
                        ' ** MonthNum = FormRef('MonthNum')), just those matching MonthNum), just scheduled accounts; From .._06f.
1460                    .RecordSource = "qryStatementParameters_AssetList_70_05"
                        '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_05_02"
1470                  End Select
1480                Case False
1490                  Select Case blnPriceHistory
                      Case True
                        ' ** PRICING HISTORY!
                        ' ** Account, linked to ActiveAssets, grouped, with add'l fields; specified
                        ' ** FormRef('accountno'), w/ TotalCost_usd, TotalMarket_usd; From .._06b.
1500                    .RecordSource = "qryStatementParameters_AssetList_70_22"
                        '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_22_02"
1510                  Case False
                        ' ** Account, linked to ActiveAssets, grouped, with add'l fields; specified
                        ' ** FormRef('accountno'), w/ TotalCost_usd, TotalMarket_usd; From .._06b.
1520                    .RecordSource = "qryStatementParameters_AssetList_70_02"
                        '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_02_02"
1530                  End Select
1540                End Select
1550              End Select
1560            Case frm.opgAccountNumber_optAll.OptionValue
                  ' ** All Accounts.
1570              Select Case frm.chkStatements
                  Case True
1580                Select Case blnNoAssets_All
                    Case True
                      ' ** Union of qryStatementParameters_AssetList_70_29 (qryStatementParameters_AssetList_70_29_01 (xx),
                      ' ** linked to qryStatementParameters_AssetList_70_29_02 (xx)), qryStatementParameters_AssetList_70_29_03
                      ' ** (qryStatementParameters_AssetList_74_65 (xx), linked to Account, with add'l fields).
1590                  .RecordSource = "qryStatementParameters_AssetList_70_29_04"
1600                Case False
1610                  Select Case blnPriceHistory
                      Case True
                        ' ** PRICING HISTORY!
                        ' ** qryStatementParameters_AssetList_70_21 (xx), linked to qryStatementParameters_AssetList_19b
                        ' ** (xx), just scheduled accounts; From .._06e.
1620                    .RecordSource = "qryStatementParameters_AssetList_70_24"
                        '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_24_02"
1630                  Case False
                        ' ** qryStatementParameters_AssetList_70_01 (Account, linked to ActiveAssets, grouped, with
                        ' ** add'l fields; all accounts, w/ TotalCost_usd, TotalMarket_usd), linked to
                        ' ** qryStatementParameters_AssetList_19b (qryStatementParameters_AssetList_19a (Account, with
                        ' ** MonthNum = FormRef('MonthNum')), just those matching MonthNum), just scheduled accounts; From .._06e.
1640                    .RecordSource = "qryStatementParameters_AssetList_70_04"
                        '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_04_02"
1650                  End Select
1660                End Select
1670              Case False
                    'NO ASSETS!
1680                Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** Account, linked to ActiveAssets, grouped, with add'l fields;
                      ' ** all accounts, w/ TotalCost_usd, TotalMarket_usd; From .._06a.
1690                  .RecordSource = "qryStatementParameters_AssetList_70_21"
                      '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_21_02"
1700                Case False
                      ' ** Account, linked to ActiveAssets, grouped, with add'l fields;
                      ' ** all accounts, w/ TotalCost_usd, TotalMarket_usd; From .._06a.
1710                  .RecordSource = "qryStatementParameters_AssetList_70_01"
                      '.rptAssetList_ForEx_Sub_Totals.Report.RecordSource = "qryStatementParameters_AssetList_81_01_02"
1720                End Select
1730              End Select
1740            End Select  ' ** opgAccountNumber.
1750          End Select  ' ** chkRelatedAccounts.
1760        End Select  ' ** gblnMessage.
            'Debug.Print "'" & .RecordSource

1770        .AsOf_lbl.Caption = "As of " & Format(CDate(frm.AssetListDate), "mmmm dd, yyyy")
1780        .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
1790        .AsOf.Caption = "Printed " & Format(Now(), "mmmm dd"", ""yyyy")
1800        .AsOf_lbl2.Caption = "As of " & Format(CDate(frm.AssetListDate), "mmmm dd, yyyy")
1810        .CoInfoBlock2.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
1820        .AsOf2.Caption = "Printed " & Format(Now(), "mmmm dd"", ""yyyy")

1830        If frm.chkDevSaveData.Visible = True And frm.chkDevSaveData = True Then
1840          blnSaveData = True
1850          lngFlds = 1&
1860          ReDim arr_varFld(F_ELEMS, 0)
              ' ******************************************************
              ' ** Array: arr_varFld()
              ' **
              ' **   Field  Element  Name                 Constant
              ' **   =====  =======  ===================  ==========
              ' **     1       0     accountno            F_ACTNO
              ' **     2       1     st_period            F_DATE
              ' **     3       2     st_forex             F_FOREX
              ' **     4       3     sa_totcost           F_CST
              ' **     5       4     sa_totmarket         F_MKT
              ' **     6       5     sa_icash_end         F_ICSH
              ' **     7       6     sa_pcash_end         F_PCSH
              ' **     8       7     sa_grandtotcost      F_TCST
              ' **     9       8     sa_grandtotmarket    F_TMKT
              ' **    10       9     st_recs              F_RECS
              ' **
              ' ******************************************************
1870          Select Case frm.opgAccountNumber
              Case frm.opgAccountNumber_optSpecified.OptionValue
1880            arr_varFld(F_ACTNO, 0) = strAccountNo
1890            arr_varFld(F_DATE, 0) = datEndDate
1900            arr_varFld(F_FOREX, 0) = CBool(True)
1910            arr_varFld(F_CST, 0) = CDbl(0)
1920            arr_varFld(F_MKT, 0) = CDbl(0)
1930            arr_varFld(F_ICSH, 0) = CDbl(0)
1940            arr_varFld(F_PCSH, 0) = CDbl(0)
1950            arr_varFld(F_TCST, 0) = CDbl(0)
1960            arr_varFld(F_TMKT, 0) = CDbl(0)
1970            arr_varFld(F_RECS, 0) = CLng(0)
1980          Case frm.opgAccountNumber_optAll.OptionValue
1990            arr_varFld(F_ACTNO, 0) = Null
2000            arr_varFld(F_DATE, 0) = datEndDate
2010            arr_varFld(F_FOREX, 0) = CBool(True)
2020            arr_varFld(F_CST, 0) = CDbl(0)
2030            arr_varFld(F_MKT, 0) = CDbl(0)
2040            arr_varFld(F_ICSH, 0) = CDbl(0)
2050            arr_varFld(F_PCSH, 0) = CDbl(0)
2060            arr_varFld(F_TCST, 0) = CDbl(0)
2070            arr_varFld(F_TMKT, 0) = CDbl(0)
2080            arr_varFld(F_RECS, 0) = CLng(0)
2090          End Select
2100        End If

2110        Select Case gblnCombineAssets
            Case True
2120          .txtIcash.Visible = False
2130          .txtPcash.Visible = False
2140          .txtPcash2.Visible = False
2150          .txtIcash2.Visible = False
2160          .txtPcashPctMarket.Visible = False
2170          .txtIcashPctMarket.Visible = False
2180          .txtCash_lbl.Visible = True
2190          .txtCash.Visible = True
2200          .txtCash2.Visible = True
2210          .txtCashPctMarket.Visible = True
2220        Case False
2230          .txtIcash.Visible = True
2240          .txtPcash.Visible = True
2250          .txtPcash2.Visible = True
2260          .txtIcash2.Visible = True
2270          .txtPcashPctMarket.Visible = True
2280          .txtIcashPctMarket.Visible = True
2290          .txtCash_lbl.Visible = False
2300          .txtCash.Visible = False
2310          .txtCash2.Visible = False
2320          .txtCashPctMarket.Visible = False
2330        End Select  ' ** gblnCombineAssets.

2340        If glngPrintRpts > 0& Then
2350          For lngZ = 0& To (glngPrintRpts - 1&)
2360            If garr_varPrintRpt(PR_ACTNO, lngZ) = gstrAccountNo Then
2370              If garr_varPrintRpt(PR_ALIST, lngZ) = True Then
2380                blnSkip = True
2390              Else
2400                garr_varPrintRpt(PR_ALIST, lngZ) = CBool(True)
2410              End If
2420              Exit For
2430            End If
2440          Next
2450        End If

2460      Else
2470        Cancel = -1
2480      End If
2490    End With

2500    If Cancel = -1 Then
2510      Beep
2520      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
2530    ElseIf blnSkip = True Then
2540      Cancel = -1
2550    End If

        ' ** VGC 02/14/2010:
        ' ** Existing equation on PercentMarket:
        '=IIf([TotalShareFace_Grand_sum]=0,'0.00%',
        'IIf((([marketvaluecurrentx]*[TotalShareface])/[TotalShareFace_Grand_sum])>1 Or
        '(([marketvaluecurrentx]*[TotalShareface])/[TotalShareFace_Grand_sum])<-1,'NA',
        'IIf(IsNull([marketvaluecurrentx])=True Or IsNull([TotalShareface])=True,'0.00%',
        'Format((([marketvaluecurrentx]*[TotalShareface])/[TotalShareFace_Grand_sum]),'Percent'))))
        ' ** Changed <-1 to <0.
        ' ** Now:
        '=IIf([TotalMarket_Grand_sum]=0,'0.00%',IIf((([MarketValueCurrentX]*[TotalShareface])/[TotalMarket_Grand_sum])>1 Or (([MarketValueCurrentX]*[TotalShareface])/[TotalMarket_Grand_sum])<0,'NA',IIf(IsNull([MarketValueCurrentX])=True Or IsNull([TotalShareface])=True,'0.00%',Format((([MarketValueCurrentX]*[TotalShareface])/[TotalMarket_Grand_sum]),'Percent'))))

        ' ** Existing equation on TotalPctMarket_Asset_sum:
        '=IIf(IsNull([TotalShareFace_Grand_sum])=True Or IsNull(Sum([marketvaluecurrentx]))=True Or IsNull(Sum([TotalShareface]))=True,0,
        'IIf([TotalShareFace_Grand_sum]=0,0,[percentMarket2x]/[TotalShareFace_Grand_sum]))
        ' ** Added IIf() to last calculation.
        ' ** Now:
        '=IIf(IsNull([TotalMarket_Grand_sum])=True Or IsNull(Sum([MarketValueCurrentX]))=True Or IsNull(Sum([TotalShareface]))=True,0,IIf([TotalMarket_Grand_sum]=0,0,IIf(([percentMarket2x]/[TotalMarket_Grand_sum])>1 Or ([percentMarket2x]/[TotalMarket_Grand_sum])<0,'NA',([percentMarket2x]/[TotalMarket_Grand_sum]))))

        ' ** Existing equation on TotalPctMarket_sum:
        '=IIf([TotalShareFace_Grand_sum]=0,0,IIf(([percentMarket3x]/[TotalShareFace_Grand_sum])=0,
        '((1-(IIf([TotalShareFace_Grand_sum]=0,0,IIf(IsNull([icash]),0,[icash])/[TotalShareFace_Grand_sum])))-
        '(IIf([TotalShareFace_Grand_sum]=0,0,IIf(IsNull([pcash]),0,[pcash])/[TotalShareFace_Grand_sum]))),
        '([percentMarket3x]/[TotalShareFace_Grand_sum])))
        ' ** Now:
        '=IIf(Nz([TotalMarket_Grand_sum],0)=0,0,IIf((IIf((Nz([percentMarket3x],0)/[TotalMarket_Grand_sum])=0,
        '((1-(Nz([icash],0)/[TotalMarket_Grand_sum]))-(Nz([pcash],0)/[TotalMarket_Grand_sum])),
        '([percentMarket3x]/[TotalMarket_Grand_sum])))>1 Or
        '(IIf((Nz([percentMarket3x],0)/[TotalMarket_Grand_sum])=0,
        '((1-(Nz([icash],0)/[TotalMarket_Grand_sum]))-(Nz([pcash],0)/[TotalMarket_Grand_sum])),
        '([percentMarket3x]/[TotalMarket_Grand_sum])))<0, 'NA',
        'IIf((Nz([percentMarket3x],0)/[TotalMarket_Grand_sum])=0,
        '((1-(Nz([icash],0)/[TotalMarket_Grand_sum]))-(Nz([pcash],0)/[TotalMarket_Grand_sum])),
        '([percentMarket3x]/[TotalMarket_Grand_sum]))))

        ' ** Existing equation on txtIcashPctMarket:
        '=IIf([TotalShareFace_Grand_sum]=0,0,IIf(IsNull([icash]),0,[icash])/[TotalShareFace_Grand_sum])
        ' ** Now:
        '=IIf([TotalMarket_Grand_sum]=0,0,IIf((IIf(IsNull([icash]),0,[icash])/[TotalMarket_Grand_sum])>1 Or
        '(IIf(IsNull([icash]),0,[icash])/[TotalMarket_Grand_sum])<0,'NA',(IIf(IsNull([icash]),0,[icash])/[TotalMarket_Grand_sum])))

        ' ** Existing equation on txtCashPctMarket:
        '=IIf([TotalShareFace_Grand_sum]=0,0,([icash]+[pcash])/[TotalShareFace_Grand_sum])
        ' ** Now:
        '=IIf([TotalMarket_Grand_sum]=0,0,IIf((([icash]+[pcash])/[TotalMarket_Grand_sum])>1 Or (([icash]+[pcash])/[TotalMarket_Grand_sum])<0,'NA',(([icash]+[pcash])/[TotalMarket_Grand_sum])))

        ' ** Existing equation on txtPcashPctMarket:
        '=IIf([TotalShareFace_Grand_sum]=0,0,IIf(IsNull([pcash]),0,[pcash])/[TotalShareFace_Grand_sum])
        ' ** Now:
        '=IIf([TotalMarket_Grand_sum]=0,0,IIf((IIf(IsNull([pcash]),0,[pcash])/[TotalMarket_Grand_sum])>1 Or (IIf(IsNull([pcash]),0,[pcash])/[TotalMarket_Grand_sum])<0,'NA',(IIf(IsNull([pcash]),0,[pcash])/[TotalMarket_Grand_sum])))

        ' ** Existing equation on TotalPctMarket_Grand_sum:
        '=IIf([TotalMarket_Grand_sum]=0,0,(Sum([MarketValueCurrentX]*[TotalShareface])+IIf(IsNull([icash]),0,[icash])+
        'IIf(IsNull([pcash]),0,[pcash]))/[TotalMarket_Grand_sum])
        ' ** Now:
        '=IIf([TotalMarket_Grand_sum]=0,0,IIf(((Sum([marketvaluecurrentx]*[TotalShareface])+
        'IIf(IsNull([icash]),0,[icash])+IIf(IsNull([pcash]),0,[pcash]))/[TotalMarket_Grand_sum])>1 Or
        '((Sum([marketvaluecurrentx]*[TotalShareface])+IIf(IsNull([icash]),0,[icash])+
        'IIf(IsNull([pcash]),0,[pcash]))/[TotalMarket_Grand_sum])<0,'NA',
        '((Sum([marketvaluecurrentx]*[TotalShareface])+
        'IIf(IsNull([icash]),0,[icash])+IIf(IsNull([pcash]),0,[pcash]))/[TotalMarket_Grand_sum])))

        ' ** Now, in English:
        'If [TotalShareFace_Grand_sum] is Zero, stop and return Zero.
        'Otherwise
        '  If [percentMarket3x] divided by [TotalShareFace_Grand_sum] is Zero, then
        '    Subtract {this} from one, then, from that difference, subtract {the other number}
        '      or
        '    Subtract [icash] divided by [TotalShareFace_Grand_sum] from one, then, from that difference,
        '      subtract [pcash] divided by [TotalShareFace_Grand_sum].
        '    and so we stop and return
        '      (1 - ([icash] / [TotalShareFace_Grand_sum])) - ([pcash] / [TotalShareFace_Grand_sum])
        '  Otherwise
        '    stop and return [percentMarket3x] divided by [TotalShareFace_Grand_sum].

        ' ** As programming:
        'If [TotalShareFace_Grand_sum] = 0 Then
        '  dblRetVal = 0
        'Else
        '  If ([percentMarket3x] / [TotalShareFace_Grand_sum]) = 0 Then
        '    dblRetVal = ([icash] / [TotalShareFace_Grand_sum])
        '    dblRetVal = (1 - dblRetVal)
        '    dblRetVal = (dblRetVal - ([pcash] / [TotalShareFace_Grand_sum]))
        '  Else
        '    dblRetVal = ([percentMarket3x] / [TotalShareFace_Grand_sum])
        '  End If
        'End If

        'Nz() everthing!

        'IIf( Nz([TotalShareFace_Grand_sum],0) = 0, 0,
        '  IIf( (Nz([percentMarket3x],0) / [TotalShareFace_Grand_sum]) = 0,
        '    ((1 - (Nz([icash],0) / [TotalShareFace_Grand_sum])) - (Nz([pcash],0) / [TotalShareFace_Grand_sum])) ,
        '    ([percentMarket3x] / [TotalShareFace_Grand_sum])
        '  )
        ')
        ' ** Voila!
        '=IIf(Nz([TotalShareFace_Grand_sum],0)=0,0,IIf((Nz([percentMarket3x],0)/[TotalShareFace_Grand_sum])=0,((1-(Nz([icash],0)/[TotalShareFace_Grand_sum]))-(Nz([pcash],0)/[TotalShareFace_Grand_sum])),([percentMarket3x]/[TotalShareFace_Grand_sum])))

EXITP:
2560    Exit Sub

ERRH:
2570    If Forms(strCallingForm).Visible = False Then
2580      Forms(strCallingForm).Visible = True
2590    End If
2600    Select Case ERR.Number
        Case 2585  ' ** This action can't be carried out while processing a form or report event.
          ' ** Ignore.
2610    Case Else
2620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2630    End Select
2640    Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

        Dim datCurrentDate As Date
        Dim varTmp00 As Variant

2710    With Me

2720      Select Case IsNull(frm.currentDate)
          Case True
2730        varTmp00 = DMax("[currentDate]", "masterasset")
2740        datCurrentDate = varTmp00
2750      Case False
2760        datCurrentDate = frm.currentDate
2770      End Select
2780      If frm.UsePriceHistory = True Then
2790        .MarketAsOf_lbl.Caption = "Market as of varies"
2800        .MarketAsOf_lbl2.Caption = "Market as of varies"
2810      Else
2820        If datCurrentDate <> datEndDate Then
2830          .MarketAsOf_lbl.Caption = "Market as of " & Format(datCurrentDate, "mm/dd/yyyy")
2840          .MarketAsOf_lbl2.Caption = "Market as of " & Format(datCurrentDate, "mm/dd/yyyy")
2850        Else
2860          .MarketAsOf_lbl.Caption = vbNullString
2870          .MarketAsOf_lbl2.Caption = vbNullString
2880        End If
2890      End If

2900      If frm.chkStatements = True Then  ' ** Doing statements.
2910        .AsOf.Visible = False
2920        .AsOf2.Visible = False
2930      End If

2940    End With

EXITP:
2950    Exit Sub

ERRH:
2960    If Forms(strCallingForm).Visible = False Then
2970      Forms(strCallingForm).Visible = True
2980    End If
2990    Select Case ERR.Number
        Case Else
3000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3010    End Select
3020    Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

3010    With Me
3020      If .Page >= 2 Then
3030        .Section(acPageHeader).Visible = True
3040      Else
3050        .Section(acPageHeader).Visible = False
3060      End If
3070    End With

EXITP:
3080    Exit Sub

ERRH:
3090    If Forms(strCallingForm).Visible = False Then
3100      Forms(strCallingForm).Visible = True
3110    End If
3120    Select Case ERR.Number
        Case Else
3130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3140    End Select
3150    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim blnFound As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long, lngE As Long

3210    With Me

3220      Select Case frm.chkLegalName
          Case True
3230        varTmp00 = FixAmps(Nz(.legalname, vbNullString))  ' ** Module Functions: modStringFuncs.
3240        If IsNull(varTmp00) = True Then
3250          varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
3260        End If
3270        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
3280        .Name_lbl.Caption = varTmp00
3290      Case False
3300        varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
3310        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
3320        .Name_lbl.Caption = varTmp00
3330      End Select  ' ** chkLegalName.

          ' ** blnNoAssets applies to singles only.
3340      If blnNoAssets = True Then
            ' ** GroupHeader2:
3350        .assettype_description_display.Visible = False
3360        .assettype_description_box.Visible = False
            ' ** Detail:
3370        .PercentMarket.Visible = False
3380        .PercentMarket_usd.Visible = False
3390        .totdesc_disp.Visible = True  ' ** Yes, show.
3400        .totdesc_disp.ForeColor = CLR_GRY
3410        .UnitCost.ForeColor = CLR_GRY
3420        .TotalCost.ForeColor = CLR_GRY
3430        .TotalCost_usd.ForeColor = CLR_GRY
3440        .UnitMarket.ForeColor = CLR_GRY
3450        .TotalMarket.ForeColor = CLR_GRY
3460        .TotalMarket_usd.ForeColor = CLR_GRY
3470        .AnnualIncome.ForeColor = CLR_GRY
3480        .AnnualIncome_usd.ForeColor = CLR_GRY
            ' ** GroupFooter2:
3490        .UnitCost_sum_line.BorderColor = CLR_GRY
3500        .Cost_sum_line.BorderColor = CLR_GRY
3510        .UnitMarket_sum_line.BorderColor = CLR_GRY
3520        .Market_sum_line.BorderColor = CLR_GRY
3530        .TotalCost_Asset_sum_line.BorderColor = CLR_GRY
3540        .TotalMarket_Asset_sum_line.BorderColor = CLR_GRY
3550        .TotalPctMarket_Asset_sum_line.BorderColor = CLR_GRY
3560        .TotalCost_Asset_sum.ForeColor = CLR_GRY
3570        .TotalMarket_Asset_sum.ForeColor = CLR_GRY
3580        .TotalPctMarket_Asset_sum.ForeColor = CLR_GRY
3590        .TotalPctMarket_Asset_sum.Visible = False
3600        .TotalPctMarket_Asset_sum_usd.Visible = False
            ' ** Subreport:
3610        .rptAssetList_ForEx_Totals_lbl.Visible = False
3620      End If

          ' ** blnNoAssets_All applies to 'All' statements only.
3630      If blnNoAssets_All = True Then
            ' ** First, turn then all back on.

            ' ** GroupHeader2:
3640        .assettype_description_display.Visible = True
3650        .assettype_description_box.Visible = True
            ' ** Detail:
3660        Select Case blnPriceHistory
            Case True
3670          .PercentMarket_usd.Visible = True
3680          .AnnualIncome_usd.Visible = True
3690        Case False
3700          .PercentMarket.Visible = True
3710          .AnnualIncome.Visible = True
3720        End Select
3730        .totdesc_disp.Visible = True
3740        .totdesc_disp.ForeColor = CLR_BLK
3750        .UnitCost.ForeColor = CLR_BLK
3760        .TotalCost.ForeColor = CLR_BLK
3770        .TotalCost_usd.ForeColor = CLR_BLK
3780        .UnitMarket.ForeColor = CLR_BLK
3790        .TotalMarket.ForeColor = CLR_BLK
3800        .TotalMarket_usd.ForeColor = CLR_BLK
3810        .AnnualIncome.ForeColor = CLR_BLK
3820        .AnnualIncome_usd.ForeColor = CLR_BLK
            ' ** GroupFooter2:
3830        .UnitCost_sum_line.BorderColor = CLR_BLK
3840        .Cost_sum_line.BorderColor = CLR_BLK
3850        .UnitMarket_sum_line.BorderColor = CLR_BLK
3860        .Market_sum_line.BorderColor = CLR_BLK
3870        .TotalCost_Asset_sum_line.BorderColor = CLR_BLK
3880        .TotalMarket_Asset_sum_line.BorderColor = CLR_BLK
3890        .TotalPctMarket_Asset_sum_line.BorderColor = CLR_BLK
3900        .UnitCost_sum_line.Visible = True
3910        .Cost_sum_line.Visible = True
3920        .UnitMarket_sum_line.Visible = True
3930        .Market_sum_line.Visible = True
3940        .TotalCost_Asset_sum_line.Visible = True
3950        .TotalMarket_Asset_sum_line.Visible = True
3960        .TotalPctMarket_Asset_sum_line.Visible = True
3970        .TotalCost_Asset_sum.ForeColor = CLR_BLK
3980        .TotalMarket_Asset_sum.ForeColor = CLR_BLK
3990        .TotalCost_Asset_sum.Visible = True
4000        .TotalMarket_Asset_sum.Visible = True
4010        .TotalPctMarket_Asset_sum.ForeColor = CLR_BLK
4020        Select Case blnPriceHistory
            Case True
4030          .TotalPctMarket_Asset_sum_usd.Visible = True
4040        Case False
4050          .TotalPctMarket_Asset_sum.Visible = True
4060        End Select
            ' ** GroupFooter1:
4070        .UnitCost_tot_sum_line.Visible = True
4080        .Cost_tot_sum_line.Visible = True
4090        .UnitMarket_tot_sum_line.Visible = True
4100        .Market_tot_sum_line.Visible = True
4110        .TotalCost_sum_line.Visible = True
4120        .TotalMarket_sum_line.Visible = True
4130        .TotalPctMarket_sum_line.Visible = True
4140        Select Case blnPriceHistory
            Case True
4150          .TotalPctMarket_sum_usd.Visible = True
4160        Case False
4170          .TotalPctMarket_sum.Visible = True
4180        End Select
4190        .TotalCost_sum.Visible = True
4200        .TotalMarket_sum.Visible = True
4210        Select Case gblnCombineAssets
            Case True
4220          .txtCash.Visible = True
4230          .txtCash2.Visible = True
4240          .txtCashPctMarket.Visible = True
4250        Case False
4260          .txtIcash.Visible = True
4270          .txtIcash2.Visible = True
4280          .txtIcashPctMarket.Visible = True
4290          .txtPcash.Visible = True
4300          .txtPcash2.Visible = True
4310          .txtPcashPctMarket.Visible = True
4320        End Select
4330        .UnitCost_Grand_sum_line.Visible = True
4340        .Cost_Grand_sum_line.Visible = True
4350        .UnitMarket_Grand_sum_line.Visible = True
4360        .Market_Grand_sum_line.Visible = True
4370        .TotalCost_Grand_sum_line.Visible = True
4380        .TotalMarket_Grand_sum_line.Visible = True
4390        .TotalPctMarket_Grand_sum_line.Visible = True
4400        .MarketValueCurrent_line.Visible = True
4410        Select Case blnPriceHistory
            Case True
4420          .TotalPctMarket_Grand_sum_usd.Visible = True
4430          .MarketValueCurrent_usd.Visible = True
4440        Case False
4450          .TotalPctMarket_Grand_sum.Visible = True
4460          .marketvaluecurrent.Visible = True
4470        End Select
4480        .TotalCost_Grand_sum.Visible = True
4490        .TotalMarket_Grand_sum.Visible = True
            ' ** Subreport:
4500        .rptAssetList_ForEx_Totals_lbl.Visible = True

4510        If .NoAssets = True Then

              ' ** GroupHeader2:
4520          .assettype_description_display.Visible = False
4530          .assettype_description_box.Visible = False
              ' ** Detail:
4540          .PercentMarket.Visible = False
4550          .PercentMarket_usd.Visible = False
4560          .totdesc_disp.Visible = True  ' ** Yes, show.
4570          .totdesc_disp.ForeColor = CLR_GRY
4580          .UnitCost.ForeColor = CLR_GRY
4590          .TotalCost.ForeColor = CLR_GRY
4600          .TotalCost_usd.ForeColor = CLR_GRY
4610          .UnitMarket.ForeColor = CLR_GRY
4620          .TotalMarket.ForeColor = CLR_GRY
4630          .TotalMarket_usd.ForeColor = CLR_GRY
4640          .AnnualIncome.ForeColor = CLR_GRY
4650          .AnnualIncome_usd.ForeColor = CLR_GRY
              ' ** GroupFooter2:
4660          .UnitCost_sum_line.BorderColor = CLR_GRY
4670          .Cost_sum_line.BorderColor = CLR_GRY
4680          .UnitMarket_sum_line.BorderColor = CLR_GRY
4690          .Market_sum_line.BorderColor = CLR_GRY
4700          .TotalCost_Asset_sum_line.BorderColor = CLR_GRY
4710          .TotalMarket_Asset_sum_line.BorderColor = CLR_GRY
4720          .TotalPctMarket_Asset_sum_line.BorderColor = CLR_GRY
4730          .TotalCost_Asset_sum.ForeColor = CLR_GRY
4740          .TotalMarket_Asset_sum.ForeColor = CLR_GRY
4750          .TotalPctMarket_Asset_sum.ForeColor = CLR_GRY
4760          .TotalPctMarket_Asset_sum.Visible = False
4770          .TotalPctMarket_Asset_sum_usd.Visible = False
              ' ** Subreport:
4780          .rptAssetList_ForEx_Totals_lbl.Visible = False

4790        ElseIf .NoTrans = True Then

              ' ** GroupHeader2:
4800          .assettype_description_display.Visible = False
4810          .assettype_description_box.Visible = False
              ' ** Detail:
4820          .totdesc_disp.Visible = True  ' ** Yes, show.
4830          .totdesc_disp.ForeColor = CLR_GRY
4840          .UnitCost.ForeColor = CLR_GRY
4850          .TotalCost.ForeColor = CLR_GRY
4860          .TotalCost_usd.ForeColor = CLR_GRY
4870          .UnitMarket.ForeColor = CLR_GRY
4880          .TotalMarket.ForeColor = CLR_GRY
4890          .TotalMarket_usd.ForeColor = CLR_GRY
4900          .PercentMarket.Visible = False
4910          .PercentMarket_usd.Visible = False
4920          .AnnualIncome.Visible = False
4930          .AnnualIncome_usd.Visible = False
              ' ** GroupFooter2:
4940          .UnitCost_sum_line.Visible = False
4950          .Cost_sum_line.Visible = False
4960          .UnitMarket_sum_line.Visible = False
4970          .Market_sum_line.Visible = False
4980          .TotalCost_Asset_sum_line.Visible = False
4990          .TotalMarket_Asset_sum_line.Visible = False
5000          .TotalPctMarket_Asset_sum_line.Visible = False
5010          .TotalCost_Asset_sum.Visible = False
5020          .TotalMarket_Asset_sum.Visible = False
5030          .TotalPctMarket_Asset_sum.Visible = False
5040          .TotalPctMarket_Asset_sum_usd.Visible = False
              ' ** GroupFooter1:
5050          .UnitCost_tot_sum_line.Visible = False
5060          .Cost_tot_sum_line.Visible = False
5070          .UnitMarket_tot_sum_line.Visible = False
5080          .Market_tot_sum_line.Visible = False
5090          .TotalCost_sum_line.Visible = False
5100          .TotalMarket_sum_line.Visible = False
5110          .TotalPctMarket_sum_line.Visible = False
5120          .TotalCost_sum.Visible = False
5130          .TotalMarket_sum.Visible = False
5140          .TotalPctMarket_sum.Visible = False
5150          .TotalPctMarket_sum_usd.Visible = False
5160          .txtIcash.Visible = False
5170          .txtIcash2.Visible = False
5180          .txtIcashPctMarket.Visible = False
5190          .txtCash.Visible = False
5200          .txtCash2.Visible = False
5210          .txtCashPctMarket.Visible = False
5220          .txtPcash.Visible = False
5230          .txtPcash2.Visible = False
5240          .txtPcashPctMarket.Visible = False
5250          .UnitCost_Grand_sum_line.Visible = False
5260          .Cost_Grand_sum_line.Visible = False
5270          .UnitMarket_Grand_sum_line.Visible = False
5280          .Market_Grand_sum_line.Visible = False
5290          .TotalCost_Grand_sum_line.Visible = False
5300          .TotalMarket_Grand_sum_line.Visible = False
5310          .TotalPctMarket_Grand_sum_line.Visible = False
5320          .MarketValueCurrent_line.Visible = False
5330          .TotalCost_Grand_sum.Visible = False
5340          .TotalMarket_Grand_sum.Visible = False
5350          .TotalPctMarket_Grand_sum.Visible = False
5360          .TotalPctMarket_Grand_sum_usd.Visible = False
5370          .marketvaluecurrent.Visible = False
5380          .MarketValueCurrent_usd.Visible = False
              ' ** Subreport:
5390          .rptAssetList_ForEx_Totals_lbl.Visible = False

5400        End If

5410      End If

5420      If blnSaveData = True And blnStatements = True Then
5430        Select Case IsNull(arr_varFld(F_ACTNO, 0))
            Case True
              ' ** First accountno.
5440          arr_varFld(F_ACTNO, 0) = .accountno
5450        Case False
5460          blnFound = False
5470          For lngX = 0& To (lngFlds - 1&)
5480            If arr_varFld(F_ACTNO, lngX) = .accountno Then
5490              blnFound = True
5500              Exit For
5510            End If
5520          Next
5530          If blnFound = False Then
5540            lngFlds = lngFlds + 1&
5550            lngE = lngFlds - 1&
5560            ReDim Preserve arr_varFld(F_ELEMS, lngE)
5570            arr_varFld(F_ACTNO, lngE) = .accountno
5580            arr_varFld(F_DATE, lngE) = datEndDate
5590            arr_varFld(F_FOREX, lngE) = CBool(True)
5600            arr_varFld(F_CST, lngE) = CDbl(0)
5610            arr_varFld(F_MKT, lngE) = CDbl(0)
5620            arr_varFld(F_ICSH, lngE) = CDbl(0)
5630            arr_varFld(F_PCSH, lngE) = CDbl(0)
5640            arr_varFld(F_TCST, lngE) = CDbl(0)
5650            arr_varFld(F_TMKT, lngE) = CDbl(0)
5660            arr_varFld(F_RECS, lngE) = CLng(0)
5670          End If
5680        End Select
5690      End If

5700    End With

EXITP:
5710    Exit Sub

ERRH:
5720    If Forms(strCallingForm).Visible = False Then
5730      Forms(strCallingForm).Visible = True
5740    End If
5750    Select Case ERR.Number
        Case Else
5760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5770    End Select
5780    Resume EXITP

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Format"

        Dim strFormat As String

5810    With Me
5820      Select Case .currsym_symbol
          Case "CHF"
5830        strFormat = "\C\H\F#,##0.00;(\C\H\F#,##0.00);\C\H\F0.00;\C\H\F0.00"
5840      Case Else
5850        strFormat = .currsym_symbol & "#,##0.00;(" & .currsym_symbol & "#,##0.00);" & .currsym_symbol & "0.00;" & .currsym_symbol & "0.00"
5860      End Select
5870      .TotalCost.Format = strFormat
5880      .TotalMarket.Format = strFormat
5890    End With

EXITP:
5900    Exit Sub

ERRH:
5910    If Forms(strCallingForm).Visible = False Then
5920      Forms(strCallingForm).Visible = True
5930    End If
5940    Select Case ERR.Number
        Case 2585  ' ** This action can't be carried out while processing a form or report event.
          ' ** Ignore.
5950    Case Else
5960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5970    End Select
5980    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

        Dim lngX As Long

6010    With Me
6020      If blnSaveData = True Then
6030        For lngX = 0& To (lngFlds - 1&)
6040          If arr_varFld(F_ACTNO, lngX) = .accountno Then
6050            arr_varFld(F_CST, lngX) = .TotalCost_sum
6060            arr_varFld(F_MKT, lngX) = .TotalMarket_sum
6070            arr_varFld(F_ICSH, lngX) = .txtIcash
6080            arr_varFld(F_PCSH, lngX) = .txtPcash
6090            arr_varFld(F_TCST, lngX) = .TotalCost_Grand_sum
6100            arr_varFld(F_TMKT, lngX) = .TotalMarket_Grand_sum
6110            Exit For
6120          End If
6130        Next
6140      End If
6150    End With

EXITP:
6160    Exit Sub

ERRH:
6170    If Forms(strCallingForm).Visible = False Then
6180      Forms(strCallingForm).Visible = True
6190    End If
6200    Select Case ERR.Number
        Case Else
6210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6220    End Select
6230    Resume EXITP

End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "ReportFooter_Format"

        'Dim dbs As DAO.Database, rst As DAO.Recordset
        'Dim strAccountNo As String, strAssetListDate As String, datBalanceDate As Date
        'Dim dblTotCost As Double, dblTotMkt As Double, dblICash As Double, dblPCash As Double, dblGrandCost As Double, dblGrandMkt As Double
        Dim blnSkip As Boolean

6310    With Me
6320      blnSkip = True
6330      If blnSkip = False Then
            'strAccountNo = .accountno
            'strAssetListDate = .AsOf_lbl.Caption
            'strAssetListDate = GetLastWord(strAssetListDate)  ' ** Module Function: modStringFuncs.
            'datBalanceDate = CDate(strAssetListDate)
            'dblTotCost = .TotalCost_sum
            'dblTotMkt = .TotalMarket_sum
            'dblICash = .txtIcash
            'dblPCash = .txtPcash
            'dblGrandCost = .TotalCost_Grand_sum
            'dblGrandMkt = .TotalMarket_Grand_sum
            'Set dbs = CurrentDb
            'With dbs
            '  Set rst = .OpenRecordset("zz_tbl_MasterTrust_Balance_01", dbOpenDynaset, dbAppendOnly)
            '  With rst
            '    .AddNew
            '    ' ** ![mtbal_id] : AutoNumber.
            '    ![accountno] = strAccountNo
            '    ![balance_date] = datBalanceDate
            '    ![total_cost] = dblTotCost
            '    ![total_market] = dblTotMkt
            '    ![ICash] = dblICash
            '    ![PCash] = dblPCash
            '    ![grand_cost] = dblGrandCost
            '    ![grand_market] = dblGrandMkt
            '    ![mtbal_datemodified] = Now()
            '    .Update
            '    .Close
            '  End With
            '  .Close
            'End With
6340      End If  ' ** blnSkip.
6350    End With

EXITP:
        'Set rst = Nothing
        'Set dbs = Nothing
6360    Exit Sub

ERRH:
6370    If Forms(strCallingForm).Visible = False Then
6380      Forms(strCallingForm).Visible = True
6390    End If
6400    Select Case ERR.Number
        Case Else
6410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6420    End Select
6430    Resume EXITP

End Sub

Private Sub Report_Close()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim blnAdd As Boolean, blnAddAll As Boolean
        Dim lngX As Long

6510    With Me

6520      Set frm = Nothing
6530      TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

6540      If blnSaveData = True Then
6550        blnAdd = False: blnAddAll = False
6560        Set dbs = CurrentDb
6570        With dbs
6580          Set rst = .OpenRecordset("zz_tbl_Statement_AssetList", dbOpenDynaset, dbConsistent)
6590          With rst
6600            If .BOF = True And .EOF = True Then
6610              blnAddAll = True
6620            End If
6630            For lngX = 0& To (lngFlds - 1&)
6640              If IsNull(arr_varFld(F_ACTNO, lngX)) = False Then
6650                If Trim(arr_varFld(F_ACTNO, lngX)) <> vbNullString Then
6660                  blnAdd = False
6670                  Select Case blnAddAll
                      Case True
6680                    blnAdd = True
6690                  Case False
6700                    .FindFirst "[accountno] = '" & arr_varFld(F_ACTNO, lngX) & "' And " & _
                          "[sa_period] = #" & Format(arr_varFld(F_DATE, lngX), "mm/dd/yyyy") & "# And " & _
                          "[sa_forex] = true"
6710                    If .NoMatch = True Then
6720                      blnAdd = True
6730                    End If
6740                  End Select
6750                  Select Case blnAdd
                      Case True
6760                    .AddNew
6770                    ![accountno] = arr_varFld(F_ACTNO, lngX)
6780                    ![sa_period] = arr_varFld(F_DATE, lngX)
6790                    ![sa_forex] = arr_varFld(F_FOREX, lngX)
6800                  Case False
6810                    .Edit
6820                  End Select
6830                  ![sa_totcost] = arr_varFld(F_CST, lngX)
6840                  ![sa_totmarket] = arr_varFld(F_MKT, lngX)
6850                  ![sa_icash_end] = arr_varFld(F_ICSH, lngX)
6860                  ![sa_pcash_end] = arr_varFld(F_PCSH, lngX)
6870                  ![sa_grandtotcost] = arr_varFld(F_TCST, lngX)
6880                  ![sa_grandtotmarket] = arr_varFld(F_TMKT, lngX)
6890                  ![sa_recs] = arr_varFld(F_RECS, lngX)
6900                  ![sa_datemodified] = Now()
6910                  .Update
6920                End If
6930              End If
6940            Next
6950            .Close
6960          End With
6970          .Close
6980        End With
6990      End If

7000      If strCallingForm <> vbNullString Then
7010        gblnReportClose = True
7020        If Forms(strCallingForm).Visible = False Then
7030          DoCmd.Restore
7040          Forms(strCallingForm).Visible = True
7050        End If
7060        Forms(strCallingForm).TimerInterval = 250&
7070      End If

7080    End With

EXITP:
7090    Set rst = Nothing
7100    Set dbs = Nothing
7110    Exit Sub

ERRH:
7120    If Forms(strCallingForm).Visible = False Then
7130      Forms(strCallingForm).Visible = True
7140    End If
7150    Select Case ERR.Number
        Case Else
7160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7170    End Select
7180    Resume EXITP

End Sub

Private Sub VerifyDates()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "VerifyDates"

        Dim varTmp00 As Variant

7210    If IsNull(frm.DateStart) = True Then
7220      If CDbl(gdatStartDate) > 0 Then
7230        If gdatEndDate = frm.AssetListDate Then
7240          frm.DateStart = gdatStartDate
7250        Else
7260          If strAccountNo = "All" Then
7270            frm.DateStart = #1/1/1900#
7280          Else
7290            varTmp00 = DLookup("[documentdate]", "account", "[accountno] = '" & strAccountNo & "'")
7300            If IsNull(varTmp00) = True Then
7310              varTmp00 = DLookup("[appointmentdate]", "account", "[accountno] = '" & strAccountNo & "'")
7320              If IsNull(varTmp00) = True Then
7330                frm.DateStart = #1/1/1900#
7340              Else
7350                frm.DateStart = CDate(varTmp00)
7360              End If
7370            Else
7380              frm.DateStart = CDate(varTmp00)
7390            End If
7400          End If
7410        End If
7420      Else
7430        If strAccountNo = "All" Then
7440          frm.DateStart = #1/1/1900#
7450        Else
7460          varTmp00 = DLookup("[documentdate]", "account", "[accountno] = '" & strAccountNo & "'")
7470          If IsNull(varTmp00) = True Then
7480            varTmp00 = DLookup("[appointmentdate]", "account", "[accountno] = '" & strAccountNo & "'")
7490            If IsNull(varTmp00) = True Then
7500              frm.DateStart = #1/1/1900#
7510            Else
7520              frm.DateStart = CDate(varTmp00)
7530            End If
7540          Else
7550            frm.DateStart = CDate(varTmp00)
7560          End If
7570        End If
7580      End If
7590    End If
7600    DoEvents
7610    If IsNull(frm.AssetListDate) = True Then
7620      If IsNull(frm.DateEnd) = True Then
7630        If CDbl(gdatEndDate) > 0 Then
7640          If gdatEndDate > frm.DateStart Then
7650            frm.AssetListDate = gdatEndDate
7660          Else
7670            frm.AssetListDate = Date
7680          End If
7690        Else
7700          frm.AssetListDate = Date
7710        End If
7720      Else
7730        If frm.DateEnd > frm.DateStart Then
7740          frm.AssetListDate = frm.DateEnd
7750        Else
7760          frm.AssetListDate = Date
7770        End If
7780      End If
7790    End If
7800    DoEvents

EXITP:
7810    Exit Sub

ERRH:
7820    If Forms(strCallingForm).Visible = False Then
7830      Forms(strCallingForm).Visible = True
7840    End If
7850    Select Case ERR.Number
        Case 2585  ' ** This action can't be carried out while processing a form or report event.
          ' ** Ignore.
7860    Case Else
7870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7880    End Select
7890    Resume EXITP

End Sub
