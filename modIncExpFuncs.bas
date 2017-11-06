Attribute VB_Name = "modIncExpFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modIncExpFuncs"

'VGC 10/07/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   {In Parent}
' #######################################

Private lngMonitorCnt As Long, lngMonitorNum As Long
Private lngTpp As Long ', lngMonitors As Long ', blnIsOpen As Boolean
' **

Public Function IncomeExpense_BuildTable(datStartDate As Date, datEndDate As Date, strAccountNo As String, blnRecvException As Boolean, Optional varArchive As Variant) As Boolean
' ** Called by:
' **   frmRpt_IncomeExpense:
' **     DoReport()
' **   frmRpt_IncomeStatement:
' **     DoReport()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "IncomeExpense_BuildTable"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnArchive As Boolean
        Dim blnRetVal As Boolean

110     blnRetVal = True

120     Select Case IsMissing(varArchive)
        Case True
130       blnArchive = False
140     Case False
150       blnArchive = CBool(varArchive)
160     End Select

170     Set dbs = CurrentDb

180     If strAccountNo = "ALL" Then
190       Select Case blnArchive
          Case True
            ' ** Union of qryIncomeExpenseReports_02 (Ledger, linked to Account, by specified
            ' ** [datstart], [datend]; All), qryIncomeExpenseReports_02a (LedgerArchive,
            ' ** linked to Account, by specified [datstart], [datend]; All).
200         Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_02b")
210       Case False
            ' ** Ledger, linked to Account, by specified [datstart], [datend]; All.
220         Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_02")
230       End Select
240       With qdf.Parameters
250         ![datStart] = datStartDate
260         ![datEnd] = datEndDate
270       End With
280     Else
290       Select Case blnArchive
          Case True
            ' ** Union of qryIncomeExpenseReports_04 (Ledger, linked to Account, by specified
            ' ** [datstart], [datend], [actno]; Account), qryIncomeExpenseReports_04a (LedgerArchived,
            ' ** linked to Account, by specified [datstart], [datend], [actno]; Account).
300         Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_04b")
310       Case False
            ' ** Ledger, linked to Account, by specified [datstart], [datend], [actno]; Account.
320         Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_04")
330       End Select
340       With qdf.Parameters
350         ![actno] = strAccountNo
360         ![datStart] = datStartDate
370         ![datEnd] = datEndDate
380       End With
390     End If
400     Set rst = qdf.OpenRecordset()

410     If MakeTempTable(rst, "tmpIncomeExpenseReports") = True Then  ' ** Module Function: modFileUtilities.
420       rst.Close

430       If strAccountNo = "ALL" Then
440         Select Case blnArchive
            Case True
              ' ** Append qryIncomeExpenseReports_02b (Union of qryIncomeExpenseReports_02 (Ledger,
              ' ** linked to Account, by specified [datstart], [datend]; All), qryIncomeExpenseReports_02a
              ' ** (LedgerArchive, linked to Account, by specified [datstart], [datend]; All)) to tmpIncomeExpenseReports.
450           Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_03a")
460         Case False
              ' ** Append qryIncomeExpenseReports_02 (Ledger, linked to Account,
              ' ** by specified [datstart], [datend]; All) to tmpIncomeExpenseReports.
470           Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_03")
480         End Select
490         With qdf.Parameters
500           ![datStart] = datStartDate
510           ![datEnd] = datEndDate
520         End With
530       Else
540         Select Case blnArchive
            Case True
              ' ** Append qryIncomeExpenseReports_04b (Union of qryIncomeExpenseReports_04 (Ledger,
              ' ** linked to Account, by specified [datstart], [datend], [actno]; Account), qryIncomeExpenseReports_04a
              ' ** (LedgerArchived, linked to Account, by specified [datstart], [datend], [actno]; Account)) to tmpIncomeExpenseReports.
550           Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_05a")
560         Case False
              ' ** Append qryIncomeExpenseReports_04 (Ledger, linked to Account,
              ' ** by specified [datstart], [datend], [actno]; Account) to tmpIncomeExpenseReports.
570           Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_05")
580         End Select
590         With qdf.Parameters
600           ![actno] = strAccountNo
610           ![datStart] = datStartDate
620           ![datEnd] = datEndDate
630         End With
640       End If
650       qdf.Execute

660       Set rst = dbs.OpenRecordset("tmpIncomeExpenseReports")
670       If rst.BOF = True And rst.EOF = True Then
680         rst.Close
690         DoCmd.Hourglass False
700         MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
710         blnRetVal = False
            ' ** Reset the variables to ensure we get a new table anytime there is an error.
720         gdatStartDate = 0
730         gdatEndDate = 0
740         gstrAccountNo = 0
750         gstrAccountName = vbNullString
760       Else
770         rst.Close

780         Select Case IsLoaded("frmRpt_IncomeExpense", acForm)  ' ** Module Functions: modFileUtilities.
            Case True
790           With Forms("frmRpt_IncomeExpense")

800             Select Case .chkDontCombineMulti
                Case True
810               Select Case .chkSweepOnly
                  Case True
                    ' ** qryIncomeExpenseReports_01_Income_i (tmpIncomeExpenseReports, with 'Sold'); Multi-Lot Not Combined, Sweep Only.
820                 dbs.QueryDefs("qryIncomeExpenseReports_01_Income").SQL = dbs.QueryDefs("qryIncomeExpenseReports_01_Income_d").SQL
                    ' ** qryIncomeExpenseReports_01_Expense_i (tmpIncomeExpenseReports, with 'Sold'); Multi-Lot Not Combined, Sweep Only.
830                 dbs.QueryDefs("qryIncomeExpenseReports_01_Expense").SQL = dbs.QueryDefs("qryIncomeExpenseReports_01_Expense_d").SQL
840               Case False
                    ' ** qryIncomeExpenseReports_01_Income_i (tmpIncomeExpenseReports, with 'Sold'); Multi-Lot Not Combined.
850                 dbs.QueryDefs("qryIncomeExpenseReports_01_Income").SQL = dbs.QueryDefs("qryIncomeExpenseReports_01_Income_b").SQL
                    ' ** qryIncomeExpenseReports_01_Expense_i (tmpIncomeExpenseReports, with 'Sold'); Multi-Lot Not Combined.
860                 dbs.QueryDefs("qryIncomeExpenseReports_01_Expense").SQL = dbs.QueryDefs("qryIncomeExpenseReports_01_Expense_b").SQL
870               End Select
880             Case False
890               Select Case .chkSweepOnly
                  Case True
                    ' ** qryIncomeExpenseReports_01_Income_h (Union of qryIncomeExpenseReports_01_Income_e
                    ' ** (tmpIncomeExpenseReports, without 'Sold'), qryIncomeExpenseReports_01_Income_g
                    ' ** (qryIncomeExpenseReports_01_Income_f (tmpIncomeExpenseReports, just 'Sold'),
                    ' ** grouped and summed by multi-lot Sale)); Multi-Lot Combined, Sweep Only.
900                 dbs.QueryDefs("qryIncomeExpenseReports_01_Income").SQL = dbs.QueryDefs("qryIncomeExpenseReports_01_Income_c").SQL
                    ' ** qryIncomeExpenseReports_01_Expense_h (Union of qryIncomeExpenseReports_01_Expense_e
                    ' ** (tmpIncomeExpenseReports, without 'Sold'), qryIncomeExpenseReports_01_Expense_g
                    ' ** (qryIncomeExpenseReports_01_Expense_f (tmpIncomeExpenseReports, just 'Sold'),
                    ' ** grouped and summed by multi-lot Sale)); Multi-Lot Combined, Sweep Only.
910                 dbs.QueryDefs("qryIncomeExpenseReports_01_Expense").SQL = dbs.QueryDefs("qryIncomeExpenseReports_01_Expense_c").SQL
920               Case False
                    ' ** qryIncomeExpenseReports_01_Income_h (Union of qryIncomeExpenseReports_01_Income_e
                    ' ** (tmpIncomeExpenseReports, without 'Sold'), qryIncomeExpenseReports_01_Income_g
                    ' ** (qryIncomeExpenseReports_01_Income_f (tmpIncomeExpenseReports, just 'Sold'),
                    ' ** grouped and summed by multi-lot Sale)); Multi-Lot Combined.
930                 dbs.QueryDefs("qryIncomeExpenseReports_01_Income").SQL = dbs.QueryDefs("qryIncomeExpenseReports_01_Income_a").SQL
                    ' ** qryIncomeExpenseReports_01_Expense_h (Union of qryIncomeExpenseReports_01_Expense_e
                    ' ** (tmpIncomeExpenseReports, without 'Sold'), qryIncomeExpenseReports_01_Expense_g
                    ' ** (qryIncomeExpenseReports_01_Expense_f (tmpIncomeExpenseReports, just 'Sold'),
                    ' ** grouped and summed by multi-lot Sale)); Multi-Lot Combined.
940                 dbs.QueryDefs("qryIncomeExpenseReports_01_Expense").SQL = dbs.QueryDefs("qryIncomeExpenseReports_01_Expense_a").SQL
950               End Select
960             End Select

970             Select Case .opgZeroCash
                Case .opgZeroCash_optInclude.OptionValue
                  ' ** Skip the query.
980             Case .opgZeroCash_optExclude.OptionValue
                  ' ** Delete records when icash and pcash equal zero.
990               Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_06")
1000              qdf.Execute
1010            End Select
                ' ** Report filter does:
                ' **   .opgUnspecified
                ' **   .opgPrincipalCash
1020          End With

1030        Case False

              ' ** qryIncomeExpenseReports_01_Income_h (Union of qryIncomeExpenseReports_01_Income_e
              ' ** (tmpIncomeExpenseReports, without 'Sold'), qryIncomeExpenseReports_01_Income_g
              ' ** (qryIncomeExpenseReports_01_Income_f (tmpIncomeExpenseReports, just 'Sold'),
              ' ** grouped and summed by multi-lot Sale)); Multi-Lot Combined.
1040          dbs.QueryDefs("qryIncomeExpenseReports_01_Income").SQL = dbs.QueryDefs("qryIncomeExpenseReports_01_Income_a").SQL
              ' ** qryIncomeExpenseReports_01_Expense_h (Union of qryIncomeExpenseReports_01_Expense_e
              ' ** (tmpIncomeExpenseReports, without 'Sold'), qryIncomeExpenseReports_01_Expense_g
              ' ** (qryIncomeExpenseReports_01_Expense_f (tmpIncomeExpenseReports, just 'Sold'),
              ' ** grouped and summed by multi-lot Sale)); Multi-Lot Combined.
1050          dbs.QueryDefs("qryIncomeExpenseReports_01_Expense").SQL = dbs.QueryDefs("qryIncomeExpenseReports_01_Expense_a").SQL
              ' ** Delete records when icash and pcash equal zero.
1060          Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_06")
1070          qdf.Execute

1080          Select Case blnRecvException
              Case True  ' ** True from Income Statement.
                ' ** Delete records, where jcomment contains '*long term capital gain*', and journaltype <>'Received'.
                ' ** VGC 02/22/2009: Exception, per Rich.
1090            Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_08")
1100            qdf.Execute
1110          Case False  ' ** False from Income/Expense Report.
                ' ** Delete records, where jcomment contains '*long term capital gain*'.
                ' ** VGC 02/24/2009: Exception removed, per Rich.
                'Set qdf = dbs.QueryDefs("qryIncomeExpenseReports_07")
                'qdf.Execute
1120          End Select

1130        End Select

1140      End If
1150    Else
1160      rst.Close
1170      DoCmd.Hourglass False
1180      MsgBox "Unable to create temporary table for reporting.", vbCritical + vbOKOnly, "Error Creating Temporary Table"
1190      blnRetVal = False
          ' ** Reset the variables to ensure we get a new table anytime there is an error.
1200      gdatStartDate = 0
1210      gdatEndDate = 0
1220      gstrAccountNo = 0
1230      gstrAccountName = vbNullString
1240    End If
1250    dbs.Close

EXITP:
1260    Set rst = Nothing
1270    Set qdf = Nothing
1280    Set dbs = Nothing
1290    IncomeExpense_BuildTable = blnRetVal
1300    Exit Function

ERRH:
1310    blnRetVal = False
1320    DoCmd.Hourglass False
1330    Select Case ERR.Number
        Case 3000  ' ** Reserved. There is no message for this error.
1340      Beep
1350      MsgBox "Trust Accountant is unable to complete your request." & vbCrLf & _
            "Exit the program, then try again.", vbCritical + vbOKOnly, "Error 3000"
1360    Case Else
1370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1380    End Select
1390    Resume EXITP

End Function

Public Sub ShowOptions_IE(strMode As String, lngDetail_Height As Long, lngDiff As Long, lngIncBtn_Top As Long, lngExpBtn_Top As Long, lngBtnLbl_Diff As Long, lngDetailLine1_Top As Long, lngAccountNo_Top As Long, lngAccountNoBox_Top As Long, lngAccountNoLbl_Top As Long, lngRemMe_Top As Long, lngAcctSrc_Top As Long, lngAcctSrcOpt_Top As Long, lngAcctSrcOptLbl_Top As Long, lngPageOf_Top As Long, lngAllBtn_Top As Long, lngHighlight_Offset As Long, lngDetailLine3_Top As Long, frm As Access.Form)
' **
' ** ShowOptions_IE(
' **   strMode As String, lngDetail_Height As Long, lngDiff As Long, lngIncBtn_Top As Long,
' **   lngExpBtn_Top As Long, lngBtnLbl_Diff As Long, lngDetailLine1_Top As Long,
' **   lngAccountNo_Top As Long, lngAccountNoBox_Top As Long, lngAccountNoLbl_Top As Long,
' **   lngRemMe_Top As Long, lngAcctSrc_Top As Long, lngAcctSrcOpt_Top As Long,
' **   lngAcctSrcOptLbl_Top As Long, lngPageOf_Top As Long, lngAllBtn_Top As Long,
' **   lngHighlight_Offset As Long, lngDetailLine3_Top As Long, frm As Access.Form
' ** )

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowOptions_IE"

        Dim lngFrm_Top As Long, lngFrm_Left As Long, lngFrm_Width As Long, lngFrm_Height As Long
        Dim lngTmp01 As Long

1410    With frm

1420      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
1430        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1440      End If

1450      lngFrm_Top = .frm_top
1460      lngFrm_Left = .frm_left
1470      lngFrm_Width = .frm_width
1480      lngFrm_Height = .frm_height

1490      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
1500      lngMonitorNum = 1&: lngTmp01 = 0&
1510      EnumMonitors frm  ' ** Module Function: modMonitorFuncs.
1520      If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

1530      Select Case strMode
          Case "Hide"
            ' ** Hide them.
1540        .opgUnspecified.Enabled = False
1550        .opgUnspecified.Visible = False
1560        .opgUnspecified_vline01.Visible = False
1570        .opgUnspecified_vline02.Visible = False
1580        .opgUnspecified_box.Visible = False
1590        .opgPrincipalCash.Enabled = False
1600        .opgPrincipalCash.Visible = False
1610        .opgPrincipalCash_vline01.Visible = False
1620        .opgPrincipalCash_vline02.Visible = False
1630        .opgPrincipalCash_box.Visible = False
1640        .opgZeroCash.Enabled = False
1650        .opgZeroCash.Visible = False
1660        .opgZeroCash_vline01.Visible = False
1670        .opgZeroCash_vline02.Visible = False
1680        .opgZeroCash_box.Visible = False
1690        .opgSummary.Enabled = False
1700        .opgSummary.Visible = False
1710        .opgSummary_vline01.Visible = False
1720        .opgSummary_vline02.Visible = False
1730        .opgSummary_box.Visible = False
1740        .chkDontCombineMulti.Enabled = False
1750        .chkDontCombineMulti.Visible = False
1760        .chkAcctEveryLine.Enabled = False
1770        .chkAcctEveryLine.Visible = False
1780        .chkSweepOnly.Enabled = False
1790        .chkSweepOnly.Visible = False
1800        .opgOptionGroups_vline05.Visible = False
1810        .opgOptionGroups_vline06.Visible = False
1820        .cmdReset.Enabled = False
1830        .cmdReset.Visible = False
1840        .cmdReset_raised_img.Visible = False
1850        .cmdReset_raised_semifocus_dots_img.Visible = False
1860        .cmdReset_raised_focus_img.Visible = False
1870        .cmdReset_raised_focus_dots_img.Visible = False
1880        .cmdReset_sunken_focus_dots_img.Visible = False
1890        .cmdReset_raised_img_dis.Visible = False
1900        .opgOptionGroups_box.Visible = False
1910        .opgOptionGroups_box2.Visible = False
1920        .opgOptionGroups_vline01.Visible = False
1930        .opgOptionGroups_vline02.Visible = False
1940        .opgOptionGroups_vline03.Visible = False
1950        .opgOptionGroups_vline04.Visible = False
1960        .opgOptionGroups_hline01.Visible = False
1970        .opgOptionGroups_hline02.Visible = False
1980        .opgOptionGroups_box3.Visible = True
            ' ** Then move everything else up.
1990        .cmbAccounts_box.Top = (.cmbAccounts_box.Top - lngDiff)
2000        .cmbAccounts_lbl.Top = (.cmbAccounts_lbl.Top - lngDiff)
2010        .cmbAccounts.Top = (.cmbAccounts.Top - lngDiff)
2020        .chkRememberMe.Top = (.chkRememberMe.Top - lngDiff)
2030        .chkRememberMe_lbl.Top = (.chkRememberMe_lbl.Top - lngDiff)
2040        .chkRememberMe_lbl2_dim.Top = (.chkRememberMe_lbl2_dim.Top - lngDiff)
2050        .chkRememberMe_lbl2_dim_hi.Top = (.chkRememberMe_lbl2_dim_hi.Top - lngDiff)
2060        .opgAccountSource.Top = (.opgAccountSource.Top - lngDiff)
2070        .opgAccountSource_box.Top = (.opgAccountSource_box.Top - lngDiff)
2080        .opgAccountSource_optNumber.Top = (.opgAccountSource_optNumber.Top - lngDiff)
2090        .opgAccountSource_optNumber_lbl.Top = (.opgAccountSource_optNumber_lbl.Top - lngDiff)
2100        .opgAccountSource_optNumber_lbl2.Top = (.opgAccountSource_optNumber_lbl2.Top - lngDiff)
2110        .opgAccountSource_optNumber_lbl2_dim_hi.Top = (.opgAccountSource_optNumber_lbl2_dim_hi.Top - lngDiff)
2120        .opgAccountSource_optName.Top = (.opgAccountSource_optName.Top - lngDiff)
2130        .opgAccountSource_optName_lbl.Top = (.opgAccountSource_optName_lbl.Top - lngDiff)
2140        .opgAccountSource_optName_lbl2.Top = (.opgAccountSource_optName_lbl2.Top - lngDiff)
2150        .opgAccountSource_optName_lbl2_dim_hi.Top = (.opgAccountSource_optName_lbl2_dim_hi.Top - lngDiff)
2160        .opgAccountSource.Height = .opgAccountSource_box.Height
2170        .chkPageOf.Top = (.chkPageOf.Top - lngDiff)
2180        .chkPageOf_lbl.Top = (.chkPageOf_lbl.Top - lngDiff)
2190        .chkPageOf_box.Top = (.chkPageOf_box.Top - lngDiff)
2200        .Detail_vline01.Top = (.Detail_vline01.Top - lngDiff)
2210        .Detail_vline02.Top = (.Detail_vline02.Top - lngDiff)
2220        .Detail_hline01.Top = (.Detail_hline01.Top - lngDiff)
2230        .Detail_hline02.Top = (.Detail_hline02.Top - lngDiff)
2240        .cmdRevIncExp_Income_lbl.Top = (.cmdRevIncExp_Income_lbl.Top - lngDiff)
2250        .cmdRevIncExp_Income_box.Top = (.cmdRevIncExp_Income_box.Top - lngDiff)
2260        .cmdRevIncExp_Expense_lbl.Top = (.cmdRevIncExp_Expense_lbl.Top - lngDiff)
2270        .cmdRevIncExp_Expense_box.Top = (.cmdRevIncExp_Expense_box.Top - lngDiff)
2280        .cmdRevIncExp_IncomePreview.Top = (.cmdRevIncExp_IncomePreview.Top - lngDiff)
2290        .cmdRevIncExp_IncomePrint.Top = (.cmdRevIncExp_IncomePrint.Top - lngDiff)
2300        .cmdRevIncExp_IncomeWord.Top = (.cmdRevIncExp_IncomeWord.Top - lngDiff)
2310        .cmdRevIncExp_IncomeExcel.Top = (.cmdRevIncExp_IncomeExcel.Top - lngDiff)
2320        .cmdRevIncExp_ExpensePreview.Top = (.cmdRevIncExp_ExpensePreview.Top - lngDiff)
2330        .cmdRevIncExp_ExpensePrint.Top = (.cmdRevIncExp_ExpensePrint.Top - lngDiff)
2340        .cmdRevIncExp_ExpenseWord.Top = (.cmdRevIncExp_ExpenseWord.Top - lngDiff)
2350        .cmdRevIncExp_ExpenseExcel.Top = (.cmdRevIncExp_ExpenseExcel.Top - lngDiff)
2360        .Detail_hline03.Top = (.Detail_hline03.Top - lngDiff)
2370        .Detail_hline04.Top = (.Detail_hline04.Top - lngDiff)
2380        .Detail_vline03.Top = (.Detail_vline03.Top - lngDiff)
2390        .Detail_vline04.Top = (.Detail_vline04.Top - lngDiff)
2400        .cmdPrintAll.Top = (.cmdPrintAll.Top - lngDiff)
2410        .cmdWordAll.Top = (.cmdWordAll.Top - lngDiff)
2420        .cmdExcelAll.Top = (.cmdExcelAll.Top - lngDiff)
2430        .cmdPrintAll_box01.Top = (.cmdPrintAll.Top - lngHighlight_Offset)
2440        .cmdPrintAll_box02.Top = (.cmdRevIncExp_IncomePrint.Top - lngHighlight_Offset)
2450        .cmdWordAll_box01.Top = (.cmdWordAll.Top - lngHighlight_Offset)
2460        .cmdWordAll_box02.Top = (.cmdRevIncExp_IncomeWord.Top - lngHighlight_Offset)
2470        .cmdExcelAll_box01.Top = (.cmdExcelAll.Top - lngHighlight_Offset)
2480        .cmdExcelAll_box02.Top = (.cmdRevIncExp_IncomeExcel.Top - lngHighlight_Offset)
            '.cmdPrintAll_box01.Top = (.cmdPrintAll_box01.Top - lngDiff)
            '.cmdPrintAll_box02.Top = (.cmdPrintAll_box02.Top - lngDiff)
            '.cmdWordAll_box01.Top = (.cmdWordAll_box01.Top - lngDiff)
            '.cmdWordAll_box02.Top = (.cmdWordAll_box02.Top - lngDiff)
            '.cmdExcelAll_box01.Top = (.cmdExcelAll_box01.Top - lngDiff)
            '.cmdExcelAll_box02.Top = (.cmdExcelAll_box02.Top - lngDiff)
2490        .GoToReport_arw_rptinc_img.Top = .GoToReport_arw_rptinc_img.Top - lngDiff
2500        .GoToReport_arw_rptexp_img.Top = .GoToReport_arw_rptexp_img.Top - lngDiff
            ' ** And finally, change the button image.
2510        .cmdMoreOptions_R_raised_img.Visible = True
2520        .cmdMoreOptions_R_raised_semifocus_dots_img.Visible = False
2530        .cmdMoreOptions_R_raised_focus_img.Visible = False
2540        .cmdMoreOptions_R_raised_focus_dots_img.Visible = False
2550        .cmdMoreOptions_R_sunken_focus_dots_img.Visible = False
2560        .cmdMoreOptions_R_raised_img_dis.Visible = False
2570        .cmdMoreOptions_L_raised_img.Visible = False
2580        .cmdMoreOptions_L_raised_semifocus_dots_img.Visible = False
2590        .cmdMoreOptions_L_raised_focus_img.Visible = False
2600        .cmdMoreOptions_L_raised_focus_dots_img.Visible = False
2610        .cmdMoreOptions_L_sunken_focus_dots_img.Visible = False
2620        .cmdMoreOptions_L_raised_img_dis.Visible = False
2630        .Detail.Height = (lngDetail_Height - lngDiff)
2640        DoCmd.SelectObject acForm, .Name, False
2650        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
2660        DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, (lngFrm_Height - lngDiff)  'lngFrm_Top
2670        If lngMonitorNum > 1& Then
2680          LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
2690        End If
2700        .chkOptionsOpen = False
2710      Case "Show"
            ' ** First move everything back down.
2720        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
2730        DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, lngFrm_Height  'lngFrm_Top
2740        If lngMonitorNum > 1& Then
2750          LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
2760        End If
2770        .Detail.Height = lngDetail_Height
2780        .cmdPrintAll.Top = lngAllBtn_Top
2790        .cmdWordAll.Top = lngAllBtn_Top
2800        .cmdExcelAll.Top = lngAllBtn_Top
2810        .Detail_hline03.Top = lngDetailLine3_Top
2820        .Detail_hline04.Top = (lngDetailLine3_Top + lngTpp)
2830        .Detail_vline03.Top = lngDetailLine3_Top
2840        .Detail_vline04.Top = lngDetailLine3_Top
2850        .cmdRevIncExp_Income_lbl.Top = (lngIncBtn_Top + lngBtnLbl_Diff)
2860        .cmdRevIncExp_Income_box.Top = (lngIncBtn_Top - (3& * lngTpp))
2870        .cmdRevIncExp_IncomePreview.Top = lngIncBtn_Top
2880        .cmdRevIncExp_IncomePrint.Top = lngIncBtn_Top
2890        .cmdRevIncExp_IncomeWord.Top = lngIncBtn_Top
2900        .cmdRevIncExp_IncomeExcel.Top = lngIncBtn_Top
2910        .cmdRevIncExp_Expense_lbl.Top = (lngExpBtn_Top + lngBtnLbl_Diff)
2920        .cmdRevIncExp_Expense_box.Top = (lngExpBtn_Top - (3& * lngTpp))
2930        .cmdRevIncExp_ExpensePreview.Top = lngExpBtn_Top
2940        .cmdRevIncExp_ExpensePrint.Top = lngExpBtn_Top
2950        .cmdRevIncExp_ExpenseWord.Top = lngExpBtn_Top
2960        .cmdRevIncExp_ExpenseExcel.Top = lngExpBtn_Top
2970        .cmdPrintAll_box01.Top = (.cmdPrintAll.Top - lngHighlight_Offset)
2980        .cmdPrintAll_box02.Top = (.cmdRevIncExp_IncomePrint.Top - lngHighlight_Offset)
2990        .cmdWordAll_box01.Top = (.cmdWordAll.Top - lngHighlight_Offset)
3000        .cmdWordAll_box02.Top = (.cmdRevIncExp_IncomeWord.Top - lngHighlight_Offset)
3010        .cmdExcelAll_box01.Top = (.cmdExcelAll.Top - lngHighlight_Offset)
3020        .cmdExcelAll_box02.Top = (.cmdRevIncExp_IncomeExcel.Top - lngHighlight_Offset)
3030        .Detail_vline01.Top = lngDetailLine1_Top
3040        .Detail_vline02.Top = lngDetailLine1_Top
3050        .Detail_hline01.Top = lngDetailLine1_Top
3060        .Detail_hline02.Top = (lngDetailLine1_Top + lngTpp)
3070        .cmbAccounts_box.Top = lngAccountNoBox_Top
3080        .cmbAccounts_lbl.Top = lngAccountNoLbl_Top
3090        .cmbAccounts.Top = lngAccountNo_Top
3100        .chkRememberMe.Top = lngRemMe_Top
3110        .chkRememberMe_lbl.Top = lngRemMe_Top
3120        .chkRememberMe_lbl2_dim.Top = lngRemMe_Top
3130        .chkRememberMe_lbl2_dim_hi.Top = (lngRemMe_Top + lngTpp)
3140        .opgAccountSource_box.Top = lngAcctSrc_Top
3150        Do Until .opgAccountSource.Top >= lngAcctSrc_Top
3160          .opgAccountSource_optNumber.Top = (.opgAccountSource_optNumber.Top + lngTpp)
3170          .opgAccountSource_optNumber_lbl.Top = (.opgAccountSource_optNumber_lbl.Top + lngTpp)
3180          .opgAccountSource_optNumber_lbl2.Top = (.opgAccountSource_optNumber_lbl2.Top + lngTpp)
3190          .opgAccountSource_optNumber_lbl2_dim_hi.Top = (.opgAccountSource_optNumber_lbl2_dim_hi.Top + lngTpp)
3200          .opgAccountSource_optName.Top = (.opgAccountSource_optName.Top + lngTpp)
3210          .opgAccountSource_optName_lbl.Top = (.opgAccountSource_optName_lbl.Top + lngTpp)
3220          .opgAccountSource_optName_lbl2.Top = (.opgAccountSource_optName_lbl2.Top + lngTpp)
3230          .opgAccountSource_optName_lbl2_dim_hi.Top = (.opgAccountSource_optName_lbl2_dim_hi.Top + lngTpp)
3240          .opgAccountSource.Top = (.opgAccountSource.Top + lngTpp)
3250        Loop
3260        .opgAccountSource.Top = lngAcctSrc_Top
3270        .opgAccountSource_optNumber.Top = lngAcctSrcOpt_Top
3280        .opgAccountSource_optNumber_lbl.Top = lngAcctSrcOptLbl_Top
3290        .opgAccountSource_optNumber_lbl2.Top = lngAcctSrcOptLbl_Top
3300        .opgAccountSource_optNumber_lbl2_dim_hi.Top = (lngAcctSrcOptLbl_Top + lngTpp)
3310        .opgAccountSource_optName.Top = lngAcctSrcOpt_Top
3320        .opgAccountSource_optName_lbl.Top = lngAcctSrcOptLbl_Top
3330        .opgAccountSource_optName_lbl2.Top = lngAcctSrcOptLbl_Top
3340        .opgAccountSource_optName_lbl2_dim_hi.Top = (lngAcctSrcOptLbl_Top + lngTpp)
3350        .opgAccountSource.Height = .opgAccountSource_box.Height
3360        .chkPageOf.Top = lngPageOf_Top
3370        .chkPageOf_lbl.Top = (lngPageOf_Top - lngTpp)
3380        .chkPageOf_box.Top = (lngPageOf_Top - (3& * lngTpp))
3390        .GoToReport_arw_rptinc_img.Top = (lngIncBtn_Top + lngTpp)
3400        .GoToReport_arw_rptexp_img.Top = (lngExpBtn_Top + lngTpp)
            ' ** Then show them.
3410        .opgOptionGroups_box3.Visible = False
3420        .opgOptionGroups_box.Visible = True
3430        .opgOptionGroups_box2.Visible = True
3440        .opgOptionGroups_vline01.Visible = True
3450        .opgOptionGroups_vline02.Visible = True
3460        .opgOptionGroups_vline03.Visible = True
3470        .opgOptionGroups_vline04.Visible = True
3480        .opgOptionGroups_hline01.Visible = True
3490        .opgOptionGroups_hline02.Visible = True
3500        .opgOptionGroups_vline05.Visible = True
3510        .opgOptionGroups_vline06.Visible = True
3520        .opgUnspecified_box.Visible = True
3530        .opgUnspecified_vline01.Visible = True
3540        .opgUnspecified_vline02.Visible = True
3550        .opgUnspecified.Visible = True
3560        .opgUnspecified.Enabled = True
3570        .opgPrincipalCash_box.Visible = True
3580        .opgPrincipalCash_vline01.Visible = True
3590        .opgPrincipalCash_vline02.Visible = True
3600        .opgPrincipalCash.Visible = True
3610        .opgPrincipalCash.Enabled = True
3620        .opgZeroCash_box.Visible = True
3630        .opgZeroCash_vline01.Visible = True
3640        .opgZeroCash_vline02.Visible = True
3650        .opgZeroCash.Visible = True
3660        .opgZeroCash.Enabled = True
3670        .opgSummary_box.Visible = True
3680        .opgSummary_vline01.Visible = True
3690        .opgSummary_vline02.Visible = True
3700        .opgSummary.Visible = True
3710        .opgSummary.Enabled = True
3720        .chkDontCombineMulti.Visible = True
3730        .chkDontCombineMulti.Enabled = True
3740        .chkAcctEveryLine.Visible = True
3750        .chkAcctEveryLine.Enabled = True
3760        .chkSweepOnly.Visible = True
3770        .chkSweepOnly.Enabled = True
3780        .cmdReset.Visible = True
3790        .cmdReset_raised_img.Visible = True
3800        .cmdReset_raised_semifocus_dots_img.Visible = False
3810        .cmdReset_raised_focus_img.Visible = False
3820        .cmdReset_raised_focus_dots_img.Visible = False
3830        .cmdReset_sunken_focus_dots_img.Visible = False
3840        .cmdReset_raised_img_dis.Visible = False
3850        .cmdReset.Enabled = True
            ' ** And change the button image.
3860        .cmdMoreOptions_L_raised_img.Visible = True
3870        .cmdMoreOptions_L_raised_semifocus_dots_img.Visible = False
3880        .cmdMoreOptions_L_raised_focus_img.Visible = False
3890        .cmdMoreOptions_L_raised_focus_dots_img.Visible = False
3900        .cmdMoreOptions_L_sunken_focus_dots_img.Visible = False
3910        .cmdMoreOptions_L_raised_img_dis.Visible = False
3920        .cmdMoreOptions_R_raised_img.Visible = False
3930        .cmdMoreOptions_R_raised_semifocus_dots_img.Visible = False
3940        .cmdMoreOptions_R_raised_focus_img.Visible = False
3950        .cmdMoreOptions_R_raised_focus_dots_img.Visible = False
3960        .cmdMoreOptions_R_sunken_focus_dots_img.Visible = False
3970        .cmdMoreOptions_R_raised_img_dis.Visible = False
3980        .chkOptionsOpen = True
3990      End Select

4000    End With

EXITP:
4010    Exit Sub

ERRH:
4020    Select Case ERR.Number
        Case Else
4030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4040    End Select
4050    Resume EXITP

End Sub

Public Sub OptionsChk_IE(frm As Access.Form)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "OptionsChk_IE"

        Dim blnFound As Boolean

4110    blnFound = False

4120    With frm

4130      If .opgUnspecified <> Val(.opgUnspecified.DefaultValue) Then
4140        blnFound = True
4150      Else
4160        If .opgPrincipalCash <> Val(.opgPrincipalCash.DefaultValue) Then
4170          blnFound = True
4180        Else
4190          If .opgZeroCash <> Val(.opgZeroCash.DefaultValue) Then
4200            blnFound = True
4210          Else
4220            If .opgSummary <> Val(.opgSummary.DefaultValue) Then
4230              blnFound = True
4240            Else
4250              If .chkDontCombineMulti <> Val(.chkDontCombineMulti.DefaultValue) Then
4260                blnFound = True
4270              Else
4280                If .chkAcctEveryLine <> Val(.chkAcctEveryLine.DefaultValue) Then
4290                  blnFound = True
4300                Else
4310                  If .chkSweepOnly <> Val(.chkSweepOnly.DefaultValue) Then
4320                    blnFound = True
4330                  End If
4340                End If
4350              End If
4360            End If
4370          End If
4380        End If
4390      End If

4400      Select Case blnFound
          Case True  ' ** Everything's default.
4410        Select Case .opgOptionGroups_box.Visible
            Case True
4420          .cmdReset.Enabled = True
4430          .cmdReset_raised_img.Visible = True
4440          .cmdReset_raised_semifocus_dots_img.Visible = False
4450          .cmdReset_raised_focus_img.Visible = False
4460          .cmdReset_raised_focus_dots_img.Visible = False
4470          .cmdReset_sunken_focus_dots_img.Visible = False
4480          .cmdReset_raised_img_dis.Visible = False
4490          .opgOptionGroups_box4.Visible = False
4500        Case False
4510          .opgOptionGroups_box4.Visible = True
4520        End Select
4530      Case False  ' ** Some non-default options.
4540        .opgOptionGroups_box4.Visible = False
4550        If .opgOptionGroups_box.Visible = True Then
4560          .cmdReset_raised_img_dis.Visible = True
4570          .cmdReset_raised_img.Visible = False
4580          .cmdReset_raised_semifocus_dots_img.Visible = False
4590          .cmdReset_raised_focus_img.Visible = False
4600          .cmdReset_raised_focus_dots_img.Visible = False
4610          .cmdReset_sunken_focus_dots_img.Visible = False
4620  On Error Resume Next
4630          .cmdReset.Enabled = False
4640  On Error GoTo ERRH
4650        End If
4660      End Select

4670    End With

EXITP:
4680    Exit Sub

ERRH:
4690    Select Case ERR.Number
        Case Else
4700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4710    End Select
4720    Resume EXITP

End Sub

Public Function Report_Criteria_Msg_IE(frm As Access.Form) As String
' ** The gstrCrtRpt_Ordinal variable is borrowed from Court Reports.

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Criteria_Msg_IE"

        Dim strRetVal As String

4810    strRetVal = vbNullString

        ' ** chkSweepOnly handled in report's Form_Open().
        ' ** chkDontCombineMulti handled in report's Form_Open().

        ' ** Construct criteria message for report footer.
4820    With frm

4830      Select Case .opgUnspecified
          Case .opgUnspecified_optExclude.OptionValue
4840        Select Case .opgPrincipalCash
            Case .opgPrincipalCash_optExclude.OptionValue
4850          Select Case .opgZeroCash
              Case .opgZeroCash_optExclude.OptionValue
                ' ** Nothing.
4860          Case .opgZeroCash_optInclude.OptionValue
4870            strRetVal = "Includes Zero-Cash Entries"
4880          End Select
4890        Case .opgPrincipalCash_optInclude.OptionValue
4900          Select Case .opgZeroCash
              Case .opgZeroCash_optExclude.OptionValue
4910            strRetVal = "Includes Principal Cash Entries"
4920          Case .opgZeroCash_optInclude.OptionValue
4930            strRetVal = "Includes Principal Cash and Zero-Cash Entries"
4940          End Select
4950        End Select
4960      Case .opgUnspecified_optInclude.OptionValue
4970        Select Case .opgPrincipalCash
            Case .opgPrincipalCash_optExclude.OptionValue
4980          Select Case .opgZeroCash
              Case .opgZeroCash_optExclude.OptionValue
4990            strRetVal = "Includes Unspecified Entries"
5000          Case .opgZeroCash_optInclude.OptionValue
5010            strRetVal = "Includes Zero-Cash and Unspecified Entries"
5020          End Select
5030        Case .opgPrincipalCash_optInclude.OptionValue
5040          Select Case .opgZeroCash
              Case .opgZeroCash_optExclude.OptionValue
5050            strRetVal = "Includes Principal Cash and Unspecified Entries"
5060          Case .opgZeroCash_optInclude.OptionValue
5070            strRetVal = "Includes Principal Cash, Zero-Cash, and Unspecified Entries"
5080          End Select
5090        End Select
5100      Case .opgUnspecified_optOnly.OptionValue
5110        Select Case .opgPrincipalCash
            Case .opgPrincipalCash_optExclude.OptionValue
5120          Select Case .opgZeroCash
              Case .opgZeroCash_optExclude.OptionValue
5130            strRetVal = "Unspecified Entries Only"
5140          Case .opgZeroCash_optInclude.OptionValue
5150            strRetVal = "Unspecified Entries Only, including Zero-Cash Entries"
5160          End Select
5170        Case .opgPrincipalCash_optInclude.OptionValue
5180          Select Case .opgZeroCash
              Case .opgZeroCash_optExclude.OptionValue
5190            strRetVal = "Unspecified Entries Only, including Principal Cash Entries"
5200          Case .opgZeroCash_optInclude.OptionValue
5210            strRetVal = "Unspecified Entries Only, including Principal Cash and Zero-Cash Entries"
5220          End Select
5230        End Select
5240      End Select

5250    End With

EXITP:
5260    Report_Criteria_Msg_IE = strRetVal
5270    Exit Function

ERRH:
5280    strRetVal = RET_ERR
5290    Select Case ERR.Number
        Case Else
5300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5310    End Select
5320    Resume EXITP

End Function

Public Sub IncomeExcel_Click_IE(strFile1 As String, strFile2 As String, strFile3 As String, strFile4 As String, frm As Access.Form)
' ** For large exports, the OutputTo errors with:
' **   2306  There are too many rows to output, based on the limitation specified by the output format or by Microsoft Access.
' ** Plain means no detail.

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "IncomeExcel_Click_IE"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strQry1 As String, strQry2 As String, strRptCap1 As String, strRptCap2 As String
        Dim strRptPath1 As String, strRptPath2 As String, strRptPathFile1 As String, strRptPathFile2 As String
        Dim blnContinue As Boolean
        Dim msgResponse As VbMsgBoxResult

5410    With frm

5420      DoCmd.Hourglass True
5430      DoEvents

5440      blnContinue = True

5450      If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
5460        DoCmd.Hourglass False
5470        msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
5480        If msgResponse <> vbRetry Then
5490          blnContinue = False
5500        End If
5510      End If

5520      If blnContinue = True Then

5530        DoCmd.Hourglass True
5540        DoEvents

5550        If .DoReport = True Then  ' ** Form Function: frmRpt_IncomeExpense.

5560          blnContinue = True

5570          If IsNull(.UserReportPath) = True Then
5580            strRptPath1 = CurrentAppPath  ' ** Module Function: modFileUtilities.
5590          Else
5600            strRptPath1 = .UserReportPath
5610          End If
5620          strRptPath2 = strRptPath1

5630          Set dbs = CurrentDb

              ' ** NOTE: chkAcctEveryLine is handled via FormRef() within the queries.

              ' ** Group 1.
5640          Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
5650            Select Case .chkDetail
                Case True
5660              Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
5670                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
5680                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_acct").SQL
5690                Case .opgPrincipalCash_optExclude.OptionValue
5700                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_acct_nopc").SQL
5710                End Select
5720              Case .opgUnspecified_optExclude.OptionValue
5730                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
5740                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_acct_un").SQL
5750                Case .opgPrincipalCash_optExclude.OptionValue
5760                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_acct_un_nopc").SQL
5770                End Select
5780              Case .opgUnspecified_optOnly.OptionValue
5790                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
5800                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_acct_uno").SQL
5810                Case .opgPrincipalCash_optExclude.OptionValue
5820                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_acct_uno_nopc").SQL
5830                End Select
5840              End Select
5850            Case False
5860              Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
5870                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
5880                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_acct").SQL
5890                Case .opgPrincipalCash_optExclude.OptionValue
5900                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_acct_nopc").SQL
5910                End Select
5920              Case .opgUnspecified_optExclude.OptionValue
5930                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
5940                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_acct_un").SQL
5950                Case .opgPrincipalCash_optExclude.OptionValue
5960                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_acct_un_nopc").SQL
5970                End Select
5980              Case .opgUnspecified_optOnly.OptionValue
5990                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6000                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_acct_uno").SQL
6010                Case .opgPrincipalCash_optExclude.OptionValue
6020                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_acct_uno_nopc").SQL
6030                End Select
6040              End Select
6050            End Select
6060          Case .opgAccountNumber_optAll.OptionValue
6070            Select Case .chkDetail
                Case True
6080              Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
6090                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6100                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_all").SQL
6110                Case .opgPrincipalCash_optExclude.OptionValue
6120                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_all_nopc").SQL
6130                End Select
6140              Case .opgUnspecified_optExclude.OptionValue
6150                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6160                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_all_un").SQL
6170                Case .opgPrincipalCash_optExclude.OptionValue
6180                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_all_un_nopc").SQL
6190                End Select
6200              Case .opgUnspecified_optOnly.OptionValue
6210                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6220                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_all_uno").SQL
6230                Case .opgPrincipalCash_optExclude.OptionValue
6240                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_all_uno_nopc").SQL
6250                End Select
6260              End Select
6270            Case False
6280              Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
6290                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6300                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_all").SQL
6310                Case .opgPrincipalCash_optExclude.OptionValue
6320                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_all_nopc").SQL
6330                End Select
6340              Case .opgUnspecified_optExclude.OptionValue
6350                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6360                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_all_un").SQL
6370                Case .opgPrincipalCash_optExclude.OptionValue
6380                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_all_un_nopc").SQL
6390                End Select
6400              Case .opgUnspecified_optOnly.OptionValue
6410                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6420                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_all_uno").SQL
6430                Case .opgPrincipalCash_optExclude.OptionValue
6440                  dbs.QueryDefs("qryIncomeExpenseReports_10").SQL = dbs.QueryDefs("qryIncomeExpenseReports_10_plain_all_uno_nopc").SQL
6450                End Select
6460              End Select
6470            End Select
6480          End Select  ' ** opgAccountNumber.

              ' ** NOTE: chkAcctEveryLine is handled via FormRef() within the queries.

              ' ** Group 2.
6490          Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
6500            Select Case .chkDetail
                Case True
6510              Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
6520                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6530                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_acct").SQL
6540                Case .opgPrincipalCash_optExclude.OptionValue
6550                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_acct_nopc").SQL
6560                End Select
6570              Case .opgUnspecified_optExclude.OptionValue
6580                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6590                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_acct_un").SQL
6600                Case .opgPrincipalCash_optExclude.OptionValue
6610                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_acct_un_nopc").SQL
6620                End Select
6630              Case .opgUnspecified_optOnly.OptionValue
6640                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6650                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_acct_uno").SQL
6660                Case .opgPrincipalCash_optExclude.OptionValue
6670                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_acct_uno_nopc").SQL
6680                End Select
6690              End Select
6700            Case False
6710              Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
6720                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6730                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_acct").SQL
6740                Case .opgPrincipalCash_optExclude.OptionValue
6750                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_acct_nopc").SQL
6760                End Select
6770              Case .opgUnspecified_optExclude.OptionValue
6780                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6790                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_acct_un").SQL
6800                Case .opgPrincipalCash_optExclude.OptionValue
6810                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_acct_un_nopc").SQL
6820                End Select
6830              Case .opgUnspecified_optOnly.OptionValue
6840                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6850                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_acct_uno").SQL
6860                Case .opgPrincipalCash_optExclude.OptionValue
6870                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_acct_uno_nopc").SQL
6880                End Select
6890              End Select
6900            End Select  ' ** chkDetail.
6910          Case .opgAccountNumber_optAll.OptionValue
6920            Select Case .chkDetail
                Case True
6930              Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
6940                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
6950                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_all").SQL
6960                Case .opgPrincipalCash_optExclude.OptionValue
6970                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_all_nopc").SQL
6980                End Select
6990              Case .opgUnspecified_optExclude.OptionValue
7000                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
7010                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_all_un").SQL
7020                Case .opgPrincipalCash_optExclude.OptionValue
7030                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_all_un_nopc").SQL
7040                End Select
7050              Case .opgUnspecified_optOnly.OptionValue
7060                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
7070                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_all_uno").SQL
7080                Case .opgPrincipalCash_optExclude.OptionValue
7090                  dbs.QueryDefs("qryIncomeExpenseReports_18").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_all_uno_nopc").SQL
7100                End Select
7110              End Select
7120            Case False
7130              Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
7140                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
7150                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_all").SQL
7160                Case .opgPrincipalCash_optExclude.OptionValue
7170                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_all_nopc").SQL
7180                End Select
7190              Case .opgUnspecified_optExclude.OptionValue
7200                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
7210                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_all_un").SQL
7220                Case .opgPrincipalCash_optExclude.OptionValue
7230                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_all_un_nopc").SQL
7240                End Select
7250              Case .opgUnspecified_optOnly.OptionValue
7260                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
7270                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_all_uno").SQL
7280                Case .opgPrincipalCash_optExclude.OptionValue
7290                  dbs.QueryDefs("qryIncomeExpenseReports_18_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_18_plain_all_uno_nopc").SQL
7300                End Select
7310              End Select
7320            End Select  ' ** chkDetail.
7330          End Select  ' ** opgAccountNumber.
7340          dbs.Close

              ' ** Options:
              ' **   chkDetail           : Below.
              ' **   chkDontCombineMulti : Handled in IncomeExpense_BuildTable().
              ' **   chkSweepOnly        : Handled in IncomeExpense_BuildTable().
              ' **   opgUnspecified      : Below.
              ' **   opgPrincipalCash    : Below.
              ' **   opgZeroCash         : Handled in IncomeExpense_BuildTable().
              ' **   opgSummary          : Below.

              ' ** Summary:
              ' **   chkDetail           : DOESN'T MATTER!
              ' **   chkDontCombineMulti : DOESN'T MATTER!
              ' **   chkSweepOnly        : DOES, BUT HANDLED ELSEWHERE!
              ' **   opgUnspecified      : DOES!
              ' **   opgPrincipalCash    : DOES!
              ' **   opgZeroCash         : DOES, BUT HANDLED ELSEWHERE!

              ' ** NOTE: chkAcctEveryLine is handled via FormRef() within the queries.

              ' ** Group 3.
7350          strQry2 = vbNullString: strRptPathFile2 = vbNullString: strRptCap2 = vbNullString
7360          Select Case .opgSummary
              Case .opgSummary_optOnly.OptionValue, .opgSummary_optInclude.OptionValue
7370            Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** Unavailable.
7380            Case .opgAccountNumber_optAll.OptionValue
7390              Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
7400                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
7410                  strQry2 = "qryIncomeExpenseReports_55_all"
7420                  strRptCap2 = "rptIncExp_Income_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
7430                Case .opgPrincipalCash_optExclude.OptionValue
7440                  strQry2 = "qryIncomeExpenseReports_55d_all"
7450                  strRptCap2 = "rptIncExp_Income_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
7460                End Select  ' ** opgPrincipalCash.
7470              Case .opgUnspecified_optExclude.OptionValue
7480                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
7490                  strQry2 = "qryIncomeExpenseReports_55c_all"
7500                  strRptCap2 = "rptIncExp_Income_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
7510                Case .opgPrincipalCash_optExclude.OptionValue
7520                  strQry2 = "qryIncomeExpenseReports_55e_all"
7530                  strRptCap2 = "rptIncExp_Income_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
7540                End Select  ' ** opgPrincipalCash.
7550              Case .opgUnspecified_optOnly.OptionValue
7560                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
7570                  strQry2 = "qryIncomeExpenseReports_55f_all"
7580                  strRptCap2 = "rptIncExp_Income_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
7590                Case .opgPrincipalCash_optExclude.OptionValue
7600                  strQry2 = "qryIncomeExpenseReports_55g_all"
7610                  strRptCap2 = "rptIncExp_Income_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
7620                End Select  ' ** opgPrincipalCash.
7630              End Select  ' ** opgUnspecified
7640            End Select  ' ** opgAccountNumber.
7650          Case .opgSummary_optExclude.OptionValue
                ' ** Nothing.
7660          End Select  ' ** opgSummary.

              ' ** NOTE: chkAcctEveryLine is handled via FormRef() within the queries.

              ' ** Group 4.
7670          strQry1 = vbNullString: strRptPathFile1 = vbNullString: strRptCap1 = vbNullString
7680          Select Case .chkDetail
              Case True
7690            Select Case .opgSummary  ' ** Unaffected by chkDetail.
                Case .opgSummary_optOnly.OptionValue
                  ' ** Variables filled above.
7700            Case Else
7710              strQry1 = "qryIncomeExpenseReports_23"
7720              strRptCap1 = "rptIncExp_Income_Detailed_"
7730              Select Case .opgSummary
                  Case .opgSummary_optInclude.OptionValue
7740                Select Case .opgAccountNumber
                    Case .opgAccountNumber_optSpecified.OptionValue
                      ' ** Unavailable.
7750                Case .opgAccountNumber_optAll.OptionValue
                      ' ** Variables filled above.
7760                  strRptCap1 = strRptCap1 & "All_"
7770                End Select  ' ** opgAccountNumber
7780              Case .opgSummary_optExclude.OptionValue
                    ' ** Nothing else.
7790                Select Case .opgAccountNumber
                    Case .opgAccountNumber_optSpecified.OptionValue
7800                  strRptCap1 = strRptCap1 & gstrAccountNo & "_"
7810                Case .opgAccountNumber_optAll.OptionValue
7820                  strRptCap1 = strRptCap1 & "All_"
7830                End Select
7840              Case .opgSummary_optOnly.OptionValue
                    ' ** Handled above.
7850              End Select  ' ** opgSummary.
7860              strRptCap1 = strRptCap1 & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
7870            End Select  ' ** opgSummary.
7880          Case False
7890            Select Case .opgSummary  ' ** Unaffected by chkDetail.
                Case .opgSummary_optOnly.OptionValue
7900              Select Case .opgAccountNumber
                  Case .opgAccountNumber_optSpecified.OptionValue
                    ' ** Unavailable.
7910              Case .opgAccountNumber_optAll.OptionValue
                    ' ** Variables filled above.
7920              End Select  ' ** opgAccountNumber.
7930            Case Else
7940              strQry1 = "qryIncomeExpenseReports_23_plain"
7950              strRptCap1 = "rptIncExp_Income_"
7960              Select Case .opgSummary
                  Case .opgSummary_optInclude.OptionValue
7970                Select Case .opgAccountNumber
                    Case .opgAccountNumber_optSpecified.OptionValue
                      ' ** Unavailable.
7980                Case .opgAccountNumber_optAll.OptionValue
                      ' ** Variables filled above.
7990                  strRptCap1 = strRptCap1 & "All_"
8000                End Select  ' ** opgAccountNumber.
8010              Case .opgSummary_optExclude.OptionValue
                    ' ** Nothing else.
8020                Select Case .opgAccountNumber
                    Case .opgAccountNumber_optSpecified.OptionValue
8030                  strRptCap1 = strRptCap1 & gstrAccountNo & "_"
8040                Case .opgAccountNumber_optAll.OptionValue
8050                  strRptCap1 = strRptCap1 & "All_"
8060                End Select
8070              Case .opgSummary_optOnly.OptionValue
                    ' ** Handled above.
8080              End Select  ' ** opgSummary.
8090              strRptCap1 = strRptCap1 & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
8100            End Select  ' ** opgSummary.
8110          End Select  ' ** chkDetail.

              ' ** Ask where to save the file.
8120          If strQry1 <> vbNullString And strRptCap1 <> vbNullString Then
8130            strRptPathFile1 = FileSaveDialog("xls", strRptCap1 & ".xls", strRptPath1, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
8140            If strRptPathFile1 = vbNullString Then
8150              blnContinue = False
8160            Else
8170              .UserReportPath = Parse_Path(strRptPathFile1)  ' ** Module Function: modFileUtilities.
8180            End If
8190          End If
8200          If blnContinue = True Then
8210            If strQry2 <> vbNullString And strRptCap2 <> vbNullString Then
8220              If strRptPathFile1 = vbNullString Then  ' ** Only ask if they didn't choose a standard report.
8230                strRptPathFile2 = FileSaveDialog("xls", strRptCap2 & ".xls", strRptPath2, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
8240                If strRptPathFile2 = vbNullString Then blnContinue = False
8250              Else  ' ** Use same path as standard report.
8260                strRptPathFile2 = Parse_Path(strRptPathFile1) & LNK_SEP & strRptCap2 & ".xls"  ' ** Module Function: modFileUtilities.
8270              End If
8280            End If
8290          End If  ' ** blnContinue.

8300          If blnContinue = True Then

8310            DoCmd.Hourglass True
8320            DoEvents

8330            If IsNull(.UserReportPath) = True Then
8340              If strRptPathFile1 <> vbNullString Then
8350                .UserReportPath = Parse_Path(strRptPathFile1)  ' ** Module Function: modFileUtilities.
8360              End If
8370            End If

                ' ** NOTE: chkAcctEveryLine is handled via FormRef() within the queries.

                'WAIT A MINUTE!
                'THESE 2 FILE VARS ARE NOT THE SAME
                'AS OUR FILE1, FILE2!
                'strRptPathFile1 = THE REPORT
                'strRptPathFile2 = SUMMARY
                ' ** Export 1.
8380            If strQry1 <> vbNullString And strRptPathFile1 <> vbNullString Then
8390              strFile1 = strRptPathFile1
8400  On Error Resume Next
8410              DoCmd.OutputTo acOutputQuery, strQry1, acFormatXLS, strRptPathFile1, False
8420              If ERR.Number <> 0 Then
8430                Select Case ERR.Number
                    Case 2306  ' ** There are too many rows to output, based on the limitation specified by the output format or by Microsoft Access.
8440  On Error GoTo ERRH
8450  On Error Resume Next
8460                  DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, strQry1, strRptPathFile1, True
8470                  If ERR.Number <> 0 Then
8480                    Select Case ERR.Number
                        Case 3000  ' ** Reserved error (-1038); there is no message for this error.
8490  On Error GoTo ERRH
                          ' ** After discovering the long Ledger description, this may no longer be necessary (or used).
8500                      IncomeExpense_Export strQry1, strRptPathFile1, .UserReportPath, "Income"  ' ** Module Function: modExcelFuncs.
                          'Debug.Print "'INCOME: 3RD TRY!"
8510                    Case Else
8520                      blnContinue = False
8530                      Set rst = dbs.OpenRecordset("tblErrorLog", dbOpenDynaset, dbConsistent)
8540                      zErrorWriteRecord ERR.Number, ERR.description, THIS_NAME, THIS_PROC, Erl, rst  ' ** Module Function: modErrorHandler.
8550                      rst.Close
8560                      Set rst = Nothing
8570                      Beep
8580                      DoCmd.Hourglass False
8590                      MsgBox "An error was detected while attempting to export the data to Excel." & vbCrLf & _
                            "  Error: " & CStr(ERR.Number) & vbCrLf & _
                            "  Description: " & ERR.description & vbCrLf & _
                            "Please contact Delta Data, Inc., for assistance.", vbInformation + vbOKOnly, "Error: " & CStr(ERR.Number)
8600  On Error GoTo ERRH
8610                    End Select
8620                  Else
8630  On Error GoTo ERRH
                        'Debug.Print "'INCOME: 2ND TRY!"
8640                  End If
8650                  If blnContinue = True Then
8660                    DoEvents
8670                    If Excel_IncExp(strRptPathFile1, "Income") = True Then  ' ** Module Function: modExcelFuncs.
8680                      DoEvents
8690                      Select Case .chkOpenExcel
                          Case True
                            ' ** Even though all references to the Excel_IncExp() objects are explicit,
                            ' ** and they're closed and quit, sometimes the process will not shut down.
                            ' ** (One suggestion is that the worksheet copy is the culprit.)
                            ' ** This, below, seems to be my only recourse.
8700                        If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
8710                          EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
8720                        End If
8730                        DoEvents
8740                        If strQry2 <> vbNullString And strRptPathFile2 <> vbNullString Then
                              ' ** Don't open it yet.
8750                        Else
8760                          OpenExe strRptPathFile1  ' ** Module Function: modShellFuncs.
8770                        End If
8780                      Case False
                            ' ** Nothing, that's it.
8790                      End Select
8800                    Else
8810                      blnContinue = False
8820                    End If
8830                  End If  ' ** blnContinue.
8840                Case Else
8850                  blnContinue = False
8860                  zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8870  On Error GoTo ERRH
8880                End Select  ' ** Err.Number.
8890              Else
8900  On Error GoTo ERRH
                    'Debug.Print "'INCOME: 1ST TRY!"
8910                If Excel_NameOnly(strRptPathFile1, "Income") = True Then  ' ** Module Function: modExcelFuncs.
8920                  Select Case .chkOpenExcel
                      Case True
8930                    If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
8940                      EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
8950                    End If
8960                    If strQry2 <> vbNullString And strRptPathFile2 <> vbNullString Then
                          ' ** Don't open it yet.
8970                    Else
8980                      OpenExe strRptPathFile1  ' ** Module Function: modShellFuncs.
8990                    End If
9000                  Case False
                        ' ** Nothing, that's it.
9010                  End Select
9020                Else
9030                  blnContinue = False
9040                End If
9050              End If
9060            End If  ' ** vbNullString.

9070          End If  ' ** blnContinue.

9080          If blnContinue = True Then

                ' ** Export 2.
9090            If strQry2 <> vbNullString And strRptPathFile2 <> vbNullString Then
9100              strFile3 = strRptPathFile2
9110  On Error Resume Next
9120              DoCmd.OutputTo acOutputQuery, strQry2, acFormatXLS, strRptPathFile2, False
9130              If ERR.Number <> 0 Then
9140                Select Case ERR.Number
                    Case 2306  ' ** There are too many rows to output, based on the limitation specified by the output format or by Microsoft Access.
9150  On Error GoTo ERRH
9160                  DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, strQry2, strRptPathFile2, True
9170                  DoEvents
9180                  Select Case .chkOpenExcel
                      Case True
9190                    OpenExe strRptPathFile2  ' ** Module Function: modShellFuncs.
9200                  Case False
                        ' ** Nothing, that's it.
9210                  End Select
9220                Case Else
9230                  blnContinue = False
9240                  zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9250  On Error GoTo ERRH
9260                End Select  ' ** Err.Number.
9270              Else
9280  On Error GoTo ERRH
9290                If Excel_NameOnly(strRptPathFile2, "Income Summary") = True Then  ' ** Module Function: modExcelFuncs.
9300                  Select Case .chkOpenExcel
                      Case True
9310                    If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
9320                      EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
9330                    End If
9340                    If strQry1 <> vbNullString And strRptPathFile1 <> vbNullString Then
9350                      OpenExe strRptPathFile1  ' ** Module Function: modShellFuncs.
9360                    End If
9370                    OpenExe strRptPathFile2  ' ** Module Function: modShellFuncs.
9380                  Case False
                        ' ** Nothing, that's it.
9390                  End Select
9400                Else
9410                  blnContinue = False
9420                End If
9430              End If
9440            End If  ' ** vbNullString.

9450          End If  ' ** blnContinue.

9460        End If  ' ** DoReport().
9470      End If  ' ** blnContinue.
9480    End With  ' ** Me.

9490    DoCmd.Hourglass False

EXITP:
9500    Set rst = Nothing
9510    Set dbs = Nothing
9520    Exit Sub

ERRH:
9530    DoCmd.Hourglass False
9540    Select Case ERR.Number
        Case Else
9550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9560    End Select
9570    Resume EXITP

End Sub

Public Sub ExpenseExcel_Click_IE(strFile1 As String, strFile2 As String, strFile3 As String, strFile4 As String, frm As Access.Form)
' ** For large exports, the OutputTo errors with:
' **   2306  There are too many rows to output, based on the limitation specified by the output format or by Microsoft Access.
' ** Plain means no detail.

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "ExpenseExcel_Click_IE"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strQry1 As String, strQry2 As String, strRptCap1 As String, strRptCap2 As String
        Dim strRptPath1 As String, strRptPath2 As String, strRptPathFile1 As String, strRptPathFile2 As String
        Dim blnContinue As Boolean
        Dim msgResponse As VbMsgBoxResult

9610    With frm

9620      DoCmd.Hourglass True
9630      DoEvents

9640      blnContinue = True

9650      If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
9660        DoCmd.Hourglass False
9670        msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
9680        If msgResponse <> vbRetry Then
9690          blnContinue = False
9700        End If
9710      End If

9720      If blnContinue = True Then

9730        DoCmd.Hourglass True
9740        DoEvents

9750        If .DoReport = True Then  ' ** Form Function: frmRpt_IncomeExpense.

9760          blnContinue = True

9770          If IsNull(.UserReportPath) = True Then
9780            strRptPath1 = CurrentAppPath  ' ** Module Function: modFileUtilities.
9790          Else
9800            strRptPath1 = .UserReportPath
9810          End If
9820          strRptPath2 = strRptPath1

9830          Set dbs = CurrentDb

              ' ** NOTE: chkAcctEveryLine is handled via FormRef() within the queries.

              ' ** Group 1.
9840          Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
9850            Select Case .chkDetail
                Case True
9860              Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
9870                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
9880                  dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_acct").SQL
9890                Case .opgPrincipalCash_optExclude.OptionValue
9900                  dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_acct_nopc").SQL
9910                End Select
9920              Case .opgUnspecified_optExclude.OptionValue
9930                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
9940                  dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_acct_un").SQL
9950                Case .opgPrincipalCash_optExclude.OptionValue
9960                  dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_acct_un_nopc").SQL
9970                End Select
9980              Case .opgUnspecified_optOnly.OptionValue
9990                Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10000                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_acct_uno").SQL
10010               Case .opgPrincipalCash_optExclude.OptionValue
10020                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_acct_uno_nopc").SQL
10030               End Select
10040             End Select
10050           Case False
10060             Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
10070               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10080                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_acct").SQL
10090               Case .opgPrincipalCash_optExclude.OptionValue
10100                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_acct_nopc").SQL
10110               End Select
10120             Case .opgUnspecified_optExclude.OptionValue
10130               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10140                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_acct_un").SQL
10150               Case .opgPrincipalCash_optExclude.OptionValue
10160                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_acct_un_nopc").SQL
10170               End Select
10180             Case .opgUnspecified_optOnly.OptionValue
10190               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10200                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_acct_uno").SQL
10210               Case .opgPrincipalCash_optExclude.OptionValue
10220                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_acct_uno_nopc").SQL
10230               End Select
10240             End Select
10250           End Select
10260         Case .opgAccountNumber_optAll.OptionValue
10270           Select Case .chkDetail
                Case True
10280             Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
10290               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10300                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_all").SQL
10310               Case .opgPrincipalCash_optExclude.OptionValue
10320                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_all_nopc").SQL
10330               End Select
10340             Case .opgUnspecified_optExclude.OptionValue
10350               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10360                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_all_un").SQL
10370               Case .opgPrincipalCash_optExclude.OptionValue
10380                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_all_un_nopc").SQL
10390               End Select
10400             Case .opgUnspecified_optOnly.OptionValue
10410               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10420                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_all_uno").SQL
10430               Case .opgPrincipalCash_optExclude.OptionValue
10440                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_all_uno_nopc").SQL
10450               End Select
10460             End Select
10470           Case False
10480             Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
10490               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10500                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_all").SQL
10510               Case .opgPrincipalCash_optExclude.OptionValue
10520                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_all_nopc").SQL
10530               End Select
10540             Case .opgUnspecified_optExclude.OptionValue
10550               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10560                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_all_un").SQL
10570               Case .opgPrincipalCash_optExclude.OptionValue
10580                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_all_un_nopc").SQL
10590               End Select
10600             Case .opgUnspecified_optOnly.OptionValue
10610               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10620                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_all_uno").SQL
10630               Case .opgPrincipalCash_optExclude.OptionValue
10640                 dbs.QueryDefs("qryIncomeExpenseReports_24").SQL = dbs.QueryDefs("qryIncomeExpenseReports_24_plain_all_uno_nopc").SQL
10650               End Select
10660             End Select
10670           End Select
10680         End Select  ' ** opgAccountNumber.

              ' ** NOTE: chkAcctEveryLine is handled via FormRef() within the queries.

              ' ** Group 2.
10690         Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
10700           Select Case .chkDetail
                Case True
10710             Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
10720               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10730                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_acct").SQL
10740               Case .opgPrincipalCash_optExclude.OptionValue
10750                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_acct_nopc").SQL
10760               End Select
10770             Case .opgUnspecified_optExclude.OptionValue
10780               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10790                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_acct_un").SQL
10800               Case .opgPrincipalCash_optExclude.OptionValue
10810                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_acct_un_nopc").SQL
10820               End Select
10830             Case .opgUnspecified_optOnly.OptionValue
10840               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10850                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_acct_uno").SQL
10860               Case .opgPrincipalCash_optExclude.OptionValue
10870                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_acct_uno_nopc").SQL
10880               End Select
10890             End Select
10900           Case False
10910             Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
10920               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10930                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_acct").SQL
10940               Case .opgPrincipalCash_optExclude.OptionValue
10950                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_acct_nopc").SQL
10960               End Select
10970             Case .opgUnspecified_optExclude.OptionValue
10980               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
10990                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_acct_un").SQL
11000               Case .opgPrincipalCash_optExclude.OptionValue
11010                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_acct_un_nopc").SQL
11020               End Select
11030             Case .opgUnspecified_optOnly.OptionValue
11040               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
11050                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_acct_uno").SQL
11060               Case .opgPrincipalCash_optExclude.OptionValue
11070                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_acct_uno_nopc").SQL
11080               End Select
11090             End Select
11100           End Select  ' ** chkDetail.
11110         Case .opgAccountNumber_optAll.OptionValue
11120           Select Case .chkDetail
                Case True
11130             Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
11140               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
11150                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_all").SQL
11160               Case .opgPrincipalCash_optExclude.OptionValue
11170                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_all_nopc").SQL
11180               End Select
11190             Case .opgUnspecified_optExclude.OptionValue
11200               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
11210                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_all_un").SQL
11220               Case .opgPrincipalCash_optExclude.OptionValue
11230                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_all_un_nopc").SQL
11240               End Select
11250             Case .opgUnspecified_optOnly.OptionValue
11260               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
11270                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_all_uno").SQL
11280               Case .opgPrincipalCash_optExclude.OptionValue
11290                 dbs.QueryDefs("qryIncomeExpenseReports_32").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_all_uno_nopc").SQL
11300               End Select
11310             End Select
11320           Case False
11330             Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
11340               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
11350                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_all").SQL
11360               Case .opgPrincipalCash_optExclude.OptionValue
11370                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_all_nopc").SQL
11380               End Select
11390             Case .opgUnspecified_optExclude.OptionValue
11400               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
11410                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_all_un").SQL
11420               Case .opgPrincipalCash_optExclude.OptionValue
11430                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_all_un_nopc").SQL
11440               End Select
11450             Case .opgUnspecified_optOnly.OptionValue
11460               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
11470                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_all_uno").SQL
11480               Case .opgPrincipalCash_optExclude.OptionValue
11490                 dbs.QueryDefs("qryIncomeExpenseReports_32_plain").SQL = dbs.QueryDefs("qryIncomeExpenseReports_32_plain_all_uno_nopc").SQL
11500               End Select
11510             End Select
11520           End Select  ' ** chkDetail.
11530         End Select  ' ** opgAccountNumber.
11540         dbs.Close

              ' ** NOTE: chkAcctEveryLine is handled via FormRef() within the queries.

              ' ** Group 3.
11550         strQry2 = vbNullString: strRptPathFile2 = vbNullString: strRptCap2 = vbNullString
11560         Select Case .opgSummary
              Case .opgSummary_optOnly.OptionValue, .opgSummary_optInclude.OptionValue
11570           Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** Unavailable.
11580           Case .opgAccountNumber_optAll.OptionValue
11590             Select Case .opgUnspecified
                  Case .opgUnspecified_optInclude.OptionValue
11600               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
11610                 strQry2 = "qryIncomeExpenseReports_65_all"
11620                 strRptCap2 = "rptIncExp_Expenses_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
11630               Case .opgPrincipalCash_optExclude.OptionValue
11640                 strQry2 = "qryIncomeExpenseReports_65d_all"
11650                 strRptCap2 = "rptIncExp_Expenses_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
11660               End Select  ' ** opgPrincipalCash.
11670             Case .opgUnspecified_optExclude.OptionValue
11680               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
11690                 strQry2 = "qryIncomeExpenseReports_65c_all"
11700                 strRptCap2 = "rptIncExp_Expenses_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
11710               Case .opgPrincipalCash_optExclude.OptionValue
11720                 strQry2 = "qryIncomeExpenseReports_65e_all"
11730                 strRptCap2 = "rptIncExp_Expenses_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
11740               End Select  ' ** opgPrincipalCash.
11750             Case .opgUnspecified_optOnly.OptionValue
11760               Select Case .opgPrincipalCash
                    Case .opgPrincipalCash_optInclude.OptionValue
11770                 strQry2 = "qryIncomeExpenseReports_65f_all"
11780                 strRptCap2 = "rptIncExp_Expenses_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
11790               Case .opgPrincipalCash_optExclude.OptionValue
11800                 strQry2 = "qryIncomeExpenseReports_65g_all"
11810                 strRptCap2 = "rptIncExp_Expenses_Summary_" & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
11820               End Select  ' ** opgPrincipalCash.
11830             End Select  ' ** opgUnspecified.
11840           End Select  ' ** opgAccountNumber.
11850         Case .opgSummary_optExclude.OptionValue
                ' ** Nothing.
11860         End Select  ' ** opgSummary.

              ' ** NOTE: chkAcctEveryLine is handled via FormRef() within the queries.

              ' ** Group 4.
11870         strQry1 = vbNullString: strRptPathFile1 = vbNullString: strRptCap1 = vbNullString
11880         Select Case .chkDetail
              Case True
11890           Select Case .opgSummary  ' ** Unaffected by chkDetail.
                Case .opgSummary_optOnly.OptionValue
                  ' ** Variables filled above.
11900           Case Else
11910             strQry1 = "qryIncomeExpenseReports_37"
11920             strRptCap1 = "rptIncExp_Expenses_Detailed_"
11930             Select Case .opgSummary
                  Case .opgSummary_optInclude.OptionValue
11940               Select Case .opgAccountNumber
                    Case .opgAccountNumber_optSpecified.OptionValue
                      ' ** Unavailable.
11950               Case .opgAccountNumber_optAll.OptionValue
                      ' ** Variables filled above.
11960                 strRptCap1 = strRptCap1 & "All_"
11970               End Select  ' ** opgAccountNumber
11980             Case .opgSummary_optExclude.OptionValue
                    ' ** Nothing else.
11990               Select Case .opgAccountNumber
                    Case .opgAccountNumber_optSpecified.OptionValue
12000                 strRptCap1 = strRptCap1 & gstrAccountNo & "_"
12010               Case .opgAccountNumber_optAll.OptionValue
12020                 strRptCap1 = strRptCap1 & "All_"
12030               End Select
12040             Case .opgSummary_optOnly.OptionValue
                    ' ** Handled above.
12050             End Select  ' ** opgSummary.
12060             strRptCap1 = strRptCap1 & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
12070           End Select  ' ** opgSummary.
12080         Case False
12090           Select Case .opgSummary  ' ** Unaffected by chkDetail.
                Case .opgSummary_optOnly.OptionValue
12100             Select Case .opgAccountNumber
                  Case .opgAccountNumber_optSpecified.OptionValue
                    ' ** Unavailable.
12110             Case .opgAccountNumber_optAll.OptionValue
                    ' ** Variables filled above.
12120             End Select  ' ** opgAccountNumber.
12130           Case Else
12140             strQry1 = "qryIncomeExpenseReports_37_plain"
12150             strRptCap1 = "rptIncExp_Expenses_"
12160             Select Case .opgSummary
                  Case .opgSummary_optInclude.OptionValue
12170               Select Case .opgAccountNumber
                    Case .opgAccountNumber_optSpecified.OptionValue
                      ' ** Unavailable.
12180               Case .opgAccountNumber_optAll.OptionValue
                      ' ** Variables filled above.
12190                 strRptCap1 = strRptCap1 & "All_"
12200               End Select  ' ** opgAccountNumber.
12210             Case .opgSummary_optExclude.OptionValue
                    ' ** Nothing else.
12220               Select Case .opgAccountNumber
                    Case .opgAccountNumber_optSpecified.OptionValue
12230                 strRptCap1 = strRptCap1 & gstrAccountNo & "_"
12240               Case .opgAccountNumber_optAll.OptionValue
12250                 strRptCap1 = strRptCap1 & "All_"
12260               End Select
12270             Case .opgSummary_optOnly.OptionValue
                    ' ** Handled above.
12280             End Select  ' ** opgSummary.
12290             strRptCap1 = strRptCap1 & Format(.DateStart, "yyyymmdd") & "_to_" & Format(.DateEnd, "yyyymmdd") 'Format(Date, "yyyymmdd")
12300           End Select  ' ** opgSummary.
12310         End Select  ' ** chkDetail.

              ' ** Ask where to save the file.
12320         If strQry1 <> vbNullString And strRptCap1 <> vbNullString Then
12330           strRptPathFile1 = FileSaveDialog("xls", strRptCap1 & ".xls", strRptPath1, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
12340           If strRptPathFile1 = vbNullString Then
12350             blnContinue = False
12360           Else
12370             .UserReportPath = Parse_Path(strRptPathFile1)  ' ** Module Function: modFileUtilities.
12380           End If
12390         End If
12400         If blnContinue = True Then
12410           If strQry2 <> vbNullString And strRptCap2 <> vbNullString Then
12420             If strRptPathFile1 = vbNullString Then  ' ** Only ask if they didn't choose a standard report.
12430               strRptPathFile2 = FileSaveDialog("xls", strRptCap2 & ".xls", strRptPath2, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
12440               If strRptPathFile2 = vbNullString Then blnContinue = False
12450             Else  ' ** Use same path as standard report.
12460               strRptPathFile2 = Parse_Path(strRptPathFile1) & LNK_SEP & strRptCap2 & ".xls"  ' ** Module Function: modFileUtilities.
12470             End If
12480           End If
12490         End If  ' ** blnContinue.

12500         If blnContinue = True Then

12510           DoCmd.Hourglass True
12520           DoEvents

12530           If IsNull(.UserReportPath) = True Then
12540             If strRptPathFile1 <> vbNullString Then
12550               .UserReportPath = Parse_Path(strRptPathFile1)  ' ** Module Function: modFileUtilities.
12560             End If
12570           End If

                ' ** NOTE: chkAcctEveryLine is handled via FormRef() within the queries.

                'WAIT A MINUTE!
                'THESE 2 FILE VARS ARE NOT THE SAME
                'AS OUR FILE1, FILE2!
                ' ** Export 1.
12580           If strQry1 <> vbNullString And strRptPathFile1 <> vbNullString Then
12590             strFile2 = strRptPathFile1
12600 On Error Resume Next
12610             DoCmd.OutputTo acOutputQuery, strQry1, acFormatXLS, strRptPathFile1, False
12620             If ERR.Number <> 0 Then
12630               Select Case ERR.Number
                    Case 2306  ' ** There are too many rows to output, based on the limitation specified by the output format or by Microsoft Access.
12640 On Error GoTo ERRH
12650 On Error Resume Next
12660                 DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, strQry1, strRptPathFile1, True
12670                 If ERR.Number <> 0 Then
12680                   Select Case ERR.Number
                        Case 3000  ' ** Reserved error (-1038); there is no message for this error.
12690 On Error GoTo ERRH
                          ' ** After discovering the long Ledger description, this may no longer be necessary (or used).
12700                     IncomeExpense_Export strQry1, strRptPathFile1, .UserReportPath, "Expense"  ' ** Module Function: modExcelFuncs.
                          'Debug.Print "'EXPENSE: 3RD TRY!"
12710                   Case Else
12720                     blnContinue = False
12730                     Set rst = dbs.OpenRecordset("tblErrorLog", dbOpenDynaset, dbConsistent)
12740                     zErrorWriteRecord ERR.Number, ERR.description, THIS_NAME, THIS_PROC, Erl, rst  ' ** Module Function: modErrorHandler.
12750                     rst.Close
12760                     Set rst = Nothing
12770                     Beep
12780                     DoCmd.Hourglass False
12790                     MsgBox "An error was detected while attempting to export the data to Excel." & vbCrLf & _
                            "  Error: " & CStr(ERR.Number) & vbCrLf & _
                            "  Description: " & ERR.description & vbCrLf & _
                            "Please contact Delta Data, Inc., for assistance.", vbInformation + vbOKOnly, "Error: " & CStr(ERR.Number)
12800 On Error GoTo ERRH
12810                   End Select
12820                 Else
12830 On Error GoTo ERRH
                        'Debug.Print "'EXPENSE: 2ND TRY!"
12840                 End If
12850                 If blnContinue = True Then
12860                   DoEvents
12870                   If Excel_IncExp(strRptPathFile1, "Expense") = True Then  ' ** Module Function: modExcelFuncs.
12880                     DoEvents
12890                     Select Case .chkOpenExcel
                          Case True
                            ' ** Even though all references to the Excel_IncExp() objects are explicit,
                            ' ** and they're closed and quit, sometimes the process will not shut down.
                            ' ** (One suggestion is that the worksheet copy is the culprit.)
                            ' ** This, below, seems to be my only recourse.
12900                       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
12910                         EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
12920                       End If
12930                       If strQry2 <> vbNullString And strRptPathFile2 <> vbNullString Then
                              ' ** Don't open it yet.
12940                       Else
12950                         OpenExe strRptPathFile1  ' ** Module Function: modShellFuncs.
12960                       End If
12970                     Case False
                            ' ** Nothing, that's it.
12980                     End Select
12990                   Else
13000                     blnContinue = False
13010                   End If
13020                 End If
13030               Case Else
13040                 blnContinue = False
13050                 zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13060 On Error GoTo ERRH
13070               End Select  ' ** Err.Number.
13080             Else
13090 On Error GoTo ERRH
                    'Debug.Print "'EXPENSE: 1ST TRY!"
13100               If Excel_NameOnly(strRptPathFile1, "Expense") = True Then  ' ** Module Function: modExcelFuncs.
13110                 Select Case .chkOpenExcel
                      Case True
13120                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
13130                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
13140                   End If
13150                   If strQry2 <> vbNullString And strRptPathFile2 <> vbNullString Then
                          ' ** Don't open it yet.
13160                   Else
13170                     OpenExe strRptPathFile1  ' ** Module Function: modShellFuncs.
13180                   End If
13190                 Case False
                        ' ** Nothing, that's it.
13200                 End Select
13210               Else
13220                 blnContinue = False
13230               End If
13240             End If
13250           End If  ' ** vbNullString.

13260         End If  ' ** blnContinue.

13270         If blnContinue = True Then

                ' ** Export 2.
13280           If strQry2 <> vbNullString And strRptPathFile2 <> vbNullString Then
13290             strFile4 = strRptPathFile2
13300 On Error Resume Next
13310             DoCmd.OutputTo acOutputQuery, strQry2, acFormatXLS, strRptPathFile2, False
13320             If ERR.Number <> 0 Then
13330               Select Case ERR.Number
                    Case 2306  ' ** There are too many rows to output, based on the limitation specified by the output format or by Microsoft Access.
13340 On Error GoTo ERRH
13350                 DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, strQry2, strRptPathFile2, True
13360                 DoEvents
13370                 Select Case .chkOpenExcel
                      Case True
13380                   OpenExe strRptPathFile2  ' ** Module Function: modShellFuncs.
13390                 Case False
                        ' ** Nothing, that's it.
13400                 End Select
13410               Case Else
13420                 blnContinue = False
13430                 zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13440 On Error GoTo ERRH
13450               End Select  ' ** Err.Number.
13460             Else
13470 On Error GoTo ERRH
13480               If Excel_NameOnly(strRptPathFile2, "Expense Summary") = True Then  ' ** Module Function: modExcelFuncs.
13490                 Select Case .chkOpenExcel
                      Case True
13500                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
13510                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
13520                   End If
13530                   If strQry1 <> vbNullString And strRptPathFile1 <> vbNullString Then
13540                     OpenExe strRptPathFile1  ' ** Module Function: modShellFuncs.
13550                   End If
13560                   OpenExe strRptPathFile2  ' ** Module Function: modShellFuncs.
13570                 Case False
                        ' ** Nothing, that's it.
13580                 End Select
13590               Else
13600                 blnContinue = False
13610               End If
13620             End If
13630           End If  ' ** vbNullString.

13640         End If  ' ** blnContinue.

13650       End If  ' ** DoReport().
13660     End If  ' ** blnContinue.
13670   End With  ' ** Me.

13680   DoCmd.Hourglass False

EXITP:
13690   Set rst = Nothing
13700   Set dbs = Nothing
13710   Exit Sub

ERRH:
13720   DoCmd.Hourglass False
13730   Select Case ERR.Number
        Case Else
13740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13750   End Select
13760   Resume EXITP

End Sub

Public Sub ArchiveSet_IE(frm As Access.Form)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "ArchiveSet_IE"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim blnHasArchs As Boolean

13810   Set dbs = CurrentDb
13820   With dbs
13830     Set rst = .OpenRecordset("LedgerArchive", dbOpenDynaset, dbReadOnly)
13840     With rst
13850       If .BOF = True And .EOF = True Then
13860         blnHasArchs = False
13870       Else
13880         .MoveLast
13890         If .RecordCount = 1 Then
13900           blnHasArchs = False
13910         Else
13920           blnHasArchs = True
13930         End If
13940       End If
13950       .Close
13960     End With
13970     Set rst = Nothing
13980     .Close
13990   End With
14000   Set dbs = Nothing

14010   With frm
14020     Select Case blnHasArchs
          Case True
14030       .chkIncludeArchive_lbl.Visible = True
14040       .chkIncludeArchive_lbl2.Visible = False
14050       .chkIncludeArchive_lbl2_dim_hi.Visible = False
14060     Case False
14070       .chkIncludeArchive = False
14080       .chkIncludeArchive.Enabled = False
14090       .chkIncludeArchive_lbl.Visible = False
14100       .chkIncludeArchive_lbl2.Visible = True
14110       .chkIncludeArchive_lbl2_dim_hi.Visible = True
14120     End Select
14130   End With

EXITP:
14140   Set rst = Nothing
14150   Set dbs = Nothing
14160   Exit Sub

ERRH:
14170   Select Case ERR.Number
        Case Else
14180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14190   End Select
14200   Resume EXITP

End Sub

Public Function DoReport_IE(frm As Access.Form) As Boolean

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport_IE"

        Dim blnRetVal As Boolean

14310   blnRetVal = True

14320   With frm

          ' ** Validate main three data values.
14330     If IsNull(.DateStart.Value) Then
14340       blnRetVal = False
14350       DoCmd.Hourglass False
14360       MsgBox "You must enter a From date to continue.", vbInformation + vbOKOnly, "Entry Required"
14370       .DateStart.SetFocus
14380     Else
14390       If IsNull(.DateEnd.Value) Then
14400         blnRetVal = False
14410         DoCmd.Hourglass False
14420         MsgBox "You must enter a To date to continue.", vbInformation + vbOKOnly, "Entry Required"
14430         .DateEnd.SetFocus
14440       Else
14450         If .DateStart.Value >= .DateEnd.Value Then
14460           blnRetVal = False
14470           DoCmd.Hourglass False
14480           MsgBox "You must enter a From date that is less than the To date to continue.", vbInformation + vbOKOnly, "Invalid Entry"
14490           .DateEnd.SetFocus
14500         Else
14510           If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue Then
14520             If IsNull(.cmbAccounts.Column(0)) Or .cmbAccounts.Column(0) = vbNullString Then
14530               blnRetVal = False
14540               DoCmd.Hourglass False
14550               MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
14560               .cmbAccounts.SetFocus
14570             End If
14580           End If
14590         End If
14600       End If
14610     End If

          ' ** Borrowing this variable from Court Reports.
14620     gstrCrtRpt_Ordinal = Report_Criteria_Msg_IE(frm)  ' ** Function: Above.

14630     If blnRetVal = True Then
            ' ** Always rebuild the temp table.

            ' ** Set global variables for report headers.
14640       gdatStartDate = .DateStart.Value
14650       gdatEndDate = .DateEnd.Value
14660       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
14670         gstrAccountNo = .cmbAccounts.Column(0)
14680         gstrAccountName = .cmbAccounts.Column(3)
14690       Case .opgAccountNumber_optAll.OptionValue
14700         gstrAccountNo = "ALL"
14710         gstrAccountName = vbNullString
14720       End Select

            ' ** They always exclude both Hidden Transactions and Non-Active Income/Expense Codes.
14730       blnRetVal = IncomeExpense_BuildTable(gdatStartDate, gdatEndDate, gstrAccountNo, False, .chkIncludeArchive)   ' ** Function: Above.

14740     End If
14750   End With

EXITP:
14760   DoReport_IE = blnRetVal
14770   Exit Function

ERRH:
14780   blnRetVal = False
14790   DoCmd.Hourglass False
14800   Select Case ERR.Number
        Case 3000  ' ** Reserved. There is no message for this error.
14810     Beep
14820     MsgBox "Trust Accountant is unable to complete your request." & vbCrLf & _
            "Exit the program, then try again.", vbCritical + vbOKOnly, "Error 3000"
14830   Case Else
14840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14850   End Select
14860   Resume EXITP

End Function

Public Sub OptGrpSet_IE(frm As Access.Form)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "OptGrpSet_IE"

14910   With frm

14920     Select Case .opgOptionGroups_box.Visible
          Case True
14930       .cmdMoreOptions_L_raised_img.Visible = True
14940       .cmdMoreOptions_L_raised_semifocus_dots_img.Visible = False
14950       .cmdMoreOptions_L_raised_focus_img.Visible = False
14960       .cmdMoreOptions_L_raised_focus_dots_img.Visible = False
14970       .cmdMoreOptions_L_sunken_focus_dots_img.Visible = False
14980       .cmdMoreOptions_L_raised_img_dis.Visible = False
14990       .cmdMoreOptions_R_raised_img.Visible = False
15000       .cmdMoreOptions_R_raised_semifocus_dots_img.Visible = False
15010       .cmdMoreOptions_R_raised_focus_img.Visible = False
15020       .cmdMoreOptions_R_raised_focus_dots_img.Visible = False
15030       .cmdMoreOptions_R_sunken_focus_dots_img.Visible = False
15040       .cmdMoreOptions_R_raised_img_dis.Visible = False
15050     Case False
15060       .cmdMoreOptions_R_raised_img.Visible = True
15070       .cmdMoreOptions_R_raised_semifocus_dots_img.Visible = False
15080       .cmdMoreOptions_R_raised_focus_img.Visible = False
15090       .cmdMoreOptions_R_raised_focus_dots_img.Visible = False
15100       .cmdMoreOptions_R_sunken_focus_dots_img.Visible = False
15110       .cmdMoreOptions_R_raised_img_dis.Visible = False
15120       .cmdMoreOptions_L_raised_img.Visible = False
15130       .cmdMoreOptions_L_raised_semifocus_dots_img.Visible = False
15140       .cmdMoreOptions_L_raised_focus_img.Visible = False
15150       .cmdMoreOptions_L_raised_focus_dots_img.Visible = False
15160       .cmdMoreOptions_L_sunken_focus_dots_img.Visible = False
15170       .cmdMoreOptions_L_raised_img_dis.Visible = False
15180     End Select

15190   End With

EXITP:
15200   Exit Sub

ERRH:
15210   Select Case ERR.Number
        Case Else
15220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15230   End Select
15240   Resume EXITP

End Sub

Public Sub Calendar_Handler_IE(strProc As String, blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean, blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean, clsMonthClass As clsMonthCal, frm As Access.Form)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Handler_IE"

        Dim strAction As String, strCtl As String
        Dim intPos01 As Integer, lngCnt As Long
        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

15310   With frm

15320     lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
15330     intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
15340     strAction = Mid(strProc, (intPos01 + 1))
15350     strCtl = Left(strProc, (intPos01 - 1))

15360     Select Case strAction
          Case "GotFocus"
15370       Select Case strCtl
            Case "cmdCalendar1"  ' ** cmdCalendar1: GotFocus().
15380         blnCalendar1_Focus = True
15390         .cmdCalendar1_raised_semifocus_dots_img.Visible = True
15400         .cmdCalendar1_raised_img.Visible = False
15410         .cmdCalendar1_raised_focus_img.Visible = False
15420         .cmdCalendar1_raised_focus_dots_img.Visible = False
15430         .cmdCalendar1_sunken_focus_dots_img.Visible = False
15440         .cmdCalendar1_raised_img_dis.Visible = False
15450       Case "cmdCalendar2"  ' ** cmdCalendar2: GotFocus().
15460         blnCalendar2_Focus = True
15470         .cmdCalendar2_raised_semifocus_dots_img.Visible = True
15480         .cmdCalendar2_raised_img.Visible = False
15490         .cmdCalendar2_raised_focus_img.Visible = False
15500         .cmdCalendar2_raised_focus_dots_img.Visible = False
15510         .cmdCalendar2_sunken_focus_dots_img.Visible = False
15520         .cmdCalendar2_raised_img_dis.Visible = False
15530       End Select
15540     Case "MouseDown"
15550       Select Case strCtl
            Case "cmdCalendar1"  ' ** cmdCalendar1: MouseDown().
15560         blnCalendar1_MouseDown = True
15570         .cmdCalendar1_sunken_focus_dots_img.Visible = True
15580         .cmdCalendar1_raised_img.Visible = False
15590         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
15600         .cmdCalendar1_raised_focus_img.Visible = False
15610         .cmdCalendar1_raised_focus_dots_img.Visible = False
15620         .cmdCalendar1_raised_img_dis.Visible = False
15630       Case "cmdCalendar2"  ' ** cmdCalendar2: MouseDown().
15640         blnCalendar2_MouseDown = True
15650         .cmdCalendar2_sunken_focus_dots_img.Visible = True
15660         .cmdCalendar2_raised_img.Visible = False
15670         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
15680         .cmdCalendar2_raised_focus_img.Visible = False
15690         .cmdCalendar2_raised_focus_dots_img.Visible = False
15700         .cmdCalendar2_raised_img_dis.Visible = False
15710       End Select
15720     Case "MouseMove"
15730       Select Case strCtl
            Case "cmdCalendar1"  ' ** cmdCalendar1: MouseMove().
15740         If blnCalendar1_MouseDown = False Then
15750           Select Case blnCalendar1_Focus
                Case True
15760             .cmdCalendar1_raised_focus_dots_img.Visible = True
15770             .cmdCalendar1_raised_focus_img.Visible = False
15780           Case False
15790             .cmdCalendar1_raised_focus_img.Visible = True
15800             .cmdCalendar1_raised_focus_dots_img.Visible = False
15810           End Select
15820           .cmdCalendar1_raised_img.Visible = False
15830           .cmdCalendar1_raised_semifocus_dots_img.Visible = False
15840           .cmdCalendar1_sunken_focus_dots_img.Visible = False
15850           .cmdCalendar1_raised_img_dis.Visible = False
15860         End If
15870       Case "cmdCalendar2"  ' ** cmdCalendar2: MouseMove().
15880         If blnCalendar2_MouseDown = False Then
15890           Select Case blnCalendar2_Focus
                Case True
15900             .cmdCalendar2_raised_focus_dots_img.Visible = True
15910             .cmdCalendar2_raised_focus_img.Visible = False
15920           Case False
15930             .cmdCalendar2_raised_focus_img.Visible = True
15940             .cmdCalendar2_raised_focus_dots_img.Visible = False
15950           End Select
15960           .cmdCalendar2_raised_img.Visible = False
15970           .cmdCalendar2_raised_semifocus_dots_img.Visible = False
15980           .cmdCalendar2_sunken_focus_dots_img.Visible = False
15990           .cmdCalendar2_raised_img_dis.Visible = False
16000         End If
16010       End Select
16020     Case "MouseUp"
16030       Select Case strCtl
            Case "cmdCalendar1"  ' ** cmdCalendar1: MouseUp().
16040         .cmdCalendar1_raised_focus_dots_img.Visible = True
16050         .cmdCalendar1_raised_img.Visible = False
16060         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
16070         .cmdCalendar1_raised_focus_img.Visible = False
16080         .cmdCalendar1_sunken_focus_dots_img.Visible = False
16090         .cmdCalendar1_raised_img_dis.Visible = False
16100         blnCalendar1_MouseDown = False
16110       Case "cmdCalendar2"  ' ** cmdCalendar2: MouseUp().
16120         .cmdCalendar2_raised_focus_dots_img.Visible = True
16130         .cmdCalendar2_raised_img.Visible = False
16140         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
16150         .cmdCalendar2_raised_focus_img.Visible = False
16160         .cmdCalendar2_sunken_focus_dots_img.Visible = False
16170         .cmdCalendar2_raised_img_dis.Visible = False
16180         blnCalendar2_MouseDown = False
16190       End Select
16200     Case "LostFocus"
16210       Select Case strCtl
            Case "cmdCalendar1"  ' ** cmdCalendar1: LostFocus().
16220         .cmdCalendar1_raised_img.Visible = True
16230         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
16240         .cmdCalendar1_raised_focus_img.Visible = False
16250         .cmdCalendar1_raised_focus_dots_img.Visible = False
16260         .cmdCalendar1_sunken_focus_dots_img.Visible = False
16270         .cmdCalendar1_raised_img_dis.Visible = False
16280         blnCalendar1_Focus = False
16290       Case "cmdCalendar2"  ' ** cmdCalendar2: LostFocus().
16300         .cmdCalendar2_raised_img.Visible = True
16310         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
16320         .cmdCalendar2_raised_focus_img.Visible = False
16330         .cmdCalendar2_raised_focus_dots_img.Visible = False
16340         .cmdCalendar2_sunken_focus_dots_img.Visible = False
16350         .cmdCalendar2_raised_img_dis.Visible = False
16360         blnCalendar2_Focus = False
16370       End Select
16380     Case "Click"
16390       Select Case strCtl
            Case "cmdCalendar1"  ' ** cmdCalendar1: Click().
16400         datStartDate = Date
16410         datEndDate = 0
16420         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
16430         If blnRetVal = True Then
16440           .DateStart = datStartDate
16450         Else
16460           .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
16470         End If
16480         .DateStart.SetFocus
16490       Case "cmdCalendar2"  ' ** cmdCalendar2: Click().
16500         datStartDate = Date
16510         datEndDate = 0
16520         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
16530         If blnRetVal = True Then
16540           .DateEnd = datStartDate
16550         Else
16560           .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
16570         End If
16580         .DateEnd.SetFocus
16590         Cancel = 0
16600         .DateEnd_Exit Cancel  ' ** Form Procedure: frmRpt_IncomeExpense
16610         If Cancel = 0 Then
16620           .opgAccountNumber.SetFocus
16630         End If
16640       End Select
16650     End Select

16660   End With

EXITP:
16670   Exit Sub

ERRH:
16680   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
16690   Case Else
16700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16710   End Select
16720   Resume EXITP

End Sub

Public Sub OptionHandler_IE(intMode As Integer, blnMoreOptions_Focus As Boolean, blnMoreOptions_MouseDown As Boolean, frm As Access.Form)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "OptionHandler_IE"

16810   With frm
16820     Select Case intMode
          Case 1
16830       blnMoreOptions_Focus = True
16840       Select Case .opgOptionGroups_box.Visible
            Case True
16850         .cmdMoreOptions_L_raised_semifocus_dots_img.Visible = True
16860         .cmdMoreOptions_L_raised_img.Visible = False
16870         .cmdMoreOptions_L_raised_focus_img.Visible = False
16880         .cmdMoreOptions_L_raised_focus_dots_img.Visible = False
16890         .cmdMoreOptions_L_sunken_focus_dots_img.Visible = False
16900         .cmdMoreOptions_L_raised_img_dis.Visible = False
16910       Case False
16920         .cmdMoreOptions_R_raised_semifocus_dots_img.Visible = True
16930         .cmdMoreOptions_R_raised_img.Visible = False
16940         .cmdMoreOptions_R_raised_focus_img.Visible = False
16950         .cmdMoreOptions_R_raised_focus_dots_img.Visible = False
16960         .cmdMoreOptions_R_sunken_focus_dots_img.Visible = False
16970         .cmdMoreOptions_R_raised_img_dis.Visible = False
16980       End Select
16990     Case 2
17000       blnMoreOptions_MouseDown = True
17010       Select Case .opgOptionGroups_box.Visible
            Case True
17020         .cmdMoreOptions_L_sunken_focus_dots_img.Visible = True
17030         .cmdMoreOptions_L_raised_img.Visible = False
17040         .cmdMoreOptions_L_raised_semifocus_dots_img.Visible = False
17050         .cmdMoreOptions_L_raised_focus_img.Visible = False
17060         .cmdMoreOptions_L_raised_focus_dots_img.Visible = False
17070         .cmdMoreOptions_L_raised_img_dis.Visible = False
17080       Case False
17090         .cmdMoreOptions_R_sunken_focus_dots_img.Visible = True
17100         .cmdMoreOptions_R_raised_img.Visible = False
17110         .cmdMoreOptions_R_raised_semifocus_dots_img.Visible = False
17120         .cmdMoreOptions_R_raised_focus_img.Visible = False
17130         .cmdMoreOptions_R_raised_focus_dots_img.Visible = False
17140         .cmdMoreOptions_R_raised_img_dis.Visible = False
17150       End Select
17160     Case 3
17170       If blnMoreOptions_MouseDown = False Then
17180         Select Case .opgOptionGroups_box.Visible
              Case True
17190           Select Case blnMoreOptions_Focus
                Case True
17200             .cmdMoreOptions_L_raised_focus_dots_img.Visible = True
17210             .cmdMoreOptions_L_raised_focus_img.Visible = False
17220           Case False
17230             .cmdMoreOptions_L_raised_focus_img.Visible = True
17240             .cmdMoreOptions_L_raised_focus_dots_img.Visible = False
17250           End Select
17260           .cmdMoreOptions_L_raised_img.Visible = False
17270           .cmdMoreOptions_L_raised_semifocus_dots_img.Visible = False
17280           .cmdMoreOptions_L_sunken_focus_dots_img.Visible = False
17290           .cmdMoreOptions_L_raised_img_dis.Visible = False
17300         Case False
17310           Select Case blnMoreOptions_Focus
                Case True
17320             .cmdMoreOptions_R_raised_focus_dots_img.Visible = True
17330             .cmdMoreOptions_R_raised_focus_img.Visible = False
17340           Case False
17350             .cmdMoreOptions_R_raised_focus_img.Visible = True
17360             .cmdMoreOptions_R_raised_focus_dots_img.Visible = False
17370           End Select
17380           .cmdMoreOptions_R_raised_img.Visible = False
17390           .cmdMoreOptions_R_raised_semifocus_dots_img.Visible = False
17400           .cmdMoreOptions_R_sunken_focus_dots_img.Visible = False
17410           .cmdMoreOptions_R_raised_img_dis.Visible = False
17420         End Select
17430       End If
17440     Case 4
17450       Select Case .opgOptionGroups_box.Visible
            Case True
17460         .cmdMoreOptions_L_raised_focus_dots_img.Visible = True
17470         .cmdMoreOptions_L_raised_img.Visible = False
17480         .cmdMoreOptions_L_raised_semifocus_dots_img.Visible = False
17490         .cmdMoreOptions_L_raised_focus_img.Visible = False
17500         .cmdMoreOptions_L_sunken_focus_dots_img.Visible = False
17510         .cmdMoreOptions_L_raised_img_dis.Visible = False
17520       Case False
17530         .cmdMoreOptions_R_raised_focus_dots_img.Visible = True
17540         .cmdMoreOptions_R_raised_img.Visible = False
17550         .cmdMoreOptions_R_raised_semifocus_dots_img.Visible = False
17560         .cmdMoreOptions_R_raised_focus_img.Visible = False
17570         .cmdMoreOptions_R_sunken_focus_dots_img.Visible = False
17580         .cmdMoreOptions_R_raised_img_dis.Visible = False
17590       End Select
17600       blnMoreOptions_MouseDown = False
17610     Case 5
17620       Select Case .opgOptionGroups_box.Visible
            Case True
17630         .cmdMoreOptions_L_raised_img.Visible = True
17640         .cmdMoreOptions_L_raised_semifocus_dots_img.Visible = False
17650         .cmdMoreOptions_L_raised_focus_img.Visible = False
17660         .cmdMoreOptions_L_raised_focus_dots_img.Visible = False
17670         .cmdMoreOptions_L_sunken_focus_dots_img.Visible = False
17680         .cmdMoreOptions_L_raised_img_dis.Visible = False
17690       Case False
17700         .cmdMoreOptions_R_raised_img.Visible = True
17710         .cmdMoreOptions_R_raised_semifocus_dots_img.Visible = False
17720         .cmdMoreOptions_R_raised_focus_img.Visible = False
17730         .cmdMoreOptions_R_raised_focus_dots_img.Visible = False
17740         .cmdMoreOptions_R_sunken_focus_dots_img.Visible = False
17750         .cmdMoreOptions_R_raised_img_dis.Visible = False
17760       End Select
17770       blnMoreOptions_Focus = False
17780     End Select
17790   End With

EXITP:
17800   Exit Sub

ERRH:
17810   Select Case ERR.Number
        Case Else
17820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17830   End Select
17840   Resume EXITP

End Sub

Public Sub DoAll_Handler_IE(strProc As String, blnPrintAll_Focus As Boolean, blnWordAll_Focus As Boolean, blnExcelAll_Focus As Boolean, blnPrintBoth As Boolean, strFile1 As String, strFile2 As String, strFile3 As String, strFile4 As String, frm As Access.Form)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "DoAll_Handler_IE"

        Dim strAction As String, strCtl As String
        Dim intPos01 As Integer, lngCnt As Long

17910   With frm

17920     lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
17930     intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
17940     strAction = Mid(strProc, (intPos01 + 1))
17950     strCtl = Left(strProc, (intPos01 - 1))

17960     Select Case strAction
          Case "Click"

17970       Select Case strCtl
            Case "cmdPrintAll"
17980         DoCmd.Hourglass True
17990         DoEvents
18000         blnPrintBoth = True
18010         .cmdRevIncExp_IncomePrint_Click  ' ** Form Procedure: frmRpt_IncomeExpense.
18020         DoCmd.Hourglass True
18030         DoEvents
18040         If blnPrintBoth = True Then
18050           .cmdRevIncExp_ExpensePrint_Click  ' ** Form Procedure: frmRpt_IncomeExpense.
18060         End If
18070         blnPrintBoth = False
18080         DoCmd.Hourglass False

18090       Case "cmdWordAll"
18100         DoCmd.Hourglass True
18110         DoEvents
18120         gblnPrintAll = True
18130         strFile1 = vbNullString: strFile2 = vbNullString: strFile3 = vbNullString: strFile4 = vbNullString
18140         .cmdRevIncExp_IncomeWord_Click  ' ** Form Procedure: frmRpt_IncomeExpense.
18150         DoEvents
18160         If gblnPrintAll = True Then
18170           .cmdRevIncExp_ExpenseWord_Click  ' ** Form Procedure: frmRpt_IncomeExpense.
18180         End If
18190         gblnPrintAll = False
18200         DoCmd.Hourglass False

18210       Case "cmdExcelAll"
18220         DoCmd.Hourglass True
18230         DoEvents
18240         gblnPrintAll = True
18250         strFile1 = vbNullString: strFile2 = vbNullString: strFile3 = vbNullString: strFile4 = vbNullString
              ' ** Excel not opened after export.
18260         .cmdRevIncExp_IncomeExcel_Click  ' ** Form Procedure: frmRpt_IncomeExpense.
18270         DoEvents
18280         If gblnPrintAll = True Then
                ' ** Excel not opened after export.
18290           .cmdRevIncExp_ExpenseExcel_Click  ' ** Form Procedure: frmRpt_IncomeExpense.
18300         End If
18310         If gblnPrintAll = True Then
18320           Select Case .chkOpenExcel
                Case True
18330             DoCmd.Hourglass True
18340             DoEvents
18350             If strFile1 <> vbNullString Then
18360               If Excel_NameOnly(strFile1, "Income") = True Then  ' ** Module Function: modExcelFuncs.
18370                 If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
18380                   EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
18390                 End If
18400                 OpenExe strFile1  ' ** Module Function: modShellFuncs.
18410               End If
18420               DoEvents
18430               If strFile2 <> vbNullString Or strFile3 <> vbNullString Or strFile4 <> vbNullString Then
18440                 ForcePause 2  ' ** Module Function: modCodeUtilities.
18450               End If
18460             End If
18470             If strFile3 <> vbNullString Then
18480               If Excel_NameOnly(strFile3, "Income Summary") = True Then  ' ** Module Function: modExcelFuncs.
18490                 If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
18500                   EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
18510                 End If
18520                 OpenExe strFile3  ' ** Module Function: modShellFuncs.
18530               End If
18540               DoEvents
18550               If strFile2 <> vbNullString Or strFile4 <> vbNullString Then
18560                 ForcePause 2  ' ** Module Function: modCodeUtilities.
18570               End If
18580             End If
18590             If strFile2 <> vbNullString Then
18600               If Excel_NameOnly(strFile2, "Expense") = True Then  ' ** Module Function: modExcelFuncs.
18610                 If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
18620                   EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
18630                 End If
18640                 OpenExe strFile2  ' ** Module Function: modShellFuncs.
18650               End If
18660               DoEvents
18670               If strFile4 <> vbNullString Then
18680                 ForcePause 2  ' ** Module Function: modCodeUtilities.
18690               End If
18700             End If
18710             If strFile4 <> vbNullString Then
18720               If Excel_NameOnly(strFile2, "Expense Summary") = True Then  ' ** Module Function: modExcelFuncs.
18730                 If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
18740                   EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
18750                 End If
18760                 OpenExe strFile2  ' ** Module Function: modShellFuncs.
18770               End If
18780               DoEvents
18790             End If
18800           Case False
                  ' ** Nothing, that's it.
18810           End Select
18820           DoEvents
18830         End If
18840         gblnPrintAll = False
18850         DoCmd.Hourglass False
18860       End Select

18870     Case "GotFocus"

18880       Select Case strCtl
            Case "cmdPrintAll"
18890         blnPrintAll_Focus = True
18900         .cmdPrintAll_box01.Visible = True
18910         .cmdPrintAll_box02.Visible = True
18920       Case "cmdWordAll"
18930         blnWordAll_Focus = True
18940         .cmdWordAll_box01.Visible = True
18950         .cmdWordAll_box02.Visible = True
18960       Case "cmdExcelAll"
18970         blnExcelAll_Focus = True
18980         .cmdExcelAll_box01.Visible = True
18990         .cmdExcelAll_box02.Visible = True
19000       End Select

19010     Case "LostFocus"

19020       Select Case strCtl
            Case "cmdPrintAll"
19030         .cmdPrintAll_box01.Visible = False
19040         .cmdPrintAll_box02.Visible = False
19050         blnPrintAll_Focus = False
19060       Case "cmdWordAll"
19070         .cmdWordAll_box01.Visible = False
19080         .cmdWordAll_box02.Visible = False
19090         blnWordAll_Focus = False
19100       Case "cmdExcelAll"
19110         .cmdExcelAll_box01.Visible = False
19120         .cmdExcelAll_box02.Visible = False
19130         blnExcelAll_Focus = False
19140       End Select

19150     Case "MouseMove"

19160       Select Case strCtl
            Case "cmdPrintAll"
19170         If gblnPrintAll = False Then
19180           .cmdPrintAll_box01.Visible = True
19190           .cmdPrintAll_box02.Visible = True
19200           If blnWordAll_Focus = False Then
19210             .cmdWordAll_box01.Visible = False
19220             .cmdWordAll_box02.Visible = False
19230           End If
19240           If blnExcelAll_Focus = False Then
19250             .cmdExcelAll_box01.Visible = False
19260             .cmdExcelAll_box02.Visible = False
19270           End If
19280         End If
19290       Case "cmdWordAll"
19300         If gblnPrintAll = False Then
19310           .cmdWordAll_box01.Visible = True
19320           .cmdWordAll_box02.Visible = True
19330           If blnPrintAll_Focus = False Then
19340             .cmdPrintAll_box01.Visible = False
19350             .cmdPrintAll_box02.Visible = False
19360           End If
19370           If blnExcelAll_Focus = False Then
19380             .cmdExcelAll_box01.Visible = False
19390             .cmdExcelAll_box02.Visible = False
19400           End If
19410         End If
19420       Case "cmdExcelAll"
19430         If gblnPrintAll = False Then
19440           .cmdExcelAll_box01.Visible = True
19450           .cmdExcelAll_box02.Visible = True
19460           If blnPrintAll_Focus = False Then
19470             .cmdPrintAll_box01.Visible = False
19480             .cmdPrintAll_box02.Visible = False
19490           End If
19500           If blnWordAll_Focus = False Then
19510             .cmdWordAll_box01.Visible = False
19520             .cmdWordAll_box02.Visible = False
19530           End If
19540         End If
19550       End Select

19560     End Select

19570   End With

EXITP:
19580   Exit Sub

ERRH:
19590   Select Case ERR.Number
        Case Else
19600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19610   End Select
19620   Resume EXITP

End Sub

Public Sub OpenApp_Handler_IE(strProc As String, frm As Access.Form)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "OpenApp_Handler_IE"

        Dim strAction As String, strCtl As String
        Dim intPos01 As Integer, lngCnt As Long

19710   With frm

19720     lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
19730     intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
19740     strAction = Mid(strProc, (intPos01 + 1))
19750     strCtl = Left(strProc, (intPos01 - 1))

19760     Select Case strAction
          Case "AfterUpdate"
19770       Select Case strCtl
            Case "chkOpenWord"
19780         Select Case .chkOpenWord
              Case True
19790           .chkOpenWord_lbl.FontBold = True
19800           .chkOpenWord_lbl_dim_hi.FontBold = True
19810           .chkOpenWord_lbl2.FontBold = True
19820           .chkOpenWord_lbl2_dim_hi.FontBold = True
19830         Case False
19840           .chkOpenWord_lbl.FontBold = False
19850           .chkOpenWord_lbl_dim_hi.FontBold = False
19860           .chkOpenWord_lbl2.FontBold = False
19870           .chkOpenWord_lbl2_dim_hi.FontBold = False
19880         End Select
19890       Case "chkOpenExcel"
19900         Select Case .chkOpenExcel
              Case True
19910           .chkOpenExcel_lbl.FontBold = True
19920           .chkOpenExcel_lbl_dim_hi.FontBold = True
19930           .chkOpenExcel_lbl2.FontBold = True
19940           .chkOpenExcel_lbl2_dim_hi.FontBold = True
19950         Case False
19960           .chkOpenExcel_lbl.FontBold = False
19970           .chkOpenExcel_lbl_dim_hi.FontBold = False
19980           .chkOpenExcel_lbl2.FontBold = False
19990           .chkOpenExcel_lbl2_dim_hi.FontBold = False
20000         End Select
20010       End Select
20020     End Select

20030   End With

EXITP:
20040   Exit Sub

ERRH:
20050   Select Case ERR.Number
        Case Else
20060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20070   End Select
20080   Resume EXITP

End Sub

Public Sub FormLoad_IE(intMode As Integer, frm As Access.Form)

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "FormLoad_IE"

20110   With frm
20120     Select Case intMode
          Case 1
20130       .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.
20140       DoCmd.Hourglass True  ' ** Make sure it's still running.
20150       DoEvents
20160       .chkRememberDates_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20170       .opgAccountSource_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20180       .chkRememberMe_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20190       .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20200       .chkIncludeArchive_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20210       .chkDetail_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20220       .chkPageOf_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20230       .opgUnspecified_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20240       .opgPrincipalCash_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20250       .opgZeroCash_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20260       .opgSummary_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20270       .chkDontCombineMulti_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20280       .chkAcctEveryLine_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20290       .chkSweepOnly_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
20300     Case 2
20310       OptionsChk_IE frm  ' ** Procedure: Above.
20320       OptGrpSet_IE frm  ' ** Procedure: Above.
20330     End Select
20340   End With

EXITP:
20350   Exit Sub

ERRH:
20360   Select Case ERR.Number
        Case Else
20370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20380   End Select
20390   Resume EXITP

End Sub

Public Sub Detail_Mouse_IE(blnCalendar1_Focus As Boolean, blnCalendar2_Focus As Boolean, blnMoreOptions_Focus As Boolean, blnReset_Focus As Boolean, blnPrintAll_Focus As Boolean, blnWordAll_Focus As Boolean, blnExcelAll_Focus As Boolean, frm As Access.Form)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Mouse_IE"

20410   With frm
20420     If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
20430       Select Case blnCalendar1_Focus
            Case True
20440         .cmdCalendar1_raised_semifocus_dots_img.Visible = True
20450         .cmdCalendar1_raised_img.Visible = False
20460       Case False
20470         .cmdCalendar1_raised_img.Visible = True
20480         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
20490       End Select
20500       .cmdCalendar1_raised_focus_dots_img.Visible = False
20510       .cmdCalendar1_raised_focus_img.Visible = False
20520       .cmdCalendar1_sunken_focus_dots_img.Visible = False
20530       .cmdCalendar1_raised_img_dis.Visible = False
20540     End If
20550     If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
20560       Select Case blnCalendar2_Focus
            Case True
20570         .cmdCalendar2_raised_semifocus_dots_img.Visible = True
20580         .cmdCalendar2_raised_img.Visible = False
20590       Case False
20600         .cmdCalendar2_raised_img.Visible = True
20610         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
20620       End Select
20630       .cmdCalendar2_raised_focus_dots_img.Visible = False
20640       .cmdCalendar2_raised_focus_img.Visible = False
20650       .cmdCalendar2_sunken_focus_dots_img.Visible = False
20660       .cmdCalendar2_raised_img_dis.Visible = False
20670     End If
20680     If .cmdMoreOptions_R_raised_focus_dots_img.Visible = True Or .cmdMoreOptions_R_raised_focus_img.Visible = True Then
20690       Select Case blnMoreOptions_Focus
            Case True
20700         .cmdMoreOptions_R_raised_semifocus_dots_img.Visible = True
20710         .cmdMoreOptions_R_raised_img.Visible = False
20720       Case False
20730         .cmdMoreOptions_R_raised_img.Visible = True
20740         .cmdMoreOptions_R_raised_semifocus_dots_img.Visible = False
20750       End Select
20760       .cmdMoreOptions_R_raised_focus_img.Visible = False
20770       .cmdMoreOptions_R_raised_focus_dots_img.Visible = False
20780       .cmdMoreOptions_R_sunken_focus_dots_img.Visible = False
20790       .cmdMoreOptions_R_raised_img_dis.Visible = False
20800     End If
20810     If .cmdMoreOptions_L_raised_focus_dots_img.Visible = True Or .cmdMoreOptions_L_raised_focus_img.Visible = True Then
20820       Select Case blnMoreOptions_Focus
            Case True
20830         .cmdMoreOptions_L_raised_semifocus_dots_img.Visible = True
20840         .cmdMoreOptions_L_raised_img.Visible = False
20850       Case False
20860         .cmdMoreOptions_L_raised_img.Visible = True
20870         .cmdMoreOptions_L_raised_semifocus_dots_img.Visible = False
20880       End Select
20890       .cmdMoreOptions_L_raised_focus_img.Visible = False
20900       .cmdMoreOptions_L_raised_focus_dots_img.Visible = False
20910       .cmdMoreOptions_L_sunken_focus_dots_img.Visible = False
20920       .cmdMoreOptions_L_raised_img_dis.Visible = False
20930     End If
20940     If .cmdReset_raised_focus_dots_img.Visible = True Or .cmdReset_raised_focus_img.Visible = True Then
20950       Select Case blnReset_Focus
            Case True
20960         .cmdReset_raised_semifocus_dots_img.Visible = True
20970         .cmdReset_raised_img.Visible = False
20980       Case False
20990         .cmdReset_raised_img.Visible = True
21000         .cmdReset_raised_semifocus_dots_img.Visible = False
21010       End Select
21020       .cmdReset_raised_focus_img.Visible = False
21030       .cmdReset_raised_focus_dots_img.Visible = False
21040       .cmdReset_sunken_focus_dots_img.Visible = False
21050       .cmdReset_raised_img_dis.Visible = False
21060     End If
21070     If blnPrintAll_Focus = False And (.cmdPrintAll_box01.Visible = True Or .cmdPrintAll_box02.Visible = True) Then
21080       .cmdPrintAll_box01.Visible = False
21090       .cmdPrintAll_box02.Visible = False
21100     End If
21110     If blnWordAll_Focus = False And (.cmdWordAll_box01.Visible = True Or .cmdWordAll_box02.Visible = True) Then
21120       .cmdWordAll_box01.Visible = False
21130       .cmdWordAll_box02.Visible = False
21140     End If
21150     If blnExcelAll_Focus = False And (.cmdExcelAll_box01.Visible = True Or .cmdExcelAll_box02.Visible = True) Then
21160       .cmdExcelAll_box01.Visible = False
21170       .cmdExcelAll_box02.Visible = False
21180     End If
21190   End With

EXITP:
21200   Exit Sub

ERRH:
21210   Select Case ERR.Number
        Case Else
21220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21230   End Select
21240   Resume EXITP

End Sub

Public Sub AcctNum_After_IE(frm As Access.Form)

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "AcctNum_After_IE"

21310   With frm
21320     Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
21330       .opgAccountNumber_optSpecified_lbl.FontBold = True
21340       .opgAccountNumber_optAll_lbl.FontBold = False
21350       .opgAccountNumber_optSpecified_lbl_box.Visible = True
21360       .opgAccountNumber_optAll_lbl_box.Visible = False
21370       .cmbAccounts.Enabled = True
21380       .cmbAccounts.BorderColor = CLR_LTBLU2
21390       .cmbAccounts.BackStyle = acBackStyleNormal
21400       .opgAccountSource.Enabled = True
21410       .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
21420       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
21430       .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
21440       .opgAccountSource_optName_lbl2_dim_hi.Visible = False
21450       .chkRememberMe.Enabled = True
21460       .chkRememberMe_lbl.Visible = True
21470       .chkRememberMe_lbl2_dim.Visible = False
21480       .chkRememberMe_lbl2_dim_hi.Visible = False
21490       .opgSummary = .opgSummary_optExclude.OptionValue
21500       .opgSummary_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
21510       .opgSummary_optInclude.Enabled = False
21520       .opgSummary_optOnly.Enabled = False
21530     Case .opgAccountNumber_optAll.OptionValue
21540       .opgAccountNumber_optSpecified_lbl.FontBold = False
21550       .opgAccountNumber_optAll_lbl.FontBold = True
21560       .opgAccountNumber_optSpecified_lbl_box.Visible = False
21570       .opgAccountNumber_optAll_lbl_box.Visible = True
21580       .cmbAccounts.Enabled = False
21590       .cmbAccounts.BorderColor = WIN_CLR_DISR
21600       .cmbAccounts.BackStyle = acBackStyleTransparent
21610       .opgAccountSource.Enabled = False
21620       .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
21630       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
21640       .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
21650       .opgAccountSource_optName_lbl2_dim_hi.Visible = True
21660       .chkRememberMe.Enabled = False
21670       .chkRememberMe_lbl.Visible = False
21680       .chkRememberMe_lbl2_dim.Visible = True
21690       .chkRememberMe_lbl2_dim_hi.Visible = True
21700       .opgSummary_optInclude.Enabled = True
21710       .opgSummary_optOnly.Enabled = True
21720     End Select
21730   End With

EXITP:
21740   Exit Sub

ERRH:
21750   Select Case ERR.Number
        Case Else
21760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21770   End Select
21780   Resume EXITP

End Sub

Public Sub AcctSrc_After_IE(frm As Access.Form)

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "AcctSrc_After_IE"

        Dim strAccountNo As String

21810   With frm
21820     strAccountNo = vbNullString
21830     If IsNull(.cmbAccounts) = False Then
21840       If Len(.cmbAccounts.Column(0)) > 0 Then
21850         strAccountNo = .cmbAccounts.Column(0)
21860       End If
21870     End If
21880     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
21890       .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
21900       .opgAccountSource_optNumber_lbl2.FontBold = True
21910       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
21920       .opgAccountSource_optName_lbl2.FontBold = False
21930       .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
21940     Case .opgAccountSource_optName.OptionValue
21950       .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
21960       .opgAccountSource_optNumber_lbl2.FontBold = False
21970       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
21980       .opgAccountSource_optName_lbl2.FontBold = True
21990       .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
22000     End Select
22010     DoEvents
22020     If strAccountNo <> vbNullString Then
22030       .cmbAccounts = strAccountNo
22040     End If
22050   End With

EXITP:
22060   Exit Sub

ERRH:
22070   Select Case ERR.Number
        Case Else
22080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22090   End Select
22100   Resume EXITP

End Sub

Public Sub OptGroup_Handler_IE(strProc As String, frm As Access.Form)

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "OptGroup_Handler_IE"

        Dim strAction As String, strCtl As String
        Dim intPos01 As Integer, lngCnt As Long

22210   With frm

22220     lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
22230     intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
22240     strAction = Mid(strProc, (intPos01 + 1))
22250     strCtl = Left(strProc, (intPos01 - 1))

22260     Select Case strAction
          Case "AfterUpdate"
22270       Select Case strCtl
            Case "opgUnspecified"
22280         .opgUnspecified_optExclude_lbl.FontBold = False
22290         .opgUnspecified_optExclude_lbl.ForeColor = CLR_VDKGRY
22300         .opgUnspecified_optInclude_lbl.FontBold = False
22310         .opgUnspecified_optInclude_lbl.ForeColor = CLR_VDKGRY
22320         .opgUnspecified_optOnly_lbl.FontBold = False
22330         .opgUnspecified_optOnly_lbl.ForeColor = CLR_VDKGRY
22340         Select Case .opgUnspecified
              Case .opgUnspecified_optExclude.OptionValue
22350           .opgUnspecified_optExclude_lbl.FontBold = True
22360           .opgUnspecified_optExclude_lbl.ForeColor = CLR_VDKGRY  ' ** Leave Exclude less prominent.
22370           .opgUnspecified_lbl.FontBold = False
22380           .opgUnspecified_lbl.ForeColor = CLR_VDKGRY
22390         Case .opgUnspecified_optInclude.OptionValue
22400           .opgUnspecified_optInclude_lbl.FontBold = True
22410           .opgUnspecified_optInclude_lbl.ForeColor = CLR_BLK
22420           .opgUnspecified_lbl.FontBold = True
22430           .opgUnspecified_lbl.ForeColor = CLR_BLK
22440         Case .opgUnspecified_optOnly.OptionValue
22450           .opgUnspecified_optOnly_lbl.FontBold = True
22460           .opgUnspecified_optOnly_lbl.ForeColor = CLR_BLK
22470           .opgUnspecified_lbl.FontBold = True
22480           .opgUnspecified_lbl.ForeColor = CLR_BLK
22490         End Select
22500         OptionsChk_IE frm  ' ** Procedure: Above.
22510       Case "opgPrincipalCash"
22520         .opgPrincipalCash_optExclude_lbl.FontBold = False
22530         .opgPrincipalCash_optExclude_lbl.ForeColor = CLR_VDKGRY
22540         .opgPrincipalCash_optInclude_lbl.FontBold = False
22550         .opgPrincipalCash_optInclude_lbl.ForeColor = CLR_VDKGRY
22560         .opgPrincipalCash_optOnly_lbl.FontBold = False
22570         .opgPrincipalCash_optOnly_lbl.ForeColor = CLR_VDKGRY
22580         Select Case .opgPrincipalCash
              Case .opgPrincipalCash_optExclude.OptionValue
22590           .opgPrincipalCash_optExclude_lbl.FontBold = True
22600           .opgPrincipalCash_optExclude_lbl.ForeColor = CLR_VDKGRY  ' ** Leave Exclude less prominent.
22610           .opgPrincipalCash_lbl.FontBold = False
22620           .opgPrincipalCash_lbl.ForeColor = CLR_VDKGRY
22630         Case .opgPrincipalCash_optInclude.OptionValue
22640           .opgPrincipalCash_optInclude_lbl.FontBold = True
22650           .opgPrincipalCash_optInclude_lbl.ForeColor = CLR_BLK
22660           .opgPrincipalCash_lbl.FontBold = True
22670           .opgPrincipalCash_lbl.ForeColor = CLR_BLK
22680         Case .opgPrincipalCash_optOnly.OptionValue
22690           .opgPrincipalCash_optOnly_lbl.FontBold = True
22700           .opgPrincipalCash_optOnly_lbl.ForeColor = CLR_BLK
22710           .opgPrincipalCash_lbl.FontBold = True
22720           .opgPrincipalCash_lbl.ForeColor = CLR_BLK
22730         End Select
22740         OptionsChk_IE frm  ' ** Procedure: Above.
22750       Case "opgZeroCash"
22760         .opgZeroCash_optExclude_lbl.FontBold = False
22770         .opgZeroCash_optExclude_lbl.ForeColor = CLR_VDKGRY
22780         .opgZeroCash_optInclude_lbl.FontBold = False
22790         .opgZeroCash_optInclude_lbl.ForeColor = CLR_VDKGRY
22800         .opgZeroCash_optOnly_lbl.FontBold = False
22810         .opgZeroCash_optOnly_lbl.ForeColor = CLR_VDKGRY
22820         Select Case .opgZeroCash
              Case .opgZeroCash_optExclude.OptionValue
22830           .opgZeroCash_optExclude_lbl.FontBold = True
22840           .opgZeroCash_optExclude_lbl.ForeColor = CLR_VDKGRY  ' ** Leave Exclude less prominent.
22850           .opgZeroCash_lbl.FontBold = False
22860           .opgZeroCash_lbl.ForeColor = CLR_VDKGRY
22870         Case .opgZeroCash_optInclude.OptionValue
22880           .opgZeroCash_optInclude_lbl.FontBold = True
22890           .opgZeroCash_optInclude_lbl.ForeColor = CLR_BLK
22900           .opgZeroCash_lbl.FontBold = True
22910           .opgZeroCash_lbl.ForeColor = CLR_BLK
22920         Case .opgZeroCash_optOnly.OptionValue
22930           .opgZeroCash_optOnly_lbl.FontBold = True
22940           .opgZeroCash_optOnly_lbl.ForeColor = CLR_BLK
22950           .opgZeroCash_lbl.FontBold = True
22960           .opgZeroCash_lbl.ForeColor = CLR_BLK
22970         End Select
22980         OptionsChk_IE frm  ' ** Procedure: Above.
22990       Case "opgSummary"
23000         .opgSummary_optExclude_lbl.FontBold = False
23010         .opgSummary_optExclude_lbl.ForeColor = CLR_VDKGRY
23020         .opgSummary_optInclude_lbl.FontBold = False
23030         .opgSummary_optInclude_lbl.ForeColor = CLR_VDKGRY
23040         .opgSummary_optOnly_lbl.FontBold = False
23050         .opgSummary_optOnly_lbl.ForeColor = CLR_VDKGRY
23060         Select Case .opgSummary
              Case .opgSummary_optExclude.OptionValue
23070           .opgSummary_optExclude_lbl.FontBold = True
23080           .opgSummary_optExclude_lbl.ForeColor = CLR_VDKGRY  ' ** Leave Exclude less prominent.
23090           .opgSummary_lbl.FontBold = False
23100           .opgSummary_lbl.ForeColor = CLR_VDKGRY
23110           .chkDetail.Enabled = True
23120           .chkDontCombineMulti.Enabled = True
23130         Case .opgSummary_optInclude.OptionValue
23140           .opgSummary_optInclude_lbl.FontBold = True
23150           .opgSummary_optInclude_lbl.ForeColor = CLR_BLK
23160           .opgSummary_lbl.FontBold = True
23170           .opgSummary_lbl.ForeColor = CLR_BLK
23180           .chkDetail.Enabled = True
23190           .chkDontCombineMulti.Enabled = True
23200         Case .opgSummary_optOnly.OptionValue
23210           .opgSummary_optOnly_lbl.FontBold = True
23220           .opgSummary_optOnly_lbl.ForeColor = CLR_BLK
23230           .opgSummary_lbl.FontBold = True
23240           .opgSummary_lbl.ForeColor = CLR_BLK
23250           .chkDetail.Enabled = False  ' ** Since the Summary is unaffected by it.
23260           .chkDontCombineMulti.Enabled = False  ' ** Since the Summary is unaffected by it.
23270         End Select
23280         OptionsChk_IE frm  ' ** Procedure: Above.
23290       End Select
23300     End Select

23310   End With

EXITP:
23320   Exit Sub

ERRH:
23330   Select Case ERR.Number
        Case Else
23340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23350   End Select
23360   Resume EXITP

End Sub

Public Sub Reset_Handler_ID(strProc As String, blnReset_Focus As Boolean, blnReset_MouseDown As Boolean, frm As Access.Form)

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "Reset_Handler_ID"

        Dim strAction As String, strCtl As String
        Dim intPos01 As Integer, lngCnt As Long

23410   With frm

23420     lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
23430     intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
23440     strAction = Mid(strProc, (intPos01 + 1))
23450     strCtl = Left(strProc, (intPos01 - 1))

23460     Select Case strAction
          Case "GotFocus"
23470       blnReset_Focus = True
23480       .cmdReset_raised_semifocus_dots_img.Visible = True
23490       .cmdReset_raised_img.Visible = False
23500       .cmdReset_raised_focus_img.Visible = False
23510       .cmdReset_raised_focus_dots_img.Visible = False
23520       .cmdReset_sunken_focus_dots_img.Visible = False
23530       .cmdReset_raised_img_dis.Visible = False
23540     Case "LostFocus"
23550       .cmdReset_raised_img.Visible = True
23560       .cmdReset_raised_semifocus_dots_img.Visible = False
23570       .cmdReset_raised_focus_img.Visible = False
23580       .cmdReset_raised_focus_dots_img.Visible = False
23590       .cmdReset_sunken_focus_dots_img.Visible = False
23600       .cmdReset_raised_img_dis.Visible = False
23610       blnReset_Focus = False
23620     Case "MouseDown"
23630       blnReset_MouseDown = True
23640       .cmdReset_sunken_focus_dots_img.Visible = True
23650       .cmdReset_raised_img.Visible = False
23660       .cmdReset_raised_semifocus_dots_img.Visible = False
23670       .cmdReset_raised_focus_img.Visible = False
23680       .cmdReset_raised_focus_dots_img.Visible = False
23690       .cmdReset_raised_img_dis.Visible = False
23700     Case "MouseUp"
23710       .cmdReset_raised_focus_dots_img.Visible = True
23720       .cmdReset_raised_img.Visible = False
23730       .cmdReset_raised_semifocus_dots_img.Visible = False
23740       .cmdReset_raised_focus_img.Visible = False
23750       .cmdReset_sunken_focus_dots_img.Visible = False
23760       .cmdReset_raised_img_dis.Visible = False
23770       blnReset_MouseDown = False
23780     Case "MouseMove"
23790       If blnReset_MouseDown = False Then
23800         Select Case blnReset_Focus
              Case True
23810           .cmdReset_raised_focus_dots_img.Visible = True
23820           .cmdReset_raised_focus_img.Visible = False
23830         Case False
23840           .cmdReset_raised_focus_img.Visible = True
23850           .cmdReset_raised_focus_dots_img.Visible = False
23860         End Select
23870         .cmdReset_raised_img.Visible = False
23880         .cmdReset_raised_semifocus_dots_img.Visible = False
23890         .cmdReset_sunken_focus_dots_img.Visible = False
23900         .cmdReset_raised_img_dis.Visible = False
23910       End If
23920     Case "Click"
23930       .opgUnspecified = .opgUnspecified.DefaultValue
23940       .opgUnspecified_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
23950       .opgPrincipalCash = .opgPrincipalCash.DefaultValue
23960       .opgPrincipalCash_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
23970       .opgZeroCash = .opgZeroCash.DefaultValue
23980       .opgZeroCash_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
23990       .opgSummary = .opgSummary.DefaultValue
24000       .opgSummary_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
24010       .chkDontCombineMulti = False
24020       .chkDontCombineMulti_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
24030       .chkAcctEveryLine = False
24040       .chkAcctEveryLine_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
24050       .chkSweepOnly = False
24060       .chkSweepOnly_AfterUpdate  ' ** Form Procedure: frmRpt_IncomeExpense.
24070       .cmdMoreOptions.SetFocus
24080       DoEvents
24090       .cmdReset.Enabled = False
24100       .cmdReset_raised_img_dis.Visible = True
24110       .cmdReset_raised_img.Visible = False
24120       .cmdReset_raised_semifocus_dots_img.Visible = False
24130       .cmdReset_raised_focus_img.Visible = False
24140       .cmdReset_raised_focus_dots_img.Visible = False
24150       .cmdReset_sunken_focus_dots_img.Visible = False
24160     End Select

24170   End With

EXITP:
24180   Exit Sub

ERRH:
24190   Select Case ERR.Number
        Case Else
24200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24210   End Select
24220   Resume EXITP

End Sub

Public Sub Prev_Click_IE(strProc As String, blnPrintBoth As Boolean, frm As Access.Form)

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "Prev_Click_IE"

        Dim strAction As String, strCtl As String
        Dim intPos01 As Integer, lngCnt As Long
        Dim strDocName As String

24310   With frm

24320     lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
24330     intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
24340     strAction = Mid(strProc, (intPos01 + 1))
24350     strCtl = Left(strProc, (intPos01 - 1))

24360     Select Case strAction
          Case "Click"
24370       Select Case strCtl
            Case "cmdRevIncExp_IncomePreview"
24380         DoCmd.Hourglass True
24390         DoEvents
24400         If .DoReport = True Then  ' ** Form Function: frmRpt_IncomeExpense.
24410           Select Case .opgSummary
                Case .opgSummary_optOnly.OptionValue
24420             strDocName = "rptIncExp_Income_Summary"
24430             DoCmd.OpenReport strDocName, acViewPreview
24440             DoCmd.Maximize
24450             DoCmd.RunCommand acCmdFitToWindow
24460           Case Else
24470             Select Case .chkDetail
                  Case True
24480               strDocName = "rptIncExp_Income_Detailed"
24490             Case False
24500               strDocName = "rptIncExp_Income"
24510             End Select
24520             DoCmd.OpenReport strDocName, acViewPreview
24530             Select Case .opgSummary
                  Case .opgSummary_optInclude.OptionValue  ' ** Don't maximize if both are being previewed.
24540               DoEvents
24550               strDocName = "rptIncExp_Income_Summary"
24560               DoCmd.OpenReport strDocName, acViewPreview
24570             Case .opgSummary_optExclude.OptionValue
24580               DoCmd.Maximize
24590               DoCmd.RunCommand acCmdFitToWindow
24600             End Select
24610           End Select
24620         End If
24630         If blnPrintBoth = False Then
24640           DoCmd.Hourglass False
24650         End If
24660       Case "cmdRevIncExp_ExpensePreview"
24670         DoCmd.Hourglass True
24680         DoEvents
24690         If .DoReport = True Then  ' ** Form Function: frmRpt_IncomeExpense.
24700           Select Case .opgSummary
                Case .opgSummary_optOnly.OptionValue
24710             strDocName = "rptIncExp_Expenses_Summary"
24720             DoCmd.OpenReport strDocName, acViewPreview
24730             DoCmd.Maximize
24740             DoCmd.RunCommand acCmdFitToWindow
24750           Case Else
24760             Select Case .chkDetail
                  Case True
24770               strDocName = "rptIncExp_Expenses_Detailed"
24780             Case False
24790               strDocName = "rptIncExp_Expenses"
24800             End Select
24810             DoCmd.OpenReport strDocName, acViewPreview
24820             Select Case .opgSummary
                  Case .opgSummary_optInclude.OptionValue  ' ** Don't maximize if both are being previewed.
24830               DoEvents
24840               strDocName = "rptIncExp_Expenses_Summary"
24850               DoCmd.OpenReport strDocName, acViewPreview
24860             Case .opgSummary_optExclude.OptionValue
24870               DoCmd.Maximize
24880               DoCmd.RunCommand acCmdFitToWindow
24890             End Select
24900           End Select
24910         End If
24920         If blnPrintBoth = False Then
24930           DoCmd.Hourglass False
24940         End If
24950       End Select
24960     End Select

24970   End With

EXITP:
24980   Exit Sub

ERRH:
24990   Select Case ERR.Number
        Case Else
25000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25010   End Select
25020   Resume EXITP

End Sub

Public Sub Remember_After_ID(strProc As String, frm As Access.Form)

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "Remember_After_ID"

        Dim strAction As String, strCtl As String
        Dim intPos01 As Integer, lngCnt As Long

25110   With frm

25120     lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
25130     intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
25140     strAction = Mid(strProc, (intPos01 + 1))
25150     strCtl = Left(strProc, (intPos01 - 1))

25160     Select Case strAction
          Case "AfterUpdate"
25170       Select Case strCtl
            Case "chkRememberDates"
25180         Select Case .chkRememberDates
              Case True
25190           .chkRememberDates_lbl.FontBold = True
25200         Case False
25210           .chkRememberDates_lbl.FontBold = False
25220         End Select
25230       Case "chkRememberMe"
25240         Select Case .chkRememberMe
              Case True
25250           .chkRememberMe_lbl.FontBold = True
25260           .chkRememberMe_lbl2_dim.FontBold = True
25270           .chkRememberMe_lbl2_dim_hi.FontBold = True
25280         Case False
25290           .chkRememberMe_lbl.FontBold = False
25300           .chkRememberMe_lbl2_dim.FontBold = False
25310           .chkRememberMe_lbl2_dim_hi.FontBold = False
25320         End Select
25330       End Select
25340     End Select

25350   End With

EXITP:
25360   Exit Sub

ERRH:
25370   Select Case ERR.Number
        Case Else
25380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25390   End Select
25400   Resume EXITP

End Sub
