VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStatementParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmStatementParameters"

'VGC 10/09/2017: CHANGES!
' ** TimerInterval set to 500& for GTR from frmMenu_Reports.

' ** For GTR, sets options and arrows, then stops, or invokes Statement Balance.

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' #########################
' ## Use VBA_RenumErrh().  37420
' #########################

' ** Shortcut Alt keys responsive from this form:
' **   Select Account:   A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Asset Word:       D {cmdAssetListWord}
' **   Trans Excel:      E {cmdTransactionsExcel}
' **   From:             F {TransDateStart}
' **   Legal Name:       G {chkLegalName}
' **   Include Archive:  H {chkIncludeArchive_Trans}
' **   Asset List:       I {chkAssetList}
' **   Statement Month:  J {cmbMonth}
' **   Statement Year:   K {StatementsYear}
' **   All Accts:        L {opgAccountNumber_optAll}
' **   Name:             M {opgAccountSource_optName}
' **   Asset Print:      N {cmdAssetListPrint}
' **   Trans Word:       O {cmdTransactionsWord}
' **   Trans Print:      P {cmdTransactionsPrint}
' **   As Of Date:       Q {AssetListDate}
' **   Transactions:     R {chkTransactions}
' **   Statements:       S {chkStatements}
' **   To:               T {TransDateEnd}
' **   Specified Acct:   U {opgAccountNumber_optSpecified}
' **   Asset Preview:    V {cmdAssetListPreview}
' **   Trans Preview:    W {cmdTransactionsPreview}
' **   Exit:             X {cmdClose}
' **   Close Reports:    Y {cmdDevCloseReports}
' **   Asset Excel:      Z {cmdAssetListExcel}
'Available letters:
'{none}

' ** Shortcut Alt-Shift keys responsive from this form:
' **   Archived Only:    H {chkArchiveOnly_Trans}

' ** Shortcut Ctrl keys responsive from this form:
' **   Print All:        A {cmdPrintStatement_All}
' **   Balance:          B {cmdBalanceTable}
' **   Order By Date:    D {opgOrderBy_optDate}
' **   Page Count:       G {chkPageOf}
' **   Annual Statement: N {cmdAnnualStatement}
' **   Combine:          O {chkCombineCash}
' **   Include Related:  R {chkRelatedAccounts}
' **   Reprint Single:   S {cmdPrintStatement_Single}
' **   Order By Type:    T {opgOrderBy_optType}
' **   Include Currency: U {chkIncludeCurrency}
' **   Save Data:        X {chkDevSaveData}
' **   Acct Summary:     Y {cmdPrintStatement_Summary}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}

' ***********************************************
' ** Easter Egg:
' ** cmdTransactionsPreview_MouseUp
' ***********************************************

' ** MsgBox() numbering:
' ** Form_Error(): "FEnn"
' ** All Others: "[A-Z]nn" for each different procedure.

' ** VGC 21/27/2012: chkIncludeArchive_Asset no longer optional.

' ** Liability Criteria:
' **   Liability with '-' (negative) icash or pcash: Expense : Purchase
' **   Liability with '+' (positive) icash or pcash: Income  : Sold

' ** Temporary tables used:
' **   AssetList      : USED IN THIS FORM, AND ?
' **   tmpAssetList1  : USED IN THIS FORM
' **   tmpAssetList2  : USED IN THIS FORM, AND ?
' **   tmpAssetList3  : USED BY COURT REPORTS
' **   tmpAccountInfo : USED IN THIS FORM, AND ?

' ** To test Account Summary report:
'gstrReportQuerySpec = "rptAccountSummary"
'gstrFormQuerySpec = "frmStatementParameters"
'gstrAccountNo = "50"
'gdatStartDate = #01/01/2015#
'gdatEndDate = #12/31/2015#
'gdatMarketDate = #12/31/2014#
'glngMonthID = 12&

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** cmbAccounts combo box constants:
Private Const CBX_A_ACTNO  As Integer = 0  ' ** accountno
'Private Const CBX_A_DESC   As Integer = 1  ' ** Desc
'Private Const CBX_A_PREDAT As Integer = 2  ' ** predate
'Private Const CBX_A_SHORT  As Integer = 3  ' ** shortname
'Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname
'Private Const CBX_A_BALDAT As Integer = 5  ' ** BalanceDate (earliest [balance date])
'Private Const CBX_A_HASREL As Integer = 6  ' ** HasRelated
'Private Const CBX_A_CASNUM As Integer = 7  ' ** CaseNum
'Private Const CBX_A_TRXDAT As Integer = 8  ' ** TransDate (earliest [transdate])

' ** cmbMonth combo box constants:
'Private Const CBX_MON_ID    As Integer = 0  ' ** month_id (same as month number)
Private Const CBX_MON_NAME  As Integer = 1  ' ** month_name
'Private Const CBX_MON_SHORT As Integer = 2  ' ** month_short

' ** Array: arr_varAcctArch().
Private lngAcctArchs As Long, arr_varAcctArch As Variant
'Private Const AR_ACTNO As Integer = 0
'Private Const AR_TDATE As Integer = 1
'Private Const AR_CNT   As Integer = 2

' ** Array: arr_varAcctFor().
Private lngAcctFors As Long, arr_varAcctFor As Variant
'Private Const F_ACTNO As Integer = 0
'Private Const F_JCNT  As Integer = 1
'Private Const F_ACNT  As Integer = 2
'Private Const F_SUPP  As Integer = 3

' ** Array: arr_varStmt().
Private lngStmts As Long, lngStmtCnt As Long, arr_varStmt As Variant
'Private Const S_ELEMS1 As Integer = 12  ' ** Array's first-element UBound().
'Private Const S_ELEMS2 As Integer = 4   ' ** Array's second-element UBound().
'Private Const S_MID   As Integer = 0  'month_id
'Private Const S_MSHT  As Integer = 1  'month_short
'Private Const S_CNT   As Integer = 2  'cnt_smt
'Private Const S_ACTNO As Integer = 3  'accountno
'Private Const S_SNAM  As Integer = 4  'shortname

Private blnAnnualStatement_Focus As Boolean, blnAnnualStatement_MouseDown As Boolean
Private blnBalanceTable_Focus As Boolean, blnBalanceTable_MouseDown As Boolean
Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean
Private blnCalendar3_Focus As Boolean, blnCalendar3_MouseDown As Boolean

Private blnResetStmts As Boolean, blnLoop As Boolean
Private datFirstDate As Date, strFirstDateMsg As String, datAssetListDate As Date, datAssetListDate_Pref As Date
Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long, blnWasGTR As Boolean
Private strRememberMe As String, blnHasForEx As Boolean, blnHasForExThis As Boolean, blnAfterFired As Boolean, blnHasForExClick As Boolean
Private blnIsOpen As Boolean, blnContinue As Boolean, blnAcctNotSched As Boolean, blnIncludeCurrency As Boolean, blnFromStmts As Boolean
Private strReportName As String, strFileName As String  ' ** String containing the file name when printing to a file.
Private blnRunPriorStatement As Boolean, blnPrintAnnualStatement As Boolean, blnAcctsSchedRpt As Boolean, blnRollbackNeeded As Boolean
Private blnAllStatements As Boolean, blnSingleStatement As Boolean, blnPrintStatements As Boolean, blnNoDataAll As Boolean
Private strCallingForm As String, lngTpp As Long
Private THAT_PROC As String, That_Erl As Long, That_Desc As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim varTmp00 As Variant

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnIsOpen = True: blnHasForEx = False: blnHasForExThis = False: blnHasForExClick = False
160           gstrFormQuerySpec = THIS_NAME
170           gstrReportQuerySpec = vbNullString
180           strRememberMe = vbNullString
190           gvarCrtRpt_FL_SpecData = Null

200           blnAnnualStatement_Focus = False: blnAnnualStatement_MouseDown = False
210           blnBalanceTable_Focus = False: blnBalanceTable_MouseDown = False
220           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
230           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False
240           blnCalendar3_Focus = False: blnCalendar3_MouseDown = False

250           gblnLegalName = False: gblnForeignCurrencies = False: blnAcctNotSched = False: blnIncludeCurrency = False
260           blnAllStatements = False: blnSingleStatement = False: blnRunPriorStatement = False: blnPrintAnnualStatement = False
270           blnNoDataAll = False: blnAfterFired = False: blnFromStmts = False: blnResetStmts = False: blnLoop = False
280           blnRollbackNeeded = False: gblnMessage = False  ' ** Used to signal that Rollbacks were necessary.
290           blnGTR_Emblem = False: lngGTR_Stat = 0&
300           datFirstDate = DateAdd("y", 1, Date)  ' ** Set this to tomorrow.
310           blnWasGTR = gblnGoToReport
320           lngStmts = 0&: lngStmtCnt = 0&

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
330           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

340           CLR_DISABLED_FG = CLR_DKGRY
350           CLR_DISABLED_BG = CLR_LTTEAL

360           ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

370           If .PopUp = True And gblnGoToReport = False Then
380             .DevMsg_lbl.Visible = True
390             gblnSwitchTo = False  ' ** Borrowing this variable from Journal/Transactions.
400           End If

              'If CurrentUser = "Superuser" Then
              '  .chkDevSaveData.Visible = True
              '  .chkNoAssets.Visible = True
              '  .chkNoAssets_All.Visible = True
              'Else
              '  .chkDevSaveData.Visible = False
              '  .chkNoAssets.Visible = False
              '  .chkNoAssets_All.Visible = False
              'End If

              ' ** Try to prevent that odd defaulting to yesterday!
410           .DateEnd = Null  ' ** Invisible blue field in header.
420           .AssetListDate = Null
430           If .AssetListDate.Enabled = False Then .AssetListDate.Enabled = True
440   On Error Resume Next
450           .AssetListDate.SetFocus
460           .AssetListDate.text = vbNullString
470   On Error GoTo ERRH
480           .cmdClose.SetFocus

490           .cmbMonth = Null
500           .StatementsYear = Null
510           DoEvents

520           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

530           If IsNull(.DateEnd) = True Then
540             .DateEnd = TodaysDate  ' ** Module Function: modUtilities.
550           End If
560           Select Case IsNull(.AssetListDate)
              Case True
570             .AssetListDate = TodaysDate  ' ** Module Function: modUtilities.
580           Case False
590             datAssetListDate_Pref = .AssetListDate
600           End Select
610           If IsNull(.StatementsYear) = True Then
620             .StatementsYear = year(DateAdd("m", -1, Now()))
630           End If
640           DoEvents

650           Set dbs = CurrentDb

              ' ** Get a list of accounts with archive transactions.
660           Set qdf = dbs.QueryDefs("qryArchiveTransactions_05")
670           Set rst = qdf.OpenRecordset
680           With rst
690             If .BOF = True And .EOF = True Then
                  ' ** They've not used TrstArch.mdb.
700               lngAcctArchs = 0&
710             Else
720               .MoveLast
730               lngAcctArchs = .RecordCount
740               .MoveFirst
750               arr_varAcctArch = .GetRows(lngAcctArchs)
                  ' ***********************************************
                  ' ** Array: arr_varAcctArch()
                  ' **
                  ' **   Field  Element  Name         Constant
                  ' **   =====  =======  ===========  ===========
                  ' **     1       0     accountno    AR_ACTNO
                  ' **     2       1     transdate    AR_TDATE
                  ' **     3       2     cnt          AR_CNT
                  ' **
                  ' ***********************************************
760             End If
770             .Close
780           End With  ' ** rst.
790           Set rst = Nothing
800           Set qdf = Nothing

810           If IsNull(.cmbAccounts) = False Then
820             If Trim(.cmbAccounts) <> vbNullString Then
830               strRememberMe = .cmbAccounts
840             End If
850           End If

              ' ** MasterAsset, grouped, with Max(currentDate).
860           Set qdf = dbs.QueryDefs("qryStatementParameters_07")
870           Set rst = qdf.OpenRecordset
880           rst.MoveFirst
890           .currentDate = rst![currentDate]
900           rst.Close
910           Set rst = Nothing
920           Set qdf = Nothing

930           dbs.Close
940           Set dbs = Nothing

950           .chkStatements_lbl3.Caption = Format(DLookup("[Statement_Date]", "Statement Date"), "mm/dd/yyyy")

960           .chkAssetList = False
970           .chkIncludeArchive_Asset.Enabled = False
980           .chkStatements = False
990           .cmdAnnualStatement.Enabled = False
1000          .cmdAnnualStatement_raised_img_dis.Visible = True
1010          .cmdAnnualStatement_raised_img.Visible = False
1020          .cmdAnnualStatement_raised_semifocus_dots_img.Visible = False
1030          .cmdAnnualStatement_raised_focus_img.Visible = False
1040          .cmdAnnualStatement_raised_focus_dots_img.Visible = False
1050          .cmdAnnualStatement_sunken_focus_dots_img.Visible = False
1060          .chkTransactions = True

      #If NoExcel Then
1070          .cmdTransactionsExcel.Enabled = False
1080          .cmdAssetListExcel.Enabled = False
      #End If

              'ALL OF THESE ARE TAKING TIME!
              ' ** EVENT CHECK: chkRememberMe!
1090          cmbAccounts_AfterUpdate  ' ** Procedure: Below.
              'Calls CmbAccts_After_SP()
              ' ** EVENT CHECK: chkRememberDates!
1100          If IsNull(.TransDateStart) = False And IsNull(.TransDateEnd) = False Then
1110            TransDateStart_Exit Cancel  ' ** Procedure: Below.
                'Sets Dates.
1120            TransDateEnd_Exit Cancel  ' ** Procedure: Below.
                'Sets Dates.
1130          End If
1140          AssetListDate_AfterUpdate  ' ** Procedure: Below.
              'Sets Dates.

1150          chkTransactions_AfterUpdate  ' ** Procedure: Below.
              'Calls ChkTrans_After_SP()
1160          opgAccountSource_AfterUpdate  ' ** Procedure: Below.
              'Sets Bold.
1170          chkRememberMe_AfterUpdate  ' ** Procedure: Below.
              'Sets Bold.
1180          chkRememberDates_Trans_AfterUpdate  ' ** Procedure: Below.
              'Sets Bold.
1190          chkLegalName_AfterUpdate  ' ** Procedure: Below.
              'Sets Bold.
1200          chkForeignExchange_AfterUpdate  ' ** Procedure: Below.
              'Sets Bold.
1210          chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
              'Sets Bold.
1220          opgOrderBy_AfterUpdate  ' ** Procedure: Below.
              'Sets Bold.
1230          chkRelatedAccounts_AfterUpdate  ' ** Procedure: Below.
              'Sets Bold.
1240          chkCombineCash_AfterUpdate  ' ** Procedure: Below.
              'Sets Bold.
1250          chkIncludeArchive_Trans_AfterUpdate  ' ** Procedure: Below.
              'Calls chkArchiveOnly_Trans_AfterUpdate()
1260          chkArchiveOnly_Trans_AfterUpdate  ' ** Procedure: Below.
              'Calls chkIncludeArchive_Trans_AfterUpdate()
              'Calls opgAccountNumber_AfterUpdate
1270          chkIncludeArchive_Asset_AfterUpdate  ' ** Procedure: Below.
              'Sets Bold.
1280          chkPageOf_AfterUpdate  ' ** Procedure: Below.
              'Sets Bold.
1290          DoEvents
1300          blnIsOpen = False

1310        Else
1320          Cancel = -1
1330        End If
1340      Else
1350        Cancel = -1
1360      End If
1370    End With

1380    If Cancel = -1 Then
1390      Beep
1400      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1410    End If

1420    DoCmd.Hourglass False

EXITP:
1430    Set rst = Nothing
1440    Set qdf = Nothing
1450    Set dbs = Nothing
1460    Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub Form_Load()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1510    With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
1520      Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
1530      clsMonthClass.hWndForm = .hwnd
          ' ** Let's default to PositionAtCursor.
1540      clsMonthClass.PositionAtCursor = True

1550      .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.
          'Runs Qrys.

          ' ** The field will always stay the same, but the variable may change.
1560      blnHasForEx = HasForEx_SP("All")  ' ** Module Function: modStatementParamFuncs1.
          'Runs Qrys.

1570      .HasForeign = blnHasForEx
1580      .HasForeign_Sched = True
1590      Select Case blnHasForEx
          Case True
1600        .chkForeignExchange = True
1610        .chkForeignExchange.Enabled = True
1620        .chkForeignExchange.Locked = True
1630        .chkForeignExchange_lbl.FontBold = True
1640        .chkForeignExchange_lbl.Caption = "Foreign Currencies Present"
1650      Case False
1660        .chkForeignExchange = False
1670        .chkForeignExchange.Enabled = False
1680        .chkForeignExchange.Locked = False
1690        .chkForeignExchange_lbl.FontBold = False
1700        .chkForeignExchange_lbl.Caption = "No Foreign Currencies Found"
1710      End Select
1720      .chkIncludeCurrency.Enabled = False
1730      .chkIncludeCurrency.Locked = False

          ' ** Load the arr_varAcctFor() array.
1740      ForExArr_Load  ' ** Module Function: modStatementParamFuncs1.
          'Runs Qrys, loads array.
1750      arr_varAcctFor = ForExArr_Get  ' ** Module Function: modStatementParamFuncs1.
          'Calls ForExArr_Load()
          '  Runs Qrys, loads array.
1760      lngAcctFors = UBound(arr_varAcctFor, 2) + 1&

1770    End With

EXITP:
1780    Exit Sub

ERRH:
100     DoCmd.Hourglass False
110     Select Case ERR.Number
        Case Else
120       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
130     End Select
140     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1820    intShiftDown = (Shift And acShiftMask) > 0
1830    intAltDown = (Shift And acAltMask) > 0
1840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality.
        ' **   Include Archive:  H {chkIncludeArchive_Trans}
        ' **   Statement Month:  J {cmbMonth}
        ' **   Statement Year:   K {StatementsYear}
        ' **   As Of Date:       Q {AssetListDate}
        ' **   Close Reports:    Y {cmdDevCloseReports}

        ' ** Shortcut Alt-Shift keys to other forms and functionality.
        ' **   Archived Only:    H {chkArchiveOnly_Trans}

        ' ** Shortcut Ctrl keys to other forms and functionality.
        ' **   Print All:        A {cmdPrintStatement_All}
        ' **   Balance:          B {cmdBalanceTable}
        ' **   Order By Date:    D {opgOrderBy_optDate}
        ' **   Page Count:       G {chkPageOf}
        ' **   Annual Statement: N {cmdAnnualStatement}
        ' **   Combine:          O {chkCombineCash}
        ' **   Include Related:  R {chkRelatedAccounts}
        ' **   Reprint Single:   S {cmdPrintStatement_Single}
        ' **   Order By Type:    T {opgOrderBy_optType}
        ' **   Include Currency: U {chkIncludeCurrency}
        ' **   Save Data:        X {chkDevSaveData}
        ' **   Acct Summary:     Y {cmdPrintStatement_Summary}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality.
        ' **   Close Timer:      T {Form_KeyDown}

        ' ** Plain keys.
1850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1860      Select Case intRetVal
          Case vbKeyEscape
1870        intRetVal = 0
1880        cmdClose_Click  ' ** Procedure: Below.
1890      End Select
1900    End If

        ' ** Alt Keys.
1910    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1920      Select Case intRetVal
          Case vbKeyH
1930        With Me
1940          intRetVal = 0
1950          If .chkIncludeArchive_Trans.Enabled = True Then
1960            .chkIncludeArchive_Trans.SetFocus
1970            .chkIncludeArchive_Trans = (Not .chkIncludeArchive_Trans)
1980            chkIncludeArchive_Trans_AfterUpdate  ' ** Procedure: Below.
1990          Else
2000            Beep
2010          End If
2020        End With
2030      Case vbKeyJ
2040        With Me
2050          intRetVal = 0
2060          .cmbMonth.SetFocus
2070        End With
2080      Case vbKeyK
2090        With Me
2100          intRetVal = 0
2110          .StatementsYear.SetFocus
2120        End With
2130      Case vbKeyQ
2140        With Me
2150          intRetVal = 0
2160          .AssetListDate.SetFocus
2170        End With
2180      Case vbKeyX
2190        intRetVal = 0
2200        cmdClose_Click  ' ** Procedure: Below.
2210      Case vbKeyY
2220        With Me
2230          intRetVal = 0
2240          If .cmdDevCloseReports.Visible = True And .cmdDevCloseReports.Enabled = True Then
2250            .cmdDevCloseReports.SetFocus
2260            cmdDevCloseReports_Click  ' ** Procedure: Below.
2270          End If
2280        End With
2290      Case vbKeyZ
2300        With Me
2310          intRetVal = 0
2320          If .cmdAssetListExcel.Enabled = True Then
2330            .cmdAssetListExcel.SetFocus  '#chk
2340            cmdAssetListExcel_Click  ' ** Procedure: Below.
2350          Else
2360            Beep
2370          End If
2380        End With
2390      End Select
2400    End If

        ' ** Alt-Shift keys.
2410    If (Not intCtrlDown) And intAltDown And intShiftDown Then
2420      Select Case intRetVal
          Case vbKeyH
2430        With Me
2440          intRetVal = 0
2450          If .chkArchiveOnly_Trans.Enabled = True Then
2460            .chkArchiveOnly_Trans.SetFocus
2470            .chkArchiveOnly_Trans = (Not .chkArchiveOnly_Trans)
2480            chkArchiveOnly_Trans_AfterUpdate  ' ** Procedure: Below.
2490          Else
2500            Beep
2510          End If
2520        End With
2530      End Select
2540    End If

        ' ** Ctrl keys.
2550    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2560      Select Case intRetVal
          Case vbKeyA
2570        With Me
2580          intRetVal = 0
2590          .cmdPrintStatement_All.SetFocus
2600          cmdPrintStatement_All_Click  ' ** Procedure: Below.
2610        End With
2620      Case vbKeyB
2630        With Me
2640          intRetVal = 0
2650          If .cmdBalanceTable.Enabled = True Then
2660            .cmdBalanceTable.SetFocus
2670            cmdBalanceTable_Click  ' ** Procedure: Below.
2680          End If
2690        End With
2700      Case vbKeyD
2710        With Me
2720          intRetVal = 0
2730          If .opgOrderBy.Enabled = True Then
2740            .opgOrderBy.SetFocus
2750            .opgOrderBy = .opgOrderBy_optDate.OptionValue
2760            opgOrderBy_AfterUpdate  ' ** Procedure: Below.
2770          End If
2780        End With
2790      Case vbKeyG
2800        With Me
2810          intRetVal = 0
2820          If .chkPageOf.Enabled = True Then
2830            .chkPageOf.SetFocus
2840            .chkPageOf = (Not .chkPageOf)
2850            chkPageOf_AfterUpdate  ' ** Procedure: Below.
2860          End If
2870        End With
2880      Case vbKeyN
2890        With Me
2900          intRetVal = 0
2910          .cmdAnnualStatement.SetFocus
2920          cmdAnnualStatement_Click  ' ** Procedure: Below.
2930        End With
2940      Case vbKeyO
2950        With Me
2960          intRetVal = 0
2970          .chkCombineCash.SetFocus
2980          .chkCombineCash = Not (.chkCombineCash)
2990          chkCombineCash_AfterUpdate  ' ** Procedure: Below.
3000        End With
3010      Case vbKeyR
3020        With Me
3030          intRetVal = 0
3040          .chkRelatedAccounts.SetFocus
3050          .chkRelatedAccounts = Not (.chkRelatedAccounts)
3060          chkRelatedAccounts_AfterUpdate  ' ** Procedure: Below.
3070        End With
3080      Case vbKeyS
3090        With Me
3100          intRetVal = 0
3110          .cmdPrintStatement_Single.SetFocus
3120          cmdPrintStatement_Single_Click  ' ** Procedure: Below.
3130        End With
3140      Case vbKeyT
3150        With Me
3160          intRetVal = 0
3170          If .opgOrderBy.Enabled = True Then
3180            .opgOrderBy.SetFocus
3190            .opgOrderBy = .opgOrderBy_optType.OptionValue
3200            opgOrderBy_AfterUpdate  ' ** Procedure: Below.
3210          End If
3220        End With
3230      Case vbKeyU
3240        With Me
3250          intRetVal = 0
3260          If .chkIncludeCurrency.Enabled = True Then
3270            .chkIncludeCurrency.SetFocus
3280            .chkIncludeCurrency = (Not .chkIncludeCurrency)
3290            chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
3300          Else
3310            Beep
3320          End If
3330        End With
3340      Case vbKeyY
3350        With Me
3360          intRetVal = 0
3370          .cmdPrintStatement_Summary.SetFocus
3380          cmdPrintStatement_Summary_Click  ' ** Procedure: Below.
3390        End With
3400      End Select
3410    End If

        ' ** Ctrl-Shift keys.
3420    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3430      Select Case intRetVal
          Case vbKeyA  ' ** As Developer.
3440        With Me
3450          intRetVal = 0
3460          If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
3470            Select Case .chkAsDev.Visible
                Case True
3480              .chkAsDev.Visible = False
3490              .chkDevMsg.Visible = False
3500              .cmdDevCloseReports.Visible = False
3510            Case False
3520              .chkAsDev.Visible = True
3530              .chkDevMsg.Visible = True
3540              .cmdDevCloseReports.Visible = True
3550            End Select
3560          End If
3570        End With
3580      Case vbKeyF
3590        With Me
3600          intRetVal = 0
3610          .FocusHolder.SetFocus
3620        End With
3630      Case vbKeyT
3640        intRetVal = 0
3650        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
3660      Case vbKeyX
3670        With Me
3680          intRetVal = 0
3690          If .chkDevSaveData.Visible = True Then
3700            .chkDevSaveData = (Not .chkDevSaveData)
3710            chkDevSaveData_AfterUpdate  ' ** Procedure: Below.
3720          End If
3730        End With
3740      End Select
3750    End If

EXITP:
3760    KeyCode = intRetVal
3770    Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1510    Resume EXITP

End Sub

Private Sub Form_Timer()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

        Static lngTmp01 As Long, lngTmp02 As Long

3810    With Me
3820      .TimerInterval = 0&
3830      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
3840        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
3850      End If
3860      If blnGTR_Emblem = True Then
3870        lngGTR_Stat = lngGTR_Stat + 1&
3880        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
3890        lngTmp01 = lngTmp01 + 1&
3900        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
3910        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
3920        If lngTmp01 > 1& Then
3930          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
3940        End If
3950        If lngTmp01 = 24& Then
3960          lngTmp01 = 0&
3970          lngTmp02 = lngTmp02 + 1&
3980          If lngTmp02 < 11& Then  ' ** Total number of complete spins before collapsing.
3990            .TimerInterval = 50&
4000          Else
4010            blnGTR_Emblem = False
4020          End If
4030        Else
4040          .TimerInterval = 50&
4050        End If
4060      End If  ' ** blnGTR_Emblem.
4070      If gblnSetFocus = True Then
4080        gblnSetFocus = False
4090        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
4100          Forms("frmLoadTimer").Modal = False
4110        End If
4120        strFormName = vbNullString
4130  On Error Resume Next
4140        Do Until strFormName = THIS_NAME
4150          DoCmd.SelectObject acForm, THIS_NAME, False
4160          strFormName = Screen.ActiveForm.Name
4170        Loop
4180  On Error GoTo ERRH
4190      ElseIf gblnReportClose = True Then
4200        gblnReportClose = False
4210        TAReports_Access2007 False  ' ** Module Function: modReportFunctions.
4220      End If
4230      If gblnGoToReport = True Then
4240        DoCmd.Hourglass True  ' ** Make sure it's still running.
4250        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48& And lngGTR_Stat < 84&) Or _
                (lngGTR_Stat > 84& And lngGTR_Stat < 120&) Or (lngGTR_Stat > 120& And lngGTR_Stat < 156&) Or _
                (lngGTR_Stat > 156& And lngGTR_Stat < 192&) Or (lngGTR_Stat > 192&)) Then
              ' ** Just let the emblem go.
4260        Else
4270          If blnGTR_Emblem = False And garr_varGoToReport(GTR_CTL6) <> "cmdTransactionsPrint" And _
                  garr_varGoToReport(GTR_CTL6) <> "cmdAssetListPrint" Then
4280            blnGTR_Emblem = True
4290            lngTmp01 = 0&: lngTmp02 = 0&
4300          End If
4310          DoEvents
4320          Select Case blnGoingToReport
              Case True
4330            Select Case blnGoingToReport2
                Case True
4340              If garr_varGoToReport(GTR_CTL2) = "cmdPrintStatement_All" Then
4350                If .GoToReport_arw_sp_printall_img.Visible = True Then
4360                  .GoToReport_arw_sp_printall_img.Visible = False
4370                  .TimerInterval = 50&
4380                  cmdPrintStatement_All_Click  ' ** Procedure: Below.
4390                ElseIf .GoToReport_arw_sp_optall_img.Visible = True Then
4400                  .cmdPrintStatement_Single.Visible = False
4410                  .GoToReport_arw_sp_printall_img.Visible = True
4420                  .GoToReport_arw_sp_optall_img.Visible = False
4430                  .cmdPrintStatement_All.SetFocus
4440                  DoEvents
4450                  .TimerInterval = 50&
4460                Else
4470                  .opgAccountNumber.SetFocus
4480                  .GoToReport_arw_sp_optall_img.Top = .opgAccountNumber_optAll_lbl.Top
4490                  .GoToReport_arw_sp_optall_img.Left = (.opgAccountNumber_optAll_lbl.Left + .opgAccountNumber_optAll_lbl.Width)
4500                  .GoToReport_arw_sp_optall_img.Visible = True
4510                  .GoToReport_arw_sp_month_img.Visible = False
4520                  DoEvents
4530                  If .opgAccountNumber <> .opgAccountNumber_optAll.OptionValue Then
4540                    .opgAccountNumber = .opgAccountNumber_optAll.OptionValue
4550                    opgAccountNumber_AfterUpdate
4560                    DoEvents
4570                  End If
4580                  .TimerInterval = 50&
4590                End If
4600              Else
4610                blnGoingToReport2 = False
4620                blnGoingToReport = False
4630                gblnGoToReport = False
4640                .opgAccountNumber_lbl.Visible = True
4650                .GoToReport_arw_sp_stmt_img.Visible = False
4660                .GoToReport_lbl_sp_stmt_img.Visible = False
4670                DoCmd.Hourglass False
4680                DoEvents
4690              End If
4700            Case False
4710              If .GoToReport_arw_sp_acct_d_img.Visible = True Or .GoToReport_arw_sp_acct_t_img.Visible = True Then
4720                .cmdBalanceTable.SetFocus
4730                .GoToReport_arw_sp_sbal_img.Top = (.cmdBalanceTable.Top + (2& * lngTpp))
4740                .GoToReport_arw_sp_sbal_img.Left = ((.cmdBalanceTable.Left - .GoToReport_arw_sp_sbal_img.Width) + lngTpp)
4750                .GoToReport_arw_sp_acct_d_img.Visible = False
4760                .GoToReport_arw_sp_acct_t_img.Visible = False
4770                .opgOrderBy.Enabled = True
4780                .GoToReport_arw_sp_sbal_img.Visible = True
4790                DoEvents
4800                .TimerInterval = 50&
4810              ElseIf .GoToReport_arw_sp_sbal_img.Visible = True Then
4820                .TimerInterval = 50&
4830                cmdBalanceTable_Click  ' ** Procedure: Below.
4840              ElseIf .GoToReport_arw_sp_trans_img.Visible = True Then
4850                .GoToReport_arw_sp_trans_img.Visible = False
4860                blnGoingToReport = False
4870                gblnGoToReport = False
4880                DoCmd.Hourglass False
4890              ElseIf .GoToReport_arw_sp_alist_img.Visible = True Then
4900                .GoToReport_arw_sp_alist_img.Visible = False
4910                blnGoingToReport = False
4920                gblnGoToReport = False
4930                DoCmd.Hourglass False
4940              ElseIf garr_varGoToReport(GTR_CTL2) = "cmdPrintStatement_All" Then
4950                blnGoingToReport2 = True
4960                .cmbMonth.SetFocus
4970                .GoToReport_arw_sp_month_img.Top = (.cmbMonth.Top + .cmbMonth.Height)
4980                .GoToReport_arw_sp_month_img.Left = (.cmbMonth.Left + (3& * lngTpp))
4990                .GoToReport_arw_sp_month_img.Visible = True
5000                .GoToReport_arw_sp_stmt_img.Visible = False
5010                .GoToReport_lbl_sp_stmt_img.Visible = False
5020                .opgAccountNumber_lbl.Visible = True
5030                .cmbMonth = "December"    'DO WE NEED THE AFTER UPDATE EVENT CALLED?
5040                .cmbMonth.SelLength = 0
5050                .cmbMonth.SelStart = 19
5060                DoEvents
5070                .TimerInterval = 50&
5080              ElseIf .GoToReport_arw_sp_stmt_img.Visible = True Then
5090                blnGTR_Emblem = False
5100                GTREmblem_Off  ' ** Procedure: Below.
5110                DoEvents
5120                DoCmd.Hourglass False
5130                blnGoingToReport2 = True
5140                .TimerInterval = (GTR_WAIT * 2&)
5150              Else
5160                blnGoingToReport = False
5170                DoCmd.Hourglass False
5180                gblnGoToReport = False
5190                blnGTR_Emblem = False
5200                GTREmblem_Off  ' ** Procedure: Below.
5210                Beep
5220              End If
5230            End Select
5240          Case False
5250            blnGoingToReport = True
5260            Select Case IsNull(garr_varGoToReport(GTR_FRM2))
                Case True
5270              Select Case garr_varGoToReport(GTR_CTL6)
                  Case "cmdTransactionsPrint"
5280                .chkTransactions.SetFocus
5290                .chkTransactions = True
5300                .GoToReport_arw_sp_trans_img.Top = ((.chkTransactions.Top + .chkTransactions.Height) - (4& * lngTpp))
5310                .GoToReport_arw_sp_trans_img.Left = (.chkTransactions.Left - (2& * lngTpp))
5320                .GoToReport_arw_sp_trans_img.Visible = True
5330                DoEvents
5340                chkTransactions_AfterUpdate  ' ** Procedure: Below.
5350                DoEvents
5360                Beep
5370                .TimerInterval = (GTR_WAIT * 2&)
5380              Case "cmdAssetListPrint"
5390                .chkAssetList.SetFocus
5400                .chkAssetList = True
5410                .GoToReport_arw_sp_alist_img.Top = ((.chkAssetList.Top + .chkAssetList.Height) - (4& * lngTpp))
5420                .GoToReport_arw_sp_alist_img.Left = (.chkAssetList.Left - (2& * lngTpp))
5430                .GoToReport_arw_sp_alist_img.Visible = True
5440                DoEvents
5450                chkAssetList_AfterUpdate  ' ** Procedure: Below.
5460                DoEvents
5470                Beep
5480                .TimerInterval = (GTR_WAIT * 2&)
5490              Case "cmdPrintStatement_Summary"
5500                .GoToReport_Emblem_01_img.Visible = True
5510                .chkStatements.SetFocus
5520                .chkStatements = True
5530                DoEvents
5540                chkStatements_AfterUpdate  ' ** Procedure: Below.
5550                DoEvents
5560                .GoToReport_arw_sp_stmt_img.Top = ((.chkStatements.Top + .chkStatements.Height) - (4& * lngTpp))
5570                .GoToReport_arw_sp_stmt_img.Left = (.chkStatements.Left - (2& * lngTpp))
5580                .GoToReport_arw_sp_stmt_img.Visible = True
5590                .GoToReport_lbl_sp_stmt_img.Visible = True
5600                .opgAccountNumber_lbl.Visible = False
5610                DoEvents
5620                Beep
5630                .TimerInterval = 50&
5640              End Select
5650            Case False
5660              Select Case garr_varGoToReport(GTR_CTL2)
                  Case "cmdPrintStatement_All"
5670                .GoToReport_Emblem_01_img.Visible = True
5680                .chkStatements.SetFocus
5690                .chkStatements = True
5700                DoEvents
5710                chkStatements_AfterUpdate  ' ** Procedure: Below.
5720                DoEvents
5730                .GoToReport_arw_sp_stmt_img.Top = ((.chkStatements.Top + .chkStatements.Height) - (4& * lngTpp))
5740                .GoToReport_arw_sp_stmt_img.Left = (.chkStatements.Left - (2& * lngTpp))
5750                .GoToReport_arw_sp_stmt_img.Visible = True
5760                .GoToReport_lbl_sp_stmt_img.Visible = True
5770                .opgAccountNumber_lbl.Visible = False
5780                DoEvents
5790                .TimerInterval = 50&
5800              Case "cmdBalanceTable"
5810                If .chkStatements = True Then
5820                  .chkTransactions.SetFocus
5830                  .chkTransactions = True
5840                  DoEvents
5850                  chkTransactions_AfterUpdate  ' ** Procedure: Below.
5860                End If
5870                .GoToReport_arw_sp_acct_d_img.Top = (.cmbAccounts.Top - (7& * lngTpp))
5880                .GoToReport_arw_sp_acct_d_img.Left = (.cmbAccounts.Left + .cmbAccounts.Width)
5890                .GoToReport_arw_sp_acct_t_img.Top = (.cmbAccounts.Top - (7& * lngTpp))
5900                .GoToReport_arw_sp_acct_t_img.Left = (.cmbAccounts.Left + .cmbAccounts.Width)
5910                .opgOrderBy.Enabled = False
5920                Select Case .opgOrderBy
                    Case .opgOrderBy_optDate.OptionValue
5930                  .GoToReport_arw_sp_acct_d_img.Visible = True
5940                Case .opgOrderBy_optType.OptionValue
5950                  .GoToReport_arw_sp_acct_t_img.Visible = True
5960                End Select
5970                .cmbAccounts.SetFocus
5980                .cmbAccounts = .cmbAccounts.Column(CBX_A_ACTNO, 0)
5990                DoEvents
6000                .TimerInterval = 50&
6010              Case Else
                    ' ** What else?
6020                blnGoingToReport = False
6030                DoCmd.Hourglass False
6040                gblnGoToReport = False
6050                blnGTR_Emblem = False
6060                GTREmblem_Off  ' ** Procedure: Below.
6070                Beep
6080              End Select
6090            End Select
6100          End Select
6110        End If
6120      End If
6130      If .PopUp = True And .DevMsg_lbl.Visible = True And gblnGoToReport = False Then
            ' ** Borrowing this variable from Journal/Transactions.
6140        Select Case gblnSwitchTo
            Case True
6150          .DevMsg_lbl.Visible = False
6160        Case False
6170          gblnSwitchTo = True
6180          .TimerInterval = 2000&
6190        End Select
6200      End If
6210      If blnResetStmts = True Then
6220        blnResetStmts = False
6230        .chkStatements = True  ' ** Should already be True.
6240        chkStatements_AfterUpdate  ' ** Procedure: Below.
6250        DoEvents
6260        If IsNull(.cmbMonth) = False Then
6270          cmbMonth_AfterUpdate  ' ** Procedure: Below.
6280        End If
6290      End If
6300      If gblnTimer = True Then
6310        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
6320          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
6330          gblnTimer = False
6340          DoCmd.Hourglass False
6350        End If
6360      End If
6370    End With

EXITP:
6380    Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

6410    Select Case DataErr
        Case 2001
          ' ** 2001: You canceled the previous operation.
6420      Response = acDataErrContinue
6430    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
6440      If Left(Me.ActiveControl.Name, 4) = "Date" Or Left(Me.ActiveControl.Name, 9) = "TransDate" Then  ' ** DateStart, DateEnd, TransDateStart, TransDateEnd.
6450        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, (Left(("Invalid Entry" & Space(40)), 40) & "FE01")
6460        Response = acDataErrContinue
6470      End If
6480      If Me.ActiveControl.Name = "StatementsYear" Then
6490        MsgBox "Please enter a valid year (YYYY).", vbInformation + vbOKOnly, (Left(("Invalid Entry" & Space(40)), 40) & "FE02")
6500        Response = acDataErrContinue
6510      End If
6520    Case 2115
          ' ** The macro or function set to the BeforeUpdate or ValidationRule property for
          ' ** this field is preventing Microsoft Access from saving the data in the field.
6530      Response = acDataErrContinue
6540    Case 2237  ' ** The text you entered isn't an item in the list.
6550      Response = acDataErrDisplay
6560    Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
6570      Beep
6580      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
6590      Response = acDataErrContinue
6600    Case 2501  ' ** 2501: The '|' action was Canceled.
6610      Response = acDataErrContinue
6620    Case Else
6630      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
6640      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
6650    End Select

EXITP:
6660    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** All the Preference queries cause the delays.

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngOrd As Long

6710    With Me

6720      DoCmd.Hourglass True
6730      DoEvents

          ' ** This is required in case user Closes Form with the
          ' ** Calendar still open. It also handles when the user
          ' ** closes the application with the Calendar still open.
6740      If Not clsMonthClass Is Nothing Then
6750        If clsMonthClass.IsCalendar Then
6760          Cancel = -1
6770        Else
6780          Set clsMonthClass = Nothing
6790        End If
6800      End If

6810      lngOrd = 0&
6820      If blnWasGTR = False Then
6830        Set dbs = CurrentDb
            ' ** Append qryForm_Close_01_03 (qryForm_Close_01_02 (qryForm_Close_01_01 (tblForm,
            ' ** by specified [frmnam]), linked to tblVBComponent_Procedure, by specified [procnam]),
            ' ** by specified  [parm1],  [parm2],  [parm3]), to tblForm_Close, by specified [ord], [actn].
6840        Set qdf = dbs.QueryDefs("qryForm_Close_01_04")
6850        If .chkRememberMe = False Then
6860          .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
6870            Pref_RemNot THIS_NAME, "Account"  ' ** Module Procedure: modPreferenceFuncs.
6880        End If
6890        If .chkRememberDates_Trans = False Then
6900          .TransDateStart = Null
6910          .TransDateEnd = Null
6920            Pref_RemNot THIS_NAME, "Dates", "Trans"  ' ** Module Procedure: modPreferenceFuncs.
6930        End If
6940        If .chkRememberDates_Asset = False Then
6950          .AssetListDate = Null
6960            Pref_RemNot THIS_NAME, "Dates", "Asset"  ' ** Module Procedure: modPreferenceFuncs.
6970        End If
6980        Set qdf = Nothing
6990        dbs.Close
7000        Set dbs = Nothing
7010        .cmbMonth = Null
7020        .StatementsYear = Null
7030        DoEvents
7040      End If

7050      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

7060      Set dbs = CurrentDb
            ' ** Empty tmpAccount.
7070        Set qdf = dbs.QueryDefs("qryStatementBalance_11_01")
7080        qdf.Execute
7090      Set qdf = Nothing
7100      dbs.Close
7110      Set dbs = Nothing
7120      DoEvents

7130      Do While Reports.Count > 0
7140        DoCmd.Close acReport, Reports(0).Name
7150      Loop
7160      DoEvents

7170      FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
7180      DoEvents

7190      ChkSpecLedgerEntry  ' ** Module Function: modUtilities.
7200      gvarCrtRpt_FL_SpecData = Null
7210      gblnLegalName = False
7220      gblnMessage = False
7230      gblnGoToReport = False
7240      gblnCrtRpt_Zero = False: gblnCrtRpt_ZeroDialog = False
7250      DoEvents

7260      gstrFormQuerySpec = vbNullString
7270      gstrReportQuerySpec = vbNullString
7280      gblnSetFocus = True
7290      gstrReturningForm = THIS_NAME
7300      Select Case strCallingForm
          Case "frmMenu_Report"
7310        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
7320      End Select

7330    End With

EXITP:
7340    Set qdf = Nothing
7350    Set dbs = Nothing
7360    Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

7410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
7420    Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

7510    DetailMouse_SP blnAnnualStatement_Focus, blnBalanceTable_Focus, blnCalendar1_Focus, blnCalendar2_Focus, blnCalendar3_Focus, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
7520    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdClose_Click()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

7610    DoCmd.Close acForm, THIS_NAME

EXITP:
7620    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7720    intShiftDown = (Shift And acShiftMask) > 0
7730    intAltDown = (Shift And acAltMask) > 0
7740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7760      Select Case intRetVal
          Case vbKeyTab
7770        With Me
7780          intRetVal = 0
7790          .chkTransactions.SetFocus
7800        End With
7810      End Select
7820    End If

        ' ** Shift keys.
7830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7840      Select Case intRetVal
          Case vbKeyTab
7850        With Me  ' ** Sorry, I'm just not sure what's enabled when.
7860          intRetVal = 0
7870          If .cmdPrintStatement_Summary.Enabled = True And .cmdPrintStatement_Summary.Visible = True Then
7880            .cmdPrintStatement_Summary.SetFocus
7890          ElseIf .cmdPrintStatement_Single.Enabled = True Then
7900            .cmdPrintStatement_Single.SetFocus
7910          ElseIf .cmdPrintStatement_All.Enabled = True Then
7920            .cmdPrintStatement_All.SetFocus
7930          ElseIf .cmdAssetListExcel.Enabled = True Then
7940            .cmdAssetListExcel.SetFocus  '#chk
7950          ElseIf .cmdAssetListWord.Enabled = True Then
7960            .cmdAssetListWord.SetFocus
7970          ElseIf .cmdTransactionsExcel.Enabled = True Then
7980            .cmdTransactionsExcel.SetFocus  '#chk
7990          ElseIf .cmdTransactionsWord.Enabled = True Then
8000            .cmdTransactionsWord.SetFocus
8010          ElseIf .cmbAccounts.Enabled = True Then
8020            .cmbAccounts.SetFocus
8030          Else
8040            .opgAccountNumber.SetFocus
8050          End If
8060        End With
8070      End Select
8080    End If

EXITP:
8090    KeyCode = intRetVal
8100    Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub chkTransactions_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTransactions_box_MouseMove"

8210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8220    Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkTransactions_GotFocus()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTransactions_GotFocus"

8310    TransAssetState_Handler_SP THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
8320    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Public Sub chkTransactions_AfterUpdate()
' ** Also called by:
' **   Form_Open(), above.

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTransactions_AfterUpdate"

8410    ChkTrans_After_SP blnIsOpen, blnRunPriorStatement, blnAcctNotSched, datAssetListDate_Pref, _
          blnHasForEx, blnHasForExThis, lngAcctFors, arr_varAcctFor, lngAcctArchs, arr_varAcctArch, _
          lngStmts, arr_varStmt, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
8420    Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl: That_Desc = ERR.description
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub chkTransactions_KeyDown(KeyCode As Integer, Shift As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTransactions_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8520    intShiftDown = (Shift And acShiftMask) > 0
8530    intAltDown = (Shift And acAltMask) > 0
8540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8560      Select Case intRetVal
          Case vbKeyTab  ' ** Let vbKeyReturn do what it normally does.
8570        With Me
8580          intRetVal = 0
8590          Select Case .chkTransactions
              Case True
8600            .TransDateStart.SetFocus
8610          Case False
8620            .chkAssetList.SetFocus
8630          End Select
8640        End With
8650      End Select
8660    End If

        ' ** Shift keys.
8670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8680      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8690        With Me
8700          intRetVal = 0
8710          .cmdClose.SetFocus
8720        End With
8730      End Select
8740    End If

EXITP:
8750    KeyCode = intRetVal
8760    Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub TransDateStart_KeyDown(KeyCode As Integer, Shift As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8820    intShiftDown = (Shift And acShiftMask) > 0
8830    intAltDown = (Shift And acAltMask) > 0
8840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8870        With Me
8880          intRetVal = 0
8890          .TransDateEnd.SetFocus
8900        End With
8910      End Select
8920    End If

        ' ** Shift keys.
8930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8950        With Me
8960          intRetVal = 0
8970          .chkTransactions.SetFocus
8980        End With
8990      End Select
9000    End If

        ' ** Ctrl keys.
9010    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
9020      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9030        With Me
9040          intRetVal = 0
9050          .cmdCalendar1.SetFocus
9060        End With
9070      End Select
9080    End If

EXITP:
9090    KeyCode = intRetVal
9100    Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub TransDateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStart_MouseMove"

9210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9220    Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub TransDateStart_Exit(Cancel As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStart_Exit"

9310    TransDateStmtYearExit_SP 1, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
9320    Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

9410    Calendar_Handler_SP THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
9420    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

9510    Calendar_Handler_SP THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
9520    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

9610    Calendar_Handler_SP THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
9620    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1480    Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

9710    Calendar_KeyDown_SP KeyCode, Shift, THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
9720    Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

9810    Calendar_Handler_SP THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
9820    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

9910    Calendar_Handler_SP THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
9920    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

10010   Calendar_Handler_SP THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
10020   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub TransDateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10120   intShiftDown = (Shift And acShiftMask) > 0
10130   intAltDown = (Shift And acAltMask) > 0
10140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10170       With Me
10180         intRetVal = 0
10190         If .chkTransactions = True And IsNull(.TransDateStart) = False Then
10200           If (IsNull(.TransDateEnd) = False Or IsNull(.TransDateEnd.text) = False) Then
10210 On Error Resume Next
10220             .opgAccountNumber.SetFocus
10230             If ERR.Number <> 0 Then
10240               If ERR.Number = 2110 Then
10250 On Error GoTo ERRH
10260                 MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, _
                        (Left(("Invalid Date" & Space(55)), 55) & "B01")
10270               Else
10280                 zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10290 On Error GoTo ERRH
10300               End If
10310             Else
10320 On Error GoTo ERRH
10330             End If
10340           End If
10350         Else
10360           .chkAssetList.SetFocus
10370         End If
10380       End With
10390     End Select
10400   End If

        ' ** Shift keys.
10410   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10420     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10430       With Me
10440         intRetVal = 0
10450         .TransDateStart.SetFocus
10460       End With
10470     End Select
10480   End If

        ' ** Ctrl keys.
10490   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10500     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10510       With Me
10520         intRetVal = 0
10530         .cmdCalendar2.SetFocus
10540       End With
10550     End Select
10560   End If

        ' ** Ctrl-Shift keys.
10570   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10580     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10590       With Me
10600         intRetVal = 0
10610         .cmdCalendar1.SetFocus
10620       End With
10630     End Select
10640   End If

EXITP:
10650   KeyCode = intRetVal
10660   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1480    Case Else
1490      intRetVal = 0
1500      THAT_PROC = THIS_PROC
1510      That_Erl = Erl
1520      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub TransDateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateEnd_MouseMove"

10710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10720   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub TransDateEnd_Exit(Cancel As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateEnd_Exit"

10810   TransDateStmtYearExit_SP 2, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
10820   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

10910   Calendar_Handler_SP THIS_PROC, blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
10920   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

11010   Calendar_Handler_SP THIS_PROC, blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
11020   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

11110   Calendar_Handler_SP THIS_PROC, blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
11120   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1480    Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

11210   Calendar_KeyDown_SP KeyCode, Shift, THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
11220   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

11310   Calendar_Handler_SP THIS_PROC, blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
11320   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

11410   Calendar_Handler_SP THIS_PROC, blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
11420   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

11510   Calendar_Handler_SP THIS_PROC, blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
11520   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub chkRememberDates_Trans_AfterUpdate()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_Trans_AfterUpdate"

11610   Remember_Handler_SP THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
11620   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkRememberDates_Trans_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_Trans_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11720   intShiftDown = (Shift And acShiftMask) > 0
11730   intAltDown = (Shift And acAltMask) > 0
11740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11770       With Me
11780         intRetVal = 0
11790         If .chkIncludeArchive_Trans.Enabled = True Then
11800           .chkIncludeArchive_Trans.SetFocus
11810         ElseIf .chkArchiveOnly_Trans.Enabled = True Then
11820           .chkArchiveOnly_Trans.SetFocus
11830         Else
11840           .chkAssetList.SetFocus
11850         End If
11860       End With
11870     End Select
11880   End If

        ' ** Shift keys.
11890   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11900     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11910       With Me
11920         intRetVal = 0
11930         .cmdCalendar2.SetFocus
11940       End With
11950     End Select
11960   End If

EXITP:
11970   KeyCode = intRetVal
11980   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub chkRememberDates_Trans_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_Trans_MouseMove"

12010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12020   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Public Sub chkIncludeArchive_Trans_AfterUpdate()

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_Trans_AfterUpdate"

12110   Archive_Handler_SP THIS_PROC, blnLoop, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
12120   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkIncludeArchive_Trans_KeyDown(KeyCode As Integer, Shift As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_Trans_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12220   intShiftDown = (Shift And acShiftMask) > 0
12230   intAltDown = (Shift And acAltMask) > 0
12240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12270       With Me
12280         intRetVal = 0
12290         .chkArchiveOnly_Trans.SetFocus
12300       End With
12310     End Select
12320   End If

        ' ** Shift keys.
12330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12350       With Me
12360         intRetVal = 0
12370         If .chkRememberDates_Trans.Enabled = True Then
12380           .chkRememberDates_Trans.SetFocus
12390         Else
12400           .chkTransactions.SetFocus
12410         End If
12420       End With
12430     End Select
12440   End If

EXITP:
12450   KeyCode = intRetVal
12460   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Public Sub chkArchiveOnly_Trans_AfterUpdate()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkArchiveOnly_Trans_AfterUpdate"

12510   Archive_Handler_SP THIS_PROC, blnLoop, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
12520   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkArchiveOnly_Trans_KeyDown(KeyCode As Integer, Shift As Integer)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkArchiveOnly_Trans_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12620   intShiftDown = (Shift And acShiftMask) > 0
12630   intAltDown = (Shift And acAltMask) > 0
12640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12670       With Me
12680         intRetVal = 0
12690         .chkAssetList.SetFocus
12700       End With
12710     End Select
12720   End If

        ' ** Shift keys.
12730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12750       With Me
12760         intRetVal = 0
12770         .chkIncludeArchive_Trans.SetFocus
12780       End With
12790     End Select
12800   End If

EXITP:
12810   KeyCode = intRetVal
12820   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub chkAssetList_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_box_MouseMove"

12910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12920   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkAssetList_GotFocus()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_GotFocus"

13010   TransAssetState_Handler_SP THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
13020   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Public Sub chkAssetList_AfterUpdate()
' ** Also called by:
' **   chkTransactions_AfterUpdate(), above.
' **   chkStatements_AfterUpdate(), below.

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_AfterUpdate"

13110   ChkAstList_After_SP blnIsOpen, blnRunPriorStatement, blnAcctNotSched, datAssetListDate_Pref, _
          blnHasForEx, blnHasForExThis, lngAcctFors, arr_varAcctFor, lngAcctArchs, arr_varAcctArch, _
          lngStmts, arr_varStmt, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
13120   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl: That_Desc = ERR.description
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub chkAssetList_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13220   intShiftDown = (Shift And acShiftMask) > 0
13230   intAltDown = (Shift And acAltMask) > 0
13240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13270       With Me
13280         intRetVal = 0
13290         Select Case .chkAssetList
              Case True
13300           .AssetListDate.SetFocus
13310         Case False
13320           .chkStatements.SetFocus
13330         End Select
13340       End With
13350     End Select
13360   End If

        ' ** Shift keys.
13370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13390       With Me
13400         intRetVal = 0
13410         If .TransDateEnd.Enabled = True Then
13420           .TransDateEnd.SetFocus
13430         Else
13440           .chkTransactions.SetFocus
13450         End If
13460       End With
13470     End Select
13480   End If

        ' ** Ctrl-Shift keys.
13490   If intCtrlDown And (Not intAltDown) And intShiftDown Then
13500     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13510       With Me
13520         intRetVal = 0
13530         If .chkArchiveOnly_Trans.Enabled = True Then
13540           .chkArchiveOnly_Trans.SetFocus
13550         ElseIf .chkIncludeArchive_Trans.Enabled = True Then
13560           .chkIncludeArchive_Trans.SetFocus
13570         ElseIf .chkRememberDates_Trans.Enabled = True Then
13580           .chkRememberDates_Trans.SetFocus
13590         Else
                ' ** Nothing.
13600         End If
13610       End With
13620     End Select
13630   End If

EXITP:
13640   KeyCode = intRetVal
13650   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub AssetListDate_GotFocus()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetListDate_GotFocus"

13710   With Me
13720     If blnGoingToReport = True Then
13730       .AssetListDate.SelLength = 0
13740       .AssetListDate.SelStart = 19
13750     End If
13760   End With

EXITP:
13770   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub AssetListDate_AfterUpdate()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetListDate_AfterUpdate"

13810   With Me
13820     If IsNull(.AssetListDate) = False Then
13830       If IsDate(.AssetListDate) = True Then
13840         datAssetListDate_Pref = .AssetListDate
13850       End If
13860     End If
13870   End With

EXITP:
13880   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub AssetListDate_KeyDown(KeyCode As Integer, Shift As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetListDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13920   intShiftDown = (Shift And acShiftMask) > 0
13930   intAltDown = (Shift And acAltMask) > 0
13940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13970       With Me
13980         intRetVal = 0
13990         Select Case .chkAssetList
              Case True
14000           If IsNull(.AssetListDate) = False Or IsNull(.AssetListDate.text) = False Then
14010             .opgAccountNumber.SetFocus
14020           End If
14030         Case False
14040           .chkStatements.SetFocus
14050         End Select
14060       End With
14070     End Select
14080   End If

        ' ** Shift keys.
14090   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14110       With Me
14120         intRetVal = 0
14130         .chkAssetList.SetFocus
14140       End With
14150     End Select
14160   End If

        ' ** Ctrl keys.
14170   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
14180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14190       With Me
14200         intRetVal = 0
14210         .cmdCalendar3.SetFocus
14220       End With
14230     End Select
14240   End If

EXITP:
14250   KeyCode = intRetVal
14260   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub AssetListDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetListDate_MouseMove"

14310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14320   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub AssetListDate_LostFocus()

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetListDate_LostFocus"

14410   AssetListDateLost_SP Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
14420   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdCalendar3_GotFocus()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_GotFocus"

14510   Calendar_Handler_SP THIS_PROC, blnCalendar3_Focus, blnCalendar3_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
14520   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_MouseDown"

14610   Calendar_Handler_SP THIS_PROC, blnCalendar3_Focus, blnCalendar3_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
14620   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar3_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_Click"

14710   Calendar_Handler_SP THIS_PROC, blnCalendar3_Focus, blnCalendar3_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
14720   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1480    Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub cmdCalendar3_KeyDown(KeyCode As Integer, Shift As Integer)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_KeyDown"

14810   Calendar_KeyDown_SP KeyCode, Shift, THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
14820   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdCalendar3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_MouseMove"

14910   Calendar_Handler_SP THIS_PROC, blnCalendar3_Focus, blnCalendar3_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
14920   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_MouseUp"

15010   Calendar_Handler_SP THIS_PROC, blnCalendar3_Focus, blnCalendar3_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
15020   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdCalendar3_LostFocus()

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_LostFocus"

15110   Calendar_Handler_SP THIS_PROC, blnCalendar3_Focus, blnCalendar3_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
15120   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub chkRememberDates_Asset_AfterUpdate()

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_Asset_AfterUpdate"

15210   Remember_Handler_SP THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
15220   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkRememberDates_Asset_KeyDown(KeyCode As Integer, Shift As Integer)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_Asset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15320   intShiftDown = (Shift And acShiftMask) > 0
15330   intAltDown = (Shift And acAltMask) > 0
15340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15370       With Me
15380         intRetVal = 0
15390         If .chkIncludeArchive_Asset.Enabled = True Then
15400           .chkIncludeArchive_Asset.SetFocus
15410         Else
15420           .chkStatements.SetFocus
15430         End If
15440       End With
15450     End Select
15460   End If

        ' ** Shift keys.
15470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15490       With Me
15500         intRetVal = 0
15510         .cmdCalendar3.SetFocus
15520       End With
15530     End Select
15540   End If

EXITP:
15550   KeyCode = intRetVal
15560   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub chkRememberDates_Asset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_Asset_MouseMove"

15610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15620   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkIncludeArchive_Asset_AfterUpdate()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_Asset_AfterUpdate"

15710   Archive_Handler_SP THIS_PROC, blnLoop, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
15720   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkStatements_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkStatements_box_MouseMove"

15810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15820   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkStatements_GotFocus()

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkStatements_GotFocus"

15910   TransAssetState_Handler_SP THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
15920   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Public Sub chkStatements_AfterUpdate()
' ** Also called by:
' **   chkAssetList_AfterUpdate(), above.

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkStatements_AfterUpdate"

16010   ChkStmt_After_SP blnIsOpen, blnRunPriorStatement, blnAcctNotSched, datAssetListDate_Pref, _
          lngAcctArchs, arr_varAcctArch, lngAcctFors, arr_varAcctFor, blnHasForEx, blnHasForExThis, _
          lngStmts, arr_varStmt, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
16020   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl: That_Desc = ERR.description
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub chkStatements_KeyDown(KeyCode As Integer, Shift As Integer)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkStatements_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16120   intShiftDown = (Shift And acShiftMask) > 0
16130   intAltDown = (Shift And acAltMask) > 0
16140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16170       With Me
16180         intRetVal = 0
16190         Select Case .chkStatements
              Case True
16200           .cmbMonth.SetFocus
16210         Case False
16220           If .cmdAnnualStatement.Enabled = True Then
16230             .cmdAnnualStatement.SetFocus
16240           ElseIf .cmdBalanceTable.Enabled = True Then
16250             .cmdBalanceTable.SetFocus
16260           Else
16270             .opgAccountNumber.SetFocus
16280           End If
16290         End Select
16300       End With
16310     End Select
16320   End If

        ' ** Shift keys.
16330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16350       With Me
16360         intRetVal = 0
16370         If .AssetListDate.Enabled = True Then
16380           .AssetListDate.SetFocus
16390         Else
16400           .chkAssetList.SetFocus
16410         End If
16420       End With
16430     End Select
16440   End If

        ' ** Ctrl-Shift keys.
16450   If intCtrlDown And (Not intAltDown) And intShiftDown Then
16460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16470       With Me
16480         intRetVal = 0
16490         If .chkIncludeArchive_Asset.Enabled = True Then
16500           .chkIncludeArchive_Asset.SetFocus
16510         ElseIf .chkRememberDates_Asset.Enabled = True Then
16520           .chkRememberDates_Asset.SetFocus
16530         Else
                ' ** Nothing.
16540         End If
16550       End With
16560     End Select
16570   End If

EXITP:
16580   KeyCode = intRetVal
16590   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub cmbMonth_Enter()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbMonth_Enter"

16610   With Me
16620     If .cmbMonth.text = vbNullString And blnGoingToReport = False Then
16630       ExpandCombo .cmbMonth  ' ** Module Procedure: modUtilities.
16640     End If
16650   End With

EXITP:
16660   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Public Sub cmbMonth_AfterUpdate()
' ** Called by:
' **   Form_Timer(), Above.
' **   modStatementParamFuncs2:
' **     CmbAccts_After_SP()

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbMonth_AfterUpdate"

16710   Month_AfterUpdate_SP Me  ' ** Module Procedure: modStatementParamFuncs2.
        'Calls CoOptions_Read()
        '  Runs Qrys, initializes variables.
        'DLookup() to "Statement Date"
        'Calls ForEx_ChkScheduled():
        '  Calls ForExArr_Load()
        '    Runs Qyrs, loads arrays.
        '  Calls AcctSched_Load()
        '    Runs Qyrs, loads arrays.
        '  Calls chkIncludeCurrency_AfterUpdate()
        '    Sets Bold.

EXITP:
16720   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub cmbMonth_KeyDown(KeyCode As Integer, Shift As Integer)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbMonth_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16820   intShiftDown = (Shift And acShiftMask) > 0
16830   intAltDown = (Shift And acAltMask) > 0
16840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16870       With Me
16880         intRetVal = 0
16890         .StatementsYear.SetFocus
16900       End With
16910     End Select
16920   End If

        ' ** Shift keys.
16930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16950       With Me
16960         intRetVal = 0
16970         .chkStatements.SetFocus
16980       End With
16990     End Select
17000   End If

EXITP:
17010   KeyCode = intRetVal
17020   Exit Sub

ERRH:
1470    intRetVal = 0
1480    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1490    Case Else
1500      THAT_PROC = THIS_PROC
1510      That_Erl = Erl
1520      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub StatementsYear_KeyDown(KeyCode As Integer, Shift As Integer)

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "StatementsYear_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17120   intShiftDown = (Shift And acShiftMask) > 0
17130   intAltDown = (Shift And acAltMask) > 0
17140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17170       With Me
17180         If .chkStatements = True Then
17190           If IsNull(.cmbMonth) = False And (IsNull(.StatementsYear) = False Or IsNull(.StatementsYear.text) = False) Then
17200             intRetVal = 0
17210             .opgAccountNumber.SetFocus
17220           End If
17230         End If
17240       End With
17250     End Select
17260   End If

        ' ** Shift keys.
17270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17290       With Me
17300         .cmbMonth.SetFocus
17310       End With
17320     End Select
17330   End If

        ' ** Ctrl keys.
17340   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
17350     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17360       With Me
17370         If .cmdAnnualStatement.Enabled = True Then
17380           .cmdAnnualStatement.SetFocus
17390         ElseIf .cmdBalanceTable.Enabled = True Then
17400           .cmdBalanceTable.SetFocus
17410         Else
                ' ** Nothing
17420         End If
17430       End With
17440     End Select
17450   End If

EXITP:
17460   KeyCode = intRetVal
17470   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub StatementsYear_Exit(Cancel As Integer)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "StatementsYear_Exit"

17510   TransDateStmtYearExit_SP 3, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
17520   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkStatements_box2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkStatements_box2_MouseMove"

17610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17620   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkStatements_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkStatements_lbl2_MouseMove"

17710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17720   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkStatements_lbl3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkStatements_lbl3_MouseMove"

17810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17820   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdAnnualStatement_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAnnualStatement_box_MouseMove"

17910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17920   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdAnnualStatement_GotFocus()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAnnualStatement_GotFocus"

18010   AnnualBalance_Handler_SP THIS_PROC, blnAnnualStatement_Focus, blnAnnualStatement_MouseDown, _
          blnBalanceTable_Focus, blnBalanceTable_MouseDown, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
18020   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdAnnualStatement_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAnnualStatement_MouseDown"

18110   AnnualBalance_Handler_SP THIS_PROC, blnAnnualStatement_Focus, blnAnnualStatement_MouseDown, _
          blnBalanceTable_Focus, blnBalanceTable_MouseDown, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
18120   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdAnnualStatement_Click()
' ** December or 4th Quarter statement must have been run.

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAnnualStatement_Click"

18210   AnnualStatement_Print Me, blnContinue, blnFromStmts, blnPrintAnnualStatement, blnPrintStatements, _
          blnAllStatements, blnSingleStatement, blnRunPriorStatement, blnAcctsSchedRpt, datFirstDate, _
          blnGoingToReport, blnGoingToReport2, blnGTR_Emblem, blnWasGTR  ' ** Module Procedure: modStatementParamFuncs1.

18220   DoCmd.Hourglass False

EXITP:
18230   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Me.PrintAnnual_chk = False
1490    Me.cmbAccounts.Locked = False
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub cmdAnnualStatement_KeyDown(KeyCode As Integer, Shift As Integer)

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAnnualStatement_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18320   intShiftDown = (Shift And acShiftMask) > 0
18330   intAltDown = (Shift And acAltMask) > 0
18340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18370       With Me
18380         If .cmdBalanceTable.Enabled = True Then
18390           .cmdBalanceTable.SetFocus
18400         Else
18410           .opgAccountNumber.SetFocus
18420         End If
18430       End With
18440     End Select
18450   End If

        ' ** Shift keys.
18460   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18470     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18480       With Me
18490         If .StatementsYear.Enabled = True Then
18500           .StatementsYear.SetFocus
18510         Else
18520           .chkStatements.SetFocus
18530         End If
18540       End With
18550     End Select
18560   End If

EXITP:
18570   KeyCode = intRetVal
18580   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub cmdAnnualStatement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAnnualStatement_MouseMove"

18610   AnnualBalance_Handler_SP THIS_PROC, blnAnnualStatement_Focus, blnAnnualStatement_MouseDown, _
          blnBalanceTable_Focus, blnBalanceTable_MouseDown, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
18620   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdAnnualStatement_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAnnualStatement_MouseUp"

18710   AnnualBalance_Handler_SP THIS_PROC, blnAnnualStatement_Focus, blnAnnualStatement_MouseDown, _
          blnBalanceTable_Focus, blnBalanceTable_MouseDown, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
18720   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdAnnualStatement_LostFocus()

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAnnualStatement_LostFocus"

18810   AnnualBalance_Handler_SP THIS_PROC, blnAnnualStatement_Focus, blnAnnualStatement_MouseDown, _
          blnBalanceTable_Focus, blnBalanceTable_MouseDown, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
18820   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdBalanceTable_GotFocus()

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBalanceTable_GotFocus"

18910   AnnualBalance_Handler_SP THIS_PROC, blnAnnualStatement_Focus, blnAnnualStatement_MouseDown, _
          blnBalanceTable_Focus, blnBalanceTable_MouseDown, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
18920   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdBalanceTable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBalanceTable_MouseDown"

19010   AnnualBalance_Handler_SP THIS_PROC, blnAnnualStatement_Focus, blnAnnualStatement_MouseDown, _
          blnBalanceTable_Focus, blnBalanceTable_MouseDown, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
19020   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdBalanceTable_Click()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBalanceTable_Click"

        Dim blnContinue2 As Boolean

19110   blnContinue2 = True
19120   CmdBalTbl_Click_SP blnContinue2, blnGoingToReport, blnGTR_Emblem, _
          THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
19130   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
1480    Case Else
1490      THAT_PROC = THIS_PROC
1500      That_Erl = Erl: That_Desc = ERR.description
1510      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdBalanceTable_KeyDown(KeyCode As Integer, Shift As Integer)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBalanceTable_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19220   intShiftDown = (Shift And acShiftMask) > 0
19230   intAltDown = (Shift And acAltMask) > 0
19240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19270       With Me
19280         .opgAccountNumber.SetFocus
19290       End With
19300     End Select
19310   End If

        ' ** Shift keys.
19320   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19330     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19340       With Me
19350         If .cmdAnnualStatement.Enabled = True Then
19360           .cmdAnnualStatement.SetFocus
19370         ElseIf .StatementsYear.Enabled = True Then
19380           .StatementsYear.SetFocus
19390         Else
19400           .chkStatements.SetFocus
19410         End If
19420       End With
19430     End Select
19440   End If

EXITP:
19450   KeyCode = intRetVal
19460   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub cmdBalanceTable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBalanceTable_MouseMove"

19510   AnnualBalance_Handler_SP THIS_PROC, blnAnnualStatement_Focus, blnAnnualStatement_MouseDown, _
          blnBalanceTable_Focus, blnBalanceTable_MouseDown, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
19520   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdBalanceTable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBalanceTable_MouseUp"

19610   AnnualBalance_Handler_SP THIS_PROC, blnAnnualStatement_Focus, blnAnnualStatement_MouseDown, _
          blnBalanceTable_Focus, blnBalanceTable_MouseDown, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
19620   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdBalanceTable_LostFocus()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBalanceTable_LostFocus"

19710   AnnualBalance_Handler_SP THIS_PROC, blnAnnualStatement_Focus, blnAnnualStatement_MouseDown, _
          blnBalanceTable_Focus, blnBalanceTable_MouseDown, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
19720   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Public Sub opgAccountNumber_AfterUpdate()
' ** Also called by:
' **   chkArchiveOnly_AfterUpdate()
' **   chkAssetList_AfterUpdate()
' **   chkStatements_AfterUpdate()
' **   chkTransactions_AfterUpdate()
' **   cmdAnnualStatement_Click()
' **   cmdAssetListPreview_Click()
' **   cmdAssetListPrint_Click()
' **   cmdAssetListExcel_Click()
' **   cmdAssetListWord_Click()
' **   cmdPrintStatement_All_Click()

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_AfterUpdate"

19810   OpgActNo_After_SP blnResetStmts, blnRunPriorStatement, strRememberMe, _
          blnHasForEx, blnHasForExThis, lngAcctArchs, arr_varAcctArch, lngAcctFors, arr_varAcctFor, _
          THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
19820   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl: That_Desc = ERR.description
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub opgAccountNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19920   intShiftDown = (Shift And acShiftMask) > 0
19930   intAltDown = (Shift And acAltMask) > 0
19940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19970       With Me
19980         If .cmbAccounts.Enabled = True Then
19990           .cmbAccounts.SetFocus
20000         Else
20010           If .chkTransactions = True And .cmdTransactionsPreview.Enabled = True Then
20020             .cmdTransactionsPreview.SetFocus
20030           ElseIf .chkAssetList = True And .cmdAssetListPreview.Enabled = True Then
20040             .cmdAssetListPreview.SetFocus
20050           ElseIf .chkStatements = True And .cmdTransactionsPreview.Enabled = True Then
20060             .cmdTransactionsPreview.SetFocus
20070           Else
20080             .cmdClose.SetFocus
20090           End If
20100         End If
20110       End With
20120     End Select
20130   End If

        ' ** Shift keys.
20140   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20150     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20160       With Me
20170         If .chkTransactions = True And .TransDateEnd.Enabled = True Then
20180           .TransDateEnd.SetFocus
20190         ElseIf .chkAssetList = True And .AssetListDate.Enabled = True Then
20200           .AssetListDate.SetFocus
20210         ElseIf .chkStatements = True And .StatementsYear.Enabled = True Then
20220           .StatementsYear.SetFocus
20230         Else
20240           .cmdClose.SetFocus
20250         End If
20260       End With
20270     End Select
20280   End If

        ' ** Ctrl keys.
20290   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
20300     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20310       With Me
20320         If .chkRelatedAccounts.Enabled = True Then
20330           .chkRelatedAccounts.SetFocus
20340         Else
20350           .chkCombineCash.SetFocus
20360         End If
20370       End With
20380     End Select
20390   End If

        ' ** Ctrl-Shift keys.
20400   If intCtrlDown And (Not intAltDown) And intShiftDown Then
20410     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20420       With Me
20430         .cmdBalanceTable.SetFocus
20440       End With
20450     End Select
20460   End If

EXITP:
20470   KeyCode = intRetVal
20480   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub opgAccountNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20520   intShiftDown = (Shift And acShiftMask) > 0
20530   intAltDown = (Shift And acAltMask) > 0
20540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20570       With Me
20580         If .cmbAccounts.Enabled = True Then
20590           .cmbAccounts.SetFocus
20600         Else
20610           If .chkTransactions = True And .cmdTransactionsPreview.Enabled = True Then
20620             .cmdTransactionsPreview.SetFocus
20630           ElseIf .chkAssetList = True And .cmdAssetListPreview.Enabled = True Then
20640             .cmdAssetListPreview.SetFocus
20650           ElseIf .chkStatements = True And .cmdTransactionsPreview.Enabled = True Then
20660             .cmdTransactionsPreview.SetFocus
20670           Else
20680             .cmdClose.SetFocus
20690           End If
20700         End If
20710       End With
20720     End Select
20730   End If

        ' ** Shift keys.
20740   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20750     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20760       With Me
20770         If .chkTransactions = True And .TransDateEnd.Enabled = True Then
20780           .TransDateEnd.SetFocus
20790         ElseIf .chkAssetList = True And .AssetListDate.Enabled = True Then
20800           .AssetListDate.SetFocus
20810         ElseIf .chkStatements = True And .StatementsYear.Enabled = True Then
20820           .StatementsYear.SetFocus
20830         Else
20840           .cmdClose.SetFocus
20850         End If
20860       End With
20870     End Select
20880   End If

        ' ** Ctrl keys.
20890   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
20900     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20910       With Me
20920         If .chkRelatedAccounts.Enabled = True Then
20930           .chkRelatedAccounts.SetFocus
20940         Else
20950           .chkCombineCash.SetFocus
20960         End If
20970       End With
20980     End Select
20990   End If

EXITP:
21000   KeyCode = intRetVal
21010   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Public Sub chkRelatedAccounts_AfterUpdate()

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRelatedAccounts_AfterUpdate"

21110   With Me
21120     Select Case .chkRelatedAccounts
          Case True
21130       .chkRelatedAccounts_lbl.FontBold = True
21140     Case False
21150       .chkRelatedAccounts_lbl.FontBold = False
21160     End Select
21170     .chkLegalName.Enabled = Not (.chkRelatedAccounts)
21180     If .chkLegalName.Enabled = False Then
21190       .chkLegalName = False
21200       .chkLegalName_lbl.FontBold = False
21210     End If
21220     gblnLegalName = .chkLegalName
21230   End With

EXITP:
21240   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkRelatedAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRelatedAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21320   intShiftDown = (Shift And acShiftMask) > 0
21330   intAltDown = (Shift And acAltMask) > 0
21340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21370       With Me
21380         intRetVal = 0
21390         If .chkCombineCash.Enabled = True Then
21400           .chkCombineCash.SetFocus
21410         ElseIf .chkLegalName.Enabled = True Then
21420           .chkLegalName.SetFocus
21430         ElseIf .chkForeignExchange.Enabled = True Then
21440           .chkForeignExchange.SetFocus
21450         ElseIf .cmbAccounts.Enabled = True Then
21460           .cmbAccounts.SetFocus
21470         ElseIf .cmdTransactionsPreview.Enabled = True Then
21480           .cmdTransactionsPreview.SetFocus
21490         ElseIf .cmdAssetListPreview.Enabled = True Then
21500           .cmdAssetListPreview.SetFocus
21510         Else
21520           .cmdClose.SetFocus
21530         End If
21540       End With
21550     End Select
21560   End If

        ' ** Shift keys.
21570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21580     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21590       With Me
21600         intRetVal = 0
21610         .opgAccountNumber.SetFocus
21620       End With
21630     End Select
21640   End If

EXITP:
21650   KeyCode = intRetVal
21660   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Public Sub chkCombineCash_AfterUpdate()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkCombineCash_AfterUpdate"

21710   With Me
21720     Select Case .chkCombineCash
          Case True
21730       .chkCombineCash_lbl.FontBold = True
21740     Case False
21750       .chkCombineCash_lbl.FontBold = False
21760     End Select
21770   End With

EXITP:
21780   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkCombineCash_KeyDown(KeyCode As Integer, Shift As Integer)

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkCombineCash_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21820   intShiftDown = (Shift And acShiftMask) > 0
21830   intAltDown = (Shift And acAltMask) > 0
21840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21870       With Me
21880         intRetVal = 0
21890         If .chkLegalName.Enabled = True Then
21900           .chkLegalName.SetFocus
21910         ElseIf .chkForeignExchange.Enabled = True Then
21920           .chkForeignExchange.SetFocus
21930         ElseIf .cmbAccounts.Enabled = True Then
21940           .cmbAccounts.SetFocus
21950         ElseIf .cmdTransactionsPreview.Enabled = True Then
21960           .cmdTransactionsPreview.SetFocus
21970         ElseIf .cmdAssetListPreview.Enabled = True Then
21980           .cmdAssetListPreview.SetFocus
21990         Else
22000           .cmdClose.SetFocus
22010         End If
22020       End With
22030     End Select
22040   End If

        ' ** Shift keys.
22050   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22070       With Me
22080         intRetVal = 0
22090         If .chkRelatedAccounts.Enabled = True Then
22100           .chkRelatedAccounts.SetFocus
22110         Else
22120           .opgAccountNumber.SetFocus
22130         End If
22140       End With
22150     End Select
22160   End If

EXITP:
22170   KeyCode = intRetVal
22180   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub chkLegalName_AfterUpdate()

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_AfterUpdate"

22210   With Me
22220     Select Case .chkLegalName
          Case True
22230       .chkLegalName_lbl.FontBold = True
22240       gblnLegalName = True
22250     Case False
22260       .chkLegalName_lbl.FontBold = False
22270       gblnLegalName = False
22280     End Select
22290   End With

EXITP:
22300   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkLegalName_KeyDown(KeyCode As Integer, Shift As Integer)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22420   intShiftDown = (Shift And acShiftMask) > 0
22430   intAltDown = (Shift And acAltMask) > 0
22440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22470       With Me
22480         intRetVal = 0
22490         If .chkForeignExchange.Enabled = True Then
22500           .chkForeignExchange.SetFocus
22510         ElseIf .cmbAccounts.Enabled = True Then
22520           .opgAccountSource.SetFocus
22530         ElseIf .cmdTransactionsPreview.Enabled = True Then
22540           .cmdTransactionsPreview.SetFocus
22550         ElseIf .cmdAssetListPreview.Enabled = True Then
22560           .cmdAssetListPreview.SetFocus
22570         Else
22580           .cmdClose.SetFocus
22590         End If
22600       End With
22610     End Select
22620   End If

        ' ** Shift keys.
22630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22650       With Me
22660         intRetVal = 0
22670         If .chkCombineCash.Enabled = True Then
22680           .chkCombineCash.SetFocus
22690         ElseIf .chkRelatedAccounts.Enabled = True Then
22700           .chkRelatedAccounts.SetFocus
22710         Else
22720           .opgAccountNumber.SetFocus
22730         End If
22740       End With
22750     End Select
22760   End If

EXITP:
22770   KeyCode = intRetVal
22780   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub chkForeignExchange_AfterUpdate()

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkForeignExchange_AfterUpdate"

22810   With Me
22820     Select Case .chkForeignExchange
          Case True
22830       .chkForeignExchange_lbl.FontBold = True
22840       gblnForeignCurrencies = True
22850     Case False
22860       .chkForeignExchange_lbl.FontBold = False
22870       gblnForeignCurrencies = False
22880     End Select
22890   End With

EXITP:
22900   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkForeignExchange_KeyDown(KeyCode As Integer, Shift As Integer)

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkForeignExchange_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23020   intShiftDown = (Shift And acShiftMask) > 0
23030   intAltDown = (Shift And acAltMask) > 0
23040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23070       With Me
23080         intRetVal = 0
23090         If .chkIncludeCurrency.Enabled = True Then
23100           .chkIncludeCurrency.SetFocus
23110         ElseIf .cmbAccounts.Enabled = True Then
23120           .cmbAccounts.SetFocus
23130         ElseIf .cmdTransactionsPreview.Enabled = True Then
23140           .cmdTransactionsPreview.SetFocus
23150         ElseIf .cmdAssetListPreview.Enabled = True Then
23160           .cmdAssetListPreview.SetFocus
23170         Else
23180           .cmdClose.SetFocus
23190         End If
23200       End With
23210     End Select
23220   End If

        ' ** Shift keys.
23230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23250       With Me
23260         intRetVal = 0
23270         If .chkLegalName.Enabled = True Then
23280           .chkLegalName.SetFocus
23290         ElseIf .chkCombineCash.Enabled = True Then
23300           .chkCombineCash.SetFocus
23310         ElseIf .chkRelatedAccounts.Enabled = True Then
23320           .chkRelatedAccounts.SetFocus
23330         Else
23340           .opgAccountNumber.SetFocus
23350         End If
23360       End With
23370     End Select
23380   End If

        ' ** Ctrl keys.
23390   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
23400     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23410       With Me
23420         intRetVal = 0
23430         If .opgAccountSource.Enabled = True Then
23440           .opgAccountSource.SetFocus
23450         ElseIf .opgOrderBy.Enabled = True Then
23460           .opgOrderBy.SetFocus
23470         Else
23480           .chkPageOf.SetFocus
23490         End If
23500       End With
23510     End Select
23520   End If

EXITP:
23530   KeyCode = intRetVal
23540   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Public Sub chkIncludeCurrency_AfterUpdate()

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

23610   With Me
23620     Select Case .chkIncludeCurrency
          Case True
23630       .chkIncludeCurrency_lbl.FontBold = True
23640     Case False
23650       .chkIncludeCurrency_lbl.FontBold = False
23660     End Select
23670   End With

EXITP:
23680   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23720   intShiftDown = (Shift And acShiftMask) > 0
23730   intAltDown = (Shift And acAltMask) > 0
23740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23770       With Me
23780         intRetVal = 0
23790         If .cmbAccounts.Enabled = True Then
23800           .opgAccountNumber.SetFocus
23810         ElseIf .cmdTransactionsPreview.Enabled = True Then
23820           .cmdTransactionsPreview.SetFocus
23830         ElseIf .cmdAssetListPreview.Enabled = True Then
23840           .cmdAssetListPreview.SetFocus
23850         Else
23860           .cmdClose.SetFocus
23870         End If
23880       End With
23890     End Select
23900   End If

        ' ** Shift keys.
23910   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23920     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23930       With Me
23940         intRetVal = 0
23950         If .chkForeignExchange.Enabled = True Then
23960           .chkForeignExchange.SetFocus
23970         ElseIf .chkLegalName.Enabled = True Then
23980           .chkLegalName.SetFocus
23990         ElseIf .chkCombineCash.Enabled = True Then
24000           .chkCombineCash.SetFocus
24010         ElseIf .chkRelatedAccounts.Enabled = True Then
24020           .chkRelatedAccounts.SetFocus
24030         Else
24040           .opgAccountNumber.SetFocus
24050         End If
24060       End With
24070     End Select
24080   End If

        ' ** Ctrl keys.
24090   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
24100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24110       With Me
24120         intRetVal = 0
24130         If .opgAccountSource.Enabled = True Then
24140           .opgAccountSource.SetFocus
24150         ElseIf .opgOrderBy.Enabled = True Then
24160           .opgOrderBy.SetFocus
24170         Else
24180           .chkPageOf.SetFocus
24190         End If
24200       End With
24210     End Select
24220   End If

EXITP:
24230   KeyCode = intRetVal
24240   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

24310   AcctSourceAfter_SP Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
24320   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24420   intShiftDown = (Shift And acShiftMask) > 0
24430   intAltDown = (Shift And acAltMask) > 0
24440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24470       With Me
24480         intRetVal = 0
24490         .chkRememberMe.SetFocus
24500       End With
24510     End Select
24520   End If

        ' ** Shift keys.
24530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24550       With Me
24560         intRetVal = 0
24570         If .chkIncludeCurrency.Enabled = True Then
24580           .chkIncludeCurrency.SetFocus
24590         ElseIf .chkForeignExchange.Enabled = True Then
24600           .chkForeignExchange.SetFocus
24610         ElseIf .chkLegalName.Enabled = True Then
24620           .chkLegalName.SetFocus
24630         ElseIf .chkCombineCash.Enabled = True Then
24640           .chkCombineCash.SetFocus
24650         ElseIf .chkRelatedAccounts.Enabled = True Then
24660           .chkRelatedAccounts.SetFocus
24670         Else
24680           .opgAccountNumber.SetFocus
24690         End If
24700       End With
24710     End Select
24720   End If

EXITP:
24730   KeyCode = intRetVal
24740   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24820   intShiftDown = (Shift And acShiftMask) > 0
24830   intAltDown = (Shift And acAltMask) > 0
24840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24870       With Me
24880         intRetVal = 0
24890         .chkRememberMe.SetFocus
24900       End With
24910     End Select
24920   End If

        ' ** Shift keys.
24930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24950       With Me
24960         intRetVal = 0
24970         If .chkIncludeCurrency.Enabled = True Then
24980           .chkIncludeCurrency.SetFocus
24990         ElseIf .chkForeignExchange.Enabled = True Then
25000           .chkForeignExchange.SetFocus
25010         ElseIf .chkLegalName.Enabled = True Then
25020           .chkLegalName.SetFocus
25030         ElseIf .chkCombineCash.Enabled = True Then
25040           .chkCombineCash.SetFocus
25050         ElseIf .chkRelatedAccounts.Enabled = True Then
25060           .chkRelatedAccounts.SetFocus
25070         Else
25080           .opgAccountNumber.SetFocus
25090         End If
25100       End With
25110     End Select
25120   End If

EXITP:
25130   KeyCode = intRetVal
25140   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

25210   Remember_Handler_SP THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
25220   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25320   intShiftDown = (Shift And acShiftMask) > 0
25330   intAltDown = (Shift And acAltMask) > 0
25340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25370       With Me
25380         intRetVal = 0
25390         .cmbAccounts.SetFocus
25400       End With
25410     End Select
25420   End If

        ' ** Shift keys.
25430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25450       With Me
25460         intRetVal = 0
25470         .opgAccountSource.SetFocus
25480       End With
25490     End Select
25500   End If

EXITP:
25510   KeyCode = intRetVal
25520   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

25610   With Me
25620     If .cmbAccounts.text = vbNullString And blnGoingToReport = False Then
25630       ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
25640       .cmbAccounts = .cmbAccounts.ItemData(0)
25650     End If
25660   End With

EXITP:
25670   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

25710   blnContinue = True

        ' ** Where is it getting the arr_varStmt() array?
        ' ** It checks and loads it if necessary.
25720   CmbAccts_After_SP blnContinue, blnAcctNotSched, blnRunPriorStatement, strRememberMe, blnAfterFired, _
          lngStmts, arr_varStmt, lngStmtCnt, blnHasForEx, blnHasForExThis, lngAcctFors, arr_varAcctFor, _
          lngAcctArchs, arr_varAcctArch, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.
        ' ** Call Stack:
        ' **   Calls AcctSched_Load()
        ' **     Runs qrys to load arrays.
        ' **   Calls cmbMonth_AfterUpdate()
        ' **     Calls Month_AfterUpdate_SP()
        ' **       Calls CoOptions_Read()
        ' **         Runs Qrys, initializes variables.
        ' **       DLookup() to "Statement Date"
        ' **       Calls ForEx_ChkScheduled():
        ' **         Calls ForExArr_Load()
        ' **           Runs Qyrs, loads arrays.
        ' **         Calls AcctSched_Load()
        ' **           Runs Qyrs, loads arrays.
        ' **         Calls chkIncludeCurrency_AfterUpdate()
        ' **           Sets Bold.
        ' **   Calls chkIncludeCurrency_AfterUpdate()
        ' **     Sets Bold.
        ' **   Calls SetArchiveOption_SP()
        ' **     Calls GetTPP()
        ' **   Calls SetAccountLastStatement()
        ' **     Runs Qrys.
        ' **   Calls SetRelatedOption()
        ' **     Sets Ctls.

EXITP:
25730   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl: That_Desc = ERR.description
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub cmbAccounts_NotInList(NewData As String, Response As Integer)

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_NotInList"

25810   MsgBox "You must select a valid account.", vbInformation + vbOKOnly, (Left(("Invalid Entry" & Space(40)), 40) & "H01")
25820   Response = acDataErrContinue

EXITP:
25830   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25920   intShiftDown = (Shift And acShiftMask) > 0
25930   intAltDown = (Shift And acAltMask) > 0
25940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25970       With Me
25980         intRetVal = 0
25990         If .chkTransactions = True Then
26000           .cmdTransactionsPreview.SetFocus
26010         ElseIf .chkAssetList = True Then
26020           .cmdAssetListPreview.SetFocus
26030         ElseIf .chkStatements = True Then
26040           If .cmdTransactionsPreview.Enabled = False Then
26050             If blnAcctNotSched = False Then
26060               .cmdTransactionsPreview.Enabled = True
26070               .cmdTransactionsPreview.SetFocus
26080             Else
26090               .opgOrderBy.SetFocus
26100             End If
26110           Else
26120 On Error Resume Next
26130             .cmdTransactionsPreview.SetFocus
26140             If ERR.Number <> 0 Then
26150 On Error GoTo ERRH
26160               .chkStatements.SetFocus
26170               DoEvents
26180               .cmbAccounts.SetFocus
26190             Else
26200 On Error GoTo ERRH
26210             End If
26220           End If
26230         End If
26240       End With
26250     End Select
26260   End If

        ' ** Shift keys.
26270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26290       With Me
26300         intRetVal = 0
26310         .opgAccountNumber.SetFocus
26320       End With
26330     End Select
26340   End If

        ' ** Ctrl keys.
26350   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
26360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26370       With Me
26380         intRetVal = 0
26390         If .opgOrderBy.Enabled = True Then
26400           .opgOrderBy.SetFocus
26410         Else
26420           .chkPageOf.SetFocus
26430         End If
26440       End With
26450     End Select
26460   End If

        ' ** Ctrl-Shift keys.
26470   If intCtrlDown And (Not intAltDown) And intShiftDown Then
26480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26490       With Me
26500         intRetVal = 0
26510         .chkRememberMe.SetFocus
26520       End With
26530     End Select
26540   End If

EXITP:
26550   KeyCode = intRetVal
26560   Exit Sub

ERRH:
1470    intRetVal = 0
1480    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1490    Case Else
1500      THAT_PROC = THIS_PROC
1510      That_Erl = Erl
1520      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub cmbAccounts_Exit(Cancel As Integer)

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Exit"

26610   With Me
26620     If blnAfterFired = False Then
26630       cmbAccounts_AfterUpdate  ' ** Procedure: Above
26640     Else
26650       blnAfterFired = False
26660     End If
26670   End With

EXITP:
26680   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub opgOrderBy_AfterUpdate()

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_AfterUpdate"

26710   With Me
26720     Select Case .opgOrderBy
          Case .opgOrderBy_optDate.OptionValue
26730       .opgOrderBy_optDate_lbl.FontBold = True
26740       .opgOrderBy_optType_lbl.FontBold = False
26750     Case .opgOrderBy_optType.OptionValue
26760       .opgOrderBy_optDate_lbl.FontBold = False
26770       .opgOrderBy_optType_lbl.FontBold = True
26780     End Select
26790   End With

EXITP:
26800   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub opgOrderBy_optDate_KeyDown(KeyCode As Integer, Shift As Integer)

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_optDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26920   intShiftDown = (Shift And acShiftMask) > 0
26930   intAltDown = (Shift And acAltMask) > 0
26940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26970       With Me
26980         intRetVal = 0
26990         .chkPageOf.SetFocus
27000       End With
27010     End Select
27020   End If

        ' ** Shift keys.
27030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27050       With Me
27060         intRetVal = 0
27070         If .cmbAccounts.Enabled = True Then
27080           .cmbAccounts.SetFocus
27090         Else
27100           .opgAccountNumber.SetFocus
27110         End If
27120       End With
27130     End Select
27140   End If

EXITP:
27150   KeyCode = intRetVal
27160   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub opgOrderBy_optType_KeyDown(KeyCode As Integer, Shift As Integer)

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_optType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27220   intShiftDown = (Shift And acShiftMask) > 0
27230   intAltDown = (Shift And acAltMask) > 0
27240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27270       With Me
27280         intRetVal = 0
27290         .chkPageOf.SetFocus
27300       End With
27310     End Select
27320   End If

        ' ** Shift keys.
27330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27350       With Me
27360         intRetVal = 0
27370         If .cmbAccounts.Enabled = True Then
27380           .cmbAccounts.SetFocus
27390         Else
27400           .opgAccountNumber.SetFocus
27410         End If
27420       End With
27430     End Select
27440   End If

EXITP:
27450   KeyCode = intRetVal
27460   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

27510   With Me
27520     Select Case .chkPageOf
          Case True
27530       .chkPageOf_lbl.FontBold = True
27540     Case False
27550       .chkPageOf_lbl.FontBold = False
27560     End Select
27570   End With

EXITP:
27580   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27620   intShiftDown = (Shift And acShiftMask) > 0
27630   intAltDown = (Shift And acAltMask) > 0
27640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27660     Select Case intRetVal
          Case vbKeyTab
27670       With Me
27680         intRetVal = 0
27690         If .cmdTransactionsPreview.Enabled = True Then
27700           .cmdTransactionsPreview.SetFocus
27710         Else
27720           .cmdAssetListPreview.SetFocus
27730         End If
27740       End With
27750     End Select
27760   End If

        ' ** Shift keys.
27770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27780     Select Case intRetVal
          Case vbKeyTab
27790       With Me
27800         intRetVal = 0
27810         If .opgOrderBy.Enabled = True Then
27820           .opgOrderBy.SetFocus
27830         ElseIf .cmbAccounts.Enabled = True Then
27840           .cmbAccounts.SetFocus
27850         Else
27860           .opgAccountNumber.SetFocus
27870         End If
27880       End With
27890     End Select
27900   End If

EXITP:
27910   KeyCode = intRetVal
27920   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub cmdTransactionsPreview_Click()

28000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransactionsPreview_Click"

28010   With Me

28020     CmdTransPreview_Click_SP blnContinue, blnAllStatements, blnFromStmts, strFirstDateMsg, _
            strFileName, strReportName, blnHasForEx, blnHasForExThis, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

28030   End With

EXITP:
28040   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub cmdTransactionsPreview_KeyDown(KeyCode As Integer, Shift As Integer)

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransactionsPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28120   intShiftDown = (Shift And acShiftMask) > 0
28130   intAltDown = (Shift And acAltMask) > 0
28140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28160     Select Case intRetVal
          Case vbKeyTab
28170       With Me
28180         intRetVal = 0
28190         .cmdTransactionsPrint.SetFocus
28200       End With
28210     Case vbKeyLeft
28220       With Me
28230         intRetVal = 0
28240         If .cmdPrintStatement_Summary.Enabled = True Then
28250           .cmdPrintStatement_Summary.SetFocus
28260         ElseIf .cmdPrintStatement_Single.Enabled = True Then
28270           .cmdPrintStatement_Single.SetFocus
28280         ElseIf .cmdPrintStatement_All.Enabled = True Then
28290           .cmdPrintStatement_All.SetFocus
28300         ElseIf .cmdAssetListExcel.Enabled = True Then
28310           .cmdAssetListExcel.SetFocus  '#chk
28320         ElseIf .cmdAssetListWord.Enabled = True Then
28330           .cmdAssetListWord.SetFocus
28340         ElseIf .cmdTransactionsExcel.Enabled = True Then
28350           .cmdTransactionsExcel.SetFocus  '#chk
28360         Else
28370           .cmdTransactionsWord.SetFocus
28380         End If
28390       End With
28400     Case vbKeyRight
28410       With Me
28420         intRetVal = 0
28430         .cmdTransactionsPrint.SetFocus
28440       End With
28450     Case vbKeyUp
28460       With Me
28470         intRetVal = 0
28480         If .cmdPrintStatement_Summary.Enabled = True Then
28490           .cmdPrintStatement_Summary.SetFocus
28500         ElseIf .cmdAssetListExcel.Enabled = True Then
28510           .cmdAssetListExcel.SetFocus  '#chk
28520         ElseIf .cmdPrintStatement_Single.Enabled = True Then
28530           .cmdPrintStatement_Single.SetFocus
28540         ElseIf .cmdTransactionsExcel.Enabled = True Then
28550           .cmdTransactionsExcel.SetFocus  '#chk
28560         Else
28570           .cmdTransactionsPrint.SetFocus
28580         End If
28590       End With
28600     Case vbKeyDown
28610       With Me
28620         intRetVal = 0
28630         .cmdTransactionsWord.SetFocus
28640       End With
28650     End Select
28660   End If

        ' ** Shift keys.
28670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28680     Select Case intRetVal
          Case vbKeyTab
28690       With Me
28700         intRetVal = 0
28710         If .cmbAccounts.Enabled = True Then
28720           .cmbAccounts.SetFocus
28730         Else
28740           .opgAccountNumber.SetFocus
28750         End If
28760       End With
28770     End Select
28780   End If

        ' ** Ctrl-Shift keys.
28790   If intCtrlDown And (Not intAltDown) And intShiftDown Then
28800     Select Case intRetVal
          Case vbKeyTab
28810       With Me
28820         intRetVal = 0
28830         .chkPageOf.SetFocus
28840       End With
28850     End Select
28860   End If

EXITP:
28870   KeyCode = intRetVal
28880   Exit Sub

ERRH:
1470    intRetVal = 0
1480    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1490    Case Else
1500      THAT_PROC = THIS_PROC
1510      That_Erl = Erl
1520      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub cmdTransactionsPreview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ** Easter Egg (Secret Key) to invoke Transaction Preview.

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransactionsPreview_MouseUp"

        ' ** This was added to assist when users do things out of sequence.
28910   If Button = acRightButton And Shift = acCtrlMask + acAltMask Then
28920     cmdTransactionsPreview_Click  ' ** Procedure: Above.
28930   End If

EXITP:
28940   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Public Sub cmdTransactionsPrint_Click()
' ** Also called by:
' **   Statements_Print(), Below.

29000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransactionsPrint_Click"

29010   CmdTransPrint_Click_SP blnContinue, blnAllStatements, blnFromStmts, strFirstDateMsg, strFileName, _
          strReportName, blnHasForEx, blnHasForExThis, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
29020   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    DoCmd.SetWarnings True
1490    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
1500      blnContinue = False
1510      If Reports.Count > 0 Then
1520        DoCmd.Close acReport, Reports(0).Name   ' ** Close report in preview.
1530      End If
1540    Case Else
1550      THAT_PROC = THIS_PROC
1560      That_Erl = Erl: That_Desc = ERR.description
1570      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1580    End Select
1590    Resume EXITP

End Sub

Private Sub cmdTransactionsPrint_KeyDown(KeyCode As Integer, Shift As Integer)

29100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransactionsPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29120   intShiftDown = (Shift And acShiftMask) > 0
29130   intAltDown = (Shift And acAltMask) > 0
29140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29160     Select Case intRetVal
          Case vbKeyTab
29170       With Me
29180         intRetVal = 0
29190         .cmdTransactionsWord.SetFocus
29200       End With
29210     Case vbKeyLeft
29220       With Me
29230         intRetVal = 0
29240         .cmdTransactionsPreview.SetFocus
29250       End With
29260     Case vbKeyRight
29270       With Me
29280         intRetVal = 0
29290         If .cmdAssetListPreview.Enabled = True Then
29300           .cmdAssetListPreview.SetFocus
29310         Else
29320           .cmdTransactionsWord.SetFocus
29330         End If
29340       End With
29350     Case vbKeyUp
29360       With Me
29370         intRetVal = 0
29380         If .cmdPrintStatement_All.Enabled = True Then
29390           .cmdPrintStatement_All.SetFocus
29400         Else
29410           .cmdTransactionsWord.SetFocus
29420         End If
29430       End With
29440     Case vbKeyDown
29450       With Me
29460         intRetVal = 0
29470         If .cmdTransactionsExcel.Enabled = True Then
29480           .cmdTransactionsExcel.SetFocus  '#chk
29490         ElseIf .cmdAssetListPreview.Enabled = True Then
29500           .cmdAssetListPreview.SetFocus
29510         ElseIf .cmdPrintStatement_Single.Enabled = True Then
29520           .cmdPrintStatement_Single.SetFocus
29530         ElseIf .cmdPrintStatement_Summary.Enabled = True Then
29540           .cmdPrintStatement_Summary.SetFocus
29550         Else
29560           .cmdTransactionsPreview.SetFocus
29570         End If
29580       End With
29590     End Select
29600   End If

        ' ** Shift keys.
29610   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29620     Select Case intRetVal
          Case vbKeyTab
29630       With Me
29640         intRetVal = 0
29650         .cmdTransactionsPreview.SetFocus
29660       End With
29670     End Select
29680   End If

EXITP:
29690   KeyCode = intRetVal
29700   Exit Sub

ERRH:
1470    intRetVal = 0
1480    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1490    Case Else
1500      THAT_PROC = THIS_PROC
1510      That_Erl = Erl
1520      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub cmdTransactionsWord_Click()

29800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransactionsWord_Click"

29810   CmdTransWord_Click_SP blnContinue, blnAllStatements, blnFromStmts, strFirstDateMsg, strFileName, _
          strReportName, blnHasForEx, blnHasForExThis, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
29820   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
1490      blnContinue = False
1500    Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdTransactionsWord_KeyDown(KeyCode As Integer, Shift As Integer)

29900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransactionsWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29920   intShiftDown = (Shift And acShiftMask) > 0
29930   intAltDown = (Shift And acAltMask) > 0
29940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29960     Select Case intRetVal
          Case vbKeyTab
29970       With Me
29980         intRetVal = 0
29990         If .cmdTransactionsExcel.Enabled = True Then
30000           .cmdTransactionsExcel.SetFocus  '#chk
30010         ElseIf .cmdAssetListPreview.Enabled = True Then
30020           .cmdAssetListPreview.SetFocus
30030         ElseIf .cmdPrintStatement_All.Enabled = True Then
30040           .cmdPrintStatement_All.SetFocus
30050         ElseIf .cmdPrintStatement_Single.Enabled = True Then
30060           .cmdPrintStatement_Single.SetFocus
30070         ElseIf .cmdPrintStatement_Summary.Enabled = True Then
30080           .cmdPrintStatement_Summary.SetFocus
30090         Else
30100           .cmdClose.SetFocus
30110         End If
30120       End With
30130     Case vbKeyLeft
30140       With Me
30150         intRetVal = 0
30160         If .cmdAssetListPrint.Enabled = True Then
30170           .cmdAssetListPrint.SetFocus
30180         Else
30190           .cmdTransactionsPrint.SetFocus
30200         End If
30210       End With
30220     Case vbKeyRight
30230       With Me
30240         intRetVal = 0
30250         If .cmdTransactionsExcel.Enabled = True Then
30260           .cmdTransactionsExcel.SetFocus  '#chk
30270         ElseIf .cmdAssetListWord.Enabled = True Then
30280           .cmdAssetListWord.SetFocus
30290         ElseIf .cmdPrintStatement_All.Enabled = True Then
30300           .cmdPrintStatement_All.SetFocus
30310         ElseIf .cmdPrintStatement_Single.Enabled = True Then
30320           .cmdPrintStatement_Single.SetFocus
30330         ElseIf .cmdPrintStatement_Summary.Enabled = True Then
30340           .cmdPrintStatement_Summary.SetFocus
30350         Else
30360           .cmdTransactionsPreview.SetFocus
30370         End If
30380       End With
30390     Case vbKeyUp
30400       With Me
30410         intRetVal = 0
30420         .cmdTransactionsPreview.SetFocus
30430       End With
30440     Case vbKeyDown
30450       With Me
30460         intRetVal = 0
30470         If .cmdPrintStatement_All.Enabled = True Then
30480           .cmdPrintStatement_All.SetFocus
30490         Else
30500           .cmdTransactionsPrint.SetFocus
30510         End If
30520       End With
30530     End Select
30540   End If

        ' ** Shift keys.
30550   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30560     Select Case intRetVal
          Case vbKeyTab
30570       With Me
30580         intRetVal = 0
30590         .cmdTransactionsPrint.SetFocus
30600       End With
30610     End Select
30620   End If

EXITP:
30630   KeyCode = intRetVal
30640   Exit Sub

ERRH:
1470    intRetVal = 0
1480    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1490    Case Else
1500      THAT_PROC = THIS_PROC
1510      That_Erl = Erl
1520      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub cmdTransactionsExcel_Click()

30700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransactionsExcel_Click"

30710   Transactions_Excel Me, blnContinue, blnFromStmts, strFirstDateMsg, strFileName, strReportName, _
          blnAllStatements, blnHasForEx, blnHasForExThis  ' ** Module Procedure: modStatementParamFuncs1.

EXITP:
30720   DoCmd.SetWarnings True
30730   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
1490      blnContinue = False
1500    Case Else
1510      THAT_PROC = THIS_PROC
1520      That_Erl = Erl
1530      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1540    End Select
1550    Resume EXITP

End Sub

Private Sub cmdTransactionsExcel_KeyDown(KeyCode As Integer, Shift As Integer)

30800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransactionsExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30820   intShiftDown = (Shift And acShiftMask) > 0
30830   intAltDown = (Shift And acAltMask) > 0
30840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30860     Select Case intRetVal
          Case vbKeyTab
30870       With Me
30880         intRetVal = 0
30890         If .cmdAssetListPreview.Enabled = True Then
30900           .cmdAssetListPreview.SetFocus
30910         ElseIf .cmdPrintStatement_All.Enabled = True Then
30920           .cmdPrintStatement_All.SetFocus
30930         ElseIf .cmdPrintStatement_Single.Enabled = True Then
30940           .cmdPrintStatement_Single.SetFocus
30950         ElseIf .cmdPrintStatement_Summary.Enabled = True And .cmdPrintStatement_Summary.Visible = True Then
30960           .cmdPrintStatement_Summary.SetFocus
30970         Else
30980           .cmdClose.SetFocus
30990         End If
31000       End With
31010     Case vbKeyLeft
31020       With Me
31030         intRetVal = 0
31040         .cmdTransactionsWord.SetFocus
31050       End With
31060     Case vbKeyRight
31070       With Me
31080         intRetVal = 0
31090         If .cmdAssetListWord.Enabled = True Then
31100           .cmdAssetListWord.SetFocus
31110         ElseIf .cmdPrintStatement_All.Enabled = True Then
31120           .cmdPrintStatement_All.SetFocus
31130         ElseIf .cmdPrintStatement_Single.Enabled = True Then
31140           .cmdPrintStatement_Single.SetFocus
31150         ElseIf .cmdPrintStatement_Summary.Enabled = True Then
31160           .cmdPrintStatement_Summary.SetFocus
31170         Else
31180           .cmdTransactionsPreview.SetFocus
31190         End If
31200       End With
31210     Case vbKeyUp
31220       With Me
31230         intRetVal = 0
31240         .cmdTransactionsPrint.SetFocus
31250       End With
31260     Case vbKeyDown
31270       With Me
31280         intRetVal = 0
31290         If .cmdAssetListPreview.Enabled = True Then
31300           .cmdAssetListPreview.SetFocus
31310         ElseIf .cmdPrintStatement_Single.Enabled = True Then
31320           .cmdPrintStatement_Single.SetFocus
31330         ElseIf .cmdPrintStatement_Summary.Enabled = True Then
31340           .cmdPrintStatement_Summary.SetFocus
31350         Else
31360           .cmdTransactionsPreview.SetFocus
31370         End If
31380       End With
31390     End Select
31400   End If

        ' ** Shift keys.
31410   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31420     Select Case intRetVal
          Case vbKeyTab
31430       With Me
31440         intRetVal = 0
31450         .cmdTransactionsWord.SetFocus
31460       End With
31470     End Select
31480   End If

EXITP:
31490   KeyCode = intRetVal
31500   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub cmdAssetListPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

31600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListPreview_MouseDown"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer

31610   With Me

31620     DoCmd.Hourglass True
31630     DoEvents

31640     intShiftDown = (Shift And acShiftMask) > 0
31650     intAltDown = (Shift And acAltMask) > 0
31660     intCtrlDown = (Shift And acCtrlMask) > 0

31670     If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue Then
31680       If IsNull(.cmbAccounts) = False Then
31690         gstrAccountNo = .cmbAccounts
              ' ** Ctrl keys.
31700         If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
31710           blnHasForExClick = True
31720           If Button = acLeftButton Then
31730             gblnSwitchTo = HasForEx_Suppress(gstrAccountNo)  ' ** Module Function: modCurrencyFuncs.
31740             If gblnSwitchTo = True Then
                    ' ** Remove the ForEx column preference.
31750               Set dbs = CurrentDb
31760               With dbs
                      ' ** tblCurrency_Account for curracct_supress = False, by specified [actno].
31770                 Set qdf = .QueryDefs("qryCurrency_17_02")
31780                 With qdf.Parameters
31790                   ![actno] = gstrAccountNo
31800                 End With
31810                 qdf.Execute
31820                 Set qdf = Nothing
31830                 .Close
31840               End With
31850               Beep
31860             End If
31870           End If
31880         End If
31890       End If
31900     End If

31910   End With

        ' ** AcButton enumeration:
        ' **   1  acLeftButton
        ' **   2  acRightButton
        ' **   4  acMiddleButton

EXITP:
31920   Set qdf = Nothing
31930   Set dbs = Nothing
31940   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub cmdAssetListPreview_Click()

32000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListPreview_Click"

32010   CmdAstListPreview_Click_SP blnContinue, datAssetListDate, blnPrintAnnualStatement, blnAllStatements, _
          blnNoDataAll, blnRollbackNeeded, strFirstDateMsg, strReportName, blnHasForExClick, blnIncludeCurrency, _
          blnHasForEx, blnHasForExThis, lngAcctFors, arr_varAcctFor, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
32020   Exit Sub

ERRH:
1470    blnContinue = False
1480    DoCmd.Hourglass False
1490    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
1500    Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdAssetListPreview_KeyDown(KeyCode As Integer, Shift As Integer)

32100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
32120   intShiftDown = (Shift And acShiftMask) > 0
32130   intAltDown = (Shift And acAltMask) > 0
32140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32160     Select Case intRetVal
          Case vbKeyTab
32170       With Me
32180         intRetVal = 0
32190         .cmdAssetListPrint.SetFocus
32200       End With
32210     Case vbKeyLeft
32220       With Me
32230         intRetVal = 0
32240         If .cmdTransactionsPrint.Enabled = True Then
32250           .cmdTransactionsPrint.SetFocus
32260         ElseIf .cmdPrintStatement_Summary.Enabled = True Then
32270           .cmdPrintStatement_Summary.SetFocus
32280         ElseIf .cmdPrintStatement_Single.Enabled = True Then
32290           .cmdPrintStatement_Single.SetFocus
32300         ElseIf .cmdPrintStatement_All.Enabled = True Then
32310           .cmdPrintStatement_All.SetFocus
32320         ElseIf .cmdAssetListExcel.Enabled = True Then
32330           .cmdAssetListExcel.SetFocus  '#chk
32340         Else
32350           .cmdAssetListWord.SetFocus
32360         End If
32370       End With
32380     Case vbKeyRight
32390       With Me
32400         intRetVal = 0
32410         .cmdAssetListPrint.SetFocus
32420       End With
32430     Case vbKeyUp
32440       With Me
32450         intRetVal = 0
32460         If .cmdTransactionsExcel.Enabled = True Then
32470           .cmdTransactionsExcel.SetFocus  '#chk
32480         ElseIf .cmdTransactionsPrint.Enabled = True Then
32490           .cmdTransactionsPrint.SetFocus
32500         ElseIf .cmdPrintStatement_All.Enabled = True Then
32510           .cmdPrintStatement_All.SetFocus
32520         ElseIf .cmdPrintStatement_Summary.Enabled = True Then
32530           .cmdPrintStatement_Summary.SetFocus
32540         ElseIf .cmdAssetListExcel.Enabled = True Then
32550           .cmdAssetListExcel.SetFocus  '#chk
32560         Else
32570           .cmdAssetListPrint.SetFocus
32580         End If
32590       End With
32600     Case vbKeyDown
32610       With Me
32620         intRetVal = 0
32630         .cmdAssetListWord.SetFocus
32640       End With
32650     End Select
32660   End If

        ' ** Shift keys.
32670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32680     Select Case intRetVal
          Case vbKeyTab
32690       With Me
32700         intRetVal = 0
32710         If .cmdTransactionsExcel.Enabled = True Then
32720           .cmdTransactionsExcel.SetFocus  '#chk
32730         ElseIf .cmdTransactionsWord.Enabled = True Then
32740           .cmdTransactionsWord.SetFocus
32750         ElseIf .cmbAccounts.Enabled = True Then
32760           .cmbAccounts.SetFocus
32770         Else
32780           .opgAccountNumber.SetFocus
32790         End If
32800       End With
32810     End Select
32820   End If

        ' ** Ctrl-Shift keys.
32830   If intCtrlDown And (Not intAltDown) And intShiftDown Then
32840     Select Case intRetVal
          Case vbKeyTab
32850       With Me
32860         intRetVal = 0
32870         If .cmdTransactionsExcel.Enabled = False And .cmdTransactionsWord.Enabled = False Then
32880           .chkPageOf.SetFocus
32890         End If
32900       End With
32910     End Select
32920   End If

EXITP:
32930   KeyCode = intRetVal
32940   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub cmdAssetListPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

33000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListPrint_MouseDown"

33010   cmdAssetListPreview_MouseDown Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
33020   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Public Sub cmdAssetListPrint_Click()
' ** Also called by:
' **   Statements_Print(), Below.

33100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListPrint_Click"

33110   CmdAstListPrint_Click_SP blnContinue, blnFromStmts, blnPrintAnnualStatement, blnAllStatements, _
          blnRollbackNeeded, blnNoDataAll, datAssetListDate, strFirstDateMsg, strReportName, _
          blnHasForExClick, blnIncludeCurrency, blnHasForEx, blnHasForExThis, lngAcctFors, arr_varAcctFor, _
          THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
33120   Exit Sub

ERRH:
1470    DoCmd.SetWarnings True
1480    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
1490    Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
1500      blnContinue = False
1510    Case Else
1520      blnContinue = False
1530      THAT_PROC = THIS_PROC
1540      That_Erl = Erl
1550      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1560    End Select
1570    Resume EXITP

End Sub

Private Sub cmdAssetListPrint_KeyDown(KeyCode As Integer, Shift As Integer)

33200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

33210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
33220   intShiftDown = (Shift And acShiftMask) > 0
33230   intAltDown = (Shift And acAltMask) > 0
33240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
33250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
33260     Select Case intRetVal
          Case vbKeyTab
33270       With Me
33280         intRetVal = 0
33290         .cmdAssetListWord.SetFocus
33300       End With
33310     Case vbKeyLeft
33320       With Me
33330         intRetVal = 0
33340         .cmdAssetListPreview.SetFocus
33350       End With
33360     Case vbKeyRight
33370       With Me
33380         intRetVal = 0
33390         If .cmdTransactionsWord.Enabled = True Then
33400           .cmdTransactionsWord.SetFocus
33410         Else
33420           .cmdAssetListWord.SetFocus
33430         End If
33440       End With
33450     Case vbKeyUp
33460       With Me
33470         intRetVal = 0
33480         If .cmdPrintStatement_Single.Enabled = True Then
33490           .cmdPrintStatement_Single.SetFocus
33500         Else
33510           .cmdAssetListWord.SetFocus
33520         End If
33530       End With
33540     Case vbKeyDown
33550       With Me
33560         intRetVal = 0
33570         If .cmdAssetListExcel.Enabled = True Then
33580           .cmdAssetListExcel.SetFocus  '#chk
33590         ElseIf .cmdPrintStatement_Summary.Enabled = True Then
33600           .cmdPrintStatement_Summary.SetFocus
33610         ElseIf .cmdTransactionsPreview.Enabled = True Then
33620           .cmdTransactionsPreview.SetFocus
33630         ElseIf .cmdPrintStatement_All.Enabled = True Then
33640           .cmdPrintStatement_All.SetFocus
33650         Else
33660           .cmdAssetListPreview.SetFocus
33670         End If
33680       End With
33690     End Select
33700   End If

        ' ** Shift keys.
33710   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33720     Select Case intRetVal
          Case vbKeyTab
33730       With Me
33740         intRetVal = 0
33750         .cmdAssetListPreview.SetFocus
33760       End With
33770     End Select
33780   End If

EXITP:
33790   KeyCode = intRetVal
33800   Exit Sub

ERRH:
1470    intRetVal = 0
1480    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1490    Case Else
1500      THAT_PROC = THIS_PROC
1510      That_Erl = Erl
1520      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub cmdAssetListWord_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

33900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListWord_MouseDown"

33910   cmdAssetListPreview_MouseDown Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
33920   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub cmdAssetListWord_Click()

34000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListWord_Click"

34010   CmdAstListWord_Click_SP blnContinue, blnAllStatements, blnPrintAnnualStatement, blnNoDataAll, _
          blnRollbackNeeded, datAssetListDate, strFirstDateMsg, blnHasForExClick, blnIncludeCurrency, _
          blnHasForEx, blnHasForExThis, lngAcctFors, arr_varAcctFor, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
34020   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
1480      blnContinue = False
1490    Case Else
1500      THAT_PROC = THIS_PROC
1510      That_Erl = Erl
1520      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub cmdAssetListWord_KeyDown(KeyCode As Integer, Shift As Integer)

34100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

34110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
34120   intShiftDown = (Shift And acShiftMask) > 0
34130   intAltDown = (Shift And acAltMask) > 0
34140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
34150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34160     Select Case intRetVal
          Case vbKeyTab
34170       With Me
34180         intRetVal = 0
34190         If .cmdAssetListExcel.Enabled = True Then
34200           .cmdAssetListExcel.SetFocus  '#chk
34210         ElseIf .cmdPrintStatement_All.Enabled = True Then
34220           .cmdPrintStatement_All.SetFocus
34230         ElseIf .cmdPrintStatement_Single.Enabled = True Then
34240           .cmdPrintStatement_Single.SetFocus
34250         ElseIf .cmdPrintStatement_Summary.Enabled = True And .cmdPrintStatement_Summary.Visible = True Then
34260           .cmdPrintStatement_Summary.SetFocus
34270         Else
34280           .cmdClose.SetFocus
34290         End If
34300       End With
34310     Case vbKeyLeft
34320       With Me
34330         intRetVal = 0
34340         If .cmdTransactionsExcel.Enabled = True Then
34350           .cmdTransactionsExcel.SetFocus
34360         ElseIf .cmdTransactionsWord.Enabled = True Then
34370           .cmdTransactionsWord.SetFocus
34380         Else
34390           .cmdAssetListPrint.SetFocus
34400         End If
34410       End With
34420     Case vbKeyRight
34430       With Me
34440         intRetVal = 0
34450         If .cmdAssetListExcel.Enabled = True Then
34460           .cmdAssetListExcel.SetFocus  '#chk
34470         ElseIf .cmdPrintStatement_All.Enabled = True Then
34480           .cmdPrintStatement_All.SetFocus
34490         ElseIf .cmdPrintStatement_Single.Enabled = True Then
34500           .cmdPrintStatement_Single.SetFocus
34510         ElseIf .cmdPrintStatement_Summary.Enabled = True Then
34520           .cmdPrintStatement_Summary.SetFocus
34530         ElseIf .cmdTransactionsPreview.Enabled = True Then
34540           .cmdTransactionsPreview.SetFocus
34550         Else
34560           .cmdAssetListPreview.SetFocus
34570         End If
34580       End With
34590     Case vbKeyUp
34600       With Me
34610         intRetVal = 0
34620         .cmdAssetListPreview.SetFocus
34630       End With
34640     Case vbKeyDown
34650       With Me
34660         intRetVal = 0
34670         If .cmdPrintStatement_Single.Enabled = True Then
34680           .cmdPrintStatement_Single.SetFocus
34690         Else
34700           .cmdAssetListPrint.SetFocus
34710         End If
34720       End With
34730     End Select
34740   End If

        ' ** Shift keys.
34750   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34760     Select Case intRetVal
          Case vbKeyTab
34770       With Me
34780         intRetVal = 0
34790         .cmdAssetListPrint.SetFocus
34800       End With
34810     End Select
34820   End If

EXITP:
34830   KeyCode = intRetVal
34840   Exit Sub

ERRH:
1470    intRetVal = 0
1480    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1490    Case Else
1500      THAT_PROC = THIS_PROC
1510      That_Erl = Erl
1520      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub cmdAssetListExcel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

34900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListExcel_MouseDown"

34910   cmdAssetListPreview_MouseDown Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
34920   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub cmdAssetListExcel_Click()

35000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListExcel_Click"

35010   AssetList_Excel_SP Me, blnContinue, strFirstDateMsg, datAssetListDate, blnPrintAnnualStatement, _
          blnAllStatements, blnNoDataAll, blnRollbackNeeded, blnHasForExClick, blnHasForEx, blnHasForExThis  ' ** Module Procedure: modStatementParamFuncs1.

EXITP:
35020   DoCmd.SetWarnings True
35030   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
1480      blnContinue = False
1490    Case Else
1500      THAT_PROC = THIS_PROC
1510      That_Erl = Erl
1520      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub cmdAssetListExcel_KeyDown(KeyCode As Integer, Shift As Integer)

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetListExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

35110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
35120   intShiftDown = (Shift And acShiftMask) > 0
35130   intAltDown = (Shift And acAltMask) > 0
35140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
35150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
35160     Select Case intRetVal
          Case vbKeyTab
35170       With Me
35180         intRetVal = 0
35190         If .cmdPrintStatement_All.Enabled = True Then
35200           .cmdPrintStatement_All.SetFocus
35210         ElseIf .cmdPrintStatement_Single.Enabled = True Then
35220           .cmdPrintStatement_Single.SetFocus
35230         ElseIf .cmdPrintStatement_Summary.Enabled = True And .cmdPrintStatement_Summary.Visible = True Then
35240           .cmdPrintStatement_Summary.SetFocus
35250         Else
35260           .cmdClose.SetFocus
35270         End If
35280       End With
35290     Case vbKeyLeft
35300       With Me
35310         intRetVal = 0
35320         .cmdAssetListWord.SetFocus
35330       End With
35340     Case vbKeyRight
35350       With Me
35360         intRetVal = 0
35370         If .cmdPrintStatement_All.Enabled = True Then
35380           .cmdPrintStatement_All.SetFocus
35390         ElseIf .cmdPrintStatement_Single.Enabled = True Then
35400           .cmdPrintStatement_Single.SetFocus
35410         ElseIf .cmdPrintStatement_Summary.Enabled = True Then
35420           .cmdPrintStatement_Summary.SetFocus
35430         ElseIf .cmdTransactionsPreview.Enabled = True Then
35440           .cmdTransactionsPreview.SetFocus
35450         Else
35460           .cmdAssetListPreview.SetFocus
35470         End If
35480       End With
35490     Case vbKeyUp
35500       With Me
35510         intRetVal = 0
35520         .cmdAssetListPrint.SetFocus
35530       End With
35540     Case vbKeyDown
35550       With Me
35560         intRetVal = 0
35570         If .cmdPrintStatement_Summary.Enabled = True Then
35580           .cmdPrintStatement_Summary.SetFocus
35590         ElseIf .cmdTransactionsPreview.Enabled = True Then
35600           .cmdTransactionsPreview.SetFocus
35610         ElseIf .cmdPrintStatement_All.Enabled = True Then
35620           .cmdPrintStatement_All.SetFocus
35630         Else
35640           .cmdAssetListPreview.SetFocus
35650         End If
35660       End With
35670     End Select
35680   End If

        ' ** Shift keys.
35690   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35700     Select Case intRetVal
          Case vbKeyTab
35710       With Me
35720         intRetVal = 0
35730         .cmdAssetListWord.SetFocus
35740       End With
35750     End Select
35760   End If

EXITP:
35770   KeyCode = intRetVal
35780   Exit Sub

ERRH:
1470    intRetVal = 0
1480    THAT_PROC = THIS_PROC
1490    That_Erl = Erl
1500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1510    Resume EXITP

End Sub

Private Sub cmdPrintStatement_All_Click()

35800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintStatement_All_Click"

35810   blnContinue = True  ' ** Unless user cancels.

35820   CmdPrintStmtAll_Click_SP blnContinue, blnAllStatements, blnPrintStatements, blnSingleStatement, _
          blnRunPriorStatement, blnFromStmts, blnAcctsSchedRpt, datFirstDate, strFirstDateMsg, blnGTR_Emblem, _
          blnGoingToReport, blnGoingToReport2, blnWasGTR, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
35830   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
1480    Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub cmdPrintStatement_All_KeyDown(KeyCode As Integer, Shift As Integer)

35900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintStatement_All_KeyDown"

35910   PrintStmt_KeyDown_SP KeyCode, Shift, THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
35920   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdPrintStatement_Single_Click()

36000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintStatement_Single_Click"

36010   CmdPrintStmtSingle_Click_SP blnContinue, blnPrintStatements, blnAllStatements, blnSingleStatement, _
          blnRunPriorStatement, blnAcctsSchedRpt, datFirstDate, blnFromStmts, blnGoingToReport, _
          blnGoingToReport2, blnGTR_Emblem, blnWasGTR, strFirstDateMsg, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
36020   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
1480    Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub cmdPrintStatement_Single_KeyDown(KeyCode As Integer, Shift As Integer)

36100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintStatement_Single_KeyDown"

36110   PrintStmt_KeyDown_SP KeyCode, Shift, THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
36120   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub cmdPrintStatement_Summary_Click()

36200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintStatement_Summary_Click"

36210   CmdPrintStmtSum_Click_SP blnContinue, blnFromStmts, strFirstDateMsg, THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modStatementParamFuncs2.

EXITP:
36220   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
1490    Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrintStatement_Summary_KeyDown(KeyCode As Integer, Shift As Integer)

36300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintStatement_Summary_KeyDown"

36310   PrintStmt_KeyDown_SP KeyCode, Shift, THIS_PROC, Me  ' ** Module Procedure: modStatementParamFuncs3.

EXITP:
36320   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkAsDev_AfterUpdate()

36400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAsDev_AfterUpdate"

36410   With Me
36420     Select Case .chkAsDev
          Case True
36430       .chkAsDev_lbl.FontBold = True
36440     Case False
36450       .chkAsDev_lbl.FontBold = False
36460     End Select
36470   End With

EXITP:
36480   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkDevMsg_AfterUpdate()

36500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkDevMsg_AfterUpdate"

36510   With Me
36520     Select Case .chkDevMsg
          Case True
36530       .chkDevMsg_lbl.FontBold = True
36540     Case False
36550       .chkDevMsg_lbl.FontBold = False
36560     End Select
36570   End With

EXITP:
36580   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Public Sub cmdDevCloseReports_Click()

36600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDevCloseReports_Click"

36610   If Reports.Count > 0& Then
36620     Do While Reports.Count > 0&
36630       DoCmd.Close acReport, Reports(0).Name
36640       DoEvents
36650     Loop
36660   End If

EXITP:
36670   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub chkDevSaveData_AfterUpdate()

36700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkDevSaveData_AfterUpdate"

36710   With Me
36720     Select Case .chkDevSaveData
          Case True
36730       .chkDevSaveData_lbl.FontBold = True
36740     Case False
36750       .chkDevSaveData_lbl.FontBold = False
36760     End Select
36770   End With

EXITP:
36780   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Public Sub cmdSummaryPrint()
' ** Called by:
' **   Statements_Print(), Below.
' **   cmdPrintStatement_Summary_Click(), Above.
' ** (Not really a command button, but I was organizing...)

36800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSummaryPrint"

36810   AccountSummary_Print Me, blnContinue, blnFromStmts, strReportName, blnHasForEx, blnHasForExThis  ' ** Module Procedure: modStatementParamFuncs1.

EXITP:
36820   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
1490      blnContinue = False
1500      If Reports.Count > 0 Then
1510        DoCmd.Close acReport, Reports(0).Name  ' ** Close report in preview.
1520      End If
1530    Case Else
1540      THAT_PROC = THIS_PROC
1550      That_Erl = Erl
1560      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1570      Resume EXITP
1580    End Select
1590    Resume EXITP

End Sub

Public Sub SetPrintStatements(blnPrint As Boolean)
' ** Called by:
' **   frmStatementParameters_Print:
' **     cmdYes_Click()
' **     cmdNo_Click()

36900 On Error GoTo ERRH

        Const THIS_PROC As String = "SetPrintStatements"

36910   blnPrintStatements = blnPrint

EXITP:
36920   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Private Sub AcctNumLock(blnLock As Boolean)

37000 On Error GoTo ERRH

        Const THIS_PROC As String = "AcctNumLock"

37010   With Me
37020     .opgAccountNumber.Locked = blnLock
37030     .cmbAccounts.Locked = blnLock
37040   End With

EXITP:
37050   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Public Sub AcctsSchedRpt_Set(blnToRpt As Boolean)

37100 On Error GoTo ERRH

        Const THIS_PROC As String = "AcctsSchedRpt_Set"

37110   blnAcctsSchedRpt = blnToRpt

EXITP:
37120   Exit Sub

ERRH:
1470    THAT_PROC = THIS_PROC
1480    That_Erl = Erl
1490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1500    Resume EXITP

End Sub

Public Sub GTREmblem_Off()

37200 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

37210   With Me
37220     blnGTR_Emblem = False
37230     For lngX = 1& To 24&
37240       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
37250     Next
37260   End With

EXITP:
37270   Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Public Sub FirstDateMsg_Set(strMsg As String)

37300 On Error GoTo ERRH

        Const THIS_PROC As String = "FirstDateMsg_Set"

37310   strFirstDateMsg = strMsg

EXITP:
37320   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Public Sub FirstDate_Set(datDate As Date)

37400 On Error GoTo ERRH

        Const THIS_PROC As String = "FirstDate_Set"

37410   datFirstDate = datDate

EXITP:
37420   Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case 13  ' ** Type mismatch.
          ' ** Ignore
1490    Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub
