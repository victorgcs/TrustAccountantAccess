VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Report"

'VGC 10/11/2017: CHANGES!
' ** TimerInterval: 100

' ** For GTR, shows arrows, sets focus, then continues.

' ** Shortcut Alt keys responsive from this form:
' **   Account Balance:  A {cmdAccountBalance}
' **   Trans By Type:    B {cmdTransByType}
' **   Close:            C {cmdClose}
' **   Load Timer:       D {chkLoadTimer}
' **   Search:           E {cmdSearch}
' **   Stmt of Cond:     F {cmdStatementOfCondition}
' **   Cap Gain/Loss:    G {cmdCapitalGainsAndLosses}
' **   Purchased/Sold:   H {cmdAssetPurchasedSold}
' **   Income Stmt:      I {cmdIncomeStatement}
' **   Tax Lot:          L {cmdTaxLot}
' **   Maturity:         M {cmdSecurityMaturity}
' **   New/Closed:       N {cmdAcctNewClosed}
' **   Cash Control:     O {cmdCashControl}
' **   Account Profile:  P {cmdAccountProfile}
' **   Income/Expense:   R {cmdRevIncExp}
' **   Statements:       S {cmdStatements}
' **   Tax:              T {cmdTaxIncDed}
' **   Court:            U {cmdCourtReports}
' **   Account Review:   W {cmdAccountReviews}
' **   Exit:             X {cmdClose}
' **   Unreal Gain/Loss: Z {cmdUnrealizedGainAndLoss}

'Available keys:
'J K Q V Y

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm}

' ** Array: arr_varCmd().
Private lngCmds As Long, arr_varCmd() As Variant
Private Const C_ELEMS As Integer = 4
Private Const C_NAM As Integer = 0
Private Const C_LFT As Integer = 1
Private Const C_TOP As Integer = 2
Private Const C_COL As Integer = 3
Private Const C_ORD As Integer = 4

Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnJustClose As Boolean, blnFormOpening As Boolean, blnJustLooking As Boolean
Private lngLeftCmd_Left As Long, lngRightCmd_Left As Long, lngMaxSortOrd As Long
Private lngTopCmd_Top As Long, lngBotCmd_Top As Long, blnTimerVis As Boolean
Private strDocName As String, strLastFocus As String, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim ctl As Access.Control
        Dim lngTmp01 As Long, lngTmp02 As Long, blnTmp03 As Boolean
        Dim lngX As Long, lngY As Long, lngE As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           DoCmd.Hourglass True
160           DoEvents

170           gblnSignal = False

180           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

190           gstrCrtRpt_Version = vbNullString  ' ** Borrowing this variable from Court Reports.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
200           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

210           blnTimerVis = False
220           Select Case .chkLoadTimer_vis
              Case True
230             .chkLoadTimer.Visible = True
240             .chkLoadTimer_lbl1.Visible = True
250             .chkLoadTimer_lbl2.Visible = True
260             blnTimerVis = True
270           Case False
280             .chkLoadTimer.Visible = False
290             .chkLoadTimer_lbl1.Visible = False
300             .chkLoadTimer_lbl2.Visible = False
310           End Select
320           chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
330           DoEvents

340           blnJustClose = False
350           blnGTR_Emblem = False: lngGTR_Stat = 0&
360           If gblnGoToReport = True Then
370             blnJustLooking = False
380             .cmdSearch_lbl1.Visible = False
390             .cmdSearch_lbl2.Visible = False
400             .cmdSearch_arw.Visible = False
410           Else
420             blnJustLooking = (Not gblnGoToReportMsg)
430           End If

440           gdatStartDate = 0: gdatEndDate = 0: gstrCrtRpt_Account = vbNullString
450           gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString
460           gstrCrtRpt_NetIncome = vbNullString: gstrCrtRpt_NetLoss = vbNullString
470           gstrCrtRpt_CashAssets_Beg = vbNullString: gstrCrtRpt_CashAssets_End = vbNullString
480           gblnCrtRpt_Zero = False: gblnCrtRpt_ZeroDialog = False
490           blnFormOpening = False: strLastFocus = vbNullString
500           DoEvents

              ' ** A little fun with numbers and arrays!
510           lngCmds = 0&
520           ReDim arr_varCmd(C_ELEMS, 0)
              ' ******************************************
              ' ** Array: arr_varCmd()
              ' **
              ' **   Element  Name            Constant
              ' **   =======  ==============  ==========
              ' **      0     Control Name    C_NAM
              ' **      1     Left            C_LFT
              ' **      2     Top             C_TOP
              ' **      3     Column          C_COL
              ' **      4     Order           C_ORD
              ' **
              ' ******************************************
530           For Each ctl In .Controls
540             With ctl
550               If .Name <> "cmdClose" And .Name <> "cmdSearch" Then
560                 Select Case .ControlType
                    Case acCommandButton
570                   lngCmds = lngCmds + 1&
580                   lngE = lngCmds - 1&
590                   ReDim Preserve arr_varCmd(C_ELEMS, lngE)
600                   arr_varCmd(C_NAM, lngE) = .Name
610                   arr_varCmd(C_LFT, lngE) = .Left
620                   arr_varCmd(C_TOP, lngE) = .Top
630                   arr_varCmd(C_COL, lngE) = 0&
640                   arr_varCmd(C_ORD, lngE) = 0&
650                 End Select
660               End If
670             End With
680           Next
690           DoEvents

              ' ** Get top-most, left-most, right-most, and bottom-most command buttons.
700           lngLeftCmd_Left = 1000000: lngRightCmd_Left = 0&
710           lngTopCmd_Top = 1000000: lngBotCmd_Top = 0&
720           For lngX = 0& To (lngCmds - 1&)
730             If arr_varCmd(C_LFT, lngX) < lngLeftCmd_Left Then
740               lngLeftCmd_Left = arr_varCmd(C_LFT, lngX)
750             End If
760             If arr_varCmd(C_LFT, lngX) > lngRightCmd_Left Then
770               lngRightCmd_Left = arr_varCmd(C_LFT, lngX)
780             End If
790             If arr_varCmd(C_TOP, lngX) < lngTopCmd_Top Then
800               lngTopCmd_Top = arr_varCmd(C_TOP, lngX)
810             End If
820             If arr_varCmd(C_TOP, lngX) > lngBotCmd_Top Then
830               lngBotCmd_Top = arr_varCmd(C_TOP, lngX)
840             End If
850           Next

              ' ** Add their column numbers.
860           For lngX = 0& To (lngCmds - 1&)
870             If arr_varCmd(C_LFT, lngX) = lngLeftCmd_Left Then
880               arr_varCmd(C_COL, lngX) = CLng(1)
890               If arr_varCmd(C_TOP, lngX) = lngTopCmd_Top Then
900                 arr_varCmd(C_ORD, lngX) = CLng(1)
910               End If
920             ElseIf arr_varCmd(C_LFT, lngX) = lngRightCmd_Left Then
930               arr_varCmd(C_COL, lngX) = CLng(2)
940               If arr_varCmd(C_TOP, lngX) = lngBotCmd_Top Then
                    'arr_varCmd(C_ORD, lngX) = lngCmds
950               End If
960             End If
970           Next
980           DoEvents

              ' ** Add the button order for column 1.
990           lngTmp02 = 3&  ' ** Skip #2 in column 1.
1000          lngTmp01 = 1000000
1010          Do While lngTmp01 > 0&

1020            blnTmp03 = False
1030            For lngX = 0& To (lngCmds - 1&)
1040              If arr_varCmd(C_COL, lngX) = 1& And arr_varCmd(C_ORD, lngX) = 0& And arr_varCmd(C_TOP, lngX) < lngTmp01 Then
1050                blnTmp03 = True
1060                lngTmp01 = arr_varCmd(C_TOP, lngX)
1070              End If
1080            Next
                ' ** This should go through once, and come out with the lowest un-ordered top in lngTmp01.

1090            For lngX = 0& To (lngCmds - 1&)
1100              If arr_varCmd(C_COL, lngX) = 1& And arr_varCmd(C_TOP, lngX) = lngTmp01 Then
1110                arr_varCmd(C_ORD, lngX) = lngTmp02
1120                lngTmp02 = lngTmp02 + 1&
1130                Exit For
1140              End If
1150            Next
                ' ** This should go through once, and put the next order (lngTmp02) into the one whose Top matches lngTmp01

1160            blnTmp03 = False
1170            For lngX = 0& To (lngCmds - 1&)
1180              If arr_varCmd(C_COL, lngX) = 1& And arr_varCmd(C_ORD, lngX) = 0& Then
1190                blnTmp03 = True
1200                Exit For
1210              End If
1220            Next
                ' ** This should then check for still-missing orders, and set blnTmp03 to True if it finds one.

1230            If blnTmp03 = True Then
                  ' ** If it found one, the loop will continue.
1240              lngTmp01 = 1000000
1250            Else
                  ' ** If all orders are filled, it should exit the loop.
1260              lngTmp01 = 0&
1270              Exit Do
1280            End If
1290            DoEvents

1300          Loop

              ' ** Add the button order for column 2 (though one will be missing because of the double-height button).
1310          For lngTmp01 = 1& To 2&
1320            For lngX = 0& To (lngCmds - 1&)
1330              If arr_varCmd(C_COL, lngX) = 1& Then
1340                For lngY = 0& To (lngCmds - 1&)
1350                  If arr_varCmd(C_COL, lngY) = 2& And arr_varCmd(C_TOP, lngY) = arr_varCmd(C_TOP, lngX) Then
1360                    arr_varCmd(C_ORD, lngY) = arr_varCmd(C_ORD, lngX)
1370                    Exit For
1380                  End If
1390                Next
1400              End If
1410            Next
1420          Next
1430          DoEvents

1440          lngMaxSortOrd = 0&
1450          For lngX = 0& To (lngCmds - 1&)
1460            If arr_varCmd(C_COL, lngX) = 2& And arr_varCmd(C_ORD, lngX) = 0& Then
1470              arr_varCmd(C_ORD, lngX) = 2&
1480            End If
1490            If arr_varCmd(C_ORD, lngX) > lngMaxSortOrd Then lngMaxSortOrd = arr_varCmd(C_ORD, lngX)
1500          Next

1510        Else
1520          Cancel = -1
1530        End If
1540      Else
1550        Cancel = -1
1560      End If
1570    End With

1580    If Cancel = -1 Then
1590      Beep
1600      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1610    End If

        ' ** Column 1:
        ' **   1  cmdStatements
        ' **   3  cmdCashControl
        ' **   4  cmdAccountBalance
        ' **   5  cmdSecurityMaturity
        ' **   6  cmdTaxLot
        ' **   7  cmdCapitalGainsAndLosses
        ' **   8  cmdUnrealizedGainAndLoss
        ' **   9  cmdTaxIncDed
        ' ** Column 2:
        ' **   1  cmdRevIncExp
        ' **   2  cmdIncomeStatement
        ' **   3  cmdCourtReports
        ' **   4  cmdTransByType
        ' **   5  cmdAccountReviews
        ' **   6  cmdAcctNewClosed
        ' **   7  cmdAssetPurchasedSold
        ' **   8  cmdStatementOfCondition
        ' **   9  cmdAccountProfile

EXITP:
1620    Set ctl = Nothing
1630    Exit Sub

ERRH:
1640    DoCmd.Hourglass False
1650    Select Case ERR.Number
        Case Else
1660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1670    End Select
1680    Resume EXITP

End Sub

Private Sub Form_Load()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngOpenForms As Long
        Dim lngX As Long

1710    If Right(gstrTrustDataLocation, 1) <> LNK_SEP Then gstrTrustDataLocation = gstrTrustDataLocation & LNK_SEP

1720    Set dbs = CurrentDb
1730    With dbs
          ' ** Delete any zero-value Ledger records! They shouldn't have been there to begin with!
1740      Set qdf = .QueryDefs("qryStatementParameters_11")
1750      qdf.Execute
1760      Set qdf = Nothing
1770      DoEvents
          ' ** Make sure first balance dates precede first transaction.
          ' ** Account, linked to qryStatementParameters_23a (Balance, grouped by accountno,
          ' ** with Min(balance_date)), qryStatementParameters_23e (qryStatementParameters_23d
          ' ** (Union of qryStatementParameters_23b (Ledger, grouped by accountno, with
          ' ** Min(transdate)), qryStatementParameters_23c (LedgerArchive, grouped by
          ' ** accountno, Min(transdate))), grouped by accountno, with Min(transdate)),
          ' ** just transdate < balance_date, with balance_date_new.
1780      Set qdf = .QueryDefs("qryStatementParameters_23f")
1790      Set rst = qdf.OpenRecordset
1800      If rst.BOF = True And rst.EOF = True Then
            ' ** No problems.
1810        rst.Close
1820        Set rst = Nothing
1830        Set qdf = Nothing
1840      Else
1850        rst.Close
1860        Set rst = Nothing
1870        Set qdf = Nothing
1880        DoEvents
            ' ** Empty tmpBalance.
1890        Set qdf = .QueryDefs("qryStatementParameters_23g")
1900        qdf.Execute
1910        Set qdf = Nothing
1920        DoEvents
            ' ** Append qryStatementParameters_23f (Account, linked to qryStatementParameters_23a
            ' ** (Balance, grouped by accountno, with Min(balance_date)), qryStatementParameters_23e
            ' ** (qryStatementParameters_23d (Union of qryStatementParameters_23b (Ledger, grouped
            ' ** by accountno, with Min(transdate)), qryStatementParameters_23c (LedgerArchive,
            ' ** grouped by accountno, Min(transdate))), grouped by accountno, with Min(transdate)),
            ' ** just transdate < balance_date, with balance_date_new) to tmpBalance.
1930        Set qdf = .QueryDefs("qryStatementParameters_23h")
1940        qdf.Execute
1950        Set qdf = Nothing
1960        DoEvents
            ' ** Update qryStatementParameters_23i (Balance, linked to tmpBalance).
1970        Set qdf = .QueryDefs("qryStatementParameters_23j")
1980        qdf.Execute
1990        Set qdf = Nothing
2000      End If
2010      .Close
2020    End With
2030    Set dbs = Nothing
2040    DoEvents

2050    lngOpenForms = Forms.Count
2060    If lngOpenForms = 1& And Forms(0).Name = THIS_NAME Then
          ' ** Obviously no problem.
2070    Else
2080      For lngX = (lngOpenForms - 1&) To 0& Step -1&
2090        If Left(Forms(lngX).Name, 7) = "frmRpt_" Or Forms(lngX).Name = "frmMenu_CourtReport" Or _
                Forms(lngX).Name = "frmStatementParameters" Then
2100  On Error Resume Next
2110          DoCmd.Close acForm, Forms(lngX).Name
2120  On Error GoTo ERRH
2130        End If
2140      Next
2150    End If

2160    DoCmd.Hourglass False

EXITP:
2170    Set rst = Nothing
2180    Set qdf = Nothing
2190    Set dbs = Nothing
2200    Exit Sub

ERRH:
2210    DoCmd.Hourglass False
2220    Select Case ERR.Number
        Case Else
2230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2240    End Select
2250    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer, Cancel As Integer

2310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2320    intShiftDown = (Shift And acShiftMask) > 0
2330    intAltDown = (Shift And acAltMask) > 0
2340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Load Timer:       D {chkLoadTimer}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Timer Chk Vis:    T {Form_KeyDown}
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Plain keys.
2350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2360      Select Case intRetVal
          Case vbKeyEscape
2370        intRetVal = 0
2380        cmdClose_Click  ' ** Procedure: Below.
2390      End Select
2400    End If

        ' ** Alt keys.
2410    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2420      Select Case intRetVal
          Case vbKeyD
2430        With Me
2440          intRetVal = 0
2450          If .chkLoadTimer.Visible = True Then
2460            If .chkLoadTimer.Enabled = True Then
2470              .chkLoadTimer.SetFocus
2480            End If
2490            .chkLoadTimer = (Not .chkLoadTimer)
2500            chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
2510          End If
2520        End With
2530      Case vbKeyX
2540        intRetVal = 0
2550        cmdClose_Click  ' ** Procedure: Below.
2560      End Select
2570    End If

        ' ** Ctrl-Shift keys.
2580    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2590      Select Case intRetVal
          Case vbKeyF
2600        With Me
2610          intRetVal = 0
2620          .FocusHolder.SetFocus
2630        End With
2640      Case vbKeyT
2650        With Me
2660          intRetVal = 0
2670          If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
2680            FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
2690          ElseIf CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
2700            Select Case .chkLoadTimer.Visible
                Case True
2710              .FocusHolder.SetFocus
2720              DoEvents
2730              .chkLoadTimer.Visible = False
2740              .chkLoadTimer_lbl1.Visible = False
2750              .chkLoadTimer_lbl2.Visible = False
2760              gblnTimer = False
2770            Case False
2780              .chkLoadTimer.Visible = True
2790              .chkLoadTimer_lbl1.Visible = True
2800              .chkLoadTimer_lbl2.Visible = True
2810              chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
2820            End Select
2830          End If
2840          .chkLoadTimer_vis = .chkLoadTimer.Visible

2850        End With
2860      Case vbKeyX
2870        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
2880          intRetVal = 0
2890          blnJustClose = True
2900          SetOption_Dev  ' ** Module Function: modUtilities.
2910          gblnSignal = True  ' ** Borrowing this variable.
2920          DoCmd.Close acForm, THIS_NAME
2930          DoCmd.Hourglass False
2940        End If
2950      End Select
2960    End If

EXITP:
2970    KeyCode = intRetVal
2980    Exit Sub

ERRH:
2990    intRetVal = 0
3000    Select Case ERR.Number
        Case Else
3010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3020    End Select
3030    Resume EXITP

End Sub

Private Sub Form_Timer()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim strFormName As String

        Static lngTmp01 As Long, lngTmp02 As Long

3110    With Me
3120      .TimerInterval = 0&
3130      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
3140        lngGTR_Stat = lngGTR_Stat + 1&
3150        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
3160        lngTmp01 = lngTmp01 + 1&
3170        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
3180        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
3190        If lngTmp01 > 1& Then
3200          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
3210        End If
3220        If lngTmp01 = 24& Then
3230          lngTmp01 = 0&
3240          lngTmp02 = lngTmp02 + 1&
3250          If lngTmp02 < 5& Then
3260            .TimerInterval = 50&
3270          Else
3280            blnGTR_Emblem = False
3290          End If
3300        Else
3310          .TimerInterval = 50&
3320        End If
3330      End If  ' ** blnGTR_Emblem.
3340      If gblnSetFocus = True Then
3350        gblnSetFocus = False
3360        strFormName = vbNullString
3370  On Error Resume Next
3380        Do Until Screen.ActiveForm.Name = strFormName
3390          DoCmd.SelectObject acForm, THIS_NAME, False
3400          strFormName = Screen.ActiveForm.Name
3410        Loop
3420  On Error GoTo ERRH
3430        DoEvents
3440      End If
3450      If blnGoingToReport = False Then
3460        Select Case gstrReturningForm
            Case "frmStatementParameters"
3470          .cmdStatements.SetFocus
3480        Case "frmRpt_CashControl"
3490          .cmdCashControl.SetFocus
3500        Case "frmRpt_AccountBalance"
3510          .cmdAccountBalance.SetFocus
3520        Case "frmRpt_Maturity"
3530          .cmdSecurityMaturity.SetFocus
3540        Case "frmRpt_TaxLot"
3550          .cmdTaxLot.SetFocus
3560        Case "frmRpt_CapitalGainAndLoss"
3570          .cmdCapitalGainsAndLosses.SetFocus
3580        Case "frmRpt_UnrealizedGainAndLoss"
3590          .cmdUnrealizedGainAndLoss.SetFocus
3600        Case "frmRpt_TaxIncomeDeductions"
3610          .cmdTaxIncDed.SetFocus
3620        Case "frmRpt_IncomeExpense"
3630          .cmdRevIncExp.SetFocus
3640        Case "frmRpt_IncomeStatement"
3650          .cmdIncomeStatement.SetFocus
3660        Case "frmMenu_CourtReport"
3670          .cmdCourtReports.SetFocus
3680        Case "frmRpt_TransactionsByType"
3690          .cmdTransByType.SetFocus
3700        Case "frmRpt_AccountReviews"
3710          .cmdAccountReviews.SetFocus
3720        Case "frmRpt_NewClosedAccounts"
3730          .cmdAcctNewClosed.SetFocus
3740        Case "frmRpt_PurchasedSold"
3750          .cmdAssetPurchasedSold.SetFocus
3760        Case "frmRpt_StatementOfCondition"
3770          .cmdStatementOfCondition.SetFocus
3780        Case "frmRpt_AccountProfile"
3790          .cmdAccountProfile.SetFocus
3800        Case "frmReportList"
3810          .cmdSearch.SetFocus
3820        Case Else
3830          .cmdStatements.SetFocus
3840        End Select
3850        gstrReturningForm = vbNullString
3860        DoCmd.Hourglass False
3870      End If
3880      If gblnGoToReport = True Then
3890        DoCmd.Hourglass True  ' ** Make sure it's still running.
3900        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
3910        Else
3920          If blnGTR_Emblem = False Then
3930            blnGTR_Emblem = True
3940            lngTmp01 = 0&: lngTmp02 = 0&
3950          End If
3960          DoEvents
3970          Select Case blnGoingToReport
              Case True
                ' ******************************************************
                ' ** Array: garr_varGoToReport()
                ' **
                ' **   Field  Element  Name                Constant
                ' **   =====  =======  ==================  ===========
                ' **     1       0     rptlistrpt_id       GTR_RLRID
                ' **     2       1     rpt_id              GTR_RID
                ' **     3       2     rpt_name            GTR_RNAM
                ' **     4       3     rpt_caption         GTR_RCAP
                ' **     5       4     rptlistrpt_order    GTR_ORD
                ' **     6       5     frm_name1           GTR_FRM1
                ' **     7       6     ctl_name1           GTR_CTL1
                ' **     8       7     ctl_invoke1         GTR_INV1
                ' **     9       8     ctl_focus1          GTR_FOC1
                ' **    10       9     frm_name2           GTR_FRM2
                ' **    11      10     ctl_name2           GTR_CTL2
                ' **    12      11     ctl_invoke2         GTR_INV2
                ' **    13      12     ctl_focus2          GTR_FOC2
                ' **    14      13     frm_name3           GTR_FRM3
                ' **    15      14     ctl_name3           GTR_CTL3
                ' **    16      15     ctl_invoke3         GTR_INV3
                ' **    17      16     ctl_focus3          GTR_FOC3
                ' **    18      17     frm_name4           GTR_FRM4
                ' **    19      18     ctl_name4           GTR_CTL4
                ' **    20      19     ctl_invoke4         GTR_INV4
                ' **    21      20     ctl_focus4          GTR_FOC4
                ' **    22      21     frm_name5           GTR_FRM5
                ' **    23      22     ctl_name5           GTR_CTL5
                ' **    24      23     ctl_invoke5         GTR_INV5
                ' **    25      24     ctl_focus5          GTR_FOC5
                ' **    26      25     frm_name6           GTR_FRM6
                ' **    27      26     ctl_name6           GTR_CTL6
                ' **    28      27     ctl_invoke6         GTR_INV6
                ' **    29      28     ctl_focus6          GTR_FOC6
                ' **
                ' ******************************************************
3980            If garr_varGoToReport(GTR_FRM1) = THIS_NAME Then
3990              Select Case garr_varGoToReport(GTR_INV1)
                  Case True
4000                Select Case garr_varGoToReport(GTR_CTL1)
                    Case "cmdStatements"
4010                  .TimerInterval = 50&
4020                  cmdStatements_Click  ' ** Procedure: Below.
4030                Case "cmdCashControl"
4040                  .TimerInterval = 50&
4050                  cmdCashControl_Click  ' ** Procedure: Below.
4060                Case "cmdAccountBalance"
4070                  .TimerInterval = 50&
4080                  cmdAccountBalance_Click  ' ** Procedure: Below.
4090                Case "cmdSecurityMaturity"
4100                  .TimerInterval = 50&
4110                  cmdSecurityMaturity_Click  ' ** Procedure: Below.
4120                Case "cmdTaxLot"
4130                  .TimerInterval = 50&
4140                  cmdTaxLot_Click  ' ** Procedure: Below.
4150                Case "cmdCapitalGainsAndLosses"
4160                  .TimerInterval = 50&
4170                  cmdCapitalGainsAndLosses_Click  ' ** Procedure: Below.
4180                Case "cmdUnrealizedGainAndLoss"
4190                  .TimerInterval = 50&
4200                  cmdUnrealizedGainAndLoss_Click  ' ** Procedure: Below.
4210                Case "cmdTaxIncDed"
4220                  .TimerInterval = 50&
4230                  cmdTaxIncDed_Click  ' ** Procedure: Below.
4240                Case "cmdRevIncExp"
4250                  .TimerInterval = 50&
4260                  cmdRevIncExp_Click  ' ** Procedure: Below.
4270                Case "cmdIncomeStatement"
4280                  .TimerInterval = 50&
4290                  cmdIncomeStatement_Click  ' ** Procedure: Below.
4300                Case "cmdCourtReports"
4310                  .TimerInterval = 50&
4320                  cmdCourtReports_Click  ' ** Procedure: Below.
4330                Case "cmdTransByType"
4340                  .TimerInterval = 50&
4350                  cmdTransByType_Click  ' ** Procedure: Below.
4360                Case "cmdAccountReviews"
4370                  .TimerInterval = 50&
4380                  cmdAccountReviews_Click  ' ** Procedure: Below.
4390                Case "cmdAcctNewClosed"
4400                  .TimerInterval = 50&
4410                  cmdAcctNewClosed_Click  ' ** Procedure: Below.
4420                Case "cmdAssetPurchasedSold"
4430                  .TimerInterval = 50&
4440                  cmdAssetPurchasedSold_Click  ' ** Procedure: Below.
4450                Case "cmdStatementOfCondition"
4460                  .TimerInterval = 50&
4470                  cmdStatementOfCondition_Click  ' ** Procedure: Below.
4480                Case "cmdAccountProfile"
4490                  .TimerInterval = 50&
4500                  cmdAccountProfile_Click  ' ** Procedure: Below.
4510                Case Else
                      ' ** Shouldn't be here.
4520                  blnGoingToReport = False
4530                  DoCmd.Hourglass False
4540                  gblnGoToReport = False
4550                  .TimerInterval = 50&
4560                  Beep
4570                End Select
4580              Case False
                    ' ** Just sit there, and we're done.
4590                blnGoingToReport = False
4600                DoCmd.Hourglass False
4610                gblnGoToReport = False
4620                Beep
4630                .TimerInterval = 50&
4640              End Select
4650            Else
                  ' ** Continuing back to frmMenu_Main.
4660              blnGoingToReport = False
4670              .TimerInterval = 50&
4680              cmdClose_Click  ' ** Procedure: Below.
4690            End If
4700          Case False
                ' ******************************************************
                ' ** Array: garr_varGoToReport()
                ' **
                ' **   Field  Element  Name                Constant
                ' **   =====  =======  ==================  ===========
                ' **     1       0     rptlistrpt_id       GTR_RLRID
                ' **     2       1     rpt_id              GTR_RID
                ' **     3       2     rpt_name            GTR_RNAM
                ' **     4       3     rpt_caption         GTR_RCAP
                ' **     5       4     rptlistrpt_order    GTR_ORD
                ' **     6       5     frm_name1           GTR_FRM1
                ' **     7       6     ctl_name1           GTR_CTL1
                ' **     8       7     ctl_invoke1         GTR_INV1
                ' **     9       8     ctl_focus1          GTR_FOC1
                ' **    10       9     frm_name2           GTR_FRM2
                ' **    11      10     ctl_name2           GTR_CTL2
                ' **    12      11     ctl_invoke2         GTR_INV2
                ' **    13      12     ctl_focus2          GTR_FOC2
                ' **    14      13     frm_name3           GTR_FRM3
                ' **    15      14     ctl_name3           GTR_CTL3
                ' **    16      15     ctl_invoke3         GTR_INV3
                ' **    17      16     ctl_focus3          GTR_FOC3
                ' **    18      17     frm_name4           GTR_FRM4
                ' **    19      18     ctl_name4           GTR_CTL4
                ' **    20      19     ctl_invoke4         GTR_INV4
                ' **    21      20     ctl_focus4          GTR_FOC4
                ' **    22      21     frm_name5           GTR_FRM5
                ' **    23      22     ctl_name5           GTR_CTL5
                ' **    24      23     ctl_invoke5         GTR_INV5
                ' **    25      24     ctl_focus5          GTR_FOC5
                ' **    26      25     frm_name6           GTR_FRM6
                ' **    27      26     ctl_name6           GTR_CTL6
                ' **    28      27     ctl_invoke6         GTR_INV6
                ' **    29      28     ctl_focus6          GTR_FOC6
                ' **
                ' ******************************************************
4710            If garr_varGoToReport(GTR_FRM1) = THIS_NAME Then
                  ' ** This is first menu in path.
4720              blnGoingToReport = True
4730              Select Case IsNull(garr_varGoToReport(GTR_FRM2))
                  Case True
                    ' ** No more menus.
4740                .Controls(garr_varGoToReport(GTR_CTL1)).SetFocus
4750                lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL1)).Top
4760                lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL1)).Left
4770                Select Case lngBtn_Left
                    Case .cmdStatements.Left
4780                  .GoToReport_arw_menu_l_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
4790                  .GoToReport_arw_menu_l_img.Visible = True
4800                Case .cmdRevIncExp.Left
4810                  .GoToReport_arw_menu_r_img.Top = lngBtn_Top + (3& * lngTpp)
4820                  .GoToReport_arw_menu_r_img.Visible = True
4830                End Select
4840                DoEvents
4850                .TimerInterval = 50&
4860              Case False
                    ' ** Moving on, must be Court Reports.
4870                .Controls(garr_varGoToReport(GTR_CTL1)).SetFocus
4880                lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL1)).Top
4890                lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL1)).Left
4900                Select Case lngBtn_Left
                    Case .cmdStatements.Left
4910                  .GoToReport_arw_menu_l_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
4920                  .GoToReport_arw_menu_l_img.Visible = True
4930                Case .cmdRevIncExp.Left
4940                  .GoToReport_arw_menu_r_img.Top = lngBtn_Top + (3& * lngTpp)
4950                  .GoToReport_arw_menu_r_img.Visible = True
4960                End Select
4970                Select Case garr_varGoToReport(GTR_INV1)
                    Case True
4980                  Select Case garr_varGoToReport(GTR_CTL1)
                      Case "cmdCourtReports"
4990                    .cmdCourtReports_box01.Visible = True
5000                    .cmdCourtReports_box02.Visible = True
5010                    DoEvents
5020                    .TimerInterval = 50&
5030                  Case "cmdStatements"
5040                    .cmdCourtReports_box01.Visible = True
5050                    .cmdCourtReports_box02.Visible = True
5060                    DoEvents
5070                    .TimerInterval = 50&
5080                  Case Else
                        ' ** Just sit there, and we're done.
5090                    blnGoingToReport = False
5100                    DoCmd.Hourglass False
5110                    gblnGoToReport = False
5120                    blnGTR_Emblem = False
5130                    GTREmblem_Off  ' ** Procedure: Below.
5140                    Beep
5150                  End Select
5160                Case False
                      ' ** Just sit there, and we're done.
5170                  blnGoingToReport = False
5180                  DoCmd.Hourglass False
5190                  gblnGoToReport = False
5200                  blnGTR_Emblem = False
5210                  GTREmblem_Off  ' ** Procedure: Below.
5220                  Beep
5230                End Select
5240              End Select
5250            Else
                  ' ** Head back to frmMenu_Main to start the path.
5260              blnGoingToReport = True
5270              .cmdClose.SetFocus
5280              .GoToReport_arw_close_img.Visible = True
5290              .cmdClose.Caption = "< Ba&ck"
5300              DoEvents
5310              .TimerInterval = 50&
5320            End If
5330          End Select
5340        End If
5350      End If
5360      If gblnGoToReport = False Then
5370        If blnJustLooking = False Then                              '2nd time here first.
5380  On Error Resume Next
5390          .cmdSearch_lbl1.Visible = False
5400          .cmdSearch_lbl2.Visible = False
5410          .cmdSearch_arw.Visible = False
5420          If blnTimerVis = True Then
5430            .chkLoadTimer.Visible = True
5440            .chkLoadTimer_lbl1.Visible = True
5450            .chkLoadTimer_lbl2.Visible = True
5460          End If
5470  On Error GoTo ERRH
5480          If gblnTimer = True Then
5490            If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
5500              Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
5510              gblnTimer = False
5520              DoCmd.Hourglass False
5530            End If
5540          End If
5550        ElseIf gblnGoToReportMsg = False Then                       '1st time here first, then once more around to above.
              ' ** Message will show first time each session.
5560          blnJustLooking = False
5570          If blnTimerVis = True Then
5580            .chkLoadTimer.Visible = False
5590            .chkLoadTimer_lbl1.Visible = False
5600            .chkLoadTimer_lbl2.Visible = False
5610          End If
5620  On Error Resume Next
5630          .cmdSearch_lbl1.Visible = True
5640          .cmdSearch_lbl2.Visible = True
5650          .cmdSearch_arw.Visible = True
5660  On Error GoTo ERRH
5670          gblnGoToReportMsg = True
5680          If gblnTimer = True Then
5690            If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
5700              Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
5710              gblnTimer = False
5720              DoCmd.Hourglass False
5730            End If
5740          End If
5750  On Error Resume Next
5760          .TimerInterval = 5000&  ' ** 5 sec.
5770  On Error GoTo ERRH
5780        End If
5790      End If
5800    End With

        'HERE! 1  Form_Open()  frmMenu_Report
        'HERE! 2  Form_Load()  frmMenu_Report
        'HERE! 3  Form_Timer()  frmMenu_Report
        'HERE! 3  Form_Timer()  frmMenu_Report

EXITP:
5810    Exit Sub

ERRH:
5820    DoCmd.Hourglass False
5830    blnGoingToReport = False
5840    gblnGoToReport = False
5850    blnJustLooking = False
5860    Select Case ERR.Number
        Case Else
5870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5880    End Select
5890    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

5910    With Me

5920      If gblnSignal = False Then  ' ** Borrowing this variable.
5930        DoCmd.Hourglass True
5940        DoEvents
5950      End If

          ' ** Make sure all reports are closed.
5960      If Reports.Count > 0 Then
5970        Do While Reports.Count > 0
5980          DoCmd.Close acReport, Reports(0).Name
5990        Loop
6000      End If

6010      If blnJustClose = False Then
6020        If gblnGoToReport = False Then
6030          blnJustLooking = False
6040          .TimerInterval = 0&
6050        End If
6060      End If

6070      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

6080      blnGoingToReport = False

6090      gdatStartDate = 0: gdatEndDate = 0: gstrCrtRpt_Account = vbNullString
6100      gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString
6110      gstrCrtRpt_NetIncome = vbNullString: gstrCrtRpt_NetLoss = vbNullString
6120      gstrCrtRpt_CashAssets_Beg = vbNullString: gstrCrtRpt_CashAssets_End = vbNullString
6130      gblnCrtRpt_Zero = False: gblnCrtRpt_ZeroDialog = False

6140      If blnJustClose = False Then

6150        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
6160        DoEvents

6170        gstrFormQuerySpec = vbNullString
6180        gblnSetFocus = True
6190        gstrReturningForm = THIS_NAME
6200        Select Case strCallingForm
            Case "frmMenu_Main"
6210          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Title"
6220          If gblnGoToReport = True Then
6230            Forms(strCallingForm).TimerInterval = 100&
6240          End If
6250        End Select

6260      End If  ' ** blnJustClose.

6270    End With

EXITP:
6280    Set qdf = Nothing
6290    Set dbs = Nothing
6300    Exit Sub

ERRH:
6310    DoCmd.Hourglass False
6320    Select Case ERR.Number
        Case Else
6330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6340    End Select
6350    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

6410    With Me
6420      If strLastFocus <> "cmdCourtReports" Then
6430        .cmdCourtReports_box01.Visible = False
6440        .cmdCourtReports_box02.Visible = False
6450      End If
6460      If .cmdClose.Caption <> "&Close" Then
6470        .cmdClose.Caption = "&Close"
6480      End If
6490    End With

EXITP:
6500    Exit Sub

ERRH:
6510    Select Case ERR.Number
        Case Else
6520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6530    End Select
6540    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

6610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6620    Exit Sub

ERRH:
6630    Select Case ERR.Number
        Case Else
6640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6650    End Select
6660    Resume EXITP

End Sub

Private Sub Header_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_box_MouseMove"

6710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6720    Exit Sub

ERRH:
6730    Select Case ERR.Number
        Case Else
6740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6750    End Select
6760    Resume EXITP

End Sub

Private Sub Header_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_MouseMove"

6810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6820    Exit Sub

ERRH:
6830    Select Case ERR.Number
        Case Else
6840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6850    End Select
6860    Resume EXITP

End Sub

Private Sub Header_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl2_MouseMove"

6910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6920    Exit Sub

ERRH:
6930    Select Case ERR.Number
        Case Else
6940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6950    End Select
6960    Resume EXITP

End Sub

Private Sub cmdClose_Click()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

7010    With Me
7020      If blnFormOpening = False Then

7030        DoCmd.Hourglass True
7040        DoEvents

7050        .TimerInterval = 0&

7060        blnJustClose = False
7070        DoCmd.Close acForm, THIS_NAME

7080      End If
7090    End With

EXITP:
7100    Exit Sub

ERRH:
7110    blnFormOpening = False
7120    DoCmd.Hourglass False
7130    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
7140    Case Else
7150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7160    End Select
7170    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7220    intShiftDown = (Shift And acShiftMask) > 0
7230    intAltDown = (Shift And acAltMask) > 0
7240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7260      Select Case intRetVal
          Case vbKeyTab
7270        With Me
7280          intRetVal = 0
7290          .cmdStatements.SetFocus
7300        End With
7310      End Select
7320    End If

        ' ** Shift keys.
7330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7340      Select Case intRetVal
          Case vbKeyTab
7350        With Me
7360          intRetVal = 0
7370          If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
7380            .chkLoadTimer.SetFocus
7390          Else
7400            .cmdSearch.SetFocus
7410          End If
7420        End With
7430      End Select
7440    End If

EXITP:
7450    KeyCode = intRetVal
7460    Exit Sub

ERRH:
7470    intRetVal = 0
7480    Select Case ERR.Number
        Case Else
7490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7500    End Select
7510    Resume EXITP

End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_MouseMove"

7610    With Me
7620      .cmdClose.Caption = "< Ba&ck"
7630    End With

EXITP:
7640    Exit Sub

ERRH:
7650    Select Case ERR.Number
        Case Else
7660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7670    End Select
7680    Resume EXITP

End Sub

Private Sub cmdReport_l_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReport_l_box01_MouseMove"

7710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7720    Exit Sub

ERRH:
7730    Select Case ERR.Number
        Case Else
7740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7750    End Select
7760    Resume EXITP

End Sub

Private Sub cmdStatements_Click()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStatements_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

7810    With Me
7820      If blnFormOpening = False Then

7830        strDocName = "frmStatementParameters"

7840        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
7850        DoEvents

7860        If .chkLoadTimer.Visible = True Then
7870          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
7880          If gblnTimer = True Then
                'gblnSetFocus = True
7890            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
7900          End If
7910        End If
7920        DoEvents

7930        DoCmd.Hourglass True
7940        DoEvents

7950        .TimerInterval = 0&

7960        gblnSetFocus = True
7970        blnJustClose = True
7980        gstrReturningForm = strDocName
7990        DoCmd.Close acForm, THIS_NAME
8000        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
8010        If gblnGoToReport = True Then
8020          Forms(strDocName).TimerInterval = 50&
8030          blnGoingToReport = False
8040          blnGTR_Emblem = False
8050        End If

8060        Do Until Screen.ActiveForm.Name = strDocName
8070          DoCmd.SelectObject acForm, strDocName, False
8080          strDocName = Screen.ActiveForm.Name
8090        Loop

8100        DoCmd.Hourglass False

8110      End If
8120    End With

EXITP:
8130    Set rst = Nothing
8140    Set qdf = Nothing
8150    Set dbs = Nothing
8160    Exit Sub

ERRH:
8170    blnFormOpening = False
8180    DoCmd.Hourglass False
8190    Select Case ERR.Number
        Case Else
8200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8210    End Select
8220    Resume EXITP

End Sub

Private Sub cmdStatements_KeyDown(KeyCode As Integer, Shift As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStatements_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8320    intShiftDown = (Shift And acShiftMask) > 0
8330    intAltDown = (Shift And acAltMask) > 0
8340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8360      Select Case intRetVal
          Case vbKeyTab
8370        With Me
8380          intRetVal = 0
8390          .cmdCashControl.SetFocus
8400        End With
8410      Case Else
8420        intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.
8430      End Select
8440    End If

        ' ** Shift keys.
8450    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8460      Select Case intRetVal
          Case vbKeyTab
8470        With Me
8480          intRetVal = 0
8490          .cmdClose.SetFocus
8500        End With
8510      End Select
8520    End If

EXITP:
8530    KeyCode = intRetVal
8540    Exit Sub

ERRH:
8550    intRetVal = 0
8560    Select Case ERR.Number
        Case Else
8570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8580    End Select
8590    Resume EXITP

End Sub

Private Sub cmdCashControl_Click()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCashControl_Click"

8610    With Me
8620      If blnFormOpening = False Then

8630        DoCmd.Hourglass True
8640        DoEvents

8650        .TimerInterval = 0&

8660        gblnSetFocus = True
8670        blnJustClose = True: blnFormOpening = True
8680        strDocName = "frmRpt_CashControl"
8690        DoCmd.Close acForm, THIS_NAME
8700        gstrReturningForm = strDocName
8710        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
8720        If gblnGoToReport = True Then
8730          Forms(strDocName).TimerInterval = 50&
8740          blnGoingToReport = False
8750          blnGTR_Emblem = False
8760        End If
8770        DoEvents

8780      End If
8790    End With

EXITP:
8800    Exit Sub

ERRH:
8810    blnFormOpening = False
8820    DoCmd.Hourglass False
8830    Select Case ERR.Number
        Case Else
8840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8850    End Select
8860    Resume EXITP

End Sub

Private Sub cmdCashControl_KeyDown(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCashControl_KeyDown"

        Dim intRetVal As Integer

8910    intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
8920    KeyCode = intRetVal
8930    Exit Sub

ERRH:
8940    intRetVal = 0
8950    Select Case ERR.Number
        Case Else
8960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8970    End Select
8980    Resume EXITP

End Sub

Private Sub cmdCashControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCashControl_MouseMove"

9010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9020    Exit Sub

ERRH:
9030    Select Case ERR.Number
        Case Else
9040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9050    End Select
9060    Resume EXITP

End Sub

Private Sub cmdAccountBalance_Click()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccountBalance_Click"

9110    With Me
9120      If blnFormOpening = False Then

9130        DoCmd.Hourglass True
9140        DoEvents

9150        .TimerInterval = 0&

9160        gblnSetFocus = True
9170        blnJustClose = True: blnFormOpening = True
9180        strDocName = "frmRpt_AccountBalance"
9190        DoCmd.Close acForm, THIS_NAME
9200        gstrReturningForm = strDocName
9210        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
9220        If gblnGoToReport = True Then
9230          Forms(strDocName).TimerInterval = 50&
9240          blnGoingToReport = False
9250          blnGTR_Emblem = False
9260        End If
9270        DoEvents

9280      End If
9290    End With

EXITP:
9300    Exit Sub

ERRH:
9310    blnFormOpening = False
9320    DoCmd.Hourglass False
9330    Select Case ERR.Number
        Case Else
9340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9350    End Select
9360    Resume EXITP

End Sub

Private Sub cmdAccountBalance_KeyDown(KeyCode As Integer, Shift As Integer)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccountBalance_KeyDown"

        Dim intRetVal As Integer

9410    intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
9420    KeyCode = intRetVal
9430    Exit Sub

ERRH:
9440    intRetVal = 0
9450    Select Case ERR.Number
        Case Else
9460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9470    End Select
9480    Resume EXITP

End Sub

Private Sub cmdSecurityMaturity_Click()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSecurityMaturity_Click"

9510    With Me
9520      If blnFormOpening = False Then

9530        DoCmd.Hourglass True
9540        DoEvents

9550        .TimerInterval = 0&

9560        gblnSetFocus = True
9570        blnJustClose = True: blnFormOpening = True
9580        strDocName = "frmRpt_Maturity"
9590        DoCmd.Close acForm, THIS_NAME
9600        gstrReturningForm = strDocName
9610        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
9620        If gblnGoToReport = True Then
9630          Forms(strDocName).TimerInterval = 50&
9640          blnGoingToReport = False
9650          blnGTR_Emblem = False
9660        End If
9670        DoEvents

9680      End If
9690    End With

EXITP:
9700    Exit Sub

ERRH:
9710    blnFormOpening = False
9720    DoCmd.Hourglass False
9730    Select Case ERR.Number
        Case Else
9740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9750    End Select
9760    Resume EXITP

End Sub

Private Sub cmdSecurityMaturity_KeyDown(KeyCode As Integer, Shift As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSecurityMaturity_KeyDown"

        Dim intRetVal As Integer

9810    intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
9820    KeyCode = intRetVal
9830    Exit Sub

ERRH:
9840    intRetVal = 0
9850    Select Case ERR.Number
        Case Else
9860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9870    End Select
9880    Resume EXITP

End Sub

Private Sub cmdTaxLot_Click()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxLot_Click"

9910    With Me
9920      If blnFormOpening = False Then

9930        DoCmd.Hourglass True
9940        DoEvents

9950        .TimerInterval = 0&

9960        gblnSetFocus = True
9970        blnJustClose = True: blnFormOpening = True
9980        DoCmd.Close acForm, THIS_NAME
9990        strDocName = "frmRpt_TaxLot"
10000       gstrReturningForm = strDocName
10010       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
10020       If gblnGoToReport = True Then
10030         Forms(strDocName).TimerInterval = 50&
10040         blnGoingToReport = False
10050         blnGTR_Emblem = False
10060       End If

10070       DoCmd.Hourglass False

10080     End If
10090   End With

EXITP:
10100   Exit Sub

ERRH:
10110   blnFormOpening = False
10120   DoCmd.Hourglass False
10130   Select Case ERR.Number
        Case Else
10140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10150   End Select
10160   Resume EXITP

End Sub

Private Sub cmdTaxLot_KeyDown(KeyCode As Integer, Shift As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxLot_KeyDown"

        Dim intRetVal As Integer

10210   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
10220   KeyCode = intRetVal
10230   Exit Sub

ERRH:
10240   intRetVal = 0
10250   Select Case ERR.Number
        Case Else
10260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10270   End Select
10280   Resume EXITP

End Sub

Private Sub cmdCapitalGainsAndLosses_Click()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCapitalGainsAndLosses_Click"

10310   With Me
10320     If blnFormOpening = False Then

10330       DoCmd.Hourglass True
10340       DoEvents

10350       .TimerInterval = 0&

10360       gblnSetFocus = True
10370       blnJustClose = True
10380       DoCmd.Close acForm, THIS_NAME
10390       strDocName = "frmRpt_CapitalGainAndLoss"
10400       gstrReturningForm = strDocName
10410       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
10420       If gblnGoToReport = True Then
10430         Forms(strDocName).TimerInterval = 50&
10440         blnGoingToReport = False
10450         blnGTR_Emblem = False
10460       End If

10470     End If
10480   End With

EXITP:
10490   Exit Sub

ERRH:
10500   blnFormOpening = False
10510   DoCmd.Hourglass False
10520   Select Case ERR.Number
        Case Else
10530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10540   End Select
10550   Resume EXITP

End Sub

Private Sub cmdCapitalGainsAndLosses_KeyDown(KeyCode As Integer, Shift As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCapitalGainsAndLosses_KeyDown"

        Dim intRetVal As Integer

10610   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
10620   KeyCode = intRetVal
10630   Exit Sub

ERRH:
10640   intRetVal = 0
10650   Select Case ERR.Number
        Case Else
10660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10670   End Select
10680   Resume EXITP

End Sub

Private Sub cmdUnrealizedGainAndLoss_Click()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUnrealizedGainAndLoss_Click"

10710   With Me
10720     If blnFormOpening = False Then

10730       DoCmd.Hourglass True
10740       DoEvents

10750       .TimerInterval = 0&

10760       gblnSetFocus = True
10770       blnJustClose = True: blnFormOpening = True
10780       DoCmd.Close acForm, THIS_NAME
10790       strDocName = "frmRpt_UnrealizedGainAndLoss"
10800       gstrReturningForm = strDocName
10810       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
10820       If gblnGoToReport = True Then
10830         Forms(strDocName).TimerInterval = 50&
10840         blnGoingToReport = False
10850         blnGTR_Emblem = False
10860       End If

10870       DoCmd.Hourglass False

10880     End If
10890   End With

EXITP:
10900   Exit Sub

ERRH:
10910   blnFormOpening = False
10920   DoCmd.Hourglass False
10930   Select Case ERR.Number
        Case Else
10940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10950   End Select
10960   Resume EXITP

End Sub

Private Sub cmdUnrealizedGainAndLoss_KeyDown(KeyCode As Integer, Shift As Integer)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUnrealizedGainAndLoss_KeyDown"

        Dim intRetVal As Integer

11010   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
11020   KeyCode = intRetVal
11030   Exit Sub

ERRH:
11040   intRetVal = 0
11050   Select Case ERR.Number
        Case Else
11060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11070   End Select
11080   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_Click()
' ** Tax Reports.
' ** (See also Income / Expense Report {cmdRevIncExp}, frmRpt_IncomeExpense.)
' ** These 2 reports are based on the entire Ledger, split between
' ** Income/Expense revcode_TYPE; (excluding ledger_HIDDEN).

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_Click"

11110   With Me
11120     If blnFormOpening = False Then

11130       DoCmd.Hourglass True
11140       DoEvents

11150       .TimerInterval = 0&

11160       gblnSetFocus = True
11170       blnJustClose = True: blnFormOpening = True
11180       DoCmd.Close acForm, THIS_NAME
11190       strDocName = "frmRpt_TaxIncomeDeductions"
11200       gstrReturningForm = strDocName
11210       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
11220       If gblnGoToReport = True Then
11230         Forms(strDocName).TimerInterval = 100&
11240         blnGoingToReport = False
11250         blnGTR_Emblem = False
11260       End If

11270       DoCmd.Hourglass False

11280     End If
11290   End With

EXITP:
11300   Exit Sub

ERRH:
11310   blnFormOpening = False
11320   DoCmd.Hourglass False
11330   Select Case ERR.Number
        Case Else
11340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11350   End Select
11360   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_KeyDown(KeyCode As Integer, Shift As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_KeyDown"

        Dim intRetVal As Integer

11410   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
11420   KeyCode = intRetVal
11430   Exit Sub

ERRH:
11440   intRetVal = 0
11450   Select Case ERR.Number
        Case Else
11460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11470   End Select
11480   Resume EXITP

End Sub

Private Sub cmdReport_r_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReport_r_box01_MouseMove"

11510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11520   Exit Sub

ERRH:
11530   Select Case ERR.Number
        Case Else
11540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11550   End Select
11560   Resume EXITP

End Sub

Private Sub cmdRevIncExp_Click()
' ** Income / Expense Report.
' ** (See also Tax Reports {cmdTaxIncDed}, frmRpt_TaxIncomeDeductions.)
' ** These 2 reports are based on the entire Ledger, with complex filters for each;
' ** (excluding ICash/Cost washouts, zero ICash/PCash, Long Term Capital Gains, and ledger_HIDDEN).

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_Click"

11610   With Me
11620     If blnFormOpening = False Then

11630       DoCmd.Hourglass True
11640       DoEvents

11650       .TimerInterval = 0&

11660       gblnSetFocus = True
11670       blnJustClose = True: blnFormOpening = True
11680       DoCmd.Close acForm, THIS_NAME
11690       strDocName = "frmRpt_IncomeExpense"
11700       gstrReturningForm = strDocName
11710       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
11720       If gblnGoToReport = True Then
11730         Forms(strDocName).TimerInterval = 50&
11740         blnGoingToReport = False
11750         blnGTR_Emblem = False
11760       End If

11770       DoCmd.Hourglass False

11780     End If
11790   End With

EXITP:
11800   Exit Sub

ERRH:
11810   blnFormOpening = False
11820   DoCmd.Hourglass False
11830   Select Case ERR.Number
        Case Else
11840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11850   End Select
11860   Resume EXITP

End Sub

Private Sub cmdRevIncExp_KeyDown(KeyCode As Integer, Shift As Integer)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_KeyDown"

        Dim intRetVal As Integer

11910   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
11920   KeyCode = intRetVal
11930   Exit Sub

ERRH:
11940   intRetVal = 0
11950   Select Case ERR.Number
        Case Else
11960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11970   End Select
11980   Resume EXITP

End Sub

Private Sub cmdIncomeStatement_Click()

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdIncomeStatement_Click"

12010   With Me
12020     If blnFormOpening = False Then

12030       DoCmd.Hourglass True
12040       DoEvents

12050       .TimerInterval = 0&

12060       gblnSetFocus = True
12070       blnJustClose = True: blnFormOpening = True
12080       DoCmd.Close acForm, THIS_NAME
12090       strDocName = "frmRpt_IncomeStatement"
12100       gstrReturningForm = strDocName
12110       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
12120       If gblnGoToReport = True Then
12130         Forms(strDocName).TimerInterval = 50&
12140         blnGoingToReport = False
12150         blnGTR_Emblem = False
12160       End If

12170       DoCmd.Hourglass False

12180     End If
12190   End With

EXITP:
12200   Exit Sub

ERRH:
12210   blnFormOpening = False
12220   DoCmd.Hourglass False
12230   Select Case ERR.Number
        Case Else
12240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12250   End Select
12260   Resume EXITP

End Sub

Private Sub cmdIncomeStatement_KeyDown(KeyCode As Integer, Shift As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdIncomeStatement_KeyDown"

        Dim intRetVal As Integer

12310   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
12320   KeyCode = intRetVal
12330   Exit Sub

ERRH:
12340   intRetVal = 0
12350   Select Case ERR.Number
        Case Else
12360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12370   End Select
12380   Resume EXITP

End Sub

Private Sub cmdIncomeStatement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdIncomeStatement_MouseMove"

12410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12420   Exit Sub

ERRH:
12430   Select Case ERR.Number
        Case Else
12440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12450   End Select
12460   Resume EXITP

End Sub

Private Sub cmdCourtReports_GotFocus()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCourtReports_GotFocus"

12510   With Me
12520     strLastFocus = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
12530     .cmdCourtReports_box01.Visible = True
12540     .cmdCourtReports_box02.Visible = True
12550   End With

EXITP:
12560   Exit Sub

ERRH:
12570   Select Case ERR.Number
        Case Else
12580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12590   End Select
12600   Resume EXITP

End Sub

Private Sub cmdCourtReports_Click()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCourtReports_Click"

12710   With Me
12720     If blnFormOpening = False Then

12730       strDocName = "frmMenu_CourtReport"

12740       FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
12750       DoEvents

12760       If .chkLoadTimer.Visible = True Then
12770         chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
12780         If gblnTimer = True Then
                'gblnSetFocus = True
12790           DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
12800         End If
12810       End If
12820       DoEvents

12830       DoCmd.Hourglass True
12840       DoEvents

12850       .TimerInterval = 0&

12860       TAReports_SetZero False  ' ** Module Function: modReportFunctions.

12870       blnJustClose = True: blnFormOpening = True
12880       gdatStartDate = 0: gdatEndDate = 0: gstrCrtRpt_Account = vbNullString
12890       gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString
12900       gstrCrtRpt_NetIncome = vbNullString: gstrCrtRpt_NetLoss = vbNullString
12910       gstrCrtRpt_CashAssets_Beg = vbNullString: gstrCrtRpt_CashAssets_End = vbNullString
12920       gblnCrtRpt_Zero = False: gblnCrtRpt_ZeroDialog = False

12930       gblnSetFocus = True
12940       DoCmd.Close acForm, THIS_NAME
12950       gstrReturningForm = strDocName
12960       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
12970       If gblnGoToReport = True Then
12980         Forms(strDocName).TimerInterval = 50&
12990         blnGoingToReport = False
13000         blnGTR_Emblem = False
13010       End If

13020       DoCmd.Hourglass False

13030     End If
13040   End With

EXITP:
13050   Exit Sub

ERRH:
13060   blnFormOpening = False
13070   DoCmd.Hourglass False
13080   Select Case ERR.Number
        Case Else
13090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13100   End Select
13110   Resume EXITP

End Sub

Private Sub cmdCourtReports_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCourtReports_KeyDown"

        Dim intRetVal As Integer

13210   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
13220   KeyCode = intRetVal
13230   Exit Sub

ERRH:
13240   intRetVal = 0
13250   Select Case ERR.Number
        Case Else
13260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13270   End Select
13280   Resume EXITP

End Sub

Private Sub cmdCourtReports_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCourtReports_MouseMove"

13310   With Me
13320     If .cmdCourtReports_box01.Visible = False Then
13330       .cmdCourtReports_box01.Visible = True
13340       .cmdCourtReports_box02.Visible = True
13350     End If
13360   End With

EXITP:
13370   Exit Sub

ERRH:
13380   Select Case ERR.Number
        Case Else
13390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13400   End Select
13410   Resume EXITP

End Sub

Private Sub cmdCourtReports_LostFocus()

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCourtReports_LostFocus"

13510   With Me
13520     .cmdCourtReports_box01.Visible = False
13530     .cmdCourtReports_box02.Visible = False
13540     strLastFocus = vbNullString
13550   End With

EXITP:
13560   Exit Sub

ERRH:
13570   Select Case ERR.Number
        Case Else
13580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13590   End Select
13600   Resume EXITP

End Sub

Private Sub cmdTransByType_Click()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransByType_Click"

13710   With Me
13720     If blnFormOpening = False Then

13730       DoCmd.Hourglass True
13740       DoEvents

13750       .TimerInterval = 0&

13760       gblnSetFocus = True
13770       blnJustClose = True: blnFormOpening = True
13780       DoCmd.Close acForm, THIS_NAME
13790       strDocName = "frmRpt_TransactionsByType"
13800       gstrReturningForm = strDocName
13810       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
13820       If gblnGoToReport = True Then
13830         Forms(strDocName).TimerInterval = 50&
13840         blnGoingToReport = False
13850         blnGTR_Emblem = False
13860       End If

13870       DoCmd.Hourglass False

13880     End If
13890   End With

EXITP:
13900   Exit Sub

ERRH:
13910   blnFormOpening = False
13920   DoCmd.Hourglass False
13930   Select Case ERR.Number
        Case Else
13940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13950   End Select
13960   Resume EXITP

End Sub

Private Sub cmdTransByType_KeyDown(KeyCode As Integer, Shift As Integer)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransByType_KeyDown"

        Dim intRetVal As Integer

14010   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
14020   KeyCode = intRetVal
14030   Exit Sub

ERRH:
14040   intRetVal = 0
14050   Select Case ERR.Number
        Case Else
14060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14070   End Select
14080   Resume EXITP

End Sub

Private Sub cmdTransByType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransByType_MouseMove"

14110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14120   Exit Sub

ERRH:
14130   Select Case ERR.Number
        Case Else
14140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14150   End Select
14160   Resume EXITP

End Sub

Private Sub cmdAccountReviews_Click()

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccountReviews_Click"

14210   With Me
14220     If blnFormOpening = False Then

14230       DoCmd.Hourglass True
14240       DoEvents

14250       .TimerInterval = 0&

14260       gblnSetFocus = True
14270       blnJustClose = True: blnFormOpening = True
14280       DoCmd.Close acForm, THIS_NAME
14290       strDocName = "frmRpt_AccountReviews"
14300       gstrReturningForm = strDocName
14310       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
14320       If gblnGoToReport = True Then
14330         Forms(strDocName).TimerInterval = 50&
14340         blnGoingToReport = False
14350         blnGTR_Emblem = False
14360       End If

14370       DoCmd.Hourglass False

14380     End If
14390   End With

EXITP:
14400   Exit Sub

ERRH:
14410   blnFormOpening = False
14420   DoCmd.Hourglass False
14430   Select Case ERR.Number
        Case Else
14440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14450   End Select
14460   Resume EXITP

End Sub

Private Sub cmdAccountReviews_KeyDown(KeyCode As Integer, Shift As Integer)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccountReviews_KeyDown"

        Dim intRetVal As Integer

14510   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
14520   KeyCode = intRetVal
14530   Exit Sub

ERRH:
14540   intRetVal = 0
14550   Select Case ERR.Number
        Case Else
14560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14570   End Select
14580   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_Click()

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_Click"

14610   With Me
14620     If blnFormOpening = False Then

14630       DoCmd.Hourglass True
14640       DoEvents

14650       .TimerInterval = 0&

14660       gblnSetFocus = True
14670       blnJustClose = True: blnFormOpening = True
14680       DoCmd.Close acForm, THIS_NAME
14690       strDocName = "frmRpt_NewClosedAccounts"
14700       gstrReturningForm = strDocName
14710       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
14720       If gblnGoToReport = True Then
14730         Forms(strDocName).TimerInterval = 100&
14740         blnGoingToReport = False
14750         blnGTR_Emblem = False
14760       End If

14770       DoCmd.Hourglass False

14780     End If
14790   End With

EXITP:
14800   Exit Sub

ERRH:
14810   blnFormOpening = False
14820   DoCmd.Hourglass False
14830   Select Case ERR.Number
        Case Else
14840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14850   End Select
14860   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_KeyDown"

        Dim intRetVal As Integer

14910   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
14920   KeyCode = intRetVal
14930   Exit Sub

ERRH:
14940   intRetVal = 0
14950   Select Case ERR.Number
        Case Else
14960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14970   End Select
14980   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_Click()

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_Click"

15010   With Me
15020     If blnFormOpening = False Then

15030       DoCmd.Hourglass True
15040       DoEvents

15050       .TimerInterval = 0&

15060       gblnSetFocus = True
15070       blnJustClose = True: blnFormOpening = True
15080       DoCmd.Close acForm, THIS_NAME
15090       strDocName = "frmRpt_PurchasedSold"
15100       gstrReturningForm = strDocName
15110       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
15120       If gblnGoToReport = True Then
15130         Forms(strDocName).TimerInterval = 100&
15140         blnGoingToReport = False
15150         blnGTR_Emblem = False
15160       End If

15170       DoCmd.Hourglass False

15180     End If
15190   End With

EXITP:
15200   Exit Sub

ERRH:
15210   blnFormOpening = False
15220   DoCmd.Hourglass False
15230   Select Case ERR.Number
        Case Else
15240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15250   End Select
15260   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_KeyDown(KeyCode As Integer, Shift As Integer)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_KeyDown"

        Dim intRetVal As Integer

15310   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
15320   KeyCode = intRetVal
15330   Exit Sub

ERRH:
15340   intRetVal = 0
15350   Select Case ERR.Number
        Case Else
15360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15370   End Select
15380   Resume EXITP

End Sub

Private Sub cmdStatementOfCondition_Click()

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStatementOfCondition_Click"

15410   With Me
15420     If blnFormOpening = False Then

15430       DoCmd.Hourglass True
15440       DoEvents

15450       .TimerInterval = 0&

15460       gblnSetFocus = True
15470       blnJustClose = True: blnFormOpening = True
15480       DoCmd.Close acForm, THIS_NAME
15490       strDocName = "frmRpt_StatementOfCondition"
15500       gstrReturningForm = strDocName
15510       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
15520       If gblnGoToReport = True Then
15530         Forms(strDocName).TimerInterval = 100&
15540         blnGoingToReport = False
15550         blnGTR_Emblem = False
15560       End If

15570       DoCmd.Hourglass False

15580     End If
15590   End With

EXITP:
15600   Exit Sub

ERRH:
15610   blnFormOpening = False
15620   DoCmd.Hourglass False
15630   Select Case ERR.Number
        Case Else
15640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15650   End Select
15660   Resume EXITP

End Sub

Private Sub cmdStatementOfCondition_KeyDown(KeyCode As Integer, Shift As Integer)

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStatementOfCondition_KeyDown"

        Dim intRetVal As Integer

15710   intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.

EXITP:
15720   KeyCode = intRetVal
15730   Exit Sub

ERRH:
15740   intRetVal = 0
15750   Select Case ERR.Number
        Case Else
15760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15770   End Select
15780   Resume EXITP

End Sub

Private Sub cmdAccountProfile_Click()

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccountProfile_Click"

15810   With Me
15820     If blnFormOpening = False Then

15830       DoCmd.Hourglass True
15840       DoEvents

15850       .TimerInterval = 0&

15860       gblnSetFocus = True
15870       blnJustClose = True: blnFormOpening = True
15880       DoCmd.Close acForm, THIS_NAME
15890       strDocName = "frmRpt_AccountProfile"
15900       gstrReturningForm = strDocName
15910       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
15920       If gblnGoToReport = True Then
15930         Forms(strDocName).TimerInterval = 100&
15940         blnGoingToReport = False
15950         blnGTR_Emblem = False
15960       End If

15970       DoCmd.Hourglass False

15980     End If
15990   End With

EXITP:
16000   Exit Sub

ERRH:
16010   blnFormOpening = False
16020   DoCmd.Hourglass False
16030   Select Case ERR.Number
        Case Else
16040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16050   End Select
16060   Resume EXITP

End Sub

Private Sub cmdAccountProfile_KeyDown(KeyCode As Integer, Shift As Integer)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccountProfile_KeyDown"

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
          Case vbKeyTab
16170       With Me
16180         intRetVal = 0
16190         .cmdSearch.SetFocus
16200       End With
16210     Case Else
16220       intRetVal = cmdKeyDown(THIS_PROC, KeyCode, Shift)  ' ** Function: Below.
16230     End Select
16240   End If

        ' ** Shift keys.
16250   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16260     Select Case intRetVal
          Case vbKeyTab
16270       With Me
16280         intRetVal = 0
16290         .cmdStatementOfCondition.SetFocus
16300       End With
16310     End Select
16320   End If

EXITP:
16330   KeyCode = intRetVal
16340   Exit Sub

ERRH:
16350   intRetVal = 0
16360   Select Case ERR.Number
        Case Else
16370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16380   End Select
16390   Resume EXITP

End Sub

Private Sub cmdSearch_Click()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSearch_Click"

16410   With Me
16420     If blnFormOpening = False Then

16430       DoCmd.Hourglass True
16440       DoEvents

16450       .TimerInterval = 0&

16460       strDocName = "frmReportList"
16470       gstrReturningForm = strDocName
16480       DoCmd.OpenForm strDocName, , , , , , THIS_NAME

16490       DoCmd.Hourglass False

16500     End If
16510   End With

EXITP:
16520   Exit Sub

ERRH:
16530   blnFormOpening = False
16540   DoCmd.Hourglass False
16550   Select Case ERR.Number
        Case Else
16560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16570   End Select
16580   Resume EXITP

End Sub

Private Sub cmdSearch_KeyDown(KeyCode As Integer, Shift As Integer)

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSearch_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16620   intShiftDown = (Shift And acShiftMask) > 0
16630   intAltDown = (Shift And acAltMask) > 0
16640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16660     Select Case intRetVal
          Case vbKeyTab
16670       With Me
16680         intRetVal = 0
16690         If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
16700           .chkLoadTimer.SetFocus
16710         Else
16720           .cmdClose.SetFocus
16730         End If
16740       End With
16750     End Select
16760   End If

        ' ** Shift keys.
16770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16780     Select Case intRetVal
          Case vbKeyTab
16790       With Me
16800         intRetVal = 0
16810         .cmdAccountProfile.SetFocus
16820       End With
16830     End Select
16840   End If

EXITP:
16850   KeyCode = intRetVal
16860   Exit Sub

ERRH:
16870   intRetVal = 0
16880   Select Case ERR.Number
        Case Else
16890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16900   End Select
16910   Resume EXITP

End Sub

Private Sub chkLoadTimer_AfterUpdate()

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_AfterUpdate"

17010   With Me
17020     If IsLoaded("frmLoadTimer", acForm) = False Then  ' ** Module Function: modFileUtilities.
            ' ** Don't turn it off if it's still playing out a timing sequence.
17030       gblnTimer = False
17040     End If
17050     Select Case .chkLoadTimer
          Case True
17060       .chkLoadTimer_lbl.FontBold = True
17070       If .chkLoadTimer.Visible = True Then  ' ** This may get turned off by the Form_Timer() sequence.
17080         gblnTimer = True
17090       End If
17100     Case False
17110       .chkLoadTimer_lbl.FontBold = False
17120     End Select
17130   End With

EXITP:
17140   Exit Sub

ERRH:
17150   Select Case ERR.Number
        Case Else
17160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17170   End Select
17180   Resume EXITP

End Sub

Private Sub chkLoadTimer_KeyDown(KeyCode As Integer, Shift As Integer)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17220   intShiftDown = (Shift And acShiftMask) > 0
17230   intAltDown = (Shift And acAltMask) > 0
17240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17260     Select Case intRetVal
          Case vbKeyTab
17270       With Me
17280         intRetVal = 0
17290         .cmdClose.SetFocus
17300       End With
17310     End Select
17320   End If

        ' ** Shift keys.
17330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17340     Select Case intRetVal
          Case vbKeyTab
17350       With Me
17360         intRetVal = 0
17370         .cmdSearch.SetFocus
17380       End With
17390     End Select
17400   End If

EXITP:
17410   KeyCode = intRetVal
17420   Exit Sub

ERRH:
17430   intRetVal = 0
17440   Select Case ERR.Number
        Case Else
17450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17460   End Select
17470   Resume EXITP

End Sub

Private Function cmdKeyDown(strProc As String, intKeyCode As Integer, intShift As Integer) As Integer

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdKeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strCtl As String
        Dim lngX As Long, lngE As Long
        Dim intRetVal As Integer

17510   intRetVal = intKeyCode

        ' ** Use bit masks to determine which key was pressed.
17520   intShiftDown = (intShift And acShiftMask) > 0
17530   intAltDown = (intShift And acAltMask) > 0
17540   intCtrlDown = (intShift And acCtrlMask) > 0

        ' ** Plain keys.
17550   With Me
17560     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17570       Select Case intRetVal
            Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight

17580         strCtl = Left(strProc, (Len(strProc) - Len("_KeyDown")))  ' ** E.g., cmdStatements_KeyDown().
17590         For lngX = 0& To (lngCmds - 1&)
17600           If arr_varCmd(C_NAM, lngX) = strCtl Then
17610             lngE = lngX
17620             Exit For
17630           End If
17640         Next

17650         Select Case intRetVal
              Case vbKeyUp
17660           intRetVal = 0
17670           If .Controls(strCtl).Top > lngTopCmd_Top Then
17680             For lngX = 0& To (lngCmds - 1&)
17690               If arr_varCmd(C_COL, lngX) = arr_varCmd(C_COL, lngE) Then
17700                 If arr_varCmd(C_ORD, lngX) = (arr_varCmd(C_ORD, lngE) - _
                          IIf(arr_varCmd(C_COL, lngE) = 1& And arr_varCmd(C_ORD, lngE) = 3&, 2&, 1&)) Then
17710                   .Controls(arr_varCmd(C_NAM, lngX)).SetFocus
17720                   Exit For
17730                 End If
17740               End If
17750             Next
17760           Else
17770             For lngX = 0& To (lngCmds - 1&)
17780               If arr_varCmd(C_COL, lngX) = arr_varCmd(C_COL, lngE) And arr_varCmd(C_ORD, lngX) = lngMaxSortOrd Then
17790                 .Controls(arr_varCmd(C_NAM, lngX)).SetFocus
17800                 Exit For
17810               End If
17820             Next
17830           End If

17840         Case vbKeyDown
17850           intRetVal = 0
17860           If .Controls(strCtl).Top < lngBotCmd_Top Then
17870             For lngX = 0& To (lngCmds - 1&)
17880               If arr_varCmd(C_COL, lngX) = arr_varCmd(C_COL, lngE) Then
17890                 If arr_varCmd(C_ORD, lngX) = (arr_varCmd(C_ORD, lngE) + _
                          IIf(arr_varCmd(C_COL, lngE) = 1& And arr_varCmd(C_ORD, lngE) = 1&, 2&, 1&)) Then
17900                   .Controls(arr_varCmd(C_NAM, lngX)).SetFocus
17910                   Exit For
17920                 End If
17930               End If
17940             Next
17950           Else
17960             For lngX = 0& To (lngCmds - 1&)
17970               If arr_varCmd(C_COL, lngX) = arr_varCmd(C_COL, lngE) And arr_varCmd(C_ORD, lngX) = 1& Then
17980                 .Controls(arr_varCmd(C_NAM, lngX)).SetFocus
17990                 Exit For
18000               End If
18010             Next
18020           End If

18030         Case vbKeyLeft
18040           intRetVal = 0
18050           If .Controls(strCtl).Left > lngLeftCmd_Left Then
18060             For lngX = 0& To (lngCmds - 1&)
18070               If arr_varCmd(C_COL, lngX) = 1& And arr_varCmd(C_ORD, lngX) = _
                        IIf(arr_varCmd(C_ORD, lngE) = 2&, 1&, arr_varCmd(C_ORD, lngE)) Then
18080                 .Controls(arr_varCmd(C_NAM, lngX)).SetFocus
18090                 Exit For
18100               End If
18110             Next
18120           Else
18130             For lngX = 0& To (lngCmds - 1&)
18140               If arr_varCmd(C_COL, lngX) = 2& And arr_varCmd(C_ORD, lngX) = arr_varCmd(C_ORD, lngE) Then
18150                 .Controls(arr_varCmd(C_NAM, lngX)).SetFocus
18160                 Exit For
18170               End If
18180             Next
18190           End If

18200         Case vbKeyRight
18210           intRetVal = 0
18220           If .Controls(strCtl).Left < lngRightCmd_Left Then
18230             For lngX = 0& To (lngCmds - 1&)
18240               If arr_varCmd(C_COL, lngX) = 2& And arr_varCmd(C_ORD, lngX) = arr_varCmd(C_ORD, lngE) Then
18250                 .Controls(arr_varCmd(C_NAM, lngX)).SetFocus
18260                 Exit For
18270               End If
18280             Next
18290           Else
18300             For lngX = 0& To (lngCmds - 1&)
18310               If arr_varCmd(C_COL, lngX) = 1& And arr_varCmd(C_ORD, lngX) = _
                        IIf(arr_varCmd(C_ORD, lngE) = 2&, 1&, arr_varCmd(C_ORD, lngE)) Then
18320                 .Controls(arr_varCmd(C_NAM, lngX)).SetFocus
18330                 Exit For
18340               End If
18350             Next
18360           End If

18370         End Select

18380       Case Else
              ' ** Don't care.
18390       End Select

18400     End If

18410   End With

EXITP:
18420   cmdKeyDown = intRetVal
18430   Exit Function

ERRH:
18440   intRetVal = 0
18450   Select Case ERR.Number
        Case Else
18460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18470   End Select
18480   Resume EXITP

End Function

Public Sub GTREmblem_Off()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

18510   With Me
18520     blnGTR_Emblem = False
18530     For lngX = 1& To 24&
18540 On Error Resume Next
18550       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
18560       If ERR.Number <> 0 Then
18570 On Error GoTo ERRH
18580         Exit For
18590       Else
18600 On Error GoTo ERRH
18610       End If
18620     Next
18630   End With

EXITP:
18640   Exit Sub

ERRH:
18650   Select Case ERR.Number
        Case Else
18660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18670   End Select
18680   Resume EXITP

End Sub
