VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_IncomeExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_IncomeExpense"

'VGC 10/07/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Load()
' ##   Form_Timer()
' ##   cmdMoreOptions_Click()
' #######################################

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' #########################
' ## Use VBA_RenumErrh().  39870
' #########################

' ** For GTR, sets options, then stops.

' ** Shortcut Alt keys responsive from this form:
' **   Select Account:   A {cmbAccounts}
' **   By Number:        B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Expense Word:     D {cmdRevIncExp_ExpenseWord}
' **   Income Excel:     E {cmdRevIncExp_IncomeExcel}
' **   From Date:        F {DateStart}
' **   Include Archive:  H {chkIncludeArchive}
' **   More Options:     I {cmdMoreOptions}
' **   All Accts:        L {opgAccountNumber_optAll}
' **   By Name:          M {opgAccountSource_optName}
' **   Print Expense:    N {cmdRevIncExp_ExpensePrint}
' **   Income Word:      O {cmdRevIncExp_IncomeWord}
' **   Print Income:     P {cmdRevIncExp_IncomePrint}
' **   Reset:            S {cmdReset}
' **   To Date:          T {DateEnd}
' **   Specified Acct:   U {opgAccountNumber_optSpecified}
' **   Preview Expense:  V {cmdRevIncExp_ExpensePreview}
' **   Preview Income:   W {cmdRevIncExp_IncomePreview}
' **   Exit:             X {cmdClose}
' **   Expense Excel:    Z {cmdRevIncExp_ExpenseExcel}
'Unused: G J Q Y

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Print All:        A {cmdPrintAll}
' **   Excel All:        E {cmdExcelAll}
' **   Page Count:       G {chkPageOf}
' **   Open Excel:       N {chkOpenExcel}
' **   Open Word:        O {chkOpenWord}
' **   Word All:         W {cmdWordAll}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   PCash Exclude:    E {opgPrincipalCash_optExclude}
' **   PCash Include:    I {opgPrincipalCash_optInclude}
' **   PCash Only:       O {opgPrincipalCash_optOnly}

' ** Shortcut Ctrl-Alt keys responsive from this form:
' **   Zero Exclude:     E {opgZeroCash_optExclude}
' **   Zero Include:     I {opgZeroCash_optInclude}
' **   Zero Only:        O {opgZeroCash_optOnly}

' ** Shortcut Alt-Shift keys responsive from this form:
' **   Show Acct Info:   A {chkAcctEveryLine}
' **   Summary Exclude:  E {opgSummary_optExclude}
' **   Show Detail:      H {chkDetail}
' **   Summary Include:  I {opgSummary_optInclude}
' **   Multi-Lot Sales:  M {chkDontCombineMulti}
' **   Summary Only:     O {opgSummary_optOnly}
' **   Sweep Only:       S {chkSweepOnly}

' ** Shortcut Ctrl-Alt-Shift keys responsive from this form:
' **   Unspec Exclude:   E {opgUnspecified_optExclude}
' **   Unspec Include:   I {opgUnspecified_optInclude}
' **   Unspec Only:      O {opgUnspecified_optOnly}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

Private blnMoreOptions_Focus As Boolean, blnMoreOptions_MouseDown As Boolean
Private blnReset_Focus As Boolean, blnReset_MouseDown As Boolean
Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean
Private blnPrintAll_Focus As Boolean, blnWordAll_Focus As Boolean, blnExcelAll_Focus As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngDetail_Height As Long, lngIncBtn_Top As Long, lngExpBtn_Top As Long, lngBtnLbl_Diff As Long
Private lngAccountNo_Top As Long, lngAccountNoBox_Top As Long, lngAccountNoLbl_Top As Long, lngRemMe_Top As Long, lngDiff As Long
Private lngAcctSrc_Top As Long, lngAcctSrcOpt_Top As Long, lngAcctSrcOptLbl_Top As Long, lngPageOf_Top As Long
Private lngAllBtn_Top As Long, lngHighlight_Offset As Long, lngDetailLine1_Top As Long, lngDetailLine3_Top As Long

Private strFile1 As String, strFile2 As String, strFile3 As String, strFile4 As String
Private lngShortcutLbl_Clr As Long, blnGoingToReport2 As Boolean, blnGoingToReport3 As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnPrintBoth As Boolean, blnGoingToReport As Boolean, blnWasGTR As Boolean, blnOptsOpened As Boolean
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           DoCmd.Hourglass True  ' ** Make sure it's still running.
170           DoEvents

180           blnWasGTR = gblnGoToReport
190           blnGTR_Emblem = False: lngGTR_Stat = 0&

200           gstrFormQuerySpec = THIS_NAME
210           blnPrintBoth = False: blnOptsOpened = False
220           blnMoreOptions_Focus = False: blnMoreOptions_MouseDown = False
230           blnReset_Focus = False: blnReset_MouseDown = False
240           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
250           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False
260           blnPrintAll_Focus = False: blnWordAll_Focus = False: blnExcelAll_Focus = False

              ' ** Reset the variables to ensure we get a new table every time we come into the form.
270           gstrAccountNo = vbNullString
280           gstrAccountName = vbNullString
290           gdatStartDate = 0: gdatEndDate = 0

      #If NoExcel Then
300           .cmdRevIncExp_IncomeExcel.Enabled = False
310           .cmdRevIncExp_ExpenseExcel.Enabled = False
320           .cmdExcelAll.Enabled = False
330           .chkOpenExcel.Enabled = False
340           .chkOpenExcel_lbl2.ForeColor = WIN_CLR_DISF
350           .chkOpenExcel_lbl2_dim_hi.Visible = True
      #Else
360           .cmdRevIncExp_IncomeExcel.Enabled = True
370           .cmdRevIncExp_ExpenseExcel.Enabled = True
380           .cmdExcelAll.Enabled = True
390           .chkOpenExcel.Enabled = True
400           .chkOpenExcel_lbl2.ForeColor = CLR_DKGRY
410           .chkOpenExcel_lbl2_dim_hi.Visible = False
      #End If

              ' ** EVENT CHECK: chkRememberMe!
              ' ** NONE!
              ' ** EVENT CHECK: chkRememberDates!
420           If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
430             DateStart_Exit Cancel  ' ** Procedure: Below.
440             DateEnd_Exit Cancel  ' ** Procedure: Below.
450           End If

460           lngShortcutLbl_Clr = CLR_VDKGRY

470           chkOpenWord_AfterUpdate  ' ** Procedure: Below.
480           chkOpenExcel_AfterUpdate  ' ** Procedure: Below.

490           blnIsOpen = False

500         Else
510           Cancel = -1
520         End If
530       Else
540         Cancel = -1
550       End If
560     End With

570     If Cancel = -1 Then
580       Beep
590       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
600     End If

610     DoCmd.Hourglass False

EXITP:
620     Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub Form_Load()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim lngTmp01 As Long

710     With Me

720       DoCmd.Hourglass True  ' ** Make sure it's still running.
730       DoEvents

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
740       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
750       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
760       clsMonthClass.PositionAtCursor = True

          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
770       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

          ' ** Just get these out of the way.
780       .detail_height_line.Top = 0&
790       .GoToReport_arw_rptinc_img_lbl.Top = 0&
800       .GoToReport_arw_rptexp_img_lbl.Top = 0&

810       lngDetail_Height = .Detail.Height
820       lngIncBtn_Top = .cmdRevIncExp_IncomePreview.Top
830       lngExpBtn_Top = .cmdRevIncExp_ExpensePreview.Top
840       lngBtnLbl_Diff = (.cmdRevIncExp_Income_lbl.Top - .cmdRevIncExp_IncomePreview.Top)
850       lngAllBtn_Top = .cmdPrintAll.Top
860       lngHighlight_Offset = (.cmdPrintAll.Top - .cmdPrintAll_box01.Top)
870       lngDetailLine1_Top = .Detail_hline01.Top
880       lngDetailLine3_Top = .Detail_hline03.Top
890       lngAccountNoBox_Top = .cmbAccounts_box.Top
900       lngAccountNoLbl_Top = .cmbAccounts_lbl.Top
910       lngAccountNo_Top = .cmbAccounts.Top
920       lngRemMe_Top = .chkRememberMe.Top
930       lngAcctSrc_Top = .opgAccountSource.Top
940       lngAcctSrcOpt_Top = .opgAccountSource_optNumber.Top
950       lngAcctSrcOptLbl_Top = .opgAccountSource_optNumber_lbl.Top
960       lngPageOf_Top = .chkPageOf.Top
970       lngDiff = ((.opgOptionGroups_box.Top + .opgOptionGroups_box.Height) - (.opgOptionGroups_box3.Top + .opgOptionGroups_box3.Height))

980       lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
990       lngMonitorNum = 1&: lngTmp01 = 0&
1000      EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
1010      If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

1020      GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

1030      FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

1040      If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
1050      DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, lngFrm_Height  'lngFrm_Top
1060      If lngMonitorNum > 1& Then
1070        LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
1080      End If

1090      .frm_top = lngFrm_Top
1100      .frm_left = lngFrm_Left
1110      .frm_width = lngFrm_Width
1120      .frm_height = lngFrm_Height

1130      ArchiveSet_IE Me  ' ** Procedure: Below.

1140      Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
1150      DoEvents

1160      FormLoad_IE 1, Me

1170      If .chkOptionsOpen = False Then
1180        ShowOptions_IE "Hide", lngDetail_Height, lngDiff, lngIncBtn_Top, lngExpBtn_Top, lngBtnLbl_Diff, _
              lngDetailLine1_Top, lngAccountNo_Top, lngAccountNoBox_Top, lngAccountNoLbl_Top, lngRemMe_Top, _
              lngAcctSrc_Top, lngAcctSrcOpt_Top, lngAcctSrcOptLbl_Top, lngPageOf_Top, lngAllBtn_Top, _
              lngHighlight_Offset, lngDetailLine3_Top, Me  ' ** Module Procedure: modIncExpFuncs.
1190      End If

1200      FormLoad_IE 2, Me

1210      DoCmd.Hourglass False

1220    End With

EXITP:
1230    Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
1310    intShiftDown = (Shift And acShiftMask) > 0
1320    intAltDown = (Shift And acAltMask) > 0
1330    intCtrlDown = (Shift And acCtrlMask) > 0

1340    intRetVal = KeyCode

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   By Number:        B {opgAccountSource_optNumber}
        ' **   By Name:          M {opgAccountSource_optName}
        ' **   Exit:             X {cmdClose}
        ' **   Expense Excel:    Z {cmdRevIncExp_ExpenseExcel}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Print All:        A {cmdPrintAll}
        ' **   Excel All:        E {cmdExcelAll}
        ' **   Page Count:       G {chkPageOf}
        ' **   Open Excel:       N {chkOpenExcel}
        ' **   Open Word:        O {chkOpenWord}
        ' **   Word All:         W {cmdWordAll}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   PCash Exclude:    E {opgPrincipalCash_optExclude}
        ' **   PCash Include:    I {opgPrincipalCash_optInclude}
        ' **   PCash Only:       O {opgPrincipalCash_optOnly}

        ' ** Shortcut Ctrl-Alt keys to other forms and functionality:
        ' **   Zero Exclude:     E {opgZeroCash_optExclude}
        ' **   Zero Include:     I {opgZeroCash_optInclude}
        ' **   Zero Only:        O {opgZeroCash_optOnly}

        ' ** Shortcut Alt-Shift keys to other forms and functionality:
        ' **   Show Acct Info:   A {chkAcctEveryLine}
        ' **   Summary Exclude:  E {opgSummary_optExclude}
        ' **   Show Detail:      H {chkDetail}
        ' **   Summary Include:  I {opgSummary_optInclude}
        ' **   Multi-Lot Sales:  M {chkDontCombineMulti}
        ' **   Summary Only:     O {opgSummary_optOnly}
        ' **   Sweep Only:       S {chkSweepOnly}

        ' ** Shortcut Ctrl-Alt-Shift keys to other forms and functionality:
        ' **   Unspec Exclude:   E {opgUnspecified_optExclude}
        ' **   Unspec Include:   I {opgUnspecified_optInclude}
        ' **   Unspec Only:      O {opgUnspecified_optOnly}

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyEscape
1370        intRetVal = 0
1380        cmdClose_Click  ' ** Procedure: Below.
1390      Case vbKeyF7
1400        intRetVal = 0
1410        cmdCalendar1_Click  ' ** Procedure: Below.
1420      Case vbKeyF8
1430        intRetVal = 0
1440        cmdCalendar2_Click  ' ** Procedure: Below.
1450      End Select
1460    End If

        ' ** Alt keys.
1470    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1480      Select Case intRetVal
          Case vbKeyB
1490        With Me
1500          intRetVal = 0
1510          If .opgAccountSource.Enabled = True Then
1520            .opgAccountSource.SetFocus
1530            .opgAccountSource = .opgAccountSource_optNumber.OptionValue
1540            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1550          Else
1560            Beep
1570          End If
1580        End With
1590      Case vbKeyM
1600        With Me
1610          intRetVal = 0
1620          If .opgAccountSource.Enabled = True Then
1630            .opgAccountSource.SetFocus
1640            .opgAccountSource = .opgAccountSource_optName.OptionValue
1650            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1660          Else
1670            Beep
1680          End If
1690        End With
1700      Case vbKeyX
1710        intRetVal = 0
1720        cmdClose_Click  ' ** Procedure: Below.
1730      Case vbKeyZ
1740        With Me
1750          intRetVal = 0
1760          If .cmdRevIncExp_ExpenseExcel.Enabled = True Then
1770            .cmdRevIncExp_ExpenseExcel.SetFocus
1780            cmdRevIncExp_ExpenseExcel_Click  ' ** Procedure: Below.
1790          End If
1800        End With
1810      End Select
1820    End If

        ' ** Ctrl keys.
1830    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1840      Select Case intRetVal
          Case vbKeyA
1850        With Me
1860          intRetVal = 0
1870          .cmdPrintAll.SetFocus
1880          cmdPrintAll_Click  ' ** Procedure: Below.
1890        End With
1900      Case vbKeyE
1910        With Me
1920          intRetVal = 0
1930          If .cmdExcelAll.Enabled = True Then
1940            .cmdExcelAll.SetFocus
1950            cmdExcelAll_Click  ' ** Procedure: Below.
1960          Else
1970            Beep
1980          End If
1990        End With
2000      Case vbKeyG
2010        With Me
2020          intRetVal = 0
2030          If .chkPageOf.Enabled = True Then
2040            .chkPageOf.SetFocus
2050            .chkPageOf = (Not .chkPageOf)
2060            chkPageOf_AfterUpdate  ' ** Procedure: Below.
2070          End If
2080        End With
2090      Case vbKeyN
2100        With Me
2110          intRetVal = 0
2120          If .chkOpenExcel.Enabled = True Then
2130            .chkOpenExcel.SetFocus
2140            .chkOpenExcel = (Not .chkOpenExcel)
2150            chkOpenExcel_AfterUpdate  ' ** Procedure: Below.
2160          Else
2170            Beep
2180          End If
2190        End With
2200      Case vbKeyO
2210        With Me
2220          intRetVal = 0
2230          .chkOpenWord.SetFocus
2240          .chkOpenWord = (Not .chkOpenWord)
2250          chkOpenWord_AfterUpdate  ' ** Procedure: Below.
2260        End With
2270      Case vbKeyW
2280        With Me
2290          intRetVal = 0
2300          .cmdWordAll.SetFocus
2310          cmdWordAll_Click  ' ** Procedure: Below.
2320        End With
2330      End Select
2340    End If

        ' ** Ctrl-Shift keys.
2350    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2360      Select Case intRetVal
          Case vbKeyE
2370        With Me
2380          intRetVal = 0
2390          If .opgPrincipalCash.Visible = True Then
2400            If .opgPrincipalCash_optExclude.Visible = True And .opgPrincipalCash_optExclude.Enabled = True Then
2410              .opgPrincipalCash.SetFocus
2420              .opgPrincipalCash = .opgPrincipalCash_optExclude.OptionValue
2430              opgPrincipalCash_AfterUpdate  ' ** Procedure: Below.
2440            Else
2450              Beep
2460            End If
2470          Else
2480            Beep
2490          End If
2500        End With
2510      Case vbKeyF
2520        With Me
2530          intRetVal = 0
2540          .FocusHolder.SetFocus
2550          .DateStart = Null
2560        End With
2570      Case vbKeyI
2580        With Me
2590          intRetVal = 0
2600          If .opgPrincipalCash.Visible = True Then
2610            If .opgPrincipalCash_optInclude.Visible = True And .opgPrincipalCash_optInclude.Enabled = True Then
2620              .opgPrincipalCash.SetFocus
2630              .opgPrincipalCash = .opgPrincipalCash_optInclude.OptionValue
2640              opgPrincipalCash_AfterUpdate  ' ** Procedure: Below.
2650            Else
2660              Beep
2670            End If
2680          Else
2690            Beep
2700          End If
2710        End With
2720      Case vbKeyO
2730        With Me
2740          intRetVal = 0
2750          If .opgPrincipalCash.Visible = True Then
2760            If .opgPrincipalCash_optOnly.Visible = True And .opgPrincipalCash_optOnly.Enabled = True Then
2770              .opgPrincipalCash.SetFocus
2780              .opgPrincipalCash = .opgPrincipalCash_optOnly.OptionValue
2790              opgPrincipalCash_AfterUpdate  ' ** Procedure: Below.
2800            Else
2810              Beep
2820            End If
2830          Else
2840            Beep
2850          End If
2860        End With
2870      End Select
2880    End If

        ' ** Ctrl-Alt keys.
2890    If intCtrlDown And intAltDown And (Not intShiftDown) Then
2900      Select Case intRetVal
          Case vbKeyE
2910        With Me
2920          intRetVal = 0
2930          If .opgZeroCash.Visible = True Then
2940            If .opgZeroCash_optExclude.Visible = True And .opgZeroCash_optExclude.Enabled = True Then
2950              .opgZeroCash.SetFocus
2960              .opgZeroCash = .opgZeroCash_optExclude.OptionValue
2970              opgZeroCash_AfterUpdate  ' ** Procedure: Below.
2980            Else
2990              Beep
3000            End If
3010          Else
3020            Beep
3030          End If
3040        End With
3050      Case vbKeyI
3060        With Me
3070          intRetVal = 0
3080          If .opgZeroCash.Visible = True Then
3090            If .opgZeroCash_optInclude.Visible = True And .opgZeroCash_optInclude.Enabled = True Then
3100              .opgZeroCash.SetFocus
3110              .opgZeroCash = .opgZeroCash_optInclude.OptionValue
3120              opgZeroCash_AfterUpdate  ' ** Procedure: Below.
3130            Else
3140              Beep
3150            End If
3160          Else
3170            Beep
3180          End If
3190        End With
3200      Case vbKeyO
3210        With Me
3220          intRetVal = 0
3230          If .opgZeroCash.Visible = True Then
3240            If .opgZeroCash_optOnly.Visible = True And .opgZeroCash_optOnly.Enabled = True Then
3250              .opgZeroCash.SetFocus
3260              .opgZeroCash = .opgZeroCash_optInclude.OptionValue
3270              opgZeroCash_AfterUpdate  ' ** Procedure: Below.
3280            Else
3290              Beep
3300            End If
3310          Else
3320            Beep
3330          End If
3340        End With
3350      End Select
3360    End If

        ' ** Alt-Shift keys.
3370    If (Not intCtrlDown) And intAltDown And intShiftDown Then
3380      Select Case intRetVal
          Case vbKeyA
3390        With Me
3400          intRetVal = 0
3410          If .chkAcctEveryLine.Visible = True And .chkAcctEveryLine.Enabled = True Then
3420            .chkAcctEveryLine.SetFocus
3430            .chkAcctEveryLine = (Not .chkAcctEveryLine)
3440            chkAcctEveryLine_AfterUpdate  ' ** Procedure: Below.
3450          End If
3460        End With
3470      Case vbKeyE
3480        With Me
3490          intRetVal = 0
3500          If .opgSummary.Visible = True Then
3510            If .opgSummary_optExclude.Visible = True And .opgSummary_optExclude.Enabled = True Then
3520              .opgSummary.SetFocus
3530              .opgSummary = .opgSummary_optExclude.OptionValue
3540              opgSummary_AfterUpdate  ' ** Procedure: Below.
3550            Else
3560              Beep
3570            End If
3580          Else
3590            Beep
3600          End If
3610        End With
3620      Case vbKeyH
3630        With Me
3640          intRetVal = 0
3650          .chkDetail.SetFocus
3660          .chkDetail = (Not .chkDetail)
3670          chkDetail_AfterUpdate  ' ** Procedure: Below.
3680        End With
3690      Case vbKeyI
3700        With Me
3710          intRetVal = 0
3720          If .opgSummary.Visible = True Then
3730            If .opgSummary_optInclude.Visible = True And .opgSummary_optInclude.Enabled = True Then
3740              .opgSummary.SetFocus
3750              .opgSummary = .opgSummary_optInclude.OptionValue
3760              opgSummary_AfterUpdate  ' ** Procedure: Below.
3770            Else
3780              Beep
3790            End If
3800          Else
3810            Beep
3820          End If
3830        End With
3840      Case vbKeyM
3850        With Me
3860          intRetVal = 0
3870          If .chkDontCombineMulti.Visible = True And .chkDontCombineMulti.Enabled = True Then
3880            .chkDontCombineMulti.SetFocus
3890            .chkDontCombineMulti = (Not .chkDontCombineMulti)
3900            chkDontCombineMulti_AfterUpdate  ' ** Procedure: Below.
3910          End If
3920        End With
3930      Case vbKeyO
3940        With Me
3950          intRetVal = 0
3960          If .opgSummary.Visible = True Then
3970            If .opgSummary_optOnly.Visible = True And .opgSummary_optOnly.Enabled = True Then
3980              .opgSummary.SetFocus
3990              .opgSummary = .opgSummary_optOnly.OptionValue
4000              opgSummary_AfterUpdate  ' ** Procedure: Below.
4010            Else
4020              Beep
4030            End If
4040          Else
4050            Beep
4060          End If
4070        End With
4080      Case vbKeyS
4090        With Me
4100          intRetVal = 0
4110          If .chkSweepOnly.Visible = True And .chkSweepOnly.Enabled = True Then
4120            .chkSweepOnly.SetFocus
4130            .chkSweepOnly = (Not .chkSweepOnly)
4140            chkSweepOnly_AfterUpdate  ' ** Procedure: Below.
4150          End If
4160        End With
4170      End Select
4180    End If

        ' ** Ctrl-Alt-Shift keys.
4190    If intCtrlDown And intAltDown And intShiftDown Then
4200      Select Case intRetVal
          Case vbKeyE
4210        With Me
4220          intRetVal = 0
4230          If .opgUnspecified.Visible = True Then
4240            If .opgUnspecified_optExclude.Visible = True And .opgUnspecified_optExclude.Enabled = True Then
4250              .opgUnspecified.SetFocus
4260              .opgUnspecified = .opgUnspecified_optExclude.OptionValue
4270              opgUnspecified_AfterUpdate  ' ** Procedure: Below.
4280            Else
4290              Beep
4300            End If
4310          Else
4320            Beep
4330          End If
4340        End With
4350      Case vbKeyI
4360        With Me
4370          intRetVal = 0
4380          If .opgUnspecified.Visible = True Then
4390            If .opgUnspecified_optInclude.Visible = True And .opgUnspecified_optInclude.Enabled = True Then
4400              .opgUnspecified.SetFocus
4410              .opgUnspecified = .opgUnspecified_optInclude.OptionValue
4420              opgUnspecified_AfterUpdate  ' ** Procedure: Below.
4430            Else
4440              Beep
4450            End If
4460          Else
4470            Beep
4480          End If
4490        End With
4500      Case vbKeyO
4510        With Me
4520          intRetVal = 0
4530          If .opgUnspecified.Visible = True Then
4540            If .opgUnspecified_optOnly.Visible = True And .opgUnspecified_optOnly.Enabled = True Then
4550              .opgUnspecified.SetFocus
4560              .opgUnspecified = .opgUnspecified_optOnly.OptionValue
4570              opgUnspecified_AfterUpdate  ' ** Procedure: Below.
4580            Else
4590              Beep
4600            End If
4610          Else
4620            Beep
4630          End If
4640        End With
4650      End Select
4660    End If

EXITP:
4670    KeyCode = intRetVal
4680    Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
670     Resume EXITP

End Sub

Private Sub Form_Timer()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim blnIsSum As Boolean
        Dim strFormName As String

        Static lngTmp01 As Long, lngTmp02 As Long

4710    With Me
4720      .TimerInterval = 0&
4730      If gblnGoToReport = True Then
4740        If garr_varGoToReport(GTR_RNAM) = "rptIncExp_Income_Summary" Or garr_varGoToReport(GTR_RNAM) = "rptIncExp_Expenses_Summary" Then
4750          blnIsSum = True
4760        End If
4770      End If
4780      If blnGTR_Emblem = True Then
4790        lngGTR_Stat = lngGTR_Stat + 1&
4800        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
4810        lngTmp01 = lngTmp01 + 1&
4820        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
4830        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
4840        If lngTmp01 > 1& Then
4850          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
4860        End If
4870        If lngTmp01 = 24& Then
4880          lngTmp01 = 0&
4890          lngTmp02 = lngTmp02 + 1&
4900          If lngTmp02 < 5& Then
4910            .TimerInterval = 50&
4920          Else
4930            blnGTR_Emblem = False
4940          End If
4950        Else
4960          .TimerInterval = 50&
4970        End If
4980      End If  ' ** blnGTR_Emblem.
4990      If gblnSetFocus = True Then
5000        gblnSetFocus = False
5010        strFormName = vbNullString
5020  On Error Resume Next
5030        Do Until strFormName = THIS_NAME
5040          DoCmd.SelectObject acForm, THIS_NAME, False
5050          strFormName = Screen.ActiveForm.Name
5060        Loop
5070  On Error GoTo ERRH
5080        CmdBars_Hide True  ' ** Module Procedure: modWindowFunctions.
5090      End If
5100      If gblnGoToReport = True Then
5110        If blnIsSum = False And ((lngGTR_Stat > 0& And lngGTR_Stat < 24&) Or (lngGTR_Stat > 24& And lngGTR_Stat < 48&) Or _
                (lngGTR_Stat > 48& And lngGTR_Stat < 72&) Or (lngGTR_Stat > 72& And lngGTR_Stat < 96&) Or (lngGTR_Stat > 96&)) Then
              ' ** Just let the emblem go.
5120          DoCmd.Hourglass True  ' ** Make sure it's still running.
5130          DoEvents
5140        ElseIf blnIsSum = True And ((lngGTR_Stat > 0& And lngGTR_Stat < 36&) Or (lngGTR_Stat > 36& And lngGTR_Stat < 72&) Or _
                (lngGTR_Stat > 72& And lngGTR_Stat < 108&) Or (lngGTR_Stat > 108& And lngGTR_Stat < 144&) Or (lngGTR_Stat > 144&)) Then
              ' ** Let the emblem go a little longer.
5150          DoCmd.Hourglass True  ' ** Make sure it's still running.
5160          DoEvents
5170        Else
5180          If blnGTR_Emblem = False And garr_varGoToReport(GTR_RNAM) <> "rptIncExp_Income" And _
                  garr_varGoToReport(GTR_RNAM) <> "rptIncExp_Expenses" Then
                ' ** No emblem for the simple ones.
5190            blnGTR_Emblem = True
5200            lngTmp01 = 0&: lngTmp02 = 0&
5210          End If
5220          DoEvents
5230          Select Case blnGoingToReport
              Case True
5240            Select Case blnGoingToReport2
                Case True
5250              Select Case blnGoingToReport3
                  Case True
5260                blnGoingToReport3 = False
5270                .GoToReport_arw_incexp_sum_img.Visible = False
5280                Select Case garr_varGoToReport(GTR_RNAM)
                    Case "rptIncExp_Income_Summary"
5290                  .cmdRevIncExp_IncomeWord.Visible = False    'INC
5300                  .GoToReport_arw_rptinc_img.Visible = True
5310                  .cmdRevIncExp_IncomePrint.SetFocus
5320                Case "rptIncExp_Expenses_Summary"
5330                  .cmdRevIncExp_ExpenseWord.Visible = False   'EXP
5340                  .GoToReport_arw_rptexp_img.Visible = True
5350                  .cmdRevIncExp_ExpensePrint.SetFocus
5360                End Select
5370                Beep
5380                DoEvents
5390                blnGTR_Emblem = False
5400                GTREmblem_Off  ' ** Procedure: Below.
5410                DoCmd.Hourglass False 'Beep
5420                DoEvents
5430                .TimerInterval = (GTR_WAIT * 2&)
                    ' ** Pointing to Print button.
5440              Case False
5450                Select Case garr_varGoToReport(GTR_RNAM)
                    Case "rptIncExp_Income_Detailed", "rptIncExp_Expenses_Detailed"
5460                  blnGoingToReport2 = False
5470                  blnGoingToReport = False
5480                  gblnGoToReport = False
5490                  .GoToReport_arw_rptinc_img.Visible = False
5500                  .GoToReport_arw_rptexp_img.Visible = False
5510                  .cmdRevIncExp_IncomeWord.Visible = True    'INC
5520                  .cmdRevIncExp_ExpenseWord.Visible = True   'EXP
5530                  .cmdClose.SetFocus
5540                  DoCmd.Hourglass False
5550                  DoEvents
5560                Case "rptIncExp_Income_Summary", "rptIncExp_Expenses_Summary"
5570                  If .GoToReport_arw_rptinc_img.Visible = True Or .GoToReport_arw_rptexp_img.Visible = True Then
5580                    blnGoingToReport2 = False
5590                    blnGoingToReport = False
5600                    gblnGoToReport = False
5610                    .GoToReport_arw_rptinc_img.Visible = False
5620                    .GoToReport_arw_rptexp_img.Visible = False
5630                    .cmdRevIncExp_IncomeWord.Visible = True    'INC
5640                    .cmdRevIncExp_ExpenseWord.Visible = True   'EXP
5650                    .cmdClose.SetFocus
5660                    DoCmd.Hourglass False
5670                    DoEvents
5680                  Else
5690                    Select Case blnOptsOpened
                        Case True
5700                      .opgSummary.SetFocus
5710                      .GoToReport_arw_incexp_opts_c_img.Visible = False
5720                      .GoToReport_arw_incexp_opts_cr_img.Visible = False
5730                      .GoToReport_arw_incexp_opts_o_img.Visible = False
5740                      .GoToReport_arw_incexp_sum_img.Visible = True
5750                      DoEvents
5760                      If .opgSummary <> .opgSummary_optOnly.OptionValue Then
5770                        .opgSummary = .opgSummary_optOnly.OptionValue
5780                        opgSummary_AfterUpdate  ' ** Procedure: Below.
5790                      End If
5800                      DoEvents
5810                      blnGoingToReport3 = True
                          ' ** Pointing to Summery option.
5820                      .TimerInterval = 50&
5830                    Case False
5840                      .GoToReport_arw_incexp_sum_img.Visible = False
5850                      Select Case garr_varGoToReport(GTR_RNAM)
                          Case "rptIncExp_Income_Summary"
5860                        .cmdRevIncExp_IncomeWord.Visible = False    'INC
5870                        .GoToReport_arw_rptinc_img.Visible = True
5880                        .cmdRevIncExp_IncomePrint.SetFocus
5890                      Case "rptIncExp_Expenses_Summary"
5900                        .cmdRevIncExp_ExpenseWord.Visible = False   'EXP
5910                        .GoToReport_arw_rptexp_img.Visible = True
5920                        .cmdRevIncExp_ExpensePrint.SetFocus
5930                      End Select
5940                      Beep
5950                      DoEvents
5960                      blnGTR_Emblem = False
5970                      GTREmblem_Off  ' ** Procedure: Below.
5980                      DoCmd.Hourglass False 'Beep
5990                      DoEvents
6000                      .TimerInterval = (GTR_WAIT * 2&)
                          ' ** Pointing to Print button.
6010                    End Select
6020                  End If
6030                End Select
6040              End Select
6050            Case False
6060              Select Case garr_varGoToReport(GTR_RNAM)
                  Case "rptIncExp_Income", "rptIncExp_Expenses"
6070                blnGoingToReport = False
6080                gblnGoToReport = False
6090                .GoToReport_arw_rptinc_img.Visible = False
6100                .GoToReport_arw_rptexp_img.Visible = False
6110                .cmdRevIncExp_IncomeWord.Visible = True    'INC
6120                .cmdRevIncExp_ExpenseWord.Visible = True   'EXP
6130                .cmdClose.SetFocus
6140                DoCmd.Hourglass False
6150                DoEvents
6160              Case "rptIncExp_Income_Detailed", "rptIncExp_Expenses_Detailed"
6170                blnGoingToReport2 = True
6180                .GoToReport_arw_incexp_det_img.Visible = False
                    ' ** lngGTR_Stat = 24
6190                .cmdRevIncExp_IncomePrint.SetFocus
6200                .cmdRevIncExp_IncomeWord.Visible = False    'INC
6210                .GoToReport_arw_rptinc_img.Visible = True
6220                DoEvents
6230                blnGTR_Emblem = False
6240                GTREmblem_Off  ' ** Procedure: Below.
6250                Beep
6260                DoCmd.Hourglass False
6270                DoEvents
6280                .TimerInterval = (GTR_WAIT * 2&)
6290              Case "rptIncExp_Income_Summary", "rptIncExp_Expenses_Summary"
6300                blnGoingToReport2 = True
6310                Select Case blnOptsOpened
                    Case True
6320                  ShowOptions_IE "Show", lngDetail_Height, lngDiff, lngIncBtn_Top, lngExpBtn_Top, _
                        lngBtnLbl_Diff, lngDetailLine1_Top, lngAccountNo_Top, lngAccountNoBox_Top, _
                        lngAccountNoLbl_Top, lngRemMe_Top, lngAcctSrc_Top, lngAcctSrcOpt_Top, _
                        lngAcctSrcOptLbl_Top, lngPageOf_Top, lngAllBtn_Top, lngHighlight_Offset, _
                        lngDetailLine3_Top, Me  ' ** Module Procedure: modIncExpFuncs.
6330                  .GoToReport_arw_incexp_opts_c_img.Visible = False
6340                  .GoToReport_arw_incexp_opts_cr_img.Visible = False
6350                  lngBtn_Top = .cmdMoreOptions.Top
6360                  lngBtn_Left = ((.cmdMoreOptions.Left + .cmdMoreOptions.Width) - lngTpp)
6370                  .GoToReport_arw_incexp_opts_o_img.Top = lngBtn_Top
6380                  .GoToReport_arw_incexp_opts_o_img.Left = lngBtn_Left
6390                  .GoToReport_arw_incexp_opts_o_img.Visible = True
6400                Case False
6410                  .opgSummary.SetFocus
6420                  .GoToReport_arw_incexp_opts_c_img.Visible = False
6430                  .GoToReport_arw_incexp_opts_cr_img.Visible = False
6440                  .GoToReport_arw_incexp_opts_o_img.Visible = False
6450                  .GoToReport_arw_incexp_sum_img.Visible = True
6460                  DoEvents
6470                  If .opgSummary <> .opgSummary_optOnly.OptionValue Then
6480                    .opgSummary = .opgSummary_optOnly.OptionValue
6490                    opgSummary_AfterUpdate  ' ** Procedure: Below.
6500                  End If
6510                End Select
6520                DoEvents
6530                .TimerInterval = 50&
6540              End Select
6550            End Select
6560          Case False
6570            DoCmd.Hourglass True  ' ** Make sure it's still running.
6580            DoEvents
6590            blnGoingToReport = True
6600            Select Case garr_varGoToReport(GTR_RNAM)
                Case "rptIncExp_Income"
6610              If .chkDetail = True Then
6620                .chkDetail = False
6630                chkDetail_AfterUpdate  ' ** Procedure: Below.
6640                DoEvents
6650              End If
6660              .cmdRevIncExp_IncomePrint.SetFocus
6670              .cmdRevIncExp_IncomeWord.Visible = False    'INC
6680              .GoToReport_arw_rptinc_img.Visible = True
6690              Beep
6700              DoCmd.Hourglass False
6710              DoEvents
6720              .TimerInterval = (GTR_WAIT * 2&)
6730            Case "rptIncExp_Expenses"
6740              If .chkDetail = True Then
6750                .chkDetail = False
6760                chkDetail_AfterUpdate  ' ** Procedure: Below.
6770                DoEvents
6780              End If
6790              .cmdRevIncExp_ExpensePrint.SetFocus
6800              .cmdRevIncExp_ExpenseWord.Visible = False   'EXP
6810              .GoToReport_arw_rptexp_img.Visible = True
6820              Beep
6830              DoCmd.Hourglass False
6840              DoEvents
6850              .TimerInterval = (GTR_WAIT * 2&)
6860            Case "rptIncExp_Income_Detailed", "rptIncExp_Expenses_Detailed"
6870              .chkDetail.SetFocus
6880              If .chkDetail = False Then
6890                .chkDetail = True
6900                chkDetail_AfterUpdate  ' ** Procedure: Below.
6910                DoEvents
6920              End If
6930              .GoToReport_arw_incexp_det_img.Visible = True
6940              DoEvents
6950              .TimerInterval = 50&
6960            Case "rptIncExp_Income_Summary", "rptIncExp_Expenses_Summary"
6970              If .opgAccountNumber <> .opgAccountNumber_optAll.OptionValue Then
6980                .opgAccountNumber = .opgAccountNumber_optAll.OptionValue
6990                opgAccountNumber_AfterUpdate  ' ** Procedure: Below.
7000                DoEvents
7010              End If
7020              If .chkDetail = True Then
7030                .chkDetail = False
7040                chkDetail_AfterUpdate  ' ** Procedure: Below.
7050                DoEvents
7060              End If
7070              .cmdMoreOptions.SetFocus
7080              Select Case .chkOptionsOpen
                  Case True
7090                lngBtn_Top = .cmdMoreOptions.Top
7100                lngBtn_Left = ((.cmdMoreOptions.Left + .cmdMoreOptions.Width) - lngTpp)
7110                .GoToReport_arw_incexp_opts_o_img.Top = lngBtn_Top
7120                .GoToReport_arw_incexp_opts_o_img.Left = lngBtn_Left
7130                .GoToReport_arw_incexp_opts_o_img.Visible = True
7140                blnOptsOpened = False
7150              Case False
                    ' ** The Form_Load() call is because the options are visible, but chkOptionsOpen defaults to False.
                    ' ** It just makes the visibility agree with the default. If it were True, a preference put it there.
7160                lngBtn_Top = .cmdMoreOptions.Top
7170                lngBtn_Left = ((.cmdMoreOptions.Left + .cmdMoreOptions.Width) - lngTpp)
7180                .GoToReport_arw_incexp_opts_c_img.Top = lngBtn_Top
7190                .GoToReport_arw_incexp_opts_cr_img.Top = lngBtn_Top
7200                .GoToReport_arw_incexp_opts_c_img.Left = lngBtn_Left
7210                .GoToReport_arw_incexp_opts_cr_img.Left = lngBtn_Left
7220                Select Case .opgOptionGroups_box4.Visible
                    Case True
7230                  .GoToReport_arw_incexp_opts_cr_img.Visible = True
7240                Case False
7250                  .GoToReport_arw_incexp_opts_c_img.Visible = True
7260                End Select
7270                blnOptsOpened = True
7280              End Select
7290              DoEvents
7300              .TimerInterval = 50&
7310            End Select
7320          End Select
7330        End If
7340      Else
7350        DoCmd.Hourglass False
7360      End If
7370    End With

EXITP:
7380    Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

7410    Select Case DataErr
        Case 2110  ' ** Access can't move the focus to the control '|'.
7420      Response = acDataErrContinue
7430    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
7440      If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateStart, DateEnd.
7450        MsgBox "You must enter the date in the format MM/DD/YYYY.", vbInformation + vbOKOnly, "Invalid Date"
7460        Response = acDataErrContinue
7470      End If
7480    Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
7490      Beep
7500      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
7510      Response = acDataErrContinue
7520    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
7530    Case Else
7540      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
7550      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
7560    End Select

EXITP:
7570    Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

7610    DoCmd.Hourglass True
7620    DoEvents

7630    With Me
7640      If .chkRememberMe = False Then
7650        .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
7660      End If
7670      If .chkRememberDates = False Then
7680        .DateStart = Null
7690        .DateEnd = Null
7700      End If
7710    End With

7720    Select Case blnWasGTR
        Case True
7730      blnWasGTR = False
7740    Case False
          ' ** Don't save settings imposed by GoToReport!
7750      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
7760    End Select

        ' ** Make sure all reports are closed.
7770    Do While Reports.Count > 0
7780      DoCmd.Close acReport, Reports(0).Name
7790    Loop

7800    gstrAccountNo = vbNullString
7810    gstrAccountName = vbNullString
7820    gdatStartDate = 0
7830    gdatEndDate = 0
7840    gstrCrtRpt_Ordinal = vbNullString  ' ** Borrowing this variable from Court Reports.

7850    gstrFormQuerySpec = vbNullString
7860    gblnSetFocus = True
7870    gstrReturningForm = THIS_NAME
7880    Select Case strCallingForm
        Case "frmMenu_Report"
7890      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
7900    End Select

EXITP:
7910    Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

8010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
8020    Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

8110    Detail_Mouse_IE blnCalendar1_Focus, blnCalendar2_Focus, blnMoreOptions_Focus, _
          blnReset_Focus, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
8120    Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

8210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8220    Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdClose_Click()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

8310    DoCmd.Hourglass True
8320    DoEvents

8330    DoCmd.Close acForm, THIS_NAME

EXITP:
8340    Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
8410    intShiftDown = (Shift And acShiftMask) > 0
8420    intAltDown = (Shift And acAltMask) > 0
8430    intCtrlDown = (Shift And acCtrlMask) > 0

8440    intRetVal = KeyCode

        ' ** Plain keys.
8450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8460      Select Case intRetVal
          Case vbKeyTab
8470        With Me
8480          intRetVal = 0
8490          .DateStart.SetFocus
8500        End With
8510      Case vbKeyUp, vbKeyLeft
8520        With Me
8530          intRetVal = 0
8540          .cmdPrintAll.SetFocus
8550        End With
8560      Case vbKeyDown, vbKeyRight
8570        With Me
8580          intRetVal = 0
8590          .DateStart.SetFocus
8600        End With
8610      End Select
8620    End If

        ' ** Shift keys.
8630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8640      Select Case intRetVal
          Case vbKeyTab
8650        With Me
8660          intRetVal = 0
8670          If .cmdExcelAll.Enabled = True Then
8680            .cmdExcelAll.SetFocus
8690          Else
8700            .cmdWordAll.SetFocus
8710          End If
8720        End With
8730      End Select
8740    End If

EXITP:
8750    KeyCode = intRetVal
8760    Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

8810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8820    Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
8910    intShiftDown = (Shift And acShiftMask) > 0
8920    intAltDown = (Shift And acAltMask) > 0
8930    intCtrlDown = (Shift And acCtrlMask) > 0

8940    intRetVal = KeyCode

        ' ** Plain keys.
8950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8970        With Me
8980          intRetVal = 0
8990          .DateEnd.SetFocus
9000        End With
9010      Case vbKeyUp
9020        With Me
9030          intRetVal = 0
9040          .cmdClose.SetFocus
9050        End With
9060      Case vbKeyDown
9070        With Me
9080          intRetVal = 0
9090          .DateEnd.SetFocus
9100        End With
9110      End Select
9120    End If

        ' ** Shift keys.
9130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9150        With Me
9160          intRetVal = 0
9170          .cmdClose.SetFocus
9180        End With
9190      End Select
9200    End If

        ' ** Ctrl keys.
9210    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
9220      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9230        With Me
9240          intRetVal = 0
9250          .cmdCalendar1.SetFocus
9260        End With
9270      End Select
9280    End If

EXITP:
9290    KeyCode = intRetVal
9300    Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub DateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_MouseMove"

9410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9420    Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub DateStart_Exit(Cancel As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

9510    With Me
9520      If IsDate(.DateEnd) = True And IsDate(.DateStart) = True Then
9530        If CDate(.DateStart) > CDate(.DateEnd) Then
9540          MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
9550          .DateEnd = Null
9560          .DateEnd.SetFocus
9570        End If
9580      End If
9590    End With

EXITP:
9600    Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

9710    Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
9720    Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

9810    Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
9820    Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

9910    Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
9920    Exit Sub

ERRH:
630     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
640     Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
10010   intShiftDown = (Shift And acShiftMask) > 0
10020   intAltDown = (Shift And acAltMask) > 0
10030   intCtrlDown = (Shift And acCtrlMask) > 0

10040   intRetVal = KeyCode

        ' ** Plain keys.
10050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10070       With Me
10080         intRetVal = 0
10090         .DateEnd.SetFocus
10100       End With
10110     End Select
10120   End If

        ' ** Shift keys.
10130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10150       With Me
10160         intRetVal = 0
10170         .DateStart.SetFocus
10180       End With
10190     End Select
10200   End If

EXITP:
10210   KeyCode = intRetVal
10220   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

10310   Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
10320   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

10410   Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
10420   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

10510   Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
10520   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
10610   intShiftDown = (Shift And acShiftMask) > 0
10620   intAltDown = (Shift And acAltMask) > 0
10630   intCtrlDown = (Shift And acCtrlMask) > 0

10640   intRetVal = KeyCode

        ' ** Plain keys.
10650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10670       With Me
10680         intRetVal = 0
10690         .opgAccountNumber.SetFocus
10700       End With
10710     Case vbKeyUp
10720       With Me
10730         intRetVal = 0
10740         .DateStart.SetFocus
10750       End With
10760     Case vbKeyDown
10770       With Me
10780         intRetVal = 0
10790         .opgAccountNumber.SetFocus
10800       End With
10810     End Select
10820   End If

        ' ** Shift keys.
10830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10850       With Me
10860         intRetVal = 0
10870         .DateStart.SetFocus
10880       End With
10890     End Select
10900   End If

        ' ** Ctrl keys.
10910   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10920     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10930       With Me
10940         intRetVal = 0
10950         .cmdCalendar2.SetFocus
10960       End With
10970     End Select
10980   End If

        ' ** Ctrl-Shift keys.
10990   If intCtrlDown And (Not intAltDown) And intShiftDown Then
11000     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11010       With Me
11020         intRetVal = 0
11030         .cmdCalendar1.SetFocus
11040       End With
11050     End Select
11060   End If

EXITP:
11070   KeyCode = intRetVal
11080   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub DateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_MouseMove"

11110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11120   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Public Sub DateEnd_Exit(Cancel As Integer)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

11210   With Me
11220     If IsNull(.DateEnd) Then
            ' ** Populate it with today's date.
11230       .DateEnd = Date
11240     Else
11250       If .DateEnd = "" Then
              ' ** Populate it with today's date.
11260         .DateEnd = Date
11270       Else
11280         If IsNull(.DateStart) = False Then
11290           If CDate(.DateStart) > CDate(.DateEnd) Then
11300             MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
11310             .DateStart = Null
11320             .DateStart.SetFocus
11330           End If
11340         End If
11350       End If
11360     End If
11370   End With

EXITP:
11380   Exit Sub

ERRH:
630     THAT_PROC = THIS_PROC
640     That_Erl = Erl
650     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
660     Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

11410   Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
11420   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

11510   Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
11520   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

11610   Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
11620   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
640     Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
11710   intShiftDown = (Shift And acShiftMask) > 0
11720   intAltDown = (Shift And acAltMask) > 0
11730   intCtrlDown = (Shift And acCtrlMask) > 0

11740   intRetVal = KeyCode

        ' ** Plain keys.
11750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11770       With Me
11780         intRetVal = 0
11790         .chkRememberDates.SetFocus
11800       End With
11810     End Select
11820   End If

        ' ** Shift keys.
11830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11850       With Me
11860         intRetVal = 0
11870         .DateEnd.SetFocus
11880       End With
11890     End Select
11900   End If

EXITP:
11910   KeyCode = intRetVal
11920   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

12010   Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
12020   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

12110   Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
12120   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

12210   Calendar_Handler_IE THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
12220   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Public Sub chkRememberDates_AfterUpdate()

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

12310   Remember_After_ID THIS_PROC, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
12320   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
12410   intShiftDown = (Shift And acShiftMask) > 0
12420   intAltDown = (Shift And acAltMask) > 0
12430   intCtrlDown = (Shift And acCtrlMask) > 0

12440   intRetVal = KeyCode

        ' ** Plain keys.
12450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12470       With Me
12480         intRetVal = 0
12490         .opgAccountNumber.SetFocus
12500       End With
12510     End Select
12520   End If

        ' ** Shift keys.
12530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12550       With Me
12560         intRetVal = 0
12570         .cmdCalendar2.SetFocus
12580       End With
12590     End Select
12600   End If

EXITP:
12610   KeyCode = intRetVal
12620   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

12710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12720   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Public Sub opgAccountNumber_AfterUpdate()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_AfterUpdate"

12810   AcctNum_After_IE Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
12820   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub opgAccountNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
12910   intShiftDown = (Shift And acShiftMask) > 0
12920   intAltDown = (Shift And acAltMask) > 0
12930   intCtrlDown = (Shift And acCtrlMask) > 0

12940   intRetVal = KeyCode

        ' ** Plain keys.
12950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12970       With Me
12980         intRetVal = 0
12990         If .cmbAccounts.Enabled = True Then
13000           .cmbAccounts.SetFocus
13010         Else
13020           .cmdRevIncExp_IncomePreview.SetFocus
13030         End If
13040       End With
13050     Case vbKeyUp, vbKeyLeft
13060       With Me
13070         intRetVal = 0
13080         .DateEnd.SetFocus
13090       End With
13100     End Select
13110   End If

        ' ** Shift keys.
13120   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13130     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13140       With Me
13150         intRetVal = 0
13160         .DateEnd.SetFocus
13170       End With
13180     End Select
13190   End If

        ' ** Ctrl keys.
13200   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
13210     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13220       With Me
13230         intRetVal = 0
13240         If .chkIncludeArchive.Enabled = True Then
13250           .chkIncludeArchive.SetFocus
13260         Else
13270           .chkDetail.SetFocus
13280         End If
13290       End With
13300     End Select
13310   End If

        ' ** Ctrl-Shift keys.
13320   If intCtrlDown And (Not intAltDown) And intShiftDown Then
13330     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13340       With Me
13350         intRetVal = 0
13360         .chkRememberDates.SetFocus
13370       End With
13380     End Select
13390   End If

EXITP:
13400   KeyCode = intRetVal
13410   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub opgAccountNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
13510   intShiftDown = (Shift And acShiftMask) > 0
13520   intAltDown = (Shift And acAltMask) > 0
13530   intCtrlDown = (Shift And acCtrlMask) > 0

13540   intRetVal = KeyCode

        ' ** Plain keys.
13550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13570       With Me
13580         intRetVal = 0
13590         If .cmbAccounts.Enabled = True Then
13600           .cmbAccounts.SetFocus
13610         Else
13620           .cmdRevIncExp_IncomePreview.SetFocus
13630         End If
13640       End With
13650     Case vbKeyDown, vbKeyRight
13660       With Me
13670         intRetVal = 0
13680         If .chkIncludeArchive.Enabled = True Then
13690           .chkIncludeArchive.SetFocus
13700         Else
13710           .chkDetail.SetFocus
13720         End If
13730       End With
13740     End Select
13750   End If

        ' ** Shift keys.
13760   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13770     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13780       With Me
13790         intRetVal = 0
13800         .DateEnd.SetFocus
13810       End With
13820     End Select
13830   End If

        ' ** Ctrl keys.
13840   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
13850     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13860       With Me
13870         intRetVal = 0
13880         If .chkIncludeArchive.Enabled = True Then
13890           .chkIncludeArchive.SetFocus
13900         Else
13910           .chkDetail.SetFocus
13920         End If
13930       End With
13940     End Select
13950   End If

        ' ** Ctrl-Shift keys.
13960   If intCtrlDown And (Not intAltDown) And intShiftDown Then
13970     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13980       With Me
13990         intRetVal = 0
14000         .chkRememberDates.SetFocus
14010       End With
14020     End Select
14030   End If

EXITP:
14040   KeyCode = intRetVal
14050   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub chkIncludeArchive_AfterUpdate()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

14110   With Me
14120     Select Case .chkIncludeArchive
          Case True
14130       .chkIncludeArchive_lbl.FontBold = True
14140     Case False
14150       .chkIncludeArchive_lbl.FontBold = False
14160     End Select
14170   End With

EXITP:
14180   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub chkIncludeArchive_KeyDown(KeyCode As Integer, Shift As Integer)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
14210   intShiftDown = (Shift And acShiftMask) > 0
14220   intAltDown = (Shift And acAltMask) > 0
14230   intCtrlDown = (Shift And acCtrlMask) > 0

14240   intRetVal = KeyCode

        ' ** Plain keys.
14250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14270       With Me
14280         intRetVal = 0
14290         .chkDetail.SetFocus
14300       End With
14310     Case vbKeyUp
14320       With Me
14330         intRetVal = 0
14340         .opgAccountNumber.SetFocus
14350       End With
14360     Case vbKeyLeft
14370       With Me
14380         intRetVal = 0
14390         .opgAccountNumber.SetFocus
14400       End With
14410     Case vbKeyRight
14420       With Me
14430         intRetVal = 0
14440         .chkDetail.SetFocus
14450       End With
14460     Case vbKeyDown
14470       With Me
14480         intRetVal = 0
14490         .chkDetail.SetFocus
14500       End With
14510     End Select
14520   End If

        ' ** Shift keys.
14530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14550       With Me
14560         intRetVal = 0
14570         .opgAccountNumber.SetFocus
14580       End With
14590     End Select
14600   End If

EXITP:
14610   KeyCode = intRetVal
14620   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub chkDetail_AfterUpdate()
' ** Public for GoToReport.

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkDetail_AfterUpdate"

14710   With Me
14720     Select Case .chkDetail
          Case True
14730       .chkDetail_lbl.FontBold = True
14740     Case False
14750       .chkDetail_lbl.FontBold = False
14760     End Select
14770   End With

EXITP:
14780   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub chkDetail_KeyDown(KeyCode As Integer, Shift As Integer)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkDetail_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
14810   intShiftDown = (Shift And acShiftMask) > 0
14820   intAltDown = (Shift And acAltMask) > 0
14830   intCtrlDown = (Shift And acCtrlMask) > 0

14840   intRetVal = KeyCode

        ' ** Plain keys.
14850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14870       With Me
14880         intRetVal = 0
14890         .cmdMoreOptions.SetFocus
14900       End With
14910     Case vbKeyUp, vbKeyLeft
14920       With Me
14930         intRetVal = 0
14940         If .chkIncludeArchive.Enabled = True Then
14950           .chkIncludeArchive.SetFocus
14960         Else
14970           .opgAccountNumber.SetFocus
14980         End If
14990       End With
15000     Case vbKeyRight
15010       With Me
15020         intRetVal = 0
15030         .cmdMoreOptions.SetFocus
15040       End With
15050     Case vbKeyDown
15060       With Me
15070         intRetVal = 0
15080         .cmdMoreOptions.SetFocus
15090       End With
15100     End Select
15110   End If

        ' ** Shift keys.
15120   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15130     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15140       With Me
15150         intRetVal = 0
15160         If .chkIncludeArchive.Enabled = True Then
15170           .chkIncludeArchive.SetFocus
15180         Else
15190           .opgAccountNumber.SetFocus
15200         End If
15210       End With
15220     End Select
15230   End If

EXITP:
15240   KeyCode = intRetVal
15250   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub opgOptionGroups_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOptionGroups_box_MouseMove"

15310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15320   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub opgOptionGroups_box2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOptionGroups_box2_MouseMove"

15410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15420   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdMoreOptions_GotFocus()

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoreOptions_GotFocus"

15510   OptionHandler_IE 1, blnMoreOptions_Focus, blnMoreOptions_MouseDown, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
15520   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdMoreOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoreOptions_MouseDown"

15610   OptionHandler_IE 2, blnMoreOptions_Focus, blnMoreOptions_MouseDown, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
15620   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdMoreOptions_Click()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoreOptions_Click"

15710   With Me

15720     Select Case .chkOptionsOpen
          Case True
            ' ** Then hide it.
15730       ShowOptions_IE "Hide", lngDetail_Height, lngDiff, lngIncBtn_Top, lngExpBtn_Top, lngBtnLbl_Diff, _
              lngDetailLine1_Top, lngAccountNo_Top, lngAccountNoBox_Top, lngAccountNoLbl_Top, lngRemMe_Top, _
              lngAcctSrc_Top, lngAcctSrcOpt_Top, lngAcctSrcOptLbl_Top, lngPageOf_Top, lngAllBtn_Top, _
              lngHighlight_Offset, lngDetailLine3_Top, Me   ' ** Module Procedure: modIncExpFuncs.
15740       DoEvents
15750       OptionsChk_IE Me  ' ** Module Procedure: modIncExpFuncs.
15760     Case False
            ' ** Then show it.
15770       ShowOptions_IE "Show", lngDetail_Height, lngDiff, lngIncBtn_Top, lngExpBtn_Top, lngBtnLbl_Diff, _
              lngDetailLine1_Top, lngAccountNo_Top, lngAccountNoBox_Top, lngAccountNoLbl_Top, lngRemMe_Top, _
              lngAcctSrc_Top, lngAcctSrcOpt_Top, lngAcctSrcOptLbl_Top, lngPageOf_Top, lngAllBtn_Top, _
              lngHighlight_Offset, lngDetailLine3_Top, Me  ' ** Module Procedure: modIncExpFuncs.
15780       DoEvents
15790       OptionsChk_IE Me  ' ** Module Procedure: modIncExpFuncs.
15800     End Select

15810   End With

EXITP:
15820   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdMoreOptions_KeyDown(KeyCode As Integer, Shift As Integer)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoreOptions_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
15910   intShiftDown = (Shift And acShiftMask) > 0
15920   intAltDown = (Shift And acAltMask) > 0
15930   intCtrlDown = (Shift And acCtrlMask) > 0

15940   intRetVal = KeyCode

        ' ** Plain keys.
15950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15960     Select Case intRetVal
          Case vbKeyTab
15970       With Me
15980         intRetVal = 0
15990         Select Case .opgOptionGroups_box.Visible
              Case True
16000           .opgUnspecified.SetFocus
16010         Case False
16020           .cmbAccounts.SetFocus
16030         End Select
16040       End With
16050     End Select
16060   End If

        ' ** Shift keys.
16070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16080     Select Case intRetVal
          Case vbKeyTab
16090       With Me
16100         intRetVal = 0
16110         .opgAccountNumber.SetFocus
16120       End With
16130     End Select
16140   End If

        ' ** Ctrl keys.
16150   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
16160     Select Case intRetVal
          Case vbKeyTab
16170       With Me
16180         intRetVal = 0
16190         If .opgOptionGroups_box.Visible = False And .cmbAccounts.Enabled = True Then
16200           .opgAccountSource.SetFocus
16210         End If
16220       End With
16230     End Select
16240   End If

        ' ** Ctrl-Shift keys.
16250   If intCtrlDown And (Not intAltDown) And intShiftDown Then
16260     Select Case intRetVal
          Case vbKeyTab
16270       With Me
16280         intRetVal = 0
16290         .chkDetail.SetFocus
16300       End With
16310     End Select
16320   End If

EXITP:
16330   KeyCode = intRetVal
16340   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmdMoreOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoreOptions_MouseMove"

16410   OptionHandler_IE 3, blnMoreOptions_Focus, blnMoreOptions_MouseDown, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
16420   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdMoreOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoreOptions_MouseUp"

16510   OptionHandler_IE 4, blnMoreOptions_Focus, blnMoreOptions_MouseDown, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
16520   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdMoreOptions_LostFocus()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoreOptions_LostFocus"

16610   OptionHandler_IE 5, blnMoreOptions_Focus, blnMoreOptions_MouseDown, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
16620   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Public Sub opgUnspecified_AfterUpdate()

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnspecified_AfterUpdate"

16710   OptGroup_Handler_IE THIS_PROC, Me  ' ** Module Procedure.  ' ** Module Procedure: modIncExpFuncs.

EXITP:
16720   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub opgUnspecified_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnspecified_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
16810   intShiftDown = (Shift And acShiftMask) > 0
16820   intAltDown = (Shift And acAltMask) > 0
16830   intCtrlDown = (Shift And acCtrlMask) > 0

16840   intRetVal = KeyCode

        ' ** Plain keys.
16850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16870       With Me
16880         intRetVal = 0
16890         .opgPrincipalCash.SetFocus
16900       End With
16910     Case vbKeyUp, vbKeyLeft
16920       With Me
16930         intRetVal = 0
16940         .cmdMoreOptions.SetFocus
16950       End With
16960     Case vbKeyDown
16970       With Me
16980         intRetVal = 0
16990         .opgPrincipalCash.SetFocus
17000       End With
17010     End Select
17020   End If

        ' ** Shift keys.
17030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17050       With Me
17060         intRetVal = 0
17070         .cmdMoreOptions.SetFocus
17080       End With
17090     End Select
17100   End If

EXITP:
17110   KeyCode = intRetVal
17120   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub opgUnspecified_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnspecified_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
17210   intShiftDown = (Shift And acShiftMask) > 0
17220   intAltDown = (Shift And acAltMask) > 0
17230   intCtrlDown = (Shift And acCtrlMask) > 0

17240   intRetVal = KeyCode

        ' ** Plain keys.
17250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17270       With Me
17280         intRetVal = 0
17290         .opgPrincipalCash.SetFocus
17300       End With
17310     Case vbKeyUp
17320       With Me
17330         intRetVal = 0
17340         .cmdMoreOptions.SetFocus
17350       End With
17360     Case vbKeyDown
17370       With Me
17380         intRetVal = 0
17390         .opgPrincipalCash.SetFocus
17400       End With
17410     End Select
17420   End If

        ' ** Shift keys.
17430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17450       With Me
17460         intRetVal = 0
17470         .cmdMoreOptions.SetFocus
17480       End With
17490     End Select
17500   End If

EXITP:
17510   KeyCode = intRetVal
17520   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub opgUnspecified_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnspecified_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
17610   intShiftDown = (Shift And acShiftMask) > 0
17620   intAltDown = (Shift And acAltMask) > 0
17630   intCtrlDown = (Shift And acCtrlMask) > 0

17640   intRetVal = KeyCode

        ' ** Plain keys.
17650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17670       With Me
17680         intRetVal = 0
17690         .opgPrincipalCash.SetFocus
17700       End With
17710     Case vbKeyUp
17720       With Me
17730         intRetVal = 0
17740         .cmdMoreOptions.SetFocus
17750       End With
17760     Case vbKeyRight
17770       With Me
17780         intRetVal = 0
17790         .chkDontCombineMulti.SetFocus
17800       End With
17810     Case vbKeyDown
17820       With Me
17830         intRetVal = 0
17840         .opgPrincipalCash.SetFocus
17850       End With
17860     End Select
17870   End If

        ' ** Shift keys.
17880   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17890     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17900       With Me
17910         intRetVal = 0
17920         .cmdMoreOptions.SetFocus
17930       End With
17940     End Select
17950   End If

EXITP:
17960   KeyCode = intRetVal
17970   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub opgPrincipalCash_AfterUpdate()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrincipalCash_AfterUpdate"

18010   OptGroup_Handler_IE THIS_PROC, Me  ' ** Module Procedure.  ' ** Module Procedure: modIncExpFuncs.

EXITP:
18020   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub opgPrincipalCash_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrincipalCash_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
18110   intShiftDown = (Shift And acShiftMask) > 0
18120   intAltDown = (Shift And acAltMask) > 0
18130   intCtrlDown = (Shift And acCtrlMask) > 0

18140   intRetVal = KeyCode

        ' ** Plain keys.
18150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18170       With Me
18180         intRetVal = 0
18190         .opgZeroCash.SetFocus
18200       End With
18210     Case vbKeyLeft
18220       With Me
18230         intRetVal = 0
18240         .chkDontCombineMulti.SetFocus
18250       End With
18260     Case vbKeyUp
18270       With Me
18280         intRetVal = 0
18290         .opgUnspecified.SetFocus
18300       End With
18310     Case vbKeyDown
18320       With Me
18330         intRetVal = 0
18340         .opgZeroCash.SetFocus
18350       End With
18360     End Select
18370   End If

        ' ** Shift keys.
18380   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18390     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18400       With Me
18410         intRetVal = 0
18420         .opgUnspecified.SetFocus
18430       End With
18440     End Select
18450   End If

EXITP:
18460   KeyCode = intRetVal
18470   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub opgPrincipalCash_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrincipalCash_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
18510   intShiftDown = (Shift And acShiftMask) > 0
18520   intAltDown = (Shift And acAltMask) > 0
18530   intCtrlDown = (Shift And acCtrlMask) > 0

18540   intRetVal = KeyCode

        ' ** Plain keys.
18550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18570       With Me
18580         intRetVal = 0
18590         .opgZeroCash.SetFocus
18600       End With
18610     Case vbKeyUp
18620       With Me
18630         intRetVal = 0
18640         .opgUnspecified.SetFocus
18650       End With
18660     Case vbKeyRight
18670       With Me
18680         If .opgPrincipalCash_optOnly.Visible = False Or .opgPrincipalCash_optOnly.Enabled = False Then
18690           intRetVal = 0
18700           .chkAcctEveryLine.SetFocus
18710         End If
18720       End With
18730     Case vbKeyDown
18740       With Me
18750         intRetVal = 0
18760         .opgZeroCash.SetFocus
18770       End With
18780     End Select
18790   End If

        ' ** Shift keys.
18800   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18810     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18820       With Me
18830         intRetVal = 0
18840         .opgUnspecified.SetFocus
18850       End With
18860     End Select
18870   End If

EXITP:
18880   KeyCode = intRetVal
18890   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub opgPrincipalCash_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrincipalCash_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
18910   intShiftDown = (Shift And acShiftMask) > 0
18920   intAltDown = (Shift And acAltMask) > 0
18930   intCtrlDown = (Shift And acCtrlMask) > 0

18940   intRetVal = KeyCode

        ' ** Plain keys.
18950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18970       With Me
18980         intRetVal = 0
18990         .opgZeroCash.SetFocus
19000       End With
19010     Case vbKeyUp
19020       With Me
19030         intRetVal = 0
19040         .opgUnspecified.SetFocus
19050       End With
19060     Case vbKeyRight
19070       With Me
19080         intRetVal = 0
19090         .chkAcctEveryLine.SetFocus
19100       End With
19110     Case vbKeyDown
19120       With Me
19130         intRetVal = 0
19140         .opgZeroCash.SetFocus
19150       End With
19160     End Select
19170   End If

        ' ** Shift keys.
19180   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19190     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19200       With Me
19210         intRetVal = 0
19220         .opgUnspecified.SetFocus
19230       End With
19240     End Select
19250   End If

EXITP:
19260   KeyCode = intRetVal
19270   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub opgZeroCash_AfterUpdate()

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgZeroCash_AfterUpdate"

19310   OptGroup_Handler_IE THIS_PROC, Me  ' ** Module Procedure.  ' ** Module Procedure: modIncExpFuncs.

EXITP:
19320   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub opgZeroCash_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgZeroCash_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
19410   intShiftDown = (Shift And acShiftMask) > 0
19420   intAltDown = (Shift And acAltMask) > 0
19430   intCtrlDown = (Shift And acCtrlMask) > 0

19440   intRetVal = KeyCode

        ' ** Plain keys.
19450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19470       With Me
19480         intRetVal = 0
19490         .opgSummary.SetFocus
19500       End With
19510     Case vbKeyLeft
19520       With Me
19530         intRetVal = 0
19540         .chkAcctEveryLine.SetFocus
19550       End With
19560     Case vbKeyUp
19570       With Me
19580         intRetVal = 0
19590         .opgPrincipalCash.SetFocus
19600       End With
19610     Case vbKeyDown
19620       With Me
19630         intRetVal = 0
19640         .opgSummary.SetFocus
19650       End With
19660     End Select
19670   End If

        ' ** Shift keys.
19680   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19690     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19700       With Me
19710         intRetVal = 0
19720         .opgPrincipalCash.SetFocus
19730       End With
19740     End Select
19750   End If

EXITP:
19760   KeyCode = intRetVal
19770   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub opgZeroCash_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgZeroCash_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
19810   intShiftDown = (Shift And acShiftMask) > 0
19820   intAltDown = (Shift And acAltMask) > 0
19830   intCtrlDown = (Shift And acCtrlMask) > 0

19840   intRetVal = KeyCode

        ' ** Plain keys.
19850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19870       With Me
19880         intRetVal = 0
19890         .opgSummary.SetFocus
19900       End With
19910     Case vbKeyUp
19920       With Me
19930         intRetVal = 0
19940         .opgPrincipalCash.SetFocus
19950       End With
19960     Case vbKeyRight
19970       With Me
19980         If .opgZeroCash_optOnly.Visible = False Or .opgZeroCash_optOnly.Enabled = False Then
19990           intRetVal = 0
20000           .chkSweepOnly.SetFocus
20010         End If
20020       End With
20030     Case vbKeyDown
20040       With Me
20050         intRetVal = 0
20060         .opgSummary.SetFocus
20070       End With
20080     End Select
20090   End If

        ' ** Shift keys.
20100   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20110     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20120       With Me
20130         intRetVal = 0
20140         .opgPrincipalCash.SetFocus
20150       End With
20160     End Select
20170   End If

EXITP:
20180   KeyCode = intRetVal
20190   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub opgZeroCash_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgZeroCash_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
20210   intShiftDown = (Shift And acShiftMask) > 0
20220   intAltDown = (Shift And acAltMask) > 0
20230   intCtrlDown = (Shift And acCtrlMask) > 0

20240   intRetVal = KeyCode

        ' ** Plain keys.
20250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20270       With Me
20280         intRetVal = 0
20290         .opgSummary.SetFocus
20300       End With
20310     Case vbKeyUp
20320       With Me
20330         intRetVal = 0
20340         .opgPrincipalCash.SetFocus
20350       End With
20360     Case vbKeyRight
20370       With Me
20380         intRetVal = 0
20390         .chkSweepOnly.SetFocus
20400       End With
20410     Case vbKeyDown
20420       With Me
20430         intRetVal = 0
20440         .opgSummary.SetFocus
20450       End With
20460     End Select
20470   End If

        ' ** Shift keys.
20480   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20490     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20500       With Me
20510         intRetVal = 0
20520         .opgPrincipalCash.SetFocus
20530       End With
20540     End Select
20550   End If

EXITP:
20560   KeyCode = intRetVal
20570   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub opgSummary_AfterUpdate()

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgSummary_AfterUpdate"

20610   OptGroup_Handler_IE THIS_PROC, Me  ' ** Module Procedure.  ' ** Module Procedure: modIncExpFuncs.

EXITP:
20620   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub opgSummary_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgSummary_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
20710   intShiftDown = (Shift And acShiftMask) > 0
20720   intAltDown = (Shift And acAltMask) > 0
20730   intCtrlDown = (Shift And acCtrlMask) > 0

20740   intRetVal = KeyCode

        ' ** Plain keys.
20750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20770       With Me
20780         intRetVal = 0
20790         .chkDontCombineMulti.SetFocus
20800       End With
20810     Case vbKeyLeft
20820       With Me
20830         intRetVal = 0
20840         .chkSweepOnly.SetFocus
20850       End With
20860     Case vbKeyUp
20870       With Me
20880         intRetVal = 0
20890         .opgZeroCash.SetFocus
20900       End With
20910     Case vbKeyDown
20920       With Me
20930         intRetVal = 0
20940         If .opgAccountSource.Enabled = True Then
20950           .opgAccountSource.SetFocus
20960         Else
20970           .cmdRevIncExp_IncomePreview.SetFocus
20980         End If
20990       End With
21000     End Select
21010   End If

        ' ** Shift keys.
21020   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21030     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21040       With Me
21050         intRetVal = 0
21060         .opgZeroCash.SetFocus
21070       End With
21080     End Select
21090   End If

EXITP:
21100   KeyCode = intRetVal
21110   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub opgSummary_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgSummary_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
21210   intShiftDown = (Shift And acShiftMask) > 0
21220   intAltDown = (Shift And acAltMask) > 0
21230   intCtrlDown = (Shift And acCtrlMask) > 0

21240   intRetVal = KeyCode

        ' ** Plain keys.
21250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21270       With Me
21280         intRetVal = 0
21290         .chkDontCombineMulti.SetFocus
21300       End With
21310     Case vbKeyUp
21320       With Me
21330         intRetVal = 0
21340         .opgZeroCash.SetFocus
21350       End With
21360     Case vbKeyRight
21370       With Me
21380         If .opgSummary_optOnly.Enabled = False Then
21390           intRetVal = 0
21400           If .cmdReset.Enabled = True Then
21410             .cmdReset.SetFocus
21420           ElseIf .opgAccountSource.Enabled = True Then
21430             .opgAccountSource.SetFocus
21440           Else
21450             .cmdRevIncExp_IncomePreview.SetFocus
21460           End If
21470         End If
21480       End With
21490     Case vbKeyDown
21500       With Me
21510         intRetVal = 0
21520         If .opgAccountSource.Enabled = True Then
21530           .opgAccountSource.SetFocus
21540         Else
21550           .cmdRevIncExp_IncomePreview.SetFocus
21560         End If
21570       End With
21580     End Select
21590   End If

        ' ** Shift keys.
21600   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21610     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21620       With Me
21630         intRetVal = 0
21640         .opgZeroCash.SetFocus
21650       End With
21660     End Select
21670   End If

EXITP:
21680   KeyCode = intRetVal
21690   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub opgSummary_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgSummary_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
21710   intShiftDown = (Shift And acShiftMask) > 0
21720   intAltDown = (Shift And acAltMask) > 0
21730   intCtrlDown = (Shift And acCtrlMask) > 0

21740   intRetVal = KeyCode

        ' ** Plain keys.
21750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21770       With Me
21780         intRetVal = 0
21790         .chkDontCombineMulti.SetFocus
21800       End With
21810     Case vbKeyUp
21820       With Me
21830         intRetVal = 0
21840         .opgZeroCash.SetFocus
21850       End With
21860     Case vbKeyRight
21870       With Me
21880         intRetVal = 0
21890         If .cmdReset.Enabled = True Then
21900           .cmdReset.SetFocus
21910         ElseIf .opgAccountSource.Enabled = True Then
21920           .opgAccountSource.SetFocus
21930         Else
21940           .cmdRevIncExp_IncomePreview.SetFocus
21950         End If
21960       End With
21970     Case vbKeyDown
21980       With Me
21990         intRetVal = 0
22000         If .opgAccountSource.Enabled = True Then
22010           .opgAccountSource.SetFocus
22020         Else
22030           .cmdRevIncExp_IncomePreview.SetFocus
22040         End If
22050       End With
22060     End Select
22070   End If

        ' ** Shift keys.
22080   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22090     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22100       With Me
22110         intRetVal = 0
22120         .opgZeroCash.SetFocus
22130       End With
22140     End Select
22150   End If

EXITP:
22160   KeyCode = intRetVal
22170   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub chkDontCombineMulti_AfterUpdate()

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkDontCombineMulti_AfterUpdate"

22210   With Me
22220     Select Case .chkDontCombineMulti
          Case True
22230       .chkDontCombineMulti_lbl.FontBold = True
22240     Case False
22250       .chkDontCombineMulti_lbl.FontBold = False
22260     End Select
22270     OptionsChk_IE Me  ' ** Module Procedure: modIncExpFuncs.
22280   End With

EXITP:
22290   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub chkDontCombineMulti_KeyDown(KeyCode As Integer, Shift As Integer)

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkDontCombineMulti_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
22310   intShiftDown = (Shift And acShiftMask) > 0
22320   intAltDown = (Shift And acAltMask) > 0
22330   intCtrlDown = (Shift And acCtrlMask) > 0

22340   intRetVal = KeyCode

        ' ** Plain keys.
22350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22370       With Me
22380         intRetVal = 0
22390         .chkAcctEveryLine.SetFocus
22400       End With
22410     Case vbKeyUp
22420       With Me
22430         intRetVal = 0
22440         .opgSummary.SetFocus
22450       End With
22460     Case vbKeyLeft
22470       With Me
22480         intRetVal = 0
22490         .opgUnspecified.SetFocus
22500       End With
22510     Case vbKeyRight
22520       With Me
22530         intRetVal = 0
22540         .opgPrincipalCash.SetFocus
22550       End With
22560     Case vbKeyDown
22570       With Me
22580         intRetVal = 0
22590         .chkAcctEveryLine.SetFocus
22600       End With
22610     End Select
22620   End If

        ' ** Shift keys.
22630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22650       With Me
22660         intRetVal = 0
22670         .opgSummary.SetFocus
22680       End With
22690     End Select
22700   End If

EXITP:
22710   KeyCode = intRetVal
22720   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub chkAcctEveryLine_AfterUpdate()

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAcctEveryLine_AfterUpdate"

22810   With Me
22820     Select Case .chkAcctEveryLine
          Case True
22830       .chkAcctEveryLine_lbl.FontBold = True
22840     Case False
22850       .chkAcctEveryLine_lbl.FontBold = False
22860     End Select
22870     OptionsChk_IE Me  ' ** Module Procedure: modIncExpFuncs.
22880   End With

EXITP:
22890   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub chkAcctEveryLine_KeyDown(KeyCode As Integer, Shift As Integer)

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAcctEveryLine_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
22910   intShiftDown = (Shift And acShiftMask) > 0
22920   intAltDown = (Shift And acAltMask) > 0
22930   intCtrlDown = (Shift And acCtrlMask) > 0

22940   intRetVal = KeyCode

        ' ** Plain keys.
22950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22970       With Me
22980         intRetVal = 0
22990         .chkSweepOnly.SetFocus
23000       End With
23010     Case vbKeyUp
23020       With Me
23030         intRetVal = 0
23040         .chkDontCombineMulti.SetFocus
23050       End With
23060     Case vbKeyLeft
23070       With Me
23080         intRetVal = 0
23090         .opgPrincipalCash.SetFocus
23100       End With
23110     Case vbKeyRight
23120       With Me
23130         intRetVal = 0
23140         .opgZeroCash.SetFocus
23150       End With
23160     Case vbKeyDown
23170       With Me
23180         intRetVal = 0
23190         .chkSweepOnly.SetFocus
23200       End With
23210     End Select
23220   End If

        ' ** Shift keys.
23230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23250       With Me
23260         intRetVal = 0
23270         .chkDontCombineMulti.SetFocus
23280       End With
23290     End Select
23300   End If

EXITP:
23310   KeyCode = intRetVal
23320   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub chkSweepOnly_AfterUpdate()

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkSweepOnly_AfterUpdate"

23410   With Me
23420     Select Case .chkSweepOnly
          Case True
23430       .chkSweepOnly_lbl.FontBold = True
23440     Case False
23450       .chkSweepOnly_lbl.FontBold = False
23460     End Select
23470     OptionsChk_IE Me  ' ** Module Procedure: modIncExpFuncs.
23480   End With

EXITP:
23490   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub chkSweepOnly_KeyDown(KeyCode As Integer, Shift As Integer)

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkSweepOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
23510   intShiftDown = (Shift And acShiftMask) > 0
23520   intAltDown = (Shift And acAltMask) > 0
23530   intCtrlDown = (Shift And acCtrlMask) > 0

23540   intRetVal = KeyCode

        ' ** Plain keys.
23550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23570       With Me
23580         intRetVal = 0
23590         If .cmdReset.Enabled = True Then
23600           .cmdReset.SetFocus
23610         ElseIf .opgAccountSource.Enabled = True Then
23620           .opgAccountSource.SetFocus
23630         Else
23640           .cmdRevIncExp_IncomePreview.SetFocus
23650         End If
23660       End With
23670     Case vbKeyUp
23680       With Me
23690         intRetVal = 0
23700         .chkAcctEveryLine.SetFocus
23710       End With
23720     Case vbKeyLeft
23730       With Me
23740         intRetVal = 0
23750         .opgZeroCash.SetFocus
23760       End With
23770     Case vbKeyRight
23780       With Me
23790         intRetVal = 0
23800         .opgSummary.SetFocus
23810       End With
23820     Case vbKeyDown
23830       With Me
23840         intRetVal = 0
23850         If .cmdReset.Enabled = True Then
23860           .cmdReset.SetFocus
23870         ElseIf .opgAccountSource.Enabled = True Then
23880           .opgAccountSource.SetFocus
23890         Else
23900           .cmdRevIncExp_IncomePreview.SetFocus
23910         End If
23920       End With
23930     End Select
23940   End If

        ' ** Shift keys.
23950   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23970       With Me
23980         intRetVal = 0
23990         .chkAcctEveryLine.SetFocus
24000       End With
24010     End Select
24020   End If

        ' ** Ctrl keys.
24030   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
24040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24050       With Me
24060         intRetVal = 0
24070         If .cmdReset.Enabled = False And .cmbAccounts.Enabled = True Then
24080           .opgAccountSource.SetFocus
24090         End If
24100       End With
24110     End Select
24120   End If

EXITP:
24130   KeyCode = intRetVal
24140   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmdReset_GotFocus()

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_GotFocus"

24210   Reset_Handler_ID THIS_PROC, blnReset_Focus, blnReset_MouseDown, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
24220   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseDown"

24310   Reset_Handler_ID THIS_PROC, blnReset_Focus, blnReset_MouseDown, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
24320   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdReset_Click()

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_Click"

24410   Reset_Handler_ID THIS_PROC, blnReset_Focus, blnReset_MouseDown, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
24420   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdReset_KeyDown(KeyCode As Integer, Shift As Integer)

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
24510   intShiftDown = (Shift And acShiftMask) > 0
24520   intAltDown = (Shift And acAltMask) > 0
24530   intCtrlDown = (Shift And acCtrlMask) > 0

24540   intRetVal = KeyCode

        ' ** Plain keys.
24550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24560     Select Case intRetVal
          Case vbKeyTab
24570       With Me
24580         intRetVal = 0
24590         .cmbAccounts.SetFocus
24600       End With
24610     End Select
24620   End If

        ' ** Shift keys.
24630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24640     Select Case intRetVal
          Case vbKeyTab
24650       With Me
24660         intRetVal = 0
24670         .chkSweepOnly.SetFocus
24680       End With
24690     End Select
24700   End If

        ' ** Ctrl keys.
24710   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
24720     Select Case intRetVal
          Case vbKeyTab
24730       With Me
24740         intRetVal = 0
24750         .opgAccountSource.SetFocus
24760       End With
24770     End Select
24780   End If

EXITP:
24790   KeyCode = intRetVal
24800   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmdReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseMove"

24910   Reset_Handler_ID THIS_PROC, blnReset_Focus, blnReset_MouseDown, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
24920   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseUp"

25010   Reset_Handler_ID THIS_PROC, blnReset_Focus, blnReset_MouseDown, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
25020   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdReset_LostFocus()

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_LostFocus"

25110   Reset_Handler_ID THIS_PROC, blnReset_Focus, blnReset_MouseDown, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
25120   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Public Sub opgAccountSource_AfterUpdate()

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

25210   AcctSrc_After_IE Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
25220   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
25310   intShiftDown = (Shift And acShiftMask) > 0
25320   intAltDown = (Shift And acAltMask) > 0
25330   intCtrlDown = (Shift And acCtrlMask) > 0

25340   intRetVal = KeyCode

        ' ** Plain keys.
25350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25370       With Me
25380         intRetVal = 0
25390         .chkRememberMe.SetFocus
25400       End With
25410     Case vbKeyUp, vbKeyLeft
25420       With Me
25430         intRetVal = 0
25440         If .cmdReset.Visible = True And .cmdReset.Enabled = True Then
25450           .cmdReset.SetFocus
25460         ElseIf .opgSummary.Visible = True And .opgSummary.Enabled = True Then
25470           .opgSummary.SetFocus
25480         Else
25490           .chkDetail.SetFocus
25500         End If
25510       End With
25520     Case vbKeyDown
25530       With Me
25540         intRetVal = 0
25550         .chkRememberMe.SetFocus
25560       End With
25570     End Select
25580   End If

        ' ** Shift keys.
25590   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25600     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25610       With Me
25620         intRetVal = 0
25630         If .cmdReset.Visible = True And .cmdReset.Enabled = True Then
25640           .cmdReset.SetFocus
25650         ElseIf .chkSweepOnly.Visible = True And .chkSweepOnly.Enabled = True Then
25660           .chkSweepOnly.SetFocus
25670         Else
25680           .cmdMoreOptions.SetFocus
25690         End If
25700       End With
25710     End Select
25720   End If

EXITP:
25730   KeyCode = intRetVal
25740   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
25810   intShiftDown = (Shift And acShiftMask) > 0
25820   intAltDown = (Shift And acAltMask) > 0
25830   intCtrlDown = (Shift And acCtrlMask) > 0

25840   intRetVal = KeyCode

        ' ** Plain keys.
25850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25870       With Me
25880         intRetVal = 0
25890         .chkRememberMe.SetFocus
25900       End With
25910     Case vbKeyUp
25920       With Me
25930         intRetVal = 0
25940         If .cmdReset.Visible = True And .cmdReset.Enabled = True Then
25950           .cmdReset.SetFocus
25960         ElseIf .opgSummary.Visible = True And .opgSummary.Enabled = True Then
25970           .opgSummary.SetFocus
25980         Else
25990           .chkDetail.SetFocus
26000         End If
26010       End With
26020     Case vbKeyDown, vbKeyRight
26030       With Me
26040         intRetVal = 0
26050         .chkRememberMe.SetFocus
26060       End With
26070     End Select
26080   End If

        ' ** Shift keys.
26090   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26110       With Me
26120         intRetVal = 0
26130         If .cmdReset.Visible = True And .cmdReset.Enabled = True Then
26140           .cmdReset.SetFocus
26150         ElseIf .chkSweepOnly.Visible = True And .chkSweepOnly.Enabled = True Then
26160           .chkSweepOnly.SetFocus
26170         Else
26180           .cmdMoreOptions.SetFocus
26190         End If
26200       End With
26210     End Select
26220   End If

EXITP:
26230   KeyCode = intRetVal
26240   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub chkRememberMe_AfterUpdate()

26300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

26310   Remember_After_ID THIS_PROC, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
26320   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
26410   intShiftDown = (Shift And acShiftMask) > 0
26420   intAltDown = (Shift And acAltMask) > 0
26430   intCtrlDown = (Shift And acCtrlMask) > 0

26440   intRetVal = KeyCode

        ' ** Plain keys.
26450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26470       With Me
26480         intRetVal = 0
26490         .cmbAccounts.SetFocus
26500       End With
26510     Case vbKeyUp, vbKeyLeft
26520       With Me
26530         intRetVal = 0
26540         .opgAccountSource.SetFocus
26550       End With
26560     Case vbKeyDown
26570       With Me
26580         intRetVal = 0
26590         .cmbAccounts.SetFocus
26600       End With
26610     End Select
26620   End If

        ' ** Shift keys.
26630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26650       With Me
26660         intRetVal = 0
26670         .opgAccountSource.SetFocus
26680       End With
26690     End Select
26700   End If

EXITP:
26710   KeyCode = intRetVal
26720   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

26810   With Me
26820     If .cmbAccounts.text = vbNullString Then
26830       ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
26840       .cmbAccounts = .cmbAccounts.ItemData(0)
26850     End If
26860   End With

EXITP:
26870   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
26910   intShiftDown = (Shift And acShiftMask) > 0
26920   intAltDown = (Shift And acAltMask) > 0
26930   intCtrlDown = (Shift And acCtrlMask) > 0

26940   intRetVal = KeyCode

        ' ** Plain keys.
26950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26970       With Me
26980         intRetVal = 0
26990         .cmdRevIncExp_IncomePreview.SetFocus
27000       End With
27010     End Select
27020   End If

        ' ** Shift keys.
27030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27050       With Me
27060         intRetVal = 0
27070         .opgAccountNumber.SetFocus
27080       End With
27090     End Select
27100   End If

        ' ** Ctrl keys.
27110   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
27120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27130       With Me
27140         intRetVal = 0
27150         .chkPageOf.SetFocus
27160       End With
27170     End Select
27180   End If

        ' ** Ctrl-Shift keys.
27190   If intCtrlDown And (Not intAltDown) And intShiftDown Then
27200     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27210       With Me
27220         intRetVal = 0
27230         .chkRememberMe.SetFocus
27240       End With
27250     End Select
27260   End If

EXITP:
27270   KeyCode = intRetVal
27280   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub chkPageOf_AfterUpdate()

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

27310   With Me
27320     Select Case .chkPageOf
          Case True
27330       .chkPageOf_lbl.FontBold = True
27340     Case False
27350       .chkPageOf_lbl.FontBold = False
27360     End Select
27370   End With

EXITP:
27380   Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
27410   intShiftDown = (Shift And acShiftMask) > 0
27420   intAltDown = (Shift And acAltMask) > 0
27430   intCtrlDown = (Shift And acCtrlMask) > 0

27440   intRetVal = KeyCode

        ' ** Plain keys.
27450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27470       With Me
27480         intRetVal = 0
27490         .cmdRevIncExp_IncomePreview.SetFocus
27500       End With
27510     End Select
27520   End If

        ' ** Shift keys.
27530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27550       With Me
27560         intRetVal = 0
27570         If .cmbAccounts.Enabled = True Then
27580           .cmbAccounts.SetFocus
27590         ElseIf .cmdReset.Visible = True And .cmdReset.Enabled = True Then
27600           .cmdReset.SetFocus
27610         ElseIf .chkSweepOnly.Visible = True And .chkSweepOnly.Enabled = True Then
27620           .chkSweepOnly.SetFocus
27630         Else
27640           .cmdMoreOptions.SetFocus
27650         End If
27660       End With
27670     End Select
27680   End If

EXITP:
27690   KeyCode = intRetVal
27700   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmdRevIncExp_IncomePreview_Click()

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_IncomePreview_Click"

27810   Prev_Click_IE THIS_PROC, blnPrintBoth, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
27820   Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub cmdRevIncExp_IncomePreview_KeyDown(KeyCode As Integer, Shift As Integer)

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_IncomePreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
27910   intShiftDown = (Shift And acShiftMask) > 0
27920   intAltDown = (Shift And acAltMask) > 0
27930   intCtrlDown = (Shift And acCtrlMask) > 0

27940   intRetVal = KeyCode

        ' ** Plain keys.
27950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27960     Select Case intRetVal
          Case vbKeyTab
27970       With Me
27980         intRetVal = 0
27990         .cmdRevIncExp_IncomePrint.SetFocus
28000       End With
28010     Case vbKeyUp, vbKeyLeft
28020       With Me
28030         intRetVal = 0
28040         If .cmbAccounts.Enabled = True Then
28050           .cmbAccounts.SetFocus
28060         Else
28070           .opgAccountNumber.SetFocus
28080         End If
28090       End With
28100     Case vbKeyDown
28110       With Me
28120         intRetVal = 0
28130         .cmdRevIncExp_ExpensePreview.SetFocus
28140       End With
28150     End Select
28160   End If

        ' ** Shift keys.
28170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28180     Select Case intRetVal
          Case vbKeyTab
28190       With Me
28200         intRetVal = 0
28210         If .cmbAccounts.Enabled = True Then
28220           .cmbAccounts.SetFocus
28230         Else
28240           .opgAccountNumber.SetFocus
28250         End If
28260       End With
28270     End Select
28280   End If

        ' ** Ctrl-Shift keys.
28290   If intCtrlDown And (Not intAltDown) And intShiftDown Then
28300     Select Case intRetVal
          Case vbKeyTab
28310       With Me
28320         intRetVal = 0
28330         .chkPageOf.SetFocus
28340       End With
28350     End Select
28360   End If

EXITP:
28370   KeyCode = intRetVal
28380   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub cmdRevIncExp_IncomePrint_Click()

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_IncomePrint_Click"

        Dim strDocName1 As String, strDocName2 As String

28410   With Me
28420     DoCmd.Hourglass True
28430     DoEvents
28440     strDocName2 = vbNullString
28450     If DoReport = True Then  ' ** Function: Below.
28460       Select Case .opgSummary
            Case .opgSummary_optOnly.OptionValue
28470         strDocName1 = "rptIncExp_Income_Summary"
28480         Select Case gblnDev_Debug
              Case True
28490           DoCmd.OpenReport strDocName1, acViewPreview
28500           DoCmd.Maximize
28510           DoCmd.RunCommand acCmdFitToWindow
28520         Case False
28530           DoCmd.SetWarnings False
28540           If IsLoaded(strDocName1, acReport) = False Then  ' ** Module Function: modFileUtilities.
28550             DoCmd.OpenReport strDocName1, acViewPreview  ' ** This gets the caption changed!
28560           End If
28570           DoCmd.OpenReport strDocName1, acViewNormal
28580           DoCmd.SetWarnings True
28590         End Select
28600       Case Else
28610         Select Case .chkDetail
              Case True
28620           strDocName1 = "rptIncExp_Income_Detailed"
28630         Case False
28640           strDocName1 = "rptIncExp_Income"
28650         End Select
28660         Select Case gblnDev_Debug
              Case True
28670           DoCmd.OpenReport strDocName1, acViewPreview
28680           Select Case .opgSummary
                Case .opgSummary_optInclude.OptionValue  ' ** Don't maximize if both are being previewed.
28690             DoEvents
28700             strDocName2 = "rptIncExp_Income_Summary"
28710             DoCmd.OpenReport strDocName2, acViewNormal
28720           Case .opgSummary_optExclude.OptionValue
28730             DoCmd.Maximize
28740             DoCmd.RunCommand acCmdFitToWindow
28750           End Select
28760         Case False
                '##GTR_Ref: rptIncExp_Income
                '##GTR_Ref: rptIncExp_Income_Detailed
28770           DoCmd.SetWarnings False
28780           If IsLoaded(strDocName1, acReport) = False Then  ' ** Module Function: modFileUtilities.
28790             DoCmd.OpenReport strDocName1, acViewPreview  ' ** This gets the caption changed!
28800           End If
28810           DoCmd.OpenReport strDocName1, acViewNormal
28820           Select Case .opgSummary
                Case .opgSummary_optInclude.OptionValue
28830             DoEvents
28840             strDocName2 = "rptIncExp_Income_Summary"
28850             If IsLoaded(strDocName2, acReport) = False Then  ' ** Module Function: modFileUtilities.
28860               DoCmd.OpenReport strDocName2, acViewPreview  ' ** This gets the caption changed!
28870             End If
                  '##GTR_Ref: rptIncExp_Income_Summary
28880             DoCmd.OpenReport strDocName2, acViewNormal
28890           Case .opgSummary_optExclude.OptionValue
                  ' ** Nothing else.
28900           End Select
28910           DoCmd.SetWarnings False
28920         End Select
28930       End Select
28940     Else
28950       blnPrintBoth = False
28960     End If

28970     DoEvents
28980     If IsLoaded(strDocName1, acReport) = True Then
28990       If gblnDev_Debug = False Then
29000         DoCmd.Close acReport, strDocName1
29010         DoEvents
29020       End If
29030     End If

29040     DoEvents
29050     If strDocName2 <> vbNullString Then
29060       If IsLoaded(strDocName2, acReport) = True Then
29070         If gblnDev_Debug = False Then
29080           DoCmd.Close acReport, strDocName2
29090           DoEvents
29100         End If
29110       End If
29120     End If

29130     If blnPrintBoth = False Then
29140       DoCmd.Hourglass False
29150     End If

29160   End With

EXITP:
29170   Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     blnPrintBoth = False
650     Select Case ERR.Number
        Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
670     End Select
680     Resume EXITP

End Sub

Private Sub cmdRevIncExp_IncomePrint_KeyDown(KeyCode As Integer, Shift As Integer)

29200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_IncomePrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
29210   intShiftDown = (Shift And acShiftMask) > 0
29220   intAltDown = (Shift And acAltMask) > 0
29230   intCtrlDown = (Shift And acCtrlMask) > 0

29240   intRetVal = KeyCode

        ' ** Plain keys.
29250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29260     Select Case intRetVal
          Case vbKeyTab
29270       With Me
29280         intRetVal = 0
29290         .cmdRevIncExp_IncomeWord.SetFocus
29300       End With
29310     Case vbKeyUp
29320       With Me
29330         intRetVal = 0
29340         .cmdPrintAll.SetFocus
29350       End With
29360     Case vbKeyDown
29370       With Me
29380         intRetVal = 0
29390         .cmdRevIncExp_ExpensePrint.SetFocus
29400       End With
29410     End Select
29420   End If

        ' ** Shift keys.
29430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29440     Select Case intRetVal
          Case vbKeyTab
29450       With Me
29460         intRetVal = 0
29470         .cmdRevIncExp_IncomePreview.SetFocus
29480       End With
29490     End Select
29500   End If

EXITP:
29510   KeyCode = intRetVal
29520   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub cmdRevIncExp_IncomeWord_Click()

29600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_IncomeWord_Click"

        Dim strRpt As String, strRptCap1 As String, strRptCap2 As String
        Dim strRptPath1 As String, strRptPath2 As String, strRptPathFile1 As String, strRptPathFile2 As String
        Dim blnContinue As Boolean

29610   With Me
29620     DoCmd.Hourglass True
29630     DoEvents
29640     If DoReport = True Then  ' ** Function: Below.

29650       blnContinue = True

29660       If IsNull(.UserReportPath) = True Then
29670         strRptPath1 = CurrentAppPath  ' ** Module Function: modFileUtilities.
29680       Else
29690         strRptPath1 = .UserReportPath
29700       End If
29710       strRptPath2 = strRptPath1

29720       strRptPathFile1 = vbNullString: strRptPathFile2 = vbNullString: strRptCap1 = vbNullString: strRptCap2 = vbNullString
29730       Select Case .opgSummary
            Case .opgSummary_optOnly.OptionValue
29740         strRptCap2 = "rptIncExp_Income_Summary_" & Format(Date, "yyyymmdd")
29750       Case Else
29760         Select Case .chkDetail
              Case True
29770           strRptCap1 = "rptIncExp_Income_Detailed_"
29780           Select Case .opgSummary
                Case .opgSummary_optInclude.OptionValue
29790             Select Case .opgAccountNumber
                  Case .opgAccountNumber_optSpecified.OptionValue
29800               strRptCap1 = strRptCap1 & gstrAccountNo & "_"
29810             Case .opgAccountNumber_optAll.OptionValue
29820               strRptCap1 = strRptCap1 & "All_"
29830             End Select
29840             strRptCap2 = "rptIncExp_Income_Summary_" & Format(Date, "yyyymmdd")
29850           Case .opgSummary_optExclude.OptionValue
                  ' ** Nothing else.
29860             Select Case .opgAccountNumber
                  Case .opgAccountNumber_optSpecified.OptionValue
29870               strRptCap1 = strRptCap1 & gstrAccountNo & "_"
29880             Case .opgAccountNumber_optAll.OptionValue
29890               strRptCap1 = strRptCap1 & "All_"
29900             End Select
29910           End Select
29920           strRptCap1 = strRptCap1 & Format(Date, "yyyymmdd")
29930         Case False
29940           strRptCap1 = "rptIncExp_Income_"
29950           Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
29960             strRptCap1 = strRptCap1 & gstrAccountNo & "_"
29970           Case .opgAccountNumber_optAll.OptionValue
29980             strRptCap1 = strRptCap1 & "All_"
29990           End Select
30000           strRptCap1 = strRptCap1 & Format(Date, "yyyymmdd")
30010         End Select
30020       End Select

            ' ** Ask where to save the file.
30030       If strRptCap1 <> vbNullString Then
30040         strRptPathFile1 = FileSaveDialog("rtf", strRptCap1 & ".rtf", strRptPath2, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
30050         If strRptPathFile1 = vbNullString Then blnContinue = False
30060       End If
30070       If blnContinue = True Then
30080         If strRptCap2 <> vbNullString Then
30090           If strRptPathFile1 = vbNullString Then  ' ** Only ask if they didn't choose a standard report.
30100             strRptPathFile2 = FileSaveDialog("rtf", strRptCap2 & ".rtf", strRptPath2, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
30110             If strRptPathFile2 = vbNullString Then blnContinue = False
30120           Else  ' ** Use same path as standard report.
30130             strRptPathFile2 = Parse_Path(strRptPathFile1) & LNK_SEP & strRptCap2 & ".rtf"  ' ** Module Function: modFileUtilities.
30140           End If
30150         End If
30160       End If

30170       If blnContinue = True Then
30180         Select Case .opgSummary
              Case .opgSummary_optOnly.OptionValue
30190           strRpt = "rptIncExp_Income_Summary"
30200           strFile3 = strRptPathFile2
30210           Select Case .chkOpenWord
                Case True
30220             DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile2, True
30230           Case False
30240             DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile2, False
30250           End Select
30260         Case Else
30270           Select Case .chkDetail
                Case True
30280             strRpt = "rptIncExp_Income_Detailed"
30290             strFile1 = strRptPathFile1
30300             Select Case .chkOpenWord
                  Case True
30310               DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile1, True
30320             Case False
30330               DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile1, False
30340             End Select
30350             DoEvents
30360             Select Case .opgSummary
                  Case .opgSummary_optInclude.OptionValue
30370               strRpt = "rptIncExp_Income_Summary"
30380               strFile3 = strRptPathFile2
30390               Select Case .chkOpenWord
                    Case True
30400                 DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile2, True
30410               Case False
30420                 DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile2, False
30430               End Select
30440             Case .opgSummary_optExclude.OptionValue
                    ' ** Nothing else.
30450             End Select
30460           Case False
30470             strRpt = "rptIncExp_Income"
30480             strFile1 = strRptPathFile1
30490             Select Case .chkOpenWord
                  Case True
30500               DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile1, True
30510             Case False
30520               DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile1, False
30530             End Select
30540           End Select
30550         End Select
30560       End If  ' ** blnContinue.

30570     End If
30580     If blnPrintBoth = False Then
30590       DoCmd.Hourglass False
30600     End If
30610   End With

EXITP:
30620   Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub cmdRevIncExp_IncomeWord_KeyDown(KeyCode As Integer, Shift As Integer)

30700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_IncomeWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
30710   intShiftDown = (Shift And acShiftMask) > 0
30720   intAltDown = (Shift And acAltMask) > 0
30730   intCtrlDown = (Shift And acCtrlMask) > 0

30740   intRetVal = KeyCode

        ' ** Plain keys.
30750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30760     Select Case intRetVal
          Case vbKeyTab
30770       With Me
30780         intRetVal = 0
30790         If .cmdRevIncExp_IncomeExcel.Enabled = True Then
30800           .cmdRevIncExp_IncomeExcel.SetFocus
30810         Else
30820           .cmdRevIncExp_ExpensePreview.SetFocus
30830         End If
30840       End With
30850     Case vbKeyUp
30860       With Me
30870         intRetVal = 0
30880         .cmdRevIncExp_ExpenseWord.SetFocus
30890       End With
30900     Case vbKeyDown
30910       With Me
30920         intRetVal = 0
30930         .cmdRevIncExp_ExpenseWord.SetFocus
30940       End With
30950     End Select
30960   End If

        ' ** Shift keys.
30970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30980     Select Case intRetVal
          Case vbKeyTab
30990       With Me
31000         intRetVal = 0
31010         .cmdRevIncExp_IncomePrint.SetFocus
31020       End With
31030     End Select
31040   End If

EXITP:
31050   KeyCode = intRetVal
31060   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub cmdRevIncExp_IncomeExcel_Click()

31100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_IncomeExcel_Click"

31110   IncomeExcel_Click_IE strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

31120   DoCmd.Hourglass False

EXITP:
31130   Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub cmdRevIncExp_IncomeExcel_KeyDown(KeyCode As Integer, Shift As Integer)

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_IncomeExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
31210   intShiftDown = (Shift And acShiftMask) > 0
31220   intAltDown = (Shift And acAltMask) > 0
31230   intCtrlDown = (Shift And acCtrlMask) > 0

31240   intRetVal = KeyCode

        ' ** Plain keys.
31250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31260     Select Case intRetVal
          Case vbKeyTab
31270       With Me
31280         intRetVal = 0
31290         .cmdRevIncExp_ExpensePreview.SetFocus
31300       End With
31310     Case vbKeyUp
31320       With Me
31330         intRetVal = 0
31340         .cmdRevIncExp_ExpenseExcel.SetFocus
31350       End With
31360     Case vbKeyDown
31370       With Me
31380         intRetVal = 0
31390         .cmdRevIncExp_ExpenseExcel.SetFocus
31400       End With
31410     End Select
31420   End If

        ' ** Shift keys.
31430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31440     Select Case intRetVal
          Case vbKeyTab
31450       With Me
31460         intRetVal = 0
31470         .cmdRevIncExp_IncomeWord.SetFocus
31480       End With
31490     End Select
31500   End If

EXITP:
31510   KeyCode = intRetVal
31520   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmdRevIncExp_ExpensePreview_Click()

31600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_ExpensePreview_Click"

31610   Prev_Click_IE THIS_PROC, blnPrintBoth, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
31620   Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub cmdRevIncExp_ExpensePreview_KeyDown(KeyCode As Integer, Shift As Integer)

31700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_ExpensePreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
31710   intShiftDown = (Shift And acShiftMask) > 0
31720   intAltDown = (Shift And acAltMask) > 0
31730   intCtrlDown = (Shift And acCtrlMask) > 0

31740   intRetVal = KeyCode

        ' ** Plain keys.
31750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31760     Select Case intRetVal
          Case vbKeyTab
31770       With Me
31780         intRetVal = 0
31790         .cmdRevIncExp_ExpensePrint.SetFocus
31800       End With
31810     Case vbKeyUp
31820       With Me
31830         intRetVal = 0
31840         .cmdRevIncExp_IncomePreview.SetFocus
31850       End With
31860     Case vbKeyDown
31870       With Me
31880         intRetVal = 0
31890         .cmdRevIncExp_IncomePreview.SetFocus
31900       End With
31910     End Select
31920   End If

        ' ** Shift keys.
31930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31940     Select Case intRetVal
          Case vbKeyTab
31950       With Me
31960         intRetVal = 0
31970         If .cmdRevIncExp_IncomeExcel.Enabled = True Then
31980           .cmdRevIncExp_IncomeExcel.SetFocus
31990         Else
32000           .cmdRevIncExp_IncomeWord.SetFocus
32010         End If
32020       End With
32030     End Select
32040   End If

EXITP:
32050   KeyCode = intRetVal
32060   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub cmdRevIncExp_ExpensePrint_Click()

32100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_ExpensePrint_Click"

        Dim strDocName1 As String, strDocName2 As String

32110   With Me
32120     DoCmd.Hourglass True
32130     DoEvents
32140     strDocName2 = vbNullString
32150     If DoReport = True Then  ' ** Function: Below.
32160       Select Case .opgSummary
            Case .opgSummary_optOnly.OptionValue
32170         strDocName1 = "rptIncExp_Expenses_Summary"
32180         Select Case gblnDev_Debug
              Case True
32190           DoCmd.OpenReport strDocName1, acViewPreview
32200           DoCmd.Maximize
32210           DoCmd.RunCommand acCmdFitToWindow
32220         Case False
32230           DoCmd.SetWarnings False
32240           If IsLoaded(strDocName1, acReport) = False Then  ' ** Module Function: modFileUtilities.
32250             DoCmd.OpenReport strDocName1, acViewPreview  ' ** This gets the caption changed!
32260           End If
32270           DoCmd.OpenReport strDocName1, acViewNormal
32280           DoCmd.SetWarnings True
32290         End Select
32300       Case Else
32310         Select Case .chkDetail
              Case True
32320           strDocName1 = "rptIncExp_Expenses_Detailed"
32330         Case False
32340           strDocName1 = "rptIncExp_Expenses"
32350         End Select
32360         Select Case gblnDev_Debug
              Case True
32370           DoCmd.OpenReport strDocName1, acViewPreview
32380           Select Case .opgSummary
                Case .opgSummary_optInclude.OptionValue  ' ** Don't maximize if both are being previewed.
32390             DoEvents
32400             strDocName2 = "rptIncExp_Expenses_Summary"
32410             DoCmd.OpenReport strDocName2, acViewNormal
32420           Case .opgSummary_optExclude.OptionValue
32430             DoCmd.Maximize
32440             DoCmd.RunCommand acCmdFitToWindow
32450           End Select
32460         Case False
                '##GTR_Ref: rptIncExp_Expenses
                '##GTR_Ref: rptIncExp_Expenses_Detailed
32470           DoCmd.SetWarnings False
32480           If IsLoaded(strDocName1, acReport) = False Then  ' ** Module Function: modFileUtilities.
32490             DoCmd.OpenReport strDocName1, acViewPreview  ' ** This gets the caption changed!
32500           End If
32510           DoCmd.OpenReport strDocName1, acViewNormal
32520           Select Case .opgSummary
                Case .opgSummary_optInclude.OptionValue
32530             DoEvents
32540             strDocName2 = "rptIncExp_Expenses_Summary"
32550             If IsLoaded(strDocName2, acReport) = False Then  ' ** Module Function: modFileUtilities.
32560               DoCmd.OpenReport strDocName2, acViewPreview  ' ** This gets the caption changed!
32570             End If
                  '##GTR_Ref: rptIncExp_Expenses_Summary
32580             DoCmd.OpenReport strDocName2, acViewNormal
32590             DoCmd.SetWarnings True
32600           Case .opgSummary_optExclude.OptionValue
                  ' ** Nothing else.
32610           End Select
32620         End Select
32630       End Select
32640     Else
32650       blnPrintBoth = False
32660     End If

32670     DoEvents
32680     If IsLoaded(strDocName1, acReport) = True Then
32690       If gblnDev_Debug = False Then
32700         DoCmd.Close acReport, strDocName1
32710         DoEvents
32720       End If
32730     End If

32740     DoEvents
32750     If strDocName2 <> vbNullString Then
32760       If IsLoaded(strDocName2, acReport) = True Then
32770         If gblnDev_Debug = False Then
32780           DoCmd.Close acReport, strDocName2
32790           DoEvents
32800         End If
32810       End If
32820     End If

32830     If blnPrintBoth = False Then
32840       DoCmd.Hourglass False
32850     End If

32860   End With

EXITP:
32870   Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     blnPrintBoth = False
650     Select Case ERR.Number
        Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
670     End Select
680     Resume EXITP

End Sub

Private Sub cmdRevIncExp_ExpensePrint_KeyDown(KeyCode As Integer, Shift As Integer)

32900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_ExpensePrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
32910   intShiftDown = (Shift And acShiftMask) > 0
32920   intAltDown = (Shift And acAltMask) > 0
32930   intCtrlDown = (Shift And acCtrlMask) > 0

32940   intRetVal = KeyCode

        ' ** Plain keys.
32950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32960     Select Case intRetVal
          Case vbKeyTab
32970       With Me
32980         intRetVal = 0
32990         .cmdRevIncExp_ExpenseWord.SetFocus
33000       End With
33010     Case vbKeyUp
33020       With Me
33030         intRetVal = 0
33040         .cmdRevIncExp_IncomePrint.SetFocus
33050       End With
33060     Case vbKeyDown
33070       With Me
33080         intRetVal = 0
33090         .cmdPrintAll.SetFocus
33100       End With
33110     End Select
33120   End If

        ' ** Shift keys.
33130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33140     Select Case intRetVal
          Case vbKeyTab
33150       With Me
33160         intRetVal = 0
33170         .cmdRevIncExp_ExpensePreview.SetFocus
33180       End With
33190     End Select
33200   End If

EXITP:
33210   KeyCode = intRetVal
33220   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub cmdRevIncExp_ExpenseWord_Click()

33300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_ExpenseWord_Click"

        Dim strRpt As String, strRptCap1 As String, strRptCap2 As String
        Dim strRptPath1 As String, strRptPath2 As String, strRptPathFile1 As String, strRptPathFile2 As String
        Dim blnContinue As Boolean

33310   With Me
33320     DoCmd.Hourglass True
33330     DoEvents
33340     If DoReport = True Then  ' ** Function: Below.

33350       blnContinue = True

33360       If IsNull(.UserReportPath) = True Then
33370         strRptPath1 = CurrentAppPath  ' ** Module Function: modFileUtilities.
33380       Else
33390         strRptPath1 = .UserReportPath
33400       End If
33410       strRptPath2 = strRptPath1

33420       strRptPathFile1 = vbNullString: strRptPathFile2 = vbNullString: strRptCap1 = vbNullString: strRptCap2 = vbNullString
33430       Select Case .opgSummary
            Case .opgSummary_optOnly.OptionValue
33440         strRptCap2 = "rptIncExp_Expenses_Summary_" & Format(Date, "yyyymmdd")
33450       Case Else
33460         Select Case .chkDetail
              Case True
33470           strRptCap1 = "rptIncExp_Expenses_Detailed_"
33480           Select Case .opgSummary
                Case .opgSummary_optInclude.OptionValue
33490             Select Case .opgAccountNumber
                  Case .opgAccountNumber_optSpecified.OptionValue
33500               strRptCap1 = strRptCap1 & gstrAccountNo & "_"
33510             Case .opgAccountNumber_optAll.OptionValue
33520               strRptCap1 = strRptCap1 & "All_"
33530             End Select
33540             strRptCap2 = "rptIncExp_Expenses_Summary_" & Format(Date, "yyyymmdd")
33550           Case .opgSummary_optExclude.OptionValue
                  ' ** Nothing else.
33560             Select Case .opgAccountNumber
                  Case .opgAccountNumber_optSpecified.OptionValue
33570               strRptCap1 = strRptCap1 & gstrAccountNo & "_"
33580             Case .opgAccountNumber_optAll.OptionValue
33590               strRptCap1 = strRptCap1 & "All_"
33600             End Select
33610           End Select
33620           strRptCap1 = strRptCap1 & Format(Date, "yyyymmdd")
33630         Case False
33640           strRptCap1 = "rptIncExp_Expenses_"
33650           Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
33660             strRptCap1 = strRptCap1 & gstrAccountNo & "_"
33670           Case .opgAccountNumber_optAll.OptionValue
33680             strRptCap1 = strRptCap1 & "All_"
33690           End Select
33700           strRptCap1 = strRptCap1 & Format(Date, "yyyymmdd")
33710         End Select
33720       End Select

            ' ** Ask where to save the file.
33730       If strRptCap1 <> vbNullString Then
33740         strRptPathFile1 = FileSaveDialog("rtf", strRptCap1 & ".rtf", strRptPath1, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
33750         If strRptPathFile1 = vbNullString Then blnContinue = False
33760       End If
33770       If blnContinue = True Then
33780         If strRptCap2 <> vbNullString Then
33790           If strRptPathFile1 = vbNullString Then  ' ** Only ask if they didn't choose a standard report.
33800             strRptPathFile2 = FileSaveDialog("rtf", strRptCap2 & ".rtf", strRptPath2, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
33810             If strRptPathFile2 = vbNullString Then blnContinue = False
33820           Else  ' ** Use same path as standard report.
33830             strRptPathFile2 = Parse_Path(strRptPathFile1) & LNK_SEP & strRptCap2 & ".rtf"  ' ** Module Function: modFileUtilities.
33840           End If
33850         End If
33860       End If  ' ** blnContinue.

33870       If blnContinue = True Then
33880         Select Case .opgSummary
              Case .opgSummary_optOnly.OptionValue
33890           strRpt = "rptIncExp_Expenses_Summary"
33900           strFile4 = strRptPathFile2
33910           Select Case .chkOpenWord
                Case True
33920             DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile2, True
33930           Case False
33940             DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile2, False
33950           End Select
33960         Case Else
33970           Select Case .chkDetail
                Case True
33980             strRpt = "rptIncExp_Expenses_Detailed"
33990             strFile2 = strRptPathFile1
34000             Select Case .chkOpenWord
                  Case True
34010               DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile1, True
34020             Case False
34030               DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile1, False
34040             End Select
34050             DoEvents
34060             Select Case .opgSummary
                  Case .opgSummary_optInclude.OptionValue
34070               strRpt = "rptIncExp_Expenses_Summary"
34080               strFile4 = strRptPathFile2
34090               Select Case .chkOpenWord
                    Case True
34100                 DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile2, True
34110               Case False
34120                 DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile2, False
34130               End Select
34140             Case .opgSummary_optExclude.OptionValue
                    ' ** Nothing else.
34150             End Select
34160           Case False
34170             strRpt = "rptIncExp_Expenses"
34180             strFile2 = strRptPathFile1
34190             Select Case .chkOpenWord
                  Case True
34200               DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile1, True
34210             Case False
34220               DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile1, False
34230             End Select
34240           End Select
34250         End Select
34260       End If  ' ** blnContinue.

34270     End If
34280     If blnPrintBoth = False Then
34290       DoCmd.Hourglass False
34300     End If
34310   End With

EXITP:
34320   Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub cmdRevIncExp_ExpenseWord_KeyDown(KeyCode As Integer, Shift As Integer)

34400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_ExpenseWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
34410   intShiftDown = (Shift And acShiftMask) > 0
34420   intAltDown = (Shift And acAltMask) > 0
34430   intCtrlDown = (Shift And acCtrlMask) > 0

34440   intRetVal = KeyCode

        ' ** Plain keys.
34450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34460     Select Case intRetVal
          Case vbKeyTab
34470       With Me
34480         intRetVal = 0
34490         If .cmdRevIncExp_ExpenseExcel.Enabled = True Then
34500           .cmdRevIncExp_ExpenseExcel.SetFocus
34510         Else
34520           .cmdPrintAll.SetFocus
34530         End If
34540       End With
34550     Case vbKeyUp
34560       With Me
34570         intRetVal = 0
34580         .cmdRevIncExp_IncomeWord.SetFocus
34590       End With
34600     Case vbKeyDown
34610       With Me
34620         intRetVal = 0
34630         .cmdRevIncExp_IncomeWord.SetFocus
34640       End With
34650     End Select
34660   End If

        ' ** Shift keys.
34670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34680     Select Case intRetVal
          Case vbKeyTab
34690       With Me
34700         intRetVal = 0
34710         .cmdRevIncExp_ExpensePrint.SetFocus
34720       End With
34730     End Select
34740   End If

EXITP:
34750   KeyCode = intRetVal
34760   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Public Sub cmdRevIncExp_ExpenseExcel_Click()

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_ExpenseExcel_Click"

34810   ExpenseExcel_Click_IE strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
34820   Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub cmdRevIncExp_ExpenseExcel_KeyDown(KeyCode As Integer, Shift As Integer)

34900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRevIncExp_ExpenseExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
34910   intShiftDown = (Shift And acShiftMask) > 0
34920   intAltDown = (Shift And acAltMask) > 0
34930   intCtrlDown = (Shift And acCtrlMask) > 0

34940   intRetVal = KeyCode

        ' ** Plain keys.
34950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34960     Select Case intRetVal
          Case vbKeyTab, vbKeyRight
34970       With Me
34980         intRetVal = 0
34990         .cmdPrintAll.SetFocus
35000       End With
35010     Case vbKeyUp
35020       With Me
35030         intRetVal = 0
35040         .cmdRevIncExp_IncomeExcel.SetFocus
35050       End With
35060     Case vbKeyDown
35070       With Me
35080         intRetVal = 0
35090         .cmdRevIncExp_IncomeExcel.SetFocus
35100       End With
35110     End Select
35120   End If

        ' ** Shift keys.
35130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35140     Select Case intRetVal
          Case vbKeyTab
35150       With Me
35160         intRetVal = 0
35170         .cmdRevIncExp_ExpenseWord.SetFocus
35180       End With
35190     End Select
35200   End If

EXITP:
35210   KeyCode = intRetVal
35220   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmdPrintAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_box01_MouseMove"

35310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
35320   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdPrintAll_GotFocus()

35400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_GotFocus"

35410   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
35420   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdPrintAll_Click()

35500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_Click"

35510   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
35520   Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     blnPrintBoth = False
650     Select Case ERR.Number
        Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
670     End Select
680     Resume EXITP

End Sub

Private Sub cmdPrintAll_KeyDown(KeyCode As Integer, Shift As Integer)

35600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
35610   intShiftDown = (Shift And acShiftMask) > 0
35620   intAltDown = (Shift And acAltMask) > 0
35630   intCtrlDown = (Shift And acCtrlMask) > 0

35640   intRetVal = KeyCode

        ' ** Plain keys.
35650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
35660     Select Case intRetVal
          Case vbKeyTab
35670       With Me
35680         intRetVal = 0
35690         .cmdClose.SetFocus
35700       End With
35710     Case vbKeyUp, vbKeyLeft
35720       With Me
35730         intRetVal = 0
35740         .cmdRevIncExp_ExpensePrint.SetFocus
35750       End With
35760     Case vbKeyDown, vbKeyRight
35770       With Me
35780         intRetVal = 0
35790         .cmdClose.SetFocus
35800       End With
35810     End Select
35820   End If

        ' ** Shift keys.
35830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35840     Select Case intRetVal
          Case vbKeyTab
35850       With Me
35860         intRetVal = 0
35870         If .cmdRevIncExp_ExpenseExcel.Enabled = True Then
35880           .cmdRevIncExp_ExpenseExcel.SetFocus
35890         Else
35900           .cmdRevIncExp_ExpenseWord.SetFocus
35910         End If
35920       End With
35930     End Select
35940   End If

EXITP:
35950   KeyCode = intRetVal
35960   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmdPrintAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

36000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_MouseMove"

36010   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
36020   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdPrintAll_LostFocus()

36100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_LostFocus"

36110   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
36120   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdWordAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

36200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_box01_MouseMove"

36210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
36220   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdWordAll_GotFocus()

36300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_GotFocus"

36310   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
36320   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdWordAll_Click()

36400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_Click"

36410   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
36420   Exit Sub

ERRH:
630     gblnPrintAll = False
640     Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
650     Case Else
660       Select Case ERR.Number
          Case Else
670         zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
680       End Select
690     End Select
700     Resume EXITP

End Sub

Private Sub cmdWordAll_KeyDown(KeyCode As Integer, Shift As Integer)

36500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

36510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
36520   intShiftDown = (Shift And acShiftMask) > 0
36530   intAltDown = (Shift And acAltMask) > 0
36540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
36550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
36560     Select Case intRetVal
          Case vbKeyTab
36570       With Me
36580         intRetVal = 0
36590         If .cmdExcelAll.Enabled = True Then
36600           .cmdExcelAll.SetFocus
36610         Else
36620           .cmdClose.SetFocus
36630         End If
36640       End With
36650     Case vbKeyUp
36660       With Me
36670         intRetVal = 0
36680         .cmdRevIncExp_ExpenseWord.SetFocus
36690       End With
36700     Case vbKeyDown
36710       With Me
36720         intRetVal = 0
36730         .cmdRevIncExp_IncomeWord.SetFocus
36740       End With
36750     Case vbKeyLeft
36760       With Me
36770         intRetVal = 0
36780         .cmdPrintAll.SetFocus
36790       End With
36800     Case vbKeyRight
36810       With Me
36820         intRetVal = 0
36830         If .cmdExcelAll.Enabled = True Then
36840           .cmdExcelAll.SetFocus
36850         Else
36860           .cmdRevIncExp_IncomePreview.SetFocus
36870         End If
36880       End With
36890     End Select
36900   End If

        ' ** Shift keys.
36910   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
36920     Select Case intRetVal
          Case vbKeyTab
36930       With Me
36940         intRetVal = 0
36950         .cmdPrintAll.SetFocus
36960       End With
36970     End Select
36980   End If

        ' ** Ctrl keys.
36990   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
37000     Select Case intRetVal
          Case vbKeyTab
37010       With Me
37020         intRetVal = 0
37030         .chkOpenWord.SetFocus
37040       End With
37050     End Select
37060   End If

EXITP:
37070   KeyCode = intRetVal
37080   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmdWordAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

37100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_MouseMove"

37110   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
37120   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdWordAll_LostFocus()

37200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_LostFocus"

37210   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
37220   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub chkOpenWord_AfterUpdate()

37300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_AfterUpdate"

37310   OpenApp_Handler_IE THIS_PROC, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
37320   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub chkOpenWord_KeyDown(KeyCode As Integer, Shift As Integer)

37400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

37410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
37420   intShiftDown = (Shift And acShiftMask) > 0
37430   intAltDown = (Shift And acAltMask) > 0
37440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
37450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
37460     Select Case intRetVal
          Case vbKeyTab
37470       With Me
37480         intRetVal = 0
37490         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
37500           .cmdExcelAll.SetFocus
37510         Else
37520           .cmdClose.SetFocus
37530         End If
37540       End With
37550     Case vbKeyUp
37560       With Me
37570         intRetVal = 0
37580         .cmdWordAll.SetFocus
37590       End With
37600     Case vbKeyDown
37610       With Me
37620         intRetVal = 0
37630         .cmdRevIncExp_IncomeWord.SetFocus
37640       End With
37650     End Select
37660   End If

        ' ** Shift keys.
37670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
37680     Select Case intRetVal
          Case vbKeyTab
37690       With Me
37700         intRetVal = 0
37710         .cmdWordAll.SetFocus
37720       End With
37730     End Select
37740   End If

        ' ** Ctrl keys.
37750   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
37760     Select Case intRetVal
          Case vbKeyTab
37770       With Me
37780         intRetVal = 0
37790         If .chkOpenExcel.Visible = True And .chkOpenExcel.Enabled = True Then
37800           .chkOpenExcel.SetFocus
37810         Else
37820           Beep
37830         End If
37840       End With
37850     End Select
37860   End If

EXITP:
37870   KeyCode = intRetVal
37880   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub chkOpenWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

37900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_MouseMove"

37910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
37920   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdExcelAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

38000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_box01_MouseMove"

38010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
38020   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdExcelAll_GotFocus()

38100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_GotFocus"

38110   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
38120   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdExcelAll_Click()

38200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_Click"

38210   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
38220   Exit Sub

ERRH:
630     gblnPrintAll = False
640     DoCmd.Hourglass False
650     Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
660     Case Else
670       Select Case ERR.Number
          Case Else
680         zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
690       End Select
700     End Select
710     Resume EXITP

End Sub

Private Sub cmdExcelAll_KeyDown(KeyCode As Integer, Shift As Integer)

38300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

38310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
38320   intShiftDown = (Shift And acShiftMask) > 0
38330   intAltDown = (Shift And acAltMask) > 0
38340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
38350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
38360     Select Case intRetVal
          Case vbKeyTab
38370       With Me
38380         intRetVal = 0
38390         .cmdClose.SetFocus
38400       End With
38410     Case vbKeyUp
38420       With Me
38430         intRetVal = 0
38440         .cmdRevIncExp_ExpenseExcel.SetFocus
38450       End With
38460     Case vbKeyDown
38470       With Me
38480         intRetVal = 0
38490         .cmdRevIncExp_IncomeExcel.SetFocus
38500       End With
38510     Case vbKeyLeft
38520       With Me
38530         intRetVal = 0
38540         .cmdWordAll.SetFocus
38550       End With
38560     Case vbKeyRight
38570       With Me
38580         intRetVal = 0
38590         .cmdRevIncExp_IncomePreview.SetFocus
38600       End With
38610     End Select
38620   End If

        ' ** Shift keys.
38630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
38640     Select Case intRetVal
          Case vbKeyTab
38650       With Me
38660         intRetVal = 0
38670         .cmdWordAll.SetFocus
38680       End With
38690     End Select
38700   End If

        ' ** Ctrl keys.
38710   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
38720     Select Case intRetVal
          Case vbKeyTab
38730       With Me
38740         intRetVal = 0
38750         .chkOpenExcel.SetFocus
38760       End With
38770     End Select
38780   End If

EXITP:
38790   KeyCode = intRetVal
38800   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub cmdExcelAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

38900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_MouseMove"

38910   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
38920   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdExcelAll_LostFocus()

39000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_LostFocus"

39010   DoAll_Handler_IE THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, _
          blnPrintBoth, strFile1, strFile2, strFile3, strFile4, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
39020   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub chkOpenExcel_AfterUpdate()

39100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_AfterUpdate"

39110   OpenApp_Handler_IE THIS_PROC, Me  ' ** Module Procedure: modIncExpFuncs.

EXITP:
39120   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub chkOpenExcel_KeyDown(KeyCode As Integer, Shift As Integer)

39200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

39210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
39220   intShiftDown = (Shift And acShiftMask) > 0
39230   intAltDown = (Shift And acAltMask) > 0
39240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
39250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
39260     Select Case intRetVal
          Case vbKeyTab
39270       With Me
39280         intRetVal = 0
39290         .cmdClose.SetFocus
39300       End With
39310     Case vbKeyUp
39320       With Me
39330         intRetVal = 0
39340         .cmdExcelAll.SetFocus
39350       End With
39360     Case vbKeyDown
39370       With Me
39380         intRetVal = 0
39390         .cmdRevIncExp_IncomeExcel.SetFocus
39400       End With
39410     End Select
39420   End If

        ' ** Shift keys.
39430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
39440     Select Case intRetVal
          Case vbKeyTab
39450       With Me
39460         intRetVal = 0
39470         .cmdExcelAll.SetFocus
39480       End With
39490     End Select
39500   End If

EXITP:
39510   KeyCode = intRetVal
39520   Exit Sub

ERRH:
630     intRetVal = 0
640     THAT_PROC = THIS_PROC
650     That_Erl = Erl
660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
670     Resume EXITP

End Sub

Private Sub chkOpenExcel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

39600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_MouseMove"

39610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
39620   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Public Function DoReport() As Boolean

39700 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim blnRetVal As Boolean

39710   blnRetVal = DoReport_IE(Me)  ' ** Module Function: modIncExpFuncs.

EXITP:
39720   DoReport = blnRetVal
39730   Exit Function

ERRH:
630     blnRetVal = False
640     DoCmd.Hourglass False
650     Select Case ERR.Number
        Case 3000  ' ** Reserved. There is no message for this error.
660       Beep
670       MsgBox "Trust Accountant is unable to complete your request." & vbCrLf & _
            "Exit the program, then try again.", vbCritical + vbOKOnly, "Error 3000"
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Function

Private Sub GTREmblem_Off()

39800 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

39810   With Me
39820     blnGTR_Emblem = False
39830     For lngX = 1& To 24&
39840       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
39850     Next
39860   End With

EXITP:
39870   Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub
