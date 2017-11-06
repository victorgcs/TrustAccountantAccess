VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_ArchivedTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_ArchivedTransactions"

'VGC 09/23/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** Shortcut Alt keys responsive from this form:
' **   Select Account:   A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Excel:            E {cmdExcel}
' **   From:             F {DateStart}
' **   Legal Name:       G {chkLegalName}
' **   Utility Menu:     I {opgReturn_optMenu}
' **   All Accts:        L {opgAccountNumber_optAll}
' **   Name:             M {opgAccountSource_optName}
' **   Combine Cash:     N {chkCombineCash}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Include Closed:   S {chkIncludeClosed}
' **   To:               T {DateEnd}
' **   Specified Acct:   U {opgAccountNumber_optSpecified}
' **   Archive Trans:    V {opgReturn_optSelect}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Sort By Date:     D {opgOrderBy_optDate}
' **   Page Count:       G {chkPageOf}
' **   Combine Cash:     O {chkCombineCash}
' **   Sort By Type:     T {opgOrderBy_optType}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean

Private datArchiveDate As Date, blnNoTrans As Boolean
Private strCallingForm As Variant
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim lngTmp01 As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strCallingForm = .OpenArgs
150           intPos01 = InStr(strCallingForm, "~")
160           If intPos01 > 0 Then

170             blnNoTrans = False

180             lngTmp01 = CLng(Val(Mid(strCallingForm, (intPos01 + 1))))
190             If lngTmp01 > 0& Then
200               datArchiveDate = CDate(lngTmp01)
210               .DateEnd = datArchiveDate
220             Else
230               blnNoTrans = True
240               .NoArchivedTrans_lbl.Visible = True
250             End If
260             strCallingForm = Left(strCallingForm, (intPos01 - 1))

270             blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
280             blnCalendar2_Focus = False: blnCalendar2_MouseDown = False

290             gvarCrtRpt_FL_SpecData = Null
300             gstrFormQuerySpec = THIS_NAME

310             .opgOrderBy = .opgOrderBy_optDate.OptionValue
320             .chkCombineCash = False

330             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
340             .cmdExcel.Enabled = False
      #Else
350             .cmdExcel.Enabled = True
      #End If

                ' ** EVENT CHECK: chkRememberMe!
                ' ** NONE!
                ' ** EVENT CHECK: chkRememberDates!
360             If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
370               DateStart_Exit Cancel  ' ** Procedure: Below.
380               DateEnd_Exit Cancel  ' ** Procedure: Below.
390             End If

400             opgOrderBy_AfterUpdate  ' ** Procedure: Below.
410             chkCombineCash_AfterUpdate  ' ** Procedure: Below.
420             chkIncludeClosed_AfterUpdate  ' ** Procedure: Below.
430             chkLegalName_AfterUpdate  ' ** Procedure: Below.
440             opgAccountSource_AfterUpdate  ' ** Procedure: Below.
450             chkRememberMe_AfterUpdate  ' ** Procedure: Below.
460             chkRememberDates_AfterUpdate  ' ** Procedure: Below.
470             chkPageOf_AfterUpdate  ' ** Procedure: Below.
480             opgAccountNumber_AfterUpdate  ' ** Procedure: Below.
490             opgReturn_AfterUpdate  ' ** Procedure: Below.
500             SetClosedAccts  ' ** Procedure: Below.

510             If blnNoTrans = True Then
520               .cmdClose.SetFocus
530               .DateStart.Enabled = False
540               .DateStart.BorderColor = WIN_CLR_DISR
550               .DateStart.BackStyle = acBackStyleTransparent
560               .cmdCalendar1.Enabled = False
570               .cmdCalendar1_raised_img_dis.Visible = True
580               .cmdCalendar1_raised_img.Visible = False
590               .cmdCalendar1_raised_semifocus_dots_img.Visible = False
600               .cmdCalendar1_raised_focus_img.Visible = False
610               .cmdCalendar1_raised_focus_dots_img.Visible = False
620               .cmdCalendar1_sunken_focus_dots_img.Visible = False
630               .DateEnd.Enabled = False
640               .DateEnd.BorderColor = WIN_CLR_DISR
650               .DateEnd.BackStyle = acBackStyleTransparent
660               .cmdCalendar2.Enabled = False
670               .cmdCalendar2_raised_img_dis.Visible = True
680               .cmdCalendar2_raised_img.Visible = False
690               .cmdCalendar2_raised_semifocus_dots_img.Visible = False
700               .cmdCalendar2_raised_focus_img.Visible = False
710               .cmdCalendar2_raised_focus_dots_img.Visible = False
720               .cmdCalendar2_sunken_focus_dots_img.Visible = False
730               .chkRememberDates.Enabled = False
740               .chkRememberDates_lbl.Visible = False
750               .chkRememberDates_lbl2_dim.Visible = True
760               .chkRememberDates_lbl2_dim_hi.Visible = True
770               .opgAccountNumber.Enabled = False
780               .opgAccountNumber_box2.Visible = True
790               .opgAccountNumber_vline01.Visible = True
800               .opgAccountNumber_vline02.Visible = True
810               .opgAccountNumber_vline03.Visible = True
820               .opgAccountNumber_vline04.Visible = True
830               .opgAccountNumber_hline01.Visible = True
840               .opgAccountNumber_hline02.Visible = True
850               .opgAccountNumber_hline03.Visible = True
860               .opgAccountNumber_optSpecified_lbl_box.Visible = False
870               .opgAccountNumber_optAll_lbl_box.Visible = False
880               .cmbAccounts.Enabled = False
890               .cmbAccounts.BorderColor = WIN_CLR_DISR
900               .cmbAccounts.BackStyle = acBackStyleTransparent
910               .opgAccountSource.Enabled = False
920               .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
930               .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
940               .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
950               .opgAccountSource_optName_lbl2_dim_hi.Visible = True
960               .chkRememberMe.Enabled = False
970               .chkRememberMe_lbl.Visible = False
980               .chkRememberMe_lbl2_dim.Visible = True
990               .chkRememberMe_lbl2_dim_hi.Visible = True
1000              .chkPageOf.Enabled = False
1010              .chkPageOf_lbl2.ForeColor = WIN_CLR_DISF
1020              .chkPageOf_lbl2_dim_hi.Visible = True
1030              .opgOrderBy.Enabled = False
1040              .chkIncludeClosed.Enabled = False
1050              .chkCombineCash.Enabled = False
1060              .chkLegalName.Enabled = False
1070              .cmdPreview.Enabled = False
1080              .cmdPrint.Enabled = False
1090              .cmdWord.Enabled = False
1100              .cmdExcel.Enabled = False
1110            End If

1120            .cmbAccounts.Requery

1130          Else
1140            Cancel = -1
1150          End If
1160        Else
1170          Cancel = -1
1180        End If
1190      Else
1200        Cancel = -1
1210      End If
1220    End With

1230    If Cancel = -1 Then
1240      Beep
1250      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1260    End If

EXITP:
1270    Exit Sub

ERRH:
1280    Select Case ERR.Number
        Case Else
1290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1300    End Select
1310    Resume EXITP

End Sub

Private Sub Form_Load()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

1410    With Me

1420      Set dbs = CurrentDb
          ' ** Empty tmpAccount.
1430      Set qdf = dbs.QueryDefs("qryArchiveTransactions_01_01")
1440      qdf.Execute
1450      Set qdf = Nothing
1460      DoEvents
          ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
          ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
          ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
          ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
          ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
          ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
          ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
          ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
1470      Set qdf = dbs.QueryDefs("qryArchiveTransactions_01_02")
1480      qdf.Execute
1490      Set qdf = Nothing
1500      DoEvents
1510      dbs.Close
1520      Set dbs = Nothing
1530      DoEvents

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
1540      Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
1550      clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
1560      clsMonthClass.PositionAtCursor = True

1570      .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.

1580    End With

EXITP:
1590    Set qdf = Nothing
1600    Set dbs = Nothing
1610    Exit Sub

ERRH:
1620    DoCmd.Hourglass False
1630    Select Case ERR.Number
        Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1720    intShiftDown = (Shift And acShiftMask) > 0
1730    intAltDown = (Shift And acAltMask) > 0
1740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Number:           B {opgAccountSource_optNumber}
        ' **   Utility Menu:     I {opgReturn_optMenu}
        ' **   Name:             M {opgAccountSource_optName}
        ' **   Archive Trans:    V {opgReturn_optSelect}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Order By Date:    D {opgOrderBy_optDate}
        ' **   Page Count:       G {chkPageOf}
        ' **   Combine Cash:     O {chkCombineCash}
        ' **   Order By Type:    T {opgOrderBy_optType}

        ' ** Plain keys.
1750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1760      Select Case intRetVal
          Case vbKeyEscape
1770        intRetVal = 0
1780        cmdClose_Click  ' ** Procedure: Below.
1790      Case vbKeyF7
1800        intRetVal = 0
1810        cmdCalendar1_Click  ' ** Procedure: Below.
1820      Case vbKeyF8
1830        intRetVal = 0
1840        cmdCalendar2_Click  ' ** Procedure: Below.
1850      End Select
1860    End If

        ' ** Alt keys.
1870    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1880      Select Case intRetVal
          Case vbKeyB
1890        With Me
1900          intRetVal = 0
1910          If .opgAccountSource.Enabled = True Then
1920            .opgAccountSource = .opgAccountSource_optNumber.OptionValue
1930            .opgAccountSource.SetFocus
1940            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1950          Else
1960            Beep
1970          End If
1980        End With
1990      Case vbKeyI
2000        With Me
2010          intRetVal = 0
2020          .opgReturn.SetFocus
2030          .opgReturn = .opgReturn_optMenu.OptionValue
2040          opgReturn_AfterUpdate  ' ** Procedure: Below.
2050        End With
2060      Case vbKeyM
2070        With Me
2080          intRetVal = 0
2090          If .opgAccountSource.Enabled = True Then
2100            .opgAccountSource = .opgAccountSource_optName.OptionValue
2110            .opgAccountSource.SetFocus
2120            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
2130          Else
2140            Beep
2150          End If
2160        End With
2170      Case vbKeyV
2180        With Me
2190          intRetVal = 0
2200          .opgReturn.SetFocus
2210          .opgReturn = .opgReturn_optSelect.OptionValue
2220          opgReturn_AfterUpdate  ' ** Procedure: Below.
2230        End With
2240      Case vbKeyX
2250        intRetVal = 0
2260        cmdClose_Click  ' ** Procedure: Below.
2270      End Select
2280    End If

        ' ** Ctrl keys.
2290    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2300      Select Case intRetVal
          Case vbKeyD
2310        With Me
2320          intRetVal = 0
2330          .opgOrderBy = .opgOrderBy_optDate.OptionValue
2340          .opgOrderBy.SetFocus
2350          opgOrderBy_AfterUpdate  ' ** Procedure: Below.
2360        End With
2370      Case vbKeyG
2380        With Me
2390          intRetVal = 0
2400          If .chkPageOf.Enabled = True Then
2410            .chkPageOf.SetFocus
2420            .chkPageOf = (Not .chkPageOf)
2430            chkPageOf_AfterUpdate  ' ** Procedure: Below.
2440          End If
2450        End With
2460      Case vbKeyO
2470        With Me
2480          intRetVal = 0
2490          .chkCombineCash.SetFocus
2500          .chkCombineCash = Not .chkCombineCash
2510          chkCombineCash_AfterUpdate  ' ** Procedure: Below.
2520        End With
2530      Case vbKeyT
2540        With Me
2550          intRetVal = 0
2560          .opgOrderBy = .opgOrderBy_optType.OptionValue
2570          .opgOrderBy.SetFocus
2580          opgOrderBy_AfterUpdate  ' ** Procedure: Below.
2590        End With
2600      End Select
2610    End If

        ' ** Ctrl-Shift keys.
2620    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2630      Select Case intRetVal
          Case vbKeyF
2640        With Me
2650          intRetVal = 0
2660          .FocusHolder.SetFocus
2670          .DateStart = Null
2680        End With
2690      End Select
2700    End If

EXITP:
2710    KeyCode = intRetVal
2720    Exit Sub

ERRH:
2730    intRetVal = 0
2740    THAT_PROC = THIS_PROC
2750    That_Erl = Erl
2760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2770    Resume EXITP

End Sub

Private Sub Form_Timer()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2810    With Me
2820      .TimerInterval = 0&
2830      If gblnSetFocus = True Then
2840        gblnSetFocus = False
2850        strFormName = vbNullString
2860  On Error Resume Next
2870        Do Until Screen.ActiveForm.Name = strFormName
2880          DoCmd.SelectObject acForm, THIS_NAME, False
2890          strFormName = Screen.ActiveForm.Name
2900        Loop
2910  On Error GoTo ERRH
2920      End If
2930      If gblnGoToReport = True Then
2940        gblnGoToReport = False
2950        .cmdPrint.SetFocus
2960        Beep
2970        DoCmd.Hourglass False
2980        DoEvents
2990      Else
3000        DoCmd.Hourglass False
3010      End If
3020    End With

EXITP:
3030    Exit Sub

ERRH:
3040    Select Case ERR.Number
        Case Else
3050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3060    End Select
3070    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3110    If DataErr <> 0 Then
3120      Select Case DataErr
          Case 3101, 8519, 2108, 2116, 3020, 2501, 2169, 7753, 3314, 2237
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
3130      Case 3162
3140        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
3150        DoCmd.CancelEvent
3160      Case 2113, 2279
            ' ** 2113: The value you entered isn't valid for this field.
            ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3170        If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateStart, DateEnd.
3180          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3190          Response = acDataErrContinue
3200        End If
3210      Case Else
3220        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3230        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3240      End Select
3250    End If

EXITP:
3260    Exit Sub

ERRH:
3270    Select Case ERR.Number
        Case Else
3280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290    End Select
3300    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

3410    With Me

3420      If .chkRememberMe = False Then
3430        .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
3440      End If
3450      If .chkRememberDates = False Then
3460        .DateStart = Null
3470        .DateEnd = Null
3480      End If

3490      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3500      Set dbs = CurrentDb
          ' ** Empty tmpAccount.
3510      Set qdf = dbs.QueryDefs("qryArchiveTransactions_01_01")
3520      qdf.Execute
3530      Set qdf = Nothing
3540      dbs.Close
3550      Set dbs = Nothing
3560      DoEvents

3570      ChkSpecLedgerEntry  ' ** Module Function: modUtilities.
3580      gvarCrtRpt_FL_SpecData = Null
3590      gstrFormQuerySpec = vbNullString
3600      gblnSetFocus = True
3610      gstrReturningForm = THIS_NAME
3620      Select Case .opgReturn
          Case .opgReturn_optSelect.OptionValue
3630        DoCmd.OpenForm "frmArchiveTransactions", , , , , , "frmMenu_Utility"
3640      Case .opgReturn_optMenu.OptionValue
3650        DoCmd.OpenForm "frmMenu_Utility", , , , , , "frmMenu_Main"
3660      End Select

3670    End With

EXITP:
3680    Set qdf = Nothing
3690    Set dbs = Nothing
3700    Exit Sub

ERRH:
3710    Select Case ERR.Number
        Case Else
3720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3730    End Select
3740    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

3810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3910    With Me
3920      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
3930        Select Case blnCalendar1_Focus
            Case True
3940          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
3950          .cmdCalendar1_raised_img.Visible = False
3960        Case False
3970          .cmdCalendar1_raised_img.Visible = True
3980          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
3990        End Select
4000        .cmdCalendar1_raised_focus_dots_img.Visible = False
4010        .cmdCalendar1_raised_focus_img.Visible = False
4020        .cmdCalendar1_sunken_focus_dots_img.Visible = False
4030        .cmdCalendar1_raised_img_dis.Visible = False
4040      End If
4050      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
4060        Select Case blnCalendar2_Focus
            Case True
4070          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
4080          .cmdCalendar2_raised_img.Visible = False
4090        Case False
4100          .cmdCalendar2_raised_img.Visible = True
4110          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
4120        End Select
4130        .cmdCalendar2_raised_focus_dots_img.Visible = False
4140        .cmdCalendar2_raised_focus_img.Visible = False
4150        .cmdCalendar2_sunken_focus_dots_img.Visible = False
4160        .cmdCalendar2_raised_img_dis.Visible = False
4170      End If
4180    End With

EXITP:
4190    Exit Sub

ERRH:
4200    Select Case ERR.Number
        Case Else
4210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4220    End Select
4230    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

4310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Private Sub cmdClose_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

4410    DoCmd.Close acForm, THIS_NAME

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4520    intShiftDown = (Shift And acShiftMask) > 0
4530    intAltDown = (Shift And acAltMask) > 0
4540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab
4570        With Me
4580          intRetVal = 0
4590          If blnNoTrans = False Then
4600            .DateStart.SetFocus
4610          Else
4620            Beep
4630          End If
4640        End With
4650      End Select
4660    End If

        ' ** Shift keys.
4670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4680      Select Case intRetVal
          Case vbKeyTab
4690        With Me
4700          intRetVal = 0
4710          If blnNoTrans = False Then
4720            If .cmdExcel.Enabled = True Then
4730              .cmdExcel.SetFocus
4740            Else
4750              .cmdWord.SetFocus
4760            End If
4770          Else
4780            Beep
4790          End If
4800        End With
4810      End Select
4820    End If

EXITP:
4830    KeyCode = intRetVal
4840    Exit Sub

ERRH:
4850    intRetVal = 0
4860    THAT_PROC = THIS_PROC
4870    That_Erl = Erl
4880    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4890    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

4910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5020    intShiftDown = (Shift And acShiftMask) > 0
5030    intAltDown = (Shift And acAltMask) > 0
5040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5070        With Me
5080          intRetVal = 0
5090          .DateEnd.SetFocus
5100        End With
5110      End Select
5120    End If

        ' ** Shift keys.
5130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5150        With Me
5160          intRetVal = 0
5170          .cmdClose.SetFocus
5180        End With
5190      End Select
5200    End If

        ' ** Ctrl keys.
5210    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
5220      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5230        With Me
5240          intRetVal = 0
5250          .cmdCalendar1.SetFocus
5260        End With
5270      End Select
5280    End If

EXITP:
5290    KeyCode = intRetVal
5300    Exit Sub

ERRH:
5310    intRetVal = 0
5320    THAT_PROC = THIS_PROC
5330    That_Erl = Erl
5340    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5350    Resume EXITP

End Sub

Private Sub DateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_MouseMove"

5410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
5420    Exit Sub

ERRH:
5430    Select Case ERR.Number
        Case Else
5440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5450    End Select
5460    Resume EXITP

End Sub

Private Sub DateStart_Exit(Cancel As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

5510    With Me
5520      If IsDate(.DateEnd) And IsDate(.DateStart) Then
5530        If CDate(.DateStart) > CDate(.DateEnd) Then
5540          MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Entry"
5550          Cancel = -1
5560        End If
5570      End If
5580    End With

EXITP:
5590    Exit Sub

ERRH:
5600    THAT_PROC = THIS_PROC
5610    That_Erl = Erl
5620    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5630    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

5710    With Me
5720      blnCalendar1_Focus = True
5730      .cmdCalendar1_raised_semifocus_dots_img.Visible = True
5740      .cmdCalendar1_raised_img.Visible = False
5750      .cmdCalendar1_raised_focus_img.Visible = False
5760      .cmdCalendar1_raised_focus_dots_img.Visible = False
5770      .cmdCalendar1_sunken_focus_dots_img.Visible = False
5780      .cmdCalendar1_raised_img_dis.Visible = False
5790    End With

EXITP:
5800    Exit Sub

ERRH:
5810    Select Case ERR.Number
        Case Else
5820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5830    End Select
5840    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

5910    With Me
5920      blnCalendar1_MouseDown = True
5930      .cmdCalendar1_sunken_focus_dots_img.Visible = True
5940      .cmdCalendar1_raised_img.Visible = False
5950      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5960      .cmdCalendar1_raised_focus_img.Visible = False
5970      .cmdCalendar1_raised_focus_dots_img.Visible = False
5980      .cmdCalendar1_raised_img_dis.Visible = False
5990    End With

EXITP:
6000    Exit Sub

ERRH:
6010    Select Case ERR.Number
        Case Else
6020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6030    End Select
6040    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

6110    With Me
6120      datStartDate = Date
6130      datEndDate = 0
6140      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
6150      If blnRetVal = True Then
6160        .DateStart = datStartDate
6170      Else
6180        .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
6190      End If
6200      .DateStart.SetFocus
6210    End With

EXITP:
6220    Exit Sub

ERRH:
6230    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
6240    Case Else
6250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6260    End Select
6270    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

6310    With Me
6320      If blnCalendar1_MouseDown = False Then
6330        Select Case blnCalendar1_Focus
            Case True
6340          .cmdCalendar1_raised_focus_dots_img.Visible = True
6350          .cmdCalendar1_raised_focus_img.Visible = False
6360        Case False
6370          .cmdCalendar1_raised_focus_img.Visible = True
6380          .cmdCalendar1_raised_focus_dots_img.Visible = False
6390        End Select
6400        .cmdCalendar1_raised_img.Visible = False
6410        .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6420        .cmdCalendar1_sunken_focus_dots_img.Visible = False
6430        .cmdCalendar1_raised_img_dis.Visible = False
6440      End If
6450    End With

EXITP:
6460    Exit Sub

ERRH:
6470    Select Case ERR.Number
        Case Else
6480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6490    End Select
6500    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

6610    With Me
6620      .cmdCalendar1_raised_focus_dots_img.Visible = True
6630      .cmdCalendar1_raised_img.Visible = False
6640      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6650      .cmdCalendar1_raised_focus_img.Visible = False
6660      .cmdCalendar1_sunken_focus_dots_img.Visible = False
6670      .cmdCalendar1_raised_img_dis.Visible = False
6680      blnCalendar1_MouseDown = False
6690    End With

EXITP:
6700    Exit Sub

ERRH:
6710    Select Case ERR.Number
        Case Else
6720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6730    End Select
6740    Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

6810    With Me
6820      .cmdCalendar1_raised_img.Visible = True
6830      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6840      .cmdCalendar1_raised_focus_img.Visible = False
6850      .cmdCalendar1_raised_focus_dots_img.Visible = False
6860      .cmdCalendar1_sunken_focus_dots_img.Visible = False
6870      .cmdCalendar1_raised_img_dis.Visible = False
6880      blnCalendar1_Focus = False
6890    End With

EXITP:
6900    Exit Sub

ERRH:
6910    Select Case ERR.Number
        Case Else
6920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6930    End Select
6940    Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer
        Dim Cancel As Integer

7010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7020    intShiftDown = (Shift And acShiftMask) > 0
7030    intAltDown = (Shift And acAltMask) > 0
7040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7070        With Me
7080          If IsNull(.DateEnd.text) Then
7090            intRetVal = 0
7100            .DateEnd = Format(Date, "mm/dd/yyyy")
7110            Cancel = 0
7120            DateEnd_Exit Cancel  ' ** Procedure: Below.
7130            .opgAccountNumber.SetFocus
7140          Else
7150            If .DateEnd.text = "__/__/____" Or .DateEnd.text = vbNullString Then
7160              intRetVal = 0
7170              .DateEnd = Format(Date, "mm/dd/yyyy")
7180              Cancel = 0
7190              DateEnd_Exit Cancel  ' ** Procedure: Below.
7200              .opgAccountNumber.SetFocus
7210            End If
7220          End If
7230        End With
7240      End Select
7250    End If

        ' ** Shift keys.
7260    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7270      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7280        With Me
7290          intRetVal = 0
7300          .DateStart.SetFocus
7310        End With
7320      End Select
7330    End If

        ' ** Ctrl keys.
7340    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
7350      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7360        With Me
7370          intRetVal = 0
7380          .cmdCalendar2.SetFocus
7390        End With
7400      End Select
7410    End If

        ' ** Ctrl-Shift keys.
7420    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7430      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7440        With Me
7450          intRetVal = 0
7460          .cmdCalendar1.SetFocus
7470        End With
7480      End Select
7490    End If

EXITP:
7500    KeyCode = intRetVal
7510    Exit Sub

ERRH:
7520    intRetVal = 0
7530    THAT_PROC = THIS_PROC
7540    That_Erl = Erl
7550    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7560    Resume EXITP

End Sub

Private Sub DateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_MouseMove"

7610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
7620    Exit Sub

ERRH:
7630    Select Case ERR.Number
        Case Else
7640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7650    End Select
7660    Resume EXITP

End Sub

Private Sub DateEnd_Exit(Cancel As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

7710    With Me
7720      If IsDate(.DateEnd) = True Then
7730        If Not IsDate(.DateStart) Then
7740          MsgBox "The start date must be prior or the same as the end date.", vbInformation + vbOKOnly, "Invalid Date"
7750          .DateStart.SetFocus
7760        Else
7770          If CDate(.DateStart) > CDate(.DateEnd) Then
7780            MsgBox "The start date must be prior or the same as the end date.", vbInformation + vbOKOnly, "Invalid Date"
7790            .DateStart.SetFocus
7800          End If
7810        End If
7820      End If
7830    End With

EXITP:
7840    Exit Sub

ERRH:
7850    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
7860    Case Else
7870      THAT_PROC = THIS_PROC
7880      That_Erl = Erl
7890      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7900    End Select
7910    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

8010    With Me
8020      blnCalendar2_Focus = True
8030      .cmdCalendar2_raised_semifocus_dots_img.Visible = True
8040      .cmdCalendar2_raised_img.Visible = False
8050      .cmdCalendar2_raised_focus_img.Visible = False
8060      .cmdCalendar2_raised_focus_dots_img.Visible = False
8070      .cmdCalendar2_sunken_focus_dots_img.Visible = False
8080      .cmdCalendar2_raised_img_dis.Visible = False
8090    End With

EXITP:
8100    Exit Sub

ERRH:
8110    Select Case ERR.Number
        Case Else
8120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8130    End Select
8140    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

8210    With Me
8220      blnCalendar2_MouseDown = True
8230      .cmdCalendar2_sunken_focus_dots_img.Visible = True
8240      .cmdCalendar2_raised_img.Visible = False
8250      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8260      .cmdCalendar2_raised_focus_img.Visible = False
8270      .cmdCalendar2_raised_focus_dots_img.Visible = False
8280      .cmdCalendar2_raised_img_dis.Visible = False
8290    End With

EXITP:
8300    Exit Sub

ERRH:
8310    Select Case ERR.Number
        Case Else
8320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8330    End Select
8340    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

8410    With Me
8420      datStartDate = Date
8430      datEndDate = 0
8440      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
8450      If blnRetVal = True Then
8460        .DateEnd = datStartDate
8470      Else
8480        .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
8490      End If
8500      .DateEnd.SetFocus
8510      Cancel = 0
8520      DateEnd_Exit Cancel  ' ** Procedure: Above.
8530      If Cancel = 0 Then
8540        .cmdPreview.SetFocus
8550      End If
8560    End With

EXITP:
8570    Exit Sub

ERRH:
8580    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
8590    Case Else
8600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8610    End Select
8620    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

8710    With Me
8720      If blnCalendar2_MouseDown = False Then
8730        Select Case blnCalendar2_Focus
            Case True
8740          .cmdCalendar2_raised_focus_dots_img.Visible = True
8750          .cmdCalendar2_raised_focus_img.Visible = False
8760        Case False
8770          .cmdCalendar2_raised_focus_img.Visible = True
8780          .cmdCalendar2_raised_focus_dots_img.Visible = False
8790        End Select
8800        .cmdCalendar2_raised_img.Visible = False
8810        .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8820        .cmdCalendar2_sunken_focus_dots_img.Visible = False
8830        .cmdCalendar2_raised_img_dis.Visible = False
8840      End If
8850    End With

EXITP:
8860    Exit Sub

ERRH:
8870    Select Case ERR.Number
        Case Else
8880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8890    End Select
8900    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

9010    With Me
9020      .cmdCalendar2_raised_focus_dots_img.Visible = True
9030      .cmdCalendar2_raised_img.Visible = False
9040      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
9050      .cmdCalendar2_raised_focus_img.Visible = False
9060      .cmdCalendar2_sunken_focus_dots_img.Visible = False
9070      .cmdCalendar2_raised_img_dis.Visible = False
9080      blnCalendar2_MouseDown = False
9090    End With

EXITP:
9100    Exit Sub

ERRH:
9110    Select Case ERR.Number
        Case Else
9120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9130    End Select
9140    Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

9210    With Me
9220      .cmdCalendar2_raised_img.Visible = True
9230      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
9240      .cmdCalendar2_raised_focus_img.Visible = False
9250      .cmdCalendar2_raised_focus_dots_img.Visible = False
9260      .cmdCalendar2_sunken_focus_dots_img.Visible = False
9270      .cmdCalendar2_raised_img_dis.Visible = False
9280      blnCalendar2_Focus = False
9290    End With

EXITP:
9300    Exit Sub

ERRH:
9310    Select Case ERR.Number
        Case Else
9320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9330    End Select
9340    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

9410    With Me
9420      Select Case .chkRememberDates
          Case True
9430        .chkRememberDates_lbl.FontBold = True
9440        .chkRememberDates_lbl2_dim.FontBold = True
9450        .chkRememberDates_lbl2_dim_hi.FontBold = True
9460      Case False
9470        .chkRememberDates_lbl.FontBold = False
9480        .chkRememberDates_lbl2_dim.FontBold = False
9490        .chkRememberDates_lbl2_dim_hi.FontBold = False
9500      End Select
9510    End With

EXITP:
9520    Exit Sub

ERRH:
9530    Select Case ERR.Number
        Case Else
9540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9550    End Select
9560    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9620    intShiftDown = (Shift And acShiftMask) > 0
9630    intAltDown = (Shift And acAltMask) > 0
9640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9670        With Me
9680          intRetVal = 0
9690          .opgAccountNumber.SetFocus
9700        End With
9710      End Select
9720    End If

        ' ** Shift keys.
9730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9750        With Me
9760          intRetVal = 0
9770          .cmdCalendar2.SetFocus
9780        End With
9790      End Select
9800    End If

EXITP:
9810    KeyCode = intRetVal
9820    Exit Sub

ERRH:
9830    intRetVal = 0
9840    THAT_PROC = THIS_PROC
9850    That_Erl = Erl
9860    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9870    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

9910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
9920    Exit Sub

ERRH:
9930    Select Case ERR.Number
        Case Else
9940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9950    End Select
9960    Resume EXITP

End Sub

Private Sub opgAccountNumber_AfterUpdate()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_AfterUpdate"

10010   With Me
10020     Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
10030       .opgAccountNumber_optSpecified_lbl.FontBold = True
10040       .opgAccountNumber_optAll_lbl.FontBold = False
10050       .opgAccountNumber_optSpecified_lbl_box.Visible = True
10060       .opgAccountNumber_optAll_lbl_box.Visible = False
10070       .cmbAccounts.Enabled = True
10080       .cmbAccounts.BorderColor = CLR_LTBLU2
10090       .cmbAccounts.BackStyle = acBackStyleNormal
10100       .opgAccountSource.Enabled = True
10110       .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
10120       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
10130       .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
10140       .opgAccountSource_optName_lbl2_dim_hi.Visible = False
10150       .chkRememberMe.Enabled = True
10160       .chkRememberMe_lbl.Visible = True
10170       .chkRememberMe_lbl2_dim.Visible = False
10180       .chkRememberMe_lbl2_dim_hi.Visible = False
10190     Case .opgAccountNumber_optAll.OptionValue
10200       .opgAccountNumber_optSpecified_lbl.FontBold = False
10210       .opgAccountNumber_optAll_lbl.FontBold = True
10220       .opgAccountNumber_optSpecified_lbl_box.Visible = False
10230       .opgAccountNumber_optAll_lbl_box.Visible = True
10240       .cmbAccounts.Enabled = False
10250       .cmbAccounts.BorderColor = WIN_CLR_DISR
10260       .cmbAccounts.BackStyle = acBackStyleTransparent
10270       .opgAccountSource.Enabled = False
10280       .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
10290       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
10300       .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
10310       .opgAccountSource_optName_lbl2_dim_hi.Visible = True
10320       .chkRememberMe.Enabled = False
10330       .chkRememberMe_lbl.Visible = False
10340       .chkRememberMe_lbl2_dim.Visible = True
10350       .chkRememberMe_lbl2_dim_hi.Visible = True
10360     End Select
10370   End With

EXITP:
10380   Exit Sub

ERRH:
10390   Select Case ERR.Number
        Case Else
10400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10410   End Select
10420   Resume EXITP

End Sub

Private Sub opgAccountNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10520   intShiftDown = (Shift And acShiftMask) > 0
10530   intAltDown = (Shift And acAltMask) > 0
10540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10570       With Me
10580         intRetVal = 0
10590         If .cmbAccounts.Enabled = True Then
10600           .cmbAccounts.SetFocus
10610         Else
10620           .cmdPreview.SetFocus
10630         End If
10640       End With
10650     End Select
10660   End If

        ' ** Shift keys.
10670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10690       With Me
10700         intRetVal = 0
10710         .DateEnd.SetFocus
10720       End With
10730     End Select
10740   End If

        ' ** Ctrl keys.
10750   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10770       With Me
10780         intRetVal = 0
10790         .opgAccountSource.SetFocus
10800       End With
10810     End Select
10820   End If

        ' ** Ctrl-Shift keys.
10830   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10850       With Me
10860         intRetVal = 0
10870         .chkRememberDates.SetFocus
10880       End With
10890     End Select
10900   End If

EXITP:
10910   KeyCode = intRetVal
10920   Exit Sub

ERRH:
10930   intRetVal = 0
10940   THAT_PROC = THIS_PROC
10950   That_Erl = Erl
10960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10970   Resume EXITP

End Sub

Private Sub opgAccountNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11020   intShiftDown = (Shift And acShiftMask) > 0
11030   intAltDown = (Shift And acAltMask) > 0
11040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11070       With Me
11080         intRetVal = 0
11090         If .cmbAccounts.Enabled = True Then
11100           .cmbAccounts.SetFocus
11110         Else
11120           .cmdPreview.SetFocus
11130         End If
11140       End With
11150     End Select
11160   End If

        ' ** Shift keys.
11170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11190       With Me
11200         intRetVal = 0
11210         .DateEnd.SetFocus
11220       End With
11230     End Select
11240   End If

        ' ** Ctrl keys.
11250   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
11260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11270       With Me
11280         intRetVal = 0
11290         .opgAccountSource.SetFocus
11300       End With
11310     End Select
11320   End If

        ' ** Ctrl-Shift keys.
11330   If intCtrlDown And (Not intAltDown) And intShiftDown Then
11340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11350       With Me
11360         intRetVal = 0
11370         .chkRememberDates.SetFocus
11380       End With
11390     End Select
11400   End If

EXITP:
11410   KeyCode = intRetVal
11420   Exit Sub

ERRH:
11430   intRetVal = 0
11440   THAT_PROC = THIS_PROC
11450   That_Erl = Erl
11460   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11470   Resume EXITP

End Sub

Private Sub chkIncludeClosed_AfterUpdate()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeClosed_AfterUpdate"

        Dim blnIsClosed As Boolean, strAccountNo As String

11510   With Me
11520     blnIsClosed = False: strAccountNo = vbNullString
11530     If IsNull(.cmbAccounts) = False Then
11540       strAccountNo = .cmbAccounts
11550       If .cmbAccounts.Column(7) = "X" Then
11560         blnIsClosed = True
11570       End If
11580     End If
11590     Select Case .chkIncludeClosed
          Case True
11600       .chkIncludeClosed_lbl.FontBold = True
11610       Select Case .opgAccountSource
            Case .opgAccountSource_optNumber.OptionValue
11620         If .cmbAccounts.RowSource <> "qryArchiveTransactions_01g" Then
11630           .cmbAccounts.RowSource = "qryArchiveTransactions_01g"
11640         End If
11650       Case .opgAccountSource_optName.OptionValue
11660         If .cmbAccounts.RowSource <> "qryArchiveTransactions_01h" Then
11670           .cmbAccounts.RowSource = "qryArchiveTransactions_01h"
11680         End If
11690       End Select
11700     Case False
11710       .chkIncludeClosed_lbl.FontBold = False
11720       Select Case .opgAccountSource
            Case .opgAccountSource_optNumber.OptionValue
11730         If .cmbAccounts.RowSource <> "qryArchiveTransactions_01e" Then
11740           .cmbAccounts.RowSource = "qryArchiveTransactions_01e"
11750         End If
11760       Case .opgAccountSource_optName.OptionValue
11770         If .cmbAccounts.RowSource <> "qryArchiveTransactions_01f" Then
11780           .cmbAccounts.RowSource = "qryArchiveTransactions_01f"
11790         End If
11800       End Select
11810     End Select
11820     .cmbAccounts.Requery
11830     If strAccountNo <> vbNullString Then
11840       Select Case blnIsClosed
            Case True
11850         Select Case .chkIncludeClosed
              Case True
11860           .cmbAccounts = strAccountNo
11870         Case False
11880           .cmbAccounts = Null
11890         End Select
11900       Case False
11910         .cmbAccounts = strAccountNo
11920       End Select
11930     End If
11940   End With

EXITP:
11950   Exit Sub

ERRH:
11960   Select Case ERR.Number
        Case Else
11970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11980   End Select
11990   Resume EXITP

End Sub

Private Sub chkIncludeClosed_KeyDown(KeyCode As Integer, Shift As Integer)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeClosed_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12020   intShiftDown = (Shift And acShiftMask) > 0
12030   intAltDown = (Shift And acAltMask) > 0
12040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12070       With Me
12080         intRetVal = 0
12090         .chkCombineCash.SetFocus
12100       End With
12110     End Select
12120   End If

        ' ** Shift keys.
12130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12150       With Me
12160         intRetVal = 0
12170         .opgAccountNumber.SetFocus
12180       End With
12190     End Select
12200   End If

EXITP:
12210   KeyCode = intRetVal
12220   Exit Sub

ERRH:
12230   intRetVal = 0
12240   THAT_PROC = THIS_PROC
12250   That_Erl = Erl
12260   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12270   Resume EXITP

End Sub

Private Sub chkCombineCash_AfterUpdate()

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkCombineCash_AfterUpdate"

12310   With Me
12320     Select Case .chkCombineCash
          Case True
12330       .chkCombineCash_lbl.FontBold = True
12340     Case False
12350       .chkCombineCash_lbl.FontBold = False
12360     End Select
12370   End With

EXITP:
12380   Exit Sub

ERRH:
12390   Select Case ERR.Number
        Case Else
12400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12410   End Select
12420   Resume EXITP

End Sub

Private Sub chkCombineCash_KeyDown(KeyCode As Integer, Shift As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkCombineCash_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12520   intShiftDown = (Shift And acShiftMask) > 0
12530   intAltDown = (Shift And acAltMask) > 0
12540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12570       With Me
12580         intRetVal = 0
12590         .chkLegalName.SetFocus
12600       End With
12610     End Select
12620   End If

        ' ** Shift keys.
12630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12650       With Me
12660         intRetVal = 0
12670         .chkIncludeClosed.SetFocus
12680       End With
12690     End Select
12700   End If

EXITP:
12710   KeyCode = intRetVal
12720   Exit Sub

ERRH:
12730   intRetVal = 0
12740   THAT_PROC = THIS_PROC
12750   That_Erl = Erl
12760   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12770   Resume EXITP

End Sub

Private Sub chkLegalName_AfterUpdate()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_AfterUpdate"

12810   With Me
12820     Select Case .chkLegalName
          Case True
12830       .chkLegalName_lbl.FontBold = True
12840     Case False
12850       .chkLegalName_lbl.FontBold = False
12860     End Select
12870   End With

EXITP:
12880   Exit Sub

ERRH:
12890   Select Case ERR.Number
        Case Else
12900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12910   End Select
12920   Resume EXITP

End Sub

Private Sub chkLegalName_KeyDown(KeyCode As Integer, Shift As Integer)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13020   intShiftDown = (Shift And acShiftMask) > 0
13030   intAltDown = (Shift And acAltMask) > 0
13040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13070       With Me
13080         intRetVal = 0
13090         If .opgAccountSource.Enabled = True Then
13100           .opgAccountSource.SetFocus
13110         Else
13120           .opgOrderBy.SetFocus
13130         End If
13140       End With
13150     End Select
13160   End If

        ' ** Shift keys.
13170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13190       With Me
13200         intRetVal = 0
13210         .chkCombineCash.SetFocus
13220       End With
13230     End Select
13240   End If

EXITP:
13250   KeyCode = intRetVal
13260   Exit Sub

ERRH:
13270   intRetVal = 0
13280   THAT_PROC = THIS_PROC
13290   That_Erl = Erl
13300   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13310   Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

13410   strAccountNo = vbNullString

13420   With Me
13430     If IsNull(.cmbAccounts) = False Then
13440       If Len(.cmbAccounts.Column(0)) > 0 Then
13450         strAccountNo = .cmbAccounts.Column(0)
13460       End If
13470     End If
13480     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
13490       .opgAccountSource_optNumber_lbl2.FontBold = True
13500       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
13510       .opgAccountSource_optName_lbl2.FontBold = False
13520       .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
13530       Select Case .chkIncludeClosed
            Case True
13540         .cmbAccounts.RowSource = "qryArchiveTransactions_01g"
13550       Case False
13560         .cmbAccounts.RowSource = "qryArchiveTransactions_01e"
13570       End Select
13580     Case .opgAccountSource_optName.OptionValue
13590       .opgAccountSource_optNumber_lbl2.FontBold = False
13600       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
13610       .opgAccountSource_optName_lbl2.FontBold = True
13620       .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
13630       Select Case .chkIncludeClosed
            Case True
13640         .cmbAccounts.RowSource = "qryArchiveTransactions_01h"
13650       Case False
13660         .cmbAccounts.RowSource = "qryArchiveTransactions_01f"
13670       End Select
13680     End Select
13690     DoEvents
13700     If strAccountNo <> vbNullString Then
13710       .cmbAccounts = strAccountNo
13720     End If
13730   End With

EXITP:
13740   Exit Sub

ERRH:
13750   Select Case ERR.Number
        Case Else
13760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13770   End Select
13780   Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13820   intShiftDown = (Shift And acShiftMask) > 0
13830   intAltDown = (Shift And acAltMask) > 0
13840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13870       With Me
13880         intRetVal = 0
13890         .chkRememberMe.SetFocus
13900       End With
13910     End Select
13920   End If

        ' ** Shift keys.
13930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13950       With Me
13960         intRetVal = 0
13970         .chkLegalName.SetFocus
13980       End With
13990     End Select
14000   End If

EXITP:
14010   KeyCode = intRetVal
14020   Exit Sub

ERRH:
14030   intRetVal = 0
14040   THAT_PROC = THIS_PROC
14050   That_Erl = Erl
14060   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14070   Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14120   intShiftDown = (Shift And acShiftMask) > 0
14130   intAltDown = (Shift And acAltMask) > 0
14140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14170       With Me
14180         intRetVal = 0
14190         .chkRememberMe.SetFocus
14200       End With
14210     End Select
14220   End If

        ' ** Shift keys.
14230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14250       With Me
14260         intRetVal = 0
14270         .chkLegalName.SetFocus
14280       End With
14290     End Select
14300   End If

EXITP:
14310   KeyCode = intRetVal
14320   Exit Sub

ERRH:
14330   intRetVal = 0
14340   THAT_PROC = THIS_PROC
14350   That_Erl = Erl
14360   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14370   Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

14410   With Me
14420     Select Case .chkRememberMe
          Case True
14430       .chkRememberMe_lbl.FontBold = True
14440       .chkRememberMe_lbl2_dim.FontBold = True
14450       .chkRememberMe_lbl2_dim_hi.FontBold = True
14460     Case False
14470       .chkRememberMe_lbl.FontBold = False
14480       .chkRememberMe_lbl2_dim.FontBold = False
14490       .chkRememberMe_lbl2_dim_hi.FontBold = False
14500     End Select
14510   End With

EXITP:
14520   Exit Sub

ERRH:
14530   Select Case ERR.Number
        Case Else
14540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14550   End Select
14560   Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14620   intShiftDown = (Shift And acShiftMask) > 0
14630   intAltDown = (Shift And acAltMask) > 0
14640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14670       With Me
14680         intRetVal = 0
14690         .cmbAccounts.SetFocus
14700       End With
14710     End Select
14720   End If

        ' ** Shift keys.
14730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14750       With Me
14760         intRetVal = 0
14770         .opgAccountSource.SetFocus
14780       End With
14790     End Select
14800   End If

EXITP:
14810   KeyCode = intRetVal
14820   Exit Sub

ERRH:
14830   intRetVal = 0
14840   THAT_PROC = THIS_PROC
14850   That_Erl = Erl
14860   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14870   Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

14910   With Me
14920     If .cmbAccounts.text = vbNullString Then
14930       ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
14940       .cmbAccounts = .cmbAccounts.ItemData(0)
14950     End If
14960   End With

EXITP:
14970   Exit Sub

ERRH:
14980   Select Case ERR.Number
        Case Else
14990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15000   End Select
15010   Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15120   intShiftDown = (Shift And acShiftMask) > 0
15130   intAltDown = (Shift And acAltMask) > 0
15140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15170       With Me
15180         intRetVal = 0
15190         .cmdPreview.SetFocus
15200       End With
15210     End Select
15220   End If

        ' ** Shift keys.
15230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15250       With Me
15260         intRetVal = 0
15270         .opgAccountNumber.SetFocus
15280       End With
15290     End Select
15300   End If

        ' ** Ctrl keys.
15310   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
15320     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15330       With Me
15340         intRetVal = 0
15350         .opgOrderBy.SetFocus
15360       End With
15370     End Select
15380   End If

        ' ** Ctrl-Shift keys.
15390   If intCtrlDown And (Not intAltDown) And intShiftDown Then
15400     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15410       With Me
15420         intRetVal = 0
15430         .chkRememberMe.SetFocus
15440       End With
15450     End Select
15460   End If

EXITP:
15470   KeyCode = intRetVal
15480   Exit Sub

ERRH:
15490   intRetVal = 0
15500   THAT_PROC = THIS_PROC
15510   That_Erl = Erl
15520   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15530   Resume EXITP

End Sub

Private Sub opgOrderBy_AfterUpdate()

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_AfterUpdate"

15610   With Me
15620     Select Case .opgOrderBy
          Case .opgOrderBy_optDate.OptionValue
15630       .opgOrderBy_optDate_lbl.FontBold = True
15640       .opgOrderBy_optType_lbl.FontBold = False
15650     Case .opgOrderBy_optType.OptionValue
15660       .opgOrderBy_optDate_lbl.FontBold = False
15670       .opgOrderBy_optType_lbl.FontBold = True
15680     End Select
15690   End With

EXITP:
15700   Exit Sub

ERRH:
15710   Select Case ERR.Number
        Case Else
15720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15730   End Select
15740   Resume EXITP

End Sub

Private Sub opgOrderBy_optDate_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_optDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15820   intShiftDown = (Shift And acShiftMask) > 0
15830   intAltDown = (Shift And acAltMask) > 0
15840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15870       With Me
15880         intRetVal = 0
15890         .chkPageOf.SetFocus
15900       End With
15910     End Select
15920   End If

        ' ** Shift keys.
15930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15950       With Me
15960         intRetVal = 0
15970         If .cmbAccounts.Enabled = True Then
15980           .cmbAccounts.SetFocus
15990         Else
16000           .chkLegalName.SetFocus
16010         End If
16020       End With
16030     End Select
16040   End If

EXITP:
16050   KeyCode = intRetVal
16060   Exit Sub

ERRH:
16070   intRetVal = 0
16080   THAT_PROC = THIS_PROC
16090   That_Erl = Erl
16100   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16110   Resume EXITP

End Sub

Private Sub opgOrderBy_optType_KeyDown(KeyCode As Integer, Shift As Integer)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_optType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16220   intShiftDown = (Shift And acShiftMask) > 0
16230   intAltDown = (Shift And acAltMask) > 0
16240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16270       With Me
16280         intRetVal = 0
16290         .chkPageOf.SetFocus
16300       End With
16310     End Select
16320   End If

        ' ** Shift keys.
16330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16350       With Me
16360         intRetVal = 0
16370         If .cmbAccounts.Enabled = True Then
16380           .cmbAccounts.SetFocus
16390         Else
16400           .chkLegalName.SetFocus
16410         End If
16420       End With
16430     End Select
16440   End If

EXITP:
16450   KeyCode = intRetVal
16460   Exit Sub

ERRH:
16470   intRetVal = 0
16480   THAT_PROC = THIS_PROC
16490   That_Erl = Erl
16500   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16510   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

16610   With Me
16620     Select Case .chkPageOf
          Case True
16630       .chkPageOf_lbl2.FontBold = True
16640       .chkPageOf_lbl2_dim_hi.FontBold = True
16650     Case False
16660       .chkPageOf_lbl2.FontBold = False
16670       .chkPageOf_lbl2_dim_hi.FontBold = False
16680     End Select
16690   End With

EXITP:
16700   Exit Sub

ERRH:
16710   Select Case ERR.Number
        Case Else
16720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16730   End Select
16740   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

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
16890         .opgReturn.SetFocus
16900       End With
16910     End Select
16920   End If

        ' ** Shift keys.
16930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16950       With Me
16960         intRetVal = 0
16970         .opgOrderBy.SetFocus
16980       End With
16990     End Select
17000   End If

EXITP:
17010   KeyCode = intRetVal
17020   Exit Sub

ERRH:
17030   intRetVal = 0
17040   THAT_PROC = THIS_PROC
17050   That_Erl = Erl
17060   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17070   Resume EXITP

End Sub

Private Sub opgReturn_AfterUpdate()

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgReturn_AfterUpdate"

17110   With Me
17120     Select Case .opgReturn
          Case .opgReturn_optSelect.OptionValue
17130       .opgReturn_optSelect_lbl.FontBold = True
17140       .opgReturn_optSelect_lbl2.FontBold = True
17150       .opgReturn_optMenu_lbl.FontBold = False
17160       .opgReturn_optMenu_lbl2.FontBold = False
17170     Case .opgReturn_optMenu.OptionValue
17180       .opgReturn_optSelect_lbl.FontBold = False
17190       .opgReturn_optSelect_lbl2.FontBold = False
17200       .opgReturn_optMenu_lbl.FontBold = True
17210       .opgReturn_optMenu_lbl2.FontBold = True
17220     End Select
17230   End With

EXITP:
17240   Exit Sub

ERRH:
17250   Select Case ERR.Number
        Case Else
17260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17270   End Select
17280   Resume EXITP

End Sub

Private Sub opgReturn_optSelect_KeyDown(KeyCode As Integer, Shift As Integer)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgReturn_optSelect_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17320   intShiftDown = (Shift And acShiftMask) > 0
17330   intAltDown = (Shift And acAltMask) > 0
17340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17370       With Me
17380         intRetVal = 0
17390         .cmdPreview.SetFocus
17400       End With
17410     End Select
17420   End If

        ' ** Shift keys.
17430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17450       With Me
17460         intRetVal = 0
17470         .chkPageOf.SetFocus
17480       End With
17490     End Select
17500   End If

EXITP:
17510   KeyCode = intRetVal
17520   Exit Sub

ERRH:
17530   intRetVal = 0
17540   THAT_PROC = THIS_PROC
17550   That_Erl = Erl
17560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17570   Resume EXITP

End Sub

Private Sub opgReturn_optMenu_KeyDown(KeyCode As Integer, Shift As Integer)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgReturn_optMenu_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17620   intShiftDown = (Shift And acShiftMask) > 0
17630   intAltDown = (Shift And acAltMask) > 0
17640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17670       With Me
17680         intRetVal = 0
17690         .cmdPreview.SetFocus
17700       End With
17710     End Select
17720   End If

        ' ** Shift keys.
17730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17750       With Me
17760         intRetVal = 0
17770         .chkPageOf.SetFocus
17780       End With
17790     End Select
17800   End If

EXITP:
17810   KeyCode = intRetVal
17820   Exit Sub

ERRH:
17830   intRetVal = 0
17840   THAT_PROC = THIS_PROC
17850   That_Erl = Erl
17860   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17870   Resume EXITP

End Sub

Private Sub cmdPreview_Click()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

17910   If DoReport = True Then  ' ** Procedure: Below.
17920     strDocName = "rptArchivedTransactions"
17930     DoCmd.OpenReport strDocName, acViewPreview
17940     DoCmd.Maximize
17950     DoCmd.RunCommand acCmdFitToWindow
17960   End If

EXITP:
17970   Exit Sub

ERRH:
17980   THAT_PROC = THIS_PROC
17990   That_Erl = Erl
18000   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18010   Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18120   intShiftDown = (Shift And acShiftMask) > 0
18130   intAltDown = (Shift And acAltMask) > 0
18140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18160     Select Case intRetVal
          Case vbKeyTab
18170       With Me
18180         intRetVal = 0
18190         .cmdPrint.SetFocus
18200       End With
18210     End Select
18220   End If

        ' ** Shift keys.
18230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18240     Select Case intRetVal
          Case vbKeyTab
18250       With Me
18260         intRetVal = 0
18270         If .cmbAccounts.Enabled = True Then
18280           .cmbAccounts.SetFocus
18290         Else
18300           .opgAccountNumber.SetFocus
18310         End If
18320       End With
18330     End Select
18340   End If

        ' ** Ctrl-Shift keys.
18350   If intCtrlDown And (Not intAltDown) And intShiftDown Then
18360     Select Case intRetVal
          Case vbKeyTab
18370       With Me
18380         intRetVal = 0
18390         .chkPageOf.SetFocus
18400       End With
18410     End Select
18420   End If

EXITP:
18430   KeyCode = intRetVal
18440   Exit Sub

ERRH:
18450   intRetVal = 0
18460   THAT_PROC = THIS_PROC
18470   That_Erl = Erl
18480   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18490   Resume EXITP

End Sub

Private Sub cmdPrint_Click()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

18510   If DoReport = True Then  ' ** Procedure: Below.
18520     strDocName = "rptArchivedTransactions"
18530     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
18540       DoCmd.OpenReport strDocName, acViewPreview
18550       DoCmd.Maximize
18560       DoCmd.RunCommand acCmdFitToWindow
18570     Else
            '##GTR_Ref: rptArchivedTransactions
18580       DoCmd.OpenReport strDocName, acViewNormal
18590     End If
18600   End If

EXITP:
18610   Exit Sub

ERRH:
18620   THAT_PROC = THIS_PROC
18630   That_Erl = Erl
18640   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18650   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18720   intShiftDown = (Shift And acShiftMask) > 0
18730   intAltDown = (Shift And acAltMask) > 0
18740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18760     Select Case intRetVal
          Case vbKeyTab
18770       With Me
18780         intRetVal = 0
18790         .cmdWord.SetFocus
18800       End With
18810     End Select
18820   End If

        ' ** Shift keys.
18830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18840     Select Case intRetVal
          Case vbKeyTab
18850       With Me
18860         intRetVal = 0
18870         .cmdPreview.SetFocus
18880       End With
18890     End Select
18900   End If

EXITP:
18910   KeyCode = intRetVal
18920   Exit Sub

ERRH:
18930   intRetVal = 0
18940   THAT_PROC = THIS_PROC
18950   That_Erl = Erl
18960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18970   Resume EXITP

End Sub

Private Sub cmdWord_Click()

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String, strRptAcct As String
        Dim strRptPath As String, strRptPathFile As String

19010   With Me
19020     If DoReport = True Then  ' ** Procedure: Below.

19030       If IsNull(.UserReportPath) = True Then
19040         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
19050       Else
19060         strRptPath = .UserReportPath
19070       End If

19080       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
19090         strRpt = "rptArchivedTransactions"
19100         strRptAcct = .cmbAccounts
19110         strRptCap = "Archived_Transactions_" & strRptAcct & "_" & _
                Format(.DateStart, "yymmdd") & "_To_" & Format(.DateEnd, "yymmdd")
19120       Case .opgAccountNumber_optAll.OptionValue
19130         strRpt = "rptArchivedTransactions"
19140         strRptCap = "Archived_Transactions_All_" & _
                Format(.DateStart, "yymmdd") & "_To_" & Format(.DateEnd, "yymmdd")
19150       End Select

19160       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

19170       If strRptPathFile <> vbNullString Then
19180         DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
19190         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
19200       End If

19210     End If
19220   End With

EXITP:
19230   Exit Sub

ERRH:
19240   Select Case ERR.Number
        Case Else
19250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19260   End Select
19270   Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19320   intShiftDown = (Shift And acShiftMask) > 0
19330   intAltDown = (Shift And acAltMask) > 0
19340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19360     Select Case intRetVal
          Case vbKeyTab
19370       With Me
19380         intRetVal = 0
19390         If .cmdExcel.Enabled = True Then
19400           .cmdExcel.SetFocus
19410         Else
19420           .cmdClose.SetFocus
19430         End If
19440       End With
19450     End Select
19460   End If

        ' ** Shift keys.
19470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19480     Select Case intRetVal
          Case vbKeyTab
19490       With Me
19500         intRetVal = 0
19510         .cmdPrint.SetFocus
19520       End With
19530     End Select
19540   End If

EXITP:
19550   KeyCode = intRetVal
19560   Exit Sub

ERRH:
19570   intRetVal = 0
19580   THAT_PROC = THIS_PROC
19590   That_Erl = Erl
19600   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19610   Resume EXITP

End Sub

Private Sub cmdExcel_Click()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String, strRptAcct As String
        Dim strRptPath As String, strRptPathFile As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

19710   With Me

19720     DoCmd.Hourglass True
19730     DoEvents

19740     blnContinue = True

19750     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
19760       DoCmd.Hourglass False
19770       msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
19780       If msgResponse <> vbRetry Then
19790         blnContinue = False
19800       End If
19810     End If

19820     If blnContinue = True Then

19830       DoCmd.Hourglass True
19840       DoEvents

19850       If DoReport = True Then  ' ** Procedure: Below.

19860         blnContinue = True

19870         If IsNull(.UserReportPath) = True Then
19880           strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
19890         Else
19900           strRptPath = .UserReportPath
19910         End If

19920         Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
                ' ** qryRpt_ArchivedTransactions_33_01_01_01 (qryRpt_ArchivedTransactions_32_01_01_01
                ' ** (Union of qryRpt_ArchivedTransactions_22_01_01_01 (xx) - qryRpt_ArchivedTransactions_30_01_01 (xx),
                ' ** shortname, icash/pcash, w/o closed), sorted, shortname, icash/pcash, w/o closed),
                ' ** For Export, shortname, icash/pcash, w/o closed, by date.
19930           strQry = "qryRpt_ArchivedTransactions_34_"
19940           Select Case .chkLegalName
                Case True
19950             strQry = strQry & "02_"
19960           Case False
19970             strQry = strQry & "01_"
19980           End Select
19990           Select Case .chkCombineCash
                Case True
20000             strQry = strQry & "02_"
20010           Case False
20020             strQry = strQry & "01_"
20030           End Select
20040           Select Case .chkIncludeClosed
                Case True
20050             strQry = strQry & "02_"
20060           Case False
20070             strQry = strQry & "01_"
20080           End Select
20090           Select Case .opgOrderBy
                Case .opgOrderBy_optDate.OptionValue
20100             strQry = strQry & "01"
20110           Case .opgOrderBy_optType.OptionValue
20120             strQry = strQry & "02"
20130           End Select
20140           strRptAcct = .cmbAccounts
20150           strRptCap = "Archived_Transactions_" & strRptAcct & "_" & _
                  Format(.DateStart, "yymmdd") & "_To_" & Format(.DateEnd, "yymmdd")
20160         Case .opgAccountNumber_optAll.OptionValue
                ' ** qryRpt_ArchivedTransactions_53_01_01_01 (qryRpt_ArchivedTransactions_52_01_01_01
                ' ** (Union of qryRpt_ArchivedTransactions_42_01_01_01 (xx) - qryRpt_ArchivedTransactions_50_01_01 (xx),
                ' ** shortname, icash/pcash, w/o closed), sorted, shortname, icash/pcash, w/o closed),
                ' ** For Export, shortname, icash/pcash, w/o closed, by date.
20170           strQry = "qryRpt_ArchivedTransactions_54_"
20180           Select Case .chkLegalName
                Case True
20190             strQry = strQry & "02_"
20200           Case False
20210             strQry = strQry & "01_"
20220           End Select
20230           Select Case .chkCombineCash
                Case True
20240             strQry = strQry & "02_"
20250           Case False
20260             strQry = strQry & "01_"
20270           End Select
20280           Select Case .chkIncludeClosed
                Case True
20290             strQry = strQry & "02_"
20300           Case False
20310             strQry = strQry & "01_"
20320           End Select
20330           Select Case .opgOrderBy
                Case .opgOrderBy_optDate.OptionValue
20340             strQry = strQry & "01"
20350           Case .opgOrderBy_optType.OptionValue
20360             strQry = strQry & "02"
20370           End Select
20380           strRptAcct = vbNullString
20390           strRptCap = "Archived_Transactions_All_" & _
                  Format(.DateStart, "yymmdd") & "_To_" & Format(.DateEnd, "yymmdd")
20400         End Select

20410         strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

20420         If strRptPathFile <> vbNullString Then
20430           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
20440             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
20450           End If
20460           DoEvents
20470           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
20480             Kill strRptPathFile
20490           End If
20500           DoEvents
20510           gstrFormQuerySpec = THIS_NAME
20520           DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
20530           If Excel_NameOnly(strRptPathFile, "Archived") = True Then  ' ** Module Function: modExcelFuncs.
20540             If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
20550               EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
20560             End If
20570             OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
20580           End If
20590           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
20600         End If

20610       End If  ' ** DoReport().
20620     End If  ' ** blnContinue.

20630     DoCmd.Hourglass False

20640   End With

EXITP:
20650   Exit Sub

ERRH:
20660   DoCmd.Hourglass False
20670   Select Case ERR.Number
        Case Else
20680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20690   End Select
20700   Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20820   intShiftDown = (Shift And acShiftMask) > 0
20830   intAltDown = (Shift And acAltMask) > 0
20840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20860     Select Case intRetVal
          Case vbKeyTab
20870       With Me
20880         intRetVal = 0
20890         .cmdClose.SetFocus
20900       End With
20910     End Select
20920   End If

        ' ** Shift keys.
20930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20940     Select Case intRetVal
          Case vbKeyTab
20950       With Me
20960         intRetVal = 0
20970         .cmdWord.SetFocus
20980       End With
20990     End Select
21000   End If

EXITP:
21010   KeyCode = intRetVal
21020   Exit Sub

ERRH:
21030   intRetVal = 0
21040   THAT_PROC = THIS_PROC
21050   That_Erl = Erl
21060   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
21070   Resume EXITP

End Sub

Private Function DoReport() As Boolean

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

21110   blnRetVal = True

21120   With Me
21130     If IsNull(.DateStart) Or IsNull(.DateEnd) Then
21140       blnRetVal = False
21150       MsgBox "You must enter appropriate dates to continue.", vbInformation + vbOKOnly, "Entry Required"
21160       .DateStart.SetFocus
21170     Else
21180       If .cmbAccounts.Enabled = True Then
21190         .cmbAccounts.SetFocus
21200         If .cmbAccounts.text = vbNullString Then
21210           blnRetVal = False
21220           MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
21230         End If
21240       End If
21250     End If

21260     If blnRetVal = True Then
21270       ChkSpecLedgerEntry  ' ** Module Function: modUtilities.
            ' ** qryRpt_ArchivedTransactions_01 gives us the balance numbers from the previous statement,
            ' ** and now uses FormRef() rather than rewriting code here.
21280     End If

21290     If blnRetVal = True Then
21300       Set dbs = CurrentDb
21310       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
21320         Set qdf = dbs.QueryDefs("qryRpt_ArchivedTransactions_02")
21330       Case .opgAccountNumber_optAll.OptionValue
21340         Set qdf = dbs.QueryDefs("qryRpt_ArchivedTransactions_03")
21350       End Select
21360       Set rst = qdf.OpenRecordset
21370       With rst
21380         If .BOF = True And .EOF = True Then
21390           blnRetVal = False
21400           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
21410         Else
                ' ** Report set to qryRpt_ArchivedTransactions_02, Report_Open()
                ' ** flips it to qryRpt_ArchivedTransactions_03 if All is chosen.
21420         End If
21430         .Close
21440       End With
21450       dbs.Close
21460     End If

21470   End With

EXITP:
21480   Set rst = Nothing
21490   Set qdf = Nothing
21500   Set dbs = Nothing
21510   DoReport = blnRetVal
21520   Exit Function

ERRH:
21530   blnRetVal = False
21540   THAT_PROC = THIS_PROC
21550   That_Erl = Erl
21560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
21570   Resume EXITP

End Function

Private Sub SetClosedAccts()

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "SetClosedAccts"

        Dim varTmp00 As Variant

21610   With Me
          ' ** Account, just closed accounts.
21620     varTmp00 = DCount("*", "qryRpt_ArchivedTransactions_05")
21630     If IsNull(varTmp00) = True Then
21640       .chkIncludeClosed = False
21650       .chkIncludeClosed.Enabled = False
21660       .chkIncludeClosed_lbl.Visible = False
21670       .chkIncludeClosed_lbl2_dim.Visible = True
21680       .chkIncludeClosed_lbl2_dim_hi.Visible = True
21690     Else
21700       If varTmp00 = 0 Then
21710         .chkIncludeClosed = False
21720         .chkIncludeClosed.Enabled = False
21730         .chkIncludeClosed_lbl.Visible = False
21740         .chkIncludeClosed_lbl2_dim.Visible = True
21750         .chkIncludeClosed_lbl2_dim_hi.Visible = True
21760       Else
21770         .chkIncludeClosed.Enabled = True
21780         .chkIncludeClosed_lbl.Visible = True
21790         .chkIncludeClosed_lbl2_dim.Visible = False
21800         .chkIncludeClosed_lbl2_dim_hi.Visible = False
21810       End If
21820     End If
21830   End With

EXITP:
21840   Exit Sub

ERRH:
21850   Select Case ERR.Number
        Case Else
21860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21870   End Select
21880   Resume EXITP

End Sub
