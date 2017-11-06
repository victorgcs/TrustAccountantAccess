VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccruedIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccruedIncome"

'VGC 04/15/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   HideIncExp_Purch()
' ##   HideIncExp_Int()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Tax Codes Purch:  A {cmbTaxCodes_Purchase}
' **   Close:            C {cmdClose}
' **   Posting Date:     D {AccrualDate}
' **   I/E Codes Purch:  E {cmbRevenueCodes_Purchase}
' **   Copy First:       F {cmdCopyFactor}
' **   Generate Dates:   G {cmdGenerateDates}
' **   Create Jrnl Ent:  J {cmdCreateEntries}
' **   Daily Factor:     L {tmpRate on frmAccruedIncome_Sub}
' **   Month:            M {AccrualMonth}
' **   I/E Codes Int:    N {cmbRevenueCodes_Interest}
' **   Comments:         O {ReinvestComments}
' **   Reinvest:         R {chkReinvest}
' **   Per Share:        S {ReinvestPerShare}
' **   Tax Codes Int:    T {cmbTaxCodes_Interest}
' **   Exit:             X {cmdClose}
' **   Year:             Y {AccrualYear}

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmAccruedIncome_Sub}

' ** CLOSES WHEN FINISHED!

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean

Private datStartDate As Date, datEndDate As Date
Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngRevID_Interest As Long, lngTaxID_Interest As Long, lngRevID_Purchase As Long, lngTaxID_Purchase As Long
Private strRevCode_Interest As String, strTaxCode_Interest As String, strRevCode_Purchase As String, strTaxCode_Purchase As String
Private blnFromCalendar As Boolean, blnKeyDown_Return As Boolean, blnHideIncExp_Purch As Boolean, blnHideIncExp_Int As Boolean
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim lngJrnlCnt As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strCallingForm = Trim(.OpenArgs)
150           intPos01 = InStr(strCallingForm, "~")
160           If intPos01 > 0 Then

170             blnIsOpen = True

180             DoCmd.Hourglass True
190             DoEvents

200             lngJrnlCnt = Val(Mid(strCallingForm, (intPos01 + 1)))
210             strCallingForm = Left(strCallingForm, (intPos01 - 1))

220             SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.
                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
230             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

240             CLR_DISABLED_FG = CLR_DKGRY
250             CLR_DISABLED_BG = CLR_LTTEAL

260             DoCmd.Hourglass True
270             DoEvents

280             blnFromCalendar = False: blnKeyDown_Return = False: blnHideIncExp_Purch = False: blnHideIncExp_Int = False
290             blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
300             lngRevID_Interest = 0&: lngTaxID_Interest = 0&: lngRevID_Purchase = 0&: lngTaxID_Purchase = 0&
310             strRevCode_Interest = vbNullString: strTaxCode_Interest = vbNullString
320             strRevCode_Purchase = vbNullString: strTaxCode_Purchase = vbNullString

                ' ** Variables are fed empty, then populated ByRef.
330             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

340             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

350             HideIncExp_Purch  ' ** Procedure: Below.
360             HideIncExp_Int  ' ** Procedure: Below.

370             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
380             If lngMonitorCnt > 1& Then
390               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
400               lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
410               If lngMonitorNum > 1& Then
420                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
430               End If
440             End If

450             DoCmd.Hourglass True
460             DoEvents

470             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
480             DoEvents

                ' ** Save preferences for later.
490             lngRevID_Interest = Nz(.cmbRevenueCodes_Interest, 0&)
500             If lngRevID_Interest > 2& Then
510               strRevCode_Interest = .cmbRevenueCodes_Interest.Column(1)
520             End If
530             lngTaxID_Interest = Nz(.cmbTaxCodes_Interest, 0&)
540             If lngTaxID_Interest > 2& Then
550               strTaxCode_Interest = .cmbTaxCodes_Interest.Column(1)
560             End If
570             lngRevID_Purchase = Nz(.cmbRevenueCodes_Purchase, 0&)
580             If lngRevID_Purchase > 2& Then
590               strRevCode_Purchase = .cmbRevenueCodes_Purchase.Column(1)
600             End If
610             lngTaxID_Purchase = Nz(.cmbTaxCodes_Purchase, 0&)
620             If lngTaxID_Purchase > 2& Then
630               strTaxCode_Purchase = .cmbTaxCodes_Purchase.Column(1)
640             End If

650             DoCmd.Hourglass True
660             DoEvents

                '.AccrualMonth = Format(DateAdd("m", -1, Date), "mmmm")
                ' ** Accrual Year had DefaultValue : =Year(Date())
670             .cmdCopyFactor.Enabled = False
680             .AccrualDate.Enabled = False
690             .AccrualDate.BorderColor = WIN_CLR_DISR
700             .AccrualDate.BackStyle = acBackStyleTransparent
710             .cmdCalendar.Enabled = False
720             .cmdCalendar_raised_img_dis.Visible = True
730             .cmdCalendar_raised_img.Visible = False
740             .cmdCalendar_raised_semifocus_dots_img.Visible = False
750             .cmdCalendar_raised_focus_img.Visible = False
760             .cmdCalendar_raised_focus_dots_img.Visible = False
770             .cmdCalendar_sunken_focus_dots_img.Visible = False
780             If blnHideIncExp_Int = False Then
790               .cmbRevenueCodes_Interest_box_lbl.ForeColor = WIN_CLR_DISF
800               .cmbRevenueCodes_Interest_box_lbl_dim_hi.Visible = True
810               .cmbRevenueCodes_Interest_display = Null
820               .cmbRevenueCodes_Interest_display.Enabled = False
830               .cmbRevenueCodes_Interest_display.Locked = False
840               .cmbRevenueCodes_Interest = Null
850               .cmbRevenueCodes_Interest.Enabled = False
860               .cmbRevenueCodes_Interest.BorderColor = WIN_CLR_DISR
870               .cmbRevenueCodes_Interest_lbl.BackStyle = acBackStyleTransparent
880               .cmbTaxCodes_Interest_display = Null
890               .cmbTaxCodes_Interest_display.Enabled = False
900               .cmbTaxCodes_Interest_display.Locked = False
910               .cmbTaxCodes_Interest = Null
920               .cmbTaxCodes_Interest.Enabled = False
930               .cmbTaxCodes_Interest.BorderColor = WIN_CLR_DISR
940               .cmbTaxCodes_Interest_lbl.BackStyle = acBackStyleTransparent
950             End If  ' ** blnHideIncExp_Int.
960             .chkReinvest.Enabled = False
970             .ReinvestPerShare.Visible = False  ' ** Always $1, no option for user.
980             .ReinvestPerShare = 1@
990             .ReinvestComments.Enabled = False
1000            .ReinvestComments.BorderColor = WIN_CLR_DISR
1010            .ReinvestComments.BackStyle = acBackStyleTransparent
1020            If blnHideIncExp_Purch = False Then
1030              .cmbRevenueCodes_Purchase_box_lbl.ForeColor = WIN_CLR_DISF
1040              .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Visible = True
1050              .cmbRevenueCodes_Purchase_display = Null
1060              .cmbRevenueCodes_Purchase_display.Enabled = False
1070              .cmbRevenueCodes_Purchase_display.Locked = False
1080              .cmbRevenueCodes_Purchase = Null
1090              .cmbRevenueCodes_Purchase.Enabled = False
1100              .cmbRevenueCodes_Purchase.BorderColor = WIN_CLR_DISR
1110              .cmbRevenueCodes_Purchase_lbl.BackStyle = acBackStyleTransparent
1120              .cmbTaxCodes_Purchase_display = Null
1130              .cmbTaxCodes_Purchase_display.Enabled = False
1140              .cmbTaxCodes_Purchase_display.Locked = False
1150              .cmbTaxCodes_Purchase = Null
1160              .cmbTaxCodes_Purchase.Enabled = False
1170              .cmbTaxCodes_Purchase.BorderColor = WIN_CLR_DISR
1180              .cmbTaxCodes_Purchase_lbl.BackStyle = acBackStyleTransparent
1190            End If  ' ** blnHideIncExp_Purch.
1200            .cmdCreateEntries.Enabled = False
1210            .frmAccruedIncome_Sub.Form.tmpDays_lbl1.Caption = "0"
1220            .frmAccruedIncome_Sub.Form.tmpDays_lbl1.ForeColor = 10070188
1230            .frmAccruedIncome_Sub.Form.tmpDays_lbl2.ForeColor = 10070188
1240            .frmAccruedIncome_Sub.Form.tmpDate_lbl.Visible = False
1250            .frmAccruedIncome_Sub.Form.tmpDate_lbl_dim.Visible = True
1260            .frmAccruedIncome_Sub.Form.tmpDate_lbl_dim_hi.Visible = True
1270            .frmAccruedIncome_Sub.Form.tmpDate_lbl_line.BorderColor = WIN_CLR_DISR
1280            .frmAccruedIncome_Sub.Form.tmpDate_lbl_line_dim_hi.Visible = True
1290            .frmAccruedIncome_Sub.Form.tmpRate_lbl.Visible = False
1300            .frmAccruedIncome_Sub.Form.tmpRate_lbl_dim.Visible = True
1310            .frmAccruedIncome_Sub.Form.tmpRate_lbl_dim_hi.Visible = True
1320            .frmAccruedIncome_Sub.Form.tmpRate_lbl_line.BorderColor = WIN_CLR_DISR
1330            .frmAccruedIncome_Sub.Form.tmpRate_lbl_line_dim_hi.Visible = True
1340            If lngJrnlCnt <> 0& Then
1350              .cmdClose.SetFocus
1360              .NoAccrued_Journal_lbl.Visible = True
1370              .AccrualMonth.Enabled = False
1380              .AccrualMonth.BorderColor = WIN_CLR_DISR
1390              .AccrualMonth.BackStyle = acBackStyleTransparent
1400              .AccrualYear.Enabled = False
1410              .AccrualYear.BorderColor = WIN_CLR_DISR
1420              .AccrualYear.BackStyle = acBackStyleTransparent
1430              .cmdGenerateDates.Enabled = False
1440            End If

1450            DoCmd.Hourglass True
1460            DoEvents

1470            Select Case gblnRevenueExpenseTracking
                Case True
1480              If blnHideIncExp_Int = False Then
1490                .cmbRevenueCodes_Interest.Visible = True
1500                .cmbRevenueCodes_Interest_lbl2.Visible = False  ' ** Option off msg.
1510              End If  ' ** blnHideIncExp_Int.
1520              If blnHideIncExp_Purch = False Then
1530                .cmbRevenueCodes_Purchase.Visible = True
1540                .cmbRevenueCodes_Purchase_lbl2.Visible = False  ' ** Option off msg.
1550              End If  ' ** blnHideIncExp_Purch.
1560            Case False
1570              If blnHideIncExp_Int = False Then
1580                .cmbRevenueCodes_Interest.Visible = False
1590                .cmbRevenueCodes_Interest_display.Visible = False
1600                .cmbRevenueCodes_Interest_lbl2.Visible = True  ' ** Option off msg.
1610              End If  ' ** blnHideIncExp_Int.
1620              If blnHideIncExp_Purch = False Then
1630                .cmbRevenueCodes_Purchase.Visible = False
1640                .cmbRevenueCodes_Purchase_display.Visible = False
1650                .cmbRevenueCodes_Purchase_lbl2.Visible = True  ' ** Option off msg.
1660              End If  ' ** blnHideIncExp_Purch.
1670            End Select

1680            Select Case gblnIncomeTaxCoding
                Case True
1690              If blnHideIncExp_Int = False Then
1700                .cmbTaxCodes_Interest.Visible = True
1710                .cmbTaxCodes_Interest_lbl2.Visible = False  ' ** Option off msg.
1720              End If  ' ** blnHideIncExp_Int.
1730              If blnHideIncExp_Purch = False Then
1740                .cmbTaxCodes_Purchase.Visible = True
1750                .cmbTaxCodes_Purchase_lbl2.Visible = False  ' ** Option off msg.
1760              End If  ' ** blnHideIncExp_Purch.
1770            Case False
1780              If blnHideIncExp_Int = False Then
1790                .cmbTaxCodes_Interest.Visible = False
1800                .cmbTaxCodes_Interest_display.Visible = False
1810                .cmbTaxCodes_Interest_lbl2.Visible = True
1820              End If  ' ** blnHideIncExp_Int.
1830              If blnHideIncExp_Purch = False Then
1840                .cmbTaxCodes_Purchase.Visible = False
1850                .cmbTaxCodes_Purchase_display.Visible = False
1860                .cmbTaxCodes_Purchase_lbl2.Visible = True
1870              End If  ' ** blnHideIncExp_Purch.
1880            End Select

1890            blnIsOpen = False

1900          Else
1910            Cancel = -1
1920          End If
1930        Else
1940          Cancel = -1
1950        End If
1960      Else
1970        Cancel = -1
1980      End If
1990    End With

2000    DoCmd.Hourglass False

2010    If Cancel = -1 Then
2020      Beep
2030      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
2040    End If

EXITP:
2050    Exit Sub

ERRH:
2060    DoCmd.Hourglass False
2070    Select Case ERR.Number
        Case Else
2080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2090    End Select
2100    Resume EXITP

End Sub

Private Sub Form_Load()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

2210    With Me

2220      DoCmd.Hourglass True
2230      DoEvents

          ' ** Create an instance of the Calendar class.
2240      Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
2250      clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
2260      clsMonthClass.PositionAtCursor = True

2270      Select Case IsNull(.cmbRevenueCodes_Interest)
          Case True
2280        .cmbRevenueCodes_Interest = REVID_INC
2290        .cmbRevenueCodes_Interest_display = Null
2300      Case False
2310        If .cmbRevenueCodes_Interest = REVID_INC Or .cmbRevenueCodes_Interest = REVID_EXP Then
2320          .cmbRevenueCodes_Interest_display = Null
2330        Else
2340          .cmbRevenueCodes_Interest_display = .cmbRevenueCodes_Interest.Column(1)
2350        End If
2360      End Select

2370      DoCmd.Hourglass True
2380      DoEvents

2390      Select Case IsNull(.cmbTaxCodes_Interest)
          Case True
2400        .cmbTaxCodes_Interest = TAXID_INC
2410        .cmbTaxCodes_Interest_display = Null
2420      Case False
2430        If .cmbTaxCodes_Interest = TAXID_INC Or .cmbTaxCodes_Interest = TAXID_DED Then
2440          .cmbTaxCodes_Interest_display = Null
2450        Else
2460          .cmbTaxCodes_Interest_display = .cmbTaxCodes_Interest.Column(1)
2470        End If
2480      End Select

2490      Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
2500        .cmbRevenueCodes_Purchase = REVID_INC
2510        .cmbRevenueCodes_Purchase_display = Null
2520      Case False
2530        If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
2540          .cmbRevenueCodes_Purchase_display = Null
2550        Else
2560          .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
2570        End If
2580      End Select

2590      Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
2600        .cmbTaxCodes_Purchase = TAXID_INC
2610        .cmbTaxCodes_Purchase_display = Null
2620      Case False
2630        If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
2640          .cmbTaxCodes_Purchase_display = Null
2650        Else
2660          .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
2670        End If
2680      End Select

2690      DoCmd.Hourglass False

2700    End With

EXITP:
2710    Exit Sub

ERRH:
2720    DoCmd.Hourglass False
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2820    intShiftDown = (Shift And acShiftMask) > 0
2830    intAltDown = (Shift And acAltMask) > 0
2840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Daily Factor:     A {tmpRate on frmAccruedIncome_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmAccruedIncome_Sub}

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyEscape
2870        intRetVal = 0
2880        cmdClose_Click  ' ** Procedure: Below.
2890      Case vbKeyF7
2900        With Me
2910          intRetVal = 0
2920          .cmdCalendar.SetFocus
2930          cmdCalendar_Click  ' ** Procedure: Below.
2940        End With
2950      End Select
2960    End If

        ' ** Alt keys.
2970    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2980      Select Case intRetVal
          Case vbKeyL
2990        With Me
3000          intRetVal = 0
3010          If .frmAccruedIncome_Sub.Form.tmpDate_lbl.Visible = True Then
3020            .frmAccruedIncome_Sub.SetFocus
3030            .frmAccruedIncome_Sub.Form.tmpRate.SetFocus
3040          Else
3050            Beep
3060          End If
3070        End With
3080      Case vbKeyX
3090        intRetVal = 0
3100        cmdClose_Click  ' ** Procedure: Below.
3110      End Select
3120    End If

        ' ** Ctrl keys.
3130    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3140      Select Case intRetVal
          Case vbKeyS
3150        intRetVal = 0
3160        With Me
3170          If .frmAccruedIncome_Sub.Form.tmpDate_lbl.Visible = True Then
3180            .frmAccruedIncome_Sub.SetFocus
3190            .frmAccruedIncome_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAccruedIncome_Sub.
3200          End If
3210        End With
3220      End Select
3230    End If

        ' ** Ctrl-Shift keys.
3240    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3250      Select Case intRetVal
          Case vbKeyF
3260        intRetVal = 0
3270        With Me
3280          .FocusHolder.SetFocus
3290        End With
3300      End Select
3310    End If

EXITP:
3320    KeyCode = intRetVal
3330    Exit Sub

ERRH:
3340    intRetVal = 0
3350    Select Case ERR.Number
        Case Else
3360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3370    End Select
3380    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3410    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3420    If Not clsMonthClass Is Nothing Then
3430      If clsMonthClass.IsCalendar Then
3440        Cancel = -1
3450      Else
3460        Set clsMonthClass = Nothing
3470      End If
3480    End If

3490    gstrFormQuerySpec = vbNullString
3500    gblnSetFocus = True
3510    gstrReturningForm = THIS_NAME
3520    Select Case strCallingForm
        Case "frmMenu_Post"
3530      If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
3540        DoCmd.SelectObject acForm, strCallingForm, False
3550        Forms(strCallingForm).UpdateTransCnt  ' ** Form Function: frmMenu_Post.
3560        Forms(strCallingForm).TimerInterval = 250&
3570      Else
3580        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
3590      End If
3600    End Select

EXITP:
3610    Exit Sub

ERRH:
3620    Select Case ERR.Number
        Case Else
3630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3640    End Select
3650    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3710    With Me
3720      If .cmdCalendar_raised_focus_dots_img.Visible = True Or .cmdCalendar_raised_focus_img.Visible = True Then
3730        Select Case blnCalendar1_Focus
            Case True
3740          .cmdCalendar_raised_semifocus_dots_img.Visible = True
3750          .cmdCalendar_raised_img.Visible = False
3760        Case False
3770          .cmdCalendar_raised_img.Visible = True
3780          .cmdCalendar_raised_semifocus_dots_img.Visible = False
3790        End Select
3800        .cmdCalendar_raised_focus_dots_img.Visible = False
3810        .cmdCalendar_raised_focus_img.Visible = False
3820        .cmdCalendar_sunken_focus_dots_img.Visible = False
3830        .cmdCalendar_raised_img_dis.Visible = False
3840      End If
3850    End With

EXITP:
3860    Exit Sub

ERRH:
3870    Select Case ERR.Number
        Case Else
3880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3890    End Select
3900    Resume EXITP

End Sub

Public Sub cmdClose_Click()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

4010    DoCmd.Close acForm, THIS_NAME

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4120    intShiftDown = (Shift And acShiftMask) > 0
4130    intAltDown = (Shift And acAltMask) > 0
4140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4160      Select Case intRetVal
          Case vbKeyTab
4170        With Me
4180          intRetVal = 0
4190          If .AccrualMonth.Enabled = True Then
4200            .AccrualMonth.SetFocus
4210          Else
4220            If .frmAccruedIncome_Sub.Enabled = True Then
4230              lngRecsCur = .frmAccruedIncome_Sub.Form.RecCnt  ' ** Form Function: frmAccruedIncome_Sub.
4240              If lngRecsCur > 0& Then
4250                .frmAccruedIncome_Sub.SetFocus
4260                .frmAccruedIncome_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAccruedIncome_Sub.
4270              Else
4280                If .cmdCopyFactor.Enabled = True Then
4290                  .cmdCopyFactor.SetFocus
4300                ElseIf .chkReinvest.Enabled = True Then
4310                  .chkReinvest.SetFocus
4320                ElseIf .cmdCreateEntries.Enabled = True Then
4330                  .cmdCreateEntries.SetFocus
4340                Else
                      ' ** Nowhere else to go!
4350                End If
4360              End If
4370            ElseIf .cmdCopyFactor.Enabled = True Then
4380              .cmdCopyFactor.SetFocus
4390            ElseIf .cmdCreateEntries.Enabled = True Then
4400              .cmdCreateEntries.SetFocus
4410            Else
                  ' ** Nowhere else to go!
4420            End If
4430          End If
4440        End With
4450      End Select
4460    End If

        ' ** Shift keys.
4470    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4480      Select Case intRetVal
          Case vbKeyTab
4490        With Me
4500          intRetVal = 0
4510          If .cmdCreateEntries.Enabled = True Then
4520            .cmdCreateEntries.SetFocus
4530          ElseIf .ReinvestComments.Enabled = True Then
4540            .ReinvestComments.SetFocus
4550          ElseIf .chkReinvest.Enabled = True Then
4560            .chkReinvest.SetFocus
4570          ElseIf .AccrualDate.Enabled = True Then
4580            .AccrualDate.SetFocus
4590          ElseIf .cmdCopyFactor.Enabled = True Then
4600            .cmdCopyFactor.SetFocus
4610          Else
4620            If .frmAccruedIncome_Sub.Enabled = True Then
4630              lngRecsCur = .frmAccruedIncome_Sub.Form.RecCnt  ' ** Form Function: frmAccruedIncome_Sub.
4640              If lngRecsCur > 0& Then
4650                .frmAccruedIncome_Sub.SetFocus
4660                .frmAccruedIncome_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccruedIncome_Sub.
4670              Else
4680                If .cmdGenerateDates.Enabled = True Then
4690                  .cmdGenerateDates.SetFocus
4700                ElseIf .AccrualYear.Enabled = True Then
4710                  .AccrualYear.SetFocus
4720                ElseIf .AccrualMonth.Enabled = True Then
4730                  .AccrualMonth.SetFocus
4740                Else
                      ' ** Nowhere else to go!
4750                End If
4760              End If
4770            Else
4780              If .cmdGenerateDates.Enabled = True Then
4790                .cmdGenerateDates.SetFocus
4800              ElseIf .AccrualYear.Enabled = True Then
4810                .AccrualYear.SetFocus
4820              ElseIf .AccrualMonth.Enabled = True Then
4830                .AccrualMonth.SetFocus
4840              Else
                    ' ** Nowhere else to go!
4850              End If
4860            End If
4870          End If
4880        End With
4890      End Select
4900    End If

EXITP:
4910    KeyCode = intRetVal
4920    Exit Sub

ERRH:
4930    intRetVal = 0
4940    Select Case ERR.Number
        Case Else
4950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4960    End Select
4970    Resume EXITP

End Sub

Private Sub AccrualMonth_AfterUpdate()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "AccrualMonth_AfterUpdate"

5010    With Me
5020      If .AccrualMonth = "December" Then
5030        .AccrualYear = year(Date) - 1
5040      Else
5050        .AccrualYear = year(Date)
5060      End If
5070    End With

EXITP:
5080    Exit Sub

ERRH:
5090    Select Case ERR.Number
        Case Else
5100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5110    End Select
5120    Resume EXITP

End Sub

Private Sub AccrualMonth_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "AccrualMonth_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5220    intShiftDown = (Shift And acShiftMask) > 0
5230    intAltDown = (Shift And acAltMask) > 0
5240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5270        With Me
5280          intRetVal = 0
5290          .AccrualYear.SetFocus
5300        End With
5310      End Select
5320    End If

        ' ** Shift keys.
5330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5350        With Me
5360          intRetVal = 0
5370          .cmdClose.SetFocus
5380        End With
5390      End Select
5400    End If

EXITP:
5410    KeyCode = intRetVal
5420    Exit Sub

ERRH:
5430    intRetVal = 0
5440    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
5450    Case Else
5460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5470    End Select
5480    Resume EXITP

End Sub

Private Sub cmdGenerateDates_Click()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdGenerateDates_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim intMonth As Integer, intYear As Integer
        Dim datEndDate As Date
        Dim intDays As Integer
        Dim blnContinue As Boolean

5510    blnContinue = True

5520    With Me

5530      .AccrualMonth.SetFocus
5540      If Len(Trim(.AccrualMonth.text)) = 0 Then
5550        blnContinue = False
5560        MsgBox "You must select a month.", vbInformation + vbOKOnly, "Entry Required"
5570      End If

5580      If blnContinue Then
5590        .AccrualYear.SetFocus
5600        If Len(Trim(.AccrualYear.text)) = 0 Then
5610          blnContinue = False
5620          MsgBox "You must enter a year.", vbInformation + vbOKOnly, "Entry Required"
5630        End If
5640      End If

5650      If blnContinue Then
5660        DoCmd.Hourglass True
5670        DoEvents
5680        intMonth = 0
5690        intYear = 0
5700        .AccrualYear.SetFocus
5710        intYear = CInt(.AccrualYear.text)
5720        .AccrualMonth.SetFocus
5730        intMonth = CInt(.AccrualMonth.Column(0))
5740        If intMonth = 0 Then
5750          blnContinue = False
5760          DoCmd.Hourglass False
5770          MsgBox "You must select a valid month.", vbInformation + vbOKOnly, "Invalid Entry"
5780        Else
5790          If (intYear > CInt(year(Now()))) Or (intYear < (CInt(year(Now())) - 1)) Then
5800            blnContinue = False
5810            DoCmd.Hourglass False
5820            MsgBox "You must enter either this year or last year.", vbInformation + vbOKOnly, "Invalid Entry"
5830          End If
5840        End If
5850      End If

          ' ** Calculate beginning and ending dates and populate table.
5860      If blnContinue Then

5870        datStartDate = CVDate(str(intMonth) + "/01/" + str(intYear))
5880        If intMonth = 12 Then
5890          datEndDate = CVDate("01/01/" + str(intYear + 1))
5900        Else
5910          datEndDate = CVDate(str(intMonth + 1) + "/01/" + str(intYear))
5920        End If
5930        datEndDate = datEndDate - 1
5940        datEndDate = datEndDate

5950        If datStartDate > Now() Then
5960          blnContinue = False
5970          DoCmd.Hourglass False
5980          MsgBox "You must choose a date to begin which is before today.", vbInformation + vbOKOnly, "Invalid Entry"
5990        Else

6000          DoCmd.Hourglass True
6010          DoEvents

6020          Set dbs = CurrentDb
6030          With dbs
                ' ** Empty tmpEdit07.
6040            Set qdf = .QueryDefs("qryAccruedIncome_05")
6050            qdf.Execute
6060            Set rst = .OpenRecordset("tmpEdit07", dbOpenDynaset, dbConsistent)
6070            With rst
6080              Do While datStartDate <= datEndDate
6090                .AddNew
6100                ![tmpDate] = datStartDate
6110                ![tmpRate] = 0#
6120                .Update
6130                datStartDate = datStartDate + 1
6140                If datStartDate > Now() Then
6150                  datStartDate = datEndDate + 1  ' ** All done.
6160                End If
6170              Loop
6180              .MoveFirst
6190              .MoveLast
6200              intDays = .RecordCount
6210              .Close
6220            End With
6230            .Close
6240          End With

6250          .frmAccruedIncome_Sub.Requery
6260          .frmAccruedIncome_Sub.Enabled = True
6270          .frmAccruedIncome_Sub.SetFocus
6280          .frmAccruedIncome_Sub.Form.tmpDays_lbl1.Caption = CStr(intDays)
6290          .frmAccruedIncome_Sub.Form.tmpDays_lbl1.ForeColor = CLR_DKGRY
6300          .frmAccruedIncome_Sub.Form.tmpDays_lbl2.ForeColor = CLR_DKGRY
6310          .cmdCopyFactor.Enabled = True
6320          Select Case IsNull(.AccrualDate)
              Case True
6330            .cmdCreateEntries.Enabled = False
6340          Case False
6350            .cmdCreateEntries.Enabled = True
6360          End Select
6370          .AccrualDate.Enabled = True
6380          .AccrualDate.BorderColor = CLR_LTBLU2
6390          .AccrualDate.BackStyle = acBackStyleNormal
6400          .cmdCalendar.Enabled = True
6410          .cmdCalendar_raised_img.Visible = True
6420          .cmdCalendar_raised_semifocus_dots_img.Visible = False
6430          .cmdCalendar_raised_focus_img.Visible = False
6440          .cmdCalendar_raised_focus_dots_img.Visible = False
6450          .cmdCalendar_sunken_focus_dots_img.Visible = False
6460          .cmdCalendar_raised_img_dis.Visible = False
6470          If blnHideIncExp_Int = False Then
6480            If gblnRevenueExpenseTracking = True Or gblnIncomeTaxCoding = True Then
6490              .cmbRevenueCodes_Interest_box_lbl.ForeColor = CLR_VDKGRY
6500              .cmbRevenueCodes_Interest_box_lbl_dim_hi.Visible = False
6510            End If
6520            .cmbRevenueCodes_Interest_display.Enabled = False
6530            .cmbRevenueCodes_Interest_display.Locked = True
6540            If lngRevID_Interest > 0& Then
6550              .cmbRevenueCodes_Interest = lngRevID_Interest
6560              .cmbRevenueCodes_Interest_display = strRevCode_Interest
6570            Else
6580              .cmbRevenueCodes_Interest = REVID_INC  ' ** Doesn't matter if it's visible or not.
6590              .cmbRevenueCodes_Interest_display = Null
6600            End If
6610            .cmbRevenueCodes_Interest.Requery
6620            .cmbRevenueCodes_Interest.Enabled = True
6630            .cmbRevenueCodes_Interest.BorderColor = CLR_LTBLU2
6640            .cmbRevenueCodes_Interest_lbl.BackStyle = acBackStyleNormal
6650            .cmbTaxCodes_Interest_display.Enabled = False
6660            .cmbTaxCodes_Interest_display.Locked = True
6670            If lngTaxID_Interest > 0& Then
6680              .cmbTaxCodes_Interest = lngTaxID_Interest
6690              .cmbTaxCodes_Interest_display = strTaxCode_Interest
6700            Else
6710              .cmbTaxCodes_Interest = TAXID_INC  ' ** Doesn't matter if it's visible or not.
6720              .cmbTaxCodes_Interest_display = Null
6730            End If
6740            .cmbTaxCodes_Interest.Requery
6750            .cmbTaxCodes_Interest.Enabled = True
6760            .cmbTaxCodes_Interest.BorderColor = CLR_LTBLU2
6770            .cmbTaxCodes_Interest_lbl.BackStyle = acBackStyleNormal
6780          End If  ' ** blnHideIncExp_Int.
6790          .chkReinvest.Enabled = True
6800          If .chkReinvest = True Then
6810            If blnHideIncExp_Purch = False Then
6820              If gblnRevenueExpenseTracking = True Or gblnIncomeTaxCoding = True Then
6830                .cmbRevenueCodes_Purchase_box_lbl.ForeColor = CLR_VDKGRY
6840                .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Visible = False
6850              End If
6860              .cmbRevenueCodes_Purchase_display.Enabled = False
6870              .cmbRevenueCodes_Purchase_display.Locked = True
6880              If lngRevID_Purchase > 0& Then
6890                .cmbRevenueCodes_Purchase = lngRevID_Purchase
6900                .cmbRevenueCodes_Purchase_display = strRevCode_Purchase
6910              Else
6920                .cmbRevenueCodes_Purchase = REVID_INC  ' ** Doesn't matter if it's visible or not.
6930                .cmbRevenueCodes_Purchase_display = Null
6940              End If
6950              .cmbRevenueCodes_Purchase.Requery
6960              .cmbRevenueCodes_Purchase.Enabled = True
6970              .cmbRevenueCodes_Purchase.BorderColor = CLR_LTBLU2
6980              .cmbRevenueCodes_Purchase_lbl.BackStyle = acBackStyleNormal
6990              .cmbTaxCodes_Purchase_display.Enabled = False
7000              .cmbTaxCodes_Purchase_display.Locked = True
7010              If lngTaxID_Purchase > 0& Then
7020                .cmbTaxCodes_Purchase = lngTaxID_Purchase
7030                .cmbTaxCodes_Purchase_display = strTaxCode_Purchase
7040              Else
7050                .cmbTaxCodes_Purchase = TAXID_INC  ' ** Doesn't matter if it's visible or not.
7060                .cmbTaxCodes_Purchase_display = Null
7070              End If
7080              .cmbTaxCodes_Purchase.Requery
7090              .cmbTaxCodes_Purchase.Enabled = True
7100              .cmbTaxCodes_Purchase.BorderColor = CLR_LTBLU2
7110              .cmbTaxCodes_Purchase_lbl.BackStyle = acBackStyleNormal
7120            End If  ' ** blnHideIncExp_Purch.
7130          End If
7140          .AccrualMonth.Enabled = False
7150          .AccrualMonth.Locked = True
7160          .AccrualMonth.ForeColor = CLR_DISABLED_FG
7170          .AccrualMonth.BackColor = CLR_DISABLED_BG
7180          .AccrualMonth.BorderColor = CLR_LTBLU2
7190          .AccrualYear.Enabled = False
7200          .AccrualYear.Locked = True
7210          .AccrualYear.ForeColor = CLR_DISABLED_FG
7220          .AccrualYear.BackColor = CLR_DISABLED_BG
7230          .AccrualYear.BorderColor = CLR_LTBLU2
7240          .cmdGenerateDates.Enabled = False
7250          .frmAccruedIncome_Sub.Form.tmpDate_lbl.Visible = True
7260          .frmAccruedIncome_Sub.Form.tmpDate_lbl_dim.Visible = False
7270          .frmAccruedIncome_Sub.Form.tmpDate_lbl_dim_hi.Visible = False
7280          .frmAccruedIncome_Sub.Form.tmpDate_lbl_line.BorderColor = CLR_DKGRY3
7290          .frmAccruedIncome_Sub.Form.tmpDate_lbl_line_dim_hi.Visible = False
7300          .frmAccruedIncome_Sub.Form.tmpRate_lbl.Visible = True
7310          .frmAccruedIncome_Sub.Form.tmpRate_lbl_dim.Visible = False
7320          .frmAccruedIncome_Sub.Form.tmpRate_lbl_dim_hi.Visible = False
7330          .frmAccruedIncome_Sub.Form.tmpRate_lbl_line.BorderColor = CLR_DKGRY3
7340          .frmAccruedIncome_Sub.Form.tmpRate_lbl_line_dim_hi.Visible = False
7350          .frmAccruedIncome_Sub.Form.tmpRate.SetFocus

7360        End If
7370      End If

7380      DoCmd.Hourglass False

7390    End With

EXITP:
7400    Set rst = Nothing
7410    Set qdf = Nothing
7420    Set dbs = Nothing
7430    Exit Sub

ERRH:
7440    DoCmd.Hourglass False
7450    Select Case ERR.Number
        Case Else
7460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7470    End Select
7480    Resume EXITP

End Sub

Private Sub cmdGenerateDates_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdGenerateDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7520    intShiftDown = (Shift And acShiftMask) > 0
7530    intAltDown = (Shift And acAltMask) > 0
7540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7560      Select Case intRetVal
          Case vbKeyTab
7570        With Me
7580          intRetVal = 0
7590          If .frmAccruedIncome_Sub.Enabled = True Then
7600            lngRecsCur = .frmAccruedIncome_Sub.Form.RecCnt  ' ** Form Function: frmAccruedIncome_Sub.
7610            If lngRecsCur > 0& Then
7620              .frmAccruedIncome_Sub.SetFocus
7630              .frmAccruedIncome_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAccruedIncome_Sub.
7640              .frmAccruedIncome_Sub.Form.tmpRate.SetFocus
7650            Else
7660              .cmdClose.SetFocus
7670            End If
7680          Else
7690            .cmdClose.SetFocus
7700          End If
7710        End With
7720      End Select
7730    End If

        ' ** Shift keys.
7740    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7750      Select Case intRetVal
          Case vbKeyTab
7760        With Me
7770          intRetVal = 0
7780          .AccrualYear.SetFocus
7790        End With
7800      End Select
7810    End If

EXITP:
7820    KeyCode = intRetVal
7830    Exit Sub

ERRH:
7840    intRetVal = 0
7850    Select Case ERR.Number
        Case Else
7860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7870    End Select
7880    Resume EXITP

End Sub

Private Sub cmdCopyFactor_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopyFactor_box_MouseMove"

7910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7920    Exit Sub

ERRH:
7930    Select Case ERR.Number
        Case Else
7940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7950    End Select
7960    Resume EXITP

End Sub

Private Sub cmdCopyFactor_Click()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopyFactor_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim dblRate As Double
        Dim lngRecs As Long
        Dim lngX As Long

8010    With Me
8020      Set dbs = CurrentDb
8030      With dbs
            ' ** tmpEdit07, order by tmpDate.
8040        Set qdf = .QueryDefs("qryAccruedIncome_03")
8050        Set rst = qdf.OpenRecordset
8060        With rst
8070          If .BOF = True And .EOF = True Then
                ' ** No date records!
8080          Else
8090            .MoveLast
8100            lngRecs = .RecordCount
8110            .MoveFirst
8120            dblRate = ![tmpRate]
8130            .MoveNext
8140            For lngX = 2& To lngRecs
8150              .Edit
8160              ![tmpRate] = dblRate
8170              .Update
8180              If lngX < lngRecs Then .MoveNext
8190            Next
8200          End If
8210          .Close
8220        End With
8230        .Close
8240      End With
8250      .frmAccruedIncome_Sub.Form.Requery
8260      .AccrualDate.SetFocus
8270    End With

EXITP:
8280    Set rst = Nothing
8290    Set qdf = Nothing
8300    Set dbs = Nothing

8310    Exit Sub

ERRH:
8320    Select Case ERR.Number
        Case Else
8330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8340    End Select
8350    Resume EXITP

End Sub

Private Sub cmdCopyFactor_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopyFactor_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8420    intShiftDown = (Shift And acShiftMask) > 0
8430    intAltDown = (Shift And acAltMask) > 0
8440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8460      Select Case intRetVal
          Case vbKeyTab
8470        With Me
8480          intRetVal = 0
8490          If .AccrualDate.Enabled = True Then
8500            .AccrualDate.SetFocus
8510          ElseIf .cmbRevenueCodes_Interest.Visible = True And .cmbRevenueCodes_Interest.Enabled = True Then
8520            .cmbRevenueCodes_Interest.SetFocus
8530          ElseIf .cmbTaxCodes_Interest.Visible = True And .cmbTaxCodes_Interest.Enabled = True Then
8540            .cmbTaxCodes_Interest.SetFocus
8550          ElseIf .chkReinvest.Enabled = True Then
8560            .chkReinvest.SetFocus
8570          ElseIf .cmdCreateEntries.Enabled = True Then
8580            .cmdCreateEntries.SetFocus
8590          Else
8600            .cmdClose.SetFocus
8610          End If
8620        End With
8630      End Select
8640    End If

        ' ** Shift keys.
8650    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8660      Select Case intRetVal
          Case vbKeyTab
8670        With Me
8680          intRetVal = 0
8690          If .frmAccruedIncome_Sub.Enabled = True Then
8700            lngRecsCur = .frmAccruedIncome_Sub.Form.RecCnt  ' ** Form Function: frmAccruedIncome_Sub.
8710            If lngRecsCur > 0& Then
8720              .frmAccruedIncome_Sub.SetFocus
8730              .frmAccruedIncome_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccruedIncome_Sub.
8740            Else
8750              If .cmdGenerateDates.Enabled = True Then
8760                .cmdGenerateDates.SetFocus
8770              Else
8780                .cmdClose.SetFocus
8790              End If
8800            End If
8810          ElseIf .cmdGenerateDates.Enabled = True Then
8820            .cmdGenerateDates.SetFocus
8830          Else
8840            .cmdClose.SetFocus
8850          End If
8860        End With
8870      End Select
8880    End If

EXITP:
8890    KeyCode = intRetVal
8900    Exit Sub

ERRH:
8910    intRetVal = 0
8920    Select Case ERR.Number
        Case Else
8930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8940    End Select
8950    Resume EXITP

End Sub

Private Sub AccrualDate_Enter()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "AccrualDate_Enter"

9010    With Me
9020      If blnFromCalendar = False Then
9030        If IsDate(.AccrualDate) = False Then
9040          .AccrualDate = vbNullString
9050        End If
9060      Else
9070        blnFromCalendar = False
9080        .AccrualDate.SelLength = 0
9090        .AccrualDate.SelStart = 99
9100      End If
9110    End With

EXITP:
9120    Exit Sub

ERRH:
9130    DoCmd.Hourglass False
9140    Select Case ERR.Number
        Case Else
9150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9160    End Select
9170    Resume EXITP

End Sub

Private Sub AccrualDate_BeforeUpdate(Cancel As Integer)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "AccrualDate_BeforeUpdate"

9210    With Me
9220      If IsNull(.AccrualDate) = False Then
9230        If IsDate(.AccrualDate) = True Then
9240          If DateCheck_Post(.AccrualDate) = False Then  ' ** Module Function: modUtilities.
9250            DoCmd.Hourglass False
                ' ** MsgBox already displayed.
9260            Cancel = -1
9270          End If
9280        End If
9290      End If
9300    End With

EXITP:
9310    Exit Sub

ERRH:
9320    DoCmd.Hourglass False
9330    Select Case ERR.Number
        Case Else
9340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9350    End Select
9360    Resume EXITP

End Sub

Private Sub AccrualDate_DblClick(Cancel As Integer)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "AccrualDate_DblClick"

9410    With Me
9420      .AccrualDate = Date
9430      .chkReinvest.SetFocus
9440    End With

EXITP:
9450    Exit Sub

ERRH:
9460    Select Case ERR.Number
        Case Else
9470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9480    End Select
9490    Resume EXITP

End Sub

Private Sub AccrualDate_AfterUpdate()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "AccrualDate_AfterUpdate"

9510    With Me
9520      Select Case IsNull(.AccrualDate)
          Case True
9530        .cmdCreateEntries.Enabled = False
9540      Case False
9550        Select Case IsDate(.AccrualDate)
            Case True
9560          .cmdCreateEntries.Enabled = True
9570        Case False
9580          .cmdCreateEntries.Enabled = False
9590        End Select
9600      End Select
9610    End With

EXITP:
9620    Exit Sub

ERRH:
9630    Select Case ERR.Number
        Case Else
9640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9650    End Select
9660    Resume EXITP

End Sub

Private Sub AccrualDate_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "AccrualDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9720    intShiftDown = (Shift And acShiftMask) > 0
9730    intAltDown = (Shift And acAltMask) > 0
9740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9770        With Me
9780          intRetVal = 0
9790          blnKeyDown_Return = True
9800          If .cmbRevenueCodes_Interest.Visible = True And .cmbRevenueCodes_Interest.Enabled = True Then
9810            .cmbRevenueCodes_Interest.SetFocus
9820          ElseIf .cmbTaxCodes_Interest.Visible = True And .cmbTaxCodes_Interest.Enabled = True Then
9830            .cmbTaxCodes_Interest.SetFocus
9840          ElseIf .chkReinvest.Enabled = True Then
9850            .chkReinvest.SetFocus
9860          ElseIf .cmdCreateEntries.Enabled = True Then
9870            .cmdCreateEntries.SetFocus
9880          Else
9890            .cmdClose.SetFocus
9900          End If
9910        End With
9920      End Select
9930    End If

        ' ** Shift keys.
9940    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9950      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9960        With Me
9970          intRetVal = 0
9980          blnKeyDown_Return = True
9990          If .cmdCopyFactor.Enabled = True Then
10000           .cmdCopyFactor.SetFocus
10010         ElseIf .frmAccruedIncome_Sub.Enabled = True Then
10020           lngRecsCur = .frmAccruedIncome_Sub.Form.RecCnt  ' ** Form Function: frmAccruedIncome_Sub.
10030           If lngRecsCur > 0& Then
10040             .frmAccruedIncome_Sub.SetFocus
10050             .frmAccruedIncome_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccruedIncome_Sub.
10060           Else
10070             If .cmdGenerateDates.Enabled = True Then
10080               .cmdGenerateDates.SetFocus
10090             Else
10100               .cmdClose.SetFocus
10110             End If
10120           End If
10130         ElseIf .cmdGenerateDates.Enabled = True Then
10140           .cmdGenerateDates.SetFocus
10150         Else
10160           .cmdClose.SetFocus
10170         End If
10180       End With
10190     End Select
10200   End If

EXITP:
10210   KeyCode = intRetVal
10220   Exit Sub

ERRH:
10230   intRetVal = 0
10240   Select Case ERR.Number
        Case Else
10250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10260   End Select
10270   Resume EXITP

End Sub

Private Sub AccrualDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "AccrualDate_MouseMove"

10310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10320   Exit Sub

ERRH:
10330   Select Case ERR.Number
        Case Else
10340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10350   End Select
10360   Resume EXITP

End Sub

Private Sub AccrualDate_Exit(Cancel As Integer)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "AccrualDate_Exit"

        Dim blnJustMove As Boolean

10410   With Me
10420     blnJustMove = False
10430     If IsNull(.AccrualDate) = False Then
10440       If .AccrualDate <> vbNullString Then
              ' ** Proceed.
10450         blnJustMove = True
10460         .cmdCreateEntries.Enabled = True
10470       End If
10480     End If
10490     If blnJustMove = False Then
10500       If blnKeyDown_Return = True Then
10510         blnKeyDown_Return = False
10520         .AccrualDate = Date
10530         .cmdCreateEntries.Enabled = True
10540       Else
10550         If IsNull(.AccrualDate.text) = False Then
10560           If Trim(.AccrualDate.text) <> vbNullString Then
                  ' ** Proceed.
10570             .cmdCreateEntries.Enabled = True
10580           Else
                  ' ** Empty, do nothing.
10590           End If
10600         Else
                ' ** Empty, do nothing.
10610         End If
10620       End If
10630     End If
10640   End With

EXITP:
10650   Exit Sub

ERRH:
10660   DoCmd.Hourglass False
10670   Select Case ERR.Number
        Case Else
10680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10690   End Select
10700   Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

10810   With Me
10820     blnCalendar1_Focus = True
10830     .cmdCalendar_raised_semifocus_dots_img.Visible = True
10840     .cmdCalendar_raised_img.Visible = False
10850     .cmdCalendar_raised_focus_img.Visible = False
10860     .cmdCalendar_raised_focus_dots_img.Visible = False
10870     .cmdCalendar_sunken_focus_dots_img.Visible = False
10880     .cmdCalendar_raised_img_dis.Visible = False
10890   End With

EXITP:
10900   Exit Sub

ERRH:
10910   Select Case ERR.Number
        Case Else
10920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10930   End Select
10940   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

11010   With Me
11020     blnCalendar1_MouseDown = True
11030     .cmdCalendar_sunken_focus_dots_img.Visible = True
11040     .cmdCalendar_raised_img.Visible = False
11050     .cmdCalendar_raised_semifocus_dots_img.Visible = False
11060     .cmdCalendar_raised_focus_img.Visible = False
11070     .cmdCalendar_raised_focus_dots_img.Visible = False
11080     .cmdCalendar_raised_img_dis.Visible = False
11090   End With

EXITP:
11100   Exit Sub

ERRH:
11110   Select Case ERR.Number
        Case Else
11120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11130   End Select
11140   Resume EXITP

End Sub

Public Sub cmdCalendar_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

11210   With Me
11220     datStartDate = Date
11230     datEndDate = 0
11240     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
11250     If blnRetVal = True Then
            ' ** Allow posting up to 1 month into the future.
11260       If datStartDate > DateAdd("m", 1, Date) Then
11270         MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
11280         .AccrualDate = CDate(Format(Date, "mm/dd/yyyy"))
11290       Else
11300         blnFromCalendar = True
11310         .AccrualDate = datStartDate
11320       End If
11330       .cmdCreateEntries.Enabled = True
11340     Else
11350       blnFromCalendar = True
11360       .AccrualDate = CDate(Format(Date, "mm/dd/yyyy"))
11370       .cmdCreateEntries.Enabled = True
11380     End If
11390     .AccrualDate.SetFocus
11400   End With

EXITP:
11410   Exit Sub

ERRH:
11420   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
11430   Case Else
11440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11450   End Select
11460   Resume EXITP

End Sub

Private Sub cmdCalendar_KeyDown(KeyCode As Integer, Shift As Integer)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11520   intShiftDown = (Shift And acShiftMask) > 0
11530   intAltDown = (Shift And acAltMask) > 0
11540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11560     Select Case intRetVal
          Case vbKeyTab
11570       With Me
11580         intRetVal = 0
11590         If .cmbRevenueCodes_Interest.Visible = True And .cmbRevenueCodes_Interest.Enabled = True Then
11600           .cmbRevenueCodes_Interest.SetFocus
11610         ElseIf .cmbTaxCodes_Interest.Visible = True And .cmbTaxCodes_Interest.Enabled = True Then
11620           .cmbTaxCodes_Interest.SetFocus
11630         ElseIf .chkReinvest.Enabled = True Then
11640           .chkReinvest.SetFocus
11650         ElseIf .cmdCreateEntries.Enabled = True Then
11660           .cmdCreateEntries.SetFocus
11670         Else
11680           .cmdClose.SetFocus
11690         End If
11700       End With
11710     End Select
11720   End If

        ' ** Shift keys.
11730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11740     Select Case intRetVal
          Case vbKeyTab
11750       With Me
11760         intRetVal = 0
11770         If .AccrualDate.Enabled = True Then
11780           .AccrualDate.SetFocus
11790         ElseIf .cmdCopyFactor.Enabled = True Then
11800           .cmdCopyFactor.SetFocus
11810         ElseIf .frmAccruedIncome_Sub.Enabled = True Then
11820           lngRecsCur = .frmAccruedIncome_Sub.Form.RecCnt  ' ** Form Function: frmAccruedIncome_Sub.
11830           If lngRecsCur > 0& Then
11840             .frmAccruedIncome_Sub.SetFocus
11850             .frmAccruedIncome_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccruedIncome_Sub.
11860           Else
11870             If .cmdGenerateDates.Enabled = True Then
11880               .cmdGenerateDates.SetFocus
11890             Else
11900               .cmdClose.SetFocus
11910             End If
11920           End If
11930         ElseIf .cmdGenerateDates.Enabled = True Then
11940           .cmdGenerateDates.SetFocus
11950         Else
11960           .cmdClose.SetFocus
11970         End If
11980       End With
11990     End Select
12000   End If

EXITP:
12010   KeyCode = intRetVal
12020   Exit Sub

ERRH:
12030   intRetVal = 0
12040   Select Case ERR.Number
        Case Else
12050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12060   End Select
12070   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

12110   With Me
12120     If blnCalendar1_MouseDown = False Then
12130       Select Case blnCalendar1_Focus
            Case True
12140         .cmdCalendar_raised_focus_dots_img.Visible = True
12150         .cmdCalendar_raised_focus_img.Visible = False
12160       Case False
12170         .cmdCalendar_raised_focus_img.Visible = True
12180         .cmdCalendar_raised_focus_dots_img.Visible = False
12190       End Select
12200       .cmdCalendar_raised_img.Visible = False
12210       .cmdCalendar_raised_semifocus_dots_img.Visible = False
12220       .cmdCalendar_sunken_focus_dots_img.Visible = False
12230       .cmdCalendar_raised_img_dis.Visible = False
12240     End If
12250   End With

EXITP:
12260   Exit Sub

ERRH:
12270   Select Case ERR.Number
        Case Else
12280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12290   End Select
12300   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

12410   With Me
12420     .cmdCalendar_raised_focus_dots_img.Visible = True
12430     .cmdCalendar_raised_img.Visible = False
12440     .cmdCalendar_raised_semifocus_dots_img.Visible = False
12450     .cmdCalendar_raised_focus_img.Visible = False
12460     .cmdCalendar_sunken_focus_dots_img.Visible = False
12470     .cmdCalendar_raised_img_dis.Visible = False
12480     blnCalendar1_MouseDown = False
12490   End With

EXITP:
12500   Exit Sub

ERRH:
12510   Select Case ERR.Number
        Case Else
12520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12530   End Select
12540   Resume EXITP

End Sub

Private Sub cmdCalendar_LostFocus()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

12610   With Me
12620     .cmdCalendar_raised_img.Visible = True
12630     .cmdCalendar_raised_semifocus_dots_img.Visible = False
12640     .cmdCalendar_raised_focus_img.Visible = False
12650     .cmdCalendar_raised_focus_dots_img.Visible = False
12660     .cmdCalendar_sunken_focus_dots_img.Visible = False
12670     .cmdCalendar_raised_img_dis.Visible = False
12680     blnCalendar1_Focus = False
12690   End With

EXITP:
12700   Exit Sub

ERRH:
12710   Select Case ERR.Number
        Case Else
12720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12730   End Select
12740   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Interest_display_GotFocus()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Interest_display_GotFocus"

12810   With Me
12820     .cmbRevenueCodes_Interest.SetFocus
12830   End With

EXITP:
12840   Exit Sub

ERRH:
12850   Select Case ERR.Number
        Case Else
12860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12870   End Select
12880   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Interest_display_KeyDown(KeyCode As Integer, Shift As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Interest_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12920   intShiftDown = (Shift And acShiftMask) > 0
12930   intAltDown = (Shift And acAltMask) > 0
12940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12970       With Me
12980         intRetVal = 0
12990         .cmbRevenueCodes_Interest.SetFocus
13000       End With
13010     End Select
13020   End If

        ' ** Shift keys.
13030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13050       With Me
13060         intRetVal = 0
13070         .AccrualDate.SetFocus
13080       End With
13090     End Select
13100   End If

EXITP:
13110   KeyCode = intRetVal
13120   Exit Sub

ERRH:
13130   intRetVal = 0
13140   Select Case ERR.Number
        Case Else
13150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13160   End Select
13170   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Interest_Enter()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Interest_Enter"

13210   With Me
13220     Select Case IsNull(.cmbRevenueCodes_Interest)
          Case True
13230       ExpandCombo .cmbRevenueCodes_Interest  ' ** Module Procedure: modUtilities.
13240     Case False
13250       If .cmbRevenueCodes_Interest = REVID_INC Or .cmbRevenueCodes_Interest = REVID_EXP Then
13260         ExpandCombo .cmbRevenueCodes_Interest  ' ** Module Procedure: modUtilities.
13270       End If
13280     End Select
13290   End With

EXITP:
13300   Exit Sub

ERRH:
13310   Select Case ERR.Number
        Case Else
13320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13330   End Select
13340   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Interest_AfterUpdate()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Interest_AfterUpdate"

13410   With Me
13420     Select Case IsNull(.cmbRevenueCodes_Interest)
          Case True
13430       .cmbRevenueCodes_Interest = REVID_INC
13440     Case False
            ' ** Nothing at the moment.
13450     End Select
13460     If .cmbRevenueCodes_Interest = REVID_INC Or .cmbRevenueCodes_Interest = REVID_EXP Then
13470       .cmbRevenueCodes_Interest_display = Null
13480     Else
13490       .cmbRevenueCodes_Interest_display = .cmbRevenueCodes_Interest.Column(1)
13500     End If
13510     lngRevID_Interest = .cmbRevenueCodes_Interest
13520   End With

EXITP:
13530   Exit Sub

ERRH:
13540   Select Case ERR.Number
        Case Else
13550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13560   End Select
13570   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Interest_KeyDown(KeyCode As Integer, Shift As Integer)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Interest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13620   intShiftDown = (Shift And acShiftMask) > 0
13630   intAltDown = (Shift And acAltMask) > 0
13640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13670       With Me
13680         intRetVal = 0
13690         If .cmbTaxCodes_Interest.Visible = True And .cmbTaxCodes_Interest.Enabled = True Then
13700           .cmbTaxCodes_Interest.SetFocus
13710         Else
13720           .chkReinvest.SetFocus
13730         End If
13740       End With
13750     End Select
13760   End If

        ' ** Shift keys.
13770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13790       With Me
13800         intRetVal = 0
13810         If .AccrualDate.Enabled = True Then
13820           .AccrualDate.SetFocus
13830         ElseIf .cmdCopyFactor.Enabled = True Then
13840           .cmdCopyFactor.SetFocus
13850         ElseIf .frmAccruedIncome_Sub.Enabled = True Then
13860           lngRecsCur = .frmAccruedIncome_Sub.Form.RecCnt  ' ** Form Function: frmAccruedIncome_Sub.
13870           If lngRecsCur > 0& Then
13880             .frmAccruedIncome_Sub.SetFocus
13890             .frmAccruedIncome_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccruedIncome_Sub.
13900           Else
13910             If .cmdGenerateDates.Enabled = True Then
13920               .cmdGenerateDates.SetFocus
13930             Else
13940               .cmdClose.SetFocus
13950             End If
13960           End If
13970         ElseIf .cmdGenerateDates.Enabled = True Then
13980           .cmdGenerateDates.SetFocus
13990         Else
14000           .cmdClose.SetFocus
14010         End If
14020       End With
14030     End Select
14040   End If

EXITP:
14050   KeyCode = intRetVal
14060   Exit Sub

ERRH:
14070   intRetVal = 0
14080   Select Case ERR.Number
        Case Else
14090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14100   End Select
14110   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Interest_display_GotFocus()

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Interest_display_GotFocus"

14210   With Me
14220     .cmbTaxCodes_Interest.SetFocus
14230   End With

EXITP:
14240   Exit Sub

ERRH:
14250   Select Case ERR.Number
        Case Else
14260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14270   End Select
14280   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Interest_display_KeyDown(KeyCode As Integer, Shift As Integer)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Interest_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14320   intShiftDown = (Shift And acShiftMask) > 0
14330   intAltDown = (Shift And acAltMask) > 0
14340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14370       With Me
14380         intRetVal = 0
14390         .cmbTaxCodes_Interest.SetFocus
14400       End With
14410     End Select
14420   End If

        ' ** Shift keys.
14430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14450       With Me
14460         intRetVal = 0
14470         If .cmbRevenueCodes_Interest.Visible = True And .cmbRevenueCodes_Interest.Enabled = True Then
14480           .cmbRevenueCodes_Interest.SetFocus
14490         ElseIf .AccrualDate.Enabled = True Then
14500           .AccrualDate.SetFocus
14510         ElseIf .cmdCopyFactor.Enabled = True Then
14520           .cmdCopyFactor.SetFocus
14530         ElseIf .frmAccruedIncome_Sub.Enabled = True Then
14540           lngRecsCur = .frmAccruedIncome_Sub.Form.RecCnt  ' ** Form Function: frmAccruedIncome_Sub.
14550           If lngRecsCur > 0& Then
14560             .frmAccruedIncome_Sub.SetFocus
14570             .frmAccruedIncome_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccruedIncome_Sub.
14580           Else
14590             If .cmdGenerateDates.Enabled = True Then
14600               .cmdGenerateDates.SetFocus
14610             Else
14620               .cmdClose.SetFocus
14630             End If
14640           End If
14650         ElseIf .cmdGenerateDates.Enabled = True Then
14660           .cmdGenerateDates.SetFocus
14670         Else
14680           .cmdClose.SetFocus
14690         End If
14700       End With
14710     End Select
14720   End If

EXITP:
14730   KeyCode = intRetVal
14740   Exit Sub

ERRH:
14750   intRetVal = 0
14760   Select Case ERR.Number
        Case Else
14770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14780   End Select
14790   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Interest_Enter()

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Interest_Enter"

14810   With Me
14820     Select Case IsNull(.cmbTaxCodes_Interest)
          Case True
14830       ExpandCombo .cmbTaxCodes_Interest  ' ** Module Procedure: modUtilities.
14840     Case False
14850       If .cmbTaxCodes_Interest = TAXID_INC Or .cmbTaxCodes_Interest = TAXID_DED Then
14860         ExpandCombo .cmbTaxCodes_Interest  ' ** Module Procedure: modUtilities.
14870       End If
14880     End Select
14890   End With

EXITP:
14900   Exit Sub

ERRH:
14910   Select Case ERR.Number
        Case Else
14920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14930   End Select
14940   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Interest_AfterUpdate()

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Interest_AfterUpdate"

15010   With Me
15020     Select Case IsNull(.cmbTaxCodes_Interest)
          Case True
15030       .cmbTaxCodes_Interest = TAXID_INC
15040     Case False
            ' ** Nothing at the moment.
15050     End Select
15060     If .cmbTaxCodes_Interest = TAXID_INC Or .cmbTaxCodes_Interest = TAXID_DED Then
15070       .cmbTaxCodes_Interest_display = Null
15080     Else
15090       .cmbTaxCodes_Interest_display = .cmbTaxCodes_Interest.Column(1)
15100     End If
15110     lngTaxID_Interest = .cmbTaxCodes_Interest
15120   End With

EXITP:
15130   Exit Sub

ERRH:
15140   Select Case ERR.Number
        Case Else
15150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15160   End Select
15170   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Interest_KeyDown(KeyCode As Integer, Shift As Integer)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Interest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15220   intShiftDown = (Shift And acShiftMask) > 0
15230   intAltDown = (Shift And acAltMask) > 0
15240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15270       With Me
15280         intRetVal = 0
15290         If .chkReinvest.Enabled = True Then
15300           .chkReinvest.SetFocus
15310         ElseIf .cmdCreateEntries.Enabled = True Then
15320           .cmdCreateEntries.SetFocus
15330         Else
15340           .cmdClose.SetFocus
15350         End If
15360       End With
15370     End Select
15380   End If

        ' ** Shift keys.
15390   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15400     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15410       With Me
15420         intRetVal = 0
15430         If .cmbRevenueCodes_Interest.Visible = True And .cmbRevenueCodes_Interest.Enabled = True Then
15440           .cmbRevenueCodes_Interest.SetFocus
15450         ElseIf .AccrualDate.Enabled = True Then
15460           .AccrualDate.SetFocus
15470         ElseIf .cmdCopyFactor.Enabled = True Then
15480           .cmdCopyFactor.SetFocus
15490         ElseIf .frmAccruedIncome_Sub.Enabled = True Then
15500           lngRecsCur = .frmAccruedIncome_Sub.Form.RecCnt  ' ** Form Function: frmAccruedIncome_Sub.
15510           If lngRecsCur > 0& Then
15520             .frmAccruedIncome_Sub.SetFocus
15530             .frmAccruedIncome_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccruedIncome_Sub.
15540           Else
15550             If .cmdGenerateDates.Enabled = True Then
15560               .cmdGenerateDates.SetFocus
15570             Else
15580               .cmdClose.SetFocus
15590             End If
15600           End If
15610         ElseIf .cmdGenerateDates.Enabled = True Then
15620           .cmdGenerateDates.SetFocus
15630         Else
15640           .cmdClose.SetFocus
15650         End If
15660       End With
15670     End Select
15680   End If

EXITP:
15690   KeyCode = intRetVal
15700   Exit Sub

ERRH:
15710   intRetVal = 0
15720   Select Case ERR.Number
        Case Else
15730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15740   End Select
15750   Resume EXITP

End Sub

Private Sub chkReinvest_AfterUpdate()

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkReinvest_AfterUpdate"

15810   With Me
15820     Select Case .chkReinvest
          Case True
15830       .chkReinvest_lbl.FontBold = True
15840       If IsNull(.ReinvestComments) = True Then
15850         .ReinvestComments = "Accrued Income Reinvestment"
15860       End If
15870       .ReinvestComments.Enabled = True
15880       .ReinvestComments.BorderColor = CLR_LTBLU2
15890       .ReinvestComments.BackStyle = acBackStyleNormal
15900       .ReinvestComments_lbl.FontBold = True
15910       Select Case IsNull(.ReinvestPerShare)
            Case True
15920         .ReinvestPerShare = 1@
15930       Case False
15940         If .ReinvestPerShare = 0@ Then
15950           .ReinvestPerShare = 1@
15960         End If
15970       End Select
15980       If blnHideIncExp_Purch = False Then
15990         If gblnRevenueExpenseTracking = True Then
16000           .cmbRevenueCodes_Purchase_box_lbl.ForeColor = CLR_VDKGRY
16010           .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Visible = False
16020           .cmbRevenueCodes_Purchase_display.Enabled = False
16030           .cmbRevenueCodes_Purchase_display.Locked = True
16040           If lngRevID_Purchase > 0& Then
16050             .cmbRevenueCodes_Purchase = lngRevID_Purchase
16060             .cmbRevenueCodes_Purchase_display = strRevCode_Purchase
16070           Else
16080             .cmbRevenueCodes_Purchase = REVID_INC  ' ** Doesn't matter if it's visible or not.
16090             .cmbRevenueCodes_Purchase_display = Null
16100           End If
16110           .cmbRevenueCodes_Purchase.Requery
16120           .cmbRevenueCodes_Purchase.Enabled = True
16130           .cmbRevenueCodes_Purchase.BorderColor = CLR_LTBLU2
16140           .cmbRevenueCodes_Purchase_lbl.BackStyle = acBackStyleNormal
16150         End If
16160         If gblnIncomeTaxCoding = True Then
16170           .cmbTaxCodes_Purchase_display.Enabled = False
16180           .cmbTaxCodes_Purchase_display.Locked = True
16190           If lngTaxID_Purchase > 0& Then
16200             .cmbTaxCodes_Purchase = lngTaxID_Purchase
16210             .cmbTaxCodes_Purchase_display = strTaxCode_Purchase
16220           Else
16230             .cmbTaxCodes_Purchase = TAXID_INC  ' ** Doesn't matter if it's visible or not.
16240             .cmbTaxCodes_Purchase_display = Null
16250           End If
16260           .cmbTaxCodes_Purchase.Requery
16270           .cmbTaxCodes_Purchase.Enabled = True
16280           .cmbTaxCodes_Purchase.BorderColor = CLR_LTBLU2
16290           .cmbTaxCodes_Purchase_lbl.BackStyle = acBackStyleNormal
16300         End If
16310       End If  ' ** blnHideIncExp_Purch.
16320       .ReinvestComments.SetFocus
16330     Case False
16340       .chkReinvest_lbl.FontBold = False
16350       .ReinvestComments.Enabled = False
16360       .ReinvestComments.BorderColor = WIN_CLR_DISR
16370       .ReinvestComments.BackStyle = acBackStyleTransparent
16380       .ReinvestComments_lbl.FontBold = False
16390       .ReinvestComments = Null
16400       If blnHideIncExp_Purch = False Then
16410         If gblnRevenueExpenseTracking = True Then
16420           .cmbRevenueCodes_Purchase_box_lbl.ForeColor = WIN_CLR_DISF
16430           .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Visible = True
16440           .cmbRevenueCodes_Purchase_display = Null
16450           .cmbRevenueCodes_Purchase_display.Enabled = False
16460           .cmbRevenueCodes_Purchase_display.Locked = False
16470           .cmbRevenueCodes_Purchase = Null
16480           .cmbRevenueCodes_Purchase.Enabled = False
16490           .cmbRevenueCodes_Purchase.BorderColor = WIN_CLR_DISR
16500           .cmbRevenueCodes_Purchase_lbl.BackStyle = acBackStyleTransparent
16510         End If
16520         If gblnIncomeTaxCoding = True Then
16530           .cmbTaxCodes_Purchase_display = Null
16540           .cmbTaxCodes_Purchase_display.Enabled = False
16550           .cmbTaxCodes_Purchase_display.Locked = False
16560           .cmbTaxCodes_Purchase = Null
16570           .cmbTaxCodes_Purchase.Enabled = False
16580           .cmbTaxCodes_Purchase.BorderColor = WIN_CLR_DISR
16590           .cmbTaxCodes_Purchase_lbl.BackStyle = acBackStyleTransparent
16600         End If
16610       End If  ' ** blnHideIncExp_Purch.
16620       If .cmdCreateEntries.Enabled = True Then
16630         .cmdCreateEntries.SetFocus
16640       End If
16650     End Select
16660   End With

EXITP:
16670   Exit Sub

ERRH:
16680   Select Case ERR.Number
        Case Else
16690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16700   End Select
16710   Resume EXITP

End Sub

Private Sub chkReinvest_KeyDown(KeyCode As Integer, Shift As Integer)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkReinvest_KeyDown"

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
16890         If .ReinvestPerShare.Visible = True And .ReinvestPerShare.Enabled = True Then
16900           .ReinvestPerShare.SetFocus
16910         ElseIf .ReinvestComments.Enabled = True Then
16920           .ReinvestComments.SetFocus
16930         ElseIf .cmdCreateEntries.Enabled = True Then
16940           .cmdCreateEntries.SetFocus
16950         Else
16960           .cmdClose.SetFocus
16970         End If
16980       End With
16990     End Select
17000   End If

        ' ** Shift keys.
17010   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17020     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17030       With Me
17040         intRetVal = 0
17050         If .cmbTaxCodes_Interest.Visible = True And .cmbTaxCodes_Interest.Enabled = True Then
17060           .cmbTaxCodes_Interest.SetFocus
17070         ElseIf .cmbRevenueCodes_Interest.Visible = True And .cmbRevenueCodes_Interest.Enabled = True Then
17080           .cmbRevenueCodes_Interest.SetFocus
17090         Else
17100           .AccrualDate.SetFocus
17110         End If
17120       End With
17130     End Select
17140   End If

EXITP:
17150   KeyCode = intRetVal
17160   Exit Sub

ERRH:
17170   intRetVal = 0
17180   Select Case ERR.Number
        Case Else
17190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17200   End Select
17210   Resume EXITP

End Sub

Private Sub ReinvestPerShare_AfterUpdate()

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "ReinvestPerShare_AfterUpdate"

17310   With Me
17320     If IsNumeric(.ReinvestPerShare.text) Then
17330       If CDbl(.ReinvestPerShare.text) >= 100000 Then
17340         MsgBox "Per Share is too large.", vbInformation + vbOKOnly, "Invalid Entry"
17350         .ReinvestPerShare = vbNullString
17360         DoCmd.CancelEvent
17370         .ReinvestPerShare.SetFocus
17380       End If
17390     Else
17400       If .ReinvestPerShare.text <> "." Then
17410         MsgBox "You must enter a numeric value.", vbInformation + vbOKOnly, "Invalid Entry"
17420         .ReinvestPerShare = vbNullString
17430         DoCmd.CancelEvent
17440         .ReinvestPerShare.SetFocus
17450       End If
17460     End If
17470   End With

EXITP:
17480   Exit Sub

ERRH:
17490   Select Case ERR.Number
        Case Else
17500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17510   End Select
17520   Resume EXITP

End Sub

Private Sub ReinvestPerShare_Exit(Cancel As Integer)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "ReinvestPerShare_Exit"

        Dim dblFullValue As Double
        Dim dblTruncatedValue As Double

17610   With Me
17620     If IsNull(.ReinvestPerShare) = False Then
17630       If IsNumeric(.ReinvestPerShare) = True Then
17640         dblFullValue = .ReinvestPerShare
17650         dblTruncatedValue = Val(Format(dblFullValue, "0.00000"))
17660         If dblFullValue <> dblTruncatedValue Then
17670           Cancel = -1
17680           MsgBox "You can only enter a price with up to 5 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
17690           .ReinvestPerShare = Null
17700         Else
17710           If dblFullValue = 0 Then
17720             Cancel = -1
17730             MsgBox "Price per share cannot be zero.", vbInformation + vbOKOnly, "Entry Required"
17740             .ReinvestPerShare = Null
17750           End If
17760         End If
17770       End If
17780     End If
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

Private Sub ReinvestComments_GotFocus()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "ReinvestComments_GotFocus"

17910   With Me
17920     .ReinvestComments.SelLength = 0
17930     .ReinvestComments.SelStart = 99
17940   End With

EXITP:
17950   Exit Sub

ERRH:
17960   Select Case ERR.Number
        Case Else
17970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17980   End Select
17990   Resume EXITP

End Sub

Private Sub ReinvestComments_KeyDown(KeyCode As Integer, Shift As Integer)

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "ReinvestComments_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18020   intShiftDown = (Shift And acShiftMask) > 0
18030   intAltDown = (Shift And acAltMask) > 0
18040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18070       With Me
18080         intRetVal = 0
18090         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
18100           .cmbRevenueCodes_Purchase.SetFocus
18110         ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
18120           .cmbTaxCodes_Purchase.SetFocus
18130         ElseIf .cmdCreateEntries.Enabled = True Then
18140           .cmdCreateEntries.SetFocus
18150         Else
18160           .cmdClose.SetFocus
18170         End If
18180       End With
18190     End Select
18200   End If

        ' ** Shift keys.
18210   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18220     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18230       With Me
18240         intRetVal = 0
18250         If .ReinvestPerShare.Visible = True And .ReinvestPerShare.Enabled = True Then
18260           .ReinvestPerShare.SetFocus
18270         Else
18280           .chkReinvest.SetFocus
18290         End If
18300       End With
18310     End Select
18320   End If

EXITP:
18330   KeyCode = intRetVal
18340   Exit Sub

ERRH:
18350   intRetVal = 0
18360   Select Case ERR.Number
        Case Else
18370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18380   End Select
18390   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_GotFocus()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_GotFocus"

18410   With Me
18420     .cmbRevenueCodes_Purchase.SetFocus
18430   End With

EXITP:
18440   Exit Sub

ERRH:
18450   Select Case ERR.Number
        Case Else
18460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18470   End Select
18480   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18520   intShiftDown = (Shift And acShiftMask) > 0
18530   intAltDown = (Shift And acAltMask) > 0
18540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18570       With Me
18580         intRetVal = 0
18590         .cmbRevenueCodes_Purchase.SetFocus
18600       End With
18610     End Select
18620   End If

        ' ** Shift keys.
18630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18650       With Me
18660         intRetVal = 0
18670         If .ReinvestComments.Enabled = True Then
18680           .ReinvestComments.SetFocus
18690         Else
18700           .chkReinvest.SetFocus
18710         End If
18720       End With
18730     End Select
18740   End If

EXITP:
18750   KeyCode = intRetVal
18760   Exit Sub

ERRH:
18770   intRetVal = 0
18780   Select Case ERR.Number
        Case Else
18790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18800   End Select
18810   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_Enter()

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_Enter"

18910   With Me
18920     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
18930       ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
18940     Case False
18950       If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
18960         ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
18970       End If
18980     End Select
18990   End With

EXITP:
19000   Exit Sub

ERRH:
19010   Select Case ERR.Number
        Case Else
19020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19030   End Select
19040   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_AfterUpdate()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_AfterUpdate"

19110   With Me
19120     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
19130       .cmbRevenueCodes_Purchase = REVID_INC
19140     Case False
            ' ** Nothing at the moment.
19150     End Select
19160     If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
19170       .cmbRevenueCodes_Purchase_display = Null
19180     Else
19190       .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
19200     End If
19210     lngRevID_Purchase = .cmbRevenueCodes_Purchase
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

Private Sub cmbRevenueCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
19370       With Me
19380         intRetVal = 0
19390         If .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
19400           .cmbTaxCodes_Purchase.SetFocus
19410         ElseIf .cmdGenerateDates.Enabled = True Then
19420           .cmdGenerateDates.SetFocus
19430         Else
19440           .cmdClose.SetFocus
19450         End If
19460       End With
19470     End Select
19480   End If

        ' ** Shift keys.
19490   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19500     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19510       With Me
19520         intRetVal = 0
19530         If .ReinvestComments.Enabled = True Then
19540           .ReinvestComments.SetFocus
19550         Else
19560           .chkReinvest.SetFocus
19570         End If
19580       End With
19590     End Select
19600   End If

EXITP:
19610   KeyCode = intRetVal
19620   Exit Sub

ERRH:
19630   intRetVal = 0
19640   Select Case ERR.Number
        Case Else
19650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19660   End Select
19670   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_GotFocus()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_GotFocus"

19710   With Me
19720     .cmbTaxCodes_Purchase.SetFocus
19730   End With

EXITP:
19740   Exit Sub

ERRH:
19750   Select Case ERR.Number
        Case Else
19760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19770   End Select
19780   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19820   intShiftDown = (Shift And acShiftMask) > 0
19830   intAltDown = (Shift And acAltMask) > 0
19840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19870       With Me
19880         intRetVal = 0
19890         .cmbTaxCodes_Purchase.SetFocus
19900       End With
19910     End Select
19920   End If

        ' ** Shift keys.
19930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19950       With Me
19960         intRetVal = 0
19970         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
19980           .cmbRevenueCodes_Purchase.SetFocus
19990         ElseIf .ReinvestComments.Enabled = True Then
20000           .ReinvestComments.SetFocus
20010         Else
20020           .chkReinvest.SetFocus
20030         End If
20040       End With
20050     End Select
20060   End If

EXITP:
20070   KeyCode = intRetVal
20080   Exit Sub

ERRH:
20090   intRetVal = 0
20100   Select Case ERR.Number
        Case Else
20110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20120   End Select
20130   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_Enter()

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_Enter"

20210   With Me
20220     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
20230       ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
20240     Case False
20250       If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
20260         ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
20270       End If
20280     End Select
20290   End With

EXITP:
20300   Exit Sub

ERRH:
20310   Select Case ERR.Number
        Case Else
20320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20330   End Select
20340   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_AfterUpdate()

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_AfterUpdate"

20410   With Me
20420     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
20430       .cmbTaxCodes_Purchase = TAXID_INC
20440     Case False
            ' ** Nothing at the moment.
20450     End Select
20460     If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
20470       .cmbTaxCodes_Purchase_display = Null
20480     Else
20490       .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
20500     End If
20510     lngTaxID_Purchase = .cmbTaxCodes_Purchase
20520   End With

EXITP:
20530   Exit Sub

ERRH:
20540   Select Case ERR.Number
        Case Else
20550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20560   End Select
20570   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20620   intShiftDown = (Shift And acShiftMask) > 0
20630   intAltDown = (Shift And acAltMask) > 0
20640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20670       With Me
20680         intRetVal = 0
20690         If .cmdCreateEntries.Enabled = True Then
20700           .cmdCreateEntries.SetFocus
20710         Else
20720           .cmdClose.SetFocus
20730         End If
20740       End With
20750     End Select
20760   End If

        ' ** Shift keys.
20770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20790       With Me
20800         intRetVal = 0
20810         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
20820           .cmbRevenueCodes_Purchase.SetFocus
20830         ElseIf .ReinvestComments.Enabled = True Then
20840           .ReinvestComments.SetFocus
20850         Else
20860           .chkReinvest.SetFocus
20870         End If
20880       End With
20890     End Select
20900   End If

EXITP:
20910   KeyCode = intRetVal
20920   Exit Sub

ERRH:
20930   intRetVal = 0
20940   Select Case ERR.Number
        Case Else
20950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20960   End Select
20970   Resume EXITP

End Sub

Private Sub cmdCreateEntries_Click()

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCreateEntries_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstDate As DAO.Recordset, rstMasterAsset As DAO.Recordset, rstInterest As DAO.Recordset
        Dim rstWork As DAO.Recordset, rstJournal As DAO.Recordset
        Dim datAssetDate As Date, dblPerShare As Double
        Dim msgResponse As VbMsgBoxResult
        Dim dblTmp01 As Double, dblTmp02 As Double
        Dim blnContinue As Boolean

21010   blnContinue = True

21020   With Me

21030     Select Case IsNull(.AccrualDate)
          Case True
21040       blnContinue = False
21050       MsgBox "You must enter a Posting Date to continue.", vbInformation + vbOKOnly, "Entry Required"
21060     Case False
21070       Select Case IsDate(.AccrualDate)
            Case True
21080         If .frmAccruedIncome_Sub.Form.tmpRate_tot = 0 Then
21090           blnContinue = False
21100           MsgBox "A Daily Factor must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
21110         End If
21120       Case False
21130         blnContinue = False
21140         MsgBox "Please enter a valid Posting Date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
21150       End Select
21160     End Select

21170     If blnContinue = True Then

21180       DoCmd.Hourglass True
21190       DoEvents

21200       lngRevID_Interest = .cmbRevenueCodes_Interest
21210       lngTaxID_Interest = .cmbTaxCodes_Interest
21220       lngRevID_Purchase = .cmbRevenueCodes_Purchase
21230       lngTaxID_Purchase = .cmbTaxCodes_Purchase
21240       dblPerShare = .ReinvestPerShare

21250       Set dbs = CurrentDb
21260       Set qdf = dbs.QueryDefs("qryAccruedIncome_03")
21270       Set rstDate = qdf.OpenRecordset
21280       datAssetDate = CDate(.AccrualDate.Value) + time()

21290       DoCmd.Hourglass False
21300       msgResponse = MsgBox("Do you want to create journal entries with a posting date of" & vbCrLf & vbCrLf & _
              "    " & CStr(datAssetDate) & "  ?" & vbCrLf & vbCrLf & _
              "If you want a different date, click 'No' and change posting date to desired date.", _
              vbQuestion + vbYesNo, "Confirm New Entries")
21310       If msgResponse = vbYes Then

21320         DoCmd.Hourglass True
21330         DoEvents

              ' ** MasterAsset, just 'Accrued Interest Asset'.
              ' ** There is (should be) only one asset -- pre-defined -- with masterasset_TYPE = 'IA'.
21340         Set qdf = dbs.QueryDefs("qryAccruedIncome_04")
21350         Set rstMasterAsset = qdf.OpenRecordset
21360         If rstMasterAsset.BOF = True And rstMasterAsset.EOF = True Then
21370           blnContinue = False
21380           DoCmd.Hourglass False
21390           rstMasterAsset.Close
21400           rstDate.Clone
21410           dbs.Close
21420           MsgBox "Interest Asset not found!", vbCritical + vbOKOnly, "Interest Asset Not Found"
21430         Else

                ' ** Empty tmpEdit03.
21440           Set qdf = dbs.QueryDefs("qryAccruedIncome_06")
21450           qdf.Execute

                ' ** Get a list of all accounts which had some face amount
                ' ** (of the 'IA' asset) at the beginning of the selected month.

                ' ** Append ActiveAssets to tmpEdit3, by specified [astno].
21460           Set qdf = dbs.QueryDefs("qryAccruedIncome_10")
21470           With qdf.Parameters
21480             ![astno] = rstMasterAsset![assetno]
21490           End With
21500           qdf.Execute

                ' ** Update tmpEdit03, set Interest = 0.
21510           Set qdf = dbs.QueryDefs("qryAccruedIncome_11")
21520           qdf.Execute

                ' ** tmpEdit03, sorted.
21530           Set qdf = dbs.QueryDefs("qryAccruedIncome_12")
21540           Set rstInterest = qdf.OpenRecordset

                ' ** Calculate beginning amount.
21550           Set rstWork = dbs.OpenRecordset("SELECT ledger.* FROM ledger WHERE assetno = " & CStr(rstMasterAsset![assetno]) & " AND " & _
                  "assetdate >= #" & CStr(datStartDate) & "# AND journaltype IN ('Deposit','Purchase','Withdrawn','Sold') " & _
                  "ORDER BY assetdate DESC;", dbOpenSnapshot)
21560           If rstWork.RecordCount > 0 Then
21570             rstWork.MoveFirst
21580             Do While Not rstWork.EOF
21590               rstInterest.FindFirst "accountno = '" & rstWork![accountno] & "'"
21600               If rstInterest.NoMatch Then  ' ** None of that asset left in ActiveAssets nor added already in this loop.
21610                 rstInterest.AddNew
21620                 rstInterest![accountno] = rstWork![accountno]
21630                 rstInterest![OnHand] = 0
21640                 rstInterest![Interest] = 0
21650               Else  ' ** Already some from activeassets or added in this loop.
21660                 rstInterest.Edit
21670               End If
21680               Select Case rstWork![journaltype]
                    Case "Deposit", "Purchase"
21690                 rstInterest![OnHand] = rstInterest![OnHand] - rstWork![shareface]
21700               Case "Withdrawn", "Sold"
21710                 rstInterest![OnHand] = rstInterest![OnHand] + rstWork![shareface]
21720               End Select
21730               rstInterest.Update  ' ** Save edit/add.
21740               rstWork.MoveNext
21750             Loop
21760           End If

                ' ** rstInterest/tmpEdit03 now contains the beginning ShareFace as OnHand for each account for the Interest Asset.
                ' ** Now process through each day, FORWARD, first accounting for any Deposits/Withdrawals;
                ' ** THEN accumulate the interest in Interest.

21770           rstDate.MoveFirst
21780           Do While Not rstDate.EOF

                  ' ** First do any Deposits/Withdrawals for the day.
21790             rstWork.FindFirst "Int(assetdate) = #" & CStr(rstDate![tmpDate]) & "#"
21800             If Not rstWork.NoMatch Then  ' ** Are some to process.
21810               Do While (Not rstWork.EOF)
21820                 rstInterest.FindFirst "accountno = '" & rstWork![accountno] & "'"
21830                 If rstInterest.NoMatch Then  ' ** Impossible; this is an error.
21840                 Else
21850                   rstInterest.Edit
21860                 End If
21870                 Select Case rstWork![journaltype]
                      Case "Deposit", "Purchase"
21880                   rstInterest![OnHand] = rstInterest![OnHand] + rstWork![shareface]
21890                 Case "Withdrawn", "Sold"
21900                   rstInterest![OnHand] = rstInterest![OnHand] - rstWork![shareface]
21910                 End Select
21920                 rstInterest.Update  ' ** Save edit.
21930                 rstWork.MoveNext
21940                 If Not rstWork.EOF Then
21950                   If (Int(rstWork![assetdate]) <> rstDate![tmpDate]) Then
21960                     rstWork.MoveLast
21970                     rstWork.MoveNext  ' ** To EOF.
21980                   End If
21990                 End If
22000               Loop
22010             End If

                  ' ** Now calculate the interest.
22020             If rstInterest.RecordCount > 0 Then
22030               rstInterest.MoveFirst
22040               Do While Not rstInterest.EOF
22050                 rstInterest.Edit
22060                 rstInterest![Interest] = rstInterest![Interest] + Round(rstInterest![OnHand] * rstDate![tmpRate], 9)
22070                 rstInterest.Update
22080                 rstInterest.MoveNext
22090               Loop
22100             End If

22110             rstDate.MoveNext

22120           Loop

                ' ** Finally, create journal records.
22130           Set rstJournal = dbs.OpenRecordset("journal", dbOpenDynaset, dbConsistent)

22140           With rstJournal

                  ' ** Make sure tmpEdit03 (rstInterest) has records appended above.
22150             If rstInterest.RecordCount > 0 Then
22160               rstInterest.MoveFirst
22170               Do While Not rstInterest.EOF

22180                 rstInterest.Edit
22190                 rstInterest![Interest] = dbl_Round(rstInterest![Interest], 2)  ' ** Module Function: modUtilities.
22200                 rstInterest.Update
22210                 If rstInterest![Interest] > 0 Then

                        ' ** Create Interest record in rstJournal.
22220                   .AddNew
22230                   ![assetno] = rstMasterAsset![assetno]
22240                   ![accountno] = rstInterest![accountno]
22250                   ![journaltype] = "Interest"
22260                   ![ICash] = rstInterest![Interest]
22270                   ![PCash] = 0
22280                   ![Cost] = 0
22290                   ![transdate] = CDate(Me.AccrualDate)
22300                   ![shareface] = rstInterest![OnHand]
22310                   ![description] = "Accrued Income"
22320                   Select Case Me.chkReinvest
                        Case True
22330                     ![Reinvested] = True
22340                   Case False
22350                     ![Reinvested] = False
22360                   End Select
22370                   ![revcode_ID] = lngRevID_Interest
22380                   ![taxcode] = lngTaxID_Interest
22390                   ![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
22400                   .Update

22410                   If Me.chkReinvest = True Then
22420                     .AddNew
22430                     ![assetno] = rstMasterAsset![assetno]
22440                     ![accountno] = rstInterest![accountno]
22450                     ![journaltype] = "Purchase"
22460                     ![ICash] = (-1# * rstInterest![Interest])
22470                     ![PCash] = 0@
22480                     dblTmp01 = dbl_Round(CDbl(rstInterest![Interest]), 2)  ' ** Module Function: modUtilities.
22490                     ![Cost] = dblTmp01
22500                     ![transdate] = CDate(Me.AccrualDate)
22510                     ![assetdate] = datAssetDate
22520                     dblTmp02 = dbl_Round(CDbl((rstInterest![Interest] / dblPerShare)), 4)  ' ** Module Function: modUtilities.
22530                     ![shareface] = dblTmp02
22540                     If dblTmp02 > 0# Then
22550                       ![pershare] = dblTmp01 / dblTmp02
22560                     End If
22570                     If IsNull(Me.ReinvestComments) = False Then
22580                       If Trim(Me.ReinvestComments) <> vbNullString Then
22590                         ![description] = Me.ReinvestComments
22600                       End If
22610                     End If
22620                     ![revcode_ID] = lngRevID_Purchase
22630                     ![taxcode] = lngTaxID_Purchase
22640                     ![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
22650                     .Update
22660                   End If

22670                 End If

22680                 rstInterest.MoveNext
22690               Loop
22700             End If

22710             DoCmd.Hourglass False
22720             If .RecordCount > 0 Then
22730               MsgBox CStr(.RecordCount) & " journal entries created." & vbCrLf & vbCrLf & _
                      "Please check them for accuracy!", vbInformation + vbOKOnly, ("Finished" & Space(40))
22740               DoCmd.Close acForm, THIS_NAME
22750             Else
22760               MsgBox "No journal entries to create!", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
22770             End If

22780             .Close
22790           End With  ' ** rstJournal.

22800           rstWork.Close
22810           rstInterest.Close
22820           rstMasterAsset.Close
22830           rstDate.Close
22840           dbs.Close

22850         End If
22860       Else
22870         rstDate.Close
22880         dbs.Close
22890       End If

22900     End If  ' ** blnContinue.

22910   End With  ' ** Me.

22920   DoCmd.Hourglass False

EXITP:
22930   Set rstJournal = Nothing
22940   Set rstWork = Nothing
22950   Set rstInterest = Nothing
22960   Set rstMasterAsset = Nothing
22970   Set rstDate = Nothing
22980   Set qdf = Nothing
22990   Set dbs = Nothing
23000   Exit Sub

ERRH:
23010   DoCmd.Hourglass False
23020   Select Case ERR.Number
        Case 91  ' ** Object variable or With block variable not set.
          ' ** Do nothing; just recordset-closing stuff.
23030   Case Else
23040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23050   End Select
23060   Resume EXITP

End Sub

Private Sub cmdCreateEntries_KeyDown(KeyCode As Integer, Shift As Integer)

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCreateEntries_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23120   intShiftDown = (Shift And acShiftMask) > 0
23130   intAltDown = (Shift And acAltMask) > 0
23140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23160     Select Case intRetVal
          Case vbKeyTab
23170       With Me
23180         intRetVal = 0
23190         .cmdClose.SetFocus
23200       End With
23210     End Select
23220   End If

        ' ** Shift keys.
23230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23240     Select Case intRetVal
          Case vbKeyTab
23250       With Me
23260         intRetVal = 0
23270         If .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
23280           .cmbTaxCodes_Purchase.SetFocus
23290         ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
23300           .cmbRevenueCodes_Purchase.SetFocus
23310         ElseIf .ReinvestComments.Enabled = True Then
23320           .ReinvestComments.SetFocus
23330         Else
23340           .chkReinvest.SetFocus
23350         End If
23360       End With
23370     End Select
23380   End If

EXITP:
23390   KeyCode = intRetVal
23400   Exit Sub

ERRH:
23410   intRetVal = 0
23420   Select Case ERR.Number
        Case Else
23430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23440   End Select
23450   Resume EXITP

End Sub

Private Sub HideIncExp_Purch()

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "HideIncExp_Purch"

        Dim lngOffset1 As Long, lngOffset2 As Long
        Dim lngTmp01 As Long

23510   With Me

23520     .cmbRevenueCodes_Purchase_box.Visible = False
23530     .cmbRevenueCodes_Purchase_box2.Visible = False
23540     .cmbRevenueCodes_Purchase_box_lbl.Visible = False
23550     .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Visible = False
23560     .cmbRevenueCodes_Purchase_hline01.Visible = False
23570     .cmbRevenueCodes_Purchase_hline02.Visible = False
23580     .cmbRevenueCodes_Purchase_hline03.Visible = False
23590     .cmbRevenueCodes_Purchase_vline01.Visible = False
23600     .cmbRevenueCodes_Purchase_vline02.Visible = False
23610     .cmbRevenueCodes_Purchase_vline03.Visible = False
23620     .cmbRevenueCodes_Purchase_vline04.Visible = False
23630     .cmbRevenueCodes_Purchase.Visible = False
23640     .cmbRevenueCodes_Purchase_display.Visible = False
23650     .cmbRevenueCodes_Purchase_lbl.Visible = False
23660     .cmbRevenueCodes_Purchase_lbl2.Visible = False
23670     .cmbTaxCodes_Purchase.Visible = False
23680     .cmbTaxCodes_Purchase_display.Visible = False
23690     .cmbTaxCodes_Purchase_lbl.Visible = False
23700     .cmbTaxCodes_Purchase_lbl2.Visible = False

23710     lngOffset1 = ((.cmbRevenueCodes_Interest_box.Top + .cmbRevenueCodes_Interest_box.Height) - .cmbRevenueCodes_Purchase_box.Height)
23720     lngTmp01 = (.Detail.Height - (.cmbRevenueCodes_Purchase_box.Top + .cmbRevenueCodes_Purchase_box.Height))
23730     lngTmp01 = ((.chkReinvest_box.Top + .chkReinvest_box.Height) + lngTmp01)
23740     lngOffset2 = (.Detail.Height - lngTmp01)

23750     .cmbRevenueCodes_Purchase_box.Top = .cmbRevenueCodes_Purchase_box.Top - lngOffset1
23760     .cmbRevenueCodes_Purchase_box2.Top = .cmbRevenueCodes_Purchase_box2.Top - lngOffset1
23770     .cmbRevenueCodes_Purchase_box_lbl.Top = .cmbRevenueCodes_Purchase_box_lbl.Top - lngOffset1
23780     .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Top = .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Top - lngOffset1
23790     .cmbRevenueCodes_Purchase_hline01.Top = .cmbRevenueCodes_Purchase_hline01.Top - lngOffset1
23800     .cmbRevenueCodes_Purchase_hline02.Top = .cmbRevenueCodes_Purchase_hline02.Top - lngOffset1
23810     .cmbRevenueCodes_Purchase_hline03.Top = .cmbRevenueCodes_Purchase_hline03.Top - lngOffset1
23820     .cmbRevenueCodes_Purchase_vline01.Top = .cmbRevenueCodes_Purchase_vline01.Top - lngOffset1
23830     .cmbRevenueCodes_Purchase_vline02.Top = .cmbRevenueCodes_Purchase_vline02.Top - lngOffset1
23840     .cmbRevenueCodes_Purchase_vline03.Top = .cmbRevenueCodes_Purchase_vline03.Top - lngOffset1
23850     .cmbRevenueCodes_Purchase_vline04.Top = .cmbRevenueCodes_Purchase_vline04.Top - lngOffset1
23860     .cmbRevenueCodes_Purchase.Top = .cmbRevenueCodes_Purchase.Top - lngOffset1
23870     .cmbRevenueCodes_Purchase_display.Top = .cmbRevenueCodes_Purchase_display.Top - lngOffset1
23880     .cmbRevenueCodes_Purchase_lbl.Top = .cmbRevenueCodes_Purchase_lbl.Top - lngOffset1
23890     .cmbRevenueCodes_Purchase_lbl2.Top = .cmbRevenueCodes_Purchase_lbl2.Top - lngOffset1
23900     .cmbTaxCodes_Purchase.Top = .cmbTaxCodes_Purchase.Top - lngOffset1
23910     .cmbTaxCodes_Purchase_display.Top = .cmbTaxCodes_Purchase_display.Top - lngOffset1
23920     .cmbTaxCodes_Purchase_lbl.Top = .cmbTaxCodes_Purchase_lbl.Top - lngOffset1
23930     .cmbTaxCodes_Purchase_lbl2.Top = .cmbTaxCodes_Purchase_lbl2.Top - lngOffset1

23940     .detail_height_line.Top = .detail_height_line.Top - lngOffset2
23950     .Detail.Height = .Detail.Height - lngOffset2

23960     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
23970     lngMonitorNum = 1&: lngTmp01 = 0&
23980     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
23990     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

24000     If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
24010     DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, (lngFrm_Height - lngOffset2)  'lngFrm_Top
24020     If lngMonitorNum > 1& Then
24030       LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
24040     End If

24050     blnHideIncExp_Purch = True
24060     DoEvents

24070   End With

EXITP:
24080   Exit Sub

ERRH:
24090   Select Case ERR.Number
        Case Else
24100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24110   End Select
24120   Resume EXITP

End Sub

Private Sub HideIncExp_Int()

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "HideIncExp_Int"

        Dim lngOffset1 As Long, lngOffset2 As Long
        Dim lngTmp01 As Long

24210   With Me

24220     GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

24230     .cmbRevenueCodes_Interest_box.Visible = False
24240     .cmbRevenueCodes_Interest_box2.Visible = False
24250     .cmbRevenueCodes_Interest_box_lbl.Visible = False
24260     .cmbRevenueCodes_Interest_box_lbl_dim_hi.Visible = False
24270     .cmbRevenueCodes_Interest_hline01.Visible = False
24280     .cmbRevenueCodes_Interest_hline02.Visible = False
24290     .cmbRevenueCodes_Interest_hline03.Visible = False
24300     .cmbRevenueCodes_Interest_vline01.Visible = False
24310     .cmbRevenueCodes_Interest_vline02.Visible = False
24320     .cmbRevenueCodes_Interest_vline03.Visible = False
24330     .cmbRevenueCodes_Interest_vline04.Visible = False
24340     .cmbRevenueCodes_Interest.Visible = False
24350     .cmbRevenueCodes_Interest_display.Visible = False
24360     .cmbRevenueCodes_Interest_lbl.Visible = False
24370     .cmbRevenueCodes_Interest_lbl2.Visible = False
24380     .cmbTaxCodes_Interest.Visible = False
24390     .cmbTaxCodes_Interest_display.Visible = False
24400     .cmbTaxCodes_Interest_lbl.Visible = False
24410     .cmbTaxCodes_Interest_lbl2.Visible = False

24420     lngOffset1 = .chkReinvest_box.Top - (.cmbRevenueCodes_Interest_hline01.Top + lngTpp)
24430     lngTmp01 = (.Detail.Height - (.chkReinvest_box.Top + .chkReinvest_box.Height))
24440     lngTmp01 = (((.cmbRevenueCodes_Interest_hline01.Top + lngTpp) + .chkReinvest_box.Height) + lngTmp01)
24450     lngOffset2 = (.Detail.Height - lngTmp01)

24460     .chkReinvest_box.Top = .chkReinvest_box.Top - lngOffset1
24470     .ReinvestPerShare.Top = .ReinvestPerShare.Top - lngOffset1
24480     .ReinvestPerShare_lbl.Top = .ReinvestPerShare_lbl.Top - lngOffset1
24490     .chkReinvest.Top = .chkReinvest.Top - lngOffset1
24500     .chkReinvest_lbl.Top = .chkReinvest_lbl.Top - lngOffset1
24510     .ReinvestComments.Top = .ReinvestComments.Top - lngOffset1
24520     .ReinvestComments_lbl.Top = .ReinvestComments_lbl.Top - lngOffset1

24530     .detail_height_line.Top = .detail_height_line.Top - lngOffset2
24540     .Detail.Height = .Detail.Height - lngOffset2

24550     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
24560     lngMonitorNum = 1&: lngTmp01 = 0&
24570     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
24580     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

24590     If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
24600     DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, (lngFrm_Height - lngOffset2)  'lngFrm_Top
24610     If lngMonitorNum > 1& Then
24620       LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
24630     End If

24640     blnHideIncExp_Int = True
24650     DoEvents

24660   End With

EXITP:
24670   Exit Sub

ERRH:
24680   Select Case ERR.Number
        Case Else
24690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24700   End Select
24710   Resume EXITP

End Sub
