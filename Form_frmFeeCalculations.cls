VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFeeCalculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmFeeCalculations"

'VGC 04/15/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   HideIncExp_Rec()
' ##   HideIncExp_Paid()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Tax Codes Rcvd:   A {cmbTaxCodes_Received}
' **   Close:            C {cmdClose}
' **   Posting Date:     D {PostDate}
' **   I/E Codes Rcvd:   E {cmbRevenueCodes_Received}
' **   Create Jrnl Ent:  J {cmdCreateEntries}
' **   Calc Month:       M {FeeMonth}
' **   I/E Codes Paid:   N {cmbRevenueCodes_Paid}
' **   Income:           O {opgBasis_optIncome}
' **   Print:            P {cmdPrint}
' **   Tax Codes Paid:   T {cmbTaxCodes_Paid}
' **   Preview:          W {cmdPreview}
' **   Year:             Y {FeeYear}
' **   Exit:             X {cmdClose}
' **   50%/50%:          5 {opgBasis_opt5050}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar}

' ** CLOSES WHEN FINISHED!

' ** Fee Frequency:
' **   1 = Monthly
' **   2 = Quarterly
' **   3 = Semiannually
' **   4 = Annually

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long

Private datStartDate As Date, datEndDate As Date
Private blnHideIncExp_Rec As Boolean, blnHideIncExp_Paid As Boolean, blnHasForEx As Boolean
Private blnHasCalculated As Boolean, blnFromCalendar As Boolean, blnKeyDown_Return As Boolean
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long ', lngRecsCur As Long
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim datPostingDate As Date

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           blnHasCalculated = False: blnHideIncExp_Rec = False: blnHideIncExp_Paid = False
170           gstrFormQuerySpec = THIS_NAME  ' ** For some reason, this gets Nulled out by the time cmdPreview/cmdPrint is clicked!
180           blnFromCalendar = False: blnKeyDown_Return = False: blnHasForEx = False
190           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
200           datStartDate = 0: datEndDate = 0

              ' ** Copyright Lebans Holdings 1999 Ltd.
              ' ** Create an instance of the Calendar class.
210           Set clsMonthClass = New clsMonthCal
              ' ** You MUST SET the class hWndForm prop!!!
220           clsMonthClass.hWndForm = Me.hwnd
              ' ** Let's default to PositionAtCursor.
230           clsMonthClass.PositionAtCursor = True

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
240           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

              ' ** Variables are fed empty, then populated ByRef.
250           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

260           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

270           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
280           If lngMonitorCnt > 1& Then
290             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
300             lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
310             If lngMonitorNum > 1& Then
320               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
330             End If
340           End If

350           HideIncExp_Rec  ' ** Procedure: Below.
360           HideIncExp_Paid  ' ** Procedure: Below.

370           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

380           .cmdCreateEntries.Enabled = False

390           Set dbs = CurrentDb
400           With dbs

410             Set grstPostingDate = .OpenRecordset("PostingDate", dbOpenDynaset, dbConsistent)
420             With grstPostingDate
430               .FindFirst "[Username] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
440               If .NoMatch = False Then
450                 datPostingDate = ![Posting_Date]
460               Else
470                 .AddNew
480                 ![Posting_Date] = Date
490                 ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
500                 .Update
510                 datPostingDate = Date
520               End If
530               .Close
540             End With

                ' ** Though all the queries are done with and without accountno criteria,
                ' ** it looks like only the without ones are used.

                ' ** By specified [actno]:
                'qryFeeCalculations_01aa

                ' ** All accounts:
                ' ** qryFeeCalculations_01c (qryFeeCalculations_01b (Account, linked to
                ' ** ActiveAssets, all accounts, with TotalCost_usd), grouped by curr_id,
                ' ** just curr_id <> 150), grouped, with cnt.
550             Set qdf = .QueryDefs("qryFeeCalculations_01ca")
560             Set rst = qdf.OpenRecordset
570             With rst
580               If .BOF = True And .EOF = True Then
                    ' ** No foreign currencies whatsoever.
590               Else
600                 .MoveFirst
610                 If ![cnt] = 0 Then
                      ' ** No foreign currencies whatsoever.
620                 Else
630                   blnHasForEx = True
640                 End If
650               End If
660               .Close
670             End With
680             Set rst = Nothing
690             Set qdf = Nothing

700             .Close
710           End With  ' ** dbs.
720           Set dbs = Nothing

730           .HasForEx = blnHasForEx

740           opgBasis_AfterUpdate  ' ** Procedure: Below.
750           chkPageOf_AfterUpdate  ' ** Procedure: Below.

760           blnIsOpen = False

770         Else
780           Cancel = -1
790         End If
800       Else
810         Cancel = -1
820       End If
830     End With

840     If Cancel = -1 Then
850       Beep
860       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
870     End If

EXITP:
880     Set rst = Nothing
890     Set qdf = Nothing
900     Set dbs = Nothing
910     Exit Sub

ERRH:
920     Select Case ERR.Number
        Case Else
930       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
940     End Select
950     Resume EXITP

End Sub

Private Sub Form_Load()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1010    With Me

          ' ** Paid is Expense.
1020      Select Case IsNull(.cmbRevenueCodes_Paid)
          Case True
1030        .cmbRevenueCodes_Paid = REVID_EXP
1040        .cmbRevenueCodes_Paid_display = Null
1050      Case False
1060        If .cmbRevenueCodes_Paid = REVID_INC Or .cmbRevenueCodes_Paid = REVID_EXP Then
1070          .cmbRevenueCodes_Paid_display = Null
1080        Else
1090          .cmbRevenueCodes_Paid_display = .cmbRevenueCodes_Paid.Column(1)
1100        End If
1110      End Select

          ' ** Paid is Deduction.
1120      Select Case IsNull(.cmbTaxCodes_Paid)
          Case True
1130        .cmbTaxCodes_Paid = TAXID_DED
1140        .cmbTaxCodes_Paid_display = Null
1150      Case False
1160        If .cmbTaxCodes_Paid = TAXID_INC Or .cmbTaxCodes_Paid = TAXID_DED Then
1170          .cmbTaxCodes_Paid_display = Null
1180        Else
1190          .cmbTaxCodes_Paid_display = .cmbTaxCodes_Paid.Column(1)
1200        End If
1210      End Select

          ' ** Received is Income.
1220      Select Case IsNull(.cmbRevenueCodes_Received)
          Case True
1230        .cmbRevenueCodes_Received = REVID_INC
1240        .cmbRevenueCodes_Received_display = Null
1250      Case False
1260        If .cmbRevenueCodes_Received = REVID_INC Or .cmbRevenueCodes_Received = REVID_EXP Then
1270          .cmbRevenueCodes_Received_display = Null
1280        Else
1290          .cmbRevenueCodes_Received_display = .cmbRevenueCodes_Received.Column(1)
1300        End If
1310      End Select

          ' ** Received is Income.
1320      Select Case IsNull(.cmbTaxCodes_Received)
          Case True
1330        .cmbTaxCodes_Received = TAXID_INC
1340        .cmbTaxCodes_Received_display = Null
1350      Case False
1360        If .cmbTaxCodes_Received = TAXID_INC Or .cmbTaxCodes_Received = TAXID_DED Then
1370          .cmbTaxCodes_Received_display = Null
1380        Else
1390          .cmbTaxCodes_Received_display = .cmbTaxCodes_Received.Column(1)
1400        End If
1410      End Select

1420      Select Case gblnRevenueExpenseTracking
          Case True
1430        If blnHideIncExp_Paid = False Then
1440          .cmbRevenueCodes_Paid.Visible = True
1450          .cmbRevenueCodes_Paid_display.Visible = True
1460          .cmbRevenueCodes_Paid_lbl2.Visible = False  ' ** Option off msg.
1470        End If  ' ** blnHideIncExp_Paid.
1480        If blnHideIncExp_Rec = False Then
1490          .cmbRevenueCodes_Received.Visible = True
1500          .cmbRevenueCodes_Received_display.Visible = True
1510          .cmbRevenueCodes_Received_lbl2.Visible = False  ' ** Option off msg.
1520        End If  ' ** blnHideIncExp_Rec.
1530      Case False
1540        If blnHideIncExp_Paid = False Then
1550          .cmbRevenueCodes_Paid.Visible = False
1560          .cmbRevenueCodes_Paid_display.Visible = False
1570          .cmbRevenueCodes_Paid_lbl2.Visible = True  ' ** Option off msg.
1580        End If  ' ** blnHideIncExp_Paid.
1590        If blnHideIncExp_Rec = False Then
1600          .cmbRevenueCodes_Received.Visible = False
1610          .cmbRevenueCodes_Received_display.Visible = False
1620          .cmbRevenueCodes_Received_lbl2.Visible = True  ' ** Option off msg.
1630        End If  ' ** blnHideIncExp_Rec.
1640      End Select

1650      Select Case gblnIncomeTaxCoding
          Case True
1660        If blnHideIncExp_Paid = False Then
1670          .cmbTaxCodes_Paid.Visible = True
1680          .cmbTaxCodes_Paid_display.Visible = True
1690          .cmbTaxCodes_Paid_lbl2.Visible = False  ' ** Option off msg.
1700        End If  ' ** blnHideIncExp_Paid.
1710        If blnHideIncExp_Rec = False Then
1720          .cmbTaxCodes_Received.Visible = True
1730          .cmbTaxCodes_Received_display.Visible = True
1740          .cmbTaxCodes_Received_lbl2.Visible = False  ' ** Option off msg.
1750        End If  ' ** blnHideIncExp_Rec.
1760      Case False
1770        If blnHideIncExp_Paid = False Then
1780          .cmbTaxCodes_Paid.Visible = False
1790          .cmbTaxCodes_Paid_display.Visible = False
1800          .cmbTaxCodes_Paid_lbl2.Visible = True
1810        End If  ' ** blnHideIncExp_Paid.
1820        If blnHideIncExp_Rec = False Then
1830          .cmbTaxCodes_Received.Visible = False
1840          .cmbTaxCodes_Received_display.Visible = False
1850          .cmbTaxCodes_Received_lbl2.Visible = True
1860        End If  ' ** blnHideIncExp_Rec.
1870      End Select

1880      If gblnRevenueExpenseTracking = False And gblnIncomeTaxCoding = False Then
1890        If blnHideIncExp_Paid = False Then
1900          .cmbRevenueCodes_Paid_box_lbl.ForeColor = WIN_CLR_DISF
1910          .cmbRevenueCodes_Paid_box_lbl_dim_hi.Visible = True
1920        End If  ' ** blnHideIncExp_Paid.
1930        If blnHideIncExp_Rec = False Then
1940          .cmbRevenueCodes_Received_box_lbl.ForeColor = WIN_CLR_DISF
1950          .cmbRevenueCodes_Received_box_lbl_dim_hi.Visible = True
1960        End If  ' ** blnHideIncExp_Rec.
1970      End If

1980      If IsNull(.FeeMonth) = True And IsNull(.FeeYear) = True Then
1990        .postdate.Enabled = False
2000        .postdate.BackStyle = acBackStyleTransparent
2010        .postdate.BorderColor = WIN_CLR_DISR
2020        .cmdCalendar.Enabled = False
2030        .cmdCalendar_raised_img_dis.Visible = True
2040        .cmdCalendar_raised_img.Visible = False
2050        .cmdCalendar_raised_semifocus_dots_img.Visible = False
2060        .cmdCalendar_raised_focus_img.Visible = False
2070        .cmdCalendar_raised_focus_dots_img.Visible = False
2080        .cmdCalendar_sunken_focus_dots_img.Visible = False
2090        .opgBasis.Enabled = False
2100        .chkPageOf.Enabled = False
2110        .chkPageOf_lbl.Visible = False
2120        .chkPageOf_lbl2_dim.Visible = True
2130        .chkPageOf_lbl2_dim_hi.Visible = True
2140        .cmdPreview.Enabled = False
2150        .cmdPrint.Enabled = False
2160        .cmdPreviewPrint_lbl.ForeColor = WIN_CLR_DISF
2170        .cmdPreviewPrint_lbl_dim_hi.Visible = True
2180      End If

2190    End With

EXITP:
2200    Exit Sub

ERRH:
2210    Select Case ERR.Number
        Case Else
2220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2230    End Select
2240    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2320    intShiftDown = (Shift And acShiftMask) > 0
2330    intAltDown = (Shift And acAltMask) > 0
2340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Shortcut F keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar}

        ' ** Plain keys.
2350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2360      Select Case intRetVal
          Case vbKeyEscape
2370        intRetVal = 0
2380        cmdClose_Click  ' ** Procedure: Below.
2390      Case vbKeyF7
2400        With Me
2410          intRetVal = 0
2420          .cmdCalendar.SetFocus
2430          cmdCalendar_Click  ' ** Procedure: Below.
2440        End With
2450      End Select
2460    End If

        ' ** Alt keys.
2470    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2480      Select Case intRetVal
          Case vbKeyX
2490        intRetVal = 0
2500        cmdClose_Click  ' ** Procedure: Below.
2510      End Select
2520    End If

        ' ** Ctrl keys.
2530    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2540      Select Case intRetVal
          Case vbKeyG
2550        With Me
2560          intRetVal = 0
2570          If .chkPageOf.Enabled = True Then
2580            .chkPageOf.SetFocus
2590            .chkPageOf = (Not .chkPageOf)
2600            chkPageOf_AfterUpdate  ' ** Procedure: Below.
2610          End If
2620        End With
2630      End Select
2640    End If

        ' ** Ctrl-Shift keys.
2650    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2660      Select Case intRetVal
          Case vbKeyF
2670        With Me
2680          intRetVal = 0
2690          .FocusHolder.SetFocus
2700        End With
2710      End Select
2720    End If

EXITP:
2730    KeyCode = intRetVal
2740    Exit Sub

ERRH:
2750    intRetVal = 0
2760    THAT_PROC = THIS_PROC
2770    That_Erl = Erl
2780    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2790    Resume EXITP

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
2870        Do Until strFormName = THIS_NAME
2880          DoCmd.SelectObject acForm, THIS_NAME, False
2890          strFormName = Screen.ActiveForm.Name
2900        Loop
2910  On Error GoTo ERRH
2920      End If
2930      .FeeMonth.SetFocus
2940      If gblnGoToReport = True Then
2950        gblnGoToReport = False
2960        .cmdClose.SetFocus
2970        Beep
2980        DoCmd.Hourglass False
2990        DoEvents
3000      End If
3010    End With

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3110    If DataErr <> 0 Then
3120      Select Case DataErr
          Case 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
3130      Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
            ' ** Probably a bad date. Ignore.
3140        Response = acDataErrContinue
3150      Case 3162
3160        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
3170        DoCmd.CancelEvent
3180      Case 3129
3190        MsgBox "You must select a valid month to continue.", vbInformation + vbOKOnly, "Invalid Entry"
3200        Me.FeeMonth.SetFocus
3210      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3220        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3230        Response = acDataErrContinue
3240        DoCmd.CancelEvent
3250      Case Else
3260        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3270        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3280      End Select
3290    End If

EXITP:
3300    Exit Sub

ERRH:
3310    Select Case ERR.Number
        Case Else
3320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3330    End Select
3340    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

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

Private Sub cmdClose_Click()

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
4190          .FeeMonth.SetFocus
4200        End With
4210      End Select
4220    End If

        ' ** Shift keys.
4230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4240      Select Case intRetVal
          Case vbKeyTab
4250        With Me
4260          intRetVal = 0
4270          If .cmdCreateEntries.Enabled = True Then
4280            .cmdCreateEntries.SetFocus
4290          ElseIf .cmbTaxCodes_Received.Visible = True And .cmbTaxCodes_Received.Enabled = True Then
4300            .cmbTaxCodes_Received.SetFocus
4310          ElseIf .cmbRevenueCodes_Received.Visible = True And .cmbRevenueCodes_Received.Enabled = True Then
4320            .cmbRevenueCodes_Received.SetFocus
4330          Else
4340            .cmdPrint.SetFocus
4350          End If
4360        End With
4370      End Select
4380    End If

EXITP:
4390    KeyCode = intRetVal
4400    Exit Sub

ERRH:
4410    intRetVal = 0
4420    THAT_PROC = THIS_PROC
4430    That_Erl = Erl
4440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4450    Resume EXITP

End Sub

Private Sub FeeMonth_AfterUpdate()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "FeeMonth_AfterUpdate"

4510    With Me
4520      If .FeeMonth.Column(1) = "December" And month(Date) = 1 Then
4530        .FeeYear = year(Date) - 1
4540      Else
4550        .FeeYear = year(Date)
4560      End If
4570      DoEvents
4580      If IsNull(.FeeMonth) = False And IsNull(.FeeYear) = False Then
4590        If .postdate.Enabled = False Then
4600          .postdate.Enabled = True
4610          .postdate.BackStyle = acBackStyleNormal
4620          .postdate.BorderColor = CLR_LTBLU2
4630          .cmdCalendar.Enabled = True
4640          .cmdCalendar_raised_img.Visible = True
4650          .cmdCalendar_raised_semifocus_dots_img.Visible = False
4660          .cmdCalendar_raised_focus_img.Visible = False
4670          .cmdCalendar_raised_focus_dots_img.Visible = False
4680          .cmdCalendar_sunken_focus_dots_img.Visible = False
4690          .cmdCalendar_raised_img_dis.Visible = False
4700          .opgBasis.Enabled = True
4710          .chkPageOf.Enabled = True
4720          .chkPageOf_lbl.Visible = True
4730          .chkPageOf_lbl2_dim.Visible = False
4740          .chkPageOf_lbl2_dim_hi.Visible = False
4750          .cmdPreview.Enabled = True
4760          .cmdPrint.Enabled = True
4770          .cmdPreviewPrint_lbl.ForeColor = CLR_DKGRY2
4780          .cmdPreviewPrint_lbl_dim_hi.Visible = False
4790        End If
4800      End If
4810    End With

EXITP:
4820    Exit Sub

ERRH:
4830    Select Case ERR.Number
        Case Else
4840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4850    End Select
4860    Resume EXITP

End Sub

Private Sub FeeMonth_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "FeeMonth_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4920    intShiftDown = (Shift And acShiftMask) > 0
4930    intAltDown = (Shift And acAltMask) > 0
4940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4970        With Me
4980          intRetVal = 0
4990          .FeeYear.SetFocus
5000        End With
5010      End Select
5020    End If

        ' ** Shift keys.
5030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5050        With Me
5060          intRetVal = 0
5070          .cmdClose.SetFocus
5080        End With
5090      End Select
5100    End If

EXITP:
5110    KeyCode = intRetVal
5120    Exit Sub

ERRH:
5130    intRetVal = 0
5140    THAT_PROC = THIS_PROC
5150    That_Erl = Erl
5160    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5170    Resume EXITP

End Sub

Private Sub FeeYear_AfterUpdate()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "FeeYear_AfterUpdate"

5210    With Me
5220      If IsNull(.FeeMonth) = False And IsNull(.FeeYear) = False Then
5230        If .postdate.Enabled = False Then
5240          .postdate.Enabled = True
5250          .postdate.BackStyle = acBackStyleNormal
5260          .postdate.BorderColor = CLR_LTBLU2
5270          .cmdCalendar.Enabled = True
5280          .cmdCalendar_raised_img.Visible = True
5290          .cmdCalendar_raised_semifocus_dots_img.Visible = False
5300          .cmdCalendar_raised_focus_img.Visible = False
5310          .cmdCalendar_raised_focus_dots_img.Visible = False
5320          .cmdCalendar_sunken_focus_dots_img.Visible = False
5330          .cmdCalendar_raised_img_dis.Visible = False
5340          .opgBasis.Enabled = True
5350          .chkPageOf.Enabled = True
5360          .chkPageOf_lbl.Visible = True
5370          .chkPageOf_lbl2_dim.Visible = False
5380          .chkPageOf_lbl2_dim_hi.Visible = False
5390          .cmdPreview.Enabled = True
5400          .cmdPrint.Enabled = True
5410          .cmdPreviewPrint_lbl.ForeColor = CLR_DKGRY2
5420          .cmdPreviewPrint_lbl_dim_hi.Visible = False
5430        End If
5440      End If
5450    End With

EXITP:
5460    Exit Sub

ERRH:
5470    Select Case ERR.Number
        Case Else
5480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5490    End Select
5500    Resume EXITP

End Sub

Private Sub PostDate_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_box_MouseMove"

5610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5620    Exit Sub

ERRH:
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Private Sub PostDate_Enter()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_Enter"

5710    With Me
5720      If blnFromCalendar = False Then
5730        If IsDate(.postdate) = False Then
5740          .postdate = vbNullString
5750        End If
5760      Else
5770        blnFromCalendar = False
5780        .postdate.SelLength = 0
5790        .postdate.SelStart = 99
5800      End If
5810    End With

EXITP:
5820    Exit Sub

ERRH:
5830    DoCmd.Hourglass False
5840    Select Case ERR.Number
        Case Else
5850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5860    End Select
5870    Resume EXITP

End Sub

Private Sub PostDate_BeforeUpdate(Cancel As Integer)
' ** I think there's a delay when getting the date from the popup calendar,
' ** causing an erroneous message about future dates.

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_BeforeUpdate"

        Dim datTmp01 As Date

5910    With Me
5920      If .postdate.text <> vbNullString Then
5930        If IsDate(.postdate.text) = True Then
              ' ** Allow 1 month from today.
5940          If CDate(.postdate.text) <= Date Then
                ' ** This is fine.
5950          Else
5960            datTmp01 = DateAdd("m", 1, Date)
5970            If CDate(.postdate.text) <= datTmp01 Then
                  ' ** This is OK, too.
5980            Else
5990              DoCmd.Hourglass False
6000              MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
6010              Cancel = -1
6020            End If
6030          End If
6040        End If
6050      End If
6060    End With

EXITP:
6070    Exit Sub

ERRH:
6080    DoCmd.Hourglass False
6090    Select Case ERR.Number
        Case Else
6100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6110    End Select
6120    Resume EXITP

End Sub

Private Sub PostDate_DblClick(Cancel As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_DblClick"

6210    With Me
6220      .postdate = Date
6230      If .cmdPreview.Enabled = True Then
6240        .cmdPreview.SetFocus
6250      End If
6260    End With

EXITP:
6270    Exit Sub

ERRH:
6280    Select Case ERR.Number
        Case Else
6290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6300    End Select
6310    Resume EXITP

End Sub

Private Sub PostDate_AfterUpdate()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_AfterUpdate"

6410    With Me
6420      If IsDate(.postdate) = True And blnHasCalculated = True Then
6430        .cmdCreateEntries.Enabled = True
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

Private Sub PostDate_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6620    intShiftDown = (Shift And acShiftMask) > 0
6630    intAltDown = (Shift And acAltMask) > 0
6640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6670        With Me
6680          intRetVal = 0
6690          blnKeyDown_Return = True
6700          .opgBasis.SetFocus
6710        End With
6720      End Select
6730    End If

        ' ** Shift keys.
6740    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6750      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6760        With Me
6770          intRetVal = 0
6780          blnKeyDown_Return = True
6790          .FeeYear.SetFocus
6800        End With
6810      End Select
6820    End If

EXITP:
6830    KeyCode = intRetVal
6840    Exit Sub

ERRH:
6850    intRetVal = 0
6860    THAT_PROC = THIS_PROC
6870    That_Erl = Erl
6880    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6890    Resume EXITP

End Sub

Private Sub PostDate_Change()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_Change"

6910    With Me
6920      If IsDate(.postdate) = True And blnHasCalculated = True Then
6930        .cmdCreateEntries.Enabled = True
6940      End If
6950    End With

EXITP:
6960    Exit Sub

ERRH:
6970    Select Case ERR.Number
        Case Else
6980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6990    End Select
7000    Resume EXITP

End Sub

Private Sub PostDate_KeyPress(KeyAscii As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_KeyPress"

7110    With Me
7120      If IsDate(.postdate) = True And blnHasCalculated = True Then
7130        .cmdCreateEntries.Enabled = True
7140      End If
7150    End With

EXITP:
7160    Exit Sub

ERRH:
7170    Select Case ERR.Number
        Case Else
7180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7190    End Select
7200    Resume EXITP

End Sub

Private Sub PostDate_KeyUp(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_KeyUp"

7310    With Me
7320      If IsDate(.postdate) = True And blnHasCalculated = True Then
7330        .cmdCreateEntries.Enabled = True
7340      End If
7350    End With

EXITP:
7360    Exit Sub

ERRH:
7370    Select Case ERR.Number
        Case Else
7380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7390    End Select
7400    Resume EXITP

End Sub

Private Sub PostDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_MouseMove"

7510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7520    Exit Sub

ERRH:
7530    Select Case ERR.Number
        Case Else
7540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7550    End Select
7560    Resume EXITP

End Sub

Private Sub PostDate_Exit(Cancel As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_Exit"

        Dim blnJustMove As Boolean

7610    With Me
7620      blnJustMove = False
7630      If IsNull(.postdate) = False Then
7640        If .postdate <> vbNullString Then
              ' ** Proceed.
7650          blnJustMove = True
7660        End If
7670      End If
7680      If blnJustMove = False Then
7690        If blnKeyDown_Return = True Then
7700          blnKeyDown_Return = False
7710          .postdate = Date
7720        Else
7730          If IsNull(.postdate.text) = False Then
7740            If Trim(.postdate.text) <> vbNullString Then
                  ' ** Proceed.
7750            Else
                  ' ** Empty, do nothing.
7760            End If
7770          Else
                ' ** Empty, do nothing.
7780          End If
7790        End If
7800      End If
7810    End With

EXITP:
7820    Exit Sub

ERRH:
7830    DoCmd.Hourglass False
7840    Select Case ERR.Number
        Case Else
7850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7860    End Select
7870    Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

7910    With Me
7920      blnCalendar1_Focus = True
7930      .cmdCalendar_raised_semifocus_dots_img.Visible = True
7940      .cmdCalendar_raised_img.Visible = False
7950      .cmdCalendar_raised_focus_img.Visible = False
7960      .cmdCalendar_raised_focus_dots_img.Visible = False
7970      .cmdCalendar_sunken_focus_dots_img.Visible = False
7980      .cmdCalendar_raised_img_dis.Visible = False
7990    End With

EXITP:
8000    Exit Sub

ERRH:
8010    Select Case ERR.Number
        Case Else
8020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8030    End Select
8040    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

8110    With Me
8120      blnCalendar1_MouseDown = True
8130      .cmdCalendar_sunken_focus_dots_img.Visible = True
8140      .cmdCalendar_raised_img.Visible = False
8150      .cmdCalendar_raised_semifocus_dots_img.Visible = False
8160      .cmdCalendar_raised_focus_img.Visible = False
8170      .cmdCalendar_raised_focus_dots_img.Visible = False
8180      .cmdCalendar_raised_img_dis.Visible = False
8190    End With

EXITP:
8200    Exit Sub

ERRH:
8210    Select Case ERR.Number
        Case Else
8220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8230    End Select
8240    Resume EXITP

End Sub

Private Sub cmdCalendar_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim blnRetVal As Boolean

8310    With Me
8320      datStartDate = Date
8330      datEndDate = 0
8340      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
8350      If blnRetVal = True Then
            ' ** Allow posting up to 1 month into the future.
8360        If datStartDate > DateAdd("m", 1, Date) Then
8370          MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
8380          .postdate = CDate(Format(Date, "mm/dd/yyyy"))
8390        Else
8400          blnFromCalendar = True
8410          .postdate = datStartDate
8420        End If
8430      Else
8440        blnFromCalendar = True
8450        .postdate = CDate(Format(Date, "mm/dd/yyyy"))
8460      End If
8470      .postdate.SetFocus
8480    End With

EXITP:
8490    Exit Sub

ERRH:
8500    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
8510    Case Else
8520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8530    End Select
8540    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

8610    With Me
8620      If blnCalendar1_MouseDown = False Then
8630        Select Case blnCalendar1_Focus
            Case True
8640          .cmdCalendar_raised_focus_dots_img.Visible = True
8650          .cmdCalendar_raised_focus_img.Visible = False
8660        Case False
8670          .cmdCalendar_raised_focus_img.Visible = True
8680          .cmdCalendar_raised_focus_dots_img.Visible = False
8690        End Select
8700        .cmdCalendar_raised_img.Visible = False
8710        .cmdCalendar_raised_semifocus_dots_img.Visible = False
8720        .cmdCalendar_sunken_focus_dots_img.Visible = False
8730        .cmdCalendar_raised_img_dis.Visible = False
8740      End If
8750    End With

EXITP:
8760    Exit Sub

ERRH:
8770    Select Case ERR.Number
        Case Else
8780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8790    End Select
8800    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

8910    With Me
8920      .cmdCalendar_raised_focus_dots_img.Visible = True
8930      .cmdCalendar_raised_img.Visible = False
8940      .cmdCalendar_raised_semifocus_dots_img.Visible = False
8950      .cmdCalendar_raised_focus_img.Visible = False
8960      .cmdCalendar_sunken_focus_dots_img.Visible = False
8970      .cmdCalendar_raised_img_dis.Visible = False
8980      blnCalendar1_MouseDown = False
8990    End With

EXITP:
9000    Exit Sub

ERRH:
9010    Select Case ERR.Number
        Case Else
9020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9030    End Select
9040    Resume EXITP

End Sub

Private Sub cmdCalendar_LostFocus()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

9110    With Me
9120      .cmdCalendar_raised_img.Visible = True
9130      .cmdCalendar_raised_semifocus_dots_img.Visible = False
9140      .cmdCalendar_raised_focus_img.Visible = False
9150      .cmdCalendar_raised_focus_dots_img.Visible = False
9160      .cmdCalendar_sunken_focus_dots_img.Visible = False
9170      .cmdCalendar_raised_img_dis.Visible = False
9180      blnCalendar1_Focus = False
9190    End With

EXITP:
9200    Exit Sub

ERRH:
9210    Select Case ERR.Number
        Case Else
9220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9230    End Select
9240    Resume EXITP

End Sub

Private Sub opgBasis_AfterUpdate()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgBasis_AfterUpdate"

9310    With Me
9320      Select Case .opgBasis
          Case .opgBasis_optIncome.OptionValue
9330        .opgBasis_optIncome_lbl.FontBold = True
9340        .opgBasis_opt5050_lbl.FontBold = False
9350      Case .opgBasis_opt5050.OptionValue
9360        .opgBasis_optIncome_lbl.FontBold = False
9370        .opgBasis_opt5050_lbl.FontBold = True
9380      End Select
9390      If .cmdCreateEntries.Enabled = True Then
9400        .cmdCreateEntries.Enabled = False
9410      End If
9420    End With

EXITP:
9430    Exit Sub

ERRH:
9440    Select Case ERR.Number
        Case Else
9450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9460    End Select
9470    Resume EXITP

End Sub

Private Sub opgBasis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgBasis_MouseMove"

9510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9520    Exit Sub

ERRH:
9530    Select Case ERR.Number
        Case Else
9540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9550    End Select
9560    Resume EXITP

End Sub

Private Sub opgBasis_optIncome_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgBasis_optIncome_KeyDown"

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
9690          If .cmdPreview.Enabled = True Then
9700            .cmdPreview.SetFocus
9710          ElseIf .cmbRevenueCodes_Paid.Visible = True And .cmbRevenueCodes_Paid.Enabled = True Then
9720            .cmbRevenueCodes_Paid.SetFocus
9730          ElseIf .cmbTaxCodes_Paid.Visible = True And .cmbTaxCodes_Paid.Enabled = True Then
9740            .cmbTaxCodes_Paid.SetFocus
9750          ElseIf .cmdCreateEntries.Enabled = True Then
9760            .cmdCreateEntries.SetFocus
9770          Else
9780            .cmdClose.SetFocus
9790          End If
9800        End With
9810      End Select
9820    End If

        ' ** Shift keys.
9830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9850        With Me
9860          intRetVal = 0
9870          .postdate.SetFocus
9880        End With
9890      End Select
9900    End If

EXITP:
9910    KeyCode = intRetVal
9920    Exit Sub

ERRH:
9930    intRetVal = 0
9940    THAT_PROC = THIS_PROC
9950    That_Erl = Erl
9960    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9970    Resume EXITP

End Sub

Private Sub opgBasis_optIncome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgBasis_optIncome_MouseMove"

10010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10020   Exit Sub

ERRH:
10030   Select Case ERR.Number
        Case Else
10040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10050   End Select
10060   Resume EXITP

End Sub

Private Sub opgBasis_opt5050_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgBasis_opt5050_KeyDown"

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
10190         If .cmdPreview.Enabled = True Then
10200           .cmdPreview.SetFocus
10210         ElseIf .cmbRevenueCodes_Paid.Visible = True And .cmbRevenueCodes_Paid.Enabled = True Then
10220           .cmbRevenueCodes_Paid.SetFocus
10230         ElseIf .cmbTaxCodes_Paid.Visible = True And .cmbTaxCodes_Paid.Enabled = True Then
10240           .cmbTaxCodes_Paid.SetFocus
10250         ElseIf .cmdCreateEntries.Enabled = True Then
10260           .cmdCreateEntries.SetFocus
10270         Else
10280           .cmdClose.SetFocus
10290         End If
10300       End With
10310     End Select
10320   End If

        ' ** Shift keys.
10330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10350       With Me
10360         intRetVal = 0
10370         .postdate.SetFocus
10380       End With
10390     End Select
10400   End If

EXITP:
10410   KeyCode = intRetVal
10420   Exit Sub

ERRH:
10430   intRetVal = 0
10440   THAT_PROC = THIS_PROC
10450   That_Erl = Erl
10460   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10470   Resume EXITP

End Sub

Private Sub opgBasis_opt5050_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgBasis_opt5050_MouseMove"

10510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10520   Exit Sub

ERRH:
10530   Select Case ERR.Number
        Case Else
10540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10550   End Select
10560   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

10610   With Me
10620     Select Case .chkPageOf
          Case True
10630       .chkPageOf_lbl.FontBold = True
10640       .chkPageOf_lbl2_dim.FontBold = True
10650       .chkPageOf_lbl2_dim_hi.FontBold = True
10660     Case False
10670       .chkPageOf_lbl.FontBold = False
10680       .chkPageOf_lbl2_dim.FontBold = False
10690       .chkPageOf_lbl2_dim_hi.FontBold = False
10700     End Select
10710   End With

EXITP:
10720   Exit Sub

ERRH:
10730   Select Case ERR.Number
        Case Else
10740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10750   End Select
10760   Resume EXITP

End Sub

Private Sub cmdPreview_Click()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngFeeFreq As Long
        Dim strDocName As String
        Dim blnContinue As Boolean

10810   blnContinue = True

10820   With Me
10830     If IsNull(.FeeMonth) = False And IsNull(.FeeYear) = False Then

10840       DoCmd.Hourglass True
10850       DoEvents

10860       Select Case .FeeMonth.Column(0)
            Case 1, 2, 4, 5, 7, 8, 10, 11
              ' ** FeeFreq 1: Jan, Feb, Apr, May, Jul, Aug, Oct, Nov.
10870         lngFeeFreq = 1&
10880       Case 3, 9
              ' ** FeeFreq 1,2:  Mar, Sep.
10890         lngFeeFreq = 2&
10900       Case 6
              ' ** FeeFreq 1,2,3: Jun.
10910         lngFeeFreq = 3&
10920       Case 12
              ' ** FeeFreq 1,2,3,4: Dec.
10930         lngFeeFreq = 4&
10940       End Select
10950       .FeeFreq = lngFeeFreq

            ' ** I don't understand why this gets Nulled out!?
10960       If gstrFormQuerySpec = vbNullString Then
10970         gstrFormQuerySpec = THIS_NAME
10980       End If

10990       CalculateFees  ' ** Procedure: Below.

11000       Set dbs = CurrentDb
11010       With dbs
              ' ** Account, linked to qryFeeCalculations_02b (qryFeeCalculations_02c
              ' ** (qryFeeCalculations_01b (Account, linked to ActiveAssets, all accounts,
              ' ** with TotalCost_usd), linked to tblCurrency, with .._usd fields), grouped
              ' ** and summed), Schedule, ScheduleDetail, with CurrentTotalMarketValue_usd.
11020         Set qdf = dbs.QueryDefs("qryFeeCalculations_03b")
11030         Set rst = qdf.OpenRecordset
11040         If rst.BOF = True And rst.EOF = True Then
11050           rst.Close
11060           Set rst = Nothing
11070           Set qdf = Nothing
11080           blnContinue = False
11090           DoCmd.Hourglass False
11100           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
11110         Else
11120           rst.Close
11130           Set rst = Nothing
11140           Set qdf = Nothing
                ' ** qryFeeCalculations_04b (Account, linked to qryFeeCalculations_03b (Account, linked
                ' ** to qryFeeCalculations_02b (qryFeeCalculations_02c (qryFeeCalculations_01b (Account,
                ' ** linked to ActiveAssets, all accounts, TotalCost_usd), linked to tblCurrency, with .._usd
                ' ** fields), grouped and summed), Schedule, ScheduleDetail, with CurrentTotalMarketValue_usd),
                ' ** by specified [feefreq]), linked to FeeCalculations, Schedule.
11150           Set qdf = dbs.QueryDefs("qryFeeCalculations_06b")
11160           With qdf.Parameters
11170             ![FeeFreq] = lngFeeFreq
11180           End With
11190           Set rst = qdf.OpenRecordset
11200           With rst
11210             If .BOF = True And .EOF = True Then
11220               blnContinue = False
11230               DoCmd.Hourglass False
11240               MsgBox "There are no Accounts scheduled for Fees in the month of " & Me.FeeMonth.Column(1) & ".", _
                      vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
11250             End If
11260             .Close
11270           End With
11280           Set rst = Nothing
11290           Set qdf = Nothing
11300         End If
11310         .Close
11320       End With  ' ** dbs.
11330       Set dbs = Nothing

11340       If blnContinue = True Then
11350         strDocName = "rptFeeCalculations"
11360         gstrReportCallingForm = THIS_NAME  ' ** When not vbNullString, sets this form .Visible = False.
11370         DoCmd.OpenReport strDocName, acViewPreview
11380         DoCmd.Maximize
11390         DoCmd.RunCommand acCmdFitToWindow
11400         DoCmd.Hourglass False
11410         blnHasCalculated = True
11420         If IsDate(.postdate) = True Then
11430           .cmdCreateEntries.Enabled = True
11440         End If
11450       Else
11460         DoCmd.Hourglass False
11470       End If  ' ** blnContinue.

11480     Else
11490       If IsNull(.FeeMonth) = True Then
11500         MsgBox "You must select a valid month to continue.", vbInformation + vbOKOnly, "Entry Required"
11510       ElseIf IsNull(.FeeYear) = True Then
11520         MsgBox "You must select a valid year to continue.", vbInformation + vbOKOnly, "Entry Required"
11530       End If
11540     End If
11550   End With  ' ** Me.

EXITP:
11560   Set rst = Nothing
11570   Set qdf = Nothing
11580   Set dbs = Nothing
11590   Exit Sub

ERRH:
11600   If Forms(THIS_NAME).Visible = False Then
11610     Forms(THIS_NAME).Visible = True
11620   End If
11630   DoCmd.Hourglass False
11640   THAT_PROC = THIS_PROC
11650   That_Erl = Erl
11660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11670   Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

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
          Case vbKeyTab
11770       With Me
11780         intRetVal = 0
11790         .cmdPrint.SetFocus
11800       End With
11810     End Select
11820   End If

        ' ** Shift keys.
11830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11840     Select Case intRetVal
          Case vbKeyTab
11850       With Me
11860         intRetVal = 0
11870         .opgBasis.SetFocus
11880       End With
11890     End Select
11900   End If

EXITP:
11910   KeyCode = intRetVal
11920   Exit Sub

ERRH:
11930   intRetVal = 0
11940   THAT_PROC = THIS_PROC
11950   That_Erl = Erl
11960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11970   Resume EXITP

End Sub

Private Sub cmdPrint_Click()

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngFeeFreq As Long
        Dim strDocName As String
        Dim blnContinue As Boolean

12010   blnContinue = True

12020   With Me
12030     If IsNull(.FeeMonth) = False And IsNull(.FeeYear) = False Then

12040       DoCmd.Hourglass True
12050       DoEvents

12060       Select Case .FeeMonth.Column(0)
            Case 1, 2, 4, 5, 7, 8, 10, 11
              ' ** FeeFreq 1: Jan, Feb, Apr, May, Jul, Aug, Oct, Nov.
12070         lngFeeFreq = 1&
12080       Case 3, 9
              ' ** FeeFreq 1,2:  Mar, Sep.
12090         lngFeeFreq = 2&
12100       Case 6
              ' ** FeeFreq 1,2,3: Jun.
12110         lngFeeFreq = 3&
12120       Case 12
              ' ** FeeFreq 1,2,3,4: Dec.
12130         lngFeeFreq = 4&
12140       End Select
12150       .FeeFreq = lngFeeFreq

            ' ** I don't understand why this gets Nulled out!?
12160       If gstrFormQuerySpec = vbNullString Then
12170         gstrFormQuerySpec = THIS_NAME
12180       End If

12190       CalculateFees  ' ** Procedure: Below.

12200       Set dbs = CurrentDb
12210       With dbs
12220         Set qdf = .QueryDefs("qryFeeCalculations_06b")
12230         With qdf.Parameters
12240           ![FeeFreq] = lngFeeFreq
12250         End With
12260         Set rst = qdf.OpenRecordset
12270         With rst
12280           If rst.BOF = True And rst.EOF = True Then
12290             blnContinue = False
12300             DoCmd.Hourglass False
12310             MsgBox "There are no Accounts scheduled for Fees in the month of " & Me.FeeMonth.Column(1) & ".", _
                    vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
12320           End If
12330           .Close
12340         End With
12350         .Close
12360       End With  ' ** dbs.

12370       If blnContinue = True Then
12380         strDocName = "rptFeeCalculations"
12390         gstrReportCallingForm = vbNullString  ' ** When vbNullString, leaves this form .Visible = True.
12400         If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
12410           DoCmd.OpenReport strDocName, acViewPreview
12420           DoCmd.Maximize
12430           DoCmd.RunCommand acCmdFitToWindow
12440         Else
                '##GTR_Ref: rptFeeCalculations
12450           DoCmd.OpenReport strDocName, acViewNormal
12460         End If
12470         DoCmd.Hourglass False
12480         blnHasCalculated = True
12490         If IsDate(.postdate) Then
12500           .cmdCreateEntries.Enabled = True
12510         End If
12520       Else
12530         DoCmd.Hourglass False
12540       End If

12550     Else
12560       If IsNull(.FeeMonth) = True Then
12570         MsgBox "You must select a valid month to continue.", vbInformation + vbOKOnly, "Entry Required"
12580       ElseIf IsNull(.FeeYear) = True Then
12590         MsgBox "You must select a valid year to continue.", vbInformation + vbOKOnly, "Entry Required"
12600       End If
12610     End If
12620   End With  ' ** Me.

EXITP:
12630   Set rst = Nothing
12640   Set qdf = Nothing
12650   Set dbs = Nothing
12660   Exit Sub

ERRH:
12670   If Forms(THIS_NAME).Visible = False Then
12680     Forms(THIS_NAME).Visible = True
12690   End If
12700   DoCmd.Hourglass False
12710   THAT_PROC = THIS_PROC
12720   That_Erl = Erl
12730   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12740   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12820   intShiftDown = (Shift And acShiftMask) > 0
12830   intAltDown = (Shift And acAltMask) > 0
12840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12860     Select Case intRetVal
          Case vbKeyTab
12870       With Me
12880         intRetVal = 0
12890         If .cmbRevenueCodes_Paid.Visible = True And .cmbRevenueCodes_Paid.Enabled = True Then
12900           .cmbRevenueCodes_Paid.SetFocus
12910         ElseIf .cmbTaxCodes_Paid.Visible = True And .cmbTaxCodes_Paid.Enabled = True Then
12920           .cmbTaxCodes_Paid.SetFocus
12930         ElseIf .cmdCreateEntries.Enabled = True Then
12940           .cmdCreateEntries.SetFocus
12950         Else
12960           .cmdClose.SetFocus
12970         End If
12980       End With
12990     End Select
13000   End If

        ' ** Shift keys.
13010   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13020     Select Case intRetVal
          Case vbKeyTab
13030       With Me
13040         intRetVal = 0
13050         .cmdPreview.SetFocus
13060       End With
13070     End Select
13080   End If

EXITP:
13090   KeyCode = intRetVal
13100   Exit Sub

ERRH:
13110   intRetVal = 0
13120   THAT_PROC = THIS_PROC
13130   That_Erl = Erl
13140   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13150   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Paid_display_GotFocus()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Paid_display_GotFocus"

13210   With Me
13220     .cmbRevenueCodes_Paid.SetFocus
13230   End With

EXITP:
13240   Exit Sub

ERRH:
13250   Select Case ERR.Number
        Case Else
13260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13270   End Select
13280   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Paid_display_KeyDown(KeyCode As Integer, Shift As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Paid_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13320   intShiftDown = (Shift And acShiftMask) > 0
13330   intAltDown = (Shift And acAltMask) > 0
13340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13370       With Me
13380         intRetVal = 0
13390         .cmbRevenueCodes_Paid.SetFocus
13400       End With
13410     End Select
13420   End If

        ' ** Shift keys.
13430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13450       With Me
13460         intRetVal = 0
13470         If .cmdPreview.Enabled = True Then
13480           .cmdPreview.SetFocus
13490         ElseIf .opgBasis.Enabled = True Then
13500           .opgBasis.SetFocus
13510         ElseIf .postdate.Enabled = True Then
13520           .postdate.SetFocus
13530         ElseIf .FeeYear.Enabled = True Then
13540           .FeeYear.SetFocus
13550         ElseIf .FeeMonth.Enabled = True Then
13560           .FeeMonth.SetFocus
13570         Else
13580           .cmdClose.SetFocus
13590         End If
13600       End With
13610     End Select
13620   End If

EXITP:
13630   KeyCode = intRetVal
13640   Exit Sub

ERRH:
13650   intRetVal = 0
13660   Select Case ERR.Number
        Case Else
13670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13680   End Select
13690   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Paid_Enter()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Paid_Enter"

13710   With Me
13720     Select Case IsNull(.cmbRevenueCodes_Paid)
          Case True
13730       ExpandCombo .cmbRevenueCodes_Paid  ' ** Module Procedure: modUtilities.
13740     Case False
13750       If .cmbRevenueCodes_Paid = REVID_INC Or .cmbRevenueCodes_Paid = REVID_EXP Then
13760         ExpandCombo .cmbRevenueCodes_Paid  ' ** Module Procedure: modUtilities.
13770       End If
13780     End Select
13790   End With

EXITP:
13800   Exit Sub

ERRH:
13810   Select Case ERR.Number
        Case Else
13820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13830   End Select
13840   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Paid_AfterUpdate()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Paid_AfterUpdate"

13910   With Me
13920     Select Case IsNull(.cmbRevenueCodes_Paid)
          Case True
13930       .cmbRevenueCodes_Paid = REVID_EXP
13940     Case False
            ' ** Nothing at the moment.
13950     End Select
13960     If .cmbRevenueCodes_Paid = REVID_INC Or .cmbRevenueCodes_Paid = REVID_EXP Then
13970       .cmbRevenueCodes_Paid_display = Null
13980     Else
13990       .cmbRevenueCodes_Paid_display = .cmbRevenueCodes_Paid.Column(1)
14000     End If
14010   End With

EXITP:
14020   Exit Sub

ERRH:
14030   Select Case ERR.Number
        Case Else
14040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14050   End Select
14060   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Paid_KeyDown(KeyCode As Integer, Shift As Integer)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Paid_KeyDown"

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
14190         If .cmbTaxCodes_Paid.Visible = True And .cmbTaxCodes_Paid.Enabled = True Then
14200           .cmbTaxCodes_Paid.SetFocus
14210         Else
14220           .cmbRevenueCodes_Received.SetFocus
14230         End If
14240       End With
14250     End Select
14260   End If

        ' ** Shift keys.
14270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14290       With Me
14300         intRetVal = 0
14310         If .cmdPrint.Enabled = True Then
14320           .cmdPrint.SetFocus
14330         ElseIf .opgBasis.Enabled = True Then
14340           .opgBasis.SetFocus
14350         ElseIf .postdate.Enabled = True Then
14360           .postdate.SetFocus
14370         ElseIf .FeeYear.Enabled = True Then
14380           .FeeYear.SetFocus
14390         ElseIf .FeeMonth.Enabled = True Then
14400           .FeeMonth.SetFocus
14410         Else
14420           .cmdClose.SetFocus
14430         End If
14440       End With
14450     End Select
14460   End If

EXITP:
14470   KeyCode = intRetVal
14480   Exit Sub

ERRH:
14490   intRetVal = 0
14500   Select Case ERR.Number
        Case Else
14510     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14520   End Select
14530   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Paid_display_GotFocus()

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Paid_display_GotFocus"

14610   With Me
14620     .cmbTaxCodes_Paid.SetFocus
14630   End With

EXITP:
14640   Exit Sub

ERRH:
14650   Select Case ERR.Number
        Case Else
14660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14670   End Select
14680   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Paid_display_KeyDown(KeyCode As Integer, Shift As Integer)

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Paid_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14720   intShiftDown = (Shift And acShiftMask) > 0
14730   intAltDown = (Shift And acAltMask) > 0
14740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14770       With Me
14780         intRetVal = 0
14790         .cmbTaxCodes_Paid.SetFocus
14800       End With
14810     End Select
14820   End If

        ' ** Shift keys.
14830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14850       With Me
14860         intRetVal = 0
14870         If .cmbRevenueCodes_Paid.Visible = True And .cmbRevenueCodes_Paid.Enabled = True Then
14880           .cmbRevenueCodes_Paid.SetFocus
14890         Else
14900           .cmdPrint.SetFocus
14910         End If
14920       End With
14930     End Select
14940   End If

EXITP:
14950   KeyCode = intRetVal
14960   Exit Sub

ERRH:
14970   intRetVal = 0
14980   Select Case ERR.Number
        Case Else
14990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15000   End Select
15010   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Paid_Enter()

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Paid_Enter"

15110   With Me
15120     Select Case IsNull(.cmbTaxCodes_Paid)
          Case True
15130       ExpandCombo .cmbTaxCodes_Paid  ' ** Module Procedure: modUtilities.
15140     Case False
15150       If .cmbTaxCodes_Paid = TAXID_INC Or .cmbTaxCodes_Paid = TAXID_DED Then
15160         ExpandCombo .cmbTaxCodes_Paid  ' ** Module Procedure: modUtilities.
15170       End If
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

Private Sub cmbTaxCodes_Paid_AfterUpdate()

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Paid_AfterUpdate"

15310   With Me
15320     Select Case IsNull(.cmbTaxCodes_Paid)
          Case True
15330       .cmbTaxCodes_Paid = TAXID_DED
15340     Case False
            ' ** Nothing at the moment.
15350     End Select
15360     If .cmbTaxCodes_Paid = TAXID_INC Or .cmbTaxCodes_Paid = TAXID_DED Then
15370       .cmbTaxCodes_Paid_display = Null
15380     Else
15390       .cmbTaxCodes_Paid_display = .cmbTaxCodes_Paid.Column(1)
15400     End If
15410   End With

EXITP:
15420   Exit Sub

ERRH:
15430   Select Case ERR.Number
        Case Else
15440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15450   End Select
15460   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Paid_KeyDown(KeyCode As Integer, Shift As Integer)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Paid_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15520   intShiftDown = (Shift And acShiftMask) > 0
15530   intAltDown = (Shift And acAltMask) > 0
15540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15570       With Me
15580         intRetVal = 0
15590         If .cmbRevenueCodes_Received.Visible = True And .cmbRevenueCodes_Received.Enabled = True Then
15600           .cmbRevenueCodes_Received.SetFocus
15610         Else
15620           .cmbTaxCodes_Paid.SetFocus
15630         End If
15640       End With
15650     End Select
15660   End If

        ' ** Shift keys.
15670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15690       With Me
15700         intRetVal = 0
15710         If .cmbRevenueCodes_Paid.Visible = True And .cmbRevenueCodes_Paid.Enabled = True Then
15720           .cmbRevenueCodes_Paid.SetFocus
15730         Else
15740           .cmdPrint.SetFocus
15750         End If
15760       End With
15770     End Select
15780   End If

EXITP:
15790   KeyCode = intRetVal
15800   Exit Sub

ERRH:
15810   intRetVal = 0
15820   Select Case ERR.Number
        Case Else
15830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15840   End Select
15850   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Received_display_GotFocus()

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Received_display_GotFocus"

15910   With Me
15920     .cmbRevenueCodes_Received.SetFocus
15930   End With

EXITP:
15940   Exit Sub

ERRH:
15950   Select Case ERR.Number
        Case Else
15960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15970   End Select
15980   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Received_display_KeyDown(KeyCode As Integer, Shift As Integer)

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Received_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16020   intShiftDown = (Shift And acShiftMask) > 0
16030   intAltDown = (Shift And acAltMask) > 0
16040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16070       With Me
16080         intRetVal = 0
16090         .cmbRevenueCodes_Received.SetFocus
16100       End With
16110     End Select
16120   End If

        ' ** Shift keys.
16130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16150       With Me
16160         intRetVal = 0
16170         If .cmbTaxCodes_Paid.Visible = True And .cmbTaxCodes_Paid.Enabled = True Then
16180           .cmbTaxCodes_Paid.SetFocus
16190         Else
16200           .cmbRevenueCodes_Paid.SetFocus
16210         End If
16220       End With
16230     End Select
16240   End If

EXITP:
16250   KeyCode = intRetVal
16260   Exit Sub

ERRH:
16270   intRetVal = 0
16280   Select Case ERR.Number
        Case Else
16290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16300   End Select
16310   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Received_Enter()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Received_Enter"

16410   With Me
16420     Select Case IsNull(.cmbRevenueCodes_Received)
          Case True
16430       ExpandCombo .cmbRevenueCodes_Received  ' ** Module Procedure: modUtilities.
16440     Case False
16450       If .cmbRevenueCodes_Received = REVID_INC Or .cmbRevenueCodes_Received = REVID_EXP Then
16460         ExpandCombo .cmbRevenueCodes_Received  ' ** Module Procedure: modUtilities.
16470       End If
16480     End Select
16490   End With

EXITP:
16500   Exit Sub

ERRH:
16510   Select Case ERR.Number
        Case Else
16520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16530   End Select
16540   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Received_AfterUpdate()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Received_AfterUpdate"

16610   With Me
16620     Select Case IsNull(.cmbRevenueCodes_Received)
          Case True
16630       .cmbRevenueCodes_Received = REVID_INC
16640     Case False
            ' ** Nothing at the moment.
16650     End Select
16660     If .cmbRevenueCodes_Received = REVID_INC Or .cmbRevenueCodes_Received = REVID_EXP Then
16670       .cmbRevenueCodes_Received_display = Null
16680     Else
16690       .cmbRevenueCodes_Received_display = .cmbRevenueCodes_Received.Column(1)
16700     End If
16710   End With

EXITP:
16720   Exit Sub

ERRH:
16730   Select Case ERR.Number
        Case Else
16740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16750   End Select
16760   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Received_KeyDown(KeyCode As Integer, Shift As Integer)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Received_KeyDown"

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
16890         If .cmbTaxCodes_Received.Visible = True And .cmbTaxCodes_Received.Enabled = True Then
16900           .cmbTaxCodes_Received.SetFocus
16910         ElseIf .cmdCreateEntries.Enabled = True Then
16920           .cmdCreateEntries.SetFocus
16930         Else
16940           .cmdClose.SetFocus
16950         End If
16960       End With
16970     End Select
16980   End If

        ' ** Shift keys.
16990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17000     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17010       With Me
17020         intRetVal = 0
17030         If .cmbTaxCodes_Paid.Visible = True And .cmbTaxCodes_Paid.Enabled = True Then
17040           .cmbTaxCodes_Paid.SetFocus
17050         Else
17060           .cmbRevenueCodes_Paid.SetFocus
17070         End If
17080       End With
17090     End Select
17100   End If

EXITP:
17110   KeyCode = intRetVal
17120   Exit Sub

ERRH:
17130   intRetVal = 0
17140   Select Case ERR.Number
        Case Else
17150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17160   End Select
17170   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Received_display_GotFocus()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Received_display_GotFocus"

17210   With Me
17220     .cmbTaxCodes_Received.SetFocus
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

Private Sub cmbTaxCodes_Received_display_KeyDown(KeyCode As Integer, Shift As Integer)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Received_display_KeyDown"

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
17390         .cmbTaxCodes_Received.SetFocus
17400       End With
17410     End Select
17420   End If

        ' ** Shift keys.
17430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17450       With Me
17460         intRetVal = 0
17470         If .cmbRevenueCodes_Received.Visible = True And .cmbRevenueCodes_Received.Enabled = True Then
17480           .cmbRevenueCodes_Received.SetFocus
17490         Else
17500           .cmbTaxCodes_Paid.SetFocus
17510         End If
17520       End With
17530     End Select
17540   End If

EXITP:
17550   KeyCode = intRetVal
17560   Exit Sub

ERRH:
17570   intRetVal = 0
17580   Select Case ERR.Number
        Case Else
17590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17600   End Select
17610   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Received_Enter()

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Received_Enter"

17710   With Me
17720     Select Case IsNull(.cmbTaxCodes_Received)
          Case True
17730       ExpandCombo .cmbTaxCodes_Received  ' ** Module Procedure: modUtilities.
17740     Case False
17750       If .cmbTaxCodes_Received = TAXID_INC Or .cmbTaxCodes_Received = TAXID_DED Then
17760         ExpandCombo .cmbTaxCodes_Received  ' ** Module Procedure: modUtilities.
17770       End If
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

Private Sub cmbTaxCodes_Received_AfterUpdate()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Received_AfterUpdate"

17910   With Me
17920     Select Case IsNull(.cmbTaxCodes_Received)
          Case True
17930       .cmbTaxCodes_Received = TAXID_INC
17940     Case False
            ' ** Nothing at the moment.
17950     End Select
17960     If .cmbTaxCodes_Received = TAXID_INC Or .cmbTaxCodes_Received = TAXID_DED Then
17970       .cmbTaxCodes_Received_display = Null
17980     Else
17990       .cmbTaxCodes_Received_display = .cmbTaxCodes_Received.Column(1)
18000     End If
18010   End With

EXITP:
18020   Exit Sub

ERRH:
18030   Select Case ERR.Number
        Case Else
18040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18050   End Select
18060   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Received_KeyDown(KeyCode As Integer, Shift As Integer)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Received_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
18170       With Me
18180         intRetVal = 0
18190         If .cmdCreateEntries.Enabled = True Then
18200           .cmdCreateEntries.SetFocus
18210         Else
18220           .cmdClose.SetFocus
18230         End If
18240       End With
18250     End Select
18260   End If

        ' ** Shift keys.
18270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18290       With Me
18300         intRetVal = 0
18310         If .cmbRevenueCodes_Received.Visible = True And .cmbRevenueCodes_Received.Enabled = True Then
18320           .cmbRevenueCodes_Received.SetFocus
18330         Else
18340           .cmbTaxCodes_Paid.SetFocus
18350         End If
18360       End With
18370     End Select
18380   End If

EXITP:
18390   KeyCode = intRetVal
18400   Exit Sub

ERRH:
18410   intRetVal = 0
18420   Select Case ERR.Number
        Case Else
18430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18440   End Select
18450   Resume EXITP

End Sub

Private Sub cmdCreateEntries_Click()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCreateEntries_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim varTotal As Variant
        Dim varValue As Variant
        Dim varSuspense As Variant
        Dim lngRevCodeID_Inc As Long, lngRevCodeID_Exp As Long
        Dim lngTaxCodeID_Inc As Long, lngTaxCodeID_Ded As Long
        Dim lngRecs As Long, lngFeeFreq As Long
        Dim lngX As Long
        Dim blnContinue As Boolean

18510   blnContinue = True

18520   With Me

18530     DoCmd.Hourglass True
18540     DoEvents

18550     Select Case .FeeMonth.Column(0)
          Case 1, 2, 4, 5, 7, 8, 10, 11
            ' ** FeeFreq 1: Jan, Feb, Apr, May, Jul, Aug, Oct, Nov.
18560       lngFeeFreq = 1&
18570     Case 3, 9
            ' ** FeeFreq 1,2:  Mar, Sep.
18580       lngFeeFreq = 2&
18590     Case 6
            ' ** FeeFreq 1,2,3: Jun.
18600       lngFeeFreq = 3&
18610     Case 12
            ' ** FeeFreq 1,2,3,4: Dec.
18620       lngFeeFreq = 4&
18630     End Select
18640     .FeeFreq = lngFeeFreq

18650     Set dbs = CurrentDb
          ' ** qryFeeCalculations_04b (Account, linked to qryFeeCalculations_03b (Account, linked to
          ' ** qryFeeCalculations_02b (qryFeeCalculations_01b (Account, linked to ActiveAssets, all accounts),
          ' ** grouped and summed), Schedule, ScheduleDetail), by specified [feefreq]),
          ' ** linked to FeeCalculations, Schedule, grouped and summed.
18660     Set qdf = dbs.QueryDefs("qryFeeCalculations_07b")  ' ** Criteria is: <=[feefreq].
18670     With qdf.Parameters
18680       ![FeeFreq] = lngFeeFreq
18690     End With
18700     Set rst1 = qdf.OpenRecordset
18710     Set rst2 = dbs.OpenRecordset("journal", dbOpenDynaset, dbConsistent)

18720     varSuspense = gblnAccountNoWithType

18730     If rst1.EOF = True And rst1.BOF = True Then
18740       blnContinue = False
18750       rst1.Close
18760       rst2.Close
18770       dbs.Close
18780       DoCmd.Hourglass False
18790       MsgBox "There are no accounts scheduled for fees in the month of " & .FeeMonth.Column(1) & ".", _
              vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
18800     Else
18810       DoCmd.Hourglass False
18820       If MsgBox("Do you want to create journal entries with a posting date of " & CStr(.postdate.Value) & "?" & _
                vbCrLf & vbCrLf & "If you want a different date, click 'No' and change posting date to desired date.", _
                vbQuestion + vbYesNo, ("Create New Entries" & Space(40))) = vbYes Then

18830         DoCmd.Hourglass True
18840         DoEvents

18850         rst1.MoveLast
18860         lngRecs = rst1.RecordCount
18870         rst1.MoveFirst

18880         DoCmd.Hourglass False
18890         MsgBox CStr(lngRecs + 1&) & " journal entries will be created." & vbCrLf & vbCrLf & _
                "  Paid: " & CStr(lngRecs) & vbCrLf & _
                "  Received: 1", vbInformation + vbOKOnly, "New Journal Entries"

18900         DoCmd.Hourglass True
18910         DoEvents

18920         lngRevCodeID_Inc = Nz(.cmbRevenueCodes_Received, REVID_INC)
18930         lngRevCodeID_Exp = Nz(.cmbRevenueCodes_Paid, REVID_EXP)

18940         lngTaxCodeID_Inc = Nz(.cmbTaxCodes_Received, TAXID_INC)
18950         lngTaxCodeID_Ded = Nz(.cmbTaxCodes_Paid, TAXID_DED)

18960         For lngX = 1& To lngRecs
18970           rst2.AddNew
18980           rst2![journaltype] = "Paid"
18990           rst2![accountno] = rst1![accountno]
19000           rst2![RecurringItem] = "Fee for Period"
19010           If (rst1![ScheduleDetail_Total] + rst1![Schedule_Base]) / rst1![Frequency] > rst1![Schedule_Minimum] / rst1![Frequency] Then
19020             varValue = (rst1![ScheduleDetail_Total] + rst1![Schedule_Base]) / rst1![Frequency]
19030           Else
19040             varValue = rst1![Schedule_Minimum] / rst1![Frequency]
19050           End If
19060           Select Case .opgBasis
                Case .opgBasis_optIncome.OptionValue
19070             rst2![PCash] = 0
19080             rst2![ICash] = Format(varValue * -1, "Currency")
19090           Case .opgBasis_opt5050.OptionValue
19100             rst2![PCash] = Format(varValue * -0.5, "Currency")
19110             rst2![ICash] = Format(varValue * -1, "Currency") - rst2![PCash]
19120           End Select
19130           rst2![transdate] = .postdate
19140           rst2![revcode_ID] = lngRevCodeID_Exp
19150           rst2![taxcode] = lngTaxCodeID_Ded
19160           rst2![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
19170           rst2.Update
19180           varTotal = varTotal + varValue
19190           If lngX < lngRecs Then rst1.MoveNext
19200         Next

19210         rst2.AddNew
19220         rst2![journaltype] = "Received"
19230         If varSuspense = "No" Then
19240           rst2![accountno] = "99-SUSPENSE"
19250         Else
19260           rst2![accountno] = "SUSPENSE"
19270         End If
19280         rst2![RecurringItem] = "Fee(s) for Period"
19290         rst2![transdate] = .postdate
19300         rst2![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
19310         rst2![PCash] = 0
19320         rst2![ICash] = Format(varTotal, "Currency")
19330         rst2![revcode_ID] = lngRevCodeID_Inc
19340         rst2![taxcode] = lngTaxCodeID_Inc
19350         rst2![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
19360         rst2.Update

19370         rst1.Close
19380         rst2.Close
19390         dbs.Close

19400         DoCmd.Hourglass False
19410         MsgBox "Journal Entries Created", vbInformation + vbOKOnly, ("Journal Entries Created" & Space(40))

19420       Else
19430         blnContinue = False
19440       End If  ' ** MsgBox.
19450     End If  ' ** EOF/BOF.

19460     If blnContinue = True Then
19470       .cmdClose.Enabled = False
19480       DoEvents
19490       DoCmd.Close acForm, THIS_NAME
19500     Else
            ' ** Else what?
19510     End If  ' ** blnContinue.

19520   End With

19530   DoCmd.Hourglass False

EXITP:
19540   Set rst1 = Nothing
19550   Set rst2 = Nothing
19560   Set qdf = Nothing
19570   Set dbs = Nothing
19580   Exit Sub

ERRH:
19590   DoCmd.Hourglass False
19600   Select Case ERR.Number
        Case Else
19610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19620   End Select
19630   Resume EXITP

End Sub

Private Sub cmdCreateEntries_KeyDown(KeyCode As Integer, Shift As Integer)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCreateEntries_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19720   intShiftDown = (Shift And acShiftMask) > 0
19730   intAltDown = (Shift And acAltMask) > 0
19740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19760     Select Case intRetVal
          Case vbKeyTab
19770       With Me
19780         intRetVal = 0
19790         .cmdClose.SetFocus
19800       End With
19810     End Select
19820   End If

        ' ** Shift keys.
19830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19840     Select Case intRetVal
          Case vbKeyTab
19850       With Me
19860         intRetVal = 0
19870         If .cmbTaxCodes_Received.Visible = True And .cmbTaxCodes_Received.Enabled = True Then
19880           .cmbTaxCodes_Received.SetFocus
19890         ElseIf .cmbRevenueCodes_Received.Visible = True And .cmbRevenueCodes_Received.Enabled = True Then
19900           .cmbRevenueCodes_Received.SetFocus
19910         Else
19920           .cmdPrint.SetFocus
19930         End If
19940       End With
19950     End Select
19960   End If

EXITP:
19970   KeyCode = intRetVal
19980   Exit Sub

ERRH:
19990   intRetVal = 0
20000   THAT_PROC = THIS_PROC
20010   That_Erl = Erl
20020   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20030   Resume EXITP

End Sub

Public Sub CalculateFees()
' ** Fee Frequency:
' **   1 = Monthly
' **   2 = Quarterly
' **   3 = Semiannually
' **   4 = Annually

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "CalculateFees"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim lngFeeFreq As Long
        Dim dblRemainder As Double
        Dim dblRate As Double
        Dim lngRecs1 As Long, lngRecs2 As Long
        Dim lngX As Long, lngY As Long
        Dim blnContinue As Boolean

20110   blnContinue = True

20120   With Me
20130     Select Case .FeeMonth.Column(0)
          Case 1, 2, 4, 5, 7, 8, 10, 11
            ' ** FeeFreq 1: Jan, Feb, Apr, May, Jul, Aug, Oct, Nov.
            ' ** These months only print if the Fee Freqency is 1 (monthly).
20140       lngFeeFreq = 1&
20150     Case 3, 9
            ' ** FeeFreq 1,2:  Mar, Sep.
            ' ** If the chosen month is March or September, only 1 (monthly), 2 (quarterly) will print.
20160       lngFeeFreq = 2&
20170     Case 6
            ' ** FeeFreq 1,2,3: Jun.
            ' ** If the chosen month is June, only 1 (monthly), 2 (quarterly), 3 (semiannually) will print.
20180       lngFeeFreq = 3&
20190     Case 12
            ' ** FeeFreq 1,2,3,4: Dec.
            ' ** If the chosen month is December, all acounts with a Fee Frequency will print,
            ' ** because December is part of 1 (monthly), 2 (quarterly), 3 (semiannually), and 4 (annually).
20200       lngFeeFreq = 4&
20210     End Select
20220   End With

        'ScheduleDetail needs to be entered in the proper order. There is no other sorting available!

20230   Set dbs = CurrentDb
20240   With dbs

          ' ** Empty FeeCalculations table.
20250     Set qdf1 = .QueryDefs("qryFeeCalculations_08")
20260     qdf1.Execute
20270     Set qdf1 = Nothing
20280     DoEvents

          ' ***********************************************************************
          ' ** All market values are based on the rates currently in tblCurrency.
          ' ***********************************************************************

          ' ** qryFeeCalculations_04b (Account, linked to qryFeeCalculations_03b (Account, linked
          ' ** to qryFeeCalculations_02b (qryFeeCalculations_01b (Account, linked to ActiveAssets,
          ' ** all accounts, with TotalCost_usd), grouped and summed), Schedule, ScheduleDetail),
          ' ** by specified [feefreq]), grouped by accountno.
20290     Set qdf1 = dbs.QueryDefs("qryFeeCalculations_09")  ' ** Criteria is: <=[feefreq].
20300     With qdf1.Parameters
20310       ![FeeFreq] = lngFeeFreq
20320     End With
20330     Set rst1 = qdf1.OpenRecordset
20340     With rst1
20350       If .BOF = True And .EOF = True Then
20360         blnContinue = False
20370       Else
              ' ** This returns just 1 field, accountno, for all
              ' ** accounts with scheduled fees for the chosen month.
20380         .MoveLast
20390         lngRecs1 = .RecordCount
20400         .MoveFirst
20410       End If
20420     End With  ' ** rst1.

20430     If blnContinue = True Then
20440       For lngX = 1& To lngRecs1

              ' ** qryFeeCalculations_04 (Account, linked to qryFeeCalculations_03 (Account, linked to
              ' ** qryFeeCalculations_02 (qryFeeCalculations_01 (Account, linked to ActiveAssets, by specified [actno]),
              ' ** grouped and summed), Schedule, ScheduleDetail, by specified [actno]), by specified [feefreq]),
              ' ** with Schedule_Base, Schedule_Minimum, ScheduleDetail_Rate, ScheduleDetail_Amount.
20450         Set qdf2 = .QueryDefs("qryFeeCalculations_05")
20460         With qdf2.Parameters
20470           ![actno] = rst1![accountno]
20480           ![FeeFreq] = lngFeeFreq
20490         End With
20500         Set rst2 = qdf2.OpenRecordset
20510         With rst2
20520           If .BOF = True And .EOF = True Then
20530             lngRecs2 = 0&
20540             dblRemainder = 0#
20550           Else
                  ' ** This returns the full set of fee details, with the various percentages,
                  ' ** for this single account, in ascending scheddets_order order.
20560             .MoveLast
20570             lngRecs2 = .RecordCount
20580             .MoveFirst
20590             Select Case blnHasForEx
                  Case True
20600               dblRemainder = Nz(![CurrentTotalMarketValue_usd], 0)
20610             Case False
20620               dblRemainder = Nz(![CurrentTotalMarketValue], 0)  ' ** All records have the same total.
20630             End Select
20640           End If
20650         End With  ' ** rst2.

              ' ** Store the initial MarketValue in the dblRemainder.
              ' ** Then subtract what is used from this number to get the remainder
              ' ** that will be used in the next rate calculation.

20660         If dblRemainder <> 0# Then
20670           For lngY = 1& To lngRecs2

20680             If lngY = lngRecs2 Then
                    ' ** This is the last record. So if the dblRemainder is
                    ' ** still greater than zero, muliply it by the last rate.

20690               Select Case IsNull(rst2![ScheduleDetail_Rate])
                    Case True
20700                 dblRate = 0#
20710               Case False
20720                 dblRate = rst2![ScheduleDetail_Rate]
20730               End Select

20740               If dblRemainder > 0# Then

                      'IF IT GETS TO THIS LAST RECORD BEFORE EXHAUSTING dblRemainder, APPLY THIS LAST RATE TO dblRemainder.
                      'THIS MEANS THAT THE LAST RATE IS APPLIED WHETHER dblRemainder IS LESS THAN OR GREATER THAN THE LAST AMOUNT.

                      ' ** Append new record to FeeCalculations, by specified [actno], [rat], [amt], [tot].
20750                 Set qdf1 = .QueryDefs("qryFeeCalculations_10")
20760                 With qdf1.Parameters
20770                   ![actno] = rst2![accountno]
20780                   ![rat] = dblRate
20790                   ![amt] = dblRemainder  ' ** This is USD.
20800                   ![tot] = (dblRemainder * dblRate)
20810                 End With
20820                 qdf1.Execute
20830                 Set qdf1 = Nothing
20840                 dblRemainder = 0#

20850               End If

20860             Else

20870               If dblRemainder >= rst2![ScheduleDetail_Amount] Then

                      'IT BEGINS HERE.
                      'WITH dblRemainder CONTAINING THE FULL AMOUNT, IT APPLIES THE 1ST RATE TO THE 1ST AMOUHNT,
                      'THEN SUBTRACTS THAT 1ST AMOUNT FROM dblRemainder.
                      'THE RATE RECORD THEN MOVES ON.
                      'IT WILL REPEAT THIS SECTION AS LONG AS dblRemainder IS GREATER THAN THE AMOUNT.

                      ' ** Append new record to FeeCalculations, by specified [actno], [rat], [amt], [tot].
20880                 Set qdf1 = .QueryDefs("qryFeeCalculations_10")
20890                 With qdf1.Parameters
20900                   ![actno] = rst2![accountno]
20910                   ![rat] = rst2![ScheduleDetail_Rate]
20920                   ![amt] = rst2![ScheduleDetail_Amount]
20930                   ![tot] = (rst2![ScheduleDetail_Amount] * rst2![ScheduleDetail_Rate])
20940                 End With
20950                 qdf1.Execute
20960                 Set qdf1 = Nothing
20970                 dblRemainder = (dblRemainder - rst2![ScheduleDetail_Amount])

20980               Else

                      'WHEN dblRemainder IS LESS THAN THE LISTED AMOUNT, IT APPLIES THIS RATE TO dblRemainder.
                      'THIS WILL BE THE FINAL RATE IF THERE ARE STILL MORE RECORDS, BUT dblRemainder IS LESS THAN THIS AMOUNT.
                      'IT MAY GET TO THE LAST RECORD WITHOUT EVER REACHING THIS PART OF THE CODE.

                      ' ** Append new record to FeeCalculations, by specified [actno], [rat], [amt], [tot].
20990                 Set qdf1 = .QueryDefs("qryFeeCalculations_10")
21000                 With qdf1.Parameters
21010                   ![actno] = rst2![accountno]
21020                   ![rat] = rst2![ScheduleDetail_Rate]
21030                   ![amt] = dblRemainder
21040                   ![tot] = (dblRemainder * rst2![ScheduleDetail_Rate])
21050                 End With
21060                 qdf1.Execute
21070                 Set qdf1 = Nothing
21080                 dblRemainder = 0#

21090               End If

21100             End If

21110             If dblRemainder = 0# Then
21120               Exit For
21130             Else
21140               If lngY < lngRecs2 Then rst2.MoveNext
21150             End If

21160           Next  ' ** rst2: lngY.
21170         End If  ' ** dblRemainder.

21180         rst2.Close
21190         Set rst2 = Nothing
21200         Set qdf2 = Nothing
21210         If lngX < lngRecs1 Then rst1.MoveNext

21220       Next  ' ** rst1: lngX.
21230       rst1.Close
21240       Set rst1 = Nothing
21250       Set qdf1 = Nothing

21260     Else
21270       rst1.Close
21280       Set rst1 = Nothing
21290       Set qdf1 = Nothing
21300     End If  ' ** blnContinue.

21310     If gstrFormQuerySpec = vbNullString Then
21320       gstrFormQuerySpec = THIS_NAME
21330     End If

21340     .Close
21350   End With  ' ** dbs.

EXITP:
21360   Set rst2 = Nothing
21370   Set rst1 = Nothing
21380   Set qdf2 = Nothing
21390   Set qdf1 = Nothing
21400   Set dbs = Nothing
21410   Exit Sub

ERRH:
21420   Select Case ERR.Number
        Case Else
21430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21440   End Select
21450   Resume EXITP

End Sub

Private Sub HideIncExp_Rec()

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "HideIncExp_Rec"

        Dim lngOffset1 As Long, lngOffset2 As Long
        Dim lngTmp01 As Long, lngTmp02 As Long

21510   With Me

21520     .cmbRevenueCodes_Received_box.Visible = False
21530     .cmbRevenueCodes_Received_box2.Visible = False
21540     .cmbRevenueCodes_Received_box_lbl.Visible = False
21550     .cmbRevenueCodes_Received_box_lbl_dim_hi.Visible = False
21560     .cmbRevenueCodes_Received_hline01.Visible = False
21570     .cmbRevenueCodes_Received_hline02.Visible = False
21580     .cmbRevenueCodes_Received_hline03.Visible = False
21590     .cmbRevenueCodes_Received_vline01.Visible = False
21600     .cmbRevenueCodes_Received_vline02.Visible = False
21610     .cmbRevenueCodes_Received_vline03.Visible = False
21620     .cmbRevenueCodes_Received_vline04.Visible = False
21630     .cmbRevenueCodes_Received.Visible = False
21640     .cmbRevenueCodes_Received_display.Visible = False
21650     .cmbRevenueCodes_Received_lbl.Visible = False
21660     .cmbRevenueCodes_Received_lbl2.Visible = False
21670     .cmbTaxCodes_Received.Visible = False
21680     .cmbTaxCodes_Received_display.Visible = False
21690     .cmbTaxCodes_Received_lbl.Visible = False
21700     .cmbTaxCodes_Received_lbl2.Visible = False

21710     lngOffset1 = ((.cmdPreviewPrint_box.Top + .cmdPreviewPrint_box.Height) - .cmbRevenueCodes_Received_box.Height)
21720     lngTmp01 = (.Detail.Height - (.cmbRevenueCodes_Received_box.Top + .cmbRevenueCodes_Received_box.Height))
21730     lngTmp01 = ((.cmbRevenueCodes_Paid_box.Top + .cmbRevenueCodes_Paid_box.Height) + lngTmp01)
21740     lngOffset2 = (.Detail.Height - lngTmp01)

21750     .cmbRevenueCodes_Received_box.Top = .cmbRevenueCodes_Received_box.Top - lngOffset1
21760     .cmbRevenueCodes_Received_box2.Top = .cmbRevenueCodes_Received_box2.Top - lngOffset1
21770     .cmbRevenueCodes_Received_box_lbl.Top = .cmbRevenueCodes_Received_box_lbl.Top - lngOffset1
21780     .cmbRevenueCodes_Received_box_lbl_dim_hi.Top = .cmbRevenueCodes_Received_box_lbl_dim_hi.Top - lngOffset1
21790     .cmbRevenueCodes_Received_hline01.Top = .cmbRevenueCodes_Received_hline01.Top - lngOffset1
21800     .cmbRevenueCodes_Received_hline02.Top = .cmbRevenueCodes_Received_hline02.Top - lngOffset1
21810     .cmbRevenueCodes_Received_hline03.Top = .cmbRevenueCodes_Received_hline03.Top - lngOffset1
21820     .cmbRevenueCodes_Received_vline01.Top = .cmbRevenueCodes_Received_vline01.Top - lngOffset1
21830     .cmbRevenueCodes_Received_vline02.Top = .cmbRevenueCodes_Received_vline02.Top - lngOffset1
21840     .cmbRevenueCodes_Received_vline03.Top = .cmbRevenueCodes_Received_vline03.Top - lngOffset1
21850     .cmbRevenueCodes_Received_vline04.Top = .cmbRevenueCodes_Received_vline04.Top - lngOffset1
21860     .cmbRevenueCodes_Received.Top = .cmbRevenueCodes_Received.Top - lngOffset1
21870     .cmbRevenueCodes_Received_display.Top = .cmbRevenueCodes_Received_display.Top - lngOffset1
21880     .cmbRevenueCodes_Received_lbl.Top = .cmbRevenueCodes_Received_lbl.Top - lngOffset1
21890     .cmbRevenueCodes_Received_lbl2.Top = .cmbRevenueCodes_Received_lbl2.Top - lngOffset1
21900     .cmbTaxCodes_Received.Top = .cmbTaxCodes_Received.Top - lngOffset1
21910     .cmbTaxCodes_Received_display.Top = .cmbTaxCodes_Received_display.Top - lngOffset1
21920     .cmbTaxCodes_Received_lbl.Top = .cmbTaxCodes_Received_lbl.Top - lngOffset1
21930     .cmbTaxCodes_Received_lbl2.Top = .cmbTaxCodes_Received_lbl2.Top - lngOffset1

21940     .detail_height_line.Top = .detail_height_line.Top - lngOffset2
21950     .Detail.Height = .Detail.Height - lngOffset2

21960     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
21970     lngMonitorNum = 1&: lngTmp02 = 0&
21980     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
21990     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
22000     If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
22010     DoCmd.MoveSize lngFrm_Left, lngTmp02, lngFrm_Width, (lngFrm_Height - lngOffset2)  'lngFrm_Top
22020     If lngMonitorNum > 1& Then
22030       LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
22040     End If

22050     blnHideIncExp_Rec = True
22060     DoEvents

22070   End With

EXITP:
22080   Exit Sub

ERRH:
22090   Select Case ERR.Number
        Case Else
22100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22110   End Select
22120   Resume EXITP

End Sub

Private Sub HideIncExp_Paid()

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "HideIncExp_Paid"

        Dim lngOffset1 As Long, lngOffset2 As Long
        Dim lngTmp01 As Long, lngTmp02 As Long

22210   With Me

22220     GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

22230     .cmbRevenueCodes_Paid_box.Visible = False
22240     .cmbRevenueCodes_Paid_box2.Visible = False
22250     .cmbRevenueCodes_Paid_box_lbl.Visible = False
22260     .cmbRevenueCodes_Paid_box_lbl_dim_hi.Visible = False
22270     .cmbRevenueCodes_Paid_hline01.Visible = False
22280     .cmbRevenueCodes_Paid_hline02.Visible = False
22290     .cmbRevenueCodes_Paid_hline03.Visible = False
22300     .cmbRevenueCodes_Paid_vline01.Visible = False
22310     .cmbRevenueCodes_Paid_vline02.Visible = False
22320     .cmbRevenueCodes_Paid_vline03.Visible = False
22330     .cmbRevenueCodes_Paid_vline04.Visible = False
22340     .cmbRevenueCodes_Paid.Visible = False
22350     .cmbRevenueCodes_Paid_display.Visible = False
22360     .cmbRevenueCodes_Paid_lbl.Visible = False
22370     .cmbRevenueCodes_Paid_lbl2.Visible = False
22380     .cmbTaxCodes_Paid.Visible = False
22390     .cmbTaxCodes_Paid_display.Visible = False
22400     .cmbTaxCodes_Paid_lbl.Visible = False
22410     .cmbTaxCodes_Paid_lbl2.Visible = False

22420     lngOffset1 = ((.cmdPreviewPrint_box.Top + .cmdPreviewPrint_box.Height) - .cmbRevenueCodes_Paid_box.Height)
22430     lngTmp01 = (.Detail.Height - (.cmbRevenueCodes_Paid_box.Top + .cmbRevenueCodes_Paid_box.Height))
22440     lngTmp01 = ((.cmdPreviewPrint_box.Top + .cmdPreviewPrint_box.Height) + lngTmp01)
22450     lngOffset2 = (.Detail.Height - lngTmp01)

22460     .cmbRevenueCodes_Paid_box.Top = .cmbRevenueCodes_Paid_box.Top - lngOffset1
22470     .cmbRevenueCodes_Paid_box2.Top = .cmbRevenueCodes_Paid_box2.Top - lngOffset1
22480     .cmbRevenueCodes_Paid_box_lbl.Top = .cmbRevenueCodes_Paid_box_lbl.Top - lngOffset1
22490     .cmbRevenueCodes_Paid_box_lbl_dim_hi.Top = .cmbRevenueCodes_Paid_box_lbl_dim_hi.Top - lngOffset1
22500     .cmbRevenueCodes_Paid_hline01.Top = .cmbRevenueCodes_Paid_hline01.Top - lngOffset1
22510     .cmbRevenueCodes_Paid_hline02.Top = .cmbRevenueCodes_Paid_hline02.Top - lngOffset1
22520     .cmbRevenueCodes_Paid_hline03.Top = .cmbRevenueCodes_Paid_hline03.Top - lngOffset1
22530     .cmbRevenueCodes_Paid_vline01.Top = .cmbRevenueCodes_Paid_vline01.Top - lngOffset1
22540     .cmbRevenueCodes_Paid_vline02.Top = .cmbRevenueCodes_Paid_vline02.Top - lngOffset1
22550     .cmbRevenueCodes_Paid_vline03.Top = .cmbRevenueCodes_Paid_vline03.Top - lngOffset1
22560     .cmbRevenueCodes_Paid_vline04.Top = .cmbRevenueCodes_Paid_vline04.Top - lngOffset1
22570     .cmbRevenueCodes_Paid.Top = .cmbRevenueCodes_Paid.Top - lngOffset1
22580     .cmbRevenueCodes_Paid_display.Top = .cmbRevenueCodes_Paid_display.Top - lngOffset1
22590     .cmbRevenueCodes_Paid_lbl.Top = .cmbRevenueCodes_Paid_lbl.Top - lngOffset1
22600     .cmbRevenueCodes_Paid_lbl2.Top = .cmbRevenueCodes_Paid_lbl2.Top - lngOffset1
22610     .cmbTaxCodes_Paid.Top = .cmbTaxCodes_Paid.Top - lngOffset1
22620     .cmbTaxCodes_Paid_display.Top = .cmbTaxCodes_Paid_display.Top - lngOffset1
22630     .cmbTaxCodes_Paid_lbl.Top = .cmbTaxCodes_Paid_lbl.Top - lngOffset1
22640     .cmbTaxCodes_Paid_lbl2.Top = .cmbTaxCodes_Paid_lbl2.Top - lngOffset1

22650     .detail_height_line.Top = .detail_height_line.Top - lngOffset2
22660     .Detail.Height = .Detail.Height - lngOffset2

22670     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
22680     lngMonitorNum = 1&: lngTmp02 = 0&
22690     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
22700     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
22710     If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
22720     DoCmd.MoveSize lngFrm_Left, lngTmp02, lngFrm_Width, (lngFrm_Height - lngOffset2)  'lngFrm_Top
22730     If lngMonitorNum > 1& Then
22740       LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
22750     End If

22760     blnHideIncExp_Paid = True
22770     DoEvents

22780   End With

EXITP:
22790   Exit Sub

ERRH:
22800   Select Case ERR.Number
        Case Else
22810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22820   End Select
22830   Resume EXITP

End Sub
