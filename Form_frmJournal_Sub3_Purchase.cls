VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournal_Sub3_Purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournal_Sub3_Purchase"

'VGC 10/01/2017: CHANGES!

' #########################
' ## Use VBA_RenumErrh().  39460
' #########################

' ** Shortcut Alt keys responsive from this form:
' **   Map:              A {cmdPurchaseMap}
' **   Close:            C {cmdPurchaseClose}
' **   Dividend:         D {opgJournal_optDividend on frmJournal}
' **   Cancel:           E {cmdPurchaseCancel}
' **   Interest:         I {opgJournal_optInterest on frmJournal}
' **   New Loc:          L {cmdPurchaseLocNew}
' **   Miscellaneous:    M {opgJournal_optMisc on frmJournal}
' **   Inc/Exp Codes:    N {cmbRevenueCodes}
' **   OK:               O {cmdPurchaseOK}
' **   Purchase:         P {opgJournal_optPurchase on frmJournal}
' **   Sale:             S {opgJournal_optSale on frmJournal}
' **   Tax Codes:        T {cmbTaxCodes}
' **   New Asset:        W {cmdPurchaseAssetNew}
' **   Exit:             X {cmdPurchaseClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Post For Trade:   D {purchaseAssetDate}
' **   Next Record:      N {MoveRec}
' **   Previous Record:  P {MoveRec}
' **   Switch:           T {cmdSwitch on frmJournal}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Mark Sale:        S {tglPurchaseSale}
' **   Close Timer:      T {Form_KeyDown}

' ** Shortcut F-keys responsive from this form:
' **   Refresh:          F5 {Recalc}
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Purchase, Deposit are always INCOME for Tax Codes.
' ** Liability is always EXPENSE for Tax Codes.

' ** NOTE: 'HA-' prefix to asset name means 'Hidden Asset'.

' **********************************************************************
' ** NOTE: Enter key only invokes command button on OK and Cancel.
' **       Though Enter would usually invoke a command button,
' **       users intuitively hit Enter to quickly move through fields.
' **********************************************************************

' *********************************************
' ** Image suffixes:
' **   std  Windows Standard Theme (Classic).
' **   des  My Dessert Theme, Beige.
' *********************************************

' *************************
' ** Date defaults:
' **   transdate = Date()
' **   assetdate = Now()
' *************************

' *********************************************************************
' *********************************************************************
' ** 1. Keystroke Event Order, Sequence:
' **      KeyDown -> KeyPress -> Change -> KeyUp

' ** 2. Keystroke Event Order, Sequence:
' **    First Control:
' **      KeyDown -> BeforeUpdate -> AfterUpdate -> Exit -> LostFocus
' **    Second Control:
' **      Enter -> GotFocus -> KeyPress -> KeyUp

' ** 3. Mouse Event Order, Sequence:
' **      MouseDown -> MouseUp -> Click

' ** 4. Mouse Event Order, Sequence:
' **    First Control:
' **      Exit -> LostFocus
' **    Second Control:
' **      Enter -> GotFocus -> MouseDown -> MouseUp -> Click

' ** 5. Keystroke Event Order, Sequence Example:
' **      Current (for the new record)
' **      ->
' **      Enter (for the first control in the new record)
' **      ->
' **      GotFocus (for the control)
' **      ->
' **      KeyDown (for the control)
' **      ->
' **      KeyPress (for the control)
' **      ->
' **      BeforeInsert (for the new record in the form)
' **      ->
' **      Change (for the control if it's a text box or combo box)
' **      ->
' **      KeyUp (for the control)
' *********************************************************************
' *********************************************************************

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Combo box column constants: purchaseAssetNo.
'Private Const CBX_A_ASTNO  As Integer = 0  'assetno
'Private Const CBX_A_DESC   As Integer = 1  'totdesc
'Private Const CBX_A_CUSIP  As Integer = 2  'cusip
Private Const CBX_A_TYPE   As Integer = 3  'assettype
Private Const CBX_A_TAX    As Integer = 4  'taxcode
Private Const CBX_A_D4D    As Integer = 5  'AssetType_D4D
Private Const CBX_A_CURRID As Integer = 6  'curr_id

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

' ** Combo box column constants: purchaseCurr_ID.
Private Const CBX_C_CURRID As Integer = 0  'curr_id
'Private Const CBX_C_CODE   As Integer = 1  'curr_code
'Private Const CBX_C_NAME   As Integer = 2  'curr_name
Private Const CBX_C_SYM    As Integer = 3  'currsym_symbol
Private Const CBX_C_DEC    As Integer = 4  'curr_decimal
'Private Const CBX_C_RATE1  As Integer = 5  'curr_rate1
Private Const CBX_C_RATE2  As Integer = 6  'curr_rate2
Private Const CBX_C_DATE   As Integer = 7  'curr_date

' ** Array: arr_varNoChar().
Private lngNoChars As Long, arr_varNoChar As Variant
Private Const N_CID   As Integer = 0
'Private Const N_CNAM  As Integer = 1
'Private Const N_FID   As Integer = 2
'Private Const N_FNAM  As Integer = 3
'Private Const N_NOCHR As Integer = 4

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean
Private blnPurchaseSale_Focus As Boolean, blnPurchaseSale_MouseDown As Boolean

Private lngCurrID As Long
Private blnGoToSaleForm As Boolean, blnCmdButton As Boolean, blnMapNewRec As Boolean
Private blnAccountNoErr As Boolean, blnStartTrans As Boolean, blnDefTypeAssigned As Boolean
Private datAssetDate_OldValue As Date, lngAssetDate_OldValue As Long, blnAssetDateChecked As Boolean
Private blnAccountNoHasFocus As Boolean, blnTypeIsNull As Boolean, blnSpecialCap As Boolean, intSpecialCapOpt As Integer
Private datPostingDate As Date, lngDateCnt As Long, lngErrCnt As Long
Private blnPCash_Exit As Boolean, blnPCash_LostFocus As Boolean, blnGoneToReport As Boolean
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmJournal" Then
160         If IsLoaded(strCallingForm) = True Then  ' ** Module Function: modFileUtilities.

170           lngDateCnt = 0&: lngErrCnt = 0&
180           blnPCash_Exit = False: blnPCash_LostFocus = False
190           blnAccountNoErr = False: blnStartTrans = False: blnDefTypeAssigned = False
200           gblnSwitchTo = False: blnMapNewRec = False
210           gblnIsLiability = False: gstrPurchaseType = vbNullString

220           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
230           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False
240           blnPurchaseSale_Focus = False: blnPurchaseSale_MouseDown = False

250           .cmdLock.Enabled = False
260           .cmdLock_open_raised_img.Visible = False
270           .cmdLock_open_raised_img_dis.Visible = False
280           .cmdLock_closed_raised_img.Visible = False

290         Else
300           Cancel = -1
310         End If
320       Else
330         Cancel = -1
340       End If
350     End With

360     If Cancel = -1 Then
370       Beep
380       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
390     End If

EXITP:
400     Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
440     Resume EXITP

End Sub

Private Sub Form_Load()
' ** Create an instance of the Calendar class.
' ** Copyright Lebans Holdings 1999 Ltd.

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me

          ' ** Create an instance of the Calendar class.
520       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
530       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
540       clsMonthClass.PositionAtCursor = True

550       blnAccountNoHasFocus = True

560       datPostingDate = GetPostDate  ' ** Module Function: modUtilities.

          ' ** Do this whether or not Currency is visible.
570       CLR_DISABLED_FG = CLR_DKGRY
580       CLR_DISABLED_BG = CLR_LTTEAL

590       .purchaseCurr_ID.Enabled = False
600       .purchaseCurr_ID.Locked = True
610       .purchaseCurr_ID.ForeColor = CLR_DISABLED_FG
620       .purchaseCurr_ID.BackColor = CLR_VLTGRN 'CLR_DISABLED_BG
630       purchaseCurr_ID_AfterUpdate  ' ** Procedure: Below.

          ' ** Set the default Posting Date (transdate), entered by the user on opening.
640       .purchaseTransDate.DefaultValue = "#" & CDate(Format(datPostingDate, "mm/dd/yyyy")) & "#"

650       gstrFormQuerySpec = .Parent.Name  ' ** Make sure this is set for the assetno combo box.

660       gblnPurchaseValidated = False
670       gblnPurchaseChanged = False

          ' ** Make sure the SpecialCap options are set.
680       If gstrCo_Name = vbNullString Then
690         CoOptions_Read  ' ** Module Function: modStartupFunctions.
700       End If

          ' ** Load SpecialCapGainLoss setting.
710       blnSpecialCap = gblnSpecialCapGainLoss
720       intSpecialCapOpt = gintSpecialCapGainLossOpt

730       .SpecialCapGainLoss = blnSpecialCap

740       .NavigationButtons = True

750       With .Parent
760         .opgJournal.Enabled = True
770         .cmdSwitch.Enabled = True
780         .cmdSwitch_raised_img.Visible = True
790         .cmdSwitch_raised_img_dis.Visible = False
800         .cmdSwitch_raised_semifocus_dots_img.Visible = False
810         .cmdSwitch_raised_focus_img.Visible = False
820         .cmdSwitch_raised_focus_dots_img.Visible = False
830         .cmdSwitch_sunken_focus_dots_img.Visible = False
840       End With
850       .cmdPurchaseClose.Enabled = True
860       .cmdPurchaseOK.Enabled = False
870       .cmdPurchaseCancel.Enabled = False
880       .Parent.NavVis True  ' ** Form Procedure: frmJournal.

890       If .NewRecord = True Then
900         .tglPurchaseSale.Enabled = False
910         .tglPurchaseSale_false_raised_img_dis.Visible = True
920         .tglPurchaseSale_false_raised_img.Visible = False
930         .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = False
940         .tglPurchaseSale_false_raised_focus_img.Visible = False
950         .tglPurchaseSale_false_raised_focus_dots_img.Visible = False
960         .tglPurchaseSale_false_sunken_focus_dots_img.Visible = False
970         .tglPurchaseSale_true_raised_img.Visible = False
980         .tglPurchaseSale_true_raised_focus_img.Visible = False
990         .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
1000        .tglPurchaseSale_true_sunken_focus_dots_img.Visible = False
1010        .tglPurchaseSale_true_raised_img_dis.Visible = False
1020      End If

1030      .cmdPurchaseAssetNew.Enabled = False
1040      .cmdPurchaseLocNew.Enabled = False
1050      .AllowAdditions = True

1060      Select Case gblnIncomeTaxCoding
          Case True
1070        .cmbTaxCodes.Visible = True
1080        .cmbTaxCodes_lbl2.Visible = False
1090      Case False
1100        .cmbTaxCodes.Visible = False
1110        .cmbTaxCodes_lbl2.Visible = True
1120      End Select

1130      Select Case gblnRevenueExpenseTracking
          Case True
1140        .cmbRevenueCodes.Visible = True
1150        .cmbRevenueCodes_lbl2.Visible = False
1160      Case False
1170        .cmbRevenueCodes.Visible = False
1180        .cmbRevenueCodes_lbl2.Visible = True
1190      End Select

1200    End With

EXITP:
1210    Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
440     Resume EXITP

End Sub

Private Sub Form_Current()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

1310    FormCurrent_Sub3 blnCmdButton, blnGoToSaleForm, blnDefTypeAssigned, blnStartTrans, blnTypeIsNull, _
          blnAccountNoErr, blnAssetDateChecked, datAssetDate_OldValue, lngAssetDate_OldValue, _
          blnSpecialCap, intSpecialCapOpt, lngNoChars, arr_varNoChar, lngCurrID, lngErrCnt, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
1320    Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
440     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim Cancel As Integer
        Dim intRetVal As Integer

1410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1420    intShiftDown = (Shift And acShiftMask) > 0
1430    intAltDown = (Shift And acAltMask) > 0
1440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Post For Trade:   D {purchaseAssetDate}
        ' **   Next Record:      N {MoveRec}
        ' **   Previous Record:  P {MoveRec}
        ' **   Switch:           T {cmdSwitch on frmJournal}

        ' ** Shortcut Ctrl-Shift to other forms and functionality:
        ' **   Mark Sale:        S {tglMiscSale}
        ' **   Close Timer:      T {Form_KeyDown}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyF5
1470        With Me
1480          intRetVal = 0
1490          .purchaseType.Requery
1500          DoEvents
1510          .purchaseAssetNo.Requery
1520          DoEvents
1530        End With
1540      Case vbKeyF7
1550        intRetVal = 0
1560        cmdCalendar1_Click  ' ** Procedure: Below.
1570      Case vbKeyF8
1580        intRetVal = 0
1590        cmdCalendar2_Click  ' ** Procedure: Below.
1600      Case vbKeyEscape
1610        intRetVal = 0
1620        If Forms("frmJournal").opgJournal.Enabled = True Then
1630          Select Case gblnPurchaseChanged
              Case True
1640            cmdPurchaseCancel_Click  ' ** Procedure: Below.
1650          Case False
1660            cmdPurchaseClose_Click  ' ** Procedure: Below.
1670          End Select
1680        Else
1690          cmdPurchaseCancel_Click  ' ** Procedure: Below.
1700        End If
1710      Case vbKeyDelete
1720        gblnPurchaseChanged = True
1730      End Select
1740    End If

        ' ** Alt keys.
1750    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1760      Select Case intRetVal
          Case vbKeyC
1770        With Me
1780          intRetVal = 0
1790          If .cmdPurchaseClose.Enabled = True Then
1800            .cmdPurchaseClose.SetFocus
1810            cmdPurchaseClose_Click  ' ** Procedure: Below.
1820          Else
1830            Beep
1840          End If
1850        End With
1860      Case vbKeyE
1870        With Me
1880          intRetVal = 0
1890          If .cmdPurchaseCancel.Enabled = True Then
1900            .cmdPurchaseCancel.SetFocus
1910            cmdPurchaseCancel_Click  ' ** Procedure: Below.
1920          Else
1930            Beep
1940          End If
1950        End With
1960      Case vbKeyX
1970        With Me
1980          intRetVal = 0
1990          If .cmdPurchaseClose.Enabled = True Then
2000            cmdPurchaseClose_Click  ' ** Procedure: Below.
2010          End If
2020        End With
2030      End Select
2040    End If

        ' ** Ctrl keys.
2050    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2060      Select Case intRetVal
          Case vbKeyD
2070        With Me
2080          intRetVal = 0
2090          If IsNull(.purchaseTransDate) = False Then
2100            .purchaseAssetDate = .purchaseTransDate
2110            purchaseAssetDate_AfterUpdate  ' ** Procedure: Below.
2120            DoEvents
2130            purchaseAssetDate_Exit Cancel  ' ** Procedure: Below.
2140            DoEvents
2150          End If
2160        End With
2170      Case vbKeyN
2180        With Me
2190          intRetVal = 0
2200          MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2210        End With
2220      Case vbKeyP
2230        With Me
2240          intRetVal = 0
2250          MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2260        End With
2270      Case vbKeyT
2280        With Me
2290          intRetVal = 0
2300          DoCmd.SelectObject acForm, .Parent.Name, False
2310          .Parent.cmdSwitch.SetFocus
2320          DoEvents
2330          .Parent.cmdSwitch_Click  ' ** Form Procedure: frmJournal.
2340        End With
2350      End Select
2360    End If

        ' ** Ctrl-Shift keys.
2370    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2380      Select Case intRetVal
          Case vbKeyF
2390        With Me
2400          intRetVal = 0
2410          .FocusHolder.SetFocus
2420        End With
2430      Case vbKeyS
2440        With Me
2450          intRetVal = 0
2460          .tglPurchaseSale.SetFocus
2470        End With
2480      Case vbKeyT
2490        intRetVal = 0
2500        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
2510      End Select
2520    End If

EXITP:
2530    KeyCode = intRetVal
2540    Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
450     Resume EXITP

End Sub

Private Sub Form_Timer()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

2610    With Me
2620      .TimerInterval = 0&
          ' ** Trying to get it to go to a new record after a map reinvest.
2630      If blnMapNewRec = True Then
2640        blnMapNewRec = False
2650        If IsNull(.purchaseID) = False Then
2660          MoveRec acCmdRecordsGoToNew  ' ** Procedure: Below.
2670        End If
2680        DoEvents
2690      End If  ' ** blnMapNewRec.
2700    End With

EXITP:
2710    Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strControl As String
        Dim blnMsgGiven As Boolean

2810    blnMsgGiven = False

2820    If DataErr <> 0 Then
2830      Select Case DataErr
          Case 13, 2110, 2115, 2108, 2116, 2135, 2169, 2474, 2501, 3020, 3163, 3314, 3315, 7753, 8519
            ' ** 13    Type mismatch.
            ' ** 2110  Microsoft Access can't move the focus to the control |.
            ' ** 2501: The '|' action was Canceled.
            ' ** 3314: The field '|' cannot contain a Null value because the Required property for this field is set to True.
            'If DataErr <> 2110 And DataErr <> 3314 Then
2840        Response = acDataErrContinue
            'Else
            '  Response = acDataErrDisplay
            'End If
2850      Case 2046  ' ** The command or action isn't available now (first or last record).
            ' ** Do nothing.
2860        Response = acDataErrContinue
2870      Case 2113  ' ** The value you entered isn't valid for this field.
2880  On Error Resume Next
2890        strControl = Screen.ActiveControl.Name
2900  On Error GoTo ERRH
2910        Select Case strControl
            Case "purchaseShareFace"
2920          strControl = "Shares"
2930        Case "purchasePCash"
2940          strControl = "Principal Cash"
2950        Case "purchaseICash"
2960          strControl = "Income Cash"
2970        Case "purchaseCost"
2980          strControl = "Cost"
2990        Case "purchaseTransDate", "purchaseAssetDate"
3000          strControl = "Date"
3010          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3020          blnMsgGiven = True
3030        Case Else
3040          strControl = vbNullString
3050        End Select
3060        If blnMsgGiven = False Then
3070          MsgBox "You must enter a numeric value in the " & strControl & " field.", vbInformation + vbOKOnly, "Entry Required"
3080        End If
3090        Response = acDataErrContinue
3100        DoCmd.CancelEvent
3110      Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
            ' ** Ignore.
3120        Response = acDataErrContinue
3130      Case 2237  ' ** The text you entered isn't an item in the list.
3140        Response = acDataErrContinue
3150        DoCmd.CancelEvent
3160      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
3170        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3180        Response = acDataErrContinue
3190        DoCmd.CancelEvent
3200      Case 3101  ' ** The Microsoft Jet database engine cannot find a record in the table '|' with key matching field(s) '|'.
3210  On Error Resume Next
3220        strControl = Screen.ActiveControl.Name
3230  On Error GoTo ERRH
3240        If Me.NewRecord = False Then  ' ** (VGC 08/19/2010)
3250          If IsNull(Me.purchaseAccountNo) Or Me.purchaseAccountNo = vbNullString Then
3260            MsgBox "The posting of this transaction is incomplete." & vbCrLf & vbCrLf & _
                  "You must either complete or cancel this transaction to continue.", vbInformation + vbOKOnly, "Incomplete Data"
3270          End If
3280          If strControl <> "purchaseAccountno" And blnAccountNoHasFocus = False Then
3290            If Me.purchaseAccountNo.Enabled = True Then
3300              blnAccountNoHasFocus = True
3310              Me.purchaseAccountNo.SetFocus
3320            Else
3330              Me.purchaseTransDate.SetFocus
3340            End If
3350          End If
3360        Else
3370          Me.Undo
3380          gblnPurchaseChanged = False
3390        End If
3400        Response = acDataErrContinue
3410      Case 3162  ' ** You tried to assign the Null value to a variable that is not a Variant data type.
3420  On Error Resume Next
3430        strControl = Screen.ActiveControl.Name
3440  On Error GoTo ERRH
3450        MsgBox "The " & strControl & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
3460        Response = acDataErrContinue
3470      Case Else
3480        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3490        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3500      End Select
3510    End If

EXITP:
3520    Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.
' ** VGC 04/13/08: Calendar window seems to be Modal, preventing the
' **               user from doing anything else till it's closed.

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3610    If Not clsMonthClass Is Nothing Then
3620      If clsMonthClass.IsCalendar Then
3630        Cancel = -1
3640      Else
3650        Set clsMonthClass = Nothing
3660      End If
3670    End If

EXITP:
3680    Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3710    DetailMouse_Sub3 blnCalendar1_Focus, blnCalendar2_Focus, blnPurchaseSale_Focus, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
3720    Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Public Sub cmdPurchaseClose_Click()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPurchaseClose_Click"

        Dim strDocName As String

3810    glngJournalForm = 0&
3820    gstrFormQuerySpec = vbNullString
3830    gblnSetFocus = True
3840    strDocName = "frmJournal"
3850    DoCmd.Close acForm, strDocName

EXITP:
3860    Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub cmdPurchaseOK_Click()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPurchaseOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, ctl As Access.Control
        Dim strMsg As String, strAdvisory As String, strType As String
        Dim strRevCode As String, lngTaxcode As Long
        Dim lngJrnlID As Long, lngJrnlMS_ID As Long, lngThisCurrID As Long, strAccountNo As String, datTransDate As Date
        Dim msgResponse As VbMsgBoxResult
        Dim lngRecs As Long
        Dim blnTmpGoToSaleForm As Boolean
        Dim intPos01 As Integer
        Dim varTmp00 As Variant, curTmp01 As Currency, strTmp02 As String, strTmp03 As String, strTmp04 As String, strTmp05 As String
        Dim lngX As Long
        Dim blnContinue As Boolean

3910    blnContinue = True

3920    DoCmd.Hourglass True
3930    DoEvents

3940    blnCmdButton = True
3950    blnTmpGoToSaleForm = blnGoToSaleForm
3960    Form_Current  ' ** Procedure: Above.
3970    blnCmdButton = False

3980    With Me

3990      If VerifyAccountNo = True Then  ' ** Function: Below.

4000        lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)

            ' ** These seem to be here as a shortcut to full qualification,
            ' ** allowing reference from other modules and forms.
            ' ** I'd like to re-establish them here,
            ' ** along with a full check of their criteria.
4010        gstrPurchaseType = Nz(.purchaseType, vbNullString)
4020        gstrPurchaseAsset = CStr(CLng(Nz(.purchaseAssetNo, 0)))
4030        gstrPurchaseShareFace = CStr(CDbl(Nz(.purchaseShareFace, 0)))
4040        gstrPurchaseAccountNumber = .purchaseAccountNo
4050        gstrPurchaseICash = CStr(CDbl(Nz(.purchaseICash, 0)))
4060        gstrPurchaseICash = Rem_Dollar(gstrPurchaseICash, lngCurrID)  ' ** Module Function: modStringFuncs.
4070        gstrPurchasePCash = CStr(CDbl(Nz(.purchasePCash, 0)))
4080        gstrPurchasePCash = Rem_Dollar(gstrPurchasePCash, lngCurrID)  ' ** Module Function: modStringFuncs.
4090        gstrPurchaseCost = CStr(CDbl(Nz(.purchaseCost, 0)))
4100        gstrPurchaseCost = Rem_Dollar(gstrPurchaseCost, lngCurrID)  ' ** Module Function: modStringFuncs.
            ' ** gintShareFaceDecimals = 4

            ' ** AccountNo, JournalType, AssetNo.
4110        If gstrPurchaseAccountNumber = vbNullString And blnAccountNoHasFocus = False Then
4120          blnContinue = False
4130          DoCmd.Hourglass False
4140          MsgBox "Please enter a valid account number.", vbInformation + vbOKOnly, "Invalid Entry"
4150          blnAccountNoHasFocus = True
4160          .purchaseAccountNo.SetFocus
4170        Else
4180          If gstrPurchaseType = vbNullString Then
4190            blnContinue = False
4200            DoCmd.Hourglass False
4210            MsgBox "You must choose a Purchase, Deposit, or Liability Journal Type to continue.", vbInformation + vbOKOnly, "Entry Required"
4220            .purchaseType.SetFocus
4230          Else
4240            If gstrPurchaseAsset = vbNullString Or gstrPurchaseAsset = "0" Then
4250              blnContinue = False
4260              DoCmd.Hourglass False
4270              MsgBox "An asset must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
4280              .purchaseAssetNo.SetFocus
4290            End If
4300          End If
4310        End If
4320      Else
4330        blnContinue = False
4340      End If

          ' ** AssetDate, TransDate.
4350      If blnContinue = True Then
4360        If IsNull(.purchaseAssetDate) = True Then
4370          blnContinue = False
4380          DoCmd.Hourglass False
4390          MsgBox "You must enter a Trade Date.", vbInformation + vbOKOnly, "Entry Required"
4400          .purchaseAssetDate.SetFocus
4410        Else
4420          .purchaseAssetDate.SetFocus
4430          If Len(.purchaseAssetDate.text) = 0 Then
4440            blnContinue = False
4450            DoCmd.Hourglass False
4460            MsgBox "You must enter a Trade Date.", vbInformation + vbOKOnly, "Entry Required"
4470          Else
4480            .cmdPurchaseOK.SetFocus
4490            If Compare_DateA_DateB(CDate(.purchaseAssetDate), ">", Date) = True Then  ' ** Module Function: modStringFuncs.
4500              blnContinue = False
4510              DoCmd.Hourglass False
4520              MsgBox "Future Trade Dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
4530              .purchaseAssetDate.SetFocus
4540            Else
4550              If IsNull(.purchaseTransDate) = True Then
4560                blnContinue = False
4570                DoCmd.Hourglass False
4580                MsgBox "You must enter a Posting Date.", vbInformation + vbOKOnly, "Entry Required"
4590                .purchaseTransDate.SetFocus
4600              Else
4610                .purchaseTransDate.SetFocus
4620                If Len(.purchaseTransDate.text) = 0 Then
4630                  blnContinue = False
4640                  DoCmd.Hourglass False
4650                  MsgBox "You must enter a Posting Date.", vbInformation + vbOKOnly, "Entry Required"
4660                Else
4670                  .cmdPurchaseOK.SetFocus
4680                  If .purchaseTransDate > DateAdd("m", 1, Date) Then
4690                    blnContinue = False
4700                    DoCmd.Hourglass False
4710                    MsgBox "Only future Posting Dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
4720                    .purchaseTransDate.SetFocus
4730                  End If
4740                End If
4750              End If
4760            End If
4770          End If
4780        End If
4790      End If  ' ** blnContinue.

          ' ** If purchaseAssetDate has survived vetting, make sure it's in assetdate. (VGC 08/19/2010)
4800      If blnContinue = True Then
4810        If IsNull(.purchaseAssetDate) = False Then
              ' ** purchaseAssetDate is expected to always only show the day portion of the date.
4820          strTmp02 = Format(.purchaseAssetDate, "mm/dd/yyyy")
4830          strTmp03 = vbNullString
4840          If IsNull(.assetdate) = False Then
4850            strTmp03 = Format(.assetdate, "mm/dd/yyyy")
4860          End If
4870          If strTmp03 <> strTmp02 Then
                ' ** The days don't match.
4880            .assetdate = .purchaseAssetDate + time()
4890          End If
4900        End If
4910      End If

          ' ** Description.
4920      If blnContinue = True Then
4930        .purchaseDescription.SetFocus
4940        If InStr(.purchaseDescription.text, Chr(34)) > 0 Then
4950          blnContinue = False
4960          DoCmd.Hourglass False
4970          MsgBox "Cannot have quote marks in text.", vbInformation + vbOKOnly, "Invalid Characters"
4980          .purchaseDescription.SetFocus
4990        End If
5000      End If  ' ** blnContinue.

          ' ** ShareFace, ICash, PCash, Cost.
5010      If blnContinue = True Then

5020        strMsg = vbNullString
5030        Set ctl = Nothing
5040        If Abs(Val(gstrPurchaseShareFace)) > 999999999.9999 Then
5050          strMsg = "Share/Face is too large."
5060          Set ctl = .purchaseShareFace
5070        ElseIf Abs(Val(gstrPurchaseICash)) > 999999999.99 Then
5080          strMsg = "Accrued/Income Cash is too large."
5090          Set ctl = .purchaseICash
5100        ElseIf Abs(Val(gstrPurchasePCash)) > 999999999.99 Then
5110          strMsg = "Principal Cash is too large."
5120          Set ctl = .purchasePCash
5130        ElseIf Abs(Val(gstrPurchaseCost)) > 999999999.99 Then
5140          strMsg = "Cost is too large."
5150          Set ctl = .purchaseCost
5160        End If
5170        If strMsg <> vbNullString Then
5180          blnContinue = False
5190          DoCmd.Hourglass False
5200          MsgBox strMsg, vbInformation + vbOKOnly, "Invalid Entry"
5210          ctl.SetFocus
5220        Else

5230          strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString: strTmp05 = vbNullString
5240          intPos01 = InStr(gstrPurchaseShareFace, ".")
5250          If intPos01 > 0 Then
5260            If Len(Mid(gstrPurchaseShareFace, (intPos01 + 1))) > gintShareFaceDecimals Then
5270              strTmp02 = "Share/Face will be rounded to " & Trim(CStr(gintShareFaceDecimals)) & " decimals."
5280              gstrPurchaseShareFace = Left(gstrPurchaseShareFace, intPos01) & _
                    Left(Mid(gstrPurchaseShareFace, (intPos01 + 1)), gintShareFaceDecimals)
5290              gstrPurchaseShareFace = CDbl(Val(gstrPurchaseShareFace))  ' ** In case it's now Zero!
5300              .purchaseShareFace = CDbl(gstrPurchaseShareFace)
5310            End If
5320          End If
5330          intPos01 = InStr(gstrPurchaseICash, ".")
5340          If intPos01 > 0 And blnContinue = True Then
5350            If Len(Mid(gstrPurchaseICash, (intPos01 + 1))) > 2 Then
5360              strTmp03 = "Accrued/Income Cash will be rounded to 2 decimals."
5370              gstrPurchaseICash = Left(gstrPurchaseICash, intPos01) & Left(Mid(gstrPurchaseICash, (intPos01 + 1)), 2)
5380              gstrPurchaseICash = CDbl(Val(gstrPurchaseICash))  ' ** In case it's now Zero!
5390              gstrPurchaseICash = Rem_Dollar(gstrPurchaseICash, lngCurrID)  ' ** Module Function: modStringFuncs.
5400              .purchaseICash = CCur(gstrPurchaseICash)
5410            End If
5420          End If
5430          intPos01 = InStr(gstrPurchasePCash, ".")
5440          If intPos01 > 0 And blnContinue = True Then
5450            If Len(Mid(gstrPurchasePCash, (intPos01 + 1))) > 2 Then
5460              strTmp04 = "Principal Cash will be rounded to 2 decimals."
5470              gstrPurchasePCash = Left(gstrPurchasePCash, intPos01) & Left(Mid(gstrPurchasePCash, (intPos01 + 1)), 2)
5480              gstrPurchasePCash = CDbl(Val(gstrPurchasePCash))  ' ** In case it's now Zero!
5490              gstrPurchasePCash = Rem_Dollar(gstrPurchasePCash, lngCurrID)  ' ** Module Function: modStringFuncs.
5500              .purchasePCash = CCur(gstrPurchasePCash)
5510            End If
5520          End If
5530          intPos01 = InStr(gstrPurchaseCost, ".")
5540          If intPos01 > 0 And blnContinue = True Then
5550            If Len(Mid(gstrPurchaseCost, (intPos01 + 1))) > 2 Then
5560              strTmp05 = "Cost will be rounded to 2 decimals."
5570              gstrPurchaseCost = Left(gstrPurchaseCost, intPos01) & Left(Mid(gstrPurchaseCost, (intPos01 + 1)), 2)
5580              gstrPurchaseCost = CDbl(Val(gstrPurchaseCost))  ' ** In case it's now Zero!
5590              gstrPurchaseCost = Rem_Dollar(gstrPurchaseCost, lngCurrID)  ' ** Module Function: modStringFuncs.
5600              .purchaseCost = CCur(gstrPurchaseCost)
5610            End If
5620          End If
5630          strAdvisory = strTmp02 & "~" & strTmp03 & "~" & strTmp04 & "~" & strTmp05
5640          If strAdvisory = "~~~" Then strAdvisory = vbNullString
5650          intPos01 = 0: strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString: strTmp05 = vbNullString
5660          If Val(gstrPurchaseShareFace) = 0 Then
5670            blnContinue = False
5680            DoCmd.Hourglass False
5690            If InStr(strAdvisory, "Share/Face") > 0 Then
5700              MsgBox "Share/Face has been rounded to " & Trim(CStr(gintShareFaceDecimals)) & " decimals," & vbCrLf & _
                    "and its value is now zero. Share/Face must be greater than zero.", vbInformation + vbOKOnly, "Entry Required"
5710            Else
5720              MsgBox "Share/Face must be greater than zero.", vbInformation + vbOKOnly, "Entry Required"
5730            End If
5740            .purchaseShareFace.SetFocus
5750          Else
5760            Select Case gstrPurchaseType
                Case "Deposit"
5770              If gstrPurchaseICash <> "0" Then  'Zero icash
5780                blnContinue = False
5790                DoCmd.Hourglass False
5800                MsgBox gstrPurchaseType & " Accrued/Income Cash must be zero.", vbInformation + vbOKOnly, "Invalid Entry"
5810                .purchaseICash.SetFocus
5820              Else
5830                If gstrPurchasePCash <> "0" Then  'Zero pcash
5840                  blnContinue = False
5850                  DoCmd.Hourglass False
5860                  MsgBox gstrPurchaseType & " Principal Cash must be zero.", vbInformation + vbOKOnly, "Invalid Entry"
5870                  .purchasePCash.SetFocus
5880                Else
5890                  If Val(gstrPurchaseCost) < 0 Then  'Positive or Zero cost  CHECK! YES, MESSAGE BELOW ALLOWS ZERO COST!
5900                    blnContinue = False
5910                    DoCmd.Hourglass False
5920                    MsgBox gstrPurchaseType & " Cost cannot be negative.", vbInformation + vbOKOnly, "Invalid Entry"
5930                    .purchaseCost.SetFocus
5940                  End If
5950                End If
5960              End If
5970            Case "Purchase"
5980              If gstrPurchaseICash <> "0" Then  '(Zero pcash, Negative icash) or (Negative pcash, Negative icash)
5990                If Abs(Val(gstrPurchaseICash)) <> Abs(Val(gstrPurchaseCost)) And (Val(gstrPurchasePCash) = 0) Then
                      ' ** Applies only if icash alone is used for Purchase.
6000                  blnContinue = False
6010                  DoCmd.Hourglass False
6020                  MsgBox "To purchase an asset with Accrued/Income Cash, Cost and Accrued/Income Cash must be the same amount.", _
                        vbInformation + vbOKOnly, "Invalid Entry"
6030                  .purchaseICash.SetFocus
6040                Else
6050                  If Abs(Val(gstrPurchasePCash)) <> Abs(Val(gstrPurchaseCost)) And (Val(gstrPurchasePCash) <> 0) Then  ' ** icash <= 0.
                        ' ** This is probably handled by the screen prior to this point, but the validation here makes it consistent.
6060                    blnContinue = False
6070                    DoCmd.Hourglass False
6080                    MsgBox "When purchasing an asset with Principal Cash, Cost and Principal Cash must be the same amount.", _
                          vbInformation + vbOKOnly, "Invalid Entry"
6090                    .purchasePCash.SetFocus
6100                  Else
6110                    If Val(gstrPurchaseICash) > 0 Then
6120                      blnContinue = False
6130                      DoCmd.Hourglass False
6140                      MsgBox gstrPurchaseType & " Accrued/Income Cash must be negative.", vbInformation + vbOKOnly, "Invalid Entry"
6150                      .purchaseICash.SetFocus
6160                    End If
6170                  End If
6180                End If
6190              End If
6200              If blnContinue = True Then
6210                If Val(gstrPurchasePCash) > 0 Then  '(Negaive pcash, Zero icash)
6220                  blnContinue = False
6230                  DoCmd.Hourglass False
6240                  MsgBox gstrPurchaseType & " Principal Cash must be negative.", vbInformation + vbOKOnly, "Invalid Entry"
6250                  .purchasePCash.SetFocus
6260                Else
6270                  If Val(gstrPurchaseICash) = 0 And Val(gstrPurchasePCash) = 0 Then
6280                    blnContinue = False
6290                    DoCmd.Hourglass False
6300                    MsgBox "A " & gstrPurchaseType & " must use either Principal or Accrued/Income Cash.", _
                          vbInformation + vbOKOnly, "Invalid Entry"
6310                    .purchasePCash.SetFocus
6320                  Else
6330                    If Val(gstrPurchaseCost) <= 0 Then  'Positive cost (or Zero?  CHECK! LET'S SAY NO, THOUGH I DON'T
6340                      blnContinue = False               'FIND A MESSAGE THAT SPECIFICALLY FORBIDS OR ALLOWS IT!)
6350                      DoCmd.Hourglass False
6360                      MsgBox "A " & gstrPurchaseType & " must have a positive Cost.", vbInformation + vbOKOnly, "Invalid Entry"
6370                      .purchaseCost.SetFocus
6380                    End If
6390                  End If
6400                End If
6410              End If
                  ' ** MESSAGE BELOW SAYS ABS(PCASH) NEEDN'T EQUAL ABS(COST)
6420            Case "Liability"  ' ** LiabilityDP
6430              If gstrPurchaseICash <> "0" Then  'Zero icash
6440                blnContinue = False
6450                DoCmd.Hourglass False
6460                MsgBox "A " & gstrPurchaseType & " cannot use Accrued/Income Cash.", vbInformation + vbOKOnly, "Invalid Entry"
6470                .purchaseICash.SetFocus
6480              Else
6490                If Val(gstrPurchasePCash) < 0 Then  'Postive or Zero pcash
6500                  blnContinue = False
6510                  DoCmd.Hourglass False
6520                  MsgBox "A " & gstrPurchaseType & " cannot have negative Principal Cash.", vbInformation + vbOKOnly, "Invalid Entry"
6530                  .purchasePCash.SetFocus
6540                Else
6550                  If Val(gstrPurchaseCost) >= 0 Then  'Negative cost (or Zero cost?  CHECK! NO, MESSAGE
6560                    blnContinue = False               'BELOW SAYS IT MUST HAVE A NEGATIVE COST!)
6570                    DoCmd.Hourglass False
6580                    MsgBox "A " & gstrPurchaseType & " must have a negative Cost.", vbInformation + vbOKOnly, "Invalid Entry"
6590                    .purchaseCost.SetFocus
6600                  End If
6610                End If
6620              End If
                  ' ** May have both Zero pcash and Zero icash
6630            End Select
6640          End If
6650        End If
6660        strMsg = vbNullString
6670        Set ctl = Nothing
6680      End If  ' ** blnContinue.

          ' ** THE SECTION BELOW MAY BE MOOT!
6690      If blnContinue = True Then
6700        If ((gstrPurchaseICash = vbNullString Or Val(gstrPurchaseICash) = 0) And _
                (gstrPurchasePCash = vbNullString Or Val(gstrPurchasePCash) = 0) And _
                (gstrPurchaseCost = vbNullString Or Val(gstrPurchaseCost) = 0)) And .purchaseType <> "Deposit" Then
              ' ** Make sure there's a value somewhere!
6710          blnContinue = False
6720          DoCmd.Hourglass False
6730          MsgBox "A " & .purchaseType & " must have a value.", vbInformation + vbOKOnly, "Entry Required"
6740          .purchaseCost.SetFocus
6750        Else
6760          If .purchaseType = "Liability" Then
6770            If gstrPurchaseCost = vbNullString Or Val(gstrPurchaseCost) = 0 Then
6780              blnContinue = False
6790              DoCmd.Hourglass False
6800              MsgBox "Cost must be less than zero.", vbInformation + vbOKOnly, "Invalid Entry"
6810              .purchaseCost.SetFocus
6820            Else
6830              If CDbl(gstrPurchaseCost) > 0 Then
6840                blnContinue = False
6850                DoCmd.Hourglass False
6860                MsgBox "Cost must be negative.", vbInformation + vbOKOnly, "Invalid Entry"
6870                .purchaseCost.SetFocus
6880              Else
6890                If gstrPurchasePCash = vbNullString Then
6900                  blnContinue = False
6910                  DoCmd.Hourglass False
6920                  MsgBox "Principal Cash must have a value.", vbInformation + vbOKOnly, "Entry Required"
6930                  .purchasePCash.SetFocus
6940                Else
6950                  If CDbl(gstrPurchaseCost) <> (CDbl(gstrPurchasePCash) * -1) Then
6960                    DoCmd.Hourglass False
6970                    If MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                            vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis") = vbNo Then
6980                      blnContinue = False
6990                      .purchaseCost.SetFocus
7000                    Else
7010                      DoCmd.Hourglass True
7020                      DoEvents
7030                    End If
7040                  End If
7050                End If
7060              End If
7070            End If
7080          Else  ' ** Not a Liability.
7090            If .purchaseType = "Purchase" Then
7100              If CDbl(gstrPurchaseCost) > 0 And CDbl(gstrPurchasePCash) > 0 Then
7110                blnContinue = False
7120                DoCmd.Hourglass False
7130                MsgBox "Purchase Principal Cash must be less than zero.", vbInformation + vbOKOnly, "Invalid Entry"
7140                .purchasePCash.SetFocus
7150              End If
7160            End If
7170          End If
7180        End If
7190      End If  ' ** blnContinue.

7200      If blnContinue = True Then
7210        If gstrPurchasePCash = vbNullString Or Val(gstrPurchasePCash) = 0 Then
7220          If .purchaseICash * -1 = .purchaseCost And .purchaseCost > 0 And .purchasePCash = 0 Then
                ' ** DO NOTHING - REINVESTMENT.
7230          Else
7240            If .purchaseType <> "Deposit" Then
7250              If IsNull(.journalSubtype) Or Nz(.journalSubtype, vbNullString) <> "Reinvest" Then
7260                If .purchaseICash <> 0 And .purchasePCash = 0 Then  ' ** Checking for Null and checking for non-Null
7270                  blnContinue = False                               ' ** in same statement can be problematic.
7280                  DoCmd.Hourglass False
7290                  MsgBox "To purchase an asset with Accrued/Income Cash, Cost and Accrued/Income Cash must be the same amount.", _
                        vbInformation + vbOKOnly, "Invalid Entry"
7300                  .purchaseCost.SetFocus
7310                Else
7320                  If .purchaseType = "Liability" Then
7330                    strType = "liability."
7340                  Else
7350                    strType = "asset."
7360                  End If
7370                  DoCmd.Hourglass False
7380                  If MsgBox("Principal Cash must be greater than zero unless setting up an existing " & _
                          strType & vbCrLf & vbCrLf & "Continue anyway?", vbInformation + vbOKCancel, "Invalid Entry") = vbCancel Then
7390                    blnContinue = False
7400                    .purchasePCash.SetFocus
7410                  Else
7420                    DoCmd.Hourglass True
7430                    DoEvents
7440                  End If
7450                End If
7460              End If
7470            End If
7480          End If
7490        End If
7500      End If  ' ** blnContinue.

7510      If blnContinue = True Then
7520        If .purchaseType = "Deposit" Then
7530          If gstrPurchaseCost = vbNullString Or Val(gstrPurchaseCost) = 0 Then
7540            DoCmd.Hourglass False
7550            msgResponse = MsgBox("Are you sure you want this Deposit to have ZERO cost?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cost Basis")
7560            If msgResponse = vbNo Then
7570              blnContinue = False
7580              .purchaseCost.SetFocus
7590            Else
7600              DoCmd.Hourglass True
7610              DoEvents
7620            End If
7630          End If
7640        End If
7650      End If  ' ** blnContinue.

7660      If blnContinue = True Then
7670        If .purchaseType = "Purchase" Then
7680          If Abs(.purchasePCash) <> Abs(.purchaseCost) Then
7690            DoCmd.Hourglass False
7700            msgResponse = MsgBox("Are you sure you want the Principal Cash to be different from Cost?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Principal And Cost Unequal")
7710            If msgResponse = vbNo Then
7720              blnContinue = False
7730              .purchaseCost.SetFocus
7740            Else
7750              DoCmd.Hourglass True
7760              DoEvents
7770            End If
7780          End If
7790        End If
7800      End If  ' ** blnContinue.

7810      If blnContinue = True Then
7820        If IsNull(.purchaseAssetNo.Column(CBX_A_TYPE)) = True Then
7830          blnContinue = False
7840          DoCmd.Hourglass False
7850          MsgBox "The asset's type cannot be determined.", vbInformation + vbOKOnly, "Asset List Error"
7860        Else
7870          strTmp02 = purchaseAssetNo.Column(CBX_A_TYPE)
7880          If IsNumeric(strTmp02) = True Then
7890            If Val(strTmp02) > 0 And Val(strTmp02) < 100 Then
7900              If strTmp02 <> "90" And .purchaseType = "Liability" Then
7910                blnContinue = False
7920                DoCmd.Hourglass False
7930                MsgBox "The chosen asset is not a Liability (Asset Type 90), so cannot use the Liability Journal Type.", _
                      vbInformation + vbOKOnly, "Invalid Entry"
7940                .purchaseAssetNo.SetFocus
7950              ElseIf strTmp02 = "90" And .purchaseType <> "Liability" Then
7960                blnContinue = False
7970                DoCmd.Hourglass False
7980                MsgBox "The chosen asset is a Liability (Asset Type 90), and must use the Liability Journal Type.", _
                      vbInformation + vbOKOnly, "Invalid Entry"
7990                .purchaseType.SetFocus
8000              End If
8010            Else
8020              blnContinue = False
8030              DoCmd.Hourglass False
8040              MsgBox "The asset's type cannot be determined.", vbInformation + vbOKOnly, "Asset List Error"
8050            End If
8060          Else
8070            blnContinue = False
8080            DoCmd.Hourglass False
8090            MsgBox "The asset's type cannot be determined.", vbInformation + vbOKOnly, "Asset List Error"
8100          End If
8110        End If
8120      End If  ' ** blnContinue.

8130      If blnContinue = True Then

8140        If Trim(Nz(.purchaseUser, vbNullString)) = vbNullString Then
8150          .purchaseUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
8160        End If

8170        strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
8180        lngTaxcode = Nz(.cmbTaxCodes, 0&)
8190        Select Case .purchaseType
            Case "Purchase", "Deposit"
              ' ** INCOME.
8200          If strRevCode <> "I" Then
8210            .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
8220          End If
8230          If IsNull(.cmbTaxCodes) = True Then
8240            .cmbTaxCodes = 0&
8250          Else
8260            If gblnLinkRevTaxCodes = True Then
8270              If lngTaxcode > 0& Then
8280                If .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
8290                  .cmbTaxCodes = .purchaseAssetNo.Column(CBX_A_TAX)  ' ** All AssetType-based Tax Codes are INCOME.
8300                End If
8310              End If
8320            End If
8330          End If
8340        Case "Liability"
              ' ** EXPENSE.
8350          If strRevCode <> "E" Then
8360            .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
8370          End If
8380          If IsNull(.cmbTaxCodes) = True Then
8390            .cmbTaxCodes = 0&
8400          Else
8410            If gblnLinkRevTaxCodes = True Then
8420              If lngTaxcode = 0& Then
8430                .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
8440              Else
8450                If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
8460                  .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
8470                End If
8480              End If
8490            End If
8500          End If
8510        End Select

            ' ** One last assetdate check.
8520        If InStr(CStr(CDbl(.assetdate)), ".") = 0 Then
8530          .assetdate = .assetdate + time()
8540        End If

8550        DoCmd.RunCommand acCmdSaveRecord

8560        Set dbs = CurrentDb

            ' ** ActiveAssets, grouped by accountno, assetno, with Min(IsAverage), by specified [actno], [astno].
8570        Set qdf = dbs.QueryDefs("qryJournal_Purchase_08")
8580        With qdf.Parameters
8590          ![actno] = gstrPurchaseAccountNumber
8600          ![astno] = CLng(gstrPurchaseAsset)
8610        End With
8620        Set rst = qdf.OpenRecordset
8630        With rst
8640          If .BOF = True And .EOF = True Then
                ' ** No existing Tax Lots for this asset.
8650            Me.purchaseIsAverage = False
8660          Else
8670            .MoveFirst
8680            Me.purchaseIsAverage = ![IsAverage]
8690          End If
8700          .Close
8710        End With

8720        .priceperunit = Abs(.purchaseCost / .purchaseShareFace)

            ' ** One last assetdate check.
8730        If InStr(CStr(CDbl(.assetdate)), ".") = 0 Then
8740          .assetdate = .assetdate + time()
8750        End If

8760        DoCmd.RunCommand acCmdSaveRecord

8770      End If  ' ** blnContinue.

8780      If blnContinue = True And strAdvisory <> vbNullString Then
8790        strAdvisory = StringReplace(strAdvisory, "~", vbCrLf)  ' ** Module Function: modStringFuncs.
8800        MsgBox strAdvisory, vbInformation + vbOKOnly, "Decimal Places Adjusted"
8810      End If

          ' ** Finished checking, set up for the next entry.
8820      If blnContinue = True Then
            ' ** dbs is open.

8830        DoCmd.Hourglass True  ' ** Make sure it's still running.
8840        DoEvents

            ' ** The field behind tglPurchaseSale IS posted!
            'If .tglMiscSale = True Then
            '  .posted = True
            'Else
            '  .posted = False
            'End If

            ' ** Save these in case the Sale/Reinvest button was pushed.
            ' ** We lose them when we go to a new record.
8850        lngJrnlID = .purchaseID
8860        strAccountNo = .purchaseAccountNo.Value
8870        datTransDate = .purchaseTransDate.Value
8880        curTmp01 = CCur(Abs(Nz(.purchaseICash, 0) + Nz(.purchasePCash, 0)))
8890        lngThisCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)

            ' ** qryJournal_Purchase_01, grouped by accountno, assetno, journal_USER,
            ' ** by specified [actno], [astno], [jrnlid].
8900        Set qdf = dbs.QueryDefs("qryJournal_Purchase_04")
8910        With qdf.Parameters
8920          ![jrnlid] = Me.purchaseID
8930          ![actno] = gstrPurchaseAccountNumber
8940          ![astno] = CLng(gstrPurchaseAsset)
8950        End With
8960        Set rst = qdf.OpenRecordset()
8970        If rst.BOF = True And rst.EOF = True Then
              ' ** No records; all's well.
8980          If .purchaseCurr_Date.Visible = True Then .purchaseCurr_Date.Visible = False
8990          .Requery
9000          .AllowAdditions = True
9010          DoCmd.GoToRecord , , acNewRec
9020          DoEvents
              ' ** These seem to trigger creation of a new record!
              '.assetdate = Null
              '.purchaseAssetDate = Null
9030          datAssetDate_OldValue = 0
9040          lngAssetDate_OldValue = 0&
9050        Else
9060          strMsg = vbNullString
9070          rst.MoveLast
9080          lngRecs = rst.RecordCount
9090          rst.MoveFirst
9100          If lngRecs > 1& Then
9110            For lngX = 1& To lngRecs
9120              strMsg = strMsg & rst![journal_USER] & ", "
9130              If lngX < lngRecs Then rst.MoveNext
9140            Next
9150            strMsg = Trim(strMsg)                     ' ** Remove trailing space.
9160            strMsg = Left(strMsg, (Len(strMsg) - 1))  ' ** Remove trailing comma.
9170          Else
9180            strMsg = Nz(rst![journal_USER], CurrentUser)  ' ** Internal Access Function: Trust Accountant login.
9190          End If
9200          strMsg = IIf(strMsg = CurrentUser, "You have", "The user" & IIf(lngRecs = 1, " " & strMsg & " has", "s " & strMsg & " have")) & _
                " already entered a purchase/deposit for this asset for this account!" & vbCrLf & vbCrLf & _
                "Do you still wish to post this transaction?"  ' ** Internal Access Function: Trust Accountant login.
9210          .AllowAdditions = True
9220          If MsgBox(strMsg, vbQuestion + vbYesNo, "Warning: Possible Duplicate Data") = vbYes Then
9230            .Requery
9240            DoCmd.GoToRecord , , acNewRec
9250            DoEvents
9260            .assetdate = Null
9270            .purchaseAssetDate = Null
9280            datAssetDate_OldValue = 0
9290            lngAssetDate_OldValue = 0&
9300          Else
9310            blnContinue = False
9320            DoCmd.RunCommand (acCmdDeleteRecord)
9330          End If
9340          If .purchaseCurr_Date.Visible = True Then .purchaseCurr_Date.Visible = False
9350        End If
9360        rst.Close
9370        Set rst = Nothing
9380        Set qdf = Nothing

9390      End If  ' ** blnContinue.

9400      If IsNothing(dbs) = False Then  ' ** Module Function: modUtilities.
9410        dbs.Close
9420      End If

9430      .purchaseICash.Enabled = True
9440      .purchaseICash.BorderColor = CLR_LTBLU2
9450      .purchaseICash.BackStyle = acBackStyleNormal
9460      .purchaseICash_lbl.BackStyle = acBackStyleNormal
9470      .purchaseICash_lbl_box.Visible = False
9480      .purchasePCash.Enabled = True
9490      .purchasePCash.BorderColor = CLR_LTBLU2
9500      .purchasePCash.BackStyle = acBackStyleNormal
9510      .purchasePCash_lbl.BackStyle = acBackStyleNormal
9520      .purchasePCash_lbl_box.Visible = False

9530      gblnPurchaseValidated = False
9540      gstrPurchaseType = vbNullString
9550      gstrPurchaseAsset = vbNullString
9560      gstrPurchaseShareFace = vbNullString
9570      gstrPurchaseAccountNumber = vbNullString
9580      gstrPurchaseICash = vbNullString
9590      gstrPurchasePCash = vbNullString
9600      gstrPurchaseCost = vbNullString

9610      .purchaseAccountNo.SetFocus

9620      Sub3Purch_Changed False, Me  ' ** Module Procedure: modPurchaseSold.

9630      DoCmd.Hourglass True  ' ** Make sure it's still running.
9640      DoEvents

9650      If blnContinue = True Then

9660        blnGoToSaleForm = blnTmpGoToSaleForm

            ' ** Sale button was pushed.
9670        If blnGoToSaleForm = True Then

              ' ** Put preliminary record into tblJournal_MiscSold.
9680          Set dbs = CurrentDb
9690          Set rst = dbs.OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
9700          With rst
9710            If .BOF = True And .EOF = True Then
                  ' ** Proceed.
9720            Else
9730              .MoveFirst
9740              .FindFirst "[Journal_ID] = " & CStr(lngJrnlID)
9750              If .NoMatch = False Then
9760                .Delete
9770              End If
9780            End If
9790            DoEvents
9800            .AddNew
                ' ** ![jrnlms_id] : AutoNumber.
9810            ![Journal_ID] = lngJrnlID
9820            ![journaltype] = "Purchase"
9830            ![jrnlms_ref_id] = 0&
9840            ![jrnlms_parent] = True
9850            ![jrnlms_child] = False
9860            ![jrnlms_note] = "This Purchase was covered by a Sold transaction."
9870            ![jrnlms_datemodified] = Now()
9880            .Update
9890            .Bookmark = .LastModified
9900            lngJrnlMS_ID = ![jrnlms_id]
9910          End With
9920          Set rst = Nothing
9930          dbs.Close
9940          Set dbs = Nothing
9950          DoEvents

              ' ** Since this won't be using a Dialog, put these somewhere.
9960          .Parent.DefPurchJrnlID = lngJrnlID
9970          .Parent.DefPurchJrnlMSID = lngJrnlMS_ID

9980          gstrSaleAsset = vbNullString
9990          gstrSaleShareFace = vbNullString
10000         gstrSaleICash = vbNullString
10010         gstrSalePCash = vbNullString
10020         gstrSaleCost = vbNullString

10030         DoCmd.SelectObject acForm, .Parent.Name, False

              ' ** Retrieve a default asset if this account has one.
10040         varTmp00 = DLookup("[taxlot]", "account", "[accountno] = '" & strAccountNo & "'")
10050         If IsNull(varTmp00) = False Then
10060           If Val(varTmp00) > 0 Then
10070             .Parent.DefAssetNo = CLng(Val(varTmp00))
10080           Else
10090             .Parent.DefAssetNo = 0&
10100           End If
10110         Else
10120           .Parent.DefAssetNo = 0&
10130         End If

10140         .Parent.DefPurchTot = curTmp01
10150         .Parent.opgJournal = .Parent.opgJournal_optSale.OptionValue
10160         .Parent.SetFocus
10170         .Parent.opgJournal.SetFocus
10180         .Parent.ShowSale  ' ** Form Procedure: frmJournal.
10190         DoEvents

10200         .Parent.frmJournal_Sub4_Sold.SetFocus
              ' ** Make sure we're on a new record!
10210         .Parent.frmJournal_Sub4_Sold.Form.MoveRec acCmdRecordsGoToNew  ' ** Form Procedure: frmJournal_Sub4_Sold.
10220         DoEvents
10230         If IsNull(.Parent.frmJournal_Sub4_Sold.Form.saleID) = False Then
10240           .Parent.frmJournal_Sub4_Sold.Form.MoveRec acCmdRecordsGoToNew  ' ** Form Procedure: frmJournal_Sub4_Sold.
10250           DoEvents
10260         End If
10270         .Parent.frmJournal_Sub4_Sold.Form.saleAccountNo = strAccountNo
10280         .Parent.frmJournal_Sub4_Sold.Form.saleTransDate = datTransDate
10290         .Parent.frmJournal_Sub4_Sold.Form.saleType = "Sold"
10300         If .Parent.DefAssetNo > 0& Then
10310           .Parent.frmJournal_Sub4_Sold.Form.saleAssetno = .Parent.DefAssetNo
10320           gstrSaleAsset = CStr(.Parent.DefAssetNo)
10330           DoEvents
10340           .Parent.frmJournal_Sub4_Sold.Form.SaleChanged True  ' ** Form Procedure: frmJournal_Sub4_Sold.
10350           subTaxCode4 blnSpecialCap, intSpecialCapOpt, .Parent.frmJournal_Sub4_Sold.Form  ' ** Module Procedure: modPurchaseSold.
10360         End If
10370         If .Parent.DefPurchTot > 0@ Then
10380           .Parent.frmJournal_Sub4_Sold.Form.saleShareFace = CDbl(.Parent.DefPurchTot)
10390           .Parent.frmJournal_Sub4_Sold.Form.salePCash = .Parent.DefPurchTot
10400         End If
10410         .Parent.frmJournal_Sub4_Sold.Form.saleCost = 0
10420         .Parent.frmJournal_Sub4_Sold.Form.CheckNum = lngJrnlID  ' ** Borrowing this field for Sale/posted reference.
10430         DoEvents
10440         .Parent.DefSaleBtn = True
10450         .Parent.frmJournal_Sub4_Sold.Form.cmdSaleOK.Enabled = True
10460         .Parent.frmJournal_Sub4_Sold.Form.cmdSaleClose.Enabled = False
10470         .Parent.frmJournal_Sub4_Sold.Form.cmdSaleLotInfo.Enabled = False
10480         .Parent.NavVis False  ' ** Form Procedure: frmJournal.
10490         gblnSaleChanged = True
10500         .Parent.frmJournal_Sub4_Sold.Form.saleAccountNo.SetFocus
10510         .Parent.frmJournal_Sub4_Sold.Form.saleAssetno.SetFocus

10520       End If
10530       blnGoToSaleForm = False  ' ** reset always.

10540     End If  ' ** blnContinue.

10550     DoCmd.Hourglass False

10560   End With  ' ** Me.

EXITP:
10570   Set ctl = Nothing
10580   Set rst = Nothing
10590   Set qdf = Nothing
10600   Set dbs = Nothing
10610   Exit Sub

ERRH:
410     DoCmd.Hourglass False
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub cmdPurchaseCancel_Click()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPurchaseCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strControl As String
        Dim lngThisJrnlID As Long, lngThatJrnlID As Long, strMsg As String
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

10710   With Me

10720     DoCmd.Hourglass True
10730     DoEvents

10740     If IsNull(.purchaseID) = True Then
10750       msgResponse = vbCancel
10760     Else

10770 On Error Resume Next
10780       strControl = Screen.ActiveControl.Name
10790 On Error GoTo ERRH

10800       lngThisJrnlID = .purchaseID
10810       lngThatJrnlID = 0&
10820       msgResponse = 0
            ' ** 2  vbCancel
            ' ** 6  vbYes
            ' ** 7  vbNo

10830 On Error Resume Next
10840       varTmp00 = .posted
10850       If ERR.Number <> 0 Then
              ' ** 3167  Record is deleted.
10860 On Error GoTo ERRH
10870       Else
10880 On Error GoTo ERRH
10890 On Error Resume Next
10900         varTmp00 = .CheckNum
10910         If ERR.Number <> 0 Then
                ' ** 3167  Record is deleted.
10920 On Error GoTo ERRH
10930         Else
10940 On Error GoTo ERRH
                ' ** Check whether this was covered using the Sale button.
10950           If .posted = True And IsNull(.CheckNum) = True Then
                  ' ** Check if the referencing entry still exists.
10960             varTmp00 = DCount("*", "journal", "[CheckNum] = " & CStr(lngThisJrnlID))
10970             If varTmp00 = 1 Then
10980               lngThatJrnlID = DLookup("[ID]", "journal", "[CheckNum] = " & CStr(lngThisJrnlID))
10990               strMsg = "This entry was covered by a Sold transaction." & vbCrLf & vbCrLf & _
                      "Yes - Delete both this Purchase and the covering Sold." & vbCrLf & _
                      "No - Delete only this Purchase." & vbCrLf & _
                      "Cancel - Delete no transactions."
11000               Beep
11010               DoCmd.Hourglass False
11020               msgResponse = MsgBox(strMsg, vbQuestion + vbYesNoCancel + vbDefaultButton3, "Delete Both Transactions")
11030               If msgResponse <> vbCancel Then
11040                 DoCmd.Hourglass True
11050                 If .purchaseCurr_Date.Visible = True Then .purchaseCurr_Date.Visible = False
11060                 DoEvents
11070               End If
11080             Else
11090               lngThatJrnlID = 0&
11100             End If
11110           End If
11120         End If
11130       End If

11140     End If

11150     If msgResponse = 0 Then
11160       Select Case blnGoneToReport
            Case True
11170         msgResponse = vbNo  ' ** Delete only Purchase.
11180       Case False
11190         DoCmd.Hourglass False
11200         msgResponse = MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo, "Delete Transaction")
11210         If msgResponse <> vbYes Then
11220           msgResponse = vbCancel
11230         Else
11240           msgResponse = vbNo  ' ** Delete only Purchase.
11250           DoCmd.Hourglass True
11260           If .purchaseCurr_Date.Visible = True Then .purchaseCurr_Date.Visible = False
11270           DoEvents
11280         End If
11290       End Select
11300     End If

11310     If msgResponse = vbCancel Then
11320       If strControl <> "purchaseAccountno" And blnAccountNoHasFocus = False Then
11330         If .purchaseAccountNo.Enabled = True Then
11340           blnAccountNoHasFocus = True
11350           .purchaseAccountNo.SetFocus
11360         Else
11370           .purchaseTransDate.SetFocus
11380         End If
11390       End If
11400     Else

            ' ** If this Purchase came from a Reinvest reference,
            ' ** that journal ID may be in CheckNum.
11410       If IsNull(.CheckNum) = False Then
11420         lngThatJrnlID = .CheckNum
11430         Set dbs = CurrentDb
11440         With dbs
                ' ** Delete both tblJournal_MiscSold records, since with
                ' ** this one gone, the parent no longer has a child.
                ' ** This one may already be gone because of Cascade Delete.
                ' **
                ' ** Delete tblJournal_MiscSold, by specified [jrnlid].
11450           Set qdf = .QueryDefs("qryJournal_MiscSold_02")
11460           With qdf.Parameters
11470             ![jrnlid] = lngThisJrnlID
11480           End With
11490           qdf.Execute
11500           Set qdf = Nothing
11510           DoEvents
                ' ** Delete tblJournal_MiscSold, by specified [jrnlid].
11520           Set qdf = .QueryDefs("qryJournal_MiscSold_02")
11530           With qdf.Parameters
11540             ![jrnlid] = lngThatJrnlID
11550           End With
11560           qdf.Execute
11570           Set qdf = Nothing
11580           DoEvents
                ' ** Update Journal, by specified [jrnlid], [pstd].
11590           Set qdf = .QueryDefs("qryJournal_Misc_11")
11600           With qdf.Parameters
11610             ![jrnlid] = lngThatJrnlID
11620             ![pstd] = False
11630           End With
11640           qdf.Execute
11650           Set qdf = Nothing
11660           .Close
11670         End With
11680         Set dbs = Nothing
11690       ElseIf msgResponse = vbNo And lngThatJrnlID > 0& Then
11700         Set dbs = CurrentDb
11710         With dbs
                ' ** Update Journal, for CheckNum = Null, by specified [chknum].
11720           Set qdf = .QueryDefs("qryJournal_Purchase_10")
11730           With qdf.Parameters
11740             ![chknum] = lngThatJrnlID
11750           End With
11760           qdf.Execute
11770           Set qdf = Nothing
11780           .Close
11790         End With
11800         Set dbs = Nothing
11810       End If
11820       DoEvents

11830       .Undo
11840       .Undo

11850       If .NewRecord = True Then
11860         DoCmd.RunCommand acCmdUndo
11870         .purchaseAccountNo = vbNullString
11880       Else
11890         DoCmd.SetWarnings False
11900         DoCmd.RunCommand acCmdSelectRecord
11910         DoCmd.RunCommand acCmdDeleteRecord
11920         DoCmd.SetWarnings True
11930       End If

11940       .purchaseICash.Enabled = True
11950       .purchaseICash.BorderColor = CLR_LTBLU2
11960       .purchaseICash.BackStyle = acBackStyleNormal
11970       .purchaseICash_lbl.BackStyle = acBackStyleNormal
11980       .purchaseICash_lbl_box.Visible = False
11990       .purchasePCash.Enabled = True
12000       .purchasePCash.BorderColor = CLR_LTBLU2
12010       .purchasePCash.BackStyle = acBackStyleNormal
12020       .purchasePCash_lbl.BackStyle = acBackStyleNormal
12030       .purchasePCash_lbl_box.Visible = False

12040       .AllowAdditions = True

12050       blnGoToSaleForm = False
12060       .tglPurchaseSale_false_raised_img_dis.Visible = True
12070       .tglPurchaseSale_false_raised_img.Visible = False
12080       .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = False
12090       .tglPurchaseSale_false_raised_focus_img.Visible = False
12100       .tglPurchaseSale_false_raised_focus_dots_img.Visible = False
12110       .tglPurchaseSale_false_sunken_focus_dots_img.Visible = False
12120       .tglPurchaseSale_true_raised_img.Visible = False
12130       .tglPurchaseSale_true_raised_focus_img.Visible = False
12140       .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
12150       .tglPurchaseSale_true_sunken_focus_dots_img.Visible = False
12160       .tglPurchaseSale_true_raised_img_dis.Visible = False
12170       .tglPurchaseSale.Enabled = False

12180       .cmbAccountHelper = Null

            ' ** Also delete covering entry.
12190       If msgResponse = vbYes And lngThatJrnlID > 0& Then
12200         Set dbs = CurrentDb
              ' ** Delete Journal, by specified [chknum].
12210         Set qdf = dbs.QueryDefs("qryJournal_Purchase_11")
12220         With qdf.Parameters
12230           ![chknum] = lngThisJrnlID  ' ** This ID is in that entry's CheckNum.
12240         End With
12250         qdf.Execute
12260         Set qdf = Nothing
12270         dbs.Close
12280         Set dbs = Nothing
12290       End If

12300       .Parent.frmJournal_Sub4_Sold.Form.Requery
12310       DoEvents

12320       If strControl <> "purchaseAccountno" And blnAccountNoHasFocus = False Then
12330         If .purchaseAccountNo.Enabled = True Then
12340           blnAccountNoHasFocus = True
12350           .purchaseAccountNo.SetFocus
12360         Else
12370           .purchaseTransDate.SetFocus
12380         End If
12390       End If

12400       gblnPurchaseValidated = False

12410       Sub3Purch_Changed False, Me  ' ** Module Procedure: modPurchaseSold.

12420       Form_Current  ' ** Procedure: Above.

12430     End If

12440     DoCmd.Hourglass False

12450   End With

EXITP:
12460   Set qdf = Nothing
12470   Set dbs = Nothing
12480   Exit Sub

ERRH:
410     DoCmd.Hourglass False
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub purchaseAccountNo_AfterUpdate()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAccountNo_AfterUpdate"

12510   Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.

EXITP:
12520   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAccountNo_KeyDown"

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
          Case vbKeyTab
            ' ** Tab key copies AccountNo. {intRetVal not cleared.}
12670       With Me
12680         If gblnTabCopyAccount = True Then
12690           If .NewRecord = True Then
12700             If IsNull(.purchaseAccountNo) Then
12710               If Len(.purchaseAccountNo.text) = 0 Then
12720                 If IsNull(.Parent.LastAcctNo) = False Then
12730                   If .Parent.LastAcctNo <> vbNullString Then
12740                     .purchaseAccountNo = .Parent.LastAcctNo
12750                     Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
12760                   End If
12770                 End If
12780               End If
12790             Else
12800               If Len(.purchaseAccountNo.text) = 0 Then
12810                 If IsNull(.Parent.LastAcctNo) = False Then
12820                   If .Parent.LastAcctNo <> vbNullString Then
12830                     .purchaseAccountNo = .Parent.LastAcctNo
12840                     Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
12850                   End If
12860                 End If
12870               End If
12880             End If
12890           End If
12900         End If
12910         gblnPurchaseChanged = True
12920         gblnPurchaseValidated = False
12930       End With
12940     Case vbKeyReturn
12950       With Me
12960         intRetVal = 0
12970         gblnPurchaseChanged = True  ' ** Just added: 09/19/2012.
12980         gblnPurchaseValidated = False
12990         .purchaseType.SetFocus
13000       End With
13010     Case vbKeyDelete
13020       gblnPurchaseValidated = False
13030     End Select
13040   End If

EXITP:
13050   KeyCode = intRetVal
13060   Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub purchaseAccountNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAccountNo_MouseMove"

13110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13120   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseAccountNo_Exit(Cancel As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAccountNo_Exit"

        Dim strAccountNo As String
        Dim blnGoToAccountHelper As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

13210   With Me

13220     blnGoToAccountHelper = False   ' ** Unless set otherwise.

13230     If ((gblnPurchaseValidated = False And gblnPurchaseChanged = True) And (lngErrCnt < 3&)) Then
13240       If IsNull(.purchaseAccountNo) = True Then
13250         lngErrCnt = lngErrCnt + 1&
13260         msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Entry Required")
13270         If msgResponse = vbYes Then
13280           blnGoToAccountHelper = True
13290         Else
13300           gblnPurchaseValidated = False
13310           If blnStartTrans = True Then
13320             Cancel = -1
13330             .purchaseAccountNo.SetFocus
13340           End If
13350         End If
13360       Else
13370         If .purchaseAccountNo = vbNullString And .purchaseAccountNo.text = vbNullString Then
13380           lngErrCnt = lngErrCnt + 1&
13390           msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                  "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Entry Required")
13400           If msgResponse = vbYes Then
13410             blnGoToAccountHelper = True
13420           Else
13430             gblnPurchaseValidated = False
13440             If blnStartTrans = True Then
13450               Cancel = -1
13460               .purchaseAccountNo.SetFocus
13470             End If
13480           End If
13490         Else
                ' ** Check to see if it is a valid account.
13500           lngErrCnt = 0&
13510           Select Case gblnAccountNoWithType
                Case True
13520             strAccountNo = Left(.purchaseAccountNo, 2) & Mid(.purchaseAccountNo, 3)
13530           Case False
13540             strAccountNo = .purchaseAccountNo
13550           End Select
13560           varTmp00 = DLookup("[accountno]", "account", "[accountno] = '" & strAccountNo & "' And ([dateclosed] Is Null)")
13570           Select Case IsNull(varTmp00)
                Case True
13580             blnStartTrans = True  ' ** Once they've started, they've got to enter a valid number!
13590             msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                    "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Invalid Entry")
13600             If msgResponse = vbYes Then
13610               blnGoToAccountHelper = True
13620             Else
13630               gblnPurchaseValidated = False
13640               Cancel = -1
13650               .purchaseAccountNo.SetFocus
13660             End If
13670             .purchaseAccountNo = vbNullString
13680             blnAccountNoHasFocus = False
13690           Case False
13700             gblnPurchaseValidated = True
13710           End Select
13720           If IsNull(.purchaseAccountNo) = False Then
13730             If .purchaseAccountNo <> vbNullString Then
13740               .purchaseAccountNo_Data = .purchaseAccountNo
13750               gstrPurchaseAccountNumber = .purchaseAccountNo.text
13760               .Parent.LastAcctNo = gstrPurchaseAccountNumber
13770               .cmdPurchaseMap.Enabled = False
13780 On Error Resume Next
13790               .Refresh
13800               If ERR.Number <> 0 Then  ' ** 3314: The field 'journal.journaltype' cannot contain a Null value because the Required property for this field is set to True.
13810                 .purchaseType = "Purchase"
13820                 blnAccountNoErr = True
13830               End If
13840 On Error GoTo ERRH
13850               .Refresh
13860             End If
13870           End If
13880         End If
13890       End If
13900     Else
13910       lngErrCnt = 0&
13920     End If

13930     blnAccountNoHasFocus = False
13940     If blnGoToAccountHelper = True Then
13950       .cmbAccountHelper.SetFocus
13960     End If

13970   End With  ' ** Me.

EXITP:
13980   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub cmdLock_Click()

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLock_Click"

14010   With Me
14020     If .cmdLock_open_raised_img.Visible = True Then
14030       .purchaseTransDate.SetFocus
14040       .purchaseAccountNo.Enabled = False
14050       .purchaseAccountNo.BorderColor = WIN_CLR_DISR
14060       .purchaseAccountNo.BackStyle = acBackStyleTransparent
14070       .purchaseAccountNo_lbl.BackStyle = acBackStyleTransparent
14080       .purchaseAccountNo_lbl_box.Visible = True
14090       .cmdLock_open_raised_img.Visible = False
14100       .cmdLock_closed_raised_img.Visible = True
14110     ElseIf .cmdLock_closed_raised_img.Visible = True Then
14120       .purchaseAccountNo.Enabled = True
14130       .purchaseAccountNo.BorderColor = CLR_LTBLU2
14140       .purchaseAccountNo.BackStyle = acBackStyleNormal
14150       .purchaseAccountNo_lbl.BackStyle = acBackStyleNormal
14160       .purchaseAccountNo_lbl_box.Visible = False
14170       .purchaseAccountNo.SetFocus
14180       .cmdLock_open_raised_img.Visible = True
14190       .cmdLock_closed_raised_img.Visible = False
14200     End If
14210   End With

EXITP:
14220   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbAccountHelper_Enter()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_Enter"

14310   With Me
14320     Select Case IsNull(.cmbAccountHelper)
          Case True
14330       ExpandCombo .cmbAccountHelper  ' ** Module Procedure: modUtilities.
14340     Case False
14350       If Trim(.cmbAccountHelper) = vbNullString Then
14360         ExpandCombo .cmbAccountHelper  ' ** Module Procedure: modUtilities.
14370       End If
14380     End Select
14390   End With

EXITP:
14400   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbAccountHelper_GotFocus()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_GotFocus"

14510   With Me
14520     .cmbAccountHelper.ForeColor = 0
14530   End With

EXITP:
14540   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbAccountHelper_AfterUpdate()

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_AfterUpdate"

14610   With Me
14620     If IsNull(.cmbAccountHelper) = False Then
14630       .purchaseAccountNo = .cmbAccountHelper
14640       .purchaseAccountNo.SetFocus
14650     End If
14660   End With

EXITP:
14670   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
420     Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub cmbAccountHelper_LostFocus()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_LostFocus"

14710   With Me
14720     gblnPurchaseValidated = False
14730     Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
14740     .cmbAccountHelper.ForeColor = CLR_WHT
14750     .purchaseAccountNo.SetFocus
14760     DoEvents
14770     .purchaseAccountNo.Value = .cmbAccountHelper.Value
14780     DoEvents
14790   End With

EXITP:
14800   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseShortName_LostFocus()

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseShortName_LostFocus"

14910   VerifyAccountNo  ' ** Function: Below.

EXITP:
14920   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseType_Enter()

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseType_Enter"

15010   With Me
15020     If IsNull(.purchaseAccountNo) = False Then
15030       If Trim(.purchaseAccountNo) <> vbNullString Then

15040         If IsNull(.purchaseType) = True Or blnAccountNoErr = True Then
15050           blnAccountNoErr = False
15060           blnTypeIsNull = True
15070           ExpandCombo .purchaseType  ' ** Module Procedure: modUtilities.
15080           .purchaseType = .purchaseType.ItemData(0)  ' ** Currious that purchaseType_AfterUpdate() stopped being triggered!?
15090           blnDefTypeAssigned = True
15100         Else
15110           blnTypeIsNull = False
15120         End If

15130 On Error Resume Next
15140         DoCmd.RunCommand acCmdSaveRecord
15150 On Error GoTo ERRH

15160         If .purchaseType <> .purchaseType.OldValue Or IsNull(.purchaseType.OldValue) Then
15170           Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
15180         End If

15190         VerifyAccountNo  ' ** Function: Below.

15200       End If
15210     End If
15220   End With

EXITP:
15230   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseType_AfterUpdate()

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseType_AfterUpdate"

15310   JournalType_After_Sub3 blnTypeIsNull, blnGoToSaleForm, blnDefTypeAssigned, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
15320   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseType_Exit(Cancel As Integer)

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseType_Exit"

15410   With Me

15420     If IsNull(.purchaseAccountNo) = False Then
15430       If Trim(.purchaseAccountNo) <> vbNullString Then
15440         VerifyAccountNo  ' ** Function: Below.
15450       End If
15460     End If

          ' ** Currious that purchaseType_AfterUpdate() stopped being triggered!?
15470     If blnDefTypeAssigned = True Then
15480       purchaseType_AfterUpdate  ' ** Procedure: Above.
15490       DoEvents
15500     End If

15510     If IsNull(.purchaseType) = False Then
15520       If .purchaseType.Column(0) = "Liability" Or .purchaseType.Column(0) = "Liability (+)" Then
15530         If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03b" Then
15540           .purchaseAssetNo.RowSource = "qryJournal_Purchase_03b"  '#curr_id
15550         End If
15560       Else
15570         If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03c" Then
15580           .purchaseAssetNo.RowSource = "qryJournal_Purchase_03c"  '#curr_id
15590         End If
15600       End If
15610     End If

15620   End With

EXITP:
15630   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseType_LostFocus()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseType_LostFocus"

15710   With Me
15720     If IsNull(.purchaseAccountNo) = False Then
15730       If Trim(.purchaseAccountNo) <> vbNullString Then

15740         gstrPurchaseType = .purchaseType.text

15750         If .purchaseType <> .purchaseType.OldValue Or IsNull(.purchaseType.OldValue) Then
15760           Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
15770         End If

15780         Sub3Purch_TaxCode Me  ' ** Module Procedure: modPurchaseSold.
15790         Sub3Purch_RevCode Me  ' ** Module Procedure: modPurchaseSold.

15800         VerifyAccountNo  ' ** Function: Below.

15810       End If
15820     End If

15830     If IsNull(.purchaseType) = False Then
15840       If .purchaseType.Column(0) = "Liability" Or .purchaseType.Column(0) = "Liability (+)" Then
15850         If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03b" Then
15860           .purchaseAssetNo.RowSource = "qryJournal_Purchase_03b"  '#curr_id
15870         End If
15880       Else
15890         If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03c" Then
15900           .purchaseAssetNo.RowSource = "qryJournal_Purchase_03c"  '#curr_id
15910         End If
15920       End If
15930     End If

15940   End With

EXITP:
15950   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Public Sub cmdPurchaseMap_Click()
' ** Public for Go-To-Report.

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPurchaseMap_Click"

        Dim strDocName As String
        Dim blnContinue As Boolean

16010   blnContinue = True

16020   With Me

16030     If .purchaseAccountNo.Enabled = True Then
16040       .purchaseAccountNo.SetFocus
16050     Else
16060       .purchaseTransDate.SetFocus
16070     End If
16080     If IsNull(.purchaseAccountNo) Then
            ' ** Open the map.
16090     Else
16100       If Len(Trim(.purchaseAccountNo.text)) = 0 Then
              ' ** Open the map.
16110       Else
16120         blnContinue = False
16130         .cmdPurchaseMap.SetFocus
16140         MsgBox "Please advance to a blank posting form.", vbInformation + vbOKOnly, "New Record Required"
16150         DoCmd.CancelEvent
16160       End If
16170     End If

16180     If blnContinue = True Then

16190       .cmdPurchaseMap.SetFocus

16200       .Parent.LoadMapDropdowns  ' ** Form Procedure: frmJournal.
16210       DoEvents

16220       strDocName = "frmMap_Split"
16230       DoCmd.OpenForm strDocName, , , , , , "frmJournal"

16240       If gblnGoToReport = True Then
16250         Forms(strDocName).TimerInterval = 100&
16260         .GoToReport_arw_mapsplit_img.Visible = False
16270         DoCmd.Hourglass True  ' ** Make sure it's still running.
16280         DoEvents
16290       End If

16300     End If

16310   End With

EXITP:
16320   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub cmdPurchaseMap_KeyDown(KeyCode As Integer, Shift As Integer)
' ** See Note, above.

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPurchaseMap_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16420   intShiftDown = (Shift And acShiftMask) > 0
16430   intAltDown = (Shift And acAltMask) > 0
16440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16460     Select Case intRetVal
          Case vbKeyTab
16470       With Me
16480         intRetVal = 0
16490         .purchaseTransDate.SetFocus
16500       End With
16510     End Select
16520   End If

        ' ** Shift keys.
16530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16540     Select Case intRetVal
          Case vbKeyTab
16550       With Me
16560         intRetVal = 0
16570         .purchaseType.SetFocus
16580       End With
16590     End Select
16600   End If

EXITP:
16610   KeyCode = intRetVal
16620   Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub purchaseTransdate_BeforeUpdate(Cancel As Integer)

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseTransdate_BeforeUpdate"

16710   With Me
16720     If .purchaseTransDate.text <> "__/__/____" And .purchaseTransDate.text <> vbNullString Then
16730       If Left(.purchaseTransDate.text, 2) > 12 Then
16740         MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Date"
16750         Cancel = -1
16760       End If
16770     End If
16780   End With

EXITP:
16790   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseTransDate_KeyDown(KeyCode As Integer, Shift As Integer)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseTransDate_KeyDown"

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
16890         .purchaseAssetNo.SetFocus
16900       End With
16910     End Select
16920   End If

        ' ** Shift keys.
16930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16950       With Me
16960         intRetVal = 0
16970         .purchaseType.SetFocus
16980       End With
16990     End Select
17000   End If

        ' ** Ctrl keys.
17010   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
17020     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17030       With Me
17040         intRetVal = 0
17050         .cmdCalendar1.SetFocus
17060       End With
17070     End Select
17080   End If

EXITP:
17090   KeyCode = intRetVal
17100   Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub purchaseTransDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseTransDate_MouseMove"

17210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17220   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseTransdate_Exit(Cancel As Integer)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseTransdate_Exit"

17310   With Me

17320     If lngDateCnt <= 1& Then
17330       If IsNull(.purchaseAccountNo) = False Then
17340         If Trim(.purchaseAccountNo) <> vbNullString Then
17350           If DateCheck_Post(.purchaseTransDate.text) = True Then  ' ** Module Function: modUtilities.
                  ' ** '.Text' needed to reliably check date.
17360             lngDateCnt = 0&
17370             If IsNull(.purchaseTransDate.OldValue) Then
17380               Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
17390             Else
17400               If .purchaseTransDate <> .purchaseTransDate.OldValue Then
17410                 Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
17420               End If
17430             End If
17440           Else
17450             If IsNull(.purchaseTransDate) = True Then
17460               .purchaseTransDate = Date
17470             Else
17480               If Trim(.purchaseTransDate) = vbNullString Then
17490 On Error Resume Next
17500                 .purchaseTransDate = Date
17510                 If ERR.Number <> 0 Then
                        ' ** 2473  The expression |2 you entered as the event property setting produced the following error: |1.
                        ' **       * The expression may not result in the name of a macro, the name of a user-defined function, or [Event Procedure].
                        ' **       * There may have been an error evaluating the function, event, or macro.
17520                   .purchaseTransDate = TodaysDate  ' ** Module Function: modUtilities.
17530 On Error GoTo ERRH
17540                 Else
17550 On Error GoTo ERRH
17560                 End If
17570               Else
17580                 lngDateCnt = lngDateCnt + 1&
17590                 Cancel = -1
17600                 .purchaseTransDate.SetFocus
17610               End If
17620             End If
17630           End If
17640         End If
17650       End If
17660     Else
17670       lngDateCnt = 0&
17680       .purchaseTransDate = CDate(Format(datPostingDate, "mm/dd/yyyy"))
17690     End If

17700     If IsNull(.purchaseType) = False Then
17710       If .purchaseType.Column(0) = "Liability" Or .purchaseType.Column(0) = "Liability (+)" Then
17720         If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03b" Then
17730           .purchaseAssetNo.RowSource = "qryJournal_Purchase_03b"  '#curr_id
17740         End If
17750       Else
17760         If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03c" Then
17770           .purchaseAssetNo.RowSource = "qryJournal_Purchase_03c"  '#curr_id
17780         End If
17790       End If
17800     End If

17810   End With

EXITP:
17820   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseTransdate_LostFocus()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseTransdate_LostFocus"

17910   With Me

17920     If IsNull(.purchaseAccountNo) = False Then
17930       If Trim(.purchaseAccountNo) <> vbNullString Then
17940         VerifyAccountNo  ' ** Function: Below.
17950       End If
17960     End If

17970     If IsNull(.purchaseType) = False Then
17980       If .purchaseType.Column(0) = "Liability" Or .purchaseType.Column(0) = "Liability (+)" Then
17990         If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03b" Then
18000           .purchaseAssetNo.RowSource = "qryJournal_Purchase_03b"  '#curr_id
18010         End If
18020       Else
18030         If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03c" Then
18040           .purchaseAssetNo.RowSource = "qryJournal_Purchase_03c"  '#curr_id
18050         End If
18060       End If
18070     End If

18080   End With

EXITP:
18090   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

18110   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
18120   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

18210   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
18220   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Public Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

18310   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
18320   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
420     Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

18410   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
18420   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

18510   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
18520   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

18610   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
18620   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseAssetNo_Enter()

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetNo_Enter"

        Static intLoopCnt As Integer

18710   With Me

          ' ** Make sure the Combo Box is populated.
18720     If intLoopCnt = 0 Then
18730       intLoopCnt = 1
18740       blnCmdButton = True
18750       Form_Current  ' ** Procedure: Above.
18760       blnCmdButton = False
18770     Else
18780       intLoopCnt = 0
18790     End If

18800     .purchaseAssetNo.SetFocus  ' ** Because the OnCurrent() event may move it.
18810     .purchaseAssetNo.Requery
18820     DoEvents
18830     If IsNull(.purchaseAssetNo) = True Or .purchaseAssetNo.text = vbNullString Then
18840       ExpandCombo .purchaseAssetNo  ' ** Module Procedure: modUtilities.
18850     End If

18860   End With

EXITP:
18870   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseAssetNo_BeforeUpdate(Cancel As Integer)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetNo_BeforeUpdate"

        Dim blnCurrID As Boolean
        Dim varTmp00 As Variant

18910   With Me
18920     If IsNull(.purchaseAssetNo) = False Then
18930       blnCurrID = .purchaseCurr_ID.Visible
18940       If blnCurrID = False Then
18950         varTmp00 = .purchaseAssetNo.Column(CBX_A_CURRID)
18960         If IsNull(varTmp00) = False Then
18970           If CLng(varTmp00) <> 150& Then
18980             Beep
18990             MsgBox "The chosen asset is in a foreign currency." & vbCrLf & vbCrLf & _
                    "You must check the 'Include Currency' checkbox on the Posting Date" & vbCrLf & _
                    "window before entering transactions not in U.S. dollars.", vbInformation + vbOKOnly, "Foreign Currency Asset"
19000             .purchaseAssetNo.Undo
19010             Cancel = -1
19020           End If
19030         End If
19040       End If
19050     End If
19060   End With

EXITP:
19070   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseAssetNo_AfterUpdate()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetNo_AfterUpdate"

19110   With Me

          ' ** Do this whether or not Currency is visible.
19120     Select Case IsNull(.purchaseAssetNo)
          Case True
19130       lngCurrID = 150&  ' ** USD.
19140     Case False
19150       Select Case IsNull(.purchaseAssetNo.Column(CBX_A_CURRID))
            Case True
19160         lngCurrID = 150&  ' ** USD.
19170       Case False
19180         lngCurrID = .purchaseAssetNo.Column(CBX_A_CURRID)
19190       End Select
19200     End Select
19210     .purchaseCurr_ID = lngCurrID
19220     purchaseCurr_ID_AfterUpdate  ' ** Procedure: Below.

19230     Sub3Purch_TaxCode Me  ' ** Module Procedure: modPurchaseSold.
19240     ChkAverage  ' ** Procedure: Below.

19250   End With

EXITP:
19260   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseAssetNo_KeyDown(KeyCode As Integer, Shift As Integer)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetNo_KeyDown"

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
19390         .purchaseShareFace.SetFocus
19400       End With
19410     End Select
19420   End If

        ' ** Shift keys.
19430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19450       With Me
19460         intRetVal = 0
19470         .purchaseTransDate.SetFocus
19480       End With
19490     End Select
19500   End If

        ' ** Ctrl keys.
19510   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
19520     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19530       With Me
19540         intRetVal = 0
19550         .cmdPurchaseAssetNew.SetFocus
19560       End With
19570     End Select
19580   End If

        ' ** Ctrl-Shift keys.
19590   If intCtrlDown And (Not intAltDown) And intShiftDown Then
19600     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19610       With Me
19620         intRetVal = 0
19630         .cmdCalendar1.SetFocus
19640       End With
19650     End Select
19660   End If

EXITP:
19670   KeyCode = intRetVal
19680   Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub purchaseAssetNo_Change()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetNo_Change"

19710   With Me
19720     If .purchaseAssetNo <> .purchaseAssetNo.OldValue Then
19730       Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
19740     End If
19750   End With

EXITP:
19760   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseAssetNo_NotInList(NewData As String, Response As Integer)

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetNo_NotInList"

19810   With Me
19820     .purchaseAssetNo.SelStart = 0
19830     .purchaseAssetNo.SelLength = Len(.purchaseAssetNo.text)
19840   End With

EXITP:
19850   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseAssetNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetNo_MouseMove"

19910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
19920   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseAssetNo_Exit(Cancel As Integer)

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetNo_Exit"

        Dim varTmp00 As Variant

20010   With Me
20020     If VerifyAccountNo = True Then  ' ** Function: Below.
20030       .purchaseAssetNo.Requery
20040       DoEvents
20050     End If
          ' ** Check curr_id.
20060     varTmp00 = .purchaseAssetNo.Column(CBX_A_CURRID)
20070     If IsNull(varTmp00) = False Then
20080       If .purchaseCurr_ID <> varTmp00 Then
20090         .purchaseCurr_ID = varTmp00
20100         purchaseCurr_ID_AfterUpdate  ' ** Procedure: Below.
20110       End If
20120     End If
20130   End With

EXITP:
20140   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseAssetNo_LostFocus()

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetNo_LostFocus"

20210   With Me
20220     gstrPurchaseAsset = .purchaseAssetNo.text
20230   End With

EXITP:
20240   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmdPurchaseAssetNew_Click()

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPurchaseAssetNew_Click"

        Dim strDocName As String

20310   strDocName = "frmAssets_Add_Purchase"
20320   DoCmd.OpenForm strDocName, acNormal, , , , acDialog, "frmJournal"

EXITP:
20330   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub cmdPurchaseAssetNew_KeyDown(KeyCode As Integer, Shift As Integer)
' ** See Note, above.

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPurchaseAssetNew_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20420   intShiftDown = (Shift And acShiftMask) > 0
20430   intAltDown = (Shift And acAltMask) > 0
20440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20460     Select Case intRetVal
          Case vbKeyTab
20470       With Me
20480         intRetVal = 0
20490         .purchaseShareFace.SetFocus
20500       End With
20510     End Select
20520   End If

        ' ** Shift keys.
20530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20540     Select Case intRetVal
          Case vbKeyTab
20550       With Me
20560         intRetVal = 0
20570         .purchaseAssetNo.SetFocus
20580       End With
20590     End Select
20600   End If

EXITP:
20610   KeyCode = intRetVal
20620   Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub purchaseShareFace_Enter()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseShareFace_Enter"

20710   With Me
20720     .purchaseShareFace.Format = "General Number"
20730   End With

EXITP:
20740   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseShareFace_KeyDown(KeyCode As Integer, Shift As Integer)

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseShareFace_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
20870       With Me
20880         intRetVal = 0
20890         .purchaseAssetDate.SetFocus
20900       End With
20910     End Select
20920   End If

        ' ** Shift keys.
20930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20950       With Me
20960         intRetVal = 0
20970         .purchaseAssetNo.SetFocus
20980       End With
20990     End Select
21000   End If

        ' ** Ctrl-Shift keys.
21010   If intCtrlDown And (Not intAltDown) And intShiftDown Then
21020     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21030       With Me
21040         intRetVal = 0
21050         .cmdPurchaseAssetNew.SetFocus
21060       End With
21070     End Select
21080   End If

EXITP:
21090   KeyCode = intRetVal
21100   Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub purchaseShareFace_Exit(Cancel As Integer)

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseShareFace_Exit"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

21210   blnContinue = True

21220   With Me

21230     If VerifyAccountNo = False Then  ' ** Function: Below.
21240       blnContinue = False
21250       If gblnDev_Debug Then
21260         subNullShareFace  ' ** Procedure: Below.
21270       Else
21280         .purchaseShareFace = vbNullString  ' ** Need in case share/face is too large.
21290       End If
21300     Else

21310       If IsNumeric(.purchaseShareFace.text) = True Then
21320         If CDbl(.purchaseShareFace.text) > 999999999.9999 Then
21330           blnContinue = False
21340           MsgBox "Share/Face is too large.", vbInformation + vbOKOnly, "Invalid Entry"
21350           If gblnDev_Debug Then
21360             subNullShareFace  ' ** Procedure: Below.
21370           Else
21380             .purchaseShareFace = vbNullString
21390           End If
21400           .purchaseShareFace = 0
21410           Cancel = -1
21420           .purchaseShareFace.SetFocus
21430         Else
21440           If .purchaseShareFace < 0 Then
21450             blnContinue = False
21460             MsgBox "You cannot enter a negative Share/Face.", vbInformation + vbOKOnly, "Invalid Entry"
21470             .purchaseShareFace = CStr(0)
21480             .Refresh
21490             Cancel = -1
21500             .purchaseShareFace.SetFocus
21510           End If
21520         End If
21530       Else
21540         blnContinue = False
21550         MsgBox "Share/Face must have a value.", vbInformation + vbOKOnly, "Entry Required"
21560         If gblnDev_Debug Then
21570           subNullShareFace  ' ** Procedure: Below.
21580         Else
21590 On Error Resume Next
21600           .purchaseShareFace = vbNullString
21610           If ERR.Number <> 0 Then
                  ' ** 2473  The expression |2 you entered as the event property setting produced the following error: |1.
                  ' **       * The expression may not result in the name of a macro, the name of a user-defined function, or [Event Procedure].
                  ' **       * There may have been an error evaluating the function, event, or macro.
21620             .purchaseShareFace = Null
21630 On Error GoTo ERRH
21640           Else
21650 On Error GoTo ERRH
21660           End If
21670         End If
21680         DoCmd.RunCommand acCmdSaveRecord
21690         Cancel = -1
21700         .purchaseShareFace.SetFocus
21710       End If

21720       If blnContinue = True Then
21730         intPos01 = InStr(1, .purchaseShareFace, ".", vbTextCompare)
21740         If intPos01 > 0 Then
21750           If Len(.purchaseShareFace) - intPos01 > gintShareFaceDecimals Then
21760             blnContinue = False
21770             MsgBox "You can only enter Share/Face with up to " & Trim(str(gintShareFaceDecimals)) & " decimals.", _
                    vbInformation + vbOKOnly, "Invalid Entry"
21780             If gblnDev_Debug Then
21790               subNullShareFace  ' ** Procedure: Below.
21800             Else
21810               .purchaseShareFace = vbNullString
21820             End If
21830             Cancel = -1
21840             .purchaseShareFace.SetFocus
21850           Else
21860             .purchaseShareFace.Format = "#,###." & String(gintShareFaceDecimals, "0")
21870           End If
21880         Else
21890           .purchaseShareFace.Format = "#,###"
21900         End If
21910       End If

21920       If blnContinue = True Then

21930         gstrPurchaseShareFace = .purchaseShareFace.text

21940         If .purchaseShareFace <> .purchaseShareFace.OldValue Or IsNull(.purchaseShareFace.OldValue) Then
21950           Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
21960         End If

21970       End If
21980     End If

21990   End With

EXITP:
22000   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseShareFace_LostFocus()

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseShareFace_LostFocus"

22110   VerifyAccountNo  ' ** Function: Below.

EXITP:
22120   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseAssetDate_Enter()

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetDate_Enter"

        Dim strTmp01 As String, dblTmp02 As Double

22210   With Me
22220     If IsNull(.assetdate) = True Then
22230       datAssetDate_OldValue = 0
22240       lngAssetDate_OldValue = 0&
22250       .purchaseAssetDate = Null
22260     Else
22270       If IsDate(.assetdate) = False Then
22280         datAssetDate_OldValue = 0
22290         lngAssetDate_OldValue = 0&
22300         .assetdate = Null
22310         .purchaseAssetDate = Null
22320       Else
22330         datAssetDate_OldValue = .assetdate
22340         dblTmp02 = CDbl(datAssetDate_OldValue)
22350         strTmp01 = CStr(dblTmp02)
22360         If InStr(strTmp01, ".") > 0 Then
22370           strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
22380         Else
22390           datAssetDate_OldValue = datAssetDate_OldValue + time()
22400         End If
22410         lngAssetDate_OldValue = CLng(strTmp01)
22420         .purchaseAssetDate.text = Format(CDate(lngAssetDate_OldValue), "mm/dd/yyyy")  ' ** Display without timestamp
22430       End If
22440     End If
22450     blnAssetDateChecked = False
22460   End With

EXITP:
22470   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseAssetDate_BeforeUpdate(Cancel As Integer)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetDate_BeforeUpdate"

22510   With Me
22520     If .purchaseAssetDate.text <> "__/__/____" And .purchaseAssetDate.text <> vbNullString Then
22530       If Left(.purchaseAssetDate.text, 2) > 12 Then
22540         MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Date"
22550         Cancel = -1
22560         .purchaseAssetDate.Undo
22570       End If
22580     End If
22590   End With

EXITP:
22600   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseAssetDate_AfterUpdate()

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetDate_AfterUpdate"

        Dim strTmp01 As String, lngTmp02 As Long, dblTmp03 As Double, datTmp04 As Date

22710   With Me
22720     If IsNull(.purchaseAssetDate) = True Then
22730       datAssetDate_OldValue = Now()
22740       lngAssetDate_OldValue = CLng(Date)
22750       .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
22760       .purchaseAssetDate = CDate(lngAssetDate_OldValue)  ' ** Display without timestamp.
22770     Else
22780       If Trim(.purchaseAssetDate) = vbNullString Then
22790         datAssetDate_OldValue = Now()
22800         lngAssetDate_OldValue = CLng(Date)
22810         .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
22820         .purchaseAssetDate = CDate(lngAssetDate_OldValue)  ' ** Display without timestamp.
22830       Else
22840         If DateCheck_Trade(.purchaseAssetDate.text) = True Then  ' ** Module Function: modUtilities.
22850           If Right(.purchaseAssetDate.text, 1) = "_" Then
22860             strTmp01 = .purchaseAssetDate.text
22870             Do While Right(strTmp01, 1) = "_"
22880               strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
22890             Loop
22900           Else
22910             strTmp01 = .purchaseAssetDate.text
22920           End If
22930           datTmp04 = CDate(strTmp01)
22940           dblTmp03 = CDbl(datTmp04)
22950           strTmp01 = CStr(dblTmp03)
22960           If InStr(strTmp01, ".") > 0 Then  ' ** There shouldn't be.
22970             strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
22980           End If
22990           lngTmp02 = CLng(strTmp01)
23000           If CDbl(datAssetDate_OldValue) = 0# Then
23010             Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
23020             datAssetDate_OldValue = datTmp04 + time()
23030             lngAssetDate_OldValue = lngTmp02
23040             .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
23050           Else
23060             If lngTmp02 <> lngAssetDate_OldValue Then
23070               Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
23080               datAssetDate_OldValue = datTmp04 + time()
23090               lngAssetDate_OldValue = lngTmp02
23100               .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
23110             Else
                    ' ** Let assetdate stand.
23120               .assetdate = datTmp04 + time()
23130             End If
23140           End If
23150         Else
23160           Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
23170           datAssetDate_OldValue = Now()
23180           lngAssetDate_OldValue = CLng(Date)
23190           .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
23200           .purchaseAssetDate = CDate(lngAssetDate_OldValue)  ' ** Display without timestamp.
23210         End If
23220       End If
23230     End If
23240     blnAssetDateChecked = True
23250   End With

EXITP:
23260   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseAssetDate_KeyDown(KeyCode As Integer, Shift As Integer)

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23320   intShiftDown = (Shift And acShiftMask) > 0
23330   intAltDown = (Shift And acAltMask) > 0
23340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23370       With Me
23380         intRetVal = 0
23390         .cmbLocations.SetFocus
23400       End With
23410     End Select
23420   End If

        ' ** Shift keys.
23430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23450       With Me
23460         intRetVal = 0
23470         .purchaseShareFace.SetFocus
23480       End With
23490     End Select
23500   End If

        ' ** Ctrl keys.
23510   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
23520     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23530       With Me
23540         intRetVal = 0
23550         .cmdCalendar2.SetFocus
23560       End With
23570     End Select
23580   End If

EXITP:
23590   KeyCode = intRetVal
23600   Exit Sub

ERRH:
410     intRetVal = 0
420     Select Case ERR.Number
        Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub purchaseAssetDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetDate_MouseMove"

23710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
23720   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseAssetDate_Exit(Cancel As Integer)

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetDate_Exit"

23810   With Me
23820     If blnAssetDateChecked = False Then
23830       If DateCheck_Trade(.purchaseAssetDate.text) = True Then  ' ** Module Function: modUtilities.
              ' ** '.Text' needed to reliably check date.
23840         If IsNull(.purchaseAssetDate.OldValue) Then
23850           Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
23860         Else
23870           If .purchaseAssetDate <> .purchaseAssetDate.OldValue Then
23880             Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
23890           End If
23900         End If
23910       Else
23920         If IsNull(.purchaseAssetDate) = True Then
23930           datAssetDate_OldValue = Now()
23940           lngAssetDate_OldValue = CLng(Date)
23950           .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
23960           .purchaseAssetDate = CDate(lngAssetDate_OldValue)  ' ** Display without timestamp.
23970         Else
23980           If Trim(.purchaseAssetDate) = vbNullString Then
23990             datAssetDate_OldValue = Now()
24000             lngAssetDate_OldValue = CLng(Date)
24010             .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
24020             .purchaseAssetDate = CDate(lngAssetDate_OldValue)  ' ** Display without timestamp.
24030           Else
24040             Cancel = -1
24050             .purchaseAssetDate.SetFocus
24060           End If
24070         End If
24080       End If
24090     Else
24100       blnAssetDateChecked = False
24110     End If
24120   End With

EXITP:
24130   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseAssetDate_LostFocus()

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseAssetDate_LostFocus"

24210   VerifyAccountNo  ' ** Function: Below.

EXITP:
24220   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

24310   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
24320   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

24410   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
24420   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Public Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

24510   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
24520   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
420     Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

24610   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
24620   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

24710   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
24720   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

24810   Calendar_Handler_Sub3 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub3PurchFuncs.

EXITP:
24820   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbLocations_Enter()

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_Enter"

24910   With Me
24920     If IsNull(.cmbLocations) = True Then
24930       .[Location_ID] = 1&
24940     Else
24950       If .cmbLocations = 0& Then
24960         .[Location_ID] = 1&
24970       End If
24980     End If
24990   End With

EXITP:
25000   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub cmbLocations_GotFocus()

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_GotFocus"

25110   With Me
25120     If .[Location_ID] = 1& Then
25130       ExpandCombo .cmbLocations  ' ** Module Procedure: modUtilities.
25140     End If
25150   End With

EXITP:
25160   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Public Sub cmbLocations_AfterUpdate()

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_AfterUpdate"

25210   With Me
25220     If IsNull(.cmbLocations) = True Then
25230       .cmbLocations = 1  ' ** Unassigned.
25240       .[Location_ID] = 1
25250     Else
25260       .[Location_ID] = .cmbLocations
25270     End If
25280   End With

EXITP:
25290   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbLocations_KeyDown(KeyCode As Integer, Shift As Integer)

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_KeyDown"

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
25390         If .purchaseICash.Enabled = True Then
25400           .purchaseICash.SetFocus
25410         ElseIf .purchasePCash.Enabled = True Then
25420           .purchasePCash.SetFocus
25430         Else
25440           .purchaseCost.SetFocus
25450         End If
25460       End With
25470     End Select
25480   End If

        ' ** Shift keys.
25490   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25500     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25510       With Me
25520         intRetVal = 0
25530         .purchaseAssetDate.SetFocus
25540       End With
25550     End Select
25560   End If

        ' ** Ctrl keys.
25570   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
25580     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25590       With Me
25600         intRetVal = 0
25610         .cmdPurchaseLocNew.SetFocus
25620       End With
25630     End Select
25640   End If

        ' ** Ctrl-Shift keys.
25650   If intCtrlDown And (Not intAltDown) And intShiftDown Then
25660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25670       With Me
25680         intRetVal = 0
25690         .cmdCalendar2.SetFocus
25700       End With
25710     End Select
25720   End If

EXITP:
25730   KeyCode = intRetVal
25740   Exit Sub

ERRH:
410     intRetVal = 0
420     Select Case ERR.Number
        Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub cmbLocations_Exit(Cancel As Integer)

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_Exit"

25810   VerifyAccountNo  ' ** Function: Below.

EXITP:
25820   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbLocations_LostFocus()

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_LostFocus"

25910   With Me
25920     If .cmbLocations <> .cmbLocations.OldValue Or IsNull(.cmbLocations.OldValue) Then
25930       Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
25940     End If
25950   End With

EXITP:
25960   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub cmdPurchaseLocNew_Click()

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPurchaseLocNew_Click"

        Dim strDocName As String

26010   strDocName = "frmLocations_Add_Purchase"
26020   DoCmd.OpenForm strDocName, acNormal, , , acFormAdd, acDialog, "frmJournal"

EXITP:
26030   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmdPurchaseLocNew_KeyDown(KeyCode As Integer, Shift As Integer)
' ** See Note, above.

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPurchaseLocNew_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26120   intShiftDown = (Shift And acShiftMask) > 0
26130   intAltDown = (Shift And acAltMask) > 0
26140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26160     Select Case intRetVal
          Case vbKeyTab
26170       With Me
26180         intRetVal = 0
26190         If .purchaseICash.Enabled = True Then
26200           .purchaseICash.SetFocus
26210         ElseIf .purchasePCash.Enabled = True Then
26220           .purchasePCash.SetFocus
26230         Else
26240           .purchaseCost.SetFocus
26250         End If
26260       End With
26270     End Select
26280   End If

        ' ** Shift keys.
26290   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26300     Select Case intRetVal
          Case vbKeyTab
26310       With Me
26320         intRetVal = 0
26330         .cmbLocations.SetFocus
26340       End With
26350     End Select
26360   End If

EXITP:
26370   KeyCode = intRetVal
26380   Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Public Sub purchaseICash_AfterUpdate()

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseICash_AfterUpdate"

        Dim dblICash As Double
        Dim varTmp00 As Variant

26410   With Me
26420     If .purchaseCurr_ID.Visible = True Then
26430       .purchaseICash_usd = Null
26440       If .purchaseICash_usd.Visible = False Then .purchaseICash_usd.Visible = True
26450       If IsNull(.purchaseICash) = False Then
26460         If .purchaseICash <> 0 Then
26470           If IsNull(.purchaseCurr_ID.Column(CBX_C_CURRID)) = False Then
26480             lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)
26490             If lngCurrID <> 150& Then  ' ** USD.
26500               varTmp00 = .purchaseCurr_ID.Column(CBX_C_RATE2)
26510               If IsNull(varTmp00) = False Then
26520                 If varTmp00 > 0 Then
26530                   dblICash = .purchaseICash
26540                   dblICash = (dblICash * varTmp00)
26550                   dblICash = Round(dblICash, 2)
26560                   .purchaseICash_usd = dblICash
26570                   If IsNull(.purchasePCash) = False Then
26580                     If .purchasePCash = 0 Then
26590                       If IsNull(.purchasePCash_usd) = True Then
26600                         purchasePCash_AfterUpdate  ' ** Procedure: Below.
26610                       End If
26620                     End If
26630                   End If
26640                   If IsNull(.purchaseCost) = False Then
26650                     If .purchaseCost = 0 Then
26660                       If IsNull(.purchaseCost_usd) = True Then
26670                         purchaseCost_AfterUpdate  ' ** Procedure: Below.
26680                       End If
26690                     End If
26700                   End If
26710                 End If
26720               End If
26730             End If
26740           End If
26750         Else
26760           .purchaseICash_usd = 0#
26770         End If
26780       End If
26790     End If
26800     DoEvents
26810   End With

EXITP:
26820   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseICash_KeyDown(KeyCode As Integer, Shift As Integer)

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseICash_KeyDown"

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
          Case vbKeyTab
26970       With Me
26980         intRetVal = 0
26990         If .purchasePCash.Enabled = True Then
27000           .purchasePCash.SetFocus
27010         Else
27020           .purchaseCost.SetFocus
27030         End If
27040       End With
27050     End Select
27060   End If

        ' ** Shift keys.
27070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27080     Select Case intRetVal
          Case vbKeyTab
27090       With Me
27100         intRetVal = 0
27110         .cmbLocations.SetFocus
27120       End With
27130     End Select
27140   End If

        ' ** Ctrl-Shift keys.
27150   If intCtrlDown And (Not intAltDown) And intShiftDown Then
27160     Select Case intRetVal
          Case vbKeyTab
27170       With Me
27180         intRetVal = 0
27190         .cmdPurchaseLocNew.SetFocus
27200       End With
27210     End Select
27220   End If

EXITP:
27230   KeyCode = intRetVal
27240   Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub purchaseICash_Exit(Cancel As Integer)

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseICash_Exit"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

27310   With Me

27320     blnContinue = True

27330     lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)

27340     If IsNumeric(Rem_Dollar(.purchaseICash, lngCurrID)) = True Then  ' ** Module Function: modStringFuncs.
27350       If CDbl(Rem_Dollar(.purchaseICash.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
27360         blnContinue = False
27370         MsgBox "Accrued/Income Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
27380         .purchaseICash = vbNullString
27390         .purchaseICash_usd = Null
27400         Cancel = -1
27410         .purchaseICash.SetFocus
27420       End If
27430     Else
27440       blnContinue = False
27450       MsgBox "Accrued/Income Cash must have a value.", vbInformation + vbOKOnly, "Entry Required"
27460       .purchaseICash = vbNullString
27470       .purchaseICash_usd = Null
27480       Cancel = -1
27490       .purchaseICash.SetFocus
27500     End If

27510     If blnContinue = True Then
27520       intPos01 = InStr(CStr(Val(CStr(.purchaseICash))), ".")
27530       If intPos01 > 0 Then
27540         If Len(.purchaseICash) - intPos01 > 2 Then
27550           blnContinue = False
27560           MsgBox "You can only enter Accrued/Income Cash with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
27570           .purchaseICash = 0
27580           .purchaseICash_usd = Null
27590           Cancel = -1
27600           .purchaseICash.SetFocus
27610         End If
27620       End If
27630     End If

27640     If blnContinue = True Then
27650       Sub3Purch_TaxCode Me  ' ** Module Procedure: modPurchaseSold.
27660     End If

27670     If blnContinue = True Then
27680       gstrPurchaseICash = .purchaseICash.text
27690       gstrPurchaseICash = Rem_Dollar(gstrPurchaseICash, lngCurrID)  ' ** Module Function: modStringFuncs.
27700       If .purchaseICash <> .purchaseICash.OldValue Or IsNull(.purchaseICash.OldValue) Then
27710         Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
27720       End If
27730     End If

27740   End With

EXITP:
27750   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseICash_LostFocus()

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseICash_LostFocus"

27810   If VerifyAccountNo = True Then  ' ** Function: Below.

27820     With Me

27830       If .purchaseICash > 0 Then
27840         .purchaseICash = .purchaseICash * -1
27850         purchaseICash_AfterUpdate  ' ** Procedure: Above.
27860       End If

27870       If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
27880         .cmbTaxCodes.Enabled = True
27890         .cmbTaxCodes.BorderColor = CLR_LTBLU2
27900         .cmbTaxCodes.BackStyle = acBackStyleNormal
27910         .cmbTaxCodes_lbl.BackStyle = acBackStyleNormal
27920         .cmbTaxCodes_lbl_box.Visible = False
27930         .cmbRevenueCodes.Enabled = True
27940         .cmbRevenueCodes.BorderColor = CLR_LTBLU2
27950         .cmbRevenueCodes.BackStyle = acBackStyleNormal
27960         .cmbRevenueCodes_lbl.BackStyle = acBackStyleNormal
27970         .cmbRevenueCodes_lbl_box.Visible = False
27980       ElseIf .purchaseICash <> 0 Then
27990         .cmbTaxCodes.Enabled = True
28000         .cmbTaxCodes.BorderColor = CLR_LTBLU2
28010         .cmbTaxCodes.BackStyle = acBackStyleNormal
28020         .cmbTaxCodes_lbl.BackStyle = acBackStyleNormal
28030         .cmbTaxCodes_lbl_box.Visible = False
28040         .cmbRevenueCodes.Enabled = True
28050         .cmbRevenueCodes.BorderColor = CLR_LTBLU2
28060         .cmbRevenueCodes.BackStyle = acBackStyleNormal
28070         .cmbRevenueCodes_lbl.BackStyle = acBackStyleNormal
28080         .cmbRevenueCodes_lbl_box.Visible = False
28090       Else
28100         .cmbTaxCodes.Enabled = False
28110         .cmbTaxCodes.BorderColor = WIN_CLR_DISR
28120         .cmbTaxCodes.BackStyle = acBackStyleTransparent
28130         .cmbTaxCodes_lbl.BackStyle = acBackStyleTransparent
28140         .cmbTaxCodes_lbl_box.Visible = True
28150         .cmbRevenueCodes.Enabled = False
28160         .cmbRevenueCodes.BorderColor = WIN_CLR_DISR
28170         .cmbRevenueCodes.BackStyle = acBackStyleTransparent
28180         .cmbRevenueCodes_lbl.BackStyle = acBackStyleTransparent
28190         .cmbRevenueCodes_lbl_box.Visible = True
28200       End If

28210       lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)

28220       If IsNumeric(.purchaseICash) = True Then
              ' ** This covers any sign changes, assures 0 = "0" not "0.00", and won't get stung if Null.
28230         gstrPurchaseICash = CStr(Val(CStr(.purchaseICash)))
28240         gstrPurchaseICash = Rem_Dollar(gstrPurchaseICash, lngCurrID)  ' ** Module Function: modStringFuncs.
28250       End If

28260     End With

28270   End If

EXITP:
28280   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Public Sub purchasePCash_AfterUpdate()

28300 On Error GoTo ERRH

        Const THIS_PROC As String = "purchasePCash_AfterUpdate"

        Dim dblPCash As Double
        Dim varTmp00 As Variant

28310   With Me
28320     If .purchaseCurr_ID.Visible = True Then
28330       .purchasePCash_usd = Null
28340       If .purchasePCash_usd.Visible = False Then .purchasePCash_usd.Visible = True
28350       If IsNull(.purchasePCash) = False Then
28360         If .purchasePCash <> 0 Then
28370           If IsNull(.purchaseCurr_ID.Column(CBX_C_CURRID)) = False Then
28380             lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)
28390             If lngCurrID <> 150& Then  ' ** USD.
28400               varTmp00 = .purchaseCurr_ID.Column(CBX_C_RATE2)
28410               If IsNull(varTmp00) = False Then
28420                 If varTmp00 > 0 Then
28430                   dblPCash = .purchasePCash
28440                   dblPCash = (dblPCash * varTmp00)
28450                   dblPCash = Round(dblPCash, 2)
28460                   .purchasePCash_usd = dblPCash
28470                   If IsNull(.purchaseICash) = False Then
28480                     If .purchaseICash = 0 Then
28490                       If IsNull(.purchaseICash_usd) = True Then
28500                         purchaseICash_AfterUpdate  ' ** Procedure: Below.
28510                       End If
28520                     End If
28530                   End If
28540                   If IsNull(.purchaseCost) = False Then
28550                     If .purchaseCost = 0 Then
28560                       If IsNull(.purchaseCost_usd) = True Then
28570                         purchaseCost_AfterUpdate  ' ** Procedure: Below.
28580                       End If
28590                     End If
28600                   End If
28610                 End If
28620               End If
28630             End If
28640           End If
28650         Else
28660           .purchasePCash_usd = 0#
28670         End If
28680       End If
28690     End If
28700     DoEvents
28710   End With

EXITP:
28720   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub purchasePCash_KeyDown(KeyCode As Integer, Shift As Integer)

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "purchasePCash_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28820   intShiftDown = (Shift And acShiftMask) > 0
28830   intAltDown = (Shift And acAltMask) > 0
28840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28860     Select Case intRetVal
          Case vbKeyTab
28870       With Me
28880         intRetVal = 0
28890         .purchaseCost.SetFocus
28900       End With
28910     End Select
28920   End If

        ' ** Shift keys.
28930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28940     Select Case intRetVal
          Case vbKeyTab
28950       With Me
28960         intRetVal = 0
28970         If .purchaseICash.Enabled = True Then
28980           .purchaseICash.SetFocus
28990         Else
29000           .cmbLocations.SetFocus
29010         End If
29020       End With
29030     End Select
29040   End If

        ' ** Ctrl-Shift keys.
29050   If intCtrlDown And (Not intAltDown) And intShiftDown Then
29060     Select Case intRetVal
          Case vbKeyTab
29070       With Me
29080         intRetVal = 0
29090         If .purchaseICash.Enabled = False Then
29100           .cmdPurchaseLocNew.SetFocus
29110         End If
29120       End With
29130     End Select
29140   End If

EXITP:
29150   KeyCode = intRetVal
29160   Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub purchasePCash_Exit(Cancel As Integer)

29200 On Error GoTo ERRH

        Const THIS_PROC As String = "purchasePCash_Exit"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

29210   With Me

29220     blnContinue = True

29230     lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)

29240     If .purchaseICash * -1 = .purchaseCost And .purchaseCost > 0 And .purchasePCash = 0 Then
            ' ** No further checking if this is a purchase from income cash.
29250       blnContinue = False
29260       gstrPurchasePCash = .purchasePCash.text
29270       gstrPurchasePCash = Rem_Dollar(gstrPurchasePCash, lngCurrID)  ' ** Module Function: modStringFuncs.
29280     Else
29290       If IsNumeric(.purchasePCash) = True Then
29300         If CDbl(Rem_Dollar(.purchasePCash.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
29310           blnContinue = False
29320           MsgBox "Principal Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
29330           .purchasePCash = vbNullString
29340           .purchasePCash_usd = Null
29350           Cancel = -1
29360           .purchasePCash.SetFocus
29370         Else
29380           If .purchaseType = "Deposit" And CDbl(Rem_Dollar(.purchasePCash.text, lngCurrID)) <> 0 Then  ' ** Module Function: modStringFuncs.
                  ' ** For a 'Deposit', never pcash nor icash, and sometimes not even cost.
29390             blnContinue = False
29400             MsgBox "Deposit Principal Cash must be zero.", vbInformation + vbOKOnly, "Invalid Entry"
29410             .purchasePCash = vbNullString
29420             .purchasePCash_usd = Null
29430             Cancel = -1
29440             .purchasePCash.SetFocus
29450           Else
29460             If .purchaseType = "Purchase" And IsNull(.purchaseCost) = False Then
                    ' ** For a 'Purchase', always pcash or icash, and always cost.
29470               If IsNumeric(.purchaseCost) = True Then
29480                 If .purchaseCost > 0 And .purchasePCash > 0 Then
29490                   .purchasePCash = (.purchasePCash * -1#)
29500                   purchasePCash_AfterUpdate  ' ** Procedure: Above.
29510                 End If
29520               End If
29530             End If
29540           End If
29550         End If
29560       Else
29570         blnContinue = False
29580         MsgBox "Principal Cash must have a value.", vbInformation + vbOKOnly, "Entry Required"
              ' ** Even if it's 0.
29590         .purchasePCash = vbNullString
29600         .purchasePCash_usd = Null
29610         Cancel = -1
29620         .purchasePCash.SetFocus
29630       End If
29640     End If

29650     If blnContinue = True Then
29660       intPos01 = InStr(CStr(Val(CStr(.purchasePCash))), ".")
29670       If intPos01 > 0 Then
29680         If Len(.purchasePCash) - intPos01 > 2 Then
29690           blnContinue = False
29700           MsgBox "You can only enter Principal Cash with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
29710           .purchasePCash = vbNullString
29720           .purchasePCash_usd = Null
29730           Cancel = -1
29740           .purchasePCash.SetFocus
29750         End If
29760       End If
29770     End If

29780     If blnContinue = True Then
29790       If .purchaseICash <> 0 And .purchasePCash <> 0 Then
29800         Select Case IsNull(.purchaseAssetNo.Column(CBX_A_D4D))
              Case True
29810           blnContinue = False
29820           MsgBox "An asset must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
29830           .purchasePCash = 0@
29840           .purchasePCash_usd = Null
29850           .purchaseAssetNo.SetFocus
29860         Case False
29870           If .purchaseAssetNo.Column(CBX_A_D4D) = True Then
29880             blnContinue = False
29890             MsgBox "A dollar-for-dollar asset may not be purchased" & vbCrLf & _
                    "with both Income Cash and Principal Cash.", vbInformation + vbOKOnly, "Invalid Entry"
29900             .purchasePCash = 0@
29910             .purchasePCash_usd = Null
29920             Cancel = -1
29930             .purchasePCash.SetFocus
29940           End If
29950         End Select
29960       End If
29970     End If

29980     If blnContinue = True Then
29990       Sub3Purch_TaxCode Me  ' ** Module Procedure: modPurchaseSold.
30000     End If

30010     If blnContinue = True Then
30020       gstrPurchasePCash = .purchasePCash.text
30030       gstrPurchasePCash = Rem_Dollar(gstrPurchasePCash, lngCurrID)  ' ** Module Function: modStringFuncs.
30040       If .purchasePCash <> .purchasePCash.OldValue Then
30050         Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
30060       End If
30070     End If

30080   End With
30090   blnPCash_Exit = True

EXITP:
30100   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchasePCash_LostFocus()

30200 On Error GoTo ERRH

        Const THIS_PROC As String = "purchasePCash_LostFocus"

30210   blnPCash_LostFocus = True

30220   If VerifyAccountNo = True Then  ' ** Function: Below.

30230     With Me
30240       If .purchaseType = "Purchase" Then
30250         If .purchasePCash > 0 Then
30260           .purchaseCost = .purchasePCash
30270           purchaseCost_AfterUpdate  ' ** Procedure: Below.
30280           .purchasePCash = .purchasePCash * -1
30290           purchasePCash_AfterUpdate  ' ** Procedure: Above.
30300         ElseIf .purchasePCash < 0 Then
30310           .purchaseCost = .purchasePCash * -1
30320           purchaseCost_AfterUpdate  ' ** Procedure: Below.
30330         ElseIf .purchasePCash = 0 Then
                ' ** Do nothing if this is a purchase from income cash.
30340         End If
30350       Else
30360         If .purchaseType = "Liability" Then
30370           If .purchasePCash < 0 Then
30380             .purchaseCost = .purchasePCash
30390             purchaseCost_AfterUpdate  ' ** Procedure: Below.
30400             .purchasePCash = .purchasePCash * -1
30410             purchasePCash_AfterUpdate  ' ** Procedure: Above.
30420           Else
30430             .purchaseCost = .purchasePCash * -1
30440             purchaseCost_AfterUpdate  ' ** Procedure: Below.
30450           End If
30460         End If
30470       End If

30480       lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)

30490       If IsNumeric(.purchasePCash) = True Then
              ' ** This covers any sign changes, assures 0 = "0" not "0.00", and won't get stung if Null.
30500         gstrPurchasePCash = CStr(Val(CStr(.purchasePCash)))
30510         gstrPurchasePCash = Rem_Dollar(gstrPurchasePCash, lngCurrID)  ' ** Module Function: modStringFuncs.
30520       End If

30530     End With

30540   End If

EXITP:
30550   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseCost_Enter()
' ** This covers the bizarre sequence of events that when a new Location is added,
' ** it somehow skips the purchasePCash_LostFocus() event!

30600 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseCost_Enter"

30610   If blnPCash_Exit = True And blnPCash_LostFocus = False Then
30620     purchasePCash_LostFocus  ' ** Procedure: Above.
30630   End If
30640   blnPCash_Exit = False: blnPCash_LostFocus = False

EXITP:
30650   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Public Sub purchaseCost_AfterUpdate()

30700 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseCost_AfterUpdate"

        Dim dblCost As Double
        Dim varTmp00 As Variant

30710   With Me
30720     If .purchaseCurr_ID.Visible = True Then
30730       .purchaseCost_usd = Null
30740       If .purchaseCost_usd.Visible = False Then .purchaseCost_usd.Visible = True
30750       If IsNull(.purchaseCost) = False Then
30760         If .purchaseCost <> 0 Then
30770           If IsNull(.purchaseCurr_ID.Column(CBX_C_CURRID)) = False Then
30780             lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)
30790             If lngCurrID <> 150& Then  ' ** USD.
30800               varTmp00 = .purchaseCurr_ID.Column(CBX_C_RATE2)
30810               If IsNull(varTmp00) = False Then
30820                 If varTmp00 > 0 Then
30830                   dblCost = .purchaseCost
30840                   dblCost = (dblCost * varTmp00)
30850                   dblCost = Round(dblCost, 2)
30860                   .purchaseCost_usd = dblCost
30870                   If IsNull(.purchaseICash) = False Then
30880                     If .purchaseICash = 0 Then
30890                       If IsNull(.purchaseICash_usd) = True Then
30900                         purchaseICash_AfterUpdate  ' ** Procedure: Above.
30910                       End If
30920                     End If
30930                   End If
30940                   If IsNull(.purchasePCash) = False Then
30950                     If .purchasePCash = 0 Then
30960                       If IsNull(.purchasePCash_usd) = True Then
30970                         purchasePCash_AfterUpdate  ' ** Procedure: Above.
30980                       End If
30990                     End If
31000                   End If
31010                 End If
31020               End If
31030             End If
31040           End If
31050         Else
31060           .purchaseCost_usd = 0#
31070         End If
31080       End If
31090     End If
31100     DoEvents
31110   End With

EXITP:
31120   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseCost_KeyDown(KeyCode As Integer, Shift As Integer)

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseCost_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
31220   intShiftDown = (Shift And acShiftMask) > 0
31230   intAltDown = (Shift And acAltMask) > 0
31240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31260     Select Case intRetVal
          Case vbKeyTab
31270       With Me
31280         intRetVal = 0
31290         .purchaseDescription.SetFocus
31300       End With
31310     End Select
31320   End If

        ' ** Shift keys.
31330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31340     Select Case intRetVal
          Case vbKeyTab
31350       With Me
31360         intRetVal = 0
31370         If .purchasePCash.Enabled = True Then
31380           .purchasePCash.SetFocus
31390         ElseIf .purchaseICash.Enabled = True Then
31400           .purchaseICash.SetFocus
31410         Else
31420           .cmbLocations.SetFocus
31430         End If
31440       End With
31450     End Select
31460   End If

        ' ** Ctrl-Shift keys.
31470   If intCtrlDown And (Not intAltDown) And intShiftDown Then
31480     Select Case intRetVal
          Case vbKeyTab
31490       With Me
31500         intRetVal = 0
31510         If .purchaseICash.Enabled = False And .purchasePCash.Enabled = False Then
31520           .cmdPurchaseLocNew.SetFocus
31530         End If
31540       End With
31550     End Select
31560   End If

EXITP:
31570   KeyCode = intRetVal
31580   Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub purchaseCost_Exit(Cancel As Integer)

31600 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseCost_Exit"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

31610   With Me

31620     blnContinue = True

31630     lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)

31640     If IsNumeric(.purchaseCost) = True Then
31650       If CDbl(Rem_Dollar(.purchaseCost.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
31660         blnContinue = False
31670         MsgBox "Cost is too large.", vbInformation + vbOKOnly, "Invalid Entry"
31680         .purchaseCost = vbNullString
31690         .purchaseCost_usd = Null
31700         Cancel = -1
31710         .purchaseCost.SetFocus
31720       End If
31730     Else
31740       blnContinue = False
31750       MsgBox "Cost must have a value.", vbInformation + vbOKOnly, "Entry Required"
            ' ** Even if it's 0.
31760       .purchaseCost = vbNullString
31770       .purchaseCost_usd = Null
31780       Cancel = -1
31790       .purchaseCost.SetFocus
31800     End If

31810     If blnContinue = True Then

31820       If .purchaseType = "Liability" And CDbl(Rem_Dollar(.purchaseCost.text, lngCurrID)) > 0 Then  ' ** Module Function: modStringFuncs.
31830         .purchaseCost.text = (CDbl(Rem_Dollar(.purchaseCost.text, lngCurrID)) * -1#)  ' ** Module Function: modStringFuncs.
31840       End If

31850       If .purchaseType = "Liability" And CDbl(Rem_Dollar(.purchaseCost.text, lngCurrID)) > 0 Then  ' ** Module Function: modStringFuncs.
31860         blnContinue = False
31870         MsgBox "Cost must be less than zero.", vbInformation + vbOKOnly, "Invalid Entry"
31880         .purchaseCost = vbNullString
31890         .purchaseCost_usd = Null
31900         Cancel = -1
31910         .purchaseCost.SetFocus
31920       Else
31930         If .purchaseType = "Deposit" And CDbl(Rem_Dollar(.purchaseCost.text, lngCurrID)) < 0 Then  ' ** Module Function: modStringFuncs.
31940           blnContinue = False
31950           MsgBox "Cost must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
31960           .purchaseCost = vbNullString
31970           .purchaseCost_usd = Null
31980           Cancel = -1
31990           .purchaseCost.SetFocus
32000         End If
32010       End If

32020     End If

32030     If blnContinue = True Then
32040       intPos01 = InStr(CStr(Val(CStr(.purchaseCost))), ".")
32050       If intPos01 > 0 Then
32060         If Len(.purchaseCost) - intPos01 > 2 Then
32070           blnContinue = False
32080           MsgBox "You can only enter Cost with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
32090           .purchaseCost = vbNullString
32100           .purchaseCost_usd = Null
32110           Cancel = -1
32120           .purchaseCost.SetFocus
32130         End If
32140       End If
32150     End If

32160     If blnContinue = True Then
32170       Sub3Purch_TaxCode Me  ' ** Module Procedure: modPurchaseSold.
32180     End If

32190     lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)

32200     If blnContinue = True Then
32210       gstrPurchaseCost = .purchaseCost.text
32220       gstrPurchaseCost = Rem_Dollar(gstrPurchaseCost, lngCurrID)  ' ** Module Function: modStringFuncs.
32230       If .purchaseCost <> .purchaseCost.OldValue Then
32240         Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
32250       End If
32260     End If

32270   End With

EXITP:
32280   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseCost_LostFocus()

32300 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseCost_LostFocus"

32310   With Me
32320     If VerifyAccountNo = True Then  ' ** Function: Below.
32330       lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)
32340       If IsNumeric(.purchaseCost) = True Then
              ' ** This covers any sign changes, assures 0 = "0" not "0.00", and won't get stung if Null.
32350         gstrPurchaseCost = CStr(Val(CStr(.purchaseCost)))
32360         gstrPurchaseCost = Rem_Dollar(gstrPurchaseCost, lngCurrID)  ' ** Module Function: modStringFuncs.
32370       End If
32380     End If
32390   End With

EXITP:
32400   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseCurr_ID_cmd_DblClick(Cancel As Integer)

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseCurr_ID_cmd_DblClick"

32510   With Me
32520     If IsNull(.purchaseCurr_Date) = True Then
32530       .purchaseCurr_Date = .purchaseCurr_ID.Column(CBX_C_DATE)
32540     End If
32550     DoEvents
32560     Select Case .purchaseCurr_Date.Visible
          Case True
32570       .purchaseCurr_Date.Visible = False
32580     Case False
32590       .purchaseCurr_Date.Visible = True
32600     End Select
32610   End With

EXITP:
32620   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Public Sub purchaseCurr_ID_AfterUpdate()

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseCurr_ID_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String, strTmp03 As String, intTmp04 As Integer, lngTmp06 As Long
        Dim lngX As Long

32710   With Me

32720     If IsNull(.purchaseCurr_ID) = False Then

32730 On Error Resume Next
32740       lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)
32750       If ERR.Number <> 0 Then
              ' ** 94  Invalid use of Null.
32760 On Error GoTo ERRH
32770         .purchaseCurr_ID.Requery
32780         .purchaseCurr_ID = 150
32790         lngCurrID = 150&  ' ** Default to USD.
32800       Else
32810 On Error GoTo ERRH
32820       End If

32830       blnFound = False
32840       If lngCurrID <> 150& Then
32850         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
32860           arr_varNoChar = .Parent.NoChar_Get  ' ** Form Function: frmJournal.
32870           lngNoChars = UBound(arr_varNoChar, 2) + 1&
32880         End If
32890         For lngX = 0& To (lngNoChars - 1&)
32900           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
32910             blnFound = True
32920           End If
32930         Next
32940       End If

32950       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
32960         .purchaseICash.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
32970         .purchaseICash.FontSize = 10
32980         .purchaseICash.FontBold = True
32990         .purchasePCash.FontName = "Courier New"
33000         .purchasePCash.FontSize = 10
33010         .purchasePCash.FontBold = True
33020         .purchaseCost.FontName = "Courier New"
33030         .purchaseCost.FontSize = 10
33040         .purchaseCost.FontBold = True
33050       Case False
33060         If .purchaseICash.FontName <> "Fixedsys" Then
33070           .purchaseICash.FontName = "Fixedsys"
33080           .purchaseICash.FontSize = 9
33090           .purchaseICash.FontBold = False
33100           .purchasePCash.FontName = "Fixedsys"
33110           .purchasePCash.FontSize = 9
33120           .purchasePCash.FontBold = False
33130           .purchaseCost.FontName = "Fixedsys"
33140           .purchaseCost.FontSize = 9
33150           .purchaseCost.FontBold = False
33160         End If
33170       End Select

33180       If lngCurrID = 150& Then  ' ** USD.
33190         strTmp01 = "Currency"
33200         intTmp04 = 2
33210         lngTmp06 = CLR_WHT
33220       Else
33230         strTmp03 = .purchaseCurr_ID.Column(CBX_C_SYM)  ' ** Symbol.
33240         If .purchaseCurr_ID.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
33250           intTmp04 = 4
33260         Else
33270           intTmp04 = .purchaseCurr_ID.Column(CBX_C_DEC)  ' ** Decimal places.
33280         End If
33290         intLen = Len(strTmp03)
33300         Select Case intLen
              Case 1
                'strTmp03 = strTmp03
33310         Case 2
33320           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2)
33330         Case 3
33340           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3)
33350         Case 4
33360           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & "\" & Mid(strTmp03, 4)
33370         Case 5
33380           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & _
                  "\" & Mid(strTmp03, 4, 1) & "\" & Mid(strTmp03, 5)
33390         End Select
33400         Select Case intTmp04
              Case 0
33410           strTmp01 = strTmp03 & "#,##0;(" & strTmp03 & "#,##0);" & strTmp03 & "0;" & strTmp03 & "0"
33420         Case 1  ' ** None have this currently.
33430           strTmp01 = strTmp03 & "#,##0.0;(" & strTmp03 & "#,##0.0);" & strTmp03 & "0.0;" & strTmp03 & "0.0"
33440         Case 2
33450           strTmp01 = strTmp03 & "#,##0.00;(" & strTmp03 & "#,##0.00);" & strTmp03 & "0.00;" & strTmp03 & "0.00"
33460         Case 3
33470           strTmp01 = strTmp03 & "#,##0.000;(" & strTmp03 & "#,##0.000);" & strTmp03 & "0.000;" & strTmp03 & "0.000"
33480         End Select
33490         lngTmp06 = CLR_VLTGRN
33500       End If
33510       .purchaseICash.Format = strTmp01
33520       .purchaseICash.DecimalPlaces = intTmp04
33530       .purchaseICash.BackColor = lngTmp06
33540       .purchasePCash.Format = strTmp01
33550       .purchasePCash.DecimalPlaces = intTmp04
33560       .purchasePCash.BackColor = lngTmp06
33570       .purchaseCost.Format = strTmp01
33580       .purchaseCost.DecimalPlaces = intTmp04
33590       .purchaseCost.BackColor = lngTmp06

33600     End If

33610     If lngCurrID <> .purchaseCurr_ID.OldValue Then
33620       Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
33630     End If

33640   End With

EXITP:
33650   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseCurr_ID_NotInList(NewData As String, Response As Integer)

33700 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseCurr_ID_NotInList"

33710   With Me
33720     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
33730     .purchaseCurr_ID.SetFocus
33740     Response = acDataErrContinue
33750   End With

EXITP:
33760   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseCurr_ID_Exit(Cancel As Integer)

33800 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseCurr_ID_Exit"

33810   With Me
33820     If .purchaseCurr_ID <> .purchaseCurr_ID.OldValue Then
33830       Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
33840     End If
33850   End With

EXITP:
33860   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub purchaseDescription_AfterUpdate()

33900 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseDescription_AfterUpdate"

33910   With Me
33920     If .purchaseDescription <> .purchaseDescription.OldValue Then
33930       Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
33940     End If
33950   End With

EXITP:
33960   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub purchaseDescription_LostFocus()

34000 On Error GoTo ERRH

        Const THIS_PROC As String = "purchaseDescription_LostFocus"

34010   VerifyAccountNo  ' ** Function: Below.

EXITP:
34020   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub tglPurchaseSale_GotFocus()

34100 On Error GoTo ERRH

        Const THIS_PROC As String = "tglPurchaseSale_GotFocus"

34110   With Me
34120     blnPurchaseSale_Focus = True
34130     Select Case .posted
          Case True
34140       .tglPurchaseSale_true_raised_focus_dots_img.Visible = True
34150       .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = False
34160     Case False
34170       .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = True
34180       .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
34190     End Select
34200     .tglPurchaseSale_false_raised_img.Visible = False
34210     .tglPurchaseSale_false_raised_focus_img.Visible = False
34220     .tglPurchaseSale_false_raised_focus_dots_img.Visible = False
34230     .tglPurchaseSale_false_sunken_focus_dots_img.Visible = False
34240     .tglPurchaseSale_false_raised_img_dis.Visible = False
34250     .tglPurchaseSale_true_raised_img.Visible = False
34260     .tglPurchaseSale_true_raised_focus_img.Visible = False
34270     .tglPurchaseSale_true_sunken_focus_dots_img.Visible = False
34280     .tglPurchaseSale_true_raised_img_dis.Visible = False
34290   End With

EXITP:
34300   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub tglPurchaseSale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

34400 On Error GoTo ERRH

        Const THIS_PROC As String = "tglPurchaseSale_MouseDown"

34410   With Me
34420     blnPurchaseSale_MouseDown = True
34430     Select Case .posted
          Case True
34440       .tglPurchaseSale_true_sunken_focus_dots_img.Visible = True
34450       .tglPurchaseSale_false_sunken_focus_dots_img.Visible = False
34460     Case False
34470       .tglPurchaseSale_false_sunken_focus_dots_img.Visible = True
34480       .tglPurchaseSale_true_sunken_focus_dots_img.Visible = False
34490     End Select
34500     .tglPurchaseSale_false_raised_img.Visible = False
34510     .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = False
34520     .tglPurchaseSale_false_raised_focus_img.Visible = False
34530     .tglPurchaseSale_false_raised_focus_dots_img.Visible = False
34540     .tglPurchaseSale_false_raised_img_dis.Visible = False
34550     .tglPurchaseSale_true_raised_img.Visible = False
34560     .tglPurchaseSale_true_raised_focus_img.Visible = False
34570     .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
34580     .tglPurchaseSale_true_raised_img_dis.Visible = False
34590   End With

EXITP:
34600   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub tglPurchaseSale_Click()
' ** I'm going to try using [posted] as a signal (since it's not used for anything else).

34700 On Error GoTo ERRH

        Const THIS_PROC As String = "tglPurchaseSale_Click"

34710   With Me
          ' ** Do I need to do anything with the images, or will that be handled by the other events?
34720     Select Case .posted
          Case True  ' ** If it's True, flip to False, and vice versa.
34730       blnGoToSaleForm = False
34740       .posted = False
34750       DoCmd.RunCommand acCmdSaveRecord
34760     Case False
34770       blnGoToSaleForm = True
34780       .posted = True
34790       DoCmd.RunCommand acCmdSaveRecord
34800     End Select
34810     Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
34820     DoEvents
34830   End With

EXITP:
34840   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub tglPurchaseSale_KeyDown(KeyCode As Integer, Shift As Integer)
' ** See Note, above.

34900 On Error GoTo ERRH

        Const THIS_PROC As String = "tglPurchaseSale_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

34910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
34920   intShiftDown = (Shift And acShiftMask) > 0
34930   intAltDown = (Shift And acAltMask) > 0
34940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
34950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
34970       With Me
34980         intRetVal = 0
34990         If .cmbRevenueCodes.Enabled = True Then
35000           .cmbRevenueCodes.SetFocus
35010         ElseIf .cmbTaxCodes.Enabled = True Then
35020           .cmbTaxCodes.SetFocus
35030         ElseIf .cmdPurchaseOK.Enabled = True Then
35040           .cmdPurchaseOK.SetFocus
35050         ElseIf .cmdPurchaseCancel.Enabled = True Then
35060           .cmdPurchaseCancel.SetFocus
35070         ElseIf .cmdPurchaseClose.Enabled = True Then
35080           .cmdPurchaseClose.SetFocus
35090         ElseIf .purchaseAccountNo.Enabled = True Then
35100           .purchaseAccountNo.SetFocus
35110         Else
35120           .purchaseTransDate.SetFocus
35130         End If
35140       End With
35150     End Select
35160   End If

        ' ** Shift keys.
35170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
35190       With Me
35200         intRetVal = 0
35210         .purchaseDescription.SetFocus
35220       End With
35230     End Select
35240   End If

EXITP:
35250   KeyCode = intRetVal
35260   Exit Sub

ERRH:
410     intRetVal = 0
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Sub

Private Sub tglPurchaseSale_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35300 On Error GoTo ERRH

        Const THIS_PROC As String = "tglPurchaseSale_MouseMove"

35310   With Me
35320     If blnPurchaseSale_MouseDown = False Then
35330       Select Case .posted
            Case True
35340         Select Case blnPurchaseSale_Focus
              Case True
35350           .tglPurchaseSale_true_raised_focus_dots_img.Visible = True
35360           .tglPurchaseSale_true_raised_focus_img.Visible = False
35370         Case False
35380           .tglPurchaseSale_true_raised_focus_img.Visible = True
35390           .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
35400         End Select
35410         .tglPurchaseSale_false_raised_focus_img.Visible = False
35420         .tglPurchaseSale_false_raised_focus_dots_img.Visible = False
35430       Case False
35440         Select Case blnPurchaseSale_Focus
              Case True
35450           .tglPurchaseSale_false_raised_focus_dots_img.Visible = True
35460           .tglPurchaseSale_false_raised_focus_img.Visible = False
35470         Case False
35480           .tglPurchaseSale_false_raised_focus_img.Visible = True
35490           .tglPurchaseSale_false_raised_focus_dots_img.Visible = False
35500         End Select
35510         .tglPurchaseSale_true_raised_focus_img.Visible = False
35520         .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
35530       End Select
35540       .tglPurchaseSale_false_raised_img.Visible = False
35550       .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = False
35560       .tglPurchaseSale_false_raised_img_dis.Visible = False
35570       .tglPurchaseSale_false_sunken_focus_dots_img.Visible = False
35580       .tglPurchaseSale_true_raised_img.Visible = False
35590       .tglPurchaseSale_true_sunken_focus_dots_img.Visible = False
35600       .tglPurchaseSale_true_raised_img_dis.Visible = False
35610     End If
35620   End With

EXITP:
35630   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub tglPurchaseSale_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

35700 On Error GoTo ERRH

        Const THIS_PROC As String = "tglPurchaseSale_MouseUp"

35710   With Me
35720     Select Case .posted
          Case True
35730       .tglPurchaseSale_true_raised_focus_dots_img.Visible = True
35740       .tglPurchaseSale_false_raised_focus_dots_img.Visible = False
35750     Case False
35760       .tglPurchaseSale_false_raised_focus_dots_img.Visible = True
35770       .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
35780     End Select
35790     .tglPurchaseSale_false_raised_img.Visible = False
35800     .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = False
35810     .tglPurchaseSale_false_raised_focus_img.Visible = False
35820     .tglPurchaseSale_false_raised_img_dis.Visible = False
35830     .tglPurchaseSale_false_sunken_focus_dots_img.Visible = False
35840     .tglPurchaseSale_true_raised_img.Visible = False
35850     .tglPurchaseSale_true_raised_focus_img.Visible = False
35860     .tglPurchaseSale_true_sunken_focus_dots_img.Visible = False
35870     .tglPurchaseSale_true_raised_img_dis.Visible = False
35880     blnPurchaseSale_MouseDown = False
35890   End With

EXITP:
35900   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub tglPurchaseSale_LostFocus()

36000 On Error GoTo ERRH

        Const THIS_PROC As String = "tglPurchaseSale_LostFocus"

36010   With Me
36020     Select Case .posted
          Case True
36030       .tglPurchaseSale_true_raised_img.Visible = True
36040       .tglPurchaseSale_false_raised_img.Visible = False
36050     Case False
36060       .tglPurchaseSale_false_raised_img.Visible = True
36070       .tglPurchaseSale_true_raised_img.Visible = False
36080     End Select
36090     .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = False
36100     .tglPurchaseSale_false_raised_focus_img.Visible = False
36110     .tglPurchaseSale_false_raised_focus_dots_img.Visible = False
36120     .tglPurchaseSale_false_sunken_focus_dots_img.Visible = False
36130     .tglPurchaseSale_false_raised_img_dis.Visible = False
36140     .tglPurchaseSale_true_raised_focus_img.Visible = False
36150     .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
36160     .tglPurchaseSale_true_sunken_focus_dots_img.Visible = False
36170     .tglPurchaseSale_true_raised_img_dis.Visible = False
36180     blnPurchaseSale_Focus = False
36190   End With

EXITP:
36200   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Enter()

36300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Enter"

36310   With Me
36320     Select Case IsNull(.cmbRevenueCodes)
          Case True
36330       ExpandCombo .cmbRevenueCodes  ' ** Module Procedure: modUtilities.
36340     Case False
36350       If .cmbRevenueCodes = REVID_INC Or .cmbRevenueCodes = REVID_EXP Then
36360         ExpandCombo .cmbRevenueCodes  ' ** Module Procedure: modUtilities.
36370       End If
36380     End Select
36390   End With

EXITP:
36400   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Change()
' ** cmbRevenueCodes:
' **   RowSource is 0-Based:
' **     Col 0: revcode_ID
' **     Col 1: revcode_DESC
' **     Col 2: revcode_TYPE
' **     Col 3: revcode_TYPE_Code (I/E)
' **     Col 4: taxcode_type
' **     Col 5: taxcode_type_Code (I/D)
' **   BoundColumn is 1-Based:
' **     Col 0: ListIndex

36500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Change"

36510   With Me
36520     If .cmbRevenueCodes <> .cmbRevenueCodes.OldValue Then
36530       Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
36540     End If
36550   End With

EXITP:
36560   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbRevenueCodes_AfterUpdate()
' ** cmbRevenueCodes:
' **   RowSource is 0-Based:
' **     Col 0: revcode_ID
' **     Col 1: revcode_DESC
' **     Col 2: revcode_TYPE
' **     Col 3: revcode_TYPE_Code (I/E)
' **     Col 4: taxcode_type
' **     Col 5: taxcode_type_Code (I/D)
' **   BoundColumn is 1-Based:
' **     Col 0: ListIndex

36600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_AfterUpdate"

        Dim strRevCode As String, lngTaxcode As Long

        ' ** Make sure tax code matches selected revenue code.
36610   With Me
36620     If IsNull(.cmbRevenueCodes) = False Then
36630       If .cmbRevenueCodes > 0 Then
36640         If gblnLinkRevTaxCodes = True Then
36650           strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
36660           If IsNull(.cmbTaxCodes) = True Then
36670             If strRevCode = "I" Then
36680               .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
36690             ElseIf strRevCode = "E" Then
36700               .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
36710             End If
36720           Else
36730             lngTaxcode = .cmbTaxCodes
36740             If lngTaxcode = 0& Then
36750               If strRevCode = "I" Then
36760                 .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
36770               ElseIf strRevCode = "E" Then
36780                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
36790               End If
36800             Else
36810               If strRevCode = "I" And .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
                      'MsgBox "This will change..."
36820                 .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
36830               ElseIf strRevCode = "E" And .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
                      'MsgBox "This will change..."
36840                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
36850               End If
36860             End If
36870           End If
36880         End If
36890       Else
36900         .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
36910       End If
36920     Else
36930       .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
36940     End If
36950   End With

EXITP:
36960   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbRevenueCodes_LostFocus()

37000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_LostFocus"

37010   VerifyAccountNo  ' ** Function: Below.

EXITP:
37020   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbTaxCodes_Enter()

37100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Enter"

37110   With Me
37120     Select Case IsNull(.cmbTaxCodes)
          Case True
37130       ExpandCombo .cmbTaxCodes  ' ** Module Procedure: modUtilities.
37140     Case False
37150       If .cmbTaxCodes = TAXID_INC Or .cmbTaxCodes = TAXID_DED Then
37160         ExpandCombo .cmbTaxCodes  ' ** Module Procedure: modUtilities.
37170       End If
37180     End Select
37190   End With

EXITP:
37200   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbTaxCodes_Change()
' ** cmbTaxCodes:
' **   RowSource is 0-Based:
' **     Col 0: taxcode
' **     Col 1: taxcode_description
' **     Col 2: taxcode_type
' **     Col 3: taxcode_type_Code (I/D)
' **     Col 4: revcode_TYPE
' **     Col 5: revcode_TYPE_Code (I/E)
' **   BoundColumn is 1-Based:
' **     Col 0: ListIndex

37300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Change"

37310   With Me
37320     If .cmbTaxCodes <> .cmbTaxCodes.OldValue Then
37330       Sub3Purch_Changed True, Me  ' ** Module Procedure: modPurchaseSold.
37340     End If
37350   End With

EXITP:
37360   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbTaxCodes_AfterUpdate()
' ** cmbTaxCodes:
' **   RowSource is 0-Based:
' **     Col 0: taxcode
' **     Col 1: taxcode_description
' **     Col 2: taxcode_type
' **     Col 3: taxcode_type_Code (I/D)
' **     Col 4: revcode_TYPE
' **     Col 5: revcode_TYPE_Code (I/E)
' **   BoundColumn is 1-Based:
' **     Col 0: ListIndex

37400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_AfterUpdate"

        Dim strRevCode As String, lngTaxcode As Long

        ' ** Make sure revenue code matches selected tax code.
37410   With Me
37420     If IsNull(.cmbTaxCodes) = False Then
37430       If .cmbTaxCodes > 0& Then
37440         If gblnLinkRevTaxCodes = True Then
37450           lngTaxcode = .cmbTaxCodes
37460           strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
37470           If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
37480             If IsNull(.cmbRevenueCodes) = True Then
37490               .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37500             Else
37510               If .cmbRevenueCodes = 0 Then
37520                 .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37530               Else
37540                 If strRevCode = "E" Then
                        'MsgBox "This will change..."
37550                   .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37560                 End If
37570               End If
37580             End If
37590           ElseIf .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
37600             If IsNull(.cmbRevenueCodes) = True Then
37610               .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
37620             Else
37630               If .cmbRevenueCodes = 0 Then
37640                 .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
37650               Else
37660                 If strRevCode = "I" Then
                        'MsgBox "This will change..."
37670                   .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
37680                 End If
37690               End If
37700             End If
37710           End If
37720         End If
37730       Else
37740         .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
37750       End If
37760     Else
37770       .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
37780     End If
37790   End With

EXITP:
37800   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub cmbTaxCodes_LostFocus()

37900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_LostFocus"

37910   VerifyAccountNo  ' ** Function: Below.

EXITP:
37920   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Public Sub subNullShareFace()

38000 On Error GoTo ERRH

        Const THIS_PROC As String = "subNullShareFace"

38010   With Me
38020     .purchaseShareFace = vbNullString
38030   End With

EXITP:
38040   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Public Function VerifyAccountNo() As Boolean

38100 On Error GoTo ERRH

        Const THIS_PROC As String = "VerifyAccountNo"

        Dim strAccountNo As String, strControl As String
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

38110   blnRetVal = True

38120   With Me

38130     If ((gblnPurchaseValidated = False And gblnPurchaseChanged = True) And (lngErrCnt < 3&)) Then

38140 On Error Resume Next
38150       strControl = Screen.ActiveControl.Name
38160 On Error GoTo ERRH

38170       If strControl <> "purchaseAccountno" And blnAccountNoHasFocus = False Then
38180         blnAccountNoHasFocus = True
38190 On Error Resume Next  ' ** Not sure why it objects here.
38200         .purchaseAccountNo.SetFocus  ' ** 2110: Microsoft Access can't move the focus to the control |.
38210 On Error GoTo ERRH
38220       End If
38230       If (IsNull(.purchaseAccountNo) = True Or .purchaseAccountNo = vbNullString) Then
38240         lngErrCnt = lngErrCnt + 1&
38250         MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Invalid Entry"
38260         gblnPurchaseValidated = False
38270       Else
              ' ** Check to see if it is a valid account
38280         lngErrCnt = 0&
38290         Select Case gblnAccountNoWithType
              Case True
38300           strAccountNo = Left(.purchaseAccountNo, 2) & Mid(.purchaseAccountNo, 3)
38310         Case False
38320           strAccountNo = .purchaseAccountNo
38330         End Select
38340         varTmp00 = DLookup("[accountno]", "account", "[accountno] = '" & strAccountNo & "' And ([dateclosed] Is Null)")
38350         Select Case IsNull(varTmp00) 'blnAccountNoHasFocus = False Then
              Case True
38360           MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Invalid Entry"
38370           blnAccountNoHasFocus = True
38380           .purchaseAccountNo.SetFocus
38390           gblnPurchaseValidated = False
38400         Case False
38410           gblnPurchaseValidated = True
38420           blnStartTrans = False
38430         End Select
38440       End If
38450     Else
38460       lngErrCnt = 0&
38470     End If

38480     If (gblnPurchaseChanged = False) Or (.NavigationButtons = True) Then
            ' ** If the changed flag isn't set, ot the navigation buttons ARE active,
            ' ** gblnPurchaseValidated may not be set before entering this funciton.
38490       blnRetVal = True
38500     Else
38510       blnRetVal = gblnPurchaseValidated
38520     End If

38530   End With

EXITP:
38540   VerifyAccountNo = blnRetVal
38550   Exit Function

ERRH:
410     blnRetVal = False
420     THAT_PROC = THIS_PROC
430     That_Erl = Erl
440     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
450     Resume EXITP

End Function

Public Sub Map_NewRec()

38600 On Error GoTo ERRH

        Const THIS_PROC As String = "Map_NewRec"

38610   With Me
38620 On Error Resume Next
38630     DoCmd.RunCommand acCmdRecordsGoToNew
38640     If ERR.Number <> 0 Then
38650       blnMapNewRec = True
38660       .TimerInterval = 500&
38670     End If
38680 On Error GoTo ERRH
38690   End With

EXITP:
38700   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
440     Resume EXITP

End Sub

Private Sub ChkAverage()

38800 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkAverage"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountNo As String, lngAssetNo As Long

38810   With Me
38820     .purchaseIsAverage = False
38830     If IsNull(.purchaseAssetNo) = False Then
38840       strAccountNo = .purchaseAccountNo
38850       lngAssetNo = .purchaseAssetNo
38860       Set dbs = CurrentDb
            ' ** ActiveAssets, just IsAverage = True, grouped, by accountno, assetno, IsAverage, with cnt.
38870       Set qdf = dbs.QueryDefs("qryJournal_Purchase_03d")
38880       Set rst = qdf.OpenRecordset
38890       If rst.BOF = True And rst.EOF = True Then
              ' ** No averaging anywhere.
38900         rst.Close
38910       Else
38920         rst.Close
              ' ** ActiveAssets, grouped, by accountno, assetno, IsAverage,
              ' ** with cnt, by specified [actno], [astno].
38930         Set qdf = dbs.QueryDefs("qryJournal_Purchase_03e")
38940         With qdf.Parameters
38950           ![actno] = strAccountNo
38960           ![astno] = lngAssetNo
38970         End With
38980         Set rst = qdf.OpenRecordset
38990         If rst.BOF = True And rst.EOF = True Then
                ' ** This asset not averaged for this account.
39000           rst.Close
39010         Else
39020           rst.MoveFirst
39030           .purchaseIsAverage = True
39040           rst.Close
39050         End If
39060       End If
39070       dbs.Close
39080     End If
39090   End With

EXITP:
39100   Set rst = Nothing
39110   Set qdf = Nothing
39120   Set dbs = Nothing
39130   Exit Sub

ERRH:
410     THAT_PROC = THIS_PROC
420     That_Erl = Erl
430     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
440     Resume EXITP

End Sub

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

39200 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

39210   Select Case intWhere
        Case acCmdRecordsGoToFirst
39220     DoCmd.RunCommand acCmdRecordsGoToFirst
39230   Case acCmdRecordsGoToLast
39240     DoCmd.RunCommand acCmdRecordsGoToLast
39250   Case acCmdRecordsGoToPrevious
39260     DoCmd.RunCommand acCmdRecordsGoToPrevious
39270   Case acCmdRecordsGoToNext
39280     DoCmd.RunCommand acCmdRecordsGoToNext
39290   Case acCmdRecordsGoToNew
39300     DoCmd.RunCommand acCmdRecordsGoToNew
39310   Case Else
39320     If IsMissing(varID) = False Then
39330       With Me
39340         Set rst = .RecordsetClone
39350         With rst
39360           .FindFirst "[purchaseID] = " & CStr(varID)
39370           If .NoMatch = False Then
39380             Me.Bookmark = .Bookmark
39390           End If
39400           .Close
39410         End With
39420       End With
39430     End If
39440   End Select

EXITP:
39450   Set rst = Nothing
39460   Exit Sub

ERRH:
410     Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
420     Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
430     Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub
