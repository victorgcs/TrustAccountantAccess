VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournal_Sub4_Sold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournal_Sub4_Sold"

'VGC 10/01/2017: CHANGES!

' #########################
' ## Use VBA_RenumErrh().  39980
' #########################

' ** Shortcut Alt keys responsive from this form:
' **   Paid Total:       A {cmdPaidTotal}
' **   Close:            C {cmdSaleClose}
' **   Dividend:         D {opgJournal_optDividend on frmJournal}
' **   Cancel:           E {cmdSaleCancel}
' **   Interest:         I {opgJournal_optInterest on frmJournal}
' **   Lot Info:         L {cmdSaleLotInfo}
' **   Miscellaneous:    M {opgJournal_optMisc on frmJournal}
' **   Inc/Exp Codes:    N {cmbRevenueCodes}
' **   OK:               O {cmdSaleOK}
' **   Purchase:         P {opgJournal_optPurchase on frmJournal}
' **   Sale:             S {opgJournal_optSale on frmJournal}
' **   Tax Codes:        T {cmbTaxCodes}
' **   Exit:             X {cmdSaleClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Next Record:      N {MoveRec}
' **   Previous Record:  P {MoveRec}
' **   Switch:           T {cmdSwitch on frmJournal}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Mark Reinvest:    R {tglSaleReinvest}
' **   Close Timer:      T {Form_KeyDown}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Paid:      F5 {RecalcPaidTotal}
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Sold, Withdrawn are always INCOME for Tax Codes.
' ** Liability is always Expense for Tax Codes.
' ** Cost Adj. is INCOME if negative, EXPENSE if positive.

'QRY: qryJournal_Sale_03                 REF: FormRef('frmJournal_Sub4_Sold')

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

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Combo box column constants: saleAssetNo.
'Private Const CBX_A_DESC   As Integer = 0  'totdesc
'Private Const CBX_A_CUSIP  As Integer = 1  'cusip
Private Const CBX_A_ASTNO  As Integer = 2  'assetno
'Private Const CBX_A_RATE   As Integer = 3  'rate
'Private Const CBX_A_DUE    As Integer = 4  'due
'Private Const CBX_A_TYPE   As Integer = 5  'assettype
'Private Const CBX_A_TDESC  As Integer = 6  'assettype_description
Private Const CBX_A_TAX    As Integer = 7  'taxcode
Private Const CBX_A_CURRID As Integer = 8  'curr_id

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

' ** Combo box column constants: saleCurr_ID.
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
Private blnSaleReinvest_Focus As Boolean, blnSaleReinvest_MouseDown As Boolean

Private lngCurrID As Long
Private blnSoldCosted As Boolean, blnCmdButton As Boolean, blnGoToSaleReinvest As Boolean
Private blnGoingToReport As Boolean, blnGoneToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long, lngGTR_ID As Long
Private blnAccountNoErr As Boolean, blnStartTrans As Boolean, blnDefTypeAssigned As Boolean
Private blnFromAcctNoKeyDown As Boolean, blnFromFormError As Boolean, blnFromElsewhere As Boolean
Private blnAccountNoHasFocus As Boolean, blnSpecialCap As Boolean, intSpecialCapOpt As Integer, blnWarned As Boolean, blnClickedLotInfo As Boolean
Private datPostingDate As Date, lngDateCnt As Long, lngErrCnt As Long
Private blnBeenToLotInfo As Boolean, blnFromSaleAssetnoEnter As Boolean
Private strCallingForm As String, lngTpp As Long
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
180           blnFromSaleAssetnoEnter = False: blnWarned = False
190           blnAccountNoErr = False: blnStartTrans = False: blnDefTypeAssigned = False
200           blnFromAcctNoKeyDown = False: blnFromFormError = False: blnClickedLotInfo = False
210           gblnSwitchTo = False: blnFromElsewhere = False
220           blnGTR_Emblem = False: lngGTR_Stat = 0&: lngGTR_ID = 0&: blnGoneToReport = False
230           blnSoldCosted = False

240           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
250           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False
260           blnSaleReinvest_Focus = False: blnSaleReinvest_MouseDown = False

270           .cmdLock.Enabled = False
280           .cmdLock_open_raised_img.Visible = False
290           .cmdLock_open_raised_img_dis.Visible = False
300           .cmdLock_closed_raised_img.Visible = False

310         Else
320           Cancel = -1
330         End If
340       Else
350         Cancel = -1
360       End If
370     End With

380     If Cancel = -1 Then
390       Beep
400       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
410     End If

EXITP:
420     Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
460     Resume EXITP

End Sub

Private Sub Form_Load()
' ** Create an instance of the Calendar class.
' ** Copyright Lebans Holdings 1999 Ltd.

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me
          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
520       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

          ' ** Create an instance of the Calendar class.
530       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
540       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
550       clsMonthClass.PositionAtCursor = True

560       blnAccountNoHasFocus = True
570       blnCmdButton = False  ' ** Set the switch to false until they push the button.
580       blnGoToSaleReinvest = False

590       datPostingDate = GetPostDate  ' ** Module Function: modUtilities.

          ' ** Do this whether or not Currency is visible.
600       CLR_DISABLED_FG = CLR_DKGRY
610       CLR_DISABLED_BG = CLR_LTTEAL

620       .saleCurr_ID.Enabled = False
630       .saleCurr_ID.Locked = True
640       .saleCurr_ID.ForeColor = CLR_DISABLED_FG
650       .saleCurr_ID.BackColor = CLR_VLTGRN
660       saleCurr_ID_AfterUpdate  ' ** Procedure: Below.

          ' ** Set the default Posting Date (transdate), entered by the user on opening.
670       .saleTransDate.DefaultValue = "#" & CDate(Format(datPostingDate, "mm/dd/yyyy")) & "#"

680       gstrFormQuerySpec = .Parent.Name  ' ** Make sure this is set for the assetno combo box.

690       gblnSaleValidated = False
700       gblnSaleChanged = False

          ' ** Make sure the SpecialCap options are set.
710       If gstrCo_Name = vbNullString Then
720         CoOptions_Read  ' ** Module Function: modStartupFunctions.
730       End If

          ' ** Load SpecialCapGainLoss setting.
740       blnSpecialCap = gblnSpecialCapGainLoss
750       intSpecialCapOpt = gintSpecialCapGainLossOpt

760       .SpecialCapGainLoss = blnSpecialCap

770       .NavigationButtons = True

780       With .Parent
790         .opgJournal.Enabled = True
800         .cmdSwitch.Enabled = True
810         .cmdSwitch_raised_img.Visible = True
820         .cmdSwitch_raised_img_dis.Visible = False
830         .cmdSwitch_raised_semifocus_dots_img.Visible = False
840         .cmdSwitch_raised_focus_img.Visible = False
850         .cmdSwitch_raised_focus_dots_img.Visible = False
860         .cmdSwitch_sunken_focus_dots_img.Visible = False
870       End With
880       .cmdSaleClose.Enabled = True
890       .cmdSaleOK.Enabled = False
900       .cmdSaleCancel.Enabled = False
910       .Parent.NavVis True  ' ** Form Procedure: frmJournal.

920       If .NewRecord = True Then
930         .tglSaleReinvest.Enabled = False
940         .tglSaleReinvest_false_raised_img_dis.Visible = True
950         .tglSaleReinvest_false_raised_img.Visible = False
960         .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = False
970         .tglSaleReinvest_false_raised_focus_img.Visible = False
980         .tglSaleReinvest_false_raised_focus_dots_img.Visible = False
990         .tglSaleReinvest_false_sunken_focus_dots_img.Visible = False
1000        .tglSaleReinvest_true_raised_img.Visible = False
1010        .tglSaleReinvest_true_raised_focus_img.Visible = False
1020        .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
1030        .tglSaleReinvest_true_sunken_focus_dots_img.Visible = False
1040        .tglSaleReinvest_true_raised_img_dis.Visible = False
1050      End If
1060      .cmdSaleLotInfo.Enabled = False

1070      blnSoldCosted = False
1080      SoldCostedSet False  ' ** Procedure: Below.

1090      .AllowAdditions = True

1100      Select Case gblnIncomeTaxCoding
          Case True
1110        .cmbTaxCodes.Visible = True
1120        .cmbTaxCodes_lbl2.Visible = False
1130      Case False
1140        .cmbTaxCodes.Visible = False
1150        .cmbTaxCodes_lbl2.Visible = True
1160      End Select

1170      Select Case gblnRevenueExpenseTracking
          Case True
1180        .cmbRevenueCodes.Visible = True
1190        .cmbRevenueCodes_lbl2.Visible = False
1200      Case False
1210        .cmbRevenueCodes.Visible = False
1220        .cmbRevenueCodes_lbl2.Visible = True
1230      End Select

1240    End With

EXITP:
1250    Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
460     Resume EXITP

End Sub

Private Sub Form_Current()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim strAccountNo As String
        Dim blnContinue As Boolean
        Dim lngTmp01 As Long, dblTmp02 As Double

1310    With Me

1320      blnContinue = True

1330      DoCmd.SelectObject acForm, .Parent.Name, False

1340      blnDefTypeAssigned = False
1350      blnStartTrans = False
1360      lngErrCnt = 0&

1370      .cmbAccountHelper = Null

1380      If .saleCurr_ID.Visible = True Then
1390        If .saleCurr_Date.Visible = True Then .saleCurr_Date.Visible = False
1400        FormCurrent_Sub4 1, blnAccountNoErr, blnBeenToLotInfo, blnClickedLotInfo, blnFromSaleAssetnoEnter, _
              blnGoToSaleReinvest, blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.
1410        If lngNoChars = 0& Or IsEmpty(arr_varNoChar) Then
1420          arr_varNoChar = .Parent.NoChar_Get  ' ** Form Function: frmJournal.
1430          lngNoChars = UBound(arr_varNoChar, 2) + 1&
1440        End If
1450      End If

1460      gstrFormQuerySpec = .Parent.Name  ' ** Make sure this is set for the assetno combo box.

1470      RecalcPaidTotal  ' ** Procedure: Below.

1480      FormCurrent_Sub4 2, blnAccountNoErr, blnBeenToLotInfo, blnClickedLotInfo, blnFromSaleAssetnoEnter, _
            blnGoToSaleReinvest, blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

          ' ** Make sure validation is reset.
1490      gblnSaleValidated = False
1500      blnBeenToLotInfo = False

1510      Select Case IsNull(.saleShareFace)
          Case True
1520        .saleShareFace.Format = "#,###"
1530      Case False
1540        If .saleShareFace <> 0 Then
1550          lngTmp01 = Int(.saleShareFace)
1560          dblTmp02 = CDbl(.saleShareFace)
1570          If lngTmp01 = dblTmp02 Then
1580            .saleShareFace.Format = "#,###"
1590          Else
1600            .saleShareFace.Format = "#,###.00"
1610          End If
1620        Else
1630          .saleShareFace.Format = "#,###"
1640        End If
1650      End Select

1660      If blnContinue = True Then

1670        lngCurrID = .saleCurr_ID.Column(CBX_C_CURRID)

1680        .saleAssetno.Requery

1690        gstrSaleType = IIf(IsNull(.saleType), vbNullString, .saleType)
1700        gstrSaleAsset = IIf(IsNull(.saleAssetno), vbNullString, .saleAssetno)
1710        gstrSaleShareFace = IIf(IsNull(.saleShareFace), vbNullString, .saleShareFace)
1720        gstrSaleAccountNumber = IIf(IsNull(.saleAccountNo), vbNullString, .saleAccountNo)
1730        gstrSaleCost = IIf(IsNull(.saleCost), vbNullString, .saleCost)
1740        gstrSaleCost = Rem_Dollar(gstrSaleCost, lngCurrID)  ' ** Module Function: modStringFuncs.
1750        gstrSalePCash = IIf(IsNull(.salePCash), vbNullString, .salePCash)
1760        gstrSalePCash = Rem_Dollar(gstrSalePCash, lngCurrID)  ' ** Module Function: modStringFuncs.

1770        strAccountNo = .saleAccountNo

1780        Select Case .NewRecord
            Case True
1790          .cmdSaleCancel.Enabled = False
1800        Case False
1810          .cmdSaleCancel.Enabled = True
1820        End Select

1830        .Refresh

1840      End If  ' ** blnContinue.

1850      If blnContinue = True Then

1860        .saleCost.Locked = True

1870        FormCurrent_Sub4 3, blnAccountNoErr, blnBeenToLotInfo, blnClickedLotInfo, blnFromSaleAssetnoEnter, _
              blnGoToSaleReinvest, blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

1880        blnSoldCosted = False
1890        SoldCostedSet False  ' ** Procedure: Below.
1900        .cmdSaleLotInfo.Enabled = False

1910        .Refresh

1920        DoEvents

1930      End If  ' ** blnContinue.

1940      If blnCmdButton = False Then  ' ** That is, only run this if Form_Current() wasn't called from another button, like cmdMiscOK_Click().
1950        FormCurrent_Sub4 4, blnAccountNoErr, blnBeenToLotInfo, blnClickedLotInfo, blnFromSaleAssetnoEnter, _
              blnGoToSaleReinvest, blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.
1960      End If  ' ** blnCmdButton.

1970      .Repaint

1980      DoEvents

          ' ** Set the currency symbol.
1990      saleCurr_ID_AfterUpdate  ' ** Procedure: Below.

2000      If .saleCurr_ID.Visible = True Then
2010        saleICash_AfterUpdate  ' ** Procedure: Below.
2020        salePCash_AfterUpdate  ' ** Procedure: Below.
2030        saleCost_AfterUpdate  ' ** Procedure: Below.
2040      End If

2050      subSetAssetCombo  ' ** Procedure: Below.
2060      subTaxCode4 blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modPurchaseSold.
2070      subRevCode4 blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modPurchaseSold.

2080    End With

EXITP:
2090    Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
460     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2120    intShiftDown = (Shift And acShiftMask) > 0
2130    intAltDown = (Shift And acAltMask) > 0
2140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Next Record:      N {MoveRec}
        ' **   Previous Record:  P {MoveRec}
        ' **   Switch:           T {cmdSwitch on frmJournal}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Mark Reinvest:    R {tglInterestReinvest}
        ' **   Close Timer:      T {Form_KeyDown}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyEscape
2170        If Forms("frmJournal").opgJournal.Enabled = True Then
2180          Select Case gblnSaleChanged
              Case True
2190            cmdSaleCancel_Click  ' ** Procedure: Below.
2200          Case False
2210            cmdSaleClose_Click  ' ** Procedure: Below.
2220          End Select
2230        Else
2240          cmdSaleCancel_Click  ' ** Procedure: Below.
2250        End If
2260      Case vbKeyF5
2270        RecalcPaidTotal  ' ** Procedure: Below.
2280      Case vbKeyF7
2290        intRetVal = 0
2300        cmdCalendar1_Click  ' ** Procedure: Below.
2310      Case vbKeyF8
2320        intRetVal = 0
2330        cmdCalendar2_Click  ' ** Procedure: Below.
2340      Case vbKeyDelete
2350        gblnSaleChanged = True
2360      End Select
2370    End If

        ' ** Alt keys.
2380    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2390      Select Case intRetVal
          Case vbKeyC
2400        With Me
2410          intRetVal = 0
2420          If .cmdSaleClose.Enabled = True Then
2430            .cmdSaleClose.SetFocus
2440            cmdSaleClose_Click  ' ** Procedure: Below.
2450          Else
2460            Beep
2470          End If
2480        End With
2490      Case vbKeyE
2500        With Me
2510          intRetVal = 0
2520          If .cmdSaleCancel.Enabled = True Then
2530            .cmdSaleCancel.SetFocus
2540            cmdSaleCancel_Click  ' ** Procedure: Below.
2550          Else
2560            Beep
2570          End If
2580        End With
2590      Case vbKeyX
2600        With Me
2610          intRetVal = 0
2620          If .cmdSaleClose.Enabled = True Then
2630            cmdSaleClose_Click  ' ** Procedure: Below.
2640          End If
2650        End With
2660      End Select
2670    End If

        ' ** Ctrl keys.
2680    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2690      Select Case intRetVal
          Case vbKeyN
2700        With Me
2710          intRetVal = 0
2720          MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2730        End With
2740      Case vbKeyP
2750        With Me
2760          intRetVal = 0
2770          MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2780        End With
2790      Case vbKeyT
2800        With Me
2810          intRetVal = 0
2820          DoCmd.SelectObject acForm, .Parent.Name, False
2830          .Parent.cmdSwitch.SetFocus
2840          DoEvents
2850          .Parent.cmdSwitch_Click  ' ** Form Procedure: frmJournal.
2860        End With
2870      End Select
2880    End If

        ' ** Ctrl-Shift keys.
2890    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2900      Select Case intRetVal
          Case vbKeyF
2910        With Me
2920          intRetVal = 0
2930          .FocusHolder.SetFocus
2940        End With
2950      Case vbKeyR
2960        With Me
2970          intRetVal = 0
2980          .tglSaleReinvest.SetFocus
2990          tglSaleReinvest_Click  ' ** Procedure: Below.
3000        End With
3010      Case vbKeyT
3020        intRetVal = 0
3030        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
3040      End Select
3050    End If

EXITP:
3060    KeyCode = intRetVal
3070    Exit Sub

ERRH:
430     intRetVal = 0
440     THAT_PROC = THIS_PROC
450     That_Erl = Erl
460     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
470     Resume EXITP

End Sub

Private Sub Form_Timer()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim intPos01 As Integer

        Static strAccountNo As String, lngAssetNo As Long

3110    With Me
3120      .TimerInterval = 0&
3130      If gblnGoToReport = True Then
            ' ** The emblem on frmJournal should be still spinning.
3140        Select Case blnGoingToReport
            Case True
3150          lngGTR_Stat = lngGTR_Stat + 1&
3160          Select Case lngGTR_Stat
              Case 1&
3170            .saleAccountNo.SetFocus
3180            .saleAccountNo = strAccountNo
3190            DoEvents
3200            saleAccountNo_AfterUpdate  ' ** Procedure: Below.
3210            DoEvents
3220            .TimerInterval = 50&
3230          Case 24&
3240            .GoToReport_arw_tl_acct_img.Visible = False
3250            .GoToReport_arw_tl_jtype_img.Top = (.saleType.Top + (2& * lngTpp))
3260            .GoToReport_arw_tl_jtype_img.Left = (.saleType.Left + .saleType.Width)
3270            .GoToReport_arw_tl_jtype_img.Visible = True
3280            .saleType.SetFocus
3290            .saleType = "Sold"
3300            .saleType.SelLength = 0
3310            .saleType.SelStart = 19
3320            .cmdLock.Enabled = False
3330            .cmdLock_open_raised_img_dis.Visible = True
3340            .cmdLock_open_raised_img.Visible = False
3350            .cmdLock_closed_raised_img.Visible = False
3360            .saleShortName.Visible = True
3370            .GoToReport_box_shortname.Visible = False
3380            saleType_AfterUpdate  ' ** Procedure: Below.
3390            DoEvents
3400            .TimerInterval = 50&
3410          Case 48&
3420            .GoToReport_arw_tl_jtype_img.Visible = False
3430            .GoToReport_arw_tl_tdate_img.Top = (.saleTransDate.Top + (2& * lngTpp))
3440            .GoToReport_arw_tl_tdate_img.Left = (.saleTransDate.Left + .saleTransDate.Width)
3450            .GoToReport_arw_tl_tdate_img.Visible = True
3460            .saleTransDate.SetFocus
3470            .saleTransDate = Date
3480            .saleTransDate.SelLength = 0
3490            .saleTransDate.SelStart = 19
3500            DoEvents
3510            .TimerInterval = 50&
3520          Case 72&
3530            .GoToReport_arw_tl_tdate_img.Visible = False
3540            .GoToReport_arw_tl_asset_img.Top = (.saleAssetno_lbl.Top + (2& * lngTpp))
3550            .GoToReport_arw_tl_asset_img.Left = (.saleAssetno_lbl.Left + .saleAssetno_lbl.Width)
3560            .GoToReport_arw_tl_asset_img.Visible = True
3570            .saleAssetno.SetFocus
3580            .saleAssetno = lngAssetNo
3590            .saleAssetno.SelLength = 0
3600            .saleAssetno.SelStart = 99
3610            saleAssetno_AfterUpdate  ' ** Procedure: Below.
3620            DoEvents
3630            .TimerInterval = 50&
3640          Case 96&
3650            .GoToReport_arw_tl_asset_img.Visible = False
3660            .GoToReport_arw_tl_share_img.Top = (.saleShareFace.Top + (2& * lngTpp))
3670            .GoToReport_arw_tl_share_img.Left = (.saleShareFace.Left + .saleShareFace.Width)
3680            .GoToReport_arw_tl_share_img.Visible = True
3690            .saleShareFace.SetFocus
3700            .saleShareFace = 10
3710            .saleShareFace.SelLength = 0
3720            .saleShareFace.SelStart = 19
3730            saleShareFace_AfterUpdate  ' ** Procedure: Below.
3740            DoEvents
3750            .TimerInterval = 50&
3760          Case 120&
3770            .GoToReport_arw_tl_share_img.Visible = False
3780            .GoToReport_arw_tl_adate_img.Top = (.saleAssetDate.Top + (2& * lngTpp))
3790            .GoToReport_arw_tl_adate_img.Left = (.saleAssetDate.Left + .saleAssetDate.Width)
3800            .GoToReport_arw_tl_adate_img.Visible = True
3810            .saleAssetDate.SetFocus
3820            .saleAssetDate = Date
3830            .saleAssetDate.SelLength = 0
3840            .saleAssetDate.SelStart = 19
3850            DoEvents
3860            .TimerInterval = 50&
3870          Case 144&
3880            .GoToReport_arw_tl_adate_img.Visible = False
3890            .GoToReport_arw_tl_taxlot_img.Top = (.cmdSaleLotInfo.Top + (2& * lngTpp))
3900            .GoToReport_arw_tl_taxlot_img.Left = (.cmdSaleLotInfo.Left + .cmdSaleLotInfo.Width)
3910            .GoToReport_arw_tl_taxlot_img.Visible = True
3920            .cmdSaleLotInfo.SetFocus
3930            .cmdSaleOK.Enabled = False
3940            DoEvents
3950            cmdSaleLotInfo_Click  ' ** Procedure: Below.
3960            DoEvents
3970          Case Else
                ' ** Just let the emblem go.
                ' ** lngGTR_Stat will NOT match frmJournal's!
3980            If lngGTR_Stat < 144& Then
3990              .TimerInterval = 50&
4000            Else
4010              .TimerInterval = 0&
4020              DoCmd.Hourglass False
4030              blnGoingToReport = False
4040              gblnGoToReport = False
4050              blnGTR_Emblem = False
4060              .Parent.GTREmblem_Set False  ' ** Form Function: frmJournal.
4070              Beep
4080            End If
4090          End Select
4100        Case False
4110          If blnGTR_Emblem = False Then
4120            blnGTR_Emblem = .Parent.GTREmblem_Get  ' ** Form Function: frmJournal.
4130          End If
4140          strAccountNo = GetSoldAsset  ' ** Module Function: modGoToReportFuncs.
4150          If strAccountNo <> vbNullString Then
4160            blnGoingToReport = True
4170            intPos01 = InStr(strAccountNo, ";")
4180            lngAssetNo = Val(Mid(strAccountNo, (intPos01 + 1)))
4190            strAccountNo = Left(strAccountNo, (intPos01 - 1))
4200            .TimerInterval = 50&
4210          Else
4220            .TimerInterval = 0&
4230            .Parent.GTREmblem_Set False  ' ** Form Function: frmJournal.
4240            blnGoingToReport = False
4250            gblnGoToReport = False
4260            blnGTR_Emblem = False
4270            .GoToReport_arw_tl_acct_img.Visible = False
4280            .cmdLock.Enabled = False
4290            .cmdLock_open_raised_img_dis.Visible = True
4300            .cmdLock_open_raised_img.Visible = False
4310            .cmdLock_closed_raised_img.Visible = False
4320            .saleShortName.Visible = True
4330            .GoToReport_box_shortname.Visible = False
4340            Beep
4350            DoCmd.Hourglass False
4360            MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                  "There are insufficient asset holdings to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
4370          End If
4380        End Select
4390      ElseIf blnGoneToReport = True Then
4400        .cmdSaleCancel.SetFocus
4410        cmdSaleCancel_Click  ' ** Procedure: Below.
4420      End If
4430    End With

EXITP:
4440    Exit Sub

ERRH:
430     DoCmd.Hourglass False
440     blnGoingToReport = False
450     gblnGoToReport = False
460     blnGTR_Emblem = False
470     Me.Parent.GTREmblem_Set False  ' ** Form Function: frmJournal.
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Private Sub Form_GotFocus()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_GotFocus"

4510    With Me
4520      If .saleAccountNo.Enabled = True Then
4530        .saleAccountNo.SetFocus
4540      Else
4550        .saleTransDate.SetFocus
4560      End If
4570    End With

EXITP:
4580    Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strControl As String
        Dim blnMsgGiven As Boolean

4610    blnMsgGiven = False

4620    If DataErr <> 0 Then
4630      Select Case DataErr
          Case 13, 2108, 2110, 2116, 2135, 2169, 2474, 2501, 3020, 3163, 3314, 3315, 7753, 8519
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
4640        Response = acDataErrContinue
4650      Case 2046  ' ** The command or action isn't available now (first or last record).
            ' ** Do nothing.
4660        Response = acDataErrContinue
4670      Case 2113  ' ** The value you entered isn't valid for this field.
4680  On Error Resume Next
4690        strControl = Screen.ActiveControl.Name
4700  On Error GoTo ERRH
4710        DoCmd.Hourglass False
4720        Select Case strControl
            Case "saleShareFace"
4730          strControl = "Share/Face."
4740        Case "salePCash"
4750          strControl = "Principal Cash"
4760        Case "saleICash"
4770          strControl = "Income Cash"
4780        Case "saleCost"
4790          strControl = "Cost"
4800        Case "saleTransDate", "saleAssetDate"
4810          strControl = "Date"
4820          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
4830          blnMsgGiven = True
4840        Case Else
4850          strControl = vbNullString
4860        End Select
4870        If blnMsgGiven = False Then
4880          MsgBox "You must enter a numeric value in the " & strControl & " field.", vbInformation + vbOKOnly, "Entry Required"
4890        End If
4900        Response = acDataErrContinue
4910        DoCmd.CancelEvent
4920      Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
            ' ** Ignore.
4930        Response = acDataErrContinue
4940      Case 2237  ' ** The text you entered isn't an item in the list.
4950        Response = acDataErrContinue
4960        DoCmd.CancelEvent
4970      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
4980        DoCmd.Hourglass False
4990        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
5000        Response = acDataErrContinue
5010        DoCmd.CancelEvent
5020      Case 3075  ' ** '|' in query expression '|'.
5030  On Error Resume Next
5040        strControl = Screen.ActiveControl.Name
5050  On Error GoTo ERRH
5060        DoCmd.Hourglass False
5070        If gstrSaleAccountNumber = vbNullString Or gstrSaleAccountNumber = "0" Then
5080          MsgBox "There must be an Account Number to continue.", vbInformation + vbOKOnly, "Entry Required"
5090          If strControl <> "saleAccountno" And blnAccountNoHasFocus = False Then
5100            If Me.saleAccountNo.Enabled = True Then
5110              blnAccountNoHasFocus = True
5120              Me.saleAccountNo.SetFocus
5130            Else
5140              Me.saleTransDate.SetFocus
5150            End If
5160          End If
5170          Exit Sub
5180        End If
5190        If gstrSaleType = vbNullString Then
5200          MsgBox "You must choose one sale type to continue.", vbInformation + vbOKOnly, "Entry Required"
5210          Me.saleType.SetFocus
5220          Exit Sub
5230        End If
5240        If gstrSaleAsset = vbNullString Or gstrSaleAsset = "0" Then
5250          MsgBox "An asset must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
5260          Me.saleAssetno.SetFocus
5270          Exit Sub
5280        End If
5290        If (gstrSaleShareFace = vbNullString Or Val(gstrSaleShareFace) = 0) And Me.saleType <> "Cost Adj." Then
5300          MsgBox "The Share/Face must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
5310          saleShareFace.SetFocus
5320          Exit Sub
5330        End If
5340        If gstrSalePCash = vbNullString Or Val(gstrSalePCash) = 0 And Me.saleType = "Sale" Then
5350          MsgBox "The Principal Cash must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
5360          Me.salePCash.SetFocus
5370          Exit Sub
5380        End If
5390        If gstrSaleCost = vbNullString Or Val(gstrSaleCost) = 0 Then
5400          MsgBox "The Cost must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
5410          Me.saleCost.SetFocus
5420          Exit Sub
5430        End If
5440      Case 3101  ' ** The Microsoft Jet database engine cannot find a record in the table '|' with key matching field(s) '|'.
5450        If IsNull(Me.saleAccountNo) Or Me.saleAccountNo = vbNullString Then
5460          DoCmd.Hourglass False
5470          MsgBox "The posting of this transaction is incomplete." & vbCrLf & vbCrLf & _
                "You must either complete or cancel this transaction to continue.", vbInformation + vbOKOnly, "Incomplete Data"
5480          If Me.cmdSaleClose.Enabled = True And Me.cmdSaleCancel.Enabled = False Then
5490            Me.cmdSaleOK.Enabled = True
5500            Me.cmdSaleCancel.Enabled = True
5510            Me.cmdSaleClose.Enabled = False
5520          End If
5530        End If
5540  On Error Resume Next
5550        strControl = Screen.ActiveControl.Name
5560  On Error GoTo ERRH
5570        If strControl <> "saleAccountno" And blnAccountNoHasFocus = False Then
5580          If Me.saleAccountNo.Enabled = True Then
5590            blnAccountNoHasFocus = True
5600  On Error Resume Next
5610            blnFromFormError = True
5620            Me.saleAccountNo.SetFocus
5630  On Error GoTo ERRH
5640          Else
5650            Me.saleTransDate.SetFocus
5660          End If
5670        End If
5680        Response = acDataErrContinue
5690      Case 3162  ' ** You tried to assign the Null value to a variable that is not a Variant data type.
5700  On Error Resume Next
5710        strControl = Screen.ActiveControl.Name
5720  On Error GoTo ERRH
5730        DoCmd.Hourglass False
5740        MsgBox "The " & strControl & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
5750        Response = acDataErrContinue
5760      Case Else
5770        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
5780        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
5790      End Select
5800      gblnSaleChanged = True
5810    End If
5820    DoCmd.Hourglass False

EXITP:
5830    Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.
' ** VGC 04/13/08: Calendar window seems to be Modal, preventing the
' **               user from doing anything else till it's closed.

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

5910    If Not clsMonthClass Is Nothing Then
5920      If clsMonthClass.IsCalendar Then
            'Cancel = -1
5930      Else
5940        Set clsMonthClass = Nothing
5950      End If
5960    End If

EXITP:
5970    Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

6010    DetailMouse_Sub4 blnCalendar1_Focus, blnCalendar2_Focus, blnSaleReinvest_Focus, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
6020    Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Public Sub cmdSaleClose_Click()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaleClose_Click"

        Dim strDocName As String

6110    glngJournalForm = 0&
6120    gstrFormQuerySpec = vbNullString
6130    gblnSetFocus = True
6140    strDocName = "frmJournal"
6150    DoCmd.Close acForm, strDocName

EXITP:
6160    Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub cmdSaleOK_Click()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaleOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, ctl As Access.Control
        Dim lngThisJrnlID As Long, lngThatJrnlID As Long, lngJrnlMS_ID As Long, lngThisCurrID As Long
        Dim datTransDate As Date
        Dim strShortName As String, dblICash As Double, dblPCash As Double
        Dim curCost As Currency, strRevCode As String, lngTaxcode As Long
        Dim msgResponse As VbMsgBoxResult
        Dim strJrnlUser As String, strMsg As String, strAdvisory As String, strDocName As String
        Dim lngRecs As Long, intPos01 As Integer
        Dim blnTmpGoToSaleReinvest As Boolean
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
        Dim dblTmp05 As Double, dblTmp06 As Double, blnTmp07 As Boolean
        Dim lngX As Long
        Dim blnContinue As Boolean

6210    blnContinue = True

6220    DoCmd.Hourglass True
6230    DoEvents

6240    blnCmdButton = True
6250    blnTmpGoToSaleReinvest = blnGoToSaleReinvest
6260    Form_Current  ' ** Procedure: Above.
6270    blnCmdButton = False

6280    With Me

6290      If VerifyAccountNo = True Then  ' ** Function: Below.

6300        lngCurrID = .saleCurr_ID.Column(CBX_C_CURRID)

            ' ** These seem to be here as a shortcut to full qualification,
            ' ** allowing reference from other modules and forms.
            ' ** I'd like to re-establish them here,
            ' ** along with a full check of their criteria.
6310        gstrSaleType = Nz(.saleType, vbNullString)
6320        gstrSaleAsset = CStr(CLng(Nz(.saleAssetno, 0)))
6330        gstrSaleShareFace = CStr(CDbl(Nz(.saleShareFace, 0)))
6340        gstrSaleAccountNumber = .saleAccountNo
6350        gstrSaleICash = CStr(CDbl(Nz(.saleICash, 0)))
6360        gstrSaleICash = Rem_Dollar(gstrSaleICash, lngCurrID)  ' ** Module Function: modStringFuncs.
6370        gstrSalePCash = CStr(CDbl(Nz(.salePCash, 0)))
6380        gstrSalePCash = Rem_Dollar(gstrSalePCash, lngCurrID)  ' ** Module Function: modStringFuncs.
6390        gstrSaleCost = CStr(CDbl(Nz(.saleCost, 0)))
6400        gstrSaleCost = Rem_Dollar(gstrSaleCost, lngCurrID)  ' ** Module Function: modStringFuncs.
            ' ** gintShareFaceDecimals = 4

            ' ** AccountNo, JournalType, AssetNo.
6410        If gstrSaleAccountNumber = vbNullString And blnAccountNoHasFocus = False Then
6420          blnContinue = False
6430          DoCmd.Hourglass False
6440          MsgBox "Please enter a valid account number.", vbInformation + vbOKOnly, "Entry Required"
6450          blnAccountNoHasFocus = True
6460          .saleAccountNo.SetFocus
6470        Else
6480          If gstrSaleType = vbNullString Then
6490            blnContinue = False
6500            DoCmd.Hourglass False
6510            MsgBox "You must choose a Sold, Withdrawn, Liability, or Cost Adj. Journal Type to continue.", _
                  vbInformation + vbOKOnly, "Entry Required"
6520            .saleType.SetFocus
6530          Else
6540            If gstrSaleAsset = vbNullString Or gstrSaleAsset = "0" Then
6550              blnContinue = False
6560              DoCmd.Hourglass False
6570              MsgBox "An asset must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
6580              .saleAssetno.SetFocus
6590            End If
6600          End If
6610        End If
6620      Else
6630        blnContinue = False
6640      End If

          ' ** AssetDate, TransDate.
6650      If blnContinue = True Then
6660        If IsNull(.saleAssetDate) = True Then
6670          blnContinue = False
6680          DoCmd.Hourglass False
6690          MsgBox "You must enter a Trade Date.", vbInformation + vbOKOnly, "Entry Required"
6700          .saleAssetDate.SetFocus
6710        Else
6720          .saleAssetDate.SetFocus
6730          If Len(.saleAssetDate.text) = 0 Then
6740            blnContinue = False
6750            DoCmd.Hourglass False
6760            MsgBox "You must enter a Trade Date.", vbInformation + vbOKOnly, "Entry Required"
6770          Else
6780            .cmdSaleOK.SetFocus
6790            If Compare_DateA_DateB(CDate(.saleAssetDate), ">", Date) = True Then  ' ** Module Function: modStringFuncs.
6800              blnContinue = False
6810              DoCmd.Hourglass False
6820              MsgBox "Future trade dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
6830              .saleAssetDate = CDate(Format(Now(), "mm/dd/yyyy"))
6840            Else
6850              If IsNull(.saleTransDate) = True Then
6860                blnContinue = False
6870                DoCmd.Hourglass False
6880                MsgBox "You must enter a Posting Date.", vbInformation + vbOKOnly, "Entry Required"
6890                .saleTransDate.SetFocus
6900              Else
6910                .saleTransDate.SetFocus
6920                If Len(.saleTransDate.text) = 0 Then
6930                  blnContinue = False
6940                  DoCmd.Hourglass False
6950                  MsgBox "You must enter a Posting Date.", vbInformation + vbOKOnly, "Entry Required"
6960                Else
6970                  .cmdSaleOK.SetFocus
6980                  If .saleTransDate > DateAdd("m", 1, Date) Then
6990                    blnContinue = False
7000                    DoCmd.Hourglass False
7010                    MsgBox "Only future Posting Dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
7020                    .saleTransDate.SetFocus
7030                  End If
7040                End If
7050              End If
7060            End If
7070          End If
7080        End If
7090      End If  ' ** blnContinue.

          ' ** Description.
7100      If blnContinue = True Then
7110        .saleDescription.SetFocus
7120        If InStr(.saleDescription.text, Chr(34)) > 0 Then
7130          blnContinue = False
7140          DoCmd.Hourglass False
7150          MsgBox "Cannot have quote marks in text.", vbInformation + vbOKOnly, "Invalid Characters"
7160          .saleDescription.SetFocus
7170        End If
7180      End If  ' ** blnContinue.

          ' ** ShareFace, ICash, PCash, Cost.
7190      If blnContinue = True Then

7200        strMsg = vbNullString
7210        Set ctl = Nothing
7220        If Abs(Val(gstrSaleShareFace)) > 999999999.9999 Then
7230          strMsg = "Share/Face is too large."
7240          Set ctl = .saleShareFace
7250        ElseIf Abs(Val(gstrSaleICash)) > 999999999.99 Then
7260          strMsg = "Accrued/Income Cash is too large."
7270          Set ctl = .saleICash
7280        ElseIf Abs(Val(gstrSalePCash)) > 999999999.99 Then
7290          strMsg = "Principal Cash is too large."
7300          Set ctl = .salePCash
7310        ElseIf Abs(Val(gstrSaleCost)) > 999999999.99 Then
7320          strMsg = "Cost is too large."
7330          Set ctl = .saleCost
7340        End If
7350        If strMsg <> vbNullString Then
7360          blnContinue = False
7370          DoCmd.Hourglass False
7380          MsgBox strMsg, vbInformation + vbOKOnly, "Invalid Entry"
7390          ctl.SetFocus
7400        Else

7410          strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString
7420          intPos01 = InStr(gstrSaleShareFace, ".")
7430          If intPos01 > 0 Then
7440            If Len(Mid(gstrSaleShareFace, (intPos01 + 1))) > gintShareFaceDecimals Then
7450              strTmp01 = "Share/Face will be rounded to " & Trim(CStr(gintShareFaceDecimals)) & " decimals."
7460              gstrSaleShareFace = Left(gstrSaleShareFace, intPos01) & _
                    Left(Mid(gstrSaleShareFace, (intPos01 + 1)), gintShareFaceDecimals)
7470              gstrSaleShareFace = CDbl(Val(gstrSaleShareFace))  ' ** In case it's now Zero!
7480              .saleShareFace = CDbl(gstrSaleShareFace)
7490            End If
7500          End If
7510          intPos01 = InStr(gstrSaleICash, ".")
7520          If intPos01 > 0 And blnContinue = True Then
7530            If Len(Mid(gstrSaleICash, (intPos01 + 1))) > 2 Then
7540              strTmp02 = "Accrued/Income Cash will be rounded to 2 decimals."
7550              gstrSaleICash = Left(gstrSaleICash, intPos01) & Left(Mid(gstrSaleICash, (intPos01 + 1)), 2)
7560              gstrSaleICash = CDbl(Val(gstrSaleICash))  ' ** In case it's now Zero!
7570              gstrSaleICash = Rem_Dollar(gstrSaleICash, lngCurrID)  ' ** Module Function: modStringFuncs.
7580              .saleICash = CCur(gstrSaleICash)
7590            End If
7600          End If
7610          intPos01 = InStr(gstrSalePCash, ".")
7620          If intPos01 > 0 And blnContinue = True Then
7630            If Len(Mid(gstrSalePCash, (intPos01 + 1))) > 2 Then
7640              strTmp03 = "Principal Cash will be rounded to 2 decimals."
7650              gstrSalePCash = Left(gstrSalePCash, intPos01) & Left(Mid(gstrSalePCash, (intPos01 + 1)), 2)
7660              gstrSalePCash = CDbl(Val(gstrSalePCash))  ' ** In case it's now Zero!
7670              gstrSalePCash = Rem_Dollar(gstrSalePCash, lngCurrID)  ' ** Module Function: modStringFuncs.
7680              .salePCash = CCur(gstrSalePCash)
7690            End If
7700          End If
7710          intPos01 = InStr(gstrSaleCost, ".")
7720          If intPos01 > 0 And blnContinue = True Then
7730            If Len(Mid(gstrSaleCost, (intPos01 + 1))) > 2 Then
7740              strTmp04 = "Cost will be rounded to 2 decimals."
7750              gstrSaleCost = Left(gstrSaleCost, intPos01) & Left(Mid(gstrSaleCost, (intPos01 + 1)), 2)
7760              gstrSaleCost = CDbl(Val(gstrSaleCost))  ' ** In case it's now Zero!
7770              gstrSaleCost = Rem_Dollar(gstrSaleCost, lngCurrID)  ' ** Module Function: modStringFuncs.
7780              .saleCost = CCur(gstrSaleCost)
7790            End If
7800          End If
7810          strAdvisory = strTmp01 & "~" & strTmp02 & "~" & strTmp03 & "~" & strTmp04
7820          If strAdvisory = "~~~" Then strAdvisory = vbNullString
7830          intPos01 = 0: strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString

7840          If Val(gstrSaleShareFace) = 0 And gstrSaleType <> "Cost Adj." Then
7850            blnContinue = False
7860            DoCmd.Hourglass False
7870            If InStr(strAdvisory, "Share/Face") > 0 Then
7880              MsgBox "Share/Face has been rounded to " & Trim(CStr(gintShareFaceDecimals)) & " decimals," & vbCrLf & _
                    "and its value is now zero. Share/Face must be greater than zero for a " & gstrSaleType & ".", _
                    vbInformation + vbOKOnly, "Entry Required"
7890            Else
7900              MsgBox "Share/Face must be greater than zero for a " & gstrSaleType & ".", vbInformation + vbOKOnly, "Entry Required"
7910            End If
7920            .saleShareFace.SetFocus
7930          ElseIf gstrSaleShareFace <> "0" And gstrSaleType = "Cost Adj." Then
7940            blnContinue = False
7950            DoCmd.Hourglass False
7960            MsgBox "Share/Face must be zero for a " & gstrSaleType, vbInformation + vbOKOnly, "Invalid Entry"
7970            .saleShareFace.SetFocus
7980          Else

7990            Select Case gstrSaleType
                Case "Withdrawn"
8000              If gstrSaleICash <> "0" Then  'Zero icash
8010                blnContinue = False
8020                DoCmd.Hourglass False
8030                MsgBox "Accrued/Income Cash must be zero for a " & gstrSaleType & ".", vbInformation + vbOKOnly, "Invalid Entry"
8040                .saleICash.SetFocus
8050              Else
8060                If gstrSalePCash <> "0" Then  'Zero pcash
8070                  blnContinue = False
8080                  DoCmd.Hourglass False
8090                  MsgBox "Principal Cash must be zero for a " & gstrSaleType & ".", vbInformation + vbOKOnly, "Invalid Entry"
8100                  .salePCash.SetFocus
8110                Else
8120                  If Val(gstrSaleCost) > 0 Then  'Negative or Zero cost  CHECK! YES, MESSAGE BELOW ALLOWS ZERO COST!
8130                    blnContinue = False
8140                    DoCmd.Hourglass False
8150                    MsgBox "A " & gstrSaleType & " cannot have a positive Cost.", vbInformation + vbOKOnly, "Invalid Entry"
8160                    .saleCost.SetFocus
8170                  End If
8180                End If
8190              End If
8200            Case "Sold"
8210              If gstrSaleICash <> "0" Then  '(Positive pcash, Positive icash)
8220                If Val(gstrSaleICash) < 0 Then
8230                  blnContinue = False
8240                  DoCmd.Hourglass False
8250                  MsgBox "Accrued/Income Cash must be positive for a " & gstrSaleType & ".", vbInformation + vbOKOnly, "Invalid Entry"
8260                  .saleICash.SetFocus
8270                End If
8280              End If
8290              If Val(gstrSalePCash) = 0 Then
8300                If blnWarned = False Then
8310                  If Val(gstrSaleICash) = 0 Then
8320                    DoCmd.Hourglass False
8330                    msgResponse = MsgBox("Are you sure you want Income and Principal Cash to be ZERO?" & vbCrLf & vbCrLf & _
                          "As would be the case for the sale of worthless shares.", vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
8340                  Else
8350                    DoCmd.Hourglass False
8360                    msgResponse = MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                          vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
8370                  End If
8380                  blnWarned = True
8390                  If msgResponse <> vbYes Then
8400                    blnContinue = False
8410                    DoCmd.Hourglass False
8420                    .salePCash.SetFocus
8430                  Else
8440                    DoCmd.Hourglass True
8450                    DoEvents
8460                  End If
8470                End If
8480              Else
8490                If Val(gstrSalePCash) < 0 Then  '(Postive pcash, Zero icash)
8500                  blnContinue = False
8510                  DoCmd.Hourglass False
8520                  MsgBox gstrSaleType & " Principal Cash cannot be negative.", vbInformation + vbOKOnly, "Invalid Entry"
8530                  .salePCash.SetFocus
8540                End If
8550              End If
8560              If Val(gstrSaleCost) > 0 Then  'Negative or Zero cost  CHECK! YES, MESSAGE BELOW ALLOWS ZERO COST.
8570                blnContinue = False
8580                DoCmd.Hourglass False
8590                MsgBox gstrSaleType & " Cost cannot be positive.", vbInformation + vbOKOnly, "Invalid Entry"
8600                .saleCost.SetFocus
8610              End If
8620            Case "Liability"  'LiabilitySW
8630              If Val(gstrSaleICash) > 0 Then  '(Negative pcash, Negative icash)
8640                blnContinue = False
8650                DoCmd.Hourglass False
8660                MsgBox "A " & gstrSaleType & " cannot have positive Accrued/Income Cash.", vbInformation + vbOKOnly, "Invalid Entry"
8670                .saleICash.SetFocus
8680              Else
8690                If Val(gstrSalePCash) > 0 Then  '(Negative pcash, Zero icash) MESSAGE BELOW ALLOWS ZERO PCASH!
8700                  blnContinue = False
8710                  DoCmd.Hourglass False
8720                  MsgBox "A " & gstrSaleType & " cannot have positive Principal Cash.", vbInformation + vbOKOnly, "Invalid Entry"
8730                  .salePCash.SetFocus
8740                Else
8750                  If Val(gstrSaleCost) < 0 Then  'Positive or Zero cost  CHECK! YES, MESSAGE BELOW ALLOWS ZERO COST.
8760                    blnContinue = False
8770                    DoCmd.Hourglass False
8780                    MsgBox gstrSaleType & " Cost cannot be negative.", vbInformation + vbOKOnly, "Invalid Entry"
8790                    .saleCost.SetFocus
8800                  End If
8810                End If
8820              End If
8830            Case "Cost Adj."
8840              If gstrSaleICash <> "0" Then  'Zero icash
8850                blnContinue = False
8860                DoCmd.Hourglass False
8870                MsgBox "A " & gstrSaleType & " must have zero Accrued/Income Cash.", vbInformation + vbOKOnly, "Invalid Entry"
8880                .saleICash.SetFocus
8890              Else
8900                If gstrSalePCash <> "0" Then  'Zero pcash
8910                  blnContinue = False
8920                  DoCmd.Hourglass False
8930                  MsgBox "A " & gstrSaleType & " must have zero Principal Cash.", vbInformation + vbOKOnly, "Invalid Entry"
8940                  .salePCash.SetFocus
8950                Else
8960                  If Val(gstrSaleCost) = 0 Then  'Positive or Negative cost. MESSAGE BELOW SAYS COST CANNOT BE ZERO.
8970                    blnContinue = False
8980                    DoCmd.Hourglass False
8990                    MsgBox "A " & gstrSaleType & " cannot have zero Cost.", vbInformation + vbOKOnly, "Invalid Entry"
9000                    .saleCost.SetFocus
9010                  End If
9020                End If
9030              End If
9040            End Select

9050          End If
9060        End If
9070      End If  ' ** blnContinue.

          ' ** THE SECTION BELOW MAY BE MOOT!
9080      If blnContinue = True Then
9090        If (gstrSaleShareFace = vbNullString Or CDbl(IIf(gstrSaleShareFace = vbNullString, "0", gstrSaleShareFace)) = 0#) And _
                .saleType <> "Cost Adj." Then
9100          blnContinue = False
9110          DoCmd.Hourglass False
9120          MsgBox "Share/Face must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
9130          saleShareFace.SetFocus
9140        Else
9150          If (gstrSalePCash = vbNullString Or Val(gstrSalePCash) = 0) And .saleType = "Sold" And blnWarned = False Then
9160            If (gstrSaleICash = vbNullString Or Val(gstrSaleICash) = 0) Then
9170              DoCmd.Hourglass False
9180              msgResponse = MsgBox("Are you sure you want Income and Principal Cash to be ZERO?" & vbCrLf & vbCrLf & _
                    "As would be the case for the sale of worthless shares.", vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
9190            Else
9200              DoCmd.Hourglass False
9210              msgResponse = MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
9220            End If
9230            If msgResponse <> vbYes Then
9240              blnContinue = False
9250              DoCmd.Hourglass False
9260              .salePCash.SetFocus
9270            Else
9280              DoCmd.Hourglass True
9290              DoEvents
9300            End If
9310          End If
9320          If blnContinue = True Then
9330            If (.saleType <> "Cost Adj.") And (.saleType <> "Withdrawn") Then
9340              If .saleType <> "Liability" Then
                    ' ** Sold.
9350                If CDbl(IIf(gstrSalePCash = vbNullString, "0", gstrSalePCash)) < 0# Then
9360                  blnContinue = False
9370                  DoCmd.Hourglass False
9380                  MsgBox .saleType & " Principal Cash cannot be negative.", vbInformation + vbOKOnly, "Invalid Entry"
9390                  .salePCash.SetFocus
9400                End If
9410              Else
                    ' ** Liability.
9420                If CDbl(IIf(gstrSalePCash = vbNullString, "0", gstrSalePCash)) = 0# Then
9430                  DoCmd.Hourglass False
9440                  If MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                          vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis") <> vbYes Then
9450                    blnContinue = False
9460                    .salePCash.SetFocus
9470                  Else
9480                    DoCmd.Hourglass True
9490                    DoEvents
9500                  End If
9510                End If
9520              End If
9530            End If
9540          End If
9550        End If
9560      End If  ' ** blnContinue.

9570      If blnContinue = True Then
9580        If (.saleType = "Cost Adj.") Then
9590          If .saleCost = 0 Then
9600            blnContinue = False
9610            DoCmd.Hourglass False
9620            MsgBox "Cannot enter a zero amount Cost Adjustment.", vbInformation + vbOKOnly, "Invalid Entry"
9630            .saleCost.SetFocus
9640          End If
9650        End If
9660      End If  ' ** blnContinue.

9670      If blnContinue = True Then
9680        If (gstrSaleICash = vbNullString) And .saleType = "Sold" Then
9690          blnContinue = False
9700          DoCmd.Hourglass False
9710          MsgBox "Income Cash must have a value.", vbInformation + vbOKOnly, "Invalid Entry"
9720          .saleICash.SetFocus
9730        Else
9740          If (.saleType <> "Cost Adj.") And (.saleType <> "Withdrawn") Then
                ' ** What's left are 'Sold', 'Liability'.
9750            If (gstrSaleICash = vbNullString) And (.saleType <> "Liability") Then
                  ' ** Sold.
9760              blnContinue = False
9770              DoCmd.Hourglass False
9780              MsgBox "Income Cash must have a value.", vbInformation + vbOKOnly, "Invalid Entry"
9790              .saleICash.SetFocus
9800            Else
9810              If (CDbl(IIf(gstrSaleICash = vbNullString, "0", gstrSaleICash)) < 0#) And (.saleType <> "Liability") Then
                    ' ** Sold.
9820                blnContinue = False
9830                DoCmd.Hourglass False
9840                MsgBox "Income Cash must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
9850                .saleICash.SetFocus
9860              End If
9870            End If
9880          End If
9890        End If
9900      End If  ' ** blnContinue.

9910      If blnContinue = True Then
9920        If IsNull(.UnitPurchaseDate) And (.saleType <> "Cost Adj.") Then
9930          blnContinue = False
9940          DoCmd.Hourglass False
9950          .saleCost.SetFocus
9960  On Error Resume Next
9970          SendKeys "{ENTER}", True
9980          If ERR.Number <> 0 Then
                ' ** 70  Permission denied.
9990  On Error GoTo ERRH
10000         Else
10010 On Error GoTo ERRH
10020         End If
10030       Else
10040         If (gstrSaleCost = vbNullString Or Val(gstrSaleCost) = 0) And (.saleType <> "Cost Adj.") Then
10050           DoCmd.Hourglass False
10060           If MsgBox("Are you sure you want this " & .saleType & " to have ZERO Cost?", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cost Basis") = vbNo Then
10070             blnContinue = False
10080             .saleCost.SetFocus
10090           Else
10100             DoCmd.Hourglass True
10110             DoEvents
10120           End If
10130         End If
10140       End If
10150     End If  ' ** blnContinue.

10160     If blnContinue = True Then

10170       DoCmd.Hourglass True  ' ** Make sure it's still running.
10180       DoEvents

10190       lngThisJrnlID = .saleID
10200       gstrAccountNo = .saleAccountNo.Value
10210       strShortName = .saleShortName.Value
10220       datTransDate = .saleTransDate.Value
10230       dblICash = Nz(.saleICash, 0)
10240       dblPCash = Nz(.salePCash, 0)
10250       glngAssetNo = .saleAssetno.Column(CBX_A_ASTNO)
10260       lngThisCurrID = .saleCurr_ID.Column(CBX_C_CURRID)

10270       If .saleType = "Cost Adj." Then
10280         SaleDistribute  ' ** Procedure: Below.
10290       End If

10300       DoCmd.Hourglass True  ' ** Make sure it's still running.
10310       DoEvents

10320       Set dbs = CurrentDb
            ' ** qryJournal_Sale_01, linked to MasterAsset; check for existing Sales/Withdrawals, by specified [actno], [astno], [jrnlid].
10330       Set qdf = dbs.QueryDefs("qryJournal_Sale_02")  ' ** This excludes Asset Types '60','80', and '81',
10340       With qdf.Parameters                            ' ** meaning they CAN enter multiple Sold!
10350         ![jrnlid] = lngThisJrnlID
10360         ![actno] = gstrAccountNo
10370         ![astno] = glngAssetNo
10380       End With
10390       Set rst = qdf.OpenRecordset()
10400       If rst.BOF = True And rst.EOF = True Then
              ' ** No records; all's well.
10410         If .saleCurr_Date.Visible = True Then .saleCurr_Date.Visible = False
10420         rst.Close
10430         Set rst = Nothing
10440         Set qdf = Nothing
10450         dbs.Close
10460         Set dbs = Nothing
10470       Else
10480         blnContinue = False
10490         strMsg = vbNullString
10500         rst.MoveLast
10510         lngRecs = rst.RecordCount
10520         rst.MoveFirst
10530         If lngRecs > 1& Then
10540           For lngX = 1& To lngRecs
10550             If IsNull(rst![journal_USER]) = True Then
10560               strJrnlUser = "{unknown}"
10570             Else
10580               If Trim(rst![journal_USER]) = vbNullString Then
10590                 strJrnlUser = "{unknown}"
10600               Else
10610                 strJrnlUser = rst![journal_USER]
10620               End If
10630             End If
10640             strMsg = strMsg & strJrnlUser & ", "
10650             If lngX < lngRecs Then rst.MoveNext
10660           Next
10670           strMsg = Trim(strMsg)                    ' ** Remove trailing space.
10680           strMsg = Left(strMsg, (Len(strMsg) - 1))  ' ** Remove trailing comma.
10690         Else
10700           If IsNull(rst![journal_USER]) = True Then
10710             strJrnlUser = "{unknown}"
10720           Else
10730             If Trim(rst![journal_USER]) = vbNullString Then
10740               strJrnlUser = "{unknown}"
10750             Else
10760               strJrnlUser = rst![journal_USER]
10770             End If
10780           End If
10790           strMsg = strJrnlUser
10800         End If
10810         strMsg = IIf(strMsg = CurrentUser, "You have", "The user" & IIf(lngRecs = 1, " " & strMsg & " has", "s " & strMsg & " have")) & _
                " already entered a sale/withdrawal for this asset for this account!" & vbCrLf & vbCrLf & _
                "You must post or delete the existing transaction before adding another one."
              ' ** Internal Access Function: Trust Accountant login.
10820         DoCmd.Hourglass False
10830         MsgBox strMsg, vbExclamation + vbOKOnly, "Multiple Sales/Withdrawals"
10840         rst.Close
10850         Set rst = Nothing
10860         Set qdf = Nothing
10870         dbs.Close
10880         Set dbs = Nothing
10890       End If

10900     End If  ' ** blnContinue.

10910     If blnContinue = True Then

10920       DoCmd.Hourglass True  ' ** Make sure it's still running.
10930       DoEvents

            ' ** VGC 11/19/2009: Just 'cause I'd like to see it there.
10940       blnTmp07 = False
10950       If IsNull(.pershare) = True Then
10960         blnTmp07 = True
10970       Else
10980         If .pershare = 0 Then
10990           blnTmp07 = True
11000         End If
11010       End If
11020       If blnTmp07 = True Then
11030         If Nz(.saleCost, 0) < 0 Then
11040           dblTmp05 = -(CDbl(.saleCost))
11050           If Nz(.saleShareFace, 0) > 0 Then
11060             dblTmp06 = CDbl(.saleShareFace)
11070             .pershare = (dblTmp05 / dblTmp06)
11080           End If
11090         End If
11100       End If

11110       If Trim(Nz(.saleUser, vbNullString)) = vbNullString Then
11120         .saleUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
11130       End If

11140       curCost = Nz(.saleCost, 0@)
11150       strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
11160       lngTaxcode = Nz(.cmbTaxCodes, 0&)
11170       Select Case gstrSaleType
            Case "Sold"
              ' ** INCOME.
11180         If strRevCode <> "I" Then
11190           .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
11200         End If
11210         If IsNull(.cmbTaxCodes) = True Then
11220           .cmbTaxCodes = 0&
11230         Else
11240           If gblnLinkRevTaxCodes = True Then
11250             If lngTaxcode = 0& Then
11260               .cmbTaxCodes = .saleAssetno.Column(CBX_A_TAX)  ' ** All AssetType-based Tax Codes are INCOME.
11270             Else
11280               If .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
11290                 .cmbTaxCodes = .saleAssetno.Column(CBX_A_TAX)  ' ** All AssetType-based Tax Codes are INCOME.
11300               End If
11310             End If
11320           End If
11330         End If
11340       Case "Withdrawn", "Cost Adj."
11350         If lngTaxcode > 0& Then  ' ** Tax Code takes precedence over Revenue Code.
11360           If gblnLinkRevTaxCodes = True Then
11370             If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
11380               If IsNull(.cmbRevenueCodes) = True Then
11390                 .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
11400               Else
11410                 If .cmbRevenueCodes = 0 Then
11420                   .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
11430                 Else
11440                   If strRevCode <> "I" Then
11450                     .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
11460                   End If
11470                 End If
11480               End If
11490             ElseIf .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
11500               If IsNull(.cmbRevenueCodes) = True Then
11510                 .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
11520               Else
11530                 If .cmbRevenueCodes = 0 Then
11540                   .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
11550                 Else
11560                   If strRevCode <> "E" Then
11570                     .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
11580                   End If
11590                 End If
11600               End If
11610             End If
11620           Else
11630             If IsNull(.cmbRevenueCodes) = True Then
11640               .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
11650             Else
11660               If .cmbRevenueCodes = 0 Then
11670                 .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
11680               End If
11690             End If
11700           End If
11710         Else
11720           If IsNull(.cmbTaxCodes) = True Then
11730             .cmbTaxCodes = 0&
11740           End If
11750           If IsNull(.cmbRevenueCodes) = True Then
11760             .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
11770           Else
11780             If .cmbRevenueCodes = 0& Then
11790               .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
11800             End If
11810           End If
11820         End If
11830       Case "Liability"
              ' ** EXPENSE.
11840         If strRevCode <> "E" Then
11850           .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
11860         End If
11870         If IsNull(.cmbTaxCodes) = True Then
11880           .cmbTaxCodes = 0&
11890         Else
11900           If gblnLinkRevTaxCodes = True Then
11910             If lngTaxcode = 0& Then
11920               .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
11930             Else
11940               If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
11950                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
11960               End If
11970             End If
11980           End If
11990         End If
12000       End Select

12010       If .Parent.DefSaleBtn = True Then
12020         If IsNull(.Parent.DefPurchJrnlID) = False And IsNull(.Parent.DefPurchJrnlMSID) = False Then
12030           lngThatJrnlID = .Parent.DefPurchJrnlID
12040           lngJrnlMS_ID = .Parent.DefPurchJrnlMSID
                ' ** Update the parent's tblJournal_MiscSold record and add this one.
12050           Set dbs = CurrentDb
12060           Set rst = dbs.OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
12070           With rst
12080             .MoveFirst
12090             .FindFirst "[jrnlms_id] = " & CStr(lngJrnlMS_ID)
12100             If .NoMatch = False Then
12110               .Edit
12120               ![jrnlms_ref_id] = lngThisJrnlID
12130               ![jrnlms_datemodified] = Now()
12140               .Update
12150               .AddNew
                    ' ** ![jrnlms_id] : AutoNumber.
12160               ![Journal_ID] = lngThisJrnlID
12170               ![journaltype] = "Sold"
12180               ![jrnlms_ref_id] = lngThatJrnlID
12190               ![jrnlms_parent] = False
12200               ![jrnlms_child] = True
12210               ![jrnlms_note] = "This is a covering Sold transaction, generated by a Purchase entry."
12220               ![jrnlms_datemodified] = Now()
12230               .Update
12240             End If
12250             .Close
12260           End With
12270           Set rst = Nothing
12280           dbs.Close
12290           Set dbs = Nothing
12300           DoEvents
12310         ElseIf IsNull(.Parent.DefPaidJrnlID) = False And IsNull(.Parent.DefPaidJrnlMSID) = False Then
12320           lngThatJrnlID = .Parent.DefPaidJrnlID
12330           lngJrnlMS_ID = .Parent.DefPaidJrnlMSID
                ' ** Update the parent's tblJournal_MiscSold record and add this one.
12340           Set dbs = CurrentDb
12350           Set rst = dbs.OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
12360           With rst
12370             .MoveFirst
12380             .FindFirst "[jrnlms_id] = " & CStr(lngJrnlMS_ID)
12390             If .NoMatch = False Then
12400               .Edit
12410               ![jrnlms_ref_id] = lngThisJrnlID
12420               ![jrnlms_datemodified] = Now()
12430               .Update
12440               .AddNew
                    ' ** ![jrnlms_id] : AutoNumber.
12450               ![Journal_ID] = lngThisJrnlID
12460               ![journaltype] = "Sold"
12470               ![jrnlms_ref_id] = lngThatJrnlID
12480               ![jrnlms_parent] = False
12490               ![jrnlms_child] = True
12500               ![jrnlms_note] = "This is a covering Sold transaction, generated by a Paid entry."
12510               ![jrnlms_datemodified] = Now()
12520               .Update
12530             End If
12540             .Close
12550           End With
12560           Set rst = Nothing
12570           dbs.Close
12580           Set dbs = Nothing
12590           DoEvents
12600         End If
12610       End If  ' ** DefSaleBtn.

12620       If .saleCurr_Date.Visible = True Then .saleCurr_Date.Visible = False

12630       If blnTmpGoToSaleReinvest = True Then
12640         .saleReinvested = True
12650       End If

12660       DoCmd.RunCommand acCmdSaveRecord

12670       DoCmd.Hourglass True  ' ** Make sure it's still running.
12680       DoEvents

12690       .Requery

12700       blnSoldCosted = False
12710       SoldCostedSet False  ' ** Procedure: Below.

12720       .AllowAdditions = True
12730       DoCmd.GoToRecord , , acNewRec

12740       Forms("frmJournal").DefSaleBtn = False
12750       Forms("frmJournal").DefAssetNo = 0&
12760       Forms("frmJournal").DefPurchTot = 0@
12770       Forms("frmJournal").DefPurchJrnlID = Null
12780       Forms("frmJournal").DefPurchJrnlMSID = Null
12790       Forms("frmJournal").DefPaidTot = 0@
12800       Forms("frmJournal").DefPaidJrnlID = Null
12810       Forms("frmJournal").DefPaidJrnlMSID = Null

12820       .cmdSaleLotInfo.Enabled = False
12830       If .saleAccountNo.Enabled = True Then
12840         .saleAccountNo.SetFocus
12850       Else
12860         .saleTransDate.SetFocus
12870       End If

12880       gblnSaleValidated = False

12890       SaleChanged False  ' ** Procedure: Below.

12900       DoCmd.Hourglass True  ' ** Make sure it's still running.
12910       DoEvents

12920       blnGoToSaleReinvest = blnTmpGoToSaleReinvest

            ' ** Reinvest button was pushed.
12930       If blnGoToSaleReinvest = True Then

              ' ** Put preliminary record into tblJournal_MiscSold.
12940         Set dbs = CurrentDb
12950         Set rst = dbs.OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
12960         With rst
12970           If .BOF = True And .EOF = True Then
                  ' ** Proceed.
12980           Else
12990             .MoveFirst
13000             .FindFirst "[Journal_ID] = " & CStr(lngThisJrnlID)
13010             If .NoMatch = False Then
13020               .Delete
13030             End If
13040           End If
13050           DoEvents
13060           .AddNew
                ' ** ![jrnlms_id] : AutoNumber.
13070           ![Journal_ID] = lngThisJrnlID
13080           ![journaltype] = "Sold"
13090           ![jrnlms_ref_id] = 0&
13100           ![jrnlms_parent] = True
13110           ![jrnlms_child] = False
13120           ![jrnlms_note] = "This Sold was reinvested, generating a Purchase transaction."
13130           ![jrnlms_datemodified] = Now()
13140           .Update
13150           .Bookmark = .LastModified
13160           lngJrnlMS_ID = ![jrnlms_id]
13170         End With
13180         Set rst = Nothing
13190         dbs.Close
13200         Set dbs = Nothing
13210         DoEvents

13220         gblnMessage = True  ' ** If this returns False, the dialog was canceled.
13230         gdblCrtRpt_CostTot = 0#  ' ** Borrowing this variable from the Court Reports.
13240         strDocName = "frmReinvest_Sold"
13250         DoCmd.OpenForm strDocName, , , , acFormAdd, acDialog, strCallingForm & "~" & gstrAccountNo & "~" & _
                Format(datTransDate, "mm/dd/yyyy") & "~" & CStr(lngThisCurrID) & "~" & CStr(dblICash) & "~" & _
                CStr(dblPCash) & "~" & CStr(glngAssetNo) & "~" & CStr(lngThisJrnlID) & "~" & strShortName

13260         DoCmd.Hourglass True
13270         DoEvents

13280         Set dbs = CurrentDb
13290         Select Case gblnMessage
              Case True
13300           Set rst = dbs.OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
13310           With rst
13320             .MoveFirst
13330             .FindFirst "[jrnlms_id] = " & CStr(lngJrnlMS_ID)
13340             If .NoMatch = False Then
13350               .Edit
13360               ![jrnlms_ref_id] = CLng(gdblCrtRpt_CostTot)  ' ** New Purchase ID from frmReinvest_Dividend.
13370               ![jrnlms_datemodified] = Now()
13380               .Update
13390             End If
13400             .AddNew
                  ' ** ![jrnlms_id] : AutoNumber.
13410             ![Journal_ID] = CLng(gdblCrtRpt_CostTot)
13420             ![journaltype] = "Purchase"
13430             ![jrnlms_ref_id] = lngThisJrnlID
13440             ![jrnlms_parent] = False
13450             ![jrnlms_child] = True
13460             ![jrnlms_note] = "This Purchase was generated by a Sold reinvestment."
13470             ![jrnlms_datemodified] = Now()
13480             .Update
13490             .Close
13500           End With
13510           Set rst = Nothing
13520           If .cmdSaleOK.Enabled = True Then
13530             .cmdSaleOK.SetFocus
13540           Else
13550             .saleAccountNo.SetFocus
13560           End If
13570         Case False
                ' ** Delete tblJournal_MiscSold, by specified [jrnlmsid].
13580           Set qdf = dbs.QueryDefs("qryJournal_MiscSold_01")
13590           With qdf.Parameters
13600             ![jrnlmsid] = lngJrnlMS_ID
13610           End With
13620           qdf.Execute
13630           Set qdf = Nothing
13640           DoEvents
                ' ** Update Journal, by specified [jrnlid], [pstd].
13650           Set qdf = dbs.QueryDefs("qryJournal_Sale_19")
13660           With qdf.Parameters
13670             ![jrnlid] = lngThisJrnlID  ' ** This new Sold record.
13680             ![pstd] = False
13690           End With
13700           qdf.Execute
13710           Set qdf = Nothing
13720           If gdblCrtRpt_CostTot <> 0# Then
                  ' ** Delete Journal, by specified [jrnlid].
13730             Set qdf = dbs.QueryDefs("qryJournal_Sale_18")
13740             With qdf.Parameters
13750               ![jrnlid] = CLng(gdblCrtRpt_CostTot)  ' ** New Purchase from frmReinvest_Sold.
13760             End With
13770             qdf.Execute dbFailOnError
13780             Set qdf = Nothing
13790             gdblCrtRpt_CostTot = 0#
13800           End If
13810           .saleAccountNo.SetFocus
13820         End Select
13830         dbs.Close
13840         Set dbs = Nothing
13850       End If
13860       blnGoToSaleReinvest = False  ' ** reset always.

13870     End If  ' ** blnContinue.

13880     DoCmd.Hourglass False

13890   End With

EXITP:
13900   Set ctl = Nothing
13910   Set rst = Nothing
13920   Set qdf = Nothing
13930   Set dbs = Nothing
13940   Exit Sub

ERRH:
430     DoCmd.Hourglass False
440     THAT_PROC = THIS_PROC
450     That_Erl = Erl
460     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
470     Resume EXITP

End Sub

Public Sub cmdSaleCancel_Click()

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaleCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngThisJrnlID As Long, lngThatJrnlID As Long
        Dim blnReinvested As Boolean, strMsg As String
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

14010   With Me

14020     DoCmd.Hourglass True
14030     DoEvents

14040     If IsNull(.saleID) = True Then
14050       msgResponse = vbCancel
14060     Else

14070       blnReinvested = .saleReinvested
14080       lngThisJrnlID = .saleID
14090       lngThatJrnlID = 0&
14100       msgResponse = 0
            ' ** 2  vbCancel
            ' ** 6  vbYes
            ' ** 7  vbNo

            ' ** Check whether this spawned a Purchase using the Reinvest button.
14110       If blnReinvested = True And .posted = True Then
              ' ** Check if the referencing entry still exists.
14120         varTmp00 = DCount("*", "journal", "[CheckNum] = " & CStr(lngThisJrnlID))
14130         If varTmp00 = 1 Then
14140           lngThatJrnlID = DLookup("[ID]", "journal", "[CheckNum] = " & CStr(lngThisJrnlID))
14150           strMsg = "This entry was used for a Purchase reinvestment." & vbCrLf & vbCrLf & _
                  "Yes - Delete both this Sold and the Purchase reinvestment." & vbCrLf & _
                  "No - Delete only this Sold." & vbCrLf & _
                  "Cancel - Delete no transactions."
14160           Beep
14170           DoCmd.Hourglass False
14180           msgResponse = MsgBox(strMsg, vbQuestion + vbYesNoCancel + vbDefaultButton3, "Delete Both Transactions")
14190           If msgResponse <> vbCancel Then
14200             DoCmd.Hourglass True
14210             If .saleCurr_Date.Visible = True Then .saleCurr_Date.Visible = False
14220             DoEvents
14230           End If
14240         Else
14250           lngThatJrnlID = 0&
14260         End If
14270       End If
14280     End If

14290     If msgResponse = 0 Then
14300       Select Case blnGoneToReport
            Case True
14310         msgResponse = vbNo  ' ** Delete only Sold.
14320       Case False
14330         DoCmd.Hourglass False
14340         msgResponse = MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo, "Delete Transaction")
14350         If msgResponse <> vbYes Then
14360           msgResponse = vbCancel
14370         Else
14380           msgResponse = vbNo  ' ** Delete only Sold.
14390           DoCmd.Hourglass True
14400           If .saleCurr_Date.Visible = True Then .saleCurr_Date.Visible = False
14410           DoEvents
14420         End If
14430       End Select
14440     End If

14450     If msgResponse = vbCancel Then
14460       If .saleAccountNo.Enabled = True Then
14470         blnAccountNoHasFocus = True
14480         .saleAccountNo.SetFocus
14490       Else
14500         .saleTransDate.SetFocus
14510       End If
14520     Else

            ' ** If this Sold came from a Sale/posted 'Paid' or 'Purchase' reference,
            ' ** that journal ID may be in CheckNum.
14530       If IsNull(.CheckNum) = False Then
14540         lngThatJrnlID = .CheckNum
14550         Set dbs = CurrentDb
14560         With dbs
                ' ** Update Journal, by specified [jrnlid], [pstd].
14570           Set qdf = .QueryDefs("qryJournal_Misc_11")
14580           With qdf.Parameters
14590             ![jrnlid] = lngThatJrnlID
14600             ![pstd] = False
14610           End With
14620           qdf.Execute
14630           Set qdf = Nothing
14640           .Close
14650         End With
14660         Set dbs = Nothing
14670       End If

14680       .Undo
14690       .Undo

            ' ** If other sales exist, their cost may have calculated this sale into
            ' ** their tax lot cost, making that other cost inaccurate once this is deleted.
            ' ** TRUE, BUT the multiple-sale check does not apply to dollar-for-dollar assets, which
            ' ** SHOULD remain at $1.00 per share regardless of how many others may or may not be sold.

14700       DoCmd.SetWarnings False
14710 On Error Resume Next
14720       DoCmd.RunCommand acCmdSelectRecord
14730       DoCmd.RunCommand acCmdDeleteRecord
14740 On Error GoTo ERRH
14750       DoCmd.SetWarnings True

14760       blnGoToSaleReinvest = False
14770       .tglSaleReinvest_false_raised_img_dis.Visible = True
14780       .tglSaleReinvest_false_raised_img.Visible = False
14790       .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = False
14800       .tglSaleReinvest_false_raised_focus_img.Visible = False
14810       .tglSaleReinvest_false_raised_focus_dots_img.Visible = False
14820       .tglSaleReinvest_false_sunken_focus_dots_img.Visible = False
14830       .tglSaleReinvest_true_raised_img.Visible = False
14840       .tglSaleReinvest_true_raised_focus_img.Visible = False
14850       .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
14860       .tglSaleReinvest_true_sunken_focus_dots_img.Visible = False
14870       .tglSaleReinvest_true_raised_img_dis.Visible = False
14880       .tglSaleReinvest.Enabled = False

14890       .cmbAccountHelper = Null

            ' ** Also delete reinvestment.
14900       If msgResponse = vbYes And lngThatJrnlID > 0& Then
14910         Set dbs = CurrentDb
              ' ** Delete Journal, by specified [jrnlid].
14920         Set qdf = dbs.QueryDefs("qryJournal_Sale_18")
14930         With qdf.Parameters
14940           ![jrnlid] = lngThatJrnlID
14950         End With
14960         qdf.Execute
14970         Set qdf = Nothing
14980         dbs.Close
14990         Set dbs = Nothing
15000       End If

15010       Forms("frmJournal").DefSaleBtn = False
15020       Forms("frmJournal").DefAssetNo = 0&
15030       Forms("frmJournal").DefPurchTot = 0@
15040       Forms("frmJournal").DefPurchJrnlID = Null
15050       Forms("frmJournal").DefPurchJrnlMSID = Null
15060       Forms("frmJournal").DefPaidTot = 0@
15070       Forms("frmJournal").DefPaidJrnlID = Null
15080       Forms("frmJournal").DefPaidJrnlMSID = Null

15090       blnSoldCosted = False
15100       SoldCostedSet False  ' ** Procedure: Below.

15110       .AllowAdditions = True

15120       .Parent.frmJournal_Sub3_Purchase.Form.Requery
15130       DoEvents

15140       If .saleAccountNo.Enabled = True Then
15150         blnAccountNoHasFocus = True
15160         .saleAccountNo.SetFocus
15170       Else
15180         .saleTransDate.SetFocus
15190       End If

15200       gblnSaleValidated = False

15210       SaleChanged False  ' ** procedure: Below.

15220       Form_Current  ' ** Procedure: Above.

15230       If blnGoneToReport = True Then
15240         If lngGTR_ID > 0& Then
15250           If IsNull(.saleID) = False Then
15260             If .saleID = lngGTR_ID Then
                    ' ** Why didn't it delete it?
15270               .TimerInterval = 250&
15280             Else
15290               blnGoneToReport = False
15300             End If
15310           Else
15320             blnGoneToReport = False
15330           End If
15340         Else
15350           blnGoneToReport = False
15360         End If
15370       End If

15380     End If

15390     DoCmd.Hourglass False

15400   End With

EXITP:
15410   Set qdf = Nothing
15420   Set dbs = Nothing
15430   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleAccountNo_AfterUpdate()

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAccountNo_AfterUpdate"

        Dim varAccountNo As Variant
        Dim blnNoGood As Boolean

15510   blnNoGood = False

15520   With Me
15530     If gstrSaleAccountNumber <> vbNullString Then
            ' ** This means it entered the record with an accountno.
15540       If IsNull(.UnitPurchaseDate) = False And IsNull(.saleAssetno) = False Then
15550         If .saleAccountNo <> gstrSaleAccountNumber Then
15560           varAccountNo = DLookup("[accountno]", "ActiveAssets", "[assetno] = " & CStr(.saleAssetno) & " " & _
                  "And [assetdate] = #" & Format(.UnitPurchaseDate, "mm/dd/yyyy hh:nn:ss AM/PM") & "#")
15570           If IsNull(varAccountNo) = True Then
15580             blnNoGood = True
15590           Else
15600             If .saleAccountNo <> varAccountNo Then
15610               blnNoGood = True
15620             End If
15630           End If
15640           If blnNoGood = True Then
15650             Beep
15660             MsgBox "The Account No. cannot be changed on a " & .saleType & "." & vbCrLf & vbCrLf & _
                    "If the Account was in error, Cancel the transaction and re-enter it.", vbInformation + vbOKOnly, "Invalid Action"
15670             .saleAccountNo = gstrSaleAccountNumber
15680           End If
15690         End If
15700       Else
              ' ** They haven't chosen a Tax Lot yet, so it's OK.
15710       End If
15720     End If
15730     If gblnGoToReport = True Then
15740       .saleAccountNo.SelLength = 0
15750       .saleAccountNo.SelStart = 16
15760     End If
15770     SaleChanged True  ' ** Procedure: Below.
15780   End With

EXITP:
15790   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub saleAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAccountNo_KeyDown"

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
          Case vbKeyTab
            ' ** Tab key copies AccountNo. {intRetVal not cleared.}
15870       With Me
15880         If gblnTabCopyAccount Then
15890           If .NewRecord = True Then
15900             If IsNull(.saleAccountNo) = True Then
15910               If Len(.saleAccountNo.text) = 0 Then
15920                 If IsNull(.Parent.LastAcctNo) = False Then
15930                   If .Parent.LastAcctNo <> vbNullString Then
15940                     .saleAccountNo = .Parent.LastAcctNo
15950                     SaleChanged True  ' ** Procedure: Below.
15960                   End If
15970                 End If
15980               End If
15990             Else
16000               If Len(.saleAccountNo.text) = 0 Then
16010                 If IsNull(.Parent.LastAcctNo) = False Then
16020                   If .Parent.LastAcctNo <> vbNullString Then
16030                     .saleAccountNo = .Parent.LastAcctNo
16040                     SaleChanged True  ' ** Procedure: Below.
16050                   End If
16060                 End If
16070               End If
16080             End If
16090           End If
16100         End If
16110         gblnSaleChanged = True
16120         gblnSaleValidated = False
16130       End With
16140     Case vbKeyReturn
16150       With Me
16160         intRetVal = 0
16170         gblnSaleChanged = True  ' ** Just added: 09/19/2012.
16180         gblnSaleValidated = False
16190 On Error Resume Next
16200         blnFromAcctNoKeyDown = True
16210         .saleType.SetFocus
16220 On Error GoTo ERRH
16230       End With
16240     Case vbKeyDelete
16250       gblnSaleValidated = False
16260     End Select
16270   End If

EXITP:
16280   KeyCode = intRetVal
16290   Exit Sub

ERRH:
430     intRetVal = 0
440     THAT_PROC = THIS_PROC
450     That_Erl = Erl
460     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
470     Resume EXITP

End Sub

Private Sub saleAccountNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAccountNo_MouseMove"

16310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16320   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Private Sub saleAccountNo_Exit(Cancel As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAccountNo_Exit"

        Dim strAccountNo As String
        Dim blnGoToAccountHelper As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

16410   With Me

16420     blnGoToAccountHelper = False  ' ** Unless set otherwise.

16430     If ((gblnSaleValidated = False And gblnSaleChanged = True) And (lngErrCnt < 3&)) Then
16440       If IsNull(.saleAccountNo) = True Then
16450         lngErrCnt = lngErrCnt + 1&
16460         msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Entry Required")
16470         If msgResponse = vbYes Then
16480           blnGoToAccountHelper = True
16490         Else
16500           gblnSaleValidated = False
16510           If blnStartTrans = True Then
16520             Cancel = -1
16530             .saleAccountNo.SetFocus
16540           End If
16550         End If
16560       Else
16570         If .saleAccountNo = vbNullString And .saleAccountNo.text = vbNullString Then
16580           lngErrCnt = lngErrCnt + 1&
16590           msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                  "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Entry Required")
16600           If msgResponse = vbYes Then
16610             blnGoToAccountHelper = True
16620           Else
16630             gblnSaleValidated = False
16640             If blnStartTrans = True Then
16650               Cancel = -1
16660               .saleAccountNo.SetFocus
16670             End If
16680           End If
16690         Else
                ' ** Check to see if it is a valid account.
16700           lngErrCnt = 0&
16710           Select Case gblnAccountNoWithType
                Case True
16720             strAccountNo = Left(.saleAccountNo, 2) & Mid(.saleAccountNo, 3)
16730           Case False
16740             strAccountNo = .saleAccountNo
16750           End Select
16760           varTmp00 = DLookup("[accountno]", "account", "[accountno] = '" & strAccountNo & "' And ([dateclosed] Is Null)")
16770           Select Case IsNull(varTmp00)
                Case True
16780             blnStartTrans = True   ' ** Once they've started, they've got to enter a valid number!
16790             msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                    "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Invalid Entry")
16800             If msgResponse = vbYes Then
16810               blnGoToAccountHelper = True
16820             Else
16830               gblnSaleValidated = False
16840               Cancel = -1
16850               .saleAccountNo.SetFocus
16860             End If
16870             .saleAccountNo = vbNullString
16880           Case False
16890             gblnSaleValidated = True
16900           End Select
16910           If IsNull(.saleAccountNo) = False Then
16920             If .saleAccountNo <> vbNullString Then
16930               .saleAccountNo_Data = .saleAccountNo
16940               gstrSaleAccountNumber = .saleAccountNo.text
16950               .Parent.LastAcctNo = gstrSaleAccountNumber
16960 On Error Resume Next
16970               .Refresh
16980               If ERR.Number <> 0 Then  ' ** 3314: The field 'journal.journaltype' cannot contain a Null value because the Required property for this field is set to True.
16990                 .saleType = "Sold"
17000                 blnAccountNoErr = True
17010               End If
17020 On Error GoTo ERRH
17030             End If
17040           End If
17050         End If
17060       End If
17070     Else
17080       lngErrCnt = 0&
17090     End If

17100     blnAccountNoHasFocus = False
17110     If blnGoToAccountHelper = True Then
17120       .cmbAccountHelper.SetFocus
17130     End If

17140   End With

EXITP:
17150   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleAccountNo_LostFocus()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAccountNo_LostFocus"

17210   With Me
17220     .saleAssetno.Requery
17230     DoEvents
17240     If IsNull(.saleAccountNo) = False Then
17250       If Trim(.saleAccountNo) <> vbNullString Then
17260         If .saleAssetno.ListCount = 0 Then
17270           MsgBox "The chosen account has no assets to sell.", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
17280         End If
17290       End If
17300     End If
17310   End With

EXITP:
17320   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub cmdLock_Click()

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLock_Click"

17410   With Me
17420     If .cmdLock_open_raised_img.Visible = True Then
17430 On Error Resume Next
17440       .saleTransDate.SetFocus
17450       Do Until ERR.Number = 0
17460 On Error GoTo ERRH
17470         .saleTransDate.Enabled = True
17480 On Error Resume Next
17490         .saleTransDate.SetFocus
17500       Loop
17510 On Error GoTo ERRH
17520       .saleAccountNo.Enabled = False
17530       .saleAccountNo.BorderColor = WIN_CLR_DISR
17540       .saleAccountNo.BackStyle = acBackStyleTransparent
17550       .saleAccountNo_lbl.BackStyle = acBackStyleTransparent
17560       .saleAccountNo_lbl_box.Visible = True
17570       .cmdLock_open_raised_img.Visible = False
17580       .cmdLock_closed_raised_img.Visible = True
17590     ElseIf .cmdLock_closed_raised_img.Visible = True Then
17600       .saleAccountNo.Enabled = True
17610       .saleAccountNo.BorderColor = CLR_LTBLU2
17620       .saleAccountNo.BackStyle = acBackStyleNormal
17630       .saleAccountNo_lbl.BackStyle = acBackStyleNormal
17640       .saleAccountNo_lbl_box.Visible = False
17650 On Error Resume Next
17660       .saleAccountNo.SetFocus
17670       Do Until ERR.Number = 0
17680 On Error GoTo ERRH
17690         .saleAccountNo.Enabled = True
17700 On Error Resume Next
17710         .saleAccountNo.SetFocus
17720       Loop
17730 On Error GoTo ERRH
17740       .cmdLock_open_raised_img.Visible = True
17750       .cmdLock_closed_raised_img.Visible = False
17760     End If
17770   End With

EXITP:
17780   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub cmbAccountHelper_Enter()

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_Enter"

17810   With Me
17820     Select Case IsNull(.cmbAccountHelper)
          Case True
17830       ExpandCombo .cmbAccountHelper  ' ** Module Procedure: modUtilities.
17840     Case False
17850       If Trim(.cmbAccountHelper) = vbNullString Then
17860         ExpandCombo .cmbAccountHelper  ' ** Module Procedure: modUtilities.
17870       End If
17880     End Select
17890   End With

EXITP:
17900   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Private Sub cmbAccountHelper_GotFocus()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_GotFocus"

18010   With Me
18020     .cmbAccountHelper.ForeColor = 0
18030   End With

EXITP:
18040   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub cmbAccountHelper_AfterUpdate()

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_AfterUpdate"

18110   With Me
18120     If IsNull(.cmbAccountHelper) = False Then
18130       .saleAccountNo = .cmbAccountHelper
18140       .saleAccountNo.SetFocus
18150     End If
18160   End With

EXITP:
18170   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
440     Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Sub

Private Sub cmbAccountHelper_LostFocus()

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_LostFocus"

18210   With Me

18220     gblnSaleValidated = False
18230     SaleChanged True  ' ** Procedure: Below.
18240     .cmbAccountHelper.ForeColor = CLR_WHT
18250     .saleAccountNo.SetFocus
18260     DoEvents
18270     .saleAccountNo.Value = .cmbAccountHelper.Value
18280     DoEvents

18290     gstrFormQuerySpec = .Parent.Name  ' ** Make sure this is set for the assetno combo box.
18300     .saleAssetno.Requery
18310     DoEvents
18320     If IsNull(.saleAccountNo) = False Then
18330       If Trim(.saleAccountNo) <> vbNullString Then
18340         If .saleAssetno.ListCount = 0 Then
18350           MsgBox "The chosen account has no assets to sell.", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
18360         End If
18370       End If
18380     End If

18390   End With

EXITP:
18400   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub saleShortName_LostFocus()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "saleShortName_LostFocus"

18510   VerifyAccountNo  ' ** Function: Below.

EXITP:
18520   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub saleType_Enter()

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "saleType_Enter"

18610   With Me
18620     If IsNull(.saleAccountNo) = False Then
18630       If Trim(.saleAccountNo) <> vbNullString Then

18640         If IsNull(.saleType) = True Or blnAccountNoErr = True Then
18650           blnAccountNoErr = False
18660           ExpandCombo .saleType  ' ** Module Procedure: modUtilities.
18670           .saleType = .saleType.ItemData(0)
18680           blnDefTypeAssigned = True  ' ** Currious that saleType_AfterUpdate() stopped being triggered!?
18690         End If

18700 On Error Resume Next
18710         DoCmd.RunCommand acCmdSaveRecord
18720 On Error GoTo ERRH

18730         If .saleType <> .saleType.OldValue Or IsNull(.saleType.OldValue) Then
18740           SaleChanged True  ' ** Procedure: Below.
18750         End If

18760         VerifyAccountNo  ' ** Function: Below.

18770       End If
18780     End If
18790   End With

EXITP:
18800   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleType_AfterUpdate()

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "saleType_AfterUpdate"

18910   JournalType_After_Sub4 blnFromElsewhere, blnDefTypeAssigned, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
18920   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleType_Exit(Cancel As Integer)

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "saleType_Exit"

19010   With Me
19020     If IsNull(.saleAccountNo) = False Then
19030       If Trim(.saleAccountNo) <> vbNullString Then
19040         VerifyAccountNo  ' ** Function: Below.
19050       End If
19060     End If
19070   End With

        ' ** Currious that saleType_AfterUpdate() stopped being triggered!?
19080   If blnDefTypeAssigned = True Then
19090     saleType_AfterUpdate  ' ** Procedure: Above.
19100   End If

EXITP:
19110   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub saleType_LostFocus()

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "saleType_LostFocus"

19210   With Me
19220     If IsNull(.saleAccountNo) = False Then
19230       If Trim(.saleAccountNo) <> vbNullString Then

19240         gstrSaleType = .saleType.text

19250         If .saleType <> .saleType.OldValue Or IsNull(.saleType.OldValue) Then
19260           SaleChanged True  ' ** Procedure: Below.
19270         End If

19280         If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
19290           .cmbTaxCodes.Enabled = True
19300           .cmbTaxCodes.BorderColor = CLR_LTBLU2
19310           .cmbTaxCodes_lbl.BackStyle = acBackStyleNormal
19320           .cmbTaxCodes_lbl_box.Visible = False
19330         ElseIf .saleType = "Withdrawn" Then
19340           .cmbTaxCodes.Enabled = True
19350           .cmbTaxCodes.BorderColor = CLR_LTBLU2
19360           .cmbTaxCodes_lbl.BackStyle = acBackStyleNormal
19370           .cmbTaxCodes_lbl_box.Visible = False
19380         End If

19390         subTaxCode4 blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modPurchaseSold.
19400         subRevCode4 blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modPurchaseSold.

19410         VerifyAccountNo  ' ** Function: Below.

19420       End If
19430     End If
19440   End With

EXITP:
19450   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleTransDate_BeforeUpdate(Cancel As Integer)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "saleTransDate_BeforeUpdate"

19510   With Me
19520     If .saleTransDate.text <> "__/__/____" And .saleTransDate.text <> vbNullString Then
19530       If Left(.saleTransDate.text, 2) > 12 Then
19540         MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Entry"
19550         Cancel = -1
19560       End If
19570     End If
19580   End With

EXITP:
19590   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleTransDate_KeyDown(KeyCode As Integer, Shift As Integer)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "saleTransDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19620   intShiftDown = (Shift And acShiftMask) > 0
19630   intAltDown = (Shift And acAltMask) > 0
19640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19670       With Me
19680         intRetVal = 0
19690         .saleAssetno.SetFocus
19700       End With
19710     End Select
19720   End If

        ' ** Shift keys.
19730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19750       With Me
19760         intRetVal = 0
19770         .saleType.SetFocus
19780       End With
19790     End Select
19800   End If

        ' ** Ctrl keys.
19810   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
19820     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19830       With Me
19840         intRetVal = 0
19850         .cmdCalendar1.SetFocus
19860       End With
19870     End Select
19880   End If

EXITP:
19890   KeyCode = intRetVal
19900   Exit Sub

ERRH:
430     intRetVal = 0
440     THAT_PROC = THIS_PROC
450     That_Erl = Erl
460     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
470     Resume EXITP

End Sub

Private Sub saleTransDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "saleTransDate_MouseMove"

20010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
20020   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Private Sub saleTransDate_Exit(Cancel As Integer)

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "saleTransDate_Exit"

20110   With Me
20120     If lngDateCnt <= 1& Then
20130       If IsNull(.saleAccountNo) = False Then
20140         If Trim(.saleAccountNo) <> vbNullString Then
20150           If DateCheck_Post(.saleTransDate.text) = True Then  ' ** Module Function: modUtilities.
                  ' ** '.Text' needed to reliably check date.
20160             lngDateCnt = 0&
20170             If IsNull(.saleTransDate.OldValue) Then
20180               SaleChanged True  ' ** Procedure: Below.
20190             Else
20200               If .saleTransDate <> .saleTransDate.OldValue Then
20210                 SaleChanged True  ' ** Procedure: Below.
20220               End If
20230             End If
20240           Else
20250             If IsNull(.saleTransDate) = True Then
20260               .saleTransDate = Date
20270             Else
20280               If Trim(.saleTransDate) = vbNullString Then
20290                 .saleTransDate = Date
20300               Else
20310                 lngDateCnt = lngDateCnt + 1&
20320                 Cancel = -1
20330                 .saleTransDate.SetFocus
20340               End If
20350             End If
20360           End If
20370         End If
20380       End If
20390     Else
20400       lngDateCnt = 0&
20410       .saleTransDate = CDate(Format(datPostingDate, "mm/dd/yyyy"))
20420     End If
20430   End With

EXITP:
20440   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub saleTransDate_LostFocus()

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "saleTransDate_LostFocus"

20510   With Me
20520     If IsNull(.saleAccountNo) = False Then
20530       If Trim(.saleAccountNo) <> vbNullString Then
20540         VerifyAccountNo  ' ** Function: Below.
20550       End If
20560     End If
20570   End With

EXITP:
20580   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

20610   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
20620   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

20710   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
20720   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Public Sub cmdCalendar1_Click()

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

20810   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
20820   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
440     Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

20910   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
20920   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

21010   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
21020   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

21110   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
21120   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Private Sub saleAssetno_Enter()

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetno_Enter"

        Dim blnLotInfo As Boolean

        Static intLoopCnt As Integer

21210   With Me

          ' ** Make sure the Combo Box is populated.
21220     If intLoopCnt = 0 Then
21230       intLoopCnt = 1
            ' ** Don't lose this if it's set!
21240       blnLotInfo = blnBeenToLotInfo
21250       blnFromSaleAssetnoEnter = True
21260       blnCmdButton = True
21270       Form_Current  ' ** Procedure: Above.
21280       blnCmdButton = False
21290       blnFromSaleAssetnoEnter = False
21300       blnBeenToLotInfo = blnLotInfo
21310     Else
21320       intLoopCnt = 0
21330     End If

21340     .saleAssetno.SetFocus  ' ** Because the OnCurrent() event may move it.
21350     .saleAssetno.Requery
21360     DoEvents
21370     If IsNull(.saleAssetno) = True Or .saleAssetno.text = vbNullString Then
21380       ExpandCombo .saleAssetno  ' ** Module Procedure: modUtilities.
21390     End If

21400   End With

EXITP:
21410   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleAssetno_BeforeUpdate(Cancel As Integer)

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetNo_BeforeUpdate"

        Dim blnCurrID As Boolean
        Dim varTmp00 As Variant

21510   With Me
21520     If IsNull(.saleAssetno) = False Then
21530       blnCurrID = .saleCurr_ID.Visible
21540       If blnCurrID = False Then
21550         varTmp00 = .saleAssetno.Column(CBX_A_CURRID)
21560         If IsNull(varTmp00) = False Then
21570           If CLng(varTmp00) <> 150& Then
21580             Beep
21590             MsgBox "The chosen asset is in a foreign currency." & vbCrLf & vbCrLf & _
                    "You must check the 'Include Currency' checkbox on the Posting Date" & vbCrLf & _
                    "window before entering transactions not in U.S. dollars.", vbInformation + vbOKOnly, "Foreign Currency Asset"
21600             .saleAssetno.Undo
21610             Cancel = -1
21620           End If
21630         End If
21640       End If
21650     End If
21660   End With

EXITP:
21670   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Public Sub saleAssetno_AfterUpdate()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetno_AfterUpdate"

21710   With Me

          ' ** Do this whether or not Currency is visible.
21720     Select Case IsNull(.saleAssetno)
          Case True
21730       lngCurrID = 150&  ' ** USD.
21740       .cmdSaleLotInfo.Enabled = False
21750     Case False
21760       .cmdSaleLotInfo.Enabled = True
21770       Select Case IsNull(.saleAssetno.Column(CBX_A_CURRID))
            Case True
21780         lngCurrID = 150&  ' ** USD.
21790       Case False
21800         lngCurrID = .saleAssetno.Column(CBX_A_CURRID)
21810       End Select
21820     End Select
21830     .saleCurr_ID = lngCurrID
21840     saleCurr_ID_AfterUpdate  ' ** Procedure: Below.

21850     subTaxCode4 blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modPurchaseSold.
21860     ChkAverage_PS_Sold Me  ' ** Module Procedure: modPurchaseSold.

21870   End With

EXITP:
21880   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleAssetno_KeyDown(KeyCode As Integer, Shift As Integer)

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21920   intShiftDown = (Shift And acShiftMask) > 0
21930   intAltDown = (Shift And acAltMask) > 0
21940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21970       With Me
21980         intRetVal = 0
21990         .saleShareFace.SetFocus
22000       End With
22010     End Select
22020   End If

        ' ** Shift keys.
22030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22050       With Me
22060         intRetVal = 0
22070         .saleTransDate.SetFocus
22080       End With
22090     End Select
22100   End If

        ' ** Ctrl-Shift keys.
22110   If intCtrlDown And (Not intAltDown) And intShiftDown Then
22120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22130       With Me
22140         intRetVal = 0
22150         .cmdCalendar1.SetFocus
22160       End With
22170     End Select
22180   End If

EXITP:
22190   KeyCode = intRetVal
22200   Exit Sub

ERRH:
430     intRetVal = 0
440     THAT_PROC = THIS_PROC
450     That_Erl = Erl
460     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
470     Resume EXITP

End Sub

Public Sub saleAssetno_Change()

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetno_Change"

22310   With Me
22320     If .saleAssetno <> .saleAssetno.OldValue And .Parent.DefSaleBtn = False Then  ' ** Let them pick an asset.
22330       SaleChanged True  ' ** Procedure: Below.
22340       .saleICash = 0
22350       .saleICash_usd = Null
22360       blnFromElsewhere = True
22370       saleICash_AfterUpdate  ' ** Procedure: Below.
22380       .salePCash = 0
22390       .salePCash_usd = Null
22400       blnFromElsewhere = True
22410       salePCash_AfterUpdate  ' ** Procedure: Below.
22420       .saleCost = 0
22430       .saleCost_usd = Null
22440       blnFromElsewhere = True
22450       saleCost_AfterUpdate  ' ** Procedure: Below.
22460     End If
22470   End With

EXITP:
22480   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleAssetno_NotInList(NewData As String, Response As Integer)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetno_NotInList"

22510   With Me
22520     .saleAssetno.SelStart = 0
22530     .saleAssetno.SelLength = Len(.saleAssetno.text)
22540   End With

EXITP:
22550   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub saleAssetNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetNo_MouseMove"

22610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
22620   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Public Sub saleAssetno_LostFocus()

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetno_LostFocus"

22710   With Me
22720     If IsNull(.saleAssetno) = False Then
22730       Select Case IsNull(.saleCost)
            Case True
22740         .cmdSaleLotInfo.Enabled = True
22750       Case False
22760         If .saleCost = 0 Then
22770           .cmdSaleLotInfo.Enabled = True
22780         Else
22790           .cmdSaleLotInfo.Enabled = False
22800         End If
22810       End Select
22820     End If
22830     gstrSaleAsset = .saleAssetno.text
22840   End With

EXITP:
22850   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Public Sub saleAssetno_Exit(Cancel As Integer)

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetno_Exit"

        Dim varTmp00 As Variant

22910   With Me
22920     VerifyAccountNo  ' ** Function: Below.
          ' ** Check curr_id.
22930     varTmp00 = .saleAssetno.Column(CBX_A_CURRID)
22940     If IsNull(varTmp00) = False Then
22950       If .saleCurr_ID <> varTmp00 Then
22960         .saleCurr_ID = varTmp00
22970         saleCurr_ID_AfterUpdate  ' ** Procedure: Below.
22980       End If
22990     End If
23000   End With

EXITP:
23010   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub cmdPaidTotal_Click()
' ** For now, I'm using the 'Posted' field in Journal (which is a Yes/No, not a date field)
' ** to signal if the user invoked the Sale button for an individual 'Paid' transaction.
' **

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPaidTotal_Click"

23110   PaidTotal_Sub4 blnFromElsewhere, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
23120   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub saleShareFace_Enter()

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "saleShareFace_Enter"

23210   With Me
23220     .saleShareFace.Format = "General Number"
23230   End With

EXITP:
23240   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Public Sub saleShareFace_AfterUpdate()

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "saleShareFace_AfterUpdate"

23310   With Me
23320     If IsNull(.saleCost) = False Then
23330       If .saleCost <> 0@ Then
23340         .saleCost = 0@
23350         .saleCost_usd = Null
23360         blnFromElsewhere = True
23370         saleCost_AfterUpdate  ' ** Procedure: Below.
23380       End If
23390     End If
23400   End With

EXITP:
23410   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Public Sub saleShareFace_Exit(Cancel As Integer)

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "saleShareFace_Exit"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

23510   blnContinue = True

23520   VerifyAccountNo  ' ** Function: Below.

23530   With Me
23540     If gblnSaleValidated Then

23550       .saleShareFace.SetFocus
23560       If IsNumeric(.saleShareFace.text) Then
23570         If CDbl(.saleShareFace.text) > 999999999.9999 Then
23580           blnContinue = False
23590           MsgBox "Share/Face is too large.", vbInformation + vbOKOnly, "Invalid Entry"
23600           .saleShareFace = vbNullString
23610           .saleShareFace = 0
23620           Cancel = -1
23630           .saleShareFace.SetFocus
23640         Else
23650           If .saleShareFace < 0 Then
23660             blnContinue = False
23670             MsgBox "You cannot enter a negative Share/Face.", vbInformation + vbOKOnly, "Invalid Entry"
23680             .saleShareFace = CStr(0)
23690             .Refresh
23700             Cancel = -1
23710             .saleShareFace.SetFocus
23720           End If
23730         End If
23740       Else
23750         blnContinue = False
23760         MsgBox "Share/Face must be numreic.", vbInformation + vbOKOnly, "Invalid Entry"
23770         .saleShareFace = vbNullString
23780         DoCmd.RunCommand acCmdSaveRecord
23790         .saleShareFace = CStr(0)
23800         Cancel = -1
23810         .saleShareFace.SetFocus
23820       End If

23830       If blnContinue = True Then
23840         intPos01 = InStr(1, .saleShareFace, ".", vbTextCompare)
23850         If intPos01 > 0 Then
23860           If Len(.saleShareFace) - intPos01 > gintShareFaceDecimals Then
23870             blnContinue = False
23880             MsgBox "You can only enter a Share/Face with up to " & Trim(str(gintShareFaceDecimals)) & " decimals.", _
                    vbInformation + vbOKOnly, "Invalid Entry"
23890             .saleShareFace = vbNullString
23900             Cancel = -1
23910             .saleShareFace.SetFocus
23920           Else
23930             .saleShareFace.Format = "#,###." & String(gintShareFaceDecimals, "0")
23940           End If
23950         Else
23960           .saleShareFace.Format = "#,###"
23970         End If
23980       End If

23990       If blnContinue = True Then
24000         gstrSaleShareFace = .saleShareFace.text
24010         If .saleShareFace <> .saleShareFace.OldValue Or IsNull(.saleShareFace.OldValue) Then
24020           SaleChanged True  ' ** Procedure: Below.
24030         End If
24040       Else
24050         If (gblnSaleChanged = False) Or (.NavigationButtons = True) Then
                ' ** Do nothing.
24060         Else
24070           .saleShareFace = vbNullString
24080         End If
24090       End If

24100     End If
24110   End With

EXITP:
24120   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Public Sub saleShareFace_LostFocus()

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "saleShareFace_LostFocus"

24210   VerifyAccountNo  ' ** Function: Below.

EXITP:
24220   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub saleAssetDate_Enter()

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetDate_Enter"

24310   With Me
24320     .saleAssetDate.text = Format(.saleAssetDate, "mm/dd/yyyy")
24330   End With

EXITP:
24340   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleAssetDate_BeforeUpdate(Cancel As Integer)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetDate_BeforeUpdate"

24410   With Me
24420     If .saleAssetDate.text <> "__/__/____" And .saleAssetDate.text <> vbNullString Then
24430       If Left(.saleAssetDate.text, 2) > 12 Then
24440         MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Entry"
24450         Cancel = -1
24460       End If
24470     End If
24480   End With

EXITP:
24490   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleAssetDate_KeyDown(KeyCode As Integer, Shift As Integer)

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24520   intShiftDown = (Shift And acShiftMask) > 0
24530   intAltDown = (Shift And acAltMask) > 0
24540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24570       With Me
24580         intRetVal = 0
24590         If .saleICash.Enabled = True Then
24600           .saleICash.SetFocus
24610         ElseIf .salePCash.Enabled = True Then
24620           .salePCash.SetFocus
24630         Else
24640           .saleCost.SetFocus
24650         End If
24660       End With
24670     End Select
24680   End If

        ' ** Shift keys.
24690   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24700     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24710       With Me
24720         intRetVal = 0
24730         .saleShareFace.SetFocus
24740       End With
24750     End Select
24760   End If

        ' ** Ctrl keys.
24770   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
24780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24790       With Me
24800         intRetVal = 0
24810         .cmdCalendar2.SetFocus
24820       End With
24830     End Select
24840   End If

EXITP:
24850   KeyCode = intRetVal
24860   Exit Sub

ERRH:
430     intRetVal = 0
440     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
450     Case Else
460       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
470     End Select
480     Resume EXITP

End Sub

Private Sub saleAssetDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetDate_MouseMove"

24910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
24920   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Private Sub saleAssetDate_Exit(Cancel As Integer)

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetDate_Exit"

25010   With Me
25020     If DateCheck_Trade(.saleAssetDate.text) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
25030       If IsNull(.saleAssetDate.OldValue) Then
25040         SaleChanged True  ' ** Procedure: Below.
25050       Else
25060         If .saleAssetDate <> .saleAssetDate.OldValue Then
25070           SaleChanged True  ' ** Procedure: Below.
25080         End If
25090       End If
25100     Else
25110       If IsNull(.saleAssetDate) = True Then
25120         .saleAssetDate = Now()
25130       Else
25140         If Trim(.saleAssetDate) = vbNullString Then
25150           .saleAssetDate = Now()
25160         Else
25170           Cancel = -1
25180           .saleAssetDate.SetFocus
25190         End If
25200       End If
25210     End If
25220   End With

EXITP:
25230   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleAssetDate_LostFocus()

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "saleAssetDate_LostFocus"

25310   VerifyAccountNo  ' ** Function: Below.

EXITP:
25320   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

25410   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
25420   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

25510   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
25520   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Public Sub cmdCalendar2_Click()

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

25610   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
25620   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
440     Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

25710   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
25720   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

25810   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
25820   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

25910   Calendar_Handler_Sub4 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, blnCalendar2_Focus, blnCalendar2_MouseDown, _
          clsMonthClass, Me  ' ** Module Procedure: modJrnlSub4SoldFuncs.

EXITP:
25920   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
450     End Select
460     Resume EXITP

End Sub

Public Sub saleICash_AfterUpdate()

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "saleICash_AfterUpdate"

        Dim dblICash As Double
        Dim varTmp00 As Variant

26010   With Me

26020     Select Case blnFromElsewhere
          Case True
26030       blnFromElsewhere = False
26040     Case False
26050       blnWarned = False
26060     End Select

26070     If .saleCurr_ID.Visible = True Then
26080       .saleICash_usd = Null
26090       If .saleICash_usd.Visible = False Then .saleICash_usd.Visible = True
26100       Select Case .saleICash.Enabled
            Case True
26110         If IsNull(.saleICash) = False Then
26120           If .saleICash <> 0 Then
26130             If IsNull(.saleCurr_ID.Column(CBX_C_CURRID)) = False Then
26140               lngCurrID = .saleCurr_ID.Column(CBX_C_CURRID)
26150               If lngCurrID <> 150& Then  ' ** USD.
26160                 varTmp00 = .saleCurr_ID.Column(CBX_C_RATE2)
26170                 If IsNull(varTmp00) = False Then
26180                   If varTmp00 > 0 Then
26190                     dblICash = .saleICash
26200                     dblICash = (dblICash * varTmp00)
26210                     dblICash = Round(dblICash, 2)
26220                     .saleICash_usd = dblICash
26230                     If IsNull(.salePCash) = False Then
26240                       If .salePCash = 0 Then
26250                         If IsNull(.salePCash_usd) = True Then
26260                           salePCash_AfterUpdate  ' ** Procedure: Below.
26270                         End If
26280                       End If
26290                     End If
26300                     If IsNull(.saleCost) = False Then
26310                       If .saleCost = 0 Then
26320                         If IsNull(.saleCost_usd) = True Then
26330                           saleCost_AfterUpdate  ' ** Procedure: Below.
26340                         End If
26350                       End If
26360                     End If
26370                   End If
26380                 End If
26390               Else
26400                 .saleICash_usd = .saleICash
26410               End If
26420             End If
26430           Else
26440             .saleICash_usd = 0#
26450           End If
26460         End If
26470       Case False
26480         .saleICash_usd = Null
26490       End Select
26500     End If
26510     DoEvents

26520   End With

EXITP:
26530   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleICash_KeyDown(KeyCode As Integer, Shift As Integer)

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "saleICash_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26620   intShiftDown = (Shift And acShiftMask) > 0
26630   intAltDown = (Shift And acAltMask) > 0
26640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26670       With Me
26680         intRetVal = 0
26690         If .salePCash.Enabled = True Then
26700           .salePCash.SetFocus
26710         Else
26720           .saleCost.SetFocus
26730         End If
26740       End With
26750     End Select
26760   End If

        ' ** Shift keys.
26770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26790       With Me
26800         intRetVal = 0
26810         .saleAssetDate.SetFocus
26820       End With
26830     End Select
26840   End If

        ' ** Ctrl-Shift keys.
26850   If intCtrlDown And (Not intAltDown) And intShiftDown Then
26860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26870       With Me
26880         intRetVal = 0
26890         .cmdCalendar2.SetFocus
26900       End With
26910     End Select
26920   End If

EXITP:
26930   KeyCode = intRetVal
26940   Exit Sub

ERRH:
430     intRetVal = 0
440     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
450     Case Else
460       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
470     End Select
480     Resume EXITP

End Sub

Private Sub saleICash_Exit(Cancel As Integer)

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "saleICash_Exit"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

27010   With Me

27020     blnContinue = True

27030     lngCurrID = .saleCurr_ID.Column(CBX_C_CURRID)

27040     If IsNumeric(.saleICash) Then
27050       If CDbl(Rem_Dollar(.saleICash.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
27060         blnContinue = False
27070         MsgBox "Accrued/Income Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
27080         .saleICash = vbNullString
27090         .saleICash_usd = Null
27100         blnFromElsewhere = True
27110         saleICash_AfterUpdate  ' ** Procedure: Above.
27120         Cancel = -1
27130         .saleICash.SetFocus
27140       End If
27150     Else
27160       blnContinue = False
27170       MsgBox "Accrued/Income Cash must have a value.", vbInformation + vbOKOnly, "Entry Required"
27180       .saleICash = vbNullString
27190       .saleICash_usd = Null
27200       blnFromElsewhere = True
27210       saleICash_AfterUpdate  ' ** Procedure: Above.
27220       Cancel = -1
27230       .saleICash.SetFocus
27240     End If

27250     If blnContinue = True Then
27260       gstrSaleICash = .saleICash.text
27270       gstrSaleICash = Rem_Dollar(gstrSaleICash, lngCurrID)  ' ** Module Function: modStringFuncs.
27280       intPos01 = InStr(1, .saleICash, ".", vbTextCompare)
27290       If intPos01 > 0 Then
27300         If Len(.saleICash) - intPos01 > 2 Then
27310           blnContinue = False
27320           MsgBox "You can only enter Accrued/Income Cash with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
27330           .saleICash = vbNullString
27340           .saleICash_usd = Null
27350           blnFromElsewhere = True
27360           saleICash_AfterUpdate  ' ** Procedure: Above.
27370           Cancel = -1
27380           .saleICash.SetFocus
27390         End If
27400       End If
27410     End If

27420     If blnContinue = True Then
27430       subTaxCode4 blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modPurchaseSold.
27440     End If

27450     If blnContinue = True Then
27460       gstrSaleICash = .saleICash.text
27470       gstrSaleICash = Rem_Dollar(gstrSaleICash, lngCurrID)  ' ** Module Function: modStringFuncs.
27480       If IsNull(.saleICash.OldValue) = True Then
27490         SaleChanged True  ' ** Procedure: Below.
27500       Else
27510         If IsNull(.saleICash) = False Then
27520           If .saleICash <> .saleICash.OldValue Then
27530             SaleChanged True  ' ** Procedure: Below.
27540           End If
27550         End If
27560       End If
27570     End If

27580   End With

EXITP:
27590   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub saleICash_LostFocus()

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "saleICash_LostFocus"

27610   If VerifyAccountNo = True Then  ' ** Function: Below.

27620     With Me

27630       If .saleType = "Liability" Then
27640         If IsNull(.saleICash) = False Then
27650           If .saleICash > 0 Then
27660             .saleICash = .saleICash * -1
27670             blnFromElsewhere = True
27680             saleICash_AfterUpdate  ' ** Procedure: Above.
27690           End If
27700         End If
27710       End If

27720       subRevCode4 blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modPurchaseSold.
27730       subTaxCode4 blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modPurchaseSold.

27740     End With

27750   End If

EXITP:
27760   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Public Sub salePCash_AfterUpdate()

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "salePCash_AfterUpdate"

        Dim dblPCash As Double
        Dim varTmp00 As Variant

27810   With Me

27820     Select Case blnFromElsewhere
          Case True
27830       blnFromElsewhere = False
27840     Case False
27850       blnWarned = False
27860     End Select

27870     If .saleCurr_ID.Visible = True Then
27880       .salePCash_usd = Null
27890       If .salePCash_usd.Visible = False Then .salePCash_usd.Visible = True
27900       Select Case .salePCash.Enabled
            Case True
27910         If IsNull(.salePCash) = False Then
27920           If .salePCash <> 0 Then
27930             If IsNull(.saleCurr_ID.Column(CBX_C_CURRID)) = False Then
27940               lngCurrID = .saleCurr_ID.Column(CBX_C_CURRID)
27950               If lngCurrID <> 150& Then  ' ** USD.
27960                 varTmp00 = .saleCurr_ID.Column(CBX_C_RATE2)
27970                 If IsNull(varTmp00) = False Then
27980                   If varTmp00 > 0 Then
27990                     dblPCash = .salePCash
28000                     dblPCash = (dblPCash * varTmp00)
28010                     dblPCash = Round(dblPCash, 2)
28020                     .salePCash_usd = dblPCash
28030                     If IsNull(.saleICash) = False Then
28040                       If .saleICash = 0 Then
28050                         If IsNull(.saleICash_usd) = True Then
28060                           saleICash_AfterUpdate  ' ** Procedure: Below.
28070                         End If
28080                       End If
28090                     End If
28100                     If IsNull(.saleCost) = False Then
28110                       If .saleCost = 0 Then
28120                         If IsNull(.saleCost_usd) = True Then
28130                           saleCost_AfterUpdate  ' ** Procedure: Below.
28140                         End If
28150                       End If
28160                     End If
28170                   End If
28180                 End If
28190               Else
28200                 .salePCash_usd = .salePCash
28210               End If
28220             End If
28230           Else
28240             .salePCash_usd = 0#
28250           End If
28260         End If
28270       Case False
28280         .salePCash_usd = Null
28290       End Select
28300     End If
28310     DoEvents

28320   End With

EXITP:
28330   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub salePCash_KeyDown(KeyCode As Integer, Shift As Integer)

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "salePCash_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28420   intShiftDown = (Shift And acShiftMask) > 0
28430   intAltDown = (Shift And acAltMask) > 0
28440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28470       With Me
28480         intRetVal = 0
28490         .saleCost.SetFocus
28500       End With
28510     End Select
28520   End If

        ' ** Shift keys.
28530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28550       With Me
28560         intRetVal = 0
28570         If .saleICash.Enabled = True Then
28580           .saleICash.SetFocus
28590         Else
28600           .saleAssetDate.SetFocus
28610         End If
28620       End With
28630     End Select
28640   End If

        ' ** Ctrl-Shift keys.
28650   If intCtrlDown And (Not intAltDown) And intShiftDown Then
28660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28670       With Me
28680         intRetVal = 0
28690         If .saleICash.Enabled = False Then
28700           .cmdCalendar2.SetFocus
28710         End If
28720       End With
28730     End Select
28740   End If

EXITP:
28750   KeyCode = intRetVal
28760   Exit Sub

ERRH:
430     intRetVal = 0
440     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
450     Case Else
460       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
470     End Select
480     Resume EXITP

End Sub

Public Sub salePCash_Exit(Cancel As Integer)

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "salePCash_Exit"

        Dim msgResponse As VbMsgBoxResult
        Dim intPos01 As Integer
        Dim blnContinue As Boolean

28810   With Me

28820     blnContinue = True

28830     If blnClickedLotInfo = False Then

28840       lngCurrID = .saleCurr_ID.Column(CBX_C_CURRID)

28850       If IsNumeric(.salePCash) Then
28860         If CDbl(Rem_Dollar(.salePCash.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
28870           blnContinue = False
28880           MsgBox "Principal Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
28890           .salePCash.text = vbNullString
28900           .salePCash_usd = Null
28910           blnFromElsewhere = True
28920           salePCash_AfterUpdate  ' ** Procedure: Above.
28930           Cancel = -1
28940           .salePCash.SetFocus
28950         Else
28960           If .saleType = "Withdrawn" And CDbl(Rem_Dollar(.salePCash.text, lngCurrID)) <> 0 Then  ' ** Module Function: modStringFuncs.
                  ' ** For a 'Withdrawn', never pcash nor icash, and sometimes not even cost.
28970             blnContinue = False
28980             MsgBox "Withdrawn Principal Cash must be zero.", vbInformation + vbOKOnly, "Invalid Entry"
28990             .salePCash = vbNullString
29000             .salePCash_usd = Null
29010             blnFromElsewhere = True
29020             salePCash_AfterUpdate  ' ** Procedure: Above.
29030             Cancel = -1
29040             .salePCash.SetFocus
29050           Else
29060             If .saleType = "Sold" And IsNull(.saleCost) = False Then
                    ' ** For a 'Sold', always pcash or icash, and always cost. NOT NECESSARILY!
29070               If IsNumeric(.saleCost) = True Then
29080                 If .saleCost < 0 And .salePCash < 0 Then
29090                   .salePCash = (.salePCash * -1#)
29100                   blnFromElsewhere = True
29110                   salePCash_AfterUpdate  ' ** Procedure: Above.
29120                 End If
29130               End If
29140             End If
29150           End If
29160         End If
29170       Else
29180         blnContinue = False
29190         MsgBox "Principal Cash must have a value.", vbInformation + vbOKOnly, "Entry Required"
              ' ** Even if it's 0.
29200         .salePCash = vbNullString
29210         .salePCash_usd = Null
29220         blnFromElsewhere = True
29230         salePCash_AfterUpdate  ' ** Procedure: Above.
29240         Cancel = -1
29250         .salePCash.SetFocus
29260       End If

29270       If blnContinue = True Then
29280         intPos01 = InStr(1, .salePCash, ".", vbTextCompare)
29290         If intPos01 > 0 Then
29300           If Len(.salePCash) - intPos01 > 2 Then
29310             blnContinue = False
29320             .salePCash = Format(.salePCash, "Currency")
29330             blnFromElsewhere = True
29340             salePCash_AfterUpdate  ' ** Procedure: Above.
29350           End If
29360         End If
29370       End If

29380       If blnContinue = True Then
29390         If .saleType = "Sold" And CDbl(Rem_Dollar(.salePCash.text, lngCurrID)) = 0 Then  ' ** Module Function: modStringFuncs.
29400           If .saleICash = 0 Then
29410             msgResponse = MsgBox("Are you sure you want Income and Principal Cash to be ZERO?" & vbCrLf & vbCrLf & _
                    "As would be the case for the sale of worthless shares.", vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
29420           Else
29430             msgResponse = MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
29440           End If
29450           blnWarned = True
29460           If msgResponse <> vbYes Then
29470             blnContinue = False
29480             Cancel = -1
29490             .salePCash.SetFocus
29500           End If
29510         End If
29520       End If

29530       If blnContinue = True Then
29540         subTaxCode4 blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modPurchaseSold.
29550       End If

29560       If blnContinue = True Then
29570         gstrSalePCash = .salePCash.text
29580         gstrSalePCash = Rem_Dollar(gstrSalePCash, lngCurrID)  ' ** Module Function: modStringFuncs.
29590         If .salePCash <> .salePCash.OldValue Or IsNull(.salePCash.OldValue) Then
29600           SaleChanged True  ' ** Procedure: Below.
29610         End If
29620       End If

29630     End If  ' ** blnClickedLotInfo.

29640   End With

EXITP:
29650   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Public Sub salePCash_LostFocus()

29700 On Error GoTo ERRH

        Const THIS_PROC As String = "salePCash_LostFocus"

29710   VerifyAccountNo  ' ** Function: Below.

29720   With Me
29730     If .saleType = "Liability" Then
29740       If .salePCash > 0 Then
29750         .salePCash = .salePCash * -1
29760         blnFromElsewhere = True
29770         salePCash_AfterUpdate  ' ** Procedure: Above.
29780       End If
29790     End If
29800   End With

29810   subRevCode4 blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modPurchaseSold.

EXITP:
29820   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Public Sub saleCost_AfterUpdate()

29900 On Error GoTo ERRH

        Const THIS_PROC As String = "saleCost_AfterUpdate"

        Dim dblCost As Double
        Dim varTmp00 As Variant

29910   With Me

29920     Select Case blnFromElsewhere
          Case True
29930       blnFromElsewhere = False
29940     Case False
            ' ** Nothing.
29950     End Select

29960     If .saleCurr_ID.Visible = True Then
29970       .saleCost_usd = Null
29980       If .saleCost_usd.Visible = False Then .saleCost_usd.Visible = True
29990       If IsNull(.saleCost) = False Then
30000         If .saleCost <> 0 Then
30010           If IsNull(.saleCurr_ID.Column(CBX_C_CURRID)) = False Then
30020             lngCurrID = .saleCurr_ID.Column(CBX_C_CURRID)
30030             If lngCurrID <> 150& Then  ' ** USD.
30040               varTmp00 = .saleCurr_ID.Column(CBX_C_RATE2)
30050               If IsNull(varTmp00) = False Then
30060                 If varTmp00 > 0 Then
30070                   dblCost = .saleCost
30080                   dblCost = (dblCost * varTmp00)
30090                   dblCost = Round(dblCost, 2)
30100                   .saleCost_usd = dblCost
30110                   If IsNull(.saleICash) = False Then
30120                     If .saleICash = 0 Then
30130                       If IsNull(.saleICash_usd) = True Then
30140                         saleICash_AfterUpdate  ' ** Procedure: Above.
30150                       End If
30160                     End If
30170                   End If
30180                   If IsNull(.salePCash) = False Then
30190                     If .salePCash = 0 Then
30200                       If IsNull(.salePCash_usd) = True Then
30210                         salePCash_AfterUpdate  ' ** Procedure: Above.
30220                       End If
30230                     End If
30240                   End If
30250                 End If
30260               End If
30270             Else
30280               .saleCost_usd = .saleCost
30290             End If
30300           End If
30310         Else
30320           .saleCost_usd = 0#
30330         End If
30340       End If
30350     End If
30360     DoEvents

30370   End With

EXITP:
30380   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleCost_KeyDown(KeyCode As Integer, Shift As Integer)

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "saleCost_KeyDown"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strMsg As String
        Dim lngRecs As Long
        Dim lngX As Long
        Dim intRetVal As Integer

30410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30420   intShiftDown = (Shift And acShiftMask) > 0
30430   intAltDown = (Shift And acAltMask) > 0
30440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30470       With Me
30480         If .saleShareFace.Locked = False Then
                ' ** Open the Lot Information form
30490           If .saleType <> "Cost Adj." Then
30500             If IsNull(.saleAssetno) = False Then
30510               If .saleAssetno > 0& Then
30520                 DoCmd.Hourglass True
30530                 DoEvents
30540                 Set dbs = CurrentDb
                      ' ** qryJournal_Sale_01, linked to MasterAsset; check for existing Sales/Withdrawals,
                      ' ** by specified [actno], [astno], [jrnlid].
30550                 Set qdf = dbs.QueryDefs("qryJournal_Sale_02")  ' ** This excludes Asset Types '60','80', and '81',
30560                 With qdf.Parameters                            ' ** meaning they CAN enter multiple Sold!
30570                   ![jrnlid] = Me.saleID
30580                   ![actno] = Me.saleAccountNo
30590                   ![astno] = Me.saleAssetno
30600                 End With
30610                 Set rst = qdf.OpenRecordset()
30620                 If rst.BOF = True And rst.EOF = True Then
                        ' ** No records; all's well.
30630                   rst.Close
30640                   dbs.Close
30650                   gstrFormQuerySpec = .Parent.Name  ' ** In case it's been Nulled out by one of the Map forms.
30660                   If .Parent.DefSaleBtn = True Then
30670                     OpenLotInfoForm .Parent.DefSaleBtn, THIS_NAME  ' ** Module Function: modPurchaseSold.
                          ' ** Return Values:
                          ' **    0 OK.
                          ' **   -1 Input missing.
                          ' **   -2 No holdings.
                          ' **   -3 Insufficient holdings.
                          ' **   -4 Zero shares.
                          ' **   -9 Data problem.
                          '.Parent.DefSaleBtn = False  'THIS IS A CONFUSED MESS!! THE PARAMETER SEEMS TO BE DOING DOUBLE DUTY!
                          'IT NEEDS TO REMAIN BECAUSE IT'S USED IN OK AND CANCEL!
30680                   Else
30690                     OpenLotInfoForm blnBeenToLotInfo, THIS_NAME  ' ** Module Function: modPurchaseSold.
                          ' ** Return Values:
                          ' **    0 OK.
                          ' **   -1 Input missing.
                          ' **   -2 No holdings.
                          ' **   -3 Insufficient holdings.
                          ' **   -4 Zero shares.
                          ' **   -9 Data problem.
30700                     blnBeenToLotInfo = False
30710                   End If
30720                 Else
30730                   strMsg = vbNullString
30740                   rst.MoveLast
30750                   lngRecs = rst.RecordCount
30760                   rst.MoveFirst
30770                   If lngRecs > 1& Then
30780                     For lngX = 1& To lngRecs
30790                       strMsg = strMsg & rst![journal_USER] & ", "
30800                       If lngX < lngRecs Then rst.MoveNext
30810                     Next
30820                     strMsg = Trim(strMsg)                     ' ** Remove trailing space.
30830                     strMsg = Left(strMsg, (Len(strMsg) - 1))  ' ** Remove trailing comma.
30840                     rst.Close
30850                   Else
30860                     Select Case IsNull(rst![journal_USER])
                          Case True  ' ** Shouldn't happen, but has!
30870                       If IsNull(Me.saleUser) = True Then
30880                         Me.saleUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
30890                       End If
30900                       rst.Close
                            ' ** Update Journal, for journal_USER = Null.
30910                       Set qdf = dbs.QueryDefs("qryJournal_Sale_14")
30920                       qdf.Execute
30930                       strMsg = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
30940                     Case False
30950                       strMsg = rst![journal_USER]
30960                       rst.Close
30970                     End Select
30980                   End If
30990                   DoCmd.Hourglass False
31000                   strMsg = IIf(strMsg = CurrentUser, "You have", "The user" & IIf(lngRecs = 1, " " & strMsg & " has", "s " & strMsg & " have")) & _
                          " already entered a sale/withdrawal for this asset for this account!" & vbCrLf & vbCrLf & _
                          "You must post or delete the existing transaction before adding another one."
                        ' ** Internal Access Function: Trust Accountant login.
31010                   MsgBox strMsg, vbExclamation + vbOKOnly, "Multiple Sales/Withdrawals"
31020                   dbs.Close
31030                 End If
31040               End If
31050             End If
31060           End If
31070         End If
31080       End With
31090     End Select
31100   End If

        ' ** Shift keys.
31110   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31130       With Me
31140         intRetVal = 0
31150         If .salePCash.Enabled = True Then
31160           .salePCash.SetFocus
31170         ElseIf .saleICash.Enabled = True Then
31180           .saleICash.SetFocus
31190         Else
31200           .saleAssetDate.SetFocus
31210         End If
31220       End With
31230     End Select
31240   End If

        ' ** Ctrl-Shift keys.
31250   If intCtrlDown And (Not intAltDown) And intShiftDown Then
31260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31270       With Me
31280         intRetVal = 0
31290         If .salePCash.Enabled = False And .saleICash.Enabled = False Then
31300           .cmdCalendar2.SetFocus
31310         End If
31320       End With
31330     End Select
31340   End If

EXITP:
31350   Set rst = Nothing
31360   Set qdf = Nothing
31370   Set dbs = Nothing
31380   KeyCode = intRetVal
31390   Exit Sub

ERRH:
430     DoCmd.Hourglass False
440     intRetVal = 0
450     THAT_PROC = THIS_PROC
460     That_Erl = Erl
470     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
480     Resume EXITP

End Sub

Private Sub saleCost_Exit(Cancel As Integer)

31400 On Error GoTo ERRH

        Const THIS_PROC As String = "saleCost_Exit"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

31410   With Me

31420     blnContinue = True

31430     lngCurrID = .saleCurr_ID.Column(CBX_C_CURRID)

31440     If IsNumeric(.saleCost) = True Then
31450       If CDbl(Rem_Dollar(.saleCost.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
31460         blnContinue = False
31470         MsgBox "Cost is too large.", vbInformation + vbOKOnly, "Invalid Entry"
31480         .saleCost = vbNullString
31490         .saleCost_usd = Null
31500         blnFromElsewhere = True
31510         saleCost_AfterUpdate  ' ** Procedure: Below.
31520         Cancel = -1
31530         .saleCost.SetFocus
31540       End If
31550     Else
31560       blnContinue = False
31570       MsgBox "Cost must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
31580       Cancel = -1
31590       .saleCost.SetFocus
31600     End If

31610     If blnContinue = True Then
31620       intPos01 = InStr(1, .saleCost, ".", vbTextCompare)
31630       If intPos01 > 0 Then
31640         If Len(.saleCost) - intPos01 > 2 Then
31650           blnContinue = False
31660           MsgBox "You can only enter Cost with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
31670           .saleCost = vbNullString
31680           .saleCost_usd = Null
31690           blnFromElsewhere = True
31700           saleCost_AfterUpdate  ' ** Procedure: Below.
31710           Cancel = -1
31720           .saleCost.SetFocus
31730         End If
31740       End If
31750     End If

31760     If blnContinue = True Then
31770       subTaxCode4 blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modPurchaseSold.
31780     End If

31790     lngCurrID = .saleCurr_ID.Column(CBX_C_CURRID)

31800     If blnContinue = True Then
31810       gstrSaleCost = .saleCost.text
31820       gstrSaleCost = Rem_Dollar(gstrSaleCost, lngCurrID)  ' ** Module Function: modStringFuncs.
31830       If .saleCost <> .saleCost.OldValue Or IsNull(.saleCost.OldValue) Then
31840         SaleChanged True  ' ** Procedure: Below.
31850       End If
31860     End If

31870   End With

EXITP:
31880   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleCost_LostFocus()

31900 On Error GoTo ERRH

        Const THIS_PROC As String = "saleCost_LostFocus"

31910   VerifyAccountNo  ' ** Function: Below.

EXITP:
31920   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub saleCurr_ID_cmd_DblClick(Cancel As Integer)

32000 On Error GoTo ERRH

        Const THIS_PROC As String = "saleCurr_ID_cmd_DblClick"

32010   With Me
32020     If IsNull(.saleCurr_Date) = True Then
32030       .saleCurr_Date = .saleCurr_ID.Column(CBX_C_DATE)
32040     End If
32050     DoEvents
32060     Select Case .saleCurr_Date.Visible
          Case True
32070       .saleCurr_Date.Visible = False
32080     Case False
32090       .saleCurr_Date.Visible = True
32100     End Select
32110   End With

EXITP:
32120   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub saleCurr_ID_AfterUpdate()

32200 On Error GoTo ERRH

        Const THIS_PROC As String = "saleCurr_ID_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String, strTmp03 As String, intTmp04 As Integer, lngTmp06 As Long
        Dim lngX As Long

32210   With Me

32220     If IsNull(.saleCurr_ID) = False Then

32230       lngCurrID = .saleCurr_ID

32240       blnFound = False
32250       If lngCurrID <> 150& Then
32260         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
32270           arr_varNoChar = .Parent.NoChar_Get  ' ** Form Function: frmJournal.
32280           lngNoChars = UBound(arr_varNoChar, 2) + 1&
32290         End If
32300         For lngX = 0& To (lngNoChars - 1&)
32310           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
32320             blnFound = True
32330           End If
32340         Next
32350       End If

32360       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
32370         .saleICash.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
32380         .saleICash.FontSize = 10
32390         .saleICash.FontBold = True
32400         .salePCash.FontName = "Courier New"
32410         .salePCash.FontSize = 10
32420         .salePCash.FontBold = True
32430         .saleCost.FontName = "Courier New"
32440         .saleCost.FontSize = 10
32450         .saleCost.FontBold = True
32460       Case False
32470         If .saleICash.FontName <> "Fixedsys" Then
32480           .saleICash.FontName = "Fixedsys"
32490           .saleICash.FontSize = 9
32500           .saleICash.FontBold = False
32510           .salePCash.FontName = "Fixedsys"
32520           .salePCash.FontSize = 9
32530           .salePCash.FontBold = False
32540           .saleCost.FontName = "Fixedsys"
32550           .saleCost.FontSize = 9
32560           .saleCost.FontBold = False
32570         End If
32580       End Select

32590       If lngCurrID = 150& Then  ' ** USD.
32600         strTmp01 = "Currency"
32610         intTmp04 = 2
32620         lngTmp06 = CLR_WHT
32630       Else
32640         strTmp03 = .saleCurr_ID.Column(CBX_C_SYM)  ' ** Symbol.
32650         If .saleCurr_ID.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
32660           intTmp04 = 4
32670         Else
32680           intTmp04 = .saleCurr_ID.Column(CBX_C_DEC)  ' ** Decimal places.
32690         End If
32700         intLen = Len(strTmp03)
32710         Select Case intLen
              Case 1
                'strTmp03 = strTmp03
32720         Case 2
32730           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2)
32740         Case 3
32750           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3)
32760         Case 4
32770           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & "\" & Mid(strTmp03, 4)
32780         Case 5
32790           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & _
                  "\" & Mid(strTmp03, 4, 1) & "\" & Mid(strTmp03, 5)
32800         End Select
32810         Select Case intTmp04
              Case 0
32820           strTmp01 = strTmp03 & "#,##0;(" & strTmp03 & "#,##0);" & strTmp03 & "0;" & strTmp03 & "0"
32830         Case 1  ' ** None have this currently.
32840           strTmp01 = strTmp03 & "#,##0.0;(" & strTmp03 & "#,##0.0);" & strTmp03 & "0.0;" & strTmp03 & "0.0"
32850         Case 2
32860           strTmp01 = strTmp03 & "#,##0.00;(" & strTmp03 & "#,##0.00);" & strTmp03 & "0.00;" & strTmp03 & "0.00"
32870         Case 3
32880           strTmp01 = strTmp03 & "#,##0.000;(" & strTmp03 & "#,##0.000);" & strTmp03 & "0.000;" & strTmp03 & "0.000"
32890         End Select
32900         lngTmp06 = CLR_VLTGRN
32910       End If
32920       .saleICash.Format = strTmp01
32930       .saleICash.DecimalPlaces = intTmp04
32940       .saleICash.BackColor = lngTmp06
32950       .salePCash.Format = strTmp01
32960       .salePCash.DecimalPlaces = intTmp04
32970       .salePCash.BackColor = lngTmp06
32980       .saleCost.Format = strTmp01
32990       .saleCost.DecimalPlaces = intTmp04
33000       .saleCost.BackColor = lngTmp06

33010     End If

33020     If lngCurrID <> .saleCurr_ID.OldValue Then
33030       SaleChanged True  ' ** Procedure: Below.
33040     End If

33050   End With

EXITP:
33060   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleCurr_ID_NotInList(NewData As String, Response As Integer)

33100 On Error GoTo ERRH

        Const THIS_PROC As String = "saleCurr_ID_NotInList"

33110   With Me
33120     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
33130     .saleCurr_ID.SetFocus
33140     Response = acDataErrContinue
33150   End With

EXITP:
33160   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleCurr_ID_Exit(Cancel As Integer)

33200 On Error GoTo ERRH

        Const THIS_PROC As String = "saleCurr_ID_Exit"

33210   With Me
33220     If .saleCurr_ID <> .saleCurr_ID.OldValue Then
33230       SaleChanged True  ' ** Procedure: Below.
33240     End If
33250   End With

EXITP:
33260   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleDescription_AfterUpdate()

33300 On Error GoTo ERRH

        Const THIS_PROC As String = "saleDescription_AfterUpdate"

33310   With Me
33320     If InStr(.saleDescription, Chr(34)) > 0 Then
33330       MsgBox "Cannot have quote marks in text." & vbCrLf & _
              "They will be removed.", vbInformation + vbOKOnly, "Invalid Characters"
33340       .saleDescription = StringReplace(.saleDescription, Chr(34), "'")  ' ** Module Function: modStringFuncs.
33350     Else
            ' ** It's OK to process the sale.
33360       SaleChanged True  ' ** Procedure: Below.
33370     End If
33380   End With

EXITP:
33390   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub saleDescription_KeyDown(KeyCode As Integer, Shift As Integer)

33400 On Error GoTo ERRH

        Const THIS_PROC As String = "saleDescription_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

33410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
33420   intShiftDown = (Shift And acShiftMask) > 0
33430   intAltDown = (Shift And acAltMask) > 0
33440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
33450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
33460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33470       With Me
33480         intRetVal = 0
33490         If .tglSaleReinvest.Visible = True And tglSaleReinvest.Enabled = True Then
33500           .tglSaleReinvest.SetFocus
33510         ElseIf .cmdSaleLotInfo.Enabled = True Then
33520           .cmdSaleLotInfo.SetFocus
33530         ElseIf .cmbRevenueCodes.Visible = True And .cmbRevenueCodes.Enabled = True Then
33540           .cmbRevenueCodes.SetFocus
33550         ElseIf .cmbTaxCodes.Visible = True And .cmbTaxCodes.Enabled = True Then
33560           .cmbTaxCodes.SetFocus
33570         ElseIf .cmdSaleOK.Enabled = True Then
33580           .cmdSaleOK.SetFocus
33590         ElseIf .cmdSaleCancel.Enabled = True Then
33600           .cmdSaleCancel.SetFocus
33610         ElseIf .cmdSaleClose.Enabled = True Then
33620           .cmdSaleClose.SetFocus
33630         Else
33640           .saleTransDate.SetFocus
33650         End If
33660       End With
33670     End Select
33680   End If

        ' ** Shift keys.
33690   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33700     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33710       With Me
33720         intRetVal = 0
33730 On Error Resume Next
33740         .saleCost.SetFocus
33750         If ERR.Number <> 0 Then
                ' ** I don't know why SetFocus errors, and GoToControl doesn't!
33760 On Error GoTo ERRH
33770           DoCmd.GoToControl "saleCost"
33780         Else
33790 On Error GoTo ERRH
33800         End If
33810       End With
33820     End Select
33830   End If

EXITP:
33840   KeyCode = intRetVal
33850   Exit Sub

ERRH:
430     intRetVal = 0
440     Select Case ERR.Number
        Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Sub

Private Sub saleDescription_Exit(Cancel As Integer)

33900 On Error GoTo ERRH

        Const THIS_PROC As String = "saleDescription_Exit"

33910   VerifyAccountNo  ' ** Function: Below.

EXITP:
33920   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub tglSaleReinvest_GotFocus()

34000 On Error GoTo ERRH

        Const THIS_PROC As String = "tglSaleReinvest_GotFocus"

34010   With Me
34020     blnSaleReinvest_Focus = True
34030     Select Case .posted
          Case True
34040       .tglSaleReinvest_true_raised_focus_dots_img.Visible = True
34050       .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = False
34060     Case False
34070       .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = True
34080       .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
34090     End Select
34100     .tglSaleReinvest_false_raised_img.Visible = False
34110     .tglSaleReinvest_false_raised_focus_img.Visible = False
34120     .tglSaleReinvest_false_raised_focus_dots_img.Visible = False
34130     .tglSaleReinvest_false_sunken_focus_dots_img.Visible = False
34140     .tglSaleReinvest_false_raised_img_dis.Visible = False
34150     .tglSaleReinvest_true_raised_img.Visible = False
34160     .tglSaleReinvest_true_raised_focus_img.Visible = False
34170     .tglSaleReinvest_true_sunken_focus_dots_img.Visible = False
34180     .tglSaleReinvest_true_raised_img_dis.Visible = False
34190   End With

EXITP:
34200   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub tglSaleReinvest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

34300 On Error GoTo ERRH

        Const THIS_PROC As String = "tglSaleReinvest_MouseDown"

34310   With Me
34320     blnSaleReinvest_MouseDown = True
34330     Select Case .posted
          Case True
34340       .tglSaleReinvest_true_sunken_focus_dots_img.Visible = True
34350       .tglSaleReinvest_false_sunken_focus_dots_img.Visible = False
34360     Case False
34370       .tglSaleReinvest_false_sunken_focus_dots_img.Visible = True
34380       .tglSaleReinvest_true_sunken_focus_dots_img.Visible = False
34390     End Select
34400     .tglSaleReinvest_false_raised_img.Visible = False
34410     .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = False
34420     .tglSaleReinvest_false_raised_focus_img.Visible = False
34430     .tglSaleReinvest_false_raised_focus_dots_img.Visible = False
34440     .tglSaleReinvest_false_raised_img_dis.Visible = False
34450     .tglSaleReinvest_true_raised_img.Visible = False
34460     .tglSaleReinvest_true_raised_focus_img.Visible = False
34470     .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
34480     .tglSaleReinvest_true_raised_img_dis.Visible = False
34490   End With

EXITP:
34500   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub tglSaleReinvest_Click()
' ** I'm going to try using [posted] as a signal (since it's not used for anything else).

34600 On Error GoTo ERRH

        Const THIS_PROC As String = "tglSaleReinvest_Click"

34610   With Me
          ' ** Do I need to do anything with the images, or will that be handled by the other events?
34620     Select Case .posted
          Case True  ' ** If it's True, flip to False, and vice versa.
34630       blnGoToSaleReinvest = False
34640       .posted = False
34650       DoCmd.RunCommand acCmdSaveRecord
34660     Case False
34670       blnGoToSaleReinvest = True
34680       .posted = True
34690       DoCmd.RunCommand acCmdSaveRecord
34700     End Select
34710     SaleChanged True  ' ** Procedure: Below.
34720     DoEvents
34730   End With

EXITP:
34740   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub tglSaleReinvest_KeyDown(KeyCode As Integer, Shift As Integer)

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "tglSaleReinvest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

34810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
34820   intShiftDown = (Shift And acShiftMask) > 0
34830   intAltDown = (Shift And acAltMask) > 0
34840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
34850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
34870       With Me
34880         intRetVal = 0
34890         If .cmdSaleLotInfo.Enabled = True Then
34900           .cmdSaleLotInfo.SetFocus
34910         ElseIf .cmbRevenueCodes.Enabled = True Then
34920           .cmbRevenueCodes.SetFocus
34930         ElseIf .cmdSaleOK.Enabled = True Then
34940           .cmdSaleOK.SetFocus
34950         ElseIf .cmdSaleCancel.Enabled = True Then
34960           .cmdSaleCancel.SetFocus
34970         ElseIf .cmdSaleClose.Enabled = True Then
34980           .cmdSaleClose.SetFocus
34990         ElseIf .saleAccountNo.Enabled = True Then
35000           .saleAccountNo.SetFocus
35010         Else
35020           .saleTransDate.SetFocus
35030         End If
35040       End With
35050     End Select
35060   End If

        ' ** Shift keys.
35070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
35090       With Me
35100         intRetVal = 0
35110         .saleDescription.SetFocus
35120       End With
35130     End Select
35140   End If

EXITP:
35150   KeyCode = intRetVal
35160   Exit Sub

ERRH:
430     intRetVal = 0
440     THAT_PROC = THIS_PROC
450     That_Erl = Erl
460     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
470     Resume EXITP

End Sub

Private Sub tglSaleReinvest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35200 On Error GoTo ERRH

        Const THIS_PROC As String = "tglSaleReinvest_MouseMove"

35210   With Me
35220     If blnSaleReinvest_MouseDown = False Then
35230       Select Case .posted
            Case True
35240         Select Case blnSaleReinvest_Focus
              Case True
35250           .tglSaleReinvest_true_raised_focus_dots_img.Visible = True
35260           .tglSaleReinvest_true_raised_focus_img.Visible = False
35270         Case False
35280           .tglSaleReinvest_true_raised_focus_img.Visible = True
35290           .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
35300         End Select
35310         .tglSaleReinvest_false_raised_focus_img.Visible = False
35320         .tglSaleReinvest_false_raised_focus_dots_img.Visible = False
35330       Case False
35340         Select Case blnSaleReinvest_Focus
              Case True
35350           .tglSaleReinvest_false_raised_focus_dots_img.Visible = True
35360           .tglSaleReinvest_false_raised_focus_img.Visible = False
35370         Case False
35380           .tglSaleReinvest_false_raised_focus_img.Visible = True
35390           .tglSaleReinvest_false_raised_focus_dots_img.Visible = False
35400         End Select
35410         .tglSaleReinvest_true_raised_focus_img.Visible = False
35420         .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
35430       End Select
35440       .tglSaleReinvest_false_raised_img.Visible = False
35450       .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = False
35460       .tglSaleReinvest_false_raised_img_dis.Visible = False
35470       .tglSaleReinvest_false_sunken_focus_dots_img.Visible = False
35480       .tglSaleReinvest_true_raised_img.Visible = False
35490       .tglSaleReinvest_true_sunken_focus_dots_img.Visible = False
35500       .tglSaleReinvest_true_raised_img_dis.Visible = False
35510     End If
35520   End With

EXITP:
35530   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub tglSaleReinvest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

35600 On Error GoTo ERRH

        Const THIS_PROC As String = "tglSaleReinvest_MouseUp"

35610   With Me
35620     Select Case .posted
          Case True
35630       .tglSaleReinvest_true_raised_focus_dots_img.Visible = True
35640       .tglSaleReinvest_false_raised_focus_dots_img.Visible = False
35650     Case False
35660       .tglSaleReinvest_false_raised_focus_dots_img.Visible = True
35670       .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
35680     End Select
35690     .tglSaleReinvest_false_raised_img.Visible = False
35700     .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = False
35710     .tglSaleReinvest_false_raised_focus_img.Visible = False
35720     .tglSaleReinvest_false_raised_img_dis.Visible = False
35730     .tglSaleReinvest_false_sunken_focus_dots_img.Visible = False
35740     .tglSaleReinvest_true_raised_img.Visible = False
35750     .tglSaleReinvest_true_raised_focus_img.Visible = False
35760     .tglSaleReinvest_true_sunken_focus_dots_img.Visible = False
35770     .tglSaleReinvest_true_raised_img_dis.Visible = False
35780     blnSaleReinvest_MouseDown = False
35790   End With

EXITP:
35800   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub tglSaleReinvest_LostFocus()

35900 On Error GoTo ERRH

        Const THIS_PROC As String = "tglSaleReinvest_LostFocus"

35910   With Me
35920     Select Case .posted
          Case True
35930       .tglSaleReinvest_true_raised_img.Visible = True
35940       .tglSaleReinvest_false_raised_img.Visible = False
35950     Case False
35960       .tglSaleReinvest_false_raised_img.Visible = True
35970       .tglSaleReinvest_true_raised_img.Visible = False
35980     End Select
35990     .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = False
36000     .tglSaleReinvest_false_raised_focus_img.Visible = False
36010     .tglSaleReinvest_false_raised_focus_dots_img.Visible = False
36020     .tglSaleReinvest_false_sunken_focus_dots_img.Visible = False
36030     .tglSaleReinvest_false_raised_img_dis.Visible = False
36040     .tglSaleReinvest_true_raised_focus_img.Visible = False
36050     .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
36060     .tglSaleReinvest_true_sunken_focus_dots_img.Visible = False
36070     .tglSaleReinvest_true_raised_img_dis.Visible = False
36080     blnSaleReinvest_Focus = False
36090   End With

EXITP:
36100   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub cmdSaleLotInfo_Click()

36200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaleLotInfo_Click"

36210   SaleLotInfo_PS blnClickedLotInfo, blnBeenToLotInfo, blnGoingToReport, blnGoneToReport, blnGTR_Emblem, lngGTR_ID, Me  ' ** Module Procedure: modPurchaseSold.

EXITP:
36220   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Private Sub cmdSaleLotInfo_KeyDown(KeyCode As Integer, Shift As Integer)
' ** See Note, above.

36300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaleLotInfo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

36310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
36320   intShiftDown = (Shift And acShiftMask) > 0
36330   intAltDown = (Shift And acAltMask) > 0
36340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
36350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
36360     Select Case intRetVal
          Case vbKeyTab
36370       With Me
36380         intRetVal = 0
36390         If .cmbRevenueCodes.Enabled = True Then
36400           .cmbRevenueCodes.SetFocus
36410         ElseIf .cmdSaleOK.Enabled = True Then
36420           .cmdSaleOK.SetFocus
36430         ElseIf .cmdSaleCancel.Enabled = True Then
36440           .cmdSaleCancel.SetFocus
36450         ElseIf .cmdSaleClose.Enabled = True Then
36460           .cmdSaleClose.SetFocus
36470         ElseIf .saleAccountNo.Enabled = True Then
36480           .saleAccountNo.SetFocus
36490         Else
36500           .saleTransDate.SetFocus
36510         End If
36520       End With
36530     End Select
36540   End If

        ' ** Shift keys.
36550   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
36560     Select Case intRetVal
          Case vbKeyTab
36570       With Me
36580         intRetVal = 0
36590         If .tglSaleReinvest.Enabled = True Then
36600           .tglSaleReinvest.SetFocus
36610         Else
36620           .saleDescription.SetFocus
36630         End If
36640       End With
36650     End Select
36660   End If

EXITP:
36670   KeyCode = intRetVal
36680   Exit Sub

ERRH:
430     intRetVal = 0
440     Select Case ERR.Number
        Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Enter()

36700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Enter"

36710   With Me
36720     Select Case IsNull(.cmbRevenueCodes)
          Case True
36730       ExpandCombo .cmbRevenueCodes  ' ** Module Procedure: modUtilities.
36740     Case False
36750       If .cmbRevenueCodes = REVID_INC Or .cmbRevenueCodes = REVID_EXP Then
36760         ExpandCombo .cmbRevenueCodes  ' ** Module Procedure: modUtilities.
36770       End If
36780     End Select
36790   End With

EXITP:
36800   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

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

36900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Change"

36910   With Me
36920     If .cmbRevenueCodes <> .cmbRevenueCodes.OldValue Then
36930       SaleChanged True  ' ** Procedure: Below.
36940     End If
36950   End With

EXITP:
36960   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

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

37000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_AfterUpdate"

        Dim strRevCode As String, lngTaxcode As Long

        ' ** Make sure tax code matches selected revenue code.
37010   With Me
37020     If IsNull(.cmbRevenueCodes) = False Then
37030       If .cmbRevenueCodes > 0 Then
37040         If gblnLinkRevTaxCodes = True Then
37050           strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
37060           If IsNull(.cmbTaxCodes) = True Then
37070             If strRevCode = "I" Then
37080               .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
37090             ElseIf strRevCode = "E" Then
37100               .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
37110             End If
37120           Else
37130             lngTaxcode = .cmbTaxCodes
37140             If lngTaxcode = 0& Then
37150               If strRevCode = "I" Then
37160                 .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
37170               ElseIf strRevCode = "E" Then
37180                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
37190               End If
37200             Else
37210               If strRevCode = "I" And .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
                      'MsgBox "This will change..."
37220                 .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
37230               ElseIf strRevCode = "E" And .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
                      'MsgBox "This will change..."
37240                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
37250               End If
37260             End If
37270           End If
37280         End If
37290       Else
37300         .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37310       End If
37320     Else
37330       .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37340     End If
37350   End With

EXITP:
37360   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub cmbRevenueCodes_LostFocus()

37400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_LostFocus"

37410   VerifyAccountNo  ' ** Function: Below.

EXITP:
37420   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub cmbTaxCodes_Enter()

37500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Enter"

37510   With Me
37520     Select Case IsNull(.cmbTaxCodes)
          Case True
37530       ExpandCombo .cmbTaxCodes  ' ** Module Procedure: modUtilities.
37540     Case False
37550       If .cmbTaxCodes = TAXID_INC Or .cmbTaxCodes = TAXID_DED Then
37560         ExpandCombo .cmbTaxCodes  ' ** Module Procedure: modUtilities.
37570       End If
37580     End Select
37590   End With

EXITP:
37600   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

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

37700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Change"

37710   With Me
37720     If .cmbTaxCodes <> .cmbTaxCodes.OldValue Then
37730       SaleChanged True  ' ** Procedure: Below.
37740     End If
37750   End With

EXITP:
37760   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

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

37800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_AfterUpdate"

        Dim strRevCode As String, lngTaxcode As Long

        ' ** Make sure revenue code matches selected tax code.
37810   With Me
37820     If IsNull(.cmbTaxCodes) = False Then
37830       If .cmbTaxCodes > 0& Then
37840         If gblnLinkRevTaxCodes = True Then
37850           lngTaxcode = .cmbTaxCodes
37860           strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
37870           If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
37880             If IsNull(.cmbRevenueCodes) = True Then
37890               .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37900             Else
37910               If .cmbRevenueCodes = 0 Then
37920                 .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37930               Else
37940                 If strRevCode = "E" Then
                        'MsgBox "This will change..."
37950                   .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37960                 End If
37970               End If
37980             End If
37990           ElseIf .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
38000             If IsNull(.cmbRevenueCodes) = True Then
38010               .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
38020             Else
38030               If .cmbRevenueCodes = 0 Then
38040                 .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
38050               Else
38060                 If strRevCode = "I" Then
                        'MsgBox "This will change..."
38070                   .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
38080                 End If
38090               End If
38100             End If
38110           End If
38120         End If
38130       Else
38140         .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
38150       End If
38160     Else
38170       .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
38180     End If
38190   End With

EXITP:
38200   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub cmbTaxCodes_LostFocus()

38300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_LostFocus"

38310   VerifyAccountNo  ' ** Function: Below.

EXITP:
38320   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Public Sub SaleChanged(blnChanged As Boolean)

38400 On Error GoTo ERRH

        Const THIS_PROC As String = "SaleChanged"

38410   With Me
38420     Select Case blnChanged
          Case True
38430       gblnSaleChanged = True
38440       .NavigationButtons = False
38450       DoCmd.SelectObject acForm, .Parent.Name, False
38460       With .Parent
38470 On Error Resume Next
38480         .FocusHolder.SetFocus
38490 On Error GoTo ERRH
38500         .opgJournal.Enabled = False
38510         .opgJournal_optSale_lbl_box.Visible = True
38520         .cmdSwitch.Enabled = False
38530         .cmdSwitch_raised_img_dis.Visible = True
38540         .cmdSwitch_raised_img.Visible = False
38550         .cmdSwitch_raised_semifocus_dots_img.Visible = False
38560         .cmdSwitch_raised_focus_img.Visible = False
38570         .cmdSwitch_raised_focus_dots_img.Visible = False
38580         .cmdSwitch_sunken_focus_dots_img.Visible = False
38590 On Error Resume Next
38600         .frmJournal_Sub4_Sold.SetFocus
38610 On Error GoTo ERRH
38620       End With
38630       .cmdSaleClose.Enabled = False
38640       .cmdSaleOK.Enabled = True
38650       .cmdSaleCancel.Enabled = True
38660       .Parent.NavVis False  ' ** Form Procedure: frmJournal.
38670     Case False
38680       gblnSaleChanged = False
38690       .NavigationButtons = True
38700       DoCmd.SelectObject acForm, .Parent.Name, False
38710       With .Parent
38720         .opgJournal.Enabled = True
38730         .opgJournal_optSale_lbl_box.Visible = False
38740         .cmdSwitch.Enabled = True
38750         .cmdSwitch_raised_img.Visible = True
38760         .cmdSwitch_raised_img_dis.Visible = False
38770         .cmdSwitch_raised_semifocus_dots_img.Visible = False
38780         .cmdSwitch_raised_focus_img.Visible = False
38790         .cmdSwitch_raised_focus_dots_img.Visible = False
38800         .cmdSwitch_sunken_focus_dots_img.Visible = False
38810       End With
38820       .cmdSaleClose.Enabled = True
38830       .cmdSaleOK.Enabled = False
38840       .cmdSaleCancel.Enabled = False
38850       .cmdSaleLotInfo.Enabled = True
38860       .Parent.NavVis True  ' ** Form Procedure: frmJournal.
38870     End Select
38880   End With

EXITP:
38890   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Public Function VerifyAccountNo() As Boolean

38900 On Error GoTo ERRH

        Const THIS_PROC As String = "VerifyAccountNo"

        Dim strAccountNo As String, strControl As String
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

38910   blnRetVal = True

38920 On Error Resume Next
38930   strControl = Screen.ActiveControl.Name
38940 On Error GoTo ERRH

38950   With Me
38960     If ((gblnSaleValidated = False And gblnSaleChanged = True) And (lngErrCnt < 3&)) Then
38970       If strControl <> "saleAccountno" And blnAccountNoHasFocus = False And .saleAccountNo.Enabled = True Then
38980         blnAccountNoHasFocus = True
38990         .saleAccountNo.SetFocus
39000       End If
39010       Select Case IsNull(.saleAccountNo)
            Case True
39020         blnRetVal = False
39030         lngErrCnt = lngErrCnt + 1&
39040         MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Entry Required"
              'DoCmd.CancelEvent
39050         gblnSaleValidated = False
39060       Case False
              ' ** Check to see if it is a valid account.
39070         lngErrCnt = 0&
39080         Select Case gblnAccountNoWithType
              Case True
39090           strAccountNo = Left(.saleAccountNo, 2) & Mid(.saleAccountNo, 3)
39100         Case False
39110           strAccountNo = .saleAccountNo
39120         End Select
39130         varTmp00 = DLookup("[accountno]", "account", "[accountno] = '" & strAccountNo & "' And ([dateclosed] Is Null)")
39140         Select Case IsNull(varTmp00)
              Case True
39150           blnRetVal = False
39160           MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Invalid Entry"
39170           gblnSaleValidated = False
39180         Case False
39190           gblnSaleValidated = True
39200           blnStartTrans = False
39210         End Select
39220       End Select
39230     Else
39240       lngErrCnt = 0&
39250     End If
39260   End With

EXITP:
39270   VerifyAccountNo = blnRetVal
39280   Exit Function

ERRH:
430     blnRetVal = False
440     THAT_PROC = THIS_PROC
450     That_Erl = Erl
460     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
470     Resume EXITP

End Function

Private Sub SaleDistribute()
' ** Called by:
' **   cmdSaleOK_Click, Above.

39300 On Error GoTo ERRH

        Const THIS_PROC As String = "SaleDistribute"

39310   DistributeCost Me  ' ** Module Procedure: modPurchaseSold.

EXITP:
39320   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
460     Resume EXITP

End Sub

Public Sub subSetAssetCombo()

39400 On Error GoTo ERRH

        Const THIS_PROC As String = "subSetAssetCombo"

39410   With Me

39420     gstrFormQuerySpec = .Parent.Name  ' ** In case it's been Nulled out by one of the Map forms.
39430     .saleAssetno.Requery

39440   End With

EXITP:
39450   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

39500 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

39510   Select Case intWhere
        Case acCmdRecordsGoToFirst
39520     DoCmd.RunCommand acCmdRecordsGoToFirst
39530   Case acCmdRecordsGoToLast
39540     DoCmd.RunCommand acCmdRecordsGoToLast
39550   Case acCmdRecordsGoToPrevious
39560     DoCmd.RunCommand acCmdRecordsGoToPrevious
39570   Case acCmdRecordsGoToNext
39580     DoCmd.RunCommand acCmdRecordsGoToNext
39590   Case acCmdRecordsGoToNew
39600     DoCmd.RunCommand acCmdRecordsGoToNew
39610   Case Else
39620     If IsMissing(varID) = False Then
39630       With Me
39640         Set rst = .RecordsetClone
39650         With rst
39660           .FindFirst "[saleID] = " & CStr(varID)
39670           If .NoMatch = False Then
39680             Me.Bookmark = .Bookmark
39690           End If
39700           .Close
39710         End With
39720       End With
39730     End If
39740   End Select

EXITP:
39750   Set rst = Nothing
39760   Exit Sub

ERRH:
430     Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
440     Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
450     Case Else
460       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
470     End Select
480     Resume EXITP

End Sub

Public Sub SoldCostedSet(blnOKCancel As Boolean)

39800 On Error GoTo ERRH

        Const THIS_PROC As String = "SoldCostedSet"

39810   SoldCostedSet_PS blnOKCancel, blnSoldCosted, Me  ' ** Module Procedure: modPurchaseSold.

EXITP:
39820   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
460     Resume EXITP

End Sub

Public Sub RecalcPaidTotal()

39900 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcPaidTotal"

        Dim varTmp00 As Variant

39910   With Me

39920     .cmdPaidTotal.Enabled = False
39930     varTmp00 = DCount("*", "journal", "[journaltype] = 'Paid'")
39940     If IsNull(varTmp00) = False Then
39950       If varTmp00 > 0 Then .cmdPaidTotal.Enabled = True
39960     End If

39970   End With

EXITP:
39980   Exit Sub

ERRH:
430     THAT_PROC = THIS_PROC
440     That_Erl = Erl
450     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
460     Resume EXITP

End Sub
