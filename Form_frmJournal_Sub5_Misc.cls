VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournal_Sub5_Misc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournal_Sub5_Misc"

'VGC 10/26/2017: CHANGES!

' #########################
' ## Use VBA_RenumErrh().  38620
' #########################

' ** Shortcut Alt keys responsive from this form:
' **   Map:              A {cmdMiscMap}
' **   Close:            C {cmdMiscClose}
' **   Dividend:         D {opgJournal_optDividend on frmJournal}
' **   Cancel:           E {cmdMiscCancel}
' **   STCGL Map:        G {cmdMiscMap_STCGL}
' **   Print Check:      H {chkChecks}
' **   Interest:         I {opgJournal_optInterest on frmJournal}
' **   LTCL Map:         L {cmdMiscMap_LTCL}
' **   Miscellaneous:    M {opgJournal_optMisc on frmJournal}
' **   Inc/Exp Codes:    N {cmbRevenueCodes}
' **   OK:               O {cmdMiscOK}
' **   Purchase:         P {opgJournal_optPurchase on frmJournal}
' **   Sale:             S {opgJournal_optSale on frmJournal}
' **   Tax Codes:        T {cmbTaxCodes}
' **   New Recur:        W {cmdMiscRecurNew}
' **   Exit:             X {cmdMiscClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Next Record:      N {MoveRec}
' **   Previous Record:  P {MoveRec}
' **   Switch:           T {cmdSwitch on frmJournal}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Mark Reinvest:    R {tglMiscReinvest}
' **   Mark Sale:        S {tglMiscSale}
' **   Close Timer:      T {Form_KeyDown}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}

' ** Paid is always EXPENSE for Tax Codes.
' ** Received is always INCOME for Tax Codes.
' ** Misc. can be either INCOME or EXPENSE for Tax Codes.

' **********************************************************************
' ** NOTE: Enter key only invokes command button on OK and Cancel.
' **       Though Enter would usually invoke any command button,
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

' ** Event Precedence:
' **   Event       Event Property                 It Occurs
' **   ==========  =============================  ================================================================================
' **   Activate    OnActivate (forms, reports)    When a form or report becomes the active window.
' **   Deactivate  OnDeactivate (forms, reports)  When a different Microsoft Access window becomes the active window, but before
' **                                              the window becomes the active window. The Deactivate event does not occur when
' **                                              when the focus moves to another application's window, a dialog box, or a
' **                                              pop-up form.
' **   Enter       OnEnter (controls)             Before a control actually receives the focus, either from a control on the same
' **                                              form or when the form opens. This event occurs before the GotFocus event.
' **   GotFocus    OnGotFocus (forms, controls)   When a control, or a form with no active or enabled controls, receives the
' **                                              focus. A form can get the focus only if all visible controls on a form are
' **                                              disabled, or there are no controls on the form.
' **   Exit        OnExit (controls)              Just before a control loses the focus to another control on the same form. This
' **                                              event occurs before the LostFocus event.
' **   LostFocus   OnLostFocus (forms, controls)  When a form or control loses the focus. A form can have the focus only if all
' **                                              visible controls on a form are disabled, or there are no controls on the form.

' ** Event Precedence Example:
' **   Exit (for the control on the main form)
' **   \
' **   LostFocus (for the control on the main form)
' **   \
' **   Enter (for the subform control)
' **   \
' **   Exit (for the control on the subform that had the focus)
' **   \
' **   LostFocus (for the control on the subform that had the focus)
' **   \
' **   Enter (for the control on the subform that the focus moved to)
' **   \
' **   GotFocus (for the control on the subform that the focus moved to)

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Memo field maximum characters.
Private Const MEMO_MAX As Integer = 40

' ** Combo box column constants: cmbRecurringItems.
'Private Const CBX_R_NAME As Integer = 0
'Private Const CBX_R_TYPE As Integer = 1
Private Const CBX_R_RID  As Integer = 2

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

' ** Though miscCurr_ID has a preference listed, it's not currently being used.
' ** I'm not sure how it can be used anyway.

' ** Combo box column constants: miscCurr_ID.Column
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
Private blnMiscMap_LTCL_Focus As Boolean, blnMiscMap_LTCL_MouseDown As Boolean
Private blnMiscMap_STCGL_Focus As Boolean, blnMiscMap_STCGL_MouseDown As Boolean
Private blnMiscSale_Focus As Boolean, blnMiscSale_MouseDown As Boolean
Private blnMiscReinvest_Focus As Boolean, blnMiscReinvest_MouseDown As Boolean

Private lngCurrID As Long, lngCurrID_Left As Long, lngICash_Left As Long, lngPCash_Left As Long
Private blnGoToSaleForm As Boolean, blnCmdButton As Boolean, blnGoToRecReinvest As Boolean
Private blnAccountNoErr As Boolean, blnStartTrans As Boolean, blnDefTypeAssigned As Boolean
Private blnFromAcctNoKeyDown As Boolean, blnFromFormError As Boolean, blnMapNewRec As Boolean
Private datPostingDate As Date, lngDateCnt As Long, lngErrCnt As Long
Private blnAccountNoHasFocus As Boolean, blnTypeIsNull As Boolean, blnSpecialCap As Boolean, intSpecialCapOpt As Integer
Private strMemo_Orig As String, strMemo_New As String
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
180           blnAccountNoErr = False: blnStartTrans = False: blnDefTypeAssigned = False
190           blnFromAcctNoKeyDown = False: blnFromFormError = False
200           gblnSwitchTo = False

210           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
220           blnMiscMap_LTCL_Focus = False: blnMiscMap_LTCL_MouseDown = False
230           blnMiscMap_STCGL_Focus = False: blnMiscMap_STCGL_MouseDown = False
240           blnMiscSale_Focus = False: blnMiscSale_MouseDown = False
250           blnMiscReinvest_Focus = False: blnMiscReinvest_MouseDown = False
260           blnMapNewRec = False

270           lngCurrID_Left = .miscCurr_ID.Left
280           lngICash_Left = .miscICash.Left
290           lngPCash_Left = .miscPCash.Left

300           .cmdLock.Enabled = False
310           .cmdLock_open_raised_img.Visible = False
320           .cmdLock_open_raised_img_dis.Visible = False
330           .cmdLock_closed_raised_img.Visible = False

340         Else
350           Cancel = -1
360         End If
370       Else
380         Cancel = -1
390       End If
400     End With

        ' ** Begins with:
        ' **   .miscAssetNo.Visible = False
        ' **   .miscAssetNo.Enabled = False  Never becomes True
        ' **   .miscAssetNo.Locked = True    Never becomes False

410     If Cancel = -1 Then
420       Beep
430       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
440     End If

EXITP:
450     Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
490     Resume EXITP

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

550       gblnMiscValidated = False
560       gblnMiscChanged = False
570       blnGoToSaleForm = False: blnGoToRecReinvest = False  ' ** Set the switch to false until they push the button.
580       blnCmdButton = False

590       datPostingDate = GetPostDate  ' ** Module Function: modUtilities.

          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
600       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

          ' ** Do this whether or not Currency is visible.
610       CLR_DISABLED_FG = CLR_DKGRY
620       CLR_DISABLED_BG = CLR_LTTEAL

630       .miscCurr_ID.Enabled = False
640       .miscCurr_ID.Locked = True
650       .miscCurr_ID_cmd.Visible = True
660       .miscCurr_ID_cmd.Enabled = True
670       .miscCurr_ID.ForeColor = CLR_DISABLED_FG
680       .miscCurr_ID.BackColor = CLR_VLTGRN
690       miscCurr_ID_AfterUpdate  ' ** Procedure: Below.

          ' ** Set the default Posting Date (transdate), entered by the user on opening.
700       .miscTransDate.DefaultValue = "#" & CDate(Format(datPostingDate, "mm/dd/yyyy")) & "#"

710       gstrFormQuerySpec = .Parent.Name  ' ** Make sure this is set for the assetno combo box.

          ' ** Make sure the SpecialCap options are set.
720       If gstrCo_Name = vbNullString Then
730         CoOptions_Read  ' ** Module Function: modStartupFunctions.
740       End If

          ' ** Load SpecialCapGainLoss setting.
750       blnSpecialCap = gblnSpecialCapGainLoss
760       intSpecialCapOpt = gintSpecialCapGainLossOpt

770       .SpecialCapGainLoss = blnSpecialCap

780       If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
790         .cmdMiscMap_LTCL.Enabled = True
800         .cmdMiscMap_LTCL.Visible = True
810         .cmdMiscMap_LTCL_raised_img.Visible = True
820         .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
830         .cmdMiscMap_LTCL_raised_focus_img.Visible = False
840         .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
850         .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
860         .cmdMiscMap_LTCL_raised_img_dis.Visible = False
870         .cmdMiscMap_STCGL.Enabled = True
880         .cmdMiscMap_STCGL.Visible = True
890         .cmdMiscMap_STCGL_raised_img.Visible = True
900         .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = False
910         .cmdMiscMap_STCGL_raised_focus_img.Visible = False
920         .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = False
930         .cmdMiscMap_STCGL_sunken_focus_dots_img.Visible = False
940         .cmdMiscMap_STCGL_raised_img_dis.Visible = False
950       Else
960         .cmdMiscMap_LTCL.Enabled = False
970         .cmdMiscMap_LTCL.Visible = False
980         .cmdMiscMap_LTCL_raised_img.Visible = False
990         .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
1000        .cmdMiscMap_LTCL_raised_focus_img.Visible = False
1010        .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
1020        .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
1030        .cmdMiscMap_LTCL_raised_img_dis.Visible = False
1040        .cmdMiscMap_STCGL.Enabled = False
1050        .cmdMiscMap_STCGL.Visible = False
1060        .cmdMiscMap_STCGL_raised_img.Visible = False
1070        .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = False
1080        .cmdMiscMap_STCGL_raised_focus_img.Visible = False
1090        .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = False
1100        .cmdMiscMap_STCGL_sunken_focus_dots_img.Visible = False
1110        .cmdMiscMap_STCGL_raised_img_dis.Visible = False
1120      End If

          ' ** Wingdings arrows:
          ' **   à  Pointing Right.
          ' **   ß  Pointing Left.
1130      .MiscXFer_ItoP_lbl.Visible = False
1140      .MiscXFer_PtoI_lbl.Visible = False

1150      .NavigationButtons = True

1160      With .Parent
1170        .opgJournal.Enabled = True
1180        .cmdSwitch.Enabled = True
1190        .cmdSwitch_raised_img.Visible = True
1200        .cmdSwitch_raised_img_dis.Visible = False
1210        .cmdSwitch_raised_semifocus_dots_img.Visible = False
1220        .cmdSwitch_raised_focus_img.Visible = False
1230        .cmdSwitch_raised_focus_dots_img.Visible = False
1240        .cmdSwitch_sunken_focus_dots_img.Visible = False
1250      End With
1260      .cmdMiscClose.Enabled = True
1270      .cmdMiscOK.Enabled = False
1280      .cmdMiscCancel.Enabled = False
1290      .Parent.NavVis True  ' ** Form Procedure: frmJournal.

1300      If .NewRecord = True Then
1310        .tglMiscSale.Enabled = False
1320        .tglMiscSale_false_raised_img_dis.Visible = True
1330        .tglMiscSale_false_raised_img.Visible = False
1340        .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
1350        .tglMiscSale_false_raised_focus_img.Visible = False
1360        .tglMiscSale_false_raised_focus_dots_img.Visible = False
1370        .tglMiscSale_false_sunken_focus_dots_img.Visible = False
1380        .tglMiscSale_true_raised_img.Visible = False
1390        .tglMiscSale_true_raised_focus_img.Visible = False
1400        .tglMiscSale_true_raised_focus_dots_img.Visible = False
1410        .tglMiscSale_true_sunken_focus_dots_img.Visible = False
1420        .tglMiscSale_true_raised_img_dis.Visible = False
1430        .tglMiscReinvest.Enabled = False
1440        .tglMiscReinvest_false_raised_img_dis.Visible = True
1450        .tglMiscReinvest_false_raised_img.Visible = False
1460        .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
1470        .tglMiscReinvest_false_raised_focus_img.Visible = False
1480        .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
1490        .tglMiscReinvest_false_sunken_focus_dots_img.Visible = False
1500        .tglMiscReinvest_true_raised_img.Visible = False
1510        .tglMiscReinvest_true_raised_focus_img.Visible = False
1520        .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
1530        .tglMiscReinvest_true_sunken_focus_dots_img.Visible = False
1540        .tglMiscReinvest_true_raised_img_dis.Visible = False
1550      End If

1560      If gblnIncomeTaxCoding Then
1570        .cmbTaxCodes.Visible = True
1580        .cmbTaxCodes_lbl2.Visible = False
1590      Else
1600        .cmbTaxCodes.Visible = False
1610        .cmbTaxCodes_lbl2.Visible = True
1620      End If

1630      If gblnRevenueExpenseTracking Then
1640        .cmbRevenueCodes.Visible = True
1650        .cmbRevenueCodes_lbl2.Visible = False
1660      Else
1670        .cmbRevenueCodes.Visible = False
1680        .cmbRevenueCodes_lbl2.Visible = True
1690      End If

1700    End With

EXITP:
1710    Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
490     Resume EXITP

End Sub

Private Sub Form_Current()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim varTmp00 As Variant

1810    With Me

1820      DoCmd.SelectObject acForm, .Parent.Name, False

1830      .MiscXFer_ItoP_lbl.Visible = False
1840      .MiscXFer_PtoI_lbl.Visible = False

1850      blnDefTypeAssigned = False
1860      blnStartTrans = False
1870      lngErrCnt = 0&

1880      .cmbAccountHelper = Null

1890      If .miscCurr_ID.Visible = True Then
1900        If .miscCurr_Date.Visible = True Then .miscCurr_Date.Visible = False
1910        .miscICash_usd = Null
1920        .miscICash_usd.Visible = False
1930        .miscICash.Format = "Currency"
1940        .miscICash.DecimalPlaces = 2
1950        .miscICash.BackColor = CLR_WHT
1960        .miscPCash_usd = Null
1970        .miscPCash_usd.Visible = False
1980        .miscPCash.Format = "Currency"
1990        .miscPCash.DecimalPlaces = 2
2000        .miscPCash.BackColor = CLR_WHT
2010        If lngNoChars = 0& Or IsEmpty(arr_varNoChar) Then
2020          arr_varNoChar = .Parent.NoChar_Get  ' ** Form Function: frmJournal.
2030          lngNoChars = UBound(arr_varNoChar, 2) + 1&
2040        End If
2050        IncludeCurrencyMisc_Sub5 True, lngTpp, lngICash_Left, lngPCash_Left, lngCurrID_Left, Me ' ** Module Procedure: modJrnlSub5MiscFuncs.
2060      Else
2070        IncludeCurrencyMisc_Sub5 False, lngTpp, lngICash_Left, lngPCash_Left, lngCurrID_Left, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
2080      End If

2090      gstrFormQuerySpec = .Parent.Name  ' ** Make sure this is set for the assetno combo box.

2100      Select Case IsNull(.miscAccountNo_Data)
          Case True
2110        .miscAccountNo = vbNullString
2120        .miscAccountNo.Enabled = True
2130        .miscAccountNo.BorderColor = CLR_LTBLU2
2140        .miscAccountNo.BackStyle = acBackStyleNormal
2150        .miscAccountNo_lbl.BackStyle = acBackStyleNormal
2160        .miscAccountNo_lbl_box.Visible = False
2170        .cmdLock.Enabled = False
2180        .cmdLock_open_raised_img_dis.Visible = True
2190        .cmdLock_open_raised_img.Visible = False
2200        .cmdLock_closed_raised_img.Visible = False
2210        .cmbAccountHelper.Enabled = True
2220        .cmbAccountHelper.BorderColor = CLR_LTBLU2
2230        .cmbAccountHelper.BackStyle = acBackStyleNormal
2240        .miscAccountNo.SetFocus
2250        .cmdMiscMap.Enabled = True
2260        If gblnGoToReport = False Then
2270          If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
2280            .cmdMiscMap_LTCL.Enabled = True
2290            .cmdMiscMap_LTCL_raised_img.Visible = True
2300            .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
2310            .cmdMiscMap_LTCL_raised_focus_img.Visible = False
2320            .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
2330            .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
2340            .cmdMiscMap_LTCL_raised_img_dis.Visible = False
2350            .cmdMiscMap_STCGL.Enabled = True
2360            .cmdMiscMap_STCGL_raised_img.Visible = True
2370            .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = False
2380            .cmdMiscMap_STCGL_raised_focus_img.Visible = False
2390            .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = False
2400            .cmdMiscMap_STCGL_sunken_focus_dots_img.Visible = False
2410            .cmdMiscMap_STCGL_raised_img_dis.Visible = False
2420          End If  ' ** blnSpecialCap.
2430        End If
2440        blnAccountNoErr = False
2450      Case False
2460        .miscAccountNo = .miscAccountNo_Data
2470        If IsNull(.miscAccountNo) = False Then
2480          gstrMiscAccountNumber = .miscAccountNo
2490        End If
2500        .miscTransDate.SetFocus  ' ** Make sure it's not on AccountNo before disabling.
2510        .miscAccountNo.Enabled = False
2520        .miscAccountNo.BorderColor = WIN_CLR_DISR
2530        .miscAccountNo.BackStyle = acBackStyleTransparent
2540        .miscAccountNo_lbl.BackStyle = acBackStyleTransparent
2550        .miscAccountNo_lbl_box.Visible = True
2560        .cmdLock.Enabled = True
2570        .cmdLock_open_raised_img_dis.Visible = False
2580        .cmdLock_open_raised_img.Visible = False
2590        .cmdLock_closed_raised_img.Visible = True
2600        .cmbAccountHelper.Enabled = False
2610        .cmbAccountHelper.BorderColor = WIN_CLR_DISR
2620        .cmbAccountHelper.BackStyle = acBackStyleTransparent
2630        .cmdMiscMap.Enabled = False
2640        If gblnGoToReport = False Then
2650          If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
2660            .cmdMiscMap_LTCL.Enabled = False
2670            .cmdMiscMap_LTCL_raised_img_dis.Visible = True
2680            .cmdMiscMap_LTCL_raised_img.Visible = False
2690            .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
2700            .cmdMiscMap_LTCL_raised_focus_img.Visible = False
2710            .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
2720            .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
2730            .cmdMiscMap_STCGL.Enabled = False
2740            .cmdMiscMap_STCGL_raised_img_dis.Visible = True
2750            .cmdMiscMap_STCGL_raised_img.Visible = False
2760            .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = False
2770            .cmdMiscMap_STCGL_raised_focus_img.Visible = False
2780            .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = False
2790            .cmdMiscMap_STCGL_sunken_focus_dots_img.Visible = False
2800          End If  ' ** blnSpecialCap.
2810        End If
2820      End Select

          ' ** Make sure validation is reset.
2830      gblnMiscValidated = False

2840      Select Case IsNull(.miscType)
          Case True
2850        blnTypeIsNull = True
2860      Case False
2870        blnTypeIsNull = False
2880      End Select

          ' ** Default.
2890      .cmbRecurringItems.Visible = True
2900      .cmbRecurringItems_lbl.Visible = True
2910      .miscAssetNo.Visible = False

2920      lngCurrID = .miscCurr_ID.Column(CBX_C_CURRID)

2930      gstrMiscType = IIf(IsNull(.miscType), vbNullString, .miscType)
2940      gstrMiscAccountNumber = IIf(IsNull(.miscAccountNo), vbNullString, .miscAccountNo)
2950      gstrMiscICash = IIf(IsNull(.miscICash), vbNullString, .miscICash)
2960      gstrMiscICash = Rem_Dollar(gstrMiscICash, lngCurrID)  ' ** Module Function: modStringFuncs.
2970      gstrMiscPCash = IIf(IsNull(.miscPCash), vbNullString, .miscPCash)
2980      gstrMiscPCash = Rem_Dollar(gstrMiscPCash, lngCurrID)  ' ** Module Function: modStringFuncs.

2990      Select Case .NewRecord
          Case True
3000        .cmdMiscCancel.Enabled = False
3010      Case False
3020        .cmdMiscCancel.Enabled = True
3030      End Select

          ' ** I believe this is how it should work:
          ' ** On a new record, curr_id should be white, letting the user
          ' **   know choosing an appropriate currency is available.
          ' ** On an existing record, always green, indicating
          ' **   it's set, and discouraging change.
3040      If .miscType = "Paid" Then
3050        .cmbRecurringItems.Enabled = True
3060        .cmbRecurringItems.BorderColor = CLR_LTBLU2
3070        .cmbRecurringItems.BackStyle = acBackStyleNormal
3080        .cmbRecurringItems_lbl.BackStyle = acBackStyleNormal
3090        .cmbRecurringItems_lbl_box.Visible = False
3100        .cmbRecurringItems.RowSource = "qryRecurringItems_01b"  ' ** Payee.
3110        .cmdMiscRecurNew.Caption = "&New Payee"
3120        .cmdMiscRecurNew.ControlTipText = "New Payee"
3130        .cmdMiscRecurNew.Enabled = True
3140        .chkChecks.Enabled = True
3150        .miscCurr_ID.Enabled = True
3160        .miscCurr_ID.Locked = False
3170        .miscCurr_ID_cmd.Visible = False
3180        .miscCurr_ID_cmd.Enabled = False
3190        .miscCurr_ID.ForeColor = CLR_DISABLED_FG
3200        .miscCurr_ID.BackColor = CLR_VLTGRN
3210        .cmbRecurringItems.Requery
3220      Else
            ' ** 'Paid' is gone, handled above.
3230        If .miscType = "Received" Then
3240          If .miscAssetNo <> vbNullString And .miscAssetNo <> 0 Then
3250            .miscAssetNo.Visible = True
3260            .cmbRecurringItems.Visible = False
3270            .cmbRecurringItems_lbl.Visible = False
3280            .cmbRecurringItems_lbl_box.Visible = False
3290            .chkChecks.Enabled = False
3300            .miscCurr_ID.Enabled = False
3310            .miscCurr_ID.Locked = True
3320            .miscCurr_ID_cmd.Visible = True
3330            .miscCurr_ID_cmd.Enabled = True
3340            .miscCurr_ID.ForeColor = CLR_DISABLED_FG
3350            .miscCurr_ID.BackColor = CLR_VLTGRN  ' ** With an assetno, this is coming from LTCG, so it could be foreign!
3360            Select Case IsNull(.miscAssetNo.Column(4))
                Case True
3370              lngCurrID = 150&  ' ** USD.
3380            Case False
3390              lngCurrID = .miscAssetNo.Column(4)
3400            End Select
3410            Select Case IsNull(.miscCurr_ID)
                Case True
3420              .miscCurr_ID = lngCurrID
3430            Case False
                  ' ** Let it stand.
3440            End Select
3450          Else
3460            .cmbRecurringItems.Enabled = True
3470            .cmbRecurringItems.BorderColor = CLR_LTBLU2
3480            .cmbRecurringItems.BackStyle = acBackStyleNormal
3490            .cmbRecurringItems_lbl.BackStyle = acBackStyleNormal
3500            .cmbRecurringItems_lbl_box.Visible = False
3510            .cmbRecurringItems.RowSource = "qryRecurringItems_01c"  ' ** Payor.
3520            .cmdMiscRecurNew.Caption = "&New Payor"
3530            .cmdMiscRecurNew.ControlTipText = "New Payor"
3540            .cmdMiscRecurNew.Enabled = True
3550            .chkChecks.Enabled = False
3560            .miscCurr_ID.Enabled = True
3570            .miscCurr_ID.Locked = False
3580            .miscCurr_ID_cmd.Visible = False
3590            .miscCurr_ID_cmd.Enabled = False
3600            .miscCurr_ID.ForeColor = CLR_DISABLED_FG
3610            .miscCurr_ID.BackColor = CLR_VLTGRN  ' ** A standard 'Received' entry. Why white?
3620            .cmbRecurringItems.Requery
3630          End If
3640        Else
              ' ** 'Paid' and 'Received' are gone, handled above.
3650          If .miscType = "Misc." Then
3660            .cmbRecurringItems.Enabled = True
3670            .cmbRecurringItems.BorderColor = CLR_LTBLU2
3680            .cmbRecurringItems.BackStyle = acBackStyleNormal
3690            .cmbRecurringItems_lbl.BackStyle = acBackStyleNormal
3700            .cmbRecurringItems_lbl_box.Visible = False
3710            .cmbRecurringItems.RowSource = "qryRecurringItems_01a"  ' ** Misc.
3720            .cmdMiscRecurNew.Caption = "&New Item"
3730            .cmdMiscRecurNew.ControlTipText = "New Item"
3740            .cmdMiscRecurNew.Enabled = True
3750            .chkChecks.Enabled = False
3760            .miscCurr_ID.Enabled = True
3770            .miscCurr_ID.Locked = False
3780            .miscCurr_ID_cmd.Visible = False
3790            .miscCurr_ID_cmd.Enabled = False
3800            .miscCurr_ID.ForeColor = CLR_DISABLED_FG
3810            .miscCurr_ID.BackColor = CLR_VLTGRN  ' ** Why white?
3820            .cmbRecurringItems.Requery
3830            DoEvents
3840            If .cmbRecurringItems.Column(CBX_R_RID) = RECUR_I_TO_P_ID Then
3850              .MiscXFer_ItoP_lbl.Visible = True
3860              .MiscXFer_PtoI_lbl.Visible = False
3870            ElseIf .cmbRecurringItems.Column(CBX_R_RID) = RECUR_P_TO_I_ID Then
3880              .MiscXFer_ItoP_lbl.Visible = False
3890              .MiscXFer_PtoI_lbl.Visible = True
3900            End If
3910          Else
                ' ** 'Paid', 'Received', and 'Misc.' have all been handled above.
                ' ** There is nothing left, so this must be a new record.
3920            .cmbRecurringItems.Enabled = False
3930            .cmbRecurringItems.BorderColor = WIN_CLR_DISR
3940            .cmbRecurringItems.BackStyle = acBackStyleTransparent
3950            .cmbRecurringItems_lbl.BackStyle = acBackStyleTransparent
3960            .cmbRecurringItems_lbl_box.Visible = True
3970            .cmdMiscRecurNew.Enabled = False
3980            .cmdMiscRecurNew.Caption = "&New Item"
3990            .cmdMiscRecurNew.ControlTipText = "New Item"
4000            .chkChecks.Enabled = False
4010            .miscCurr_ID.Enabled = False
4020            .miscCurr_ID.Locked = True
4030            .miscCurr_ID_cmd.Visible = True
4040            .miscCurr_ID_cmd.Enabled = True
4050            .miscCurr_ID.ForeColor = CLR_BLK
4060            .miscCurr_ID.BackColor = CLR_WHT
4070          End If
4080        End If
4090      End If

4100      strMemo_Orig = vbNullString: strMemo_New = vbNullString
4110      Select Case .chkChecks
          Case True
4120        .chkChecks_lbl.FontBold = True
4130        .miscMemo = Null
4140        .miscMemo.Enabled = True
4150        .miscMemo.Visible = True
4160        .miscMemo_lbl2.Visible = True
4170        varTmp00 = DLookup("[JrnlMemo_Memo]", "tblJournal_Memo", "[Journal_ID] = " & CStr(.miscID) & " And " & _
              "[accountno] = '" & .miscAccountNo_Data & "'")
4180        If IsNull(varTmp00) = False Then
4190          If Trim(varTmp00) <> vbNullString Then
4200            strMemo_Orig = Trim(varTmp00)
4210            .miscMemo = Left(strMemo_Orig, MEMO_MAX)
4220          End If
4230        End If
4240      Case False
4250        .chkChecks_lbl.FontBold = False
4260        .miscMemo.Visible = False
4270        .miscMemo.Enabled = False
4280        .miscMemo_lbl2.Visible = False
4290        .miscMemo = Null
4300      End Select

4310      If blnCmdButton = False Then  ' ** That is, only run this if Form_Current() wasn't called from another button, like cmdMiscOK_Click().
4320        blnGoToSaleForm = False
4330        .tglMiscSale_false_raised_img_dis.Visible = True
4340        .tglMiscSale_false_raised_img.Visible = False
4350        .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
4360        .tglMiscSale_false_raised_focus_img.Visible = False
4370        .tglMiscSale_false_raised_focus_dots_img.Visible = False
4380        .tglMiscSale_false_sunken_focus_dots_img.Visible = False
4390        .tglMiscSale_true_raised_img.Visible = False
4400        .tglMiscSale_true_raised_focus_img.Visible = False
4410        .tglMiscSale_true_raised_focus_dots_img.Visible = False
4420        .tglMiscSale_true_sunken_focus_dots_img.Visible = False
4430        .tglMiscSale_true_raised_img_dis.Visible = False
4440        .tglMiscSale.Enabled = False
4450        If blnTypeIsNull = False Then
4460          If .miscType = "Paid" Then
4470            Select Case .posted
                Case True
4480              .tglMiscSale.Enabled = False
4490              .tglMiscSale_true_raised_img_dis.Visible = True
4500              .tglMiscSale_false_raised_img_dis.Visible = False
4510            Case False
4520              .tglMiscSale.Enabled = True
4530              .tglMiscSale_false_raised_img.Visible = True
4540              .tglMiscSale_false_raised_img_dis.Visible = False
4550            End Select
4560          End If
4570        End If
4580        blnGoToRecReinvest = False
4590        .tglMiscReinvest_false_raised_img_dis.Visible = True
4600        .tglMiscReinvest_false_raised_img.Visible = False
4610        .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
4620        .tglMiscReinvest_false_raised_focus_img.Visible = False
4630        .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
4640        .tglMiscReinvest_false_sunken_focus_dots_img.Visible = False
4650        .tglMiscReinvest_true_raised_img.Visible = False
4660        .tglMiscReinvest_true_raised_focus_img.Visible = False
4670        .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
4680        .tglMiscReinvest_true_sunken_focus_dots_img.Visible = False
4690        .tglMiscReinvest_true_raised_img_dis.Visible = False
4700        .tglMiscReinvest.Enabled = False
4710        If blnTypeIsNull = False Then
4720          If .miscType = "Received" Then
4730            Select Case .posted
                Case True
4740              .tglMiscReinvest.Enabled = False
4750              .tglMiscReinvest_true_raised_img_dis.Visible = True
4760              .tglMiscReinvest_false_raised_img_dis.Visible = False
4770            Case False
4780              .tglMiscReinvest.Enabled = True
4790              .tglMiscReinvest_false_raised_img.Visible = True
4800              .tglMiscReinvest_false_raised_img_dis.Visible = False
4810            End Select
4820          End If
4830        End If
4840      End If  ' ** blnCmdButton.

4850      .Repaint

4860      DoEvents

          ' ** Set the currency symbol.
4870      miscCurr_ID_AfterUpdate  ' ** Procedure: Below.

4880      If .miscCurr_ID.Visible = True Then
4890        miscICash_AfterUpdate  ' ** Procedure: Below.
4900        miscPCash_AfterUpdate  ' ** Procedure: Below.
4910      End If

4920      RevCodeMisc_Sub5 Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
4930      TaxCodeMisc_Sub5 Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

4940    End With

EXITP:
4950    Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
490     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5020    intShiftDown = (Shift And acShiftMask) > 0
5030    intAltDown = (Shift And acAltMask) > 0
5040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Next Record:      N {MoveRec}
        ' **   Previous Record:  P {MoveRec}
        ' **   Switch:           T {cmdSwitch on frmJournal}

        ' ** Shortcut Ctrl-Shift to other forms and functionality:
        ' **   Mark Reinvest:    R {tglMiscReinvest}
        ' **   Mark Sale:        S {tglMiscSale}
        ' **   Close Timer:      T {Form_KeyDown}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}

        ' ** Plain keys.
5050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5060      Select Case intRetVal
          Case vbKeyEscape
5070        intRetVal = 0
5080        If Forms("frmJournal").opgJournal.Enabled = True Then
5090          Select Case gblnMiscChanged
              Case True
5100            cmdMiscCancel_Click  ' ** Procedure: Below.
5110          Case False
5120            cmdMiscClose_Click  ' ** Procedure: Below.
5130          End Select
5140        Else
5150          cmdMiscCancel_Click  ' ** Procedure: Below.
5160        End If
5170      Case vbKeyDelete
5180        gblnMiscChanged = True
5190      Case vbKeyF7
5200        intRetVal = 0
5210        cmdCalendar1_Click  ' ** Procedure: Below.
5220      End Select
5230    End If

        ' ** Alt keys.
5240    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
5250      Select Case intRetVal
          Case vbKeyC
5260        With Me
5270          intRetVal = 0
5280          If .cmdMiscClose.Enabled = True Then
5290            .cmdMiscClose.SetFocus
5300            cmdMiscClose_Click  ' ** Procedure: Below.
5310          Else
5320            Beep
5330          End If
5340        End With
5350      Case vbKeyE
5360        With Me
5370          intRetVal = 0
5380          If .cmdMiscCancel.Enabled = True Then
5390            .cmdMiscCancel.SetFocus
5400            cmdMiscCancel_Click  ' ** Procedure: Below.
5410          Else
5420            Beep
5430          End If
5440        End With
5450      Case vbKeyX
5460        With Me
5470          intRetVal = 0
5480          If .cmdMiscClose.Enabled = True Then
5490            cmdMiscClose_Click  ' ** Procedure: Below.
5500          End If
5510        End With
5520      End Select
5530    End If

        ' ** Ctrl keys.
5540    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
5550      Select Case intRetVal
          Case vbKeyN
5560        With Me
5570          intRetVal = 0
5580          MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5590        End With
5600      Case vbKeyP
5610        With Me
5620          intRetVal = 0
5630          MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5640        End With
5650      Case vbKeyT
5660        With Me
5670          intRetVal = 0
5680          DoCmd.SelectObject acForm, .Parent.Name, False
5690          .Parent.cmdSwitch.SetFocus
5700          DoEvents
5710          .Parent.cmdSwitch_Click  ' ** Form Procedure: frmJournal.
5720        End With
5730      End Select
5740    End If

        ' ** Ctrl-Shift keys.
5750    If intCtrlDown And (Not intAltDown) And intShiftDown Then
5760      Select Case intRetVal
          Case vbKeyF
5770        With Me
5780          intRetVal = 0
5790          .FocusHolder.SetFocus
5800        End With
5810      Case vbKeyR
5820        With Me
5830          intRetVal = 0
5840          .tglMiscReinvest.SetFocus
5850          tglMiscReinvest_Click  ' ** Procedure: Below.
5860        End With
5870      Case vbKeyS
5880        With Me
5890          intRetVal = 0
5900          .tglMiscSale.SetFocus
5910          tglMiscSale_Click  ' ** Procedure: Below.
5920        End With
5930      Case vbKeyT
5940        intRetVal = 0
5950        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
5960      End Select
5970    End If

EXITP:
5980    KeyCode = intRetVal
5990    Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
500     Resume EXITP

End Sub

Private Sub Form_Timer()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

6010    With Me
6020      .TimerInterval = 0&
          ' ** Trying to get it to go to a new record after a map.
6030      If blnMapNewRec = True Then
6040        blnMapNewRec = False
6050        If IsNull(.miscID) = False Then
6060          MoveRec acCmdRecordsGoToNew  ' ** Procedure: Below.
6070        End If
6080      End If
6090    End With

EXITP:
6100    Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub Form_LostFocus()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_LostFocus"

6210    VerifyAccountNo  ' ** Function: Below.

EXITP:
6220    Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim blnMsgGiven As Boolean
        Dim strControl As String

6310    blnMsgGiven = False

6320    If DataErr <> 0 Then
6330      Select Case DataErr
          Case 13, 2108, 2110, 2116, 2135, 2169, 2474, 2501, 3020, 3163, 3314, 3315, 7753, 8519
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
6340        Response = acDataErrContinue
6350      Case 2046  ' ** The command or action isn't available now (first or last record).
            ' ** Do nothing.
6360        Response = acDataErrContinue
6370      Case 2113  ' ** The value you entered isn't valid for this field.
6380  On Error Resume Next
6390        strControl = Screen.ActiveControl.Name
6400  On Error GoTo ERRH
6410        Select Case strControl
            Case "miscPCash"
6420          strControl = "Principal Cash"
6430        Case "miscICash"
6440          strControl = "Income Cash"
6450        Case "miscCost"
6460          strControl = "Cost"
6470        Case "miscTransDate"
6480          strControl = "Date"
6490          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
6500          blnMsgGiven = True
6510        Case Else
6520          strControl = vbNullString
6530        End Select
6540        If blnMsgGiven = False Then
6550          MsgBox "You must enter a numeric value in the " & strControl & " field.", vbInformation + vbOKOnly, "Entry Required"
6560        End If
6570        Response = acDataErrContinue
6580        DoCmd.CancelEvent
6590      Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
            ' ** Ignore.
6600        Response = acDataErrContinue
6610      Case 2237  ' ** The text you entered isn't an item in the list.
6620        Response = acDataErrContinue
6630        DoCmd.CancelEvent
6640      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
6650        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
6660        Response = acDataErrContinue
6670        DoCmd.CancelEvent
6680      Case 3101  ' ** The Microsoft Jet database engine cannot find a record in the table '|' with key matching field(s) '|'.
6690        If IsNull(Me.miscAccountNo) Or Me.miscAccountNo = vbNullString Then
6700          MsgBox "The posting of this transaction is incomplete." & vbCrLf & vbCrLf & _
                "You must either complete or cancel this transaction to continue.", vbInformation + vbOKOnly, "Incomplete Data"
6710          If Me.cmdMiscClose.Enabled = True And Me.cmdMiscCancel.Enabled = False Then
6720            Me.cmdMiscOK.Enabled = True
6730            Me.cmdMiscCancel.Enabled = True
6740            Me.cmdMiscClose.Enabled = False
6750          End If
6760        End If
6770  On Error Resume Next
6780        strControl = Screen.ActiveControl.Name
6790  On Error GoTo ERRH
6800        If strControl <> "miscAccountno" And blnAccountNoHasFocus = False Then
6810          If Me.miscAccountNo.Enabled = True Then
6820            blnAccountNoHasFocus = True
6830  On Error Resume Next
6840            blnFromFormError = True
6850            Me.miscAccountNo.SetFocus
6860  On Error GoTo ERRH
6870          Else
6880            Me.miscTransDate.SetFocus
6890          End If
6900        End If
6910        Response = acDataErrContinue
6920      Case 3162  ' ** You tried to assign the Null value to a variable that is not a Variant data type.
6930  On Error Resume Next
6940        strControl = Screen.ActiveControl.Name
6950  On Error GoTo ERRH
6960        MsgBox "The " & strControl & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
6970        Response = acDataErrContinue
6980      Case Else
6990        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
7000        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
7010      End Select
7020      gblnMiscChanged = True
7030    End If

EXITP:
7040    Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.
' ** VGC 04/13/08: Calendar window seems to be Modal, preventing the
' **               user from doing anything else till it's closed.

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

7110    If Not clsMonthClass Is Nothing Then
7120      If clsMonthClass.IsCalendar Then
7130        Cancel = -1
7140      Else
7150        Set clsMonthClass = Nothing
7160      End If
7170    End If

EXITP:
7180    Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

7210    DetailMouse_Sub5 blnCalendar1_Focus, blnMiscMap_LTCL_Focus, blnMiscMap_STCGL_Focus, _
          blnMiscSale_Focus, blnMiscReinvest_Focus, blnSpecialCap, intSpecialCapOpt, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
7220    Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Public Sub cmdMiscClose_Click()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscClose_Click"

        Dim strDocName As String

7310    glngJournalForm = 0&
7320    gstrFormQuerySpec = vbNullString
7330    gblnSetFocus = True
7340    strDocName = "frmJournal"
7350    DoCmd.Close acForm, strDocName

EXITP:
7360    Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub cmdMiscClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscClose_MouseMove"

7410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7420    Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdMiscOK_Click()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngThisJrnlID As Long, lngJrnlMS_ID As Long, lngThisCurrID As Long
        Dim strShortName As String, dblICash As Double, dblPCash As Double, datTransDate As Date
        Dim blnSaveVisible_cmbRecurringItems As Boolean, blnSaveEnabled_miscAssetNo As Boolean
        Dim blnTmpGoToSaleForm As Boolean, blnTmpGoToRecReinvest As Boolean
        Dim strRevCode As String, lngTaxcode As Long, strDocName As String
        Dim varTmp00 As Variant, curTmp01 As Currency, strTmp02 As String, strTmp03 As String
        Dim blnContinue As Boolean

7510    With Me

7520      blnContinue = True

7530      DoCmd.Hourglass True
7540      DoEvents

7550      blnCmdButton = True
7560      blnTmpGoToSaleForm = blnGoToSaleForm
7570      blnTmpGoToRecReinvest = blnGoToRecReinvest
7580      strTmp02 = strMemo_Orig
7590      strTmp03 = strMemo_New
7600      Form_Current  ' ** Procedure: Above.
          ' ** Even if I reassign the two memo strings here, they get nulled out again somewhere below!
7610      blnCmdButton = False
7620      VerifyAccountNo  ' ** Function: Below.

7630      If ((gstrMiscPCash = vbNullString Or Val(gstrMiscPCash) = 0) And _
              (gstrMiscICash = vbNullString Or Val(gstrMiscICash) = 0)) Then
7640        blnContinue = False
7650        DoCmd.Hourglass False
7660        MsgBox "Must enter an amount in the Income Cash or Principal Cash field.", vbInformation + vbOKOnly, "Entry Required"
7670        .miscICash.SetFocus
7680      Else
7690        If .miscAccountNo = vbNullString Or IsNull(.miscAccountNo) Then
7700          blnContinue = False
7710          DoCmd.Hourglass False
7720          MsgBox "Please enter a valid account number.", vbInformation + vbOKOnly, "Invalid Entry"
7730          .miscAccountNo.SetFocus
7740        Else
7750          blnSaveVisible_cmbRecurringItems = .cmbRecurringItems.Visible
7760          blnSaveEnabled_miscAssetNo = .miscAssetNo.Enabled
              ' ** 07/21/2008: Make sure there's something there.
7770          If (blnSaveVisible_cmbRecurringItems = True And (.cmbRecurringItems = vbNullString Or IsNull(.cmbRecurringItems) = True)) Or _
                  (blnSaveEnabled_miscAssetNo = True And (.miscAssetNo = vbNullString Or IsNull(.miscAssetNo) = True)) Then
7780            blnContinue = False
7790            DoCmd.Hourglass False
7800            MsgBox "Please enter a Description/Recurring.", vbInformation + vbOKOnly, "Entry Required"
7810            .cmbRecurringItems.SetFocus
7820          Else
7830            If .miscType = vbNullString Or IsNull(.miscType) Then
7840              blnContinue = False
7850              DoCmd.Hourglass False
7860              MsgBox "Please enter a valid type 'Misc.', 'Paid', or 'Received'. ", vbInformation + vbOKOnly, "Invalid Entry"
7870              .miscType.SetFocus
7880            Else
                  ' ** For some unknown reason, strMemo_New gets Nulled out between here...
7890              If .cmbRecurringItems.Visible = False Then
7900                .cmbRecurringItems.Visible = True
7910                .miscAssetNo.Visible = False
7920              End If
7930              If .cmbRecurringItems.Enabled = False Then
7940                .cmbRecurringItems.Enabled = True
7950                .cmbRecurringItems.BorderColor = CLR_LTBLU2
7960                .cmbRecurringItems.BackStyle = acBackStyleNormal
7970                .cmbRecurringItems_lbl.BackStyle = acBackStyleNormal
7980                .cmbRecurringItems_lbl_box.Visible = False
7990              End If
8000              .cmbRecurringItems.SetFocus
                  ' ** And here...
8010              If InStr(.cmbRecurringItems.text, Chr(34)) > 0 Then
8020                blnContinue = False
8030                DoCmd.Hourglass False
8040                MsgBox "Cannot have quote marks in text.", vbInformation + vbOKOnly, "Invalid Characters"
8050                .cmbRecurringItems.SetFocus
8060              Else
8070                If (((.cmbRecurringItems.Column(CBX_R_RID) = RECUR_I_TO_P_ID) Or _
                        (.cmbRecurringItems.Column(CBX_R_RID) = RECUR_P_TO_I_ID)) And _
                        (((.miscICash <> 0) And (.miscPCash = 0)) Or ((.miscICash = 0) And (.miscPCash <> 0)))) Then
8080                  blnContinue = False
8090                  DoCmd.Hourglass False
8100                  MsgBox "Cash transfers must be of equal amounts.", vbInformation + vbOKOnly, "Invalid Entry"
8110                  If .miscICash = 0 Then
8120                    .miscICash.SetFocus
8130                  ElseIf .miscPCash = 0 Then
8140                    .miscPCash.SetFocus
8150                  End If
8160                Else
8170                  .cmdMiscOK.SetFocus
8180                  If blnSaveVisible_cmbRecurringItems = False Then
8190                    .cmbRecurringItems.Visible = False
8200                    .cmbRecurringItems_lbl_box.Visible = False
8210                    .miscAssetNo.Visible = True
8220                  End If
8230                  If Not blnSaveEnabled_miscAssetNo Then
8240                    .cmbRecurringItems.Enabled = False
8250                    .cmbRecurringItems.BorderColor = WIN_CLR_DISR
8260                    .cmbRecurringItems.BackStyle = acBackStyleTransparent
8270                    .cmbRecurringItems_lbl.BackStyle = acBackStyleTransparent
8280                    .cmbRecurringItems_lbl_box.Visible = True
8290                  End If
8300                  .miscDescription.SetFocus
8310                  If InStr(.miscDescription.text, Chr(34)) > 0 Then
8320                    blnContinue = False
8330                    DoCmd.Hourglass False
8340                    MsgBox "Cannot have quote marks in text.", vbInformation + vbOKOnly, "Invalid Characters"
8350                    .miscDescription.SetFocus
8360                  Else
8370                    If IsNull(.miscTransDate) = True Then
8380                      blnContinue = False
8390                      DoCmd.Hourglass False
8400                      MsgBox "A Posting Date is required.", vbInformation + vbOKOnly, "Entry Required"
8410                      .miscTransDate.SetFocus
8420                    Else
8430                      If .miscTransDate > DateAdd("m", 1, Date) Then
8440                        blnContinue = False
8450                        DoCmd.Hourglass False
8460                        MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
8470                        .miscTransDate.SetFocus
8480                      End If
8490                    End If
8500                  End If
8510                  .Parent.frmJournal_Sub4_Sold.Form.RecalcPaidTotal  ' ** Form Procedure: frmJournal_Sub4_Sold.
8520                End If
8530              End If
8540            End If
8550          End If
8560        End If
8570      End If

8580      If blnContinue = True Then

8590        DoCmd.Hourglass True  ' ** Make sure it's still running.
8600        DoEvents

8610        If .miscType = "Misc." Then
8620          If IsNull(.cmbRecurringItems) = False Then
8630            If .cmbRecurringItems <> vbNullString Then
8640              If .cmbRecurringItems.Column(CBX_R_RID) = RECUR_I_TO_P_ID Then
8650                If IsNull(.miscICash) = False And IsNull(.miscPCash) = False Then
8660                  If .miscICash <> 0 And .miscPCash <> 0 Then
8670                    If Abs(.miscICash) <> Abs(.miscPCash) Then
8680                      blnContinue = False
8690                      DoCmd.Hourglass False
8700                      MsgBox "Cash transfers must be of equal amounts.", vbInformation + vbOKOnly, "Invalid Entry"
8710                      .miscPCash.SetFocus
8720                    End If
8730                    If .miscICash > 0 Then
8740                      .miscICash = .miscICash * -1
8750                    End If
8760                    If .miscPCash < 0 Then
8770                      .miscPCash = .miscPCash * -1
8780                    End If
8790                  End If
8800                End If
8810              End If
8820              If .cmbRecurringItems.Column(CBX_R_RID) = RECUR_P_TO_I_ID Then
8830                If IsNull(.miscPCash) = False And IsNull(.miscICash) = False Then
8840                  If .miscICash <> 0 And .miscPCash <> 0 Then
8850                    If Abs(.miscICash) <> Abs(.miscPCash) Then
8860                      blnContinue = False
8870                      DoCmd.Hourglass False
8880                      MsgBox "Cash transfers must be of equal amounts.", vbInformation + vbOKOnly, "Invalid Entry"
8890                      .miscPCash.SetFocus
8900                    End If
8910                    If .miscPCash > 0 Then
8920                      .miscPCash = .miscPCash * -1
8930                    End If
8940                    If .miscICash < 0 Then
8950                      .miscICash = .miscICash * -1
8960                    End If
8970                  End If
8980                End If
8990              End If
9000            End If
9010          End If
9020        End If

9030        DoEvents
9040      End If

9050      If blnContinue = True Then

9060        DoCmd.Hourglass True  ' ** Make sure it's still running.
9070        DoEvents

9080        If Nz(.miscUser, vbNullString) = vbNullString Then
9090          .miscUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
9100        End If

9110        strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
9120        lngTaxcode = Nz(.cmbTaxCodes, 0&)
9130        Select Case .miscType
            Case "Paid"
              ' ** EXPENSE.
9140          If strRevCode <> "E" Then
9150            .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
9160          End If
9170          If IsNull(.cmbTaxCodes) = True Then
9180            .cmbTaxCodes = 0&
9190          Else
9200            If gblnLinkRevTaxCodes = True Then
9210              If .cmbTaxCodes = 0& Then
9220                .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
9230              Else
9240                If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
9250                  .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
9260                End If
9270              End If
9280            End If
9290          End If
9300        Case "Received"
              ' ** INCOME.
9310          If strRevCode <> "I" Then
9320            .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
9330          End If
9340          If IsNull(.cmbTaxCodes) = True Then
9350            .cmbTaxCodes = 0&
9360          Else
9370            If gblnLinkRevTaxCodes = True Then
9380              If .cmbTaxCodes = 0& Then
9390                .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
9400              Else
9410                If .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
9420                  .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
9430                End If
9440              End If
9450            End If
9460          End If
9470        Case "Misc."  ' ** No cash consideration, but Tax Code takes precedence over Revenue Code!
              ' ** ALL.
9480          If IsNull(.cmbTaxCodes) = True Then
9490            .cmbTaxCodes = 0&
9500          End If
9510          If gblnLinkRevTaxCodes = True Then
9520            If .cmbTaxCodes = 0& Then
9530              If strRevCode = "I" Then
9540                .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
9550              ElseIf strRevCode = "E" Then
9560                .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
9570              End If
9580            Else
9590              If strRevCode = vbNullString Then
9600                If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
9610                  .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
9620                ElseIf .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
9630                  .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
9640                End If
9650              Else
9660                If strRevCode = "I" And .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
9670                  .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
9680                ElseIf strRevCode = "E" And .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
9690                  .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
9700                End If
9710              End If
9720            End If
9730          End If
9740          If strRevCode = vbNullString Then
9750            .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
9760          End If
9770        End Select

9780        If blnTmpGoToRecReinvest = True Then
9790          .miscReinvested = True
9800        End If

9810        DoCmd.RunCommand acCmdSaveRecord

9820        DoCmd.Hourglass True  ' ** Make sure it's still running.
9830        DoEvents

            ' ** These get nulled out above (not including the Form_Current() at the top).
9840        strMemo_Orig = strTmp02
9850        strMemo_New = strTmp03

            ' ** The field behind tglMiscSale, tblMiscReinvest IS posted!
            'If .tglMiscSale = True Then
            '  .posted = True
            'Else
            '  .posted = False
            'End If

            ' ** Save these in case the Sale/Reinvest button was pushed.
            ' ** We lose them when we go to a new record.
9860        lngThisJrnlID = .miscID
9870        gstrAccountNo = .miscAccountNo.Value
9880        strShortName = .miscShortname.Value
9890        datTransDate = .miscTransDate.Value
9900        dblICash = Nz(.miscICash, 0)
9910        dblPCash = Nz(.miscPCash, 0)
9920        curTmp01 = CCur(Abs(Nz(.miscICash, 0) + Nz(.miscPCash, 0)))
9930        lngThisCurrID = .miscCurr_ID.Column(CBX_C_CURRID)

9940        If (.miscType = "Paid") And (strMemo_Orig <> strMemo_New) Then
9950          MemoUpdate_Sub5 lngThisJrnlID, gstrAccountNo, datTransDate, strMemo_New, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
9960        End If

9970        If .miscCurr_Date.Visible = True Then .miscCurr_Date.Visible = False

9980  On Error Resume Next
9990        DoCmd.GoToRecord , , acNewRec
10000 On Error GoTo ERRH

10010       .miscAccountNo.SetFocus

10020       gblnMiscValidated = False

10030       ChangedMisc_Sub5 False, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

10040       DoCmd.Hourglass True  ' ** Make sure it's still running.
10050       DoEvents

10060       blnGoToSaleForm = blnTmpGoToSaleForm

            ' ** Sale button was pushed.
10070       If blnGoToSaleForm = True Then

              ' ** Put preliminary record into tblJournal_MiscSold.
10080         Set dbs = CurrentDb
10090         Set rst = dbs.OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
10100         With rst
10110           If .BOF = True And .EOF = True Then
                  ' ** Proceed.
10120           Else
10130             .MoveFirst
10140             .FindFirst "[Journal_ID] = " & CStr(lngThisJrnlID)
10150             If .NoMatch = False Then
10160               .Delete
10170             End If
10180           End If
10190           DoEvents
10200           .AddNew
                ' ** ![jrnlms_id] : AutoNumber.
10210           ![Journal_ID] = lngThisJrnlID
10220           ![journaltype] = "Paid"
10230           ![jrnlms_ref_id] = 0&
10240           ![jrnlms_parent] = True
10250           ![jrnlms_child] = False
10260           ![jrnlms_note] = "This Paid transaction was covered by generating a Sold transaction."
10270           ![jrnlms_datemodified] = Now()
10280           .Update
10290           .Bookmark = .LastModified
10300           lngJrnlMS_ID = ![jrnlms_id]
10310         End With
10320         Set rst = Nothing
10330         dbs.Close
10340         Set dbs = Nothing
10350         DoEvents

              ' ** Since this won't be using a Dialog, put these somewhere.
10360         .Parent.DefPaidJrnlID = lngThisJrnlID
10370         .Parent.DefPaidJrnlMSID = lngJrnlMS_ID

10380         gstrSaleAsset = vbNullString
10390         gstrSaleShareFace = vbNullString
10400         gstrSaleICash = vbNullString
10410         gstrSalePCash = vbNullString
10420         gstrSaleCost = vbNullString

10430         DoCmd.SelectObject acForm, .Parent.Name, False

              ' ** Retrieve a default asset if this account has one.
10440         varTmp00 = DLookup("[taxlot]", "account", "[accountno] = '" & gstrAccountNo & "'")
10450         If IsNull(varTmp00) = False Then
10460           If Val(varTmp00) > 0 Then
10470             .Parent.DefAssetNo = CLng(Val(varTmp00))
10480           Else
10490             .Parent.DefAssetNo = 0&
10500           End If
10510         Else
10520           .Parent.DefAssetNo = 0&
10530         End If

10540         .Parent.DefPaidTot = curTmp01
10550         .Parent.opgJournal = .Parent.opgJournal_optSale.OptionValue
10560         .Parent.SetFocus
10570         .Parent.opgJournal.SetFocus
10580         .Parent.ShowSale  ' ** Form Procedure: frmJournal.
10590         DoEvents

10600         .Parent.frmJournal_Sub4_Sold.SetFocus
              ' ** Make sure we're on a new record!
10610         .Parent.frmJournal_Sub4_Sold.Form.MoveRec acCmdRecordsGoToNew  ' ** Form Procedure: frmJournal_Sub4_Sold.
10620         DoEvents
10630         If IsNull(.Parent.frmJournal_Sub4_Sold.Form.saleID) = False Then
10640           .Parent.frmJournal_Sub4_Sold.Form.MoveRec acCmdRecordsGoToNew  ' ** Form Procedure: frmJournal_Sub4_Sold.
10650           DoEvents
10660         End If
10670         .Parent.frmJournal_Sub4_Sold.Form.saleAccountNo = gstrAccountNo
10680         .Parent.frmJournal_Sub4_Sold.Form.saleTransDate = datTransDate
10690         .Parent.frmJournal_Sub4_Sold.Form.saleType = "Sold"
10700         If .Parent.DefAssetNo > 0& Then
10710           .Parent.frmJournal_Sub4_Sold.Form.saleAssetno = .Parent.DefAssetNo
10720           gstrSaleAsset = CStr(.Parent.DefAssetNo)
10730           DoEvents
10740           .Parent.frmJournal_Sub4_Sold.Form.SaleChanged True  ' ** Form Procedure: frmJournal_Sub4_Sold.
10750           subTaxCode4 blnSpecialCap, intSpecialCapOpt, .Parent.frmJournal_Sub4_Sold.Form  ' ** Module Procedure: modPurchaseSold.
10760         End If
10770         If .Parent.DefPaidTot > 0@ Then
10780           .Parent.frmJournal_Sub4_Sold.Form.saleShareFace = CDbl(.Parent.DefPaidTot)
10790           .Parent.frmJournal_Sub4_Sold.Form.salePCash = .Parent.DefPaidTot
10800         End If
10810         .Parent.frmJournal_Sub4_Sold.Form.saleCost = 0
10820         .Parent.frmJournal_Sub4_Sold.Form.CheckNum = lngThisJrnlID  ' ** Borrowing this field for Sale/posted reference.
10830         DoEvents
10840         .Parent.DefSaleBtn = True
10850         .Parent.frmJournal_Sub4_Sold.Form.cmdSaleOK.Enabled = True
10860         .Parent.frmJournal_Sub4_Sold.Form.cmdSaleClose.Enabled = False
10870         .Parent.frmJournal_Sub4_Sold.Form.cmdSaleLotInfo.Enabled = False
10880         .Parent.NavVis False  ' ** Form Procedure: frmJournal.
10890         gblnSaleChanged = True
10900         .Parent.frmJournal_Sub4_Sold.Form.saleAccountNo.SetFocus
10910         .Parent.frmJournal_Sub4_Sold.Form.saleAssetno.SetFocus

10920       End If
10930       blnGoToSaleForm = False  ' ** reset always.

10940       DoCmd.Hourglass True  ' ** Make sure it's still running.
10950       DoEvents

10960       blnGoToRecReinvest = blnTmpGoToRecReinvest

            ' ** Reinvest button was pushed.
10970       If blnGoToRecReinvest = True Then

              ' ** Put preliminary record into tblJournal_MiscSold.
10980         Set dbs = CurrentDb
10990         Set rst = dbs.OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
11000         With rst
11010           If .BOF = True And .EOF = True Then
                  ' ** Proceed.
11020           Else
11030             .MoveFirst
11040             .FindFirst "[Journal_ID] = " & CStr(lngThisJrnlID)
11050             If .NoMatch = False Then
11060               .Delete
11070             End If
11080           End If
11090           DoEvents
11100           .AddNew
                ' ** ![jrnlms_id] : AutoNumber.
11110           ![Journal_ID] = lngThisJrnlID
11120           ![journaltype] = "Received"
11130           ![jrnlms_ref_id] = 0&
11140           ![jrnlms_parent] = True
11150           ![jrnlms_child] = False
11160           ![jrnlms_note] = "This Received was reinvested, generating a Purchase transaction."
11170           ![jrnlms_datemodified] = Now()
11180           .Update
11190           .Bookmark = .LastModified
11200           lngJrnlMS_ID = ![jrnlms_id]
11210         End With
11220         Set rst = Nothing
11230         dbs.Close
11240         Set dbs = Nothing
11250         DoEvents

11260         gblnMessage = True  ' ** If this returns False, the dialog was canceled.
11270         gdblCrtRpt_CostTot = 0#  ' ** Borrowing this variable from the Court Reports.
11280         strDocName = "frmReinvest_Received"
11290         DoCmd.OpenForm strDocName, , , , acFormAdd, acDialog, strCallingForm & "~" & gstrAccountNo & "~" & _
                Format(datTransDate, "mm/dd/yyyy") & "~" & CStr(lngThisCurrID) & "~" & _
                CStr(dblICash) & "~" & CStr(dblPCash) & "~" & CStr(lngThisJrnlID) & "~" & strShortName

11300         DoCmd.Hourglass True
11310         DoEvents

11320         Select Case gblnMessage
              Case True
11330           Set dbs = CurrentDb
11340           With dbs
11350             Set rst = .OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
11360             With rst
11370               .MoveFirst
11380               .FindFirst "[jrnlms_id] = " & CStr(lngJrnlMS_ID)
11390               If .NoMatch = False Then
11400                 .Edit
11410                 ![jrnlms_ref_id] = CLng(gdblCrtRpt_CostTot)  ' ** New Purchase ID from frmReinvest_Dividend.
11420                 ![jrnlms_datemodified] = Now()
11430                 .Update
11440               End If
11450               .AddNew
                    ' ** ![jrnlms_id] : AutoNumber.
11460               ![Journal_ID] = CLng(gdblCrtRpt_CostTot)
11470               ![journaltype] = "Purchase"
11480               ![jrnlms_ref_id] = lngThisJrnlID
11490               ![jrnlms_parent] = False
11500               ![jrnlms_child] = True
11510               ![jrnlms_note] = "This Purchase was generated by a Received reinvestment."
11520               ![jrnlms_datemodified] = Now()
11530               .Update
11540               .Close
11550             End With
11560             Set rst = Nothing
11570             .Close
11580           End With
11590           Set dbs = Nothing
11600           DoEvents
11610           If .cmdMiscOK.Enabled = True Then
11620             .cmdMiscOK.SetFocus
11630           Else
11640             .miscAccountNo.SetFocus
11650           End If
11660         Case False
11670           Set dbs = CurrentDb
11680           With dbs
                  ' ** Delete tblJournal_MiscSold, by specified [jrnlmsid].
11690             Set qdf = dbs.QueryDefs("qryJournal_MiscSold_01")
11700             With qdf.Parameters
11710               ![jrnlmsid] = lngJrnlMS_ID
11720             End With
11730             qdf.Execute
11740             Set qdf = Nothing
11750             DoEvents
                  ' ** Update Journal, by specified [jrnlid], [pstd].
11760             Set qdf = .QueryDefs("qryJournal_Misc_11")
11770             With qdf.Parameters
11780               ![jrnlid] = lngThisJrnlID
11790               ![pstd] = False
11800             End With
11810             qdf.Execute
11820             Set qdf = Nothing
11830             If gdblCrtRpt_CostTot <> 0# Then
                    ' ** Delete Journal, by specified [jrnlid].
11840               Set qdf = .QueryDefs("qryJournal_Misc_10")
11850               With qdf.Parameters
11860                 ![jrnlid] = CLng(gdblCrtRpt_CostTot)
11870               End With
11880               qdf.Execute dbFailOnError
11890               Set qdf = Nothing
11900               gdblCrtRpt_CostTot = 0#
11910             End If
11920             .Close
11930           End With
11940           Set dbs = Nothing
11950           DoEvents
11960           .miscAccountNo.SetFocus
11970         End Select

11980       End If  ' ** blnGoToRecReinvest.
11990       blnGoToRecReinvest = False  ' ** reset always.

12000     End If  ' ** blnContinue.

12010     DoCmd.Hourglass False

12020   End With

EXITP:
12030   Set rst = Nothing
12040   Set qdf = Nothing
12050   Set dbs = Nothing
12060   Exit Sub

ERRH:
460     DoCmd.Hourglass False
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub cmdMiscCancel_Click()

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngThisJrnlID As Long, lngThatJrnlID As Long
        Dim blnReinvested As Boolean, strMsg As String
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

12110   With Me

12120     DoCmd.Hourglass True
12130     DoEvents

12140     If IsNull(.miscID) = True Then
12150       msgResponse = vbCancel
12160     Else

12170       blnReinvested = .miscReinvested
12180       lngThisJrnlID = .miscID
12190       lngThatJrnlID = 0&
12200       msgResponse = 0
            ' ** 2  vbCancel
            ' ** 6  vbYes
            ' ** 7  vbNo

            ' ** Check whether this spawned a Purchase using the Reinvest button or was covered using the Sale button.
12210       If blnReinvested = True And .posted = True And .miscType = "Received" Then  'Reinvest
              ' ** Check if the referencing entry still exists.
12220         varTmp00 = DCount("*", "journal", "[CheckNum] = " & CStr(lngThisJrnlID))
12230         If varTmp00 = 1 Then
12240           lngThatJrnlID = DLookup("[ID]", "journal", "[CheckNum] = " & CStr(lngThisJrnlID))
12250           strMsg = "This entry was used for a Purchase reinvestment." & vbCrLf & vbCrLf & _
                  "Yes - Delete both this Received and the Purchase reinvestment." & vbCrLf & _
                  "No - Delete only this Received." & vbCrLf & _
                  "Cancel - Delete no transactions."
12260           Beep
12270           DoCmd.Hourglass False
12280           msgResponse = MsgBox(strMsg, vbQuestion + vbYesNoCancel + vbDefaultButton3, "Delete Both Transactions")
12290           If msgResponse <> vbCancel Then
12300             DoCmd.Hourglass True
12310             If .miscCurr_Date.Visible = True Then .miscCurr_Date.Visible = False
12320             DoEvents
12330           End If
12340         Else
12350           lngThatJrnlID = 0&
12360         End If
12370       ElseIf blnReinvested = False And .posted = True And .miscType = "Paid" Then  'Sale
              ' ** Check if the referencing entry still exists.
12380         varTmp00 = DCount("*", "journal", "[CheckNum] = " & CStr(lngThisJrnlID))
12390         If varTmp00 = 1 Then
12400           lngThatJrnlID = DLookup("[ID]", "journal", "[CheckNum] = " & CStr(lngThisJrnlID))
12410           strMsg = "This entry was covered by a Sold transaction." & vbCrLf & vbCrLf & _
                  "Yes - Delete both this Paid and the covering Sold." & vbCrLf & _
                  "No - Delete only this Paid." & vbCrLf & _
                  "Cancel - Delete no transactions."
12420           Beep
12430           DoCmd.Hourglass False
12440           msgResponse = MsgBox(strMsg, vbQuestion + vbYesNoCancel + vbDefaultButton3, "Delete Both Transactions")
12450           If msgResponse <> vbCancel Then
12460             DoCmd.Hourglass True
12470             If .miscCurr_Date.Visible = True Then .miscCurr_Date.Visible = False
12480             DoEvents
12490           End If
12500         Else
12510           lngThatJrnlID = 0&
12520         End If
12530       End If

12540     End If

12550     If msgResponse = 0 Then
12560       DoCmd.Hourglass False
12570       msgResponse = MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo, "Delete Transaction")
12580       If msgResponse <> vbYes Then
12590         msgResponse = vbCancel
12600       Else
12610         msgResponse = vbNo  ' ** Delete only Misc.
12620         DoCmd.Hourglass True
12630         If .miscCurr_Date.Visible = True Then .miscCurr_Date.Visible = False
12640         DoEvents
12650       End If
12660     End If

12670     If msgResponse = vbCancel Then
12680       If .miscAccountNo.Enabled = True Then
12690         .miscAccountNo.SetFocus
12700       Else
12710         .miscTransDate.SetFocus
12720       End If
12730     Else

12740       If .posted = True And lngThatJrnlID > 0& Then
              ' ** Delete both, whether or not reinvestment/cover is deleted,
              ' ** since it would no longer have a parent.
12750         Set dbs = CurrentDb
12760         With dbs
                ' ** Delete tblJournal_MiscSold, by specified [jrnlid].
12770           Set qdf = .QueryDefs("qryJournal_MiscSold_02")
12780           With qdf.Parameters
12790             ![jrnlid] = lngThisJrnlID
12800           End With
12810           qdf.Execute
12820           Set qdf = Nothing
12830           DoEvents
                ' ** Delete tblJournal_MiscSold, by specified [jrnlid].
12840           Set qdf = .QueryDefs("qryJournal_MiscSold_02")
12850           With qdf.Parameters
12860             ![jrnlid] = lngThatJrnlID
12870           End With
12880           qdf.Execute
12890           Set qdf = Nothing
12900           DoEvents
                ' ** If Purchase or Sold not being deleted, clear out their CheckNum.
12910           If msgResponse = vbNo Then
                  ' ** Update Journal, for CheckNum = Null, by specified [jrnalid].
12920             Set qdf = .QueryDefs("qryJournal_Misc_12")
12930             With qdf.Parameters
12940               ![jrnalid] = lngThatJrnlID
12950             End With
12960             qdf.Execute
12970             Set qdf = Nothing
12980             DoEvents
12990           End If
13000           .Close
13010         End With
13020         Set dbs = Nothing
13030       End If
13040       DoEvents

13050       .Undo
13060       .Undo

13070       DoCmd.SetWarnings False
13080 On Error Resume Next
13090       DoCmd.RunCommand acCmdSelectRecord
13100       DoCmd.RunCommand acCmdDeleteRecord
13110 On Error GoTo ERRH
13120       DoCmd.SetWarnings True

13130       blnGoToSaleForm = False
13140       .tglMiscSale_false_raised_img_dis.Visible = True
13150       .tglMiscSale_false_raised_img.Visible = False
13160       .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
13170       .tglMiscSale_false_raised_focus_img.Visible = False
13180       .tglMiscSale_false_raised_focus_dots_img.Visible = False
13190       .tglMiscSale_false_sunken_focus_dots_img.Visible = False
13200       .tglMiscSale_true_raised_img.Visible = False
13210       .tglMiscSale_true_raised_focus_img.Visible = False
13220       .tglMiscSale_true_raised_focus_dots_img.Visible = False
13230       .tglMiscSale_true_sunken_focus_dots_img.Visible = False
13240       .tglMiscSale_true_raised_img_dis.Visible = False
13250       .tglMiscSale.Enabled = False
13260       blnGoToRecReinvest = False
13270       .tglMiscReinvest_false_raised_img_dis.Visible = True
13280       .tglMiscReinvest_false_raised_img.Visible = False
13290       .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
13300       .tglMiscReinvest_false_raised_focus_img.Visible = False
13310       .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
13320       .tglMiscReinvest_false_sunken_focus_dots_img.Visible = False
13330       .tglMiscReinvest_true_raised_img.Visible = False
13340       .tglMiscReinvest_true_raised_focus_img.Visible = False
13350       .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
13360       .tglMiscReinvest_true_sunken_focus_dots_img.Visible = False
13370       .tglMiscReinvest_true_raised_img_dis.Visible = False
13380       .tglMiscReinvest.Enabled = False

            ' ** Also delete reinvestment/cover.
13390       If msgResponse = vbYes And lngThatJrnlID > 0& Then
13400         Set dbs = CurrentDb
              ' ** Delete Journal, by specified [jrnlid].
13410         Set qdf = dbs.QueryDefs("qryJournal_Misc_10")
13420         With qdf.Parameters
13430           ![jrnlid] = lngThatJrnlID
13440         End With
13450         qdf.Execute
13460         Set qdf = Nothing
13470         dbs.Close
13480         Set dbs = Nothing
13490       End If

13500       .Parent.frmJournal_Sub3_Purchase.Form.Requery
13510       DoEvents
13520       .Parent.frmJournal_Sub4_Sold.Form.Requery
13530       DoEvents

13540       .cmbAccountHelper = Null
13550       If .miscAccountNo.Enabled = True Then
13560         .miscAccountNo.SetFocus
13570       Else
13580         .miscTransDate.SetFocus
13590       End If

13600       gblnMiscValidated = False

13610       ChangedMisc_Sub5 False, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

13620       Form_Current  ' ** Procedure: Above.

13630     End If

13640     DoCmd.Hourglass False

13650   End With

EXITP:
13660   Exit Sub

ERRH:
460     DoCmd.Hourglass False
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Public Sub cmdMiscMap_Click()
' ** Public for Go-To-Report.

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_Click"

        Dim strDocName As String
        Dim varTmp00 As Variant

13710   With Me

13720     DoCmd.Hourglass True
13730     DoEvents
13740     gblnSetFocus = True

13750     .Parent.LoadMapDropdowns  ' ** Form Procedure: frmJournal.
13760     DoEvents

13770     strDocName = "frmMap_Rec"
13780     DoCmd.OpenForm strDocName, , , , , , "frmJournal"

13790     If gblnGoToReport = True Then
13800       varTmp00 = DMax("[ID]", "journal")  ' ** Save this so we can delete any fake Journal records.
13810       Select Case IsNull(varTmp00)
            Case True
13820         glngTaxCode_Distribution = 0&  ' ** Borrowing this variable from the Court Reports.
13830       Case False
13840         glngTaxCode_Distribution = varTmp00
13850       End Select
13860       Forms(strDocName).TimerInterval = 100&
13870       .GoToReport_arw_maprec_img.Visible = False
13880       If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
13890         .cmdMiscMap_LTCL.Enabled = True
13900         .cmdMiscMap_LTCL_raised_img.Visible = True
13910         .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
13920         .cmdMiscMap_LTCL_raised_focus_img.Visible = False
13930         .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
13940         .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
13950         .cmdMiscMap_LTCL_raised_img_dis.Visible = False
13960       End If
13970       DoCmd.Hourglass True  ' ** Make sure it's still running.
13980       DoEvents
13990     End If

14000   End With

        'blnSpecialCap

EXITP:
14010   Exit Sub

ERRH:
460     DoCmd.Hourglass False
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub cmdMiscMap_KeyDown(KeyCode As Integer, Shift As Integer)
' ** See Note, above.

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_KeyDown"

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
          Case vbKeyTab
14170       With Me
14180         intRetVal = 0
14190         If .cmdMiscMap_LTCL.Enabled = True Then
14200           .cmdMiscMap_LTCL.SetFocus
14210         Else
14220           .miscTransDate.SetFocus
14230         End If
14240       End With
14250     End Select
14260   End If

        ' ** Shift keys.
14270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14280     Select Case intRetVal
          Case vbKeyTab
14290       With Me
14300         intRetVal = 0
14310         .miscType.SetFocus
14320       End With
14330     End Select
14340   End If

EXITP:
14350   KeyCode = intRetVal
14360   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub cmdMiscMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_MouseMove"

14410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14420   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdMiscMap_LTCL_GotFocus()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_LTCL_GotFocus"

14510   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
14520   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdMiscMap_LTCL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_LTCL_MouseDown"

14610   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
14620   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Public Sub cmdMiscMap_LTCL_Click()
' ** Public for Go-To-Report.

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_LTCL_Click"

14710   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
14720   Exit Sub

ERRH:
460     DoCmd.Hourglass False
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub cmdMiscMap_LTCL_KeyDown(KeyCode As Integer, Shift As Integer)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_LTCL_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14820   intShiftDown = (Shift And acShiftMask) > 0
14830   intAltDown = (Shift And acAltMask) > 0
14840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14860     Select Case intRetVal
          Case vbKeyTab
14870       With Me
14880         intRetVal = 0
14890         .cmdMiscMap_STCGL.SetFocus
14900       End With
14910     End Select
14920   End If

        ' ** Shift keys.
14930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14940     Select Case intRetVal
          Case vbKeyTab
14950       With Me
14960         intRetVal = 0
14970         .cmdMiscMap.SetFocus
14980       End With
14990     End Select
15000   End If

EXITP:
15010   KeyCode = intRetVal
15020   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub cmdMiscMap_LTCL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_LTCL_MouseMove"

15110   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
15120   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdMiscMap_LTCL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_LTCL_MouseUp"

15210   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
15220   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdMiscMap_LTCL_LostFocus()

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_LTCL_LostFocus"

15310   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
15320   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdMiscMap_STCGL_GotFocus()

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_STCGL_GotFocus"

15410   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
15420   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdMiscMap_STCGL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_STCGL_MouseDown"

15510   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
15520   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Public Sub cmdMiscMap_STCGL_Click()
' ** Public for Go-To-Report.

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_STCGL_Click"

15610   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
15620   Exit Sub

ERRH:
460     DoCmd.Hourglass False
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub cmdMiscMap_STCGL_KeyDown(KeyCode As Integer, Shift As Integer)

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_STCGL_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15720   intShiftDown = (Shift And acShiftMask) > 0
15730   intAltDown = (Shift And acAltMask) > 0
15740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15760     Select Case intRetVal
          Case vbKeyTab
15770       With Me
15780         intRetVal = 0
15790         .cmdMiscClose.SetFocus
15800       End With
15810     End Select
15820   End If

        ' ** Shift keys.
15830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15840     Select Case intRetVal
          Case vbKeyTab
15850       With Me
15860         intRetVal = 0
15870         .cmdMiscMap_LTCL.SetFocus
15880       End With
15890     End Select
15900   End If

EXITP:
15910   KeyCode = intRetVal
15920   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub cmdMiscMap_STCGL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_STCGL_MouseMove"

16010   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
16020   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdMiscMap_STCGL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_STCGL_MouseUp"

16110   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
16120   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdMiscMap_STCGL_LostFocus()

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscMap_STCGL_LostFocus"

16210   Map_Handler_Sub5 THIS_PROC, blnMiscMap_LTCL_Focus, blnMiscMap_LTCL_MouseDown, _
          blnMiscMap_STCGL_Focus, blnMiscMap_STCGL_MouseDown, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
16220   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub miscAccountNo_AfterUpdate()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "miscAccountNo_AfterUpdate"

16310   ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
16320   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "miscAccountNo_KeyDown"

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
            ' ** Tab key copies AccountNo. {intRetVal not cleared.}
16470       With Me
16480         If gblnTabCopyAccount Then
16490           If .NewRecord = True Then
16500             If IsNull(.miscAccountNo) = True Then
16510               If Len(.miscAccountNo.text) = 0 Then
16520                 If IsNull(.Parent.LastAcctNo) = False Then
16530                   If .Parent.LastAcctNo <> vbNullString Then
16540                     .miscAccountNo = .Parent.LastAcctNo
16550                     ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
16560                   End If
16570                 End If
16580               End If
16590             Else
16600               If Len(.miscAccountNo.text) = 0 Then
16610                 If IsNull(.Parent.LastAcctNo) = False Then
16620                   If .Parent.LastAcctNo <> vbNullString Then
16630                     .miscAccountNo = .Parent.LastAcctNo
16640                     ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
16650                   End If
16660                 End If
16670               End If
16680             End If
16690           End If
16700         End If
16710         gblnMiscChanged = True
16720         gblnMiscValidated = False
16730       End With
16740     Case vbKeyReturn
16750       With Me
16760         intRetVal = 0
16770         gblnMiscChanged = True  ' ** Just added: 09/19/2012.
16780         gblnMiscValidated = False
16790 On Error Resume Next
16800         blnFromAcctNoKeyDown = True
16810         .miscType.SetFocus
16820 On Error GoTo ERRH
16830       End With
16840     Case vbKeyDelete
16850       gblnMiscValidated = False
16860     End Select
16870   End If

EXITP:
16880   KeyCode = intRetVal
16890   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub miscAccountNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "miscAccountNo_MouseMove"

16910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16920   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub miscAccountNo_Exit(Cancel As Integer)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "miscAccountNo_Exit"

        Dim strAccountNo As String
        Dim blnGoToAccountHelper As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

17010   With Me

17020     blnGoToAccountHelper = False   ' ** Unless set otherwise.

17030     If ((gblnMiscValidated = False And gblnMiscChanged = True) And (lngErrCnt < 3&)) Then
17040       If IsNull(.miscAccountNo) = True Then
17050         lngErrCnt = lngErrCnt + 1&
17060         msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Entry Required")
17070         If msgResponse = vbYes Then
17080           blnGoToAccountHelper = True
17090         Else
17100           gblnMiscValidated = False
17110           If blnStartTrans = True Then
17120             Cancel = -1
17130             .miscAccountNo.SetFocus
17140           End If
17150         End If
17160       Else
17170         If .miscAccountNo = vbNullString And .miscAccountNo.text = vbNullString Then
17180           lngErrCnt = lngErrCnt + 1&
17190           msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                  "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Entry Required")
17200           If msgResponse = vbYes Then
17210             blnGoToAccountHelper = True
17220           Else
17230             gblnMiscValidated = False
17240             If blnStartTrans = True Then
17250               Cancel = -1
17260               .miscAccountNo.SetFocus
17270             End If
17280           End If
17290         Else
                ' ** Check to see if it is a valid account.
17300           lngErrCnt = 0&
17310           Select Case gblnAccountNoWithType
                Case True
17320             strAccountNo = Left(.miscAccountNo, 2) & Mid(.miscAccountNo, 3)
17330           Case False
17340             strAccountNo = .miscAccountNo
17350           End Select
17360           varTmp00 = DLookup("[accountno]", "account", "[accountno] = '" & strAccountNo & "' And ([dateclosed] Is Null)")
17370           Select Case IsNull(varTmp00)
                Case True
17380             blnStartTrans = True   ' ** Once they've started, they've got to enter a valid number!
17390             msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                    "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Invalid Entry")
17400             If msgResponse = vbYes Then
17410               blnGoToAccountHelper = True
17420             Else
17430               gblnMiscValidated = False
17440               Cancel = -1
17450               .miscAccountNo.SetFocus
17460             End If
17470             .miscAccountNo = vbNullString
17480           Case False
17490             gblnMiscValidated = True
17500           End Select
17510           If IsNull(.miscAccountNo) = False Then
17520             If .miscAccountNo <> vbNullString Then
17530               .miscAccountNo_Data = .miscAccountNo
17540               gstrMiscAccountNumber = .miscAccountNo.text
17550               .Parent.LastAcctNo = gstrMiscAccountNumber
17560               .cmdMiscMap.Enabled = False
17570 On Error Resume Next
17580               .Refresh
17590               If ERR.Number <> 0 Then  ' ** 3314: The field 'journal.journaltype' cannot contain a Null value because the Required property for this field is set to True.
17600                 .miscType = "Misc."
17610                 blnAccountNoErr = True
17620               End If
17630 On Error GoTo ERRH
17640             End If
17650           End If
17660         End If
17670       End If
17680     Else
17690       lngErrCnt = 0&
17700     End If

17710     If blnGoToAccountHelper Then
17720       .cmbAccountHelper.SetFocus
17730     End If

17740   End With

EXITP:
17750   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub cmdLock_Click()

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLock_Click"

17810   With Me
17820     If .cmdLock_open_raised_img.Visible = True Then
17830       .miscTransDate.SetFocus
17840       .miscAccountNo.Enabled = False
17850       .miscAccountNo.BorderColor = WIN_CLR_DISR
17860       .miscAccountNo.BackStyle = acBackStyleTransparent
17870       .miscAccountNo_lbl.BackStyle = acBackStyleTransparent
17880       .miscAccountNo_lbl_box.Visible = True
17890       .cmdLock_open_raised_img.Visible = False
17900       .cmdLock_closed_raised_img.Visible = True
17910     ElseIf .cmdLock_closed_raised_img.Visible = True Then
17920       .miscAccountNo.Enabled = True
17930       .miscAccountNo.BorderColor = CLR_LTBLU2
17940       .miscAccountNo.BackStyle = acBackStyleNormal
17950       .miscAccountNo_lbl.BackStyle = acBackStyleNormal
17960       .miscAccountNo_lbl_box.Visible = False
17970       .miscAccountNo.SetFocus
17980       .cmdLock_open_raised_img.Visible = True
17990       .cmdLock_closed_raised_img.Visible = False
18000     End If
18010   End With

EXITP:
18020   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub cmbAccountHelper_Enter()

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_Enter"

18110   With Me
18120     Select Case IsNull(.cmbAccountHelper)
          Case True
18130       ExpandCombo .cmbAccountHelper  ' ** Module Procedure: modUtilities.
18140     Case False
18150       If Trim(.cmbAccountHelper) = vbNullString Then
18160         ExpandCombo .cmbAccountHelper  ' ** Module Procedure: modUtilities.
18170       End If
18180     End Select
18190   End With

EXITP:
18200   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmbAccountHelper_GotFocus()

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_GotFocus"

18310   With Me
18320     .cmbAccountHelper.ForeColor = 0
18330   End With

EXITP:
18340   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub cmbAccountHelper_AfterUpdate()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_AfterUpdate"

18410   With Me
18420     If IsNull(.cmbAccountHelper) = False Then
18430       .miscAccountNo = .cmbAccountHelper
18440       .miscAccountNo.SetFocus
18450     End If
18460   End With

EXITP:
18470   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
470     Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub

Private Sub cmbAccountHelper_LostFocus()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_LostFocus"

18510   With Me
18520     gblnMiscValidated = False
18530     ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
18540     .cmbAccountHelper.ForeColor = CLR_WHT
18550     .miscAccountNo.SetFocus
18560     DoEvents
18570     .miscAccountNo.Value = .cmbAccountHelper.Value
18580     DoEvents
18590   End With

EXITP:
18600   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscShortname_LostFocus()

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "miscShortname_LostFocus"

18710   VerifyAccountNo  ' ** Function: Below.

EXITP:
18720   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscType_Enter()

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "miscType_Enter"

18810   With Me
18820     If IsNull(.miscAccountNo) = False Then
18830       If Trim(.miscAccountNo) <> vbNullString Then

18840         If IsNull(.miscType) = True Or blnAccountNoErr = True Then
18850           blnAccountNoErr = False
18860           blnTypeIsNull = True
18870           ExpandCombo .miscType  ' ** Module Procedure: modUtilities.
18880           .miscType = .miscType.ItemData(0)
18890           blnDefTypeAssigned = True  ' ** Currious that miscType_AfterUpdate() stopped being triggered!?
18900         Else
18910           blnTypeIsNull = False
18920         End If

18930 On Error Resume Next
18940         DoCmd.RunCommand acCmdSaveRecord
18950 On Error GoTo ERRH

18960         If .miscType <> .miscType.OldValue Or IsNull(.miscType.OldValue) Then
18970           ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
18980         End If

18990         VerifyAccountNo  ' ** Function: Below.

19000       End If
19010     End If
19020   End With

EXITP:
19030   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscType_AfterUpdate()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "miscType_AfterUpdate"

19110   With Me

19120     If .miscType <> "Misc." And .miscType <> "Paid" And .miscType <> "Received" Then
19130       MsgBox "Please enter a valid type 'Misc.' or 'Paid', 'Received'. ", vbInformation + vbOKOnly, "Invalid Entry"
19140       .miscType.SetFocus
19150     Else

19160       RevCodeMisc_Sub5 Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
19170       TaxCodeMisc_Sub5 Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

19180       If blnTypeIsNull = True Or .miscType <> .miscType.OldValue Or blnDefTypeAssigned = True Then

19190         blnDefTypeAssigned = False

19200         ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
19210         .miscICash = 0
19220         .miscICash_usd = Null
19230         miscICash_AfterUpdate  ' ** Procedure: Below.
19240         .miscPCash = 0
19250         .miscPCash_usd = Null
19260         miscPCash_AfterUpdate  ' ** Procedure: Below.

              ' ** Usually, making a choice here means it's a new record.
              ' ** But, it could be changed after other fields have been entered.
              ' ** So, how should the curr_id BackColor be handled?
              ' ** Since this always Zero's out values, it'll always be new!
              ' ** Although, if curr_id isn't 150, then obviously they had something else in mind!
19270         If .miscType = "Paid" Then
19280           .cmbRecurringItems.Enabled = True
19290           .cmbRecurringItems.BorderColor = CLR_LTBLU2
19300           .cmbRecurringItems.BackStyle = acBackStyleNormal
19310           .cmbRecurringItems_lbl.BackStyle = acBackStyleNormal
19320           .cmbRecurringItems_lbl_box.Visible = False
19330           .cmbRecurringItems.RowSource = "qryRecurringItems_01b"  ' ** Payee.
19340           .cmdMiscRecurNew.Caption = "&New Payee"
19350           .cmdMiscRecurNew.ControlTipText = "New Payee"
19360           .cmdMiscRecurNew.Enabled = True
19370           .chkChecks.Enabled = True
19380           .miscCurr_ID.Enabled = True
19390           .miscCurr_ID.Locked = False
19400           .miscCurr_ID_cmd.Visible = False
19410           .miscCurr_ID_cmd.Enabled = False
19420           If .miscCurr_ID = 150& Then
19430             .miscCurr_ID.ForeColor = CLR_BLK
19440             .miscCurr_ID.BackColor = CLR_WHT
19450           Else
19460             .miscCurr_ID.ForeColor = CLR_DISABLED_FG
19470             .miscCurr_ID.BackColor = CLR_VLTGRN
19480           End If
19490           .cmbRecurringItems.Requery
19500         Else
                ' ** 'Paid' is gone, handled above.
19510           If .miscType = "Received" Then
19520             .cmbRecurringItems.Enabled = True
19530             .cmbRecurringItems.BorderColor = CLR_LTBLU2
19540             .cmbRecurringItems.BackStyle = acBackStyleNormal
19550             .cmbRecurringItems_lbl.BackStyle = acBackStyleNormal
19560             .cmbRecurringItems_lbl_box.Visible = False
19570             .cmbRecurringItems.RowSource = "qryRecurringItems_01c"  ' ** Payor.
19580             .cmdMiscRecurNew.Caption = "&New Payor"
19590             .cmdMiscRecurNew.ControlTipText = "New Payor"
19600             .cmdMiscRecurNew.Enabled = True
19610             .chkChecks.Enabled = False
19620             .miscCurr_ID.Enabled = True
19630             .miscCurr_ID.Locked = False
19640             .miscCurr_ID_cmd.Visible = False
19650             .miscCurr_ID_cmd.Enabled = False
19660             If .miscCurr_ID = 150& Then
19670               .miscCurr_ID.ForeColor = CLR_BLK
19680               .miscCurr_ID.BackColor = CLR_WHT
19690             Else
19700               .miscCurr_ID.ForeColor = CLR_DISABLED_FG
19710               .miscCurr_ID.BackColor = CLR_VLTGRN
19720             End If
19730             .cmbRecurringItems.Requery
19740           Else
                  ' ** 'Paid' and 'Received' are gone, handled above.
19750             If .miscType = "Misc." Then
19760               .cmbRecurringItems.Enabled = True
19770               .cmbRecurringItems.BorderColor = CLR_LTBLU2
19780               .cmbRecurringItems.BackStyle = acBackStyleNormal
19790               .cmbRecurringItems_lbl.BackStyle = acBackStyleNormal
19800               .cmbRecurringItems_lbl_box.Visible = False
19810               .cmbRecurringItems.RowSource = "qryRecurringItems_01a"  ' ** Misc.
19820               .cmdMiscRecurNew.Caption = "&New Item"
19830               .cmdMiscRecurNew.ControlTipText = "New Item"
19840               .cmdMiscRecurNew.Enabled = True
19850               .chkChecks.Enabled = False
19860               .miscCurr_ID.Enabled = True
19870               .miscCurr_ID.Locked = False
19880               .miscCurr_ID_cmd.Visible = False
19890               .miscCurr_ID_cmd.Enabled = False
19900               If .miscCurr_ID = 150& Then
19910                 .miscCurr_ID.ForeColor = CLR_BLK
19920                 .miscCurr_ID.BackColor = CLR_WHT
19930               Else
19940                 .miscCurr_ID.ForeColor = CLR_DISABLED_FG
19950                 .miscCurr_ID.BackColor = CLR_VLTGRN
19960               End If
19970               .cmbRecurringItems.Requery
19980             Else
                    ' ** 'Paid', 'Received', and 'Misc.' are all gone, handled above.
                    ' ** There is nothing left, so this must be a new record.
19990               .cmbRecurringItems.Enabled = False
20000               .cmbRecurringItems.BorderColor = WIN_CLR_DISR
20010               .cmbRecurringItems.BackStyle = acBackStyleTransparent
20020               .cmbRecurringItems_lbl.BackStyle = acBackStyleTransparent
20030               .cmbRecurringItems_lbl_box.Visible = True
20040               .cmdMiscRecurNew.Enabled = False
20050               .cmdMiscRecurNew.Caption = "&New Item"
20060               .cmdMiscRecurNew.ControlTipText = "New Item"
20070               .chkChecks.Enabled = False
20080               .miscCurr_ID.Enabled = False
20090               .miscCurr_ID.Locked = True
20100               .miscCurr_ID_cmd.Visible = True
20110               .miscCurr_ID_cmd.Enabled = True
20120               If .miscCurr_ID = 150& Then
20130                 .miscCurr_ID.ForeColor = CLR_BLK
20140                 .miscCurr_ID.BackColor = CLR_WHT
20150               Else
20160                 .miscCurr_ID.ForeColor = CLR_DISABLED_FG
20170                 .miscCurr_ID.BackColor = CLR_VLTGRN
20180               End If
20190             End If
20200           End If
20210         End If

20220         blnGoToSaleForm = False
20230         If .miscType = "Paid" Then
20240           .tglMiscSale.Enabled = True
20250           .tglMiscSale_false_raised_img.Visible = True
20260           .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
20270           .tglMiscSale_false_raised_focus_img.Visible = False
20280           .tglMiscSale_false_raised_focus_dots_img.Visible = False
20290           .tglMiscSale_false_sunken_focus_dots_img.Visible = False
20300           .tglMiscSale_false_raised_img_dis.Visible = False
20310           .tglMiscSale_true_raised_img.Visible = False
20320           .tglMiscSale_true_raised_focus_img.Visible = False
20330           .tglMiscSale_true_raised_focus_dots_img.Visible = False
20340           .tglMiscSale_true_sunken_focus_dots_img.Visible = False
20350           .tglMiscSale_true_raised_img_dis.Visible = False
20360         End If

20370         blnGoToRecReinvest = False
20380         If .miscType = "Received" Then
20390           .tglMiscReinvest.Enabled = True
20400           .tglMiscReinvest_false_raised_img.Visible = True
20410           .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
20420           .tglMiscReinvest_false_raised_focus_img.Visible = False
20430           .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
20440           .tglMiscReinvest_false_sunken_focus_dots_img.Visible = False
20450           .tglMiscReinvest_false_raised_img_dis.Visible = False
20460           .tglMiscReinvest_true_raised_img.Visible = False
20470           .tglMiscReinvest_true_raised_focus_img.Visible = False
20480           .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
20490           .tglMiscReinvest_true_sunken_focus_dots_img.Visible = False
20500           .tglMiscReinvest_true_raised_img_dis.Visible = False
20510         End If

20520         If IsNull(.miscType.OldValue) = False Then
20530           DoCmd.RunCommand acCmdSaveRecord
20540         End If

20550       End If

20560       miscICash_AfterUpdate  ' ** Procedure: Below.
20570       miscPCash_AfterUpdate  ' ** Procedure: Below.

20580     End If

20590   End With

EXITP:
20600   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "miscType_MouseMove"

20710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
20720   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub miscType_Exit(Cancel As Integer)

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "miscType_Exit"

20810   With Me
20820     If IsNull(.miscAccountNo) = False Then
20830       If Trim(.miscAccountNo) <> vbNullString Then
20840         VerifyAccountNo  ' ** Function: Below.
20850       End If
20860     End If
20870   End With

        ' ** Currious that miscType_AfterUpdate() stopped being triggered!?
20880   If blnDefTypeAssigned = True Then
20890     miscType_AfterUpdate  ' ** Procedure: Above.
20900   End If

EXITP:
20910   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscType_LostFocus()

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "miscType_LostFocus"

21010   With Me
21020     If IsNull(.miscAccountNo) = False Then
21030       If Trim(.miscAccountNo) <> vbNullString Then

21040         gstrMiscType = .miscType.text

21050         If .miscType <> .miscType.OldValue Or IsNull(.miscType.OldValue) Then
21060           ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
21070         End If

21080         TaxCodeMisc_Sub5 Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
21090         RevCodeMisc_Sub5 Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

21100         VerifyAccountNo  ' ** Function: Below.

21110       End If
21120     End If
21130   End With

EXITP:
21140   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscTransDate_BeforeUpdate(Cancel As Integer)

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "miscTransDate_BeforeUpdate"

21210   With Me
21220     If .miscTransDate.text <> "__/__/____" And .miscTransDate.text <> vbNullString Then
21230       If Left(.miscTransDate.text, 2) > 12 Then
21240         MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
21250         Cancel = -1
21260       End If
21270     End If
21280   End With

EXITP:
21290   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscTransDate_KeyDown(KeyCode As Integer, Shift As Integer)

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "miscTransDate_KeyDown"

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
21390         .cmbRecurringItems.SetFocus
21400       End With
21410     End Select
21420   End If

        ' ** Shift keys.
21430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21450       With Me
21460         intRetVal = 0
21470         .miscType.SetFocus
21480       End With
21490     End Select
21500   End If

        ' ** Ctrl keys.
21510   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
21520     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21530       With Me
21540         intRetVal = 0
21550         .cmdCalendar1.SetFocus
21560       End With
21570     End Select
21580   End If

EXITP:
21590   KeyCode = intRetVal
21600   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub miscTransDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "miscTransDate_MouseMove"

21710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
21720   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub miscTransDate_Exit(Cancel As Integer)

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "miscTransDate_Exit"

21810   With Me
21820     If lngDateCnt <= 1& Then
21830       If IsNull(.miscAccountNo) = False Then
21840         If Trim(.miscAccountNo) <> vbNullString Then
21850           If DateCheck_Post(.miscTransDate.text) = True Then  ' ** Module Function: modUtilities.
                  ' ** '.Text' needed to reliably check date.
21860             lngDateCnt = 0&
21870             If IsNull(.miscTransDate.OldValue) Then
21880               ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
21890             Else
21900               If .miscTransDate <> .miscTransDate.OldValue Then
21910                 ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
21920               End If
21930             End If
21940           Else
21950             If IsNull(.miscTransDate) = True Then
21960               .miscTransDate = Date
21970             Else
21980               If Trim(.miscTransDate) = vbNullString Then
21990 On Error Resume Next
22000                 .miscTransDate = Date
22010                 If ERR.Number <> 0 Then
                        ' ** 2473  The expression |2 you entered as the event property setting produced the following error: |1.
                        ' **       * The expression may not result in the name of a macro, the name of a user-defined function, or [Event Procedure].
                        ' **       * There may have been an error evaluating the function, event, or macro.
22020                   .miscTransDate = TodaysDate  ' ** Module Function: modUtilities.
22030 On Error GoTo ERRH
22040                 Else
22050 On Error GoTo ERRH
22060                 End If
22070               Else
22080                 lngDateCnt = lngDateCnt + 1&
22090                 Cancel = -1
22100                 .miscTransDate.SetFocus
22110               End If
22120             End If
22130           End If
22140         End If
22150       End If
22160     Else
22170       lngDateCnt = 0&
22180       .miscTransDate = CDate(Format(datPostingDate, "mm/dd/yyyy"))
22190     End If
22200   End With

EXITP:
22210   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscTransDate_LostFocus()

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "miscTransDate_LostFocus"

22310   With Me
22320     If IsNull(.miscAccountNo) = False Then
22330       If Trim(.miscAccountNo) <> vbNullString Then
22340         VerifyAccountNo  ' ** Function: Below.
22350       End If
22360     End If
22370   End With

EXITP:
22380   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

22410   Calendar_Handler_Sub5 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
22420   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

22510   Calendar_Handler_Sub5 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
22520   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Public Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

22610   Calendar_Handler_Sub5 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
22620   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
470     Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

22710   Calendar_Handler_Sub5 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
22720   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

22810   Calendar_Handler_Sub5 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
22820   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

22910   Calendar_Handler_Sub5 THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
22920   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmbRecurringItems_Enter()

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_Enter"

        Dim strTmp02 As String, strTmp03 As String

        Static intLoopCnt As Integer

23010   strTmp02 = strMemo_Orig
23020   strTmp03 = strMemo_New

        ' ** Make sure the Combo Box is populated.
23030   If intLoopCnt = 0 Then
23040     intLoopCnt = 1
23050     blnCmdButton = True
23060     Form_Current  ' ** Procedure: Above.
23070   Else
23080     intLoopCnt = 0
23090   End If

23100   With Me

23110     blnCmdButton = False
23120     .cmbRecurringItems.SetFocus
23130     .cmbRecurringItems.Requery
23140     DoEvents

23150     If IsNull(.cmbRecurringItems) = True Or .cmbRecurringItems.text = vbNullString Then
23160       ExpandCombo .cmbRecurringItems  ' ** Module Procedure: modUtilities.
23170     End If

23180   End With

23190   strMemo_Orig = strTmp02
23200   strMemo_New = strTmp03

EXITP:
23210   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub cmbRecurringItems_AfterUpdate()
' ** CBX_R_NAME As Integer = 0
' ** CBX_R_TYPE As Integer = 1
' ** CBX_R_RID   As Integer = 2
' ** Wingdings arrows:
' **   à  Pointing Right.
' **   ß  Pointing Left.

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_AfterUpdate"

23310   With Me
23320     .MiscXFer_ItoP_lbl.Visible = False
23330     .MiscXFer_PtoI_lbl.Visible = False
23340     If .cmbRecurringItems.Column(CBX_R_RID) = RECUR_I_TO_P_ID Then
23350       .MiscXFer_ItoP_lbl.Visible = True
23360       .MiscXFer_PtoI_lbl.Visible = False
23370       If IsNull(.miscICash) = False And IsNull(.miscPCash) = False Then
23380         If .miscICash <> 0 And .miscPCash <> 0 Then
23390           If Abs(.miscICash) <> Abs(.miscPCash) Then
23400             MsgBox "Cash transfers must be of equal amounts.", vbInformation + vbOKOnly, "Invalid Entry"
23410             .miscPCash.SetFocus
23420           End If
23430           If .miscICash > 0 Then
23440             .miscICash = .miscICash * -1
23450             miscICash_AfterUpdate
23460           End If
23470           If .miscPCash < 0 Then
23480             .miscPCash = .miscPCash * -1
23490             miscPCash_AfterUpdate
23500           End If
23510         End If
23520       End If
23530     ElseIf .cmbRecurringItems.Column(CBX_R_RID) = RECUR_P_TO_I_ID Then
23540       .MiscXFer_ItoP_lbl.Visible = False
23550       .MiscXFer_PtoI_lbl.Visible = True
23560       If IsNull(.miscPCash) = False And IsNull(.miscICash) = False Then
23570         If .miscICash <> 0 And .miscPCash <> 0 Then
23580           If Abs(.miscICash) <> Abs(.miscPCash) Then
23590             MsgBox "Cash transfers must be of equal amounts.", vbInformation + vbOKOnly, "Invalid Entry"
23600             .miscPCash.SetFocus
23610           End If
23620           If .miscICash < 0 Then
23630             .miscICash = .miscICash * -1
23640             miscICash_AfterUpdate
23650           End If
23660           If .miscPCash > 0 Then
23670             .miscPCash = .miscPCash * -1
23680             miscPCash_AfterUpdate
23690           End If
23700         End If
23710       End If
23720     End If
23730   End With

EXITP:
23740   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub cmbRecurringItems_KeyDown(KeyCode As Integer, Shift As Integer)

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23820   intShiftDown = (Shift And acShiftMask) > 0
23830   intAltDown = (Shift And acAltMask) > 0
23840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23870       With Me
23880         intRetVal = 0
23890         If .miscCurr_ID.Visible = True Then
23900           .miscCurr_ID.SetFocus
23910         Else
23920           .miscICash.SetFocus
23930         End If
23940       End With
23950     End Select
23960   End If

        ' ** Shift keys.
23970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23980     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23990       With Me
24000         intRetVal = 0
24010         .miscTransDate.SetFocus
24020       End With
24030     End Select
24040   End If

        ' ** Ctrl keys.
24050   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
24060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24070       With Me
24080         intRetVal = 0
24090         .cmdMiscRecurNew.SetFocus
24100       End With
24110     End Select
24120   End If

        ' ** Ctrl-Shift keys.
24130   If intCtrlDown And (Not intAltDown) And intShiftDown Then
24140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24150       With Me
24160         intRetVal = 0
24170         .cmdCalendar1.SetFocus
24180       End With
24190     End Select
24200   End If

EXITP:
24210   KeyCode = intRetVal
24220   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub cmbRecurringItems_Change()

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_Change"

24310   With Me
24320     If .cmbRecurringItems.Value <> .cmbRecurringItems.OldValue Then
24330       ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
24340     End If
24350   End With

EXITP:
24360   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub cmbRecurringItems_NotInList(NewData As String, Response As Integer)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_NotInList"

24410   With Me
24420     .cmbRecurringItems.SelStart = 0
24430     .cmbRecurringItems.SelLength = Len(.cmbRecurringItems.text)
24440   End With

EXITP:
24450   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub cmbRecurringItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_MouseMove"

24510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
24520   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub cmdMiscRecurNew_Click()

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscRecurNew_Click"

        Dim strDocName As String, strJournalType As String

24610   With Me
24620     strJournalType = .miscType
24630     strDocName = "frmRecurringItems_Add_Misc"
24640     gblnSetFocus = True
24650     DoCmd.OpenForm strDocName, , , , acFormAdd, acDialog, "frmJournal" & "~" & strJournalType
24660   End With

EXITP:
24670   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub cmdMiscRecurNew_KeyDown(KeyCode As Integer, Shift As Integer)
' ** See Note, above.

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMiscRecurNew_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24720   intShiftDown = (Shift And acShiftMask) > 0
24730   intAltDown = (Shift And acAltMask) > 0
24740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24760     Select Case intRetVal
          Case vbKeyTab
24770       With Me
24780         intRetVal = 0
24790         .miscICash.SetFocus
24800       End With
24810     End Select
24820   End If

        ' ** Shift keys.
24830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24840     Select Case intRetVal
          Case vbKeyTab
24850       With Me
24860         intRetVal = 0
24870         .cmbRecurringItems.SetFocus
24880       End With
24890     End Select
24900   End If

EXITP:
24910   KeyCode = intRetVal
24920   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub miscAssetNo_AfterUpdate()

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "miscAssetNo_AfterUpdate"

25010   With Me

          ' ** Do this whether or not Currency is visible.
25020     Select Case IsNull(.miscAssetNo)
          Case True
25030       lngCurrID = 150&  ' ** USD.
25040     Case False
25050       Select Case IsNull(.miscAssetNo.Column(3))
            Case True
25060         lngCurrID = 150&  ' ** USD.
25070       Case False
25080         lngCurrID = .miscAssetNo.Column(3)
25090       End Select
25100     End Select
25110     .miscCurr_ID = lngCurrID
25120     miscCurr_ID_AfterUpdate  ' ** Procedure: Below.

25130   End With

EXITP:
25140   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
490     Resume EXITP

End Sub

Private Sub miscAssetNo_NotInList(NewData As String, Response As Integer)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "miscAssetNo_NotInList"

25210   With Me
25220     .miscAssetNo.SelStart = 0
25230     .miscAssetNo.SelLength = Len(.miscAssetNo.text)
25240   End With

EXITP:
25250   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscAssetNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "miscAssetNo_MouseMove"

25310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
25320   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
480     End Select
490     Resume EXITP

End Sub

Private Sub miscAssetNo_Exit(Cancel As Integer)

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "miscAssetNo_Exit"

        Dim varTmp00 As Variant

25410   With Me
          ' ** Check curr_id.
25420     varTmp00 = .miscAssetNo.Column(3)
25430     If IsNull(varTmp00) = False Then
25440       If .miscCurr_ID <> varTmp00 Then
25450         .miscCurr_ID = varTmp00
25460         miscCurr_ID_AfterUpdate  ' ** Procedure: Below.
25470       End If
25480     End If
25490   End With

EXITP:
25500   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscICash_AfterUpdate()

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "miscICash_AfterUpdate"

        Dim dblICash As Double
        Dim varTmp00 As Variant

25610   With Me
25620     If .miscCurr_ID.Visible = True Then
25630       .miscICash_usd = Null
25640       If .miscICash_usd.Visible = False Then .miscICash_usd.Visible = True
25650       If IsNull(.miscCurr_ID.Column(CBX_C_CURRID)) = False Then
25660         lngCurrID = .miscCurr_ID.Column(CBX_C_CURRID)
25670         If lngCurrID <> 150& Then  ' ** USD.
25680           If IsNull(.miscICash) = False Then
25690             If .miscICash <> 0 Then
25700               varTmp00 = .miscCurr_ID.Column(CBX_C_RATE2)
25710               If IsNull(varTmp00) = False Then
25720                 If varTmp00 > 0 Then
25730                   dblICash = .miscICash
25740                   dblICash = (dblICash * varTmp00)
25750                   dblICash = Round(dblICash, 2)
25760                   .miscICash_usd = dblICash
25770                   If IsNull(.miscPCash) = False Then
25780                     If .miscPCash = 0 Then
25790                       If IsNull(.miscPCash_usd) = True Then
25800                         miscPCash_AfterUpdate  ' ** Procedure: Below.
25810                       End If
25820                     End If
25830                   End If
25840                 End If
25850               End If
25860             Else
25870               .miscICash_usd = 0#
25880             End If
25890           Else
25900             .miscICash_usd = Null
25910           End If  ' ** Null.
25920         Else
25930           .miscICash_usd = Null
25940         End If  ' ** USD.
25950       Else
25960         .miscICash_usd = Null
25970       End If  ' ** CBX_C_CURRID.
25980     End If
25990     DoEvents
26000   End With

EXITP:
26010   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscICash_KeyDown(KeyCode As Integer, Shift As Integer)

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "miscICash_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
26170       With Me
26180         intRetVal = 0
26190         .miscPCash.SetFocus
26200       End With
26210     End Select
26220   End If

        ' ** Shift keys.
26230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26250       With Me
26260         intRetVal = 0
26270         If .miscCurr_ID.Visible = True Then
26280           .miscCurr_ID.SetFocus
26290         Else
26300           .cmbRecurringItems.SetFocus
26310         End If
26320       End With
26330     End Select
26340   End If

        ' ** Ctrl-Shift keys.
26350   If intCtrlDown And (Not intAltDown) And intShiftDown Then
26360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26370       With Me
26380         intRetVal = 0
26390         If .miscCurr_ID.Visible = False Then
26400           .cmdMiscRecurNew.SetFocus
26410         End If
26420       End With
26430     End Select
26440   End If

EXITP:
26450   KeyCode = intRetVal
26460   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub miscICash_Exit(Cancel As Integer)

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "miscICash_Exit"

        Dim blnContinue As Boolean
        Dim intPos01 As Integer

26510   With Me

26520     blnContinue = True

26530 On Error Resume Next
26540     lngCurrID = .miscCurr_ID.Column(CBX_C_CURRID)
26550     If ERR.Number <> 0 Then
            ' ** 94  Invalid use of Null.
26560 On Error GoTo ERRH
26570       .miscCurr_ID.Requery
26580       .miscCurr_ID = 150
26590       lngCurrID = 150&  ' ** Default to USD.
26600     Else
26610 On Error GoTo ERRH
26620     End If

26630     If IsNumeric(.miscICash) Then
26640       If CDbl(Rem_Dollar(.miscICash.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
26650         blnContinue = False
26660         MsgBox "Income Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
26670         .miscICash = vbNullString
26680         .miscICash_usd = Null
26690         miscICash_AfterUpdate  ' ** Procedure: Above.
26700         Cancel = -1
26710         .miscICash.SetFocus
26720       Else
26730         If .miscType = "Received" And CDbl(Rem_Dollar(.miscICash.text, lngCurrID)) < 0# Then  ' ** Module Function: modStringFuncs.
26740           blnContinue = False
26750           MsgBox "Income Cash must be greater than zero for a Received transaction.", vbInformation + vbOKOnly, "Invalid Entry"
26760           .miscICash = vbNullString
26770           .miscICash_usd = Null
26780           miscICash_AfterUpdate  ' ** Procedure: Above.
26790           Cancel = -1
26800           .miscICash.SetFocus
26810         End If
26820       End If
26830     Else
26840       blnContinue = False
26850       MsgBox "You must enter a numeric value in the Income Cash field.", vbInformation + vbOKOnly, "Entry Required"
26860       .miscICash = vbNullString
26870       .miscICash_usd = Null
26880       miscICash_AfterUpdate  ' ** Procedure: Above.
26890       Cancel = -1
26900       .miscICash.SetFocus
26910     End If

26920     If blnContinue = True Then
26930       intPos01 = InStr(1, .miscICash, ".", vbTextCompare)
26940       If intPos01 > 0 Then
26950         If Len(.miscICash) - intPos01 > 2 Then
26960           blnContinue = False
26970           MsgBox "You can only enter Income Cash with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
26980           .miscICash = vbNullString
26990           .miscICash_usd = Null
27000           miscICash_AfterUpdate  ' ** Procedure: Above.
27010           Cancel = -1
27020           .miscICash.SetFocus
27030         End If
27040       End If
27050     End If

27060     If blnContinue = True Then

27070       Select Case .miscType
            Case "Paid"
27080         If .miscICash > 0 Then
27090           .miscICash = .miscICash * -1
27100           miscICash_AfterUpdate  ' ** Procedure: Above.
27110         End If
27120       Case "Misc."
27130         If IsNull(.cmbRecurringItems) = False Then
27140           If .cmbRecurringItems <> vbNullString Then
27150             If .cmbRecurringItems.Column(CBX_R_RID) = RECUR_I_TO_P_ID Then
27160               If IsNull(.miscICash) = False Then
27170                 If IsNull(.miscPCash) = False Then
27180                   If .miscICash <> 0 And .miscPCash <> 0 Then
27190                     If Abs(.miscICash) <> Abs(.miscPCash) Then
27200                       MsgBox "Cash transfers must be of equal amounts.", vbInformation + vbOKOnly, "Invalid Entry"
27210                       Cancel = -1
27220                       .miscPCash.SetFocus
27230                     End If
27240                     If .miscPCash < 0 Then
27250                       .miscPCash = .miscPCash * -1
27260                       miscPCash_AfterUpdate  ' ** Procedure: Below.
27270                     End If
27280                   End If
27290                 End If
27300                 If .miscICash > 0 Then
27310                   .miscICash = .miscICash * -1
27320                   miscICash_AfterUpdate  ' ** Procedure: Above.
27330                 End If
27340               End If
27350             End If
27360             If .cmbRecurringItems.Column(CBX_R_RID) = RECUR_P_TO_I_ID Then
27370               If IsNull(.miscICash) = False Then
27380                 If IsNull(.miscPCash) = False Then
27390                   If .miscICash <> 0 And .miscPCash <> 0 Then
27400                     If Abs(.miscICash) <> Abs(.miscPCash) Then
27410                       MsgBox "Cash transfers must be of equal amounts.", vbInformation + vbOKOnly, "Invalid Entry"
27420                       Cancel = -1
27430                       .miscPCash.SetFocus
27440                     End If
27450                     If .miscPCash > 0 Then
27460                       .miscPCash = .miscPCash * -1
27470                       miscPCash_AfterUpdate  ' ** Procedure: Below.
27480                     End If
27490                   End If
27500                 End If
27510                 If .miscICash < 0 Then
27520                   .miscICash = .miscICash * -1
27530                   miscICash_AfterUpdate  ' ** Procedure: Above.
27540                 End If
27550               End If
27560             End If
27570           End If
27580         End If
27590       Case Else
              ' ** Anything else?
27600       End Select

27610 On Error Resume Next  ' ** In case its focus was changed above.
27620       gstrMiscICash = .miscICash.text
27630 On Error GoTo ERRH
27640       gstrMiscICash = Rem_Dollar(gstrMiscICash, lngCurrID)  ' ** Module Function: modStringFuncs.

27650       If blnContinue = True Then
27660         RevCodeMisc_Sub5 Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
27670         TaxCodeMisc_Sub5 Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
27680       End If

27690       If .miscICash <> .miscICash.OldValue Or IsNull(.miscICash.OldValue) Then
27700         ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
27710       End If

27720     End If
27730   End With

EXITP:
27740   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscICash_LostFocus()

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "miscICash_LostFocus"

27810   VerifyAccountNo  ' ** Function: Below.

EXITP:
27820   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscPCash_AfterUpdate()

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "miscPCash_AfterUpdate"

        Dim dblPCash As Double
        Dim varTmp00 As Variant

27910   With Me
27920     If .miscCurr_ID.Visible = True Then
27930       .miscPCash_usd = Null
27940       If .miscPCash_usd.Visible = False Then .miscPCash_usd.Visible = True
27950       If IsNull(.miscCurr_ID.Column(CBX_C_CURRID)) = False Then
27960         lngCurrID = .miscCurr_ID.Column(CBX_C_CURRID)
27970         If lngCurrID <> 150& Then  ' ** USD.
27980           If IsNull(.miscPCash) = False Then
27990             If .miscPCash <> 0 Then
28000               varTmp00 = .miscCurr_ID.Column(CBX_C_RATE2)
28010               If IsNull(varTmp00) = False Then
28020                 If varTmp00 > 0 Then
28030                   dblPCash = .miscPCash
28040                   dblPCash = (dblPCash * varTmp00)
28050                   dblPCash = Round(dblPCash, 2)
28060                   .miscPCash_usd = dblPCash
28070                   If IsNull(.miscICash) = False Then
28080                     If .miscICash = 0 Then
28090                       If IsNull(.miscICash_usd) = True Then
28100                         miscICash_AfterUpdate  ' ** Procedure: Above.
28110                       End If
28120                     End If
28130                   End If
28140                 End If
28150               End If
28160             Else
28170               .miscPCash_usd = 0#
28180             End If
28190           Else
28200             .miscPCash_usd = Null
28210           End If  ' ** Null
28220         Else
28230           .miscPCash_usd = Null
28240         End If  ' ** USD.
28250       Else
28260         .miscPCash_usd = Null
28270       End If  ' ** CBX_C_CURRID.
28280     End If
28290     DoEvents

28300   End With

EXITP:
28310   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscPCash_KeyDown(KeyCode As Integer, Shift As Integer)

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "miscPCash_KeyDown"

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
28490         .miscDescription.SetFocus
28500       End With
28510     End Select
28520   End If

        ' ** Shift keys.
28530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28550       With Me
28560         intRetVal = 0
28570         .miscICash.SetFocus
28580       End With
28590     End Select
28600   End If

        ' ** Ctrl-Shift keys.
28610   If intCtrlDown And (Not intAltDown) And intShiftDown Then
28620     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28630       With Me
28640         intRetVal = 0
28650         If .miscICash.Enabled = False Then
28660           If .miscCurr_ID.Visible = False Then
28670             .cmdMiscRecurNew.SetFocus
28680           End If
28690         End If
28700       End With
28710     End Select
28720   End If

EXITP:
28730   KeyCode = intRetVal
28740   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub miscPCash_Exit(Cancel As Integer)

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "miscPCash_Exit"

        Dim blnContinue As Boolean
        Dim intPos01 As Integer

28810   With Me

28820     blnContinue = True

28830     lngCurrID = .miscCurr_ID.Column(CBX_C_CURRID)

28840     If IsNumeric(.miscPCash) Then
28850       If CDbl(Rem_Dollar(.miscPCash.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
28860         blnContinue = False
28870         MsgBox "Principal Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
28880         .miscPCash = vbNullString
28890         .miscPCash_usd = Null
28900         miscPCash_AfterUpdate  ' ** Procedure: Below.
28910         Cancel = -1
28920         .miscPCash.SetFocus
28930       Else
28940         If .miscType = "Received" And CDbl(Rem_Dollar(.miscPCash.text, lngCurrID)) < 0# Then  ' ** Module Function: modStringFuncs.
28950           blnContinue = False
28960           MsgBox "Principal Cash must be greater than zero for a Received transaction.", vbInformation + vbOKOnly, "Invalid Entry"
28970           .miscPCash = vbNullString
28980           .miscPCash_usd = Null
28990           miscPCash_AfterUpdate  ' ** Procedure: Below.
29000           Cancel = -1
29010           .miscPCash.SetFocus
29020         End If
29030       End If
29040     Else
29050       blnContinue = False
29060       MsgBox "You must enter a numeric value in the Principal Cash field.", vbInformation + vbOKOnly, "Entry Required"
29070       .miscPCash = vbNullString
29080       .miscPCash_usd = Null
29090       miscPCash_AfterUpdate  ' ** Procedure: Below.
29100       Cancel = -1
29110       .miscPCash.SetFocus
29120     End If

29130     If blnContinue = True Then
29140       intPos01 = InStr(1, Me.miscPCash, ".", vbTextCompare)
29150       If intPos01 > 0 Then
29160         If Len(Me.miscPCash) - intPos01 > 2 Then
29170           blnContinue = False
29180           MsgBox "You can only enter Principal Cash with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
29190           .miscPCash = vbNullString
29200           .miscPCash_usd = Null
29210           miscPCash_AfterUpdate  ' ** Procedure: Below.
29220           Cancel = -1
29230           .miscPCash.SetFocus
29240         End If
29250       End If
29260     End If

29270     If blnContinue = True Then

29280       If .miscType = "Paid" Then
29290         If .miscPCash > 0 Then
29300           .miscPCash = .miscPCash * -1
29310           miscPCash_AfterUpdate  ' ** Procedure: Below.
29320         End If
29330       ElseIf .miscType = "Misc." Then
29340         If IsNull(.cmbRecurringItems) = False Then
29350           If .cmbRecurringItems <> vbNullString Then
29360             If .cmbRecurringItems.Column(CBX_R_RID) = RECUR_I_TO_P_ID Then
29370               If IsNull(.miscPCash) = False Then
29380                 If IsNull(.miscICash) = False Then
29390                   If .miscICash <> 0 And .miscPCash <> 0 Then
29400                     If Abs(.miscICash) <> Abs(.miscPCash) Then
29410                       MsgBox "Cash transfers must be of equal amounts.", vbInformation + vbOKOnly, "Invalid Entry"
29420                       Cancel = -1
29430                       .miscPCash.SetFocus
29440                       .miscPCash.SelStart = 0
29450                       .miscPCash.SelLength = 19
29460                     End If
29470                     If .miscICash > 0 Then
29480                       .miscICash = .miscICash * -1
29490                       miscICash_AfterUpdate  ' ** Procedure: Above.
29500                     End If
29510                   End If
29520                 End If
29530                 If .miscPCash < 0 Then
29540                   .miscPCash = .miscPCash * -1
29550                   miscPCash_AfterUpdate  ' ** Procedure: Below.
29560                 End If
29570               End If
29580             End If
29590             If .cmbRecurringItems.Column(CBX_R_RID) = RECUR_P_TO_I_ID Then
29600               If IsNull(.miscPCash) = False Then
29610                 If IsNull(.miscICash) = False Then
29620                   If .miscICash <> 0 And .miscPCash <> 0 Then
29630                     If Abs(.miscICash) <> Abs(.miscPCash) Then
29640                       MsgBox "Cash transfers must be of equal amounts.", vbInformation + vbOKOnly, "Invalid Entry"
29650                       Cancel = -1
29660                       .miscPCash.SetFocus
29670                       .miscPCash.SelStart = 0
29680                       .miscPCash.SelLength = 19
29690                     End If
29700                     If .miscICash < 0 Then
29710                       .miscICash = .miscICash * -1
29720                       miscICash_AfterUpdate  ' ** Procedure: Above.
29730                     End If
29740                   End If
29750                 End If
29760                 If .miscPCash > 0 Then
29770                   .miscPCash = .miscPCash * -1
29780                   miscPCash_AfterUpdate  ' ** Procedure: Below.
29790                 End If
29800               End If
29810             End If
29820           End If
29830         End If
29840       Else
              ' ** Anything else?
29850       End If

29860 On Error Resume Next  ' ** In case its focus was changed above.
29870       gstrMiscPCash = .miscPCash.text
29880 On Error GoTo ERRH
29890       gstrMiscPCash = Rem_Dollar(gstrMiscPCash, lngCurrID) ' ** Module Function: modStringFuncs.

29900       If blnContinue = True Then
29910         RevCodeMisc_Sub5 Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
29920         TaxCodeMisc_Sub5 Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
29930       End If

29940       If .miscPCash <> .miscPCash.OldValue Or IsNull(.miscPCash.OldValue) Then
29950         ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
29960       End If

29970     End If  ' ** blnContinue.
29980   End With

EXITP:
29990   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscPCash_LostFocus()

30000 On Error GoTo ERRH

        Const THIS_PROC As String = "miscPCash_LostFocus"

30010   VerifyAccountNo  ' ** Function: Below.

30020   With Me
30030     lngCurrID = .miscCurr_ID.Column(CBX_C_CURRID)
30040     If .miscType = "Paid" Then
30050       If .miscPCash > 0 Then
30060         .miscPCash = .miscPCash * -1
30070         miscPCash_AfterUpdate  ' ** Procedure: Above.
30080         gstrMiscPCash = Format(.miscPCash, "0.00")
30090         gstrMiscPCash = Rem_Dollar(gstrMiscPCash, lngCurrID)  ' ** Module Function: modStringFuncs.
30100       End If
30110     End If
30120   End With

EXITP:
30130   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscCurr_ID_cmd_DblClick(Cancel As Integer)

30200 On Error GoTo ERRH

        Const THIS_PROC As String = "miscCurr_ID_cmd_DblClick"

30210   miscCurr_ID_DblClick Cancel  ' ** Procedure: Below.

EXITP:
30220   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
490     Resume EXITP

End Sub

Private Sub miscCurr_ID_DblClick(Cancel As Integer)

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "miscCurr_ID_DblClick"

30310   With Me
30320     If IsNull(.miscCurr_Date) = True Then
30330       .miscCurr_Date = .miscCurr_ID.Column(CBX_C_DATE)
30340     End If
30350     DoEvents
30360     Select Case .miscCurr_Date.Visible
          Case True
30370       .miscCurr_Date.Visible = False
30380     Case False
30390       .miscCurr_Date.Visible = True
30400     End Select
30410   End With

EXITP:
30420   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscCurr_ID_AfterUpdate()

30500 On Error GoTo ERRH

        Const THIS_PROC As String = "miscCurr_ID_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String, strTmp03 As String, intTmp04 As Integer, lngTmp06 As Long
        Dim lngX As Long

30510   With Me

30520     If IsNull(.miscCurr_ID) = False Then

30530       lngCurrID = .miscCurr_ID.Column(CBX_C_CURRID)

30540       blnFound = False
30550       If lngCurrID <> 150& Then
30560         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
30570           arr_varNoChar = .Parent.NoChar_Get  ' ** Form Function: frmJournal.
30580           lngNoChars = UBound(arr_varNoChar, 2) + 1&
30590         End If
30600         For lngX = 0& To (lngNoChars - 1&)
30610           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
30620             blnFound = True
30630           End If
30640         Next
30650       End If

30660       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
30670         .miscICash.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
30680         .miscICash.FontSize = 10
30690         .miscICash.FontBold = True
30700         .miscPCash.FontName = "Courier New"
30710         .miscPCash.FontSize = 10
30720         .miscPCash.FontBold = True
30730       Case False
30740         If .miscICash.FontName <> "Fixedsys" Then
30750           .miscICash.FontName = "Fixedsys"
30760           .miscICash.FontSize = 9
30770           .miscICash.FontBold = False
30780           .miscPCash.FontName = "Fixedsys"
30790           .miscPCash.FontSize = 9
30800           .miscPCash.FontBold = False
30810         End If
30820       End Select

30830       If lngCurrID = 150& Then  ' ** USD.
30840         strTmp01 = "Currency"
30850         intTmp04 = 2
30860         lngTmp06 = CLR_WHT
30870       Else
30880         strTmp03 = .miscCurr_ID.Column(CBX_C_SYM)  ' ** Symbol.
30890         If .miscCurr_ID.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
30900           intTmp04 = 4
30910         Else
30920           intTmp04 = .miscCurr_ID.Column(CBX_C_DEC)  ' ** Decimal places.
30930         End If
30940         intLen = Len(strTmp03)
30950         Select Case intLen
              Case 1
                'strTmp03 = strTmp03
30960         Case 2
30970           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2)
30980         Case 3
30990           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3)
31000         Case 4
31010           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & "\" & Mid(strTmp03, 4)
31020         Case 5
31030           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & _
                  "\" & Mid(strTmp03, 4, 1) & "\" & Mid(strTmp03, 5)
31040         End Select
31050         Select Case intTmp04
              Case 0
31060           strTmp01 = strTmp03 & "#,##0;(" & strTmp03 & "#,##0);" & strTmp03 & "0;" & strTmp03 & "0"
31070         Case 1  ' ** None have this currently.
31080           strTmp01 = strTmp03 & "#,##0.0;(" & strTmp03 & "#,##0.0);" & strTmp03 & "0.0;" & strTmp03 & "0.0"
31090         Case 2
31100           strTmp01 = strTmp03 & "#,##0.00;(" & strTmp03 & "#,##0.00);" & strTmp03 & "0.00;" & strTmp03 & "0.00"
31110         Case 3
31120           strTmp01 = strTmp03 & "#,##0.000;(" & strTmp03 & "#,##0.000);" & strTmp03 & "0.000;" & strTmp03 & "0.000"
31130         End Select
31140         lngTmp06 = CLR_VLTGRN
31150       End If
31160       .miscICash.Format = strTmp01
31170       .miscICash.DecimalPlaces = intTmp04
31180       .miscICash.BackColor = lngTmp06
31190       .miscPCash.Format = strTmp01
31200       .miscPCash.DecimalPlaces = intTmp04
31210       .miscPCash.BackColor = lngTmp06
31220     End If

31230     miscICash_AfterUpdate  ' ** Procedure: Above.
31240     miscPCash_AfterUpdate  ' ** Procedure: Above.

31250     If lngCurrID <> .miscCurr_ID.OldValue Then
31260       ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
31270     End If

31280   End With

EXITP:
31290   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscCurr_ID_KeyDown(KeyCode As Integer, Shift As Integer)

31300 On Error GoTo ERRH

        Const THIS_PROC As String = "miscCurr_ID_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
31320   intShiftDown = (Shift And acShiftMask) > 0
31330   intAltDown = (Shift And acAltMask) > 0
31340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31370       With Me
31380         intRetVal = 0
31390         .miscICash.SetFocus
31400       End With
31410     End Select
31420   End If

        ' ** Shift keys.
31430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31450       With Me
31460         intRetVal = 0
31470         .cmbRecurringItems.SetFocus
31480       End With
31490     End Select
31500   End If

        ' ** Ctrl-Shift keys.
31510   If intCtrlDown And (Not intAltDown) And intShiftDown Then
31520     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31530       With Me
31540         intRetVal = 0
31550         .cmdMiscRecurNew.SetFocus
31560       End With
31570     End Select
31580   End If

EXITP:
31590   KeyCode = intRetVal
31600   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub miscCurr_ID_NotInList(NewData As String, Response As Integer)

31700 On Error GoTo ERRH

        Const THIS_PROC As String = "miscCurr_ID_NotInList"

31710   With Me
31720     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
31730     .miscCurr_ID.SetFocus
31740     Response = acDataErrContinue
31750   End With

EXITP:
31760   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscCurr_ID_Exit(Cancel As Integer)

31800 On Error GoTo ERRH

        Const THIS_PROC As String = "miscCurr_ID_Exit"

31810   With Me
31820     If .miscCurr_ID <> .miscCurr_ID.OldValue Then
31830       ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
31840     End If
31850   End With

EXITP:
31860   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscDescription_AfterUpdate()

31900 On Error GoTo ERRH

        Const THIS_PROC As String = "miscDescription_AfterUpdate"

31910   With Me
31920     If .miscDescription <> .miscAccountNo.OldValue Then
31930       ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
31940     End If
31950   End With

EXITP:
31960   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub miscDescription_KeyDown(KeyCode As Integer, Shift As Integer)

32000 On Error GoTo ERRH

        Const THIS_PROC As String = "miscDescription_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
32020   intShiftDown = (Shift And acShiftMask) > 0
32030   intAltDown = (Shift And acAltMask) > 0
32040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
32070       With Me
32080         intRetVal = 0
32090         If .chkChecks.Enabled = True Then
32100           .chkChecks.SetFocus
32110         ElseIf .tglMiscSale.Visible = True And .tglMiscSale.Enabled = True Then
32120           .tglMiscSale.SetFocus
32130         ElseIf .tglMiscReinvest.Visible = True And .tglMiscReinvest.Enabled = True Then
32140           .tglMiscReinvest.SetFocus
32150         ElseIf .cmbRevenueCodes.Visible = True And .cmbRevenueCodes.Enabled = True Then
32160           .cmbRevenueCodes.SetFocus
32170         ElseIf .cmbTaxCodes.Visible = True And .cmbTaxCodes.Enabled = True Then
32180           .cmbTaxCodes.SetFocus
32190         ElseIf .cmdMiscOK.Enabled = True Then
32200           .cmdMiscOK.SetFocus
32210         ElseIf .cmdMiscCancel.Enabled = True Then
32220           .cmdMiscCancel.SetFocus
32230         ElseIf .cmdMiscClose.Enabled = True Then
32240           .cmdMiscClose.SetFocus
32250         ElseIf .miscAccountNo.Enabled = True Then
32260           .miscAccountNo.SetFocus
32270         Else
32280           .miscTransDate.SetFocus
32290         End If
32300       End With
32310     End Select
32320   End If

        ' ** Shift keys.
32330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
32350       With Me
32360         intRetVal = 0
32370         .miscICash.SetFocus
32380       End With
32390     End Select
32400   End If

EXITP:
32410   KeyCode = intRetVal
32420   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub miscDescription_LostFocus()

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "miscDescription_LostFocus"

32510   VerifyAccountNo  ' ** Function: Below.

EXITP:
32520   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub tglMiscSale_GotFocus()

32600 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscSale_GotFocus"

32610   Tgl_Handler_Sub5 THIS_PROC, blnMiscSale_Focus, blnMiscSale_MouseDown, _
          blnMiscReinvest_Focus, blnMiscReinvest_MouseDown, blnGoToSaleForm, blnGoToRecReinvest, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
32620   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub tglMiscSale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscSale_MouseDown"

32710   Tgl_Handler_Sub5 THIS_PROC, blnMiscSale_Focus, blnMiscSale_MouseDown, _
          blnMiscReinvest_Focus, blnMiscReinvest_MouseDown, blnGoToSaleForm, blnGoToRecReinvest, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
32720   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub tglMiscSale_Click()
' ** I'm going to try using [posted] as a signal (since it's not used for anything else).

32800 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscSale_Click"

32810   Tgl_Handler_Sub5 THIS_PROC, blnMiscSale_Focus, blnMiscSale_MouseDown, _
          blnMiscReinvest_Focus, blnMiscReinvest_MouseDown, blnGoToSaleForm, blnGoToRecReinvest, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
32820   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub tglMiscSale_KeyDown(KeyCode As Integer, Shift As Integer)
' ** See Note, above.

32900 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscSale_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
32920   intShiftDown = (Shift And acShiftMask) > 0
32930   intAltDown = (Shift And acAltMask) > 0
32940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
32970       With Me
32980         intRetVal = 0
32990         If .cmbRevenueCodes.Enabled = True Then
33000           .cmbRevenueCodes.SetFocus
33010         ElseIf .cmbTaxCodes.Enabled = True Then
33020           .cmbTaxCodes.SetFocus
33030         ElseIf .cmdMiscOK.Enabled = True Then
33040           .cmdMiscOK.SetFocus
33050         ElseIf .cmdMiscCancel.Enabled = True Then
33060           .cmdMiscCancel.SetFocus
33070         ElseIf .cmdMiscClose.Enabled = True Then
33080           .cmdMiscClose.SetFocus
33090         ElseIf .miscAccountNo.Enabled = True Then
33100           .miscAccountNo.SetFocus
33110         Else
33120           .miscTransDate.SetFocus
33130         End If
33140       End With
33150     End Select
33160   End If

        ' ** Shift keys.
33170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33190       With Me
33200         intRetVal = 0
33210         If .miscMemo.Visible = True And .miscMemo.Enabled = True Then
33220           .miscMemo.SetFocus
33230         ElseIf .chkChecks.Enabled = True Then
33240           .chkChecks.SetFocus
33250         Else
33260           .miscDescription.SetFocus
33270         End If
33280       End With
33290     End Select
33300   End If

EXITP:
33310   KeyCode = intRetVal
33320   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub tglMiscSale_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

33400 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscSale_MouseMove"

33410   Tgl_Handler_Sub5 THIS_PROC, blnMiscSale_Focus, blnMiscSale_MouseDown, _
          blnMiscReinvest_Focus, blnMiscReinvest_MouseDown, blnGoToSaleForm, blnGoToRecReinvest, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
33420   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub tglMiscSale_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

33500 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscSale_MouseUp"

33510   Tgl_Handler_Sub5 THIS_PROC, blnMiscSale_Focus, blnMiscSale_MouseDown, _
          blnMiscReinvest_Focus, blnMiscReinvest_MouseDown, blnGoToSaleForm, blnGoToRecReinvest, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
33520   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub tglMiscSale_LostFocus()

33600 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscSale_LostFocus"

33610   Tgl_Handler_Sub5 THIS_PROC, blnMiscSale_Focus, blnMiscSale_MouseDown, _
          blnMiscReinvest_Focus, blnMiscReinvest_MouseDown, blnGoToSaleForm, blnGoToRecReinvest, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
33620   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub tglMiscReinvest_GotFocus()

33700 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscReinvest_GotFocus"

33710   Tgl_Handler_Sub5 THIS_PROC, blnMiscSale_Focus, blnMiscSale_MouseDown, _
          blnMiscReinvest_Focus, blnMiscReinvest_MouseDown, blnGoToSaleForm, blnGoToRecReinvest, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
33720   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub tglMiscReinvest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

33800 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscReinvest_MouseDown"

33810   Tgl_Handler_Sub5 THIS_PROC, blnMiscSale_Focus, blnMiscSale_MouseDown, _
          blnMiscReinvest_Focus, blnMiscReinvest_MouseDown, blnGoToSaleForm, blnGoToRecReinvest, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
33820   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub tglMiscReinvest_Click()
' ** I'm going to try using [posted] as a signal (since it's not used for anything else).

33900 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscReinvest_Click"

33910   With Me
33920   End With

EXITP:
33930   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub tglMiscReinvest_KeyDown(KeyCode As Integer, Shift As Integer)

34000 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscReinvest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

34010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
34020   intShiftDown = (Shift And acShiftMask) > 0
34030   intAltDown = (Shift And acAltMask) > 0
34040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
34050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
34070       With Me
34080         intRetVal = 0
34090         If .cmbRevenueCodes.Enabled = True Then
34100           .cmbRevenueCodes.SetFocus
34110         ElseIf .cmbTaxCodes.Enabled = True Then
34120           .cmbTaxCodes.SetFocus
34130         ElseIf .cmdMiscOK.Enabled = True Then
34140           .cmdMiscOK.SetFocus
34150         ElseIf .cmdMiscCancel.Enabled = True Then
34160           .cmdMiscCancel.SetFocus
34170         ElseIf .cmdMiscClose.Enabled = True Then
34180           .cmdMiscClose.SetFocus
34190         ElseIf .miscAccountNo.Enabled = True Then
34200           .miscAccountNo.SetFocus
34210         Else
34220           .miscTransDate.SetFocus
34230         End If
34240       End With
34250     End Select
34260   End If

        ' ** Shift keys.
34270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
34290       With Me
34300         intRetVal = 0
34310         If .miscMemo.Visible = True And .miscMemo.Enabled = True Then
34320           .miscMemo.SetFocus
34330         ElseIf .chkChecks.Enabled = True Then
34340           .chkChecks.SetFocus
34350         Else
34360           .miscDescription.SetFocus
34370         End If
34380       End With
34390     End Select
34400   End If

EXITP:
34410   KeyCode = intRetVal
34420   Exit Sub

ERRH:
460     intRetVal = 0
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Sub

Private Sub tglMiscReinvest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

34500 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscReinvest_MouseMove"

34510   Tgl_Handler_Sub5 THIS_PROC, blnMiscSale_Focus, blnMiscSale_MouseDown, _
          blnMiscReinvest_Focus, blnMiscReinvest_MouseDown, blnGoToSaleForm, blnGoToRecReinvest, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
34520   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub tglMiscReinvest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

34600 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscReinvest_MouseUp"

34610   Tgl_Handler_Sub5 THIS_PROC, blnMiscSale_Focus, blnMiscSale_MouseDown, _
          blnMiscReinvest_Focus, blnMiscReinvest_MouseDown, blnGoToSaleForm, blnGoToRecReinvest, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
34620   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub tglMiscReinvest_LostFocus()

34700 On Error GoTo ERRH

        Const THIS_PROC As String = "tglMiscReinvest_LostFocus"

34710   Tgl_Handler_Sub5 THIS_PROC, blnMiscSale_Focus, blnMiscSale_MouseDown, _
          blnMiscReinvest_Focus, blnMiscReinvest_MouseDown, blnGoToSaleForm, blnGoToRecReinvest, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
34720   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Private Sub chkChecks_AfterUpdate()

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkChecks_AfterUpdate"

34810   With Me
34820     Select Case .chkChecks
          Case True
34830       .chkChecks_lbl.FontBold = True
34840       .miscMemo.Enabled = True
34850       .miscMemo.Visible = True
34860       .miscMemo_lbl2.Visible = True
34870       .miscMemo.SetFocus
34880     Case False
34890       .chkChecks_lbl.FontBold = False
34900       .miscMemo.Visible = False
34910       .miscMemo.Enabled = False
34920       .miscMemo_lbl2.Visible = False
34930     End Select
34940     ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
34950   End With

EXITP:
34960   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscMemo_KeyDown(KeyCode As Integer, Shift As Integer)

35000 On Error GoTo ERRH

        Const THIS_PROC As String = "miscMemo_KeyDown"

        Dim intLen As Integer

35010   With Me
35020     If IsNull(.miscMemo.text) = False Then
35030       If Trim(.miscMemo.text) <> vbNullString Then
35040         intLen = Len(Trim(.miscMemo.text))
35050         If intLen > MEMO_MAX Then
35060           KeyCode = 0
35070           Beep
35080           .miscMemo.text = Left(Trim(.miscMemo.text), MEMO_MAX)
35090           .miscMemo.SelStart = 99
35100         End If
35110       End If
35120     End If
35130   End With

EXITP:
35140   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub miscMemo_AfterUpdate()

35200 On Error GoTo ERRH

        Const THIS_PROC As String = "miscMemo_AfterUpdate"

35210   With Me
35220     If IsNull(.miscMemo) = False Then
35230       If Trim(.miscMemo) <> vbNullString Then
35240         strMemo_New = Trim(.miscMemo)
35250         If Len(strMemo_New) > MEMO_MAX Then
35260           strMemo_New = Left(strMemo_New, MEMO_MAX)
35270           .miscMemo = strMemo_New
35280         End If
35290       End If
35300     End If
35310   End With

35320   ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
35330   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Enter()

35400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Enter"

35410   With Me
35420     Select Case IsNull(.cmbRevenueCodes)
          Case True
35430       ExpandCombo .cmbRevenueCodes  ' ** Module Procedure: modUtilities.
35440     Case False
35450       If .cmbRevenueCodes = REVID_INC Or .cmbRevenueCodes = REVID_EXP Then
35460         ExpandCombo .cmbRevenueCodes  ' ** Module Procedure: modUtilities.
35470       End If
35480     End Select
35490   End With

EXITP:
35500   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

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

35600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Change"

35610   With Me
35620     If .cmbRevenueCodes <> .cmbRevenueCodes.OldValue Then
35630       ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
35640     End If
35650   End With

EXITP:
35660   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

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

35700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_AfterUpdate"

        Dim strRevCode As String, lngTaxcode As Long

        ' ** Make sure tax code matches selected revenue code.
35710   With Me
35720     If IsNull(.cmbRevenueCodes) = False Then
35730       If .cmbRevenueCodes > 0 Then
35740         If gblnLinkRevTaxCodes = True Then
35750           strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
35760           If IsNull(.cmbTaxCodes) = True Then
35770             If strRevCode = "I" Then
35780               .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
35790             ElseIf strRevCode = "E" Then
35800               .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
35810             End If
35820           Else
35830             If .cmbTaxCodes = 0& Then
35840               If strRevCode = "I" Then
35850                 .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
35860               ElseIf strRevCode = "E" Then
35870                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
35880               End If
35890             Else
35900               lngTaxcode = .cmbTaxCodes
35910               Select Case .miscType
                    Case "Paid"
35920                 If strRevCode = "E" And .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
35930                   .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
35940                 End If
35950               Case "Received"
35960                 If strRevCode = "I" And .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
35970                   .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
35980                 End If
35990               Case "Misc."
36000                 If strRevCode = "I" And .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
                        'MsgBox "This will change..."
36010                   .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
36020                 ElseIf strRevCode = "E" And .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
                        'MsgBox "This will change..."
36030                   .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
36040                 End If
36050               End Select
36060             End If
36070           End If
36080         End If
36090       Else
36100         Select Case .miscType
              Case "Paid"
36110           .cmbRevenueCodes = REVID_EXP
36120         Case "Received"
36130           .cmbRevenueCodes = REVID_INC
36140         Case "Misc."
36150           .cmbRevenueCodes = REVID_INC
36160         End Select
36170       End If
36180     Else
36190       Select Case .miscType
            Case "Paid"
36200         .cmbRevenueCodes = REVID_EXP
36210       Case "Received"
36220         .cmbRevenueCodes = REVID_INC
36230       Case "Misc."
36240         .cmbRevenueCodes = REVID_INC
36250       End Select
36260     End If
36270   End With

EXITP:
36280   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub cmbRevenueCodes_LostFocus()

36300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_LostFocus"

36310   VerifyAccountNo  ' ** Function: Below.

EXITP:
36320   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub cmbTaxCodes_Enter()

36400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Enter"

36410   With Me
36420     Select Case IsNull(.cmbTaxCodes)
          Case True
36430       ExpandCombo .cmbTaxCodes  ' ** Module Procedure: modUtilities.
36440     Case False
36450       If .cmbTaxCodes = TAXID_INC Or .cmbTaxCodes = TAXID_DED Then
36460         ExpandCombo .cmbTaxCodes  ' ** Module Procedure: modUtilities.
36470       End If
36480     End Select
36490   End With

EXITP:
36500   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

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

36600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Change"

36610   With Me
36620     If .cmbTaxCodes <> .cmbTaxCodes.OldValue Then
36630       ChangedMisc_Sub5 True, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.
36640     End If
36650   End With

EXITP:
36660   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

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

36700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_AfterUpdate"

        Dim strRevCode As String, lngTaxcode As Long

        ' ** Make sure revenue code matches selected tax code.
36710   With Me
36720     If IsNull(.cmbTaxCodes) = False Then
36730       If .cmbTaxCodes > 0& Then
36740         If gblnLinkRevTaxCodes = True Then
36750           lngTaxcode = .cmbTaxCodes
36760           Select Case .miscType
                Case "Paid"
                  ' ** EXPENSE.
36770             If IsNull(.cmbRevenueCodes) = False Then
36780               If .cmbRevenueCodes > 0& Then
36790                 strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
36800                 If strRevCode <> "E" Then
36810                   .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
36820                 End If
36830               Else
36840                 .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
36850               End If
36860             Else
36870               .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
36880             End If
36890           Case "Received"
                  ' ** INCOME.
36900             If IsNull(.cmbRevenueCodes) = False Then
36910               If .cmbRevenueCodes > 0& Then
36920                 strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
36930                 If strRevCode <> "I" Then
36940                   .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
36950                 End If
36960               Else
36970                 .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
36980               End If
36990             Else
37000               .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37010             End If
37020           Case "Misc."
                  ' ** ALL.
37030             If IsNull(.cmbRevenueCodes) = False Then
37040               If .cmbRevenueCodes > 0& Then
37050                 strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
37060                 If lngTaxcode > 0& Then
37070                   If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
                          ' ** INCOME.
37080                     If strRevCode = "E" Then
                            'MsgBox "This will change..."
37090                       .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37100                     End If
37110                   ElseIf .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
                          ' ** EXPENSE.
37120                     If strRevCode = "I" Then
                            'MsgBox "This will change..."
37130                       .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
37140                     End If
37150                   End If
37160                 End If
37170               Else
37180                 If lngTaxcode > 0& Then
37190                   If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
37200                     .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37210                   ElseIf .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
37220                     .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
37230                   End If
37240                 End If
37250               End If
37260             Else
37270               If lngTaxcode > 0& Then
37280                 If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
37290                   .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
37300                 ElseIf .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
37310                   .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
37320                 End If
37330               End If
37340             End If
37350           End Select
37360         End If
37370       Else
37380         Select Case .miscType
              Case "Paid"
37390           .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
37400         Case "Received"
37410           .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
37420         Case "Misc."
37430           .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
37440         End Select
37450       End If
37460     Else
37470       Select Case .miscType
            Case "Paid"
37480         .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
37490       Case "Received"
37500         .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
37510       Case "Misc."
37520         .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
37530       End Select
37540     End If
37550   End With

EXITP:
37560   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub cmbTaxCodes_LostFocus()

37600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_LostFocus"

37610   VerifyAccountNo  ' ** Function: Below.

EXITP:
37620   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Public Function VerifyAccountNo() As Boolean

37700 On Error GoTo ERRH

        Const THIS_PROC As String = "VerifyAccountNo"

        Dim strAccountNo As String
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

37710   With Me

37720     blnRetVal = True

37730     If ((gblnMiscValidated = False And gblnMiscChanged = True) And (lngErrCnt < 3&)) Then
37740       Select Case IsNull(.miscAccountNo)
            Case True
37750         blnRetVal = False
37760         lngErrCnt = lngErrCnt + 1&
37770         .miscAccountNo.SetFocus
37780         MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Invalid Entry"
37790         gblnMiscValidated = False
37800       Case False
              ' ** Check to see if it is a valid account.
37810         lngErrCnt = 0&
37820         Select Case gblnAccountNoWithType
              Case True
37830           strAccountNo = Left(.miscAccountNo, 2) & Mid(.miscAccountNo, 3)
37840         Case False
37850           strAccountNo = .miscAccountNo
37860         End Select
37870         varTmp00 = DLookup("[accountno]", "account", "[accountno] = '" & strAccountNo & "' And ([dateclosed] Is Null)")
37880         Select Case IsNull(varTmp00)
              Case True
37890           blnRetVal = False
37900           MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Invalid Entry"
37910           gblnMiscValidated = False
37920         Case False
37930           gblnMiscValidated = True
37940           blnStartTrans = False
37950         End Select
37960       End Select
37970     Else
37980       lngErrCnt = 0&
37990     End If
38000   End With

EXITP:
38010   VerifyAccountNo = blnRetVal
38020   Exit Function

ERRH:
460     blnRetVal = False
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
500     Resume EXITP

End Function

Public Sub Map_NewRec()

38100 On Error GoTo ERRH

        Const THIS_PROC As String = "Map_NewRec"

38110   With Me
38120 On Error Resume Next
38130     DoCmd.RunCommand acCmdRecordsGoToNew
38140     If ERR.Number <> 0 Then
38150       blnMapNewRec = True
38160       .TimerInterval = 500&
38170     End If
38180 On Error GoTo ERRH
38190   End With

EXITP:
38200   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

38300 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

38310   Select Case intWhere
        Case acCmdRecordsGoToFirst
38320     DoCmd.RunCommand acCmdRecordsGoToFirst
38330   Case acCmdRecordsGoToLast
38340     DoCmd.RunCommand acCmdRecordsGoToLast
38350   Case acCmdRecordsGoToPrevious
38360     DoCmd.RunCommand acCmdRecordsGoToPrevious
38370   Case acCmdRecordsGoToNext
38380     DoCmd.RunCommand acCmdRecordsGoToNext
38390   Case acCmdRecordsGoToNew
38400     DoCmd.RunCommand acCmdRecordsGoToNew
38410   Case Else
38420     If IsMissing(varID) = False Then
38430       With Me
38440         Set rst = .RecordsetClone
38450         With rst
38460           .FindFirst "[miscID] = " & CStr(varID)
38470           If .NoMatch = False Then
38480             Me.Bookmark = .Bookmark
38490           End If
38500           .Close
38510         End With
38520       End With
38530     End If
38540   End Select

EXITP:
38550   Set rst = Nothing
38560   Exit Sub

ERRH:
460     Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
470     Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
480     Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub IncludeCurrency(blnShow As Boolean)

38600 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

38610   IncludeCurrencyMisc_Sub5 blnShow, lngTpp, lngICash_Left, lngPCash_Left, lngCurrID_Left, Me  ' ** Module Procedure: modJrnlSub5MiscFuncs.

EXITP:
38620   Exit Sub

ERRH:
460     THAT_PROC = THIS_PROC
470     That_Erl = Erl
480     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
490     Resume EXITP

End Sub
