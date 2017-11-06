VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournal_Sub1_Dividend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournal_Sub1_Dividend"

'VGC 10/01/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Map:              A {cmdDividendMap}
' **   Close:            C {cmdDividendClose}
' **   Dividend:         D {opgJournal_optDividend on frmJournal}
' **   Cancel:           E {cmdDividendCancel}
' **   Show All Assets:  H {chkShowAllAssets}
' **   Interest:         I {opgJournal_optInterest on frmJournal}
' **   Miscellaneous:    M {opgJournal_optMisc on frmJournal}
' **   Inc/Exp Codes:    N {cmbRevenueCodes}
' **   OK:               O {cmdDividendOK}
' **   Purchase:         P {opgJournal_optPurchase on frmJournal}
' **   Sale:             S {opgJournal_optSale on frmJournal}
' **   Tax Codes:        T {cmbTaxCodes}
' **   Exit:             X {cmdDividendClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Next Record:      N {MoveRec}
' **   Previous Record:  P {MoveRec}
' **   Switch:           T {cmdSwitch on frmJournal}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Mark Reinvest:    R {tglDividendReinvest}
' **   Close Timer:      T {Form_KeyDown}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}

' ** Dividends are always INCOME for Tax Codes.

' ** IsAverage used for chkShowAllAssets.

'QRY: qryJournal_Dividend_02             REF: FormRef('frmJournal_Sub1_Dividend')
'QRY: qryStatementParameters_Dividend_01 REF: FormRef('frmJournal_Sub1_Dividend')

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

' ** Combo box column constants: dividendAssetNo.
'Private Const CBX_A_DESC   As Integer = 0  'totdesc
'Private Const CBX_A_CUSIP  As Integer = 1  'cusip
Private Const CBX_A_ASTNO  As Integer = 2  'assetno
'Private Const CBX_A_TYPE   As Integer = 3  'assettype
Private Const CBX_A_TAX    As Integer = 4  'taxcode
Private Const CBX_A_CURRID As Integer = 5  'curr_id

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

' ** Combo box column constants: dividendCurr_ID.
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
Private blnDividendReinvest_Focus As Boolean, blnDividendReinvest_MouseDown As Boolean

Private lngCurrID As Long
Private blnCmdButton As Boolean, blnGoToDivReinvest As Boolean
Private blnAccountNoErr As Boolean, blnStartTrans As Boolean, blnMapNewRec As Boolean
Private datPostingDate As Date, lngDateCnt As Long, lngErrCnt As Long
Private blnCheckingAccount As Boolean, blnAccountValidated As Boolean, blnExitingAssetNo As Boolean
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
180           blnAccountNoErr = False: blnStartTrans = False
190           gblnSwitchTo = False: blnMapNewRec = False

200           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
210           blnDividendReinvest_Focus = False: blnDividendReinvest_MouseDown = False

220           If .cmbRevenueCodes.DefaultValue = vbNullString Then
230             .cmbRevenueCodes.DefaultValue = REVID_ORDDIV  ' ** Ordinary Dividend.
240           End If

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
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub Form_Load()
' ** Create an instance of the Calendar class.
' ** Copyright Lebans Holdings 1999 Ltd.

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me

520       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
530       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
540       clsMonthClass.PositionAtCursor = True

550       blnCheckingAccount = False   ' ** Default.
560       blnExitingAssetNo = False    ' ** Default.
570       blnAccountValidated = False  ' ** Default.
580       blnGoToDivReinvest = False  ' ** Set the switch to false until they push the button.
590       blnCmdButton = False

600       datPostingDate = GetPostDate  ' ** Module Function: modUtilities.

          ' ** Set the form's default JournalType.
610       .dividendType.DefaultValue = Chr(34) & "Dividend" & Chr(34)

          ' ** Do this whether or not Currency is visible.
620       CLR_DISABLED_FG = CLR_DKGRY
630       CLR_DISABLED_BG = CLR_LTTEAL

640       .dividendCurr_ID.Enabled = False
650       .dividendCurr_ID.Locked = True
660       .dividendCurr_ID.ForeColor = CLR_DISABLED_FG
670       .dividendCurr_ID.BackColor = CLR_VLTGRN
680       dividendCurr_ID_AfterUpdate  ' ** Procedure: Below.

          ' ** Set the default Posting Date (transdate), entered by the user on opening.
690       .dividendTransDate.DefaultValue = "#" & CDate(Format(datPostingDate, "mm/dd/yyyy")) & "#"

700       gstrFormQuerySpec = .Parent.Name  ' ** Make sure this is set for the assetno combo box.

710       gblnDividendValidated = False
720       gblnDividendChanged = False

730       .NavigationButtons = True

740       With .Parent
750         .opgJournal.Enabled = True
760         .cmdSwitch.Enabled = True
770         .cmdSwitch_raised_img.Visible = True
780         .cmdSwitch_raised_img_dis.Visible = False
790         .cmdSwitch_raised_semifocus_dots_img.Visible = False
800         .cmdSwitch_raised_focus_img.Visible = False
810         .cmdSwitch_raised_focus_dots_img.Visible = False
820         .cmdSwitch_sunken_focus_dots_img.Visible = False
830       End With
840       .cmdDividendClose.Enabled = True
850       .cmdDividendOK.Enabled = False
860       .cmdDividendCancel.Enabled = False
870       .Parent.NavVis True  ' ** Form Procedure: frmJournal.

880       If .NewRecord = True Then
890         .tglDividendReinvest.Enabled = False
900         .tglDividendReinvest_false_raised_img_dis.Visible = True
910         .tglDividendReinvest_false_raised_img.Visible = False
920         .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = False
930         .tglDividendReinvest_false_raised_focus_img.Visible = False
940         .tglDividendReinvest_false_raised_focus_dots_img.Visible = False
950         .tglDividendReinvest_false_sunken_focus_dots_img.Visible = False
960         .tglDividendReinvest_true_raised_img.Visible = False
970         .tglDividendReinvest_true_raised_focus_img.Visible = False
980         .tglDividendReinvest_true_raised_focus_dots_img.Visible = False
990         .tglDividendReinvest_true_sunken_focus_dots_img.Visible = False
1000        .tglDividendReinvest_true_raised_img_dis.Visible = False
1010      End If

1020      Select Case gblnIncomeTaxCoding
          Case True
1030        .cmbTaxCodes.Visible = True
1040        .cmbTaxCodes_lbl2.Visible = False
1050      Case False
1060        .cmbTaxCodes.Visible = False
1070        .cmbTaxCodes_lbl2.Visible = True
1080      End Select

1090      Select Case gblnRevenueExpenseTracking
          Case True
1100        .cmbRevenueCodes.Visible = True
1110        .cmbRevenueCodes_lbl2.Visible = False
1120      Case False
1130        .cmbRevenueCodes.Visible = False
1140        .cmbRevenueCodes_lbl2.Visible = True
1150      End Select
1160      .Refresh

1170    End With

EXITP:
1180    Exit Sub

ERRH:
1190    THAT_PROC = THIS_PROC
1200    That_Erl = Erl
1210    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1220    Resume EXITP

End Sub

Private Sub Form_Current()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim strAccountNo As String
        Dim intPos01 As Integer
        Dim blnContinue As Boolean

1310    With Me

1320      blnContinue = True

1330      DoCmd.SelectObject acForm, .Parent.Name, False

1340      blnCheckingAccount = False   ' ** Default.
1350      blnExitingAssetNo = False    ' ** Default.
1360      blnAccountValidated = False  ' ** Default.
1370      blnStartTrans = False
1380      lngErrCnt = 0&

1390      .cmbAccountHelper = Null

1400      If .dividendCurr_ID.Visible = True Then
1410        If .dividendCurr_Date.Visible = True Then .dividendCurr_Date.Visible = False
1420        .dividendICash_usd = Null
1430        .dividendICash_usd.Visible = False
1440        .dividendICash.Format = "Currency"
1450        .dividendICash.DecimalPlaces = 2
1460        .dividendICash.BackColor = CLR_WHT
1470        .dividendPerShare_usd = Null
1480        .dividendPerShare_usd.Visible = False
1490        .dividendPerShare.Format = "Currency"
1500        .dividendPerShare.DecimalPlaces = 4
1510        .dividendPerShare.BackColor = CLR_WHT
1520        If lngNoChars = 0& Or IsEmpty(arr_varNoChar) Then
1530          arr_varNoChar = .Parent.NoChar_Get  ' ** Form Function: frmJournal.
1540          lngNoChars = UBound(arr_varNoChar, 2) + 1&
1550        End If
1560      End If

1570      gstrFormQuerySpec = .Parent.Name  ' ** Make sure this is set for the assetno combo box.

1580      Select Case IsNull(.dividendAccountNo_Data)
          Case True
1590        .dividendAccountNo = vbNullString
1600        .dividendAccountNo.Enabled = True
1610        .dividendAccountNo.BorderColor = CLR_LTBLU2
1620        .dividendAccountNo.BackStyle = acBackStyleNormal
1630        .dividendAccountNo_lbl.BackStyle = acBackStyleNormal
1640        .dividendAccountNo_lbl_box.Visible = False
1650        .cmdLock.Enabled = False
1660        .cmdLock_open_raised_img_dis.Visible = True
1670        .cmdLock_open_raised_img.Visible = False
1680        .cmdLock_closed_raised_img.Visible = False
1690        .cmbAccountHelper.Enabled = True
1700        .cmbAccountHelper.BorderColor = CLR_LTBLU2
1710        .cmbAccountHelper.BackStyle = acBackStyleNormal
1720        .dividendAccountNo.SetFocus
1730        .cmdDividendMap.Enabled = True
1740        blnAccountNoErr = False
1750      Case False
1760        .dividendAccountNo = .dividendAccountNo_Data
1770        gstrDividendAccountNumber = .dividendAccountNo
1780        .dividendTransDate.SetFocus  ' ** Make sure it's not on AccountNo before disabling.
1790        .dividendAccountNo.Enabled = False
1800        .dividendAccountNo.BorderColor = WIN_CLR_DISR
1810        .dividendAccountNo.BackStyle = acBackStyleTransparent
1820        .dividendAccountNo_lbl.BackStyle = acBackStyleTransparent
1830        .dividendAccountNo_lbl_box.Visible = True
1840        .cmdLock.Enabled = True
1850        .cmdLock_open_raised_img_dis.Visible = False
1860        .cmdLock_open_raised_img.Visible = False
1870        .cmdLock_closed_raised_img.Visible = True
1880        .cmbAccountHelper.Enabled = False
1890        .cmbAccountHelper.BorderColor = WIN_CLR_DISR
1900        .cmbAccountHelper.BackStyle = acBackStyleTransparent
1910        .cmdDividendMap.Enabled = False
1920      End Select

          ' ** Make sure validation is reset.
1930      gblnDividendValidated = False

1940      Select Case IsNull(.dividendShareFace)
          Case True
1950        intPos01 = 0
1960      Case False
1970        intPos01 = InStr(1, .dividendShareFace, ".", vbTextCompare)
1980      End Select
1990      If intPos01 > 0 Then
2000        If Len(.dividendShareFace) - intPos01 > gintShareFaceDecimals Then
2010          blnContinue = False
2020          MsgBox "You can only enter a share/face with up to " & Trim(str(gintShareFaceDecimals)) & " decimals.", _
                vbInformation + vbOKOnly, "Invalid Entry"
2030          .dividendShareFace = vbNullString
2040          DoCmd.CancelEvent
2050        Else
2060          .dividendShareFace.Format = "#,###." & String(gintShareFaceDecimals, "0")
2070          .Refresh
2080        End If
2090      Else
2100        .dividendShareFace.Format = "#,###"
2110        .Refresh
2120      End If

2130      If blnContinue = True Then

2140        lngCurrID = .dividendCurr_ID.Column(CBX_C_CURRID)

2150        If IsNull(.dividendAssetNo) = False Then
2160          .dividendAssetNo.Requery
2170        End If
2180        gstrDividendAsset = IIf(IsNull(.dividendAssetNo), vbNullString, .dividendAssetNo)
2190        gstrDividendShareFace = IIf(IsNull(.dividendShareFace), vbNullString, .dividendShareFace)
2200        gstrDividendAccountNumber = IIf(IsNull(.dividendAccountNo), vbNullString, .dividendAccountNo)
2210        gstrDividendPerShare = IIf(IsNull(.dividendPerShare), vbNullString, .dividendPerShare)
2220        gstrDividendICash = IIf(IsNull(.dividendICash), vbNullString, .dividendICash)
2230        gstrDividendICash = Rem_Dollar(gstrDividendICash, lngCurrID)  ' ** Module Function: modStringFuncs.

2240        strAccountNo = .dividendAccountNo

2250        Select Case .NewRecord
            Case True
2260          .cmdDividendCancel.Enabled = False
2270        Case False
2280          .cmdDividendCancel.Enabled = True
2290        End Select

2300        If strAccountNo <> vbNullString Then
              ' ** I'm using this to indicate chkShowAllAssets.
2310          Select Case .IsAverage
              Case True
2320            If .dividendAssetNo.RowSource <> "qryJournal_Dividend_02g" Then
                  ' ** qryJournal_Dividend_02c (MasterAsset table, with totdesc)
                  ' ** linked to AssetType, Dividend = True, all assets.
2330              .dividendAssetNo.RowSource = "qryJournal_Dividend_02g"
2340            End If
2350          Case False
                ' ** Queries do not have Dividend = True criteria.
2360            If .dividendAccountNo = "INCOME O/U" Or .dividendAccountNo = "99-INCOME O/U" Then
2370              If .dividendAssetNo.RowSource <> "qryJournal_Dividend_02f" Then
                    ' ** qryJournal_Dividend_02b (Journal linked to Account, MasterAsset)
                    ' ** linked to AssetType, by specified FormRef('frmJournal_Sub1_Dividend').
2380                .dividendAssetNo.RowSource = "qryJournal_Dividend_02f"
2390              End If
2400            Else
2410              If .dividendAssetNo.RowSource <> "qryJournal_Dividend_02d" Then
                    ' ** qryJournal_Dividend_02a (MasterAsset table linked to ActiveAssets, Account)
                    ' ** linked to AssetType, by specified FormRef('frmJournal_Sub1_Dividend').
2420                .dividendAssetNo.RowSource = "qryJournal_Dividend_02d"
2430              End If
2440            End If
2450          End Select
2460          .dividendAssetNo.Requery
2470        End If

2480        .Refresh

2490      End If  ' ** blnContinue.

2500      If blnCmdButton = False Then  ' ** That is, only run this if Form_Current() wasn't called from another button, like cmdMiscOK_Click().
2510        blnGoToDivReinvest = False
2520        .tglDividendReinvest_false_raised_img_dis.Visible = True
2530        .tglDividendReinvest_false_raised_img.Visible = False
2540        .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = False
2550        .tglDividendReinvest_false_raised_focus_img.Visible = False
2560        .tglDividendReinvest_false_raised_focus_dots_img.Visible = False
2570        .tglDividendReinvest_false_sunken_focus_dots_img.Visible = False
2580        .tglDividendReinvest_true_raised_img.Visible = False
2590        .tglDividendReinvest_true_raised_focus_img.Visible = False
2600        .tglDividendReinvest_true_raised_focus_dots_img.Visible = False
2610        .tglDividendReinvest_true_sunken_focus_dots_img.Visible = False
2620        .tglDividendReinvest_true_raised_img_dis.Visible = False
2630        .tglDividendReinvest.Enabled = False
2640        Select Case .posted
            Case True
2650          .tglDividendReinvest.Enabled = False
2660          .tglDividendReinvest_true_raised_img_dis.Visible = True
2670          .tglDividendReinvest_false_raised_img_dis.Visible = False
2680        Case False
              ' ** Let's wait before enabling this.
2690          If .NewRecord = False Then
2700            .tglDividendReinvest.Enabled = True
2710            .tglDividendReinvest_false_raised_img.Visible = True
2720            .tglDividendReinvest_false_raised_img_dis.Visible = False
2730          End If
2740        End Select
2750      End If  ' ** blnCmdButton.

2760      .Repaint

2770      DoEvents

          ' ** Set the currency symbol.
2780      dividendCurr_ID_AfterUpdate  ' ** Procedure: Below.

2790      If .dividendCurr_ID.Visible = True Then
2800        dividendICash_AfterUpdate  ' ** Procedure: Below.
2810      End If

2820    End With

        'HERE! 1  Form_Current()  frmJournal_Sub1_Dividend
        'HERE! 2  Form_Current()  frmJournal_Sub1_Dividend
        'HERE! 3  Form_Open()  frmJournal
        'HERE! 4  Form_Load()  frmJournal
        'HERE! 5  Form_Load()  frmJournal
        'HERE! 6  ShowDividend()  frmJournal
        'HERE! 7  Form_Open()  frmJournal
        'HERE! 8  Form_Timer()  frmJournal
        'HERE! 6  ShowDividend()  frmJournal
        'HERE! 9  Form_Timer()  frmJournal
        'HERE! 8  Form_Timer()  frmJournal

EXITP:
2830    Exit Sub

ERRH:
2840    THAT_PROC = THIS_PROC
2850    That_Erl = Erl
2860    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2870    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2920    intShiftDown = (Shift And acShiftMask) > 0
2930    intAltDown = (Shift And acAltMask) > 0
2940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Next Record:      N {MoveRec}
        ' **   Previous Record:  P {MoveRec}
        ' **   Switch:           T {cmdSwitch on frmJournal}

        ' ** Shortcut Ctrl-Shift to other forms and functionality:
        ' **   Mark Reinvest:    R {tglDividendReinvest}
        ' **   Close Timer:      T {Form_KeyDown}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}

        ' ** Plain keys.
2950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2960      Select Case intRetVal
          Case vbKeyEscape
2970        intRetVal = 0
2980        If Forms(strCallingForm).opgJournal.Enabled = True Then
2990          Select Case gblnDividendChanged
              Case True
3000            cmdDividendCancel_Click  ' ** Procedure: Below.
3010          Case False
3020            cmdDividendClose_Click  ' ** Procedure: Below.
3030          End Select
3040        Else
3050          cmdDividendCancel_Click  ' ** Procedure: Below.
3060        End If
3070      Case vbKeyDelete
3080        gblnDividendChanged = True
3090      Case vbKeyF7
3100        intRetVal = 0
3110        cmdCalendar1_Click  ' ** Procedure: Below.
3120      End Select
3130    End If

        ' ** Alt keys.
3140    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
3150      Select Case intRetVal
          Case vbKeyC
3160        With Me
3170          intRetVal = 0
3180          If .cmdDividendClose.Enabled = True Then
3190            .cmdDividendClose.SetFocus
3200            cmdDividendClose_Click  ' ** Procedure: Below.
3210          Else
3220            Beep
3230          End If
3240        End With
3250      Case vbKeyE
3260        With Me
3270          intRetVal = 0
3280          If .cmdDividendCancel.Enabled = True Then
3290            .cmdDividendCancel.SetFocus
3300            cmdDividendCancel_Click  ' ** Procedure: Below.
3310          Else
3320            Beep
3330          End If
3340        End With
3350      Case vbKeyX
3360        With Me
3370          intRetVal = 0
3380          If .cmdDividendClose.Enabled = True Then
3390            cmdDividendClose_Click  ' ** Procedure: Below.
3400          End If
3410        End With
3420      End Select
3430    End If

        ' ** Ctrl keys.
3440    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3450      Select Case intRetVal
          Case vbKeyN
3460        With Me
3470          intRetVal = 0
3480          MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3490        End With
3500      Case vbKeyP
3510        With Me
3520          intRetVal = 0
3530          MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3540        End With
3550      Case vbKeyT
3560        With Me
3570          intRetVal = 0
3580          DoCmd.SelectObject acForm, .Parent.Name, False
3590          .Parent.cmdSwitch.SetFocus
3600          DoEvents
3610          .Parent.cmdSwitch_Click  ' ** Form Procedure: frmJournal.
3620        End With
3630      End Select
3640    End If

        ' ** Ctrl-Shift keys.
3650    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3660      Select Case intRetVal
          Case vbKeyF
3670        With Me
3680          intRetVal = 0
3690          .FocusHolder.SetFocus
3700        End With
3710      Case vbKeyR
3720        With Me
3730          intRetVal = 0
3740          .tglDividendReinvest.SetFocus
3750          tglDividendReinvest_Click  ' ** Procedure: Below.
3760        End With
3770      Case vbKeyT
3780        intRetVal = 0
3790        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
3800      End Select
3810    End If

EXITP:
3820    KeyCode = intRetVal
3830    Exit Sub

ERRH:
3840    intRetVal = 0
3850    THAT_PROC = THIS_PROC
3860    That_Erl = Erl
3870    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
3880    Resume EXITP

End Sub

Private Sub Form_GotFocus()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_GotFocus"

3910    With Me
3920      If .dividendAccountNo.Enabled = True Then
3930        .dividendAccountNo.SetFocus
3940      Else
3950        .dividendTransDate.SetFocus
3960      End If
3970    End With

EXITP:
3980    Exit Sub

ERRH:
3990    THAT_PROC = THIS_PROC
4000    That_Erl = Erl
4010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
4020    Resume EXITP

End Sub

Private Sub Form_Timer()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

4110    With Me
4120      .TimerInterval = 0&
          ' ** Trying to get it to go to a new record after a map reinvest.
4130      If blnMapNewRec = True Then
4140        blnMapNewRec = False
4150        If IsNull(.dividendID) = False Then
4160          MoveRec acCmdRecordsGoToNew  ' ** Procedure: Below.
4170        End If
4180        DoEvents
4190      End If  ' ** blnMapNewRec.
4200    End With

EXITP:
4210    Exit Sub

ERRH:
4220    Select Case ERR.Number
        Case Else
4230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4240    End Select
4250    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strControl As String
        Dim blnMsgGiven As Boolean

4310    blnMsgGiven = False

4320    If DataErr <> 0 Then
4330      Select Case DataErr
          Case 13, 2108, 2110, 2116, 2135, 2169, 2474, 2501, 3020, 3163, 3314, 3315, 7753, 8519
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
4340        Response = acDataErrContinue
4350      Case 2046  ' ** The command or action isn't available now (first or last record).
            ' ** Do nothing.
4360        Response = acDataErrContinue
4370      Case 2113  ' ** The value you entered isn't valid for this field.
4380  On Error Resume Next
4390        strControl = Screen.ActiveControl.Name
4400  On Error GoTo ERRH
4410        Select Case strControl
            Case "dividendICash"
4420          strControl = "Income Cash"
4430        Case "dividendShareFace"
4440          strControl = "Shares"
4450        Case "dividendPerShare"
4460          strControl = "Per Share"
4470        Case "dividendTransDate"
4480          strControl = "Date"
4490          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
4500          blnMsgGiven = True
4510        Case Else
4520          strControl = vbNullString
4530        End Select
4540        If blnMsgGiven = False Then
4550          MsgBox "You must enter a numeric value in the " & strControl & " field.", vbInformation + vbOKOnly, "Entry Required"
4560        End If
4570        Response = acDataErrContinue
4580        DoCmd.CancelEvent
4590      Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
            ' ** Ignore.
4600        Response = acDataErrContinue
4610      Case 2237  ' ** The text you entered isn't an item in the list.
4620        Response = acDataErrContinue
4630        DoCmd.CancelEvent
4640      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
4650        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
4660        Response = acDataErrContinue
4670        DoCmd.CancelEvent
4680      Case 3101  ' ** The Microsoft Jet database engine cannot find a record in the table '|' with key matching field(s) '|'.
4690        If IsNull(Me.dividendAccountNo) Or Me.dividendAccountNo = vbNullString Then
4700          MsgBox "The posting of this transaction is incomplete." & vbCrLf & vbCrLf & _
                "You must either complete or cancel this transaction to continue.", vbInformation + vbOKOnly, "Incomplete Data"
4710        End If
4720        If Me.dividendAccountNo.Enabled = True Then
4730          Me.dividendAccountNo.SetFocus
4740        Else
4750          Me.dividendTransDate.SetFocus
4760        End If
4770        Response = acDataErrContinue
4780      Case 3162  ' ** You tried to assign the Null value to a variable that is not a Variant data type.
4790  On Error Resume Next
4800        strControl = Screen.ActiveControl.Name
4810  On Error GoTo ERRH
4820        MsgBox "The " & strControl & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
4830        Response = acDataErrContinue
4840      Case Else
4850        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
4860        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
4870      End Select
4880    End If

        ' ** AcDataError enumeration:
        ' **   0  acDataErrContinue  Ignore the error and continue without displaying the default Microsoft Access
        ' **                         error message. A custom error message may be displayed in place of the default
        ' **                         error message.
        ' **   1  acDataErrDisplay   Display the default Microsoft Access error message. (Default)
        ' **   2  acDataErrAdded     Don't display the default Microsoft Access error message. The entry may be
        ' **                         added to the combo box list in the NotInList event procedure. After the entry
        ' **                         is added, Microsoft Access updates the list by requerying the combo box.
        ' **                         Microsoft Access then rechecks the string against the combo box list, and saves
        ' **                         the value in the NewData argument in the field the combo box is bound to. If
        ' **                         the string is not in the list, then Microsoft Access displays an error message.

EXITP:
4890    Exit Sub

ERRH:
4900    Select Case ERR.Number
        Case Else
4910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4920    End Select
4930    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.
' ** VGC 04/13/08: Calendar window seems to be Modal, preventing the
' **               user from doing anything else till it's closed.

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

5010    If Not clsMonthClass Is Nothing Then
5020      If clsMonthClass.IsCalendar Then
5030        Cancel = -1
5040      Else
5050        Set clsMonthClass = Nothing
5060      End If
5070    End If

EXITP:
5080    Exit Sub

ERRH:
5090    Select Case ERR.Number
        Case Else
5100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5110    End Select
5120    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

5210    With Me
5220      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
5230        Select Case blnCalendar1_Focus
            Case True
5240          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
5250          .cmdCalendar1_raised_img.Visible = False
5260        Case False
5270          .cmdCalendar1_raised_img.Visible = True
5280          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5290        End Select
5300        .cmdCalendar1_raised_focus_img.Visible = False
5310        .cmdCalendar1_raised_focus_dots_img.Visible = False
5320        .cmdCalendar1_sunken_focus_dots_img.Visible = False
5330        .cmdCalendar1_raised_img_dis.Visible = False
5340      End If
5350      If .tglDividendReinvest_true_raised_focus_img.Visible = True Or .tglDividendReinvest_true_raised_focus_dots_img.Visible = True Or _
              .tglDividendReinvest_false_raised_focus_img.Visible = True Or .tglDividendReinvest_false_raised_focus_dots_img.Visible = True Then
5360        Select Case .posted
            Case True
5370          Select Case blnDividendReinvest_Focus
              Case True
5380            .tglDividendReinvest_true_raised_focus_dots_img.Visible = True  ' ** Same for ON.
5390            .tglDividendReinvest_true_raised_img.Visible = False
5400          Case False
5410            .tglDividendReinvest_true_raised_img.Visible = True
5420            .tglDividendReinvest_true_raised_focus_dots_img.Visible = False
5430          End Select
5440          .tglDividendReinvest_false_raised_img.Visible = False
5450          .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = False
5460        Case False
5470          Select Case blnDividendReinvest_Focus
              Case True
5480            .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = True
5490            .tglDividendReinvest_false_raised_img.Visible = False
5500          Case False
5510            .tglDividendReinvest_false_raised_img.Visible = True
5520            .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = False
5530          End Select
5540          .tglDividendReinvest_true_raised_img.Visible = False
5550          .tglDividendReinvest_true_raised_focus_dots_img.Visible = False
5560        End Select
5570        .tglDividendReinvest_false_raised_focus_img.Visible = False
5580        .tglDividendReinvest_false_raised_focus_dots_img.Visible = False
5590        .tglDividendReinvest_false_sunken_focus_dots_img.Visible = False
5600        .tglDividendReinvest_false_raised_img_dis.Visible = False
5610        .tglDividendReinvest_true_raised_focus_img.Visible = False
5620        .tglDividendReinvest_true_sunken_focus_dots_img.Visible = False
5630        .tglDividendReinvest_true_raised_img_dis.Visible = False
5640      End If
5650    End With

EXITP:
5660    Exit Sub

ERRH:
5670    Select Case ERR.Number
        Case Else
5680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5690    End Select
5700    Resume EXITP

End Sub

Public Sub cmdDividendClose_Click()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDividendClose_Click"

        Dim strDocName As String

5810    glngJournalForm = 0&
5820    gstrFormQuerySpec = vbNullString
5830    gblnSetFocus = True
5840    strDocName = "frmJournal"
5850    DoCmd.Close acForm, strDocName

EXITP:
5860    Exit Sub

ERRH:
5870    THAT_PROC = THIS_PROC
5880    That_Erl = Erl
5890    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5900    Resume EXITP

End Sub

Private Sub cmdDividendOK_Click()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDividendOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngThisJrnlID As Long, lngJrnlMS_ID As Long, lngThisCurrID As Long, datTransDate As Date
        Dim strShortName As String, dblICash As Double
        Dim strJrnlUser As String, strMsg As String, strDocName As String
        Dim lngRecs As Long
        Dim blnTmpGoToDivReinvest As Boolean
        Dim lngX As Long
        Dim blnContinue As Boolean

6010    blnContinue = True

6020    DoCmd.Hourglass True
6030    DoEvents

6040    blnCmdButton = True
6050    blnTmpGoToDivReinvest = blnGoToDivReinvest
6060    Form_Current  ' ** Procedure: Above.
6070    blnCmdButton = False
6080    VerifyAccountNo  ' ** Function: Below.

6090    With Me

6100      If gstrDividendAccountNumber = vbNullString Then
6110        blnContinue = False
6120        DoCmd.Hourglass False
6130        MsgBox "Please enter a valid account number.", vbInformation + vbOKOnly, "Entry Required"
6140        dividendAccountNo.SetFocus
6150      Else
6160        If gstrDividendAsset = vbNullString Or gstrDividendAsset = "0" Then
6170          blnContinue = False
6180          DoCmd.Hourglass False
6190          MsgBox "An asset must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
6200          dividendAssetNo.SetFocus
6210        Else
6220          If gstrDividendShareFace = vbNullString Or Val(gstrDividendShareFace) = 0 Then
6230            blnContinue = False
6240            DoCmd.Hourglass False
6250            MsgBox "Shares must be greater than zero.", vbInformation + vbOKOnly, "Entry Required"
6260            dividendShareFace.SetFocus
6270          Else
6280            If gstrDividendICash = vbNullString Or Val(gstrDividendICash) = 0 Then
6290              blnContinue = False
6300              DoCmd.Hourglass False
6310              MsgBox "Income cash must be greater than zero.", vbInformation + vbOKOnly, "Entry Required"
6320              dividendICash.SetFocus
6330            Else
6340              If CDbl(gstrDividendICash) <= 0 Then
6350                blnContinue = False
6360                DoCmd.Hourglass False
6370                MsgBox "Income cash must be greater than zero.", vbInformation + vbOKOnly, "Entry Required"
6380                .dividendICash.SetFocus
6390              Else
6400                .dividendDescription.SetFocus
6410                If InStr(.dividendDescription.text, Chr(34)) > 0 Then
6420                  blnContinue = False
6430                  DoCmd.Hourglass False
6440                  MsgBox "Cannot have quote marks in text.", vbInformation + vbOKOnly, "Invalid Characters"
6450                Else
6460                  If IsNull(.dividendTransDate) = True Then
6470                    blnContinue = False
6480                    DoCmd.Hourglass False
6490                    MsgBox "A Posting Date is required.", vbInformation + vbOKOnly, "Entry Required"
6500                    .dividendTransDate.SetFocus
6510                  Else
6520                    If .dividendTransDate > DateAdd("m", 1, Date) Then
6530                      blnContinue = False
6540                      DoCmd.Hourglass False
6550                      MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
6560                      .dividendTransDate.SetFocus
6570                    End If
6580                  End If
6590                End If
6600              End If
6610            End If
6620          End If
6630        End If
6640      End If

6650      If blnContinue = True Then

6660        DoCmd.Hourglass True  ' ** Make sure it's still running.
6670        DoEvents

6680        If Trim(Nz(.dividendUser, vbNullString)) = vbNullString Then
6690          .dividendUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
6700        End If
6710        If IsNull(.cmbTaxCodes) = True Then
6720          .cmbTaxCodes = 0&
6730        Else
6740          If gblnLinkRevTaxCodes = True Then
6750            If .cmbTaxCodes = 0& Then
6760              .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
6770            Else
6780              If .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
6790                .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
6800              End If
6810            End If
6820          End If
6830        End If
6840        If Nz(.cmbRevenueCodes, 0&) = 0& Then
6850          .cmbRevenueCodes = REVID_ORDDIV  ' ** Ordinary Dividend.
6860        End If

6870        If blnTmpGoToDivReinvest = True Then
6880          .dividendReinvested = True
6890        End If

6900        DoCmd.RunCommand acCmdSaveRecord
6910        DoEvents

6920        lngThisJrnlID = .dividendID
6930        gstrAccountNo = .dividendAccountNo.Value
6940        glngAssetNo = .dividendAssetNo.Column(CBX_A_ASTNO)
6950        strShortName = .dividendShortName.Value
6960        datTransDate = .dividendTransDate.Value
6970        dblICash = Nz(.dividendICash, 0)
6980        lngThisCurrID = .dividendCurr_ID.Column(CBX_C_CURRID)

6990        Set dbs = CurrentDb
            ' ** qryJournal_Dividend_01 (Journal, linked to account, for 'Dividend'),
            ' ** check for existing Dividends, by specified [actno], [astno], [jrnlid].
7000        Set qdf = dbs.QueryDefs("qryJournal_Dividend_04")
7010        With qdf.Parameters
7020          ![jrnlid] = lngThisJrnlID
7030          ![actno] = gstrAccountNo
7040          ![astno] = glngAssetNo
7050        End With
7060        Set rst = qdf.OpenRecordset()
7070        If rst.BOF = True And rst.EOF = True Then
              ' ** No records; all's well.
7080          If .dividendCurr_Date.Visible = True Then .dividendCurr_Date.Visible = False
7090          .Requery
7100          DoCmd.GoToRecord , , acNewRec
7110        Else
7120          strMsg = vbNullString
7130          rst.MoveLast
7140          lngRecs = rst.RecordCount
7150          rst.MoveFirst
7160          If lngRecs > 1& Then
7170            For lngX = 1& To lngRecs
7180              If IsNull(rst![journal_USER]) = True Then
7190                strJrnlUser = "{unknown}"
7200              Else
7210                If Trim(rst![journal_USER]) = vbNullString Then
7220                  strJrnlUser = "{unknown}"
7230                Else
7240                  strJrnlUser = rst![journal_USER]
7250                End If
7260              End If
7270              strMsg = strMsg & strJrnlUser & ", "
7280              If lngX < lngRecs Then rst.MoveNext
7290            Next
7300            strMsg = Trim(strMsg)                    ' ** Remove trailing space.
7310            strMsg = Left(strMsg, (Len(strMsg) - 1))  ' ** Remove trailing comma.
7320          Else
7330            If IsNull(rst![journal_USER]) = True Then
7340              strJrnlUser = "{unknown}"
7350            Else
7360              If Trim(rst![journal_USER]) = vbNullString Then
7370                strJrnlUser = "{unknown}"
7380              Else
7390                strJrnlUser = rst![journal_USER]
7400              End If
7410            End If
7420            strMsg = strJrnlUser
7430          End If
7440          DoCmd.Hourglass False
7450          strMsg = IIf(strMsg = CurrentUser, "You have", "The user" & IIf(lngRecs = 1, " " & strMsg & " has", "s " & strMsg & " have")) & _
                " already entered a dividend for this asset!" & vbCrLf & vbCrLf & "Do you still wish to post this dividend?"
              ' ** Internal Access Function: Trust Accountant login.
7460          If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Warning: Possible Duplicate Data") = vbYes Then
7470            DoCmd.Hourglass True
7480            DoEvents
7490            .Requery
7500            DoCmd.GoToRecord , , acNewRec
7510          Else
7520            blnGoToDivReinvest = False
7530            blnTmpGoToDivReinvest = False
7540            DoCmd.RunCommand (acCmdDeleteRecord)
7550          End If
7560          If .dividendCurr_Date.Visible = True Then .dividendCurr_Date.Visible = False
7570        End If
7580        rst.Close
7590        Set rst = Nothing
7600        Set qdf = Nothing
7610        dbs.Close
7620        Set dbs = Nothing

7630        If .dividendAccountNo.Enabled = True Then
7640          .dividendAccountNo.SetFocus
7650        Else
7660          .dividendTransDate.SetFocus
7670        End If

7680        gblnDividendValidated = False

7690        DividendChanged False  ' ** Procedure: Below.

7700        DoCmd.Hourglass True  ' ** Make sure it's still running
7710        DoEvents

7720        blnGoToDivReinvest = blnTmpGoToDivReinvest

            ' ** Reinvest button was pushed.
7730        If blnGoToDivReinvest = True Then

              ' ** Put preliminary record into tblJournal_MiscSold.
7740          Set dbs = CurrentDb
7750          Set rst = dbs.OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
7760          With rst
7770            If .BOF = True And .EOF = True Then
                  ' ** Proceed.
7780            Else
7790              .MoveFirst
7800              .FindFirst "[Journal_ID] = " & CStr(lngThisJrnlID)
7810              If .NoMatch = False Then
7820                .Delete
7830              End If
7840            End If
7850            DoEvents
7860            .AddNew
                ' ** ![jrnlms_id] : AutoNumber.
7870            ![Journal_ID] = lngThisJrnlID
7880            ![journaltype] = "Dividend"
7890            ![jrnlms_ref_id] = 0&
7900            ![jrnlms_parent] = True
7910            ![jrnlms_child] = False
7920            ![jrnlms_note] = "This Dividend was reinvested, generating a Purchase transaction."
7930            ![jrnlms_datemodified] = Now()
7940            .Update
7950            .Bookmark = .LastModified
7960            lngJrnlMS_ID = ![jrnlms_id]
7970          End With
7980          Set rst = Nothing
7990          dbs.Close
8000          Set dbs = Nothing
8010          DoEvents

8020          gblnMessage = True  ' ** If this returns False, the dialog was canceled.
8030          gdblCrtRpt_CostTot = 0#  ' ** Borrowing this variable from the Court Reports.
8040          strDocName = "frmReinvest_Dividend"
8050          DoCmd.OpenForm strDocName, , , , acFormAdd, acDialog, strCallingForm & "~" & gstrAccountNo & "~" & _
                Format(datTransDate, "mm/dd/yyyy") & "~" & CStr(lngThisCurrID) & "~" & _
                CStr(dblICash) & "~" & CStr(glngAssetNo) & "~" & CStr(lngThisJrnlID) & "~" & strShortName

8060          DoCmd.Hourglass True
8070          DoEvents

8080          Set dbs = CurrentDb
8090          Select Case gblnMessage
              Case True
8100            Set rst = dbs.OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
8110            With rst
8120              .MoveFirst
8130              .FindFirst "[jrnlms_id] = " & CStr(lngJrnlMS_ID)
8140              If .NoMatch = False Then
8150                .Edit
8160                ![jrnlms_ref_id] = CLng(gdblCrtRpt_CostTot)  ' ** New Purchase ID from frmReinvest_Dividend.
8170                ![jrnlms_datemodified] = Now()
8180                .Update
8190              End If
8200              .AddNew
                  ' ** ![jrnlms_id] : AutoNumber.
8210              ![Journal_ID] = CLng(gdblCrtRpt_CostTot)
8220              ![journaltype] = "Purchase"
8230              ![jrnlms_ref_id] = lngThisJrnlID
8240              ![jrnlms_parent] = False
8250              ![jrnlms_child] = True
8260              ![jrnlms_note] = "This Purchase was generated by a Dividend reinvestment."
8270              ![jrnlms_datemodified] = Now()
8280              .Update
8290              .Close
8300            End With
8310            Set rst = Nothing
8320            If .cmdDividendOK.Enabled = True Then
8330              .cmdDividendOK.SetFocus
8340            Else
8350              .dividendAccountNo.SetFocus
8360            End If
8370          Case False
                ' ** Delete tblJournal_MiscSold, by specified [jrnlmsid].
8380            Set qdf = dbs.QueryDefs("qryJournal_MiscSold_01")
8390            With qdf.Parameters
8400              ![jrnlmsid] = lngJrnlMS_ID
8410            End With
8420            qdf.Execute
8430            Set qdf = Nothing
8440            DoEvents
                ' ** Update Journal, by specified [jrnlid], [pstd].
8450            Set qdf = dbs.QueryDefs("qryJournal_Dividend_11")
8460            With qdf.Parameters
8470              ![jrnlid] = lngThisJrnlID  ' ** This new Interest record.
8480              ![pstd] = False
8490            End With
8500            qdf.Execute
8510            Set qdf = Nothing
8520            If gdblCrtRpt_CostTot <> 0# Then
                  ' ** Delete Journal, by specified [jrnlid].
8530              Set qdf = dbs.QueryDefs("qryJournal_Dividend_10")
8540              With qdf.Parameters
8550                ![jrnlid] = CLng(gdblCrtRpt_CostTot)  ' ** New Purchase from frmReinvest_Interest.
8560              End With
8570              qdf.Execute dbFailOnError
8580              Set qdf = Nothing
8590              gdblCrtRpt_CostTot = 0#
8600            End If
8610            .dividendAccountNo.SetFocus
8620          End Select

8630          dbs.Close
8640          Set dbs = Nothing

8650        End If  ' ** blnGoToDivReinvest.
8660        blnGoToDivReinvest = False  ' ** reset always.

8670      End If  ' ** blnContinue.

8680      DoCmd.Hourglass False

8690    End With

EXITP:
8700    Set rst = Nothing
8710    Set qdf = Nothing
8720    Set dbs = Nothing
8730    Exit Sub

ERRH:
8740    DoCmd.Hourglass False
8750    THAT_PROC = THIS_PROC
8760    That_Erl = Erl
8770    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8780    Resume EXITP

End Sub

Private Sub cmdDividendCancel_Click()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDividendCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngThisJrnlID As Long, lngThatJrnlID As Long
        Dim blnReinvested As Boolean, strMsg As String
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

8810    With Me

8820      DoCmd.Hourglass True
8830      DoEvents

8840      If IsNull(.dividendID) = True Then
8850        msgResponse = vbCancel
8860      Else

8870        blnReinvested = .dividendReinvested
8880        lngThisJrnlID = .dividendID
8890        lngThatJrnlID = 0&
8900        msgResponse = 0
            ' ** 2  vbCancel
            ' ** 6  vbYes
            ' ** 7  vbNo

            ' ** Check whether this spawned a Purchase using the Reinvest button.
8910        If blnReinvested = True And .posted = True Then
              ' ** Check if the referencing entry still exists.
8920          varTmp00 = DCount("*", "journal", "[CheckNum] = " & CStr(lngThisJrnlID))
8930          If varTmp00 = 1 Then
8940            lngThatJrnlID = DLookup("[ID]", "journal", "[CheckNum] = " & CStr(lngThisJrnlID))
8950            strMsg = "This entry was used for a Purchase reinvestment." & vbCrLf & vbCrLf & _
                  "Yes - Delete both this Dividend and the Purchase reinvestment." & vbCrLf & _
                  "No - Delete only this Dividend." & vbCrLf & _
                  "Cancel - Delete no transactions."
8960            Beep
8970            DoCmd.Hourglass False
8980            msgResponse = MsgBox(strMsg, vbQuestion + vbYesNoCancel + vbDefaultButton3, "Delete Both Transactions")
8990            If msgResponse <> vbCancel Then
9000              DoCmd.Hourglass True
9010              If .dividendCurr_Date.Visible = True Then .dividendCurr_Date.Visible = False
9020              DoEvents
9030            End If
9040          Else
9050            lngThatJrnlID = 0&
9060          End If
9070        End If

9080      End If

9090      If msgResponse = 0 Then
9100        DoCmd.Hourglass False
9110        msgResponse = MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo, "Delete Transaction")
9120        If msgResponse <> vbYes Then
9130          msgResponse = vbCancel
9140        Else
9150          msgResponse = vbNo  ' ** Delete only Dividend.
9160          DoCmd.Hourglass True
9170          If .dividendCurr_Date.Visible = True Then .dividendCurr_Date.Visible = False
9180          DoEvents
9190        End If
9200      End If

9210      If msgResponse = vbCancel Then
9220        If .dividendAccountNo.Enabled = True Then
9230          .dividendAccountNo.SetFocus
9240        Else
9250          .dividendTransDate.SetFocus
9260        End If
9270      Else

9280        .Undo
9290        .Undo

9300        If blnReinvested = True Then
              ' ** Delete both, whether or not reinvestment is deleted,
              ' ** since it would no longer have a parent.
9310          Set dbs = CurrentDb
              ' ** Delete tblJournal_MiscSold, by specified [jrnlid].
9320          Set qdf = dbs.QueryDefs("qryJournal_MiscSold_02")
9330          With qdf.Parameters
9340            ![jrnlid] = lngThisJrnlID
9350          End With
9360          qdf.Execute
9370          Set qdf = Nothing
9380          DoEvents
              ' ** Delete tblJournal_MiscSold, by specified [jrnlid].
9390          Set qdf = dbs.QueryDefs("qryJournal_MiscSold_02")
9400          With qdf.Parameters
9410            ![jrnlid] = lngThatJrnlID
9420          End With
9430          qdf.Execute
9440          Set qdf = Nothing
9450          dbs.Close
9460          Set dbs = Nothing
9470          DoEvents
9480        End If  ' ** blnReinvested.

9490        DoCmd.SetWarnings False
9500  On Error Resume Next
9510        DoCmd.RunCommand acCmdSelectRecord
9520        DoCmd.RunCommand acCmdDeleteRecord
9530  On Error GoTo ERRH
9540        DoCmd.SetWarnings True

9550        blnGoToDivReinvest = False
9560        .tglDividendReinvest_false_raised_img_dis.Visible = True
9570        .tglDividendReinvest_false_raised_img.Visible = False
9580        .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = False
9590        .tglDividendReinvest_false_raised_focus_img.Visible = False
9600        .tglDividendReinvest_false_raised_focus_dots_img.Visible = False
9610        .tglDividendReinvest_false_sunken_focus_dots_img.Visible = False
9620        .tglDividendReinvest_true_raised_img.Visible = False
9630        .tglDividendReinvest_true_raised_focus_img.Visible = False
9640        .tglDividendReinvest_true_raised_focus_dots_img.Visible = False
9650        .tglDividendReinvest_true_sunken_focus_dots_img.Visible = False
9660        .tglDividendReinvest_true_raised_img_dis.Visible = False
9670        .tglDividendReinvest.Enabled = False

9680        .cmbAccountHelper = Null

            ' ** Also delete reinvestment.
9690        If msgResponse = vbYes And lngThatJrnlID > 0& Then
9700          Set dbs = CurrentDb
              ' ** Delete Journal, by specified [jrnlid].
9710          Set qdf = dbs.QueryDefs("qryJournal_Dividend_10")
9720          With qdf.Parameters
9730            ![jrnlid] = lngThatJrnlID
9740          End With
9750          qdf.Execute
9760          Set qdf = Nothing
9770          dbs.Close
9780          Set dbs = Nothing
9790        End If

9800        .Parent.frmJournal_Sub3_Purchase.Form.Requery
9810        DoEvents

9820        If .dividendAccountNo.Enabled = True Then
9830          .dividendAccountNo.SetFocus
9840        Else
9850          .dividendTransDate.SetFocus
9860        End If

9870        gblnDividendValidated = False

9880        DividendChanged False  ' ** Procedure: Below.

9890        Form_Current  ' ** Procedure: Above.

9900      End If  ' ** msgResponse.

9910      DoCmd.Hourglass False

9920    End With

EXITP:
9930    Set qdf = Nothing
9940    Set dbs = Nothing
9950    Exit Sub

ERRH:
9960    DoCmd.Hourglass False
9970    THAT_PROC = THIS_PROC
9980    That_Erl = Erl
9990    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10000   Resume EXITP

End Sub

Private Sub cmbAccountHelper_Enter()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_Enter"

10110   With Me
10120     Select Case IsNull(.cmbAccountHelper)
          Case True
10130       ExpandCombo .cmbAccountHelper  ' ** Module Procedure: modUtilities.
10140     Case False
10150       If Trim(.cmbAccountHelper) = vbNullString Then
10160         ExpandCombo .cmbAccountHelper  ' ** Module Procedure: modUtilities.
10170       End If
10180     End Select
10190   End With

EXITP:
10200   Exit Sub

ERRH:
10210   Select Case ERR.Number
        Case Else
10220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10230   End Select
10240   Resume EXITP

End Sub

Private Sub cmbAccountHelper_GotFocus()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_GotFocus"

10310   With Me
10320     .cmbAccountHelper.ForeColor = 0
10330   End With

EXITP:
10340   Exit Sub

ERRH:
10350   Select Case ERR.Number
        Case Else
10360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10370   End Select
10380   Resume EXITP

End Sub

Private Sub cmbAccountHelper_AfterUpdate()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_AfterUpdate"

10410   With Me
10420     If IsNull(.cmbAccountHelper) = False Then
10430       .dividendAccountNo = .cmbAccountHelper
10440       .dividendAccountNo.SetFocus
10450     End If
10460   End With

EXITP:
10470   Exit Sub

ERRH:
10480   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
10490   Case Else
10500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10510   End Select
10520   Resume EXITP

End Sub

Private Sub cmbAccountHelper_LostFocus()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_LostFocus"

10610   With Me
10620     gblnDividendValidated = False
10630     DividendChanged True  ' ** Procedure: Below.
10640     .cmbAccountHelper.ForeColor = CLR_WHT
10650     .dividendAccountNo.SetFocus
10660     DoEvents
10670     .dividendAccountNo.Value = .cmbAccountHelper.Value
10680     DoEvents
10690   End With

EXITP:
10700   Exit Sub

ERRH:
10710   Select Case ERR.Number
        Case Else
10720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10730   End Select
10740   Resume EXITP

End Sub

Private Sub dividendAccountNo_AfterUpdate()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAccountNo_AfterUpdate"

10810   DividendChanged True  ' ** Procedure: Below.

EXITP:
10820   Exit Sub

ERRH:
10830   THAT_PROC = THIS_PROC
10840   That_Erl = Erl
10850   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10860   Resume EXITP

End Sub

Private Sub dividendAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAccountNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10920   intShiftDown = (Shift And acShiftMask) > 0
10930   intAltDown = (Shift And acAltMask) > 0
10940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10960     Select Case intRetVal
          Case vbKeyTab
            ' ** Tab key copies AccountNo. {intRetVal not cleared.}
10970       With Me
10980         If gblnTabCopyAccount = True Then
10990           If .NewRecord = True Then
11000             If IsNull(.dividendAccountNo) = True Then
11010               If Len(.dividendAccountNo.text) = 0 Then
11020                 If IsNull(.Parent.LastAcctNo) = False Then
11030                   If .Parent.LastAcctNo <> vbNullString Then
11040                     .dividendAccountNo = .Parent.LastAcctNo
11050                     DividendChanged True  ' ** Procedure: Below.
11060                   End If
11070                 End If
11080               End If
11090             Else
11100               If Len(.dividendAccountNo.text) = 0 Then
11110                 If IsNull(.Parent.LastAcctNo) = False Then
11120                   If .Parent.LastAcctNo <> vbNullString Then
11130                     .dividendAccountNo = .Parent.LastAcctNo
11140                     DividendChanged True  ' ** Procedure: Below.
11150                   End If
11160                 End If
11170               End If
11180             End If
11190           End If
11200         End If
11210         gblnDividendChanged = True
11220         gblnDividendValidated = False
11230       End With
11240     Case vbKeyReturn
11250       With Me
11260         intRetVal = 0
11270         gblnDividendChanged = True  ' ** Just added: 09/19/2012.
11280         gblnDividendValidated = False
11290         .dividendTransDate.SetFocus
11300       End With
11310     Case vbKeyDelete
11320       gblnDividendValidated = False
11330     End Select
11340   End If

EXITP:
11350   KeyCode = intRetVal
11360   Exit Sub

ERRH:
11370   intRetVal = 0
11380   THAT_PROC = THIS_PROC
11390   That_Erl = Erl
11400   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11410   Resume EXITP

End Sub

Private Sub dividendAccountNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAccountNo_MouseMove"

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

Private Sub dividendAccountNo_Exit(Cancel As Integer)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAccountNo_Exit"

        Dim strAccountNo As String
        Dim blnGoToAccountHelper As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

11610   With Me

11620     blnGoToAccountHelper = False  ' ** Unless set otherwise.

11630     If ((gblnDividendValidated = False And gblnDividendChanged = True) And (lngErrCnt < 3&)) Then
11640       If IsNull(.dividendAccountNo) = True Then
11650         lngErrCnt = lngErrCnt + 1&
11660         msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Entry Required")
11670         If msgResponse = vbYes Then
11680           blnGoToAccountHelper = True
11690         Else
11700           gblnDividendValidated = False
11710           If blnStartTrans = True Then
11720             Cancel = -1
11730             .dividendAccountNo.SetFocus
11740           End If
11750         End If
11760       Else
11770         If .dividendAccountNo = vbNullString And .dividendAccountNo.text = vbNullString Then
11780           lngErrCnt = lngErrCnt + 1&
11790           msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                  "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Entry Required")
11800           If msgResponse = vbYes Then
11810             blnGoToAccountHelper = True
11820           Else
11830             gblnDividendValidated = False
11840             If blnStartTrans = True Then
11850               Cancel = -1
11860               .dividendAccountNo.SetFocus
11870             End If
11880           End If
11890         Else
                ' ** Check to see if it is a valid account.
11900           lngErrCnt = 0&
11910           Select Case gblnAccountNoWithType
                Case True
11920             strAccountNo = Left(.dividendAccountNo, 2) & Mid(.dividendAccountNo, 3)
11930           Case False
11940             strAccountNo = .dividendAccountNo
11950           End Select
11960           varTmp00 = DLookup("[accountno]", "account", "[accountno] = '" & strAccountNo & "' And ([dateclosed] Is Null)")
11970           Select Case IsNull(varTmp00)
                Case True
11980             blnStartTrans = True  ' ** Once they've started, they've got to enter a valid number!
11990             msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                    "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Invalid Entry")
12000             If msgResponse = vbYes Then
12010               blnGoToAccountHelper = True
12020             Else
12030               gblnDividendValidated = False
12040               Cancel = -1
12050               .dividendAccountNo.SetFocus
12060             End If
12070             .dividendAccountNo = vbNullString
12080           Case False
12090             gblnDividendValidated = True
12100           End Select
12110           If IsNull(.dividendAccountNo) = False Then
12120             If .dividendAccountNo.text <> vbNullString Then
12130               .dividendAccountNo_Data = .dividendAccountNo
12140               gstrDividendAccountNumber = .dividendAccountNo.text
12150               .Parent.LastAcctNo = gstrDividendAccountNumber
12160               .cmdDividendMap.Enabled = False
12170 On Error Resume Next
12180               .Refresh
12190               If ERR.Number <> 0 Then  ' ** 3314: The field 'journal.journaltype' cannot contain a Null value because the Required property for this field is set to True.
12200                 .dividendType = "Dividend"
12210                 blnAccountNoErr = True
12220               End If
12230 On Error GoTo ERRH
12240               .tglDividendReinvest.Enabled = True
12250               .tglDividendReinvest_false_raised_img.Visible = True
12260               .tglDividendReinvest_false_raised_img_dis.Visible = False
12270             End If
12280           End If
12290         End If
12300       End If
12310     Else
12320       lngErrCnt = 0&
12330     End If

12340     If blnGoToAccountHelper = True Then
12350       .cmbAccountHelper.SetFocus
12360     End If

12370   End With  ' ** Me.

EXITP:
12380   Exit Sub

ERRH:
12390   THAT_PROC = THIS_PROC
12400   That_Erl = Erl
12410   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12420   Resume EXITP

End Sub

Private Sub cmdLock_Click()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLock_Click"

12510   With Me
12520     If .cmdLock_open_raised_img.Visible = True Then
12530       .dividendTransDate.SetFocus
12540       .dividendAccountNo.Enabled = False
12550       .dividendAccountNo.BorderColor = WIN_CLR_DISR
12560       .dividendAccountNo_lbl.BackStyle = acBackStyleTransparent
12570       .dividendAccountNo_lbl_box.Visible = True
12580       .cmdLock_open_raised_img.Visible = False
12590       .cmdLock_closed_raised_img.Visible = True
12600     ElseIf .cmdLock_closed_raised_img.Visible = True Then
12610       .dividendAccountNo.Enabled = True
12620       .dividendAccountNo.BorderColor = CLR_LTBLU2
12630       .dividendAccountNo_lbl.BackStyle = acBackStyleNormal
12640       .dividendAccountNo_lbl_box.Visible = False
12650       .dividendAccountNo.SetFocus
12660       .cmdLock_open_raised_img.Visible = True
12670       .cmdLock_closed_raised_img.Visible = False
12680     End If
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

Private Sub dividendShortName_LostFocus()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendShortName_LostFocus"

12810   VerifyAccountNo  ' ** Function: Below.

EXITP:
12820   Exit Sub

ERRH:
12830   THAT_PROC = THIS_PROC
12840   That_Erl = Erl
12850   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12860   Resume EXITP

End Sub

Private Sub dividendType_LostFocus()

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendType_LostFocus"

12910   VerifyAccountNo  ' ** Function: Below.

12920   With Me
12930     .dividendType.SetFocus
12940     gstrDividendType = .dividendType.text
12950   End With

EXITP:
12960   Exit Sub

ERRH:
12970   THAT_PROC = THIS_PROC
12980   That_Erl = Erl
12990   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13000   Resume EXITP

End Sub

Public Sub cmdDividendMap_Click()
' ** Public for Go-To-Report.

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDividendMap_Click"

        Dim blnContinue As Boolean
        Dim strDocName As String
        Dim varTmp00 As Variant

13110   blnContinue = True

13120   With Me

13130     If .dividendAccountNo.Enabled = True Then
13140       .dividendAccountNo.SetFocus
13150     Else
13160       .dividendTransDate.SetFocus
13170     End If
13180     If IsNull(.dividendAccountNo) Then
            ' ** Open the map.
13190     Else
13200       If Len(Trim(.dividendAccountNo.text)) = 0 Then
              ' ** Open the map.
13210       Else
13220         blnContinue = False
13230         .cmdDividendMap.SetFocus
13240         MsgBox "Please advance to a blank posting form.", vbInformation + vbOKOnly, "New Record Required"
13250         DoCmd.CancelEvent
13260       End If
13270     End If

13280     If blnContinue = True Then

13290       .cmdDividendMap.SetFocus

13300       .Parent.LoadMapDropdowns  ' ** Form Procedure: frmJournal.
13310       DoEvents

13320       strDocName = "frmMap_Div"
13330       DoCmd.OpenForm strDocName, , , , , , strCallingForm

13340       If gblnGoToReport = True Then
13350         varTmp00 = DMax("[ID]", "journal")  ' ** Save this so we can delete any fake Journal records.
13360         Select Case IsNull(varTmp00)
              Case True
13370           glngTaxCode_Distribution = 0&  ' ** Borrowing this variable from the Court Reports.
13380         Case False
13390           glngTaxCode_Distribution = varTmp00
13400         End Select
13410         Forms(strDocName).TimerInterval = 100&
13420         .GoToReport_arw_mapdiv_img.Visible = False
13430         DoCmd.Hourglass True  ' ** Make sure it's still running.
13440         DoEvents
13450       End If

13460     End If

13470   End With

EXITP:
13480   Exit Sub

ERRH:
13490   DoCmd.Hourglass False
13500   If gblnGoToReport = True Then
13510     gblnGoToReport = False
13520     Me.Parent.GTREmblem_Set False  ' ** Form Procedure: frmJournal.
13530   End If
13540   THAT_PROC = THIS_PROC
13550   That_Erl = Erl
13560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13570   Resume EXITP

End Sub

Private Sub cmdDividendMap_KeyDown(KeyCode As Integer, Shift As Integer)
' ** See Note, above.

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDividendMap_KeyDown"

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
          Case vbKeyTab
13670       With Me
13680         intRetVal = 0
13690         .dividendTransDate.SetFocus
13700       End With
13710     End Select
13720   End If

        ' ** Shift keys.
13730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13740     Select Case intRetVal
          Case vbKeyTab
13750       With Me
13760         intRetVal = 0
13770         .dividendType.SetFocus
13780       End With
13790     End Select
13800   End If

EXITP:
13810   KeyCode = intRetVal
13820   Exit Sub

ERRH:
13830   intRetVal = 0
13840   THAT_PROC = THIS_PROC
13850   That_Erl = Erl
13860   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13870   Resume EXITP

End Sub

Private Sub dividendTransDate_BeforeUpdate(Cancel As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendTransDate_BeforeUpdate"

13910   With Me
13920     If .dividendTransDate.text <> "__/__/____" And .dividendTransDate.text <> vbNullString Then
13930       If Left(.dividendTransDate.text, 2) > 12 Then
13940         MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Date"
13950         Cancel = -1
13960       End If
13970     End If
13980   End With

EXITP:
13990   Exit Sub

ERRH:
14000   THAT_PROC = THIS_PROC
14010   That_Erl = Erl
14020   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14030   Resume EXITP

End Sub

Private Sub dividendTransDate_KeyDown(KeyCode As Integer, Shift As Integer)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendTransDate_KeyDown"

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
14190         .dividendAssetNo.SetFocus
14200       End With
14210     End Select
14220   End If

        ' ** Shift keys.
14230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14250       With Me
14260         intRetVal = 0
14270         .dividendType.SetFocus
14280       End With
14290     End Select
14300   End If

        ' ** Ctrl keys.
14310   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
14320     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14330       With Me
14340         intRetVal = 0
14350         .cmdCalendar1.SetFocus
14360       End With
14370     End Select
14380   End If

EXITP:
14390   KeyCode = intRetVal
14400   Exit Sub

ERRH:
14410   intRetVal = 0
14420   THAT_PROC = THIS_PROC
14430   That_Erl = Erl
14440   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14450   Resume EXITP

End Sub

Private Sub dividendTransDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendTransDate_MouseMove"

14510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14520   Exit Sub

ERRH:
14530   Select Case ERR.Number
        Case Else
14540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14550   End Select
14560   Resume EXITP

End Sub

Private Sub dividendTransDate_Exit(Cancel As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendTransDate_Exit"

14610   With Me
14620     If lngDateCnt <= 1& Then
14630       If IsNull(.dividendAccountNo) = False Then
14640         If Trim(.dividendAccountNo) <> vbNullString Then
14650           If DateCheck_Post(.dividendTransDate.text) = True Then  ' ** Module Function: modUtilities.
                  ' ** '.Text' needed to reliably check date.
14660             lngDateCnt = 0&
14670             If IsNull(.dividendTransDate.OldValue) Then
14680               DividendChanged True  ' ** Procedure: Below.
14690             Else
14700               If .dividendTransDate <> .dividendTransDate.OldValue Then
14710                 DividendChanged True  ' ** Procedure: Below.
14720               End If
14730             End If
14740           Else
14750             If IsNull(.dividendTransDate) = True Then
14760               .dividendTransDate = Date
14770             Else
14780               If Trim(.dividendTransDate) = vbNullString Then
14790                 .dividendTransDate = Date
14800               Else
14810                 lngDateCnt = lngDateCnt + 1&
14820                 Cancel = -1
14830                 .dividendTransDate.SetFocus
14840               End If
14850             End If
14860           End If
14870         End If
14880       End If
14890     Else
14900       lngDateCnt = 0&
14910       .dividendTransDate = CDate(Format(datPostingDate, "mm/dd/yyyy"))
14920     End If
14930   End With

EXITP:
14940   Exit Sub

ERRH:
14950   THAT_PROC = THIS_PROC
14960   That_Erl = Erl
14970   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14980   Resume EXITP

End Sub

Private Sub dividendTransDate_LostFocus()

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendTransDate_LostFocus"

15010   With Me
15020     If IsNull(.dividendAccountNo) = False Then
15030       If Trim(.dividendAccountNo) <> vbNullString Then
15040         VerifyAccountNo  ' ** Function: Below.
15050       End If
15060     End If
15070   End With

EXITP:
15080   Exit Sub

ERRH:
15090   THAT_PROC = THIS_PROC
15100   That_Erl = Erl
15110   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15120   Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

15210   With Me
15220     blnCalendar1_Focus = True
15230     .cmdCalendar1_raised_semifocus_dots_img.Visible = True
15240     .cmdCalendar1_raised_img.Visible = False
15250     .cmdCalendar1_raised_focus_img.Visible = False
15260     .cmdCalendar1_raised_focus_dots_img.Visible = False
15270     .cmdCalendar1_sunken_focus_dots_img.Visible = False
15280     .cmdCalendar1_raised_img_dis.Visible = False
15290   End With

EXITP:
15300   Exit Sub

ERRH:
15310   Select Case ERR.Number
        Case Else
15320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15330   End Select
15340   Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

15410   With Me
15420     blnCalendar1_MouseDown = True
15430     .cmdCalendar1_sunken_focus_dots_img.Visible = True
15440     .cmdCalendar1_raised_img.Visible = False
15450     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
15460     .cmdCalendar1_raised_focus_img.Visible = False
15470     .cmdCalendar1_raised_focus_dots_img.Visible = False
15480     .cmdCalendar1_raised_img_dis.Visible = False
15490   End With

EXITP:
15500   Exit Sub

ERRH:
15510   Select Case ERR.Number
        Case Else
15520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15530   End Select
15540   Resume EXITP

End Sub

Public Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

15610   With Me
15620     datStartDate = Date
15630     datEndDate = 0
15640     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
15650     If blnRetVal = True Then
            ' ** Allow posting up to 1 month into the future.
15660       If datStartDate > DateAdd("m", 1, Date) Then
15670         MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
15680         .dividendTransDate = CDate(Format(Date, "mm/dd/yyyy"))
15690       Else
15700         .dividendTransDate = datStartDate
15710       End If
15720     Else
15730       .dividendTransDate = CDate(Format(Date, "mm/dd/yyyy"))
15740     End If
15750     .dividendAssetNo.SetFocus
15760   End With

EXITP:
15770   Exit Sub

ERRH:
15780   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
15790   Case Else
15800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15810   End Select
15820   Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

15910   With Me
15920     If blnCalendar1_MouseDown = False Then
15930       Select Case blnCalendar1_Focus
            Case True
15940         .cmdCalendar1_raised_focus_dots_img.Visible = True
15950         .cmdCalendar1_raised_focus_img.Visible = False
15960       Case False
15970         .cmdCalendar1_raised_focus_img.Visible = True
15980         .cmdCalendar1_raised_focus_dots_img.Visible = False
15990       End Select
16000       .cmdCalendar1_raised_img.Visible = False
16010       .cmdCalendar1_raised_semifocus_dots_img.Visible = False
16020       .cmdCalendar1_sunken_focus_dots_img.Visible = False
16030       .cmdCalendar1_raised_img_dis.Visible = False
16040     End If
16050   End With

EXITP:
16060   Exit Sub

ERRH:
16070   Select Case ERR.Number
        Case Else
16080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16090   End Select
16100   Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

16210   With Me
16220     .cmdCalendar1_raised_focus_dots_img.Visible = True
16230     .cmdCalendar1_raised_img.Visible = False
16240     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
16250     .cmdCalendar1_raised_focus_img.Visible = False
16260     .cmdCalendar1_sunken_focus_dots_img.Visible = False
16270     .cmdCalendar1_raised_img_dis.Visible = False
16280     blnCalendar1_MouseDown = False
16290   End With

EXITP:
16300   Exit Sub

ERRH:
16310   Select Case ERR.Number
        Case Else
16320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16330   End Select
16340   Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

16410   With Me
16420     .cmdCalendar1_raised_img.Visible = True
16430     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
16440     .cmdCalendar1_raised_focus_img.Visible = False
16450     .cmdCalendar1_raised_focus_dots_img.Visible = False
16460     .cmdCalendar1_sunken_focus_dots_img.Visible = False
16470     .cmdCalendar1_raised_img_dis.Visible = False
16480     blnCalendar1_Focus = False
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

Private Sub chkShowAllAssets_AfterUpdate()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAllAssets_AfterUpdate"

16610   With Me
16620     Select Case .chkShowAllAssets
          Case True
16630       If .dividendAssetNo.RowSource <> "qryJournal_Dividend_02g" Then
              ' ** qryJournal_Dividend_02c (MasterAsset table, with totdesc)
              ' ** linked to AssetType, Dividend = True, all assets.
16640         .dividendAssetNo.RowSource = "qryJournal_Dividend_02g"
16650       Else
16660         .dividendAssetNo.Requery
16670       End If
16680     Case False
16690       If .dividendAssetNo.RowSource <> "qryJournal_Dividend_02e" Then
              ' ** qryJournal_Dividend_02a (MasterAsset table linked to ActiveAssets, Account) linked
              ' ** to AssetType, Dividend = True, by specified FormRef('frmJournal_Sub1_Dividend').
16700         .dividendAssetNo.RowSource = "qryJournal_Dividend_02e"
16710       Else
16720         .dividendAssetNo.Requery
16730       End If
16740     End Select
16750   End With

        ' ** These queries are used on the Form_Current() event.
        ' **   qryJournal_Dividend_02d
        ' **   qryJournal_Dividend_02f

EXITP:
16760   Exit Sub

ERRH:
16770   Select Case ERR.Number
        Case Else
16780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16790   End Select
16800   Resume EXITP

End Sub

Private Sub chkShowAllAssets_KeyDown(KeyCode As Integer, Shift As Integer)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAllAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16920   intShiftDown = (Shift And acShiftMask) > 0
16930   intAltDown = (Shift And acAltMask) > 0
16940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16970       With Me
16980         intRetVal = 0
16990         .dividendAssetNo.SetFocus
17000       End With
17010     End Select
17020   End If

        ' ** Shift keys.
17030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17050       With Me
17060         intRetVal = 0
17070         .dividendTransDate.SetFocus
17080       End With
17090     End Select
17100   End If

EXITP:
17110   KeyCode = intRetVal
17120   Exit Sub

ERRH:
17130   intRetVal = 0
17140   THAT_PROC = THIS_PROC
17150   That_Erl = Erl
17160   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17170   Resume EXITP

End Sub

Private Sub dividendAssetNo_Enter()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAssetNo_Enter"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnContinue As Boolean
        Dim blnHasDiv As Boolean
        Dim strSQL As String
        Dim varTmp00 As Variant

17210   blnContinue = True

17220   With Me

17230     If Not blnAccountValidated Then
17240       blnCheckingAccount = True
17250       blnExitingAssetNo = True
17260       VerifyAccountNo  ' ** Function: Below.
17270       blnCheckingAccount = False
17280       blnExitingAssetNo = False
17290       If blnAccountValidated = False Then
17300         blnContinue = False
17310         blnExitingAssetNo = True
17320       Else
17330         .dividendAssetNo.SetFocus
17340       End If
17350     End If

17360     If blnContinue = True Then

17370       If .dividendAccountNo = "INCOME O/U" Or .dividendAccountNo = "99-INCOME O/U" Then
              'strSQL ISN'T USED ANYWHERE!
              'IF CODE GETS HERE, ROW SOURCE DEFAULTS TO EXISTING qryJournal_Dividend_02!
              ' ** qryJournal_Dividend_02f.
17380         strSQL = "SELECT DISTINCTROW qryJournal_Dividend_02b.totdesc, qryJournal_Dividend_02b.cusip, " & _
                "qryJournal_Dividend_02b.assetno, assettype.taxcode, qryJournal_Dividend_02b.curr_id " & _
                "FROM qryJournal_Dividend_02b INNER JOIN assettype ON qryJournal_Dividend_02b.assettype = assettype.assettype " & _
                "WHERE (((qryJournal_Dividend_02b.accountno)='" & .dividendAccountNo & "')) " & _
                "ORDER BY qryJournal_Dividend_02b.assettype, qryJournal_Dividend_02b.totdesc;"  '#curr_id
17390       Else

              ' ** Check to see if this account has any Dividend assets!
17400         blnHasDiv = True
17410         Set dbs = CurrentDb
17420         Set qdf = dbs.QueryDefs("qryJournal_Dividend_06")
17430         With qdf.Parameters
17440           ![actno] = Me.dividendAccountNo
17450         End With
17460         Set rst = qdf.OpenRecordset
17470         With rst
17480           If .BOF = True And .EOF = True Then
17490             blnHasDiv = False
17500           End If
17510           .Close
17520         End With
17530         dbs.Close

17540         Select Case .chkShowAllAssets
              Case True
17550           If .dividendAssetNo.RowSource <> "qryJournal_Dividend_02g" Then
                  ' ** qryJournal_Dividend_02c (MasterAsset table, with totdesc)
                  ' ** linked to AssetType, Dividend = True, all assets.
17560             .dividendAssetNo.RowSource = "qryJournal_Dividend_02g"
17570           Else
17580             .dividendAssetNo.Requery
17590           End If
17600         Case False
                ' ** Only assets for this account.
17610           If blnHasDiv = False Then
17620             blnContinue = False
17630             DoCmd.Hourglass False
17640             MsgBox "The chosen account has no Dividend assets.", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
17650           Else
17660             If .dividendAssetNo.RowSource <> "qryJournal_Dividend_02e" Then
                    ' ** qryJournal_Dividend_02a (MasterAsset table linked to ActiveAssets, Account) linked
                    ' ** to AssetType, Dividend = True, by specified FormRef('frmJournal_Sub1_Dividend').
17670               .dividendAssetNo.RowSource = "qryJournal_Dividend_02e"
17680             Else
17690               .dividendAssetNo.Requery
17700             End If
17710           End If
17720         End Select
17730       End If

17740       If blnContinue = True Then
17750         .dividendAssetNo.Requery
17760         DoEvents
17770         If IsNull(.dividendAssetNo) = True Then
17780           ExpandCombo .dividendAssetNo  ' ** Module Procedure: modUtilities.
17790         Else
17800           varTmp00 = vbNullString
17810 On Error Resume Next
17820           varTmp00 = .dividendAssetNo.text
17830           If ERR.Number <> 0 Then  ' ** When tabbing through quickly, it may not get here before it's already moved on!
17840 On Error GoTo ERRH
17850           Else
17860 On Error GoTo ERRH
17870             If varTmp00 = vbNullString Then
17880               ExpandCombo .dividendAssetNo  ' ** Module Procedure: modUtilities.
17890             End If
17900           End If
17910         End If
17920       End If

17930     End If  ' ** blnContinue.

17940   End With  ' ** Me.

EXITP:
17950   Set rst = Nothing
17960   Set qdf = Nothing
17970   Set dbs = Nothing
17980   Exit Sub

ERRH:
17990   THAT_PROC = THIS_PROC
18000   That_Erl = Erl
18010   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18020   Resume EXITP

End Sub

Private Sub dividendAssetNo_NotInList(NewData As String, Response As Integer)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAssetNo_NotInList"

18110   With Me
18120     .dividendAssetNo.SelStart = 0
18130     .dividendAssetNo.SelLength = Len(.dividendAssetNo.text)
18140   End With

EXITP:
18150   Exit Sub

ERRH:
18160   Select Case ERR.Number
        Case Else
18170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18180   End Select
18190   Resume EXITP

End Sub

Private Sub dividendAssetNo_BeforeUpdate(Cancel As Integer)

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAssetNo_BeforeUpdate"

        Dim blnCurrID As Boolean
        Dim varTmp00 As Variant

18210   With Me
18220     If IsNull(.dividendAssetNo) = False Then
18230       blnCurrID = .dividendCurr_ID.Visible
18240       If blnCurrID = False Then
18250         varTmp00 = .dividendAssetNo.Column(CBX_A_CURRID)
18260         If IsNull(varTmp00) = False Then
18270           If CLng(varTmp00) <> 150& Then
18280             Beep
18290             MsgBox "The chosen asset is in a foreign currency." & vbCrLf & vbCrLf & _
                    "You must check the 'Include Currency' checkbox on the Posting Date" & vbCrLf & _
                    "window before entering transactions not in U.S. dollars.", vbInformation + vbOKOnly, "Foreign Currency Asset"
18300             .dividendAssetNo.Undo
18310             Cancel = -1
18320           End If
18330         End If
18340       End If
18350     End If
18360   End With

EXITP:
18370   Exit Sub

ERRH:
18380   Select Case ERR.Number
        Case Else
18390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18400   End Select
18410   Resume EXITP

End Sub

Private Sub dividendAssetNo_AfterUpdate()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAssetNo_AfterUpdate"

18510   With Me

          ' ** Do this whether or not Currency is visible.
18520     Select Case IsNull(.dividendAssetNo)
          Case True
18530       lngCurrID = 150&  ' ** USD.
18540     Case False
18550       Select Case IsNull(.dividendAssetNo.Column(CBX_A_CURRID))
            Case True
18560         lngCurrID = 150&  ' ** USD.
18570       Case False
18580         lngCurrID = .dividendAssetNo.Column(CBX_A_CURRID)
18590       End Select
18600     End Select
18610     .dividendCurr_ID = lngCurrID
18620     dividendCurr_ID_AfterUpdate  ' ** Procedure: Below.

18630     If IsNull(.dividendAssetNo) = False Then

18640       If .dividendAssetNo <> .dividendAssetNo.OldValue Then
18650         DividendChanged True  ' ** Procedure: Below.
18660       End If

18670       If gblnIncomeTaxCoding = True Then
18680         .cmbTaxCodes = dividendAssetNo.Column(CBX_A_TAX)  ' ** All AssetType-based Tax Codes are INCOME.
18690         .cmbTaxCodes.Requery
18700       End If

18710     End If

18720   End With

EXITP:
18730   Exit Sub

ERRH:
18740   THAT_PROC = THIS_PROC
18750   That_Erl = Erl
18760   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18770   Resume EXITP

End Sub

Private Sub dividendAssetNo_KeyDown(KeyCode As Integer, Shift As Integer)

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAssetNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18820   intShiftDown = (Shift And acShiftMask) > 0
18830   intAltDown = (Shift And acAltMask) > 0
18840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18870       With Me
18880         intRetVal = 0
18890         .dividendShareFace.SetFocus
18900       End With
18910     End Select
18920   End If

        ' ** Shift keys.
18930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18950       With Me
18960         intRetVal = 0
18970         .dividendTransDate.SetFocus
18980       End With
18990     End Select
19000   End If

        ' ** Ctrl-Shift keys.
19010   If intCtrlDown And (Not intAltDown) And intShiftDown Then
19020     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19030       With Me
19040         intRetVal = 0
19050         .cmdCalendar1.SetFocus
19060       End With
19070     End Select
19080   End If

EXITP:
19090   KeyCode = intRetVal
19100   Exit Sub

ERRH:
19110   intRetVal = 0
19120   THAT_PROC = THIS_PROC
19130   That_Erl = Erl
19140   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19150   Resume EXITP

End Sub

Private Sub dividendAssetNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAssetNo_MouseMove"

19210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
19220   Exit Sub

ERRH:
19230   Select Case ERR.Number
        Case Else
19240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19250   End Select
19260   Resume EXITP

End Sub

Private Sub dividendAssetNo_Exit(Cancel As Integer)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAssetNo_Exit"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim varTmp00 As Variant

        ' ** Set shares box value.
19310   With Me
19320     If IsNull(.dividendAssetNo) = False Then
19330       If .dividendAssetNo <> vbNullString And IsNull(.dividendAccountNo) = False Then
19340         If .dividendAccountNo <> vbNullString Then
19350           If blnExitingAssetNo = False Then
19360             Set dbs = CurrentDb
19370             Set rst = dbs.OpenRecordset("SELECT SUM(shareface) AS sumsf FROM activeAssets WHERE assetno = " & _
                    CStr(.dividendAssetNo) & " AND accountno = '" & .dividendAccountNo & "'")
19380             If rst.BOF = True And rst.EOF = True Then
19390               .dividendShareFace = 0
19400             Else
19410               rst.MoveLast
19420               If rst.RecordCount > 0 Then
19430                 rst.MoveFirst
19440                 .dividendShareFace = rst![sumsf]
19450               Else
19460                 .dividendShareFace = 0
19470               End If
19480             End If
19490             rst.Close
19500             dbs.Close
19510           End If
                ' ** Check curr_id.
19520           varTmp00 = .dividendAssetNo.Column(CBX_A_CURRID)
19530           If IsNull(varTmp00) = False Then
19540             If .dividendCurr_ID <> varTmp00 Then
19550               .dividendCurr_ID = varTmp00
19560               dividendCurr_ID_AfterUpdate  ' ** Procedure: Below.
19570             End If
19580           End If
19590         End If
19600       End If
19610     End If
19620   End With

EXITP:
19630   Set rst = Nothing
19640   Set dbs = Nothing
19650   Exit Sub

ERRH:
19660   THAT_PROC = THIS_PROC
19670   That_Erl = Erl
19680   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19690   Resume EXITP

End Sub

Private Sub dividendAssetNo_LostFocus()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendAssetNo_LostFocus"

19710   With Me
19720     If blnExitingAssetNo = False Then

19730       VerifyAccountNo  ' ** Function: Below.

19740       gstrDividendAsset = .dividendAssetNo.text
19750       If .dividendAssetNo <> .dividendAssetNo.OldValue Then
19760         DividendChanged True  ' ** Procedure: Below.
19770       End If

19780     End If
19790   End With

EXITP:
19800   Exit Sub

ERRH:
19810   THAT_PROC = THIS_PROC
19820   That_Erl = Erl
19830   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19840   Resume EXITP

End Sub

Private Sub dividendShareFace_Enter()

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendShareFace_Enter"

19910   With Me
19920     .dividendShareFace.Format = "General Number"
19930   End With

EXITP:
19940   Exit Sub

ERRH:
19950   THAT_PROC = THIS_PROC
19960   That_Erl = Erl
19970   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19980   Resume EXITP

End Sub

Private Sub dividendShareFace_Exit(Cancel As Integer)

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendShareFace_Exit"

        Dim blnContinue As Boolean
        Dim intPos01 As Integer

20010   blnContinue = True

20020   VerifyAccountNo  ' ** Function: Below.

20030   With Me

20040     If blnAccountValidated Then
20050       .dividendShareFace.SetFocus
20060       If IsNumeric(.dividendShareFace.text) Then
20070         If CDbl(.dividendShareFace.text) > 999999999.9999 Then
20080           blnContinue = False
20090           MsgBox "Shares are too large.", vbInformation + vbOKOnly, "Invalid Entry"
20100           .dividendShareFace = vbNullString
20110           .dividendShareFace = 0
20120           Cancel = -1
20130         Else
20140           If .dividendShareFace < 0 Then
20150             blnContinue = False
20160             MsgBox "You cannot enter a negative Shares.", vbInformation + vbOKOnly, "Invalid Entry"
20170             .dividendShareFace = CStr(0)
20180             .dividendShareFace.SetFocus
20190             .Refresh
20200             Cancel = -1
20210           End If
20220         End If
20230       Else
20240         If IsNull(.dividendShareFace) = True Then
20250           blnContinue = False
20260           .dividendShareFace = 0
20270         Else
20280           blnContinue = False
20290           MsgBox "Shares must have a value.", vbInformation + vbOKOnly, "Entry Required"
20300           .dividendShareFace = vbNullString
20310           DoCmd.RunCommand acCmdSaveRecord
20320           Cancel = -1
20330           .dividendShareFace.SetFocus
20340         End If
20350       End If

20360       If blnContinue = True Then

20370         intPos01 = InStr(1, .dividendShareFace, ".", vbTextCompare)
20380         If intPos01 > 0 Then
20390           If Len(.dividendShareFace) - intPos01 > gintShareFaceDecimals Then
20400             blnContinue = False
20410             MsgBox "You can only enter Share with up to " & Trim(str(gintShareFaceDecimals)) & " decimals.", _
                    vbInformation + vbOKOnly, "Invalid Entry"
20420             .dividendShareFace = vbNullString
20430             Cancel = -1
20440           Else
20450             .dividendShareFace.Format = "#,###." & String(gintShareFaceDecimals, "0")
20460           End If
20470         Else
20480           .dividendShareFace.Format = "#,###"
20490         End If

20500         If blnContinue = True Then
20510           gstrDividendShareFace = .dividendShareFace.text
20520           If .dividendShareFace <> .dividendShareFace.OldValue Then
20530             DividendChanged True  ' ** Procedure: Below.
20540           End If
20550         End If

20560       End If
20570     Else
20580       If (gblnDividendChanged = False) Or (.NavigationButtons = True) Then
              ' ** Do nothing.
20590       Else
20600         .dividendShareFace = vbNullString
20610         gstrDividendShareFace = vbNullString
20620       End If
20630     End If

20640   End With

EXITP:
20650   Exit Sub

ERRH:
20660   THAT_PROC = THIS_PROC
20670   That_Erl = Erl
20680   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20690   Resume EXITP

End Sub

Private Sub dividendShareFace_LostFocus()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendShareFace_LostFocus"

20710   With Me

20720     VerifyAccountNo  ' ** Function: Below.

20730     If IsNull(.dividendShareFace) = False Then
20740       If .dividendShareFace > 0 Then
20750         If IsNull(.dividendPerShare) = True And IsNull(.dividendICash) = True Then
                ' ** Do nothing.
20760         ElseIf IsNull(.dividendPerShare) = True Then
20770           If .dividendICash = 0 Then
                  ' ** Do nothing.
20780           Else
20790             .dividendPerShare = (.dividendICash / .dividendShareFace)
20800             gstrDividendPerShare = CStr(.dividendICash / .dividendShareFace)
20810             .Refresh
20820           End If
20830         ElseIf IsNull(.dividendICash) = True Then
20840           If .dividendPerShare = 0 Then
                  ' ** Do nothing.
20850           Else
20860             .dividendICash = (.dividendPerShare * .dividendShareFace)
20870             gstrDividendICash = CStr(.dividendPerShare * .dividendShareFace)
20880             .Refresh
20890           End If
20900         Else
20910           If .dividendPerShare = 0 And .dividendICash > 0 Then
20920             .dividendPerShare = (.dividendICash / .dividendShareFace)
20930             gstrDividendPerShare = CStr(.dividendICash / .dividendShareFace)
20940             .Refresh
20950           ElseIf .dividendPerShare > 0 And .dividendICash = 0 Then
20960             .dividendICash = (.dividendPerShare * .dividendShareFace)
20970             gstrDividendICash = CStr(.dividendPerShare * .dividendShareFace)
20980             .Refresh
20990           Else
21000             .dividendPerShare = (.dividendICash / .dividendShareFace)
21010             gstrDividendPerShare = CStr(.dividendICash / .dividendShareFace)
21020             .Refresh
21030           End If
21040         End If
21050       End If
21060     End If

21070   End With

EXITP:
21080   Exit Sub

ERRH:
21090   THAT_PROC = THIS_PROC
21100   That_Erl = Erl
21110   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
21120   Resume EXITP

End Sub

Private Sub dividendPerShare_Change()

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendPerShare_Change"

        Dim blnContinue As Boolean
        Dim intPos01 As Integer

21210   With Me

21220     blnContinue = True

21230     lngCurrID = .dividendCurr_ID.Column(CBX_C_CURRID)

21240     If IsNumeric(Rem_Dollar(.dividendPerShare.text, lngCurrID)) Then  ' ** Module Function: modStringFuncs.
21250       If CDbl(Rem_Dollar(.dividendPerShare.text, lngCurrID)) > 999999999.9999 Then  ' ** Module Function: modStringFuncs.
21260         blnContinue = False
21270         MsgBox "Per Share is too large.", vbInformation + vbOKOnly, "Invalid Entry"
21280         .dividendPerShare = vbNullString
21290         DoCmd.CancelEvent
21300         .dividendPerShare.SetFocus
21310       End If
21320     Else
21330       If .dividendPerShare.text <> "." Then
21340         blnContinue = False
21350         MsgBox "Per Share must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
21360         .dividendPerShare = vbNullString
21370         DoCmd.CancelEvent
21380         .dividendPerShare.SetFocus
21390       End If
21400     End If

21410     If blnContinue = True Then
21420       intPos01 = InStr(1, .dividendPerShare.text, ".", vbTextCompare)
21430       If intPos01 > 0 Then
21440         If Len(.dividendPerShare.text) - intPos01 > 5 Then
21450           MsgBox "You can only enter Per Share with up to 5 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
21460           .dividendPerShare = vbNullString
21470           DoCmd.CancelEvent
21480           .dividendPerShare.SetFocus
21490         End If
21500       End If
21510     End If

21520   End With

EXITP:
21530   Exit Sub

ERRH:
21540   Select Case ERR.Number
        Case Else
21550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21560   End Select
21570   Resume EXITP

End Sub

Private Sub dividendPerShare_Exit(Cancel As Integer)

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendPerShare_Exit"

21610   With Me
21620     lngCurrID = .dividendCurr_ID.Column(CBX_C_CURRID)
21630     gstrDividendPerShare = .dividendPerShare.text
21640     gstrDividendPerShare = Rem_Dollar(gstrDividendPerShare, lngCurrID)  ' ** Module Function: modStringFuncs.
21650     If .dividendCurr_ID.Visible = True Then
21660       If .dividendPerShare_usd.Visible = False Then .dividendPerShare_usd.Visible = True
21670       .dividendPerShare_usd = Round((.dividendPerShare * .dividendCurr_ID.Column(CBX_C_RATE2)), 5)
21680     End If
21690     If .dividendPerShare <> .dividendPerShare.OldValue Then
21700       DividendChanged True  ' ** Procedure: Below.
21710     End If
21720   End With

EXITP:
21730   Exit Sub

ERRH:
21740   THAT_PROC = THIS_PROC
21750   That_Erl = Erl
21760   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
21770   Resume EXITP

End Sub

Private Sub dividendPerShare_LostFocus()

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendPerShare_LostFocus"

21810   With Me

21820     VerifyAccountNo  ' ** Function: Below.

21830     If IsNull(.dividendShareFace) = False And IsNull(.dividendPerShare) = False Then
21840       If .dividendShareFace > 0 Then
21850         lngCurrID = .dividendCurr_ID.Column(CBX_C_CURRID)
21860         If .dividendPerShare = 0 Then
21870           If IsNull(.dividendICash) = True Then
                  ' ** Do nothing.
21880           Else
21890             If .dividendICash = 0 Then
                    ' ** Do nothing.
21900             Else
21910               .dividendPerShare = (.dividendICash / .dividendShareFace)
21920               gstrDividendPerShare = CStr(.dividendICash / .dividendShareFace)
21930               If .dividendCurr_ID.Visible = True Then
21940                 If .dividendPerShare_usd.Visible = False Then .dividendPerShare_usd.Visible = True
21950                 .dividendPerShare_usd = Round(((.dividendICash * .dividendCurr_ID.Column(CBX_C_RATE2)) / .dividendShareFace), 5)
21960               End If
21970               .Refresh
21980             End If
21990           End If
22000         Else
22010           .dividendICash = Round((.dividendShareFace * .dividendPerShare), 2)
22020           gstrDividendICash = CStr(Round((.dividendShareFace * .dividendPerShare), 2))
22030           .Refresh
22040           dividendICash_AfterUpdate  ' ** Procedure: Below.
22050         End If
22060       End If
22070     Else
22080       If IsNull(.dividendShareFace) = False And IsNull(.dividendICash) = False Then
22090         If .dividendShareFace > 0 And .dividendICash > 0 Then
22100           .dividendPerShare = (.dividendICash / .dividendShareFace)
22110           gstrDividendPerShare = CStr(.dividendICash / .dividendShareFace)
22120           If .dividendCurr_ID.Visible = True Then
22130             If .dividendPerShare_usd.Visible = False Then .dividendPerShare_usd.Visible = True
22140             .dividendPerShare_usd = ((.dividendICash * .dividendCurr_ID.Column(CBX_C_RATE2)) / .dividendShareFace)
22150           End If
22160           .Refresh
22170         End If
22180       End If
22190     End If

22200   End With

EXITP:
22210   Exit Sub

ERRH:
22220   THAT_PROC = THIS_PROC
22230   That_Erl = Erl
22240   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
22250   Resume EXITP

End Sub

Private Sub dividendICash_AfterUpdate()

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendICash_AfterUpdate"

        Dim dblICash As Double
        Dim varTmp00 As Variant

22310   With Me
22320     If .dividendCurr_ID.Visible = True Then
22330       .dividendICash_usd = Null
22340       If .dividendICash_usd.Visible = False Then .dividendICash_usd.Visible = True
22350       If IsNull(.dividendICash) = False Then
22360         If .dividendICash <> 0 Then
22370           If IsNull(.dividendCurr_ID.Column(CBX_C_CURRID)) = False Then
22380             lngCurrID = .dividendCurr_ID.Column(CBX_C_CURRID)
22390             If lngCurrID <> 150& Then  ' ** USD.
22400               varTmp00 = .dividendCurr_ID.Column(CBX_C_RATE2)
22410               If IsNull(varTmp00) = False Then
22420                 If varTmp00 > 0 Then
22430                   dblICash = .dividendICash
22440                   dblICash = (dblICash * varTmp00)
22450                   dblICash = Round(dblICash, 2)
22460                   .dividendICash_usd = dblICash
22470                 End If
22480               End If
22490             End If
22500           End If
22510         End If
22520       End If
22530     End If
22540     DoEvents
22550   End With

EXITP:
22560   Exit Sub

ERRH:
22570   Select Case ERR.Number
        Case Else
22580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22590   End Select
22600   Resume EXITP

End Sub

Private Sub dividendICash_Change()

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendICash_Change"

        Dim blnContinue As Boolean
        Dim intPos01 As Integer

22710   With Me

22720     blnContinue = True

22730     lngCurrID = .dividendCurr_ID.Column(CBX_C_CURRID)

22740     If IsNumeric(Rem_Dollar(.dividendICash.text, lngCurrID)) Then  ' ** Module Function: modStringFuncs.
22750       If CDbl(Rem_Dollar(.dividendICash.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
22760         blnContinue = False
22770         MsgBox "Income Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
22780         .dividendICash = vbNullString
22790         .dividendICash_usd = Null
22800         DoCmd.CancelEvent
22810         .dividendICash.SetFocus
22820       End If
22830     Else
22840       If .dividendICash.text <> "." Then  ' ** That is, if it's not numeric, don't give the message
22850         blnContinue = False               ' ** if all they've entered so far is a decimal point.
22860         MsgBox "Income Cash must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
22870         .dividendICash = vbNullString
22880         .dividendICash_usd = Null
22890         DoCmd.CancelEvent
22900         .dividendICash.SetFocus
22910       End If
22920     End If

22930     If blnContinue = True Then
22940       intPos01 = InStr(1, .dividendICash.text, ".", vbTextCompare)
22950       If intPos01 > 0 Then
22960         If Len(.dividendICash.text) - intPos01 > 2 Then
22970           .dividendICash = CDbl(Format(.dividendICash, "0.00"))
22980           dividendICash_AfterUpdate  ' ** Procedure: Above.
22990         End If
23000       End If
23010     End If

23020   End With

EXITP:
23030   Exit Sub

ERRH:
23040   Select Case ERR.Number
        Case Else
23050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23060   End Select
23070   Resume EXITP

End Sub

Private Sub dividendICash_Exit(Cancel As Integer)

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendICash_Exit"

        Dim blnContinue As Boolean
        Dim dblTmp01 As Double
        Dim intPos01 As Integer

23110   With Me

23120     blnContinue = True

23130     lngCurrID = .dividendCurr_ID.Column(CBX_C_CURRID)

23140     If IsNumeric(.dividendICash) Then
23150       If lngCurrID <> 150& Then
23160         dblTmp01 = .dividendICash_usd
23170       Else
23180         dblTmp01 = .dividendICash
23190       End If
23200       If dblTmp01 > 999999999.99 Then 'CDbl(.dividendICash.text) > 999999999.99 Then
23210         blnContinue = False
23220         MsgBox "Income Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
23230         .dividendICash = vbNullString
23240         .dividendICash_usd = Null
23250         Cancel = -1
23260         .dividendICash.SetFocus
23270       End If
23280     Else
23290       blnContinue = False
23300       MsgBox "Income Cash must have a value.", vbInformation + vbOKOnly, "Entry Required"
23310       .dividendICash = vbNullString
23320       .dividendICash_usd = Null
23330       Cancel = -1
23340       .dividendICash.SetFocus
23350     End If

23360     If blnContinue = True Then
23370       intPos01 = InStr(1, .dividendICash, ".", vbTextCompare)
23380       If intPos01 > 0 Then
23390         If Len(.dividendICash) - intPos01 > 2 Then
23400           blnContinue = False
23410           MsgBox "You can only enter Income Cash with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
23420           .dividendICash = vbNullString
23430           .dividendICash_usd = Null
23440           Cancel = -1
23450           .dividendICash.SetFocus
23460         End If
23470       End If
23480     End If

23490     If blnContinue = True Then
23500       If .dividendICash.text = vbNullString Or .dividendICash.text = "0" Then
23510         blnContinue = False
23520         MsgBox "Income Cash must be greater than zero.", vbInformation + vbOKOnly, "Entry Required"
23530         Cancel = -1
23540         .dividendICash.SetFocus
23550       End If
23560     End If

23570     If blnContinue = True Then
23580       gstrDividendICash = Nz(.dividendICash.text, vbNullString)
23590       gstrDividendICash = Rem_Dollar(gstrDividendICash, lngCurrID)  ' ** Module Function: modStringFuncs.
23600       If .dividendICash <> .dividendICash.OldValue Then
23610         DividendChanged True  ' ** Procedure: Below.
23620       End If
23630     End If

23640   End With

EXITP:
23650   Exit Sub

ERRH:
23660   THAT_PROC = THIS_PROC
23670   That_Erl = Erl
23680   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23690   Resume EXITP

End Sub

Private Sub dividendICash_LostFocus()

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendICash_LostFocus"

23710   With Me

23720     VerifyAccountNo  ' ** Function: Below.

23730     .dividendICash_usd = Null
23740     If IsNull(.dividendICash) = False Then
23750       If .dividendICash > 0 Then
23760         If .dividendCurr_ID.Visible = True Then
23770           If .dividendICash_usd.Visible = False Then .dividendICash_usd.Visible = True
23780           .dividendICash_usd = (.dividendICash * .dividendCurr_ID.Column(CBX_C_RATE2))
23790         End If
23800         If IsNull(.dividendShareFace) = False Then
23810           If .dividendShareFace > 0 Then
23820             .dividendPerShare = (.dividendICash / .dividendShareFace)
23830             gstrDividendPerShare = CStr(.dividendICash / .dividendShareFace)
23840             If .dividendCurr_ID.Visible = True Then
23850               If .dividendICash_usd.Visible = False Then .dividendICash_usd.Visible = True
23860               .dividendPerShare_usd = Round((.dividendICash_usd / .dividendShareFace), 5)
23870             End If
23880             .Refresh
23890           End If
23900         End If
23910       End If
23920     End If

23930   End With

EXITP:
23940   Exit Sub

ERRH:
23950   THAT_PROC = THIS_PROC
23960   That_Erl = Erl
23970   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23980   Resume EXITP

End Sub

Private Sub dividendCurr_ID_cmd_DblClick(Cancel As Integer)

24000 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendCurr_ID_cmd_DblClick"

24010   With Me
24020     If IsNull(.dividendCurr_Date) = True Then
24030       .dividendCurr_Date = .dividendCurr_ID.Column(CBX_C_DATE)
24040     End If
24050     DoEvents
24060     Select Case .dividendCurr_Date.Visible
          Case True
24070       .dividendCurr_Date.Visible = False
24080     Case False
24090       .dividendCurr_Date.Visible = True
24100     End Select
24110   End With

EXITP:
24120   Exit Sub

ERRH:
24130   Select Case ERR.Number
        Case Else
24140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24150   End Select
24160   Resume EXITP

End Sub

Private Sub dividendCurr_ID_AfterUpdate()

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendCurr_ID_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, intTmp04 As Integer, intTmp05 As Integer, lngTmp06 As Long
        Dim lngX As Long

24210   With Me

24220     If IsNull(.dividendCurr_ID) = False Then

24230       lngCurrID = .dividendCurr_ID.Column(CBX_C_CURRID)

24240       blnFound = False
24250       If lngCurrID <> 150& Then
24260         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
24270           arr_varNoChar = .Parent.NoChar_Get  ' ** Form Function: frmJournal.
24280           lngNoChars = UBound(arr_varNoChar, 2) + 1&
24290         End If
24300         For lngX = 0& To (lngNoChars - 1&)
24310           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
24320             blnFound = True
24330           End If
24340         Next
24350       End If

24360       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
24370         .dividendPerShare.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
24380         .dividendPerShare.FontSize = 10
24390         .dividendPerShare.FontBold = True
24400         .dividendICash.FontName = "Courier New"
24410         .dividendICash.FontSize = 10
24420         .dividendICash.FontBold = True
24430       Case False
24440         If .dividendPerShare.FontName <> "Fixedsys" Then
24450           .dividendPerShare.FontName = "Fixedsys"
24460           .dividendPerShare.FontSize = 9
24470           .dividendPerShare.FontBold = False
24480           .dividendICash.FontName = "Fixedsys"
24490           .dividendICash.FontSize = 9
24500           .dividendICash.FontBold = False
24510         End If
24520       End Select

24530       If lngCurrID = 150& Then  ' ** USD.
24540         strTmp01 = "Currency"
24550         intTmp04 = 2
24560         strTmp02 = "Currency"
24570         intTmp05 = 4
24580         lngTmp06 = CLR_WHT
24590       Else
24600         strTmp03 = .dividendCurr_ID.Column(CBX_C_SYM)  ' ** Symbol.
24610         If .dividendCurr_ID.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
24620           intTmp04 = 4
24630         Else
24640           intTmp04 = .dividendCurr_ID.Column(CBX_C_DEC)  ' ** Decimal places.
24650         End If
24660         intTmp05 = 4
24670         intLen = Len(strTmp03)
24680         Select Case intLen
              Case 1
                'strTmp03 = strTmp03
24690         Case 2
24700           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2)
24710         Case 3
24720           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3)
24730         Case 4
24740           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & "\" & Mid(strTmp03, 4)
24750         Case 5
24760           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & _
                  "\" & Mid(strTmp03, 4, 1) & "\" & Mid(strTmp03, 5)
24770         End Select
24780         Select Case intTmp04
              Case 0
24790           strTmp01 = strTmp03 & "#,##0;(" & strTmp03 & "#,##0);" & strTmp03 & "0;" & strTmp03 & "0"
24800           strTmp02 = strTmp03 & "#,##0.0000;(" & strTmp03 & "#,##0.0000);" & strTmp03 & "0.0000;" & strTmp03 & "0.0000"
24810         Case 1  ' ** None have this currently.
24820           strTmp01 = strTmp03 & "#,##0.0;(" & strTmp03 & "#,##0.0);" & strTmp03 & "0.0;" & strTmp03 & "0.0"
24830           strTmp02 = strTmp03 & "#,##0.0000;(" & strTmp03 & "#,##0.0000);" & strTmp03 & "0.0000;" & strTmp03 & "0.0000"
24840         Case 2
24850           strTmp01 = strTmp03 & "#,##0.00;(" & strTmp03 & "#,##0.00);" & strTmp03 & "0.00;" & strTmp03 & "0.00"
24860           strTmp02 = strTmp03 & "#,##0.0000;(" & strTmp03 & "#,##0.0000);" & strTmp03 & "0.0000;" & strTmp03 & "0.0000"
24870         Case 3
24880           strTmp01 = strTmp03 & "#,##0.000;(" & strTmp03 & "#,##0.000);" & strTmp03 & "0.000;" & strTmp03 & "0.000"
24890           strTmp02 = strTmp03 & "#,##0.0000;(" & strTmp03 & "#,##0.0000);" & strTmp03 & "0.0000;" & strTmp03 & "0.0000"
24900         End Select
24910         lngTmp06 = CLR_VLTGRN
24920       End If
24930       .dividendICash.Format = strTmp01
24940       .dividendICash.DecimalPlaces = intTmp04
24950       .dividendICash.BackColor = lngTmp06
24960       .dividendPerShare.Format = strTmp02
24970       .dividendPerShare.DecimalPlaces = intTmp05
24980       .dividendPerShare.BackColor = lngTmp06

24990     End If

25000     If lngCurrID <> .dividendCurr_ID.OldValue Then
25010       DividendChanged True  ' ** Procedure: Below.
25020     End If

25030   End With

EXITP:
25040   Exit Sub

ERRH:
25050   THAT_PROC = THIS_PROC
25060   That_Erl = Erl
25070   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
25080   Resume EXITP

End Sub

Private Sub dividendCurr_ID_NotInList(NewData As String, Response As Integer)

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendCurr_ID_NotInList"

25110   With Me
25120     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
25130     .dividendCurr_ID.SetFocus
25140     Response = acDataErrContinue
25150   End With

EXITP:
25160   Exit Sub

ERRH:
25170   THAT_PROC = THIS_PROC
25180   That_Erl = Erl
25190   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
25200   Resume EXITP

End Sub

Private Sub dividendCurr_ID_Exit(Cancel As Integer)

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendCurr_ID_Exit"

25310   With Me
25320     If .dividendCurr_ID <> .dividendCurr_ID.OldValue Then
25330       DividendChanged True  ' ** Procedure: Below.
25340     End If
25350   End With

EXITP:
25360   Exit Sub

ERRH:
25370   THAT_PROC = THIS_PROC
25380   That_Erl = Erl
25390   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
25400   Resume EXITP

End Sub

Private Sub dividendDescription_AfterUpdate()

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendDescription_AfterUpdate"

25510   With Me
25520     If .dividendDescription <> .dividendDescription.OldValue Then
25530       DividendChanged True  ' ** Procedure: Below.
25540     End If
25550   End With

EXITP:
25560   Exit Sub

ERRH:
25570   THAT_PROC = THIS_PROC
25580   That_Erl = Erl
25590   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
25600   Resume EXITP

End Sub

Private Sub dividendDescription_KeyDown(KeyCode As Integer, Shift As Integer)

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendDescription_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25720   intShiftDown = (Shift And acShiftMask) > 0
25730   intAltDown = (Shift And acAltMask) > 0
25740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25770       With Me
25780         intRetVal = 0
25790         If .tglDividendReinvest.Visible = True And .tglDividendReinvest.Enabled = True Then
25800           .tglDividendReinvest.SetFocus
25810         ElseIf .cmbRevenueCodes.Visible = True And .cmbRevenueCodes.Enabled = True Then
25820           .cmbRevenueCodes.SetFocus
25830         ElseIf .cmbTaxCodes.Visible = True And .cmbTaxCodes.Enabled = True Then
25840           .cmbTaxCodes.SetFocus
25850         ElseIf .cmdDividendOK.Enabled = True Then
25860           .cmdDividendOK.SetFocus
25870         ElseIf .cmdDividendCancel.Enabled = True Then
25880           .cmdDividendCancel.SetFocus
25890         ElseIf .cmdDividendClose.Enabled = True Then
25900           .cmdDividendClose.SetFocus
25910         ElseIf .dividendAccountNo.Enabled = True Then
25920           .dividendAccountNo.SetFocus
25930         Else
25940           .dividendTransDate.SetFocus
25950         End If
25960       End With
25970     End Select
25980   End If

        ' ** Shift keys.
25990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26000     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26010       With Me
26020         intRetVal = 0
26030         .dividendICash.SetFocus
26040       End With
26050     End Select
26060   End If

EXITP:
26070   KeyCode = intRetVal
26080   Exit Sub

ERRH:
26090   intRetVal = 0
26100   THAT_PROC = THIS_PROC
26110   That_Erl = Erl
26120   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
26130   Resume EXITP

End Sub

Private Sub dividendDescription_LostFocus()

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "dividendDescription_LostFocus"

26210   VerifyAccountNo  ' ** Function: Below.

EXITP:
26220   Exit Sub

ERRH:
26230   THAT_PROC = THIS_PROC
26240   That_Erl = Erl
26250   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
26260   Resume EXITP

End Sub

Private Sub tglDividendReinvest_GotFocus()

26300 On Error GoTo ERRH

        Const THIS_PROC As String = "tglDividendReinvest_GotFocus"

26310   With Me
26320     blnDividendReinvest_Focus = True
26330     Select Case .posted
          Case True
26340       .tglDividendReinvest_true_raised_focus_dots_img.Visible = True
26350       .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = False
26360     Case False
26370       .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = True
26380       .tglDividendReinvest_true_raised_focus_dots_img.Visible = False
26390     End Select
26400     .tglDividendReinvest_false_raised_img.Visible = False
26410     .tglDividendReinvest_false_raised_focus_img.Visible = False
26420     .tglDividendReinvest_false_raised_focus_dots_img.Visible = False
26430     .tglDividendReinvest_false_sunken_focus_dots_img.Visible = False
26440     .tglDividendReinvest_false_raised_img_dis.Visible = False
26450     .tglDividendReinvest_true_raised_img.Visible = False
26460     .tglDividendReinvest_true_raised_focus_img.Visible = False
26470     .tglDividendReinvest_true_sunken_focus_dots_img.Visible = False
26480     .tglDividendReinvest_true_raised_img_dis.Visible = False
26490   End With

EXITP:
26500   Exit Sub

ERRH:
26510   THAT_PROC = THIS_PROC
26520   That_Erl = Erl
26530   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
26540   Resume EXITP

End Sub

Private Sub tglDividendReinvest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "tglDividendReinvest_MouseDown"

26610   With Me
26620     blnDividendReinvest_MouseDown = True
26630     Select Case .posted
          Case True
26640       .tglDividendReinvest_true_sunken_focus_dots_img.Visible = True
26650       .tglDividendReinvest_false_sunken_focus_dots_img.Visible = False
26660     Case False
26670       .tglDividendReinvest_false_sunken_focus_dots_img.Visible = True
26680       .tglDividendReinvest_true_sunken_focus_dots_img.Visible = False
26690     End Select
26700     .tglDividendReinvest_false_raised_img.Visible = False
26710     .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = False
26720     .tglDividendReinvest_false_raised_focus_img.Visible = False
26730     .tglDividendReinvest_false_raised_focus_dots_img.Visible = False
26740     .tglDividendReinvest_false_raised_img_dis.Visible = False
26750     .tglDividendReinvest_true_raised_img.Visible = False
26760     .tglDividendReinvest_true_raised_focus_img.Visible = False
26770     .tglDividendReinvest_true_raised_focus_dots_img.Visible = False
26780     .tglDividendReinvest_true_raised_img_dis.Visible = False
26790   End With

EXITP:
26800   Exit Sub

ERRH:
26810   THAT_PROC = THIS_PROC
26820   That_Erl = Erl
26830   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
26840   Resume EXITP

End Sub

Public Sub tglDividendReinvest_Click()
' ** I'm going to try using [posted] as a signal (since it's not used for anything else).

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "tglDividendReinvest_Click"

26910   With Me
          ' ** Do I need to do anything with the images, or will that be handled by the other events?
26920     Select Case .posted
          Case True  ' ** If it's True, flip to False, and vice versa.
26930       blnGoToDivReinvest = False
26940       .posted = False
26950       DoCmd.RunCommand acCmdSaveRecord
26960     Case False
26970       blnGoToDivReinvest = True
26980       .posted = True
26990       DoCmd.RunCommand acCmdSaveRecord
27000     End Select
27010     DividendChanged True  ' ** Procedure: Below.
27020     DoEvents
27030   End With

EXITP:
27040   Exit Sub

ERRH:
27050   THAT_PROC = THIS_PROC
27060   That_Erl = Erl
27070   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
27080   Resume EXITP

End Sub

Private Sub tglDividendReinvest_KeyDown(KeyCode As Integer, Shift As Integer)

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "tglDividendReinvest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27120   intShiftDown = (Shift And acShiftMask) > 0
27130   intAltDown = (Shift And acAltMask) > 0
27140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27170       With Me
27180         intRetVal = 0
27190         If .cmbRevenueCodes.Enabled = True Then
27200           .cmbRevenueCodes.SetFocus
27210         ElseIf .cmdDividendOK.Enabled = True Then
27220           .cmdDividendOK.SetFocus
27230         ElseIf .cmdDividendCancel.Enabled = True Then
27240           .cmdDividendCancel.SetFocus
27250         ElseIf .cmdDividendClose.Enabled = True Then
27260           .cmdDividendClose.SetFocus
27270         ElseIf .dividendAccountNo.Enabled = True Then
27280           .dividendAccountNo.SetFocus
27290         Else
27300           .dividendTransDate.SetFocus
27310         End If
27320       End With
27330     End Select
27340   End If

        ' ** Shift keys.
27350   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27370       With Me
27380         intRetVal = 0
27390         .dividendDescription.SetFocus
27400       End With
27410     End Select
27420   End If

EXITP:
27430   KeyCode = intRetVal
27440   Exit Sub

ERRH:
27450   intRetVal = 0
27460   THAT_PROC = THIS_PROC
27470   That_Erl = Erl
27480   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
27490   Resume EXITP

End Sub

Private Sub tglDividendReinvest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "tglDividendReinvest_MouseMove"

27510   With Me
27520     If blnDividendReinvest_MouseDown = False Then
27530       Select Case .posted
            Case True
27540         Select Case blnDividendReinvest_Focus
              Case True
27550           .tglDividendReinvest_true_raised_focus_dots_img.Visible = True
27560           .tglDividendReinvest_true_raised_focus_img.Visible = False
27570         Case False
27580           .tglDividendReinvest_true_raised_focus_img.Visible = True
27590           .tglDividendReinvest_true_raised_focus_dots_img.Visible = False
27600         End Select
27610         .tglDividendReinvest_false_raised_focus_img.Visible = False
27620         .tglDividendReinvest_false_raised_focus_dots_img.Visible = False
27630       Case False
27640         Select Case blnDividendReinvest_Focus
              Case True
27650           .tglDividendReinvest_false_raised_focus_dots_img.Visible = True
27660           .tglDividendReinvest_false_raised_focus_img.Visible = False
27670         Case False
27680           .tglDividendReinvest_false_raised_focus_img.Visible = True
27690           .tglDividendReinvest_false_raised_focus_dots_img.Visible = False
27700         End Select
27710         .tglDividendReinvest_true_raised_focus_img.Visible = False
27720         .tglDividendReinvest_true_raised_focus_dots_img.Visible = False
27730       End Select
27740       .tglDividendReinvest_false_raised_img.Visible = False
27750       .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = False
27760       .tglDividendReinvest_false_raised_img_dis.Visible = False
27770       .tglDividendReinvest_false_sunken_focus_dots_img.Visible = False
27780       .tglDividendReinvest_true_raised_img.Visible = False
27790       .tglDividendReinvest_true_sunken_focus_dots_img.Visible = False
27800       .tglDividendReinvest_true_raised_img_dis.Visible = False
27810     End If
27820   End With

EXITP:
27830   Exit Sub

ERRH:
27840   THAT_PROC = THIS_PROC
27850   That_Erl = Erl
27860   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
27870   Resume EXITP

End Sub

Private Sub tglDividendReinvest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "tglDividendReinvest_MouseUp"

27910   With Me
27920     Select Case .posted
          Case True
27930       .tglDividendReinvest_true_raised_focus_dots_img.Visible = True
27940       .tglDividendReinvest_false_raised_focus_dots_img.Visible = False
27950     Case False
27960       .tglDividendReinvest_false_raised_focus_dots_img.Visible = True
27970       .tglDividendReinvest_true_raised_focus_dots_img.Visible = False
27980     End Select
27990     .tglDividendReinvest_false_raised_img.Visible = False
28000     .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = False
28010     .tglDividendReinvest_false_raised_focus_img.Visible = False
28020     .tglDividendReinvest_false_raised_img_dis.Visible = False
28030     .tglDividendReinvest_false_sunken_focus_dots_img.Visible = False
28040     .tglDividendReinvest_true_raised_img.Visible = False
28050     .tglDividendReinvest_true_raised_focus_img.Visible = False
28060     .tglDividendReinvest_true_sunken_focus_dots_img.Visible = False
28070     .tglDividendReinvest_true_raised_img_dis.Visible = False
28080     blnDividendReinvest_MouseDown = False
28090   End With

EXITP:
28100   Exit Sub

ERRH:
28110   THAT_PROC = THIS_PROC
28120   That_Erl = Erl
28130   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
28140   Resume EXITP

End Sub

Private Sub tglDividendReinvest_LostFocus()

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "tglDividendReinvest_LostFocus"

28210   With Me
28220     Select Case .posted
          Case True
28230       .tglDividendReinvest_true_raised_img.Visible = True
28240       .tglDividendReinvest_false_raised_img.Visible = False
28250     Case False
28260       .tglDividendReinvest_false_raised_img.Visible = True
28270       .tglDividendReinvest_true_raised_img.Visible = False
28280     End Select
28290     .tglDividendReinvest_false_raised_semifocus_dots_img.Visible = False
28300     .tglDividendReinvest_false_raised_focus_img.Visible = False
28310     .tglDividendReinvest_false_raised_focus_dots_img.Visible = False
28320     .tglDividendReinvest_false_sunken_focus_dots_img.Visible = False
28330     .tglDividendReinvest_false_raised_img_dis.Visible = False
28340     .tglDividendReinvest_true_raised_focus_img.Visible = False
28350     .tglDividendReinvest_true_raised_focus_dots_img.Visible = False
28360     .tglDividendReinvest_true_sunken_focus_dots_img.Visible = False
28370     .tglDividendReinvest_true_raised_img_dis.Visible = False
28380     blnDividendReinvest_Focus = False
28390   End With

EXITP:
28400   Exit Sub

ERRH:
28410   THAT_PROC = THIS_PROC
28420   That_Erl = Erl
28430   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
28440   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Enter()

28500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Enter"

28510   With Me
28520     Select Case IsNull(.cmbRevenueCodes)
          Case True
28530       ExpandCombo .cmbRevenueCodes  ' ** Module Procedure: modUtilities.
28540     Case False
28550       If .cmbRevenueCodes = REVID_INC Or .cmbRevenueCodes = REVID_EXP Then  ' ** Just for the Unspecified's.
28560         ExpandCombo .cmbRevenueCodes  ' ** Module Procedure: modUtilities.
28570       End If
28580     End Select
28590   End With

EXITP:
28600   Exit Sub

ERRH:
28610   Select Case ERR.Number
        Case Else
28620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28630   End Select
28640   Resume EXITP

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

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Change"

28710   With Me
28720     If .cmbRevenueCodes <> .cmbRevenueCodes.OldValue Then
28730       DividendChanged True  ' ** Procedure: Below.
28740     End If
28750   End With

EXITP:
28760   Exit Sub

ERRH:
28770   Select Case ERR.Number
        Case Else
28780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28790   End Select
28800   Resume EXITP

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

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_AfterUpdate"

        Dim strRevCode As String, lngTaxcode As Long

        ' ** Make sure tax code matches selected revenue code.
28910   With Me
28920     If IsNull(.cmbRevenueCodes) = False Then
28930       If .cmbRevenueCodes > 0 Then
28940         If gblnLinkRevTaxCodes = True Then
28950           strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
28960           If IsNull(.cmbTaxCodes) = True Then
28970             If strRevCode = "I" Then
28980               .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
28990             ElseIf strRevCode = "E" Then
29000               .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
29010             End If
29020           Else
29030             lngTaxcode = .cmbTaxCodes
29040             If lngTaxcode = 0& Then
29050               If strRevCode = "I" Then
29060                 .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
29070               ElseIf strRevCode = "E" Then
29080                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
29090               End If
29100             Else
29110               If strRevCode = "I" And .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
                      'MsgBox "This will change..."
29120                 .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
29130               ElseIf strRevCode = "E" And .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
                      'MsgBox "This will change..."
29140                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
29150               End If
29160             End If
29170           End If
29180         End If
29190       Else
29200         .cmbRevenueCodes = REVID_ORDDIV  ' ** Ordinary Dividend.
29210       End If
29220     Else
29230       .cmbRevenueCodes = REVID_ORDDIV  ' ** Ordinary Dividend.
29240     End If
29250   End With

EXITP:
29260   Exit Sub

ERRH:
29270   Select Case ERR.Number
        Case Else
29280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29290   End Select
29300   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_LostFocus()

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_LostFocus"

29410   VerifyAccountNo  ' ** Function: Below.

EXITP:
29420   Exit Sub

ERRH:
29430   THAT_PROC = THIS_PROC
29440   That_Erl = Erl
29450   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
29460   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Enter()

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Enter"

29510   With Me
29520     Select Case IsNull(.cmbTaxCodes)
          Case True
29530       ExpandCombo .cmbTaxCodes  ' ** Module Procedure: modUtilities.
29540     Case False
            'If .cmbTaxCodes = TAXID_INC Or .cmbTaxCodes = TAXID_DED Then
29550       ExpandCombo .cmbTaxCodes  ' ** Module Procedure: modUtilities.
            'End If
29560     End Select
29570   End With

EXITP:
29580   Exit Sub

ERRH:
29590   Select Case ERR.Number
        Case Else
29600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29610   End Select
29620   Resume EXITP

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

29700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Change"

29710   With Me
29720     If .cmbTaxCodes <> .cmbTaxCodes.OldValue Then
29730       DividendChanged True  ' ** Procedure: Below.
29740     End If
29750   End With

EXITP:
29760   Exit Sub

ERRH:
29770   THAT_PROC = THIS_PROC
29780   That_Erl = Erl
29790   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
29800   Resume EXITP

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

29900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_AfterUpdate"

        Dim strRevCode As String, lngTaxcode As Long

        ' ** Make sure revenue code matches selected tax code.
29910   With Me
29920     If IsNull(.cmbTaxCodes) = False Then
29930       If .cmbTaxCodes > 0& Then
29940         If gblnLinkRevTaxCodes = True Then
29950           lngTaxcode = .cmbTaxCodes
29960           strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
29970           If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
29980             If IsNull(.cmbRevenueCodes) = True Then
29990               .cmbRevenueCodes = REVID_ORDDIV  ' ** Ordinary Dividend.
30000             Else
30010               If .cmbRevenueCodes = 0 Then
30020                 .cmbRevenueCodes = REVID_ORDDIV  ' ** Ordinary Dividend.
30030               Else
30040                 If strRevCode = "E" Then
                        'MsgBox "This will change..."
30050                   .cmbRevenueCodes = REVID_ORDDIV  ' ** Ordinary Dividend.
30060                 End If
30070               End If
30080             End If
30090           ElseIf .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
30100             If IsNull(.cmbRevenueCodes) = True Then
30110               .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
30120             Else
30130               If .cmbRevenueCodes = 0 Then
30140                 .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
30150               Else
30160                 If strRevCode = "I" Then
                        'MsgBox "This will change..."
30170                   .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
30180                 End If
30190               End If
30200             End If
30210           End If
30220         End If
30230       Else
30240         .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
30250       End If
30260     Else
30270       .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
30280     End If
30290   End With

EXITP:
30300   Exit Sub

ERRH:
30310   Select Case ERR.Number
        Case Else
30320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30330   End Select
30340   Resume EXITP

End Sub

Private Sub cmbTaxCodes_LostFocus()

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_LostFocus"

30410   VerifyAccountNo  ' ** Function: Below.

EXITP:
30420   Exit Sub

ERRH:
30430   THAT_PROC = THIS_PROC
30440   That_Erl = Erl
30450   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
30460   Resume EXITP

End Sub

Public Sub DividendChanged(blnChanged As Boolean)

30500 On Error GoTo ERRH

        Const THIS_PROC As String = "DividendChanged"

30510   With Me
30520     Select Case blnChanged
          Case True
30530       gblnDividendChanged = True
30540       .NavigationButtons = False
30550       DoCmd.SelectObject acForm, .Parent.Name, False
30560       With .Parent
30570         .FocusHolder.SetFocus
30580         .opgJournal.Enabled = False
30590         .opgJournal_optDividend_lbl_box.Visible = True
30600         .cmdSwitch.Enabled = False
30610         .cmdSwitch_raised_img_dis.Visible = True
30620         .cmdSwitch_raised_img.Visible = False
30630         .cmdSwitch_raised_semifocus_dots_img.Visible = False
30640         .cmdSwitch_raised_focus_img.Visible = False
30650         .cmdSwitch_raised_focus_dots_img.Visible = False
30660         .cmdSwitch_sunken_focus_dots_img.Visible = False
30670         .frmJournal_Sub1_Dividend.SetFocus
30680       End With
30690       .cmdDividendClose.Enabled = False
30700       .cmdDividendOK.Enabled = True
30710       .cmdDividendCancel.Enabled = True
30720       .Parent.NavVis False  ' ** Form Procedure: frmJournal.
30730     Case False
30740       gblnDividendChanged = False
30750       .NavigationButtons = True
30760       DoCmd.SelectObject acForm, .Parent.Name, False
30770       With .Parent
30780         .opgJournal.Enabled = True
30790         .opgJournal_optDividend_lbl_box.Visible = False
30800         .cmdSwitch.Enabled = True
30810         .cmdSwitch_raised_img.Visible = True
30820         .cmdSwitch_raised_img_dis.Visible = False
30830         .cmdSwitch_raised_semifocus_dots_img.Visible = False
30840         .cmdSwitch_raised_focus_img.Visible = False
30850         .cmdSwitch_raised_focus_dots_img.Visible = False
30860         .cmdSwitch_sunken_focus_dots_img.Visible = False
30870       End With
30880       .cmdDividendClose.Enabled = True
30890       .cmdDividendOK.Enabled = False
30900       .cmdDividendCancel.Enabled = False
30910       .Parent.NavVis True  ' ** Form Procedure: frmJournal.
30920     End Select
30930   End With

EXITP:
30940   Exit Sub

ERRH:
30950   THAT_PROC = THIS_PROC
30960   That_Erl = Erl
30970   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
30980   Resume EXITP

End Sub

Public Function VerifyAccountNo() As Boolean

31000 On Error GoTo ERRH

        Const THIS_PROC As String = "VerifyAccountNo"

        Dim strAccountNo As String
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

31010   blnRetVal = True

31020   With Me
31030     If ((gblnDividendValidated = False And gblnDividendChanged = True) Or blnCheckingAccount = True) And (lngErrCnt < 3&) Then
31040       blnAccountValidated = False     ' ** Unless proven otherwise.
31050       Select Case IsNull(.dividendAccountNo)
            Case True
31060         blnRetVal = False
31070         lngErrCnt = lngErrCnt + 1&
31080         .dividendAccountNo.SetFocus
31090         MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Entry Required"
31100         gblnDividendValidated = False
31110       Case False
              ' ** Check to see if it is a valid account.
31120         lngErrCnt = 0&
31130         Select Case gblnAccountNoWithType
              Case True
31140           strAccountNo = Left(.dividendAccountNo, 2) & Mid(.dividendAccountNo, 3)
31150         Case False
31160           strAccountNo = .dividendAccountNo
31170         End Select
31180         varTmp00 = DLookup("[accountno]", "account", "[accountno] = '" & strAccountNo & "' And ([dateclosed] Is Null)")
31190         Select Case IsNull(varTmp00)
              Case True
31200           blnRetVal = False
31210           MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Invalid Entry"
31220           gblnDividendValidated = False
31230         Case False
31240           blnAccountValidated = True
31250           gblnDividendValidated = True
31260           blnStartTrans = False
31270         End Select
31280       End Select
31290     Else
31300       lngErrCnt = 0&
31310     End If
31320   End With

EXITP:
31330   VerifyAccountNo = blnRetVal
31340   Exit Function

ERRH:
31350   blnRetVal = False
31360   THAT_PROC = THIS_PROC
31370   That_Erl = Erl
31380   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
31390   Resume EXITP

End Function

Public Sub Map_NewRec()

31400 On Error GoTo ERRH

        Const THIS_PROC As String = "Map_NewRec"

31410   With Me
31420 On Error Resume Next
31430     DoCmd.RunCommand acCmdRecordsGoToNew
31440     If ERR.Number <> 0 Then
31450       blnMapNewRec = True
31460       .TimerInterval = 500&
31470     End If
31480 On Error GoTo ERRH
31490   End With

EXITP:
31500   Exit Sub

ERRH:
31510   THAT_PROC = THIS_PROC
31520   That_Erl = Erl
31530   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
31540   Resume EXITP

End Sub

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

31600 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

31610   Select Case intWhere
        Case acCmdRecordsGoToFirst
31620     DoCmd.RunCommand acCmdRecordsGoToFirst
31630   Case acCmdRecordsGoToLast
31640     DoCmd.RunCommand acCmdRecordsGoToLast
31650   Case acCmdRecordsGoToPrevious
31660     DoCmd.RunCommand acCmdRecordsGoToPrevious
31670   Case acCmdRecordsGoToNext
31680     DoCmd.RunCommand acCmdRecordsGoToNext
31690   Case acCmdRecordsGoToNew
31700     DoCmd.RunCommand acCmdRecordsGoToNew
31710   Case Else
31720     If IsMissing(varID) = False Then
31730       With Me
31740         Set rst = .RecordsetClone
31750         With rst
31760           .FindFirst "[dividendID] = " & CStr(varID)
31770           If .NoMatch = False Then
31780             Me.Bookmark = .Bookmark
31790           End If
31800           .Close
31810         End With
31820       End With
31830     End If
31840   End Select

EXITP:
31850   Set rst = Nothing
31860   Exit Sub

ERRH:
31870   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
31880   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
31890   Case Else
31900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31910   End Select
31920   Resume EXITP

End Sub
