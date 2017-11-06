VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournal_Sub2_Interest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournal_Sub2_Interest"

'VGC 10/01/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Map:              A {cmdInterestMap}
' **   Close:            C {cmdInterestClose}
' **   Dividend:         D {opgJournal_optDividend on frmJournal}
' **   Cancel:           E {cmdInterestCancel}
' **   Show All Assets:  H {chkShowAllAssets}
' **   Interest:         I {opgJournal_optInterest on frmJournal}
' **   Miscellaneous:    M {opgJournal_optMisc on frmJournal}
' **   Inc/Exp Codes:    N {cmbRevenueCodes}
' **   OK:               O {cmdInterestOK}
' **   Purchase:         P {opgJournal_optPurchase on frmJournal}
' **   Sale:             S {opgJournal_optSale on frmJournal}
' **   Tax Codes:        T {cmbTaxCodes}
' **   Exit:             X {cmdInterestClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Next Record:      N {MoveRec}
' **   Previous Record:  P {MoveRec}
' **   Switch:           T {cmdSwitch on frmJournal}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Mark Reinvest:    R {tglInterestReinvest}
' **   Close Timer:      T {Form_KeyDown}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}

' ** Interest is always INCOME for Tax Codes.

' ** IsAverage used for chkShowAllAssets.

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

' ** Combo box column constants: interestAssetNo.
'Private Const CBX_A_DESC   As Integer = 0  'totdesc
'Private Const CBX_A_CUSIP  As Integer = 1  'cusip
Private Const CBX_A_ASTNO  As Integer = 2  'assetno
'Private Const CBX_A_TYPE   As Integer = 3  'assettype
'Private Const CBX_A_DUE    As Integer = 4  'due
'Private Const CBX_A_RATE   As Integer = 5  'rate
Private Const CBX_A_TAX    As Integer = 6  'taxcode
Private Const CBX_A_CURRID As Integer = 7  'curr_id

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

' ** Combo box column constants: interestCurr_ID.
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
Private blnInterestReinvest_Focus As Boolean, blnInterestReinvest_MouseDown As Boolean

Private lngCurrID As Long, blnMapNewRec As Boolean
Private blnCmdButton As Boolean, blnGoToIntReinvest As Boolean
Private blnAccountNoErr As Boolean, blnStartTrans As Boolean, blnTmpEdit01Used As Boolean
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
180           blnAccountNoErr = False: blnStartTrans = False: blnTmpEdit01Used = False
190           gblnSwitchTo = False

200           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
210           blnInterestReinvest_Focus = False: blnInterestReinvest_MouseDown = False
220           blnMapNewRec = False

230           If .cmbRevenueCodes.DefaultValue = vbNullString Then
240             .cmbRevenueCodes.DefaultValue = REVID_INTINC  ' ** Interest Income.
250           End If

260           .cmdLock.Enabled = False
270           .cmdLock_open_raised_img.Visible = False
280           .cmdLock_open_raised_img_dis.Visible = False
290           .cmdLock_closed_raised_img.Visible = False

300         Else
310           Cancel = -1
320         End If
330       Else
340         Cancel = -1
350       End If
360     End With

370     If Cancel = -1 Then
380       Beep
390       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
400     End If

EXITP:
410     Exit Sub

ERRH:
420     Select Case ERR.Number
        Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

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

550       blnCheckingAccount = False     ' ** Default.
560       blnExitingAssetNo = False      ' ** Default.
570       blnAccountValidated = False    ' ** Default.
580       blnCmdButton = False  ' ** Set the switch to false until they push the button.
590       blnGoToIntReinvest = False

600       datPostingDate = GetPostDate  ' ** Module Function: modUtilities.

          ' ** Set the form's default JournalType.
610       .interestType.DefaultValue = Chr(34) & "Interest" & Chr(34)

          ' ** Do this whether or not Currency is visible.
620       CLR_DISABLED_FG = CLR_DKGRY
630       CLR_DISABLED_BG = CLR_LTTEAL

640       .interestCurr_ID.Enabled = False
650       .interestCurr_ID.Locked = True
660       .interestCurr_ID.ForeColor = CLR_DISABLED_FG
670       .interestCurr_ID.BackColor = CLR_VLTGRN
680       interestCurr_ID_AfterUpdate  ' ** Procedure: Below.

          ' ** Set the default Posting Date (transdate), entered by the user on opening.
690       .interestTransDate.DefaultValue = "#" & CDate(Format(datPostingDate, "mm/dd/yyyy")) & "#"

700       gstrFormQuerySpec = .Parent.Name  ' ** Make sure this is set for the assetno combo box.

710       gblnInterestValidated = False
720       gblnInterestChanged = False

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
840       .cmdInterestClose.Enabled = True
850       .cmdInterestOK.Enabled = False
860       .cmdInterestCancel.Enabled = False
870       .Parent.NavVis True  ' ** Form Procedure: frmJournal.

880       If .NewRecord = True Then
890         .tglInterestReinvest.Enabled = False
900         .tglInterestReinvest_false_raised_img_dis.Visible = True
910         .tglInterestReinvest_false_raised_img.Visible = False
920         .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = False
930         .tglInterestReinvest_false_raised_focus_img.Visible = False
940         .tglInterestReinvest_false_raised_focus_dots_img.Visible = False
950         .tglInterestReinvest_false_sunken_focus_dots_img.Visible = False
960         .tglInterestReinvest_true_raised_img.Visible = False
970         .tglInterestReinvest_true_raised_focus_img.Visible = False
980         .tglInterestReinvest_true_raised_focus_dots_img.Visible = False
990         .tglInterestReinvest_true_sunken_focus_dots_img.Visible = False
1000        .tglInterestReinvest_true_raised_img_dis.Visible = False
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

1160    End With

EXITP:
1170    Exit Sub

ERRH:
1180    Select Case ERR.Number
        Case Else
1190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1200    End Select
1210    Resume EXITP

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

1340      blnCheckingAccount = False     ' ** Default.
1350      blnExitingAssetNo = False      ' ** Default.
1360      blnAccountValidated = False    ' ** Default.
1370      blnStartTrans = False
1380      lngErrCnt = 0&

1390      .cmbAccountHelper = Null

1400      If .interestCurr_ID.Visible = True Then
1410        If .interestCurr_Date.Visible = True Then .interestCurr_Date.Visible = False
1420        .interestICash_usd = Null
1430        .interestICash_usd.Visible = False
1440        .interestICash.Format = "Currency"
1450        .interestICash.DecimalPlaces = 2
1460        .interestICash.BackColor = CLR_WHT
1470        If lngNoChars = 0& Or IsEmpty(arr_varNoChar) Then
1480          arr_varNoChar = .Parent.NoChar_Get  ' ** Form Function: frmJournal.
1490          lngNoChars = UBound(arr_varNoChar, 2) + 1&
1500        End If
1510      End If

1520      gstrFormQuerySpec = .Parent.Name  ' ** Make sure this is set for the assetno combo box.

1530      Select Case IsNull(.interestAccountNo_Data)
          Case True
1540        .interestAccountNo = vbNullString
1550        .interestAccountNo.Enabled = True
1560        .interestAccountNo.BorderColor = CLR_LTBLU2
1570        .interestAccountNo.BackStyle = acBackStyleNormal
1580        .interestAccountNo_lbl.BackStyle = acBackStyleNormal
1590        .interestAccountNo_lbl_box.Visible = False
1600        .cmdLock.Enabled = False
1610        .cmdLock_open_raised_img_dis.Visible = True
1620        .cmdLock_open_raised_img.Visible = False
1630        .cmdLock_closed_raised_img.Visible = False
1640        .cmbAccountHelper.Enabled = True
1650        .cmbAccountHelper.BorderColor = CLR_LTBLU2
1660        .cmbAccountHelper.BackStyle = acBackStyleNormal
1670        .interestAccountNo.SetFocus
1680        .cmdInterestMap.Enabled = True
1690        blnAccountNoErr = False
1700      Case False
1710        .interestAccountNo = .interestAccountNo_Data
1720        gstrInterestAccountNumber = .interestAccountNo
1730        .interestTransDate.SetFocus  ' ** Make sure it's not on AccountNo before disabling.
1740        .interestAccountNo.Enabled = False
1750        .interestAccountNo.BorderColor = WIN_CLR_DISR
1760        .interestAccountNo.BackStyle = acBackStyleTransparent
1770        .interestAccountNo_lbl.BackStyle = acBackStyleTransparent
1780        .interestAccountNo_lbl_box.Visible = True
1790        .cmdLock.Enabled = True
1800        .cmdLock_open_raised_img_dis.Visible = False
1810        .cmdLock_open_raised_img.Visible = False
1820        .cmdLock_closed_raised_img.Visible = True
1830        .cmbAccountHelper.Enabled = False
1840        .cmbAccountHelper.BorderColor = WIN_CLR_DISR
1850        .cmbAccountHelper.BackStyle = acBackStyleTransparent
1860        .cmdInterestMap.Enabled = False
1870      End Select

          ' ** Make sure validation is reset.
1880      gblnInterestValidated = False

1890      Select Case IsNull(.interestShareFace)
          Case True
1900        intPos01 = 0
1910      Case False
1920        intPos01 = InStr(1, .interestShareFace, ".", vbTextCompare)
1930      End Select
1940      If intPos01 > 0 Then
1950        If Len(.interestShareFace) - intPos01 > gintShareFaceDecimals Then
1960          blnContinue = False
1970          MsgBox "You can only enter a share/face with up to " & Trim(str(gintShareFaceDecimals)) & " decimals.", _
                vbInformation + vbOKOnly, "Invalid Entry"
1980          .interestShareFace = vbNullString
1990          DoCmd.CancelEvent
2000        Else
2010          .interestShareFace.Format = "#,###." & String(gintShareFaceDecimals, "0")
2020          .Refresh
2030        End If
2040      Else
2050        .interestShareFace.Format = "#,###"
2060        .Refresh
2070      End If

2080      If blnContinue = True Then

2090        lngCurrID = .interestCurr_ID.Column(CBX_C_CURRID)

2100        If IsNull(.interestAssetNo) = False Then
2110          .interestAssetNo.Requery
2120        End If
2130        gstrInterestAsset = IIf(IsNull(.interestAssetNo), vbNullString, .interestAssetNo)
2140        gstrInterestShareFace = IIf(IsNull(.interestShareFace), vbNullString, .interestShareFace)
2150        gstrInterestAccountNumber = IIf(IsNull(.interestAccountNo), vbNullString, .interestAccountNo)
2160        gstrInterestICash = IIf(IsNull(.interestICash), vbNullString, .interestICash)
2170        gstrInterestICash = Rem_Dollar(gstrInterestICash, lngCurrID)  ' ** Module Function: modStringFuncs.

2180        strAccountNo = .interestAccountNo

2190        Select Case .NewRecord
            Case True
2200          .cmdInterestCancel.Enabled = False
2210        Case False
2220          .cmdInterestCancel.Enabled = True
2230        End Select

2240        If strAccountNo <> vbNullString Then
              ' ** I'm using this to indicate chkShowAllAssets.
2250          Select Case .IsAverage
              Case True
2260            If .interestAssetNo.RowSource <> "qryJournal_Interest_02g" Then
                  ' ** qryJournal_Interest_02c (MasterAsset table, with totdesc)
                  ' ** linked to AssetType, Interest = True, all assets.
2270              .interestAssetNo.RowSource = "qryJournal_Interest_02g"
2280            End If
2290          Case False
                ' ** Queries do not have Dividend = True criteria.
2300            If .interestAccountNo = "INCOME O/U" Or .interestAccountNo = "99-INCOME O/U" Then
2310              If .interestAssetNo.RowSource <> "qryJournal_Interest_02f" Then
                    ' ** qryJournal_Interest_02b (Journal table linked to Account, MasterAsset)
                    ' ** linked to AssetType, by specified FormRef('frmJournal_Sub2_Interest').
2320                .interestAssetNo.RowSource = "qryJournal_Interest_02f"
2330              End If
2340            Else
2350              If .interestAssetNo.RowSource <> "qryJournal_Interest_02d" Then
                    ' ** qryJournal_Interest_02a (MasterAsset table linked to ActiveAssets, Account)
                    ' ** linked to AssetType, by specified FormRef('frmJournal_Sub2_Interest').
2360                .interestAssetNo.RowSource = "qryJournal_Interest_02d"
2370              End If
2380            End If
2390          End Select
2400          .interestAssetNo.Requery
2410        End If

2420        .Refresh

2430      End If  ' ** blnContinue.

2440      If blnCmdButton = False Then  ' ** That is, only run this if Form_Current() wasn't called from another button, like cmdMiscOK_Click().
2450        blnGoToIntReinvest = False
2460        .tglInterestReinvest_false_raised_img_dis.Visible = True
2470        .tglInterestReinvest_false_raised_img.Visible = False
2480        .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = False
2490        .tglInterestReinvest_false_raised_focus_img.Visible = False
2500        .tglInterestReinvest_false_raised_focus_dots_img.Visible = False
2510        .tglInterestReinvest_false_sunken_focus_dots_img.Visible = False
2520        .tglInterestReinvest_true_raised_img.Visible = False
2530        .tglInterestReinvest_true_raised_focus_img.Visible = False
2540        .tglInterestReinvest_true_raised_focus_dots_img.Visible = False
2550        .tglInterestReinvest_true_sunken_focus_dots_img.Visible = False
2560        .tglInterestReinvest_true_raised_img_dis.Visible = False
2570        .tglInterestReinvest.Enabled = False
2580        Select Case .posted
            Case True
2590          .tglInterestReinvest.Enabled = False
2600          .tglInterestReinvest_true_raised_img_dis.Visible = True
2610          .tglInterestReinvest_false_raised_img_dis.Visible = False
2620        Case False
              ' ** Let's wait before enabling this.
2630          If .NewRecord = False Then
2640            .tglInterestReinvest.Enabled = True
2650            .tglInterestReinvest_false_raised_img.Visible = True
2660            .tglInterestReinvest_false_raised_img_dis.Visible = False
2670          End If
2680        End Select
2690      End If  ' ** blnCmdButton.

2700      .Repaint

2710      DoEvents

          ' ** Set the currency symbol.
2720      interestCurr_ID_AfterUpdate  ' ** Procedure: Below.

2730      If .interestCurr_ID.Visible = True Then
2740        interestICash_AfterUpdate  ' ** Procedure: Below.
2750      End If

2760    End With

EXITP:
2770    Exit Sub

ERRH:
2780    THAT_PROC = THIS_PROC
2790    That_Erl = Erl
2800    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2810    Resume EXITP

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

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Mark Reinvest:    R {tglInterestReinvest}
        ' **   Close Timer:      T {Form_KeyDown}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}

        ' ** Plain keys.
2950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2960      Select Case KeyCode
          Case vbKeyEscape
2970        intRetVal = 0
2980        If Forms("frmJournal").opgJournal.Enabled = True Then
2990          Select Case gblnInterestChanged
              Case True
3000            cmdInterestCancel_Click  ' ** Procedure: Below.
3010          Case False
3020            cmdInterestClose_Click  ' ** Procedure: Below.
3030          End Select
3040        Else
3050          cmdInterestCancel_Click  ' ** Procedure: Below.
              'MsgBox "The posting of this transaction is incomplete." & vbCrLf & vbCrLf & _
              '  "You must either complete or cancel this transaction to continue.", vbInformation + vbOKOnly, "Incomplete Data"
              'DoCmd.RunCommand acCmdSaveRecord
3060        End If
3070      Case vbKeyDelete
3080        gblnInterestChanged = True
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
3180          If .cmdInterestClose.Enabled = True Then
3190            .cmdInterestClose.SetFocus
3200            cmdInterestClose_Click  ' ** Procedure: Below.
3210          Else
3220            Beep
3230          End If
3240        End With
3250      Case vbKeyE
3260        With Me
3270          intRetVal = 0
3280          If .cmdInterestCancel.Enabled = True Then
3290            .cmdInterestCancel.SetFocus
3300            cmdInterestCancel_Click  ' ** Procedure: Below.
3310          Else
3320            Beep
3330          End If
3340        End With
3350      Case vbKeyX
3360        With Me
3370          intRetVal = 0
3380          If .cmdInterestClose.Enabled = True Then
3390            cmdInterestClose_Click  ' ** Procedure: Below.
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
3740          .tglInterestReinvest.SetFocus
3750          tglInterestReinvest_Click  ' ** Procedure: Below.
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
3920      If .interestAccountNo.Enabled = True Then
3930        .interestAccountNo.SetFocus
3940      Else
3950        .interestTransDate.SetFocus
3960      End If
3970    End With

EXITP:
3980    Exit Sub

ERRH:
3990    THAT_PROC = THIS_PROC
4000    That_Erl = Erl
4010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
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
4150        If IsNull(.interestID) = False Then
4160          MoveRec acCmdRecordsGoToNew  ' ** Procedure: Below.
4170        End If
4180        DoEvents
4190      End If  ' ** blnMapNewRec.
4200    End With

EXITP:
4210    Exit Sub

ERRH:
4220    THAT_PROC = THIS_PROC
4230    That_Erl = Erl
4240    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
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
            Case "interestShareFace"
4420          strControl = "Shares"
4430        Case "interestPerShare"
4440          strControl = "Per Share"
4450        Case "interestICash"
4460          strControl = "Income Cash"
4470        Case "txtDate", "interesttransdate"
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
4690        If IsNull(Me.interestAccountNo) Or Me.interestAccountNo = vbNullString Then
4700          MsgBox "The posting of this transaction is incomplete." & vbCrLf & vbCrLf & _
                "You must either complete or cancel this transaction to continue.", vbInformation + vbOKOnly, "Incomplete Data"
4710        End If
4720        If Me.interestAccountNo.Enabled = True Then
4730          Me.interestAccountNo.SetFocus
4740        Else
4750          Me.interestTransDate.SetFocus
4760        End If
4770        Response = acDataErrContinue
4780      Case 3162  ' ** You tried to assign the Null value to a variable that is not a Variant data type.
4790  On Error Resume Next
4800        strControl = Screen.ActiveControl.Name
4810  On Error GoTo ERRH
4820        MsgBox "The " & strControl & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
            ' ** Do nothing.
4830      Case Else
4840        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
4850        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
4860      End Select
4870    End If

EXITP:
4880    Exit Sub

ERRH:
4890    Select Case ERR.Number
        Case Else
4900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
4910    End Select
4920    Resume EXITP

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
5350      If .tglInterestReinvest_true_raised_focus_img.Visible = True Or .tglInterestReinvest_true_raised_focus_dots_img.Visible = True Or _
              .tglInterestReinvest_false_raised_focus_img.Visible = True Or .tglInterestReinvest_false_raised_focus_dots_img.Visible = True Then
5360        Select Case .posted
            Case True
5370          Select Case blnInterestReinvest_Focus
              Case True
5380            .tglInterestReinvest_true_raised_focus_dots_img.Visible = True  ' ** Same for ON.
5390            .tglInterestReinvest_true_raised_img.Visible = False
5400          Case False
5410            .tglInterestReinvest_true_raised_img.Visible = True
5420            .tglInterestReinvest_true_raised_focus_dots_img.Visible = False
5430          End Select
5440          .tglInterestReinvest_false_raised_img.Visible = False
5450          .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = False
5460        Case False
5470          Select Case blnInterestReinvest_Focus
              Case True
5480            .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = True
5490            .tglInterestReinvest_false_raised_img.Visible = False
5500          Case False
5510            .tglInterestReinvest_false_raised_img.Visible = True
5520            .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = False
5530          End Select
5540          .tglInterestReinvest_true_raised_img.Visible = False
5550          .tglInterestReinvest_true_raised_focus_dots_img.Visible = False
5560        End Select
5570        .tglInterestReinvest_false_raised_focus_img.Visible = False
5580        .tglInterestReinvest_false_raised_focus_dots_img.Visible = False
5590        .tglInterestReinvest_false_sunken_focus_dots_img.Visible = False
5600        .tglInterestReinvest_false_raised_img_dis.Visible = False
5610        .tglInterestReinvest_true_raised_focus_img.Visible = False
5620        .tglInterestReinvest_true_sunken_focus_dots_img.Visible = False
5630        .tglInterestReinvest_true_raised_img_dis.Visible = False
5640      End If
5650    End With

EXITP:
5660    Exit Sub

ERRH:
5670    Select Case ERR.Number
        Case Else
5680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
5690    End Select
5700    Resume EXITP

End Sub

Public Sub cmdInterestClose_Click()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdInterestClose_Click"

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

Private Sub cmdInterestOK_Click()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdInterestOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngJrnlID As Long, lngJrnlMS_ID As Long, lngThisCurrID As Long, datTransDate As Date
        Dim strShortName As String, dblICash As Double
        Dim strJrnlUser As String, strMsg As String, strDocName As String
        Dim lngRecs As Long
        Dim blnTmpGoToIntReinvest As Boolean
        Dim lngX As Long
        Dim blnContinue As Boolean

6010    blnContinue = True

6020    DoCmd.Hourglass True
6030    DoEvents

6040    blnCmdButton = True
6050    blnTmpGoToIntReinvest = blnGoToIntReinvest
6060    Form_Current  ' ** Procedure: Above.
6070    blnCmdButton = False
6080    VerifyAccountNo  ' ** Function: Below.

6090    With Me

6100      If gstrInterestAccountNumber = vbNullString Or gstrInterestAccountNumber = "0" Then
6110        blnContinue = False
6120        DoCmd.Hourglass False
6130        MsgBox "Please enter a valid account number.", vbInformation + vbOKOnly, "Invalid Entry"
6140        interestAccountNo.SetFocus
6150      Else
6160        If gstrInterestAsset = vbNullString Or gstrInterestAsset = "0" Then
6170          blnContinue = False
6180          DoCmd.Hourglass False
6190          MsgBox "An asset must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
6200          interestAssetNo.SetFocus
6210        Else
6220          If gstrInterestShareFace = vbNullString Or Val(gstrInterestShareFace) = 0 Then
6230            blnContinue = False
6240            DoCmd.Hourglass False
6250            MsgBox "Face must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
6260            interestShareFace.SetFocus
6270          Else
6280            If gstrInterestICash = vbNullString Or Val(gstrInterestICash) = 0 Then
6290              blnContinue = False
6300              DoCmd.Hourglass False
6310              MsgBox "Income Cash must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
6320              interestICash.SetFocus
6330            Else
6340              If CDbl(gstrInterestICash) <= 0 Then
6350                blnContinue = False
6360                DoCmd.Hourglass False
6370                MsgBox "Income Cash must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
6380                .interestICash.SetFocus
6390              Else
6400                .interestDescription.SetFocus
6410                If InStr(.interestDescription.text, Chr(34)) > 0 Then
6420                  blnContinue = False
6430                  DoCmd.Hourglass False
6440                  MsgBox "Cannot have quote marks in text.", vbInformation + vbOKOnly, "Invalid Characters"
6450                Else
6460                  If IsNull(.interestTransDate) = True Then
                        ' ** This extra check is the result of Jim Townsand's (First Clarksdale) problem.
6470                    blnContinue = False
6480                    DoCmd.Hourglass False
6490                    MsgBox "A Posting Date must be entered.", vbInformation + vbOKOnly, "Entry Required"
6500                    .interestTransDate.SetFocus
6510                  Else
6520                    If .interestTransDate > DateAdd("m", 1, Date) Then
6530                      blnContinue = False
6540                      DoCmd.Hourglass False
6550                      MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
6560                      .interestTransDate.SetFocus
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

6680        If Trim(Nz(.interestUser, vbNullString)) = vbNullString Then
6690          .interestUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
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
6850          .cmbRevenueCodes = REVID_INTINC  ' ** Interest Income.
6860        End If

6870        If blnTmpGoToIntReinvest = True Then
6880          .interestReinvested = True
6890        End If

6900        DoCmd.RunCommand acCmdSaveRecord

6910        DoCmd.Hourglass True  ' ** Make sure it's still running.
6920        DoEvents

6930        lngJrnlID = .interestID
6940        gstrAccountNo = .interestAccountNo.Value
6950        glngAssetNo = .interestAssetNo.Column(CBX_A_ASTNO)
6960        strShortName = .interestShortName.Value
6970        datTransDate = .interestTransDate.Value
6980        dblICash = Nz(.interestICash, 0)
6990        lngThisCurrID = .interestCurr_ID.Column(CBX_C_CURRID)

7000        Set dbs = CurrentDb
            ' ** qryJournal_Interest_01 (Journal, linked to account, for 'Interest'),
            ' ** check for existing Interests, by specified [actno], [astno], [jrnlid].
7010        Set qdf = dbs.QueryDefs("qryJournal_Interest_04")
7020        With qdf.Parameters
7030          ![jrnlid] = lngJrnlID
7040          ![actno] = gstrAccountNo
7050          ![astno] = glngAssetNo
7060        End With
7070        Set rst = qdf.OpenRecordset()
7080        If rst.BOF = True And rst.EOF = True Then
              ' ** No records; all's well.
7090          If .interestCurr_Date.Visible = True Then .interestCurr_Date.Visible = False
7100          rst.Close
7110          .Requery
7120          DoCmd.GoToRecord , , acNewRec
7130        Else
7140          strMsg = vbNullString
7150          rst.MoveLast
7160          lngRecs = rst.RecordCount
7170          rst.MoveFirst
7180          If lngRecs > 1& Then
7190            For lngX = 1& To lngRecs
7200              If IsNull(rst![journal_USER]) = True Then
7210                strJrnlUser = "{unknown}"
7220              Else
7230                If Trim(rst![journal_USER]) = vbNullString Then
7240                  strJrnlUser = "{unknown}"
7250                Else
7260                  strJrnlUser = rst![journal_USER]
7270                End If
7280              End If
7290              strMsg = strMsg & strJrnlUser & ", "
7300              If lngX < lngRecs Then rst.MoveNext
7310            Next
7320            strMsg = Trim(strMsg)                    ' ** Remove trailing space.
7330            strMsg = Left(strMsg, (Len(strMsg) - 1))  ' ** Remove trailing comma.
7340          Else
7350            If IsNull(rst![journal_USER]) = True Then
7360              strJrnlUser = "{unknown}"
7370            Else
7380              If Trim(rst![journal_USER]) = vbNullString Then
7390                strJrnlUser = "{unknown}"
7400              Else
7410                strJrnlUser = rst![journal_USER]
7420              End If
7430            End If
7440            strMsg = strJrnlUser
7450          End If
7460          strMsg = IIf(strMsg = CurrentUser, "You have", "The user" & IIf(lngRecs = 1, " " & strMsg & " has", "s " & strMsg & " have")) & _
                " already entered interest for this asset!" & vbCrLf & vbCrLf & "Do you still wish to post this interest?"
              ' ** Internal Access Function: Trust Accountant login.
7470          If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Warning: Possible Duplicate Data") = vbYes Then
7480            .Requery
7490            DoCmd.GoToRecord , , acNewRec
7500          Else
7510            blnGoToIntReinvest = False
7520            blnTmpGoToIntReinvest = False
7530            DoCmd.RunCommand (acCmdDeleteRecord)
7540          End If
7550          If .interestCurr_Date.Visible = True Then .interestCurr_Date.Visible = False
7560          rst.Close
7570        End If
7580        Set rst = Nothing
7590        Set qdf = Nothing
7600        dbs.Close
7610        Set dbs = Nothing

7620        DoCmd.Hourglass True  ' ** Make sure it's still running.
7630        DoEvents

7640        If .interestAccountNo.Enabled = True Then
7650          .interestAccountNo.SetFocus
7660        Else
7670          .interestTransDate.SetFocus
7680        End If

7690        gblnInterestValidated = False

7700        InterestChanged False  ' ** Procedure: Below.

7710        DoCmd.Hourglass True  ' ** Make sure it's still running.
7720        DoEvents

7730        blnGoToIntReinvest = blnTmpGoToIntReinvest

            ' ** Reinvest button was pushed.
7740        If blnGoToIntReinvest = True Then

              ' ** Put preliminary record into tblJournal_MiscSold.
7750          Set dbs = CurrentDb
7760          Set rst = dbs.OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
7770          With rst
7780            If .BOF = True And .EOF = True Then
                  ' ** Proceed.
7790            Else
7800              .MoveFirst
7810              .FindFirst "[Journal_ID] = " & CStr(lngJrnlID)
7820              If .NoMatch = False Then
7830                .Delete
7840              End If
7850            End If
7860            DoEvents
7870            .AddNew
                ' ** ![jrnlms_id] : AutoNumber.
7880            ![Journal_ID] = lngJrnlID
7890            ![journaltype] = "Interest"
7900            ![jrnlms_ref_id] = 0&
7910            ![jrnlms_parent] = True
7920            ![jrnlms_child] = False
7930            ![jrnlms_note] = "This Interest was reinvested, generating a Purchase transaction."
7940            ![jrnlms_datemodified] = Now()
7950            .Update
7960            .Bookmark = .LastModified
7970            lngJrnlMS_ID = ![jrnlms_id]
7980          End With
7990          Set rst = Nothing
8000          dbs.Close
8010          Set dbs = Nothing
8020          DoEvents

8030          gblnMessage = True  ' ** If this returns False, the dialog was canceled.
8040          gdblCrtRpt_CostTot = 0#  ' ** Borrowing this variable from the Court Reports.
8050          strDocName = "frmReinvest_Interest"
8060          DoCmd.OpenForm strDocName, , , , acFormAdd, acDialog, strCallingForm & "~" & gstrAccountNo & "~" & _
                Format(datTransDate, "mm/dd/yyyy") & "~" & CStr(lngThisCurrID) & "~" & _
                CStr(dblICash) & "~" & CStr(glngAssetNo) & "~" & CStr(lngJrnlID) & "~" & strShortName

8070          DoCmd.Hourglass True
8080          DoEvents

8090          Set dbs = CurrentDb
8100          Select Case gblnMessage
              Case True
8110            Set rst = dbs.OpenRecordset("tblJournal_MiscSold", dbOpenDynaset, dbConsistent)
8120            With rst
8130              .MoveFirst
8140              .FindFirst "[jrnlms_id] = " & CStr(lngJrnlMS_ID)
8150              If .NoMatch = False Then
8160                .Edit
8170                ![jrnlms_ref_id] = CLng(gdblCrtRpt_CostTot)  ' ** New Purchase ID from frmReinvest_Dividend.
8180                ![jrnlms_datemodified] = Now()
8190                .Update
8200              End If
8210              .AddNew
                  ' ** ![jrnlms_id] : AutoNumber.
8220              ![Journal_ID] = CLng(gdblCrtRpt_CostTot)
8230              ![journaltype] = "Purchase"
8240              ![jrnlms_ref_id] = lngJrnlID
8250              ![jrnlms_parent] = False
8260              ![jrnlms_child] = True
8270              ![jrnlms_note] = "This Purchase was generated by an Interest reinvestment."
8280              ![jrnlms_datemodified] = Now()
8290              .Update
8300              .Close
8310            End With
8320            Set rst = Nothing
8330            If .cmdInterestOK.Enabled = True Then
8340              .cmdInterestOK.SetFocus
8350            Else
8360              .interestAccountNo.SetFocus
8370            End If
8380          Case False
                ' ** Delete tblJournal_MiscSold, by specified [jrnlmsid].
8390            Set qdf = dbs.QueryDefs("qryJournal_MiscSold_01")
8400            With qdf.Parameters
8410              ![jrnlmsid] = lngJrnlMS_ID
8420            End With
8430            qdf.Execute
8440            Set qdf = Nothing
8450            DoEvents
                ' ** Update Journal, by specified [jrnlid], [pstd].
8460            Set qdf = dbs.QueryDefs("qryJournal_Interest_14")
8470            With qdf.Parameters
8480              ![jrnlid] = lngJrnlID  ' ** This new Interest record.
8490              ![pstd] = False
8500            End With
8510            qdf.Execute
8520            Set qdf = Nothing
8530            If gdblCrtRpt_CostTot <> 0# Then
                  ' ** Delete Journal, by specified [jrnlid].
8540              Set qdf = dbs.QueryDefs("qryJournal_Interest_13")
8550              With qdf.Parameters
8560                ![jrnlid] = CLng(gdblCrtRpt_CostTot)  ' ** New Purchase from frmReinvest_Interest.
8570              End With
8580              qdf.Execute dbFailOnError
8590              Set qdf = Nothing
8600              gdblCrtRpt_CostTot = 0#
8610            End If
8620            .interestAccountNo.SetFocus
8630          End Select

8640          dbs.Close
8650          Set dbs = Nothing

8660        End If  ' ** blnGoToIntReinvest
8670        blnGoToIntReinvest = False  ' ** reset always.

8680      End If  ' ** blnContinue

8690      DoCmd.Hourglass False

8700    End With

EXITP:
8710    Set rst = Nothing
8720    Set qdf = Nothing
8730    Set dbs = Nothing
8740    Exit Sub

ERRH:
8750    DoCmd.Hourglass False
8760    THAT_PROC = THIS_PROC
8770    That_Erl = Erl
8780    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8790    Resume EXITP

End Sub

Private Sub cmdInterestCancel_Click()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdInterestCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngThisJrnlID As Long, lngThatJrnlID As Long
        Dim blnReinvested As Boolean, strMsg As String
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

8810    With Me

8820      DoCmd.Hourglass True
8830      DoEvents

8840      If IsNull(.interestID) = True Then
8850        msgResponse = vbCancel
8860      Else

8870        blnReinvested = .interestReinvested
8880        lngThisJrnlID = .interestID
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
                  "Yes - Delete both this Interest and the Purchase reinvestment." & vbCrLf & _
                  "No - Delete only this Interest." & vbCrLf & _
                  "Cancel - Delete no transactions."
8960            Beep
8970            DoCmd.Hourglass False
8980            msgResponse = MsgBox(strMsg, vbQuestion + vbYesNoCancel + vbDefaultButton3, "Delete Both Transactions")
8990            If msgResponse <> vbCancel Then
9000              DoCmd.Hourglass True
9010              If .interestCurr_Date.Visible = True Then .interestCurr_Date.Visible = False
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
9150          msgResponse = vbNo  ' ** Delete only Interest.
9160          DoCmd.Hourglass True
9170          If .interestCurr_Date.Visible = True Then .interestCurr_Date.Visible = False
9180          DoEvents
9190        End If
9200      End If

9210      If msgResponse = vbCancel Then
9220        If .interestAccountNo.Enabled = True Then
9230          .interestAccountNo.SetFocus
9240        Else
9250          .interestTransDate.SetFocus
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
              ' ** Delete tblJournal_MiscSold, by specified [jrnlidref].
9390          Set qdf = dbs.QueryDefs("qryJournal_MiscSold_03")
9400          With qdf.Parameters
9410            ![jrnlidref] = lngThatJrnlID
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

9550        blnGoToIntReinvest = False
9560        .tglInterestReinvest_false_raised_img_dis.Visible = True
9570        .tglInterestReinvest_false_raised_img.Visible = False
9580        .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = False
9590        .tglInterestReinvest_false_raised_focus_img.Visible = False
9600        .tglInterestReinvest_false_raised_focus_dots_img.Visible = False
9610        .tglInterestReinvest_false_sunken_focus_dots_img.Visible = False
9620        .tglInterestReinvest_true_raised_img.Visible = False
9630        .tglInterestReinvest_true_raised_focus_img.Visible = False
9640        .tglInterestReinvest_true_raised_focus_dots_img.Visible = False
9650        .tglInterestReinvest_true_sunken_focus_dots_img.Visible = False
9660        .tglInterestReinvest_true_raised_img_dis.Visible = False
9670        .tglInterestReinvest.Enabled = False

9680        .cmbAccountHelper = Null

            ' ** Also delete reinvestment.
9690        If msgResponse = vbYes And lngThatJrnlID > 0& Then
9700          Set dbs = CurrentDb
              ' ** Delete Journal, by specified [jrnlid].
9710          Set qdf = dbs.QueryDefs("qryJournal_Interest_13")
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

9820        If .interestAccountNo.Enabled = True Then
9830          .interestAccountNo.SetFocus
9840        Else
9850          .interestTransDate.SetFocus
9860        End If

9870        gblnInterestValidated = False

9880        InterestChanged False  ' ** Procedure: Below.

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

Private Sub interestAccountNo_AfterUpdate()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAccountNo_AfterUpdate"

10110   InterestChanged True  ' ** Procedure: Below.

EXITP:
10120   Exit Sub

ERRH:
10130   Select Case ERR.Number
        Case Else
10140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10150   End Select
10160   Resume EXITP

End Sub

Private Sub interestAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAccountNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10220   intShiftDown = (Shift And acShiftMask) > 0
10230   intAltDown = (Shift And acAltMask) > 0
10240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10260     Select Case KeyCode
          Case vbKeyTab
            ' ** Tab key copies AccountNo. {intRetVal not cleared.}
10270       With Me
10280         If gblnTabCopyAccount = True Then
10290           If .NewRecord = True Then
10300             If IsNull(.interestAccountNo) = True Then
10310               If Len(.interestAccountNo.text) = 0 Then
10320                 If IsNull(.Parent.LastAcctNo) = False Then
10330                   If .Parent.LastAcctNo <> vbNullString Then
10340                     .interestAccountNo = .Parent.LastAcctNo
10350                     InterestChanged True  ' ** Procedure: Below.
10360                   End If
10370                 End If
10380               End If
10390             Else
10400               If Len(.interestAccountNo.text) = 0 Then
10410                 If IsNull(.Parent.LastAcctNo) = False Then
10420                   If .Parent.LastAcctNo <> vbNullString Then
10430                     .interestAccountNo = .Parent.LastAcctNo
10440                     InterestChanged True  ' ** Procedure: Below.
10450                   End If
10460                 End If
10470               End If
10480             End If
10490           End If
10500         End If
10510         gblnInterestChanged = True
10520         gblnInterestValidated = False
10530       End With
10540     Case vbKeyReturn
10550       With Me
10560         intRetVal = 0
10570         gblnInterestChanged = True  ' ** Just added: 09/19/2012.
10580         gblnInterestValidated = False
10590         .interestTransDate.SetFocus
10600       End With
10610     Case vbKeyDelete
10620       gblnInterestValidated = False
10630     End Select
10640   End If

EXITP:
10650   KeyCode = intRetVal
10660   Exit Sub

ERRH:
10670   intRetVal = 0
10680   THAT_PROC = THIS_PROC
10690   That_Erl = Erl
10700   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10710   Resume EXITP

End Sub

Private Sub interestAccountNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAccountNo_MouseMove"

10810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10820   Exit Sub

ERRH:
10830   Select Case ERR.Number
        Case Else
10840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
10850   End Select
10860   Resume EXITP

End Sub

Private Sub interestAccountNo_Exit(Cancel As Integer)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAccountNo_Exit"

        Dim strAccountNo As String
        Dim blnGoToAccountHelper As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

10910   With Me

10920     blnGoToAccountHelper = False  ' ** Unless set otherwise.

10930     If ((gblnInterestValidated = False And gblnInterestChanged = True) And (lngErrCnt < 3&)) Then
10940       If IsNull(.interestAccountNo) = True Then
10950         lngErrCnt = lngErrCnt + 1&
10960         msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Entry Required")
10970         If msgResponse = vbYes Then
10980           blnGoToAccountHelper = True
10990         Else
11000           gblnInterestValidated = False
11010           If blnStartTrans = True Then
11020             Cancel = -1
11030             .interestAccountNo.SetFocus
11040           End If
11050         End If
11060       Else
11070         lngErrCnt = lngErrCnt + 1&
11080         If .interestAccountNo = vbNullString And .interestAccountNo.text = vbNullString Then
11090           msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                  "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Entry Required")
11100           If msgResponse = vbYes Then
11110             blnGoToAccountHelper = True
11120           Else
11130             gblnInterestValidated = False
11140             If blnStartTrans = True Then
11150               Cancel = -1
11160               .interestAccountNo.SetFocus
11170             End If
11180           End If
11190         Else
                ' ** Check to see if it is a valid account.
11200           lngErrCnt = 0&
11210           Select Case gblnAccountNoWithType
                Case True
11220             strAccountNo = Left(.interestAccountNo, 2) & Mid(.interestAccountNo, 3)
11230           Case False
11240             strAccountNo = .interestAccountNo
11250           End Select
11260           varTmp00 = DLookup("[accountno]", "account", "[accountno] = '" & strAccountNo & "' And ([dateclosed] Is Null)")
11270           Select Case IsNull(varTmp00)
                Case True
11280             blnStartTrans = True  ' ** Once they've started, they've got to enter a valid number!
11290             msgResponse = MsgBox("Please enter a valid account number to continue." & vbCrLf & vbCrLf & _
                    "Would you like to pick an account from a list?", vbQuestion + vbYesNo, "Invalid Entry")
11300             If msgResponse = vbYes Then
11310               blnGoToAccountHelper = True
11320             Else
11330               gblnInterestValidated = False
11340               Cancel = -1
11350               .interestAccountNo.SetFocus
11360             End If
11370             .interestAccountNo = vbNullString
11380           Case False
11390             gblnInterestValidated = True
11400           End Select
11410           If IsNull(.interestAccountNo) = False Then
11420             If .interestAccountNo <> vbNullString Then
11430               .interestAccountNo_Data = .interestAccountNo
11440               gstrInterestAccountNumber = .interestAccountNo.text
11450               .Parent.LastAcctNo = gstrInterestAccountNumber
11460               .cmdInterestMap.Enabled = False
11470 On Error Resume Next
11480               .Refresh
11490               If ERR.Number <> 0 Then  ' ** 3314: The field 'journal.journaltype' cannot contain a Null value because the Required property for this field is set to True.
11500                 .interestType = "Interest"
11510                 blnAccountNoErr = True
11520               End If
11530 On Error GoTo ERRH
11540               .tglInterestReinvest.Enabled = True
11550               .tglInterestReinvest_false_raised_img.Visible = True
11560               .tglInterestReinvest_false_raised_img_dis.Visible = False
11570             End If
11580           End If
11590         End If
11600       End If
11610     Else
11620       lngErrCnt = 0&
11630     End If

11640     If blnGoToAccountHelper = True Then
11650       .cmbAccountHelper.SetFocus
11660     End If

11670   End With  ' ** Me.

EXITP:
11680   Exit Sub

ERRH:
11690   THAT_PROC = THIS_PROC
11700   That_Erl = Erl
11710   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11720   Resume EXITP

End Sub

Private Sub cmdLock_Click()

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLock_Click"

11810   With Me
11820     If .cmdLock_open_raised_img.Visible = True Then
11830       .interestTransDate.SetFocus
11840       .interestAccountNo.Enabled = False
11850       .interestAccountNo.BorderColor = WIN_CLR_DISR
11860       .interestAccountNo_lbl.BackStyle = acBackStyleTransparent
11870       .interestAccountNo_lbl_box.Visible = True
11880       .cmdLock_open_raised_img.Visible = False
11890       .cmdLock_closed_raised_img.Visible = True
11900     ElseIf .cmdLock_closed_raised_img.Visible = True Then
11910       .interestAccountNo.Enabled = True
11920       .interestAccountNo.BorderColor = CLR_LTBLU2
11930       .interestAccountNo_lbl.BackStyle = acBackStyleNormal
11940       .interestAccountNo_lbl_box.Visible = False
11950       .interestAccountNo.SetFocus
11960       .cmdLock_open_raised_img.Visible = True
11970       .cmdLock_closed_raised_img.Visible = False
11980     End If
11990   End With

EXITP:
12000   Exit Sub

ERRH:
12010   Select Case ERR.Number
        Case Else
12020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12030   End Select
12040   Resume EXITP

End Sub

Private Sub cmbAccountHelper_Enter()

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_Enter"

12110   With Me
12120     Select Case IsNull(.cmbAccountHelper)
          Case True
12130       ExpandCombo .cmbAccountHelper  ' ** Module Procedure: modUtilities.
12140     Case False
12150       If Trim(.cmbAccountHelper) = vbNullString Then
12160         ExpandCombo .cmbAccountHelper  ' ** Module Procedure: modUtilities.
12170       End If
12180     End Select
12190   End With

EXITP:
12200   Exit Sub

ERRH:
12210   Select Case ERR.Number
        Case Else
12220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
12230   End Select
12240   Resume EXITP

End Sub

Private Sub cmbAccountHelper_GotFocus()

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_GotFocus"

12310   With Me
12320     .cmbAccountHelper.ForeColor = 0
12330   End With

EXITP:
12340   Exit Sub

ERRH:
12350   Select Case ERR.Number
        Case Else
12360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12370   End Select
12380   Resume EXITP

End Sub

Private Sub cmbAccountHelper_AfterUpdate()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_AfterUpdate"

12410   With Me
12420     If IsNull(.cmbAccountHelper) = False Then
12430       .interestAccountNo = .cmbAccountHelper
12440       .interestAccountNo.SetFocus
12450     End If
12460   End With

EXITP:
12470   Exit Sub

ERRH:
12480   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
12490   Case Else
12500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12510   End Select
12520   Resume EXITP

End Sub

Private Sub cmbAccountHelper_LostFocus()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccountHelper_LostFocus"

12610   With Me
12620     gblnInterestValidated = False
12630     InterestChanged True  ' ** Procedure: Below.
12640     .cmbAccountHelper.ForeColor = CLR_WHT
12650     .interestAccountNo.SetFocus
12660     DoEvents
12670     .interestAccountNo.Value = .cmbAccountHelper.Value
12680     DoEvents
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

Private Sub interestShortName_LostFocus()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "interestShortName_LostFocus"

12810   VerifyAccountNo  ' ** Function: Below.

EXITP:
12820   Exit Sub

ERRH:
12830   Select Case ERR.Number
        Case Else
12840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12850   End Select
12860   Resume EXITP

End Sub

Private Sub interestType_Exit(Cancel As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "interestType_Exit"

12910   With Me
12920     gstrInterestType = .interestType.text
12930   End With

EXITP:
12940   Exit Sub

ERRH:
12950   Select Case ERR.Number
        Case Else
12960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12970   End Select
12980   Resume EXITP

End Sub

Private Sub interestType_LostFocus()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "interestType_LostFocus"

13010   VerifyAccountNo  ' ** Function: Below.

EXITP:
13020   Exit Sub

ERRH:
13030   Select Case ERR.Number
        Case Else
13040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13050   End Select
13060   Resume EXITP

End Sub

Public Sub cmdInterestMap_Click()
' ** Public for Go-To-Report.

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdInterestMap_Click"

        Dim blnContinue As Boolean
        Dim strDocName As String
        Dim varTmp00 As Variant

13110   blnContinue = True

13120   With Me

13130     If .interestAccountNo.Enabled = True Then
13140       .interestAccountNo.SetFocus
13150     Else
13160       .interestTransDate.SetFocus
13170     End If
13180     If IsNull(.interestAccountNo) Then
            ' ** Open the map.
13190     Else
13200       If Len(Trim(.interestAccountNo.text)) = 0 Then
              ' ** Open the map.
13210       Else
13220         blnContinue = False
13230         .cmdInterestMap.SetFocus
13240         MsgBox "Please advance to a blank posting form.", vbInformation + vbOKOnly, "New Record Required"
13250         DoCmd.CancelEvent
13260       End If
13270     End If

13280     If blnContinue = True Then

13290       .cmdInterestMap.SetFocus

13300       .Parent.LoadMapDropdowns  ' ** Form Procedure: frmJournal.
13310       DoEvents

13320       strDocName = "frmMap_Int"
13330       DoCmd.OpenForm strDocName, , , , , , "frmJournal"

13340       If gblnGoToReport = True Then
13350         varTmp00 = DMax("[ID]", "journal")  ' ** Save this so we can delete any fake Journal records.
13360         Select Case IsNull(varTmp00)
              Case True
13370           glngTaxCode_Distribution = 0&  ' ** Borrowing this variable from the Court Reports.
13380         Case False
13390           glngTaxCode_Distribution = varTmp00
13400         End Select
13410         Forms(strDocName).TimerInterval = 100&
13420         .GoToReport_arw_mapint_img.Visible = False
13430         DoCmd.Hourglass True  ' ** Make sure it's still running.
13440         DoEvents
13450       End If

13460     End If

13470   End With

EXITP:
13480   Exit Sub

ERRH:
13490   THAT_PROC = THIS_PROC
13500   That_Erl = Erl
13510   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13520   Resume EXITP

End Sub

Private Sub cmdInterestMap_KeyDown(KeyCode As Integer, Shift As Integer)
' ** See Note, above.

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdInterestMap_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13620   intShiftDown = (Shift And acShiftMask) > 0
13630   intAltDown = (Shift And acAltMask) > 0
13640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13660     Select Case KeyCode
          Case vbKeyTab
13670       With Me
13680         intRetVal = 0
13690         .interestTransDate.SetFocus
13700       End With
13710     End Select
13720   End If

        ' ** Shift keys.
13730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13740     Select Case KeyCode
          Case vbKeyTab
13750       With Me
13760         intRetVal = 0
13770         .interestType.SetFocus
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

Private Sub interestTransDate_BeforeUpdate(Cancel As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "interestTransDate_BeforeUpdate"

13910   With Me
13920     If .interestTransDate.text <> "__/__/____" And .interestTransDate.text <> vbNullString Then
13930       If Left(.interestTransDate.text, 2) > 12 Then
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

Private Sub interestTransDate_KeyDown(KeyCode As Integer, Shift As Integer)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "interestTransDate_KeyDown"

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
14190         .interestAssetNo.SetFocus
14200       End With
14210     End Select
14220   End If

        ' ** Shift keys.
14230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14250       With Me
14260         intRetVal = 0
14270         .interestType.SetFocus
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

Private Sub interestTransDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "interestTransDate_MouseMove"

14510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14520   Exit Sub

ERRH:
14530   Select Case ERR.Number
        Case Else
14540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
14550   End Select
14560   Resume EXITP

End Sub

Private Sub interestTransDate_Exit(Cancel As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "interestTransDate_Exit"

14610   With Me
14620     If lngDateCnt <= 1& Then
14630       If IsNull(.interestAccountNo) = False Then
14640         If Trim(.interestAccountNo) <> vbNullString Then
14650           If DateCheck_Post(.interestTransDate.text) = True Then  ' ** Module Function: modUtilities.
                  ' ** '.Text' needed to reliably check date.
14660             lngDateCnt = 0&
14670             If IsNull(.interestTransDate.OldValue) Then
14680               InterestChanged True  ' ** Procedure: Below.
14690             Else
14700               If .interestTransDate <> .interestTransDate.OldValue Then
14710                 InterestChanged True  ' ** Procedure: Below.
14720               End If
14730             End If
14740           Else
14750             If IsNull(.interestTransDate) = True Then
14760               .interestTransDate = Date
14770             Else
14780               If Trim(.interestTransDate) = vbNullString Then
14790                 .interestTransDate = Date
14800               Else
14810                 lngDateCnt = lngDateCnt + 1&
14820                 Cancel = -1
14830                 .interestTransDate.SetFocus
14840               End If
14850             End If
14860           End If
14870         End If
14880       End If
14890     Else
14900       lngDateCnt = 0&
14910       .interestTransDate = CDate(Format(datPostingDate, "mm/dd/yyyy"))
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

Private Sub interestTransDate_LostFocus()

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "interestTransDate_LostFocus"

15010   With Me
15020     If IsNull(.interestAccountNo) = False Then
15030       If Trim(.interestAccountNo) <> vbNullString Then
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
15320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
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
15520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
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
15680         .interestTransDate = CDate(Format(Date, "mm/dd/yyyy"))
15690       Else
15700         .interestTransDate = datStartDate
15710       End If
15720     Else
15730       .interestTransDate = CDate(Format(Date, "mm/dd/yyyy"))
15740     End If
15750     .interestAssetNo.SetFocus
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
16080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
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
16320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
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
16520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
16530   End Select
16540   Resume EXITP

End Sub

Private Sub chkShowAllAssets_AfterUpdate()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAllAssets_AfterUpdate"

16610   With Me
16620     Select Case .chkShowAllAssets
          Case True
16630       If .interestAssetNo.RowSource <> "qryJournal_Interest_02g" Then
              ' ** qryJournal_Interest_02c (MasterAsset table, with totdesc)
              ' ** linked to AssetType, Interest = True, all assets.
16640         .interestAssetNo.RowSource = "qryJournal_Interest_02g"
16650       Else
16660         .interestAssetNo.Requery
16670       End If
16680     Case False
16690       If .interestAssetNo.RowSource <> "qryJournal_Interest_02e" Then
              ' ** qryJournal_Interest_02a (MasterAsset table linked to ActiveAssets, Account) linked
              ' ** to AssetType; Interest = True, by specified FormRef('frmJournal_Sub2_Interest').
16700         .interestAssetNo.RowSource = "qryJournal_Interest_02e"
16710       Else
16720         .interestAssetNo.Requery
16730       End If
16740     End Select
16750   End With

        ' ** These queries are used on the Form_Current() event.
        ' **   qryJournal_
        ' **   qryJournal_

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
16990         .interestAssetNo.SetFocus
17000       End With
17010     End Select
17020   End If

        ' ** Shift keys.
17030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17050       With Me
17060         intRetVal = 0
17070         .interestTransDate.SetFocus
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

Private Sub interestAssetNo_Enter()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAssetNo_Enter"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnContinue As Boolean
        Dim blnHasInt As Boolean
        Dim strSQL As String
        Dim varTmp00 As Variant

17210   blnContinue = True

17220   With Me

17230     DoCmd.Hourglass True
17240     DoEvents

17250     If Not blnAccountValidated Then
17260       blnCheckingAccount = True
17270       blnExitingAssetNo = True
17280       VerifyAccountNo  ' ** Function: Below.
17290       blnCheckingAccount = False
17300       blnExitingAssetNo = False
17310       If Not blnAccountValidated Then
17320         blnContinue = False
17330         blnExitingAssetNo = True
17340       Else
17350         .interestAssetNo.SetFocus
17360       End If
17370     End If

17380     If blnContinue = True Then

17390       If .interestAccountNo = "INCOME O/U" Or .interestAccountNo = "99-INCOME O/U" Then
              ' ** qryJournal_Interest_02e.
17400         strSQL = "SELECT DISTINCTROW qryJournal_Interest_02b.totdesc, qryJournal_Interest_02b.cusip, " & _
                "qryJournal_Interest_02b.assetno, qryJournal_Interest_02b.rate, qryJournal_Interest_02b.due, assettype.taxcode, " & _
                "qryJournal_Interest_02b.curr_id " & _
                "FROM qryJournal_Interest_02b INNER JOIN assettype ON qryJournal_Interest_02b.assettype = assettype.assettype " & _
                "WHERE (((qryJournal_Interest_02b.accountno)='" & .interestAccountNo & "')) " & _
                "ORDER BY qryJournal_Interest_02b.assettype, qryJournal_Interest_02b.totdesc;"  '#curr_id
17410       Else

              ' ** Check to see if this account has any Interest assets!
17420         blnHasInt = True
17430         Set dbs = CurrentDb
17440         Set qdf = dbs.QueryDefs("qryJournal_Interest_06")
17450         With qdf.Parameters
17460           ![actno] = Me.interestAccountNo
17470         End With
17480         Set rst = qdf.OpenRecordset
17490         With rst
17500           If .BOF = True And .EOF = True Then
17510             blnHasInt = False
17520           End If
17530           .Close
17540         End With
17550         dbs.Close

17560         Select Case .chkShowAllAssets
              Case True
17570           If .interestAssetNo.RowSource <> "qryJournal_Interest_02g" Then
                  ' ** qryJournal_Interest_02c (MasterAsset table, with totdesc)
                  ' ** linked to AssetType, Interest = True, all assets.
17580             .interestAssetNo.RowSource = "qryJournal_Interest_02g"
17590           Else
17600             .interestAssetNo.Requery
17610           End If
17620         Case False
                ' ** Only assets for this account.
17630           If blnHasInt = False Then
17640             blnContinue = False
17650             DoCmd.Hourglass False
17660             MsgBox "The chosen account has no Interest assets.", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
17670           Else
17680             If .interestAssetNo.RowSource <> "qryJournal_Interest_02e" Then
                    ' ** qryJournal_Interest_02a (MasterAsset table linked to ActiveAssets, Account) linked
                    ' ** to AssetTypej, Interest = True, by specified FormRef('frmJournal_Sub2_Interest')
17690               .interestAssetNo.RowSource = "qryJournal_Interest_02e"
17700             Else
17710               .interestAssetNo.Requery
17720             End If
17730           End If
17740         End Select
17750       End If

17760       If blnContinue = True Then
17770         .interestAssetNo.Requery
17780         DoEvents
17790         If IsNull(.interestAssetNo) = True Then
17800           ExpandCombo .interestAssetNo  ' ** Module Procedure: modUtilities.
17810         Else
17820           varTmp00 = vbNullString
17830 On Error Resume Next
17840           varTmp00 = .interestAssetNo.text
17850           If ERR.Number <> 0 Then  ' ** When tabbing through quickly, it may not get here before it's already moved on!
17860 On Error GoTo ERRH
17870           Else
17880 On Error GoTo ERRH
17890             If varTmp00 = vbNullString Then
17900               ExpandCombo .interestAssetNo  ' ** Module Procedure: modUtilities.
17910             End If
17920           End If
17930         End If
17940       End If  ' ** blnContinue.

17950     End If  ' ** blnContinue.

17960     DoCmd.Hourglass False

17970   End With  ' ** Me.

EXITP:
17980   Set rst = Nothing
17990   Set qdf = Nothing
18000   Set dbs = Nothing
18010   Exit Sub

ERRH:
18020   DoCmd.Hourglass False
18030   THAT_PROC = THIS_PROC
18040   That_Erl = Erl
18050   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18060   Resume EXITP

End Sub

Private Sub interestAssetNo_BeforeUpdate(Cancel As Integer)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAssetNo_BeforeUpdate"

        Dim blnCurrID As Boolean
        Dim varTmp00 As Variant

18110   With Me
18120     If IsNull(.interestAssetNo) = False Then
18130       blnCurrID = .interestCurr_ID.Visible
18140       If blnCurrID = False Then
18150         varTmp00 = .interestAssetNo.Column(CBX_A_CURRID)
18160         If IsNull(varTmp00) = False Then
18170           If CLng(varTmp00) <> 150& Then
18180             Beep
18190             MsgBox "The chosen asset is in a foreign currency." & vbCrLf & vbCrLf & _
                    "You must check the 'Include Currency' checkbox on the Posting Date" & vbCrLf & _
                    "window before entering transactions not in U.S. dollars.", vbInformation + vbOKOnly, "Foreign Currency Asset"
18200             .interestAssetNo.Undo
18210             Cancel = -1
18220           End If
18230         End If
18240       End If
18250     End If
18260   End With

EXITP:
18270   Exit Sub

ERRH:
18280   Select Case ERR.Number
        Case Else
18290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
18300   End Select
18310   Resume EXITP

End Sub

Private Sub interestAssetNo_AfterUpdate()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAssetNo_AfterUpdate"

18410   With Me

          ' ** Do this whether or not Currency is visible.
18420     Select Case IsNull(.interestAssetNo)
          Case True
18430       lngCurrID = 150&  ' ** USD.
18440     Case False
18450       Select Case IsNull(.interestAssetNo.Column(CBX_A_CURRID))
            Case True
18460         lngCurrID = 150&  ' ** USD.
18470       Case False
18480         lngCurrID = .interestAssetNo.Column(CBX_A_CURRID)
18490       End Select
18500     End Select
18510     .interestCurr_ID = lngCurrID
18520     interestCurr_ID_AfterUpdate  ' ** Procedure: Below.

18530     If IsNull(.interestAssetNo) = False Then

18540       If .interestAssetNo <> .interestAssetNo.OldValue Then
18550         InterestChanged True  ' ** Procedure: Below.
18560       End If

18570       If gblnIncomeTaxCoding = True Then
18580         .cmbTaxCodes = .interestAssetNo.Column(CBX_A_TAX)  ' ** All AssetType-based Tax Codes are INC.
18590         .cmbTaxCodes.Requery
18600       End If

18610     End If

18620   End With

EXITP:
18630   Exit Sub

ERRH:
18640   Select Case ERR.Number
        Case Else
18650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18660   End Select
18670   Resume EXITP

End Sub

Private Sub interestAssetNo_KeyDown(KeyCode As Integer, Shift As Integer)

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAssetNo_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
18770       With Me
18780         intRetVal = 0
18790         .interestShareFace.SetFocus
18800       End With
18810     End Select
18820   End If

        ' ** Shift keys.
18830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18850       With Me
18860         intRetVal = 0
18870         .interestTransDate.SetFocus
18880       End With
18890     End Select
18900   End If

        ' ** Ctrl-Shift keys.
18910   If intCtrlDown And (Not intAltDown) And intShiftDown Then
18920     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18930       With Me
18940         intRetVal = 0
18950         .cmdCalendar1.SetFocus
18960       End With
18970     End Select
18980   End If

EXITP:
18990   KeyCode = intRetVal
19000   Exit Sub

ERRH:
19010   intRetVal = 0
19020   THAT_PROC = THIS_PROC
19030   That_Erl = Erl
19040   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19050   Resume EXITP

End Sub

Private Sub interestAssetNo_NotInList(NewData As String, Response As Integer)

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAssetNo_NotInList"

19110   With Me
19120     .interestAssetNo.SelStart = 0
19130     .interestAssetNo.SelLength = Len(.interestAssetNo.text)
19140   End With

EXITP:
19150   Exit Sub

ERRH:
19160   Select Case ERR.Number
        Case Else
19170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19180   End Select
19190   Resume EXITP

End Sub

Private Sub interestAssetNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAssetNo_MouseMove"

19210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
19220   Exit Sub

ERRH:
19230   Select Case ERR.Number
        Case Else
19240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
19250   End Select
19260   Resume EXITP

End Sub

Private Sub interestAssetNo_Exit(Cancel As Integer)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAssetNo_Exit"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim varTmp00 As Variant

        ' ** Set shares box value.
19310   With Me
19320     If IsNull(.interestAssetNo) = False Then
19330       If .interestAssetNo <> vbNullString And IsNull(.interestAccountNo) = False Then
19340         If .interestAccountNo <> vbNullString Then
19350           If blnExitingAssetNo = False Then
19360             Set dbs = CurrentDb
19370             Set rst = dbs.OpenRecordset("SELECT SUM(shareface) AS sumsf FROM activeAssets WHERE assetno = " & _
                    CStr(.interestAssetNo) & " AND accountno = '" & .interestAccountNo & "'")
19380             If rst.RecordCount > 0 Then
19390               rst.MoveFirst
19400               .interestShareFace = rst![sumsf]
19410             End If
19420             rst.Close
19430             dbs.Close
19440           End If
                ' ** Check curr_id.
19450           varTmp00 = .interestAssetNo.Column(CBX_A_CURRID)
19460           If IsNull(varTmp00) = False Then
19470             If .interestCurr_ID <> varTmp00 Then
19480               .interestCurr_ID = varTmp00
19490               interestCurr_ID_AfterUpdate  ' ** Procedure: Below.
19500             End If
19510           End If
19520         End If
19530       End If
19540     End If
19550   End With

EXITP:
19560   Set rst = Nothing
19570   Set dbs = Nothing
19580   Exit Sub

ERRH:
19590   Select Case ERR.Number
        Case Else
19600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19610   End Select
19620   Resume EXITP

End Sub

Private Sub interestAssetNo_LostFocus()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "interestAssetNo_LostFocus"

19710   With Me
19720     If Not blnExitingAssetNo Then

19730       VerifyAccountNo  ' ** Function: Below.

19740       gstrInterestAsset = .interestAssetNo.text
19750       If .interestAssetNo <> .interestAssetNo.OldValue Then
19760         InterestChanged True  ' ** Procedure: Below.
19770       End If

19780     End If
19790   End With

EXITP:
19800   Exit Sub

ERRH:
19810   Select Case ERR.Number
        Case Else
19820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19830   End Select
19840   Resume EXITP

End Sub

Private Sub interestShareFace_Enter()

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "interestShareFace_Enter"

19910   With Me
19920     .interestShareFace.Format = "General Number"
19930   End With

EXITP:
19940   Exit Sub

ERRH:
19950   Select Case ERR.Number
        Case Else
19960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19970   End Select
19980   Resume EXITP

End Sub

Private Sub interestShareFace_Exit(Cancel As Integer)

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "interestShareFace_Exit"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

20010   blnContinue = True

20020   VerifyAccountNo  ' ** Function: Below.

20030   With Me

20040     If blnAccountValidated Then

20050       .interestShareFace.SetFocus
20060       If IsNumeric(.interestShareFace.text) Then
20070         If CDbl(.interestShareFace.text) > 999999999.9999 Then  ' ** TO THE NINES! (Hopefully, not at sixes and sevens.)
20080           blnContinue = False
20090           MsgBox "Face is too large.", vbInformation + vbOKOnly, "Invalid Entry"
20100           .interestShareFace = vbNullString
20110           .interestShareFace = 0
20120           Cancel = -1
20130           .interestShareFace.SetFocus
20140         Else
20150           If .interestShareFace < 0 Then
20160             blnContinue = False
20170             MsgBox "You cannot enter a negative Face.", vbInformation + vbOKOnly, "Invalid Entry"
20180             .interestShareFace = CStr(0)
20190             Cancel = -1
20200             .interestShareFace.SetFocus
20210           End If
20220         End If
20230       Else
20240         If IsNull(.interestShareFace) = True Then
20250           blnContinue = False
20260           .interestShareFace = 0
20270         Else
20280           blnContinue = False
20290           MsgBox "Face must have a value.", vbInformation + vbOKOnly, "Entry Required"
20300           .interestShareFace = vbNullString
20310           DoCmd.RunCommand acCmdSaveRecord
20320           Cancel = -1
20330           .interestShareFace.SetFocus
20340         End If
20350       End If

20360       If blnContinue = True Then

20370         intPos01 = InStr(1, .interestShareFace, ".", vbTextCompare)
20380         If intPos01 > 0 Then
20390           If Len(.interestShareFace) - intPos01 > gintShareFaceDecimals Then
20400             blnContinue = False
20410             MsgBox "You can only enter a Face with up to " & Trim(str(gintShareFaceDecimals)) & " decimals.", _
                    vbInformation + vbOKOnly, "Invalid Entry"
20420             .interestShareFace = vbNullString
20430             Cancel = -1
20440             .interestShareFace.SetFocus
20450           Else
20460             .interestShareFace.Format = "#,###." & String(gintShareFaceDecimals, "0")
20470           End If
20480         Else
20490           .interestShareFace.Format = "#,###"
20500         End If

20510         If blnContinue = True Then
20520           gstrInterestShareFace = .interestShareFace.text
20530           If .interestShareFace <> .interestShareFace.OldValue Or IsNull(.interestShareFace.OldValue) Then
20540             InterestChanged True  ' ** Procedure: Below.
20550           End If
20560         End If

20570       End If

20580     Else
20590       If (gblnInterestChanged = False) Or (.NavigationButtons = True) Then
              ' ** Do nothing.
20600       Else
20610         .interestShareFace = vbNullString
20620         gstrInterestShareFace = vbNullString
20630       End If
20640     End If

20650   End With

EXITP:
20660   Exit Sub

ERRH:
20670   THAT_PROC = THIS_PROC
20680   That_Erl = Erl
20690   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20700   Resume EXITP

End Sub

Private Sub interestShareFace_LostFocus()

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "interestShareFace_LostFocus"

20810   VerifyAccountNo  ' ** Function: Below.

EXITP:
20820   Exit Sub

ERRH:
20830   Select Case ERR.Number
        Case Else
20840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20850   End Select
20860   Resume EXITP

End Sub

Private Sub interestICash_Change()

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "interestICash_Change"

        Dim blnContinue As Boolean
        Dim intPos01 As Integer

20910   With Me

20920     blnContinue = True

20930     lngCurrID = .interestCurr_ID.Column(CBX_C_CURRID)

20940     If IsNumeric(Rem_Dollar(.interestICash.text, lngCurrID)) Then  ' ** Module Function: modStringFuncs.
20950       If CDbl(Rem_Dollar(.interestICash.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
20960         blnContinue = False
20970         MsgBox "Income Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
20980         .interestICash = vbNullString
20990         .interestICash_usd = Null
21000         DoCmd.CancelEvent
21010         .interestICash.SetFocus
21020       End If
21030     Else
21040       If .interestICash.text <> "." Then  ' ** That is, if it's not numeric, don't give the message
21050         blnContinue = False               ' ** if all they've entered so far is a decimal point.
21060         MsgBox "Income Cash must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
21070         .interestICash = vbNullString
21080         .interestICash_usd = Null
21090         DoCmd.CancelEvent
21100         .interestICash.SetFocus
21110       End If
21120     End If

21130     If blnContinue = True Then
21140       intPos01 = InStr(1, .interestICash.text, ".", vbTextCompare)
21150       If intPos01 > 0 Then
21160         If Len(.interestICash.text) - intPos01 > 2 Then
21170           .interestICash = CDbl(Format(.interestICash, "0.00"))
21180           interestICash_AfterUpdate  ' ** Procedure: Above.
21190         End If
21200       End If
21210     End If

21220   End With

EXITP:
21230   Exit Sub

ERRH:
21240   Select Case ERR.Number
        Case Else
21250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21260   End Select
21270   Resume EXITP

End Sub

Private Sub interestICash_AfterUpdate()

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "interestICash_AfterUpdate"

        Dim dblICash As Double
        Dim varTmp00 As Variant

21310   With Me
21320     If .interestCurr_ID.Visible = True Then
21330       .interestICash_usd = Null
21340       If .interestICash_usd.Visible = False Then .interestICash_usd.Visible = True
21350       If IsNull(.interestICash) = False Then
21360         If .interestICash <> 0 Then
21370           If IsNull(.interestCurr_ID.Column(CBX_C_CURRID)) = False Then
21380             lngCurrID = .interestCurr_ID.Column(CBX_C_CURRID)
21390             If lngCurrID <> 150& Then  ' ** USD.
21400               varTmp00 = .interestCurr_ID.Column(CBX_C_RATE2)
21410               If IsNull(varTmp00) = False Then
21420                 If varTmp00 > 0 Then
21430                   dblICash = .interestICash
21440                   dblICash = (dblICash * varTmp00)
21450                   dblICash = Round(dblICash, 2)
21460                   .interestICash_usd = dblICash
21470                 End If
21480               End If
21490             End If
21500           End If
21510         End If
21520       End If
21530     End If
21540     DoEvents
21550   End With

EXITP:
21560   Exit Sub

ERRH:
21570   Select Case ERR.Number
        Case Else
21580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
21590   End Select
21600   Resume EXITP

End Sub

Private Sub interestICash_Exit(Cancel As Integer)

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "interestICash_Exit"

        Dim intPos01 As Integer
        Dim dblTmp01 As Double
        Dim blnContinue As Boolean

21710   With Me

21720     blnContinue = True

21730     lngCurrID = .interestCurr_ID.Column(CBX_C_CURRID)

21740     If IsNumeric(.interestICash) Then
21750       If lngCurrID <> 150& Then
21760         dblTmp01 = .interestICash_usd
21770       Else
21780         dblTmp01 = .interestICash
21790       End If
21800       If dblTmp01 > 999999999.99 Then 'CDbl(.interestICash.text) > 999999999.99 Then
21810         blnContinue = False
21820         MsgBox "Income Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
21830         .interestICash = vbNullString
21840         .interestICash_usd = Null
21850         Cancel = -1
21860         .interestICash.SetFocus
21870       End If
21880     Else
21890       blnContinue = False
21900       MsgBox "Income Cash must have a value.", vbInformation + vbOKOnly, "Entry Required"
21910       .interestICash = vbNullString
21920       .interestICash_usd = Null
21930       Cancel = -1
21940       .interestICash.SetFocus
21950     End If

21960     If blnContinue = True Then
21970       intPos01 = InStr(1, Nz(.interestICash, 0), ".", vbTextCompare)
21980       If intPos01 > 0 Then
21990         If Len(.interestICash) - intPos01 > 2 Then
22000           blnContinue = False
22010           MsgBox "You can only enter Income Cash with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
22020           .interestICash = vbNullString
22030           .interestICash_usd = Null
22040           Cancel = -1
22050           .interestICash.SetFocus
22060         End If
22070       End If
22080     End If

22090     If blnContinue = True Then
22100       If .interestICash.text = vbNullString Or .interestICash.text = "0" Then
22110         blnContinue = False
22120         MsgBox "Income Cash must be greater than zero.", vbInformation + vbOKOnly, "Entry Required"
22130         Cancel = -1
22140         .interestICash.SetFocus
22150       End If
22160     End If

22170     If blnContinue = True Then
22180       gstrInterestICash = .interestICash.text
22190       gstrInterestICash = Rem_Dollar(gstrInterestICash, lngCurrID)  ' ** Module Function: modStringFuncs.
22200       If .interestICash <> .interestICash.OldValue Then
22210         InterestChanged True  ' ** Procedure: Below.
22220       End If
22230     End If

22240   End With

EXITP:
22250   Exit Sub

ERRH:
22260   THAT_PROC = THIS_PROC
22270   That_Erl = Erl
22280   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
22290   Resume EXITP

End Sub

Private Sub interestICash_LostFocus()

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "interestICash_LostFocus"

22310   VerifyAccountNo  ' ** Function: Below.

EXITP:
22320   Exit Sub

ERRH:
22330   Select Case ERR.Number
        Case Else
22340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22350   End Select
22360   Resume EXITP

End Sub

Private Sub interestCurr_ID_cmd_DblClick(Cancel As Integer)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "interestCurr_ID_cmd_DblClick"

22410   With Me
22420     If IsNull(.interestCurr_Date) = True Then
22430       .interestCurr_Date = .interestCurr_ID.Column(CBX_C_DATE)
22440     End If
22450     DoEvents
22460     Select Case .interestCurr_Date.Visible
          Case True
22470       .interestCurr_Date.Visible = False
22480     Case False
22490       .interestCurr_Date.Visible = True
22500     End Select
22510   End With

EXITP:
22520   Exit Sub

ERRH:
22530   Select Case ERR.Number
        Case Else
22540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22550   End Select
22560   Resume EXITP

End Sub

Private Sub interestCurr_ID_AfterUpdate()

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "interestCurr_ID_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String, strTmp03 As String, intTmp04 As Integer, lngTmp06 As Long
        Dim lngX As Long

22610   With Me

22620     If IsNull(.interestCurr_ID) = False Then

22630       lngCurrID = .interestCurr_ID.Column(CBX_C_CURRID)

22640       blnFound = False
22650       If lngCurrID <> 150& Then
22660         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
22670           arr_varNoChar = .Parent.NoChar_Get  ' ** Form Function: frmJournal.
22680           lngNoChars = UBound(arr_varNoChar, 2) + 1&
22690         End If
22700         For lngX = 0& To (lngNoChars - 1&)
22710           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
22720             blnFound = True
22730           End If
22740         Next
22750       End If

22760       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
22770         .interestICash.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
22780         .interestICash.FontSize = 10
22790         .interestICash.FontBold = True
22800       Case False
22810         If .interestICash.FontName <> "Fixedsys" Then
22820           .interestICash.FontName = "Fixedsys"
22830           .interestICash.FontSize = 9
22840           .interestICash.FontBold = False
22850         End If
22860       End Select

22870       If lngCurrID = 150& Then  ' ** USD.
22880         strTmp01 = "Currency"
22890         intTmp04 = 2
22900         lngTmp06 = CLR_WHT
22910       Else
22920         strTmp03 = .interestCurr_ID.Column(CBX_C_SYM)  ' ** Symbol.
22930         If .interestCurr_ID.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
22940           intTmp04 = 4
22950         Else
22960           intTmp04 = .interestCurr_ID.Column(CBX_C_DEC)  ' ** Decimal places.
22970         End If
22980         intLen = Len(strTmp03)
22990         Select Case intLen
              Case 1
                'strTmp03 = strTmp03
23000         Case 2
23010           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2)
23020         Case 3
23030           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3)
23040         Case 4
23050           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & "\" & Mid(strTmp03, 4)
23060         Case 5
23070           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & _
                  "\" & Mid(strTmp03, 4, 1) & "\" & Mid(strTmp03, 5)
23080         End Select
23090         Select Case intTmp04
              Case 0
23100           strTmp01 = strTmp03 & "#,##0;(" & strTmp03 & "#,##0);" & strTmp03 & "0;" & strTmp03 & "0"
23110         Case 1  ' ** None have this currently.
23120           strTmp01 = strTmp03 & "#,##0.0;(" & strTmp03 & "#,##0.0);" & strTmp03 & "0.0;" & strTmp03 & "0.0"
23130         Case 2
23140           strTmp01 = strTmp03 & "#,##0.00;(" & strTmp03 & "#,##0.00);" & strTmp03 & "0.00;" & strTmp03 & "0.00"
23150         Case 3
23160           strTmp01 = strTmp03 & "#,##0.000;(" & strTmp03 & "#,##0.000);" & strTmp03 & "0.000;" & strTmp03 & "0.000"
23170         End Select
23180         lngTmp06 = CLR_VLTGRN
23190       End If
23200       .interestICash.Format = strTmp01
23210       .interestICash.DecimalPlaces = intTmp04
23220       .interestICash.BackColor = lngTmp06

23230     End If

23240     If lngCurrID <> .interestCurr_ID.OldValue Then
23250       InterestChanged True  ' ** Procedure: Below.
23260     End If

23270   End With

EXITP:
23280   Exit Sub

ERRH:
23290   THAT_PROC = THIS_PROC
23300   That_Erl = Erl
23310   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23320   Resume EXITP

End Sub

Private Sub interestCurr_ID_NotInList(NewData As String, Response As Integer)

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "interestCurr_ID_NotInList"

23410   With Me
23420     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
23430     .interestCurr_ID.SetFocus
23440     Response = acDataErrContinue
23450   End With

EXITP:
23460   Exit Sub

ERRH:
23470   THAT_PROC = THIS_PROC
23480   That_Erl = Erl
23490   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23500   Resume EXITP

End Sub

Private Sub interestCurr_ID_Exit(Cancel As Integer)

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "interestCurr_ID_Exit"

23610   With Me
23620     If .interestCurr_ID <> .interestCurr_ID.OldValue Then
23630       InterestChanged True  ' ** Procedure: Below.
23640     End If
23650   End With

EXITP:
23660   Exit Sub

ERRH:
23670   THAT_PROC = THIS_PROC
23680   That_Erl = Erl
23690   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23700   Resume EXITP

End Sub

Private Sub interestDescription_AfterUpdate()

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "interestDescription_AfterUpdate"

23810   With Me
23820     If .interestDescription <> .interestDescription.OldValue Then
23830       InterestChanged True  ' ** Procedure: Below.
23840     End If
23850   End With

EXITP:
23860   Exit Sub

ERRH:
23870   Select Case ERR.Number
        Case Else
23880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23890   End Select
23900   Resume EXITP

End Sub

Private Sub InterestDescription_KeyDown(KeyCode As Integer, Shift As Integer)

24000 On Error GoTo ERRH

        Const THIS_PROC As String = "InterestDescription_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24020   intShiftDown = (Shift And acShiftMask) > 0
24030   intAltDown = (Shift And acAltMask) > 0
24040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24070       With Me
24080         intRetVal = 0
24090         If .tglInterestReinvest.Visible = True And .tglInterestReinvest.Enabled = True Then
24100           .tglInterestReinvest.SetFocus
24110         ElseIf .cmbRevenueCodes.Visible = True And .cmbRevenueCodes.Enabled = True Then
24120           .cmbRevenueCodes.SetFocus
24130         ElseIf .cmbTaxCodes.Visible = True And .cmbTaxCodes.Enabled = True Then
24140           .cmbTaxCodes.SetFocus
24150         ElseIf .cmdInterestOK.Enabled = True Then
24160           .cmdInterestOK.SetFocus
24170         ElseIf .cmdInterestCancel.Enabled = True Then
24180           .cmdInterestCancel.SetFocus
24190         ElseIf .cmdInterestClose.Enabled = True Then
24200           .cmdInterestClose.SetFocus
24210         ElseIf .interestAccountNo.Enabled = True Then
24220           .interestAccountNo.SetFocus
24230         Else
24240           .interestTransDate.SetFocus
24250         End If
24260       End With
24270     End Select
24280   End If

        ' ** Shift keys.
24290   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24300     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24310       With Me
24320         intRetVal = 0
24330         .interestICash.SetFocus
24340       End With
24350     End Select
24360   End If

EXITP:
24370   KeyCode = intRetVal
24380   Exit Sub

ERRH:
24390   intRetVal = 0
24400   THAT_PROC = THIS_PROC
24410   That_Erl = Erl
24420   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24430   Resume EXITP

End Sub

Private Sub interestDescription_LostFocus()

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "interestDescription_LostFocus"

24510   VerifyAccountNo  ' ** Function: Below.

EXITP:
24520   Exit Sub

ERRH:
24530   Select Case ERR.Number
        Case Else
24540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24550   End Select
24560   Resume EXITP

End Sub

Private Sub tglInterestReinvest_GotFocus()

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "tglInterestReinvest_GotFocus"

24610   With Me
24620     blnInterestReinvest_Focus = True
24630     Select Case .posted
          Case True
24640       .tglInterestReinvest_true_raised_focus_dots_img.Visible = True
24650       .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = False
24660     Case False
24670       .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = True
24680       .tglInterestReinvest_true_raised_focus_dots_img.Visible = False
24690     End Select
24700     .tglInterestReinvest_false_raised_img.Visible = False
24710     .tglInterestReinvest_false_raised_focus_img.Visible = False
24720     .tglInterestReinvest_false_raised_focus_dots_img.Visible = False
24730     .tglInterestReinvest_false_sunken_focus_dots_img.Visible = False
24740     .tglInterestReinvest_false_raised_img_dis.Visible = False
24750     .tglInterestReinvest_true_raised_img.Visible = False
24760     .tglInterestReinvest_true_raised_focus_img.Visible = False
24770     .tglInterestReinvest_true_sunken_focus_dots_img.Visible = False
24780     .tglInterestReinvest_true_raised_img_dis.Visible = False
24790   End With

EXITP:
24800   Exit Sub

ERRH:
24810   THAT_PROC = THIS_PROC
24820   That_Erl = Erl
24830   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24840   Resume EXITP

End Sub

Private Sub tglInterestReinvest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "tglInterestReinvest_MouseDown"

24910   With Me
24920     blnInterestReinvest_MouseDown = True
24930     Select Case .posted
          Case True
24940       .tglInterestReinvest_true_sunken_focus_dots_img.Visible = True
24950       .tglInterestReinvest_false_sunken_focus_dots_img.Visible = False
24960     Case False
24970       .tglInterestReinvest_false_sunken_focus_dots_img.Visible = True
24980       .tglInterestReinvest_true_sunken_focus_dots_img.Visible = False
24990     End Select
25000     .tglInterestReinvest_false_raised_img.Visible = False
25010     .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = False
25020     .tglInterestReinvest_false_raised_focus_img.Visible = False
25030     .tglInterestReinvest_false_raised_focus_dots_img.Visible = False
25040     .tglInterestReinvest_false_raised_img_dis.Visible = False
25050     .tglInterestReinvest_true_raised_img.Visible = False
25060     .tglInterestReinvest_true_raised_focus_img.Visible = False
25070     .tglInterestReinvest_true_raised_focus_dots_img.Visible = False
25080     .tglInterestReinvest_true_raised_img_dis.Visible = False
25090   End With

EXITP:
25100   Exit Sub

ERRH:
25110   THAT_PROC = THIS_PROC
25120   That_Erl = Erl
25130   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
25140   Resume EXITP

End Sub

Public Sub tglInterestReinvest_Click()
' ** I'm going to try using [posted] as a signal (since it's not used for anything else).

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "tglInterestReinvest_Click"

25210   With Me
          ' ** Do I need to do anything with the images, or will that be handled by the other events?
25220     Select Case .posted
          Case True  ' ** If it's True, flip to False, and vice versa.
25230       blnGoToIntReinvest = False
25240       .posted = False
25250       DoCmd.RunCommand acCmdSaveRecord
25260     Case False
25270       blnGoToIntReinvest = True
25280       .posted = True
25290       DoCmd.RunCommand acCmdSaveRecord
25300     End Select
25310     InterestChanged True  ' ** Procedure: Below.
25320     DoEvents
25330   End With

EXITP:
25340   Exit Sub

ERRH:
25350   THAT_PROC = THIS_PROC
25360   That_Erl = Erl
25370   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
25380   Resume EXITP

End Sub

Private Sub tglInterestReinvest_KeyDown(KeyCode As Integer, Shift As Integer)

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "tglInterestReinvest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25420   intShiftDown = (Shift And acShiftMask) > 0
25430   intAltDown = (Shift And acAltMask) > 0
25440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25470       With Me
25480         intRetVal = 0
25490         If .cmbRevenueCodes.Enabled = True Then
25500           .cmbRevenueCodes.SetFocus
25510         ElseIf .cmdInterestOK.Enabled = True Then
25520           .cmdInterestOK.SetFocus
25530         ElseIf .cmdInterestCancel.Enabled = True Then
25540           .cmdInterestCancel.SetFocus
25550         ElseIf .cmdInterestClose.Enabled = True Then
25560           .cmdInterestClose.SetFocus
25570         ElseIf .interestAccountNo.Enabled = True Then
25580           .interestAccountNo.SetFocus
25590         Else
25600           .interestTransDate.SetFocus
25610         End If
25620       End With
25630     End Select
25640   End If

        ' ** Shift keys.
25650   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25670       With Me
25680         intRetVal = 0
25690         .interestDescription.SetFocus
25700       End With
25710     End Select
25720   End If

EXITP:
25730   KeyCode = intRetVal
25740   Exit Sub

ERRH:
25750   intRetVal = 0
25760   THAT_PROC = THIS_PROC
25770   That_Erl = Erl
25780   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
25790   Resume EXITP

End Sub

Private Sub tglInterestReinvest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "tglInterestReinvest_MouseMove"

25810   With Me
25820     If blnInterestReinvest_MouseDown = False Then
25830       Select Case .posted
            Case True
25840         Select Case blnInterestReinvest_Focus
              Case True
25850           .tglInterestReinvest_true_raised_focus_dots_img.Visible = True
25860           .tglInterestReinvest_true_raised_focus_img.Visible = False
25870         Case False
25880           .tglInterestReinvest_true_raised_focus_img.Visible = True
25890           .tglInterestReinvest_true_raised_focus_dots_img.Visible = False
25900         End Select
25910         .tglInterestReinvest_false_raised_focus_img.Visible = False
25920         .tglInterestReinvest_false_raised_focus_dots_img.Visible = False
25930       Case False
25940         Select Case blnInterestReinvest_Focus
              Case True
25950           .tglInterestReinvest_false_raised_focus_dots_img.Visible = True
25960           .tglInterestReinvest_false_raised_focus_img.Visible = False
25970         Case False
25980           .tglInterestReinvest_false_raised_focus_img.Visible = True
25990           .tglInterestReinvest_false_raised_focus_dots_img.Visible = False
26000         End Select
26010         .tglInterestReinvest_true_raised_focus_img.Visible = False
26020         .tglInterestReinvest_true_raised_focus_dots_img.Visible = False
26030       End Select
26040       .tglInterestReinvest_false_raised_img.Visible = False
26050       .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = False
26060       .tglInterestReinvest_false_raised_img_dis.Visible = False
26070       .tglInterestReinvest_false_sunken_focus_dots_img.Visible = False
26080       .tglInterestReinvest_true_raised_img.Visible = False
26090       .tglInterestReinvest_true_sunken_focus_dots_img.Visible = False
26100       .tglInterestReinvest_true_raised_img_dis.Visible = False
26110     End If
26120   End With

EXITP:
26130   Exit Sub

ERRH:
26140   THAT_PROC = THIS_PROC
26150   That_Erl = Erl
26160   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
26170   Resume EXITP

End Sub

Private Sub tglInterestReinvest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "tglInterestReinvest_MouseUp"

26210   With Me
26220     Select Case .posted
          Case True
26230       .tglInterestReinvest_true_raised_focus_dots_img.Visible = True
26240       .tglInterestReinvest_false_raised_focus_dots_img.Visible = False
26250     Case False
26260       .tglInterestReinvest_false_raised_focus_dots_img.Visible = True
26270       .tglInterestReinvest_true_raised_focus_dots_img.Visible = False
26280     End Select
26290     .tglInterestReinvest_false_raised_img.Visible = False
26300     .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = False
26310     .tglInterestReinvest_false_raised_focus_img.Visible = False
26320     .tglInterestReinvest_false_raised_img_dis.Visible = False
26330     .tglInterestReinvest_false_sunken_focus_dots_img.Visible = False
26340     .tglInterestReinvest_true_raised_img.Visible = False
26350     .tglInterestReinvest_true_raised_focus_img.Visible = False
26360     .tglInterestReinvest_true_sunken_focus_dots_img.Visible = False
26370     .tglInterestReinvest_true_raised_img_dis.Visible = False
26380     blnInterestReinvest_MouseDown = False
26390   End With

EXITP:
26400   Exit Sub

ERRH:
26410   THAT_PROC = THIS_PROC
26420   That_Erl = Erl
26430   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
26440   Resume EXITP

End Sub

Private Sub tglInterestReinvest_LostFocus()

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "tglInterestReinvest_LostFocus"

26510   With Me
26520     Select Case .posted
          Case True
26530       .tglInterestReinvest_true_raised_img.Visible = True
26540       .tglInterestReinvest_false_raised_img.Visible = False
26550     Case False
26560       .tglInterestReinvest_false_raised_img.Visible = True
26570       .tglInterestReinvest_true_raised_img.Visible = False
26580     End Select
26590     .tglInterestReinvest_false_raised_semifocus_dots_img.Visible = False
26600     .tglInterestReinvest_false_raised_focus_img.Visible = False
26610     .tglInterestReinvest_false_raised_focus_dots_img.Visible = False
26620     .tglInterestReinvest_false_sunken_focus_dots_img.Visible = False
26630     .tglInterestReinvest_false_raised_img_dis.Visible = False
26640     .tglInterestReinvest_true_raised_focus_img.Visible = False
26650     .tglInterestReinvest_true_raised_focus_dots_img.Visible = False
26660     .tglInterestReinvest_true_sunken_focus_dots_img.Visible = False
26670     .tglInterestReinvest_true_raised_img_dis.Visible = False
26680     blnInterestReinvest_Focus = False
26690   End With

EXITP:
26700   Exit Sub

ERRH:
26710   THAT_PROC = THIS_PROC
26720   That_Erl = Erl
26730   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
26740   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Enter()

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Enter"

26810   With Me
26820     Select Case IsNull(.cmbRevenueCodes)
          Case True
26830       ExpandCombo .cmbRevenueCodes  ' ** Module Procedure: modUtilities.
26840     Case False
26850       If .cmbRevenueCodes = REVID_INC Or .cmbRevenueCodes = REVID_EXP Then  ' ** For just the Unspecified's.
26860         ExpandCombo .cmbRevenueCodes  ' ** Module Procedure: modUtilities.
26870       End If
26880     End Select
26890   End With

EXITP:
26900   Exit Sub

ERRH:
26910   Select Case ERR.Number
        Case Else
26920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26930   End Select
26940   Resume EXITP

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

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Change"

27010   With Me
27020     If .cmbRevenueCodes <> .cmbRevenueCodes.OldValue Then
27030       InterestChanged True  ' ** Procedure: Below.
27040     End If
27050   End With

EXITP:
27060   Exit Sub

ERRH:
27070   Select Case ERR.Number
        Case Else
27080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27090   End Select
27100   Resume EXITP

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

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_AfterUpdate"

        Dim strRevCode As String, lngTaxcode As Long

        ' ** Make sure tax code matches selected revenue code.
27210   With Me
27220     If IsNull(.cmbRevenueCodes) = False Then
27230       If .cmbRevenueCodes > 0 Then
27240         If gblnLinkRevTaxCodes = True Then
27250           strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
27260           If IsNull(.cmbTaxCodes) = True Then
27270             If strRevCode = "I" Then
27280               .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
27290             ElseIf strRevCode = "E" Then
27300               .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
27310             End If
27320           Else
27330             lngTaxcode = .cmbTaxCodes
27340             If lngTaxcode = 0& Then
27350               If strRevCode = "I" Then
27360                 .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
27370               ElseIf strRevCode = "E" Then
27380                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
27390               End If
27400             Else
27410               If strRevCode = "I" And .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
                      'MsgBox "This will change..."
27420                 .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
27430               ElseIf strRevCode = "E" And .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
                      'MsgBox "This will change..."
27440                 .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
27450               End If
27460             End If
27470           End If
27480         End If
27490       Else
27500         .cmbRevenueCodes = REVID_INTINC  ' ** Interest Income.
27510       End If
27520     Else
27530       .cmbRevenueCodes = REVID_INTINC  ' ** Interest Income.
27540     End If
27550   End With

EXITP:
27560   Exit Sub

ERRH:
27570   Select Case ERR.Number
        Case Else
27580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27590   End Select
27600   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_LostFocus()

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_LostFocus"

27710   VerifyAccountNo  ' ** Function: Below.

EXITP:
27720   Exit Sub

ERRH:
27730   Select Case ERR.Number
        Case Else
27740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27750   End Select
27760   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Enter()

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Enter"

27810   With Me
27820     Select Case IsNull(.cmbTaxCodes)
          Case True
27830       ExpandCombo .cmbTaxCodes  ' ** Module Procedure: modUtilities.
27840     Case False
            'If .cmbTaxCodes = TAXID_INC Or .cmbTaxCodes = TAXID_DED Then
27850       ExpandCombo .cmbTaxCodes  ' ** Module Procedure: modUtilities.
            'End If
27860     End Select
27870   End With

EXITP:
27880   Exit Sub

ERRH:
27890   Select Case ERR.Number
        Case Else
27900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27910   End Select
27920   Resume EXITP

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

28000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Change"

28010   With Me
28020     If .cmbTaxCodes <> .cmbTaxCodes.OldValue Then
28030       InterestChanged True  ' ** Procedure: Below.
28040     End If
28050   End With

EXITP:
28060   Exit Sub

ERRH:
28070   Select Case ERR.Number
        Case Else
28080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28090   End Select
28100   Resume EXITP

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

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_AfterUpdate"

        Dim strRevCode As String, lngTaxcode As Long

        ' ** Make sure revenue code matches selected tax code.
28210   With Me
28220     If IsNull(.cmbTaxCodes) = False Then
28230       If .cmbTaxCodes > 0& Then
28240         If gblnLinkRevTaxCodes = True Then
28250           lngTaxcode = .cmbTaxCodes
28260           strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
28270           If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
28280             If IsNull(.cmbRevenueCodes) = True Then
28290               .cmbRevenueCodes = REVID_INTINC  ' ** Interest Income.
28300             Else
28310               If .cmbRevenueCodes = 0 Then
28320                 .cmbRevenueCodes = REVID_INTINC  ' ** Interest Income.
28330               Else
28340                 If strRevCode = "E" Then
                        'MsgBox "This will change..."
28350                   .cmbRevenueCodes = REVID_INTINC  ' ** Interest Income.
28360                 End If
28370               End If
28380             End If
28390           ElseIf .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
28400             If IsNull(.cmbRevenueCodes) = True Then
28410               .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
28420             Else
28430               If .cmbRevenueCodes = 0 Then
28440                 .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
28450               Else
28460                 If strRevCode = "I" Then
                        'MsgBox "This will change..."
28470                   .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
28480                 End If
28490               End If
28500             End If
28510           End If
28520         End If
28530       Else
28540         .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
28550       End If
28560     Else
28570       .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
28580     End If
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

Private Sub cmbTaxCodes_LostFocus()

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_LostFocus"

28710   VerifyAccountNo  ' ** Function: Below.

EXITP:
28720   Exit Sub

ERRH:
28730   Select Case ERR.Number
        Case Else
28740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28750   End Select
28760   Resume EXITP

End Sub

Public Sub InterestChanged(blnChanged As Boolean)

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "InterestChanged"

28810   With Me
28820     Select Case blnChanged
          Case True
28830       gblnInterestChanged = True
28840       .NavigationButtons = False
28850       DoCmd.SelectObject acForm, .Parent.Name, False
28860       With .Parent
28870         .FocusHolder.SetFocus
28880         .opgJournal.Enabled = False
28890         .opgJournal_optInterest_lbl_box.Visible = True
28900         .cmdSwitch.Enabled = False
28910         .cmdSwitch_raised_img_dis.Visible = True
28920         .cmdSwitch_raised_img.Visible = False
28930         .cmdSwitch_raised_semifocus_dots_img.Visible = False
28940         .cmdSwitch_raised_focus_img.Visible = False
28950         .cmdSwitch_raised_focus_dots_img.Visible = False
28960         .cmdSwitch_sunken_focus_dots_img.Visible = False
28970         .frmJournal_Sub2_Interest.SetFocus
28980       End With
28990       .cmdInterestClose.Enabled = False
29000       .cmdInterestCancel.Enabled = True
29010       .cmdInterestOK.Enabled = True
29020       .Parent.NavVis False  ' ** Form Procedure: frmJournal.
29030     Case False
29040       gblnInterestChanged = False
29050       .NavigationButtons = True
29060       DoCmd.SelectObject acForm, .Parent.Name, False
29070       With .Parent
29080         .opgJournal.Enabled = True
29090         .opgJournal_optInterest_lbl_box.Visible = False
29100         .cmdSwitch.Enabled = True
29110         .cmdSwitch_raised_img.Visible = True
29120         .cmdSwitch_raised_img_dis.Visible = False
29130         .cmdSwitch_raised_semifocus_dots_img.Visible = False
29140         .cmdSwitch_raised_focus_img.Visible = False
29150         .cmdSwitch_raised_focus_dots_img.Visible = False
29160         .cmdSwitch_sunken_focus_dots_img.Visible = False
29170       End With
29180       .cmdInterestClose.Enabled = True
29190       .cmdInterestCancel.Enabled = False
29200       .cmdInterestOK.Enabled = False
29210       .Parent.NavVis True  ' ** Form Procedure: frmJournal.
29220     End Select
29230   End With

EXITP:
29240   Exit Sub

ERRH:
29250   THAT_PROC = THIS_PROC
29260   That_Erl = Erl
29270   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
29280   Resume EXITP

End Sub

Public Function VerifyAccountNo() As Boolean

29300 On Error GoTo ERRH

        Const THIS_PROC As String = "VerifyAccountNo"

        Dim strAccountNo As String
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

29310   blnRetVal = True

29320   With Me
29330     If (((gblnInterestValidated = False And gblnInterestChanged = True) Or blnCheckingAccount = True) And (lngErrCnt < 3&)) Then
29340       blnAccountValidated = False     ' ** Unless proven otherwise.
29350       Select Case IsNull(interestAccountNo)
            Case True
29360         blnRetVal = False
29370         lngErrCnt = lngErrCnt + 1&
29380         .interestAccountNo.SetFocus
29390         MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Invalid Entry"
              'DoCmd.CancelEvent
29400         gblnInterestValidated = False
29410       Case False
              ' ** Check to see if it is a valid account.
29420         lngErrCnt = 0&
29430         Select Case gblnAccountNoWithType
              Case True
29440           strAccountNo = Left(.interestAccountNo, 2) & Mid(.interestAccountNo, 3)
29450         Case False
29460           strAccountNo = .interestAccountNo
29470         End Select
29480         varTmp00 = DLookup("[accountno]", "account", "[accountno] = '" & strAccountNo & "' And ([dateclosed] Is Null)")
29490         Select Case IsNull(varTmp00)
              Case True
29500           blnRetVal = False
29510           MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Invalid Entry"
29520           gblnInterestValidated = False
29530         Case False
29540           blnAccountValidated = True
29550           gblnInterestValidated = True
29560           blnStartTrans = False
29570         End Select
29580       End Select
29590     Else
29600       lngErrCnt = 0&
29610     End If
29620   End With

EXITP:
29630   VerifyAccountNo = blnRetVal
29640   Exit Function

ERRH:
29650   blnRetVal = False
29660   THAT_PROC = THIS_PROC
29670   That_Erl = Erl
29680   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
29690   Resume EXITP

End Function

Public Sub Map_NewRec()

29700 On Error GoTo ERRH

        Const THIS_PROC As String = "Map_NewRec"

29710   With Me
29720 On Error Resume Next
29730     DoCmd.RunCommand acCmdRecordsGoToNew
29740     If ERR.Number <> 0 Then
29750       blnMapNewRec = True
29760       .TimerInterval = 500&
29770     End If
29780 On Error GoTo ERRH
29790   End With

EXITP:
29800   Exit Sub

ERRH:
29810   THAT_PROC = THIS_PROC
29820   That_Erl = Erl
29830   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
29840   Resume EXITP

End Sub

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

29900 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

29910   Select Case intWhere
        Case acCmdRecordsGoToFirst
29920     DoCmd.RunCommand acCmdRecordsGoToFirst
29930   Case acCmdRecordsGoToLast
29940     DoCmd.RunCommand acCmdRecordsGoToLast
29950   Case acCmdRecordsGoToPrevious
29960     DoCmd.RunCommand acCmdRecordsGoToPrevious
29970   Case acCmdRecordsGoToNext
29980     DoCmd.RunCommand acCmdRecordsGoToNext
29990   Case acCmdRecordsGoToNew
30000     DoCmd.RunCommand acCmdRecordsGoToNew
30010   Case Else
30020     If IsMissing(varID) = False Then
30030       With Me
30040         Set rst = .RecordsetClone
30050         With rst
30060           .FindFirst "[interestID] = " & CStr(varID)
30070           If .NoMatch = False Then
30080             Me.Bookmark = .Bookmark
30090           End If
30100           .Close
30110         End With
30120       End With
30130     End If
30140   End Select

EXITP:
30150   Set rst = Nothing
30160   Exit Sub

ERRH:
30170   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
30180   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
30190   Case Else
30200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30210   End Select
30220   Resume EXITP

End Sub
