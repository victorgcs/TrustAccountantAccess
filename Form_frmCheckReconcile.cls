VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckReconcile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckReconcile"

'VGC 10/03/2017: CHANGES!

' #########################
' ## Use VBA_RenumErrh().  39230
' #########################

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Out Checks:       D {frmCheckReconcile_Sub_OutChecks}
' **   BS Debits:        E {frmCheckReconcile_Sub_BSDebits}
' **   Total Cash:       H {opgCheckingType_optTotalCash}
' **   TA Credits:       I {frmCheckReconcile_Sub_TACredits}
' **   Uncheck All:      K {cmdCheckNone}
' **   Check All:        L {cmdCheckAll}
' **   Name:             M {opgAccountSource_optName}
' **   BS Balance:       N {cracct_bsbalance_display}
' **   TA Debits:        O {frmCheckReconcile_Sub_TADebits}
' **   Print Report:     P {cmdPrintReport}
' **   BS Credits:       R {frmCheckReconcile_Sub_BSCredits}
' **   Asset:            S {cmbAssets}
' **   Balance Through:  T {DateEnd}
' **   Update:           U {cmdUpdate}
' **   Individual:       V {opgCheckingType_optIndividual}
' **   Preview Report:   W {cmdPreviewReport}
' **   Exit:             X {cmdClose}
'Unused: F G J Q Y Z

' ** Shortcut F-Keys responsive from this form:
' **   Recalc Tots:      F5 {RecalcTots_CR}
' **   Date Picker:      F7 {cmdCalendar}

' ** Shortcut Ctrl keys responsive from this form:
' **   2nd Asset:        A {cmdAddAsset}
' **   Delete:           D {cmdDelete}
' **   Next Asset:       N {cmdAssetNext}
' **   Previous Asset:   P {cmdAssetPrevious}
' **   Clear:            R {cmdClear}
' **   Save:             S {cmdSave here, and on each subform}

' ** Shortcut Alt-Shift keys responsive from this form:
' **   Show Paid:        H {chkShowPaid}
' **   By Cusip:         I {opgAssetSource_optCusip}
' **   By Name:          N {opgAssetSource_optName}
' **   By Type:          Y {opgAssetSource_optType}

'BSTotalCredits : =IIf([cmbAccounts]="CRTC01",DLookUp("CalcAmount","qryCheckReconcile_BSCreditTotal_01b"),DLookUp("CalcAmount","qryCheckReconcile_BSCreditTotal_01a"))
'BSTotalDebits  : =IIf([cmbAccounts]="CRTC01",DLookUp("CalcAmount","qryCheckReconcile_BSDebitTotal_01b"),DLookUp("CalcAmount","qryCheckReconcile_BSDebitTotal_01a"))
'TATotalCredits : =IIf([cmbAccounts]="CRTC01",DLookUp("CalcAmount","qryCheckReconcile_TACreditTotal_01b"),DLookUp("CalcAmount","qryCheckReconcile_TACreditTotal_01a"))
'TATotalDebits  : =IIf([cmbAccounts]="CRTC01",DLookUp("CalcAmount","qryCheckReconcile_TADebitTotal_01b"),DLookUp("CalcAmount","qryCheckReconcile_TADebitTotal_01a"))
'TotalOutstandingChecks : =IIf([cmbAccounts]="CRTC01",DLookUp("CalcTotal","qryCheckReconcile_OutChecks_09b"),DLookUp("CalcTotal","qryCheckReconcile_OutChecks_09a"))
'TotalOutstandingChecks2: =IIf([cmbAccounts]="CRTC01",DLookUp("CalcTotal","qryCheckReconcile_OutChecks_09b"),DLookUp("CalcTotal","qryCheckReconcile_OutChecks_09a"))
'DateStart: =#01/01/1900#

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

' ** Combo box column constants: cmbAccounts.
Private Const CBX_ACT_ACTNO   As Integer = 0
Private Const CBX_ACT_DESC    As Integer = 1
'Private Const CBX_ACT_STMDATE As Integer = 2
'Private Const CBX_ACT_SHORT   As Integer = 3
'Private Const CBX_ACT_LEGAL   As Integer = 4
'Private Const CBX_ACT_BALDATE As Integer = 5
'Private Const CBX_ACT_HASREL  As Integer = 6
'Private Const CBX_ACT_DEFAST  As Integer = 7

' ** Combo box column constants: cmbAssets.
Private Const CBX_AST_ASTNO   As Integer = 0
'Private Const CBX_AST_DESC    As Integer = 1
'Private Const CBX_AST_CUSIP   As Integer = 2
'Private Const CBX_AST_ATYPE   As Integer = 3

Private lngChks As Long, arr_varChk() As Variant
Private Const C_ELEMS As Integer = 12  ' ** Array's first-element UBound().
Private Const C_ACTID  As Integer = 0
Private Const C_STGID  As Integer = 1
Private Const C_JNO    As Integer = 2
Private Const C_ACTNO  As Integer = 3
Private Const C_CHKID  As Integer = 4
Private Const C_CNUM   As Integer = 5
Private Const C_PAID   As Integer = 6
Private Const C_DESC   As Integer = 7
Private Const C_HDESC  As Integer = 8
Private Const C_ASGN   As Integer = 9
Private Const C_ASTNO1 As Integer = 10
Private Const C_ASTNO2 As Integer = 11
Private Const C_FND    As Integer = 12

Private Const JRNL_PAID As String = "There are unposted PAID entries in the Journal."
Private Const JRNL_PAID_ACCT As String = "There are unposted PAID entries in the Journal for this account."

Private Const ASSIGN_MSG As String = "Associated with CUSIP:"

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnAddAsset_Focus As Boolean, blnAddAsset_MouseDown As Boolean
Private blnPreviewReport_Focus As Boolean, blnPreviewReport_MouseDown As Boolean
Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean
Private blnCheckAll_Focus As Boolean, blnCheckAll_MouseDown As Boolean
Private blnCheckNone_Focus As Boolean, blnCheckNone_MouseDown As Boolean

' ** Array: arr_varAsset_ThisAcct().
Private lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant
Private Const A_CRACTID As Integer = 0
'Private Const A_ACTNO   As Integer = 1
Private Const A_ASTNO   As Integer = 2
Private Const A_CUSIP   As Integer = 3

Private blnGoingToReport As Boolean
Private lngClear_Left As Long, lngTpp As Long
Private blnNotTillDateEntered As Boolean, blnFromTimer As Boolean
Private lngAssetNo_Move As Long
Private blnUnpostedPaid As Boolean, blnGetTotCashAcct As Boolean, blnSetDateEndFocus As Boolean
Private lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim strFilter As String
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

              ' ** Don't turn this on!
150           blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170           .OutChksBox_Left = .OutChecks_box.Left
180           lngClear_Left = .cmdClear.Left

190           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
200           DoEvents

210           gstrFormQuerySpec = THIS_NAME  ' ** This global var is used by Public FormRef().
220           gdatStartDate = 0: gdatEndDate = 0
230           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
240           blnAddAsset_Focus = False: blnAddAsset_MouseDown = False
250           blnPreviewReport_Focus = False: blnPreviewReport_MouseDown = False
260           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
270           blnCheckAll_Focus = False: blnCheckAll_MouseDown = False
280           blnCheckNone_Focus = False: blnCheckNone_MouseDown = False
290           lngAssetNo_Move = -1&: lngAssets_ThisAcct = 0&
300           blnGetTotCashAcct = False: blnSetDateEndFocus = False
310           blnNotTillDateEntered = False: blnFromTimer = False

320           CLR_DISABLED_FG = CLR_DKGRY
330           CLR_DISABLED_BG = CLR_LTTEAL

340           .cmdUpdate.Enabled = False
350           Select Case gblnSeparateCheckingAccounts
              Case True
360             .chkSeparateCheckingAccounts_lbl.Visible = False
370             .Bank_Name.Visible = True
380             .Bank_AccountNumber.Visible = True
390           Case False
400             .chkSeparateCheckingAccounts_lbl.Visible = True
410             .Bank_Name.Visible = False
420             .Bank_AccountNumber.Visible = False
430           End Select

              ' ** EVENT CHECK: chkRememberMe!
440           cmbAccounts_AfterUpdate  ' ** Procedure: Below.

450           Select Case .opgCheckingType
              Case .opgCheckingType_optIndividual.OptionValue
460             strFilter = "[accountno] = '0'"
470           Case .opgCheckingType_optTotalCash.OptionValue
480             strFilter = "[accountno] = 'CRTC01'"
490           End Select

              ' ** Here's where it jumps to Form_Load().
500           .Filter = strFilter
510           .FilterOn = True

              ' ** Back to here when Form_Load() finishes.
520           opgAccountSource_AfterUpdate  ' ** Procedure: Below.
530           opgAssetSource_AfterUpdate  ' ** Procedure: Below.
540           chkShowPaid_AfterUpdate  ' ** Procedure: Below.

550           CheckingType_Set lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

560           chkRememberMe_AfterUpdate  ' ** Procedure: Below.

              ' ** Then straight through to here.
570         Else
580           Cancel = -1
590         End If
600       Else
610         Cancel = -1
620       End If
630     End With

640     If Cancel = -1 Then
650       Beep
660       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
670     End If

EXITP:
680     Exit Sub

ERRH:
690     DoCmd.Hourglass False
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub Form_Load()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strMsg As String, strAccountNo As String

710     With Me

720       SetOption "Confirm Record Changes", True

          ' ** Create an instance of the Calendar class.
730       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
740       clsMonthClass.hWndForm = Me.hwnd

750       gstrFormQuerySpec = THIS_NAME
760       DoEvents

770       AccountsRowSource_Set Me  ' ** Module Procedure: modCheckReconcile.

780       If IsNull(.cmbAccounts) = True And .opgCheckingType = .opgCheckingType_optTotalCash.OptionValue Then
790         .cmbAccounts = "CRTC01"
800         DoEvents
810       ElseIf IsNull(.cmbAccounts) = False And .opgCheckingType = .opgCheckingType_optIndividual.OptionValue Then
820         If .cmbAccounts = "CRTC01" Then
830           .cmbAccounts = Null
840         End If
850       End If

860       strMsg = vbNullString: blnUnpostedPaid = False
870       Set dbs = CurrentDb
880       Select Case IsNull(.cmbAccounts)
          Case True

890         With dbs
              ' ** qryCheckReconcile_11a (Journal, just 'Paid', grouped by
              ' ** journaltype, PrintCheck, with cnt), grouped and summed.  '## OK
900           Set qdf = .QueryDefs("qryCheckReconcile_11b")
910           Set rst = qdf.OpenRecordset
920           With rst
930             If .BOF = True And .EOF = True Then
                  ' ** No Paid's in the Journal.
940             Else
950               .MoveFirst
960               If IsNull(![cnt]) = False Then
970                 If ![cnt] > 0 Then
980                   strMsg = JRNL_PAID
990                 End If
1000              End If
1010            End If
1020            .Close
1030          End With
1040        End With
            ' ** For Individual, jumps out of Select.

1050      Case False

1060        AccountsRowSource_Set Me  ' ** Module Procedure: modCheckReconcile.
1070        DoEvents

1080        strAccountNo = .cmbAccounts.Column(CBX_ACT_ACTNO)
1090        Select Case .opgCheckingType
            Case .opgCheckingType_optIndividual.OptionValue
1100          With dbs
1110            Select Case strAccountNo
                Case "CRTC01"
                  ' ** qryCheckReconcile_11e (Journal, just 'Paid', grouped by journaltype,
                  ' ** PrintCheck, with cnt, for accountno = 'CRTC01'), grouped and summed.
1120              Set qdf = .QueryDefs("qryCheckReconcile_11f")
1130            Case Else
                  ' ** qryCheckReconcile_11c (Journal, just 'Paid', grouped by journaltype,
                  ' ** PrintCheck, with cnt, by specified [actno]), grouped and summed.
1140              Set qdf = .QueryDefs("qryCheckReconcile_11d")
1150              With qdf.Parameters
1160                ![actno] = strAccountNo
1170              End With
1180            End Select
1190            Set rst = qdf.OpenRecordset
1200            With rst
1210              If .BOF = True And .EOF = True Then
                    ' ** No Paid's in the Journal for this account.
1220              Else
1230                .MoveFirst
1240                If IsNull(![cnt]) = False Then
1250                  If ![cnt] > 0 Then
1260                    strMsg = JRNL_PAID_ACCT
1270                  End If
1280                End If
1290              End If
1300              .Close
1310            End With  ' ** rst.
1320          End With  ' ** dbs.
1330        Case .opgCheckingType_optTotalCash.OptionValue
1340          With dbs
                ' ** qryCheckReconcile_15c (Journal, just 'Paid', grouped by journaltype,
                ' ** PrintCheck, with cnt), grouped and summed.  '## OK
1350            Set qdf = .QueryDefs("qryCheckReconcile_16d")
1360            Set rst = qdf.OpenRecordset
1370            With rst
1380              If .BOF = True And .EOF = True Then
                    ' ** No Paid's in the Journal whatsoever.
1390              Else
1400                .MoveFirst
1410                If IsNull(![cnt]) = False Then
1420                  If ![cnt] > 0 Then
1430                    strMsg = JRNL_PAID_ACCT
1440                  End If
1450                End If
1460              End If
1470              .Close
1480            End With  ' ** rst.
1490          End With  ' ** dbs.
1500          .cmbAssets = .assetno
1510          .DateEnd = .cracct_date
1520          SetSubForms True, True  ' ** Module Procedure: modCheckReconcile.
1530          blnGetTotCashAcct = True
              '.TimerInterval = 500&
1540        End Select

1550      End Select
1560      dbs.Close

1570      Select Case .opgCheckingType
          Case .opgCheckingType_optIndividual.OptionValue
1580        If .frmCheckReconcile_Sub_BSCredits.Form.RecordSource <> "qryCheckReconcile_BSCredit_01a" Then
1590          .frmCheckReconcile_Sub_BSCredits.Form.RecordSource = "qryCheckReconcile_BSCredit_01a"
1600        End If
1610        If .frmCheckReconcile_Sub_BSDebits.Form.RecordSource <> "qryCheckReconcile_BSDebit_01a" Then
1620          .frmCheckReconcile_Sub_BSDebits.Form.RecordSource = "qryCheckReconcile_BSDebit_01a"
1630        End If
1640        If .frmCheckReconcile_Sub_TACredits.Form.RecordSource <> "qryCheckReconcile_TACredit_01a" Then
1650          .frmCheckReconcile_Sub_TACredits.Form.RecordSource = "qryCheckReconcile_TACredit_01a"
1660        End If
1670        If .frmCheckReconcile_Sub_TADebits.Form.RecordSource <> "qryCheckReconcile_TADebit_01a" Then
1680          .frmCheckReconcile_Sub_TADebits.Form.RecordSource = "qryCheckReconcile_TADebit_01a"
1690        End If
1700      Case .opgCheckingType_optTotalCash.OptionValue
1710        If .frmCheckReconcile_Sub_BSCredits.Form.RecordSource <> "qryCheckReconcile_BSCredit_01b" Then
1720          .frmCheckReconcile_Sub_BSCredits.Form.RecordSource = "qryCheckReconcile_BSCredit_01b"
1730        End If
1740        If .frmCheckReconcile_Sub_BSDebits.Form.RecordSource <> "qryCheckReconcile_BSDebit_01b" Then
1750          .frmCheckReconcile_Sub_BSDebits.Form.RecordSource = "qryCheckReconcile_BSDebit_01b"
1760        End If
1770        If .frmCheckReconcile_Sub_TACredits.Form.RecordSource <> "qryCheckReconcile_TACredit_01b" Then
1780          .frmCheckReconcile_Sub_TACredits.Form.RecordSource = "qryCheckReconcile_TACredit_01b"
1790        End If
1800        If .frmCheckReconcile_Sub_TADebits.Form.RecordSource <> "qryCheckReconcile_TADebit_01b" Then
1810          .frmCheckReconcile_Sub_TADebits.Form.RecordSource = "qryCheckReconcile_TADebit_01b"
1820        End If
1830      End Select

1840      If strMsg <> vbNullString Then
1850        blnUnpostedPaid = True
1860        .UnpostedJournalMsg_lbl.Caption = strMsg
1870        .UnpostedJournalMsg_lbl.Visible = True
1880        .cmdClear_box.Visible = False
1890        .cmdClear_lbl1.Visible = False
1900        .cmdClear_lbl1_dim_hi.Visible = False
1910        .cmdClear_lbl2.Visible = False
1920        .cmdClear_lbl2_dim_hi.Visible = False
1930        .cmdClear_vline01.Visible = False
1940        .cmdClear_vline02.Visible = False
1950        .cmdDelete_lbl1.Visible = False
1960        .cmdDelete_lbl1_dim_hi.Visible = False
1970        .cmdDelete_lbl2.Visible = False
1980        .cmdDelete_lbl2_dim_hi.Visible = False
1990        .cmdClear.Left = ((.cmdDelete.Left - .cmdClear.Width) - (16& * lngTpp))
2000      End If

2010      DoCmd.Hourglass False

2020    End With

EXITP:
2030    Set rst = Nothing
2040    Set qdf = Nothing
2050    Set dbs = Nothing
2060    Exit Sub

ERRH:
100     DoCmd.Hourglass False
110     Select Case ERR.Number
        Case Else
120       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
130     End Select
140     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strAccountNo As String
        Dim intRetVal As Integer

2110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2120    intShiftDown = (Shift And acShiftMask) > 0
2130    intAltDown = (Shift And acAltMask) > 0
2140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F keys to other forms and functionality:
        ' **   Recalc Tots:      F5 {RecalcTots_CR}
        ' **   Date Picker:      F7 {cmdCalendar}

        ' ** Shortcut Ctrl to other forms and functionality.
        ' **   2nd Asset:        A {cmdAddAsset}
        ' **   Delete:           D {cmdDelete}
        ' **   Next Asset:       N {cmdAssetNext}
        ' **   Previous Asset:   P {cmdAssetPrevious}
        ' **   Clear:            R {cmdClear}
        ' **   Save:             S {cmdSave}

        ' ** Shortcut Alt-Shift keys responsive from this form:
        ' **   Show Paid:        H {chkShowPaid}
        ' **   By Cusip:         I {opgAssetSource_optCusip}
        ' **   By Name:          N {opgAssetSource_optName}
        ' **   By Type:          Y {opgAssetSource_optType}

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal  ' ** No vbKeyEscape
          Case vbKeyF5
2170        With Me
2180          DoCmd.Hourglass True
2190          DoEvents
2200          strAccountNo = .cmbAccounts.Column(CBX_ACT_ACTNO)
2210          If lngAssets_ThisAcct = 0& Or IsEmpty(arr_varAsset_ThisAcct) = True Then
2220            arr_varAsset_ThisAcct = AssetsThisArray_Get  ' ** Function: Below.
2230            lngAssets_ThisAcct = UBound(arr_varAsset_ThisAcct, 2) + 1&
2240          End If
2250          RecalcTots_CR strAccountNo, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.
2260          DoCmd.Hourglass False
2270        End With
2280      Case vbKeyF7
2290        With Me
2300          intRetVal = 0
2310          If .cmdCalendar.Enabled = True Then
2320            .cmdCalendar.SetFocus
2330            cmdCalendar_Click  ' ** Procedure: Below.
2340          Else
2350            Beep
2360          End If
2370        End With
2380      End Select
2390    End If

        ' ** Alt keys.
2400    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2410      Select Case intRetVal
          Case vbKeyE
2420        With Me
2430          intRetVal = 0
2440          If .frmCheckReconcile_Sub_BSDebits.Enabled = True Then
2450            lngRecsCur = .frmCheckReconcile_Sub_BSDebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSDebits.
2460            .frmCheckReconcile_Sub_BSDebits.SetFocus
2470            If lngRecsCur > 0& Then
2480              .frmCheckReconcile_Sub_BSDebits.Form.FocusHolder.SetFocus
2490            Else
2500              .frmCheckReconcile_Sub_BSDebits.Form.critem_description.SetFocus
2510            End If
2520          Else
2530            Beep
2540          End If
2550        End With
2560      Case vbKeyI
2570        With Me
2580          intRetVal = 0
2590          If .frmCheckReconcile_Sub_TACredits.Enabled = True Then
2600            lngRecsCur = .frmCheckReconcile_Sub_TACredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TACredits.
2610            .frmCheckReconcile_Sub_TACredits.SetFocus
2620            If lngRecsCur > 0& Then
2630              .frmCheckReconcile_Sub_TACredits.Form.FocusHolder.SetFocus
2640            Else
2650              .frmCheckReconcile_Sub_TACredits.Form.critem_description.SetFocus
2660            End If
2670          Else
2680            Beep
2690          End If
2700        End With
2710      Case vbKeyO
2720        With Me
2730          intRetVal = 0
2740          If .frmCheckReconcile_Sub_TADebits.Enabled = True Then
2750            lngRecsCur = .frmCheckReconcile_Sub_TADebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TADebits.
2760            .frmCheckReconcile_Sub_TADebits.SetFocus
2770            If lngRecsCur > 0& Then
2780              .frmCheckReconcile_Sub_TADebits.Form.FocusHolder.SetFocus
2790            Else
2800              .frmCheckReconcile_Sub_TADebits.Form.critem_description.SetFocus
2810            End If
2820          Else
2830            Beep
2840          End If
2850        End With
2860      Case vbKeyR
2870        With Me
2880          intRetVal = 0
2890          If .frmCheckReconcile_Sub_BSCredits.Enabled = True Then
2900            lngRecsCur = .frmCheckReconcile_Sub_BSCredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSCredits.
2910            .frmCheckReconcile_Sub_BSCredits.SetFocus
2920            If lngRecsCur > 0& Then
2930              .frmCheckReconcile_Sub_BSCredits.Form.FocusHolder.SetFocus
2940            Else
2950              .frmCheckReconcile_Sub_BSCredits.Form.critem_description.SetFocus
2960            End If
2970          Else
2980            Beep
2990          End If
3000        End With
3010      Case vbKeyX
3020        intRetVal = 0
3030        cmdClose_Click  ' ** Procedure: Below.
3040      End Select
3050    End If

        ' ** Ctrl keys.
3060    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3070      Select Case intRetVal
          Case vbKeyA
3080        With Me
3090          intRetVal = 0
3100          If .cmdAddAsset.Visible = True And .cmdAddAsset.Enabled = True Then
3110            .cmdAddAsset.SetFocus
3120            cmdAddAsset_Click  ' ** Procedure: Below.
3130          Else
3140            Beep
3150          End If
3160        End With
3170      Case vbKeyD
3180        With Me
3190          intRetVal = 0
3200          If .cmdDelete.Visible = True And .cmdDelete.Enabled = True Then
3210            .cmdDelete.SetFocus
3220            cmdDelete_Click  ' ** Procedure: Below.
3230          Else
3240            Beep
3250          End If
3260        End With
3270      Case vbKeyN
3280        With Me
3290          intRetVal = 0
3300          If .cmdAssetNext.Visible = True And .cmdAssetNext.Enabled = True Then
3310            .cmdAssetNext.SetFocus
3320            cmdAssetNext_Click  ' ** Procedure: Below.
3330          End If
3340        End With
3350      Case vbKeyP
3360        With Me
3370          intRetVal = 0
3380          If .cmdAssetPrevious.Visible = True And .cmdAssetPrevious.Enabled = True Then
3390            .cmdAssetPrevious.SetFocus
3400            cmdAssetPrevious_Click  ' ** Procedure: Below.
3410          End If
3420        End With
3430      Case vbKeyR
3440        With Me
3450          intRetVal = 0
3460          If .cmdClear.Visible = True And .cmdClear.Enabled = True Then
3470            .cmdClear.SetFocus
3480            cmdClear_Click  ' ** Procedure: Below.
3490          Else
3500            Beep
3510          End If
3520        End With
3530      Case vbKeyS
3540        With Me
3550          intRetVal = 0
3560          cmdSave_Click  ' ** Procedure: Below.
3570          lngRecsCur = .frmCheckReconcile_Sub_BSCredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSCredits.
3580          If lngRecsCur > 0& Then
3590            .frmCheckReconcile_Sub_BSCredits.Form.cmdSave_Click  ' ** Form Procedure: frmCheckReconcile_Sub_BSCredits.
3600          End If
3610          lngRecsCur = .frmCheckReconcile_Sub_BSDebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSDebits.
3620          If lngRecsCur > 0& Then
3630            .frmCheckReconcile_Sub_BSDebits.Form.cmdSave_Click  ' ** Form Procedure: frmCheckReconcile_Sub_BSDebits.
3640          End If
3650          lngRecsCur = .frmCheckReconcile_Sub_TACredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TACredits.
3660          If lngRecsCur > 0& Then
3670            .frmCheckReconcile_Sub_TACredits.Form.cmdSave_Click  ' ** Form Procedure: frmCheckReconcile_Sub_TACredits.
3680          End If
3690          lngRecsCur = .frmCheckReconcile_Sub_TADebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TADebits.
3700          If lngRecsCur > 0& Then
3710            .frmCheckReconcile_Sub_TADebits.Form.cmdSave_Click  ' ** Form Procedure: frmCheckReconcile_Sub_TADebits.
3720          End If
3730          lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
3740          If lngRecsCur > 0& Then
3750            .frmCheckReconcile_Sub_OutChecks.Form.cmdSave_Click  ' ** Form Procedure: frmCheckReconcile_Sub_OutChecks.
3760          End If
3770        End With
3780      End Select
3790    End If

        ' ** Alt-Shift keys.
3800    If (Not intCtrlDown) And intAltDown And intShiftDown Then
3810      Select Case intRetVal
          Case vbKeyH
3820        With Me
3830          intRetVal = 0
3840          If .chkShowPaid.Enabled = True Then
3850            .chkShowPaid.SetFocus
3860            .chkShowPaid = (Not .chkShowPaid)
3870            chkShowPaid_AfterUpdate  ' ** Procedure: Below.
3880          End If
3890        End With
3900      Case vbKeyI
3910        With Me
3920          intRetVal = 0
3930          If .opgAssetSource.Enabled = True Then
3940            .opgAssetSource.SetFocus
3950            .opgAssetSource = .opgAssetSource_optCusip.OptionValue
3960            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
3970          End If
3980        End With
3990      Case vbKeyN
4000        With Me
4010          intRetVal = 0
4020          If .opgAssetSource.Enabled = True Then
4030            .opgAssetSource.SetFocus
4040            .opgAssetSource = .opgAssetSource_optName.OptionValue
4050            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
4060          End If
4070        End With
4080      Case vbKeyY
4090        With Me
4100          intRetVal = 0
4110          If .opgAssetSource.Enabled = True Then
4120            .opgAssetSource.SetFocus
4130            .opgAssetSource = .opgAssetSource_optType.OptionValue
4140            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
4150          End If
4160        End With
4170      End Select
4180    End If

        ' ** Ctrl-Shift keys.
4190    If intCtrlDown And (Not intAltDown) And intShiftDown Then
4200      Select Case intRetVal
          Case vbKeyF
4210        With Me
4220          intRetVal = 0
4230          .FocusHolder.SetFocus
4240        End With
4250      End Select
4260    End If

EXITP:
4270    KeyCode = intRetVal
4280    Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

4310    With Me
4320      .cracct_datemodified = Now()
4330    End With

EXITP:
4340    Exit Sub

ERRH:
690     DoCmd.Hourglass False
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub Form_Timer()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

        Static blnUnposted As Boolean

4410    With Me
4420      .TimerInterval = 0&
4430      If gblnSetFocus = True Then
4440        gblnSetFocus = False
4450        strFormName = vbNullString
4460  On Error Resume Next
4470        Do Until Screen.ActiveForm.Name = strFormName
4480          DoCmd.SelectObject acForm, THIS_NAME, False
4490          strFormName = Screen.ActiveForm.Name
4500        Loop
4510  On Error GoTo ERRH
4520        gstrFormQuerySpec = THIS_NAME
4530      End If
4540      If blnGetTotCashAcct = True Then
4550        blnGetTotCashAcct = False
4560        If .cmbAccounts = "CRTC01" Then
4570          If InStr(.Filter, "CRTC01") = 0 Then
4580            blnFromTimer = True
4590            .Filter = "[accountno] = 'CRTC01'"
4600            .FilterOn = True
4610            DoEvents
4620            blnFromTimer = False
4630          End If
4640          AccountsRowSource_Set Me  ' ** Module Procedure: modCheckReconcile.
4650          If .accountno = "CRTC01" Then
4660            cmbAccounts_AfterUpdate  ' ** Procedure: Below.
4670            If .DateEnd <> .cracct_date Then
4680              Select Case IsNull(.cracct_date)
                  Case True
4690                If IsDate(.DateEnd) = False Then  ' ** I can't get it to stick!
4700                  .DateEnd = Date
4710                End If
4720              Case False
4730                .DateEnd = .cracct_date
4740              End Select
4750            End If
4760            If IsNull(.cmbAssets) = True Then
4770              .cmbAssets = .assetno
4780            End If
4790            .DateEnd.SetFocus
4800            blnGetTotCashAcct = False
4810            blnSetDateEndFocus = True  ' ** I don't why all these things won't work on load!!
4820            .TimerInterval = 500&
4830          End If
4840        End If

4850        If .frmCheckReconcile_Sub_BSCredits.Form.RecordSource <> "qryCheckReconcile_BSCredit_01b" Then
4860          .frmCheckReconcile_Sub_BSCredits.Form.RecordSource = "qryCheckReconcile_BSCredit_01b"
4870        End If
4880        If .frmCheckReconcile_Sub_BSDebits.Form.RecordSource <> "qryCheckReconcile_BSDebit_01b" Then
4890          .frmCheckReconcile_Sub_BSDebits.Form.RecordSource = "qryCheckReconcile_BSDebit_01b"
4900        End If
4910        If .frmCheckReconcile_Sub_TACredits.Form.RecordSource <> "qryCheckReconcile_TACredit_01b" Then
4920          .frmCheckReconcile_Sub_TACredits.Form.RecordSource = "qryCheckReconcile_TACredit_01b"
4930        End If
4940        If .frmCheckReconcile_Sub_TADebits.Form.RecordSource <> "qryCheckReconcile_TADebit_01b" Then
4950          .frmCheckReconcile_Sub_TADebits.Form.RecordSource = "qryCheckReconcile_TADebit_01b"
4960        End If

4970      ElseIf blnSetDateEndFocus = True Then
4980        blnGetTotCashAcct = False
4990        If IsDate(.DateEnd) = False Then  ' ** I can't get it to stick!
5000          .DateEnd = Date
5010        End If
5020  On Error Resume Next
5030        .DateEnd.SelLength = 0
5040        .DateEnd.SelStart = 19
5050  On Error GoTo ERRH
5060      End If
5070      If gblnGoToReport = True Then
5080        If .UnpostedJournalMsg_lbl.Visible = True Then
5090          blnUnposted = True
5100        Else
5110          blnUnposted = False
5120        End If
5130        Select Case blnGoingToReport
            Case True
5140          DoCmd.Hourglass False  ' ** It should already be off.
5150          gblnGoToReport = False
5160          blnGoingToReport = False
5170          .GoToReport_arw_pr_img.Visible = False
5180          If blnUnposted = True Then
5190            .UnpostedJournalMsg_lbl.Visible = True
5200          End If
5210          If gblnMessage = True Then
5220            gblnMessage = False
5230            .cmdClose.SetFocus
5240            .cmdPrintReport.Enabled = False
5250            .cmdPrintReport_raised_img_dis.Visible = True
5260            .cmdPrintReport_raised_img.Visible = False
5270          End If
5280          DoEvents
5290        Case False
5300          DoCmd.Hourglass True  ' ** Make sure it's still running.
5310          DoEvents
5320          blnGoingToReport = True: gblnMessage = False
5330          If blnUnposted = True Then
5340            .UnpostedJournalMsg_lbl.Visible = False
5350          End If
5360          If .cmdPrintReport.Enabled = False Then  ' ** Just for this.
5370            .cmdPrintReport.Enabled = True
5380            gblnMessage = True
5390          End If
5400          .cmdPrintReport.SetFocus
5410          DoEvents
5420          .GoToReport_arw_pr_img.Visible = True
5430          Beep
5440          DoCmd.Hourglass False
5450          DoEvents
5460          .TimerInterval = (GTR_WAIT * 2&)
5470        End Select
5480      End If
5490    End With

EXITP:
5500    Exit Sub

ERRH:
690     DoCmd.Hourglass False
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

5610    If DataErr <> 0 Then
5620      Select Case DataErr
          Case 2501  ' ** The '|' action was Canceled.
5630        Response = acDataErrContinue
5640      Case 2113, 2279
            ' ** 2113: The value you entered isn't valid for this field.
            ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
5650        If Right(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateStart, DateEnd.
5660          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
5670          Response = acDataErrContinue
5680        End If
5690      Case 2237  ' ** The text you entered isn't an item in the list.
5700        Response = acDataErrDisplay
5710      Case Else
5720        zErrorHandler THIS_NAME, THIS_PROC, DataErr, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5730      End Select
5740    End If

5750    gstrFormQuerySpec = THIS_NAME

EXITP:
5760    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim msgResponse As VbMsgBoxResult
        Dim lngX As Long
        Dim blnRetVal As Boolean

5810    msgResponse = vbYes
5820    For lngX = 0& To (lngChks - 1&)
5830      If arr_varChk(C_CHKID, lngX) <> 0& Then
5840        Beep
5850        msgResponse = MsgBox("You have made changes to the Outstanding Checks, but have not saved them." & vbCrLf & vbCrLf & _
              "Are you sure you want to abandon changes?", _
              vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved")
5860        Exit For
5870      End If
5880    Next

5890    If msgResponse = vbYes Then

5900      With Me
5910        If .chkRememberMe = False Then
5920          .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
5930        End If
5940      End With

5950      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

5960      Set dbs = CurrentDb
            ' ** Empty tmpAccount.
5970        Set qdf = dbs.QueryDefs("qryCheckReconcile_03_01")
5980        qdf.Execute
5990      Set qdf = Nothing
6000      dbs.Close
6010      Set dbs = Nothing
6020      DoEvents

6030      SetOption "Confirm Record Changes", False  ' ** Application default setting.

6040      blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.

6050      gstrFormQuerySpec = vbNullString
6060      gdatStartDate = 0
6070      gdatEndDate = 0

6080      Select Case strCallingForm
          Case "frmMenu_Post"
6090        gblnSetFocus = True
6100        gstrReturningForm = THIS_NAME
6110        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
6120        DoCmd.Hourglass False
6130      End Select

6140    Else
6150      Cancel = -1
6160    End If

EXITP:
6170    Set qdf = Nothing
6180    Set dbs = Nothing
6190    Exit Sub

ERRH:
690     DoCmd.Hourglass False
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

6210    Detail_Mouse_CR blnAddAsset_Focus, blnCalendar1_Focus, blnPreviewReport_Focus, _
          blnPrintReport_Focus, blnCheckAll_Focus, blnCheckNone_Focus, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
6220    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

6310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6320    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

6410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6420    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Public Sub cmdClose_Click()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

6510    DoCmd.Hourglass True
6520    DoEvents

6530    DoCmd.Close acForm, THIS_NAME

EXITP:
6540    Exit Sub

ERRH:
690     DoCmd.Hourglass False
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
6670        With Me
6680          intRetVal = 0
6690          .opgCheckingType.SetFocus
6700        End With
6710      End Select
6720    End If

        ' ** Shift keys.
6730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6740      Select Case intRetVal
          Case vbKeyTab
6750        With Me
6760          intRetVal = 0
6770          If .cmdDelete.Enabled = True Then
6780            .cmdDelete.SetFocus
6790          ElseIf .cmdClear.Enabled = True Then
6800            .cmdClear.SetFocus
6810          ElseIf .cmdPrintReport.Enabled = True Then
6820            .cmdPrintReport.SetFocus
6830          ElseIf .cmdPreviewReport.Enabled = True Then
6840            .cmdPreviewReport.SetFocus
6850          ElseIf .chkShowPaid.Enabled = True Then
6860            .chkShowPaid.SetFocus
6870          Else
6880            .cmbAccounts.SetFocus
6890          End If
6900        End With
6910      End Select
6920    End If

EXITP:
6930    KeyCode = intRetVal
6940    Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub cmdSave_Click()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

7010    DoCmd.RunCommand acCmdSaveRecord
7020    DoEvents

EXITP:
7030    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPreviewReport_GotFocus()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_GotFocus"

7110    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
7120    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPreviewReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_MouseDown"

7210    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
7220    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPreviewReport_Click()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_Click"

7310    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
7320    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPreviewReport_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7420    intShiftDown = (Shift And acShiftMask) > 0
7430    intAltDown = (Shift And acAltMask) > 0
7440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7460      Select Case intRetVal
          Case vbKeyTab
7470        With Me
7480          intRetVal = 0
7490          .cmdPrintReport.SetFocus
7500        End With
7510      End Select
7520    End If

        ' ** Shift keys.
7530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7540      Select Case intRetVal
          Case vbKeyTab
7550        With Me
7560          intRetVal = 0
7570          If .chkShowPaid.Enabled = True Then
7580            .chkShowPaid.SetFocus
7590          ElseIf .cracct_bsbalance_display.Enabled = True Then
7600            .cracct_bsbalance_display.SetFocus
7610          ElseIf .DateEnd.Enabled = True Then
7620            .DateEnd.SetFocus
7630          Else
7640            .cmbAccounts.SetFocus
7650          End If
7660        End With
7670      End Select
7680    End If

EXITP:
7690    KeyCode = intRetVal
7700    Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub cmdPreviewReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_MouseMove"

7810    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
7820    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPreviewReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_MouseUp"

7910    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
7920    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPreviewReport_LostFocus()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_LostFocus"

8010    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
8020    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

8110    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
8120    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

8210    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
8220    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

8310    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
8320    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
8490          If .cmdClear.Enabled = True Then
8500            .cmdClear.SetFocus
8510          ElseIf .cmdDelete.Enabled = True Then
8520            .cmdDelete.SetFocus
8530          Else
8540            .cmdClose.SetFocus
8550          End If
8560        End With
8570      End Select
8580    End If

        ' ** Shift keys.
8590    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8600      Select Case intRetVal
          Case vbKeyTab
8610        With Me
8620          intRetVal = 0
8630          .cmdPreviewReport.SetFocus
8640        End With
8650      End Select
8660    End If

EXITP:
8670    KeyCode = intRetVal
8680    Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

8710    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
8720    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

8810    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
8820    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

8910    PrevPrint_Handler_CR THIS_PROC, blnPreviewReport_Focus, blnPreviewReport_MouseDown, _
          blnPrintReport_Focus, blnPrintReport_MouseDown, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
8920    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Public Sub cmdClear_Click()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClear_Click"

9010    With Me
9020      If .opgCheckingType = .opgCheckingType_optTotalCash.OptionValue Then
9030        .chkRememberMe = False
9040        chkRememberMe_AfterUpdate  ' ** Procedure: Below.
9050        .opgCheckingType = .opgCheckingType_optIndividual.OptionValue
9060        opgCheckingType_AfterUpdate  ' ** Procedure: Below.
9070      End If
9080      .cmbAccounts = Null
9090      lngAssets_ThisAcct = 0&
9100      arr_varAsset_ThisAcct = Empty
9110      cmbAccounts_AfterUpdate  ' ** Procedure: Below.
9120      DoEvents
9130      .cmbAccounts.SetFocus
9140    End With

EXITP:
9150    Exit Sub

ERRH:
690     Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
700     Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Public Sub cmdDelete_Click()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strAccountNo As String, lngAssetNo As Long, lngAssets As Long
        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String

9210    With Me
9220      strAccountNo = vbNullString: lngAssetNo = 0&: lngAssets = 0&
9230      strTmp01 = vbNullString
9240      If IsNull(.cmbAccounts) = False Then
9250        If Trim(.cmbAccounts) <> vbNullString Then
9260          strAccountNo = .cmbAccounts.Column(CBX_ACT_ACTNO)
9270          If IsNull(.cmbAssets) = False Then
9280            lngAssetNo = .cmbAssets
9290          End If
9300          lngAssets = Val(Mid(.cmbAssets_Cnt_lbl.Caption, (InStr(.cmbAssets_Cnt_lbl.Caption, "of") + 3)))
9310          If lngAssets > 1& Then strTmp01 = vbCrLf & "    " & .cmbAssets.Column(1)
9320          Beep
9330          msgResponse = MsgBox("Are you sure you want to delete" & vbCrLf & _
                "the Check Reconciliation record for:" & vbCrLf & vbCrLf & _
                "  " & .cmbAccounts.Column(CBX_ACT_DESC) & strTmp01, _
                vbQuestion + vbYesNo + vbDefaultButton2, "Delete Check Reconciliation Record")
9340          If msgResponse = vbYes Then
9350            DoCmd.Hourglass True
9360            DoEvents
9370            .opgCheckingType.SetFocus
9380            DoEvents
9390            Select Case .opgCheckingType
                Case .opgCheckingType_optIndividual.OptionValue
9400              cmdClear_Click  ' ** Procedure: Above.
9410              Set dbs = CurrentDb
9420              With dbs
                    ' ** Delete tblCheckReconcile_Account, by specified [actno], [astno].
9430                Set qdf = .QueryDefs("qryCheckReconcile_10a")
9440                With qdf.Parameters
9450                  ![actno] = strAccountNo
9460                  ![astno] = lngAssetNo
9470                End With
9480                qdf.Execute
9490                .Close
9500              End With
9510              DoEvents
9520            Case .opgCheckingType_optTotalCash.OptionValue
9530              Set dbs = CurrentDb
9540              With dbs
                    ' ** Empty tblCheckReconcile_Staging.  '## OK
9550                Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_20")
9560                qdf.Execute
                    ' ** Delete tblCheckReconcile_Account, for accountno = 'CRTC01'.
9570                Set qdf = .QueryDefs("qryCheckReconcile_10b")
9580                qdf.Execute
                    ' ** Update Account, just accountno = 'CRTC01', for Bank_Name = Null, Bank_AccountNumber = Null.
9590                Set qdf = .QueryDefs("qryCheckReconcile_18")
9600                qdf.Execute
9610                .Close
9620              End With
9630              DoEvents
9640              .Refresh
9650              .Requery
9660              .opgCheckingType = .opgCheckingType_optIndividual.OptionValue
9670              opgCheckingType_AfterUpdate  ' ** Procedure: Below.
9680            End Select
9690            .cmdDelete.Enabled = False
9700            .cmdDelete_lbl1.ForeColor = WIN_CLR_DISF
9710            .cmdDelete_lbl2.ForeColor = WIN_CLR_DISF
9720            If .UnpostedJournalMsg_lbl.Visible = False Then
9730              .cmdDelete_lbl1_dim_hi.Visible = True
9740              .cmdDelete_lbl2_dim_hi.Visible = True
9750            End If
9760            .cmdClear.Enabled = False
9770            .cmdClear_lbl1.ForeColor = WIN_CLR_DISF
9780            .cmdClear_lbl2.ForeColor = WIN_CLR_DISF
9790            If .UnpostedJournalMsg_lbl.Visible = False Then
9800              .cmdClear_lbl1_dim_hi.Visible = True
9810              .cmdClear_lbl2_dim_hi.Visible = True
9820            End If
9830            .cmdCheckAll.Enabled = False
9840            .cmdCheckAll_raised_img_dis.Visible = True
9850            .cmdCheckAll_raised_img.Visible = False
9860            .cmdCheckNone.Enabled = False
9870            .cmdCheckNone_raised_img_dis.Visible = True
9880            .cmdCheckNone_raised_img.Visible = False
9890            lngChks = 0&
9900            ReDim arr_varChk(C_ELEMS, 0)
9910            lngAssets_ThisAcct = 0&
9920            arr_varAsset_ThisAcct = Empty
9930          End If
9940        Else
9950          Beep
9960        End If
9970      Else
9980        Beep
9990      End If
10000   End With

10010   DoCmd.Hourglass False

EXITP:
10020   Set qdf = Nothing
10030   Set dbs = Nothing
10040   Exit Sub

ERRH:
690     DoCmd.Hourglass False
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_KeyDown"

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
          Case vbKeyTab
10170       With Me
10180         intRetVal = 0
10190         .cmdClose.SetFocus
10200       End With
10210     End Select
10220   End If

        ' ** Shift keys.
10230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10240     Select Case intRetVal
          Case vbKeyTab
10250       With Me
10260         intRetVal = 0
10270         .cmdClear.SetFocus
10280       End With
10290     End Select
10300   End If

EXITP:
10310   KeyCode = intRetVal
10320   Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub opgCheckingType_AfterUpdate()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCheckingType_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strMsg As String, strAccountNo As String
        Dim varTmp00 As Variant

10410   With Me
10420     Select Case .opgCheckingType
          Case .opgCheckingType_optIndividual.OptionValue
10430       .opgCheckingType_optIndividual_lbl.FontBold = True
10440       .opgCheckingType_optTotalCash_lbl.FontBold = False
10450       .cmbAccounts = Null
10460       .chkShowPaid = False
10470       If InStr(.Filter, "CRTC01") > 0 Then
10480         .Filter = "[accountno] = '0'"
10490         .FilterOn = True
10500         DoEvents
10510       End If
10520       CheckingType_Set lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.
10530       chkShowPaid_AfterUpdate  ' ** Procedure: Below.
10540       DoEvents
10550       strMsg = vbNullString: blnUnpostedPaid = False
10560       Set dbs = CurrentDb
10570       With dbs
              ' ** qryCheckReconcile_11a (Journal, just 'Paid', grouped by
              ' ** journaltype, PrintCheck, with cnt), grouped and summed.  '## OK
10580         Set qdf = .QueryDefs("qryCheckReconcile_11b")
10590         Set rst = qdf.OpenRecordset
10600         With rst
10610           If .BOF = True And .EOF = True Then
                  ' ** No Paid's in the Journal.
10620           Else
10630             .MoveFirst
10640             If IsNull(![cnt]) = False Then
10650               If ![cnt] > 0 Then
10660                 strMsg = JRNL_PAID
10670               End If
10680             End If
10690           End If
10700           .Close
10710         End With  ' ** rst.
10720         .Close
10730       End With  ' ** dbs.
10740       If strMsg <> vbNullString Then
10750         blnUnpostedPaid = True
10760         .UnpostedJournalMsg_lbl.Caption = strMsg
10770         .UnpostedJournalMsg_lbl.Visible = True
10780         .cmdClear_box.Visible = False
10790         .cmdClear_lbl1.Visible = False
10800         .cmdClear_lbl1_dim_hi.Visible = False
10810         .cmdClear_lbl2.Visible = False
10820         .cmdClear_lbl2_dim_hi.Visible = False
10830         .cmdClear_vline01.Visible = False
10840         .cmdClear_vline02.Visible = False
10850         .cmdDelete_lbl1.Visible = False
10860         .cmdDelete_lbl1_dim_hi.Visible = False
10870         .cmdDelete_lbl2.Visible = False
10880         .cmdDelete_lbl2_dim_hi.Visible = False
10890         .cmdClear.Left = ((.cmdDelete.Left - .cmdClear.Width) - (16& * lngTpp))
10900       End If
10910       If .frmCheckReconcile_Sub_BSCredits.Form.RecordSource <> "qryCheckReconcile_BSCredit_01a" Then
10920         .frmCheckReconcile_Sub_BSCredits.Form.RecordSource = "qryCheckReconcile_BSCredit_01a"
10930       End If
10940       If .frmCheckReconcile_Sub_BSDebits.Form.RecordSource <> "qryCheckReconcile_BSDebit_01a" Then
10950         .frmCheckReconcile_Sub_BSDebits.Form.RecordSource = "qryCheckReconcile_BSDebit_01a"
10960       End If
10970       If .frmCheckReconcile_Sub_TACredits.Form.RecordSource <> "qryCheckReconcile_TACredit_01a" Then
10980         .frmCheckReconcile_Sub_TACredits.Form.RecordSource = "qryCheckReconcile_TACredit_01a"
10990       End If
11000       If .frmCheckReconcile_Sub_TADebits.Form.RecordSource <> "qryCheckReconcile_TADebit_01a" Then
11010         .frmCheckReconcile_Sub_TADebits.Form.RecordSource = "qryCheckReconcile_TADebit_01a"
11020       End If
11030       CheckingType_Set lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.
11040       .cmbAccounts.SetFocus
11050     Case .opgCheckingType_optTotalCash.OptionValue
11060       .opgCheckingType_optTotalCash_lbl.FontBold = True
11070       .opgCheckingType_optIndividual_lbl.FontBold = False
11080       .cmbAccounts = Null
11090       CheckingType_Set lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.
11100       Set dbs = CurrentDb
11110       varTmp00 = DLookup("[shortname]", "account", "[accountno] = 'CRTC01'")
11120       If IsNull(varTmp00) = True Then
              ' ** Append qryCheckReconcile_12 (Check Reconcile Total Cash, as Account entry) to Account.  '## OK
11130         Set qdf = dbs.QueryDefs("qryCheckReconcile_13")
11140         qdf.Execute
11150         Set qdf = Nothing
              ' ** Append accountno = 'CRTC01' to Balance.  '## OK
11160         Set qdf = dbs.QueryDefs("qryCheckReconcile_14")
11170         qdf.Execute
11180         Set qdf = Nothing
11190       End If
            ' ** Append accountno = 'CRTC01' to tblCheckReconcile_Account.  '## OK
11200       Set qdf = dbs.QueryDefs("qryCheckReconcile_15")
11210       qdf.Execute
11220       Set qdf = Nothing
11230       DoEvents
11240       .cmbAccounts.Requery
11250       .cmbAccounts = "CRTC01"
11260       DoEvents
11270       strMsg = vbNullString: blnUnpostedPaid = False
11280       strAccountNo = .cmbAccounts.Column(CBX_ACT_ACTNO)
11290       With dbs
              ' ** qryCheckReconcile_15c (Journal, just 'Paid', grouped by journaltype,
              ' ** PrintCheck, with cnt), grouped and summed.  '## OK
11300         Set qdf = .QueryDefs("qryCheckReconcile_16d")
11310         Set rst = qdf.OpenRecordset
11320         With rst
11330           If .BOF = True And .EOF = True Then
                  ' ** No Paid's in the Journal whatsoever.
11340           Else
11350             .MoveFirst
11360             If IsNull(![cnt]) = False Then
11370               If ![cnt] > 0 Then
11380                 strMsg = JRNL_PAID_ACCT
11390               End If
11400             End If
11410           End If
11420           .Close
11430         End With  ' ** rst.
11440         .Close
11450       End With  ' ** dbs.
11460       SetSubForms False, False  ' ** Module Procedure: modCheckReconcile.
11470       If strMsg <> vbNullString Then
11480         blnUnpostedPaid = True
11490         .UnpostedJournalMsg_lbl.Caption = strMsg
11500         .UnpostedJournalMsg_lbl.Visible = True
11510         .cmdClear_box.Visible = False
11520         .cmdClear_lbl1.Visible = False
11530         .cmdClear_lbl1_dim_hi.Visible = False
11540         .cmdClear_lbl2.Visible = False
11550         .cmdClear_lbl2_dim_hi.Visible = False
11560         .cmdClear_vline01.Visible = False
11570         .cmdClear_vline02.Visible = False
11580         .cmdDelete_lbl1.Visible = False
11590         .cmdDelete_lbl1_dim_hi.Visible = False
11600         .cmdDelete_lbl2.Visible = False
11610         .cmdDelete_lbl2_dim_hi.Visible = False
11620         .cmdClear.Left = ((.cmdDelete.Left - .cmdClear.Width) - (16& * lngTpp))
11630       End If
11640       blnGetTotCashAcct = True
11650       If .frmCheckReconcile_Sub_BSCredits.Form.RecordSource <> "qryCheckReconcile_BSCredit_01b" Then
11660         .frmCheckReconcile_Sub_BSCredits.Form.RecordSource = "qryCheckReconcile_BSCredit_01b"
11670       End If
11680       If .frmCheckReconcile_Sub_BSDebits.Form.RecordSource <> "qryCheckReconcile_BSDebit_01b" Then
11690         .frmCheckReconcile_Sub_BSDebits.Form.RecordSource = "qryCheckReconcile_BSDebit_01b"
11700       End If
11710       If .frmCheckReconcile_Sub_TACredits.Form.RecordSource <> "qryCheckReconcile_TACredit_01b" Then
11720         .frmCheckReconcile_Sub_TACredits.Form.RecordSource = "qryCheckReconcile_TACredit_01b"
11730       End If
11740       If .frmCheckReconcile_Sub_TADebits.Form.RecordSource <> "qryCheckReconcile_TADebit_01b" Then
11750         .frmCheckReconcile_Sub_TADebits.Form.RecordSource = "qryCheckReconcile_TADebit_01b"
11760       End If
11770       .TimerInterval = 500&
11780     End Select
11790   End With

EXITP:
11800   Set rst = Nothing
11810   Set qdf = Nothing
11820   Set dbs = Nothing
11830   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
700     Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub opgCheckingType_optIndividual_KeyDown(KeyCode As Integer, Shift As Integer)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCheckingType_optIndividual_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11920   intShiftDown = (Shift And acShiftMask) > 0
11930   intAltDown = (Shift And acAltMask) > 0
11940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11970       With Me
11980         intRetVal = 0
11990         If .cmbAccounts.Enabled = True Then
12000           .cmbAccounts.SetFocus
12010         ElseIf .DateEnd.Enabled = True Then
12020           .DateEnd.SetFocus
12030         Else
12040           .cmdClose.SetFocus
12050         End If
12060       End With
12070     End Select
12080   End If

        ' ** Shift keys.
12090   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12110       With Me
12120         intRetVal = 0
12130         .cmdClose.SetFocus
12140       End With
12150     End Select
12160   End If

EXITP:
12170   KeyCode = intRetVal
12180   Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub opgCheckingType_optTotalCash_KeyDown(KeyCode As Integer, Shift As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCheckingType_optTotalCash_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12220   intShiftDown = (Shift And acShiftMask) > 0
12230   intAltDown = (Shift And acAltMask) > 0
12240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12270       With Me
12280         intRetVal = 0
12290         If .cmbAccounts.Enabled = True Then
12300           .cmbAccounts.SetFocus
12310         ElseIf .DateEnd.Enabled = True Then
12320           .DateEnd.SetFocus
12330         Else
12340           .cmdClose.SetFocus
12350         End If
12360       End With
12370     End Select
12380   End If

        ' ** Shift keys.
12390   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12400     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12410       With Me
12420         intRetVal = 0
12430         .cmdClose.SetFocus
12440       End With
12450     End Select
12460   End If

EXITP:
12470   KeyCode = intRetVal
12480   Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

12510   With Me

12520     strAccountNo = vbNullString

12530     If IsNull(.cmbAccounts) = False Then
12540       If Len(.cmbAccounts.Column(CBX_ACT_ACTNO)) > 0 Then
12550         strAccountNo = .cmbAccounts.Column(CBX_ACT_ACTNO)
12560       End If
12570     End If

12580     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
12590       .opgAccountSource_optNumber_lbl.FontBold = True
12600       .opgAccountSource_optName_lbl.FontBold = False
12610     Case .opgAccountSource_optName.OptionValue
12620       .opgAccountSource_optNumber_lbl.FontBold = False
12630       .opgAccountSource_optName_lbl.FontBold = True
12640     End Select

12650     AccountsRowSource_Set Me  ' ** Module Procedure: modCheckReconcile.

12660     .cmbAccounts.Requery
12670     DoEvents

12680   End With

EXITP:
12690   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

12710   With Me
12720     Select Case .chkRememberMe
          Case True
12730       .chkRememberMe_lbl.FontBold = True
12740       .chkRememberMe_lbl2_dim.FontBold = True
12750       .chkRememberMe_lbl2_dim_hi.FontBold = True
12760     Case False
12770       .chkRememberMe_lbl.FontBold = False
12780       .chkRememberMe_lbl2_dim.FontBold = False
12790       .chkRememberMe_lbl2_dim_hi.FontBold = False
12800     End Select
12810   End With

EXITP:
12820   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmbAccounts_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_box_MouseMove"

12910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12920   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmbAccounts_GotFocus()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_GotFocus"

13010   With Me
13020     If IsNull(.cmbAccounts) = False Then
13030       If Trim(.cmbAccounts) <> vbNullString Then
13040         .cmbAccounts.SelLength = 0
13050         .cmbAccounts.SelStart = 99
13060       End If
13070     End If
13080   End With

EXITP:
13090   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

13110   With Me
13120     If .cmbAccounts.text = vbNullString Then
13130       ExpandCombo .cmbAccounts ' ** Module Procedure: modUtilities.
13140     End If
13150   End With

EXITP:
13160   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmbAccounts_BeforeUpdate(Cancel As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_BeforeUpdate"

13210   lngAssetNo_Move = -1&

EXITP:
13220   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Public Sub cmbAccounts_AfterUpdate()
' ** To append checks to tblCheckReconcile_Check, a cracct_id is required for the parent record.
' ** As this procedure runs, that value is often not yet in the on-screen field.
' ** What to do?

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountNo As String, strFilter As String, datDateEnd As Date
        Dim blnAdd As Boolean, blnOutChk As Boolean, blnCrDb As Boolean
        Dim lngRecs As Long, lngAssetNo As Long, lngCRAcctID As Long
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

13310   With Me

13320     If IsNull(.cmbAccounts) = False And .opgCheckingType <> .opgCheckingType_optTotalCash.OptionValue Then
13330       If .cmbAccounts = "CRTC01" Then
13340         .cmbAccounts = Null
13350         DoEvents
13360       End If
13370     End If

13380     blnOutChk = False: blnCrDb = False: blnNotTillDateEntered = False
13390     strTmp01 = "0 of 0"
13400     datDateEnd = Nz(.DateEnd, Date)

13410     Set dbs = CurrentDb

13420     If IsNull(.cmbAccounts) = False Then
13430       If Trim(.cmbAccounts) <> vbNullString Then

13440         If IsNull(.cmbAccounts) = True And .opgCheckingType = .opgCheckingType_optTotalCash.OptionValue Then
13450           .cmbAccounts = "CRTC01"
13460           DoEvents
13470         End If

              ' ** This is really annoying!
13480         Do Until IsNull(.cmbAccounts.Column(CBX_ACT_ACTNO)) = False
13490           AccountsRowSource_Set Me  ' ** Module Procedure: modCheckReconcile.
13500           .cmbAccounts.Requery
13510           If IsNull(.cmbAccounts) = True And .opgCheckingType = .opgCheckingType_optTotalCash.OptionValue Then
13520             .cmbAccounts = "CRTC01"
13530             DoEvents
13540           End If
13550         Loop

13560         strAccountNo = .cmbAccounts.Column(CBX_ACT_ACTNO)
13570         blnAdd = False
13580         lngAssetNo = 0&: lngCRAcctID = 0&

13590         If lngAssetNo_Move >= 0& Then  ' ** From MoveRec().

13600           .cmbAssets.SetFocus  ' ** Since it's a move, it's assumed cmbAssets is already enabled.
13610           If .cracct_id = arr_varAsset_ThisAcct(A_CRACTID, 0) Then
13620             strTmp01 = "1 of " & CStr(lngAssets_ThisAcct)
13630             .cmdAssetPrevious.Enabled = False
13640             .cmdAssetNext.Enabled = True
13650           Else
13660             strTmp01 = "2 of " & CStr(lngAssets_ThisAcct)
13670             .cmdAssetPrevious.Enabled = True
13680             .cmdAssetNext.Enabled = False
13690           End If

13700           SetMultiAsset lngAssets_ThisAcct  ' ** Module Procedure: modCheckReconcile.

13710           lngAssetNo = lngAssetNo_Move
13720           lngCRAcctID = .cracct_id

                ' ** This is used by FormRef() to enable appending the tblCheckReconcile_Check records.
13730           .cracct_id_tmp = lngCRAcctID

13740         Else

13750           lngAssets_ThisAcct = 0&
13760           arr_varAsset_ThisAcct = Empty

13770           With dbs
13780             Select Case strAccountNo
                  Case "CRTC01"
                    ' ** tblCheckReconcile_Account, for accountno = 'CRTC01'.
13790               Set qdf = .QueryDefs("qryCheckReconcile_06d")
13800             Case Else
                    ' ** tblCheckReconcile_Account, by specified [actno].
13810               Set qdf = .QueryDefs("qryCheckReconcile_06a")
13820               With qdf.Parameters
13830                 ![actno] = strAccountNo
13840               End With
13850             End Select
13860             Set rst = qdf.OpenRecordset
13870             With rst
13880               If .BOF = True And .EOF = True Then
13890                 blnAdd = True
13900               End If
13910               .Close
13920             End With  ' ** rst.
13930             Set rst = Nothing
13940             Set qdf = Nothing
13950           End With  ' ** dbs.

13960           If blnAdd = False Then

13970             With dbs
13980               Select Case strAccountNo
                    Case "CRTC01"
                      ' ** tblCheckReconcile_Account, grouped, with cnt of assetno's, for accountno = 'CRTC01'.
13990                 Set qdf = .QueryDefs("qryCheckReconcile_09c")
14000               Case Else
                      ' ** tblCheckReconcile_Account, grouped, with cnt of assetno's, by specified [actno].
14010                 Set qdf = .QueryDefs("qryCheckReconcile_09a")
14020                 With qdf.Parameters
14030                   ![actno] = strAccountNo
14040                 End With
14050               End Select
14060               Set rst = qdf.OpenRecordset
14070               With rst
14080                 If .BOF = True And .EOF = True Then
                        ' ** A record, but still assetno = 0.
14090                 Else
14100                   If IsNull(![cnt]) = False Then
14110                     lngAssets_ThisAcct = ![cnt]
14120                   End If
14130                 End If
14140                 .Close
14150               End With  ' ** rst.
14160               Set rst = Nothing
14170               Set qdf = Nothing
14180             End With  ' ** dbs.

14190             With dbs
14200               Select Case strAccountNo
                    Case "CRTC01"
                      ' ** tblCheckReconcile_Account, for accountno = 'CRTC01'; Cartesian.
14210                 Set qdf = .QueryDefs("qryCheckReconcile_09d")
14220               Case Else
                      ' ** tblCheckReconcile_Account, by specified [actno].
14230                 Set qdf = .QueryDefs("qryCheckReconcile_09b")
14240                 With qdf.Parameters
14250                   ![actno] = strAccountNo
14260                 End With
14270               End Select
14280               Set rst = qdf.OpenRecordset
14290               With rst
14300                 .MoveLast
14310                 .MoveFirst
14320                 arr_varAsset_ThisAcct = .GetRows(lngAssets_ThisAcct)
                      ' ***********************************************
                      ' ** Array: arr_varAsset_ThisAcct()
                      ' **
                      ' **   Field  Element  Name         Constant
                      ' **   =====  =======  ===========  ===========
                      ' **     1       0     cracct_id    A_CRACTID
                      ' **     2       1     accountno    A_ACTNO
                      ' **     3       2     assetno      A_ASTNO
                      ' **     4       3     cusip        A_CUSIP
                      ' **
                      ' ***********************************************
14330                 .Close
14340               End With  ' ** rst.
14350               Set rst = Nothing
14360               Set qdf = Nothing
14370             End With  ' ** dbs.

14380           End If  ' ** blnAdd.

14390           Select Case lngAssets_ThisAcct
                Case 0&
14400             strTmp01 = "0 of 0"
14410             .cmdAssetPrevious.Enabled = False
14420             .cmdAssetNext.Enabled = False
14430           Case 1&
14440             strTmp01 = "1 of 1"
14450             .cmdAssetPrevious.Enabled = False
14460             .cmdAssetNext.Enabled = False
14470           Case Else
14480             strTmp01 = "1 of " & CStr(lngAssets_ThisAcct)
14490             .cmdAssetPrevious.Enabled = False
14500             .cmdAssetNext.Enabled = True
14510           End Select  ' ** lngAssets_ThisAcct.

14520           SetMultiAsset lngAssets_ThisAcct  ' ** Module Procedure: modCheckReconcile.

14530           If blnAdd = True Then
14540             lngAssetNo = 0&
14550           Else
14560             With dbs
14570               Select Case strAccountNo
                    Case "CRTC01"
                      ' ** tblCheckReconcile_Account, for accountno = 'CRTC01'.
14580                 Set qdf = .QueryDefs("qryCheckReconcile_06d")
14590               Case Else
                      ' ** tblCheckReconcile_Account, by specified [actno].
14600                 Set qdf = .QueryDefs("qryCheckReconcile_06a")
14610                 With qdf.Parameters
14620                   ![actno] = strAccountNo
14630                 End With
14640               End Select
14650               Set rst = qdf.OpenRecordset
14660               With rst
14670                 If .BOF = True And .EOF = True Then
                        ' ** Already caught.
14680                 Else
14690                   .MoveFirst
14700                   lngAssetNo = ![assetno]
14710                 End If
14720                 .Close
14730               End With
14740               Set rst = Nothing
14750               Set qdf = Nothing
14760             End With
14770           End If

14780           If blnAdd = True Then
14790             With dbs
14800               Select Case strAccountNo
                    Case "CRTC01"
                      ' ** Append new record to tblCheckReconcile_Account, for accountno = 'CRTC01'.
14810                 Set qdf = .QueryDefs("qryCheckReconcile_07b")
14820                 qdf.Execute
14830                 Set qdf = Nothing
14840                 DoEvents
14850               Case Else
                      ' ** Append new record to tblCheckReconcile_Account, by specified [actno].
14860                 Set qdf = .QueryDefs("qryCheckReconcile_07a")
14870                 With qdf.Parameters
14880                   ![actno] = strAccountNo
14890                 End With
14900                 qdf.Execute
14910                 Set qdf = Nothing
14920                 DoEvents
14930               End Select
14940               DoEvents
14950               varTmp00 = DLookup("[assetno]", "tblCheckReconcile_Account", "[accountno] = '" & strAccountNo & "'")
14960               lngAssetNo = varTmp00
14970               varTmp00 = DLookup("[cracct_id]", "tblCheckReconcile_Account", "[accountno] = '" & strAccountNo & "'")
14980               lngCRAcctID = varTmp00
14990             End With  ' ** dbs.
15000           Else
15010             If lngCRAcctID = 0& Then
15020               varTmp00 = DLookup("[cracct_id]", "tblCheckReconcile_Account", "[accountno] = '" & strAccountNo & "'")
15030               lngCRAcctID = varTmp00
15040             End If
15050           End If  ' ** blnAdd.

                ' ** This is used by FormRef() to enable appending the tblCheckReconcile_Check records.
15060           .cracct_id_tmp = lngCRAcctID

15070           With dbs
                  ' ** Empty tblCheckReconcile_Staging.  '## OK
15080             Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_20")
15090             qdf.Execute
15100             Set qdf = Nothing
15110           End With  ' ** dbs.

15120           Select Case .opgCheckingType
                Case .opgCheckingType_optIndividual.OptionValue
15130             With dbs
                    ' ** Append qryCheckReconcile_OutChecks_21a (Ledger, just 'Paid', by specified [actno])
                    ' ** to tblCheckReconcile_Staging, by specified [actno], [cractid], [datend].
15140               Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_22a")
15150               With qdf.Parameters
15160                 ![actno] = strAccountNo
15170                 ![cractid] = lngCRAcctID
15180                 ![datEnd] = datDateEnd
15190               End With
15200             End With  ' ** dbs.
15210           Case .opgCheckingType_optTotalCash.OptionValue
15220             With dbs
                    ' ** Append qryCheckReconcile_OutChecks_21b (Ledger, just 'Paid', for accountno = 'CRTC01')
                    ' ** to tblCheckReconcile_Staging, for accountno = 'CRTC01', by specified [cractid], [datend].
15230               Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_22b")
15240               With qdf.Parameters
15250                 ![cractid] = lngCRAcctID
15260                 ![datEnd] = datDateEnd
15270               End With
15280             End With  ' ** dbs.
15290           End Select
15300           qdf.Execute dbFailOnError
15310           Set qdf = Nothing

15320           UpdateStageTbl lngAssets_ThisAcct, arr_varAsset_ThisAcct, dbs  ' ** Module Procedure: modCheckReconcile.

15330           lngChks = 0&
15340           ReDim arr_varChk(C_ELEMS, 0)

15350           With dbs

                  ' ** tblCheckReconcile_Staging, just needed fields.  '## OK
15360             Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_23")
15370             Set rst = qdf.OpenRecordset
15380             With rst
15390               If .BOF = True And .EOF = True Then
                      ' ** No unpaid checks.
15400               Else
15410                 .MoveLast
15420                 lngRecs = .RecordCount
15430                 .MoveFirst
15440                 For lngX = 1& To lngRecs
15450                   lngChks = lngChks + 1&
15460                   lngE = lngChks - 1&
15470                   ReDim Preserve arr_varChk(C_ELEMS, lngE)
15480                   arr_varChk(C_ACTID, lngE) = ![cracct_id]
15490                   arr_varChk(C_STGID, lngE) = ![crstage_id]
15500                   arr_varChk(C_JNO, lngE) = ![journalno]
15510                   arr_varChk(C_ACTNO, lngE) = ![accountno]
15520                   arr_varChk(C_CHKID, lngE) = CLng(0)  ' ** croutchk_id.
15530                   arr_varChk(C_CNUM, lngE) = ![CheckNumx]
15540                   arr_varChk(C_PAID, lngE) = ![CheckPaidx]
15550                   arr_varChk(C_DESC, lngE) = ![Descriptionx]
15560                   arr_varChk(C_HDESC, lngE) = ![crstage_hasdesc]
15570                   arr_varChk(C_ASGN, lngE) = ![crstage_assign]
15580                   arr_varChk(C_ASTNO1, lngE) = ![crstage_asset1]
15590                   arr_varChk(C_ASTNO2, lngE) = ![crstage_asset2]
15600                   arr_varChk(C_FND, lngE) = CBool(True)
15610                   If lngX < lngRecs Then .MoveNext
15620                 Next
15630               End If
15640               .Close
15650             End With  ' ** rst.
15660             Set rst = Nothing
15670             Set qdf = Nothing

15680           End With  ' ** dbs.

15690         End If  ' ** lngAssetNo_Move.

15700         If lngAssetNo_Move >= 0& Then
                ' ** Filter already in place.
15710         Else
15720           Select Case .opgCheckingType
                Case .opgCheckingType_optTotalCash.OptionValue
15730             If InStr(.Filter, "CRTC01") = 0 Then
15740               strFilter = "[accountno] = 'CRTC01'"
15750               .Filter = strFilter
15760               .FilterOn = True
15770             End If
15780           Case .opgCheckingType_optIndividual.OptionValue
15790             strFilter = "[accountno] = '" & strAccountNo & "' And [assetno] = " & CStr(lngAssetNo)
15800             .Filter = strFilter
15810             .FilterOn = True
15820           End Select

15830         End If
15840         DoEvents

15850         .cracct_id.Visible = True
15860         .cmbAssets.Requery
15870         .cmbAssets = .assetno
15880         If ((.assetno = 0&) And (IsNull(.cracct_date) = True)) Then
15890           blnNotTillDateEntered = True
15900         End If
15910         If (((.cmbAssets.Enabled = False) And (.opgCheckingType <> .opgCheckingType_optTotalCash.OptionValue)) And _
                  (blnNotTillDateEntered = False)) Then
15920           .cmbAssets.Enabled = True
15930           .cmbAssets.BorderColor = CLR_LTBLU2
15940           .cmbAssets.BackStyle = acBackStyleNormal
15950           .cmbAssets_lbl.BackStyle = acBackStyleNormal
15960           .cmbAssets_lbl_box.Visible = False
15970           .opgAssetSource.Enabled = True
15980           .opgAssetSource_optType_lbl2.ForeColor = CLR_VDKGRY
15990           .opgAssetSource_optType_lbl2_dim_hi.Visible = False
16000           .opgAssetSource_optName_lbl2.ForeColor = CLR_VDKGRY
16010           .opgAssetSource_optName_lbl2_dim_hi.Visible = False
16020           .opgAssetSource_optCusip_lbl2.ForeColor = CLR_VDKGRY
16030           .opgAssetSource_optCusip_lbl2_dim_hi.Visible = False
16040           .cmdAddAsset.Enabled = True
16050           .cmdAddAsset_raised_img.Visible = True
16060           .cmdAddAsset_raised_semifocus_dots_img.Visible = False
16070           .cmdAddAsset_raised_focus_img.Visible = False
16080           .cmdAddAsset_raised_focus_dots_img.Visible = False
16090           .cmdAddAsset_sunken_focus_dots_img.Visible = False
16100           .cmdAddAsset_raised_img_dis.Visible = False
16110         End If

16120         If IsNull(.cracct_date) = False Then
16130           .DateEnd = .cracct_date
16140         End If
16150         If .DateEnd.Enabled = False Then
16160           .DateEnd.Enabled = True
16170           .DateEnd.BorderColor = CLR_LTBLU2
16180           .DateEnd.BackStyle = acBackStyleNormal
16190           .DateEnd_lbl.BackStyle = acBackStyleNormal
16200           .DateEnd_lbl_box.Visible = False
16210           .cmdCalendar.Enabled = True
16220           .cmdCalendar_raised_img.Visible = True
16230           .cmdCalendar_raised_semifocus_dots_img.Visible = False
16240           .cmdCalendar_raised_focus_img.Visible = False
16250           .cmdCalendar_raised_focus_dots_img.Visible = False
16260           .cmdCalendar_sunken_focus_dots_img.Visible = False
16270           .cmdCalendar_raised_img_dis.Visible = False
16280         End If

16290         If .Bank_Name.Enabled = False Then
16300           .Bank_Name.Enabled = True
16310           .Bank_Name.Locked = True
16320           .Bank_Name.BorderColor = CLR_LTBLU2
16330           .Bank_Name.BackStyle = acBackStyleNormal
16340           .Bank_AccountNumber.Enabled = True
16350           .Bank_AccountNumber.Locked = True
16360           .Bank_AccountNumber.BorderColor = CLR_LTBLU2
16370           .Bank_AccountNumber.BackStyle = acBackStyleNormal
16380         End If

16390         .cracct_bsbalance_display = .cracct_bsbalance
16400         If ((.cracct_bsbalance_display.Enabled = False) And (blnNotTillDateEntered = False)) Then
16410           .cracct_bsbalance_display.Enabled = True
16420           .cracct_bsbalance_display.BorderColor = CLR_LTBLU2
16430           .cracct_bsbalance_display.BackStyle = acBackStyleNormal
16440         End If
16450         DoEvents

16460         With dbs

                ' ** Empty tblCheckReconcile_Check.  '## OK
16470           Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_02")
16480           qdf.Execute
16490           Set qdf = Nothing

16500           Select Case Me.chkShowPaid
                Case True
                  ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check, by specified [datend]; All.  '## OK
16510             Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_25")
16520             With qdf.Parameters
16530               ![datEnd] = Me.DateEnd
16540             End With
16550             qdf.Execute
16560             Set qdf = Nothing
16570           Case False
                  ' ** For multi-asset, we need just Zeroes, or their asset!
16580             If lngAssets_ThisAcct > 1& Then
16590               If lngAssetNo_Move >= 0& Then
16600                 If lngAssetNo_Move > 0& Then
                        ' ** Zeroes and this association.
16610                   If lngAssets_ThisAcct > 1& Then
                          ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                          ' ** for CheckPaid = False, by specified [datend], [astno]; Zeroes/Match/ChangedMatch.
16620                     Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_31")  '####  UNKNOWN  ####
16630                   Else
                          ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                          ' ** for CheckPaid = False, by specified [datend], [astno]; Zeroes/Match/Changed.
16640                     Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_29")  '####  UNKNOWN  ####
16650                   End If
16660                   With qdf.Parameters
16670                     ![datEnd] = Me.DateEnd
16680                     ![astno] = lngAssetNo_Move
16690                   End With
16700                 Else
                        ' ** Zeroes only.
                        ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                        ' ** for CheckPaid = False, by specified [datend]; Zeroes.
16710                   Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_30")  '####  UNKNOWN  ####
16720                   With qdf.Parameters
16730                     ![datEnd] = Me.DateEnd
16740                   End With
16750                 End If
16760               Else
                      ' ** Zeroes, and the first association.
16770                 If lngAssets_ThisAcct > 1& Then
                        ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                        ' ** for CheckPaid = False, by specified [datend], [astno]; Zeroes/Match/ChangedMatch.
16780                   Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_31")  '####  UNKNOWN  ####
16790                 Else
                        ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                        ' ** for CheckPaid = False, by specified [datend], [astno]; Zeroes/Match/Changed.
16800                   Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_29")  '####  UNKNOWN  ####
16810                 End If
16820                 With qdf.Parameters
16830                   ![datEnd] = Me.DateEnd
16840                   ![astno] = arr_varAsset_ThisAcct(A_ASTNO, 0)
16850                 End With
16860               End If
16870             Else
                    ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                    ' ** for CheckPaid = False, by specified [datend]; No Criteria.  '## OK
16880               Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_28")
16890               With qdf.Parameters
16900                 ![datEnd] = Nz(Me.DateEnd, Date)
16910               End With
16920             End If
16930             qdf.Execute dbFailOnError
16940             Set qdf = Nothing
16950           End Select

16960         End With  ' ** dbs.
16970         DoEvents

16980         .Requery
16990         DoEvents
17000         If lngAssetNo_Move >= 0& Then
17010           lngAssetNo_Move = 0&
17020         End If

17030         .frmCheckReconcile_Sub_OutChecks.Form.Requery
17040         DoEvents
17050         lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
17060         .OutChecksCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")

17070         If blnNotTillDateEntered = False Then
17080           blnOutChk = True
17090         End If

17100         If .opgCheckingType = .opgCheckingType_optTotalCash.OptionValue Then
17110           .cmbAssets = 0&
17120           .assetno = 0&
17130           blnCrDb = True
17140         ElseIf IsNull(.cmbAssets) = False Then
17150           If .cmbAssets > 0 Then
17160             blnCrDb = True
17170           End If
17180         Else
17190           .cmbAssets = 0&
17200           .assetno = 0&
17210         End If

17220         If .cmdDelete.Enabled = False Then
17230           .cmdDelete.Enabled = True
17240           .cmdDelete_lbl1.ForeColor = CLR_DKGRY
17250           .cmdDelete_lbl1_dim_hi.Visible = False
17260           .cmdDelete_lbl2.ForeColor = CLR_DKGRY
17270           .cmdDelete_lbl2_dim_hi.Visible = False
17280           .cmdClear.Enabled = True
17290           .cmdClear_lbl1.ForeColor = CLR_DKGRY
17300           .cmdClear_lbl1_dim_hi.Visible = False
17310           .cmdClear_lbl2.ForeColor = CLR_DKGRY
17320           .cmdClear_lbl2_dim_hi.Visible = False
17330         End If

17340         If blnNotTillDateEntered = False Then
17350           .cmdCheckAll.Enabled = True
17360           .cmdCheckAll_raised_img.Visible = True
17370           .cmdCheckAll_raised_img_dis.Visible = False
17380           .cmdCheckNone.Enabled = True
17390           .cmdCheckNone_raised_img.Visible = True
17400           .cmdCheckNone_raised_img_dis.Visible = False
17410         End If

17420         With .frmCheckReconcile_Sub_BSCredits.Form
17430           .Detail.BackColor = CLR_WHT
17440           .critem_description.BackStyle = acBackStyleNormal
17450           .critem_amount.BackStyle = acBackStyleNormal
17460         End With
17470         With .frmCheckReconcile_Sub_BSDebits.Form
17480           .Detail.BackColor = CLR_WHT
17490           .critem_description.BackStyle = acBackStyleNormal
17500           .critem_amount.BackStyle = acBackStyleNormal
17510         End With
17520         With .frmCheckReconcile_Sub_TACredits.Form
17530           .Detail.BackColor = CLR_WHT
17540           .critem_description.BackStyle = acBackStyleNormal
17550           .critem_amount.BackStyle = acBackStyleNormal
17560         End With
17570         With .frmCheckReconcile_Sub_TADebits.Form
17580           .Detail.BackColor = CLR_WHT
17590           .critem_description.BackStyle = acBackStyleNormal
17600           .critem_amount.BackStyle = acBackStyleNormal
17610         End With
17620         With .frmCheckReconcile_Sub_OutChecks.Form
17630           .Detail.BackColor = CLR_WHT
17640           .FormHeader.BackColor = CLR_WHT
17650         End With

17660         .BSCreditsCnt_lbl.Caption = "0 Items"
17670         .BSDebitsCnt_lbl.Caption = "0 Items"
17680         .TACreditsCnt_lbl.Caption = "0 Items"
17690         .TADebitsCnt_lbl.Caption = "0 Items"
17700         .OutChecksCnt_lbl.Caption = "0 Items"

17710         If blnUnpostedPaid = True Then
17720           Select Case strAccountNo
                Case "CRTC01"
                  ' ** qryCheckReconcile_11e (Journal, just 'Paid', grouped by journaltype,
                  ' ** PrintCheck, with cnt, for accountno = 'CRTC01'), grouped and summed.
17730             Set qdf = dbs.QueryDefs("qryCheckReconcile_11f")
17740           Case Else
                  ' ** qryCheckReconcile_11c (Journal, just 'Paid', grouped by journaltype,
                  ' ** PrintCheck, with cnt, by specified [actno]), grouped and summed.
17750             Set qdf = dbs.QueryDefs("qryCheckReconcile_11d")
17760             With qdf.Parameters
17770               ![actno] = strAccountNo
17780             End With
17790           End Select
17800           Set rst = qdf.OpenRecordset
17810           If rst.BOF = True And rst.EOF = True Then
17820             .UnpostedJournalMsg_lbl.Visible = False
17830             .cmdClear_box.Visible = True
17840             .cmdClear.Left = lngClear_Left
17850             .cmdClear_lbl1.Visible = True
17860             .cmdClear_lbl2.Visible = True
17870             .cmdClear_vline01.Visible = True
17880             .cmdClear_vline02.Visible = True
17890             If .cmdClear.Enabled = False Then
17900               .cmdClear_lbl1_dim_hi.Visible = True
17910               .cmdClear_lbl2_dim_hi.Visible = True
17920             End If
17930             .cmdDelete_lbl1.Visible = True
17940             .cmdDelete_lbl2.Visible = True
17950             If .cmdDelete.Enabled = False Then
17960               .cmdDelete_lbl2_dim_hi.Visible = True
17970               .cmdDelete_lbl1_dim_hi.Visible = True
17980             End If
17990           Else
18000             rst.MoveFirst
18010             Select Case IsNull(rst![cnt])
                  Case True
18020               .UnpostedJournalMsg_lbl.Visible = False
18030               .cmdClear_box.Visible = True
18040               .cmdClear.Left = lngClear_Left
18050               .cmdClear_lbl1.Visible = True
18060               .cmdClear_lbl2.Visible = True
18070               .cmdClear_vline01.Visible = True
18080               .cmdClear_vline02.Visible = True
18090               If .cmdClear.Enabled = False Then
18100                 .cmdClear_lbl1_dim_hi.Visible = True
18110                 .cmdClear_lbl2_dim_hi.Visible = True
18120               End If
18130               .cmdDelete_lbl1.Visible = True
18140               .cmdDelete_lbl2.Visible = True
18150               If .cmdDelete.Enabled = False Then
18160                 .cmdDelete_lbl2_dim_hi.Visible = True
18170                 .cmdDelete_lbl1_dim_hi.Visible = True
18180               End If
18190             Case False
18200               If rst![cnt] > 0 Then
18210                 .UnpostedJournalMsg_lbl.Caption = JRNL_PAID_ACCT
18220                 .UnpostedJournalMsg_lbl.Visible = True
18230                 .cmdClear_box.Visible = False
18240                 .cmdClear_lbl1.Visible = False
18250                 .cmdClear_lbl1_dim_hi.Visible = False
18260                 .cmdClear_lbl2.Visible = False
18270                 .cmdClear_lbl2_dim_hi.Visible = False
18280                 .cmdClear_vline01.Visible = False
18290                 .cmdClear_vline02.Visible = False
18300                 .cmdDelete_lbl1.Visible = False
18310                 .cmdDelete_lbl1_dim_hi.Visible = False
18320                 .cmdDelete_lbl2.Visible = False
18330                 .cmdDelete_lbl2_dim_hi.Visible = False
18340                 .cmdClear.Left = ((.cmdDelete.Left - .cmdClear.Width) - (16& * lngTpp))
18350               Else
18360                 .UnpostedJournalMsg_lbl.Visible = False
18370                 .cmdClear_box.Visible = True
18380                 .cmdClear.Left = lngClear_Left
18390                 .cmdClear_lbl1.Visible = True
18400                 .cmdClear_lbl2.Visible = True
18410                 .cmdClear_vline01.Visible = True
18420                 .cmdClear_vline02.Visible = True
18430                 If .cmdClear.Enabled = False Then
18440                   .cmdClear_lbl1_dim_hi.Visible = True
18450                   .cmdClear_lbl2_dim_hi.Visible = True
18460                 End If
18470                 .cmdDelete_lbl1.Visible = True
18480                 .cmdDelete_lbl2.Visible = True
18490                 If .cmdDelete.Enabled = False Then
18500                   .cmdDelete_lbl2_dim_hi.Visible = True
18510                   .cmdDelete_lbl1_dim_hi.Visible = True
18520                 End If
18530               End If
18540             End Select
18550           End If
18560           rst.Close
18570           Set rst = Nothing
18580           Set qdf = Nothing
18590         End If

18600         .cmdPreviewReport.Enabled = True
18610         .cmdPreviewReport_raised_img.Visible = True
18620         .cmdPreviewReport_raised_semifocus_dots_img.Visible = False
18630         .cmdPreviewReport_raised_focus_img.Visible = False
18640         .cmdPreviewReport_raised_focus_dots_img.Visible = False
18650         .cmdPreviewReport_sunken_focus_dots_img.Visible = False
18660         .cmdPreviewReport_raised_img_dis.Visible = False

18670         .cmdPrintReport.Enabled = True
18680         .cmdPrintReport_raised_img.Visible = True
18690         .cmdPrintReport_raised_semifocus_dots_img.Visible = False
18700         .cmdPrintReport_raised_focus_img.Visible = False
18710         .cmdPrintReport_raised_focus_dots_img.Visible = False
18720         .cmdPrintReport_sunken_focus_dots_img.Visible = False
18730         .cmdPrintReport_raised_img_dis.Visible = False

18740       Else
              ' ** It's not Null, but it is a NullString.

18750         strFilter = "[accountno] = '0'"
18760         .Filter = strFilter
18770         .FilterOn = True
18780         .cmbAssets = Null
18790         .DateEnd = Null
18800         .cracct_bsbalance_display = Null

18810         With dbs
                ' ** Empty tblCheckReconcile_Check.  '## OK
18820           Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_02")
18830           qdf.Execute
18840           Set qdf = Nothing
18850         End With

18860         .frmCheckReconcile_Sub_OutChecks.Form.Requery
18870         lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
18880         .OutChecksCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
18890         .cracct_id.Visible = False
18900         .cmbAssets.Enabled = False
18910         .cmbAssets.BorderColor = WIN_CLR_DISR
18920         .cmbAssets.BackStyle = acBackStyleTransparent
18930         .cmbAssets_lbl.BackStyle = acBackStyleTransparent
18940         .cmbAssets_lbl_box.Visible = True
18950         .cmbAssets_Cnt_lbl.Caption = "0 of 0"
18960         .opgAssetSource.Enabled = False
18970         .opgAssetSource_optType_lbl2.ForeColor = WIN_CLR_DISF
18980         .opgAssetSource_optType_lbl2_dim_hi.Visible = True
18990         .opgAssetSource_optName_lbl2.ForeColor = WIN_CLR_DISF
19000         .opgAssetSource_optName_lbl2_dim_hi.Visible = True
19010         .opgAssetSource_optCusip_lbl2.ForeColor = WIN_CLR_DISF
19020         .opgAssetSource_optCusip_lbl2_dim_hi.Visible = True
19030         .cmdAddAsset.Enabled = False
19040         .cmdAddAsset_raised_img_dis.Visible = True
19050         .cmdAddAsset_raised_img.Visible = False
19060         .cmdAddAsset_raised_semifocus_dots_img.Visible = False
19070         .cmdAddAsset_raised_focus_img.Visible = False
19080         .cmdAddAsset_raised_focus_dots_img.Visible = False
19090         .cmdAddAsset_sunken_focus_dots_img.Visible = False
19100         .DateEnd.Enabled = False
19110         .DateEnd.BorderColor = WIN_CLR_DISR
19120         .DateEnd.BackStyle = acBackStyleTransparent
19130         .DateEnd_lbl.BackStyle = acBackStyleTransparent
19140         .DateEnd_lbl_box.Visible = True
19150         .cmdCalendar.Enabled = False
19160         .cmdCalendar_raised_img_dis.Visible = True
19170         .cmdCalendar_raised_img.Visible = False
19180         .cmdCalendar_raised_semifocus_dots_img.Visible = False
19190         .cmdCalendar_raised_focus_img.Visible = False
19200         .cmdCalendar_raised_focus_dots_img.Visible = False
19210         .cmdCalendar_sunken_focus_dots_img.Visible = False
19220         .Bank_Name.Enabled = False
19230         .Bank_Name.Locked = False
19240         .Bank_Name.BorderColor = WIN_CLR_DISR
19250         .Bank_Name.BackStyle = acBackStyleTransparent
19260         .Bank_AccountNumber.Enabled = False
19270         .Bank_AccountNumber.Locked = False
19280         .Bank_AccountNumber.BorderColor = WIN_CLR_DISR
19290         .Bank_AccountNumber.BackStyle = acBackStyleTransparent
19300         .cracct_bsbalance_display.Enabled = False
19310         .cracct_bsbalance_display.BorderColor = WIN_CLR_DISR
19320         .cracct_bsbalance_display.BackStyle = acBackStyleTransparent
19330 On Error Resume Next
19340         strTmp01 = Screen.ActiveControl.Name  ' ** If a control doesn't have the focus, it'll error.
19350 On Error GoTo ERRH
19360         If strTmp01 = "cmdDelete" Or strTmp01 = "cmdClear" Then
19370           .cmdClose.SetFocus
19380         End If
19390         .cmdDelete.Enabled = False
19400         .cmdDelete_lbl1.ForeColor = WIN_CLR_DISF
19410         .cmdDelete_lbl2.ForeColor = WIN_CLR_DISF
19420         If .UnpostedJournalMsg_lbl.Visible = False Then
19430           .cmdDelete_lbl1_dim_hi.Visible = True
19440           .cmdDelete_lbl2_dim_hi.Visible = True
19450         End If
19460         .cmdClear.Enabled = False
19470         .cmdClear_lbl1.ForeColor = WIN_CLR_DISF
19480         .cmdClear_lbl2.ForeColor = WIN_CLR_DISF
19490         If .UnpostedJournalMsg_lbl.Visible = False Then
19500           .cmdClear_lbl1_dim_hi.Visible = True
19510           .cmdClear_lbl2_dim_hi.Visible = True
19520         End If
19530         .cmdCheckAll.Enabled = False
19540         .cmdCheckAll_raised_img_dis.Visible = True
19550         .cmdCheckAll_raised_img.Visible = False
19560         .cmdCheckNone.Enabled = False
19570         .cmdCheckNone_raised_img_dis.Visible = True
19580         .cmdCheckNone_raised_img.Visible = False
19590         .BSCreditsCnt_lbl.Caption = "0 Items"
19600         .BSDebitsCnt_lbl.Caption = "0 Items"
19610         .TACreditsCnt_lbl.Caption = "0 Items"
19620         .TADebitsCnt_lbl.Caption = "0 Items"
19630         .OutChecksCnt_lbl.Caption = "0 Items"
19640         If blnUnpostedPaid = True Then
19650           .UnpostedJournalMsg_lbl.Caption = JRNL_PAID
19660           .UnpostedJournalMsg_lbl.Visible = True
19670           .cmdClear_box.Visible = False
19680           .cmdClear_lbl1.Visible = False
19690           .cmdClear_lbl1_dim_hi.Visible = False
19700           .cmdClear_lbl2.Visible = False
19710           .cmdClear_lbl2_dim_hi.Visible = False
19720           .cmdClear_vline01.Visible = False
19730           .cmdClear_vline02.Visible = False
19740           .cmdDelete_lbl1.Visible = False
19750           .cmdDelete_lbl1_dim_hi.Visible = False
19760           .cmdDelete_lbl2.Visible = False
19770           .cmdDelete_lbl2_dim_hi.Visible = False
19780           .cmdClear.Left = ((.cmdDelete.Left - .cmdClear.Width) - (16& * lngTpp))
19790         End If
19800         .cmdPreviewReport.Enabled = False
19810         .cmdPreviewReport_raised_img_dis.Visible = True
19820         .cmdPreviewReport_raised_img.Visible = False
19830         .cmdPreviewReport_raised_semifocus_dots_img.Visible = False
19840         .cmdPreviewReport_raised_focus_img.Visible = False
19850         .cmdPreviewReport_raised_focus_dots_img.Visible = False
19860         .cmdPreviewReport_sunken_focus_dots_img.Visible = False
19870         .cmdPrintReport.Enabled = False
19880         .cmdPrintReport_raised_img_dis.Visible = True
19890         .cmdPrintReport_raised_img.Visible = False
19900         .cmdPrintReport_raised_semifocus_dots_img.Visible = False
19910         .cmdPrintReport_raised_focus_img.Visible = False
19920         .cmdPrintReport_raised_focus_dots_img.Visible = False
19930         .cmdPrintReport_sunken_focus_dots_img.Visible = False

19940         With .frmCheckReconcile_Sub_BSCredits.Form
19950           .Detail.BackColor = MY_CLR_BGE
19960           .critem_description.BackStyle = acBackStyleTransparent
19970           .critem_amount.BackStyle = acBackStyleTransparent
19980         End With
19990         With .frmCheckReconcile_Sub_BSDebits.Form
20000           .Detail.BackColor = MY_CLR_BGE
20010           .critem_description.BackStyle = acBackStyleTransparent
20020           .critem_amount.BackStyle = acBackStyleTransparent
20030         End With
20040         With .frmCheckReconcile_Sub_TACredits.Form
20050           .Detail.BackColor = MY_CLR_BGE
20060           .critem_description.BackStyle = acBackStyleTransparent
20070           .critem_amount.BackStyle = acBackStyleTransparent
20080         End With
20090         With .frmCheckReconcile_Sub_TADebits.Form
20100           .Detail.BackColor = MY_CLR_BGE
20110           .critem_description.BackStyle = acBackStyleTransparent
20120           .critem_amount.BackStyle = acBackStyleTransparent
20130         End With
20140         With .frmCheckReconcile_Sub_OutChecks.Form
20150           .Detail.BackColor = MY_CLR_BGE
20160           .FormHeader.BackColor = MY_CLR_BGE
20170         End With

20180       End If  ' ** vbNullString.

20190     Else
            ' ** This is Null.

20200       strFilter = "[accountno] = '0'"
20210       .Filter = strFilter
20220       .FilterOn = True
20230       .cmbAssets = Null
20240       .DateEnd = Null
20250       .cracct_bsbalance_display = Null

20260       With dbs
              ' ** Empty tblCheckReconcile_Check.  '## OK
20270         Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_02")
20280         qdf.Execute
20290         Set qdf = Nothing
20300       End With

20310       .frmCheckReconcile_Sub_OutChecks.Form.Requery
20320       lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
20330       .OutChecksCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
20340       .cracct_id.Visible = False
20350       .cmbAssets.Enabled = False
20360       .cmbAssets.BorderColor = WIN_CLR_DISR
20370       .cmbAssets.BackStyle = acBackStyleTransparent
20380       .cmbAssets_lbl.BackStyle = acBackStyleTransparent
20390       .cmbAssets_lbl_box.Visible = True
20400       .cmbAssets_Cnt_lbl.Caption = "0 of 0"
20410       .opgAssetSource.Enabled = False
20420       .opgAssetSource_optType_lbl2.ForeColor = WIN_CLR_DISF
20430       .opgAssetSource_optType_lbl2_dim_hi.Visible = True
20440       .opgAssetSource_optName_lbl2.ForeColor = WIN_CLR_DISF
20450       .opgAssetSource_optName_lbl2_dim_hi.Visible = True
20460       .opgAssetSource_optCusip_lbl2.ForeColor = WIN_CLR_DISF
20470       .opgAssetSource_optCusip_lbl2_dim_hi.Visible = True
20480       .cmdAddAsset.Enabled = False
20490       .cmdAddAsset_raised_img_dis.Visible = True
20500       .cmdAddAsset_raised_img.Visible = False
20510       .cmdAddAsset_raised_semifocus_dots_img.Visible = False
20520       .cmdAddAsset_raised_focus_img.Visible = False
20530       .cmdAddAsset_raised_focus_dots_img.Visible = False
20540       .cmdAddAsset_sunken_focus_dots_img.Visible = False
20550       .DateEnd.Enabled = False
20560       .DateEnd.BorderColor = WIN_CLR_DISR
20570       .DateEnd.BackStyle = acBackStyleTransparent
20580       .DateEnd_lbl.BackStyle = acBackStyleTransparent
20590       .DateEnd_lbl_box.Visible = True
20600       .cmdCalendar.Enabled = False
20610       .cmdCalendar_raised_img_dis.Visible = True
20620       .cmdCalendar_raised_img.Visible = False
20630       .cmdCalendar_raised_semifocus_dots_img.Visible = False
20640       .cmdCalendar_raised_focus_img.Visible = False
20650       .cmdCalendar_raised_focus_dots_img.Visible = False
20660       .cmdCalendar_sunken_focus_dots_img.Visible = False
20670       .Bank_Name.Enabled = False
20680       .Bank_Name.Locked = False
20690       .Bank_Name.BorderColor = WIN_CLR_DISR
20700       .Bank_Name.BackStyle = acBackStyleTransparent
20710       .Bank_AccountNumber.Enabled = False
20720       .Bank_AccountNumber.Locked = False
20730       .Bank_AccountNumber.BorderColor = WIN_CLR_DISR
20740       .Bank_AccountNumber.BackStyle = acBackStyleTransparent
20750       .cracct_bsbalance_display.Enabled = False
20760       .cracct_bsbalance_display.BorderColor = WIN_CLR_DISR
20770       .cracct_bsbalance_display.BackStyle = acBackStyleTransparent
20780 On Error Resume Next
20790       strTmp01 = Screen.ActiveControl.Name  ' ** If a control doesn't have the focus, it'll error.
20800 On Error GoTo ERRH
20810       If strTmp01 = "cmdDelete" Or strTmp01 = "cmdClear" Then
20820         .cmdClose.SetFocus
20830       End If
20840       .cmdDelete.Enabled = False
20850       .cmdDelete_lbl1.ForeColor = WIN_CLR_DISF
20860       .cmdDelete_lbl2.ForeColor = WIN_CLR_DISF
20870       If .UnpostedJournalMsg_lbl.Visible = False Then
20880         .cmdDelete_lbl1_dim_hi.Visible = True
20890         .cmdDelete_lbl2_dim_hi.Visible = True
20900       End If
20910       .cmdClear.Enabled = False
20920       .cmdClear_lbl1.ForeColor = WIN_CLR_DISF
20930       .cmdClear_lbl2.ForeColor = WIN_CLR_DISF
20940       If .UnpostedJournalMsg_lbl.Visible = False Then
20950         .cmdClear_lbl1_dim_hi.Visible = True
20960         .cmdClear_lbl2_dim_hi.Visible = True
20970       End If
20980       .cmdCheckAll.Enabled = False
20990       .cmdCheckAll_raised_img_dis.Visible = True
21000       .cmdCheckAll_raised_img.Visible = False
21010       .cmdCheckNone.Enabled = False
21020       .cmdCheckNone_raised_img_dis.Visible = True
21030       .cmdCheckNone_raised_img.Visible = False
21040       .BSCreditsCnt_lbl.Caption = "0 Items"
21050       .BSDebitsCnt_lbl.Caption = "0 Items"
21060       .TACreditsCnt_lbl.Caption = "0 Items"
21070       .TADebitsCnt_lbl.Caption = "0 Items"
21080       .OutChecksCnt_lbl.Caption = "0 Items"
21090       If blnUnpostedPaid = True Then
21100         .UnpostedJournalMsg_lbl.Caption = JRNL_PAID
21110         .UnpostedJournalMsg_lbl.Visible = True
21120         .cmdClear_box.Visible = False
21130         .cmdClear_lbl1.Visible = False
21140         .cmdClear_lbl1_dim_hi.Visible = False
21150         .cmdClear_lbl2.Visible = False
21160         .cmdClear_lbl2_dim_hi.Visible = False
21170         .cmdClear_vline01.Visible = False
21180         .cmdClear_vline02.Visible = False
21190         .cmdDelete_lbl1.Visible = False
21200         .cmdDelete_lbl1_dim_hi.Visible = False
21210         .cmdDelete_lbl2.Visible = False
21220         .cmdDelete_lbl2_dim_hi.Visible = False
21230         .cmdClear.Left = ((.cmdDelete.Left - .cmdClear.Width) - (16& * lngTpp))
21240       End If
21250       .cmdPreviewReport.Enabled = False
21260       .cmdPreviewReport_raised_img_dis.Visible = True
21270       .cmdPreviewReport_raised_img.Visible = False
21280       .cmdPreviewReport_raised_semifocus_dots_img.Visible = False
21290       .cmdPreviewReport_raised_focus_img.Visible = False
21300       .cmdPreviewReport_raised_focus_dots_img.Visible = False
21310       .cmdPreviewReport_sunken_focus_dots_img.Visible = False
21320       .cmdPrintReport.Enabled = False
21330       .cmdPrintReport_raised_img_dis.Visible = True
21340       .cmdPrintReport_raised_img.Visible = False
21350       .cmdPrintReport_raised_semifocus_dots_img.Visible = False
21360       .cmdPrintReport_raised_focus_img.Visible = False
21370       .cmdPrintReport_raised_focus_dots_img.Visible = False
21380       .cmdPrintReport_sunken_focus_dots_img.Visible = False

21390       With .frmCheckReconcile_Sub_BSCredits.Form
21400         .Detail.BackColor = MY_CLR_BGE
21410         .critem_description.BackStyle = acBackStyleTransparent
21420         .critem_amount.BackStyle = acBackStyleTransparent
21430       End With
21440       With .frmCheckReconcile_Sub_BSDebits.Form
21450         .Detail.BackColor = MY_CLR_BGE
21460         .critem_description.BackStyle = acBackStyleTransparent
21470         .critem_amount.BackStyle = acBackStyleTransparent
21480       End With
21490       With .frmCheckReconcile_Sub_TACredits.Form
21500         .Detail.BackColor = MY_CLR_BGE
21510         .critem_description.BackStyle = acBackStyleTransparent
21520         .critem_amount.BackStyle = acBackStyleTransparent
21530       End With
21540       With .frmCheckReconcile_Sub_TADebits.Form
21550         .Detail.BackColor = MY_CLR_BGE
21560         .critem_description.BackStyle = acBackStyleTransparent
21570         .critem_amount.BackStyle = acBackStyleTransparent
21580       End With
21590       With .frmCheckReconcile_Sub_OutChecks.Form
21600         .Detail.BackColor = MY_CLR_BGE
21610         .FormHeader.BackColor = MY_CLR_BGE
21620       End With

21630     End If  ' ** IsNull().

21640     dbs.Close
21650     Set dbs = Nothing

21660     cmdSave_Click  ' ** Procedure: Above.

21670     .frmCheckReconcile_Sub_BSCredits.Form.Requery
21680     .frmCheckReconcile_Sub_BSDebits.Form.Requery
21690     .frmCheckReconcile_Sub_TACredits.Form.Requery
21700     .frmCheckReconcile_Sub_TADebits.Form.Requery

21710     .cmbAssets_Cnt_lbl.Caption = strTmp01
21720     If lngAssets_ThisAcct > 1& Then
            ' ** Previous/Next already appropriately enabled.
21730       .cmbAssets_Cnt_lbl.ForeColor = CLR_DKGRY
21740       .cmdAddAsset.Enabled = False
21750       .cmdAddAsset_raised_img_dis.Visible = True
21760       .cmdAddAsset_raised_img.Visible = False
21770       .cmdAddAsset_raised_semifocus_dots_img.Visible = False
21780       .cmdAddAsset_raised_focus_img.Visible = False
21790       .cmdAddAsset_raised_focus_dots_img.Visible = False
21800       .cmdAddAsset_sunken_focus_dots_img.Visible = False
21810     ElseIf lngAssets_ThisAcct > 0& And .opgCheckingType <> .opgCheckingType_optTotalCash.OptionValue And .cmbAssets > 0 Then
21820       .cmbAssets_Cnt_lbl.ForeColor = WIN_CLR_DISF
21830       .cmdAssetPrevious.Enabled = False
21840       .cmdAssetNext.Enabled = False
21850       .cmdAddAsset.Enabled = True
21860       .cmdAddAsset_raised_img.Visible = True
21870       .cmdAddAsset_raised_semifocus_dots_img.Visible = False
21880       .cmdAddAsset_raised_focus_img.Visible = False
21890       .cmdAddAsset_raised_focus_dots_img.Visible = False
21900       .cmdAddAsset_sunken_focus_dots_img.Visible = False
21910       .cmdAddAsset_raised_img_dis.Visible = False
21920     Else
21930       If IsNull(.cmbAccounts) = True Then
21940         .cmbAssets_Cnt_lbl.ForeColor = WIN_CLR_DISF
21950         .cmdAssetPrevious.Enabled = False
21960         .cmdAssetNext.Enabled = False
21970         .cmdAddAsset.Enabled = False
21980         .cmdAddAsset_raised_img_dis.Visible = True
21990         .cmdAddAsset_raised_img.Visible = False
22000         .cmdAddAsset_raised_semifocus_dots_img.Visible = False
22010         .cmdAddAsset_raised_focus_img.Visible = False
22020         .cmdAddAsset_raised_focus_dots_img.Visible = False
22030         .cmdAddAsset_sunken_focus_dots_img.Visible = False
22040       Else
22050         If Trim(.cmbAccounts) = vbNullString Then
22060           .cmbAssets_Cnt_lbl.ForeColor = WIN_CLR_DISF
22070           .cmdAssetPrevious.Enabled = False
22080           .cmdAssetNext.Enabled = False
22090           .cmdAddAsset.Enabled = False
22100           .cmdAddAsset_raised_img_dis.Visible = True
22110           .cmdAddAsset_raised_img.Visible = False
22120           .cmdAddAsset_raised_semifocus_dots_img.Visible = False
22130           .cmdAddAsset_raised_focus_img.Visible = False
22140           .cmdAddAsset_raised_focus_dots_img.Visible = False
22150           .cmdAddAsset_sunken_focus_dots_img.Visible = False
22160         End If
22170       End If
22180     End If

22190     SetSubForms blnOutChk, blnCrDb  ' ** Module Procedure: modCheckReconcile.

22200     FormValidate  ' ** Procedure: Below.

22210     If .DateEnd.Enabled = True And IsNull(.DateEnd) = True Then
22220       .DateEnd.SetFocus
22230     ElseIf .cmbAssets.Enabled = True And .cmbAssets = 0 Then
22240       .cmbAssets.SetFocus
22250     ElseIf .cracct_bsbalance_display.Enabled = True Then
22260       .cracct_bsbalance_display.SetFocus
22270     Else
22280       .cmdClose.SetFocus
22290     End If

22300   End With  ' ** Me.

        'HERE!  Form_Open()
        'HERE! 1  cmbAccounts_AfterUpdate()
        'HERE!  CheckingType_Set()
        'HERE! 1  cmbAccounts_AfterUpdate()

EXITP:
22310   Set rst = Nothing
22320   Set qdf = Nothing
22330   Set rst = Nothing
22340   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intLoop As Integer
        Dim intRetVal As Integer

22410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22420   intShiftDown = (Shift And acShiftMask) > 0
22430   intAltDown = (Shift And acAltMask) > 0
22440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22470       With Me
22480         intRetVal = 0
22490         If .DateEnd.Enabled = True Then
22500           .DateEnd.SetFocus
22510         Else
22520           intLoop = 0
22530 On Error Resume Next
22540           Do Until Screen.ActiveControl.Name = "cmdClose"
                  ' ** I don't know why this errors!
22550             .cmdClose.SetFocus
22560             intLoop = intLoop + 1&
22570             If intLoop > 500 Then Exit Do
22580           Loop
22590 On Error GoTo ERRH
22600         End If
22610       End With
22620     End Select
22630   End If

        ' ** Shift keys.
22640   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22650     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22660       With Me
22670         intRetVal = 0
22680         .opgCheckingType.SetFocus
22690       End With
22700     End Select
22710   End If

EXITP:
22720   KeyCode = intRetVal
22730   Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub cmbAccounts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_MouseMove"

22810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
22820   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub DateEnd_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_box_MouseMove"

22910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
22920   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub DateEnd_AfterUpdate()

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_AfterUpdate"

        Dim lngRecs As Long
        Dim lngTmp01 As Long, arr_varTmp02 As Variant
        Dim lngX As Long, lngY As Long, lngE As Long

23010   With Me
23020     If IsNull(.DateEnd) = False Then
23030       If IsDate(.DateEnd) = True Then
23040         .cracct_date = .DateEnd
23050         cmdSave_Click  ' ** Procedure: Above.
23060         DoEvents
23070         If blnNotTillDateEntered = True Then
23080           blnNotTillDateEntered = False
23090           If .cmbAssets.Enabled = False Then
23100             .cmbAssets.Enabled = True
23110             .cmbAssets.BorderColor = CLR_LTBLU2
23120             .cmbAssets.BackStyle = acBackStyleNormal
23130             .cmbAssets_lbl.BackStyle = acBackStyleNormal
23140             .cmbAssets_lbl_box.Visible = False
23150             .opgAssetSource.Enabled = True
23160             .opgAssetSource_optType_lbl2.ForeColor = CLR_VDKGRY
23170             .opgAssetSource_optType_lbl2_dim_hi.Visible = False
23180             .opgAssetSource_optName_lbl2.ForeColor = CLR_VDKGRY
23190             .opgAssetSource_optName_lbl2_dim_hi.Visible = False
23200             .opgAssetSource_optCusip_lbl2.ForeColor = CLR_VDKGRY
23210             .opgAssetSource_optCusip_lbl2_dim_hi.Visible = False
23220           End If
23230           If .cmbAssets > 0 And .cmdAddAsset.Enabled = False Then
23240             .cmdAddAsset.Enabled = True
23250             .cmdAddAsset_raised_img.Visible = True
23260             .cmdAddAsset_raised_semifocus_dots_img.Visible = False
23270             .cmdAddAsset_raised_focus_img.Visible = False
23280             .cmdAddAsset_raised_focus_dots_img.Visible = False
23290             .cmdAddAsset_sunken_focus_dots_img.Visible = False
23300             .cmdAddAsset_raised_img_dis.Visible = False
23310           End If
23320           If .cracct_bsbalance_display.Enabled = False Then
23330             .cracct_bsbalance_display.Enabled = True
23340             .cracct_bsbalance_display.BorderColor = CLR_LTBLU2
23350             .cracct_bsbalance_display.BackStyle = acBackStyleNormal
23360           End If
23370           If .cmdCheckAll.Enabled = False Then
23380             .cmdCheckAll.Enabled = True
23390             .cmdCheckAll_raised_img.Visible = True
23400             .cmdCheckAll_raised_semifocus_dots_img.Visible = False
23410             .cmdCheckAll_raised_focus_img.Visible = False
23420             .cmdCheckAll_raised_focus_dots_img.Visible = False
23430             .cmdCheckAll_sunken_focus_dots_img.Visible = False
23440             .cmdCheckAll_raised_img_dis.Visible = False
23450             .cmdCheckNone.Enabled = True
23460             .cmdCheckNone_raised_img.Visible = True
23470             .cmdCheckNone_raised_semifocus_dots_img.Visible = False
23480             .cmdCheckNone_raised_focus_img.Visible = False
23490             .cmdCheckNone_raised_focus_dots_img.Visible = False
23500             .cmdCheckNone_sunken_focus_dots_img.Visible = False
23510             .cmdCheckNone_raised_img_dis.Visible = False
23520           End If
23530         End If
23540       End If
23550     End If
23560     DoEvents
23570     lngTmp01 = lngChks
23580     arr_varTmp02 = arr_varChk
23590     ResetCheckSub lngTmp01, arr_varTmp02, lngAssets_ThisAcct, arr_varAsset_ThisAcct, lngAssetNo_Move, Me  ' ** Module Function: modCheckReconcile.
23600     lngRecs = lngTmp01
23610     lngChks = 0&
23620     ReDim arr_varChk(C_ELEMS, 0)
23630     For lngX = 0& To (lngRecs - 1&)
23640       lngChks = lngChks + 1&
23650       lngE = lngChks - 1&
23660       ReDim Preserve arr_varChk(C_ELEMS, lngE)
23670       For lngY = 0& To C_ELEMS
23680         arr_varChk(lngY, lngE) = arr_varTmp02(lngY, lngX)
23690       Next
23700     Next
23710     DoEvents
23720     If .cmbAssets.Enabled = True Then
23730       .cmbAssets.SetFocus
23740     End If
23750   End With

EXITP:
23760   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

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
23890         If .Bank_Name.Enabled = True And .Bank_Name.Locked = False Then
23900 On Error Resume Next
23910           .Bank_Name.SetFocus
23920 On Error GoTo ERRH
23930         ElseIf .cmbAssets.Enabled = True Then
23940           .cmbAssets.SetFocus
23950         ElseIf .cracct_bsbalance_display.Enabled = True Then
23960           .cracct_bsbalance_display.SetFocus
23970         ElseIf .cmdPreviewReport.Enabled = True Then
23980           .cmdPreviewReport.SetFocus
23990         ElseIf .cmdClear.Enabled = True Then
24000           .cmdClear.SetFocus
24010         ElseIf .cmdDelete.Enabled = True Then
24020           .cmdDelete.SetFocus
24030         Else
24040           .cmdClose.SetFocus
24050         End If
24060       End With
24070     End Select
24080   End If

        ' ** Shift keys.
24090   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24110       With Me
24120         intRetVal = 0
24130         If .cmbAccounts.Enabled = True Then
24140           .cmbAccounts.SetFocus
24150         Else
24160           .opgCheckingType.SetFocus
24170         End If
24180       End With
24190     End Select
24200   End If

EXITP:
24210   KeyCode = intRetVal
24220   Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore
710     Case Else
720       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
730     End Select
740     Resume EXITP

End Sub

Private Sub DateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_MouseMove"

24310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
24320   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub DateEnd_Exit(Cancel As Integer)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

24410   With Me
24420     If IsNull(.DateEnd) = True And blnFromTimer = False Then
24430       .DateEnd = Date
24440       DateEnd_AfterUpdate  ' ** Procedure: Above.
24450     End If
24460   End With

EXITP:
24470   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

24510   With Me
24520     blnCalendar1_Focus = True
24530     .cmdCalendar_raised_semifocus_dots_img.Visible = True
24540     .cmdCalendar_raised_img.Visible = False
24550     .cmdCalendar_raised_focus_img.Visible = False
24560     .cmdCalendar_raised_focus_dots_img.Visible = False
24570     .cmdCalendar_sunken_focus_dots_img.Visible = False
24580     .cmdCalendar_raised_img_dis.Visible = False
24590   End With

EXITP:
24600   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

24710   With Me
24720     blnCalendar1_MouseDown = True
24730     .cmdCalendar_sunken_focus_dots_img.Visible = True
24740     .cmdCalendar_raised_img.Visible = False
24750     .cmdCalendar_raised_semifocus_dots_img.Visible = False
24760     .cmdCalendar_raised_focus_img.Visible = False
24770     .cmdCalendar_raised_focus_dots_img.Visible = False
24780     .cmdCalendar_raised_img_dis.Visible = False
24790   End With

EXITP:
24800   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Public Sub cmdCalendar_Click()

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim blnRetVal As Boolean
        Dim datStartDate As Date, datEndDate As Date

24910   With Me
24920     datStartDate = Date
24930     datEndDate = 0
24940     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
24950     If blnRetVal = True Then
24960       .DateEnd = datStartDate
24970     Else
24980       .DateEnd = Date
24990     End If
25000     DateEnd_AfterUpdate  ' ** Procedure: Below.
25010   End With

EXITP:
25020   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCalendar_KeyDown(KeyCode As Integer, Shift As Integer)

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25120   intShiftDown = (Shift And acShiftMask) > 0
25130   intAltDown = (Shift And acAltMask) > 0
25140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25160     Select Case intRetVal
          Case vbKeyTab
25170       With Me
25180         intRetVal = 0
25190         If .Bank_Name.Enabled = True And .Bank_Name.Locked = False Then
25200           .Bank_Name.SetFocus
25210         ElseIf .cmbAssets.Enabled = True Then
25220           .cmbAssets.SetFocus
25230         ElseIf .cracct_bsbalance_display.Enabled = True Then
25240           .cracct_bsbalance_display.SetFocus
25250         Else
25260           .cmdClose.SetFocus
25270         End If
25280       End With
25290     End Select
25300   End If

        ' ** Shift keys.
25310   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25320     Select Case intRetVal
          Case vbKeyTab
25330       With Me
25340         intRetVal = 0
25350         .DateEnd.SetFocus
25360       End With
25370     End Select
25380   End If

EXITP:
25390   KeyCode = intRetVal
25400   Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

25510   With Me
25520     If blnCalendar1_MouseDown = False Then
25530       Select Case blnCalendar1_Focus
            Case True
25540         .cmdCalendar_raised_focus_dots_img.Visible = True
25550         .cmdCalendar_raised_focus_img.Visible = False
25560       Case False
25570         .cmdCalendar_raised_focus_img.Visible = True
25580         .cmdCalendar_raised_focus_dots_img.Visible = False
25590       End Select
25600       .cmdCalendar_raised_img.Visible = False
25610       .cmdCalendar_raised_semifocus_dots_img.Visible = False
25620       .cmdCalendar_sunken_focus_dots_img.Visible = False
25630       .cmdCalendar_raised_img_dis.Visible = False
25640     End If
25650   End With

EXITP:
25660   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

25710   With Me
25720     .cmdCalendar_raised_focus_dots_img.Visible = True
25730     .cmdCalendar_raised_img.Visible = False
25740     .cmdCalendar_raised_semifocus_dots_img.Visible = False
25750     .cmdCalendar_raised_focus_img.Visible = False
25760     .cmdCalendar_sunken_focus_dots_img.Visible = False
25770     .cmdCalendar_raised_img_dis.Visible = False
25780     blnCalendar1_MouseDown = False
25790   End With

EXITP:
25800   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCalendar_LostFocus()

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

25910   With Me
25920     .cmdCalendar_raised_img.Visible = True
25930     .cmdCalendar_raised_semifocus_dots_img.Visible = False
25940     .cmdCalendar_raised_focus_img.Visible = False
25950     .cmdCalendar_raised_focus_dots_img.Visible = False
25960     .cmdCalendar_sunken_focus_dots_img.Visible = False
25970     .cmdCalendar_raised_img_dis.Visible = False
25980     blnCalendar1_Focus = False
25990   End With

EXITP:
26000   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub Bank_Name_KeyDown(KeyCode As Integer, Shift As Integer)

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_Name_KeyDown"

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
26190         If .Bank_AccountNumber.Enabled = True And .Bank_AccountNumber.Locked = False Then
26200           .Bank_AccountNumber.SetFocus
26210         ElseIf .cmbAssets.Enabled = True Then
26220           .cmbAssets.SetFocus
26230         ElseIf .cracct_bsbalance_display.Enabled = True Then
26240           .cracct_bsbalance_display.SetFocus
26250         Else
26260           .cmdClose.SetFocus
26270         End If
26280       End With
26290     End Select
26300   End If

        ' ** Shift keys.
26310   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26320     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26330       With Me
26340         intRetVal = 0
26350         .DateEnd.SetFocus
26360       End With
26370     End Select
26380   End If

EXITP:
26390   KeyCode = intRetVal
26400   Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub Bank_AccountNumber_KeyDown(KeyCode As Integer, Shift As Integer)

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_AccountNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26520   intShiftDown = (Shift And acShiftMask) > 0
26530   intAltDown = (Shift And acAltMask) > 0
26540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26570       With Me
26580         intRetVal = 0
26590         If .cmbAssets.Enabled = True Then
26600           .cmbAssets.SetFocus
26610         ElseIf .cracct_bsbalance_display.Enabled = True Then
26620           .cracct_bsbalance_display.SetFocus
26630         Else
26640           .cmdClose.SetFocus
26650         End If
26660       End With
26670     End Select
26680   End If

        ' ** Shift keys.
26690   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26700     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26710       With Me
26720         intRetVal = 0
26730         .Bank_Name.SetFocus
26740       End With
26750     End Select
26760   End If

EXITP:
26770   KeyCode = intRetVal
26780   Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub opgAssetSource_AfterUpdate()

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_AfterUpdate"

        Dim lngAssetNo As Long

26810   With Me

26820     lngAssetNo = 0&

26830     If IsNull(.cmbAssets) = False Then
26840       If Len(.cmbAssets.Column(CBX_AST_ASTNO)) > 0 Then
26850         lngAssetNo = .cmbAssets.Column(CBX_AST_ASTNO)
26860       End If
26870     End If

26880     Select Case .opgAssetSource
          Case .opgAssetSource_optType.OptionValue
26890       .opgAssetSource_optType_lbl2.FontBold = True
26900       .opgAssetSource_optType_lbl2_dim_hi.FontBold = True
26910       .opgAssetSource_optName_lbl2.FontBold = False
26920       .opgAssetSource_optName_lbl2_dim_hi.FontBold = False
26930       .opgAssetSource_optCusip_lbl2.FontBold = False
26940       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = False
26950       If .cmbAssets.RowSource <> "qryCheckReconcile_05a" Then
26960         .cmbAssets.RowSource = "qryCheckReconcile_05a"
26970       End If
26980     Case .opgAssetSource_optName.OptionValue
26990       .opgAssetSource_optType_lbl2.FontBold = False
27000       .opgAssetSource_optType_lbl2_dim_hi.FontBold = False
27010       .opgAssetSource_optName_lbl2.FontBold = True
27020       .opgAssetSource_optName_lbl2_dim_hi.FontBold = True
27030       .opgAssetSource_optCusip_lbl2.FontBold = False
27040       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = False
27050       If .cmbAssets.RowSource <> "qryCheckReconcile_05b" Then
27060         .cmbAssets.RowSource = "qryCheckReconcile_05b"
27070       End If
27080     Case .opgAssetSource_optCusip.OptionValue
27090       .opgAssetSource_optType_lbl2.FontBold = False
27100       .opgAssetSource_optType_lbl2_dim_hi.FontBold = False
27110       .opgAssetSource_optName_lbl2.FontBold = False
27120       .opgAssetSource_optName_lbl2_dim_hi.FontBold = False
27130       .opgAssetSource_optCusip_lbl2.FontBold = True
27140       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = True
27150       If .cmbAssets.RowSource <> "qryCheckReconcile_05c" Then
27160         .cmbAssets.RowSource = "qryCheckReconcile_05c"
27170       End If
27180     End Select

27190     .cmbAssets.Requery
27200     DoEvents

27210   End With

EXITP:
27220   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmbAssets_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_box_MouseMove"

27310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
27320   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdAddAsset_GotFocus()

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAddAsset_GotFocus"

27410   With Me
27420     blnAddAsset_Focus = True
27430     .cmdAddAsset_raised_semifocus_dots_img.Visible = True
27440     .cmdAddAsset_raised_img.Visible = False
27450     .cmdAddAsset_raised_focus_img.Visible = False
27460     .cmdAddAsset_raised_focus_dots_img.Visible = False
27470     .cmdAddAsset_sunken_focus_dots_img.Visible = False
27480     .cmdAddAsset_raised_img_dis.Visible = False
27490   End With

EXITP:
27500   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdAddAsset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAddAsset_MouseDown"

27610   With Me
27620     blnAddAsset_MouseDown = True
27630     .cmdAddAsset_sunken_focus_dots_img.Visible = True
27640     .cmdAddAsset_raised_img.Visible = False
27650     .cmdAddAsset_raised_semifocus_dots_img.Visible = False
27660     .cmdAddAsset_raised_focus_img.Visible = False
27670     .cmdAddAsset_raised_focus_dots_img.Visible = False
27680     .cmdAddAsset_raised_img_dis.Visible = False
27690   End With

EXITP:
27700   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Public Sub cmdAddAsset_Click()

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAddAsset_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCRAcctID1 As Long, lngCRAcctID2 As Long
        Dim strAccountNo As String, datDateEnd As Date

27810   With Me
27820     If .cmbAssets > 0& Then

27830       lngCRAcctID1 = .cracct_id
27840       strAccountNo = .cmbAccounts.Column(CBX_ACT_ACTNO)
27850       datDateEnd = .DateEnd
27860       Set dbs = CurrentDb
27870       With dbs
27880         Set rst = .OpenRecordset("tblCheckReconcile_Account", dbOpenDynaset, dbConsistent)
27890         With rst
27900           .AddNew
27910           ![accountno] = strAccountNo
27920           ![assetno] = 0&
                ' ** ![cracct_id] : AutoNumber.
27930           ![cracct_date] = datDateEnd
27940           ![cracct_bsbalance] = 0@
27950           ![cracct_datemodified] = Now()
27960           .Update
27970           .Bookmark = .LastModified
27980           lngCRAcctID2 = ![cracct_id]
27990           .Close
28000         End With
28010         Set rst = Nothing
28020         DoEvents
28030         .Close
28040       End With

28050       MoveRec 0, lngCRAcctID2  ' ** Procedure: Below.
28060       DoEvents

28070       cmbAccounts_AfterUpdate  ' ** Procedure: Above.

28080       .cmdAssetPrevious.Enabled = True
28090       .cmdAssetNext.Enabled = True

28100       .cmbAssets.SetFocus

28110     Else
28120       Beep
28130     End If
28140   End With

        'UPDATE cracct_datemodified!

EXITP:
28150   Set rst = Nothing
28160   Set qdf = Nothing
28170   Set dbs = Nothing
28180   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdAddAsset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAddAsset_MouseMove"

28210   With Me
28220     If blnAddAsset_MouseDown = False Then
28230       Select Case blnAddAsset_Focus
            Case True
28240         .cmdAddAsset_raised_focus_dots_img.Visible = True
28250         .cmdAddAsset_raised_focus_img.Visible = False
28260       Case False
28270         .cmdAddAsset_raised_focus_img.Visible = True
28280         .cmdAddAsset_raised_focus_dots_img.Visible = False
28290       End Select
28300       .cmdAddAsset_raised_img.Visible = False
28310       .cmdAddAsset_raised_semifocus_dots_img.Visible = False
28320       .cmdAddAsset_sunken_focus_dots_img.Visible = False
28330       .cmdAddAsset_raised_img_dis.Visible = False
28340     End If
28350   End With

EXITP:
28360   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdAddAsset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAddAsset_MouseUp"

28410   With Me
28420     .cmdAddAsset_raised_focus_dots_img.Visible = True
28430     .cmdAddAsset_raised_img.Visible = False
28440     .cmdAddAsset_raised_semifocus_dots_img.Visible = False
28450     .cmdAddAsset_raised_focus_img.Visible = False
28460     .cmdAddAsset_sunken_focus_dots_img.Visible = False
28470     .cmdAddAsset_raised_img_dis.Visible = False
28480     blnAddAsset_MouseDown = False
28490   End With

EXITP:
28500   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdAddAsset_LostFocus()

28600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAddAsset_LostFocus"

28610   With Me
28620     .cmdAddAsset_raised_img.Visible = True
28630     .cmdAddAsset_raised_semifocus_dots_img.Visible = False
28640     .cmdAddAsset_raised_focus_img.Visible = False
28650     .cmdAddAsset_raised_focus_dots_img.Visible = False
28660     .cmdAddAsset_sunken_focus_dots_img.Visible = False
28670     .cmdAddAsset_raised_img_dis.Visible = False
28680     blnAddAsset_Focus = False
28690   End With

EXITP:
28700   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Public Sub cmdAssetPrevious_Click()

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPrevious_Click"

28810   If lngAssets_ThisAcct > 1& Then
28820     DoCmd.Hourglass True
28830     DoEvents
28840     MoveRec acCmdRecordsGoToPrevious, 0, True  ' ** Procedure: Below.
28850   Else
28860     Beep
28870   End If

EXITP:
28880   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Public Sub cmdAssetNext_Click()

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetNext_Click"

28910   If lngAssets_ThisAcct > 1& Then
28920     DoCmd.Hourglass True
28930     DoEvents
28940     MoveRec acCmdRecordsGoToNext, 0, True  ' ** Procedure: Below.
28950   Else
28960     Beep
28970   End If

EXITP:
28980   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmbAssets_Cnt_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_Cnt_lbl_MouseMove"

29010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
29020   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmbAssets_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_MouseMove"

29110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
29120   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmbAssets_GotFocus()

29200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_GotFocus"

29210   With Me
29220     If IsNull(.cmbAssets) = False Then
29230       If .cmbAssets > 0 Then
29240         .cmbAssets.SelLength = 0
29250         .cmbAssets.SelStart = 99
29260       End If
29270     End If
29280   End With

EXITP:
29290   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmbAssets_Enter()

29300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_Enter"

29310   With Me
29320     If .cmbAssets.text = vbNullString And IsNull(.cmbAccounts) = False Then
29330       ExpandCombo .cmbAssets  ' ** Module Procedure: modUtilities.
29340     End If
29350   End With

EXITP:
29360   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmbAssets_AfterUpdate()

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

        Dim blnOutChk As Boolean, blnCrDb As Boolean

29410   With Me

29420     blnOutChk = False: blnCrDb = False
29430     If IsNull(.cmbAssets) = False Then
29440       If .cmbAssets > 0 Then
29450         .assetno = .cmbAssets
29460         If IsNull(.cmbAccounts) = False Then
29470           If Trim(.cmbAccounts) <> vbNullString Then
29480             blnCrDb = True
29490             blnOutChk = True
29500             If .cmdAddAsset.Enabled = False Then
29510               .cmdAddAsset.Enabled = True
29520               .cmdAddAsset_raised_img.Visible = True
29530               .cmdAddAsset_raised_semifocus_dots_img.Visible = False
29540               .cmdAddAsset_raised_focus_img.Visible = False
29550               .cmdAddAsset_raised_focus_dots_img.Visible = False
29560               .cmdAddAsset_sunken_focus_dots_img.Visible = False
29570               .cmdAddAsset_raised_img_dis.Visible = False
29580             End If
29590           End If
29600         End If
29610       Else
29620         .assetno = 0
29630       End If
29640     Else
29650       .assetno = 0
29660       If IsNull(.cmbAccounts) = False Then
29670         If Trim(.cmbAccounts) <> vbNullString Then
29680           blnOutChk = True
29690         End If
29700       End If
29710     End If

29720     cmdSave_Click  ' ** Procedure: Above.

29730     .frmCheckReconcile_Sub_BSCredits.Form.Requery
29740     .frmCheckReconcile_Sub_BSDebits.Form.Requery
29750     .frmCheckReconcile_Sub_TACredits.Form.Requery
29760     .frmCheckReconcile_Sub_TADebits.Form.Requery

29770     SetSubForms blnOutChk, blnCrDb  ' ** Module Procedure: modCheckReconcile.

29780     FormValidate  ' ** Procedure: Below.

29790     If blnOutChk = True And blnCrDb = True Then
29800       .cracct_bsbalance_display.SetFocus
29810     End If

29820   End With

EXITP:
29830   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub BSBalance_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29900 On Error GoTo ERRH

        Const THIS_PROC As String = "BSBalance_box_MouseMove"

29910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
29920   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cracct_bsbalance_display_GotFocus()

30000 On Error GoTo ERRH

        Const THIS_PROC As String = "cracct_bsbalance_display_GotFocus"

30010   With Me
30020     If IsNull(.cracct_bsbalance_display) = False Then
30030       If .cracct_bsbalance_display <> 0 Then
30040         .cracct_bsbalance_display.SelLength = 0
30050         .cracct_bsbalance_display.SelStart = 0
30060       End If
30070     End If
30080   End With

EXITP:
30090   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cracct_bsbalance_display_BeforeUpdate(Cancel As Integer)

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "cracct_bsbalance_display_BeforeUpdate"

30110   With Me
30120     If IsNull(.cmbAccounts) = False Then
30130       If Trim(.cmbAccounts) <> vbNullString Then
30140         UpdateBSTotal_CR Me, lngAssets_ThisAcct, arr_varAsset_ThisAcct, THIS_PROC ' ** Module Procedure: modCheckReconcile.
30150       Else
30160         If .cracct_bsbalance_display.text <> vbNullString Then
30170           Cancel = -1
30180         End If
30190       End If
30200     Else
30210       If .cracct_bsbalance_display.text <> vbNullString Then
30220         Cancel = -1
30230       End If
30240     End If
30250   End With

EXITP:
30260   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cracct_bsbalance_display_AfterUpdate()

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "cracct_bsbalance_display_AfterUpdate"

30310   With Me
30320     .cracct_bsbalance = .cracct_bsbalance_display
30330     cmdSave_Click  ' ** Procedure: Above.
30340   End With

EXITP:
30350   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cracct_bsbalance_display_KeyDown(KeyCode As Integer, Shift As Integer)

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "cracct_bsbalance_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
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
30480         intRetVal = 0
30490         lngRecsCur = .frmCheckReconcile_Sub_BSCredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSCredits.
30500         .frmCheckReconcile_Sub_BSCredits.SetFocus
30510         If lngRecsCur > 0& Then
30520           .frmCheckReconcile_Sub_BSCredits.Form.FocusHolder.SetFocus
30530         Else
30540           .frmCheckReconcile_Sub_BSCredits.Form.critem_description.SetFocus
30550         End If
30560       End With
30570     End Select
30580   End If

        ' ** Shift keys.
30590   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30600     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30610       With Me
30620         intRetVal = 0
30630         If .cmbAssets.Enabled = True Then
30640           .cmbAssets.SetFocus
30650         ElseIf .Bank_AccountNumber.Enabled = True And .Bank_AccountNumber.Locked = False Then
30660           .Bank_AccountNumber.SetFocus
30670         ElseIf .DateEnd.Enabled = True Then
30680           .DateEnd.SetFocus
30690         ElseIf .cmbAccounts.Enabled = True Then
30700           .cmbAccounts.SetFocus
30710         Else
30720           .opgCheckingType.SetFocus
30730         End If
30740       End With
30750     End Select
30760   End If

EXITP:
30770   KeyCode = intRetVal
30780   Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub cracct_bsbalance_display_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30800 On Error GoTo ERRH

        Const THIS_PROC As String = "cracct_bsbalance_display_MouseMove"

30810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
30820   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub frmCheckReconcile_Sub_BSCredits_Enter()

30900 On Error GoTo ERRH

        Const THIS_PROC As String = "frmCheckReconcile_Sub_BSCredits_Enter"

30910   With Me
30920     With .frmCheckReconcile_Sub_BSCredits.Form
30930       .Header_Arrow_lbl.Visible = True
30940       .Header_Arrow_lbl_line1.Visible = True
30950       .Header_Arrow_lbl_line2.Visible = True
30960     End With
30970   End With

EXITP:
30980   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub frmCheckReconcile_Sub_BSCredits_Exit(Cancel As Integer)

31000 On Error GoTo ERRH

        Const THIS_PROC As String = "frmCheckReconcile_Sub_BSCredits_Exit"

31010   With Me
31020     With .frmCheckReconcile_Sub_BSCredits.Form
31030       .Header_Arrow_lbl.Visible = False
31040       .Header_Arrow_lbl_line1.Visible = False
31050       .Header_Arrow_lbl_line2.Visible = False
31060     End With
31070   End With

EXITP:
31080   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub frmCheckReconcile_Sub_BSDebits_Enter()

31100 On Error GoTo ERRH

        Const THIS_PROC As String = "frmCheckReconcile_Sub_BSDebits_Enter"

31110   With Me
31120     With .frmCheckReconcile_Sub_BSDebits.Form
31130       .Header_Arrow_lbl.Visible = True
31140       .Header_Arrow_lbl_line1.Visible = True
31150       .Header_Arrow_lbl_line2.Visible = True
31160     End With
31170   End With

EXITP:
31180   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub frmCheckReconcile_Sub_BSDebits_Exit(Cancel As Integer)

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "frmCheckReconcile_Sub_BSDebits_Exit"

31210   With Me
31220     With .frmCheckReconcile_Sub_BSDebits.Form
31230       .Header_Arrow_lbl.Visible = False
31240       .Header_Arrow_lbl_line1.Visible = False
31250       .Header_Arrow_lbl_line2.Visible = False
31260     End With
31270   End With

EXITP:
31280   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub frmCheckReconcile_Sub_TACredits_Enter()

31300 On Error GoTo ERRH

        Const THIS_PROC As String = "frmCheckReconcile_Sub_TACredits_Enter"

31310   With Me
31320     With .frmCheckReconcile_Sub_TACredits.Form
31330       .Header_Arrow_lbl.Visible = True
31340       .Header_Arrow_lbl_line1.Visible = True
31350       .Header_Arrow_lbl_line2.Visible = True
31360     End With
31370   End With

EXITP:
31380   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub frmCheckReconcile_Sub_TACredits_Exit(Cancel As Integer)

31400 On Error GoTo ERRH

        Const THIS_PROC As String = "frmCheckReconcile_Sub_TACredits_Exit"

31410   With Me
31420     With .frmCheckReconcile_Sub_TACredits.Form
31430       .Header_Arrow_lbl.Visible = False
31440       .Header_Arrow_lbl_line1.Visible = False
31450       .Header_Arrow_lbl_line2.Visible = False
31460     End With
31470   End With

EXITP:
31480   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub frmCheckReconcile_Sub_TADebits_Enter()

31500 On Error GoTo ERRH

        Const THIS_PROC As String = "frmCheckReconcile_Sub_TADebits_Enter"

31510   With Me
31520     With .frmCheckReconcile_Sub_TADebits.Form
31530       .Header_Arrow_lbl.Visible = True
31540       .Header_Arrow_lbl_line1.Visible = True
31550       .Header_Arrow_lbl_line2.Visible = True
31560     End With
31570   End With

EXITP:
31580   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub frmCheckReconcile_Sub_TADebits_Exit(Cancel As Integer)

31600 On Error GoTo ERRH

        Const THIS_PROC As String = "frmCheckReconcile_Sub_TADebits_Exit"

31610   With Me
31620     With .frmCheckReconcile_Sub_TADebits.Form
31630       .Header_Arrow_lbl.Visible = False
31640       .Header_Arrow_lbl_line1.Visible = False
31650       .Header_Arrow_lbl_line2.Visible = False
31660     End With
31670   End With

EXITP:
31680   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Public Sub cmdUpdate_Click()

31700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngX As Long

31710   With Me
31720     DoCmd.Hourglass True
31730     DoEvents
31740     Set dbs = CurrentDb
31750     With dbs
31760       For lngX = 0& To (lngChks - 1&)
31770         If arr_varChk(C_CHKID, lngX) > 0& Then
31780           Select Case IsNull(arr_varChk(C_DESC, lngX))
                Case True
                  'OK, IT'S NULL NOW, BUT WAS IT?
                  'IS THIS RIGHT? BOTH OF THEM?
                  'IT LOOKS LIKE IT'S CLEARING THEM IN ORDER TO REPLACE THE DESCRIPTION.
                  ' ** Update qryCheckReconcile_08a (Ledger, by specified [jno], with HasDesc),
                  ' ** for HasDesc = True, by specified [chknum], [chkpad].  '## OK
31790             Set qdf = .QueryDefs("qryCheckReconcile_08b")
31800             With qdf.Parameters
31810               ![jno] = arr_varChk(C_JNO, lngX)
31820               ![chknum] = arr_varChk(C_CNUM, lngX)
31830               ![chkpad] = arr_varChk(C_PAID, lngX)
31840             End With
31850             qdf.Execute dbFailOnError
                  ' ** Update qryCheckReconcile_08a (Ledger, by specified [jno], with HasDesc),
                  ' ** for HasDesc = False, by specified [chknum], [chkpad].  '## OK
31860             Set qdf = .QueryDefs("qryCheckReconcile_08c")
31870             With qdf.Parameters
31880               ![jno] = arr_varChk(C_JNO, lngX)
31890               ![chknum] = arr_varChk(C_CNUM, lngX)
31900               ![chkpad] = arr_varChk(C_PAID, lngX)
31910             End With
31920             qdf.Execute dbFailOnError
                  ' ** Update tblCheckReconcile_Staging, by specified
                  ' ** [jno], [chknum], [chkpad], [asgn], [astno1], [astno2].
31930             Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_12a")
31940             With qdf.Parameters
31950               ![jno] = arr_varChk(C_JNO, lngX)
31960               ![chknum] = arr_varChk(C_CNUM, lngX)
31970               ![chkpad] = arr_varChk(C_PAID, lngX)
31980               ![asgn] = arr_varChk(C_ASGN, lngX)
31990               ![astno1] = arr_varChk(C_ASTNO1, lngX)
32000               ![astno2] = arr_varChk(C_ASTNO2, lngX)
32010             End With
32020             qdf.Execute dbFailOnError
32030           Case False
                  ' ** Update Ledger, by specified [jno], [chknum], [chkpad], [jdesc].  '## OK
32040             Set qdf = .QueryDefs("qryCheckReconcile_08d")
32050             With qdf.Parameters
32060               ![jno] = arr_varChk(C_JNO, lngX)
32070               ![chknum] = arr_varChk(C_CNUM, lngX)
32080               ![chkpad] = arr_varChk(C_PAID, lngX)
32090               ![jdesc] = arr_varChk(C_DESC, lngX)
32100             End With
32110             qdf.Execute dbFailOnError
                  ' ** Update tblCheckReconcile_Staging, by specified [jno],
                  ' ** [chknum], [chkpad], [jdesc], [asgn], [astno1], [astno2].
32120             Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_12b")
32130             With qdf.Parameters
32140               ![jno] = arr_varChk(C_JNO, lngX)
32150               ![chknum] = arr_varChk(C_CNUM, lngX)
32160               ![chkpad] = arr_varChk(C_PAID, lngX)
32170               ![jdesc] = arr_varChk(C_DESC, lngX)
32180               ![asgn] = arr_varChk(C_ASGN, lngX)
32190               ![astno1] = arr_varChk(C_ASTNO1, lngX)
32200               ![astno2] = arr_varChk(C_ASTNO2, lngX)
32210             End With
32220             qdf.Execute dbFailOnError
32230           End Select
32240           arr_varChk(C_CHKID, lngX) = CLng(0)
32250         End If
32260         arr_varChk(C_FND, lngX) = True
32270       Next
32280       .Close
32290     End With  ' ** dbs.
32300     .cmdClose.SetFocus
32310     .cmdUpdate.Enabled = False
32320   End With

32330   DoCmd.Hourglass False

EXITP:
32340   Set qdf = Nothing
32350   Set dbs = Nothing
32360   Exit Sub

ERRH:
690     DoCmd.Hourglass False
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub OutChecks_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

32400 On Error GoTo ERRH

        Const THIS_PROC As String = "OutChecks_box_MouseMove"

32410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
32420   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub frmCheckReconcile_Sub_OutChecks_Enter()

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "frmCheckReconcile_Sub_OutChecks_Enter"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

32510   With Me
32520     .OutChecks_Arrow_lbl.Visible = True
32530     .OutChecks_Arrow_lbl_line1.Visible = True
32540     .OutChecks_Arrow_lbl_line2.Visible = True
32550   End With

32560   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
32570   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub frmCheckReconcile_Sub_OutChecks_Exit(Cancel As Integer)

32600 On Error GoTo ERRH

        Const THIS_PROC As String = "frmCheckReconcile_Sub_OutChecks_Exit"

32610   With Me
32620     .OutChecks_Arrow_lbl.Visible = False
32630     .OutChecks_Arrow_lbl_line1.Visible = False
32640     .OutChecks_Arrow_lbl_line2.Visible = False
32650   End With

EXITP:
32660   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub TotalOutstandingChecks_box2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "TotalOutstandingChecks_box2_MouseMove"

32710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
32720   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub chkShowPaid_AfterUpdate()

32800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowPaid_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnChanged As Boolean
        Dim lngX As Long

32810   With Me

32820     DoCmd.Hourglass True
32830     DoEvents

          ' ** This is used by FormRef() to enable appending the tblCheckReconcile_Check records.
32840     .cracct_id_tmp = .cracct_id

32850     Select Case .chkShowPaid
          Case True
            ' ** Show all.

32860       .chkShowPaid_lbl.FontBold = True

32870       Set dbs = CurrentDb
32880       With dbs

              ' ** Update tblCheckReconcile_Staging from array.
32890         Set rst = .OpenRecordset("tblCheckReconcile_Staging", dbOpenDynaset, dbConsistent)
32900         With rst
32910           For lngX = 0& To (lngChks - 1&)
32920             If arr_varChk(C_CHKID, lngX) > 0& Then
                    ' ** Only has number if it was changed.
32930               .FindFirst "[journalno] = " & CStr(arr_varChk(C_JNO, lngX))
32940               If .NoMatch = False Then
32950                 .Edit
32960                 ![CheckNum_new] = arr_varChk(C_CNUM, lngX)
32970                 ![CheckPaid_new] = arr_varChk(C_PAID, lngX)
32980                 ![description] = arr_varChk(C_DESC, lngX)
32990                 ![crstage_assign] = arr_varChk(C_ASGN, lngX)
33000                 ![crstage_asset1] = arr_varChk(C_ASTNO1, lngX)
33010                 ![crstage_asset2] = arr_varChk(C_ASTNO2, lngX)
33020                 ![crstage_datemodified] = Now()
33030                 .Update
33040               End If
33050             End If
33060           Next
33070           .Close
33080         End With

              ' ** Empty tblCheckReconcile_Check.  '## OK
33090         Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_02")
33100         qdf.Execute

              ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check, by specified [datend]; All.  '## OK
33110         Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_25")
33120         With qdf.Parameters
33130           ![datEnd] = Me.DateEnd
33140         End With
33150         qdf.Execute

33160         .Close
33170       End With  ' ** dbs.
33180       DoEvents

33190       .frmCheckReconcile_Sub_OutChecks.Form.Requery
33200       DoEvents
33210       lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
33220       .OutChecksCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
33230       .frmCheckReconcile_Sub_OutChecks.Form.Requery
33240       lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
33250       .OutChecksCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")

33260     Case False
            ' ** Show unpaid.

33270       .chkShowPaid_lbl.FontBold = False
33280       blnChanged = False
33290       For lngX = 0& To (lngChks - 1&)
33300         If arr_varChk(C_CHKID, lngX) > 0& Then
33310           blnChanged = True
33320           Exit For
33330         End If
33340       Next

33350       Select Case blnChanged
            Case True
              ' ** Changes made to original list.

33360         Set dbs = CurrentDb
33370         With dbs

                ' ** Update tblCheckReconcile_Staging from array.
33380           Set rst = .OpenRecordset("tblCheckReconcile_Staging", dbOpenDynaset, dbConsistent)
33390           With rst
33400             For lngX = 0& To (lngChks - 1&)
33410               If arr_varChk(C_CHKID, lngX) > 0& Then
                      ' ** Only has number if it was changed.
33420                 .FindFirst "[journalno] = " & CStr(arr_varChk(C_JNO, lngX))
33430                 If .NoMatch = False Then
33440                   .Edit
33450                   ![CheckNum_new] = arr_varChk(C_CNUM, lngX)
33460                   ![CheckPaid_new] = arr_varChk(C_PAID, lngX)
33470                   ![description] = arr_varChk(C_DESC, lngX)
33480                   ![crstage_assign] = arr_varChk(C_ASGN, lngX)
33490                   ![crstage_asset1] = arr_varChk(C_ASTNO1, lngX)
33500                   ![crstage_asset2] = arr_varChk(C_ASTNO2, lngX)
33510                   ![crstage_datemodified] = Now()
33520                   .Update
33530                 End If
33540               End If
33550             Next
33560             .Close
33570           End With

                ' ** Empty tblCheckReconcile_Check.  '## OK
33580           Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_02")
33590           qdf.Execute

33600           If lngAssets_ThisAcct > 1& Then
33610             If Me.cmbAssets > 0& Then
                    ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                    ' ** for CheckPaid = False, by specified [datend], [astno]; Zeroes/Match.
33620               Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_29")  '####  UNKNOWN  ####
33630               With qdf.Parameters
33640                 ![datEnd] = Me.DateEnd
33650                 ![astno] = Me.cmbAssets
33660               End With
33670             Else
                    ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                    ' ** for CheckPaid = False, by specified [datend]; Zeroes.
33680               Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_30")  '####  UNKNOWN  ####
33690               With qdf.Parameters
33700                 ![datEnd] = Me.DateEnd
33710               End With
33720             End If
33730           Else
                  ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                  ' ** for CheckPaid = False, by specified [datend]; No Criteria.
33740             Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_28")  '## OK
33750             With qdf.Parameters
33760               ![datEnd] = Me.DateEnd
33770             End With
33780           End If
33790           qdf.Execute

33800           .Close
33810         End With  ' ** dbs.

33820       Case False
              ' ** No changes made to original list.

33830         Set dbs = CurrentDb
33840         With dbs

                ' ** Empty tblCheckReconcile_Check.  '## OK
33850           Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_02")
33860           qdf.Execute

33870           If lngAssets_ThisAcct > 1& Then
33880             If Me.cmbAssets > 0& Then
                    ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                    ' ** for CheckPaid = False, by specified [datend], [astno]; Zeroes/Match.
33890               Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_29")  '####  UNKNOWN  ####
33900               With qdf.Parameters
33910                 ![datEnd] = Me.DateEnd
33920                 ![astno] = Me.cmbAssets
33930               End With
33940             Else
                    ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                    ' ** for CheckPaid = False, by specified [datend]; Zeroes.
33950               Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_30")  '####  UNKNOWN  ####
33960               With qdf.Parameters
33970                 ![datEnd] = Me.DateEnd
33980               End With
33990             End If
34000           Else
                  ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                  ' ** for CheckPaid = False, by specified [datend]; No Criteria.  '## OK
34010             Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_28")
34020             With qdf.Parameters
34030               ![datEnd] = Me.DateEnd
34040             End With
34050           End If
34060           qdf.Execute

34070           .Close
34080         End With  ' ** dbs.

34090         .frmCheckReconcile_Sub_OutChecks.Form.Requery
34100         DoEvents
34110         lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
34120         .OutChecksCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")

34130       End Select  ' ** blnChanged.

34140       .frmCheckReconcile_Sub_OutChecks.Form.Requery
34150       lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
34160       .OutChecksCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")

34170     End Select  ' ** chkShowPaid.

34180     If lngRecsCur > 0& Then
34190       .frmCheckReconcile_Sub_OutChecks.SetFocus
34200     End If

34210   End With  ' ** Me.
34220   DoCmd.Hourglass False

EXITP:
34230   Set rst = Nothing
34240   Set qdf = Nothing
34250   Set dbs = Nothing
34260   Exit Sub

ERRH:
690     DoCmd.Hourglass False
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub chkShowPaid_KeyDown(KeyCode As Integer, Shift As Integer)

34300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowPaid_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

34310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
34320   intShiftDown = (Shift And acShiftMask) > 0
34330   intAltDown = (Shift And acAltMask) > 0
34340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
34350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
34370       With Me
34380         intRetVal = 0
34390         If .cmdPreviewReport.Enabled = True Then
34400           .cmdPreviewReport.SetFocus
34410         ElseIf .cmdPrintReport.Enabled = True Then
34420           .cmdPrintReport.SetFocus
34430         ElseIf .cmdClear.Enabled = True Then
34440           .cmdClear.SetFocus
34450         ElseIf .cmdDelete.Enabled = True Then
34460           .cmdDelete.SetFocus
34470         Else
34480           .cmdClose.SetFocus
34490         End If
34500       End With
34510     End Select
34520   End If

        ' ** Shift keys.
34530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
34550       With Me
34560         intRetVal = 0
34570         If .frmCheckReconcile_Sub_OutChecks.Enabled = True Then
34580           .frmCheckReconcile_Sub_OutChecks.SetFocus
34590         ElseIf .DateEnd.Enabled = True Then
34600           .DateEnd.SetFocus
34610         ElseIf .cmbAccounts.Enabled = True Then
34620           .cmbAccounts.SetFocus
34630         Else
34640           .opgCheckingType.SetFocus
34650         End If
34660       End With
34670     End Select
34680   End If

EXITP:
34690   KeyCode = intRetVal
34700   Exit Sub

ERRH:
690     intRetVal = 0
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub chkShowPaid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowPaid_MouseMove"

34810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
34820   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckAll_GotFocus()

34900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckAll_GotFocus"

34910   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
34920   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

35000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckAll_MouseDown"

35010   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
35020   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckAll_Click()

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckAll_Click"

35110   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
35120   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckAll_MouseMove"

35210   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
35220   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

35300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckAll_MouseUp"

35310   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
35320   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckAll_LostFocus()

35400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckAll_LostFocus"

35410   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
35420   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckNone_GotFocus()

35500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckNone_GotFocus"

35510   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
35520   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

35600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckNone_MouseDown"

35610   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
35620   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckNone_Click()

35700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckNone_Click"

35710   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
35720   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckNone_MouseMove"

35810   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
35820   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

35900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckNone_MouseUp"

35910   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
35920   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub cmdCheckNone_LostFocus()

36000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckNone_LostFocus"

36010   Check_Handler_CR THIS_PROC, blnCheckAll_Focus, blnCheckAll_MouseDown, _
          blnCheckNone_Focus, blnCheckNone_MouseDown, Me  ' ** Module Procedure: modCheckReconcile.

EXITP:
36020   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub UnpostedJournalMsg_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

36100 On Error GoTo ERRH

        Const THIS_PROC As String = "UnpostedJournalMsg_lbl_MouseMove"

36110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
36120   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Public Sub FormValidate()

36200 On Error GoTo ERRH

        Const THIS_PROC As String = "FormValidate"

36210   gstrFormQuerySpec = THIS_NAME
36220   With Me
36230     If IsDate(.DateEnd) = False Then
36240       .TABalance = 0
36250       UpdateTATotal_CR Me  ' ** Module Procedure: modCheckReconcile.
36260     Else
36270       If IsNull(.cmbAccounts) = False And IsNull(.cmbAssets) = False Then
36280         gstrFormQuerySpec = THIS_NAME
36290         CalcTABalance  ' ** Procedure: Below.
36300         gstrFormQuerySpec = THIS_NAME
36310         UpdateBSTotal_CR Me, lngAssets_ThisAcct, arr_varAsset_ThisAcct, THIS_PROC ' ** Module Procedure: modCheckReconcile.
36320         gstrFormQuerySpec = THIS_NAME
36330       End If
36340     End If
36350   End With

EXITP:
36360   Exit Sub

ERRH:
690     DoCmd.Hourglass False
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Public Sub CalcTABalance()

36400 On Error GoTo ERRH

        Const THIS_PROC As String = "CalcTABalance"

36410   gstrFormQuerySpec = THIS_NAME

36420   With Me
36430     GetBal_End_CR .cmbAccounts.Column(CBX_ACT_ACTNO), Me  ' ** Module Procedure: modCheckReconcile.
36440 On Error Resume Next
36450     .frmCheckReconcile_Sub_OutChecks.Form.Requery
36460 On Error GoTo ERRH
36470     gstrFormQuerySpec = THIS_NAME
36480     TotOutChks Me, .cmbAccounts.Column(CBX_ACT_ACTNO)  ' ** Module Function: modCheckReconcile.
          '.TotalOutstandingChecks.Requery
          '.TotalOutstandingChecks2.Requery
36490     lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
36500     .OutChecksCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
36510     UpdateTATotal_CR Me  ' ** Module Procedure: modCheckReconcile.
36520   End With
36530   gstrFormQuerySpec = THIS_NAME

EXITP:
36540   gstrFormQuerySpec = THIS_NAME
36550   Exit Sub

ERRH:
690     DoCmd.Hourglass False
700     Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
710     Case Else
720       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
730     End Select
740     Resume EXITP

End Sub

Public Sub UpdateCheckArray(lngJrnlNo As Long, lngCROutChk_ID As Long, blnCheckPaid As Boolean, varCheckNum As Variant, varDescription As Variant)

36600 On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateCheckArray"

        Dim intPos01 As Integer
        Dim strTmp02 As String
        Dim lngX As Long, lngY As Long

36610   If lngCROutChk_ID > 0& Then
36620     For lngX = 0& To (lngChks - 1&)
36630       If arr_varChk(C_JNO, lngX) = lngJrnlNo Then
36640         arr_varChk(C_CHKID, lngX) = lngCROutChk_ID
36650         arr_varChk(C_PAID, lngX) = blnCheckPaid
36660         arr_varChk(C_CNUM, lngX) = varCheckNum
36670         arr_varChk(C_DESC, lngX) = varDescription
36680         If IsNull(varDescription) = False Then
36690           strTmp02 = varDescription
36700           intPos01 = InStr(strTmp02, ASSIGN_MSG)
36710           If intPos01 > 0 Then
36720             strTmp02 = Mid(strTmp02, intPos01)
36730             intPos01 = InStr(strTmp02, ";")
36740             If intPos01 > 0 Then strTmp02 = Left(strTmp02, (intPos01 - 1))
36750             intPos01 = InStr(strTmp02, ":")
36760             If intPos01 > 0 Then
36770               strTmp02 = Trim(Mid(strTmp02, (intPos01 + 1)))
36780               For lngY = 0& To (lngAssets_ThisAcct - 1&)
36790                 If arr_varAsset_ThisAcct(A_CUSIP, lngY) = strTmp02 Then
36800                   arr_varChk(C_ASGN, lngX) = CInt(lngY + 1&)
36810                   Select Case lngY
                        Case 1&
36820                     arr_varChk(C_ASTNO1, lngX) = arr_varAsset_ThisAcct(A_ASTNO, lngY)
36830                     arr_varChk(C_ASTNO2, lngX) = 0&
36840                   Case 2&
36850                     arr_varChk(C_ASTNO1, lngX) = 0&
36860                     arr_varChk(C_ASTNO2, lngX) = arr_varAsset_ThisAcct(A_ASTNO, lngY)
36870                   End Select
36880                   Exit For
36890                 End If
36900               Next
36910             Else
                    ' ** Shouldn't happen!
36920               arr_varChk(C_ASGN, lngX) = 0
36930               arr_varChk(C_ASTNO1, lngX) = 0&
36940               arr_varChk(C_ASTNO2, lngX) = 0&
36950             End If
36960           Else
36970             arr_varChk(C_ASGN, lngX) = 0
36980             arr_varChk(C_ASTNO1, lngX) = 0&
36990             arr_varChk(C_ASTNO2, lngX) = 0&
37000           End If
37010         Else
37020           arr_varChk(C_ASGN, lngX) = 0
37030           arr_varChk(C_ASTNO1, lngX) = 0&
37040           arr_varChk(C_ASTNO2, lngX) = 0&
37050         End If
37060         Exit For
37070       End If
37080     Next
37090   End If

EXITP:
37100   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Public Function RecCnt() As Long

37200 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

37210   lngRetVal = 0&

37220   With Me
37230     Set rst = .RecordsetClone
37240     With rst
37250       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
37260       Else
37270         .MoveLast
37280         lngRetVal = .RecordCount
37290       End If
37300       .Close
37310     End With
37320   End With

EXITP:
37330   Set rst = Nothing
37340   RecCnt = lngRetVal
37350   Exit Function

ERRH:
690     lngRetVal = 0&
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant, Optional varAssetMove As Variant)

37400 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCRAcctID As Long, strAccountNo As String
        Dim strFilter As String
        Dim blnAssetMove As Boolean
        Dim lngRecs As Long

37410   With Me
37420     strFilter = vbNullString: lngAssetNo_Move = -1&: strAccountNo = vbNullString
37430     Select Case IsNull(.cracct_id)
          Case True
37440       lngCRAcctID = 0&
37450     Case False
37460       lngCRAcctID = .cracct_id
37470     End Select
37480     Select Case IsMissing(varAssetMove)
          Case True
37490       blnAssetMove = False
37500     Case False
37510       blnAssetMove = CBool(varAssetMove)
37520       If blnAssetMove = True Then
37530         strAccountNo = .accountno
37540       End If
37550     End Select
37560     Set dbs = CurrentDb
37570     Set qdf = dbs.QueryDefs(.RecordSource)
37580     Set rst = qdf.OpenRecordset
37590     With rst
37600       If .BOF = True And .EOF = True Then
              ' ** How on Earth did they get here?
37610         Beep
37620       Else
37630         .MoveLast
37640         lngRecs = .RecordCount
37650         .MoveFirst
37660         Select Case intWhere
              Case acCmdRecordsGoToFirst
37670           Select Case blnAssetMove
                Case True
37680             If ![accountno] = strAccountNo Then
37690               lngAssetNo_Move = ![assetno]
37700               strFilter = "[accountno] = '" & ![accountno] & "' And [assetno] = " & CStr(lngAssetNo_Move)
37710             Else
37720               Beep
37730             End If
37740           Case False
37750             lngAssetNo_Move = ![assetno]
37760             strFilter = "[accountno] = '" & ![accountno] & "' And [assetno] = " & CStr(lngAssetNo_Move)
37770           End Select
37780         Case acCmdRecordsGoToLast
37790           .MoveLast
37800           Select Case blnAssetMove
                Case True
37810             If ![accountno] = strAccountNo Then
37820               lngAssetNo_Move = ![assetno]
37830               strFilter = "[accountno] = '" & ![accountno] & "' And [assetno] = " & CStr(lngAssetNo_Move)
37840             Else
37850               Beep
37860             End If
37870           Case False
37880             lngAssetNo_Move = ![assetno]
37890             strFilter = "[accountno] = '" & ![accountno] & "' And [assetno] = " & CStr(lngAssetNo_Move)
37900           End Select
37910         Case acCmdRecordsGoToPrevious
37920           If lngCRAcctID > 0& And lngCRAcctID <> ![cracct_id] And lngRecs > 1& Then
37930             .FindFirst "[cracct_id] = " & CStr(lngCRAcctID)
37940             If .NoMatch = False Then
37950               .MovePrevious
37960               Select Case blnAssetMove
                    Case True
37970                 If ![accountno] = strAccountNo Then
37980                   lngAssetNo_Move = ![assetno]
37990                   strFilter = "[accountno] = '" & ![accountno] & "' And [assetno] = " & CStr(lngAssetNo_Move)
38000                 Else
38010                   Beep
38020                 End If
38030               Case False
38040                 lngAssetNo_Move = ![assetno]
38050                 strFilter = "[accountno] = '" & ![accountno] & "' And [assetno] = " & CStr(lngAssetNo_Move)
38060               End Select
38070             Else
38080               Beep
38090             End If
38100           Else
38110             Beep
38120           End If
38130         Case acCmdRecordsGoToNext
38140           If lngCRAcctID > 0& And lngRecs > 1& Then
38150             If lngCRAcctID = ![cracct_id] Then
38160               .MoveNext
38170               Select Case blnAssetMove
                    Case True
38180                 If ![accountno] = strAccountNo Then
38190                   lngAssetNo_Move = ![assetno]
38200                   strFilter = "[accountno] = '" & ![accountno] & "' And [assetno] = " & CStr(lngAssetNo_Move)
38210                 Else
38220                   Beep
38230                 End If
38240               Case False
38250                 lngAssetNo_Move = ![assetno]
38260                 strFilter = "[accountno] = '" & ![accountno] & "' And [assetno] = " & CStr(lngAssetNo_Move)
38270               End Select
38280             Else
38290               .FindFirst "[cracct_id] = " & CStr(lngCRAcctID)
38300               If .NoMatch = False Then
38310                 If .AbsolutePosition < (lngRecs - 1&) Then
38320                   .MoveNext
38330                   Select Case blnAssetMove
                        Case True
38340                     If ![accountno] = strAccountNo Then
38350                       lngAssetNo_Move = ![assetno]
38360                       strFilter = "[accountno] = '" & ![accountno] & "' And [assetno] = " & CStr(lngAssetNo_Move)
38370                     Else
38380                       Beep
38390                     End If
38400                   Case False
38410                     lngAssetNo_Move = ![assetno]
38420                     strFilter = "[accountno] = '" & ![accountno] & "' And [assetno] = " & CStr(lngAssetNo_Move)
38430                   End Select
38440                 Else
38450                   Beep
38460                 End If
38470               Else
38480                 Beep
38490               End If
38500             End If
38510           Else
38520             Beep
38530           End If
38540         Case Else
38550           If IsMissing(varID) = False And IsMissing(varAssetMove) = True Then
38560             If varID <> ![cracct_id] And lngRecs > 1& Then
38570               .FindFirst "[cracct_id] = " & CStr(varID)
38580               If .NoMatch = False Then
38590                 lngAssetNo_Move = ![assetno]
38600                 strFilter = "[accountno] = '" & ![accountno] & "' And [assetno] = " & CStr(lngAssetNo_Move)
38610               Else
38620                 Beep
38630               End If
38640             End If
38650           End If
38660         End Select
38670       End If
38680       .Close
38690     End With
38700     dbs.Close
38710     If strFilter <> vbNullString Then
38720       .Filter = strFilter
38730       .FilterOn = True
38740       cmbAccounts_AfterUpdate  ' ** Procedure: Above
38750     End If
38760   End With

EXITP:
38770   Set rst = Nothing
38780   Set qdf = Nothing
38790   Set dbs = Nothing
38800   Exit Sub

ERRH:
690     Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
700     Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
710     Case Else
720       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
730     End Select
740     Resume EXITP

End Sub

Public Function AssetNumGet(intMode As Integer, lngNum As Long) As Variant

38900 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetNumGet"

        Dim lngX As Long
        Dim varRetVal As Variant

38910   Select Case intMode
        Case 1, 3
38920     varRetVal = 0&
38930   Case 2
38940     varRetVal = vbNullString
38950   End Select

        ' ** This should only be the assets assigned to this account; one or two.
38960   If intMode = 3 Then
38970     varRetVal = lngAssets_ThisAcct
38980   Else
38990     For lngX = 0& To (lngAssets_ThisAcct - 1&)
39000       If (lngX + 1&) = lngNum Then
39010         Select Case intMode
              Case 1
39020           varRetVal = arr_varAsset_ThisAcct(A_ASTNO, lngX)
39030         Case 2
39040           varRetVal = arr_varAsset_ThisAcct(A_CUSIP, lngX)
39050         End Select
39060         Exit For
39070       End If
39080     Next
39090   End If

EXITP:
39100   AssetNumGet = varRetVal
39110   Exit Function

ERRH:
690     varRetVal = 0&
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Function

Public Function AssetsThisArray_Get() As Variant

39200 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetsThisArray_Get"

        Dim arr_varRetVal As Variant

39210   arr_varRetVal = arr_varAsset_ThisAcct

EXITP:
39220   AssetsThisArray_Get = arr_varRetVal
39230   Exit Function

ERRH:
690     arr_varRetVal = RET_ERR
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Function
