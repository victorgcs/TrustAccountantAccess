VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTransaction_Audit_Sub_Criteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmTransaction_Audit_Sub_Criteria"

'VGC 10/05/2017: CHANGES!

' #########################
' ## Use VBA_RenumErrh().  39560
' #########################

' ** Shortcut Alt keys responsive from this form:
' **   AccountNo:        A {cmbAccounts}
' **   Acct Number:      B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose on frmTransaction_Audit}
' **   Date Posted:      D {TransDateStart}
' **   Excel:            E {cmdExcel on frmTransaction_Audit}
' **   Form View:        F {opgView_optForm on frmTransaction_Audit}
' **   Include Archive:  H {chkIncludeArchive on frmTransaction_Audit}
' **   Asset CUSIP:      I {opgAssetSource_optCUSIP}
' **   JournalNo:        J {journalno on frmTransaction_Audit_Sub, frmTransaction_Audit_Sub_ds}
' **   Show Criteria:    K {cmdWidenToCriteria on frmTransaction_Audit}
' **   Select All:       L {cmdSelectAll on frmTransaction_Audit}
' **   Acct Name:        M {opgAccountSource_optName}
' **   Select None:      N {cmdSelectNone on frmTransaction_Audit}
' **   Word:             O {cmdWord on frmTransaction_Audit}
' **   Print:            P {cmdPrint on frmTransaction_Audit}
' **   Clear All:        R {cmdClearAll on frmTransaction_Audit}
' **   Assets:           S {cmbAssets}
' **   Datasheet View:   T {opgView_optDatasheet on frmTransaction_Audit}
' **   View Fields:      V {ckgFlds_cmd on frmTransaction_Audit}
' **   Preview:          W {cmdPreview on frmTransaction_Audit}
' **   Exit:             X {cmdClose on frmTransaction_Audit}
' **   Asset Type:       Y {opgAssetSource_optType}
' **   Size/Position:    Z {chkSaveSizePos on frmTransaction_Audit}
' **   JournalType1:     1 {cmbJournalType1}
' **   JournalType2:     2 {cmbJournalType2}
' **   JournalType3:     3 {cmbJournalType3}
'G  Q  U

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker Posting - From:  F5  {cmdCalendar1}
' **   Date Picker Posting - To:    F6  {cmdCalendar2}
' **   Date Picker Trade - From:    F7  {cmdCalendar3}
' **   Date Picker Trade - To:      F8  {cmdCalendar4}
' **   Date Picker Purchase - From: F9  {cmdCalendar5}
' **   Date Picker Purchase - To:   F10 {cmdCalendar6}
' **   Date Picker Posted - From:   F11 {cmdCalendar7}
' **   Date Picker Posted - To:     F12 {cmdCalendar8}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf on frmTransaction_Audit}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown on frmTransaction_Audit}

' ** Shortcut Alt-Shift keys responsive from this form:
' **   Include Hidden:   C {opgHidden_optInclude}
' **   Tax Deduction:    D {chkTaxcodeType_Deduction}
' **   Exclude Hidden:   E {opgHidden_optExclude}
' **   Tax Income:       I {chkTaxcodeType_Income}
' **   Asset Name:       M {opgAssetSource_optName}
' **   Narrow Font:      N {chkNarrowFont on frmTransaction_Audit}
' **   Only Hidden:      O {opgHidden_optOnly}
' **   Shareface Total:  T {chkIncludeSharefaceTot on frmTransaction_Audit}

' ** Shortcut Ctrl-Alt keys responsive from this form:
' **   Inc/Exp Expense:  E {chkRevcodeType_Expense}
' **   Inc/Exp Income:   I {chkRevcodeType_Income}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

' ** Filter constants: frmTransaction_Audit_Sub.
Private Const ANDF         As String = " And "  ' ** Filter 'And'.
Private Const ORF          As String = " Or "  ' ** Filter 'Or'.
Private Const JRNL_NUM     As String = "[journalno] = "
Private Const JRNL_TYPE    As String = "[journaltype] = '"
Private Const TRANS_START  As String = "[transdate] >= #"
Private Const TRANS_END    As String = "[transdate] <= #"
Private Const ACCT_NUM     As String = "[accountno] = '"
Private Const ASSET_NUM    As String = "[assetno] = "
Private Const CURR_NUM     As String = "[curr_id] = "
Private Const ASSET_START  As String = "[assetdate] >= #"
Private Const ASSET_END    As String = "[assetdate] <= #"
Private Const PURCH_START  As String = "[PurchaseDate] >= #"
Private Const PURCH_END    As String = "[PurchaseDate] <= #"
Private Const COMM_DESC    As String = "[ledger_description] Like '*"
Private Const RECUR_ITEM   As String = "[RecurringItem] Like '*"
Private Const REV_CODE     As String = "[revcode_ID] = "
Private Const TAX_CODE     As String = "[taxcode] = "
Private Const LOC_NUM      As String = "[Location_ID] = "
Private Const CHK_NUM      As String = "[CheckNum] = "
Private Const CHK_NUM1     As String = "[CheckNum] >= "
Private Const CHK_NUM2     As String = "[CheckNum] <= "
Private Const JRNL_USER    As String = "[journal_USER] = '"
Private Const POSTED_START As String = "[posted] >= #"
Private Const POSTED_END   As String = "[posted] <= #"
Private Const HIDDEN_TRX1  As String = "[ledger_HIDDEN] = True"
Private Const HIDDEN_TRX2  As String = "[ledger_HIDDEN] = False"

' ** Array: arr_varFilt(), arr_varFilt_ds().
Private lngFilts As Long, arr_varFilt As Variant
Private lngFilts_ds As Long, arr_varFilt_ds As Variant
'Private Const F_ELEMS As Integer = 13  ' ** Array's first-element UBound().
'Private Const F_IDX   As Integer = 0
'Private Const F_NAM   As Integer = 1
Private Const F_CONST As Integer = 2
'Private Const F_CTL   As Integer = 3
'Private Const F_CLBL  As Integer = 4
'Private Const F_FLD   As Integer = 5
'Private Const F_FLBL  As Integer = 6
'Private Const F_CTL2  As Integer = 7
'Private Const F_CLBL2 As Integer = 8
'Private Const F_FLD2  As Integer = 9
'Private Const F_FLBL2 As Integer = 10
'Private Const F_CLBL3 As Integer = 11
'Private Const F_FLD3  As Integer = 12
'Private Const F_FLBL3 As Integer = 13

' ** Array: arr_varCal().
Private lngCals As Long, arr_varCal() As Variant
Private Const C_ELEMS As Integer = 4  ' ** Array's first-element UBound().
Private Const C_CNAM  As Integer = 0
Private Const C_FOCUS As Integer = 1
Private Const C_DOWN  As Integer = 2
Private Const C_ABLE  As Integer = 3
Private Const C_FLD   As Integer = 4

Private rstAll1 As DAO.Recordset, rstAll2 As DAO.Recordset

Private strFilter01 As String, strFilter02 As String, dblFilterRecs As Double
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmTransaction_Audit" Then

            ' ** These are the ones added later.
160         .cmbCurrencies.Enabled = False
170         .cmbCurrencies.BackStyle = acBackStyleTransparent
180         .cmbCurrencies.BorderColor = WIN_CLR_DISR
190         .cmbLocations.Enabled = False
200         .cmbLocations.BackStyle = acBackStyleTransparent
210         .cmbLocations.BorderColor = WIN_CLR_DISR
220         .CheckNum.Enabled = False
230         .CheckNum.BackStyle = acBackStyleTransparent
240         .CheckNum.BorderColor = WIN_CLR_DISR

250       Else
260         Cancel = -1
270       End If
280     End With

290     If Cancel = -1 Then
300       Beep
310       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
320     End If

EXITP:
330     Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub Form_Load()
' ** Create an instance of the Calendar class.
' ** Copyright Lebans Holdings 1999 Ltd.

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me

420       Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
430       DoEvents

440       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
450       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
460       clsMonthClass.PositionAtCursor = True

470       Calendar_Handler THIS_PROC  ' ** Procedure: Below.

480     End With

EXITP:
490     Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim ctl As Access.Control
        Dim intRetVal As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Date Picker Posting - From:  F5  {cmdCalendar1}
        ' **   Date Picker Posting - To:    F6  {cmdCalendar2}
        ' **   Date Picker Trade - From:    F7  {cmdCalendar3}
        ' **   Date Picker Trade - To:      F8  {cmdCalendar4}
        ' **   Date Picker Purchase - From: F9  {cmdCalendar5}
        ' **   Date Picker Purchase - To:   F10 {cmdCalendar6}
        ' **   Date Picker Posted - From:   F11 {cmdCalendar7}
        ' **   Date Picker Posted - To:     F12 {cmdCalendar8}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Date Posted:      D {TransDateStart}
        ' **   Include Archive:  H {chkIncludeArchive on frmTransaction_Audit}
        ' **   JournalNo:        J {journalno on frmTransaction_Audit_Sub, frmTransaction_Audit_Sub_ds}
        ' **   Show Criteria:    K {cmdWidenToCriteria on frmTransaction_Audit}
        ' **   Clear All:        R {cmdClearAll on frmTransaction_Audit}
        ' **   Exit:             X {cmdClose on frmTransaction_Audit}

        ' ** Shortcut Alt-Shift keys to other forms and functionality:
        ' **   Include Hidden:   C {opgHidden_optInclude}
        ' **   Tax Deduction:    D {chkTaxcodeType_Deduction}
        ' **   Exclude Hidden:   E {opgHidden_optExclude}
        ' **   Tax Income:       I {chkTaxcodeType_Income}
        ' **   Asset Name:       M {opgAssetSource_optName}
        ' **   Narrow Font:      N {chkNarrowFont on frmTransaction_Audit}
        ' **   Only Hidden:      O {opgHidden_optOnly}
        ' **   Shareface Total:  T {chkIncludeSharefaceTot on frmTransaction_Audit}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf on frmTransaction_Audit}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown on frmTransaction_Audit}

        ' ** Shortcut Ctrl-Alt keys to other forms and functionality:
        ' **   Inc/Exp Expense:  E {chkRevcodeType_Expense}
        ' **   Inc/Exp Income:   I {chkRevcodeType_Income}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyF5
            ' **   Date Picker Posting - From:  F5  {cmdCalendar1}
570         With Me
580           intRetVal = 0
590           If cmdCalendar1.Enabled = True Then
600             .cmdCalendar1.SetFocus
610             cmdCalendar1_Click  ' ** Procedure: Below.
620           End If
630         End With
640       Case vbKeyF6
            ' **   Date Picker Posting - To:    F6  {cmdCalendar2}
650         With Me
660           intRetVal = 0
670           If cmdCalendar2.Enabled = True Then
680             .cmdCalendar2.SetFocus
690             cmdCalendar2_Click  ' ** Procedure: Below.
700           End If
710         End With
720       Case vbKeyF7
            ' **   Date Picker Trade - From:    F7  {cmdCalendar3}
730         With Me
740           intRetVal = 0
750           If cmdCalendar3.Enabled = True Then
760             .cmdCalendar3.SetFocus
770             cmdCalendar3_Click  ' ** Procedure: Below.
780           End If
790         End With
800       Case vbKeyF8
            ' **   Date Picker Trade - To:      F8  {cmdCalendar4}
810         With Me
820           intRetVal = 0
830           If cmdCalendar4.Enabled = True Then
840             .cmdCalendar4.SetFocus
850             cmdCalendar4_Click  ' ** Procedure: Below.
860           End If
870         End With
880       Case vbKeyF9
            ' **   Date Picker Purchase - From: F9  {cmdCalendar5}
890         With Me
900           intRetVal = 0
910           If cmdCalendar5.Enabled = True Then
920             .cmdCalendar5.SetFocus
930             cmdCalendar5_Click  ' ** Procedure: Below.
940           End If
950         End With
960       Case vbKeyF10
            ' **   Date Picker Purchase - To:   F10 {cmdCalendar6}
970         With Me
980           intRetVal = 0
990           If cmdCalendar6.Enabled = True Then
1000            .cmdCalendar6.SetFocus
1010            cmdCalendar6_Click  ' ** Procedure: Below.
1020          End If
1030        End With
1040      Case vbKeyF11
            ' **   Date Picker Posted - From:   F11 {cmdCalendar7}
1050        With Me
1060          intRetVal = 0
1070          If cmdCalendar7.Enabled = True Then
1080            .cmdCalendar7.SetFocus
1090            cmdCalendar7_Click  ' ** Procedure: Below.
1100          End If
1110        End With
1120      Case vbKeyF12
            ' **   Date Picker Posted - To:     F12 {cmdCalendar8}
1130        With Me
1140          intRetVal = 0
1150          If cmdCalendar8.Enabled = True Then
1160            .cmdCalendar8.SetFocus
1170            cmdCalendar8_Click  ' ** Procedure: Below.
1180          End If
1190        End With
1200      End Select
1210    End If

        ' ** Alt keys.
1220    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1230      Select Case intRetVal
          Case vbKeyD
1240        With Me
1250          intRetVal = 0
1260          .TransDateStart.SetFocus
1270        End With
1280      Case vbKeyH
1290        With Me
1300          intRetVal = 0
1310          If .Parent.chkIncludeArchive.Enabled = True Then
1320            DoCmd.SelectObject acForm, .Parent.Name, False
1330            .Parent.chkIncludeArchive.SetFocus
1340            .Parent.chkIncludeArchive = (Not .Parent.chkIncludeArchive)
1350            .Parent.chkIncludeArchive_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit.
1360          Else
1370            Beep
1380          End If
1390        End With
1400      Case vbKeyJ
1410        With Me
1420          intRetVal = 0
1430          DoCmd.SelectObject acForm, .Parent.Name, False
1440          Select Case .Parent.opgView
              Case .Parent.opgView_optForm.OptionValue
1450            If .Parent.frmTransaction_Audit_Sub.Form.journalno.Visible = True And _
                    .Parent.frmTransaction_Audit_Sub.Form.journalno.Enabled = True Then
1460              .Parent.frmTransaction_Audit_Sub.Form.journalno.SetFocus
1470            Else
1480              Beep
1490            End If
1500          Case .Parent.opgView_optDatasheet.OptionValue
1510            If .Parent.frmTransaction_Audit_Sub_ds.Form.journalno.Visible = True And _
                    .Parent.frmTransaction_Audit_Sub_ds.Form.journalno.Enabled = True Then
1520              .Parent.frmTransaction_Audit_Sub_ds.Form.[Journal Number].SetFocus
1530            Else
1540              Beep
1550            End If
1560          End Select
1570        End With
1580      Case vbKeyK
1590        With Me
1600          intRetVal = 0
1610          If .Parent.cmdWidenToCriteria.Enabled = True Then
1620            DoCmd.SelectObject acForm, .Parent.Name, False
1630            .Parent.cmdWidenToCriteria.SetFocus
1640            .Parent.cmdWidenToCriteria_Click  ' ** Form Procedure: frmTransaction_Audit.
1650          Else
1660            Beep
1670          End If
1680        End With
1690      Case vbKeyR
1700        With Me
1710          If dblFilterRecs = 0# Then
1720            DoCmd.SelectObject acForm, .Parent.Name, False
1730            .Parent.cmdClearAll.SetFocus
1740            .Parent.cmdClearAll_Click  ' ** Form Procedure: frmTransaction_Audit.
1750          End If
1760        End With
1770      Case vbKeyX
1780        With Me
1790          intRetVal = 0
1800          .Parent.cmdClose_Click  ' ** Form Procedure: frmTransaction_Audit.
1810        End With
1820      End Select
1830    End If

        ' ** Alt-Shift keys.
1840    If (Not intCtrlDown) And intAltDown And intShiftDown Then
1850      Select Case intRetVal
          Case vbKeyC
1860        With Me
1870          intRetVal = 0
1880          .opgHidden.SetFocus
1890          .opgHidden = .opgHidden_optInclude.OptionValue
1900          opgHidden_AfterUpdate  ' ** Procedure: Below.
1910        End With
1920      Case vbKeyD
1930        With Me
1940          intRetVal = 0
1950          .chkTaxcodeType_Deduction = (Not .chkTaxcodeType_Deduction)
1960          .chkTaxcodeType_Deduction_AfterUpdate  ' ** Procedure: Below.
1970        End With
1980      Case vbKeyE
1990        With Me
2000          intRetVal = 0
2010          .opgHidden.SetFocus
2020          .opgHidden = .opgHidden_optExclude.OptionValue
2030          opgHidden_AfterUpdate  ' ** Procedure: Below.
2040        End With
2050      Case vbKeyI
2060        With Me
2070          intRetVal = 0
2080          .chkTaxcodeType_Income = (Not .chkTaxcodeType_Income)
2090          chkTaxcodeType_Income_AfterUpdate  ' ** Procedure: Below.
2100        End With
2110      Case vbKeyM
2120        With Me
2130          intRetVal = 0
2140          .opgAssetSource = .opgAssetSource_optName.OptionValue
2150          .opgAssetSource.SetFocus
2160          opgAssetSource_AfterUpdate  ' ** Procedure: Below.
2170        End With
2180      Case vbKeyN
2190        With Me
2200          intRetVal = 0
2210          DoCmd.SelectObject acForm, .Parent.Name, False
2220          .Parent.chkNarrowFont.SetFocus
2230          .Parent.chkNarrowFont = (Not .Parent.chkNarrowFont)
2240          .Parent.chkNarrowFont_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit.
2250        End With
2260      Case vbKeyO
2270        With Me
2280          intRetVal = 0
2290          .opgHidden.SetFocus
2300          .opgHidden = .opgHidden_optOnly.OptionValue
2310          opgHidden_AfterUpdate  ' ** Procedure: Below.
2320        End With
2330      Case vbKeyT
2340        With Me
2350          intRetVal = 0
2360          DoCmd.SelectObject acForm, .Parent.Name, False
2370          .Parent.chkIncludeSharefaceTot.SetFocus
2380          .Parent.chkIncludeSharefaceTot = (Not .Parent.chkIncludeSharefaceTot)
2390          .Parent.chkIncludeSharefaceTot_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit.
2400        End With
2410      End Select
2420    End If

        ' ** Ctrl keys.
2430    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2440      Select Case intRetVal
          Case vbKeyG
2450        With Me
2460          intRetVal = 0
2470          DoCmd.SelectObject acForm, .Parent.Name, False
2480          .Parent.chkPageOf.SetFocus
2490          .Parent.chkPageOf = (Not .Parent.chkPageOf)
2500          .Parent.chkPageOf_AfterUpdate  ' ** Procedure: frmTransaction_Audit.
2510        End With
            'Case vbKeyTab, vbKeyReturn  ' ** Need this for moving to and from calendar buttons.
            '  With Me
            '    intRetVal = 0
            '    DoCmd.SelectObject acForm, .Parent.Name, False
            '    Select Case .Parent.opgView
            '    Case .Parent.opgView_optForm.OptionValue
            '      .Parent.frmTransaction_Audit_Sub.SetFocus
            '    Case .Parent.opgView_optDatasheet.OptionValue
            '      .Parent.frmTransaction_Audit_Sub_ds.SetFocus
            '    End Select
            '  End With
2520      End Select
2530    End If

        ' ** Ctrl-Shift keys.
2540    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2550      Select Case intRetVal
          Case vbKeyF
2560        With Me
2570          intRetVal = 0
2580          DoCmd.SelectObject acForm, .Parent.Name, False
2590          .Parent.FocusHolder.SetFocus
2600        End With
2610      Case vbKeyT
2620        intRetVal = 0
2630        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
            'Case vbKeyTab, vbKeyReturn  ' ** Need this for moving to and from calendar buttons.
            '  With Me
            '    intRetVal = 0
            '    DoCmd.SelectObject acForm, .Parent.Name, False
            '    .Parent.opgView.SetFocus
            '  End With
2640      End Select
2650    End If

        ' ** Ctrl-Alt keys.
2660    If intCtrlDown And intAltDown And (Not intShiftDown) Then
2670      Select Case intRetVal
          Case vbKeyE
2680        With Me
2690          intRetVal = 0
2700          .chkRevcodeType_Expense.SetFocus
2710          .chkRevcodeType_Expense = (Not .chkRevcodeType_Expense)
2720          chkRevcodeType_Expense_AfterUpdate  ' ** Procedure: Below.
2730        End With
2740      Case vbKeyI
2750        With Me
2760          intRetVal = 0
2770          .chkRevcodeType_Income.SetFocus
2780          .chkRevcodeType_Income = (Not .chkRevcodeType_Income)
2790          chkRevcodeType_Income_AfterUpdate  ' ** Procedure: Below.
2800        End With
2810      End Select
2820    End If

EXITP:
2830    Set ctl = Nothing
2840    KeyCode = intRetVal
2850    Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
360     Case Else
370       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
380     End Select
390     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.
' ** VGC 04/13/08: Calendar window seems to be Modal, preventing the
' **               user from doing anything else till it's closed.

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2910    If Not clsMonthClass Is Nothing Then
2920      If clsMonthClass.IsCalendar Then
2930        Cancel = -1
2940      Else
2950        Set clsMonthClass = Nothing
2960      End If
2970    End If

EXITP:
2980    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3010    DetailMouse_Crit_TA lngCals, arr_varCal, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
3020    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub FilterBy_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "FilterBy_box_MouseMove"

3110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3120    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub journalno_AfterUpdate()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_AfterUpdate"

        Dim frm As Access.Form

3210    With Me
3220      If IsNull(.journalno) = False Then
3230        DoCmd.Hourglass True
3240        DoEvents
3250        Set frm = .Parent
3260        strFilter01 = JRNL_NUM & CStr(.journalno)
3270        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
3280        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
3290        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
3300        strFilter02 = JRNL_NUM & CStr(.journalno)
3310        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
3320        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll2  ' ** Form Function: frmTransaction_Audit_Sub_ds.
3330        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
3340        frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
3350        DoCmd.Hourglass False
3360        DoEvents
3370      End If
3380    End With

EXITP:
3390    Set frm = Nothing
3400    Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub journalno_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3520    intShiftDown = (Shift And acShiftMask) > 0
3530    intAltDown = (Shift And acAltMask) > 0
3540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3570        With Me
3580          intRetVal = 0
3590          If .cmbJournalType1.Enabled = True Then
3600            .cmbJournalType1.SetFocus
3610          ElseIf .TransDateStart.Enabled = True Then
3620            .TransDateStart.SetFocus
3630          ElseIf .cmbAccounts.Enabled = True Then
3640            .cmbAccounts.SetFocus
3650          ElseIf .cmbAssets.Enabled = True Then
3660            .cmbAssets.SetFocus
3670          ElseIf .cmbCurrencies.Enabled = True Then
3680            .cmbCurrencies.SetFocus
3690          ElseIf .AssetDateStart.Enabled = True Then
3700            .AssetDateStart.SetFocus
3710          ElseIf .PurchaseDateStart.Enabled = True Then
3720            .PurchaseDateStart.SetFocus
3730          ElseIf .ledger_description.Enabled = True Then
3740            .ledger_description.SetFocus
3750          ElseIf .cmbRecurringItems.Enabled = True Then
3760            .cmbRecurringItems.SetFocus
3770          ElseIf .cmbRevenueCodes.Enabled = True Then
3780            .cmbRevenueCodes.SetFocus
3790          ElseIf .chkRevcodeType_Income.Enabled = True Then
3800            .chkRevcodeType_Income.SetFocus
3810          ElseIf .cmbTaxCodes.Enabled = True Then
3820            .cmbTaxCodes.SetFocus
3830          ElseIf .chkTaxcodeType_Income.Enabled = True Then
3840            .chkTaxcodeType_Income.SetFocus
3850          ElseIf .CheckNum.Enabled = True Then
3860            .CheckNum.SetFocus
3870          ElseIf .cmbUsers.Enabled = True Then
3880            .cmbUsers.SetFocus
3890          ElseIf .PostedDateStart.Enabled = True Then
3900            .PostedDateStart.SetFocus
3910          ElseIf .opgHidden.Enabled = True Then
3920            .opgHidden.SetFocus
3930          Else
3940            DoCmd.SelectObject acForm, .Parent.Name, False
3950            Select Case .Parent.opgView
                Case .Parent.opgView_optForm.OptionValue
3960              .Parent.frmTransaction_Audit_Sub.SetFocus
3970            Case .Parent.opgView_optDatasheet.OptionValue
3980              .Parent.frmTransaction_Audit_Sub_ds.SetFocus
3990            End Select
4000          End If
4010        End With
4020      End Select
4030    End If

        ' ** Shift keys.
4040    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4050      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4060        With Me
4070          intRetVal = 0
4080          DoCmd.SelectObject acForm, .Parent.Name, False
4090          .Parent.cmdWidenToCriteria.SetFocus
4100        End With
4110      End Select
4120    End If

EXITP:
4130    KeyCode = intRetVal
4140    Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmbJournalType_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalType_lbl_MouseMove"

4210    With Me
4220      .Parent.Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.
4230    End With

EXITP:
4240    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmbJournalType1_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalType1_AfterUpdate"

4310    JType_After_TA THIS_PROC, strFilter01, strFilter02, dblFilterRecs, rstAll1, Me  ' ** Module Procedure: modTrasnsactionAuditFuncs2.

EXITP:
4320    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmbJournalType1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalType1_MouseMove"

4410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4420    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmbJournalType2_AfterUpdate()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalType2_AfterUpdate"

4510    JType_After_TA THIS_PROC, strFilter01, strFilter02, dblFilterRecs, rstAll1, Me  ' ** Module Procedure: modTrasnsactionAuditFuncs2.

EXITP:
4520    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmbJournalType2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalType2_MouseMove"

4610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4620    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmbJournalType3_AfterUpdate()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalType3_AfterUpdate"

4710    JType_After_TA THIS_PROC, strFilter01, strFilter02, dblFilterRecs, rstAll1, Me  ' ** Module Procedure: modTrasnsactionAuditFuncs2.

EXITP:
4720    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmbJournalType3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalType3_MouseMove"

4810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4820    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub TransDateStart_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStart_lbl2_MouseMove"

4910    With Me
4920      .Parent.Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.
4930    End With

EXITP:
4940    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub TransDateStart_BeforeUpdate(Cancel As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStart_BeforeUpdate"

5010    With Me
5020      If .TransDateStart > Date Then
5030        MsgBox "Please enter a date prior to today.", vbInformation + vbOKOnly, "Invalid Date"
5040        Cancel = -1
5050      End If
5060    End With

EXITP:
5070    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub TransDateStart_AfterUpdate()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStart_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

5110    With Me
5120      DoCmd.Hourglass True
5130      DoEvents
5140      Set frm = .Parent
5150      If IsNull(.TransDateStart) = False Then
5160        If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
5170          strFilter01 = TRANS_START & .TransDateStart & "#"
5180        Else
              ' ** There are clauses present.

5190          FilterRec_GetArr  ' ** Procedure: Below.

5200          intPos01 = InStr(strFilter01, TRANS_START)
5210          If intPos01 = 0& Then
                ' ** This clause isn't present.
5220            intPos03 = 0&: intPos04 = 0&
5230            For lngX = (lngFilts - 1&) To 0& Step -1&
5240              If arr_varFilt(F_CONST, lngX) = TRANS_START Then
5250                intPos04 = -1
5260              ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
5270                intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
5280                If intPos03 > 0 Then
5290                  intPos04 = 0&
5300                  Exit For
5310                End If
5320              End If
5330            Next
5340            If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
5350              strFilter01 = TRANS_START & .TransDateStart & "#" & ANDF & strFilter01
5360            Else
                  ' ** There's a clause before this one.
5370              intPos02 = InStr(intPos03, strFilter01, ANDF)
5380              If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
5390                strFilter01 = strFilter01 & ANDF & TRANS_START & .TransDateStart & "#"
5400              Else
                    ' ** Add this clause to the middle of the filter.
5410                strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & TRANS_START & .TransDateStart & "#" & Mid(strFilter01, intPos02)
5420              End If
5430            End If
5440          Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
5450            intPos02 = InStr((intPos01 + Len(TRANS_START) + 1), strFilter01, "#")  ' ** Find the closing paren.
5460            If intPos02 > 0 Then
5470              strFilter01 = Left(strFilter01, (intPos01 - 1)) & TRANS_START & .TransDateStart & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
5480            Else
5490              strFilter01 = Left(strFilter01, (intPos01 - 1)) & TRANS_START & .TransDateStart & "#"
5500            End If
5510          End If
5520        End If
5530        strFilter02 = strFilter01
5540        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
5550        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
5560        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
5570        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
5580        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
5590        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
5600      Else
5610        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
5620        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem TRANS_START  ' ** Form Procedure: frmTransaction_Audit_Sub.
5630        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
5640        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub.
5650        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
5660        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem TRANS_START  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
5670        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
5680        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
5690      End If
5700      frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
5710      DoEvents
          '.TransDateEnd.SetFocus
5720      DoCmd.Hourglass False
5730      DoEvents
5740    End With

EXITP:
5750    Set frm = Nothing
5760    Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub TransDateStart_KeyDown(KeyCode As Integer, Shift As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5820    intShiftDown = (Shift And acShiftMask) > 0
5830    intAltDown = (Shift And acAltMask) > 0
5840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5870        With Me
5880          intRetVal = 0
5890          .TransDateEnd.SetFocus
5900        End With
5910      End Select
5920    End If

        ' ** Shift keys.
5930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5950        With Me
5960          intRetVal = 0
5970          If .cmbJournalType3.Enabled = True Then
5980            .cmbJournalType3.SetFocus
5990          ElseIf .cmbJournalType2.Enabled = True Then
6000            .cmbJournalType2.SetFocus
6010          ElseIf .cmbJournalType1.Enabled = True Then
6020            .cmbJournalType1.SetFocus
6030          Else

6040          End If
6050        End With
6060      End Select
6070    End If

        ' ** Ctrl keys.
6080    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
6090      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6100        With Me
6110          intRetVal = 0
6120          .cmdCalendar1.SetFocus
6130        End With
6140      End Select
6150    End If

EXITP:
6160    KeyCode = intRetVal
6170    Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub TransDateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStart_MouseMove"

6210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6220    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

6310    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
6320    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

6410    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
6420    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmdCalendar1_Click()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

6510    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
6520    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

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
6690          .TransDateEnd.SetFocus
6700        End With
6710      End Select
6720    End If

        ' ** Shift keys.
6730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6740      Select Case intRetVal
          Case vbKeyTab
6750        With Me
6760          intRetVal = 0
6770          .TransDateStart.SetFocus
6780        End With
6790      End Select
6800    End If

EXITP:
6810    KeyCode = intRetVal
6820    Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

6910    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
6920    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

7010    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
7020    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

7110    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
7120    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub TransDateEnd_BeforeUpdate(Cancel As Integer)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateEnd_BeforeUpdate"

7210    With Me
7220      If IsNull(.TransDateStart) = False Then
7230        If .TransDateEnd < .TransDateStart Then
7240          MsgBox "Please enter a date the same as or after the From date.", vbInformation + vbOKOnly, "Invalid Date"
7250          Cancel = -1
7260        End If
7270      End If
7280    End With

EXITP:
7290    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub TransDateEnd_AfterUpdate()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateEnd_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

7310    With Me
7320      DoCmd.Hourglass True
7330      DoEvents
7340      Set frm = .Parent
7350      If IsNull(.TransDateEnd) = False Then
7360        If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
7370          strFilter01 = TRANS_END & .TransDateEnd & "#"
7380        Else
              ' ** There are clauses present.

7390          FilterRec_GetArr  ' ** Procedure: Below.

7400          intPos01 = InStr(strFilter01, TRANS_END)
7410          If intPos01 = 0& Then
                ' ** This clause isn't present.
7420            intPos03 = 0&: intPos04 = 0&
7430            For lngX = (lngFilts - 1&) To 0& Step -1&
7440              If arr_varFilt(F_CONST, lngX) = TRANS_END Then
7450                intPos04 = -1
7460              ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
7470                intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
7480                If intPos03 > 0 Then
7490                  intPos04 = 0&
7500                  Exit For
7510                End If
7520              End If
7530            Next
7540            If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
7550              strFilter01 = TRANS_END & .TransDateEnd & "#" & ANDF & strFilter01
7560            Else
                  ' ** There's a clause before this one.
7570              intPos02 = InStr(intPos03, strFilter01, ANDF)
7580              If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
7590                strFilter01 = strFilter01 & ANDF & TRANS_END & .TransDateEnd & "#"
7600              Else
                    ' ** Add this clause to the middle of the filter.
7610                strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & TRANS_END & .TransDateEnd & "#" & Mid(strFilter01, intPos02)
7620              End If
7630            End If
7640          Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
7650            intPos02 = InStr((intPos01 + Len(TRANS_END) + 1), strFilter01, "#")  ' ** Find the closing paren.
7660            If intPos02 > 0 Then
7670              strFilter01 = Left(strFilter01, (intPos01 - 1)) & TRANS_END & .TransDateEnd & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
7680            Else
7690              strFilter01 = Left(strFilter01, (intPos01 - 1)) & TRANS_END & .TransDateEnd & "#"
7700            End If
7710          End If
7720        End If
7730        strFilter02 = strFilter01
7740        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
7750        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
7760        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
7770        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
7780        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
7790        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
7800      Else
7810        strFilter02 = strFilter01
7820        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
7830        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem TRANS_END  ' ** Form Procedure: frmTransaction_Audit_Sub.
7840        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
7850        frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub.
7860        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
7870        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem TRANS_END  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
7880        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
7890        frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
7900        If strFilter01 <> vbNullString Then
7910          frm.CurrentFilter2 = strFilter02
7920          frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
7930          frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
7940          frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
7950          frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
7960          frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
7970          frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
7980        End If
7990      End If
8000      frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .cmbAccounts.Enabled = True Then
          '  .cmbAccounts.SetFocus
          'End If
8010      DoCmd.Hourglass False
8020      DoEvents
8030    End With

EXITP:
8040    Set frm = Nothing
8050    Exit Sub

ERRH:
340       DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub TransDateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8120    intShiftDown = (Shift And acShiftMask) > 0
8130    intAltDown = (Shift And acAltMask) > 0
8140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8160      Select Case intRetVal
          Case vbKeyReturn  ' ** Let vbKeyTab be handled normally.
8170        With Me
8180          If IsNull(.TransDateEnd) Then
8190            If .TransDateEnd.text = vbNullString Then
                  ' ** Populate it with today's date.
8200              .TransDateEnd = Date
8210              TransDateEnd_AfterUpdate  ' ** Procedure: Below.
8220            End If
8230          End If
8240        End With
8250      End Select
8260    End If

        ' ** Ctrl keys.
8270    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8280      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8290        With Me
8300          .cmdCalendar2.SetFocus
8310        End With
8320      End Select
8330    End If

        ' ** Ctrl-Shift keys.
8340    If intCtrlDown And (Not intAltDown) And intShiftDown Then
8350      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8360        With Me
8370          .cmdCalendar1.SetFocus
8380        End With
8390      End Select
8400    End If

EXITP:
8410    KeyCode = intRetVal
8420    Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub TransDateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateEnd_MouseMove"

8510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8520    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

8610    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
8620    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

8710    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
8720    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmdCalendar2_Click()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

8810    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
8820    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8920    intShiftDown = (Shift And acShiftMask) > 0
8930    intAltDown = (Shift And acAltMask) > 0
8940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8960      Select Case intRetVal
          Case vbKeyTab
8970        With Me
8980          intRetVal = 0
8990          If .cmbAccounts.Enabled = True Then
9000            .cmbAccounts.SetFocus
9010          Else

9020          End If
9030        End With
9040      End Select
9050    End If

        ' ** Shift keys.
9060    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9070      Select Case intRetVal
          Case vbKeyTab
9080        With Me
9090          intRetVal = 0
9100          .TransDateEnd.SetFocus
9110        End With
9120      End Select
9130    End If

EXITP:
9140    KeyCode = intRetVal
9150    Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

9210    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
9220    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

9310    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
9320    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

9410    Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
9420    Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmbAccounts_AfterUpdate()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

9510    With Me
9520      DoCmd.Hourglass True
9530      DoEvents
9540      Set frm = .Parent
9550      If IsNull(.cmbAccounts) = False Then
9560        If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
9570          strFilter01 = ACCT_NUM & .cmbAccounts & "'"
9580        Else
              ' ** There are clauses present.

9590          FilterRec_GetArr  ' ** Procedure: Below.

9600      DoCmd.Hourglass True
9610      DoEvents
9620          intPos01 = InStr(strFilter01, ACCT_NUM)
9630          If intPos01 = 0& Then
                ' ** This clause isn't present.
9640            intPos03 = 0&: intPos04 = 0&
9650            For lngX = (lngFilts - 1&) To 0& Step -1&
9660              If arr_varFilt(F_CONST, lngX) = ACCT_NUM Then
9670                intPos04 = -1
9680              ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
9690                intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
9700                If intPos03 > 0 Then
9710                  intPos04 = 0&
9720                  Exit For
9730                End If
9740              End If
9750            Next
9760      DoCmd.Hourglass True
9770      DoEvents
9780            If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
9790              strFilter01 = ACCT_NUM & .cmbAccounts & "'" & ANDF & strFilter01
9800            Else
                  ' ** There's a clause before this one.
9810              intPos02 = InStr(intPos03, strFilter01, ANDF)
9820              If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
9830                strFilter01 = strFilter01 & ANDF & ACCT_NUM & .cmbAccounts & "'"
9840              Else
                    ' ** Add this clause to the middle of the filter.
9850                strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & ACCT_NUM & .cmbAccounts & "'" & Mid(strFilter01, intPos02)
9860              End If
9870            End If
9880          Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
9890            intPos02 = InStr((intPos01 + Len(ACCT_NUM) + 1), strFilter01, "'")  ' ** Find the closing paren.
9900            If intPos02 > 0 Then
9910              strFilter01 = Left(strFilter01, (intPos01 - 1)) & ACCT_NUM & .cmbAccounts & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
9920            Else
9930              strFilter01 = Left(strFilter01, (intPos01 - 1)) & ACCT_NUM & .cmbAccounts & "'"
9940            End If
9950          End If
9960        End If
9970      DoCmd.Hourglass True
9980      DoEvents
9990        strFilter02 = strFilter01
10000       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
10010     DoCmd.Hourglass True
10020     DoEvents
10030       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
10040     DoCmd.Hourglass True
10050     DoEvents
10060       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
10070     DoCmd.Hourglass True
10080     DoEvents
10090       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
10100     DoCmd.Hourglass True
10110     DoEvents
10120       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
10130     DoCmd.Hourglass True
10140     DoEvents
10150       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
10160     DoCmd.Hourglass True
10170     DoEvents
10180     Else
10190       strFilter02 = strFilter01
10200       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
10210     DoCmd.Hourglass True
10220     DoEvents
10230       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem ACCT_NUM  ' ** Form Procedure: frmTransaction_Audit_Sub.
10240     DoCmd.Hourglass True
10250     DoEvents
10260       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
10270     DoCmd.Hourglass True
10280     DoEvents
10290       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub.
10300     DoCmd.Hourglass True
10310     DoEvents
10320       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
10330     DoCmd.Hourglass True
10340     DoEvents
10350       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem ACCT_NUM  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
10360     DoCmd.Hourglass True
10370     DoEvents
10380       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
10390     DoCmd.Hourglass True
10400     DoEvents
10410       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
10420     DoCmd.Hourglass True
10430     DoEvents
10440     End If
10450     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .cmbAssets.Enabled = True Then
          '  .cmbAssets.SetFocus
          'End If
10460     DoCmd.Hourglass False
10470     DoEvents
10480   End With

EXITP:
10490   Set frm = Nothing
10500   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10620   intShiftDown = (Shift And acShiftMask) > 0
10630   intAltDown = (Shift And acAltMask) > 0
10640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10670       With Me
10680         intRetVal = 0
10690         If .cmbAssets.Enabled = True Then
10700           .cmbAssets.SetFocus
10710         Else

10720         End If
10730       End With
10740     End Select
10750   End If

        ' ** Shift keys.
10760   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10770     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10780       With Me
10790         intRetVal = 0
10800         If .TransDateEnd.Enabled = True Then
10810           .TransDateEnd.SetFocus
10820         Else

10830         End If
10840       End With
10850     End Select
10860   End If

        ' ** Ctrl keys.
10870   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10880     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10890       With Me
10900         intRetVal = 0
10910         .opgAccountSource.SetFocus
10920       End With
10930     End Select
10940   End If

        ' ** Ctrl-Shift keys.
10950   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10970       With Me
10980         intRetVal = 0
10990         If .cmdCalendar2.Enabled = True Then
11000           .cmdCalendar2.SetFocus
11010         Else
11020           Beep
11030         End If
11040       End With
11050     End Select
11060   End If

EXITP:
11070   KeyCode = intRetVal
11080   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmbAccounts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_MouseMove"

11110   With Me
11120     Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.
11130     .Parent.Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.
11140   End With

EXITP:
11150   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub opgAccountSource_AfterUpdate()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

11210   AcctSource_After_TA Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
11220   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub opgAccountSource_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_MouseMove"

11310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11320   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub opgAccountSource_optName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_MouseMove"

11410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11420   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_MouseMove"

11510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11520   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmbAssets_AfterUpdate()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

11610   With Me
11620     DoCmd.Hourglass True
11630     DoEvents
11640     Set frm = .Parent
11650     If IsNull(.cmbAssets) = False Then
11660       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
11670         strFilter01 = ASSET_NUM & CStr(.cmbAssets)
11680       Else
              ' ** There are clauses present.

11690         FilterRec_GetArr  ' ** Procedure: Below.

11700         intPos01 = InStr(strFilter01, ASSET_NUM)
11710         If intPos01 = 0& Then
                ' ** This clause isn't present.
11720           intPos03 = 0&: intPos04 = 0&
11730           For lngX = (lngFilts - 1&) To 0& Step -1&
11740             If arr_varFilt(F_CONST, lngX) = ASSET_NUM Then
11750               intPos04 = -1
11760             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
11770               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
11780               If intPos03 > 0 Then
11790                 intPos04 = 0&
11800                 Exit For
11810               End If
11820             End If
11830           Next
11840           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
11850             strFilter01 = ASSET_NUM & CStr(.cmbAssets) & ANDF & strFilter01
11860           Else
                  ' ** There's a clause before this one.
11870             intPos02 = InStr(intPos03, strFilter01, ANDF)
11880             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
11890               strFilter01 = strFilter01 & ANDF & ASSET_NUM & CStr(.cmbAssets)
11900             Else
                    ' ** Add this clause to the middle of the filter.
11910               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & ASSET_NUM & CStr(.cmbAssets) & Mid(strFilter01, intPos02)
11920             End If
11930           End If
11940         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
11950           intPos02 = InStr((intPos01 + Len(ASSET_NUM) + 1), strFilter01, ANDF)  ' ** Find the next ' And '.
11960           If intPos02 > 0 Then
11970             strFilter01 = Left(strFilter01, (intPos01 - 1)) & ASSET_NUM & CStr(.cmbAssets) & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
11980           Else
11990             strFilter01 = Left(strFilter01, (intPos01 - 1)) & ASSET_NUM & CStr(.cmbAssets)
12000           End If
12010         End If
12020       End If
12030       strFilter02 = strFilter01
12040       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
12050       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
12060       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
12070       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
12080       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
12090       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
12100     Else
12110       strFilter02 = strFilter01
12120       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
12130       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem ASSET_NUM  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
12140       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
12150       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
12160       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
12170       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem ASSET_NUM  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
12180       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
12190       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
12200     End If
12210     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .cmbCurrencies.Enabled = True Then
          '  .cmbCurrencies.SetFocus
          'ElseIf .AssetDateStart.Enabled = True Then
          '  .AssetDateStart.SetFocus
          'End If
12220     DoCmd.Hourglass False
12230     DoEvents
12240   End With

EXITP:
12250   Set frm = Nothing
12260   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12320   intShiftDown = (Shift And acShiftMask) > 0
12330   intAltDown = (Shift And acAltMask) > 0
12340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12370       With Me
12380         intRetVal = 0
12390         If .cmbCurrencies.Enabled = True Then
12400           .cmbCurrencies.SetFocus
12410         ElseIf .AssetDateStart.Enabled = True Then
12420           .AssetDateStart.SetFocus
12430         Else

12440         End If
12450       End With
12460     End Select
12470   End If

        ' ** Shift keys.
12480   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12490     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12500       With Me
12510         intRetVal = 0
12520         If .cmbAccounts.Enabled = True Then
12530           .cmbAccounts.SetFocus
12540         Else

12550         End If
12560       End With
12570     End Select
12580   End If

        ' ** Ctrl keys.
12590   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
12600     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12610       With Me
12620         intRetVal = 0
12630         .opgAssetSource.SetFocus
12640       End With
12650     End Select
12660   End If

        ' ** Ctrl-Shift keys.
12670   If intCtrlDown And (Not intAltDown) And intShiftDown Then
12680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12690       With Me
12700         intRetVal = 0
12710         If .opgAccountSource.Enabled = True Then
12720           .opgAccountSource.SetFocus
12730         Else
12740           Beep
12750         End If
12760       End With
12770     End Select
12780   End If

EXITP:
12790   KeyCode = intRetVal
12800   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub opgAssetSource_AfterUpdate()

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_AfterUpdate"

12910   AssetSource_After_TA Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
12920   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub cmbCurrencies_AfterUpdate()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbCurrencies_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

13010   With Me
13020     DoCmd.Hourglass True
13030     DoEvents
13040     Set frm = .Parent
13050     If IsNull(.cmbCurrencies) = False Then
13060       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
13070         strFilter01 = CURR_NUM & CStr(.cmbCurrencies)
13080       Else
              ' ** There are clauses present.

13090         FilterRec_GetArr  ' ** Procedure: Below.

13100         intPos01 = InStr(strFilter01, CURR_NUM)
13110         If intPos01 = 0& Then
                ' ** This clause isn't present.
13120           intPos03 = 0&: intPos04 = 0&
13130           For lngX = (lngFilts - 1&) To 0& Step -1&
13140             If arr_varFilt(F_CONST, lngX) = CURR_NUM Then
13150               intPos04 = -1
13160             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
13170               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
13180               If intPos03 > 0 Then
13190                 intPos04 = 0&
13200                 Exit For
13210               End If
13220             End If
13230           Next
13240           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
13250             strFilter01 = CURR_NUM & CStr(.cmbCurrencies) & ANDF & strFilter01
13260           Else
                  ' ** There's a clause before this one.
13270             intPos02 = InStr(intPos03, strFilter01, ANDF)
13280             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
13290               strFilter01 = strFilter01 & ANDF & CURR_NUM & CStr(.cmbCurrencies)
13300             Else
                    ' ** Add this clause to the middle of the filter.
13310               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & CURR_NUM & CStr(.cmbCurrencies) & Mid(strFilter01, intPos02)
13320             End If
13330           End If
13340         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
13350           intPos02 = InStr((intPos01 + Len(CURR_NUM) + 1), strFilter01, ANDF)  ' ** Find the next ' And '.
13360           If intPos02 > 0 Then
13370             strFilter01 = Left(strFilter01, (intPos01 - 1)) & CURR_NUM & CStr(.cmbCurrencies) & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
13380           Else
13390             strFilter01 = Left(strFilter01, (intPos01 - 1)) & CURR_NUM & CStr(.cmbCurrencies)
13400           End If
13410         End If
13420       End If
13430       strFilter02 = strFilter01
13440       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
13450       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
13460       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
13470       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
13480       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
13490       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
13500     Else
13510       strFilter02 = strFilter01
13520       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
13530       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem CURR_NUM  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
13540       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
13550       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
13560       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
13570       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem CURR_NUM  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
13580       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
13590       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
13600     End If
13610     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .AssetDateStart.Enabled = True Then
          '  .AssetDateStart.SetFocus
          'End If
13620     DoCmd.Hourglass False
13630     DoEvents
13640   End With

EXITP:
13650   Set frm = Nothing
13660   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmbCurrencies_KeyDown(KeyCode As Integer, Shift As Integer)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbCurrencies_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13720   intShiftDown = (Shift And acShiftMask) > 0
13730   intAltDown = (Shift And acAltMask) > 0
13740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13770       With Me
13780         intRetVal = 0
13790         If .AssetDateStart.Enabled = True Then
13800           .AssetDateStart.SetFocus
13810         Else

13820         End If
13830       End With
13840     End Select
13850   End If

        ' ** Shift keys.
13860   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13870     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13880       With Me
13890         intRetVal = 0
13900         If .cmbAssets.Enabled = True Then
13910           .cmbAssets.SetFocus
13920         Else

13930         End If
13940       End With
13950     End Select
13960   End If

        ' ** Ctrl-Shift keys.
13970   If intCtrlDown And (Not intAltDown) And intShiftDown Then
13980     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13990       With Me
14000         intRetVal = 0
14010         If .opgAssetSource.Enabled = True Then
14020           .opgAssetSource.SetFocus
14030         Else
14040           Beep
14050         End If
14060       End With
14070     End Select
14080   End If

EXITP:
14090   KeyCode = intRetVal
14100   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub AssetDateStart_BeforeUpdate(Cancel As Integer)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetDateStart_BeforeUpdate"

14210   With Me
14220     If .AssetDateStart > Date Then
14230       MsgBox "Please enter a date prior to today.", vbInformation + vbOKOnly, "Invalid Date"
14240       Cancel = -1
14250     End If
14260   End With

EXITP:
14270   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub AssetDateStart_AfterUpdate()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetDateStart_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

14310   With Me
14320     DoCmd.Hourglass True
14330     DoEvents
14340     Set frm = .Parent
14350     If IsNull(.AssetDateStart) = False Then
14360       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
14370         strFilter01 = ASSET_START & .AssetDateStart & "#"
14380       Else
              ' ** There are clauses present.

14390         FilterRec_GetArr  ' ** Procedure: Below.

              ' ** There are clauses present.
14400         intPos01 = InStr(strFilter01, ASSET_START)
14410         If intPos01 = 0& Then
                ' ** This clause isn't present.
14420           intPos03 = 0&: intPos04 = 0&
14430           For lngX = (lngFilts - 1&) To 0& Step -1&
14440             If arr_varFilt(F_CONST, lngX) = ASSET_START Then
14450               intPos04 = -1
14460             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
14470               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
14480               If intPos03 > 0 Then
14490                 intPos04 = 0&
14500                 Exit For
14510               End If
14520             End If
14530           Next
14540           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
14550             strFilter01 = ASSET_START & .AssetDateStart & "#" & ANDF & strFilter01
14560           Else
                  ' ** There's a clause before this one.
14570             intPos02 = InStr(intPos03, strFilter01, ANDF)
14580             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
14590               strFilter01 = strFilter01 & ANDF & ASSET_START & .AssetDateStart & "#"
14600             Else
                    ' ** Add this clause to the middle of the filter.
14610               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & ASSET_START & .AssetDateStart & "#" & Mid(strFilter01, intPos02)
14620             End If
14630           End If
14640         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
14650           intPos02 = InStr((intPos01 + Len(ASSET_START) + 1), strFilter01, "#")  ' ** Find the closing paren.
14660           If intPos02 > 0 Then
14670             strFilter01 = Left(strFilter01, (intPos01 - 1)) & ASSET_START & .AssetDateStart & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
14680           Else
14690             strFilter01 = Left(strFilter01, (intPos01 - 1)) & ASSET_START & .AssetDateStart & "#"
14700           End If
14710         End If
14720       End If
14730       strFilter02 = strFilter01
14740       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
14750       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
14760       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
14770       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
14780       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
14790       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
14800     Else
14810       strFilter02 = strFilter01
14820       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
14830       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem ASSET_START  ' ** Form Procedure: frmTransaction_Audit_Sub.
14840       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
14850       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub.
14860       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
14870       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem ASSET_START  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
14880       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
14890       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
14900     End If
14910     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          '.AssetDateEnd.SetFocus
14920     DoCmd.Hourglass False
14930     DoEvents
14940   End With

EXITP:
14950   Set frm = Nothing
14960   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub AssetDateStart_KeyDown(KeyCode As Integer, Shift As Integer)

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetDateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15020   intShiftDown = (Shift And acShiftMask) > 0
15030   intAltDown = (Shift And acAltMask) > 0
15040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15070       With Me
15080         intRetVal = 0
15090         .AssetDateEnd.SetFocus
15100       End With
15110     End Select
15120   End If

        ' ** Shift keys.
15130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15150       With Me
15160         intRetVal = 0
15170         If .cmbCurrencies.Enabled = True Then
15180           .cmbCurrencies.SetFocus
15190         ElseIf .cmbAssets.Enabled = True Then
15200           .cmbAssets.SetFocus
15210         Else

15220         End If
15230       End With
15240     End Select
15250   End If

        ' ** Ctrl keys.
15260   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
15270     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15280       With Me
15290         intRetVal = 0
15300         .cmdCalendar3.SetFocus
15310       End With
15320     End Select
15330   End If

        ' ** Ctrl-Shift keys.
15340   If intCtrlDown And (Not intAltDown) And intShiftDown Then
15350     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15360       With Me
15370         intRetVal = 0
15380         If .cmbCurrencies.Enabled = False Then
15390           If .opgAssetSource.Enabled = True Then
15400             .opgAssetSource.SetFocus
15410           Else
15420             Beep
15430           End If
15440         Else
15450           Beep
15460         End If
15470       End With
15480     End Select
15490   End If

EXITP:
15500   KeyCode = intRetVal
15510   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar3_GotFocus()

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_GotFocus"

15610   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
15620   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_MouseDown"

15710   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
15720   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmdCalendar3_Click()

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_Click"

15810   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
15820   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar3_KeyDown(KeyCode As Integer, Shift As Integer)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15920   intShiftDown = (Shift And acShiftMask) > 0
15930   intAltDown = (Shift And acAltMask) > 0
15940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15960     Select Case intRetVal
          Case vbKeyTab
15970       With Me
15980         intRetVal = 0
15990         .AssetDateEnd.SetFocus
16000       End With
16010     End Select
16020   End If

        ' ** Shift keys.
16030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16040     Select Case intRetVal
          Case vbKeyTab
16050       With Me
16060         intRetVal = 0
16070         .AssetDateStart.SetFocus
16080       End With
16090     End Select
16100   End If

EXITP:
16110   KeyCode = intRetVal
16120   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_MouseMove"

16210     Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
16220   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_MouseUp"

16310   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
16320   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar3_LostFocus()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar3_LostFocus"

16410   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
16420   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub AssetDateEnd_BeforeUpdate(Cancel As Integer)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetDateEnd_BeforeUpdate"

16510   With Me
16520     If IsNull(.AssetDateStart) = False Then
16530       If .AssetDateEnd < .AssetDateStart Then
16540         MsgBox "Please enter a date the same as or after the From date.", vbInformation + vbOKOnly, "Invalid Date"
16550         Cancel = -1
16560       End If
16570     End If
16580   End With

EXITP:
16590   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub AssetDateEnd_AfterUpdate()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetDateEnd_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

16610   With Me
16620     DoCmd.Hourglass True
16630     DoEvents
16640     Set frm = .Parent
16650     If IsNull(.AssetDateEnd) = False Then
16660       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
16670         strFilter01 = ASSET_END & .AssetDateEnd & "#"
16680       Else
              ' ** There are clauses present.

16690         FilterRec_GetArr  ' ** Procedure: Below.

16700         intPos01 = InStr(strFilter01, ASSET_END)
16710         If intPos01 = 0& Then
                ' ** This clause isn't present.
16720           intPos03 = 0&: intPos04 = 0&
16730           For lngX = (lngFilts - 1&) To 0& Step -1&
16740             If arr_varFilt(F_CONST, lngX) = ASSET_END Then
16750               intPos04 = -1
16760             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
16770               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
16780               If intPos03 > 0 Then
16790                 intPos04 = 0&
16800                 Exit For
16810               End If
16820             End If
16830           Next
16840           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
16850             strFilter01 = ASSET_END & .AssetDateEnd & "#" & ANDF & strFilter01
16860           Else
                  ' ** There's a clause before this one.
16870             intPos02 = InStr(intPos03, strFilter01, ANDF)
16880             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
16890               strFilter01 = strFilter01 & ANDF & ASSET_END & .AssetDateEnd & "#"
16900             Else
                    ' ** Add this clause to the middle of the filter.
16910               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & ASSET_END & .AssetDateEnd & "#" & Mid(strFilter01, intPos02)
16920             End If
16930           End If
16940         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
16950           intPos02 = InStr((intPos01 + Len(ASSET_END) + 1), strFilter01, "#")  ' ** Find the closing paren.
16960           If intPos02 > 0 Then
16970             strFilter01 = Left(strFilter01, (intPos01 - 1)) & ASSET_END & .AssetDateEnd & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
16980           Else
16990             strFilter01 = Left(strFilter01, (intPos01 - 1)) & ASSET_END & .AssetDateEnd & "#"
17000           End If
17010         End If
17020       End If
17030       strFilter02 = strFilter01
17040       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
17050       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
17060       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
17070       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
17080       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
17090       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
17100     Else
17110       strFilter02 = strFilter01
17120       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
17130       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem ASSET_END  ' ** Form Procedure: frmTransaction_Audit_Sub.
17140       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
17150       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub.
17160       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
17170       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem ASSET_END  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
17180       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
17190       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
17200     End If
17210     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .PurchaseDateStart.Enabled = True Then
          '  .PurchaseDateStart.SetFocus
          'End If
17220     DoCmd.Hourglass False
17230     DoEvents
17240   End With

EXITP:
17250   Set frm = Nothing
17260   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub AssetDateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetDateEnd_KeyDown"

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
17390         If .PurchaseDateStart.Enabled = True Then
17400           .PurchaseDateStart.SetFocus
17410         Else

17420         End If
17430       End With
17440     End Select
17450   End If

        ' ** Shift keys.
17460   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17470     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17480       With Me
17490         intRetVal = 0
17500         .AssetDateStart.SetFocus
17510       End With
17520     End Select
17530   End If

        ' ** Ctrl keys.
17540   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
17550     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17560       With Me
17570         intRetVal = 0
17580         .cmdCalendar4.SetFocus
17590       End With
17600     End Select
17610   End If

        ' ** Ctrl-Shift keys.
17620   If intCtrlDown And (Not intAltDown) And intShiftDown Then
17630     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17640       With Me
17650         intRetVal = 0
17660         .cmdCalendar3.SetFocus
17670       End With
17680     End Select
17690   End If

EXITP:
17700   KeyCode = intRetVal
17710   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar4_GotFocus()

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar4_GotFocus"

17810   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
17820   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar4_MouseDown"

17910   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
17920   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmdCalendar4_Click()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar4_Click"

18010   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
18020   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar4_KeyDown(KeyCode As Integer, Shift As Integer)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar4_KeyDown"

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
18190         If .PurchaseDateStart.Enabled = True Then
18200           .PurchaseDateStart.SetFocus
18210         Else

18220         End If
18230       End With
18240     End Select
18250   End If

        ' ** Shift keys.
18260   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18270     Select Case intRetVal
          Case vbKeyTab
18280       With Me
18290         intRetVal = 0
18300         .AssetDateEnd.SetFocus
18310       End With
18320     End Select
18330   End If

EXITP:
18340   KeyCode = intRetVal
18350   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar4_MouseMove"

18410     Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
18420   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar4_MouseUp"

18510   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
18520   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar4_LostFocus()

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar4_LostFocus"

18610   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
18620   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub PurchaseDateStart_BeforeUpdate(Cancel As Integer)

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDateStart_BeforeUpdate"

18710   With Me
18720     If .PurchaseDateStart > Date Then
18730       MsgBox "Please enter a date prior to today.", vbInformation + vbOKOnly, "Invalid Date"
18740       Cancel = -1
18750     End If
18760   End With

EXITP:
18770   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub PurchaseDateStart_AfterUpdate()

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDateStart_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

18810   With Me
18820     DoCmd.Hourglass True
18830     DoEvents
18840     Set frm = .Parent
18850     If IsNull(.PurchaseDateStart) = False Then
18860       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
18870         strFilter01 = PURCH_START & .PurchaseDateStart & "#"
18880       Else
              ' ** There are clauses present.

18890         FilterRec_GetArr  ' ** Procedure: Below.

18900         intPos01 = InStr(strFilter01, PURCH_START)
18910         If intPos01 = 0& Then
                ' ** This clause isn't present.
18920           intPos03 = 0&: intPos04 = 0&
18930           For lngX = (lngFilts - 1&) To 0& Step -1&
18940             If arr_varFilt(F_CONST, lngX) = PURCH_START Then
18950               intPos04 = -1
18960             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
18970               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
18980               If intPos03 > 0 Then
18990                 intPos04 = 0&
19000                 Exit For
19010               End If
19020             End If
19030           Next
19040           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
19050             strFilter01 = PURCH_START & .PurchaseDateStart & "#" & ANDF & strFilter01
19060           Else
                  ' ** There's a clause before this one.
19070             intPos02 = InStr(intPos03, strFilter01, ANDF)
19080             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
19090               strFilter01 = strFilter01 & ANDF & PURCH_START & .PurchaseDateStart & "#"
19100             Else
                    ' ** Add this clause to the middle of the filter.
19110               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & PURCH_START & .PurchaseDateStart & "#" & Mid(strFilter01, intPos02)
19120             End If
19130           End If
19140         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
19150           intPos02 = InStr((intPos01 + Len(PURCH_START) + 1), strFilter01, "#")  ' ** Find the closing paren.
19160           If intPos02 > 0 Then
19170             strFilter01 = Left(strFilter01, (intPos01 - 1)) & PURCH_START & .PurchaseDateStart & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
19180           Else
19190             strFilter01 = Left(strFilter01, (intPos01 - 1)) & PURCH_START & .PurchaseDateStart & "#"
19200           End If
19210         End If
19220       End If
19230       strFilter02 = strFilter01
19240       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
19250       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
19260       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
19270       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
19280       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
19290       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
19300     Else
19310       strFilter02 = strFilter01
19320       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
19330       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem PURCH_START  ' ** Form Procedure: frmTransaction_Audit_Sub.
19340       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
19350       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub.
19360       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
19370       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem PURCH_START  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
19380       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
19390       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
19400     End If
19410     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          '.PurchaseDateEnd.SetFocus
19420     DoCmd.Hourglass False
19430     DoEvents
19440   End With

EXITP:
19450   Set frm = Nothing
19460   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub PurchaseDateStart_KeyDown(KeyCode As Integer, Shift As Integer)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19520   intShiftDown = (Shift And acShiftMask) > 0
19530   intAltDown = (Shift And acAltMask) > 0
19540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19570       With Me
19580         intRetVal = 0
19590         .PurchaseDateEnd.SetFocus
19600       End With
19610     End Select
19620   End If

        ' ** Shift keys.
19630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19650       With Me
19660         intRetVal = 0
19670         If .AssetDateEnd.Enabled = True Then
19680           .AssetDateEnd.SetFocus
19690         Else

19700         End If
19710       End With
19720     End Select
19730   End If

        ' ** Ctrl keys.
19740   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
19750     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19760       With Me
19770         intRetVal = 0
19780         .cmdCalendar5.SetFocus
19790       End With
19800     End Select
19810   End If

        ' ** Ctrl-Shift keys.
19820   If intCtrlDown And (Not intAltDown) And intShiftDown Then
19830     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19840       With Me
19850         intRetVal = 0
19860         If .cmdCalendar4.Enabled = True Then
19870           .cmdCalendar4.SetFocus
19880         Else
19890           Beep
19900         End If
19910       End With
19920     End Select
19930   End If

EXITP:
19940   KeyCode = intRetVal
19950   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar5_GotFocus()

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar5_GotFocus"

20010   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
20020   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar5_MouseDown"

20110   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
20120   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmdCalendar5_Click()

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar5_Click"

20210   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
20220   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar5_KeyDown(KeyCode As Integer, Shift As Integer)

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar5_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20320   intShiftDown = (Shift And acShiftMask) > 0
20330   intAltDown = (Shift And acAltMask) > 0
20340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20360     Select Case intRetVal
          Case vbKeyTab
20370       With Me
20380         intRetVal = 0
20390         .PurchaseDateEnd.SetFocus
20400       End With
20410     End Select
20420   End If

        ' ** Shift keys.
20430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20440     Select Case intRetVal
          Case vbKeyTab
20450       With Me
20460         intRetVal = 0
20470         .PurchaseDateStart.SetFocus
20480       End With
20490     End Select
20500   End If

EXITP:
20510   KeyCode = intRetVal
20520   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar5_MouseMove"

20610     Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
20620   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar5_MouseUp"

20710   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
20720   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar5_LostFocus()

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar5_LostFocus"

20810   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
20820   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub PurchaseDateEnd_BeforeUpdate(Cancel As Integer)

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDateEnd_BeforeUpdate"

20910   With Me
20920     If IsNull(.PurchaseDateStart) = False Then
20930       If .PurchaseDateEnd < .PurchaseDateStart Then
20940         MsgBox "Please enter a date the same as or after the From date.", vbInformation + vbOKOnly, "Invalid Date"
20950         Cancel = -1
20960       End If
20970     End If
20980   End With

EXITP:
20990   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub PurchaseDateEnd_AfterUpdate()

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDateEnd_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

21010   With Me
21020     DoCmd.Hourglass True
21030     DoEvents
21040     Set frm = .Parent
21050     If IsNull(.PurchaseDateEnd) = False Then
21060       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
21070         strFilter01 = PURCH_END & .PurchaseDateEnd & "#"
21080       Else
              ' ** There are clauses present.

21090         FilterRec_GetArr  ' ** Procedure: Below.

21100         intPos01 = InStr(strFilter01, PURCH_END)
21110         If intPos01 = 0& Then
                ' ** This clause isn't present.
21120           intPos03 = 0&: intPos04 = 0&
21130           For lngX = (lngFilts - 1&) To 0& Step -1&
21140             If arr_varFilt(F_CONST, lngX) = PURCH_END Then
21150               intPos04 = -1
21160             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
21170               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
21180               If intPos03 > 0 Then
21190                 intPos04 = 0&
21200                 Exit For
21210               End If
21220             End If
21230           Next
21240           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
21250             strFilter01 = PURCH_END & .PurchaseDateEnd & "#" & ANDF & strFilter01
21260           Else
                  ' ** There's a clause before this one.
21270             intPos02 = InStr(intPos03, strFilter01, ANDF)
21280             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
21290               strFilter01 = strFilter01 & ANDF & PURCH_END & .PurchaseDateEnd & "#"
21300             Else
                    ' ** Add this clause to the middle of the filter.
21310               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & PURCH_END & .PurchaseDateEnd & "#" & Mid(strFilter01, intPos02)
21320             End If
21330           End If
21340         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
21350           intPos02 = InStr((intPos01 + Len(PURCH_END) + 1), strFilter01, "#")  ' ** Find the closing paren.
21360           If intPos02 > 0 Then
21370             strFilter01 = Left(strFilter01, (intPos01 - 1)) & PURCH_END & .PurchaseDateEnd & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
21380           Else
21390             strFilter01 = Left(strFilter01, (intPos01 - 1)) & PURCH_END & .PurchaseDateEnd & "#"
21400           End If
21410         End If
21420       End If
21430       strFilter02 = strFilter01
21440       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
21450       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
21460       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
21470       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
21480       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
21490       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
21500     Else
21510       strFilter02 = strFilter01
21520       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
21530       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem PURCH_END  ' ** Form Procedure: frmTransaction_Audit_Sub.
21540       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
21550       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub.
21560       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
21570       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem PURCH_END  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
21580       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
21590       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
21600     End If
21610     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .ledger_description.Enabled = True Then
          '  .ledger_description.SetFocus
          'End If
21620     DoCmd.Hourglass False
21630     DoEvents
21640   End With

EXITP:
21650   Set frm = Nothing
21660   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub PurchaseDateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21720   intShiftDown = (Shift And acShiftMask) > 0
21730   intAltDown = (Shift And acAltMask) > 0
21740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21770       With Me
21780         intRetVal = 0
21790         If .ledger_description.Enabled = True Then
21800           .ledger_description.SetFocus
21810         Else

21820         End If
21830       End With
21840     End Select
21850   End If

        ' ** Shift keys.
21860   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21870     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21880       With Me
21890         intRetVal = 0
21900         .PurchaseDateStart.SetFocus
21910       End With
21920     End Select
21930   End If

        ' ** Ctrl keys.
21940   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
21950     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21960       With Me
21970         intRetVal = 0
21980         .cmdCalendar6.SetFocus
21990       End With
22000     End Select
22010   End If

        ' ** Ctrl-Shift keys.
22020   If intCtrlDown And (Not intAltDown) And intShiftDown Then
22030     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22040       With Me
22050         intRetVal = 0
22060         .cmdCalendar5.SetFocus
22070       End With
22080     End Select
22090   End If

EXITP:
22100   KeyCode = intRetVal
22110   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar6_GotFocus()

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar6_GotFocus"

22210   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
22220   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar6_MouseDown"

22310   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
22320   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmdCalendar6_Click()

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar6_Click"

22410   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
22420   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar6_KeyDown(KeyCode As Integer, Shift As Integer)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar6_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22520   intShiftDown = (Shift And acShiftMask) > 0
22530   intAltDown = (Shift And acAltMask) > 0
22540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22560     Select Case intRetVal
          Case vbKeyTab
22570       With Me
22580         intRetVal = 0
22590         If .ledger_description.Enabled = True Then
22600           .ledger_description.SetFocus
22610         Else

22620         End If
22630       End With
22640     End Select
22650   End If

        ' ** Shift keys.
22660   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22670     Select Case intRetVal
          Case vbKeyTab
22680       With Me
22690         intRetVal = 0
22700         .PurchaseDateEnd.SetFocus
22710       End With
22720     End Select
22730   End If

EXITP:
22740   KeyCode = intRetVal
22750   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar6_MouseMove"

22810   With Me

22820     Calendar_Handler THIS_PROC  ' ** Procedure: Below.

22830   End With

EXITP:
22840   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar6_MouseUp"

22910   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
22920   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar6_LostFocus()

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar6_LostFocus"

23010   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
23020   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub ledger_description_AfterUpdate()

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

23110   With Me
23120     DoCmd.Hourglass True
23130     DoEvents
23140     Set frm = .Parent
23150     If IsNull(.ledger_description) = False Then
23160       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
23170         strFilter01 = COMM_DESC & .ledger_description & "*'"  ' ** Wildcard at both beginning and end.
23180       Else
              ' ** There are clauses present.

23190         FilterRec_GetArr  ' ** Procedure: Below.

23200         intPos01 = InStr(strFilter01, COMM_DESC)
23210         If intPos01 = 0& Then
                ' ** This clause isn't present.
23220           intPos03 = 0&: intPos04 = 0&
23230           For lngX = (lngFilts - 1&) To 0& Step -1&
23240             If arr_varFilt(F_CONST, lngX) = COMM_DESC Then
23250               intPos04 = -1
23260             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
23270               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
23280               If intPos03 > 0 Then
23290                 intPos04 = 0&
23300                 Exit For
23310               End If
23320             End If
23330           Next
23340           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
23350             strFilter01 = COMM_DESC & .ledger_description & "*'" & ANDF & strFilter01
23360           Else
                  ' ** There's a clause before this one.
23370             intPos02 = InStr(intPos03, strFilter01, ANDF)
23380             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
23390               strFilter01 = strFilter01 & ANDF & COMM_DESC & .ledger_description & "*'"
23400             Else
                    ' ** Add this clause to the middle of the filter.
23410               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & COMM_DESC & .ledger_description & "*'" & Mid(strFilter01, intPos02)
23420             End If
23430           End If
23440         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
23450           intPos02 = InStr((intPos01 + Len(COMM_DESC) + 1), strFilter01, "*'")  ' ** Find the closing asterisk/paren.
23460           If intPos02 > 0 Then
23470             strFilter01 = Left(strFilter01, (intPos01 - 1)) & COMM_DESC & .ledger_description & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
23480           Else
23490             strFilter01 = Left(strFilter01, (intPos01 - 1)) & COMM_DESC & .ledger_description & "*'"
23500           End If
23510         End If
23520       End If
23530       strFilter02 = strFilter01
23540       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
23550       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
23560       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
23570       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
23580       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
23590       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
23600     Else
23610       strFilter02 = strFilter01
23620       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
23630       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem COMM_DESC  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
23640       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
23650       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
23660       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
23670       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem COMM_DESC  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
23680       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
23690       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
23700     End If
23710     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .cmbRecurringItems.Enabled = True Then
          '  .cmbRecurringItems.SetFocus
          'End If
23720     DoCmd.Hourglass False
23730     DoEvents
23740   End With

EXITP:
23750   Set frm = Nothing
23760   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub ledger_description_KeyDown(KeyCode As Integer, Shift As Integer)

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_KeyDown"

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
23890         If .cmbRecurringItems.Enabled = True Then
23900           .cmbRecurringItems.SetFocus
23910         Else

23920         End If
23930       End With
23940     End Select
23950   End If

        ' ** Shift keys.
23960   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23970     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23980       With Me
23990         intRetVal = 0
24000         If .PurchaseDateEnd.Enabled = True Then
24010           .PurchaseDateEnd.SetFocus
24020         Else

24030         End If
24040       End With
24050     End Select
24060   End If

        ' ** Ctrl-Shift keys.
24070   If intCtrlDown And (Not intAltDown) And intShiftDown Then
24080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24090       With Me
24100         intRetVal = 0
24110         If .cmdCalendar6.Enabled = True Then
24120           .cmdCalendar6.SetFocus
24130         Else
24140           Beep
24150         End If
24160       End With
24170     End Select
24180   End If

EXITP:
24190   KeyCode = intRetVal
24200   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub cmbRecurringItems_AfterUpdate()

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

24310   With Me
24320     DoCmd.Hourglass True
24330     DoEvents
24340     Set frm = .Parent
24350     If IsNull(.cmbRecurringItems) = False Then
24360       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
24370         strFilter01 = RECUR_ITEM & .cmbRecurringItems.Column(1) & "*'"
24380       Else
              ' ** There are clauses present.

24390         FilterRec_GetArr  ' ** Procedure: Below.

24400         intPos01 = InStr(strFilter01, RECUR_ITEM)
24410         If intPos01 = 0& Then
                ' ** This clause isn't present.
24420           intPos03 = 0&: intPos04 = 0&
24430           For lngX = (lngFilts - 1&) To 0& Step -1&
24440             If arr_varFilt(F_CONST, lngX) = RECUR_ITEM Then
24450               intPos04 = -1
24460             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
24470               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
24480               If intPos03 > 0 Then
24490                 intPos04 = 0&
24500                 Exit For
24510               End If
24520             End If
24530           Next
24540           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
24550             strFilter01 = RECUR_ITEM & .cmbRecurringItems.Column(1) & "*'" & ANDF & strFilter01
24560           Else
                  ' ** There's a clause before this one.
24570             intPos02 = InStr(intPos03, strFilter01, ANDF)
24580             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
24590               strFilter01 = strFilter01 & ANDF & RECUR_ITEM & .cmbRecurringItems.Column(1) & "*'"
24600             Else
                    ' ** Add this clause to the middle of the filter.
24610               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & RECUR_ITEM & .cmbRecurringItems.Column(1) & "*'" & Mid(strFilter01, intPos02)
24620             End If
24630           End If
24640         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
24650           intPos02 = InStr((intPos01 + Len(RECUR_ITEM) + 1), strFilter01, "*'")  ' ** Find the closing asterisk/paren.
24660           If intPos02 > 0 Then
24670             strFilter01 = Left(strFilter01, (intPos01 - 1)) & RECUR_ITEM & .cmbRecurringItems.Column(1) & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
24680           Else
24690             strFilter01 = Left(strFilter01, (intPos01 - 1)) & RECUR_ITEM & .cmbRecurringItems.Column(1) & "*'"
24700           End If
24710         End If
24720       End If
24730       strFilter02 = strFilter01
24740       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
24750       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
24760       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
24770       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
24780       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
24790       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
24800     Else
24810       strFilter02 = strFilter01
24820       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
24830       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem RECUR_ITEM  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
24840       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
24850       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
24860       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
24870       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem RECUR_ITEM  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
24880       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
24890       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
24900     End If
24910     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .cmbRevenueCodes.Enabled = True Then
          '  .cmbRevenueCodes.SetFocus
          'ElseIf .chkRevcodeType_Income.Enabled = True Then
          '  .chkRevcodeType_Income.SetFocus
          'End If
24920     DoCmd.Hourglass False
24930     DoEvents
24940   End With

EXITP:
24950   Set frm = Nothing
24960   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub cmbRevenueCodes_AfterUpdate()

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

25010   With Me
25020     DoCmd.Hourglass True
25030     DoEvents
25040     Set frm = .Parent
25050     If IsNull(.cmbRevenueCodes) = False Then
25060       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
25070         strFilter01 = REV_CODE & CStr(.cmbRevenueCodes)
25080       Else
              ' ** There are clauses present.

25090         FilterRec_GetArr  ' ** Procedure: Below.

25100         intPos01 = InStr(strFilter01, REV_CODE)
25110         If intPos01 = 0& Then
                ' ** This clause isn't present.
25120           intPos03 = 0&: intPos04 = 0&
25130           For lngX = (lngFilts - 1&) To 0& Step -1&
25140             If arr_varFilt(F_CONST, lngX) = REV_CODE Then
25150               intPos04 = -1
25160             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
25170               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
25180               If intPos03 > 0 Then
25190                 intPos04 = 0&
25200                 Exit For
25210               End If
25220             End If
25230           Next
25240           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
25250             strFilter01 = REV_CODE & CStr(.cmbRevenueCodes) & ANDF & strFilter01
25260           Else
                  ' ** There's a clause before this one.
25270             intPos02 = InStr(intPos03, strFilter01, ANDF)
25280             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
25290               strFilter01 = strFilter01 & ANDF & REV_CODE & CStr(.cmbRevenueCodes)
25300             Else
                    ' ** Add this clause to the middle of the filter.
25310               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & REV_CODE & CStr(.cmbRevenueCodes) & Mid(strFilter01, intPos02)
25320             End If
25330           End If
25340         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
25350           intPos02 = InStr((intPos01 + Len(REV_CODE) + 1), strFilter01, ANDF)  ' ** Find the next ' And '.
25360           If intPos02 > 0 Then
25370             strFilter01 = Left(strFilter01, (intPos01 - 1)) & REV_CODE & CStr(.cmbRevenueCodes) & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
25380           Else
25390             strFilter01 = Left(strFilter01, (intPos01 - 1)) & REV_CODE & CStr(.cmbRevenueCodes)
25400           End If
25410         End If
25420       End If
25430       strFilter02 = strFilter01
25440       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
25450       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
25460       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
25470       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
25480       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
25490       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
25500     Else
25510       strFilter02 = strFilter01
25520       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
25530       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem REV_CODE  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
25540       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
25550       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
25560       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
25570       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem REV_CODE  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
25580       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
25590       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
25600     End If
25610     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .cmbTaxCodes.Enabled = True Then
          '  .cmbTaxCodes.SetFocus
          'ElseIf .chkTaxcodeType_Income.Enabled = True Then
          '  .chkTaxcodeType_Income.SetFocus
          'End If
25620     DoCmd.Hourglass False
25630     DoEvents
25640   End With

EXITP:
25650   Set frm = Nothing
25660   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmbRevenueCodes_KeyDown(KeyCode As Integer, Shift As Integer)

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_KeyDown"

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
25790         If .cmbTaxCodes.Enabled = True Then
25800           .cmbTaxCodes.SetFocus
25810         Else

25820         End If
25830       End With
25840     End Select
25850   End If

        ' ** Shift keys.
25860   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25870     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25880       With Me
25890         intRetVal = 0
25900         If .cmbRecurringItems.Enabled = True Then
25910           .cmbRecurringItems.SetFocus
25920         Else

25930         End If
25940       End With
25950     End Select
25960   End If

        ' ** Ctrl keys.
25970   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
25980     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25990       With Me
26000         intRetVal = 0
26010         .chkRevcodeType_Income.SetFocus
26020       End With
26030     End Select
26040   End If

EXITP:
26050   KeyCode = intRetVal
26060   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub chkRevcodeType_Income_AfterUpdate()

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRevcodeType_Income_AfterUpdate"

26110   With Me
26120     DoCmd.Hourglass True
26130     DoEvents
26140     Select Case .chkRevcodeType_Income
          Case True
26150       If .cmbRevenueCodes.Enabled = False Then
26160         .cmbRevenueCodes.Enabled = True
26170         .cmbRevenueCodes.BorderColor = CLR_LTBLU2
26180         .cmbRevenueCodes.BackStyle = acBackStyleNormal
26190       End If
26200       Select Case .chkRevcodeType_Expense
            Case True
26210         If .cmbRevenueCodes.RowSource <> "qryTransaction_Audit_10_03_Revcode_Both" Then
26220           .cmbRevenueCodes.RowSource = "qryTransaction_Audit_10_03_Revcode_Both"
26230         End If
26240       Case False
26250         If .cmbRevenueCodes.RowSource <> "qryTransaction_Audit_10_03_Revcode_Inc" Then
26260           .cmbRevenueCodes.RowSource = "qryTransaction_Audit_10_03_Revcode_Inc"
26270         End If
26280       End Select
26290       .chkRevcodeType_Income_lbl.FontBold = True
26300     Case False
26310       Select Case .chkRevcodeType_Expense
            Case True
26320         If .cmbRevenueCodes.Enabled = False Then .cmbRevenueCodes.Enabled = True
26330         If .cmbRevenueCodes.RowSource <> "qryTransaction_Audit_10_03_Revcode_Exp" Then
26340           .cmbRevenueCodes.RowSource = "qryTransaction_Audit_10_03_Revcode_Exp"
26350         End If
26360       Case False
              ' ** When no choice remains, disable combo box.
26370         .cmbRevenueCodes.Enabled = False
26380         .cmbRevenueCodes.BorderColor = WIN_CLR_DISR
26390         .cmbRevenueCodes.BackStyle = acBackStyleTransparent
26400       End Select
26410       .chkRevcodeType_Income_lbl.FontBold = False
26420     End Select
26430     DoCmd.Hourglass False
26440     DoEvents
26450   End With

EXITP:
26460   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub chkRevcodeType_Income_KeyDown(KeyCode As Integer, Shift As Integer)

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRevcodeType_Income_KeyDown"

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
26590         .chkRevcodeType_Expense.SetFocus
26600       End With
26610     End Select
26620   End If

        ' ** Shift keys.
26630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26650       With Me
26660         intRetVal = 0
26670         .cmbRevenueCodes.SetFocus
26680       End With
26690     End Select
26700   End If

EXITP:
26710   KeyCode = intRetVal
26720   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub chkRevcodeType_Expense_AfterUpdate()

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRevcodeType_Expense_AfterUpdate"

26810   With Me
26820     DoCmd.Hourglass True
26830     DoEvents
26840     Select Case .chkRevcodeType_Expense
          Case True
26850       If .cmbRevenueCodes.Enabled = False Then
26860         .cmbRevenueCodes.Enabled = True
26870         .cmbRevenueCodes.BorderColor = CLR_LTBLU2
26880         .cmbRevenueCodes.BackStyle = acBackStyleNormal
26890       End If
26900       Select Case .chkRevcodeType_Income
            Case True
26910         If .cmbRevenueCodes.RowSource <> "qryTransaction_Audit_10_03_Revcode_Both" Then
26920           .cmbRevenueCodes.RowSource = "qryTransaction_Audit_10_03_Revcode_Both"
26930         End If
26940       Case False
26950         If .cmbRevenueCodes.RowSource <> "qryTransaction_Audit_10_03_Revcode_Exp" Then
26960           .cmbRevenueCodes.RowSource = "qryTransaction_Audit_10_03_Revcode_Exp"
26970         End If
26980       End Select
26990       .chkRevcodeType_Expense_lbl.FontBold = True
27000     Case False
27010       Select Case .chkRevcodeType_Income
            Case True
27020         If .cmbRevenueCodes.Enabled = False Then .cmbRevenueCodes.Enabled = True
27030         If .cmbRevenueCodes.RowSource <> "qryTransaction_Audit_10_03_Revcode_Inc" Then
27040           .cmbRevenueCodes.RowSource = "qryTransaction_Audit_10_03_Revcode_Inc"
27050         End If
27060       Case False
              ' ** When no choice remains, disable combo box.
27070         .cmbRevenueCodes.Enabled = False
27080         .cmbRevenueCodes.BorderColor = WIN_CLR_DISR
27090         .cmbRevenueCodes.BackStyle = acBackStyleTransparent
27100       End Select
27110       .chkRevcodeType_Expense_lbl.FontBold = False
27120     End Select
27130     DoCmd.Hourglass False
27140     DoEvents
27150   End With

EXITP:
27160   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub chkRevcodeType_Expense_KeyDown(KeyCode As Integer, Shift As Integer)

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRevcodeType_Expense_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27220   intShiftDown = (Shift And acShiftMask) > 0
27230   intAltDown = (Shift And acAltMask) > 0
27240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27270       With Me
27280         intRetVal = 0
27290         If .cmbTaxCodes.Enabled = True Then
27300           .cmbTaxCodes.SetFocus
27310         Else

27320         End If
27330       End With
27340     End Select
27350   End If

        ' ** Shift keys.
27360   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27370     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27380       With Me
27390         intRetVal = 0
27400         .chkRevcodeType_Income.SetFocus
27410       End With
27420     End Select
27430   End If

EXITP:
27440   KeyCode = intRetVal
27450   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub cmbTaxCodes_AfterUpdate()

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

27510   With Me
27520     DoCmd.Hourglass True
27530     DoEvents
27540     Set frm = .Parent
27550     If IsNull(.cmbTaxCodes) = False Then
27560       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
27570         strFilter01 = TAX_CODE & CStr(.cmbTaxCodes)
27580       Else
              ' ** There are clauses present.

27590         FilterRec_GetArr  ' ** Procedure: Below.

27600         intPos01 = InStr(strFilter01, TAX_CODE)
27610         If intPos01 = 0& Then
                ' ** This clause isn't present.
27620           intPos03 = 0&: intPos04 = 0&
27630           For lngX = (lngFilts - 1&) To 0& Step -1&
27640             If arr_varFilt(F_CONST, lngX) = TAX_CODE Then
27650               intPos04 = -1
27660             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
27670               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
27680               If intPos03 > 0 Then
27690                 intPos04 = 0&
27700                 Exit For
27710               End If
27720             End If
27730           Next
27740           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
27750             strFilter01 = TAX_CODE & CStr(.cmbTaxCodes) & ANDF & strFilter01
27760           Else
                  ' ** There's a clause before this one.
27770             intPos02 = InStr(intPos03, strFilter01, ANDF)
27780             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
27790               strFilter01 = strFilter01 & ANDF & TAX_CODE & CStr(.cmbTaxCodes)
27800             Else
                    ' ** Add this clause to the middle of the filter.
27810               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & TAX_CODE & CStr(.cmbTaxCodes) & Mid(strFilter01, intPos02)
27820             End If
27830           End If
27840         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
27850           intPos02 = InStr((intPos01 + Len(TAX_CODE) + 1), strFilter01, ANDF)  ' ** Find the next ' And '.
27860           If intPos02 > 0 Then
27870             strFilter01 = Left(strFilter01, (intPos01 - 1)) & TAX_CODE & CStr(.cmbTaxCodes) & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
27880           Else
27890             strFilter01 = Left(strFilter01, (intPos01 - 1)) & TAX_CODE & CStr(.cmbTaxCodes)
27900           End If
27910         End If
27920       End If
27930       strFilter02 = strFilter01
27940       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
27950       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
27960       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
27970       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
27980       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
27990       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
28000     Else
28010       strFilter02 = strFilter01
28020       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
28030       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem TAX_CODE  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
28040       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
28050       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
28060       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
28070       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem TAX_CODE  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
28080       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
28090       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
28100     End If
28110     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .CheckNum.Enabled = True Then
          '  .CheckNum.SetFocus
          'End If
28120     DoCmd.Hourglass False
28130     DoEvents
28140   End With

EXITP:
28150   Set frm = Nothing
28160   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmbTaxCodes_KeyDown(KeyCode As Integer, Shift As Integer)

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28220   intShiftDown = (Shift And acShiftMask) > 0
28230   intAltDown = (Shift And acAltMask) > 0
28240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28270       With Me
28280         intRetVal = 0
28290         If .cmbLocations.Enabled = True Then
28300           .cmbLocations.SetFocus
28310         Else

28320         End If
28330       End With
28340     End Select
28350   End If

        ' ** Shift keys.
28360   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28370     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28380       With Me
28390         intRetVal = 0
28400         If .cmbRevenueCodes.Enabled = True Then
28410           .cmbRevenueCodes.SetFocus
28420         Else

28430         End If
28440       End With
28450     End Select
28460   End If

        ' ** Ctrl keys.
28470   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
28480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28490       With Me
28500         intRetVal = 0
28510         .chkTaxcodeType_Income.SetFocus
28520       End With
28530     End Select
28540   End If

        ' ** Ctrl-Shift keys.
28550   If intCtrlDown And (Not intAltDown) And intShiftDown Then
28560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28570       With Me
28580         intRetVal = 0
28590         If .chkRevcodeType_Expense.Enabled = True Then
28600           .chkRevcodeType_Expense.SetFocus
28610         Else
28620           Beep
28630         End If
28640       End With
28650     End Select
28660   End If

EXITP:
28670   KeyCode = intRetVal
28680   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub chkTaxcodeType_Income_AfterUpdate()

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkTaxcodeType_Income_AfterUpdate"

        Dim frm As Access.Form

28710   With Me
28720     DoCmd.Hourglass True
28730     DoEvents
28740     Select Case .chkTaxcodeType_Income
          Case True
28750       If .cmbTaxCodes.Enabled = False Then
28760         .cmbTaxCodes.Enabled = True
28770         .cmbTaxCodes.BorderColor = CLR_LTBLU2
28780         .cmbTaxCodes.BackStyle = acBackStyleNormal
28790       End If
28800       Select Case .chkTaxcodeType_Deduction
            Case True
28810         If .cmbTaxCodes.RowSource <> "qryTransaction_Audit_10_04_Taxcode_Both" Then
28820           .cmbTaxCodes.RowSource = "qryTransaction_Audit_10_04_Taxcode_Both"
28830         End If
28840       Case False
28850         If .cmbTaxCodes.RowSource <> "qryTransaction_Audit_10_04_Taxcode_Inc" Then
28860           .cmbTaxCodes.RowSource = "qryTransaction_Audit_10_04_Taxcode_Inc"
28870         End If
28880       End Select
28890       .chkTaxcodeType_Income_lbl.FontBold = True
28900     Case False
28910       Select Case .chkTaxcodeType_Deduction
            Case True
28920         If .cmbTaxCodes.Enabled = False Then .cmbTaxCodes.Enabled = True
28930         If .cmbTaxCodes.RowSource <> "qryTransaction_Audit_10_04_Taxcode_Ded" Then
28940           .cmbTaxCodes.RowSource = "qryTransaction_Audit_10_04_Taxcode_Ded"
28950         End If
28960       Case False
              ' ** When no choice remains, disable combo box.
28970         .cmbTaxCodes.Enabled = False
28980         .cmbTaxCodes.BorderColor = WIN_CLR_DISR
28990         .cmbTaxCodes.BackStyle = acBackStyleTransparent
29000       End Select
29010       .chkTaxcodeType_Income_lbl.FontBold = False
29020     End Select
29030     DoCmd.Hourglass False
29040     DoEvents
29050   End With

EXITP:
29060   Set frm = Nothing
29070   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub chkTaxcodeType_Income_KeyDown(KeyCode As Integer, Shift As Integer)

29100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkTaxcodeType_Income_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29120   intShiftDown = (Shift And acShiftMask) > 0
29130   intAltDown = (Shift And acAltMask) > 0
29140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29170       With Me
29180         intRetVal = 0
29190         .chkTaxcodeType_Deduction.SetFocus
29200       End With
29210     End Select
29220   End If

        ' ** Shift keys.
29230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29250       With Me
29260         intRetVal = 0
29270         .cmbTaxCodes.SetFocus
29280       End With
29290     End Select
29300   End If

EXITP:
29310   KeyCode = intRetVal
29320   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub chkTaxcodeType_Deduction_AfterUpdate()

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkTaxcodeType_Deduction_AfterUpdate"

        Dim frm As Access.Form

29410   With Me
29420     DoCmd.Hourglass True
29430     DoEvents
29440     Select Case .chkTaxcodeType_Deduction
          Case True
29450       If .cmbTaxCodes.Enabled = False Then
29460         .cmbTaxCodes.Enabled = True
29470         .cmbTaxCodes.BorderColor = CLR_LTBLU2
29480         .cmbTaxCodes.BackStyle = acBackStyleNormal
29490       End If
29500       Select Case .chkTaxcodeType_Income
            Case True
29510         If .cmbTaxCodes.RowSource <> "qryTransaction_Audit_10_04_Taxcode_Both" Then
29520           .cmbTaxCodes.RowSource = "qryTransaction_Audit_10_04_Taxcode_Both"
29530         End If
29540       Case False
29550         If .cmbTaxCodes.RowSource <> "qryTransaction_Audit_10_04_Taxcode_Ded" Then
29560           .cmbTaxCodes.RowSource = "qryTransaction_Audit_10_04_Taxcode_Ded"
29570         End If
29580       End Select
29590       .chkTaxcodeType_Deduction_lbl.FontBold = True
29600     Case False
29610       Select Case .chkTaxcodeType_Income
            Case True
29620         If .cmbTaxCodes.Enabled = False Then .cmbTaxCodes.Enabled = True
29630         If .cmbTaxCodes.RowSource <> "qryTransaction_Audit_10_04_Taxcode_Inc" Then
29640           .cmbTaxCodes.RowSource = "qryTransaction_Audit_10_04_Taxcode_Inc"
29650         End If
29660       Case False
              ' ** When no choice remains, disable combo box.
29670         .cmbTaxCodes.Enabled = False
29680         .cmbTaxCodes.BorderColor = WIN_CLR_DISR
29690         .cmbTaxCodes.BackStyle = acBackStyleTransparent
29700       End Select
29710       .chkTaxcodeType_Deduction_lbl.FontBold = False
29720     End Select
29730     DoCmd.Hourglass False
29740     DoEvents
29750   End With

EXITP:
29760   Set frm = Nothing
29770   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub chkTaxcodeType_Deduction_KeyDown(KeyCode As Integer, Shift As Integer)

29800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkTaxcodeType_Deduction_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29820   intShiftDown = (Shift And acShiftMask) > 0
29830   intAltDown = (Shift And acAltMask) > 0
29840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29870       With Me
29880         intRetVal = 0
29890         If .cmbLocations.Enabled = True Then
29900           .cmbLocations.SetFocus
29910         Else

29920         End If
29930       End With
29940     End Select
29950   End If

        ' ** Shift keys.
29960   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29970     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29980       With Me
29990         intRetVal = 0
30000         .chkTaxcodeType_Income.SetFocus
30010       End With
30020     End Select
30030   End If

EXITP:
30040   KeyCode = intRetVal
30050   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub cmbLocations_AfterUpdate()

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngTmp01 As Long
        Dim lngX As Long

30110   With Me
30120     DoCmd.Hourglass True
30130     DoEvents
30140     Set frm = .Parent
30150     If IsNull(.cmbLocations) = False Then
30160       lngTmp01 = .cmbLocations
30170       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
30180         strFilter01 = LOC_NUM & CStr(lngTmp01)
30190       Else
              ' ** There are other clauses present.
30200         FilterRec_GetArr  ' ** Procedure: Below.
30210         intPos01 = InStr(strFilter01, LOC_NUM)
30220         If intPos01 = 0& Then
                ' ** This clause isn't present.
30230           intPos03 = 0&: intPos04 = 0&
30240           For lngX = (lngFilts - 1&) To 0& Step -1&
30250             If arr_varFilt(F_CONST, lngX) = LOC_NUM Then
30260               intPos04 = -1
30270             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
30280               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
30290               If intPos03 > 0 Then
30300                 intPos04 = 0&
30310                 Exit For
30320               End If
30330             End If
30340           Next
30350           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
30360             strFilter01 = LOC_NUM & CStr(lngTmp01) & ANDF & strFilter01
30370           Else
                  ' ** There's a clause before this one.
30380             intPos02 = InStr(intPos03, strFilter01, ANDF)
30390             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
30400               strFilter01 = strFilter01 & ANDF & LOC_NUM & CStr(lngTmp01)
30410             Else
                    ' ** Add this clause to the middle of the filter.
30420               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & LOC_NUM & CStr(lngTmp01) & Mid(strFilter01, intPos02)
30430             End If
30440           End If
30450         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
30460           intPos02 = InStr((intPos01 + Len(LOC_NUM) + 1), strFilter01, "'")  ' ** Find the closing paren.
30470           If intPos02 > 0 Then
30480             strFilter01 = Left(strFilter01, (intPos01 - 1)) & LOC_NUM & CStr(lngTmp01) & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
30490           Else
30500             strFilter01 = Left(strFilter01, (intPos01 - 1)) & LOC_NUM & CStr(lngTmp01)
30510           End If
30520         End If
30530       End If
30540       strFilter02 = strFilter01
30550       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
30560       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
30570       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
30580       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
30590       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
30600       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
30610     Else
30620       strFilter02 = strFilter01
30630       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
30640       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem LOC_NUM  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
30650       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
30660       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
30670       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
30680       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem LOC_NUM  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
30690       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
30700       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
30710     End If
30720     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
30730     DoCmd.Hourglass False
30740     DoEvents
30750   End With

EXITP:
30760   Set frm = Nothing
30770   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmbLocations_KeyDown(KeyCode As Integer, Shift As Integer)

30800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30820   intShiftDown = (Shift And acShiftMask) > 0
30830   intAltDown = (Shift And acAltMask) > 0
30840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30870       With Me
30880         intRetVal = 0
30890         If .CheckNum.Enabled = True Then
30900           .CheckNum.SetFocus
30910         Else

30920         End If
30930       End With
30940     End Select
30950   End If

        ' ** Shift keys.
30960   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30970     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30980       With Me
30990         intRetVal = 0
31000         If .cmbTaxCodes.Enabled = True Then
31010           .cmbTaxCodes.SetFocus
31020         Else

31030         End If
31040       End With
31050     End Select
31060   End If

        ' ** Ctrl-Shift keys.
31070   If intCtrlDown And (Not intAltDown) And intShiftDown Then
31080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31090       With Me
31100         intRetVal = 0
31110         If .chkTaxcodeType_Deduction.Enabled = True Then
31120           .chkTaxcodeType_Deduction.SetFocus
31130         Else
31140           Beep
31150         End If
31160       End With
31170     End Select
31180   End If

EXITP:
31190   KeyCode = intRetVal
31200   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub CheckNum_AfterUpdate()

31300 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
        Dim lngX As Long

31310   With Me
31320     DoCmd.Hourglass True
31330     DoEvents
31340     Set frm = .Parent
31350     If IsNull(.CheckNum) = False Then
31360       strTmp01 = Trim(.CheckNum)
31370       intPos01 = InStr(strTmp01, "-")
31380       If intPos01 > 0 Then
31390         If Left(strTmp01, 1) <> "-" And Right(strTmp01, 1) <> "-" Then
                ' ** A range of numbers.
31400           strTmp02 = Trim(Mid(strTmp01, (intPos01 + 1)))
31410           strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
31420         End If
31430       End If
31440       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
31450         If intPos01 = 0 Then
                ' ** Simple equality.
31460           strFilter01 = CHK_NUM & CStr(strTmp01)
31470         Else
31480           If Right(strTmp01, 1) = "-" Then
                  ' ** Greater than.
31490             strFilter01 = CHK_NUM1 & Rem_Dash(strTmp01)  ' ** Module Function: modStringFuncs.
31500           ElseIf Left(strTmp01, 1) = "-" Then
                  ' ** Less than.
31510             strFilter01 = CHK_NUM2 & Rem_Dash(strTmp01)  ' ** Module Function: modStringFuncs.
31520           Else
                  ' ** A range of numbers.
31530             strFilter01 = "(" & CHK_NUM1 & strTmp01 & ANDF & CHK_NUM2 & strTmp02 & ")"
31540           End If
31550         End If
31560       Else
              ' ** There are clauses present.

31570         FilterRec_GetArr  ' ** Procedure: Below.

31580         intPos01 = InStr(strFilter01, CHK_NUM)
31590         intPos02 = InStr(strFilter01, CHK_NUM1)
31600         intPos03 = InStr(strFilter01, CHK_NUM2)
31610         If intPos01 = 0& And intPos02 = 0& And intPos03 = 0& Then
                ' ** This clause isn't present.
31620           intPos03 = 0&: intPos04 = 0&
31630           For lngX = (lngFilts - 1&) To 0& Step -1&
31640             If arr_varFilt(F_CONST, lngX) = CHK_NUM Then  ' ** Only the 1st one is present.
31650               intPos04 = -1
31660             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
31670               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
31680               If intPos03 > 0 Then
31690                 intPos04 = 0&
31700                 Exit For
31710               End If
31720             End If
31730           Next
31740           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
31750             If strTmp01 <> vbNullString And strTmp02 <> vbNullString Then
                    ' ** A range of numbers.
31760               strFilter01 = CHK_NUM1 & strTmp01 & ANDF & CHK_NUM2 & strTmp02 & ANDF & strFilter01
31770             ElseIf Right(strTmp01, 1) = "-" Then
                    ' ** Greater than.
31780               strTmp01 = Rem_Dash(strTmp01)  ' ** Module Function: modStringFuncs.
31790               strFilter01 = CHK_NUM1 & strTmp01 & ANDF & strFilter01
31800             ElseIf Left(strTmp01, 1) = "-" Then
                    ' ** Less than.
31810               strTmp01 = Rem_Dash(strTmp01)  ' ** Module Function: modStringFuncs.
31820               strFilter01 = CHK_NUM2 & strTmp01 & ANDF & strFilter01
31830             Else
                    ' ** Simple equality.
31840               strFilter01 = CHK_NUM & strTmp01 & ANDF & strFilter01
31850             End If
31860           Else
                  ' ** There's a clause before this one.
31870             intPos02 = InStr(intPos03, strFilter01, ANDF)
31880             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
31890               If strTmp01 <> vbNullString And strTmp02 <> vbNullString Then
                      ' ** A range of numbers.
31900                 strFilter01 = strFilter01 & ANDF & CHK_NUM1 & strTmp01 & ANDF & CHK_NUM2 & strTmp02
31910               ElseIf Right(strTmp01, 1) = "-" Then
                      ' ** Greater than.
31920                 strTmp01 = Rem_Dash(strTmp01)  ' ** Module Function: modStringFuncs.
31930                 strFilter01 = strFilter01 & ANDF & CHK_NUM1 & strTmp01
31940               ElseIf Left(strTmp01, 1) = "-" Then
                      ' ** Less than.
31950                 strTmp01 = Rem_Dash(strTmp01)  ' ** Module Function: modStringFuncs.
31960                 strFilter01 = strFilter01 & ANDF & CHK_NUM2 & strTmp01
31970               Else
                      ' ** Simple equality.
31980                 strFilter01 = strFilter01 & ANDF & CHK_NUM & strTmp01
31990               End If
32000             Else
                    ' ** Add this clause to the middle of the filter.
32010               If strTmp01 <> vbNullString And strTmp02 <> vbNullString Then
                      ' ** A range of numbers.
32020                 strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & CHK_NUM1 & strTmp01 & ANDF & CHK_NUM2 & strTmp02 & _
                        Mid(strFilter01, intPos02)
32030               ElseIf Right(strTmp01, 1) = "-" Then
                      ' ** Greater than.
32040                 strTmp01 = Rem_Dash(strTmp01)  ' ** Module Function: modStringFuncs.
32050                 strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & CHK_NUM1 & strTmp01 & Mid(strFilter01, intPos02)
32060               ElseIf Left(strTmp01, 1) = "-" Then
                      ' ** Less than.
32070                 strTmp01 = Rem_Dash(strTmp01)  ' ** Module Function: modStringFuncs.
32080                 strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & CHK_NUM2 & strTmp01 & Mid(strFilter01, intPos02)
32090               Else
                      ' ** Simple equality.
32100                 strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & CHK_NUM & strTmp01 & Mid(strFilter01, intPos02)
32110               End If
32120             End If
32130           End If
32140         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
                'REMEMBER, THESE POSITIONS ARE FOR AN EXISTING CLAUSE,
                'AND SAY NOTHING ABOUT WHAT'S TO REPLACE IT!
32150           strTmp03 = vbNullString: strTmp04 = vbNullString
                ' ** strTmp01 and/or strTmp02 have the new input.
32160           If intPos01 > 0 Then
                  ' ** Simple equality.
32170             intPos04 = InStr((intPos01 + Len(CHK_NUM) + 1), strFilter01, ANDF)  ' ** Find the next ' And '.
32180             strTmp03 = Left(strFilter01, (intPos01 - 1))
32190             If intPos04 > 0 Then
32200               strTmp04 = Mid(strFilter01, intPos04)
32210             End If
32220           ElseIf intPos02 > 0 And intPos03 > 0 Then
                  ' ** Range of numbers.
32230             intPos04 = InStr((intPos03 + Len(CHK_NUM2) + 1), strFilter01, ANDF)
32240             strTmp03 = Left(strFilter01, (intPos02 - 1))
32250             If intPos04 > 0 Then
32260               strTmp04 = Mid(strFilter01, intPos04)
32270             End If
32280           ElseIf intPos02 > 0 Then
                  ' ** Greater than.
32290             intPos04 = InStr((intPos02 + Len(CHK_NUM1) + 1), strFilter01, ANDF)
32300             strTmp03 = Left(strFilter01, (intPos02 - 1))
32310             If intPos04 > 0 Then
32320               strTmp04 = Mid(strFilter01, intPos04)
32330             End If
32340           ElseIf intPos03 > 0 Then
                  ' ** Less than.
32350             intPos04 = InStr((intPos03 + Len(CHK_NUM2) + 1), strFilter01, ANDF)
32360             strTmp03 = Left(strFilter01, (intPos03 - 1))
32370             If intPos04 > 0 Then
32380               strTmp04 = Mid(strFilter01, intPos04)
32390             End If
32400           End If
                ' ** It doesn't matter whether strTmp04 has anything or not.
32410           If strTmp01 <> vbNullString And strTmp02 <> vbNullString Then
                  ' ** A range of numbers
32420             strFilter01 = strTmp03 & ANDF & CHK_NUM1 & strTmp01 & ANDF & CHK_NUM2 & strTmp02 & strTmp04
32430           ElseIf Right(strTmp01, 1) = "-" Then
                  ' ** Greater than.
32440             strTmp01 = Rem_Dash(strTmp01)  ' ** Module Function: modStringFuncs.
32450             strFilter01 = strTmp03 & ANDF & CHK_NUM1 & strTmp01 & strTmp04
32460           ElseIf Left(strTmp01, 1) = "-" Then
                  ' ** Less than.
32470             strTmp01 = Rem_Dash(strTmp01)  ' ** Module Function: modStringFuncs.
32480             strFilter01 = strTmp03 & ANDF & CHK_NUM2 & strTmp01 & strTmp04
32490           Else
                  ' ** Simple equality.
32500             strFilter01 = strTmp03 & ANDF & CHK_NUM & strTmp01 & strTmp04
32510           End If
32520         End If
32530       End If
32540       strFilter02 = strFilter01
32550       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
32560       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
32570       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
32580       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
32590       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
32600       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
32610     Else
32620       strFilter02 = strFilter01
32630       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
32640       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem CHK_NUM  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
32650       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
32660       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
32670       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
32680       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem CHK_NUM  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
32690       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
32700       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
32710     End If
32720     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .cmbUsers.Enabled = True Then
          '  .cmbUsers.SetFocus
          'End If
32730     DoCmd.Hourglass False
32740     DoEvents
32750   End With

EXITP:
32760   Set frm = Nothing
32770   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub cmbUsers_AfterUpdate()

32800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbUsers_AfterUpdate"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

32810   With Me
32820     DoCmd.Hourglass True
32830     DoEvents
32840     Set frm = .Parent
32850     If IsNull(.cmbUsers) = False Then
32860       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
32870         strFilter01 = JRNL_USER & .cmbUsers & "'"
32880       Else
              ' ** There are clauses present.

32890         FilterRec_GetArr  ' ** Procedure: Below.

32900         intPos01 = InStr(strFilter01, JRNL_USER)
32910         If intPos01 = 0& Then
                ' ** This clause isn't present.
32920           intPos03 = 0&: intPos04 = 0&
32930           For lngX = (lngFilts - 1&) To 0& Step -1&
32940             If arr_varFilt(F_CONST, lngX) = JRNL_USER Then
32950               intPos04 = -1
32960             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
32970               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
32980               If intPos03 > 0 Then
32990                 intPos04 = 0&
33000                 Exit For
33010               End If
33020             End If
33030           Next
33040           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
33050             strFilter01 = JRNL_USER & .cmbUsers & "'" & ANDF & strFilter01
33060           Else
                  ' ** There's a clause before this one.
33070             intPos02 = InStr(intPos03, strFilter01, ANDF)
33080             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
33090               strFilter01 = strFilter01 & ANDF & JRNL_USER & .cmbUsers & "'"
33100             Else
                    ' ** Add this clause to the middle of the filter.
33110               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & JRNL_USER & .cmbUsers & "'" & Mid(strFilter01, intPos02)
33120             End If
33130           End If
33140         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
33150           intPos02 = InStr((intPos01 + Len(JRNL_USER) + 1), strFilter01, "'")  ' ** Find the closing paren.
33160           If intPos02 > 0 Then
33170             strFilter01 = Left(strFilter01, (intPos01 - 1)) & JRNL_USER & .cmbUsers & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
33180           Else
33190             strFilter01 = Left(strFilter01, (intPos01 - 1)) & JRNL_USER & .cmbUsers & "'"
33200           End If
33210         End If
33220       End If
33230       strFilter02 = strFilter01
33240       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
33250       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
33260       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
33270       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
33280       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
33290       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
33300     Else
33310       strFilter02 = strFilter01
33320       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
33330       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem JRNL_USER  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
33340       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
33350       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
33360       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
33370       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem JRNL_USER  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
33380       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
33390       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
33400     End If
33410     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .PostedDateStart.Enabled = True Then
          '  .PostedDateStart.SetFocus
          'End If
33420     DoCmd.Hourglass False
33430     DoEvents
33440   End With

EXITP:
33450   Set frm = Nothing
33460   Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub PostedDateStart_BeforeUpdate(Cancel As Integer)

33500 On Error GoTo ERRH

        Const THIS_PROC As String = "PostedDateStart_BeforeUpdate"

33510   With Me
33520     If .PostedDateStart > Date Then
33530       MsgBox "Please enter a date prior to today.", vbInformation + vbOKOnly, "Invalid Date"
33540       Cancel = -1
33550     End If
33560   End With

EXITP:
33570   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub PostedDateStart_AfterUpdate()

33600 On Error GoTo ERRH

        Const THIS_PROC As String = "PostedDateStart_AfterUpdate"

33610   PostedDateStartAfterUpdate_TA strFilter01, strFilter02, dblFilterRecs, rstAll1, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
33620   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub PostedDateStart_KeyDown(KeyCode As Integer, Shift As Integer)

33700 On Error GoTo ERRH

        Const THIS_PROC As String = "PostedDateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

33710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
33720   intShiftDown = (Shift And acShiftMask) > 0
33730   intAltDown = (Shift And acAltMask) > 0
33740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
33750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
33760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33770       With Me
33780         intRetVal = 0
33790         .PostedDateEnd.SetFocus
33800       End With
33810     End Select
33820   End If

        ' ** Shift keys.
33830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33850       With Me
33860         intRetVal = 0
33870         If .cmbUsers.Enabled = True Then
33880           .cmbUsers.SetFocus
33890         Else

33900         End If
33910       End With
33920     End Select
33930   End If

        ' ** Ctrl keys.
33940   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
33950     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33960       With Me
33970         intRetVal = 0
33980         .cmdCalendar7.SetFocus
33990       End With
34000     End Select
34010   End If

EXITP:
34020   KeyCode = intRetVal
34030   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar7_GotFocus()

34100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar7_GotFocus"

34110   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
34120   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

34200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar7_MouseDown"

34210   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
34220   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmdCalendar7_Click()

34300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar7_Click"

34310   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
34320   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar7_KeyDown(KeyCode As Integer, Shift As Integer)

34400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar7_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

34410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
34420   intShiftDown = (Shift And acShiftMask) > 0
34430   intAltDown = (Shift And acAltMask) > 0
34440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
34450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34460     Select Case intRetVal
          Case vbKeyTab
34470       With Me
34480         intRetVal = 0
34490         .PostedDateEnd.SetFocus
34500       End With
34510     End Select
34520   End If

        ' ** Shift keys.
34530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34540     Select Case intRetVal
          Case vbKeyTab
34550       With Me
34560         intRetVal = 0
34570         .PostedDateStart.SetFocus
34580       End With
34590     End Select
34600   End If

EXITP:
34610   KeyCode = intRetVal
34620   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

34700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar7_MouseMove"

34710   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
34720   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar7_MouseUp"

34810   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
34820   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar7_LostFocus()

34900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar7_LostFocus"

34910   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
34920   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub PostedDateEnd_BeforeUpdate(Cancel As Integer)

35000 On Error GoTo ERRH

        Const THIS_PROC As String = "PostedDateEnd_BeforeUpdate"

35010   With Me
35020     If IsNull(.PostedDateStart) = False Then
35030       If .PostedDateEnd < .PostedDateStart Then
35040         MsgBox "Please enter a date the same as or after the From date.", vbInformation + vbOKOnly, "Invalid Date"
35050         Cancel = -1
35060       End If
35070     End If
35080   End With

EXITP:
35090   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub PostedDateEnd_AfterUpdate()

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "PostedDateEnd_AfterUpdate"

35110   PostedDateEndAfterUpdate_TA strFilter01, strFilter02, dblFilterRecs, rstAll1, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
35120   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub PostedDateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

35200 On Error GoTo ERRH

        Const THIS_PROC As String = "PostedDateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

35210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
35220   intShiftDown = (Shift And acShiftMask) > 0
35230   intAltDown = (Shift And acAltMask) > 0
35240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
35250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
35260     Select Case intRetVal
          Case vbKeyReturn  ' ** Let vbKeyTab be handled normally.
35270       With Me
35280         If IsNull(.PostedDateEnd) Then
35290           If .PostedDateEnd.text = vbNullString Then
                  ' ** Populate it with today's date.
35300             .PostedDateEnd = CDate(Format(Date, "mm/dd/yyyy") & " 11:59:59 PM")  ' ** Since this is Posted, make sure it's through Midnight.
35310             PostedDateEnd_AfterUpdate  ' ** Procedure: Below.
35320           End If
35330         End If
35340       End With
35350     End Select
35360   End If

        ' ** Ctrl keys.
35370   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
35380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
35390       With Me
35400         intRetVal = 0
35410         .cmdCalendar8.SetFocus
35420       End With
35430     End Select
35440   End If

        ' ** Ctrl-Shift keys.
35450   If intCtrlDown And (Not intAltDown) And intShiftDown Then
35460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
35470       With Me
35480         intRetVal = 0
35490         .cmdCalendar7.SetFocus
35500       End With
35510     End Select
35520   End If

EXITP:
35530   KeyCode = intRetVal
35540   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar8_GotFocus()

35600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar8_GotFocus"

35610   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
35620   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

35700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar8_MouseDown"

35710   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
35720   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub cmdCalendar8_Click()

35800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar8_Click"

35810   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
35820   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar8_KeyDown(KeyCode As Integer, Shift As Integer)

35900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar8_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

35910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
35920   intShiftDown = (Shift And acShiftMask) > 0
35930   intAltDown = (Shift And acAltMask) > 0
35940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
35950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
35960     Select Case intRetVal
          Case vbKeyTab
35970       With Me
35980         intRetVal = 0
35990         If .opgHidden.Enabled = True Then
36000           .opgHidden.SetFocus
36010         Else

36020         End If
36030       End With
36040     End Select
36050   End If

        ' ** Shift keys.
36060   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
36070     Select Case intRetVal
          Case vbKeyTab
36080       With Me
36090         intRetVal = 0
36100         .PostedDateEnd.SetFocus
36110       End With
36120     End Select
36130   End If

EXITP:
36140   KeyCode = intRetVal
36150   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub cmdCalendar8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

36200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar8_MouseMove"

36210   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
36220   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

36300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar8_MouseUp"

36310   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
36320   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub cmdCalendar8_LostFocus()

36400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar8_LostFocus"

36410   Calendar_Handler THIS_PROC  ' ** Procedure: Below.

EXITP:
36420   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub opgHidden_AfterUpdate()

36500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgHidden_AfterUpdate"

36510   opgHiddenAfterUpdate_TA strFilter01, strFilter02, dblFilterRecs, rstAll1, Me

EXITP:
36520   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub opgHidden_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

36600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgHidden_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

36610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
36620   intShiftDown = (Shift And acShiftMask) > 0
36630   intAltDown = (Shift And acAltMask) > 0
36640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
36650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
36660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
36670       With Me
36680         intRetVal = 0
36690         DoCmd.SelectObject acForm, .Parent.Name, False
36700         Select Case .Parent.opgView
              Case .Parent.opgView_optForm.OptionValue
36710           .Parent.frmTransaction_Audit_Sub.SetFocus
36720         Case .Parent.opgView_optDatasheet.OptionValue
36730           .Parent.frmTransaction_Audit_Sub_ds.SetFocus
36740         End Select
36750       End With
36760     End Select
36770   End If

        ' ** Shift keys.
36780   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
36790     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
36800       With Me
36810         intRetVal = 0
36820         If .PostedDateEnd.Enabled = True Then
36830           .PostedDateEnd.SetFocus
36840         ElseIf .cmbUsers.Enabled = True Then
36850           .cmbUsers.SetFocus
36860         ElseIf .CheckNum.Enabled = True Then
36870           .CheckNum.SetFocus
36880         ElseIf .chkTaxcodeType_Deduction.Enabled = True Then
36890           .chkTaxcodeType_Deduction.SetFocus
36900         ElseIf .cmbTaxCodes.Enabled = True Then
36910           .cmbTaxCodes.SetFocus
36920         ElseIf .chkRevcodeType_Expense.Enabled = True Then
36930           .chkRevcodeType_Expense.SetFocus
36940         ElseIf .cmbRevenueCodes.Enabled = True Then
36950           .cmbRevenueCodes.SetFocus
36960         ElseIf .cmbRecurringItems.Enabled = True Then
36970           .cmbRecurringItems.SetFocus
36980         ElseIf .ledger_description.Enabled = True Then
36990           .ledger_description.SetFocus
37000         ElseIf .PurchaseDateEnd.Enabled = True Then
37010           .PurchaseDateEnd.SetFocus
37020         ElseIf .AssetDateEnd.Enabled = True Then
37030           .AssetDateEnd.SetFocus
37040         ElseIf .cmbCurrencies.Enabled = True Then
37050           .cmbCurrencies.SetFocus
37060         ElseIf .cmbAssets.Enabled = True Then
37070           .cmbAssets.SetFocus
37080         ElseIf .cmbAccounts.Enabled = True Then
37090           .cmbAccounts.SetFocus
37100         ElseIf .TransDateStart.Enabled = True Then
37110           .TransDateStart.SetFocus
37120         ElseIf .cmbJournalType1.Enabled = True Then
37130           .cmbJournalType1.SetFocus
37140         ElseIf .journalno.Enabled = True Then
37150           .journalno.SetFocus
37160         Else
37170           DoCmd.SelectObject acForm, .Parent.Name, False
37180           .Parent.opgView.SetFocus
37190         End If
37200       End With
37210     End Select
37220   End If

        ' ** Ctrl-Shift keys.
37230   If intCtrlDown And (Not intAltDown) And intShiftDown Then
37240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
37250       With Me
37260         intRetVal = 0
37270         If .cmdCalendar8.Enabled = True Then
37280           .cmdCalendar8.SetFocus
37290         Else
37300           Beep
37310         End If
37320       End With
37330     End Select
37340   End If

EXITP:
37350   KeyCode = intRetVal
37360   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub opgHidden_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

37400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgHidden_optExclude_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
37470       With Me
37480         intRetVal = 0
37490         DoCmd.SelectObject acForm, .Parent.Name, False
37500         Select Case .Parent.opgView
              Case .Parent.opgView_optForm.OptionValue
37510           .Parent.frmTransaction_Audit_Sub.SetFocus
37520         Case .Parent.opgView_optDatasheet.OptionValue
37530           .Parent.frmTransaction_Audit_Sub_ds.SetFocus
37540         End Select
37550       End With
37560     End Select
37570   End If

        ' ** Shift keys.
37580   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
37590     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
37600       With Me
37610         intRetVal = 0
37620         If .PostedDateEnd.Enabled = True Then
37630           .PostedDateEnd.SetFocus
37640         ElseIf .cmbUsers.Enabled = True Then
37650           .cmbUsers.SetFocus
37660         ElseIf .CheckNum.Enabled = True Then
37670           .CheckNum.SetFocus
37680         ElseIf .chkTaxcodeType_Deduction.Enabled = True Then
37690           .chkTaxcodeType_Deduction.SetFocus
37700         ElseIf .cmbTaxCodes.Enabled = True Then
37710           .cmbTaxCodes.SetFocus
37720         ElseIf .chkRevcodeType_Expense.Enabled = True Then
37730           .chkRevcodeType_Expense.SetFocus
37740         ElseIf .cmbRevenueCodes.Enabled = True Then
37750           .cmbRevenueCodes.SetFocus
37760         ElseIf .cmbRecurringItems.Enabled = True Then
37770           .cmbRecurringItems.SetFocus
37780         ElseIf .ledger_description.Enabled = True Then
37790           .ledger_description.SetFocus
37800         ElseIf .PurchaseDateEnd.Enabled = True Then
37810           .PurchaseDateEnd.SetFocus
37820         ElseIf .AssetDateEnd.Enabled = True Then
37830           .AssetDateEnd.SetFocus
37840         ElseIf .cmbCurrencies.Enabled = True Then
37850           .cmbCurrencies.SetFocus
37860         ElseIf .cmbAssets.Enabled = True Then
37870           .cmbAssets.SetFocus
37880         ElseIf .cmbAccounts.Enabled = True Then
37890           .cmbAccounts.SetFocus
37900         ElseIf .TransDateStart.Enabled = True Then
37910           .TransDateStart.SetFocus
37920         ElseIf .cmbJournalType1.Enabled = True Then
37930           .cmbJournalType1.SetFocus
37940         ElseIf .journalno.Enabled = True Then
37950           .journalno.SetFocus
37960         Else
37970           DoCmd.SelectObject acForm, .Parent.Name, False
37980           .Parent.opgView.SetFocus
37990         End If
38000       End With
38010     End Select
38020   End If

        ' ** Ctrl-Shift keys.
38030   If intCtrlDown And (Not intAltDown) And intShiftDown Then
38040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
38050       With Me
38060         intRetVal = 0
38070         If .cmdCalendar8.Enabled = True Then
38080           .cmdCalendar8.SetFocus
38090         Else
38100           Beep
38110         End If
38120       End With
38130     End Select
38140   End If

EXITP:
38150   KeyCode = intRetVal
38160   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub opgHidden_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

38200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgHidden_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

38210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
38220   intShiftDown = (Shift And acShiftMask) > 0
38230   intAltDown = (Shift And acAltMask) > 0
38240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
38250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
38260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
38270       With Me
38280         intRetVal = 0
38290         DoCmd.SelectObject acForm, .Parent.Name, False
38300         Select Case .Parent.opgView
              Case .Parent.opgView_optForm.OptionValue
38310           .Parent.frmTransaction_Audit_Sub.SetFocus
38320         Case .Parent.opgView_optDatasheet.OptionValue
38330           .Parent.frmTransaction_Audit_Sub_ds.SetFocus
38340         End Select
38350       End With
38360     End Select
38370   End If

        ' ** Shift keys.
38380   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
38390     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
38400       With Me
38410         intRetVal = 0
38420         If .PostedDateEnd.Enabled = True Then
38430           .PostedDateEnd.SetFocus
38440         ElseIf .cmbUsers.Enabled = True Then
38450           .cmbUsers.SetFocus
38460         ElseIf .CheckNum.Enabled = True Then
38470           .CheckNum.SetFocus
38480         ElseIf .chkTaxcodeType_Deduction.Enabled = True Then
38490           .chkTaxcodeType_Deduction.SetFocus
38500         ElseIf .cmbTaxCodes.Enabled = True Then
38510           .cmbTaxCodes.SetFocus
38520         ElseIf .chkRevcodeType_Expense.Enabled = True Then
38530           .chkRevcodeType_Expense.SetFocus
38540         ElseIf .cmbRevenueCodes.Enabled = True Then
38550           .cmbRevenueCodes.SetFocus
38560         ElseIf .cmbRecurringItems.Enabled = True Then
38570           .cmbRecurringItems.SetFocus
38580         ElseIf .ledger_description.Enabled = True Then
38590           .ledger_description.SetFocus
38600         ElseIf .PurchaseDateEnd.Enabled = True Then
38610           .PurchaseDateEnd.SetFocus
38620         ElseIf .AssetDateEnd.Enabled = True Then
38630           .AssetDateEnd.SetFocus
38640         ElseIf .cmbCurrencies.Enabled = True Then
38650           .cmbCurrencies.SetFocus
38660         ElseIf .cmbAssets.Enabled = True Then
38670           .cmbAssets.SetFocus
38680         ElseIf .cmbAccounts.Enabled = True Then
38690           .cmbAccounts.SetFocus
38700         ElseIf .TransDateStart.Enabled = True Then
38710           .TransDateStart.SetFocus
38720         ElseIf .cmbJournalType1.Enabled = True Then
38730           .cmbJournalType1.SetFocus
38740         ElseIf .journalno.Enabled = True Then
38750           .journalno.SetFocus
38760         Else
38770           DoCmd.SelectObject acForm, .Parent.Name, False
38780           .Parent.opgView.SetFocus
38790         End If
38800       End With
38810     End Select
38820   End If

        ' ** Ctrl-Shift keys.
38830   If intCtrlDown And (Not intAltDown) And intShiftDown Then
38840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
38850       With Me
38860         intRetVal = 0
38870         If .cmdCalendar8.Enabled = True Then
38880           .cmdCalendar8.SetFocus
38890         Else
38900           Beep
38910         End If
38920       End With
38930     End Select
38940   End If

EXITP:
38950   KeyCode = intRetVal
38960   Exit Sub

ERRH:
340     intRetVal = 0
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Public Sub RstAllSet(strQryName As String, Optional varMode As Variant)

39000 On Error GoTo ERRH

        Const THIS_PROC As String = "RstAllSet"

39010   Select Case IsMissing(varMode)
        Case True
39020     RstAllSet_TA rstAll1, rstAll2, strQryName  ' ** Module Procedure: modTransactionAuditFuncs1.
39030   Case False
39040     RstAllSet_TA rstAll1, rstAll2, strQryName, varMode  ' ** Module Procedure: modTransactionAuditFuncs1.
39050   End Select

EXITP:
39060   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub FilterRec_GetArr()
' ** Note: FilterRec vs. FilterRecs.

39100 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRec_GetArr"

39110   If lngFilts = 0& Or lngFilts_ds = 0& Then
39120     arr_varFilt = Empty
39130     arr_varFilt = FilterRecs_GetArr(1)  ' ** Module Function: modTransactionAuditFuncs1.
39140     If IsEmpty(arr_varFilt) = True Then
39150       FilterRecs_Load  ' ** Module Function: modTransactionAuditFuncs1.
39160       DoEvents
39170       arr_varFilt = FilterRecs_GetArr(1)  ' ** Module Function: modTransactionAuditFuncs1.
39180     End If
39190     lngFilts = (UBound(arr_varFilt, 2) + 1&)
39200     arr_varFilt_ds = Empty
39210     arr_varFilt_ds = FilterRecs_GetArr(2)  ' ** Module Function: modTransactionAuditFuncs1.
39220     lngFilts_ds = (UBound(arr_varFilt_ds, 2) + 1&)
39230   End If

EXITP:
39240   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub FilterRecs_Set(strFltr As String, dblFltrRecs As Double, Optional varMode As Variant)
' ** Initialize the local strFilter01 and dblFilterRecs with frmTransaction_Audit_Sub's copy.

39300 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Set"

        Dim intMode As Integer

39310   Select Case IsMissing(varMode)
        Case True
39320     intMode = 1
39330   Case False
39340     intMode = varMode
39350   End Select
39360   Select Case intMode
        Case 1
39370     strFilter01 = strFltr
39380   Case 2
39390     strFilter02 = strFltr
39400   End Select
39410   dblFilterRecs = dblFltrRecs

EXITP:
39420   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Public Sub Calendar_Handler(strProc As String, Optional varAble As Variant)

39500 On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Handler"

39510   Select Case IsMissing(varAble)
        Case True
39520     Calendar_Handler_TA strProc, lngCals, arr_varCal, clsMonthClass, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
39530   Case False
39540     Calendar_Handler_TA strProc, lngCals, arr_varCal, clsMonthClass, Me, varAble  ' ** Module Procedure: modTransactionAuditFuncs1.
39550   End Select

EXITP:
39560   Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub
