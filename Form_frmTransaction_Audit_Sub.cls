VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTransaction_Audit_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmTransaction_Audit_Sub"

'VGC 10/05/2017: CHANGES!

' #########################
' ## Use VBA_RenumErrh().  39720
' #########################

' ** Shortcut Alt keys responsive from this form:
' **   AccountNo:        A {cmbAccounts on frmTransaction_Audit_Sub_Criteria}
' **   Acct Number:      B {opgAccountSource_optNumber on frmTransaction_Audit_Sub_Criteria}
' **   Close:            C {cmdClose on frmTransaction_Audit}
' **   Date Posted:      D {TransDateStart on frmTransaction_Audit_Sub_Criteria}
' **   Excel:            E {cmdExcel on frmTransaction_Audit}
' **   Form View:        F {opgView_optForm on frmTransaction_Audit}
' **   Include Archive:  H {chkIncludeArchive on frmTransaction_Audit}
' **   Asset CUSIP:      I {opgAssetSource_optCUSIP on frmTransaction_Audit_Sub_Criteria}
' **   JournalNo:        J {journalno}
' **   Show Criteria:    K {cmdWidenToCriteria on frmTransaction_Audit}
' **   Select All:       L {cmdSelectAll on frmTransaction_Audit}
' **   Acct Name:        M {opgAccountSource_optName on frmTransaction_Audit_Sub_Criteria}
' **   Select None:      N {cmdSelectNone on frmTransaction_Audit}
' **   Word:             O {cmdWord on frmTransaction_Audit}
' **   Print:            P {cmdPrint on frmTransaction_Audit}
' **   Clear All:        R {cmdClearAll on frmTransaction_Audit}
' **   Assets:           S {cmbAssets on frmTransaction_Audit_Sub_Criteria}
' **   Datasheet View:   T {opgView_optDatasheet on frmTransaction_Audit}
' **   View Fields:      V {ckgFlds_cmd on frmTransaction_Audit}
' **   Preview:          W {cmdPreview on frmTransaction_Audit}
' **   Exit:             X {cmdClose on frmTransaction_Audit}
' **   Asset Type:       Y {opgAssetSource_optType on frmTransaction_Audit_Sub_Criteria}
' **   Size/Position:    Z {chkSaveSizePos on frmTransaction_Audit}
' **   JournalType1:     1 {cmbJournalType1 on frmTransaction_Audit_Sub_Criteria}
' **   JournalType2:     2 {cmbJournalType2 on frmTransaction_Audit_Sub_Criteria}
' **   JournalType3:     3 {cmbJournalType3 on frmTransaction_Audit_Sub_Criteria}
'G  Q  U

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker Posting - From:  F5  {cmdCalendar1 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Posting - To:    F6  {cmdCalendar2 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Trade - From:    F7  {cmdCalendar3 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Trade - To:      F8  {cmdCalendar4 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Purchase - From: F9  {cmdCalendar5 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Purchase - To:   F10 {cmdCalendar6 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Posted - From:   F11 {cmdCalendar7 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Posted - To:     F12 {cmdCalendar8 on frmTransaction_Audit_Sub_Criteria}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf on frmTransaction_Audit}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown on frmTransaction_Audit}

' ** Shortcut Alt-Shift keys responsive from this form:
' **   Include Hidden:   C {opgHidden_optInclude on frmTransaction_Audit_Sub_Criteria}
' **   Tax Deduction:    D {chkTaxcodeType_Deduction on frmTransaction_Audit_Sub_Criteria}
' **   Exclude Hidden:   E {opgHidden_optExclude on frmTransaction_Audit_Sub_Criteria}
' **   Tax Income:       I {chkTaxcodeType_Income on frmTransaction_Audit_Sub_Criteria}
' **   Asset Name:       M {opgAssetSource_optName on frmTransaction_Audit_Sub_Criteria}
' **   Narrow Font:      N {chkNarrowFont on frmTransaction_Audit}
' **   Only Hidden:      O {opgHidden_optOnly on frmTransaction_Audit_Sub_Criteria}
' **   Shareface Total:  T {chkIncludeSharefaceTot on frmTransaction_Audit}

' ** Shortcut Ctrl-Alt keys responsive from this form:
' **   Inc/Exp Expense:  E {chkRevcodeType_Expense on frmTransaction_Audit_Sub_Criteria}
' **   Inc/Exp Income:   I {chkRevcodeType_Income on frmTransaction_Audit_Sub_Criteria}

' ** FocusHolders:
' **   FocusHolder : Beginning (Left) of Detail line.
' **   FocusHolder2: End (Right) of Detail line.
' **   FocusHolder3: All the way Left in FormHeader section.
' **   FocusHolder4: At ICash Left in FormHeader section.
' **   FocusHolder5: At AssetDate Left in FormHeader section.
' **   FocusHolder6: At revcode_DESC Left in FormHeader section.

' **      Field                             Twips  Pixels
' **      ================================  =====  ======
' ** 1.   journalno.Width                     780     52
' ** 2.   journaltype.Width                  1020     68
' ** 3.   transdate.Width                     900     60
' ** 4.   accountno.Width                    1200     80
' ** 5.   shortname.Width                    2160    144
' ** 6.   cusip.Width                        1080     72
' ** 7.   asset_description.Width            2100    140
' ** 8.   shareface.Width                    1080     72
' ** 9.   icash.Width                        1440     96
' ** 10.  pcash.Width                        1440     96
' ** 11.  cost.Width                         1440     96
' ** 12.  curr_id.Width                       840     56
' ** 13.  assetdate.Width                    1575    105
' ** 14.  PurchaseDate.Width                 1575    105
' ** 15.  ledger_description.Width           1980    132
' ** 16.  RecurringItem.Width                1980    132
' ** 17.  revcode_DESC.Width                 1920    128
' ** 18.  revcode_TYPE_Description.Width      810     54
' ** 19.  taxcode_description.Width          1920    128
' ** 20.  taxcode_type_description.Width      915     61
' ** 21.  Location_Name.Width                1920    128
' ** 22.  CheckNum.Width                      780     52
' ** 23.  journal_USER.Width                 1020     68
' ** 24.  posted.Width                       1575    105
' ** 25.  ledger_HIDDEN_lbl.Width             720     48

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const ANDF As String = " And "  ' ** Filter 'And'.
Private Const ORF As String = " Or "  ' ** Filter 'Or'.

' ** Array: arr_varFld().
Private lngFlds As Long, arr_varFld As Variant
Private Const FLD_ELEMS As Integer = 13  ' ** Array's first-element UBound().
Private Const F_CNAM     As Integer = 0
Private Const F_FNAM     As Integer = 1
Private Const F_LFT      As Integer = 2
Private Const F_WDT      As Integer = 3
Private Const F_LBL1     As Integer = 4
Private Const F_LBL2     As Integer = 5
Private Const F_LBL_LFT  As Integer = 6
Private Const F_LBL_WDT  As Integer = 7
Private Const F_LIN      As Integer = 8
Private Const F_LIN_LFT  As Integer = 9
Private Const F_LIN_WDT  As Integer = 10
Private Const F_SRT_ADJ  As Integer = 11
Private Const F_VIS      As Integer = 12
Private Const F_CHK_ELEM As Integer = 13

' ** Array: arr_varChk().
Private lngChks As Long, arr_varChk As Variant
'Private Const C_ELEMS As Integer = 17  ' ** Array's first-element UBound().
Private Const C_FNAM  As Integer = 0
'Private Const C_CHKBX As Integer = 1
Private Const C_INCL  As Integer = 2
Private Const C_FOC   As Integer = 3
'Private Const C_MOUS  As Integer = 4
'Private Const C_DIS   As Integer = 5
Private Const C_CMD   As Integer = 6
Private Const C_OFR   As Integer = 7
Private Const C_OFRD  As Integer = 8
Private Const C_OFRF  As Integer = 9
Private Const C_OFRFD As Integer = 10
Private Const C_OFDIS As Integer = 11
Private Const C_ONR   As Integer = 12
Private Const C_ONRD  As Integer = 13
Private Const C_ONRF  As Integer = 14
Private Const C_ONRFD As Integer = 15
Private Const C_ONSD  As Integer = 16
Private Const C_ONDIS As Integer = 17

' ** Array: arr_varFilt().
Private lngFilts As Long, arr_varFilt As Variant
'Private Const F_ELEMS As Integer = 12  ' ** Array's first-element UBound().
'Private Const F_NAM   As Integer = 0
'Private Const F_CONST As Integer = 1
Private Const F_CTL   As Integer = 2
'Private Const F_CLBL  As Integer = 3
Private Const F_FLD   As Integer = 4
Private Const F_FLBL  As Integer = 5
Private Const F_CTL2  As Integer = 6
Private Const F_CLBL2 As Integer = 7
Private Const F_FLD2  As Integer = 8
Private Const F_FLBL2 As Integer = 9
Private Const F_CLBL3 As Integer = 10
Private Const F_FLD3  As Integer = 11
Private Const F_FLBL3 As Integer = 12

Private Const strSortOrig As String = "[journalno]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Top As Long, lngSortLbl_Width As Long
Private lngSortLine_Left As Long, lngSortLine_Top As Long, lngSortLine_Width As Long

Private frmCrit As Access.Form, frmTrans As Access.Form, frmMaint As Access.Form
Private lngFldSep As Long, lngForm_Width As Long
Private lngTotalRecs As Long, lngRecsCur As Long, lngTpp As Long
Private strFilter As String, dblFilterRecs As Double, blnFromTgl As Boolean
Private intPos01 As Integer, intPos02 As Integer
Private strCallingForm As String, strToFoc1 As String, strToFoc2 As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmTransaction_Audit" Then

160         strFilter = vbNullString: dblFilterRecs = 0#
170         strToFoc1 = vbNullString: strToFoc2 = vbNullString
180         blnFromTgl = False

190         Set frmTrans = Forms(strCallingForm)
200         Set frmMaint = Forms("frmMenu_Maintenance")

            ' ****************************************
            ' ** Secondary Progress Bar setup.
            ' ****************************************
210         Set gctlPBar_Box1_2 = .Parent.ProgBar_box
220         Set gctlPBar_Box2_2 = .Parent.ProgBar_box2
230         Set gctlPBar_Lbl_2 = .Parent.ProgBar_lbl
            ' ****************************************
            ' ****************************************
240         gctlPBar_Box1_2.Visible = True
250         gctlPBar_Box2_2.Visible = True
260         ProgBar_Width_Trans frmTrans, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
270         ProgBar_Width_Trans frmTrans, True, 1  ' ** Module Procedure: modTransactionAuditFuncs1.
280         If CurrentUser = "Superuser" Then
290           gctlPBar_Lbl_2.Visible = True
300         Else
310           gctlPBar_Lbl_2.Visible = False
320         End If
330         DoEvents
            ' ****************************************

            ' ****************************************
340         gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
350         frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
360         gctlPBar_Lbl.Caption = "Step 4."
370         DoEvents
380         ProgBar_Width_Trans frmTrans, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
390         gctlPBar_Lbl_2.Caption = "Step 4."
400         DoEvents
            ' ****************************************

410         Set frmCrit = .Parent.frmTransaction_Audit_Sub_Criteria.Form

            ' ****************************************
            ' ** Step 5. Initialize subform
            ' ****************************************
420         gdblPBar_ThisStep = 5#
            ' ****************************************

430         lngSortLbl_Width = .Sort_lbl.Width
440         lngFldSep = (.shortname.Left - (.accountno.Left + .accountno.Width))  ' ** Arbitrary.
450         lngForm_Width = .Width

460         CLR_DISABLED_FG = CLR_DKGRY
470         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
480         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

            ' ** This will then call ChkArray_Set(), below.
            'IT GETS HERE BEFORE TRANS AUDIT LOADS THE FIELD ARRAY!
490         PrintTgls_Load Me ' ** Module Procedure: modTransactionAuditFuncs1.

            ' ****************************************
500         gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
510         frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
520         gctlPBar_Lbl.Caption = "Step 5."
530         DoEvents
540         ProgBar_Width_Trans frmTrans, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
550         gctlPBar_Lbl_2.Caption = "Step 5."
560         DoEvents
            ' ****************************************

570       Else
580         Cancel = -1
590       End If
600     End With

610     If Cancel = -1 Then
620       Beep
630       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
640     End If

EXITP:
650     Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub Form_Load()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim ctl As Access.Control

        ' ****************************************
        ' ** Step 6. Load preferences
        ' ****************************************
710     gdblPBar_ThisStep = 6#
        ' ****************************************
720     Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
730     DoEvents

        ' ****************************************
740     gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
750     frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
760     gctlPBar_Lbl.Caption = "Step 6."
770     DoEvents
780     ProgBar_Width_Trans frmTrans, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
790     gctlPBar_Lbl_2.Caption = "Step 6."
800     DoEvents
        ' ****************************************

        ' ****************************************
        ' ** Step 7. Set print check boxes
        ' ****************************************
810     gdblPBar_ThisStep = 7#
        ' ****************************************

820     With Me
830       For Each ctl In .FormHeader.Controls
840         With ctl
850           If .ControlType = acCheckBox And .Visible = True Then
860             Print_Chk (.Name & "_AfterUpdate")  ' ** Procedure: Below.
870           End If
880         End With
890       Next
900     End With
910     DoEvents

        ' ****************************************
920     gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
930     frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
940     gctlPBar_Lbl.Caption = "Step 7."
950     DoEvents
960     ProgBar_Width_Trans frmTrans, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
970     gctlPBar_Lbl_2.Caption = "Step 7."
980     DoEvents
        ' ****************************************

        ' ****************************************
        ' ** Step 8. Sort subform
        ' ****************************************
990     gdblPBar_ThisStep = 8#
        ' ****************************************

1000    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1010    Set ctl = Nothing
1020    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim ctl As Access.Control
        Dim intRetVal As Integer

1110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1120    intShiftDown = (Shift And acShiftMask) > 0
1130    intAltDown = (Shift And acAltMask) > 0
1140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Date Picker Posting - From:  F5  {cmdCalendar1 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Posting - To:    F6  {cmdCalendar2 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Trade - From:    F7  {cmdCalendar3 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Trade - To:      F8  {cmdCalendar4 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Purchase - From: F9  {cmdCalendar5 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Purchase - To:   F10 {cmdCalendar6 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Posted - From:   F11 {cmdCalendar7 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Posted - To:     F12 {cmdCalendar8 on frmTransaction_Audit_Sub_Criteria}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Date Posted:      D {TransDateStart on frmTransaction_Audit_Sub_Criteria}
        ' **   Include Archive:  H {chkIncludeArchive on frmTransaction_Audit}
        ' **   JournalNo:        J {journalno}
        ' **   Show Criteria:    K {cmdWidenToCriteria on frmTransaction_Audit}
        ' **   Clear All:        R {cmdClearAll on frmTransaction_Audit}
        ' **   Exit:             X {cmdClose on frmTransaction_Audit}

        ' ** Shortcut Alt-Shift keys to other forms and functionality:
        ' **   Include Hidden:   C {opgHidden_optInclude on frmTransaction_Audit_Sub_Criteria}
        ' **   Tax Deduction:    D {chkTaxcodeType_Deduction on frmTransaction_Audit_Sub_Criteria}
        ' **   Exclude Hidden:   E {opgHidden_optExclude on frmTransaction_Audit_Sub_Criteria}
        ' **   Tax Income:       I {chkTaxcodeType_Income on frmTransaction_Audit_Sub_Criteria}
        ' **   Asset Name:       M {opgAssetSource_optName on frmTransaction_Audit_Sub_Criteria}
        ' **   Narrow Font:      N {chkNarrowFont on frmTransaction_Audit}
        ' **   Only Hidden:      O {opgHidden_optOnly on frmTransaction_Audit_Sub_Criteria}
        ' **   Shareface Total:  T {chkIncludeSharefaceTot on frmTransaction_Audit}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf on frmTransaction_Audit}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown on frmTransaction_Audit}

        ' ** Shortcut Ctrl-Alt keys to other forms and functionality:
        ' **   Inc/Exp Expense:  E {chkRevcodeType_Expense on frmTransaction_Audit_Sub_Criteria}
        ' **   Inc/Exp Income:   I {chkRevcodeType_Income on frmTransaction_Audit_Sub_Criteria}

        ' ** Plain keys.
1150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyF5
            ' **   Date Picker Posting - From:  F5  {cmdCalendar1}
1170        With Me
1180          intRetVal = 0
1190          If frmCrit.cmdCalendar1.Enabled = True Then
1200            DoCmd.SelectObject acForm, .Parent.Name, False
1210            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1220            frmCrit.cmdCalendar1.SetFocus
1230            frmCrit.cmdCalendar1_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1240          End If
1250        End With
1260      Case vbKeyF6
            ' **   Date Picker Posting - To:    F6  {cmdCalendar2}
1270        With Me
1280          intRetVal = 0
1290          If frmCrit.cmdCalendar2.Enabled = True Then
1300            DoCmd.SelectObject acForm, .Parent.Name, False
1310            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1320            frmCrit.cmdCalendar2.SetFocus
1330            frmCrit.cmdCalendar2_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1340          End If
1350        End With
1360      Case vbKeyF7
            ' **   Date Picker Trade - From:    F7  {cmdCalendar3}
1370        With Me
1380          intRetVal = 0
1390          If frmCrit.cmdCalendar3.Enabled = True Then
1400            DoCmd.SelectObject acForm, .Parent.Name, False
1410            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1420            frmCrit.cmdCalendar3.SetFocus
1430            frmCrit.cmdCalendar3_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1440          End If
1450        End With
1460      Case vbKeyF8
            ' **   Date Picker Trade - To:      F8  {cmdCalendar4}
1470        With Me
1480          intRetVal = 0
1490          If frmCrit.cmdCalendar4.Enabled = True Then
1500            DoCmd.SelectObject acForm, .Parent.Name, False
1510            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1520            frmCrit.cmdCalendar4.SetFocus
1530            frmCrit.cmdCalendar4_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1540          End If
1550        End With
1560      Case vbKeyF9
            ' **   Date Picker Purchase - From: F9  {cmdCalendar5}
1570        With Me
1580          intRetVal = 0
1590          If frmCrit.cmdCalendar5.Enabled = True Then
1600            DoCmd.SelectObject acForm, .Parent.Name, False
1610            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1620            frmCrit.cmdCalendar5.SetFocus
1630            frmCrit.cmdCalendar5_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1640          End If
1650        End With
1660      Case vbKeyF10
            ' **   Date Picker Purchase - To:   F10 {cmdCalendar6}
1670        With Me
1680          intRetVal = 0
1690          If frmCrit.cmdCalendar6.Enabled = True Then
1700            DoCmd.SelectObject acForm, .Parent.Name, False
1710            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1720            frmCrit.cmdCalendar6.SetFocus
1730            frmCrit.cmdCalendar6_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1740          End If
1750        End With
1760      Case vbKeyF11
            ' **   Date Picker Posted - From:   F11 {cmdCalendar7}
1770        With Me
1780          intRetVal = 0
1790          If frmCrit.cmdCalendar7.Enabled = True Then
1800            DoCmd.SelectObject acForm, .Parent.Name, False
1810            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1820            frmCrit.cmdCalendar7.SetFocus
1830            frmCrit.cmdCalendar7_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1840          End If
1850        End With
1860      Case vbKeyF12
            ' **   Date Picker Posted - To:     F12 {cmdCalendar8}
1870        With Me
1880          intRetVal = 0
1890          If frmCrit.cmdCalendar8.Enabled = True Then
1900            DoCmd.SelectObject acForm, .Parent.Name, False
1910            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1920            frmCrit.cmdCalendar8.SetFocus
1930            frmCrit.cmdCalendar8_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1940          End If
1950        End With
1960      Case vbKeyUp
1970        intRetVal = 0
1980        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1990      Case vbKeyDown
2000        intRetVal = 0
2010        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2020      End Select
2030    End If

        ' ** Alt keys.
2040    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2050      Select Case intRetVal
          Case vbKeyD
2060        With Me
2070          intRetVal = 0
2080          DoCmd.SelectObject acForm, .Parent.Name, False
2090          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
2100          frmCrit.TransDateStart.SetFocus
2110        End With
2120      Case vbKeyH
2130        With Me
2140          intRetVal = 0
2150          If .Parent.chkIncludeArchive.Enabled = True Then
2160            DoCmd.SelectObject acForm, .Parent.Name, False
2170            .Parent.chkIncludeArchive.SetFocus
2180            .Parent.chkIncludeArchive = (Not .Parent.chkIncludeArchive)
2190            .Parent.chkIncludeArchive_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit.
2200          Else
2210            Beep
2220          End If
2230        End With
2240      Case vbKeyJ
2250        With Me
2260          intRetVal = 0
2270          If .journalno.Visible = True And .journalno.Enabled = True Then
2280            .journalno.SetFocus
2290          Else
2300            Beep
2310          End If
2320        End With
2330      Case vbKeyK
2340        With Me
2350          intRetVal = 0
2360          If .Parent.cmdWidenToCriteria.Enabled = True Then
2370            DoCmd.SelectObject acForm, .Parent.Name, False
2380            .Parent.cmdWidenToCriteria.SetFocus
2390            .Parent.cmdWidenToCriteria_Click  ' ** Form Procedure: frmTransaction_Audit.
2400          Else
2410            Beep
2420          End If
2430        End With
2440      Case vbKeyR
2450        With Me
2460          If dblFilterRecs = 0# Then
2470            DoCmd.SelectObject acForm, .Parent.Name, False
2480            .Parent.cmdClearAll.SetFocus
2490            .Parent.cmdClearAll_Click  ' ** Form Procedure: frmTransaction_Audit.
2500          End If
2510        End With
2520      Case vbKeyX
2530        With Me
2540          intRetVal = 0
2550          .Parent.cmdClose_Click  ' ** Form Procedure: frmTransaction_Audit.
2560        End With
2570      End Select
2580    End If

        ' ** Alt-Shift keys.
2590    If (Not intCtrlDown) And intAltDown And intShiftDown Then
2600      Select Case intRetVal
          Case vbKeyC
2610        With Me
2620          intRetVal = 0
2630          DoCmd.SelectObject acForm, .Parent.Name, False
2640          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
2650          frmCrit.opgHidden.SetFocus
2660          frmCrit.opgHidden = frmCrit.opgHidden_optInclude.OptionValue
2670          frmCrit.opgHidden_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
2680        End With
2690      Case vbKeyD
2700        With Me
2710          intRetVal = 0
2720          DoCmd.SelectObject acForm, .Parent.Name, False
2730          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
2740          frmCrit.chkTaxcodeType_Deduction.SetFocus
2750          frmCrit.chkTaxcodeType_Deduction = (Not frmCrit.chkTaxcodeType_Deduction)
2760          frmCrit.chkTaxcodeType_Deduction_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
2770        End With
2780      Case vbKeyE
2790        With Me
2800          intRetVal = 0
2810          DoCmd.SelectObject acForm, .Parent.Name, False
2820          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
2830          frmCrit.opgHidden.SetFocus
2840          frmCrit.opgHidden = frmCrit.opgHidden_optExclude.OptionValue
2850          frmCrit.opgHidden_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
2860        End With
2870      Case vbKeyI
2880        With Me
2890          intRetVal = 0
2900          DoCmd.SelectObject acForm, .Parent.Name, False
2910          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
2920          frmCrit.chkTaxcodeType_Income.SetFocus
2930          frmCrit.chkTaxcodeType_Income = (Not frmCrit.chkTaxcodeType_Income)
2940          frmCrit.chkTaxcodeType_Income_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria
2950        End With
2960      Case vbKeyM
2970        With Me
2980          intRetVal = 0
2990          DoCmd.SelectObject acForm, .Parent.Name, False
3000          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
3010          frmCrit.opgAssetSource = frmCrit.opgAssetSource_optName.OptionValue
3020          frmCrit.opgAssetSource.SetFocus
3030          frmCrit.opgAssetSource_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
3040        End With
3050      Case vbKeyN
3060        With Me
3070          intRetVal = 0
3080          DoCmd.SelectObject acForm, .Parent.Name, False
3090          .Parent.chkNarrowFont.SetFocus
3100          .Parent.chkNarrowFont = (Not .Parent.chkNarrowFont)
3110          .Parent.chkNarrowFont_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit.
3120        End With
3130      Case vbKeyO
3140        With Me
3150          intRetVal = 0
3160          DoCmd.SelectObject acForm, .Parent.Name, False
3170          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
3180          frmCrit.opgHidden.SetFocus
3190          frmCrit.opgHidden = frmCrit.opgHidden_optOnly.OptionValue
3200          frmCrit.opgHidden_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
3210        End With
3220      Case vbKeyT
3230        With Me
3240          intRetVal = 0
3250          DoCmd.SelectObject acForm, .Parent.Name, False
3260          .Parent.chkIncludeSharefaceTot.SetFocus
3270          .Parent.chkIncludeSharefaceTot = (Not .Parent.chkIncludeSharefaceTot)
3280          .Parent.chkIncludeSharefaceTot_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit.
3290        End With
3300      End Select
3310    End If

        ' ** Ctrl keys.
3320    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3330      Select Case intRetVal
          Case vbKeyG
3340        With Me
3350          intRetVal = 0
3360          DoCmd.SelectObject acForm, .Parent.Name, False
3370          .Parent.chkPageOf.SetFocus
3380          .Parent.chkPageOf = (Not .Parent.chkPageOf)
3390          .Parent.chkPageOf_AfterUpdate  ' ** Procedure: frmTransaction_Audit.
3400        End With
3410      Case vbKeyTab, vbKeyReturn
3420        With Me
3430          intRetVal = 0
3440          DoCmd.SelectObject acForm, .Parent.Name, False
3450          .Parent.cmdSelectAll.SetFocus
3460        End With
3470      Case vbKeyPageUp, vbKeyHome
3480        intRetVal = 0
3490        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
3500      Case vbKeyPageDown, vbKeyEnd
3510        intRetVal = 0
3520        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
3530      End Select
3540    End If

        ' ** Ctrl-Shift keys.
3550    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3560      Select Case intRetVal
          Case vbKeyF
3570        With Me
3580          intRetVal = 0
3590          DoCmd.SelectObject acForm, .Parent.Name, False
3600          .Parent.FocusHolder.SetFocus
3610        End With
3620      Case vbKeyT
3630        intRetVal = 0
3640        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
3650      Case vbKeyTab, vbKeyReturn
3660        With Me
3670          intRetVal = 0
3680          DoCmd.SelectObject acForm, .Parent.Name, False
3690          .Parent.opgView.SetFocus
3700        End With
3710      End Select
3720    End If

        ' ** Ctrl-Alt keys.
3730    If intCtrlDown And intAltDown And (Not intShiftDown) Then
3740      Select Case intRetVal
          Case vbKeyE
3750        With Me
3760          intRetVal = 0
3770          DoCmd.SelectObject acForm, .Parent.Name, False
3780          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
3790          frmCrit.chkRevcodeType_Expense.SetFocus
3800          frmCrit.chkRevcodeType_Expense = (Not frmCrit.chkRevcodeType_Expense)
3810          frmCrit.chkRevcodeType_Expense_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
3820        End With
3830      Case vbKeyI
3840        With Me
3850          intRetVal = 0
3860          DoCmd.SelectObject acForm, .Parent.Name, False
3870          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
3880          frmCrit.chkRevcodeType_Income.SetFocus
3890          frmCrit.chkRevcodeType_Income = (Not frmCrit.chkRevcodeType_Income)
3900          frmCrit.chkRevcodeType_Income_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
3910        End With
3920      End Select
3930    End If

EXITP:
3940    Set ctl = Nothing
3950    KeyCode = intRetVal
3960    Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub Form_Timer()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strControl As String
        Dim lngLoopCnt As Long
        Dim strTmp01 As String

4010    With Me
          ' ** Kluge!
4020      .TimerInterval = 0
4030      If strToFoc2 <> vbNullString Then
            ' ** journal_USER_GotFocus, posted_GotFocus, ledger_HIDDEN_GotFocus.
4040        strControl = Left(strToFoc2, (Len(strToFoc2) - Len("_GotFocus")))
4050        .Controls(strControl).SetFocus
4060        If strToFoc1 <> vbNullString Then
4070          .TimerInterval = 1000&
4080        End If
4090      End If
4100      If strToFoc1 <> vbNullString Then
4110        strControl = Left(strToFoc1, (Len(strToFoc1) - Len("_GotFocus")))
4120        lngLoopCnt = 0&
4130        strTmp01 = "FocusHolder"
4140      End If
4150    End With

EXITP:
4160    Exit Sub

ERRH:
660     strToFoc2 = vbNullString
670     Me.ledger_HIDDEN.Enabled = True
680     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
690     Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        ' ** Pref_Save() in frmTransaction_Audit.

EXITP:
4210    Set frmCrit = Nothing
4220    Set frmTrans = Nothing
4230    Set frmMaint = Nothing
4240    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

4310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
4320    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Public Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

        Dim lngX As Long

4410    With Me

4420      ChkArray_Pop  ' ** Procedure: Below.

4430      For lngX = 0& To (lngChks - 1&)
4440        If .Controls(arr_varChk(C_OFRFD, lngX)).Visible = True Or .Controls(arr_varChk(C_OFRF, lngX)).Visible = True Or _
                .Controls(arr_varChk(C_ONRFD, lngX)).Visible = True Or .Controls(arr_varChk(C_ONRF, lngX)).Visible = True Then
4450          Select Case arr_varChk(C_INCL, lngX)
              Case True
4460            Select Case arr_varChk(C_FOC, lngX)
                Case True
4470              .Controls(arr_varChk(C_ONRD, lngX)).Visible = True
4480              .Controls(arr_varChk(C_ONR, lngX)).Visible = False
4490            Case False
4500              .Controls(arr_varChk(C_ONR, lngX)).Visible = True
4510              .Controls(arr_varChk(C_ONRD, lngX)).Visible = False
4520            End Select
4530            .Controls(arr_varChk(C_OFR, lngX)).Visible = False
4540            .Controls(arr_varChk(C_OFRD, lngX)).Visible = False
4550          Case False
4560            Select Case arr_varChk(C_FOC, lngX)
                Case True
4570              .Controls(arr_varChk(C_OFRD, lngX)).Visible = True
4580              .Controls(arr_varChk(C_OFR, lngX)).Visible = False
4590            Case False
4600              .Controls(arr_varChk(C_OFR, lngX)).Visible = True
4610              .Controls(arr_varChk(C_OFRD, lngX)).Visible = False
4620            End Select
4630            .Controls(arr_varChk(C_ONR, lngX)).Visible = False
4640            .Controls(arr_varChk(C_ONRD, lngX)).Visible = False
4650          End Select
4660          .Controls(arr_varChk(C_OFRF, lngX)).Visible = False
4670          .Controls(arr_varChk(C_OFRFD, lngX)).Visible = False
4680          .Controls(arr_varChk(C_OFDIS, lngX)).Visible = False
4690          .Controls(arr_varChk(C_ONRF, lngX)).Visible = False
4700          .Controls(arr_varChk(C_ONRFD, lngX)).Visible = False
4710          .Controls(arr_varChk(C_ONSD, lngX)).Visible = False
4720          .Controls(arr_varChk(C_ONDIS, lngX)).Visible = False
4730        End If
4740      Next
4750    End With

EXITP:
4760    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub FocusHolder_GotFocus()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_GotFocus"

4810    With Me
          ' ** Kluge!
4820      strToFoc1 = "journalno_GotFocus"
4830      strToFoc2 = vbNullString
4840      .TimerInterval = 500&
4850    End With

EXITP:
4860    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journalno_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_lbl_DblClick"

4910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4920    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journalno_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_lbl_MouseMove"

5010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5020    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub FocusHolder3_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder3_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5120    intShiftDown = (Shift And acShiftMask) > 0
5130    intAltDown = (Shift And acAltMask) > 0
5140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5170        With Me
5180          intRetVal = 0
5190          .FocusHolder.SetFocus
5200        End With
5210      End Select
5220    End If

        ' ** Shift keys.
5230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5250        With Me
5260          intRetVal = 0
5270          .FocusHolder2.SetFocus
5280        End With
5290      End Select
5300    End If

EXITP:
5310    KeyCode = intRetVal
5320    Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Sub

Private Sub journalno_tgl_GotFocus()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_tgl_GotFocus"

5410    ChkArray_Pop  ' ** Procedure: Below.

5420    PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
5430    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journalno_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_tgl_MouseDown"

5510    ChkArray_Pop  ' ** Procedure: Below.

5520    PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
5530    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journalno_tgl_Click()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_tgl_Click"

5610    ChkArray_Pop  ' ** Procedure: Below.

5620    PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
5630    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journalno_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5720    intShiftDown = (Shift And acShiftMask) > 0
5730    intAltDown = (Shift And acAltMask) > 0
5740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5770        With Me
5780          If .journaltype_tgl.Visible = True And .journaltype_tgl.Enabled = True Then
5790            intRetVal = 0
5800            .journaltype_tgl.SetFocus
5810          End If
5820        End With
5830      End Select
5840    End If

        ' ** Shift keys.
5850    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5870        With Me
5880          intRetVal = 0
5890          .FocusHolder.SetFocus
5900        End With
5910      End Select
5920    End If

EXITP:
5930    KeyCode = intRetVal
5940    Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub journalno_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_tgl_MouseMove"

6010    ChkArray_Pop  ' ** Procedure: Below.

6020    PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
6030    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journalno_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_tgl_MouseUp"

6110    ChkArray_Pop  ' ** Procedure: Below.

6120    PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
6130    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journalno_tgl_LostFocus()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_tgl_LostFocus"

6210    ChkArray_Pop  ' ** Procedure: Below.

6220    PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
6230    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Public Sub journalno_GotFocus()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_GotFocus"

6310    With Me
          ' ** Kluge!
6320      strToFoc1 = vbNullString: strToFoc2 = vbNullString
6330      .journalno.SelLength = 0
6340      .journalno.SelStart = 9
6350      blnFromTgl = False
6360    End With

EXITP:
6370    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journalno_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6420    intShiftDown = (Shift And acShiftMask) > 0
6430    intAltDown = (Shift And acAltMask) > 0
6440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6470        With Me
6480          intRetVal = 0
6490          NextFocus THIS_PROC, True, lngFlds, arr_varFld, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
6500        End With
6510      End Select
6520    End If

        ' ** Shift keys.
6530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6550        With Me
6560          intRetVal = 0
6570          NextFocus THIS_PROC, False, lngFlds, arr_varFld, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
6580        End With
6590      End Select
6600    End If

EXITP:
6610    KeyCode = intRetVal
6620    Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub journaltype_lbl_DblClick(Cancel As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_lbl_DblClick"

6710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6720    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journaltype_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_lbl_MouseMove"

6810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6820    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journaltype_tgl_GotFocus()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_tgl_GotFocus"

6910    ChkArray_Pop  ' ** Procedure: Below.

6920    PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
6930    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journaltype_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_tgl_MouseDown"

7010    ChkArray_Pop  ' ** Procedure: Below.

7020    PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
7030    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journaltype_tgl_Click()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_tgl_Click"

7110    ChkArray_Pop  ' ** Procedure: Below.

7120    PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
7130    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journaltype_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7220    intShiftDown = (Shift And acShiftMask) > 0
7230    intAltDown = (Shift And acAltMask) > 0
7240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7270        With Me
7280          If .transdate_tgl.Visible = True And .transdate_tgl.Enabled = True Then
7290            intRetVal = 0
7300            .transdate_tgl.SetFocus
7310          End If
7320        End With
7330      End Select
7340    End If

        ' ** Shift keys.
7350    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7370        With Me
7380          If .journalno_tgl.Visible = True And .journalno_tgl.Enabled = True Then
7390            intRetVal = 0
7400            .journalno_tgl.SetFocus
7410          End If
7420        End With
7430      End Select
7440    End If

EXITP:
7450    KeyCode = intRetVal
7460    Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub journaltype_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_tgl_MouseMove"

7510    ChkArray_Pop  ' ** Procedure: Below.

7520    PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
7530    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journaltype_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_tgl_MouseUp"

7610    ChkArray_Pop  ' ** Procedure: Below.

7620    PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
7630    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journaltype_tgl_LostFocus()

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_tgl_LostFocus"

7710    ChkArray_Pop  ' ** Procedure: Below.

7720    PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
7730    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journaltype_GotFocus()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_GotFocus"

7810    With Me
7820      If blnFromTgl = True Then
7830        blnFromTgl = False
7840        .journaltype.SelLength = 0
7850        .journaltype.SelStart = 19
7860      End If
7870    End With

EXITP:
7880    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

7910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7920    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub transdate_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_MouseMove"

8010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8020    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub transdate_tgl_GotFocus()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_tgl_GotFocus"

8110    ChkArray_Pop  ' ** Procedure: Below.

8120    PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
8130    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub transdate_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_tgl_MouseDown"

8210    ChkArray_Pop  ' ** Procedure: Below.

8220    PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
8230    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub transdate_tgl_Click()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_tgl_Click"

8310    ChkArray_Pop  ' ** Procedure: Below.

8320    PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
8330    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub transdate_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_tgl_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
8470        With Me
8480          If .accountno_tgl.Visible = True And .accountno_tgl.Enabled = True Then
8490            intRetVal = 0
8500            .accountno_tgl.SetFocus
8510          End If
8520        End With
8530      End Select
8540    End If

        ' ** Shift keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8570        With Me
8580          If .journaltype_tgl.Visible = True And .journaltype_tgl.Enabled = True Then
8590            intRetVal = 0
8600            .journaltype_tgl.SetFocus
8610          End If
8620        End With
8630      End Select
8640    End If

EXITP:
8650    KeyCode = intRetVal
8660    Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub transdate_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_tgl_MouseMove"

8710    ChkArray_Pop  ' ** Procedure: Below.

8720    PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
8730    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub transdate_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_tgl_MouseUp"

8810    ChkArray_Pop  ' ** Procedure: Below.

8820    PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
8830    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub transdate_tgl_LostFocus()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_tgl_LostFocus"

8910    ChkArray_Pop  ' ** Procedure: Below.

8920    PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
8930    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub transdate_GotFocus()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_GotFocus"

9010    With Me
9020      If blnFromTgl = True Then
9030        blnFromTgl = False
9040        .transdate.SelLength = 0
9050        .transdate.SelStart = 19
9060      End If
9070    End With

EXITP:
9080    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

9110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
9120    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub accountno_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_MouseMove"

9210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9220    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub accountno_tgl_GotFocus()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_tgl_GotFocus"

9310    ChkArray_Pop  ' ** Procedure: Below.

9320    PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
9330    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub accountno_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_tgl_MouseDown"

9410    ChkArray_Pop  ' ** Procedure: Below.

9420    PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
9430    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub accountno_tgl_Click()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_tgl_Click"

9510    ChkArray_Pop  ' ** Procedure: Below.

9520    PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
9530    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub accountno_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_tgl_KeyDown"

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
9680          If .shortname_tgl.Visible = True And .shortname_tgl.Enabled = True Then
9690            intRetVal = 0
9700            .shortname_tgl.SetFocus
9710          End If
9720        End With
9730      End Select
9740    End If

        ' ** Shift keys.
9750    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9770        With Me
9780          If .transdate_tgl.Visible = True And .transdate_tgl.Enabled = True Then
9790            intRetVal = 0
9800            .transdate_tgl.SetFocus
9810          End If
9820        End With
9830      End Select
9840    End If

EXITP:
9850    KeyCode = intRetVal
9860    Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub accountno_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_tgl_MouseMove"

9910    ChkArray_Pop  ' ** Procedure: Below.

9920    PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
9930    Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub accountno_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_tgl_MouseUp"

10010   ChkArray_Pop  ' ** Procedure: Below.

10020   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
10030   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub accountno_tgl_LostFocus()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_tgl_LostFocus"

10110   ChkArray_Pop  ' ** Procedure: Below.

10120   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
10130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub accountno_GotFocus()

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_GotFocus"

10210   With Me
10220     If blnFromTgl = True Then
10230       blnFromTgl = False
10240       .accountno.SelLength = 0
10250       .accountno.SelStart = 19
10260     End If
10270   End With

EXITP:
10280   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

10310   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
10320   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shortname_tgl_GotFocus()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_tgl_GotFocus"

10410   ChkArray_Pop  ' ** Procedure: Below.

10420   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
10430   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shortname_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_tgl_MouseDown"

10510   ChkArray_Pop  ' ** Procedure: Below.

10520   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
10530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shortname_tgl_Click()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_tgl_Click"

10610   ChkArray_Pop  ' ** Procedure: Below.

10620   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
10630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shortname_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10720   intShiftDown = (Shift And acShiftMask) > 0
10730   intAltDown = (Shift And acAltMask) > 0
10740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10770       With Me
10780         If .cusip_tgl.Visible = True And .cusip_tgl.Enabled = True Then
10790           intRetVal = 0
10800           .cusip_tgl.SetFocus
10810         End If
10820       End With
10830     End Select
10840   End If

        ' ** Shift keys.
10850   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10870       With Me
10880         If .accountno_tgl.Visible = True And .accountno_tgl.Enabled = True Then
10890           intRetVal = 0
10900           .accountno_tgl.SetFocus
10910         End If
10920       End With
10930     End Select
10940   End If

EXITP:
10950   KeyCode = intRetVal
10960   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub shortname_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_tgl_MouseMove"

11010   ChkArray_Pop  ' ** Procedure: Below.

11020   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
11030   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shortname_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_tgl_MouseUp"

11110   ChkArray_Pop  ' ** Procedure: Below.

11120   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
11130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shortname_tgl_LostFocus()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_tgl_LostFocus"

11210   ChkArray_Pop  ' ** Procedure: Below.

11220   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
11230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shortname_GotFocus()

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_GotFocus"

11310   With Me
11320     If blnFromTgl = True Then
11330       blnFromTgl = False
11340       .shortname.SelLength = 0
11350       .shortname.SelStart = 0
11360     End If
11370   End With

EXITP:
11380   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cusip_lbl_DblClick(Cancel As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_lbl_DblClick"

11410   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
11420   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cusip_tgl_GotFocus()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_tgl_GotFocus"

11510   ChkArray_Pop  ' ** Procedure: Below.

11520   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
11530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cusip_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_tgl_MouseDown"

11610   ChkArray_Pop  ' ** Procedure: Below.

11620   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
11630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cusip_tgl_Click()

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_tgl_Click"

11710   ChkArray_Pop  ' ** Procedure: Below.

11720   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
11730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cusip_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11820   intShiftDown = (Shift And acShiftMask) > 0
11830   intAltDown = (Shift And acAltMask) > 0
11840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11870       With Me
11880         If .asset_description_tgl.Visible = True And .asset_description_tgl.Enabled = True Then
11890           intRetVal = 0
11900           .asset_description_tgl.SetFocus
11910         End If
11920       End With
11930     End Select
11940   End If

        ' ** Shift keys.
11950   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11970       With Me
11980         If .shortname_tgl.Visible = True And .shortname_tgl.Enabled = True Then
11990           intRetVal = 0
12000           .shortname_tgl.SetFocus
12010         End If
12020       End With
12030     End Select
12040   End If

EXITP:
12050   KeyCode = intRetVal
12060   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub cusip_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_tgl_MouseMove"

12110   ChkArray_Pop  ' ** Procedure: Below.

12120   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
12130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cusip_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_tgl_MouseUp"

12210   ChkArray_Pop  ' ** Procedure: Below.

12220   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
12230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cusip_tgl_LostFocus()

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_tgl_LostFocus"

12310   ChkArray_Pop  ' ** Procedure: Below.

12320   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
12330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cusip_GotFocus()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_GotFocus"

12410   With Me
12420     If blnFromTgl = True Then
12430       blnFromTgl = False
12440       .cusip.SelLength = 0
12450       .cusip.SelStart = 19
12460     End If
12470   End With

EXITP:
12480   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub asset_description_lbl_DblClick(Cancel As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_lbl_DblClick"

12510   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
12520   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub asset_description_tgl_GotFocus()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_tgl_GotFocus"

12610   ChkArray_Pop  ' ** Procedure: Below.

12620   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
12630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub asset_description_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_tgl_MouseDown"

12710   ChkArray_Pop  ' ** Procedure: Below.

12720   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
12730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub asset_description_tgl_Click()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_tgl_Click"

12810   ChkArray_Pop  ' ** Procedure: Below.

12820   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
12830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub asset_description_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_tgl_KeyDown"

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
12980         If .shareface_tgl.Visible = True And .shareface_tgl.Enabled = True Then
12990           intRetVal = 0
13000           .shareface_tgl.SetFocus
13010         End If
13020       End With
13030     End Select
13040   End If

        ' ** Shift keys.
13050   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13070       With Me
13080         If .cusip_tgl.Visible = True And .cusip_tgl.Enabled = True Then
13090           intRetVal = 0
13100           .cusip_tgl.SetFocus
13110         End If
13120       End With
13130     End Select
13140   End If

EXITP:
13150   KeyCode = intRetVal
13160   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub asset_description_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_tgl_MouseMove"

13210   ChkArray_Pop  ' ** Procedure: Below.

13220   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
13230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub asset_description_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_tgl_MouseUp"

13310   ChkArray_Pop  ' ** Procedure: Below.

13320   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
13330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub asset_description_tgl_LostFocus()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_tgl_LostFocus"

13410   ChkArray_Pop  ' ** Procedure: Below.

13420   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
13430   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub asset_description_GotFocus()

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_GotFocus"

13510   With Me
13520     If blnFromTgl = True Then
13530       blnFromTgl = False
13540       .asset_description.SelLength = 0
13550       .asset_description.SelStart = 0
13560     End If
13570   End With

EXITP:
13580   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shareface_lbl_DblClick(Cancel As Integer)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_lbl_DblClick"

13610   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
13620   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shareface_tgl_GotFocus()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_tgl_GotFocus"

13710   ChkArray_Pop  ' ** Procedure: Below.

13720   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
13730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shareface_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_tgl_MouseDown"

13810   ChkArray_Pop  ' ** Procedure: Below.

13820   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
13830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shareface_tgl_Click()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_tgl_Click"

13910   ChkArray_Pop  ' ** Procedure: Below.

13920   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
13930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shareface_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14020   intShiftDown = (Shift And acShiftMask) > 0
14030   intAltDown = (Shift And acAltMask) > 0
14040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14070       With Me
14080         If .icash_tgl.Visible = True And .icash_tgl.Enabled = True Then
14090           intRetVal = 0
14100           .icash_tgl.SetFocus
14110         End If
14120       End With
14130     End Select
14140   End If

        ' ** Shift keys.
14150   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14170       With Me
14180         If .asset_description_tgl.Visible = True And .asset_description_tgl.Enabled = True Then
14190           intRetVal = 0
14200           .asset_description_tgl.SetFocus
14210         End If
14220       End With
14230     End Select
14240   End If

EXITP:
14250   KeyCode = intRetVal
14260   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub shareface_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_tgl_MouseMove"

14310   ChkArray_Pop  ' ** Procedure: Below.

14320   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
14330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shareface_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_tgl_MouseUp"

14410   ChkArray_Pop  ' ** Procedure: Below.

14420   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
14430   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shareface_tgl_LostFocus()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_tgl_LostFocus"

14510   ChkArray_Pop  ' ** Procedure: Below.

14520   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
14530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub shareface_GotFocus()

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_GotFocus"

14610   With Me
14620     If blnFromTgl = True Then
14630       blnFromTgl = False
14640       .shareface.SelLength = 0
14650       .shareface.SelStart = 0
14660     End If
14670   End With

EXITP:
14680   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub icash_lbl_DblClick(Cancel As Integer)

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_lbl_DblClick"

14710   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
14720   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub FocusHolder4_KeyDown(KeyCode As Integer, Shift As Integer)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder4_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
14870       With Me
14880         intRetVal = 0
14890         If .ICash.Visible = True And .ICash.Enabled = True Then
14900           .ICash.SetFocus
14910         Else
14920           .FocusHolder.SetFocus
14930         End If
14940       End With
14950     End Select
14960   End If

        ' ** Shift keys.
14970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14980     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14990       With Me
15000         intRetVal = 0
15010         .FocusHolder3.SetFocus
15020       End With
15030     End Select
15040   End If

EXITP:
15050   KeyCode = intRetVal
15060   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Sub

Private Sub icash_tgl_GotFocus()

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_tgl_GotFocus"

15110   ChkArray_Pop  ' ** Procedure: Below.

15120   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
15130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub icash_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_tgl_MouseDown"

15210   ChkArray_Pop  ' ** Procedure: Below.

15220   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
15230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub icash_tgl_Click()

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_tgl_Click"

15310   ChkArray_Pop  ' ** Procedure: Below.

15320   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
15330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub icash_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15420   intShiftDown = (Shift And acShiftMask) > 0
15430   intAltDown = (Shift And acAltMask) > 0
15440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15470       With Me
15480         If .pcash_tgl.Visible = True And .pcash_tgl.Enabled = True Then
15490           intRetVal = 0
15500           .pcash_tgl.SetFocus
15510         End If
15520       End With
15530     End Select
15540   End If

        ' ** Shift keys.
15550   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15570       With Me
15580         If .shareface_tgl.Visible = True And .shareface_tgl.Enabled = True Then
15590           intRetVal = 0
15600           .shareface_tgl.SetFocus
15610         End If
15620       End With
15630     End Select
15640   End If

EXITP:
15650   KeyCode = intRetVal
15660   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub icash_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_tgl_MouseMove"

15710   ChkArray_Pop  ' ** Procedure: Below.

15720   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
15730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub icash_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_tgl_MouseUp"

15810   ChkArray_Pop  ' ** Procedure: Below.

15820   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
15830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub icash_tgl_LostFocus()

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_tgl_LostFocus"

15910   ChkArray_Pop  ' ** Procedure: Below.

15920   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
15930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub icash_GotFocus()

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_GotFocus"

16010   With Me
16020     If blnFromTgl = True Then
16030       blnFromTgl = False
16040       .ICash.SelLength = 0
16050       .ICash.SelStart = 0
16060     End If
16070   End With

EXITP:
16080   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub pcash_lbl_DblClick(Cancel As Integer)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_lbl_DblClick"

16110   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
16120   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub pcash_tgl_GotFocus()

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_tgl_GotFocus"

16210   ChkArray_Pop  ' ** Procedure: Below.

16220   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
16230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub pcash_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_tgl_MouseDown"

16310   ChkArray_Pop  ' ** Procedure: Below.

16320   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
16330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub pcash_tgl_Click()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_tgl_Click"

16410   ChkArray_Pop  ' ** Procedure: Below.

16420   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
16430   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub pcash_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16520   intShiftDown = (Shift And acShiftMask) > 0
16530   intAltDown = (Shift And acAltMask) > 0
16540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16570       With Me
16580         If .cost_tgl.Visible = True And .cost_tgl.Enabled = True Then
16590           intRetVal = 0
16600           .cost_tgl.SetFocus
16610         End If
16620       End With
16630     End Select
16640   End If

        ' ** Shift keys.
16650   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16670       With Me
16680         If .icash_tgl.Visible = True And .icash_tgl.Enabled = True Then
16690           intRetVal = 0
16700           .icash_tgl.SetFocus
16710         End If
16720       End With
16730     End Select
16740   End If

EXITP:
16750   KeyCode = intRetVal
16760   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub pcash_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_tgl_MouseMove"

16810   ChkArray_Pop  ' ** Procedure: Below.

16820   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
16830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub pcash_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_tgl_MouseUp"

16910   ChkArray_Pop  ' ** Procedure: Below.

16920   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
16930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub pcash_tgl_LostFocus()

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_tgl_LostFocus"

17010   ChkArray_Pop  ' ** Procedure: Below.

17020   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
17030   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub pcash_GotFocus()

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_GotFocus"

17110   With Me
17120     If blnFromTgl = True Then
17130       blnFromTgl = False
17140       .PCash.SelLength = 0
17150       .PCash.SelStart = 0
17160     End If
17170   End With

EXITP:
17180   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cost_lbl_DblClick(Cancel As Integer)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_lbl_DblClick"

17210   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
17220   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cost_tgl_GotFocus()

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_tgl_GotFocus"

17310   ChkArray_Pop  ' ** Procedure: Below.

17320   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
17330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cost_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_tgl_MouseDown"

17410   ChkArray_Pop  ' ** Procedure: Below.

17420   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
17430   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cost_tgl_Click()

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_tgl_Click"

17510   ChkArray_Pop  ' ** Procedure: Below.

17520   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
17530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cost_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_tgl_KeyDown"

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
17680         If .curr_id_tgl.Visible = True And .curr_id_tgl.Enabled = True Then
17690           .curr_id_tgl.SetFocus
17700         ElseIf .assetdate_tgl.Visible = True And .assetdate_tgl.Enabled = True Then
17710           intRetVal = 0
17720           .assetdate_tgl.SetFocus
17730         End If
17740       End With
17750     End Select
17760   End If

        ' ** Shift keys.
17770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17790       With Me
17800         If .pcash_tgl.Visible = True And .pcash_tgl.Enabled = True Then
17810           intRetVal = 0
17820           .pcash_tgl.SetFocus
17830         End If
17840       End With
17850     End Select
17860   End If

EXITP:
17870   KeyCode = intRetVal
17880   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub cost_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_tgl_MouseMove"

17910   ChkArray_Pop  ' ** Procedure: Below.

17920   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
17930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cost_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_tgl_MouseUp"

18010   ChkArray_Pop  ' ** Procedure: Below.

18020   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
18030   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cost_tgl_LostFocus()

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_tgl_LostFocus"

18110   ChkArray_Pop  ' ** Procedure: Below.

18120   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
18130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cost_GotFocus()

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_GotFocus"

18210   With Me
18220     If blnFromTgl = True Then
18230       blnFromTgl = False
18240       .Cost.SelLength = 0
18250       .Cost.SelStart = 0
18260     End If
18270   End With

EXITP:
18280   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

18310   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
18320   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub curr_id_tgl_GotFocus()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_tgl_GotFocus"

18410   ChkArray_Pop  ' ** Procedure: Below.

18420   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
18430   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub curr_id_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_tgl_MouseDown"

18510   ChkArray_Pop  ' ** Procedure: Below.

18520   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
18530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub curr_id_tgl_Click()

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_tgl_Click"

18610   ChkArray_Pop  ' ** Procedure: Below.

18620   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
18630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub curr_id_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_tgl_KeyDown"

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
18780         If .assetdate_tgl.Visible = True And .assetdate_tgl.Enabled = True Then
18790           intRetVal = 0
18800           .assetdate_tgl.SetFocus
18810         End If
18820       End With
18830     End Select
18840   End If

        ' ** Shift keys.
18850   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18870       With Me
18880         If .cost_tgl.Visible = True And .cost_tgl.Enabled = True Then
18890           intRetVal = 0
18900           .cost_tgl.SetFocus
18910         End If
18920       End With
18930     End Select
18940   End If

EXITP:
18950   KeyCode = intRetVal
18960   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub curr_id_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_tgl_MouseMove"

19010   ChkArray_Pop  ' ** Procedure: Below.

19020   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
19030   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub curr_id_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_tgl_MouseUp"

19110   ChkArray_Pop  ' ** Procedure: Below.

19120   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
19130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub curr_id_tgl_LostFocus()

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_tgl_LostFocus"

19210   ChkArray_Pop  ' ** Procedure: Below.

19220   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
19230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub curr_id_GotFocus()

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_GotFocus"

19310   With Me
19320     If blnFromTgl = True Then
19330       blnFromTgl = False
19340       .curr_id.SelLength = 0
19350       .curr_id.SelStart = 0
19360     End If
19370   End With

EXITP:
19380   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub assetdate_lbl_DblClick(Cancel As Integer)

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_lbl_DblClick"

19410   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19420   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub FocusHolder5_KeyDown(KeyCode As Integer, Shift As Integer)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder5_KeyDown"

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
19590         If .assetdate.Visible = True And .assetdate.Enabled = True Then
19600           .assetdate.SetFocus
19610         Else
19620           .FocusHolder.SetFocus
19630         End If
19640       End With
19650     End Select
19660   End If

        ' ** Shift keys.
19670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19690       With Me
19700         intRetVal = 0
19710         .FocusHolder4.SetFocus
19720       End With
19730     End Select
19740   End If

EXITP:
19750   KeyCode = intRetVal
19760   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Sub

Private Sub assetdate_tgl_GotFocus()

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_tgl_GotFocus"

19810   ChkArray_Pop  ' ** Procedure: Below.

19820   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
19830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub assetdate_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_tgl_MouseDown"

19910   ChkArray_Pop  ' ** Procedure: Below.

19920   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
19930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub assetdate_tgl_Click()

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_tgl_Click"

20010   ChkArray_Pop  ' ** Procedure: Below.

20020   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
20030   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub assetdate_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20120   intShiftDown = (Shift And acShiftMask) > 0
20130   intAltDown = (Shift And acAltMask) > 0
20140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20170       With Me
20180         If .PurchaseDate_tgl.Visible = True And .PurchaseDate_tgl.Enabled = True Then
20190           intRetVal = 0
20200           .PurchaseDate_tgl.SetFocus
20210         End If
20220       End With
20230     End Select
20240   End If

        ' ** Shift keys.
20250   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20270       With Me
20280         If .curr_id_tgl.Visible = True And .curr_id_tgl.Enabled = True Then
20290           intRetVal = 0
20300           .curr_id_tgl.SetFocus
20310         ElseIf .cost_tgl.Visible = True And .cost_tgl.Enabled = True Then
20320           intRetVal = 0
20330           .cost_tgl.SetFocus
20340         End If
20350       End With
20360     End Select
20370   End If

EXITP:
20380   KeyCode = intRetVal
20390   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub assetdate_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_tgl_MouseMove"

20410   ChkArray_Pop  ' ** Procedure: Below.

20420   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
20430   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub assetdate_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_tgl_MouseUp"

20510   ChkArray_Pop  ' ** Procedure: Below.

20520   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
20530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub assetdate_tgl_LostFocus()

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_tgl_LostFocus"

20610   ChkArray_Pop  ' ** Procedure: Below.

20620   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
20630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub assetdate_GotFocus()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_GotFocus"

20710   With Me
20720     If blnFromTgl = True Then
20730       blnFromTgl = False
20740       .assetdate.SelLength = 0
20750       .assetdate.SelStart = 19
20760     End If
20770   End With

EXITP:
20780   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub PurchaseDate_lbl_DblClick(Cancel As Integer)

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDate_lbl_DblClick"

20810   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
20820   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub PurchaseDate_tgl_GotFocus()

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDate_tgl_GotFocus"

20910   ChkArray_Pop  ' ** Procedure: Below.

20920   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
20930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub PurchaseDate_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDate_tgl_MouseDown"

21010   ChkArray_Pop  ' ** Procedure: Below.

21020   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
21030   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub PurchaseDate_tgl_Click()

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDate_tgl_Click"

21110   ChkArray_Pop  ' ** Procedure: Below.

21120   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
21130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub PurchaseDate_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDate_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21220   intShiftDown = (Shift And acShiftMask) > 0
21230   intAltDown = (Shift And acAltMask) > 0
21240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21270       With Me
21280         If .ledger_description_tgl.Visible = True And .ledger_description_tgl.Enabled = True Then
21290           intRetVal = 0
21300           .ledger_description_tgl.SetFocus
21310         End If
21320       End With
21330     End Select
21340   End If

        ' ** Shift keys.
21350   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21370       With Me
21380         If .assetdate_tgl.Visible = True And .assetdate_tgl.Enabled = True Then
21390           intRetVal = 0
21400           .assetdate_tgl.SetFocus
21410         End If
21420       End With
21430     End Select
21440   End If

EXITP:
21450   KeyCode = intRetVal
21460   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub PurchaseDate_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDate_tgl_MouseMove"

21510   ChkArray_Pop  ' ** Procedure: Below.

21520   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
21530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub PurchaseDate_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDate_tgl_MouseUp"

21610   ChkArray_Pop  ' ** Procedure: Below.

21620   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
21630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub PurchaseDate_tgl_LostFocus()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDate_tgl_LostFocus"

21710   ChkArray_Pop  ' ** Procedure: Below.

21720   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
21730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub PurchaseDate_GotFocus()

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDate_GotFocus"

21810   With Me
21820     If blnFromTgl = True Then
21830       blnFromTgl = False
21840       .PurchaseDate.SelLength = 0
21850       .PurchaseDate.SelStart = 19
21860     End If
21870   End With

EXITP:
21880   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_description_lbl_DblClick(Cancel As Integer)

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_lbl_DblClick"

21910   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
21920   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_description_tgl_GotFocus()

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_tgl_GotFocus"

22010   ChkArray_Pop  ' ** Procedure: Below.

22020   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
22030   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_description_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_tgl_MouseDown"

22110   ChkArray_Pop  ' ** Procedure: Below.

22120   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
22130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_description_tgl_Click()

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_tgl_Click"

22210   ChkArray_Pop  ' ** Procedure: Below.

22220   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
22230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_description_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22320   intShiftDown = (Shift And acShiftMask) > 0
22330   intAltDown = (Shift And acAltMask) > 0
22340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22370       With Me
22380         If .RecurringItem_tgl.Visible = True And .RecurringItem_tgl.Enabled = True Then
22390           intRetVal = 0
22400           .RecurringItem_tgl.SetFocus
22410         End If
22420       End With
22430     End Select
22440   End If

        ' ** Shift keys.
22450   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22470       With Me
22480         If .PurchaseDate_tgl.Visible = True And .PurchaseDate_tgl.Enabled = True Then
22490           intRetVal = 0
22500           .PurchaseDate_tgl.SetFocus
22510         End If
22520       End With
22530     End Select
22540   End If

EXITP:
22550   KeyCode = intRetVal
22560   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub ledger_description_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_tgl_MouseMove"

22610   ChkArray_Pop  ' ** Procedure: Below.

22620   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
22630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_description_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_tgl_MouseUp"

22710   ChkArray_Pop  ' ** Procedure: Below.

22720   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
22730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_description_tgl_LostFocus()

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_tgl_LostFocus"

22810   ChkArray_Pop  ' ** Procedure: Below.

22820   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
22830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_description_GotFocus()

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_GotFocus"

22910   With Me
22920     If blnFromTgl = True Then
22930       blnFromTgl = False
22940       .ledger_description.SelLength = 0
22950       .ledger_description.SelStart = 0
22960     End If
22970   End With

EXITP:
22980   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub RecurringItem_lbl_DblClick(Cancel As Integer)

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_lbl_DblClick"

23010   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
23020   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub RecurringItem_tgl_GotFocus()

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_tgl_GotFocus"

23110   ChkArray_Pop  ' ** Procedure: Below.

23120   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
23130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub RecurringItem_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_tgl_MouseDown"

23210   ChkArray_Pop  ' ** Procedure: Below.

23220   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
23230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub RecurringItem_tgl_Click()

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_tgl_Click"

23310   ChkArray_Pop  ' ** Procedure: Below.

23320   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
23330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub RecurringItem_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23420   intShiftDown = (Shift And acShiftMask) > 0
23430   intAltDown = (Shift And acAltMask) > 0
23440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23470       With Me
23480         If .revcode_DESC_tgl.Visible = True And .revcode_DESC_tgl.Enabled = True Then
23490           intRetVal = 0
23500           .revcode_DESC_tgl.SetFocus
23510         End If
23520       End With
23530     End Select
23540   End If

        ' ** Shift keys.
23550   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23570       With Me
23580         If .ledger_description_tgl.Visible = True And .ledger_description_tgl.Enabled = True Then
23590           intRetVal = 0
23600           .ledger_description_tgl.SetFocus
23610         End If
23620       End With
23630     End Select
23640   End If

EXITP:
23650   KeyCode = intRetVal
23660   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub RecurringItem_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_tgl_MouseMove"

23710   ChkArray_Pop  ' ** Procedure: Below.

23720   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
23730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub RecurringItem_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_tgl_MouseUp"

23810   ChkArray_Pop  ' ** Procedure: Below.

23820   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
23830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub RecurringItem_tgl_LostFocus()

23900 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_tgl_LostFocus"

23910   ChkArray_Pop  ' ** Procedure: Below.

23920   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
23930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub RecurringItem_GotFocus()

24000 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_GotFocus"

24010   With Me
24020     If blnFromTgl = True Then
24030       blnFromTgl = False
24040       .RecurringItem.SelLength = 0
24050       .RecurringItem.SelStart = 0
24060     End If
24070   End With

EXITP:
24080   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_DESC_lbl_DblClick(Cancel As Integer)

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_lbl_DblClick"

24110   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
24120   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub FocusHolder6_KeyDown(KeyCode As Integer, Shift As Integer)

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder6_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24220   intShiftDown = (Shift And acShiftMask) > 0
24230   intAltDown = (Shift And acAltMask) > 0
24240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24270       With Me
24280         intRetVal = 0
24290         If .revcode_DESC.Visible = True And .revcode_DESC.Enabled = True Then
24300           .revcode_DESC.SetFocus
24310         Else
24320           .FocusHolder2.SetFocus
24330         End If
24340       End With
24350     End Select
24360   End If

        ' ** Shift keys.
24370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24390       With Me
24400         intRetVal = 0
24410         .FocusHolder5.SetFocus
24420       End With
24430     End Select
24440   End If

EXITP:
24450   KeyCode = intRetVal
24460   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Sub

Private Sub revcode_DESC_tgl_GotFocus()

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_tgl_GotFocus"

24510   ChkArray_Pop  ' ** Procedure: Below.

24520   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
24530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_DESC_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_tgl_MouseDown"

24610   ChkArray_Pop  ' ** Procedure: Below.

24620   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
24630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_DESC_tgl_Click()

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_tgl_Click"

24710   ChkArray_Pop  ' ** Procedure: Below.

24720   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
24730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_DESC_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24820   intShiftDown = (Shift And acShiftMask) > 0
24830   intAltDown = (Shift And acAltMask) > 0
24840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24870       With Me
24880         If .revcode_TYPE_Description_tgl.Visible = True And .revcode_TYPE_Description_tgl.Enabled = True Then
24890           intRetVal = 0
24900           .revcode_TYPE_Description_tgl.SetFocus
24910         End If
24920       End With
24930     End Select
24940   End If

        ' ** Shift keys.
24950   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24970       With Me
24980         If .RecurringItem_tgl.Visible = True And .RecurringItem_tgl.Enabled = True Then
24990           intRetVal = 0
25000           .RecurringItem_tgl.SetFocus
25010         End If
25020       End With
25030     End Select
25040   End If

EXITP:
25050   KeyCode = intRetVal
25060   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub revcode_DESC_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_tgl_MouseMove"

25110   ChkArray_Pop  ' ** Procedure: Below.

25120   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
25130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_DESC_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_tgl_MouseUp"

25210   ChkArray_Pop  ' ** Procedure: Below.

25220   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
25230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_DESC_tgl_LostFocus()

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_tgl_LostFocus"

25310   ChkArray_Pop  ' ** Procedure: Below.

25320   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
25330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_DESC_GotFocus()

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_GotFocus"

25410   With Me
25420     If blnFromTgl = True Then
25430       blnFromTgl = False
25440       .revcode_DESC.SelLength = 0
25450       .revcode_DESC.SelStart = 0
25460     End If
25470   End With

EXITP:
25480   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_TYPE_Description_lbl_DblClick(Cancel As Integer)

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_Description_lbl_DblClick"

25510   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
25520   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_TYPE_Description_tgl_GotFocus()

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_Description_tgl_GotFocus"

25610   ChkArray_Pop  ' ** Procedure: Below.

25620   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
25630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_TYPE_Description_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_Description_tgl_MouseDown"

25710   ChkArray_Pop  ' ** Procedure: Below.

25720   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
25730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_TYPE_Description_tgl_Click()

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_Description_tgl_Click"

25810   ChkArray_Pop  ' ** Procedure: Below.

25820   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
25830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_TYPE_Description_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_Description_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25920   intShiftDown = (Shift And acShiftMask) > 0
25930   intAltDown = (Shift And acAltMask) > 0
25940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25970       With Me
25980         If .taxcode_description_tgl.Visible = True And .taxcode_description_tgl.Enabled = True Then
25990           intRetVal = 0
26000           .taxcode_description_tgl.SetFocus
26010         End If
26020       End With
26030     End Select
26040   End If

        ' ** Shift keys.
26050   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26070       With Me
26080         If .revcode_DESC_tgl.Visible = True And .revcode_DESC_tgl.Enabled = True Then
26090           intRetVal = 0
26100           .revcode_DESC_tgl.SetFocus
26110         End If
26120       End With
26130     End Select
26140   End If

EXITP:
26150   KeyCode = intRetVal
26160   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub revcode_TYPE_Description_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_Description_tgl_MouseMove"

26210   ChkArray_Pop  ' ** Procedure: Below.

26220   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
26230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_TYPE_Description_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

26300 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_Description_tgl_MouseUp"

26310   ChkArray_Pop  ' ** Procedure: Below.

26320   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
26330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_TYPE_Description_tgl_LostFocus()

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_Description_tgl_LostFocus"

26410   ChkArray_Pop  ' ** Procedure: Below.

26420   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
26430   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub revcode_TYPE_Description_GotFocus()

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_Description_GotFocus"

26510   With Me
26520     If blnFromTgl = True Then
26530       blnFromTgl = False
26540       .revcode_TYPE_Description.SelLength = 0
26550       .revcode_TYPE_Description.SelStart = 19
26560     End If
26570   End With

EXITP:
26580   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_description_lbl_DblClick(Cancel As Integer)

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_lbl_DblClick"

26610   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
26620   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_description_tgl_GotFocus()

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_tgl_GotFocus"

26710   ChkArray_Pop  ' ** Procedure: Below.

26720   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
26730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_description_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_tgl_MouseDown"

26810   ChkArray_Pop  ' ** Procedure: Below.

26820   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
26830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_description_tgl_Click()

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_tgl_Click"

26910   ChkArray_Pop  ' ** Procedure: Below.

26920   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
26930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_description_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27020   intShiftDown = (Shift And acShiftMask) > 0
27030   intAltDown = (Shift And acAltMask) > 0
27040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27070       With Me
27080         If .taxcode_type_description_tgl.Visible = True And .taxcode_type_description_tgl.Enabled = True Then
27090           intRetVal = 0
27100           .taxcode_type_description_tgl.SetFocus
27110         End If
27120       End With
27130     End Select
27140   End If

        ' ** Shift keys.
27150   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27170       With Me
27180         If .revcode_TYPE_Description_tgl.Visible = True And .revcode_TYPE_Description_tgl.Enabled = True Then
27190           intRetVal = 0
27200           .revcode_TYPE_Description_tgl.SetFocus
27210         End If
27220       End With
27230     End Select
27240   End If

EXITP:
27250   KeyCode = intRetVal
27260   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub taxcode_description_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_tgl_MouseMove"

27310   ChkArray_Pop  ' ** Procedure: Below.

27320   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
27330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_description_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_tgl_MouseUp"

27410   ChkArray_Pop  ' ** Procedure: Below.

27420   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
27430   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_description_tgl_LostFocus()

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_tgl_LostFocus"

27510   ChkArray_Pop  ' ** Procedure: Below.

27520   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
27530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_description_GotFocus()

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_GotFocus"

27610   With Me
27620     If blnFromTgl = True Then
27630       blnFromTgl = False
27640       .taxcode_description.SelLength = 0
27650       .taxcode_description.SelStart = 0
27660     End If
27670   End With

EXITP:
27680   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_type_description_lbl_DblClick(Cancel As Integer)

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_description_lbl_DblClick"

27710   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
27720   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_type_description_tgl_GotFocus()

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_description_tgl_GotFocus"

27810   ChkArray_Pop  ' ** Procedure: Below.

27820   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
27830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_type_description_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_description_tgl_MouseDown"

27910   ChkArray_Pop  ' ** Procedure: Below.

27920   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
27930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_type_description_tgl_Click()

28000 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_description_tgl_Click"

28010   ChkArray_Pop  ' ** Procedure: Below.

28020   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
28030   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_type_description_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_description_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28120   intShiftDown = (Shift And acShiftMask) > 0
28130   intAltDown = (Shift And acAltMask) > 0
28140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28170       With Me
28180         If .Location_Name_tgl.Visible = True And .Location_Name_tgl.Enabled = True Then
28190           intRetVal = 0
28200           .Location_Name_tgl.SetFocus
28210         ElseIf .CheckNum_tgl.Visible = True And .CheckNum_tgl.Enabled = True Then
28220           intRetVal = 0
28230           .CheckNum_tgl.SetFocus
28240         ElseIf .journal_USER_tgl.Visible = True And .journal_USER_tgl.Enabled = True Then
28250           intRetVal = 0
28260           .journal_USER_tgl.SetFocus
28270         End If
28280       End With
28290     End Select
28300   End If

        ' ** Shift keys.
28310   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28320     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28330       With Me
28340         If .taxcode_description_tgl.Visible = True And .taxcode_description_tgl.Enabled = True Then
28350           intRetVal = 0
28360           .taxcode_description_tgl.SetFocus
28370         End If
28380       End With
28390     End Select
28400   End If

EXITP:
28410   KeyCode = intRetVal
28420   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub taxcode_type_description_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28500 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_description_tgl_MouseMove"

28510   ChkArray_Pop  ' ** Procedure: Below.

28520   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
28530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_type_description_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

28600 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_description_tgl_MouseUp"

28610   ChkArray_Pop  ' ** Procedure: Below.

28620   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
28630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_type_description_tgl_LostFocus()

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_description_tgl_LostFocus"

28710   ChkArray_Pop  ' ** Procedure: Below.

28720   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
28730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub taxcode_type_description_GotFocus()

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_description_GotFocus"

28810   With Me
28820     If blnFromTgl = True Then
28830       blnFromTgl = False
28840       .TaxCode_Type_Description.SelLength = 0
28850       .TaxCode_Type_Description.SelStart = 19
28860     End If
28870   End With

EXITP:
28880   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub Location_Name_lbl_DblClick(Cancel As Integer)

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_Name_lbl_DblClick"

28910   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
28920   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub Location_Name_tgl_GotFocus()

29000 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_Name_tgl_GotFocus"

29010   ChkArray_Pop  ' ** Procedure: Below.

29020   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
29030   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub Location_Name_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

29100 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_Name_tgl_MouseDown"

29110   ChkArray_Pop  ' ** Procedure: Below.

29120   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
29130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub Location_Name_tgl_Click()

29200 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_Name_tgl_Click"

29210   ChkArray_Pop  ' ** Procedure: Below.

29220   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
29230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub Location_Name_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

29300 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_Name_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29320   intShiftDown = (Shift And acShiftMask) > 0
29330   intAltDown = (Shift And acAltMask) > 0
29340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29370       With Me
29380         If .CheckNum_tgl.Visible = True And .CheckNum_tgl.Enabled = True Then
29390           intRetVal = 0
29400           .CheckNum_tgl.SetFocus
29410         ElseIf .journal_USER_tgl.Visible = True And .journal_USER_tgl.Enabled = True Then
29420           intRetVal = 0
29430           .journal_USER_tgl.SetFocus
29440         End If
29450       End With
29460     End Select
29470   End If

        ' ** Shift keys.
29480   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29490     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29500       With Me
29510         If .taxcode_type_description_tgl.Visible = True And .taxcode_type_description_tgl.Enabled = True Then
29520           intRetVal = 0
29530           .taxcode_type_description_tgl.SetFocus
29540         End If
29550       End With
29560     End Select
29570   End If

EXITP:
29580   KeyCode = intRetVal
29590   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub Location_Name_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29600 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_Name_tgl_MouseMove"

29610   ChkArray_Pop  ' ** Procedure: Below.

29620   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
29630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub Location_Name_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

29700 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_Name_tgl_MouseUp"

29710   ChkArray_Pop  ' ** Procedure: Below.

29720   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
29730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub Location_Name_tgl_LostFocus()

29800 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_Name_tgl_LostFocus"

29810   ChkArray_Pop  ' ** Procedure: Below.

29820   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
29830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub Location_Name_GotFocus()

29900 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_Name_GotFocus"

29910   With Me
29920     If blnFromTgl = True Then
29930       blnFromTgl = False
29940       .Location_Name.SelLength = 0
29950       .Location_Name.SelStart = 0
29960     End If
29970   End With

EXITP:
29980   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub CheckNum_lbl_DblClick(Cancel As Integer)

30000 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_lbl_DblClick"

30010   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
30020   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub CheckNum_tgl_GotFocus()

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_tgl_GotFocus"

30110   ChkArray_Pop  ' ** Procedure: Below.

30120   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
30130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub CheckNum_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

30200 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_tgl_MouseDown"

30210   ChkArray_Pop  ' ** Procedure: Below.

30220   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
30230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub CheckNum_tgl_Click()

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_tgl_Click"

30310   ChkArray_Pop  ' ** Procedure: Below.

30320   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
30330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub CheckNum_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_tgl_KeyDown"

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
30480         If .journal_USER_tgl.Visible = True And .journal_USER_tgl.Enabled = True Then
30490           intRetVal = 0
30500           .journal_USER_tgl.SetFocus
30510         End If
30520       End With
30530     End Select
30540   End If

        ' ** Shift keys.
30550   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30570       With Me
30580         If .Location_Name_tgl.Visible = True And .Location_Name_tgl.Enabled = True Then
30590           intRetVal = 0
30600           .Location_Name_tgl.SetFocus
30610         ElseIf .taxcode_type_description_tgl.Visible = True And .taxcode_type_description_tgl.Enabled = True Then
30620           intRetVal = 0
30630           .taxcode_type_description_tgl.SetFocus
30640         End If
30650       End With
30660     End Select
30670   End If

EXITP:
30680   KeyCode = intRetVal
30690   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub CheckNum_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30700 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_tgl_MouseMove"

30710   ChkArray_Pop  ' ** Procedure: Below.

30720   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
30730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub CheckNum_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

30800 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_tgl_MouseUp"

30810   ChkArray_Pop  ' ** Procedure: Below.

30820   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
30830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub CheckNum_tgl_LostFocus()

30900 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_tgl_LostFocus"

30910   ChkArray_Pop  ' ** Procedure: Below.

30920   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
30930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub CheckNum_GotFocus()

31000 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_GotFocus"

31010   With Me
31020     If blnFromTgl = True Then
31030       blnFromTgl = False
31040       .CheckNum.SelLength = 0
31050       .CheckNum.SelStart = 0
31060     End If
31070   End With

EXITP:
31080   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journal_USER_lbl_DblClick(Cancel As Integer)

31100 On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_lbl_DblClick"

31110   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
31120   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journal_USER_tgl_GotFocus()

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_tgl_GotFocus"

31210   ChkArray_Pop  ' ** Procedure: Below.

31220   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
31230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journal_USER_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

31300 On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_tgl_MouseDown"

31310   ChkArray_Pop  ' ** Procedure: Below.

31320   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
31330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journal_USER_tgl_Click()

31400 On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_tgl_Click"

31410   ChkArray_Pop  ' ** Procedure: Below.

31420   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
31430   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journal_USER_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

31500 On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
31520   intShiftDown = (Shift And acShiftMask) > 0
31530   intAltDown = (Shift And acAltMask) > 0
31540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31570       With Me
31580         If .posted_tgl.Visible = True And .posted_tgl.Enabled = True Then
31590           intRetVal = 0
31600           .posted_tgl.SetFocus
31610         End If
31620       End With
31630     End Select
31640   End If

        ' ** Shift keys.
31650   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31670       With Me
31680         If .CheckNum_tgl.Visible = True And .CheckNum_tgl.Enabled = True Then
31690           intRetVal = 0
31700           .CheckNum_tgl.SetFocus
31710         ElseIf .Location_Name_tgl.Visible = True And .Location_Name_tgl.Enabled = True Then
31720           intRetVal = 0
31730           .Location_Name_tgl.SetFocus
31740         ElseIf .taxcode_type_description_tgl.Visible = True And .taxcode_type_description_tgl.Enabled = True Then
31750           intRetVal = 0
31760           .taxcode_type_description_tgl.SetFocus
31770         End If
31780       End With
31790     End Select
31800   End If

EXITP:
31810   KeyCode = intRetVal
31820   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub journal_USER_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

31900 On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_tgl_MouseMove"

31910   ChkArray_Pop  ' ** Procedure: Below.

31920   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
31930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journal_USER_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

32000 On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_tgl_MouseUp"

32010   ChkArray_Pop  ' ** Procedure: Below.

32020   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
32030   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journal_USER_tgl_LostFocus()

32100 On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_tgl_LostFocus"

32110   ChkArray_Pop  ' ** Procedure: Below.

32120   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
32130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub journal_USER_GotFocus()

32200 On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_GotFocus"

        Dim lngLoopCnt As Long

32210   With Me
          ' ** Kluge!
32220     If strToFoc2 = vbNullString Then
32230       strToFoc2 = THIS_PROC
32240       lngLoopCnt = 0&
32250       Do Until Screen.ActiveControl.Name = "FocusHolder2"
32260 On Error Resume Next
32270         .FocusHolder2.SetFocus  ' ** It seems to want to resist this.
32280 On Error GoTo ERRH
32290         lngLoopCnt = lngLoopCnt + 1&
32300         If lngLoopCnt >= 100& Then Exit Do
32310       Loop
32320       .journal_USER.SetFocus
32330     Else
32340       strToFoc2 = vbNullString
32350     End If
32360     .journal_USER.SelLength = 0
32370     .journal_USER.SelStart = 19
32380   End With

EXITP:
32390   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
670     Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Sub

Private Sub posted_lbl_DblClick(Cancel As Integer)

32400 On Error GoTo ERRH

        Const THIS_PROC As String = "posted_lbl_DblClick"

32410   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
32420   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub posted_tgl_GotFocus()

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "posted_tgl_GotFocus"

32510   ChkArray_Pop  ' ** Procedure: Below.

32520   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
32530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub posted_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

32600 On Error GoTo ERRH

        Const THIS_PROC As String = "posted_tgl_MouseDown"

32610   ChkArray_Pop  ' ** Procedure: Below.

32620   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
32630   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub posted_tgl_Click()

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "posted_tgl_Click"

32710   ChkArray_Pop  ' ** Procedure: Below.

32720   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
32730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub posted_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

32800 On Error GoTo ERRH

        Const THIS_PROC As String = "posted_tgl_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
32820   intShiftDown = (Shift And acShiftMask) > 0
32830   intAltDown = (Shift And acAltMask) > 0
32840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
32870       With Me
32880         If .ledger_HIDDEN_tgl.Visible = True And .ledger_HIDDEN_tgl.Enabled = True Then
32890           intRetVal = 0
32900           .ledger_HIDDEN_tgl.SetFocus
32910         End If
32920       End With
32930     End Select
32940   End If

        ' ** Shift keys.
32950   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
32970       With Me
32980         If .journal_USER_tgl.Visible = True And .journal_USER_tgl.Enabled = True Then
32990           intRetVal = 0
33000           .journal_USER_tgl.SetFocus
33010         End If
33020       End With
33030     End Select
33040   End If

EXITP:
33050   KeyCode = intRetVal
33060   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub posted_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

33100 On Error GoTo ERRH

        Const THIS_PROC As String = "posted_tgl_MouseMove"

33110   ChkArray_Pop  ' ** Procedure: Below.

33120   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
33130   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub posted_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

33200 On Error GoTo ERRH

        Const THIS_PROC As String = "posted_tgl_MouseUp"

33210   ChkArray_Pop  ' ** Procedure: Below.

33220   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
33230   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub posted_tgl_LostFocus()

33300 On Error GoTo ERRH

        Const THIS_PROC As String = "posted_tgl_LostFocus"

33310   ChkArray_Pop  ' ** Procedure: Below.

33320   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
33330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub posted_GotFocus()

33400 On Error GoTo ERRH

        Const THIS_PROC As String = "posted_GotFocus"

        Dim lngLoopCnt As Long

33410   With Me
          ' ** Kluge!
33420     If strToFoc2 = vbNullString Then
33430       strToFoc2 = THIS_PROC
33440       lngLoopCnt = 0&
33450       Do Until Screen.ActiveControl.Name = "FocusHolder2"
33460 On Error Resume Next
33470         .FocusHolder2.SetFocus
33480 On Error GoTo ERRH
33490         lngLoopCnt = lngLoopCnt + 1&
33500         If lngLoopCnt >= 100& Then Exit Do
33510       Loop
33520       .posted.SetFocus
33530     Else
33540       strToFoc2 = vbNullString
33550     End If
33560     .posted.SelLength = 0
33570     .posted.SelStart = 19
33580   End With

EXITP:
33590   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
670     Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Sub

Private Sub ledger_HIDDEN_lbl_DblClick(Cancel As Integer)

33600 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_lbl_DblClick"

33610   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
33620   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_HIDDEN_tgl_GotFocus()

33700 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_tgl_GotFocus"

33710   ChkArray_Pop  ' ** Procedure: Below.

33720   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
33730   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_HIDDEN_tgl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

33800 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_tgl_MouseDown"

33810   ChkArray_Pop  ' ** Procedure: Below.

33820   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
33830   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_HIDDEN_tgl_Click()

33900 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_tgl_Click"

33910   ChkArray_Pop  ' ** Procedure: Below.

33920   PrintTgls_Click THIS_PROC, blnFromTgl, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
33930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_HIDDEN_tgl_KeyDown(KeyCode As Integer, Shift As Integer)

34000 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_tgl_KeyDown"

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
34090         .FocusHolder2.SetFocus
34100       End With
34110     End Select
34120   End If

        ' ** Shift keys.
34130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
34150       With Me
34160         If .posted_tgl.Visible = True And .posted_tgl.Enabled = True Then
34170           intRetVal = 0
34180           .posted_tgl.SetFocus
34190         End If
34200       End With
34210     End Select
34220   End If

EXITP:
34230   KeyCode = intRetVal
34240   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub ledger_HIDDEN_tgl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

34300 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_tgl_MouseMove"

34310   ChkArray_Pop  ' ** Procedure: Below.

34320   PrintTgls_Move THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
34330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_HIDDEN_tgl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

34400 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_tgl_MouseUp"

34410   ChkArray_Pop  ' ** Procedure: Below.

34420   PrintTgls_Mouse THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
34430   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_HIDDEN_tgl_LostFocus()

34500 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_tgl_LostFocus"

34510   ChkArray_Pop  ' ** Procedure: Below.

34520   PrintTgls_Focus THIS_PROC, lngChks, arr_varChk, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
34530   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ledger_HIDDEN_GotFocus()

34600 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_GotFocus"

        Dim lngLoopCnt As Long

34610   With Me
          ' ** Kluge!
34620     If strToFoc2 = vbNullString Then
34630       strToFoc2 = THIS_PROC
34640       lngLoopCnt = 0&
34650       Do Until Screen.ActiveControl.Name = "FocusHolder2"
34660 On Error Resume Next
34670         .FocusHolder2.SetFocus
34680 On Error GoTo ERRH
34690         lngLoopCnt = lngLoopCnt + 1&
34700         If lngLoopCnt >= 100& Then Exit Do
34710       Loop
34720     Else
34730       strToFoc2 = vbNullString
34740     End If
34750   End With

EXITP:
34760   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
670     Case 2474  ' ** The expression you entered requires the control to be in the active window.
          ' ** Ignore
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub ledger_HIDDEN_KeyDown(KeyCode As Integer, Shift As Integer)

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_KeyDown"

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
34890         lngRecsCur = RecCnt  ' ** Function: Below.
34900         NextFocus THIS_PROC, True, lngFlds, arr_varFld, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
34910       End With
34920     End Select
34930   End If

        ' ** Shift keys.
34940   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34950     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
34960       With Me
34970         intRetVal = 0
34980         NextFocus THIS_PROC, False, lngFlds, arr_varFld, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
34990       End With
35000     End Select
35010   End If

EXITP:
35020   KeyCode = intRetVal
35030   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Sub

Private Sub FocusHolder2_GotFocus()

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder2_GotFocus"

35110   With Me
          ' ** Kluge!
35120     If strToFoc2 <> vbNullString Then
35130       .TimerInterval = 100&
35140     End If
35150   End With

EXITP:
35160   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub FocusHolder2_KeyDown(KeyCode As Integer, Shift As Integer)

35200 On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder2_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
35270       With Me
35280         intRetVal = 0
35290         lngRecsCur = RecCnt  ' ** Function: Below.
35300         If .CurrentRecord < lngRecsCur Then
35310           MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
35320           .FocusHolder.SetFocus
                ' ** It keeps wanting to come back here! WHY?
35330           .ledger_HIDDEN.Enabled = False
35340         Else
35350           DoCmd.SelectObject acForm, .Parent.Name, False
35360           .Parent.cmdSelectAll.SetFocus
35370         End If
35380       End With
35390     End Select
35400   End If

        ' ** Shift keys.
35410   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35420     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
35430       With Me
35440         intRetVal = 0
35450         NextFocus "ledger_HIDDEN_KeyDown", False, lngFlds, arr_varFld, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
35460       End With
35470     End Select
35480   End If

EXITP:
35490   KeyCode = intRetVal
35500   Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Sub

Public Function FilterRecs_Cnt(rstAll As DAO.Recordset) As Boolean

35600 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Cnt"

        Dim blnRetVal As Boolean

35610   blnRetVal = FilterRecs_Cnt_TA(strFilter, dblFilterRecs, rstAll, Me)  ' ** Module Function: modTransactionAuditFuncs1.

EXITP:
35620   FilterRecs_Cnt = blnRetVal
35630   Exit Function

ERRH:
660     blnRetVal = False
670     DoCmd.Hourglass False
680     Select Case ERR.Number
        Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Function

Public Sub FilterRecs_Rem(strClause As String)
' ** Remove a clause from the filter string.

35700 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Rem"

35710   FilterRecs_Rem_TA strClause, strFilter, dblFilterRecs, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
35720   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Public Sub FilterRecs_Clr(strProc As String, blnOn As Boolean)

35800 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Clr"

35810   FilterRecs_Clr_TA strProc, blnOn, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
35820   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Public Sub FilterRecs_Set(strFltr As String, dblFltrRecs As Double)
' ** Initialize the local strFilter and dblFilterRecs with frmTransaction_Audit's copy.

35900 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Set"

35910   strFilter = strFltr
35920   dblFilterRecs = dblFltrRecs

EXITP:
35930   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Public Sub TotalRecs_Set(lngTotRecs As Long)

36000 On Error GoTo ERRH

        Const THIS_PROC As String = "TotalRecs_Set"

36010   lngTotalRecs = lngTotRecs

EXITP:
36020   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Public Sub Print_Chk(strProc As String)

36100 On Error GoTo ERRH

        Const THIS_PROC As String = "Print_Chk"

36110   Print_Chk_TA strProc, 1, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
36120   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Public Function Print_ChkCnt() As Long

36200 On Error GoTo ERRH

        Const THIS_PROC As String = "Print_ChkCnt"

        Dim lngRetVal As Long

36210   lngRetVal = Print_ChkCnt_TA(1, Me)   ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
36220   Print_ChkCnt = lngRetVal
36230   Exit Function

ERRH:
660     lngRetVal = 0&
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Function

Public Sub ChkArray_Set(arr_varTmp01 As Variant)

36300 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkArray_Set"

36310   arr_varChk = arr_varTmp01
36320   lngChks = UBound(arr_varChk, 2) + 1&

EXITP:
36330   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub ChkArray_Pop()

36400 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkArray_Pop"

36410   If lngChks = 0& Or IsEmpty(arr_varChk) = True Then
36420     arr_varChk = PrintTgls_GetArr  ' ** Module Procedure: modTransactionAuditFuncs1.
36430     If IsEmpty(arr_varChk) = True Then
36440       PrintTgls_Load Me  ' ** Module Procedure: modTransactionAuditFuncs1.
36450       DoEvents
36460       If lngChks = 0& Then
36470         arr_varChk = PrintTgls_GetArr  ' ** Module Procedure: modTransactionAuditFuncs1.
36480         lngChks = UBound(arr_varChk, 2) + 1&
36490       End If
36500     Else
36510       lngChks = UBound(arr_varChk, 2) + 1&
36520     End If
36530   End If

EXITP:
36540   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Public Sub ShowFields_Sub(strProc As String, intMode As Integer)
' ** This will either load the field array, or update it.

36600 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowFields_Sub"

36610   ShowFields_Sub_TA strProc, intMode, lngFlds, arr_varFld, lngFldSep, blnFromTgl, lngSortLbl_Width, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
36620   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Public Function FldArray_Get() As Variant
' ** This allows another module to retrieve the field array from here.

36700 On Error GoTo ERRH

        Const THIS_PROC As String = "FldArray_Get"

        Dim arr_varRetVal As Variant

36710   arr_varFld = FldArray_Get_TA  ' ** Module Procedure: modTransactionAuditFuncs1.
36720   arr_varRetVal = arr_varFld

EXITP:
36730   FldArray_Get = arr_varRetVal
36740   Exit Function

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Function

Public Function CalcSpace(strProc As String) As Boolean

36800 On Error GoTo ERRH

        Const THIS_PROC As String = "CalcSpace"

        Dim strCtlName As String
        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

36810   With Me
36820     blnRetVal = False
36830     strCtlName = Left(strProc, (CharPos(strProc, 2, "_") - 1))  ' ** Module Function: modStringFuncs.
36840     lngTmp01 = .journalno.Left
36850     lngTmp02 = 0&
          ' ** Make sure we've got an up-to-date array.
36860     arr_varFld = FldArray_Get_TA  ' ** Module Procedure: modTransactionAuditFuncs1.
36870     For lngX = 0& To (lngFlds - 1&)
36880       If arr_varFld(F_VIS, lngX) = True Then
36890         lngTmp01 = (lngTmp01 + arr_varFld(F_WDT, lngX) + lngFldSep)
36900       ElseIf arr_varFld(F_CNAM, lngX) = strCtlName Then
36910         lngTmp02 = (arr_varFld(F_WDT, lngX) + lngFldSep)
36920       End If
36930     Next  ' ** lngX.
36940     If (lngTmp01 + lngTmp02) <= lngForm_Width Then
36950       blnRetVal = True
36960     End If
36970   End With

EXITP:
36980   CalcSpace = blnRetVal
36990   Exit Function

ERRH:
660     blnRetVal = False
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Function

Public Sub FirstFocus()

37000 On Error GoTo ERRH

        Const THIS_PROC As String = "FirstFocus"

        Dim lngX As Long

37010   With Me
37020     For lngX = 0& To (lngFlds - 1&)
37030       If arr_varFld(F_VIS, lngX) = True Then
37040         .Controls(arr_varFld(F_FNAM, lngX)).SetFocus
37050         Exit For
37060       End If
37070     Next
37080   End With

EXITP:
37090   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Public Function RecCnt() As Long

37100 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

37110   lngRetVal = 0&

37120   With Me
37130     Set rst = .RecordsetClone
37140     With rst
37150       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
37160       Else
37170         .MoveLast
37180         lngRetVal = .RecordCount
37190       End If
37200       .Close
37210     End With
37220   End With

EXITP:
37230   Set rst = Nothing
37240   RecCnt = lngRetVal
37250   Exit Function

ERRH:
660     lngRetVal = 0&
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

37300 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

37310   Select Case intWhere
        Case acCmdRecordsGoToFirst
37320     DoCmd.RunCommand acCmdRecordsGoToFirst
37330   Case acCmdRecordsGoToLast
37340     DoCmd.RunCommand acCmdRecordsGoToLast
37350   Case acCmdRecordsGoToPrevious
37360     DoCmd.RunCommand acCmdRecordsGoToPrevious
37370   Case acCmdRecordsGoToNext
37380     DoCmd.RunCommand acCmdRecordsGoToNext
37390   Case Else
37400     If IsMissing(varID) = False Then
37410       With Me
37420         Set rst = .RecordsetClone
37430         With rst
37440           .FindFirst "[journalno] = " & CStr(varID)
37450           If .NoMatch = False Then
37460             Me.Bookmark = .Bookmark
37470           End If
37480           .Close
37490         End With
37500       End With
37510     End If
37520   End Select

EXITP:
37530   Set rst = Nothing
37540   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
670     Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
680     Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

37600 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim ctl As Access.Control
        Dim strCalled As String, strSortAsc As String
        Dim blnFormView As Boolean
        Dim intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [alphasort], [journaltype], [journalno]"

37610   With Me
37620     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
37630       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
37640     End If
37650     Select Case .CurrentView
          Case acCurViewFormBrowse
37660       blnFormView = True
37670     Case acCurViewDatasheet
37680       blnFormView = False
37690     End Select
37700     .Controls(strSortLbl).Visible = False
37710     .Controls(strSortLine).Visible = False
37720     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push of the right side of the form.
37730     If strProc = "Form_Load" Then
37740       strCalled = "journalno"
37750       strSortNow = strSortOrig
37760       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
37770       lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (6& * lngTpp))
37780       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
37790       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
37800       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
37810       If blnFormView = True Then
37820         .Controls(strSortLbl).Top = lngSortLbl_Top
37830         .Controls(strSortLbl).Left = lngSortLbl_Left
37840         .Controls(strSortLine).Top = lngSortLine_Top
37850         .Controls(strSortLine).Left = lngSortLine_Left
37860         .Controls(strSortLine).Width = lngSortLine_Width
37870       End If
37880       .Controls(strSortLbl).Caption = strArwUp
37890       .Controls(strSortLbl).ForeColor = CLR_DKBLU
37900     Else
37910       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
37920       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
37930       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
37940       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
37950       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
37960       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
37970       If blnFormView = True Then
37980         .Controls(strSortLbl).Top = lngSortLbl_Top
37990         .Controls(strSortLbl).Left = lngSortLbl_Left
38000         .Controls(strSortLine).Top = lngSortLine_Top
38010         .Controls(strSortLine).Left = lngSortLine_Left
38020         .Controls(strSortLine).Width = lngSortLine_Width
38030       End If
38040       If strCalled = "journalno" Then
38050         strSortAsc = "[journalno] DESC"
38060         If strSortNow = strSortOrig Then
38070           strSortNow = strSortAsc
38080           .Controls(strSortLbl).Caption = strArwDn
38090           .Controls(strSortLbl).ForeColor = CLR_DKRED
38100         Else
38110           strSortNow = strSortOrig
38120           .Controls(strSortLbl).Caption = strArwUp
38130           .Controls(strSortLbl).ForeColor = CLR_DKBLU
38140         End If
38150         lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
38160         .Controls(strSortLbl).Left = lngSortLbl_Left
38170       Else
38180         .Controls(strSortLbl).Caption = strArwUp
38190         .Controls(strSortLbl).ForeColor = CLR_DKBLU
38200         Select Case strCalled
              Case "journaltype"
38210           strSortAsc = "[journaltype]" & strStdAsc
38220         Case "transdate"
38230           strSortAsc = "[transdate]" & strStdAsc
38240           lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
38250           .Controls(strSortLbl).Left = lngSortLbl_Left
38260         Case "accountno"
38270           strSortAsc = "[alphasort], [journaltype], [journalno]"
38280         Case "shortname"
38290           strSortAsc = "[shortname]" & strStdAsc
38300         Case "cusip"
38310           strSortAsc = "[cusip]" & strStdAsc
38320         Case "asset_description"
38330           strSortAsc = "[asset_description]" & strStdAsc
38340         Case "shareface"
38350           strSortAsc = "[shareface]" & strStdAsc
38360           lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
38370           .Controls(strSortLbl).Left = lngSortLbl_Left
38380         Case "icash"
38390           strSortAsc = "[icash]" & strStdAsc
38400           lngSortLbl_Left = (lngSortLbl_Left + (1& * lngTpp))
38410           .Controls(strSortLbl).Left = lngSortLbl_Left
38420         Case "pcash"
38430           strSortAsc = "[pcash]" & strStdAsc
38440           lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
38450           .Controls(strSortLbl).Left = lngSortLbl_Left
38460         Case "cost"
38470           strSortAsc = "[cost]" & strStdAsc
38480         Case "curr_id"
38490           strSortAsc = "[curr_code]" & strStdAsc
38500           lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
38510           lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
38520           lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
38530           .Controls(strSortLbl).Left = lngSortLbl_Left
38540           .Controls(strSortLine).Left = lngSortLine_Left
38550           .Controls(strSortLine).Width = lngSortLine_Width
38560         Case "assetdate"
38570           strSortAsc = "[assetdate]" & strStdAsc
38580         Case "PurchaseDate"
38590           strSortAsc = "[PurchaseDate]" & strStdAsc
38600           lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
38610           lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
38620           lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
38630           .Controls(strSortLbl).Left = lngSortLbl_Left
38640           .Controls(strSortLine).Left = lngSortLine_Left
38650           .Controls(strSortLine).Width = lngSortLine_Width
38660         Case "ledger_description"
38670           strSortAsc = "[ledger_description]" & strStdAsc
38680         Case "RecurringItem"
38690           strSortAsc = "[RecurringItem]" & strStdAsc
38700         Case "revcode_DESC"
38710           strSortAsc = "[revcode_DESC]" & strStdAsc
38720         Case "revcode_TYPE_Description"
38730           strSortAsc = "[revcode_TYPE_Description], [revcode_DESC]" & strStdAsc
38740           lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
38750           .Controls(strSortLbl).Left = lngSortLbl_Left
38760         Case "taxcode_description"
38770           strSortAsc = "[taxcode_description]" & strStdAsc
38780         Case "taxcode_type_description"
38790           strSortAsc = "[taxcode_type_description], [taxcode_description]" & strStdAsc
38800         Case "Location_Name"
38810           strSortAsc = "[Location_Name]" & strStdAsc
38820         Case "CheckNum"
38830           strSortAsc = "[CheckNum]" & strStdAsc
38840           lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
38850           .Controls(strSortLbl).Left = lngSortLbl_Left
38860         Case "journal_USER"
38870           strSortAsc = "[journal_USER]" & strStdAsc
38880         Case "posted"
38890           strSortAsc = "[posted]" & strStdAsc
38900         Case "ledger_HIDDEN"
38910           strSortAsc = "[ledger_HIDDEN]" & strStdAsc
38920           .Controls(strSortLbl).Width = (lngSortLbl_Width - (1& * lngTpp))
38930           lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
38940           .Controls(strSortLbl).Left = lngSortLbl_Left
38950         End Select
38960         If strSortNow = strSortAsc Then
38970           intCnt = (CharCnt(strSortAsc, ",") + 1)  ' ** Module Function: modStringFuncs.
38980           Select Case intCnt
                Case 1
38990             strTmp01 = strSortAsc & " DESC"
39000           Case 2
39010             intPos01 = InStr(strSortAsc, ",")
39020             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
39030             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
39040           Case 3
39050             intPos01 = InStr(strSortAsc, ",")
39060             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
39070             strTmp02 = Mid(strSortAsc, intPos01)
39080             intPos01 = InStr(2, strTmp02, ",")
39090             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
39100             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
39110           Case 4
39120             intPos01 = InStr(strSortAsc, ",")
39130             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
39140             strTmp02 = Mid(strSortAsc, intPos01)
39150             intPos01 = InStr(2, strTmp02, ",")
39160             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
39170             strTmp02 = Mid(strTmp02, intPos01)
39180             intPos01 = InStr(2, strTmp02, ",")
39190             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
39200             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
39210           Case 5
39220             intPos01 = InStr(strSortAsc, ",")
39230             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
39240             strTmp02 = Mid(strSortAsc, intPos01)
39250             intPos01 = InStr(2, strTmp02, ",")
39260             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
39270             strTmp02 = Mid(strTmp02, intPos01)
39280             intPos01 = InStr(2, strTmp02, ",")
39290             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
39300             strTmp02 = Mid(strTmp02, intPos01)
39310             intPos01 = InStr(2, strTmp02, ",")
39320             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
39330             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
39340           End Select
39350           strSortNow = strTmp01
39360           .Controls(strSortLbl).Caption = strArwDn
39370           .Controls(strSortLbl).ForeColor = CLR_DKRED
39380         Else
39390           strSortNow = strSortAsc
39400         End If
39410       End If
39420     End If
39430     .Controls(strSortLbl).Visible = True
39440     .Controls(strSortLine).Visible = True
39450 On Error Resume Next
39460     .OrderBy = strSortNow
39470     .OrderByOn = True
39480 On Error GoTo ERRH
39490   End With

        ' ** Potential Sorts:
        ' **   [journalno]
        ' **   [journaltype], [alphasort], [journaltype], [journalno]
        ' **   [transdate], [alphasort], [journaltype], [journalno]
        ' **   [alphasort], [journaltype], [journalno]
        ' **   [shortname], [alphasort], [journaltype], [journalno]
        ' **   [cusip], [alphasort], [journaltype], [journalno]
        ' **   [asset_description], [alphasort], [journaltype], [journalno]
        ' **   [shareface], [alphasort], [journaltype], [journalno]
        ' **   [icash], [alphasort], [journaltype], [journalno]
        ' **   [pcash], [alphasort], [journaltype], [journalno]
        ' **   [cost], [alphasort], [journaltype], [journalno]
        ' **   [curr_code], [alphasort], [journaltype], [journalno]
        ' **   [assetdate], [alphasort], [journaltype], [journalno]
        ' **   [PurchaseDate], [alphasort], [journaltype], [journalno]
        ' **   [ledger_description], [alphasort], [journaltype], [journalno]
        ' **   [RecurringItem], [alphasort], [journaltype], [journalno]
        ' **   [revcode_DESC], [alphasort], [journaltype], [journalno]
        ' **   [revcode_TYPE_Description], [revcode_DESC], [alphasort], [journaltype], [journalno]
        ' **   [taxcode_description], [alphasort], [journaltype], [journalno]
        ' **   [taxcode_type_description], [taxcode_description], [alphasort], [journaltype], [journalno]
        ' **   [Location_Name], [alphasort], [journaltype], [journalno]
        ' **   [CheckNum], [alphasort], [journaltype], [journalno]
        ' **   [journal_USER], [alphasort], [journaltype], [journalno]
        ' **   [posted], [alphasort], [journaltype], [journalno]
        ' **   [ledger_HIDDEN], [alphasort], [journaltype], [journalno]

EXITP:
39500   Set ctl = Nothing
39510   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Public Function SortNow_Get() As String

39600 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

39610   strRetVal = strSortNow

EXITP:
39620   SortNow_Get = strRetVal
39630   Exit Function

ERRH:
660     strRetVal = vbNullString
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Function

Public Sub SortNow_Set(strSort As String)

39700 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Set"

39710   strSortNow = strSort

EXITP:
39720   Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub
