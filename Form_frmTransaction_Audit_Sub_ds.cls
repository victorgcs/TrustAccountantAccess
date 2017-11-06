VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTransaction_Audit_Sub_ds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmTransaction_Audit_Sub_ds"

'VGC 10/05/2017: CHANGES!

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

'THE DATASHEET MAY NEED ITS VISIBLE FIELD NAMES INSTEAD OF THE TABLE FIELD NAMES IN THE FILTER!
'Journal Number
'Journal Type
'Posting Date
'Account Number
'Name
'CUSIP
'Asset
'Share / Face
'Income Cash
'Principal Cash
'Cost
'Trade Date
'Original Trade Date
'Comments
'Recurring Item
'Inc/Exp Codes
'Inc/Exp
'Tax Codes
'Inc/Ded
'User
'Date Posted
'Hidden

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const ANDF As String = " And "  ' ** Filter 'And'.
Private Const ORF As String = " Or "  ' ** Filter 'Or'.

' ** Array: arr_varFilt_ds().
Private lngFilts_ds As Long, arr_varFilt_ds As Variant
'Private Const F_ELEMS As Integer = 13  ' ** Array's first-element UBound().
'Private Const F_IDX   As Integer = 0
'Private Const F_NAM   As Integer = 1
'Private Const F_CONST As Integer = 2
Private Const F_CTL   As Integer = 3
'Private Const F_CLBL  As Integer = 4
Private Const F_FLD   As Integer = 5
Private Const F_FLBL  As Integer = 6
Private Const F_CTL2  As Integer = 7
Private Const F_CLBL2 As Integer = 8
Private Const F_FLD2  As Integer = 9
Private Const F_FLBL2 As Integer = 10
Private Const F_CLBL3 As Integer = 11
Private Const F_FLD3  As Integer = 12
Private Const F_FLBL3 As Integer = 13

' ** Array: arr_varCol().
Private lngCols As Long, arr_varCol() As Variant
Private Const C_ELEMS As Integer = 4  ' ** Array's first-element UBound().
Private Const C_VIEWCHK As Integer = 0
Private Const C_VIEWVAL As Integer = 1
Private Const C_COLNAM  As Integer = 2
Private Const C_HASSP   As Integer = 3
Private Const C_WDT     As Integer = 4

Private Const strSortOrig As String = "[journalno]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
'Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
'Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String

Private frmCrit As Access.Form
Private lngTotalRecs As Long, lngRecsCur As Long, lngTpp As Long
Private strFilter As String, dblFilterRecs As Double, blnColWidth As Boolean
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
180         blnColWidth = False

190         Set frmCrit = .Parent.frmTransaction_Audit_Sub_Criteria.Form

200         CLR_DISABLED_FG = CLR_DKGRY
210         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
220         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

230       Else
240         Cancel = -1
250       End If
260     End With

270     If Cancel = -1 Then
280       Beep
290       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
300     End If

EXITP:
310     Exit Sub

ERRH:
320     Select Case ERR.Number
        Case Else
330       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
340     End Select
350     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim ctl As Access.Control
        Dim lngX As Long

410     With Me

420       Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
430       DoEvents

440       For Each ctl In .FormHeader.Controls
450         With ctl
460           If .ControlType = acCheckBox And .Visible = True Then
470             Print_Chk (.Name & "_AfterUpdate")  ' ** Procedure: Below.
480           End If
490         End With
500       Next

510       SortNow THIS_PROC  ' ** Procedure: Below.
520       DoEvents

          ' ** Hide columns that shouldn't be visible.
530       .alphasort.ColumnHidden = True
540       .curr_id.ColumnHidden = True
550       .curr_code.ColumnHidden = True
560       .Location_ID.ColumnHidden = True
570       .Location_Name.ColumnHidden = True
580       .CheckNum.ColumnHidden = True

          'HAVE PREFERENCES BEEN LOADED IN THE PARENT?
590       ColArray_Load  ' ** Procedure: Below.

          ' ** Hide columns based on check boxes.
600       For lngX = 0& To (lngCols - 1&)
610         If arr_varCol(C_VIEWVAL, lngX) = False Then
620           .Controls(arr_varCol(C_COLNAM, lngX)).ColumnHidden = True
630         End If
640       Next

          ' ** Column widths handled in Form_Timer().

650     End With
660     DoEvents

EXITP:
670     Set ctl = Nothing
680     Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim ctl As Access.Control
        Dim intRetVal As Integer

810     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
820     intShiftDown = (Shift And acShiftMask) > 0
830     intAltDown = (Shift And acAltMask) > 0
840     intCtrlDown = (Shift And acCtrlMask) > 0

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
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyF5
            ' **   Date Picker Posting - From:  F5  {cmdCalendar1}
870         With Me
880           intRetVal = 0
890           If frmCrit.cmdCalendar1.Enabled = True Then
900             DoCmd.SelectObject acForm, .Parent.Name, False
910             .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
920             frmCrit.cmdCalendar1.SetFocus
930             frmCrit.cmdCalendar1_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
940           End If
950         End With
960       Case vbKeyF6
            ' **   Date Picker Posting - To:    F6  {cmdCalendar2}
970         With Me
980           intRetVal = 0
990           If frmCrit.cmdCalendar2.Enabled = True Then
1000            DoCmd.SelectObject acForm, .Parent.Name, False
1010            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1020            frmCrit.cmdCalendar2.SetFocus
1030            frmCrit.cmdCalendar2_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1040          End If
1050        End With
1060      Case vbKeyF7
            ' **   Date Picker Trade - From:    F7  {cmdCalendar3}
1070        With Me
1080          intRetVal = 0
1090          If frmCrit.cmdCalendar3.Enabled = True Then
1100            DoCmd.SelectObject acForm, .Parent.Name, False
1110            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1120            frmCrit.cmdCalendar3.SetFocus
1130            frmCrit.cmdCalendar3_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1140          End If
1150        End With
1160      Case vbKeyF8
            ' **   Date Picker Trade - To:      F8  {cmdCalendar4}
1170        With Me
1180          intRetVal = 0
1190          If frmCrit.cmdCalendar4.Enabled = True Then
1200            DoCmd.SelectObject acForm, .Parent.Name, False
1210            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1220            frmCrit.cmdCalendar4.SetFocus
1230            frmCrit.cmdCalendar4_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1240          End If
1250        End With
1260      Case vbKeyF9
            ' **   Date Picker Purchase - From: F9  {cmdCalendar5}
1270        With Me
1280          intRetVal = 0
1290          If frmCrit.cmdCalendar5.Enabled = True Then
1300            DoCmd.SelectObject acForm, .Parent.Name, False
1310            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1320            frmCrit.cmdCalendar5.SetFocus
1330            frmCrit.cmdCalendar5_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1340          End If
1350        End With
1360      Case vbKeyF10
            ' **   Date Picker Purchase - To:   F10 {cmdCalendar6}
1370        With Me
1380          intRetVal = 0
1390          If frmCrit.cmdCalendar6.Enabled = True Then
1400            DoCmd.SelectObject acForm, .Parent.Name, False
1410            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1420            frmCrit.cmdCalendar6.SetFocus
1430            frmCrit.cmdCalendar6_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1440          End If
1450        End With
1460      Case vbKeyF11
            ' **   Date Picker Posted - From:   F11 {cmdCalendar7}
1470        With Me
1480          intRetVal = 0
1490          If frmCrit.cmdCalendar7.Enabled = True Then
1500            DoCmd.SelectObject acForm, .Parent.Name, False
1510            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1520            frmCrit.cmdCalendar7.SetFocus
1530            frmCrit.cmdCalendar7_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1540          End If
1550        End With
1560      Case vbKeyF12
            ' **   Date Picker Posted - To:     F12 {cmdCalendar8}
1570        With Me
1580          intRetVal = 0
1590          If frmCrit.cmdCalendar8.Enabled = True Then
1600            DoCmd.SelectObject acForm, .Parent.Name, False
1610            .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1620            frmCrit.cmdCalendar8.SetFocus
1630            frmCrit.cmdCalendar8_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1640          End If
1650        End With
1660      Case vbKeyUp
1670        intRetVal = 0
1680        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1690      Case vbKeyDown
1700        intRetVal = 0
1710        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1720      End Select
1730    End If

        ' ** Alt keys.
1740    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1750      Select Case intRetVal
          Case vbKeyD
1760        With Me
1770          intRetVal = 0
1780          DoCmd.SelectObject acForm, .Parent.Name, False
1790          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
1800          frmCrit.TransDateStart.SetFocus
1810        End With
1820      Case vbKeyH
1830        With Me
1840          intRetVal = 0
1850          If .Parent.chkIncludeArchive.Enabled = True Then
1860            DoCmd.SelectObject acForm, .Parent.Name, False
1870            .Parent.chkIncludeArchive.SetFocus
1880            .Parent.chkIncludeArchive = (Not .Parent.chkIncludeArchive)
1890            .Parent.chkIncludeArchive_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit.
1900          Else
1910            Beep
1920          End If
1930        End With
1940      Case vbKeyJ
1950        With Me
1960          intRetVal = 0
1970          If .[Journal Number].Visible = True And .[Journal Number].Enabled = True Then
1980            .[Journal Number].SetFocus
1990          Else
2000            Beep
2010          End If
2020        End With
2030      Case vbKeyK
2040        With Me
2050          intRetVal = 0
2060          If .Parent.cmdWidenToCriteria.Enabled = True Then
2070            DoCmd.SelectObject acForm, .Parent.Name, False
2080            .Parent.cmdWidenToCriteria.SetFocus
2090            .Parent.cmdWidenToCriteria_Click  ' ** Form Procedure: frmTransaction_Audit.
2100          Else
2110            Beep
2120          End If
2130        End With
2140      Case vbKeyR
2150        With Me
2160          If dblFilterRecs = 0# Then
2170            DoCmd.SelectObject acForm, .Parent.Name, False
2180            .Parent.cmdClearAll.SetFocus
2190            .Parent.cmdClearAll_Click  ' ** Form Procedure: frmTransaction_Audit.
2200          End If
2210        End With
2220      Case vbKeyX
2230        With Me
2240          intRetVal = 0
2250          .Parent.cmdClose_Click  ' ** Form Procedure: frmTransaction_Audit.
2260        End With
2270      End Select
2280    End If

        ' ** Alt-Shift keys.
2290    If (Not intCtrlDown) And intAltDown And intShiftDown Then
2300      Select Case intRetVal
          Case vbKeyC
2310        With Me
2320          intRetVal = 0
2330          DoCmd.SelectObject acForm, .Parent.Name, False
2340          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
2350          frmCrit.opgHidden.SetFocus
2360          frmCrit.opgHidden = frmCrit.opgHidden_optInclude.OptionValue
2370          frmCrit.opgHidden_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
2380        End With
2390      Case vbKeyD
2400        With Me
2410          intRetVal = 0
2420          DoCmd.SelectObject acForm, .Parent.Name, False
2430          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
2440          frmCrit.chkTaxcodeType_Deduction.SetFocus
2450          frmCrit.chkTaxcodeType_Deduction = (Not frmCrit.chkTaxcodeType_Deduction)
2460          frmCrit.chkTaxcodeType_Deduction_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
2470        End With
2480      Case vbKeyE
2490        With Me
2500          intRetVal = 0
2510          DoCmd.SelectObject acForm, .Parent.Name, False
2520          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
2530          frmCrit.opgHidden.SetFocus
2540          frmCrit.opgHidden = frmCrit.opgHidden_optExclude.OptionValue
2550          frmCrit.opgHidden_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
2560        End With
2570      Case vbKeyI
2580        With Me
2590          intRetVal = 0
2600          DoCmd.SelectObject acForm, .Parent.Name, False
2610          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
2620          frmCrit.chkTaxcodeType_Income.SetFocus
2630          frmCrit.chkTaxcodeType_Income = (Not frmCrit.chkTaxcodeType_Income)
2640          frmCrit.chkTaxcodeType_Income_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria
2650        End With
2660      Case vbKeyM
2670        With Me
2680          intRetVal = 0
2690          DoCmd.SelectObject acForm, .Parent.Name, False
2700          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
2710          frmCrit.opgAssetSource = frmCrit.opgAssetSource_optName.OptionValue
2720          frmCrit.opgAssetSource.SetFocus
2730          frmCrit.opgAssetSource_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
2740        End With
2750      Case vbKeyN
2760        With Me
2770          intRetVal = 0
2780          DoCmd.SelectObject acForm, .Parent.Name, False
2790          .Parent.chkNarrowFont.SetFocus
2800          .Parent.chkNarrowFont = (Not .Parent.chkNarrowFont)
2810          .Parent.chkNarrowFont_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit.
2820        End With
2830      Case vbKeyO
2840        With Me
2850          intRetVal = 0
2860          DoCmd.SelectObject acForm, .Parent.Name, False
2870          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
2880          frmCrit.opgHidden.SetFocus
2890          frmCrit.opgHidden = frmCrit.opgHidden_optOnly.OptionValue
2900          frmCrit.opgHidden_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
2910        End With
2920      Case vbKeyT
2930        With Me
2940          intRetVal = 0
2950          DoCmd.SelectObject acForm, .Parent.Name, False
2960          .Parent.chkIncludeSharefaceTot.SetFocus
2970          .Parent.chkIncludeSharefaceTot = (Not .Parent.chkIncludeSharefaceTot)
2980          .Parent.chkIncludeSharefaceTot_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit.
2990        End With
3000      End Select
3010    End If

        ' ** Ctrl keys.
3020    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3030      Select Case intRetVal
          Case vbKeyG
3040        With Me
3050          intRetVal = 0
3060          DoCmd.SelectObject acForm, .Parent.Name, False
3070          .Parent.chkPageOf.SetFocus
3080          .Parent.chkPageOf = (Not .Parent.chkPageOf)
3090          .Parent.chkPageOf_AfterUpdate  ' ** Procedure: frmTransaction_Audit.
3100        End With
3110      Case vbKeyTab, vbKeyReturn
3120        With Me
3130          intRetVal = 0
3140          DoCmd.SelectObject acForm, .Parent.Name, False
3150          .Parent.cmdSelectAll.SetFocus
3160        End With
3170      Case vbKeyPageUp, vbKeyHome
3180        intRetVal = 0
3190        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
3200      Case vbKeyPageDown, vbKeyEnd
3210        intRetVal = 0
3220        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
3230      End Select
3240    End If

        ' ** Ctrl-Shift keys.
3250    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3260      Select Case intRetVal
          Case vbKeyF
3270        With Me
3280          intRetVal = 0
3290          DoCmd.SelectObject acForm, .Parent.Name, False
3300          .Parent.FocusHolder.SetFocus
3310        End With
3320      Case vbKeyT
3330        intRetVal = 0
3340        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
3350      Case vbKeyTab, vbKeyReturn
3360        With Me
3370          intRetVal = 0
3380          DoCmd.SelectObject acForm, .Parent.Name, False
3390          .Parent.chkHidden.SetFocus
3400        End With
3410      End Select
3420    End If

        ' ** Ctrl-Alt keys.
3430    If intCtrlDown And intAltDown And (Not intShiftDown) Then
3440      Select Case intRetVal
          Case vbKeyE
3450        With Me
3460          intRetVal = 0
3470          DoCmd.SelectObject acForm, .Parent.Name, False
3480          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
3490          frmCrit.chkRevcodeType_Expense.SetFocus
3500          frmCrit.chkRevcodeType_Expense = (Not frmCrit.chkRevcodeType_Expense)
3510          frmCrit.chkRevcodeType_Expense_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
3520        End With
3530      Case vbKeyI
3540        With Me
3550          intRetVal = 0
3560          DoCmd.SelectObject acForm, .Parent.Name, False
3570          .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
3580          frmCrit.chkRevcodeType_Income.SetFocus
3590          frmCrit.chkRevcodeType_Income = (Not frmCrit.chkRevcodeType_Income)
3600          frmCrit.chkRevcodeType_Income_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
3610        End With
3620      End Select
3630    End If

EXITP:
3640    Set ctl = Nothing
3650    KeyCode = intRetVal
3660    Exit Sub

ERRH:
3670    intRetVal = 0
3680    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3690    Case Else
3700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3710    End Select
3720    Resume EXITP

End Sub

Private Sub Form_Timer()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim frm As Access.Form
        Dim strControl As String
        Dim lngLoopCnt As Long
        Dim lngX As Long

3810    With Me
          ' ** Kluge!
3820      .TimerInterval = 0
3830      If strToFoc2 <> vbNullString Then
            ' ** journal_USER_GotFocus, posted_GotFocus, ledger_HIDDEN_GotFocus.
3840        strControl = Left(strToFoc2, (Len(strToFoc2) - Len("_GotFocus")))
3850        .Controls(strControl).SetFocus
3860        If strToFoc1 <> vbNullString Then
3870          .TimerInterval = 1000&
3880        End If
3890      End If
3900      If strToFoc1 <> vbNullString Then
3910        strControl = Left(strToFoc1, (Len(strToFoc1) - Len("_GotFocus")))
3920        lngLoopCnt = 0&
3930        Do Until Screen.ActiveControl.Name = "Journal Number"
3940  On Error Resume Next
3950          .[Journal Number].SetFocus
3960  On Error GoTo ERRH
3970          lngLoopCnt = lngLoopCnt + 1&
3980          If lngLoopCnt >= 100& Then Exit Do
3990        Loop
4000      End If
4010      If blnColWidth = True Then
4020        blnColWidth = False
            ' ** Adjust column widths according to preferences.
4030        Set frm = .Parent
4040        If frm.col00_width <> 0 Then  ' ** Default is 0, but hidden = -1.
4050          For lngX = 0& To (lngCols - 1&)
4060            arr_varCol(C_WDT, lngX) = frm.Controls("col" & Right("00" & CStr(lngX), 2) & "_width")
4070            If arr_varCol(C_WDT, lngX) = -1 Or arr_varCol(C_WDT, lngX) = 0 Then
4080              .Controls(arr_varCol(C_COLNAM, lngX)).ColumnHidden = True
4090            Else
4100              .Controls(arr_varCol(C_COLNAM, lngX)).ColumnWidth = arr_varCol(C_WDT, lngX)
4110            End If
4120          Next
4130        End If
4140      End If
4150    End With

EXITP:
4160    Set frm = Nothing
4170    Exit Sub

ERRH:
4180    strToFoc2 = vbNullString
4190    Me.Hidden.Enabled = True
4200    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
4210    Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        ' ** Pref_Save() in frmTransaction_Audit.

EXITP:
4310    Set frmCrit = Nothing
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Public Sub Journal_Number_GotFocus()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Journal_Number_GotFocus"

4410    With Me
          ' ** Kluge!
4420      strToFoc1 = vbNullString: strToFoc2 = vbNullString
4430      .[Journal Number].SelLength = 0
4440      .[Journal Number].SelStart = 9
4450    End With

EXITP:
4460    Exit Sub

ERRH:
4470    Select Case ERR.Number
        Case Else
4480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4490    End Select
4500    Resume EXITP

End Sub

Private Sub Journal_Number_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Journal_Number_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4620    intShiftDown = (Shift And acShiftMask) > 0
4630    intAltDown = (Shift And acAltMask) > 0
4640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4670        With Me
4680          If .[Journal Type].Visible = True And .[Journal Type].Enabled = True Then
4690            intRetVal = 0
4700            .[Journal Type].SetFocus
4710          End If
4720        End With
4730      End Select
4740    End If

        ' ** Shift keys.
4750    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4770        With Me
4780          intRetVal = 0
4790          If .CurrentRecord > 1 Then
4800            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
                '.[Date Posted].SetFocus
4810          Else
4820            DoCmd.SelectObject acForm, .Parent.Name, False
                '.Parent.chkHidden.SetFocus
4830          End If
4840        End With
4850      End Select
4860    End If

EXITP:
4870    KeyCode = intRetVal
4880    Exit Sub

ERRH:
4890    intRetVal = 0
4900    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
4910    Case Else
4920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4930    End Select
4940    Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

5010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5020    Exit Sub

ERRH:
5030    Select Case ERR.Number
        Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Sub

Private Sub transdate_chk_AfterUpdate()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_chk_AfterUpdate"

5110    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
5120    Exit Sub

ERRH:
5130    Select Case ERR.Number
        Case Else
5140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5150    End Select
5160    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

5210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5220    Exit Sub

ERRH:
5230    Select Case ERR.Number
        Case Else
5240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5250    End Select
5260    Resume EXITP

End Sub

Private Sub accountno_chk_AfterUpdate()

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_chk_AfterUpdate"

5310    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

5410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5420    Exit Sub

ERRH:
5430    Select Case ERR.Number
        Case Else
5440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5450    End Select
5460    Resume EXITP

End Sub

Private Sub shortname_chk_AfterUpdate()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_chk_AfterUpdate"

5510    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
5520    Exit Sub

ERRH:
5530    Select Case ERR.Number
        Case Else
5540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5550    End Select
5560    Resume EXITP

End Sub

Private Sub cusip_lbl_DblClick(Cancel As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_lbl_DblClick"

5610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5620    Exit Sub

ERRH:
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Private Sub cusip_chk_AfterUpdate()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_chk_AfterUpdate"

5710    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
5720    Exit Sub

ERRH:
5730    Select Case ERR.Number
        Case Else
5740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5750    End Select
5760    Resume EXITP

End Sub

Private Sub asset_description_lbl_DblClick(Cancel As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_lbl_DblClick"

5810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5820    Exit Sub

ERRH:
5830    Select Case ERR.Number
        Case Else
5840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5850    End Select
5860    Resume EXITP

End Sub

Private Sub asset_description_chk_AfterUpdate()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_chk_AfterUpdate"

5910    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
5920    Exit Sub

ERRH:
5930    Select Case ERR.Number
        Case Else
5940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5950    End Select
5960    Resume EXITP

End Sub

Private Sub shareface_lbl_DblClick(Cancel As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_lbl_DblClick"

6010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6020    Exit Sub

ERRH:
6030    Select Case ERR.Number
        Case Else
6040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6050    End Select
6060    Resume EXITP

End Sub

Private Sub shareface_chk_AfterUpdate()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_chk_AfterUpdate"

6110    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
6120    Exit Sub

ERRH:
6130    Select Case ERR.Number
        Case Else
6140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6150    End Select
6160    Resume EXITP

End Sub

Private Sub icash_lbl_DblClick(Cancel As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_lbl_DblClick"

6210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6220    Exit Sub

ERRH:
6230    Select Case ERR.Number
        Case Else
6240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6250    End Select
6260    Resume EXITP

End Sub

Private Sub icash_chk_AfterUpdate()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_chk_AfterUpdate"

6310    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Sub

Private Sub pcash_lbl_DblClick(Cancel As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_lbl_DblClick"

6410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6420    Exit Sub

ERRH:
6430    Select Case ERR.Number
        Case Else
6440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6450    End Select
6460    Resume EXITP

End Sub

Private Sub pcash_chk_AfterUpdate()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_chk_AfterUpdate"

6510    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
6520    Exit Sub

ERRH:
6530    Select Case ERR.Number
        Case Else
6540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6550    End Select
6560    Resume EXITP

End Sub

Private Sub cost_lbl_DblClick(Cancel As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_lbl_DblClick"

6610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6620    Exit Sub

ERRH:
6630    Select Case ERR.Number
        Case Else
6640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6650    End Select
6660    Resume EXITP

End Sub

Private Sub cost_chk_AfterUpdate()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_chk_AfterUpdate"

6710    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
6720    Exit Sub

ERRH:
6730    Select Case ERR.Number
        Case Else
6740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6750    End Select
6760    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

6810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6820    Exit Sub

ERRH:
6830    Select Case ERR.Number
        Case Else
6840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6850    End Select
6860    Resume EXITP

End Sub

Private Sub curr_id_chk_AfterUpdate()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_chk_AfterUpdate"

6910    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
6920    Exit Sub

ERRH:
6930    Select Case ERR.Number
        Case Else
6940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6950    End Select
6960    Resume EXITP

End Sub

Private Sub assetdate_lbl_DblClick(Cancel As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_lbl_DblClick"

7010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7020    Exit Sub

ERRH:
7030    Select Case ERR.Number
        Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Sub

Private Sub assetdate_chk_AfterUpdate()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_chk_AfterUpdate"

7110    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
7120    Exit Sub

ERRH:
7130    Select Case ERR.Number
        Case Else
7140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7150    End Select
7160    Resume EXITP

End Sub

Private Sub PurchaseDate_lbl_DblClick(Cancel As Integer)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDate_lbl_DblClick"

7210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7220    Exit Sub

ERRH:
7230    Select Case ERR.Number
        Case Else
7240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7250    End Select
7260    Resume EXITP

End Sub

Private Sub PurchaseDate_chk_AfterUpdate()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "PurchaseDate_chk_AfterUpdate"

7310    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
7320    Exit Sub

ERRH:
7330    Select Case ERR.Number
        Case Else
7340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7350    End Select
7360    Resume EXITP

End Sub

Private Sub ledger_description_lbl_DblClick(Cancel As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_lbl_DblClick"

7410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7420    Exit Sub

ERRH:
7430    Select Case ERR.Number
        Case Else
7440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7450    End Select
7460    Resume EXITP

End Sub

Private Sub ledger_description_chk_AfterUpdate()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_description_chk_AfterUpdate"

7510    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
7520    Exit Sub

ERRH:
7530    Select Case ERR.Number
        Case Else
7540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7550    End Select
7560    Resume EXITP

End Sub

Private Sub RecurringItem_lbl_DblClick(Cancel As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_lbl_DblClick"

7610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7620    Exit Sub

ERRH:
7630    Select Case ERR.Number
        Case Else
7640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7650    End Select
7660    Resume EXITP

End Sub

Private Sub RecurringItem_chk_AfterUpdate()

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_chk_AfterUpdate"

7710    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
7720    Exit Sub

ERRH:
7730    Select Case ERR.Number
        Case Else
7740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7750    End Select
7760    Resume EXITP

End Sub

Private Sub revcode_DESC_lbl_DblClick(Cancel As Integer)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_lbl_DblClick"

7810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7820    Exit Sub

ERRH:
7830    Select Case ERR.Number
        Case Else
7840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7850    End Select
7860    Resume EXITP

End Sub

Private Sub revcode_DESC_chk_AfterUpdate()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_chk_AfterUpdate"

7910    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
7920    Exit Sub

ERRH:
7930    Select Case ERR.Number
        Case Else
7940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7950    End Select
7960    Resume EXITP

End Sub

Private Sub revcode_TYPE_Description_lbl_DblClick(Cancel As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_Description_lbl_DblClick"

8010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8020    Exit Sub

ERRH:
8030    Select Case ERR.Number
        Case Else
8040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8050    End Select
8060    Resume EXITP

End Sub

Private Sub revcode_TYPE_Description_chk_AfterUpdate()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_Description_chk_AfterUpdate"

8110    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
8120    Exit Sub

ERRH:
8130    Select Case ERR.Number
        Case Else
8140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8150    End Select
8160    Resume EXITP

End Sub

Private Sub taxcode_description_lbl_DblClick(Cancel As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_lbl_DblClick"

8210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8220    Exit Sub

ERRH:
8230    Select Case ERR.Number
        Case Else
8240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8250    End Select
8260    Resume EXITP

End Sub

Private Sub taxcode_description_chk_AfterUpdate()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_chk_AfterUpdate"

8310    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
8320    Exit Sub

ERRH:
8330    Select Case ERR.Number
        Case Else
8340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8350    End Select
8360    Resume EXITP

End Sub

Private Sub taxcode_type_description_lbl_DblClick(Cancel As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_description_lbl_DblClick"

8410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8420    Exit Sub

ERRH:
8430    Select Case ERR.Number
        Case Else
8440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8450    End Select
8460    Resume EXITP

End Sub

Private Sub taxcode_type_description_chk_AfterUpdate()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_description_chk_AfterUpdate"

8510    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
8520    Exit Sub

ERRH:
8530    Select Case ERR.Number
        Case Else
8540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8550    End Select
8560    Resume EXITP

End Sub

Private Sub Location_Name_lbl_DblClick(Cancel As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "Location_Name_lbl_DblClick"

8610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8620    Exit Sub

ERRH:
8630    Select Case ERR.Number
        Case Else
8640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8650    End Select
8660    Resume EXITP

End Sub

Private Sub Location_Name_chk_AfterUpdate()

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "Location_Name_chk_AfterUpdate"

8710    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
8720    Exit Sub

ERRH:
8730    Select Case ERR.Number
        Case Else
8740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8750    End Select
8760    Resume EXITP

End Sub

Private Sub CheckNum_lbl_DblClick(Cancel As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_lbl_DblClick"

8810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8820    Exit Sub

ERRH:
8830    Select Case ERR.Number
        Case Else
8840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8850    End Select
8860    Resume EXITP

End Sub

Private Sub CheckNum_chk_AfterUpdate()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_chk_AfterUpdate"

8910    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
8920    Exit Sub

ERRH:
8930    Select Case ERR.Number
        Case Else
8940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8950    End Select
8960    Resume EXITP

End Sub

Private Sub journal_USER_lbl_DblClick(Cancel As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_lbl_DblClick"

9010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
9020    Exit Sub

ERRH:
9030    Select Case ERR.Number
        Case Else
9040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9050    End Select
9060    Resume EXITP

End Sub

Private Sub journal_USER_chk_AfterUpdate()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_chk_AfterUpdate"

9110    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
9120    Exit Sub

ERRH:
9130    Select Case ERR.Number
        Case Else
9140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9150    End Select
9160    Resume EXITP

End Sub

Private Sub User_GotFocus()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "User_GotFocus"

        Dim lngLoopCnt As Long

9210    With Me
          ' ** Kluge!
9220      If strToFoc2 = vbNullString Then
9230        strToFoc2 = THIS_PROC
9240        lngLoopCnt = 0&
9250        Do Until Screen.ActiveControl.Name = "User"
9260  On Error Resume Next
9270          .User.SetFocus
9280  On Error GoTo ERRH
9290          lngLoopCnt = lngLoopCnt + 1&
9300          If lngLoopCnt >= 100& Then Exit Do
9310        Loop
9320      Else
9330        strToFoc2 = vbNullString
9340      End If
9350    End With

EXITP:
9360    Exit Sub

ERRH:
9370    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
9380    Case Else
9390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9400    End Select
9410    Resume EXITP

End Sub

Private Sub posted_lbl_DblClick(Cancel As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "posted_lbl_DblClick"

9510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
9520    Exit Sub

ERRH:
9530    Select Case ERR.Number
        Case Else
9540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9550    End Select
9560    Resume EXITP

End Sub

Private Sub posted_chk_AfterUpdate()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "posted_chk_AfterUpdate"

9610    Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
9620    Exit Sub

ERRH:
9630    Select Case ERR.Number
        Case Else
9640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9650    End Select
9660    Resume EXITP

End Sub

Private Sub Date_Posted_GotFocus()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_Posted_GotFocus"

        Dim lngLoopCnt As Long

9710    With Me
          ' ** Kluge!
9720      If strToFoc2 = vbNullString Then
9730        strToFoc2 = THIS_PROC
9740        lngLoopCnt = 0&
9750        Do Until Screen.ActiveControl.Name = "Date Posted"
9760  On Error Resume Next
9770          .[Date Posted].SetFocus
9780  On Error GoTo ERRH
9790          lngLoopCnt = lngLoopCnt + 1&
9800          If lngLoopCnt >= 100& Then Exit Do
9810        Loop
9820      Else
9830        strToFoc2 = vbNullString
9840      End If
9850    End With

EXITP:
9860    Exit Sub

ERRH:
9870    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
9880    Case Else
9890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9900    End Select
9910    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_lbl_DblClick(Cancel As Integer)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_lbl_DblClick"

10010   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
10020   Exit Sub

ERRH:
10030   Select Case ERR.Number
        Case Else
10040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10050   End Select
10060   Resume EXITP

End Sub

Private Sub ledger_HIDDEN_chk_AfterUpdate()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_chk_AfterUpdate"

10110   Print_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
10120   Exit Sub

ERRH:
10130   Select Case ERR.Number
        Case Else
10140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10150   End Select
10160   Resume EXITP

End Sub

Private Sub Hidden_GotFocus()

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "Hidden_GotFocus"

        Dim lngLoopCnt As Long

10210   With Me
          ' ** Kluge!
10220     If strToFoc2 = vbNullString Then
10230       strToFoc2 = THIS_PROC
10240       lngLoopCnt = 0&
10250       Do Until Screen.ActiveControl.Name = "Hidden"
10260 On Error Resume Next
10270         .Hidden.SetFocus
10280 On Error GoTo ERRH
10290         lngLoopCnt = lngLoopCnt + 1&
10300         If lngLoopCnt >= 100& Then Exit Do
10310       Loop
10320     Else
10330       strToFoc2 = vbNullString
10340     End If
10350   End With

EXITP:
10360   Exit Sub

ERRH:
10370   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
10380   Case Else
10390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10400   End Select
10410   Resume EXITP

End Sub

Private Sub Hidden_KeyDown(KeyCode As Integer, Shift As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "Hidden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10520   intShiftDown = (Shift And acShiftMask) > 0
10530   intAltDown = (Shift And acAltMask) > 0
10540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10570       With Me
10580         intRetVal = 0
10590         lngRecsCur = RecCnt  ' ** Function: Below.
10600         If .CurrentRecord < lngRecsCur Then
10610           MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
10620           .[Journal Number].SetFocus
                ' ** It keeps wanting to come back here! WHY?
                '.Hidden.Enabled = False
10630         Else
10640           DoCmd.SelectObject acForm, .Parent.Name, False
10650           .Parent.cmdSelectAll.SetFocus
10660         End If
10670       End With
10680     End Select
10690   End If

        ' ** Shift keys.
10700   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10710     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10720       With Me
              'intRetVal = 0
              '.[Date Posted].SetFocus
10730       End With
10740     End Select
10750   End If

EXITP:
10760   KeyCode = intRetVal
10770   Exit Sub

ERRH:
10780   intRetVal = 0
10790   Select Case ERR.Number
        Case Else
10800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10810   End Select
10820   Resume EXITP

End Sub

Public Function FilterRecs_Cnt(rstAll As DAO.Recordset) As Boolean

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Cnt"

        Dim rst As DAO.Recordset
        Dim ctl As Access.Control
        Dim blnRetVal As Boolean

10910   blnRetVal = True

10920   If strFilter <> vbNullString Then
10930     With Me
10940       DoCmd.Hourglass True
10950       DoEvents
10960       rstAll.Filter = strFilter
10970       .Parent.CurrentFilter2 = strFilter
10980 On Error Resume Next
10990       Set rst = rstAll.OpenRecordset
11000       If ERR.Number <> 0 Then
11010 On Error GoTo ERRH
11020         Set rst = Nothing
11030         strFilter = SwapNames(strFilter, True)  ' ** Function: Below.
11040         rstAll.Filter = strFilter
11050         Set rst = rstAll.OpenRecordset
11060       Else
11070 On Error GoTo ERRH
11080       End If
11090       If rst.BOF = True And rst.EOF = True Then
11100         blnRetVal = False
11110       Else
11120         rst.MoveLast
11130         dblFilterRecs = rst.RecordCount
11140         .Parent.FilterRecs_Set strFilter, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit.
11150       End If
11160       rst.Close
11170       rstAll.Filter = vbNullString
11180       .Parent.CurrentFilter2 = vbNullString
11190       If blnRetVal = False Then
11200         .Parent.FilterRecs = "None"
11210         .Parent.FilterRecs.ForeColor = CLR_DKRED
11220         .Parent.FilterRecs_lbl.ForeColor = CLR_DKRED
11230         For Each ctl In .Section("Detail").Controls
11240           With ctl
11250             If .Visible = True Then
11260               If .ControlType = acTextBox Then
11270                 .BackColor = CLR_GRY2
11280               End If
11290             End If
11300           End With
11310         Next
11320       Else
11330         .Filter = SwapNames(strFilter)  ' ** Function: Below.
11340         .Parent.CurrentFilter2 = strFilter
11350         .FilterOn = True
11360         .Parent.FilterRecs = Format(dblFilterRecs, "#,##0")
11370         .Parent.FilterRecs.ForeColor = CLR_BLU
11380         .Parent.FilterRecs_lbl.ForeColor = CLR_BLU
11390         For Each ctl In .Section("Detail").Controls
11400           With ctl
11410             If .Visible = True Then
11420               If .ControlType = acTextBox Then
11430                 .BackColor = CLR_WHT
11440               End If
11450             End If
11460           End With
11470         Next
11480       End If
11490       DoCmd.Hourglass False
11500     End With
11510   Else
11520     With Me
11530       If .Filter <> vbNullString Then
11540         .Filter = vbNullString
11550         .Parent.CurrentFilter2 = vbNullString
11560       End If
11570       If .FilterOn = True Then
11580         .FilterOn = False
11590       End If
11600       .Parent.FilterRecs = "All"
11610       .Parent.FilterRecs.ForeColor = CLR_DISABLED_FG
11620       .Parent.FilterRecs.BackColor = CLR_DISABLED_BG
11630       .Parent.FilterRecs_lbl.ForeColor = CLR_VDKGRY
11640       For Each ctl In .Section("Detail").Controls
11650         With ctl
11660           If .Visible = True Then
11670             If .ControlType = acTextBox Then
11680               .BackColor = CLR_WHT
11690             End If
11700           End If
11710         End With
11720       Next
11730     End With
11740   End If

EXITP:
11750   Set ctl = Nothing
11760   Set rst = Nothing
11770   FilterRecs_Cnt = blnRetVal
11780   Exit Function

ERRH:
11790   DoCmd.Hourglass False
11800   Select Case ERR.Number
        Case Else
11810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11820   End Select
11830   Resume EXITP

End Function

Public Sub FilterRecs_Rem(strClause As String)
' ** Remove a clause from the filter string.

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Rem"

        Dim blnIsJType As Boolean, blnIsCheckNum As Boolean, blnIsHidden As Boolean
        Dim lngMultiCnt As Long
        Dim intPos03 As Integer, intPos04 As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const CHK_NUM      As String = "[CheckNum] = "
        Const CHK_NUM1     As String = "[CheckNum] >= "
        Const CHK_NUM2     As String = "[CheckNum] <= "
        Const HIDDEN_TRX1  As String = "[ledger_HIDDEN] = True"
        Const HIDDEN_TRX2  As String = "[ledger_HIDDEN] = False"

11910   If strClause <> vbNullString And strFilter <> vbNullString Then
11920     If InStr(strClause, "[journaltype]") > 0 Or InStr(strClause, "[Journal Type]") > 0 Then blnIsJType = True Else blnIsJType = False
11930     If InStr(strClause, "[CheckNum]") > 0 Or InStr(strClause, "[Check Number]") > 0 Then blnIsCheckNum = True Else blnIsCheckNum = False
11940     If InStr(strClause, "[ledger_HIDDEN]") > 0 Or InStr(strClause, "[Hidden]") > 0 Then blnIsHidden = True Else blnIsHidden = False
11950     With Me
11960       intPos01 = InStr(strFilter, strClause)
11970       If blnIsCheckNum = True And intPos01 = 0 Then
11980         intPos01 = InStr(strFilter, StringReplace(strClause, "CheckNum", "Check Number"))   ' ** Module Function: modStringFuncs.
11990         If intPos01 = 0 Then
12000           intPos01 = InStr(strFilter, CHK_NUM1)
12010           If intPos01 = 0 Then
12020             intPos01 = InStr(strFilter, StringReplace(CHK_NUM1, "CheckNum", "Check Number"))
12030             If intPos01 = 0 Then
12040               intPos01 = InStr(strFilter, CHK_NUM2)
12050               If intPos01 = 0 Then
12060                 intPos01 = InStr(strFilter, StringReplace(CHK_NUM2, "CheckNum", "Check Number"))
12070               End If
12080             End If
12090           End If
12100         End If
12110       End If
12120       If blnIsHidden = True And intPos01 = 0 Then
12130         intPos01 = InStr(strFilter, StringReplace(HIDDEN_TRX1, "ledger_HIDDEN", "Hidden"))
12140         If intPos01 = 0 Then
12150           intPos01 = InStr(strFilter, HIDDEN_TRX2)
12160           If intPos01 = 0 Then
12170             intPos01 = InStr(strFilter, StringReplace(HIDDEN_TRX2, "ledger_HIDDEN", "Hidden"))
12180           End If
12190         End If
12200       End If
12210       If intPos01 > 0 Then
              ' ** This clause is present
12220         Select Case blnIsCheckNum
              Case True
12230           intPos01 = InStr(strFilter, CHK_NUM)
12240           intPos02 = InStr(strFilter, CHK_NUM1)
12250           intPos03 = InStr(strFilter, CHK_NUM2)
12260           If intPos01 = 0 And intPos02 = 0 And intPos03 = 0 Then
12270             intPos01 = InStr(strFilter, StringReplace(CHK_NUM, "CheckNum", "Check Number"))
12280             intPos02 = InStr(strFilter, StringReplace(CHK_NUM1, "CheckNum", "Check Number"))
12290             intPos03 = InStr(strFilter, StringReplace(CHK_NUM2, "CheckNum", "Check Number"))
12300           End If
12310           strTmp01 = vbNullString: strTmp02 = vbNullString
12320           If intPos01 > 0 Then
                  ' ** Simple equaltiy.
12330             strTmp01 = Left(strFilter, (intPos01 - 1))
12340             intPos04 = InStr((intPos01 + Len(CHK_NUM) + 1), strFilter, ANDF)
12350             If intPos04 > 0 Then
12360               strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
12370             End If
12380           ElseIf intPos02 > 0 And intPos03 > 0 Then
                  ' ** A range of check numbers.
12390             strTmp01 = Left(strFilter, (intPos02 - 1))
12400             If Right(Trim(strTmp01), 1) = "(" Then strTmp01 = Left(Trim(strTmp01), (Len(Trim(strTmp01)) - 1))
12410             intPos04 = InStr((intPos03 + Len(CHK_NUM2) + 1), strFilter, ANDF)  ' ** This will miss closing paren.
12420             If intPos04 > 0 Then
12430               strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
12440             End If
12450           ElseIf intPos02 > 0 Then
                  ' ** Greater than.
12460             strTmp01 = Left(strFilter, (intPos02 - 1))
12470             intPos04 = InStr((intPos02 + Len(CHK_NUM1) + 1), strFilter, ANDF)
12480             If intPos04 > 0 Then
12490               strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
12500             End If
12510           ElseIf intPos03 > 0 Then
                  ' ** Less than.
12520             strTmp01 = Left(strFilter, (intPos03 - 1))
12530             intPos04 = InStr((intPos03 + Len(CHK_NUM2) + 1), strFilter, ANDF)
12540             If intPos04 > 0 Then
12550               strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
12560             End If
12570           End If
12580           If Trim(strTmp01) <> vbNullString Then
12590             If Right(Trim(strTmp01), 4) = " And" Then
                    ' ** Remove this clause's preceeding 'And'.
12600               strTmp01 = Trim(Left(Trim(strTmp01), (Len(Trim(strTmp01)) - 3)))
12610             End If
12620           Else
                  ' ** Clause is beginning of line.
12630             strTmp01 = Trim(strTmp01)
12640           End If
12650           strFilter = strTmp01 & strTmp02  ' ** Both may be vbNullString.
12660         Case False
12670           Select Case blnIsHidden
                Case True
12680             intPos01 = InStr(strFilter, HIDDEN_TRX1)
12690             intPos02 = InStr(strFilter, HIDDEN_TRX2)
12700             If intPos01 = 0 And intPos02 = 0 Then
12710               intPos01 = InStr(strFilter, StringReplace(HIDDEN_TRX1, "ledger_HIDDEN", "Hidden"))
12720               intPos02 = InStr(strFilter, StringReplace(HIDDEN_TRX2, "ledger_HIDDEN", "Hidden"))
12730             End If
12740             strTmp01 = vbNullString: strTmp02 = vbNullString
12750             If intPos01 > 0 Then
                    ' ** Only hidden transactions.
12760               strTmp01 = Left(strFilter, (intPos01 - 1))
12770               intPos04 = InStr(intPos01, strFilter, ANDF)
12780               If intPos04 > 0 Then
12790                 strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
12800               End If
12810             ElseIf intPos02 > 0 Then
                    ' ** Excludes hidden transactions.
12820               strTmp01 = Left(strFilter, (intPos02 - 1))
12830               intPos04 = InStr(intPos02, strFilter, ANDF)
12840               If intPos04 > 0 Then
12850                 strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
12860               End If
12870             End If
12880             If Trim(strTmp01) <> vbNullString Then
12890               If Right(Trim(strTmp01), 4) = " And" Then
                      ' ** Remove this clause's preceeding 'And'.
12900                 strTmp01 = Trim(Left(Trim(strTmp01), (Len(Trim(strTmp01)) - 3)))
12910               End If
12920             Else
                    ' ** Clause is beginning of line.
12930               strTmp01 = Trim(strTmp01)
12940             End If
12950             strFilter = strTmp01 & strTmp02  ' ** Both may be vbNullString.
12960           Case False
12970             Select Case blnIsJType
                  Case True
12980               lngMultiCnt = CharCnt(strFilter, strClause, True)  ' ** Module Function: modStringFuncs.
12990               Select Case lngMultiCnt
                    Case 1&  ' ** For 1 occurance only.
13000                 intPos02 = InStr(intPos01, strFilter, ")" & ANDF)  ' ** Searches for following ') And '.
13010               Case 2&, 3&  ' ** For 1st of 2 or 3 occurances.
13020                 intPos02 = InStr(intPos01, strFilter, ORF)  ' ** Searches for following ' Or '.
13030               End Select
13040             Case False
13050               intPos02 = InStr(intPos01, strFilter, ANDF)  ' ** Searches for following ' And '.
13060             End Select
13070             If intPos01 = 1 Or (blnIsJType = True And intPos01 = 2) Then
                    ' ** This clause starts the filter.
13080               If intPos02 = 0 Then
                      ' ** This clause is the only clause.
13090                 strFilter = vbNullString
13100               Else
                      ' ** There's another clause after this one.
13110                 Select Case blnIsJType
                      Case True
13120                   Select Case lngMultiCnt
                        Case 1&  ' ** For 1 occurance only.
13130                     strFilter = Mid(strFilter, (intPos02 + Len(")" & ANDF)))  ' ** Pick up after ') And '.
13140                   Case 2&, 3&  ' ** For 1st of 2 or 3 occurances.
13150                     strFilter = Mid(strFilter, (intPos02 + Len(ORF)))  ' ** Pick up after ' Or '.
13160                   End Select
13170                 Case False
13180                   strFilter = Mid(strFilter, (intPos02 + Len(ANDF)))  ' ** Pick up after ' And '.
13190                 End Select
13200               End If
13210             Else
                    ' ** There's another clause before this one.
13220               If intPos02 = 0 Then
                      ' ** This clause is the last one.
13230                 Select Case blnIsJType
                      Case True  ' ** For 1 occurance only (otherwise it wouldn't be the last one).
13240                   strFilter = Left(strFilter, (intPos01 - Len(ANDF & "(")))  ' ** Remove final ' And ('.
13250                 Case False
13260                   strFilter = Left(strFilter, (intPos01 - Len(ANDF)))  ' ** Remove final ' And '.
13270                 End Select
13280               Else
                      ' ** There's a clause both before and after this one.
13290                 Select Case blnIsJType
                      Case True
13300                   Select Case lngMultiCnt
                        Case 1&  ' ** For 1 occurance only.
13310                     strFilter = Left(strFilter, (intPos01 - 2)) & Mid(strFilter, (intPos02 + Len(")" & ANDF)))
13320                   Case 2&, 3&  ' ** For 1st of 2 or 3 occurances.
13330                     strFilter = Left(strFilter, (intPos01 - 1)) & Mid(strFilter, (intPos02 + Len(ORF)))
13340                   End Select
13350                 Case False
13360                   strFilter = Left(strFilter, (intPos01 - 1)) & Mid(strFilter, (intPos02 + Len(ANDF)))  ' ** Don't leave 2 spaces or 2 And's.
13370                 End Select
13380               End If
13390             End If
13400           End Select
13410         End Select
13420         .Filter = SwapNames(strFilter)  ' ** Function: Below.
13430         DoEvents
13440         If IsNothing(frmCrit) = True Then  ' ** Module Function: modUtilities.
13450           Set frmCrit = .Parent.frmTransaction_Audit_Sub_Criteria.Form
13460         End If
13470         .Parent.CurrentFilter2 = strFilter
13480         If strFilter = vbNullString Then
13490           .FilterOn = False
13500         Else
13510           .FilterOn = True
13520         End If
13530         dblFilterRecs = RecCnt  ' ** Function: Below.
13540         .Parent.FilterRecs_Set strFilter, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit.
13550         frmCrit.FilterRecs_Set strFilter, dblFilterRecs, 2  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
13560       End If
13570     End With
13580   End If

EXITP:
13590   Exit Sub

ERRH:
13600   Select Case ERR.Number
        Case Else
13610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13620   End Select
13630   Resume EXITP

End Sub

Public Sub FilterRecs_Clr(strProc As String, blnOn As Boolean)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Clr"

        Dim strCtl As String, strCtlLabel As String, strCtl2 As String, strCtlLabel2 As String, strCtlLabel3 As String
        Dim strFld As String, strFldLabelA As String, strFldLabelB As String
        Dim strFld2 As String, strFldLabel2 As String, strFld3 As String, strFldLabel3 As String
        Dim intPos01 As Integer
        Dim lngX As Long

13710   strCtl = Left(strProc, (Len(strProc) - Len("_AfterUpdate")))  ' ** PostedDateStart_AfterUpdate -> PostedDateStart
13720   strCtlLabel = strCtl & "_lbl"                                  ' ** PostedDateStart -> PostedDateStart_lbl
13730   strFld = vbNullString: strFldLabelA = vbNullString: strFldLabelB = vbNullString
13740   strFld2 = vbNullString: strFldLabel2 = vbNullString: strFld3 = vbNullString: strFldLabel3 = vbNullString
13750   strCtl2 = vbNullString: strCtlLabel2 = vbNullString: strCtlLabel3 = vbNullString

13760   arr_varFilt_ds = Empty
13770   arr_varFilt_ds = FilterRecs_GetArr(2)  ' ** Module Function: modTransactionAuditFuncs1.
13780   If IsEmpty(arr_varFilt_ds) = True Then
13790     FilterRecs_Load  ' ** Module Function: modTransactionAuditFuncs1.
13800     DoEvents
13810     arr_varFilt_ds = FilterRecs_GetArr(2)  ' ** Module Function: modTransactionAuditFuncs1.
13820   End If
13830   If IsEmpty(arr_varFilt_ds) = False Then

13840     lngFilts_ds = (UBound(arr_varFilt_ds, 2) + 1&)

          ' ******************************************************
          ' ** Array: arr_varFilt_ds()
          ' **
          ' **   Field  Element  Name                 Constant
          ' **   =====  =======  ===================  ==========
          ' **     1       0     taf_index            F_NAM
          ' **     2       1     vbdec_name1          F_NAM
          ' **     3       2     vbdec_value1         F_CONST
          ' **     4       3     ctl_name1            F_CTL
          ' **     5       4     ctl_name_lbl1        F_CLBL
          ' **     6       5     fld_ctl_name1        F_FLD
          ' **     7       6     fld_ctl_name_lbl1    F_FLBL
          ' **     8       7     ctl_name2            F_CTL2
          ' **     9       8     ctl_name_lbl2        F_CLBL2
          ' **    10       9     fld_ctl_name2        F_FLD2
          ' **    11      10     fld_ctl_name_lbl2    F_FLBL2
          ' **    12      11     ctl_name_lbl3        F_CLBL3
          ' **    13      12     fld_ctl_name3        F_FLD3
          ' **    14      13     fld_ctl_name_lbl3    F_FLBL3
          ' **
          ' ******************************************************

13850     For lngX = 0& To (lngFilts_ds - 1&)
13860       If (arr_varFilt_ds(F_CTL, lngX) = strCtl) Or (arr_varFilt_ds(F_CTL, lngX) = "cmbJournalType1" And _
                (strCtl = "cmbJournalType2" Or strCtl = "cmbJournalType3")) Then
13870         strFld = arr_varFilt_ds(F_FLD, lngX)
13880         strFldLabelA = arr_varFilt_ds(F_FLBL, lngX)  ' ** Not displayed: [revcode_ID], [assetno].
13890         intPos01 = InStr(strFldLabelA, "~")
13900         If intPos01 > 0 Then
13910           strFldLabelB = Mid(strFldLabelA, (intPos01 + 1))
13920           strFldLabelA = Left(strFldLabelA, (intPos01 - 1))
13930         End If
13940         If arr_varFilt_ds(F_CTL2, lngX) <> vbNullString Then
                ' ** Two criteria controls affect one field.
                ' ** For example: {TransDateStart, TransDateEnd}, {AssetDateStart, AssetDateEnd}, {PostedDateStart, PostedDateEnd}
13950           strCtl2 = arr_varFilt_ds(F_CTL2, lngX)
13960           If arr_varFilt_ds(F_CLBL2, lngX) <> vbNullString Then
13970             strCtlLabel2 = arr_varFilt_ds(F_CLBL2, lngX)
13980           End If
13990           If arr_varFilt_ds(F_CLBL3, lngX) <> vbNullString Then
14000             strCtlLabel3 = arr_varFilt_ds(F_CLBL3, lngX)
14010           End If
14020         ElseIf arr_varFilt_ds(F_CLBL2, lngX) <> vbNullString Then
14030           strCtlLabel2 = arr_varFilt_ds(F_CLBL2, lngX)
14040         End If
14050         If arr_varFilt_ds(F_FLD2, lngX) <> vbNullString Then
                ' ** Two fields affected by one criteria control.
                ' ** For example: {accountno, shortname}, {assetno, asset_description}, {assetdate, PurchaseDate}, {revcode_ID, revcode_DESC}
14060           strFld2 = arr_varFilt_ds(F_FLD2, lngX)
14070           If arr_varFilt_ds(F_FLBL2, lngX) <> vbNullString Then
14080             strFldLabel2 = arr_varFilt_ds(F_FLBL2, lngX)
14090           End If
14100           If arr_varFilt_ds(F_FLD3, lngX) <> vbNullString Then
                  ' ** For example: {assetno, asset_description, cusip}
14110             strFld3 = arr_varFilt_ds(F_FLD3, lngX)
14120             If arr_varFilt_ds(F_FLBL3, lngX) <> vbNullString Then
14130               strFldLabel3 = arr_varFilt_ds(F_FLBL3, lngX)
14140             End If
14150           End If
14160         End If
14170         Exit For
14180       End If
14190     Next

14200     If strFld <> vbNullString Then
14210       With Me
14220         If IsNothing(frmCrit) = True Then  ' ** Module Function: modUtilities.
14230           Set frmCrit = .Parent.frmTransaction_Audit_Sub_Criteria.Form
14240         End If
14250         Select Case blnOn
              Case True
                ' ** Set parent label colors.
14260           frmCrit.Controls(strCtlLabel).ForeColor = CLR_BLU
14270           If strCtlLabel2 <> vbNullString Then
14280             frmCrit.Controls(strCtlLabel2).ForeColor = CLR_BLU
14290           End If
14300           If strCtlLabel3 <> vbNullString Then
14310             frmCrit.Controls(strCtlLabel3).ForeColor = CLR_BLU
14320           End If
                ' ** Set field label colors.
                'If strFldLabelA <> vbNullString Then
                '  .Controls(strFldLabelA).ForeColor = CLR_BLU
                '  .Controls(strFldLabelA).Tag = "Filter"
                'End If
                'If strFldLabelB <> vbNullString Then
                '  .Controls(strFldLabelB).ForeColor = CLR_BLU
                '  .Controls(strFldLabelB).Tag = "Filter"
                'End If
                'If strFldLabel2 <> vbNullString Then
                '  .Controls(strFldLabel2).ForeColor = CLR_BLU
                '  .Controls(strFldLabel2).Tag = "Filter"
                'End If
                'If strFldLabel3 <> vbNullString Then
                '  .Controls(strFldLabel3).ForeColor = CLR_BLU
                '  .Controls(strFldLabel3).Tag = "Filter"
                'End If
14330         Case False
                ' ** Set parent label colors.
14340           frmCrit.Controls(strCtlLabel).ForeColor = CLR_VDKGRY
14350           If strCtlLabel2 <> vbNullString Then
14360             frmCrit.Controls(strCtlLabel2).ForeColor = CLR_VDKGRY
14370           End If
14380           If strCtlLabel3 <> vbNullString Then
14390             frmCrit.Controls(strCtlLabel3).ForeColor = CLR_VDKGRY
14400           End If
                ' ** Set field label colors.
14410           If strCtl <> "cmbJournaltype2" And strCtl <> "cmbJournaltype3" Then
                  'If strFldLabelA <> vbNullString Then
                  '  .Controls(strFldLabelA).ForeColor = CLR_DKGRY2
                  '  .Controls(strFldLabelA).Tag = vbNullString
                  'End If
                  'If strFldLabelB <> vbNullString Then
                  '  .Controls(strFldLabelB).ForeColor = CLR_DKGRY2
                  '  .Controls(strFldLabelB).Tag = vbNullString
                  'End If
                  'If strFldLabel2 <> vbNullString Then
                  '  .Controls(strFldLabel2).ForeColor = CLR_DKGRY2
                  '  .Controls(strFldLabel2).Tag = vbNullString
                  'End If
                  'If strFldLabel3 <> vbNullString Then
                  '  .Controls(strFldLabel3).ForeColor = CLR_DKGRY2
                  '  .Controls(strFldLabel3).Tag = vbNullString
                  'End If
14420           End If
14430         End Select
14440       End With
14450     End If

14460   End If

EXITP:
14470   Set frmCrit = Nothing
14480   Exit Sub

ERRH:
14490   Select Case ERR.Number
        Case Else
14500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14510   End Select
14520   Resume EXITP

End Sub

Public Sub Print_Chk(strProc As String)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "Print_Chk"

14610   Print_Chk_TA strProc, 2, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
14620   Exit Sub

ERRH:
14630   Select Case ERR.Number
        Case Else
14640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14650   End Select
14660   Resume EXITP

End Sub

Public Function Print_ChkCnt() As Long

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "Print_ChkCnt"

        Dim lngRetVal As Long

14710   lngRetVal = Print_ChkCnt_TA(2, Me)   ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
14720   Print_ChkCnt = lngRetVal
14730   Exit Function

ERRH:
14740   lngRetVal = 0&
14750   Select Case ERR.Number
        Case Else
14760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14770   End Select
14780   Resume EXITP

End Function

Public Function RecCnt() As Long

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

14810   lngRetVal = 0&

14820   With Me
14830     Set rst = .RecordsetClone
14840     With rst
14850       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
14860       Else
14870         .MoveLast
14880         lngRetVal = .RecordCount
14890       End If
14900       .Close
14910     End With
14920   End With

EXITP:
14930   Set rst = Nothing
14940   RecCnt = lngRetVal
14950   Exit Function

ERRH:
14960   lngRetVal = 0&
14970   Select Case ERR.Number
        Case Else
14980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14990   End Select
15000   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

15110   Select Case intWhere
        Case acCmdRecordsGoToFirst
15120     DoCmd.RunCommand acCmdRecordsGoToFirst
15130   Case acCmdRecordsGoToLast
15140     DoCmd.RunCommand acCmdRecordsGoToLast
15150   Case acCmdRecordsGoToPrevious
15160     DoCmd.RunCommand acCmdRecordsGoToPrevious
15170   Case acCmdRecordsGoToNext
15180     DoCmd.RunCommand acCmdRecordsGoToNext
15190   Case Else
15200     If IsMissing(varID) = False Then
15210       With Me
15220         Set rst = .RecordsetClone
15230         With rst
15240           .FindFirst "[journalno] = " & CStr(varID)
15250           If .NoMatch = False Then
15260             Me.Bookmark = .Bookmark
15270           End If
15280           .Close
15290         End With
15300       End With
15310     End If
15320   End Select

EXITP:
15330   Set rst = Nothing
15340   Exit Sub

ERRH:
15350   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
15360   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
15370   Case Else
15380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15390   End Select
15400   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim ctl As Access.Control
        Dim strCalled As String, strSortAsc As String
        Dim blnFormView As Boolean
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [alphasort], [journaltype], [journalno]"

15510   With Me
15520     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
15530       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
15540     End If
15550     Select Case .CurrentView
          Case acCurViewFormBrowse
15560       blnFormView = True
15570     Case acCurViewDatasheet
15580       blnFormView = False
15590     End Select
15600     .Controls(strSortLbl).Visible = False
15610     .Controls(strSortLine).Visible = False
15620     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push of the right side of the form.
15630     If strProc = "Form_Load" Then
15640       strCalled = "Journal Number"
15650       strSortNow = strSortOrig
15660     Else
15670       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
15680       If strCalled = strCalled = "journalno" Or "Journal Number" Or strCalled = "Journal_Number" Then
15690         strSortAsc = "[journalno] DESC"
15700         If strSortNow = strSortOrig Then
15710           strSortNow = strSortAsc
15720         Else
15730           strSortNow = strSortOrig
15740         End If
15750       Else
15760         Select Case strCalled
              Case "journaltype", "Journal Type", "Journal_Type"
15770           strSortAsc = "[journaltype]" & strStdAsc
15780         Case "transdate", "Posting Date", "Posting_Date"
15790           strSortAsc = "[transdate]" & strStdAsc
15800         Case "accountno", "Account Number", "Account_Number"
15810           strSortAsc = "[alphasort], [journaltype], [journalno]"
15820         Case "shortname", "Name"
15830           strSortAsc = "[shortname]" & strStdAsc
15840         Case "cusip", "CUSIP"
15850           strSortAsc = "[cusip]" & strStdAsc
15860         Case "asset_description", "Asset"
15870           strSortAsc = "[asset_description]" & strStdAsc
15880         Case "shareface", "Share / Face", "Share_/_Face"
15890           strSortAsc = "[shareface]" & strStdAsc
15900         Case "icash", "Income Cash", "Income_Cash"
15910           strSortAsc = "[icash]" & strStdAsc
15920         Case "pcash", "Principal Cash", "Principal_Cash"
15930           strSortAsc = "[pcash]" & strStdAsc
15940         Case "cost", "Cost"
15950           strSortAsc = "[cost]" & strStdAsc
15960         Case "curr_id]"
15970           strSortAsc = "[curr_code]" & strStdAsc
15980         Case "assetdate", "Trade Date", "Trade_Date"
15990           strSortAsc = "[assetdate]" & strStdAsc
16000         Case "PurchaseDate", "Original Trade Date", "Original_Trade_Date"
16010           strSortAsc = "[PurchaseDate]" & strStdAsc
16020         Case "ledger_description", "Comments"
16030           strSortAsc = "[ledger_description]" & strStdAsc
16040         Case "RecurringItem", "Recurring Item", "Recurring_Item"
16050           strSortAsc = "[RecurringItem]" & strStdAsc
16060         Case "revcode_DESC", "Inc/Exp Codes", "Inc/Exp_Codes"
16070           strSortAsc = "[revcode_DESC]" & strStdAsc
16080         Case "revcode_TYPE_Description", "Inc/Exp"
16090           strSortAsc = "[revcode_TYPE_Description], [revcode_DESC]" & strStdAsc
16100         Case "taxcode_description", "Tax Codes", "Tax_Codes"
16110           strSortAsc = "[taxcode_description]" & strStdAsc
16120         Case "taxcode_type_description", "Inc/Ded"
16130           strSortAsc = "[taxcode_type_description], [taxcode_description]" & strStdAsc
16140         Case "Location_Name]"
16150           strSortAsc = "[Location_Name]" & strStdAsc
16160         Case "CheckNum"
16170           strSortAsc = "[CheckNum]" & strStdAsc
16180         Case "journal_USER", "User"
16190           strSortAsc = "[journal_USER]" & strStdAsc
16200         Case "posted", "Date Posted", "Date_Posted"
16210           strSortAsc = "[posted]" & strStdAsc
16220         Case "ledger_HIDDEN", "Hidden"
16230           strSortAsc = "[ledger_HIDDEN]" & strStdAsc
16240         End Select
16250         If strSortNow = strSortAsc Then
16260           intCnt = (CharCnt(strSortAsc, ",") + 1)  ' ** Module Function: modStringFuncs.
16270           Select Case intCnt
                Case 1
16280             strSortNow = strSortAsc & " DESC"
16290           Case 2
16300             intPos01 = InStr(strSortAsc, ",")
16310             strSortNow = Left(strSortAsc, (intPos01 - 1)) & " DESC" & Mid(strSortAsc, intPos01) & " DESC"
16320           Case 3
16330             intPos01 = InStr(strSortAsc, ",")
16340             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
16350             strTmp02 = Mid(strSortAsc, intPos01)
16360             intPos01 = InStr(2, strTmp02, ",")
16370             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16380             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
16390             strSortNow = strTmp01
16400           Case 4
16410             intPos01 = InStr(strSortAsc, ",")
16420             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
16430             strTmp02 = Mid(strSortAsc, intPos01)
16440             intPos01 = InStr(2, strTmp02, ",")
16450             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16460             strTmp02 = Mid(strTmp02, intPos01)
16470             intPos01 = InStr(2, strTmp02, ",")
16480             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16490             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
16500             strSortNow = strTmp01
16510           Case 5
16520             intPos01 = InStr(strSortAsc, ",")
16530             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
16540             strTmp02 = Mid(strSortAsc, intPos01)
16550             intPos01 = InStr(2, strTmp02, ",")
16560             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16570             strTmp02 = Mid(strTmp02, intPos01)
16580             intPos01 = InStr(2, strTmp02, ",")
16590             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16600             strTmp02 = Mid(strTmp02, intPos01)
16610             intPos01 = InStr(2, strTmp02, ",")
16620             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16630             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
16640             strSortNow = strTmp01
16650           End Select
16660         Else
16670           strSortNow = strSortAsc
16680         End If
16690       End If
16700     End If
16710 On Error Resume Next
16720     .OrderBy = strSortNow
16730     .OrderByOn = True
16740 On Error GoTo ERRH
16750   End With

EXITP:
16760   Set ctl = Nothing
16770   Exit Sub

ERRH:
16780   Select Case ERR.Number
        Case Else
16790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16800   End Select
16810   Resume EXITP

End Sub

Public Function SortNow_Get() As String

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

16910   strRetVal = strSortNow

EXITP:
16920   SortNow_Get = strRetVal
16930   Exit Function

ERRH:
16940   strRetVal = vbNullString
16950   Select Case ERR.Number
        Case Else
16960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16970   End Select
16980   Resume EXITP

End Function

Public Sub SortNow_Set(strSort As String)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Set"

17010   strSortNow = strSort

EXITP:
17020   Exit Sub

ERRH:
17030   Select Case ERR.Number
        Case Else
17040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17050   End Select
17060   Resume EXITP

End Sub

Public Sub FilterRecs_Set(strFltr As String, dblFltrRecs As Double)
' ** Initialize the local strFilter and dblFilterRecs with frmTransaction_Audit's copy.

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Set"

17110   strFilter = strFltr
17120   dblFilterRecs = dblFltrRecs
17130   strFilter = SwapNames(strFilter)  ' ** Function: Below.

EXITP:
17140   Exit Sub

ERRH:
17150   Select Case ERR.Number
        Case Else
17160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17170   End Select
17180   Resume EXITP

End Sub

Public Sub TotalRecs_Set(lngTotRecs As Long)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "TotalRecs_Set"

17210   lngTotalRecs = lngTotRecs

EXITP:
17220   Exit Sub

ERRH:
17230   Select Case ERR.Number
        Case Else
17240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17250   End Select
17260   Resume EXITP

End Sub

Public Function SwapNames(strFilt As String, Optional varReverse As Variant) As String
' ** Replace table field names with this form's visible field names.

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "SwapNames"

        Dim blnReverse As Boolean
        Dim strRetVal As String

17310   If strFilt <> vbNullString Then
17320     Select Case IsMissing(varReverse)
          Case True
17330       blnReverse = False
17340     Case False
17350       blnReverse = varReverse
17360     End Select
17370     strRetVal = strFilt
17380     Select Case blnReverse
          Case True
17390       strRetVal = StringReplace(strRetVal, "[Journal Number]", "[journalno]")  ' ** Module Function: modStringFuncs.
17400       strRetVal = StringReplace(strRetVal, "[Journal Type]", "[journaltype]")  ' ** Module Function: modStringFuncs.
17410       strRetVal = StringReplace(strRetVal, "[Posting Date]", "[transdate]")  ' ** Module Function: modStringFuncs.
17420       strRetVal = StringReplace(strRetVal, "[Account Number]", "[accountno]")  ' ** Module Function: modStringFuncs.
17430       strRetVal = StringReplace(strRetVal, "[Name]", "[shortname]")  ' ** Module Function: modStringFuncs.
17440       strRetVal = StringReplace(strRetVal, "[CUSIP]", "[cusix]")  ' ** Module Function: modStringFuncs.
17450       strRetVal = StringReplace(strRetVal, "[cusix]", "[cusip]")  ' ** Module Function: modStringFuncs.
17460       strRetVal = StringReplace(strRetVal, "[Asset]", "[asset_description]")  ' ** Module Function: modStringFuncs.
17470       strRetVal = StringReplace(strRetVal, "[Share / Face]", "[shareface]")  ' ** Module Function: modStringFuncs.
17480       strRetVal = StringReplace(strRetVal, "[Income Cash]", "[icash]")  ' ** Module Function: modStringFuncs.
17490       strRetVal = StringReplace(strRetVal, "[Principal Cash]", "[pcash]")  ' ** Module Function: modStringFuncs.
17500       strRetVal = StringReplace(strRetVal, "[Cost]", "[cosx]")  ' ** Module Function: modStringFuncs.
17510       strRetVal = StringReplace(strRetVal, "[cosx]", "[cost]")  ' ** Module Function: modStringFuncs.
17520       strRetVal = StringReplace(strRetVal, "[Currency]", "[curr_id]")  ' ** Module Function: modStringFuncs.
17530       strRetVal = StringReplace(strRetVal, "[Trade Date]", "[assetdate]")  ' ** Module Function: modStringFuncs.
17540       strRetVal = StringReplace(strRetVal, "[Original Trade Date]", "[PurchaseDate]")  ' ** Module Function: modStringFuncs.
17550       strRetVal = StringReplace(strRetVal, "[Comments]", "[ledger_description]")  ' ** Module Function: modStringFuncs.
17560       strRetVal = StringReplace(strRetVal, "[Recurring Item]", "[RecurringItem]")  ' ** Module Function: modStringFuncs.
17570       strRetVal = StringReplace(strRetVal, "[Inc/Exp Codes]", "[revcode_DESC]")  ' ** Module Function: modStringFuncs.
17580       strRetVal = StringReplace(strRetVal, "[Inc/Exp]", "[revcode_TYPE_Description]")  ' ** Module Function: modStringFuncs.
17590       strRetVal = StringReplace(strRetVal, "[Tax Codes]", "[taxcode_description]")  ' ** Module Function: modStringFuncs.
17600       strRetVal = StringReplace(strRetVal, "[Inc/Ded]", "[taxcode_type_description]")  ' ** Module Function: modStringFuncs.
17610       strRetVal = StringReplace(strRetVal, "[Location]", "[Location_Name]")  ' ** Module Function: modStringFuncs.
17620       strRetVal = StringReplace(strRetVal, "[Check Number]", "[CheckNum]")  ' ** Module Function: modStringFuncs.
17630       strRetVal = StringReplace(strRetVal, "[User]", "[journal_USER]")  ' ** Module Function: modStringFuncs.
17640       strRetVal = StringReplace(strRetVal, "[Date Posted]", "[posted]")  ' ** Module Function: modStringFuncs.
17650       strRetVal = StringReplace(strRetVal, "[Hidden]", "[ledger_HIDDEN]")  ' ** Module Function: modStringFuncs.
17660     Case False
17670       strRetVal = StringReplace(strRetVal, "[journalno]", "[Journal Number]")  ' ** Module Function: modStringFuncs.
17680       strRetVal = StringReplace(strRetVal, "[journaltype]", "[Journal Type]")  ' ** Module Function: modStringFuncs.
17690       strRetVal = StringReplace(strRetVal, "[transdate]", "[Posting Date]")  ' ** Module Function: modStringFuncs.
17700       strRetVal = StringReplace(strRetVal, "[accountno]", "[Account Number]")  ' ** Module Function: modStringFuncs.
17710       strRetVal = StringReplace(strRetVal, "[shortname]", "[Name]")  ' ** Module Function: modStringFuncs.
17720       strRetVal = StringReplace(strRetVal, "[cusip]", "[CUSIX]")  ' ** Module Function: modStringFuncs.
17730       strRetVal = StringReplace(strRetVal, "[CUSIX]", "[CUSIP]")  ' ** Module Function: modStringFuncs.
17740       strRetVal = StringReplace(strRetVal, "[asset_description]", "[Asset]")  ' ** Module Function: modStringFuncs.
17750       strRetVal = StringReplace(strRetVal, "[shareface]", "[Share / Face]")  ' ** Module Function: modStringFuncs.
17760       strRetVal = StringReplace(strRetVal, "[icash]", "[Income Cash]")  ' ** Module Function: modStringFuncs.
17770       strRetVal = StringReplace(strRetVal, "[pcash]", "[Principal Cash]")  ' ** Module Function: modStringFuncs.
17780       strRetVal = StringReplace(strRetVal, "[cost]", "[Cosx]")  ' ** Module Function: modStringFuncs.
17790       strRetVal = StringReplace(strRetVal, "[Cosx]", "[Cost]")  ' ** Module Function: modStringFuncs.
17800       strRetVal = StringReplace(strRetVal, "[curr_id]", "[Currency]")  ' ** Module Function: modStringFuncs.
17810       strRetVal = StringReplace(strRetVal, "[assetdate]", "[Trade Date]")  ' ** Module Function: modStringFuncs.
17820       strRetVal = StringReplace(strRetVal, "[PurchaseDate]", "[Original Trade Date]")  ' ** Module Function: modStringFuncs.
17830       strRetVal = StringReplace(strRetVal, "[ledger_description]", "[Comments]")  ' ** Module Function: modStringFuncs.
17840       strRetVal = StringReplace(strRetVal, "[RecurringItem]", "[Recurring Item]")  ' ** Module Function: modStringFuncs.
17850       strRetVal = StringReplace(strRetVal, "[revcode_DESC]", "[Inc/Exp Codes]")  ' ** Module Function: modStringFuncs.
17860       strRetVal = StringReplace(strRetVal, "[revcode_ID]", "[Inc/Exp Codes]")  ' ** Module Function: modStringFuncs.
17870       strRetVal = StringReplace(strRetVal, "[revcode_TYPE_Description]", "[Inc/Exp]")  ' ** Module Function: modStringFuncs.
17880       strRetVal = StringReplace(strRetVal, "[revcode_TYPE]", "[Inc/Exp]")  ' ** Module Function: modStringFuncs.
17890       strRetVal = StringReplace(strRetVal, "[taxcode_description]", "[Tax Codes]")  ' ** Module Function: modStringFuncs.
17900       strRetVal = StringReplace(strRetVal, "[taxcode]", "[Tax Codes]")  ' ** Module Function: modStringFuncs.
17910       strRetVal = StringReplace(strRetVal, "[taxcode_type_description]", "[Inc/Ded]")  ' ** Module Function: modStringFuncs.
17920       strRetVal = StringReplace(strRetVal, "[taxcode_type]", "[Inc/Ded]")  ' ** Module Function: modStringFuncs.
17930       strRetVal = StringReplace(strRetVal, "[Location_Name]", "[Location]")  ' ** Module Function: modStringFuncs.
17940       strRetVal = StringReplace(strRetVal, "[CheckNum]", "[Check Number]")  ' ** Module Function: modStringFuncs.
17950       strRetVal = StringReplace(strRetVal, "[journal_USER]", "[User]")  ' ** Module Function: modStringFuncs.
17960       strRetVal = StringReplace(strRetVal, "[posted]", "[Date Posted]")  ' ** Module Function: modStringFuncs.
17970       strRetVal = StringReplace(strRetVal, "[ledger_HIDDEN]", "[Hidden]")  ' ** Module Function: modStringFuncs.
17980     End Select
17990   End If

EXITP:
18000   SwapNames = strRetVal
18010   Exit Function

ERRH:
18020   Select Case ERR.Number
        Case Else
18030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18040   End Select
18050   Resume EXITP

End Function

Private Sub ColArray_Load()

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "ColArray_Load"

        Dim frm As Access.Form, ctl As Access.Control
        Dim lngX As Long, lngE As Long

18110   With Me
18120     If lngCols = 0& Or IsEmpty(arr_varCol) = True Then

18130       lngCols = 0&
18140       ReDim arr_varCol(C_ELEMS, 0)

18150       Set frm = .Parent
18160       With frm
18170         For Each ctl In .Detail.Controls
18180           With ctl
18190             If .ControlType = acCheckBox Then
                    ' ** ckgFlds_chkJournalType.
18200               If Left(.Name, 8) = "ckgFlds_" Then
18210                 lngCols = lngCols + 1&
18220                 lngE = lngCols - 1&
18230                 ReDim Preserve arr_varCol(C_ELEMS, lngE)
18240                 arr_varCol(C_VIEWCHK, lngE) = .Name
18250                 arr_varCol(C_VIEWVAL, lngE) = .Value
18260                 arr_varCol(C_COLNAM, lngE) = Null
18270                 arr_varCol(C_HASSP, lngE) = CBool(False)
18280                 arr_varCol(C_WDT, lngE) = CLng(0)
18290               End If
18300             End If
18310           End With  ' ** ctl.
18320         Next  ' ** ctl.
18330       End With  ' ** frm.

            'SHOULD THESE BE PUT INTO THE TAG?
            ' ** Cross-reference check boxes with columns.
18340       For lngX = 0& To (lngCols - 1&)
18350         Select Case arr_varCol(C_VIEWCHK, lngX)
              Case "ckgFlds_chkJournalno"
18360           arr_varCol(C_COLNAM, lngX) = "Journal Number"
18370           arr_varCol(C_HASSP, lngX) = CBool(True)
18380         Case "ckgFlds_chkJournalType"
18390           arr_varCol(C_COLNAM, lngX) = "Journal Type"
18400           arr_varCol(C_HASSP, lngX) = CBool(True)
18410         Case "ckgFlds_chkTransDate"
18420           arr_varCol(C_COLNAM, lngX) = "Posting Date"
18430           arr_varCol(C_HASSP, lngX) = CBool(True)
18440         Case "ckgFlds_chkAccountNo"
18450           arr_varCol(C_COLNAM, lngX) = "Account Number"
18460           arr_varCol(C_HASSP, lngX) = CBool(True)
18470         Case "ckgFlds_chkShortName"
18480           arr_varCol(C_COLNAM, lngX) = "Name"
18490         Case "ckgFlds_chkCusip"
18500           arr_varCol(C_COLNAM, lngX) = "CUSIP"
18510         Case "ckgFlds_chkAssetDescription"
18520           arr_varCol(C_COLNAM, lngX) = "Asset"
18530         Case "ckgFlds_chkShareFace"
18540           arr_varCol(C_COLNAM, lngX) = "Share / Face"
18550           arr_varCol(C_HASSP, lngX) = CBool(True)
18560         Case "ckgFlds_chkICash"
18570           arr_varCol(C_COLNAM, lngX) = "Income Cash"
18580           arr_varCol(C_HASSP, lngX) = CBool(True)
18590         Case "ckgFlds_chkPCash"
18600           arr_varCol(C_COLNAM, lngX) = "Principal Cash"
18610           arr_varCol(C_HASSP, lngX) = CBool(True)
18620         Case "ckgFlds_chkCost"
18630           arr_varCol(C_COLNAM, lngX) = "Cost"
18640         Case "ckgFlds_chkCurrID"
18650           arr_varCol(C_COLNAM, lngX) = "Currency"
18660         Case "ckgFlds_chkAssetDate"
18670           arr_varCol(C_COLNAM, lngX) = "Trade Date"
18680           arr_varCol(C_HASSP, lngX) = CBool(True)
18690         Case "ckgFlds_chkPurchaseDate"
18700           arr_varCol(C_COLNAM, lngX) = "Original Trade Date"
18710           arr_varCol(C_HASSP, lngX) = CBool(True)
18720         Case "ckgFlds_chkLedgerDescription"
18730           arr_varCol(C_COLNAM, lngX) = "Comments"
18740         Case "ckgFlds_chkRecurringItem"
18750           arr_varCol(C_COLNAM, lngX) = "Recurring Item"
18760           arr_varCol(C_HASSP, lngX) = CBool(True)
18770         Case "ckgFlds_chkRevCodeDesc"
18780           arr_varCol(C_COLNAM, lngX) = "Inc/Exp Codes"
18790           arr_varCol(C_HASSP, lngX) = CBool(True)
18800         Case "ckgFlds_chkRevCodeTypeDescription"
18810           arr_varCol(C_COLNAM, lngX) = "Inc/Exp"
18820         Case "ckgFlds_chkTaxCodeDescription"
18830           arr_varCol(C_COLNAM, lngX) = "Tax Codes"
18840           arr_varCol(C_HASSP, lngX) = CBool(True)
18850         Case "ckgFlds_chkTaxCodeTypeDescription"
18860           arr_varCol(C_COLNAM, lngX) = "Inc/Ded"
18870         Case "ckgFlds_chkLocationName"
18880           arr_varCol(C_COLNAM, lngX) = "Location"
18890         Case "ckgFlds_chkCheckNum"
18900           arr_varCol(C_COLNAM, lngX) = "Check Number"
18910           arr_varCol(C_HASSP, lngX) = CBool(True)
18920         Case "ckgFlds_chkJournalUser"
18930           arr_varCol(C_COLNAM, lngX) = "User"
18940         Case "ckgFlds_chkPosted"
18950           arr_varCol(C_COLNAM, lngX) = "Date Posted"
18960           arr_varCol(C_HASSP, lngX) = CBool(True)
18970         Case "ckgFlds_chkLedgerHidden"
18980           arr_varCol(C_COLNAM, lngX) = "Hidden"
18990         End Select

19000       Next  ' ** lngX.

19010       For lngX = 0& To (lngCols - 1&)
19020         arr_varCol(C_WDT, lngX) = .Controls(arr_varCol(C_COLNAM, lngX)).ColumnWidth
19030       Next  ' ** lngX.

19040     End If
19050   End With

        'ckgFlds_chkJournalno
        'ckgFlds_chkJournalType
        'ckgFlds_chkTransDate
        'ckgFlds_chkAccountNo
        'ckgFlds_chkShortName
        'ckgFlds_chkCusip
        'ckgFlds_chkAssetDescription
        'ckgFlds_chkShareFace
        'ckgFlds_chkICash
        'ckgFlds_chkPCash
        'ckgFlds_chkCost
        'ckgFlds_chkCurrID
        'ckgFlds_chkAssetDate
        'ckgFlds_chkPurchaseDate
        'ckgFlds_chkLedgerDescription
        'ckgFlds_chkRecurringItem
        'ckgFlds_chkRevCodeDesc
        'ckgFlds_chkRevCodeTypeDescription
        'ckgFlds_chkTaxCodeDescription
        'ckgFlds_chkTaxCodeTypeDescription
        'ckgFlds_chkLocationName
        'ckgFlds_chkCheckNum
        'ckgFlds_chkJournalUser
        'ckgFlds_chkPosted
        'ckgFlds_chkLedgerHidden

        'Journal Number
        'Journal Type
        'Posting Date
        'Account Number
        'Name
        'CUSIP
        'Asset
        'Share / Face
        'Income Cash
        'Principal Cash
        'Cost
        'Currency
        'Trade Date
        'Original Trade Date
        'Comments
        'Recurring Item
        'Inc/Exp Codes
        'Inc/Exp
        'Tax Codes
        'Inc/Ded
        'Location
        'Check Number
        'User
        'Date Posted
        'Hidden

EXITP:
19060   Set ctl = Nothing
19070   Set frm = Nothing
19080   Exit Sub

ERRH:
19090   Select Case ERR.Number
        Case Else
19100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19110   End Select
19120   Resume EXITP

End Sub

Public Sub ColArray_Width()

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "ColArray_Width"

        Dim frm As Access.Form
        Dim lngX As Long

19210   With Me
19220     Set frm = .Parent
19230     For lngX = 0& To (lngCols - 1&)
19240       arr_varCol(C_WDT, lngX) = .Controls(arr_varCol(C_COLNAM, lngX)).ColumnWidth
            ' ** col13_width
19250       frm.Controls("col" & Right("00" & CStr(lngX), 2) & "_width") = arr_varCol(C_WDT, lngX)
19260     Next  ' ** lngX.
19270   End With

EXITP:
19280   Set frm = Nothing
19290   Exit Sub

ERRH:
19300   Select Case ERR.Number
        Case Else
19310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19320   End Select
19330   Resume EXITP

End Sub

Public Sub ColWidth_Set(blnSet As Boolean)

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "ColWidth_Set"

19410   With Me
19420     blnColWidth = blnSet
19430     .TimerInterval = 100&
19440   End With

EXITP:
19450   Exit Sub

ERRH:
19460   Select Case ERR.Number
        Case Else
19470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19480   End Select
19490   Resume EXITP

End Sub
