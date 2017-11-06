VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountTaxCodes_BlockAssign_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountTaxCodes_BlockAssign_Sub"

'VGC 10/02/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Assign:           A {cmdAssign on frmAccountTaxCodes_BlockAssign}
' **   Cancel:           C {cmdCancel on frmAccountTaxCodes_BlockAssign}
' **   Tax Code:         D {cmbTaxcodes on frmAccountTaxCodes_BlockAssign}}
' **   Clear:            E {cmdClear on frmAccountTaxCodes_BlockAssign}
' **   Show JournalNo:   H {chkShowJournalNo on frmAccountTaxCodes_BlockAssign}
' **   Filter by JType:  J {cmbJournalTypes on frmAccountTaxCodes_BlockAssign}
' **   Checked Only:     K {chkShowCheckedOnly on frmAccountTaxCodes_BlockAssign}
' **   Select None:      N {cmdSelectNone on frmAccountTaxCodes_BlockAssign}
' **   OK:               O {cmdOK on frmAccountTaxCodes_BlockAssign}
' **   Select All:       S {cmdSelectAll on frmAccountTaxCodes_BlockAssign}
' **   Include Currency: U {chkIncludeCurrency on frmAccountTaxCodes_BlockAssign}
' **   Unassigned Only:  W {chkShowUnassignedOnly on frmAccountTaxCodes_BlockAssign}
' **   Check This:       X {chkTaxChk}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:           F5 {RecalcTots}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ** Array: arr_varCtl().
Private lngCtls As Long, arr_varCtl() As Variant
Private Const C_ELEMS As Integer = 7  ' ** Array's first-element UBound().
Private Const C_CNAM   As Integer = 0
Private Const C_LFT    As Integer = 1
Private Const C_L1_NAM As Integer = 2
Private Const C_L1_LFT As Integer = 3
Private Const C_L2_NAM As Integer = 4
Private Const C_L2_LFT As Integer = 5
Private Const C_LN_NAM As Integer = 6
Private Const C_LN_LFT As Integer = 7

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[transdate], [JournalType_Order], [totdesc]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngForm_Width As Long, lngJournalNo_Width As Long, lngCurrID_Width As Long

Private blnRecalcTots As Boolean
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccountTaxCodes_BlockAssign" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

180         blnRecalcTots = False

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

200         .form_width_line.Left = 0&  ' ** Just get this out of the way.

210         lngSortLbl_Width = .Sort_lbl.Width
220         lngForm_Width = .Width
230         lngJournalNo_Width = (.transdate.Left - .journalno.Left)  ' ** The amount of space it uses.
240         lngCurrID_Width = ((.taxcode.Left - .curr_id.Left) - (4& * lngTpp))
250         .Parent.SetJNoWidth lngJournalNo_Width  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign.
260         .Parent.SetCurrWidth lngCurrID_Width  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign.

270         CtlArrayLoad  ' ** Procedure: Below.

280         gblnHasForEx = HasForEx_All  ' ** Module Function: modCurrencyFuncs.

290         Select Case gblnHasForEx
            Case True
300           Select Case gblnHasForExThis
              Case True
310             .curr_id.Visible = True
320             .curr_id_lbl.Visible = True
330             .curr_id_lbl_line.Visible = True
340             .curr_id_box.Visible = True
350             .curr_id_forex_bg.Visible = True
360             .ICash.Visible = False
370             .icash_str.Visible = True
380             .PCash.Visible = False
390             .pcash_str.Visible = True
400             .Cost.Visible = False
410             .cost_str.Visible = True
420           Case False
430             .curr_id.Visible = False
440             .curr_id_lbl.Visible = False
450             .curr_id_lbl_line.Visible = False
460             .curr_id_box.Visible = False
470             .curr_id_forex_bg.Visible = False
480             .curr_id.Left = 0&
490             .curr_id_lbl.Left = 0&
500             .curr_id_lbl_line.Left = 0&
510             .curr_id_box.Left = 0&
520             .curr_id_forex_bg.Left = 0&
530             .ICash.Visible = True
540             .icash_str.Visible = False
550             .PCash.Visible = True
560             .pcash_str.Visible = False
570             .Cost.Visible = True
580             .cost_str.Visible = False
590           End Select
600         Case False
610           .curr_id.Visible = False
620           .curr_id_lbl.Visible = False
630           .curr_id_lbl_line.Visible = False
640           .curr_id_box.Visible = False
650           .curr_id_forex_bg.Visible = False
660           .curr_id.Left = 0&
670           .curr_id_lbl.Left = 0&
680           .curr_id_lbl_line.Left = 0&
690           .curr_id_box.Left = 0&
700           .curr_id_forex_bg.Left = 0&
710           .ICash.Visible = True
720           .icash_str.Visible = False
730           .PCash.Visible = True
740           .pcash_str.Visible = False
750           .Cost.Visible = True
760           .cost_str.Visible = False
770         End Select

780   On Error Resume Next
790         .FocusHolder.SetFocus
800   On Error GoTo ERRH

810       Else
820         Cancel = -1
830       End If
840     End With

850     If Cancel = -1 Then
860       Beep
870       DoCmd.Hourglass False
880       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
890     End If

EXITP:
900     Exit Sub

ERRH:
910     Select Case ERR.Number
        Case Else
920       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
930     End Select
940     Resume EXITP

End Sub

Private Sub Form_Load()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1010    With Me

1020      .journalno.ForeColor = CLR_DISABLED_FG
1030      .journalno.BackColor = CLR_DISABLED_BG
1040      .transdate.ForeColor = CLR_DISABLED_FG
1050      .transdate.BackColor = CLR_DISABLED_BG
1060      .journaltype.ForeColor = CLR_DISABLED_FG
1070      .journaltype.BackColor = CLR_DISABLED_BG
1080      .totdesc.ForeColor = CLR_DISABLED_FG
1090      .totdesc.BackColor = CLR_DISABLED_BG
1100      .ICash.ForeColor = CLR_DISABLED_FG
1110      .ICash.BackColor = CLR_DISABLED_BG
1120      .icash_str.ForeColor = CLR_DISABLED_FG
1130      .icash_str.BackColor = CLR_DISABLED_BG
1140      .PCash.ForeColor = CLR_DISABLED_FG
1150      .PCash.BackColor = CLR_DISABLED_BG
1160      .pcash_str.ForeColor = CLR_DISABLED_FG
1170      .pcash_str.BackColor = CLR_DISABLED_BG
1180      .Cost.ForeColor = CLR_DISABLED_FG
1190      .Cost.BackColor = CLR_DISABLED_BG
1200      .cost_str.ForeColor = CLR_DISABLED_FG
1210      .cost_str.BackColor = CLR_DISABLED_BG
1220      .curr_id.ForeColor = CLR_DISABLED_FG
1230      .curr_id.BackColor = CLR_DISABLED_BG
1240      .taxcode.ForeColor = CLR_DISABLED_FG
1250      .taxcode.BackColor = CLR_DISABLED_BG

1260      SortNow THIS_PROC  ' ** Procedure: Below.

1270    End With

EXITP:
1280    Exit Sub

ERRH:
1290    Select Case ERR.Number
        Case Else
1300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1310    End Select
1320    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1420    intShiftDown = (Shift And acShiftMask) > 0
1430    intAltDown = (Shift And acAltMask) > 0
1440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc:           F5 {RecalcTots}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Rev Code:         D {cmbRevenueCodes on frmAccountTaxCodes_BlockAssign}}
        ' **   Filter by JType:  J {cmbJournalTypes on frmAccountTaxCodes_BlockAssign}
        ' **   Check This:       X {chkTaxChk}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyF5
1470        intRetVal = 0
1480        RecalcTots  ' ** Procedure: Below.
1490      Case vbKeyUp
1500        intRetVal = 0
1510        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1520      Case vbKeyDown
1530        intRetVal = 0
1540        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1550      Case vbKeyEscape
1560        With Me
1570          intRetVal = 0
1580          .Parent.cmdCancel_Click  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign.
1590        End With
1600      End Select
1610    End If

        ' ** Alt keys.
1620    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1630      Select Case intRetVal
          Case vbKeyD
1640        With Me
1650          intRetVal = 0
1660          DoCmd.SelectObject acForm, .Parent.Name, False
1670          .Parent.cmbTaxCodes.SetFocus
1680        End With
1690      Case vbKeyJ
1700        With Me
1710          intRetVal = 0
1720          DoCmd.SelectObject acForm, .Parent.Name, False
1730          .Parent.cmbJournalTypes.SetFocus
1740        End With
1750      Case vbKeyX
1760        With Me
1770          intRetVal = 0
1780          .chkTaxChk.SetFocus
1790          .chkTaxChk = (Not .chkTaxChk)
1800          chkTaxChk_AfterUpdate  ' ** Procedure: Below.
1810        End With
1820      End Select
1830    End If

        ' ** Ctrl keys.
1840    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1850      Select Case intRetVal
          Case vbKeyS
1860        intRetVal = 0
1870        cmdSave_Click  ' ** Procedure: Below.
1880      Case vbKeyTab, vbKeyReturn
1890        With Me
1900          intRetVal = 0
1910          DoCmd.SelectObject acForm, .Parent.Name, False
1920          .Parent.cmSelectAll.SetFocus
1930        End With
1940      Case vbKeyPageUp, vbKeyHome
1950        intRetVal = 0
1960        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1970      Case vbKeyPageDown, vbKeyEnd
1980        intRetVal = 0
1990        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2000      End Select
2010    End If

        ' ** Ctrl-Shift keys.
2020    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2030      Select Case intRetVal
          Case vbKeyF
2040        With Me
2050          intRetVal = 0
2060          DoCmd.SelectObject acForm, .Parent.Name, False
2070          .Parent.FocusHolder.SetFocus
2080        End With
2090      Case vbKeyTab, vbKeyReturn
2100        With Me
2110          intRetVal = 0
2120          DoCmd.SelectObject acForm, .Parent.Name, False
2130          .Parent.cmbTaxCodes.SetFocus
2140        End With
2150      End Select
2160    End If

EXITP:
2170    KeyCode = intRetVal
2180    Exit Sub

ERRH:
2190    intRetVal = 0
2200    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2210    Case Else
2220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2230    End Select
2240    Resume EXITP

End Sub

Private Sub Form_Timer()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

2310    With Me
2320      .TimerInterval = 0&
2330      If blnRecalcTots = True Then
2340        blnRecalcTots = False
2350        RecalcTots  ' ** Procedure: Below.
2360      End If
2370    End With

EXITP:
2380    Exit Sub

ERRH:
2390    Select Case ERR.Number
        Case Else
2400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2410    End Select
2420    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2510    If DataErr <> 0 Then
2520      Select Case DataErr
          Case 2115  ' ** The macro or function set to the BeforeUpdate or ValidationRule property for this field is preventing Microsoft Access from saving the data in the field.
            ' ** RecalcTots getting piled up, Form_Timer() calling here and parent form.
2530        gblnSignal = False
2540        Response = acDataErrContinue
2550      Case Else
2560        zErrorHandler THIS_NAME, THIS_PROC, DataErr, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2570      End Select
2580    End If

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
2590    Exit Sub

ERRH:
2600    Select Case ERR.Number
        Case Else
2610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2620    End Select
2630    Resume EXITP

End Sub

Public Sub cmdSave_Click()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2710    DoCmd.RunCommand acCmdSaveRecord
2720    DoEvents

EXITP:
2730    Exit Sub

ERRH:
2740    Select Case ERR.Number
        Case Else
2750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2760    End Select
2770    Resume EXITP

End Sub

Private Sub journalno_lbl_DblClick(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_lbl_DblClick"

2810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

2910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2920    Exit Sub

ERRH:
2930    Select Case ERR.Number
        Case Else
2940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2950    End Select
2960    Resume EXITP

End Sub

Private Sub journaltype_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_lbl_DblClick"

3010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub chkTaxChk_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTaxChk_lbl_DblClick"

3110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Public Sub chkTaxChk_AfterUpdate()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTaxChk_AfterUpdate"

3210    With Me

3220      cmdSave_Click  ' ** Procedure: Above.

3230      blnRecalcTots = True
3240      .TimerInterval = 500&

3250    End With

EXITP:
3260    Exit Sub

ERRH:
3270    Select Case ERR.Number
        Case Else
3280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290    End Select
3300    Resume EXITP

End Sub

Private Sub chkTaxChk_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTaxChk_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3420    intShiftDown = (Shift And acShiftMask) > 0
3430    intAltDown = (Shift And acAltMask) > 0
3440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3470        With Me
3480          intRetVal = 0
3490          lngRecsCur = RecCnt  ' ** Function: Below.
3500          If .CurrentRecord < lngRecsCur Then
3510            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3520            .chkTaxChk.SetFocus
3530          Else
3540            DoCmd.SelectObject acForm, .Parent.Name, False
3550            .Parent.cmdSelectAll.SetFocus
3560          End If
3570        End With
3580      End Select
3590    End If

        ' ** Shift keys.
3600    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3610      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3620        With Me
3630          intRetVal = 0
3640          If .CurrentRecord > 1 Then
3650            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3660            .chkTaxChk.SetFocus
3670          Else
3680            DoCmd.SelectObject acForm, .Parent.Name, False
3690            .Parent.cmbTaxCodes.SetFocus
3700          End If
3710        End With
3720      End Select
3730    End If

EXITP:
3740    KeyCode = intRetVal
3750    Exit Sub

ERRH:
3760    intRetVal = 0
3770    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3780    Case Else
3790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3800    End Select
3810    Resume EXITP

End Sub

Private Sub totdesc_lbl_DblClick(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_DblClick"

3910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub icash_lbl_DblClick(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_lbl_DblClick"

4010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub pcash_lbl_DblClick(Cancel As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_lbl_DblClick"

4110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub cost_lbl_DblClick(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_lbl_DblClick"

4210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4220    Exit Sub

ERRH:
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

4310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Private Sub taxcode_lbl_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_lbl_DblClick"

4410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub taxcode_description_display_GotFocus()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_display_GotFocus"

4510    With Me
4520      .taxcode.SetFocus
4530    End With

EXITP:
4540    Exit Sub

ERRH:
4550    Select Case ERR.Number
        Case Else
4560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4570    End Select
4580    Resume EXITP

End Sub

Private Sub taxcode_AfterUpdate()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_AfterUpdate"

4610    With Me
4620      If IsNull(.taxcode) = True Then
4630        .taxcode = TAXTYP_INC
4640      End If
4650      cmdSave_Click  ' ** Procedure: Above.
4660    End With

EXITP:
4670    Exit Sub

ERRH:
4680    Select Case ERR.Number
        Case Else
4690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4700    End Select
4710    Resume EXITP

End Sub

Private Sub taxcode_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4820    intShiftDown = (Shift And acShiftMask) > 0
4830    intAltDown = (Shift And acAltMask) > 0
4840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4870        With Me
4880          intRetVal = 0
4890          lngRecsCur = RecCnt  ' ** Function: Below.
4900          If .CurrentRecord < lngRecsCur Then
4910            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4920            .chkTaxChk.SetFocus
4930          Else
4940            DoCmd.SelectObject acForm, .Parent.Name, False
4950            .Parent.cmdSelectAll.SetFocus
4960          End If
4970        End With
4980      End Select
4990    End If

        ' ** Shift keys.
5000    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5010      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5020        With Me
5030          intRetVal = 0
5040          .chkTaxChk.SetFocus
5050        End With
5060      End Select
5070    End If

EXITP:
5080    KeyCode = intRetVal
5090    Exit Sub

ERRH:
5100    intRetVal = 0
5110    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5120    Case Else
5130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5140    End Select
5150    Resume EXITP

End Sub

Public Sub RecalcTots()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngTaxcode As Long
        Dim varTmp00 As Variant, lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long, lngTmp05 As Long

5210    With Me
5220      If gblnSignal = False Then
5230        gblnSignal = True  ' ** Trying to prevent too many calls from piling up.

5240        lngRecsCur = RecCnt  ' ** Function: Below.
5250        If lngRecsCur > 0& Then

5260          .chkTaxChk_sum.Requery
5270          Select Case IsNull(.Parent.cmbTaxCodes)
              Case True
5280            lngTaxcode = 0&
5290          Case False
5300            lngTaxcode = .Parent.cmbTaxCodes.Column(0)
5310          End Select

5320          varTmp00 = .chkTaxChk_sum
5330          If IsNull(varTmp00) = False Then
5340            lngTmp01 = Abs(varTmp00)
5350            .Parent.chkTaxChk_sum = lngTmp01
5360          Else
5370            lngTmp01 = 0&
5380            .Parent.chkTaxChk_sum = lngTmp01
5390          End If

5400          varTmp00 = .tcbas_changed_sum
5410          If IsNull(varTmp00) = False Then
5420            lngTmp02 = Abs(varTmp00)
5430          Else
5440            lngTmp02 = 0&
5450          End If

5460          Set dbs = CurrentDb
              ' ** tblTaxCode_BlockAssign_Staging, grouped and summed,
              ' ** with chkTaxChk_sum, tcbas_changed_sum, cnt_recs.
5470          Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_Block_07")
5480          Set rst = qdf.OpenRecordset
5490          If rst.BOF = True And rst.EOF = True Then
                ' ** Well, who knows?
5500            lngTmp03 = 0&: lngTmp04 = 0&: lngTmp05 = 0&
5510          Else
                ' ** Remember, this recordset is the entire table, and this subform may be filtered.
                ' ** Also, sums will be negative, since they're on boolean fields.
5520            rst.MoveFirst
5530            lngTmp03 = Nz(Abs(rst![chkTaxChk_sum]), 0)   ' ** Total checks in entire group.
5540            lngTmp04 = Nz(Abs(rst![tcbas_changed_sum]), 0)  ' ** Total changed in entire group.
5550            lngTmp05 = rst![cnt_recs]                       ' ** Total records in entire group.
5560            .Parent.tcbas_changed_sum = lngTmp04
5570          End If
5580          rst.Close
5590          dbs.Close

5600        Else
5610          lngTmp01 = 0&: lngTmp03 = 0&: lngTmp04 = 0&: lngTmp05 = 0&
5620        End If  ' ** lngRecsCur.

            ' ** lngTmp01  : Checked items in this subform.
            ' ** lngTmp02  : Changed items in this subform.
            ' ** lngRecsCur: Records in this subform.
            ' ** lngTmp03  : Checked items in entire group.
            ' ** lngTmp04  : Changed items in entire group.
            ' ** lngTmp05  : Records in entire group.
5630        If lngRecsCur = lngTmp05 And lngRecsCur > 0& Then
              ' ** Subform not filtered.
5640          If lngTmp03 > 0& Then
5650            .Parent.cmdClear.Enabled = True
5660          Else
5670            .Parent.cmdClear.Enabled = False
5680          End If
5690          If lngTaxcode > 0& And lngTmp03 > 0& Then
5700            .Parent.cmdAssign.Enabled = True
5710          Else
5720            .Parent.cmdAssign.Enabled = False
5730          End If
5740          If lngTmp04 > 0& Then
5750            .Parent.cmdOK.Enabled = True
5760          Else
5770            .Parent.cmdOK.Enabled = False
5780          End If
5790        ElseIf lngRecsCur > 0& Then
              ' ** Subform filtered.
5800          If lngTmp01 > 0& Then
5810            .Parent.cmdClear.Enabled = True
5820          Else
5830            .Parent.cmdClear.Enabled = False
5840          End If
5850          If lngTaxcode > 0& And lngTmp01 > 0& Then
5860            .Parent.cmdAssign.Enabled = True
5870          Else
5880            .Parent.FocusHolder.SetFocus
5890            .Parent.cmdAssign.Enabled = False
5900          End If
5910          If lngTmp02 > 0& Then
5920            .Parent.cmdOK.Enabled = True
5930          Else
5940            .Parent.cmdOK.Enabled = False
5950          End If
5960        Else
5970          .Parent.cmdClear.Enabled = False
5980          .Parent.cmdAssign.Enabled = False
5990          .Parent.cmdOK.Enabled = False
6000        End If  ' ** lngRecsCur.

6010        gblnSignal = False
6020      End If
6030    End With

EXITP:
6040    Set rst = Nothing
6050    Set qdf = Nothing
6060    Set dbs = Nothing
6070    Exit Sub

ERRH:
6080    gblnSignal = False
6090    Select Case ERR.Number
        Case Else
6100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6110    End Select
6120    Resume EXITP

End Sub

Public Sub ShowJournalNo_Sub(blnShow As Boolean)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowJournalNo_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim blnCurrID As Boolean
        Dim lngTmp01 As Long
        Dim lngX As Long

6210    With Me

6220      blnCurrID = .Parent.chkIncludeCurrency
6230      blnSortHere = False: blnResort = False

          ' **********************************************
          ' ** Array: arr_varCtl()
          ' **
          ' **   Field  Element  Name         Constant
          ' **   =====  =======  ===========  ==========
          ' **     1       0     fld name     C_CNAM
          ' **     2       1     fld left     C_LFT
          ' **     3       2     lbl1 name    C_L1_NAM
          ' **     4       3     lbl1 left    C_L1_LFT
          ' **     5       4     lbl2 name    C_L2_NAM
          ' **     6       5     lbl2 left    C_L2_LFT
          ' **     7       6     line name    C_LN_NAM
          ' **     8       7     line left    C_LN_LFT
          ' **
          ' **********************************************

6240      Select Case blnShow
          Case True
6250        Select Case blnCurrID
            Case True
6260          .Width = lngForm_Width
6270        Case False
6280          .Width = (lngForm_Width - lngCurrID_Width)
6290        End Select
6300        For lngX = (lngCtls - 1&) To 0& Step -1&
6310          If arr_varCtl(C_CNAM, lngX) = "journalno" Then
6320            .journalno.Visible = True
6330            .journalno_lbl.Visible = True
6340            .journalno_lbl2.Visible = True
6350            .journalno_lbl_line.Visible = True
6360          Else
6370            If blnCurrID = False And arr_varCtl(C_CNAM, lngX) = "curr_id" Then
                  ' ** Skip it.
6380            Else
6390              If .Sort_line.Left = .Controls(arr_varCtl(C_L1_NAM, lngX)).Left Then blnSortHere = True
6400              If blnCurrID = False And arr_varCtl(C_CNAM, lngX) = "taxcode" Then
6410                lngTmp01 = lngCurrID_Width
6420              Else
6430                lngTmp01 = 0&
6440              End If
6450              .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) - lngTmp01)
6460              .Controls(arr_varCtl(C_L1_NAM, lngX)).Left = (arr_varCtl(C_L1_LFT, lngX) - lngTmp01)
6470              If IsNull(arr_varCtl(C_L2_NAM, lngX)) = False Then
6480                .Controls(arr_varCtl(C_L2_NAM, lngX)).Left = (arr_varCtl(C_L2_LFT, lngX) - lngTmp01)
6490              End If
6500              .Controls(arr_varCtl(C_LN_NAM, lngX)).Left = (arr_varCtl(C_LN_LFT, lngX) - lngTmp01)
6510              Select Case arr_varCtl(C_CNAM, lngX)
                  Case "icash", "pcash", "cost"
6520                .Controls(arr_varCtl(C_CNAM, lngX) & "_str").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
6530                .Controls(arr_varCtl(C_CNAM, lngX) & "_box").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
6540                .Controls(arr_varCtl(C_CNAM, lngX) & "_forex_bg").Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTpp)
6550              Case "curr_id"
6560                .Controls(arr_varCtl(C_CNAM, lngX) & "_box").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
6570                .Controls(arr_varCtl(C_CNAM, lngX) & "_forex_bg").Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTpp)
6580              Case "taxcode"
6590                .Controls(arr_varCtl(C_CNAM, lngX) & "_description_display").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
6600              Case Else
                    ' ** Nothing else.
6610              End Select
6620              If blnSortHere = True Then
6630                .Sort_line.Left = .Controls(arr_varCtl(C_L1_NAM, lngX)).Left
6640                .Sort_lbl.Left = ((.Controls(arr_varCtl(C_L1_NAM, lngX)).Left + _
                      .Controls(arr_varCtl(C_L1_NAM, lngX)).Width) - lngSortLbl_Width)
6650              End If
6660            End If
6670          End If
6680          blnSortHere = False
6690        Next
6700        Select Case blnCurrID
            Case True
6710          .IsArch_bg1.Width = lngForm_Width
6720          .IsArch_bg2.Width = lngForm_Width
6730        Case False
6740          .IsArch_bg1.Width = (lngForm_Width - lngCurrID_Width)
6750          .IsArch_bg2.Width = (lngForm_Width - lngCurrID_Width)
6760        End Select
6770        .IsIncome.Left = (.taxcode.Left - (8& * lngTpp))
6780        .IsDeduction.Left = (.taxcode.Left - (8& * lngTpp))
6790      Case False
6800        For lngX = 0& To (lngCtls - 1&)
6810          If arr_varCtl(C_CNAM, lngX) = "journalno" Then
6820            If .Sort_line.Left = .journalno_lbl.Left Then blnSortHere = True
6830            .journalno.Visible = False
6840            .journalno_lbl.Visible = False
6850            .journalno_lbl2.Visible = False
6860            .journalno_lbl_line.Visible = False
6870            If blnSortHere = True Then
6880              .Sort_line.Visible = False
6890              .Sort_lbl.Visible = False
6900              blnResort = True
6910            End If
6920          Else
6930            If blnCurrID = False And arr_varCtl(C_CNAM, lngX) = "curr_id" Then
                  ' ** Skip it.
6940            Else
6950              If .Sort_line.Left = .Controls(arr_varCtl(C_L1_NAM, lngX)).Left Then blnSortHere = True
6960              If blnCurrID = False And arr_varCtl(C_CNAM, lngX) = "taxcode" Then
6970                lngTmp01 = lngCurrID_Width
6980              Else
6990                lngTmp01 = 0&
7000              End If
7010              .Controls(arr_varCtl(C_CNAM, lngX)).Left = ((arr_varCtl(C_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
7020              .Controls(arr_varCtl(C_L1_NAM, lngX)).Left = ((arr_varCtl(C_L1_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
7030              If IsNull(arr_varCtl(C_L2_NAM, lngX)) = False Then
7040                .Controls(arr_varCtl(C_L2_NAM, lngX)).Left = ((arr_varCtl(C_L2_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
7050              End If
7060              .Controls(arr_varCtl(C_LN_NAM, lngX)).Left = ((arr_varCtl(C_LN_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
7070              Select Case arr_varCtl(C_CNAM, lngX)
                  Case "icash", "pcash", "cost"
7080                .Controls(arr_varCtl(C_CNAM, lngX) & "_str").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
7090                .Controls(arr_varCtl(C_CNAM, lngX) & "_box").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
7100                .Controls(arr_varCtl(C_CNAM, lngX) & "_forex_bg").Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTpp)
7110              Case "curr_id"
7120                .Controls(arr_varCtl(C_CNAM, lngX) & "_box").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
7130                .Controls(arr_varCtl(C_CNAM, lngX) & "_forex_bg").Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTpp)
7140              Case "taxcode"
7150                .Controls(arr_varCtl(C_CNAM, lngX) & "_description_display").Left = .Controls(arr_varCtl(C_LN_NAM, lngX)).Left
7160              Case Else
                    ' ** Nothing else.
7170              End Select
7180              If blnSortHere = True Then
7190                .Sort_line.Left = .Controls(arr_varCtl(C_L1_NAM, lngX)).Left
7200                .Sort_lbl.Left = ((.Controls(arr_varCtl(C_L1_NAM, lngX)).Left + _
                      .Controls(arr_varCtl(C_L1_NAM, lngX)).Width) - lngSortLbl_Width)
7210              End If
7220            End If
7230          End If
7240          blnSortHere = False
7250        Next
7260        .IsIncome.Left = (.taxcode.Left - (8& * lngTpp))
7270        .IsDeduction.Left = (.taxcode.Left - (8& * lngTpp))
7280        Select Case blnCurrID
            Case True
7290          .IsArch_bg1.Width = (lngForm_Width - lngJournalNo_Width)
7300          .IsArch_bg2.Width = (lngForm_Width - lngJournalNo_Width)
7310          .Width = (lngForm_Width - lngJournalNo_Width)
7320        Case False
7330          .IsArch_bg1.Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
7340          .IsArch_bg2.Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
7350          .Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
7360        End Select
7370        If blnResort = True Then
7380          SortNow "Form_Load"  ' ** Procedure: Below.
7390        End If
7400      End Select

7410    End With

EXITP:
7420    Exit Sub

ERRH:
7430    Select Case ERR.Number
        Case Else
7440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7450    End Select
7460    Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnInclude As Boolean)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim blnCurrID As Boolean
        Dim lngX As Long, lngE As Long, lngF As Long

7510    With Me

7520      blnCurrID = .Parent.chkIncludeCurrency
7530      blnSortHere = False: blnResort = False

          ' **********************************************
          ' ** Array: arr_varCtl()
          ' **
          ' **   Field  Element  Name         Constant
          ' **   =====  =======  ===========  ==========
          ' **     1       0     fld name     C_CNAM
          ' **     2       1     fld left     C_LFT
          ' **     3       2     lbl1 name    C_L1_NAM
          ' **     4       3     lbl1 left    C_L1_LFT
          ' **     5       4     lbl2 name    C_L2_NAM
          ' **     6       5     lbl2 left    C_L2_LFT
          ' **     7       6     line name    C_LN_NAM
          ' **     8       7     line left    C_LN_LFT
          ' **
          ' **********************************************

          ' ** Since this only involves 2 fields, we'll do it this way.
7540      For lngX = 0& To (lngCtls - 1&)
7550        If arr_varCtl(C_CNAM, lngX) = "curr_id" Then
7560          lngE = lngX
7570        ElseIf arr_varCtl(C_CNAM, lngX) = "taxcode" Then
7580          lngF = lngX
7590        End If
7600      Next

7610      Select Case blnInclude
          Case True
7620        Select Case .Parent.chkShowJournalNo
            Case True
7630          .Width = lngForm_Width
7640          If .Sort_line.Left = .taxcode_lbl.Left Then blnSortHere = True
7650          .taxcode.Left = arr_varCtl(C_LFT, lngF)
7660          .taxcode_lbl.Left = arr_varCtl(C_L1_LFT, lngF)
7670          .taxcode_lbl2.Left = arr_varCtl(C_L2_LFT, lngF)
7680          .taxcode_lbl_line.Left = arr_varCtl(C_LN_LFT, lngF)
7690          .taxcode_description_display.Left = .taxcode.Left
7700          If blnSortHere = True Then
7710            .Sort_line.Left = .taxcode_lbl.Left
7720            .Sort_lbl.Left = ((.taxcode_lbl.Left + .taxcode_lbl.Width) - lngSortLbl_Width)
7730          End If
7740          blnSortHere = False
7750          .curr_id.Visible = True
7760          .curr_id_lbl.Visible = True
7770          .curr_id_lbl_line.Visible = True
7780          .curr_id_box.Visible = True
7790          .curr_id_forex_bg.Visible = True
7800          .IsArch_bg1.Width = lngForm_Width
7810          .IsArch_bg2.Width = lngForm_Width
7820        Case False
7830          .Width = (lngForm_Width - lngJournalNo_Width)
7840          If .Sort_line.Left = .taxcode_lbl.Left Then blnSortHere = True
7850          .taxcode.Left = (arr_varCtl(C_LFT, lngF) - lngJournalNo_Width)
7860          .taxcode_lbl.Left = (arr_varCtl(C_L1_LFT, lngF) - lngJournalNo_Width)
7870          .taxcode_lbl2.Left = (arr_varCtl(C_L2_LFT, lngF) - lngJournalNo_Width)
7880          .taxcode_lbl_line.Left = (arr_varCtl(C_LN_LFT, lngF) - lngJournalNo_Width)
7890          .taxcode_description_display.Left = .taxcode.Left
7900          If blnSortHere = True Then
7910            .Sort_line.Left = .taxcode_lbl.Left
7920            .Sort_lbl.Left = ((.taxcode_lbl.Left + .taxcode_lbl.Width) - lngSortLbl_Width)
7930          End If
7940          blnSortHere = False
7950          .curr_id.Left = (arr_varCtl(C_LFT, lngE) - lngJournalNo_Width)
7960          .curr_id_lbl.Left = (arr_varCtl(C_L1_LFT, lngE) - lngJournalNo_Width)
7970          .curr_id_lbl_line.Left = (arr_varCtl(C_LN_LFT, lngE) - lngJournalNo_Width)
7980          .curr_id_box.Left = .curr_id.Left
7990          .curr_id_forex_bg.Left = (.curr_id.Left - lngTpp)
8000          .curr_id.Visible = True
8010          .curr_id_lbl.Visible = True
8020          .curr_id_lbl_line.Visible = True
8030          .curr_id_box.Visible = True
8040          .curr_id_forex_bg.Visible = True
8050          .IsArch_bg1.Width = (lngForm_Width - lngJournalNo_Width)
8060          .IsArch_bg2.Width = (lngForm_Width - lngJournalNo_Width)
8070        End Select
8080        .IsIncome.Left = (.taxcode.Left - (8& * lngTpp))
8090        .IsDeduction.Left = (.taxcode.Left - (8& * lngTpp))
8100      Case False
8110        If .Sort_line.Left = .curr_id_lbl.Left Then blnSortHere = True
8120        .curr_id.Visible = False
8130        .curr_id_lbl.Visible = False
8140        .curr_id_lbl_line.Visible = False
8150        .curr_id_box.Visible = False
8160        .curr_id_forex_bg.Visible = False
8170        If blnSortHere = True Then
8180          .Sort_line.Visible = False
8190          .Sort_lbl.Visible = False
8200          blnResort = True
8210        End If
8220        blnSortHere = False
8230        If .Sort_line.Left = .taxcode_lbl.Left Then blnSortHere = True
8240        Select Case .Parent.chkShowJournalNo
            Case True
8250          .taxcode.Left = (arr_varCtl(C_LFT, lngF) - lngCurrID_Width)
8260          .taxcode_lbl.Left = (arr_varCtl(C_L1_LFT, lngF) - lngCurrID_Width)
8270          .taxcode_lbl2.Left = (arr_varCtl(C_L2_LFT, lngF) - lngCurrID_Width)
8280          .taxcode_lbl_line.Left = (arr_varCtl(C_LN_LFT, lngF) - lngCurrID_Width)
8290        Case False
8300          .taxcode.Left = ((arr_varCtl(C_LFT, lngF) - lngCurrID_Width) - lngJournalNo_Width)
8310          .taxcode_lbl.Left = ((arr_varCtl(C_L1_LFT, lngF) - lngCurrID_Width) - lngJournalNo_Width)
8320          .taxcode_lbl2.Left = ((arr_varCtl(C_L2_LFT, lngF) - lngCurrID_Width) - lngJournalNo_Width)
8330          .taxcode_lbl_line.Left = ((arr_varCtl(C_LN_LFT, lngF) - lngCurrID_Width) - lngJournalNo_Width)
8340        End Select
8350        .taxcode_description_display.Left = .taxcode.Left
8360        If blnSortHere = True Then
8370          .Sort_line.Left = .taxcode_lbl.Left
8380          .Sort_lbl.Left = ((.taxcode_lbl.Left + .taxcode_lbl.Width) - lngSortLbl_Width)
8390        End If
8400        If blnResort = True Then
8410          SortNow "Form_Load"  ' ** Procedure: Below.
8420        End If
8430        .IsIncome.Left = (.taxcode.Left - (8& * lngTpp))
8440        .IsDeduction.Left = (.taxcode.Left - (8& * lngTpp))
8450        Select Case .Parent.chkShowJournalNo
            Case True
8460          .IsArch_bg1.Width = (lngForm_Width - lngCurrID_Width)
8470          .IsArch_bg2.Width = (lngForm_Width - lngCurrID_Width)
8480          .Width = (lngForm_Width - lngCurrID_Width)
8490        Case False
8500          .IsArch_bg1.Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
8510          .IsArch_bg2.Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
8520          .Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
8530        End Select
8540      End Select

8550    End With

EXITP:
8560    Exit Sub

ERRH:
8570    Select Case ERR.Number
        Case Else
8580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8590    End Select
8600    Resume EXITP

End Sub

Public Sub CtlArrayLoad()

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "CtlArrayLoad"

        Dim lngE As Long

8710    With Me

8720      If lngCtls = 0& Or IsEmpty(arr_varCtl) = True Then

8730        lngCtls = 0&
8740        ReDim arr_varCtl(C_ELEMS, 0)

            ' **********************************************
            ' ** Array: arr_varCtl()
            ' **
            ' **   Field  Element  Name         Constant
            ' **   =====  =======  ===========  ==========
            ' **     1       0     fld name     C_CNAM
            ' **     2       1     fld left     C_LFT
            ' **     3       2     lbl1 name    C_L1_NAM
            ' **     4       3     lbl1 left    C_L1_LFT
            ' **     5       4     lbl2 name    C_L2_NAM
            ' **     6       5     lbl2 left    C_L2_LFT
            ' **     7       6     line name    C_LN_NAM
            ' **     8       7     line left    C_LN_LFT
            ' **
            ' **********************************************

8750        lngCtls = lngCtls + 1&
8760        lngE = lngCtls - 1&
8770        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
8780        arr_varCtl(C_CNAM, lngE) = "journalno"
8790        arr_varCtl(C_LFT, lngE) = .journalno.Left
8800        arr_varCtl(C_L1_NAM, lngE) = "journalno_lbl"
8810        arr_varCtl(C_L1_LFT, lngE) = .journalno_lbl.Left
8820        arr_varCtl(C_L2_NAM, lngE) = "journalno_lbl2"
8830        arr_varCtl(C_L2_LFT, lngE) = .journalno_lbl2.Left
8840        arr_varCtl(C_LN_NAM, lngE) = "journalno_lbl_line"
8850        arr_varCtl(C_LN_LFT, lngE) = .journalno_lbl_line.Left

8860        lngCtls = lngCtls + 1&
8870        lngE = lngCtls - 1&
8880        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
8890        arr_varCtl(C_CNAM, lngE) = "transdate"
8900        arr_varCtl(C_LFT, lngE) = .transdate.Left
8910        arr_varCtl(C_L1_NAM, lngE) = "transdate_lbl"
8920        arr_varCtl(C_L1_LFT, lngE) = .transdate_lbl.Left
8930        arr_varCtl(C_L2_NAM, lngE) = "transdate_lbl2"
8940        arr_varCtl(C_L2_LFT, lngE) = .transdate_lbl2.Left
8950        arr_varCtl(C_LN_NAM, lngE) = "transdate_lbl_line"
8960        arr_varCtl(C_LN_LFT, lngE) = .transdate_lbl_line.Left

8970        lngCtls = lngCtls + 1&
8980        lngE = lngCtls - 1&
8990        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
9000        arr_varCtl(C_CNAM, lngE) = "journaltype"
9010        arr_varCtl(C_LFT, lngE) = .journaltype.Left
9020        arr_varCtl(C_L1_NAM, lngE) = "journaltype_lbl"
9030        arr_varCtl(C_L1_LFT, lngE) = .journaltype_lbl.Left
9040        arr_varCtl(C_L2_NAM, lngE) = "journaltype_lbl2"
9050        arr_varCtl(C_L2_LFT, lngE) = .journaltype_lbl2.Left
9060        arr_varCtl(C_LN_NAM, lngE) = "journaltype_lbl_line"
9070        arr_varCtl(C_LN_LFT, lngE) = .journaltype_lbl_line.Left

9080        lngCtls = lngCtls + 1&
9090        lngE = lngCtls - 1&
9100        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
9110        arr_varCtl(C_CNAM, lngE) = "chkTaxChk"
9120        arr_varCtl(C_LFT, lngE) = .chkTaxChk.Left
9130        arr_varCtl(C_L1_NAM, lngE) = "chkTaxChk_lbl"
9140        arr_varCtl(C_L1_LFT, lngE) = .chkTaxChk_lbl.Left
9150        arr_varCtl(C_L2_NAM, lngE) = Null
9160        arr_varCtl(C_L2_LFT, lngE) = Null
9170        arr_varCtl(C_LN_NAM, lngE) = "chkTaxChk_lbl_line"
9180        arr_varCtl(C_LN_LFT, lngE) = .chkTaxChk_lbl_line.Left

9190        lngCtls = lngCtls + 1&
9200        lngE = lngCtls - 1&
9210        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
9220        arr_varCtl(C_CNAM, lngE) = "totdesc"
9230        arr_varCtl(C_LFT, lngE) = .totdesc.Left
9240        arr_varCtl(C_L1_NAM, lngE) = "totdesc_lbl"
9250        arr_varCtl(C_L1_LFT, lngE) = .totdesc_lbl.Left
9260        arr_varCtl(C_L2_NAM, lngE) = Null
9270        arr_varCtl(C_L2_LFT, lngE) = Null
9280        arr_varCtl(C_LN_NAM, lngE) = "totdesc_lbl_line"
9290        arr_varCtl(C_LN_LFT, lngE) = .totdesc_lbl_line.Left

9300        lngCtls = lngCtls + 1&
9310        lngE = lngCtls - 1&
9320        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
9330        arr_varCtl(C_CNAM, lngE) = "icash"
9340        arr_varCtl(C_LFT, lngE) = .ICash.Left
9350        arr_varCtl(C_L1_NAM, lngE) = "icash_lbl"
9360        arr_varCtl(C_L1_LFT, lngE) = .icash_lbl.Left
9370        arr_varCtl(C_L2_NAM, lngE) = "icash_lbl2"
9380        arr_varCtl(C_L2_LFT, lngE) = .icash_lbl2.Left
9390        arr_varCtl(C_LN_NAM, lngE) = "icash_lbl_line"
9400        arr_varCtl(C_LN_LFT, lngE) = .icash_lbl_line.Left

9410        lngCtls = lngCtls + 1&
9420        lngE = lngCtls - 1&
9430        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
9440        arr_varCtl(C_CNAM, lngE) = "pcash"
9450        arr_varCtl(C_LFT, lngE) = .PCash.Left
9460        arr_varCtl(C_L1_NAM, lngE) = "pcash_lbl"
9470        arr_varCtl(C_L1_LFT, lngE) = .pcash_lbl.Left
9480        arr_varCtl(C_L2_NAM, lngE) = "pcash_lbl2"
9490        arr_varCtl(C_L2_LFT, lngE) = .pcash_lbl2.Left
9500        arr_varCtl(C_LN_NAM, lngE) = "pcash_lbl_line"
9510        arr_varCtl(C_LN_LFT, lngE) = .pcash_lbl_line.Left

9520        lngCtls = lngCtls + 1&
9530        lngE = lngCtls - 1&
9540        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
9550        arr_varCtl(C_CNAM, lngE) = "cost"
9560        arr_varCtl(C_LFT, lngE) = .Cost.Left
9570        arr_varCtl(C_L1_NAM, lngE) = "cost_lbl"
9580        arr_varCtl(C_L1_LFT, lngE) = .cost_lbl.Left
9590        arr_varCtl(C_L2_NAM, lngE) = Null
9600        arr_varCtl(C_L2_LFT, lngE) = Null
9610        arr_varCtl(C_LN_NAM, lngE) = "cost_lbl_line"
9620        arr_varCtl(C_LN_LFT, lngE) = .cost_lbl_line.Left

9630        lngCtls = lngCtls + 1&
9640        lngE = lngCtls - 1&
9650        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
9660        arr_varCtl(C_CNAM, lngE) = "curr_id"
9670        arr_varCtl(C_LFT, lngE) = .curr_id.Left
9680        arr_varCtl(C_L1_NAM, lngE) = "curr_id_lbl"
9690        arr_varCtl(C_L1_LFT, lngE) = .curr_id_lbl.Left
9700        arr_varCtl(C_L2_NAM, lngE) = Null
9710        arr_varCtl(C_L2_LFT, lngE) = Null
9720        arr_varCtl(C_LN_NAM, lngE) = "curr_id_lbl_line"
9730        arr_varCtl(C_LN_LFT, lngE) = .curr_id_lbl_line.Left

9740        lngCtls = lngCtls + 1&
9750        lngE = lngCtls - 1&
9760        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
9770        arr_varCtl(C_CNAM, lngE) = "taxcode"
9780        arr_varCtl(C_LFT, lngE) = .taxcode.Left
9790        arr_varCtl(C_L1_NAM, lngE) = "taxcode_lbl"
9800        arr_varCtl(C_L1_LFT, lngE) = .taxcode_lbl.Left
9810        arr_varCtl(C_L2_NAM, lngE) = "taxcode_lbl2"
9820        arr_varCtl(C_L2_LFT, lngE) = .taxcode_lbl2.Left
9830        arr_varCtl(C_LN_NAM, lngE) = "taxcode_lbl_line"
9840        arr_varCtl(C_LN_LFT, lngE) = .taxcode_lbl_line.Left

9850      End If

9860    End With

EXITP:
9870    Exit Sub

ERRH:
9880    Select Case ERR.Number
        Case Else
9890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9900    End Select
9910    Resume EXITP

End Sub

Public Function RecCnt() As Long

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

10010   lngRetVal = 0&

10020   With Me
10030     Set rst = .RecordsetClone
10040     With rst
10050       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
10060       Else
10070         .MoveLast
10080         lngRetVal = .RecordCount
10090       End If
10100       .Close
10110     End With
10120   End With

EXITP:
10130   Set rst = Nothing
10140   RecCnt = lngRetVal
10150   Exit Function

ERRH:
10160   lngRetVal = 0&
10170   Select Case ERR.Number
        Case Else
10180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10190   End Select
10200   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

10310   Select Case intWhere
        Case acCmdRecordsGoToFirst
10320     DoCmd.RunCommand acCmdRecordsGoToFirst
10330   Case acCmdRecordsGoToLast
10340     DoCmd.RunCommand acCmdRecordsGoToLast
10350   Case acCmdRecordsGoToPrevious
10360     DoCmd.RunCommand acCmdRecordsGoToPrevious
10370   Case acCmdRecordsGoToNext
10380     DoCmd.RunCommand acCmdRecordsGoToNext
10390   Case Else
10400     If IsMissing(varID) = False Then
10410       With Me
10420         Set rst = .RecordsetClone
10430         With rst
10440           .FindFirst "[tcbas_id] = '" & varID & "'"
10450           If .NoMatch = False Then
10460             Me.Bookmark = .Bookmark
10470           End If
10480           .Close
10490         End With
10500       End With
10510     End If
10520   End Select

EXITP:
10530   Set rst = Nothing
10540   Exit Sub

ERRH:
10550   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
10560   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
10570   Case Else
10580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10590   End Select
10600   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [transdate], [JournalType_Order], [totdesc]"

10710   With Me
10720     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
10730       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
10740     End If
10750     .Controls(strSortLbl).Visible = False
10760     .Controls(strSortLine).Visible = False
10770     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
10780     If strProc = "Form_Load" Then
10790       strCalled = "transdate"
10800       strSortNow = strSortOrig
10810       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
10820       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
10830       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
10840       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
10850       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
10860       .Controls(strSortLbl).Top = lngSortLbl_Top
10870       .Controls(strSortLbl).Left = lngSortLbl_Left
10880       .Controls(strSortLine).Top = lngSortLine_Top
10890       .Controls(strSortLine).Left = lngSortLine_Left
10900       .Controls(strSortLine).Width = lngSortLine_Width
10910       .Controls(strSortLbl).Caption = strArwUp
10920       .Controls(strSortLbl).ForeColor = CLR_DKBLU
10930     Else
10940       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
10950       Select Case strCalled
            Case "journalno", "curr_id"
10960         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
10970         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
10980       Case Else
10990         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
11000         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
11010       End Select
11020       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
11030       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
11040       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
11050       .Controls(strSortLbl).Top = lngSortLbl_Top
11060       .Controls(strSortLbl).Left = lngSortLbl_Left
11070       .Controls(strSortLine).Top = lngSortLine_Top
11080       .Controls(strSortLine).Left = lngSortLine_Left
11090       .Controls(strSortLine).Width = lngSortLine_Width
11100       If strCalled = "transdate" Then
11110         strSortAsc = strSortOrig
11120         If strSortNow = strSortAsc Then
11130           strSortNow = "[transdate] DESC, [JType_sortOrder] DESC, [totdesc] DESC"
11140           .Controls(strSortLbl).Caption = strArwDn
11150           .Controls(strSortLbl).ForeColor = CLR_DKRED
11160         Else
11170           strSortNow = strSortAsc
11180           .Controls(strSortLbl).Caption = strArwUp
11190           .Controls(strSortLbl).ForeColor = CLR_DKBLU
11200         End If
11210       Else
11220         .Controls(strSortLbl).Caption = strArwUp
11230         .Controls(strSortLbl).ForeColor = CLR_DKBLU
11240         Select Case strCalled
              Case "journalno"
11250           strSortAsc = "[journalno]"
11260           lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
11270           .Controls(strSortLbl).Left = lngSortLbl_Left
11280         Case "journaltype"
11290           strSortAsc = "[journaltypex], [transdate], [totdesc]"
11300           lngSortLbl_Left = (lngSortLbl_Left + lngTpp)
11310           .Controls(strSortLbl).Left = lngSortLbl_Left
11320         Case "chkTaxChk"
11330           strSortAsc = "[chkTaxChk]" & strStdAsc
11340           lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
11350           .Controls(strSortLbl).Left = lngSortLbl_Left
11360         Case "totdesc"
11370           strSortAsc = "[totdesc], [transdate], [JournalType_Order]"
11380         Case "icash"
11390           strSortAsc = "[icash]" & strStdAsc
11400         Case "pcash"
11410           strSortAsc = "[pcash]" & strStdAsc
11420         Case "cost"
11430           strSortAsc = "[cost]" & strStdAsc
11440         Case "curr_id"
11450           strSortAsc = "[curr_code]" & strStdAsc
11460           lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
11470           .Controls(strSortLbl).Left = lngSortLbl_Left
11480         Case "taxcode"
11490           strSortAsc = "[taxcode_type], [taxcode_description]" & strStdAsc
11500         End Select
11510         If strSortNow = strSortAsc Then
11520           intCnt = CharCnt(strSortAsc, ",") + 1
11530           Select Case intCnt
                Case 1
11540             strTmp01 = strSortAsc & " DESC"
11550           Case 2
11560             intPos01 = InStr(strSortAsc, ",")
11570             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
11580             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
11590           Case 3
11600             intPos01 = InStr(strSortAsc, ",")
11610             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
11620             strTmp02 = Mid(strSortAsc, intPos01)
11630             intPos01 = InStr(2, strTmp02, ",")
11640             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11650             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
11660           Case 4
11670             intPos01 = InStr(strSortAsc, ",")
11680             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
11690             strTmp02 = Mid(strSortAsc, intPos01)
11700             intPos01 = InStr(2, strTmp02, ",")
11710             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11720             strTmp02 = Mid(strSortAsc, intPos01)
11730             intPos01 = InStr(2, strTmp02, ",")
11740             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11750             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
11760           Case 5
11770             intPos01 = InStr(strSortAsc, ",")
11780             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
11790             strTmp02 = Mid(strSortAsc, intPos01)
11800             intPos01 = InStr(2, strTmp02, ",")
11810             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11820             strTmp02 = Mid(strSortAsc, intPos01)
11830             intPos01 = InStr(2, strTmp02, ",")
11840             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11850             strTmp02 = Mid(strSortAsc, intPos01)
11860             intPos01 = InStr(2, strTmp02, ",")
11870             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11880             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
11890           End Select
11900           strSortNow = strTmp01
11910           .Controls(strSortLbl).Caption = strArwDn
11920           .Controls(strSortLbl).ForeColor = CLR_DKRED
11930         Else
11940           strSortNow = strSortAsc
11950         End If
11960       End If
11970     End If
11980     .Controls(strSortLbl).Visible = True
11990     .Controls(strSortLine).Visible = True
12000     .OrderBy = strSortNow
12010     .OrderByOn = True
12020 On Error Resume Next
12030     .FocusHolder.SetFocus
12040 On Error GoTo ERRH
12050   End With

        ' ** Potential Sorts:

EXITP:
12060   Exit Sub

ERRH:
12070   Select Case ERR.Number
        Case Else
12080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12090   End Select
12100   Resume EXITP

End Sub
