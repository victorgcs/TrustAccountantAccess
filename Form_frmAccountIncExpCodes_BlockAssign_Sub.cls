VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountIncExpCodes_BlockAssign_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountIncExpCodes_BlockAssign_Sub"

'VGC 10/02/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Assign:           A {cmdAssign on frmAccountIncExpCodes_BlockAssign}
' **   Cancel:           C {cmdCancel on frmAccountIncExpCodes_BlockAssign}
' **   Rev Code:         D {cmbRevenueCodes on frmAccountIncExpCodes_BlockAssign}}
' **   Clear:            E {cmdClear on frmAccountIncExpCodes_BlockAssign}
' **   Show JournalNo:   H {chkShowJournalNo on frmAccountIncExpCodes_BlockAssign}
' **   Filter by JType:  J {cmbJournalTypes on frmAccountIncExpCodes_BlockAssign}
' **   Checked Only:     K {chkShowCheckedOnly on frmAccountIncExpCodes_BlockAssign}
' **   Select None:      N {cmdSelectNone on frmAccountIncExpCodes_BlockAssign}
' **   OK:               O {cmdOK on frmAccountIncExpCodes_BlockAssign}
' **   Select All:       S {cmdSelectAll on frmAccountIncExpCodes_BlockAssign}
' **   Include Currency: U {chkIncludeCurrency on frmAccountIncExpCodes_BlockAssign}
' **   Unassigned Only:  W {chkShowUnassignedOnly on frmAccountIncExpCodes_BlockAssign}
' **   Check This:       X {chkIncExpChk}

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
150       If strCallingForm = "frmAccountIncExpCodes_BlockAssign" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

180         blnRecalcTots = False

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

200         .form_width_line.Left = 0&  ' ** Just get this out of the way.

210         lngSortLbl_Width = .Sort_lbl.Width
220         lngForm_Width = .Width
230         lngJournalNo_Width = (.transdate.Left - .journalno.Left)  ' ** The amount of space it uses.
240         lngCurrID_Width = ((.revcode_ID.Left - .curr_id.Left) - (4& * lngTpp))
250         .Parent.SetJNoWidth lngJournalNo_Width  ' ** Form Procedure: frmAccountIncExpCodes_BlockAssign.
260         .Parent.SetCurrWidth lngCurrID_Width  ' ** Form Procedure: frmAccountIncExpCodes_BlockAssign.

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
1240      .revcode_ID.ForeColor = CLR_DISABLED_FG
1250      .revcode_ID.BackColor = CLR_DISABLED_BG

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
        ' **   Rev Code:         D {cmbRevenueCodes on frmAccountIncExpCodes_BlockAssign}}
        ' **   Filter by JType:  J {cmbJournalTypes on frmAccountIncExpCodes_BlockAssign}
        ' **   Check This:       X {chkIncExpChk}

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
1580          .Parent.cmdCancel_Click  ' ** Form Procedure: frmAccountIncExpCodes_BlockAssign.
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
1670          .Parent.cmbRevenueCodes.SetFocus
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
1780          .chkIncExpChk.SetFocus
1790          .chkIncExpChk = (Not .chkIncExpChk)
1800          chkIncExpChk_AfterUpdate  ' ** Procedure: Below.
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
2130          .Parent.cmbRevenueCodes.SetFocus
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

Private Sub chkIncExpChk_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncExpChk_lbl_DblClick"

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

Public Sub chkIncExpChk_AfterUpdate()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncExpChk_AfterUpdate"

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

Private Sub chkIncExpChk_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncExpChk_KeyDown"

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
3520            .chkIncExpChk.SetFocus
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
3660            .chkIncExpChk.SetFocus
3670          Else
3680            DoCmd.SelectObject acForm, .Parent.Name, False
3690            .Parent.cmbRevenueCodes.SetFocus
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

Private Sub revcode_ID_lbl_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_lbl_DblClick"

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

Public Sub RecalcTots()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRevcodeID As Long
        Dim varTmp00 As Variant, lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long, lngTmp05 As Long

4510    With Me
4520      If gblnSignal = False Then
4530        gblnSignal = True  ' ** Trying to prevent too many calls from piling up.

4540        lngRecsCur = RecCnt  ' ** Function: Below.
4550        If lngRecsCur > 0& Then

4560          .chkIncExpChk_sum.Requery
4570          Select Case IsNull(.Parent.cmbRevenueCodes)
              Case True
4580            lngRevcodeID = 0&
4590          Case False
4600            lngRevcodeID = .Parent.cmbRevenueCodes.Column(0)
4610          End Select

4620          varTmp00 = .chkIncExpChk_sum
4630          If IsNull(varTmp00) = False Then
4640            lngTmp01 = Abs(varTmp00)
4650            .Parent.chkIncExpChk_sum = lngTmp01
4660          Else
4670            lngTmp01 = 0&
4680            .Parent.chkIncExpChk_sum = lngTmp01
4690          End If

4700          varTmp00 = .rcbas_changed_sum
4710          If IsNull(varTmp00) = False Then
4720            lngTmp02 = Abs(varTmp00)
4730          Else
4740            lngTmp02 = 0&
4750          End If

4760          Set dbs = CurrentDb
              ' ** tblRevCode_BlockAssign_Staging, grouped and summed,
              ' ** with chkIncExpChk_sum, rcbas_changed_sum, cnt_recs.
4770          Set qdf = dbs.QueryDefs("qryAccountProfile_IncExpCodes_Block_07")
4780          Set rst = qdf.OpenRecordset
4790          If rst.BOF = True And rst.EOF = True Then
                ' ** Well, who knows?
4800            lngTmp03 = 0&: lngTmp04 = 0&: lngTmp05 = 0&
4810          Else
                ' ** Remember, this recordset is the entire table, and this subform may be filtered.
                ' ** Also, sums will be negative, since they're on boolean fields.
4820            rst.MoveFirst
4830            lngTmp03 = Nz(Abs(rst![chkIncExpChk_sum]), 0)   ' ** Total checks in entire group.
4840            lngTmp04 = Nz(Abs(rst![rcbas_changed_sum]), 0)  ' ** Total changed in entire group.
4850            lngTmp05 = rst![cnt_recs]                       ' ** Total records in entire group.
4860            .Parent.rcbas_changed_sum = lngTmp04
4870          End If
4880          rst.Close
4890          dbs.Close

4900        Else
4910          lngTmp01 = 0&: lngTmp03 = 0&: lngTmp04 = 0&: lngTmp05 = 0&
4920        End If  ' ** lngRecsCur.

            ' ** lngTmp01  : Checked items in this subform.
            ' ** lngTmp02  : Changed items in this subform.
            ' ** lngRecsCur: Records in this subform.
            ' ** lngTmp03  : Checked items in entire group.
            ' ** lngTmp04  : Changed items in entire group.
            ' ** lngTmp05  : Records in entire group.
4930        If lngRecsCur = lngTmp05 And lngRecsCur > 0& Then
              ' ** Subform not filtered.
4940          If lngTmp03 > 0& Then
4950            .Parent.cmdClear.Enabled = True
4960          Else
4970            .Parent.cmdClear.Enabled = False
4980          End If
4990          If lngRevcodeID > 0& And lngTmp03 > 0& Then
5000            .Parent.cmdAssign.Enabled = True
5010          Else
5020            .Parent.cmdAssign.Enabled = False
5030          End If
5040          If lngTmp04 > 0& Then
5050            .Parent.cmdOK.Enabled = True
5060          Else
5070            .Parent.cmdOK.Enabled = False
5080          End If
5090        ElseIf lngRecsCur > 0& Then
              ' ** Subform filtered.
5100          If lngTmp01 > 0& Then
5110            .Parent.cmdClear.Enabled = True
5120          Else
5130            .Parent.cmdClear.Enabled = False
5140          End If
5150          If lngRevcodeID > 0& And lngTmp01 > 0& Then
5160            .Parent.cmdAssign.Enabled = True
5170          Else
5180            .Parent.cmdAssign.Enabled = False
5190          End If
5200          If lngTmp02 > 0& Then
5210            .Parent.cmdOK.Enabled = True
5220          Else
5230            .Parent.cmdOK.Enabled = False
5240          End If
5250        Else
5260          .Parent.cmdClear.Enabled = False
5270          .Parent.cmdAssign.Enabled = False
5280          .Parent.cmdOK.Enabled = False
5290        End If  ' ** lngRecsCur.

5300        gblnSignal = False
5310      End If
5320    End With

EXITP:
5330    Set rst = Nothing
5340    Set qdf = Nothing
5350    Set dbs = Nothing
5360    Exit Sub

ERRH:
5370    gblnSignal = False
5380    Select Case ERR.Number
        Case Else
5390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5400    End Select
5410    Resume EXITP

End Sub

Public Sub ShowJournalNo_Sub(blnShow As Boolean)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowJournalNo_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim blnCurrID As Boolean
        Dim lngTmp01 As Long
        Dim lngX As Long

5510    With Me

5520      blnCurrID = .Parent.chkIncludeCurrency
5530      blnSortHere = False: blnResort = False

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

5540      Select Case blnShow
          Case True
5550        Select Case blnCurrID
            Case True
5560          .Width = lngForm_Width
5570        Case False
5580          .Width = (lngForm_Width - lngCurrID_Width)
5590        End Select
5600        For lngX = (lngCtls - 1&) To 0& Step -1&
5610          If arr_varCtl(C_CNAM, lngX) = "journalno" Then
5620            .journalno.Visible = True
5630            .journalno_lbl.Visible = True
5640            .journalno_lbl2.Visible = True
5650            .journalno_lbl_line.Visible = True
5660          Else
5670            If blnCurrID = False And arr_varCtl(C_CNAM, lngX) = "curr_id" Then
                  ' ** Skip it.
5680            Else
5690              If .Sort_line.Left = .Controls(arr_varCtl(C_L1_NAM, lngX)).Left Then blnSortHere = True
5700              If blnCurrID = False And arr_varCtl(C_CNAM, lngX) = "revcode_ID" Then
5710                lngTmp01 = lngCurrID_Width
5720              Else
5730                lngTmp01 = 0&
5740              End If
5750              .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) - lngTmp01)
5760              .Controls(arr_varCtl(C_L1_NAM, lngX)).Left = (arr_varCtl(C_L1_LFT, lngX) - lngTmp01)
5770              If IsNull(arr_varCtl(C_L2_NAM, lngX)) = False Then
5780                .Controls(arr_varCtl(C_L2_NAM, lngX)).Left = (arr_varCtl(C_L2_LFT, lngX) - lngTmp01)
5790              End If
5800              .Controls(arr_varCtl(C_LN_NAM, lngX)).Left = (arr_varCtl(C_LN_LFT, lngX) - lngTmp01)
5810              Select Case arr_varCtl(C_CNAM, lngX)
                  Case "icash", "pcash", "cost"
5820                .Controls(arr_varCtl(C_CNAM, lngX) & "_str").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
5830                .Controls(arr_varCtl(C_CNAM, lngX) & "_box").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
5840                .Controls(arr_varCtl(C_CNAM, lngX) & "_forex_bg").Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTpp)
5850              Case "curr_id"
5860                .Controls(arr_varCtl(C_CNAM, lngX) & "_box").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
5870                .Controls(arr_varCtl(C_CNAM, lngX) & "_forex_bg").Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTpp)
5880              Case Else
                    ' ** Nothing else.
5890              End Select
5900              If blnSortHere = True Then
5910                .Sort_line.Left = .Controls(arr_varCtl(C_L1_NAM, lngX)).Left
5920                .Sort_lbl.Left = ((.Controls(arr_varCtl(C_L1_NAM, lngX)).Left + _
                      .Controls(arr_varCtl(C_L1_NAM, lngX)).Width) - lngSortLbl_Width)
5930              End If
5940            End If
5950          End If
5960          blnSortHere = False
5970        Next
5980        Select Case blnCurrID
            Case True
5990          .IsArch_bg1.Width = lngForm_Width
6000          .IsArch_bg2.Width = lngForm_Width
6010        Case False
6020          .IsArch_bg1.Width = (lngForm_Width - lngCurrID_Width)
6030          .IsArch_bg2.Width = (lngForm_Width - lngCurrID_Width)
6040        End Select
6050        .IsIncome.Left = (.revcode_ID.Left - (8& * lngTpp))
6060        .IsExpense.Left = (.revcode_ID.Left - (8& * lngTpp))
6070      Case False
6080        For lngX = 0& To (lngCtls - 1&)
6090          If arr_varCtl(C_CNAM, lngX) = "journalno" Then
6100            If .Sort_line.Left = .journalno_lbl.Left Then blnSortHere = True
6110            .journalno.Visible = False
6120            .journalno_lbl.Visible = False
6130            .journalno_lbl2.Visible = False
6140            .journalno_lbl_line.Visible = False
6150            If blnSortHere = True Then
6160              .Sort_line.Visible = False
6170              .Sort_lbl.Visible = False
6180              blnResort = True
6190            End If
6200          Else
6210            If blnCurrID = False And arr_varCtl(C_CNAM, lngX) = "curr_id" Then
                  ' ** Skip it.
6220            Else
6230              If .Sort_line.Left = .Controls(arr_varCtl(C_L1_NAM, lngX)).Left Then blnSortHere = True
6240              If blnCurrID = False And arr_varCtl(C_CNAM, lngX) = "revcode_ID" Then
6250                lngTmp01 = lngCurrID_Width
6260              Else
6270                lngTmp01 = 0&
6280              End If
6290              .Controls(arr_varCtl(C_CNAM, lngX)).Left = ((arr_varCtl(C_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
6300              .Controls(arr_varCtl(C_L1_NAM, lngX)).Left = ((arr_varCtl(C_L1_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
6310              If IsNull(arr_varCtl(C_L2_NAM, lngX)) = False Then
6320                .Controls(arr_varCtl(C_L2_NAM, lngX)).Left = ((arr_varCtl(C_L2_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
6330              End If
6340              .Controls(arr_varCtl(C_LN_NAM, lngX)).Left = ((arr_varCtl(C_LN_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
6350              Select Case arr_varCtl(C_CNAM, lngX)
                  Case "icash", "pcash", "cost"
6360                .Controls(arr_varCtl(C_CNAM, lngX) & "_str").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
6370                .Controls(arr_varCtl(C_CNAM, lngX) & "_box").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
6380                .Controls(arr_varCtl(C_CNAM, lngX) & "_forex_bg").Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTpp)
6390              Case "curr_id"
6400                .Controls(arr_varCtl(C_CNAM, lngX) & "_box").Left = .Controls(arr_varCtl(C_CNAM, lngX)).Left
6410                .Controls(arr_varCtl(C_CNAM, lngX) & "_forex_bg").Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTpp)
6420              Case Else
                    ' ** Nothing else.
6430              End Select
6440              If blnSortHere = True Then
6450                .Sort_line.Left = .Controls(arr_varCtl(C_L1_NAM, lngX)).Left
6460                .Sort_lbl.Left = ((.Controls(arr_varCtl(C_L1_NAM, lngX)).Left + _
                      .Controls(arr_varCtl(C_L1_NAM, lngX)).Width) - lngSortLbl_Width)
6470              End If
6480            End If
6490          End If
6500          blnSortHere = False
6510        Next
6520        .IsIncome.Left = (.revcode_ID.Left - (8& * lngTpp))
6530        .IsExpense.Left = (.revcode_ID.Left - (8& * lngTpp))
6540        Select Case blnCurrID
            Case True
6550          .IsArch_bg1.Width = (lngForm_Width - lngJournalNo_Width)
6560          .IsArch_bg2.Width = (lngForm_Width - lngJournalNo_Width)
6570          .Width = (lngForm_Width - lngJournalNo_Width)
6580        Case False
6590          .IsArch_bg1.Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
6600          .IsArch_bg2.Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
6610          .Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
6620        End Select
6630        If blnResort = True Then
6640          SortNow "Form_Load"  ' ** Procedure: Below.
6650        End If
6660      End Select

6670    End With

EXITP:
6680    Exit Sub

ERRH:
6690    Select Case ERR.Number
        Case Else
6700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6710    End Select
6720    Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnInclude As Boolean)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim blnCurrID As Boolean
        Dim lngX As Long, lngE As Long, lngF As Long

6810    With Me

6820      blnCurrID = .Parent.chkIncludeCurrency
6830      blnSortHere = False: blnResort = False

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
6840      For lngX = 0& To (lngCtls - 1&)
6850        If arr_varCtl(C_CNAM, lngX) = "curr_id" Then
6860          lngE = lngX
6870        ElseIf arr_varCtl(C_CNAM, lngX) = "revcode_ID" Then
6880          lngF = lngX
6890        End If
6900      Next

6910      Select Case blnInclude
          Case True
6920        Select Case .Parent.chkShowJournalNo
            Case True
6930          .Width = lngForm_Width
6940          If .Sort_line.Left = .revcode_ID_lbl.Left Then blnSortHere = True
6950          .revcode_ID.Left = arr_varCtl(C_LFT, lngF)
6960          .revcode_ID_lbl.Left = arr_varCtl(C_L1_LFT, lngF)
6970          .revcode_ID_lbl2.Left = arr_varCtl(C_L2_LFT, lngF)
6980          .revcode_ID_lbl_line.Left = arr_varCtl(C_LN_LFT, lngF)
6990          If blnSortHere = True Then
7000            .Sort_line.Left = .revcode_ID_lbl.Left
7010            .Sort_lbl.Left = ((.revcode_ID_lbl.Left + .revcode_ID_lbl.Width) - lngSortLbl_Width)
7020          End If
7030          blnSortHere = False
7040          .curr_id.Visible = True
7050          .curr_id_lbl.Visible = True
7060          .curr_id_lbl_line.Visible = True
7070          .curr_id_box.Visible = True
7080          .curr_id_forex_bg.Visible = True
7090          .IsArch_bg1.Width = lngForm_Width
7100          .IsArch_bg2.Width = lngForm_Width
7110        Case False
7120          .Width = (lngForm_Width - lngJournalNo_Width)
7130          If .Sort_line.Left = .revcode_ID_lbl.Left Then blnSortHere = True
7140          .revcode_ID.Left = (arr_varCtl(C_LFT, lngF) - lngJournalNo_Width)
7150          .revcode_ID_lbl.Left = (arr_varCtl(C_L1_LFT, lngF) - lngJournalNo_Width)
7160          .revcode_ID_lbl2.Left = (arr_varCtl(C_L2_LFT, lngF) - lngJournalNo_Width)
7170          .revcode_ID_lbl_line.Left = (arr_varCtl(C_LN_LFT, lngF) - lngJournalNo_Width)
7180          If blnSortHere = True Then
7190            .Sort_line.Left = .revcode_ID_lbl.Left
7200            .Sort_lbl.Left = ((.revcode_ID_lbl.Left + .revcode_ID_lbl.Width) - lngSortLbl_Width)
7210          End If
7220          blnSortHere = False
7230          .curr_id.Left = (arr_varCtl(C_LFT, lngE) - lngJournalNo_Width)
7240          .curr_id_lbl.Left = (arr_varCtl(C_L1_LFT, lngE) - lngJournalNo_Width)
7250          .curr_id_lbl_line.Left = (arr_varCtl(C_LN_LFT, lngE) - lngJournalNo_Width)
7260          .curr_id_box.Left = .curr_id.Left
7270          .curr_id_forex_bg.Left = (.curr_id.Left - lngTpp)
7280          .curr_id.Visible = True
7290          .curr_id_lbl.Visible = True
7300          .curr_id_lbl_line.Visible = True
7310          .curr_id_box.Visible = True
7320          .curr_id_forex_bg.Visible = True
7330          .IsArch_bg1.Width = (lngForm_Width - lngJournalNo_Width)
7340          .IsArch_bg2.Width = (lngForm_Width - lngJournalNo_Width)
7350        End Select
7360        .IsIncome.Left = (.revcode_ID.Left - (8& * lngTpp))
7370        .IsExpense.Left = (.revcode_ID.Left - (8& * lngTpp))
7380      Case False
7390        If .Sort_line.Left = .curr_id_lbl.Left Then blnSortHere = True
7400        .curr_id.Visible = False
7410        .curr_id_lbl.Visible = False
7420        .curr_id_lbl_line.Visible = False
7430        .curr_id_box.Visible = False
7440        .curr_id_forex_bg.Visible = False
7450        If blnSortHere = True Then
7460          .Sort_line.Visible = False
7470          .Sort_lbl.Visible = False
7480          blnResort = True
7490        End If
7500        blnSortHere = False
7510        If .Sort_line.Left = .revcode_ID_lbl.Left Then blnSortHere = True
7520        Select Case .Parent.chkShowJournalNo
            Case True
7530          .revcode_ID.Left = (arr_varCtl(C_LFT, lngF) - lngCurrID_Width)
7540          .revcode_ID_lbl.Left = (arr_varCtl(C_L1_LFT, lngF) - lngCurrID_Width)
7550          .revcode_ID_lbl2.Left = (arr_varCtl(C_L2_LFT, lngF) - lngCurrID_Width)
7560          .revcode_ID_lbl_line.Left = (arr_varCtl(C_LN_LFT, lngF) - lngCurrID_Width)
7570        Case False
7580          .revcode_ID.Left = ((arr_varCtl(C_LFT, lngF) - lngCurrID_Width) - lngJournalNo_Width)
7590          .revcode_ID_lbl.Left = ((arr_varCtl(C_L1_LFT, lngF) - lngCurrID_Width) - lngJournalNo_Width)
7600          .revcode_ID_lbl2.Left = ((arr_varCtl(C_L2_LFT, lngF) - lngCurrID_Width) - lngJournalNo_Width)
7610          .revcode_ID_lbl_line.Left = ((arr_varCtl(C_LN_LFT, lngF) - lngCurrID_Width) - lngJournalNo_Width)
7620        End Select
7630        If blnSortHere = True Then
7640          .Sort_line.Left = .revcode_ID_lbl.Left
7650          .Sort_lbl.Left = ((.revcode_ID_lbl.Left + .revcode_ID_lbl.Width) - lngSortLbl_Width)
7660        End If
7670        If blnResort = True Then
7680          SortNow "Form_Load"  ' ** Procedure: Below.
7690        End If
7700        .IsIncome.Left = (.revcode_ID.Left - (8& * lngTpp))
7710        .IsExpense.Left = (.revcode_ID.Left - (8& * lngTpp))
7720        Select Case .Parent.chkShowJournalNo
            Case True
7730          .IsArch_bg1.Width = (lngForm_Width - lngCurrID_Width)
7740          .IsArch_bg2.Width = (lngForm_Width - lngCurrID_Width)
7750          .Width = (lngForm_Width - lngCurrID_Width)
7760        Case False
7770          .IsArch_bg1.Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
7780          .IsArch_bg2.Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
7790          .Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
7800        End Select
7810      End Select

7820    End With

EXITP:
7830    Exit Sub

ERRH:
7840    Select Case ERR.Number
        Case Else
7850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7860    End Select
7870    Resume EXITP

End Sub

Public Sub CtlArrayLoad()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "CtlArrayLoad"

        Dim lngE As Long

7910    With Me

7920      If lngCtls = 0& Or IsEmpty(arr_varCtl) = True Then

7930        lngCtls = 0&
7940        ReDim arr_varCtl(C_ELEMS, 0)

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

7950        lngCtls = lngCtls + 1&
7960        lngE = lngCtls - 1&
7970        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
7980        arr_varCtl(C_CNAM, lngE) = "journalno"
7990        arr_varCtl(C_LFT, lngE) = .journalno.Left
8000        arr_varCtl(C_L1_NAM, lngE) = "journalno_lbl"
8010        arr_varCtl(C_L1_LFT, lngE) = .journalno_lbl.Left
8020        arr_varCtl(C_L2_NAM, lngE) = "journalno_lbl2"
8030        arr_varCtl(C_L2_LFT, lngE) = .journalno_lbl2.Left
8040        arr_varCtl(C_LN_NAM, lngE) = "journalno_lbl_line"
8050        arr_varCtl(C_LN_LFT, lngE) = .journalno_lbl_line.Left

8060        lngCtls = lngCtls + 1&
8070        lngE = lngCtls - 1&
8080        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
8090        arr_varCtl(C_CNAM, lngE) = "transdate"
8100        arr_varCtl(C_LFT, lngE) = .transdate.Left
8110        arr_varCtl(C_L1_NAM, lngE) = "transdate_lbl"
8120        arr_varCtl(C_L1_LFT, lngE) = .transdate_lbl.Left
8130        arr_varCtl(C_L2_NAM, lngE) = "transdate_lbl2"
8140        arr_varCtl(C_L2_LFT, lngE) = .transdate_lbl2.Left
8150        arr_varCtl(C_LN_NAM, lngE) = "transdate_lbl_line"
8160        arr_varCtl(C_LN_LFT, lngE) = .transdate_lbl_line.Left

8170        lngCtls = lngCtls + 1&
8180        lngE = lngCtls - 1&
8190        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
8200        arr_varCtl(C_CNAM, lngE) = "journaltype"
8210        arr_varCtl(C_LFT, lngE) = .journaltype.Left
8220        arr_varCtl(C_L1_NAM, lngE) = "journaltype_lbl"
8230        arr_varCtl(C_L1_LFT, lngE) = .journaltype_lbl.Left
8240        arr_varCtl(C_L2_NAM, lngE) = "journaltype_lbl2"
8250        arr_varCtl(C_L2_LFT, lngE) = .journaltype_lbl2.Left
8260        arr_varCtl(C_LN_NAM, lngE) = "journaltype_lbl_line"
8270        arr_varCtl(C_LN_LFT, lngE) = .journaltype_lbl_line.Left

8280        lngCtls = lngCtls + 1&
8290        lngE = lngCtls - 1&
8300        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
8310        arr_varCtl(C_CNAM, lngE) = "chkIncExpChk"
8320        arr_varCtl(C_LFT, lngE) = .chkIncExpChk.Left
8330        arr_varCtl(C_L1_NAM, lngE) = "chkIncExpChk_lbl"
8340        arr_varCtl(C_L1_LFT, lngE) = .chkIncExpChk_lbl.Left
8350        arr_varCtl(C_L2_NAM, lngE) = Null
8360        arr_varCtl(C_L2_LFT, lngE) = Null
8370        arr_varCtl(C_LN_NAM, lngE) = "chkIncExpChk_lbl_line"
8380        arr_varCtl(C_LN_LFT, lngE) = .chkIncExpChk_lbl_line.Left

8390        lngCtls = lngCtls + 1&
8400        lngE = lngCtls - 1&
8410        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
8420        arr_varCtl(C_CNAM, lngE) = "totdesc"
8430        arr_varCtl(C_LFT, lngE) = .totdesc.Left
8440        arr_varCtl(C_L1_NAM, lngE) = "totdesc_lbl"
8450        arr_varCtl(C_L1_LFT, lngE) = .totdesc_lbl.Left
8460        arr_varCtl(C_L2_NAM, lngE) = Null
8470        arr_varCtl(C_L2_LFT, lngE) = Null
8480        arr_varCtl(C_LN_NAM, lngE) = "totdesc_lbl_line"
8490        arr_varCtl(C_LN_LFT, lngE) = .totdesc_lbl_line.Left

8500        lngCtls = lngCtls + 1&
8510        lngE = lngCtls - 1&
8520        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
8530        arr_varCtl(C_CNAM, lngE) = "icash"
8540        arr_varCtl(C_LFT, lngE) = .ICash.Left
8550        arr_varCtl(C_L1_NAM, lngE) = "icash_lbl"
8560        arr_varCtl(C_L1_LFT, lngE) = .icash_lbl.Left
8570        arr_varCtl(C_L2_NAM, lngE) = "icash_lbl2"
8580        arr_varCtl(C_L2_LFT, lngE) = .icash_lbl2.Left
8590        arr_varCtl(C_LN_NAM, lngE) = "icash_lbl_line"
8600        arr_varCtl(C_LN_LFT, lngE) = .icash_lbl_line.Left

8610        lngCtls = lngCtls + 1&
8620        lngE = lngCtls - 1&
8630        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
8640        arr_varCtl(C_CNAM, lngE) = "pcash"
8650        arr_varCtl(C_LFT, lngE) = .PCash.Left
8660        arr_varCtl(C_L1_NAM, lngE) = "pcash_lbl"
8670        arr_varCtl(C_L1_LFT, lngE) = .pcash_lbl.Left
8680        arr_varCtl(C_L2_NAM, lngE) = "pcash_lbl2"
8690        arr_varCtl(C_L2_LFT, lngE) = .pcash_lbl2.Left
8700        arr_varCtl(C_LN_NAM, lngE) = "pcash_lbl_line"
8710        arr_varCtl(C_LN_LFT, lngE) = .pcash_lbl_line.Left

8720        lngCtls = lngCtls + 1&
8730        lngE = lngCtls - 1&
8740        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
8750        arr_varCtl(C_CNAM, lngE) = "cost"
8760        arr_varCtl(C_LFT, lngE) = .Cost.Left
8770        arr_varCtl(C_L1_NAM, lngE) = "cost_lbl"
8780        arr_varCtl(C_L1_LFT, lngE) = .cost_lbl.Left
8790        arr_varCtl(C_L2_NAM, lngE) = Null
8800        arr_varCtl(C_L2_LFT, lngE) = Null
8810        arr_varCtl(C_LN_NAM, lngE) = "cost_lbl_line"
8820        arr_varCtl(C_LN_LFT, lngE) = .cost_lbl_line.Left

8830        lngCtls = lngCtls + 1&
8840        lngE = lngCtls - 1&
8850        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
8860        arr_varCtl(C_CNAM, lngE) = "curr_id"
8870        arr_varCtl(C_LFT, lngE) = .curr_id.Left
8880        arr_varCtl(C_L1_NAM, lngE) = "curr_id_lbl"
8890        arr_varCtl(C_L1_LFT, lngE) = .curr_id_lbl.Left
8900        arr_varCtl(C_L2_NAM, lngE) = Null
8910        arr_varCtl(C_L2_LFT, lngE) = Null
8920        arr_varCtl(C_LN_NAM, lngE) = "curr_id_lbl_line"
8930        arr_varCtl(C_LN_LFT, lngE) = .curr_id_lbl_line.Left

8940        lngCtls = lngCtls + 1&
8950        lngE = lngCtls - 1&
8960        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
8970        arr_varCtl(C_CNAM, lngE) = "revcode_ID"
8980        arr_varCtl(C_LFT, lngE) = .revcode_ID.Left
8990        arr_varCtl(C_L1_NAM, lngE) = "revcode_ID_lbl"
9000        arr_varCtl(C_L1_LFT, lngE) = .revcode_ID_lbl.Left
9010        arr_varCtl(C_L2_NAM, lngE) = "revcode_ID_lbl2"
9020        arr_varCtl(C_L2_LFT, lngE) = .revcode_ID_lbl2.Left
9030        arr_varCtl(C_LN_NAM, lngE) = "revcode_ID_lbl_line"
9040        arr_varCtl(C_LN_LFT, lngE) = .revcode_ID_lbl_line.Left

9050      End If

9060    End With

EXITP:
9070    Exit Sub

ERRH:
9080    Select Case ERR.Number
        Case Else
9090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9100    End Select
9110    Resume EXITP

End Sub

Public Function RecCnt() As Long

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

9210    lngRetVal = 0&

9220    With Me
9230      Set rst = .RecordsetClone
9240      With rst
9250        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
9260        Else
9270          .MoveLast
9280          lngRetVal = .RecordCount
9290        End If
9300        .Close
9310      End With
9320    End With

EXITP:
9330    Set rst = Nothing
9340    RecCnt = lngRetVal
9350    Exit Function

ERRH:
9360    lngRetVal = 0&
9370    Select Case ERR.Number
        Case Else
9380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9390    End Select
9400    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

9510    Select Case intWhere
        Case acCmdRecordsGoToFirst
9520      DoCmd.RunCommand acCmdRecordsGoToFirst
9530    Case acCmdRecordsGoToLast
9540      DoCmd.RunCommand acCmdRecordsGoToLast
9550    Case acCmdRecordsGoToPrevious
9560      DoCmd.RunCommand acCmdRecordsGoToPrevious
9570    Case acCmdRecordsGoToNext
9580      DoCmd.RunCommand acCmdRecordsGoToNext
9590    Case Else
9600      If IsMissing(varID) = False Then
9610        With Me
9620          Set rst = .RecordsetClone
9630          With rst
9640            .FindFirst "[rcbas_id] = '" & varID & "'"
9650            If .NoMatch = False Then
9660              Me.Bookmark = .Bookmark
9670            End If
9680            .Close
9690          End With
9700        End With
9710      End If
9720    End Select

EXITP:
9730    Set rst = Nothing
9740    Exit Sub

ERRH:
9750    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
9760    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
9770    Case Else
9780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9790    End Select
9800    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [transdate], [JournalType_Order], [totdesc]"

9910    With Me
9920      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
9930        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
9940      End If
9950      .Controls(strSortLbl).Visible = False
9960      .Controls(strSortLine).Visible = False
9970      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
9980      If strProc = "Form_Load" Then
9990        strCalled = "transdate"
10000       strSortNow = strSortOrig
10010       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
10020       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
10030       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
10040       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
10050       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
10060       .Controls(strSortLbl).Top = lngSortLbl_Top
10070       .Controls(strSortLbl).Left = lngSortLbl_Left
10080       .Controls(strSortLine).Top = lngSortLine_Top
10090       .Controls(strSortLine).Left = lngSortLine_Left
10100       .Controls(strSortLine).Width = lngSortLine_Width
10110       .Controls(strSortLbl).Caption = strArwUp
10120       .Controls(strSortLbl).ForeColor = CLR_DKBLU
10130     Else
10140       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
10150       Select Case strCalled
            Case "journalno", "curr_id"
10160         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
10170         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
10180       Case Else
10190         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
10200         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
10210       End Select
10220       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
10230       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
10240       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
10250       .Controls(strSortLbl).Top = lngSortLbl_Top
10260       .Controls(strSortLbl).Left = lngSortLbl_Left
10270       .Controls(strSortLine).Top = lngSortLine_Top
10280       .Controls(strSortLine).Left = lngSortLine_Left
10290       .Controls(strSortLine).Width = lngSortLine_Width
10300       If strCalled = "transdate" Then
10310         strSortAsc = strSortOrig
10320         If strSortNow = strSortAsc Then
10330           strSortNow = "[transdate] DESC, [JType_sortOrder] DESC, [totdesc] DESC"
10340           .Controls(strSortLbl).Caption = strArwDn
10350           .Controls(strSortLbl).ForeColor = CLR_DKRED
10360         Else
10370           strSortNow = strSortAsc
10380           .Controls(strSortLbl).Caption = strArwUp
10390           .Controls(strSortLbl).ForeColor = CLR_DKBLU
10400         End If
10410       Else
10420         .Controls(strSortLbl).Caption = strArwUp
10430         .Controls(strSortLbl).ForeColor = CLR_DKBLU
10440         Select Case strCalled
              Case "journalno"
10450           strSortAsc = "[journalno]"
10460           lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
10470           .Controls(strSortLbl).Left = lngSortLbl_Left
10480         Case "journaltype"
10490           strSortAsc = "[journaltypex], [transdate], [totdesc]"
10500           lngSortLbl_Left = (lngSortLbl_Left + lngTpp)
10510           .Controls(strSortLbl).Left = lngSortLbl_Left
10520         Case "chkIncExpChk"
10530           strSortAsc = "[chkIncExpChk]" & strStdAsc
10540           lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
10550           .Controls(strSortLbl).Left = lngSortLbl_Left
10560         Case "totdesc"
10570           strSortAsc = "[totdesc], [transdate], [JournalType_Order]"
10580         Case "icash"
10590           strSortAsc = "[icash]" & strStdAsc
10600         Case "pcash"
10610           strSortAsc = "[pcash]" & strStdAsc
10620         Case "cost"
10630           strSortAsc = "[cost]" & strStdAsc
10640         Case "curr_id"
10650           strSortAsc = "[curr_code]" & strStdAsc
10660           lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
10670           .Controls(strSortLbl).Left = lngSortLbl_Left
10680         Case "revcode_ID"
10690           strSortAsc = "[revcode_TYPE], [revcode_DESC]" & strStdAsc
10700         End Select
10710         If strSortNow = strSortAsc Then
10720           intCnt = CharCnt(strSortAsc, ",") + 1
10730           Select Case intCnt
                Case 1
10740             strTmp01 = strSortAsc & " DESC"
10750           Case 2
10760             intPos01 = InStr(strSortAsc, ",")
10770             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10780             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
10790           Case 3
10800             intPos01 = InStr(strSortAsc, ",")
10810             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10820             strTmp02 = Mid(strSortAsc, intPos01)
10830             intPos01 = InStr(2, strTmp02, ",")
10840             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10850             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
10860           Case 4
10870             intPos01 = InStr(strSortAsc, ",")
10880             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10890             strTmp02 = Mid(strSortAsc, intPos01)
10900             intPos01 = InStr(2, strTmp02, ",")
10910             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10920             strTmp02 = Mid(strSortAsc, intPos01)
10930             intPos01 = InStr(2, strTmp02, ",")
10940             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10950             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
10960           Case 5
10970             intPos01 = InStr(strSortAsc, ",")
10980             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10990             strTmp02 = Mid(strSortAsc, intPos01)
11000             intPos01 = InStr(2, strTmp02, ",")
11010             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11020             strTmp02 = Mid(strSortAsc, intPos01)
11030             intPos01 = InStr(2, strTmp02, ",")
11040             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11050             strTmp02 = Mid(strSortAsc, intPos01)
11060             intPos01 = InStr(2, strTmp02, ",")
11070             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11080             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
11090           End Select
11100           strSortNow = strTmp01
11110           .Controls(strSortLbl).Caption = strArwDn
11120           .Controls(strSortLbl).ForeColor = CLR_DKRED
11130         Else
11140           strSortNow = strSortAsc
11150         End If
11160       End If
11170     End If
11180     .Controls(strSortLbl).Visible = True
11190     .Controls(strSortLine).Visible = True
11200     .OrderBy = strSortNow
11210     .OrderByOn = True
11220 On Error Resume Next
11230     .FocusHolder.SetFocus
11240 On Error GoTo ERRH
11250   End With

        ' ** Potential Sorts:

EXITP:
11260   Exit Sub

ERRH:
11270   Select Case ERR.Number
        Case Else
11280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11290   End Select
11300   Resume EXITP

End Sub
