VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAdminOfficer_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAdminOfficer_Sub"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd on frmAdminOfficer}
' **   Close:            C {cmdClose on frmAdminOfficer}
' **   Delete:           D {cmdDelete on frmAdminOfficer}
' **   Officer:          F {officer, cmbAdminNos}
' **   Show Accounts:    H {chkShowAccounts on frmAdminOfficer}
' **   Print Report:     P {cmdPrintReport on frmAdminOfficer}
' **   Unassigned:       U {chkShowUnassigned on frmAdminOfficer}
' **   Exit:             X {cmdClose on frmAdminOfficer}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[officer]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Top As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private lngOfficer_Width As Long

Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim strFilter As String

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAdminOfficer" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190         lngSortLbl_Width = .Sort_lbl.Width
200         lngOfficer_Width = .officer.Width

210         strFilter = "[adminno] <> 1"
220         .Filter = strFilter
230         .FilterOn = True

240       Else
250         Cancel = -1
260       End If
270     End With

280     If Cancel = -1 Then
290       Beep
300       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
310     End If

EXITP:
320     Exit Sub

ERRH:
330     Select Case ERR.Number
        Case Else
340       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
350     End Select
360     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .adminno.ForeColor = CLR_DISABLED_FG
430       .adminno.BackColor = CLR_DISABLED_BG
440       .assigned.ForeColor = CLR_DISABLED_FG
450       .assigned.BackColor = CLR_DISABLED_BG
460       .accountno.ForeColor = CLR_DISABLED_FG
470       .accountno.BackColor = CLR_DISABLED_BG
480       .shortname.ForeColor = CLR_DISABLED_FG
490       .shortname.BackColor = CLR_DISABLED_BG
500     End With

        ' ** Handled by AfterUpdat's in frmAdminOffier.
        'SortNow THIS_PROC

EXITP:
510     Exit Sub

ERRH:
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Private Sub Form_Current()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

610     With Me
620       lngRecsCur = RecCnt  ' ** Function: Below
630       If lngRecsCur > 0& Then
640         If (.adminno = 1 Or .officer = "{Unassigned}") And .Parent.chkShowUnassigned = False Then
              ' ** Though it shouldn't be showing anyway.
650           .officer.Locked = True
660           .Parent.cmdDelete.Enabled = False
670         Else
680           Select Case .Parent.chkShowAccounts
              Case True
690             .officer.Locked = True
700             .Parent.cmdDelete.Enabled = False
710           Case False
720             .officer.Locked = False
730             .Parent.cmdDelete.Enabled = True
740           End Select
750         End If
760       Else
770         .Parent.cmdDelete.Enabled = False
780       End If
790     End With

EXITP:
800     Exit Sub

ERRH:
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

910     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
920     intShiftDown = (Shift And acShiftMask) > 0
930     intAltDown = (Shift And acAltMask) > 0
940     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyUp
970         intRetVal = 0
980         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
990       Case vbKeyDown
1000        intRetVal = 0
1010        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1020      Case vbKeyEscape
1030        With Me
1040          intRetVal = 0
1050          .Parent.cmdClose_Click  ' ** Form Procedure: frmAdminOfficer.
1060        End With
1070      End Select
1080    End If

        ' ** Alt keys.
1090    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1100      Select Case intRetVal
          Case vbKeyX
1110        With Me
1120          intRetVal = 0
1130          .Parent.cmdClose_Click  ' ** Form Procedure: frmAdminOfficer.
1140        End With
1150      End Select
1160    End If

        ' ** Ctrl keys.
1170    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1180      Select Case intRetVal
          Case vbKeyS
1190        intRetVal = 0
1200        cmdSave_Click  ' ** Procedure: Below.
1210      Case vbKeyTab, vbKeyReturn
1220        With Me
1230          intRetVal = 0
1240          DoCmd.SelectObject acForm, .Parent.Name, False
1250          If .Parent.chkShowAccounts.Enabled = True Then
1260            .Parent.chkShowAccounts.SetFocus
1270          ElseIf Parent.chkShowUnassigned.Enabled = True Then
1280            .Parent.chkShowUnassigned.SetFocus
1290          ElseIf .Parent.cmdPrintReport.Enabled = True Then
1300            .Parent.cmdPrintReport.SetFocus
1310          ElseIf .Parent.cmdAdd.Enabled = True Then
1320            .Parent.cmdAdd.SetFocus
1330          ElseIf .Parent.cmdDelete.Enabled = True Then
1340            .Parent.cmdDelete.SetFocus
1350          Else
1360            .Parent.cmdClose.SetFocus
1370          End If
1380        End With
1390      End Select
1400    End If

        ' ** Ctrl-Shift keys.
1410    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1420      Select Case intRetVal
          Case vbKeyF
1430        With Me
1440          intRetVal = 0
1450          DoCmd.SelectObject acForm, .Parent.Name, False
1460          .Parent.FocusHolder.SetFocus
1470        End With
1480      Case vbKeyTab, vbKeyReturn
1490        With Me
1500          intRetVal = 0
1510          DoCmd.SelectObject acForm, .Parent.Name, False
1520          .Parent.cmdClose.SetFocus
1530        End With
1540      End Select
1550    End If

EXITP:
1560    KeyCode = intRetVal
1570    Exit Sub

ERRH:
1580    intRetVal = 0
1590    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
1600    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
1610    Case Else
1620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1630    End Select
1640    Resume EXITP

End Sub

Public Sub cmdSave_Click()
' ** This is a fake click event.

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1710    DoCmd.RunCommand acCmdSaveRecord

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub adminno_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "adminno_lbl_DblClick"

1810    SortNow THIS_PROC

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub officer_lbl_DblClick(Cancel As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "officer_lbl_DblClick"

1910    SortNow THIS_PROC

EXITP:
1920    Exit Sub

ERRH:
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Private Sub officer_GotFocus()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "officer_GotFocus"

2010    With Me
2020      .officer.SelLength = 0
2030      .officer.SelStart = 99
2040    End With

EXITP:
2050    Exit Sub

ERRH:
2060    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No admin officers.
2070    Case Else
2080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2090    End Select
2100    Resume EXITP

End Sub

Private Sub officer_AfterUpdate()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "officer_AfterUpdate"

2210    cmdSave_Click  ' ** Procedure: Above.

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Private Sub officer_KeyDown(KeyCode As Integer, Shift As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "officer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2320    intShiftDown = (Shift And acShiftMask) > 0
2330    intAltDown = (Shift And acAltMask) > 0
2340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2370        With Me
2380          intRetVal = 0
2390          Select Case .Parent.chkShowAccounts
              Case True
2400            .accountno.SetFocus
2410          Case False
2420            lngRecsCur = RecCnt  ' ** Function: Below.
2430            If .CurrentRecord < lngRecsCur Then
2440              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2450            Else
2460              DoCmd.SelectObject acForm, .Parent.Name, False
2470              If .Parent.chkShowAccounts.Enabled = True Then
2480                .Parent.chkShowAccounts.SetFocus
2490              ElseIf Parent.chkShowUnassigned.Enabled = True Then
2500                .Parent.chkShowUnassigned.SetFocus
2510              ElseIf .Parent.cmdPrintReport.Enabled = True Then
2520                .Parent.cmdPrintReport.SetFocus
2530              ElseIf .Parent.cmdAdd.Enabled = True Then
2540                .Parent.cmdAdd.SetFocus
2550              ElseIf .Parent.cmdDelete.Enabled = True Then
2560                .Parent.cmdDelete.SetFocus
2570              Else
2580                .Parent.cmdClose.SetFocus
2590              End If
2600            End If
2610          End Select
2620        End With
2630      End Select
2640    End If

        ' ** Shift keys.
2650    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2670        With Me
2680          intRetVal = 0
2690          If .CurrentRecord > 1 Then
2700            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2710            Select Case .Parent.chkShowAccounts
                Case True
2720              .shortname.SetFocus
2730            Case False
2740              .officer.SetFocus
2750            End Select
2760          Else
2770            DoCmd.SelectObject acForm, .Parent.Name, False
2780            .Parent.cmdClose.SetFocus
2790          End If
2800        End With
2810      End Select
2820    End If

EXITP:
2830    KeyCode = intRetVal
2840    Exit Sub

ERRH:
2850    intRetVal = 0
2860    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
2870    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
2880    Case Else
2890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2900    End Select
2910    Resume EXITP

End Sub

Private Sub cmbAdminNos_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAdminNos_lbl_DblClick"

3010    SortNow THIS_PROC

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub cmbAdminNos_AfterUpdate()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAdminNos_AfterUpdate"

        Dim varTmp00 As Variant

3110    With Me
3120      If IsNull(.cmbAdminNos) = False Then
3130        cmdSave_Click  ' ** Procedure: Above.
3140        DoEvents
3150        varTmp00 = DCount("*", "qryAdminOfficer_12")
3160        If IsNull(varTmp00) = True Then
3170          .Parent.Unassigned_lbl.Caption = "0 Accounts Unassigned"
3180        Else
3190          If varTmp00 = 0 Then
3200            .Parent.Unassigned_lbl.Caption = "0 Accounts Unassigned"
3210          Else
3220            .Parent.Unassigned_lbl.Caption = CStr(varTmp00) & " Account" & IIf(varTmp00 = 1, vbNullString, "s") & " Unassigned"
3230          End If
3240        End If
3250      End If
3260    End With

EXITP:
3270    Exit Sub

ERRH:
3280    Select Case ERR.Number
        Case Else
3290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3300    End Select
3310    Resume EXITP

End Sub

Private Sub cmbAdminNos_NotInList(NewData As String, Response As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAdminNos_NotInList"

3410    With Me
3420      .cmbAdminNos = 1
3430      MsgBox "Please select an Administrator from the list.", vbInformation + vbOKOnly, "Invalid Entry"
3440      Response = acDataErrContinue
3450    End With

EXITP:
3460    Exit Sub

ERRH:
3470    Select Case ERR.Number
        Case Else
3480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3490    End Select
3500    Resume EXITP

End Sub

Private Sub cmbAdminNos_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAdminNos_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3620    intShiftDown = (Shift And acShiftMask) > 0
3630    intAltDown = (Shift And acAltMask) > 0
3640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3670        With Me
3680          intRetVal = 0
3690          .accountno.SetFocus
3700        End With
3710      End Select
3720    End If

        ' ** Shift keys.
3730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3750        With Me
3760          intRetVal = 0
3770          If .CurrentRecord > 1 Then
3780            .MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3790            .shortname.SetFocus
3800          Else
3810            DoCmd.SelectObject acForm, .Parent.Name, False
3820            .Parent.cmdClose.SetFocus
3830          End If
3840        End With
3850      End Select
3860    End If

EXITP:
3870    KeyCode = intRetVal
3880    Exit Sub

ERRH:
3890    intRetVal = 0
3900    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
3910    Case 2110  ' ** Microsoft Access can't move the focus to the control |.
          ' ** Ignore.
3920    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
3930    Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub assigned_lbl_DblClick(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "assigned_lbl_DblClick"

4010    SortNow THIS_PROC

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

4110    With Me
4120      If .Parent.chkShowAccounts = True Or .Parent.chkShowUnassigned = True Then
4130        SortNow THIS_PROC
4140      End If
4150    End With

EXITP:
4160    Exit Sub

ERRH:
4170    Select Case ERR.Number
        Case Else
4180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4190    End Select
4200    Resume EXITP

End Sub

Private Sub accountno_GotFocus()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_GotFocus"

4310    With Me
4320      .accountno.SelLength = 0
4330      .accountno.SelStart = 19
4340    End With

EXITP:
4350    Exit Sub

ERRH:
4360    Select Case ERR.Number
        Case Else
4370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4380    End Select
4390    Resume EXITP

End Sub

Private Sub accountno_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4420    intShiftDown = (Shift And acShiftMask) > 0
4430    intAltDown = (Shift And acAltMask) > 0
4440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4470        With Me
4480          intRetVal = 0
4490          .shortname.SetFocus
4500        End With
4510      End Select
4520    End If

        ' ** Shift keys.
4530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4550        With Me
4560          intRetVal = 0
4570          .officer.SetFocus
4580        End With
4590      End Select
4600    End If

EXITP:
4610    KeyCode = intRetVal
4620    Exit Sub

ERRH:
4630    intRetVal = 0
4640    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
4650    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4660    Case Else
4670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4680    End Select
4690    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

4710    With Me
4720      If .Parent.chkShowAccounts = True Or .Parent.chkShowUnassigned = True Then
4730        SortNow THIS_PROC
4740      End If
4750    End With

EXITP:
4760    Exit Sub

ERRH:
4770    Select Case ERR.Number
        Case Else
4780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4790    End Select
4800    Resume EXITP

End Sub

Private Sub shortname_GotFocus()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_GotFocus"

4910    With Me
4920      .shortname.SelLength = 0
4930      .shortname.SelStart = 0
4940    End With

EXITP:
4950    Exit Sub

ERRH:
4960    Select Case ERR.Number
        Case Else
4970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4980    End Select
4990    Resume EXITP

End Sub

Private Sub shortname_KeyDown(KeyCode As Integer, Shift As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5020    intShiftDown = (Shift And acShiftMask) > 0
5030    intAltDown = (Shift And acAltMask) > 0
5040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5070        With Me
5080          intRetVal = 0
5090          lngRecsCur = RecCnt  ' ** Function: Below.
5100          If .CurrentRecord < lngRecsCur Then
5110            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5120            .officer.SetFocus
5130          Else
5140            DoCmd.SelectObject acForm, .Parent.Name, False
5150            If .Parent.chkShowAccounts.Enabled = True Then
5160              .Parent.chkShowAccounts.SetFocus
5170            ElseIf Parent.chkShowUnassigned.Enabled = True Then
5180              .Parent.chkShowUnassigned.SetFocus
5190            ElseIf .Parent.cmdPrintReport.Enabled = True Then
5200              .Parent.cmdPrintReport.SetFocus
5210            ElseIf .Parent.cmdAdd.Enabled = True Then
5220              .Parent.cmdAdd.SetFocus
5230            ElseIf .Parent.cmdDelete.Enabled = True Then
5240              .Parent.cmdDelete.SetFocus
5250            Else
5260              .Parent.cmdClose.SetFocus
5270            End If
5280          End If
5290        End With
5300      End Select
5310    End If

        ' ** Shift keys.
5320    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5330      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5340        With Me
5350          intRetVal = 0
5360          .accountno.SetFocus
5370        End With
5380      End Select
5390    End If

EXITP:
5400    KeyCode = intRetVal
5410    Exit Sub

ERRH:
5420    intRetVal = 0
5430    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
5440    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5450    Case Else
5460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5470    End Select
5480    Resume EXITP

End Sub

Public Sub ShowAccounts(blnShow As Boolean)
' ** Resorting is handled in AfterUpdates() on frmAdminOfficer.

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAccounts"

        Dim blnSortHere As Boolean, blnResort As Boolean

5510    With Me
5520      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5530        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5540      End If
5550      blnSortHere = False: blnResort = False
5560      Select Case blnShow
          Case True
5570  On Error Resume Next
5580        .FocusHolder.SetFocus
5590  On Error GoTo ERRH
5600        If .Sort_line.Left = .assigned_lbl.Left Then blnSortHere = True
5610        .assigned.Visible = False
5620        .assigned_lbl.Visible = False
5630        .assigned_lbl2.Visible = False
5640        .assigned_lbl_line.Visible = False
5650        If blnSortHere = True Then
5660          .Sort_line.Visible = False
5670          .Sort_lbl.Visible = False
5680          blnResort = True
5690        End If
5700        blnSortHere = False
5710        If .Sort_line.Left = .officer_lbl.Left Then blnSortHere = True
5720        .officer.Width = ((.accountno.Left - .officer.Left) - (4& * lngTpp))
5730        .officer.Locked = True
5740        .officer.ForeColor = CLR_DISABLED_FG
5750        .officer.BackColor = CLR_DISABLED_BG
5760        .officer_lbl.Width = .officer.Width
5770        .officer_lbl_line.Width = (.officer_lbl.Width + lngTpp)
5780        If blnSortHere = True Then
5790          .Sort_line.Width = (.officer_lbl.Width + lngTpp)
5800          .Sort_lbl.Left = ((.officer_lbl.Left + .officer_lbl.Width) - lngSortLbl_Width)
5810        End If
5820        blnSortHere = False
5830        .accountno.Visible = True
5840        .accountno_lbl.Left = .accountno.Left
5850        .accountno_lbl.Visible = True
5860        .accountno_lbl_line.Left = .accountno_lbl.Left
5870        .accountno_lbl_line.Visible = True
5880        .shortname.Visible = True
5890        .shortname_lbl.Left = .shortname.Left
5900        .shortname_lbl.Visible = True
5910        .shortname_lbl_line.Left = .shortname_lbl.Left
5920        .shortname_lbl_line.Visible = True
5930        .RecordSource = "qryAdminOfficer_08"
5940        .Filter = "[adminno] > 1"
5950        .FilterOn = True
5960        DoEvents
5970      Case False
5980  On Error Resume Next
5990        .FocusHolder.SetFocus
6000  On Error GoTo ERRH
6010        If .Sort_line.Left = .accountno_lbl.Left Then blnSortHere = True
6020        .accountno.Visible = False
6030        .accountno_lbl.Visible = False
6040        .accountno_lbl_line.Visible = False
6050        If blnSortHere = True Then
6060          .Sort_line.Visible = False
6070          .Sort_lbl.Visible = False
6080          blnResort = True
6090        End If
6100        blnSortHere = False
6110        If .Sort_line.Left = .shortname_lbl.Left Then blnSortHere = True
6120        .shortname.Visible = False
6130        .shortname_lbl.Visible = False
6140        .shortname_lbl_line.Visible = False
6150        If blnSortHere = True Then
6160          .Sort_line.Visible = False
6170          .Sort_lbl.Visible = False
6180          blnResort = True
6190        End If
6200        blnSortHere = False
6210        If .Sort_line.Left = .officer_lbl.Left Then blnSortHere = True
6220        .officer.Width = lngOfficer_Width
6230        .officer.ForeColor = CLR_BLK
6240        .officer.BackColor = CLR_WHT
6250        .officer.Locked = False
6260        .officer_lbl.Width = .officer.Width
6270        .officer_lbl_line.Width = (.officer_lbl.Width + lngTpp)
6280        If blnSortHere = True Then
6290          .Sort_line.Width = (.officer_lbl.Width + lngTpp)
6300          .Sort_lbl.Left = ((.officer_lbl.Left + .officer_lbl.Width) - lngSortLbl_Width)
6310        End If
6320        blnSortHere = False
6330        .assigned.Visible = True
6340        .assigned_lbl.Visible = True
6350        .assigned_lbl2.Visible = True
6360        .assigned_lbl_line.Visible = True
6370        .RecordSource = "qryAdminOfficer_06"
6380        .Filter = "[adminno] > 1"
6390        .FilterOn = True
6400        DoEvents
6410      End Select
6420    End With

EXITP:
6430    Exit Sub

ERRH:
6440    Select Case ERR.Number
        Case Else
6450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6460    End Select
6470    Resume EXITP

End Sub

Public Sub ShowUnassigned(blnShow As Boolean)
' ** Resorting is handled in AfterUpdates() on frmAdminOfficer.

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowUnassigned"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim strFilter As String

6510    With Me
6520      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
6530        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6540      End If
6550      blnSortHere = False: blnResort = False
6560      lngRecsCur = RecCnt  ' ** Function: Below.
6570      If lngRecsCur > 0& Then
6580        .FocusHolder.SetFocus
6590      End If
6600      Select Case blnShow
          Case True
6610        .cmbAdminNos.Visible = True
6620        .cmbAdminNos_lbl.Visible = True
6630        .cmbAdminNos_lbl_line.Visible = True
6640        If .Sort_line.Left = .adminno_lbl.Left Then blnSortHere = True
6650        .adminno.Visible = False
6660        .adminno_lbl.Visible = False
6670        .adminno_lbl_line.Visible = False
6680        If blnSortHere = True Then
6690          .Sort_line.Visible = False
6700          .Sort_lbl.Visible = False
6710          blnResort = True
6720        End If
6730        blnSortHere = False
6740        If .Sort_line.Left = .officer_lbl.Left Then blnSortHere = True
6750        .officer.Visible = False
6760        .officer_lbl.Visible = False
6770        .officer_lbl_line.Visible = False
6780        If blnSortHere = True Then
6790          .Sort_line.Visible = False
6800          .Sort_lbl.Visible = False
6810          blnResort = True
6820        End If
6830        blnSortHere = False
6840        If .Sort_line.Left = .assigned_lbl.Left Then blnSortHere = True
6850        .assigned.Visible = False
6860        .assigned_lbl.Visible = False
6870        .assigned_lbl2.Visible = False
6880        .assigned_lbl_line.Visible = False
6890        If blnSortHere = True Then
6900          .Sort_line.Visible = False
6910          .Sort_lbl.Visible = False
6920          blnResort = True
6930        End If
6940        blnSortHere = False
6950        If .Sort_line.Left = .accountno_lbl.Left Then blnSortHere = True
6960        .accountno.Visible = True
6970        .accountno_lbl.Visible = True
6980        .accountno_lbl_line.Visible = True
6990        If blnSortHere = True Then
7000          .Sort_line.Visible = True  ' ** These would have already been there and visible.
7010          .Sort_lbl.Visible = True
7020        End If
7030        blnSortHere = False
7040        If .Sort_line.Left = .shortname_lbl.Left Then blnSortHere = True
7050        .shortname.Visible = True
7060        .shortname_lbl.Visible = True
7070        .shortname_lbl_line.Visible = True
7080        If blnSortHere = True Then
7090          .Sort_line.Visible = True  ' ** These would have already been there and visible.
7100          .Sort_lbl.Visible = True
7110        End If
7120        blnSortHere = False
7130        If .RecordSource <> "qryAdminOfficer_12" Then
              ' ** Account, just adminno = 1.
7140          .RecordSource = "qryAdminOfficer_12"
7150          strFilter = vbNullString
7160          .Filter = strFilter
7170          .FilterOn = False
7180        End If
7190      Case False
7200        If .Sort_line.Left = .cmbAdminNos_lbl.Left Then blnSortHere = True
7210        .cmbAdminNos.Visible = False
7220        .cmbAdminNos_lbl.Visible = False
7230        .cmbAdminNos_lbl_line.Visible = False
7240        If blnSortHere = True Then
7250          .Sort_line.Visible = False
7260          .Sort_lbl.Visible = False
7270          blnResort = True
7280        End If
7290        blnSortHere = False
7300        Select Case .Parent.chkShowAccounts
            Case True
7310          .adminno.Visible = True           ' ** Coming from Unassigned, these couldn't have been the sort.
7320          .adminno_lbl.Visible = True
7330          .adminno_lbl_line.Visible = True
7340          .officer.Visible = True           ' ** "
7350          .officer_lbl.Visible = True
7360          .officer_lbl_line.Visible = True
7370          .assigned.Visible = False         ' ** "
7380          .assigned_lbl.Visible = False
7390          .assigned_lbl2.Visible = False
7400          .assigned_lbl_line.Visible = False
7410          If .Sort_line.Left = .accountno_lbl.Left Then blnSortHere = True
7420          .accountno.Visible = True
7430          .accountno_lbl.Visible = True
7440          .accountno_lbl_line.Visible = True
7450          If blnSortHere = True Then
7460            .Sort_line.Visible = True  ' ** These would have already been there and visible.
7470            .Sort_lbl.Visible = True
7480          End If
7490          blnSortHere = False
7500          If .Sort_line.Left = .shortname_lbl.Left Then blnSortHere = True
7510          .shortname.Visible = True
7520          .shortname_lbl.Visible = True
7530          .shortname_lbl_line.Visible = True
7540          If blnSortHere = True Then
7550            .Sort_line.Visible = True  ' ** These would have already been there and visible.
7560            .Sort_lbl.Visible = True
7570          End If
7580          blnSortHere = False
7590          If .RecordSource <> "qryAdminOfficer_08" Then
                ' ** AdminOfficer, linked to qryAdminOfficer_05_01
                ' ** (Account, just dateclosed = Null), sorted.
7600            .RecordSource = "qryAdminOfficer_08"
7610            strFilter = "[adminno] > 1"
7620            .Filter = strFilter
7630            .FilterOn = True
7640            DoEvents
7650          End If
7660        Case False
7670          .adminno.Visible = True           ' ** Coming from Unassigned, these couldn't have been the sort.
7680          .adminno_lbl.Visible = True
7690          .adminno_lbl_line.Visible = True
7700          .officer.Visible = True           ' ** "
7710          .officer_lbl.Visible = True
7720          .officer_lbl_line.Visible = True
7730          .assigned.Visible = True          ' ** "
7740          .assigned_lbl.Visible = True
7750          .assigned_lbl2.Visible = True
7760          .assigned_lbl_line.Visible = True
7770          If .Sort_line.Left = .accountno_lbl.Left Then blnSortHere = True
7780          .accountno.Visible = False
7790          .accountno_lbl.Visible = False
7800          .accountno_lbl_line.Visible = False
7810          If blnSortHere = True Then
7820            .Sort_line.Visible = False
7830            .Sort_lbl.Visible = False
7840            blnResort = True
7850          End If
7860          blnSortHere = False
7870          If .Sort_line.Left = .shortname_lbl.Left Then blnSortHere = True
7880          .shortname.Visible = False
7890          .shortname_lbl.Visible = False
7900          .shortname_lbl_line.Visible = False
7910          If blnSortHere = True Then
7920            .Sort_line.Visible = False
7930            .Sort_lbl.Visible = False
7940            blnResort = True
7950          End If
7960          blnSortHere = False
7970          If .RecordSource <> "qryAdminOfficer_06" Then
                ' ** AdminOfficer, with DLookups() to qryAdminOfficer_05
                ' ** (AdminOfficer, linked to qryAdminOfficer_05_01 (Account,
                ' ** just dateclosed = Null), grouped by adminno, with cnt).
7980            .RecordSource = "qryAdminOfficer_06"
7990            strFilter = "[adminno] > 1"
8000            .Filter = strFilter
8010            .FilterOn = True
8020            DoEvents
8030          End If
8040        End Select
8050      End Select
8060    End With

EXITP:
8070    Exit Sub

ERRH:
8080    Select Case ERR.Number
        Case Else
8090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8100    End Select
8110    Resume EXITP

End Sub

Public Function RecCnt() As Long

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

8210    lngRetVal = 0&

8220    With Me
8230      Set rst = .RecordsetClone
8240      With rst
8250        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
8260        Else
8270  On Error Resume Next
8280          .MoveLast
8290  On Error GoTo ERRH
8300          lngRetVal = .RecordCount
8310        End If
8320        .Close
8330      End With
8340    End With

EXITP:
8350    Set rst = Nothing
8360    RecCnt = lngRetVal
8370    Exit Function

ERRH:
8380    lngRetVal = 0&
8390    Select Case ERR.Number
        Case Else
8400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8410    End Select
8420    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

8510    Select Case intWhere
        Case acCmdRecordsGoToFirst
8520      DoCmd.RunCommand acCmdRecordsGoToFirst
8530    Case acCmdRecordsGoToLast
8540      DoCmd.RunCommand acCmdRecordsGoToLast
8550    Case acCmdRecordsGoToPrevious
8560      DoCmd.RunCommand acCmdRecordsGoToPrevious
8570    Case acCmdRecordsGoToNext
8580      DoCmd.RunCommand acCmdRecordsGoToNext
8590    Case Else
8600      If IsMissing(varID) = False Then
8610        With Me
8620          Set rst = .RecordsetClone
8630          With rst
8640            .FindFirst "[adminno] = " & CStr(varID)
8650            If .NoMatch = False Then
8660              Me.Bookmark = .Bookmark
8670            End If
8680            .Close
8690          End With
8700        End With
8710      End If
8720    End Select

EXITP:
8730    Set rst = Nothing
8740    Exit Sub

ERRH:
8750    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8760    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8770    Case Else
8780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8790    End Select
8800    Resume EXITP

End Sub

Public Sub SortNow(strProc As String, Optional varShow As Variant)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim blnAccts As Boolean, blnDontReverse As Boolean
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [officer]"

8910    With Me
8920      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
8930        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
8940      End If
8950      blnAccts = .Parent.chkShowAccounts
8960      Select Case IsMissing(varShow)
          Case True
8970        blnDontReverse = False
8980      Case False
8990        blnDontReverse = CBool(varShow)
9000      End Select
9010      .Controls(strSortLbl).Visible = False
9020      .Controls(strSortLine).Visible = False
9030      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
9040      If strProc = "Form_Load" Then
9050        strCalled = "officer"
9060        strSortNow = strSortOrig
9070        If blnAccts = True Then
9080          strSortNow = strSortNow & ", [alphasort]"
9090        End If
9100        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9110        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
9120        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9130        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
9140        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
9150        .Controls(strSortLbl).Top = lngSortLbl_Top
9160        .Controls(strSortLbl).Left = lngSortLbl_Left
9170        .Controls(strSortLine).Top = lngSortLine_Top
9180        .Controls(strSortLine).Left = lngSortLine_Left
9190        .Controls(strSortLine).Width = lngSortLine_Width
9200        .Controls(strSortLbl).Caption = strArwUp
9210        .Controls(strSortLbl).ForeColor = CLR_DKBLU
9220      Else
9230        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
9240        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9250        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
9260        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9270        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
9280        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
9290        .Controls(strSortLbl).Top = lngSortLbl_Top
9300        .Controls(strSortLbl).Left = lngSortLbl_Left
9310        .Controls(strSortLine).Top = lngSortLine_Top
9320        .Controls(strSortLine).Left = lngSortLine_Left
9330        .Controls(strSortLine).Width = lngSortLine_Width
9340        If strCalled = "officer" Then
9350          strSortAsc = strSortOrig
9360          If strSortNow = strSortAsc And blnDontReverse = False Then
9370            strSortNow = "[officer] DESC"
9380            If blnAccts = True Then
9390              strSortNow = strSortNow & ", [alphasort] DESC"
9400            End If
9410            .Controls(strSortLbl).Caption = strArwDn
9420            .Controls(strSortLbl).ForeColor = CLR_DKRED
9430          Else
9440            strSortNow = strSortAsc
9450            If blnAccts = True Then
9460              strSortNow = strSortNow & ", [alphasort]"
9470            End If
9480            .Controls(strSortLbl).Caption = strArwUp
9490            .Controls(strSortLbl).ForeColor = CLR_DKBLU
9500          End If
9510        Else
9520          .Controls(strSortLbl).Caption = strArwUp
9530          .Controls(strSortLbl).ForeColor = CLR_DKBLU
9540          Select Case strCalled
              Case "adminno"
9550            Select Case blnAccts
                Case True
9560              strSortAsc = "[adminno], [alphasort]"
9570            Case False
9580              strSortAsc = "[adminno]"
9590            End Select
9600          Case "cmbAdminNos"
9610            strSortAsc = "[officer], [alphasort]"
9620          Case "assigned"
9630            strSortAsc = "[assigned]" & strStdAsc
9640            lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
9650            .Controls(strSortLbl).Left = lngSortLbl_Left
9660          Case "accountno"
9670            strSortAsc = "[alphasort]"
9680            lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
9690            .Controls(strSortLbl).Left = lngSortLbl_Left
9700          Case "shortname"
9710            strSortAsc = "[shortname]"
9720          End Select
9730          If strSortNow = strSortAsc And blnDontReverse = False Then
9740            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
9750            Select Case intCnt
                Case 1
9760              strTmp01 = strSortAsc & " DESC"
9770            Case 2
9780              intPos01 = InStr(strSortAsc, ",")
9790              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9800              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
9810            Case 3
9820              intPos01 = InStr(strSortAsc, ",")
9830              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9840              strTmp02 = Mid(strSortAsc, intPos01)
9850              intPos01 = InStr(2, strTmp02, ",")
9860              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9870              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
9880            End Select
9890            strSortNow = strTmp01
9900            .Controls(strSortLbl).Caption = strArwDn
9910            .Controls(strSortLbl).ForeColor = CLR_DKRED
9920          Else
9930            strSortNow = strSortAsc
9940          End If
9950        End If
9960      End If
9970      .Controls(strSortLbl).Visible = True
9980      .Controls(strSortLine).Visible = True
9990      .OrderBy = strSortNow
10000     .OrderByOn = True

10010     Select Case .Parent.chkShowUnassigned
          Case True
10020       .Parent.SortNow_Set strSortNow, 3
10030     Case False
10040       Select Case .Parent.chkShowAccounts
            Case True
10050         .Parent.SortNow_Set strSortNow, 2
10060       Case False
10070         .Parent.SortNow_Set strSortNow, 1
10080       End Select
10090     End Select

10100   End With

        ' ** Potential Sorts:
        ' **   [adminno]
        ' **   [officer], [accountno]
        ' **   [assigned], [officer]
        ' **   [accountno]
        ' **   [shortname]

EXITP:
10110   Exit Sub

ERRH:
10120   Select Case ERR.Number
        Case Else
10130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10140   End Select
10150   Resume EXITP

End Sub

Public Function SortNow_Get() As String

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

10210   strRetVal = strSortNow

EXITP:
10220   SortNow_Get = strRetVal
10230   Exit Function

ERRH:
10240   strRetVal = strSortOrig
10250   Select Case ERR.Number
        Case Else
10260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10270   End Select
10280   Resume EXITP

End Function
