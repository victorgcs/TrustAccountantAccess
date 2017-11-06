VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmInvestmentObjective_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmInvestmentObjective_Sub"

'VGC 04/19/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmInvestmentObjective}
' **   Show Accounts:    H {chkShowAccounts on frmInvestmentObjective}
' **   Objective:        O {invobj_name, cmbInvestmentObjectives}
' **   Print Report:     P {cmdPrintReport on frmInvestmentObjective}
' **   Show Unassigned:  U {chkShowUnassigned on frmInvestmentObjective}
' **   Exit:             X {cmdClose on frmInvestmentObjective}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:          F5 {RecalcTots}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[invobj_id]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngObjName_Width As Long
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
150       If strCallingForm = "frmInvestmentObjective" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170         lngSortLbl_Width = .Sort_lbl.Width
180         lngObjName_Width = .invobj_name.Width

190         CLR_DISABLED_FG = CLR_DKGRY
200         CLR_DISABLED_BG = CLR_LTTEAL

210       Else
220         Cancel = -1
230       End If
240     End With

250     If Cancel = -1 Then
260       Beep
270       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
280     End If

EXITP:
290     Exit Sub

ERRH:
300     Select Case ERR.Number
        Case Else
310       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
320     End Select
330     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .invobj_id.ForeColor = CLR_DISABLED_FG
430       .invobj_id.BackColor = CLR_DISABLED_BG
440       .invobj_name.ForeColor = CLR_DISABLED_FG
450       .invobj_name.BackColor = CLR_DISABLED_BG
460       .acct_count.ForeColor = CLR_DISABLED_FG
470       .acct_count.BackColor = CLR_DISABLED_BG
480       .accountno.ForeColor = CLR_DISABLED_FG
490       .accountno.BackColor = CLR_DISABLED_BG
500       .shortname.ForeColor = CLR_DISABLED_FG
510       .shortname.BackColor = CLR_DISABLED_BG
520     End With

        ' ** Handled by AfterUpdat's in frmAdminOffier.
        'SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
530     Exit Sub

ERRH:
540     Select Case ERR.Number
        Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
560     End Select
570     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc:          F5 {RecalcTots}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Objective:        O {invobj_name, cmbInvestmentObjectives}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyF5
670         intRetVal = 0
680         RecalcTots  ' ** Procedure: Below.
690       Case vbKeyUp
700         intRetVal = 0
710         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
720       Case vbKeyDown
730         intRetVal = 0
740         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
750       Case vbKeyEscape
760         With Me
770           intRetVal = 0
780           .Parent.cmdClose_Click  ' ** Form Procedure: frmInvestmentObjectives.
790         End With
800       End Select
810     End If

        ' ** Alt keys.
820     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
830       Select Case intRetVal
          Case vbKeyO
840         With Me
850           intRetVal = 0
860           If .invobj_name.Visible = True And .invobj_name.Enabled = True Then
870             .invobj_name.SetFocus
880           ElseIf .cmbInvestmentObjectives.Visible = True And .cmbInvestmentObjectives.Enabled = True Then
890             .cmbInvestmentObjectives.SetFocus
900           End If
910         End With
920       Case vbKeyX
930         With Me
940           intRetVal = 0
950           .Parent.cmdClose_Click  ' ** Form Procedure: frmInvestmentObjectives.
960         End With
970       End Select
980     End If

        ' ** Ctrl keys.
990     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1000      Select Case intRetVal
          Case vbKeyS
1010        intRetVal = 0
1020        cmdSave_Click  ' ** Procedure: Below.
1030      Case vbKeyTab, vbKeyReturn
1040        With Me
1050          intRetVal = 0
1060          DoCmd.SelectObject acForm, .Parent.Name, False
1070          If .Parent.chkShowAccounts.Enabled = True Then
1080            .Parent.chkShowAccounts.SetFocus
1090          ElseIf .Parent.chkShowUnassigned.Enabled = True Then
1100            .Parent.chkShowUnassigned.SetFocus
1110          ElseIf .Parent.cmdPrintReport.Enabled = True Then
1120            .Parent.cmdPrintReport.SetFocus
1130          Else
1140            .Parent.cmdClose.SetFocus
1150          End If
1160        End With
1170      End Select
1180    End If

        ' ** Ctrl-Shift keys.
1190    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1200      Select Case intRetVal
          Case vbKeyF
1210        With Me
1220          intRetVal = 0
1230          DoCmd.SelectObject acForm, .Parent.Name, False
1240          .Parent.FocusHolder.SetFocus
1250        End With
1260      Case vbKeyTab, vbKeyReturn
1270        With Me
1280          intRetVal = 0
1290          DoCmd.SelectObject acForm, .Parent.Name, False
1300          .Parent.cmdClose.SetFocus
1310        End With
1320      End Select
1330    End If

EXITP:
1340    KeyCode = intRetVal
1350    Exit Sub

ERRH:
1360    intRetVal = 0
1370    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1380    Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdSave_Click()
' ** This is a fake click event!

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1510    DoCmd.RunCommand acCmdSelectRecord
1520    DoCmd.RunCommand acCmdSaveRecord

EXITP:
1530    Exit Sub

ERRH:
1540    Select Case ERR.Number
        Case Else
1550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1560    End Select
1570    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1620    intShiftDown = (Shift And acShiftMask) > 0
1630    intAltDown = (Shift And acAltMask) > 0
1640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1670        With Me
1680          intRetVal = 0
1690          lngRecsCur = RecCnt  ' ** Function: Below.
1700          If .CurrentRecord < lngRecsCur Then
1710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1720            .FocusHolder.SetFocus
1730          Else
1740            DoCmd.SelectObject acForm, .Parent.Name, False
1750            If .Parent.chkShowAccounts.Enabled = True Then
1760              .Parent.chkShowAccounts.SetFocus
1770            ElseIf .Parent.chkShowUnassigned.Enabled = True Then
1780              .Parent.chkShowUnassigned.SetFocus
1790            ElseIf .Parent.cmdPrintReport.Enabled = True Then
1800              .Parent.cmdPrintReport.SetFocus
1810            Else
1820              .Parent.cmdClose.SetFocus
1830            End If
1840          End If
1850        End With
1860      End Select
1870    End If

        ' ** Shift keys.
1880    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1890      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1900        With Me
1910          intRetVal = 0
1920          If .CurrentRecord > 1 Then
1930            MoveRec acCmdRecordsGoToPrevious
1940            .FocusHolder.SetFocus
1950          Else
1960            DoCmd.SelectObject acForm, .Parent.Name, False
1970            .Parent.cmdClose.SetFocus
1980          End If
1990        End With
2000      End Select
2010    End If

EXITP:
2020    KeyCode = intRetVal
2030    Exit Sub

ERRH:
2040    intRetVal = 0
2050    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2060    Case Else
2070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2080    End Select
2090    Resume EXITP

End Sub

Private Sub invobj_id_lbl_DblClick(Cancel As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "invobj_id_lbl_DblClick"

2110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub invobj_name_lbl_DblClick(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "invobj_name_lbl_DblClick"

2210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Private Sub invobj_name_GotFocus()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "invobj_name_GotFocus"

2310    With Me
2320      .invobj_name.SelLength = 0
2330      .invobj_name.SelStart = 99
2340    End With

EXITP:
2350    Exit Sub

ERRH:
2360    Select Case ERR.Number
        Case Else
2370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2380    End Select
2390    Resume EXITP

End Sub

Private Sub invobj_name_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "invobj_name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2420    intShiftDown = (Shift And acShiftMask) > 0
2430    intAltDown = (Shift And acAltMask) > 0
2440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2470        With Me
2480          intRetVal = 0
2490          Select Case .Parent.chkShowAccounts
              Case True
2500            .accountno.SetFocus
2510          Case False
2520            lngRecsCur = RecCnt  ' ** Function: Below.
2530            If .CurrentRecord < lngRecsCur Then
2540              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2550              .invobj_name.SetFocus
2560            Else
2570              DoCmd.SelectObject acForm, .Parent.Name, False
2580              If .Parent.chkShowAccounts.Enabled = True Then
2590                .Parent.chkShowAccounts.SetFocus
2600              ElseIf .Parent.chkShowUnassigned.Enabled = True Then
2610                .Parent.chkShowUnassigned.SetFocus
2620              ElseIf .Parent.cmdPrintReport.Enabled = True Then
2630                .Parent.cmdPrintReport.SetFocus
2640              Else
2650                .Parent.cmdClose.SetFocus
2660              End If
2670            End If
2680          End Select
2690        End With
2700      End Select
2710    End If

        ' ** Shift keys.
2720    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2730      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2740        With Me
2750          intRetVal = 0
2760          If .CurrentRecord > 1 Then
2770            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2780            Select Case .Parent.chkShowAccounts
                Case True
2790              .shortname.SetFocus
2800            Case False
2810              .invobj_name.SetFocus
2820            End Select
2830          Else
2840            DoCmd.SelectObject acForm, .Parent.Name, False
2850            .Parent.cmdClose.SetFocus
2860          End If
2870        End With
2880      End Select
2890    End If

EXITP:
2900    KeyCode = intRetVal
2910    Exit Sub

ERRH:
2920    intRetVal = 0
2930    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2940    Case Else
2950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2960    End Select
2970    Resume EXITP

End Sub

Private Sub cmbInvestmentObjectives_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbInvestmentObjectives_lbl_DblClick"

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

Private Sub cmbInvestmentObjectives_GotFocus()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbInvestmentObjectives_GotFocus"

3110    With Me
3120      .cmbInvestmentObjectives.SelLength = 0
3130      .cmbInvestmentObjectives.SelStart = 19
3140    End With

EXITP:
3150    Exit Sub

ERRH:
3160    Select Case ERR.Number
        Case Else
3170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3180    End Select
3190    Resume EXITP

End Sub

Private Sub cmbInvestmentObjectives_AfterUpdate()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbInvestmentObjectives_AfterUpdate"


3210    With Me
3220      cmdSave_Click  ' ** Procedure: Above.
3230      DoEvents
3240      RecalcTots  ' ** Procedure: Below.
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

Private Sub cmbInvestmentObjectives_NotInList(NewData As String, Response As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbInvestmentObjectives_NotInList"

3410    With Me
3420      .cmbInvestmentObjectives = Null
3430      MsgBox "Please select a valid objective from the list.", vbInformation + vbOKOnly, "Invalid Entry"
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

Private Sub cmbInvestmentObjectives_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbInvestmentObjectives_KeyDown"

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
3690          lngRecsCur = RecCnt  ' ** Function: Below.
3700          If .CurrentRecord < lngRecsCur Then
3710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below
3720            .cmbInvestmentObjectives.SetFocus
3730          Else
3740            DoCmd.SelectObject acForm, .Parent.Name, False
3750            If .Parent.chkShowAccounts.Enabled = True Then
3760              .Parent.chkShowAccounts.SetFocus
3770            ElseIf .Parent.chkShowUnassigned.Enabled = True Then
3780              .Parent.chkShowUnassigned.SetFocus
3790            ElseIf .Parent.cmdPrintReport.Enabled = True Then
3800              .Parent.cmdPrintReport.SetFocus
3810            Else
3820              .Parent.cmdClose.SetFocus
3830            End If
3840          End If
3850        End With
3860      End Select
3870    End If

        ' ** Shift keys.
3880    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3890      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3900        With Me
3910          intRetVal = 0
3920          If .CurrentRecord > 1 Then
3930            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3940            .cmbInvestmentObjectives.SetFocus
3950          Else
3960            DoCmd.SelectObject acForm, .Parent.Name, False
3970            .Parent.cmdClose.SetFocus
3980          End If
3990        End With
4000      End Select
4010    End If

EXITP:
4020    KeyCode = intRetVal
4030    Exit Sub

ERRH:
4040    intRetVal = 0
4050    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4060    Case Else
4070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4080    End Select
4090    Resume EXITP

End Sub

Private Sub acct_count_lbl_DblClick(Cancel As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "acct_count_lbl_DblClick"

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

Private Sub accountno_lbl_DblClick(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

4210    With Me
4220      If .Parent.chkShowAccounts = True Or .Parent.chkShowUnassigned = True Then
4230        SortNow THIS_PROC  ' ** Procedure: Below.
4240      End If
4250    End With

EXITP:
4260    Exit Sub

ERRH:
4270    Select Case ERR.Number
        Case Else
4280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4290    End Select
4300    Resume EXITP

End Sub

Private Sub accountno_GotFocus()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_GotFocus"

4410    With Me
4420      .accountno.SelLength = 0
4430      .accountno.SelStart = 19
4440    End With

EXITP:
4450    Exit Sub

ERRH:
4460    Select Case ERR.Number
        Case Else
4470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4480    End Select
4490    Resume EXITP

End Sub

Private Sub accountno_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4520    intShiftDown = (Shift And acShiftMask) > 0
4530    intAltDown = (Shift And acAltMask) > 0
4540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4570        With Me
4580          intRetVal = 0
4590          .shortname.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4650        With Me
4660          intRetVal = 0
4670          .invobj_name.SetFocus
4680        End With
4690      End Select
4700    End If

EXITP:
4710    KeyCode = intRetVal
4720    Exit Sub

ERRH:
4730    intRetVal = 0
4740    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4750    Case Else
4760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4770    End Select
4780    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

4810    With Me
4820      If .Parent.chkShowAccounts = True Or .Parent.chkShowUnassigned = True Then
4830        SortNow THIS_PROC  ' ** Procedure: Below.
4840      End If
4850    End With

EXITP:
4860    Exit Sub

ERRH:
4870    Select Case ERR.Number
        Case Else
4880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4890    End Select
4900    Resume EXITP

End Sub

Private Sub shortname_GotFocus()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_GotFocus"

5010    With Me
5020      .shortname.SelLength = 0
5030      .shortname.SelStart = 0
5040    End With

EXITP:
5050    Exit Sub

ERRH:
5060    Select Case ERR.Number
        Case Else
5070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5080    End Select
5090    Resume EXITP

End Sub

Private Sub shortname_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_KeyDown"

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
5190          lngRecsCur = RecCnt  ' ** Function: Below.
5200          If .CurrentRecord < lngRecsCur Then
5210            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5220            .invobj_name.SetFocus
5230          Else
5240            DoCmd.SelectObject acForm, .Parent.Name, False
5250            If .Parent.chkShowAccounts.Enabled = True Then
5260              .Parent.chkShowAccounts.SetFocus
5270            ElseIf .Parent.chkShowUnassigned.Enabled = True Then
5280              .Parent.chkShowUnassigned.SetFocus
5290            ElseIf .Parent.cmdPrintReport.Enabled = True Then
5300              .Parent.cmdPrintReport.SetFocus
5310            Else
5320              .Parent.cmdClose.SetFocus
5330            End If
5340          End If
5350        End With
5360      End Select
5370    End If

        ' ** Shift keys.
5380    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5390      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5400        With Me
5410          intRetVal = 0
5420          .accountno.SetFocus
5430        End With
5440      End Select
5450    End If

EXITP:
5460    KeyCode = intRetVal
5470    Exit Sub

ERRH:
5480    intRetVal = 0
5490    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5500    Case Else
5510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5520    End Select
5530    Resume EXITP

End Sub

Public Sub ShowAccounts(blnShow As Boolean)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAccounts"

        Dim blnSortHere As Boolean, blnResort As Boolean

5610    With Me
5620      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5630        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5640      End If
5650      blnSortHere = False: blnResort = False
5660      .FocusHolder.SetFocus
5670      DoEvents
5680      Select Case blnShow
          Case True
5690        If .Sort_line.Left = .acct_count_lbl.Left Then blnSortHere = True
5700        .acct_count.Visible = False
5710        .acct_count_lbl.Visible = False
5720        .acct_count_lbl2.Visible = False
5730        .acct_count_lbl_line.Visible = False
5740        If blnSortHere = True Then
5750          .Sort_line.Visible = False
5760          .Sort_lbl.Visible = False
5770          blnResort = True
5780        End If
5790        blnSortHere = False
5800        If .Sort_line.Left = .invobj_name_lbl.Left Then blnSortHere = True
5810        .invobj_name.Width = ((.accountno.Left - .invobj_name.Left) - (4& * lngTpp))
5820        .invobj_name_lbl.Width = .invobj_name.Width
5830        .invobj_name_lbl_line.Width = (.invobj_name_lbl.Width + lngTpp)
5840        If blnSortHere = True Then
5850          .Sort_line.Width = (.invobj_name_lbl.Width + lngTpp)
5860          .Sort_lbl.Left = ((.invobj_name_lbl.Left + .invobj_name_lbl.Width) - lngSortLbl_Width)
5870        End If
5880        blnSortHere = False
5890        .accountno.Visible = True
5900        .accountno_lbl.Visible = True
5910        .accountno_lbl_line.Visible = True
5920        .shortname.Visible = True
5930        .shortname_lbl.Visible = True
5940        .shortname_lbl_line.Visible = True
5950        .RecordSource = "qryInvestmentObjective_04"
5960      Case False
5970        If .Sort_line.Left = .accountno_lbl.Left Then blnSortHere = True
5980        .accountno.Visible = False
5990        .accountno_lbl.Visible = False
6000        .accountno_lbl_line.Visible = False
6010        If blnSortHere = True Then
6020          .Sort_line.Visible = False
6030          .Sort_lbl.Visible = False
6040          blnResort = True
6050        End If
6060        blnSortHere = False
6070        If .Sort_line.Left = .shortname_lbl.Left Then blnSortHere = True
6080        .shortname.Visible = False
6090        .shortname_lbl.Visible = False
6100        .shortname_lbl_line.Visible = False
6110        If blnSortHere = True Then
6120          .Sort_line.Visible = False
6130          .Sort_lbl.Visible = False
6140          blnResort = True
6150        End If
6160        blnSortHere = False
6170        .acct_count.Visible = True
6180        .acct_count_lbl.Visible = True
6190        .acct_count_lbl2.Visible = True
6200        .acct_count_lbl_line.Visible = True
6210        If .Sort_line.Left = .invobj_name_lbl.Left Then blnSortHere = True
6220        .invobj_name.Width = lngObjName_Width
6230        .invobj_name_lbl.Width = lngObjName_Width
6240        .invobj_name_lbl_line.Width = (lngObjName_Width + lngTpp)
6250        If blnSortHere = True Then
6260          .Sort_line.Width = (.invobj_name_lbl.Width + lngTpp)
6270          .Sort_lbl.Left = ((.invobj_name_lbl.Left + .invobj_name_lbl.Width) - lngSortLbl_Width)
6280        End If
6290        blnSortHere = False
6300        .RecordSource = "qryInvestmentObjective_03"
6310      End Select
6320    End With

EXITP:
6330    Exit Sub

ERRH:
6340    Select Case ERR.Number
        Case Else
6350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6360    End Select
6370    Resume EXITP

End Sub

Public Sub ShowUnassigned(blnShow As Boolean)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowUnassigned"

        Dim blnSortHere As Boolean, blnResort As Boolean

6410    With Me
6420      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
6430        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6440      End If
6450      blnSortHere = False: blnResort = False
6460      .FocusHolder.SetFocus
6470      DoEvents
6480      Select Case blnShow
          Case True
6490        If .Sort_line.Left = .invobj_id_lbl.Left Then blnSortHere = True
6500        .invobj_id.Visible = False
6510        .invobj_id_lbl.Visible = False
6520        .invobj_id_lbl_line.Visible = False
6530        If blnSortHere = True Then
6540          .Sort_line.Visible = False
6550          .Sort_lbl.Visible = False
6560          blnResort = True
6570        End If
6580        blnSortHere = False
6590        If .Sort_line.Left = .invobj_name_lbl.Left Then blnSortHere = True
6600        .invobj_name.Visible = False
6610        .invobj_name_lbl.Visible = False
6620        .invobj_name_lbl_line.Visible = False
6630        If blnSortHere = True Then
6640          .Sort_line.Visible = False
6650          .Sort_lbl.Visible = False
6660          blnResort = True
6670        End If
6680        blnSortHere = False
6690        If .Sort_line.Left = .acct_count_lbl.Left Then blnSortHere = True
6700        .acct_count.Visible = False
6710        .acct_count_lbl.Visible = False
6720        .acct_count_lbl2.Visible = False
6730        .acct_count_lbl_line.Visible = False
6740        If blnSortHere = True Then
6750          .Sort_line.Visible = False
6760          .Sort_lbl.Visible = False
6770          blnResort = True
6780        End If
6790        blnSortHere = False
6800        .cmbInvestmentObjectives.Visible = True
6810        .cmbInvestmentObjectives_lbl.Visible = True
6820        .cmbInvestmentObjectives_lbl_line.Visible = True
6830        .accountno.Visible = True
6840        .accountno_lbl.Visible = True
6850        .accountno_lbl_line.Visible = True
6860        .shortname.Visible = True
6870        .shortname_lbl.Visible = True
6880        .shortname_lbl_line.Visible = True
6890      Case False
6900        If .Sort_line.Left = .cmbInvestmentObjectives_lbl.Left Then blnSortHere = True
6910        .cmbInvestmentObjectives.Visible = False
6920        .cmbInvestmentObjectives_lbl.Visible = False
6930        .cmbInvestmentObjectives_lbl_line.Visible = False
6940        If blnSortHere = True Then
6950          .Sort_line.Visible = False
6960          .Sort_lbl.Visible = True
6970          blnResort = True
6980        End If
6990        blnSortHere = False
7000        .invobj_id.Visible = True
7010        .invobj_id_lbl.Visible = True
7020        .invobj_id_lbl_line.Visible = True
7030        Select Case .Parent.chkShowAccounts
            Case True
7040          .acct_count.Visible = False
7050          .acct_count_lbl.Visible = False
7060          .acct_count_lbl2.Visible = False
7070          .acct_count_lbl_line.Visible = False
7080          .invobj_name.Width = ((.accountno.Left - .invobj_name.Left) - (4& * lngTpp))
7090          .invobj_name_lbl.Width = .invobj_name.Width
7100          .invobj_name_lbl_line.Width = (.invobj_name_lbl.Width + lngTpp)
7110          .invobj_name.Visible = True
7120          .invobj_name_lbl.Visible = True
7130          .invobj_name_lbl_line.Visible = True
7140          .accountno.Visible = True
7150          .accountno_lbl.Visible = True
7160          .accountno_lbl_line.Visible = True
7170          .shortname.Visible = True
7180          .shortname_lbl.Visible = True
7190          .shortname_lbl_line.Visible = True
7200        Case False
7210          If .Sort_line.Left = .accountno_lbl.Left Then blnSortHere = True
7220          .accountno.Visible = False
7230          .accountno_lbl.Visible = False
7240          .accountno_lbl_line.Visible = False
7250          If blnSortHere = True Then
7260            .Sort_line.Visible = False
7270            .Sort_lbl.Visible = False
7280            blnResort = True
7290          End If
7300          blnSortHere = False
7310          If .Sort_line.Left = .shortname_lbl.Left Then blnSortHere = True
7320          .shortname.Visible = False
7330          .shortname_lbl.Visible = False
7340          .shortname_lbl_line.Visible = False
7350          If blnSortHere = True Then
7360            .Sort_line.Visible = False
7370            .Sort_lbl.Visible = False
7380            blnResort = True
7390          End If
7400          blnSortHere = False
7410          .invobj_name.Width = lngObjName_Width
7420          .invobj_name_lbl.Width = .invobj_name.Width
7430          .invobj_name_lbl_line.Width = (.invobj_name_lbl.Width + lngTpp)
7440          .invobj_name.Visible = True
7450          .invobj_name_lbl.Visible = True
7460          .invobj_name_lbl_line.Visible = True
7470          .acct_count.Visible = True
7480          .acct_count_lbl.Visible = True
7490          .acct_count_lbl2.Visible = True
7500          .acct_count_lbl_line.Visible = True
7510        End Select
7520      End Select
7530    End With

EXITP:
7540    Exit Sub

ERRH:
7550    Select Case ERR.Number
        Case Else
7560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7570    End Select
7580    Resume EXITP

End Sub

Public Sub RecalcTots()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim varTmp00 As Variant, strTmp01 As String

7610    With Me
          ' ** tblInvestmentObjective_Staging, grouped by invobj_name, just '{Unassigned}', with cnt_null.
7620      varTmp00 = DLookup("[cnt_null]", "qryInvestmentObjective_11")
7630      If IsNull(varTmp00) = True Then
7640        varTmp00 = 0
7650      End If
7660      If varTmp00 = 0 Then
7670        strTmp01 = "0 Accounts Unassigned"
7680      ElseIf varTmp00 = 1 Then
7690        strTmp01 = "1 Account Unassigned"
7700      Else
7710        strTmp01 = CStr(varTmp00) & " Accounts Unassigned"
7720      End If
7730      .Parent.Unassigned_lbl.Caption = strTmp01
7740      DoEvents
7750    End With

EXITP:
7760    Exit Sub

ERRH:
7770    Select Case ERR.Number
        Case Else
7780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7790    End Select
7800    Resume EXITP

End Sub

Public Function RecCnt() As Long

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

7910    lngRetVal = 0&

7920    With Me
7930      Set rst = .RecordsetClone
7940      With rst
7950        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
7960        Else
7970          .MoveLast
7980          lngRetVal = .RecordCount
7990        End If
8000        .Close
8010      End With
8020    End With

EXITP:
8030    Set rst = Nothing
8040    RecCnt = lngRetVal
8050    Exit Function

ERRH:
8060    lngRetVal = 0&
8070    Select Case ERR.Number
        Case Else
8080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8090    End Select
8100    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

8210    Select Case intWhere
        Case acCmdRecordsGoToFirst
8220      DoCmd.RunCommand acCmdRecordsGoToFirst
8230    Case acCmdRecordsGoToLast
8240      DoCmd.RunCommand acCmdRecordsGoToLast
8250    Case acCmdRecordsGoToPrevious
8260      DoCmd.RunCommand acCmdRecordsGoToPrevious
8270    Case acCmdRecordsGoToNext
8280      DoCmd.RunCommand acCmdRecordsGoToNext
8290    Case Else
8300      If IsMissing(varID) = False Then
8310        With Me
8320          Set rst = .RecordsetClone
8330          With rst
8340            .FindFirst "[invobj_id] = " & CStr(varID)
8350            If .NoMatch = False Then
8360              Me.Bookmark = .Bookmark
8370            End If
8380            .Close
8390          End With
8400        End With
8410      End If
8420    End Select

EXITP:
8430    Set rst = Nothing
8440    Exit Sub

ERRH:
8450    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8460    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8470    Case Else
8480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8490    End Select
8500    Resume EXITP

End Sub

Public Sub SortNow(strProc As String, Optional varShow As Variant, Optional varIsDesc As Variant)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim blnAccts As Boolean, blnDontReverse As Boolean, blnIsDesc As Boolean
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [invobj_id]"

8610    With Me
8620      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
8630        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
8640      End If
8650      blnAccts = .Parent.chkShowAccounts
8660      Select Case IsMissing(varShow)
          Case True
8670        blnDontReverse = False
8680      Case False
8690        blnDontReverse = CBool(varShow)
8700      End Select
8710      Select Case IsMissing(varIsDesc)
          Case True
8720        blnIsDesc = False
8730      Case False
8740        blnIsDesc = CBool(varIsDesc)
8750      End Select
8760      .Controls(strSortLbl).Visible = False
8770      .Controls(strSortLine).Visible = False
8780      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the the right side of the form.
8790      If strProc = "Form_Load" Then
8800        strCalled = "invobj_id"
8810        strSortNow = strSortOrig
8820        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
8830        lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (4& * lngTpp))
8840        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
8850        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
8860        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
8870        .Controls(strSortLbl).Top = lngSortLbl_Top
8880        .Controls(strSortLbl).Left = lngSortLbl_Left
8890        .Controls(strSortLine).Top = lngSortLine_Top
8900        .Controls(strSortLine).Left = lngSortLine_Left
8910        .Controls(strSortLine).Width = lngSortLine_Width
8920        .Controls(strSortLbl).Caption = strArwUp
8930        .Controls(strSortLbl).ForeColor = CLR_DKBLU
8940      Else
8950        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
8960        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
8970        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
8980        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
8990        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
9000        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
9010        .Controls(strSortLbl).Top = lngSortLbl_Top
9020        .Controls(strSortLbl).Left = lngSortLbl_Left
9030        .Controls(strSortLine).Top = lngSortLine_Top
9040        .Controls(strSortLine).Left = lngSortLine_Left
9050        .Controls(strSortLine).Width = lngSortLine_Width
9060        If strCalled = "invobj_id" Then
9070          strSortAsc = strSortOrig
9080          If ((strSortNow = strSortAsc And blnDontReverse = False) Or (blnIsDesc = True)) Then
9090            strSortNow = "[invobj_id] DESC"
9100            .Controls(strSortLbl).Caption = strArwDn
9110            .Controls(strSortLbl).ForeColor = CLR_DKRED
9120          Else
9130            strSortNow = strSortAsc
9140            .Controls(strSortLbl).Caption = strArwUp
9150            .Controls(strSortLbl).ForeColor = CLR_DKBLU
9160          End If
9170          lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
9180          .Controls(strSortLbl).Left = lngSortLbl_Left
9190        Else
9200          .Controls(strSortLbl).Caption = strArwUp
9210          .Controls(strSortLbl).ForeColor = CLR_DKBLU
9220          Select Case strCalled
              Case "invobj_name"
9230            strSortAsc = "[invobj_name], [accountno]"
9240            If blnDontReverse = True Then
9250              strSortNow = vbNullString
9260            End If
9270          Case "cmbInvestmentObjectives"
9280            strSortAsc = "[invobj_name], [accountno]"
9290          Case "acct_count"
9300            strSortAsc = "[acct_count]" & strStdAsc
9310            lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
9320            .Controls(strSortLbl).Left = lngSortLbl_Left
9330          Case "accountno"
9340            strSortAsc = "[alphasort]"
9350          Case "shortname"
9360            strSortAsc = "[shortname]"
9370          End Select
9380          If ((strSortNow = strSortAsc And blnDontReverse = False) Or (blnIsDesc = True)) Then
9390            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
9400            Select Case intCnt
                Case 1
9410              strTmp01 = strSortAsc & " DESC"
9420            Case 2
9430              intPos01 = InStr(strSortAsc, ",")
9440              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9450              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
9460            Case 3
9470              intPos01 = InStr(strSortAsc, ",")
9480              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9490              strTmp02 = Mid(strSortAsc, intPos01)
9500              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9510              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
9520            End Select
9530            strSortNow = strTmp01
9540            .Controls(strSortLbl).Caption = strArwDn
9550            .Controls(strSortLbl).ForeColor = CLR_DKRED
9560          Else
9570            strSortNow = strSortAsc
9580          End If
9590        End If
9600      End If
9610      .Controls(strSortLbl).Visible = True
9620      .Controls(strSortLine).Visible = True
9630      .OrderBy = strSortNow
9640      .OrderByOn = True

9650      Select Case .Parent.chkShowUnassigned
          Case True
9660        .Parent.SortNow_Set strSortNow, 3
9670      Case False
9680        Select Case .Parent.chkShowAccounts
            Case True
9690          .Parent.SortNow_Set strSortNow, 2
9700        Case False
9710          .Parent.SortNow_Set strSortNow, 1
9720        End Select
9730      End Select

9740    End With

        ' ** Potential Sorts:
        ' **   [invobj_id]
        ' **   [invobj_name], [accountno]
        ' **   [acct_count], [invobj_id]
        ' **   [accountno]
        ' **   [shortname]

EXITP:
9750    Exit Sub

ERRH:
9760    Select Case ERR.Number
        Case Else
9770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9780    End Select
9790    Resume EXITP

End Sub

Public Function SortNow_Get() As String

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

9810    strRetVal = strSortNow

EXITP:
9820    SortNow_Get = strRetVal
9830    Exit Function

ERRH:
9840    strRetVal = vbNullString
9850    Select Case ERR.Number
        Case Else
9860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9870    End Select
9880    Resume EXITP

End Function
