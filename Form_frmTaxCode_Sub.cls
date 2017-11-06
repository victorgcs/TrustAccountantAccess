VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTaxCode_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmTaxCode_Sub"

'VGC 09/02/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmTaxCode}
' **   Description:      D {taxcode_description}
' **   Print Report:     P {cmdPrintReport on frmTaxCode}
' **   Tax Code:         T {taxcode_inc, taxcode_ded}
' **   Exit:             X {cmdClose on frmTaxCode}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Toggle RevCodes:  R {ShowRevCodes}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[taxcode_type], [taxcode_order]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

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
150       If strCallingForm = "frmTaxCode" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170         lngSortLbl_Width = .Sort_lbl.Width

180         CLR_DISABLED_FG = CLR_DKGRY
190         CLR_DISABLED_BG = CLR_LTTEAL

200       Else
210         Cancel = -1
220       End If
230     End With

240     If Cancel = -1 Then
250       Beep
260       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
270     End If

EXITP:
280     Exit Sub

ERRH:
290     Select Case ERR.Number
        Case Else
300       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
310     End Select
320     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .taxcode_inc.ForeColor = CLR_DISABLED_FG
430       .taxcode_inc.BackColor = CLR_DISABLED_BG
440       .taxcode_ded.ForeColor = CLR_DISABLED_FG
450       .taxcode_ded.BackColor = CLR_DISABLED_BG
460       .taxcode_description.ForeColor = CLR_DISABLED_FG
470       .taxcode_description.BackColor = CLR_DISABLED_BG
480       .revcode_TYPE.ForeColor = CLR_DISABLED_FG
490       .revcode_TYPE.BackColor = CLR_DISABLED_BG
500       Select Case gblnLinkRevTaxCodes
          Case True
510         .revcode_TYPE.Visible = True
520         .revcode_TYPE_lbl.Visible = True
530         .revcode_TYPE_lbl2.Visible = True
540         .revcode_TYPE_lbl_line.Visible = True
550         .revcode_TYPE_lbl_dim.Visible = False
560         .revcode_TYPE_lbl2_dim.Visible = False
570         .revcode_TYPE_lbl_line_dim.Visible = False
580         .revcode_TYPE_lbl_dim_hi.Visible = False
590         .revcode_TYPE_lbl2_dim_hi.Visible = False
600         .revcode_TYPE_lbl_line_dim_hi.Visible = False
610       Case False
620         .revcode_TYPE.Visible = False
630         .revcode_TYPE_lbl.Visible = False
640         .revcode_TYPE_lbl2.Visible = False
650         .revcode_TYPE_lbl_line.Visible = False
660         .revcode_TYPE_lbl_dim.Visible = True
670         .revcode_TYPE_lbl2_dim.Visible = True
680         .revcode_TYPE_lbl_line_dim.Visible = True
690         .revcode_TYPE_lbl_dim_hi.Visible = True
700         .revcode_TYPE_lbl2_dim_hi.Visible = True
710         .revcode_TYPE_lbl_line_dim_hi.Visible = True
720       End Select
730     End With

740     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
750     Exit Sub

ERRH:
760     Select Case ERR.Number
        Case Else
770       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
780     End Select
790     Resume EXITP

End Sub

Private Sub Form_Current()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

810     With Me
820       If .taxcode = TAXID_INC Or .taxcode = TAXID_DED Then
830         .taxcode_active.Locked = True
840       Else
850         .taxcode_active.Locked = False
860       End If
870     End With

EXITP:
880     Exit Sub

ERRH:
890     Select Case ERR.Number
        Case Else
900       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
910     End Select
920     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1020    intShiftDown = (Shift And acShiftMask) > 0
1030    intAltDown = (Shift And acAltMask) > 0
1040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Description:      D {taxcode_description}
        ' **   Tax Code:         T {taxcode_inc, taxcode_ded}
        ' **   Exit:             X {cmdClose on frmTaxCode}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1060      Select Case intRetVal
          Case vbKeyUp
1070        intRetVal = 0
1080        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1090      Case vbKeyDown
1100        intRetVal = 0
1110        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1120      Case vbKeyEscape
1130        With Me
1140          intRetVal = 0
1150          .Parent.cmdClose_Click  ' ** Form Procedure: frmTaxCode.
1160        End With
1170      End Select
1180    End If

        ' ** Alt keys.
1190    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1200      Select Case intRetVal
          Case vbKeyD
1210        With Me
1220          intRetVal = 0
1230          .taxcode_description.SetFocus
1240        End With
1250      Case vbKeyT
1260        With Me
1270          intRetVal = 0
1280          Select Case .taxcode_type
              Case TAXTYP_INC
1290            .taxcode_inc.SetFocus
1300          Case TAXTYP_DED
1310            .taxcode_ded.SetFocus
1320          End Select
1330        End With
1340      Case vbKeyX
1350        With Me
1360          intRetVal = 0
1370          .Parent.cmdClose_Click  ' ** Form Procedure: frmTaxCode.
1380        End With
1390      End Select
1400    End If

        ' ** Ctrl keys.
1410    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1420      Select Case intRetVal
          Case vbKeyS
1430        intRetVal = 0
1440        cmdSave_Click  ' ** Procedure: Below.
1450      Case vbKeyTab, vbKeyReturn
1460        With Me
1470          intRetVal = 0
1480          DoCmd.SelectObject acForm, .Parent.Name, False
1490          If .Parent.cmdPrintReport.Enabled = True Then
1500            .Parent.cmdPrintReport.SetFocus
1510          Else
1520            .Parent.cmdClose.SetFocus
1530          End If
1540        End With
1550      End Select
1560    End If

        ' ** Ctrl-Shift keys.
1570    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1580      Select Case intRetVal
          Case vbKeyF
1590        With Me
1600          intRetVal = 0
1610          DoCmd.SelectObject acForm, .Parent.Name, False
1620          .Parent.FocusHolder.SetFocus
1630        End With
1640      Case vbKeyR
1650        intRetVal = 0
1660        ShowRevCodes  ' ** Procedure: Below.
1670      Case vbKeyTab, vbKeyReturn
1680        With Me
1690          intRetVal = 0
1700          DoCmd.SelectObject acForm, .Parent.Name, False
1710          .Parent.cmdClose.SetFocus
1720        End With
1730      End Select
1740    End If

EXITP:
1750    KeyCode = intRetVal
1760    Exit Sub

ERRH:
1770    intRetVal = 0
1780    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1790    Case Else
1800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1810    End Select
1820    Resume EXITP

End Sub

Public Sub cmdSave_Click()
' ** This is a fake click event!

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1910    With Me
1920      DoCmd.RunCommand acCmdSelectRecord
1930      DoCmd.RunCommand acCmdSaveRecord
1940      .taxcode_active.SetFocus
1950    End With

EXITP:
1960    Exit Sub

ERRH:
1970    Select Case ERR.Number
        Case Else
1980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1990    End Select
2000    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2120    intShiftDown = (Shift And acShiftMask) > 0
2130    intAltDown = (Shift And acAltMask) > 0
2140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2170        With Me
2180          intRetVal = 0
2190          lngRecsCur = RecCnt  ' ** Function: Below.
2200          If .CurrentRecord < lngRecsCur Then
2210            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2220            .FocusHolder.SetFocus
2230          Else
2240            DoCmd.SelectObject acForm, .Parent.Name, False
2250            If .Parent.cmdPrintReport.Enabled = True Then
2260              .Parent.cmdPrintReport.SetFocus
2270            Else
2280              .Parent.cmdClose.SetFocus
2290            End If
2300          End If
2310        End With
2320      End Select
2330    End If

        ' ** Shift keys.
2340    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2350      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2360        With Me
2370          intRetVal = 0
2380          If .CurrentRecord > 1 Then
2390            MoveRec acCmdRecordsGoToPrevious
2400            .FocusHolder.SetFocus
2410          Else
2420            DoCmd.SelectObject acForm, .Parent.Name, False
2430            .Parent.cmdClose.SetFocus
2440          End If
2450        End With
2460      End Select
2470    End If

EXITP:
2480    KeyCode = intRetVal
2490    Exit Sub

ERRH:
2500    intRetVal = 0
2510    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2520    Case Else
2530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2540    End Select
2550    Resume EXITP

End Sub

Private Sub taxcode_type_lbl_DblClick(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_type_lbl_DblClick"

2610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub opgType_optIncome_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optIncome_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2770        With Me
2780          intRetVal = 0
2790          Select Case .taxcode_type
              Case TAXTYP_INC
2800            .taxcode_inc.SetFocus
2810          Case TAXTYP_DED
2820            .taxcode_ded.SetFocus
2830          End Select
2840        End With
2850      End Select
2860    End If

        ' ** Shift keys.
2870    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2880      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2890        With Me
2900          intRetVal = 0
2910          If .CurrentRecord > 1 Then
2920            MoveRec acCmdRecordsGoToPrevious
2930            .FocusHolder.SetFocus
2940          Else
2950            DoCmd.SelectObject acForm, .Parent.Name, False
2960            .Parent.cmdClose.SetFocus
2970          End If
2980        End With
2990      End Select
3000    End If

EXITP:
3010    KeyCode = intRetVal
3020    Exit Sub

ERRH:
3030    intRetVal = 0
3040    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3050    Case Else
3060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3070    End Select
3080    Resume EXITP

End Sub

Private Sub opgType_optDeduction_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optDeduction_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3120    intShiftDown = (Shift And acShiftMask) > 0
3130    intAltDown = (Shift And acAltMask) > 0
3140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3170        With Me
3180          intRetVal = 0
3190          Select Case .taxcode_type
              Case TAXTYP_INC
3200            .taxcode_inc.SetFocus
3210          Case TAXTYP_DED
3220            .taxcode_ded.SetFocus
3230          End Select
3240        End With
3250      End Select
3260    End If

        ' ** Shift keys.
3270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3280      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3290        With Me
3300          intRetVal = 0
3310          If .CurrentRecord > 1 Then
3320            MoveRec acCmdRecordsGoToPrevious
3330            .FocusHolder.SetFocus
3340          Else
3350            DoCmd.SelectObject acForm, .Parent.Name, False
3360            .Parent.cmdClose.SetFocus
3370          End If
3380        End With
3390      End Select
3400    End If

EXITP:
3410    KeyCode = intRetVal
3420    Exit Sub

ERRH:
3430    intRetVal = 0
3440    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3450    Case Else
3460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3470    End Select
3480    Resume EXITP

End Sub

Private Sub taxcode_lbl_DblClick(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_lbl_DblClick"

3510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub taxcode_inc_GotFocus()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_inc_GotFocus"

3610    With Me
3620      .taxcode_inc.SelLength = 0
3630      .taxcode_inc.SelStart = 9
3640    End With

EXITP:
3650    Exit Sub

ERRH:
3660    Select Case ERR.Number
        Case Else
3670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3680    End Select
3690    Resume EXITP

End Sub

Private Sub taxcode_ded_GotFocus()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_ded_GotFocus"

3710    With Me
3720      .taxcode_ded.SelLength = 0
3730      .taxcode_ded.SelStart = 9
3740    End With

EXITP:
3750    Exit Sub

ERRH:
3760    Select Case ERR.Number
        Case Else
3770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3780    End Select
3790    Resume EXITP

End Sub

Private Sub taxcode_inc_KeyDown(KeyCode As Integer, Shift As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_inc_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3820    intShiftDown = (Shift And acShiftMask) > 0
3830    intAltDown = (Shift And acAltMask) > 0
3840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3870        With Me
3880          intRetVal = 0
3890          .taxcode_description.SetFocus
3900        End With
3910      End Select
3920    End If

        ' ** Shift keys.
3930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3950        With Me
3960          intRetVal = 0
3970          .opgType.SetFocus
3980        End With
3990      End Select
4000    End If

EXITP:
4010    KeyCode = intRetVal
4020    Exit Sub

ERRH:
4030    intRetVal = 0
4040    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4050    Case Else
4060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4070    End Select
4080    Resume EXITP

End Sub

Private Sub taxcode_ded_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_ded_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4120    intShiftDown = (Shift And acShiftMask) > 0
4130    intAltDown = (Shift And acAltMask) > 0
4140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4170        With Me
4180          intRetVal = 0
4190          .taxcode_description.SetFocus
4200        End With
4210      End Select
4220    End If

        ' ** Shift keys.
4230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4250        With Me
4260          intRetVal = 0
4270          .opgType.SetFocus
4280        End With
4290      End Select
4300    End If

EXITP:
4310    KeyCode = intRetVal
4320    Exit Sub

ERRH:
4330    intRetVal = 0
4340    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4350    Case Else
4360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4370    End Select
4380    Resume EXITP

End Sub

Private Sub taxcode_description_lbl_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_lbl_DblClick"

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

Private Sub taxcode_description_GotFocus()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_GotFocus"

4510    With Me
4520      .taxcode_description.SelLength = 0
4530      .taxcode_description.SelStart = 99
4540    End With

EXITP:
4550    Exit Sub

ERRH:
4560    Select Case ERR.Number
        Case Else
4570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4580    End Select
4590    Resume EXITP

End Sub

Private Sub taxcode_description_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_KeyDown"

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
4680          intRetVal = 0
4690          .taxcode_active.SetFocus
4700        End With
4710      End Select
4720    End If

        ' ** Shift keys.
4730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4750        With Me
4760          intRetVal = 0
4770          Select Case .taxcode_type
              Case TAXTYP_INC
4780            .taxcode_inc.SetFocus
4790          Case TAXTYP_DED
4800            .taxcode_ded.SetFocus
4810          End Select
4820        End With
4830      End Select
4840    End If

EXITP:
4850    KeyCode = intRetVal
4860    Exit Sub

ERRH:
4870    intRetVal = 0
4880    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4890    Case Else
4900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4910    End Select
4920    Resume EXITP

End Sub

Private Sub taxcode_active_lbl_DblClick(Cancel As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_active_lbl_DblClick"

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

Private Sub taxcode_active_AfterUpdate()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_active_AfterUpdate"

5110    cmdSave_Click  ' ** Procedure: Above.

EXITP:
5120    Exit Sub

ERRH:
5130    Select Case ERR.Number
        Case Else
5140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5150    End Select
5160    Resume EXITP

End Sub

Private Sub taxcode_active_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_active_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5220    intShiftDown = (Shift And acShiftMask) > 0
5230    intAltDown = (Shift And acAltMask) > 0
5240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5270        With Me
5280          intRetVal = 0
5290          If .revcode_TYPE.Visible = True And .revcode_TYPE.Enabled = True Then
5300            .revcode_TYPE.SetFocus
5310          Else
5320            lngRecsCur = RecCnt  ' ** Function: Below.
5330            If .CurrentRecord < lngRecsCur Then
5340              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5350              .FocusHolder.SetFocus
5360            Else
5370              DoCmd.SelectObject acForm, .Parent.Name
5380              If .Parent.cmdPrintReport.Enabled = True Then
5390                .Parent.cmdPrintReport.SetFocus
5400              Else
5410                .Parent.cmdClose.SetFocus
5420              End If
5430            End If
5440          End If
5450        End With
5460      End Select
5470    End If

        ' ** Shift keys.
5480    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5490      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5500        With Me
5510          intRetVal = 0
5520          .taxcode_description.SetFocus
5530        End With
5540      End Select
5550    End If

EXITP:
5560    KeyCode = intRetVal
5570    Exit Sub

ERRH:
5580    intRetVal = 0
5590    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5600    Case Else
5610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5620    End Select
5630    Resume EXITP

End Sub

Private Sub revcode_TYPE_lbl_DblClick(Cancel As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_lbl_DblClick"

5710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5720    Exit Sub

ERRH:
5730    Select Case ERR.Number
        Case Else
5740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5750    End Select
5760    Resume EXITP

End Sub

Private Sub revcode_TYPE_GotFocus()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_GotFocus"

5810    With Me
5820      .revcode_TYPE.SelLength = 0
5830      .revcode_TYPE.SelStart = 19
5840    End With

EXITP:
5850    Exit Sub

ERRH:
5860    Select Case ERR.Number
        Case Else
5870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5880    End Select
5890    Resume EXITP

End Sub

Private Sub revcode_TYPE_KeyDown(KeyCode As Integer, Shift As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_TYPE_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5920    intShiftDown = (Shift And acShiftMask) > 0
5930    intAltDown = (Shift And acAltMask) > 0
5940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5970        With Me
5980          intRetVal = 0
5990          lngRecsCur = RecCnt  ' ** Function: Below.
6000          If .CurrentRecord < lngRecsCur Then
6010            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
6020            .FocusHolder.SetFocus
6030          Else
6040            DoCmd.SelectObject acForm, .Parent.Name, False
6050            If .Parent.cmdPrintReport.Enabled = True Then
6060              .Parent.cmdPrintReport.SetFocus
6070            Else
6080              .Parent.cmdClose.SetFocus
6090            End If
6100          End If
6110        End With
6120      End Select
6130    End If

        ' ** Shift keys.
6140    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6150      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6160        With Me
6170          intRetVal = 0
6180          .taxcode_active.SetFocus
6190        End With
6200      End Select
6210    End If

EXITP:
6220    KeyCode = intRetVal
6230    Exit Sub

ERRH:
6240    intRetVal = 0
6250    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6260    Case Else
6270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6280    End Select
6290    Resume EXITP

End Sub

Public Sub ShowRevCodes()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowRevCodes"

        Dim blnShow As Boolean

6310    With Me
6320      blnShow = (Not .revcode_TYPE.Visible)
6330      .FocusHolder.SetFocus
6340      Select Case blnShow
          Case True
6350        .revcode_TYPE.Visible = True
6360        .revcode_TYPE_lbl.Visible = True
6370        .revcode_TYPE_lbl2.Visible = True
6380        .revcode_TYPE_lbl_line.Visible = True
6390        .revcode_TYPE_lbl_dim.Visible = False
6400        .revcode_TYPE_lbl2_dim.Visible = False
6410        .revcode_TYPE_lbl_line_dim.Visible = False
6420        .revcode_TYPE_lbl_dim_hi.Visible = False
6430        .revcode_TYPE_lbl2_dim_hi.Visible = False
6440        .revcode_TYPE_lbl_line_dim_hi.Visible = False
6450      Case False
6460        .revcode_TYPE.Visible = False
6470        .revcode_TYPE_lbl.Visible = False
6480        .revcode_TYPE_lbl2.Visible = False
6490        .revcode_TYPE_lbl_line.Visible = False
6500        .revcode_TYPE_lbl_dim.Visible = True
6510        .revcode_TYPE_lbl2_dim.Visible = True
6520        .revcode_TYPE_lbl_line_dim.Visible = True
6530        .revcode_TYPE_lbl_dim_hi.Visible = True
6540        .revcode_TYPE_lbl2_dim_hi.Visible = True
6550        .revcode_TYPE_lbl_line_dim_hi.Visible = True
6560      End Select
6570    End With

EXITP:
6580    Exit Sub

ERRH:
6590    Select Case ERR.Number
        Case Else
6600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6610    End Select
6620    Resume EXITP

End Sub

Public Function RecCnt() As Long

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

6710    lngRetVal = 0&

6720    With Me
6730      Set rst = .RecordsetClone
6740      With rst
6750        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
6760        Else
6770          .MoveLast
6780          lngRetVal = .RecordCount
6790        End If
6800        .Close
6810      End With
6820    End With

EXITP:
6830    Set rst = Nothing
6840    RecCnt = lngRetVal
6850    Exit Function

ERRH:
6860    lngRetVal = 0&
6870    Select Case ERR.Number
        Case Else
6880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6890    End Select
6900    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

7010    Select Case intWhere
        Case acCmdRecordsGoToFirst
7020      DoCmd.RunCommand acCmdRecordsGoToFirst
7030    Case acCmdRecordsGoToLast
7040      DoCmd.RunCommand acCmdRecordsGoToLast
7050    Case acCmdRecordsGoToPrevious
7060      DoCmd.RunCommand acCmdRecordsGoToPrevious
7070    Case acCmdRecordsGoToNext
7080      DoCmd.RunCommand acCmdRecordsGoToNext
7090    Case Else
7100      If IsMissing(varID) = False Then
7110        With Me
7120          Set rst = .RecordsetClone
7130          With rst
7140            .FindFirst "[taxcode] = '" & varID & "'"
7150            If .NoMatch = False Then
7160              Me.Bookmark = .Bookmark
7170            End If
7180            .Close
7190          End With
7200        End With
7210      End If
7220    End Select

EXITP:
7230    Set rst = Nothing
7240    Exit Sub

ERRH:
7250    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
7260    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
7270    Case Else
7280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7290    End Select
7300    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [taxcode_type], [taxcode_order]"

7410    With Me
7420      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
7430        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7440      End If
7450      .Controls(strSortLbl).Visible = False
7460      .Controls(strSortLine).Visible = False
7470      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the the right side of the form.
7480      If strProc = "Form_Load" Then
7490        strCalled = "taxcode_type"
7500        strSortNow = strSortOrig
7510        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7520        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7530        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7540        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7550        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7560        .Controls(strSortLbl).Top = lngSortLbl_Top
7570        .Controls(strSortLbl).Left = lngSortLbl_Left
7580        .Controls(strSortLine).Top = lngSortLine_Top
7590        .Controls(strSortLine).Left = lngSortLine_Left
7600        .Controls(strSortLine).Width = lngSortLine_Width
7610        .Controls(strSortLbl).Caption = strArwUp
7620        .Controls(strSortLbl).ForeColor = CLR_DKBLU
7630      Else
7640        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
7650        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7660        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7670        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7680        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7690        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7700        .Controls(strSortLbl).Top = lngSortLbl_Top
7710        .Controls(strSortLbl).Left = lngSortLbl_Left
7720        .Controls(strSortLine).Top = lngSortLine_Top
7730        .Controls(strSortLine).Left = lngSortLine_Left
7740        .Controls(strSortLine).Width = lngSortLine_Width
7750        If strCalled = "taxcode_type" Then
7760          strSortAsc = strSortOrig
7770          If strSortNow = strSortAsc Then
7780            strSortNow = "[taxcode_type] DESC, [taxcode_order] DESC"
7790            .Controls(strSortLbl).Caption = strArwDn
7800            .Controls(strSortLbl).ForeColor = CLR_DKRED
7810          Else
7820            strSortNow = strSortAsc
7830            .Controls(strSortLbl).Caption = strArwUp
7840            .Controls(strSortLbl).ForeColor = CLR_DKBLU
7850          End If
7860        Else
7870          .Controls(strSortLbl).Caption = strArwUp
7880          .Controls(strSortLbl).ForeColor = CLR_DKBLU
7890          Select Case strCalled
              Case "taxcode"
7900            lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
7910            .Controls(strSortLbl).Left = lngSortLbl_Left
7920            strSortAsc = "[taxcode]"
7930          Case "taxcode_description"
7940            strSortAsc = "[taxcode_description]"
7950          Case "taxcode_active"
7960            strSortAsc = "[taxcode_active]" & strStdAsc
7970            lngSortLbl_Left = (lngSortLbl_Left + (10& * lngTpp))
7980            .Controls(strSortLbl).Left = lngSortLbl_Left
7990          Case "revcode_TYPE"
8000            strSortAsc = "[revcode_TYPE]" & strStdAsc
8010            lngSortLbl_Left = (lngSortLbl_Left + (1& * lngTpp))
8020            .Controls(strSortLbl).Left = lngSortLbl_Left
8030          End Select
8040          If strSortNow = strSortAsc Then
8050            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
8060            Select Case intCnt
                Case 1
8070              strTmp01 = strSortAsc & " DESC"
8080            Case 2
8090              intPos01 = InStr(strSortAsc, ",")
8100              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8110              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
8120            Case 3
8130              intPos01 = InStr(strSortAsc, ",")
8140              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8150              strTmp02 = Mid(strSortAsc, intPos01)
8160              intPos01 = InStr(2, strTmp02, ",")
8170              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
8180              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
8190            Case 4
8200              intPos01 = InStr(strSortAsc, ",")
8210              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8220              strTmp02 = Mid(strSortAsc, intPos01)
8230              intPos01 = InStr(2, strTmp02, ",")
8240              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
8250              strTmp02 = Mid(strTmp02, intPos01)
8260              intPos01 = InStr(2, strTmp02, ",")
8270              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
8280              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
8290            End Select
8300            strSortNow = strTmp01
8310            .Controls(strSortLbl).Caption = strArwDn
8320            .Controls(strSortLbl).ForeColor = CLR_DKRED
8330          Else
8340            strSortNow = strSortAsc
8350          End If
8360        End If
8370      End If
8380      .Controls(strSortLbl).Visible = True
8390      .Controls(strSortLine).Visible = True
8400      .OrderBy = strSortNow
8410      .OrderByOn = True
8420    End With

        ' ** Potential Sorts:
        ' **   [taxcode_type], [taxcode_order]
        ' **   [taxcode]
        ' **   [taxcode_description]
        ' **   [taxcode_order], [taxcode_type]
        ' **   [taxcode_active], [taxcode_type], [taxcode_order]
        ' **   [revcode_TYPE], [taxcode_type], [taxcode_order]

EXITP:
8430    Exit Sub

ERRH:
8440    Select Case ERR.Number
        Case Else
8450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8460    End Select
8470    Resume EXITP

End Sub

Public Function SortNow_Get() As String

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

8510    strRetVal = strSortNow

EXITP:
8520    SortNow_Get = strRetVal
8530    Exit Function

ERRH:
8540    strRetVal = vbNullString
8550    Select Case ERR.Number
        Case Else
8560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8570    End Select
8580    Resume EXITP

End Function
