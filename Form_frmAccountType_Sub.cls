VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountType_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountType_Sub"

'VGC 04/19/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmAccountType}
' **   Show Accounts:    H {chkShowAccounts on frmAccountType}
' **   Print Report:     P {cmdPrintReport on frmAccountType}
' **   Exit:             X {cmdClose on frmAccountType}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[accounttype]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngRecsCur As Long, lngTpp As Long
Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccountType" Then

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
420       .accounttype.ForeColor = CLR_DISABLED_FG
430       .accounttype.BackColor = CLR_DISABLED_BG
440       .accounttype_description.ForeColor = CLR_DISABLED_FG
450       .accounttype_description.BackColor = CLR_DISABLED_BG
460       .accounttypegroup_name.ForeColor = CLR_DISABLED_FG
470       .accounttypegroup_name.BackColor = CLR_DISABLED_BG
480       .accounttypegroup_sequence.ForeColor = CLR_DISABLED_FG
490       .accounttypegroup_sequence.BackColor = CLR_DISABLED_BG
500     End With

510     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
520     Exit Sub

ERRH:
530     Select Case ERR.Number
        Case Else
540       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
550     End Select
560     Resume EXITP

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

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyUp
670         intRetVal = 0
680         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
690       Case vbKeyDown
700         intRetVal = 0
710         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
720       Case vbKeyEscape
730         With Me
740           intRetVal = 0
750           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountAssets.
760         End With
770       End Select
780     End If

        ' ** Alt keys.
790     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
800       Select Case intRetVal
          Case vbKeyX
810         With Me
820           intRetVal = 0
830           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountAssets.
840         End With
850       End Select
860     End If

        ' ** Ctrl keys.
870     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
880       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
890         With Me
900           intRetVal = 0
910           DoCmd.SelectObject acForm, .Parent.Name, False
920           .Parent.chkShowAccounts
930         End With
940       End Select
950     End If

        ' ** Ctrl-Shift keys.
960     If intCtrlDown And (Not intAltDown) And intShiftDown Then
970       Select Case intRetVal
          Case vbKeyF
980         With Me
990           intRetVal = 0
1000          DoCmd.SelectObject acForm, .Parent.Name, False
1010          .Parent.FocusHolder.SetFocus
1020        End With
1030      Case vbKeyTab, vbKeyReturn
1040        With Me
1050          intRetVal = 0
1060          DoCmd.SelectObject acForm, .Parent.Name, False
1070          .Parent.cmdClose.SetFocus
1080        End With
1090      End Select
1100    End If

EXITP:
1110    KeyCode = intRetVal
1120    Exit Sub

ERRH:
1130    intRetVal = 0
1140    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1150    Case Else
1160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1170    End Select
1180    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1270        With Me
1280          intRetVal = 0
1290          lngRecsCur = RecCnt  ' ** Function: Below.
1300          If .CurrentRecord < lngRecsCur Then
1310            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1320            .FocusHolder.SetFocus
1330          Else
1340            DoCmd.SelectObject acForm, .Parent.Name, False
1350            .Parent.chkShowAccounts.SetFocus
1360          End If
1370        End With
1380      End Select
1390    End If

        ' ** Shift keys.
1400    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1410      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1420        With Me
1430          intRetVal = 0
1440          If .CurrentRecord > 1 Then
1450            MoveRec acCmdRecordsGoToPrevious
1460            .FocusHolder.SetFocus
1470          Else
1480            DoCmd.SelectObject acForm, .Parent.Name, False
1490            .Parent.cmdClose.SetFocus
1500          End If
1510        End With
1520      End Select
1530    End If

EXITP:
1540    KeyCode = intRetVal
1550    Exit Sub

ERRH:
1560    intRetVal = 0
1570    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1580    Case Else
1590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1600    End Select
1610    Resume EXITP

End Sub

Private Sub accounttype_lbl_DblClick(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_lbl_DblClick"

1710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub accounttype_GotFocus()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_GotFocus"

1810    With Me
1820      .accounttype.SelLength = 0
1830      .accounttype.SelStart = 9
1840    End With

EXITP:
1850    Exit Sub

ERRH:
1860    Select Case ERR.Number
        Case Else
1870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1880    End Select
1890    Resume EXITP

End Sub

Private Sub accounttype_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1920    intShiftDown = (Shift And acShiftMask) > 0
1930    intAltDown = (Shift And acAltMask) > 0
1940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1970        With Me
1980          intRetVal = 0
1990          .accounttype_description.SetFocus
2000        End With
2010      End Select
2020    End If

        ' ** Shift keys.
2030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2050        With Me
2060          intRetVal = 0
2070          If .CurrentRecord > 1 Then
2080            MoveRec acCmdRecordsGoToPrevious
2090            .FocusHolder.SetFocus
2100          Else
2110            DoCmd.SelectObject acForm, .Parent.Name, False
2120            .Parent.cmdClose.SetFocus
2130          End If
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

Private Sub accounttype_description_lbl_DblClick(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_description_lbl_DblClick"

2310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Private Sub accounttype_description_GotFocus()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_description_GotFocus"

2410    With Me
2420      .accounttype_description.SelLength = 0
2430      .accounttype_description.SelStart = 99
2440    End With

EXITP:
2450    Exit Sub

ERRH:
2460    Select Case ERR.Number
        Case Else
2470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2480    End Select
2490    Resume EXITP

End Sub

Private Sub accounttype_description_KeyDown(KeyCode As Integer, Shift As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2520    intShiftDown = (Shift And acShiftMask) > 0
2530    intAltDown = (Shift And acAltMask) > 0
2540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2570        With Me
2580          intRetVal = 0
2590          If .accounttypegroup_name.Visible = True And .accounttypegroup_name.Enabled = True Then
2600            .accounttypegroup_name.SetFocus
2610          ElseIf .accountno.Visible = True And .accountno.Enabled = True Then
2620            .accountno.SetFocus
2630          End If
2640        End With
2650      End Select
2660    End If

        ' ** Shift keys.
2670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2680      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2690        With Me
2700          intRetVal = 0
2710          .accounttype.SetFocus
2720        End With
2730      End Select
2740    End If

EXITP:
2750    KeyCode = intRetVal
2760    Exit Sub

ERRH:
2770    intRetVal = 0
2780    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2790    Case Else
2800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2810    End Select
2820    Resume EXITP

End Sub

Private Sub accounttypegroup_name_lbl_DblClick(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_name_lbl_DblClick"

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

Private Sub accounttypegroup_name_GotFocus()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_name_GotFocus"

3010    With Me
3020      .accounttypegroup_name.SelLength = 0
3030      .accounttypegroup_name.SelStart = 99
3040    End With

EXITP:
3050    Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub accounttypegroup_name_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_name_KeyDown"

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
3190          .accounttypegroup_sequence.SetFocus
3200        End With
3210      End Select
3220    End If

        ' ** Shift keys.
3230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3250        With Me
3260          intRetVal = 0
3270          .accounttype_description.SetFocus
3280        End With
3290      End Select
3300    End If

EXITP:
3310    KeyCode = intRetVal
3320    Exit Sub

ERRH:
3330    intRetVal = 0
3340    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3350    Case Else
3360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3370    End Select
3380    Resume EXITP

End Sub

Private Sub accounttypegroup_sequence_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_sequence_lbl_DblClick"

3410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub accounttypegroup_sequence_GotFocus()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_sequence_GotFocus"

3510    With Me
3520      .accounttypegroup_sequence.SelLength = 0
3530      .accounttypegroup_sequence.SelStart = 9
3540    End With

EXITP:
3550    Exit Sub

ERRH:
3560    Select Case ERR.Number
        Case Else
3570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3580    End Select
3590    Resume EXITP

End Sub

Private Sub accounttypegroup_sequence_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_sequence_KeyDown"

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
3710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3720            .FocusHolder.SetFocus
3730          Else
3740            DoCmd.SelectObject acForm, .Parent.Name, False
3750            .Parent.chkShowAccounts.SetFocus
3760          End If
3770        End With
3780      End Select
3790    End If

        ' ** Shift keys.
3800    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3810      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3820        With Me
3830          intRetVal = 0
3840          .accounttypegroup_name.SetFocus
3850        End With
3860      End Select
3870    End If

EXITP:
3880    KeyCode = intRetVal
3890    Exit Sub

ERRH:
3900    intRetVal = 0
3910    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3920    Case Else
3930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3940    End Select
3950    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

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

Private Sub accountno_GotFocus()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_GotFocus"

4110    With Me
4120      .accountno.SelLength = 0
4130      .accountno.SelStart = 19
4140    End With

EXITP:
4150    Exit Sub

ERRH:
4160    Select Case ERR.Number
        Case Else
4170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4180    End Select
4190    Resume EXITP

End Sub

Private Sub accountno_KeyDown(KeyCode As Integer, Shift As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4220    intShiftDown = (Shift And acShiftMask) > 0
4230    intAltDown = (Shift And acAltMask) > 0
4240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4270        With Me
4280          intRetVal = 0
4290          .shortname.SetFocus
4300        End With
4310      End Select
4320    End If

        ' ** Shift keys.
4330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4350        With Me
4360          intRetVal = 0
4370          .accounttype_description.SetFocus
4380        End With
4390      End Select
4400    End If

EXITP:
4410    KeyCode = intRetVal
4420    Exit Sub

ERRH:
4430    intRetVal = 0
4440    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4450    Case Else
4460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4470    End Select
4480    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

4510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub shortname_GotFocus()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_GotFocus"

4610    With Me
4620      .shortname.SelLength = 0
4630      .shortname.SelStart = 0
4640    End With

EXITP:
4650    Exit Sub

ERRH:
4660    Select Case ERR.Number
        Case Else
4670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4680    End Select
4690    Resume EXITP

End Sub

Private Sub shortname_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4720    intShiftDown = (Shift And acShiftMask) > 0
4730    intAltDown = (Shift And acAltMask) > 0
4740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4770        With Me
4780          intRetVal = 0
4790          lngRecsCur = RecCnt  ' ** Function: Below.
4800          If .CurrentRecord < lngRecsCur Then
4810            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4820            .FocusHolder.SetFocus
4830          Else
4840            DoCmd.SelectObject acForm, .Parent.Name, False
4850            .Parent.chkShowAccounts.SetFocus
4860          End If
4870        End With
4880      End Select
4890    End If

        ' ** Shift keys.
4900    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4910      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4920        With Me
4930          intRetVal = 0
4940          .accountno.SetFocus
4950        End With
4960      End Select
4970    End If

EXITP:
4980    KeyCode = intRetVal
4990    Exit Sub

ERRH:
5000    intRetVal = 0
5010    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5020    Case Else
5030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5040    End Select
5050    Resume EXITP

End Sub

Public Sub ShowAccount(blnShow As Boolean)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAccount"

        Dim blnSortHere As Boolean, blnResort As Boolean, blnIsDesc As Boolean
        Dim strFilter As String

5110    With Me
5120      blnSortHere = False: blnResort = False
5130      Select Case blnShow
          Case True
5140        If .Sort_line.Left = .accounttypegroup_name_lbl.Left Then blnSortHere = True
5150        .accounttypegroup_name.Visible = False
5160        .accounttypegroup_name_lbl.Visible = False
5170        .accounttypegroup_name_lbl_line.Visible = False
5180        If blnSortHere = True Then
5190          .Sort_line.Visible = False
5200          .Sort_lbl.Visible = False
5210          blnResort = True
5220        End If
5230        blnSortHere = False
5240        If .Sort_line.Left = .accounttypegroup_sequence_lbl.Left Then blnSortHere = True
5250        .accounttypegroup_sequence.Visible = False
5260        .accounttypegroup_sequence_lbl.Visible = False
5270        .accounttypegroup_sequence_lbl2.Visible = False
5280        .accounttypegroup_sequence_lbl_line.Visible = False
5290        If blnSortHere = True Then
5300          .Sort_line.Visible = False
5310          .Sort_lbl.Visible = False
5320          blnResort = True
5330        End If
5340        blnSortHere = False
5350        .accountno.Visible = True
5360        .accountno_lbl.Visible = True
5370        .accountno_lbl2.Visible = True
5380        .accountno_lbl_line.Visible = True
5390        .shortname.Visible = True
5400        .shortname_lbl.Visible = True
5410        .shortname_lbl_line.Visible = True
            ' ** A check box could be added for this.
5420        strFilter = "[IsClosed] = False"
5430        .Filter = strFilter
5440        .FilterOn = True
5450      Case False
5460        If .Sort_line.Left = .accountno_lbl.Left Then blnSortHere = True
5470        .accountno.Visible = False
5480        .accountno_lbl.Visible = False
5490        .accountno_lbl2.Visible = False
5500        .accountno_lbl_line.Visible = False
5510        If blnSortHere = True Then
5520          .Sort_line.Visible = False
5530          .Sort_lbl.Visible = False
5540          blnResort = True
5550        End If
5560        blnSortHere = False
5570        If .Sort_line.Left = .shortname_lbl.Left Then blnSortHere = True
5580        .shortname.Visible = False
5590        .shortname_lbl.Visible = False
5600        .shortname_lbl_line.Visible = False
5610        If blnSortHere = True Then
5620          .Sort_line.Visible = False
5630          .Sort_lbl.Visible = False
5640          blnResort = True
5650        End If
5660        blnSortHere = False
5670        .accounttypegroup_name.Visible = True
5680        .accounttypegroup_name_lbl.Visible = True
5690        .accounttypegroup_name_lbl_line.Visible = True
5700        .accounttypegroup_sequence.Visible = True
5710        .accounttypegroup_sequence_lbl.Visible = True
5720        .accounttypegroup_sequence_lbl2.Visible = True
5730        .accounttypegroup_sequence_lbl_line.Visible = True
            ' ** A check box could be added for this.
5740        strFilter = vbNullString
5750        .Filter = strFilter
5760        .FilterOn = False
5770      End Select
5780      DoEvents
5790      If InStr(strSortNow, " DESC") > 0 Then blnIsDesc = True Else blnIsDesc = False
5800      If blnResort = True Then
5810        SortNow "Form_Load"  ' ** Procedure: Below.
5820      ElseIf Left(strSortNow, 13) = "[accounttype]" Then
5830        SortNow "accounttype_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
5840      ElseIf Left(strSortNow, 25) = "[accounttype_description]" Then
5850        SortNow "accounttype_description_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
5860      ElseIf Left(strSortNow, 23) = "[accounttypegroup_name]" Then
5870        SortNow "accounttypegroup_name_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
5880      ElseIf Left(strSortNow, 27) = "[accounttypegroup_sequence]" Then
5890        SortNow "accounttypegroup_sequence_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
5900      ElseIf Left(strSortNow, 11) = "[accountno]" Then
5910        SortNow "accountno_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
5920      ElseIf Left(strSortNow, 11) = "[shortname]" Then
5930        SortNow "shortname_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
5940      End If
5950    End With

EXITP:
5960    Exit Sub

ERRH:
5970    Select Case ERR.Number
        Case Else
5980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5990    End Select
6000    Resume EXITP

End Sub

Public Function RecCnt() As Long

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

6110    lngRetVal = 0&

6120    With Me
6130      Set rst = .RecordsetClone
6140      With rst
6150        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
6160        Else
6170          .MoveLast
6180          lngRetVal = .RecordCount
6190        End If
6200        .Close
6210      End With
6220    End With

EXITP:
6230    Set rst = Nothing
6240    RecCnt = lngRetVal
6250    Exit Function

ERRH:
6260    lngRetVal = 0&
6270    Select Case ERR.Number
        Case Else
6280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6290    End Select
6300    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

6410    Select Case intWhere
        Case acCmdRecordsGoToFirst
6420      DoCmd.RunCommand acCmdRecordsGoToFirst
6430    Case acCmdRecordsGoToLast
6440      DoCmd.RunCommand acCmdRecordsGoToLast
6450    Case acCmdRecordsGoToPrevious
6460      DoCmd.RunCommand acCmdRecordsGoToPrevious
6470    Case acCmdRecordsGoToNext
6480      DoCmd.RunCommand acCmdRecordsGoToNext
6490    Case Else
6500      If IsMissing(varID) = False Then
6510        With Me
6520          Set rst = .RecordsetClone
6530          With rst
6540            .FindFirst "[accounttype] = '" & varID & "'"
6550            If .NoMatch = False Then
6560              Me.Bookmark = .Bookmark
6570            End If
6580            .Close
6590          End With
6600        End With
6610      End If
6620    End Select

EXITP:
6630    Set rst = Nothing
6640    Exit Sub

ERRH:
6650    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
6660    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
6670    Case Else
6680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6690    End Select
6700    Resume EXITP

End Sub

Public Sub SortNow(strProc As String, Optional varShow As Variant, Optional varIsDesc As Variant)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim blnDontReverse As Boolean, blnIsDesc As Boolean
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [accounttype], [accountno]"

6810    With Me
6820      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
6830        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6840      End If
6850      Select Case IsMissing(varShow)
          Case True
6860        blnDontReverse = False
6870      Case False
6880        blnDontReverse = CBool(varShow)
6890      End Select
6900      Select Case IsMissing(varIsDesc)
          Case True
6910        blnIsDesc = False
6920      Case False
6930        blnIsDesc = CBool(varIsDesc)
6940      End Select
6950      .Controls(strSortLbl).Visible = False
6960      .Controls(strSortLine).Visible = False
6970      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
6980      If strProc = "Form_Load" Then
6990        strCalled = "accounttype"
7000        strSortNow = strSortOrig
7010        If .Parent.chkShowAccounts = True Then
7020          strSortNow = strSortNow & ", [accountno]"
7030        End If
7040        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7050        lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (6& * lngTpp))
7060        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7070        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7080        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7090        .Controls(strSortLbl).Top = lngSortLbl_Top
7100        .Controls(strSortLbl).Left = lngSortLbl_Left
7110        .Controls(strSortLine).Top = lngSortLine_Top
7120        .Controls(strSortLine).Left = lngSortLine_Left
7130        .Controls(strSortLine).Width = lngSortLine_Width
7140        .Controls(strSortLbl).Caption = strArwUp
7150        .Controls(strSortLbl).ForeColor = CLR_DKBLU
7160      Else
7170        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
7180        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7190        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7200        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7210        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7220        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7230        .Controls(strSortLbl).Top = lngSortLbl_Top
7240        .Controls(strSortLbl).Left = lngSortLbl_Left
7250        .Controls(strSortLine).Top = lngSortLine_Top
7260        .Controls(strSortLine).Left = lngSortLine_Left
7270        .Controls(strSortLine).Width = lngSortLine_Width
7280        If strCalled = "accounttype" Then
7290          lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
7300          .Controls(strSortLbl).Left = lngSortLbl_Left
7310          strSortAsc = strSortOrig
7320          If .Parent.chkShowAccounts = True Then
7330            strSortAsc = strSortAsc & ", [accountno]"
7340          End If
7350          If ((strSortNow = strSortAsc And blnDontReverse = False) Or (blnIsDesc = True)) Then
7360            strSortNow = "[accounttype] DESC"
7370            If .Parent.chkShowAccounts = True Then
7380              strSortNow = strSortNow & ", [accountno] DESC"
7390            End If
7400            .Controls(strSortLbl).Caption = strArwDn
7410            .Controls(strSortLbl).ForeColor = CLR_DKRED
7420          Else
7430            strSortNow = strSortAsc
7440            .Controls(strSortLbl).Caption = strArwUp
7450            .Controls(strSortLbl).ForeColor = CLR_DKBLU
7460          End If
7470        Else
7480          .Controls(strSortLbl).Caption = strArwUp
7490          .Controls(strSortLbl).ForeColor = CLR_DKBLU
7500          Select Case strCalled
              Case "accounttype_description"
7510            strSortAsc = "[accounttype_description]"
7520            If .Parent.chkShowAccounts = True Then
7530              strSortAsc = strSortAsc & ", [accountno]"
7540            End If
7550          Case "accounttypegroup_name"
7560            strSortAsc = "[accounttypegroup_name]" & strStdAsc
7570          Case "accounttypegroup_sequence"
7580            strSortAsc = "[accounttypegroup_sequence]" & strStdAsc
7590          Case "accountno"
7600            strSortAsc = "[accountno]"
7610          Case "shortname"
7620            strSortAsc = "[shortname]"
7630          End Select
7640          If ((strSortNow = strSortAsc And blnDontReverse = False) Or (blnIsDesc = True)) Then
7650            intCnt = CharCnt(strSortAsc, ",") + 1
7660            Select Case intCnt
                Case 1
7670              strTmp01 = strSortAsc & " DESC"
7680            Case 2
7690              intPos01 = InStr(strSortAsc, ",")
7700              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
7710              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
7720            Case 3
7730              intPos01 = InStr(strSortAsc, ",")
7740              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
7750              strTmp02 = Mid(strSortAsc, intPos01)
7760              intPos01 = InStr(2, strTmp02, ",")
7770              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
7780              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
7790            End Select
7800            strSortNow = strTmp01
7810            .Controls(strSortLbl).Caption = strArwDn
7820            .Controls(strSortLbl).ForeColor = CLR_DKRED
7830          Else
7840            strSortNow = strSortAsc
7850          End If
7860        End If
7870      End If
7880      .Controls(strSortLbl).Visible = True
7890      .Controls(strSortLine).Visible = True
7900      .OrderBy = strSortNow
7910      .OrderByOn = True
7920    End With

        ' ** Potential Sorts:
        ' **   [accounttype]
        ' **   [accounttype], [accountno]
        ' **   [accounttype_description]
        ' **   [accounttype_description], [accountno]
        ' **   [accounttypegroup_name], [accounttype]
        ' **   [accounttypegroup_sequence], [accounttype]
        ' **   [accountno]
        ' **   [shortname]

EXITP:
7930    Exit Sub

ERRH:
7940    Select Case ERR.Number
        Case Else
7950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7960    End Select
7970    Resume EXITP

End Sub

Public Function SortNow_Get() As String

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

8010    strRetVal = strSortNow

EXITP:
8020    SortNow_Get = strRetVal
8030    Exit Function

ERRH:
8040    strRetVal = strSortOrig
8050    Select Case ERR.Number
        Case Else
8060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8070    End Select
8080    Resume EXITP

End Function
