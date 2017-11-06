VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_Account_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_Account_Sub"

'VGC 08/31/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   All:              A {opgFilter_optAll on frmCurrency_Account}
' **   Close:            C {cmdClose on frmCurrency_Account}
' **   Default Suppress: D {chkDefaultSuppress on frmCurrency_Account}
' **   Suppress On/Off:  F {curracct_suppress}
' **   Option Suppress:  H {chkOptionSuppress on frmCurrency_Account}
' **   Non-Suppressed:   N {opgFilter_optNonSuppressed on frmCurrency_Account}
' **   Regenerate:       R {cmdRegenerate on frmCurrency_Account}
' **   Suppressed:       S {opgFilter_optSuppressed on frmCurrency_Account}
' **   Exit:             X {cmdClose on frmCurrency_Account}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[curracct_sort]"
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
150       If strCallingForm = "frmCurrency_Account" Then

160         lngSortLbl_Width = .Sort_lbl.Width

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

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

410     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
420     Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Form_Current()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

510     With Me
520       If (.curracct_jno > 0 Or .curracct_aa > 0) And .curracct_suppress = False Then
            ' ** Don't let them change this to True.
530         .curracct_suppress.Locked = True
540       Else
550         .curracct_suppress.Locked = False
560       End If
570     End With

EXITP:
580     Exit Sub

ERRH:
590     Select Case ERR.Number
        Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Default Suppress: D {chkDefaultSuppress on frmCurrency_Account}
        ' **   Foreign Check:    F {curracct_suppress}
        ' **   Option Suppress:  H {chkOptionSuppress on frmCurrency_Account}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         With Me
780           intRetVal = 0
790           .Parent.cmdClose_Click  ' ** Form Procedure: frmCurrency_Account.
800         End With
810       Case vbKeyUp
820         With Me
830           intRetVal = 0
840           MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
850         End With
860       Case vbKeyDown
870         With Me
880           intRetVal = 0
890           MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
900         End With
910       End Select
920     End If

        ' ** Alt keys.
930     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
940       Select Case intRetVal
          Case vbKeyD
950         With Me
960           intRetVal = 0
970           DoCmd.SelectObject acForm, .Parent.Name, False
980           .Parent.chkDefaultSuppress.SetFocus
990           .Parent.chkDefaultSuppress = (Not .Parent.chkDefaultSuppress)
1000          .Parent.chkDefaultSuppress_AfterUpdate  ' ** Form Procedure: frmCurrency_Account.
1010        End With
1020      Case vbKeyF
1030        With Me
1040          intRetVal = 0
1050          .curracct_sort.SetFocus  ' ** Set focus only, dont't flip.
1060        End With
1070      Case vbKeyH
1080        With Me
1090          intRetVal = 0
1100          DoCmd.SelectObject acForm, .Parent.Name, False
1110          .Parent.chkOptionSuppress.SetFocus
1120          .Parent.chkOptionSuppress = (Not .Parent.chkOptionSuppress)
1130          .Parent.chkOptionSuppress_AfterUpdate  ' ** Procedure: Below.
1140        End With
1150      Case vbKeyX
1160        With Me
1170          intRetVal = 0
1180          .Parent.cmdClose_Click  ' ** Form Procedure: frmCurrency_Account.
1190        End With
1200      End Select
1210    End If

        ' ** Ctrl keys.
1220    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1230      Select Case intRetVal
          Case vbKeyS
1240        intRetVal = 0
1250        cmdSave_Click  ' ** Procedure: Below.
1260      Case vbKeyTab, vbKeyReturn
1270        With Me
1280          intRetVal = 0
1290          DoCmd.SelectObject acForm, .Parent.Name, False
1300          .Parent.opgFilter.SetFocus
1310        End With
1320      Case vbKeyPageUp
1330        intRetVal = 0
1340        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1350      Case vbKeyPageDown
1360        intRetVal = 0
1370        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1380      End Select
1390    End If

        ' ** Ctrl-Shift keys.
1400    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1410      Select Case intRetVal
          Case vbKeyF
1420        With Me
1430          intRetVal = 0
1440          DoCmd.SelectObject acForm, .Parent.Name, False
1450          .Parent.FocusHolder.SetFocus
1460        End With
1470      Case vbKeyTab, vbKeyReturn
1480        With Me
1490          intRetVal = 0
1500          DoCmd.SelectObject acForm, .Parent.Name, False
1510          .Parent.cmdClose.SetFocus
1520        End With
1530      End Select
1540    End If

EXITP:
1550    KeyCode = intRetVal
1560    Exit Sub

ERRH:
1570    intRetVal = 0
1580    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1590    Case Else
1600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1610    End Select
1620    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

1710    With Me
1720      .curracct_datemodified = Now()
1730    End With

EXITP:
1740    Exit Sub

ERRH:
1750    Select Case ERR.Number
        Case Else
1760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1770    End Select
1780    Resume EXITP

End Sub

Public Sub cmdSave_Click()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1810    DoCmd.RunCommand acCmdSaveRecord
1820    DoEvents

EXITP:
1830    Exit Sub

ERRH:
1840    Select Case ERR.Number
        Case Else
1850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1860    End Select
1870    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

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
1990          .curracct_suppress.SetFocus
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
2080            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2090            .curracct_suppress.SetFocus
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

Private Sub accountno_lbl_DblClick(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

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

Private Sub accountno_zero_GotFocus()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_zero_GotFocus"

2410    With Me
2420      If .curracct_jno > 0 Then
2430        .accountno_notzero.SetFocus
2440      Else
2450        .accountno_zero.SelLength = 0
2460        .accountno_zero.SelStart = 19
2470      End If
2480    End With

EXITP:
2490    Exit Sub

ERRH:
2500    Select Case ERR.Number
        Case Else
2510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2520    End Select
2530    Resume EXITP

End Sub

Private Sub accountno_zero_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_zero_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2620    intShiftDown = (Shift And acShiftMask) > 0
2630    intAltDown = (Shift And acAltMask) > 0
2640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2670        With Me
2680          intRetVal = 0
2690          If .curracct_jno = 0 Then
2700            If .shortname_zero.Enabled = True Then
2710              .shortname_zero.SetFocus
2720            Else
2730              .curracct_suppress.SetFocus
2740            End If
2750          Else
2760            If .shortname_notzero.Enabled = True Then
2770              .shortname_notzero.SetFocus
2780            Else
2790              .curracct_suppress.SetFocus
2800            End If
2810          End If
2820        End With
2830      End Select
2840    End If

        ' ** Shift keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2870        With Me
2880          intRetVal = 0
2890          If .CurrentRecord > 1 Then
2900            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2910            .curracct_suppress.SetFocus
2920          Else
2930            DoCmd.SelectObject acForm, .Parent.Name, False
2940            .Parent.cmdClose.SetFocus
2950          End If
2960        End With
2970      End Select
2980    End If

EXITP:
2990    KeyCode = intRetVal
3000    Exit Sub

ERRH:
3010    intRetVal = 0
3020    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3030    Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub accountno_notzero_GotFocus()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_notzero_GotFocus"

3110    With Me
3120      If .curracct_jno = 0 Then
3130        .accountno_zero.SetFocus
3140      Else
3150        .accountno_notzero.SelLength = 0
3160        .accountno_notzero.SelStart = 19
3170      End If
3180    End With

EXITP:
3190    Exit Sub

ERRH:
3200    Select Case ERR.Number
        Case Else
3210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3220    End Select
3230    Resume EXITP

End Sub

Private Sub accountno_notzero_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_notzero_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3320    intShiftDown = (Shift And acShiftMask) > 0
3330    intAltDown = (Shift And acAltMask) > 0
3340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3370        With Me
3380          intRetVal = 0
3390          If .curracct_jno = 0 Then
3400            If .shortname_zero.Enabled = True Then
3410              .shortname_zero.SetFocus
3420            Else
3430              .curracct_suppress.SetFocus
3440            End If
3450          Else
3460            If .shortname_notzero.Enabled = True Then
3470              .shortname_notzero.SetFocus
3480            Else
3490              .curracct_suppress.SetFocus
3500            End If
3510          End If
3520        End With
3530      End Select
3540    End If

        ' ** Shift keys.
3550    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3570        With Me
3580          intRetVal = 0
3590          If .CurrentRecord > 1 Then
3600            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3610            .curracct_suppress.SetFocus
3620          Else
3630            DoCmd.SelectObject acForm, .Parent.Name, False
3640            .Parent.cmdClose.SetFocus
3650          End If
3660        End With
3670      End Select
3680    End If

EXITP:
3690    KeyCode = intRetVal
3700    Exit Sub

ERRH:
3710    intRetVal = 0
3720    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3730    Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

3810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub shortname_zero_GotFocus()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_zero_GotFocus"

3910    With Me
3920      If .curracct_jno > 0 Then
3930        .shortname_notzero.SetFocus
3940      Else
3950        .shortname_zero.SelLength = 0
3960        .shortname_zero.SelStart = 99
3970      End If
3980    End With

EXITP:
3990    Exit Sub

ERRH:
4000    Select Case ERR.Number
        Case Else
4010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4020    End Select
4030    Resume EXITP

End Sub

Private Sub shortname_zero_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_zero_KeyDown"

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
4190          If .curracct_jno = 0 Then
4200            If .curracct_jno_zero.Enabled = True Then
4210              .curracct_jno_zero.SetFocus
4220            Else
4230              .curracct_suppress.SetFocus
4240            End If
4250          Else
4260            If .curracct_jno_notzero.Enabled = True Then
4270              .curracct_jno_notzero.SetFocus
4280            Else
4290              .curracct_suppress.SetFocus
4300            End If
4310          End If
4320        End With
4330      End Select
4340    End If

        ' ** Shift keys.
4350    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4370        With Me
4380          intRetVal = 0
4390          If .curracct_jno = 0 Then
4400            .accountno_zero.SetFocus
4410          Else
4420            .accountno_notzero.SetFocus
4430          End If
4440        End With
4450      End Select
4460    End If

EXITP:
4470    KeyCode = intRetVal
4480    Exit Sub

ERRH:
4490    intRetVal = 0
4500    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4510    Case Else
4520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4530    End Select
4540    Resume EXITP

End Sub

Private Sub shortname_notzero_GotFocus()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_notzero_GotFocus"

4610    With Me
4620      If .curracct_jno = 0 Then
4630        .shortname_zero.SetFocus
4640      Else
4650        .shortname_notzero.SelLength = 0
4660        .shortname_notzero.SelStart = 99
4670      End If
4680    End With

EXITP:
4690    Exit Sub

ERRH:
4700    Select Case ERR.Number
        Case Else
4710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4720    End Select
4730    Resume EXITP

End Sub

Private Sub shortname_notzero_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_notzero_KeyDown"

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
4890          If .curracct_jno = 0 Then
4900            If .curracct_jno_zero.Enabled = True Then
4910              .curracct_jno_zero.SetFocus
4920            Else
4930              .curracct_suppress.SetFocus
4940            End If
4950          Else
4960            If .curracct_jno_notzero.Enabled = True Then
4970              .curracct_jno_notzero.SetFocus
4980            Else
4990              .curracct_suppress.SetFocus
5000            End If
5010          End If
5020        End With
5030      End Select
5040    End If

        ' ** Shift keys.
5050    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5070        With Me
5080          intRetVal = 0
5090          If .curracct_jno = 0 Then
5100            .accountno_zero.SetFocus
5110          Else
5120            .accountno_notzero.SetFocus
5130          End If
5140        End With
5150      End Select
5160    End If

EXITP:
5170    KeyCode = intRetVal
5180    Exit Sub

ERRH:
5190    intRetVal = 0
5200    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5210    Case Else
5220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5230    End Select
5240    Resume EXITP

End Sub

Private Sub curracct_jno_lbl_DblClick(Cancel As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_jno_lbl_DblClick"

5310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Private Sub curracct_jno_zero_GotFocus()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_jno_zero_GotFocus"

5410    With Me
5420      If .curracct_jno > 0 Then
5430        .curracct_jno_notzero.SetFocus
5440      Else
5450        .curracct_jno_zero.SelLength = 0
5460        .curracct_jno_zero.SelStart = 9
5470      End If
5480    End With

EXITP:
5490    Exit Sub

ERRH:
5500    Select Case ERR.Number
        Case Else
5510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5520    End Select
5530    Resume EXITP

End Sub

Private Sub curracct_jno_zero_KeyDown(KeyCode As Integer, Shift As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_jno_zero_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5620    intShiftDown = (Shift And acShiftMask) > 0
5630    intAltDown = (Shift And acAltMask) > 0
5640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5670        With Me
5680          intRetVal = 0
5690          If .curracct_aa = 0 Then
5700            If .curracct_aa_zero.Enabled = True Then
5710              .curracct_aa_zero.SetFocus
5720            Else
5730              .curracct_suppress.SetFocus
5740            End If
5750          Else
5760            If .curracct_aa_notzero.Enabled = True Then
5770              .curracct_aa_notzero.SetFocus
5780            Else
5790              .curracct_suppress.SetFocus
5800            End If
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
5890          .shortname.SetFocus
5900        End With
5910      End Select
5920    End If

EXITP:
5930    KeyCode = intRetVal
5940    Exit Sub

ERRH:
5950    intRetVal = 0
5960    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5970    Case Else
5980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5990    End Select
6000    Resume EXITP

End Sub

Private Sub curracct_jno_notzero_GotFocus()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_jno_notzero_GotFocus"

6110    With Me
6120      If .curracct_jno = 0 Then
6130        .curracct_jno_zero.SetFocus
6140      Else
6150        .curracct_jno_notzero.SelLength = 0
6160        .curracct_jno_notzero.SelStart = 9
6170      End If
6180    End With

EXITP:
6190    Exit Sub

ERRH:
6200    Select Case ERR.Number
        Case Else
6210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6220    End Select
6230    Resume EXITP

End Sub

Private Sub curracct_jno_notzero_KeyDown(KeyCode As Integer, Shift As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_jno_notzero_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6320    intShiftDown = (Shift And acShiftMask) > 0
6330    intAltDown = (Shift And acAltMask) > 0
6340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6370        With Me
6380          intRetVal = 0
6390          If .curracct_aa = 0 Then
6400            If .curracct_aa_zero.Enabled = True Then
6410              .curracct_aa_zero.SetFocus
6420            Else
6430              .curracct_suppress.SetFocus
6440            End If
6450          Else
6460            If .curracct_aa_notzero.Enabled = True Then
6470              .curracct_aa_notzero.SetFocus
6480            Else
6490              .curracct_suppress.SetFocus
6500            End If
6510          End If
6520        End With
6530      End Select
6540    End If

        ' ** Shift keys.
6550    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6570        With Me
6580          intRetVal = 0
6590          .shortname.SetFocus
6600        End With
6610      End Select
6620    End If

EXITP:
6630    KeyCode = intRetVal
6640    Exit Sub

ERRH:
6650    intRetVal = 0
6660    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6670    Case Else
6680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6690    End Select
6700    Resume EXITP

End Sub

Private Sub curracct_aa_lbl_DblClick(Cancel As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_aa_lbl_DblClick"

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

Private Sub curracct_aa_zero_GotFocus()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_aa_zero_GotFocus"

6910    With Me
6920      If .curracct_aa > 0 Then
6930        .curracct_aa_notzero.SetFocus
6940      Else
6950        .curracct_aa_zero.SelLength = 0
6960        .curracct_aa_zero.SelStart = 9
6970      End If
6980    End With

EXITP:
6990    Exit Sub

ERRH:
7000    Select Case ERR.Number
        Case Else
7010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7020    End Select
7030    Resume EXITP

End Sub

Private Sub curracct_aa_zero_KeyDown(KeyCode As Integer, Shift As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_aa_zero_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7120    intShiftDown = (Shift And acShiftMask) > 0
7130    intAltDown = (Shift And acAltMask) > 0
7140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7170        With Me
7180          intRetVal = 0
7190          .curracct_suppress.SetFocus
7200        End With
7210      End Select
7220    End If

        ' ** Shift keys.
7230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7250        With Me
7260          intRetVal = 0
7270          If .curracct_jno = 0 Then
7280            If .curracct_jno_zero.Enabled = True Then
7290              .curracct_jno_zero.SetFocus
7300            Else
7310              .curracct_suppress.SetFocus
7320            End If
7330          Else
7340            If .curracct_jno_notzero.Enabled = True Then
7350              .curracct_jno_notzero.SetFocus
7360            Else
7370              .curracct_suppress.SetFocus
7380            End If
7390          End If
7400        End With
7410      End Select
7420    End If

EXITP:
7430    KeyCode = intRetVal
7440    Exit Sub

ERRH:
7450    intRetVal = 0
7460    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7470    Case Else
7480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7490    End Select
7500    Resume EXITP

End Sub

Private Sub curracct_aa_notzero_GotFocus()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_aa_notzero_GotFocus"

7610    With Me
7620      If .curracct_aa = 0 Then
7630        .curracct_aa_zero.SetFocus
7640      Else
7650        .curracct_aa_notzero.SelLength = 0
7660        .curracct_aa_notzero.SelStart = 9
7670      End If
7680    End With

EXITP:
7690    Exit Sub

ERRH:
7700    Select Case ERR.Number
        Case Else
7710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7720    End Select
7730    Resume EXITP

End Sub

Private Sub curracct_aa_notzero_KeyDown(KeyCode As Integer, Shift As Integer)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_aa_notzero_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7820    intShiftDown = (Shift And acShiftMask) > 0
7830    intAltDown = (Shift And acAltMask) > 0
7840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7870        With Me
7880          intRetVal = 0
7890          .curracct_suppress.SetFocus
7900        End With
7910      End Select
7920    End If

        ' ** Shift keys.
7930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7950        With Me
7960          intRetVal = 0
7970          If .curracct_jno = 0 Then
7980            If .curracct_jno_zero.Enabled = True Then
7990              .curracct_jno_zero.SetFocus
8000            Else
8010              .curracct_suppress.SetFocus
8020            End If
8030          Else
8040            If .curracct_jno_notzero.Enabled = True Then
8050              .curracct_jno_notzero.SetFocus
8060            Else
8070              .curracct_suppress.SetFocus
8080            End If
8090          End If
8100        End With
8110      End Select
8120    End If

EXITP:
8130    KeyCode = intRetVal
8140    Exit Sub

ERRH:
8150    intRetVal = 0
8160    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8170    Case Else
8180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8190    End Select
8200    Resume EXITP

End Sub

Private Sub curracct_suppress_lbl_DblClick(Cancel As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_suppress_lbl_DblClick"

8310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8320    Exit Sub

ERRH:
8330    Select Case ERR.Number
        Case Else
8340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8350    End Select
8360    Resume EXITP

End Sub

Private Sub curracct_suppress_AfterUpdate()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_suppress_AfterUpdate"

8410    cmdSave_Click  ' ** Procedure: Above

EXITP:
8420    Exit Sub

ERRH:
8430    Select Case ERR.Number
        Case Else
8440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8450    End Select
8460    Resume EXITP

End Sub

Private Sub curracct_suppress_KeyDown(KeyCode As Integer, Shift As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "curracct_suppress_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8520    intShiftDown = (Shift And acShiftMask) > 0
8530    intAltDown = (Shift And acAltMask) > 0
8540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8570        With Me
8580          intRetVal = 0
8590          lngRecsCur = RecCnt  ' ** Function: Below.
8600          If .CurrentRecord < lngRecsCur Then
8610            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
8620            .curracct_suppress.SetFocus
8630          Else
8640            DoCmd.SelectObject acForm, .Parent.Name, False
8650            .Parent.opgFilter.SetFocus
8660          End If
8670        End With
8680      End Select
8690    End If

        ' ** Shift keys.
8700    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8710      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8720        With Me
8730          intRetVal = 0
8740          If .CurrentRecord > 1 Then
8750            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
8760            .curracct_suppress.SetFocus
8770          Else
8780            DoCmd.SelectObject acForm, .Parent.Name, False
8790            .Parent.cmdClose.SetFocus
8800          End If
8810        End With
8820      End Select
8830    End If

EXITP:
8840    KeyCode = intRetVal
8850    Exit Sub

ERRH:
8860    intRetVal = 0
8870    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8880    Case Else
8890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8900    End Select
8910    Resume EXITP

End Sub

Public Function RecCnt() As Long

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

9010    lngRetVal = 0&

9020    With Me
9030      Set rst = .RecordsetClone
9040      With rst
9050        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
9060        Else
9070          .MoveLast
9080          lngRetVal = .RecordCount
9090        End If
9100        .Close
9110      End With
9120    End With

EXITP:
9130    Set rst = Nothing
9140    RecCnt = lngRetVal
9150    Exit Function

ERRH:
9160    lngRetVal = 0&
9170    Select Case ERR.Number
        Case Else
9180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9190    End Select
9200    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

9310    Select Case intWhere
        Case acCmdRecordsGoToFirst
9320      DoCmd.RunCommand acCmdRecordsGoToFirst
9330    Case acCmdRecordsGoToLast
9340      DoCmd.RunCommand acCmdRecordsGoToLast
9350    Case acCmdRecordsGoToPrevious
9360      DoCmd.RunCommand acCmdRecordsGoToPrevious
9370    Case acCmdRecordsGoToNext
9380      DoCmd.RunCommand acCmdRecordsGoToNext
9390    Case Else
9400      If IsMissing(varID) = False Then
9410        With Me
9420          Set rst = .RecordsetClone
9430          With rst
9440            .FindFirst "[curracct_id] = " & varID
9450            If .NoMatch = False Then
9460              Me.Bookmark = .Bookmark
9470            End If
9480            .Close
9490          End With
9500        End With
9510      End If
9520    End Select

EXITP:
9530    Set rst = Nothing
9540    Exit Sub

ERRH:
9550    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
9560    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
9570    Case Else
9580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9590    End Select
9600    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [curracct_sort]"

9710    With Me
9720      If lngTpp = 0 Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
9730        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
9740      End If
9750      .Controls(strSortLbl).Visible = False
9760      .Controls(strSortLine).Visible = False
9770      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
9780      If strProc = "Form_Load" Then
9790        strCalled = "accountno"
9800        strSortNow = strSortOrig
9810        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9820        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
9830        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9840        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
9850        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
9860        .Controls(strSortLbl).Top = lngSortLbl_Top
9870        .Controls(strSortLbl).Left = lngSortLbl_Left
9880        .Controls(strSortLine).Top = lngSortLine_Top
9890        .Controls(strSortLine).Left = lngSortLine_Left
9900        .Controls(strSortLine).Width = lngSortLine_Width
9910        .Controls(strSortLbl).Caption = strArwUp
9920        .Controls(strSortLbl).ForeColor = CLR_DKBLU
9930      Else
9940        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
9950        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9960        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
9970        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9980        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
9990        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
10000       .Controls(strSortLbl).Top = lngSortLbl_Top
10010       .Controls(strSortLbl).Left = lngSortLbl_Left
10020       .Controls(strSortLine).Top = lngSortLine_Top
10030       .Controls(strSortLine).Left = lngSortLine_Left
10040       .Controls(strSortLine).Width = lngSortLine_Width
10050       If strCalled = "accountno" Then
10060         strSortAsc = strSortOrig
10070         If strSortNow = strSortAsc Then
10080           strSortNow = "[curracct_sort] DESC"
10090           .Controls(strSortLbl).Caption = strArwDn
10100           .Controls(strSortLbl).ForeColor = CLR_DKRED
10110         Else
10120           strSortNow = strSortAsc
10130           .Controls(strSortLbl).Caption = strArwUp
10140           .Controls(strSortLbl).ForeColor = CLR_DKBLU
10150         End If
10160       Else
10170         .Controls(strSortLbl).Caption = strArwUp
10180         .Controls(strSortLbl).ForeColor = CLR_DKBLU
10190         Select Case strCalled
              Case "shortname"
10200           strSortAsc = "[shortname]" & strStdAsc
10210         Case "curracct_jno"
10220           strSortAsc = "[curracct_jno]"
10230         Case "curracct_aa"
10240           strSortAsc = "[curracct_aa]"
10250         Case "curracct_suppress"
10260           strSortAsc = "[curracct_suppress]"
10270           lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
10280           .Controls(strSortLbl).Left = lngSortLbl_Left
10290         End Select
10300         If strSortNow = strSortAsc Then
10310           intCnt = CharCnt(strSortAsc, ",") + 1
10320           Select Case intCnt
                Case 1
10330             strTmp01 = strSortAsc & " DESC"
10340           Case 2
10350             intPos01 = InStr(strSortAsc, ",")
10360             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10370             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
10380           Case 3
10390             intPos01 = InStr(strSortAsc, ",")
10400             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10410             strTmp02 = Mid(strTmp01, intPos01)
10420             intPos01 = InStr(2, strTmp02, ",")
10430             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10440             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
10450           End Select
10460           strSortNow = strTmp01
10470           .Controls(strSortLbl).Caption = strArwDn
10480           .Controls(strSortLbl).ForeColor = CLR_DKRED
10490         Else
10500           strSortNow = strSortAsc
10510         End If
10520       End If
10530     End If
10540     .Controls(strSortLbl).Visible = True
10550     .Controls(strSortLine).Visible = True
10560     .OrderBy = strSortNow
10570     .OrderByOn = True
10580   End With

EXITP:
10590   Exit Sub

ERRH:
10600   Select Case ERR.Number
        Case Else
10610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10620   End Select
10630   Resume EXITP

End Sub
