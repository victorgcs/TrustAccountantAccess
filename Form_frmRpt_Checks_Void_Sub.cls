VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks_Void_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks_Void_Sub"

'VGC 07/05/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Bank Accot Num:   A {chkbank_acctnum}
' **   Account Number:   B {accountno}
' **   Copy:             C {cmdCopy on frmRpt_Checks_Void}
' **   Bank Name:        E {chkbank_name}
' **   Check Number:     K {chkvoid_chknum}
' **   Account Name:     M {shortname}
' **   Select None:      N {cmdSelectNone on frmRpt_Checks_Void}
' **   OK:               O {cmdOK on frmRpt_Checks_Void}
' **   Payee:            P {chkvoid_payee}
' **   Select All:       S {cmdSelectAll on frmRpt_Checks_Void}
' **   Void:             V {chkvoid_void}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:           F5 {RecalcTots}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[chkvoid_chknum]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private strSortNow As String, lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngTpp As Long, lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmRpt_Checks_Void" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190         lngSortLbl_Width = .Sort_lbl.Width

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

420       .chkvoid_chknum.ForeColor = CLR_DISABLED_FG
430       .chkvoid_chknum.BackColor = CLR_DISABLED_BG
440       .accountno.ForeColor = CLR_DISABLED_FG
450       .accountno.BackColor = CLR_DISABLED_BG
460       .shortname.ForeColor = CLR_DISABLED_FG
470       .shortname.BackColor = CLR_DISABLED_BG
480       .transdate.ForeColor = CLR_DISABLED_FG
490       .transdate.BackColor = CLR_DISABLED_BG
500       .chkvoid_payee.ForeColor = CLR_DISABLED_FG
510       .chkvoid_payee.BackColor = CLR_DISABLED_BG

520       SortNow THIS_PROC  ' ** Procedure: Below.

530     End With

EXITP:
540     Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub Form_Current()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim blnAble As Boolean
        Dim varTmp00 As Variant

610     With Me
620       If .InChkBnkStg = True And .InChkBnk = False And .changed = False Then
630         .changed = True
640         cmdSave_Click  ' ** Procedure: Below.
650       End If
660       DoEvents
670       If .chkvoid_void = True And (IsNull(.chkbank_name) = True Or IsNull(.chkbank_acctnum) = True) Then
680         .Parent.MsgSubhead_lbl1.Visible = True
690         .Parent.MsgSubhead_lbl2.Visible = True
700       Else
710         .Parent.MsgSubhead_lbl1.Visible = False
720         .Parent.MsgSubhead_lbl2.Visible = False
730       End If
740       DoEvents
750       blnAble = False
760       If IsNull(.chkbank_name) = False And IsNull(.chkbank_acctnum) = False Then  ' ** Bank info is present.
770   On Error Resume Next
780         varTmp00 = .chkvoid_void_sum
790         If ERR.Number = 0 Then
800   On Error GoTo ERRH
810           If varTmp00 <> 0 Then  ' ** Voids are present. (It doesn't matter if this one is voided or not.)
820             .Parent.cmdCopy.Enabled = True
830             blnAble = True
840           Else
850             .Parent.cmdCopy.Enabled = False
860           End If
870         Else
880   On Error GoTo ERRH
890           .Parent.cmdCopy.Enabled = False
900         End If
910       Else
920         .Parent.cmdCopy.Enabled = False
930       End If
940       Select Case blnAble
          Case True
950         .Parent.cmdCopy_raised_img.Visible = True
960         .Parent.cmdCopy_raised_semifocus_dots_img.Visible = False
970         .Parent.cmdCopy_raised_focus_img.Visible = False
980         .Parent.cmdCopy_raised_focus_dots_img.Visible = False
990         .Parent.cmdCopy_sunken_focus_dots_img.Visible = False
1000        .Parent.cmdCopy_raised_img_dis.Visible = False
1010      Case False
1020        .Parent.cmdCopy_raised_img_dis.Visible = True
1030        .Parent.cmdCopy_raised_img.Visible = False
1040        .Parent.cmdCopy_raised_semifocus_dots_img.Visible = False
1050        .Parent.cmdCopy_raised_focus_img.Visible = False
1060        .Parent.cmdCopy_raised_focus_dots_img.Visible = False
1070        .Parent.cmdCopy_sunken_focus_dots_img.Visible = False
1080      End Select
1090      If .Parent.chkvoid_set = 0 Then
1100        .Parent.chkvoid_set = .chkvoid_set
1110      End If
1120      DoEvents
1130    End With

EXITP:
1140    Exit Sub

ERRH:
1150    Select Case ERR.Number
        Case Else
1160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1170    End Select
1180    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Bank Accot Num:   A {chkbank_acctnum}
        ' **   Account Number:   B {accountno}
        ' **   Bank Name:        E {chkbank_name}
        ' **   Check Number:     K {chkvoid_chknum}
        ' **   Account Name:     M {shortname}
        ' **   Payee:            P {chkvoid_payee}
        ' **   Void:             V {chkvoid_void}

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc:           F5 {RecalcTots}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyUp
1270        intRetVal = 0
1280        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1290      Case vbKeyDown
1300        intRetVal = 0
1310        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1320      Case vbKeyF5
1330        intRetVal = 0
1340        RecalcTots  ' ** Procedure: Below.
1350      End Select
1360    End If

        ' ** Alt keys.
1370    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1380      Select Case intRetVal
          Case vbKeyA
1390        With Me
1400          intRetVal = 0
1410          .chkbank_acctnum.SetFocus
1420        End With
1430      Case vbKeyB
1440        With Me
1450          intRetVal = 0
1460          .accountno.SetFocus
1470        End With
1480      Case vbKeyE
1490        With Me
1500          intRetVal = 0
1510          .chkbank_name.SetFocus
1520        End With
1530      Case vbKeyK
1540        With Me
1550          intRetVal = 0
1560          .chkvoid_chknum.SetFocus
1570        End With
1580      Case vbKeyM
1590        With Me
1600          intRetVal = 0
1610          .shortname.SetFocus
1620        End With
1630      Case vbKeyP
1640        With Me
1650          intRetVal = 0
1660          .chkvoid_payee.SetFocus
1670        End With
1680      Case vbKeyV
1690        With Me
1700          intRetVal = 0
1710          .chkvoid_void.SetFocus
1720          .chkvoid_void = (Not .chkvoid_void)
1730          chkvoid_void_AfterUpdate  ' ** Procedure Below.
1740        End With
1750      End Select
1760    End If

        ' ** Ctrl keys.
1770    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1780      Select Case intRetVal
          Case vbKeyS
1790        intRetVal = 0
1800        cmdSave_Click  ' ** Procedure: Below.
1810      Case vbKeyTab, vbKeyReturn
1820        With Me
1830          intRetVal = 0
1840          DoCmd.SelectObject acForm, .Parent.Name, False
1850          .Parent.cmdOK.SetFocus
1860        End With
1870      Case vbKeyPageUp, vbKeyHome
1880        intRetVal = 0
1890        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1900      Case vbKeyPageDown, vbKeyEnd
1910        intRetVal = 0
1920        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1930      End Select
1940    End If

        ' ** Ctrl-Shift keys.
1950    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1960      Select Case intRetVal
          Case vbKeyF
1970        With Me
1980          intRetVal = 0
1990          DoCmd.SelectObject acForm, .Parent.Name, False
2000          .Parent.FocusHolder.SetFocus
2010        End With
2020      Case vbKeyTab, vbKeyReturn
2030        With Me
2040          intRetVal = 0
2050          DoCmd.SelectObject acForm, .Parent.Name, False
2060          .Parent.cmdOK.SetFocus
2070        End With
2080      End Select
2090    End If

EXITP:
2100    KeyCode = intRetVal
2110    Exit Sub

ERRH:
2120    intRetVal = 0
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

2210    With Me
2220      .chkvoids_datemodified = Now()
2230    End With

EXITP:
2240    Exit Sub

ERRH:
2250    Select Case ERR.Number
        Case Else
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

End Sub

Public Sub cmdSave_Click()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2310    DoCmd.RunCommand acCmdSaveRecord

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case 3022  ' ** The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.
          ' ** Ignore! I don't know why it happens on some machines and not mine.
2340    Case Else
2350      DoCmd.Hourglass False
2360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2370    End Select
2380    Resume EXITP

End Sub

Private Sub chkvoid_chknum_lbl_DblClick(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_chknum_lbl_DblClick"

2410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub chkvoid_chknum_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_chknum_GotFocus"

2510    With Me
2520  On Error Resume Next
          ' ** 2185  You can't reference a property or method for a control unless the control has the focus (no records).
2530      .chkvoid_chknum.SelLength = 0
2540      .chkvoid_chknum.SelStart = 19
2550  On Error GoTo ERRH
2560    End With

EXITP:
2570    Exit Sub

ERRH:
2580    Select Case ERR.Number
        Case Else
2590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2600    End Select
2610    Resume EXITP

End Sub

Private Sub chkvoid_chknum_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_chknum_KeyDown"

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
2790          .chkvoid_void.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2850        With Me
2860          intRetVal = 0
2870          If .CurrentRecord > 1 Then
2880            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2890            .chkbank_acctnum.SetFocus
2900          Else
2910            DoCmd.SelectObject acForm, .Parent.Name, False
2920            .Parent.cmdOK.SetFocus
2930          End If
2940        End With
2950      End Select
2960    End If

EXITP:
2970    KeyCode = intRetVal
2980    Exit Sub

ERRH:
2990    intRetVal = 0
3000    Select Case ERR.Number
        Case Else
3010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3020    End Select
3030    Resume EXITP

End Sub

Private Sub chkvoid_void_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_void_lbl_DblClick"

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

Public Sub chkvoid_void_AfterUpdate()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_void_AfterUpdate"

3210    With Me
3220      .changed = True
3230      cmdSave_Click  ' ** Procedure: Above.
3240      DoEvents
3250      RecalcTots  ' ** Procedure: Below.
3260      DoEvents
3270      If IsNull(.chkbank_name) = True Or IsNull(.chkbank_acctnum) = True Then
3280        .Parent.MsgSubhead_lbl1.Visible = True
3290        .Parent.MsgSubhead_lbl2.Visible = True
3300      Else
3310        .Parent.MsgSubhead_lbl1.Visible = False
3320        .Parent.MsgSubhead_lbl2.Visible = False
3330      End If
3340    End With

EXITP:
3350    Exit Sub

ERRH:
3360    Select Case ERR.Number
        Case Else
3370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3380    End Select
3390    Resume EXITP

End Sub

Private Sub chkvoid_void_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_void_KeyDown"

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
3490          .accountno.SetFocus
3500        End With
3510      End Select
3520    End If

        ' ** Shift keys.
3530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3550        With Me
3560          intRetVal = 0
3570          .chkvoid_chknum.SetFocus
3580        End With
3590      End Select
3600    End If

EXITP:
3610    KeyCode = intRetVal
3620    Exit Sub

ERRH:
3630    intRetVal = 0
3640    Select Case ERR.Number
        Case Else
3650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3660    End Select
3670    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

3710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub accountno_GotFocus()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_GotFocus"

3810    With Me
3820      .accountno.SelLength = 0
3830      .accountno.SelStart = 19
3840    End With

EXITP:
3850    Exit Sub

ERRH:
3860    Select Case ERR.Number
        Case Else
3870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3880    End Select
3890    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

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

Private Sub shortname_GotFocus()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_GotFocus"

4010    With Me
4020      .shortname.SelLength = 0
4030      .shortname.SelStart = 0
4040    End With

EXITP:
4050    Exit Sub

ERRH:
4060    Select Case ERR.Number
        Case Else
4070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4080    End Select
4090    Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

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

Private Sub transdate_GotFocus()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_GotFocus"

4210    With Me
4220      .transdate.SelLength = 0
4230      .transdate.SelStart = 19
4240    End With

EXITP:
4250    Exit Sub

ERRH:
4260    Select Case ERR.Number
        Case Else
4270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4280    End Select
4290    Resume EXITP

End Sub

Private Sub chkvoid_payee_lbl_DblClick(Cancel As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_payee_lbl_DblClick"

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

Private Sub chkvoid_payee_GotFocus()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_payee_GotFocus"

4410    With Me
4420      .chkvoid_payee.SelLength = 0
4430      .chkvoid_payee.SelStart = 0
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

Private Sub chkvoid_amount_lbl_DblClick(Cancel As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_amount_lbl_DblClick"

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

Private Sub chkvoid_amount_GotFocus()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_amount_GotFocus"

4610    With Me
4620      .chkvoid_amount.SelLength = 0
4630      .chkvoid_amount.SelStart = 0
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

Private Sub chkbank_name_lbl_DblClick(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_name_lbl_DblClick"

4710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4720    Exit Sub

ERRH:
4730    Select Case ERR.Number
        Case Else
4740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4750    End Select
4760    Resume EXITP

End Sub

Private Sub chkbank_name_GotFocus()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_name_GotFocus"

4810    With Me
4820      .chkbank_name.SelLength = 0
4830      .chkbank_name.SelStart = 0
4840    End With

EXITP:
4850    Exit Sub

ERRH:
4860    Select Case ERR.Number
        Case Else
4870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4880    End Select
4890    Resume EXITP

End Sub

Private Sub chkbank_name_AfterUpdate()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_name_AfterUpdate"

4910    With Me

4920      .changed = True

4930      cmdSave_Click  ' ** Procedure: Above.

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

Private Sub chkbank_name_KeyDown(KeyCode As Integer, Shift As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_name_KeyDown"

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
5090          .chkbank_acctnum.SetFocus
5100        End With
5110      End Select
5120    End If

        ' ** Shift keys.
5130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5150        With Me
5160          intRetVal = 0
5170          .chkvoid_amount.SetFocus
5180        End With
5190      End Select
5200    End If

EXITP:
5210    KeyCode = intRetVal
5220    Exit Sub

ERRH:
5230    intRetVal = 0
5240    Select Case ERR.Number
        Case Else
5250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5260    End Select
5270    Resume EXITP

End Sub

Private Sub chkbank_acctnum_lbl_DblClick(Cancel As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_acctnum_lbl_DblClick"

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

Private Sub chkbank_acctnum_GotFocus()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_acctnum_GotFocus"

5410    With Me
5420      .chkbank_acctnum.SelLength = 0
5430      .chkbank_acctnum.SelStart = 19
5440    End With

EXITP:
5450    Exit Sub

ERRH:
5460    Select Case ERR.Number
        Case Else
5470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5480    End Select
5490    Resume EXITP

End Sub

Private Sub chkbank_acctnum_AfterUpdate()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_acctnum_AfterUpdate"

5510    With Me

5520      .changed = True

5530      cmdSave_Click  ' ** Procedure: Above.

5540    End With

EXITP:
5550    Exit Sub

ERRH:
5560    Select Case ERR.Number
        Case Else
5570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5580    End Select
5590    Resume EXITP

End Sub

Private Sub chkbank_acctnum_KeyDown(KeyCode As Integer, Shift As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_acctnum_KeyDown"

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
5690          lngRecsCur = RecCnt  ' ** Function: Below.
5700          If .CurrentRecord < lngRecsCur Then
5710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5720            .chkvoid_chknum.SetFocus
5730          Else
5740            DoCmd.SelectObject acForm, .Parent.Name, False
5750            .Parent.cmdOK.SetFocus
5760          End If
5770        End With
5780      End Select
5790    End If

        ' ** Shift keys.
5800    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5810      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5820        With Me
5830          intRetVal = 0
5840          .chkbank_name.SetFocus
5850        End With
5860      End Select
5870    End If

EXITP:
5880    KeyCode = intRetVal
5890    Exit Sub

ERRH:
5900    intRetVal = 0
5910    Select Case ERR.Number
        Case Else
5920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5930    End Select
5940    Resume EXITP

End Sub

Public Sub RecalcTots()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

6010    With Me
6020      .Recalc
6030      .chkvoid_void_sum.Requery
6040      DoEvents
6050  On Error Resume Next
          ' ** 2113: The value you entered isn't valid for this field (no records).
6060      .Parent.chkvoid_void_sum = .chkvoid_void_sum
6070  On Error GoTo ERRH
6080    End With

EXITP:
6090    Exit Sub

ERRH:
6100    Select Case ERR.Number
        Case Else
6110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6120    End Select
6130    Resume EXITP

End Sub

Public Function RecCnt() As Long

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

6210    lngRetVal = 0&

6220    With Me
6230      Set rst = .RecordsetClone
6240      With rst
6250        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
6260        Else
6270          .MoveLast
6280          lngRetVal = .RecordCount
6290        End If
6300        .Close
6310      End With
6320    End With

EXITP:
6330    Set rst = Nothing
6340    RecCnt = lngRetVal
6350    Exit Function

ERRH:
6360    lngRetVal = 0&
6370    Select Case ERR.Number
        Case Else
6380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6390    End Select
6400    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

6510    Select Case intWhere
        Case acCmdRecordsGoToFirst
6520      DoCmd.RunCommand acCmdRecordsGoToFirst
6530    Case acCmdRecordsGoToLast
6540      DoCmd.RunCommand acCmdRecordsGoToLast
6550    Case acCmdRecordsGoToPrevious
6560      DoCmd.RunCommand acCmdRecordsGoToPrevious
6570    Case acCmdRecordsGoToNext
6580      DoCmd.RunCommand acCmdRecordsGoToNext
6590    Case Else
6600      If IsMissing(varID) = False Then
6610        With Me
6620          Set rst = .RecordsetClone
6630          With rst
6640            .FindFirst "[chkvoid_chknum] = " & CStr(varID)
6650            If .NoMatch = False Then
6660              Me.Bookmark = .Bookmark
6670            End If
6680            .Close
6690          End With
6700        End With
6710      End If
6720    End Select

EXITP:
6730    Set rst = Nothing
6740    Exit Sub

ERRH:
6750    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
6760    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
6770    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
6780    Case Else
6790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6800    End Select
6810    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [chkvoid_chknum]"

6910    With Me
6920      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
6930        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6940      End If
6950      .Controls(strSortLbl).Visible = False
6960      .Controls(strSortLine).Visible = False
6970      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form
6980      If strProc = "Form_Load" Then
6990        strCalled = "chkvoid_chknum"
7000        strSortNow = strSortOrig
7010        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7020        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7030        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7040        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7050        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7060        .Controls(strSortLbl).Left = lngSortLbl_Left
7070        .Controls(strSortLbl).Top = lngSortLbl_Top
7080        .Controls(strSortLine).Top = lngSortLine_Top
7090        .Controls(strSortLine).Left = lngSortLine_Left
7100        .Controls(strSortLine).Width = lngSortLine_Width
7110        .Controls(strSortLbl).Caption = strArwUp
7120        .Controls(strSortLbl).ForeColor = CLR_DKBLU
7130      Else
7140        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
7150        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7160        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7170        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7180        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7190        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7200        .Controls(strSortLbl).Left = lngSortLbl_Left
7210        .Controls(strSortLbl).Top = lngSortLbl_Top
7220        .Controls(strSortLine).Top = lngSortLine_Top
7230        .Controls(strSortLine).Left = lngSortLine_Left
7240        .Controls(strSortLine).Width = lngSortLine_Width
7250        If strCalled = "chkvoid_chknum" Then
7260          strSortAsc = strSortOrig
7270          If strSortNow = strSortAsc Then
7280            strSortNow = "[chkvoid_chknum] DESC"
7290            .Controls(strSortLbl).Caption = strArwDn
7300            .Controls(strSortLbl).ForeColor = CLR_DKRED
7310          Else
7320            strSortNow = strSortAsc
7330            .Controls(strSortLbl).Caption = strArwUp
7340            .Controls(strSortLbl).ForeColor = CLR_DKBLU
7350          End If
7360        Else
7370          .Controls(strSortLbl).Caption = strArwUp
7380          .Controls(strSortLbl).ForeColor = CLR_DKBLU
7390          Select Case strCalled
              Case "chkvoid_void"
7400            strSortAsc = "[chkvoid_void]" & strStdAsc
7410            lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
7420            .Controls(strSortLbl).Left = lngSortLbl_Left
7430          Case "accountno"
7440            strSortAsc = "[accountno]" & strStdAsc
7450          Case "shortname"
7460            strSortAsc = "[shortname]" & strStdAsc
7470          Case "transdate"
7480            strSortAsc = "[transdate]" & strStdAsc
7490          Case "chkvoid_payee"
7500            strSortAsc = "[chkvoid_payee]" & strStdAsc
7510          Case "chkbank_name"
7520            strSortAsc = "[chkbank_name]" & strStdAsc
7530          Case "chkbank_acctnum"
7540            strSortAsc = "[chkbank_acctnum]" & strStdAsc
7550          End Select
7560          If strSortNow = strSortAsc Then
7570            intCnt = CharCnt(strSortAsc, ",") + 1
7580            Select Case intCnt
                Case 1
7590              strTmp01 = strSortAsc & " DESC"
7600            Case 2
7610              intPos01 = InStr(strSortAsc, ",")
7620              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
7630              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
7640            Case 3
7650              intPos01 = InStr(strSortAsc, ",")
7660              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
7670              strTmp02 = Mid(strSortAsc, intPos01)
7680              intPos01 = InStr(2, strTmp02, ",")
7690              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
7700              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
7710            End Select
7720            strSortNow = strTmp01
7730            .Controls(strSortLbl).Caption = strArwDn
7740            .Controls(strSortLbl).ForeColor = CLR_DKRED
7750          Else
7760            strSortNow = strSortAsc
7770          End If
7780        End If
7790      End If
7800      .Controls(strSortLbl).Visible = True
7810      .Controls(strSortLine).Visible = True
7820      .OrderBy = strSortNow
7830      .OrderByOn = True
7840      .TimerInterval = 500&
7850    End With

        'chkvoid_chknum
        'chkvoid_void
        'accountno
        'shortname
        'transdate
        'chkvoid_payee
        'chkbank_name
        'chkbank_acctnum

EXITP:
7860    Exit Sub

ERRH:
7870    Select Case ERR.Number
        Case Else
7880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7890    End Select
7900    Resume EXITP

End Sub
