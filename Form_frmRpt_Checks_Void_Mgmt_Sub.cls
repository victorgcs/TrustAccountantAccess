VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks_Void_Mgmt_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks_Void_Mgmt_Sub"

'VGC 07/13/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Bank Acct Number: A {chkbank_acctnum}
' **   Account Number:   B {accountno}
' **   Close:            C {cmdClose on frmRpt_Checks_Void_Mgmt}
' **   Bank Name:        E {chkbank_name}
' **   Check Number:     K {chkvoid_chknum}
' **   Account Name:     M {shortname}
' **   Print Report:     P {cmdPrintReport on frmRpt_Checks_Void_Mgmt}
' **   Exit:             X {cmdClose on frmRpt_Checks_Void_Mgmt}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[chkbank_name], [chkvoid_chknum]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strSortBox As String = "Sort_box"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private strSortNow As String, lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private lngSortBox_Top As Long, lngSortBox_Left As Long

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
150       If strCallingForm = "frmRpt_Checks_Void_Mgmt" Then

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

420       .chkbank_name.ForeColor = CLR_DISABLED_FG
430       .chkbank_name.BackColor = CLR_DISABLED_BG
440       .chkbank_acctnum.ForeColor = CLR_DISABLED_FG
450       .chkbank_acctnum.BackColor = CLR_DISABLED_BG
460       .chkvoid_chknum.ForeColor = CLR_DISABLED_FG
470       .chkvoid_chknum.BackColor = CLR_DISABLED_BG
480       .chkvoid_date.ForeColor = CLR_DISABLED_FG
490       .chkvoid_date.BackColor = CLR_DISABLED_BG
500       .accountno.ForeColor = CLR_DISABLED_FG
510       .accountno.BackColor = CLR_DISABLED_BG
520       .shortname.ForeColor = CLR_DISABLED_FG
530       .shortname.BackColor = CLR_DISABLED_BG
540       .transdate.ForeColor = CLR_DISABLED_FG
550       .transdate.BackColor = CLR_DISABLED_BG
560       .chkvoid_payee.ForeColor = CLR_DISABLED_FG
570       .chkvoid_payee.BackColor = CLR_DISABLED_BG
580       .chkvoid_amount.ForeColor = CLR_DISABLED_FG
590       .chkvoid_amount.BackColor = CLR_DISABLED_BG

600       SortNow THIS_PROC  ' ** Procedure: Below.

610     End With

EXITP:
620     Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

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
        ' **   Bank Acct Number: A {chkbank_acctnum}
        ' **   Account Number:   B {accountno}
        ' **   Bank Name:        E {chkbank_name}
        ' **   Check Number:     K {chkvoid_chknum}
        ' **   Account Name:     M {shortname}
        ' **   Exit:             X {cmdClose on frmRpt_Checks_Void_Mgmt}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyUp
770         intRetVal = 0
780         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
790       Case vbKeyDown
800         intRetVal = 0
810         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
820       Case vbKeyEscape
830         With Me
840           intRetVal = 0
850           .Parent.cmdClose_Click  ' ** Form Procedure: frmRpt_Checks_Void_Mgmt.
860         End With
870       End Select
880     End If

        ' ** Alt keys.
890     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
900       Select Case intRetVal
          Case vbKeyA
910         With Me
920           intRetVal = 0
930           .chkbank_acctnum.SetFocus
940         End With
950       Case vbKeyB
960         With Me
970           intRetVal = 0
980           .accountno.SetFocus
990         End With
1000      Case vbKeyE
1010        With Me
1020          intRetVal = 0
1030          .chkbank_name.SetFocus
1040        End With
1050      Case vbKeyK
1060        With Me
1070          intRetVal = 0
1080          .chkvoid_chknum.SetFocus
1090        End With
1100      Case vbKeyM
1110        With Me
1120          intRetVal = 0
1130          .shortname.SetFocus
1140        End With
1150      Case vbKeyX
1160        With Me
1170          intRetVal = 0
1180          .Parent.cmdClose_Click  ' ** Form Procedure: frmRpt_Checks_Void_Mgmt.
1190        End With
1200      End Select
1210    End If

        ' ** Ctrl keys.
1220    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1230      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1240        With Me
1250          intRetVal = 0
1260          DoCmd.SelectObject acForm, .Parent.Name, False
1270          If .Parent.cmdPrintReport.Enabled = True Then
1280            .Parent.cmdPrintReport.SetFocus
1290          Else
1300            .Parent.cmdClose.SetFocus
1310          End If
1320        End With
1330      Case vbKeyPageUp, vbKeyHome
1340        intRetVal = 0
1350        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1360      Case vbKeyPageDown, vbKeyEnd
1370        intRetVal = 0
1380        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
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
        Case Else
1600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1610    End Select
1620    Resume EXITP

End Sub

Private Sub chkbank_name_lbl_DblClick(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_name_lbl_DblClick"

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

Private Sub chkbank_name_lbl_img_cmd_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_name_lbl_img_cmd_DblClick"

1810    chkbank_name_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub chkbank_name_GotFocus()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_name_GotFocus"

1910    With Me
1920      .chkbank_name.SelLength = 0
1930      .chkbank_name.SelStart = 0
1940    End With

EXITP:
1950    Exit Sub

ERRH:
1960    Select Case ERR.Number
        Case Else
1970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1980    End Select
1990    Resume EXITP

End Sub

Private Sub chkbank_name_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2020    intShiftDown = (Shift And acShiftMask) > 0
2030    intAltDown = (Shift And acAltMask) > 0
2040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2070        With Me
2080          intRetVal = 0
2090          .chkbank_acctnum.SetFocus
2100        End With
2110      End Select
2120    End If

        ' ** Shift keys.
2130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2150        With Me
2160          intRetVal = 0
2170          If .CurrentRecord > 1 Then
2180            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2190            .chkvoid_amount.SetFocus
2200          Else
2210            DoCmd.SelectObject acForm, .Parent.Name, False
2220            .Parent.cmdClose.SetFocus
2230          End If
2240        End With
2250      End Select
2260    End If

EXITP:
2270    KeyCode = intRetVal
2280    Exit Sub

ERRH:
2290    intRetVal = 0
2300    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2310    Case Else
2320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2330    End Select
2340    Resume EXITP

End Sub

Private Sub chkbank_acctnum_lbl_DblClick(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_acctnum_lbl_DblClick"

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

Private Sub chkbank_acctnum_lbl_img_cmd_DblClick(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_acctnum_lbl_img_cmd_DblClick"

2510    chkbank_acctnum_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub chkbank_acctnum_GotFocus()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_acctnum_GotFocus"

2610    With Me
2620      .chkbank_acctnum.SelLength = 0
2630      .chkbank_acctnum.SelStart = 19
2640    End With

EXITP:
2650    Exit Sub

ERRH:
2660    Select Case ERR.Number
        Case Else
2670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2680    End Select
2690    Resume EXITP

End Sub

Private Sub chkvoid_chknum_lbl_DblClick(Cancel As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_chknum_lbl_DblClick"

2710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2720    Exit Sub

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub chkvoid_chknum_lbl_img_cmd_DblClick(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_chknum_lbl_img_cmd_DblClick"

2810    chkvoid_chknum_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub chkvoid_chknum_GotFocus()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_chknum_GotFocus"

2910    With Me
2920      .chkvoid_chknum.SelLength = 0
2930      .chkvoid_chknum.SelStart = 19
2940    End With

EXITP:
2950    Exit Sub

ERRH:
2960    Select Case ERR.Number
        Case Else
2970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2980    End Select
2990    Resume EXITP

End Sub

Private Sub chkvoid_date_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_date_lbl_DblClick"

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

Private Sub chkvoid_date_lbl_img_cmd_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_date_lbl_img_cmd_DblClick"

3110    chkvoid_date_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub chkvoid_date_GotFocus()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_date_GotFocus"

3210    With Me
3220      .chkvoid_date.SelLength = 0
3230      .chkvoid_date.SelStart = 19
3240    End With

EXITP:
3250    Exit Sub

ERRH:
3260    Select Case ERR.Number
        Case Else
3270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3280    End Select
3290    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

3310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub accountno_lbl_img_cmd_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_img_cmd_DblClick"

3410    accountno_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub accountno_GotFocus()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_GotFocus"

3510    With Me
3520      .accountno.SelLength = 0
3530      .accountno.SelStart = 19
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

Private Sub shortname_lbl_DblClick(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

3610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub shortname_lbl_img_cmd_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_img_cmd_DblClick"

3710    shortname_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub shortname_GotFocus()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_GotFocus"

3810    With Me
3820      .shortname.SelLength = 0
3830      .shortname.SelStart = 0
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

Private Sub transdate_lbl_DblClick(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

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

Private Sub transdate_lbl_img_cmd_DblClick(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_img_cmd_DblClick"

4010    transdate_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub transdate_GotFocus()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_GotFocus"

4110    With Me
4120      .transdate.SelLength = 0
4130      .transdate.SelStart = 19
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

Private Sub chkvoid_payee_lbl_DblClick(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_payee_lbl_DblClick"

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

Private Sub chkvoid_payee_lbl_img_cmd_DblClick(Cancel As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_payee_lbl_img_cmd_DblClick"

4310    chkvoid_payee_lbl_DblClick Cancel  ' ** Procedure: Above.

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

Private Sub chkvoid_amount_lbl_img_cmd_DblClick(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_amount_lbl_img_cmd_DblClick"

4610    chkvoid_amount_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub chkvoid_amount_GotFocus()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_amount_GotFocus"

4710    With Me
4720      .chkvoid_amount.SelLength = 0
4730      .chkvoid_amount.SelStart = 0
4740    End With

EXITP:
4750    Exit Sub

ERRH:
4760    Select Case ERR.Number
        Case Else
4770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4780    End Select
4790    Resume EXITP

End Sub

Private Sub chkvoid_amount_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkvoid_amount_KeyDown"

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
4920            .chkbank_name.SetFocus
4930          Else
4940            DoCmd.SelectObject acForm, .Parent.Name, False
4950            If .Parent.cmdPrintReport.Enabled = True Then
4960              .Parent.cmdPrintReport.SetFocus
4970            Else
4980              .Parent.cmdClose.SetFocus
4990            End If
5000          End If
5010        End With
5020      End Select
5030    End If

        ' ** Shift keys.
5040    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5050      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5060        With Me
5070          intRetVal = 0
5080          .chkvoid_payee.SetFocus
5090        End With
5100      End Select
5110    End If

EXITP:
5120    KeyCode = intRetVal
5130    Exit Sub

ERRH:
5140    intRetVal = 0
5150    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5160    Case Else
5170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5180    End Select
5190    Resume EXITP

End Sub

Public Function RecCnt() As Long

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

5210    lngRetVal = 0&

5220    With Me
5230      Set rst = .RecordsetClone
5240      With rst
5250        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
5260        Else
5270          .MoveLast
5280          lngRetVal = .RecordCount
5290        End If
5300        .Close
5310      End With
5320    End With

EXITP:
5330    Set rst = Nothing
5340    RecCnt = lngRetVal
5350    Exit Function

ERRH:
5360    lngRetVal = 0&
5370    Select Case ERR.Number
        Case Else
5380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5390    End Select
5400    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

5510    Select Case intWhere
        Case acCmdRecordsGoToFirst
5520      DoCmd.RunCommand acCmdRecordsGoToFirst
5530    Case acCmdRecordsGoToLast
5540      DoCmd.RunCommand acCmdRecordsGoToLast
5550    Case acCmdRecordsGoToPrevious
5560      DoCmd.RunCommand acCmdRecordsGoToPrevious
5570    Case acCmdRecordsGoToNext
5580      DoCmd.RunCommand acCmdRecordsGoToNext
5590    Case Else
5600      If IsMissing(varID) = False Then
5610        With Me
5620          Set rst = .RecordsetClone
5630          With rst
5640            .FindFirst "[chkvoid_id] = " & CStr(varID)
5650            If .NoMatch = False Then
5660              Me.Bookmark = .Bookmark
5670            End If
5680            .Close
5690          End With
5700        End With
5710      End If
5720    End Select

EXITP:
5730    Set rst = Nothing
5740    Exit Sub

ERRH:
5750    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
5760    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5770    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5780    Case Else
5790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5800    End Select
5810    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [chkbank_name], [chkvoid_chknum]"

5910    With Me
5920      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5930        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5940      End If
5950      .Controls(strSortLbl).Visible = False
5960      .Controls(strSortLine).Visible = False
5970      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form
5980      .Controls(strSortBox).Visible = False
5990      If strProc = "Form_Load" Then
6000        strCalled = "chkbank_name"
6010        strSortNow = strSortOrig
6020        lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (4& * lngTpp))
6030        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
6040        lngSortBox_Top = .Controls(strCalled & "_lbl_img_cmd").Top
6050        lngSortBox_Left = (((.Controls(strCalled & "_lbl_img_cmd").Left + .Controls(strCalled & "_lbl_img_cmd").Width) - .Controls(strSortBox).Width) - lngTpp)
6060        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
6070        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
6080        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
6090        .Controls(strSortLbl).Left = lngSortLbl_Left
6100        .Controls(strSortLbl).Top = lngSortLbl_Top
6110        .Controls(strSortBox).Top = lngSortBox_Top
6120        .Controls(strSortBox).Left = lngSortBox_Left
6130        .Controls(strSortLine).Top = lngSortLine_Top
6140        .Controls(strSortLine).Left = lngSortLine_Left
6150        .Controls(strSortLine).Width = lngSortLine_Width
6160        .Controls(strSortLbl).Caption = strArwUp
6170        .Controls(strSortLbl).ForeColor = CLR_DKBLU
            'lngSortBox_Top
            'lngSortBox_Left
6180      Else
6190        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
6200        lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (4& * lngTpp))
6210        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
6220        lngSortBox_Top = .Controls(strCalled & "_lbl_img_cmd").Top
6230        lngSortBox_Left = (((.Controls(strCalled & "_lbl_img_cmd").Left + .Controls(strCalled & "_lbl_img_cmd").Width) - .Controls(strSortBox).Width) - lngTpp)
6240        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
6250        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
6260        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
6270        .Controls(strSortLbl).Left = lngSortLbl_Left
6280        .Controls(strSortLbl).Top = lngSortLbl_Top
6290        .Controls(strSortBox).Top = lngSortBox_Top
6300        .Controls(strSortBox).Left = lngSortBox_Left
6310        .Controls(strSortLine).Top = lngSortLine_Top
6320        .Controls(strSortLine).Left = lngSortLine_Left
6330        .Controls(strSortLine).Width = lngSortLine_Width
6340        If strCalled = "chkbank_name" Then
6350          strSortAsc = strSortOrig
6360          If strSortNow = strSortAsc Then
6370            strSortNow = "[chkbank_name] DESC, [chkvoid_chknum] DESC"
6380            .Controls(strSortLbl).Caption = strArwDn
6390            .Controls(strSortLbl).ForeColor = CLR_DKRED
6400          Else
6410            strSortNow = strSortAsc
6420            .Controls(strSortLbl).Caption = strArwUp
6430            .Controls(strSortLbl).ForeColor = CLR_DKBLU
6440          End If
6450        Else
6460          .Controls(strSortLbl).Caption = strArwUp
6470          .Controls(strSortLbl).ForeColor = CLR_DKBLU
6480          Select Case strCalled
              Case "chkbank_acctnum"
6490            strSortAsc = "[chkbank_acctnum]" & strStdAsc
6500          Case "chkvoid_chknum"
6510            strSortAsc = "[chkvoid_chknum]" & strStdAsc
6520          Case "chkvoid_date"
6530            strSortAsc = "[chkvoid_date]" & strStdAsc
6540          Case "accountno"
6550            strSortAsc = "[accountno]" & strStdAsc
6560          Case "shortname"
6570            strSortAsc = "[shortname]" & strStdAsc
6580          Case "transdate"
6590            strSortAsc = "[transdate]" & strStdAsc
6600          Case "chkvoid_payee"
6610            strSortAsc = "[chkvoid_payee]" & strStdAsc
6620          Case "chkvoid_amount"
6630            strSortAsc = "[chkvoid_amount]" & strStdAsc
6640          End Select
6650          If strSortNow = strSortAsc Then
6660            intCnt = CharCnt(strSortAsc, ",") + 1
6670            Select Case intCnt
                Case 1
6680              strTmp01 = strSortAsc & " DESC"
6690            Case 2
6700              intPos01 = InStr(strSortAsc, ",")
6710              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6720              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
6730            Case 3
6740              intPos01 = InStr(strSortAsc, ",")
6750              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6760              strTmp02 = Mid(strSortAsc, intPos01)
6770              intPos01 = InStr(2, strTmp02, ",")
6780              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
6790              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
6800            End Select
6810            strSortNow = strTmp01
6820            .Controls(strSortLbl).Caption = strArwDn
6830            .Controls(strSortLbl).ForeColor = CLR_DKRED
6840          Else
6850            strSortNow = strSortAsc
6860          End If
6870        End If
6880      End If
6890      .Controls(strSortLbl).Visible = True
6900      .Controls(strSortLine).Visible = True
6910      .Controls(strSortBox).Visible = True
6920      .OrderBy = strSortNow
6930      .OrderByOn = True
6940      .TimerInterval = 500&
6950    End With

        'chkbank_name
        'chkbank_acctnum
        'chkvoid_chknum
        'chkvoid_date
        'accountno
        'shortname
        'transdate
        'chkvoid_payee
        'chkvoid_amount

EXITP:
6960    Exit Sub

ERRH:
6970    Select Case ERR.Number
        Case Else
6980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6990    End Select
7000    Resume EXITP

End Sub
