VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_CourtReports_CA_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_CourtReports_CA_Input"

'VGC 01/16/2012: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Final:            F {opgVer_optFinal}
' **   Interim:          I {opgVer_optInterim}
' **   OK:               O {cmdOK}
' **   Summary Ord/Ver:  1 {cmbOrd}
' **   Beg Cash:         2 {CashAssets_Beg}
' **   Net Income:       3 {NetIncome}
' **   Net Loss:         4 {NetLoss}
' **   End Cash:         5 {CashAssets_End}

Private blnOrderlyClose As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim ctl As Access.Control
        Dim intX As Integer

110     blnOrderlyClose = False

120     With Me
130       If IsNull(.OpenArgs) = False Then
140         If Trim(.OpenArgs) <> vbNullString Then

150           strCallingForm = Trim(.OpenArgs)

160           .PictureData = .Form_Rpt_CourtReports_CA_Input_img

170           .cmbOrd = .cmbOrd.Column(0, 0)

180           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

190           If gstrCrtRpt_Ordinal <> vbNullString Then
200             For intX = 0 To (.cmbOrd.ListCount - 1)
210               If .cmbOrd.Column(1, intX) = gstrCrtRpt_Ordinal Then
220                 .cmbOrd = .cmbOrd.Column(0, intX)
230                 Exit For
240               End If
250             Next
260           End If
270           gstrCrtRpt_Ordinal = .cmbOrd.Column(1)

280           If IsNumeric(gstrCrtRpt_Version) = True Then gstrCrtRpt_Version = vbNullString
290           If gstrCrtRpt_Version <> vbNullString Then
300             For Each ctl In .Controls
310               With ctl
320                 If .ControlType = acOptionButton Then
330                   If .Parent.Name = Me.opgVer.Name Then
340                     If InStr(.Name, gstrCrtRpt_Version) > 0 Then
350                       Me.opgVer = .OptionValue
360                       Exit For
370                     End If
380                   End If
390                 End If
400               End With
410             Next
420           Else
430             .opgVer = .opgVer.DefaultValue
440             For Each ctl In .Controls
450               With ctl
460                 If .ControlType = acOptionButton Then
470                   If .Parent.Name = Me.opgVer.Name Then
480                     If .OptionValue = Me.opgVer.DefaultValue Then
490                       gstrCrtRpt_Version = Mid(.Name, (InStr(.Name, "Opt") + 3))
500                       Exit For
510                     End If
520                   End If
530                 End If
540               End With
550             Next
560           End If

570           opgVer_AfterUpdate  ' ** Procedure: Below.

580           If gstrCrtRpt_CashAssets_Beg = vbNullString Then
590             .CashAssets_Beg = 0
600             gstrCrtRpt_CashAssets_Beg = "0"
610           Else
620             .CashAssets_Beg = Val(gstrCrtRpt_CashAssets_Beg)
630           End If
640           .CashAssets_Beg_display_lbl.Caption = Format(.CashAssets_Beg, "Currency")

650           If gstrCrtRpt_NetIncome = vbNullString Then
660             .NetIncome = 0
670             gstrCrtRpt_NetIncome = "0"
680           Else
690             .NetIncome = Val(gstrCrtRpt_NetIncome)
700           End If
710           .NetIncome_display_lbl.Caption = Format(.NetIncome, "Currency")

720           If gstrCrtRpt_NetLoss = vbNullString Then
730             .NetLoss = 0
740             gstrCrtRpt_NetLoss = "0"
750           Else
760             .NetLoss = Val(gstrCrtRpt_NetLoss)
770           End If
780           .NetLoss_display_lbl.Caption = Format(.NetLoss, "Currency")

790           If gstrCrtRpt_CashAssets_End = vbNullString Then
800             .CashAssets_End = 0
810             gstrCrtRpt_CashAssets_End = "0"
820           Else
830             .CashAssets_End = Val(gstrCrtRpt_CashAssets_End)
840           End If
850           .CashAssets_End_display_lbl.Caption = Format(.CashAssets_End, "Currency")

860           .cmdOK.SetFocus

870         Else
880           Cancel = -1
890         End If
900       Else
910         Cancel = -1
920       End If
930     End With

940     If Cancel = -1 Then
950       blnOrderlyClose = True
960       gblnMessage = False
970       Beep
980       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
990     End If

EXITP:
1000    Exit Sub

ERRH:
1010    Select Case ERR.Number
        Case Else
1020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1030    End Select
1040    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1120    intShiftDown = (Shift And acShiftMask) > 0
1130    intAltDown = (Shift And acAltMask) > 0
1140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyEscape
1170        intRetVal = 0
1180        cmdCancel_Click  ' ** Procedure: Below.
1190      End Select
1200    End If

        ' ** Alt keys.
1210    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1220      Select Case intRetVal
          Case vbKey1
1230        With Me
1240          intRetVal = 0
1250          .cmbOrd.SetFocus
1260        End With
1270      Case vbKey2
1280        With Me
1290          intRetVal = 0
1300          .CashAssets_Beg.SetFocus
1310        End With
1320      Case vbKey3
1330        With Me
1340          intRetVal = 0
1350          .NetIncome.SetFocus
1360        End With
1370      Case vbKey4
1380        With Me
1390          intRetVal = 0
1400          .NetLoss.SetFocus
1410        End With
1420      Case vbKey5
1430        With Me
1440          intRetVal = 0
1450          .CashAssets_End.SetFocus
1460        End With
1470      End Select
1480    End If

EXITP:
1490    KeyCode = intRetVal
1500    Exit Sub

ERRH:
1510    intRetVal = 0
1520    Select Case ERR.Number
        Case Else
1530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1540    End Select
1550    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1610    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1620    If blnOrderlyClose = False Then
1630      gblnMessage = False
1640    End If

EXITP:
1650    Exit Sub

ERRH:
1660    Select Case ERR.Number
        Case Else
1670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1680    End Select
1690    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOk_Click"

        Dim blnContinue As Boolean

1710    blnContinue = True

1720    With Forms("frmRpt_CourtReports_CA")
1730      .Ordinal = Me.cmbOrd.Column(1)
1740      Select Case Me.opgVer
          Case Me.opgVer_optFinal.OptionValue
1750        .Version = "Final"
1760      Case Me.opgVer_optInterim.OptionValue
1770        .Version = "Interim"
1780      End Select
1790    End With
1800    With Me
1810      If IsNumeric(.CashAssets_Beg) = False Then
1820        blnContinue = False
1830        Beep
1840        MsgBox "Beginning Cash Assets must be a numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
1850        .CashAssets_Beg.SetFocus
1860      ElseIf IsNumeric(.NetIncome) = False Then
1870        blnContinue = False
1880        Beep
1890        MsgBox "Net Income must be a numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
1900        .NetIncome.SetFocus
1910      ElseIf IsNumeric(.NetLoss) = False Then
1920        blnContinue = False
1930        Beep
1940        MsgBox "Net Loss must be a numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
1950        .NetLoss.SetFocus
1960      ElseIf IsNumeric(.CashAssets_End) = False Then
1970        blnContinue = False
1980        Beep
1990        MsgBox "Ending Cash Assets must be a numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
2000        .CashAssets_End.SetFocus
2010      End If
2020    End With

2030    If blnContinue = True Then
2040      blnOrderlyClose = True
2050      DoCmd.Close acForm, THIS_NAME
2060      DoCmd.SelectObject acForm, "frmRpt_CourtReports_CA", False
2070    End If

EXITP:
2080    Exit Sub

ERRH:
2090    Select Case ERR.Number
        Case Else
2100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2110    End Select
2120    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2220    intShiftDown = (Shift And acShiftMask) > 0
2230    intAltDown = (Shift And acAltMask) > 0
2240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2260      Select Case intRetVal
          Case vbKeyTab
2270        With Me
2280          intRetVal = 0
2290          .cmdCancel.SetFocus
2300        End With
2310      End Select
2320    End If

        ' ** Shift keys.
2330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyTab
2350        With Me
2360          intRetVal = 0
2370          .CashAssets_End.SetFocus
2380        End With
2390      End Select
2400    End If

EXITP:
2410    KeyCode = intRetVal
2420    Exit Sub

ERRH:
2430    intRetVal = 0
2440    Select Case ERR.Number
        Case Else
2450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2460    End Select
2470    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

2510    blnOrderlyClose = True
2520    gblnMessage = False
2530    With Forms("frmRpt_CourtReports_CA")
2540      .Ordinal = vbNullString  ' ** But leave the Public variable.
2550      .Version = vbNullString  ' ** But leave the Public variable.
2560    End With
2570    DoCmd.Close acForm, THIS_NAME
2580    DoCmd.SelectObject acForm, "frmRpt_CourtReports_CA", False

EXITP:
2590    Exit Sub

ERRH:
2600    Select Case ERR.Number
        Case Else
2610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2620    End Select
2630    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
          Case vbKeyTab
2770        With Me
2780          intRetVal = 0
2790          .cmbOrd.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab
2850        With Me
2860          intRetVal = 0
2870          .cmdOK.SetFocus
2880        End With
2890      End Select
2900    End If

EXITP:
2910    KeyCode = intRetVal
2920    Exit Sub

ERRH:
2930    intRetVal = 0
2940    Select Case ERR.Number
        Case Else
2950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2960    End Select
2970    Resume EXITP

End Sub

Private Sub cmbOrd_GotFocus()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_GotFocus"

3010    With Me
3020      .cmbOrd.SelLength = 0
3030      .cmbOrd.SelStart = 99
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

Private Sub cmbOrd_AfterUpdate()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_AfterUpdate"

3110    With Me
3120      If IsNull(.cmbOrd) = False Then
3130        gstrCrtRpt_Ordinal = .cmbOrd.Column(1)
3140      Else
3150        .cmbOrd = .cmbOrd.DefaultValue
3160        gstrCrtRpt_Ordinal = .cmbOrd.Column(1)
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

Private Sub cmbOrd_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_KeyDown"

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
3390          .opgVer.SetFocus
3400        End With
3410      End Select
3420    End If

        ' ** Shift keys.
3430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3450        With Me
3460          intRetVal = 0
3470          .cmdCancel.SetFocus
3480        End With
3490      End Select
3500    End If

EXITP:
3510    KeyCode = intRetVal
3520    Exit Sub

ERRH:
3530    intRetVal = 0
3540    Select Case ERR.Number
        Case Else
3550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3560    End Select
3570    Resume EXITP

End Sub

Private Sub opgVer_AfterUpdate()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgVer_AfterUpdate"

3610    With Me
3620      Select Case .opgVer
          Case .opgVer_optFinal.OptionValue
3630        .opgVer_optFinal_lbl.FontBold = True
3640        .opgVer_optInterim_lbl.FontBold = False
3650        gstrCrtRpt_Version = Mid(.opgVer_optFinal.Name, (InStr(.opgVer_optFinal.Name, "Opt") + 3))
3660      Case .opgVer_optInterim.OptionValue
3670        .opgVer_optFinal_lbl.FontBold = False
3680        .opgVer_optInterim_lbl.FontBold = True
3690        gstrCrtRpt_Version = Mid(.opgVer_optInterim.Name, (InStr(.opgVer_optInterim.Name, "Opt") + 3))
3700      End Select
3710    End With

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub CashAssets_Beg_AfterUpdate()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "CashAssets_Beg_AfterUpdate"

3810    With Me
3820      If IsNull(.CashAssets_Beg) = False Then
3830        If IsNumeric(.CashAssets_Beg) = True Then
3840          gstrCrtRpt_CashAssets_Beg = CStr(.CashAssets_Beg)
3850          .CashAssets_Beg_display_lbl.Caption = Format(.CashAssets_Beg, "Currency")
3860        Else
3870          MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
3880        End If
3890      Else
3900        .CashAssets_Beg = 0&
3910        gstrCrtRpt_CashAssets_Beg = "0"
3920        .CashAssets_Beg_display_lbl.Caption = Format(.CashAssets_Beg, "Currency")
3930      End If
3940    End With

EXITP:
3950    Exit Sub

ERRH:
3960    Select Case ERR.Number
        Case Else
3970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3980    End Select
3990    Resume EXITP

End Sub

Private Sub NetIncome_AfterUpdate()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "NetIncome_AfterUpdate"

4010    With Me
4020      If IsNull(.NetIncome) = False Then
4030        If IsNumeric(.NetIncome) = True Then
4040          gstrCrtRpt_NetIncome = CStr(.NetIncome)
4050          .NetIncome_display_lbl.Caption = Format(.NetIncome, "Currency")
4060        Else
4070          MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
4080        End If
4090      Else
4100        .NetIncome = 0
4110        gstrCrtRpt_NetIncome = "0"
4120        .NetIncome_display_lbl.Caption = Format(.NetIncome, "Currency")
4130      End If
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

Private Sub NetLoss_AfterUpdate()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "NetLoss_AfterUpdate"

4210    With Me
4220      If IsNull(.NetLoss) = False Then
4230        If IsNumeric(.NetLoss) = True Then
4240          gstrCrtRpt_NetLoss = CStr(.NetLoss)
4250          .NetLoss_display_lbl.Caption = Format(.NetLoss, "Currency")
4260        Else
4270          MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
4280        End If
4290      Else
4300        .NetLoss = 0
4310        gstrCrtRpt_NetLoss = "0"
4320        .NetLoss_display_lbl.Caption = Format(.NetLoss, "Currency")
4330      End If
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

Private Sub CashAssets_End_AfterUpdate()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "CashAssets_End_AfterUpdate"

4410    With Me
4420      If IsNull(.CashAssets_End) = False Then
4430        If IsNumeric(.CashAssets_End) = True Then
4440          gstrCrtRpt_CashAssets_End = CStr(.CashAssets_End)
4450          .CashAssets_End_display_lbl.Caption = Format(.CashAssets_End, "Currency")
4460        Else
4470          MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
4480        End If
4490      Else
4500        .CashAssets_End = 0
4510        gstrCrtRpt_CashAssets_End = "0"
4520        .CashAssets_End_display_lbl.Caption = Format(.CashAssets_End, "Currency")
4530      End If
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

Private Sub CashAssets_End_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "CashAssets_End_KeyDown"

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
4690          .cmdOK.SetFocus
4700        End With
4710      End Select
4720    End If

        ' ** Shift keys.
4730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4750        With Me
4760          intRetVal = 0
4770          .NetLoss.SetFocus
4780        End With
4790      End Select
4800    End If

EXITP:
4810    KeyCode = intRetVal
4820    Exit Sub

ERRH:
4830    intRetVal = 0
4840    Select Case ERR.Number
        Case Else
4850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4860    End Select
4870    Resume EXITP

End Sub
