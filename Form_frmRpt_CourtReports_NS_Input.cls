VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_CourtReports_NS_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_CourtReports_NS_Input"

'VGC 12/07/2012: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Final:            F {opgVer_optFinal}
' **   Interim:          I {opgVer_optInterim}
' **   OK:               O {cmdOK}
' **   Summary Ord/Ver:  1 {cmbOrd}
' **   Beg Cash:         2 {CashAssets_Beg}
' **   Net Income:       3 {NetIncome}

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

160           .PictureData = .Form_Rpt_CourtReports_NS_Input_img

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

720           .cmdOK.SetFocus

730         Else
740           Cancel = -1
750         End If
760       Else
770         Cancel = -1
780       End If
790     End With

800     If Cancel = -1 Then
810       blnOrderlyClose = True
820       gblnMessage = False
830       Beep
840       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
850     End If

EXITP:
860     Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

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

        ' ** Plain keys.
1050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1060      Select Case intRetVal
          Case vbKeyEscape
1070        intRetVal = 0
1080        cmdCancel_Click  ' ** Procedure: Below.
1090      End Select
1100    End If

        ' ** Alt keys.
1110    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1120      Select Case intRetVal
          Case vbKey1
1130        With Me
1140          intRetVal = 0
1150          .cmbOrd.SetFocus
1160        End With
1170      Case vbKey2
1180        With Me
1190          intRetVal = 0
1200          .CashAssets_Beg.SetFocus
1210        End With
1220      Case vbKey3
1230        With Me
1240          intRetVal = 0
1250          .NetIncome.SetFocus
1260        End With
1270      End Select
1280    End If

EXITP:
1290    KeyCode = intRetVal
1300    Exit Sub

ERRH:
1310    intRetVal = 0
1320    Select Case ERR.Number
        Case Else
1330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1340    End Select
1350    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1410    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1420    If blnOrderlyClose = False Then
1430      gblnMessage = False
1440    End If

EXITP:
1450    Exit Sub

ERRH:
1460    Select Case ERR.Number
        Case Else
1470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1480    End Select
1490    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOk_Click"

        Dim blnContinue As Boolean

1510    blnContinue = True

1520    With Me
1530      If IsNumeric(.CashAssets_Beg) = False Then
1540        blnContinue = False
1550        Beep
1560        MsgBox "Beginning Cash Assets must be a numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
1570        .CashAssets_Beg.SetFocus
1580      End If
1590    End With

1600    If blnContinue = True Then
1610      blnOrderlyClose = True
1620      DoCmd.Close acForm, THIS_NAME
1630      DoCmd.SelectObject acForm, "frmRpt_CourtReports_NS", False
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

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1720    intShiftDown = (Shift And acShiftMask) > 0
1730    intAltDown = (Shift And acAltMask) > 0
1740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1760      Select Case intRetVal
          Case vbKeyTab
1770        With Me
1780          intRetVal = 0
1790          .cmdCancel.SetFocus
1800        End With
1810      End Select
1820    End If

        ' ** Shift keys.
1830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1840      Select Case intRetVal
          Case vbKeyTab
1850        With Me
1860          intRetVal = 0
1870          .NetIncome.SetFocus
1880        End With
1890      End Select
1900    End If

EXITP:
1910    KeyCode = intRetVal
1920    Exit Sub

ERRH:
1930    intRetVal = 0
1940    Select Case ERR.Number
        Case Else
1950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1960    End Select
1970    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

2010    blnOrderlyClose = True
2020    gblnMessage = False
2030    DoCmd.Close acForm, THIS_NAME
2040    DoCmd.SelectObject acForm, "frmRpt_CourtReports_NS", False

EXITP:
2050    Exit Sub

ERRH:
2060    Select Case ERR.Number
        Case Else
2070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2080    End Select
2090    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
          Case vbKeyTab
2170        With Me
2180          intRetVal = 0
2190          .cmbOrd.SetFocus
2200        End With
2210      End Select
2220    End If

        ' ** Shift keys.
2230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyTab
2250        With Me
2260          intRetVal = 0
2270          .cmdOK.SetFocus
2280        End With
2290      End Select
2300    End If

EXITP:
2310    KeyCode = intRetVal
2320    Exit Sub

ERRH:
2330    intRetVal = 0
2340    Select Case ERR.Number
        Case Else
2350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2360    End Select
2370    Resume EXITP

End Sub

Private Sub cmbOrd_GotFocus()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_GotFocus"

2410    With Me
2420      .cmbOrd.SelLength = 0
2430      .cmbOrd.SelStart = 99
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

Private Sub cmbOrd_AfterUpdate()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_AfterUpdate"

2510    With Me
2520      If IsNull(.cmbOrd) = False Then
2530        gstrCrtRpt_Ordinal = .cmbOrd.Column(1)
2540      Else
2550        .cmbOrd = .cmbOrd.DefaultValue
2560        gstrCrtRpt_Ordinal = .cmbOrd.Column(1)
2570      End If
2580    End With

EXITP:
2590    Exit Sub

ERRH:
2600    Select Case ERR.Number
        Case Else
2610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2620    End Select
2630    Resume EXITP

End Sub

Private Sub cmbOrd_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_KeyDown"

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
2790          .opgVer.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2850        With Me
2860          intRetVal = 0
2870          .cmdCancel.SetFocus
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

Private Sub opgVer_AfterUpdate()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgVer_AfterUpdate"

3010    With Me
3020      Select Case .opgVer
          Case .opgVer_optFinal.OptionValue
3030        .opgVer_optFinal_lbl.FontBold = True
3040        .opgVer_optInterim_lbl.FontBold = False
3050        gstrCrtRpt_Version = Mid(.opgVer_optFinal.Name, (InStr(.opgVer_optFinal.Name, "Opt") + 3))
3060      Case .opgVer_optInterim.OptionValue
3070        .opgVer_optFinal_lbl.FontBold = False
3080        .opgVer_optInterim_lbl.FontBold = True
3090        gstrCrtRpt_Version = Mid(.opgVer_optInterim.Name, (InStr(.opgVer_optInterim.Name, "Opt") + 3))
3100      End Select
3110    End With

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub CashAssets_Beg_AfterUpdate()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "CashAssets_Beg_AfterUpdate"

3210    With Me
3220      If IsNull(.CashAssets_Beg) = False Then
3230        If IsNumeric(.CashAssets_Beg) = True Then
3240          gstrCrtRpt_CashAssets_Beg = CStr(.CashAssets_Beg)
3250          .CashAssets_Beg_display_lbl.Caption = Format(.CashAssets_Beg, "Currency")
3260        Else
3270          MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
3280        End If
3290      Else
3300        .CashAssets_Beg = 0&
3310        gstrCrtRpt_CashAssets_Beg = "0"
3320        .CashAssets_Beg_display_lbl.Caption = Format(.CashAssets_Beg, "Currency")
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

Private Sub CashAssets_Beg_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "CashAssets_Beg_KeyDown"

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
3490          .NetIncome.SetFocus
3500        End With
3510      End Select
3520    End If

        ' ** Shift keys.
3530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3550        With Me
3560          intRetVal = 0
3570          .opgVer.SetFocus
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

Private Sub NetIncome_AfterUpdate()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "NetIncome_AfterUpdate"

3710    With Me
3720      If IsNull(.NetIncome) = False Then
3730        If IsNumeric(.NetIncome) = True Then
3740          gstrCrtRpt_NetIncome = CStr(.NetIncome)
3750          .NetIncome_display_lbl.Caption = Format(.NetIncome, "Currency")
3760        Else
3770          MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
3780        End If
3790      Else
3800        .NetIncome = 0
3810        gstrCrtRpt_NetIncome = "0"
3820        .NetIncome_display_lbl.Caption = Format(.NetIncome, "Currency")
3830      End If
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

Private Sub NetIncome_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "NetIncome_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3920    intShiftDown = (Shift And acShiftMask) > 0
3930    intAltDown = (Shift And acAltMask) > 0
3940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3970        With Me
3980          intRetVal = 0
3990          .cmdOK.SetFocus
4000        End With
4010      End Select
4020    End If

        ' ** Shift keys.
4030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4050        With Me
4060          intRetVal = 0
4070          .CashAssets_Beg.SetFocus
4080        End With
4090      End Select
4100    End If

EXITP:
4110    KeyCode = intRetVal
4120    Exit Sub

ERRH:
4130    intRetVal = 0
4140    Select Case ERR.Number
        Case Else
4150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4160    End Select
4170    Resume EXITP

End Sub
