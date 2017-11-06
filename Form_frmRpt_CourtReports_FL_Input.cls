VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_CourtReports_FL_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_CourtReports_FL_Input"

'VGC 12/07/2012: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Final:            F {opgVer_optFinal}
' **   Interim:          I {opgVer_optInterim}
' **   OK:               O {cmdOK}
' **   Summary Ord/Ver:  1 {cmbOrd}
' **   Beg Cash:         2 {CashAssets_Beg}

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

160           .PictureData = .Form_Rpt_CourtReports_FL_Input_img

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

650           .cmdOK.SetFocus

660         Else
670           Cancel = -1
680         End If
690       Else
700         Cancel = -1
710       End If
720     End With

730     If Cancel = -1 Then
740       blnOrderlyClose = True
750       gblnMessage = False
760       Beep
770       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
780     End If

EXITP:
790     Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

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
          Case vbKeyEscape
970         intRetVal = 0
980         cmdCancel_Click  ' ** Procedure: Below.
990       End Select
1000    End If

        ' ** Alt keys.
1010    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1020      Select Case intRetVal
          Case vbKey1
1030        With Me
1040          intRetVal = 0
1050          .cmbOrd.SetFocus
1060        End With
1070      Case vbKey2
1080        With Me
1090          intRetVal = 0
1100          .CashAssets_Beg.SetFocus
1110        End With
1120      End Select
1130    End If

EXITP:
1140    KeyCode = intRetVal
1150    Exit Sub

ERRH:
1160    intRetVal = 0
1170    Select Case ERR.Number
        Case Else
1180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1190    End Select
1200    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1310    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1320    If blnOrderlyClose = False Then
1330      gblnMessage = False
1340    End If

EXITP:
1350    Exit Sub

ERRH:
1360    Select Case ERR.Number
        Case Else
1370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1380    End Select
1390    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOk_Click"

        Dim blnContinue As Boolean

1410    blnContinue = True

1420    With Forms("frmRpt_CourtReports_FL")
1430      .Ordinal = Me.cmbOrd.Column(1)
1440      Select Case Me.opgVer
          Case Me.opgVer_optFinal.OptionValue
1450        .Version = "Final"
1460      Case Me.opgVer_optInterim.OptionValue
1470        .Version = "Interim"
1480      End Select
1490    End With
1500    With Me
1510      If IsNumeric(.CashAssets_Beg) = False Then
1520        blnContinue = False
1530        Beep
1540        MsgBox "Beginning Cash Assets must be a numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
1550        .CashAssets_Beg.SetFocus
1560      End If
1570    End With

1580    If blnContinue = True Then
1590      blnOrderlyClose = True
1600      DoCmd.Close acForm, THIS_NAME
1610      DoCmd.SelectObject acForm, "frmRpt_CourtReports_FL", False
1620    End If

EXITP:
1630    Exit Sub

ERRH:
1640    Select Case ERR.Number
        Case Else
1650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1660    End Select
1670    Resume EXITP

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
1870          .CashAssets_Beg.SetFocus
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
2030    With Forms("frmRpt_CourtReports_FL")
2040      .Ordinal = vbNullString  ' ** But leave the Public variable.
2050      .Version = vbNullString  ' ** But leave the Public variable.
2060    End With
2070    DoCmd.Close acForm, THIS_NAME
2080    DoCmd.SelectObject acForm, "frmRpt_CourtReports_FL", False

EXITP:
2090    Exit Sub

ERRH:
2100    Select Case ERR.Number
        Case Else
2110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2120    End Select
2130    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
2290          .cmbOrd.SetFocus
2300        End With
2310      End Select
2320    End If

        ' ** Shift keys.
2330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyTab
2350        With Me
2360          intRetVal = 0
2370          .cmdOK.SetFocus
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

Private Sub cmbOrd_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_GotFocus"

2510    With Me
2520      .cmbOrd.SelLength = 0
2530      .cmbOrd.SelStart = 99
2540    End With

EXITP:
2550    Exit Sub

ERRH:
2560    Select Case ERR.Number
        Case Else
2570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2580    End Select
2590    Resume EXITP

End Sub

Private Sub cmbOrd_AfterUpdate()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_AfterUpdate"

2610    With Me
2620      If IsNull(.cmbOrd) = False Then
2630        gstrCrtRpt_Ordinal = .cmbOrd.Column(1)
2640      Else
2650        .cmbOrd = .cmbOrd.DefaultValue
2660        gstrCrtRpt_Ordinal = .cmbOrd.Column(1)
2670      End If
2680    End With

EXITP:
2690    Exit Sub

ERRH:
2700    Select Case ERR.Number
        Case Else
2710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2720    End Select
2730    Resume EXITP

End Sub

Private Sub cmbOrd_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2820    intShiftDown = (Shift And acShiftMask) > 0
2830    intAltDown = (Shift And acAltMask) > 0
2840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2870        With Me
2880          intRetVal = 0
2890          .opgVer.SetFocus
2900        End With
2910      End Select
2920    End If

        ' ** Shift keys.
2930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2950        With Me
2960          intRetVal = 0
2970          .cmdCancel.SetFocus
2980        End With
2990      End Select
3000    End If

EXITP:
3010    KeyCode = intRetVal
3020    Exit Sub

ERRH:
3030    intRetVal = 0
3040    Select Case ERR.Number
        Case Else
3050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3060    End Select
3070    Resume EXITP

End Sub

Private Sub opgVer_AfterUpdate()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgVer_AfterUpdate"

3110    With Me
3120      Select Case .opgVer
          Case .opgVer_optFinal.OptionValue
3130        .opgVer_optFinal_lbl.FontBold = True
3140        .opgVer_optInterim_lbl.FontBold = False
3150        gstrCrtRpt_Version = Mid(.opgVer_optFinal.Name, (InStr(.opgVer_optFinal.Name, "Opt") + 3))
3160      Case .opgVer_optInterim.OptionValue
3170        .opgVer_optFinal_lbl.FontBold = False
3180        .opgVer_optInterim_lbl.FontBold = True
3190        gstrCrtRpt_Version = Mid(.opgVer_optInterim.Name, (InStr(.opgVer_optInterim.Name, "Opt") + 3))
3200      End Select
3210    End With

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub CashAssets_Beg_AfterUpdate()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "CashAssets_Beg_AfterUpdate"

3310    With Me
3320      If IsNull(.CashAssets_Beg) = False Then
3330        If IsNumeric(.CashAssets_Beg) = True Then
3340          gstrCrtRpt_CashAssets_Beg = CStr(.CashAssets_Beg)
3350          .CashAssets_Beg_display_lbl.Caption = Format(.CashAssets_Beg, "Currency")
3360        Else
3370          MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
3380        End If
3390      Else
3400        .CashAssets_Beg = 0&
3410        gstrCrtRpt_CashAssets_Beg = "0"
3420        .CashAssets_Beg_display_lbl.Caption = Format(.CashAssets_Beg, "Currency")
3430      End If
3440    End With

EXITP:
3450    Exit Sub

ERRH:
3460    Select Case ERR.Number
        Case Else
3470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3480    End Select
3490    Resume EXITP

End Sub

Private Sub CashAssets_Beg_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "CashAssets_Beg_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3520    intShiftDown = (Shift And acShiftMask) > 0
3530    intAltDown = (Shift And acAltMask) > 0
3540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3570        With Me
3580          intRetVal = 0
3590          .cmdOK.SetFocus
3600        End With
3610      End Select
3620    End If

        ' ** Shift keys.
3630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3650        With Me
3660          intRetVal = 0
3670          .opgVer.SetFocus
3680        End With
3690      End Select
3700    End If

EXITP:
3710    KeyCode = intRetVal
3720    Exit Sub

ERRH:
3730    intRetVal = 0
3740    Select Case ERR.Number
        Case Else
3750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3760    End Select
3770    Resume EXITP

End Sub
