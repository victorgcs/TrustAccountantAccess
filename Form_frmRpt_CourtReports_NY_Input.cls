VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_CourtReports_NY_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_CourtReports_NY_Input"

'VGC 12/07/2012: CHANGES!
'FC 11/21/2011: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Final:            F {opgVer_optFinal}
' **   Interim:          I {opgVer_optInterim}
' **   OK:               O {cmdOK}
' **   Summary Ord/Ver:  1 {cmbOrd}
' **   Invested Income:  2 {NetIncome}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           .PictureData = .Form_Rpt_CourtReports_NY_Input_img

160           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

170           If gblnCrtRpt_NY_InvIncChange = True Then
180             .NetIncome = gcurCrtRpt_NY_InputNew
190           Else
200             .NetIncome = 0
210           End If

220           cmbOrd_AfterUpdate  ' ** Procedure: Below.
230           opgVer_AfterUpdate  ' ** Procedure: Below.

240         Else
250           Cancel = -1
260         End If
270       Else
280         Cancel = -1
290       End If
300     End With

310     If Cancel = -1 Then
320       Beep
330       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
340     End If

EXITP:
350     Exit Sub

ERRH:
360     Select Case ERR.Number
        Case Else
370       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
380     End Select
390     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

410     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
420     intShiftDown = (Shift And acShiftMask) > 0
430     intAltDown = (Shift And acAltMask) > 0
440     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
450     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
460       Select Case intRetVal
          Case vbKeyEscape
470         intRetVal = 0
480         cmdCancel_Click  ' ** Procedure: Below.
490       End Select
500     End If

        ' ** Alt keys.
510     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
520       Select Case intRetVal
          Case vbKey1
530         With Me
540           intRetVal = 0
550           .cmbOrd.SetFocus
560         End With
570       Case vbKey2
580         With Me
590           intRetVal = 0
600           .NetIncome.SetFocus
610         End With
620       End Select
630     End If

EXITP:
640     KeyCode = intRetVal
650     Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

810     Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

820     Select Case strCallingForm
        Case "frmRpt_CourtReports_NY"
830       DoCmd.SelectObject acForm, strCallingForm, False
840     End Select

EXITP:
850     Exit Sub

ERRH:
860     Select Case ERR.Number
        Case Else
870       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
880     End Select
890     Resume EXITP

End Sub

Private Sub cmdOK_Click()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim frm As Access.Form
        Dim blnContinue As Boolean

910     blnContinue = True

920     Set frm = Forms("frmRpt_CourtReports_NY")

930     With Me
940       If IsNull(.NetIncome) = True Then
950         blnContinue = False
960         Beep
970         MsgBox "Invested Income must be a numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
980       Else
990         If IsNumeric(.NetIncome) = False Then
1000          blnContinue = False
1010          Beep
1020          MsgBox "Invested Income must be a numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
1030        End If
1040      End If

1050      If blnContinue = True Then

1060        frm.Ordinal = .cmbOrd.Column(1)
1070        Select Case .opgVer
            Case .opgVer_optFinal.OptionValue
1080          frm.Version = "Final"
1090        Case .opgVer_optInterim.OptionValue
1100          frm.Version = "Interim"
1110        End Select
1120        frm.CashAssets_Beg = .NetIncome

1130        gstrCrtRpt_CashAssets_Beg = .NetIncome
1140        gblnCrtRpt_NY_InvIncChange = True
1150        gcurCrtRpt_NY_InputNew = .NetIncome

1160        DoCmd.Close acForm, THIS_NAME

1170      End If

1180    End With

EXITP:
1190    Set frm = Nothing
1200    Exit Sub

ERRH:
1210    Select Case ERR.Number
        Case Else
1220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1230    End Select
1240    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1320    intShiftDown = (Shift And acShiftMask) > 0
1330    intAltDown = (Shift And acAltMask) > 0
1340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyTab
1370        With Me
1380          intRetVal = 0
1390          .cmdCancel.SetFocus
1400        End With
1410      End Select
1420    End If

        ' ** Shift keys.
1430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1440      Select Case intRetVal
          Case vbKeyTab
1450        With Me
1460          intRetVal = 0
1470          .NetIncome.SetFocus
1480        End With
1490      End Select
1500    End If

EXITP:
1510    KeyCode = intRetVal
1520    Exit Sub

ERRH:
1530    intRetVal = 0
1540    Select Case ERR.Number
        Case Else
1550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1560    End Select
1570    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim frm As Access.Form

1610    Set frm = Forms("frmRpt_CourtReports_NY")

1620    With frm
1630      .Ordinal = vbNullString  ' ** But leave the Public variable.
1640      .Version = vbNullString  ' ** But leave the Public variable.
1650      .CashAssets_Beg = vbNullString ' ** But leave the Public variable.
1660    End With

1670    gblnMessage = False
1680    DoCmd.Close acForm, THIS_NAME

EXITP:
1690    Set frm = Nothing
1700    Exit Sub

ERRH:
1710    Select Case ERR.Number
        Case Else
1720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1730    End Select
1740    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1820    intShiftDown = (Shift And acShiftMask) > 0
1830    intAltDown = (Shift And acAltMask) > 0
1840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1860      Select Case intRetVal
          Case vbKeyTab
1870        With Me
1880          intRetVal = 0
1890          .cmbOrd.SetFocus
1900        End With
1910      End Select
1920    End If

        ' ** Shift keys.
1930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1940      Select Case intRetVal
          Case vbKeyTab
1950        With Me
1960          intRetVal = 0
1970          .cmdOK.SetFocus
1980        End With
1990      End Select
2000    End If

EXITP:
2010    KeyCode = intRetVal
2020    Exit Sub

ERRH:
2030    intRetVal = 0
2040    Select Case ERR.Number
        Case Else
2050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2060    End Select
2070    Resume EXITP

End Sub

Private Sub cmbOrd_GotFocus()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_GotFocus"

2110    With Me
2120      .cmbOrd.SelLength = 0
2130      .cmbOrd.SelStart = 99
2140    End With

EXITP:
2150    Exit Sub

ERRH:
2160    Select Case ERR.Number
        Case Else
2170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2180    End Select
2190    Resume EXITP

End Sub

Private Sub cmbOrd_AfterUpdate()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_AfterUpdate"

2210    With Me
2220      If IsNull(.cmbOrd) = False Then
2230        gstrCrtRpt_Ordinal = .cmbOrd.Column(1)
2240      Else
2250        .cmbOrd = .cmbOrd.DefaultValue
2260        gstrCrtRpt_Ordinal = .cmbOrd.Column(1)
2270      End If
2280    End With

EXITP:
2290    Exit Sub

ERRH:
2300    Select Case ERR.Number
        Case Else
2310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2320    End Select
2330    Resume EXITP

End Sub

Private Sub cmbOrd_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbOrd_KeyDown"

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
2490          .opgVer.SetFocus
2500        End With
2510      End Select
2520    End If

        ' ** Shift keys.
2530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2550        With Me
2560          intRetVal = 0
2570          .cmdCancel.SetFocus
2580        End With
2590      End Select
2600    End If

EXITP:
2610    KeyCode = intRetVal
2620    Exit Sub

ERRH:
2630    intRetVal = 0
2640    Select Case ERR.Number
        Case Else
2650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2660    End Select
2670    Resume EXITP

End Sub

Private Sub opgVer_AfterUpdate()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgVer_AfterUpdate"

2710    With Me
2720      Select Case .opgVer
          Case .opgVer_optFinal.OptionValue
2730        .opgVer_optFinal_lbl.FontBold = True
2740        .opgVer_optInterim_lbl.FontBold = False
2750        gstrCrtRpt_Version = Mid(.opgVer_optFinal.Name, (InStr(.opgVer_optFinal.Name, "Opt") + 3))
2760      Case .opgVer_optInterim.OptionValue
2770        .opgVer_optFinal_lbl.FontBold = False
2780        .opgVer_optInterim_lbl.FontBold = True
2790        gstrCrtRpt_Version = Mid(.opgVer_optInterim.Name, (InStr(.opgVer_optInterim.Name, "Opt") + 3))
2800      End Select
2810    End With

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub opgVer_optFinal_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgVer_optFinal_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2920    intShiftDown = (Shift And acShiftMask) > 0
2930    intAltDown = (Shift And acAltMask) > 0
2940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2970        With Me
2980          intRetVal = 0
2990          .NetIncome.SetFocus
3000        End With
3010      End Select
3020    End If

        ' ** Shift keys.
3030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3050        With Me
3060          intRetVal = 0
3070          .cmbOrd.SetFocus
3080        End With
3090      End Select
3100    End If

EXITP:
3110    KeyCode = intRetVal
3120    Exit Sub

ERRH:
3130    intRetVal = 0
3140    Select Case ERR.Number
        Case Else
3150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3160    End Select
3170    Resume EXITP

End Sub

Private Sub opgVer_optInterim_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgVer_optInterim_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3220    intShiftDown = (Shift And acShiftMask) > 0
3230    intAltDown = (Shift And acAltMask) > 0
3240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3270        With Me
3280          intRetVal = 0
3290          .NetIncome.SetFocus
3300        End With
3310      End Select
3320    End If

        ' ** Shift keys.
3330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3350        With Me
3360          intRetVal = 0
3370          .cmbOrd.SetFocus
3380        End With
3390      End Select
3400    End If

EXITP:
3410    KeyCode = intRetVal
3420    Exit Sub

ERRH:
3430    intRetVal = 0
3440    Select Case ERR.Number
        Case Else
3450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3460    End Select
3470    Resume EXITP

End Sub

Private Sub NetIncome_AfterUpdate()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "NetIncome_AfterUpdate"

3510    With Me
3520      If IsNull(.NetIncome) = False Then
3530        If IsNumeric(.NetIncome) = True Then
3540          gstrCrtRpt_CashAssets_Beg = CStr(.NetIncome)
3550          gcurCrtRpt_NY_InputAmtForm = CCur(gstrCrtRpt_CashAssets_Beg)
3560          .NetIncome_display_lbl.Caption = Format(.NetIncome, "Currency")
3570        Else
3580          Beep
3590          MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
3600          .NetIncome.SetFocus
3610        End If
3620      Else
3630        .NetIncome = 0
3640        gstrCrtRpt_CashAssets_Beg = "0"
3650        .NetIncome_display_lbl.Caption = Format(.NetIncome, "Currency")
3660      End If
3670    End With

EXITP:
3680    Exit Sub

ERRH:
3690    Select Case ERR.Number
        Case Else
3700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3710    End Select
3720    Resume EXITP

End Sub

Private Sub NetIncome_KeyDown(KeyCode As Integer, Shift As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "NetIncome_KeyDown"

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
3890          .cmdOK.SetFocus
3900        End With
3910      End Select
3920    End If

        ' ** Shift keys.
3930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3950        With Me
3960          intRetVal = 0
3970          .opgVer.SetFocus
3980        End With
3990      End Select
4000    End If

EXITP:
4010    KeyCode = intRetVal
4020    Exit Sub

ERRH:
4030    intRetVal = 0
4040    Select Case ERR.Number
        Case Else
4050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4060    End Select
4070    Resume EXITP

End Sub
