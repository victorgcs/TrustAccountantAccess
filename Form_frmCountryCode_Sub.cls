VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCountryCode_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCountryCode_Sub"

'VGC 04/17/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Active:           A {country_active}
' **   Close:            C {cmdClose on frmCountryCode}
' **   Misc Exclude:     D {opgMisc_optExclude on frmCountryCode}
' **   Active Exclude:   E {opgActive_optExclude on frmCountryCode}
' **   Active Include:   I {opgActive_optInclude on frmCountryCode}
' **   Misc Only:        L {opgMisc_optOnly on frmCountryCode}
' **   Name:             M {country_name}
' **   Currencies:       N {cmdCurrencies on frmCountryCode}
' **   Active Only:      O {opgActive_optOnly on frmCountryCode}
' **   Print Report:     P {cmdPrintReport on frmCountryCode}
' **   Misc Include:     U {opgMisc_optInclude on frmCountryCode}
' **   Exit:             X {cmdClose on frmCountryCode}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   2-Ltr Exclude:    E {opgTwoLetter_optExclude on frmCountryCode}
' **   2-Ltr Include:    I {opgTwoLetter_optInclude on frmCountryCode}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[country_name]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngTpp As Long
Private lngForm_Width As Long, lngName_Left As Long, lngISO_Left As Long, lngDemonym_Left As Long
Private lngActive_Left As Long, lngNA_Left As Long, lngActiveLbl_Left As Long, lngNALbl_Left As Long
Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmCountryCode" Then

160         lngSortLbl_Width = .Sort_lbl.Width

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

200         .form_width_line.Left = 0&  ' ** Just get it out of the way.
210         lngForm_Width = .Width  ' ** Includes Misc, doesn't inlucde Code2.
220         lngName_Left = .country_name.Left
230         lngISO_Left = .country_iso.Left
240         lngActive_Left = .country_active.Left
250         lngActiveLbl_Left = .country_active_lbl.Left
260         lngNA_Left = .country_na.Left
270         lngNALbl_Left = .country_na_lbl.Left
280         lngDemonym_Left = .country_demonym.Left

290       Else
300         Cancel = -1
310       End If
320     End With

330     If Cancel = -1 Then
340       Beep
350       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
360     End If

EXITP:
370     Exit Sub

ERRH:
380     Select Case ERR.Number
        Case Else
390       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
400     End Select
410     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me
520       .country_code3.ForeColor = CLR_DISABLED_FG
530       .country_code3.BackColor = CLR_DISABLED_BG
540       .country_code2.ForeColor = CLR_DISABLED_FG
550       .country_code2.BackColor = CLR_DISABLED_BG
560       .country_name.ForeColor = CLR_DISABLED_FG
570       .country_name.BackColor = CLR_DISABLED_BG
580       .country_iso.ForeColor = CLR_DISABLED_FG
590       .country_iso.BackColor = CLR_DISABLED_BG
600       .country_na_box.BackColor = CLR_DISABLED_BG
610     End With

620     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
630     Exit Sub

ERRH:
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub Form_Current()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim strDocName As String
710     With Me
720       .Parent.currency_cnt = .currency_cnt
730       strDocName = "frmCountry_Currency"
740       If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
750         Forms(strDocName).cmbCountries = .country_id
760       End If
770     End With

EXITP:
780     Exit Sub

ERRH:
790     Select Case ERR.Number
        Case Else
800       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
810     End Select
820     Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Active:           A {country_active}
        ' **   Name:             M {country_name}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality
        ' **   2-Ltr Exclude:    E {opgTwoLetter_optExclude on frmCountryCode}
        ' **   2-Ltr Include:    I {opgTwoLetter_optInclude on frmCountryCode}

        ' ** Plain keys.
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyUp
970         intRetVal = 0
980         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
990       Case vbKeyDown
1000        intRetVal = 0
1010        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1020      Case vbKeyEscape
1030        With Me
1040          intRetVal = 0
1050          .Parent.cmdClose_Click  ' ** Form Procedure: frmCountryCode.
1060        End With
1070      End Select
1080    End If

        ' ** Alt keys.
1090    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1100      Select Case intRetVal
          Case vbKeyA
1110        With Me
1120          intRetVal = 0
1130          .country_active.SetFocus
1140          .country_active = (Not .country_active)
1150          .cmdSave_Click  ' ** Procedure: Below.
1160        End With
1170      Case vbKeyM
1180        With Me
1190          intRetVal = 0
1200          .country_name.SetFocus
1210        End With
1220      Case vbKeyX
1230        With Me
1240          intRetVal = 0
1250          .Parent.cmdClose_Click  ' ** Form Procedure: frmCountryCode.
1260        End With
1270      End Select
1280    End If

        ' ** Ctrl keys.
1290    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1300      Select Case intRetVal
          Case vbKeyS
1310        With Me
1320          intRetVal = 0
1330          cmdSave_Click  ' ** Procedure: Below.
1340        End With
1350      Case vbKeyTab, vbKeyReturn
1360        With Me
1370          intRetVal = 0
1380          DoCmd.SelectObject acForm, .Parent.Name, False
1390          .Parent.opgActive.SetFocus
1400        End With
1410      Case vbKeyUp, vbKeyPageUp
1420        intRetVal = 0
1430        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1440      Case vbKeyDown, vbKeyPageDown
1450        intRetVal = 0
1460        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1470      End Select
1480    End If

        ' ** Ctrl-Shift keys.
1490    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1500      Select Case intRetVal
          Case vbKeyE
1510        With Me
1520          intRetVal = 0
1530          DoCmd.SelectObject acForm, .Parent.Name, False
1540          .Parent.opgTwoLetter = .Parent.opgTwoLetter_optExclude.OptionValue
1550          .Parent.opgTwoLetter.SetFocus
1560          .Parent.opgTwoLetter_AfterUpdate  ' ** Form Procedure: frmCountryCode.
1570        End With
1580      Case vbKeyI
1590        With Me
1600          intRetVal = 0
1610          DoCmd.SelectObject acForm, .Parent.Name, False
1620          .Parent.opgTwoLetter = .Parent.opgTwoLetter_optInclude.OptionValue
1630          .Parent.opgTwoLetter.SetFocus
1640          .Parent.opgTwoLetter_AfterUpdate  ' ** Form Procedure: frmCountryCode.
1650        End With
1660      Case vbKeyTab, vbKeyReturn
1670        With Me
1680          intRetVal = 0
1690          DoCmd.SelectObject acForm, .Parent.Name, False
1700          .Parent.cmdClose.SetFocus
1710        End With
1720      End Select
1730    End If

EXITP:
1740    KeyCode = intRetVal
1750    Exit Sub

ERRH:
1760    intRetVal = 0
1770    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1780    Case Else
1790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1800    End Select
1810    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

1910    With Me
1920      .country_datemodified = Now()
1930      .countrystage_datemodified = Now()
1940      .country_username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1950      .countrystage_username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1960    End With

EXITP:
1970    Exit Sub

ERRH:
1980    Select Case ERR.Number
        Case Else
1990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2000    End Select
2010    Resume EXITP

End Sub

Public Sub cmdSave_Click()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2110    With Me

2120      DoCmd.RunCommand acCmdSaveRecord
2130      DoEvents

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

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
2270        With Me
2280          intRetVal = 0
2290          lngRecsCur = RecCnt  ' ** Function: Below.
2300          If .CurrentRecord < lngRecsCur Then
2310            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2320            .FocusHolder.SetFocus
2330          Else
2340            DoCmd.SelectObject acForm, .Parent.Name, False
2350            .Parent.opgActive.SetFocus
2360          End If
2370        End With
2380      End Select
2390    End If

        ' ** Shift keys.
2400    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2410      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2420        With Me
2430          intRetVal = 0
2440          If .CurrentRecord > 1 Then
2450            MoveRec acCmdRecordsGoToPrevious
2460            .FocusHolder.SetFocus
2470          Else
2480            DoCmd.SelectObject acForm, .Parent.Name, False
2490            .Parent.cmdClose.SetFocus
2500          End If
2510        End With
2520      End Select
2530    End If

EXITP:
2540    KeyCode = intRetVal
2550    Exit Sub

ERRH:
2560    intRetVal = 0
2570    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2580    Case Else
2590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2600    End Select
2610    Resume EXITP

End Sub

Private Sub country_code3_lbl_DblClick(Cancel As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "country_code3_lbl_DblClick"

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

Private Sub country_code3_GotFocus()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "country_code3_GotFocus"

2810    With Me
2820      .country_code3.SelLength = 0
2830      .country_code3.SelStart = 9
2840    End With

EXITP:
2850    Exit Sub

ERRH:
2860    Select Case ERR.Number
        Case Else
2870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2880    End Select
2890    Resume EXITP

End Sub

Private Sub country_code3_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "country_code3_KeyDown"

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
2990          If .country_code2.Visible = True And .country_code2.Enabled = True Then
3000            .country_code2.SetFocus
3010          Else
3020            .country_name.SetFocus
3030          End If
3040        End With
3050      End Select
3060    End If

        ' ** Shift keys.
3070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3080      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3090        With Me
3100          intRetVal = 0
3110          If .CurrentRecord > 1 Then
3120            MoveRec acCmdRecordsGoToPrevious
3130            If .country_na.Visible = True And .country_na.Enabled = True Then
3140              .country_na.SetFocus
3150            Else
3160              .country_active.SetFocus
3170            End If
3180          Else
3190            DoCmd.SelectObject acForm, .Parent.Name, False
3200            .Parent.cmdClose.SetFocus
3210          End If
3220        End With
3230      End Select
3240    End If

EXITP:
3250    KeyCode = intRetVal
3260    Exit Sub

ERRH:
3270    intRetVal = 0
3280    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3290    Case Else
3300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3310    End Select
3320    Resume EXITP

End Sub

Private Sub country_code2_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "country_code2_lbl_DblClick"

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

Private Sub country_code2_GotFocus()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "country_code2_GotFocus"

3510    With Me
3520      .country_code2.SelLength = 0
3530      .country_code2.SelStart = 9
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

Private Sub country_name_lbl_DblClick(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "country_name_lbl_DblClick"

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

Private Sub country_name_GotFocus()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "country_name_GotFocus"

3710    With Me
3720      .country_name.SelLength = 0
3730      .country_name.SelStart = 99
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

Private Sub country_iso_lbl_DblClick(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "country_iso_lbl_DblClick"

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

Private Sub country_iso_GotFocus()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "country_iso_GotFocus"

3910    With Me
3920      .country_iso.SelLength = 0
3930      .country_iso.SelStart = 9
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

Private Sub country_active_lbl_DblClick(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "country_active_lbl_DblClick"

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

Private Sub country_active_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "country_active_KeyDown"

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
4190          If .country_na.Visible = True And .country_na.Enabled = True Then
4200            .country_na.SetFocus
4210          Else
4220            lngRecsCur = RecCnt  ' ** Function: Below.
4230            If .CurrentRecord < lngRecsCur Then
4240              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4250              .country_active.SetFocus
4260            Else
4270              DoCmd.SelectObject acForm, .Parent.Name, False
4280              .Parent.opgActive.SetFocus
4290            End If
4300          End If
4310        End With
4320      Case vbKeyUp
4330        With Me
4340          intRetVal = 0
4350          If .CurrentRecord > 1 Then
4360            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4370            .country_active.SetFocus
4380          Else
4390            DoCmd.SelectObject acForm, .Parent.Name, False
4400            .Parent.cmdClose.SetFocus
4410          End If
4420        End With
4430      Case vbKeyDown
4440        With Me
4450          intRetVal = 0
4460          lngRecsCur = RecCnt  ' ** Function: Below.
4470          If .CurrentRecord < lngRecsCur Then
4480            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4490            .country_active.SetFocus
4500          Else
4510            DoCmd.SelectObject acForm, .Parent.Name, False
4520            .Parent.opgActive.SetFocus
4530          End If

4540        End With
4550      End Select
4560    End If

        ' ** Shift keys.
4570    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4580      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4590        With Me
4600          intRetVal = 0
4610          .country_iso.SetFocus
4620        End With
4630      End Select
4640    End If

EXITP:
4650    KeyCode = intRetVal
4660    Exit Sub

ERRH:
4670    intRetVal = 0
4680    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4690    Case Else
4700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4710    End Select
4720    Resume EXITP

End Sub

Private Sub country_demonym_lbl_DblClick(Cancel As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "country_demonym_lbl_DblClick"

4810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4820    Exit Sub

ERRH:
4830    Select Case ERR.Number
        Case Else
4840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4850    End Select
4860    Resume EXITP

End Sub

Private Sub country_na_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "country_na_lbl_DblClick"

4910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub country_na_KeyDown(KeyCode As Integer, Shift As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "country_na_KeyDown"

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
5090          lngRecsCur = RecCnt  ' ** Function: Below.
5100          If .CurrentRecord < lngRecsCur Then
5110            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5120            .country_active.SetFocus
5130          Else
5140            DoCmd.SelectObject acForm, .Parent.Name, False
5150            .Parent.opgActive.SetFocus
5160          End If
5170        End With
5180      End Select
5190    End If

        ' ** Shift keys.
5200    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5210      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5220        With Me
5230          intRetVal = 0
5240          .country_iso.SetFocus
5250        End With
5260      End Select
5270    End If

EXITP:
5280    KeyCode = intRetVal
5290    Exit Sub

ERRH:
5300    intRetVal = 0
5310    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5320    Case Else
5330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5340    End Select
5350    Resume EXITP

End Sub

Public Sub ShowCode2(blnShow As Boolean)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowCode2"

        Dim lngAddedWidth As Long, lngNewFrmWidth As Long
        Dim intPos01 As Integer
        Dim strTmp01 As String, lngTmp02 As Long

5410    With Me
5420      .FocusHolder.SetFocus
5430      Select Case blnShow
          Case True
5440        lngAddedWidth = (.country_code2_lbl.Width + (4& * lngTpp))
5450        lngNewFrmWidth = (lngForm_Width + lngAddedWidth)
5460        Select Case .Parent.opgMisc
            Case .Parent.opgMisc_optInclude.OptionValue, .Parent.opgMisc_optOnly.OptionValue
5470          .Width = lngNewFrmWidth
5480          .country_na.Left = (lngNA_Left + lngAddedWidth)
5490          .country_na_lbl.Left = (lngNALbl_Left + lngAddedWidth)
5500          .country_na_lbl_line.Left = (lngNALbl_Left + lngAddedWidth)
5510          .country_na_box.Left = (lngNALbl_Left + lngAddedWidth)
5520        Case .Parent.opgMisc_optExclude.OptionValue
5530          lngNewFrmWidth = (lngNewFrmWidth - (.country_na_lbl.Width + (4& * lngTpp)))
5540          .Width = lngNewFrmWidth
5550        End Select
5560        .country_demonym.Left = (lngDemonym_Left + lngAddedWidth)
5570        .country_demonym_lbl.Left = (lngDemonym_Left + lngAddedWidth)
5580        .country_demonym_lbl_line.Left = (lngDemonym_Left + lngAddedWidth)
5590        .country_active.Left = (lngActive_Left + lngAddedWidth)
5600        .country_active_lbl.Left = (lngActiveLbl_Left + lngAddedWidth)
5610        .country_active_lbl_line.Left = (lngActiveLbl_Left + lngAddedWidth)
5620        .country_active_box1.Left = ((lngActiveLbl_Left + lngAddedWidth) - lngTpp)
5630        .country_active_box2.Left = (.country_active_box1.Left - lngTpp)
5640        .country_name.Left = (lngName_Left + lngAddedWidth)
5650        .country_name_lbl.Left = (lngName_Left + lngAddedWidth)
5660        .country_name_lbl_line.Left = (lngName_Left + lngAddedWidth)
5670        .country_iso.Left = (lngISO_Left + lngAddedWidth)
5680        .country_iso_lbl.Left = (lngISO_Left + lngAddedWidth)
5690        .country_iso_lbl_line.Left = (lngISO_Left + lngAddedWidth)
5700        .country_code2.Visible = True
5710        .country_code2_lbl.Visible = True
5720        .country_code2_lbl_line.Visible = True
5730      Case False
5740        .country_code2.Visible = False
5750        .country_code2_lbl.Visible = False
5760        .country_code2_lbl_line.Visible = False
5770        .country_iso.Left = lngISO_Left
5780        .country_iso_lbl.Left = lngISO_Left
5790        .country_iso_lbl_line.Left = lngISO_Left
5800        .country_name.Left = lngName_Left
5810        .country_name_lbl.Left = lngName_Left
5820        .country_name_lbl_line.Left = lngName_Left
5830        .country_active.Left = lngActive_Left
5840        .country_active_lbl.Left = lngActiveLbl_Left
5850        .country_active_lbl_line.Left = lngActiveLbl_Left
5860        .country_active_box1.Left = (lngActiveLbl_Left - lngTpp)
5870        .country_active_box2.Left = (.country_active_box1.Left - lngTpp)
5880        .country_demonym.Left = lngDemonym_Left
5890        .country_demonym_lbl.Left = lngDemonym_Left
5900        .country_demonym_lbl_line.Left = lngDemonym_Left
5910        Select Case .Parent.opgMisc
            Case .Parent.opgMisc_optInclude.OptionValue, .Parent.opgMisc_optOnly.OptionValue
5920          .country_na.Left = lngNA_Left
5930          .country_na_lbl.Left = lngNALbl_Left
5940          .country_na_lbl_line.Left = lngNALbl_Left
5950          .country_na_box.Left = lngNALbl_Left
5960          .Width = lngForm_Width
5970        Case .Parent.opgMisc_optExclude.OptionValue
5980          lngTmp02 = (.country_na_lbl.Width + (4& * lngTpp))
5990          .Width = (lngForm_Width - lngTmp02)
6000        End Select
6010        If Left(strSortNow, 15) = "[country_code2]" Then
6020          SortNow "Form_Load"  ' ** Procedure: Below.
6030        End If
6040      End Select
6050      strTmp01 = strSortNow
6060      lngTmp02 = .Controls(strSortLbl).Left - .Controls(strSortLine).Left
6070      intPos01 = InStr(strTmp01, " ")
6080      If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, intPos01))
6090      If Left(strTmp01, 1) = "[" Then strTmp01 = Mid(strTmp01, 2)
6100      If Right(strTmp01, 1) = "]" Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
6110      .Controls(strSortLine).Left = .Controls(strTmp01 & "_lbl").Left
6120      .Controls(strSortLbl).Left = .Controls(strSortLine).Left + lngTmp02
6130    End With

EXITP:
6140    Exit Sub

ERRH:
6150    Select Case ERR.Number
        Case Else
6160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6170    End Select
6180    Resume EXITP

End Sub

Public Sub ShowMisc(blnShow As Boolean)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowMisc"

        Dim lngAddedWidth As Long, lngNewFrmWidth As Long
        Dim lngTmp02 As Long

6210    With Me
6220      .FocusHolder.SetFocus
6230      lngAddedWidth = (.country_code2_lbl.Width + (4& * lngTpp))
6240      Select Case blnShow
          Case True
6250        Select Case .Parent.opgTwoLetter
            Case .Parent.opgTwoLetter_optInclude.OptionValue
6260          lngNewFrmWidth = lngForm_Width + lngAddedWidth
6270          If .Width < lngNewFrmWidth Then
6280            .Width = lngNewFrmWidth
6290          End If
6300          .country_na.Left = (lngNA_Left + lngAddedWidth)
6310          .country_na_lbl.Left = (lngNALbl_Left + lngAddedWidth)
6320          .country_na_lbl_line.Left = .country_na_lbl.Left
6330          .country_na_box.Left = .country_na_lbl.Left
6340        Case .Parent.opgTwoLetter_optExclude.OptionValue
6350          .country_na.Left = lngNA_Left
6360          .country_na_lbl.Left = lngNALbl_Left
6370          .country_na_lbl_line.Left = .country_na_lbl.Left
6380          .country_na_box.Left = .country_na_lbl.Left
6390        End Select
6400        .country_na.Visible = True
6410        .country_na_lbl.Visible = True
6420        .country_na_lbl_line.Visible = True
6430        .country_na_box.Visible = True
6440      Case False
6450        lngTmp02 = (.country_na_lbl.Width + (4& * lngTpp))
6460        .country_na.Visible = False
6470        .country_na_lbl.Visible = False
6480        .country_na_lbl_line.Visible = False
6490        .country_na_box.Visible = False
6500        .country_na.Left = 0&
6510        .country_na_lbl.Left = 0&
6520        .country_na_lbl_line.Left = 0&
6530        .country_na_box.Left = 0&
6540        Select Case .Parent.opgTwoLetter
            Case .Parent.opgTwoLetter_optInclude.OptionValue
6550          lngNewFrmWidth = ((lngForm_Width + lngAddedWidth) - lngTmp02)
6560        Case .Parent.opgTwoLetter_optExclude.OptionValue
6570          lngNewFrmWidth = (lngForm_Width - lngTmp02)
6580        End Select
6590        If .Width > lngNewFrmWidth Then
6600          .Width = lngNewFrmWidth
6610        End If
6620      End Select
6630    End With

EXITP:
6640    Exit Sub

ERRH:
6650    Select Case ERR.Number
        Case Else
6660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6670    End Select
6680    Resume EXITP

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
7140            .FindFirst "[country_id] = '" & varID & "'"
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

        Const strStdAsc As String = ", [country_name]"

7410    With Me
7420      .Controls(strSortLbl).Visible = False
7430      .Controls(strSortLine).Visible = False
7440      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
7450      If strProc = "Form_Load" Then
7460        strCalled = "country_name"
7470        strSortNow = strSortOrig
7480        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7490        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7500        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7510        .Controls(strSortLbl).Left = lngSortLbl_Left
7520        .Controls(strSortLbl).Caption = strArwUp
7530        .Controls(strSortLbl).ForeColor = CLR_DKBLU
7540        .Controls(strSortLine).Left = lngSortLine_Left
7550        .Controls(strSortLine).Width = lngSortLine_Width
7560      Else
7570        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
7580        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7590        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7600        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7610        .Controls(strSortLbl).Left = lngSortLbl_Left
7620        .Controls(strSortLine).Left = lngSortLine_Left
7630        .Controls(strSortLine).Width = lngSortLine_Width
7640        If strCalled = "country_name" Then
7650          strSortAsc = strSortOrig
7660          If strSortNow = strSortAsc Then
7670            strSortNow = "[country_name] DESC"
7680            .Controls(strSortLbl).Caption = strArwDn
7690            .Controls(strSortLbl).ForeColor = CLR_DKRED
7700          Else
7710            strSortNow = strSortAsc
7720            .Controls(strSortLbl).Caption = strArwUp
7730            .Controls(strSortLbl).ForeColor = CLR_DKBLU
7740          End If
7750        Else
7760          .Controls(strSortLbl).Caption = strArwUp
7770          .Controls(strSortLbl).ForeColor = CLR_DKBLU
7780          Select Case strCalled
              Case "country_code3"
7790            strSortAsc = "[country_code3]"
7800          Case "country_code2"
7810            strSortAsc = "[country_code2]"
7820            lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
7830            .Controls(strSortLbl).Left = lngSortLbl_Left
7840          Case "country_iso"
7850            strSortAsc = "[country_iso]"
7860          Case "country_active"
7870            strSortAsc = "[country_active]" & strStdAsc
7880            lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
7890            .Controls(strSortLbl).Left = lngSortLbl_Left
7900          Case "country_demonym"
7910            strSortAsc = "[country_demonym]" & strStdAsc
7920          Case "country_na"
7930            strSortAsc = "[country_na]" & strStdAsc
7940            lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
7950            .Controls(strSortLbl).Left = lngSortLbl_Left
7960          End Select
7970          If strSortNow = strSortAsc Then
7980            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
7990            Select Case intCnt
                Case 1
8000              strTmp01 = strSortAsc & " DESC"
8010            Case 2
8020              intPos01 = InStr(strSortAsc, ",")
8030              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8040              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
8050            Case 3
8060              intPos01 = InStr(strSortAsc, ",")
8070              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8080              strTmp02 = Mid(strSortAsc, intPos01)
8090              intPos01 = InStr(2, strTmp02, ",")
8100              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
8110              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
8120            Case 4
8130              intPos01 = InStr(strSortAsc, ",")
8140              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8150              strTmp02 = Mid(strSortAsc, intPos01)
8160              intPos01 = InStr(2, strTmp02, ",")
8170              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
8180              strTmp02 = Mid(strTmp02, intPos01)
8190              intPos01 = InStr(2, strTmp02, ",")
8200              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
8210              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
8220            End Select
8230            strSortNow = strTmp01
8240            .Controls(strSortLbl).Caption = strArwDn
8250            .Controls(strSortLbl).ForeColor = CLR_DKRED
8260          Else
8270            strSortNow = strSortAsc
8280          End If
8290        End If
8300      End If
8310      .Controls(strSortLbl).Visible = True
8320      .Controls(strSortLine).Visible = True
8330      .OrderBy = strSortNow
8340      .OrderByOn = True
8350    End With

EXITP:
8360    Exit Sub

ERRH:
8370    Select Case ERR.Number
        Case Else
8380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8390    End Select
8400    Resume EXITP

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
