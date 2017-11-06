VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_SystemTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_SystemTable"

'VGC 10/11/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Acct Types:       A {cmdAcctTypes}
' **   Close:            C {cmdClose}
' **   Load Timer:       D {chkLoadTimer}
' **   State Codes:      E {cmdStateCodes}
' **   Invest Objects:   I {cmdInvestmentObjectives}
' **   Journal Types:    J {cmdJrnlTypes}
' **   Acct Groups:      O {cmdAcctTypeGroups}
' **   Asset Types:      S {cmdAssetTypes}
' **   Tax Codes:        T {cmdTaxCodes}
' **   Asset Groups:     U {cmdAssetTypeGroups}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm}

'7.5938"  frmAssetType
'6.6146"  frmTaxCode              A
'6.2813"  frmAccountType          A
'5.1667"  frmInvestmentObjective  B
'4.9896"  frmStateCode            B

'3.3646"  frmAccountTypeGrouping
'3.3646"  frmAssetTypeGrouping
'3.3646"  frmJournalType

Private lngTpp As Long
Private blnJustClose As Boolean, blnFormOpening As Boolean
Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           gblnSignal = False

160           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

170           gstrCrtRpt_Version = vbNullString  ' ** Borrowing this variable from Court Reports.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190           blnJustClose = False: blnFormOpening = False
200           blnGTR_Emblem = False: lngGTR_Stat = 0&

210           Select Case .chkLoadTimer_vis
              Case True
220             .chkLoadTimer.Visible = True
230             .chkLoadTimer.Enabled = False
                '.chkLoadTimer_lbl1.Visible = True
240           Case False
250             .chkLoadTimer.Visible = False
                '.chkLoadTimer_lbl1.Visible = False
260           End Select
270           chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
280           DoEvents

290         Else
300           Cancel = -1
310         End If
320       Else
330         Cancel = -1
340       End If
350     End With

360     DoCmd.Hourglass False

370     If Cancel = -1 Then
380       Beep
390       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
400     End If

EXITP:
410     Exit Sub

ERRH:
420     DoCmd.Hourglass False
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strUsername As String
        Dim intRetVal As Integer, Cancel As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Load Timer:       D {chkLoadTimer}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Timer Chk Vis:    T {Form_KeyDown}
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         intRetVal = 0
580         cmdClose_Click  ' ** Procedure: Below.
590       End Select
600     End If

        ' ** Alt keys.
610     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
620       Select Case intRetVal
          Case vbKeyD
630         With Me
640           intRetVal = 0
650           If .chkLoadTimer.Visible = True Then
660             If .chkLoadTimer.Enabled = True Then
670               .chkLoadTimer.SetFocus
680             End If
690             .chkLoadTimer = (Not .chkLoadTimer)
700             chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
710           End If
720         End With
730       Case vbKeyX
740         intRetVal = 0
750         cmdClose_Click  ' ** Procedure: Below.
760       End Select
770     End If

        ' ** Ctrl-Shift keys.
780     If intCtrlDown And (Not intAltDown) And intShiftDown Then
790       Select Case intRetVal
          Case vbKeyF
800         With Me
810           intRetVal = 0
820           .FocusHolder.SetFocus
830         End With
840       Case vbKeyT
850         With Me
860           intRetVal = 0
870           If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
880             FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
890           ElseIf CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
900             Select Case .chkLoadTimer.Visible
                Case True
910               .FocusHolder.SetFocus
920               DoEvents
930               .chkLoadTimer.Visible = False
                  '.chkLoadTimer_lbl1.Visible = False
940               gblnTimer = False
950             Case False
960               .chkLoadTimer.Visible = True
970               .chkLoadTimer.Enabled = False
                  '.chkLoadTimer_lbl1.Visible = True
980               chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
990             End Select
1000          End If
1010          .chkLoadTimer_vis = .chkLoadTimer.Visible
1020        End With
1030      Case vbKeyX
1040        strUsername = GetUserName  ' ** Module Function: modFileUtilities.
1050        If strUsername = gstrDevUserName Then
1060          intRetVal = 0
1070          blnJustClose = True
1080          SetOption_Dev  ' ** Module Function: modUtilities.
1090          gblnSignal = True  ' ** Borrowing this variable.
1100          DoCmd.Close acForm, THIS_NAME
1110          DoCmd.Hourglass False
1120        End If
1130      End Select
1140    End If

EXITP:
1150    KeyCode = intRetVal
1160    Exit Sub

ERRH:
1170    intRetVal = 0
1180    Select Case ERR.Number
        Case Else
1190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1200    End Select
1210    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim strFormName As String

        Static lngTmp01 As Long, lngTmp02 As Long

1310    With Me
1320      .TimerInterval = 0&
1330      If blnGTR_Emblem = True Then
1340        lngGTR_Stat = lngGTR_Stat + 1&
1350        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1360        lngTmp01 = lngTmp01 + 1&
1370        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1380        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1390        If lngTmp01 > 1& Then
1400          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1410        End If
1420        If lngTmp01 = 24& Then
1430          lngTmp01 = 0&
1440          lngTmp02 = lngTmp02 + 1&
1450          If lngTmp02 < 5& Then
1460            .TimerInterval = 50&
1470          Else
1480            blnGTR_Emblem = False
1490          End If
1500        Else
1510          .TimerInterval = 50&
1520        End If
1530      End If  ' ** blnGTR_Emblem.
1540      If gblnSetFocus = True Then
1550        gblnSetFocus = False
1560        strFormName = vbNullString
1570  On Error Resume Next
1580        Do Until strFormName = THIS_NAME
1590          DoCmd.SelectObject acForm, THIS_NAME, False
1600          strFormName = Screen.ActiveForm.Name
1610        Loop
1620  On Error GoTo ERRH
1630        If IsMaximized(Me) = True Then
1640          DoCmd.Restore
1650        End If
1660        DoEvents
1670      End If
1680      If blnGoingToReport = False Then
1690        Select Case gstrReturningForm
            Case "frmAccountType"
1700          .cmdAcctTypes.SetFocus
1710        Case "frmAccountTypeGrouping"
1720          .cmdAcctTypeGroups.SetFocus
1730        Case "frmAssetType"
1740          .cmdAssetTypes.SetFocus
1750        Case "frmAssetTypeGrouping"
1760          .cmdAssetTypeGroups.SetFocus
1770        Case "frmJournalType"
1780          .cmdJrnlTypes.SetFocus
1790        Case "frmTaxCode"
1800          .cmdTaxCodes.SetFocus
1810        Case "frmInvestmentObjective"
1820          .cmdInvestmentObjectives.SetFocus
1830        Case "frmStateCode"
1840          .cmdStateCodes.SetFocus
1850        Case Else
1860          .cmdAcctTypes.SetFocus
1870        End Select
1880        gstrReturningForm = vbNullString
1890        DoCmd.Hourglass False
1900      End If
1910      If gblnGoToReport = True Then
1920        DoCmd.Hourglass True  ' ** Make sure it's still running.
1930        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
1940        Else
1950          If blnGTR_Emblem = False Then
1960            blnGTR_Emblem = True
1970            lngTmp01 = 0&: lngTmp02 = 0&
1980          End If
1990          DoEvents
2000          Select Case blnGoingToReport
              Case True
2010            Select Case garr_varGoToReport(GTR_INV3)
                Case True
2020              Select Case garr_varGoToReport(GTR_CTL3)
                  Case "cmdAcctTypes"
2030                .TimerInterval = 50&
2040                cmdAcctTypes_Click  ' ** Procedure: Below.
2050              Case "cmdAcctTypeGroups"
2060                .TimerInterval = 50&
2070                cmdAcctTypeGroups_Click  ' ** Procedure: Below.
2080              Case "cmdAssetTypes"
2090                .TimerInterval = 50&
2100                cmdAssetTypes_Click  ' ** Procedure: Below.
2110              Case "cmdAssetTypeGroups"
2120                .TimerInterval = 50&
2130                cmdAssetTypeGroups_Click  ' ** Procedure: Below.
2140              Case "cmdInvestmentObjectives"
2150                .TimerInterval = 50&
2160                cmdInvestmentObjectives_Click  ' ** Procedure: Below.
2170              Case "cmdJrnlTypes"
2180                .TimerInterval = 50&
2190                cmdJrnlTypes_Click  ' ** Procedure: Below.
2200              Case "cmdStateCodes"
2210                .TimerInterval = 50&
2220                cmdStateCodes_Click  ' ** Procedure: Below.
2230              Case "cmdTaxCodes"
2240                .TimerInterval = 50&
2250                cmdTaxCodes_Click  ' ** Procedure: Below.
2260              Case Else
                    ' ** Shouldn't happen.
2270                blnGoingToReport = False
2280                DoCmd.Hourglass False
2290                gblnGoToReport = False
2300                blnGTR_Emblem = False
2310                GTREmblem_Off  ' ** Procedure: Below.
2320                Beep
2330              End Select
2340            Case False
                  ' ** Just sit there, and we're done.
2350              blnGoingToReport = False
2360              DoCmd.Hourglass False
2370              gblnGoToReport = False
2380              blnGTR_Emblem = False
2390              GTREmblem_Off  ' ** Procedure: Below.
2400              Beep
2410            End Select
2420          Case False
2430            If garr_varGoToReport(GTR_FRM3) = THIS_NAME Then
                  ' ** This is second menu in path.
2440              blnGoingToReport = True
                  ' ** No more menus.
2450              .Controls(garr_varGoToReport(GTR_CTL3)).SetFocus
2460              lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL3)).Top
2470              lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL3)).Left
2480              .GoToReport_arw_menu_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
2490              .GoToReport_arw_menu_img.Visible = True
2500              DoEvents
2510              .TimerInterval = 50&
2520            Else
                  ' ** Shouldn't be here, then!
2530              blnGoingToReport = False
2540              DoCmd.Hourglass False
2550              gblnGoToReport = False
2560              blnGTR_Emblem = False
2570              GTREmblem_Off  ' ** Procedure: Below.
2580              Beep
2590            End If
2600          End Select
2610        End If
2620      End If
2630      If gblnTimer = True Then
2640        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
2650          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
2660          gblnTimer = False
2670          DoCmd.Hourglass False
2680        End If
2690      End If
2700    End With

EXITP:
2710    Exit Sub

ERRH:
2720    DoCmd.Hourglass False
2730    blnGoingToReport = False
2740    gblnGoToReport = False
2750    Select Case ERR.Number
        Case Else
2760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2770    End Select
2780    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2810    With Me

2820      If gblnSignal = False Then  ' ** Borrowing this variable.
2830        DoCmd.Hourglass True
2840        DoEvents
2850      End If

          ' ** Make sure all reports are closed.
2860      Do While Reports.Count > 0
2870        DoCmd.Close acReport, Reports(0).Name
2880      Loop

2890      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2900      gblnAdmin = False
2910      gstrFormQuerySpec = vbNullString

2920      If blnJustClose = False Then

2930        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
2940        DoEvents

2950        gblnSetFocus = True
2960        gstrReturningForm = THIS_NAME
2970        Select Case strCallingForm
            Case "frmMenu_Maintenance"
2980          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Utility"
2990        End Select

3000      End If

3010    End With

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3110    With Me
3120      If .cmdClose.Caption <> "&Close" Then
3130        .cmdClose.Caption = "&Close"
3140      End If
3150    End With

EXITP:
3160    Exit Sub

ERRH:
3170    Select Case ERR.Number
        Case Else
3180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3190    End Select
3200    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

3310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub Header_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_box_MouseMove"

3410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub Header_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_MouseMove"

3510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub Header_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl2_MouseMove"

3610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub cmdClose_Click()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3710    With Me
3720      If blnFormOpening = False Then

3730        DoCmd.Hourglass True
3740        DoEvents

3750        .TimerInterval = 0&

3760        blnJustClose = False
3770        DoCmd.Close acForm, THIS_NAME

3780      End If
3790    End With

EXITP:
3800    Exit Sub

ERRH:
3810    blnFormOpening = False
3820    DoCmd.Hourglass False
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
3970        With Me
3980          intRetVal = 0
3990          .cmdAcctTypes.SetFocus
4000        End With
4010      End Select
4020    End If

        ' ** Shift keys.
4030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4040      Select Case intRetVal
          Case vbKeyTab
4050        With Me
4060          intRetVal = 0
4070          If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
4080            .chkLoadTimer.SetFocus
4090          Else
4100            .cmdTaxCodes.SetFocus
4110          End If
4120        End With
4130      End Select
4140    End If

EXITP:
4150    KeyCode = intRetVal
4160    Exit Sub

ERRH:
4170    intRetVal = 0
4180    Select Case ERR.Number
        Case Else
4190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4200    End Select
4210    Resume EXITP

End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_MouseMove"

4310    With Me
4320      .cmdClose.Caption = "< Ba&ck"
4330    End With

EXITP:
4340    Exit Sub

ERRH:
4350    Select Case ERR.Number
        Case Else
4360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4370    End Select
4380    Resume EXITP

End Sub

Private Sub cmdAcctTypes_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctTypes_Click"

        Dim strDocName As String

4410    With Me
4420      If blnFormOpening = False Then

4430        DoCmd.Hourglass True
4440        DoEvents

4450        .TimerInterval = 0&

4460        gblnSetFocus = True
4470        blnJustClose = True: blnFormOpening = True
4480        strDocName = "frmAccountType"
4490        gstrReturningForm = strDocName
4500        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
4510        If gblnGoToReport = True Then
4520          Forms(strDocName).TimerInterval = 50&
4530          blnGoingToReport = False
4540          blnGTR_Emblem = False
4550        End If
4560        DoCmd.Close acForm, THIS_NAME

4570      End If
4580    End With

EXITP:
4590    Exit Sub

ERRH:
4600    blnFormOpening = False
4610    DoCmd.Hourglass False
4620    Select Case ERR.Number
        Case Else
4630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4640    End Select
4650    Resume EXITP

End Sub

Private Sub cmdAcctTypes_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctTypes_KeyDown"

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
          Case vbKeyTab
4770        With Me
4780          intRetVal = 0
4790          .cmdAcctTypeGroups.SetFocus
4800        End With
4810      End Select
4820    End If

        ' ** Shift keys.
4830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4840      Select Case intRetVal
          Case vbKeyTab
4850        With Me
4860          intRetVal = 0
4870          .cmdClose.SetFocus
4880        End With
4890      End Select
4900    End If

EXITP:
4910    KeyCode = intRetVal
4920    Exit Sub

ERRH:
4930    intRetVal = 0
4940    Select Case ERR.Number
        Case Else
4950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4960    End Select
4970    Resume EXITP

End Sub

Private Sub cmdAcctTypeGroups_Click()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctTypeGroups_Click"

        Dim strDocName As String

5010    With Me
5020      If blnFormOpening = False Then

5030        DoCmd.Hourglass True
5040        DoEvents

5050        .TimerInterval = 0&

5060        gblnSetFocus = True
5070        blnJustClose = True: blnFormOpening = True
5080        strDocName = "frmAccountTypeGrouping"
5090        gstrReturningForm = strDocName
5100        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
5110        If gblnGoToReport = True Then
5120          Forms(strDocName).TimerInterval = 50&
5130          blnGoingToReport = False
5140          blnGTR_Emblem = False
5150        End If
5160        DoCmd.Close acForm, THIS_NAME

5170      End If
5180    End With

EXITP:
5190    Exit Sub

ERRH:
5200    blnFormOpening = False
5210    DoCmd.Hourglass False
5220    Select Case ERR.Number
        Case Else
5230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5240    End Select
5250    Resume EXITP

End Sub

Private Sub cmdAssetTypes_Click()

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetTypes_Click"

        Dim strDocName As String

5310    With Me
5320      If blnFormOpening = False Then

5330        DoCmd.Hourglass True
5340        DoEvents

5350        .TimerInterval = 0&

5360        gblnSetFocus = True
5370        blnJustClose = True: blnFormOpening = True
5380        strDocName = "frmAssetType"
5390        gstrReturningForm = strDocName
5400        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
5410        If gblnGoToReport = True Then
5420          Forms(strDocName).TimerInterval = 50&
5430          blnGoingToReport = False
5440          blnGTR_Emblem = False
5450        End If
5460        DoCmd.Close acForm, THIS_NAME

5470      End If
5480    End With

EXITP:
5490    Exit Sub

ERRH:
5500    blnFormOpening = False
5510    DoCmd.Hourglass False
5520    Select Case ERR.Number
        Case Else
5530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5540    End Select
5550    Resume EXITP

End Sub

Private Sub cmdAssetTypeGroups_Click()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetTypeGroups_Click"

        Dim strDocName As String

5610    With Me
5620      If blnFormOpening = False Then

5630        DoCmd.Hourglass True
5640        DoEvents

5650        .TimerInterval = 0&

5660        gblnSetFocus = True
5670        blnJustClose = True: blnFormOpening = True
5680        strDocName = "frmAssetTypeGrouping"
5690        gstrReturningForm = strDocName
5700        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
5710        If gblnGoToReport = True Then
5720          Forms(strDocName).TimerInterval = 50&
5730          blnGoingToReport = False
5740          blnGTR_Emblem = False
5750        End If
5760        DoCmd.Close acForm, THIS_NAME

5770      End If
5780    End With

EXITP:
5790    Exit Sub

ERRH:
5800    blnFormOpening = False
5810    DoCmd.Hourglass False
5820    Select Case ERR.Number
        Case Else
5830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5840    End Select
5850    Resume EXITP

End Sub

Private Sub cmdInvestmentObjectives_Click()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdInvestmentObjectives_Click"

        Dim strDocName As String

5910    With Me
5920      If blnFormOpening = False Then

5930        DoCmd.Hourglass True
5940        DoEvents

5950        .TimerInterval = 0&

5960        gblnSetFocus = True
5970        blnJustClose = True: blnFormOpening = True
5980        strDocName = "frmInvestmentObjective"
5990        gstrReturningForm = strDocName
6000        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6010        If gblnGoToReport = True Then
6020          Forms(strDocName).TimerInterval = 50&
6030          blnGoingToReport = False
6040          blnGTR_Emblem = False
6050        End If
6060        DoCmd.Close acForm, THIS_NAME

6070      End If
6080    End With

EXITP:
6090    Exit Sub

ERRH:
6100    blnFormOpening = False
6110    DoCmd.Hourglass False
6120    Select Case ERR.Number
        Case Else
6130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6140    End Select
6150    Resume EXITP

End Sub

Private Sub cmdJrnlTypes_Click()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdJrnlTypes_Click"

        Dim strDocName As String

6210    With Me
6220      If blnFormOpening = False Then

6230        DoCmd.Hourglass True
6240        DoEvents

6250        .TimerInterval = 0&

6260        gblnSetFocus = True
6270        blnJustClose = True: blnFormOpening = True
6280        strDocName = "frmJournalType"
6290        gstrReturningForm = strDocName
6300        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6310        If gblnGoToReport = True Then
6320          Forms(strDocName).TimerInterval = 50&
6330          blnGoingToReport = False
6340          blnGTR_Emblem = False
6350        End If
6360        DoCmd.Close acForm, THIS_NAME

6370      End If
6380    End With

EXITP:
6390    Exit Sub

ERRH:
6400    blnFormOpening = False
6410    DoCmd.Hourglass False
6420    Select Case ERR.Number
        Case Else
6430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6440    End Select
6450    Resume EXITP

End Sub

Private Sub cmdStateCodes_Click()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStateCodes_Click"

        Dim strDocName As String

6510    With Me
6520      If blnFormOpening = False Then

6530        DoCmd.Hourglass True
6540        DoEvents

6550        .TimerInterval = 0&

6560        gblnSetFocus = True
6570        blnJustClose = True: blnFormOpening = True
6580        strDocName = "frmStateCode"
6590        gstrReturningForm = strDocName
6600        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6610        If gblnGoToReport = True Then
6620          Forms(strDocName).TimerInterval = 50&
6630          blnGoingToReport = False
6640          blnGTR_Emblem = False
6650        End If
6660        DoCmd.Close acForm, THIS_NAME

6670      End If
6680    End With

EXITP:
6690    Exit Sub

ERRH:
6700    blnFormOpening = False
6710    DoCmd.Hourglass False
6720    Select Case ERR.Number
        Case Else
6730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6740    End Select
6750    Resume EXITP

End Sub

Private Sub cmdTaxCodes_Click()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxCodes_Click"

        Dim strDocName As String

6810    With Me
6820      If blnFormOpening = False Then

6830        DoCmd.Hourglass True
6840        DoEvents

6850        .TimerInterval = 0&

6860        gblnSetFocus = True
6870        blnJustClose = True: blnFormOpening = True
6880        strDocName = "frmTaxCode"
6890        gstrReturningForm = strDocName
6900        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6910        If gblnGoToReport = True Then
6920          Forms(strDocName).TimerInterval = 50&
6930          blnGoingToReport = False
6940          blnGTR_Emblem = False
6950        End If
6960        DoCmd.Close acForm, THIS_NAME

6970      End If
6980    End With

EXITP:
6990    Exit Sub

ERRH:
7000    blnFormOpening = False
7010    DoCmd.Hourglass False
7020    Select Case ERR.Number
        Case Else
7030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7040    End Select
7050    Resume EXITP

End Sub

Private Sub cmdTaxCodes_KeyDown(KeyCode As Integer, Shift As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxCodes_KeyDown"

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
          Case vbKeyTab
7170        With Me
7180          intRetVal = 0
7190          If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
7200            .chkLoadTimer.SetFocus
7210          Else
7220            .cmdClose.SetFocus
7230          End If
7240        End With
7250      End Select
7260    End If

        ' ** Shift keys.
7270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7280      Select Case intRetVal
          Case vbKeyTab
7290        With Me
7300          intRetVal = 0
7310          .cmdStateCodes.SetFocus
7320        End With
7330      End Select
7340    End If

EXITP:
7350    KeyCode = intRetVal
7360    Exit Sub

ERRH:
7370    intRetVal = 0
7380    Select Case ERR.Number
        Case Else
7390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7400    End Select
7410    Resume EXITP

End Sub

Private Sub chkLoadTimer_AfterUpdate()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_AfterUpdate"

7510    With Me
7520      If IsLoaded("frmLoadTimer", acForm) = False Then  ' ** Module Function: modFileUtilities.
            ' ** Don't turn it off if it's still playing out a timing sequence.
7530        gblnTimer = False
7540      End If
7550      Select Case .chkLoadTimer
          Case True
7560        .chkLoadTimer_lbl.FontBold = True
7570        If .chkLoadTimer.Visible = True Then  ' ** This may get turned off by the Form_Timer() sequence.
7580          gblnTimer = True
7590        End If
7600      Case False
7610        .chkLoadTimer_lbl.FontBold = False
7620      End Select
7630    End With

EXITP:
7640    Exit Sub

ERRH:
7650    Select Case ERR.Number
        Case Else
7660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7670    End Select
7680    Resume EXITP

End Sub

Private Sub chkLoadTimer_KeyDown(KeyCode As Integer, Shift As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7720    intShiftDown = (Shift And acShiftMask) > 0
7730    intAltDown = (Shift And acAltMask) > 0
7740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7760      Select Case intRetVal
          Case vbKeyTab
7770        With Me
7780          intRetVal = 0
7790          .cmdClose.SetFocus
7800        End With
7810      End Select
7820    End If

        ' ** Shift keys.
7830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7840      Select Case intRetVal
          Case vbKeyTab
7850        With Me
7860          intRetVal = 0
7870          .cmdTaxCodes.SetFocus
7880        End With
7890      End Select
7900    End If

EXITP:
7910    KeyCode = intRetVal
7920    Exit Sub

ERRH:
7930    intRetVal = 0
7940    Select Case ERR.Number
        Case Else
7950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7960    End Select
7970    Resume EXITP

End Sub

Public Sub GTREmblem_Off()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

8010    With Me
8020      blnGTR_Emblem = False
8030      For lngX = 1& To 24&
8040  On Error Resume Next
8050        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
8060        If ERR.Number <> 0 Then
8070  On Error GoTo ERRH
8080          Exit For
8090        Else
8100  On Error GoTo ERRH
8110        End If
8120      Next
8130    End With

EXITP:
8140    Exit Sub

ERRH:
8150    Select Case ERR.Number
        Case Else
8160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8170    End Select
8180    Resume EXITP

End Sub
