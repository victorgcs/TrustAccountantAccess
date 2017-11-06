VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_CourtReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_CourtReport"

'VGC 10/11/2017: CHANGES!

' ** For GTR, sets focus, then continues.

' ** Shortcut Alt keys responsive from this form:
' **   California:       A {cmdCourtReports_California}
' **   Close:            C {cmdClose}
' **   Load Timer:       D {chkLoadTimer}
' **   Florida:          F {cmdCourtReports_Florida}
' **   National:         N {cmdCourtReports_NationalStandard}
' **   Exit:             X {cmdClose}
' **   New York:         Y {cmdCourtReports_NewYork}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm}

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
230             .chkLoadTimer_lbl1.Visible = True
240             .chkLoadTimer_lbl2.Visible = True
250             .chkLoadTimer_lbl3.Visible = True
260             .chkLoadTimer_lbl4.Visible = True
270           Case False
280             .chkLoadTimer.Visible = False
290             .chkLoadTimer_lbl1.Visible = False
300             .chkLoadTimer_lbl2.Visible = False
310             .chkLoadTimer_lbl3.Visible = False
320             .chkLoadTimer_lbl4.Visible = False
330           End Select
340           chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
350           DoEvents

360           gstrAccountNo = vbNullString
370           gstrCrtRpt_Account = vbNullString
380           gdatStartDate = 0
390           gdatEndDate = 0

400           DoCmd.Hourglass False

410         Else
420           Cancel = -1
430         End If
440       Else
450         Cancel = -1
460       End If
470     End With

480     If Cancel = -1 Then
490       Beep
500       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
510     End If

EXITP:
520     Exit Sub

ERRH:
530     DoCmd.Hourglass False
540     Select Case ERR.Number
        Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
560     End Select
570     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strUsername As String
        Dim intRetVal As Integer, Cancel As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Load Timer:       D {chkLoadTimer}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Timer Chk Vis:    T {Form_KeyDown}
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyEscape
670         intRetVal = 0
680         cmdClose_Click  ' ** Procedure: Below.
690       End Select
700     End If

        ' ** Alt keys.
710     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
720       Select Case intRetVal
          Case vbKeyD
730         With Me
740           intRetVal = 0
750           If .chkLoadTimer.Visible = True Then
760             If .chkLoadTimer.Enabled = True Then
770               .chkLoadTimer.SetFocus
780             End If
790             .chkLoadTimer = (Not .chkLoadTimer)
800             chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
810           End If
820         End With
830       Case vbKeyX
840         intRetVal = 0
850         cmdClose_Click  ' ** Procedure: Below.
860       End Select
870     End If

        ' ** Ctrl-Shift keys.
880     If intCtrlDown And (Not intAltDown) And intShiftDown Then
890       Select Case intRetVal
          Case vbKeyF
900         With Me
910           intRetVal = 0
920           .FocusHolder.SetFocus
930         End With
940       Case vbKeyT
950         With Me
960           intRetVal = 0
970           If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
980             FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
990           ElseIf CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
1000            Select Case .chkLoadTimer.Visible
                Case True
1010              .FocusHolder.SetFocus
1020              DoEvents
1030              .chkLoadTimer.Visible = False
1040              .chkLoadTimer_lbl1.Visible = False
1050              .chkLoadTimer_lbl2.Visible = False
1060              .chkLoadTimer_lbl3.Visible = False
1070              .chkLoadTimer_lbl4.Visible = False
1080              gblnTimer = False
1090            Case False
1100              .chkLoadTimer.Visible = True
1110              .chkLoadTimer_lbl1.Visible = True
1120              .chkLoadTimer_lbl2.Visible = True
1130              .chkLoadTimer_lbl3.Visible = True
1140              .chkLoadTimer_lbl4.Visible = True
1150              chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
1160            End Select
1170          End If
1180          .chkLoadTimer_vis = .chkLoadTimer.Visible
1190        End With
1200      Case vbKeyX
1210        strUsername = GetUserName  ' ** Module Function: modFileUtilities.
1220        If strUsername = gstrDevUserName Then
1230          intRetVal = 0
1240          blnJustClose = True
1250          SetOption_Dev  ' ** Module Function: modUtilities.
1260          gblnSignal = True  ' ** Borrowing this variable.
1270          DoCmd.Close acForm, THIS_NAME
1280          DoCmd.Hourglass False
1290        End If
1300      End Select
1310    End If

EXITP:
1320    KeyCode = intRetVal
1330    Exit Sub

ERRH:
1340    intRetVal = 0
1350    Select Case ERR.Number
        Case Else
1360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1370    End Select
1380    Resume EXITP

End Sub

Private Sub Form_Timer()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim strFormName As String

        Static lngTmp01 As Long, lngTmp02 As Long

1410    With Me
1420      .TimerInterval = 0&
1430      If blnGTR_Emblem = True Then
1440        lngGTR_Stat = lngGTR_Stat + 1&
1450        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1460        lngTmp01 = lngTmp01 + 1&
1470        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1480        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1490        If lngTmp01 > 1& Then
1500          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1510        End If
1520        If lngTmp01 = 24& Then
1530          lngTmp01 = 0&
1540          lngTmp02 = lngTmp02 + 1&
1550          If lngTmp02 < 5& Then
1560            .TimerInterval = 50&
1570          Else
1580            blnGTR_Emblem = False
1590          End If
1600        Else
1610          .TimerInterval = 50&
1620        End If
1630      End If  ' ** blnGTR_Emblem.
1640      If gblnSetFocus = True Then
1650        gblnSetFocus = False
1660        strFormName = vbNullString
1670  On Error Resume Next
1680        Do Until strFormName = THIS_NAME
1690          DoCmd.SelectObject acForm, THIS_NAME, False
1700          strFormName = Screen.ActiveForm.Name
1710        Loop
1720  On Error GoTo ERRH
1730        DoEvents
1740      End If
1750      If blnGoingToReport = False Then
1760        Select Case gstrReturningForm
            Case "frmRpt_CourtReports_NS"
1770          .cmdCourtReports_NationalStandard.SetFocus
1780        Case "frmRpt_CourtReports_CA"
1790          .cmdCourtReports_California.SetFocus
1800        Case "frmRpt_CourtReports_FL"
1810          .cmdCourtReports_Florida.SetFocus
1820        Case "frmRpt_CourtReports_NY"
1830          .cmdCourtReports_NewYork.SetFocus
1840        Case Else
1850          .cmdCourtReports_NationalStandard.SetFocus
1860        End Select
1870        gstrReturningForm = vbNullString
1880        DoCmd.Hourglass False
1890      End If
1900      If gblnGoToReport = True Then
1910        DoCmd.Hourglass True  ' ** Make sure it's still running.
1920        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
1930        Else
1940          If blnGTR_Emblem = False Then
1950            blnGTR_Emblem = True
1960            lngTmp01 = 0&: lngTmp02 = 0&
1970          End If
1980          DoEvents
1990          Select Case blnGoingToReport
              Case True
2000            Select Case garr_varGoToReport(GTR_INV2)
                Case True
2010              Select Case garr_varGoToReport(GTR_CTL2)
                  Case "cmdCourtReports_NationalStandard"
2020                .TimerInterval = 50&
2030                cmdCourtReports_NationalStandard_Click  ' ** Procedure: Below.
2040              Case "cmdCourtReports_California"
2050                .TimerInterval = 50&
2060                cmdCourtReports_California_Click  ' ** Procedure: Below.
2070              Case "cmdCourtReports_Florida"
2080                .TimerInterval = 50&
2090                cmdCourtReports_Florida_Click  ' ** Procedure: Below.
2100              Case "cmdCourtReports_NewYork"
2110                .TimerInterval = 50&
2120                cmdCourtReports_NewYork_Click  ' ** Procedure: Below.
2130              Case Else
                    ' ** Can't imagine what it could be.
2140                blnGoingToReport = False
2150                DoCmd.Hourglass False
2160                gblnGoToReport = False
2170                blnGTR_Emblem = False
2180                GTREmblem_Off  ' ** Procedure: Below.
2190                Beep
2200              End Select
2210            Case False
                  ' ** Just sit there, and we're done.
2220              blnGoingToReport = False
2230              DoCmd.Hourglass False
2240              gblnGoToReport = False
2250              blnGTR_Emblem = False
2260              GTREmblem_Off  ' ** Procedure: Below.
2270              Beep
2280            End Select
2290          Case False
2300            If garr_varGoToReport(GTR_FRM2) = THIS_NAME Then
                  ' ** This is second menu in path.
2310              blnGoingToReport = True
2320              Select Case IsNull(garr_varGoToReport(GTR_FRM3))
                  Case True
                    ' ** No more menus.
2330                .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
2340                lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL2)).Top
2350                lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL2)).Left
2360                .GoToReport_arw_menu_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
2370                .GoToReport_arw_menu_img.Visible = True
2380                DoEvents
2390                .TimerInterval = 50&
2400              Case False
                    ' ** There are no more from here!
2410                blnGoingToReport = False
2420                DoCmd.Hourglass False
2430                gblnGoToReport = False
2440                blnGTR_Emblem = False
2450                GTREmblem_Off  ' ** Procedure: Below.
2460                Beep
2470              End Select
2480            Else
                  ' ** Just sit there, and we're done.
2490              blnGoingToReport = False
2500              DoCmd.Hourglass False
2510              gblnGoToReport = False
2520              blnGTR_Emblem = False
2530              GTREmblem_Off  ' ** Procedure: Below.
2540              Beep
2550            End If
2560          End Select
2570        End If
2580      End If  ' ** gblnGoToReport
2590      If gblnTimer = True Then
2600        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
2610          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
2620          gblnTimer = False
2630          DoCmd.Hourglass False
2640        End If
2650      End If
2660    End With

        'HERE! 1  Form_Open()  frmMenu_CourtReport  True
        'HERE! 2  Form_Timer()  frmMenu_CourtReport  True

EXITP:
2670    Exit Sub

ERRH:
2680    DoCmd.Hourglass False
2690    blnGoingToReport = False
2700    gblnGoToReport = False
2710    Select Case ERR.Number
        Case Else
2720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2730    End Select
2740    Resume EXITP

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

2900      blnGoingToReport = False

2910      gstrAccountNo = vbNullString
2920      gstrCrtRpt_Account = vbNullString
2930      gdatStartDate = 0
2940      gdatEndDate = 0

2950      If blnJustClose = False Then

2960        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
2970        DoEvents

2980        gstrFormQuerySpec = vbNullString
2990        gblnSetFocus = True
3000        gstrReturningForm = THIS_NAME
3010        Select Case strCallingForm
            Case "frmMenu_Report"
3020          DoCmd.OpenForm "frmMenu_Report", , , , , , "frmMenu_Main"
3030        End Select

3040      End If

3050    End With

EXITP:
3060    Exit Sub

ERRH:
3070    DoCmd.Hourglass False
3080    Select Case ERR.Number
        Case Else
3090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3100    End Select
3110    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3210    With Me
3220      If .cmdClose.Caption <> "&Close" Then
3230        .cmdClose.Caption = "&Close"
3240      End If
3250    End With

EXITP:
3260    Exit Sub

ERRH:
3270    Select Case ERR.Number
        Case Else
3280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290    End Select
3300    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

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

Private Sub Header_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_box_MouseMove"

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

Private Sub Header_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_MouseMove"

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

Private Sub Header_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl2_MouseMove"

3710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub cmdClose_Click()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3810    With Me
3820      If blnFormOpening = False Then

3830        DoCmd.Hourglass True
3840        DoEvents

3850        .TimerInterval = 0&

3860        blnJustClose = False
3870        DoCmd.Close acForm, THIS_NAME

3880      End If
3890    End With

EXITP:
3900    Exit Sub

ERRH:
3910    blnFormOpening = False
3920    DoCmd.Hourglass False
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4020    intShiftDown = (Shift And acShiftMask) > 0
4030    intAltDown = (Shift And acAltMask) > 0
4040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4060      Select Case intRetVal
          Case vbKeyTab
4070        With Me
4080          intRetVal = 0
4090          .cmdCourtReports_NationalStandard.SetFocus
4100        End With
4110      End Select
4120    End If

        ' ** Shift keys.
4130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4140      Select Case intRetVal
          Case vbKeyTab
4150        With Me
4160          intRetVal = 0
4170          If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
4180            .chkLoadTimer.SetFocus
4190          Else
4200            .cmdCourtReports_Florida.SetFocus
4210          End If
4220        End With
4230      End Select
4240    End If

EXITP:
4250    KeyCode = intRetVal
4260    Exit Sub

ERRH:
4270    intRetVal = 0
4280    Select Case ERR.Number
        Case Else
4290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4300    End Select
4310    Resume EXITP

End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_MouseMove"

4410    With Me
4420      .cmdClose.Caption = "< Ba&ck"
4430    End With

EXITP:
4440    Exit Sub

ERRH:
4450    Select Case ERR.Number
        Case Else
4460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4470    End Select
4480    Resume EXITP

End Sub

Private Sub cmdCourtReports_NationalStandard_Click()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCourtReports_NationalStandard_Click"

        Dim strDocName As String

4510    With Me
4520      If blnFormOpening = False Then

4530        strDocName = "frmRpt_CourtReports_NS"

4540        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
4550        DoEvents

4560        If .chkLoadTimer.Visible = True Then
4570          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
4580          If gblnTimer = True Then
                'gblnSetFocus = True
4590            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
4600          End If
4610        End If
4620        DoEvents

4630        DoCmd.Hourglass True
4640        DoEvents

4650        .TimerInterval = 0&

4660        gblnSetFocus = True: blnJustClose = True
4670        DoCmd.Close acForm, THIS_NAME
4680        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
4690        If gblnGoToReport = True Then
4700          Forms(strDocName).TimerInterval = 50&
4710          blnGoingToReport = False
4720          blnGTR_Emblem = False
4730        End If

4740        DoCmd.Hourglass False

4750      End If
4760    End With

EXITP:
4770    Exit Sub

ERRH:
4780    blnFormOpening = False
4790    DoCmd.Hourglass False
4800    Select Case ERR.Number
        Case Else
4810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4820    End Select
4830    Resume EXITP

End Sub

Private Sub cmdCourtReports_NationalStandard_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCourtReports_NationalStandard_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4920    intShiftDown = (Shift And acShiftMask) > 0
4930    intAltDown = (Shift And acAltMask) > 0
4940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4960      Select Case intRetVal
          Case vbKeyTab
4970        With Me
4980          intRetVal = 0
4990          .cmdCourtReports_California.SetFocus
5000        End With
5010      End Select
5020    End If

        ' ** Shift keys.
5030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5040      Select Case intRetVal
          Case vbKeyTab
5050        With Me
5060          intRetVal = 0
5070          .cmdClose.SetFocus
5080        End With
5090      End Select
5100    End If

EXITP:
5110    KeyCode = intRetVal
5120    Exit Sub

ERRH:
5130    intRetVal = 0
5140    Select Case ERR.Number
        Case Else
5150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5160    End Select
5170    Resume EXITP

End Sub

Private Sub cmdCourtReports_California_Click()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCourtReports_California_Click"

        Dim strDocName As String

5210    With Me
5220      If blnFormOpening = False Then

5230        strDocName = "frmRpt_CourtReports_CA"

5240        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
5250        DoEvents

5260        If .chkLoadTimer.Visible = True Then
5270          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
5280          If gblnTimer = True Then
                'gblnSetFocus = True
5290            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
5300          End If
5310        End If
5320        DoEvents

5330        DoCmd.Hourglass True
5340        DoEvents

5350        .TimerInterval = 0&

5360        gblnSetFocus = True: blnJustClose = True
5370        DoCmd.Close acForm, THIS_NAME
5380        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
5390        If gblnGoToReport = True Then
5400          Forms(strDocName).TimerInterval = 50&
5410          blnGoingToReport = False
5420          blnGTR_Emblem = False
5430        End If

5440        DoCmd.Hourglass False

5450      End If
5460    End With

EXITP:
5470    Exit Sub

ERRH:
5480    blnFormOpening = False
5490    DoCmd.Hourglass False
5500    Select Case ERR.Number
        Case Else
5510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5520    End Select
5530    Resume EXITP

End Sub

Private Sub cmdCourtReports_Florida_Click()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCourtReports_Florida_Click"

        Dim strDocName As String

5610    With Me
5620      If blnFormOpening = False Then

5630        strDocName = "frmRpt_CourtReports_FL"

5640        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
5650        DoEvents

5660        If .chkLoadTimer.Visible = True Then
5670          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
5680          If gblnTimer = True Then
                'gblnSetFocus = True
5690            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
5700          End If
5710        End If
5720        DoEvents

5730        DoCmd.Hourglass True
5740        DoEvents

5750        .TimerInterval = 0&

5760        gblnSetFocus = True: blnJustClose = True
5770        DoCmd.Close acForm, THIS_NAME
5780        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
5790        If gblnGoToReport = True Then
5800          Forms(strDocName).TimerInterval = 50&
5810          blnGoingToReport = False
5820          blnGTR_Emblem = False
5830        End If

5840        DoCmd.Hourglass False

5850      End If
5860    End With

EXITP:
5870    Exit Sub

ERRH:
5880    blnFormOpening = False
5890    DoCmd.Hourglass False
5900    Select Case ERR.Number
        Case Else
5910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5920    End Select
5930    Resume EXITP

End Sub

Private Sub cmdCourtReports_NewYork_Click()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCourtReports_NewYork_Click"

        Dim strDocName As String

6010    With Me
6020      If blnFormOpening = False Then

6030        strDocName = "frmRpt_CourtReports_NY"

6040        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
6050        DoEvents

6060        If .chkLoadTimer.Visible = True Then
6070          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
6080          If gblnTimer = True Then
                'gblnSetFocus = True
6090            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
6100          End If
6110        End If
6120        DoEvents

6130        DoCmd.Hourglass True
6140        DoEvents

6150        .TimerInterval = 0&

6160        gblnSetFocus = True: blnJustClose = True
6170        DoCmd.Close acForm, THIS_NAME
6180        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6190        If gblnGoToReport = True Then
6200          Forms(strDocName).TimerInterval = 50&
6210          blnGoingToReport = False
6220          blnGTR_Emblem = False
6230        End If

6240        DoCmd.Hourglass False

6250      End If
6260    End With

EXITP:
6270    Exit Sub

ERRH:
6280    blnFormOpening = False
6290    DoCmd.Hourglass False
6300    Select Case ERR.Number
        Case Else
6310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6320    End Select
6330    Resume EXITP

End Sub

Private Sub cmdCourtReports_NewYork_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCourtReports_NewYork_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6420    intShiftDown = (Shift And acShiftMask) > 0
6430    intAltDown = (Shift And acAltMask) > 0
6440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6460      Select Case intRetVal
          Case vbKeyTab
6470        With Me
6480          intRetVal = 0
6490          If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
6500            .chkLoadTimer.SetFocus
6510          Else
6520            .cmdClose.SetFocus
6530          End If
6540        End With
6550      End Select
6560    End If

        ' ** Shift keys.
6570    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6580      Select Case intRetVal
          Case vbKeyTab
6590        With Me
6600          intRetVal = 0
6610          .cmdCourtReports_Florida.SetFocus
6620        End With
6630      End Select
6640    End If

EXITP:
6650    KeyCode = intRetVal
6660    Exit Sub

ERRH:
6670    intRetVal = 0
6680    Select Case ERR.Number
        Case Else
6690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6700    End Select
6710    Resume EXITP

End Sub

Private Sub chkLoadTimer_AfterUpdate()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_AfterUpdate"

6810    With Me
6820      If IsLoaded("frmLoadTimer", acForm) = False Then  ' ** Module Function: modFileUtilities.
            ' ** Don't turn it off if it's still playing out a timing sequence.
6830        gblnTimer = False
6840      End If
6850      Select Case .chkLoadTimer
          Case True
6860        .chkLoadTimer_lbl.FontBold = True
6870        If .chkLoadTimer.Visible = True Then  ' ** This may get turned off by the Form_Timer() sequence.
6880          gblnTimer = True
6890        End If
6900      Case False
6910        .chkLoadTimer_lbl.FontBold = False
6920      End Select
6930    End With

EXITP:
6940    Exit Sub

ERRH:
6950    Select Case ERR.Number
        Case Else
6960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6970    End Select
6980    Resume EXITP

End Sub

Private Sub chkLoadTimer_KeyDown(KeyCode As Integer, Shift As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7020    intShiftDown = (Shift And acShiftMask) > 0
7030    intAltDown = (Shift And acAltMask) > 0
7040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7060      Select Case intRetVal
          Case vbKeyTab
7070        With Me
7080          intRetVal = 0
7090          .cmdClose.SetFocus
7100        End With
7110      End Select
7120    End If

        ' ** Shift keys.
7130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7140      Select Case intRetVal
          Case vbKeyTab
7150        With Me
7160          intRetVal = 0
7170          .cmdCourtReports_NewYork.SetFocus
7180        End With
7190      End Select
7200    End If

EXITP:
7210    KeyCode = intRetVal
7220    Exit Sub

ERRH:
7230    intRetVal = 0
7240    Select Case ERR.Number
        Case Else
7250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7260    End Select
7270    Resume EXITP

End Sub

Public Sub GTREmblem_Off()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

7310    With Me
7320      blnGTR_Emblem = False
7330      For lngX = 1& To 24&
7340  On Error Resume Next
7350        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
7360        If ERR.Number <> 0 Then
7370  On Error GoTo ERRH
7380          Exit For
7390        Else
7400  On Error GoTo ERRH
7410        End If
7420      Next
7430    End With

EXITP:
7440    Exit Sub

ERRH:
7450    Select Case ERR.Number
        Case Else
7460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7470    End Select
7480    Resume EXITP

End Sub
