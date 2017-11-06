VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPortfolioModeling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmPortfolioModeling"

'VGC 09/23/2017: CHANGES!

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Account Select:   A {opgReturn_optSelect}
' **   Close:            C {cmdClose}
' **   Model:            M {Model_input on frmPortfolioModeling_Sub}
' **   Invest Objective: O {cmbInvestmentObjectives}
' **   Print Report:     P {cmdPrintReport}
' **   Asset Menu:       S {opgReturn_optMenu}
' **   Exit:             X {cmdClose}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:           F5 {RecalcTots on frmPortfolioModeling_Sub}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmPortfolioModeling_Sub}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean
Private blnRecalcTotsNotReady As Boolean, blnViewOnly As Boolean
Private lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim usr As DAO.User, grp As DAO.Group
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           CLR_DISABLED_FG = CLR_DKGRY
160           CLR_DISABLED_BG = CLR_LTTEAL

170           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

180           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

190           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

              ' ** Check user status, and lock all fields if ViewOnly.
200           blnViewOnly = False
210           For Each grp In DBEngine.Workspaces(0).Groups
220             If grp.Name = "ViewOnly" Then
230               For Each usr In grp.Users
240                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
250                   blnViewOnly = True
260                   Exit For
270                 End If
280               Next
290             End If
300           Next

310           If blnViewOnly = True Then
320             .frmPortfolioModeling_Sub.Form.Model_input.Locked = True
330           End If

340           opgReturn_AfterUpdate  ' ** Procedure: Below.

350         Else
360           Cancel = -1
370         End If
380       Else
390         Cancel = -1
400       End If
410     End With

420     If Cancel = -1 Then
430       DoCmd.Hourglass False
440       Beep
450       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
460     End If

EXITP:
470     Set usr = Nothing
480     Set grp = Nothing
490     Exit Sub

ERRH:
500     DoCmd.Hourglass False
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me
620       .accountno.ForeColor = CLR_DISABLED_FG
630       .accountno.BackColor = CLR_DISABLED_BG
640       .shortname.ForeColor = CLR_DISABLED_FG
650       .shortname.BackColor = CLR_DISABLED_BG
660       .ModelTot.ForeColor = CLR_DISABLED_FG
670       .ModelTot.BackColor = CLR_DISABLED_BG
680       .CurrentTot.ForeColor = CLR_DISABLED_FG
690       .CurrentTot.BackColor = CLR_DISABLED_BG
700       .DifferenceTot.ForeColor = CLR_DISABLED_FG
710       .DifferenceTot.BackColor = CLR_DISABLED_BG
720       .ICash.ForeColor = CLR_DISABLED_FG
730       .ICash.BackColor = CLR_DISABLED_BG
740       DoCmd.Hourglass False
750     End With

EXITP:
760     Exit Sub

ERRH:
770     DoCmd.Hourglass False
780     Select Case ERR.Number
        Case Else
790       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
800     End Select
810     Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionslity:
        ' **   Model:            M {Model_input on frmPortfolioModeling_Sub}
        ' **   Account Select:   A {opgReturn_optSelect}
        ' **   Asset Menu:       S {opgReturn_optMenu}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmPortfolioModeling_Sub}

        ' ** Plain keys.
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyEscape
970         intRetVal = 0
980         cmdClose_Click  ' ** Procedure: Below.
990       Case vbKeyUp
1000        With Me
1010          intRetVal = 0
1020          .frmPortfolioModeling_Sub.SetFocus
1030          .frmPortfolioModeling_Sub.Form.MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmPortfolioModeling_Sub.
1040        End With
1050      Case vbKeyDown
1060        With Me
1070          intRetVal = 0
1080          .frmPortfolioModeling_Sub.SetFocus
1090          .frmPortfolioModeling_Sub.Form.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmPortfolioModeling_Sub.
1100        End With
1110      Case vbKeyF5
1120        With Me
1130          intRetVal = 0
1140          .frmPortfolioModeling_Sub.Form.RecalcTots  ' ** Form Procedure: frmPortfolioModeling_Sub.
1150        End With
1160      End Select
1170    End If

        ' ** Alt keys.
1180    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1190      Select Case intRetVal
          Case vbKeyA
1200        With Me
1210          intRetVal = 0
1220          .opgReturn.SetFocus
1230          .opgReturn = .opgReturn_optSelect.OptionValue
1240          opgReturn_AfterUpdate  ' ** Procedure: Below.
1250        End With
1260      Case vbKeyM
1270        With Me
1280          intRetVal = 0
1290          .frmPortfolioModeling_Sub.SetFocus
1300          .frmPortfolioModeling_Sub.Form.Model_input.SetFocus
1310        End With
1320      Case vbKeyS
1330        With Me
1340          intRetVal = 0
1350          .opgReturn.SetFocus
1360          .opgReturn = .opgReturn_optMenu.OptionValue
1370          opgReturn_AfterUpdate  ' ** Procedure: Below.
1380        End With
1390      Case vbKeyX
1400        intRetVal = 0
1410        cmdClose_Click  ' ** Procedure: Below.
1420      End Select
1430    End If

        ' ** Ctrl keys.
1440    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1450      Select Case intRetVal
          Case vbKeyS
1460        With Me
1470          intRetVal = 0
1480          .frmPortfolioModeling_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmPortfolioModeling_Sub.
1490        End With
1500      Case vbKeyPageUp, vbKeyHome
1510        With Me
1520          intRetVal = 0
1530          .frmPortfolioModeling_Sub.SetFocus
1540          .frmPortfolioModeling_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmPortfolioModeling_Sub.
1550        End With
1560      Case vbKeyPageDown, vbKeyEnd
1570        With Me
1580          intRetVal = 0
1590          .frmPortfolioModeling_Sub.SetFocus
1600          .frmPortfolioModeling_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmPortfolioModeling_Sub.
1610        End With
1620      End Select
1630    End If

        ' ** Ctrl-Shift keys.
1640    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1650      Select Case intRetVal
          Case vbKeyF
1660        With Me
1670          intRetVal = 0
1680          .FocusHolder.SetFocus
1690        End With
1700      End Select
1710    End If

EXITP:
1720    KeyCode = intRetVal
1730    Exit Sub

ERRH:
1740    intRetVal = 0
1750    Select Case ERR.Number
        Case Else
1760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1770    End Select
1780    Resume EXITP

End Sub

Private Sub Form_Timer()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1810    With Me
1820      .TimerInterval = 0&
1830      If gblnSetFocus = True Then
1840        gblnSetFocus = False
1850        strFormName = vbNullString
1860  On Error Resume Next
1870        Do Until strFormName = THIS_NAME
1880          DoCmd.SelectObject acForm, THIS_NAME, False
1890          strFormName = Screen.ActiveForm.Name
1900        Loop
1910  On Error GoTo ERRH
1920      End If  ' ** gblnSetFocus.
1930      .frmPortfolioModeling_Sub.Form.RecalcTots  ' ** Form Procedure: frmPortfolioModeling_Sub.
1940      DoEvents
1950      If blnRecalcTotsNotReady = True Then
1960        .TimerInterval = 250&
1970      End If
1980      If gblnGoToReport = True Then
1990        DoCmd.Hourglass True  ' ** Make sure it's still running.
2000        DoEvents
2010        Select Case blnGoingToReport
            Case True
2020          blnGoingToReport = False
2030          gblnGoToReport = False
2040          .GoToReport_arw_pr_img.Visible = False
2050          .cmdClose.SetFocus
2060          DoCmd.Hourglass False
2070        Case False
2080          blnGoingToReport = True
2090          .GoToReport_arw_pr_img.Visible = True
2100          If .cmdPrintReport.Enabled = True Then
2110            .cmdPrintReport.SetFocus
2120          End If
2130          Beep
2140          DoEvents
2150          .TimerInterval = (GTR_WAIT * 2&)
2160        End Select
2170      End If
2180    End With

EXITP:
2190    Exit Sub

ERRH:
2200    Select Case ERR.Number
        Case Else
2210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
2220    End Select
2230    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

2310    With Me

2320      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2330      If blnViewOnly = False Then
2340        Set dbs = CurrentDb
2350        With dbs
              ' ** Update qryPortfolioModeling_09_01 (PortfolioModel,
              ' ** linked to tblPortfolioModeling2, with Model_new).
2360          Set qdf = .QueryDefs("qryPortfolioModeling_09_02")
2370          qdf.Execute
2380          Set qdf = Nothing
2390          .Close
2400        End With
2410        Set dbs = Nothing
2420      End If

2430      gblnGoToReport = False
2440      gblnSetFocus = True
2450      gstrReturningForm = THIS_NAME

2460      Select Case .opgReturn
          Case .opgReturn_optSelect.OptionValue
2470        strDocName = "frmPortfolioModeling_Select"
2480        DoCmd.OpenForm strDocName, , , , , , "frmMenu_Asset"
2490      Case .opgReturn_optMenu.OptionValue
2500        strDocName = "frmMenu_Asset"
2510        DoCmd.OpenForm strDocName, , , , , , "frmMenu_Main"
2520      End Select

2530    End With

EXITP:
2540    Set qdf = Nothing
2550    Set dbs = Nothing
2560    Exit Sub

ERRH:
2570    Select Case ERR.Number
        Case Else
2580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
2590    End Select
2600    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2710    With Me
2720      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2730        Select Case blnPrintReport_Focus
            Case True
2740          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2750          .cmdPrintReport_raised_img.Visible = False
2760        Case False
2770          .cmdPrintReport_raised_img.Visible = True
2780          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2790        End Select
2800        .cmdPrintReport_raised_focus_img.Visible = False
2810        .cmdPrintReport_raised_focus_dots_img.Visible = False
2820        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2830        .cmdPrintReport_raised_img_dis.Visible = False
2840      End If
2850    End With

EXITP:
2860    Exit Sub

ERRH:
2870    Select Case ERR.Number
        Case Else
2880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2890    End Select
2900    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Public Sub cmdClose_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3110    DoCmd.Close acForm, THIS_NAME

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
3270        With Me
3280          intRetVal = 0
3290          .cmbInvestmentObjectives.SetFocus
3300        End With
3310      End Select
3320    End If

        ' ** Shift keys.
3330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3340      Select Case intRetVal
          Case vbKeyTab
3350        With Me
3360          intRetVal = 0
3370          If .cmdPrintReport.Enabled = True Then
3380            .cmdPrintReport.SetFocus
3390          Else
3400            .opgReturn.SetFocus
3410          End If
3420        End With
3430      End Select
3440    End If

EXITP:
3450    KeyCode = intRetVal
3460    Exit Sub

ERRH:
3470    intRetVal = 0
3480    Select Case ERR.Number
        Case Else
3490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3500    End Select
3510    Resume EXITP

End Sub

Private Sub cmbInvestmentObjectives_Enter()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbInvestmentObjectives_Enter"

3610    With Me
3620      If IsNull(.cmbInvestmentObjectives) = True Then
3630        ExpandCombo .cmbInvestmentObjectives  ' ** Module Procedure: modUtilities.
3640      End If
3650    End With

EXITP:
3660    Exit Sub

ERRH:
3670    Select Case ERR.Number
        Case Else
3680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3690    End Select
3700    Resume EXITP

End Sub

Private Sub cmbInvestmentObjectives_AfterUpdate()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbInvestmentObjectives_AfterUpdate"

3810    With Me
3820      lngRecsCur = .frmPortfolioModeling_Sub.Form.RecCnt  ' ** Form Function: frmPortfolioModeling_Sub.
3830      If lngRecsCur > 0& Then
3840        If IsNull(.cmbInvestmentObjectives) = False Then
3850          .frmPortfolioModeling_Sub.Form.investmentobj = .cmbInvestmentObjectives
3860        Else
3870          .frmPortfolioModeling_Sub.Form.investmentobj = Null
3880        End If
3890        .frmPortfolioModeling_Sub.SetFocus
3900        .frmPortfolioModeling_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmPortfolioModeling_Sub.
3910        DoEvents
3920        .frmPortfolioModeling_Sub.Form.Model_input.SetFocus
3930        .frmPortfolioModeling_Sub.Form.Model_input.SelLength = 0
3940        .frmPortfolioModeling_Sub.Form.Model_input.SelStart = 0
3950      End If
3960    End With

EXITP:
3970    Exit Sub

ERRH:
3980    Select Case ERR.Number
        Case Else
3990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4000    End Select
4010    Resume EXITP

End Sub

Private Sub cmbInvestmentObjectives_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbInvestmentObjectives_KeyDown"

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
          Case vbKeyTab
4170        With Me
4180          intRetVal = 0
4190          lngRecsCur = .frmPortfolioModeling_Sub.Form.RecCnt  ' ** Form function: frmPortfolioModeling_Sub.
4200          If lngRecsCur > 0& Then
4210            .frmPortfolioModeling_Sub.SetFocus
4220            .frmPortfolioModeling_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmPortfolioModeling_Sub.
4230          Else
4240            .opgReturn.SetFocus
4250          End If
4260        End With
4270      End Select
4280    End If

        ' ** Shift keys.
4290    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4300      Select Case intRetVal
          Case vbKeyTab
4310        With Me
4320          intRetVal = 0
4330          .cmdClose.SetFocus
4340        End With
4350      End Select
4360    End If

EXITP:
4370    KeyCode = intRetVal
4380    Exit Sub

ERRH:
4390    intRetVal = 0
4400    Select Case ERR.Number
        Case Else
4410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4420    End Select
4430    Resume EXITP

End Sub

Private Sub frmPortfolioModeling_Sub_Enter()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "frmPortfolioModeling_Sub_Enter"

4510    With Me
4520      .frmPortfolioModeling_Sub.Form.FromEnter_Set True  ' ** Form Procedure: frmPortfolioModeling_Sub.
4530    End With

EXITP:
4540    Exit Sub

ERRH:
4550    Select Case ERR.Number
        Case Else
4560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4570    End Select
4580    Resume EXITP

End Sub

Public Sub opgReturn_AfterUpdate()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgReturn_AfterUpdate"

4610    With Me
4620      Select Case .opgReturn
          Case .opgReturn_optSelect.OptionValue
4630        .opgReturn_optSelect_lbl.FontBold = True
4640        .opgReturn_optSelect_lbl2.FontBold = True
4650        .opgReturn_optMenu_lbl.FontBold = False
4660        .opgReturn_optMenu_lbl2.FontBold = False
4670      Case .opgReturn_optMenu.OptionValue
4680        .opgReturn_optSelect_lbl.FontBold = False
4690        .opgReturn_optSelect_lbl2.FontBold = False
4700        .opgReturn_optMenu_lbl.FontBold = True
4710        .opgReturn_optMenu_lbl2.FontBold = True
4720      End Select
4730    End With

EXITP:
4740    Exit Sub

ERRH:
4750    DoCmd.Hourglass False
4760    Select Case ERR.Number
        Case Else
4770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4780    End Select
4790    Resume EXITP

End Sub

Private Sub opgReturn_optSelect_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgReturn_optSelect_KeyDown"

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
          Case vbKeyTab
4870        With Me
4880          intRetVal = 0
4890          If .cmdPrintReport.Enabled = True Then
4900            .cmdPrintReport.SetFocus
4910          Else
4920            .cmdClose.SetFocus
4930          End If
4940        End With
4950      End Select
4960    End If

        ' ** Shift keys.
4970    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4980      Select Case intRetVal
          Case vbKeyTab
4990        With Me
5000          intRetVal = 0
5010          lngRecsCur = .frmPortfolioModeling_Sub.Form.RecCnt  ' ** Form Function: frmPortfolioModeling_Sub.
5020          If lngRecsCur > 0& Then
5030            .frmPortfolioModeling_Sub.SetFocus
5040            .frmPortfolioModeling_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmPortfolioModeling_Sub.
5050          Else
5060            .cmbInvestmentObjectives.SetFocus
5070          End If
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

Private Sub opgReturn_optMenu_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgReturn_optMenu_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5220    intShiftDown = (Shift And acShiftMask) > 0
5230    intAltDown = (Shift And acAltMask) > 0
5240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5260      Select Case intRetVal
          Case vbKeyTab
5270        With Me
5280          intRetVal = 0
5290          If .cmdPrintReport.Enabled = True Then
5300            .cmdPrintReport.SetFocus
5310          Else
5320            .cmdClose.SetFocus
5330          End If
5340        End With
5350      End Select
5360    End If

        ' ** Shift keys.
5370    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5380      Select Case intRetVal
          Case vbKeyTab
5390        With Me
5400          intRetVal = 0
5410          lngRecsCur = .frmPortfolioModeling_Sub.Form.RecCnt  ' ** Form Function: frmPortfolioModeling_Sub.
5420          If lngRecsCur > 0& Then
5430            .frmPortfolioModeling_Sub.SetFocus
5440            .frmPortfolioModeling_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmPortfolioModeling_Sub.
5450          Else
5460            .cmbInvestmentObjectives.SetFocus
5470          End If
5480        End With
5490      End Select
5500    End If

EXITP:
5510    KeyCode = intRetVal
5520    Exit Sub

ERRH:
5530    intRetVal = 0
5540    Select Case ERR.Number
        Case Else
5550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5560    End Select
5570    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

5610    With Me
5620      blnPrintReport_Focus = True
5630      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
5640      .cmdPrintReport_raised_img.Visible = False
5650      .cmdPrintReport_raised_focus_img.Visible = False
5660      .cmdPrintReport_raised_focus_dots_img.Visible = False
5670      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5680      .cmdPrintReport_raised_img_dis.Visible = False
5690    End With

EXITP:
5700    Exit Sub

ERRH:
5710    Select Case ERR.Number
        Case Else
5720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5730    End Select
5740    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

5810    With Me
5820      blnPrintReport_MouseDown = True
5830      .cmdPrintReport_sunken_focus_dots_img.Visible = True
5840      .cmdPrintReport_raised_img.Visible = False
5850      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5860      .cmdPrintReport_raised_focus_img.Visible = False
5870      .cmdPrintReport_raised_focus_dots_img.Visible = False
5880      .cmdPrintReport_raised_img_dis.Visible = False
5890    End With

EXITP:
5900    Exit Sub

ERRH:
5910    Select Case ERR.Number
        Case Else
5920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5930    End Select
5940    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

6000  On Error GoTo ERRH

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String
        Dim strSQL As String

        Const THIS_PROC As String = "cmdPrintReport_Click"

6010    With Me

6020      strSQL = "SELECT tblPortfolioModeling2.assettype, tblPortfolioModeling2.assettype_description, tblPortfolioModeling2.accountno, " & _
            "tblPortfolioModeling2.shortname, tblPortfolioModeling2.Model, tblPortfolioModeling2.Current, " & _
            "tblPortfolioModeling2.Difference," & CoInfo & ", TotalCurrent, cash " & _
            "FROM tblPortfolioModeling2 " & _
            "WHERE (tblPortfolioModeling2.Model <> 0) OR (tblPortfolioModeling2.Current <> 0) OR (tblPortfolioModeling2.Difference <> 0);"

6030      Set dbs = CurrentDb
6040      Set rst = dbs.OpenRecordset(strSQL)
6050      If rst.EOF Then
6060        MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
6070        rst.Close
6080        dbs.Close
6090      Else
6100        rst.Close
6110        Set qdf = dbs.QueryDefs("qryRpt_PortfolioModeling_01")
6120        qdf.SQL = strSQL
6130        dbs.Close
6140        strDocName = "rptPortfolioModeling"
6150        If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
6160          DoCmd.OpenReport strDocName, acViewPreview
6170          DoCmd.Maximize
6180          DoCmd.RunCommand acCmdFitToWindow
6190          .Visible = False
6200        Else
              '##GTR_Ref: rptPortfolioModeling
6210          DoCmd.OpenReport strDocName, acViewNormal
6220        End If
6230      End If

6240    End With

EXITP:
6250    Set rst = Nothing
6260    Set qdf = Nothing
6270    Set dbs = Nothing
6280    Exit Sub

ERRH:
6290    Me.Visible = True
6300    DoCmd.Restore
6310    Select Case ERR.Number
        Case Else
6320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
6330    End Select
6340    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
6490          .cmdClose.SetFocus
6500        End With
6510      End Select
6520    End If

        ' ** Shift keys.
6530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6540      Select Case intRetVal
          Case vbKeyTab
6550        With Me
6560          intRetVal = 0
6570          .opgReturn.SetFocus
6580        End With
6590      End Select
6600    End If

EXITP:
6610    KeyCode = intRetVal
6620    Exit Sub

ERRH:
6630    intRetVal = 0
6640    Select Case ERR.Number
        Case Else
6650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6660    End Select
6670    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

6710    With Me
6720      If blnPrintReport_MouseDown = False Then
6730        Select Case blnPrintReport_Focus
            Case True
6740          .cmdPrintReport_raised_focus_dots_img.Visible = True
6750          .cmdPrintReport_raised_focus_img.Visible = False
6760        Case False
6770          .cmdPrintReport_raised_focus_img.Visible = True
6780          .cmdPrintReport_raised_focus_dots_img.Visible = False
6790        End Select
6800        .cmdPrintReport_raised_img.Visible = False
6810        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6820        .cmdPrintReport_sunken_focus_dots_img.Visible = False
6830        .cmdPrintReport_raised_img_dis.Visible = False
6840      End If
6850    End With

EXITP:
6860    Exit Sub

ERRH:
6870    Select Case ERR.Number
        Case Else
6880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6890    End Select
6900    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

7010    With Me
7020      .cmdPrintReport_raised_focus_dots_img.Visible = True
7030      .cmdPrintReport_raised_img.Visible = False
7040      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7050      .cmdPrintReport_raised_focus_img.Visible = False
7060      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7070      .cmdPrintReport_raised_img_dis.Visible = False
7080      blnPrintReport_MouseDown = False
7090    End With

EXITP:
7100    Exit Sub

ERRH:
7110    Select Case ERR.Number
        Case Else
7120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7130    End Select
7140    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

7210    With Me
7220      .cmdPrintReport_raised_img.Visible = True
7230      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7240      .cmdPrintReport_raised_focus_img.Visible = False
7250      .cmdPrintReport_raised_focus_dots_img.Visible = False
7260      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7270      .cmdPrintReport_raised_img_dis.Visible = False
7280      blnPrintReport_Focus = False
7290    End With

EXITP:
7300    Exit Sub

ERRH:
7310    Select Case ERR.Number
        Case Else
7320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7330    End Select
7340    Resume EXITP

End Sub

Public Sub RecalcTots_NotReady(blnReady As Boolean)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots_NotReady"

7410    blnRecalcTotsNotReady = blnReady

EXITP:
7420    Exit Sub

ERRH:
7430    Select Case ERR.Number
        Case Else
7440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7450    End Select
7460    Resume EXITP

End Sub
