VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmInvestmentObjective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmInvestmentObjective"

'VGC 09/23/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Show Accounts:    H {chkShowAccounts}
' **   Objective:        O {invobj_name, cmbInvestmentObjectives on frmInvestmentObjective_Sub}
' **   Print Report:     P {cmdPrintReport}
' **   Show Unassigned:  U {chkShowUnassigned}
' **   Exit:             X {cmdClose}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:          F5 {RecalcTots on frmInvestmentObjective_Sub}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmInvestmentObjective_Sub}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, lngPrintReportLbl_Left As Long, blnWasGTR As Boolean
Private lngUnassigned As Long, lngUnassignedLbl_Left As Long, lngTpp As Long
Private strSortNow1 As String, strSortNow2 As String, strSortNow3 As String
Private strCallingForm As String, lngRecsCur As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean
        Dim varTmp00 As Variant

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnWasGTR = gblnGoToReport
160           strSortNow1 = vbNullString: strSortNow3 = vbNullString: strSortNow3 = vbNullString

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
180           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

190           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left
200           lngUnassignedLbl_Left = .Unassigned_lbl.Left

210           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

220           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

230           varTmp00 = DLookup("[cnt]", "qryInvestmentObjective_05")
240           If IsNull(varTmp00) = True Then
250             lngUnassigned = 0&
260           Else
270             lngUnassigned = CLng(varTmp00)
280           End If
290           If lngUnassigned > 0& Then
300             If lngUnassigned = 1& Then
310               .Unassigned_lbl.Caption = "1 Account Unassigned"
320             Else
330               .Unassigned_lbl.Caption = CStr(lngUnassigned) & " Accounts Unassigned"
340             End If
350             .Unassigned_lbl.Visible = True
360           End If

370           If lngUnassigned = 0& Then
380             .chkShowUnassigned = False
390             .chkShowUnassigned.Enabled = False
400           Else
410             .chkShowUnassigned.Enabled = True
420           End If

430           chkShowAccounts_AfterUpdate  ' ** Procedure: Below.
440           chkShowUnassigned_AfterUpdate  ' ** Procedure: Below.

450         Else
460           Cancel = -1
470         End If
480       Else
490         Cancel = -1
500       End If
510     End With

520     If Cancel = -1 Then
530       Beep
540       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
550     End If

EXITP:
560     Exit Sub

ERRH:
570     Select Case ERR.Number
        Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim frm As Access.Form
        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc:          F5 {RecalcTots on frmInvestmentObjective_Sub}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Objective:        O {invobj_name, cmbInvestmentObjectives on frmInvestmentObjective_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmInvestmentObjective_Sub}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyF5
770         With Me
780           intRetVal = 0
790           .frmInvestmentObjective_Sub.Form.RecalcTots  ' ** Form Procedure: frmInvestmentObjective_Sub.
800         End With
810       Case vbKeyEscape
820         intRetVal = 0
830         cmdClose_Click  ' ** Procedure: Below.
840       End Select
850     End If

        ' ** Alt keys.
860     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
870       Select Case intRetVal
          Case vbKeyO
880         With Me
890           intRetVal = 0
900           Set frm = .frmInvestmentObjective_Sub.Form
910           .frmInvestmentObjective_Sub.SetFocus
920           If frm.invobj_name.Visible = True And frm.invobj_name.Enabled = True Then
930             frm.invobj_name.SetFocus
940           ElseIf frm.cmbInvestmentObjectives.Visible = True And frm.cmbInvestmentObjectives.Enabled = True Then
950             frm.cmbInvestmentObjectives.SetFocus
960           End If
970         End With
980       Case vbKeyX
990         intRetVal = 0
1000        cmdClose_Click  ' ** Procedure: Below.
1010      End Select
1020    End If

        ' ** Ctrl keys.
1030    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1040      Select Case intRetVal
          Case vbKeyS
1050        With Me
1060          intRetVal = 0
1070          If .chkShowUnassigned = True Then
1080            .frmInvestmentObjective_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmInvestmentObjective_Sub.
1090          End If
1100        End With
1110      End Select
1120    End If

        ' ** Ctrl-Shift keys.
1130    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1140      Select Case intRetVal
          Case vbKeyF
1150        With Me
1160          intRetVal = 0
1170          .FocusHolder.SetFocus
1180        End With
1190      End Select
1200    End If

EXITP:
1210    Set frm = Nothing
1220    KeyCode = intRetVal
1230    Exit Sub

ERRH:
1240    intRetVal = 0
1250    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1260    Case Else
1270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1280    End Select
1290    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1310    With Me
1320      .TimerInterval = 0&
1330      If gblnSetFocus = True Then
1340        gblnSetFocus = False
1350        strFormName = vbNullString
1360  On Error Resume Next
1370        Do Until strFormName = THIS_NAME
1380          DoCmd.SelectObject acForm, THIS_NAME, False
1390          strFormName = Screen.ActiveForm.Name
1400        Loop
1410  On Error GoTo ERRH
1420        .frmInvestmentObjective_Sub.SetFocus
1430      End If
1440      If gblnGoToReport = True Then
1450        Select Case blnGoingToReport
            Case True
1460          Select Case blnGoingToReport2
              Case True
1470            DoCmd.Hourglass False  ' ** It should already be off.
1480            gblnGoToReport = False
1490            blnGoingToReport = False
1500            .GoToReport_arw_pr_img.Visible = False
1510            .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1520            DoEvents
1530          Case False
1540            Select Case garr_varGoToReport(GTR_RNAM)
                Case "rptListOfInvestmentObjectives_Detailed", "rptListOfInvestmentObjectives_Unassigned"
1550              blnGoingToReport2 = True
1560              .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1570              .GoToReport_arw_det_img.Visible = False
1580              .GoToReport_arw_pr_img.Visible = True
1590              .Unassigned_lbl.Left = lngUnassignedLbl_Left
1600              .cmdPrintReport.SetFocus
1610              Beep
1620              DoCmd.Hourglass False
1630              DoEvents
1640              .TimerInterval = (GTR_WAIT * 2&)
1650            Case Else
1660              DoCmd.Hourglass False  ' ** It should already be off.
1670              gblnGoToReport = False
1680              blnGoingToReport = False
1690              .GoToReport_arw_pr_img.Visible = False
1700              .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1710              DoEvents
1720            End Select
1730          End Select
1740        Case False
1750          DoCmd.Hourglass True  ' ** Make sure it's still running.
1760          DoEvents
1770          blnGoingToReport = True
1780          Select Case garr_varGoToReport(GTR_RNAM)
              Case "rptListOfInvestmentObjectives"
1790            .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1800            .GoToReport_arw_pr_img.Visible = True
1810            .cmdPrintReport.SetFocus
1820            DoEvents
1830            If .chkShowAccounts = True Then
1840              .chkShowAccounts = False
1850              chkShowAccounts_AfterUpdate  ' ** Procedure: Below.
1860              DoEvents
1870            End If
1880            If .chkShowUnassigned = True Then
1890              .chkShowUnassigned = False
1900              chkShowUnassigned_AfterUpdate  ' ** Procedure: Below.
1910              DoEvents
1920            End If
1930            Beep
1940            DoCmd.Hourglass False
1950            DoEvents
1960            .TimerInterval = (GTR_WAIT * 2&)
1970          Case "rptListOfInvestmentObjectives_Detailed"
1980            .GoToReport_arw_det_img.Top = (.chkShowAccounts_lbl.Top - lngTpp)
1990            .GoToReport_arw_det_img.Visible = True
2000            DoEvents
2010            If .chkShowUnassigned = True Then
2020              .chkShowUnassigned = False
2030              chkShowUnassigned_AfterUpdate  ' ** Procedure: Below.
2040              DoEvents
2050            End If
2060            If .chkShowAccounts = False Then
2070              .chkShowAccounts = True
2080              chkShowAccounts_AfterUpdate  ' ** Procedure: Below.
2090            End If
2100            DoEvents
2110            .TimerInterval = GTR_WAIT
2120          Case "rptListOfInvestmentObjectives_Unassigned"
2130            .Unassigned_lbl.Left = (.Unassigned_lbl.Left + .GoToReport_arw_det_img.Width)
2140            .GoToReport_arw_det_img.Top = (.chkShowUnassigned_lbl.Top - lngTpp)
2150            .GoToReport_arw_det_img.Visible = True
2160            DoEvents
2170            If .chkShowUnassigned = False Then
2180              .chkShowUnassigned = True
2190              chkShowUnassigned_AfterUpdate  ' ** Procedure: Below.
2200              DoEvents
2210            End If
2220            If .chkShowAccounts = True Then
2230              .chkShowAccounts = False
2240              chkShowAccounts_AfterUpdate  ' ** Procedure: Below.
2250              DoEvents
2260            End If
2270            DoEvents
2280            .TimerInterval = GTR_WAIT
2290          End Select
2300        End Select
2310      Else
2320        DoCmd.Hourglass False
2330      End If
2340    End With

EXITP:
2350    Exit Sub

ERRH:
2360    Select Case ERR.Number
        Case Else
2370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2380    End Select
2390    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

2410    With Me

2420      DoCmd.Hourglass True
2430      DoEvents

2440      varTmp00 = DCount("*", "tblInvestmentObjective_Staging")
2450      If varTmp00 > 0 Then
2460        Set dbs = CurrentDb
            ' ** Update qryInvestmentObjective_10_01 (Account, linked to tblInvestmentObjective_Staging,
            ' ** just invobj_name <> '{Unassigned}, with investmentobj_new).
2470        Set qdf = dbs.QueryDefs("qryInvestmentObjective_10_02")
2480        qdf.Execute
2490        Set qdf = Nothing
            ' ** Empty tblInvestmentObjective_Staging.
2500        Set qdf = dbs.QueryDefs("qryInvestmentObjective_09_02")
2510        qdf.Execute
2520        Set qdf = Nothing
2530        dbs.Close
2540        Set dbs = Nothing
2550      End If

2560      Select Case blnWasGTR
          Case True
2570        blnWasGTR = False
2580      Case False
            ' ** Don't save settings imposed by GoToReport!
2590        Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
2600      End Select

2610      blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

2620      gblnSetFocus = False
2630      gstrReturningForm = THIS_NAME
2640      Select Case strCallingForm
          Case "frmMenu_SystemTable"
2650        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Maintenance"
2660      End Select

2670    End With

EXITP:
2680    Set qdf = Nothing
2690    Set dbs = Nothing
2700    Exit Sub

ERRH:
2710    DoCmd.Hourglass False
2720    Select Case ERR.Number
        Case Else
2730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2740    End Select
2750    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2810    With Me
2820      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2830        Select Case blnPrintReport_Focus
            Case True
2840          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2850          .cmdPrintReport_raised_img.Visible = False
2860        Case False
2870          .cmdPrintReport_raised_img.Visible = True
2880          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2890        End Select
2900        .cmdPrintReport_raised_focus_img.Visible = False
2910        .cmdPrintReport_raised_focus_dots_img.Visible = False
2920        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2930        .cmdPrintReport_raised_img_dis.Visible = False
2940      End If
2950    End With

EXITP:
2960    Exit Sub

ERRH:
2970    Select Case ERR.Number
        Case Else
2980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2990    End Select
3000    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Public Sub cmdClose_Click()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3210    DoCmd.Close acForm, THIS_NAME

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
3370        With Me
3380          intRetVal = 0
3390          .frmInvestmentObjective_Sub.SetFocus
3400          .frmInvestmentObjective_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmInvestmentObjective_Sub.
3410          .frmInvestmentObjective_Sub.Form.FocusHolder.SetFocus
3420        End With
3430      End Select
3440    End If

        ' ** Shift keys.
3450    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3460      Select Case intRetVal
          Case vbKeyTab
3470        With Me
3480          intRetVal = 0
3490          If .cmdPrintReport.Enabled = True Then
3500            .cmdPrintReport.SetFocus
3510          ElseIf .chkShowUnassigned.Enabled = True Then
3520            .chkShowUnassigned.SetFocus
3530          Else
3540            .chkShowAccounts.SetFocus
3550          End If
3560        End With
3570      End Select
3580    End If

EXITP:
3590    KeyCode = intRetVal
3600    Exit Sub

ERRH:
3610    intRetVal = 0
3620    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3630    Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

3710    With Me
3720      blnPrintReport_Focus = True
3730      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3740      .cmdPrintReport_raised_img.Visible = False
3750      .cmdPrintReport_raised_focus_img.Visible = False
3760      .cmdPrintReport_raised_focus_dots_img.Visible = False
3770      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3780      .cmdPrintReport_raised_img_dis.Visible = False
3790    End With

EXITP:
3800    Exit Sub

ERRH:
3810    Select Case ERR.Number
        Case Else
3820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3830    End Select
3840    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

3910    With Me
3920      blnPrintReport_MouseDown = True
3930      .cmdPrintReport_sunken_focus_dots_img.Visible = True
3940      .cmdPrintReport_raised_img.Visible = False
3950      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3960      .cmdPrintReport_raised_focus_img.Visible = False
3970      .cmdPrintReport_raised_focus_dots_img.Visible = False
3980      .cmdPrintReport_raised_img_dis.Visible = False
3990    End With

EXITP:
4000    Exit Sub

ERRH:
4010    Select Case ERR.Number
        Case Else
4020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4030    End Select
4040    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String, strSortNow As String
        Dim blnContinue As Boolean

4110    With Me

4120      blnContinue = True

4130      DoCmd.Hourglass True
4140      DoEvents

4150      Select Case .chkShowUnassigned
          Case True
4160        strDocName = "rptListOfInvestmentObjectives_Unassigned"
4170      Case False
4180        Select Case .chkShowAccounts
            Case True
4190          strSortNow = .frmInvestmentObjective_Sub.Form.SortNow_Get  ' ** Form Function: frmInvestmentObjective_Sub.
4200          If strSortNow = "[invobj_name], [accountno]" Then
4210            strDocName = "rptListOfInvestmentObjectives_Detailed"
4220          Else
4230            strDocName = "rptListOfInvestmentObjectives_Detailed_Sort"
4240          End If
4250        Case False
4260          strDocName = "rptListOfInvestmentObjectives"
4270        End Select
4280      End Select

4290    End With

4300    If blnContinue = True Then
4310      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4320        DoCmd.OpenReport strDocName, acViewPreview
4330        DoCmd.Maximize
4340        DoCmd.RunCommand acCmdFitToWindow
4350        Me.Visible = False
4360      Else
            '##GTR_Ref: rptListOfInvestmentObjectives
            '##GTR_Ref: rptListOfInvestmentObjectives_Detailed
            '##GTR_Ref: rptListOfInvestmentObjectives_Unassigned
4370        DoCmd.OpenReport strDocName, acViewNormal
4380      End If
4390    End If

4400    DoCmd.Hourglass False

EXITP:
4410    Exit Sub

ERRH:
4420    DoCmd.Hourglass False
4430    Me.Visible = True
4440    DoCmd.Restore
4450    Select Case ERR.Number
        Case Else
4460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4470    End Select
4480    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4520    intShiftDown = (Shift And acShiftMask) > 0
4530    intAltDown = (Shift And acAltMask) > 0
4540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab
4570        With Me
4580          intRetVal = 0
4590          .cmdClose.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab
4650        With Me
4660          intRetVal = 0
4670          If .chkShowUnassigned.Enabled = True Then
4680            .chkShowUnassigned.SetFocus
4690          ElseIf .chkShowAccounts.Enabled = True Then
4700            .chkShowAccounts.SetFocus
4710          Else
4720            .frmInvestmentObjective_Sub.SetFocus
4730            .frmInvestmentObjective_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmInvestmentObjective_Sub.
4740            .frmInvestmentObjective_Sub.Form.FocusHolder.SetFocus
4750          End If
4760        End With
4770      End Select
4780    End If

EXITP:
4790    KeyCode = intRetVal
4800    Exit Sub

ERRH:
4810    intRetVal = 0
4820    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4830    Case Else
4840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4850    End Select
4860    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

4910    With Me
4920      If blnPrintReport_MouseDown = False Then
4930        Select Case blnPrintReport_Focus
            Case True
4940          .cmdPrintReport_raised_focus_dots_img.Visible = True
4950          .cmdPrintReport_raised_focus_img.Visible = False
4960        Case False
4970          .cmdPrintReport_raised_focus_img.Visible = True
4980          .cmdPrintReport_raised_focus_dots_img.Visible = False
4990        End Select
5000        .cmdPrintReport_raised_img.Visible = False
5010        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5020        .cmdPrintReport_sunken_focus_dots_img.Visible = False
5030        .cmdPrintReport_raised_img_dis.Visible = False
5040      End If
5050    End With

EXITP:
5060    Exit Sub

ERRH:
5070    Select Case ERR.Number
        Case Else
5080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5090    End Select
5100    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

5210    With Me
5220      .cmdPrintReport_raised_focus_dots_img.Visible = True
5230      .cmdPrintReport_raised_img.Visible = False
5240      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5250      .cmdPrintReport_raised_focus_img.Visible = False
5260      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5270      .cmdPrintReport_raised_img_dis.Visible = False
5280      blnPrintReport_MouseDown = False
5290    End With

EXITP:
5300    Exit Sub

ERRH:
5310    Select Case ERR.Number
        Case Else
5320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5330    End Select
5340    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

5410    With Me
5420      .cmdPrintReport_raised_img.Visible = True
5430      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5440      .cmdPrintReport_raised_focus_img.Visible = False
5450      .cmdPrintReport_raised_focus_dots_img.Visible = False
5460      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5470      .cmdPrintReport_raised_img_dis.Visible = False
5480      blnPrintReport_Focus = False
5490    End With

EXITP:
5500    Exit Sub

ERRH:
5510    Select Case ERR.Number
        Case Else
5520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5530    End Select
5540    Resume EXITP

End Sub

Private Sub Unassigned_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "Unassigned_lbl_MouseMove"

5610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5620    Exit Sub

ERRH:
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Public Sub chkShowAccounts_AfterUpdate()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAccounts_AfterUpdate"

        Dim frm As Access.Form
        Dim strSortNow As String, blnDesc As Boolean

5710    With Me
5720      Set frm = .frmInvestmentObjective_Sub.Form
5730      lngRecsCur = frm.RecCnt  ' ** Form Function: frmInvestmentObjective_Sub.
5740      Select Case .chkShowAccounts
          Case True
5750        .chkShowAccounts_lbl.FontBold = True
5760        frm.ShowAccounts True  ' ** Form Procedure: frmInvestmentObjective_Sub.
5770        DoEvents
5780        strSortNow = frm.SortNow_Get  ' ** Form Function: frmInvestmentObjective_Sub.
5790        If strSortNow2 <> vbNullString Then
              ' ** Use the last sort this option used.
5800          If InStr(strSortNow2, " DESC") > 0 Then
5810            blnDesc = True
5820          Else
5830            blnDesc = False
5840          End If
5850          If Left(strSortNow2, 11) = "[invobj_id]" Then
5860            frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
5870            If blnDesc = True Then
5880              frm.SortNow "invobj_id_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
5890            End If
5900          ElseIf Left(strSortNow2, 13) = "[invobj_name]" Then
5910            frm.SortNow "invobj_name_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
5920            If blnDesc = True Then
5930              frm.SortNow "invobj_name_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
5940            End If
5950          ElseIf Left(strSortNow2, 11) = "[alphasort]" Then
5960            frm.SortNow "accountno_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
5970            If blnDesc = True Then
5980              frm.SortNow "accountno_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
5990            End If
6000          ElseIf Left(strSortNow2, 11) = "[shortname]" Then
6010            frm.SortNow "shortname_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
6020            If blnDesc = True Then
6030              frm.SortNow "shortname_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
6040            End If
6050          Else
6060            frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
6070          End If
6080        Else
6090          frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
6100        End If
6110        DoEvents
6120        strSortNow2 = frm.SortNow_Get  ' ** Form Procedure: frmInvestmentObjective_Sub.
6130      Case False
6140        .chkShowAccounts_lbl.FontBold = False
6150        frm.ShowAccounts False  ' ** Form Procedure: frmInvestmentObjective_Sub.
6160        DoEvents
6170        strSortNow = frm.SortNow_Get  ' ** Form Function: frmInvestmentObjective_Sub.
6180        If strSortNow1 <> vbNullString Then
              ' ** Use the last sort this option used.
6190          If InStr(strSortNow1, " DESC") > 0 Then
6200            blnDesc = True
6210          Else
6220            blnDesc = False
6230          End If
6240          If Left(strSortNow1, 11) = "[invobj_id]" Then
6250            frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
6260            If blnDesc = True Then
6270              frm.SortNow "invobj_id_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
6280            End If
6290          ElseIf Left(strSortNow1, 13) = "[invobj_name]" Then
6300            frm.SortNow "invobj_name_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
6310            If blnDesc = True Then
6320              frm.SortNow "invobj_name_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
6330            End If
6340          ElseIf Left(strSortNow1, 12) = "[acct_count]" Then
6350            frm.SortNow "acct_count_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
6360            If blnDesc = True Then
6370              frm.SortNow "acct_count_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
6380            End If
6390          Else
6400            frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
6410          End If
6420        Else
6430          frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
6440        End If
6450        DoEvents
6460        strSortNow1 = frm.SortNow_Get  ' ** Form Procedure: frmInvestmentObjective_Sub.
6470      End Select
6480      DoEvents
6490      If lngRecsCur > 0& Then
6500        .frmInvestmentObjective_Sub.SetFocus
6510      End If
6520    End With

EXITP:
6530    Set frm = Nothing
6540    Exit Sub

ERRH:
6550    Select Case ERR.Number
        Case Else
6560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6570    End Select
6580    Resume EXITP

End Sub

Private Sub chkShowAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6620    intShiftDown = (Shift And acShiftMask) > 0
6630    intAltDown = (Shift And acAltMask) > 0
6640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6670        With Me
6680          intRetVal = 0
6690          If .chkShowUnassigned.Enabled = True Then
6700            .chkShowUnassigned.SetFocus
6710          ElseIf .cmdPrintReport.Enabled = True Then
6720            .cmdPrintReport.SetFocus
6730          Else
6740            .cmdClose.SetFocus
6750          End If
6760        End With
6770      End Select
6780    End If

        ' ** Shift keys.
6790    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6800      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6810        With Me
6820          intRetVal = 0
6830          .frmInvestmentObjective_Sub.SetFocus
6840          .frmInvestmentObjective_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmInvestmentObjective_Sub.
6850          .frmInvestmentObjective_Sub.Form.FocusHolder.SetFocus
6860        End With
6870      End Select
6880    End If

EXITP:
6890    KeyCode = intRetVal
6900    Exit Sub

ERRH:
6910    intRetVal = 0
6920    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6930    Case Else
6940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6950    End Select
6960    Resume EXITP

End Sub

Private Sub chkShowUnassigned_AfterUpdate()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowUnassigned_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, frm As Access.Form
        Dim strSortNow As String, blnDesc As Boolean
        Dim varTmp00 As Variant

7010    With Me
7020      DoCmd.Hourglass True
7030      DoEvents
7040      Set frm = .frmInvestmentObjective_Sub.Form
7050      Select Case .chkShowUnassigned
          Case True
7060        .chkShowUnassigned_lbl.FontBold = True
7070        .chkShowAccounts.Enabled = False
7080        Set dbs = CurrentDb
7090        With dbs
              ' ** Empty tblInvestmentObjective_Staging.
7100          Set qdf = .QueryDefs("qryInvestmentObjective_09_02")
7110          qdf.Execute
7120          Set qdf = Nothing
7130          DoEvents
              ' ** Append qryInvestmentObjective_09_01 (Account, just investmentobj = Null, as
              ' ** new tblInvestmentObjective_Staging records) to tblInvestmentObjective_Staging.
7140          Set qdf = .QueryDefs("qryInvestmentObjective_09_03")
7150          qdf.Execute
7160          Set qdf = Nothing
7170          DoEvents
7180          .Close
7190        End With
7200        Set dbs = Nothing
7210        If frm.RecordSource <> "qryInvestmentObjective_07_01" Then
              ' ** tblInvestmentObjective_Staging, all records.
7220          frm.RecordSource = "qryInvestmentObjective_07_01"
7230        End If
7240        frm.ShowUnassigned True  ' ** Form Procedure: frmInvestmentObjective_Sub.
7250        DoEvents
7260        strSortNow = frm.SortNow_Get  ' ** Form Function: frmInvestmentObjective_Sub.
7270        If strSortNow3 <> vbNullString Then
              ' ** Use the last sort this option used.
7280          If InStr(strSortNow3, " DESC") > 0 Then
7290            blnDesc = True
7300          Else
7310            blnDesc = False
7320          End If
7330          If Left(strSortNow3, 13) = "[invobj_name]" Then
7340            frm.SortNow "invobj_name_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
7350            If blnDesc = True Then
7360              frm.SortNow "invobj_name_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
7370            End If
7380          ElseIf Left(strSortNow3, 11) = "[alphasort]" Then
7390            frm.SortNow "accountno_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
7400            If blnDesc = True Then
7410              frm.SortNow "accountno_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
7420            End If
7430          ElseIf Left(strSortNow3, 11) = "[shortname]" Then
7440            frm.SortNow "shortname_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
7450            If blnDesc = True Then
7460              frm.SortNow "shortname_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
7470            End If
7480          Else
7490            frm.SortNow "accountno_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
7500          End If
7510        Else
7520          frm.SortNow "accountno_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
7530        End If
7540        DoEvents
7550        strSortNow3 = frm.SortNow_Get  ' ** Form Procedure: frmInvestmentObjective_Sub.
7560        .frmInvestmentObjective_Sub.SetFocus
7570        .frmInvestmentObjective_Sub.Form.cmbInvestmentObjectives.SetFocus
7580      Case False
7590        .chkShowUnassigned_lbl.FontBold = False
7600        .chkShowAccounts.Enabled = True
7610        DoEvents
7620        Set dbs = CurrentDb
7630        With dbs
              ' ** Update qryInvestmentObjective_10_01 (Account, linked to tblInvestmentObjective_Staging,
              ' ** just invobj_name <> '{Unassigned}, with investmentobj_new).
7640          Set qdf = .QueryDefs("qryInvestmentObjective_10_02")
7650          qdf.Execute
7660          Set qdf = Nothing
7670          .Close
7680        End With
7690        Set dbs = Nothing
7700        Select Case .chkShowAccounts
            Case True
7710          If frm.RecordSource <> "qryInvestmentObjective_04" Then
                ' ** qryInvestmentObjective_03 (InvestmentObjective, linked to
                ' ** qryInvestmentObjective_03_01 (Account, grouped by investmentobj,
                ' ** with cnt)), linked back to Account, with all accounts.'
7720            frm.RecordSource = "qryInvestmentObjective_04"
7730          End If
7740        Case False
7750          If frm.RecordSource <> "qryInvestmentObjective_03" Then
                ' ** InvestmentObjective, linked to qryInvestmentObjective_03_01
                ' ** (Account, grouped by investmentobj, with cnt)
7760            frm.RecordSource = "qryInvestmentObjective_03"
7770          End If
7780        End Select
7790        frm.ShowUnassigned False  ' ** Form Procedure: frmInvestmentObjective_Sub.
7800        DoEvents
7810        strSortNow = frm.SortNow_Get  ' ** Form Function: frmInvestmentObjective_Sub.
7820        Select Case .chkShowAccounts
            Case True
7830          If strSortNow2 <> vbNullString Then
                ' ** Use the last sort this option used.
7840            If InStr(strSortNow2, " DESC") > 0 Then
7850              blnDesc = True
7860            Else
7870              blnDesc = False
7880            End If
7890            If Left(strSortNow2, 11) = "[invobj_id]" Then
7900              frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
7910              If blnDesc = True Then
7920                frm.SortNow "invobj_id_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
7930              End If
7940            ElseIf Left(strSortNow2, 13) = "[invobj_name]" Then
7950              frm.SortNow "invobj_name_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
7960              If blnDesc = True Then
7970                frm.SortNow "invobj_name_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
7980              End If
7990            ElseIf Left(strSortNow2, 11) = "[alphasort]" Then
8000              frm.SortNow "accountno_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
8010              If blnDesc = True Then
8020                frm.SortNow "accountno_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
8030              End If
8040            ElseIf Left(strSortNow2, 11) = "[shortname]" Then
8050              frm.SortNow "shortname_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
8060              If blnDesc = True Then
8070                frm.SortNow "shortname_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
8080              End If
8090            Else
8100              frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
8110            End If
8120          Else
8130            frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
8140          End If
8150          DoEvents
8160          strSortNow2 = frm.SortNow_Get  ' ** Form Procedure: frmInvestmentObjective_Sub.
8170        Case False
8180          If strSortNow1 <> vbNullString Then
                ' ** Use the last sort this option used.
8190            If InStr(strSortNow1, " DESC") > 0 Then
8200              blnDesc = True
8210            Else
8220              blnDesc = False
8230            End If
8240            If Left(strSortNow1, 11) = "[invobj_id]" Then
8250              frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
8260              If blnDesc = True Then
8270                frm.SortNow "invobj_id_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
8280              End If
8290            ElseIf Left(strSortNow1, 13) = "[invobj_name]" Then
8300              frm.SortNow "invobj_name_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
8310              If blnDesc = True Then
8320                frm.SortNow "invobj_name_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
8330              End If
8340            ElseIf Left(strSortNow1, 12) = "[acct_count]" Then
8350              frm.SortNow "acct_count_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
8360              If blnDesc = True Then
8370                frm.SortNow "acct_count_lbl_DblClick", False  ' ** Form Procedure: frmInvestmentObjective_Sub.
8380              End If
8390            Else
8400              frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
8410            End If
8420          Else
8430            frm.SortNow "invobj_id_lbl_DblClick", True  ' ** Form Procedure: frmInvestmentObjective_Sub.
8440          End If
8450          DoEvents
8460          strSortNow1 = frm.SortNow_Get  ' ** Form Procedure: frmInvestmentObjective_Sub.
8470        End Select
            ' ** Account, just investmentobj = Null.
8480        varTmp00 = DCount("*", "qryInvestmentObjective_08")
8490        If IsNull(varTmp00) = True Then
8500          .Unassigned_lbl.Visible = False
8510          .chkShowAccounts.SetFocus
8520          .chkShowUnassigned.Enabled = False
8530        Else
8540          If varTmp00 = 0 Then
8550            .Unassigned_lbl.Visible = False
8560            .chkShowAccounts.SetFocus
8570            .chkShowUnassigned.Enabled = False
8580          Else
                ' ** Unassigned accounts remain.
8590          End If
8600        End If
8610        .frmInvestmentObjective_Sub.SetFocus
8620      End Select
8630      DoCmd.Hourglass False
8640    End With

EXITP:
8650    Set frm = Nothing
8660    Set qdf = Nothing
8670    Set dbs = Nothing
8680    Exit Sub

ERRH:
8690    DoCmd.Hourglass False
8700    Select Case ERR.Number
        Case Else
8710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8720    End Select
8730    Resume EXITP

End Sub

Private Sub chkShowUnassigned_KeyDown(KeyCode As Integer, Shift As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowUnassigned_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8820    intShiftDown = (Shift And acShiftMask) > 0
8830    intAltDown = (Shift And acAltMask) > 0
8840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8870        With Me
8880          intRetVal = 0
8890          If .cmdPrintReport.Enabled = True Then
8900            .cmdPrintReport.SetFocus
8910          Else
8920            .cmdClose.SetFocus
8930          End If
8940        End With
8950      End Select
8960    End If

        ' ** Shift keys.
8970    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8980      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8990        With Me
9000          intRetVal = 0
9010          If .chkShowAccounts.Enabled = True Then
9020            .chkShowAccounts.SetFocus
9030          Else
9040            .frmInvestmentObjective_Sub.SetFocus
9050            .frmInvestmentObjective_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmInvestmentObjective_Sub.
9060            .frmInvestmentObjective_Sub.Form.FocusHolder.SetFocus
9070          End If
9080        End With
9090      End Select
9100    End If

EXITP:
9110    KeyCode = intRetVal
9120    Exit Sub

ERRH:
9130    intRetVal = 0
9140    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
9150    Case Else
9160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9170    End Select
9180    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

9210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9220    Exit Sub

ERRH:
9230    Select Case ERR.Number
        Case Else
9240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9250    End Select
9260    Resume EXITP

End Sub

Public Sub SortNow_Set(strSort As String, intMode As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Set"

9310    Select Case intMode
        Case 1
9320      strSortNow1 = strSort
9330    Case 2
9340      strSortNow2 = strSort
9350    Case 3
9360      strSortNow3 = strSort
9370    End Select

EXITP:
9380    Exit Sub

ERRH:
9390    Select Case ERR.Number
        Case Else
9400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9410    End Select
9420    Resume EXITP

End Sub
