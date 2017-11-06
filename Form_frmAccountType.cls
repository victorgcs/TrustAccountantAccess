VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountType"

'VGC 09/23/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Show Accounts:    H {chkShowAccounts}
' **   Print Report:     P {cmdPrintReport}
' **   Exit:             X {cmdClose}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, lngPrintReportLbl_Left As Long, blnWasGTR As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

160           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

170           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left
180           blnWasGTR = gblnGoToReport

190           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

200           chkShowAccounts_AfterUpdate  ' ** Procedure: Below.

210         Else
220           Cancel = -1
230         End If
240       Else
250         Cancel = -1
260       End If
270     End With

280     If Cancel = -1 Then
290       Beep
300       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
310     End If

EXITP:
320     Exit Sub

ERRH:
330     Select Case ERR.Number
        Case Else
340       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
350     End Select
360     Resume EXITP

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
480         cmdClose_Click  ' ** Procedure: Below.
490       End Select
500     End If

        ' ** Alt keys.
510     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
520       Select Case intRetVal
          Case vbKeyX
530         intRetVal = 0
540         cmdClose_Click  ' ** Procedure: Below.
550       End Select
560     End If

        ' ** Ctrl-Shift keys.
570     If intCtrlDown And (Not intAltDown) And intShiftDown Then
580       Select Case intRetVal
          Case vbKeyF
590         With Me
600           intRetVal = 0
610           .FocusHolder.SetFocus
620         End With
630       End Select
640     End If

EXITP:
650     KeyCode = intRetVal
660     Exit Sub

ERRH:
670     intRetVal = 0
680     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
690     Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub Form_Timer()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

810     With Me
820       .TimerInterval = 0&
830       If gblnSetFocus = True Then
840         gblnSetFocus = False
850         strFormName = vbNullString
860   On Error Resume Next
870         Do Until strFormName = THIS_NAME
880           DoCmd.SelectObject acForm, THIS_NAME, False
890           strFormName = Screen.ActiveForm.Name
900         Loop
910   On Error GoTo ERRH
920         .frmAccountType_Sub.SetFocus
930       End If
940       If gblnGoToReport = True Then
950         Select Case blnGoingToReport
            Case True
960           Select Case blnGoingToReport2
              Case True
970             DoCmd.Hourglass False  ' ** It should already be off.
980             gblnGoToReport = False
990             blnGoingToReport = False
1000            blnGoingToReport2 = False
1010            .GoToReport_arw_pr_img.Visible = False
1020            .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1030            DoEvents
1040          Case False
1050            If garr_varGoToReport(GTR_RNAM) = "rptListOfAccountTypes_Detailed" Then
1060              blnGoingToReport2 = True
1070              .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1080              .GoToReport_arw_pr_img.Visible = True
1090              .GoToReport_arw_det_img.Visible = False
1100              .cmdPrintReport.SetFocus
1110              Beep
1120              DoCmd.Hourglass False
1130              DoEvents
1140              .TimerInterval = (GTR_WAIT * 2&)
1150            Else
1160              DoCmd.Hourglass False  ' ** It should already be off.
1170              gblnGoToReport = False
1180              blnGoingToReport = False
1190              .GoToReport_arw_pr_img.Visible = False
1200              .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1210              DoEvents
1220            End If
1230          End Select
1240        Case False
1250          DoCmd.Hourglass True  ' ** Make sure it's still running.
1260          DoEvents
1270          blnGoingToReport = True
1280          If garr_varGoToReport(GTR_RNAM) = "rptListOfAccountTypes_Detailed" Then
1290            .GoToReport_arw_det_img.Visible = True
1300            DoEvents
1310            If .chkShowAccounts = False Then
1320              .chkShowAccounts = True
1330              chkShowAccounts_AfterUpdate  ' ** Procedure: Below.
1340            End If
1350            DoEvents
1360            .TimerInterval = GTR_WAIT
1370          Else
1380            If .chkShowAccounts = True Then
1390              .chkShowAccounts = False
1400              chkShowAccounts_AfterUpdate  ' ** Procedure: Below.
1410            End If
1420            DoEvents
1430            .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1440            .GoToReport_arw_pr_img.Visible = True
1450            .cmdPrintReport.SetFocus
1460            Beep
1470            DoCmd.Hourglass False
1480            DoEvents
1490            .TimerInterval = (GTR_WAIT * 2&)
1500          End If
1510        End Select
1520      Else
1530        DoCmd.Hourglass False
1540      End If
1550    End With

EXITP:
1560    Exit Sub

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim blnRetVal As Boolean

1710    With Me

1720      DoCmd.Hourglass True
1730      DoEvents

1740      Select Case blnWasGTR
          Case True
1750        blnWasGTR = False
1760      Case False
            ' ** Don't save settings imposed by GoToReport!
1770        Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
1780      End Select

1790      blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

1800      gstrReturningForm = THIS_NAME
1810      gblnSetFocus = True
1820      strDocName = "frmMenu_SystemTable"
1830      DoCmd.OpenForm strDocName, , , , , , "frmMenu_Maintenance"

1840    End With

EXITP:
1850    Exit Sub

ERRH:
1860    DoCmd.Hourglass False
1870    Select Case ERR.Number
        Case Else
1880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1890    End Select
1900    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2010    With Me
2020      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2030        Select Case blnPrintReport_Focus
            Case True
2040          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2050          .cmdPrintReport_raised_img.Visible = False
2060        Case False
2070          .cmdPrintReport_raised_img.Visible = True
2080          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2090        End Select
2100        .cmdPrintReport_raised_focus_img.Visible = False
2110        .cmdPrintReport_raised_focus_dots_img.Visible = False
2120        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2130        .cmdPrintReport_raised_img_dis.Visible = False
2140      End If
2150    End With

EXITP:
2160    Exit Sub

ERRH:
2170    Select Case ERR.Number
        Case Else
2180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2190    End Select
2200    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2410    DoCmd.Close acForm, THIS_NAME

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2520    intShiftDown = (Shift And acShiftMask) > 0
2530    intAltDown = (Shift And acAltMask) > 0
2540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2560      Select Case intRetVal
          Case vbKeyTab
2570        With Me
2580          intRetVal = 0
2590          .frmAccountType_Sub.SetFocus
2600          .frmAccountType_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAccountType_Sub.
2610          .frmAccountType_Sub.Form.FocusHolder.SetFocus
2620        End With
2630      End Select
2640    End If

        ' ** Shift keys.
2650    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2660      Select Case intRetVal
          Case vbKeyTab
2670        With Me
2680          intRetVal = 0
2690          If .cmdPrintReport.Enabled = True Then
2700            .cmdPrintReport.SetFocus
2710          Else
2720            .chkShowAccounts.SetFocus
2730          End If
2740        End With
2750      End Select
2760    End If

EXITP:
2770    KeyCode = intRetVal
2780    Exit Sub

ERRH:
2790    intRetVal = 0
2800    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2810    Case Else
2820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2830    End Select
2840    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

2910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2920    Exit Sub

ERRH:
2930    Select Case ERR.Number
        Case Else
2940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2950    End Select
2960    Resume EXITP

End Sub

Private Sub chkShowAccounts_AfterUpdate()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAccounts_AfterUpdate"

3010    With Me
3020      Select Case .chkShowAccounts
          Case True
3030        .chkShowAccounts_lbl.FontBold = True
3040        If .frmAccountType_Sub.Form.RecordSource <> "qryAccountType_Sub_02" Then
              ' ** qryAccountType_Sub_01 (AccountType, linked to AccountTypeGrouping),
              ' ** linked to Account, with accountno, shortname.
3050          .frmAccountType_Sub.Form.RecordSource = "qryAccountType_Sub_02"
3060        End If
3070        .frmAccountType_Sub.Form.ShowAccount True  ' ** Form Procedure: frmAccountType_Sub.
3080      Case False
3090        .chkShowAccounts_lbl.FontBold = False
3100        If .frmAccountType_Sub.Form.RecordSource <> "qryAccountType_Sub_01" Then
              ' ** AccountType, linked to AccountTypeGrouping.
3110          .frmAccountType_Sub.Form.RecordSource = "qryAccountType_Sub_01"
3120        End If
3130        .frmAccountType_Sub.Form.ShowAccount False  ' ** Form Procedure: frmAccountType_Sub.
3140      End Select
3150      .frmAccountType_Sub.Form.Requery
3160      .frmAccountType_Sub.SetFocus
3170    End With

EXITP:
3180    Exit Sub

ERRH:
3190    Select Case ERR.Number
        Case Else
3200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3210    End Select
3220    Resume EXITP

End Sub

Private Sub chkShowAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAccounts_KeyDown"

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
3390          If .cmdPrintReport.Enabled = True Then
3400            .cmdPrintReport.SetFocus
3410          Else
3420            .cmdClose.SetFocus
3430          End If
3440        End With
3450      End Select
3460    End If

        ' ** Shift keys.
3470    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3480      Select Case intRetVal
          Case vbKeyTab
3490        With Me
3500          intRetVal = 0
3510          .frmAccountType_Sub.SetFocus
3520          .frmAccountType_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountType_Sub.
3530          .frmAccountType_Sub.Form.FocusHolder.SetFocus
3540        End With
3550      End Select
3560    End If

EXITP:
3570    KeyCode = intRetVal
3580    Exit Sub

ERRH:
3590    intRetVal = 0
3600    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3610    Case Else
3620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3630    End Select
3640    Resume EXITP

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

        Dim strDocName As String

4110    With Me
4120      Select Case .chkShowAccounts
          Case True
4130        strDocName = "rptListOfAccountTypes_Detailed"
4140      Case False
4150        strDocName = "rptListOfAccountTypes"
4160      End Select
4170      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4180        DoCmd.OpenReport strDocName, acViewPreview
4190        DoCmd.Maximize
4200        DoCmd.RunCommand acCmdFitToWindow
4210        .Visible = False
4220      Else
            '##GTR_Ref: rptListOfAccountTypes_Detailed
            '##GTR_Ref: rptListOfAccountTypes
4230        DoCmd.OpenReport strDocName, acViewNormal
4240      End If
4250    End With

EXITP:
4260    Exit Sub

ERRH:
4270    Me.Visible = True
4280    DoCmd.Restore
4290    Select Case ERR.Number
        Case Else
4300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4310    End Select
4320    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4420    intShiftDown = (Shift And acShiftMask) > 0
4430    intAltDown = (Shift And acAltMask) > 0
4440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4460      Select Case intRetVal
          Case vbKeyTab
4470        With Me
4480          intRetVal = 0
4490          .cmdClose.SetFocus
4500        End With
4510      End Select
4520    End If

        ' ** Shift keys.
4530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4540      Select Case intRetVal
          Case vbKeyTab
4550        With Me
4560          intRetVal = 0
4570          .chkShowAccounts.SetFocus
4580        End With
4590      End Select
4600    End If

EXITP:
4610    KeyCode = intRetVal
4620    Exit Sub

ERRH:
4630    intRetVal = 0
4640    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4650    Case Else
4660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4670    End Select
4680    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

4710    With Me
4720      If blnPrintReport_MouseDown = False Then
4730        Select Case blnPrintReport_Focus
            Case True
4740          .cmdPrintReport_raised_focus_dots_img.Visible = True
4750          .cmdPrintReport_raised_focus_img.Visible = False
4760        Case False
4770          .cmdPrintReport_raised_focus_img.Visible = True
4780          .cmdPrintReport_raised_focus_dots_img.Visible = False
4790        End Select
4800        .cmdPrintReport_raised_img.Visible = False
4810        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4820        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4830        .cmdPrintReport_raised_img_dis.Visible = False
4840      End If
4850    End With

EXITP:
4860    Exit Sub

ERRH:
4870    Select Case ERR.Number
        Case Else
4880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4890    End Select
4900    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

5010    With Me
5020      .cmdPrintReport_raised_focus_dots_img.Visible = True
5030      .cmdPrintReport_raised_img.Visible = False
5040      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5050      .cmdPrintReport_raised_focus_img.Visible = False
5060      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5070      .cmdPrintReport_raised_img_dis.Visible = False
5080      blnPrintReport_MouseDown = False
5090    End With

EXITP:
5100    Exit Sub

ERRH:
5110    Select Case ERR.Number
        Case Else
5120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5130    End Select
5140    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

5210    With Me
5220      .cmdPrintReport_raised_img.Visible = True
5230      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5240      .cmdPrintReport_raised_focus_img.Visible = False
5250      .cmdPrintReport_raised_focus_dots_img.Visible = False
5260      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5270      .cmdPrintReport_raised_img_dis.Visible = False
5280      blnPrintReport_Focus = False
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
