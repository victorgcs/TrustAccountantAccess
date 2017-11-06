VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssetTypeGrouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssetTypeGrouping"

'VGC 09/23/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Show Types:       H {chkShowAssetType}
' **   Print Report:     P {cmdPrintReport}
' **   Exit:             X {cmdClose}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, lngPrintReportLbl_Left As Long, blnWasGTR As Boolean
Private lngTpp As Long
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

150           blnWasGTR = gblnGoToReport

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
170           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

180           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

190           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

200           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

210           chkShowAssetType_AfterUpdate  ' ** Procedure: Below.

220         Else
230           Cancel = -1
240         End If
250       Else
260         Cancel = -1
270       End If
280     End With

290     If Cancel = -1 Then
300       Beep
310       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
320     End If

EXITP:
330     Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

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
920         .frmAssetTypeGrouping_Sub.SetFocus
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
1020            .cmdPrintReport_lbl2.Visible = True
1030            DoEvents
1040          Case False
1050            If garr_varGoToReport(GTR_RNAM) = "rptListOfAssetTypeGroups_Detailed" Then
1060              blnGoingToReport2 = True
1070              .cmdPrintReport_lbl2.Visible = False
1080              .GoToReport_arw_det_img.Visible = False
1090              .GoToReport_arw_pr_img.Visible = True
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
1200              .cmdPrintReport_lbl2.Visible = True
1210              DoEvents
1220            End If
1230          End Select
1240        Case False
1250          DoCmd.Hourglass True  ' ** Make sure it's still running.
1260          DoEvents
1270          blnGoingToReport = True
1280          If garr_varGoToReport(GTR_RNAM) = "rptListOfAssetTypeGroups_Detailed" Then
1290            .GoToReport_arw_det_img.Visible = True
1300            DoEvents
1310            If .chkShowAssetType = False Then
1320              .chkShowAssetType = True
1330              chkShowAssetType_AfterUpdate  ' ** Procedure: Below.
1340            End If
1350            DoEvents
1360            .TimerInterval = GTR_WAIT
1370          Else
1380            .cmdPrintReport_lbl2.Visible = False
1390            .GoToReport_arw_pr_img.Visible = True
1400            .cmdPrintReport.SetFocus
1410            DoEvents
1420            If .chkShowAssetType = True Then
1430              .chkShowAssetType = False
1440              chkShowAssetType_AfterUpdate  ' ** Procedure: Below.
1450              DoEvents
1460            End If
1470            Beep
1480            DoCmd.Hourglass False
1490            DoEvents
1500            .TimerInterval = (GTR_WAIT & 2&)
1510          End If
1520        End Select
1530      Else
1540        DoCmd.Hourglass False
1550      End If
1560    End With

EXITP:
1570    Exit Sub

ERRH:
1580    Select Case ERR.Number
        Case Else
1590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1600    End Select
1610    Resume EXITP

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
2590          .frmAssetTypeGrouping_Sub.SetFocus
2600          .frmAssetTypeGrouping_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmassettypeGrouping_Sub.
2610          .frmAssetTypeGrouping_Sub.Form.FocusHolder.SetFocus
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
2720            .chkShowAssetType.SetFocus
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

Public Sub chkShowAssetType_AfterUpdate()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssetType_AfterUpdate"

2910    With Me
2920      Select Case .chkShowAssetType
          Case True
2930        .chkShowAssetType_lbl.FontBold = True
2940        .frmAssetTypeGrouping_Sub.Form.ShowAssetType True
2950      Case False
2960        .chkShowAssetType_lbl.FontBold = False
2970        .frmAssetTypeGrouping_Sub.Form.ShowAssetType False
2980      End Select
2990      .frmAssetTypeGrouping_Sub.SetFocus
3000    End With

EXITP:
3010    Exit Sub

ERRH:
3020    Select Case ERR.Number
        Case Else
3030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3040    End Select
3050    Resume EXITP

End Sub

Private Sub chkShowAssetType_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssetType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3120    intShiftDown = (Shift And acShiftMask) > 0
3130    intAltDown = (Shift And acAltMask) > 0
3140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3170        With Me
3180          intRetVal = 0
3190          If .cmdPrintReport.Enabled = True Then
3200            .cmdPrintReport.SetFocus
3210          Else
3220            .cmdClose.SetFocus
3230          End If
3240        End With
3250      End Select
3260    End If

        ' ** Shift keys.
3270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3280      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3290        With Me
3300          intRetVal = 0
3310          .frmAssetTypeGrouping_Sub.SetFocus
3320          .frmAssetTypeGrouping_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmassettypeGrouping_Sub.
3330          .frmAssetTypeGrouping_Sub.Form.FocusHolder.SetFocus
3340        End With
3350      End Select
3360    End If

EXITP:
3370    KeyCode = intRetVal
3380    Exit Sub

ERRH:
3390    intRetVal = 0
3400    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3410    Case Else
3420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3430    End Select
3440    Resume EXITP

End Sub

Private Sub chkShowAssetType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssetType_MouseMove"

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

Private Sub cmdPrintReport_GotFocus()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

3610    With Me
3620      blnPrintReport_Focus = True
3630      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3640      .cmdPrintReport_raised_img.Visible = False
3650      .cmdPrintReport_raised_focus_img.Visible = False
3660      .cmdPrintReport_raised_focus_dots_img.Visible = False
3670      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3680      .cmdPrintReport_raised_img_dis.Visible = False
3690    End With

EXITP:
3700    Exit Sub

ERRH:
3710    Select Case ERR.Number
        Case Else
3720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3730    End Select
3740    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

3810    With Me
3820      blnPrintReport_MouseDown = True
3830      .cmdPrintReport_sunken_focus_dots_img.Visible = True
3840      .cmdPrintReport_raised_img.Visible = False
3850      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3860      .cmdPrintReport_raised_focus_img.Visible = False
3870      .cmdPrintReport_raised_focus_dots_img.Visible = False
3880      .cmdPrintReport_raised_img_dis.Visible = False
3890    End With

EXITP:
3900    Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

4010    With Me
4020      Select Case .chkShowAssetType
          Case True
4030        strDocName = "rptListOfAssetTypeGroups_Detailed"
4040      Case False
4050        strDocName = "rptListOfAssetTypeGroups"
4060      End Select
4070      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4080        DoCmd.OpenReport strDocName, acViewPreview
4090        DoCmd.Maximize
4100        DoCmd.RunCommand acCmdFitToWindow
4110        .Visible = False
4120      Else
            '##GTR_Ref: rptListOfAssetTypeGroups
            '##GTR_Ref: rptListOfAssetTypeGroups_Detailed
4130        DoCmd.OpenReport strDocName, acViewNormal
4140      End If
4150    End With

EXITP:
4160    Exit Sub

ERRH:
4170    Me.Visible = True
4180    DoCmd.Restore
4190    Select Case ERR.Number
        Case Else
4200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4210    End Select
4220    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4320    intShiftDown = (Shift And acShiftMask) > 0
4330    intAltDown = (Shift And acAltMask) > 0
4340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4360      Select Case intRetVal
          Case vbKeyTab
4370        With Me
4380          intRetVal = 0
4390          .cmdClose.SetFocus
4400        End With
4410      End Select
4420    End If

        ' ** Shift keys.
4430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4440      Select Case intRetVal
          Case vbKeyTab
4450        With Me
4460          intRetVal = 0
4470          .chkShowAssetType.SetFocus
4480        End With
4490      End Select
4500    End If

EXITP:
4510    KeyCode = intRetVal
4520    Exit Sub

ERRH:
4530    intRetVal = 0
4540    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4550    Case Else
4560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4570    End Select
4580    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

4610    With Me
4620      If blnPrintReport_MouseDown = False Then
4630        Select Case blnPrintReport_Focus
            Case True
4640          .cmdPrintReport_raised_focus_dots_img.Visible = True
4650          .cmdPrintReport_raised_focus_img.Visible = False
4660        Case False
4670          .cmdPrintReport_raised_focus_img.Visible = True
4680          .cmdPrintReport_raised_focus_dots_img.Visible = False
4690        End Select
4700        .cmdPrintReport_raised_img.Visible = False
4710        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4720        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4730        .cmdPrintReport_raised_img_dis.Visible = False
4740      End If
4750    End With

EXITP:
4760    Exit Sub

ERRH:
4770    Select Case ERR.Number
        Case Else
4780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4790    End Select
4800    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

4910    With Me
4920      .cmdPrintReport_raised_focus_dots_img.Visible = True
4930      .cmdPrintReport_raised_img.Visible = False
4940      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4950      .cmdPrintReport_raised_focus_img.Visible = False
4960      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4970      .cmdPrintReport_raised_img_dis.Visible = False
4980      blnPrintReport_MouseDown = False
4990    End With

EXITP:
5000    Exit Sub

ERRH:
5010    Select Case ERR.Number
        Case Else
5020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5030    End Select
5040    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

5110    With Me
5120      .cmdPrintReport_raised_img.Visible = True
5130      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5140      .cmdPrintReport_raised_focus_img.Visible = False
5150      .cmdPrintReport_raised_focus_dots_img.Visible = False
5160      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5170      .cmdPrintReport_raised_img_dis.Visible = False
5180      blnPrintReport_Focus = False
5190    End With

EXITP:
5200    Exit Sub

ERRH:
5210    Select Case ERR.Number
        Case Else
5220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5230    End Select
5240    Resume EXITP

End Sub

Private Sub cmdPrintReport_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_lbl2_MouseMove"

5310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

5410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5420    Exit Sub

ERRH:
5430    Select Case ERR.Number
        Case Else
5440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5450    End Select
5460    Resume EXITP

End Sub
