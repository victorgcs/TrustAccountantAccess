VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmIncomeExpenseCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmIncomeExpenseCodes"

'VGC 10/02/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd}
' **   Close:            C {cmdClose}
' **   Edit:             E {cmdEdit}
' **   Print Report:     P {cmdPrintReport}
' **   Renumber:         R {cmdRenumber}
' **   Update:           U {cmdUpdate}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmIncomeExpenseCodes_Sub}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Tax Codes:        T {ShowTaxCodes on frmIncomeExpenseCodes_Sub}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, lngPrintReportLbl_Left As Long, lngRecsCur As Long
Private blnEdited As Boolean, blnMoveAll As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.

150           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

160           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

170           CLR_DISABLED_FG = CLR_DKGRY
180           CLR_DISABLED_BG = CLR_LTTEAL

              ' ** Default until they edit something or use cmdEdit.
190           blnEdited = False: gblnClosing = False
200           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

210           .DefaultCodes_lbl.Visible = False

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

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Tax Codes:        T {ShowTaxCodes on frmIncomeExpenseCodes_Sub}

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

        ' ** Ctrl keys.
570     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
580       Select Case intRetVal
          Case vbKeyS
590         With Me
600           intRetVal = 0
610           .frmIncomeExpenseCodes_Sub.SetFocus
620           .frmIncomeExpenseCodes_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.
630         End With
640       End Select
650     End If

        ' ** Ctrl-Shift keys.
660     If intCtrlDown And (Not intAltDown) And intShiftDown Then
670       Select Case intRetVal
          Case vbKeyF
680         With Me
690           intRetVal = 0
700           .FocusHolder.SetFocus
710         End With
720       Case vbKeyT
730         With Me
740           intRetVal = 0
750           .frmIncomeExpenseCodes_Sub.Form.ShowTaxCodes  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.
760         End With
770       End Select
780     End If

EXITP:
790     KeyCode = intRetVal
800     Exit Sub

ERRH:
810     intRetVal = 0
820     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
830     Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Sub

Private Sub Form_Timer()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

910     With Me
920       .TimerInterval = 0&
930       If gblnSetFocus = True Then
940         gblnSetFocus = False
950         strFormName = vbNullString
960   On Error Resume Next
970         Do Until strFormName = THIS_NAME
980           DoCmd.SelectObject acForm, THIS_NAME, False
990           strFormName = Screen.ActiveForm.Name
1000        Loop
1010  On Error GoTo ERRH
1020      End If
1030      .frmIncomeExpenseCodes_Sub.SetFocus
1040      If gblnGoToReport = True Then
1050        Select Case blnGoingToReport
            Case True
1060          DoCmd.Hourglass False  ' ** It should already be off.
1070          blnGoingToReport = False
1080          gblnGoToReport = False
1090          .GoToReport_arw_pr_img.Visible = False
1100          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1110          DoEvents
1120        Case False
1130          DoCmd.Hourglass True  ' ** Make sure it's still running.
1140          DoEvents
1150          blnGoingToReport = True
1160          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1170          .GoToReport_arw_pr_img.Visible = True
1180          .cmdPrintReport.SetFocus
1190          DoEvents
1200          Beep
1210          DoCmd.Hourglass False
1220          DoEvents
1230          .TimerInterval = (GTR_WAIT * 2&)
1240        End Select
1250      Else
1260        DoCmd.Hourglass False
1270      End If
1280    End With

EXITP:
1290    Exit Sub

ERRH:
1300    Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Ignore.
1310    Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim blnRetVal As Boolean

1410    With Me

1420      blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

1430      Set dbs = CurrentDb
          ' ** Empty tmpRevCodeEdit.
1440      Set qdf = dbs.QueryDefs("qryRevCodes_12")
1450      qdf.Execute
1460      Set qdf = Nothing
1470      dbs.Close

1480      Select Case strCallingForm
          Case "frmMenu_Utility"
1490        gblnSetFocus = True
1500        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
1510      Case vbNullString
1520        gblnSetFocus = True
1530        strDocName = "frmMenu_Utility"
1540        DoCmd.OpenForm strDocName, , , , , , "frmMenu_Main"
1550      Case Else
1560        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
1570      End Select
1580      .frmIncomeExpenseCodes_Sub.Form.RecordSource = ""

1590    End With

EXITP:
1600    Set qdf = Nothing
1610    Set dbs = Nothing
1620    Exit Sub

ERRH:
1630    Select Case ERR.Number
        Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1710    With Me
1720      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
1730        Select Case blnPrintReport_Focus
            Case True
1740          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
1750          .cmdPrintReport_raised_img.Visible = False
1760        Case False
1770          .cmdPrintReport_raised_img.Visible = True
1780          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
1790        End Select
1800        .cmdPrintReport_raised_focus_img.Visible = False
1810        .cmdPrintReport_raised_focus_dots_img.Visible = False
1820        .cmdPrintReport_sunken_focus_dots_img.Visible = False
1830        .cmdPrintReport_raised_img_dis.Visible = False
1840      End If
1850    End With

EXITP:
1860    Exit Sub

ERRH:
1870    Select Case ERR.Number
        Case Else
1880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1890    End Select
1900    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim blnContinue As Boolean

2110    With Me

2120      blnContinue = True  ' ** Unless proven otherwise.

2130      If blnEdited = True Then
2140        If MsgBox("You have made changes that you have not saved." & vbCrLf & vbCrLf & _
                "Are you sure you want to exit without saving?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved") = vbNo Then
2150          blnContinue = False
2160        End If
2170      End If

2180      If blnContinue = True Then

2190        gblnClosing = True
2200        .frmIncomeExpenseCodes_Sub.Form.ClosingForm gblnClosing  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.

2210        DoCmd.Close acForm, THIS_NAME

2220      End If

2230    End With

EXITP:
2240    Exit Sub

ERRH:
2250    Select Case ERR.Number
        Case Else
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2320    intShiftDown = (Shift And acShiftMask) > 0
2330    intAltDown = (Shift And acAltMask) > 0
2340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2360      Select Case intRetVal
          Case vbKeyTab
2370        With Me
2380          intRetVal = 0
2390          .frmIncomeExpenseCodes_Sub.SetFocus
2400          .frmIncomeExpenseCodes_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.
2410          .frmIncomeExpenseCodes_Sub.Form.revcode_DESC.SetFocus
2420        End With
2430      End Select
2440    End If

        ' ** Shift keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2460      Select Case intRetVal
          Case vbKeyTab
2470        With Me
2480          intRetVal = 0
2490          If .cmdUpdate.Enabled = True Then
2500            .cmdUpdate.SetFocus
2510          ElseIf .cmdEdit.Enabled = True Then
2520            .cmdEdit.SetFocus
2530          ElseIf .cmdAdd.Enabled = True Then
2540            .cmdAdd.SetFocus
2550          ElseIf .cmdRenumber.Enabled = True Then
2560            .cmdRenumber.SetFocus
2570          ElseIf .cmdPrintReport.Enabled = True Then
2580            .cmdPrintReport.SetFocus
2590          Else
                ' ** That's far enough!
2600          End If
2610        End With
2620      End Select
2630    End If

EXITP:
2640    KeyCode = intRetVal
2650    Exit Sub

ERRH:
2660    intRetVal = 0
2670    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2680    Case Else
2690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2700    End Select
2710    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

2810    With Me
2820      blnPrintReport_Focus = True
2830      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2840      .cmdPrintReport_raised_img.Visible = False
2850      .cmdPrintReport_raised_focus_img.Visible = False
2860      .cmdPrintReport_raised_focus_dots_img.Visible = False
2870      .cmdPrintReport_sunken_focus_dots_img.Visible = False
2880      .cmdPrintReport_raised_img_dis.Visible = False
2890    End With

EXITP:
2900    Exit Sub

ERRH:
2910    Select Case ERR.Number
        Case Else
2920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2930    End Select
2940    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

3010    With Me
3020      blnPrintReport_MouseDown = True
3030      .cmdPrintReport_sunken_focus_dots_img.Visible = True
3040      .cmdPrintReport_raised_img.Visible = False
3050      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3060      .cmdPrintReport_raised_focus_img.Visible = False
3070      .cmdPrintReport_raised_focus_dots_img.Visible = False
3080      .cmdPrintReport_raised_img_dis.Visible = False
3090    End With

EXITP:
3100    Exit Sub

ERRH:
3110    Select Case ERR.Number
        Case Else
3120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3130    End Select
3140    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String, strSortNow As String

3210    With Me
3220      strSortNow = .frmIncomeExpenseCodes_Sub.Form.SortNow_Get  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.
3230      If strSortNow = "[revcode_TYPE], [revcode_SORTORDER]" Then
3240        strDocName = "rptListOfIncExpCodes"
3250      Else
3260        strDocName = "rptListOfIncExpCodes_Sort"
3270      End If
3280      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
3290        DoCmd.OpenReport strDocName, acViewPreview
3300        DoCmd.Maximize
3310        DoCmd.RunCommand acCmdFitToWindow
3320        .Visible = False
3330      Else
            '##GTR_Ref: rptListOfIncExpCodes
3340        DoCmd.OpenReport strDocName, acViewNormal
3350      End If
3360    End With

EXITP:
3370    Exit Sub

ERRH:
3380    Me.Visible = True
3390    DoCmd.Restore
3400    Select Case ERR.Number
        Case Else
3410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3420    End Select
3430    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
          Case vbKeyTab
3570        With Me
3580          intRetVal = 0
3590          If .cmdRenumber.Enabled = True Then
3600            .cmdRenumber.SetFocus
3610          ElseIf .cmdAdd.Enabled = True Then
3620            .cmdAdd.SetFocus
3630          ElseIf .cmdEdit.Enabled = True Then
3640            .cmdEdit.SetFocus
3650          ElseIf .cmdUpdate.Enabled = True Then
3660            .cmdUpdate.SetFocus
3670          Else
3680            .cmdClose.SetFocus
3690          End If
3700        End With
3710      End Select
3720    End If

        ' ** Shift keys.
3730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3740      Select Case intRetVal
          Case vbKeyTab
3750        With Me
3760          intRetVal = 0
3770          .frmIncomeExpenseCodes_Sub.SetFocus
3780          With .frmIncomeExpenseCodes_Sub.Form
3790            .MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.
3800            If .taxcode_type.Visible = True And .taxcode_type.Enabled = True Then
3810              .taxcode_type.SetFocus
3820            Else
3830              .revcode_ACTIVE.SetFocus
3840            End If
3850          End With
3860        End With
3870      End Select
3880    End If

EXITP:
3890    KeyCode = intRetVal
3900    Exit Sub

ERRH:
3910    intRetVal = 0
3920    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
3930    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
3940    Case Else
3950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3960    End Select
3970    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

4010    With Me
4020      If blnPrintReport_MouseDown = False Then
4030        Select Case blnPrintReport_Focus
            Case True
4040          .cmdPrintReport_raised_focus_dots_img.Visible = True
4050          .cmdPrintReport_raised_focus_img.Visible = False
4060        Case False
4070          .cmdPrintReport_raised_focus_img.Visible = True
4080          .cmdPrintReport_raised_focus_dots_img.Visible = False
4090        End Select
4100        .cmdPrintReport_raised_img.Visible = False
4110        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4120        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4130        .cmdPrintReport_raised_img_dis.Visible = False
4140      End If
4150    End With

EXITP:
4160    Exit Sub

ERRH:
4170    Select Case ERR.Number
        Case Else
4180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4190    End Select
4200    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

4310    With Me
4320      .cmdPrintReport_raised_focus_dots_img.Visible = True
4330      .cmdPrintReport_raised_img.Visible = False
4340      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4350      .cmdPrintReport_raised_focus_img.Visible = False
4360      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4370      .cmdPrintReport_raised_img_dis.Visible = False
4380      blnPrintReport_MouseDown = False
4390    End With

EXITP:
4400    Exit Sub

ERRH:
4410    Select Case ERR.Number
        Case Else
4420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4430    End Select
4440    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

4510    With Me
4520      .cmdPrintReport_raised_img.Visible = True
4530      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4540      .cmdPrintReport_raised_focus_img.Visible = False
4550      .cmdPrintReport_raised_focus_dots_img.Visible = False
4560      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4570      .cmdPrintReport_raised_img_dis.Visible = False
4580      blnPrintReport_Focus = False
4590    End With

EXITP:
4600    Exit Sub

ERRH:
4610    Select Case ERR.Number
        Case Else
4620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4630    End Select
4640    Resume EXITP

End Sub

Private Sub cmdRenumber_Click()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRenumber_Click"

4710    With Me

4720      .frmIncomeExpenseCodes_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.

          ' ** blnEdited will get changed to True, if not already so.
4730      RevCode_Renum THIS_NAME, blnEdited, blnMoveAll  ' ** Module Function: modRevCodeFuncs.

4740      If blnEdited = True And .cmdUpdate.Enabled = False Then
4750        .cmdUpdate.Enabled = True
4760      End If

4770    End With

EXITP:
4780    Exit Sub

ERRH:
4790    Select Case ERR.Number
        Case Else
4800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4810    End Select
4820    Resume EXITP

End Sub

Private Sub cmdRenumber_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRenumber_KeyDown"

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
4990          If .cmdAdd.Enabled = True Then
5000            .cmdAdd.SetFocus
5010          ElseIf .cmdEdit.Enabled = True Then
5020            .cmdEdit.SetFocus
5030          ElseIf .cmdUpdate.Enabled = True Then
5040            .cmdUpdate.SetFocus
5050          Else
5060            .cmdClose.SetFocus
5070          End If
5080        End With
5090      End Select
5100    End If

        ' ** Shift keys.
5110    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5120      Select Case intRetVal
          Case vbKeyTab
5130        With Me
5140          intRetVal = 0
5150          If .cmdPrintReport.Enabled = True Then
5160            .cmdPrintReport.SetFocus
5170          Else
5180            .frmIncomeExpenseCodes_Sub.SetFocus
5190            With .frmIncomeExpenseCodes_Sub.Form
5200              .MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.
5210              If .taxcode_type.Visible = True And .taxcode_type.Enabled = True Then
5220                .taxcode_type.SetFocus
5230              Else
5240                .revcode_ACTIVE.SetFocus
5250              End If
5260            End With
5270          End If
5280        End With
5290      End Select
5300    End If

EXITP:
5310    KeyCode = intRetVal
5320    Exit Sub

ERRH:
5330    intRetVal = 0
5340    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5350    Case Else
5360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5370    End Select
5380    Resume EXITP

End Sub

Private Sub cmdAdd_Click()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

        Dim lngRevcodeID As Long
        Dim strDocName As String

5410    gdblCrtRpt_CostTot = 0#  ' ** Borrowing this Public variable.

5420    strDocName = "frmIncomeExpenseCodes_Add"
5430    DoCmd.OpenForm strDocName, , , , acFormAdd, acDialog, THIS_NAME

5440    With Me
5450      If gdblCrtRpt_CostTot > 0 Then
5460        If .chkEditable = False Then
5470          cmdEdit_Click  ' ** Procedure: Above.
5480          DoEvents
5490        End If
5500        blnEdited = True
5510        .cmdUpdate.Enabled = True
5520        .cmdRenumber.Enabled = True
5530        lngRevcodeID = CLng(gdblCrtRpt_CostTot)
5540        .frmIncomeExpenseCodes_Sub.Form.Requery
5550        DoEvents
5560        .frmIncomeExpenseCodes_Sub.Form.MoveRec 0, lngRevcodeID  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.
5570        gdblCrtRpt_CostTot = 0#
5580      End If
5590    End With

EXITP:
5600    Exit Sub

ERRH:
5610    Select Case ERR.Number
        Case Else
5620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5630    End Select
5640    Resume EXITP

End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5720    intShiftDown = (Shift And acShiftMask) > 0
5730    intAltDown = (Shift And acAltMask) > 0
5740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5760      Select Case intRetVal
          Case vbKeyTab
5770        With Me
5780          intRetVal = 0
5790          If .cmdEdit.Enabled = True Then
5800            .cmdEdit.SetFocus
5810          ElseIf .cmdUpdate.Enabled = True Then
5820            .cmdUpdate.SetFocus
5830          Else
5840            .cmdClose.SetFocus
5850          End If
5860        End With
5870      End Select
5880    End If

        ' ** Shift keys.
5890    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5900      Select Case intRetVal
          Case vbKeyTab
5910        With Me
5920          intRetVal = 0
5930          If .cmdRenumber.Enabled = True Then
5940            .cmdRenumber.SetFocus
5950          ElseIf .cmdPrintReport.Enabled = True Then
5960            .cmdPrintReport.SetFocus
5970          Else
5980            .frmIncomeExpenseCodes_Sub.SetFocus
5990            With .frmIncomeExpenseCodes_Sub.Form
6000              .MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.
6010              If .taxcode_type.Visible = True And .taxcode_type.Enabled = True Then
6020                .taxcode_type.SetFocus
6030              Else
6040                .revcode_ACTIVE.SetFocus
6050              End If
6060            End With
6070          End If
6080        End With
6090      End Select
6100    End If

EXITP:
6110    KeyCode = intRetVal
6120    Exit Sub

ERRH:
6130    intRetVal = 0
6140    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6150    Case Else
6160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6170    End Select
6180    Resume EXITP

End Sub

Private Sub cmdEdit_Click()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_Click"

        Dim frm As Access.Form

6210    With Me
6220      Set frm = .frmIncomeExpenseCodes_Sub.Form
6230      Select Case .chkEditable
          Case True
            ' ** It's currently editable, so this turns it off.
6240        If blnEdited = False And .cmdUpdate.Enabled = False Then
              ' ** Toggle only if they've made no changes.
6250          With frm
6260            .opgIncomeExpense.Locked = True
6270            .opgIncomeExpense_optIncome_box.ForeColor = CLR_IEI_BLU_DIS
6280            .Legend_Income_box.BackColor = CLR_IEI_BLU_DIS
6290            .opgIncomeExpense_optExpense_box.ForeColor = CLR_IEE_RED_DIS
6300            .Legend_Expense_box.BackColor = CLR_IEE_RED_DIS
6310            .opgIncomeExpense_def_box.ForeColor = .Detail.BackColor
6320            .opgIncomeExpense_norm_box.ForeColor = .Detail.BackColor
6330            .revcode_DESC.Locked = True
6340            .revcode_DESC.ForeColor = CLR_DISABLED_FG
6350            .revcode_DESC.BackColor = CLR_DISABLED_BG
6360            .revcode_DESC.BackColor = CLR_DISABLED_BG
6370            .revcode_DESC_def_box.ForeColor = CLR_DISABLED_BG
6380            .revcode_DESC_norm_box.ForeColor = CLR_DISABLED_BG
6390            .revcode_SORTORDER.Locked = True
6400            .revcode_SORTORDER.ForeColor = CLR_DISABLED_FG
6410            .revcode_SORTORDER.BackColor = CLR_DISABLED_BG
6420            .revcode_SORTORDER.BackColor = CLR_DISABLED_BG
6430            .revcode_SORTORDER_def_box.ForeColor = CLR_DISABLED_BG
6440            .revcode_SORTORDER_norm_box.ForeColor = CLR_DISABLED_BG
6450            .IsInc_Def.ForeColor = CLR_DISABLED_BG
6460            .Rev_SortOrd_Inc_Def.ForeColor = CLR_DISABLED_FG
6470            .IsInc_Norm.ForeColor = CLR_DISABLED_BG
6480            .Rev_SortOrd_Inc_Norm.ForeColor = CLR_DISABLED_FG
6490            .IsExp_Def.ForeColor = CLR_DISABLED_BG
6500            .Rev_SortOrd_Exp_Def.ForeColor = CLR_DISABLED_FG
6510            .IsExp_Norm.ForeColor = CLR_DISABLED_BG
6520            .Rev_SortOrd_Exp_Norm.ForeColor = CLR_DISABLED_FG
6530            .revcode_ACTIVE.Locked = True
6540            .revcode_active_locked1.ForeColor = CLR_DISABLED_BG
6550            .revcode_active_locked2.ForeColor = CLR_DISABLED_BG
6560            .revcode_active_unlocked1.ForeColor = CLR_DISABLED_BG
6570            .revcode_active_unlocked2.ForeColor = CLR_DISABLED_BG
6580            .IsDef_bg1.ForeColor = .Detail.BackColor
6590            .IsDef_bg2.ForeColor = .Detail.BackColor

6600          End With
6610          .chkEditable = False
6620        Else
6630          Beep
6640        End If
6650      Case False
            ' ** It's currently locked, so this turns it on.
6660        With frm
6670          .opgIncomeExpense.Enabled = True
6680          If .revcode_ID > REVID_INTINC Then
6690            .opgIncomeExpense.Locked = False
6700          End If
6710          .opgIncomeExpense_optIncome_box.ForeColor = CLR_IEI_BLU
6720          .Legend_Income_box.BackColor = CLR_IEI_BLU
6730          .opgIncomeExpense_optExpense_box.ForeColor = CLR_IEE_RED
6740          .Legend_Expense_box.BackColor = CLR_IEE_RED
6750          .opgIncomeExpense_def_box.ForeColor = CLR_DISABLED_BG
6760          .opgIncomeExpense_norm_box.ForeColor = MY_CLR_LTBGE
6770          .revcode_DESC.Enabled = True
6780          If .revcode_ID > REVID_INTINC Then
6790            .revcode_DESC.Locked = False
6800          End If
6810          .revcode_DESC.ForeColor = CLR_BLK
6820          .revcode_DESC.BackColor = CLR_WHT
6830          .revcode_DESC_def_box.ForeColor = CLR_DISABLED_BG
6840          .revcode_DESC_norm_box.ForeColor = CLR_WHT
6850          .revcode_SORTORDER.Enabled = True
6860          If .revcode_ID > REVID_INTINC Then
6870            .revcode_SORTORDER.Locked = False
6880          End If
6890          .revcode_SORTORDER.ForeColor = CLR_BLK
6900          .revcode_SORTORDER.BackColor = CLR_WHT
6910          .revcode_SORTORDER_def_box.ForeColor = CLR_DISABLED_BG
6920          .revcode_SORTORDER_norm_box.ForeColor = CLR_WHT
6930          .IsInc_Def.ForeColor = CLR_DISABLED_BG
6940          .Rev_SortOrd_Inc_Def.ForeColor = CLR_DISABLED_FG
6950          .IsInc_Norm.ForeColor = CLR_WHT
6960          .Rev_SortOrd_Inc_Norm.ForeColor = CLR_BLK
6970          .IsExp_Def.ForeColor = CLR_DISABLED_BG
6980          .Rev_SortOrd_Exp_Def.ForeColor = CLR_DISABLED_FG
6990          .IsExp_Norm.ForeColor = CLR_WHT
7000          .Rev_SortOrd_Exp_Norm.ForeColor = CLR_BLK
7010          .revcode_ACTIVE.Enabled = True
7020          If .revcode_ID > REVID_INTINC Then
7030            .revcode_ACTIVE.Locked = False
7040          End If
7050          .revcode_active_locked1.ForeColor = CLR_DISABLED_BG
7060          .revcode_active_locked2.ForeColor = CLR_DISABLED_BG
7070          .revcode_active_unlocked1.ForeColor = CLR_WHT
7080          .revcode_active_unlocked2.ForeColor = CLR_WHT
7090          .IsDef_bg1.ForeColor = CLR_DISABLED_BG
7100          .IsDef_bg2.ForeColor = CLR_DISABLED_BG
7110        End With
7120        .chkEditable = True
7130      End Select
7140      lngRecsCur = .frmIncomeExpenseCodes_Sub.Form.RecCnt  ' ** Form Function: frmIncomeExpenseCodes_Sub.
7150      If lngRecsCur > 6& Then
7160        .frmIncomeExpenseCodes_Sub.SetFocus
7170      End If
7180      DoEvents
7190    End With  ' ** Me.

EXITP:
7200    Set frm = Nothing
7210    Exit Sub

ERRH:
7220    Select Case ERR.Number
        Case Else
7230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7240    End Select
7250    Resume EXITP

End Sub

Private Sub cmdEdit_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7320    intShiftDown = (Shift And acShiftMask) > 0
7330    intAltDown = (Shift And acAltMask) > 0
7340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7360      Select Case intRetVal
          Case vbKeyTab
7370        With Me
7380          intRetVal = 0
7390          If .cmdUpdate.Enabled = True Then
7400            .cmdUpdate.SetFocus
7410          Else
7420            .cmdClose.SetFocus
7430          End If
7440        End With
7450      End Select
7460    End If

        ' ** Shift keys.
7470    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7480      Select Case intRetVal
          Case vbKeyTab
7490        With Me
7500          intRetVal = 0
7510          If .cmdAdd.Enabled = True Then
7520            .cmdAdd.SetFocus
7530          ElseIf .cmdRenumber.Enabled = True Then
7540            .cmdRenumber.SetFocus
7550          ElseIf .cmdPrintReport.Enabled = True Then
7560            .cmdPrintReport.SetFocus
7570          Else
7580            .frmIncomeExpenseCodes_Sub.SetFocus
7590            With .frmIncomeExpenseCodes_Sub.Form
7600              .MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.
7610              If .taxcode_type.Visible = True And .taxcode_type.Enabled = True Then
7620                .taxcode_type.SetFocus
7630              Else
7640                .revcode_ACTIVE.SetFocus
7650              End If
7660            End With
7670          End If
7680        End With
7690      End Select
7700    End If

EXITP:
7710    KeyCode = intRetVal
7720    Exit Sub

ERRH:
7730    intRetVal = 0
7740    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7750    Case Else
7760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7770    End Select
7780    Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

7810    With Me

7820      DoCmd.Hourglass True
7830      DoEvents

7840      .frmIncomeExpenseCodes_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.

7850      RevCode_Update THIS_NAME, blnEdited  ' ** Module Function: modRevCodeFuncs.

7860      .frmIncomeExpenseCodes_Sub.SetFocus
7870      .frmIncomeExpenseCodes_Sub.Form.FocusHolder.SetFocus

7880      DoCmd.Hourglass False

7890    End With

EXITP:
7900    Exit Sub

ERRH:
7910    DoCmd.Hourglass False
7920    Select Case ERR.Number
        Case Else
7930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7940    End Select
7950    Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8020    intShiftDown = (Shift And acShiftMask) > 0
8030    intAltDown = (Shift And acAltMask) > 0
8040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8060      Select Case intRetVal
          Case vbKeyTab
8070        With Me
8080          intRetVal = 0
8090          .cmdClose.SetFocus
8100        End With
8110      End Select
8120    End If

        ' ** Shift keys.
8130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8140      Select Case intRetVal
          Case vbKeyTab
8150        With Me
8160          intRetVal = 0
8170          If .cmdEdit.Enabled = True Then
8180            .cmdEdit.SetFocus
8190          ElseIf .cmdAdd.Enabled = True Then
8200            .cmdAdd.SetFocus
8210          ElseIf .cmdRenumber.Enabled = True Then
8220            .cmdRenumber.SetFocus
8230          ElseIf .cmdPrintReport.Enabled = True Then
8240            .cmdPrintReport.SetFocus
8250          Else
8260            .frmIncomeExpenseCodes_Sub.SetFocus
8270            With .frmIncomeExpenseCodes_Sub.Form
8280              .MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmIncomeExpenseCodes_Sub.
8290              If .taxcode_type.Visible = True And .taxcode_type.Enabled = True Then
8300                .taxcode_type.SetFocus
8310              Else
8320                .revcode_ACTIVE.SetFocus
8330              End If
8340            End With
8350          End If
8360        End With
8370      End Select
8380    End If

EXITP:
8390    KeyCode = intRetVal
8400    Exit Sub

ERRH:
8410    intRetVal = 0
8420    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8430    Case Else
8440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8450    End Select
8460    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

8510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8520    Exit Sub

ERRH:
8530    Select Case ERR.Number
        Case Else
8540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8550    End Select
8560    Resume EXITP

End Sub

Public Sub ClosingForm(blnClose As Boolean)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "ClosingForm"

8610    gblnClosing = blnClose

EXITP:
8620    Exit Sub

ERRH:
8630    Select Case ERR.Number
        Case Else
8640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8650    End Select
8660    Resume EXITP

End Sub

Public Sub EditingForm(blnEdit As Boolean)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "EditingForm"

8710    blnEdited = blnEdit

EXITP:
8720    Exit Sub

ERRH:
8730    Select Case ERR.Number
        Case Else
8740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8750    End Select
8760    Resume EXITP

End Sub
