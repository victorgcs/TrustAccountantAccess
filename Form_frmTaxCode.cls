VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTaxCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmTaxCode"

'VGC 09/23/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Description:      D {taxcode_description on frmTaxCode_Sub}
' **   Print Report:     P {cmdPrintReport}
' **   Tax Code:         T {taxcode_inc, taxcode_ded on frmTaxCode_Sub}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmTaxCode_Sub}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Toggle RevCodes:  R {ShowRevCodes on frmTaxCode_Sub}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, lngPrintReportLbl_Left As Long
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

180         Else
190           Cancel = -1
200         End If
210       Else
220         Cancel = -1
230       End If
240     End With

250     If Cancel = -1 Then
260       Beep
270       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
280     End If

EXITP:
290     Exit Sub

ERRH:
300     Select Case ERR.Number
        Case Else
310       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
320     End Select
330     Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Description:      D {taxcode_description on frmTaxCode_Sub}
        ' **   Tax Code:         T {taxcode_inc, taxcode_ded on frmTaxCode_Sub}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmTaxCode_Sub}

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
          Case vbKeyD
530         With Me
540           intRetVal = 0
550           .frmTaxCode_Sub.SetFocus
560           .frmTaxCode_Sub.Form.taxcode_description.SetFocus
570         End With
580       Case vbKeyT
590         With Me
600           intRetVal = 0
610           .frmTaxCode_Sub.SetFocus
620           Select Case .frmTaxCode_Sub.Form.taxcode_type
              Case TAXTYP_INC
630             .frmTaxCode_Sub.Form.taxcode_inc.SetFocus
640           Case TAXTYP_DED
650             .frmTaxCode_Sub.Form.taxcode_ded.SetFocus
660           End Select
670         End With
680       Case vbKeyX
690         intRetVal = 0
700         cmdClose_Click  ' ** Procedure: Below.
710       End Select
720     End If

        ' ** Ctrl keys.
730     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
740       Select Case intRetVal
          Case vbKeyS
750         With Me
760           intRetVal = 0
770           .frmTaxCode_Sub.SetFocus
780           .frmTaxCode_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmTaxCode_Sub.
790         End With
800       End Select
810     End If

        ' ** Ctrl-Shift keys.
820     If intCtrlDown And (Not intAltDown) And intShiftDown Then
830       Select Case intRetVal
          Case vbKeyF
840         With Me
850           intRetVal = 0
860           .FocusHolder.SetFocus
870         End With
880       Case vbKeyR
890         With Me
900           intRetVal = 0
910           .frmTaxCode_Sub.Form.ShowRevCodes  ' ** Form Procedure: frmTaxCode_Sub.
920         End With
930       End Select
940     End If

EXITP:
950     KeyCode = intRetVal
960     Exit Sub

ERRH:
970     intRetVal = 0
980     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
990     Case Else
1000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1010    End Select
1020    Resume EXITP

End Sub

Private Sub Form_Timer()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1110    With Me
1120      .TimerInterval = 0&
1130      If gblnSetFocus = True Then
1140        gblnSetFocus = False
1150        strFormName = vbNullString
1160  On Error Resume Next
1170        Do Until strFormName = THIS_NAME
1180          DoCmd.SelectObject acForm, THIS_NAME, False
1190          strFormName = Screen.ActiveForm.Name
1200        Loop
1210  On Error GoTo ERRH
1220        .frmTaxCode_Sub.SetFocus
1230      End If
1240      If gblnGoToReport = True Then
1250        Select Case blnGoingToReport
            Case True
1260          DoCmd.Hourglass False  ' ** It should already be off.
1270          gblnGoToReport = False
1280          blnGoingToReport = False
1290          .GoToReport_arw_pr_img.Visible = False
1300          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1310          DoEvents
1320        Case False
1330          DoCmd.Hourglass True  ' ** Make sure it's still running.
1340          DoEvents
1350          blnGoingToReport = True
1360          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1370          .GoToReport_arw_pr_img.Visible = True
1380          .cmdPrintReport.SetFocus
1390          Beep
1400          DoCmd.Hourglass False
1410          DoEvents
1420          .TimerInterval = (GTR_WAIT * 2&)
1430        End Select
1440      Else
1450        DoCmd.Hourglass False
1460      End If
1470    End With

EXITP:
1480    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim blnRetVal As Boolean

1610    With Me

1620      DoCmd.Hourglass True
1630      DoEvents

1640      blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

1650      gstrReturningForm = THIS_NAME
1660      gblnSetFocus = True
1670      strDocName = "frmMenu_SystemTable"
1680      DoCmd.OpenForm strDocName, , , , , , "frmMenu_Maintenance"

1690    End With

EXITP:
1700    Exit Sub

ERRH:
1710    DoCmd.Hourglass False
1720    Select Case ERR.Number
        Case Else
1730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1740    End Select
1750    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1810    With Me
1820      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
1830        Select Case blnPrintReport_Focus
            Case True
1840          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
1850          .cmdPrintReport_raised_img.Visible = False
1860        Case False
1870          .cmdPrintReport_raised_img.Visible = True
1880          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
1890        End Select
1900        .cmdPrintReport_raised_focus_img.Visible = False
1910        .cmdPrintReport_raised_focus_dots_img.Visible = False
1920        .cmdPrintReport_sunken_focus_dots_img.Visible = False
1930        .cmdPrintReport_raised_img_dis.Visible = False
1940      End If
1950    End With

EXITP:
1960    Exit Sub

ERRH:
1970    Select Case ERR.Number
        Case Else
1980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1990    End Select
2000    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2210    DoCmd.Close acForm, THIS_NAME

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

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
2390          .frmTaxCode_Sub.SetFocus
2400          .frmTaxCode_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmTaxCode_Sub.
2410          .frmTaxCode_Sub.Form.FocusHolder.SetFocus
2420        End With
2430      End Select
2440    End If

        ' ** Shift keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2460      Select Case intRetVal
          Case vbKeyTab
2470        With Me
2480          intRetVal = 0
2490          If .cmdPrintReport.Enabled = True Then
2500            .cmdPrintReport.SetFocus
2510          Else
2520            .frmTaxCode_Sub.SetFocus
2530            .frmTaxCode_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmTaxCode_Sub.
2540            .frmTaxCode_Sub.Form.FocusHolder.SetFocus
2550          End If
2560        End With
2570      End Select
2580    End If

EXITP:
2590    KeyCode = intRetVal
2600    Exit Sub

ERRH:
2610    intRetVal = 0
2620    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2630    Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

2710    With Me
2720      blnPrintReport_Focus = True
2730      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2740      .cmdPrintReport_raised_img.Visible = False
2750      .cmdPrintReport_raised_focus_img.Visible = False
2760      .cmdPrintReport_raised_focus_dots_img.Visible = False
2770      .cmdPrintReport_sunken_focus_dots_img.Visible = False
2780      .cmdPrintReport_raised_img_dis.Visible = False
2790    End With

EXITP:
2800    Exit Sub

ERRH:
2810    Select Case ERR.Number
        Case Else
2820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2830    End Select
2840    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

2910    With Me
2920      blnPrintReport_MouseDown = True
2930      .cmdPrintReport_sunken_focus_dots_img.Visible = True
2940      .cmdPrintReport_raised_img.Visible = False
2950      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2960      .cmdPrintReport_raised_focus_img.Visible = False
2970      .cmdPrintReport_raised_focus_dots_img.Visible = False
2980      .cmdPrintReport_raised_img_dis.Visible = False
2990    End With

EXITP:
3000    Exit Sub

ERRH:
3010    Select Case ERR.Number
        Case Else
3020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3030    End Select
3040    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String, strSortNow As String

3110    With Me
3120      strSortNow = .frmTaxCode_Sub.Form.SortNow_Get  ' ** Form Function: frmTaxCode_Sub.
3130      If strSortNow = "[taxcode_type], [taxcode_order]" Then
3140        strDocName = "rptListOfTaxCodes"
3150      Else
3160        strDocName = "rptListOfTaxCodes_Sort"
3170      End If
3180      If gblnDev_Debug Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
3190        DoCmd.OpenReport strDocName, acViewPreview
3200        DoCmd.Maximize
3210        DoCmd.RunCommand acCmdFitToWindow
3220        .Visible = False
3230      Else
            '##GTR_Ref: rptListOfTaxCodes
3240        DoCmd.OpenReport strDocName, acViewNormal
3250      End If
3260    End With

EXITP:
3270    Exit Sub

ERRH:
3280    Me.Visible = True
3290    DoCmd.Restore
3300    Select Case ERR.Number
        Case Else
3310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3320    End Select
3330    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3420    intShiftDown = (Shift And acShiftMask) > 0
3430    intAltDown = (Shift And acAltMask) > 0
3440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3460      Select Case intRetVal
          Case vbKeyTab
3470        With Me
3480          intRetVal = 0
3490          .cmdClose.SetFocus
3500        End With
3510      End Select
3520    End If

        ' ** Shift keys.
3530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3540      Select Case intRetVal
          Case vbKeyTab
3550        With Me
3560          intRetVal = 0
3570          .frmTaxCode_Sub.SetFocus
3580          .frmTaxCode_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmTaxCode_Sub.
3590          .frmTaxCode_Sub.Form.FocusHolder.SetFocus
3600        End With
3610      End Select
3620    End If

EXITP:
3630    KeyCode = intRetVal
3640    Exit Sub

ERRH:
3650    intRetVal = 0
3660    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3670    Case Else
3680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3690    End Select
3700    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

3810    With Me
3820      If blnPrintReport_MouseDown = False Then
3830        Select Case blnPrintReport_Focus
            Case True
3840          .cmdPrintReport_raised_focus_dots_img.Visible = True
3850          .cmdPrintReport_raised_focus_img.Visible = False
3860        Case False
3870          .cmdPrintReport_raised_focus_img.Visible = True
3880          .cmdPrintReport_raised_focus_dots_img.Visible = False
3890        End Select
3900        .cmdPrintReport_raised_img.Visible = False
3910        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3920        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3930        .cmdPrintReport_raised_img_dis.Visible = False
3940      End If
3950    End With

EXITP:
3960    Exit Sub

ERRH:
3970    Select Case ERR.Number
        Case Else
3980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3990    End Select
4000    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

4110    With Me
4120      .cmdPrintReport_raised_focus_dots_img.Visible = True
4130      .cmdPrintReport_raised_img.Visible = False
4140      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4150      .cmdPrintReport_raised_focus_img.Visible = False
4160      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4170      .cmdPrintReport_raised_img_dis.Visible = False
4180      blnPrintReport_MouseDown = False
4190    End With

EXITP:
4200    Exit Sub

ERRH:
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

4310    With Me
4320      .cmdPrintReport_raised_img.Visible = True
4330      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4340      .cmdPrintReport_raised_focus_img.Visible = False
4350      .cmdPrintReport_raised_focus_dots_img.Visible = False
4360      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4370      .cmdPrintReport_raised_img_dis.Visible = False
4380      blnPrintReport_Focus = False
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

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

4510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub
