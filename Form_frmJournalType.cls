VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournalType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournalType"

'VGC 09/26/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Print Report:     P {cmdPrintReport}
' **   Exit:             X {cmdClose}

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
920         .frmJournalType_Sub.SetFocus
930       End If
940       If gblnGoToReport = True Then
950         Select Case blnGoingToReport
            Case True
960           DoCmd.Hourglass False  ' ** It should already be off.
970           gblnGoToReport = False
980           blnGoingToReport = False
990           .GoToReport_arw_pr_img.Visible = False
1000          .cmdPrintReport_lbl2.Visible = True
1010          DoEvents
1020        Case False
1030          DoCmd.Hourglass True  ' ** Make sure it's still running.
1040          DoEvents
1050          blnGoingToReport = True
1060          .cmdPrintReport_lbl2.Visible = False
1070          .GoToReport_arw_pr_img.Visible = True
1080          .cmdPrintReport.SetFocus
1090          Beep
1100          DoCmd.Hourglass False
1110          DoEvents
1120          .TimerInterval = (GTR_WAIT * 2&)
1130        End Select
1140      Else
1150        DoCmd.Hourglass False
1160      End If
1170    End With

EXITP:
1180    Exit Sub

ERRH:
1190    Select Case ERR.Number
        Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1210    End Select
1220    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim blnRetVal As Boolean

1310    With Me

1320      DoCmd.Hourglass True
1330      DoEvents

1340      blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

1350      gstrReturningForm = THIS_NAME
1360      gblnSetFocus = True
1370      strDocName = "frmMenu_SystemTable"
1380      DoCmd.OpenForm strDocName, , , , , , "frmMenu_Maintenance"

1390    End With

EXITP:
1400    Exit Sub

ERRH:
1410    DoCmd.Hourglass False
1420    Select Case ERR.Number
        Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1510    With Me
1520      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
1530        Select Case blnPrintReport_Focus
            Case True
1540          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
1550          .cmdPrintReport_raised_img.Visible = False
1560        Case False
1570          .cmdPrintReport_raised_img.Visible = True
1580          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
1590        End Select
1600        .cmdPrintReport_raised_focus_img.Visible = False
1610        .cmdPrintReport_raised_focus_dots_img.Visible = False
1620        .cmdPrintReport_sunken_focus_dots_img.Visible = False
1630        .cmdPrintReport_raised_img_dis.Visible = False
1640      End If
1650    End With

EXITP:
1660    Exit Sub

ERRH:
1670    Select Case ERR.Number
        Case Else
1680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1690    End Select
1700    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

1810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Public Sub cmdClose_Click()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

1910    DoCmd.Close acForm, THIS_NAME

EXITP:
1920    Exit Sub

ERRH:
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2020    intShiftDown = (Shift And acShiftMask) > 0
2030    intAltDown = (Shift And acAltMask) > 0
2040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2060      Select Case intRetVal
          Case vbKeyTab
2070        With Me
2080          intRetVal = 0
2090          .frmJournalType_Sub.SetFocus
2100          .frmJournalType_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournalType_Sub.
2110          .frmJournalType_Sub.Form.FocusHolder.SetFocus
2120        End With
2130      End Select
2140    End If

        ' ** Shift keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2160      Select Case intRetVal
          Case vbKeyTab
2170        With Me
2180          intRetVal = 0
2190          If .cmdPrintReport.Enabled = True Then
2200            .cmdPrintReport.SetFocus
2210          Else
2220            .frmJournalType_Sub.SetFocus
2230            .frmJournalType_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournalType_Sub.
2240            .frmJournalType_Sub.Form.FocusHolder.SetFocus
2250          End If
2260        End With
2270      End Select
2280    End If

EXITP:
2290    KeyCode = intRetVal
2300    Exit Sub

ERRH:
2310    intRetVal = 0
2320    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2330    Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

2410    With Me
2420      blnPrintReport_Focus = True
2430      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2440      .cmdPrintReport_raised_img.Visible = False
2450      .cmdPrintReport_raised_focus_img.Visible = False
2460      .cmdPrintReport_raised_focus_dots_img.Visible = False
2470      .cmdPrintReport_sunken_focus_dots_img.Visible = False
2480      .cmdPrintReport_raised_img_dis.Visible = False
2490    End With

EXITP:
2500    Exit Sub

ERRH:
2510    Select Case ERR.Number
        Case Else
2520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2530    End Select
2540    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

2610    With Me
2620      blnPrintReport_MouseDown = True
2630      .cmdPrintReport_sunken_focus_dots_img.Visible = True
2640      .cmdPrintReport_raised_img.Visible = False
2650      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2660      .cmdPrintReport_raised_focus_img.Visible = False
2670      .cmdPrintReport_raised_focus_dots_img.Visible = False
2680      .cmdPrintReport_raised_img_dis.Visible = False
2690    End With

EXITP:
2700    Exit Sub

ERRH:
2710    Select Case ERR.Number
        Case Else
2720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2730    End Select
2740    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

2810    With Me
2820      strDocName = "rptListOfJournalTypes"
2830      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
2840        DoCmd.OpenReport strDocName, acViewPreview
2850        DoCmd.Maximize
2860        DoCmd.RunCommand acCmdFitToWindow
2870        .Visible = False
2880      Else
            '##GTR_Ref: rptListOfJournalTypes
2890        DoCmd.OpenReport strDocName, acViewNormal
2900      End If
2910    End With

EXITP:
2920    Exit Sub

ERRH:
2930    Me.Visible = True
2940    DoCmd.Restore
2950    Select Case ERR.Number
        Case Else
2960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2970    End Select
2980    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3020    intShiftDown = (Shift And acShiftMask) > 0
3030    intAltDown = (Shift And acAltMask) > 0
3040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3060      Select Case intRetVal
          Case vbKeyTab
3070        With Me
3080          intRetVal = 0
3090          .cmdClose.SetFocus
3100        End With
3110      End Select
3120    End If

        ' ** Shift keys.
3130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3140      Select Case intRetVal
          Case vbKeyTab
3150        With Me
3160          intRetVal = 0
3170          .frmJournalType_Sub.SetFocus
3180          .frmJournalType_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournalType_Sub.
3190          .frmJournalType_Sub.Form.FocusHolder.SetFocus
3200        End With
3210      End Select
3220    End If

EXITP:
3230    KeyCode = intRetVal
3240    Exit Sub

ERRH:
3250    intRetVal = 0
3260    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3270    Case Else
3280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290    End Select
3300    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

3410    With Me
3420      If blnPrintReport_MouseDown = False Then
3430        Select Case blnPrintReport_Focus
            Case True
3440          .cmdPrintReport_raised_focus_dots_img.Visible = True
3450          .cmdPrintReport_raised_focus_img.Visible = False
3460        Case False
3470          .cmdPrintReport_raised_focus_img.Visible = True
3480          .cmdPrintReport_raised_focus_dots_img.Visible = False
3490        End Select
3500        .cmdPrintReport_raised_img.Visible = False
3510        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3520        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3530        .cmdPrintReport_raised_img_dis.Visible = False
3540      End If
3550    End With

EXITP:
3560    Exit Sub

ERRH:
3570    Select Case ERR.Number
        Case Else
3580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3590    End Select
3600    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

3710    With Me
3720      .cmdPrintReport_raised_focus_dots_img.Visible = True
3730      .cmdPrintReport_raised_img.Visible = False
3740      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3750      .cmdPrintReport_raised_focus_img.Visible = False
3760      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3770      .cmdPrintReport_raised_img_dis.Visible = False
3780      blnPrintReport_MouseDown = False
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

Private Sub cmdPrintReport_LostFocus()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

3910    With Me
3920      .cmdPrintReport_raised_img.Visible = True
3930      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3940      .cmdPrintReport_raised_focus_img.Visible = False
3950      .cmdPrintReport_raised_focus_dots_img.Visible = False
3960      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3970      .cmdPrintReport_raised_img_dis.Visible = False
3980      blnPrintReport_Focus = False
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

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

4110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub
