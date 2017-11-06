VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStatementParameters_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmStatementParameters_Print"

'VGC 12/27/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   No:               N {cmdNo}
' **   Print:            P {cmdPrint}
' **   Preview:          W {cmdPreview}
' **   Yes:              Y {cmdYes}

' ** When Previewed, this form is canceled.
' ** When Printed, this form remains.

Private blnGoingToReport As Boolean
Private strCallingForm As String
Private blnFromPreview As Boolean
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           intPos01 = InStr(Trim(.OpenArgs), "~")
150           If intPos01 > 0 Then

160             strCallingForm = Left(Trim(.OpenArgs), (intPos01 - 1))
170             .Msg_lbl.Caption = Mid(Trim(.OpenArgs), (intPos01 + 1))

180             .PictureData = .Form_StatementParameters_Print_img

190             blnFromPreview = False

200             If Forms(strCallingForm).PrintAnnual_chk = True Then
210               .cmdPreview.Visible = False
220               .cmdPreview_lbl.Visible = False
230               .cmdPreview_box.Visible = False
240               .cmdPrint.Visible = False
250             End If

260             .cmdYes.SetFocus

270           Else
280             Cancel = -1
290           End If
300         Else
310           Cancel = -1
320         End If
330       Else
340         Cancel = -1
350       End If
360     End With

370     If Cancel = -1 Then
380       Beep
390       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
400     End If

EXITP:
410     Exit Sub

ERRH:
420     Select Case ERR.Number
        Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         intRetVal = 0
580         cmdNo_Click  ' ** Procedure: Below.
590       End Select
600     End If

EXITP:
610     KeyCode = intRetVal
620     Exit Sub

ERRH:
630     intRetVal = 0
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub Form_Timer()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

710     With Me
720       .TimerInterval = 0&
730       If gblnGoToReport = True Then
740         DoCmd.Hourglass True  ' ** Make sure it's still running.
750         DoEvents
760         Select Case blnGoingToReport
            Case True
770           blnGoingToReport = False
780           gblnGoToReport = False
790           .GoToReport_arw_spp_print_img.Visible = False
800           .cmdNo.SetFocus
810           DoCmd.Hourglass False
820         Case False
830           blnGoingToReport = True
840           .GoToReport_arw_spp_print_img.Visible = True
850           .cmdPrint.SetFocus
860           DoEvents
870           .cmdYes.Enabled = False
880           Forms(strCallingForm).GTREmblem_Off  ' ** Form Procedure: frmStatementParameters.
890           Forms(strCallingForm).TimerInterval = 0&
900           Forms(strCallingForm).GoToReport_arw_sp_printall_img.Visible = False
910           Forms(strCallingForm).cmdPrintStatement_Single.Visible = True
920           Beep
930           .TimerInterval = (GTR_WAIT * 2&)
940         End Select
950       End If
960     End With

EXITP:
970     Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdYes_Click()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdYes_Click"

1110    DoCmd.Close acForm, THIS_NAME
1120    Forms(strCallingForm).SetPrintStatements True  ' ** Form Procedure: frmStatementParameters.
1130    DoCmd.SelectObject acForm, strCallingForm, False
1140    DoEvents

EXITP:
1150    Exit Sub

ERRH:
1160    Select Case ERR.Number
        Case Else
1170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1180    End Select
1190    Resume EXITP

End Sub

Private Sub cmdNo_Click()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdNo_Click"

1210    DoCmd.Close acForm, THIS_NAME
1220    Forms(strCallingForm).SetPrintStatements False  ' ** Form Procedure: frmStatementParameters.
1230    If blnFromPreview = True Then
1240      DoCmd.SelectObject acReport, "rptStatement_Qualified", False
1250      DoCmd.Maximize
1260      DoCmd.RunCommand acCmdFitToWindow
1270      DoCmd.Hourglass False
1280    Else
1290      DoCmd.SelectObject acForm, strCallingForm, False
1300      DoEvents
1310    End If

EXITP:
1320    Exit Sub

ERRH:
1330    Select Case ERR.Number
        Case Else
1340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1350    End Select
1360    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

        ' ** In order to Preview this list to the screen, this
        ' ** form needs to be closed, so it'll have to be Canceled.
1410    blnFromPreview = True
1420    Forms(strCallingForm).AcctsSchedRpt_Set True  ' ** Form Procedure: frmStatementParameters.
1430    strDocName = "rptStatement_Qualified"
1440    DoCmd.OpenReport strDocName, acViewPreview
1450    DoCmd.Maximize
1460    DoCmd.RunCommand acCmdFitToWindow
1470    cmdNo_Click  ' ** Procedure: Above.

EXITP:
1480    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

1610    strDocName = "rptStatement_Qualified"
1620    If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
1630      DoCmd.OpenReport strDocName, acViewPreview
1640    Else
          '##GTR_Ref: rptStatement_Qualified
1650      DoCmd.OpenReport strDocName, acViewNormal
1660    End If

EXITP:
1670    Exit Sub

ERRH:
1680    Select Case ERR.Number
        Case Else
1690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1700    End Select
1710    Resume EXITP

End Sub
