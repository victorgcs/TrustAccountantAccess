VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Post_Clear_Multi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Post_Clear_Multi"

'VGC 12/06/2012: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   OK:               O {cmdOK}
' **   Users:            U {lbxUser}

Private strMode As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strCallingForm = Trim(.OpenArgs)
150           intPos01 = InStr(strCallingForm, "~")
160           If intPos01 > 0 Then

170             strMode = Mid(strCallingForm, (intPos01 + 1))
180             strCallingForm = Left(strCallingForm, (intPos01 - 1))

190             .PictureData = .Form_Menu_Post_Clear_Multi_img

200             Select Case strMode
                Case "PostJournal"
210               .Caption = "Post Journal Entries By User"
220               .Msg_lbl1.Visible = True
230               .Msg_lbl2.Visible = False
240               .Msg_lbl3.Visible = False
250               .Msg_lbl4.Visible = False
260               .lbxUser_lbl.Caption = "These Users currently have" & vbCrLf & _
                    "Journal Entries to be Posted." & vbCrLf & _
                    "Make a selection to continue."
270               .lbxUser.RowSource = "qryJournal_User_01d"
280             Case "ClearJournal"
290               .Caption = "Clear Journal Entries By User"
300               .Msg_lbl1.Visible = False
310               .Msg_lbl2.Visible = True
320               .Msg_lbl3.Visible = False
330               .Msg_lbl4.Visible = False
340               .lbxUser_lbl.Caption = "These Users currently have" & vbCrLf & _
                    "Journal entries. Make a" & vbCrLf & _
                    "selection to clear the entries."
350               .lbxUser.RowSource = "qryJournal_Clear_01f"
360             Case "ClearUncommitted"
370               .Caption = "Clear Uncommitted Entries By User"
380               .Msg_lbl1.Visible = False
390               .Msg_lbl2.Visible = False
400               .Msg_lbl3.Visible = True
410               .Msg_lbl4.Visible = False
420               .lbxUser_lbl.Caption = "These Users currently have" & vbCrLf & _
                    "uncommitted entries. Make a" & vbCrLf & _
                    "selection to clear the entries."
430               .lbxUser.RowSource = "qryJournal_Clear_01l"
440             Case "ClearChecks"
450               .Caption = "Clear Journal Checks By User"
460               .Msg_lbl1.Visible = False
470               .Msg_lbl2.Visible = False
480               .Msg_lbl3.Visible = False
490               .Msg_lbl4.Visible = True
500               .lbxUser_lbl.Caption = "These Users currently have" & vbCrLf & _
                    "checks in the Journal. Make a" & vbCrLf & _
                    "selection to clear the checks."
510               .lbxUser.RowSource = "qryPrintChecks_11f"
520             End Select
530             .lbxUser.Requery
540             .lbxUser = vbNullString

550             Beep
                'gstrJournalUser = vbNullString  ' ** Don't clear this on this form!
560             DoCmd.Hourglass False

570           Else
580             Cancel = -1
590           End If
600         Else
610           Cancel = -1
620         End If
630       Else
640         Cancel = -1
650       End If
660     End With

670     If Cancel = -1 Then
680       Beep
690       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
700     End If

EXITP:
710     Exit Sub

ERRH:
720     Select Case ERR.Number
        Case Else
730       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
740     End Select
750     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        Const THIS_PROC As String = "Form_KeyDown"

        ' ** Use bit masks to determine which key was pressed.
810     intShiftDown = (Shift And acShiftMask) > 0
820     intAltDown = (Shift And acAltMask) > 0
830     intCtrlDown = (Shift And acCtrlMask) > 0

840     intRetVal = KeyCode

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyEscape
870         intRetVal = 0
880         cmdCancel_Click  ' ** Procedure: Below.
890       End Select
900     End If

EXITP:
910     KeyCode = intRetVal
920     Exit Sub

ERRH:
930     intRetVal = 0
940     Select Case ERR.Number
        Case Else
950       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
960     End Select
970     Resume EXITP

End Sub

Private Sub Form_Close()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Close"

1010    If strCallingForm <> vbNullString Then
1020      DoCmd.SelectObject acForm, strCallingForm, False
1030    End If

EXITP:
1040    Exit Sub

ERRH:
1050    Select Case ERR.Number
        Case Else
1060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1070    End Select
1080    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

1110    With Me
1120      If IsNull(.lbxUser) = False Then
1130        If Trim(.lbxUser) <> vbNullString Then
1140          DoCmd.Hourglass True
1150          Select Case strMode
              Case "ClearUncommitted"
1160            gstrCrtRpt_NetLoss = Trim(.lbxUser)  ' ** Borrowing this variable from the Court Reports.
1170          Case Else
1180            gstrJournalUser = Trim(.lbxUser)
1190          End Select
1200          DoEvents
1210          DoCmd.Close acForm, THIS_NAME
1220        Else
1230          MsgBox "Please choose a User, or click Cancel to exit.", vbInformation + vbOKOnly, "Entry Required"
1240          .lbxUser.SetFocus
1250        End If
1260      Else
1270        MsgBox "Please choose a User, or click Cancel to exit.", vbInformation + vbOKOnly, "Entry Required"
1280        .lbxUser.SetFocus
1290      End If
1300    End With

EXITP:
1310    Exit Sub

ERRH:
1320    Select Case ERR.Number
        Case Else
1330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1340    End Select
1350    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

1410    Select Case strMode
        Case "ClearUncommitted"
1420      gstrCrtRpt_NetLoss = vbNullString  ' ** Borrowing this variable from the Court Reports.
1430    Case Else
1440      gstrJournalUser = vbNullString
1450    End Select
1460    DoCmd.Close acForm, THIS_NAME

EXITP:
1470    Exit Sub

ERRH:
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub lbxUser_DblClick(Cancel As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxUser_DblClick"

1610    With Me
1620      If IsNull(.lbxUser) = False Then
1630        If Trim(lbxUser) <> vbNullString Then
1640          cmdOK_Click  ' ** Procedure: Above.
1650        End If
1660      End If
1670    End With

EXITP:
1680    Exit Sub

ERRH:
1690    Select Case ERR.Number
        Case Else
1700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1710    End Select
1720    Resume EXITP

End Sub
