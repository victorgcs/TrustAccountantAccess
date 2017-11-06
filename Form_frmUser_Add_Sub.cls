VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUser_Add_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmUser_Add_Sub"

'VGC 12/18/2012: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel on frmUser_Add}
' **   Password Help:    H {cmdHelp on frmUser_Add}
' **   OK:               O {cmdOK on frmUser_Add}
' **   User Name:        U {Username}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmUser_Add" Then

160         DoCmd.RunCommand acCmdRecordsGoToNew

170       Else
180         Cancel = -1
190       End If
200     End With

210     If Cancel = -1 Then
220       Beep
230       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
240     End If

EXITP:
250     Exit Sub

ERRH:
260     Select Case ERR.Number
        Case Else
270       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
280     End Select
290     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

300   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

310     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
320     intShiftDown = (Shift And acShiftMask) > 0
330     intAltDown = (Shift And acAltMask) > 0
340     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
350     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
360       Select Case intRetVal
          Case vbKeyEscape
370         With Me
380           intRetVal = 0
390           .Parent.cmdCancel_Click  ' ** Form Procedure: frmUser_Add.
400         End With
410       End Select
420     End If

EXITP:
430     KeyCode = intRetVal
440     Exit Sub

ERRH:
450     intRetVal = 0
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

510     With Me
520       .Parent.Detail_MouseMove Button, Shift, X, Y  ' ** Form Procedure: frmUser_Add.
530     End With

EXITP:
540     Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub Username_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Username_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shift keys.
650     If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
660       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
670         With Me
680           intRetVal = 0
690           DoCmd.SelectObject acForm, .Parent.Name, False
700           .Parent.cmdCancel.SetFocus
710         End With
720       End Select
730     End If

EXITP:
740     KeyCode = intRetVal
750     Exit Sub

ERRH:
760     Select Case ERR.Number
        Case Else
770       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
780     End Select
790     Resume EXITP

End Sub

Private Sub Primary_Group_Enter()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Primary_Group_Enter"

810     With Me
820       If IsNull(.Primary_Group) = True Then
830         ExpandCombo  ' ** Module Procedure: modUtilities.
840       ElseIf .Primary_Group = vbNullString Then
850         ExpandCombo  ' ** Module Procedure: modUtilities.
860       End If
870     End With

EXITP:
880     Exit Sub

ERRH:
890     Select Case ERR.Number
        Case Else
900       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
910     End Select
920     Resume EXITP

End Sub

Private Sub NewPassword_GotFocus()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "NewPassword_GotFocus"

1010    With Me
1020      .NewPassword.SelStart = 0
1030      .NewPassword.SelLength = Len(Trim(.NewPassword.text))
1040    End With

EXITP:
1050    Exit Sub

ERRH:
1060    Select Case ERR.Number
        Case Else
1070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1080    End Select
1090    Resume EXITP

End Sub

Private Sub VerifyNewPassword_GotFocus()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "VerifyNewPassword_GotFocus"

1110    With Me
1120      .VerifyNewPassword.SelStart = 0
1130      .VerifyNewPassword.SelLength = Len(Trim(.VerifyNewPassword.text))
1140    End With

EXITP:
1150    Exit Sub

ERRH:
1160    Select Case ERR.Number
        Case Else
1170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1180    End Select
1190    Resume EXITP

End Sub

Private Sub VerifyNewPassword_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "VerifyNewPassword_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1270        With Me
1280          intRetVal = 0
1290          DoCmd.SelectObject acForm, .Parent.Name, False
1300          .Parent.cmdOK.SetFocus
1310        End With
1320      End Select
1330    End If

EXITP:
1340    KeyCode = intRetVal
1350    Exit Sub

ERRH:
1360    Select Case ERR.Number
        Case Else
1370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1380    End Select
1390    Resume EXITP

End Sub

Public Sub MoveRec(strUsername As String)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

1410    With Me
1420      Set rst = .RecordsetClone
1430      With rst
1440        .FindFirst "[Username] = '" & strUsername & "'"
1450        Select Case .NoMatch
            Case True
1460          Beep
1470          MsgBox "Record not found!"
1480        Case False
1490          Me.Bookmark = .Bookmark
1500        End Select
1510        .Close
1520      End With
1530    End With

EXITP:
1540    Set rst = Nothing
1550    Exit Sub

ERRH:
1560    Select Case ERR.Number
        Case Else
1570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1580    End Select
1590    Resume EXITP

End Sub
