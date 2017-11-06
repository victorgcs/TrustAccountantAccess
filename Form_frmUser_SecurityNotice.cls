VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUser_SecurityNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'VGC 12/07/2012: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   OK:               O {cmdOK}
' **   Username:         U {UserPass}

' ** Use this in Immediate Window to test:
' **   DoCmd.OpenForm "frmUser_SecurityNotice", acNormal, , , , acDialog, "BillL           BillLar1       " & vbCrLf & "MaryS           MarySmi1       "
' ** 16 characters per column

Private Const THIS_NAME  As String = "frmUser_SecurityNotice"
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           .UserPass = .OpenArgs

150           .PictureData = .Form_User_SecurityNotice_img

160           If IsLoaded("frmLicense", acForm) = True Then  ' ** Module Function: modFileUtilities.
170             Forms![frmLicense].Visible = False
180           End If

190           .cmdOK.SetFocus

200         Else
210           Cancel = -1
220         End If
230       Else
240         Cancel = -1
250       End If
260     End With

270     If Cancel = -1 Then
280       Beep
290       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
300     End If

EXITP:
310     Exit Sub

ERRH:
320     Select Case ERR.Number
        Case Else
330       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
340     End Select
350     Resume EXITP

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
480         cmdOK_Click  ' ** Procedure: Below.
490       End Select
500     End If

EXITP:
510     KeyCode = intRetVal
520     Exit Sub

ERRH:
530     intRetVal = 0
540     Select Case ERR.Number
        Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
560     End Select
570     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String

610     strDocName = "frmLicense"
620     If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
630       Forms(strDocName).Visible = True
640     End If

EXITP:
650     Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub cmdOK_Click()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

710     DoCmd.Close acForm, THIS_NAME

EXITP:
720     Exit Sub

ERRH:
730     Select Case ERR.Number
        Case Else
740       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
750     End Select
760     Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

810     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
820     intShiftDown = (Shift And acShiftMask) > 0
830     intAltDown = (Shift And acAltMask) > 0
840     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyTab
870         With Me
880           intRetVal = 0
890           .UserPass.SetFocus
900         End With
910       End Select
920     End If

        ' ** Shift keys.
930     If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
940       Select Case intRetVal
          Case vbKeyTab
950         With Me
960           intRetVal = 0
970           .UserPass.SetFocus
980         End With
990       End Select
1000    End If

EXITP:
1010    KeyCode = intRetVal
1020    Exit Sub

ERRH:
1030    intRetVal = 0
1040    Select Case ERR.Number
        Case Else
1050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1060    End Select
1070    Resume EXITP

End Sub

Private Sub UserPass_GotFocus()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "UserPass_GotFocus"

1110    With Me
1120      .UserPass.SelLength = 0
1130    End With

EXITP:
1140    Exit Sub

ERRH:
1150    Select Case ERR.Number
        Case Else
1160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1170    End Select
1180    Resume EXITP

End Sub

Private Sub UserPass_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "UserPass_KeyDown"

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
          Case vbKeyTab
1270        With Me
1280          intRetVal = 0
1290          .cmdOK.SetFocus
1300        End With
1310      End Select
1320    End If

        ' ** Shift keys.
1330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1340      Select Case intRetVal
          Case vbKeyTab
1350        With Me
1360          intRetVal = 0
1370          .cmdOK.SetFocus
1380        End With
1390      End Select
1400    End If

EXITP:
1410    KeyCode = intRetVal
1420    Exit Sub

ERRH:
1430    intRetVal = 0
1440    Select Case ERR.Number
        Case Else
1450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1460    End Select
1470    Resume EXITP

End Sub
