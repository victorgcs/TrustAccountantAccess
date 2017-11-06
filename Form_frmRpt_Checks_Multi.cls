VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks_Multi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks_Multi"

'VGC 03/22/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   OK:               O {cmdOK}
' **   Users:            U {lbxUser}

' ** lbxUser list box column constants:
Private Const LBX_U_USER As Integer = 0  ' ** journal_USER
'Private Const LBX_U_CNT  As Integer = 1  ' ** cnt_chks

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           gstrFormQuerySpec = THIS_NAME

160           .PictureData = .Form_Print_Checks_Multi_img

170           .lbxUser.Requery

180           Beep
190           DoCmd.Hourglass False

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
330       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
340     End Select
350     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

400   On Error GoTo ERRH

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        Const THIS_PROC As String = "Form_KeyDown"

        ' ** Use bit masks to determine which key was pressed.
410     intShiftDown = (Shift And acShiftMask) > 0
420     intAltDown = (Shift And acAltMask) > 0
430     intCtrlDown = (Shift And acCtrlMask) > 0

440     intRetVal = KeyCode

        ' ** Plain keys.
450     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
460       Select Case intRetVal
          Case vbKeyEscape
470         intRetVal = 0
480         cmdCancel_Click  ' ** Procedure: Below.
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

610     If strCallingForm <> vbNullString Then
620       DoCmd.SelectObject acForm, strCallingForm, False
630     End If

EXITP:
640     Exit Sub

ERRH:
650     Select Case ERR.Number
        Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
670     End Select
680     Resume EXITP

End Sub

Private Sub cmdOK_Click()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

710     With Me
720       If IsNull(.lbxUser) = False Then
730         If Trim(.lbxUser) <> vbNullString Then
740           DoCmd.Hourglass True
750           DoEvents
760           gstrJournalUser = .lbxUser.Column(LBX_U_USER)
770           DoCmd.Close acForm, THIS_NAME
780         Else
790           MsgBox "Please choose a User, or click Cancel to exit.", vbInformation + vbOKOnly, "Entry Required"
800           .lbxUser.SetFocus
810         End If
820       Else
830         MsgBox "Please choose a User, or click Cancel to exit.", vbInformation + vbOKOnly, "Entry Required"
840         .lbxUser.SetFocus
850       End If
860     End With

EXITP:
870     Exit Sub

ERRH:
880     Select Case ERR.Number
        Case Else
890       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
900     End Select
910     Resume EXITP

End Sub

Private Sub cmdCancel_Click()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

1010    gstrJournalUser = vbNullString
1020    DoCmd.Close acForm, THIS_NAME

EXITP:
1030    Exit Sub

ERRH:
1040    Select Case ERR.Number
        Case Else
1050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1060    End Select
1070    Resume EXITP

End Sub

Private Sub lbxUser_DblClick(Cancel As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxUser_DblClick"

1110    With Me
1120      If IsNull(.lbxUser) = False Then
1130        If Trim(lbxUser) <> vbNullString Then
1140          cmdOK_Click  ' ** Procedure: Above.
1150        End If
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
