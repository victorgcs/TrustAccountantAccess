VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEULA_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmEULA_Sub"

'VGC 08/28/2012: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Accept:           A {cmdAcceptDecline_Accept on frmEULA}
' **   Decline:          D {cmdAcceptDecline_Decline on frmEULA}
' **   EULA:             E {frmEULA_Sub on frmEULA}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Form:       X {cmdCloseForm on frmEula}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmEULA" Then

160       Else
170         Cancel = -1
180       End If
190     End With

200     If Cancel = -1 Then
210       Beep
220       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
230     End If

EXITP:
240     Exit Sub

ERRH:
250     Select Case ERR.Number
        Case Else
260       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
270     End Select
280     Resume EXITP

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

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Form:       X {Close on frmEULA}

        ' ** Plain keys.
350     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
360       Select Case intRetVal
          Case vbKeyEscape
370         With Me
380           intRetVal = 0
390           .Parent.cmdAcceptDecline_Decline_Click  ' ** Form Procedure: frmEULA.
400         End With
410       End Select
420     End If

        ' ** Ctrl-Shift keys.
430     If intCtrlDown And (Not intAltDown) And intShiftDown Then
440       Select Case intRetVal
          Case vbKeyX
450         If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
460           intRetVal = 0
470           DoCmd.Close acForm, strCallingForm
480           SetOption_Dev  ' ** Module Function: modUtilities.
490         End If
500       End Select
510     End If

EXITP:
520     KeyCode = intRetVal
530     Exit Sub

ERRH:
540     intRetVal = 0
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
670         With Me
680           intRetVal = 0
690           DoCmd.SelectObject acForm, .Parent.Name, False
700           .Parent.cmdAcceptDecline_Decline.SetFocus
710         End With
720       End Select
730     End If

        ' ** Shift keys.
740     If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
750       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
760         With Me
770           intRetVal = 0
780           DoCmd.SelectObject acForm, .Parent.Name, False
790           .Parent.cmdAcceptDecline_Accept.SetFocus
800         End With
810       End Select
820     End If

EXITP:
830     KeyCode = intRetVal
840     Exit Sub

ERRH:
850     intRetVal = 0
860     Select Case ERR.Number
        Case Else
870       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
880     End Select
890     Resume EXITP

End Sub
