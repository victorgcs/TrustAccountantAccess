VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTransaction_Audit_Sub_Graphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmTransaction_Audit_Sub_Graphics"

'VGC 09/08/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Exit:             X {cmdClose}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmTransaction_Audit" Then

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

        ' ** Alt keys.
350     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
360       Select Case intRetVal
          Case vbKeyX
370         With Me
380           intRetVal = 0
390           .Parent.cmdClose_Click  ' ** Form Procedure: frmTransaction_Audit.
400         End With
410       End Select
420     End If

        ' ** Ctrl-Shift keys.
430     If intCtrlDown And (Not intAltDown) And intShiftDown Then
440       Select Case intRetVal
          Case vbKeyF
450         With Me
460           intRetVal = 0
470           DoCmd.SelectObject acForm, .Parent.Name, False
480           .Parent.FocusHolder.SetFocus
490         End With
500       End Select
510     End If

EXITP:
520     KeyCode = intRetVal
530     Exit Sub

ERRH:
540     intRetVal = 0
550     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
560     Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub
