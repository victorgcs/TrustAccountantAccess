VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmVersion_Input_Notice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmVersion_Input_Notice"

'VGC 11/15/2015: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   OK:               O {cmdOK}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           .PictureData = .Form_Version_Input_Notice_img

160         Else
170           Cancel = -1
180         End If
190       Else
200         Cancel = -1
210       End If
220     End With

230     If Cancel = -1 Then
240       Beep
250       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
260     End If

EXITP:
270     Exit Sub

ERRH:
280     Select Case ERR.Number
        Case Else
290       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
300     End Select
310     Resume EXITP

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
470         cmdOK_Click  ' ** Procedure: Below.
480       End Select
490     End If

EXITP:
500     KeyCode = intRetVal
510     Exit Sub

ERRH:
520     intRetVal = 0
530     Select Case ERR.Number
        Case Else
540       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
550     End Select
560     Resume EXITP

End Sub

Private Sub cmdOK_Click()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

610     DoCmd.Close acForm, THIS_NAME

EXITP:
620     Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub
