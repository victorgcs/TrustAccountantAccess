VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssets_Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssets_Update"

'VGC 01/26/2013: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   OK Remain:        K {cmdOKRemain}
' **   OK Exit:          O {cmdOKExit}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)
150           Beep

160           .PictureData = .Form_Assets_Update_img

170         Else
180           Cancel = -1
190         End If
200       Else
210         Cancel = -1
220       End If
230     End With

240     If Cancel = -1 Then
250       Beep
260       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
270     End If

EXITP:
280     Exit Sub

ERRH:
290     Select Case ERR.Number
        Case Else
300       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
310     End Select
320     Resume EXITP

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
470         With Me
480           intRetVal = 0
490           cmdOKRemain_Click  ' ** Procedure: Below.
500         End With
510       End Select
520     End If

EXITP:
530     KeyCode = intRetVal
540     Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdOKExit_Click()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOKExit_Click"

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

Private Sub cmdOKRemain_Click()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOKRemain_Click"

710     gblnMessage = False
720     DoCmd.Close acForm, THIS_NAME

EXITP:
730     Exit Sub

ERRH:
740     Select Case ERR.Number
        Case Else
750       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
760     End Select
770     Resume EXITP

End Sub
