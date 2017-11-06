VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReferences_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmReferences_Sub"

'VGC 12/07/2014: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmReferences}
' **   Refs:             R {frmReferences_Sub on frmReferences}
' **   Exit:             X {cmdClose on frmReferences}

Private lngLastFocus As Long, strLastFocus As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmReferences" Then

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
        ' **   Close Form:       X {cmdCloseForm on frmReferences}

        ' ** Alt keys.
350     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
360       Select Case intRetVal
          Case vbKeyX
370         With Me
380           intRetVal = 0
390           .Parent.cmdClose_Click  ' ** Form Procedure: frmReferences.
400         End With
410       End Select
420     End If

        ' ** Ctrl-Shift keys.
430     If intCtrlDown And (Not intAltDown) And intShiftDown Then
440       Select Case intRetVal
          Case vbKeyX
450         With Me
460           intRetVal = 0
470           .Parent.cmdClose_Click  ' ** Form Procedure: frmReferences.
480         End With
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

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_MouseWheel"

        Dim strTmp01 As String

610     With Me
620       If Count > 0 Then
630         If lngLastFocus = 0& Then
640           lngLastFocus = 1&
650           strLastFocus = "FocusHolder01"
660         End If
670         If lngLastFocus < 40& Then
680           If lngLastFocus <= 10& Then
690             lngLastFocus = 11&
700           Else
710             lngLastFocus = lngLastFocus + 1&
720           End If
730           strTmp01 = strLastFocus
740           strTmp01 = Left(strTmp01, (Len(strTmp01) - 2)) & Right("00" & CStr(lngLastFocus), 2)
750           .Controls(strTmp01).SetFocus
760         End If
770       Else
780         If lngLastFocus = 0& Then
790           lngLastFocus = 1&
800           strLastFocus = "FocusHolder01"
810         End If
820         If lngLastFocus > 1& Then
830           If lngLastFocus >= 30& Then
840             lngLastFocus = 29&
850           Else
860             lngLastFocus = lngLastFocus - 1&
870           End If
880           strTmp01 = strLastFocus
890           strTmp01 = Left(strTmp01, (Len(strTmp01) - 2)) & Right("00" & CStr(lngLastFocus), 2)
900           .Controls(strTmp01).SetFocus
910         End If
920       End If
930     End With

EXITP:
940     Exit Sub

ERRH:
950     Select Case ERR.Number
        Case Else
960       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
970     End Select
980     Resume EXITP

End Sub

Private Sub Ref_display_GotFocus()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Ref_display_GotFocus"

1010    With Me
1020      .Ref_display.SelLength = 0
1030      .Ref_display.SelStart = 0
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

Private Sub FocusHolder01_KeyDown(KeyCode As Integer, Shift As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder01_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1120    intShiftDown = (Shift And acShiftMask) > 0
1130    intAltDown = (Shift And acAltMask) > 0
1140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1170        With Me
1180          intRetVal = 0
1190          .FocusHolder02.SetFocus
1200        End With
1210      End Select
1220    End If

        ' ** Shift keys.
1230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1250        With Me
1260          intRetVal = 0
1270          DoCmd.SelectObject acForm, .Parent.Name, False
1280          .Parent.cmdClose.SetFocus
1290        End With
1300      End Select
1310    End If

EXITP:
1320    KeyCode = intRetVal
1330    Exit Sub

ERRH:
1340    intRetVal = 0
1350    Select Case ERR.Number
        Case Else
1360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1370    End Select
1380    Resume EXITP

End Sub

Private Sub FocusHolder40_KeyDown(KeyCode As Integer, Shift As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder40_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1420    intShiftDown = (Shift And acShiftMask) > 0
1430    intAltDown = (Shift And acAltMask) > 0
1440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1470        With Me
1480          intRetVal = 0
1490          DoCmd.SelectObject acForm, .Parent.Name, False
1500          .Parent.cmdClose.SetFocus
1510        End With
1520      End Select
1530    End If

        ' ** Shift keys.
1540    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1550      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1560        With Me
1570          intRetVal = 0
1580          .FocusHolder39.SetFocus
1590        End With
1600      End Select
1610    End If

EXITP:
1620    KeyCode = intRetVal
1630    Exit Sub

ERRH:
1640    intRetVal = 0
1650    Select Case ERR.Number
        Case Else
1660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1670    End Select
1680    Resume EXITP

End Sub
