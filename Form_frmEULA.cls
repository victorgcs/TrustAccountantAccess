VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEULA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmEULA"

'VGC 12/07/2012: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Accept:           A {cmdAcceptDecline_Accept}
' **   Decline:          D {cmdAcceptDecline_Decline}
' **   EULA:             E {frmEULA_Sub}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Form:       X {cmdCloseForm}

Private strUsername As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           DoCmd.Restore

150           strCallingForm = Trim(.OpenArgs)

              ' ** Call MouseWheelOFF() first, because that's the one that loads the MouseHood.dll library.
160           blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.
170           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

180           gstrAccept = vbNullString
190           gdatAccept = 1

200           .cmdAcceptDecline_Decline.SetFocus

210         Else
              'Cancel = -1
220         End If
230       Else
            'Cancel = -1
240       End If

250     End With

260     DoCmd.Hourglass False

270     If Cancel = -1 Then
          ' ** Just for the record.
280       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Allowed"
290     End If

EXITP:
300     Exit Sub

ERRH:
310     Select Case ERR.Number
        Case Else
320       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
330     End Select
340     Resume EXITP

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

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Form:       X {Close}

        ' ** Plain keys.
450     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
460       Select Case intRetVal
          Case vbKeyEscape
470         intRetVal = 0
480         cmdAcceptDecline_Decline_Click  ' ** Procedure: Below.
490       End Select
500     End If

        ' ** Ctrl-Shift keys.
510     If intCtrlDown And (Not intAltDown) And intShiftDown Then
520       Select Case intRetVal
          Case vbKeyX
530         strUsername = GetUserName  ' ** Module Function: modFileUtilities.
540         If strUsername = gstrDevUserName Then
550           intRetVal = 0
560           DoCmd.Close acForm, THIS_NAME
570           SetOption_Dev  ' ** Module Function: modUtilities.
580         End If
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

        Dim strFormName As String

710     With Me
720       .TimerInterval = 0&
730       If gblnSetFocus = True Then
740         gblnSetFocus = False
750         strFormName = vbNullString
760   On Error Resume Next
770         Do Until strFormName = THIS_NAME
780           DoCmd.SelectObject acForm, THIS_NAME, False
790           strFormName = Screen.ActiveForm.Name
800         Loop
810   On Error GoTo ERRH
820       End If
830     End With

EXITP:
840     Exit Sub

ERRH:
850     Select Case ERR.Number
        Case Else
860       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
870     End Select
880     Resume EXITP

End Sub

Private Sub frmEULA_Sub_Enter()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "frmEULA_Sub_Enter"

910     With Me
920       .frmEULA_Sub.Form.Detail_hline01.Visible = True
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

Private Sub frmEULA_Sub_Exit(Cancel As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "frmEULA_Sub_Exit"

1010    With Me
1020      .frmEULA_Sub.Form.Detail_hline01.Visible = False
1030    End With

EXITP:
1040    Exit Sub

ERRH:
1050    Select Case ERR.Number
        Case Else
1060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1070    End Select
1080    Resume EXITP

End Sub

Private Sub cmdAcceptDecline_Accept_Click()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcceptDecline_Accept_Click"

        Dim blnRetVal As Boolean

1110    With Me
1120      gstrAccept = .cmdAcceptDecline_Accept.Caption
1130      gdatAccept = Now()
1140      blnRetVal = DemoLicense_Set  ' ** Module Function: modSecurityFunctions.
1150      If blnRetVal = True Then
1160        DoCmd.Close acForm, THIS_NAME
1170      Else
1180        Beep
1190        MsgBox "Your license agreement could not be accepted." & vbCrLf & vbCrLf & _
              "Please contact Delta Data, Inc.", vbCritical + vbOKOnly, "License Error"
1200        DoCmd.Quit
1210      End If
1220    End With

EXITP:
1230    Exit Sub

ERRH:
1240    Select Case ERR.Number
        Case Else
1250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1260    End Select
1270    Resume EXITP

End Sub

Private Sub cmdAcceptDecline_Accept_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcceptDecline_Accept_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1320    intShiftDown = (Shift And acShiftMask) > 0
1330    intAltDown = (Shift And acAltMask) > 0
1340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyTab
1370        With Me
1380          intRetVal = 0
1390          .frmEULA_Sub.SetFocus
1400        End With
1410      End Select
1420    End If

        ' ** Shift keys.
1430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1440      Select Case intRetVal
          Case vbKeyTab
1450        With Me
1460          intRetVal = 0
1470          .cmdAcceptDecline_Decline.SetFocus
1480        End With
1490      End Select
1500    End If

EXITP:
1510    KeyCode = intRetVal
1520    Exit Sub

ERRH:
1530    intRetVal = 0
1540    Select Case ERR.Number
        Case Else
1550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1560    End Select
1570    Resume EXITP

End Sub

Public Sub cmdAcceptDecline_Decline_Click()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcceptDecline_Decline_Click"

1610    With Me

1620      gstrAccept = .cmdAcceptDecline_Decline.Caption

1630      strUsername = GetUserName  ' ** Module Function: modFileUtilities.
1640      If strUsername = gstrDevUserName Then
1650        DoCmd.Close acForm, THIS_NAME
1660      Else
1670        DoCmd.Quit
1680      End If

1690    End With

EXITP:
1700    Exit Sub

ERRH:
1710    Select Case ERR.Number
        Case Else
1720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1730    End Select
1740    Resume EXITP

End Sub

Private Sub cmdAcceptDecline_Decline_KeyDown(KeyCode As Integer, Shift As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcceptDecline_Decline_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1820    intShiftDown = (Shift And acShiftMask) > 0
1830    intAltDown = (Shift And acAltMask) > 0
1840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1860      Select Case intRetVal
          Case vbKeyTab
1870        With Me
1880          intRetVal = 0
1890          .cmdAcceptDecline_Accept.SetFocus
1900        End With
1910      End Select
1920    End If

        ' ** Shift keys.
1930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1940      Select Case intRetVal
          Case vbKeyTab
1950        With Me
1960          intRetVal = 0
1970          .frmEULA_Sub.SetFocus
1980        End With
1990      End Select
2000    End If

EXITP:
2010    KeyCode = intRetVal
2020    Exit Sub

ERRH:
2030    intRetVal = 0
2040    Select Case ERR.Number
        Case Else
2050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2060    End Select
2070    Resume EXITP

End Sub
