VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks_Confirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks_Confirm"

'VGC 07/04/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   No:           N {opgConfirm_optNo}
' **   OK:           O {cmdOK}
' **   Some:         S {opgConfirm_optPartial}
' **   Yes:          Y {opgConfirm_optYes}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           .PictureData = .Form_Rpt_Checks_Confirm_img

160           opgConfirm_AfterUpdate  ' ** Procedure: Below.

170           .cmdOK.SetFocus

180           DoCmd.Hourglass False

190         Else
200           Cancel = -1
210         End If
220       Else
230         Cancel = -1
240       End If
250     End With

260     DoCmd.Hourglass False

270     If Cancel = -1 Then
280       Beep
290       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
300     End If

EXITP:
310     Exit Sub

ERRH:
320     DoCmd.Hourglass False
330     Select Case ERR.Number
        Case Else
340       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
350     End Select
360     Resume EXITP

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
            ' ** Do not allow Escape (what would it mean?).
470         intRetVal = 0
480       End Select
490     End If

        ' ** Ctrl-Shift keys.
500     If intCtrlDown And (Not intAltDown) And intShiftDown Then
510       Select Case intRetVal
          Case vbKeyF
520         With Me
530           intRetVal = 0
540           .FocusHolder.SetFocus
550         End With
560       End Select
570     End If

EXITP:
580     KeyCode = intRetVal
590     Exit Sub

ERRH:
600     intRetVal = 0
610     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
620     Case Else
630       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
640     End Select
650     Resume EXITP

End Sub

Private Sub Form_Timer()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

710     With Me
720       .TimerInterval = 0&

730     End With

EXITP:
740     Exit Sub

ERRH:
750     Select Case ERR.Number
        Case Else
760       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
770     End Select
780     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

810     gstrReturningForm = THIS_NAME
820     gblnSetFocus = True
830     Select Case strCallingForm
        Case "frmRpt_Checks"
840       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
850         DoCmd.SelectObject acForm, strCallingForm, False
860       Else
870         DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
880       End If
890     End Select

EXITP:
900     Exit Sub

ERRH:
910     DoCmd.Hourglass False
920     Select Case ERR.Number
        Case Else
930       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
940     End Select
950     Resume EXITP

End Sub

Private Sub cmdOK_Click()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

1010    With Me

1020      DoCmd.Hourglass True
1030      DoEvents

1040      Forms(strCallingForm).opgConfirm = .opgConfirm
1050      DoEvents

1060      DoCmd.Close acForm, THIS_NAME

1070    End With

EXITP:
1080    Exit Sub

ERRH:
1090    DoCmd.Hourglass False
1100    Select Case ERR.Number
        Case Else
1110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1120    End Select
1130    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
1290          .opgConfirm.SetFocus
1300        End With
1310      End Select
1320    End If

        ' ** Shift keys.
1330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1340      Select Case intRetVal
          Case vbKeyTab
1350        With Me
1360          intRetVal = 0
1370          .opgConfirm.SetFocus
1380        End With
1390      End Select
1400    End If

EXITP:
1410    KeyCode = intRetVal
1420    Exit Sub

ERRH:
1430    intRetVal = 0
1440    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1450    Case Else
1460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1470    End Select
1480    Resume EXITP

End Sub

Private Sub opgConfirm_AfterUpdate()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgConfirm_AfterUpdate"

1510    With Me
1520      Select Case .opgConfirm
          Case .opgConfirm_optYes.OptionValue
1530        .opgConfirm_optYes_lbl.FontBold = True
1540        .opgConfirm_optNo_lbl.FontBold = False
1550        .opgConfirm_optPartial_lbl.FontBold = False
1560      Case .opgConfirm_optNo.OptionValue
1570        .opgConfirm_optYes_lbl.FontBold = False
1580        .opgConfirm_optNo_lbl.FontBold = True
1590        .opgConfirm_optPartial_lbl.FontBold = False
1600      Case .opgConfirm_optPartial.OptionValue
1610        .opgConfirm_optYes_lbl.FontBold = False
1620        .opgConfirm_optNo_lbl.FontBold = False
1630        .opgConfirm_optPartial_lbl.FontBold = True
1640      End Select
1650    End With

EXITP:
1660    Exit Sub

ERRH:
1670    Select Case ERR.Number
        Case Else
1680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1690    End Select
1700    Resume EXITP

End Sub

Private Sub opgConfirm_optYes_KeyDown(KeyCode As Integer, Shift As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgConfirm_optYes_KeyDown"

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
1890          .cmdOK.SetFocus
1900        End With
1910      End Select
1920    End If

        ' ** Shift keys.
1930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1940      Select Case intRetVal
          Case vbKeyTab
1950        With Me
1960          intRetVal = 0
1970          .cmdOK.SetFocus
1980        End With
1990      End Select
2000    End If

EXITP:
2010    KeyCode = intRetVal
2020    Exit Sub

ERRH:
2030    intRetVal = 0
2040    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2050    Case Else
2060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2070    End Select
2080    Resume EXITP

End Sub

Private Sub opgConfirm_optNo_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgConfirm_optNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2120    intShiftDown = (Shift And acShiftMask) > 0
2130    intAltDown = (Shift And acAltMask) > 0
2140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyTab
2170        With Me
2180          intRetVal = 0
2190          .cmdOK.SetFocus
2200        End With
2210      End Select
2220    End If

        ' ** Shift keys.
2230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyTab
2250        With Me
2260          intRetVal = 0
2270          .cmdOK.SetFocus
2280        End With
2290      End Select
2300    End If

EXITP:
2310    KeyCode = intRetVal
2320    Exit Sub

ERRH:
2330    intRetVal = 0
2340    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2350    Case Else
2360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2370    End Select
2380    Resume EXITP

End Sub

Private Sub opgConfirm_optPartial_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgConfirm_optPartial_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2420    intShiftDown = (Shift And acShiftMask) > 0
2430    intAltDown = (Shift And acAltMask) > 0
2440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2460      Select Case intRetVal
          Case vbKeyTab
2470        With Me
2480          intRetVal = 0
2490          .cmdOK.SetFocus
2500        End With
2510      End Select
2520    End If

        ' ** Shift keys.
2530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2540      Select Case intRetVal
          Case vbKeyTab
2550        With Me
2560          intRetVal = 0
2570          .cmdOK.SetFocus
2580        End With
2590      End Select
2600    End If

EXITP:
2610    KeyCode = intRetVal
2620    Exit Sub

ERRH:
2630    intRetVal = 0
2640    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2650    Case Else
2660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2670    End Select
2680    Resume EXITP

End Sub
