VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSiteMap_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmSiteMap_Sub"

'VGC 03/26/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmSiteMap}
' **   Print Report:     P {cmdPrintReport on frmSiteMap}
' **   Exit:             X {cmdClose on frmSiteMap}
' **   Save Size/Pos:    Z {chkSaveSizePos on frmSiteMap}

' ** Shortcut Ctrl keys responsive from this form:
' **   Journal Map:      J {Level02_02_01_s02_arrow}

Private strFocusHolder As String
Private lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmSiteMap" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170         blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

180         SetUnavailable  ' ** Procedure: Below.

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
        Dim Cancel As Integer
        Dim intRetVal As Integer

410     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
420     intShiftDown = (Shift And acShiftMask) > 0
430     intAltDown = (Shift And acAltMask) > 0
440     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Journal Map:      J {Level02_02_01_s02_arrow}

        ' ** Plain keys.
450     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
460       Select Case intRetVal
          Case vbKeyEscape
470         With Me
480           intRetVal = 0
490           .Parent.cmdClose_Click  ' ** Form Procedure: frmSiteMap.
500         End With
510       End Select
520     End If

        ' ** Alt keys.
530     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
540       Select Case intRetVal
          Case vbKeyX
550         With Me
560           intRetVal = 0
570           .Parent.cmdClose_Click  ' ** Form Procedure: frmSiteMap.
580         End With
590       End Select
600     End If

        ' ** Ctrl keys.
610     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
620       Select Case intRetVal
          Case vbKeyJ
630         With Me
640           intRetVal = 0
650           Cancel = 0
660           Level02_02_01_s02_arrow_DblClick Cancel  ' ** Procedure: Below.
670         End With
680       Case vbKeyTab, vbKeyReturn
690         With Me
700           intRetVal = 0
710           DoCmd.SelectObject acForm, .Parent.Name, False
720           If .Parent.cmdPrintReport.Enabled = True Then
730             .Parent.cmdPrintReport.SetFocus
740           Else
750             .Parent.chkSaveSizePos.SetFocus
760           End If
770         End With
780       End Select
790     End If

        ' ** Ctrl-Shift keys.
800     If intCtrlDown And (Not intAltDown) And intShiftDown Then
810       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
820         With Me
830           intRetVal = 0
840           DoCmd.SelectObject acForm, .Parent.Name, False
850           .Parent.cmdClose.SetFocus
860         End With
870       End Select
880     End If

EXITP:
890     KeyCode = intRetVal
900     Exit Sub

ERRH:
910     intRetVal = 0
920     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
930     Case Else
940       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
950     End Select
960     Resume EXITP

End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_MouseWheel"

1010    With Me
1020      If Count > 0 Then
1030        Select Case strFocusHolder
            Case vbNullString
1040          .FocusHolder01.SetFocus  ' ** Start with the first.
1050        Case "FocusHolder01"
1060          .FocusHolder07.SetFocus  ' ** To 1st bottom one.
1070        Case "FocusHolder02"
1080          .FocusHolder07.SetFocus  ' ** To 1st bottom one.
1090        Case "FocusHolder03"
1100          .FocusHolder07.SetFocus  ' ** To 1st bottom one.
1110        Case "FocusHolder04"
1120          .FocusHolder07.SetFocus  ' ** To 1st bottom one.
1130        Case "FocusHolder05"
1140          .FocusHolder07.SetFocus  ' ** To 1st bottom one.
1150        Case "FocusHolder06"
1160          .FocusHolder07.SetFocus  ' ** To 1st bottom one.
1170        Case "FocusHolder07"
1180          .FocusHolder08.SetFocus  ' ** To 2nd bottom one.
1190        Case "FocusHolder08"
1200          .FocusHolder09.SetFocus  ' ** To 3rd bottom one.
1210        Case "FocusHolder09"
1220          .FocusHolder10.SetFocus  ' ** To 4th bottom one.
1230        Case "FocusHolder10"
1240          .FocusHolder11.SetFocus  ' ** To 5th bottom one.
1250        Case "FocusHolder11"
1260          .FocusHolder12.SetFocus  ' ** To 6th bottom one.
1270        Case "FocusHolder12"
1280          .FocusHolder12.SetFocus  ' ** Stay put.
1290        End Select
1300      Else  ' ** < 0.
1310        Select Case strFocusHolder
            Case vbNullString
1320          .FocusHolder01.SetFocus  ' ** Start with the first.
1330        Case "FocusHolder01"
1340          .FocusHolder01.SetFocus  ' ** Stay put.
1350        Case "FocusHolder02"
1360          .FocusHolder01.SetFocus  ' ** To 1st top one.
1370        Case "FocusHolder03"
1380          .FocusHolder02.SetFocus  ' ** To 2nd top one.
1390        Case "FocusHolder04"
1400          .FocusHolder03.SetFocus  ' ** To 3rd top one.
1410        Case "FocusHolder05"
1420          .FocusHolder04.SetFocus  ' ** To 4th top one.
1430        Case "FocusHolder06"
1440          .FocusHolder05.SetFocus  ' ** To 5th top one.
1450        Case "FocusHolder07"
1460          .FocusHolder06.SetFocus  ' ** To 6th top one.
1470        Case "FocusHolder08"
1480          .FocusHolder06.SetFocus  ' ** To 6th top one.
1490        Case "FocusHolder09"
1500          .FocusHolder06.SetFocus  ' ** To 6th top one.
1510        Case "FocusHolder10"
1520          .FocusHolder06.SetFocus  ' ** To 6th top one.
1530        Case "FocusHolder11"
1540          .FocusHolder06.SetFocus  ' ** To 6th top one.
1550        Case "FocusHolder12"
1560          .FocusHolder06.SetFocus  ' ** To 6th top one.
1570        End Select
1580      End If
          'Debug.Print "'COUNT: " & CStr(Count)
1590    End With

EXITP:
1600    Exit Sub

ERRH:
1610    Select Case ERR.Number
        Case Else
1620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1630    End Select
1640    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1710    With Me
1720      If .Level02_02_01_s02_arrow_a.Visible = True Then
1730        .Level02_02_01_s02_arrow.Visible = True
1740        .Level02_02_01_s02_arrow_shadow.Visible = True
1750        .Level02_02_01_s02_arrow_a.Visible = False
1760        .Level02_02_01_s02_arrow_a_shadow.Visible = False
1770        .Level02_02_01_s02_arrow_b.Visible = False
1780        .Level02_02_01_s02_arrow_c.Visible = False
1790        .Level02_02_01_s02_arrow_d.Visible = False
1800        .Level02_02_01_s02_line1.Visible = True
1810        .Level02_02_01_s02_line1_shadow.Visible = True
1820        .Level02_02_01_s02_line2.Visible = True
1830        .Level02_02_01_s02_line2_shadow.Visible = True
1840        .Level02_02_01_s02_line3.Visible = False
1850        .Level02_02_01_s02_line3_shadow.Visible = False
1860        .Level02_02_01_s02_line4.Visible = False
1870        .Level02_02_01_s02_line4_shadow.Visible = False
1880      End If
1890      If .Level02_02_01_s03_arrow_a.Visible = True Then
1900        .Level02_02_01_s03_arrow.Visible = True
1910        .Level02_02_01_s03_arrow_shadow.Visible = True
1920        .Level02_02_01_s03_arrow_a.Visible = False
1930        .Level02_02_01_s03_arrow_a_shadow.Visible = False
1940        .Level02_02_01_s03_arrow_b.Visible = False
1950        .Level02_02_01_s03_arrow_c.Visible = False
1960        .Level02_02_01_s03_arrow_d.Visible = False
1970        .Level02_02_01_s03_line1.Visible = True
1980        .Level02_02_01_s03_line1_shadow.Visible = True
1990        .Level02_02_01_s03_line2.Visible = True
2000        .Level02_02_01_s03_line2_shadow.Visible = True
2010        .Level02_02_01_s03_line3.Visible = False
2020        .Level02_02_01_s03_line3_shadow.Visible = False
2030        .Level02_02_01_s03_line4.Visible = False
2040        .Level02_02_01_s03_line4_shadow.Visible = False
2050      End If
2060    End With

EXITP:
2070    Exit Sub

ERRH:
2080    Select Case ERR.Number
        Case Else
2090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2100    End Select
2110    Resume EXITP

End Sub

Private Sub FocusHolder01_GotFocus()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder01_GotFocus"

2210    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Private Sub FocusHolder01_KeyDown(KeyCode As Integer, Shift As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder01_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2320    intShiftDown = (Shift And acShiftMask) > 0
2330    intAltDown = (Shift And acAltMask) > 0
2340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2370        With Me
2380          intRetVal = 0
2390          .FocusHolder02.SetFocus
2400        End With
2410      End Select
2420    End If

        ' ** Shift keys.
2430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2450        With Me
2460          intRetVal = 0
2470          DoCmd.SelectObject acForm, .Parent.Name, False
2480          .Parent.cmdClose.SetFocus
2490        End With
2500      End Select
2510    End If

EXITP:
2520    KeyCode = intRetVal
2530    Exit Sub

ERRH:
2540    intRetVal = 0
2550    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2560    Case Else
2570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2580    End Select
2590    Resume EXITP

End Sub

Private Sub FocusHolder02_GotFocus()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder02_GotFocus"

2610    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub FocusHolder03_GotFocus()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder03_GotFocus"

2710    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
2720    Exit Sub

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub FocusHolder04_GotFocus()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder04_GotFocus"

2810    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub FocusHolder05_GotFocus()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder05_GotFocus"

2910    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
2920    Exit Sub

ERRH:
2930    Select Case ERR.Number
        Case Else
2940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2950    End Select
2960    Resume EXITP

End Sub

Private Sub FocusHolder06_GotFocus()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder06_GotFocus"

3010    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub FocusHolder07_GotFocus()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder07_GotFocus"

3110    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub FocusHolder08_GotFocus()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder08_GotFocus"

3210    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub FocusHolder09_GotFocus()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder09_GotFocus"

3310    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub FocusHolder10_GotFocus()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder10_GotFocus"

3410    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub FocusHolder11_GotFocus()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder11_GotFocus"

3510    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub FocusHolder12_GotFocus()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder12_GotFocus"

3610    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub FocusHolder12_KeyDown(KeyCode As Integer, Shift As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder12_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3720    intShiftDown = (Shift And acShiftMask) > 0
3730    intAltDown = (Shift And acAltMask) > 0
3740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3770        With Me
3780          intRetVal = 0
3790          DoCmd.SelectObject acForm, .Parent.Name, False
3800          If .Parent.cmdPrintReport.Enabled = True Then
3810            .Parent.cmdPrintReport.SetFocus
3820          Else
3830            .Parent.chkSaveSizePos.SetFocus
3840          End If
3850        End With
3860      End Select
3870    End If

        ' ** Shift keys.
3880    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3890      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3900        With Me
3910          intRetVal = 0
3920          .FocusHolder11.SetFocus
3930        End With
3940      End Select
3950    End If

EXITP:
3960    KeyCode = intRetVal
3970    Exit Sub

ERRH:
3980    intRetVal = 0
3990    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4000    Case Else
4010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4020    End Select
4030    Resume EXITP

End Sub

Private Sub Level02_02_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_box_MouseMove"

4110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub Level02_02_01_s02_lbl_DblClick(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s02_lbl_DblClick"

4210    Level02_02_01_s02_arrow_DblClick Cancel  ' ** Procedure: Below.

EXITP:
4220    Exit Sub

ERRH:
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Private Sub Level02_02_01_s02_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s02_lbl_MouseMove"

        'Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.
4310    Level02_02_01_s02_arrow_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Private Sub Level02_02_01_s02_lbl2_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s02_lbl2_DblClick"

4410    Level02_02_01_s02_arrow_DblClick Cancel  ' ** Procedure: Below.

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub Level02_02_01_s02_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s02_lbl2_MouseMove"

        'Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.
4510    Level02_02_01_s02_arrow_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Public Sub Level02_02_01_s02_arrow_DblClick(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s02_arrow_DblClick"

        Dim strDocName As String

4610    strDocName = "frmSiteMap_Journal"
4620    Select Case IsLoaded(strDocName, acForm)  ' ** Module Function: modFileUtilities.
        Case True
4630      DoCmd.Close acForm, strDocName
4640    Case False
4650      gblnSetFocus = True
4660      DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & "Classic"
4670    End Select

EXITP:
4680    Exit Sub

ERRH:
4690    Select Case ERR.Number
        Case Else
4700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4710    End Select
4720    Resume EXITP

End Sub

Private Sub Level02_02_01_s02_arrow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s02_arrow_MouseMove"

4810    With Me
4820      .Level02_02_01_s02_arrow.Visible = False
4830      .Level02_02_01_s02_arrow_shadow.Visible = False
4840      .Level02_02_01_s02_arrow_a.Visible = True
4850      .Level02_02_01_s02_arrow_a_shadow.Visible = True
4860      .Level02_02_01_s02_arrow_b.Visible = True
4870      .Level02_02_01_s02_arrow_c.Visible = True
4880      .Level02_02_01_s02_arrow_d.Visible = True
4890      .Level02_02_01_s02_line1.Visible = False
4900      .Level02_02_01_s02_line1_shadow.Visible = False
4910      .Level02_02_01_s02_line2.Visible = False
4920      .Level02_02_01_s02_line2_shadow.Visible = False
4930      .Level02_02_01_s02_line3.Visible = True
4940      .Level02_02_01_s02_line3_shadow.Visible = True
4950      .Level02_02_01_s02_line4.Visible = True
4960      .Level02_02_01_s02_line4_shadow.Visible = True
4970      If .Level02_02_01_s03_arrow_a.Visible = True Then
4980        .Level02_02_01_s03_arrow.Visible = True
4990        .Level02_02_01_s03_arrow_shadow.Visible = True
5000        .Level02_02_01_s03_line1.Visible = True
5010        .Level02_02_01_s03_line1_shadow.Visible = True
5020        .Level02_02_01_s03_line2.Visible = True
5030        .Level02_02_01_s03_line2_shadow.Visible = True
5040        .Level02_02_01_s03_arrow_a.Visible = False
5050        .Level02_02_01_s03_arrow_a_shadow.Visible = False
5060        .Level02_02_01_s03_arrow_b.Visible = False
5070        .Level02_02_01_s03_arrow_c.Visible = False
5080        .Level02_02_01_s03_arrow_d.Visible = False
5090        .Level02_02_01_s03_line3.Visible = False
5100        .Level02_02_01_s03_line3_shadow.Visible = False
5110        .Level02_02_01_s03_line4.Visible = False
5120        .Level02_02_01_s03_line4_shadow.Visible = False
5130      End If
5140    End With

EXITP:
5150    Exit Sub

ERRH:
5160    Select Case ERR.Number
        Case Else
5170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5180    End Select
5190    Resume EXITP

End Sub

Private Sub Level02_02_01_s02_arrow_d_DblClick(Cancel As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s02_arrow_d_DblClick"

5210    Level02_02_01_s02_arrow_DblClick Cancel  ' ** Procedure: Above.

EXITP:
5220    Exit Sub

ERRH:
5230    Select Case ERR.Number
        Case Else
5240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5250    End Select
5260    Resume EXITP

End Sub

Private Sub Level02_02_01_s03_lbl_DblClick(Cancel As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s03_lbl_DblClick"

5310    Level02_02_01_s03_arrow_DblClick Cancel  ' ** Procedure: Below.

EXITP:
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Private Sub Level02_02_01_s03_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s03_lbl_MouseMove"

        'Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.
5410    Level02_02_01_s03_arrow_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
5420    Exit Sub

ERRH:
5430    Select Case ERR.Number
        Case Else
5440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5450    End Select
5460    Resume EXITP

End Sub

Private Sub Level02_02_01_s03_lbl2_DblClick(Cancel As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s03_lbl2_DblClick"

5510    Level02_02_01_s03_arrow_DblClick Cancel  ' ** Procedure: Below.

EXITP:
5520    Exit Sub

ERRH:
5530    Select Case ERR.Number
        Case Else
5540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5550    End Select
5560    Resume EXITP

End Sub

Private Sub Level02_02_01_s03_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s03_lbl2_MouseMove"

        'Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.
5610    Level02_02_01_s03_arrow_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
5620    Exit Sub

ERRH:
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Private Sub Level02_02_01_s03_arrow_DblClick(Cancel As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s03_arrow_DblClick"

        Dim strDocName As String

5710    strDocName = "frmSiteMap_Journal"
5720    Select Case IsLoaded(strDocName, acForm)  ' ** Module Function: modFileUtilities.
        Case True
5730      DoCmd.Close acForm, strDocName
5740    Case False
5750      gblnSetFocus = True
5760      DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & "Columnar"
5770    End Select

EXITP:
5780    Exit Sub

ERRH:
5790    Select Case ERR.Number
        Case Else
5800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5810    End Select
5820    Resume EXITP

End Sub

Private Sub Level02_02_01_s03_arrow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s03_arrow_MouseMove"

5910    With Me
5920      .Level02_02_01_s03_arrow.Visible = False
5930      .Level02_02_01_s03_arrow_shadow.Visible = False
5940      .Level02_02_01_s03_arrow_a.Visible = True
5950      .Level02_02_01_s03_arrow_a_shadow.Visible = True
5960      .Level02_02_01_s03_arrow_b.Visible = True
5970      .Level02_02_01_s03_arrow_c.Visible = True
5980      .Level02_02_01_s03_arrow_d.Visible = True
5990      .Level02_02_01_s03_line1.Visible = False
6000      .Level02_02_01_s03_line1_shadow.Visible = False
6010      .Level02_02_01_s03_line2.Visible = False
6020      .Level02_02_01_s03_line2_shadow.Visible = False
6030      .Level02_02_01_s03_line3.Visible = True
6040      .Level02_02_01_s03_line3_shadow.Visible = True
6050      .Level02_02_01_s03_line4.Visible = True
6060      .Level02_02_01_s03_line4_shadow.Visible = True
6070      If .Level02_02_01_s02_arrow_a.Visible = True Then
6080        .Level02_02_01_s02_arrow.Visible = True
6090        .Level02_02_01_s02_arrow_shadow.Visible = True
6100        .Level02_02_01_s02_line1.Visible = True
6110        .Level02_02_01_s02_line1_shadow.Visible = True
6120        .Level02_02_01_s02_line2.Visible = True
6130        .Level02_02_01_s02_line2_shadow.Visible = True
6140        .Level02_02_01_s02_arrow_a.Visible = False
6150        .Level02_02_01_s02_arrow_a_shadow.Visible = False
6160        .Level02_02_01_s02_arrow_b.Visible = False
6170        .Level02_02_01_s02_arrow_c.Visible = False
6180        .Level02_02_01_s02_arrow_d.Visible = False
6190        .Level02_02_01_s02_line3.Visible = False
6200        .Level02_02_01_s02_line3_shadow.Visible = False
6210        .Level02_02_01_s02_line4.Visible = False
6220        .Level02_02_01_s02_line4_shadow.Visible = False
6230      End If
6240    End With

EXITP:
6250    Exit Sub

ERRH:
6260    Select Case ERR.Number
        Case Else
6270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6280    End Select
6290    Resume EXITP

End Sub

Private Sub Level02_02_01_s03_arrow_d_DblClick(Cancel As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_02_01_s03_arrow_d_DblClick"

6310    Level02_02_01_s03_arrow_DblClick Cancel  ' ** Procedure: Above.

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Sub

Private Sub Level02_03_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "Level02_03_box_MouseMove"

6410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6420    Exit Sub

ERRH:
6430    Select Case ERR.Number
        Case Else
6440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6450    End Select
6460    Resume EXITP

End Sub

Private Sub SetUnavailable()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "SetUnavailable"

        Dim blnRetVal As Boolean

        Const CLR_FOREX_DRK As Long = 12713856
        Const CLR_FOREX_MED As Long = 14745536
        Const CLR_FOREX_LGT As Long = 16121835

6510    With Me
6520      If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
6530        blnRetVal = True
6540      Else
6550        blnRetVal = False
6560      End If
6570      Select Case blnRetVal
          Case True
            ' ** Administration.
6580        .Level03_02_09_lbl.Visible = True
6590        .Level03_02_09_lbl_box1.Visible = True
6600        .Level03_02_09_lbl_box2.Visible = True
6610        .Level03_02_09_arrow.Visible = True
6620        .Level03_02_09_arrow_shadow.Visible = True
6630        .Level03_02_09_line1.Visible = True
6640        .Level03_02_09_line1_shadow.Visible = True
6650        .Level03_02_09_line1a.Visible = True
6660        .Level03_02_09_line1a_shadow.Visible = True
6670        .Level03_02_09_line1b.Visible = True
6680        .Level03_02_09_line1b_shadow.Visible = True
6690        .Level03_02_09_line1c.Visible = True
6700        .Level03_02_09_line1c_shadow.Visible = True
6710        .Level03_02_09_line1d.Visible = True
6720        .Level03_02_09_line1d_shadow.Visible = True
6730        .Level03_02_09_line2.Visible = True
6740        .Level03_02_09_line2_shadow.Visible = True
6750        .Level03_02_09_line2a.Visible = True
6760        .Level03_02_09_line2a_shadow.Visible = True
6770        .Level03_02_09_line2b.Visible = True
6780        .Level03_02_09_line2b_shadow.Visible = True
6790        .Level03_02_09_line2c.Visible = True
6800        .Level03_02_09_line2c_shadow.Visible = True
6810        .Level03_02_09_line2d.Visible = True
6820        .Level03_02_09_line2d_shadow.Visible = True
6830      Case False
            ' ** Administration.
6840        .Level03_02_09_lbl.Visible = False
6850        .Level03_02_09_lbl_box1.Visible = False
6860        .Level03_02_09_lbl_box2.Visible = False
6870        .Level03_02_09_arrow.Visible = False
6880        .Level03_02_09_arrow_shadow.Visible = False
6890        .Level03_02_09_line1.Visible = False
6900        .Level03_02_09_line1_shadow.Visible = False
6910        .Level03_02_09_line1a.Visible = False
6920        .Level03_02_09_line1a_shadow.Visible = False
6930        .Level03_02_09_line1b.Visible = False
6940        .Level03_02_09_line1b_shadow.Visible = False
6950        .Level03_02_09_line1c.Visible = False
6960        .Level03_02_09_line1c_shadow.Visible = False
6970        .Level03_02_09_line1d.Visible = False
6980        .Level03_02_09_line1d_shadow.Visible = False
6990        .Level03_02_09_line2.Visible = False
7000        .Level03_02_09_line2_shadow.Visible = False
7010        .Level03_02_09_line2a.Visible = False
7020        .Level03_02_09_line2a_shadow.Visible = False
7030        .Level03_02_09_line2b.Visible = False
7040        .Level03_02_09_line2b_shadow.Visible = False
7050        .Level03_02_09_line2c.Visible = False
7060        .Level03_02_09_line2c_shadow.Visible = False
7070        .Level03_02_09_line2d.Visible = False
7080        .Level03_02_09_line2d_shadow.Visible = False
7090      End Select

7100      blnRetVal = Has_ForEx  ' ** Function: Below.
7110      Select Case blnRetVal
          Case True
            ' ** Foreign Exchange.
7120        .cmdForEx_l_hline01.BorderColor = CLR_FOREX_DRK
7130        .cmdForEx_l_hline02.BorderColor = CLR_FOREX_LGT
7140        .cmdForEx_l_hline03.BorderColor = CLR_FOREX_DRK
7150        .cmdForEx_l_hline04.BorderColor = CLR_FOREX_LGT
7160        .cmdForEx_l_vline01.BorderColor = CLR_FOREX_DRK
7170        .cmdForEx_l_vline02.BorderColor = CLR_FOREX_LGT
7180        .cmdForEx_l_vline03.BorderColor = CLR_FOREX_DRK
7190        .cmdForEx_l_vline04.BorderColor = CLR_FOREX_LGT
7200        .cmdForEx_t_hline01.BorderColor = CLR_FOREX_MED
7210        .cmdForEx_t_hline02.BorderColor = CLR_FOREX_LGT
7220        .cmdForEx_t_hline03.BorderColor = CLR_FOREX_MED
7230        .cmdForEx_t_hline04.BorderColor = CLR_FOREX_LGT
7240        .cmdForEx_t_vline01.BorderColor = CLR_FOREX_MED
7250        .cmdForEx_t_vline02.BorderColor = CLR_FOREX_LGT
7260        .cmdForEx_t_vline03.BorderColor = CLR_FOREX_MED
7270        .cmdForEx_t_vline04.BorderColor = CLR_FOREX_LGT
7280        .Level03_02_04_lbl.ForeColor = CLR_VDKGRY
7290        .Level03_02_04_lbl_dim_hi.Visible = False
7300        .Level03_02_04_lbl_box1.BackColor = CLR_LTTEAL2
7310        .Level03_02_04_arrow.ForeColor = CLR_DKGRY
7320        .Level03_02_04_arrow_shadow.ForeColor = CLR_GRY3
7330        .Level03_02_04_line1.BorderColor = CLR_DKGRY
7340        .Level03_02_04_line1_shadow.BorderColor = CLR_GRY3
7350        .Level03_02_04_line2.BorderColor = CLR_DKGRY
7360        .Level03_02_04_line2_shadow.BorderColor = CLR_GRY3
7370        .Level03_02_04_line3.BorderColor = CLR_DKGRY
7380        .Level03_02_04_line3_shadow.BorderColor = CLR_GRY3
7390        .Level03_02_04_line4.BorderColor = CLR_DKGRY
7400        .Level03_02_04_line4_shadow.BorderColor = CLR_GRY3
7410        .Level03_02_04_line5.BorderColor = CLR_DKGRY
7420        .Level03_02_04_line5_shadow.BorderColor = CLR_GRY3
7430        .Level04_01_lbl.BackStyle = acBackStyleNormal
7440        .Level04_01_lbl.ForeColor = CLR_WHT
7450        .Level04_01_lbl_dim_hi.Visible = False
7460        .Level04_01_01_lbl.ForeColor = CLR_VDKGRY
7470        .Level04_01_01_lbl_dim_hi.Visible = False
7480        .Level04_01_01_lbl_box1.BackColor = CLR_LTTEAL2
7490        .Level04_01_02_lbl.ForeColor = CLR_VDKGRY
7500        .Level04_01_02_lbl_dim_hi.Visible = False
7510        .Level04_01_02_lbl_box1.BackColor = CLR_LTTEAL2
7520        .Level04_01_03_lbl.ForeColor = CLR_VDKGRY
7530        .Level04_01_03_lbl_dim_hi.Visible = False
7540        .Level04_01_03_lbl_box1.BackColor = CLR_LTTEAL2
7550        .Level04_01_04_lbl.ForeColor = CLR_VDKGRY
7560        .Level04_01_04_lbl_dim_hi.Visible = False
7570        .Level04_01_04_lbl_box1.BackColor = CLR_LTTEAL2
7580      Case False
            ' ** Foreign Exchange.
7590        .cmdForEx_l_hline01.BorderColor = WIN_CLR_DISF
7600        .cmdForEx_l_hline02.BorderColor = CLR_LTGRY
7610        .cmdForEx_l_hline03.BorderColor = WIN_CLR_DISF
7620        .cmdForEx_l_hline04.BorderColor = CLR_LTGRY
7630        .cmdForEx_l_vline01.BorderColor = WIN_CLR_DISF
7640        .cmdForEx_l_vline02.BorderColor = CLR_LTGRY
7650        .cmdForEx_l_vline03.BorderColor = WIN_CLR_DISF
7660        .cmdForEx_l_vline04.BorderColor = CLR_LTGRY
7670        .cmdForEx_t_hline01.BorderColor = CLR_GRY3
7680        .cmdForEx_t_hline02.BorderColor = CLR_LTGRY
7690        .cmdForEx_t_hline03.BorderColor = CLR_GRY3
7700        .cmdForEx_t_hline04.BorderColor = CLR_LTGRY
7710        .cmdForEx_t_vline01.BorderColor = CLR_GRY3
7720        .cmdForEx_t_vline02.BorderColor = CLR_LTGRY
7730        .cmdForEx_t_vline03.BorderColor = CLR_GRY3
7740        .cmdForEx_t_vline04.BorderColor = CLR_LTGRY
7750        .Level03_02_04_lbl.ForeColor = WIN_CLR_DISF
7760        .Level03_02_04_lbl_dim_hi.Visible = True
7770        .Level03_02_04_lbl_box1.BackColor = WIN_CLR_DISR
7780        .Level03_02_04_arrow.ForeColor = WIN_CLR_DISF
7790        .Level03_02_04_arrow_shadow.ForeColor = CLR_LTGRY
7800        .Level03_02_04_line1.BorderColor = WIN_CLR_DISF
7810        .Level03_02_04_line1_shadow.BorderColor = CLR_LTGRY
7820        .Level03_02_04_line2.BorderColor = WIN_CLR_DISF
7830        .Level03_02_04_line2_shadow.BorderColor = CLR_LTGRY
7840        .Level03_02_04_line3.BorderColor = WIN_CLR_DISF
7850        .Level03_02_04_line3_shadow.BorderColor = CLR_LTGRY
7860        .Level03_02_04_line4.BorderColor = WIN_CLR_DISF
7870        .Level03_02_04_line4_shadow.BorderColor = CLR_LTGRY
7880        .Level03_02_04_line5.BorderColor = WIN_CLR_DISF
7890        .Level03_02_04_line5_shadow.BorderColor = CLR_LTGRY
7900        .Level04_01_lbl.BackStyle = acBackStyleTransparent
7910        .Level04_01_lbl.ForeColor = WIN_CLR_DISF
7920        .Level04_01_lbl_dim_hi.Visible = True
7930        .Level04_01_01_lbl.ForeColor = WIN_CLR_DISF
7940        .Level04_01_01_lbl_dim_hi.Visible = True
7950        .Level04_01_01_lbl_box1.BackColor = WIN_CLR_DISR
7960        .Level04_01_02_lbl.ForeColor = WIN_CLR_DISF
7970        .Level04_01_02_lbl_dim_hi.Visible = True
7980        .Level04_01_02_lbl_box1.BackColor = WIN_CLR_DISR
7990        .Level04_01_03_lbl.ForeColor = WIN_CLR_DISF
8000        .Level04_01_03_lbl_dim_hi.Visible = True
8010        .Level04_01_03_lbl_box1.BackColor = WIN_CLR_DISR
8020        .Level04_01_04_lbl.ForeColor = WIN_CLR_DISF
8030        .Level04_01_04_lbl_dim_hi.Visible = True
8040        .Level04_01_04_lbl_box1.BackColor = WIN_CLR_DISR
8050      End Select

8060      blnRetVal = Has_ChkReg  ' ** Function: Below.
8070      Select Case blnRetVal
          Case True
            ' ** Check Register.
8080        .Level03_02_05_lbl.ForeColor = CLR_VDKGRY
8090        .Level03_02_05_lbl_dim_hi.Visible = False
8100        .Level03_02_05_lbl_box1.BackColor = CLR_LTTEAL2
8110      Case False
            ' ** Check Register.
8120        .Level03_02_05_lbl.ForeColor = WIN_CLR_DISF
8130        .Level03_02_05_lbl_dim_hi.Visible = True
8140        .Level03_02_05_lbl_box1.BackColor = WIN_CLR_DISR
8150      End Select

8160      blnRetVal = Has_Form1099  ' ** Function: Below.
8170      Select Case blnRetVal
          Case True
            ' ** Form 1099.
8180        .Level02_03_18_lbl.ForeColor = CLR_VDKGRY
8190        .Level02_03_18_lbl_dim_hi.Visible = False
8200        .Level02_03_18_lbl_box1.BackColor = CLR_LTTEAL2
8210      Case False
            ' ** Form 1099.
8220        .Level02_03_18_lbl.ForeColor = WIN_CLR_DISF
8230        .Level02_03_18_lbl_dim_hi.Visible = True
8240        .Level02_03_18_lbl_box1.BackColor = WIN_CLR_DISR
8250      End Select

          'Select Case blnRetVal
          'Case True
          '.._lbl.ForeColor = CLR_VDKGRY
          '.._lbl_dim_hi.Visible = False
          '.._lbl_box1.BackColor = CLR_LTTEAL2
          '.._arrow.ForeColor = CLR_DKGRY
          '.._arrow_shadow.ForeColor = CLR_GRY3
          '.._line1.BorderColor = CLR_DKGRY
          '.._line1_shadow.BorderColor = CLR_GRY3
          'Case False
          '.._lbl.ForeColor = WIN_CLR_DISF
          '.._lbl_dim_hi.Visible = True
          '.._lbl_box1.BackColor = WIN_CLR_DISR
          '.._arrow.ForeColor = WIN_CLR_DISF
          '.._arrow_shadow.ForeColor = CLR_LTGRY
          '.._line1.BorderColor = WIN_CLR_DISF
          '.._line1_shadow.BorderColor = CLR_LTGRY
          'End Select

8260    End With

EXITP:
8270    Exit Sub

ERRH:
8280    Select Case ERR.Number
        Case Else
8290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8300    End Select
8310    Resume EXITP

End Sub

Private Function Has_ForEx() As Boolean

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "Has_ForEx"

        Dim dbs As DAO.Database, cntr As DAO.Container, doc As DAO.Document
        Dim blnRetVal As Boolean

8410    blnRetVal = False

8420    Set dbs = CurrentDb
8430    With dbs
8440      Set cntr = .Containers("Forms")
8450      With cntr
8460        For Each doc In .Documents
8470          With doc
8480            If .Name = "frmMenu_ForeignExchange" Then
8490              blnRetVal = True
8500              Exit For
8510            End If
8520          End With  ' ** doc.
8530        Next  ' ** doc.
8540      End With  ' ** cntr.
8550    End With  ' ** dbs.

EXITP:
8560    Set doc = Nothing
8570    Set cntr = Nothing
8580    Set dbs = Nothing
8590    Has_ForEx = blnRetVal
8600    Exit Function

ERRH:
8610    blnRetVal = False
8620    Select Case ERR.Number
        Case Else
8630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8640    End Select
8650    Resume EXITP

End Function

Private Function Has_ChkReg() As Boolean

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "Has_ChkReg"

        Dim dbs As DAO.Database, cntr As DAO.Container, doc As DAO.Document
        Dim blnRetVal As Boolean

8710    blnRetVal = False

8720    Set dbs = CurrentDb
8730    With dbs
8740      Set cntr = .Containers("Forms")
8750      With cntr
8760        For Each doc In .Documents
8770          With doc
8780            If .Name = "frmCheckRegister" Then
8790              blnRetVal = True
8800              Exit For
8810            End If
8820          End With  ' ** doc.
8830        Next  ' ** doc.
8840      End With  ' ** cntr.
8850    End With  ' ** dbs.

EXITP:
8860    Set doc = Nothing
8870    Set cntr = Nothing
8880    Set dbs = Nothing
8890    Has_ChkReg = blnRetVal
8900    Exit Function

ERRH:
8910    blnRetVal = False
8920    Select Case ERR.Number
        Case Else
8930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8940    End Select
8950    Resume EXITP

End Function

Private Function Has_Form1099() As Boolean

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "Has_Form1099"

        Dim dbs As DAO.Database, cntr As DAO.Container, doc As DAO.Document
        Dim blnRetVal As Boolean

9010    blnRetVal = False

9020    Set dbs = CurrentDb
9030    With dbs
9040      Set cntr = .Containers("Forms")
9050      With cntr
9060        For Each doc In .Documents
9070          With doc
9080            If .Name = "frmRpt_Form1099" Then
9090              blnRetVal = True
9100              Exit For
9110            End If
9120          End With  ' ** doc.
9130        Next  ' ** doc.
9140      End With  ' ** cntr.
9150    End With  ' ** dbs.

EXITP:
9160    Set doc = Nothing
9170    Set cntr = Nothing
9180    Set dbs = Nothing
9190    Has_Form1099 = blnRetVal
9200    Exit Function

ERRH:
9210    blnRetVal = False
9220    Select Case ERR.Number
        Case Else
9230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9240    End Select
9250    Resume EXITP

End Function
