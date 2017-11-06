VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSiteMap_Journal_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmSiteMap_Journal_Sub"

'VGC 03/26/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmSiteMap_Journal}
' **   Exit:             X {cmdClose on frmSiteMap_Journal}

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
150       If strCallingForm = "frmSiteMap_Journal" Then

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
470         With Me
480           intRetVal = 0
490           .Parent.cmdClose_Click  ' ** Form Procedure: frmSiteMap_Journal.
500         End With
510       End Select
520     End If

        ' ** Alt keys.
530     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
540       Select Case intRetVal
          Case vbKeyX
550         With Me
560           intRetVal = 0
570           .Parent.cmdClose_Click  ' ** Form Procedure: frmSiteMap_Journal.
580         End With
590       End Select
600     End If

        ' ** Ctrl keys.
610     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
620       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
630         With Me
640           intRetVal = 0
650           DoCmd.SelectObject acForm, .Parent.Name, False
660           .Parent.cmdClose.SetFocus
670         End With
680       End Select
690     End If

        ' ** Ctrl-Shift keys.
700     If intCtrlDown And (Not intAltDown) And intShiftDown Then
710       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
720         With Me
730           intRetVal = 0
740           DoCmd.SelectObject acForm, .Parent.Name, False
750           .Parent.cmdClose.SetFocus
760         End With
770       End Select
780     End If

EXITP:
790     KeyCode = intRetVal
800     Exit Sub

ERRH:
810     intRetVal = 0
820     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
830     Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_MouseWheel"

910     With Me
920       If Count > 0 Then
930         Select Case strFocusHolder
            Case vbNullString
940           .FocusHolder01.SetFocus  ' ** Start with the first.
950         Case "FocusHolder01"
960           .FocusHolder07.SetFocus  ' ** To 1st bottom one.
970         Case "FocusHolder02"
980           .FocusHolder07.SetFocus  ' ** To 1st bottom one.
990         Case "FocusHolder03"
1000          .FocusHolder07.SetFocus  ' ** To 1st bottom one.
1010        Case "FocusHolder04"
1020          .FocusHolder07.SetFocus  ' ** To 1st bottom one.
1030        Case "FocusHolder05"
1040          .FocusHolder07.SetFocus  ' ** To 1st bottom one.
1050        Case "FocusHolder06"
1060          .FocusHolder07.SetFocus  ' ** To 1st bottom one.
1070        Case "FocusHolder07"
1080          .FocusHolder08.SetFocus  ' ** To 2nd bottom one.
1090        Case "FocusHolder08"
1100          .FocusHolder09.SetFocus  ' ** To 3rd bottom one.
1110        Case "FocusHolder09"
1120          .FocusHolder10.SetFocus  ' ** To 4th bottom one.
1130        Case "FocusHolder10"
1140          .FocusHolder11.SetFocus  ' ** To 5th bottom one.
1150        Case "FocusHolder11"
1160          .FocusHolder12.SetFocus  ' ** To 6th bottom one.
1170        Case "FocusHolder12"
1180          .FocusHolder12.SetFocus  ' ** Stay put.
1190        End Select
1200      Else  ' ** < 0.
1210        Select Case strFocusHolder
            Case vbNullString
1220          .FocusHolder01.SetFocus  ' ** Start with the first.
1230        Case "FocusHolder01"
1240          .FocusHolder01.SetFocus  ' ** Stay put.
1250        Case "FocusHolder02"
1260          .FocusHolder01.SetFocus  ' ** To 1st top one.
1270        Case "FocusHolder03"
1280          .FocusHolder02.SetFocus  ' ** To 2nd top one.
1290        Case "FocusHolder04"
1300          .FocusHolder03.SetFocus  ' ** To 3rd top one.
1310        Case "FocusHolder05"
1320          .FocusHolder04.SetFocus  ' ** To 4th top one.
1330        Case "FocusHolder06"
1340          .FocusHolder05.SetFocus  ' ** To 5th top one.
1350        Case "FocusHolder07"
1360          .FocusHolder06.SetFocus  ' ** To 6th top one.
1370        Case "FocusHolder08"
1380          .FocusHolder06.SetFocus  ' ** To 6th top one.
1390        Case "FocusHolder09"
1400          .FocusHolder06.SetFocus  ' ** To 6th top one.
1410        Case "FocusHolder10"
1420          .FocusHolder06.SetFocus  ' ** To 6th top one.
1430        Case "FocusHolder11"
1440          .FocusHolder06.SetFocus  ' ** To 6th top one.
1450        Case "FocusHolder12"
1460          .FocusHolder06.SetFocus  ' ** To 6th top one.
1470        End Select
1480      End If
          'Debug.Print "'COUNT: " & CStr(Count)
1490    End With

EXITP:
1500    Exit Sub

ERRH:
1510    Select Case ERR.Number
        Case Else
1520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub FocusHolder01_GotFocus()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder01_GotFocus"

1610    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
1620    Exit Sub

ERRH:
1630    Select Case ERR.Number
        Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Private Sub FocusHolder01_KeyDown(KeyCode As Integer, Shift As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder01_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1720    intShiftDown = (Shift And acShiftMask) > 0
1730    intAltDown = (Shift And acAltMask) > 0
1740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1770        With Me
1780          intRetVal = 0
1790          .FocusHolder02.SetFocus
1800        End With
1810      End Select
1820    End If

        ' ** Shift keys.
1830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1850        With Me
1860          intRetVal = 0
1870          DoCmd.SelectObject acForm, .Parent.Name, False
1880          .Parent.cmdClose.SetFocus
1890        End With
1900      End Select
1910    End If

EXITP:
1920    KeyCode = intRetVal
1930    Exit Sub

ERRH:
1940    intRetVal = 0
1950    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1960    Case Else
1970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1980    End Select
1990    Resume EXITP

End Sub

Private Sub FocusHolder02_GotFocus()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder02_GotFocus"

2010    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Private Sub FocusHolder03_GotFocus()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder03_GotFocus"

2110    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub FocusHolder04_GotFocus()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder04_GotFocus"

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

Private Sub FocusHolder05_GotFocus()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder05_GotFocus"

2310    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Private Sub FocusHolder06_GotFocus()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder06_GotFocus"

2410    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub FocusHolder07_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder07_GotFocus"

2510    strFocusHolder = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub FocusHolder08_GotFocus()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder08_GotFocus"

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

Private Sub FocusHolder09_GotFocus()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder09_GotFocus"

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

Private Sub FocusHolder10_GotFocus()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder10_GotFocus"

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

Private Sub FocusHolder11_GotFocus()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder11_GotFocus"

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

Private Sub FocusHolder12_GotFocus()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder12_GotFocus"

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

Private Sub FocusHolder12_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder12_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3120    intShiftDown = (Shift And acShiftMask) > 0
3130    intAltDown = (Shift And acAltMask) > 0
3140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3170        With Me
3180          intRetVal = 0
3190          DoCmd.SelectObject acForm, .Parent.Name, False
3200          .Parent.cmdClose.SetFocus
3210        End With
3220      End Select
3230    End If

        ' ** Shift keys.
3240    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3250      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3260        With Me
3270          intRetVal = 0
3280          .FocusHolder11.SetFocus
3290        End With
3300      End Select
3310    End If

EXITP:
3320    KeyCode = intRetVal
3330    Exit Sub

ERRH:
3340    intRetVal = 0
3350    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3360    Case Else
3370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3380    End Select
3390    Resume EXITP

End Sub

Private Sub SetUnavailable()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "SetUnavailable"

        Dim blnRetVal As Boolean

3410    With Me
3420      blnRetVal = Has_LTCL_STCGL  ' ** Function: Below.
3430      Select Case blnRetVal
          Case True
            ' ** Classic.
3440        .Level02_02_01_s02e_02_lbl.ForeColor = CLR_VDKGRY
3450        .Level02_02_01_s02e_02_lbl_dim_hi.Visible = False
3460        .Level02_02_01_s02e_02_lbl_box1.BackColor = CLR_LTTEAL2
3470        .Level02_02_01_s02e_02_s01_lbl.ForeColor = CLR_DKGRY2
3480        .Level02_02_01_s02e_02_s01_lbl_dim_hi.Visible = False
3490        .Level02_02_01_s02e_02_s01_lbl_box1.BackColor = CLR_LTTEAL2
3500        .Level02_02_01_s02e_02_arrow.ForeColor = CLR_DKGRY
3510        .Level02_02_01_s02e_02_arrow_shadow.ForeColor = CLR_GRY3
3520        .Level02_02_01_s02e_02_line1.BorderColor = CLR_DKGRY
3530        .Level02_02_01_s02e_02_line1_shadow.BorderColor = CLR_GRY3
3540        .Level02_02_01_s02e_02_line2.BorderColor = CLR_DKGRY
3550        .Level02_02_01_s02e_02_line2_shadow.BorderColor = CLR_GRY3
3560        .Level02_02_01_s02e_03_lbl.ForeColor = CLR_VDKGRY
3570        .Level02_02_01_s02e_03_lbl_dim_hi.Visible = False
3580        .Level02_02_01_s02e_03_lbl_box1.BackColor = CLR_LTTEAL2
3590        .Level02_02_01_s02e_03_s01_lbl.ForeColor = CLR_DKGRY2
3600        .Level02_02_01_s02e_03_s01_lbl_dim_hi.Visible = False
3610        .Level02_02_01_s02e_03_s01_lbl_box1.BackColor = CLR_LTTEAL2
3620        .Level02_02_01_s02e_03_s02_lbl.ForeColor = CLR_DKGRY2
3630        .Level02_02_01_s02e_03_s02_lbl_dim_hi.Visible = False
3640        .Level02_02_01_s02e_03_s02_lbl_box1.BackColor = CLR_LTTEAL2
3650        .Level02_02_01_s02e_03_arrow1.ForeColor = CLR_DKGRY
3660        .Level02_02_01_s02e_03_arrow1_shadow.ForeColor = CLR_GRY3
3670        .Level02_02_01_s02e_03_arrow2.ForeColor = CLR_DKGRY
3680        .Level02_02_01_s02e_03_arrow2_shadow.ForeColor = CLR_GRY3
3690        .Level02_02_01_s02e_03_line1.BorderColor = CLR_DKGRY
3700        .Level02_02_01_s02e_03_line1_shadow.BorderColor = CLR_GRY3
3710        .Level02_02_01_s02e_03_line2.BorderColor = CLR_DKGRY
3720        .Level02_02_01_s02e_03_line2_shadow.BorderColor = CLR_GRY3
3730        .Level02_02_01_s02e_03_line3.BorderColor = CLR_DKGRY
3740        .Level02_02_01_s02e_03_line3_shadow.BorderColor = CLR_GRY3
            ' ** Columnar.
3750        .Level02_02_01_s03_04_lbl.ForeColor = CLR_VDKGRY
3760        .Level02_02_01_s03_04_lbl_dim_hi.Visible = False
3770        .Level02_02_01_s03_04_lbl_box1.BackColor = CLR_LTTEAL2
3780        .Level02_02_01_s03_04_s01_lbl.ForeColor = CLR_DKGRY2
3790        .Level02_02_01_s03_04_s01_lbl_dim_hi.Visible = False
3800        .Level02_02_01_s03_04_s01_lbl_box1.BackColor = CLR_LTTEAL2
3810        .Level02_02_01_s03_04_arrow.ForeColor = CLR_DKGRY
3820        .Level02_02_01_s03_04_arrow_shadow.ForeColor = CLR_GRY3
3830        .Level02_02_01_s03_04_line1.BorderColor = CLR_DKGRY
3840        .Level02_02_01_s03_04_line1_shadow.BorderColor = CLR_GRY3
3850        .Level02_02_01_s03_04_line2.BorderColor = CLR_DKGRY
3860        .Level02_02_01_s03_04_line2_shadow.BorderColor = CLR_GRY3
3870        .Level02_02_01_s03_05_lbl.ForeColor = CLR_VDKGRY
3880        .Level02_02_01_s03_05_lbl_dim_hi.Visible = False
3890        .Level02_02_01_s03_05_lbl_box1.BackColor = CLR_LTTEAL2
3900        .Level02_02_01_s03_05_s01_lbl.ForeColor = CLR_DKGRY2
3910        .Level02_02_01_s03_05_s01_lbl_dim_hi.Visible = False
3920        .Level02_02_01_s03_05_s01_lbl_box1.BackColor = CLR_LTTEAL2
3930        .Level02_02_01_s03_05_s02_lbl.ForeColor = CLR_DKGRY2
3940        .Level02_02_01_s03_05_s02_lbl_dim_hi.Visible = False
3950        .Level02_02_01_s03_05_s02_lbl_box1.BackColor = CLR_LTTEAL2
3960        .Level02_02_01_s03_05_arrow1.ForeColor = CLR_DKGRY
3970        .Level02_02_01_s03_05_arrow1_shadow.ForeColor = CLR_GRY3
3980        .Level02_02_01_s03_05_arrow2.ForeColor = CLR_DKGRY
3990        .Level02_02_01_s03_05_arrow2_shadow.ForeColor = CLR_GRY3
4000        .Level02_02_01_s03_05_line1.BorderColor = CLR_DKGRY
4010        .Level02_02_01_s03_05_line1_shadow.BorderColor = CLR_GRY3
4020        .Level02_02_01_s03_05_line2.BorderColor = CLR_DKGRY
4030        .Level02_02_01_s03_05_line2_shadow.BorderColor = CLR_GRY3
4040        .Level02_02_01_s03_05_line3.BorderColor = CLR_DKGRY
4050        .Level02_02_01_s03_05_line3_shadow.BorderColor = CLR_GRY3
4060      Case False
            ' ** Classic.
4070        .Level02_02_01_s02e_02_lbl.ForeColor = WIN_CLR_DISF
4080        .Level02_02_01_s02e_02_lbl_dim_hi.Visible = True
4090        .Level02_02_01_s02e_02_lbl_box1.BackColor = WIN_CLR_DISR
4100        .Level02_02_01_s02e_02_s01_lbl.ForeColor = WIN_CLR_DISF
4110        .Level02_02_01_s02e_02_s01_lbl_dim_hi.Visible = True
4120        .Level02_02_01_s02e_02_s01_lbl_box1.BackColor = WIN_CLR_DISR
4130        .Level02_02_01_s02e_02_arrow.ForeColor = WIN_CLR_DISF
4140        .Level02_02_01_s02e_02_arrow_shadow.ForeColor = CLR_LTGRY
4150        .Level02_02_01_s02e_02_line1.BorderColor = WIN_CLR_DISF
4160        .Level02_02_01_s02e_02_line1_shadow.BorderColor = CLR_LTGRY
4170        .Level02_02_01_s02e_02_line2.BorderColor = WIN_CLR_DISF
4180        .Level02_02_01_s02e_02_line2_shadow.BorderColor = CLR_LTGRY
4190        .Level02_02_01_s02e_03_lbl.ForeColor = WIN_CLR_DISF
4200        .Level02_02_01_s02e_03_lbl_dim_hi.Visible = True
4210        .Level02_02_01_s02e_03_lbl_box1.BackColor = WIN_CLR_DISR
4220        .Level02_02_01_s02e_03_s01_lbl.ForeColor = WIN_CLR_DISF
4230        .Level02_02_01_s02e_03_s01_lbl_dim_hi.Visible = True
4240        .Level02_02_01_s02e_03_s01_lbl_box1.BackColor = WIN_CLR_DISR
4250        .Level02_02_01_s02e_03_s02_lbl.ForeColor = WIN_CLR_DISF
4260        .Level02_02_01_s02e_03_s02_lbl_dim_hi.Visible = True
4270        .Level02_02_01_s02e_03_s02_lbl_box1.BackColor = WIN_CLR_DISR
4280        .Level02_02_01_s02e_03_arrow1.ForeColor = WIN_CLR_DISF
4290        .Level02_02_01_s02e_03_arrow1_shadow.ForeColor = CLR_LTGRY
4300        .Level02_02_01_s02e_03_arrow2.ForeColor = WIN_CLR_DISF
4310        .Level02_02_01_s02e_03_arrow2_shadow.ForeColor = CLR_LTGRY
4320        .Level02_02_01_s02e_03_line1.BorderColor = WIN_CLR_DISF
4330        .Level02_02_01_s02e_03_line1_shadow.BorderColor = CLR_LTGRY
4340        .Level02_02_01_s02e_03_line2.BorderColor = WIN_CLR_DISF
4350        .Level02_02_01_s02e_03_line2_shadow.BorderColor = CLR_LTGRY
4360        .Level02_02_01_s02e_03_line3.BorderColor = WIN_CLR_DISF
4370        .Level02_02_01_s02e_03_line3_shadow.BorderColor = CLR_LTGRY
            ' ** Columnar.
4380        .Level02_02_01_s03_04_lbl.ForeColor = WIN_CLR_DISF
4390        .Level02_02_01_s03_04_lbl_dim_hi.Visible = True
4400        .Level02_02_01_s03_04_lbl_box1.BackColor = WIN_CLR_DISR
4410        .Level02_02_01_s03_04_s01_lbl.ForeColor = WIN_CLR_DISF
4420        .Level02_02_01_s03_04_s01_lbl_dim_hi.Visible = True
4430        .Level02_02_01_s03_04_s01_lbl_box1.BackColor = WIN_CLR_DISR
4440        .Level02_02_01_s03_04_arrow.ForeColor = WIN_CLR_DISF
4450        .Level02_02_01_s03_04_arrow_shadow.ForeColor = CLR_LTGRY
4460        .Level02_02_01_s03_04_line1.BorderColor = WIN_CLR_DISF
4470        .Level02_02_01_s03_04_line1_shadow.BorderColor = CLR_LTGRY
4480        .Level02_02_01_s03_04_line2.BorderColor = WIN_CLR_DISF
4490        .Level02_02_01_s03_04_line2_shadow.BorderColor = CLR_LTGRY
4500        .Level02_02_01_s03_05_lbl.ForeColor = WIN_CLR_DISF
4510        .Level02_02_01_s03_05_lbl_dim_hi.Visible = True
4520        .Level02_02_01_s03_05_lbl_box1.BackColor = WIN_CLR_DISR
4530        .Level02_02_01_s03_05_s01_lbl.ForeColor = WIN_CLR_DISF
4540        .Level02_02_01_s03_05_s01_lbl_dim_hi.Visible = True
4550        .Level02_02_01_s03_05_s01_lbl_box1.BackColor = WIN_CLR_DISR
4560        .Level02_02_01_s03_05_s02_lbl.ForeColor = WIN_CLR_DISF
4570        .Level02_02_01_s03_05_s02_lbl_dim_hi.Visible = True
4580        .Level02_02_01_s03_05_s02_lbl_box1.BackColor = WIN_CLR_DISR
4590        .Level02_02_01_s03_05_arrow1.ForeColor = WIN_CLR_DISF
4600        .Level02_02_01_s03_05_arrow1_shadow.ForeColor = CLR_LTGRY
4610        .Level02_02_01_s03_05_arrow2.ForeColor = WIN_CLR_DISF
4620        .Level02_02_01_s03_05_arrow2_shadow.ForeColor = CLR_LTGRY
4630        .Level02_02_01_s03_05_line1.BorderColor = WIN_CLR_DISF
4640        .Level02_02_01_s03_05_line1_shadow.BorderColor = CLR_LTGRY
4650        .Level02_02_01_s03_05_line2.BorderColor = WIN_CLR_DISF
4660        .Level02_02_01_s03_05_line2_shadow.BorderColor = CLR_LTGRY
4670        .Level02_02_01_s03_05_line3.BorderColor = WIN_CLR_DISF
4680        .Level02_02_01_s03_05_line3_shadow.BorderColor = CLR_LTGRY
4690      End Select
4700    End With

EXITP:
4710    Exit Sub

ERRH:
4720    Select Case ERR.Number
        Case Else
4730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4740    End Select
4750    Resume EXITP

End Sub

Private Function Has_LTCL_STCGL() As Boolean

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "Has_LTCL_STCGL"

        Dim dbs As DAO.Database, cntr As DAO.Container, doc As DAO.Document
        Dim blnRetVal As Boolean

4810   blnRetVal = False

4820    Set dbs = CurrentDb
4830    With dbs
4840      Set cntr = .Containers("Forms")
4850      With cntr
4860        For Each doc In .Documents
4870          With doc
4880            If .Name = "frmMap_Misc_LTCL" Or .Name = "frmMap_Misc_STCGL" Then
4890              blnRetVal = True
4900              Exit For
4910            End If
4920          End With  ' ** doc.
4930        Next  ' ** doc.
4940      End With  ' ** cntr.
4950    End With  ' ** dbs.

EXITP:
4960    Set doc = Nothing
4970    Set cntr = Nothing
4980    Set dbs = Nothing
4990    Has_LTCL_STCGL = blnRetVal
5000    Exit Function

ERRH:
5010    blnRetVal = False
5020    Select Case ERR.Number
        Case Else
5030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5040    End Select
5050    Resume EXITP

End Function
