VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountProfile_ReviewFreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountProfile_ReviewFreq"

'VGC 03/22/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Select None:      N {cmdSelectNone}
' **   OK:               O {cmdOK}
' **   Select All:       S {cmdSelectAll}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmAccountProfile_ReviewFreq_Sub}

Private blnSelectAll_Focus As Boolean, blnSelectAll_MouseDown As Boolean
Private blnSelectNone_Focus As Boolean, blnSelectNone_MouseDown As Boolean

' ** Array: arr_varMonth().
Private lngMonths As Long, arr_varMonth() As Variant
'Private Const M_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const M_CNAM As Integer = 0
Private Const M_VAL  As Integer = 1

Private strAccountNo As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim usr As DAO.User, grp As DAO.Group, frm As Access.Form, ctl As Access.Control
        Dim blnFound As Boolean
        Dim intPos01 As Integer

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)
150           intPos01 = InStr(strCallingForm, "~")
160           strAccountNo = Mid(strCallingForm, (intPos01 + 1))
170           strCallingForm = Left(strCallingForm, (intPos01 - 1))

180           blnSelectAll_Focus = False: blnSelectAll_MouseDown = False
190           blnSelectNone_Focus = False: blnSelectNone_MouseDown = False

200           .cmdOK.Enabled = False

210           lngMonths = 12&
220           ReDim arr_varMonth(1, 12)

              ' ** Check user status, and lock all fields if ViewOnly.
230           blnFound = False
240           For Each grp In DBEngine.Workspaces(0).Groups
250             If grp.Name = "ViewOnly" Then
260               For Each usr In grp.Users
270                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
280                   blnFound = True
290                   Exit For
300                 End If
310               Next
320             End If
330           Next

340           Set frm = .frmAccountProfile_ReviewFreq_Sub.Form

350           If blnFound = True Then
360             .cmdCancel.SetFocus
370             For Each ctl In frm.Detail.Controls
380               With ctl
390                 If .ControlType = acCheckBox Then
400                   .Locked = True
410                 End If
420               End With
430             Next
440           End If

              ' ** Gather original values.
450           With frm
460             arr_varMonth(M_CNAM, 1) = .revjan.Name
470             arr_varMonth(M_VAL, 1) = .revjan
480             arr_varMonth(M_CNAM, 2) = .revfeb.Name
490             arr_varMonth(M_VAL, 2) = .revfeb
500             arr_varMonth(M_CNAM, 3) = .revmar.Name
510             arr_varMonth(M_VAL, 3) = .revmar
520             arr_varMonth(M_CNAM, 4) = .revapr.Name
530             arr_varMonth(M_VAL, 4) = .revapr
540             arr_varMonth(M_CNAM, 5) = .revmay.Name
550             arr_varMonth(M_VAL, 5) = .revmay
560             arr_varMonth(M_CNAM, 6) = .revjun.Name
570             arr_varMonth(M_VAL, 6) = .revjun
580             arr_varMonth(M_CNAM, 7) = .revjul.Name
590             arr_varMonth(M_VAL, 7) = .revjul
600             arr_varMonth(M_CNAM, 8) = .revaug.Name
610             arr_varMonth(M_VAL, 8) = .revaug
620             arr_varMonth(M_CNAM, 9) = .revsep.Name
630             arr_varMonth(M_VAL, 9) = .revsep
640             arr_varMonth(M_CNAM, 10) = .revoct.Name
650             arr_varMonth(M_VAL, 10) = .revoct
660             arr_varMonth(M_CNAM, 11) = .revnov.Name
670             arr_varMonth(M_VAL, 11) = .revnov
680             arr_varMonth(M_CNAM, 12) = .revdec.Name
690             arr_varMonth(M_VAL, 12) = .revdec
700             .MonthVals_Set arr_varMonth  ' ** Form Procedure: frmAccountProfile_ReviewFreq_Sub.
710           End With

720           .accountno = strAccountNo

730         Else
740           Cancel = -1
750         End If
760       Else
770         Cancel = -1
780       End If
790     End With

800     If Cancel = -1 Then
810       Beep
820       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
830     End If

EXITP:
840     Set ctl = Nothing
850     Set frm = Nothing
860     Set usr = Nothing
870     Set grp = Nothing
880     Exit Sub

ERRH:
890     Select Case ERR.Number
        Case Else
900       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
910     End Select
920     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1020    intShiftDown = (Shift And acShiftMask) > 0
1030    intAltDown = (Shift And acAltMask) > 0
1040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmAccountProfile_ReviewFreq_Sub}

        ' ** Plain keys.
1050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1060      Select Case intRetVal
          Case vbKeyEscape
1070        intRetVal = 0
1080        cmdCancel_Click  ' ** Procedure: Below.
1090      End Select
1100    End If

        ' ** Ctrl keys.
1110    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1120      Select Case intRetVal
          Case vbKeyS
1130        With Me
1140          intRetVal = 0
1150          .frmAccountProfile_ReviewFreq_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAccountProfile_ReviewFreq_Sub.
1160        End With
1170      End Select
1180    End If

        ' ** Ctrl-Shift keys.
1190    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1200      Select Case intRetVal
          Case vbKeyF
1210        With Me
1220          intRetVal = 0
1230          .FocusHolder.SetFocus
1240        End With
1250      End Select
1260    End If

EXITP:
1270    KeyCode = intRetVal
1280    Exit Sub

ERRH:
1290    intRetVal = 0
1300    Select Case ERR.Number
        Case Else
1310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1320    End Select
1330    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1410    Pref_Save THIS_NAME & "_Sub"  ' ** Module Procedure: modPreferenceFuncs.

EXITP:
1420    Exit Sub

ERRH:
1430    Select Case ERR.Number
        Case Else
1440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1450    End Select
1460    Resume EXITP

End Sub

Public Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1510    With Me
1520      If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
1530        Select Case blnSelectAll_Focus
            Case True
1540          .cmdSelectAll_raised_semifocus_dots_img.Visible = True
1550          .cmdSelectAll_raised_img.Visible = False
1560        Case False
1570          .cmdSelectAll_raised_img.Visible = True
1580          .cmdSelectAll_raised_semifocus_dots_img.Visible = False
1590        End Select
1600        .cmdSelectAll_raised_focus_img.Visible = False
1610        .cmdSelectAll_raised_focus_dots_img.Visible = False
1620        .cmdSelectAll_sunken_focus_dots_img.Visible = False
1630        .cmdSelectAll_raised_img_dis.Visible = False
1640      End If
1650      If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
1660        Select Case blnSelectNone_Focus
            Case True
1670          .cmdSelectNone_raised_semifocus_dots_img.Visible = True
1680          .cmdSelectNone_raised_img.Visible = False
1690        Case False
1700          .cmdSelectNone_raised_img.Visible = True
1710          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
1720        End Select
1730        .cmdSelectNone_raised_focus_img.Visible = False
1740        .cmdSelectNone_raised_focus_dots_img.Visible = False
1750        .cmdSelectNone_sunken_focus_dots_img.Visible = False
1760        .cmdSelectNone_raised_img_dis.Visible = False
1770      End If
1780    End With

EXITP:
1790    Exit Sub

ERRH:
1800    Select Case ERR.Number
        Case Else
1810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1820    End Select
1830    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

1910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
1920    Exit Sub

ERRH:
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Private Sub cmdOK_Click()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

2010    With Me

2020      .frmAccountProfile_ReviewFreq_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAccountProfile_ReviewFreq_Sub.

2030      DoCmd.Close acForm, THIS_NAME

2040    End With

EXITP:
2050    Exit Sub

ERRH:
2060    Select Case ERR.Number
        Case Else
2070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2080    End Select
2090    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
2190          .cmdCancel.SetFocus
2200        End With
2210      End Select
2220    End If

        ' ** Shift keys.
2230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyTab
2250        With Me
2260          intRetVal = 0
2270          .cmdSelectNone.SetFocus
2280        End With
2290      End Select
2300    End If

EXITP:
2310    KeyCode = intRetVal
2320    Exit Sub

ERRH:
2330    intRetVal = 0
2340    Select Case ERR.Number
        Case Else
2350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2360    End Select
2370    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim frm As Access.Form
        Dim blnContinue As Boolean
        Dim lngX As Long

2410    With Me
2420      blnContinue = True
2430      If .changed = True Then
2440        Beep
2450        If MsgBox("You have made changes, but have not confirmed them by clicking OK." & vbCrLf & vbCrLf & _
                "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton1, "Abandon Changes") = vbYes Then
              ' ** Abandon changes.
2460        Else
2470          blnContinue = False
2480        End If
2490        If blnContinue = True Then
              ' ** Undo changes.
2500          Set frm = .frmAccountProfile_ReviewFreq_Sub.Form
2510          With frm
2520            For lngX = 1& To lngMonths
2530              .Controls(arr_varMonth(M_CNAM, lngX)) = arr_varMonth(M_VAL, lngX)
2540            Next
2550            .cmdSave_Click  ' ** Form Procedure: frmAccountProfile_ReviewFreq_Sub.
2560          End With
2570        End If
2580      End If
2590      DoEvents
2600      If blnContinue = True Then
2610        DoCmd.Close acForm, THIS_NAME
2620      End If
2630    End With

EXITP:
2640    Set frm = Nothing
2650    Exit Sub

ERRH:
2660    Select Case ERR.Number
        Case Else
2670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2680    End Select
2690    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyTab
2770        With Me
2780          intRetVal = 0
2790          .frmAccountProfile_ReviewFreq_Sub.SetFocus
2800          .frmAccountProfile_ReviewFreq_Sub.Form.revjan.SetFocus
2810        End With
2820      End Select
2830    End If

        ' ** Shift keys.
2840    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2850      Select Case intRetVal
          Case vbKeyTab
2860        With Me
2870          intRetVal = 0
2880          If .cmdOK.Enabled = True Then
2890            .cmdOK.SetFocus
2900          Else
2910            .cmdSelectNone.SetFocus
2920          End If
2930        End With
2940      End Select
2950    End If

EXITP:
2960    KeyCode = intRetVal
2970    Exit Sub

ERRH:
2980    intRetVal = 0
2990    Select Case ERR.Number
        Case Else
3000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3010    End Select
3020    Resume EXITP

End Sub

Private Sub cmdSelectAll_GotFocus()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_GotFocus"

3110    With Me
3120      blnSelectAll_Focus = True
3130      .cmdSelectAll_raised_semifocus_dots_img.Visible = True
3140      .cmdSelectAll_raised_img.Visible = False
3150      .cmdSelectAll_raised_focus_img.Visible = False
3160      .cmdSelectAll_raised_focus_dots_img.Visible = False
3170      .cmdSelectAll_sunken_focus_dots_img.Visible = False
3180      .cmdSelectAll_raised_img_dis.Visible = False
3190    End With

EXITP:
3200    Exit Sub

ERRH:
3210    Select Case ERR.Number
        Case Else
3220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3230    End Select
3240    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseDown"

3310    With Me
3320      blnSelectAll_MouseDown = True
3330      .cmdSelectAll_sunken_focus_dots_img.Visible = True
3340      .cmdSelectAll_raised_img.Visible = False
3350      .cmdSelectAll_raised_semifocus_dots_img.Visible = False
3360      .cmdSelectAll_raised_focus_img.Visible = False
3370      .cmdSelectAll_raised_focus_dots_img.Visible = False
3380      .cmdSelectAll_raised_img_dis.Visible = False
3390    End With

EXITP:
3400    Exit Sub

ERRH:
3410    Select Case ERR.Number
        Case Else
3420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3430    End Select
3440    Resume EXITP

End Sub

Private Sub cmdSelectAll_Click()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_Click"

        Dim ctl As Access.Control

3510    With Me
3520      For Each ctl In .frmAccountProfile_ReviewFreq_Sub.Form.Controls
3530        With ctl
3540          If .ControlType = acCheckBox Then
3550            .Value = True
3560          End If
3570        End With
3580      Next
3590      .frmAccountProfile_ReviewFreq_Sub.Form.ShowMonth  ' ** Form Procedure: frmAccountProfile_ReviewFreq_Sub.
3600    End With

EXITP:
3610    Set ctl = Nothing
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub cmdSelectAll_KeyDown(KeyCode As Integer, Shift As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_KeyDown"

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
          Case vbKeyTab
3770        With Me
3780          intRetVal = 0
3790          .cmdSelectNone.SetFocus
3800        End With
3810      End Select
3820    End If

        ' ** Shift keys.
3830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3840      Select Case intRetVal
          Case vbKeyTab
3850        With Me
3860          intRetVal = 0
3870          .frmAccountProfile_ReviewFreq_Sub.SetFocus
3880          .frmAccountProfile_ReviewFreq_Sub.Form.revdec.SetFocus
3890        End With
3900      End Select
3910    End If

EXITP:
3920    KeyCode = intRetVal
3930    Exit Sub

ERRH:
3940    intRetVal = 0
3950    Select Case ERR.Number
        Case Else
3960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3970    End Select
3980    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseMove"

4010    With Me
4020      If blnSelectAll_MouseDown = False Then
4030        Select Case blnSelectAll_Focus
            Case True
4040          .cmdSelectAll_raised_focus_dots_img.Visible = True
4050          .cmdSelectAll_raised_focus_img.Visible = False
4060        Case False
4070          .cmdSelectAll_raised_focus_img.Visible = True
4080          .cmdSelectAll_raised_focus_dots_img.Visible = False
4090        End Select
4100        .cmdSelectAll_raised_img.Visible = False
4110        .cmdSelectAll_raised_semifocus_dots_img.Visible = False
4120        .cmdSelectAll_sunken_focus_dots_img.Visible = False
4130        .cmdSelectAll_raised_img_dis.Visible = False
4140      End If
4150      If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
4160        Select Case blnSelectNone_Focus
            Case True
4170          .cmdSelectNone_raised_semifocus_dots_img.Visible = True
4180          .cmdSelectNone_raised_img.Visible = False
4190        Case False
4200          .cmdSelectNone_raised_img.Visible = True
4210          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
4220        End Select
4230        .cmdSelectNone_raised_focus_img.Visible = False
4240        .cmdSelectNone_raised_focus_dots_img.Visible = False
4250        .cmdSelectNone_sunken_focus_dots_img.Visible = False
4260        .cmdSelectNone_raised_img_dis.Visible = False
4270      End If
4280    End With

EXITP:
4290    Exit Sub

ERRH:
4300    Select Case ERR.Number
        Case Else
4310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4320    End Select
4330    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseUp"

4410    With Me
4420      .cmdSelectAll_raised_focus_dots_img.Visible = True
4430      .cmdSelectAll_raised_img.Visible = False
4440      .cmdSelectAll_raised_semifocus_dots_img.Visible = False
4450      .cmdSelectAll_raised_focus_img.Visible = False
4460      .cmdSelectAll_sunken_focus_dots_img.Visible = False
4470      .cmdSelectAll_raised_img_dis.Visible = False
4480      blnSelectAll_MouseDown = False
4490    End With

EXITP:
4500    Exit Sub

ERRH:
4510    Select Case ERR.Number
        Case Else
4520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4530    End Select
4540    Resume EXITP

End Sub

Private Sub cmdSelectAll_LostFocus()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_LostFocus"

4610    With Me
4620      .cmdSelectAll_raised_img.Visible = True
4630      .cmdSelectAll_raised_semifocus_dots_img.Visible = False
4640      .cmdSelectAll_raised_focus_img.Visible = False
4650      .cmdSelectAll_raised_focus_dots_img.Visible = False
4660      .cmdSelectAll_sunken_focus_dots_img.Visible = False
4670      .cmdSelectAll_raised_img_dis.Visible = False
4680      blnSelectAll_Focus = False
4690    End With

EXITP:
4700    Exit Sub

ERRH:
4710    Select Case ERR.Number
        Case Else
4720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4730    End Select
4740    Resume EXITP

End Sub

Private Sub cmdSelectNone_GotFocus()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_GotFocus"

4810    With Me
4820      blnSelectNone_Focus = True
4830      .cmdSelectNone_raised_semifocus_dots_img.Visible = True
4840      .cmdSelectNone_raised_img.Visible = False
4850      .cmdSelectNone_raised_focus_img.Visible = False
4860      .cmdSelectNone_raised_focus_dots_img.Visible = False
4870      .cmdSelectNone_sunken_focus_dots_img.Visible = False
4880      .cmdSelectNone_raised_img_dis.Visible = False
4890    End With

EXITP:
4900    Exit Sub

ERRH:
4910    Select Case ERR.Number
        Case Else
4920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4930    End Select
4940    Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseDown"

5010    With Me
5020      blnSelectNone_MouseDown = True
5030      .cmdSelectNone_sunken_focus_dots_img.Visible = True
5040      .cmdSelectNone_raised_img.Visible = False
5050      .cmdSelectNone_raised_semifocus_dots_img.Visible = False
5060      .cmdSelectNone_raised_focus_img.Visible = False
5070      .cmdSelectNone_raised_focus_dots_img.Visible = False
5080      .cmdSelectNone_raised_img_dis.Visible = False
5090    End With

EXITP:
5100    Exit Sub

ERRH:
5110    Select Case ERR.Number
        Case Else
5120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5130    End Select
5140    Resume EXITP

End Sub

Private Sub cmdSelectNone_Click()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_Click"

        Dim ctl As Access.Control

5210    With Me
5220      For Each ctl In .frmAccountProfile_ReviewFreq_Sub.Form.Controls
5230        With ctl
5240          If .ControlType = acCheckBox Then
5250            .Value = False
5260          End If
5270        End With
5280      Next
5290      .frmAccountProfile_ReviewFreq_Sub.Form.ShowMonth  ' ** Form Procedure: frmAccountProfile_ReviewFreq_Sub.
5300    End With

EXITP:
5310    Set ctl = Nothing
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Private Sub cmdSelectNone_KeyDown(KeyCode As Integer, Shift As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5420    intShiftDown = (Shift And acShiftMask) > 0
5430    intAltDown = (Shift And acAltMask) > 0
5440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5460      Select Case intRetVal
          Case vbKeyTab
5470        With Me
5480          intRetVal = 0
5490          If .cmdOK.Enabled = True Then
5500            .cmdOK.SetFocus
5510          Else
5520            .cmdCancel.SetFocus
5530          End If
5540        End With
5550      End Select
5560    End If

        ' ** Shift keys.
5570    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5580      Select Case intRetVal
          Case vbKeyTab
5590        With Me
5600          intRetVal = 0
5610          .cmdSelectAll.SetFocus
5620        End With
5630      End Select
5640    End If

EXITP:
5650    KeyCode = intRetVal
5660    Exit Sub

ERRH:
5670    intRetVal = 0
5680    Select Case ERR.Number
        Case Else
5690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5700    End Select
5710    Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseMove"

5810    With Me
5820      If blnSelectNone_MouseDown = False Then
5830        Select Case blnSelectNone_Focus
            Case True
5840          .cmdSelectNone_raised_focus_dots_img.Visible = True
5850          .cmdSelectNone_raised_focus_img.Visible = False
5860        Case False
5870          .cmdSelectNone_raised_focus_img.Visible = True
5880          .cmdSelectNone_raised_focus_dots_img.Visible = False
5890        End Select
5900        .cmdSelectNone_raised_img.Visible = False
5910        .cmdSelectNone_raised_semifocus_dots_img.Visible = False
5920        .cmdSelectNone_sunken_focus_dots_img.Visible = False
5930        .cmdSelectNone_raised_img_dis.Visible = False
5940      End If
5950      If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
5960        Select Case blnSelectAll_Focus
            Case True
5970          .cmdSelectAll_raised_semifocus_dots_img.Visible = True
5980          .cmdSelectAll_raised_img.Visible = False
5990        Case False
6000          .cmdSelectAll_raised_img.Visible = True
6010          .cmdSelectAll_raised_semifocus_dots_img.Visible = False
6020        End Select
6030        .cmdSelectAll_raised_focus_img.Visible = False
6040        .cmdSelectAll_raised_focus_dots_img.Visible = False
6050        .cmdSelectAll_sunken_focus_dots_img.Visible = False
6060        .cmdSelectAll_raised_img_dis.Visible = False
6070      End If
6080    End With

EXITP:
6090    Exit Sub

ERRH:
6100    Select Case ERR.Number
        Case Else
6110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6120    End Select
6130    Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseUp"

6210    With Me
6220      .cmdSelectNone_raised_focus_dots_img.Visible = True
6230      .cmdSelectNone_raised_img.Visible = False
6240      .cmdSelectNone_raised_semifocus_dots_img.Visible = False
6250      .cmdSelectNone_raised_focus_img.Visible = False
6260      .cmdSelectNone_sunken_focus_dots_img.Visible = False
6270      .cmdSelectNone_raised_img_dis.Visible = False
6280      blnSelectNone_MouseDown = False
6290    End With

EXITP:
6300    Exit Sub

ERRH:
6310    Select Case ERR.Number
        Case Else
6320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6330    End Select
6340    Resume EXITP

End Sub

Private Sub cmdSelectNone_LostFocus()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_LostFocus"

6410    With Me
6420      .cmdSelectNone_raised_img.Visible = True
6430      .cmdSelectNone_raised_semifocus_dots_img.Visible = False
6440      .cmdSelectNone_raised_focus_img.Visible = False
6450      .cmdSelectNone_raised_focus_dots_img.Visible = False
6460      .cmdSelectNone_sunken_focus_dots_img.Visible = False
6470      .cmdSelectNone_raised_img_dis.Visible = False
6480      blnSelectNone_Focus = False
6490    End With

EXITP:
6500    Exit Sub

ERRH:
6510    Select Case ERR.Number
        Case Else
6520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6530    End Select
6540    Resume EXITP

End Sub

Public Sub MonthVals_Chk()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "MonthVals_Chk"

6610    With Me
6620      Select Case .changed
          Case True
6630        .cmdOK.Enabled = True
6640      Case False
6650        .cmdOK.Enabled = False
6660      End Select
6670    End With

EXITP:
6680    Exit Sub

ERRH:
6690    Select Case ERR.Number
        Case Else
6700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6710    End Select
6720    Resume EXITP

End Sub
