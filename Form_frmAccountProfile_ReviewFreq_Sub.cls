VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountProfile_ReviewFreq_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountProfile_ReviewFreq_Sub"

'VGC 03/22/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel on frmAccountProfile_ReviewFreq}
' **   Select None:      N {cmdSelectNone on frmAccountProfile_ReviewFreq}
' **   OK:               O {cmdOK on frmAccountProfile_ReviewFreq}
' **   Select All:       S {cmdSelectAll on frmAccountProfile_ReviewFreq}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ** NOTE: Prefs only used when called from frmAccountProfile_Add.

' ** Array: arr_varMonth().
Private lngMonths As Long, arr_varMonth As Variant
'Private Const M_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const M_CNAM As Integer = 0
Private Const M_VAL  As Integer = 1

Private strAccountNo As String
Private strCallingForm As String, strCallersCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim strFilter As String
        Dim intPos01 As Integer

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccountProfile_ReviewFreq" Then

160         strCallersCallingForm = .Parent.OpenArgs
170         intPos01 = InStr(strCallersCallingForm, "~")
180         strAccountNo = Mid(strCallersCallingForm, (intPos01 + 1))
190         strCallersCallingForm = Left(strCallersCallingForm, (intPos01 - 1))

200         Select Case strCallersCallingForm
            Case "frmAccountProfile"
210           If .RecordSource <> "qryAccountFreq_Rev_01_02" Then
220             .RecordSource = "qryAccountFreq_Rev_01_02"
230           End If
240         Case "frmAccountProfile_Add"
250           .RecordSource = "qryAccountFreq_Rev_02"
260           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
270         End Select

280         strFilter = "[accountno] = '" & strAccountNo & "'"
290         .Filter = strFilter
300         .FilterOn = True

310         ShowMonth  ' ** Procedure: Below.

320       Else
330         Cancel = -1
340       End If
350     End With

360     If Cancel = -1 Then
370       Beep
380       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
390     End If

EXITP:
400     Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         With Me
580           intRetVal = 0
590           DoCmd.SelectObject acForm, .Parent.Name, False
600           .Parent.cmdCancel_Click  ' ** Form Procedure: frmAccountProfile_ReviewFreq.
610         End With
620       End Select
630     End If

        ' ** Ctrl keys.
640     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
650       Select Case intRetVal
          Case vbKeyS
660         intRetVal = 0
670         cmdSave_Click  ' ** Procedure: Below.
680       End Select
690     End If

        ' ** Ctrl-Shift keys.
700     If intCtrlDown And (Not intAltDown) And intShiftDown Then
710       Select Case intRetVal
          Case vbKeyF
720         With Me
730           intRetVal = 0
740           DoCmd.SelectObject acForm, .Parent.Name, False
750           .Parent.FocusHolder.SetFocus
760         End With
770       End Select
780     End If

EXITP:
790     KeyCode = intRetVal
800     Exit Sub

ERRH:
810     intRetVal = 0
820     Select Case ERR.Number
        Case Else
830       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
840     End Select
850     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

910     If strCallersCallingForm = "frmAccountProfile_Add" Then
          ' ** Doesn't work from here, so moved to parent.
          ' ** Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
920     End If

EXITP:
930     Exit Sub

ERRH:
940     Select Case ERR.Number
        Case Else
950       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
960     End Select
970     Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1010    With Me
1020      .Parent.Detail_MouseMove Button, Shift, X, Y  ' ** Form Procedure: frmAccountProfile_ReviewFreq.
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

Public Sub cmdSave_Click()
' ** This is a fake click event.

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1110    DoCmd.RunCommand acCmdSaveRecord
1120    DoEvents

EXITP:
1130    Exit Sub

ERRH:
1140    Select Case ERR.Number
        Case Else
1150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1160    End Select
1170    Resume EXITP

End Sub

Private Sub revjan_AfterUpdate()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "revjan_AfterUpdate"

1210    ShowMonth  ' ** Procedure: Below.

EXITP:
1220    Exit Sub

ERRH:
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub revjan_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "revjan_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
1370        With Me
1380          intRetVal = 0
1390          .revfeb.SetFocus
1400        End With
1410      End Select
1420    End If

        ' ** Shift keys.
1430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1450        With Me
1460          intRetVal = 0
1470          DoCmd.SelectObject acForm, .Parent.Name, False
1480          .Parent.cmdCancel.SetFocus
1490        End With
1500      End Select
1510    End If

EXITP:
1520    KeyCode = intRetVal
1530    Exit Sub

ERRH:
1540    intRetVal = 0
1550    Select Case ERR.Number
        Case Else
1560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1570    End Select
1580    Resume EXITP

End Sub

Private Sub revfeb_AfterUpdate()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "revfeb_AfterUpdate"

1610    ShowMonth  ' ** Procedure: Below.

EXITP:
1620    Exit Sub

ERRH:
1630    Select Case ERR.Number
        Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Private Sub revmar_AfterUpdate()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "revmar_AfterUpdate"

1710    ShowMonth  ' ** Procedure: Below.

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub revapr_AfterUpdate()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "revapr_AfterUpdate"

1810    ShowMonth  ' ** Procedure: Below.

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub revmay_AfterUpdate()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "revmay_AfterUpdate"

1910    ShowMonth  ' ** Procedure: Below.

EXITP:
1920    Exit Sub

ERRH:
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Private Sub revjun_AfterUpdate()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "revjun_AfterUpdate"

2010    ShowMonth  ' ** Procedure: Below.

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Private Sub revjul_AfterUpdate()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "revjul_AfterUpdate"

2110    ShowMonth  ' ** Procedure: Below.

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub revaug_AfterUpdate()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "revaug_AfterUpdate"

2210    ShowMonth  ' ** Procedure: Below.

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Private Sub revsep_AfterUpdate()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "revsep_AfterUpdate"

2310    ShowMonth  ' ** Procedure: Below.

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Private Sub revoct_AfterUpdate()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "revoct_AfterUpdate"

2410    ShowMonth  ' ** Procedure: Below.

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub revnov_AfterUpdate()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "revnov_AfterUpdate"

2510    ShowMonth  ' ** Procedure: Below.

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub revdec_AfterUpdate()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "revdec_AfterUpdate"

2610    ShowMonth  ' ** Procedure: Below.

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub revdec_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "revdec_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
2770        With Me
2780          intRetVal = 0
2790          DoCmd.SelectObject acForm, .Parent.Name, False
2800          .Parent.cmdSelectAll.SetFocus
2810        End With
2820      End Select
2830    End If

        ' ** Shift keys.
2840    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2850      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2860        With Me
2870          intRetVal = 0
2880          .revnov.SetFocus
2890        End With
2900      End Select
2910    End If

EXITP:
2920    KeyCode = intRetVal
2930    Exit Sub

ERRH:
2940    intRetVal = 0
2950    Select Case ERR.Number
        Case Else
2960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2970    End Select
2980    Resume EXITP

End Sub

Public Sub ShowMonth()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowMonth"

        Dim ctl As Access.Control
        Dim blnFound As Boolean
        Dim lngX As Long

3010    With Me
3020      cmdSave_Click  ' ** Procedure: Above.
3030      blnFound = False
3040      For Each ctl In .Detail.Controls
3050        If ctl.ControlType = acCheckBox Then
3060          .Controls(ctl.Name & "_lbl2").Visible = ctl.Value
3070          For lngX = 1& To lngMonths
3080            If arr_varMonth(M_CNAM, lngX) = ctl.Name Then
3090              If ctl.Value <> arr_varMonth(M_VAL, lngX) Then
3100                blnFound = True
3110              End If
3120              Exit For
3130            End If
3140          Next
3150        End If
3160      Next
3170      .Parent.changed = blnFound
3180      .Parent.MonthVals_Chk  ' ** Form Procedure: frmAccountProfile_ReviewFreq.
3190    End With

EXITP:
3200    Set ctl = Nothing
3210    Exit Sub

ERRH:
3220    Select Case ERR.Number
        Case Else
3230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3240    End Select
3250    Resume EXITP

End Sub

Public Sub MonthVals_Set(arr_varMon As Variant)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "MonthVals_Set"

3310    arr_varMonth = arr_varMon
3320    lngMonths = UBound(arr_varMonth, 2)

EXITP:
3330    Exit Sub

ERRH:
3340    Select Case ERR.Number
        Case Else
3350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3360    End Select
3370    Resume EXITP

End Sub
