VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountProfile_StatementFreq_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountProfile_StatementFreq_Sub"

'VGC 03/22/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Number of Copies: B {numCopies}
' **   Cancel:           C {cmdCancel on frmAccountProfile_StatementFreq}
' **   Select None:      N {cmdSelectNone on frmAccountProfile_StatementFreq}
' **   OK:               O {cmdOK on frmAccountProfile_StatementFreq}
' **   Select All:       S {cmdSelectAll on frmAccountProfile_StatementFreq}

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
150       If strCallingForm = "frmAccountProfile_StatementFreq" Then

160         strCallersCallingForm = .Parent.OpenArgs
170         intPos01 = InStr(strCallersCallingForm, "~")
180         strAccountNo = Mid(strCallersCallingForm, (intPos01 + 1))
190         strCallersCallingForm = Left(strCallersCallingForm, (intPos01 - 1))

200         Select Case strCallersCallingForm
            Case "frmAccountProfile"
210           If .RecordSource <> "qryAccountFreq_Stmt_01_02" Then
220             .RecordSource = "qryAccountFreq_Stmt_01_02"
230           End If
240         Case "frmAccountProfile_Add"
250           .RecordSource = "qryAccountFreq_Stmt_02"
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

Private Sub smtjan_AfterUpdate()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "smtjan_AfterUpdate"

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

Private Sub smtjan_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "smtjan_KeyDown"

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
1390          .smtfeb.SetFocus
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

Private Sub smtfeb_AfterUpdate()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "smtfeb_AfterUpdate"

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

Private Sub smtmar_AfterUpdate()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "smtmar_AfterUpdate"

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

Private Sub smtapr_AfterUpdate()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "smtapr_AfterUpdate"

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

Private Sub smtmay_AfterUpdate()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "smtmay_AfterUpdate"

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

Private Sub smtjun_AfterUpdate()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "smtjun_AfterUpdate"

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

Private Sub smtjul_AfterUpdate()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "smtjul_AfterUpdate"

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

Private Sub smtaug_AfterUpdate()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "smtaug_AfterUpdate"

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

Private Sub smtsep_AfterUpdate()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "smtsep_AfterUpdate"

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

Private Sub smtoct_AfterUpdate()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "smtoct_AfterUpdate"

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

Private Sub smtnov_AfterUpdate()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "smtnov_AfterUpdate"

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

Private Sub smtdec_AfterUpdate()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "smtdec_AfterUpdate"

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

Private Sub smtdec_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "smtdec_KeyDown"

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
2790          .numCopies.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2850        With Me
2860          intRetVal = 0
2870          .smtnov.SetFocus
2880        End With
2890      End Select
2900    End If

EXITP:
2910    KeyCode = intRetVal
2920    Exit Sub

ERRH:
2930    intRetVal = 0
2940    Select Case ERR.Number
        Case Else
2950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2960    End Select
2970    Resume EXITP

End Sub

Private Sub numCopies_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "numCopies_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3020    intShiftDown = (Shift And acShiftMask) > 0
3030    intAltDown = (Shift And acAltMask) > 0
3040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3070        With Me
3080          intRetVal = 0
3090          DoCmd.SelectObject acForm, .Parent.Name, False
3100          If .Parent.cmdOK.Enabled = True Then
3110            .Parent.cmdOK.SetFocus
3120          Else
3130            .Parent.cmdCancel.SetFocus
3140          End If
3150        End With
3160      End Select
3170    End If

        ' ** Shift keys.
3180    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3190      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3200        With Me
3210          intRetVal = 0
3220          .smtdec.SetFocus
3230        End With
3240      End Select
3250    End If

EXITP:
3260    KeyCode = intRetVal
3270    Exit Sub

ERRH:
3280    intRetVal = 0
3290    Select Case ERR.Number
        Case Else
3300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3310    End Select
3320    Resume EXITP

End Sub

Public Sub ShowMonth()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowMonth"

        Dim ctl As Access.Control
        Dim blnFound As Boolean
        Dim lngX As Long

3410    With Me
3420      cmdSave_Click  ' ** Procedure: Above.
3430      blnFound = False
3440      For Each ctl In .Detail.Controls
3450        If ctl.ControlType = acCheckBox Then
3460          .Controls(ctl.Name & "_lbl2").Visible = ctl.Value
3470          For lngX = 1& To lngMonths
3480            If arr_varMonth(M_CNAM, lngX) = ctl.Name Then
3490              If ctl.Value <> arr_varMonth(M_VAL, lngX) Then
3500                blnFound = True
3510              End If
3520              Exit For
3530            End If
3540          Next
3550        End If
3560      Next
3570      .Parent.changed = blnFound
3580      .Parent.MonthVals_Chk  ' ** Form Procedure: frmAccountProfile_StatementFreq.
3590    End With

EXITP:
3600    Set ctl = Nothing
3610    Exit Sub

ERRH:
3620    Select Case ERR.Number
        Case Else
3630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3640    End Select
3650    Resume EXITP

End Sub

Public Sub MonthVals_Set(arr_varMon As Variant)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "MonthVals_Set"

3710    arr_varMonth = arr_varMon
3720    lngMonths = UBound(arr_varMonth, 2)

EXITP:
3730    Exit Sub

ERRH:
3740    Select Case ERR.Number
        Case Else
3750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3760    End Select
3770    Resume EXITP

End Sub
