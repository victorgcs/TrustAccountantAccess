VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountSearch"

'VGC 10/19/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Cancel:           C {cmdCancel}
' **   Name:             M {opgAccountSource_optName}
' **   OK:               O {cmdOK}

' ** I got a little carried away with this form.
' ** It's probably way more code than necessary.

' ** Combo box column constants: cmbAccounts.
Private Const CBX_ACCT_NUM  As Integer = 0  ' ** accountno
'Private Const CBX_ACCT_DISP As Integer = 1  ' ** Desc: accountno/shortname, qryAccountNoDropDown_01; shortname, qryAccountNoDropDown_02.

' ** List box column constants: lbxAccounts.
Private Const LBX_ACCT_NUM  As Integer = 0  ' ** accountno
'Private Const LBX_ACCT_DISP As Integer = 1  ' ** Desc: accountno/shortname, qryAccountNoDropDown_01; shortname, qryAccountNoDropDown_02.

Private strCallingForm As String
Private blnExitNow As Boolean, blnCmbAcctKeyReturn As Boolean
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = .OpenArgs

150           .PictureData = .Form_AccountSearch_img

160           blnExitNow = False: blnCmbAcctKeyReturn = False

170           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

180           opgAccountSource_AfterUpdate  ' ** Procedure: Below.

190         Else
200           Cancel = -1
210         End If
220       Else
230         Cancel = -1
240       End If
250     End With

260     If Cancel = -1 Then
270       Beep
280       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
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

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     opgAccountSource_AfterUpdate  ' ** Procedure: Below.

EXITP:
420     Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

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

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         intRetVal = 0
580         gstrAccountNo = vbNullString
590         blnExitNow = True
600         cmdCancel_Click  ' ** Procedure: Below.
610       End Select
620     End If

EXITP:
630     KeyCode = intRetVal
640     Exit Sub

ERRH:
650     intRetVal = 0
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

End Sub

Private Sub Form_Timer()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

710     With Me
720       .TimerInterval = 0&
730       If blnExitNow = True Then
740         DoCmd.Close acForm, THIS_NAME
750       End If
760     End With

EXITP:
770     Exit Sub

ERRH:
780     Select Case ERR.Number
        Case Else
790       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
800     End Select
810     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

910     With Me

920       Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

          ' ** Check if we've got a good account number.
930       If gstrAccountNo <> vbNullString Then  ' ** Might be from cmdCancel.
940         If IsNull(.cmbAccounts) = False Then
950           If .cmbAccounts <> vbNullString Then
960             gstrAccountNo = .cmbAccounts.Column(CBX_ACCT_NUM)
970           End If
980         End If
990       End If

1000      Select Case strCallingForm
          Case "frmMenu_Account"
1010        DoCmd.SelectObject acForm, strCallingForm, False
1020      End Select

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

Private Sub cmdOK_Click()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

1110    With Me
1120      If IsNull(.cmbAccounts) = False Then
1130        If .cmbAccounts <> vbNullString Then
1150          gstrAccountNo = .cmbAccounts
1160          blnExitNow = True
1170          .TimerInterval = 100&
1180        Else
1190          Beep
1200          MsgBox "Choose an Account, or click Cancel to exit.", vbInformation + vbOKOnly, "Invalid Entry"
1210        End If
1220      Else
1230        Beep
1240        MsgBox "Choose an Account, or click Cancel to exit.", vbInformation + vbOKOnly, "Invalid Entry"
1250      End If
1260    End With

EXITP:
1270    Exit Sub

ERRH:
1280    Select Case ERR.Number
        Case Else
1290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1300    End Select
1310    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

1410    gstrAccountNo = vbNullString
1420    DoCmd.Close acForm, THIS_NAME

EXITP:
1430    Exit Sub

ERRH:
1440    Select Case ERR.Number
        Case Else
1450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1460    End Select
1470    Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo1 As String, strAccountNo2 As String

1510    strAccountNo1 = vbNullString: strAccountNo2 = vbNullString

1520    With Me
1530      If IsNull(.cmbAccounts) = False Then
1540        If Len(.cmbAccounts.Column(CBX_ACCT_NUM)) > 0 Then
1550          strAccountNo1 = .cmbAccounts.Column(CBX_ACCT_NUM)
1560        End If
1570      End If
1580      If IsNull(.lbxAccounts) = False Then
1590        If Len(.lbxAccounts.Column(LBX_ACCT_NUM)) > 0 Then
1600          strAccountNo2 = .lbxAccounts.Column(LBX_ACCT_NUM)
1610        End If
1620      End If
1630      Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
1640        .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
1650        .lbxAccounts.RowSource = "qryAccountNoDropDown_03"
1660        .opgAccountSource_optNumber_lbl.FontBold = True
1670        .opgAccountSource_optName_lbl.FontBold = False
1680      Case .opgAccountSource_optName.OptionValue
1690        .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
1700        .lbxAccounts.RowSource = "qryAccountNoDropDown_04"
1710        .opgAccountSource_optNumber_lbl.FontBold = False
1720        .opgAccountSource_optName_lbl.FontBold = True
1730      End Select
1740      DoEvents
1750      If strAccountNo1 <> vbNullString Then
1760        .cmbAccounts = strAccountNo1
1770      End If
1780      If strAccountNo2 <> vbNullString Then
1790        .lbxAccounts = strAccountNo2
1800      End If
1810    End With

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

1910    With Me

1920      gstrAccountNo = vbNullString
1930      If IsNull(.cmbAccounts) = False Then
1940        If .cmbAccounts <> vbNullString Then
1950          gstrAccountNo = .cmbAccounts.Column(CBX_ACCT_NUM)
1960          .lbxAccounts = gstrAccountNo
1970        End If
1980      End If

1990      If blnCmbAcctKeyReturn = True And gstrAccountNo <> vbNullString Then
            ' ** Access seems to get confused, and spews an error if I try to close the form from here.
            ' ** Error 2585: This action can't be carried out while processing a form or report event.
2000        blnCmbAcctKeyReturn = False
2010        blnExitNow = True
2020        .TimerInterval = 100&
2030      End If

2040    End With

EXITP:
2050    Exit Sub

ERRH:
2060    Select Case ERR.Number
        Case 2585  ' ** This action can't be carried out while processing a form or report event.
          ' ** Ignore.
2070    Case Else
2080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2090    End Select
2100    Resume EXITP

End Sub

Private Sub cmbAccounts_NotInList(NewData As String, Response As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_NotInList"

        Dim intX As Integer

2210    With Me
2220      blnCmbAcctKeyReturn = False
2230      .lbxAccounts = Null
2240      For intX = 0 To (.lbxAccounts.ListCount - 1)
2250        .lbxAccounts.Selected(intX) = False
2260      Next
2270      Response = acDataErrDisplay
          ' ** acDataErrDisplay
          ' ** acDataErrContinue
          ' ** acDataErrAdded
2280      .Repaint  ' ** Sometimes the List Box looks offset, like not quite repainted.
2290    End With    ' ** I'm not sure what causes it, or how to fix it.

EXITP:
2300    Exit Sub

ERRH:
2310    Select Case ERR.Number
        Case Else
2320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2330    End Select
2340    Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)
' ** Track keystrokes and mirror in listbox
' ** lbxAccounts, then return unchanged KeyCode.
' ** I'm not sure whether the KeyCode even matters here now,
' ** but it's working, so I'll let it be.

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2420    intShiftDown = (Shift And acShiftMask) > 0
2430    intAltDown = (Shift And acAltMask) > 0
2440    intCtrlDown = (Shift And acCtrlMask) > 0

2450    With Me

2460      blnCmbAcctKeyReturn = False

2470      If .cmbAccounts.ListIndex >= 0 Then
            ' ** Filter-out non-printable characters.
2480        If SkipKey(intRetVal) = False Then  ' ** Module Function: modStringFuncs.
              ' ** Plain keys.
2490          If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2500            .lbxAccounts.Selected(.cmbAccounts.ListIndex) = True
2510          End If
              ' ** Shift keys (for initial cap or caps lock).
2520          If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2530            .lbxAccounts.Selected(.cmbAccounts.ListIndex) = True
2540          End If
2550        End If
2560      End If

          ' ** Plain keys.
2570      If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2580        Select Case intRetVal
            Case vbKeyTab
2590          With Me
2600            intRetVal = 0
2610            .lbxAccounts.SetFocus
2620          End With
2630        Case vbKeyReturn
              ' ** This variable may no longer be necessary!
2640          blnCmbAcctKeyReturn = True
2650        End Select
2660      End If

2670    End With

EXITP:
2680    KeyCode = intRetVal
2690    Exit Sub

ERRH:
2700    Select Case ERR.Number
        Case Else
2710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2720    End Select
2730    Resume EXITP

End Sub

Private Sub cmbAccounts_Exit(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Exit"

2810    With Me
2820      If blnExitNow = True Then
2830        If blnCmbAcctKeyReturn = True Then
2840          If .cmbAccounts.text = vbNullString Then
2850            gstrAccountNo = vbNullString
2860            blnCmbAcctKeyReturn = False
2870            blnExitNow = False
2880          End If
2890        End If
2900      Else
2910        If IsNull(.cmbAccounts) = False Then
2920          If .cmbAccounts <> vbNullString Then
2930            If .cmbAccounts.text <> vbNullString Then
2940              gstrAccountNo = .cmbAccounts
2950              blnExitNow = True
2960            End If
2970          End If
2980        End If
2990      End If
3000    End With

EXITP:
3010    Exit Sub

ERRH:
3020    Select Case ERR.Number
        Case Else
3030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3040    End Select
3050    Resume EXITP

End Sub

Private Sub lbxAccounts_AfterUpdate()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxAccounts_AfterUpdate"

3110    With Me
3120      gstrAccountNo = vbNullString
3130      If IsNull(.lbxAccounts) = False Then
3140        If .lbxAccounts <> vbNullString Then
3150          gstrAccountNo = .lbxAccounts.Column(LBX_ACCT_NUM)
3160          .cmbAccounts = gstrAccountNo
3170        End If
3180      End If
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

Private Sub lbxAccounts_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxAccounts_DblClick"

3310    With Me
3320      gstrAccountNo = vbNullString: blnExitNow = False
3330      If IsNull(.lbxAccounts) = False Then
3340        If .lbxAccounts <> vbNullString Then
3350          gstrAccountNo = .lbxAccounts.Column(LBX_ACCT_NUM)
3360          blnExitNow = True
3370        End If
3380      End If
3390      If blnExitNow = True Then
3400        .TimerInterval = 100&
3410      End If
3420    End With

EXITP:
3430    Exit Sub

ERRH:
3440    Select Case ERR.Number
        Case Else
3450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3460    End Select
3470    Resume EXITP

End Sub

Private Sub lbxAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3520    intShiftDown = (Shift And acShiftMask) > 0
3530    intAltDown = (Shift And acAltMask) > 0
3540    intCtrlDown = (Shift And acCtrlMask) > 0

3550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3560      Select Case intRetVal
          Case vbKeyTab
3570        With Me
3580          intRetVal = 0
3590          If .cmdOK.Enabled = True Then
3600            .cmdOK.SetFocus
3610          Else
3620            .cmdCancel.SetFocus
3630          End If
3640        End With
3650      Case vbKeyReturn
3660        With Me
3670          gstrAccountNo = vbNullString: blnExitNow = False
3680          If IsNull(.lbxAccounts) = False Then
3690            If .lbxAccounts <> vbNullString Then
3700              gstrAccountNo = .lbxAccounts.Column(LBX_ACCT_NUM)
3710              blnExitNow = True
3720            End If
3730          End If
3740          If blnExitNow = True Then
3750            .TimerInterval = 100&
3760          End If
3770        End With
3780      End Select
3790    End If

EXITP:
3800    KeyCode = intRetVal
3810    Exit Sub

ERRH:
3820    Select Case ERR.Number
        Case Else
3830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3840    End Select
3850    Resume EXITP

End Sub
