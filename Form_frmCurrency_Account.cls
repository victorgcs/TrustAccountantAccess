VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_Account"

'VGC 08/31/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   All:              A {opgFilter_optAll}
' **   Close:            C {cmdClose}
' **   Default Suppress: D {chkDefaultSuppress}
' **   Suppress On/Off:  F {curracct_suppress on frmCurrency_Account_Sub}
' **   Option Suppress:  H {chkOptionSuppress}
' **   Non-Suppressed:   N {opgFilter_optNonSuppressed}
' **   Regenerate:       R {cmdRegenerate}
' **   Suppressed:       S {opgFilter_optSuppressed}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmCurrency_Account_Sub}

Private lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

160           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
170           DoEvents

180           opgFilter_AfterUpdate  ' ** Procedure: Below.
190           DoEvents
200           chkDefaultSuppress_AfterUpdate  ' ** Procedure: Below.
210           chkOptionSuppress_AfterUpdate  ' ** Procedure: Below.

220           .NoForEx_lbl.Visible = False

230           GetCounts  ' ** Procedure: Below.

240           lngRecsCur = .frmCurrency_Account_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Account_Sub.
250           .frmCurrency_Account_Sub.SetFocus
260           If lngRecsCur > 0 Then
270             .frmCurrency_Account_Sub.Form.FocusHolder.SetFocus
280           Else
290             .frmCurrency_Account_Sub.Form.FocusHolder2.SetFocus
300           End If

310         Else
320           Cancel = -1
330         End If
340       Else
350         Cancel = -1
360       End If
370     End With

380     DoCmd.Hourglass False

390     If Cancel = -1 Then
400       Beep
410       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
420     End If

EXITP:
430     Exit Sub

ERRH:
440     DoCmd.Hourglass False
450     Select Case ERR.Number
        Case Else
460       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
470     End Select
480     Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Default Suppress: D {chkDefaultSuppress}
        ' **   Foreign Check:    F {curracct_suppress on frmCurrency_Account_Sub}
        ' **   Option Suppress:  H {chkOptionSuppress}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmCurrency_Account_Sub}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         intRetVal = 0
580         cmdClose_Click  ' ** Procedure: Below.
590       End Select
600     End If

        ' ** Alt keys.
610     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
620       Select Case intRetVal
          Case vbKeyD
630         With Me
640           intRetVal = 0
650           .chkDefaultSuppress.SetFocus
660           .chkDefaultSuppress = (Not .chkDefaultSuppress)
670           chkDefaultSuppress_AfterUpdate  ' ** Procedure: Below.
680         End With
690       Case vbKeyF
700         With Me
710           intRetVal = 0
720           .frmCurrency_Account_Sub.SetFocus
730           .frmCurrency_Account_Sub.Form.curracct_sort.SetFocus  ' ** Set focus only, dont't flip.
740         End With
750       Case vbKeyH
760         With Me
770           intRetVal = 0
780           .chkOptionSuppress.SetFocus
790           .chkOptionSuppress = (Not .chkOptionSuppress)
800           chkOptionSuppress_AfterUpdate  ' ** Procedure: Below.
810         End With
820       Case vbKeyX
830         intRetVal = 0
840         cmdClose_Click  ' ** Procedure: Below.
850       End Select
860     End If

        ' ** Ctrl keys.
870     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
880       Select Case intRetVal
          Case vbKeyS
890         With Me
900           intRetVal = 0
910           .frmCurrency_Account_Sub.SetFocus
920           .frmCurrency_Account_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmCurrency_Account_Sub.
930         End With
940       End Select
950     End If

        ' ** Ctrl-Shift keys.
960     If intCtrlDown And (Not intAltDown) And intShiftDown Then
970       Select Case intRetVal
          Case vbKeyF
980         With Me
990           intRetVal = 0
1000          .FocusHolder.SetFocus
1010        End With
1020      End Select
1030    End If

EXITP:
1040    KeyCode = intRetVal
1050    Exit Sub

ERRH:
1060    intRetVal = 0
1070    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1080    Case Else
1090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1100    End Select
1110    Resume EXITP

End Sub

Private Sub Form_Timer()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1210    With Me
1220      .TimerInterval = 0&
1230      If gblnSetFocus = True Then
1240        gblnSetFocus = False
1250        strFormName = vbNullString
1260  On Error Resume Next
1270        Do Until strFormName = THIS_NAME
1280          DoCmd.SelectObject acForm, THIS_NAME, False
1290          strFormName = Screen.ActiveForm.Name
1300        Loop
1310  On Error GoTo ERRH
1320      End If
1330      DoCmd.Hourglass False
1340    End With

EXITP:
1350    Exit Sub

ERRH:
1360    DoCmd.Hourglass False
1370    Select Case ERR.Number
        Case Else
1380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1390    End Select
1400    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1510    With Me

1520      DoCmd.Hourglass True
1530      DoEvents

1540      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1550      Select Case strCallingForm
          Case "frmMenu_ForeignExchange"
1560        gstrReturningForm = THIS_NAME
1570        gblnSetFocus = True
1580        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Maintenance"
1590      End Select

1600    End With

EXITP:
1610    Exit Sub

ERRH:
1620    DoCmd.Hourglass False
1630    Select Case ERR.Number
        Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Public Sub cmdClose_Click()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

1710    DoCmd.Close acForm, THIS_NAME

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
1890          .frmCurrency_Account_Sub.SetFocus
1900        End With
1910      End Select
1920    End If

        ' ** Shift keys.
1930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1940      Select Case intRetVal
          Case vbKeyTab
1950        With Me
1960          intRetVal = 0
1970          .cmdRegenerate.SetFocus
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

Private Sub cmdRegenerate_Click()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRegenerate_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim msgResponse As VbMsgBoxResult

2110    With Me

2120      Beep
2130      msgResponse = MsgBox("This will regenerate the Account/Currency list above." & vbCrLf & vbCrLf & _
            "Proceed?", vbQuestion + vbYesNo, "Regenerate Account/Currency List")

2140      If msgResponse = vbYes Then

2150        DoCmd.Hourglass True
2160        DoEvents

2170        Set dbs = CurrentDb
            ' ** Empty tblCurrency_Account.
2180        Set qdf = dbs.QueryDefs("qryCurrency_Account_03")
2190        qdf.Execute
2200        Set qdf = Nothing
2210        dbs.Close
2220        Set dbs = Nothing
2230        DoEvents

2240        gblnHasForEx = HasForEx_All  ' ** Module Function: modCurrencyFuncs.
2250        HasForEx_Load  ' ** Module Procedure: modCurrencyFuncs.
2260        DoEvents

2270        .frmCurrency_Account_Sub.Form.Requery
2280        DoEvents

2290        Beep

2300      End If

2310      DoCmd.Hourglass False

2320    End With

EXITP:
2330    Set qdf = Nothing
2340    Set dbs = Nothing
2350    Exit Sub

ERRH:
2360    DoCmd.Hourglass False
2370    Select Case ERR.Number
        Case Else
2380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2390    End Select
2400    Resume EXITP

End Sub

Private Sub cmdRegenerate_KeyDown(KeyCode As Integer, Shift As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRegenerate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2520    intShiftDown = (Shift And acShiftMask) > 0
2530    intAltDown = (Shift And acAltMask) > 0
2540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2560      Select Case intRetVal
          Case vbKeyTab
2570        With Me
2580          intRetVal = 0
2590          .cmdClose.SetFocus
2600        End With
2610      End Select
2620    End If

        ' ** Shift keys.
2630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2640      Select Case intRetVal
          Case vbKeyTab
2650        With Me
2660          intRetVal = 0
2670          .chkOptionSuppress.SetFocus
2680        End With
2690      End Select
2700    End If

EXITP:
2710    KeyCode = intRetVal
2720    Exit Sub

ERRH:
2730    intRetVal = 0
2740    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2750    Case Else
2760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2770    End Select
2780    Resume EXITP

End Sub

Private Sub opgFilter_AfterUpdate()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgFilter_AfterUpdate"

        Dim strFilter As String

2810    With Me
2820      Select Case .opgFilter
          Case .opgFilter_optAll.OptionValue
2830        .opgFilter_optAll_lbl.FontBold = True
2840        .opgFilter_optSuppressed_lbl.FontBold = False
2850        .opgFilter_optNonSuppressed_lbl.FontBold = False
2860        strFilter = vbNullString
2870        .frmCurrency_Account_Sub.Form.Filter = strFilter
2880        .frmCurrency_Account_Sub.Form.FilterOn = False
2890      Case .opgFilter_optSuppressed.OptionValue
2900        .opgFilter_optAll_lbl.FontBold = False
2910        .opgFilter_optSuppressed_lbl.FontBold = True
2920        .opgFilter_optNonSuppressed_lbl.FontBold = False
2930        strFilter = "[curracct_suppress] = True"
2940        .frmCurrency_Account_Sub.Form.Filter = strFilter
2950        .frmCurrency_Account_Sub.Form.FilterOn = True
2960      Case .opgFilter_optNonSuppressed.OptionValue
2970        .opgFilter_optAll_lbl.FontBold = False
2980        .opgFilter_optSuppressed_lbl.FontBold = False
2990        .opgFilter_optNonSuppressed_lbl.FontBold = True
3000        strFilter = "[curracct_suppress] = False"
3010        .frmCurrency_Account_Sub.Form.Filter = strFilter
3020        .frmCurrency_Account_Sub.Form.FilterOn = True
3030      End Select
3040    End With

EXITP:
3050    Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub opgFilter_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgFilter_optAll_KeyDown"

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
          Case vbKeyTab
3170        With Me
3180          intRetVal = 0
3190          .chkDefaultSuppress.SetFocus
3200        End With
3210      End Select
3220    End If

        ' ** Shift keys.
3230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3240      Select Case intRetVal
          Case vbKeyTab
3250        With Me
3260          intRetVal = 0
3270          lngRecsCur = .frmCurrency_Account_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Account_Sub.
3280          .frmCurrency_Account_Sub.SetFocus
3290          If lngRecsCur = 0 Then
3300            .frmCurrency_Account_Sub.Form.FocusHolder2.SetFocus
3310          Else
3320            .frmCurrency_Account_Sub.Form.FocusHolder.SetFocus
3330          End If
3340        End With
3350      End Select
3360    End If

EXITP:
3370    KeyCode = intRetVal
3380    Exit Sub

ERRH:
3390    intRetVal = 0
3400    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3410    Case Else
3420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3430    End Select
3440    Resume EXITP

End Sub

Private Sub opgFilter_optSuppressed_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgFilter_optSuppressed_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3520    intShiftDown = (Shift And acShiftMask) > 0
3530    intAltDown = (Shift And acAltMask) > 0
3540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3560      Select Case intRetVal
          Case vbKeyTab
3570        With Me
3580          intRetVal = 0
3590          .chkDefaultSuppress.SetFocus
3600        End With
3610      End Select
3620    End If

        ' ** Shift keys.
3630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3640      Select Case intRetVal
          Case vbKeyTab
3650        With Me
3660          intRetVal = 0
3670          lngRecsCur = .frmCurrency_Account_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Account_Sub.
3680          .frmCurrency_Account_Sub.SetFocus
3690          If lngRecsCur = 0 Then
3700            .frmCurrency_Account_Sub.Form.FocusHolder2.SetFocus
3710          Else
3720            .frmCurrency_Account_Sub.Form.FocusHolder.SetFocus
3730          End If
3740        End With
3750      End Select
3760    End If

EXITP:
3770    KeyCode = intRetVal
3780    Exit Sub

ERRH:
3790    intRetVal = 0
3800    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3810    Case Else
3820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3830    End Select
3840    Resume EXITP

End Sub

Private Sub opgFilter_optNonSuppressed_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgFilter_optNonSuppressed_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3920    intShiftDown = (Shift And acShiftMask) > 0
3930    intAltDown = (Shift And acAltMask) > 0
3940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3960      Select Case intRetVal
          Case vbKeyTab
3970        With Me
3980          intRetVal = 0
3990          .chkDefaultSuppress.SetFocus
4000        End With
4010      End Select
4020    End If

        ' ** Shift keys.
4030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4040      Select Case intRetVal
          Case vbKeyTab
4050        With Me
4060          intRetVal = 0
4070          lngRecsCur = .frmCurrency_Account_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Account_Sub.
4080          .frmCurrency_Account_Sub.SetFocus
4090          If lngRecsCur = 0 Then
4100            .frmCurrency_Account_Sub.Form.FocusHolder2.SetFocus
4110          Else
4120            .frmCurrency_Account_Sub.Form.FocusHolder.SetFocus
4130          End If
4140        End With
4150      End Select
4160    End If

EXITP:
4170    KeyCode = intRetVal
4180    Exit Sub

ERRH:
4190    intRetVal = 0
4200    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4210    Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Sub

Public Sub chkDefaultSuppress_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkDefaultSuppress_AfterUpdate"

4310    With Me
4320      Select Case .chkDefaultSuppress
          Case True
4330        .chkDefaultSuppress_lbl.FontBold = True
4340        .chkDefaultSuppress_lbl2.FontBold = True
4350        .chkDefaultSuppress_box.Visible = True
4360        DoEvents
4370        If .chkOptionSuppress = True Then
4380          .chkOptionSuppress = False
4390          chkOptionSuppress_AfterUpdate  ' ** Procedure: Below.
4400        End If
4410      Case False
4420        .chkDefaultSuppress_lbl.FontBold = False
4430        .chkDefaultSuppress_lbl2.FontBold = False
4440        .chkDefaultSuppress_box.Visible = False
4450        If .chkOptionSuppress = False Then
4460          .chkOptionSuppress = True
4470          chkOptionSuppress_AfterUpdate  ' ** Procedure: Below.
4480        End If
4490      End Select
4500    End With

EXITP:
4510    Exit Sub

ERRH:
4520    Select Case ERR.Number
        Case Else
4530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4540    End Select
4550    Resume EXITP

End Sub

Private Sub chkDefaultSuppress_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkDefaultSuppress_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4620    intShiftDown = (Shift And acShiftMask) > 0
4630    intAltDown = (Shift And acAltMask) > 0
4640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4660      Select Case intRetVal
          Case vbKeyTab
4670        With Me
4680          intRetVal = 0
4690          .chkOptionSuppress.SetFocus
4700        End With
4710      End Select
4720    End If

        ' ** Shift keys.
4730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4740      Select Case intRetVal
          Case vbKeyTab
4750        With Me
4760          intRetVal = 0
4770          .opgFilter.SetFocus
4780        End With
4790      End Select
4800    End If

EXITP:
4810    KeyCode = intRetVal
4820    Exit Sub

ERRH:
4830    intRetVal = 0
4840    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4850    Case Else
4860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4870    End Select
4880    Resume EXITP

End Sub

Public Sub chkOptionSuppress_AfterUpdate()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkOptionSuppress_AfterUpdate"

4910    With Me
4920      Select Case .chkOptionSuppress
          Case True
4930        .chkOptionSuppress_lbl.FontBold = True
4940        .chkOptionSuppress_lbl2.FontBold = True
4950        .chkOptionSuppress_box.Visible = True
4960        DoEvents
4970        If .chkDefaultSuppress = True Then
4980          .chkDefaultSuppress = False
4990          chkDefaultSuppress_AfterUpdate  ' ** Procedure: Above.
5000        End If
5010      Case False
5020        .chkOptionSuppress_lbl.FontBold = False
5030        .chkOptionSuppress_lbl2.FontBold = False
5040        .chkOptionSuppress_box.Visible = False
5050        If .chkDefaultSuppress = False Then
5060          .chkDefaultSuppress = True
5070          chkDefaultSuppress_AfterUpdate  ' ** Procedure: Above.
5080        End If
5090      End Select
5100    End With

EXITP:
5110    Exit Sub

ERRH:
5120    Select Case ERR.Number
        Case Else
5130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5140    End Select
5150    Resume EXITP

End Sub

Private Sub chkOptionSuppress_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkOptionSuppress_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5220    intShiftDown = (Shift And acShiftMask) > 0
5230    intAltDown = (Shift And acAltMask) > 0
5240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5260      Select Case intRetVal
          Case vbKeyTab
5270        With Me
5280          intRetVal = 0
5290          .cmdRegenerate.SetFocus
5300        End With
5310      End Select
5320    End If

        ' ** Shift keys.
5330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5340      Select Case intRetVal
          Case vbKeyTab
5350        With Me
5360          intRetVal = 0
5370          .chkDefaultSuppress.SetFocus
5380        End With
5390      End Select
5400    End If

EXITP:
5410    KeyCode = intRetVal
5420    Exit Sub

ERRH:
5430    intRetVal = 0
5440    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5450    Case Else
5460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5470    End Select
5480    Resume EXITP

End Sub

Private Sub GetCounts()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "GetCounts"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCntJNo As Long, lngCntAA As Long
        Dim blnRegen As Boolean
        Dim varTmp00 As Variant

5510    With Me

5520      blnRegen = False
5530      lngCntJNo = 0&: lngCntAA = 0&

5540      varTmp00 = DCount("*", "tblCurrency_Account")
5550      Select Case IsNull(varTmp00)
          Case True
5560        blnRegen = True
5570      Case False
5580        If varTmp00 = 0 Then
5590          blnRegen = True
5600        End If
5610      End Select
5620      DoEvents

5630      If blnRegen = True Then
5640        HasForEx_Load  ' ** Module Procedure: modCurrencyFuncs.
5650      End If
5660      DoEvents

5670      Set dbs = CurrentDb
5680      With dbs
            ' ** qryCurrency_11_03 (Account, linked to qryCurrency_11_02 (qryCurrency_11_01 (Union of
            ' ** qryCurrency_09_04 (qryCurrency_09_03 (Union of qryCurrency_09_01 (Ledger, grouped by
            ' ** accountno, just curr_id <> 150, with cnt_jno), qryCurrency_09_02 (LedgerArchive,
            ' ** grouped by accountno, just curr_id <> 150, with cnt_jno)), grouped and summed, by
            ' ** accountno, with cnt_jno), qryCurrency_10_01 (ActiveAssets, grouped by accountno,
            ' ** just curr_id <> 150, with cnt_aa)), grouped, with Max(cnt_jno), Max(cnt_aa)),
            ' ** with acct_sort), grouped and summed.
5690        Set qdf = .QueryDefs("qryCurrency_Account_02")
5700        Set rst = qdf.OpenRecordset
5710        With rst
5720          .MoveFirst
5730          lngCntJNo = ![cnt_jno]
5740          lngCntAA = ![Cnt_AA]
5750          .Close
5760        End With
5770        .Close
5780      End With
5790      Set dbs = Nothing

5800      If lngCntJNo = 0& And lngCntAA = 0& Then
5810        .NoForEx_lbl.Visible = True
5820      ElseIf lngCntJNo > 0& And lngCntAA = 0& Then
5830        .NoForExAssets_lbl.Visible = True
5840      ElseIf lngCntJNo = 0& And lngCntAA > 0& Then
            ' ** This shouldn't be able to happen!
5850      Else
            ' ** Both present, no message.
5860      End If

5870    End With

EXITP:
5880    Set rst = Nothing
5890    Set qdf = Nothing
5900    Set dbs = Nothing
5910    Exit Sub

ERRH:
5920    DoCmd.Hourglass False
5930    Select Case ERR.Number
        Case Else
5940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5950    End Select
5960    Resume EXITP

End Sub
