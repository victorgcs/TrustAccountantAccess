VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_AccountProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_AccountProfile"

'VGC 07/14/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Excel:            E {cmdExcel}
' **   All:              L {opgAccountNumber_optAll}
' **   Name:             M {opgAccountSource_optName}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Specified:        U {opgAccountNumber_optSpecified}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
160           .cmdExcel.Enabled = False
      #Else
170           .cmdExcel.Enabled = True
      #End If

              ' ** EVENT CHECK: chkRememberMe!
              ' ** NONE!

180           opgAccountSource_AfterUpdate  ' ** Procedure: Below.
190           chkRememberMe_AfterUpdate  ' ** Procedure: Below.
200           chkPageOf_AfterUpdate  ' ** Procedure: Below.
210           opgAccountNumber_AfterUpdate  ' ** Procedure: Below.

220         Else
230           Cancel = -1
240         End If
250       Else
260         Cancel = -1
270       End If
280     End With

290     If Cancel = -1 Then
300       Beep
310       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
320     End If

330     DoCmd.Hourglass False

EXITP:
340     Exit Sub

ERRH:
350     DoCmd.Hourglass False
360     Select Case ERR.Number
        Case Else
370       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
380     End Select
390     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.
430     End With

EXITP:
440     Exit Sub

ERRH:
450     DoCmd.Hourglass False
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

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
        ' **   Number:           B {opgAccountSource_optNumber}
        ' **   Name:             M {opgAccountSource_optName}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

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
          Case vbKeyB
630         With Me
640           intRetVal = 0
650           If .opgAccountSource.Enabled = True Then
660             .opgAccountSource = .opgAccountSource_optNumber.OptionValue
670             .opgAccountSource.SetFocus
680             opgAccountSource_AfterUpdate  ' ** Procedure: Below.
690           Else
700             Beep
710           End If
720         End With
730       Case vbKeyM
740         With Me
750           intRetVal = 0
760           If .opgAccountSource.Enabled = True Then
770             .opgAccountSource = .opgAccountSource_optName.OptionValue
780             .opgAccountSource.SetFocus
790             opgAccountSource_AfterUpdate  ' ** Procedure: Below.
800           Else
810             Beep
820           End If
830         End With
840       Case vbKeyX
850         intRetVal = 0
860         cmdClose_Click  ' ** Procedure: Below.
870       End Select
880     End If

        ' ** Ctrl keys.
890     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
900       Select Case intRetVal
          Case vbKeyG
910         With Me
920           intRetVal = 0
930           If .chkPageOf.Enabled = True Then
940             .chkPageOf.SetFocus
950             .chkPageOf = (Not .chkPageOf)
960             chkPageOf_AfterUpdate  ' ** Procedure: Below.
970           End If
980         End With
990       End Select
1000    End If

        ' ** Ctrl-Shift keys.
1010    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1020      Select Case intRetVal
          Case vbKeyF
1030        With Me
1040          intRetVal = 0
1050          .FocusHolder.SetFocus
1060        End With
1070      End Select
1080    End If

EXITP:
1090    KeyCode = intRetVal
1100    Exit Sub

ERRH:
1110    intRetVal = 0
1120    Select Case ERR.Number
        Case Else
1130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1140    End Select
1150    Resume EXITP

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
1330      If gblnGoToReport = True Then
1340        If .cmdPrint.Enabled = True Then
1350          .cmdPrint.SetFocus
1360        Else
1370          .cmdClose.SetFocus
1380        End If
1390        gblnGoToReport = False
1400        Beep
1410        DoCmd.Hourglass False
1420        DoEvents
1430      Else
1440        DoCmd.Hourglass False
1450      End If
1460    End With

EXITP:
1470    Exit Sub

ERRH:
1480    DoCmd.Hourglass False
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1610    DoCmd.Hourglass True
1620    DoEvents

1630    With Me
1640      If .chkRememberMe = False Then
1650        .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
1660      End If
1670    End With

1680    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
1690    Do While Reports.Count > 0
1700      DoCmd.Close acReport, Reports(0).Name
1710    Loop

1720    gstrFormQuerySpec = vbNullString
1730    gblnSetFocus = True
1740    gstrReturningForm = THIS_NAME
1750    Select Case strCallingForm
        Case "frmMenu_Report"
1760      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
1770    End Select

EXITP:
1780    Exit Sub

ERRH:
1790    Select Case ERR.Number
        Case Else
1800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1810    End Select
1820    Resume EXITP

End Sub

Private Sub cmdClose_Click()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

1910    DoCmd.Hourglass True
1920    DoEvents

1930    DoCmd.Close acForm, THIS_NAME

EXITP:
1940    Exit Sub

ERRH:
1950    Select Case ERR.Number
        Case Else
1960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1970    End Select
1980    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2020    intShiftDown = (Shift And acShiftMask) > 0
2030    intAltDown = (Shift And acAltMask) > 0
2040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2060      Select Case intRetVal
          Case vbKeyTab
2070        With Me
2080          intRetVal = 0
2090          .opgAccountNumber.SetFocus
2100        End With
2110      End Select
2120    End If

        ' ** Shift keys.
2130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2140      Select Case intRetVal
          Case vbKeyTab
2150        With Me
2160          intRetVal = 0
2170          If .cmdExcel.Enabled = True Then
2180            .cmdExcel.SetFocus
2190          Else
2200            .cmdWord.SetFocus
2210          End If
2220        End With
2230      End Select
2240    End If

EXITP:
2250    KeyCode = intRetVal
2260    Exit Sub

ERRH:
2270    intRetVal = 0
2280    Select Case ERR.Number
        Case Else
2290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2300    End Select
2310    Resume EXITP

End Sub

Private Sub opgAccountNumber_AfterUpdate()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_AfterUpdate"

2410    With Me
2420      Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
2430        .opgAccountNumber_optSpecified_lbl.FontBold = True
2440        .opgAccountNumber_optAll_lbl.FontBold = False
2450        .opgAccountNumber_optSpecified_lbl_box.Visible = True
2460        .opgAccountNumber_optAll_lbl_box.Visible = False
2470        .cmbAccounts.Enabled = True
2480        .cmbAccounts.BorderColor = CLR_LTBLU2
2490        .cmbAccounts.BackStyle = acBackStyleNormal
2500        .opgAccountSource.Enabled = True
2510        .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
2520        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
2530        .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
2540        .opgAccountSource_optName_lbl2_dim_hi.Visible = False
2550        .chkRememberMe.Enabled = True
2560        .chkRememberMe_lbl.Visible = True
2570        .chkRememberMe_lbl2_dim.Visible = False
2580        .chkRememberMe_lbl2_dim_hi.Visible = False
2590      Case .opgAccountNumber_optAll.OptionValue
2600        .opgAccountNumber_optSpecified_lbl.FontBold = False
2610        .opgAccountNumber_optAll_lbl.FontBold = True
2620        .opgAccountNumber_optSpecified_lbl_box.Visible = False
2630        .opgAccountNumber_optAll_lbl_box.Visible = True
2640        .cmbAccounts.Enabled = False
2650        .cmbAccounts.BorderColor = WIN_CLR_DISR
2660        .cmbAccounts.BackStyle = acBackStyleTransparent
2670        .opgAccountSource.Enabled = False
2680        .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
2690        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
2700        .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
2710        .opgAccountSource_optName_lbl2_dim_hi.Visible = True
2720        .chkRememberMe.Enabled = False
2730        .chkRememberMe_lbl.Visible = False
2740        .chkRememberMe_lbl2_dim.Visible = True
2750        .chkRememberMe_lbl2_dim_hi.Visible = True
2760      End Select
2770    End With

EXITP:
2780    Exit Sub

ERRH:
2790    Select Case ERR.Number
        Case Else
2800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2810    End Select
2820    Resume EXITP

End Sub

Private Sub opgAccountNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2920    intShiftDown = (Shift And acShiftMask) > 0
2930    intAltDown = (Shift And acAltMask) > 0
2940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2970        With Me
2980          intRetVal = 0
2990          If .cmbAccounts.Enabled = True Then
3000            .cmbAccounts.SetFocus
3010          Else
3020            .cmdPreview.SetFocus
3030          End If
3040        End With
3050      End Select
3060    End If

        ' ** Shift keys.
3070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3080      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3090        With Me
3100          intRetVal = 0
3110          .cmdClose.SetFocus
3120        End With
3130      End Select
3140    End If

EXITP:
3150    KeyCode = intRetVal
3160    Exit Sub

ERRH:
3170    intRetVal = 0
3180    Select Case ERR.Number
        Case Else
3190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3200    End Select
3210    Resume EXITP

End Sub

Private Sub opgAccountNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3320    intShiftDown = (Shift And acShiftMask) > 0
3330    intAltDown = (Shift And acAltMask) > 0
3340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3370        With Me
3380          intRetVal = 0
3390          If .cmbAccounts.Enabled = True Then
3400            .cmbAccounts.SetFocus
3410          Else
3420            .cmdPreview.SetFocus
3430          End If
3440        End With
3450      End Select
3460    End If

        ' ** Shift keys.
3470    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3480      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3490        With Me
3500          intRetVal = 0
3510          .cmdClose.SetFocus
3520        End With
3530      End Select
3540    End If

EXITP:
3550    KeyCode = intRetVal
3560    Exit Sub

ERRH:
3570    intRetVal = 0
3580    Select Case ERR.Number
        Case Else
3590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3600    End Select
3610    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

3710    With Me
3720      Select Case .chkPageOf
          Case True
3730        .chkPageOf_lbl.FontBold = True
3740      Case False
3750        .chkPageOf_lbl.FontBold = False
3760      End Select
3770    End With

EXITP:
3780    Exit Sub

ERRH:
3790    Select Case ERR.Number
        Case Else
3800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3810    End Select
3820    Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

3910    strAccountNo = vbNullString

3920    With Me
3930      If IsNull(.cmbAccounts) = False Then
3940        If Len(.cmbAccounts.Column(0)) > 0 Then
3950          strAccountNo = .cmbAccounts.Column(0)
3960        End If
3970      End If
3980      Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
3990        .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
4000        .opgAccountSource_optNumber_lbl2.FontBold = True
4010        .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
4020        .opgAccountSource_optName_lbl2.FontBold = False
4030        .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
4040      Case .opgAccountSource_optName.OptionValue
4050        .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
4060        .opgAccountSource_optNumber_lbl2.FontBold = False
4070        .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
4080        .opgAccountSource_optName_lbl2.FontBold = True
4090        .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
4100      End Select
4110      DoEvents
4120      If strAccountNo <> vbNullString Then
4130        .cmbAccounts = strAccountNo
4140      End If
4150    End With

EXITP:
4160    Exit Sub

ERRH:
4170    Select Case ERR.Number
        Case Else
4180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4190    End Select
4200    Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

4310    With Me
4320      Select Case .chkRememberMe
          Case True
4330        .chkRememberMe_lbl.FontBold = True
4340        .chkRememberMe_lbl2_dim.FontBold = True
4350        .chkRememberMe_lbl2_dim_hi.FontBold = True
4360      Case False
4370        .chkRememberMe_lbl.FontBold = False
4380        .chkRememberMe_lbl2_dim.FontBold = False
4390        .chkRememberMe_lbl2_dim_hi.FontBold = False
4400      End Select
4410    End With

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

4510    With Me
4520      If .cmbAccounts.text = vbNullString Then
4530        ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
4540        .cmbAccounts = .cmbAccounts.ItemData(0)
4550      End If
4560    End With

EXITP:
4570    Exit Sub

ERRH:
4580    Select Case ERR.Number
        Case Else
4590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4600    End Select
4610    Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4720    intShiftDown = (Shift And acShiftMask) > 0
4730    intAltDown = (Shift And acAltMask) > 0
4740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4770        With Me
4780          intRetVal = 0
4790          .cmdPreview.SetFocus
4800        End With
4810      End Select
4820    End If

        ' ** Shift keys.
4830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4850        With Me
4860          intRetVal = 0
4870          .opgAccountNumber.SetFocus
4880        End With
4890      End Select
4900    End If

EXITP:
4910    KeyCode = intRetVal
4920    Exit Sub

ERRH:
4930    intRetVal = 0
4940    Select Case ERR.Number
        Case Else
4950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4960    End Select
4970    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

5010    With Me
5020      If DoReport = True Then  ' ** Function: Below.
5030        Select Case gblnSeparateCheckingAccounts
            Case True
5040          strDocName = "rptAccountProfile_SepChkAcct"
5050        Case False
5060          strDocName = "rptAccountProfile_SngChkAcct"
5070        End Select
5080        Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
5090          gstrAccountNo = .cmbAccounts
5100          DoCmd.OpenReport strDocName, acViewPreview
5110        Case .opgAccountNumber_optAll.OptionValue
5120          .cmbAccounts = "All"
5130          DoCmd.OpenReport strDocName, acViewPreview
5140        End Select
5150        DoCmd.Maximize
5160        DoCmd.RunCommand acCmdFitToWindow
5170      End If
5180    End With

EXITP:
5190    Exit Sub

ERRH:
5200    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
5210    Case Else
5220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5230    End Select
5240    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5320    intShiftDown = (Shift And acShiftMask) > 0
5330    intAltDown = (Shift And acAltMask) > 0
5340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5360      Select Case intRetVal
          Case vbKeyTab
5370        With Me
5380          intRetVal = 0
5390          .cmdPrint.SetFocus
5400        End With
5410      End Select
5420    End If

        ' ** Shift keys.
5430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5440      Select Case intRetVal
          Case vbKeyTab
5450        With Me
5460          intRetVal = 0
5470          If .cmbAccounts.Enabled = True Then
5480            .cmbAccounts.SetFocus
5490          Else
5500            .opgAccountNumber.SetFocus
5510          End If
5520        End With
5530      End Select
5540    End If

EXITP:
5550    KeyCode = intRetVal
5560    Exit Sub

ERRH:
5570    intRetVal = 0
5580    Select Case ERR.Number
        Case Else
5590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5600    End Select
5610    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String
        Dim intMode As Integer

5710    With Me
5720      If DoReport = True Then  ' ** Function: Below.
5730        Select Case gblnSeparateCheckingAccounts
            Case True
5740          strDocName = "rptAccountProfile_SepChkAcct"
5750        Case False
5760          strDocName = "rptAccountProfile_SngChkAcct"
5770        End Select
5780        If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
5790          intMode = acViewPreview
5800        Else
5810          intMode = acViewNormal
5820        End If
5830        Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
              '##GTR_Ref: rptAccountProfile_SepChkAcct
              '##GTR_Ref: rptAccountProfile_SngChkAcct
5840          gstrAccountNo = .cmbAccounts
5850          DoCmd.OpenReport strDocName, intMode
5860        Case .opgAccountNumber_optAll.OptionValue
5870          .cmbAccounts = "All"
5880          DoCmd.OpenReport strDocName, intMode
5890        End Select
5900      End If
5910    End With

EXITP:
5920    Exit Sub

ERRH:
5930    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
5940    Case Else
5950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5960    End Select
5970    Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6020    intShiftDown = (Shift And acShiftMask) > 0
6030    intAltDown = (Shift And acAltMask) > 0
6040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6060      Select Case intRetVal
          Case vbKeyTab
6070        With Me
6080          intRetVal = 0
6090          .cmdWord.SetFocus
6100        End With
6110      End Select
6120    End If

        ' ** Shift keys.
6130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6140      Select Case intRetVal
          Case vbKeyTab
6150        With Me
6160          intRetVal = 0
6170          .cmdPreview.SetFocus
6180        End With
6190      End Select
6200    End If

EXITP:
6210    KeyCode = intRetVal
6220    Exit Sub

ERRH:
6230    intRetVal = 0
6240    Select Case ERR.Number
        Case Else
6250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6260    End Select
6270    Resume EXITP

End Sub

Private Sub cmdWord_Click()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

6310    With Me
6320      If DoReport = True Then  ' ** Function: Below.

6330        If IsNull(.UserReportPath) = True Then
6340          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
6350        Else
6360          strRptPath = .UserReportPath
6370        End If

6380        Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
6390          gstrAccountNo = .cmbAccounts
6400        Case .opgAccountNumber_optAll.OptionValue
6410          gstrAccountNo = "All"
6420        End Select

6430        Select Case gblnSeparateCheckingAccounts
            Case True
6440          strRpt = "rptAccountProfile_SepChkAcct"
6450        Case False
6460          strRpt = "rptAccountProfile_SngChkAcct"
6470        End Select
6480        strRptCap = "AccountProfile_" & gstrAccountNo & "_" & Format$(Date, "yyyymmdd")

6490        strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

6500        If strRptPathFile <> vbNullString Then
6510          DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
6520          .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
6530        End If

6540      End If
6550    End With

EXITP:
6560    Exit Sub

ERRH:
6570    Select Case ERR.Number
        Case Else
6580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6590    End Select
6600    Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6720    intShiftDown = (Shift And acShiftMask) > 0
6730    intAltDown = (Shift And acAltMask) > 0
6740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6760      Select Case intRetVal
          Case vbKeyTab
6770        With Me
6780          intRetVal = 0
6790          If .cmdExcel.Enabled = True Then
6800            .cmdExcel.SetFocus
6810          Else
6820            .cmdClose.SetFocus
6830          End If
6840        End With
6850      End Select
6860    End If

        ' ** Shift keys.
6870    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6880      Select Case intRetVal
          Case vbKeyTab
6890        With Me
6900          intRetVal = 0
6910          .cmdPrint.SetFocus
6920        End With
6930      End Select
6940    End If

EXITP:
6950    KeyCode = intRetVal
6960    Exit Sub

ERRH:
6970    intRetVal = 0
6980    Select Case ERR.Number
        Case Else
6990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7000    End Select
7010    Resume EXITP

End Sub

Private Sub cmdExcel_Click()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim strTmp01 As String

7110    With Me
7120      If DoReport = True Then  ' ** Function: Below.

7130        If IsNull(.UserReportPath) = True Then
7140          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
7150        Else
7160          strRptPath = .UserReportPath
7170        End If

7180        Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
7190          gstrAccountNo = .cmbAccounts
7200          strTmp01 = "Account Profile"
7210        Case .opgAccountNumber_optAll.OptionValue
7220          gstrAccountNo = "All"
7230          strTmp01 = "Account Profiles"
7240        End Select
7250        strRptCap = "AccountProfile_" & gstrAccountNo & "_" & Format$(Date, "yyyymmdd")

7260        strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

7270        If strRptPathFile <> vbNullString Then

7280          gstrFormQuerySpec = THIS_NAME

7290          Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
7300            Select Case gblnSeparateCheckingAccounts
                Case True
                  ' ** qryRpt_AccountProfile_11_02_08 (xx), For Export.
7310              strQry = "qryRpt_AccountProfile_11_01_09"
7320            Case False
                  ' ** qryRpt_AccountProfile_11_02_08 (xx), For Export.
7330              strQry = "qryRpt_AccountProfile_11_02_09"
7340            End Select
7350          Case .opgAccountNumber_optAll.OptionValue
7360            Select Case gblnSeparateCheckingAccounts
                Case True
                  ' ** qryRpt_AccountProfile_10_01_09 (xx), For Export.
7370              strQry = "qryRpt_AccountProfile_10_01_10"
7380            Case False
                  ' ** qryRpt_AccountProfile_10_02_09 (xx), For Export.
7390              strQry = "qryRpt_AccountProfile_10_02_10"
7400            End Select
7410          End Select

7420          DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False

7430          If Excel_NameOnly(strRptPathFile, strTmp01) = True Then  ' ** Module Function: modExcelFuncs.
7440            If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
7450              EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
7460            End If
7470            OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
7480          End If

7490          .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.

7500        End If

7510      End If
7520    End With

EXITP:
7530    Exit Sub

ERRH:
7540    Select Case ERR.Number
        Case Else
7550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7560    End Select
7570    Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7620    intShiftDown = (Shift And acShiftMask) > 0
7630    intAltDown = (Shift And acAltMask) > 0
7640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7660      Select Case intRetVal
          Case vbKeyTab
7670        With Me
7680          intRetVal = 0
7690          .cmdClose.SetFocus
7700        End With
7710      End Select
7720    End If

        ' ** Shift keys.
7730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7740      Select Case intRetVal
          Case vbKeyTab
7750        With Me
7760          intRetVal = 0
7770          .cmdWord.SetFocus
7780        End With
7790      End Select
7800    End If

EXITP:
7810    KeyCode = intRetVal
7820    Exit Sub

ERRH:
7830    intRetVal = 0
7840    Select Case ERR.Number
        Case Else
7850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7860    End Select
7870    Resume EXITP

End Sub

Private Function DoReport() As Boolean

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

7910    blnRetVal = True

7920    With Me

7930      Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
7940        If IsNull(.cmbAccounts) = True Then
7950          blnRetVal = False
7960          MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
7970        Else
7980          If .cmbAccounts = vbNullString Then
7990            blnRetVal = False
8000            MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
8010          End If
8020        End If
8030      Case .opgAccountNumber_optAll.OptionValue
            ' ** Nothing required.
8040      End Select

8050      If blnRetVal = True Then

8060        gstrFormQuerySpec = THIS_NAME

8070        Set dbs = CurrentDb

8080        Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
8090          Select Case gblnSeparateCheckingAccounts
              Case True
                ' ** Account, w/add'l fields, for separate checking; by specified FormRef('accountno').
8100            Set qdf = dbs.QueryDefs("qryRpt_AccountProfile_03b")
8110          Case False
                ' ** Account, w/add'l fields, for single checking; by specified FormRef('accountno').
8120            Set qdf = dbs.QueryDefs("qryRpt_AccountProfile_03d")
8130          End Select
8140        Case .opgAccountNumber_optAll.OptionValue
8150          Select Case gblnSeparateCheckingAccounts
              Case True
                ' ** Account, w/add'l fields, for separate checking; all accounts.
8160            Set qdf = dbs.QueryDefs("qryRpt_AccountProfile_03a")
8170          Case False
                ' ** Account, w/add'l fields, for single checking; all accounts.
8180            Set qdf = dbs.QueryDefs("qryRpt_AccountProfile_03c")
8190          End Select
8200        End Select

8210        Set rst = qdf.OpenRecordset
8220        With rst
8230          If .BOF = True And .EOF = True Then
8240            blnRetVal = False
8250            MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
8260          End If
8270          .Close
8280        End With

8290        dbs.Close

8300      End If  ' ** blnRetVal.

8310    End With  ' ** Me.

EXITP:
8320    Set qdf = Nothing
8330    Set rst = Nothing
8340    Set dbs = Nothing
8350    DoReport = blnRetVal
8360    Exit Function

ERRH:
8370    blnRetVal = False
8380    Select Case ERR.Number
        Case Else
8390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8400    End Select
8410    Resume EXITP

End Function
