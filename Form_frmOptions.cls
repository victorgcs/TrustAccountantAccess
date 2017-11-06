VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmOptions"

'VGC 09/22/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Account Type:     A {AccountNoWithType}
' **   Cancel:           C {cmdCancel}
' **   Disallow:         D {opgDataEntry_optDisallow}
' **   Cap Gain/Loss:    G {SpecialCapGainLoss}
' **   Income/Expense:   I {RevenueExpenseTracking}
' **   Tab Key:          K {TabCopyAccount}
' **   Link:             L {LinkRevTaxCodes}
' **   OK:               O {cmdOK}
' **   Checking:         S {SeparateCheckingAccounts}
' **   Income Tax:       T {IncomeTaxCoding}
' **   Allow:            W {opgDataEntry_optAllow}
' **   Exit:             X {cmdOK}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ***********************************************
' ** Easter Egg (Secret Key):
' ** Header_lbl_DblClick  Force w/ Type
' ***********************************************

Private blnForce As Boolean
Private blnTmp_IncomeTaxCoding As Boolean
Private blnTmp_RevenueExpenseTracking As Boolean
Private blnTmp_AccountNoWithType As Boolean
Private blnTmp_SeparateCheckingAccounts As Boolean
Private blnTmp_TabCopyAccount As Boolean
Private blnTmp_LinkRevTaxCodes As Boolean
Private blnTmp_SpecialCapGainLoss As Boolean
Private intTmp_SpecialCapGainLossOpt As Integer
Private intTmp_DataEntry As Integer

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     blnRetVal = True

120     With Me
130       If IsNull(.OpenArgs) = False Then
140         If Trim(.OpenArgs) <> vbNullString Then

150           strCallingForm = Trim(.OpenArgs)

              ' ** Assure the CompanyInformation variables are set.
160           blnRetVal = CoOptions_Read  ' ** Module Function: modStartupFuncs.
170           Select Case blnRetVal
              Case True

                ' ** Start with the opening choices.
180             blnTmp_IncomeTaxCoding = gblnIncomeTaxCoding
190             blnTmp_RevenueExpenseTracking = gblnRevenueExpenseTracking
200             blnTmp_AccountNoWithType = gblnAccountNoWithType
210             blnTmp_SeparateCheckingAccounts = gblnSeparateCheckingAccounts
220             blnTmp_TabCopyAccount = gblnTabCopyAccount
230             blnTmp_LinkRevTaxCodes = gblnLinkRevTaxCodes
240             blnTmp_SpecialCapGainLoss = gblnSpecialCapGainLoss
250             intTmp_DataEntry = gintSpecialCapGainLossOpt
260             DoEvents

270           Case False
280             Cancel = -1
290             DoCmd.Hourglass False
300             MsgBox "Unable to read options!", vbCritical + vbOKOnly, "Vital Information Not Found"
310           End Select

320         Else
330           Cancel = -1
340         End If
350       Else
360         Cancel = -1
370       End If
380     End With

390     If Cancel = -1 And blnRetVal = True Then
400       DoCmd.Hourglass False
410       Beep
420       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
430     End If

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

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me
520       If gblnMessage = True Then
530         gblnMessage = False
540         .CoInfo_lbl.Visible = True
550       End If
560       If .IncomeTaxCoding = True Then
570         .IncomeTaxCoding_lbl.FontBold = True
580         .IncomeTaxCoding_box.Visible = True
590       End If
600       If .RevenueExpenseTracking = True Then
610         .RevenueExpenseTracking_lbl.FontBold = True
620         .RevenueExpenseTracking_box.Visible = True
630       End If
640       If .AccountNoWithType = True Then
650         .AccountNoWithType_lbl.FontBold = True
660         .AccountNoWithType_box.Visible = True
670       End If
680       If .SeparateCheckingAccounts = True Then
690         .SeparateCheckingAccounts_lbl.FontBold = True
700         .SeparateCheckingAccounts_box.Visible = True
710       End If
720       If .TabCopyAccount = True Then
730         .TabCopyAccount_lbl.FontBold = True
740         .TabCopyAccount_box.Visible = True
750       End If
760       DoEvents
770       If .LinkRevTaxCodes = True Then
780         If .IncomeTaxCoding = False Or .RevenueExpenseTracking = False Then
790           Beep
800           MsgBox "The option 'Link Income/Expense Codes With Tax Codes' is invalid" & vbCrLf & _
                "without both 'Income Tax Tracking' and 'Income/Expense Tracking'" & vbCrLf & _
                "also chosen." & vbCrLf & vbCrLf & "The option will be deselected.", _
                vbInformation + vbOKOnly, "Invalid Entry"
810           .LinkRevTaxCodes = False
820           .LinkRevTaxCodes_box.Visible = False
830           blnTmp_LinkRevTaxCodes = False
840         Else
850           .LinkRevTaxCodes_lbl.FontBold = True
860           .LinkRevTaxCodes_box.Visible = True
870         End If
880       End If
890       DoEvents
900       If .IncomeTaxCoding = True And .RevenueExpenseTracking = True Then
910         .LinkRevTaxCodes.Enabled = True
920         .LinkRevTaxCodes_lbl2.ForeColor = CLR_DKGRY2
930         .LinkRevTaxCodes_lbl3.ForeColor = CLR_DKGRY2
940         .LinkRevTaxCodes_lbl2_dim_hi.Visible = False
950         .LinkRevTaxCodes_lbl3_dim_hi.Visible = False
960       Else
970         .LinkRevTaxCodes.Enabled = False
980         .LinkRevTaxCodes_lbl2.ForeColor = WIN_CLR_DISF
990         .LinkRevTaxCodes_lbl3.ForeColor = WIN_CLR_DISF
1000        .LinkRevTaxCodes_lbl2_dim_hi.Visible = True
1010        .LinkRevTaxCodes_lbl3_dim_hi.Visible = True
1020      End If
1030      .opgDataEntry.Visible = True
1040      .opgDataEntry_box.Visible = True
1050      .opgDataEntry_vline01.Visible = True
1060      .opgDataEntry_vline02.Visible = True
1070      DoEvents
1080      Select Case .SpecialCapGainLoss
          Case True
1090        .SpecialCapGainLoss_lbl.FontBold = True
1100        .SpecialCapGainLoss_box.Visible = True
1110        Select Case CurrentUser  ' ** Internal Access Function: Trust Accountant login.
            Case "TAAdmin"
1120          .opgDataEntry = .opgDataEntry_optAllow.OptionValue
1130          .opgDataEntry.Enabled = False
1140        Case "Superuser"
1150          .opgDataEntry = 0
1160          .opgDataEntry.Enabled = False
1170        Case Else
1180          .opgDataEntry = .opgDataEntry_optAllow.OptionValue
1190          .opgDataEntry.Enabled = True
1200        End Select
1210      Case False
1220        .SpecialCapGainLoss_lbl.FontBold = False
1230        .SpecialCapGainLoss_box.Visible = False
1240        .opgDataEntry = 0
1250        .opgDataEntry.Enabled = False
1260      End Select
1270      DoEvents
1280      opgDataEntry_AfterUpdate  ' ** Procedure: Below
          ' ** This happens after all Form_Open() code,
          ' ** so visibility has already been set.
1290      DoEvents
1300      DoCmd.Hourglass False
1310    End With

EXITP:
1320    Exit Sub

ERRH:
1330    DoCmd.Hourglass False
1340    Select Case ERR.Number
        Case Else
1350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1360    End Select
1370    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1420    intShiftDown = (Shift And acShiftMask) > 0
1430    intAltDown = (Shift And acAltMask) > 0
1440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyEscape
1470        intRetVal = 0
1480        cmdCancel_Click  ' ** Procedure: Below.
1490      End Select
1500    End If

        ' ** Alt keys.
1510    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1520      Select Case intRetVal
          Case vbKeyX
1530        intRetVal = 0
1540        cmdCancel_Click  ' ** Procedure: Below.
1550      End Select
1560    End If

        ' ** Ctrl keys.
1570    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1580      Select Case intRetVal
          Case vbKeyS
1590        intRetVal = 0
1600        cmdSave_Click  ' ** Procedure: Below.
1610      End Select
1620    End If

        ' ** Ctrl-Shift keys.
1630    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1640      Select Case intRetVal
          Case vbKeyF
1650        With Me
1660          intRetVal = 0
1670          .FocusHolder.SetFocus
1680        End With
1690      End Select
1700    End If

EXITP:
1710    KeyCode = intRetVal
1720    Exit Sub

ERRH:
1730    intRetVal = 0
1740    Select Case ERR.Number
        Case Else
1750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1760    End Select
1770    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

1810    With Me
1820      .Username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1830      .CoInfo_DateModified = Now()
1840    End With

EXITP:
1850    Exit Sub

ERRH:
1860    Select Case ERR.Number
        Case Else
1870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1880    End Select
1890    Resume EXITP

End Sub

Private Sub Form_Timer()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1910    With Me
1920      .TimerInterval = 0&
1930      If gblnSetFocus = True Then
1940        gblnSetFocus = False
1950        strFormName = vbNullString
1960  On Error Resume Next
1970        Do Until strFormName = THIS_NAME
1980          DoCmd.SelectObject acForm, THIS_NAME, False
1990          strFormName = Screen.ActiveForm.Name
2000        Loop
2010  On Error GoTo ERRH
2020      End If
2030      DoEvents
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

Private Sub Form_Unload(Cancel As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2110    DoCmd.Hourglass False
2120    DoEvents

2130    Select Case strCallingForm
        Case "frmMenu_Utility"
2140      gstrReturningForm = THIS_NAME
2150      gblnSetFocus = True
2160      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2170    End Select

EXITP:
2180    Exit Sub

ERRH:
2190    Select Case ERR.Number
        Case Else
2200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2210    End Select
2220    Resume EXITP

End Sub

Private Sub cmdOK_Click()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnContinue As Boolean

2310    blnContinue = True

2320    With Me

2330      If gblnIncomeTaxCoding = blnTmp_IncomeTaxCoding And gblnRevenueExpenseTracking = blnTmp_RevenueExpenseTracking And _
              gblnAccountNoWithType = blnTmp_AccountNoWithType And gblnSeparateCheckingAccounts = blnTmp_SeparateCheckingAccounts And _
              gblnTabCopyAccount = blnTmp_TabCopyAccount And gblnLinkRevTaxCodes = blnTmp_LinkRevTaxCodes And _
              gblnSpecialCapGainLoss = blnTmp_SpecialCapGainLoss And gintSpecialCapGainLossOpt = intTmp_DataEntry Then
2340        MsgBox "No changes will be made to the program options.", vbInformation + vbOKOnly, "Nothing To Do"
2350        DoCmd.Close acForm, THIS_NAME
2360      Else

2370        DoCmd.Hourglass True
2380        DoEvents

2390        If blnTmp_AccountNoWithType <> gblnAccountNoWithType Then
2400          blnContinue = AccountNoWithType_Set  ' ** Function: Below.
2410        End If

2420        If blnContinue = True Then

2430          DoCmd.Hourglass True
2440          DoEvents

2450          Set dbs = CurrentDb
2460          With dbs

2470            Set rst = .OpenRecordset("CompanyInformation", dbOpenDynaset, dbConsistent)
2480            With rst
2490              .MoveFirst
2500              .Edit
2510              ![IncomeTaxCoding] = blnTmp_IncomeTaxCoding
2520              ![RevenueExpenseTracking] = blnTmp_RevenueExpenseTracking
2530              ![AccountNoWithType] = blnTmp_AccountNoWithType
2540              ![SeparateCheckingAccounts] = blnTmp_SeparateCheckingAccounts
2550              ![TabCopyAccount] = blnTmp_TabCopyAccount
2560              ![LinkRevTaxCodes] = blnTmp_LinkRevTaxCodes
2570              ![SpecialCapGainLoss] = blnTmp_SpecialCapGainLoss
2580              ![SpecialCapGainLossOpt] = intTmp_DataEntry
2590              ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2600              ![CoInfo_DateModified] = Now()
2610              .Update
2620              .Close
2630            End With
2640            Set rst = Nothing

2650            Select Case blnTmp_SpecialCapGainLoss
                Case True
                  ' ** Update qryOptions_SpecialCap_01 (AssetType, just '80', with Dividend_new = True).
2660              Set qdf = .QueryDefs("qryOptions_SpecialCap_02")
2670            Case False
                  ' ** Update qryOptions_SpecialCap_01 (AssetType, just '80', with Dividend_new = False).
2680              Set qdf = .QueryDefs("qryOptions_SpecialCap_04")
2690            End Select
2700            qdf.Execute
2710            Set qdf = Nothing

2720            .Close
2730          End With
2740          Set dbs = Nothing

2750          gblnIncomeTaxCoding = blnTmp_IncomeTaxCoding
2760          gblnRevenueExpenseTracking = blnTmp_RevenueExpenseTracking
2770          gblnAccountNoWithType = blnTmp_AccountNoWithType
2780          gblnSeparateCheckingAccounts = blnTmp_SeparateCheckingAccounts
2790          gblnTabCopyAccount = blnTmp_TabCopyAccount
2800          gblnLinkRevTaxCodes = blnTmp_LinkRevTaxCodes
2810          gblnSpecialCapGainLoss = blnTmp_SpecialCapGainLoss
2820          gintSpecialCapGainLossOpt = intTmp_DataEntry

2830          DoCmd.Close acForm, THIS_NAME

2840        End If

2850      End If

2860    End With

EXITP:
2870    Set rst = Nothing
2880    Set qdf = Nothing
2890    Set dbs = Nothing
2900    Exit Sub

ERRH:
2910    DoCmd.Hourglass False
2920    Select Case ERR.Number
        Case Else
2930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2940    End Select
2950    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
          Case vbKeyTab
3070        With Me
3080          intRetVal = 0
3090          .cmdCancel.SetFocus
3100        End With
3110      End Select
3120    End If

        ' ** Shift keys.
3130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3140      Select Case intRetVal
          Case vbKeyTab
3150        With Me
3160          intRetVal = 0
3170          If .opgDataEntry.Visible = True And .opgDataEntry.Enabled = True Then
3180            .opgDataEntry.SetFocus
3190          Else
3200            .SpecialCapGainLoss.SetFocus
3210          End If
3220        End With
3230      End Select
3240    End If

EXITP:
3250    KeyCode = intRetVal
3260    Exit Sub

ERRH:
3270    intRetVal = 0
3280    Select Case ERR.Number
        Case Else
3290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3300    End Select
3310    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

3410    DoCmd.Hourglass True  'THIS WON'T TURN ON RIGHT AWAY!
3420    DoEvents

3430    DoCmd.Close acForm, THIS_NAME

EXITP:
3440    Exit Sub

ERRH:
3450    DoCmd.Hourglass False
3460    Select Case ERR.Number
        Case Else
3470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3480    End Select
3490    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
3590          .IncomeTaxCoding.SetFocus
3600        End With
3610      End Select
3620    End If

        ' ** Shift keys.
3630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3640      Select Case intRetVal
          Case vbKeyTab
3650        With Me
3660          intRetVal = 0
3670          .cmdOK.SetFocus
3680        End With
3690      End Select
3700    End If

EXITP:
3710    KeyCode = intRetVal
3720    Exit Sub

ERRH:
3730    intRetVal = 0
3740    Select Case ERR.Number
        Case Else
3750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3760    End Select
3770    Resume EXITP

End Sub

Private Sub cmdSave_Click()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3810    DoCmd.RunCommand acCmdSaveRecord
3820    DoEvents

EXITP:
3830    Exit Sub

ERRH:
3840    Select Case ERR.Number
        Case Else
3850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3860    End Select
3870    Resume EXITP

End Sub

Private Sub Header_lbl_DblClick(Cancel As Integer)
' ** Easter Egg (Secret Key) to force AccountNo With Type.

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_DblClick"

        Dim blnContinue As Boolean

3910    blnContinue = True

3920    With Me
3930      glngInstance = 0&
3940      DoCmd.OpenForm "frmOptions_Force", , , , , acDialog, THIS_NAME
3950      Select Case glngInstance
          Case 0&  ' ** Cancel.
            ' ** Do nothing.
3960        blnContinue = False
3970      Case 1&  ' ** Add.
3980        blnForce = True
3990        .AccountNoWithType = True
4000      Case 2&  ' ** Remove.
4010        blnForce = True
4020        .AccountNoWithType = False
4030      End Select
4040      If blnContinue = True Then
4050        AccountNoWithType_AfterUpdate  ' ** Procedure: Below.
4060      End If
4070    End With

EXITP:
4080    Exit Sub

ERRH:
4090    Select Case ERR.Number
        Case Else
4100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4110    End Select
4120    Resume EXITP

End Sub

Private Sub IncomeTaxCoding_AfterUpdate()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "IncomeTaxCoding_AfterUpdate"

4210    With Me
4220      Select Case .IncomeTaxCoding
          Case True
4230        .IncomeTaxCoding_lbl.FontBold = True
4240        .IncomeTaxCoding_box.Visible = True
4250        If .RevenueExpenseTracking = True And .LinkRevTaxCodes.Enabled = False Then
4260          .LinkRevTaxCodes.Enabled = True
4270          .LinkRevTaxCodes_lbl2.ForeColor = CLR_DKGRY2
4280          .LinkRevTaxCodes_lbl3.ForeColor = CLR_DKGRY2
4290          .LinkRevTaxCodes_lbl2_dim_hi.Visible = False
4300          .LinkRevTaxCodes_lbl3_dim_hi.Visible = False
4310        End If
4320      Case False
4330        .IncomeTaxCoding_lbl.FontBold = False
4340        .IncomeTaxCoding_box.Visible = False
4350        If .LinkRevTaxCodes.Enabled = True Then
4360          If .LinkRevTaxCodes = True Then
4370            Beep
4380            MsgBox "The option 'Link Income/Expense Codes With Tax Codes' is invalid" & vbCrLf & _
                  "without both 'Income Tax Tracking' and 'Income/Expense Tracking'" & vbCrLf & _
                  "also chosen." & vbCrLf & vbCrLf & "That option will be deselected.", _
                  vbInformation + vbOKOnly, "Invalid Entry"
4390            .LinkRevTaxCodes = False
4400            .LinkRevTaxCodes_box.Visible = False
4410            blnTmp_LinkRevTaxCodes = False
4420          End If
4430          .LinkRevTaxCodes.Enabled = False
4440          .LinkRevTaxCodes_lbl2.ForeColor = WIN_CLR_DISF
4450          .LinkRevTaxCodes_lbl3.ForeColor = WIN_CLR_DISF
4460          .LinkRevTaxCodes_lbl2_dim_hi.Visible = True
4470          .LinkRevTaxCodes_lbl3_dim_hi.Visible = True
4480        End If
4490      End Select
4500      cmdSave_Click  ' ** Procedure: Above.
4510      blnTmp_IncomeTaxCoding = .IncomeTaxCoding
4520    End With

EXITP:
4530    Exit Sub

ERRH:
4540    Select Case ERR.Number
        Case Else
4550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4560    End Select
4570    Resume EXITP

End Sub

Private Sub IncomeTaxCoding_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "IncomeTaxCoding_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
4670        With Me
4680          intRetVal = 0
4690          .RevenueExpenseTracking.SetFocus
4700        End With
4710      End Select
4720    End If

        ' ** Shift keys.
4730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4750        With Me
4760          intRetVal = 0
4770          .cmdCancel.SetFocus
4780        End With
4790      End Select
4800    End If

EXITP:
4810    KeyCode = intRetVal
4820    Exit Sub

ERRH:
4830    intRetVal = 0
4840    Select Case ERR.Number
        Case Else
4850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4860    End Select
4870    Resume EXITP

End Sub

Private Sub RevenueExpenseTracking_AfterUpdate()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "RevenueExpenseTracking_AfterUpdate"

4910    With Me
4920      Select Case .RevenueExpenseTracking
          Case True
4930        .RevenueExpenseTracking_lbl.FontBold = True
4940        .RevenueExpenseTracking_box.Visible = True
4950        If .IncomeTaxCoding = True And .LinkRevTaxCodes.Enabled = False Then
4960          .LinkRevTaxCodes.Enabled = True
4970          .LinkRevTaxCodes_lbl2.ForeColor = CLR_DKGRY2
4980          .LinkRevTaxCodes_lbl3.ForeColor = CLR_DKGRY2
4990          .LinkRevTaxCodes_lbl2_dim_hi.Visible = False
5000          .LinkRevTaxCodes_lbl3_dim_hi.Visible = False
5010        End If
5020      Case False
5030        .RevenueExpenseTracking_lbl.FontBold = False
5040        .RevenueExpenseTracking_box.Visible = False
5050        If .LinkRevTaxCodes.Enabled = True Then
5060          If .LinkRevTaxCodes = True Then
5070            Beep
5080            MsgBox "The option 'Link Income/Expense Codes With Tax Codes' is invalid" & vbCrLf & _
                  "without both 'Income Tax Tracking' and 'Income/Expense Tracking'" & vbCrLf & _
                  "also chosen." & vbCrLf & vbCrLf & "That option will be deselected.", _
                  vbInformation + vbOKOnly, "Invalid Entry"
5090            .LinkRevTaxCodes = False
5100            .LinkRevTaxCodes_box.Visible = False
5110            blnTmp_LinkRevTaxCodes = False
5120          End If
5130          .LinkRevTaxCodes.Enabled = False
5140          .LinkRevTaxCodes_lbl2.ForeColor = WIN_CLR_DISF
5150          .LinkRevTaxCodes_lbl3.ForeColor = WIN_CLR_DISF
5160          .LinkRevTaxCodes_lbl2_dim_hi.Visible = True
5170          .LinkRevTaxCodes_lbl3_dim_hi.Visible = True
5180        End If
5190      End Select
5200      cmdSave_Click  ' ** Procedure: Above.
5210      blnTmp_RevenueExpenseTracking = .RevenueExpenseTracking
5220    End With

EXITP:
5230    Exit Sub

ERRH:
5240    Select Case ERR.Number
        Case Else
5250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5260    End Select
5270    Resume EXITP

End Sub

Private Sub SeparateCheckingAccounts_AfterUpdate()

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingAccounts_AfterUpdate"

5310    With Me
5320      Select Case .SeparateCheckingAccounts
          Case True
5330        .SeparateCheckingAccounts_lbl.FontBold = True
5340        .SeparateCheckingAccounts_box.Visible = True
5350      Case False
5360        .SeparateCheckingAccounts_lbl.FontBold = False
5370        .SeparateCheckingAccounts_box.Visible = False
5380      End Select
5390      cmdSave_Click  ' ** Procedure: Above.
5400      blnTmp_SeparateCheckingAccounts = .SeparateCheckingAccounts
5410    End With

EXITP:
5420    Exit Sub

ERRH:
5430    Select Case ERR.Number
        Case Else
5440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5450    End Select
5460    Resume EXITP

End Sub

Private Sub TabCopyAccount_AfterUpdate()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "TabCopyAccount_AfterUpdate"

5510    With Me
5520      Select Case .TabCopyAccount
          Case True
5530        .TabCopyAccount_lbl.FontBold = True
5540        .TabCopyAccount_box.Visible = True
5550      Case False
5560        .TabCopyAccount_lbl.FontBold = False
5570        .TabCopyAccount_box.Visible = False
5580      End Select
5590      cmdSave_Click  ' ** Procedure: Above.
5600      blnTmp_TabCopyAccount = .TabCopyAccount
5610    End With

EXITP:
5620    Exit Sub

ERRH:
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Private Sub AccountNoWithType_AfterUpdate()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "AccountNoWithType_AfterUpdate"

        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String, strTmp02 As String

5710    With Me

5720      If blnForce = False Then
5730        Select Case .AccountNoWithType
            Case True
5740          strTmp01 = "This will add an Account Type prefix to all your account numbers."
5750          strTmp02 = Left("Add Account Type Prefix" & Space(80), 80)
5760        Case False
5770          strTmp01 = "This will remove the Account Type prefix from all your account numbers."
5780          strTmp02 = Left("Remove Account Type Prefix" & Space(80), 80)
5790        End Select
5800        Beep
5810        msgResponse = MsgBox(strTmp01 & vbCrLf & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo, strTmp02)
5820      Else
5830        msgResponse = vbYes
5840      End If

5850      If msgResponse <> vbYes Then
5860        .AccountNoWithType = gblnAccountNoWithType
5870      Else
5880        blnTmp_AccountNoWithType = .AccountNoWithType
5890      End If

5900      If blnTmp_AccountNoWithType <> gblnAccountNoWithType Then
5910        .AccountNoWithType_OK_lbl.Visible = True
5920      Else
5930        .AccountNoWithType_OK_lbl.Visible = False
5940      End If

5950      Select Case .AccountNoWithType
          Case True
5960        .AccountNoWithType_lbl.FontBold = True
5970        .AccountNoWithType_box.Visible = True
5980      Case False
5990        .AccountNoWithType_lbl.FontBold = False
6000        .AccountNoWithType_box.Visible = False
6010      End Select

6020    End With

EXITP:
6030    Exit Sub

ERRH:
6040    Select Case ERR.Number
        Case Else
6050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6060    End Select
6070    Resume EXITP

End Sub

Private Sub LinkRevTaxCodes_BeforeUpdate(Cancel As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "LinkRevTaxCodes_BeforeUpdate"

6110    With Me
6120      If .IncomeTaxCoding = False Or .RevenueExpenseTracking = False Then
6130        Cancel = -1
6140        Beep
6150        MsgBox "Both Income Tax Tracking and Income/Expense Tracking must be chosen" & vbCrLf & _
              "before this option can be used.", vbInformation + vbOKOnly, "Invalid Entry"
6160        .LinkRevTaxCodes.Undo
6170      End If
6180    End With

EXITP:
6190    Exit Sub

ERRH:
6200    Select Case ERR.Number
        Case Else
6210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6220    End Select
6230    Resume EXITP

End Sub

Private Sub LinkRevTaxCodes_AfterUpdate()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "LinkRevTaxCodes_AfterUpdate"

6310    With Me
6320      Select Case .LinkRevTaxCodes
          Case True
6330        .LinkRevTaxCodes_lbl.FontBold = True
6340        .LinkRevTaxCodes_box.Visible = True
6350      Case False
6360        .LinkRevTaxCodes_lbl.FontBold = False
6370        .LinkRevTaxCodes_box.Visible = False
6380      End Select
6390      cmdSave_Click  ' ** Procedure: Above.
6400      blnTmp_LinkRevTaxCodes = .LinkRevTaxCodes
6410    End With

EXITP:
6420    Exit Sub

ERRH:
6430    Select Case ERR.Number
        Case Else
6440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6450    End Select
6460    Resume EXITP

End Sub

Private Sub LinkRevTaxCodes_KeyDown(KeyCode As Integer, Shift As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "LinkRevTaxCodes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6520    intShiftDown = (Shift And acShiftMask) > 0
6530    intAltDown = (Shift And acAltMask) > 0
6540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6570        With Me
6580          intRetVal = 0
6590          .SpecialCapGainLoss.SetFocus
6600        End With
6610      End Select
6620    End If

        ' ** Shift keys.
6630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6650        With Me
6660          intRetVal = 0
6670          .AccountNoWithType.SetFocus
6680        End With
6690      End Select
6700    End If

EXITP:
6710    KeyCode = intRetVal
6720    Exit Sub

ERRH:
6730    intRetVal = 0
6740    Select Case ERR.Number
        Case Else
6750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6760    End Select
6770    Resume EXITP

End Sub

Private Sub SpecialCapGainLoss_AfterUpdate()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "SpecialCapGainLoss_AfterUpdate"

        Dim strCurrentUser As String

        ' ** Only Admins level have access to this form.
6810    strCurrentUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.

6820    With Me
6830      Select Case .SpecialCapGainLoss
          Case True
6840        .SpecialCapGainLoss_lbl.FontBold = True
6850        .SpecialCapGainLoss_box.Visible = True
6860        Select Case strCurrentUser
            Case "TAAdmin"
              ' ** Everyone just gets this setting.
6870          .opgDataEntry = .opgDataEntry_optAllow.OptionValue
6880          .opgDataEntry.Visible = True
6890          .opgDataEntry.Enabled = False
6900          .opgDataEntry_vline01.Visible = True
6910          .opgDataEntry_vline02.Visible = True
6920          .opgDataEntry_box.Visible = True
6930        Case "Superuser"
              ' ** Only I get this setting.
6940          .opgDataEntry = 0
6950          .opgDataEntry.Visible = True
6960          .opgDataEntry.Enabled = False
6970          .opgDataEntry_vline01.Visible = True
6980          .opgDataEntry_vline02.Visible = True
6990          .opgDataEntry_box.Visible = True
7000        Case Else  ' ** This is an Admins level user that isn't TAAdmin.
              ' ** Data Entry may or not get this setting.
7010          .opgDataEntry = .opgDataEntry_optAllow.OptionValue
7020          .opgDataEntry.Visible = True
7030          .opgDataEntry.Enabled = True
7040          .opgDataEntry_vline01.Visible = True
7050          .opgDataEntry_vline02.Visible = True
7060          .opgDataEntry_box.Visible = True
7070        End Select
7080        opgDataEntry_AfterUpdate  ' ** Procedure: Below
7090      Case False
7100        .SpecialCapGainLoss_lbl.FontBold = False
7110        .SpecialCapGainLoss_box.Visible = False
7120        Select Case strCurrentUser
            Case "TAAdmin"
              ' ** Everyone just gets this setting.
7130          .opgDataEntry = 0
7140          .opgDataEntry.Visible = True
7150          .opgDataEntry.Enabled = True
7160          .opgDataEntry_vline01.Visible = True
7170          .opgDataEntry_vline02.Visible = True
7180          .opgDataEntry_box.Visible = True
7190        Case "Superuser"
              ' ** Only I get this setting.
7200          .opgDataEntry = 0
7210          .opgDataEntry.Visible = True
7220          .opgDataEntry.Enabled = False
7230          .opgDataEntry_vline01.Visible = True
7240          .opgDataEntry_vline02.Visible = True
7250          .opgDataEntry_box.Visible = True
7260        Case Else  ' ** This is an Admins level user that isn't TAAdmin.
              ' ** They can't not have it, while Data Entry does.
7270          .opgDataEntry = 0
7280          .opgDataEntry.Visible = True
7290          .opgDataEntry.Enabled = False
7300          .opgDataEntry_vline01.Visible = True
7310          .opgDataEntry_vline02.Visible = True
7320          .opgDataEntry_box.Visible = True
7330        End Select
7340        opgDataEntry_AfterUpdate  ' ** Procedure: Below
7350      End Select
7360      cmdSave_Click  ' ** Procedure: Above.
7370      blnTmp_SpecialCapGainLoss = .SpecialCapGainLoss
7380      intTmp_SpecialCapGainLossOpt = .opgDataEntry
7390    End With

EXITP:
7400    Exit Sub

ERRH:
7410    Select Case ERR.Number
        Case Else
7420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7430    End Select
7440    Resume EXITP

End Sub

Private Sub SpecialCapGainLoss_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "SpecialCapGainLoss_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7520    intShiftDown = (Shift And acShiftMask) > 0
7530    intAltDown = (Shift And acAltMask) > 0
7540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7570        With Me
7580          intRetVal = 0
7590          If .opgDataEntry.Visible = True And .opgDataEntry.Enabled = True Then
7600            .opgDataEntry.SetFocus
7610          Else
7620            .cmdOK.SetFocus
7630          End If
7640        End With
7650      End Select
7660    End If

        ' ** Shift keys.
7670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7680      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7690        With Me
7700          intRetVal = 0
7710          .LinkRevTaxCodes.SetFocus
7720        End With
7730      End Select
7740    End If

EXITP:
7750    KeyCode = intRetVal
7760    Exit Sub

ERRH:
7770    intRetVal = 0
7780    Select Case ERR.Number
        Case Else
7790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7800    End Select
7810    Resume EXITP

End Sub

Private Sub opgDataEntry_AfterUpdate()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgDataEntry_AfterUpdate"

7910    With Me
7920      Select Case .opgDataEntry
          Case .opgDataEntry_optAllow.OptionValue
7930        .opgDataEntry_optAllow_lbl.FontBold = True
7940        .opgDataEntry_optDisallow_lbl.FontBold = False
7950      Case .opgDataEntry_optDisallow.OptionValue
7960        .opgDataEntry_optAllow_lbl.FontBold = False
7970        .opgDataEntry_optDisallow_lbl.FontBold = True
7980      Case Else  ' ** When Zero.
7990        .opgDataEntry_optAllow_lbl.FontBold = False
8000        .opgDataEntry_optDisallow_lbl.FontBold = False
8010      End Select
8020      cmdSave_Click  ' ** Procedure: Above.
8030      intTmp_SpecialCapGainLossOpt = .opgDataEntry
8040    End With

EXITP:
8050    Exit Sub

ERRH:
8060    Select Case ERR.Number
        Case Else
8070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8080    End Select
8090    Resume EXITP

End Sub

Private Sub opgDataEntry_optAllow_KeyDown(KeyCode As Integer, Shift As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgDataEntry_optAllow_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8120    intShiftDown = (Shift And acShiftMask) > 0
8130    intAltDown = (Shift And acAltMask) > 0
8140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8170        With Me
8180          intRetVal = 0
8190          .cmdOK.SetFocus
8200        End With
8210      End Select
8220    End If

        ' ** Shift keys.
8230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8250        With Me
8260          intRetVal = 0
8270          .SpecialCapGainLoss.SetFocus
8280        End With
8290      End Select
8300    End If

EXITP:
8310    KeyCode = intRetVal
8320    Exit Sub

ERRH:
8330    intRetVal = 0
8340    Select Case ERR.Number
        Case Else
8350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8360    End Select
8370    Resume EXITP

End Sub

Private Sub opgDataEntry_optDisallow_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgDataEntry_optDisallow_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8420    intShiftDown = (Shift And acShiftMask) > 0
8430    intAltDown = (Shift And acAltMask) > 0
8440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8470        With Me
8480          intRetVal = 0
8490          .cmdOK.SetFocus
8500        End With
8510      End Select
8520    End If

        ' ** Shift keys.
8530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8550        With Me
8560          intRetVal = 0
8570          .SpecialCapGainLoss.SetFocus
8580        End With
8590      End Select
8600    End If

EXITP:
8610    KeyCode = intRetVal
8620    Exit Sub

ERRH:
8630    intRetVal = 0
8640    Select Case ERR.Number
        Case Else
8650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8660    End Select
8670    Resume EXITP

End Sub

Private Function AccountNoWithType_Set() As Boolean

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "AccountNoWithType_Set"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngXRecs As Long, arr_varXRec As Variant
        Dim lngAcctsChanged As Long, lngHasRels As Long
        Dim lngAccts As Long, arr_varAcct() As Variant
        Dim lngRels As Long, arr_varRel() As Variant
        Dim lngRelAccts As Long, arr_varRelAcct As Variant
        Dim lngTooLongs As Long, msgResponse As VbMsgBoxResult
        Dim blnParentFound As Boolean, blnRemovePresent As Boolean, blnUpdateArchive As Boolean
        Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim blnContinue As Boolean
        Dim blnRetVal As Boolean

        ' ** Array: arr_varAcct().
        Const A_ELEMS As Integer = 7  ' ** Array's first-element UBound().
        Const A_ACTO  As Integer = 0  ' ** Old.
        Const A_TYP   As Integer = 1
        Const A_NOPFX As Integer = 2
        Const A_PFX   As Integer = 3
        Const A_ACTN  As Integer = 4  ' ** New.
        Const A_RELS  As Integer = 5
        Const A_RACTO As Integer = 6  ' ** Old.
        Const A_RACTN As Integer = 7  ' ** New.

        ' ** Array: arr_varRel().
        Const R_ELEMS As Integer = 7  ' ** Array's first-element UBound().
        Const R_ACTN  As Integer = 0
        Const R_TYP   As Integer = 1
        Const R_RELS  As Integer = 2
        Const R_RACTO As Integer = 3
        Const R_RACTN As Integer = 4
        Const R_ARR   As Integer = 5
        Const R_CHKD  As Integer = 6
        Const R_AELEM As Integer = 7

        ' ** Array: arr_varRelAcct().
        Const RA_ACTO As Integer = 0
        Const RA_PFX  As Integer = 1
        'Const RA_TYPO As Integer = 2
        Const RA_TYPN As Integer = 3
        Const RA_ACTN As Integer = 4

        ' ** Array: arr_varXRec().
        Const X_ACTO     As Integer = 0
        Const X_RACTO    As Integer = 1
        'Const X_HAS_PFX  As Integer = 2
        'Const X_HAS_RPFX As Integer = 3
        Const X_TYP      As Integer = 4
        'Const X_LEN      As Integer = 5
        Const X_HAS_RELS As Integer = 6
        Const X_PFX      As Integer = 7
        Const X_NOPFX    As Integer = 8

8710    blnRetVal = True

8720    blnContinue = True
8730    blnRemovePresent = False

8740    DoCmd.Hourglass True
8750    DoEvents

8760    With Me

          ' ** Write first, then update account numbers.
8770      If CoOptions_Write("AccountNoWithType", .AccountNoWithType) = True Then
            'blnTmp_AccountNoWithType = .AccountNoWithType.Value

8780        Set dbs = CurrentDb
8790        With dbs

              ' ** qryCompanyInformation_05, grouped, just accountno too long, with cnt.
8800          Set qdf = .QueryDefs("qryCompanyInformation_06")
8810          Set rst = qdf.OpenRecordset
8820          With rst
8830            If .BOF = True And .EOF = True Then
                  ' ** We're OK.
8840            Else
8850              .MoveFirst
8860              If IsNull(![cnt]) = True Then
                    ' ** We're still OK.
8870              Else
8880                lngTooLongs = ![cnt]
8890              End If
8900            End If
8910            .Close
8920          End With  ' ** rst.

8930          If lngTooLongs > 0& And Me.AccountNoWithType = True And blnForce = False Then
8940            If lngTooLongs = 1& Then
8950              strTmp01 = "One of your accounts has an existing number that" & vbCrLf & _
                    "is too long to accept the Account Type prefix."
8960            Else
8970              strTmp01 = CStr(lngTooLongs) & " of your accounts have existing numbers that" & vbCrLf & _
                    "are too long to accept the Account Type prefix."
8980            End If
8990            DoCmd.Hourglass False
9000            msgResponse = MsgBox(strTmp01 & vbCrLf & vbCrLf & _
                  "Do you wish to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Results Will Be Mixed")
9010            If msgResponse <> vbYes Then blnContinue = False
9020          End If

              'CHECK FOR NON-TYPE PREFIXES!!
              'IF THEY HAVE LOTS OF PREFIXES, AND ONLY ONE IS OFF, ASSUME IT'S A TYPO!
              'IF THEY HAVE LOTS OF PRIFIXES, WITH A VARIETY OF TYPE AND NON-TYPE NUMBERS, ASSUME ALL ARE NON-TYPE.
              'IF THEY HAVE ONLY ONE WITH A PREFIX, AND IT MATCHES ONE OF OURS, ASSUME IT NEEDS TO BE CLEANED UP.
              'IF THEY HAVE ONLY ONE WITH A PREFIX, AND IT DOESN'T MATCH ONE OF OURS, LEAVE IT ALONE, AND APPEND OURS.
              'IS 'ONE' AN OFFICIAL THRESHOLD?

9030          lngAccts = 0&
9040          ReDim arr_varAcct(A_ELEMS, 0)

9050          If blnContinue = True Then

9060            DoCmd.Hourglass True
9070            DoEvents

                ' ** qryCompanyInformation_03, with accountno_pfx, accountno_nopfx.
9080            Set qdf = .QueryDefs("qryCompanyInformation_05")  '#CRTC01
9090            Set rst = qdf.OpenRecordset
                ' ********************************************************
                ' ** Array: arr_varXRec()
                ' **
                ' **   Field  Element  Name                 Constant
                ' **   =====  =======  ===================  ============
                ' **     1       0     accountno            X_ACTO
                ' **     2       1     related_accountno    X_RACTO
                ' **     3       2     TypePfx              X_HAS_PFX
                ' **     4       3     TypePfx_Rel          X_HAS_RPFX
                ' **     5       4     accounttype          X_TYP
                ' **     6       5     actno_len            X_LEN
                ' **     7       6     HasRels              X_HAS_RELS
                ' **     8       7     accountno_pfx        X_PFX
                ' **     9       8     accountno_nopfx      X_NOPFX
                ' **
                ' ********************************************************
9100            With rst
9110              If .BOF = True And .EOF = True Then
                    ' ** There are no accounts!
9120                blnContinue = False
9130              Else
                    ' ** If their choice contradicts what's been found, it doesn't matter,
                    ' ** since we'll go through each record and only update as necessary.
                    ' ** And, accountno length is factored into accountno_pfx (too long or short, doesn't get one).
9140                .MoveLast
9150                lngXRecs = .RecordCount
9160                .MoveFirst
9170                arr_varXRec = .GetRows(lngXRecs)
9180                .MoveFirst
9190                lngAcctsChanged = 0&
9200                For lngX = 1& To lngXRecs
9210                  For lngY = 0& To (lngXRecs - 1&)
9220                    If arr_varXRec(X_ACTO, lngY) = ![accountno] Then
9230                      lngZ = lngY
9240                      Exit For
9250                    End If
9260                  Next
9270                  If arr_varXRec(X_HAS_RELS, lngZ) = True Then lngHasRels = lngHasRels + 1&
9280                  lngAccts = lngAccts + 1&  ' ** Array will be all accounts, to be used if lngHasRels > 0, otherwise it won't be used at all.
9290                  lngE = lngAccts - 1&
9300                  ReDim Preserve arr_varAcct(A_ELEMS, lngE)
9310                  arr_varAcct(A_ACTO, lngE) = arr_varXRec(X_ACTO, lngZ)
9320                  arr_varAcct(A_TYP, lngE) = arr_varXRec(X_TYP, lngZ)
9330                  arr_varAcct(A_NOPFX, lngE) = arr_varXRec(X_NOPFX, lngZ)
9340                  arr_varAcct(A_PFX, lngE) = arr_varXRec(X_PFX, lngZ)
9350                  Select Case arr_varXRec(X_HAS_RELS, lngZ)
                      Case True
9360                    arr_varAcct(A_RELS, lngE) = 1&  ' ** We'll count how many, later.
9370                    arr_varAcct(A_RACTO, lngE) = arr_varXRec(X_RACTO, lngZ)
9380                  Case False
9390                    arr_varAcct(A_RELS, lngE) = 0&
9400                    arr_varAcct(A_RACTO, lngE) = vbNullString  ' ** Old.
9410                  End Select
9420                  arr_varAcct(A_RACTN, lngE) = vbNullString    ' ** New.
9430                  Select Case Me.AccountNoWithType
                      Case True
9440                    If ![accountno] <> arr_varXRec(X_PFX, lngZ) Then
9450                      .Edit
9460                      ![accountno] = arr_varXRec(X_PFX, lngZ)
9470                      arr_varAcct(A_ACTN, lngE) = arr_varXRec(X_PFX, lngZ)
                          'WHAT IF, AT THIS MOMENT, OUR PRIFIX HAPPENS TO MATCH THEIR NON-TYPE PRIFIX?
                          'AND THE CONFLICTING NON-TYPE ONE, ONCE IT GETS OUR TYPE PREFIX, WILL THEN NO LONGER MATCH?
                          ' ** On Error Resume Next
9480                      .Update
                          ' ** If Err.Number <> 0 Then
9490                      lngAcctsChanged = lngAcctsChanged + 1&
9500                    End If
9510                  Case False
9520                    If ![accountno] <> arr_varXRec(X_NOPFX, lngZ) Then
9530                      .Edit
9540                      ![accountno] = arr_varXRec(X_NOPFX, lngZ)
9550                      arr_varAcct(A_ACTN, lngE) = arr_varXRec(X_NOPFX, lngZ)
9560                      .Update
9570                      lngAcctsChanged = lngAcctsChanged + 1&
9580                    End If
9590                  End Select
9600                  If lngX < lngXRecs Then .MoveNext
9610                Next  ' ** lngX.
9620              End If  ' ** Has accounts.
9630              .Close
9640            End With  ' ** rst.

9650          End If  ' ** blnContinue.

              ' ** Now go back and do the related_accountno's.
9660          If blnContinue = True Then
9670            If lngHasRels > 0& Then

                  ' ** Collect the subset of related_accountno's.
9680              lngRels = 0&
9690              ReDim arr_varRel(R_ELEMS, 0)
                  ' ***********************************************************
                  ' ** Array: arr_varRel()
                  ' **
                  ' **   Field  Element  Name                      Constant
                  ' **   =====  =======  ========================  ==========
                  ' **     1       0     accountno_new             R_ACTN
                  ' **     2       1     accounttype               R_TYP
                  ' **     3       2     Related Count             R_RELS
                  ' **     4       3     related_accountno_old     R_RACTO
                  ' **     5       4     related_accountno_new     R_RACTN
                  ' **     6       5     arr_varRelAcct() array    R_ARR
                  ' **     7       6     Checked                   R_CHKD
                  ' **     8       7     arr_varAcct() Element     R_AELEM
                  ' **
                  ' ***********************************************************
9700              For lngX = 0& To (lngAccts - 1&)
9710                If arr_varAcct(A_RELS, lngX) > 0& Then
9720                  lngRels = lngRels + 1&
9730                  lngE = lngRels - 1&
9740                  ReDim Preserve arr_varRel(R_ELEMS, lngE)
9750                  arr_varRel(R_ACTN, lngE) = arr_varAcct(A_ACTN, lngX)
9760                  arr_varRel(R_TYP, lngE) = arr_varAcct(A_TYP, lngX)
9770                  arr_varRel(R_RELS, lngE) = CLng(1)
9780                  arr_varRel(R_RACTO, lngE) = arr_varAcct(A_RACTO, lngX)
9790                  arr_varRel(R_RACTN, lngE) = vbNullString
9800                  arr_varRel(R_ARR, lngE) = Empty
9810                  arr_varRel(R_CHKD, lngE) = CBool(False)
9820                  arr_varRel(R_AELEM, lngE) = lngX
9830                End If
9840              Next  ' ** lngX.

                  ' ** Parse each one out.
9850              For lngX = 0& To (lngRels - 1&)
9860                arr_varRelAcct = HasTypePfx(arr_varRel(R_RACTO, lngX), True)  ' ** Module Function: modStringFuncs.
9870  On Error Resume Next
9880                lngRelAccts = CLng((UBound(arr_varRelAcct, 2) + 1&))
9890                If ERR.Number <> 0 Then
9900                  blnContinue = False
9910                  DoCmd.Hourglass False
9920                  Beep
9930                  MsgBox "An error occurred while updating Account Numbers." & vbCrLf & _
                        "    Error: " & CStr(ERR.Number) & vbCrLf & vbCrLf & _
                        "Check your Account Number list for errors.", vbCritical + vbOKOnly, "Account Type Error"
9940  On Error GoTo ERRH
9950                Else
9960  On Error GoTo ERRH
                      ' ***************************************************
                      ' ** Array: arr_varRelAcct()
                      ' **
                      ' **   Field  Element  Name              Constant
                      ' **   =====  =======  ================= ==========
                      ' **     1       0     accountno_old     RA_ACTO
                      ' **     2       1     TypePfx (T/F)     RA_PFX
                      ' **     3       2     accounttype_old   RA_TYPO
                      ' **     4       3     accounttype_new   RA_TYPN
                      ' **     5       4     accountno_new     RA_ACTN
                      ' **
                      ' ***************************************************
9970                  If arr_varRelAcct(RA_ACTN, 0) = RET_ERR Then
                        ' ** Skip the rest.
9980                    blnContinue = False
9990                    Exit For
10000                 Else
10010                   If arr_varRelAcct(RA_ACTN, 0) = "{Remove}" Then
10020                     blnRemovePresent = True
10030                     arr_varRel(R_RELS, lngX) = CLng(0)
10040                     arr_varRel(R_RACTN, lngX) = arr_varRelAcct(RA_ACTN, 0)
10050                     arr_varRel(R_ARR, lngX) = Empty
10060                   Else
10070                     arr_varRel(R_RELS, lngX) = lngRelAccts
10080                     arr_varRel(R_ARR, lngX) = arr_varRelAcct
10090                   End If
10100                 End If
10110                 lngRelAccts = 0&
10120                 arr_varRelAcct = Empty
10130               End If
10140             Next  ' ** lngX.

10150             If blnContinue = True Then

                    ' ** Assemble the new related_accountno field.
10160               Select Case Me.AccountNoWithType
                    Case True  ' ** Add AccountType prefixes.
                      ' ** Assemble the new type-prefixed related_accountno field.
10170                 For lngX = 0& To (lngRels - 1&)
10180                   If arr_varRel(R_RELS, lngX) > 0& And arr_varRel(R_RACTN, lngX) <> "{Remove}" Then
10190                     arr_varRelAcct = arr_varRel(R_ARR, lngX)
10200                     lngRelAccts = (UBound(arr_varRelAcct, 2) + 1&)
10210                     For lngY = 0& To (lngRelAccts - 1&)
10220                       If arr_varRelAcct(RA_PFX, lngY) = True Then  ' ** It already had a prefix.
10230                         If arr_varRelAcct(RA_ACTO, lngY) = arr_varRel(R_ACTN, lngX) Then
                                ' ** This is the parent.
10240                           arr_varRelAcct(RA_TYPN, lngY) = arr_varRel(R_TYP, lngX)   ' ** Get the correct accounttype.
10250                           arr_varRelAcct(RA_ACTN, lngY) = arr_varRel(R_ACTN, lngX)  ' ** Get accountno_new.
10260                         Else
10270                           For lngZ = 0& To (lngRels - 1&)
10280                             If arr_varRel(R_ACTN, lngZ) = arr_varRelAcct(RA_ACTO, lngY) Then
10290                               arr_varRelAcct(RA_TYPN, lngY) = arr_varRel(R_TYP, lngZ)   ' ** Get its correct accounttype.
10300                               arr_varRelAcct(RA_ACTN, lngY) = arr_varRel(R_ACTN, lngZ)  ' ** Make sure it has the proper accountno_new.
10310                               Exit For
10320                             End If
10330                           Next
10340                         End If
10350                       Else
10360                         For lngZ = 0& To (lngRels - 1&)
10370                           If arr_varAcct(A_NOPFX, arr_varRel(R_AELEM, lngZ)) = arr_varRelAcct(RA_ACTO, lngY) Then
10380                             arr_varRelAcct(RA_TYPN, lngY) = arr_varRel(R_TYP, lngZ)   ' ** Though we could check whether it's
10390                             arr_varRelAcct(RA_ACTN, lngY) = arr_varRel(R_ACTN, lngZ)  ' ** the parent, Just get them all this way.
10400                             Exit For
10410                           End If
10420                         Next
10430                       End If
10440                     Next  ' ** lngY.
10450                     strTmp01 = vbNullString
10460                     For lngY = 0& To (lngRelAccts - 1&)
10470                       strTmp01 = strTmp01 & arr_varRelAcct(RA_ACTN, lngY) & ","
10480                     Next
10490                     If Right(strTmp01, 1) = "," Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
10500                     arr_varRel(R_RACTN, lngX) = strTmp01
10510                   End If
10520                   lngRelAccts = 0&
10530                   arr_varRelAcct = Empty
10540                 Next  ' ** lngX.
10550               Case False  ' ** Remove AccountType prefixes.
                      ' ** Assemble the new non-type related_accountno field.
10560                 For lngX = 0& To (lngRels - 1&)
10570                   If arr_varRel(R_RELS, lngX) > 0& And arr_varRel(R_RACTN, lngX) <> "{Remove}" Then
10580                     arr_varRelAcct = arr_varRel(R_ARR, lngX)
10590                     lngRelAccts = (UBound(arr_varRelAcct, 2) + 1&)
10600                     For lngY = 0& To (lngRelAccts - 1&)
10610                       If arr_varRelAcct(RA_PFX, lngY) = True Then
10620                         arr_varRelAcct(RA_ACTN, lngY) = Mid(arr_varRelAcct(RA_ACTO, lngY), 4)
10630                       Else
10640                         arr_varRelAcct(RA_ACTN, lngY) = arr_varRelAcct(RA_ACTO, lngY)
10650                       End If
10660                     Next  ' ** lngY.
10670                     strTmp01 = vbNullString
10680                     For lngY = 0& To (lngRelAccts - 1&)
10690                       strTmp01 = strTmp01 & arr_varRelAcct(RA_ACTN, lngY) & ","
10700                     Next
10710                     If Right(strTmp01, 1) = "," Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
10720                     arr_varRel(R_RACTN, lngX) = strTmp01
10730                   End If
10740                   lngRelAccts = 0&
10750                   arr_varRelAcct = Empty
10760                 Next  ' ** lngX.
10770               End Select

                    ' ** Now check to see that the sets match among the participants!
10780               For lngX = 0& To (lngRels - 1&)
10790                 If arr_varRel(R_CHKD, lngX) = False Then  ' ** The later ones will have been checked by the first in the set.
10800                   If arr_varRel(R_RELS, lngX) > 0& And arr_varRel(R_RACTN, lngX) <> "{Remove}" Then
10810                     arr_varRelAcct = arr_varRel(R_ARR, lngX)
10820                     lngRelAccts = (UBound(arr_varRelAcct, 2) + 1&)
10830                     blnParentFound = False
10840                     For lngY = 0& To (lngRelAccts - 1&)
10850                       If arr_varRelAcct(RA_ACTN, lngY) = arr_varRel(R_ACTN, lngX) Then
10860                         If blnParentFound = False Then
10870                           blnParentFound = True
10880                         Else
                                ' ** It was in the list twice?
                                'SOMEHOW PULL THE DUPE OUT OF THE LIST.
10890                         End If
10900                       Else
                              ' ** See that this one's related_accountno matches!
                              ' ** If someone says A is related to B, and B is related to C, then all must include A, B, and C!
                              ' ** If one has 3 rels and another has 4 rels, tough! Whichever is encountered first wins the day.
10910                         For lngZ = 0& To (lngRels - 1&)
10920                           If lngZ <> lngX Then
10930                             If arr_varRel(R_ACTN, lngZ) = arr_varRelAcct(RA_ACTN, lngY) Then
10940                               If arr_varRel(R_CHKD, lngZ) = False Then  ' ** If it's already been checked, skip it!
10950                                 If arr_varRel(R_RACTN, lngZ) <> arr_varRel(R_RACTN, lngX) Then
10960                                   arr_varRel(R_RACTN, lngZ) = arr_varRel(R_RACTN, lngX)
10970                                 End If
10980                                 If arr_varRel(R_RELS, lngZ) <> arr_varRel(R_RELS, lngX) Then
10990                                   arr_varRel(R_RELS, lngZ) = arr_varRel(R_RELS, lngX)
11000                                 End If
11010                                 arr_varRel(R_CHKD, lngZ) = CBool(True)
11020                               End If
11030                               Exit For
11040                             End If
11050                           End If
11060                         Next
11070                       End If
11080                     Next
11090                   End If
11100                   arr_varRel(R_CHKD, lngX) = CBool(True)
11110                 End If
11120               Next  ' ** lngX.

                    'DOES ANYTHING CHECK TO MAKE SURE MISSING ONES HAVE BEEN ADDED?
                    ' ** Finally, update the Account table.
11130               Set qdf = .QueryDefs("qryCompanyInformation_05")  '#CRTC01
11140               Set rst = qdf.OpenRecordset
11150               With rst
11160                 .MoveFirst
11170                 For lngX = 0& To (lngRels - 1&)
11180                   If arr_varRel(R_RELS, lngX) > 0& And arr_varRel(R_RACTN, lngX) <> "{Remove}" Then
11190                     .FindFirst "[accountno] = '" & arr_varRel(R_ACTN, lngX) & "'"
11200                     If .NoMatch = False Then
11210                       .Edit
11220                       ![related_accountno] = arr_varRel(R_RACTN, lngX)
11230                       .Update
11240                     Else
                            ' ** Since we pulled this from the first walk through, it should always be found!
11250                     End If
11260                   Else
11270                     .FindFirst "[accountno] = '" & arr_varRel(R_ACTN, lngX) & "'"
11280                     If .NoMatch = False Then
11290                       .Edit
11300                       ![related_accountno] = Null
11310                       .Update
11320                     End If
11330                   End If
11340                 Next
11350                 .Close
11360               End With  ' ** rst.

11370               blnUpdateArchive = False
11380               varTmp00 = DCount("[journalno]", "LedgerArchive")
11390               If IsNull(varTmp00) = False Then
11400                 If varTmp00 > 0 Then

                        ' ** qryCompanyInformation_09 (LedgerArchive, with TypePfx, actno_len), grouped by TypePfx, with cnt.
11410                   Set qdf = .QueryDefs("qryCompanyInformation_09a")
11420                   Set rst = qdf.OpenRecordset
11430                   With rst
11440                     If .BOF = True And .EOF = True Then
                            ' ** That's screwy! It's gotta be one or the other!
11450                     Else
11460                       .MoveLast
11470                       If .RecordCount > 1 Then
                              ' ** This means there are some with and some without a prefix.
11480                         blnUpdateArchive = True
11490                       Else
11500                         If ![TypePfx] <> blnTmp_AccountNoWithType Then
11510                           blnUpdateArchive = True
11520                         End If
11530                       End If
11540                     End If
11550                     .Close
11560                   End With

11570                   If blnUpdateArchive = True Then

                          ' ** Empty tmpCompanyInformation.
11580                     Set qdf = .QueryDefs("qryCompanyInformation_20")
11590                     qdf.Execute

                          ' ** Append qryCompanyInformation_12 (qryCompanyInformation_10
                          ' ** (qryCompanyInformation_09 (LedgerArchive, with TypePfx, actno_len), with accountno_type,
                          ' ** accountno_notype), linked to qryCompanyInformation_11 (Account, with accountno_type,
                          ' ** accountno_notype), with accountno_new) to tmpCompanyInformation.
11600                     Set qdf = .QueryDefs("qryCompanyInformation_12a")  '#CRTC01
11610                     qdf.Execute

                          ' ** Update qryCompanyInformation_14 (LedgerArchive, with DLookups() to qryCompanyInformation_13
                          ' ** (tmpCompanyInformation, grouped by accountno, just those in need of change)).
11620                     Set qdf = .QueryDefs("qryCompanyInformation_15")
11630                     qdf.Execute

11640                   End If  ' ** blnUpdateArchive

11650                 End If  ' ** varTmp00.
11660               End If  ' ** DCount().

11670             End If  ' ** blnContinue.

11680           End If  ' ** lngHasRels.
11690         End If  ' ** blnContinue.

11700         .Close
11710       End With  ' ** dbs.

11720       If blnContinue = True Then
11730         If lngAcctsChanged = 1& Then
11740           strTmp01 = "One account number has been updated."
11750           strTmp02 = Left("Update Successful" & Space(50), 50)
11760         ElseIf lngAcctsChanged > 1& Then
11770           strTmp01 = CStr(lngAcctsChanged) & " account numbers have been updated."
11780           strTmp02 = Left("Update Successful" & Space(50), 50)
11790         Else
                ' ** I'm not sure how it would get here, but perhaps if it were being forced, and no changes needed to be made...
11800           strTmp01 = "No account numbers needed updating."
11810           strTmp02 = Left("Confirmation Successful" & Space(50), 50)
11820         End If
11830         DoCmd.Hourglass False
11840         Beep
11850         MsgBox strTmp01, vbInformation + vbOKOnly, strTmp02
11860       End If  ' ** blnContinue.

11870     Else
11880       DoCmd.Hourglass False
11890       MsgBox "Your changes have not been saved.", vbCritical + vbOKOnly, "Save Failed"
11900     End If  ' ** CoOptions_Write().

11910   End With

11920   DoCmd.Hourglass False

EXITP:
11930   Set rst = Nothing
11940   Set qdf = Nothing
11950   Set dbs = Nothing
11960   AccountNoWithType_Set = blnRetVal
11970   Exit Function

ERRH:
11980   DoCmd.Hourglass False
11990   blnRetVal = False
12000   Select Case ERR.Number
        Case Else
12010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12020   End Select
12030   Resume EXITP

End Function
