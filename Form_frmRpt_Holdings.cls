VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Holdings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Holdings"

'VGC 04/16/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, sets options, then stops.

' ** Shortcut Alt keys responsive from this form:
' **   Specified:        A {opgAssetNumber_optSpecified}
' **   Close:            C {cmdClose}
' **   Detail:           D {chkDetail}
' **   Excel:            E {cmdExcel}
' **   By Cusip:         I {opgAssetSource_optCusip}
' **   All:              L {opgAssetNumber_optAll}
' **   By Name:          N {opgAssetSource_optName}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Select Asset:     S {cmbAssets}
' **   Total Mkt Val:    T {chkTotalMarketValue}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}
' **   By Type:          Y {opgAssetSource_optType}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

Private blnGoingToReport As Boolean, blnWasGTR As Boolean, blnIsOpen As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           Select Case gblnGoToReport
              Case True
160             blnWasGTR = True
170           Case False
180             blnWasGTR = False
190           End Select

200           .IsWord = False
210           .IsExcel = False

220           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
230           .cmdExcel.Enabled = False
      #Else
240           .cmdExcel.Enabled = True
      #End If

              ' ** EVENT CHECK: chkRememberAsset!
              ' ** NONE!

250           opgAssetNumber_AfterUpdate  ' ** Procedure: Below.
260           chkDetail_AfterUpdate  ' ** Procedure: Below.
270           chkRememberAsset_AfterUpdate  ' ** Procedure: Below.
280           chkPageOf_AfterUpdate  ' ** Procedure: Below.
290           blnIsOpen = True
300           chkTotalMarketValue_AfterUpdate  ' ** Procedure: Below.
310           blnIsOpen = False

320         Else
330           Cancel = -1
340         End If
350       Else
360         Cancel = -1
370       End If
380     End With

390     DoCmd.Hourglass False

400     If Cancel = -1 Then
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
        ' **   By Cusip:         I {opgAssetSource_optCusip}
        ' **   By Name:          N {opgAssetSource_optName}
        ' **   By Type:          Y {opgAssetSource_optType}

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
          Case vbKeyI
630         With Me
640           intRetVal = 0
650           If .opgAssetSource.Enabled = True Then
660             .opgAssetSource.SetFocus
670             .opgAssetSource = .opgAssetSource_optCusip.OptionValue
680             opgAssetSource_AfterUpdate  ' ** Procedure: Below.
690           Else
700             Beep
710           End If
720         End With
730       Case vbKeyN
740         With Me
750           intRetVal = 0
760           If .opgAssetSource.Enabled = True Then
770             .opgAssetSource.SetFocus
780             .opgAssetSource = .opgAssetSource_optName.OptionValue
790             opgAssetSource_AfterUpdate  ' ** Procedure: Below.
800           Else
810             Beep
820           End If
830         End With
840       Case vbKeyX
850         intRetVal = 0
860         cmdClose_Click  ' ** Procedure: Below.
870       Case vbKeyY
880         With Me
890           intRetVal = 0
900           If .opgAssetSource.Enabled = True Then
910             .opgAssetSource.SetFocus
920             .opgAssetSource = .opgAssetSource_optType.OptionValue
930             opgAssetSource_AfterUpdate  ' ** Procedure: Below.
940           Else
950             Beep
960           End If
970         End With
980       End Select
990     End If

        ' ** Ctrl keys.
1000    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1010      Select Case intRetVal
          Case vbKeyG
1020        With Me
1030          intRetVal = 0
1040          If .chkPageOf.Enabled = True Then
1050            .chkPageOf.SetFocus
1060            .chkPageOf = (Not .chkPageOf)
1070            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1080          End If
1090        End With
1100      End Select
1110    End If

        ' ** Ctrl-Shift keys.
1120    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1130      Select Case intRetVal
          Case vbKeyF
1140        With Me
1150          intRetVal = 0
1160          .FocusHolder.SetFocus
1170        End With
1180      End Select
1190    End If

EXITP:
1200    KeyCode = intRetVal
1210    Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1310    With Me
1320      .TimerInterval = 0&
1330      If gblnSetFocus = True Then
1340        gblnSetFocus = False
1350        strFormName = vbNullString
1360  On Error Resume Next
1370        Do Until Screen.ActiveForm.Name = strFormName
1380          DoCmd.SelectObject acForm, THIS_NAME, False
1390          strFormName = Screen.ActiveForm.Name
1400        Loop
1410  On Error GoTo ERRH
1420      End If
1430      If gblnGoToReport = True Then
1440        Select Case blnGoingToReport
            Case True
1450          .GoToReport_arw_det_img.Visible = False
1460          If .cmdPrint.Enabled = True Then
1470            .cmdPrint.SetFocus
1480          Else
1490            .cmdClose.SetFocus
1500          End If
1510          DoEvents
1520          blnGoingToReport = False
1530          gblnGoToReport = False
1540          Beep
1550          DoCmd.Hourglass False
1560          DoEvents
1570        Case False
1580          DoCmd.Hourglass True  ' ** Make sure it's still runing.
1590          DoEvents
1600          If garr_varGoToReport(GTR_RNAM) = "rptHoldings_Detailed" Then
1610            blnGoingToReport = True
1620            .GoToReport_arw_det_img.Visible = True
1630            DoEvents
1640            If .chkDetail = False Then
1650              .chkDetail = True
1660              chkDetail_AfterUpdate  ' ** Procedure: Below.
1670              DoEvents
1680            End If
1690            .TimerInterval = (GTR_WAIT * 2&)
1700          Else
1710            If .chkDetail = True Then
1720              .chkDetail = False
1730              chkDetail_AfterUpdate  ' ** Procedure: Below.
1740              DoEvents
1750            End If
1760            If .cmdPrint.Enabled = True Then
1770              .cmdPrint.SetFocus
1780            Else
1790              .cmdClose.SetFocus
1800            End If
1810            DoEvents
1820            blnGoingToReport = False
1830            gblnGoToReport = False
1840            Beep
1850            DoCmd.Hourglass False
1860            DoEvents
1870          End If
1880        End Select
1890      End If
1900    End With

EXITP:
1910    Exit Sub

ERRH:
1920    DoCmd.Hourglass False
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2010    With Me
2020      If .chkRememberAsset = False Then
2030        .cmbAssets = Null  ' ** So it won't be there when Pref_Save() runs.
2040      End If
2050    End With

        ' ** Make sure all reports are closed.
2060    Do While Reports.Count > 0
2070      DoCmd.Close acReport, Reports(0).Name
2080    Loop

2090    Select Case blnWasGTR
        Case True
2100      blnWasGTR = False
2110    Case False
          ' ** Don't save settings imposed by GoToReport!
2120      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
2130    End Select

2140    gstrFormQuerySpec = vbNullString
2150    gblnSetFocus = True
2160    gstrReturningForm = THIS_NAME
2170    Select Case strCallingForm
        Case "frmMenu_Asset"
2180      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2190    End Select

EXITP:
2200    Exit Sub

ERRH:
2210    Select Case ERR.Number
        Case Else
2220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2230    End Select
2240    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2310    DoCmd.Close acForm, THIS_NAME

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
2490          .opgAssetNumber.SetFocus
2500        End With
2510      End Select
2520    End If

        ' ** Shift keys.
2530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2540      Select Case intRetVal
          Case vbKeyTab
2550        With Me
2560          intRetVal = 0
2570          If .cmdExcel.Enabled = True Then
2580            .cmdExcel.SetFocus
2590          Else
2600            .cmdWord.SetFocus
2610          End If
2620        End With
2630      End Select
2640    End If

EXITP:
2650    KeyCode = intRetVal
2660    Exit Sub

ERRH:
2670    intRetVal = 0
2680    Select Case ERR.Number
        Case Else
2690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2700    End Select
2710    Resume EXITP

End Sub

Private Sub opgAssetNumber_AfterUpdate()

2800  On Error GoTo ERRH:

        Const THIS_PROC As String = "opgAssetNumber_AfterUpdate"

2810    With Me
2820      Select Case .opgAssetNumber
          Case .opgAssetNumber_optSpecified.OptionValue
2830        .opgAssetNumber_optSpecified_lbl.FontBold = True
2840        .opgAssetNumber_optAll_lbl.FontBold = False
2850        .opgAssetNumber_optSpecified_lbl_box.Visible = True
2860        .opgAssetNumber_optAll_lbl_box.Visible = False
2870        .cmbAssets.Enabled = True
2880        .cmbAssets.BorderColor = CLR_LTBLU2
2890        .cmbAssets.BackStyle = acBackStyleNormal
2900        .chkRememberAsset.Enabled = True
2910        .chkRememberAsset_lbl.Visible = True
2920        .chkRememberAsset_lbl2_dim.Visible = False
2930        .chkRememberAsset_lbl2_dim_hi.Visible = False
2940        .opgAssetSource.Enabled = True
2950        .opgAssetSource_optType_lbl2.ForeColor = CLR_VDKGRY
2960        .opgAssetSource_optType_lbl2_dim_hi.Visible = False
2970        .opgAssetSource_optName_lbl2.ForeColor = CLR_VDKGRY
2980        .opgAssetSource_optName_lbl2_dim_hi.Visible = False
2990        .opgAssetSource_optCusip_lbl2.ForeColor = CLR_VDKGRY
3000        .opgAssetSource_optCusip_lbl2_dim_hi.Visible = False
3010        .chkTotalMarketValue.Enabled = True
3020        .chkTotalMarketValue_lbl2.ForeColor = CLR_BLK
3030        .chkTotalMarketValue_lbl2_dim_hi.Visible = False
3040        If .chkTotalMarketValue = True Then
3050          .TotalMarketValue.Enabled = True
3060          .TotalMarketValue.BorderColor = CLR_LTBLU2
3070          .TotalMarketValue.BackStyle = acBackStyleNormal
3080        End If
3090      Case .opgAssetNumber_optAll.OptionValue
3100        .opgAssetNumber_optSpecified_lbl.FontBold = False
3110        .opgAssetNumber_optAll_lbl.FontBold = True
3120        .opgAssetNumber_optSpecified_lbl_box.Visible = False
3130        .opgAssetNumber_optAll_lbl_box.Visible = True
3140        .cmbAssets.Enabled = False
3150        .cmbAssets.BorderColor = WIN_CLR_DISR
3160        .cmbAssets.BackStyle = acBackStyleTransparent
3170        .chkRememberAsset.Enabled = False
3180        .chkRememberAsset_lbl.Visible = False
3190        .chkRememberAsset_lbl2_dim.Visible = True
3200        .chkRememberAsset_lbl2_dim_hi.Visible = True
3210        .opgAssetSource.Enabled = False
3220        .opgAssetSource_optType_lbl2.ForeColor = WIN_CLR_DISF
3230        .opgAssetSource_optType_lbl2_dim_hi.Visible = True
3240        .opgAssetSource_optName_lbl2.ForeColor = WIN_CLR_DISF
3250        .opgAssetSource_optName_lbl2_dim_hi.Visible = True
3260        .opgAssetSource_optCusip_lbl2.ForeColor = WIN_CLR_DISF
3270        .opgAssetSource_optCusip_lbl2_dim_hi.Visible = True
3280        .chkTotalMarketValue.Enabled = False
3290        .chkTotalMarketValue_lbl2.ForeColor = WIN_CLR_DISF
3300        .chkTotalMarketValue_lbl2_dim_hi.Visible = True
3310        If .chkTotalMarketValue = True Then
3320          .TotalMarketValue.Enabled = False
3330          .TotalMarketValue.BorderColor = WIN_CLR_DISR
3340          .TotalMarketValue.BackStyle = acBackStyleTransparent
3350        End If
3360      End Select
3370    End With

EXITP:
3380    Exit Sub

ERRH:
3390    Select Case ERR.Number
        Case Else
3400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3410    End Select
3420    Resume EXITP

End Sub

Private Sub opgAssetNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetNumber_optSpecified_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
3570        With Me
3580          intRetVal = 0
3590          If .cmbAssets.Enabled = True Then
3600            .cmbAssets.SetFocus
3610          Else
3620            .cmdPreview.SetFocus
3630          End If
3640        End With
3650      End Select
3660    End If

        ' ** Shift keys.
3670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3680      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3690        With Me
3700          intRetVal = 0
3710          .cmdClose.SetFocus
3720        End With
3730      End Select
3740    End If

        ' ** Ctrl keys.
3750    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3770        With Me
3780          intRetVal = 0
3790          .chkDetail.SetFocus
3800        End With
3810      End Select
3820    End If

EXITP:
3830    KeyCode = intRetVal
3840    Exit Sub

ERRH:
3850    intRetVal = 0
3860    Select Case ERR.Number
        Case Else
3870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3880    End Select
3890    Resume EXITP

End Sub

Private Sub opgAssetNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetNumber_optAll_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
3970        With Me
3980          intRetVal = 0
3990          If .cmbAssets.Enabled = True Then
4000            .cmbAssets.SetFocus
4010          Else
4020            .cmdPreview.SetFocus
4030          End If
4040        End With
4050      End Select
4060    End If

        ' ** Shift keys.
4070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4080      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4090        With Me
4100          intRetVal = 0
4110          .cmdClose.SetFocus
4120        End With
4130      End Select
4140    End If

        ' ** Ctrl keys.
4150    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4170        With Me
4180          intRetVal = 0
4190          .chkDetail.SetFocus
4200        End With
4210      End Select
4220    End If

EXITP:
4230    KeyCode = intRetVal
4240    Exit Sub

ERRH:
4250    intRetVal = 0
4260    Select Case ERR.Number
        Case Else
4270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4280    End Select
4290    Resume EXITP

End Sub

Public Sub chkDetail_AfterUpdate()
' ** Public for GoToReport.

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkDetail_AfterUpdate"

4310    With Me
4320      Select Case .chkDetail
          Case True
4330        .chkDetail_lbl.FontBold = True
4340      Case False
4350        .chkDetail_lbl.FontBold = False
4360      End Select
4370    End With

EXITP:
4380    Exit Sub

ERRH:
4390    Select Case ERR.Number
        Case Else
4400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4410    End Select
4420    Resume EXITP

End Sub

Private Sub chkDetail_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkDetail_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4520    intShiftDown = (Shift And acShiftMask) > 0
4530    intAltDown = (Shift And acAltMask) > 0
4540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4570        With Me
4580          intRetVal = 0
4590          .chkTotalMarketValue.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4650        With Me
4660          intRetVal = 0
4670          .opgAssetNumber.SetFocus
4680        End With
4690      End Select
4700    End If

EXITP:
4710    KeyCode = intRetVal
4720    Exit Sub

ERRH:
4730    intRetVal = 0
4740    Select Case ERR.Number
        Case Else
4750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4760    End Select
4770    Resume EXITP

End Sub

Private Sub chkTotalMarketValue_AfterUpdate()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTotalMarketValue_AfterUpdate"

4810    With Me
4820      Select Case .chkTotalMarketValue
          Case True
4830        .chkTotalMarketValue_lbl.FontBold = True
4840        .chkTotalMarketValue_lbl2.FontBold = True
4850        .chkTotalMarketValue_lbl2_dim_hi.FontBold = True
4860        .TotalMarketValue.Enabled = True
            '.TotalMarketValue_lbl_box.Visible = False
            '.TotalMarketValue_lbl.BackStyle = acBackStyleNormal
4870        .TotalMarketValue.BorderColor = CLR_LTBLU2
4880        .TotalMarketValue.BackStyle = acBackStyleNormal
4890        If blnIsOpen = False Then
4900          .TotalMarketValue.SetFocus
4910        End If
4920      Case False
4930        .chkTotalMarketValue_lbl.FontBold = False
4940        .chkTotalMarketValue_lbl2.FontBold = False
4950        .chkTotalMarketValue_lbl2_dim_hi.FontBold = False
4960        .TotalMarketValue.Enabled = False
            '.TotalMarketValue_lbl_box.Visible = True
            '.TotalMarketValue_lbl.BackStyle = acBackStyleTransparent
4970        .TotalMarketValue.BorderColor = WIN_CLR_DISR
4980        .TotalMarketValue.BackStyle = acBackStyleTransparent
4990      End Select
5000    End With

EXITP:
5010    Exit Sub

ERRH:
5020    Select Case ERR.Number
        Case Else
5030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5040    End Select
5050    Resume EXITP

End Sub

Private Sub chkTotalMarketValue_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTotalMarketValue_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5120    intShiftDown = (Shift And acShiftMask) > 0
5130    intAltDown = (Shift And acAltMask) > 0
5140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5170        With Me
5180          intRetVal = 0
5190          Select Case .chkTotalMarketValue
              Case True
5200            .TotalMarketValue.SetFocus
5210          Case False
5220            If .opgAssetSource.Enabled = True Then
5230              .opgAssetSource.SetFocus
5240            Else
5250              .chkPageOf.SetFocus
5260            End If
5270          End Select
5280        End With
5290      End Select
5300    End If

        ' ** Shift keys.
5310    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5320      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5330        With Me
5340          intRetVal = 0
5350          .chkDetail.SetFocus
5360        End With
5370      End Select
5380    End If

EXITP:
5390    KeyCode = intRetVal
5400    Exit Sub

ERRH:
5410    intRetVal = 0
5420    Select Case ERR.Number
        Case Else
5430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5440    End Select
5450    Resume EXITP

End Sub

Private Sub TotalMarketValue_KeyDown(KeyCode As Integer, Shift As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "TotalMarketValue_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5520    intShiftDown = (Shift And acShiftMask) > 0
5530    intAltDown = (Shift And acAltMask) > 0
5540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5570        With Me
5580          intRetVal = 0
5590          .cmbAssets.SetFocus
5600        End With
5610      End Select
5620    End If

        ' ** Shift keys.
5630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5650        With Me
5660          intRetVal = 0
5670          .chkTotalMarketValue.SetFocus
5680        End With
5690      End Select
5700    End If

EXITP:
5710    KeyCode = intRetVal
5720    Exit Sub

ERRH:
5730    intRetVal = 0
5740    Select Case ERR.Number
        Case Else
5750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5760    End Select
5770    Resume EXITP

End Sub

Private Sub opgAssetSource_AfterUpdate()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_AfterUpdate"

5810    With Me
5820      Select Case .opgAssetSource
          Case .opgAssetSource_optType.OptionValue
5830        .opgAssetSource_optType_lbl2.FontBold = True
5840        .opgAssetSource_optType_lbl2_dim_hi.FontBold = True
5850        .opgAssetSource_optName_lbl2.FontBold = False
5860        .opgAssetSource_optName_lbl2_dim_hi.FontBold = False
5870        .opgAssetSource_optCusip_lbl2.FontBold = False
5880        .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = False
5890        If .cmbAssets.RowSource <> "qryHoldings_03a" Then
5900          .cmbAssets.RowSource = "qryHoldings_03a"
5910        End If
5920        .cmbAssets.ListWidth = 6690&
5930        .cmbAssets.ColumnWidths = "0;4875;1200;360;0"
5940      Case .opgAssetSource_optName.OptionValue
5950        .opgAssetSource_optType_lbl2.FontBold = False
5960        .opgAssetSource_optType_lbl2_dim_hi.FontBold = False
5970        .opgAssetSource_optName_lbl2.FontBold = True
5980        .opgAssetSource_optName_lbl2_dim_hi.FontBold = True
5990        .opgAssetSource_optCusip_lbl2.FontBold = False
6000        .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = False
6010        If .cmbAssets.RowSource <> "qryHoldings_03b" Then
6020          .cmbAssets.RowSource = "qryHoldings_03b"
6030        End If
6040        .cmbAssets.ListWidth = 6690&
6050        .cmbAssets.ColumnWidths = "0;4875;1200;360;0"
6060      Case .opgAssetSource_optCusip.OptionValue
6070        .opgAssetSource_optType_lbl2.FontBold = False
6080        .opgAssetSource_optType_lbl2_dim_hi.FontBold = False
6090        .opgAssetSource_optName_lbl2.FontBold = False
6100        .opgAssetSource_optName_lbl2_dim_hi.FontBold = False
6110        .opgAssetSource_optCusip_lbl2.FontBold = True
6120        .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = True
6130        If .cmbAssets.RowSource <> "qryHoldings_03c" Then
6140          .cmbAssets.RowSource = "qryHoldings_03c"
6150        End If
6160        .cmbAssets.ListWidth = 6690&
6170        .cmbAssets.ColumnWidths = "0;6075;0;360;0"
6180      End Select
6190    End With

EXITP:
6200    Exit Sub

ERRH:
6210    Select Case ERR.Number
        Case Else
6220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6230    End Select
6240    Resume EXITP

End Sub

Private Sub opgAssetSource_optType_KeyDown(KeyCode As Integer, Shift As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6320    intShiftDown = (Shift And acShiftMask) > 0
6330    intAltDown = (Shift And acAltMask) > 0
6340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6370        With Me
6380          intRetVal = 0
6390          .chkRememberAsset.SetFocus
6400        End With
6410      End Select
6420    End If

        ' ** Shift keys.
6430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6450        With Me
6460          intRetVal = 0
6470          If .TotalMarketValue.Enabled = True Then
6480            .TotalMarketValue.SetFocus
6490          Else
6500            .chkTotalMarketValue.SetFocus
6510          End If
6520        End With
6530      End Select
6540    End If

EXITP:
6550    KeyCode = intRetVal
6560    Exit Sub

ERRH:
6570    intRetVal = 0
6580    Select Case ERR.Number
        Case Else
6590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6600    End Select
6610    Resume EXITP

End Sub

Private Sub opgAssetSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optName_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
6770        With Me
6780          intRetVal = 0
6790          .chkRememberAsset.SetFocus
6800        End With
6810      End Select
6820    End If

        ' ** Shift keys.
6830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6850        With Me
6860          intRetVal = 0
6870          If .TotalMarketValue.Enabled = True Then
6880            .TotalMarketValue.SetFocus
6890          Else
6900            .chkTotalMarketValue.SetFocus
6910          End If
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

Private Sub opgAssetSource_optCusip_KeyDown(KeyCode As Integer, Shift As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optCusip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7120    intShiftDown = (Shift And acShiftMask) > 0
7130    intAltDown = (Shift And acAltMask) > 0
7140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7170        With Me
7180          intRetVal = 0
7190          .chkRememberAsset.SetFocus
7200        End With
7210      End Select
7220    End If

        ' ** Shift keys.
7230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7250        With Me
7260          intRetVal = 0
7270          If .TotalMarketValue.Enabled = True Then
7280            .TotalMarketValue.SetFocus
7290          Else
7300            .chkTotalMarketValue.SetFocus
7310          End If
7320        End With
7330      End Select
7340    End If

EXITP:
7350    KeyCode = intRetVal
7360    Exit Sub

ERRH:
7370    intRetVal = 0
7380    Select Case ERR.Number
        Case Else
7390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7400    End Select
7410    Resume EXITP

End Sub

Private Sub chkRememberAsset_AfterUpdate()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberAsset_AfterUpdate"

7510    With Me
7520      Select Case .chkRememberAsset
          Case True
7530        .chkRememberAsset_lbl.FontBold = True
7540        .chkRememberAsset_lbl2_dim.FontBold = True
7550        .chkRememberAsset_lbl2_dim_hi.FontBold = True
7560      Case False
7570        .chkRememberAsset_lbl.FontBold = False
7580        .chkRememberAsset_lbl2_dim.FontBold = False
7590        .chkRememberAsset_lbl2_dim_hi.FontBold = False
7600      End Select
7610    End With

EXITP:
7620    Exit Sub

ERRH:
7630    Select Case ERR.Number
        Case Else
7640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7650    End Select
7660    Resume EXITP

End Sub

Private Sub chkRememberAsset_KeyDown(KeyCode As Integer, Shift As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberAsset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7720    intShiftDown = (Shift And acShiftMask) > 0
7730    intAltDown = (Shift And acAltMask) > 0
7740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7770        With Me
7780          intRetVal = 0
7790          .cmbAssets.SetFocus
7800        End With
7810      End Select
7820    End If

        ' ** Shift keys.
7830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7850        With Me
7860          intRetVal = 0
7870          .opgAssetSource.SetFocus
7880        End With
7890      End Select
7900    End If

EXITP:
7910    KeyCode = intRetVal
7920    Exit Sub

ERRH:
7930    intRetVal = 0
7940    Select Case ERR.Number
        Case Else
7950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7960    End Select
7970    Resume EXITP

End Sub

Private Sub cmbAssets_Enter()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_Enter"

8010    With Me
8020      If .cmbAssets.text = vbNullString Then
8030        ExpandCombo .cmbAssets  ' ** Module Procedure: modUtilities.
8040        .cmbAssets = .cmbAssets.ItemData(0)
8050      End If
8060    End With

EXITP:
8070    Exit Sub

ERRH:
8080    Select Case ERR.Number
        Case Else
8090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8100    End Select
8110    Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8220    intShiftDown = (Shift And acShiftMask) > 0
8230    intAltDown = (Shift And acAltMask) > 0
8240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8270        With Me
8280          intRetVal = 0
8290          .cmdPreview.SetFocus
8300        End With
8310      End Select
8320    End If

        ' ** Shift keys.
8330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8350        With Me
8360          intRetVal = 0
8370          .opgAssetNumber.SetFocus
8380        End With
8390      End Select
8400    End If

        ' ** Ctrl keys.
8410    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8420      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8430        With Me
8440          intRetVal = 0
8450          .chkPageOf.SetFocus
8460        End With
8470      End Select
8480    End If

        ' ** Ctrl-Shift keys.
8490    If intCtrlDown And (Not intAltDown) And intShiftDown Then
8500      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8510        With Me
8520          intRetVal = 0
8530          .chkRememberAsset.SetFocus
8540        End With
8550      End Select
8560    End If

EXITP:
8570    KeyCode = intRetVal
8580    Exit Sub

ERRH:
8590    intRetVal = 0
8600    Select Case ERR.Number
        Case Else
8610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8620    End Select
8630    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

8710    With Me
8720      Select Case .chkPageOf
          Case True
8730        .chkPageOf_lbl.FontBold = True
8740      Case False
8750        .chkPageOf_lbl.FontBold = False
8760      End Select
8770    End With

EXITP:
8780    Exit Sub

ERRH:
8790    DoCmd.Hourglass False
8800    Select Case ERR.Number
        Case Else
8810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8820    End Select
8830    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8920    intShiftDown = (Shift And acShiftMask) > 0
8930    intAltDown = (Shift And acAltMask) > 0
8940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8970        With Me
8980          intRetVal = 0
8990          .cmdPreview.SetFocus
9000        End With
9010      End Select
9020    End If

        ' ** Shift keys.
9030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9050        With Me
9060          intRetVal = 0
9070          .cmbAssets.SetFocus
9080        End With
9090      End Select
9100    End If

EXITP:
9110    KeyCode = intRetVal
9120    Exit Sub

ERRH:
9130    intRetVal = 0
9140    Select Case ERR.Number
        Case Else
9150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9160    End Select
9170    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

9210    With Me
9220      DoCmd.Hourglass True
9230      DoEvents
9240      .IsWord = False
9250      .IsExcel = False
9260      If DoReport = True Then  ' ** Function: Below.
9270        Select Case .chkDetail
            Case True
9280          strDocName = "rptHoldings_Detailed"
9290        Case False
9300          strDocName = "rptHoldings"
9310        End Select
9320        Select Case .opgAssetNumber
            Case .opgAssetNumber_optSpecified.OptionValue
9330          DoCmd.OpenReport strDocName, acViewPreview, , "[assetno] = " & CStr(.cmbAssets)
9340        Case .opgAssetNumber_optAll.OptionValue
9350          DoCmd.OpenReport strDocName, acViewPreview, , vbNullString
9360        End Select
9370        DoCmd.Maximize
9380        DoCmd.RunCommand acCmdFitToWindow
9390      End If
9400      DoCmd.Hourglass False
9410    End With

EXITP:
9420    Exit Sub

ERRH:
9430    DoCmd.Hourglass False
9440    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
9450    Case Else
9460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9470    End Select
9480    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9520    intShiftDown = (Shift And acShiftMask) > 0
9530    intAltDown = (Shift And acAltMask) > 0
9540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9560      Select Case intRetVal
          Case vbKeyTab
9570        With Me
9580          intRetVal = 0
9590          .cmdPrint.SetFocus
9600        End With
9610      End Select
9620    End If

        ' ** Shift keys.
9630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9640      Select Case intRetVal
          Case vbKeyTab
9650        With Me
9660          intRetVal = 0
9670          If .cmbAssets.Enabled = True Then
9680            .cmbAssets.SetFocus
9690          Else
9700            .opgAssetNumber.SetFocus
9710          End If
9720        End With
9730      End Select
9740    End If

        ' ** Ctrl-Shift keys.
9750    If intCtrlDown And (Not intAltDown) And intShiftDown Then
9760      Select Case intRetVal
          Case vbKeyTab
9770        With Me
9780          intRetVal = 0
9790          .chkPageOf.SetFocus
9800        End With
9810      End Select
9820    End If

EXITP:
9830    KeyCode = intRetVal
9840    Exit Sub

ERRH:
9850    intRetVal = 0
9860    Select Case ERR.Number
        Case Else
9870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9880    End Select
9890    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String, intMode As Integer

9910    With Me
9920      DoCmd.Hourglass True
9930      DoEvents
9940      .IsWord = False
9950      .IsExcel = False
9960      If DoReport = True Then  ' ** Function: Below.
9970        Select Case .chkDetail
            Case True
9980          strDocName = "rptHoldings_Detailed"
9990        Case False
10000         strDocName = "rptHoldings"
10010       End Select
10020       If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
10030         intMode = acViewPreview
10040       Else
10050         intMode = acViewNormal
10060       End If
10070       Select Case .opgAssetNumber
            Case .opgAssetNumber_optSpecified.OptionValue
              '##GTR_Ref: rptHoldings
              '##GTR_Ref: rptHoldings_Detailed
10080         DoCmd.OpenReport strDocName, intMode, , "[assetno] = " & CStr(.cmbAssets)
10090       Case .opgAssetNumber_optAll.OptionValue
10100         DoCmd.OpenReport strDocName, intMode, , vbNullString
10110       End Select
10120       If intMode = acViewPreview Then
10130         DoCmd.Maximize
10140         DoCmd.RunCommand acCmdFitToWindow
10150       End If
10160     End If
10170     DoCmd.Hourglass False
10180   End With

EXITP:
10190   Exit Sub

ERRH:
10200   DoCmd.Hourglass False
10210   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
10220   Case Else
10230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10240   End Select
10250   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10320   intShiftDown = (Shift And acShiftMask) > 0
10330   intAltDown = (Shift And acAltMask) > 0
10340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10360     Select Case intRetVal
          Case vbKeyTab
10370       With Me
10380         intRetVal = 0
10390         If .cmdWord.Enabled = True Then
10400           .cmdWord.SetFocus
10410         Else
10420           .cmdClose.SetFocus
10430         End If
10440       End With
10450     End Select
10460   End If

        ' ** Shift keys.
10470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10480     Select Case intRetVal
          Case vbKeyTab
10490       With Me
10500         intRetVal = 0
10510         .cmdPreview.SetFocus
10520       End With
10530     End Select
10540   End If

EXITP:
10550   KeyCode = intRetVal
10560   Exit Sub

ERRH:
10570   intRetVal = 0
10580   Select Case ERR.Number
        Case Else
10590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10600   End Select
10610   Resume EXITP

End Sub

Private Sub cmdWord_Click()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

10710   With Me

10720     DoCmd.Hourglass True
10730     DoEvents

10740     .IsWord = True
10750     .IsExcel = False

10760     If DoReport = True Then ' ** Function: Below.

10770       If IsNull(.UserReportPath) = True Then
10780         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
10790       Else
10800         strRptPath = .UserReportPath
10810       End If

10820       Select Case .opgAssetNumber
            Case .opgAssetNumber_optSpecified.OptionValue
10830         Select Case .chkDetail
              Case True
10840           strRptCap = "rptHoldings_Detailed_" & Format(Date, "yyyymmdd")
10850         Case False
10860           strRptCap = "rptHoldings_" & Format(Date, "yyyymmdd")
10870         End Select
10880       Case .opgAssetNumber_optAll.OptionValue
10890         Select Case .chkDetail
              Case True
10900           strRptCap = "rptHoldings_Detailed_All_" & Format(Date, "yyyymmdd")
10910         Case False
10920           strRptCap = "rptHoldings_All_" & Format(Date, "yyyymmdd")
10930         End Select
10940       End Select

10950       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

10960       If strRptPathFile <> vbNullString Then
10970         Select Case .chkDetail
              Case True
10980           strRpt = "rptHoldings_Detailed"
10990         Case False
11000           strRpt = "rptHoldings"
11010         End Select
              ' ** AssetNo is filtered in Report_Open().
11020         DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
11030         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
11040       End If

11050     End If
11060     DoCmd.Hourglass False
11070   End With

EXITP:
11080   Exit Sub

ERRH:
11090   DoCmd.Hourglass False
11100   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
11110   Case Else
11120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11130   End Select
11140   Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11220   intShiftDown = (Shift And acShiftMask) > 0
11230   intAltDown = (Shift And acAltMask) > 0
11240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11260     Select Case intRetVal
          Case vbKeyTab
11270       With Me
11280         intRetVal = 0
11290         If .cmdExcel.Enabled = True Then
11300           .cmdExcel.SetFocus
11310         Else
11320           .cmdClose.SetFocus
11330         End If
11340       End With
11350     End Select
11360   End If

        ' ** Shift keys.
11370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11380     Select Case intRetVal
          Case vbKeyTab
11390       With Me
11400         intRetVal = 0
11410         .cmdPrint.SetFocus
11420       End With
11430     End Select
11440   End If

EXITP:
11450   KeyCode = intRetVal
11460   Exit Sub

ERRH:
11470   intRetVal = 0
11480   Select Case ERR.Number
        Case Else
11490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11500   End Select
11510   Resume EXITP

End Sub

Private Sub cmdExcel_Click()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String, strMacro As String, strExcelFile As String
        Dim blnAll As Boolean, blnFound As Boolean
        Dim strTmp01 As String

11610   With Me

11620     DoCmd.Hourglass True
11630     DoEvents

11640     .IsWord = False
11650     .IsExcel = True

11660     If DoReport = True Then  ' ** Function: Below.

11670       If IsNull(.UserReportPath) = True Then
11680         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
11690       Else
11700         strRptPath = .UserReportPath
11710       End If

11720       Select Case .opgAssetNumber
            Case .opgAssetNumber_optSpecified.OptionValue
11730         blnAll = False
11740         Select Case .chkDetail
              Case True
11750           strRptCap = "rptHoldings_Detailed_" & Format(Date, "yyyymmdd")
11760         Case False
11770           strRptCap = "rptHoldings_" & Format(Date, "yyyymmdd")
11780         End Select
11790       Case .opgAssetNumber_optAll.OptionValue
11800         blnAll = True
11810         Select Case .chkDetail
              Case True
11820           strRptCap = "rptHoldings_Detailed_All_" & Format(Date, "yyyymmdd")
11830         Case False
11840           strRptCap = "rptHoldings_All_" & Format(Date, "yyyymmdd")
11850         End Select
11860       End Select

11870       DoCmd.Hourglass False
11880       strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

11890       If strRptPathFile <> vbNullString Then
11900         DoCmd.Hourglass True
11910         DoEvents
11920         If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
11930           EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
11940         End If
11950         DoEvents
11960         If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
11970           Kill strRptPathFile
11980         End If
11990         DoEvents
12000         gstrFormQuerySpec = THIS_NAME
12010         Select Case .opgAssetNumber
              Case .opgAssetNumber_optSpecified.OptionValue
12020           Select Case .chkDetail
                Case True
                  ' ** qryRpt_Holdings_Detailed_04_14 (xx), linked to qryRpt_Holdings_Detailed_04_15 (xx), one, with detail; For Export.
12030             strQry = "qryRpt_Holdings_Detailed_04_16"
12040           Case False
                  ' ** qryRpt_Holdings_04_11 (xx), one, no detail; For Export.
12050             strQry = "qryRpt_Holdings_04_12"
12060           End Select
12070         Case .opgAssetNumber_optAll.OptionValue
12080           Select Case .chkDetail
                Case True
                  ' ** qryRpt_Holdings_Detailed_03_14 (xx), linked to qryRpt_Holdings_Detailed_03_15 (xx), all, with detail; For Export.
12090             strQry = "qryRpt_Holdings_Detailed_03_16"
12100           Case False
                  ' ** qryRpt_Holdings_03_12 (xx), all, no detail; For Export.
12110             strQry = "qryRpt_Holdings_03_13"
12120           End Select
12130         End Select

              ' ** Now set up the macro.
12140         Select Case .opgAssetNumber
              Case .opgAssetNumber_optSpecified.OptionValue
12150           Select Case .chkDetail
                Case True
12160             strMacro = "mcrExcelExport_Holdings_One_Detail"
12170             strExcelFile = "Holdings_Detailed_xxx.xls"
12180           Case False
12190             strMacro = "mcrExcelExport_Holdings_One_NoDetail"
12200             strExcelFile = "Holdings_xxx.xls"
12210           End Select
12220         Case .opgAssetNumber_optAll.OptionValue
12230           Select Case .chkDetail
                Case True
12240             strMacro = "mcrExcelExport_Holdings_All_Detail"
12250             strExcelFile = "Holdings_Detailed_All_xxx.xls"
12260           Case False
12270             strMacro = "mcrExcelExport_Holdings_All_NoDetail"
12280             strExcelFile = "Holdings_All_xxx.xls"
12290           End Select
12300         End Select

              ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
              ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
12310         DoCmd.RunMacro strMacro
              'DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False

              ' ** The macros specify the queries, above, but cannot be given a dynamic file name.
              ' ** So, it's exported to 'Holdings_..._xxx.xls', which is then renamed.
12320         blnFound = False
12330         If FileExists(CurrentAppPath & LNK_SEP & strExcelFile) = True Then  ' ** Module Function: modFileUtilities.
12340           blnFound = True
12350           strTmp01 = CurrentAppPath & LNK_SEP & strExcelFile  ' ** Module Function: modFileUtilities.
12360         ElseIf FileExists(Parse_Path(strRptPathFile) & LNK_SEP & strExcelFile) = True Then   ' ** Module Function: modFileUtilities.
12370           blnFound = True
12380           strTmp01 = Parse_Path(strRptPathFile) & LNK_SEP & strExcelFile  ' ** Module Function: modFileUtilities.
12390         End If

12400         If blnFound = True Then
12410           Name (strTmp01) As (strRptPathFile)
12420           DoEvents
12430           Select Case .chkDetail
                Case True
                  ' ** Name, Trade Date, Cost, Location:
                  ' **   Specified, Detail.
                  ' **   All, Detail.
12440             If Excel_Holdings(strRptPathFile, blnAll) = True Then  ' ** Module Function: modExcelFuncs.
12450               If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
12460                 EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
12470               End If
12480               OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
12490             End If
12500           Case False
                  ' ** Name Only:
                  ' **   Specified, No Detail.
                  ' **   All, No Detail.
12510             strTmp01 = "Holdings"
12520             If blnAll = True Then strTmp01 = strTmp01 & " All"
12530             If Excel_NameOnly(strRptPathFile, strTmp01) = True Then  ' ** Module Function: modExcelFuncs.
12540               If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
12550                 EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
12560               End If
12570               OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
12580             End If
12590           End Select
12600           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
12610         End If  ' ** blnFound.

12620       End If  ' ** vbNullString.

            '="Holders of:  " & [Description] & " " & IIf([rate]>0," " & Format([rate],"#,##0.000%"),"") & " " & IIf([due] Is Not Null,"  Due " & Format([due],"mm/dd/yyyy"))

12630     End If
12640     DoCmd.Hourglass False
12650   End With

EXITP:
12660   Exit Sub

ERRH:
12670   DoCmd.Hourglass False
12680   Select Case ERR.Number
        Case Else
12690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12700   End Select
12710   Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12820   intShiftDown = (Shift And acShiftMask) > 0
12830   intAltDown = (Shift And acAltMask) > 0
12840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12860     Select Case intRetVal
          Case vbKeyTab
12870       With Me
12880         intRetVal = 0
12890         .cmdClose.SetFocus
12900       End With
12910     End Select
12920   End If

        ' ** Shift keys.
12930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12940     Select Case intRetVal
          Case vbKeyTab
12950       With Me
12960         intRetVal = 0
12970         If .cmdWord.Enabled = True Then
12980           .cmdWord.SetFocus
12990         Else
13000           .cmdPrint.SetFocus
13010         End If
13020       End With
13030     End Select
13040   End If

EXITP:
13050   KeyCode = intRetVal
13060   Exit Sub

ERRH:
13070   intRetVal = 0
13080   Select Case ERR.Number
        Case Else
13090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13100   End Select
13110   Resume EXITP

End Sub

Private Function DoReport() As Boolean

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

13210   blnRetVal = True

13220   With Me

13230     Select Case .opgAssetNumber
          Case .opgAssetNumber_optSpecified.OptionValue
13240       If IsNull(.cmbAssets) = True Then
13250         blnRetVal = False
13260         DoCmd.Hourglass False
13270         MsgBox "You must select an asset to continue.", vbInformation + vbOKOnly, "Entry Required"
13280       Else
13290         If .cmbAssets = vbNullString Then
13300           blnRetVal = False
13310           DoCmd.Hourglass False
13320           MsgBox "You must select an asset to continue.", vbInformation + vbOKOnly, "Entry Required"
13330         End If
13340       End If
13350     Case .opgAssetNumber_optAll.OptionValue
            ' ** Nothing required.
13360     End Select

13370     If blnRetVal = True Then

13380       gstrFormQuerySpec = THIS_NAME

13390       Set dbs = CurrentDb

13400       Select Case .opgAssetNumber
            Case .opgAssetNumber_optSpecified.OptionValue
13410         Select Case .chkDetail
              Case True
                ' ** qryRpt_Holdings_Detailed_02_01 (ActiveAssets, grouped and summed,
                ' ** by accountno, assetno, assetdate.), linked to Location, by specified FormRef('assetno').
13420           Set qdf = dbs.QueryDefs("qryRpt_Holdings_Detailed_02_02")
13430         Case False
                ' ** ActiveAssets, grouped and summed, by accountno, assetno, assetdate, by specified FormRef('assetno').
13440           Set qdf = dbs.QueryDefs("qryRpt_Holdings_02")
13450         End Select
13460       Case .opgAssetNumber_optAll.OptionValue
13470         Select Case .chkDetail
              Case True
                ' ** qryRpt_Holdings_Detailed_01_01 (ActiveAssets, grouped and summed,
                ' ** by accountno, assetno, assetdate), linked to Location.
13480           Set qdf = dbs.QueryDefs("qryRpt_Holdings_Detailed_01_02")
13490         Case False
                ' ** ActiveAssets, grouped and summed, by accountno, assetno, assetdate.
13500           Set qdf = dbs.QueryDefs("qryRpt_Holdings_01")
13510         End Select
13520       End Select

13530       Set rst = qdf.OpenRecordset
13540       With rst
13550         If .BOF = True And .EOF = True Then
13560           blnRetVal = False
13570           DoCmd.Hourglass False
13580           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
13590         End If
13600         .Close
13610       End With

13620       dbs.Close

13630     End If  ' ** blnRetVal.

13640   End With  ' ** Me.

EXITP:
13650   Set qdf = Nothing
13660   Set rst = Nothing
13670   Set dbs = Nothing
13680   DoReport = blnRetVal
13690   Exit Function

ERRH:
13700   blnRetVal = False
13710   DoCmd.Hourglass False
13720   Select Case ERR.Number
        Case Else
13730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13740   End Select
13750   Resume EXITP

End Function
