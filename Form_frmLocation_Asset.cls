VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLocation_Asset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLocation_Asset"

'VGC 07/23/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Select Account:   A {cmbAccounts}
' **   Acct Number:      B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Description:      D {asset_description on frmLocation_Asset_Sub}
' **   Select Asset:     E {cmbAssets}
' **   Current Loc:      I {cmbLocations_Current}
' **   All Group:        L {opgGroup_optAll}
' **   Acct Name:        M {opgAccountSource_optName}
' **   Select None:      N {cmdSelectNone}
' **   Account Group:    O {opgGroup_optAccount}
' **   Reset:            R {cmdResetFilter}
' **   Select All:       S {cmdSelectAll}
' **   Asset Group:      T {opgGroup_optAsset}
' **   Update:           U {cmdUpdate}
' **   Move:             V {locasset_mark on frmLocation_Asset_Sub}
' **   New Loc:          W {cmbLocations_New}
' **   Exit:             X {cmdClose}
' **   Accnt/Asset Grp:  / {opgGroup_optAccountAsset}
'F  G  H  J  K  P Q  Z

' ** Shortcut Alt-Shift keys responsive from this form:
' **   Asset Cusip:      I {opgAssetSource_optCusip}
' **   Asset Name:       N {opgAssetSource_optName}
' **   Asset Type:       Y {opgAssetSource_optType}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnSelectAll_Focus As Boolean, blnSelectAll_MouseDown As Boolean
Private blnSelectNone_Focus As Boolean, blnSelectNone_MouseDown As Boolean
Private blnReset_Focus As Boolean, blnReset_MouseDown As Boolean

Private blnFocusCurrent As Boolean, blnFocusNew As Boolean, blnRecalcTots As Boolean
Private lngRecsCur As Long, lngTpp As Long
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

150           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
160           DoEvents

170           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190           opgAccountSource_AfterUpdate  ' ** Procedure: Below.
200           opgAssetSource_AfterUpdate  ' ** Procedure: Below.
210           opgGroup_AfterUpdate  ' ** Procedure: Below.

220           CLR_DISABLED_FG = CLR_DKGRY
230           CLR_DISABLED_BG = CLR_LTTEAL

240           blnFocusCurrent = False: blnFocusNew = False: blnRecalcTots = False
250           blnSelectAll_Focus = False: blnSelectAll_MouseDown = False
260           blnSelectNone_Focus = False: blnSelectNone_MouseDown = False
270           blnReset_Focus = False: blnReset_MouseDown = False

280           .cmdUpdate.Enabled = False

290           .Cnt_AA.Locked = False
300           .Cnt_AA.BackStyle = acBackStyleTransparent
310           .Cnt_AA.BorderColor = WIN_CLR_DISR
320           .Cnt_L.Visible = False
330           .Cnt_LA.Visible = False
340           .Cnt_J.Visible = False
350           .Cnt_JC.Visible = False

360           .Cnt_AA.ForeColor = CLR_DISABLED_FG
370           .Cnt_AA.BackColor = CLR_DISABLED_BG
380           .Cnt_L.ForeColor = CLR_DISABLED_FG
390           .Cnt_L.BackColor = CLR_DISABLED_BG
400           .Cnt_LA.ForeColor = CLR_DISABLED_FG
410           .Cnt_LA.BackColor = CLR_DISABLED_BG
420           .Cnt_J.ForeColor = CLR_DISABLED_FG
430           .Cnt_J.BackColor = CLR_DISABLED_BG
440           .Cnt_JC.ForeColor = CLR_DISABLED_FG
450           .Cnt_JC.BackColor = CLR_DISABLED_BG

460           DoCmd.Hourglass False

470         Else
480           Cancel = -1
490         End If
500       Else
510         Cancel = -1
520       End If

530     End With

540     If Cancel = -1 Then
550       Beep
560       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
570     End If

EXITP:
580     Exit Sub

ERRH:
590     DoCmd.Hourglass False
600     Select Case ERR.Number
        Case Else
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End Select
630     Resume EXITP

End Sub

Private Sub Form_Load()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim frm As Access.Form
        Dim strFilter As String

710     With Me
720       Set frm = .frmLocation_Asset_Sub.Form
730       strFilter = "[Location_ID] = 0"  ' ** Empty the subform.
740       frm.Filter = strFilter
750       frm.FilterOn = True
760       frm.ColHeaders False  ' ** Form Procedure: frmLocation_Asset_Sub.
770     End With

EXITP:
780     Set frm = Nothing
790     Exit Sub

ERRH:
800     DoCmd.Hourglass False
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

910     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
920     intShiftDown = (Shift And acShiftMask) > 0
930     intAltDown = (Shift And acAltMask) > 0
940     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Acct Number:      B {opgAccountSource_optNumber}
        ' **   Description:      D {asset_description on frmLocation_Asset_Sub}
        ' **   Acct Name:        M {opgAccountSource_optName}
        ' **   Move:             V {locasset_mark on frmLocation_Asset_Sub}

        ' ** Shortcut Alt-Shift keys to other forms and functionality:
        ' **   Asset Cusip:      I {opgAssetSource_optCusip}
        ' **   Asset Name:       N {opgAssetSource_optName}
        ' **   Asset Type:       Y {opgAssetSource_optType}

        ' ** Plain keys.
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyEscape
970         intRetVal = 0
980         cmdClose_Click  ' ** Procedure: Below.
990       End Select
1000    End If

        ' ** Alt keys.
1010    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1020      Select Case intRetVal
          Case vbKeyB  ' ** Because the '_lbl' is empty.
1030        With Me
1040          intRetVal = 0
1050          If .opgAccountSource.Enabled = True Then
1060            .opgAccountSource.SetFocus
1070            .opgAccountSource = .opgAccountSource_optNumber.OptionValue
1080            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1090          Else
1100            Beep
1110          End If
1120        End With
1130      Case vbKeyD
1140        With Me
1150          intRetVal = 0
1160          lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
1170          If lngRecsCur > 0& Then
1180            .frmLocation_Asset_Sub.SetFocus
1190            .frmLocation_Asset_Sub.Form.FocusHolder.SetFocus
1200          Else
1210            Beep
1220          End If
1230        End With
1240      Case vbKeyM  ' ** Because the '_lbl' is empty.
1250        With Me
1260          intRetVal = 0
1270          If .opgAccountSource.Enabled = True Then
1280            .opgAccountSource.SetFocus
1290            .opgAccountSource = .opgAccountSource_optName.OptionValue
1300            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1310          Else
1320            Beep
1330          End If
1340        End With
1350      Case vbKeyV
1360        With Me
1370          intRetVal = 0
1380          lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
1390          If lngRecsCur > 0 Then
1400            .frmLocation_Asset_Sub.SetFocus
1410            .frmLocation_Asset_Sub.Form.locasset_mark.SetFocus
1420          Else
1430            Beep
1440          End If
1450        End With
1460      Case vbKeyX
1470        intRetVal = 0
1480        cmdClose_Click  ' ** Procedure: Below.
1490      End Select
1500    End If

        ' ** Alt-Shift keys.
1510    If (Not intCtrlDown) And intAltDown And intShiftDown Then
1520      Select Case intRetVal
          Case vbKeyI
1530        With Me
1540          intRetVal = 0
1550          If .opgAssetSource.Enabled = True Then
1560            .opgAssetSource.SetFocus
1570            .opgAssetSource = .opgAssetSource_optCusip.OptionValue
1580            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
1590          Else
1600            Beep
1610          End If
1620        End With
1630      Case vbKeyN
1640        With Me
1650          intRetVal = 0
1660          If .opgAssetSource.Enabled = True Then
1670            .opgAssetSource.SetFocus
1680            .opgAssetSource = .opgAssetSource_optName.OptionValue
1690            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
1700          Else
1710            Beep
1720          End If
1730        End With
1740      Case vbKeyY
1750        With Me
1760          intRetVal = 0
1770          If .opgAssetSource.Enabled = True Then
1780            .opgAssetSource.SetFocus
1790            .opgAssetSource = .opgAssetSource_optType.OptionValue
1800            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
1810          Else
1820            Beep
1830          End If
1840        End With
1850      End Select
1860    End If

EXITP:
1870    KeyCode = intRetVal
1880    Exit Sub

ERRH:
1890    intRetVal = 0
1900    Select Case ERR.Number
        Case Else
1910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1920    End Select
1930    Resume EXITP

End Sub

Private Sub Form_Timer()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim lngTmp01 As Long

2010    With Me
2020      .TimerInterval = 0&
2030      If gblnSetFocus = True Then
2040        gblnSetFocus = False
2050        strFormName = vbNullString
2060  On Error Resume Next
2070        Do Until strFormName = THIS_NAME
2080          DoCmd.SelectObject acForm, THIS_NAME, False
2090          strFormName = Screen.ActiveForm.Name
2100        Loop
2110  On Error GoTo ERRH
2120      ElseIf blnRecalcTots = True Then
2130        blnRecalcTots = False
2140  On Error Resume Next
2150        lngTmp01 = Abs(Nz(.frmLocation_Asset_Sub.Form.locasset_mark_sum, 0))
2160        If ERR.Number <> 0 Then
2170  On Error GoTo ERRH
2180          lngTmp01 = 0&
2190        Else
2200  On Error GoTo ERRH
2210        End If
2220        .locasset_mark_sum = lngTmp01
2230        If lngTmp01 > 0 Then
2240          .locasset_mark_sum.ForeColor = CLR_GRN
2250        Else
2260          .locasset_mark_sum.ForeColor = CLR_BLU
2270        End If
2280        lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
2290        If lngRecsCur > 0 Then
2300          .frmLocation_Asset_Sub.Form.ColHeaders True  ' ** Form Procedure: frmLocation_Asset_Sub.
2310        Else
2320          .cmbLocations_Current.SetFocus
2330          .cmdUpdate.Enabled = False
2340          .frmLocation_Asset_Sub.Form.ColHeaders False  ' ** Form Procedure: frmLocation_Asset_Sub.
2350        End If
2360      ElseIf blnFocusCurrent = True Then
2370        blnFocusCurrent = False
2380        .cmbLocations_Current.SetFocus
2390        .cmbLocations_Current.SelLength = 0
2400        .cmbLocations_Current.SelStart = 0
2410        blnRecalcTots = True
2420        .TimerInterval = 500&
2430      ElseIf blnFocusNew = True Then
2440        blnFocusNew = False
2450        .cmbLocations_New.SetFocus
2460        .cmbLocations_New.SelLength = 0
2470        .cmbLocations_New.SelStart = 0
2480        blnRecalcTots = True
2490        .TimerInterval = 500&
2500      End If
2510    End With

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim msgResponse As VbMsgBoxResult

2610    With Me

2620      If .cmdUpdate.Enabled = True And .locasset_mark_sum <> 0 Then
2630        Beep
2640        msgResponse = MsgBox("You have made changes, but have not confirmed them" & vbCrLf & _
              "by clicking the Update button." & vbCrLf & vbCrLf & _
              "Are you sure you want to abandon changes?", vbQuestion + vbYesNo, "Changes Not Saved")
2650      Else
2660        msgResponse = vbYes
2670      End If

2680      If msgResponse = vbYes Then

2690        Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2700        Set dbs = CurrentDb
            ' ** Empty tmpAccount.
2710        Set qdf = dbs.QueryDefs("qryLocation_Asset_03_01")
2720        qdf.Execute
2730        Set qdf = Nothing
2740        DoEvents
2750        dbs.Close
2760        Set dbs = Nothing
2770        DoEvents

2780        DoCmd.Hourglass True
2790        DoEvents
2800        Select Case strCallingForm
            Case "frmLocations"
2810          gstrReturningForm = THIS_NAME
2820          gblnSetFocus = True
2830          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Utility"
2840        End Select

2850      End If

2860    End With

EXITP:
2870    Set qdf = Nothing
2880    Set dbs = Nothing
2890    Exit Sub

ERRH:
2900    DoCmd.Hourglass False
2910    Select Case ERR.Number
        Case Else
2920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2930    End Select
2940    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3010    With Me
3020      If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
3030        Select Case blnSelectAll_Focus
            Case True
3040          .cmdSelectAll_raised_semifocus_dots_img.Visible = True
3050          .cmdSelectAll_raised_img.Visible = False
3060        Case False
3070          .cmdSelectAll_raised_img.Visible = True
3080          .cmdSelectAll_raised_semifocus_dots_img.Visible = False
3090        End Select
3100        .cmdSelectAll_raised_focus_img.Visible = False
3110        .cmdSelectAll_raised_focus_dots_img.Visible = False
3120        .cmdSelectAll_sunken_focus_dots_img.Visible = False
3130        .cmdSelectAll_raised_img_dis.Visible = False
3140      End If
3150      If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
3160        Select Case blnSelectNone_Focus
            Case True
3170          .cmdSelectNone_raised_semifocus_dots_img.Visible = True
3180          .cmdSelectNone_raised_img.Visible = False
3190        Case False
3200          .cmdSelectNone_raised_img.Visible = True
3210          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
3220        End Select
3230        .cmdSelectNone_raised_focus_img.Visible = False
3240        .cmdSelectNone_raised_focus_dots_img.Visible = False
3250        .cmdSelectNone_sunken_focus_dots_img.Visible = False
3260        .cmdSelectNone_raised_img_dis.Visible = False
3270      End If
3280      If .cmdResetFilter_raised_focus_dots_img.Visible = True Or .cmdResetFilter_raised_focus_img.Visible = True Then
3290        Select Case blnReset_Focus
            Case True
3300          .cmdResetFilter_raised_semifocus_dots_img.Visible = True
3310          .cmdResetFilter_raised_img.Visible = False
3320        Case False
3330          .cmdResetFilter_raised_img.Visible = True
3340          .cmdResetFilter_raised_semifocus_dots_img.Visible = False
3350        End Select
3360        .cmdResetFilter_raised_focus_img.Visible = False
3370        .cmdResetFilter_raised_focus_dots_img.Visible = False
3380        .cmdResetFilter_sunken_focus_dots_img.Visible = False
3390        .cmdResetFilter_raised_img_dis.Visible = False
3400      End If
3410    End With

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub cmdClose_Click()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3610    DoCmd.Close acForm, THIS_NAME

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
3790          .cmbLocations_Current.SetFocus
3800        End With
3810      End Select
3820    End If

        ' ** Shift keys.
3830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3840      Select Case intRetVal
          Case vbKeyTab
3850        With Me
3860          intRetVal = 0
3870          If .cmdUpdate.Enabled = True Then
3880            .cmdUpdate.SetFocus
3890          ElseIf .cmdSelectNone.Enabled = True Then
3900            .cmdSelectNone.SetFocus
3910          Else
3920            lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
3930            If lngRecsCur > 0& Then
3940              .frmLocation_Asset_Sub.SetFocus
3950            Else
3960              If .cmdResetFilter.Enabled = True Then
3970                .cmdResetFilter.SetFocus
3980              ElseIf .cmbAssets.Enabled = True Then
3990                .cmbAssets.SetFocus
4000              ElseIf .cmbAccounts.Enabled = True Then
4010                .cmbAccounts.SetFocus
4020              Else
4030                .opgGroup.SetFocus
4040              End If
4050            End If
4060          End If
4070        End With
4080      End Select
4090    End If

EXITP:
4100    KeyCode = intRetVal
4110    Exit Sub

ERRH:
4120    intRetVal = 0
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim frm As Access.Form
        Dim lngLocID_Old As Long, lngLocID_New As Long, strLoc_Old As String, strLoc_New As String
        Dim lngAssetNo As Long
        Dim lngAA_Cnt As Long, lngL_Cnt As Long, lngLA_Cnt As Long, lngJ_Cnt As Long, lngJC_Cnt As Long
        Dim strAccountNo As String, strLocName As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long

4210    blnContinue = True

4220    With Me
4230      Select Case IsNull(.cmbLocations_Current)
          Case True
4240        blnContinue = False
4250        Beep
4260        MsgBox "A Current Location and a New Location must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
4270        .cmbLocations_Current.SetFocus
4280      Case False
4290        lngLocID_Old = .cmbLocations_Current
4300        Select Case IsNull(.cmbLocations_New)
            Case True
4310          blnContinue = False
4320          Beep
4330          MsgBox "A New Location must be chosen to continue.", vbInformation + vbOKOnly, "Entry Required"
4340          .cmbLocations_New.SetFocus
4350        Case False
4360          lngLocID_New = .cmbLocations_New
4370          If lngLocID_New = lngLocID_Old Then
4380            blnContinue = False
4390            Beep
4400            MsgBox "The New Location must be different from the Current Location.", vbInformation + vbOKOnly, "Invalid Entry"
4410            .cmbLocations_New.SetFocus
4420          Else
4430            strLocName = .cmbLocations_New.Column(1)
4440            If .Cnt_AA = 0 And .Cnt_L = 0 And .Cnt_LA = 0 And .Cnt_J = 0 And .Cnt_JC = 0 Then
4450              blnContinue = False
4460              Beep
4470              MsgBox "No records were found for the chosen location.", vbInformation + vbOKOnly, "Nothing To Do"
4480              .cmbLocations_Current.SetFocus
4490            Else
4500              Select Case .opgGroup
                  Case .opgGroup_optAll.OptionValue
                    ' ** Ready to go.
4510              Case .opgGroup_optAccount.OptionValue
4520                Select Case IsNull(.cmbAccounts)
                    Case True
4530                  blnContinue = False
4540                  Beep
4550                  MsgBox "You must select an Account to continue.", vbInformation + vbOKOnly, "Entry Required"
4560                  .cmbAccounts.SetFocus
4570                Case False
                      ' ** Ready to go.
4580                  strAccountNo = .cmbAccounts
4590                End Select
4600              Case .opgGroup_optAsset.OptionValue
4610                Select Case IsNull(.cmbAssets)
                    Case True
4620                  blnContinue = False
4630                  Beep
4640                  MsgBox "You must select an Asset to continue.", vbInformation + vbOKOnly, "Entry Required"
4650                  .cmbAssets.SetFocus
4660                Case False
                      ' ** Ready to go.
4670                  lngAssetNo = .cmbAssets
4680                End Select
4690              Case .opgGroup_optAccountAsset.OptionValue
4700                Select Case IsNull(.cmbAccounts)
                    Case True
4710                  blnContinue = False
4720                  Beep
4730                  MsgBox "You must select an Account to continue.", vbInformation + vbOKOnly, "Entry Required"
4740                  .cmbAccounts.SetFocus
4750                Case False
4760                  strAccountNo = .cmbAccounts
4770                  Select Case IsNull(.cmbAssets)
                      Case True
4780                    blnContinue = False
4790                    Beep
4800                    MsgBox "You must select an Asset to continue.", vbInformation + vbOKOnly, "Entry Required"
4810                    .cmbAssets.SetFocus
4820                  Case False
                        ' ** Ready to go.
4830                    lngAssetNo = .cmbAssets
4840                  End Select
4850                End Select
4860              End Select
4870            End If
4880          End If
4890        End Select
4900      End Select

4910      If blnContinue = True Then
4920        strLoc_Old = .cmbLocations_Current.Column(1)
4930        strLoc_New = .cmbLocations_New.Column(1)
4940        varTmp00 = .locasset_mark_sum
4950        Select Case IsNull(varTmp00)
            Case True
4960          strTmp01 = "these assets"
4970        Case False
4980          If varTmp00 = 1 Then
4990            strTmp01 = "this asset"
5000          Else
5010            strTmp01 = "these assets"
5020          End If
5030        End Select
5040        msgResponse = MsgBox("Are you sure you want to reassign " & strTmp01 & "?" & vbCrLf & _
              "From:" & vbCrLf & "    " & strLoc_Old & vbCrLf & _
              "To:" & vbCrLf & "    " & strLoc_New, vbQuestion + vbYesNo, "Change Location")
5050        If msgResponse <> vbYes Then
5060          blnContinue = False
5070        End If
5080      End If

5090      If blnContinue = True Then

5100        DoCmd.Hourglass True
5110        DoEvents

5120        lngAA_Cnt = .Cnt_AA
5130        lngL_Cnt = .Cnt_L
5140        lngLA_Cnt = .Cnt_LA
5150        lngJ_Cnt = .Cnt_J
5160        lngJC_Cnt = .Cnt_JC

5170        Set dbs = CurrentDb

5180        For lngX = 1& To 9&

5190          Select Case .opgGroup
              Case .opgGroup_optAll.OptionValue
5200            Select Case lngX
                Case 1&
                  ' ** Update qryLocation_Asset_15ab (ActiveAssets, linked to qryLocation_Asset_15aa
                  ' ** (tmpLocation_Asset, Location_ID2 <> Null, locasset_mark = True), with Location_ID_new).
5210              Set qdf = dbs.QueryDefs("qryLocation_Asset_15ba")
5220            Case 2&
                  ' ** Update qryLocation_Asset_16ab (Ledger, linked to qryLocation_Asset_16aa (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True), with Location_ID_new, by assetdate).
5230              Set qdf = dbs.QueryDefs("qryLocation_Asset_16ba")
5240            Case 3&
                  ' ** Update qryLocation_Asset_16ac (Ledger, linked to qryLocation_Asset_16aa (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True), with Location_ID_new, by PurchaseDate).
5250              Set qdf = dbs.QueryDefs("qryLocation_Asset_16bb")
5260            Case 4&
                  ' ** Update qryLocation_Asset_17ab (LedgerArchive, linked to qryLocation_Asset_17aa (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True), with Location_ID_new, by assetdate).
5270              Set qdf = dbs.QueryDefs("qryLocation_Asset_17ba")
5280            Case 5&
                  ' ** Update qryLocation_Asset_17ac (LedgerArchive, linked to qryLocation_Asset_17aa (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True), with Location_ID_new, by PurchaseDate).
5290              Set qdf = dbs.QueryDefs("qryLocation_Asset_17bb")
5300            Case 6&
                  ' ** Update qryLocation_Asset_18ab (Journal, linked to qryLocation_Asset_18aa (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True), with Location_ID_new, by assetdate).
5310              Set qdf = dbs.QueryDefs("qryLocation_Asset_18ba")
5320            Case 7&
                  ' ** Update qryLocation_Asset_18ac (Journal, linked to qryLocation_Asset_18aa (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True), with Location_ID_new, by PurchaseDate).
5330              Set qdf = dbs.QueryDefs("qryLocation_Asset_18bb")
5340            Case 8&
                  ' ** Update qryLocation_Asset_19ab (tblJournal_Column, linked to qryLocation_Asset_19aa (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True), with Location_ID_new, by assetdate).
5350              Set qdf = dbs.QueryDefs("qryLocation_Asset_19ba")
5360            Case 9&
                  ' ** Update qryLocation_Asset_19ac (tblJournal_Column, linked to qryLocation_Asset_19aa (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True), with Location_ID_new, by PurchaseDate).
5370              Set qdf = dbs.QueryDefs("qryLocation_Asset_19bb")
5380            End Select
5390          Case .opgGroup_optAccount.OptionValue
5400            Select Case lngX
                Case 1&
                  ' ** Update qryLocation_Asset_15cb (ActiveAssets, linked to qryLocation_Asset_15ca (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [actno]), with Location_ID_new).
5410              Set qdf = dbs.QueryDefs("qryLocation_Asset_15da")
5420            Case 2&
                  ' ** Update qryLocation_Asset_16cb (Ledger, linked to qryLocation_Asset_16ca (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [actno]), with Location_ID_new, by assetdate).
5430              Set qdf = dbs.QueryDefs("qryLocation_Asset_16da")
5440            Case 3&
                  ' ** Update qryLocation_Asset_16cc (Ledger, linked to qryLocation_Asset_16ca (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [actno]), with Location_ID_new, by PurchaseDate).
5450              Set qdf = dbs.QueryDefs("qryLocation_Asset_16db")
5460            Case 4&
                  ' ** Update qryLocation_Asset_17cb (LedgerArchive, linked to qryLocation_Asset_17ca (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [actno]), with Location_ID_new, by assetdate).
5470              Set qdf = dbs.QueryDefs("qryLocation_Asset_17da")
5480            Case 5&
                  ' ** Update qryLocation_Asset_17cc (LedgerArchive, linked to qryLocation_Asset_17ca (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [actno]), with Location_ID_new, by PurchaseDate).
5490              Set qdf = dbs.QueryDefs("qryLocation_Asset_17db")
5500            Case 6&
                  ' ** Update qryLocation_Asset_18cb (Journal, linked to qryLocation_Asset_18ca (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [actno]), with Location_ID_new, by assetdate).
5510              Set qdf = dbs.QueryDefs("qryLocation_Asset_18da")
5520            Case 7&
                  ' ** Update qryLocation_Asset_18cc (Journal, linked to qryLocation_Asset_18ca (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [actno]), with Location_ID_new, by PurchaseDate).
5530              Set qdf = dbs.QueryDefs("qryLocation_Asset_18db")
5540            Case 8&
                  ' ** Update qryLocation_Asset_19cb (tblJournal_Column, linked to qryLocation_Asset_19ca (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [actno]), with Location_ID_new, by assetdate).
5550              Set qdf = dbs.QueryDefs("qryLocation_Asset_19da")
5560            Case 9&
                  ' ** Update qryLocation_Asset_19cc (tblJournal_Column, linked to qryLocation_Asset_19ca (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [actno]), with Location_ID_new, by PurchaseDate).
5570              Set qdf = dbs.QueryDefs("qryLocation_Asset_19db")
5580            End Select
5590            With qdf.Parameters
5600              ![actno] = strAccountNo
5610            End With
5620          Case .opgGroup_optAsset.OptionValue
5630            Select Case lngX
                Case 1&
                  ' ** Update qryLocation_Asset_15eb (ActiveAssets, linked to qryLocation_Asset_15ea (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [astno]), with Location_ID_new).
5640              Set qdf = dbs.QueryDefs("qryLocation_Asset_15fa")
5650            Case 2&
                  ' ** Update qryLocation_Asset_16eb (Ledger, linked to qryLocation_Asset_16ea (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [astno]), with Location_ID_new, by assetdate).
5660              Set qdf = dbs.QueryDefs("qryLocation_Asset_16fa")
5670            Case 3&
                  ' ** Update qryLocation_Asset_16ec (Ledger, linked to qryLocation_Asset_16ea (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [astno]), with Location_ID_new, by PurchaseDate).
5680              Set qdf = dbs.QueryDefs("qryLocation_Asset_16fb")
5690            Case 4&
                  ' ** Update qryLocation_Asset_17eb (LedgerArchive, linked to qryLocation_Asset_17ea (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [astno]), with Location_ID_new, by assetdate).
5700              Set qdf = dbs.QueryDefs("qryLocation_Asset_17fa")
5710            Case 5&
                  ' ** Update qryLocation_Asset_17ec (LedgerArchive, linked to qryLocation_Asset_17ea (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [astno]), with Location_ID_new, by PurchaseDate).
5720              Set qdf = dbs.QueryDefs("qryLocation_Asset_17fb")
5730            Case 6&
                  ' ** Update qryLocation_Asset_18eb (Journal, linked to qryLocation_Asset_18ea (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [astno]), with Location_ID_new, by assetdate).
5740              Set qdf = dbs.QueryDefs("qryLocation_Asset_18fa")
5750            Case 7&
                  ' ** Update qryLocation_Asset_18ec (Journal, linked to qryLocation_Asset_18ea (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [astno]), with Location_ID_new, by PurchaseDate).
5760              Set qdf = dbs.QueryDefs("qryLocation_Asset_18fb")
5770            Case 8&
                  ' ** Update qryLocation_Asset_19eb (tblJournal_Column, linked to qryLocation_Asset_19ea (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [astno]), with Location_ID_new, by assetdate).
5780              Set qdf = dbs.QueryDefs("qryLocation_Asset_19fa")
5790            Case 9&
                  ' ** Update qryLocation_Asset_19ec (tblJournal_Column, linked to qryLocation_Asset_19ea (tmpLocation_Asset,
                  ' ** Location_ID2 <> Null, locasset_mark = True, by specified [astno]), with Location_ID_new, by PurchaseDate).
5800              Set qdf = dbs.QueryDefs("qryLocation_Asset_19fb")
5810            End Select
5820            With qdf.Parameters
5830              ![astno] = lngAssetNo
5840            End With
5850          Case .opgGroup_optAccountAsset.OptionValue
5860            Select Case lngX
                Case 1&
                  ' ** Update qryLocation_Asset_15gb (ActiveAssets, linked to qryLocation_Asset_15ga
                  ' ** (tmpLocation_Asset, Location_ID2 <> Null, locasset_mark = True,
                  ' ** by specified [actno], [astno]), with Location_ID_new.).
5870              Set qdf = dbs.QueryDefs("qryLocation_Asset_15ha")
5880            Case 2&
                  ' ** Update qryLocation_Asset_16gb (Ledger, linked to qryLocation_Asset_16ga
                  ' ** (tmpLocation_Asset, Location_ID2 <> Null, locasset_mark = True,
                  ' ** by specified [actno], [astno]), with Location_ID_new, by assetdate).
5890              Set qdf = dbs.QueryDefs("qryLocation_Asset_16ha")
5900            Case 3&
                  ' ** Update qryLocation_Asset_16gc (Ledger, linked to qryLocation_Asset_16ga
                  ' ** (tmpLocation_Asset, Location_ID2 <> Null, locasset_mark = True,
                  ' ** by specified [actno], [astno]), with Location_ID_new, by PurchaseDate).
5910              Set qdf = dbs.QueryDefs("qryLocation_Asset_16hb")
5920            Case 4&
                  ' ** Update qryLocation_Asset_17gb (LedgerArchive, linked to qryLocation_Asset_17ga
                  ' ** (tmpLocation_Asset, Location_ID2 <> Null, locasset_mark = True,
                  ' ** by specified [actno], [astno]), with Location_ID_new, by assetdate).
5930              Set qdf = dbs.QueryDefs("qryLocation_Asset_17ha")
5940            Case 5&
                  ' ** Update qryLocation_Asset_17gc (LedgerArchive, linked to qryLocation_Asset_17ga
                  ' ** (tmpLocation_Asset, Location_ID2 <> Null, locasset_mark = True,
                  ' ** by specified [actno], [astno]), with Location_ID_new, by PurchaseDate).
5950              Set qdf = dbs.QueryDefs("qryLocation_Asset_17hb")
5960            Case 6&
                  ' ** Update qryLocation_Asset_18gb (Journal, linked to qryLocation_Asset_18ga
                  ' ** (tmpLocation_Asset, Location_ID2 <> Null, locasset_mark = True, by
                  ' ** specified [actno], [astno]), with Location_ID_new, by assetdate).
5970              Set qdf = dbs.QueryDefs("qryLocation_Asset_18ha")
5980            Case 7&
                  ' ** Update qryLocation_Asset_18gc (Journal, linked to qryLocation_Asset_18ga
                  ' ** (tmpLocation_Asset, Location_ID2 <> Null, locasset_mark = True,
                  ' ** by specified [actno], [astno]), with Location_ID_new, by PurchaseDate).
5990              Set qdf = dbs.QueryDefs("qryLocation_Asset_18hb")
6000            Case 8&
                  ' ** Update qryLocation_Asset_19gb (tblJournal_Column, linked to qryLocation_Asset_19ga
                  ' ** (tmpLocation_Asset, Location_ID2 <> Null, locasset_mark = True,
                  ' ** by specified [actno], [astno]), with Location_ID_new, by assetdate).
6010              Set qdf = dbs.QueryDefs("qryLocation_Asset_19ha")
6020            Case 9&
                  ' ** Update qryLocation_Asset_19gc (tblJournal_Column, linked to qryLocation_Asset_19ga
                  ' ** (tmpLocation_Asset, Location_ID2 <> Null, locasset_mark = True,
                  ' ** by specified [actno], [astno]), with Location_ID_new, by PurchaseDate).
6030              Set qdf = dbs.QueryDefs("qryLocation_Asset_19hb")
6040            End Select
6050            With qdf.Parameters
6060              ![actno] = strAccountNo
6070              ![astno] = lngAssetNo
6080            End With
6090          End Select
6100          qdf.Execute dbFailOnError
6110          Set qdf = Nothing

6120        Next  ' ** lngX.

6130        Set frm = .frmLocation_Asset_Sub.Form
6140        With frm
6150          .accountno.ForeColor = CLR_LTTEAL
6160          .cusip.ForeColor = CLR_LTTEAL
6170          .asset_description.ForeColor = CLR_LTTEAL
6180          .assettype.ForeColor = CLR_LTTEAL
6190          .assetdate.ForeColor = CLR_LTTEAL
6200          .transdate.ForeColor = CLR_LTTEAL
6210          .shareface.ForeColor = CLR_LTTEAL
6220        End With

6230        With dbs
              ' ** Empty tmpLocation_Asset.
6240          Set qdf = .QueryDefs("qryLocation_Asset_Sub_01")
6250          qdf.Execute
6260          Set qdf = Nothing
              ' ** Append qryLocation_Asset_Sub_02(ActiveAssets,
              ' ** linked to Location, MasterAsset) to tmpLocation_Asset.
6270          Set qdf = .QueryDefs("qryLocation_Asset_Sub_03")
6280          qdf.Execute
6290          Set qdf = Nothing
6300        End With

6310        dbs.Close
6320        Set dbs = Nothing

6330        frm.Requery

6340        With frm
6350          .accountno.ForeColor = CLR_DKGRY
6360          .cusip.ForeColor = CLR_DKGRY
6370          .asset_description.ForeColor = CLR_DKGRY
6380          .assettype.ForeColor = CLR_DKGRY
6390          .assetdate.ForeColor = CLR_DKGRY
6400          .transdate.ForeColor = CLR_DKGRY
6410          .shareface.ForeColor = CLR_DKGRY
6420        End With

6430        Loc_Counts  ' ** Function: Below.
6440        DoEvents

6450        .cmbLocations_Current.SetFocus
6460        frm.RecalcTots_Set True  ' ** Form Procedure: frmLocation_Asset_Sub.
6470        DoEvents
6480        .cmdUpdate.Enabled = False

6490        DoCmd.Hourglass False
6500        MsgBox "All records have been successfully reassigned.", vbInformation + vbOKOnly, "Reassignment Successful"

6510      End If

6520    End With

EXITP:
6530    Set frm = Nothing
6540    Set qdf = Nothing
6550    Set dbs = Nothing
6560    Exit Sub

ERRH:
6570    DoCmd.Hourglass False
6580    Select Case ERR.Number
        Case Else
6590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6600    End Select
6610    Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

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
6790          .cmdClose.SetFocus
6800        End With
6810      End Select
6820    End If

        ' ** Shift keys.
6830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6840      Select Case intRetVal
          Case vbKeyTab
6850        With Me
6860          intRetVal = 0
6870          If .cmdSelectNone.Enabled = True Then
6880            .cmdSelectNone.SetFocus
6890          Else
6900            lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
6910            If lngRecsCur > 0& Then
6920              .frmLocation_Asset_Sub.SetFocus
6930            Else
6940              If .cmdResetFilter.Enabled = True Then
6950                .cmdResetFilter.SetFocus
6960              ElseIf .cmbAssets.Enabled = True Then
6970                .cmbAssets.SetFocus
6980              ElseIf .cmbAccounts.Enabled = True Then
6990                .cmbAccounts.SetFocus
7000              Else
7010                .opgGroup.SetFocus
7020              End If
7030            End If
7040          End If
7050        End With
7060      End Select
7070    End If

EXITP:
7080    KeyCode = intRetVal
7090    Exit Sub

ERRH:
7100    intRetVal = 0
7110    Select Case ERR.Number
        Case Else
7120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7130    End Select
7140    Resume EXITP

End Sub

Private Sub cmbLocations_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_box_MouseMove"

7210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7220    Exit Sub

ERRH:
7230    Select Case ERR.Number
        Case Else
7240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7250    End Select
7260    Resume EXITP

End Sub

Private Sub cmbLocations_Current_AfterUpdate()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_Current_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim frm As Access.Form
        Dim lngLocID As Long, strLocName As String, lngAssetNo As Long
        Dim strAccountNo As String, strFilter As String
        Dim lngRecs As Long
        Dim lngX As Long

7310    With Me

7320      DoCmd.Hourglass True
7330      DoEvents

7340      .Cnt_AA.Locked = False
7350      .Cnt_AA = 0&
7360      .Cnt_AA.BackStyle = acBackStyleTransparent
7370      .Cnt_AA.BorderColor = WIN_CLR_DISR
7380      .Cnt_L.Visible = False
7390      .Cnt_L = 0&
7400      .Cnt_LA.Visible = False
7410      .Cnt_LA = 0&
7420      .Cnt_J.Visible = False
7430      .Cnt_J = 0&
7440      .Cnt_JC.Visible = False
7450      .Cnt_JC = 0&
7460      DoEvents

7470      Set dbs = CurrentDb
7480      With dbs
            ' ** Update tmpLocation_Asset, for locasset_mark = False.
7490        Set qdf = .QueryDefs("qryLocation_Asset_06b")
7500        qdf.Execute
7510        Set qdf = Nothing
7520        .Close
7530      End With
7540      Set dbs = Nothing

7550      Set frm = .frmLocation_Asset_Sub.Form

7560      Select Case IsNull(.cmbLocations_Current)
          Case True
7570        strFilter = "[Location_ID] = 0"
7580      Case False
7590        lngLocID = .cmbLocations_Current
7600        Select Case .opgGroup
            Case .opgGroup_optAll.OptionValue
7610          strFilter = "[Location_ID] = " & CStr(lngLocID)
7620        Case .opgGroup_optAccount.OptionValue
7630          Select Case IsNull(.cmbAccounts)
              Case True
7640            strAccountNo = "0"
7650          Case False
7660            strAccountNo = .cmbAccounts
7670          End Select
7680          strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [accountno] = '" & strAccountNo & "'"
7690        Case .opgGroup_optAsset.OptionValue
7700          Select Case IsNull(.cmbAssets)
              Case True
7710            lngAssetNo = 0&
7720          Case False
7730            lngAssetNo = .cmbAssets
7740          End Select
7750          strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [assetno] = " & CStr(lngAssetNo)
7760        Case .opgGroup_optAccountAsset.OptionValue
7770          Select Case IsNull(.cmbAccounts)
              Case True
7780            strAccountNo = "0"
7790          Case False
7800            strAccountNo = .cmbAccounts
7810          End Select
7820          strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [accountno] = '" & strAccountNo & "'"
7830          Select Case IsNull(.cmbAssets)
              Case True
7840            lngAssetNo = 0&
7850          Case False
7860            lngAssetNo = .cmbAssets
7870          End Select
7880          strFilter = strFilter & " And [assetno] = " & CStr(lngAssetNo)
7890        End Select
7900      End Select

7910      frm.Filter = strFilter
7920      frm.FilterOn = True

7930      If lngLocID > 0& Then
7940        If IsNull(.cmbLocations_New) = False Then
7950          If .cmbLocations_New = lngLocID Then
7960            .cmbLocations_New = Null
7970          End If
7980        End If
7990      End If

8000      Loc_Counts  ' ** Function: Below.
8010      DoEvents

8020      Validate  ' ** Function: Below.
8030      DoEvents

8040      lngRecsCur = frm.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
8050      If lngLocID > 0& And IsNull(.cmbLocations_New) = False Then
8060        If lngRecsCur > 0& Then
8070          lngLocID = .cmbLocations_New.Column(0)
8080          strLocName = .cmbLocations_New.Column(1)
8090          Set rst = frm.RecordsetClone
8100          With rst
8110            .MoveLast
8120            lngRecs = .RecordCount
8130            .MoveFirst
8140            For lngX = 1& To lngRecs
8150              .Edit
8160              ![locasset_mark] = True
8170              ![Location_ID2] = lngLocID
8180              ![Location_Name2] = strLocName
8190              ![locasset_datemodified] = Now()
8200              .Update
8210              If lngX < lngRecs Then .MoveNext
8220            Next
8230            .Close
8240          End With
8250          Set rst = Nothing
8260          frm.ColHeaders True  ' ** Form Procedure: frmLocation_Asset_Sub.
8270        Else
8280          frm.ColHeaders False  ' ** Form Procedure: frmLocation_Asset_Sub.
8290        End If
8300        frm.Refresh
8310      ElseIf lngRecsCur > 0& Then
8320        frm.ColHeaders True  ' ** Form Procedure: frmLocation_Asset_Sub.
8330      Else
8340        frm.ColHeaders False  ' ** Form Procedure: frmLocation_Asset_Sub.
8350      End If

8360      ResetValidate  ' ** Function: Below.

8370      blnFocusCurrent = True
8380      .TimerInterval = 250&

8390      DoCmd.Hourglass False

8400    End With

EXITP:
8410    Set rst = Nothing
8420    Set qdf = Nothing
8430    Set dbs = Nothing
8440    Set frm = Nothing
8450    Exit Sub

ERRH:
8460    DoCmd.Hourglass False
8470    Select Case ERR.Number
        Case Else
8480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8490    End Select
8500    Resume EXITP

End Sub

Private Sub cmbLocations_Current_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_Current_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8620    intShiftDown = (Shift And acShiftMask) > 0
8630    intAltDown = (Shift And acAltMask) > 0
8640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8670        With Me
8680          intRetVal = 0
8690          .cmbLocations_New.SetFocus
8700        End With
8710      End Select
8720    End If

        ' ** Shift keys.
8730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8750        With Me
8760          intRetVal = 0
8770          .cmdClose.SetFocus
8780        End With
8790      End Select
8800    End If

EXITP:
8810    KeyCode = intRetVal
8820    Exit Sub

ERRH:
8830    intRetVal = 0
8840    Select Case ERR.Number
        Case Else
8850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8860    End Select
8870    Resume EXITP

End Sub

Private Sub cmbLocations_New_AfterUpdate()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_New_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim frm As Access.Form
        Dim lngLocID As Long, strLocName As String
        Dim lngRecs As Long
        Dim strFilter As String, strQryName As String, strSQL As String
        Dim intPos01 As Integer
        Dim lngX As Long

8910    With Me

8920      DoCmd.Hourglass True
8930      DoEvents

8940      If Validate = True Then  ' ** Function: Below.

8950        DoEvents

8960        Set frm = .frmLocation_Asset_Sub.Form
8970        frm.cmdSave_Click  ' ** Form Function: frmLocation_Asset_Sub.
8980        DoEvents

8990        lngLocID = .cmbLocations_New.Column(0)
9000        strLocName = .cmbLocations_New.Column(1)
9010        strFilter = vbNullString

9020        Set dbs = CurrentDb
9030        With dbs
              ' ** Update tmpLocation_Asset, for locasset_mark = False.
9040          Set qdf = .QueryDefs("qryLocation_Asset_06b")
9050          qdf.Execute
9060          Set qdf = Nothing
9070          .Close
9080        End With
9090        Set dbs = Nothing
9100        DoEvents

9110        lngRecsCur = frm.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
9120        If lngRecsCur > 0& Then
9130          DoEvents
9140          frm.Refresh
9150          DoEvents
9160          strFilter = frm.Filter
              ' ** tmpLocation_Asset, with or without form Filter.
9170          strQryName = "qryLocation_Asset_Sub_05"
9180          Set dbs = CurrentDb
9190          Set qdf = dbs.QueryDefs(strQryName)
9200          If strFilter <> vbNullString Then
                ' ** WHERE (((tmpLocation_Asset.Location_ID)=6));
9210            strSQL = qdf.SQL
9220            intPos01 = InStr(strSQL, "WHERE ")
9230            If intPos01 > 0 Then
9240              strSQL = Trim(Left(strSQL, (intPos01 - 1)))
9250            Else
9260              strSQL = Left(strSQL, (InStr(strSQL, ";") - 1))
9270            End If
9280            strSQL = strSQL & vbCrLf & "WHERE (" & strFilter & ");"
9290            qdf.SQL = strSQL
9300          End If
              'Set rst = frm.RecordsetClone  ' ** Errors!
9310          Set rst = qdf.OpenRecordset
9320          With rst
9330            .MoveLast
9340            lngRecs = .RecordCount
9350            .MoveFirst
9360            For lngX = 1& To lngRecs
9370              .Edit
9380              ![locasset_mark] = True
9390              ![Location_ID2] = lngLocID
9400              ![Location_Name2] = strLocName
9410              ![locasset_datemodified] = Now()
9420              .Update
9430              If lngX < lngRecs Then .MoveNext
9440            Next
9450            .Close
9460          End With
9470          Set rst = Nothing
9480          frm.Refresh
9490          frm.ColHeaders True  ' ** Form Procedure: frmLocation_Asset_Sub.
9500        Else
9510          frm.ColHeaders False  ' ** Form Procedure: frmLocation_Asset_Sub.
9520        End If

9530        ResetValidate  ' ** Function: Below.

9540        blnFocusNew = True
9550        .TimerInterval = 250&

9560      End If

9570      DoCmd.Hourglass False

9580    End With

EXITP:
9590    Set frm = Nothing
9600    Set rst = Nothing
9610    Set qdf = Nothing
9620    Set dbs = Nothing
9630    Exit Sub

ERRH:
9640    DoCmd.Hourglass False
9650    Select Case ERR.Number
        Case Else
9660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9670    End Select
9680    Resume EXITP

End Sub

Private Sub cmbLocations_New_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_New_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9720    intShiftDown = (Shift And acShiftMask) > 0
9730    intAltDown = (Shift And acAltMask) > 0
9740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9770        With Me
9780          intRetVal = 0
9790          .opgGroup.SetFocus
9800        End With
9810      End Select
9820    End If

        ' ** Shift keys.
9830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9850        With Me
9860          intRetVal = 0
9870          .cmbLocations_Current.SetFocus
9880        End With
9890      End Select
9900    End If

EXITP:
9910    KeyCode = intRetVal
9920    Exit Sub

ERRH:
9930    intRetVal = 0
9940    Select Case ERR.Number
        Case Else
9950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9960    End Select
9970    Resume EXITP

End Sub

Private Sub opgGroup_AfterUpdate()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGroup_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim frm As Access.Form
        Dim lngLocID As Long, strLocName As String, lngAssetNo As Long
        Dim strAccountNo As String, strFilter As String
        Dim lngRecs As Long
        Dim lngX As Long

10010   With Me

10020     DoCmd.Hourglass True
10030     DoEvents

10040     Set dbs = CurrentDb
10050     With dbs
            ' ** Update tmpLocation_Asset, for locasset_mark = False.
10060       Set qdf = .QueryDefs("qryLocation_Asset_06b")
10070       qdf.Execute
10080       Set qdf = Nothing
10090       .Close
10100     End With
10110     Set dbs = Nothing

10120     Set frm = .frmLocation_Asset_Sub.Form

10130     Select Case .opgGroup
          Case .opgGroup_optAll.OptionValue
10140       .opgGroup_optAll_lbl.FontBold = True
10150       .opgGroup_optAccount_lbl.FontBold = False
10160       .opgGroup_optAsset_lbl.FontBold = False
10170       .opgGroup_optAccountAsset_lbl.FontBold = False
10180       .cmbAccounts.Enabled = False
10190       .cmbAccounts.BackStyle = acBackStyleTransparent
10200       .cmbAccounts.BorderColor = WIN_CLR_DISR
10210       .opgAccountSource.Enabled = False
10220       .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
10230       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
10240       .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
10250       .opgAccountSource_optName_lbl2_dim_hi.Visible = True
10260       .cmbAssets.Enabled = False
10270       .cmbAssets.BackStyle = acBackStyleTransparent
10280       .cmbAssets.BorderColor = WIN_CLR_DISR
10290       .opgAssetSource.Enabled = False
10300       .opgAssetSource_optType_lbl2.ForeColor = WIN_CLR_DISF
10310       .opgAssetSource_optType_lbl2_dim_hi.Visible = True
10320       .opgAssetSource_optName_lbl2.ForeColor = WIN_CLR_DISF
10330       .opgAssetSource_optName_lbl2_dim_hi.Visible = True
10340       .opgAssetSource_optCusip_lbl2.ForeColor = WIN_CLR_DISF
10350       .opgAssetSource_optCusip_lbl2_dim_hi.Visible = True
10360       Select Case IsNull(.cmbLocations_Current)
            Case True
10370         strFilter = "[Location_ID] = 0"
10380         frm.Filter = strFilter
10390         frm.FilterOn = True
10400       Case False
10410         lngLocID = .cmbLocations_Current
10420         strFilter = "[Location_ID] = " & CStr(lngLocID)
10430         frm.Filter = strFilter
10440         frm.FilterOn = True
10450       End Select
10460     Case .opgGroup_optAccount.OptionValue
10470       .opgGroup_optAll_lbl.FontBold = False
10480       .opgGroup_optAccount_lbl.FontBold = True
10490       .opgGroup_optAsset_lbl.FontBold = False
10500       .opgGroup_optAccountAsset_lbl.FontBold = False
10510       .cmbAccounts.Enabled = True
10520       .cmbAccounts.BackStyle = acBackStyleNormal
10530       .cmbAccounts.BorderColor = CLR_LTBLU2
10540       .opgAccountSource.Enabled = True
10550       .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
10560       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
10570       .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
10580       .opgAccountSource_optName_lbl2_dim_hi.Visible = False
10590       .cmbAssets.Enabled = False
10600       .cmbAssets.BackStyle = acBackStyleTransparent
10610       .cmbAssets.BorderColor = WIN_CLR_DISR
10620       .opgAssetSource.Enabled = False
10630       .opgAssetSource_optType_lbl2.ForeColor = WIN_CLR_DISF
10640       .opgAssetSource_optType_lbl2_dim_hi.Visible = True
10650       .opgAssetSource_optName_lbl2.ForeColor = WIN_CLR_DISF
10660       .opgAssetSource_optName_lbl2_dim_hi.Visible = True
10670       .opgAssetSource_optCusip_lbl2.ForeColor = WIN_CLR_DISF
10680       .opgAssetSource_optCusip_lbl2_dim_hi.Visible = True
10690       Select Case IsNull(.cmbLocations_Current)
            Case True
10700         strFilter = "[Location_ID] = 0"
10710         frm.Filter = strFilter
10720         frm.FilterOn = True
10730       Case False
10740         lngLocID = .cmbLocations_Current
10750         Select Case IsNull(.cmbAccounts)
              Case True
10760           strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [accountno] = '0'"
10770           frm.Filter = strFilter
10780           frm.FilterOn = True
10790         Case False
10800           strAccountNo = .cmbAccounts
10810           strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [accountno] = '" & strAccountNo & "'"
10820           frm.Filter = strFilter
10830           frm.FilterOn = True
10840         End Select
10850       End Select
10860     Case .opgGroup_optAsset.OptionValue
10870       .opgGroup_optAll_lbl.FontBold = False
10880       .opgGroup_optAccount_lbl.FontBold = False
10890       .opgGroup_optAsset_lbl.FontBold = True
10900       .opgGroup_optAccountAsset_lbl.FontBold = False
10910       .cmbAccounts.Enabled = False
10920       .cmbAccounts.BackStyle = acBackStyleTransparent
10930       .cmbAccounts.BorderColor = WIN_CLR_DISR
10940       .opgAccountSource.Enabled = False
10950       .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
10960       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
10970       .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
10980       .opgAccountSource_optName_lbl2_dim_hi.Visible = True
10990       .cmbAssets.Enabled = True
11000       .cmbAssets.BackStyle = acBackStyleNormal
11010       .cmbAssets.BorderColor = CLR_LTBLU2
11020       .opgAssetSource.Enabled = True
11030       .opgAssetSource_optType_lbl2.ForeColor = CLR_VDKGRY
11040       .opgAssetSource_optType_lbl2_dim_hi.Visible = False
11050       .opgAssetSource_optName_lbl2.ForeColor = CLR_VDKGRY
11060       .opgAssetSource_optName_lbl2_dim_hi.Visible = False
11070       .opgAssetSource_optCusip_lbl2.ForeColor = CLR_VDKGRY
11080       .opgAssetSource_optCusip_lbl2_dim_hi.Visible = False
11090       Select Case IsNull(.cmbLocations_Current)
            Case True
11100         strFilter = "[Location_ID] = 0"
11110         frm.Filter = strFilter
11120         frm.FilterOn = True
11130       Case False
11140         lngLocID = .cmbLocations_Current
11150         Select Case IsNull(.cmbAssets)
              Case True
11160           strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [assetno] = 0"
11170           frm.Filter = strFilter
11180           frm.FilterOn = True
11190         Case False
11200           lngAssetNo = .cmbAssets
11210           strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [assetno] = " & CStr(lngAssetNo)
11220           frm.Filter = strFilter
11230           frm.FilterOn = True
11240         End Select
11250       End Select
11260     Case .opgGroup_optAccountAsset.OptionValue
11270       .opgGroup_optAll_lbl.FontBold = False
11280       .opgGroup_optAccount_lbl.FontBold = False
11290       .opgGroup_optAsset_lbl.FontBold = False
11300       .opgGroup_optAccountAsset_lbl.FontBold = True
11310       .cmbAccounts.Enabled = True
11320       .cmbAccounts.BackStyle = acBackStyleNormal
11330       .cmbAccounts.BorderColor = CLR_LTBLU2
11340       .opgAccountSource.Enabled = True
11350       .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
11360       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
11370       .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
11380       .opgAccountSource_optName_lbl2_dim_hi.Visible = False
11390       .cmbAssets.Enabled = True
11400       .cmbAssets.BackStyle = acBackStyleNormal
11410       .cmbAssets.BorderColor = CLR_LTBLU2
11420       .opgAssetSource.Enabled = True
11430       .opgAssetSource_optType_lbl2.ForeColor = CLR_VDKGRY
11440       .opgAssetSource_optType_lbl2_dim_hi.Visible = False
11450       .opgAssetSource_optName_lbl2.ForeColor = CLR_VDKGRY
11460       .opgAssetSource_optName_lbl2_dim_hi.Visible = False
11470       .opgAssetSource_optCusip_lbl2.ForeColor = CLR_VDKGRY
11480       .opgAssetSource_optCusip_lbl2_dim_hi.Visible = False
11490       Select Case IsNull(.cmbLocations_Current)
            Case True
11500         strFilter = "[Location_ID] = 0"
11510         frm.Filter = strFilter
11520         frm.FilterOn = True
11530       Case False
11540         lngLocID = .cmbLocations_Current
11550         Select Case IsNull(.cmbAccounts)
              Case True
11560           strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [accountno] = '0'"
11570         Case False
11580           strAccountNo = .cmbAccounts
11590           strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [accountno] = '" & strAccountNo & "'"
11600         End Select
11610         Select Case IsNull(.cmbAssets)
              Case True
11620           strFilter = strFilter & " And [assetno] = 0"
11630           frm.Filter = strFilter
11640           frm.FilterOn = True
11650         Case False
11660           lngAssetNo = .cmbAssets
11670           strFilter = strFilter & " And [assetno] = " & CStr(lngAssetNo)
11680           frm.Filter = strFilter
11690           frm.FilterOn = True
11700         End Select
11710       End Select
11720     End Select

11730     Loc_Counts  ' ** Function: Below.
11740     DoEvents

11750     Validate  ' ** Function: Below.
11760     DoEvents

11770     lngRecsCur = frm.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
11780     If lngRecsCur > 0& Then
11790       If lngLocID > 0& Then
11800         If IsNull(.cmbLocations_New) = False Then
11810           lngLocID = .cmbLocations_New.Column(0)
11820           strLocName = .cmbLocations_New.Column(1)
11830           Set rst = frm.RecordsetClone
11840           With rst
11850             .MoveLast
11860             lngRecs = .RecordCount
11870             .MoveFirst
11880             For lngX = 1& To lngRecs
11890               .Edit
11900               ![locasset_mark] = True
11910               ![Location_ID2] = lngLocID
11920               ![Location_Name2] = strLocName
11930               ![locasset_datemodified] = Now()
11940               .Update
11950               If lngX < lngRecs Then .MoveNext
11960             Next
11970             .Close
11980           End With
11990           Set rst = Nothing
12000           frm.Refresh
12010         End If
12020       End If
12030       frm.ColHeaders True  ' ** Form Procedure: frmLocation_Asset_Sub.
12040     Else
12050       frm.ColHeaders False  ' ** Form Procedure: frmLocation_Asset_Sub.
12060     End If

12070     ResetValidate  ' ** Function: Below.

12080     lngRecsCur = frm.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
12090     If lngRecsCur > 0& Then
12100       .frmLocation_Asset_Sub.SetFocus
12110     Else
12120       If .opgGroup = .opgGroup_optAccount.OptionValue And .cmbAccounts.Enabled = True Then
12130         .cmbAccounts.SetFocus
12140       ElseIf .opgGroup = .opgGroup_optAsset.OptionValue And .cmbAssets.Enabled = True Then
12150         .cmbAssets.SetFocus
12160       ElseIf .opgGroup = .opgGroup_optAccountAsset.OptionValue And .cmbAccounts.Enabled = True Then
12170         .cmbAccounts.SetFocus
12180       End If
12190     End If
12200     .frmLocation_Asset_Sub.Form.RecalcTots_Set True  ' ** Form Procedure: frmLocation_Asset_Sub.
12210     DoEvents

12220     DoCmd.Hourglass False

12230   End With

EXITP:
12240   Set frm = Nothing
12250   Set rst = Nothing
12260   Set qdf = Nothing
12270   Set dbs = Nothing
12280   Exit Sub

ERRH:
12290   DoCmd.Hourglass False
12300   Select Case ERR.Number
        Case Else
12310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12320   End Select
12330   Resume EXITP

End Sub

Private Sub opgGroup_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGroup_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12420   intShiftDown = (Shift And acShiftMask) > 0
12430   intAltDown = (Shift And acAltMask) > 0
12440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12470       With Me
12480         intRetVal = 0
12490         If .cmbAccounts.Enabled = True Then
12500           .cmbAccounts.SetFocus
12510         ElseIf .cmbAssets.Enabled = True Then
12520           .cmbAssets.SetFocus
12530         ElseIf .cmdResetFilter.Enabled = True Then
12540           .cmdResetFilter.SetFocus
12550         Else
12560           lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
12570           If lngRecsCur > 0& Then
12580             .frmLocation_Asset_Sub.SetFocus
12590           Else
12600             If .cmdSelectAll.Enabled = True Then
12610               .cmdSelectAll.SetFocus
12620             ElseIf .cmdUpdate.Enabled = True Then
12630               .cmdUpdate.SetFocus
12640             Else
12650               .cmdClose.SetFocus
12660             End If
12670           End If
12680         End If
12690       End With
12700     End Select
12710   End If

        ' ** Shift keys.
12720   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12730     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12740       With Me
12750         intRetVal = 0
12760         .cmbLocations_New.SetFocus
12770       End With
12780     End Select
12790   End If

        ' ** Ctrl keys.
12800   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
12810     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12820       With Me
12830         intRetVal = 0
12840         If .opgAccountSource.Enabled = True Then
12850           .opgAccountSource.SetFocus
12860         ElseIf .opgAssetSource.Enabled = True Then
12870           .opgAssetSource.SetFocus
12880         Else
12890           Beep
12900         End If
12910       End With
12920     End Select
12930   End If

EXITP:
12940   KeyCode = intRetVal
12950   Exit Sub

ERRH:
12960   intRetVal = 0
12970   Select Case ERR.Number
        Case Else
12980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12990   End Select
13000   Resume EXITP

End Sub

Private Sub opgGroup_optAccount_KeyDown(KeyCode As Integer, Shift As Integer)

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGroup_optAccount_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13120   intShiftDown = (Shift And acShiftMask) > 0
13130   intAltDown = (Shift And acAltMask) > 0
13140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13170       With Me
13180         intRetVal = 0
13190         If .cmbAccounts.Enabled = True Then
13200           .cmbAccounts.SetFocus
13210         ElseIf .cmbAssets.Enabled = True Then
13220           .cmbAssets.SetFocus
13230         ElseIf .cmdResetFilter.Enabled = True Then
13240           .cmdResetFilter.SetFocus
13250         Else
13260           lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
13270           If lngRecsCur > 0& Then
13280             .frmLocation_Asset_Sub.SetFocus
13290           Else
13300             If .cmdSelectAll.Enabled = True Then
13310               .cmdSelectAll.SetFocus
13320             ElseIf .cmdUpdate.Enabled = True Then
13330               .cmdUpdate.SetFocus
13340             Else
13350               .cmdClose.SetFocus
13360             End If
13370           End If
13380         End If
13390       End With
13400     End Select
13410   End If

        ' ** Shift keys.
13420   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13430     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13440       With Me
13450         intRetVal = 0
13460         .cmbLocations_New.SetFocus
13470       End With
13480     End Select
13490   End If

        ' ** Ctrl keys.
13500   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
13510     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13520       With Me
13530         intRetVal = 0
13540         If .opgAccountSource.Enabled = True Then
13550           .opgAccountSource.SetFocus
13560         ElseIf .opgAssetSource.Enabled = True Then
13570           .opgAssetSource.SetFocus
13580         Else
13590           Beep
13600         End If
13610       End With
13620     End Select
13630   End If

EXITP:
13640   KeyCode = intRetVal
13650   Exit Sub

ERRH:
13660   intRetVal = 0
13670   Select Case ERR.Number
        Case Else
13680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13690   End Select
13700   Resume EXITP

End Sub

Private Sub opgGroup_optAsset_KeyDown(KeyCode As Integer, Shift As Integer)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGroup_optAsset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13820   intShiftDown = (Shift And acShiftMask) > 0
13830   intAltDown = (Shift And acAltMask) > 0
13840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13870       With Me
13880         intRetVal = 0
13890         If .cmbAccounts.Enabled = True Then
13900           .cmbAccounts.SetFocus
13910         ElseIf .cmbAssets.Enabled = True Then
13920           .cmbAssets.SetFocus
13930         ElseIf .cmdResetFilter.Enabled = True Then
13940           .cmdResetFilter.SetFocus
13950         Else
13960           lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
13970           If lngRecsCur > 0& Then
13980             .frmLocation_Asset_Sub.SetFocus
13990           Else
14000             If .cmdSelectAll.Enabled = True Then
14010               .cmdSelectAll.SetFocus
14020             ElseIf .cmdUpdate.Enabled = True Then
14030               .cmdUpdate.SetFocus
14040             Else
14050               .cmdClose.SetFocus
14060             End If
14070           End If
14080         End If
14090       End With
14100     End Select
14110   End If

        ' ** Shift keys.
14120   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14130     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14140       With Me
14150         intRetVal = 0
14160         .cmbLocations_New.SetFocus
14170       End With
14180     End Select
14190   End If

        ' ** Ctrl keys.
14200   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
14210     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14220       With Me
14230         intRetVal = 0
14240         If .opgAccountSource.Enabled = True Then
14250           .opgAccountSource.SetFocus
14260         ElseIf .opgAssetSource.Enabled = True Then
14270           .opgAssetSource.SetFocus
14280         Else
14290           Beep
14300         End If
14310       End With
14320     End Select
14330   End If

EXITP:
14340   KeyCode = intRetVal
14350   Exit Sub

ERRH:
14360   intRetVal = 0
14370   Select Case ERR.Number
        Case Else
14380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14390   End Select
14400   Resume EXITP

End Sub

Private Sub opgGroup_optAccountAsset_KeyDown(KeyCode As Integer, Shift As Integer)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGroup_optAccountAsset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14520   intShiftDown = (Shift And acShiftMask) > 0
14530   intAltDown = (Shift And acAltMask) > 0
14540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14570       With Me
14580         intRetVal = 0
14590         If .cmbAccounts.Enabled = True Then
14600           .cmbAccounts.SetFocus
14610         ElseIf .cmbAssets.Enabled = True Then
14620           .cmbAssets.SetFocus
14630         ElseIf .cmdResetFilter.Enabled = True Then
14640           .cmdResetFilter.SetFocus
14650         Else
14660           lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
14670           If lngRecsCur > 0& Then
14680             .frmLocation_Asset_Sub.SetFocus
14690           Else
14700             If .cmdSelectAll.Enabled = True Then
14710               .cmdSelectAll.SetFocus
14720             ElseIf .cmdUpdate.Enabled = True Then
14730               .cmdUpdate.SetFocus
14740             Else
14750               .cmdClose.SetFocus
14760             End If
14770           End If
14780         End If
14790       End With
14800     End Select
14810   End If

        ' ** Shift keys.
14820   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14830     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14840       With Me
14850         intRetVal = 0
14860         .cmbLocations_New.SetFocus
14870       End With
14880     End Select
14890   End If

        ' ** Ctrl keys.
14900   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
14910     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14920       With Me
14930         intRetVal = 0
14940         If .opgAccountSource.Enabled = True Then
14950           .opgAccountSource.SetFocus
14960         ElseIf .opgAssetSource.Enabled = True Then
14970           .opgAssetSource.SetFocus
14980         Else
14990           Beep
15000         End If
15010       End With
15020     End Select
15030   End If

EXITP:
15040   KeyCode = intRetVal
15050   Exit Sub

ERRH:
15060   intRetVal = 0
15070   Select Case ERR.Number
        Case Else
15080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15090   End Select
15100   Resume EXITP

End Sub

Public Sub opgAccountSource_AfterUpdate()

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

15210   With Me
15220     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
15230       .opgAccountSource_optNumber_lbl2.FontBold = True
15240       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
15250       .opgAccountSource_optName_lbl2.FontBold = False
15260       .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
15270       If .cmbAccounts.RowSource <> "qryLocation_Asset_03a" Then
15280         .cmbAccounts.RowSource = "qryLocation_Asset_03a"
15290       End If
15300     Case .opgAccountSource_optName.OptionValue
15310       .opgAccountSource_optNumber_lbl2.FontBold = False
15320       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
15330       .opgAccountSource_optName_lbl2.FontBold = True
15340       .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
15350       If .cmbAccounts.RowSource <> "qryLocation_Asset_03b" Then
15360         .cmbAccounts.RowSource = "qryLocation_Asset_03b"
15370       End If
15380     End Select
15390   End With

EXITP:
15400   Exit Sub

ERRH:
15410   Select Case ERR.Number
        Case Else
15420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15430   End Select
15440   Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15520   intShiftDown = (Shift And acShiftMask) > 0
15530   intAltDown = (Shift And acAltMask) > 0
15540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15570       With Me
15580         intRetVal = 0
15590         .cmbAccounts.SetFocus
15600       End With
15610     End Select
15620   End If

        ' ** Shift keys.
15630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15650       With Me
15660         intRetVal = 0
15670         .opgGroup.SetFocus
15680       End With
15690     End Select
15700   End If

EXITP:
15710   KeyCode = intRetVal
15720   Exit Sub

ERRH:
15730   intRetVal = 0
15740   Select Case ERR.Number
        Case Else
15750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15760   End Select
15770   Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15820   intShiftDown = (Shift And acShiftMask) > 0
15830   intAltDown = (Shift And acAltMask) > 0
15840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15870       With Me
15880         intRetVal = 0
15890         .cmbAccounts.SetFocus
15900       End With
15910     End Select
15920   End If

        ' ** Shift keys.
15930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15950       With Me
15960         intRetVal = 0
15970         .opgGroup.SetFocus
15980       End With
15990     End Select
16000   End If

EXITP:
16010   KeyCode = intRetVal
16020   Exit Sub

ERRH:
16030   intRetVal = 0
16040   Select Case ERR.Number
        Case Else
16050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16060   End Select
16070   Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

16110   With Me
16120     If .cmbAccounts.text = vbNullString Then
16130       ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
16140     End If
16150   End With

EXITP:
16160   Exit Sub

ERRH:
16170   Select Case ERR.Number
        Case Else
16180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16190   End Select
16200   Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim frm As Access.Form
        Dim lngLocID As Long, strLocName As String, lngAssetNo As Long
        Dim strAccountNo As String, strFilter As String
        Dim lngRecs As Long
        Dim lngX As Long

16310   With Me

16320     DoCmd.Hourglass True
16330     DoEvents

16340     Set dbs = CurrentDb
16350     With dbs
            ' ** Update tmpLocation_Asset, for locasset_mark = False.
16360       Set qdf = .QueryDefs("qryLocation_Asset_06b")
16370       qdf.Execute
16380       Set qdf = Nothing
16390       .Close
16400     End With
16410     Set dbs = Nothing

16420     .Cnt_AA.Locked = False
16430     .Cnt_AA = 0&
16440     .Cnt_AA.BackStyle = acBackStyleTransparent
16450     .Cnt_AA.BorderColor = WIN_CLR_DISR
16460     .Cnt_L.Visible = False
16470     .Cnt_L = 0&
16480     .Cnt_LA.Visible = False
16490     .Cnt_LA = 0&
16500     .Cnt_J.Visible = False
16510     .Cnt_J = 0&
16520     .Cnt_JC.Visible = False
16530     .Cnt_JC = 0&
16540     DoEvents

16550     Set frm = .frmLocation_Asset_Sub.Form

16560     Select Case IsNull(.cmbLocations_Current)
          Case True
16570       lngLocID = 0&
16580     Case False
16590       lngLocID = .cmbLocations_Current
16600     End Select

16610     Select Case IsNull(.cmbAccounts)
          Case True
16620       strAccountNo = "0"
16630     Case False
16640       strAccountNo = .cmbAccounts
16650     End Select

16660     Select Case .opgGroup
          Case .opgGroup_optAccountAsset.OptionValue
16670       Select Case IsNull(.cmbAssets)
            Case True
16680         lngAssetNo = 0&
16690       Case False
16700         lngAssetNo = .cmbAssets
16710       End Select
16720       strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [accountno] = '" & strAccountNo & "' And [assetno] = " & CStr(lngAssetNo)
16730     Case Else
16740       strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [accountno] = '" & strAccountNo & "'"
16750     End Select
16760     frm.Filter = strFilter
16770     frm.FilterOn = True

16780     Loc_Counts  ' ** Function: Below.
16790     DoEvents

16800     Validate  ' ** Function: Below.
16810     DoEvents

16820     lngRecsCur = frm.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
16830     If lngRecsCur > 0& Then
16840       If lngLocID > 0& Then
16850         If IsNull(.cmbLocations_New) = False Then
16860           lngLocID = .cmbLocations_New.Column(0)
16870           strLocName = .cmbLocations_New.Column(1)
16880           Set rst = frm.RecordsetClone
16890           With rst
16900             .MoveLast
16910             lngRecs = .RecordCount
16920             .MoveFirst
16930             For lngX = 1& To lngRecs
16940               .Edit
16950               ![locasset_mark] = True
16960               ![Location_ID2] = lngLocID
16970               ![Location_Name2] = strLocName
16980               ![locasset_datemodified] = Now()
16990               .Update
17000               If lngX < lngRecs Then .MoveNext
17010             Next
17020             .Close
17030           End With
17040           Set rst = Nothing
17050           frm.Refresh
17060         End If
17070       End If
17080       frm.ColHeaders True  ' ** Form Procedure: frmLocation_Asset_Sub.
17090     Else
17100       frm.ColHeaders False  ' ** Form Procedure: frmLocation_Asset_Sub.
17110     End If

17120     ResetValidate  ' ** Function: Below.

17130     frm.RecalcTots_Set True  ' ** Form Procedure: frmLocation_Asset_Sub.

17140     DoCmd.Hourglass False

17150   End With

EXITP:
17160   Set frm = Nothing
17170   Set rst = Nothing
17180   Set qdf = Nothing
17190   Set dbs = Nothing
17200   Exit Sub

ERRH:
17210   DoCmd.Hourglass False
17220   Select Case ERR.Number
        Case Else
17230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17240   End Select
17250   Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17320   intShiftDown = (Shift And acShiftMask) > 0
17330   intAltDown = (Shift And acAltMask) > 0
17340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17370       With Me
17380         intRetVal = 0
17390         If .cmbAssets.Enabled = True Then
17400           .cmbAssets.SetFocus
17410         ElseIf .cmdResetFilter.Enabled = True Then
17420           .cmdResetFilter.SetFocus
17430         Else
17440           lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
17450           If lngRecsCur > 0& Then
17460             .frmLocation_Asset_Sub.SetFocus
17470           Else
17480             If .cmdSelectAll.Enabled = True Then
17490               .cmdSelectAll.SetFocus
17500             ElseIf .cmdUpdate.Enabled = True Then
17510               .cmdUpdate.SetFocus
17520             Else
17530               .cmdClose.SetFocus
17540             End If
17550           End If
17560         End If
17570       End With
17580     End Select
17590   End If

        ' ** Shift keys.
17600   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17610     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17620       With Me
17630         intRetVal = 0
17640         .opgGroup.SetFocus
17650       End With
17660     End Select
17670   End If

        ' ** Ctrl keys.
17680   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
17690     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17700       With Me
17710         intRetVal = 0
17720         If .opgAssetSource.Enabled = True Then
17730           .opgAssetSource.SetFocus
17740         Else
17750           Beep
17760         End If
17770       End With
17780     End Select
17790   End If

        ' ** Ctrl-Shift keys.
17800   If intCtrlDown And (Not intAltDown) And intShiftDown Then
17810     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17820       With Me
17830         intRetVal = 0
17840         .opgAccountSource.SetFocus
17850       End With
17860     End Select
17870   End If

EXITP:
17880   KeyCode = intRetVal
17890   Exit Sub

ERRH:
17900   intRetVal = 0
17910   Select Case ERR.Number
        Case Else
17920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17930   End Select
17940   Resume EXITP

End Sub

Public Sub opgAssetSource_AfterUpdate()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_AfterUpdate"

18010   With Me
18020     Select Case .opgAssetSource
          Case .opgAssetSource_optType.OptionValue
18030       .opgAssetSource_optType_lbl2.FontBold = True
18040       .opgAssetSource_optType_lbl2_dim_hi.FontBold = True
18050       .opgAssetSource_optName_lbl2.FontBold = False
18060       .opgAssetSource_optName_lbl2_dim_hi.FontBold = False
18070       .opgAssetSource_optCusip_lbl2.FontBold = False
18080       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = False
18090       If .cmbAssets.RowSource <> "qryLocation_Asset_04a" Then
18100         .cmbAssets.RowSource = "qryLocation_Asset_04a"
18110       End If
            ' ** These should perhaps be translated using lngTpp!
18120       .cmbAssets.ColumnWidths = "0;4620;1185;0;360"
18130       .cmbAssets.ListWidth = 6420
18140     Case .opgAssetSource_optName.OptionValue
18150       .opgAssetSource_optType_lbl2.FontBold = False
18160       .opgAssetSource_optType_lbl2_dim_hi.FontBold = False
18170       .opgAssetSource_optName_lbl2.FontBold = True
18180       .opgAssetSource_optName_lbl2_dim_hi.FontBold = True
18190       .opgAssetSource_optCusip_lbl2.FontBold = False
18200       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = False
18210       If .cmbAssets.RowSource <> "qryLocation_Asset_04b" Then
18220         .cmbAssets.RowSource = "qryLocation_Asset_04b"
18230       End If
            ' ** These should perhaps be translated using lngTpp!
18240       .cmbAssets.ColumnWidths = "0;4620;1185;0;360"
18250       .cmbAssets.ListWidth = 6420
18260     Case .opgAssetSource_optCusip.OptionValue
18270       .opgAssetSource_optType_lbl2.FontBold = False
18280       .opgAssetSource_optType_lbl2_dim_hi.FontBold = False
18290       .opgAssetSource_optName_lbl2.FontBold = False
18300       .opgAssetSource_optName_lbl2_dim_hi.FontBold = False
18310       .opgAssetSource_optCusip_lbl2.FontBold = True
18320       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = True
18330       If .cmbAssets.RowSource <> "qryLocation_Asset_04c" Then
18340         .cmbAssets.RowSource = "qryLocation_Asset_04c"
18350       End If
            ' ** These should perhaps be translated using lngTpp!
18360       .cmbAssets.ColumnWidths = "0;5940;1185;0;360"
18370       .cmbAssets.ListWidth = 7740
18380     End Select
18390   End With

EXITP:
18400   Exit Sub

ERRH:
18410   Select Case ERR.Number
        Case Else
18420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18430   End Select
18440   Resume EXITP

End Sub

Private Sub opgAssetSource_optType_KeyDown(KeyCode As Integer, Shift As Integer)

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18520   intShiftDown = (Shift And acShiftMask) > 0
18530   intAltDown = (Shift And acAltMask) > 0
18540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18570       With Me
18580         intRetVal = 0
18590         .cmbAssets.SetFocus
18600       End With
18610     End Select
18620   End If

        ' ** Shift keys.
18630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18650       With Me
18660         intRetVal = 0
18670         If .cmbAccounts.Enabled = True Then
18680           .cmbAccounts.SetFocus
18690         Else
18700           .opgGroup.SetFocus
18710         End If
18720       End With
18730     End Select
18740   End If

EXITP:
18750   KeyCode = intRetVal
18760   Exit Sub

ERRH:
18770   intRetVal = 0
18780   Select Case ERR.Number
        Case Else
18790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18800   End Select
18810   Resume EXITP

End Sub

Private Sub opgAssetSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18920   intShiftDown = (Shift And acShiftMask) > 0
18930   intAltDown = (Shift And acAltMask) > 0
18940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18970       With Me
18980         intRetVal = 0
18990         .cmbAssets.SetFocus
19000       End With
19010     End Select
19020   End If

        ' ** Shift keys.
19030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19050       With Me
19060         intRetVal = 0
19070         If .cmbAccounts.Enabled = True Then
19080           .cmbAccounts.SetFocus
19090         Else
19100           .opgGroup.SetFocus
19110         End If
19120       End With
19130     End Select
19140   End If

EXITP:
19150   KeyCode = intRetVal
19160   Exit Sub

ERRH:
19170   intRetVal = 0
19180   Select Case ERR.Number
        Case Else
19190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19200   End Select
19210   Resume EXITP

End Sub

Private Sub opgAssetSource_optCusip_KeyDown(KeyCode As Integer, Shift As Integer)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optCusip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19320   intShiftDown = (Shift And acShiftMask) > 0
19330   intAltDown = (Shift And acAltMask) > 0
19340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19370       With Me
19380         intRetVal = 0
19390         .cmbAssets.SetFocus
19400       End With
19410     End Select
19420   End If

        ' ** Shift keys.
19430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19450       With Me
19460         intRetVal = 0
19470         If .cmbAccounts.Enabled = True Then
19480           .cmbAccounts.SetFocus
19490         Else
19500           .opgGroup.SetFocus
19510         End If
19520       End With
19530     End Select
19540   End If

EXITP:
19550   KeyCode = intRetVal
19560   Exit Sub

ERRH:
19570   intRetVal = 0
19580   Select Case ERR.Number
        Case Else
19590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19600   End Select
19610   Resume EXITP

End Sub

Private Sub cmbAssets_AfterUpdate()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim frm As Access.Form
        Dim strAccountNo As String, lngLocID As Long, strLocName As String, lngAssetNo As Long
        Dim strFilter As String
        Dim lngRecs As Long
        Dim lngX As Long

19710   With Me

19720     DoCmd.Hourglass True
19730     DoEvents

19740     Set dbs = CurrentDb
19750     With dbs
            ' ** Update tmpLocation_Asset, for locasset_mark = False.
19760       Set qdf = .QueryDefs("qryLocation_Asset_06b")
19770       qdf.Execute
19780       Set qdf = Nothing
19790       .Close
19800     End With
19810     Set dbs = Nothing

19820     .Cnt_AA.Locked = False
19830     .Cnt_AA = 0&
19840     .Cnt_AA.BackStyle = acBackStyleTransparent
19850     .Cnt_AA.BorderColor = WIN_CLR_DISR
19860     .Cnt_L.Visible = False
19870     .Cnt_L = 0&
19880     .Cnt_LA.Visible = False
19890     .Cnt_LA = 0&
19900     .Cnt_J.Visible = False
19910     .Cnt_J = 0&
19920     .Cnt_JC.Visible = False
19930     .Cnt_JC = 0&
19940     DoEvents

19950     Set frm = .frmLocation_Asset_Sub.Form
19960     Select Case IsNull(.cmbLocations_Current)
          Case True
19970       lngLocID = 0&
19980     Case False
19990       lngLocID = .cmbLocations_Current
20000     End Select

20010     Select Case IsNull(.cmbAssets)
          Case True
20020       lngAssetNo = 0&
20030     Case False
20040       lngAssetNo = .cmbAssets
20050     End Select

20060     Select Case .opgGroup
          Case .opgGroup_optAccountAsset.OptionValue
20070       Select Case IsNull(.cmbAccounts)
            Case True
20080         strAccountNo = "0"
20090       Case False
20100         strAccountNo = .cmbAccounts
20110       End Select
20120       strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [accountno] = '" & strAccountNo & "' And [assetno] = " & CStr(lngAssetNo)
20130     Case Else
20140       strFilter = "[Location_ID] = " & CStr(lngLocID) & " And [assetno] = " & CStr(lngAssetNo)
20150     End Select
20160     frm.Filter = strFilter
20170     frm.FilterOn = True

20180     Loc_Counts  ' ** Function: Below.
20190     DoEvents

20200     Validate  ' ** Function: Below.
20210     DoEvents

20220     lngRecsCur = frm.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
20230     If lngRecsCur > 0& Then
20240       If lngLocID > 0& Then
20250         If IsNull(.cmbLocations_New) = False Then
20260           lngLocID = .cmbLocations_New.Column(0)
20270           strLocName = .cmbLocations_New.Column(1)
20280           Set rst = frm.RecordsetClone
20290           With rst
20300             .MoveLast
20310             lngRecs = .RecordCount
20320             .MoveFirst
20330             For lngX = 1& To lngRecs
20340               .Edit
20350               ![locasset_mark] = True
20360               ![Location_ID2] = lngLocID
20370               ![Location_Name2] = strLocName
20380               ![locasset_datemodified] = Now()
20390               .Update
20400               If lngX < lngRecs Then .MoveNext
20410             Next
20420             .Close
20430           End With
20440           Set rst = Nothing
20450           frm.Refresh
20460         End If
20470       End If
20480       frm.ColHeaders True  ' ** Form Procedure: frmLocation_Asset_Sub.
20490     Else
20500       frm.ColHeaders False  ' ** Form Procedure: frmLocation_Asset_Sub.
20510     End If

20520     ResetValidate  ' ** Function: Below.

20530     frm.RecalcTots_Set True  ' ** Form Procedure: frmLocation_Asset_Sub.

20540     DoCmd.Hourglass False

20550   End With

EXITP:
20560   Set frm = Nothing
20570   Set rst = Nothing
20580   Set qdf = Nothing
20590   Set dbs = Nothing
20600   Exit Sub

ERRH:
20610   DoCmd.Hourglass False
20620   Select Case ERR.Number
        Case Else
20630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20640   End Select
20650   Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20720   intShiftDown = (Shift And acShiftMask) > 0
20730   intAltDown = (Shift And acAltMask) > 0
20740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20770       With Me
20780         intRetVal = 0
20790         If .cmdResetFilter.Enabled = True Then
20800           .cmdResetFilter.SetFocus
20810         Else
20820           lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
20830           If lngRecsCur > 0& Then
20840             .frmLocation_Asset_Sub.SetFocus
20850           Else
20860             If .cmdSelectAll.Enabled = True Then
20870               .cmdSelectAll.SetFocus
20880             ElseIf .cmdUpdate.Enabled = True Then
20890               .cmdUpdate.SetFocus
20900             Else
20910               .cmdClose.SetFocus
20920             End If
20930           End If
20940         End If
20950       End With
20960     End Select
20970   End If

        ' ** Shift keys.
20980   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20990     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21000       With Me
21010         intRetVal = 0
21020         If .cmbAccounts.Enabled = True Then
21030           .cmbAccounts.SetFocus
21040         Else
21050           .opgGroup.SetFocus
21060         End If
21070       End With
21080     End Select
21090   End If

        ' ** Ctrl-Shift keys.
21100   If intCtrlDown And (Not intAltDown) And intShiftDown Then
21110     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21120       With Me
21130         intRetVal = 0
21140         .opgAssetSource.SetFocus
21150       End With
21160     End Select
21170   End If

EXITP:
21180   KeyCode = intRetVal
21190   Exit Sub

ERRH:
21200   intRetVal = 0
21210   Select Case ERR.Number
        Case Else
21220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21230   End Select
21240   Resume EXITP

End Sub

Private Sub cmdResetFilter_GotFocus()

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_GotFocus"

21310   With Me
21320     blnReset_Focus = True
21330     .cmdResetFilter_raised_semifocus_dots_img.Visible = True
21340     .cmdResetFilter_raised_img.Visible = False
21350     .cmdResetFilter_raised_focus_img.Visible = False
21360     .cmdResetFilter_raised_focus_dots_img.Visible = False
21370     .cmdResetFilter_sunken_focus_dots_img.Visible = False
21380     .cmdResetFilter_raised_img_dis.Visible = False
21390   End With

EXITP:
21400   Exit Sub

ERRH:
21410   Select Case ERR.Number
        Case Else
21420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21430   End Select
21440   Resume EXITP

End Sub

Private Sub cmdResetFilter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_MouseDown"

21510   With Me
21520     blnReset_MouseDown = True
21530     .cmdResetFilter_sunken_focus_dots_img.Visible = True
21540     .cmdResetFilter_raised_img.Visible = False
21550     .cmdResetFilter_raised_semifocus_dots_img.Visible = False
21560     .cmdResetFilter_raised_focus_img.Visible = False
21570     .cmdResetFilter_raised_focus_dots_img.Visible = False
21580     .cmdResetFilter_raised_img_dis.Visible = False
21590   End With

EXITP:
21600   Exit Sub

ERRH:
21610   Select Case ERR.Number
        Case Else
21620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21630   End Select
21640   Resume EXITP

End Sub

Private Sub cmdResetFilter_Click()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_Click"

21710   With Me
21720     .cmbLocations_Current = Null
21730     .cmbLocations_New = Null
21740     .cmbAccounts = Null
21750     .cmbAssets = Null
21760     If .opgGroup <> .opgGroup_optAll.OptionValue Then
21770       .opgGroup = .opgGroup.DefaultValue
21780     End If
21790     opgGroup_AfterUpdate  ' ** Procedure: Above.
21800     DoEvents
21810     ResetValidate  ' ** Function: Below.
21820   End With

EXITP:
21830   Exit Sub

ERRH:
21840   Select Case ERR.Number
        Case Else
21850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21860   End Select
21870   Resume EXITP

End Sub

Private Sub cmdResetFilter_KeyDown(KeyCode As Integer, Shift As Integer)

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21920   intShiftDown = (Shift And acShiftMask) > 0
21930   intAltDown = (Shift And acAltMask) > 0
21940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21970       With Me
21980         intRetVal = 0
21990         lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
22000         If lngRecsCur > 0& Then
22010           .frmLocation_Asset_Sub.SetFocus
22020         Else
22030           If .cmdSelectAll.Enabled = True Then
22040             .cmdSelectAll.SetFocus
22050           ElseIf .cmdUpdate.Enabled = True Then
22060             .cmdUpdate.SetFocus
22070           Else
22080             .cmdClose.SetFocus
22090           End If
22100         End If
22110       End With
22120     End Select
22130   End If

        ' ** Shift keys.
22140   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22150     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22160       With Me
22170         intRetVal = 0
22180         If .cmbAssets.Enabled = True Then
22190           .cmbAssets.SetFocus
22200         ElseIf .cmbAccounts.Enabled = True Then
22210           .cmbAccounts.SetFocus
22220         Else
22230           .opgGroup.SetFocus
22240         End If
22250       End With
22260     End Select
22270   End If

EXITP:
22280   KeyCode = intRetVal
22290   Exit Sub

ERRH:
22300   intRetVal = 0
22310   Select Case ERR.Number
        Case Else
22320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22330   End Select
22340   Resume EXITP

End Sub

Private Sub cmdResetFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_MouseMove"

22410   With Me
22420     If blnReset_MouseDown = False Then
22430       Select Case blnReset_Focus
            Case True
22440         .cmdResetFilter_raised_focus_dots_img.Visible = True
22450         .cmdResetFilter_raised_focus_img.Visible = False
22460       Case False
22470         .cmdResetFilter_raised_focus_img.Visible = True
22480         .cmdResetFilter_raised_focus_dots_img.Visible = False
22490       End Select
22500       .cmdResetFilter_raised_img.Visible = False
22510       .cmdResetFilter_raised_semifocus_dots_img.Visible = False
22520       .cmdResetFilter_sunken_focus_dots_img.Visible = False
22530       .cmdResetFilter_raised_img_dis.Visible = False
22540     End If
22550   End With

EXITP:
22560   Exit Sub

ERRH:
22570   Select Case ERR.Number
        Case Else
22580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22590   End Select
22600   Resume EXITP

End Sub

Private Sub cmdResetFilter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_MouseUp"

22710   With Me
22720     .cmdResetFilter_raised_focus_dots_img.Visible = True
22730     .cmdResetFilter_raised_img.Visible = False
22740     .cmdResetFilter_raised_semifocus_dots_img.Visible = False
22750     .cmdResetFilter_raised_focus_img.Visible = False
22760     .cmdResetFilter_sunken_focus_dots_img.Visible = False
22770     .cmdResetFilter_raised_img_dis.Visible = False
22780     blnReset_MouseDown = False
22790   End With

EXITP:
22800   Exit Sub

ERRH:
22810   Select Case ERR.Number
        Case Else
22820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22830   End Select
22840   Resume EXITP

End Sub

Private Sub cmdResetFilter_LostFocus()

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_LostFocus"

22910   With Me
22920     .cmdResetFilter_raised_img.Visible = True
22930     .cmdResetFilter_raised_semifocus_dots_img.Visible = False
22940     .cmdResetFilter_raised_focus_img.Visible = False
22950     .cmdResetFilter_raised_focus_dots_img.Visible = False
22960     .cmdResetFilter_sunken_focus_dots_img.Visible = False
22970     .cmdResetFilter_raised_img_dis.Visible = False
22980     blnReset_Focus = False
22990   End With

EXITP:
23000   Exit Sub

ERRH:
23010   Select Case ERR.Number
        Case Else
23020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23030   End Select
23040   Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

23110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
23120   Exit Sub

ERRH:
23130   Select Case ERR.Number
        Case Else
23140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23150   End Select
23160   Resume EXITP

End Sub

Private Sub locasset_mark_sum_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "locasset_mark_sum_box_MouseMove"

23210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
23220   Exit Sub

ERRH:
23230   Select Case ERR.Number
        Case Else
23240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23250   End Select
23260   Resume EXITP

End Sub

Private Sub locasset_mark_sum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "locasset_mark_sum_MouseMove"

23310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
23320   Exit Sub

ERRH:
23330   Select Case ERR.Number
        Case Else
23340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23350   End Select
23360   Resume EXITP

End Sub

Private Sub cmdSelectAll_GotFocus()

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_GotFocus"

23410   With Me
23420     blnSelectAll_Focus = True
23430     .cmdSelectAll_raised_semifocus_dots_img.Visible = True
23440     .cmdSelectAll_raised_img.Visible = False
23450     .cmdSelectAll_raised_focus_img.Visible = False
23460     .cmdSelectAll_raised_focus_dots_img.Visible = False
23470     .cmdSelectAll_sunken_focus_dots_img.Visible = False
23480     .cmdSelectAll_raised_img_dis.Visible = False
23490   End With

EXITP:
23500   Exit Sub

ERRH:
23510   Select Case ERR.Number
        Case Else
23520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23530   End Select
23540   Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseDown"

23610   With Me
23620     blnSelectAll_MouseDown = True
23630     .cmdSelectAll_sunken_focus_dots_img.Visible = True
23640     .cmdSelectAll_raised_img.Visible = False
23650     .cmdSelectAll_raised_semifocus_dots_img.Visible = False
23660     .cmdSelectAll_raised_focus_img.Visible = False
23670     .cmdSelectAll_raised_focus_dots_img.Visible = False
23680     .cmdSelectAll_raised_img_dis.Visible = False
23690   End With

EXITP:
23700   Exit Sub

ERRH:
23710   Select Case ERR.Number
        Case Else
23720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23730   End Select
23740   Resume EXITP

End Sub

Private Sub cmdSelectAll_Click()

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngLocID As Long, strLocName As String

23810   With Me

23820     DoCmd.Hourglass True
23830     DoEvents

23840     lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
23850     If lngRecsCur > 0& Then
23860       Select Case IsNull(.cmbLocations_New)
            Case True
23870         Beep
23880       Case False
23890         lngLocID = .cmbLocations_New.Column(0)
23900         strLocName = .cmbLocations_New.Column(1)
23910         Set dbs = CurrentDb
23920         With dbs
                ' ** Update tmpLocation_Asset, for locasset_mark = True, by specified [locnew], [locnam].
23930           Set qdf = .QueryDefs("qryLocation_Asset_06a")
23940           With qdf.Parameters
23950             ![locnew] = lngLocID
23960             ![locnam] = strLocName
23970           End With
23980           qdf.Execute
23990           Set qdf = Nothing
24000           .Close
24010         End With
24020         Set dbs = Nothing
24030         .frmLocation_Asset_Sub.Form.Refresh
24040         DoEvents
24050         .frmLocation_Asset_Sub.Form.RecalcTots_Set True  ' ** Form Procedure: frmLocation_Asset_Sub.
24060       End Select
24070     Else
24080       Beep
24090     End If

24100     lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
24110     If lngRecsCur > 0& Then
24120       .frmLocation_Asset_Sub.SetFocus
24130     End If

24140     DoCmd.Hourglass False

24150   End With

EXITP:
24160   Set qdf = Nothing
24170   Set dbs = Nothing
24180   Exit Sub

ERRH:
24190   DoCmd.Hourglass False
24200   Select Case ERR.Number
        Case Else
24210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24220   End Select
24230   Resume EXITP

End Sub

Private Sub cmdSelectAll_KeyDown(KeyCode As Integer, Shift As Integer)

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24320   intShiftDown = (Shift And acShiftMask) > 0
24330   intAltDown = (Shift And acAltMask) > 0
24340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24370       With Me
24380         intRetVal = 0
24390         .cmdSelectNone.SetFocus
24400       End With
24410     End Select
24420   End If

        ' ** Shift keys.
24430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24450       With Me
24460         intRetVal = 0
24470         lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
24480         If lngRecsCur > 0& Then
24490           .frmLocation_Asset_Sub.SetFocus
24500         Else
24510           If .cmdResetFilter.Enabled = True Then
24520             .cmdResetFilter.SetFocus
24530           ElseIf .cmbAssets.Enabled = True Then
24540             .cmbAssets.SetFocus
24550           ElseIf .cmbAccounts.Enabled = True Then
24560             .cmbAccounts.SetFocus
24570           Else
24580             .opgGroup.SetFocus
24590           End If
24600         End If
24610       End With
24620     End Select
24630   End If

EXITP:
24640   KeyCode = intRetVal
24650   Exit Sub

ERRH:
24660   intRetVal = 0
24670   Select Case ERR.Number
        Case Else
24680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24690   End Select
24700   Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseMove"

24810   With Me
24820     If blnSelectAll_MouseDown = False Then
24830       Select Case blnSelectAll_Focus
            Case True
24840         .cmdSelectAll_raised_focus_dots_img.Visible = True
24850         .cmdSelectAll_raised_focus_img.Visible = False
24860       Case False
24870         .cmdSelectAll_raised_focus_img.Visible = True
24880         .cmdSelectAll_raised_focus_dots_img.Visible = False
24890       End Select
24900       .cmdSelectAll_raised_img.Visible = False
24910       .cmdSelectAll_raised_semifocus_dots_img.Visible = False
24920       .cmdSelectAll_sunken_focus_dots_img.Visible = False
24930       .cmdSelectAll_raised_img_dis.Visible = False
24940     End If
24950     If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
24960       Select Case blnSelectNone_Focus
            Case True
24970         .cmdSelectNone_raised_semifocus_dots_img.Visible = True
24980         .cmdSelectNone_raised_img.Visible = False
24990       Case False
25000         .cmdSelectNone_raised_img.Visible = True
25010         .cmdSelectNone_raised_semifocus_dots_img.Visible = False
25020       End Select
25030       .cmdSelectNone_raised_focus_img.Visible = False
25040       .cmdSelectNone_raised_focus_dots_img.Visible = False
25050       .cmdSelectNone_sunken_focus_dots_img.Visible = False
25060       .cmdSelectNone_raised_img_dis.Visible = False
25070     End If
25080   End With

EXITP:
25090   Exit Sub

ERRH:
25100   Select Case ERR.Number
        Case Else
25110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25120   End Select
25130   Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseUp"

25210   With Me
25220     .cmdSelectAll_raised_focus_dots_img.Visible = True
25230     .cmdSelectAll_raised_img.Visible = False
25240     .cmdSelectAll_raised_semifocus_dots_img.Visible = False
25250     .cmdSelectAll_raised_focus_img.Visible = False
25260     .cmdSelectAll_sunken_focus_dots_img.Visible = False
25270     .cmdSelectAll_raised_img_dis.Visible = False
25280     blnSelectAll_MouseDown = False
25290   End With

EXITP:
25300   Exit Sub

ERRH:
25310   Select Case ERR.Number
        Case Else
25320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25330   End Select
25340   Resume EXITP

End Sub

Private Sub cmdSelectAll_LostFocus()

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_LostFocus"

25410   With Me
25420     .cmdSelectAll_raised_img.Visible = True
25430     .cmdSelectAll_raised_semifocus_dots_img.Visible = False
25440     .cmdSelectAll_raised_focus_img.Visible = False
25450     .cmdSelectAll_raised_focus_dots_img.Visible = False
25460     .cmdSelectAll_sunken_focus_dots_img.Visible = False
25470     .cmdSelectAll_raised_img_dis.Visible = False
25480     blnSelectAll_Focus = False
25490   End With

EXITP:
25500   Exit Sub

ERRH:
25510   Select Case ERR.Number
        Case Else
25520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25530   End Select
25540   Resume EXITP

End Sub

Private Sub cmdSelectNone_GotFocus()

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_GotFocus"

25610   With Me
25620     blnSelectNone_Focus = True
25630     .cmdSelectNone_raised_semifocus_dots_img.Visible = True
25640     .cmdSelectNone_raised_img.Visible = False
25650     .cmdSelectNone_raised_focus_img.Visible = False
25660     .cmdSelectNone_raised_focus_dots_img.Visible = False
25670     .cmdSelectNone_sunken_focus_dots_img.Visible = False
25680     .cmdSelectNone_raised_img_dis.Visible = False
25690   End With

EXITP:
25700   Exit Sub

ERRH:
25710   Select Case ERR.Number
        Case Else
25720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25730   End Select
25740   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseDown"

25810   With Me
25820     blnSelectNone_MouseDown = True
25830     .cmdSelectNone_sunken_focus_dots_img.Visible = True
25840     .cmdSelectNone_raised_img.Visible = False
25850     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
25860     .cmdSelectNone_raised_focus_img.Visible = False
25870     .cmdSelectNone_raised_focus_dots_img.Visible = False
25880     .cmdSelectNone_raised_img_dis.Visible = False
25890   End With

EXITP:
25900   Exit Sub

ERRH:
25910   Select Case ERR.Number
        Case Else
25920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25930   End Select
25940   Resume EXITP

End Sub

Private Sub cmdSelectNone_Click()

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

26010   With Me

26020     DoCmd.Hourglass True
26030     DoEvents

26040     lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
26050     If lngRecsCur > 0& And .locasset_mark_sum > 0 Then
26060       Set dbs = CurrentDb
26070       With dbs
              ' ** Update tmpLocation_Asset, for locasset_mark = False.
26080         Set qdf = .QueryDefs("qryLocation_Asset_06b")
26090         qdf.Execute
26100         Set qdf = Nothing
26110         .Close
26120       End With
26130       Set dbs = Nothing
26140       .frmLocation_Asset_Sub.Form.Refresh
26150       DoEvents
26160       .frmLocation_Asset_Sub.Form.RecalcTots_Set True  ' ** Form Procedure: frmLocation_Asset_Sub.
26170       blnRecalcTots = True
26180       .TimerInterval = 1000&
26190     Else
26200       Beep
26210     End If

26220     lngRecsCur = .frmLocation_Asset_Sub.Form.RecCnt  ' ** Form Function: frmLocation_Asset_Sub.
26230     If lngRecsCur > 0& Then
26240       .frmLocation_Asset_Sub.SetFocus
26250     End If

26260     DoCmd.Hourglass False

26270   End With

EXITP:
26280   Set qdf = Nothing
26290   Set dbs = Nothing
26300   Exit Sub

ERRH:
26310   DoCmd.Hourglass False
26320   Select Case ERR.Number
        Case Else
26330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26340   End Select
26350   Resume EXITP

End Sub

Private Sub cmdSelectNone_KeyDown(KeyCode As Integer, Shift As Integer)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26420   intShiftDown = (Shift And acShiftMask) > 0
26430   intAltDown = (Shift And acAltMask) > 0
26440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26470       With Me
26480         intRetVal = 0
26490         If .cmdUpdate.Enabled = True Then
26500           .cmdUpdate.SetFocus
26510         Else
26520           .cmdClose.SetFocus
26530         End If
26540       End With
26550     End Select
26560   End If

        ' ** Shift keys.
26570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26580     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26590       With Me
26600         intRetVal = 0
26610         .cmdSelectAll.SetFocus
26620       End With
26630     End Select
26640   End If

EXITP:
26650   KeyCode = intRetVal
26660   Exit Sub

ERRH:
26670   intRetVal = 0
26680   Select Case ERR.Number
        Case Else
26690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26700   End Select
26710   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseMove"

26810   With Me
26820     If blnSelectNone_MouseDown = False Then
26830       Select Case blnSelectNone_Focus
            Case True
26840         .cmdSelectNone_raised_focus_dots_img.Visible = True
26850         .cmdSelectNone_raised_focus_img.Visible = False
26860       Case False
26870         .cmdSelectNone_raised_focus_img.Visible = True
26880         .cmdSelectNone_raised_focus_dots_img.Visible = False
26890       End Select
26900       .cmdSelectNone_raised_img.Visible = False
26910       .cmdSelectNone_raised_semifocus_dots_img.Visible = False
26920       .cmdSelectNone_sunken_focus_dots_img.Visible = False
26930       .cmdSelectNone_raised_img_dis.Visible = False
26940     End If
26950     If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
26960       Select Case blnSelectAll_Focus
            Case True
26970         .cmdSelectAll_raised_semifocus_dots_img.Visible = True
26980         .cmdSelectAll_raised_img.Visible = False
26990       Case False
27000         .cmdSelectAll_raised_img.Visible = True
27010         .cmdSelectAll_raised_semifocus_dots_img.Visible = False
27020       End Select
27030       .cmdSelectAll_raised_focus_img.Visible = False
27040       .cmdSelectAll_raised_focus_dots_img.Visible = False
27050       .cmdSelectAll_sunken_focus_dots_img.Visible = False
27060       .cmdSelectAll_raised_img_dis.Visible = False
27070     End If
27080   End With

EXITP:
27090   Exit Sub

ERRH:
27100   Select Case ERR.Number
        Case Else
27110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27120   End Select
27130   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseUp"

27210   With Me
27220     .cmdSelectNone_raised_focus_dots_img.Visible = True
27230     .cmdSelectNone_raised_img.Visible = False
27240     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
27250     .cmdSelectNone_raised_focus_img.Visible = False
27260     .cmdSelectNone_sunken_focus_dots_img.Visible = False
27270     .cmdSelectNone_raised_img_dis.Visible = False
27280     blnSelectNone_MouseDown = False
27290   End With

EXITP:
27300   Exit Sub

ERRH:
27310   Select Case ERR.Number
        Case Else
27320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27330   End Select
27340   Resume EXITP

End Sub

Private Sub cmdSelectNone_LostFocus()

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_LostFocus"

27410   With Me
27420     .cmdSelectNone_raised_img.Visible = True
27430     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
27440     .cmdSelectNone_raised_focus_img.Visible = False
27450     .cmdSelectNone_raised_focus_dots_img.Visible = False
27460     .cmdSelectNone_sunken_focus_dots_img.Visible = False
27470     .cmdSelectNone_raised_img_dis.Visible = False
27480     blnSelectNone_Focus = False
27490   End With

EXITP:
27500   Exit Sub

ERRH:
27510   Select Case ERR.Number
        Case Else
27520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27530   End Select
27540   Resume EXITP

End Sub

Private Function Loc_Counts() As Boolean

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Counts"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngLocID As Long, lngAssetNo As Long
        Dim strAccountNo As String
        Dim lngAA_Cnt As Long, lngL_Cnt As Long, lngLA_Cnt As Long, lngJ_Cnt As Long, lngJC_Cnt As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

27610   blnRetVal = True

27620   lngLocID = 0&
27630   strAccountNo = "-1"
27640   lngAssetNo = -1&
27650   lngAA_Cnt = 0&: lngL_Cnt = 0&: lngLA_Cnt = 0&: lngJ_Cnt = 0&: lngJC_Cnt = 0&

27660   With Me
27670     Set dbs = CurrentDb
27680     Select Case IsNull(.cmbLocations_Current)
          Case True
            ' ** Let them stand.
27690     Case False
27700       lngLocID = .cmbLocations_Current
27710       Select Case .opgGroup
            Case .opgGroup_optAll.OptionValue
              ' ** No further criteria.
27720         With dbs
27730           For lngX = 1& To 5&
27740             Select Case lngX
                  Case 1&
                    ' ** qryLocation_Asset_10a (ActiveAssets, linked to Location,
                    ' ** MasterAsset, with uniqueid), grouped by Location_ID, with cnt_aa.
27750               Set qdf = .QueryDefs("qryLocation_Asset_10b")
27760             Case 2&
                    ' ** qryLocation_Asset_11a (Ledger, linked to Location),
                    ' ** grouped by Location_ID, with cnt_l.
27770               Set qdf = .QueryDefs("qryLocation_Asset_11b")
27780             Case 3&
                    ' ** qryLocation_Asset_12a (LedgerArchive, linked to Location),
                    ' ** grouped by Location_ID, with cnt_la.
27790               Set qdf = .QueryDefs("qryLocation_Asset_12b")
27800             Case 4&
                    ' ** qryLocation_Asset_13a (Journal, linked to Location),
                    ' ** grouped by Location_ID, with cnt_j.
27810               Set qdf = .QueryDefs("qryLocation_Asset_13b")
27820             Case 5&
                    ' ** qryLocation_Asset_14a (tblJournal_Column, linked to Location),
                    ' ** grouped by Location_ID, with cnt_jc.
27830               Set qdf = .QueryDefs("qryLocation_Asset_14b")
27840             End Select
27850             Set rst = qdf.OpenRecordset
27860             With rst
27870               If .BOF = True And .EOF = True Then
                      ' ** Let them stand.
27880               Else
27890                 .FindFirst "[Location_ID] = " & CStr(lngLocID)
27900                 Select Case .NoMatch
                      Case True
                        ' ** Let them stand.
27910                 Case False
27920                   Select Case lngX
                        Case 1&
27930                     lngAA_Cnt = ![Cnt_AA]
27940                   Case 2&
27950                     lngL_Cnt = ![Cnt_L]
27960                   Case 3&
27970                     lngLA_Cnt = ![Cnt_LA]
27980                   Case 4&
27990                     lngJ_Cnt = ![Cnt_J]
28000                   Case 5&
28010                     lngJC_Cnt = ![Cnt_JC]
28020                   End Select
28030                 End Select
28040               End If
28050               .Close
28060             End With  ' ** rst.
28070             Set rst = Nothing
28080             Set qdf = Nothing
28090           Next  ' ** lngX.
28100         End With  ' ** dbs.
28110       Case .opgGroup_optAccount.OptionValue
28120         Select Case IsNull(.cmbAccounts)
              Case True
                ' ** Let them stand.
28130           strAccountNo = "0"
28140         Case False
28150           strAccountNo = .cmbAccounts
28160           With dbs
28170             For lngX = 1& To 5&
28180               Select Case lngX
                    Case 1&
                      ' ** qryLocation_Asset_10a (xx), grouped by Location_ID, accountno, with cnt_aa_actno.
28190                 Set qdf = .QueryDefs("qryLocation_Asset_10c")
28200               Case 2&
                      ' ** qryLocation_Asset_11a (xx), grouped by Location_ID, accountno, with cnt_l_actno.
28210                 Set qdf = .QueryDefs("qryLocation_Asset_11c")
28220               Case 3&
                      ' ** qryLocation_Asset_12a (xx), grouped by Location_ID, accountno, with cnt_la_actno.
28230                 Set qdf = .QueryDefs("qryLocation_Asset_12c")
28240               Case 4&
                      ' ** qryLocation_Asset_13a (xx), grouped by Location_ID, accountno, with cnt_j_actno.
28250                 Set qdf = .QueryDefs("qryLocation_Asset_13c")
28260               Case 5&
                      ' ** qryLocation_Asset_14a (xx), grouped by Location_ID, accountno, with cnt_jc_actno.
28270                 Set qdf = .QueryDefs("qryLocation_Asset_14c")
28280               End Select
28290               Set rst = qdf.OpenRecordset
28300               With rst
28310                 If .BOF = True And .EOF = True Then
                        ' ** Let them stand.
28320                 Else
28330                   .FindFirst "[Location_ID] = " & CStr(lngLocID) & " And [accountno] = '" & strAccountNo & "'"
28340                   Select Case .NoMatch
                        Case True
                          ' ** Let them stand.
28350                   Case False
28360                     Select Case lngX
                          Case 1&
28370                       lngAA_Cnt = ![cnt_aa_actno]
28380                     Case 2&
28390                       lngL_Cnt = ![cnt_l_actno]
28400                     Case 3&
28410                       lngLA_Cnt = ![cnt_la_actno]
28420                     Case 4&
28430                       lngJ_Cnt = ![cnt_j_actno]
28440                     Case 5&
28450                       lngJC_Cnt = ![cnt_jc_actno]
28460                     End Select
28470                   End Select
28480                 End If
28490                 .Close
28500               End With  ' ** rst.
28510               Set rst = Nothing
28520               Set qdf = Nothing
28530             Next  ' ** lngX.
28540           End With  ' ** dbs.
28550         End Select
28560       Case .opgGroup_optAsset.OptionValue
28570         Select Case IsNull(.cmbAssets)
              Case True
28580           lngAssetNo = 0&
28590         Case False
28600           lngAssetNo = .cmbAssets
28610           With dbs
28620             For lngX = 1& To 5&
28630               Select Case lngX
                    Case 1&
                      ' ** qryLocation_Asset_10a (xx), grouped by Location_ID, assetno, with cnt_aa_astno.
28640                 Set qdf = .QueryDefs("qryLocation_Asset_10d")
28650               Case 2&
                      ' ** qryLocation_Asset_11a (xx), grouped by Location_ID, assetno, with cnt_l_astno.
28660                 Set qdf = .QueryDefs("qryLocation_Asset_11d")
28670               Case 3&
                      ' ** qryLocation_Asset_12a (xx), grouped by Location_ID, assetno, with cnt_la_astno.
28680                 Set qdf = .QueryDefs("qryLocation_Asset_12d")
28690               Case 4&
                      ' ** qryLocation_Asset_13a (xx), grouped by Location_ID, assetno, with cnt_j_astno.
28700                 Set qdf = .QueryDefs("qryLocation_Asset_13d")
28710               Case 5&
                      ' ** qryLocation_Asset_14a (xx), grouped by Location_ID, assetno, with cnt_jc_astno.
28720                 Set qdf = .QueryDefs("qryLocation_Asset_14d")
28730               End Select
28740               Set rst = qdf.OpenRecordset
28750               With rst
28760                 If .BOF = True And .EOF = True Then
                        ' ** Let them stand.
28770                 Else
28780                   .FindFirst "[Location_ID] = " & CStr(lngLocID) & " And [assetno] = " & CStr(lngAssetNo)
28790                   Select Case .NoMatch
                        Case True
                          ' ** Let them stand.
28800                   Case False
28810                     Select Case lngX
                          Case 1&
28820                       lngAA_Cnt = ![cnt_aa_astno]
28830                     Case 2&
28840                       lngL_Cnt = ![cnt_l_astno]
28850                     Case 3&
28860                       lngLA_Cnt = ![cnt_la_astno]
28870                     Case 4&
28880                       lngJ_Cnt = ![cnt_j_astno]
28890                     Case 5&
28900                       lngJC_Cnt = ![cnt_jc_astno]
28910                     End Select
28920                   End Select
28930                 End If
28940                 .Close
28950               End With  ' ** rst.
28960               Set rst = Nothing
28970               Set qdf = Nothing
28980             Next  ' ** lngX.
28990           End With  ' ** dbs.
29000         End Select
29010       Case .opgGroup_optAccountAsset.OptionValue
29020         Select Case IsNull(.cmbAccounts)
              Case True
                ' ** Let them stand.
29030           strAccountNo = "0"
29040         Case False
29050           strAccountNo = .cmbAccounts
29060         End Select
29070         Select Case IsNull(.cmbAssets)
              Case True
                ' ** Let them stand.
29080           lngAssetNo = 0&
29090         Case False
29100           lngAssetNo = .cmbAssets
29110         End Select
29120         If strAccountNo <> "0" And lngAssetNo > 0& Then
29130           With dbs
29140             For lngX = 1& To 5&
29150               Select Case lngX
                    Case 1&
                      ' ** qryLocation_Asset_10a (xx), grouped by Location_ID, accountno, assetno, with cnt_aa_actno_astno.
29160                 Set qdf = .QueryDefs("qryLocation_Asset_10e")
29170               Case 2&
                      ' ** qryLocation_Asset_11a (xx), grouped by Location_ID, accountno, assetno, with cnt_l_actno_astno.
29180                 Set qdf = .QueryDefs("qryLocation_Asset_11e")
29190               Case 3&
                      ' ** qryLocation_Asset_12a (xx), grouped by Location_ID, accountno, assetno, with cnt_la_actno_astno.
29200                 Set qdf = .QueryDefs("qryLocation_Asset_12e")
29210               Case 4&
                      ' ** qryLocation_Asset_13a (xx), grouped by Location_ID, accountno, assetno, with cnt_j_actno_astno.
29220                 Set qdf = .QueryDefs("qryLocation_Asset_13e")
29230               Case 5&
                      ' ** qryLocation_Asset_14a (xx), grouped by Location_ID, accountno, assetno, with cnt_jc_actno_astno.
29240                 Set qdf = .QueryDefs("qryLocation_Asset_14e")
29250               End Select
29260               Set rst = qdf.OpenRecordset
29270               With rst
29280                 If .BOF = True And .EOF = True Then
                        ' ** Let them stand.
29290                 Else
29300                   .FindFirst "[Location_ID] = " & CStr(lngLocID) & " And [accountno] = '" & strAccountNo & "' And " & _
                          "[assetno] = " & CStr(lngAssetNo)
29310                   Select Case .NoMatch
                        Case True
                          ' ** Let them stand.
29320                   Case False
29330                     Select Case lngX
                          Case 1&
29340                       lngAA_Cnt = ![cnt_aa_actno_astno]
29350                     Case 2&
29360                       lngL_Cnt = ![cnt_l_actno_astno]
29370                     Case 3&
29380                       lngLA_Cnt = ![cnt_la_actno_astno]
29390                     Case 4&
29400                       lngJ_Cnt = ![cnt_j_actno_astno]
29410                     Case 5&
29420                       lngJC_Cnt = ![cnt_jc_actno_astno]
29430                     End Select
29440                   End Select
29450                 End If
29460                 .Close
29470               End With  ' ** rst.
29480               Set rst = Nothing
29490               Set qdf = Nothing
29500             Next  ' ** lngX.
29510           End With  ' ** dbs.
29520         End If
29530       End Select
29540     End Select
          ' ** COMBOS HERE!
29550     dbs.Close
29560     Set dbs = Nothing

29570     .Cnt_AA = lngAA_Cnt
29580     If lngAA_Cnt > 0& Then
29590       .Cnt_AA.Locked = True
29600       .Cnt_AA.BackStyle = acBackStyleNormal
29610       .Cnt_AA.BorderColor = CLR_LTBLU2
29620     Else
29630       .Cnt_AA.Locked = False
29640       .Cnt_AA.BackStyle = acBackStyleTransparent
29650       .Cnt_AA.BorderColor = WIN_CLR_DISR
29660     End If
29670     .Cnt_L = lngL_Cnt
29680     If lngL_Cnt > 0& Then
29690       .Cnt_L.Visible = True
29700     Else
29710       .Cnt_L.Visible = False
29720     End If
29730     .Cnt_LA = lngLA_Cnt
29740     If lngLA_Cnt > 0& Then
29750       .Cnt_LA.Visible = True
29760     Else
29770       .Cnt_LA.Visible = False
29780     End If
29790     .Cnt_J = lngJ_Cnt
29800     If lngJ_Cnt > 0& Then
29810       .Cnt_J.Visible = True
29820     Else
29830       .Cnt_J.Visible = False
29840     End If
29850     .Cnt_JC = lngJC_Cnt
29860     If lngJC_Cnt > 0& Then
29870       .Cnt_JC.Visible = True
29880     Else
29890       .Cnt_JC.Visible = False
29900     End If

29910   End With

EXITP:
29920   Set rst = Nothing
29930   Set qdf = Nothing
29940   Set dbs = Nothing
29950   Loc_Counts = blnRetVal
29960   Exit Function

ERRH:
29970   blnRetVal = False
29980   Select Case ERR.Number
        Case Else
29990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30000   End Select
30010   Resume EXITP

End Function

Private Function Validate() As Boolean

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "Validate"

        Dim blnRetVal As Boolean

30110   blnRetVal = False

30120   With Me
30130     If .Cnt_AA > 0 Or .Cnt_L > 0 Or .Cnt_LA > 0 Or .Cnt_J > 0 Or .Cnt_JC > 0 Then
30140       Select Case IsNull(.cmbLocations_Current)
            Case True
30150         .cmdUpdate.Enabled = False
30160       Case False
30170         Select Case IsNull(.cmbLocations_New)
              Case True
30180           .cmdUpdate.Enabled = False
30190         Case False
30200           If .cmbLocations_New <> .cmbLocations_Current Then
30210             blnRetVal = True
30220             .cmdUpdate.Enabled = True
30230           Else
30240             .cmdUpdate.Enabled = False
30250           End If
30260         End Select
30270       End Select
30280     Else
30290       .cmdUpdate.Enabled = False
30300     End If
30310   End With

EXITP:
30320   Validate = blnRetVal
30330   Exit Function

ERRH:
30340   blnRetVal = False
30350   Select Case ERR.Number
        Case Else
30360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30370   End Select
30380   Resume EXITP

End Function

Public Function ResetValidate() As Boolean
' ** True: Enabled, False: Disabled.

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "ResetValidate"

        Dim strControlName As String
        Dim blnRetVal As Boolean

30410   blnRetVal = False

30420 On Error Resume Next
30430   strControlName = Screen.ActiveControl.Name
30440 On Error GoTo ERRH

30450   With Me
30460     Select Case IsNull(.cmbLocations_Current)
          Case True
30470       Select Case IsNull(.cmbLocations_New)
            Case True
30480         Select Case .opgGroup
              Case .opgGroup_optAll.OptionValue
30490           Select Case IsNull(.cmbAccounts)
                Case True
30500             Select Case IsNull(.cmbAssets)
                  Case True
                    ' ** Nothing out of place.
30510             Case False
30520               blnRetVal = True
30530             End Select
30540           Case False
30550             blnRetVal = True
30560           End Select
30570         Case Else
30580           blnRetVal = True
30590         End Select
30600       Case False
30610         blnRetVal = True
30620       End Select
30630     Case False
30640       blnRetVal = True
30650     End Select
30660     Select Case blnRetVal
          Case True
            ' ** Not default, allow reset to default.
30670       If .cmdResetFilter.Enabled = False Then
30680         .cmdResetFilter.Enabled = True
30690         .cmdResetFilter_raised_img.Visible = True
30700         .cmdResetFilter_raised_semifocus_dots_img.Visible = False
30710         .cmdResetFilter_raised_focus_img.Visible = False
30720         .cmdResetFilter_raised_focus_dots_img.Visible = False
30730         .cmdResetFilter_sunken_focus_dots_img.Visible = False
30740         .cmdResetFilter_raised_img_dis.Visible = False
30750       End If
30760     Case False
            ' ** Default, don't allow reset since it's already there.
30770       If .cmdResetFilter.Enabled = True Then
30780         If strControlName = "cmdResetFilter" Then
30790           .cmbLocations_Current.SetFocus
30800           DoEvents
30810         End If
30820         .cmdResetFilter.Enabled = False
30830         .cmdResetFilter_raised_img_dis.Visible = True
30840         .cmdResetFilter_raised_img.Visible = False
30850         .cmdResetFilter_raised_semifocus_dots_img.Visible = False
30860         .cmdResetFilter_raised_focus_img.Visible = False
30870         .cmdResetFilter_raised_focus_dots_img.Visible = False
30880         .cmdResetFilter_sunken_focus_dots_img.Visible = False
30890       End If
30900     End Select
30910   End With

EXITP:
30920   ResetValidate = blnRetVal
30930   Exit Function

ERRH:
30940   Select Case ERR.Number
        Case Else
30950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30960   End Select
30970   Resume EXITP

End Function
