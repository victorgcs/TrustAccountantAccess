VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSweeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmSweeper"

'VGC 04/15/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Posting Date:     D {SweepDate}
' **   Cover PC ODs:     E {opgType_optCover}
' **   Income Cash:      H {opgCash_optIncome}
' **   By Cusip:         I {opgAssetSource_optCusip}
' **   Create Jrnl Ent:  J {cmdCreateEntries}
' **   By Name:          N {opgAssetSource_optName}
' **   Inc/Exp Codes:    O {cmbRevenueCodes_Purchase}
' **   Principal Cash:   P {opgCash_optPrincipal}
' **   Asset:            S {cmbAssets}
' **   Tax Codes:        T {cmbTaxCodes_Purchase}
' **   Invest:           V {opgType_optInvest}
' **   Exit:             X {cmdClose}
' **   By Type:          Y {opgAssetSource_optType}

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar}

' ** CLOSES WHEN FINISHED!

' *****************************************************
' ** NOTE: Only AssetTypes 60 and 80 can be swept to!
' *****************************************************

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long

Private lngCntSweeps As Long
Private blnNoSweeps As Boolean, blnFromCalendar As Boolean, blnKeyDown_Return As Boolean, blnHideIncExp As Boolean
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long ', lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           blnFromCalendar = False: blnKeyDown_Return = False: blnHideIncExp = False
170           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

              ' ** Variables are fed empty, then populated ByRef.
190           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

200           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

210           HideIncExp  ' ** Procedure: Below.
220           DoEvents

230           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
240           If lngMonitorCnt > 1& Then
250             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
260             lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
270             If lngMonitorNum > 1& Then
280               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
290             End If
300           End If

310           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
320           DoEvents

              ' ** EVENT CHECK: chkRememberAsset!
330           cmbAssets_AfterUpdate  ' ** Procedure: Below.

340           opgType_AfterUpdate  ' ** Procedure: Below.
350           opgCash_AfterUpdate  ' ** Procedure: Below.
360           opgAssetSource_AfterUpdate  ' ** Procedure: Below.
370           chkRememberAsset_AfterUpdate  ' ** Procedure: Below.

380           blnIsOpen = False

390         Else
400           Cancel = -1
410         End If
420       Else
430         Cancel = -1
440       End If
450     End With

460     If Cancel = -1 Then
470       Beep
480       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
490     End If

EXITP:
500     Exit Sub

ERRH:
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim blnNoSweepJrnlNo As Boolean, blnNoSweepActNo As Boolean, blnNoSweepAstNo As Boolean
        Dim varTmp00 As Variant

610     blnNoSweeps = False
620     blnNoSweepJrnlNo = False: blnNoSweepActNo = False: blnNoSweepAstNo = False

        ' ** Create an instance of the Calendar class.
630     Set clsMonthClass = New clsMonthCal
        ' ** You MUST SET the class hWndForm prop!!!
640     clsMonthClass.hWndForm = Me.hwnd
        ' ** Let's default to PositionAtCursor.
650     clsMonthClass.PositionAtCursor = True

660     With Me

670       Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
680         .cmbRevenueCodes_Purchase = REVID_INC
690         .cmbRevenueCodes_Purchase_display = Null
700       Case False
710         If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
720           .cmbRevenueCodes_Purchase_display = Null
730         Else
740           .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
750         End If
760       End Select

770       Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
780         .cmbTaxCodes_Purchase = TAXID_INC
790         .cmbTaxCodes_Purchase_display = Null
800       Case False
810         If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
820           .cmbTaxCodes_Purchase_display = Null
830         Else
840           .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
850         End If
860       End Select

870       .NoSweep_Account_lbl.Visible = False
880       .NoSweep_Asset_lbl.Visible = False
890       .NoSweep_Journal_lbl.Visible = False

          ' ** Account, grouped, with cnt_sweeps.
900       varTmp00 = DLookup("[cnt_sweeps]", "qrySweeper_06")
910       If IsNull(varTmp00) = True Then
920         blnNoSweeps = True: blnNoSweepActNo = True
930       Else
940         If varTmp00 = 0 Then
950           blnNoSweeps = True: blnNoSweepActNo = True
960         End If
970       End If

          ' ** RowSource, combo box; MasterAsset, just sweeps.
980       lngCntSweeps = 0&
990       varTmp00 = DCount("*", "qrySweeper_05")
1000      If IsNull(varTmp00) = True Then
1010        blnNoSweeps = True: blnNoSweepAstNo = True
1020      Else
1030        If varTmp00 = 0 Then
1040          blnNoSweeps = True: blnNoSweepAstNo = True
1050        Else
1060          varTmp00 = DLookup("[cnt_sweeps]", "qrySweeper_08")
1070          If IsNull(varTmp00) = False Then
1080            lngCntSweeps = varTmp00
1090          End If
1100        End If
1110      End If

1120      varTmp00 = DCount("*", "journal")
1130      If IsNull(varTmp00) = False Then
1140        If varTmp00 > 0 Then
1150          blnNoSweeps = True: blnNoSweepJrnlNo = True
1160        End If
1170      End If

1180      .cmdCreateEntries.Enabled = False  ' ** Until they put in a date.

1190      Select Case blnNoSweeps
          Case True
1200        .cmdClose.SetFocus
            '.SweepDate_lbl.BackStyle = acBackStyleTransparent
1210        .SweepDate.Enabled = False
1220        .SweepDate = Null
1230        .SweepDate.BorderColor = WIN_CLR_DISR
1240        .SweepDate.BackStyle = acBackStyleTransparent
1250        .cmdCalendar.Enabled = False
1260        .cmdCalendar_raised_img_dis.Visible = True
1270        .cmdCalendar_raised_img.Visible = False
1280        .cmdCalendar_raised_semifocus_dots_img.Visible = False
1290        .cmdCalendar_raised_focus_img.Visible = False
1300        .cmdCalendar_raised_focus_dots_img.Visible = False
1310        .cmdCalendar_sunken_focus_dots_img.Visible = False
1320        .opgCash.Enabled = False
1330        .opgType.Enabled = False
            '.cmbAssets_lbl.BackStyle = acBackStyleTransparent
1340        .cmbAssets.Enabled = False
1350        .cmbAssets.BorderColor = WIN_CLR_DISR
1360        .cmbAssets.BackStyle = acBackStyleTransparent
1370        .opgAssetSource.Enabled = False
1380        .opgAssetSource_optType_lbl2.ForeColor = WIN_CLR_DISF
1390        .opgAssetSource_optType_lbl2_dim_hi.Visible = True
1400        .opgAssetSource_optName_lbl2.ForeColor = WIN_CLR_DISF
1410        .opgAssetSource_optName_lbl2_dim_hi.Visible = True
1420        .opgAssetSource_optCusip_lbl2.ForeColor = WIN_CLR_DISF
1430        .opgAssetSource_optCusip_lbl2_dim_hi.Visible = True
1440        .chkRememberAsset.Enabled = False
1450        .chkRememberAsset_lbl.Visible = False
1460        .chkRememberAsset_lbl2_dim.Visible = True
1470        .chkRememberAsset_lbl2_dim_hi.Visible = True
1480        If blnHideIncExp = False Then
1490          .cmbRevenueCodes_Purchase_display = Null
1500          .cmbRevenueCodes_Purchase_display.Enabled = False
1510          .cmbRevenueCodes_Purchase_display.Locked = False
1520          .cmbRevenueCodes_Purchase = Null
1530          .cmbRevenueCodes_Purchase.Enabled = False
1540          .cmbRevenueCodes_Purchase.BorderColor = WIN_CLR_DISR
              '.cmbRevenueCodes_Purchase_lbl.BackStyle = acBackStyleTransparent
1550          .cmbTaxCodes_Purchase_display = Null
1560          .cmbTaxCodes_Purchase_display.Enabled = False
1570          .cmbTaxCodes_Purchase_display.Locked = False
1580          .cmbTaxCodes_Purchase = Null
1590          .cmbTaxCodes_Purchase.Enabled = False
1600          .cmbTaxCodes_Purchase.BorderColor = WIN_CLR_DISR
              '.cmbTaxCodes_Purchase_lbl.BackStyle = acBackStyleTransparent
1610        End If  ' ** blnHideIncExp.
1620        If blnNoSweepActNo = True Then .NoSweep_Account_lbl.Visible = True
1630        If blnNoSweepAstNo = True Then .NoSweep_Asset_lbl.Visible = True
1640        If blnNoSweepJrnlNo = True Then .NoSweep_Journal_lbl.Visible = True
1650      Case False
1660        .NoSweep_Account_lbl.Visible = False
1670        .NoSweep_Asset_lbl.Visible = False
1680        .NoSweep_Journal_lbl.Visible = False
            '.SweepDate_lbl.BackStyle = acBackStyleNormal
1690        .SweepDate.Enabled = True
1700        .SweepDate.BorderColor = CLR_LTBLU2
1710        .SweepDate.BackStyle = acBackStyleNormal
1720        .cmdCalendar.Enabled = True
1730        .cmdCalendar_raised_img.Visible = True
1740        .cmdCalendar_raised_semifocus_dots_img.Visible = False
1750        .cmdCalendar_raised_focus_img.Visible = False
1760        .cmdCalendar_raised_focus_dots_img.Visible = False
1770        .cmdCalendar_sunken_focus_dots_img.Visible = False
1780        .cmdCalendar_raised_img_dis.Visible = False
1790        .opgType.Enabled = True
1800        If .opgType = .opgType_optCover.OptionValue Then
1810          .opgCash.Enabled = False
1820        Else
1830          .opgCash.Enabled = True
1840        End If
            '.cmbAssets_lbl.BackStyle = acBackStyleNormal
1850        .cmbAssets.Enabled = True
1860        .cmbAssets.BorderColor = CLR_LTBLU2
1870        .cmbAssets.BackStyle = acBackStyleNormal
1880        .opgAssetSource.Enabled = True
1890        .opgAssetSource_optType_lbl2.ForeColor = CLR_VDKGRY
1900        .opgAssetSource_optType_lbl2_dim_hi.Visible = False
1910        .opgAssetSource_optName_lbl2.ForeColor = CLR_VDKGRY
1920        .opgAssetSource_optName_lbl2_dim_hi.Visible = False
1930        .opgAssetSource_optCusip_lbl2.ForeColor = CLR_VDKGRY
1940        .opgAssetSource_optCusip_lbl2_dim_hi.Visible = False
1950        .chkRememberAsset.Enabled = True
1960        .chkRememberAsset_lbl.Visible = True
1970        .chkRememberAsset_lbl2_dim.Visible = False
1980        .chkRememberAsset_lbl2_dim_hi.Visible = False
1990      End Select

2000      If blnHideIncExp = False Then

2010        Select Case gblnRevenueExpenseTracking
            Case True
2020          .cmbRevenueCodes_Purchase.Visible = True
2030          .cmbRevenueCodes_Purchase_lbl2.Visible = False
2040        Case False
2050          .cmbRevenueCodes_Purchase.Visible = False
2060          .cmbRevenueCodes_Purchase_display.Visible = False
2070          .cmbRevenueCodes_Purchase_lbl2.Visible = True
2080        End Select

2090        Select Case gblnIncomeTaxCoding
            Case True
2100          .cmbTaxCodes_Purchase.Visible = True
2110          .cmbTaxCodes_Purchase_lbl2.Visible = False
2120        Case False
2130          .cmbTaxCodes_Purchase.Visible = False
2140          .cmbTaxCodes_Purchase_display.Visible = False
2150          .cmbTaxCodes_Purchase_lbl2.Visible = True
2160        End Select

2170        If gblnRevenueExpenseTracking = False And gblnIncomeTaxCoding = False Then
2180          .cmbRevenueCodes_Purchase_box_lbl.ForeColor = WIN_CLR_DISF
2190          .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Visible = True
2200        End If

2210      End If  ' ** blnHideIncExp.

2220      If lngCntSweeps > 0& Then
2230        .cmbAssets = .cmbAssets.Column(0, 0)
2240      End If

2250    End With

EXITP:
2260    Exit Sub

ERRH:
2270    Select Case ERR.Number
        Case Else
2280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2290    End Select
2300    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2420    intShiftDown = (Shift And acShiftMask) > 0
2430    intAltDown = (Shift And acAltMask) > 0
2440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   By Cusip:         I {opgAssetSource_optCusip}
        ' **   By Name:          N {opgAssetSource_optName}
        ' **   Exit:             X {cmdClose}
        ' **   By Type:          Y {opgAssetSource_optType}

        ' ** Shortcut F keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar}

        ' ** Plain keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2460      Select Case intRetVal
          Case vbKeyEscape
2470        intRetVal = 0
2480        cmdClose_Click  ' ** Procedure: Below.
2490      Case vbKeyF7
2500        With Me
2510          intRetVal = 0
2520          If .cmdCalendar.Enabled = True Then
2530            .cmdCalendar.SetFocus
2540            cmdCalendar_Click  ' ** Procedure: Below.
2550          End If
2560        End With
2570      End Select
2580    End If

        ' ** Alt keys.
2590    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2600      Select Case intRetVal
          Case vbKeyI
2610        With Me
2620          intRetVal = 0
2630          If .opgAssetSource.Enabled = True Then
2640            .opgAssetSource.SetFocus
2650            .opgAssetSource = .opgAssetSource_optCusip.OptionValue
2660            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
2670          Else
2680            Beep
2690          End If
2700        End With
2710      Case vbKeyN
2720        With Me
2730          intRetVal = 0
2740          If .opgAssetSource.Enabled = True Then
2750            .opgAssetSource.SetFocus
2760            .opgAssetSource = .opgAssetSource_optName.OptionValue
2770            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
2780          Else
2790            Beep
2800          End If
2810        End With
2820      Case vbKeyX
2830        intRetVal = 0
2840        cmdClose_Click  ' ** Procedure: Below.
2850      Case vbKeyY
2860        With Me
2870          intRetVal = 0
2880          If .opgAssetSource.Enabled = True Then
2890            .opgAssetSource.SetFocus
2900            .opgAssetSource = .opgAssetSource_optType.OptionValue
2910            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
2920          Else
2930            Beep
2940          End If
2950        End With
2960      End Select
2970    End If

EXITP:
2980    KeyCode = intRetVal
2990    Exit Sub

ERRH:
3000    intRetVal = 0
3010    Select Case ERR.Number
        Case Else
3020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3030    End Select
3040    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3110    Select Case DataErr
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
3120    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3130      If Right(Me.ActiveControl.Name, 4) = "Date" Then  ' ** SweepDate.
3140        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3150        Response = acDataErrContinue
3160      End If
3170    Case Else
3180      zErrorHandler THIS_NAME, "Undefined Form Error", DataErr  ' ** Module Function: modErrorHandler.
3190    End Select

EXITP:
3200    Exit Sub

ERRH:
3210    Select Case ERR.Number
        Case Else
3220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3230    End Select
3240    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3310    With Me
3320      If .chkRememberAsset = False Then
3330        .cmbAssets = Null  ' ** So it won't be there when Pref_Save() runs.
3340      End If
3350    End With

3360    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3370    If Not clsMonthClass Is Nothing Then
3380      If clsMonthClass.IsCalendar Then
3390        Cancel = -1
3400      Else
3410        Set clsMonthClass = Nothing
3420      End If
3430    End If

3440    gstrFormQuerySpec = vbNullString
3450    gblnSetFocus = True
3460    gstrReturningForm = THIS_NAME
3470    Select Case strCallingForm
        Case "frmMenu_Post"
3480      If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
3490        DoCmd.SelectObject acForm, strCallingForm, False
3500        Forms(strCallingForm).UpdateTransCnt  ' ** Form Function: frmMenu_Post.
3510        Forms(strCallingForm).TimerInterval = 250&
3520      Else
3530        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
3540      End If
3550    End Select

EXITP:
3560    Exit Sub

ERRH:
3570    Select Case ERR.Number
        Case Else
3580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3590    End Select
3600    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3710    With Me
3720      If .cmdCalendar_raised_focus_dots_img.Visible = True Or .cmdCalendar_raised_focus_img.Visible = True Then
3730        Select Case blnCalendar1_Focus
            Case True
3740          .cmdCalendar_raised_semifocus_dots_img.Visible = True
3750          .cmdCalendar_raised_img.Visible = False
3760        Case False
3770          .cmdCalendar_raised_img.Visible = True
3780          .cmdCalendar_raised_semifocus_dots_img.Visible = False
3790        End Select
3800        .cmdCalendar_raised_focus_dots_img.Visible = False
3810        .cmdCalendar_raised_focus_img.Visible = False
3820        .cmdCalendar_sunken_focus_dots_img.Visible = False
3830        .cmdCalendar_raised_img_dis.Visible = False
3840      End If
3850    End With

EXITP:
3860    Exit Sub

ERRH:
3870    Select Case ERR.Number
        Case Else
3880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3890    End Select
3900    Resume EXITP

End Sub

Private Sub cmdClose_Click()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

4010    DoCmd.Close acForm, THIS_NAME

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4120    intShiftDown = (Shift And acShiftMask) > 0
4130    intAltDown = (Shift And acAltMask) > 0
4140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4160      Select Case intRetVal
          Case vbKeyTab
4170        With Me
4180          intRetVal = 0
4190          If .SweepDate.Enabled = True Then
4200            .SweepDate.SetFocus
4210          Else
                ' ** Nothing else.
4220          End If
4230        End With
4240      End Select
4250    End If

        ' ** Shift keys.
4260    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4270      Select Case intRetVal
          Case vbKeyTab
4280        With Me
4290          intRetVal = 0
4300          If .cmdCreateEntries.Enabled = True Then
4310            .cmdCreateEntries.SetFocus
4320          ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
4330            .cmbTaxCodes_Purchase.SetFocus
4340          ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
4350            .cmbRevenueCodes_Purchase.SetFocus
4360          ElseIf .cmbAssets.Enabled = True Then
4370            .cmbAssets.SetFocus
4380          Else
                ' ** Nothing else.
4390          End If
4400        End With
4410      End Select
4420    End If

EXITP:
4430    KeyCode = intRetVal
4440    Exit Sub

ERRH:
4450    intRetVal = 0
4460    Select Case ERR.Number
        Case Else
4470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4480    End Select
4490    Resume EXITP

End Sub

Private Sub SweepDate_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "SweepDate_box_MouseMove"

4510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub SweepDate_Enter()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "SweepDate_Enter"

4610    With Me
4620      If blnFromCalendar = False Then
4630        If IsDate(.SweepDate) = False Then
4640          .SweepDate = vbNullString
4650        End If
4660      Else
4670        blnFromCalendar = False
4680        .SweepDate.SelLength = 0
4690        .SweepDate.SelStart = 99
4700      End If
4710    End With

EXITP:
4720    Exit Sub

ERRH:
4730    DoCmd.Hourglass False
4740    Select Case ERR.Number
        Case Else
4750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4760    End Select
4770    Resume EXITP

End Sub

Private Sub SweepDate_BeforeUpdate(Cancel As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "SweepDate_BeforeUpdate"

4810    With Me
4820      If IsNull(.SweepDate) = False Then
4830        If IsDate(.SweepDate) = True Then
4840          If DateCheck_Post(.SweepDate) = False Then  ' ** Module Function: modUtilities.
4850            DoCmd.Hourglass False
                ' ** MsgBox already displayed.
4860            Cancel = -1
4870          End If
4880        End If
4890      End If
4900    End With

EXITP:
4910    Exit Sub

ERRH:
4920    DoCmd.Hourglass False
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub SweepDate_DblClick(Cancel As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "SweepDate_DblClick"

5010    With Me
5020      .SweepDate = Date
5030    End With

EXITP:
5040    Exit Sub

ERRH:
5050    Select Case ERR.Number
        Case Else
5060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5070    End Select
5080    Resume EXITP

End Sub

Private Sub SweepDate_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "SweepDate_KeyDown"

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
5190          blnKeyDown_Return = True
5200          .opgType.SetFocus
5210        End With
5220      End Select
5230    End If

        ' ** Shift keys.
5240    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5250      Select Case intRetVal
          Case vbKeyTab
5260        With Me
5270          intRetVal = 0
5280          .cmdClose.SetFocus
5290        End With
5300      Case vbKeyReturn
5310        With Me
5320          intRetVal = 0
5330          blnKeyDown_Return = True
5340          .cmdClose.SetFocus
5350        End With
5360      End Select
5370    End If

EXITP:
5380    KeyCode = intRetVal
5390    Exit Sub

ERRH:
5400    intRetVal = 0
5410    Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Probably a bad date. Ignore.
5420    Case Else
5430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5440    End Select
5450    Resume EXITP

End Sub

Private Sub SweepDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "SweepDate_MouseMove"

5510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5520    Exit Sub

ERRH:
5530    Select Case ERR.Number
        Case Else
5540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5550    End Select
5560    Resume EXITP

End Sub

Private Sub SweepDate_Exit(Cancel As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "SweepDate_Exit"

        Dim blnJustMove As Boolean

5610    With Me

5620      blnJustMove = False
5630      If IsNull(.SweepDate) = False Then
5640        If IsDate(.SweepDate) = True Then
              ' ** Proceed.
5650          blnJustMove = True
5660        End If
5670      End If

5680      If blnJustMove = False Then
5690        If blnKeyDown_Return = True Then
5700          blnKeyDown_Return = False
5710          .SweepDate = Date
5720        Else
5730          If IsNull(.SweepDate.text) = False Then
5740            If Trim(.SweepDate.text) <> vbNullString Then
                  ' ** Proceed.
5750            Else
                  ' ** Empty, do nothing.
5760            End If
5770          Else
                ' ** Empty, do nothing.
5780          End If
5790        End If
5800      End If

5810      If IsNull(.cmbAssets) = False Then
5820        If .cmbAssets > 0& Then
5830          If IsDate(.SweepDate) = True Then
5840            .cmdCreateEntries.Enabled = True
5850          Else
5860            .cmdCreateEntries.Enabled = False
5870          End If
5880        Else
5890          .cmdCreateEntries.Enabled = False
5900        End If
5910      Else
5920        .cmdCreateEntries.Enabled = False
5930      End If

5940    End With

EXITP:
5950    Exit Sub

ERRH:
5960    Select Case ERR.Number
        Case Else
5970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5980    End Select
5990    Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

6010    With Me
6020      blnCalendar1_Focus = True
6030      .cmdCalendar_raised_semifocus_dots_img.Visible = True
6040      .cmdCalendar_raised_img.Visible = False
6050      .cmdCalendar_raised_focus_img.Visible = False
6060      .cmdCalendar_raised_focus_dots_img.Visible = False
6070      .cmdCalendar_sunken_focus_dots_img.Visible = False
6080      .cmdCalendar_raised_img_dis.Visible = False
6090    End With

EXITP:
6100    Exit Sub

ERRH:
6110    Select Case ERR.Number
        Case Else
6120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6130    End Select
6140    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

6210    With Me
6220      blnCalendar1_MouseDown = True
6230      .cmdCalendar_sunken_focus_dots_img.Visible = True
6240      .cmdCalendar_raised_img.Visible = False
6250      .cmdCalendar_raised_semifocus_dots_img.Visible = False
6260      .cmdCalendar_raised_focus_img.Visible = False
6270      .cmdCalendar_raised_focus_dots_img.Visible = False
6280      .cmdCalendar_raised_img_dis.Visible = False
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

Private Sub cmdCalendar_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

6410    With Me
6420      datStartDate = Date
6430      datEndDate = 0
6440      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
6450      If blnRetVal = True Then
            ' ** Allow posting up to 1 month into the future.
6460        If datStartDate > DateAdd("m", 1, Date) Then
6470          MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
6480          .SweepDate = CDate(Format(Date, "mm/dd/yyyy"))
6490        Else
6500          blnFromCalendar = True
6510          .SweepDate = datStartDate
6520        End If
6530      Else
6540        blnFromCalendar = True
6550        .SweepDate = CDate(Format(Date, "mm/dd/yyyy"))
6560      End If
6570      .SweepDate.SetFocus
6580    End With

EXITP:
6590    Exit Sub

ERRH:
6600    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
6610    Case Else
6620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6630    End Select
6640    Resume EXITP

End Sub

Private Sub cmdCalendar_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_KeyDown"

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
6790          .opgType.SetFocus
6800        End With
6810      End Select
6820    End If

        ' ** Shift keys.
6830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6840      Select Case intRetVal
          Case vbKeyTab
6850        With Me
6860          intRetVal = 0
6870          .SweepDate.SetFocus
6880        End With
6890      End Select
6900    End If

EXITP:
6910    KeyCode = intRetVal
6920    Exit Sub

ERRH:
6930    intRetVal = 0
6940    Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Probably a bad date. Ignore.
6950    Case Else
6960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6970    End Select
6980    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

7010    With Me
7020      If blnCalendar1_MouseDown = False Then
7030        Select Case blnCalendar1_Focus
            Case True
7040          .cmdCalendar_raised_focus_dots_img.Visible = True
7050          .cmdCalendar_raised_focus_img.Visible = False
7060        Case False
7070          .cmdCalendar_raised_focus_img.Visible = True
7080          .cmdCalendar_raised_focus_dots_img.Visible = False
7090        End Select
7100        .cmdCalendar_raised_img.Visible = False
7110        .cmdCalendar_raised_semifocus_dots_img.Visible = False
7120        .cmdCalendar_sunken_focus_dots_img.Visible = False
7130        .cmdCalendar_raised_img_dis.Visible = False
7140      End If
7150    End With

EXITP:
7160    Exit Sub

ERRH:
7170    Select Case ERR.Number
        Case Else
7180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7190    End Select
7200    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

7310    With Me
7320      .cmdCalendar_raised_focus_dots_img.Visible = True
7330      .cmdCalendar_raised_img.Visible = False
7340      .cmdCalendar_raised_semifocus_dots_img.Visible = False
7350      .cmdCalendar_raised_focus_img.Visible = False
7360      .cmdCalendar_sunken_focus_dots_img.Visible = False
7370      .cmdCalendar_raised_img_dis.Visible = False
7380      blnCalendar1_MouseDown = False
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

Private Sub cmdCalendar_LostFocus()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

7510    With Me
7520      .cmdCalendar_raised_img.Visible = True
7530      .cmdCalendar_raised_semifocus_dots_img.Visible = False
7540      .cmdCalendar_raised_focus_img.Visible = False
7550      .cmdCalendar_raised_focus_dots_img.Visible = False
7560      .cmdCalendar_sunken_focus_dots_img.Visible = False
7570      .cmdCalendar_raised_img_dis.Visible = False
7580      blnCalendar1_Focus = False
7590    End With

EXITP:
7600    Exit Sub

ERRH:
7610    Select Case ERR.Number
        Case Else
7620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7630    End Select
7640    Resume EXITP

End Sub

Private Sub NoSweep_Journal_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "NoSweep_Journal_lbl_MouseMove"

7710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7720    Exit Sub

ERRH:
7730    Select Case ERR.Number
        Case Else
7740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7750    End Select
7760    Resume EXITP

End Sub

Private Sub NoSweep_Account_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "NoSweep_Account_lbl_MouseMove"

7810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7820    Exit Sub

ERRH:
7830    Select Case ERR.Number
        Case Else
7840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7850    End Select
7860    Resume EXITP

End Sub

Private Sub NoSweep_Asset_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "NoSweep_Asset_lbl_MouseMove"

7910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7920    Exit Sub

ERRH:
7930    Select Case ERR.Number
        Case Else
7940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7950    End Select
7960    Resume EXITP

End Sub

Private Sub opgType_AfterUpdate()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_AfterUpdate"

8010    With Me
8020      Select Case .opgType
          Case .opgType_optInvest.OptionValue
8030        .opgType_optInvest_lbl.FontBold = True
8040        .opgType_optCover_lbl.FontBold = False
8050        .opgCash.Enabled = True
8060      Case .opgType_optCover.OptionValue
8070        .opgType_optInvest_lbl.FontBold = False
8080        .opgType_optCover_lbl.FontBold = True
8090        .opgCash = opgCash_optPrincipal.OptionValue
8100        opgCash_AfterUpdate  ' ** Procedure: Below.
8110        .opgCash.Enabled = False
8120      End Select
8130    End With

EXITP:
8140    Exit Sub

ERRH:
8150    Select Case ERR.Number
        Case Else
8160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8170    End Select
8180    Resume EXITP

End Sub

Private Sub opgType_optInvest_KeyDown(KeyCode As Integer, Shift As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optInvest_KeyDown"

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
8290          If .opgCash.Enabled = True Then
8300            .opgCash.SetFocus
8310          ElseIf .cmbAssets.Enabled = True Then
8320            .cmbAssets.SetFocus
8330          ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
8340            .cmbRevenueCodes_Purchase.SetFocus
8350          ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
8360            .cmbTaxCodes_Purchase.SetFocus
8370          ElseIf .cmdCreateEntries.Enabled = True Then
8380            .cmdCreateEntries.SetFocus
8390          Else
8400            .cmdClose.SetFocus
8410          End If
8420        End With
8430      End Select
8440    End If

        ' ** Shift keys.
8450    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8470        With Me
8480          intRetVal = 0
8490          .SweepDate.SetFocus
8500        End With
8510      End Select
8520    End If

        ' ** Ctrl keys.
8530    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8550        With Me
8560          intRetVal = 0
8570          If .opgCash.Enabled = False Then
8580            If .opgAssetSource.Enabled = True Then
8590              .opgAssetSource.SetFocus
8600            ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
8610              .cmbRevenueCodes_Purchase.SetFocus
8620            ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
8630              .cmbTaxCodes_Purchase.SetFocus
8640            ElseIf .cmdCreateEntries.Enabled = True Then
8650              .cmdCreateEntries.SetFocus
8660            Else
8670              .cmdClose.SetFocus
8680            End If
8690          End If
8700        End With
8710      End Select
8720    End If

EXITP:
8730    KeyCode = intRetVal
8740    Exit Sub

ERRH:
8750    intRetVal = 0
8760    Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Probably a bad date. Ignore.
8770    Case Else
8780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8790    End Select
8800    Resume EXITP

End Sub

Private Sub opgType_optCover_KeyDown(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optCover_KeyDown"

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
8990          If .opgCash.Enabled = True Then
9000            .opgCash.SetFocus
9010          ElseIf .cmbAssets.Enabled = True Then
9020            .cmbAssets.SetFocus
9030          ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
9040            .cmbRevenueCodes_Purchase.SetFocus
9050          ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
9060            .cmbTaxCodes_Purchase.SetFocus
9070          ElseIf .cmdCreateEntries.Enabled = True Then
9080            .cmdCreateEntries.SetFocus
9090          Else
9100            .cmdClose.SetFocus
9110          End If
9120        End With
9130      End Select
9140    End If

        ' ** Shift keys.
9150    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9170        With Me
9180          intRetVal = 0
9190          .SweepDate.SetFocus
9200        End With
9210      End Select
9220    End If

        ' ** Ctrl keys.
9230    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
9240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9250        With Me
9260          intRetVal = 0
9270          If .opgCash.Enabled = False Then
9280            If .opgAssetSource.Enabled = True Then
9290              .opgAssetSource.SetFocus
9300            ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
9310              .cmbRevenueCodes_Purchase.SetFocus
9320            ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
9330              .cmbTaxCodes_Purchase.SetFocus
9340            ElseIf .cmdCreateEntries.Enabled = True Then
9350              .cmdCreateEntries.SetFocus
9360            Else
9370              .cmdClose.SetFocus
9380            End If
9390          End If
9400        End With
9410      End Select
9420    End If

EXITP:
9430    KeyCode = intRetVal
9440    Exit Sub

ERRH:
9450    intRetVal = 0
9460    Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Probably a bad date. Ignore.
9470    Case Else
9480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9490    End Select
9500    Resume EXITP

End Sub

Private Sub opgCash_AfterUpdate()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgCash_AfterUpdate"

9610    With Me
9620      Select Case .opgCash
          Case .opgCash_optIncome.OptionValue
9630        .opgCash_optIncome_lbl.FontBold = True
9640        .opgCash_optPrincipal_lbl.FontBold = False
9650      Case .opgCash_optPrincipal.OptionValue
9660        .opgCash_optIncome_lbl.FontBold = False
9670        .opgCash_optPrincipal_lbl.FontBold = True
9680      End Select
9690    End With

EXITP:
9700    Exit Sub

ERRH:
9710    Select Case ERR.Number
        Case Else
9720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9730    End Select
9740    Resume EXITP

End Sub

Private Sub opgCash_optIncome_KeyDown(KeyCode As Integer, Shift As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgCash_optIncome_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9820    intShiftDown = (Shift And acShiftMask) > 0
9830    intAltDown = (Shift And acAltMask) > 0
9840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9870        With Me
9880          intRetVal = 0
9890          If .cmbAssets.Enabled = True Then
9900            .cmbAssets.SetFocus
9910          ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
9920            .cmbRevenueCodes_Purchase.SetFocus
9930          ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
9940            .cmbTaxCodes_Purchase.SetFocus
9950          ElseIf .cmdCreateEntries.Enabled = True Then
9960            .cmdCreateEntries.SetFocus
9970          Else
9980            .cmdClose.SetFocus
9990          End If
10000       End With
10010     End Select
10020   End If

        ' ** Shift keys.
10030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10050       With Me
10060         intRetVal = 0
10070         .opgType.SetFocus
10080       End With
10090     End Select
10100   End If

        ' ** Ctrl keys.
10110   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10130       With Me
10140         intRetVal = 0
10150         If .opgAssetSource.Enabled = True Then
10160           .opgAssetSource.SetFocus
10170         ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
10180           .cmbRevenueCodes_Purchase.SetFocus
10190         ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
10200           .cmbTaxCodes_Purchase.SetFocus
10210         ElseIf .cmdCreateEntries.Enabled = True Then
10220           .cmdCreateEntries.SetFocus
10230         Else
10240           .cmdClose.SetFocus
10250         End If
10260       End With
10270     End Select
10280   End If

EXITP:
10290   KeyCode = intRetVal
10300   Exit Sub

ERRH:
10310   intRetVal = 0
10320   Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Probably a bad date. Ignore.
10330   Case Else
10340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10350   End Select
10360   Resume EXITP

End Sub

Private Sub opgCash_optPrincipal_KeyDown(KeyCode As Integer, Shift As Integer)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCash_optPrincipal_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10420   intShiftDown = (Shift And acShiftMask) > 0
10430   intAltDown = (Shift And acAltMask) > 0
10440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10470       With Me
10480         intRetVal = 0
10490         If .cmbAssets.Enabled = True Then
10500           .cmbAssets.SetFocus
10510         ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
10520           .cmbRevenueCodes_Purchase.SetFocus
10530         ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
10540           .cmbTaxCodes_Purchase.SetFocus
10550         ElseIf .cmdCreateEntries.Enabled = True Then
10560           .cmdCreateEntries.SetFocus
10570         Else
10580           .cmdClose.SetFocus
10590         End If
10600       End With
10610     End Select
10620   End If

        ' ** Shift keys.
10630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10650       With Me
10660         intRetVal = 0
10670         .opgType.SetFocus
10680       End With
10690     End Select
10700   End If

        ' ** Ctrl keys.
10710   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10720     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10730       With Me
10740         intRetVal = 0
10750         If .opgAssetSource.Enabled = True Then
10760           .opgAssetSource.SetFocus
10770         ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
10780           .cmbRevenueCodes_Purchase.SetFocus
10790         ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
10800           .cmbTaxCodes_Purchase.SetFocus
10810         ElseIf .cmdCreateEntries.Enabled = True Then
10820           .cmdCreateEntries.SetFocus
10830         Else
10840           .cmdClose.SetFocus
10850         End If
10860       End With
10870     End Select
10880   End If

EXITP:
10890   KeyCode = intRetVal
10900   Exit Sub

ERRH:
10910   intRetVal = 0
10920   Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Probably a bad date. Ignore.
10930   Case Else
10940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10950   End Select
10960   Resume EXITP

End Sub

Private Sub opgAssetSource_AfterUpdate()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_AfterUpdate"

11010   With Me
11020     Select Case .opgAssetSource
          Case .opgAssetSource_optType.OptionValue
11030       .opgAssetSource_optType_lbl2.FontBold = True
11040       .opgAssetSource_optType_lbl2_dim_hi.FontBold = True
11050       .opgAssetSource_optName_lbl2.FontBold = False
11060       .opgAssetSource_optName_lbl2_dim_hi.FontBold = False
11070       .opgAssetSource_optCusip_lbl2.FontBold = False
11080       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = False
11090       If .cmbAssets.RowSource <> "qrySweeper_05a" Then
11100         .cmbAssets.RowSource = "qrySweeper_05a"
11110       End If
11120       .cmbAssets.ListWidth = 6810
11130       .cmbAssets.ColumnWidths = "0;5010;1185;360;0"
11140     Case .opgAssetSource_optName.OptionValue
11150       .opgAssetSource_optType_lbl2.FontBold = False
11160       .opgAssetSource_optType_lbl2_dim_hi.FontBold = False
11170       .opgAssetSource_optName_lbl2.FontBold = True
11180       .opgAssetSource_optName_lbl2_dim_hi.FontBold = True
11190       .opgAssetSource_optCusip_lbl2.FontBold = False
11200       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = False
11210       If .cmbAssets.RowSource <> "qrySweeper_05b" Then
11220         .cmbAssets.RowSource = "qrySweeper_05b"
11230       End If
11240       .cmbAssets.ListWidth = 6810
11250       .cmbAssets.ColumnWidths = "0;5010;1185;360;0"
11260     Case .opgAssetSource_optCusip.OptionValue
11270       .opgAssetSource_optType_lbl2.FontBold = False
11280       .opgAssetSource_optType_lbl2_dim_hi.FontBold = False
11290       .opgAssetSource_optName_lbl2.FontBold = False
11300       .opgAssetSource_optName_lbl2_dim_hi.FontBold = False
11310       .opgAssetSource_optCusip_lbl2.FontBold = True
11320       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = True
11330       If .cmbAssets.RowSource <> "qrySweeper_05c" Then
11340         .cmbAssets.RowSource = "qrySweeper_05c"
11350       End If
11360       .cmbAssets.ListWidth = 6810
11370       .cmbAssets.ColumnWidths = "0;6195;0;360;0"
11380     End Select
11390   End With

EXITP:
11400   Exit Sub

ERRH:
11410   Select Case ERR.Number
        Case Else
11420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11430   End Select
11440   Resume EXITP

End Sub

Private Sub opgAssetSource_optType_KeyDown(KeyCode As Integer, Shift As Integer)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11520   intShiftDown = (Shift And acShiftMask) > 0
11530   intAltDown = (Shift And acAltMask) > 0
11540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11570       With Me
11580         intRetVal = 0
11590         .chkRememberAsset.SetFocus
11600       End With
11610     End Select
11620   End If

        ' ** Shift keys.
11630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11650       With Me
11660         intRetVal = 0
11670         If .opgCash.Enabled = True Then
11680           .opgCash.SetFocus
11690         Else
11700           .opgType.SetFocus
11710         End If
11720       End With
11730     End Select
11740   End If

EXITP:
11750   KeyCode = intRetVal
11760   Exit Sub

ERRH:
11770   intRetVal = 0
11780   Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Probably a bad date. Ignore.
11790   Case Else
11800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11810   End Select
11820   Resume EXITP

End Sub

Private Sub opgAssetSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11920   intShiftDown = (Shift And acShiftMask) > 0
11930   intAltDown = (Shift And acAltMask) > 0
11940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11970       With Me
11980         intRetVal = 0
11990         .chkRememberAsset.SetFocus
12000       End With
12010     End Select
12020   End If

        ' ** Shift keys.
12030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12050       With Me
12060         intRetVal = 0
12070         If .opgCash.Enabled = True Then
12080           .opgCash.SetFocus
12090         Else
12100           .opgType.SetFocus
12110         End If
12120       End With
12130     End Select
12140   End If

EXITP:
12150   KeyCode = intRetVal
12160   Exit Sub

ERRH:
12170   intRetVal = 0
12180   Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Probably a bad date. Ignore.
12190   Case Else
12200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12210   End Select
12220   Resume EXITP

End Sub

Private Sub opgAssetSource_optCusip_KeyDown(KeyCode As Integer, Shift As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optCusip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12320   intShiftDown = (Shift And acShiftMask) > 0
12330   intAltDown = (Shift And acAltMask) > 0
12340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12370       With Me
12380         intRetVal = 0
12390         .chkRememberAsset.SetFocus
12400       End With
12410     End Select
12420   End If

        ' ** Shift keys.
12430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12450       With Me
12460         intRetVal = 0
12470         If .opgCash.Enabled = True Then
12480           .opgCash.SetFocus
12490         Else
12500           .opgType.SetFocus
12510         End If
12520       End With
12530     End Select
12540   End If

EXITP:
12550   KeyCode = intRetVal
12560   Exit Sub

ERRH:
12570   intRetVal = 0
12580   Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Probably a bad date. Ignore.
12590   Case Else
12600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12610   End Select
12620   Resume EXITP

End Sub

Private Sub chkRememberAsset_AfterUpdate()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberAsset_AfterUpdate"

12710   With Me
12720     Select Case .chkRememberAsset
          Case True
12730       .chkRememberAsset_lbl.FontBold = True
12740       .chkRememberAsset_lbl2_dim.FontBold = True
12750       .chkRememberAsset_lbl2_dim_hi.FontBold = True
12760     Case False
12770       .chkRememberAsset_lbl.FontBold = False
12780       .chkRememberAsset_lbl2_dim.FontBold = False
12790       .chkRememberAsset_lbl2_dim_hi.FontBold = False
12800     End Select
12810   End With

EXITP:
12820   Exit Sub

ERRH:
12830   Select Case ERR.Number
        Case Else
12840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12850   End Select
12860   Resume EXITP

End Sub

Private Sub chkRememberAsset_KeyDown(KeyCode As Integer, Shift As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberAsset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12920   intShiftDown = (Shift And acShiftMask) > 0
12930   intAltDown = (Shift And acAltMask) > 0
12940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12970       With Me
12980         intRetVal = 0
12990         .cmbAssets.SetFocus
13000       End With
13010     End Select
13020   End If

        ' ** Shift keys.
13030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13050       With Me
13060         intRetVal = 0
13070         .opgAssetSource.SetFocus
13080       End With
13090     End Select
13100   End If

EXITP:
13110   KeyCode = intRetVal
13120   Exit Sub

ERRH:
13130   intRetVal = 0
13140   Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Probably a bad date. Ignore.
13150   Case Else
13160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13170   End Select
13180   Resume EXITP

End Sub

Private Sub cmbAssets_Enter()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_Enter"

13210   With Me
13220     If IsNull(.cmbAssets) = True Or lngCntSweeps > 1& Then
13230       ExpandCombo .cmbAssets  ' ** Module Procedure: modUtilities.
13240     End If
13250   End With

EXITP:
13260   Exit Sub

ERRH:
13270   Select Case ERR.Number
        Case Else
13280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13290   End Select
13300   Resume EXITP

End Sub

Private Sub cmbAssets_AfterUpdate()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

13410   With Me
13420     If IsNull(.cmbAssets) = False And IsNull(.SweepDate) = False Then
13430       If .cmbAssets > 0 And IsDate(.SweepDate) = True Then
13440         .cmdCreateEntries.Enabled = True
13450       Else
13460         .cmdCreateEntries.Enabled = False
13470       End If
13480     Else
13490       .cmdCreateEntries.Enabled = False
13500     End If
13510   End With

EXITP:
13520   Exit Sub

ERRH:
13530   Select Case ERR.Number
        Case Else
13540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13550   End Select
13560   Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13620   intShiftDown = (Shift And acShiftMask) > 0
13630   intAltDown = (Shift And acAltMask) > 0
13640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13670       With Me
13680         intRetVal = 0
13690         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
13700           .cmbRevenueCodes_Purchase.SetFocus
13710         ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
13720           .cmbTaxCodes_Purchase.SetFocus
13730         ElseIf .cmdCreateEntries.Enabled = True Then
13740           .cmdCreateEntries.SetFocus
13750         Else
13760           .cmdClose.SetFocus
13770         End If
13780       End With
13790     End Select
13800   End If

        ' ** Shift keys.
13810   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13820     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13830       With Me
13840         intRetVal = 0
13850         If .opgCash.Enabled = True Then
13860           .opgCash.SetFocus
13870         Else
13880           .opgType.SetFocus
13890         End If
13900       End With
13910     End Select
13920   End If

        ' ** Ctrl-Shift keys.
13930   If intCtrlDown And (Not intAltDown) And intShiftDown Then
13940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13950       With Me
13960         intRetVal = 0
13970         .chkRememberAsset.SetFocus
13980       End With
13990     End Select
14000   End If

EXITP:
14010   KeyCode = intRetVal
14020   Exit Sub

ERRH:
14030   intRetVal = 0
14040   Select Case ERR.Number
        Case Else
14050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14060   End Select
14070   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_GotFocus()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_GotFocus"

14110   With Me
14120     .cmbRevenueCodes_Purchase.SetFocus
14130   End With

EXITP:
14140   Exit Sub

ERRH:
14150   Select Case ERR.Number
        Case Else
14160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14170   End Select
14180   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14220   intShiftDown = (Shift And acShiftMask) > 0
14230   intAltDown = (Shift And acAltMask) > 0
14240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14270       With Me
14280         intRetVal = 0
14290         .cmbRevenueCodes_Purchase.SetFocus
14300       End With
14310     End Select
14320   End If

        ' ** Shift keys.
14330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14350       With Me
14360         intRetVal = 0
14370         .cmbAssets.SetFocus
14380       End With
14390     End Select
14400   End If

EXITP:
14410   KeyCode = intRetVal
14420   Exit Sub

ERRH:
14430   intRetVal = 0
14440   Select Case ERR.Number
        Case Else
14450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14460   End Select
14470   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_Enter()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_Enter"

14510   With Me
14520     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
14530       ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
14540     Case False
14550       If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
14560         ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
14570       End If
14580     End Select
14590   End With

EXITP:
14600   Exit Sub

ERRH:
14610   Select Case ERR.Number
        Case Else
14620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14630   End Select
14640   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_AfterUpdate()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_AfterUpdate"

14710   With Me
14720     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
14730       .cmbRevenueCodes_Purchase = REVID_INC
14740     Case False
            ' ** Nothing at the moment.
14750     End Select
14760     If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
14770       .cmbRevenueCodes_Purchase_display = Null
14780     Else
14790       .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
14800     End If
14810   End With

EXITP:
14820   Exit Sub

ERRH:
14830   Select Case ERR.Number
        Case Else
14840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14850   End Select
14860   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14920   intShiftDown = (Shift And acShiftMask) > 0
14930   intAltDown = (Shift And acAltMask) > 0
14940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14970       With Me
14980         intRetVal = 0
14990         If .cmbTaxCodes_Purchase.Enabled = True Then
15000           .cmbTaxCodes_Purchase.SetFocus
15010         ElseIf .cmdCreateEntries.Enabled = True Then
15020           .cmdCreateEntries.SetFocus
15030         Else
15040           .cmdClose.SetFocus
15050         End If
15060       End With
15070     End Select
15080   End If

        ' ** Shift keys.
15090   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15110       With Me
15120         intRetVal = 0
15130         .cmbAssets.SetFocus
15140       End With
15150     End Select
15160   End If

EXITP:
15170   KeyCode = intRetVal
15180   Exit Sub

ERRH:
15190   intRetVal = 0
15200   Select Case ERR.Number
        Case Else
15210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15220   End Select
15230   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_GotFocus()

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_GotFocus"

15310   With Me
15320     .cmbTaxCodes_Purchase.SetFocus
15330   End With

EXITP:
15340   Exit Sub

ERRH:
15350   Select Case ERR.Number
        Case Else
15360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15370   End Select
15380   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15420   intShiftDown = (Shift And acShiftMask) > 0
15430   intAltDown = (Shift And acAltMask) > 0
15440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15470       With Me
15480         intRetVal = 0
15490         .cmbTaxCodes_Purchase.SetFocus
15500       End With
15510     End Select
15520   End If

        ' ** Shift keys.
15530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15550       With Me
15560         intRetVal = 0
15570         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
15580           .cmbRevenueCodes_Purchase.SetFocus
15590         Else
15600           .cmbAssets.SetFocus
15610         End If
15620       End With
15630     End Select
15640   End If

EXITP:
15650   KeyCode = intRetVal
15660   Exit Sub

ERRH:
15670   intRetVal = 0
15680   Select Case ERR.Number
        Case Else
15690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15700   End Select
15710   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_Enter()

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_Enter"

15810   With Me
15820     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
15830       ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
15840     Case False
15850       If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
15860         ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
15870       End If
15880     End Select
15890   End With

EXITP:
15900   Exit Sub

ERRH:
15910   Select Case ERR.Number
        Case Else
15920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15930   End Select
15940   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_AfterUpdate()

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_AfterUpdate"

16010   With Me
16020     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
16030       .cmbTaxCodes_Purchase = TAXID_INC
16040     Case False
            ' ** Nothing at the moment.
16050     End Select
16060     If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
16070       .cmbTaxCodes_Purchase_display = Null
16080     Else
16090       .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
16100     End If
16110   End With

EXITP:
16120   Exit Sub

ERRH:
16130   Select Case ERR.Number
        Case Else
16140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16150   End Select
16160   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16220   intShiftDown = (Shift And acShiftMask) > 0
16230   intAltDown = (Shift And acAltMask) > 0
16240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16270       With Me
16280         intRetVal = 0
16290         If .cmdCreateEntries.Enabled = True Then
16300           .cmdCreateEntries.SetFocus
16310         Else
16320           .cmdClose.SetFocus
16330         End If
16340       End With
16350     End Select
16360   End If

        ' ** Shift keys.
16370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16390       With Me
16400         intRetVal = 0
16410         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
16420           .cmbRevenueCodes_Purchase.SetFocus
16430         Else
16440           .cmbAssets.SetFocus
16450         End If
16460       End With
16470     End Select
16480   End If

EXITP:
16490   KeyCode = intRetVal
16500   Exit Sub

ERRH:
16510   intRetVal = 0
16520   Select Case ERR.Number
        Case Else
16530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16540   End Select
16550   Resume EXITP

End Sub

Private Sub cmdCreateEntries_Click()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCreateEntries_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstAccount As DAO.Recordset, rstJournal As DAO.Recordset
        Dim rstActiveAssets As DAO.Recordset
        Dim strMsg As String, intStyle As Integer, strTitle1 As String
        Dim dblAmountToSweep As Double, dblAmountToSweepHere As Double
        Dim blnOut As Boolean, blnSomeSwept As Boolean
        Dim lngRecs As Long
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant
        Dim lngX As Long
        Dim blnContinue As Boolean

16610   blnContinue = True
16620   blnSomeSwept = False  ' ** Unless proven otherwise.

16630   With Me
16640     If IsNull(.cmbAssets.Value) Then
            ' ** Have to choose an asset.
16650       blnContinue = False
16660       strMsg = "You must choose an asset to sweep " & IIf(.opgType = .opgType_optInvest.OptionValue, "to.", "from.")
16670       intStyle = vbQuestion + vbOKOnly
16680       strTitle1 = "Sweep Cash"
16690       MsgBox strMsg, intStyle, strTitle1
16700       .cmbAssets.SetFocus
16710     Else

16720       DoCmd.Hourglass True
16730       DoEvents

16740       Set dbs = CurrentDb

16750       Select Case .opgCash
            Case .opgCash_optIncome.OptionValue
              ' ** Income cash.
16760         Select Case .opgType
              Case .opgType_optInvest.OptionValue
                ' ** Account, just icash > 0, account_SWEEP = True.
16770           Set qdf = dbs.QueryDefs("qrySweeper_01")
16780           Set rstAccount = qdf.OpenRecordset
16790         Case .opgType_optCover.OptionValue   ' ** Reverse.
                ' ** Account, just icash < 0, account_SWEEP = True.
16800           Set qdf = dbs.QueryDefs("qrySweeper_02")
16810           Set rstAccount = qdf.OpenRecordset
16820         End Select
16830       Case .opgCash_optPrincipal.OptionValue  ' ** Principal cash.
16840         Select Case .opgType
              Case .opgType_optInvest.OptionValue
                ' ** Account, just pcash > 0, account_SWEEP = True.
16850           Set qdf = dbs.QueryDefs("qrySweeper_03")
16860           Set rstAccount = qdf.OpenRecordset
16870         Case .opgType_optCover.OptionValue  ' ** Reverse
                ' ** Account, just pcash < 0, account_SWEEP = True.
16880           Set qdf = dbs.QueryDefs("qrySweeper_04")
16890           Set rstAccount = qdf.OpenRecordset
16900         End Select
16910       End Select

16920       If rstAccount.BOF = True And rstAccount.EOF = True Then
16930         blnContinue = False
16940         DoCmd.Hourglass False
16950         strMsg = "No accounts to sweep!"
16960         intStyle = vbExclamation + vbOKOnly + vbDefaultButton1
16970         strTitle1 = ("Nothing To Do" & Space(40))
16980         MsgBox strMsg, intStyle, strTitle1
16990       Else
17000         varTmp00 = DCount("*", "journal")
17010         If IsNull(varTmp00) = False Then
17020           If varTmp00 > 0 Then
17030             blnContinue = False
17040             DoCmd.Hourglass False
17050             strMsg = "Existing journal entries must be posted or cleared before performing this operation."
17060             intStyle = vbExclamation + vbOKOnly + vbDefaultButton1
17070             strTitle1 = "Post Journal"
17080             MsgBox strMsg, intStyle, strTitle1
17090           End If
17100         End If
17110       End If
17120       If blnContinue = True Then
17130         DoCmd.Hourglass False
17140         msgResponse = MsgBox("Do you want to create journal entries with a posting date of " & CStr(.SweepDate.Value) & "?" & vbCrLf & vbCrLf & _
                "If you want a different date, click 'no' and change posting date to desired date.", _
                vbQuestion + vbYesNo, "Confirm New Entries")
17150         If msgResponse = vbYes Then
                ' ** Process the accounts.
17160           DoCmd.Hourglass True
17170           DoEvents
17180           Set rstJournal = dbs.OpenRecordset("journal", dbOpenDynaset, dbConsistent)
17190           rstAccount.MoveLast
17200           lngRecs = rstAccount.RecordCount
17210           rstAccount.MoveFirst
17220           For lngX = 1& To lngRecs
17230             Select Case .opgType
                  Case .opgType_optInvest.OptionValue  ' ** Regular sweep.
                    ' ** Create a purchase entry in the journal for each account.
17240               rstJournal.AddNew
17250               rstJournal![assetno] = .cmbAssets
17260               rstJournal![accountno] = rstAccount![accountno]
17270               rstJournal![assetdate] = CDate(.SweepDate + time())
17280               rstJournal![journaltype] = "Purchase"
17290               rstJournal![transdate] = CDate(.SweepDate)
17300               Select Case .opgCash
                    Case .opgCash_optIncome.OptionValue  ' ** Income cash.
17310                 rstJournal![ICash] = CCur(Round(CDbl(rstAccount![ICash]), 2) * -1#)
17320                 rstJournal![PCash] = 0@
17330                 rstJournal![Cost] = CCur(Round(CDbl(rstAccount![ICash]), 2))
17340                 rstJournal![shareface] = CDbl(Round(Abs(CDbl(rstAccount![ICash])), 2))  ' ** Yes, rounded to 2.
17350                 rstJournal![pershare] = 1#
17360               Case .opgCash_optPrincipal.OptionValue  ' ** Principal cash.
17370                 rstJournal![ICash] = 0@
17380                 rstJournal![PCash] = CCur(Round(CDbl(rstAccount![PCash]), 2) * -1#)
17390                 rstJournal![Cost] = CCur(Round(CDbl(rstAccount![PCash]), 2))
17400                 rstJournal![shareface] = CDbl(Round(Abs(CDbl(rstAccount![PCash])), 2))  ' ** Yes, rounded to 2.
17410                 rstJournal![pershare] = 1#
17420               End Select
17430               rstJournal![rate] = 0#
17440               rstJournal![Location_ID] = 1&
17450               rstJournal![revcode_ID] = Nz(Me.cmbRevenueCodes_Purchase, REVID_INC)
17460               rstJournal![taxcode] = Nz(Me.cmbTaxCodes_Purchase, TAXID_INC)
17470               rstJournal![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
17480               rstJournal.Update
17490             Case .opgType_optCover.OptionValue  ' ** Reverse sweep.
                    ' ** Create a sold entry in the journal for each account.
17500               Set rstActiveAssets = dbs.OpenRecordset("SELECT activeassets.* FROM activeassets WHERE accountno = '" & _
                      rstAccount![accountno] & "' AND assetno = " & CStr(.cmbAssets.Value) & " ORDER BY assetdate;", dbOpenSnapshot)
17510               If rstActiveAssets.RecordCount > 0 Then  ' ** They own some of it.
17520                 rstActiveAssets.MoveFirst
17530                 dblAmountToSweep = CDbl(Round(Abs(IIf(.opgCash = .opgCash_optIncome.OptionValue, _
                        CDbl(rstAccount![ICash]), CDbl(rstAccount![PCash]))), 2))  ' ** Yes, rounded to 2 here.
17540                 blnOut = False
17550                 blnSomeSwept = True  ' ** Now, at least one journal record will be created.
                      ' ** Process through rstActiveAssets until dblAmountToSweep is zero.
17560                 Do While Round(dblAmountToSweep, 4) > 0 And blnOut = False
17570                   dblAmountToSweepHere = IIf(Round(CDbl(rstActiveAssets![shareface]), 4) > dblAmountToSweep, _
                          dblAmountToSweep, Round(CDbl(rstActiveAssets![shareface]), 4))  ' ** Rounded to 4 within the loop.
17580                   rstJournal.AddNew
17590                   rstJournal![assetno] = .cmbAssets
17600                   rstJournal![accountno] = rstAccount![accountno]
17610                   rstJournal![assetdate] = CDate(.SweepDate + time())
17620                   rstJournal![journaltype] = "Sold"
17630                   rstJournal![transdate] = CDate(.SweepDate)
17640                   Select Case .opgCash
                        Case .opgCash_optIncome.OptionValue  ' ** Income cash.
                          ' ** Do nothing.
17650                   Case .opgCash_optPrincipal.OptionValue  ' ** Principal cash.
17660                     rstJournal![PCash] = CCur(Round(dblAmountToSweepHere, 2))
17670                     rstJournal![ICash] = 0@
17680                     rstJournal![Cost] = CCur((dblAmountToSweepHere * _
                            (Round(CDbl(rstActiveAssets![Cost]), 2) / Round(CDbl(rstActiveAssets![shareface]), 4))) * -1)
17690                     rstJournal![shareface] = dblAmountToSweepHere
17700                     rstJournal![pershare] = 0#
17710                     rstJournal![PurchaseDate] = rstActiveAssets![assetdate]
17720                     rstJournal![rate] = 0#
17730                     rstJournal![Location_ID] = 1&
17740                     rstJournal![revcode_ID] = Nz(Me.cmbRevenueCodes_Purchase, REVID_INC)
17750                     rstJournal![taxcode] = Nz(Me.cmbTaxCodes_Purchase, TAXID_INC)
17760                     rstJournal![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
17770                   End Select
17780                   If IsNull(rstJournal![shareface]) = True Then rstJournal![shareface] = 0#
17790                   If IsNull(rstJournal![ICash]) = True Then rstJournal![ICash] = 0@
17800                   If IsNull(rstJournal![PCash]) = True Then rstJournal![PCash] = 0@
17810                   If IsNull(rstJournal![Cost]) = True Then rstJournal![Cost] = 0@
17820                   If rstJournal![shareface] = 0# And rstJournal![ICash] = 0@ And rstJournal![PCash] = 0@ And rstJournal![Cost] = 0@ Then
17830                     rstJournal.CancelUpdate
                          ' ** DbUpdate enumeration for CancelUpdate:
                          ' **   1  dbUpdateRegular        Cancels pending changes that aren’t cached. (Default)
                          ' **   4  dbUpdateBatch          Cancels pending changes in the update cache.
                          ' ** DbUpdate enumeration for Update:
                          ' **   1  dbUpdateRegular        Pending changes are not cached and are written to disk immediately. (Default)
                          ' **   2  dbUpdateCurrentRecord  Only the current record's pending changes are written to disk.
                          ' **   4  dbUpdateBatch          All pending changes in the update cache are written to disk.
17840                   Else
17850                     rstJournal.Update
17860                   End If
17870                   dblAmountToSweep = Round((dblAmountToSweep - dblAmountToSweepHere), 4)
17880                   rstActiveAssets.MoveNext
17890                   If rstActiveAssets.EOF Then
17900                     blnOut = True
17910                   End If
17920                 Loop
17930                 rstActiveAssets.Close
17940               End If
17950             End Select
17960             If lngX < lngRecs Then rstAccount.MoveNext
17970           Next
17980           strMsg = "Operation complete!"
17990           If .opgType = .opgType_optCover.OptionValue And blnSomeSwept = False Then
                  ' ** Did reverse sweep, but no overdraft accounts owned the sweep asset.
18000             blnContinue = False
18010             strMsg = strMsg & vbCrLf & vbCrLf & _
                    "(However, no journal entries were created because no overdraft accounts owned the sweep asset.)"
18020           End If
18030           intStyle = vbInformation + vbOKOnly + vbDefaultButton1
18040           strTitle1 = "Sweep Cash"
18050           DoCmd.Hourglass False
18060           MsgBox strMsg, intStyle, strTitle1
18070           rstJournal.Close
18080           If blnContinue = True Then
18090             DoCmd.Close acForm, THIS_NAME
18100           End If
18110         End If ' ** msgResponse.
18120       End If  ' ** blnContinue.
18130       rstAccount.Close
18140       dbs.Close
18150     End If  ' ** cmbAssets.

18160   End With

18170   DoCmd.Hourglass False

EXITP:
18180   Set rstActiveAssets = Nothing
18190   Set rstJournal = Nothing
18200   Set rstAccount = Nothing
18210   Set qdf = Nothing
18220   Set dbs = Nothing
18230   Exit Sub

ERRH:
18240   DoCmd.Hourglass False
18250   Select Case ERR.Number
        Case Else
18260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18270   End Select
18280   Resume EXITP

End Sub

Private Sub cmdCreateEntries_KeyDown(KeyCode As Integer, Shift As Integer)

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCreateEntries_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18320   intShiftDown = (Shift And acShiftMask) > 0
18330   intAltDown = (Shift And acAltMask) > 0
18340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18360     Select Case intRetVal
          Case vbKeyTab
18370       With Me
18380         intRetVal = 0
18390         .cmdClose.SetFocus
18400       End With
18410     End Select
18420   End If

        ' ** Shift keys.
18430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18440     Select Case intRetVal
          Case vbKeyTab
18450       With Me
18460         intRetVal = 0
18470         If .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
18480           .cmbTaxCodes_Purchase.SetFocus
18490         ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
18500           .cmbRevenueCodes_Purchase.SetFocus
18510         Else
18520           .cmbAssets.SetFocus
18530         End If
18540       End With
18550     End Select
18560   End If

EXITP:
18570   KeyCode = intRetVal
18580   Exit Sub

ERRH:
18590   intRetVal = 0
18600   Select Case ERR.Number
        Case Else
18610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18620   End Select
18630   Resume EXITP

End Sub

Private Sub HideIncExp()

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "HideIncExp"

        Dim lngOffset1 As Long, lngOffset2 As Long
        Dim lngTmp01 As Long, lngTmp02 As Long

18710   With Me

18720     .cmbRevenueCodes_Purchase_box.Visible = False
18730     .cmbRevenueCodes_Purchase_box2.Visible = False
18740     .cmbRevenueCodes_Purchase_box_lbl.Visible = False
18750     .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Visible = False
18760     .cmbRevenueCodes_Purchase_hline01.Visible = False
18770     .cmbRevenueCodes_Purchase_hline02.Visible = False
18780     .cmbRevenueCodes_Purchase_hline03.Visible = False
18790     .cmbRevenueCodes_Purchase_vline01.Visible = False
18800     .cmbRevenueCodes_Purchase_vline02.Visible = False
18810     .cmbRevenueCodes_Purchase_vline03.Visible = False
18820     .cmbRevenueCodes_Purchase_vline04.Visible = False
18830     .cmbRevenueCodes_Purchase.Visible = False
18840     .cmbRevenueCodes_Purchase_display.Visible = False
18850     .cmbRevenueCodes_Purchase_lbl.Visible = False
18860     .cmbRevenueCodes_Purchase_lbl2.Visible = False
18870     .cmbTaxCodes_Purchase.Visible = False
18880     .cmbTaxCodes_Purchase_display.Visible = False
18890     .cmbTaxCodes_Purchase_lbl.Visible = False
18900     .cmbTaxCodes_Purchase_lbl2.Visible = False

18910     lngOffset1 = ((.cmbAssets_box.Top + .cmbAssets_box.Height) - .cmbRevenueCodes_Purchase_box.Height)
18920     lngTmp01 = (.Detail.Height - (.cmbRevenueCodes_Purchase_box.Top + .cmbRevenueCodes_Purchase_box.Height))
18930     lngTmp01 = ((.cmbAssets_box.Top + .cmbAssets_box.Height) + lngTmp01)
18940     lngOffset2 = (.Detail.Height - lngTmp01)

18950     .cmbRevenueCodes_Purchase_box.Top = .cmbRevenueCodes_Purchase_box.Top - lngOffset1
18960     .cmbRevenueCodes_Purchase_box2.Top = .cmbRevenueCodes_Purchase_box2.Top - lngOffset1
18970     .cmbRevenueCodes_Purchase_box_lbl.Top = .cmbRevenueCodes_Purchase_box_lbl.Top - lngOffset1
18980     .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Top = .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Top - lngOffset1
18990     .cmbRevenueCodes_Purchase_hline01.Top = .cmbRevenueCodes_Purchase_hline01.Top - lngOffset1
19000     .cmbRevenueCodes_Purchase_hline02.Top = .cmbRevenueCodes_Purchase_hline02.Top - lngOffset1
19010     .cmbRevenueCodes_Purchase_hline03.Top = .cmbRevenueCodes_Purchase_hline03.Top - lngOffset1
19020     .cmbRevenueCodes_Purchase_vline01.Top = .cmbRevenueCodes_Purchase_vline01.Top - lngOffset1
19030     .cmbRevenueCodes_Purchase_vline02.Top = .cmbRevenueCodes_Purchase_vline02.Top - lngOffset1
19040     .cmbRevenueCodes_Purchase_vline03.Top = .cmbRevenueCodes_Purchase_vline03.Top - lngOffset1
19050     .cmbRevenueCodes_Purchase_vline04.Top = .cmbRevenueCodes_Purchase_vline04.Top - lngOffset1
19060     .cmbRevenueCodes_Purchase.Top = .cmbRevenueCodes_Purchase.Top - lngOffset1
19070     .cmbRevenueCodes_Purchase_display.Top = .cmbRevenueCodes_Purchase_display.Top - lngOffset1
19080     .cmbRevenueCodes_Purchase_lbl.Top = .cmbRevenueCodes_Purchase_lbl.Top - lngOffset1
19090     .cmbRevenueCodes_Purchase_lbl2.Top = .cmbRevenueCodes_Purchase_lbl2.Top - lngOffset1
19100     .cmbTaxCodes_Purchase.Top = .cmbTaxCodes_Purchase.Top - lngOffset1
19110     .cmbTaxCodes_Purchase_display.Top = .cmbTaxCodes_Purchase_display.Top - lngOffset1
19120     .cmbTaxCodes_Purchase_lbl.Top = .cmbTaxCodes_Purchase_lbl.Top - lngOffset1
19130     .cmbTaxCodes_Purchase_lbl2.Top = .cmbTaxCodes_Purchase_lbl2.Top - lngOffset1

19140     .detail_height_line.Top = .detail_height_line.Top - lngOffset2
19150     .Detail.Height = .Detail.Height - lngOffset2

19160     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
19170     lngMonitorNum = 1&: lngTmp02 = 0&
19180     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
19190     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
19200     If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
19210     DoCmd.MoveSize lngFrm_Left, lngTmp02, lngFrm_Width, (lngFrm_Height - lngOffset2)  'lngFrm_Top
19220     If lngMonitorNum > 1& Then
19230       LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
19240     End If

19250     blnHideIncExp = True
19260     DoEvents

19270   End With

EXITP:
19280   Exit Sub

ERRH:
19290   Select Case ERR.Number
        Case Else
19300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19310   End Select
19320   Resume EXITP

End Sub
