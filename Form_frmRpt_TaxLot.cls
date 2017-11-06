VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_TaxLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_TaxLot"

'VGC 07/16/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   cmdExcelOptions_Click()
' #######################################

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Asset Desc:       D {opgOrderBy_optAssetDescription}
' **   Excel:            E {cmdExcel}
' **   Excel Options:    I {cmdExcelOptions}
' **   All:              L {opgAccountNumber_optAll}
' **   Name:             M {opgAccountSource_optName}
' **   Options:          N {ckgAltOptional_box}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Reset:            S {cmdReset}
' **   Asset Type:       T {opgOrderBy_optAssetType}
' **   Specified:        U {opgAccountNumber_optSpecified}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

'QRY: qryTaxLotReport_01                 REF: FormRef('AccountNo')

' ** Array: arr_varChkGrp().
Private lngChkGrps As Long, arr_varChkGrp() As Variant
Private Const CG_ELEMS As Integer = 10  ' ** Array's first-element UBound().
Private Const CG_NAM As Integer = 0
Private Const CG_TYP As Integer = 1
Private Const CG_TOP As Integer = 2
Private Const CG_LFT As Integer = 3
Private Const CG_HGT As Integer = 4
Private Const CG_DEF As Integer = 5
Private Const CG_PAR As Integer = 6
Private Const CG_VAL As Integer = 7
Private Const CG_FAL As Integer = 8
Private Const CG_FSN As Integer = 9
Private Const CG_OPT As Integer = 10

' ** Array: arr_varAcctCtl().
Private lngAcctCtls As Long, arr_varAcctCtl() As Variant
Private Const A_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const A_CNAM As Integer = 0
Private Const A_TYP  As Integer = 1
Private Const A_TOP  As Integer = 2

Private Const MAX_OPTS As Long = 14&

Private blnExcelOptions_Focus As Boolean, blnExcelOptions_MouseDown As Boolean
Private blnReset_Focus As Boolean, blnReset_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngFrm_Diff As Long, lngDetail_Height As Long, lngReset_Top As Long
Private lngSeparation As Long, lngDiff As Long, lngRights_Offset As Long

Private strSortField_All As String, strSortField_Single As String
Private strAcctField_All As String, strAcctField_Single As String
Private strTypeField_All As String, strTypeField_Single As String
Private strDescField_All As String, strDescField_Single As String

Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long ', lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim varTmp00 As Variant, blnTmp01 As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
150           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

              ' ** Variables are fed empty, then populated ByRef.
160           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

170           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

180           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
190           If lngMonitorCnt > 1& Then
200             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
210             lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
220             If lngMonitorNum > 1& Then
230               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
240             End If
250           End If

260           .detail_height_line.Top = 0&  ' ** Just get this out of the way.
270           .cmdReset_lbl.Top = 0&
280           lngSeparation = .ckgAlt_box3.Top - (.opgAccountNumber_box.Top + .opgAccountNumber_box.Height)
290           lngDiff = ((.cmbAccounts_box.Top - (.ckgAlt_box3.Top + .ckgAlt_box3.Height)) - lngSeparation)
300           lngFrm_Diff = lngDiff
310           lngDetail_Height = .Detail.Height
320           lngRights_Offset = (.Rights_Microsoft_lbl2.Top - .Rights_Microsoft_lbl1.Top)
330           lngReset_Top = .cmdReset.Top

340           ChkGroup_Show "Load"  ' ** Procedure: Below.
350           AcctCtls_Move "Load"  ' ** Procedure: Below.

360           blnTmp01 = False
              ' ** tblPreference_User, just 'chkOptionsOpen'.
370           varTmp00 = DCount("*", "qryRpt_TaxLot_02")
380           If IsNull(varTmp00) = True Then
390             blnTmp01 = True  ' ** True mean 'Yes, they have no bananas'.
400           Else
410             If varTmp00 = 0 Then
420               blnTmp01 = True
430             End If
440           End If

450           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

460           strSortField_All = vbNullString: strSortField_Single = vbNullString
470           strAcctField_All = vbNullString: strAcctField_Single = vbNullString
480           strTypeField_All = vbNullString: strTypeField_Single = vbNullString
490           strDescField_All = vbNullString: strDescField_Single = vbNullString
500           blnExcelOptions_Focus = False: blnExcelOptions_MouseDown = False
510           blnReset_Focus = False: blnReset_MouseDown = False
520           gblnMessage = False

      #If NoExcel Then
530           .cmdExcel.Enabled = False
540           .chkOptionsOpen = False
      #Else
550           .cmdExcel.Enabled = True
560           .cmdExcelOptions.Enabled = True
570           .cmdExcelOptions_R_raised_img.Visible = True
580           .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = False
590           .cmdExcelOptions_R_raised_focus_img.Visible = False
600           .cmdExcelOptions_R_raised_focus_dots_img.Visible = False
610           .cmdExcelOptions_R_sunken_focus_dots_img.Visible = False
620           .cmdExcelOptions_R_raised_img_dis.Visible = False
      #End If

630           blnIsOpen = True
640           chkRememberMe_AfterUpdate  ' ** Procedure: Below.
650           opgAccountSource_AfterUpdate  ' ** Procedure: Below.
660           opgAccountNumber_AfterUpdate  ' ** Procedure: Below.
670           If blnTmp01 = True Then
                ' ** If they have no prefs, meaning first time
                ' ** opening, then close the Excel options.
680             .chkOptionsOpen = False
690           End If
700           Select Case .chkOptionsOpen
              Case True
710             ChkGroup_Show "Show"  ' ** Procedure: Below.
720             AcctCtls_Move "Show"  ' ** Procedure: Below.
730           Case False
740             ChkGroup_Show "Hide"  ' ** Procedure: Below.
750             AcctCtls_Move "Hide"  ' ** Procedure: Below.
760           End Select
770           ChkGroup_Default  ' ** Procedure: Below.
780           opgOrderBy_AfterUpdate  ' ** Procedure: Below.
790           chkPageOf_AfterUpdate  ' ** Procedure: Below.
800           cmdReset_Check  ' ** Procedure: Below.

      #If NoExcel Then
810           .cmdExcelOptions.Enabled = False
820           .cmdExcelOptions_R_raised_img_dis.Visible = True
830           .cmdExcelOptions_R_raised_img.Visible = False
840           .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = False
850           .cmdExcelOptions_R_raised_focus_img.Visible = False
860           .cmdExcelOptions_R_raised_focus_dots_img.Visible = False
870           .cmdExcelOptions_R_sunken_focus_dots_img.Visible = False
      #Else
880           Select Case .ckgAlt_box.Visible
              Case True
890             .cmdExcelOptions_L_raised_img.Visible = True
900             .cmdExcelOptions_L_raised_semifocus_dots_img.Visible = False
910             .cmdExcelOptions_L_raised_focus_img.Visible = False
920             .cmdExcelOptions_L_raised_focus_dots_img.Visible = False
930             .cmdExcelOptions_L_sunken_focus_dots_img.Visible = False
940             .cmdExcelOptions_L_raised_img_dis.Visible = False
950             .cmdExcelOptions_R_raised_img.Visible = False
960             .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = False
970             .cmdExcelOptions_R_raised_focus_img.Visible = False
980             .cmdExcelOptions_R_raised_focus_dots_img.Visible = False
990             .cmdExcelOptions_R_sunken_focus_dots_img.Visible = False
1000            .cmdExcelOptions_R_raised_img_dis.Visible = False
1010          Case False
1020            .cmdExcelOptions_R_raised_img.Visible = True
1030            .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = False
1040            .cmdExcelOptions_R_raised_focus_img.Visible = False
1050            .cmdExcelOptions_R_raised_focus_dots_img.Visible = False
1060            .cmdExcelOptions_R_sunken_focus_dots_img.Visible = False
1070            .cmdExcelOptions_R_raised_img_dis.Visible = False
1080            .cmdExcelOptions_L_raised_img.Visible = False
1090            .cmdExcelOptions_L_raised_semifocus_dots_img.Visible = False
1100            .cmdExcelOptions_L_raised_focus_img.Visible = False
1110            .cmdExcelOptions_L_raised_focus_dots_img.Visible = False
1120            .cmdExcelOptions_L_sunken_focus_dots_img.Visible = False
1130            .cmdExcelOptions_L_raised_img_dis.Visible = False
1140          End Select
      #End If

              ' ** EVENT CHECK: chkRememberMe!
              ' ** NONE!

1150          DoEvents
1160          blnIsOpen = False

1170        Else
1180          Cancel = -1
1190        End If
1200      Else
1210        Cancel = -1
1220      End If
1230    End With

1240    If Cancel = -1 Then
1250      Beep
1260      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1270    End If

1280    DoCmd.Hourglass False

EXITP:
1290    Exit Sub

ERRH:
1300    DoCmd.Hourglass False
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Options:          N {ckgAltOptional_box}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyEscape
1470        intRetVal = 0
1480        cmdClose_Click  ' ** Procedure: Below.
1490      End Select
1500    End If

        ' ** Alt keys.
1510    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1520      Select Case intRetVal
          Case vbKeyN
1530        With Me
1540          intRetVal = 0
1550          If .ckgAlt_opt06_ShortName.Visible = True And .ckgAlt_opt06_ShortName.Enabled = True Then
1560            .ckgAlt_opt06_ShortName.SetFocus
1570          Else
1580            Beep
1590          End If
1600        End With
1610      Case vbKeyX
1620        intRetVal = 0
1630        cmdClose_Click  ' ** Procedure: Below.
1640      End Select
1650    End If

        ' ** Ctrl keys.
1660    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1670      Select Case intRetVal
          Case vbKeyG
1680        With Me
1690          intRetVal = 0
1700          If .chkPageOf.Enabled = True Then
1710            .chkPageOf.SetFocus
1720            .chkPageOf = (Not .chkPageOf)
1730            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1740          End If
1750        End With
1760      End Select
1770    End If

        ' ** Ctrl-Shift keys.
1780    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1790      Select Case intRetVal
          Case vbKeyF
1800        With Me
1810          intRetVal = 0
1820          .FocusHolder.SetFocus
1830        End With
1840      End Select
1850    End If

EXITP:
1860    KeyCode = intRetVal
1870    Exit Sub

ERRH:
1880    intRetVal = 0
1890    Select Case ERR.Number
        Case Else
1900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1910    End Select
1920    Resume EXITP

End Sub

Private Sub Form_Timer()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

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
2120      End If
2130      If gblnGoToReport = True Then
2140        If .cmdPrint.Enabled = True Then
2150          .cmdPrint.SetFocus
2160        Else
2170          .cmdClose.SetFocus
2180        End If
2190        gblnGoToReport = False
2200        Beep
2210        DoCmd.Hourglass False
2220        DoEvents
2230      Else
2240        DoCmd.Hourglass False
2250      End If
2260    End With

EXITP:
2270    Exit Sub

ERRH:
2280    DoCmd.Hourglass False
2290    Select Case ERR.Number
        Case Else
2300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2310    End Select
2320    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2410    DoCmd.Hourglass True
2420    DoEvents

2430    With Me
2440      If .chkRememberMe = False Then
2450        .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
2460      End If
2470    End With

2480    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
2490    Do While Reports.Count > 0
2500      DoCmd.Close acReport, Reports(0).Name
2510    Loop

2520    gstrFormQuerySpec = vbNullString
2530    gblnSetFocus = True
2540    gstrReturningForm = THIS_NAME
2550    Select Case strCallingForm
        Case "frmMenu_Report"
2560      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2570    End Select

EXITP:
2580    Exit Sub

ERRH:
2590    Select Case ERR.Number
        Case Else
2600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2610    End Select
2620    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2710    With Me
2720      If .cmdExcelOptions_R_raised_focus_dots_img.Visible = True Or .cmdExcelOptions_R_raised_focus_img.Visible = True Then
2730        Select Case blnExcelOptions_Focus
            Case True
2740          .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = True
2750          .cmdExcelOptions_R_raised_img.Visible = False
2760        Case False
2770          .cmdExcelOptions_R_raised_img.Visible = True
2780          .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = False
2790        End Select
2800        .cmdExcelOptions_R_raised_focus_img.Visible = False
2810        .cmdExcelOptions_R_raised_focus_dots_img.Visible = False
2820        .cmdExcelOptions_R_sunken_focus_dots_img.Visible = False
2830        .cmdExcelOptions_R_raised_img_dis.Visible = False
2840      End If
2850      If .cmdExcelOptions_L_raised_focus_dots_img.Visible = True Or .cmdExcelOptions_L_raised_focus_img.Visible = True Then
2860        Select Case blnExcelOptions_Focus
            Case True
2870          .cmdExcelOptions_L_raised_semifocus_dots_img.Visible = True
2880          .cmdExcelOptions_L_raised_img.Visible = False
2890        Case False
2900          .cmdExcelOptions_L_raised_img.Visible = True
2910          .cmdExcelOptions_L_raised_semifocus_dots_img.Visible = False
2920        End Select
2930        .cmdExcelOptions_L_raised_focus_img.Visible = False
2940        .cmdExcelOptions_L_raised_focus_dots_img.Visible = False
2950        .cmdExcelOptions_L_sunken_focus_dots_img.Visible = False
2960        .cmdExcelOptions_L_raised_img_dis.Visible = False
2970      End If
2980      If .cmdReset_raised_focus_dots_img.Visible = True Or .cmdReset_raised_focus_img.Visible = True Then
2990        Select Case blnReset_Focus
            Case True
3000          .cmdReset_raised_semifocus_dots_img.Visible = True
3010          .cmdReset_raised_img.Visible = False
3020        Case False
3030          .cmdReset_raised_img.Visible = True
3040          .cmdReset_raised_semifocus_dots_img.Visible = False
3050        End Select
3060        .cmdReset_raised_focus_img.Visible = False
3070        .cmdReset_raised_focus_dots_img.Visible = False
3080        .cmdReset_sunken_focus_dots_img.Visible = False
3090        .cmdReset_raised_img_dis.Visible = False
3100      End If
3110    End With

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub cmdClose_Click()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3210    DoCmd.Hourglass True
3220    DoEvents

3230    DoCmd.Close acForm, THIS_NAME

EXITP:
3240    Exit Sub

ERRH:
3250    Select Case ERR.Number
        Case Else
3260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3270    End Select
3280    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
3370        With Me
3380          intRetVal = 0
3390          .opgAccountNumber.SetFocus
3400        End With
3410      End Select
3420    End If

        ' ** Shift keys.
3430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3440      Select Case intRetVal
          Case vbKeyTab
3450        With Me
3460          intRetVal = 0
3470          If .cmdExcel.Enabled = True Then
3480            .cmdExcel.SetFocus
3490          Else
3500            .cmdWord.SetFocus
3510          End If
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

Private Sub opgAccountNumber_AfterUpdate()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_AfterUpdate"

3710    With Me
3720      Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
3730        .opgAccountNumber_optSpecified_lbl.FontBold = True
3740        .opgAccountNumber_optAll_lbl.FontBold = False
3750        .opgAccountNumber_optSpecified_lbl_box.Visible = True
3760        .opgAccountNumber_optAll_lbl_box.Visible = False
3770        .cmbAccounts.Enabled = True
3780        .cmbAccounts.BorderColor = CLR_LTBLU2
3790        .cmbAccounts.BackStyle = acBackStyleNormal
3800        .opgAccountSource.Enabled = True
3810        .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
3820        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
3830        .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
3840        .opgAccountSource_optName_lbl2_dim_hi.Visible = False
3850        .chkRememberMe.Enabled = True
3860        .chkRememberMe_lbl.Visible = True
3870        .chkRememberMe_lbl2_dim.Visible = False
3880        .chkRememberMe_lbl2_dim_hi.Visible = False
3890      Case .opgAccountNumber_optAll.OptionValue
3900        .opgAccountNumber_optSpecified_lbl.FontBold = False
3910        .opgAccountNumber_optAll_lbl.FontBold = True
3920        .opgAccountNumber_optSpecified_lbl_box.Visible = False
3930        .opgAccountNumber_optAll_lbl_box.Visible = True
3940        .cmbAccounts.Enabled = False
3950        .cmbAccounts.BorderColor = WIN_CLR_DISR
3960        .cmbAccounts.BackStyle = acBackStyleTransparent
3970        .opgAccountSource.Enabled = False
3980        .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
3990        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
4000        .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
4010        .opgAccountSource_optName_lbl2_dim_hi.Visible = True
4020        .chkRememberMe.Enabled = False
4030        .chkRememberMe_lbl.Visible = False
4040        .chkRememberMe_lbl2_dim.Visible = True
4050        .chkRememberMe_lbl2_dim_hi.Visible = True
4060      End Select
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

Private Sub opgAccountNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4220    intShiftDown = (Shift And acShiftMask) > 0
4230    intAltDown = (Shift And acAltMask) > 0
4240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4270        With Me
4280          intRetVal = 0
4290          If .cmbAccounts.Enabled = True Then
4300            .cmbAccounts.SetFocus
4310          Else
4320            .cmdPreview.SetFocus
4330          End If
4340        End With
4350      End Select
4360    End If

        ' ** Shift keys.
4370    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4380      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4390        With Me
4400          intRetVal = 0
4410          .cmdClose.SetFocus
4420        End With
4430      End Select
4440    End If

        ' ** Ctrl keys.
4450    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4470        With Me
4480          intRetVal = 0
4490          If .cmdExcel.Enabled = True Then
4500            .cmdExcelOptions.SetFocus
4510          Else
4520            .opgAccountSource.SetFocus
4530          End If
4540        End With
4550      End Select
4560    End If

EXITP:
4570    KeyCode = intRetVal
4580    Exit Sub

ERRH:
4590    intRetVal = 0
4600    Select Case ERR.Number
        Case Else
4610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4620    End Select
4630    Resume EXITP

End Sub

Private Sub opgAccountNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optAll_KeyDown"

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
4790          If .cmbAccounts.Enabled = True Then
4800            .cmbAccounts.SetFocus
4810          Else
4820            .cmdPreview.SetFocus
4830          End If
4840        End With
4850      End Select
4860    End If

        ' ** Shift keys.
4870    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4880      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4890        With Me
4900          intRetVal = 0
4910          .cmdClose.SetFocus
4920        End With
4930      End Select
4940    End If

        ' ** Ctrl keys.
4950    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4970        With Me
4980          intRetVal = 0
4990          If .cmdExcel.Enabled = True Then
5000            .cmdExcelOptions.SetFocus
5010          Else
5020            .opgAccountSource.SetFocus
5030          End If
5040        End With
5050      End Select
5060    End If

EXITP:
5070    KeyCode = intRetVal
5080    Exit Sub

ERRH:
5090    intRetVal = 0
5100    Select Case ERR.Number
        Case Else
5110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5120    End Select
5130    Resume EXITP

End Sub

Private Sub ckgAlt_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_box_MouseMove"

5210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5220    Exit Sub

ERRH:
5230    Select Case ERR.Number
        Case Else
5240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5250    End Select
5260    Resume EXITP

End Sub

Private Sub ckgAlt_box2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_box2_MouseMove"

5310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Private Sub ckgAlt_box3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_box3_MouseMove"

5410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5420    Exit Sub

ERRH:
5430    Select Case ERR.Number
        Case Else
5440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5450    End Select
5460    Resume EXITP

End Sub

Private Sub cmdExcelOptions_GotFocus()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelOptions_GotFocus"

5510    With Me
5520      blnExcelOptions_Focus = True
5530      Select Case .ckgAlt_box.Visible
          Case True
5540        .cmdExcelOptions_L_raised_semifocus_dots_img.Visible = True
5550        .cmdExcelOptions_L_raised_img.Visible = False
5560        .cmdExcelOptions_L_raised_focus_img.Visible = False
5570        .cmdExcelOptions_L_raised_focus_dots_img.Visible = False
5580        .cmdExcelOptions_L_sunken_focus_dots_img.Visible = False
5590        .cmdExcelOptions_L_raised_img_dis.Visible = False
5600      Case False
5610        .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = True
5620        .cmdExcelOptions_R_raised_img.Visible = False
5630        .cmdExcelOptions_R_raised_focus_img.Visible = False
5640        .cmdExcelOptions_R_raised_focus_dots_img.Visible = False
5650        .cmdExcelOptions_R_sunken_focus_dots_img.Visible = False
5660        .cmdExcelOptions_R_raised_img_dis.Visible = False
5670      End Select
5680    End With

EXITP:
5690    Exit Sub

ERRH:
5700    Select Case ERR.Number
        Case Else
5710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5720    End Select
5730    Resume EXITP

End Sub

Private Sub cmdExcelOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelOptions_MouseDown"

5810    With Me
5820      blnExcelOptions_MouseDown = True
5830      Select Case .ckgAlt_box.Visible
          Case True
5840        .cmdExcelOptions_L_sunken_focus_dots_img.Visible = True
5850        .cmdExcelOptions_L_raised_img.Visible = False
5860        .cmdExcelOptions_L_raised_semifocus_dots_img.Visible = False
5870        .cmdExcelOptions_L_raised_focus_img.Visible = False
5880        .cmdExcelOptions_L_raised_focus_dots_img.Visible = False
5890        .cmdExcelOptions_L_raised_img_dis.Visible = False
5900      Case False
5910        .cmdExcelOptions_R_sunken_focus_dots_img.Visible = True
5920        .cmdExcelOptions_R_raised_img.Visible = False
5930        .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = False
5940        .cmdExcelOptions_R_raised_focus_img.Visible = False
5950        .cmdExcelOptions_R_raised_focus_dots_img.Visible = False
5960        .cmdExcelOptions_R_raised_img_dis.Visible = False
5970      End Select
5980    End With

EXITP:
5990    Exit Sub

ERRH:
6000    Select Case ERR.Number
        Case Else
6010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6020    End Select
6030    Resume EXITP

End Sub

Private Sub cmdExcelOptions_Click()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelOptions_Click"

6110    With Me

6120      Select Case .ckgAlt_box.Visible
          Case True
6130        ChkGroup_Show "Hide"  ' ** Procedure: Below.
6140        AcctCtls_Move "Hide"  ' ** Procedure: Below.
6150        .opgOrderBy_optAssetType_lbl2.Visible = False
6160        .opgOrderBy_optAssetDescription_lbl2.Visible = False
6170      Case False
6180        ChkGroup_Show "Show"  ' ** Procedure: Below.
6190        AcctCtls_Move "Show"  ' ** Procedure: Below.
6200        Select Case .opgOrderBy
            Case .opgOrderBy_optAssetDescription.OptionValue
6210          .opgOrderBy_optAssetDescription_lbl2.Visible = True
6220          .ckgAlt_opt07_Description_lbl2.Visible = True
6230          ckgAlt_opt07_Description_AfterUpdate  ' ** Procedure: Below.
6240        Case .opgOrderBy_optAssetType.OptionValue
6250          .opgOrderBy_optAssetType_lbl2.Visible = True
6260          .ckgAlt_opt08_AssetType_lbl2.Visible = True
6270          ckgAlt_opt08_AssetType_AfterUpdate  ' ** Procedure: Below.
6280        End Select
6290      End Select
6300      cmdReset_Check  ' ** Procedure: Below.
6310    End With

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Sub

Private Sub cmdExcelOptions_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelOptions_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6420    intShiftDown = (Shift And acShiftMask) > 0
6430    intAltDown = (Shift And acAltMask) > 0
6440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6470        With Me
6480          intRetVal = 0
6490          Select Case .ckgAlt_box.Visible
              Case True
6500            .ckgAlt_opt06_ShortName.SetFocus
6510          Case False
6520            .opgAccountSource.SetFocus
6530          End Select
6540        End With
6550      End Select
6560    End If

        ' ** Shift keys.
6570    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6580      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6590        With Me
6600          intRetVal = 0
6610          .opgAccountNumber.SetFocus
6620        End With
6630      End Select
6640    End If

EXITP:
6650    KeyCode = intRetVal
6660    Exit Sub

ERRH:
6670    intRetVal = 0
6680    Select Case ERR.Number
        Case Else
6690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6700    End Select
6710    Resume EXITP

End Sub

Private Sub cmdExcelOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelOptions_MouseMove"

6810    With Me
6820      If blnExcelOptions_MouseDown = False Then
6830        Select Case .ckgAlt_box.Visible
            Case True
6840          Select Case blnExcelOptions_Focus
              Case True
6850            .cmdExcelOptions_L_raised_focus_dots_img.Visible = True
6860            .cmdExcelOptions_L_raised_focus_img.Visible = False
6870          Case False
6880            .cmdExcelOptions_L_raised_focus_img.Visible = True
6890            .cmdExcelOptions_L_raised_focus_dots_img.Visible = False
6900          End Select
6910          .cmdExcelOptions_L_raised_img.Visible = False
6920          .cmdExcelOptions_L_raised_semifocus_dots_img.Visible = False
6930          .cmdExcelOptions_L_sunken_focus_dots_img.Visible = False
6940          .cmdExcelOptions_L_raised_img_dis.Visible = False
6950        Case False
6960          Select Case blnExcelOptions_Focus
              Case True
6970            .cmdExcelOptions_R_raised_focus_dots_img.Visible = True
6980            .cmdExcelOptions_R_raised_focus_img.Visible = False
6990          Case False
7000            .cmdExcelOptions_R_raised_focus_img.Visible = True
7010            .cmdExcelOptions_R_raised_focus_dots_img.Visible = False
7020          End Select
7030          .cmdExcelOptions_R_raised_img.Visible = False
7040          .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = False
7050          .cmdExcelOptions_R_sunken_focus_dots_img.Visible = False
7060          .cmdExcelOptions_R_raised_img_dis.Visible = False
7070        End Select
7080      End If
7090    End With

EXITP:
7100    Exit Sub

ERRH:
7110    Select Case ERR.Number
        Case Else
7120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7130    End Select
7140    Resume EXITP

End Sub

Private Sub cmdExcelOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelOptions_MouseUp"

7210    With Me
7220      Select Case .ckgAlt_box.Visible
          Case True
7230        .cmdExcelOptions_L_raised_focus_dots_img.Visible = True
7240        .cmdExcelOptions_L_raised_img.Visible = False
7250        .cmdExcelOptions_L_raised_semifocus_dots_img.Visible = False
7260        .cmdExcelOptions_L_raised_focus_img.Visible = False
7270        .cmdExcelOptions_L_sunken_focus_dots_img.Visible = False
7280        .cmdExcelOptions_L_raised_img_dis.Visible = False
7290      Case False
7300        .cmdExcelOptions_R_raised_focus_dots_img.Visible = True
7310        .cmdExcelOptions_R_raised_img.Visible = False
7320        .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = False
7330        .cmdExcelOptions_R_raised_focus_img.Visible = False
7340        .cmdExcelOptions_R_sunken_focus_dots_img.Visible = False
7350        .cmdExcelOptions_R_raised_img_dis.Visible = False
7360      End Select
7370      blnExcelOptions_MouseDown = False
7380    End With

EXITP:
7390    Exit Sub

ERRH:
7400    Select Case ERR.Number
        Case Else
7410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7420    End Select
7430    Resume EXITP

End Sub

Private Sub cmdExcelOptions_LostFocus()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelOptions_LostFocus"

7510    With Me
7520      Select Case .ckgAlt_box.Visible
          Case True
7530        .cmdExcelOptions_L_raised_img.Visible = True
7540        .cmdExcelOptions_L_raised_semifocus_dots_img.Visible = False
7550        .cmdExcelOptions_L_raised_focus_img.Visible = False
7560        .cmdExcelOptions_L_raised_focus_dots_img.Visible = False
7570        .cmdExcelOptions_L_sunken_focus_dots_img.Visible = False
7580        .cmdExcelOptions_L_raised_img_dis.Visible = False
7590      Case False
7600        .cmdExcelOptions_R_raised_img.Visible = True
7610        .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = False
7620        .cmdExcelOptions_R_raised_focus_img.Visible = False
7630        .cmdExcelOptions_R_raised_focus_dots_img.Visible = False
7640        .cmdExcelOptions_R_sunken_focus_dots_img.Visible = False
7650        .cmdExcelOptions_R_raised_img_dis.Visible = False
7660      End Select
7670      blnExcelOptions_Focus = False
7680    End With

EXITP:
7690    Exit Sub

ERRH:
7700    Select Case ERR.Number
        Case Else
7710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7720    End Select
7730    Resume EXITP

End Sub

Private Sub ckgAltOptional_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAltOptional_box_MouseMove"

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

Private Sub ckgAlt_opt01_AccountNo_AfterUpdate()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt01_AccountNo_AfterUpdate"

7910    ChkGroup_Select THIS_PROC  ' ** Procedure: Below.

EXITP:
7920    Exit Sub

ERRH:
7930    DoCmd.Hourglass False
7940    Select Case ERR.Number
        Case Else
7950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7960    End Select
7970    Resume EXITP

End Sub

Private Sub ckgAlt_opt02_CUSIP_AfterUpdate()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt02_CUSIP_AfterUpdate"

8010    ChkGroup_Select THIS_PROC  ' ** Procedure: Below.

EXITP:
8020    Exit Sub

ERRH:
8030    DoCmd.Hourglass False
8040    Select Case ERR.Number
        Case Else
8050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8060    End Select
8070    Resume EXITP

End Sub

Private Sub ckgAlt_opt03_AssetDate_AfterUpdate()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt03_AssetDate_AfterUpdate"

8110    ChkGroup_Select THIS_PROC  ' ** Procedure: Below.

EXITP:
8120    Exit Sub

ERRH:
8130    DoCmd.Hourglass False
8140    Select Case ERR.Number
        Case Else
8150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8160    End Select
8170    Resume EXITP

End Sub

Private Sub ckgAlt_opt04_ShareFace_AfterUpdate()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt04_ShareFace_AfterUpdate"

8210    ChkGroup_Select THIS_PROC  ' ** Procedure: Below.

EXITP:
8220    Exit Sub

ERRH:
8230    DoCmd.Hourglass False
8240    Select Case ERR.Number
        Case Else
8250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8260    End Select
8270    Resume EXITP

End Sub

Private Sub ckgAlt_opt05_Cost_AfterUpdate()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt05_Cost_AfterUpdate"

8310    ChkGroup_Select THIS_PROC  ' ** Procedure: Below.

EXITP:
8320    Exit Sub

ERRH:
8330    DoCmd.Hourglass False
8340    Select Case ERR.Number
        Case Else
8350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8360    End Select
8370    Resume EXITP

End Sub

Private Sub ckgAlt_opt06_ShortName_AfterUpdate()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt06_ShortName_AfterUpdate"

8410    ChkGroup_Select THIS_PROC  ' ** Procedure: Below.
8420    cmdReset_Check  ' ** Procedure: Below.

EXITP:
8430    Exit Sub

ERRH:
8440    DoCmd.Hourglass False
8450    Select Case ERR.Number
        Case Else
8460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8470    End Select
8480    Resume EXITP

End Sub

Private Sub ckgAlt_opt06_ShortName_KeyDown(KeyCode As Integer, Shift As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt06_ShortName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8520    intShiftDown = (Shift And acShiftMask) > 0
8530    intAltDown = (Shift And acAltMask) > 0
8540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8570        With Me
8580          intRetVal = 0
8590          .ckgAlt_opt07_Description.SetFocus
8600        End With
8610      End Select
8620    End If

        ' ** Shift keys.
8630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8650        With Me
8660          intRetVal = 0
8670          .cmdExcelOptions.SetFocus
8680        End With
8690      End Select
8700    End If

EXITP:
8710    KeyCode = intRetVal
8720    Exit Sub

ERRH:
8730    intRetVal = 0
8740    Select Case ERR.Number
        Case Else
8750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8760    End Select
8770    Resume EXITP

End Sub

Private Sub ckgAlt_opt07_Description_AfterUpdate()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt07_Description_AfterUpdate"

8810    With Me
8820      ChkGroup_Select THIS_PROC  ' ** Procedure: Below.
8830      Select Case .opgOrderBy
          Case .opgOrderBy_optAssetDescription.OptionValue
            ' ** Sort by Asset Description.
8840        If (.ckgAlt_opt07_Description = False And .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectFlat) Or blnIsOpen = True Then
              ' ** They clicked while it was flat, so we go to full showing.
8850          .ckgAlt_opt07_Description = True
8860          .ckgAlt_opt07_Description.BorderStyle = acBorderStyleSolid
8870          .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectSunken
8880          .ckgAlt_opt07_Description.Top = .ckgAlt_opt12_HoldingPeriod.Top
8890          .ckgAlt_opt07_Description.Left = .ckgAlt_opt06_ShortName.Left
8900          If .ckgAltOptional_box.BackColor <> MY_CLR_VLTBGE Then  ' ** In case they're all checked.
8910            .ckgAlt_opt07_Description_lbl.ForeColor = CLR_VDKGRY
8920          End If
8930          .ckgAlt_opt07_Description_lbl.FontBold = True
8940          .ckgAlt_opt07_Description_box.Visible = True
8950        ElseIf .ckgAlt_opt07_Description = False And .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectSunken Then
              ' ** They clicked while it was sunken, so we go to flat dimmed.
8960          .ckgAlt_opt07_Description = True
8970          .ckgAlt_opt07_Description.BorderStyle = acBorderStyleTransparent
8980          .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectFlat
8990          .ckgAlt_opt07_Description.Top = (.ckgAlt_opt12_HoldingPeriod.Top + lngTpp)
9000          .ckgAlt_opt07_Description.Left = (.ckgAlt_opt06_ShortName.Left + lngTpp)
9010          .ckgAlt_opt07_Description_lbl.ForeColor = CLR_DKGRY3  ' ** This is lighter than all checked.
9020          .ckgAlt_opt07_Description_lbl.FontBold = False
9030          .ckgAlt_opt07_Description_box.Visible = False
9040        ElseIf .ckgAlt_opt07_Description = True Then
              ' ** Might be on open.
9050          If .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectFlat Then
                ' ** This is second time opened after a close.
9060            .ckgAlt_opt07_Description.BorderStyle = acBorderStyleSolid
9070            .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectSunken
9080            .ckgAlt_opt07_Description.Top = .ckgAlt_opt12_HoldingPeriod.Top
9090            .ckgAlt_opt07_Description.Left = .ckgAlt_opt06_ShortName.Left
9100            If .ckgAltOptional_box.BackColor <> MY_CLR_VLTBGE Then  ' ** In case they're all checked.
9110              .ckgAlt_opt07_Description_lbl.ForeColor = CLR_VDKGRY
9120            End If
9130            .ckgAlt_opt07_Description_lbl.FontBold = True
9140            .ckgAlt_opt07_Description_box.Visible = True
9150          ElseIf .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectSunken Then
                ' ** This is first time opened after a close.
                ' ** It ends up alternating between here and just above!
                ' ** We want it to open the same way it closed!
9160            .ckgAlt_opt07_Description.BorderStyle = acBorderStyleTransparent
9170            .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectFlat
9180            .ckgAlt_opt07_Description.Top = (.ckgAlt_opt12_HoldingPeriod.Top + lngTpp)
9190            .ckgAlt_opt07_Description.Left = (.ckgAlt_opt06_ShortName.Left + lngTpp)
9200            .ckgAlt_opt07_Description_lbl.ForeColor = CLR_DKGRY3  ' ** This is lighter than all checked.
9210            .ckgAlt_opt07_Description_lbl.FontBold = False
9220            .ckgAlt_opt07_Description_box.Visible = False
9230          End If
9240        End If
9250      Case .opgOrderBy_optAssetType.OptionValue
            ' ** Sort by Asset Type.
9260        If .ckgAlt_opt07_Description.SpecialEffect <> acSpecialEffectSunken Then
9270          .ckgAlt_opt07_Description.BorderStyle = acBorderStyleSolid
9280          .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectSunken
9290          .ckgAlt_opt07_Description.Top = .ckgAlt_opt12_HoldingPeriod.Top
9300          .ckgAlt_opt07_Description.Left = .ckgAlt_opt06_ShortName.Left
9310        Else
              ' ** Let it stand.
9320        End If
9330        If .ckgAltOptional_box.BackColor <> MY_CLR_VLTBGE Then
9340          .ckgAlt_opt07_Description_lbl.ForeColor = CLR_VDKGRY
9350        End If
9360      End Select  ' ** chkOrderBy_AssetType.
9370      cmdReset_Check  ' ** Procedure: Below.
9380    End With

EXITP:
9390    Exit Sub

ERRH:
9400    DoCmd.Hourglass False
9410    Select Case ERR.Number
        Case Else
9420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9430    End Select
9440    Resume EXITP

End Sub

Private Sub ckgAlt_opt08_AssetType_AfterUpdate()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt08_AssetType_AfterUpdate"

9510    ChkGroup_Select THIS_PROC  ' ** Procedure: Below.
9520    cmdReset_Check  ' ** Procedure: Below.

EXITP:
9530    Exit Sub

ERRH:
9540    DoCmd.Hourglass False
9550    Select Case ERR.Number
        Case Else
9560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9570    End Select
9580    Resume EXITP

End Sub

Private Sub ckgAlt_opt09_PricePerUnit_AfterUpdate()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt09_PricePerUnit_AfterUpdate"

9610    ChkGroup_Select THIS_PROC  ' ** Procedure: Below.
9620    cmdReset_Check  ' ** Procedure: Below.

EXITP:
9630    Exit Sub

ERRH:
9640    DoCmd.Hourglass False
9650    Select Case ERR.Number
        Case Else
9660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9670    End Select
9680    Resume EXITP

End Sub

Private Sub ckgAlt_opt10_CurrentMarketValue_AfterUpdate()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt10_CurrentMarketValue_AfterUpdate"

9710    ChkGroup_Select THIS_PROC  ' ** Procedure: Below.
9720    cmdReset_Check  ' ** Procedure: Below.

EXITP:
9730    Exit Sub

ERRH:
9740    DoCmd.Hourglass False
9750    Select Case ERR.Number
        Case Else
9760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9770    End Select
9780    Resume EXITP

End Sub

Private Sub ckgAlt_opt11_MarketValue_AfterUpdate()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt11_MarketValue_AfterUpdate"

9810    ChkGroup_Select THIS_PROC  ' ** Procedure: Below.
9820    cmdReset_Check  ' ** Procedure: Below.

EXITP:
9830    Exit Sub

ERRH:
9840    DoCmd.Hourglass False
9850    Select Case ERR.Number
        Case Else
9860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9870    End Select
9880    Resume EXITP

End Sub

Private Sub ckgAlt_opt12_HoldingPeriod_AfterUpdate()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt12_HoldingPeriod_AfterUpdate"

9910    ChkGroup_Select THIS_PROC  ' ** Procedure: Below.
9920    cmdReset_Check  ' ** Procedure: Below.

EXITP:
9930    Exit Sub

ERRH:
9940    DoCmd.Hourglass False
9950    Select Case ERR.Number
        Case Else
9960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9970    End Select
9980    Resume EXITP

End Sub

Private Sub ckgAlt_opt13_LocationName_AfterUpdate()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt13_LocationName_AfterUpdate"

10010   ChkGroup_Select THIS_PROC  ' ** Procedure: Below.
10020   cmdReset_Check  ' ** Procedure: Below.

EXITP:
10030   Exit Sub

ERRH:
10040   DoCmd.Hourglass False
10050   Select Case ERR.Number
        Case Else
10060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10070   End Select
10080   Resume EXITP

End Sub

Private Sub ckgAlt_opt14_AssetDate_Time_AfterUpdate()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt14_AssetDate_Time_AfterUpdate"

10110   ChkGroup_Select THIS_PROC  ' ** Procedure: Below.
10120   cmdReset_Check  ' ** Procedure: Below.

EXITP:
10130   Exit Sub

ERRH:
10140   DoCmd.Hourglass False
10150   Select Case ERR.Number
        Case Else
10160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10170   End Select
10180   Resume EXITP

End Sub

Private Sub ckgAlt_opt14_AssetDate_Time_KeyDown(KeyCode As Integer, Shift As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt14_AssetDate_Time_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10220   intShiftDown = (Shift And acShiftMask) > 0
10230   intAltDown = (Shift And acAltMask) > 0
10240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10270       With Me
10280         intRetVal = 0
10290         If .cmdReset.Enabled = True Then
10300           .cmdReset.SetFocus
10310         Else
10320           .opgAccountSource.SetFocus
10330         End If
10340       End With
10350     End Select
10360   End If

        ' ** Shift keys.
10370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10390       With Me
10400         intRetVal = 0
10410         .ckgAlt_opt13_LocationName.SetFocus
10420       End With
10430     End Select
10440   End If

EXITP:
10450   KeyCode = intRetVal
10460   Exit Sub

ERRH:
10470   intRetVal = 0
10480   Select Case ERR.Number
        Case Else
10490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10500   End Select
10510   Resume EXITP

End Sub

Private Sub ckgAlt_opt14_AssetDate_Time_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgAlt_opt14_AssetDate_Time_MouseMove"

10610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10620   Exit Sub

ERRH:
10630   Select Case ERR.Number
        Case Else
10640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10650   End Select
10660   Resume EXITP

End Sub

Private Sub cmdReset_GotFocus()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_GotFocus"

10710   With Me
10720     blnReset_Focus = True
10730     .cmdReset_raised_semifocus_dots_img.Visible = True
10740     .cmdReset_raised_img.Visible = False
10750     .cmdReset_raised_focus_img.Visible = False
10760     .cmdReset_raised_focus_dots_img.Visible = False
10770     .cmdReset_sunken_focus_dots_img.Visible = False
10780     .cmdReset_raised_img_dis.Visible = False
10790   End With

EXITP:
10800   Exit Sub

ERRH:
10810   Select Case ERR.Number
        Case Else
10820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10830   End Select
10840   Resume EXITP

End Sub

Private Sub cmdReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseDown"

10910   With Me
10920     blnReset_MouseDown = True
10930     .cmdReset_sunken_focus_dots_img.Visible = True
10940     .cmdReset_raised_img.Visible = False
10950     .cmdReset_raised_semifocus_dots_img.Visible = False
10960     .cmdReset_raised_focus_img.Visible = False
10970     .cmdReset_raised_focus_dots_img.Visible = False
10980     .cmdReset_raised_img_dis.Visible = False
10990   End With

EXITP:
11000   Exit Sub

ERRH:
11010   Select Case ERR.Number
        Case Else
11020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11030   End Select
11040   Resume EXITP

End Sub

Private Sub cmdReset_Click()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_Click"

11110   With Me
11120     .FocusHolder.SetFocus
11130     DoEvents
11140     .opgOrderBy = .opgOrderBy_optAssetDescription.OptionValue
11150     opgOrderBy_AfterUpdate  ' ** Procedure: Below.
11160     DoEvents
11170     .ckgAlt_opt06_ShortName = False
11180     ckgAlt_opt06_ShortName_AfterUpdate  ' ** Procedure: Below.
11190     DoEvents
11200     .ckgAlt_opt07_Description = True
          ' ** This is so it doesn flip.
11210     .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectFlat
11220     ckgAlt_opt07_Description_AfterUpdate  ' ** Procedure: Below.
11230     DoEvents
11240     .ckgAlt_opt08_AssetType = False
11250     ckgAlt_opt08_AssetType_AfterUpdate  ' ** Procedure: Below.
11260     DoEvents
11270     .ckgAlt_opt09_PricePerUnit = False
11280     ckgAlt_opt09_PricePerUnit_AfterUpdate  ' ** Procedure: Below.
11290     DoEvents
11300     .ckgAlt_opt10_CurrentMarketValue = False
11310     ckgAlt_opt10_CurrentMarketValue_AfterUpdate  ' ** Procedure: Below.
11320     DoEvents
11330     .ckgAlt_opt11_MarketValue = False
11340     ckgAlt_opt11_MarketValue_AfterUpdate  ' ** Procedure: Below.
11350     DoEvents
11360     .ckgAlt_opt12_HoldingPeriod = False
11370     ckgAlt_opt12_HoldingPeriod_AfterUpdate  ' ** Procedure: Below.
11380     DoEvents
11390     .ckgAlt_opt13_LocationName = False
11400     ckgAlt_opt13_LocationName_AfterUpdate  ' ** Procedure: Below.
11410     DoEvents
11420     .ckgAlt_opt14_AssetDate_Time = False
11430     ckgAlt_opt14_AssetDate_Time_AfterUpdate  ' ** Procedure: Below.
11440     DoEvents
11450     .FocusHolder.SetFocus
11460     DoEvents
11470     .cmdReset.Enabled = False
11480     .cmdReset_raised_img_dis.Visible = True
11490     .cmdReset_raised_img.Visible = False
11500     .cmdReset_raised_semifocus_dots_img.Visible = False
11510     .cmdReset_raised_focus_img.Visible = False
11520     .cmdReset_raised_focus_dots_img.Visible = False
11530     .cmdReset_sunken_focus_dots_img.Visible = False
11540   End With

EXITP:
11550   Exit Sub

ERRH:
11560   Select Case ERR.Number
        Case Else
11570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11580   End Select
11590   Resume EXITP

End Sub

Private Sub cmdReset_Check()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_Check"

        Dim blnFound As Boolean

11610   With Me
11620     blnFound = False  ' ** All default.
11630     If .opgOrderBy <> .opgOrderBy_optAssetDescription.OptionValue Then
11640       blnFound = True
11650     ElseIf .ckgAlt_opt06_ShortName <> False Then
11660       blnFound = True
11670     ElseIf .ckgAlt_opt07_Description <> True Then
11680       blnFound = True
11690     ElseIf .ckgAlt_opt08_AssetType <> False Then
11700       blnFound = True
11710     ElseIf .ckgAlt_opt09_PricePerUnit <> False Then
11720       blnFound = True
11730     ElseIf .ckgAlt_opt10_CurrentMarketValue <> False Then
11740       blnFound = True
11750     ElseIf .ckgAlt_opt11_MarketValue <> False Then
11760       blnFound = True
11770     ElseIf .ckgAlt_opt12_HoldingPeriod <> False Then
11780       blnFound = True
11790     ElseIf .ckgAlt_opt13_LocationName <> False Then
11800       blnFound = True
11810     ElseIf .ckgAlt_opt14_AssetDate_Time <> False Then
11820       blnFound = True
11830     End If
11840     Select Case .ckgAltOptional_box.Visible
          Case True
11850       .ckgAlt_box4.Visible = False
11860       If .cmdReset.Enabled <> blnFound Then
11870         Select Case blnFound
              Case True
                ' ** Enable cmdReset.
11880           .cmdReset.Enabled = True
11890           .cmdReset_raised_img.Visible = True
11900           .cmdReset_raised_img_dis.Visible = False
11910         Case False
                ' ** Disable cmdReset.
11920           .cmdReset.Enabled = False
11930           .cmdReset_raised_img_dis.Visible = True
11940           .cmdReset_raised_img.Visible = False
11950         End Select
11960         .cmdReset_raised_semifocus_dots_img.Visible = False
11970         .cmdReset_raised_focus_img.Visible = False
11980         .cmdReset_raised_focus_dots_img.Visible = False
11990         .cmdReset_sunken_focus_dots_img.Visible = False
12000       End If
12010     Case False
12020       If blnFound = True Then
12030         .ckgAlt_box4.Visible = True
12040       Else
12050         .ckgAlt_box4.Visible = False
12060       End If
12070     End Select
12080   End With

EXITP:
12090   Exit Sub

ERRH:
12100   Select Case ERR.Number
        Case Else
12110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12120   End Select
12130   Resume EXITP

End Sub

Private Sub cmdReset_KeyDown(KeyCode As Integer, Shift As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12220   intShiftDown = (Shift And acShiftMask) > 0
12230   intAltDown = (Shift And acAltMask) > 0
12240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12270       With Me
12280         intRetVal = 0
12290         .opgAccountSource.SetFocus
12300       End With
12310     End Select
12320   End If

        ' ** Shift keys.
12330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12350       With Me
12360         intRetVal = 0
12370         .ckgAlt_opt14_AssetDate_Time.SetFocus
12380       End With
12390     End Select
12400   End If

EXITP:
12410   KeyCode = intRetVal
12420   Exit Sub

ERRH:
12430   intRetVal = 0
12440   Select Case ERR.Number
        Case Else
12450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12460   End Select
12470   Resume EXITP

End Sub

Private Sub cmdReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseMove"

12510   With Me
12520     If blnReset_MouseDown = False Then
12530       Select Case blnReset_Focus
            Case True
12540         .cmdReset_raised_focus_dots_img.Visible = True
12550         .cmdReset_raised_focus_img.Visible = False
12560       Case False
12570         .cmdReset_raised_focus_img.Visible = True
12580         .cmdReset_raised_focus_dots_img.Visible = False
12590       End Select
12600       .cmdReset_raised_img.Visible = False
12610       .cmdReset_raised_semifocus_dots_img.Visible = False
12620       .cmdReset_sunken_focus_dots_img.Visible = False
12630       .cmdReset_raised_img_dis.Visible = False
12640     End If
12650   End With

EXITP:
12660   Exit Sub

ERRH:
12670   Select Case ERR.Number
        Case Else
12680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12690   End Select
12700   Resume EXITP

End Sub

Private Sub cmdReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseUp"

12810   With Me
12820     .cmdReset_raised_focus_dots_img.Visible = True
12830     .cmdReset_raised_img.Visible = False
12840     .cmdReset_raised_semifocus_dots_img.Visible = False
12850     .cmdReset_raised_focus_img.Visible = False
12860     .cmdReset_sunken_focus_dots_img.Visible = False
12870     .cmdReset_raised_img_dis.Visible = False
12880     blnReset_MouseDown = False
12890   End With

EXITP:
12900   Exit Sub

ERRH:
12910   Select Case ERR.Number
        Case Else
12920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12930   End Select
12940   Resume EXITP

End Sub

Private Sub cmdReset_LostFocus()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_LostFocus"

13010   With Me
13020     .cmdReset_raised_img.Visible = True
13030     .cmdReset_raised_semifocus_dots_img.Visible = False
13040     .cmdReset_raised_focus_img.Visible = False
13050     .cmdReset_raised_focus_dots_img.Visible = False
13060     .cmdReset_sunken_focus_dots_img.Visible = False
13070     .cmdReset_raised_img_dis.Visible = False
13080     blnReset_Focus = False
13090   End With

EXITP:
13100   Exit Sub

ERRH:
13110   Select Case ERR.Number
        Case Else
13120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13130   End Select
13140   Resume EXITP

End Sub

Private Sub Rights_Microsoft_lbl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "Rights_Microsoft_lbl1_MouseMove"

13210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13220   Exit Sub

ERRH:
13230   Select Case ERR.Number
        Case Else
13240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13250   End Select
13260   Resume EXITP

End Sub

Private Sub Rights_Microsoft_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "Rights_Microsoft_lbl2_MouseMove"

13310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13320   Exit Sub

ERRH:
13330   Select Case ERR.Number
        Case Else
13340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13350   End Select
13360   Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

13410   strAccountNo = vbNullString

13420   With Me
13430     If IsNull(.cmbAccounts) = False Then
13440       If Len(.cmbAccounts.Column(0)) > 0 Then
13450         strAccountNo = .cmbAccounts.Column(0)
13460       End If
13470     End If
13480     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
13490       .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
13500       .opgAccountSource_optNumber_lbl2.FontBold = True
13510       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
13520       .opgAccountSource_optName_lbl2.FontBold = False
13530       .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
13540     Case .opgAccountSource_optName.OptionValue
13550       .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
13560       .opgAccountSource_optNumber_lbl2.FontBold = False
13570       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
13580       .opgAccountSource_optName_lbl2.FontBold = True
13590       .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
13600     End Select
13610     DoEvents
13620     If strAccountNo <> vbNullString Then
13630       .cmbAccounts = strAccountNo
13640     End If
13650   End With

EXITP:
13660   Exit Sub

ERRH:
13670   Select Case ERR.Number
        Case Else
13680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13690   End Select
13700   Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

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
13890         .chkRememberMe.SetFocus
13900       End With
13910     End Select
13920   End If

        ' ** Shift keys.
13930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13950       With Me
13960         intRetVal = 0
13970         Select Case .ckgAlt_box.Visible
              Case True
13980           If .cmdReset.Enabled = True Then
13990             .cmdReset.SetFocus
14000           Else
14010             .ckgAlt_opt14_AssetDate_Time.SetFocus
14020           End If
14030         Case False
14040           If .cmdExcel.Enabled = True Then
14050             .cmdExcelOptions.SetFocus
14060           Else
14070             .opgAccountNumber.SetFocus
14080           End If
14090         End Select
14100       End With
14110     End Select
14120   End If

EXITP:
14130   KeyCode = intRetVal
14140   Exit Sub

ERRH:
14150   intRetVal = 0
14160   Select Case ERR.Number
        Case Else
14170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14180   End Select
14190   Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

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
14290         .chkRememberMe.SetFocus
14300       End With
14310     End Select
14320   End If

        ' ** Shift keys.
14330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14350       With Me
14360         intRetVal = 0
14370         Select Case .ckgAlt_box.Visible
              Case True
14380           If .cmdReset.Enabled = True Then
14390             .cmdReset.SetFocus
14400           Else
14410             .ckgAlt_opt14_AssetDate_Time.SetFocus
14420           End If
14430         Case False
14440           If .cmdExcel.Enabled = True Then
14450             .cmdExcelOptions.SetFocus
14460           Else
14470             .opgAccountNumber.SetFocus
14480           End If
14490         End Select
14500       End With
14510     End Select
14520   End If

EXITP:
14530   KeyCode = intRetVal
14540   Exit Sub

ERRH:
14550   intRetVal = 0
14560   Select Case ERR.Number
        Case Else
14570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14580   End Select
14590   Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

14610   With Me
14620     Select Case .chkRememberMe
          Case True
14630       .chkRememberMe_lbl.FontBold = True
14640       .chkRememberMe_lbl2_dim.FontBold = True
14650       .chkRememberMe_lbl2_dim_hi.FontBold = True
14660     Case False
14670       .chkRememberMe_lbl.FontBold = False
14680       .chkRememberMe_lbl2_dim.FontBold = False
14690       .chkRememberMe_lbl2_dim_hi.FontBold = False
14700     End Select
14710   End With

EXITP:
14720   Exit Sub

ERRH:
14730   Select Case ERR.Number
        Case Else
14740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14750   End Select
14760   Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14820   intShiftDown = (Shift And acShiftMask) > 0
14830   intAltDown = (Shift And acAltMask) > 0
14840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14870       With Me
14880         intRetVal = 0
14890         .cmbAccounts.SetFocus
14900       End With
14910     End Select
14920   End If

        ' ** Shift keys.
14930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14950       With Me
14960         intRetVal = 0
14970         .opgAccountSource.SetFocus
14980       End With
14990     End Select
15000   End If

EXITP:
15010   KeyCode = intRetVal
15020   Exit Sub

ERRH:
15030   intRetVal = 0
15040   Select Case ERR.Number
        Case Else
15050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15060   End Select
15070   Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

15110   With Me
15120     If .cmbAccounts.text = vbNullString Then
15130       ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
15140       .cmbAccounts = .cmbAccounts.ItemData(0)
15150     End If
15160   End With

EXITP:
15170   Exit Sub

ERRH:
15180   Select Case ERR.Number
        Case Else
15190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15200   End Select
15210   Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15320   intShiftDown = (Shift And acShiftMask) > 0
15330   intAltDown = (Shift And acAltMask) > 0
15340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15370       With Me
15380         intRetVal = 0
15390         .cmdPreview.SetFocus
15400       End With
15410     End Select
15420   End If

        ' ** Shift keys.
15430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15450       With Me
15460         intRetVal = 0
15470         .opgAccountNumber.SetFocus
15480       End With
15490     End Select
15500   End If

        ' ** Ctrl keys.
15510   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
15520     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15530       With Me
15540         intRetVal = 0
15550         .opgOrderBy.SetFocus
15560       End With
15570     End Select
15580   End If

        ' ** Ctrl-Shift keys.
15590   If intCtrlDown And (Not intAltDown) And intShiftDown Then
15600     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15610       With Me
15620         intRetVal = 0
15630         .chkRememberMe.SetFocus
15640       End With
15650     End Select
15660   End If

EXITP:
15670   KeyCode = intRetVal
15680   Exit Sub

ERRH:
15690   intRetVal = 0
15700   Select Case ERR.Number
        Case Else
15710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15720   End Select
15730   Resume EXITP

End Sub

Private Sub opgOrderBy_AfterUpdate()

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_AfterUpdate"

15810   With Me
15820     Select Case .opgOrderBy
          Case .opgOrderBy_optAssetDescription.OptionValue
            ' ** Sort by Asset Description.
15830       .opgOrderBy_optAssetType_lbl.FontBold = False
15840       .opgOrderBy_optAssetDescription_lbl.FontBold = True
15850       Select Case .chkOptionsOpen
            Case True
15860         If .ckgAlt_opt07_Description = False Then
15870           .ckgAlt_opt07_Description = True
15880           ckgAlt_opt07_Description_AfterUpdate  ' ** Procedure: Below.
15890           DoEvents
15900         End If
15910         .opgOrderBy_optAssetDescription_lbl2.Visible = True
15920         .ckgAlt_opt07_Description_lbl2.Visible = True
15930         .opgOrderBy_optAssetType_lbl2.Visible = False
15940         .ckgAlt_opt08_AssetType_lbl2.Visible = False
15950       Case False
15960         .opgOrderBy_optAssetType_lbl2.Visible = False
15970         .opgOrderBy_optAssetType_lbl2.Visible = False
15980         .opgOrderBy_optAssetDescription_lbl2.Visible = False
15990       End Select
            ' ** Make sure Type is normal.
16000       .ckgAlt_opt08_AssetType.Locked = False
16010     Case .opgOrderBy_optAssetType.OptionValue
            ' ** Sort by Asset Type.
16020       .opgOrderBy_optAssetType_lbl.FontBold = True
16030       .opgOrderBy_optAssetDescription_lbl.FontBold = False
16040       Select Case .chkOptionsOpen
            Case True
16050         If .ckgAlt_opt08_AssetType = False Then
16060           .ckgAlt_opt08_AssetType = True
16070           ckgAlt_opt08_AssetType_AfterUpdate  ' ** Procedure: Below.
16080           DoEvents
16090           .ckgAlt_opt08_AssetType.Locked = True
16100         ElseIf .ckgAlt_opt08_AssetType.Locked = False Then
16110           .ckgAlt_opt08_AssetType.Locked = True
16120         End If
16130         .opgOrderBy_optAssetType_lbl2.Visible = True
16140         .ckgAlt_opt08_AssetType_lbl2.Visible = True
16150         .opgOrderBy_optAssetDescription_lbl2.Visible = False
16160         .ckgAlt_opt07_Description_lbl2.Visible = False
16170       Case False
16180         .ckgAlt_opt08_AssetType.Locked = False
16190         .opgOrderBy_optAssetType_lbl2.Visible = False
16200         .opgOrderBy_optAssetDescription_lbl2.Visible = False
16210       End Select
            ' ** Make sure Description is normal.
16220       .ckgAlt_opt07_Description.BorderStyle = acBorderStyleSolid
16230       .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectSunken
16240       .ckgAlt_opt07_Description.Top = .ckgAlt_opt12_HoldingPeriod.Top
16250       .ckgAlt_opt07_Description.Left = .ckgAlt_opt06_ShortName.Left
16260       If .ckgAlt_opt07_Description_lbl.ForeColor = CLR_DKGRY3 Then
16270         .ckgAlt_opt07_Description_lbl.ForeColor = CLR_VDKGRY
16280       End If
16290     End Select  ' ** opgOrderBy.
16300     cmdReset_Check  ' ** Procedure: Below.
16310   End With

EXITP:
16320   Exit Sub

ERRH:
16330   Select Case ERR.Number
        Case Else
16340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16350   End Select
16360   Resume EXITP

End Sub

Private Sub opgOrderBy_optAssetDescription_KeyDown(KeyCode As Integer, Shift As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_optAssetDescription_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16420   intShiftDown = (Shift And acShiftMask) > 0
16430   intAltDown = (Shift And acAltMask) > 0
16440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16470       With Me
16480         intRetVal = 0
16490         .chkPageOf.SetFocus
16500       End With
16510     End Select
16520   End If

        ' ** Shift keys.
16530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16550       With Me
16560         intRetVal = 0
16570         If .cmbAccounts.Enabled = True Then
16580           .cmbAccounts.SetFocus
16590         Else
16600           If .cmdExcel.Enabled = True Then
16610             .cmdExcelOptions.SetFocus
16620           Else
16630             .opgAccountNumber.SetFocus
16640           End If
16650         End If
16660       End With
16670     End Select
16680   End If

EXITP:
16690   KeyCode = intRetVal
16700   Exit Sub

ERRH:
16710   intRetVal = 0
16720   Select Case ERR.Number
        Case Else
16730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16740   End Select
16750   Resume EXITP

End Sub

Private Sub opgOrderBy_optAssetType_KeyDown(KeyCode As Integer, Shift As Integer)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_optAssetType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16820   intShiftDown = (Shift And acShiftMask) > 0
16830   intAltDown = (Shift And acAltMask) > 0
16840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16870       With Me
16880         intRetVal = 0
16890         .chkPageOf.SetFocus
16900       End With
16910     End Select
16920   End If

        ' ** Shift keys.
16930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16950       With Me
16960         intRetVal = 0
16970         If .cmbAccounts.Enabled = True Then
16980           .cmbAccounts.SetFocus
16990         Else
17000           If .cmdExcel.Enabled = True Then
17010             .cmdExcelOptions.SetFocus
17020           Else
17030             .opgAccountNumber.SetFocus
17040           End If
17050         End If
17060       End With
17070     End Select
17080   End If

EXITP:
17090   KeyCode = intRetVal
17100   Exit Sub

ERRH:
17110   intRetVal = 0
17120   Select Case ERR.Number
        Case Else
17130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17140   End Select
17150   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

17210   With Me
17220     Select Case .chkPageOf
          Case True
17230       .chkPageOf_lbl.FontBold = True
17240     Case False
17250       .chkPageOf_lbl.FontBold = False
17260     End Select
17270   End With

EXITP:
17280   Exit Sub

ERRH:
17290   DoCmd.Hourglass False
17300   Select Case ERR.Number
        Case Else
17310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17320   End Select
17330   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17420   intShiftDown = (Shift And acShiftMask) > 0
17430   intAltDown = (Shift And acAltMask) > 0
17440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17470       With Me
17480         intRetVal = 0
17490         .cmdPreview.SetFocus
17500       End With
17510     End Select
17520   End If

        ' ** Shift keys.
17530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17550       With Me
17560         intRetVal = 0
17570         .opgOrderBy.SetFocus
17580       End With
17590     End Select
17600   End If

EXITP:
17610   KeyCode = intRetVal
17620   Exit Sub

ERRH:
17630   intRetVal = 0
17640   Select Case ERR.Number
        Case Else
17650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17660   End Select
17670   Resume EXITP

End Sub

Private Sub cmdPreview_Click()

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

17710   With Me
17720     If DoReport = True Then  ' ** Function: Below.

17730       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
17740         gstrAccountNo = .cmbAccounts.Column(0)
17750         Select Case .opgOrderBy
              Case .opgOrderBy_optAssetDescription.OptionValue
17760           strDocName = "rptTaxLot_One_SortDesc"
17770         Case .opgOrderBy_optAssetType.OptionValue
17780           strDocName = "rptTaxLot_One_SortType"
17790         End Select
17800       Case .opgAccountNumber_optAll.OptionValue
17810         gstrAccountNo = "All"
17820         Select Case .opgOrderBy
              Case .opgOrderBy_optAssetDescription.OptionValue
17830           strDocName = "rptTaxLot_all_SortDesc"
17840         Case .opgOrderBy_optAssetType.OptionValue
17850           strDocName = "rptTaxLot_all_SortType"
17860         End Select
17870       End Select

17880       DoCmd.OpenReport strDocName, acViewPreview
17890       DoCmd.Maximize
17900       DoCmd.RunCommand acCmdFitToWindow

17910     End If
17920   End With

EXITP:
17930   Exit Sub

ERRH:
17940   Select Case ERR.Number
        Case Else
17950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17960   End Select
17970   Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18020   intShiftDown = (Shift And acShiftMask) > 0
18030   intAltDown = (Shift And acAltMask) > 0
18040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18060     Select Case intRetVal
          Case vbKeyTab
18070       With Me
18080         intRetVal = 0
18090         .cmdPrint.SetFocus
18100       End With
18110     End Select
18120   End If

        ' ** Shift keys.
18130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18140     Select Case intRetVal
          Case vbKeyTab
18150       With Me
18160         intRetVal = 0
18170         If .cmbAccounts.Enabled = True Then
18180           .cmbAccounts.SetFocus
18190         Else
18200           .opgAccountNumber.SetFocus
18210         End If
18220       End With
18230     End Select
18240   End If

        ' ** Ctrl-Shift keys.
18250   If intCtrlDown And (Not intAltDown) And intShiftDown Then
18260     Select Case intRetVal
          Case vbKeyTab
18270       With Me
18280         intRetVal = 0
18290         .chkPageOf.SetFocus
18300       End With
18310     End Select
18320   End If

EXITP:
18330   KeyCode = intRetVal
18340   Exit Sub

ERRH:
18350   intRetVal = 0
18360   Select Case ERR.Number
        Case Else
18370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18380   End Select
18390   Resume EXITP

End Sub

Private Sub cmdPrint_Click()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

18410   With Me
18420     If DoReport = True Then  ' ** Function: Below.

18430       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
18440         gstrAccountNo = .cmbAccounts.Column(0)
18450         Select Case .opgOrderBy
              Case .opgOrderBy_optAssetDescription.OptionValue
18460           strDocName = "rptTaxLot_One_SortDesc"
18470         Case .opgOrderBy_optAssetType.OptionValue
18480           strDocName = "rptTaxLot_One_SortType"
18490         End Select
18500       Case .opgAccountNumber_optAll.OptionValue
18510         gstrAccountNo = "All"
18520         Select Case .opgOrderBy
              Case .opgOrderBy_optAssetDescription.OptionValue
18530           strDocName = "rptTaxLot_All_SortDesc"
18540         Case .opgOrderBy_optAssetType.OptionValue
18550           strDocName = "rptTaxLot_All_SortType"
18560         End Select
18570       End Select

18580       If gblnDev_Debug = True Then
18590         DoCmd.OpenReport strDocName, acViewPreview
18600         DoCmd.Maximize
18610         DoCmd.RunCommand acCmdFitToWindow
18620       Else
              '##GTR_Ref: rptTaxLot_All_SortDesc
              '##GTR_Ref: rptTaxLot_All_SortType
              '##GTR_Ref: rptTaxLot_One_SortDesc
              '##GTR_Ref: rptTaxLot_One_SortType
18630         DoCmd.OpenReport strDocName, acViewNormal
18640       End If

18650     End If
18660   End With

EXITP:
18670   Exit Sub

ERRH:
18680   Select Case ERR.Number
        Case Else
18690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18700   End Select
18710   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18820   intShiftDown = (Shift And acShiftMask) > 0
18830   intAltDown = (Shift And acAltMask) > 0
18840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18860     Select Case intRetVal
          Case vbKeyTab
18870       With Me
18880         intRetVal = 0
18890         .cmdWord.SetFocus
18900       End With
18910     End Select
18920   End If

        ' ** Shift keys.
18930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18940     Select Case intRetVal
          Case vbKeyTab
18950       With Me
18960         intRetVal = 0
18970         .cmdPreview.SetFocus
18980       End With
18990     End Select
19000   End If

EXITP:
19010   KeyCode = intRetVal
19020   Exit Sub

ERRH:
19030   intRetVal = 0
19040   Select Case ERR.Number
        Case Else
19050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19060   End Select
19070   Resume EXITP

End Sub

Private Sub cmdWord_Click()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String
        Dim strRptCap As String, strRptPath As String, strRptPathFile As String

19110   With Me
19120     DoCmd.Hourglass True
19130     DoEvents
19140     If DoReport = True Then  ' ** Function: Below.

19150       strRpt = vbNullString
19160       strRptCap = vbNullString: strRptPathFile = vbNullString

19170       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
19180         gstrAccountNo = .cmbAccounts.Column(0)
19190         strRptCap = "rptTaxLots_" & gstrAccountNo & "_" & Format(Date, "yymmdd")
19200         Select Case .opgOrderBy
              Case .opgOrderBy_optAssetDescription.OptionValue
19210           strRpt = "rptTaxLot_One_SortDesc"
19220         Case .opgOrderBy_optAssetType.OptionValue
19230           strRpt = "rptTaxLot_One_SortType"
19240         End Select
19250       Case .opgAccountNumber_optAll.OptionValue
19260         gstrAccountNo = "All"
19270         strRptCap = "rptTaxLots_All_" & Format(Date, "yymmdd")
19280         Select Case .opgOrderBy
              Case .opgOrderBy_optAssetDescription.OptionValue
19290           strRpt = "rptTaxLot_All_SortDesc"
19300         Case .opgOrderBy_optAssetType.OptionValue
19310           strRpt = "rptTaxLot_All_SortType"
19320         End Select
19330       End Select

19340       strRptCap = StringReplace(strRptCap, "/", "_")  ' ** Module Function: modStringFuncs.

19350       If IsNull(.UserReportPath) = True Then
19360         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
19370       Else
19380         strRptPath = .UserReportPath
19390       End If

19400       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

19410       If strRptPathFile <> vbNullString Then

19420         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.

19430         If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
19440           Kill strRptPathFile
19450         End If

19460         DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True

19470       End If

19480     End If
19490     DoCmd.Hourglass False
19500   End With

EXITP:
19510   Exit Sub

ERRH:
19520   DoCmd.Hourglass False
19530   Select Case ERR.Number
        Case Else
19540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19550   End Select
19560   Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19620   intShiftDown = (Shift And acShiftMask) > 0
19630   intAltDown = (Shift And acAltMask) > 0
19640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19660     Select Case intRetVal
          Case vbKeyTab
19670       With Me
19680         intRetVal = 0
19690         If .cmdExcel.Enabled = True Then
19700           .cmdExcel.SetFocus
19710         Else
19720           .cmdClose.SetFocus
19730         End If
19740       End With
19750     End Select
19760   End If

        ' ** Shift keys.
19770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19780     Select Case intRetVal
          Case vbKeyTab
19790       With Me
19800         intRetVal = 0
19810         .cmdPrint.SetFocus
19820       End With
19830     End Select
19840   End If

EXITP:
19850   KeyCode = intRetVal
19860   Exit Sub

ERRH:
19870   intRetVal = 0
19880   Select Case ERR.Number
        Case Else
19890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19900   End Select
19910   Resume EXITP

End Sub

Private Sub cmdExcel_Click()

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strQry As String, strSQL As String, strSource As String
        Dim strRptCap As String, strRptPath As String, strRptPathFile As String
        Dim blnIsTwo As Boolean, blnTypeSort As Boolean, blnIsAll As Boolean, blnDescSortOnly As Boolean
        Dim intPos01 As Integer, intPos02 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim lngX As Long, lngY As Long, intZ As Integer

20010   With Me

20020     DoCmd.Hourglass True
20030     DoEvents

20040     If DoReport = True Then  ' ** Function: Below.

20050       strQry = vbNullString: blnIsAll = False
20060       strRptCap = vbNullString: strRptPathFile = vbNullString

20070       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
20080         gstrAccountNo = .cmbAccounts.Column(0)
20090         strRptCap = "rptTaxLots_" & gstrAccountNo & "_" & Format(Date, "yymmdd")
20100         Select Case .opgOrderBy
              Case .opgOrderBy_optAssetDescription.OptionValue
20110           strQry = "qryTaxLotReport_07a_01"
20120         Case .opgOrderBy_optAssetType.OptionValue
20130           strQry = "qryTaxLotReport_07b_01"
20140         End Select
20150       Case .opgAccountNumber_optAll.OptionValue
20160         gstrAccountNo = "All"
20170         blnIsAll = True
20180         strRptCap = "rptTaxLots_All_" & Format(Date, "yymmdd")
20190         Select Case .opgOrderBy
              Case .opgOrderBy_optAssetDescription.OptionValue
20200           strQry = "qryTaxLotReport_07c_01"
20210         Case .opgOrderBy_optAssetType.OptionValue
20220           strQry = "qryTaxLotReport_07d_01"
20230         End Select
20240       End Select

20250       Select Case .chkOptionsOpen
            Case True

              ' ** The 2's all have these fields:
              ' **   x opt01 : accountno   : Default
              ' **   x opt02 : CUSIP       : Default
              ' **   x opt03 : assetdate   : Default
              ' **   x opt04 : shareface   : Default
              ' **   x opt05 : cost        : Default
              ' **   opt09 : Unit Price      : priceperunit
              ' **   opt10 : Market Price    : marketvaluecurrent
              ' **   opt11 : Market Value    : shareface * marketvaluecurrent
              ' **   opt12 : Holding Period  :
              ' **   opt13 : Location        : Location_Name

              ' ** If their choice happens to match above, use the existing 2's.
20260         blnIsTwo = True
20270         For lngX = 0& To (lngChkGrps - 1&)
20280           If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
20290             Select Case arr_varChkGrp(CG_OPT, lngX)
                  Case 1&, 2&, 3&, 4&, 5&, 9&, 10&, 11&, 12&, 13&
20300               If arr_varChkGrp(CG_VAL, lngX) = False Then
20310                 blnIsTwo = False
20320                 Exit For
20330               End If
20340             Case Else
20350               If arr_varChkGrp(CG_VAL, lngX) = True Then
20360                 blnIsTwo = False
20370                 Exit For
20380               End If
20390             End Select
20400           End If
20410         Next

20420         Select Case blnIsTwo
              Case True
20430           strQry = Left(strQry, (Len(strQry) - 1)) & "2"
20440         Case False

                ' ** qryTaxLotReport_07a_03a, single, no sort, all fields, by specified FormRef('AccountNo').
                ' ** qryTaxLotReport_07a_03b
                ' ** SELECT qryTaxLotReport_07a_03a.[Account Num], qryTaxLotReport_07a_03a.[Short Name],
                ' **   qryTaxLotReport_07a_03a.Asset, qryTaxLotReport_07a_03a.CUSIP,
                ' **   qryTaxLotReport_07a_03a.Type, qryTaxLotReport_07a_03a.[Trade Date],
                ' **   qryTaxLotReport_07a_03a.[Trade DateTime], qryTaxLotReport_07a_03a.[Share/Face],
                ' **   qryTaxLotReport_07a_03a.Cost, qryTaxLotReport_07a_03a.[Unit Price],
                ' **   qryTaxLotReport_07a_03a.[Market Price], qryTaxLotReport_07a_03a.[Market Value],
                ' **   qryTaxLotReport_07a_03a.[Holding Period], qryTaxLotReport_07a_03a.Location, qryTaxLotReport_07a_03a.sort1
                ' ** FROM qryTaxLotReport_07a_03a;

                ' ** qryTaxLotReport_07c_03a, all, no sort, all fields.
                ' ** qryTaxLotReport_07c_03b
                ' ** SELECT qryTaxLotReport_07c_03a.[Account Num], qryTaxLotReport_07c_03a.[Short Name],
                ' **   qryTaxLotReport_07c_03a.Asset, qryTaxLotReport_07c_03a.CUSIP,
                ' **   qryTaxLotReport_07c_03a.Type, qryTaxLotReport_07c_03a.[Trade Date],
                ' **   qryTaxLotReport_07c_03a.[Trade DateTime], qryTaxLotReport_07c_03a.[Share/Face],
                ' **   qryTaxLotReport_07c_03a.Cost, qryTaxLotReport_07c_03a.[Unit Price],
                ' **   qryTaxLotReport_07c_03a.[Market Value], qryTaxLotReport_07c_03a.[Market Price],
                ' **   qryTaxLotReport_07c_03a.[Holding Period], qryTaxLotReport_07c_03a.Location, qryTaxLotReport_07c_03a.sort1
                ' ** FROM qryTaxLotReport_07c_03a;

                ' ** Remember, these have got both [Trade Date], and [Trade DateTime].

20450           Select Case blnIsAll
                Case True
20460             strSortField_All = ChkGroup_GetFields(blnIsAll)  ' ** Function: Below.
20470           Case False
20480             strSortField_Single = ChkGroup_GetFields(blnIsAll)  ' ** Function: Below.
20490           End Select

                'For lngX = 0& To (lngChkGrps - 1&)
                '  If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
                '    Debug.Print "'" & arr_varChkGrp(CG_FSN, lngX)
                '  End If
                'Next

20500           Select Case .opgOrderBy
                Case .opgOrderBy_optAssetDescription.OptionValue
20510             blnTypeSort = False
20520           Case .opgOrderBy_optAssetType.OptionValue
20530             blnTypeSort = True
20540           End Select

20550           Select Case blnTypeSort
                Case True
                  ' ** If they've chosen sorting by Asset Type, and it's not checked, then check it.
                  ' ** ckgAlt_opt08_AssetType.
20560             For lngX = 0& To (lngChkGrps - 1&)
20570               If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
20580                 If Right(arr_varChkGrp(CG_NAM, lngX), 10) = "_AssetType" Then
20590                   If arr_varChkGrp(CG_VAL, lngX) = False Then
20600                     .Controls(arr_varChkGrp(CG_NAM, lngX)) = True
20610                     ckgAlt_opt08_AssetType_AfterUpdate  ' ** Procedure: Above.
20620                     arr_varChkGrp(CG_VAL, lngX) = True
20630                     DoEvents
20640                   End If  ' ** CG_VAL.
20650                   Exit For
20660                 End If  ' ** CG_NAM.
20670               End If  ' ** acCheckBox.
20680             Next  ' ** lngX.
                  ' ** Since Asset Type still uses Asset Description in the sort, we'll need that, too.
                  ' ** ckgAlt_opt07_Description.
20690             For lngX = 0& To (lngChkGrps - 1&)
20700               If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
20710                 If Right(arr_varChkGrp(CG_NAM, lngX), 12) = "_Description" Then
20720                   If arr_varChkGrp(CG_VAL, lngX) = False Then
20730                     .Controls(arr_varChkGrp(CG_NAM, lngX)) = True
20740                     ckgAlt_opt07_Description_AfterUpdate  ' ** Procedure: Above.
20750                     .ckgAlt_opt07_Description.BorderStyle = acBorderStyleTransparent
20760                     .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectFlat
20770                     .ckgAlt_opt07_Description.Top = (.ckgAlt_opt12_HoldingPeriod.Top + lngTpp)
20780                     .ckgAlt_opt07_Description.Left = (.ckgAlt_opt06_ShortName.Left + lngTpp)
20790                     .ckgAlt_opt07_Description_lbl.ForeColor = CLR_DKGRY3
20800                     .ckgAlt_opt07_Description_box.Visible = False
20810                     arr_varChkGrp(CG_VAL, lngX) = True
20820                     DoEvents
20830                   End If  ' ** CG_VAL.
20840                   Exit For
20850                 End If  ' ** CG_NAM.
20860               End If  ' ** acCheckBox.
20870             Next  ' ** lngX.

20880           Case False
                  ' ** If they've chosen sorting by Asset Description, and it's not checked,
                  ' ** check it, but do something to indicate it won't be showing.
                  ' ** ckgAlt_opt07_Description.
20890             For lngX = 0& To (lngChkGrps - 1&)
20900               If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
20910                 If Right(arr_varChkGrp(CG_NAM, lngX), 12) = "_Description" Then
20920                   If arr_varChkGrp(CG_VAL, lngX) = False Then
20930                     .Controls(arr_varChkGrp(CG_NAM, lngX)) = True
20940                     ckgAlt_opt07_Description_AfterUpdate  ' ** Procedure: Above.
20950                     .ckgAlt_opt07_Description.BorderStyle = acBorderStyleTransparent
20960                     .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectFlat
20970                     .ckgAlt_opt07_Description.Top = (.ckgAlt_opt12_HoldingPeriod.Top + lngTpp)
20980                     .ckgAlt_opt07_Description.Left = (.ckgAlt_opt06_ShortName.Left + lngTpp)
20990                     .ckgAlt_opt07_Description_lbl.ForeColor = CLR_DKGRY3
21000                     .ckgAlt_opt07_Description_box.Visible = False
21010                     arr_varChkGrp(CG_VAL, lngX) = True
21020                     DoEvents
21030                   End If  ' ** CG_VAL.
21040                   Exit For
21050                 End If  ' ** CG_NAM.
21060               End If  ' ** acCheckBox.
21070             Next  ' ** lngX.
21080           End Select  ' ** blnTypeSort.

21090           Select Case .ckgAlt_opt07_Description
                Case True
21100             If .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectFlat Then
21110               blnDescSortOnly = True
21120             Else
21130               blnDescSortOnly = False
21140             End If
21150           Case False
21160             blnDescSortOnly = True
21170           End Select

                ' ** Now assemble the new SQL.
21180           strSQL = "SELECT ": strSource = vbNullString
21190           For lngX = 1& To MAX_OPTS
21200             For lngY = 0& To (lngChkGrps - 1&)
21210               If arr_varChkGrp(CG_TYP, lngY) = acCheckBox Then
21220                 If arr_varChkGrp(CG_OPT, lngY) = lngX Then  ' ** Do these in order.
21230                   If arr_varChkGrp(CG_VAL, lngY) = True Then
21240                     Select Case blnIsAll
                          Case True
21250                       strSQL = strSQL & arr_varChkGrp(CG_FAL, lngY) & ", "
21260                       If strSource = vbNullString Then  ' ** Just pick it off the first one.
21270                         strSource = Left(arr_varChkGrp(CG_FAL, lngY), (InStr(arr_varChkGrp(CG_FAL, lngY), ".") - 1))
21280                       End If
21290                     Case False
21300                       strSQL = strSQL & arr_varChkGrp(CG_FSN, lngY) & ", "
21310                       If strSource = vbNullString Then  ' ** Just pick it off the first one.
21320                         strSource = Left(arr_varChkGrp(CG_FSN, lngY), (InStr(arr_varChkGrp(CG_FSN, lngY), ".") - 1))
21330                       End If
21340                     End Select
21350                     Exit For
21360                   End If
21370                 End If
21380               End If
21390             Next  ' ** lngY.
21400           Next  ' ** lngX.

                ' ** And add the final sort field.
21410           If Right(strSQL, 2) = ", " Then strSQL = Left(strSQL, (Len(strSQL) - 2))
21420           Select Case blnIsAll
                Case True
21430             strSQL = strSQL & ", " & strSortField_All
21440           Case False
21450             strSQL = strSQL & ", " & strSortField_Single
21460           End Select
21470           strSQL = strSQL & vbCrLf & "FROM " & strSource & ";"

21480           Set dbs = CurrentDb
21490           With dbs

21500             Select Case blnIsAll
                  Case True
                    ' ** qryTaxLotReport_07c_03a, all, assembled by code.
21510               Set qdf = .QueryDefs("qryTaxLotReport_07c_03c")
21520             Case False
                    ' ** qryTaxLotReport_07a_03a, single, assembled by code.
21530               Set qdf = .QueryDefs("qryTaxLotReport_07a_03c")
21540             End Select
21550             qdf.SQL = strSQL
21560             Set qdf = Nothing
                  ' ** These now have the fields, but no sort, and column widths haven't been checked.

                  ' ** Variations needed:
                  ' **   Single, Asset Description
                  ' **   Single, Asset Type
                  ' **   All, Asset Description
                  ' **   All, Asset Type
                  ' ** Sort fields saved:
                  ' **   strSortField_All
                  ' **   strSortField_Single
                  ' **   strAcctField_All
                  ' **   strAcctField_Single
                  ' **   strTypeField_All
                  ' **   strTypeField_Single
                  ' **   strDescField_All
                  ' **   strDescField_Single

                  ' ** Assemble the ORDER BY clause.
21570             strTmp02 = "ORDER BY "
21580             strTmp01 = vbNullString: strTmp03 = vbNullString
21590             intPos01 = 0: intPos02 = 0
21600             Select Case blnIsAll
                  Case True

21610               If strAcctField_All = vbNullString Then
21620                 For lngX = 0& To (lngChkGrps - 1&)
21630                   If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
21640                     If InStr(arr_varChkGrp(CG_FAL, lngX), ".[Account Num") > 0 Then
21650                       strAcctField_All = arr_varChkGrp(CG_FAL, lngX)
21660                       Exit For
21670                     End If
21680                   End If
21690                 Next
21700               End If  ' ** strAcctField_All.
21710               If strTypeField_All = vbNullString Then
21720                 For lngX = 0& To (lngChkGrps - 1&)
21730                   If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
21740                     If InStr(arr_varChkGrp(CG_FAL, lngX), ".Type") > 0 Then
21750                       strTypeField_All = arr_varChkGrp(CG_FAL, lngX)
21760                       Exit For
21770                     End If
21780                   End If
21790                 Next
21800               End If  ' ** strTypeField_All.
21810               If strDescField_All = vbNullString Then
21820                 For lngX = 0& To (lngChkGrps - 1&)
21830                   If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
21840                     If InStr(arr_varChkGrp(CG_FAL, lngX), ".Asset") > 0 Then
21850                       strDescField_All = arr_varChkGrp(CG_FAL, lngX)
21860                       Exit For
21870                     End If
21880                   End If
21890                 Next
21900               End If  ' ** strDescField_All.

                    ' ** qryTaxLotReport_07c_03a, all, assembled by code.
21910               Set qdf = .QueryDefs("qryTaxLotReport_07c_03c")
21920               strSQL = qdf.SQL
21930               Set qdf = Nothing
21940               strTmp01 = strSQL
21950               If Right(strTmp01, 1) = ";" Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
21960               strTmp02 = strTmp02 & strSortField_All & ", "
21970               strTmp02 = strTmp02 & strAcctField_All & ", "

                    ' ** Remove sort1 from select.
21980               intPos01 = InStr(strTmp01, ".sort1")
21990               If intPos01 > 0 Then
22000                 intPos02 = intPos01 + Len(".sort1")  ' ** 1 after sort1.
22010                 If Mid(strTmp01, intPos02, 1) = "]" Then intPos01 = intPos01 + 1
22020                 strTmp03 = Mid(strTmp01, intPos02)  ' ** Everything after sort1.
22030                 For intZ = intPos01 To 1 Step -1
22040                   If Mid(strTmp01, intZ, 1) = "," Then
22050                     strTmp03 = Left(strTmp01, (intZ - 1)) & strTmp03
22060                     Exit For
22070                   End If
22080                 Next
22090                 strTmp01 = strTmp03
22100               End If  ' ** sort1.

                    ' ** Remove Asset from select if only for sort.
22110               If blnDescSortOnly = True Then
22120                 intPos02 = InStr(strTmp01, "FROM ")
22130                 intPos01 = InStr(strTmp01, ".Asset")
22140                 If intPos01 > 0 Then
22150                   If intPos01 < intPos02 Then
22160                     intPos02 = intPos01 + Len(".Asset")
22170                     If Mid(strTmp01, intPos02, 1) = "]" Then intPos02 = intPos02 + 1
22180                     strTmp03 = Mid(strTmp01, intPos02)  ' ** Everything after Asset.
22190                     For intZ = intPos01 To 1 Step -1
22200                       If Mid(strTmp01, intZ, 1) = "," Then
22210                         strTmp03 = Left(strTmp01, (intZ - 1)) & strTmp03
22220                         Exit For
22230                       End If
22240                     Next
22250                     strTmp01 = strTmp03
22260                   End If
22270                 End If
22280               End If  ' ** Asset.

22290               Select Case blnTypeSort
                    Case True
                      ' ** ORDER BY qryTaxLotReport_06a.sort1, qryTaxLotReport_06a.[Account Num], qryTaxLotReport_06a.assettype, qryTaxLotReport_06a.Asset
22300                 strTmp02 = strTmp02 & strTypeField_All & ", "
22310                 strTmp02 = strTmp02 & strDescField_All
22320                 strTmp02 = strTmp02 & ";"
22330               Case False
                      ' ** ORDER BY qryTaxLotReport_06b.sort1, qryTaxLotReport_06b.[Account Num], qryTaxLotReport_06b.Asset
22340                 strTmp02 = strTmp02 & strDescField_All
22350                 strTmp02 = strTmp02 & ";"
22360               End Select
22370               strTmp01 = strTmp01 & vbCrLf & strTmp02  ' ** Completed new SQL.

22380             Case False

22390               If strAcctField_Single = vbNullString Then
22400                 For lngX = 0& To (lngChkGrps - 1&)
22410                   If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
22420                     If InStr(arr_varChkGrp(CG_FSN, lngX), ".[Account Num") > 0 Then
22430                       strAcctField_Single = arr_varChkGrp(CG_FSN, lngX)
22440                       Exit For
22450                     End If
22460                   End If
22470                 Next
22480               End If  ' ** strAcctField_Single.
22490               If strTypeField_Single = vbNullString Then
22500                 For lngX = 0& To (lngChkGrps - 1&)
22510                   If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
22520                     If InStr(arr_varChkGrp(CG_FSN, lngX), ".Type") > 0 Then
22530                       strTypeField_Single = arr_varChkGrp(CG_FSN, lngX)
22540                       Exit For
22550                     End If
22560                   End If
22570                 Next
22580               End If  ' ** strTypeField_Single.
22590               If strDescField_Single = vbNullString Then
22600                 For lngX = 0& To (lngChkGrps - 1&)
22610                   If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
22620                     If InStr(arr_varChkGrp(CG_FSN, lngX), ".Asset") > 0 Then
22630                       strDescField_Single = arr_varChkGrp(CG_FSN, lngX)
22640                       Exit For
22650                     End If
22660                   End If
22670                 Next
22680               End If  ' ** strDescField_Single.

                    ' ** qryTaxLotReport_07a_03a, single, assembled by code.
22690               Set qdf = .QueryDefs("qryTaxLotReport_07a_03c")
22700               strSQL = qdf.SQL
22710               Set qdf = Nothing
22720               strTmp01 = strSQL
22730               If Right(strTmp01, 1) = ";" Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
22740               strTmp02 = strTmp02 & strSortField_Single & ", "
22750               strTmp02 = strTmp02 & strAcctField_Single & ", "

                    ' ** Remove sort1 from select.
22760               intPos01 = InStr(strTmp01, ".sort1")
22770               If intPos01 > 0 Then
22780                 intPos02 = intPos01 + Len(".sort1")  ' ** 1 after sort1.
22790                 If Mid(strTmp01, intPos02, 1) = "]" Then intPos01 = intPos01 + 1
22800                 strTmp03 = Mid(strTmp01, intPos02)  ' ** Everything after sort1.
22810                 For intZ = intPos01 To 1 Step -1
22820                   If Mid(strTmp01, intZ, 1) = "," Then
22830                     strTmp03 = Left(strTmp01, (intZ - 1)) & strTmp03
22840                     Exit For
22850                   End If
22860                 Next
22870                 strTmp01 = strTmp03
22880               End If  ' ** sort1.

                    ' ** Remove Asset from select if only for sort.
22890               If blnDescSortOnly = True Then
22900                 intPos02 = InStr(strTmp01, "FROM ")
22910                 intPos01 = InStr(strTmp01, ".Asset")
22920                 If intPos01 > 0 Then
22930                   If intPos01 < intPos02 Then
22940                     intPos02 = intPos01 + Len(".Asset")
22950                     If Mid(strTmp01, intPos02, 1) = "]" Then intPos02 = intPos02 + 1
22960                     strTmp03 = Mid(strTmp01, intPos02)  ' ** Everything after Asset.
22970                     For intZ = intPos01 To 1 Step -1
22980                       If Mid(strTmp01, intZ, 1) = "," Then
22990                         strTmp03 = Left(strTmp01, (intZ - 1)) & strTmp03
23000                         Exit For
23010                       End If
23020                     Next
23030                     strTmp01 = strTmp03
23040                   End If
23050                 End If
23060               End If  ' ** Asset.

23070               Select Case blnTypeSort
                    Case True
                      ' ** ORDER BY qryTaxLotReport_06b.sort1, qryTaxLotReport_06b.[Account Num], qryTaxLotReport_06b.assettype, qryTaxLotReport_06b.Asset
23080                 strTmp02 = strTmp02 & strTypeField_Single & ", "
23090                 strTmp02 = strTmp02 & strDescField_Single
23100                 strTmp02 = strTmp02 & ";"
23110               Case False
                      ' ** ORDER BY qryTaxLotReport_06a.sort1, qryTaxLotReport_06a.[Account Num], qryTaxLotReport_06a.Asset
23120                 strTmp02 = strTmp02 & strDescField_Single
23130                 strTmp02 = strTmp02 & ";"
23140               End Select
23150               strTmp01 = strTmp01 & vbCrLf & strTmp02  ' ** Completed new SQL.

23160             End Select  ' ** blnIsAll.

23170             Select Case blnIsAll
                  Case True
23180               Select Case blnTypeSort
                    Case True
                      ' ** qryTaxLotReport_07c_03a, all, sorted by AssetType, assembled by code.
23190                 strQry = "qryTaxLotReport_07d_04"
23200               Case False
                      ' ** qryTaxLotReport_07c_03a, all, sorted by Asset_Description, assembled by code.
23210                 strQry = "qryTaxLotReport_07c_04"
23220               End Select
23230             Case False
23240               Select Case blnTypeSort
                    Case True
                      ' ** qryTaxLotReport_07a_03a, single, sorted by AssetType, assembled by code.
23250                 strQry = "qryTaxLotReport_07b_04"
23260               Case False
                      ' ** qryTaxLotReport_07a_03a, single, sorted by Asset_Description, assembled by code.
23270                 strQry = "qryTaxLotReport_07a_04"
23280               End Select
23290             End Select
23300             Set qdf = .QueryDefs(strQry)
23310             qdf.SQL = strTmp01
23320             Set qdf = Nothing

23330             .Close
23340           End With  ' ** dbs.
23350           Set dbs = Nothing

23360         End Select  ' ** blnIsTwo.

23370       Case False
              ' ** Queries stand.
23380       End Select  ' ** chkOptionsOpen.

23390       strRptCap = StringReplace(strRptCap, "/", "_")  ' ** Module Function: modStringFuncs.

23400       If strQry <> vbNullString Then

23410         If IsNull(.UserReportPath) = True Then
23420           strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
23430         Else
23440           strRptPath = .UserReportPath
23450         End If

23460         DoCmd.Hourglass False
23470         strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

23480         If strRptPathFile <> vbNullString Then

23490           DoCmd.Hourglass True
23500           DoEvents

23510           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.

23520           If FileExists(CurrentAppPath & LNK_SEP & "TaxLots_xxx.xls") = True Then  ' ** Module Functions: modFileUtilities.
23530             Kill (CurrentAppPath & LNK_SEP & "TaxLots_xxx.xls")
23540           End If

23550           If CurDir <> CurrentAppPath Then  ' ** Module Function: modFileUtilities.
                  ' ** Since I'm not specifying a path in the macro, I want to make sure it's here.
23560             ChDir CurrentAppPath  ' ** Module Function: modFileUtilities.
23570           End If

                '#AVG  BASE QUERIES ASSIGN EITHER averagepriceperunit OR priceperunit!

                ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
23580           Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
23590             Select Case .opgOrderBy
                  Case .opgOrderBy_optAssetDescription.OptionValue
23600               Select Case .chkOptionsOpen
                    Case True
23610                 Select Case blnIsTwo
                      Case True
                        ' ** qryTaxLotReport_07a_02.
23620                   DoCmd.RunMacro "mcrExcelExport_TaxLots_One_Desc_02"
23630                 Case False
                        ' ** qryTaxLotReport_07a_04.
23640                   DoCmd.RunMacro "mcrExcelExport_TaxLots_One_Desc_04"
23650                 End Select
23660               Case False
                      ' ** qryTaxLotReport_07a_01.
23670                 DoCmd.RunMacro "mcrExcelExport_TaxLots_One_Desc_01"
23680               End Select
23690             Case .opgOrderBy_optAssetType.OptionValue
23700               Select Case .chkOptionsOpen
                    Case True
23710                 Select Case blnIsTwo
                      Case True
                        ' ** qryTaxLotReport_07b_02.
23720                   DoCmd.RunMacro "mcrExcelExport_TaxLots_One_Type_02"
23730                 Case False
                        ' ** qryTaxLotReport_07b_04.
23740                   DoCmd.RunMacro "mcrExcelExport_TaxLots_One_Type_04"
23750                 End Select
23760               Case False
                      ' ** qryTaxLotReport_07b_01.
23770                 DoCmd.RunMacro "mcrExcelExport_TaxLots_One_Type_01"
23780               End Select
23790             End Select
23800           Case .opgAccountNumber_optAll.OptionValue
23810             Select Case .opgOrderBy
                  Case .opgOrderBy_optAssetDescription.OptionValue
23820               Select Case .chkOptionsOpen
                    Case True
23830                 Select Case blnIsTwo
                      Case True
                        ' ** qryTaxLotReport_07c_02.
23840                   DoCmd.RunMacro "mcrExcelExport_TaxLots_All_Desc_02"
23850                 Case False
                        ' ** qryTaxLotReport_07c_04.
23860                   DoCmd.RunMacro "mcrExcelExport_TaxLots_All_Desc_04"
23870                 End Select
23880               Case False
                      ' ** qryTaxLotReport_07c_01.
23890                 DoCmd.RunMacro "mcrExcelExport_TaxLots_All_Desc_01"
23900               End Select
23910             Case .opgOrderBy_optAssetType.OptionValue
23920               Select Case .chkOptionsOpen
                    Case True
23930                 Select Case blnIsTwo
                      Case True
                        ' ** qryTaxLotReport_07d_02.
23940                   DoCmd.RunMacro "mcrExcelExport_TaxLots_All_Type_02"
23950                 Case False
                        ' ** qryTaxLotReport_07d_04.
23960                   DoCmd.RunMacro "mcrExcelExport_TaxLots_All_Type_04"
23970                 End Select
23980               Case False
                      ' ** qryTaxLotReport_07d_01.
23990                 DoCmd.RunMacro "mcrExcelExport_TaxLots_All_Type_01"
24000               End Select
24010             End Select
24020           End Select

                ' ** The macro specifies qryTaxLotReport_07x, but cannot be given a dynamic file name.
                ' ** So, it's exported to 'TaxLots_xxx.xls', which is then renamed.
24030           If FileExists(CurrentAppPath & LNK_SEP & "TaxLots_xxx.xls") = True Then  ' ** Module Functions: modFileUtilities.
24040             If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
24050               Kill strRptPathFile
24060             End If
24070             Name (CurrentAppPath & LNK_SEP & "TaxLots_xxx.xls") As (strRptPathFile)  ' ** Module Function: modFileUtilities.
                  ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
24080             OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
24090           End If

24100         End If  ' ** strRptPathFile.

24110       End If  ' ** strQry.

24120     End If  ' ** DoReport().

24130     DoCmd.Hourglass False

24140   End With

EXITP:
24150   Set qdf = Nothing
24160   Set dbs = Nothing
24170   Exit Sub

ERRH:
24180   DoCmd.Hourglass False
24190   Select Case ERR.Number
        Case Else
24200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24210   End Select
24220   Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

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
          Case vbKeyTab
24370       With Me
24380         intRetVal = 0
24390         .cmdClose.SetFocus
24400       End With
24410     End Select
24420   End If

        ' ** Shift keys.
24430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24440     Select Case intRetVal
          Case vbKeyTab
24450       With Me
24460         intRetVal = 0
24470         .cmdWord.SetFocus
24480       End With
24490     End Select
24500   End If

EXITP:
24510   KeyCode = intRetVal
24520   Exit Sub

ERRH:
24530   intRetVal = 0
24540   Select Case ERR.Number
        Case Else
24550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24560   End Select
24570   Resume EXITP

End Sub

Private Function DoReport() As Boolean

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountNo As String, strMsg As String
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

24610   blnRetVal = True

24620   With Me

24630     Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
24640       If IsNull(.cmbAccounts) = True Then
24650         blnRetVal = False
24660         DoCmd.Hourglass False
24670         MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
24680       Else
24690         If .cmbAccounts = vbNullString Then
24700           blnRetVal = False
24710           DoCmd.Hourglass False
24720           MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
24730         End If
24740       End If
24750     Case .opgAccountNumber_optAll.OptionValue
            ' ** Nothing required.
24760     End Select

24770     If blnRetVal = True Then

24780       gstrFormQuerySpec = THIS_NAME

24790       Set dbs = CurrentDb

24800       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
              ' ** ActiveAssets, with add'l fields, UniqueID; single, by specified FormRef().
24810         Set qdf = dbs.QueryDefs("qryTaxLotReport_01")
24820         strAccountNo = .cmbAccounts
24830       Case .opgAccountNumber_optAll.OptionValue
              ' ** ActiveAssets, with add'l fields, UniqueID; all.
24840         Set qdf = dbs.QueryDefs("qryTaxLotReport_02")
24850         strAccountNo = "All"
24860       End Select

24870       Set rst = qdf.OpenRecordset
24880       With rst
24890         If .BOF = True And .EOF = True Then
24900           blnRetVal = False
24910           DoCmd.Hourglass False
24920           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
24930         End If
24940         .Close
24950       End With
24960       Set qdf = Nothing
24970       Set rst = Nothing

24980       If blnRetVal = True Then

24990         Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
                ' ** ActiveAssets, linked to MasterAsset, just marketvaluecurrent = Null, by specified [actno].
25000           Set qdf = dbs.QueryDefs("qryTaxLotReport_09")
25010           With qdf.Parameters
25020             ![actno] = strAccountNo
25030           End With
25040         Case .opgAccountNumber_optAll.OptionValue
                ' ** ActiveAssets, linked to MasterAsset, just marketvaluecurrent = Null, for all.
25050           Set qdf = dbs.QueryDefs("qryTaxLotReport_08")
25060         End Select

25070         Set rst = qdf.OpenRecordset
25080         With rst
25090           If .BOF = True And .EOF = True Then
                  ' ** All's well!
25100           Else
25110             blnRetVal = False
25120             .MoveLast
25130             lngRecs = .RecordCount
25140             .MoveFirst
25150             If lngRecs = 1& Then
25160               strMsg = "An asset currently held has no Market Value." & vbCrLf & vbCrLf
25170               strMsg = strMsg & "Asset: " & ![description_masterasset] & vbCrLf
25180               strMsg = strMsg & "CUSIP: " & ![cusip] & vbCrLf & vbCrLf
25190               strMsg = strMsg & "A value must be entered before this report can be run."
25200             Else
25210               If lngRecs = 2& Then
25220                 strMsg = "Two assets currently held have no Market Value." & vbCrLf & vbCrLf
25230                 strMsg = strMsg & "Asset: " & ![description_masterasset] & vbCrLf
25240                 strMsg = strMsg & "CUSIP: " & ![cusip] & vbCrLf & vbCrLf
25250                 .MoveNext
25260                 strMsg = strMsg & "Asset: " & ![description_masterasset] & vbCrLf
25270                 strMsg = strMsg & "CUSIP: " & ![cusip] & vbCrLf & vbCrLf
25280                 strMsg = strMsg & "Values must be entered before this report can be run."
25290               Else
25300                 strMsg = "Several assets currently held have no Market Value." & vbCrLf & vbCrLf
25310                 For lngX = 1& To lngRecs
25320                   strMsg = strMsg & "Asset: " & ![description_masterasset] & vbCrLf
25330                   If lngX < lngRecs Then .MoveNext
25340                 Next
25350                 strMsg = strMsg & vbCrLf
25360                 strMsg = strMsg & "Values must be entered before this report can be run."
25370               End If
25380             End If
25390           End If
25400           .Close
25410         End With

25420         dbs.Close

25430         If blnRetVal = False Then
25440           Beep
25450           MsgBox strMsg, vbInformation + vbOKOnly, "Entry Required"
25460         End If

25470       Else
25480         dbs.Close
25490       End If  ' ** blnRetVal.

25500     End If  ' ** blnRetVal.

          ' ** Record their current values.
25510     For lngX = 0& To (lngChkGrps - 1&)
25520       If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
25530         arr_varChkGrp(CG_VAL, lngX) = .Controls(arr_varChkGrp(CG_NAM, lngX))
25540       End If
25550     Next

25560   End With  ' ** Me.

EXITP:
25570   Set qdf = Nothing
25580   Set rst = Nothing
25590   Set dbs = Nothing
25600   DoReport = blnRetVal
25610   Exit Function

ERRH:
25620   blnRetVal = False
25630   DoCmd.Hourglass False
25640   Select Case ERR.Number
        Case Else
25650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25660   End Select
25670   Resume EXITP

End Function

Private Sub ChkGroup_Default()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkGroup_Default"

        Dim strProc As String

25710   With Me

25720     strProc = "ckgAlt_opt01_AccountNo_AfterUpdate"
25730     ChkGroup_Select strProc  ' ** Procedure: Below.
25740     strProc = "ckgAlt_opt02_CUSIP_AfterUpdate"
25750     ChkGroup_Select strProc  ' ** Procedure: Below.
25760     strProc = "ckgAlt_opt03_AssetDate_AfterUpdate"
25770     ChkGroup_Select strProc  ' ** Procedure: Below.
25780     strProc = "ckgAlt_opt04_ShareFace_AfterUpdate"
25790     ChkGroup_Select strProc  ' ** Procedure: Below.
25800     strProc = "ckgAlt_opt05_Cost_AfterUpdate"
25810     ChkGroup_Select strProc  ' ** Procedure: Below.

          ' ** When all True, fill in the box.
25820     If .ckgAlt_opt01_AccountNo = True And .ckgAlt_opt02_CUSIP = True And .ckgAlt_opt03_AssetDate = True And _
              .ckgAlt_opt04_ShareFace = True And .ckgAlt_opt05_Cost = True Then
25830       .ckgAltStandard_box.BackColor = MY_CLR_VLTBGE
25840       .ckgAltStandard_box2.BackColor = MY_CLR_VLTBGE
25850       .ckgAltStandard_hline03.BorderColor = MY_CLR_VLTBGE
            ' ** Opening ForeColor: CLR_VDKGRY
            ' ** Lighter: CLR_DKGRY
25860       .ckgAltStandard_lbl.ForeColor = CLR_DKGRY
25870       .ckgAlt_opt01_AccountNo_lbl.ForeColor = CLR_DKGRY
25880       .ckgAlt_opt02_CUSIP_lbl.ForeColor = CLR_DKGRY
25890       .ckgAlt_opt03_AssetDate_lbl.ForeColor = CLR_DKGRY
25900       .ckgAlt_opt04_ShareFace_lbl.ForeColor = CLR_DKGRY
25910       .ckgAlt_opt05_Cost_lbl.ForeColor = CLR_DKGRY
25920     End If

25930   End With

EXITP:
25940   Exit Sub

ERRH:
25950   Select Case ERR.Number
        Case Else
25960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25970   End Select
25980   Resume EXITP

End Sub

Private Sub ChkGroup_Select(strProc As String)

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkGroup_Select"

        Dim strChk As String, strLbl As String, strBox As String
        Dim blnAllTrue As Boolean
        Dim lngX As Long, lngY As Long

26010   With Me

26020     strChk = Left(strProc, (Len(strProc) - Len("_AfterUpdate")))
26030     strLbl = strChk & "_lbl"
26040     strBox = strChk & "_box"
26050     Select Case .Controls(strChk)
          Case True
26060       .Controls(strLbl).FontBold = True
26070       If .chkOptionsOpen = True Then
26080         .Controls(strBox).Visible = True
26090       End If
26100     Case False
26110       .Controls(strLbl).FontBold = False
26120       .Controls(strBox).Visible = False
26130     End Select

26140     blnAllTrue = True
26150     For lngX = 0& To (lngChkGrps - 1&)
26160       If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
26170         If arr_varChkGrp(CG_DEF, lngX) = False Then
26180           If .Controls(arr_varChkGrp(CG_NAM, lngX)) = False Then
26190             blnAllTrue = False  ' ** Only need one to hit.
26200             Exit For
26210           End If
26220         End If
26230       End If
26240     Next

          ' ** When all True, fill in the box.
26250     Select Case blnAllTrue
          Case True
26260       .ckgAltOptional_box.BackColor = MY_CLR_VLTBGE
26270       .ckgAltOptional_box2.BackColor = MY_CLR_VLTBGE
26280       .ckgAltOptional_hline03.BorderColor = MY_CLR_VLTBGE
26290       .ckgAltOptional_lbl.ForeColor = CLR_DKGRY
26300       For lngX = 0& To (lngChkGrps - 1&)
26310         If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
26320           If arr_varChkGrp(CG_DEF, lngX) = False Then
26330             For lngY = 0& To (lngChkGrps - 1&)
26340               If arr_varChkGrp(CG_TYP, lngY) = acLabel Then
26350                 If arr_varChkGrp(CG_PAR, lngY) = lngX Then
26360                   .Controls(arr_varChkGrp(CG_NAM, lngY)).ForeColor = CLR_DKGRY
26370                   Exit For
26380                 End If
26390               End If
26400             Next
26410           End If
26420         End If
26430       Next
26440     Case False
26450       If .ckgAltOptional_box.BackColor = MY_CLR_VLTBGE Then
26460         .ckgAltOptional_box.BackColor = MY_CLR_LTBGE
26470         .ckgAltOptional_box2.BackColor = MY_CLR_LTBGE
26480         .ckgAltOptional_hline03.BorderColor = MY_CLR_LTBGE
26490         .ckgAltOptional_lbl.ForeColor = CLR_VDKGRY
26500       End If
26510       For lngX = 0& To (lngChkGrps - 1&)
26520         If arr_varChkGrp(CG_TYP, lngX) = acLabel And arr_varChkGrp(CG_PAR, lngX) >= 0& Then
26530           If arr_varChkGrp(CG_DEF, arr_varChkGrp(CG_PAR, lngX)) = False Then
26540             If .Controls(arr_varChkGrp(CG_NAM, lngX)).ForeColor = CLR_DKGRY Then
26550               .Controls(arr_varChkGrp(CG_NAM, lngX)).ForeColor = CLR_VDKGRY
26560             End If
26570           End If
26580         End If
26590       Next
26600     End Select

26610   End With

EXITP:
26620   Exit Sub

ERRH:
26630   Select Case ERR.Number
        Case Else
26640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26650   End Select
26660   Resume EXITP

End Sub

Private Function ChkGroup_GetFields(blnIsAll As Boolean) As String

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkGroup_GetFields"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strSQL As String
        Dim blnGetFields As Boolean
        Dim intPos01 As Integer, intPos02 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim lngX As Long
        Dim strRetVal As String

26710   strRetVal = vbNullString

        ' ** qryTaxLotReport_07a_03a, single, no sort, all fields, by specified FormRef('AccountNo').
        ' ** qryTaxLotReport_07a_03b
        ' ** SELECT qryTaxLotReport_07a_03a.[Account Num], qryTaxLotReport_07a_03a.[Short Name],
        ' **   qryTaxLotReport_07a_03a.Asset, qryTaxLotReport_07a_03a.CUSIP,
        ' **   qryTaxLotReport_07a_03a.Type, qryTaxLotReport_07a_03a.[Trade Date],
        ' **   qryTaxLotReport_07a_03a.[Trade DateTime], qryTaxLotReport_07a_03a.[Share/Face],
        ' **   qryTaxLotReport_07a_03a.Cost, qryTaxLotReport_07a_03a.[Unit Price],
        ' **   qryTaxLotReport_07a_03a.[Market Price], qryTaxLotReport_07a_03a.[Market Value],
        ' **   qryTaxLotReport_07a_03a.[Holding Period], qryTaxLotReport_07a_03a.Location, qryTaxLotReport_07a_03a.sort1
        ' ** FROM qryTaxLotReport_07a_03a;

        ' ** qryTaxLotReport_07c_03a, all, no sort, all fields.
        ' ** qryTaxLotReport_07c_03b
        ' ** SELECT qryTaxLotReport_07c_03a.[Account Num], qryTaxLotReport_07c_03a.[Short Name],
        ' **   qryTaxLotReport_07c_03a.Asset, qryTaxLotReport_07c_03a.CUSIP,
        ' **   qryTaxLotReport_07c_03a.Type, qryTaxLotReport_07c_03a.[Trade Date],
        ' **   qryTaxLotReport_07c_03a.[Trade DateTime], qryTaxLotReport_07c_03a.[Share/Face],
        ' **   qryTaxLotReport_07c_03a.Cost, qryTaxLotReport_07c_03a.[Unit Price],
        ' **   qryTaxLotReport_07c_03a.[Market Value], qryTaxLotReport_07c_03a.[Market Price],
        ' **   qryTaxLotReport_07c_03a.[Holding Period], qryTaxLotReport_07c_03a.Location, qryTaxLotReport_07c_03a.sort1
        ' ** FROM qryTaxLotReport_07c_03a;

        ' ** Remember, this has got both [Trade Date], and [Trade DateTime].

        ' ** Don't reassemble the field list if it's already in the array.
26720   blnGetFields = True
26730   For lngX = 0& To (lngChkGrps - 1&)
26740     If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
26750       Select Case blnIsAll
            Case True
26760         If arr_varChkGrp(CG_FAL, lngX) <> vbNullString Then
26770           blnGetFields = False
26780           Exit For
26790         End If
26800       Case False
26810         If arr_varChkGrp(CG_FSN, lngX) <> vbNullString Then
26820           blnGetFields = False
26830           Exit For
26840         End If
26850       End Select
26860     End If
26870   Next

26880   If blnGetFields = True Then

26890     Set dbs = CurrentDb
26900     With dbs

26910       Select Case blnIsAll
            Case True
              ' ** qryTaxLotReport_07c_03a, all, no sort, all fields.
26920         Set qdf = .QueryDefs("qryTaxLotReport_07c_03b")
26930       Case False
              ' ** qryTaxLotReport_07a_03a, single, no sort, all fields, by specified FormRef('AccountNo').
26940         Set qdf = .QueryDefs("qryTaxLotReport_07a_03b")
26950       End Select

26960       strSQL = qdf.SQL
26970       Set qdf = Nothing
26980       strTmp01 = strSQL
26990       intPos01 = InStr(strSQL, "FROM ")
27000       strTmp01 = Left(strSQL, (intPos01 - 1))  ' ** Strip FROM.
27010       If Asc(Right(strTmp01, 1)) = 13 Or Asc(Right(strTmp01, 1)) = 10 Then
27020         strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))  ' ** Strip CRLF.
27030       End If
27040       If Asc(Right(strTmp01, 1)) = 13 Or Asc(Right(strTmp01, 1)) = 10 Then
27050         strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))  ' ** Strip CRLF.
27060       End If
27070       If Left(strTmp01, 7) = "SELECT " Then
27080         strTmp01 = Mid(strTmp01, 8)  ' ** Strip SELECT.
27090       End If
27100       strTmp01 = Trim(strTmp01)

            ' ** Put the SQL fields into the array.
27110       For lngX = 0& To (lngChkGrps - 1&)
27120         If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
27130           strTmp02 = vbNullString: strTmp03 = vbNullString
27140           Select Case arr_varChkGrp(CG_OPT, lngX)
                Case 1&
                  ' ** ckgAlt_opt01_AccountNo.
27150             strTmp03 = "Account Num"
27160           Case 2&
                  ' ** ckgAlt_opt02_CUSIP.
27170             strTmp03 = "CUSIP"
27180           Case 3&
                  ' ** ckgAlt_opt03_AssetDate.
27190             strTmp03 = "Trade Date"  '!!!
27200           Case 4&
                  ' ** ckgAlt_opt04_ShareFace.
27210             strTmp03 = "Share/Face"
27220           Case 5&
                  ' ** ckgAlt_opt05_Cost.
27230             strTmp03 = "Cost"
27240           Case 6&
                  ' ** ckgAlt_opt06_ShortName.
27250             strTmp03 = "Short Name"
27260           Case 7&
                  ' ** ckgAlt_opt07_Description.
27270             strTmp03 = "Asset"
27280           Case 8&
                  ' ** ckgAlt_opt08_AssetType.
27290             strTmp03 = "Type"
27300           Case 9&
                  ' ** ckgAlt_opt09_PricePerUnit.
27310             strTmp03 = "Unit Price"
27320           Case 10&
                  ' ** ckgAlt_opt10_CurrentMarketValue.
27330             strTmp03 = "Market Price"
27340           Case 11&
                  ' ** ckgAlt_opt11_MarketValue.
27350             strTmp03 = "Market Value"
27360           Case 12&
                  ' ** ckgAlt_opt12_HoldingPeriod.
27370             strTmp03 = "Holding Period"
27380           Case 13&
                  ' ** ckgAlt_opt13_LocationName.
27390             strTmp03 = "Location"
27400           Case 14&
                  ' ** ckgAlt_opt14_AssetDate_Time.
27410             strTmp03 = "Trade DateTime"  '!!!
27420           End Select
27430           If Left(strTmp03, 10) = "Trade Date" Then
27440             intPos02 = InStr(strTmp01, "Trade DateTime")
27450             intPos01 = InStr(strTmp01, "Trade Date")
27460             If intPos01 = intPos02 Then intPos01 = InStr((intPos02 + 1), strTmp01, "Trade Date")
27470             If strTmp03 = "Trade DateTime" Then intPos01 = intPos02
27480             intPos01 = intPos01 + Len(strTmp03)
27490           Else
27500             intPos01 = InStr(strTmp01, strTmp03) + Len(strTmp03)
27510           End If
27520           If intPos01 >= Len(strTmp01) Then
27530             intPos01 = Len(strTmp01)
27540           Else
27550             If Mid(strTmp01, intPos01, 1) = "]" Then intPos01 = intPos01 + 1
27560           End If
27570           strTmp02 = Left(strTmp01, intPos01)  ' ** Ends in our field.
27580           If Right(strTmp02, 1) = "," Then strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
27590           intPos01 = InStr(strTmp02, ",")
27600           Do While intPos01 > 0
27610             strTmp02 = Trim(Mid(strTmp02, (intPos01 + 1)))  ' ** Strip off preceding fields.
27620             intPos01 = InStr(strTmp02, ",")
27630           Loop
27640           Select Case blnIsAll
                Case True
27650             arr_varChkGrp(CG_FAL, lngX) = strTmp02
27660           Case False
27670             arr_varChkGrp(CG_FSN, lngX) = strTmp02
27680           End Select
27690         End If  ' ** acCheckBox.
27700       Next  ' ** lngX.

27710       .Close
27720     End With  ' ** dbs.
27730     Set dbs = Nothing

27740     strTmp01 = vbNullString
27750     intPos01 = InStr(strSQL, "sort1")
27760     If intPos01 > 0 Then
27770       intPos02 = intPos01 + Len("sort1")  ' ** 1 after sort1.
27780       If Mid(strSQL, intPos02, 1) = "]" Then intPos02 = (intPos02 + 1)
27790       strTmp01 = Left(strSQL, (intPos02 - 1))
27800       intPos01 = InStr(strTmp01, ",")
27810       Do While intPos01 > 0
27820         strTmp01 = Trim(Mid(strTmp01, (intPos01 + 1)))  ' ** Strip off preceding fields.
27830         intPos01 = InStr(strTmp01, ",")
27840       Loop
27850       strRetVal = strTmp01
27860     End If

27870   Else

27880     Select Case blnIsAll
          Case True
27890       If strSortField_All = vbNullString Then blnGetFields = True
27900     Case False
27910       If strSortField_Single = vbNullString Then blnGetFields = True
27920     End Select

27930     Select Case blnGetFields
          Case True

27940       Set dbs = CurrentDb
27950       With dbs
27960         Select Case blnIsAll
              Case True
                ' ** qryTaxLotReport_07c_03a, all, no sort, all fields.
27970           Set qdf = .QueryDefs("qryTaxLotReport_07c_03b")
27980         Case False
                ' ** qryTaxLotReport_07a_03a, single, no sort, all fields, by specified FormRef('AccountNo').
27990           Set qdf = .QueryDefs("qryTaxLotReport_07a_03b")
28000         End Select
28010         strSQL = qdf.SQL
28020         Set qdf = Nothing
28030         .Close
28040       End With
28050       Set dbs = Nothing

28060       strTmp01 = vbNullString
28070       intPos01 = InStr(strSQL, "sort1")
28080       If intPos01 > 0 Then
28090         intPos02 = intPos01 + Len("sort1")  ' ** 1 after sort1.
28100         If Mid(strSQL, intPos02, 1) = "]" Then intPos02 = (intPos02 + 1)
28110         strTmp01 = Left(strSQL, (intPos02 - 1))
28120         intPos01 = InStr(strTmp01, ",")
28130         Do While intPos01 > 0
28140           strTmp01 = Trim(Mid(strTmp01, (intPos01 + 1)))  ' ** Strip off preceding fields.
28150           intPos01 = InStr(strTmp01, ",")
28160         Loop
28170         strRetVal = strTmp01
28180       End If

28190     Case False
28200       Select Case blnIsAll
            Case True
28210         strRetVal = strSortField_All
28220       Case False
28230         strRetVal = strSortField_Single
28240       End Select
28250     End Select

28260   End If  ' ** blnGetFields.

EXITP:
28270   Set qdf = Nothing
28280   Set dbs = Nothing
28290   ChkGroup_GetFields = strRetVal
28300   Exit Function

ERRH:
28310   strRetVal = vbNullString
28320   Select Case ERR.Number
        Case Else
28330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28340   End Select
28350   Resume EXITP

End Function

Private Sub ChkGroup_Show(strMode As String)

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkGroup_Show"

        Dim ctl As Access.Control, chk As Access.CheckBox, lbl As Access.Label, lin As Access.Line, box As Access.Rectangle
        Dim strCtl1 As String, strCtl2 As String, strCtl3 As String
        Dim lngTmp01 As Long
        Dim lngX As Long, lngY As Long, lngE As Long

28410   With Me

28420     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
28430     lngMonitorNum = 1&: lngTmp01 = 0&
28440     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
28450     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

28460     Select Case strMode
          Case "Load"

28470       lngChkGrps = 0&
28480       ReDim arr_varChkGrp(CG_ELEMS, 0)

28490       For Each ctl In .Controls
28500         If Left(ctl.Name, 6) = "ckgAlt" And ctl.Name <> "ckgAlt_box4" Then
28510           Select Case ctl.ControlType
                Case acCheckBox
28520             Set chk = .Controls(ctl.Name)
28530             With chk
28540               lngChkGrps = lngChkGrps + 1&
28550               lngE = lngChkGrps - 1&
28560               ReDim Preserve arr_varChkGrp(CG_ELEMS, lngE)
28570               arr_varChkGrp(CG_NAM, lngE) = .Name
28580               arr_varChkGrp(CG_TYP, lngE) = acCheckBox
28590               arr_varChkGrp(CG_TOP, lngE) = .Top
28600               arr_varChkGrp(CG_LFT, lngE) = .Left
28610               arr_varChkGrp(CG_HGT, lngE) = Null
28620               arr_varChkGrp(CG_DEF, lngE) = CBool(.DefaultValue)
28630               arr_varChkGrp(CG_PAR, lngE) = CLng(-1)
28640               arr_varChkGrp(CG_VAL, lngE) = CBool(False)  ' ** Start everybody False.
28650               arr_varChkGrp(CG_FAL, lngE) = vbNullString
28660               arr_varChkGrp(CG_FSN, lngE) = vbNullString
28670               arr_varChkGrp(CG_OPT, lngE) = Val(Mid(.Name, 11, 2))  ' ** ckgAlt_opt01_AccountNo.
28680             End With
28690           Case acLabel
28700             Set lbl = .Controls(ctl.Name)
28710             With lbl
28720               If Left(.Name, 10) = "ckgAlt_opt" Then
28730                 lngChkGrps = lngChkGrps + 1&
28740                 lngE = lngChkGrps - 1&
28750                 ReDim Preserve arr_varChkGrp(CG_ELEMS, lngE)
28760                 arr_varChkGrp(CG_NAM, lngE) = .Name
28770                 arr_varChkGrp(CG_TYP, lngE) = acLabel
28780                 arr_varChkGrp(CG_TOP, lngE) = .Top
28790                 arr_varChkGrp(CG_LFT, lngE) = .Left
28800                 arr_varChkGrp(CG_HGT, lngE) = Null
28810                 arr_varChkGrp(CG_DEF, lngE) = CBool(False)
28820                 arr_varChkGrp(CG_PAR, lngE) = CLng(-1)
28830                 arr_varChkGrp(CG_VAL, lngE) = CBool(False)
28840                 arr_varChkGrp(CG_FAL, lngE) = vbNullString
28850                 arr_varChkGrp(CG_FSN, lngE) = vbNullString
28860                 arr_varChkGrp(CG_OPT, lngE) = Val(Mid(.Name, 11, 2))  ' ** ckgAlt_opt01_AccountNo_lbl.
28870               Else
                      ' ** I may be doing something different with the main box labels.
28880                 lngChkGrps = lngChkGrps + 1&
28890                 lngE = lngChkGrps - 1&
28900                 ReDim Preserve arr_varChkGrp(CG_ELEMS, lngE)
28910                 arr_varChkGrp(CG_NAM, lngE) = .Name
28920                 arr_varChkGrp(CG_TYP, lngE) = acLabel
28930                 arr_varChkGrp(CG_TOP, lngE) = .Top
28940                 arr_varChkGrp(CG_LFT, lngE) = .Left
28950                 arr_varChkGrp(CG_HGT, lngE) = Null
28960                 arr_varChkGrp(CG_DEF, lngE) = CBool(False)
28970                 arr_varChkGrp(CG_PAR, lngE) = CLng(-1)
28980                 arr_varChkGrp(CG_VAL, lngE) = CBool(False)
28990                 arr_varChkGrp(CG_FAL, lngE) = vbNullString
29000                 arr_varChkGrp(CG_FSN, lngE) = vbNullString
29010                 arr_varChkGrp(CG_OPT, lngE) = CLng(0)
29020               End If
29030             End With
29040           Case acLine
29050             Set lin = .Controls(ctl.Name)
29060             With lin
29070               lngChkGrps = lngChkGrps + 1&
29080               lngE = lngChkGrps - 1&
29090               ReDim Preserve arr_varChkGrp(CG_ELEMS, lngE)
29100               arr_varChkGrp(CG_NAM, lngE) = .Name
29110               arr_varChkGrp(CG_TYP, lngE) = acLine
29120               arr_varChkGrp(CG_TOP, lngE) = .Top
29130               arr_varChkGrp(CG_LFT, lngE) = .Left
29140               arr_varChkGrp(CG_HGT, lngE) = .Height  ' ** May be Zero for horizontal lines.
29150               arr_varChkGrp(CG_DEF, lngE) = CBool(False)
29160               arr_varChkGrp(CG_PAR, lngE) = CLng(-1)
29170               arr_varChkGrp(CG_VAL, lngE) = CBool(False)
29180               arr_varChkGrp(CG_FAL, lngE) = vbNullString
29190               arr_varChkGrp(CG_FSN, lngE) = vbNullString
29200               arr_varChkGrp(CG_OPT, lngE) = CLng(0)
29210             End With
29220           Case acRectangle
29230             Set box = .Controls(ctl.Name)
29240             With box
29250               If .Name <> "ckgAlt_box3" Then
29260                 lngChkGrps = lngChkGrps + 1&
29270                 lngE = lngChkGrps - 1&
29280                 ReDim Preserve arr_varChkGrp(CG_ELEMS, lngE)
29290                 arr_varChkGrp(CG_NAM, lngE) = .Name
29300                 arr_varChkGrp(CG_TYP, lngE) = acRectangle
29310                 arr_varChkGrp(CG_TOP, lngE) = .Top
29320                 arr_varChkGrp(CG_LFT, lngE) = .Left
29330                 arr_varChkGrp(CG_HGT, lngE) = .Height
29340                 arr_varChkGrp(CG_DEF, lngE) = CBool(False)
29350                 arr_varChkGrp(CG_PAR, lngE) = CLng(-1)
29360                 arr_varChkGrp(CG_VAL, lngE) = CBool(False)
29370                 arr_varChkGrp(CG_FAL, lngE) = vbNullString
29380                 arr_varChkGrp(CG_FSN, lngE) = vbNullString
29390                 arr_varChkGrp(CG_OPT, lngE) = Val(Mid(.Name, 11, 2))  ' ** ckgAlt_opt01_AccountNo_box.
29400               End If
29410             End With
29420           Case acCommandButton
                  ' ** Don't need them.
29430           End Select
29440         End If
29450       Next  ' ** ctl.

            ' ** Put the parent's element number in the label and box.
29460       For lngX = 0& To (lngChkGrps - 1&)
29470         If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
29480           arr_varChkGrp(CG_PAR, lngX) = CLng(0)
29490           strCtl1 = arr_varChkGrp(CG_NAM, lngX)
29500           strCtl2 = strCtl1 & "_lbl"
29510           strCtl3 = strCtl1 & "_box"
29520           For lngY = 0& To (lngChkGrps - 1&)
29530             If arr_varChkGrp(CG_TYP, lngY) = acLabel Then
29540               If arr_varChkGrp(CG_NAM, lngY) = strCtl2 Then
29550                 arr_varChkGrp(CG_PAR, lngY) = lngX
29560                 Exit For
29570               End If
29580             End If
29590           Next
29600           For lngY = 0& To (lngChkGrps - 1&)
29610             If arr_varChkGrp(CG_TYP, lngY) = acRectangle Then
29620               If arr_varChkGrp(CG_NAM, lngY) = strCtl3 Then
29630                 arr_varChkGrp(CG_PAR, lngY) = lngX
29640                 Exit For
29650               End If
29660             End If
29670           Next
29680         End If
29690       Next

29700     Case "Hide"

            ' ** Hide them.
29710       .ckgAlt_box.Visible = False
29720       .ckgAlt_box2.Visible = False
29730       .ckgAlt_vline01.Visible = False
29740       .ckgAlt_vline02.Visible = False
29750       .ckgAlt_vline03.Visible = False
29760       .ckgAlt_vline04.Visible = False
29770       .ckgAlt_hline01.Visible = False
29780       .ckgAlt_hline02.Visible = False
29790       .ckgAlt_box3.Visible = True

            ' ** First make everything invisible.
29800       For lngX = 0& To (lngChkGrps - 1&)
29810         strCtl1 = arr_varChkGrp(CG_NAM, lngX)
29820         .Controls(strCtl1).Visible = False
29830       Next

            ' ** Now move them all up and out of the way.
29840       For lngX = 0& To (lngChkGrps - 1&)
29850         strCtl1 = arr_varChkGrp(CG_NAM, lngX)
29860         Select Case arr_varChkGrp(CG_TYP, lngX)
              Case acCheckBox
29870           .Controls(strCtl1).Top = (4& * lngTpp)  ' ** Though the Twips-Per-Pixel really doesn't matter.
29880           arr_varChkGrp(CG_VAL, lngX) = .Controls(strCtl1)
29890         Case acLabel
29900           .Controls(strCtl1).Top = (4& * lngTpp)
29910         Case acLine
29920           .Controls(strCtl1).Top = (4& * lngTpp)
29930           If Right(strCtl1, 8) = "_vline05" Or Right(strCtl1, 8) = "_vline06" Then
29940             .Controls(strCtl1).Height = (4& * lngTpp)
29950           End If
29960         Case acRectangle
29970           If InStr(strCtl1, "_opt") > 0 Then
29980             .Controls(strCtl1).Top = (4& * lngTpp)
29990           ElseIf Right(strCtl1, 5) = "_box2" Then
30000             .Controls(strCtl1).Top = (4& * lngTpp)
30010           Else
30020             .Controls(strCtl1).Top = (4& * lngTpp)
30030             .Controls(strCtl1).Height = (4& * lngTpp)
30040           End If
30050         End Select
30060       Next

            ' ** Also move these.
30070       .cmdExcel_lbl.Visible = False
30080       .cmdExcel_lbl.Top = (4& * lngTpp)
30090       .Rights_Microsoft_lbl1.Visible = False
30100       .Rights_Microsoft_lbl1.Top = (4& * lngTpp)
30110       .Rights_Microsoft_lbl2.Visible = False
30120       .Rights_Microsoft_lbl2.Top = (4& * lngTpp)
30130       gblnMessage = .cmdReset.Enabled
30140       .cmdReset.Enabled = False
30150       .cmdReset.Top = 0&
30160       .cmdReset_raised_img.Visible = False
30170       .cmdReset_raised_img.Top = 0&
30180       .cmdReset_raised_semifocus_dots_img.Visible = False
30190       .cmdReset_raised_semifocus_dots_img.Top = 0&
30200       .cmdReset_raised_focus_img.Visible = False
30210       .cmdReset_raised_focus_img.Top = 0&
30220       .cmdReset_raised_focus_dots_img.Visible = False
30230       .cmdReset_raised_focus_dots_img.Top = 0&
30240       .cmdReset_sunken_focus_dots_img.Visible = False
30250       .cmdReset_sunken_focus_dots_img.Top = 0&
30260       .cmdReset_raised_img_dis.Visible = False
30270       .cmdReset_raised_img_dis.Top = 0&

            ' ** In order that ckgAlt_opt07_Description opens the same way it closed,
            ' ** rather than alternating, flip it as it's closing.
30280       Select Case .ckgAlt_opt07_Description.SpecialEffect
            Case acSpecialEffectFlat
              ' ** TRUE   FLAT
              ' ** FALSE  FLAT
30290         .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectSunken
30300       Case acSpecialEffectSunken
              ' ** TRUE   SUNKEN
              ' ** FALSE  SUNKEN
30310         .ckgAlt_opt07_Description.SpecialEffect = acSpecialEffectFlat
30320       End Select

            ' ** And change the button image.
30330       .cmdExcelOptions_R_raised_img.Visible = True
30340       .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = False
30350       .cmdExcelOptions_R_raised_focus_img.Visible = False
30360       .cmdExcelOptions_R_raised_focus_dots_img.Visible = False
30370       .cmdExcelOptions_R_sunken_focus_dots_img.Visible = False
30380       .cmdExcelOptions_R_raised_img_dis.Visible = False
30390       .cmdExcelOptions_L_raised_img.Visible = False
30400       .cmdExcelOptions_L_raised_semifocus_dots_img.Visible = False
30410       .cmdExcelOptions_L_raised_focus_img.Visible = False
30420       .cmdExcelOptions_L_raised_focus_dots_img.Visible = False
30430       .cmdExcelOptions_L_sunken_focus_dots_img.Visible = False
30440       .cmdExcelOptions_L_raised_img_dis.Visible = False
30450       .Detail.Height = (lngDetail_Height - lngDiff)

            ' ** Now close up the form.
30460       DoCmd.SelectObject acForm, THIS_NAME, False
30470       If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
30480       DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, (lngFrm_Height - lngFrm_Diff)  'lngFrm_Top
30490       If lngMonitorNum > 1& Then
30500         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
30510       End If
30520       .chkOptionsOpen = False

30530     Case "Show"

            ' ** Show them.
30540       .ckgAlt_box3.Visible = False
30550       .ckgAlt_box.Visible = True
30560       .ckgAlt_box2.Visible = True
30570       .ckgAlt_vline01.Visible = True
30580       .ckgAlt_vline02.Visible = True
30590       .ckgAlt_vline03.Visible = True
30600       .ckgAlt_vline04.Visible = True
30610       .ckgAlt_hline01.Visible = True
30620       .ckgAlt_hline02.Visible = True

            ' ** Start by opening up the form.
30630       DoCmd.SelectObject acForm, THIS_NAME, False
30640       If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
30650       DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, lngFrm_Height  'lngFrm_Top
30660       If lngMonitorNum > 1& Then
30670         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
30680       End If

            ' ** Now move them back to where they were.
30690       For lngX = 0& To (lngChkGrps - 1&)
30700         strCtl1 = arr_varChkGrp(CG_NAM, lngX)
30710         Select Case arr_varChkGrp(CG_TYP, lngX)
              Case acCheckBox
30720           .Controls(strCtl1).Top = arr_varChkGrp(CG_TOP, lngX)
30730           arr_varChkGrp(CG_VAL, lngX) = .Controls(strCtl1)
30740         Case acLabel
30750           .Controls(strCtl1).Top = arr_varChkGrp(CG_TOP, lngX)
30760         Case acLine
30770           .Controls(strCtl1).Top = arr_varChkGrp(CG_TOP, lngX)
30780           If Right(strCtl1, 8) = "_vline05" Or Right(strCtl1, 8) = "_vline06" Then
30790             .Controls(strCtl1).Height = arr_varChkGrp(CG_HGT, lngX)
30800           End If
30810         Case acRectangle
30820           .Controls(strCtl1).Top = arr_varChkGrp(CG_TOP, lngX)
30830           If InStr(strCtl1, "_opt") = 0 And Right(strCtl1, 5) <> "_box2" Then
30840             .Controls(strCtl1).Height = arr_varChkGrp(CG_HGT, lngX)
30850           End If
30860         End Select
30870       Next

            ' ** Then restore these.
30880       .cmdExcel_lbl.Visible = True
30890       .cmdExcel_lbl.Top = .ckgAltOptional_box.Top
30900       lngTmp01 = (((.ckgAltOptional_box.Top + .ckgAltOptional_box.Height) - .Rights_Microsoft_lbl2.Height) + (5& * lngTpp))
30910       .Rights_Microsoft_lbl2.Top = lngTmp01
30920       .Rights_Microsoft_lbl1.Top = (.Rights_Microsoft_lbl2.Top - lngRights_Offset)
30930       .Rights_Microsoft_lbl1.Visible = True
30940       .Rights_Microsoft_lbl2.Visible = True
30950       .cmdReset.Top = lngReset_Top
30960       Select Case blnIsOpen
            Case True
30970         .cmdReset.Enabled = True
30980       Case False
30990         .cmdReset.Enabled = gblnMessage
31000       End Select
31010       .cmdReset_raised_img.Top = lngReset_Top
31020       .cmdReset_raised_semifocus_dots_img.Top = lngReset_Top
31030       .cmdReset_raised_focus_img.Top = lngReset_Top
31040       .cmdReset_raised_focus_dots_img.Top = lngReset_Top
31050       .cmdReset_sunken_focus_dots_img.Top = lngReset_Top
31060       .cmdReset_raised_img_dis.Top = lngReset_Top
31070       Select Case .cmdReset.Enabled
            Case True
31080         .cmdReset_raised_img.Visible = True
31090         .cmdReset_raised_img_dis.Visible = False
31100       Case False
31110         .cmdReset_raised_img_dis.Visible = True
31120         .cmdReset_raised_img.Visible = False
31130       End Select
31140       .cmdReset_raised_semifocus_dots_img.Visible = False
31150       .cmdReset_raised_focus_img.Visible = False
31160       .cmdReset_raised_focus_dots_img.Visible = False
31170       .cmdReset_sunken_focus_dots_img.Visible = False

            ' ** and finally make them visible.
31180       For lngX = 0& To (lngChkGrps - 1&)
31190         strCtl1 = arr_varChkGrp(CG_NAM, lngX)
31200         Select Case arr_varChkGrp(CG_TYP, lngX)
              Case acCheckBox
31210           .Controls(strCtl1).Visible = True
31220         Case acLabel
31230           If Right(strCtl1, 5) = "_lbl2" Then
31240             Select Case .opgOrderBy
                  Case .opgOrderBy_optAssetDescription.OptionValue
31250               If InStr(strCtl1, "Description") > 0 Then
31260                 .Controls(strCtl1).Visible = True
31270               End If
31280             Case .opgOrderBy_optAssetType.OptionValue
31290               If InStr(strCtl1, "AssetType") > 0 Then
31300                 .Controls(strCtl1).Visible = True
31310               End If
31320             End Select
31330           Else
31340             .Controls(strCtl1).Visible = True
31350           End If
31360         Case acLine
31370           .Controls(strCtl1).Visible = True
31380         Case acRectangle
31390           If InStr(strCtl1, "_opt") = 0 Then
31400             .Controls(strCtl1).Visible = True
31410           Else
31420             strCtl2 = Left(strCtl1, (Len(strCtl1) - 4))
31430             If .Controls(strCtl2) = True Then
31440               .Controls(strCtl1).Visible = True
31450             End If
31460           End If
31470         End Select
31480       Next

            ' ** Check FontBold.
31490       For lngX = 0& To (lngChkGrps - 1&)
31500         If Left(arr_varChkGrp(CG_NAM, lngX), 10) = "ckgAlt_opt" Then
31510           If arr_varChkGrp(CG_TYP, lngX) = acCheckBox Then
31520             strCtl1 = arr_varChkGrp(CG_NAM, lngX) & "_AfterUpdate"
31530             ChkGroup_Select strCtl1  ' ** Procedure: Above.
31540           End If
31550         End If
31560       Next

            ' ** Check .opgOrderBy and .ckgAlt_opt08_AssetType.
31570       If .opgOrderBy = .opgOrderBy_optAssetType.OptionValue Then
31580         If .ckgAlt_opt08_AssetType = False Then
31590           .ckgAlt_opt08_AssetType = True
31600           ckgAlt_opt08_AssetType_AfterUpdate  ' ** Procedure: Above.
31610           DoEvents
31620           .ckgAlt_opt08_AssetType.Locked = True
31630         End If
31640       End If

            ' ** And change the button image.
31650       .cmdExcelOptions_L_raised_img.Visible = True
31660       .cmdExcelOptions_L_raised_semifocus_dots_img.Visible = False
31670       .cmdExcelOptions_L_raised_focus_img.Visible = False
31680       .cmdExcelOptions_L_raised_focus_dots_img.Visible = False
31690       .cmdExcelOptions_L_sunken_focus_dots_img.Visible = False
31700       .cmdExcelOptions_L_raised_img_dis.Visible = False
31710       .cmdExcelOptions_R_raised_img.Visible = False
31720       .cmdExcelOptions_R_raised_semifocus_dots_img.Visible = False
31730       .cmdExcelOptions_R_raised_focus_img.Visible = False
31740       .cmdExcelOptions_R_raised_focus_dots_img.Visible = False
31750       .cmdExcelOptions_R_sunken_focus_dots_img.Visible = False
31760       .cmdExcelOptions_R_raised_img_dis.Visible = False
31770       .chkOptionsOpen = True

31780     End Select
31790     DoEvents
31800     .Repaint
31810   End With

EXITP:
31820   Set chk = Nothing
31830   Set lbl = Nothing
31840   Set lin = Nothing
31850   Set box = Nothing
31860   Set ctl = Nothing
31870   Exit Sub

ERRH:
31880   Select Case ERR.Number
        Case Else
31890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31900   End Select
31910   Resume EXITP

End Sub

Private Sub AcctCtls_Move(strAction As String)

32000 On Error GoTo ERRH

        Const THIS_PROC As String = "AcctCtls_Move"

        Dim ctl As Access.Control
        Dim lngX As Long, lngE As Long

32010   With Me
32020     Select Case strAction
          Case "Load"

32030       lngAcctCtls = 0&
32040       ReDim arr_varAcctCtl(A_ELEMS, 0)

32050       For Each ctl In .Detail.Controls
32060         With ctl
32070           If Left(.Name, 11) = "cmbAccounts" Then
                  ' ** cmbAccounts
                  ' ** cmbAccounts_lbl
                  ' ** cmbAccounts_lbl_box
                  ' ** cmbAccounts_box
32080             lngAcctCtls = lngAcctCtls + 1&
32090             lngE = lngAcctCtls - 1&
32100             ReDim Preserve arr_varAcctCtl(A_ELEMS, lngE)
32110             arr_varAcctCtl(A_CNAM, lngE) = .Name
32120             arr_varAcctCtl(A_TYP, lngE) = .ControlType
32130             arr_varAcctCtl(A_TOP, lngE) = .Top
32140           ElseIf Left(.Name, 16) = "opgAccountSource" Then
                  ' ** opgAccountSource
                  ' ** opgAccountSource_box
                  ' ** opgAccountSource_optNumber
                  ' ** opgAccountSource_optNumber_lbl
                  ' ** opgAccountSource_optNumber_lbl2
                  ' ** opgAccountSource_optNumber_lbl2_dim_hi
                  ' ** opgAccountSource_optName
                  ' ** opgAccountSource_optName_lbl
                  ' ** opgAccountSource_optName_lbl2
                  ' ** opgAccountSource_optName_lbl2_dim_hi
32150             lngAcctCtls = lngAcctCtls + 1&
32160             lngE = lngAcctCtls - 1&
32170             ReDim Preserve arr_varAcctCtl(A_ELEMS, lngE)
32180             arr_varAcctCtl(A_CNAM, lngE) = .Name
32190             arr_varAcctCtl(A_TYP, lngE) = .ControlType
32200             arr_varAcctCtl(A_TOP, lngE) = .Top
32210           ElseIf Left(.Name, 13) = "chkRememberMe" Then
                  ' ** chkRememberMe
                  ' ** chkRememberMe_lbl
                  ' ** chkRememberMe_lbl2_dim
                  ' ** chkRememberMe_lbl2_dim_hi
32220             lngAcctCtls = lngAcctCtls + 1&
32230             lngE = lngAcctCtls - 1&
32240             ReDim Preserve arr_varAcctCtl(A_ELEMS, lngE)
32250             arr_varAcctCtl(A_CNAM, lngE) = .Name
32260             arr_varAcctCtl(A_TYP, lngE) = .ControlType
32270             arr_varAcctCtl(A_TOP, lngE) = .Top
32280           ElseIf Left(.Name, 10) = "opgOrderBy" Then
                  ' ** opgOrderBy
                  ' ** opgOrderBy_box
                  ' ** opgOrderBy_optAssetDescription
                  ' ** opgOrderBy_optAssetDescription_lbl
                  ' ** opgOrderBy_optAssetDescription_lbl2
                  ' ** opgOrderBy_optAssetType
                  ' ** opgOrderBy_optAssetType_lbl
                  ' ** opgOrderBy_optAssetType_lbl2
32290             lngAcctCtls = lngAcctCtls + 1&
32300             lngE = lngAcctCtls - 1&
32310             ReDim Preserve arr_varAcctCtl(A_ELEMS, lngE)
32320             arr_varAcctCtl(A_CNAM, lngE) = .Name
32330             arr_varAcctCtl(A_TYP, lngE) = .ControlType
32340             arr_varAcctCtl(A_TOP, lngE) = .Top
32350           ElseIf Left(.Name, 9) = "chkPageOf" Then
                  ' ** chkPageOf
                  ' ** chkPageOf_box
                  ' ** chkPageOf_lbl
32360             lngAcctCtls = lngAcctCtls + 1&
32370             lngE = lngAcctCtls - 1&
32380             ReDim Preserve arr_varAcctCtl(A_ELEMS, lngE)
32390             arr_varAcctCtl(A_CNAM, lngE) = .Name
32400             arr_varAcctCtl(A_TYP, lngE) = .ControlType
32410             arr_varAcctCtl(A_TOP, lngE) = .Top
32420           End If
32430         End With  ' ** ctl.
32440       Next  ' ** ctl.

            'For lngX = 0& To (lngAcctCtls - 1&)
            '  Debug.Print "'" & arr_varAcctCtl(A_CNAM, lngX)
            'Next

32450     Case "Hide"
32460       For lngX = 0& To (lngAcctCtls - 1&)
32470         .Controls(arr_varAcctCtl(A_CNAM, lngX)).Top = (arr_varAcctCtl(A_TOP, lngX) - lngDiff)
32480       Next
32490       .opgAccountSource.Height = .opgAccountSource_box.Height
32500       .opgOrderBy.Height = .opgOrderBy_box.Height

32510     Case "Show"

32520       For lngX = 0& To (lngAcctCtls - 1&)
32530         .Controls(arr_varAcctCtl(A_CNAM, lngX)).Top = arr_varAcctCtl(A_TOP, lngX)
32540       Next
32550       .opgAccountSource.Height = .opgAccountSource_box.Height
32560       .opgOrderBy.Height = .opgOrderBy_box.Height

32570     End Select  ' ** strAction
32580   End With

EXITP:
32590   Set ctl = Nothing
32600   Exit Sub

ERRH:
32610   DoCmd.Hourglass False
32620   Select Case ERR.Number
        Case Else
32630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32640   End Select
32650   Resume EXITP

End Sub
