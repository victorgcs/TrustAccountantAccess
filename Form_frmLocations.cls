VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLocations"

'VGC 07/23/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Load()
' ##   EnableCountry_SetFrmWidth()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd}
' **   Close:            C {cmdClose}
' **   Delete:           D {cmdDelete}
' **   Enable Country:   E {chkEnableCountry}
' **   Name:             N {Loc_Name on frmLocations_Sub}
' **   Print Report:     P {cmdPrintReport}
' **   Move Assets:      V {cmdMoveAssets}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmLocations_Sub}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private lngFrm_Left_Orig As Long, lngFrm_Top_Orig As Long, lngFrm_Width_Orig As Long, lngFrm_Height_Orig As Long
Private lngForm_Width As Long, lngCountry_Diff As Long, lngSub_Width As Long
Private lngClose_Left As Long, lngDelete_Left As Long, lngAdd_Left As Long, lngMoveAssets_Left As Long
Private lngEnableCountry_Left As Long, lngEnableCountryLbl_Left As Long

Private blnGoingToReport As Boolean, lngPrintReportLbl_Left As Long, blnWasGTR As Boolean
Private blnJustClose As Boolean
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean
        Dim lngTmp01 As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
170           blnJustClose = False
180           blnWasGTR = gblnGoToReport

190           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

200           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
210           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

220           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
230           lngMonitorNum = 1&: lngTmp01 = 0&
240           EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
250           If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

              ' ** Variables are fed empty, then populated ByRef.
260           GetFormDimensions Me, lngFrm_Left_Orig, lngFrm_Top_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig  ' ** Module Function: modWindowFunctions.

270           FrmDimSave THIS_NAME, lngFrm_Left_Orig, lngFrm_Top_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig  ' ** Module Function: modMonitorFuncs.

280           If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top_Orig
290           DoCmd.MoveSize lngFrm_Left_Orig, lngTmp01, lngFrm_Width_Orig, lngFrm_Height_Orig  'lngFrm_Top_Orig
300           If lngMonitorNum > 1& Then
310             LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
320           End If

330           lngForm_Width = .Width
340           lngSub_Width = .frmLocations_Sub.Width
350           lngClose_Left = .cmdClose.Left
360           lngDelete_Left = .cmdDelete.Left
370           lngAdd_Left = .cmdAdd.Left
380           lngMoveAssets_Left = .cmdMoveAssets.Left
390           lngEnableCountry_Left = .chkEnableCountry.Left
400           lngEnableCountryLbl_Left = .chkEnableCountry_lbl.Left

410           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

420           lngCountry_Diff = (.frmLocations_Sub.Form.Width - .frmLocations_Sub.Form.Loc_Country.Left)

430           SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.
              'SystemColor_Move Me  ' ** Module Procedure: modSystemColorFuncs.
440           StateCodeQrySet Me  ' ** Module Function: modPreferenceFuncs.

450           chkEnableCountry_AfterUpdate  ' ** Procedure: Below.

460           blnIsOpen = False

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

710     With Me
720       If IsNull(.Loc_State_Pref) = False Then
730         .frmLocations_Sub.Form.Loc_State.DefaultValue = Chr(34) & .Loc_State_Pref & Chr(34)
740       End If
750     End With

EXITP:
760     Exit Sub

ERRH:
770     DoCmd.Hourglass False
780     Select Case ERR.Number
        Case Else
790       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
800     End Select
810     Resume EXITP

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

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmLocations_Sub}

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
          Case vbKeyN
1030        With Me
1040          intRetVal = 0               ' ** Needed because labels aren't connected to
1050          .frmLocations_Sub.SetFocus  ' ** text boxes in order to permit DblClicking.
1060          .frmLocations_Sub.Form.Loc_Name.SetFocus
1070        End With
1080      Case vbKeyX
1090        intRetVal = 0
1100        cmdClose_Click  ' ** Procedure: Below.
1110      End Select
1120    End If

        ' ** Ctrl keys.
1130    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1140      Select Case intRetVal
          Case vbKeyS
1150        With Me
1160          intRetVal = 0
1170          .frmLocations_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmLocations_Sub.
1180        End With
1190      End Select
1200    End If

        ' ** Ctrl-Shift keys.
1210    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1220      Select Case intRetVal
          Case vbKeyF
1230        With Me
1240          intRetVal = 0
1250          .FocusHolder.SetFocus
1260        End With
1270      End Select
1280    End If

EXITP:
1290    KeyCode = intRetVal
1300    Exit Sub

ERRH:
1310    intRetVal = 0
1320    Select Case ERR.Number
        Case Else
1330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1340    End Select
1350    Resume EXITP

End Sub

Private Sub Form_Timer()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1410    With Me
1420      .TimerInterval = 0&
1430      If gblnSetFocus = True Then
1440        gblnSetFocus = False
1450        strFormName = vbNullString
1460  On Error Resume Next
1470        Do Until strFormName = THIS_NAME
1480          DoCmd.SelectObject acForm, THIS_NAME, False
1490          strFormName = Screen.ActiveForm.Name
1500        Loop
1510  On Error GoTo ERRH
1520        DoCmd.Hourglass False
1530      End If
1540      If gstrReturningForm = "frmLocation_Asset" Then
1550        gstrReturningForm = vbNullString
1560        .cmdMoveAssets.SetFocus
1570        DoCmd.Hourglass False
1580      End If
1590      If gdblCrtRpt_CostTot <> 0# Then  ' ** Borrowing this Public variable.
1600        .frmLocations_Sub.Form.Requery
1610        DoEvents
1620        .frmLocations_Sub.Form.MoveRec 0, CLng(gdblCrtRpt_CostTot)  ' ** Form Procedure: frmLocations_Sub.
1630        gdblCrtRpt_CostTot = 0#
1640      End If
1650      If gblnGoToReport = True Then
1660        Select Case blnGoingToReport
            Case True
1670          DoCmd.Hourglass False  ' ** It should already be off.
1680          blnGoingToReport = False
1690          gblnGoToReport = False
1700          .GoToReport_arw_pr_img.Visible = False
1710          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1720          DoEvents
1730        Case False
1740          DoCmd.Hourglass True  ' ** Make sure it's still running.
1750          DoEvents
1760          blnGoingToReport = True
1770          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1780          .GoToReport_arw_pr_img.Visible = True
1790          DoEvents
1800          Beep
1810          DoCmd.Hourglass False
1820          DoEvents
1830          .TimerInterval = (GTR_WAIT * 2&)
1840        End Select
1850      End If
1860    End With

EXITP:
1870    Exit Sub

ERRH:
1880    Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Ignore.
1890    Case Else
1900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1910    End Select
1920    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim blnRetVal As Boolean

2010    Select Case blnWasGTR
        Case True
2020      blnWasGTR = False
2030    Case False
          ' ** Don't save settings imposed by GoToReport!
2040      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
2050    End Select

2060    If blnJustClose = False Then

2070      blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel

2080      Select Case strCallingForm
          Case "frmMenu_Utility"
2090        gblnSetFocus = True
2100        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2110      Case "frmJournal_Columns"
2120        gblnSetFocus = True
2130        DoCmd.SelectObject acForm, strCallingForm, False
2140        Forms(strCallingForm).cmdAssetNew.Enabled = True
2150        Forms(strCallingForm).cmdRecurNew.Enabled = True
2160        Forms(strCallingForm).TimerInterval = 250&
2170      End Select

2180    End If

EXITP:
2190    Exit Sub

ERRH:
2200    Select Case ERR.Number
        Case Else
2210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2220    End Select
2230    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2310    With Me
2320      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2330        Select Case blnPrintReport_Focus
            Case True
2340          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2350          .cmdPrintReport_raised_img.Visible = False
2360        Case False
2370          .cmdPrintReport_raised_img.Visible = True
2380          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2390        End Select
2400        .cmdPrintReport_raised_focus_img.Visible = False
2410        .cmdPrintReport_raised_focus_dots_img.Visible = False
2420        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2430        .cmdPrintReport_raised_img_dis.Visible = False
2440      End If
2450    End With

EXITP:
2460    Exit Sub

ERRH:
2470    Select Case ERR.Number
        Case Else
2480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2490    End Select
2500    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2710    DoCmd.Close acForm, THIS_NAME

EXITP:
2720    Exit Sub

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2820    intShiftDown = (Shift And acShiftMask) > 0
2830    intAltDown = (Shift And acAltMask) > 0
2840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyTab
2870        With Me
2880          intRetVal = 0
2890          lngRecsCur = .frmLocations_Sub.Form.RecCnt  ' ** Form Function: frmLocations_Sub.
2900          If lngRecsCur > 0& Then
2910            .frmLocations_Sub.SetFocus
2920            .frmLocations_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmLocations_Sub.
2930            .frmLocations_Sub.Form.Loc_Name.SetFocus
2940          Else
2950            .chkEnableCountry.SetFocus
2960          End If
2970        End With
2980      End Select
2990    End If

        ' ** Shift keys.
3000    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3010      Select Case intRetVal
          Case vbKeyTab
3020        With Me
3030          intRetVal = 0
3040          If .cmdDelete.Enabled = True Then
3050            .cmdDelete.SetFocus
3060          ElseIf .cmdAdd.Enabled = True Then
3070            .cmdAdd.SetFocus
3080          ElseIf .cmdMoveAssets.Enabled = True Then
3090            .cmdMoveAssets.SetFocus
3100          ElseIf .cmdPrintReport.Enabled = True Then
3110            .cmdPrintReport.SetFocus
3120          Else
3130            .chkEnableCountry.SetFocus
3140          End If
3150        End With
3160      End Select
3170    End If

EXITP:
3180    KeyCode = intRetVal
3190    Exit Sub

ERRH:
3200    intRetVal = 0
3210    Select Case ERR.Number
        Case Else
3220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3230    End Select
3240    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

3310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub chkEnableCountry_AfterUpdate()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkEnableCountry_AfterUpdate"

3410    With Me

3420      Select Case .chkEnableCountry
          Case True
3430        .chkEnableCountry_lbl.FontBold = True
3440        .frmLocations_Sub.Form.EnableCountry True  ' ** Form Procedure: frmLocations_Sub.
3450        DoEvents
3460        EnableCountry_SetFrmWidth  ' ** Procedure: Below
3470      Case False
3480        .chkEnableCountry_lbl.FontBold = False
3490        .frmLocations_Sub.Form.EnableCountry False  ' ** Form Procedure: frmLocations_Sub.
3500        DoEvents
3510        EnableCountry_SetFrmWidth  ' ** Procedure: Below
3520      End Select
3530      DoEvents
3540      lngRecsCur = .frmLocations_Sub.Form.RecCnt  ' ** Form Function: frmLocations_Sub.
3550      If lngRecsCur > 0& Then
3560        .frmLocations_Sub.SetFocus
3570      Else
3580        .cmdAdd.SetFocus
3590      End If

3600      If lngMonitorCnt > 1& And blnIsOpen = False Then
3610        EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
3620        lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
3630        If lngMonitorNum > 1& Then
3640          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
3650        End If
3660      End If

3670    End With

EXITP:
3680    Exit Sub

ERRH:
3690    Select Case ERR.Number
        Case Else
3700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3710    End Select
3720    Resume EXITP

End Sub

Private Sub chkEnableCountry_KeyDown(KeyCode As Integer, Shift As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkEnableCountry_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3820    intShiftDown = (Shift And acShiftMask) > 0
3830    intAltDown = (Shift And acAltMask) > 0
3840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3860      Select Case intRetVal
          Case vbKeyTab
3870        With Me
3880          intRetVal = 0
3890          If .cmdPrintReport.Enabled = True Then
3900            .cmdPrintReport.SetFocus
3910          ElseIf .cmdMoveAssets.Enabled = True Then
3920            .cmdMoveAssets.SetFocus
3930          ElseIf .cmdAdd.Enabled = True Then
3940            .cmdAdd.SetFocus
3950          ElseIf .cmdDelete.Enabled = True Then
3960            .cmdDelete.SetFocus
3970          Else
3980            .cmdClose.SetFocus
3990          End If
4000        End With
4010      End Select
4020    End If

        ' ** Shift keys.
4030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4040      Select Case intRetVal
          Case vbKeyTab
4050        With Me
4060          intRetVal = 0
4070          lngRecsCur = .frmLocations_Sub.Form.RecCnt  ' ** Form Function: frmLocations_Sub.
4080          If lngRecsCur > 0& Then
4090            .frmLocations_Sub.SetFocus
4100            .frmLocations_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmLocations_Sub.
4110            .frmLocations_Sub.Form.Loc_Name.SetFocus
4120          Else
4130            .cmdClose.SetFocus
4140          End If
4150        End With
4160      End Select
4170    End If

EXITP:
4180    KeyCode = intRetVal
4190    Exit Sub

ERRH:
4200    intRetVal = 0
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

4310    With Me
4320      blnPrintReport_Focus = True
4330      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
4340      .cmdPrintReport_raised_img.Visible = False
4350      .cmdPrintReport_raised_focus_img.Visible = False
4360      .cmdPrintReport_raised_focus_dots_img.Visible = False
4370      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4380      .cmdPrintReport_raised_img_dis.Visible = False
4390    End With

EXITP:
4400    Exit Sub

ERRH:
4410    Select Case ERR.Number
        Case Else
4420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4430    End Select
4440    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

4510    With Me
4520      blnPrintReport_MouseDown = True
4530      .cmdPrintReport_sunken_focus_dots_img.Visible = True
4540      .cmdPrintReport_raised_img.Visible = False
4550      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4560      .cmdPrintReport_raised_focus_img.Visible = False
4570      .cmdPrintReport_raised_focus_dots_img.Visible = False
4580      .cmdPrintReport_raised_img_dis.Visible = False
4590    End With

EXITP:
4600    Exit Sub

ERRH:
4610    Select Case ERR.Number
        Case Else
4620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4630    End Select
4640    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

4710    With Me
4720      strDocName = "rptListOfLocations"
4730      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4740        DoCmd.OpenReport strDocName, acViewPreview
4750        DoCmd.Maximize
4760        DoCmd.RunCommand acCmdFitToWindow
4770        .Visible = False
4780      Else
            '##GTR_Ref: rptListOfLocations
4790        DoCmd.OpenReport strDocName, acViewNormal
4800      End If
4810    End With

EXITP:
4820    Exit Sub

ERRH:
4830    Me.Visible = True
4840    DoCmd.Restore
4850    Select Case ERR.Number
        Case Else
4860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4870    End Select
4880    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4920    intShiftDown = (Shift And acShiftMask) > 0
4930    intAltDown = (Shift And acAltMask) > 0
4940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4960      Select Case intRetVal
          Case vbKeyTab
4970        With Me
4980          intRetVal = 0
4990          If .cmdMoveAssets.Enabled = True Then
5000            .cmdMoveAssets.SetFocus
5010          ElseIf .cmdAdd.Enabled = True Then
5020            .cmdAdd.SetFocus
5030          ElseIf .cmdDelete.Enabled = True Then
5040            .cmdDelete.SetFocus
5050          Else
5060            .cmdClose.SetFocus
5070          End If
5080        End With
5090      End Select
5100    End If

        ' ** Shift keys.
5110    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5120      Select Case intRetVal
          Case vbKeyTab
5130        With Me
5140          intRetVal = 0
5150          .chkEnableCountry.SetFocus
5160        End With
5170      End Select
5180    End If

EXITP:
5190    KeyCode = intRetVal
5200    Exit Sub

ERRH:
5210    intRetVal = 0
5220    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
5230    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5240    Case Else
5250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5260    End Select
5270    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

5310    With Me
5320      If blnPrintReport_MouseDown = False Then
5330        Select Case blnPrintReport_Focus
            Case True
5340          .cmdPrintReport_raised_focus_dots_img.Visible = True
5350          .cmdPrintReport_raised_focus_img.Visible = False
5360        Case False
5370          .cmdPrintReport_raised_focus_img.Visible = True
5380          .cmdPrintReport_raised_focus_dots_img.Visible = False
5390        End Select
5400        .cmdPrintReport_raised_img.Visible = False
5410        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5420        .cmdPrintReport_sunken_focus_dots_img.Visible = False
5430        .cmdPrintReport_raised_img_dis.Visible = False
5440      End If
5450    End With

EXITP:
5460    Exit Sub

ERRH:
5470    Select Case ERR.Number
        Case Else
5480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5490    End Select
5500    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

5610    With Me
5620      .cmdPrintReport_raised_focus_dots_img.Visible = True
5630      .cmdPrintReport_raised_img.Visible = False
5640      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5650      .cmdPrintReport_raised_focus_img.Visible = False
5660      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5670      .cmdPrintReport_raised_img_dis.Visible = False
5680      blnPrintReport_MouseDown = False
5690    End With

EXITP:
5700    Exit Sub

ERRH:
5710    Select Case ERR.Number
        Case Else
5720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5730    End Select
5740    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

5810    With Me
5820      .cmdPrintReport_raised_img.Visible = True
5830      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5840      .cmdPrintReport_raised_focus_img.Visible = False
5850      .cmdPrintReport_raised_focus_dots_img.Visible = False
5860      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5870      .cmdPrintReport_raised_img_dis.Visible = False
5880      blnPrintReport_Focus = False
5890    End With

EXITP:
5900    Exit Sub

ERRH:
5910    Select Case ERR.Number
        Case Else
5920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5930    End Select
5940    Resume EXITP

End Sub

Private Sub cmdMoveAssets_Click()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveAssets_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String
        Dim blnContinue As Boolean

6010    blnContinue = True

6020    DoCmd.Hourglass True
6030    DoEvents

6040    Set dbs = CurrentDb
6050    With dbs
          ' ** ActiveAssets, linked to Location, MasterAsset.
6060      Set qdf = .QueryDefs("qryLocation_Asset_Sub_02")
6070      Set rst = qdf.OpenRecordset
6080      If rst.BOF = True And rst.EOF = True Then
6090        blnContinue = False
6100        DoCmd.Hourglass False
6110        MsgBox "There are no assets held by accounts.", vbInformation + vbOKOnly, "Nothing To Do"
6120      End If
6130      rst.Close
6140      Set rst = Nothing
6150      Set qdf = Nothing
6160      If blnContinue = True Then
            ' ** Empty tmpLocation_Asset.
6170        Set qdf = .QueryDefs("qryLocation_Asset_Sub_01")
6180        qdf.Execute
6190        Set qdf = Nothing
6200        DoEvents
            ' ** Append qryLocation_Asset_Sub_02(ActiveAssets,
            ' ** linked to Location, MasterAsset) to tmpLocation_Asset.
6210        Set qdf = .QueryDefs("qryLocation_Asset_Sub_03")
6220        qdf.Execute
6230        Set qdf = Nothing
6240        DoEvents
            ' ** Empty tmpAccount.
6250        Set qdf = .QueryDefs("qryLocation_Asset_03_01")
6260        qdf.Execute
6270        Set qdf = Nothing
6280        DoEvents
            ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
            ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
            ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
            ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
            ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
            ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
            ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
            ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
6290        Set qdf = .QueryDefs("qryLocation_Asset_03_02")
6300        qdf.Execute
6310        Set qdf = Nothing
6320        DoEvents
6330      End If
6340      .Close
6350    End With
6360    Set dbs = Nothing
6370    DoEvents

6380    If blnContinue = True Then
6390      strDocName = "frmLocation_Asset"
6400      gblnSetFocus = True
6410      blnJustClose = True
6420      DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6430      DoCmd.Close acForm, THIS_NAME
6440    End If

EXITP:
6450    Set rst = Nothing
6460    Set qdf = Nothing
6470    Set dbs = Nothing
6480    Exit Sub

ERRH:
6490    DoCmd.Hourglass False
6500    Select Case ERR.Number
        Case Else
6510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6520    End Select
6530    Resume EXITP

End Sub

Private Sub cmdMoveAssets_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6620    intShiftDown = (Shift And acShiftMask) > 0
6630    intAltDown = (Shift And acAltMask) > 0
6640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6660      Select Case intRetVal
          Case vbKeyTab
6670        With Me
6680          intRetVal = 0
6690          If .cmdAdd.Enabled = True Then
6700            .cmdAdd.SetFocus
6710          ElseIf .cmdDelete.Enabled = True Then
6720            .cmdDelete.SetFocus
6730          Else
6740            .cmdClose.SetFocus
6750          End If
6760        End With
6770      End Select
6780    End If

        ' ** Shift keys.
6790    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6800      Select Case intRetVal
          Case vbKeyTab
6810        With Me
6820          intRetVal = 0
6830          If .cmdPrintReport.Enabled = True Then
6840            .cmdPrintReport.SetFocus
6850          Else
6860            .chkEnableCountry.SetFocus
6870          End If
6880        End With
6890      End Select
6900    End If

EXITP:
6910    KeyCode = intRetVal
6920    Exit Sub

ERRH:
6930    intRetVal = 0
6940    Select Case ERR.Number
        Case Else
6950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6960    End Select
6970    Resume EXITP

End Sub

Private Sub cmdAdd_Click()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

7010    DoCmd.OpenForm "frmLocations_Add", acNormal, , , acFormAdd, acDialog, THIS_NAME

EXITP:
7020    Exit Sub

ERRH:
7030    Select Case ERR.Number
        Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_KeyDown"

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
          Case vbKeyTab
7170        With Me
7180          intRetVal = 0
7190          If .cmdDelete.Enabled = True Then
7200            .cmdDelete.SetFocus
7210          Else
7220            .cmdClose.SetFocus
7230          End If
7240        End With
7250      End Select
7260    End If

        ' ** Shift keys.
7270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7280      Select Case intRetVal
          Case vbKeyTab
7290        With Me
7300          intRetVal = 0
7310          If .cmdMoveAssets.Enabled = True Then
7320            .cmdMoveAssets.SetFocus
7330          ElseIf .cmdPrintReport.Enabled = True Then
7340            .cmdPrintReport.SetFocus
7350          Else
7360            .chkEnableCountry.SetFocus
7370          End If
7380        End With
7390      End Select
7400    End If

EXITP:
7410    KeyCode = intRetVal
7420    Exit Sub

ERRH:
7430    intRetVal = 0
7440    Select Case ERR.Number
        Case Else
7450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7460    End Select
7470    Resume EXITP

End Sub

Private Sub cmdDelete_Click()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngLocID As Long
        Dim strAssets As String
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnContinue As Boolean

7510    blnContinue = True

7520    With Me

7530      Set dbs = CurrentDb
          ' ** Location, without '{Unassigned}'.
7540      Set qdf = dbs.QueryDefs("qryLocation_06")
7550      Set rst = qdf.OpenRecordset
7560      If rst.BOF = True And rst.EOF = True Then
7570        blnContinue = False
7580        MsgBox "No Locations to delete.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
7590        rst.Close
7600        dbs.Close
7610      Else
7620        If .frmLocations_Sub.Form.NewRecord Then
7630          blnContinue = False
7640          If rst.BOF = True And rst.EOF = True Then
                ' ** There are no other locations.
7650          Else
7660            .frmLocations_Sub.Form.MoveRec acCmdRecordsGoToFirst
7670          End If
7680          rst.Close
7690          dbs.Close
7700        ElseIf .frmLocations_Sub.Form.Location_ID = 1& And .frmLocations_Sub.Form.Loc_Name = "{Unassigned}" Then
7710          blnContinue = False
7720          MsgBox "You cannot delete the '{Unassigned}' location.", vbInformation + vbOKOnly, "Invalid Action"
7730        Else
7740          rst.Close
7750          lngLocID = Me.frmLocations_Sub.Form.Location_ID
              ' ** ActiveAssets, linked to MasterAsset, grouped by assetno, by specified [locid].
7760          Set qdf = dbs.QueryDefs("qryLocation_07")
7770          With qdf.Parameters
7780            ![locid] = lngLocID
7790          End With
7800          Set rst = qdf.OpenRecordset
7810          If rst.BOF = True And rst.EOF = True Then
7820            rst.Close
7830            If MsgBox("Are you sure you want to delete '" & .frmLocations_Sub.Form.Loc_Name & "'?", vbQuestion + vbYesNo, "Delete Location") = vbYes Then
7840              .frmLocations_Sub.Form.DeleteRec lngLocID  ' ** Form Procedure: frmLocations_Sub.
7850            End If
7860          Else
7870            strAssets = vbNullString
7880            rst.MoveLast
7890            lngRecs = rst.RecordCount
7900            rst.MoveFirst
7910            For lngX = 1& To lngRecs
7920              strAssets = strAssets & "'" & rst![description_ma] & "', "
7930              If lngX < lngRecs Then rst.MoveNext
7940            Next
7950            rst.Close
7960            strAssets = Trim(strAssets)
7970            If Right(strAssets, 1) = "," Then strAssets = Left(strAssets, (Len(strAssets) - 1))
7980            MsgBox "This Location cannot be deleted because the following asset(s) are assigned to it:" & vbCrLf & vbCrLf & _
                  strAssets & vbCrLf & vbCrLf & _
                  "If you wish to delete this location, you will first need to move the assets to another location.", _
                  vbInformation + vbOKOnly, "Invalid Action"
7990          End If
8000          dbs.Close
8010        End If
8020      End If

8030    End With

EXITP:
8040    Set rst = Nothing
8050    Set qdf = Nothing
8060    Set dbs = Nothing
8070    Exit Sub

ERRH:
8080    Select Case ERR.Number
        Case Else
8090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8100    End Select
8110    Resume EXITP

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_KeyDown"

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
          Case vbKeyTab
8270        With Me
8280          intRetVal = 0
8290          .cmdClose.SetFocus
8300        End With
8310      End Select
8320    End If

        ' ** Shift keys.
8330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8340      Select Case intRetVal
          Case vbKeyTab
8350        With Me
8360          intRetVal = 0
8370          If .cmdAdd.Enabled = True Then
8380            .cmdAdd.SetFocus
8390          ElseIf .cmdMoveAssets.Enabled = True Then
8400            .cmdMoveAssets.SetFocus
8410          ElseIf .cmdPrintReport.Enabled = True Then
8420            .cmdPrintReport.SetFocus
8430          Else
8440            .chkEnableCountry.SetFocus
8450          End If
8460        End With
8470      End Select
8480    End If

EXITP:
8490    KeyCode = intRetVal
8500    Exit Sub

ERRH:
8510    intRetVal = 0
8520    Select Case ERR.Number
        Case Else
8530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8540    End Select
8550    Resume EXITP

End Sub

Private Sub EnableCountry_SetFrmWidth()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "EnableCountry_SetFrmWidth"

        Dim lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
        Dim lngTmp01 As Long, lngTmp02 As Long

8610    With Me

          ' ** Variables are fed empty, then populated ByRef.
8620      GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

8630      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
8640      lngMonitorNum = 1&: lngTmp02 = 0&
8650      EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
8660      If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

8670      Select Case .chkEnableCountry
          Case True
            ' ** Design View set at USA-Only width.
8680        If lngFrm_Width <> (lngFrm_Width_Orig + lngCountry_Diff) Then
8690          .Width = (lngForm_Width + lngCountry_Diff)
8700          lngTmp01 = (lngFrm_Left_Orig - (lngCountry_Diff / 2))
8710          If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top_Orig
8720          DoCmd.MoveSize lngTmp01, lngTmp02, (lngFrm_Width_Orig + lngCountry_Diff), lngFrm_Height_Orig  'lngFrm_Top_Orig
8730          If lngMonitorNum > 1& Then
8740            LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
8750          End If
8760          .frmLocations_Sub.Width = (lngSub_Width + lngCountry_Diff)
8770          .frmLocations_Sub_box.Width = (.frmLocations_Sub.Width + (2& * lngTpp))
8780          .Nav_box01.Width = .frmLocations_Sub.Width
8790          .cmdClose.Left = (lngClose_Left + lngCountry_Diff)
8800          .cmdDelete.Left = (lngDelete_Left + lngCountry_Diff)
8810          .cmdAdd.Left = (lngAdd_Left + lngCountry_Diff)
8820          .cmdMoveAssets.Left = (lngMoveAssets_Left + lngCountry_Diff)
8830          .Header_vline01.Left = .Width
8840          .Header_vline02.Left = .Width
8850          .Detail_vline01.Left = .Width
8860          .Detail_vline02.Left = .Width
8870          .Footer_vline01.Left = .Width
8880          .Footer_vline02.Left = .Width
8890          .Header_hline01.Width = .Width
8900          .Header_hline02.Width = .Width
8910          .Detail_hline01.Width = .Width
8920          .Detail_hline02.Width = .Width
8930          .Footer_hline01.Width = .Width
8940          .Footer_hline02.Width = .Width
8950        End If
8960      Case False
8970        If lngFrm_Width <> lngFrm_Width_Orig Then
8980          .frmLocations_Sub.Width = lngSub_Width
8990          .frmLocations_Sub_box.Width = (.frmLocations_Sub.Width + (2& * lngTpp))
9000          .Nav_box01.Width = .frmLocations_Sub.Width
9010          .cmdClose.Left = lngClose_Left
9020          .cmdMoveAssets.Left = lngMoveAssets_Left
9030          .cmdDelete.Left = lngDelete_Left
9040          .cmdAdd.Left = lngAdd_Left
9050          .Header_vline01.Left = lngForm_Width
9060          .Header_vline02.Left = lngForm_Width
9070          .Detail_vline01.Left = lngForm_Width
9080          .Detail_vline02.Left = lngForm_Width
9090          .Footer_vline01.Left = lngForm_Width
9100          .Footer_vline02.Left = lngForm_Width
9110          .Header_hline01.Width = lngForm_Width
9120          .Header_hline02.Width = lngForm_Width
9130          .Detail_hline01.Width = lngForm_Width
9140          .Detail_hline02.Width = lngForm_Width
9150          .Footer_hline01.Width = lngForm_Width
9160          .Footer_hline02.Width = lngForm_Width
9170          .Width = lngForm_Width
9180          If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top_Orig
9190          DoCmd.MoveSize lngFrm_Left_Orig, lngTmp02, lngFrm_Width_Orig, lngFrm_Height_Orig  'lngFrm_Top_Orig
9200          If lngMonitorNum > 1& Then
9210            LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
9220          End If
9230        End If

9240      End Select
9250    End With

EXITP:
9260    Exit Sub

ERRH:
9270    Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Confusion.
9280    Case Else
9290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9300    End Select
9310    Resume EXITP

End Sub
