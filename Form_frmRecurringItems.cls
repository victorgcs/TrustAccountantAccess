VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRecurringItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRecurringItems"

'VGC 04/16/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   EnableCountry_SetFrmWidth()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd}
' **   Close:            C {cmdClose}
' **   Delete:           D {cmdDelete}
' **   Enable Country:   E {chkEnableCountry}
' **   Form 1099:        F {RecurringItemAux_1099 on frmRecurringItems_Sub}
' **   Paid:             I {opgType_optPaid}
' **   All:              L {opgType_optAll}
' **   Miscellaneous:    M {opgType_optMisc}
' **   Description:      N {Recur_Name on frmRecurringItems_Sub}
' **   Print Report:     P {cmdPrintReport}
' **   Received:         R {opgType_optReceived}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmRecurringItems_Sub}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private lngFrm_Left_Orig As Long, lngFrm_Top_Orig As Long, lngFrm_Width_Orig As Long, lngFrm_Height_Orig As Long
Private lngForm_Width As Long, lngCountry_Diff As Long, lngSub_Width As Long
Private lngClose_Left As Long, lngDelete_Left As Long, lngAdd_Left As Long
Private lngEnableCountry_Left As Long, lngEnableCountryLbl_Left As Long

Private blnGoingToReport As Boolean, lngPrintReportLbl_Left As Long, blnResetHeaders As Boolean
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

160           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

170           blnResetHeaders = False
180           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

190           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
200           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

              ' ** Variables are fed empty, then populated ByRef.
210           GetFormDimensions Me, lngFrm_Left_Orig, lngFrm_Top_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig  ' ** Module Function: modWindowFunctions.

220           FrmDimSave THIS_NAME, lngFrm_Left_Orig, lngFrm_Top_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig  ' ** Module Function: modMonitorFuncs.

230           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
240           lngMonitorNum = 1&: lngTmp01 = 0&
250           EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
260           If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

270           If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top_Orig
280           DoCmd.MoveSize lngFrm_Left_Orig, lngTmp01, lngFrm_Width_Orig, lngFrm_Height_Orig  'lngFrm_Top_Orig
290           If lngMonitorNum > 1& Then
300             LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
310           End If

320           lngForm_Width = .Width
330           lngSub_Width = .frmRecurringItems_Sub.Width
340           lngClose_Left = .cmdClose.Left
350           lngDelete_Left = .cmdDelete.Left
360           lngAdd_Left = .cmdAdd.Left
370           lngEnableCountry_Left = .chkEnableCountry.Left
380           lngEnableCountryLbl_Left = .chkEnableCountry_lbl.Left

390           SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.
              'SystemColor_Move Me  ' ** Module Procedure: modSystemColorFuncs.
400           StateCodeQrySet Me  ' ** Module Function: modPreferenceFuncs.

410           .opgType = 0

420           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

430           If .opgType = 0 Then
440             With .frmRecurringItems_Sub.Form
450               .FilterOn = True
460               .Filter = "[Recur_Type] = 'Misc'"
470               .Refresh
480             End With
490             .opgType = .opgType_optMisc.OptionValue
500           End If

510           lngCountry_Diff = (.frmRecurringItems_Sub.Form.RecurringItemAux_1099_lbl.Left - .frmRecurringItems_Sub.Form.Recur_Country.Left)

520           opgType_AfterUpdate  ' ** Procedure: Below.
530           chkEnableCountry_AfterUpdate  ' ** Procedure: Below.

540           .frmRecurringItems_Sub.Form.EnableAddressInfo  ' ** Procedure: Below.

550           lngRecsCur = .frmRecurringItems_Sub.Form.RecCnt  ' ** Form Function: frmRecurringItems_Sub.
560           If lngRecsCur > 0& Then
570             .frmRecurringItems_Sub.SetFocus
580             .frmRecurringItems_Sub.Form.FocusHolder.SetFocus
590           Else
600             .cmdAdd.SetFocus
610           End If

620           blnIsOpen = False

630         Else
640           Cancel = -1
650         End If
660       Else
670         Cancel = -1
680       End If
690     End With

700     If Cancel = -1 Then
710       Beep
720       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
730     End If

EXITP:
740     Exit Sub

ERRH:
750     Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Empty.
760     Case Else
770       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
780     End Select
790     Resume EXITP

End Sub

Private Sub Form_Load()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

810     With Me
820       If IsNull(.Recur_State_Pref) = False Then
830         .frmRecurringItems_Sub.Form.Recur_State.DefaultValue = Chr(34) & .Recur_State_Pref & Chr(34)
840       End If
850     End With

EXITP:
860     Exit Sub

ERRH:
870     Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Empty.
880     Case Else
890       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
900     End Select
910     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1020    intShiftDown = (Shift And acShiftMask) > 0
1030    intAltDown = (Shift And acAltMask) > 0
1040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Form 1099:        F {RecurringItemAux_1099 on frmRecurringItems_Sub}
        ' **   Description:      N {Recur_Name on frmRecurringItems_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmRecurringItems_Sub}

        ' ** Plain keys.
1050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1060      Select Case intRetVal
          Case vbKeyEscape
1070        intRetVal = 0
1080        cmdClose_Click  ' ** Procedure: Below.
1090      End Select
1100    End If

        ' ** Alt keys.
1110    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1120      Select Case intRetVal
          Case vbKeyF
1130        With Me
1140          intRetVal = 0
1150          lngRecsCur = .frmRecurringItems_Sub.Form.RecCnt  ' ** Form Function: frmRecurringItems_Sub.
1160          If lngRecsCur > 0& Then
1170            If .frmRecurringItems_Sub.Form.RecurringItemAux_1099.Enabled = True Then
1180              .frmRecurringItems_Sub.SetFocus
1190              .frmRecurringItems_Sub.Form.RecurringItemAux_1099.SetFocus
1200              .frmRecurringItems_Sub.Form.RecurringItemAux_1099 = (Not .frmRecurringItems_Sub.Form.RecurringItemAux_1099)
1210              .frmRecurringItems_Sub.Form.RecurringItemAux_1099_AfterUpdate  ' ** Form Procedure: frmRecurringItems_Sub.
1220            Else
1230              Beep
1240            End If
1250          Else
1260            Beep
1270          End If
1280        End With
1290      Case vbKeyN
1300        With Me
1310          intRetVal = 0
1320          .frmRecurringItems_Sub.SetFocus
1330          .frmRecurringItems_Sub.Form.Recur_Name.SetFocus
1340        End With
1350      Case vbKeyX
1360        intRetVal = 0
1370        cmdClose_Click  ' ** Procedure: Below.
1380      End Select
1390    End If

        ' ** Ctrl keys.
1400    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1410      Select Case intRetVal
          Case vbKeyS
1420        With Me
1430          intRetVal = 0
1440          lngRecsCur = .frmRecurringItems_Sub.Form.RecCnt  ' ** Form Function: frmRecurringItems_Sub.
1450          If lngRecsCur > 0& Then
1460            .frmRecurringItems_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmRecurringItems_Sub.
1470          Else
1480            Beep
1490          End If
1500        End With
1510      End Select
1520    End If

        ' ** Ctrl-Shift keys.
1530    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1540      Select Case intRetVal
          Case vbKeyF
1550        With Me
1560          intRetVal = 0
1570          .FocusHolder.SetFocus
1580        End With
1590      End Select
1600    End If

EXITP:
1610    KeyCode = intRetVal
1620    Exit Sub

ERRH:
1630    intRetVal = 0
1640    Select Case ERR.Number
        Case Else
1650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1660    End Select
1670    Resume EXITP

End Sub

Private Sub Form_Timer()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1710    With Me
1720      .TimerInterval = 0&
1730      If gblnSetFocus = True Then
1740        gblnSetFocus = False
1750        strFormName = vbNullString
1760  On Error Resume Next
1770        Do Until strFormName = THIS_NAME
1780          DoCmd.SelectObject acForm, THIS_NAME, False
1790          strFormName = Screen.ActiveForm.Name
1800        Loop
1810  On Error GoTo ERRH
1820        CmdBars_Hide True  ' ** Module Procedure: modWindowFunctions.
1830        If blnResetHeaders = True Then
1840          .TimerInterval = 250&
1850        End If
1860      ElseIf blnResetHeaders = True Then
1870        blnResetHeaders = False
1880        Select Case .frmRecurringItems_Sub.Form.Recur_Type
            Case "Misc"
1890          .frmRecurringItems_Sub.Form.RecurHeaders False  ' ** Form Procedure: frmRecurringItems_Sub.
1900        Case "Payee"
1910          .frmRecurringItems_Sub.Form.Recur_Address.Enabled = True
1920          .frmRecurringItems_Sub.Form.Recur_City.Enabled = True
1930          .frmRecurringItems_Sub.Form.Recur_State.Enabled = True
1940          .frmRecurringItems_Sub.Form.Recur_Zip.Enabled = True
1950          If .chkEnableCountry = True Then
1960            .frmRecurringItems_Sub.Form.Recur_Country.Enabled = True
1970            .frmRecurringItems_Sub.Form.Recur_PostalCode.Enabled = True
1980          End If
1990          .frmRecurringItems_Sub.Form.RecurringItemAux_1099.Enabled = True
2000          .frmRecurringItems_Sub.Form.RecurringItemAux_TaxNo.Enabled = True
2010          .frmRecurringItems_Sub.Form.RecurHeaders True  ' ** Form Procedure: frmRecurringItems_Sub.
2020        Case "Payor"
2030          .frmRecurringItems_Sub.Form.RecurHeaders False  ' ** Form Procedure: frmRecurringItems_Sub.
2040        End Select
100       .FocusHolder.SetFocus
110       DoEvents
2050      End If
2060      If gdblCrtRpt_CostTot <> 0# Then  ' ** Borrowing this Public variable.
2070        Application.Echo False
2080        .frmRecurringItems_Sub.Form.FilterOn = False
2090        .frmRecurringItems_Sub.Form.Requery
2100        DoEvents
2110        .frmRecurringItems_Sub.Form.FilterOn = True
2120        Application.Echo True
2130        .frmRecurringItems_Sub.Form.MoveRec 0, CLng(gdblCrtRpt_CostTot)  ' ** Form Procedure: frmRecurringItems_Sub.
2140        .frmRecurringItems_Sub.Form.Recur_Name.ForeColor = CLR_BLK
2150        gdblCrtRpt_CostTot = 0#
2160      End If
2170      If gblnGoToReport = True Then
2180        Select Case blnGoingToReport
            Case True
2190          DoCmd.Hourglass False  ' ** It should already be off.
2200          blnGoingToReport = False
2210          gblnGoToReport = False
2220          .GoToReport_arw_pr_img.Visible = False
2230          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
2240          DoEvents
2250        Case False
2260          DoCmd.Hourglass True  ' ** Make sure it's still running.
2270          DoEvents
2280          blnGoingToReport = True
2290          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
2300          .GoToReport_arw_pr_img.Visible = True
2310          .cmdPrintReport.SetFocus
2320          DoEvents
2330          Beep
2340          DoCmd.Hourglass False
2350          DoEvents
2360          .TimerInterval = (GTR_WAIT * 2&)
2370        End Select
2380      Else
2390        DoCmd.Hourglass False
2400      End If
2410    End With

EXITP:
2420    Exit Sub

ERRH:
2430    Application.Echo True
2440    Select Case ERR.Number
        Case Else
2450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2460    End Select
2470    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim blnRetVal As Boolean

2510    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2520    blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.

2530    Select Case strCallingForm
        Case "frmMenu_Utility"
2540      gblnSetFocus = True
2550      gstrReturningForm = THIS_NAME
2560      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2570    Case "frmJournal_Columns"
2580      gblnSetFocus = True
2590      DoCmd.SelectObject acForm, strCallingForm, False
2600      Forms(strCallingForm).cmdAssetNew.Enabled = True
2610      Forms(strCallingForm).cmdLocNew.Enabled = True
2620      Forms(strCallingForm).TimerInterval = 250&
2630    End Select

EXITP:
2640    Exit Sub

ERRH:
2650    Select Case ERR.Number
        Case Else
2660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2670    End Select
2680    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2710    With Me
2720      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2730        Select Case blnPrintReport_Focus
            Case True
2740          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2750          .cmdPrintReport_raised_img.Visible = False
2760        Case False
2770          .cmdPrintReport_raised_img.Visible = True
2780          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2790        End Select
2800        .cmdPrintReport_raised_focus_img.Visible = False
2810        .cmdPrintReport_raised_focus_dots_img.Visible = False
2820        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2830        .cmdPrintReport_raised_img_dis.Visible = False
2840      End If
2850    End With

EXITP:
2860    Exit Sub

ERRH:
2870    Select Case ERR.Number
        Case Else
2880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2890    End Select
2900    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Public Sub cmdClose_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3110    DoCmd.Close acForm, THIS_NAME

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3220    intShiftDown = (Shift And acShiftMask) > 0
3230    intAltDown = (Shift And acAltMask) > 0
3240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3260      Select Case intRetVal
          Case vbKeyTab
3270        With Me
3280          intRetVal = 0
3290          lngRecsCur = .frmRecurringItems_Sub.Form.RecCnt  ' ** Form Function: frmRecurringItems_Sub.
3300          If lngRecsCur > 0& Then
3310            .frmRecurringItems_Sub.SetFocus
3320            .frmRecurringItems_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Function: frmRecurringItems_Sub.
3330            .frmRecurringItems_Sub.Form.Recur_Name.SetFocus
3340          ElseIf .cmdPrintReport.Enabled = True Then
3350            .cmdPrintReport.SetFocus
3360          Else
3370            .opgType.SetFocus
3380          End If
3390        End With
3400      End Select
3410    End If

        ' ** Shift keys.
3420    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3430      Select Case intRetVal
          Case vbKeyTab
3440        With Me
3450          intRetVal = 0
3460          If .cmdDelete.Enabled = True Then
3470            .cmdDelete.SetFocus
3480          ElseIf .cmdAdd.Enabled = True Then
3490            .cmdAdd.SetFocus
3500          Else
3510            .opgType.SetFocus
3520          End If
3530        End With
3540      End Select
3550    End If

EXITP:
3560    KeyCode = intRetVal
3570    Exit Sub

ERRH:
3580    intRetVal = 0
3590    Select Case ERR.Number
        Case Else
3600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3610    End Select
3620    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

3710    With Me
3720      blnPrintReport_Focus = True
3730      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3740      .cmdPrintReport_raised_img.Visible = False
3750      .cmdPrintReport_raised_focus_img.Visible = False
3760      .cmdPrintReport_raised_focus_dots_img.Visible = False
3770      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3780      .cmdPrintReport_raised_img_dis.Visible = False
3790    End With

EXITP:
3800    Exit Sub

ERRH:
3810    Select Case ERR.Number
        Case Else
3820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3830    End Select
3840    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

3910    With Me
3920      blnPrintReport_MouseDown = True
3930      .cmdPrintReport_sunken_focus_dots_img.Visible = True
3940      .cmdPrintReport_raised_img.Visible = False
3950      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3960      .cmdPrintReport_raised_focus_img.Visible = False
3970      .cmdPrintReport_raised_focus_dots_img.Visible = False
3980      .cmdPrintReport_raised_img_dis.Visible = False
3990    End With

EXITP:
4000    Exit Sub

ERRH:
4010    Select Case ERR.Number
        Case Else
4020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4030    End Select
4040    Resume EXITP

End Sub

Public Sub cmdPrintReport_Click()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String, strSortNow As String

4110    With Me
4120      strSortNow = .frmRecurringItems_Sub.Form.SortNow_Get  ' ** Form Function: frmRecurringItems_Sub.
4130      If Left(strSortNow, 12) = "[Recur_Type]" Then
4140        strDocName = "rptListOfRecurringItems_01"
4150      Else
4160        strDocName = "rptListOfRecurringItems_02"
4170      End If
4180      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4190        DoCmd.OpenReport strDocName, acViewPreview
4200        DoCmd.Maximize
4210        DoCmd.RunCommand acCmdFitToWindow
4220        .Visible = False
4230      Else
            '##GTR_Ref: rptListOfRecurringItems_01
4240        DoCmd.OpenReport strDocName, acViewNormal
4250      End If
4260    End With

EXITP:
4270    Exit Sub

ERRH:
4280    Me.Visible = True
4290    DoCmd.Restore
4300    Select Case ERR.Number
        Case Else
4310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4320    End Select
4330    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4420    intShiftDown = (Shift And acShiftMask) > 0
4430    intAltDown = (Shift And acAltMask) > 0
4440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4460      Select Case intRetVal
          Case vbKeyTab
4470        With Me
4480          intRetVal = 0
4490          .opgType.SetFocus
4500        End With
4510      End Select
4520    End If

        ' ** Shift keys.
4530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4540      Select Case intRetVal
          Case vbKeyTab
4550        With Me
4560          intRetVal = 0
4570          lngRecsCur = .frmRecurringItems_Sub.Form.RecCnt  ' ** Form Function: frmRecurringItems_Sub.
4580          If lngRecsCur > 0& Then
4590            .frmRecurringItems_Sub.SetFocus
4600            .frmRecurringItems_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmRecurringItems_Sub.
4610          Else
4620            .cmdClose.SetFocus
4630          End If
4640        End With
4650      End Select
4660    End If

EXITP:
4670    KeyCode = intRetVal
4680    Exit Sub

ERRH:
4690    intRetVal = 0
4700    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
4710    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4720    Case Else
4730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4740    End Select
4750    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

4810    With Me
4820      If blnPrintReport_MouseDown = False Then
4830        Select Case blnPrintReport_Focus
            Case True
4840          .cmdPrintReport_raised_focus_dots_img.Visible = True
4850          .cmdPrintReport_raised_focus_img.Visible = False
4860        Case False
4870          .cmdPrintReport_raised_focus_img.Visible = True
4880          .cmdPrintReport_raised_focus_dots_img.Visible = False
4890        End Select
4900        .cmdPrintReport_raised_img.Visible = False
4910        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4920        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4930        .cmdPrintReport_raised_img_dis.Visible = False
4940      End If
4950    End With

EXITP:
4960    Exit Sub

ERRH:
4970    Select Case ERR.Number
        Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

5110    With Me
5120      .cmdPrintReport_raised_focus_dots_img.Visible = True
5130      .cmdPrintReport_raised_img.Visible = False
5140      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5150      .cmdPrintReport_raised_focus_img.Visible = False
5160      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5170      .cmdPrintReport_raised_img_dis.Visible = False
5180      blnPrintReport_MouseDown = False
5190    End With

EXITP:
5200    Exit Sub

ERRH:
5210    Select Case ERR.Number
        Case Else
5220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5230    End Select
5240    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

5310    With Me
5320      .cmdPrintReport_raised_img.Visible = True
5330      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5340      .cmdPrintReport_raised_focus_img.Visible = False
5350      .cmdPrintReport_raised_focus_dots_img.Visible = False
5360      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5370      .cmdPrintReport_raised_img_dis.Visible = False
5380      blnPrintReport_Focus = False
5390    End With

EXITP:
5400    Exit Sub

ERRH:
5410    Select Case ERR.Number
        Case Else
5420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5430    End Select
5440    Resume EXITP

End Sub

Private Sub cmdAdd_Click()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

        Dim frm As Access.Form
        Dim blnTmp01 As Boolean

5510    With Me

5520      DoCmd.OpenForm "frmRecurringItems_Add", acNormal, , , , , THIS_NAME
          'DoCmd.GoToRecord acDataForm, "frmRecurringItems_Add", acNewRec

5530      blnTmp01 = .chkEnableCountry

5540      Set frm = Forms("frmRecurringItems_Add")

5550      Select Case .opgType
          Case .opgType_optMisc.OptionValue
5560        With frm
5570          .Recur_Type = "Misc"
5580          .Caption = "Add New Recurring Item"
5590          .Recur_lbl.Caption = "Recurring Item"
5600          .Header_lbl_Item.Visible = True
5610          .Header_lbl_Payment.Visible = False
5620          .Header_lbl_Receipt.Visible = False
5630          .Recur_Address.Visible = False
5640          .Recur_City.Visible = False
5650          .Recur_State.Visible = False
5660          .Recur_Zip.Visible = False
5670          Select Case blnTmp01
              Case True
5680            .opgLoc.Visible = False
5690            .opgLoc_box.Visible = False
5700            .Recur_Country.Visible = False
5710            .Recur_PostalCode.Visible = False
5720          Case False
5730            .Recur_Country.Visible = False
5740            .Recur_PostalCode.Visible = False
5750            .opgLoc.Visible = False
5760            .opgLoc_box.Visible = False
5770          End Select
5780          .RecurringItemAux_1099.Visible = False
5790          .RecurringItemAux_TaxNo.Visible = False
5800        End With
5810      Case .opgType_optPaid.OptionValue
5820        With frm
5830          .Recur_Type = "Payee"
5840          .Caption = "Add New Recurring Payment"
5850          .Recur_lbl.Caption = "Recurrent Payment"
5860          .Header_lbl_Item.Visible = False
5870          .Header_lbl_Payment.Visible = True
5880          .Header_lbl_Receipt.Visible = False
5890          .Recur_Address.Visible = True
5900          .Recur_City.Visible = True
5910          .Recur_State.Visible = True
5920          .Recur_Zip.Visible = True
5930          Select Case blnTmp01
              Case True
5940            .opgLoc.Visible = True
5950            .opgLoc_box.Visible = True
5960            .Recur_Country.Visible = True
5970            .Recur_PostalCode.Visible = True
5980          Case False
5990            .Recur_Country.Visible = True
6000            .Recur_Country.Enabled = False
6010            .Recur_PostalCode.Visible = True
6020            .Recur_PostalCode.Enabled = False
6030            .opgLoc.Visible = True
6040            .opgLoc_box.Visible = True
6050            .opgLoc.Enabled = False
6060          End Select
6070          .RecurringItemAux_1099.Visible = True
6080          .RecurringItemAux_1099.Enabled = True
6090          .RecurringItemAux_TaxNo.Visible = True
6100          .RecurringItemAux_TaxNo.Enabled = False
6110          .RecurringItemAux_TaxNo.BorderColor = WIN_CLR_DISR
6120          .RecurringItemAux_TaxNo.BackStyle = acBackStyleTransparent
6130        End With
6140      Case .opgType_optReceived.OptionValue
6150        With frm
6160          .Recur_Type = "Payor"
6170          .Caption = "Add New Recurring Receipt"
6180          .Recur_lbl.Caption = "Recurring Receipt"
6190          .Header_lbl_Item.Visible = False
6200          .Header_lbl_Payment.Visible = False
6210          .Header_lbl_Receipt.Visible = True
6220          .Recur_Address.Visible = False
6230          .Recur_City.Visible = False
6240          .Recur_State.Visible = False
6250          .Recur_Zip.Visible = False
6260          Select Case blnTmp01
              Case True
6270            .opgLoc.Visible = False
6280            .opgLoc_box.Visible = False
6290            .Recur_Country.Visible = False
6300            .Recur_PostalCode.Visible = False
6310          Case False
6320            .Recur_Country.Visible = False
6330            .Recur_PostalCode.Visible = False
6340            .opgLoc.Visible = False
6350            .opgLoc_box.Visible = False
6360          End Select
6370          .RecurringItemAux_1099.Visible = False
6380          .RecurringItemAux_TaxNo.Visible = False
6390        End With
6400      End Select

6410    End With

EXITP:
6420    Set frm = Nothing
6430    Exit Sub

ERRH:
6440    Select Case ERR.Number
        Case Else
6450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6460    End Select
6470    Resume EXITP

End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_KeyDown"

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
          Case vbKeyTab
6570        With Me
6580          intRetVal = 0
6590          If .cmdDelete.Enabled = True Then
6600            .cmdDelete.SetFocus
6610          Else
6620            .cmdClose.SetFocus
6630          End If
6640        End With
6650      End Select
6660    End If

        ' ** Shift keys.
6670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6680      Select Case intRetVal
          Case vbKeyTab
6690        With Me
6700          intRetVal = 0
6710          If .chkEnableCountry.Enabled = True Then
6720            .chkEnableCountry.SetFocus
6730          Else
6740            .opgType.SetFocus
6750          End If
6760        End With
6770      End Select
6780    End If

EXITP:
6790    KeyCode = intRetVal
6800    Exit Sub

ERRH:
6810    intRetVal = 0
6820    Select Case ERR.Number
        Case Else
6830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6840    End Select
6850    Resume EXITP

End Sub

Private Sub cmdDelete_Click()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngRecurID As Long

6910    With Me
6920      If .frmRecurringItems_Sub.Form.Recur_Name = RECUR_I_TO_P Or _
              .frmRecurringItems_Sub.Form.Recur_Name = RECUR_P_TO_I Or _
              .frmRecurringItems_Sub.Form.RecurringItem_ID = RECUR_I_TO_P_ID Or _
              .frmRecurringItems_Sub.Form.RecurringItem_ID = RECUR_P_TO_I_ID Then
6930        MsgBox "You can not delete this item.", vbInformation + vbOKOnly, "Invalid Action"
6940      Else
6950        lngRecurID = .frmRecurringItems_Sub.Form.RecurringItem_ID
6960        If MsgBox("Are you sure you want to delete '" & .frmRecurringItems_Sub.Form.Recur_Name & "'?", _
                vbQuestion + vbYesNo, "Delete Recurring Item") = vbYes Then
6970          Set dbs = CurrentDb
6980          With dbs
                ' ** Delete RecurringItems, by specified [recurid].
6990            Set qdf = .QueryDefs("qryRecurringItems_09")
7000            With qdf.Parameters
7010              ![recurid] = lngRecurID
7020            End With
7030            qdf.Execute
7040            Set qdf = Nothing
7050            .Close
7060          End With
7070          Set dbs = Nothing
7080          DoEvents
7090          .frmRecurringItems_Sub.Form.Requery
7100          DoEvents
7110        End If
7120      End If
7130    End With

EXITP:
7140    Set qdf = Nothing
7150    Set dbs = Nothing
7160    Exit Sub

ERRH:
7170    Select Case ERR.Number
        Case 2427  ' ** You entered an expression that has no value.
7180      MsgBox "There are no recurring items to delete!", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
7190    Case Else
7200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7210    End Select
7220    Resume EXITP

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7320    intShiftDown = (Shift And acShiftMask) > 0
7330    intAltDown = (Shift And acAltMask) > 0
7340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7360      Select Case intRetVal
          Case vbKeyTab
7370        With Me
7380          intRetVal = 0
7390          .cmdClose.SetFocus
7400        End With
7410      End Select
7420    End If

        ' ** Shift keys.
7430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7440      Select Case intRetVal
          Case vbKeyTab
7450        With Me
7460          intRetVal = 0
7470          If .cmdAdd.Enabled = True Then
7480            .cmdAdd.SetFocus
7490          ElseIf .chkEnableCountry.Enabled = True Then
7500            .chkEnableCountry.SetFocus
7510          Else
7520            .opgType.SetFocus
7530          End If
7540        End With
7550      End Select
7560    End If

EXITP:
7570    KeyCode = intRetVal
7580    Exit Sub

ERRH:
7590    intRetVal = 0
7600    Select Case ERR.Number
        Case Else
7610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7620    End Select
7630    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

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

Private Sub opgType_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_box_MouseMove"

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

Private Sub opgType_box2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_box2_MouseMove"

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

Public Sub opgType_AfterUpdate()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_AfterUpdate"

8010    With Me
8020      Select Case .opgType
          Case .opgType_optMisc.OptionValue
            ' ** Miscellaneous Items.
8030        .frmRecurringItems_Sub.Form.Filter = "[Recur_Type] = 'Misc'"
8040        .frmRecurringItems_Sub.Form.FilterOn = True
8050        .opgType_optMisc_lbl.FontBold = True
8060        .opgType_optPaid_lbl.FontBold = False
8070        .opgType_optReceived_lbl.FontBold = False
8080        .opgType_optAll_lbl.FontBold = False
8090        .opgType_lbl_Item.Visible = True
8100        .opgType_lbl_Payment.Visible = False
8110        .opgType_lbl_Receipt.Visible = False
8120        .cmdAdd.Enabled = True
8130        .chkEnableCountry.Enabled = False
8140      Case .opgType_optPaid.OptionValue
            ' ** Payees.
8150        .frmRecurringItems_Sub.Form.Filter = "[Recur_Type] = 'Payee'"
8160        .frmRecurringItems_Sub.Form.FilterOn = True
8170        .opgType_optMisc_lbl.FontBold = False
8180        .opgType_optPaid_lbl.FontBold = True
8190        .opgType_optReceived_lbl.FontBold = False
8200        .opgType_optAll_lbl.FontBold = False
8210        .opgType_lbl_Item.Visible = False
8220        .opgType_lbl_Payment.Visible = True
8230        .opgType_lbl_Receipt.Visible = False
8240        .cmdAdd.Enabled = True
8250        .chkEnableCountry.Enabled = True
8260      Case .opgType_optReceived.OptionValue
            ' ** Payors.
8270        .frmRecurringItems_Sub.Form.Filter = "[Recur_Type] = 'Payor'"
8280        .frmRecurringItems_Sub.Form.FilterOn = True
8290        .opgType_optMisc_lbl.FontBold = False
8300        .opgType_optPaid_lbl.FontBold = False
8310        .opgType_optReceived_lbl.FontBold = True
8320        .opgType_optAll_lbl.FontBold = False
8330        .opgType_lbl_Item.Visible = False
8340        .opgType_lbl_Payment.Visible = False
8350        .opgType_lbl_Receipt.Visible = True
8360        .cmdAdd.Enabled = True
8370        .chkEnableCountry.Enabled = False
8380      Case .opgType_optAll.OptionValue
            ' ** Show all.
8390        .frmRecurringItems_Sub.Form.Filter = vbNullString
8400        .frmRecurringItems_Sub.Form.FilterOn = False
8410        .opgType_optMisc_lbl.FontBold = False
8420        .opgType_optPaid_lbl.FontBold = False
8430        .opgType_optReceived_lbl.FontBold = False
8440        .opgType_optAll_lbl.FontBold = True
8450        .opgType_lbl_Item.Visible = True
8460        .opgType_lbl_Payment.Visible = True
8470        .opgType_lbl_Receipt.Visible = True
8480        .cmdAdd.Enabled = False
8490        .chkEnableCountry.Enabled = True
8500        blnResetHeaders = True
8510        .TimerInterval = 1000&
8520      End Select
8530      .frmRecurringItems_Sub.Form.Refresh
8540      .frmRecurringItems_Sub.Form.EnableAddressInfo  ' ** Form Procedure: frmRecurringItems_Sub.
8550      .frmRecurringItems_Sub.SetFocus
8560      lngRecsCur = .frmRecurringItems_Sub.Form.RecCnt  ' ** Form Function: frmRecurringItems_Sub.
8570      If lngRecsCur > 0& Then
8580        .frmRecurringItems_Sub.Form.Recur_Name.SetFocus
8590      End If
8600    End With

EXITP:
8610    Exit Sub

ERRH:
8620    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Empty.
8630    Case Else
8640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8650    End Select
8660    Resume EXITP

End Sub

Private Sub opgType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_MouseMove"

8710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8720    Exit Sub

ERRH:
8730    Select Case ERR.Number
        Case Else
8740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8750    End Select
8760    Resume EXITP

End Sub

Private Sub opgType_optMisc_KeyDown(KeyCode As Integer, Shift As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optMisc_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8820    intShiftDown = (Shift And acShiftMask) > 0
8830    intAltDown = (Shift And acAltMask) > 0
8840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8870        With Me
8880          intRetVal = 0
8890          If .chkEnableCountry.Enabled = True Then
8900            .chkEnableCountry.SetFocus
8910          ElseIf .cmdAdd.Enabled = True Then
8920            .cmdAdd.SetFocus
8930          ElseIf .cmdDelete.Enabled = True Then
8940            .cmdDelete.SetFocus
8950          Else
8960            .cmdClose.SetFocus
8970          End If
8980        End With
8990      End Select
9000    End If

        ' ** Shift keys.
9010    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9020      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9030        With Me
9040          intRetVal = 0
9050          If .cmdPrintReport.Enabled = True Then
9060            .cmdPrintReport.SetFocus
9070          Else
9080            lngRecsCur = .frmRecurringItems_Sub.Form.RecCnt  ' ** Form Function: frmRecurringItems_Sub.
9090            If lngRecsCur > 0& Then
9100              .frmRecurringItems_Sub.SetFocus
9110              .frmRecurringItems_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmRecurringItems_Sub.
9120              If .frmRecurringItems_Sub.Form.Recur_Zip.Enabled = True And _
                      .frmRecurringItems_Sub.Form.Recur_Zip.Visible = True Then
9130                .frmRecurringItems_Sub.Form.Recur_Zip.SetFocus
9140              Else
9150                .frmRecurringItems_Sub.Form.Recur_Name.SetFocus
9160              End If
9170            Else
9180              .cmdClose.SetFocus
9190            End If
9200          End If
9210        End With
9220      End Select
9230    End If

EXITP:
9240    KeyCode = intRetVal
9250    Exit Sub

ERRH:
9260    intRetVal = 0
9270    Select Case ERR.Number
        Case Else
9280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9290    End Select
9300    Resume EXITP

End Sub

Private Sub opgType_optMisc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optMisc_MouseMove"

9410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9420    Exit Sub

ERRH:
9430    Select Case ERR.Number
        Case Else
9440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9450    End Select
9460    Resume EXITP

End Sub

Private Sub opgType_optPaid_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optPaid_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
9570        With Me
9580          intRetVal = 0
9590          If .chkEnableCountry.Enabled = True Then
9600            .chkEnableCountry.SetFocus
9610          ElseIf .cmdAdd.Enabled = True Then
9620            .cmdAdd.SetFocus
9630          ElseIf .cmdDelete.Enabled = True Then
9640            .cmdDelete.SetFocus
9650          Else
9660            .cmdClose.SetFocus
9670          End If
9680        End With
9690      End Select
9700    End If

        ' ** Shift keys.
9710    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9720      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9730        With Me
9740          intRetVal = 0
9750          If .cmdPrintReport.Enabled = True Then
9760            .cmdPrintReport.SetFocus
9770          Else
9780            lngRecsCur = .frmRecurringItems_Sub.Form.RecCnt  ' ** Form Function: frmRecurringItems_Sub.
9790            If lngRecsCur > 0& Then
9800              .frmRecurringItems_Sub.SetFocus
9810              .frmRecurringItems_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmRecurringItems_Sub.
9820              If .frmRecurringItems_Sub.Form.Recur_Zip.Enabled = True And _
                      .frmRecurringItems_Sub.Form.Recur_Zip.Visible = True Then
9830                .frmRecurringItems_Sub.Form.Recur_Zip.SetFocus
9840              Else
9850                .frmRecurringItems_Sub.Form.Recur_Name.SetFocus
9860              End If
9870            Else
9880              .cmdClose.SetFocus
9890            End If
9900          End If
9910        End With
9920      End Select
9930    End If

EXITP:
9940    KeyCode = intRetVal
9950    Exit Sub

ERRH:
9960    intRetVal = 0
9970    Select Case ERR.Number
        Case Else
9980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9990    End Select
10000   Resume EXITP

End Sub

Private Sub opgType_optPaid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optPaid_MouseMove"

10110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10120   Exit Sub

ERRH:
10130   Select Case ERR.Number
        Case Else
10140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10150   End Select
10160   Resume EXITP

End Sub

Private Sub opgType_optReceived_KeyDown(KeyCode As Integer, Shift As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optReceived_KeyDown"

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
10290         If .chkEnableCountry.Enabled = True Then
10300           .chkEnableCountry.SetFocus
10310         ElseIf .cmdAdd.Enabled = True Then
10320           .cmdAdd.SetFocus
10330         ElseIf .cmdDelete.Enabled = True Then
10340           .cmdDelete.SetFocus
10350         Else
10360           .cmdClose.SetFocus
10370         End If
10380       End With
10390     End Select
10400   End If

        ' ** Shift keys.
10410   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10420     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10430       With Me
10440         intRetVal = 0
10450         If .cmdPrintReport.Enabled = True Then
10460           .cmdPrintReport.SetFocus
10470         Else
10480           lngRecsCur = .frmRecurringItems_Sub.Form.RecCnt  ' ** Form Function: frmRecurringItems_Sub.
10490           If lngRecsCur > 0& Then
10500             .frmRecurringItems_Sub.SetFocus
10510             .frmRecurringItems_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmRecurringItems_Sub.
10520             If .frmRecurringItems_Sub.Form.Recur_Zip.Enabled = True And _
                      .frmRecurringItems_Sub.Form.Recur_Zip.Visible = True Then
10530               .frmRecurringItems_Sub.Form.Recur_Zip.SetFocus
10540             Else
10550               .frmRecurringItems_Sub.Form.Recur_Name.SetFocus
10560             End If
10570           Else
10580             .cmdClose.SetFocus
10590           End If
10600         End If
10610       End With
10620     End Select
10630   End If

EXITP:
10640   KeyCode = intRetVal
10650   Exit Sub

ERRH:
10660   intRetVal = 0
10670   Select Case ERR.Number
        Case Else
10680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10690   End Select
10700   Resume EXITP

End Sub

Private Sub opgType_optReceived_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optReceived_MouseMove"

10810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10820   Exit Sub

ERRH:
10830   Select Case ERR.Number
        Case Else
10840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10850   End Select
10860   Resume EXITP

End Sub

Private Sub opgType_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10920   intShiftDown = (Shift And acShiftMask) > 0
10930   intAltDown = (Shift And acAltMask) > 0
10940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10970       With Me
10980         intRetVal = 0
10990         If .chkEnableCountry.Enabled = True Then
11000           .chkEnableCountry.SetFocus
11010         ElseIf .cmdAdd.Enabled = True Then
11020           .cmdAdd.SetFocus
11030         ElseIf .cmdDelete.Enabled = True Then
11040           .cmdDelete.SetFocus
11050         Else
11060           .cmdClose.SetFocus
11070         End If
11080       End With
11090     End Select
11100   End If

        ' ** Shift keys.
11110   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11130       With Me
11140         intRetVal = 0
11150         If .cmdPrintReport.Enabled = True Then
11160           .cmdPrintReport.SetFocus
11170         Else
11180           lngRecsCur = .frmRecurringItems_Sub.Form.RecCnt  ' ** Form Function: frmRecurringItems_Sub.
11190           If lngRecsCur > 0& Then
11200             .frmRecurringItems_Sub.SetFocus
11210             .frmRecurringItems_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmRecurringItems_Sub.
11220             If .frmRecurringItems_Sub.Form.Recur_Zip.Enabled = True And _
                      .frmRecurringItems_Sub.Form.Recur_Zip.Visible = True Then
11230               .frmRecurringItems_Sub.Form.Recur_Zip.SetFocus
11240             Else
11250               .frmRecurringItems_Sub.Form.Recur_Name.SetFocus
11260             End If
11270           Else
11280             .cmdClose.SetFocus
11290           End If
11300         End If
11310       End With
11320     End Select
11330   End If

EXITP:
11340   KeyCode = intRetVal
11350   Exit Sub

ERRH:
11360   intRetVal = 0
11370   Select Case ERR.Number
        Case Else
11380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11390   End Select
11400   Resume EXITP

End Sub

Private Sub opgType_optAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optAll_MouseMove"

11510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11520   Exit Sub

ERRH:
11530   Select Case ERR.Number
        Case Else
11540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11550   End Select
11560   Resume EXITP

End Sub

Private Sub chkEnableCountry_AfterUpdate()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkEnableCountry_AfterUpdate"

11610   With Me
11620     Select Case .chkEnableCountry
          Case True
11630       .chkEnableCountry_lbl.FontBold = True
11640       .frmRecurringItems_Sub.Form.EnableCountry True  ' ** Form Procedure: frmRecurringItems_Sub.
11650       Select Case .opgType
            Case .opgType_optPaid.OptionValue
11660         .frmRecurringItems_Sub.Form.RecurHeaders True  ' ** Form Procedure: frmRecurringItems_Sub.
11670       Case .opgType_optAll.OptionValue
11680         If .frmRecurringItems_Sub.Form.Recur_Type = "Payee" Then
11690           .frmRecurringItems_Sub.Form.RecurHeaders True  ' ** Form Procedure: frmRecurringItems_Sub.
11700         Else
11710           .frmRecurringItems_Sub.Form.RecurHeaders False  ' ** Form Procedure: frmRecurringItems_Sub.
11720         End If
11730       Case Else
11740         .frmRecurringItems_Sub.Form.RecurHeaders False  ' ** Form Procedure: frmRecurringItems_Sub.
11750       End Select
11760       DoEvents
11770       EnableCountry_SetFrmWidth  ' ** Procedure: Below
11780     Case False
11790       .chkEnableCountry_lbl.FontBold = False
11800       .frmRecurringItems_Sub.Form.EnableCountry False  ' ** Form Procedure: frmRecurringItems_Sub.
11810       DoEvents
11820       EnableCountry_SetFrmWidth  ' ** Procedure: Below
11830     End Select
11840     DoEvents
11850     lngRecsCur = .frmRecurringItems_Sub.Form.RecCnt  ' ** Form Function: frmRecurringItems_Sub.
11860     If lngRecsCur > 0& Then
11870       .frmRecurringItems_Sub.SetFocus
11880     Else
11890       .cmdAdd.SetFocus
11900     End If
11910     blnResetHeaders = True  ' ** To make sure it's pushed left.
11920     .TimerInterval = 250&
11930   End With

EXITP:
11940   Exit Sub

ERRH:
11950   Select Case ERR.Number
        Case Else
11960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11970   End Select
11980   Resume EXITP

End Sub

Private Sub chkEnableCountry_KeyDown(KeyCode As Integer, Shift As Integer)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkEnableCountry_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12020   intShiftDown = (Shift And acShiftMask) > 0
12030   intAltDown = (Shift And acAltMask) > 0
12040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12070       With Me
12080         intRetVal = 0
12090         If .cmdAdd.Enabled = True Then
12100           .cmdAdd.SetFocus
12110         ElseIf .cmdDelete.Enabled = True Then
12120           .cmdDelete.SetFocus
12130         Else
12140           .cmdClose.SetFocus
12150         End If
12160       End With
12170     End Select
12180   End If

        ' ** Shift keys.
12190   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12200     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12210       With Me
12220         intRetVal = 0
12230         .opgType.SetFocus
12240       End With
12250     End Select
12260   End If

EXITP:
12270   KeyCode = intRetVal
12280   Exit Sub

ERRH:
12290   intRetVal = 0
12300   Select Case ERR.Number
        Case Else
12310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12320   End Select
12330   Resume EXITP

End Sub

Private Sub EnableCountry_SetFrmWidth()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "EnableCountry_SetFrmWidth"

        Dim lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, dblTmp04 As Double

12410   With Me

12420     GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

12430     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
12440     lngMonitorNum = 1&: lngTmp03 = 0&
12450     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
12460     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

12470     Select Case .chkEnableCountry
          Case True
            ' ** Design View set at USA-Only width.
12480       If lngFrm_Width <> (lngFrm_Width_Orig + lngCountry_Diff) Then
12490         .Width = (lngForm_Width + lngCountry_Diff)
12500         lngTmp01 = (lngFrm_Left_Orig - (lngCountry_Diff / 2))
12510         If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top_Orig
12520         DoCmd.MoveSize lngTmp01, lngTmp03, (lngFrm_Width_Orig + lngCountry_Diff), lngFrm_Height_Orig  'lngFrm_Top_Orig
12530         If lngMonitorNum > 1& Then
12540           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
12550         End If
12560         .frmRecurringItems_Sub.Width = (lngSub_Width + lngCountry_Diff)
12570         .frmRecurringItems_Sub_box.Width = (.frmRecurringItems_Sub.Width + (2& * lngTpp))
12580         .Nav_box01.Width = .frmRecurringItems_Sub.Width
12590         .cmdClose.Left = (lngClose_Left + lngCountry_Diff)
12600         .cmdDelete.Left = (lngDelete_Left + lngCountry_Diff)
12610         .cmdAdd.Left = (lngAdd_Left + lngCountry_Diff)
12620         dblTmp04 = (lngCountry_Diff / 2)
12630         dblTmp04 = (dblTmp04 / lngTpp)
12640         If InStr(CStr(dblTmp04), ".") > 0 Then
12650           lngTmp02 = CInt(Left(CStr(dblTmp04), (InStr(CStr(dblTmp04), ".") - 1)))
12660           lngTmp02 = (lngTmp02 * lngTpp)
12670         Else
12680           lngTmp02 = CInt(dblTmp04)
12690           lngTmp02 = (lngTmp02 * lngTpp)
12700         End If
12710         .Header_vline01.Left = .Width
12720         .Header_vline02.Left = .Width
12730         .Detail_vline01.Left = .Width
12740         .Detail_vline02.Left = .Width
12750         .Footer_vline01.Left = .Width
12760         .Footer_vline02.Left = .Width
12770         .Header_hline01.Width = .Width
12780         .Header_hline02.Width = .Width
12790         .Detail_hline01.Width = .Width
12800         .Detail_hline02.Width = .Width
12810         .Footer_hline01.Width = .Width
12820         .Footer_hline02.Width = .Width
12830       End If
12840     Case False
12850       If lngFrm_Width <> lngFrm_Width_Orig Then
12860         If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top_Orig
12870         DoCmd.MoveSize lngFrm_Left_Orig, lngTmp03, lngFrm_Width_Orig, lngFrm_Height_Orig  'lngFrm_Top_Orig
12880         If lngMonitorNum > 1& Then
12890           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
12900         End If
12910         .cmdClose.Left = lngClose_Left
12920         .cmdDelete.Left = lngDelete_Left
12930         .cmdAdd.Left = lngAdd_Left
12940         .Header_vline01.Left = lngForm_Width
12950         .Header_vline02.Left = lngForm_Width
12960         .Detail_vline01.Left = lngForm_Width
12970         .Detail_vline02.Left = lngForm_Width
12980         .Footer_vline01.Left = lngForm_Width
12990         .Footer_vline02.Left = lngForm_Width
13000         .Header_hline01.Width = lngForm_Width
13010         .Header_hline02.Width = lngForm_Width
13020         .Detail_hline01.Width = lngForm_Width
13030         .Detail_hline02.Width = lngForm_Width
13040         .Footer_hline01.Width = lngForm_Width
13050         .Footer_hline02.Width = lngForm_Width
13060         .Width = lngForm_Width
13070       End If
13080     End Select

13090   End With

EXITP:
13100   Exit Sub

ERRH:
13110   Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Confusion.
13120   Case Else
13130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13140   End Select
13150   Resume EXITP

End Sub
