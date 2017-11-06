VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_History"

'VGC 03/22/2017: CHANGES!

'IF THE HISTORY DATE IS THE SAME AS THE CURRENT DATE, NO EDIT!
'THEY MUST DO THAT IN THE CURRENCY RATE WINDOW!
'AND NO ACTIVE CHANGING!

' ** Shortcut Alt keys responsive from this form:
' **   Active Only Date: A {chkActiveOnly_Dates}
' **   Close:            C {cmdClose}
' **   Prior Dates:      D {lbxHistoryDate}
' **   Edit:             E {cmdEdit}
' **   Currency Code:    O {curr_code on frmCurrency_History_Sub}
' **   Print Report:     P {cmdPrintReport}
' **   Prior Rates:      R {lbxHistoryRate}
' **   Restore:          S {cmdRestore}
' **   Active Only Rate: V {chkActiveOnly_Rates}
' **   Exit:             X {cmdClose}

' ** List box column constants: lbxHistoryDate.
Private Const LBX_DT_DATE As Integer = 0  ' ** curr_date
'Private Const LBX_DT_DEL  As Integer = 1  ' ** date_del
'Private Const LBX_DT_CNT  As Integer = 2  ' ** cnt

' ** List box column constants: lbxHistoryRate.
Private Const LBX_RT_CURID As Integer = 0  ' ** curr_id
Private Const LBX_RT_CODE  As Integer = 1  ' ** curr_code
'Private Const LBX_RT_NAME  As Integer = 2  ' ** curr_name
'Private Const LBX_RT_DEL   As Integer = 3  ' ** rate_del
'Private Const LBX_RT_CNT   As Integer = 4  ' ** cnt

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, lngPrintReportLbl_Left As Long
Private blnIsOpen As Boolean, blnFromActiveOnly As Boolean, blnIsDesc As Boolean
Private lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnIsOpen = True: blnFromActiveOnly = False: blnIsDesc = False

160           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

170           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

180           If gblnGoToReport = True Then
190             DoCmd.Hourglass True  ' ** Make sure it's still running.
200             DoEvents
210             .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
220             .GoToReport_arw_pr_img.Visible = True
230             If .cmdPrintReport.Enabled = True Then
240               .cmdPrintReport.SetFocus
250             End If
260             DoEvents
270           End If

280           .lbxHistoryDate.RowSource = "qryCurrency_History_02_02"
290           .Sort_Date_lbl_desc_img.Visible = True
300           .Sort_Date_lbl_asc_img.Visible = False

310           .lbxHistoryRate.RowSource = "qryCurrency_History_03_01"
320           .Sort_Code_line1.Visible = True
330           .Sort_Code_line2.Visible = True
340           .Sort_Code_lbl_asc_img.Visible = True
350           .Sort_Code_lbl_desc_img.Visible = False
360           .Sort_Name_line1.Visible = False
370           .Sort_Name_line2.Visible = False
380           .Sort_Name_lbl_asc_img.Visible = False
390           .Sort_Name_lbl_desc_img.Visible = False

400           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

410           chkActiveOnly_Dates_AfterUpdate  ' ** Procedure: Below.
420           chkActiveOnly_Rates_AfterUpdate  ' ** Procedure: Below.

430           If IsNull(.lbxHistoryDate) = True Then
440             .lbxHistoryDate = .lbxHistoryDate.Column(0, 0)
450           End If

460           lbxHistoryDate_AfterUpdate  ' ** Procedure: Below.
470           .cmdEdit.Enabled = False

480           blnIsOpen = False

490         Else
500           Cancel = -1
510         End If
520       Else
530         Cancel = -1
540       End If
550     End With

560     DoCmd.Hourglass False

570     If Cancel = -1 Then
580       DoCmd.Hourglass False
590       Beep
600       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
610     End If

EXITP:
620     Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Prior Dates:      D {lbxHistoryDate}
        ' **   Currency Code:    O {curr_code on frmCurrency_History_Sub}
        ' **   Prior Rates:      R {lbxHistoryRate}
        ' **   Exit:             X {cmdClose}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         intRetVal = 0
780         cmdClose_Click  ' ** Procedure: Below.
790       End Select
800     End If

        ' ** Alt keys.
810     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
820       Select Case intRetVal
          Case vbKeyD
830         With Me
840           intRetVal = 0
850           .lbxHistoryDate.SetFocus
860         End With
870       Case vbKeyO
880         With Me
890           intRetVal = 0
900           lngRecsCur = .frmCurrency_History_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_History_Sub.
910           If lngRecsCur > 0& Then
920             .frmCurrency_History_Sub.SetFocus
930             .frmCurrency_History_Sub.Form.curr_code.SetFocus
940           Else
950             Beep
960           End If
970         End With
980       Case vbKeyR
990         With Me
1000          intRetVal = 0
1010          .lbxHistoryRate.SetFocus
1020        End With
1030      Case vbKeyX
1040        intRetVal = 0
1050        cmdClose_Click  ' ** Procedure: Below.
1060      End Select
1070    End If

        ' ** Ctrl-Shift keys.
1080    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1090      Select Case intRetVal
          Case vbKeyF
1100        With Me
1110          intRetVal = 0
1120          .FocusHolder.SetFocus
1130        End With
1140      End Select
1150    End If

EXITP:
1160    KeyCode = intRetVal
1170    Exit Sub

ERRH:
1180    intRetVal = 0
1190    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1200    Case Else
1210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1220    End Select
1230    Resume EXITP

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
1420      End If  ' ** gblnSetFocus.
1430      If gblnGoToReport = True Then
1440        DoCmd.Hourglass True  ' ** Make sure it's still running.
1450        DoEvents
1460        Select Case blnGoingToReport
            Case True
1470          gblnGoToReport = False
1480          blnGoingToReport = False
1490          .GoToReport_arw_pr_img.Visible = False
1500          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1510          DoCmd.Hourglass False
1520          DoEvents
1530        Case False
1540          blnGoingToReport = True
1550          .GoToReport_arw_pr_img.Visible = True
1560          DoEvents
1570          Forms(strCallingForm).TimerInterval = 0&
1580          DoCmd.SelectObject acForm, THIS_NAME, False  ' ** Don't let it slip back to the other form!
1590          DoEvents
1600          Forms(strCallingForm).GTREmblem_Off  ' ** Form Procedure: frmCurrency_Rate.
1610          DoCmd.SelectObject acForm, THIS_NAME, False
1620          DoEvents
1630          Forms(strCallingForm).GoToReport_arw_hist_img.Visible = False
1640          DoCmd.SelectObject acForm, THIS_NAME, False
1650          DoEvents
1660          DoCmd.SelectObject acForm, THIS_NAME, False
1670          gblnSetFocus = True
1680          Beep
1690          .TimerInterval = (GTR_WAIT * 2&)
1700        End Select
1710      End If
1720    End With

EXITP:
1730    Exit Sub

ERRH:
1740    DoCmd.Hourglass False
1750    Select Case ERR.Number
        Case Else
1760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1770    End Select
1780    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1810    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1820    gstrReturningForm = THIS_NAME
1830    gblnSetFocus = True
1840    Select Case strCallingForm
        Case "frmCurrency_Rate"
1850      If IsLoaded(strCallingForm, acForm) = True Then
1860        DoCmd.SelectObject acForm, strCallingForm, False
1870        Forms(strCallingForm).TimerInterval = 250&
1880      Else
1890        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_ForeignExchange"
1900      End If
1910    Case Else
          ' ** Nothing at the moment.
1920    End Select

EXITP:
1930    Exit Sub

ERRH:
1940    Select Case ERR.Number
        Case Else
1950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1960    End Select
1970    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2010    With Me
2020      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2030        Select Case blnPrintReport_Focus
            Case True
2040          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2050          .cmdPrintReport_raised_img.Visible = False
2060        Case False
2070          .cmdPrintReport_raised_img.Visible = True
2080          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2090        End Select
2100        .cmdPrintReport_raised_focus_img.Visible = False
2110        .cmdPrintReport_raised_focus_dots_img.Visible = False
2120        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2130        .cmdPrintReport_raised_img_dis.Visible = False
2140      End If
2150    End With

EXITP:
2160    Exit Sub

ERRH:
2170    Select Case ERR.Number
        Case Else
2180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2190    End Select
2200    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2410    DoCmd.Close acForm, THIS_NAME

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
2510    intShiftDown = (Shift And acShiftMask) > 0
2520    intAltDown = (Shift And acAltMask) > 0
2530    intCtrlDown = (Shift And acCtrlMask) > 0

2540    intRetVal = KeyCode

        ' ** Plain keys.
2550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2560      Select Case intRetVal
          Case vbKeyTab
2570        With Me
2580          intRetVal = 0
2590          If .lbxHistoryDate.Enabled = True Then
2600            .lbxHistoryDate.SetFocus
2610          ElseIf .lbxHistoryRate.Enabled = True Then
2620            .lbxHistoryRate.SetFocus
2630          End If
2640        End With
2650      End Select
2660    End If

        ' ** Shift keys.
2670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2680      Select Case intRetVal
          Case vbKeyTab
2690        With Me
2700          intRetVal = 0
2710          If .cmdEdit.Enabled = True Then
2720            .cmdEdit.SetFocus
2730          ElseIf .cmdRestore.Enabled = True Then
2740            .cmdRestore.SetFocus
2750          ElseIf .cmdPrintReport.Enabled = True Then
2760            .cmdPrintReport.SetFocus
2770          Else
2780            lngRecsCur = .frmCurrency_History_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_History_Sub.
2790            If lngRecsCur > 0& Then
2800              .frmCurrency_History_Sub.SetFocus
2810              .frmCurrency_History_Sub.Form.FocusHolder.SetFocus
2820            Else
2830              .chkActiveOnly_Rates.SetFocus
2840            End If
2850          End If
2860        End With
2870      End Select
2880    End If

EXITP:
2890    KeyCode = intRetVal
2900    Exit Sub

ERRH:
2910    intRetVal = 0
2920    Select Case ERR.Number
        Case Else
2930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2940    End Select
2950    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

3010    With Me
3020      blnPrintReport_Focus = True
3030      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3040      .cmdPrintReport_raised_img.Visible = False
3050      .cmdPrintReport_raised_focus_img.Visible = False
3060      .cmdPrintReport_raised_focus_dots_img.Visible = False
3070      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3080      .cmdPrintReport_raised_img_dis.Visible = False
3090    End With

EXITP:
3100    Exit Sub

ERRH:
3110    Select Case ERR.Number
        Case Else
3120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3130    End Select
3140    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

3210    With Me
3220      blnPrintReport_MouseDown = True
3230      .cmdPrintReport_sunken_focus_dots_img.Visible = True
3240      .cmdPrintReport_raised_img.Visible = False
3250      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3260      .cmdPrintReport_raised_focus_img.Visible = False
3270      .cmdPrintReport_raised_focus_dots_img.Visible = False
3280      .cmdPrintReport_raised_img_dis.Visible = False
3290    End With

EXITP:
3300    Exit Sub

ERRH:
3310    Select Case ERR.Number
        Case Else
3320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3330    End Select
3340    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

3410    With Me
3420      lngRecsCur = .frmCurrency_History_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_History_Sub.
3430      If lngRecsCur > 0& Then
3440        strDocName = "rptCurrencyRateHistory"
3450        If gblnDev_Debug Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
3460          DoCmd.OpenReport strDocName, acViewPreview
3470          DoCmd.Maximize
3480          DoCmd.RunCommand acCmdFitToWindow
3490          .Visible = False
3500        Else
              '##GTR_Ref: rptCurrencyRateHistory
3510          DoCmd.OpenReport strDocName, acViewNormal
3520        End If
3530      Else
3540        Beep
3550      End If
3560    End With

EXITP:
3570    Exit Sub

ERRH:
3580    Select Case ERR.Number
        Case Else
3590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3600    End Select
3610    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
3790          If .cmdRestore.Enabled = True Then
3800            .cmdRestore.SetFocus
3810          ElseIf .cmdEdit.Enabled = True Then
3820            .cmdEdit.SetFocus
3830          Else
3840            .cmdClose.SetFocus
3850          End If
3860        End With
3870      End Select
3880    End If

        ' ** Shift keys.
3890    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3900      Select Case intRetVal
          Case vbKeyTab
3910        With Me
3920          intRetVal = 0
3930          lngRecsCur = .frmCurrency_History_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_History_Sub.
3940          If lngRecsCur > 0& Then
3950            .frmCurrency_History_Sub.SetFocus
3960            .frmCurrency_History_Sub.Form.FocusHolder.SetFocus
3970          Else
3980            .chkActiveOnly_Rates.SetFocus
3990          End If
4000        End With
4010      End Select
4020    End If

EXITP:
4030    KeyCode = intRetVal
4040    Exit Sub

ERRH:
4050    intRetVal = 0
4060    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4070    Case Else
4080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4090    End Select
4100    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

4210    With Me
4220      If blnPrintReport_MouseDown = False Then
4230        Select Case blnPrintReport_Focus
            Case True
4240          .cmdPrintReport_raised_focus_dots_img.Visible = True
4250          .cmdPrintReport_raised_focus_img.Visible = False
4260        Case False
4270          .cmdPrintReport_raised_focus_img.Visible = True
4280          .cmdPrintReport_raised_focus_dots_img.Visible = False
4290        End Select
4300        .cmdPrintReport_raised_img.Visible = False
4310        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4320        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4330        .cmdPrintReport_raised_img_dis.Visible = False
4340      End If
4350    End With

EXITP:
4360    Exit Sub

ERRH:
4370    Select Case ERR.Number
        Case Else
4380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4390    End Select
4400    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

4510    With Me
4520      .cmdPrintReport_raised_focus_dots_img.Visible = True
4530      .cmdPrintReport_raised_img.Visible = False
4540      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4550      .cmdPrintReport_raised_focus_img.Visible = False
4560      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4570      .cmdPrintReport_raised_img_dis.Visible = False
4580      blnPrintReport_MouseDown = False
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

Private Sub cmdPrintReport_LostFocus()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

4710    With Me
4720      .cmdPrintReport_raised_img.Visible = True
4730      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4740      .cmdPrintReport_raised_focus_img.Visible = False
4750      .cmdPrintReport_raised_focus_dots_img.Visible = False
4760      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4770      .cmdPrintReport_raised_img_dis.Visible = False
4780      blnPrintReport_Focus = False
4790    End With

EXITP:
4800    Exit Sub

ERRH:
4810    Select Case ERR.Number
        Case Else
4820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4830    End Select
4840    Resume EXITP

End Sub

Private Sub cmdRestore_Click()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRestore_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim datCurrDate As Date, lngCurrID As Long
        Dim strDocName As String, strCurrCode As String, strCurrName As String
        Dim blnAll As Boolean

4910    With Me
4920      blnAll = False: gblnMessage = False
4930      lngRecsCur = .frmCurrency_History_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_History_Sub.
4940      If lngRecsCur > 0& Then
4950        If IsNull(.frmCurrency_History_Sub.Form.source) = False Then
4960          Select Case IsDate(.frmCurrency_History_Sub.Form.source)
              Case True
4970            If IsNull(.lbxHistoryDate) = False Then
4980              datCurrDate = CDate(.lbxHistoryDate)
4990              lngCurrID = .frmCurrency_History_Sub.Form.curr_id
5000              strCurrCode = .frmCurrency_History_Sub.Form.curr_code
5010              strCurrName = .frmCurrency_History_Sub.Form.curr_name
5020              gblnMessage = True
5030              strDocName = "frmCurrency_History_Restore"
5040              DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME & "~Date~" & Format(datCurrDate, "mm/dd/yyyy") & "~" & _
                    CStr(lngCurrID) & "~" & strCurrCode & "~" & strCurrName
5050              If gblnMessage = True Then
5060                DoCmd.Hourglass True
5070                DoEvents
5080                Select Case IsNull(.curr_date_all)
                    Case True
                      ' ** Just one rate and date.
5090                  lngCurrID = .curr_id_one
5100                  datCurrDate = .curr_date_one
5110                Case False
                      ' ** All rates for the date.
5120                  blnAll = True
5130                  lngCurrID = 0&
5140                  datCurrDate = .curr_date_all
5150                End Select
5160                Set dbs = CurrentDb
5170                With dbs
5180                  Select Case blnAll
                      Case True
                        ' ** Update qryCurrency_History_Restore_02_02 (tblCurrency, linked to
                        ' ** qryCurrency_History_Restore_02_01 (tblCurrency_History, by specified
                        ' ** [curdat]), with curr_date_new, curr_rate1_new, curr_rate2_new).
5190                    Set qdf = .QueryDefs("qryCurrency_History_Restore_02_03")
5200                    With qdf.Parameters
5210                      ![curdat] = datCurrDate
5220                    End With
5230                    qdf.Execute
5240                    Set qdf = Nothing
5250                  Case False
                        ' ** Update qryCurrency_History_Restore_03_02 (tblCurrency, linked to
                        ' ** qryCurrency_History_Restore_03_01 (tblCurrency_History, by specified
                        ' ** [curdat], [curid]), with curr_date_new, curr_rate1_new, curr_rate2_new).
5260                    Set qdf = .QueryDefs("qryCurrency_History_Restore_03_03")
5270                    With qdf.Parameters
5280                      ![curdat] = datCurrDate
5290                      ![curid] = lngCurrID
5300                    End With
5310                    qdf.Execute
5320                    Set qdf = Nothing
5330                  End Select
5340                  .Close
5350                End With
5360                Set dbs = Nothing
5370              End If
5380            End If  ' ** IsNull().
5390          Case False
5400            If IsNull(.lbxHistoryRate) = False Then
5410              datCurrDate = .frmCurrency_History_Sub.Form.curr_date
5420              lngCurrID = .lbxHistoryRate
5430              strCurrCode = .frmCurrency_History_Sub.Form.curr_code
5440              strCurrName = .frmCurrency_History_Sub.Form.curr_name
5450              gblnMessage = True
5460              strDocName = "frmCurrency_History_Restore"
5470              DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME & "~Rate~" & Format(datCurrDate, "mm/dd/yyyy") & "~" & _
                    CStr(lngCurrID) & "~" & strCurrCode & "~" & strCurrName
5480              If gblnMessage = True Then
5490                DoCmd.Hourglass True
5500                DoEvents
5510                lngCurrID = .curr_id_one
5520                datCurrDate = .curr_date_one
5530                Set dbs = CurrentDb
5540                With dbs
                      ' ** Update qryCurrency_History_Restore_03_02 (tblCurrency, linked to
                      ' ** qryCurrency_History_Restore_03_01 (tblCurrency_History, by specified
                      ' ** [curdat], [curid]), with curr_date_new, curr_rate1_new, curr_rate2_new).
5550                  Set qdf = .QueryDefs("qryCurrency_History_Restore_03_03")
5560                  With qdf.Parameters
5570                    ![curdat] = datCurrDate
5580                    ![curid] = lngCurrID
5590                  End With
5600                  qdf.Execute
5610                  Set qdf = Nothing
5620                  .Close
5630                End With
5640                Set dbs = Nothing
5650              End If
5660            End If  ' ** IsNull().
5670          End Select  ' ** source.
5680        End If  ' ** source.
5690      End If  ' ** lngRecsCur.
5700      If IsLoaded(strCallingForm, acForm) = True And gblnMessage = True Then  ' ** Module Function: modFileUtilities.
5710        Forms(strCallingForm).frmCurrency_Rate_Sub.Form.Refresh
5720      End If
5730      Beep
5740      DoCmd.Hourglass False
5750    End With

EXITP:
5760    Set qdf = Nothing
5770    Set dbs = Nothing
5780    Exit Sub

ERRH:
5790    DoCmd.Hourglass False
5800    Select Case ERR.Number
        Case Else
5810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5820    End Select
5830    Resume EXITP

End Sub

Private Sub cmdRestore_KeyDown(KeyCode As Integer, Shift As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRestore_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5920    intShiftDown = (Shift And acShiftMask) > 0
5930    intAltDown = (Shift And acAltMask) > 0
5940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5960      Select Case intRetVal
          Case vbKeyTab
5970        With Me
5980          intRetVal = 0
5990          If .cmdEdit.Enabled = True Then
6000            .cmdEdit.SetFocus
6010          Else
6020            .cmdClose.SetFocus
6030          End If
6040        End With
6050      End Select
6060    End If

        ' ** Shift keys.
6070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6080      Select Case intRetVal
          Case vbKeyTab
6090        With Me
6100          intRetVal = 0
6110          .cmdPrintReport.SetFocus
6120        End With
6130      End Select
6140    End If

EXITP:
6150    KeyCode = intRetVal
6160    Exit Sub

ERRH:
6170    intRetVal = 0
6180    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6190    Case Else
6200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6210    End Select
6220    Resume EXITP

End Sub

Private Sub Sort_Date_lbl_cmd_DblClick(Cancel As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "Sort_Date_lbl_cmd_DblClick"

6310    SortList THIS_PROC  ' ** Procedure: Below.

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Sub

Private Sub lbxHistoryDate_GotFocus()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxHistoryDate_GotFocus"

6410    With Me
6420      .lbxHistoryDate.ForeColor = CLR_BLK
6430      .lbxHistoryDate.BackColor = CLR_WHT
6440      .lbxHistoryRate.ForeColor = CLR_DKGRY   ' ** 92/92/92
6450      .lbxHistoryRate.BackColor = 16448250  ' ** 250/250/250
6460    End With

EXITP:
6470    Exit Sub

ERRH:
6480    Select Case ERR.Number
        Case Else
6490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6500    End Select
6510    Resume EXITP

End Sub

Private Sub lbxHistoryDate_AfterUpdate()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxHistoryDate_AfterUpdate"

        Dim strFilter As String

6610    With Me
6620      If IsNull(.lbxHistoryDate) = False Then
6630        .lbxHistoryRate = Null
6640        strFilter = "[curr_date] = #" & .lbxHistoryDate.Column(LBX_DT_DATE) & "#"
6650        If .chkActiveOnly_Dates = True Then
6660          strFilter = strFilter & " And [curr_active] = True"
6670        End If
6680        .frmCurrency_History_Sub.Form.Filter = strFilter
6690        .frmCurrency_History_Sub.Form.FilterOn = True
6700        .frmCurrency_History_Sub.Form.source = .lbxHistoryDate.Column(LBX_DT_DATE)
6710        DoEvents
6720        .frmCurrency_History_Sub.Form.Refresh
6730        DoEvents
6740      End If
6750    End With

EXITP:
6760    Exit Sub

ERRH:
6770    Select Case ERR.Number
        Case Else
6780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6790    End Select
6800    Resume EXITP

End Sub

Private Sub lbxHistoryDate_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxHistoryDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
6910    intShiftDown = (Shift And acShiftMask) > 0
6920    intAltDown = (Shift And acAltMask) > 0
6930    intCtrlDown = (Shift And acCtrlMask) > 0

6940    intRetVal = KeyCode

        ' ** Plain keys.
6950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6960      Select Case intRetVal
          Case vbKeyTab
6970        With Me
6980          intRetVal = 0
6990          .chkActiveOnly_Dates.SetFocus
7000        End With
7010      End Select
7020    End If

        ' ** Shift keys.
7030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7040      Select Case intRetVal
          Case vbKeyTab
7050        With Me
7060          intRetVal = 0
7070          .cmdClose.SetFocus
7080        End With
7090      End Select
7100    End If

EXITP:
7110    KeyCode = intRetVal
7120    Exit Sub

ERRH:
7130    intRetVal = 0
7140    Select Case ERR.Number
        Case Else
7150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7160    End Select
7170    Resume EXITP

End Sub

Private Sub lbxHistoryDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxHistoryDate_MouseMove"

        Dim strControl As String

7210    With Me
7220  On Error Resume Next
7230      strControl = Screen.ActiveControl.Name
7240  On Error GoTo ERRH
7250      If strControl <> vbNullString Then
7260        If strControl <> "lbxHistoryDate" Then
7270          .lbxHistoryDate.SetFocus
7280        End If
7290      End If
7300    End With

EXITP:
7310    Exit Sub

ERRH:
7320    Select Case ERR.Number
        Case Else
7330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7340    End Select
7350    Resume EXITP

End Sub

Private Sub chkActiveOnly_Dates_AfterUpdate()

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkActiveOnly_Dates_AfterUpdate"

        Dim strFilter As String
        Dim intPos01 As Integer

7410    With Me
7420      Select Case .chkActiveOnly_Dates
          Case True
7430        .chkActiveOnly_Dates_lbl.FontBold = True
7440        strFilter = .frmCurrency_History_Sub.Form.Filter
7450        intPos01 = InStr(strFilter, "And [curr_active] = True")
7460        If intPos01 = 0 Then
7470          strFilter = strFilter & " And [curr_active] = True"
7480        End If
7490      Case False
7500        .chkActiveOnly_Dates_lbl.FontBold = False
7510        strFilter = .frmCurrency_History_Sub.Form.Filter
7520        intPos01 = InStr(strFilter, "And [curr_active] = True")
7530        If intPos01 > 0 Then
7540          strFilter = Trim(Left(strFilter, (intPos01 - 1)))
7550        End If
7560      End Select
7570      .frmCurrency_History_Sub.Form.Filter = strFilter
7580      .frmCurrency_History_Sub.Form.FilterOn = True
7590      .frmCurrency_History_Sub.Form.Refresh
7600      blnFromActiveOnly = True
7610      If .Sort_Date_lbl_desc_img.Visible = True Then blnIsDesc = True Else blnIsDesc = False
7620      SortList "Sort_Date_lbl_cmd_DblClick"  ' ** Procedure: Below.
7630      blnFromActiveOnly = False
7640    End With

EXITP:
7650    Exit Sub

ERRH:
7660    Select Case ERR.Number
        Case Else
7670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7680    End Select
7690    Resume EXITP

End Sub

Private Sub chkActiveOnly_Dates_KeyDown(KeyCode As Integer, Shift As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkActiveOnly_Dates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
7710    intShiftDown = (Shift And acShiftMask) > 0
7720    intAltDown = (Shift And acAltMask) > 0
7730    intCtrlDown = (Shift And acCtrlMask) > 0

7740    intRetVal = KeyCode

        ' ** Plain keys.
7750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7770        With Me
7780          intRetVal = 0
7790          .lbxHistoryRate.SetFocus
7800        End With
7810      End Select
7820    End If

        ' ** Shift keys.
7830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7850        With Me
7860          intRetVal = 0
7870          .lbxHistoryDate.SetFocus
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

Private Sub Sort_Code_lbl_cmd_DblClick(Cancel As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "Sort_Code_lbl_cmd_DblClick"

8010    SortList THIS_PROC  ' ** Procedure: Below.

EXITP:
8020    Exit Sub

ERRH:
8030    Select Case ERR.Number
        Case Else
8040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8050    End Select
8060    Resume EXITP

End Sub

Private Sub Sort_Name_lbl_cmd_DblClick(Cancel As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "Sort_Name_lbl_cmd_DblClick"

8110    SortList THIS_PROC  ' ** Procedure: Below.

EXITP:
8120    Exit Sub

ERRH:
8130    Select Case ERR.Number
        Case Else
8140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8150    End Select
8160    Resume EXITP

End Sub

Private Sub lbxHistoryRate_GotFocus()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxHistoryRate_GotFocus"

8210    With Me
8220      .lbxHistoryRate.ForeColor = CLR_BLK
8230      .lbxHistoryRate.BackColor = CLR_WHT
8240      .lbxHistoryDate.ForeColor = CLR_DKGRY   ' ** 92/92/92
8250      .lbxHistoryDate.BackColor = 16448250  ' ** 250/250/250
8260    End With

EXITP:
8270    Exit Sub

ERRH:
8280    Select Case ERR.Number
        Case Else
8290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8300    End Select
8310    Resume EXITP

End Sub

Private Sub lbxHistoryRate_AfterUpdate()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxHistoryRate_AfterUpdate"

        Dim strFilter As String

8410    With Me
8420      If IsNull(.lbxHistoryRate) = False Then
8430        .lbxHistoryDate = Null
8440        strFilter = "[curr_id] = " & CStr(.lbxHistoryRate.Column(LBX_RT_CURID))
8450        If .chkActiveOnly_Rates = True Then
8460          strFilter = strFilter & " And [curr_active] = True"
8470        End If
8480        .frmCurrency_History_Sub.Form.Filter = strFilter
8490        .frmCurrency_History_Sub.Form.FilterOn = True
8500        DoEvents
8510        .frmCurrency_History_Sub.Form.source = .lbxHistoryRate.Column(LBX_RT_CODE)
8520        DoEvents
8530        .frmCurrency_History_Sub.Form.Refresh
8540        DoEvents
8550      End If
8560    End With

EXITP:
8570    Exit Sub

ERRH:
8580    Select Case ERR.Number
        Case Else
8590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8600    End Select
8610    Resume EXITP

End Sub

Private Sub lbxHistoryRate_KeyDown(KeyCode As Integer, Shift As Integer)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxHistoryRate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
8710    intShiftDown = (Shift And acShiftMask) > 0
8720    intAltDown = (Shift And acAltMask) > 0
8730    intCtrlDown = (Shift And acCtrlMask) > 0

8740    intRetVal = KeyCode

        ' ** Plain keys.
8750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8760      Select Case intRetVal
          Case vbKeyTab
8770        With Me
8780          intRetVal = 0
8790          .chkActiveOnly_Rates.SetFocus
8800        End With
8810      End Select
8820    End If

        ' ** Shift keys.
8830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8840      Select Case intRetVal
          Case vbKeyTab
8850        With Me
8860          intRetVal = 0
8870          .chkActiveOnly_Dates.SetFocus
8880        End With
8890      End Select
8900    End If

EXITP:
8910    KeyCode = intRetVal
8920    Exit Sub

ERRH:
8930    intRetVal = 0
8940    Select Case ERR.Number
        Case Else
8950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8960    End Select
8970    Resume EXITP

End Sub

Private Sub lbxHistoryRate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxHistoryRate_MouseMove"

        Dim strControl As String

9010    With Me
9020  On Error Resume Next
9030      strControl = Screen.ActiveControl.Name
9040  On Error GoTo ERRH
9050      If strControl <> vbNullString Then
9060        If strControl <> "lbxHistoryRate" Then
9070          .lbxHistoryRate.SetFocus
9080        End If
9090      End If
9100    End With

EXITP:
9110    Exit Sub

ERRH:
9120    Select Case ERR.Number
        Case Else
9130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9140    End Select
9150    Resume EXITP

End Sub

Private Sub chkActiveOnly_Rates_AfterUpdate()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkActiveOnly_Rates_AfterUpdate"

        Dim strFilter As String
        Dim intPos01 As Integer

9210    With Me
9220      Select Case .chkActiveOnly_Rates
          Case True
9230        .chkActiveOnly_Rates_lbl.FontBold = True
            ' ** tblCurrency_History, grouped by curr_id, with cnt, for curr_active = True.
9240        If .lbxHistoryRate.RowSource <> "qryCurrency_History_03_02" Then
9250          .lbxHistoryRate.RowSource = "qryCurrency_History_03_02"
9260        End If
9270        strFilter = .frmCurrency_History_Sub.Form.Filter
9280        intPos01 = InStr(strFilter, "And [curr_active] = True")
9290        If intPos01 = 0 Then
9300          strFilter = strFilter & " And [curr_active] = True"
9310        End If
9320      Case False
9330        .chkActiveOnly_Rates_lbl.FontBold = False
            ' ** tblCurrency_History, grouped by curr_id, with cnt.
9340        If .lbxHistoryRate.RowSource <> "qryCurrency_History_03_01" Then
9350          .lbxHistoryRate.RowSource = "qryCurrency_History_03_01"
9360        End If
9370        strFilter = .frmCurrency_History_Sub.Form.Filter
9380        intPos01 = InStr(strFilter, "And [curr_active] = True")
9390        If intPos01 > 0 Then
9400          strFilter = Trim(Left(strFilter, (intPos01 - 1)))
9410        End If
9420      End Select
9430      .frmCurrency_History_Sub.Form.Filter = strFilter
9440      .frmCurrency_History_Sub.Form.FilterOn = True
9450      .frmCurrency_History_Sub.Form.Refresh
9460      blnFromActiveOnly = True
9470      If .Sort_Code_line1.Visible = True Then
9480        If .Sort_Code_lbl_desc_img.Visible = True Then blnIsDesc = True Else blnIsDesc = False
9490        SortList "Sort_Code_lbl_cmd_DblClick"  ' ** Procedure: Below.
9500      ElseIf .Sort_Name_line1.Visible = True Then
9510        If .Sort_Name_lbl_desc_img.Visible = True Then blnIsDesc = True Else blnIsDesc = False
9520        SortList "Sort_Name_lbl_cmd_DblClick"  ' ** Procedure: Below.
9530      End If
9540      blnFromActiveOnly = False
9550    End With

EXITP:
9560    Exit Sub

ERRH:
9570    Select Case ERR.Number
        Case Else
9580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9590    End Select
9600    Resume EXITP

End Sub

Private Sub chkActiveOnly_Rates_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkActiveOnly_Rates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
9710    intShiftDown = (Shift And acShiftMask) > 0
9720    intAltDown = (Shift And acAltMask) > 0
9730    intCtrlDown = (Shift And acCtrlMask) > 0

9740    intRetVal = KeyCode

        ' ** Plain keys.
9750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9760      Select Case intRetVal
          Case vbKeyTab
9770        With Me
9780          intRetVal = 0
9790          lngRecsCur = .frmCurrency_History_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_History_Sub.
9800          If lngRecsCur > 0& Then
9810            .frmCurrency_History_Sub.SetFocus
9820            .frmCurrency_History_Sub.Form.FocusHolder.SetFocus
9830          ElseIf .cmdPrintReport.Enabled = True Then
9840            .cmdPrintReport.SetFocus
9850          Else
9860            .cmdClose.SetFocus
9870          End If
9880        End With
9890      End Select
9900    End If

        ' ** Shift keys.
9910    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9920      Select Case intRetVal
          Case vbKeyTab
9930        With Me
9940          intRetVal = 0
9950          .lbxHistoryRate.SetFocus
9960        End With
9970      End Select
9980    End If

EXITP:
9990    KeyCode = intRetVal
10000   Exit Sub

ERRH:
10010   intRetVal = 0
10020   Select Case ERR.Number
        Case Else
10030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10040   End Select
10050   Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

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

Private Sub SortList(strProc As String)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "SortList"

        Dim strCalled As String
        Dim intPos01 As Integer

        'THIS IS WAY MORE COMPLEX THAN IT HAS TO BE!!!!!!!!
        'I JUST COULDN'T SEE THE SIMPLIFYING PATTERNS!!

10210   With Me
10220     intPos01 = InStr(strProc, "_lbl_cmd")
10230     strCalled = Left(strProc, (intPos01 - 1))
10240     Select Case strCalled
          Case "Sort_Date"
10250       Select Case .chkActiveOnly_Dates
            Case True
10260         Select Case blnFromActiveOnly
              Case True
10270           Select Case blnIsDesc
                Case True
                  ' ** tblCurrency_History, grouped by curr_date, with cnt; Date, Active, Desc.
10280             .lbxHistoryDate.RowSource = "qryCurrency_History_02_04"
10290             .Sort_Date_lbl_desc_img.Visible = True
10300             .Sort_Date_lbl_asc_img.Visible = False
10310           Case False
                  ' ** tblCurrency_History, grouped by curr_date, with cnt; Date, Active, Asc.
10320             .lbxHistoryDate.RowSource = "qryCurrency_History_02_03"
10330             .Sort_Date_lbl_asc_img.Visible = True
10340             .Sort_Date_lbl_desc_img.Visible = False
10350           End Select
10360         Case False
10370           If ((.lbxHistoryDate.RowSource <> "qryCurrency_History_02_03") And (.lbxHistoryDate.RowSource <> "qryCurrency_History_02_04")) Or _
                    ((.lbxHistoryDate.RowSource = "qryCurrency_History_02_04") And (blnIsOpen = False)) Then
                  ' ** tblCurrency_History, grouped by curr_date, with cnt; Date, Active, Asc.
10380             .lbxHistoryDate.RowSource = "qryCurrency_History_02_03"
10390             .Sort_Date_lbl_asc_img.Visible = True
10400             .Sort_Date_lbl_desc_img.Visible = False
10410           Else
                  ' ** tblCurrency_History, grouped by curr_date, with cnt; Date, Active, Desc.
10420             .lbxHistoryDate.RowSource = "qryCurrency_History_02_04"
10430             .Sort_Date_lbl_desc_img.Visible = True
10440             .Sort_Date_lbl_asc_img.Visible = False
10450           End If
10460         End Select
10470       Case False
10480         Select Case blnFromActiveOnly
              Case True
10490           Select Case blnIsDesc
                Case True
                  ' ** tblCurrency_History, grouped by curr_date, with cnt; Date, All, Desc.
10500             .lbxHistoryDate.RowSource = "qryCurrency_History_02_02"
10510             .Sort_Date_lbl_desc_img.Visible = True
10520             .Sort_Date_lbl_asc_img.Visible = False
10530           Case False
                  ' ** tblCurrency_History, grouped by curr_date, with cnt; Date, All, Asc.
10540             .lbxHistoryDate.RowSource = "qryCurrency_History_02_01"
10550             .Sort_Date_lbl_asc_img.Visible = True
10560             .Sort_Date_lbl_desc_img.Visible = False
10570           End Select
10580         Case False
10590           If ((.lbxHistoryDate.RowSource <> "qryCurrency_History_02_01") And (.lbxHistoryDate.RowSource <> "qryCurrency_History_02_02")) Or _
                    ((.lbxHistoryDate.RowSource = "qryCurrency_History_02_02") And (blnIsOpen = False)) Then
                  ' ** tblCurrency_History, grouped by curr_date, with cnt; Date, All, Asc.
10600             .lbxHistoryDate.RowSource = "qryCurrency_History_02_01"
10610             .Sort_Date_lbl_asc_img.Visible = True
10620             .Sort_Date_lbl_desc_img.Visible = False
10630           Else
                  ' ** tblCurrency_History, grouped by curr_date, with cnt; Date, All, Desc.
10640             .lbxHistoryDate.RowSource = "qryCurrency_History_02_02"
10650             .Sort_Date_lbl_desc_img.Visible = True
10660             .Sort_Date_lbl_asc_img.Visible = False
10670           End If
10680         End Select
10690       End Select
10700     Case "Sort_Code"
10710       .Sort_Name_line1.Visible = False
10720       .Sort_Name_line2.Visible = False
10730       .Sort_Name_lbl_asc_img.Visible = False
10740       .Sort_Name_lbl_desc_img.Visible = False
10750       .Sort_Code_line1.Visible = True
10760       .Sort_Code_line2.Visible = True
10770       Select Case .chkActiveOnly_Rates
            Case True
10780         Select Case blnFromActiveOnly
              Case True
10790           Select Case blnIsDesc
                Case True
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Code, Active, Desc.
10800             .lbxHistoryRate.RowSource = "qryCurrency_History_03_04"
10810             .Sort_Code_lbl_desc_img.Visible = True
10820             .Sort_Code_lbl_asc_img.Visible = False
10830           Case False
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Code, Active, Asc.
10840             .lbxHistoryRate.RowSource = "qryCurrency_History_03_03"
10850             .Sort_Code_lbl_asc_img.Visible = True
10860             .Sort_Code_lbl_desc_img.Visible = False
10870           End Select
10880         Case False
10890           If ((.lbxHistoryRate.RowSource <> "qryCurrency_History_03_03") And (.lbxHistoryRate.RowSource <> "qryCurrency_History_03_04")) Or _
                    ((.lbxHistoryRate.RowSource = "qryCurrency_History_03_03") And (blnIsOpen = True)) Or _
                    ((.lbxHistoryRate.RowSource = "qryCurrency_History_03_04") And (blnIsOpen = False)) Then
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Code, Active, Asc.
10900             .lbxHistoryRate.RowSource = "qryCurrency_History_03_03"
10910             .Sort_Code_lbl_asc_img.Visible = True
10920             .Sort_Code_lbl_desc_img.Visible = False
10930           Else
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Code, Active, Desc.
10940             .lbxHistoryRate.RowSource = "qryCurrency_History_03_04"
10950             .Sort_Code_lbl_desc_img.Visible = True
10960             .Sort_Code_lbl_asc_img.Visible = False
10970           End If
10980         End Select
10990       Case False
11000         Select Case blnFromActiveOnly
              Case True
11010           Select Case blnIsDesc
                Case True
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Code, All, Desc.
11020             .lbxHistoryRate.RowSource = "qryCurrency_History_03_02"
11030             .Sort_Code_lbl_desc_img.Visible = True
11040             .Sort_Code_lbl_asc_img.Visible = False
11050           Case False
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Code, All, Asc.
11060             .lbxHistoryRate.RowSource = "qryCurrency_History_03_01"
11070             .Sort_Code_lbl_asc_img.Visible = True
11080             .Sort_Code_lbl_desc_img.Visible = False
11090           End Select
11100         Case False
11110           If ((.lbxHistoryRate.RowSource <> "qryCurrency_History_03_01") And (.lbxHistoryRate.RowSource = "qryCurrency_History_03_02")) Or _
                    ((.lbxHistoryRate.RowSource = "qryCurrency_History_03_01") And (blnIsOpen = True)) Or _
                    ((.lbxHistoryRate.RowSource = "qryCurrency_History_03_02") And (blnIsOpen = False)) Then
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Code, All, Asc.
11120             .lbxHistoryRate.RowSource = "qryCurrency_History_03_01"
11130             .Sort_Code_lbl_asc_img.Visible = True
11140             .Sort_Code_lbl_desc_img.Visible = False
11150           Else
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Code, All, Desc.
11160             .lbxHistoryRate.RowSource = "qryCurrency_History_03_02"
11170             .Sort_Code_lbl_desc_img.Visible = True
11180             .Sort_Code_lbl_asc_img.Visible = False
11190           End If
11200         End Select
11210       End Select
11220     Case "Sort_Name"
11230       .Sort_Code_line1.Visible = False
11240       .Sort_Code_line2.Visible = False
11250       .Sort_Code_lbl_asc_img.Visible = False
11260       .Sort_Code_lbl_desc_img.Visible = False
11270       .Sort_Name_line1.Visible = True
11280       .Sort_Name_line2.Visible = True
11290       Select Case .chkActiveOnly_Rates
            Case True
11300         Select Case blnFromActiveOnly
              Case True
11310           Select Case blnIsDesc
                Case True
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Name, Active, Desc.
11320             .lbxHistoryRate.RowSource = "qryCurrency_History_03_08"
11330             .Sort_Name_lbl_desc_img.Visible = True
11340             .Sort_Name_lbl_asc_img.Visible = False
11350           Case False
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Name, Active, Asc.
11360             .lbxHistoryRate.RowSource = "qryCurrency_History_03_07"
11370             .Sort_Name_lbl_asc_img.Visible = True
11380             .Sort_Name_lbl_desc_img.Visible = False
11390           End Select
11400         Case False
11410           If ((.lbxHistoryRate.RowSource <> "qryCurrency_History_03_07") And (.lbxHistoryRate.RowSource <> "qryCurrency_History_03_08")) Or _
                    ((.lbxHistoryRate.RowSource = "qryCurrency_History_03_07") And (blnIsOpen = True)) Or _
                    ((.lbxHistoryRate.RowSource = "qryCurrency_History_03_08") And (blnIsOpen = False)) Then
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Name, Active, Asc.
11420             .lbxHistoryRate.RowSource = "qryCurrency_History_03_07"
11430             .Sort_Name_lbl_asc_img.Visible = True
11440             .Sort_Name_lbl_desc_img.Visible = False
11450           Else
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Name, Active, Desc.
11460             .lbxHistoryRate.RowSource = "qryCurrency_History_03_08"
11470             .Sort_Name_lbl_desc_img.Visible = True
11480             .Sort_Name_lbl_asc_img.Visible = False
11490           End If
11500         End Select
11510       Case False
11520         Select Case blnFromActiveOnly
              Case True
11530           Select Case blnIsDesc
                Case True
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Name, All, Desc.
11540             .lbxHistoryRate.RowSource = "qryCurrency_History_03_06"
11550             .Sort_Name_lbl_desc_img.Visible = True
11560             .Sort_Name_lbl_asc_img.Visible = False
11570           Case False
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Name, All, Asc.
11580             .lbxHistoryRate.RowSource = "qryCurrency_History_03_05"
11590             .Sort_Name_lbl_asc_img.Visible = True
11600             .Sort_Name_lbl_desc_img.Visible = False
11610           End Select
11620         Case False
11630           If ((.lbxHistoryRate.RowSource <> "qryCurrency_History_03_05") And (.lbxHistoryRate.RowSource <> "qryCurrency_History_03_06")) Or _
                    ((.lbxHistoryRate.RowSource = "qryCurrency_History_03_05") And (blnIsOpen = True)) Or _
                    ((.lbxHistoryRate.RowSource = "qryCurrency_History_03_06") And (blnIsOpen = False)) Then
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Name, All, Asc.
11640             .lbxHistoryRate.RowSource = "qryCurrency_History_03_05"
11650             .Sort_Name_lbl_asc_img.Visible = True
11660             .Sort_Name_lbl_desc_img.Visible = False
11670           Else
                  ' ** tblCurrency_History, grouped by curr_id, with cnt; Name, All, Desc.
11680             .lbxHistoryRate.RowSource = "qryCurrency_History_03_06"
11690             .Sort_Name_lbl_desc_img.Visible = True
11700             .Sort_Name_lbl_asc_img.Visible = False
11710           End If
11720         End Select
11730       End Select
11740     End Select
11750   End With

EXITP:
11760   Exit Sub

ERRH:
11770   Select Case ERR.Number
        Case Else
11780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11790   End Select
11800   Resume EXITP

End Sub
