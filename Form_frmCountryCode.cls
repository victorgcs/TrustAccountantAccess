VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCountryCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCountryCode"

'VGC 09/23/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   AdjustFormWidth()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Active Only:      A {country_active on frmCountryCode_Sub}
' **   Close:            C {cmdClose}
' **   Misc Exclude:     D {opgMisc_optExclude}
' **   Active Exclude:   E {opgActive_optExclude}
' **   Active Include:   I {opgActive_optInclude}
' **   Misc Only:        L {opgMisc_optOnly}
' **   Name:             M {country_name on frmCountryCode_Sub}
' **   Currencies:       N {cmdCurrencies}
' **   Active Only:      O {opgActive_optOnly}
' **   Print Report:     P {cmdPrintReport}
' **   Misc Include:     U {opgMisc_optInclude}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmCountryCode_Sub}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   2-Ltr Exclude:    E {opgTwoLetter_optExclude}
' **   2-Ltr Include:    I {opgTwoLetter_optInclude}

' ** Countries with multiple currency listings.
' **   Show currency beneath? With list box for multiples?

Private Const FAND As String = " And "  ' ** Filter And.

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean
Private blnCurrencies_Focus As Boolean, blnCurrencies_MouseDown As Boolean

Private lngFrm_Top As Long, lngFrm_Left As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngSub_Width As Long, lngClose_Left As Long

Private blnGoingToReport As Boolean, lngPrintReportLbl_Left As Long
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

170           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
200           blnCurrencies_Focus = False: blnCurrencies_MouseDown = False

210           .form_width_line.Left = 0&  ' ** Just get it out of the way.
220           lngForm_Width = .Width
230           lngSub_Width = .frmCountryCode_Sub.Width
240           lngClose_Left = .cmdClose.Left

              ' ** Variables are fed empty, then populated ByRef.
250           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

260           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

270           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
280           If lngMonitorCnt > 1& Then
290             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
300             lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
310             If lngMonitorNum > 1& Then
320               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
330             End If
340           End If

350           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

360           opgActive_AfterUpdate  ' ** Procedure: Below.
370           opgMisc_AfterUpdate  ' ** Procedure: Below.
380           opgTwoLetter_AfterUpdate  ' ** Procedure: Below.

390           .frmCountryCode_Sub.SetFocus

400           blnIsOpen = False

410         Else
420           Cancel = -1
430         End If
440       Else
450         Cancel = -1
460       End If
470     End With

480     DoCmd.Hourglass False

490     If Cancel = -1 Then
500       Beep
510       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
520     End If

EXITP:
530     Exit Sub

ERRH:
540     DoCmd.Hourglass False
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Active:           A {country_active on frmCountryCode_Sub}
        ' **   Name:             M {country_name on frmCountryCode_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmCountryCode_Sub}

        ' ** Shortcut Ctrl-Shift to other forms and functionality:
        ' **   2-Ltr Exclude:    E {opgTwoLetter_optExclude}
        ' **   2-Ltr Include:    I {opgTwoLetter_optInclude}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyEscape
670         intRetVal = 0
680         cmdClose_Click  ' ** Procedure: Below.
690       End Select
700     End If

        ' ** Alt keys.
710     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
720       Select Case intRetVal
          Case vbKeyA
730         With Me
740           intRetVal = 0
750           .frmCountryCode_Sub.SetFocus
760           .frmCountryCode_Sub.Form.country_active.SetFocus
770           .frmCountryCode_Sub.Form.country_active = (Not .frmCountryCode_Sub.Form.country_active)
780           .frmCountryCode_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmCountryCode_Sub.
790         End With
800       Case vbKeyM
810         With Me
820           intRetVal = 0
830           .frmCountryCode_Sub.SetFocus
840           .frmCountryCode_Sub.Form.country_name.SetFocus
850         End With
860       Case vbKeyX
870         intRetVal = 0
880         cmdClose_Click  ' ** Procedure: Below.
890       End Select
900     End If

        ' ** Ctrl keys.
910     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
920       Select Case intRetVal
          Case vbKeyS
930         With Me
940           intRetVal = 0
950           .frmCountryCode_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmCountryCode_Sub.
960         End With
970       End Select
980     End If

        ' ** Ctrl-Shift keys.
990     If intCtrlDown And (Not intAltDown) And intShiftDown Then
1000      Select Case intRetVal
          Case vbKeyE
1010        With Me
1020          intRetVal = 0
1030          .opgTwoLetter = .opgTwoLetter_optExclude.OptionValue
1040          .opgTwoLetter.SetFocus
1050          opgTwoLetter_AfterUpdate  ' ** Procedure: Below.
1060        End With
1070      Case vbKeyI
1080        With Me
1090          intRetVal = 0
1100          .opgTwoLetter = .opgTwoLetter_optInclude.OptionValue
1110          .opgTwoLetter.SetFocus
1120          opgTwoLetter_AfterUpdate  ' ** Procedure: Below.
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
1370        Do Until strFormName = THIS_NAME
1380          DoCmd.SelectObject acForm, THIS_NAME, False
1390          strFormName = Screen.ActiveForm.Name
1400        Loop
1410  On Error GoTo ERRH
1420        .frmCountryCode_Sub.SetFocus
1430      End If
1440      If gblnGoToReport = True Then
1450        Select Case blnGoingToReport
            Case True
1460          DoCmd.Hourglass False
1470          gblnGoToReport = False
1480          blnGoingToReport = False
1490          .GoToReport_arw_pr_img.Visible = False
1500          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1510          DoEvents
1520        Case False
1530          DoCmd.Hourglass True  ' ** Make sure it's still running.
1540          DoEvents
1550          blnGoingToReport = True
1560          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1570          .GoToReport_arw_pr_img.Visible = True
1580          DoEvents
1590          Beep
1600          DoCmd.Hourglass False
1610          DoEvents
1620          .TimerInterval = (GTR_WAIT * 2&)
1630        End Select
1640      End If
1650      DoCmd.Hourglass False
1660    End With

EXITP:
1670    Exit Sub

ERRH:
1680    DoCmd.Hourglass False
1690    Select Case ERR.Number
        Case Else
1700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1710    End Select
1720    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String
        Dim blnRetVal As Boolean

1810    With Me

1820      DoCmd.Hourglass True
1830      DoEvents

1840      strDocName = "frmCountry_Currency"
1850      If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
1860        DoCmd.Close acForm, strDocName
1870        DoEvents
1880      End If

          ' ** Make sure all reports are closed.
1890      Do While Reports.Count > 0
1900        DoCmd.Close acReport, Reports(0).Name
1910        DoEvents
1920      Loop

1930      Set dbs = CurrentDb
          ' ** tblCountry, linked to tblCountry_Staging, with country_active_new.
1940      Set qdf = dbs.QueryDefs("qryCountryCode_02")
1950      Set rst = qdf.OpenRecordset
1960      If rst.BOF = True And rst.EOF = True Then
            ' ** No changes.
1970        rst.Close
1980        Set rst = Nothing
1990        Set qdf = Nothing
2000      Else
2010        rst.Close
2020        Set rst = Nothing
2030        Set qdf = Nothing
            ' ** Update qryCountryCode_02 (tblCountry, linked to tblCountry_Staging, with country_active_new).
2040        Set qdf = dbs.QueryDefs("qryCountryCode_03")
2050        qdf.Execute
2060        Set qdf = Nothing
            ' ** Empty tblCountry_Staging.
2070        Set qdf = dbs.QueryDefs("qryCountryCode_Sub_04")
2080        qdf.Execute
2090        Set qdf = Nothing
2100      End If
2110      dbs.Close
2120      Set dbs = Nothing

2130      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2140      Select Case strCallingForm
          Case "frmMenu_ForeignExchange"
2150        gstrReturningForm = THIS_NAME
2160        gblnSetFocus = True
2170        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Maintenance"
2180      End Select

2190    End With

EXITP:
2200    Set rst = Nothing
2210    Set qdf = Nothing
2220    Set dbs = Nothing
2230    Exit Sub

ERRH:
2240    DoCmd.Hourglass False
2250    Select Case ERR.Number
        Case Else
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

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
2450      If .cmdCurrencies_raised_focus_dots_img.Visible = True Or .cmdCurrencies_raised_focus_img.Visible = True Then
2460        Select Case blnCurrencies_Focus
            Case True
2470          .cmdCurrencies_raised_semifocus_dots_img.Visible = True
2480          .cmdCurrencies_raised_img.Visible = False
2490        Case False
2500          .cmdCurrencies_raised_img.Visible = True
2510          .cmdCurrencies_raised_semifocus_dots_img.Visible = False
2520        End Select
2530        .cmdCurrencies_raised_focus_img.Visible = False
2540        .cmdCurrencies_raised_focus_dots_img.Visible = False
2550        .cmdCurrencies_sunken_focus_dots_img.Visible = False
2560        .cmdCurrencies_raised_img_dis.Visible = False
2570      End If
2580    End With

EXITP:
2590    Exit Sub

ERRH:
2600    Select Case ERR.Number
        Case Else
2610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2620    End Select
2630    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2720    Exit Sub

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2810    DoCmd.Close acForm, THIS_NAME

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
2970        With Me
2980          intRetVal = 0
2990          .frmCountryCode_Sub.SetFocus
3000          .frmCountryCode_Sub.Form.FocusHolder.SetFocus
3010        End With
3020      End Select
3030    End If

        ' ** Shift keys.
3040    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3050      Select Case intRetVal
          Case vbKeyTab
3060        With Me
3070          intRetVal = 0
3080          .opgTwoLetter.SetFocus
3090        End With
3100      End Select
3110    End If

EXITP:
3120    KeyCode = intRetVal
3130    Exit Sub

ERRH:
3140    intRetVal = 0
3150    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3160    Case Else
3170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3180    End Select
3190    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

3210    With Me
3220      blnPrintReport_Focus = True
3230      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3240      .cmdPrintReport_raised_img.Visible = False
3250      .cmdPrintReport_raised_focus_img.Visible = False
3260      .cmdPrintReport_raised_focus_dots_img.Visible = False
3270      .cmdPrintReport_sunken_focus_dots_img.Visible = False
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

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

3410    With Me
3420      blnPrintReport_MouseDown = True
3430      .cmdPrintReport_sunken_focus_dots_img.Visible = True
3440      .cmdPrintReport_raised_img.Visible = False
3450      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3460      .cmdPrintReport_raised_focus_img.Visible = False
3470      .cmdPrintReport_raised_focus_dots_img.Visible = False
3480      .cmdPrintReport_raised_img_dis.Visible = False
3490    End With

EXITP:
3500    Exit Sub

ERRH:
3510    Select Case ERR.Number
        Case Else
3520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3530    End Select
3540    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String, strFilter As String

3610    With Me
3620      strFilter = .frmCountryCode_Sub.Form.Filter
3630      strDocName = "rptListOfCountryCodes"
3640      If gblnDev_Debug Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
3650        DoCmd.OpenReport strDocName, acViewPreview, , strFilter
3660        DoCmd.Maximize
3670        DoCmd.RunCommand acCmdFitToWindow
3680        .Visible = False
3690      Else
            '##GTR_Ref: rptListOfCountryCodes
3700        DoCmd.OpenReport strDocName, acViewNormal, , strFilter
3710      End If
3720    End With

EXITP:
3730    Exit Sub

ERRH:
3740    Select Case ERR.Number
        Case Else
3750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3760    End Select
3770    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
3890          .opgTwoLetter.SetFocus
3900        End With
3910      End Select
3920    End If

        ' ** Shift keys.
3930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3940      Select Case intRetVal
          Case vbKeyTab
3950        With Me
3960          intRetVal = 0
3970          .opgMisc.SetFocus
3980        End With
3990      End Select
4000    End If

EXITP:
4010    KeyCode = intRetVal
4020    Exit Sub

ERRH:
4030    intRetVal = 0
4040    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4050    Case Else
4060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4070    End Select
4080    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

4110    With Me
4120      If blnPrintReport_MouseDown = False Then
4130        Select Case blnPrintReport_Focus
            Case True
4140          .cmdPrintReport_raised_focus_dots_img.Visible = True
4150          .cmdPrintReport_raised_focus_img.Visible = False
4160        Case False
4170          .cmdPrintReport_raised_focus_img.Visible = True
4180          .cmdPrintReport_raised_focus_dots_img.Visible = False
4190        End Select
4200        .cmdPrintReport_raised_img.Visible = False
4210        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4220        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4230        .cmdPrintReport_raised_img_dis.Visible = False
4240      End If
4250    End With

EXITP:
4260    Exit Sub

ERRH:
4270    Select Case ERR.Number
        Case Else
4280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4290    End Select
4300    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

4410    With Me
4420      .cmdPrintReport_raised_focus_dots_img.Visible = True
4430      .cmdPrintReport_raised_img.Visible = False
4440      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4450      .cmdPrintReport_raised_focus_img.Visible = False
4460      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4470      .cmdPrintReport_raised_img_dis.Visible = False
4480      blnPrintReport_MouseDown = False
4490    End With

EXITP:
4500    Exit Sub

ERRH:
4510    Select Case ERR.Number
        Case Else
4520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4530    End Select
4540    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

4610    With Me
4620      .cmdPrintReport_raised_img.Visible = True
4630      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4640      .cmdPrintReport_raised_focus_img.Visible = False
4650      .cmdPrintReport_raised_focus_dots_img.Visible = False
4660      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4670      .cmdPrintReport_raised_img_dis.Visible = False
4680      blnPrintReport_Focus = False
4690    End With

EXITP:
4700    Exit Sub

ERRH:
4710    Select Case ERR.Number
        Case Else
4720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4730    End Select
4740    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

4810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4820    Exit Sub

ERRH:
4830    Select Case ERR.Number
        Case Else
4840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4850    End Select
4860    Resume EXITP

End Sub

Private Sub opgActive_AfterUpdate()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_AfterUpdate"

        Dim strFilter As String

4910    With Me

4920      Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
4930        .opgActive_optInclude_lbl.FontBold = True
4940        .opgActive_optExclude_lbl.FontBold = False
4950        .opgActive_optOnly_lbl.FontBold = False
4960        strFilter = vbNullString
4970      Case .opgActive_optExclude.OptionValue
4980        .opgActive_optInclude_lbl.FontBold = False
4990        .opgActive_optExclude_lbl.FontBold = True
5000        .opgActive_optOnly_lbl.FontBold = False
5010        strFilter = "[country_active] = False"
5020      Case .opgActive_optOnly.OptionValue
5030        .opgActive_optInclude_lbl.FontBold = False
5040        .opgActive_optExclude_lbl.FontBold = False
5050        .opgActive_optOnly_lbl.FontBold = True
5060        strFilter = "[country_active] = True"
5070      End Select

5080      Select Case .opgMisc
          Case .opgMisc_optInclude.OptionValue
            ' ** No change.
5090      Case .opgMisc_optExclude.OptionValue
5100        If strFilter <> vbNullString Then strFilter = strFilter & FAND
5110        strFilter = strFilter & "[country_na] = False"
5120      Case .opgMisc_optOnly.OptionValue
5130        If strFilter <> vbNullString Then strFilter = strFilter & FAND
5140        strFilter = strFilter & "[country_na] = True"
5150      End Select

5160      .frmCountryCode_Sub.Form.Filter = strFilter
5170      If strFilter = vbNullString Then
5180        .frmCountryCode_Sub.Form.FilterOn = False
5190      Else
5200        .frmCountryCode_Sub.Form.FilterOn = True
5210      End If
5220      DoEvents

5230      lngRecsCur = .frmCountryCode_Sub.Form.RecCnt  ' ** Form Function: frmCountryCode_Sub.
5240      If lngRecsCur > 0& Then
5250        .frmCountryCode_Sub.SetFocus
5260      End If

5270    End With

EXITP:
5280    Exit Sub

ERRH:
5290    Select Case ERR.Number
        Case Else
5300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5310    End Select
5320    Resume EXITP

End Sub

Private Sub opgActive_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5420    intShiftDown = (Shift And acShiftMask) > 0
5430    intAltDown = (Shift And acAltMask) > 0
5440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5470        With Me
5480          intRetVal = 0
5490          .opgMisc.SetFocus
5500        End With
5510      End Select
5520    End If

        ' ** Shift keys.
5530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5550        With Me
5560          intRetVal = 0
5570          .frmCountryCode_Sub.SetFocus
5580          .frmCountryCode_Sub.Form.FocusHolder.SetFocus
5590        End With
5600      End Select
5610    End If

EXITP:
5620    KeyCode = intRetVal
5630    Exit Sub

ERRH:
5640    intRetVal = 0
5650    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5660    Case Else
5670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5680    End Select
5690    Resume EXITP

End Sub

Private Sub opgActive_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5720    intShiftDown = (Shift And acShiftMask) > 0
5730    intAltDown = (Shift And acAltMask) > 0
5740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5770        With Me
5780          intRetVal = 0
5790          .opgMisc.SetFocus
5800        End With
5810      End Select
5820    End If

        ' ** Shift keys.
5830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5850        With Me
5860          intRetVal = 0
5870          .frmCountryCode_Sub.SetFocus
5880          .frmCountryCode_Sub.Form.FocusHolder.SetFocus
5890        End With
5900      End Select
5910    End If

EXITP:
5920    KeyCode = intRetVal
5930    Exit Sub

ERRH:
5940    intRetVal = 0
5950    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5960    Case Else
5970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5980    End Select
5990    Resume EXITP

End Sub

Private Sub opgActive_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optOnly_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
6070        With Me
6080          intRetVal = 0
6090          .opgMisc.SetFocus
6100        End With
6110      End Select
6120    End If

        ' ** Shift keys.
6130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6150        With Me
6160          intRetVal = 0
6170          .frmCountryCode_Sub.SetFocus
6180          .frmCountryCode_Sub.Form.FocusHolder.SetFocus
6190        End With
6200      End Select
6210    End If

EXITP:
6220    KeyCode = intRetVal
6230    Exit Sub

ERRH:
6240    intRetVal = 0
6250    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6260    Case Else
6270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6280    End Select
6290    Resume EXITP

End Sub

Private Sub opgActive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_MouseMove"

6310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Sub

Private Sub opgMisc_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgMisc_box_MouseMove"

6410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6420    Exit Sub

ERRH:
6430    Select Case ERR.Number
        Case Else
6440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6450    End Select
6460    Resume EXITP

End Sub

Private Sub opgMisc_AfterUpdate()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgMisc_AfterUpdate"

        Dim strFilter As String

6510    With Me

6520      Select Case .opgMisc
          Case .opgMisc_optInclude.OptionValue
6530        .opgMisc_optInclude_lbl.FontBold = True
6540        .opgMisc_optExclude_lbl.FontBold = False
6550        .opgMisc_optOnly_lbl.FontBold = False
6560        strFilter = vbNullString
6570        .frmCountryCode_Sub.Form.ShowMisc True  ' ** Form Procedure: frmCountryCode_Sub.
6580        AdjustFormWidth THIS_PROC  ' ** Procedure: Below.
6590      Case .opgMisc_optExclude.OptionValue
6600        .opgMisc_optInclude_lbl.FontBold = False
6610        .opgMisc_optExclude_lbl.FontBold = True
6620        .opgMisc_optOnly_lbl.FontBold = False
6630        strFilter = "[country_na] = False"
6640        If .frmCountryCode_Sub.Form.Sort_line.Left = .frmCountryCode_Sub.Form.country_na_lbl.Left Then
6650          .frmCountryCode_Sub.Form.SortNow "Form_Load"  ' ** Form Procedure: frmCountryCode_Sub.
6660        End If
6670        .frmCountryCode_Sub.Form.ShowMisc False  ' ** Form Procedure: frmCountryCode_Sub.
6680        AdjustFormWidth THIS_PROC  ' ** Procedure: Below.
6690      Case .opgMisc_optOnly.OptionValue
6700        .opgMisc_optInclude_lbl.FontBold = False
6710        .opgMisc_optExclude_lbl.FontBold = False
6720        .opgMisc_optOnly_lbl.FontBold = True
6730        strFilter = "[country_na] = True"
6740        .frmCountryCode_Sub.Form.ShowMisc True  ' ** Form Procedure: frmCountryCode_Sub.
6750        AdjustFormWidth THIS_PROC  ' ** Procedure: Below.
6760      End Select

6770      Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
            ' ** No change.
6780      Case .opgActive_optExclude.OptionValue
6790        If strFilter <> vbNullString Then strFilter = strFilter & FAND
6800        strFilter = strFilter & "[country_active] = False"
6810      Case .opgActive_optOnly.OptionValue
6820        If strFilter <> vbNullString Then strFilter = strFilter & FAND
6830        strFilter = strFilter & "[country_active] = True"
6840      End Select

6850      .frmCountryCode_Sub.Form.Filter = strFilter
6860      If strFilter = vbNullString Then
6870        .frmCountryCode_Sub.Form.FilterOn = False
6880      Else
6890        .frmCountryCode_Sub.Form.FilterOn = True
6900      End If
6910      DoEvents

6920      lngRecsCur = .frmCountryCode_Sub.Form.RecCnt  ' ** Form Function: frmCountryCode_Sub.
6930      If lngRecsCur > 0& Then
6940        .frmCountryCode_Sub.SetFocus
6950      End If

6960    End With

EXITP:
6970    Exit Sub

ERRH:
6980    Select Case ERR.Number
        Case Else
6990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7000    End Select
7010    Resume EXITP

End Sub

Private Sub opgMisc_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgMisc_optInclude_KeyDown"

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
7190          If .cmdPrintReport.Enabled = True Then
7200            .cmdPrintReport.SetFocus
7210          Else
7220            .opgTwoLetter.SetFocus
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
7310          .opgActive.SetFocus
7320        End With
7330      End Select
7340    End If

EXITP:
7350    KeyCode = intRetVal
7360    Exit Sub

ERRH:
7370    intRetVal = 0
7380    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7390    Case Else
7400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7410    End Select
7420    Resume EXITP

End Sub

Private Sub opgMisc_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgMisc_optExclude_KeyDown"

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
          Case vbKeyTab
7570        With Me
7580          intRetVal = 0
7590          If .cmdPrintReport.Enabled = True Then
7600            .cmdPrintReport.SetFocus
7610          Else
7620            .opgTwoLetter.SetFocus
7630          End If
7640        End With
7650      End Select
7660    End If

        ' ** Shift keys.
7670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7680      Select Case intRetVal
          Case vbKeyTab
7690        With Me
7700          intRetVal = 0
7710          .opgActive.SetFocus
7720        End With
7730      End Select
7740    End If

EXITP:
7750    KeyCode = intRetVal
7760    Exit Sub

ERRH:
7770    intRetVal = 0
7780    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7790    Case Else
7800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7810    End Select
7820    Resume EXITP

End Sub

Private Sub opgMisc_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgMisc_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7920    intShiftDown = (Shift And acShiftMask) > 0
7930    intAltDown = (Shift And acAltMask) > 0
7940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7960      Select Case intRetVal
          Case vbKeyTab
7970        With Me
7980          intRetVal = 0
7990          If .cmdPrintReport.Enabled = True Then
8000            .cmdPrintReport.SetFocus
8010          Else
8020            .opgTwoLetter.SetFocus
8030          End If
8040        End With
8050      End Select
8060    End If

        ' ** Shift keys.
8070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8080      Select Case intRetVal
          Case vbKeyTab
8090        With Me
8100          intRetVal = 0
8110          .opgActive.SetFocus
8120        End With
8130      End Select
8140    End If

EXITP:
8150    KeyCode = intRetVal
8160    Exit Sub

ERRH:
8170    intRetVal = 0
8180    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8190    Case Else
8200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8210    End Select
8220    Resume EXITP

End Sub

Private Sub opgMisc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgMisc_MouseMove"

8310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8320    Exit Sub

ERRH:
8330    Select Case ERR.Number
        Case Else
8340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8350    End Select
8360    Resume EXITP

End Sub

Private Sub currency_cnt_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "currency_cnt_box_MouseMove"

8410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8420    Exit Sub

ERRH:
8430    Select Case ERR.Number
        Case Else
8440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8450    End Select
8460    Resume EXITP

End Sub

Private Sub currency_cnt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "currency_cnt_MouseMove"

8510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8520    Exit Sub

ERRH:
8530    Select Case ERR.Number
        Case Else
8540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8550    End Select
8560    Resume EXITP

End Sub

Private Sub cmdCurrencies_GotFocus()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrencies_GotFocus"

8610    With Me
8620      blnCurrencies_Focus = True
8630      .cmdCurrencies_raised_semifocus_dots_img.Visible = True
8640      .cmdCurrencies_raised_img.Visible = False
8650      .cmdCurrencies_raised_focus_img.Visible = False
8660      .cmdCurrencies_raised_focus_dots_img.Visible = False
8670      .cmdCurrencies_sunken_focus_dots_img.Visible = False
8680      .cmdCurrencies_raised_img_dis.Visible = False
8690    End With

EXITP:
8700    Exit Sub

ERRH:
8710    Select Case ERR.Number
        Case Else
8720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8730    End Select
8740    Resume EXITP

End Sub

Private Sub cmdCurrencies_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrencies_MouseDown"

8810    With Me
8820      blnCurrencies_MouseDown = True
8830      .cmdCurrencies_sunken_focus_dots_img.Visible = True
8840      .cmdCurrencies_raised_img.Visible = False
8850      .cmdCurrencies_raised_semifocus_dots_img.Visible = False
8860      .cmdCurrencies_raised_focus_img.Visible = False
8870      .cmdCurrencies_raised_focus_dots_img.Visible = False
8880      .cmdCurrencies_raised_img_dis.Visible = False
8890    End With

EXITP:
8900    Exit Sub

ERRH:
8910    Select Case ERR.Number
        Case Else
8920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8930    End Select
8940    Resume EXITP

End Sub

Private Sub cmdCurrencies_Click()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrencies_Click"

        Dim strDocName As String

9010    strDocName = "frmCountry_Currency"
9020    Select Case IsLoaded(strDocName, acForm)
        Case True
9030      DoCmd.Close acForm, strDocName
9040    Case False
9050      gblnSetFocus = True
9060      DoCmd.OpenForm strDocName, , , , , , THIS_NAME
9070    End Select

EXITP:
9080    Exit Sub

ERRH:
9090    Select Case ERR.Number
        Case Else
9100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9110    End Select
9120    Resume EXITP

End Sub

Private Sub cmdCurrencies_KeyDown(KeyCode As Integer, Shift As Integer)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrencies_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9220    intShiftDown = (Shift And acShiftMask) > 0
9230    intAltDown = (Shift And acAltMask) > 0
9240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9260      Select Case intRetVal
          Case vbKeyTab
9270        With Me
9280          intRetVal = 0
9290          If .cmdPrintReport.Enabled = True Then
9300            .cmdPrintReport.SetFocus
9310          ElseIf .opgTwoLetter.Enabled = True Then
9320            .opgTwoLetter.SetFocus
9330          Else
9340            .cmdClose.SetFocus
9350          End If
9360        End With
9370      End Select
9380    End If

        ' ** Shift keys.
9390    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9400      Select Case intRetVal
          Case vbKeyTab
9410        With Me
9420          intRetVal = 0
9430          .opgMisc.SetFocus
9440        End With
9450      End Select
9460    End If

EXITP:
9470    KeyCode = intRetVal
9480    Exit Sub

ERRH:
9490    intRetVal = 0
9500    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
9510    Case Else
9520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9530    End Select
9540    Resume EXITP

End Sub

Private Sub cmdCurrencies_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrencies_MouseMove"

9610    With Me
9620      If blnCurrencies_MouseDown = False Then
9630        Select Case blnCurrencies_Focus
            Case True
9640          .cmdCurrencies_raised_focus_dots_img.Visible = True
9650          .cmdCurrencies_raised_focus_img.Visible = False
9660        Case False
9670          .cmdCurrencies_raised_focus_img.Visible = True
9680          .cmdCurrencies_raised_focus_dots_img.Visible = False
9690        End Select
9700        .cmdCurrencies_raised_img.Visible = False
9710        .cmdCurrencies_raised_semifocus_dots_img.Visible = False
9720        .cmdCurrencies_sunken_focus_dots_img.Visible = False
9730        .cmdCurrencies_raised_img_dis.Visible = False
9740      End If
9750    End With

EXITP:
9760    Exit Sub

ERRH:
9770    Select Case ERR.Number
        Case Else
9780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9790    End Select
9800    Resume EXITP

End Sub

Private Sub cmdCurrencies_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrencies_MouseUp"

9910    With Me
9920      .cmdCurrencies_raised_focus_dots_img.Visible = True
9930      .cmdCurrencies_raised_img.Visible = False
9940      .cmdCurrencies_raised_semifocus_dots_img.Visible = False
9950      .cmdCurrencies_raised_focus_img.Visible = False
9960      .cmdCurrencies_sunken_focus_dots_img.Visible = False
9970      .cmdCurrencies_raised_img_dis.Visible = False
9980      blnCurrencies_MouseDown = False
9990    End With

EXITP:
10000   Exit Sub

ERRH:
10010   Select Case ERR.Number
        Case Else
10020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10030   End Select
10040   Resume EXITP

End Sub

Private Sub cmdCurrencies_LostFocus()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrencies_LostFocus"

10110   With Me
10120     .cmdCurrencies_raised_img.Visible = True
10130     .cmdCurrencies_raised_semifocus_dots_img.Visible = False
10140     .cmdCurrencies_raised_focus_img.Visible = False
10150     .cmdCurrencies_raised_focus_dots_img.Visible = False
10160     .cmdCurrencies_sunken_focus_dots_img.Visible = False
10170     .cmdCurrencies_raised_img_dis.Visible = False
10180     blnCurrencies_Focus = False
10190   End With

EXITP:
10200   Exit Sub

ERRH:
10210   Select Case ERR.Number
        Case Else
10220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10230   End Select
10240   Resume EXITP

End Sub

Public Sub opgTwoLetter_AfterUpdate()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgTwoLetter_AfterUpdate"

10310   With Me
10320     Select Case .opgTwoLetter
          Case .opgTwoLetter_optInclude.OptionValue
10330       .opgTwoLetter_optInclude_lbl.FontBold = True
10340       .opgTwoLetter_optExclude_lbl.FontBold = False
10350       .frmCountryCode_Sub.Form.ShowCode2 True  ' ** Form Procedure: frmCountryCode_Sub.
10360       AdjustFormWidth THIS_PROC  ' ** Procedure: Below.
10370     Case .opgTwoLetter_optExclude.OptionValue
10380       .opgTwoLetter_optInclude_lbl.FontBold = False
10390       .opgTwoLetter_optExclude_lbl.FontBold = True
10400       .frmCountryCode_Sub.Form.ShowCode2 False  ' ** Form Procedure: frmCountryCode_Sub.
10410       AdjustFormWidth THIS_PROC  ' ** Procedure: Below.
10420     End Select
10430     .frmCountryCode_Sub.SetFocus
10440   End With

EXITP:
10450   Exit Sub

ERRH:
10460   Select Case ERR.Number
        Case Else
10470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10480   End Select
10490   Resume EXITP

End Sub

Private Sub opgTwoLetter_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgTwoLetter_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10520   intShiftDown = (Shift And acShiftMask) > 0
10530   intAltDown = (Shift And acAltMask) > 0
10540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10560     Select Case intRetVal
          Case vbKeyTab
10570       With Me
10580         intRetVal = 0
10590         .cmdClose.SetFocus
10600       End With
10610     End Select
10620   End If

        ' ** Shift keys.
10630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10640     Select Case intRetVal
          Case vbKeyTab
10650       With Me
10660         intRetVal = 0
10670         .opgMisc.SetFocus
10680       End With
10690     End Select
10700   End If

EXITP:
10710   KeyCode = intRetVal
10720   Exit Sub

ERRH:
10730   intRetVal = 0
10740   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
10750   Case Else
10760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10770   End Select
10780   Resume EXITP

End Sub

Private Sub opgTwoLetter_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgTwoLetter_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10820   intShiftDown = (Shift And acShiftMask) > 0
10830   intAltDown = (Shift And acAltMask) > 0
10840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10860     Select Case intRetVal
          Case vbKeyTab
10870       With Me
10880         intRetVal = 0
10890         .cmdClose.SetFocus
10900       End With
10910     End Select
10920   End If

        ' ** Shift keys.
10930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10940     Select Case intRetVal
          Case vbKeyTab
10950       With Me
10960         intRetVal = 0
10970         .opgMisc.SetFocus
10980       End With
10990     End Select
11000   End If

EXITP:
11010   KeyCode = intRetVal
11020   Exit Sub

ERRH:
11030   intRetVal = 0
11040   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
11050   Case Else
11060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11070   End Select
11080   Resume EXITP

End Sub

Private Sub AdjustFormWidth(strProc As String)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "AdjustFormWidth"

        Dim lngAddedWidth1 As Long, lngAddedWidth2 As Long
        Dim lngNewFrmWidth1 As Long, lngNewFrmWidth2 As Long, lngNewFrmWidth3 As Long, lngNewFrmWidth4 As Long
        Dim strTmp01 As String, lngTmp02 As Long

11110   With Me

11120     strTmp01 = Left(strProc, (Len(strProc) - Len("_AfterUpdate")))
11130     lngAddedWidth1 = (.frmCountryCode_Sub.Form.country_code2_lbl.Width + (4& * lngTpp))
11140     lngAddedWidth2 = (.frmCountryCode_Sub.Form.country_na_lbl.Width + (4& * lngTpp))
11150     lngNewFrmWidth1 = (lngForm_Width - lngAddedWidth2)                     ' ** No Misc, no Code2.
11160     lngNewFrmWidth2 = lngForm_Width                                        ' ** Misc, no Code2 (opening).
11170     lngNewFrmWidth3 = ((lngForm_Width - lngAddedWidth2) + lngAddedWidth1)  ' ** Code2, no Misc.
11180     lngNewFrmWidth4 = (lngForm_Width + lngAddedWidth1)                     ' ** Misc, Code2.

11190     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
11200     lngMonitorNum = 1&: lngTmp02 = 0&
11210     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
11220     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

11230     Select Case .opgMisc
          Case .opgMisc_optInclude.OptionValue, .opgMisc_optOnly.OptionValue
11240       Select Case .opgTwoLetter
            Case .opgTwoLetter_optInclude.OptionValue
11250         If .Width < lngNewFrmWidth4 Then
11260           .Width = lngNewFrmWidth4
11270         End If
11280         .cmdClose.Left = (lngClose_Left + lngAddedWidth1)
11290         .Header_vline01.Left = lngNewFrmWidth4
11300         .Header_vline02.Left = lngNewFrmWidth4
11310         .Detail_vline01.Left = lngNewFrmWidth4
11320         .Detail_vline02.Left = lngNewFrmWidth4
11330         .Footer_vline01.Left = lngNewFrmWidth4
11340         .Footer_vline02.Left = lngNewFrmWidth4
11350         .Header_hline01.Width = lngNewFrmWidth4
11360         .Header_hline02.Width = lngNewFrmWidth4
11370         .Detail_hline01.Width = lngNewFrmWidth4
11380         .Detail_hline02.Width = lngNewFrmWidth4
11390         .Footer_hline01.Width = lngNewFrmWidth4
11400         .Footer_hline02.Width = lngNewFrmWidth4
11410         .frmCountryCode_Sub.Width = (lngSub_Width + lngAddedWidth1)
11420         .frmCountryCode_Sub_box.Width = (.frmCountryCode_Sub.Width + (2& * lngTpp))
11430         .Nav_box01.Width = .frmCountryCode_Sub.Width
11440         If .Width > lngNewFrmWidth4 Then
11450           .Width = lngNewFrmWidth4
11460         End If
              ' ** I don't know why this offsets by 8 pixels.
11470         If lngMonitorNum = 1& Then lngTmp02 = (lngFrm_Top - (8& * lngTpp))
11480         DoCmd.MoveSize (lngFrm_Left - (8& * lngTpp)), lngTmp02, (lngFrm_Width + lngAddedWidth1), lngFrm_Height  '(lngFrm_Top - (8& * lngTpp))
11490         If lngMonitorNum > 1& Then
11500           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
11510         End If
11520       Case .opgTwoLetter_optExclude.OptionValue
11530         If .Width < lngNewFrmWidth2 Then
11540           .Width = lngNewFrmWidth2
11550         End If
11560         .cmdClose.Left = lngClose_Left
11570         .Header_vline01.Left = lngNewFrmWidth2
11580         .Header_vline02.Left = lngNewFrmWidth2
11590         .Detail_vline01.Left = lngNewFrmWidth2
11600         .Detail_vline02.Left = lngNewFrmWidth2
11610         .Footer_vline01.Left = lngNewFrmWidth2
11620         .Footer_vline02.Left = lngNewFrmWidth2
11630         .Header_hline01.Width = lngNewFrmWidth2
11640         .Header_hline02.Width = lngNewFrmWidth2
11650         .Detail_hline01.Width = lngNewFrmWidth2
11660         .Detail_hline02.Width = lngNewFrmWidth2
11670         .Footer_hline01.Width = lngNewFrmWidth2
11680         .Footer_hline02.Width = lngNewFrmWidth2
11690         .frmCountryCode_Sub.Width = lngSub_Width
11700         .frmCountryCode_Sub_box.Width = (.frmCountryCode_Sub.Width + (2& * lngTpp))
11710         .Nav_box01.Width = .frmCountryCode_Sub.Width
11720         If .Width > lngNewFrmWidth2 Then
11730           .Width = lngNewFrmWidth2
11740         End If
11750         If lngMonitorNum = 1& Then lngTmp02 = (lngFrm_Top - (8& * lngTpp))
11760         DoCmd.MoveSize (lngFrm_Left - (8& * lngTpp)), lngTmp02, lngFrm_Width, lngFrm_Height  '(lngFrm_Top - (8& * lngTpp))
11770         If lngMonitorNum > 1& Then
11780           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
11790         End If
11800       End Select
11810     Case .opgMisc_optExclude.OptionValue
11820       Select Case .opgTwoLetter
            Case .opgTwoLetter_optInclude.OptionValue
11830         If .Width < lngNewFrmWidth3 Then
11840           .Width = lngNewFrmWidth3
11850         End If
11860         .cmdClose.Left = ((lngClose_Left - lngAddedWidth2) + lngAddedWidth1)
11870         .Header_vline01.Left = lngNewFrmWidth3
11880         .Header_vline02.Left = lngNewFrmWidth3
11890         .Detail_vline01.Left = lngNewFrmWidth3
11900         .Detail_vline02.Left = lngNewFrmWidth3
11910         .Footer_vline01.Left = lngNewFrmWidth3
11920         .Footer_vline02.Left = lngNewFrmWidth3
11930         .Header_hline01.Width = lngNewFrmWidth3
11940         .Header_hline02.Width = lngNewFrmWidth3
11950         .Detail_hline01.Width = lngNewFrmWidth3
11960         .Detail_hline02.Width = lngNewFrmWidth3
11970         .Footer_hline01.Width = lngNewFrmWidth3
11980         .Footer_hline02.Width = lngNewFrmWidth3
11990         .frmCountryCode_Sub.Width = ((lngSub_Width - lngAddedWidth2) + lngAddedWidth1)
12000         .frmCountryCode_Sub_box.Width = (.frmCountryCode_Sub.Width + (2& * lngTpp))
12010         .Nav_box01.Width = .frmCountryCode_Sub.Width
12020         If .Width > lngNewFrmWidth3 Then
12030           .Width = lngNewFrmWidth3
12040         End If
12050         If lngMonitorNum = 1& Then lngTmp02 = (lngFrm_Top - (8& * lngTpp))
12060         DoCmd.MoveSize (lngFrm_Left - (8& * lngTpp)), lngTmp02, ((lngFrm_Width - lngAddedWidth2) + lngAddedWidth1), lngFrm_Height  '(lngFrm_Top - (8& * lngTpp))
12070         If lngMonitorNum > 1& Then
12080           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
12090         End If
12100       Case .opgTwoLetter_optExclude.OptionValue
12110         If .Width < lngNewFrmWidth1 Then  ' ** Unlikely.
12120           .Width = lngNewFrmWidth1
12130         End If
12140         .cmdClose.Left = (lngClose_Left - lngAddedWidth2)
12150         .Header_vline01.Left = lngNewFrmWidth1
12160         .Header_vline02.Left = lngNewFrmWidth1
12170         .Detail_vline01.Left = lngNewFrmWidth1
12180         .Detail_vline02.Left = lngNewFrmWidth1
12190         .Footer_vline01.Left = lngNewFrmWidth1
12200         .Footer_vline02.Left = lngNewFrmWidth1
12210         .Header_hline01.Width = lngNewFrmWidth1
12220         .Header_hline02.Width = lngNewFrmWidth1
12230         .Detail_hline01.Width = lngNewFrmWidth1
12240         .Detail_hline02.Width = lngNewFrmWidth1
12250         .Footer_hline01.Width = lngNewFrmWidth1
12260         .Footer_hline02.Width = lngNewFrmWidth1
12270         .frmCountryCode_Sub.Width = (lngSub_Width - lngAddedWidth2)
12280         .frmCountryCode_Sub_box.Width = (.frmCountryCode_Sub.Width + (2& * lngTpp))
12290         .Nav_box01.Width = .frmCountryCode_Sub.Width
12300         If .Width > lngNewFrmWidth1 Then
12310           .Width = lngNewFrmWidth1
12320         End If
12330         If lngMonitorNum = 1& Then lngTmp02 = (lngFrm_Top - (8& * lngTpp))
12340         DoCmd.MoveSize (lngFrm_Left - (8& * lngTpp)), lngTmp02, (lngFrm_Width - lngAddedWidth2), lngFrm_Height  '(lngFrm_Top - (8& * lngTpp))
12350         If lngMonitorNum > 1& Then
12360           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
12370         End If
12380       End Select
12390     End Select

12400   End With

EXITP:
12410   Exit Sub

ERRH:
12420   Select Case ERR.Number
        Case Else
12430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12440   End Select
12450   Resume EXITP

End Sub
