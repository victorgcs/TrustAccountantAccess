VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_ForeignExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_ForeignExchange"

'VGC 10/11/2017: CHANGES!
' ** TimerInverval: 100

' ** Shortcut Alt keys responsive from this form:
' **   Currency Account: A {cmdCurrencyAccount}
' **   Close:            C {cmdClose}
' **   Load Timer:       D {chkLoadTimer}
' **   Countries:        N {cmdCountries}
' **   Rates:            R {cmdCurrRates}
' **   Symbols:          S {cmdCurrSymbols}
' **   Currencies:       U {cmdCurrencies}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm}

Private lngTpp As Long, blnFormOpening As Boolean, blnJustClose As Boolean
Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           gblnSignal = False

160           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

170           gstrCrtRpt_Version = vbNullString  ' ** Borrowing this variable from Court Reports.

180           If gblnGoToReport = True Then
190             GTRLoad  ' ** Procedure: Below.
200           End If

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
210           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

220           blnJustClose = False: blnFormOpening = False
230           blnGTR_Emblem = False: lngGTR_Stat = 0&

240           Select Case .chkLoadTimer_vis
              Case True
250             .chkLoadTimer.Visible = True
260             .chkLoadTimer.Enabled = False
                '.chkLoadTimer_lbl1.Visible = True
270           Case False
280             .chkLoadTimer.Visible = False
                '.chkLoadTimer_lbl1.Visible = False
290           End Select
300           chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
310           DoEvents

320         Else
330           Cancel = -1
340         End If
350       Else
360         Cancel = -1
370       End If
380     End With

390     If Cancel = -1 Then
400       Beep
410       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
420     End If

EXITP:
430     Exit Sub

ERRH:
440     Select Case ERR.Number
        Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer, Cancel As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Load Timer:       D {chkLoadTimer}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Timer Chk Vis:    T {Form_KeyDown}
        ' **   Close Form:       X {cmdCloseForm}

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
          Case vbKeyD
630         With Me
640           intRetVal = 0
650           If .chkLoadTimer.Visible = True Then
660             If .chkLoadTimer.Enabled = True Then
670               .chkLoadTimer.SetFocus
680             End If
690             .chkLoadTimer = (Not .chkLoadTimer)
700             chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
710           End If
720         End With
730       Case vbKeyX
740         intRetVal = 0
750         cmdClose_Click  ' ** Procedure: Below.
760       End Select
770     End If

        ' ** Ctrl-Shift keys.
780     If intCtrlDown And (Not intAltDown) And intShiftDown Then
790       Select Case intRetVal
          Case vbKeyF
800         With Me
810           intRetVal = 0
820           .FocusHolder.SetFocus
830         End With
840       Case vbKeyT
850         With Me
860           intRetVal = 0
870           If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
880             FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
890           ElseIf CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
900             Select Case .chkLoadTimer.Visible
                Case True
910               .FocusHolder.SetFocus
920               DoEvents
930               .chkLoadTimer.Visible = False
                  '.chkLoadTimer_lbl1.Visible = False
940               gblnTimer = False
950             Case False
960               .chkLoadTimer.Visible = True
970               .chkLoadTimer.Enabled = False
                  '.chkLoadTimer_lbl1.Visible = True
980               chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
990             End Select
1000          End If
1010          .chkLoadTimer_vis = .chkLoadTimer.Visible
1020        End With
1030      Case vbKeyX
1040        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
1050          intRetVal = 0
1060          blnJustClose = True
1070          SetOption_Dev  ' ** Module Function: modUtilities.
1080          gblnSignal = True  ' ** Borrowing this variable.
1090          DoCmd.Close acForm, THIS_NAME
1100          DoCmd.Hourglass False
1110        End If
1120      End Select
1130    End If

EXITP:
1140    KeyCode = intRetVal
1150    Exit Sub

ERRH:
1160    intRetVal = 0
1170    Select Case ERR.Number
        Case Else
1180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1190    End Select
1200    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim strFormName As String

        Static lngTmp01 As Long, lngTmp02 As Long

1310    With Me
1320      .TimerInterval = 0&
1330      If blnGTR_Emblem = True Then
1340        lngGTR_Stat = lngGTR_Stat + 1&
1350        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1360        lngTmp01 = lngTmp01 + 1&
1370        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1380        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1390        If lngTmp01 > 1& Then
1400          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1410        End If
1420        If lngTmp01 = 24& Then
1430          lngTmp01 = 0&
1440          lngTmp02 = lngTmp02 + 1&
1450          If lngTmp02 < 5& Then
1460            .TimerInterval = 50&
1470          Else
1480            blnGTR_Emblem = False
1490          End If
1500        Else
1510          .TimerInterval = 50&
1520        End If
1530      End If  ' ** blnGTR_Emblem.
1540      If gblnSetFocus = True Then
1550        gblnSetFocus = False
1560        strFormName = vbNullString
1570  On Error Resume Next
1580        Do Until strFormName = THIS_NAME
1590          DoCmd.SelectObject acForm, THIS_NAME, False
1600          strFormName = Screen.ActiveForm.Name
1610        Loop
1620  On Error GoTo ERRH
1630        If IsMaximized(Me) = True Then  ' ** Module Function: modWindowFunctions.
1640          DoCmd.Restore
1650        End If
1660        DoEvents
1670      End If
1680      If blnGoingToReport = False Then
1690        Select Case gstrReturningForm
            Case "frmCurrency_Rate"
1700          .cmdCurrRates.SetFocus
1710        Case "frmCurrency"
1720          .cmdCurrencies.SetFocus
1730        Case "frmCurrency_Symbol"
1740          .cmdCurrSymbols.SetFocus
1750        Case "frmCountryCode"
1760          .cmdCountries.SetFocus
1770        Case "frmCurrency_Account"
1780          .cmdCurrencyAccount.SetFocus
1790        Case Else
1800          .cmdCurrRates.SetFocus
1810        End Select
1820        gstrReturningForm = vbNullString
1830        DoCmd.Hourglass False
1840      End If
1850      If gblnGoToReport = True Then
1860        DoCmd.Hourglass True  ' ** Make sure it's still running.
1870        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
1880        Else
1890          If blnGTR_Emblem = False Then
1900            blnGTR_Emblem = True
1910            lngTmp01 = 0&: lngTmp02 = 0&
1920          End If
1930          DoEvents
1940          Select Case blnGoingToReport
              Case True
1950            Select Case garr_varGoToReport(GTR_INV3)
                Case True
1960              Select Case garr_varGoToReport(GTR_CTL3)
                  Case "cmdCurrRates"
1970                .TimerInterval = 50&
1980                cmdCurrRates_Click  ' ** Procedure: Below.
1990              Case "cmdCurrencies"
2000                .TimerInterval = 50&
2010                cmdCurrencies_Click  ' ** Procedure: Below.
2020              Case "cmdCurrSymbols"
2030                .TimerInterval = 50&
2040                cmdCurrSymbols_Click  ' ** Procedure: Below.
2050              Case "cmdCountries"
2060                .TimerInterval = 50&
2070                cmdCountries_Click  ' ** Procedure: Below.
2080              Case Else
                    ' ** Shouldn't happen.
2090                blnGoingToReport = False
2100                DoCmd.Hourglass False
2110                gblnGoToReport = False
2120                blnGTR_Emblem = False
2130                GTREmblem_Off  ' ** Procedure: Below.
2140                Beep
2150              End Select
2160            Case False
                  ' ** Just sit there, and we're done.
2170              blnGoingToReport = False
2180              DoCmd.Hourglass False
2190              gblnGoToReport = False
2200              blnGTR_Emblem = False
2210              GTREmblem_Off  ' ** Procedure: Below.
2220              Beep
2230            End Select
2240          Case False
2250            If garr_varGoToReport(GTR_FRM3) = THIS_NAME Then
                  ' ** This is second menu in path.
2260              blnGoingToReport = True
                  ' ** No more menus.
2270              .Controls(garr_varGoToReport(GTR_CTL3)).SetFocus
2280              lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL3)).Top
2290              lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL3)).Left
2300              .GoToReport_arw_menu_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
2310              .GoToReport_arw_menu_img.Visible = True
2320              DoEvents
2330              .TimerInterval = 50&
2340            Else
                  ' ** Shouldn't be here, then!
2350              blnGoingToReport = False
2360              DoCmd.Hourglass False
2370              gblnGoToReport = False
2380              blnGTR_Emblem = False
2390              GTREmblem_Off  ' ** Procedure: Below.
2400              Beep
2410            End If
2420          End Select
2430        End If
2440      End If
2450      If gblnTimer = True Then
2460        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
2470          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
2480          gblnTimer = False
2490          DoCmd.Hourglass False
2500        End If
2510      End If
2520    End With

EXITP:
2530    Exit Sub

ERRH:
2540    DoCmd.Hourglass False
2550    Select Case ERR.Number
        Case Else
2560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2570    End Select
2580    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2610    With Me

2620      If gblnSignal = False Then  ' ** Borrowing this variable.
2630        DoCmd.Hourglass True
2640        DoEvents
2650      End If

          ' ** Make sure all reports are closed.
2660      Do While Reports.Count > 0
2670        DoCmd.Close acReport, Reports(0).Name
2680      Loop

2690      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2700      gblnAdmin = False
2710      gstrFormQuerySpec = vbNullString

2720      If blnJustClose = False Then

2730        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
2740        DoEvents

2750        gblnSetFocus = True
2760        gstrReturningForm = THIS_NAME
2770        Select Case strCallingForm
            Case "frmMenu_Maintenance"
2780          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Utility"
2790        End Select

2800      Else
2810        DoCmd.Hourglass False
2820      End If

2830    End With

EXITP:
2840    Exit Sub

ERRH:
2850    DoCmd.Hourglass False
2860    Select Case ERR.Number
        Case Else
2870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2880    End Select
2890    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2910    With Me
2920      If .cmdClose.Caption <> "&Close" Then
2930        .cmdClose.Caption = "&Close"
2940      End If
2950    End With

EXITP:
2960    Exit Sub

ERRH:
2970    Select Case ERR.Number
        Case Else
2980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2990    End Select
3000    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

3110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub Header_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_box_MouseMove"

3210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub Header_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_MouseMove"

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

Private Sub Header_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl2_MouseMove"

3410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub cmdClose_Click()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3510    With Me
3520      If blnFormOpening = False Then

3530        DoCmd.Hourglass True
3540        DoEvents

3550        .TimerInterval = 0&

3560        blnJustClose = False
3570        DoCmd.Close acForm, THIS_NAME

3580      End If
3590    End With

EXITP:
3600    Exit Sub

ERRH:
3610    blnFormOpening = False
3620    DoCmd.Hourglass False
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
3790          .cmdCurrRates.SetFocus
3800        End With
3810      End Select
3820    End If

        ' ** Shift keys.
3830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3840      Select Case intRetVal
          Case vbKeyTab
3850        With Me
3860          intRetVal = 0
3870          If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
3880            .chkLoadTimer.SetFocus
3890          Else
3900            .cmdCurrencyAccount.SetFocus
3910          End If
3920        End With
3930      End Select
3940    End If

EXITP:
3950    KeyCode = intRetVal
3960    Exit Sub

ERRH:
3970    intRetVal = 0
3980    Select Case ERR.Number
        Case Else
3990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4000    End Select
4010    Resume EXITP

End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_MouseMove"

4110    With Me
4120      .cmdClose.Caption = "< Ba&ck"
4130    End With

EXITP:
4140    Exit Sub

ERRH:
4150    Select Case ERR.Number
        Case Else
4160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4170    End Select
4180    Resume EXITP

End Sub

Private Sub cmdCurrRates_Click()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrRates_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

4210    With Me
4220      If blnFormOpening = False Then

4230        DoCmd.Hourglass True
4240        DoEvents

4250        .TimerInterval = 0&

4260        Set dbs = CurrentDb
4270        With dbs
              ' ** Empty tblCurrency_Rate_Staging.
4280          Set qdf = .QueryDefs("qryCurrency_Rate_Sub_03")
4290          qdf.Execute
4300          Set qdf = Nothing
              ' ** Append qryCurrency_Rate_Sub_02 (qryCurrency_Rate_Sub_01 (tblCurrency, linked to tblCurrency_Symbol),
              ' ** linked to qryCurrency_Rate_Sub_01b (tblCountry_Currency, grouped by curr_id, with country_cnt),
              ' ** qryCurrency_Rate_Sub_01c (qryCurrency_Rate_Sub_01 (tblCurrency, linked to tblCurrency_Symbol), linked to
              ' ** tblCountry_Currency, tblCountry_Currency_Primary, tblCurrency, just curr_unit = True, with curr_id_unit_par,
              ' ** curr_name_unit_par), qryCurrency_Rate_Sub_01d (qryCurrency_Rate_Sub_01 (tblCurrency, linked to
              ' ** tblCurrency_Symbol), linked to tblCurrency_Country_Primary, tblCountry, with country_id, country_name,
              ' ** country_cod3), with add'l fields) to tblCurrency_Rate_Staging, without curr_id = 4, 190, 40, 198, 201.
4310          Set qdf = .QueryDefs("qryCurrency_Rate_Sub_04")
4320          qdf.Execute
4330          Set qdf = Nothing
              ' ** Append qryCurrency_Rate_Sub_04_01_01 (qryCurrency_Rate_Sub_02 (qryCurrency_Rate_Sub_01 (tblCurrency, linked
              ' ** to tblCurrency_Symbol), linked to qryCurrency_Rate_Sub_01b (tblCountry_Currency, grouped by curr_id, with
              ' ** country_cnt), qryCurrency_Rate_Sub_01c (qryCurrency_Rate_Sub_01 (tblCurrency, linked to tblCurrency_Symbol),
              ' ** linked to tblCountry_Currency, tblCountry_Currency_Primary, tblCurrency, just curr_unit = True, with
              ' ** curr_id_unit_par, curr_name_unit_par), qryCurrency_Rate_Sub_01d (qryCurrency_Rate_Sub_01 (tblCurrency,
              ' ** linked to tblCurrency_Symbol), linked to tblCurrency_Country_Primary, tblCountry, with country_id,
              ' ** country_name, country_cod3), with add'l fields), linked to qryCurrency_Rate_Sub_01e (tblForm_Graphics,
              ' ** just 'frmCurrency_Rate_Sub'), just curr_id = 4, 'Armenian Dram') to tblCurrency_Rate_Staging.
4340          Set qdf = .QueryDefs("qryCurrency_Rate_Sub_04_01_02")
4350          qdf.Execute
4360          Set qdf = Nothing
              ' ** Append qryCurrency_Rate_Sub_04_02_01 (qryCurrency_Rate_Sub_02 (qryCurrency_Rate_Sub_01 (tblCurrency, linked
              ' ** to tblCurrency_Symbol), linked to qryCurrency_Rate_Sub_01b (tblCountry_Currency, grouped by curr_id, with
              ' ** country_cnt), qryCurrency_Rate_Sub_01c (qryCurrency_Rate_Sub_01 (tblCurrency, linked to tblCurrency_Symbol),
              ' ** linked to tblCountry_Currency, tblCountry_Currency_Primary, tblCurrency, just curr_unit = True, with
              ' ** curr_id_unit_par, curr_name_unit_par), qryCurrency_Rate_Sub_01d (qryCurrency_Rate_Sub_01 (tblCurrency,
              ' ** linked to tblCurrency_Symbol), linked to tblCurrency_Country_Primary, tblCountry, with country_id,
              ' ** country_name, country_cod3), with add'l fields), linked to qryCurrency_Rate_Sub_01e (tblForm_Graphics,
              ' ** just 'frmCurrency_Rate_Sub'), just curr_id = 190, 'Bitcoin') to tblCurrency_Rate_Staging.
4370          Set qdf = .QueryDefs("qryCurrency_Rate_Sub_04_02_02")
4380          qdf.Execute
4390          Set qdf = Nothing
              ' ** Append qryCurrency_Rate_Sub_04_03_01 (qryCurrency_Rate_Sub_02 (qryCurrency_Rate_Sub_01 (tblCurrency, linked
              ' ** to tblCurrency_Symbol), linked to qryCurrency_Rate_Sub_01b (tblCountry_Currency, grouped by curr_id, with
              ' ** country_cnt), qryCurrency_Rate_Sub_01c (qryCurrency_Rate_Sub_01 (tblCurrency, linked to tblCurrency_Symbol),
              ' ** linked to tblCountry_Currency, tblCountry_Currency_Primary, tblCurrency, just curr_unit = True, with
              ' ** curr_id_unit_par, curr_name_unit_par), qryCurrency_Rate_Sub_01d (qryCurrency_Rate_Sub_01 (tblCurrency,
              ' ** linked to tblCurrency_Symbol), linked to tblCurrency_Country_Primary, tblCountry, with country_id,
              ' ** country_name, country_cod3), with add'l fields), linked to qryCurrency_Rate_Sub_01e (tblForm_Graphics,
              ' ** just 'frmCurrency_Rate_Sub'), just curr_id = 40, 'Cape Verdean Escudo') to tblCurrency_Rate_Staging.
4400          Set qdf = .QueryDefs("qryCurrency_Rate_Sub_04_03_02")
4410          qdf.Execute
4420          Set qdf = Nothing
              ' ** Append qryCurrency_Rate_Sub_04_04_01 (qryCurrency_Rate_Sub_02 (qryCurrency_Rate_Sub_01 (tblCurrency, linked
              ' ** to tblCurrency_Symbol), linked to qryCurrency_Rate_Sub_01b (tblCountry_Currency, grouped by curr_id, with
              ' ** country_cnt), qryCurrency_Rate_Sub_01c (qryCurrency_Rate_Sub_01 (tblCurrency, linked to tblCurrency_Symbol),
              ' ** linked to tblCountry_Currency, tblCountry_Currency_Primary, tblCurrency, just curr_unit = True, with
              ' ** curr_id_unit_par, curr_name_unit_par), qryCurrency_Rate_Sub_01d (qryCurrency_Rate_Sub_01 (tblCurrency,
              ' ** linked to tblCurrency_Symbol), linked to tblCurrency_Country_Primary, tblCountry, with country_id,
              ' ** country_name, country_cod3), with add'l fields), linked to qryCurrency_Rate_Sub_01e (tblForm_Graphics,
              ' ** just 'frmCurrency_Rate_Sub'), just curr_id = 198, 'Litecoin') to tblCurrency_Rate_Staging.
4430          Set qdf = .QueryDefs("qryCurrency_Rate_Sub_04_04_02")
4440          qdf.Execute
4450          Set qdf = Nothing
              ' ** Append qryCurrency_Rate_Sub_04_05_01 (qryCurrency_Rate_Sub_02 (qryCurrency_Rate_Sub_01 (tblCurrency, linked
              ' ** to tblCurrency_Symbol), linked to qryCurrency_Rate_Sub_01b (tblCountry_Currency, grouped by curr_id, with
              ' ** country_cnt), qryCurrency_Rate_Sub_01c (qryCurrency_Rate_Sub_01 (tblCurrency, linked to tblCurrency_Symbol),
              ' ** linked to tblCountry_Currency, tblCountry_Currency_Primary, tblCurrency, just curr_unit = True, with
              ' ** curr_id_unit_par, curr_name_unit_par), qryCurrency_Rate_Sub_01d (qryCurrency_Rate_Sub_01 (tblCurrency,
              ' ** linked to tblCurrency_Symbol), linked to tblCurrency_Country_Primary, tblCountry, with country_id,
              ' ** country_name, country_cod3), with add'l fields), linked to qryCurrency_Rate_Sub_01e (tblForm_Graphics,
              ' ** just 'frmCurrency_Rate_Sub'), just curr_id = 201, 'Pridnestrovie Ruble') to tblCurrency_Rate_Staging.
4460          Set qdf = .QueryDefs("qryCurrency_Rate_Sub_04_05_02")
4470          qdf.Execute
4480          Set qdf = Nothing
              ' ** Append qryCurrency_Rate_Sub_04_06_01 (qryCurrency_Rate_Sub_02 (qryCurrency_Rate_Sub_01 (tblCurrency, linked
              ' ** to tblCurrency_Symbol), linked to qryCurrency_Rate_Sub_01b (tblCountry_Currency, grouped by curr_id, with
              ' ** country_cnt), qryCurrency_Rate_Sub_01c (qryCurrency_Rate_Sub_01 (tblCurrency, linked to tblCurrency_Symbol),
              ' ** linked to tblCountry_Currency, tblCountry_Currency_Primary, tblCurrency, just curr_unit = True, with
              ' ** curr_id_unit_par, curr_name_unit_par), qryCurrency_Rate_Sub_01d (qryCurrency_Rate_Sub_01 (tblCurrency,
              ' ** linked to tblCurrency_Symbol), linked to tblCurrency_Country_Primary, tblCountry, with country_id,
              ' ** country_name, country_cod3), with add'l fields), linked to qryCurrency_Rate_Sub_01e (tblForm_Graphics,
              ' ** just 'frmCurrency_Rate_Sub'), just curr_id = 202, 'Portuguese Escudo') to tblCurrency_Rate_Staging.
4490          Set qdf = .QueryDefs("qryCurrency_Rate_Sub_04_06_02")
4500          qdf.Execute
4510          Set qdf = Nothing
4520          .Close
4530        End With
4540        Set dbs = Nothing

4550        gblnSetFocus = True
4560        blnJustClose = True
4570        strDocName = "frmCurrency_Rate"
4580        gstrReturningForm = strDocName
4590        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
4600        If gblnGoToReport = True Then
4610          Forms(strDocName).TimerInterval = 50&
4620          blnGoingToReport = False
4630          blnGTR_Emblem = False
4640        End If
4650        DoCmd.Close acForm, THIS_NAME

4660      End If
4670    End With

EXITP:
4680    Set qdf = Nothing
4690    Set dbs = Nothing
4700    Exit Sub

ERRH:
4710    blnFormOpening = False
4720    DoCmd.Hourglass False
4730    Select Case ERR.Number
        Case Else
4740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4750    End Select
4760    Resume EXITP

End Sub

Private Sub cmdCurrRates_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrRates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4820    intShiftDown = (Shift And acShiftMask) > 0
4830    intAltDown = (Shift And acAltMask) > 0
4840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4860      Select Case intRetVal
          Case vbKeyTab
4870        With Me
4880          intRetVal = 0
4890          .cmdCurrencies.SetFocus
4900        End With
4910      End Select
4920    End If

        ' ** Shift keys.
4930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4940      Select Case intRetVal
          Case vbKeyTab
4950        With Me
4960          intRetVal = 0
4970          .cmdClose.SetFocus
4980        End With
4990      End Select
5000    End If

EXITP:
5010    KeyCode = intRetVal
5020    Exit Sub

ERRH:
5030    intRetVal = 0
5040    Select Case ERR.Number
        Case Else
5050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5060    End Select
5070    Resume EXITP

End Sub

Private Sub cmdCurrencies_Click()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrencies_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

5110    With Me
5120      If blnFormOpening = False Then

5130        DoCmd.Hourglass True
5140        DoEvents

5150        .TimerInterval = 0&

5160        Set dbs = CurrentDb
5170        With dbs
              ' ** Empty tblCurrency_Staging.
5180          Set qdf = .QueryDefs("qryCurrency_Sub_05")
5190          qdf.Execute
5200          Set qdf = Nothing
              ' ** Append qryCurrency_Sub_04 (qryCurrency_Sub_04_02 (qryCurrency_Sub_04_01 (qryCurrency_Sub_01
              ' ** (tblCurrency, all records, all fields), linked to qryCurrency_Sub_02 (qryCurrency_Sub_01
              ' ** (tblCurrency, all records, all fields), linked to tblCurrency_Country_Primary),
              ' ** tblCurrency_Symbol), with add'l fields), linked to qryCurrency_Sub_04_03 (tblCountry_Currency,
              ' ** grouped by curr_id, with country_cnt), tblCountry) to tblCurrency_Staging.
5210          Set qdf = .QueryDefs("qryCurrency_Sub_06")
5220          qdf.Execute
5230          Set qdf = Nothing
              ' ** Append qryCurrency_Sub_04_06 (qryCurrency_Sub_04_05 (qryCurrency_Sub_04_02 (qryCurrency_Sub_04_01
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to qryCurrency_Sub_02
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to tblCurrency_Country_Primary),
              ' ** tblCurrency_Symbol), with add'l fields), linked to qryCurrency_Sub_04_03 (tblCountry_Currency,
              ' ** grouped by curr_id, with country_cnt), tblCountry), linked to qryCurrency_Sub_04_04 (tblForm_Graphics,
              ' ** just 'frmCurrency_Sub'), just Armenian Dram) to tblCurrency_Staging, just Armenian Dram.
5240          Set qdf = .QueryDefs("qryCurrency_Sub_06_01")
5250          qdf.Execute
5260          Set qdf = Nothing
              ' ** Append qryCurrency_Sub_04_07 (qryCurrency_Sub_04_05 (qryCurrency_Sub_04_02 (qryCurrency_Sub_04_01
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to qryCurrency_Sub_02
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to tblCurrency_Country_Primary),
              ' ** tblCurrency_Symbol), with add'l fields), linked to qryCurrency_Sub_04_03 (tblCountry_Currency,
              ' ** grouped by curr_id, with country_cnt), tblCountry), linked to qryCurrency_Sub_04_04 (tblForm_Graphics,
              ' ** just 'frmCurrency_Sub'), just Cape Verdean Escudo) to tblCurrency_Staging, just Cape Verdean Escudo.
5270          Set qdf = .QueryDefs("qryCurrency_Sub_06_02")
5280          qdf.Execute
5290          Set qdf = Nothing
              ' ** Append qryCurrency_Sub_04_08 (qryCurrency_Sub_04_05 (qryCurrency_Sub_04_02 (qryCurrency_Sub_04_01
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to qryCurrency_Sub_02
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to tblCurrency_Country_Primary),
              ' ** tblCurrency_Symbol), with add'l fields), linked to qryCurrency_Sub_04_03 (tblCountry_Currency,
              ' ** grouped by curr_id, with country_cnt), tblCountry), linked to qryCurrency_Sub_04_04 (tblForm_Graphics,
              ' ** just 'frmCurrency_Sub'), just Bitcoin) to tblCurrency_Staging, just Bitcoin.
5300          Set qdf = .QueryDefs("qryCurrency_Sub_06_03")
5310          qdf.Execute
5320          Set qdf = Nothing
              ' ** Append qryCurrency_Sub_04_09 (qryCurrency_Sub_04_05 (qryCurrency_Sub_04_02 (qryCurrency_Sub_04_01
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to qryCurrency_Sub_02
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to tblCurrency_Country_Primary),
              ' ** tblCurrency_Symbol), with add'l fields), linked to qryCurrency_Sub_04_03 (tblCountry_Currency,
              ' ** grouped by curr_id, with country_cnt), tblCountry), linked to qryCurrency_Sub_04_04 (tblForm_Graphics,
              ' ** just 'frmCurrency_Sub'), just Pridnestrovie Ruble) to tblCurrency_Staging, just Pridnestrovie Ruble.
5330          Set qdf = .QueryDefs("qryCurrency_Sub_06_04")
5340          qdf.Execute
5350          Set qdf = Nothing
              ' ** Append qryCurrency_Sub_04_10 (qryCurrency_Sub_04_05 (qryCurrency_Sub_04_02 (qryCurrency_Sub_04_01
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to qryCurrency_Sub_02
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to tblCurrency_Country_Primary),
              ' ** tblCurrency_Symbol), with add'l fields), linked to qryCurrency_Sub_04_03 (tblCountry_Currency,
              ' ** grouped by curr_id, with country_cnt), tblCountry), linked to qryCurrency_Sub_04_04 (tblForm_Graphics,
              ' ** just 'frmCurrency_Sub'), just Litecoin) to tblCurrency_Staging, just Litecoin.
5360          Set qdf = .QueryDefs("qryCurrency_Sub_06_05")
5370          qdf.Execute
5380          Set qdf = Nothing
              ' ** Append qryCurrency_Sub_04_11 (qryCurrency_Sub_04_05 (qryCurrency_Sub_04_02 (qryCurrency_Sub_04_01
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to qryCurrency_Sub_02
              ' ** (qryCurrency_Sub_01 (tblCurrency, all records, all fields), linked to tblCurrency_Country_Primary),
              ' ** tblCurrency_Symbol), with add'l fields), linked to qryCurrency_Sub_04_03 (tblCountry_Currency,
              ' ** grouped by curr_id, with country_cnt), tblCountry), linked to qryCurrency_Sub_04_04 (tblForm_Graphics,
              ' ** just 'frmCurrency_Sub'), just Portuguese Escudo) to tblCurrency_Staging, just Portuguese Escudo.
5390          Set qdf = .QueryDefs("qryCurrency_Sub_06_06")
5400          qdf.Execute
5410          Set qdf = Nothing
5420          .Close
5430        End With
5440        Set dbs = Nothing

5450        gblnSetFocus = True
5460        blnJustClose = True
5470        strDocName = "frmCurrency"
5480        gstrReturningForm = strDocName
5490        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
5500        If gblnGoToReport = True Then
5510          Forms(strDocName).TimerInterval = 50&
5520          blnGoingToReport = False
5530          blnGTR_Emblem = False
5540        End If
5550        DoCmd.Close acForm, THIS_NAME

5560      End If
5570    End With

EXITP:
5580    Set qdf = Nothing
5590    Set dbs = Nothing
5600    Exit Sub

ERRH:
5610    blnFormOpening = False
5620    DoCmd.Hourglass False
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Private Sub cmdCurrencies_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrencies_KeyDown"

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
          Case vbKeyTab
5770        With Me
5780          intRetVal = 0
5790          .cmdCurrSymbols.SetFocus
5800        End With
5810      End Select
5820    End If

        ' ** Shift keys.
5830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5840      Select Case intRetVal
          Case vbKeyTab
5850        With Me
5860          intRetVal = 0
5870          .cmdCurrRates.SetFocus
5880        End With
5890      End Select
5900    End If

EXITP:
5910    KeyCode = intRetVal
5920    Exit Sub

ERRH:
5930    intRetVal = 0
5940    Select Case ERR.Number
        Case Else
5950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5960    End Select
5970    Resume EXITP

End Sub

Private Sub cmdCurrSymbols_Click()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrSymbols_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String

6010    With Me
6020      If blnFormOpening = False Then

6030        DoCmd.Hourglass True
6040        DoEvents

6050        .TimerInterval = 0&

6060        blnFormOpening = True

6070        Set dbs = CurrentDb
6080        With dbs
6090          Set rst = .OpenRecordset("tblCurrency_Symbol_Staging", dbOpenDynaset, dbReadOnly)
6100          If rst.BOF = True And rst.EOF = True Then
6110            rst.Close
6120            Set rst = Nothing
                ' ** Append qryCurrency_Symbol_Sub_01 (tblCurrency_Symbol, linked to
                ' ** tblCurrency_Country_Primary) to tblCurrency_Symbol_Staging, without Metals.
6130            Set qdf = .QueryDefs("qryCurrency_Symbol_Sub_03")
6140            qdf.Execute
6150            Set qdf = Nothing
                ' ** Append qryCurrency_Symbol_Sub_01 (tblCurrency_Symbol, linked to tblCurrency_Country_Primary),
                ' ** with qryCurrency_Symbol_Sub_05_01_01_Gold (qryCurrency_Symbol_Sub_04 (tblForm_Graphics, just
                ' ** 'frmCurrency_Symbol_Sub'), just 'Gold'), to tblCurrency_Symbol_Staging; Cartesian.
6160            Set qdf = .QueryDefs("qryCurrency_Symbol_Sub_05_01_02_Gold")
6170            qdf.Execute
6180            Set qdf = Nothing
                ' ** Append qryCurrency_Symbol_Sub_01 (tblCurrency_Symbol, linked to tblCurrency_Country_Primary),
                ' ** with qryCurrency_Symbol_Sub_05_02_01_Silver (qryCurrency_Symbol_Sub_04 (tblForm_Graphics, just
                ' ** 'frmCurrency_Symbol_Sub'), just 'Silver'), to tblCurrency_Symbol_Staging; Cartesian.
6190            Set qdf = .QueryDefs("qryCurrency_Symbol_Sub_05_02_02_Silver")
6200            qdf.Execute
6210            Set qdf = Nothing
                ' ** Append qryCurrency_Symbol_Sub_01 (tblCurrency_Symbol, linked to tblCurrency_Country_Primary),
                ' ** with qryCurrency_Symbol_Sub_05_03_01_Platinum (qryCurrency_Symbol_Sub_04 (tblForm_Graphics, just
                ' ** 'frmCurrency_Symbol_Sub'), just 'Platinum'), to tblCurrency_Symbol_Staging; Cartesian.
6220            Set qdf = .QueryDefs("qryCurrency_Symbol_Sub_05_03_02_Platinum")
6230            qdf.Execute
6240            Set qdf = Nothing
                ' ** Append qryCurrency_Symbol_Sub_01 (tblCurrency_Symbol, linked to tblCurrency_Country_Primary),
                ' ** with qryCurrency_Symbol_Sub_05_04_01_Palladium (qryCurrency_Symbol_Sub_04 (tblForm_Graphics, just
                ' ** 'frmCurrency_Symbol_Sub'), just 'Palladium'), to tblCurrency_Symbol_Staging; Cartesian.
6250            Set qdf = .QueryDefs("qryCurrency_Symbol_Sub_05_04_02_Palladium")
6260            qdf.Execute
6270            Set qdf = Nothing
                ' ** Append qryCurrency_Symbol_Sub_01 (tblCurrency_Symbol, linked to tblCurrency_Country_Primary),
                ' ** with qryCurrency_Symbol_Sub_05_05_01_Armenia (qryCurrency_Symbol_Sub_04 (tblForm_Graphics, just
                ' ** 'frmCurrency_Symbol_Sub'), just 'Armenian Dram'), to tblCurrency_Symbol_Staging; Cartesian.
6280            Set qdf = .QueryDefs("qryCurrency_Symbol_Sub_05_05_02_Armenia")
6290            qdf.Execute
6300            Set qdf = Nothing
                ' ** Append qryCurrency_Symbol_Sub_01 (tblCurrency_Symbol, linked to tblCurrency_Country_Primary),
                ' ** with qryCurrency_Symbol_Sub_05_06_01_Cape_Verde (qryCurrency_Symbol_Sub_04 (tblForm_Graphics, just
                ' ** 'frmCurrency_Symbol_Sub'), just 'Cape Verdean Escudo'), to tblCurrency_Symbol_Staging; Cartesian.
6310            Set qdf = .QueryDefs("qryCurrency_Symbol_Sub_05_06_02_Cape_Verde")
6320            qdf.Execute
6330            Set qdf = Nothing
                ' ** Append qryCurrency_Symbol_Sub_01 (tblCurrency_Symbol, linked to tblCurrency_Country_Primary),
                ' ** with qryCurrency_Symbol_Sub_05_07_01_Bitcoin (qryCurrency_Symbol_Sub_04 (tblForm_Graphics, just
                ' ** 'frmCurrency_Symbol_Sub'), just 'Bitcoin'), to tblCurrency_Symbol_Staging; Cartesian.
6340            Set qdf = .QueryDefs("qryCurrency_Symbol_Sub_05_07_02_Bitcoin")
6350            qdf.Execute
6360            Set qdf = Nothing
                ' ** Append qryCurrency_Symbol_Sub__01 (tblCurrency_Symbol, linked to tblCurrency_Country_Primary),
                ' ** with qryCurrency_Symbol_Sub_05_08_01_Pridnestrovie (qryCurrency_Symbol_Sub_04 (tblForm_Graphics, just
                ' ** 'frmCurrency_Symbol_Sub'), just 'Pridnestrovie Ruble'), to tblCurrency_Symbol_Staging; Cartesian.
6370            Set qdf = .QueryDefs("qryCurrency_Symbol_Sub_05_08_02_Pridnestrovie")
6380            qdf.Execute
6390            Set qdf = Nothing
                ' ** Append qryCurrency_Symbol_Sub_01 (tblCurrency_Symbol, linked to tblCurrency_Country_Primary),
                ' ** with qryCurrency_Symbol_Sub_05_09_01_Litecoin (qryCurrency_Symbol_Sub_04 (tblForm_Graphics, just
                ' ** 'frmCurrency_Symbol_Sub'), just 'Litecoin'), to tblCurrency_Symbol_Staging; Cartesian.
6400            Set qdf = .QueryDefs("qryCurrency_Symbol_Sub_05_09_02_Litecoin")
6410            qdf.Execute
6420            Set qdf = Nothing
                ' ** Append qryCurrency_Symbol_Sub_01 (tblCurrency_Symbol, linked to tblCurrency_Country_Primary,
                ' ** qryCurrency_Symbol_Sub_01_01 (tblCurrency, just units, linked to parent currency)), with
                ' ** qryCurrency_Symbol_Sub_05_10_01_Portugal (qryCurrency_Symbol_Sub_04 (tblForm_Graphics, just
                ' ** 'frmCurrency_Symbol_Sub'), just 'Portuguese Escudo'), to tblCurrency_Symbol_Staging; Cartesian.
6430            Set qdf = .QueryDefs("qryCurrency_Symbol_Sub_05_10_02_Portugal")
6440            qdf.Execute
6450            Set qdf = Nothing
6460          Else
6470            rst.Close
6480            Set rst = Nothing
6490          End If
6500          .Close
6510        End With
6520        Set dbs = Nothing

6530        gblnSetFocus = True
6540        blnJustClose = True
6550        strDocName = "frmCurrency_Symbol"
6560        gstrReturningForm = strDocName
6570        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6580        If gblnGoToReport = True Then
6590          Forms(strDocName).TimerInterval = 50&
6600          blnGoingToReport = False
6610          blnGTR_Emblem = False
6620        End If
6630        DoCmd.Close acForm, THIS_NAME

6640      End If
6650    End With

EXITP:
6660    Exit Sub

ERRH:
6670    blnFormOpening = False
6680    DoCmd.Hourglass False
6690    Select Case ERR.Number
        Case Else
6700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6710    End Select
6720    Resume EXITP

End Sub

Private Sub cmdCurrSymbols_KeyDown(KeyCode As Integer, Shift As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrSymbols_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6820    intShiftDown = (Shift And acShiftMask) > 0
6830    intAltDown = (Shift And acAltMask) > 0
6840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6860      Select Case intRetVal
          Case vbKeyTab
6870        With Me
6880          intRetVal = 0
6890          .cmdCountries.SetFocus
6900        End With
6910      End Select
6920    End If

        ' ** Shift keys.
6930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6940      Select Case intRetVal
          Case vbKeyTab
6950        With Me
6960          intRetVal = 0
6970          .cmdCurrencies.SetFocus
6980        End With
6990      End Select
7000    End If

EXITP:
7010    KeyCode = intRetVal
7020    Exit Sub

ERRH:
7030    intRetVal = 0
7040    Select Case ERR.Number
        Case Else
7050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7060    End Select
7070    Resume EXITP

End Sub

Private Sub cmdCountries_Click()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

7110    With Me
7120      If blnFormOpening = False Then

7130        DoCmd.Hourglass True
7140        DoEvents

7150        .TimerInterval = 0&

7160        Set dbs = CurrentDb
7170        With dbs
              ' ** Empty tblCountry_Staging.
7180          Set qdf = .QueryDefs("qryCountryCode_Sub_04")
7190          qdf.Execute
7200          Set qdf = Nothing
              ' ** Append qryCountryCode_Sub_03 (qryCountryCode_Sub_01 (tblCountry, all records,
              ' ** all fields), linked to qryCountryCode_Sub_02 (tblCountry_Currency, grouped by
              ' ** country_id, with currency_cnt)) to tblCountry_Staging.
7210          Set qdf = .QueryDefs("qryCountryCode_Sub_05")
7220          qdf.Execute
7230          Set qdf = Nothing
7240          .Close
7250        End With
7260        Set dbs = Nothing

7270        gblnSetFocus = True
7280        blnJustClose = True
7290        strDocName = "frmCountryCode"
7300        gstrReturningForm = strDocName
7310        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
7320        If gblnGoToReport = True Then
7330          Forms(strDocName).TimerInterval = 50&
7340          blnGoingToReport = False
7350          blnGTR_Emblem = False
7360        End If
7370        DoCmd.Close acForm, THIS_NAME

            ' ** This just doesn't want to turn off within Countries!
7380        DoCmd.Hourglass False

7390      End If
7400    End With

EXITP:
7410    Set qdf = Nothing
7420    Set dbs = Nothing
7430    Exit Sub

ERRH:
7440    blnFormOpening = False
7450    DoCmd.Hourglass False
7460    Select Case ERR.Number
        Case Else
7470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7480    End Select
7490    Resume EXITP

End Sub

Private Sub cmdCountries_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_KeyDown"

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
7590          .cmdCurrencyAccount.SetFocus
7600        End With
7610      End Select
7620    End If

        ' ** Shift keys.
7630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7640      Select Case intRetVal
          Case vbKeyTab
7650        With Me
7660          intRetVal = 0
7670          .cmdCurrSymbols.SetFocus
7680        End With
7690      End Select
7700    End If

EXITP:
7710    KeyCode = intRetVal
7720    Exit Sub

ERRH:
7730    intRetVal = 0
7740    Select Case ERR.Number
        Case Else
7750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7760    End Select
7770    Resume EXITP

End Sub

Private Sub cmdCurrencyAccount_Click()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrencyAccount_Click"

        Dim strDocName As String

7810    With Me
7820      If blnFormOpening = False Then

            'DoCmd.Hourglass True
7830        DoEvents

7840        .TimerInterval = 0&

7850        blnFormOpening = True

7860        gblnSetFocus = True
7870        blnJustClose = True
7880        strDocName = "frmCurrency_Account"
7890        gstrReturningForm = strDocName
7900        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
7910        DoCmd.Close acForm, THIS_NAME

7920        DoCmd.Hourglass False

7930      End If
7940    End With

EXITP:
7950    Exit Sub

ERRH:
7960    blnFormOpening = False
7970    DoCmd.Hourglass False
7980    Select Case ERR.Number
        Case Else
7990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8000    End Select
8010    Resume EXITP

End Sub

Private Sub cmdCurrencyAccount_KeyDown(KeyCode As Integer, Shift As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCurrencyAccount_KeyDown"

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
          Case vbKeyTab
8170        With Me
8180          intRetVal = 0
8190          If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
8200            .chkLoadTimer.SetFocus
8210          Else
8220            .cmdClose.SetFocus
8230          End If
8240        End With
8250      End Select
8260    End If

        ' ** Shift keys.
8270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8280      Select Case intRetVal
          Case vbKeyTab
8290        With Me
8300          intRetVal = 0
8310          .cmdCountries.SetFocus
8320        End With
8330      End Select
8340    End If

EXITP:
8350    KeyCode = intRetVal
8360    Exit Sub

ERRH:
8370    intRetVal = 0
8380    Select Case ERR.Number
        Case Else
8390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8400    End Select
8410    Resume EXITP

End Sub

Private Sub chkLoadTimer_AfterUpdate()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_AfterUpdate"

8510    With Me
8520      If IsLoaded("frmLoadTimer", acForm) = False Then  ' ** Module Function: modFileUtilities.
            ' ** Don't turn it off if it's still playing out a timing sequence.
8530        gblnTimer = False
8540      End If
8550      Select Case .chkLoadTimer
          Case True
8560        .chkLoadTimer_lbl.FontBold = True
8570        If .chkLoadTimer.Visible = True Then  ' ** This may get turned off by the Form_Timer() sequence.
8580          gblnTimer = True
8590        End If
8600      Case False
8610        .chkLoadTimer_lbl.FontBold = False
8620      End Select
8630    End With

EXITP:
8640    Exit Sub

ERRH:
8650    Select Case ERR.Number
        Case Else
8660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8670    End Select
8680    Resume EXITP

End Sub

Private Sub chkLoadTimer_KeyDown(KeyCode As Integer, Shift As Integer)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8720    intShiftDown = (Shift And acShiftMask) > 0
8730    intAltDown = (Shift And acAltMask) > 0
8740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8760      Select Case intRetVal
          Case vbKeyTab
8770        With Me
8780          intRetVal = 0
8790          .cmdClose.SetFocus
8800        End With
8810      End Select
8820    End If

        ' ** Shift keys.
8830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8840      Select Case intRetVal
          Case vbKeyTab
8850        With Me
8860          intRetVal = 0
8870          .cmdCurrencyAccount.SetFocus
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

Private Sub GTRLoad()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "GTRLoad"

        Dim strTmp01 As String
        Dim lngX As Long

9010    With Me
9020      .GoToReport_arw_menu_img.ControlSource = "GoToReport_arw_menu_img"
9030      For lngX = 1& To 24&
9040        strTmp01 = "GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img"
9050        .Controls(strTmp01).ControlSource = strTmp01
9060      Next
9070      DoEvents
9080    End With

EXITP:
9090    Exit Sub

ERRH:
9100    Select Case ERR.Number
        Case Else
9110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9120    End Select
9130    Resume EXITP

End Sub

Public Sub GTREmblem_Off()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

9210    With Me
9220      blnGTR_Emblem = False
9230      For lngX = 1& To 24&
9240  On Error Resume Next
9250        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
9260        If ERR.Number <> 0 Then
9270  On Error GoTo ERRH
9280          Exit For
9290        Else
9300  On Error GoTo ERRH
9310        End If
9320      Next
9330    End With

EXITP:
9340    Exit Sub

ERRH:
9350    Select Case ERR.Number
        Case Else
9360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9370    End Select
9380    Resume EXITP

End Sub
