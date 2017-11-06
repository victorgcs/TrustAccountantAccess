VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_Rate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_Rate"

'VGC 10/02/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** Shortcut Alt keys responsive from this form:
' **   Active:           A {curr_active on frmCurrency_Rate_Sub}
' **   Close:            C {cmdClose}
' **   Code:             D {opgCurrencySource_optCurrCode}
' **   Edit:             E {cmdEdit}
' **   Find:             F {cmbCurrencies}
' **   History:          H {cmdHistory}
' **   ISO:              I {opgCurrencySource_optCurrISO}
' **   Excel:            L {cmdExcel}
' **   Name:             M {opgCurrencySource_optCurrName}
' **   Countries:        N {cmdCountries}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Currency Name:    R {curr_name on frmCurrency_Rate_Sub}
' **   Reset:            S {cmdResetFilter}
' **   Date:             T {curr_date on frmCurrency_Rate_Sub}
' **   Update:           U {cmdUpdate}
' **   Convert 1:        V {curr_convert1a on frmCurrency_Rate_Sub}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}
' **   Country:          Y {opgCurrencySource_optCountry}
' **   Convert 2:        Z {curr_convert2a on frmCurrency_Rate_Sub}
' **   Rate 1:           1 {curr_rate1 on frmCurrency_Rate_Sub}
' **   Rate 2:           2 {curr_rate2 on frmCurrency_Rate_Sub}
'Unused:
'B  G  J  K  Q

' ** Shortcut Ctrl keys responsive from this form:
' **   Exclude Active:   E {opgActive_optExclude}
' **   Page Count:       G {chkPageOf}
' **   Include Active:   I {opgActive_optInclude}
' **   Include Notes:    N {chkIncludeNotes}
' **   Only Active:      O {opgActive_optOnly}
' **   Save:             S {cmdSave on frmCurrency_Rate_Sub}
' **   Word:             W {opgCurrencySource_optCurrWord}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Exclude Funds:    E {opgFund_optExclude}
' **   Include Funds:    I {opgFund_optInclude}
' **   Only Metals:      L {opgMetal_optOnly}
' **   Include Metals:   N {opgMetal_optInclude}
' **   Only Funds:       O {opgFund_optOnly}
' **   Exclude Metals:   X {opgMetal_optExclude}

' ** Shortcut Ctrl-Alt keys responsive from this form:
' **   Exclude Units:    D {opgUnit_optExclude}
' **   Exclude BMUs:     E {opgBMU_optExclude}
' **   Include BMUs:     I {opgBMU_optInclude}
' **   Only Misc:        L {opgAlt_optOnly}
' **   Include Misc:     N {opgAlt_optInclude}
' **   Only BMUs:        O {opgBMU_optOnly}
' **   Include Units:    U {opgUnit_optInclude}
' **   Exclude Misc:     X {opgAlt_optExclude}
' **   Only Units:       Y {opgUnit_optOnly}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const FAND As String = " And "  ' ** Filter And.

' ** cmbCurrencies combo box constants:
Private Const CBX_C_CURRID As Integer = 0

Private blnClearFilter_Focus As Boolean, blnClearFilter_MouseDown As Boolean
Private blnCountries_Focus As Boolean, blnCountries_MouseDown As Boolean

Private lngClose_Left As Long, lngBtn_Offset As Long
Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnFromCombo As Boolean
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

150           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170           CLR_DISABLED_FG = CLR_DKGRY
180           CLR_DISABLED_BG = CLR_LTTEAL

190           gstrFormQuerySpec = THIS_NAME
200           blnClearFilter_Focus = False: blnClearFilter_MouseDown = False
210           blnCountries_Focus = False: blnCountries_MouseDown = False
220           blnGTR_Emblem = False: lngGTR_Stat = 0&
230           blnFromCombo = False

240           .GoToReport_Emblem_lbl.Left = 0&  ' ** Just get this out of the way.

250           lngClose_Left = .cmdClose.Left
260           lngBtn_Offset = (.cmdClose.Left - .cmdUpdate.Left)

270           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
280           DoEvents

290           If gblnGoToReport = True Then
300             Select Case IsNull(garr_varGoToReport(GTR_CTL4))
                Case True
310               If .cmdPrint.Enabled = True Then
320                 .cmdPrint.SetFocus
330               End If
340             Case False
350               .cmdClose.Left = .GTR_Close_line.Left
360               .cmdUpdate.Left = (.cmdClose.Left - lngBtn_Offset)
370               .cmdEdit.Left = (.cmdUpdate.Left - lngBtn_Offset)
380               .cmdHistory.Left = (.cmdEdit.Left - lngBtn_Offset)
390               DoEvents
400               If garr_varGoToReport(GTR_CTL4) = "cmdHistory" Then
410                 If .GoToReport_arw_hist_img.Visible = False Then
420                   .GoToReport_arw_hist_img.Top = (.cmdHistory.Top + (3& * lngTpp))
430                   .GoToReport_arw_hist_img.Left = ((.cmdHistory.Left - .GoToReport_arw_hist_img.Width) + (2& * lngTpp))
440                   .GoToReport_arw_hist_img.Visible = True
450                   If .cmdHistory.Enabled = True Then
460                     .cmdHistory.SetFocus
470                   End If
480                   DoEvents
490                 End If
500               End If
510             End Select
520           End If

530           opgActive_AfterUpdate  ' ** Procedure: Below.
540           opgFund_AfterUpdate  ' ** Procedure: Below.
550           opgMetal_AfterUpdate  ' ** Procedure: Below.
560           opgBMU_AfterUpdate  ' ** Procedure: Below.
570           opgAlt_AfterUpdate  ' ** Procedure: Below.
580           opgUnit_AfterUpdate  ' ** Procedure: Below.
590           opgCurrencySource_AfterUpdate  ' ** Procedure: Below.
600           chkIncludeNotes_AfterUpdate  ' ** Procedure: Below.
610           chkPageOf_AfterUpdate  ' ** Procedure: Below.
620           DoEvents

630           .cmdEdit.Enabled = True
640           .cmdUpdate.Enabled = False

      #If NoExcel Then
650           .cmdExcel.Enabled = False
      #Else
660           .cmdExcel.Enabled = True
      #End If

670           ResetFilter_Set  ' ** Procedure: Below.

680           .frmCurrency_Rate_Sub.SetFocus
690           .frmCurrency_Rate_Sub.Form.FocusHolder.SetFocus

700         Else
710           Cancel = -1
720         End If
730       Else
740         Cancel = -1
750       End If
760     End With

770     DoCmd.Hourglass False

780     If Cancel = -1 Then
790       Beep
800       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
810     End If

EXITP:
820     Exit Sub

ERRH:
830     DoCmd.Hourglass False
840     Select Case ERR.Number
        Case Else
850       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
860     End Select
870     Resume EXITP

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
        ' **   Convert 1:        V {curr_convert1a on frmCurrency_Rate_Sub}
        ' **   Convert 2:        Z {curr_convert2a on frmCurrency_Rate_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Exclude Active:   E {opgActive_optExclude}
        ' **   Page Count:       G {chkPageOf}
        ' **   Include Active:   I {opgActive_optInclude}
        ' **   Include Notes:    N {chkIncludeNotes}
        ' **   Only Active:      O {opgActive_optOnly}
        ' **   Save:             S {cmdSave on frmCurrency_Rate_Sub}
        ' **   Word:             W {opgCurrencySource_optCurrWord}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Exclude Funds:    E {opgFund_optExclude}
        ' **   Include Funds:    I {opgFund_optInclude}
        ' **   Only Metals:      L {opgMetal_optOnly}
        ' **   Include Metals:   N {opgMetal_optInclude}
        ' **   Only Funds:       O {opgFund_optOnly}
        ' **   Exclude Metals:   X {opgMetal_optExclude}

        ' ** Shortcut Ctrl-Alt keys to other forms and functionality:
        ' **   Exclude Units:    D {opgUnit_optExclude}
        ' **   Exclude BMUs:     E {opgBMU_optExclude}
        ' **   Include BMUs:     I {opgBMU_optInclude}
        ' **   Only Misc:        L {opgAlt_optOnly}
        ' **   Include Misc:     N {opgAlt_optInclude}
        ' **   Only BMUs:        O {opgBMU_optOnly}
        ' **   Include Units:    U {opgUnit_optInclude}
        ' **   Exclude Misc:     X {opgAlt_optExclude}
        ' **   Only Units:       Y {opgUnit_optOnly}

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
          Case vbKeyV
1030        With Me
1040          intRetVal = 0
1050          .frmCurrency_Rate_Sub.SetFocus
1060          .frmCurrency_Rate_Sub.Form.curr_convert1a.SetFocus
1070        End With
1080      Case vbKeyX
1090        intRetVal = 0
1100        cmdClose_Click  ' ** Procedure: Below.
1110      Case vbKeyZ
1120        With Me
1130          intRetVal = 0
1140          .frmCurrency_Rate_Sub.SetFocus
1150          .frmCurrency_Rate_Sub.Form.curr_convert2a.SetFocus
1160        End With
1170      End Select
1180    End If

        ' ** Ctrl keys.
1190    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1200      Select Case intRetVal
          Case vbKeyE
1210        With Me
1220          intRetVal = 0
1230          .opgActive.SetFocus
1240          .opgActive = .opgActive_optExclude.OptionValue
1250          opgActive_AfterUpdate  ' ** Procedure: Below.
1260        End With
1270      Case vbKeyG
1280        With Me
1290          intRetVal = 0
1300          .chkPageOf.SetFocus
1310          .chkPageOf = (Not .chkPageOf)
1320          chkPageOf_AfterUpdate  ' ** Procedure: Below.
1330        End With
1340      Case vbKeyI
1350        With Me
1360          intRetVal = 0
1370          .opgActive.SetFocus
1380          .opgActive = .opgActive_optInclude.OptionValue
1390          opgActive_AfterUpdate  ' ** Procedure: Below.
1400        End With
1410      Case vbKeyN
1420        With Me
1430          intRetVal = 0
1440          .chkIncludeNotes.SetFocus
1450          .chkIncludeNotes = (Not .chkIncludeNotes)
1460          chkIncludeNotes_AfterUpdate  ' ** Procedure: Below.
1470        End With
1480      Case vbKeyO
1490        With Me
1500          intRetVal = 0
1510          .opgActive.SetFocus
1520          .opgActive = .opgActive_optOnly.OptionValue
1530          opgActive_AfterUpdate  ' ** Procedure: Below.
1540        End With
1550      Case vbKeyS
1560        With Me
1570          intRetVal = 0
1580          .frmCurrency_Rate_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmCurrency_Sub.
1590        End With
1600      Case vbKeyW
1610        With Me
1620          intRetVal = 0
1630          .opgCurrencySource.SetFocus
1640          .opgCurrencySource = .opgCurrencySource_optCurrWord.OptionValue
1650          opgCurrencySource_AfterUpdate  ' ** Procedure: Below.
1660        End With
1670      End Select
1680    End If

        ' ** Ctrl-Shift keys.
1690    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1700      Select Case intRetVal
          Case vbKeyE
1710        With Me
1720          intRetVal = 0
1730          .opgFund.SetFocus
1740          .opgFund = .opgFund_optExclude.OptionValue
1750          opgFund_AfterUpdate  ' ** Procedure: Below.
1760        End With
1770      Case vbKeyF
1780        With Me
1790          intRetVal = 0
1800          .FocusHolder.SetFocus
1810        End With
1820      Case vbKeyI
1830        With Me
1840          intRetVal = 0
1850          .opgFund.SetFocus
1860          .opgFund = .opgFund_optInclude.OptionValue
1870          opgFund_AfterUpdate  ' ** Procedure: Below.
1880        End With
1890      Case vbKeyL
1900        With Me
1910          intRetVal = 0
1920          .opgMetal.SetFocus
1930          .opgMetal = .opgMetal_optOnly.OptionValue
1940          opgMetal_AfterUpdate  ' ** Procedure: Below.
1950        End With
1960      Case vbKeyN
1970        With Me
1980          intRetVal = 0
1990          .opgMetal.SetFocus
2000          .opgMetal = .opgMetal_optInclude.OptionValue
2010          opgMetal_AfterUpdate  ' ** Procedure: Below.
2020        End With
2030      Case vbKeyO
2040        With Me
2050          intRetVal = 0
2060          .opgFund.SetFocus
2070          .opgFund = .opgFund_optOnly.OptionValue
2080          opgFund_AfterUpdate  ' ** Procedure: Below.
2090        End With
2100      Case vbKeyX
2110        With Me
2120          intRetVal = 0
2130          .opgMetal.SetFocus
2140          .opgMetal = .opgMetal_optExclude.OptionValue
2150          opgMetal_AfterUpdate  ' ** Procedure: Below.
2160        End With
2170      End Select
2180    End If

        ' ** Ctrl-Alt keys.
2190    If intCtrlDown And intAltDown And (Not intShiftDown) Then
2200      Select Case intRetVal
          Case vbKeyD
2210        With Me
2220          intRetVal = 0
2230          .opgUnit.SetFocus
2240          .opgUnit = .opgUnit_optExclude.OptionValue
2250          opgUnit_AfterUpdate  ' ** Procedure: Below.
2260        End With
2270      Case vbKeyE
2280        With Me
2290          intRetVal = 0
2300          .opgBMU.SetFocus
2310          .opgBMU = .opgBMU_optExclude.OptionValue
2320          opgBMU_AfterUpdate  ' ** Procedure: Below.
2330        End With
2340      Case vbKeyI
2350        With Me
2360          intRetVal = 0
2370          .opgBMU.SetFocus
2380          .opgBMU = .opgBMU_optInclude.OptionValue
2390          opgBMU_AfterUpdate  ' ** Procedure: Below.
2400        End With
2410      Case vbKeyL
2420        With Me
2430          intRetVal = 0
2440          .opgAlt.SetFocus
2450          .opgAlt = .opgAlt_optOnly.OptionValue
2460          opgAlt_AfterUpdate  ' ** Procedure: Below.
2470        End With
2480      Case vbKeyN
2490        With Me
2500          intRetVal = 0
2510          .opgAlt.SetFocus
2520          .opgAlt = .opgAlt_optInclude.OptionValue
2530          opgAlt_AfterUpdate  ' ** Procedure: Below.
2540        End With
2550      Case vbKeyO
2560        With Me
2570          intRetVal = 0
2580          .opgBMU.SetFocus
2590          .opgBMU = .opgBMU_optOnly.OptionValue
2600          opgBMU_AfterUpdate  ' ** Procedure: Below.
2610        End With
2620      Case vbKeyU
2630        With Me
2640          intRetVal = 0
2650          .opgUnit.SetFocus
2660          .opgUnit = .opgUnit_optInclude.OptionValue
2670          opgUnit_AfterUpdate  ' ** Procedure: Below.
2680        End With
2690      Case vbKeyX
2700        With Me
2710          intRetVal = 0
2720          .opgAlt.SetFocus
2730          .opgAlt = .opgAlt_optExclude.OptionValue
2740          opgAlt_AfterUpdate  ' ** Procedure: Below.
2750        End With
2760      Case vbKeyY
2770        With Me
2780          intRetVal = 0
2790          .opgUnit.SetFocus
2800          .opgUnit = .opgUnit_optOnly.OptionValue
2810          opgUnit_AfterUpdate  ' ** Procedure: Below.
2820        End With
2830      End Select
2840    End If

EXITP:
2850    KeyCode = intRetVal
2860    Exit Sub

ERRH:
2870    intRetVal = 0
2880    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2890    Case Else
2900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2910    End Select
2920    Resume EXITP

End Sub

Private Sub Form_Timer()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

        Static lngTmp01 As Long, lngTmp02 As Long

3010    With Me
3020      .TimerInterval = 0&
3030      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
3040        lngGTR_Stat = lngGTR_Stat + 1&
3050        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
3060        lngTmp01 = lngTmp01 + 1&
3070        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
3080        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
3090        If lngTmp01 > 1& Then
3100          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
3110        End If
3120        If lngTmp01 = 24& Then
3130          lngTmp01 = 0&
3140          lngTmp02 = lngTmp02 + 1&
3150          If lngTmp02 < 5& Then
3160            .TimerInterval = 50&
3170          Else
3180            blnGTR_Emblem = False
3190          End If
3200        Else
3210          .TimerInterval = 50&
3220        End If
3230      End If  ' ** blnGTR_Emblem.
3240      If gblnSetFocus = True Then
3250        gblnSetFocus = False
3260        strFormName = vbNullString
3270  On Error Resume Next
3280        Do Until strFormName = THIS_NAME
3290          DoCmd.SelectObject acForm, THIS_NAME, False
3300          strFormName = Screen.ActiveForm.Name
3310        Loop
3320  On Error GoTo ERRH
3330        If gstrReturningForm <> vbNullString Then
3340          Select Case gstrReturningForm
              Case "frmCurrency_History"
3350            .cmdHistory.SetFocus
3360          Case Else
3370            .frmCurrency_Rate_Sub.SetFocus
3380          End Select
3390        Else
3400          .frmCurrency_Rate_Sub.SetFocus
3410        End If
3420      End If
3430      If gblnGoToReport = True Then
3440        DoCmd.Hourglass True  ' ** Make sure it's still running.
3450        DoEvents
3460        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
3470        Else
3480          If blnGTR_Emblem = False And IsNull(garr_varGoToReport(GTR_CTL4)) = False Then
3490            blnGTR_Emblem = True
3500            lngTmp01 = 0&: lngTmp02 = 0&
3510          End If
3520          DoEvents
3530          Select Case blnGoingToReport
              Case True
3540            Select Case IsNull(garr_varGoToReport(GTR_CTL4))
                Case True
3550              gblnGoToReport = False
3560              blnGoingToReport = False
3570              DoCmd.Hourglass False
3580              DoEvents
3590            Case False
3600              Select Case garr_varGoToReport(GTR_CTL4)
                  Case "cmdHistory"
3610                .TimerInterval = 50&
3620                cmdHistory_Click  ' ** Procedure: Below.
3630              End Select
3640            End Select
3650          Case False
3660            DoCmd.Hourglass True  ' ** Make sure it's still running.
3670            DoEvents
3680            blnGoingToReport = True
3690            Select Case IsNull(garr_varGoToReport(GTR_CTL4))
                Case True
3700              If .cmdPrint.Enabled = True Then
3710                .cmdPrint.SetFocus
3720              End If
3730              DoEvents
3740              Beep
3750              DoCmd.Hourglass False
3760              DoEvents
3770              .TimerInterval = 50&
3780            Case False
3790              Select Case garr_varGoToReport(GTR_CTL4)
                  Case "cmdHistory"
3800                Select Case .cmdHistory.Enabled
                    Case True
3810                  .cmdHistory.SetFocus
3820                  .GoToReport_arw_hist_img.Top = (.cmdHistory.Top + lngTpp)
3830                  .GoToReport_arw_hist_img.Left = ((.cmdHistory.Left - .GoToReport_arw_hist_img.Width) + (2& * lngTpp))
3840                  .GoToReport_arw_hist_img.Visible = True
3850                  DoEvents
3860                  .TimerInterval = GTR_WAIT
3870                Case False
                      ' ** Should be enabled if they're here!
3880                  blnGoingToReport = False
3890                  DoCmd.Hourglass False
3900                  gblnGoToReport = False
3910                  blnGTR_Emblem = False
3920                  GTREmblem_Off  ' ** Procedure: Below.
3930                  Beep
3940                End Select
3950              Case Else
                    ' ** Nothing else
3960                blnGoingToReport = False
3970                DoCmd.Hourglass False
3980                gblnGoToReport = False
3990                blnGTR_Emblem = False
4000                GTREmblem_Off  ' ** Procedure: Below.
4010                Beep
4020              End Select
4030            End Select
4040          End Select
4050        End If
4060      End If
4070      DoCmd.Hourglass False
4080    End With

EXITP:
4090    Exit Sub

ERRH:
4100    DoCmd.Hourglass False
4110    Select Case ERR.Number
        Case Else
4120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4130    End Select
4140    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

4210    With Me

4220      DoCmd.Hourglass True
4230      DoEvents

4240      strDocName = "frmCurrency_Country"
4250      If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
4260        DoCmd.Close acForm, strDocName
4270        DoEvents
4280      End If
4290      strDocName = "frmCurrency_History"
4300      If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
4310        DoCmd.Close acForm, strDocName
4320        DoEvents
4330      End If

          ' ** Make sure all reports are closed.
4340      Do While Reports.Count > 0
4350        DoCmd.Close acReport, Reports(0).Name
4360        DoEvents
4370      Loop

4380      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

4390      Set dbs = CurrentDb
          ' ** Empty tblCurrency_Rate_Staging.
4400      Set qdf = dbs.QueryDefs("qryCurrency_Rate_Sub_03")
4410      qdf.Execute
4420      Set qdf = Nothing
4430      dbs.Close
4440      Set dbs = Nothing

4450      Select Case strCallingForm
          Case "frmMenu_ForeignExchange"
4460        gstrReturningForm = THIS_NAME
4470        gblnSetFocus = True
4480        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Maintenance"
4490      End Select

4500    End With

EXITP:
4510    Set qdf = Nothing
4520    Set dbs = Nothing
4530    Exit Sub

ERRH:
4540    DoCmd.Hourglass False
4550    Select Case ERR.Number
        Case Else
4560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4570    End Select
4580    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4610    With Me
4620      If .cmdResetFilter_raised_focus_dots_img.Visible = True Or .cmdResetFilter_raised_focus_img.Visible = True Then
4630        Select Case blnClearFilter_Focus
            Case True
4640          .cmdResetFilter_raised_semifocus_dots_img.Visible = True
4650          .cmdResetFilter_raised_img.Visible = False
4660        Case False
4670          .cmdResetFilter_raised_img.Visible = True
4680          .cmdResetFilter_raised_semifocus_dots_img.Visible = False
4690        End Select
4700        .cmdResetFilter_raised_focus_img.Visible = False
4710        .cmdResetFilter_raised_focus_dots_img.Visible = False
4720        .cmdResetFilter_sunken_focus_dots_img.Visible = False
4730        .cmdResetFilter_raised_img_dis.Visible = False
4740      End If
4750      If .cmdCountries_raised_focus_dots_img.Visible = True Or .cmdCountries_raised_focus_img.Visible = True Then
4760        Select Case blnCountries_Focus
            Case True
4770          .cmdCountries_raised_semifocus_dots_img.Visible = True
4780          .cmdCountries_raised_img.Visible = False
4790        Case False
4800          .cmdCountries_raised_img.Visible = True
4810          .cmdCountries_raised_semifocus_dots_img.Visible = False
4820        End Select
4830        .cmdCountries_raised_focus_img.Visible = False
4840        .cmdCountries_raised_focus_dots_img.Visible = False
4850        .cmdCountries_sunken_focus_dots_img.Visible = False
4860        .cmdCountries_raised_img_dis.Visible = False
4870      End If
4880    End With

EXITP:
4890    Exit Sub

ERRH:
4900    Select Case ERR.Number
        Case Else
4910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4920    End Select
4930    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

5010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5020    Exit Sub

ERRH:
5030    Select Case ERR.Number
        Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Sub

Private Sub cmdClose_Click()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim msgResponse As VbMsgBoxResult

5110    With Me
5120      DoCmd.Hourglass True
5130      DoEvents
5140      If .currratestage_changed = True And .cmdUpdate.Enabled = True Then
5150        Beep
5160        DoCmd.Hourglass False
5170        msgResponse = MsgBox("You have made changes, but have not saved them." & vbCrLf & vbCrLf & _
              "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved")
5180      Else
5190        msgResponse = vbYes
5200      End If
5210      If msgResponse = vbYes Then
5220        DoCmd.Hourglass True
5230        DoEvents
5240        .TimerInterval = 0&
5250        DoCmd.Close acForm, THIS_NAME
5260      Else
5270        .cmdUpdate.SetFocus
5280      End If
5290    End With

EXITP:
5300    Exit Sub

ERRH:
5310    DoCmd.Hourglass False
5320    Select Case ERR.Number
        Case Else
5330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5340    End Select
5350    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
5470        With Me
5480          intRetVal = 0
5490          .frmCurrency_Rate_Sub.SetFocus
5500          .frmCurrency_Rate_Sub.Form.FocusHolder.SetFocus
5510        End With
5520      End Select
5530    End If

        ' ** Shift keys.
5540    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5550      Select Case intRetVal
          Case vbKeyTab
5560        With Me
5570          intRetVal = 0
5580          If .cmdUpdate.Enabled = True Then
5590            .cmdUpdate.SetFocus
5600          ElseIf .cmdEdit.Enabled = True Then
5610            .cmdEdit.SetFocus
5620          ElseIf .cmdHistory.Enabled = True Then
5630            .cmdHistory.SetFocus
5640          ElseIf .cmdExcel.Enabled = True Then
5650            .cmdExcel.SetFocus
5660          ElseIf .cmdCountries.Enabled = True Then
5670            .cmdCountries.SetFocus
5680          ElseIf .cmdResetFilter.Enabled = True Then
5690            .cmdResetFilter.SetFocus
5700          Else
5710            .opgUnit.SetFocus
5720          End If
5730        End With
5740      End Select
5750    End If

EXITP:
5760    KeyCode = intRetVal
5770    Exit Sub

ERRH:
5780    intRetVal = 0
5790    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5800    Case Else
5810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5820    End Select
5830    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String, strFilter As String

5910    With Me
5920      Select Case .frmCurrency_Rate_Sub.Form.FilterOn
          Case True
5930        strFilter = .frmCurrency_Rate_Sub.Form.Filter
5940      Case False
5950        strFilter = vbNullString
5960      End Select
5970      strDocName = "rptCurrencyRates"
5980      DoCmd.OpenReport strDocName, acViewPreview, , strFilter
5990      DoCmd.Maximize
6000      DoCmd.RunCommand acCmdFitToWindow
6010    End With

EXITP:
6020    Exit Sub

ERRH:
6030    Select Case ERR.Number
        Case Else
6040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6050    End Select
6060    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6120    intShiftDown = (Shift And acShiftMask) > 0
6130    intAltDown = (Shift And acAltMask) > 0
6140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6160      Select Case intRetVal
          Case vbKeyTab
6170        With Me
6180          intRetVal = 0
6190          .cmdPrint.SetFocus
6200        End With
6210      End Select
6220    End If

        ' ** Shift keys.
6230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6240      Select Case intRetVal
          Case vbKeyTab
6250        With Me
6260          intRetVal = 0
6270          If .cmdCountries.Enabled = True Then
6280            .cmdCountries.SetFocus
6290          ElseIf .cmdResetFilter.Enabled = True Then
6300            .cmdResetFilter.SetFocus
6310          Else
6320            .opgUnit.SetFocus
6330          End If
6340        End With
6350      End Select
6360    End If

        ' ** Ctrl-Shift keys.
6370    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6380      Select Case intRetVal
          Case vbKeyTab
6390        With Me
6400          intRetVal = 0
6410          .chkPageOf.SetFocus
6420        End With
6430      End Select
6440    End If

EXITP:
6450    KeyCode = intRetVal
6460    Exit Sub

ERRH:
6470    intRetVal = 0
6480    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6490    Case Else
6500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6510    End Select
6520    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String, strFilter As String

6610    With Me
6620      Select Case .frmCurrency_Rate_Sub.Form.FilterOn
          Case True
6630        strFilter = .frmCurrency_Rate_Sub.Form.Filter
6640      Case False
6650        strFilter = vbNullString
6660      End Select
6670      strDocName = "rptCurrencyRates"
6680      If gblnDev_Debug = True Or CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
6690        DoCmd.OpenReport strDocName, acViewPreview, , strFilter
6700        DoCmd.Maximize
6710        DoCmd.RunCommand acCmdFitToWindow
6720      Else
            '##GTR_Ref: rptCurrencyRates
6730        DoCmd.OpenReport strDocName, acViewNormal, , strFilter
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

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6920    intShiftDown = (Shift And acShiftMask) > 0
6930    intAltDown = (Shift And acAltMask) > 0
6940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6960      Select Case intRetVal
          Case vbKeyTab
6970        With Me
6980          intRetVal = 0
6990          .cmdWord.SetFocus
7000        End With
7010      End Select
7020    End If

        ' ** Shift keys.
7030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7040      Select Case intRetVal
          Case vbKeyTab
7050        With Me
7060          intRetVal = 0
7070          .cmdPreview.SetFocus
7080        End With
7090      End Select
7100    End If

EXITP:
7110    KeyCode = intRetVal
7120    Exit Sub

ERRH:
7130    intRetVal = 0
7140    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7150    Case Else
7160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7170    End Select
7180    Resume EXITP

End Sub

Private Sub cmdWord_Click()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

7210    With Me

7220      If IsNull(.UserReportPath) = True Then
7230        strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
7240      Else
7250        strRptPath = .UserReportPath
7260      End If
7270      strRptCap = "rptCurrencyRates_" & Format$(Date, "yyyymmdd")

7280      strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

7290      If strRptPathFile <> vbNullString Then
7300        strRpt = "rptCurrencyRates"
7310        DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
7320        .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
7330      End If

7340    End With

EXITP:
7350    Exit Sub

ERRH:
7360    Select Case ERR.Number
        Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
7370      Beep
7380      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
7390    Case Else
7400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7410    End Select
7420    Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

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
7590          If .cmdExcel.Enabled = True Then
7600            .cmdExcel.SetFocus
7610          ElseIf .cmdHistory.Enabled = True Then
7620            .cmdHistory.SetFocus
7630          ElseIf .cmdEdit.Enabled = True Then
7640            .cmdEdit.SetFocus
7650          ElseIf .cmdUpdate.Enabled = True Then
7660            .cmdUpdate.SetFocus
7670          Else
7680            .cmdClose.SetFocus
7690          End If
7700        End With
7710      End Select
7720    End If

        ' ** Shift keys.
7730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7740      Select Case intRetVal
          Case vbKeyTab
7750        With Me
7760          intRetVal = 0
7770          .cmdPrint.SetFocus
7780        End With
7790      End Select
7800    End If

EXITP:
7810    KeyCode = intRetVal
7820    Exit Sub

ERRH:
7830    intRetVal = 0
7840    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7850    Case Else
7860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7870    End Select
7880    Resume EXITP

End Sub

Private Sub cmdExcel_Click()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strQry As String, strRptCap As String, strOrderBy As String
        Dim strRptPath As String, strRptPathFile As String
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, blnTmp04 As Boolean

7910    With Me
7920      lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
7930      If lngRecsCur > 0& Then

7940        gstrFormQuerySpec = THIS_NAME

7950        Set dbs = CurrentDb

            ' ** Empty tmpCurrencyRateReport.
7960        Set qdf = dbs.QueryDefs("qryRpt_CurrencyRates_10")
7970        qdf.Execute
7980        Set qdf = Nothing

            ' ** Append qryRpt_CurrencyRates_11 (tblCurrency, with curr_rate1x, curr_rate2x,
            ' ** by specified [activ], [fun], [metl], [bmu], [alt], [unt]) to tmpCurrencyRateReport.
7990        Set qdf = dbs.QueryDefs("qryRpt_CurrencyRates_12")
8000        Select Case .opgActive
            Case .opgActive_optInclude.OptionValue, .opgActive_optOnly.OptionValue
8010          blnTmp04 = True
8020        Case .opgActive_optExclude.OptionValue
8030          blnTmp04 = False
8040        End Select
8050        With qdf.Parameters
8060          ![activ] = blnTmp04
8070        End With
8080        Select Case .opgFund
            Case .opgFund_optInclude.OptionValue, .opgFund_optOnly.OptionValue
8090          blnTmp04 = True
8100        Case .opgFund_optExclude.OptionValue
8110          blnTmp04 = False
8120        End Select
8130        With qdf.Parameters
8140          ![fun] = blnTmp04
8150        End With
8160        Select Case .opgMetal
            Case .opgMetal_optInclude.OptionValue, .opgMetal_optOnly.OptionValue
8170          blnTmp04 = True
8180        Case .opgMetal_optExclude.OptionValue
8190          blnTmp04 = False
8200        End Select
8210        With qdf.Parameters
8220          ![metl] = blnTmp04
8230        End With
8240        Select Case .opgBMU
            Case .opgBMU_optInclude.OptionValue, .opgBMU_optOnly.OptionValue
8250          blnTmp04 = True
8260        Case .opgBMU_optExclude.OptionValue
8270          blnTmp04 = False
8280        End Select
8290        With qdf.Parameters
8300          ![bmu] = blnTmp04
8310        End With
8320        Select Case .opgAlt
            Case .opgAlt_optInclude.OptionValue, .opgAlt_optOnly.OptionValue
8330          blnTmp04 = True
8340        Case .opgAlt_optExclude.OptionValue
8350          blnTmp04 = False
8360        End Select
8370        With qdf.Parameters
8380          ![Alt] = blnTmp04
8390        End With
8400        Select Case .opgUnit
            Case .opgUnit_optInclude.OptionValue, .opgUnit_optOnly.OptionValue
8410          blnTmp04 = True
8420        Case .opgUnit_optExclude.OptionValue
8430          blnTmp04 = False
8440        End Select
8450        With qdf.Parameters
8460          ![unt] = blnTmp04
8470        End With
8480        qdf.Execute
8490        Set qdf = Nothing
8500        dbs.Close
8510        Set dbs = Nothing

8520        If IsNull(.UserReportPath) = True Then
8530          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
8540        Else
8550          strRptPath = .UserReportPath
8560        End If
8570        strRptCap = "rptCurrencyRates_" & Format$(Date, "yyyymmdd")

8580        strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

8590        If strRptPathFile <> vbNullString Then

8600          strOrderBy = .frmCurrency_Rate_Sub.Form.OrderBy

8610          strQry = "qryRpt_CurrencyRates_16_"

8620          Select Case .opgActive
              Case .opgActive_optInclude.OptionValue, .opgActive_optOnly.OptionValue
8630            strQry = strQry & "01_"
8640          Case .opgActive_optExclude.OptionValue
8650            strQry = strQry & "00_"
8660          End Select

8670          Select Case .opgFund
              Case .opgFund_optInclude.OptionValue, .opgFund_optOnly.OptionValue
8680            strQry = strQry & "01_"
8690          Case .opgFund_optExclude.OptionValue
8700            strQry = strQry & "00_"
8710          End Select

8720          Select Case .opgMetal
              Case .opgMetal_optInclude.OptionValue, .opgMetal_optOnly.OptionValue
8730            strQry = strQry & "01_"
8740          Case .opgMetal_optExclude.OptionValue
8750            strQry = strQry & "00_"
8760          End Select

8770          Select Case .opgBMU
              Case .opgBMU_optInclude.OptionValue, .opgBMU_optOnly.OptionValue
8780            strQry = strQry & "01_"
8790          Case .opgBMU_optExclude.OptionValue
8800            strQry = strQry & "00_"
8810          End Select

8820          Select Case .opgAlt
              Case .opgAlt_optInclude.OptionValue, .opgAlt_optOnly.OptionValue
8830            strQry = strQry & "01"
8840          Case .opgAlt_optExclude.OptionValue
8850            strQry = strQry & "00"
8860          End Select

8870          Select Case .opgUnit
              Case .opgUnit_optInclude.OptionValue, .opgUnit_optOnly.OptionValue
8880            strQry = strQry & "01"
8890          Case .opgUnit_optExclude.OptionValue
8900            strQry = strQry & "00"
8910          End Select

8920          blnTmp04 = .chkIncludeNotes

8930          Set dbs = CurrentDb
8940          strTmp01 = dbs.QueryDefs(strQry).SQL
8950          strQry = "qryRpt_CurrencyRates_16"
8960          Set qdf = dbs.QueryDefs(strQry)
8970          With qdf
                ' ** SELECT qryRpt_CurrencyRates_15.[Currency Code], qryRpt_CurrencyRates_15.[Currency Name],
                ' **   qryRpt_CurrencyRates_15.Active, qryRpt_CurrencyRates_15.[Rate 1],
                ' **   qryRpt_CurrencyRates_15.[Rate 2], qryRpt_CurrencyRates_15.[Rate Date],
                ' **   qryRpt_CurrencyRates_15.ISO, qryRpt_CurrencyRates_15.Fund, qryRpt_CurrencyRates_15.Metal,
                ' **   qryRpt_CurrencyRates_15.BMU, qryRpt_CurrencyRates_15.Misc
                ' ** FROM qryRpt_CurrencyRates_15
                ' ** ORDER BY qryRpt_CurrencyRates_15.sort;
8980            strOrderBy = StringReplace(strOrderBy, "[curr_code]", "[Currency Code]")  ' ** Module Function: modStringReplace.
8990            strOrderBy = StringReplace(strOrderBy, "[curr_name]", "[Currency Name]")  ' ** Module Function: modStringReplace.
9000            strOrderBy = StringReplace(strOrderBy, "[curr_active]", "[Active]")  ' ** Module Function: modStringReplace.
9010            strOrderBy = StringReplace(strOrderBy, "[curr_rate1]", "[Rate 1]")  ' ** Module Function: modStringReplace.
9020            strOrderBy = StringReplace(strOrderBy, "[curr_rate2]", "[Rate 2]")  ' ** Module Function: modStringReplace.
9030            strOrderBy = StringReplace(strOrderBy, "[curr_date]", "[Rate Date]")  ' ** Module Function: modStringReplace.
9040            strOrderBy = StringReplace(strOrderBy, "[curr_iso]", "[ISO]")  ' ** Module Function: modStringReplace.
9050            strOrderBy = StringReplace(strOrderBy, "[curr_fund]", "[Fund]")  ' ** Module Function: modStringReplace.
9060            strOrderBy = StringReplace(strOrderBy, "[curr_metal]", "[Metal]")  ' ** Module Function: modStringReplace.
9070            strOrderBy = StringReplace(strOrderBy, "[curr_bmu]", "[BMU]")  ' ** Module Function: modStringReplace.
9080            strOrderBy = StringReplace(strOrderBy, "[curr_alt]", "[Misc]")  ' ** Module Function: modStringReplace.
9090            strOrderBy = StringReplace(strOrderBy, "[curr_unit]", "[Unit]")  ' ** Module Function: modStringReplace.
9100            strOrderBy = StringReplace(strOrderBy, "[curr_notes]", "[Notes]")  ' ** Module Function: modStringReplace.
9110            strOrderBy = StringReplace(strOrderBy, "[curr_word1]", "[Currency Word]")  ' ** Module Function: modStringReplace.
                ' ** [curr_code]
                ' ** [curr_name]
                ' ** [curr_active], [curr_name]
                ' ** [curr_rate1], [curr_word1], [curr_name]
                ' ** [curr_rate2], [curr_word1], [curr_name]
                ' ** [curr_date], [curr_name]
                ' ** [curr_iso]
                ' ** [curr_fund], [curr_name]
                ' ** [curr_metal], [curr_name]
                ' ** [curr_bmu], [curr_name]
                ' ** [curr_alt], [curr_name]
                ' ** [curr_word1], [curr_name]
                ' ** [curr_notes], [curr_name]
9120            strOrderBy = StringReplace(strOrderBy, "[", "qryRpt_CurrencyRates_15.[")  ' ** Module Function: modStringReplace.
9130            intPos01 = InStr(strTmp01, ";")
9140            strTmp01 = Left(strTmp01, (intPos01 - 1)) & ", " & strOrderBy & ";"
9150            If blnTmp04 = True Then
                  ' ** Include Notes. ([Currency Word] is not shown.)
9160              intPos01 = InStr(strTmp01, "FROM qry")
9170              strTmp02 = Left(strTmp01, (intPos01 - 1))
9180              strTmp03 = Mid(strTmp01, intPos01)
9190              If Right(strTmp02, 1) = vbLf Then
9200                strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
9210                strTmp03 = vbLf & strTmp03
9220              End If
9230              If Right(strTmp02, 1) = vbCr Then
9240                strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
9250                strTmp03 = vbCr & strTmp03
9260              End If
9270              If Right(strTmp02, 1) = vbLf Then
9280                strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
9290                strTmp03 = vbLf & strTmp03
9300              End If
9310              strTmp02 = strTmp02 & ", qryRpt_CurrencyRates_15.[Notes] "
9320              strTmp01 = strTmp02 & strTmp03
9330            End If
9340            .SQL = strTmp01
9350          End With
9360          Set qdf = Nothing
9370          dbs.Close
9380          Set dbs = Nothing

              ' ** qryRpt_CurrencyRates_15, without Active, Misc, BMU, Metal, Fund.
              'qryRpt_CurrencyRates_16_00_00_00_00_00
              ' ** qryRpt_CurrencyRates_15, with Misc, without Active, BMU, Metal, Fund.
              'qryRpt_CurrencyRates_16_00_00_00_00_01
              ' ** qryRpt_CurrencyRates_15, with BMU, without Active, Misc, Metal, Fund.
              'qryRpt_CurrencyRates_16_00_00_00_01_00
              ' ** qryRpt_CurrencyRates_15, with BMU, Misc, without Active, Metal, Fund.
              'qryRpt_CurrencyRates_16_00_00_00_01_01
              ' ** qryRpt_CurrencyRates_15, with Metal, without Active, Misc, BMU, Fund.
              'qryRpt_CurrencyRates_16_00_00_01_00_00
              ' ** qryRpt_CurrencyRates_15, with Metal, Misc, without Active, BMU, Fund.
              'qryRpt_CurrencyRates_16_00_00_01_00_01
              ' ** qryRpt_CurrencyRates_15, with Metal, BMU, without Active, Misc, Fund.
              'qryRpt_CurrencyRates_16_00_00_01_01_00
              ' ** qryRpt_CurrencyRates_15, with Metal, BMU, Misc, without Active, Fund.
              'qryRpt_CurrencyRates_16_00_00_01_01_01
              ' ** qryRpt_CurrencyRates_15, with Fund, without Active, Misc, BMU, Metal.
              'qryRpt_CurrencyRates_16_00_01_00_00_00
              ' ** qryRpt_CurrencyRates_15, with Fund, Misc, without Active, BMU, Metal.
              'qryRpt_CurrencyRates_16_00_01_00_00_01
              ' ** qryRpt_CurrencyRates_15, with Fund, BMU, without Active, Misc, Metal.
              'qryRpt_CurrencyRates_16_00_01_00_01_00
              ' ** qryRpt_CurrencyRates_15, with Fund, BMU, Misc, without Active, Metal.
              'qryRpt_CurrencyRates_16_00_01_00_01_01
              ' ** qryRpt_CurrencyRates_15, with Fund, Metal, without Active, Misc, BMU.
              'qryRpt_CurrencyRates_16_00_01_01_00_00
              ' ** qryRpt_CurrencyRates_15, with Fund, Metal, Misc, without Active, BMU.
              'qryRpt_CurrencyRates_16_00_01_01_00_01
              ' ** qryRpt_CurrencyRates_15, with Fund, Metal, BMU, without Active, Misc.
              'qryRpt_CurrencyRates_16_00_01_01_01_00
              ' ** qryRpt_CurrencyRates_15, with Fund, Metal, BMU, Misc, without Active.
              'qryRpt_CurrencyRates_16_00_01_01_01_01
              ' ** qryRpt_CurrencyRates_15, with Active, without Misc, BMU, Metal, Fund.
              'qryRpt_CurrencyRates_16_01_00_00_00_00
              ' ** qryRpt_CurrencyRates_15, with Active, Misc, without BMU, Metal, Fund.
              'qryRpt_CurrencyRates_16_01_00_00_00_01
              ' ** qryRpt_CurrencyRates_15, with Active, BMU, without Misc, Metal, Fund.
              'qryRpt_CurrencyRates_16_01_00_00_01_00
              ' ** qryRpt_CurrencyRates_15, with Active, BMU, Misc, without Metal, Fund.
              'qryRpt_CurrencyRates_16_01_00_00_01_01
              ' ** qryRpt_CurrencyRates_15, with Active, Metal, without Misc, BMU, Fund.
              'qryRpt_CurrencyRates_16_01_00_01_00_00
              ' ** qryRpt_CurrencyRates_15, with Active, Metal, Misc, without BMU, Fund.
              'qryRpt_CurrencyRates_16_01_00_01_00_01
              ' ** qryRpt_CurrencyRates_15, with Active, Metal, BMU, without Misc, Fund.
              'qryRpt_CurrencyRates_16_01_00_01_01_00
              ' ** qryRpt_CurrencyRates_15, with Active, Metal, BMU, Misc, without Fund.
              'qryRpt_CurrencyRates_16_01_00_01_01_01
              ' ** qryRpt_CurrencyRates_15, with Active, Fund, without Misc, BMU, Metal.
              'qryRpt_CurrencyRates_16_01_01_00_00_00
              ' ** qryRpt_CurrencyRates_15, with Active, Fund, Misc, without BMU, Metal.
              'qryRpt_CurrencyRates_16_01_01_00_00_01
              ' ** qryRpt_CurrencyRates_15, with Active, Fund, BMU, without Misc, Metal.
              'qryRpt_CurrencyRates_16_01_01_00_01_00
              ' ** qryRpt_CurrencyRates_15, with Active, Fund, BMU, Misc, without Metal.
              'qryRpt_CurrencyRates_16_01_01_00_01_01
              ' ** qryRpt_CurrencyRates_15, with Active, Fund, Metal, without Misc, BMU.
              'qryRpt_CurrencyRates_16_01_01_01_00_00
              ' ** qryRpt_CurrencyRates_15, with Active, Fund, Metal, Misc, without  BMU.
              'qryRpt_CurrencyRates_16_01_01_01_00_01
              ' ** qryRpt_CurrencyRates_15, with Active, Fund, Metal, BMU, without Misc.
              'qryRpt_CurrencyRates_16_01_01_01_01_00
              ' ** qryRpt_CurrencyRates_15, with Active, Fund, Metal, BMU, Misc.
              'qryRpt_CurrencyRates_16_01_01_01_01_01

9390          DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, True
9400          .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
9410        End If

9420      Else
9430        Beep
9440      End If
9450    End With

EXITP:
9460    Set qdf = Nothing
9470    Set dbs = Nothing
9480    Exit Sub

ERRH:
9490    Select Case ERR.Number
        Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
9500      Beep
9510      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
9520    Case Else
9530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9540    End Select
9550    Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9620    intShiftDown = (Shift And acShiftMask) > 0
9630    intAltDown = (Shift And acAltMask) > 0
9640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9660      Select Case intRetVal
          Case vbKeyTab
9670        With Me
9680          intRetVal = 0
9690          If .cmdHistory.Enabled = True Then
9700            .cmdHistory.SetFocus
9710          ElseIf .cmdEdit.Enabled = True Then
9720            .cmdEdit.SetFocus
9730          ElseIf .cmdUpdate.Enabled = True Then
9740            .cmdUpdate.SetFocus
9750          Else
9760            .cmdClose.SetFocus
9770          End If
9780        End With
9790      End Select
9800    End If

        ' ** Shift keys.
9810    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9820      Select Case intRetVal
          Case vbKeyTab
9830        With Me
9840          intRetVal = 0
9850          .cmdWord.SetFocus
9860        End With
9870      End Select
9880    End If

EXITP:
9890    KeyCode = intRetVal
9900    Exit Sub

ERRH:
9910    intRetVal = 0
9920    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
9930    Case Else
9940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9950    End Select
9960    Resume EXITP

End Sub

Private Sub cmdHistory_Click()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistory_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String
        Dim lngRecs As Long
        Dim blnContinue As Boolean

10010   blnContinue = True

10020   With Me

10030     DoCmd.Hourglass True
10040     DoEvents

10050     Set dbs = CurrentDb

          ' ** If data has been changed, they must update first.
          'MAKE SURE INDEX ON curr_date IS IN ALL MDB'S!

          ' ** tblCurrency_Rate_Staging, just currratestage_changed = True.
10060     Set qdf = dbs.QueryDefs("qryCurrency_Rate_Sub_06")
10070     Set rst = qdf.OpenRecordset
10080     If rst.BOF = True And rst.EOF = True Then
            ' ** No pending updates.
10090       rst.Close
10100     Else
10110       blnContinue = False
10120       DoCmd.Hourglass False
10130       Beep
10140       MsgBox "Changed records were found." & vbCrLf & _
              "You must update to commit those changes before proceding.", vbInformation + vbOKOnly, "Update Pending"
10150       rst.Close
10160     End If
10170     Set rst = Nothing
10180     Set qdf = Nothing

10190     If blnContinue = True Then
            ' ** tblCurrency, linked to tblCurrency_History,
            ' ** just differing rates for same curr_date.
10200       Set qdf = dbs.QueryDefs("qryCurrency_History_04")
10210       Set rst = qdf.OpenRecordset
10220       If rst.BOF = True And rst.EOF = True Then
              ' ** No conflicts.
10230         rst.Close
10240       Else
10250         rst.MoveLast
10260         lngRecs = rst.RecordCount
10270         rst.MoveFirst
10280         rst.Close
10290         strDocName = "frmCurrency_History_Date"
10300         gblnMessage = True
10310         DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME
10320         blnContinue = gblnMessage  ' ** Cancel sets gblnMessage = False.
10330       End If
10340       Set rst = Nothing
10350       Set qdf = Nothing
10360     End If  ' ** blnContinue.

10370     If blnContinue = True Then

            ' ** tblCurrency, not in qryCurrency_History_04 (tblCurrency, linked
            ' ** to tblCurrency_History, just differing rates for same curr_date),
            ' ** qryCurrency_History_05 (tblCurrency, linked to tblCurrency_History,
            ' ** just same dates, same rates), just new dates.
10380       Set qdf = dbs.QueryDefs("qryCurrency_History_06")
10390       Set rst = qdf.OpenRecordset
10400       If rst.BOF = True And rst.EOF = True Then
              ' ** Nothing new, so just proceed to form.
10410         rst.Close
10420         lngRecs = 0&
10430       Else
10440         rst.MoveLast
10450         lngRecs = rst.RecordCount
10460         rst.Close
10470       End If
10480       Set rst = Nothing
10490       Set qdf = Nothing

10500       If lngRecs > 0& Then
              ' ** Append qryCurrency_History_06 (tblCurrency, not in
              ' ** qryCurrency_History_04 (tblCurrency, linked to tblCurrency_History,
              ' ** just differing rates for same curr_date), qryCurrency_History_05
              ' ** (tblCurrency, linked to tblCurrency_History, just same dates, same
              ' ** rates), just new dates) to tblCurrency_History.
10510         Set qdf = dbs.QueryDefs("qryCurrency_History_07")
10520         qdf.Execute dbFailOnError
10530         Set qdf = Nothing
10540       End If

10550       strDocName = "frmCurrency_History"
10560       gblnSetFocus = True
10570       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
10580       If gblnGoToReport = True Then
10590         Forms(strDocName).TimerInterval = 50&
10600         blnGoingToReport = False
10610         blnGTR_Emblem = False
10620       End If

10630     End If  ' ** blnContinue

10640   End With

EXITP:
10650   Set rst = Nothing
10660   Set qdf = Nothing
10670   Set dbs = Nothing
10680   Exit Sub

ERRH:
10690   DoCmd.Hourglass False
10700   Select Case ERR.Number
        Case Else
10710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10720   End Select
10730   Resume EXITP

End Sub

Private Sub cmdHistory_KeyDown(KeyCode As Integer, Shift As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistory_KeyDown"

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
10890         If .cmdEdit.Enabled = True Then
10900           .cmdEdit.SetFocus
10910         ElseIf .cmdUpdate.Enabled = True Then
10920           .cmdUpdate.SetFocus
10930         Else
10940           .cmdClose.SetFocus
10950         End If
10960       End With
10970     End Select
10980   End If

        ' ** Shift keys.
10990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11000     Select Case intRetVal
          Case vbKeyTab
11010       With Me
11020         intRetVal = 0
11030         If .cmdExcel.Enabled = True Then
11040           .cmdExcel.SetFocus
11050         ElseIf .cmdWord.Enabled = True Then
11060           .cmdWord.SetFocus
11070         ElseIf .cmdCountries.Enabled = True Then
11080           .cmdCountries.SetFocus
11090         ElseIf .cmdResetFilter.Enabled = True Then
11100           .cmdResetFilter.SetFocus
11110         Else
11120           .opgUnit.SetFocus
11130         End If
11140       End With
11150     End Select
11160   End If

EXITP:
11170   KeyCode = intRetVal
11180   Exit Sub

ERRH:
11190   intRetVal = 0
11200   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
11210   Case Else
11220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11230   End Select
11240   Resume EXITP

End Sub

Private Sub cmdEdit_Click()

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_Click"

11310   With Me
11320     If .frmCurrency_Rate_Sub.Form.curr_rate1.Locked = True Then

11330       .frmCurrency_Rate_Sub.Form.curr_rate1.Enabled = True
11340       .frmCurrency_Rate_Sub.Form.curr_rate1.Locked = False
11350       .frmCurrency_Rate_Sub.Form.curr_rate1.ForeColor = CLR_VDKGRY
11360       .frmCurrency_Rate_Sub.Form.curr_rate1.BackColor = CLR_WHT
11370       .frmCurrency_Rate_Sub.Form.curr_rate2.Enabled = True
11380       .frmCurrency_Rate_Sub.Form.curr_rate2.Locked = False
11390       .frmCurrency_Rate_Sub.Form.curr_rate2.ForeColor = CLR_VDKGRY
11400       .frmCurrency_Rate_Sub.Form.curr_rate2.BackColor = CLR_WHT
11410       .frmCurrency_Rate_Sub.Form.curr_date.Enabled = True
11420       .frmCurrency_Rate_Sub.Form.curr_date.Locked = False
11430       .frmCurrency_Rate_Sub.Form.curr_date.ForeColor = CLR_VDKGRY
11440       .frmCurrency_Rate_Sub.Form.curr_date.BackColor = CLR_WHT
11450       .frmCurrency_Rate_Sub.Form.curr_convert1a.Enabled = True
11460       .frmCurrency_Rate_Sub.Form.curr_convert1a.Locked = False
11470       .frmCurrency_Rate_Sub.Form.curr_convert1a.ForeColor = CLR_VDKGRY
11480       .frmCurrency_Rate_Sub.Form.curr_convert1a.BackColor = CLR_WHT
11490       .frmCurrency_Rate_Sub.Form.curr_convert1a_box.BackColor = CLR_WHT
11500       .frmCurrency_Rate_Sub.Form.curr_convert1b.Enabled = True
11510       .frmCurrency_Rate_Sub.Form.curr_convert1b.Locked = True
11520       .frmCurrency_Rate_Sub.Form.curr_convert1b.ForeColor = CLR_VDKGRY
11530       .frmCurrency_Rate_Sub.Form.curr_convert1b.BackColor = CLR_WHT
11540       .frmCurrency_Rate_Sub.Form.curr_convert1b_box.BackColor = CLR_WHT
11550       .frmCurrency_Rate_Sub.Form.curr_convert2a.Enabled = True
11560       .frmCurrency_Rate_Sub.Form.curr_convert2a.Locked = False
11570       .frmCurrency_Rate_Sub.Form.curr_convert2a.ForeColor = CLR_VDKGRY
11580       .frmCurrency_Rate_Sub.Form.curr_convert2a.BackColor = CLR_WHT
11590       .frmCurrency_Rate_Sub.Form.curr_convert2a_box.BackColor = CLR_WHT
11600       .frmCurrency_Rate_Sub.Form.curr_convert2b.Enabled = True
11610       .frmCurrency_Rate_Sub.Form.curr_convert2b.Locked = True
11620       .frmCurrency_Rate_Sub.Form.curr_convert2b.ForeColor = CLR_VDKGRY
11630       .frmCurrency_Rate_Sub.Form.curr_convert2b.BackColor = CLR_WHT
11640       .frmCurrency_Rate_Sub.Form.curr_convert2b_box.BackColor = CLR_WHT
11650       .frmCurrency_Rate_Sub.Form.curr_notes.Enabled = True
11660       .frmCurrency_Rate_Sub.Form.curr_notes.Locked = False
11670       .frmCurrency_Rate_Sub.Form.curr_notes.ForeColor = CLR_VDKGRY
11680       .frmCurrency_Rate_Sub.Form.curr_notes.BackColor = CLR_WHT
11690       .frmCurrency_Rate_Sub.Form.curr_active.Enabled = True
11700       .frmCurrency_Rate_Sub.Form.curr_active.Locked = False
11710       .frmCurrency_Rate_Sub.Form.curr_active_box1.ForeColor = CLR_WHT
11720       .frmCurrency_Rate_Sub.Form.curr_active_box2.ForeColor = CLR_LTBLU2

11730       .frmCurrency_Rate_Sub.Form.currsym_symbol_box1a.Visible = True
11740       .frmCurrency_Rate_Sub.Form.currsym_symbol_box1b.Visible = True
11750       .frmCurrency_Rate_Sub.Form.currsym_symbol_box2a.Visible = True
11760       .frmCurrency_Rate_Sub.Form.currsym_symbol_box2b.Visible = True

11770       .frmCurrency_Rate_Sub.Form.currsym_symbol1b_lbl_img.Visible = True
11780       .frmCurrency_Rate_Sub.Form.currsym_symbol1b_lbl_img_dis.Visible = False
11790       .frmCurrency_Rate_Sub.Form.currsym_symbol2a_lbl_img.Visible = True
11800       .frmCurrency_Rate_Sub.Form.currsym_symbol2a_lbl_img_dis.Visible = False

11810       .frmCurrency_Rate_Sub.SetFocus
11820       .frmCurrency_Rate_Sub.Form.curr_name.SetFocus

11830     Else
11840       If .currratestage_changed = False And .cmdUpdate.Enabled = False Then

11850         .frmCurrency_Rate_Sub.SetFocus
11860         .frmCurrency_Rate_Sub.Form.FocusHolder.SetFocus

11870         .frmCurrency_Rate_Sub.Form.curr_rate1.Enabled = True
11880         .frmCurrency_Rate_Sub.Form.curr_rate1.Locked = True
11890         .frmCurrency_Rate_Sub.Form.curr_rate1.ForeColor = CLR_DISABLED_FG
11900         .frmCurrency_Rate_Sub.Form.curr_rate1.BackColor = CLR_DISABLED_BG
11910         .frmCurrency_Rate_Sub.Form.curr_rate2.Enabled = True
11920         .frmCurrency_Rate_Sub.Form.curr_rate2.Locked = True
11930         .frmCurrency_Rate_Sub.Form.curr_rate2.ForeColor = CLR_DISABLED_FG
11940         .frmCurrency_Rate_Sub.Form.curr_rate2.BackColor = CLR_DISABLED_BG
11950         .frmCurrency_Rate_Sub.Form.curr_date.Enabled = True
11960         .frmCurrency_Rate_Sub.Form.curr_date.Locked = True
11970         .frmCurrency_Rate_Sub.Form.curr_date.ForeColor = CLR_DISABLED_FG
11980         .frmCurrency_Rate_Sub.Form.curr_date.BackColor = CLR_DISABLED_BG
11990         .frmCurrency_Rate_Sub.Form.curr_convert1a.Enabled = True
12000         .frmCurrency_Rate_Sub.Form.curr_convert1a.Locked = True
12010         .frmCurrency_Rate_Sub.Form.curr_convert1a.ForeColor = CLR_DISABLED_FG
12020         .frmCurrency_Rate_Sub.Form.curr_convert1a.BackColor = CLR_DISABLED_BG
12030         .frmCurrency_Rate_Sub.Form.curr_convert1a_box.BackColor = CLR_DISABLED_BG
12040         .frmCurrency_Rate_Sub.Form.curr_convert1b.Enabled = True
12050         .frmCurrency_Rate_Sub.Form.curr_convert1b.Locked = True
12060         .frmCurrency_Rate_Sub.Form.curr_convert1b.ForeColor = CLR_DISABLED_FG
12070         .frmCurrency_Rate_Sub.Form.curr_convert1b.BackColor = CLR_DISABLED_BG
12080         .frmCurrency_Rate_Sub.Form.curr_convert1b_box.BackColor = CLR_DISABLED_BG
12090         .frmCurrency_Rate_Sub.Form.curr_convert2a.Enabled = True
12100         .frmCurrency_Rate_Sub.Form.curr_convert2a.Locked = True
12110         .frmCurrency_Rate_Sub.Form.curr_convert2a.ForeColor = CLR_DISABLED_FG
12120         .frmCurrency_Rate_Sub.Form.curr_convert2a.BackColor = CLR_DISABLED_BG
12130         .frmCurrency_Rate_Sub.Form.curr_convert2a_box.BackColor = CLR_DISABLED_BG
12140         .frmCurrency_Rate_Sub.Form.curr_convert2b.Enabled = True
12150         .frmCurrency_Rate_Sub.Form.curr_convert2b.Locked = True
12160         .frmCurrency_Rate_Sub.Form.curr_convert2b.ForeColor = CLR_DISABLED_FG
12170         .frmCurrency_Rate_Sub.Form.curr_convert2b.BackColor = CLR_DISABLED_BG
12180         .frmCurrency_Rate_Sub.Form.curr_convert2b_box.BackColor = CLR_DISABLED_BG
12190         .frmCurrency_Rate_Sub.Form.curr_notes.Enabled = True
12200         .frmCurrency_Rate_Sub.Form.curr_notes.Locked = True
12210         .frmCurrency_Rate_Sub.Form.curr_notes.ForeColor = CLR_DISABLED_FG
12220         .frmCurrency_Rate_Sub.Form.curr_notes.BackColor = CLR_DISABLED_BG
12230         .frmCurrency_Rate_Sub.Form.curr_active.Enabled = True
12240         .frmCurrency_Rate_Sub.Form.curr_active.Locked = True
12250         .frmCurrency_Rate_Sub.Form.curr_active_box1.ForeColor = MY_CLR_BGE 'CLR_IEI_BLU_DIS
12260         .frmCurrency_Rate_Sub.Form.curr_active_box2.ForeColor = WIN_CLR_DISR 'CLR_LTBLU2

12270         .frmCurrency_Rate_Sub.Form.currsym_symbol_box1a.Visible = False
12280         .frmCurrency_Rate_Sub.Form.currsym_symbol_box1b.Visible = False
12290         .frmCurrency_Rate_Sub.Form.currsym_symbol_box2a.Visible = False
12300         .frmCurrency_Rate_Sub.Form.currsym_symbol_box2b.Visible = False

12310         .frmCurrency_Rate_Sub.Form.currsym_symbol1b_lbl_img_dis.Visible = True
12320         .frmCurrency_Rate_Sub.Form.currsym_symbol1b_lbl_img.Visible = False
12330         .frmCurrency_Rate_Sub.Form.currsym_symbol2a_lbl_img_dis.Visible = False
12340         .frmCurrency_Rate_Sub.Form.currsym_symbol2a_lbl_img.Visible = False

12350       End If
12360     End If
12370   End With

EXITP:
12380   Exit Sub

ERRH:
12390   Select Case ERR.Number
        Case Else
12400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12410   End Select
12420   Resume EXITP

End Sub

Private Sub cmdEdit_KeyDown(KeyCode As Integer, Shift As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12520   intShiftDown = (Shift And acShiftMask) > 0
12530   intAltDown = (Shift And acAltMask) > 0
12540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12560     Select Case intRetVal
          Case vbKeyTab
12570       With Me
12580         intRetVal = 0
12590         If .cmdUpdate.Enabled = True Then
12600           .cmdUpdate.SetFocus
12610         Else
12620           .cmdClose.SetFocus
12630         End If
12640       End With
12650     End Select
12660   End If

        ' ** Shift keys.
12670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12680     Select Case intRetVal
          Case vbKeyTab
12690       With Me
12700         intRetVal = 0
12710         If .cmdHistory.Enabled = True Then
12720           .cmdHistory.SetFocus
12730         ElseIf .cmdExcel.Enabled = True Then
12740           .cmdExcel.SetFocus
12750         ElseIf .cmdCountries.Enabled = True Then
12760           .cmdCountries.SetFocus
12770         ElseIf .cmdResetFilter.Enabled = True Then
12780           .cmdResetFilter.SetFocus
12790         Else
12800           .opgUnit.SetFocus
12810         End If
12820       End With
12830     End Select
12840   End If

EXITP:
12850   KeyCode = intRetVal
12860   Exit Sub

ERRH:
12870   intRetVal = 0
12880   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
12890   Case Else
12900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12910   End Select
12920   Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnContinue As Boolean

13010   blnContinue = True

13020   With Me

13030     DoCmd.Hourglass True
13040     DoEvents

13050     Set dbs = CurrentDb

          ' ** tblCurrency_Rate_Staging, just currratestage_changed = True.
13060     Set qdf = dbs.QueryDefs("qryCurrency_Rate_Sub_06")
13070     Set rst = qdf.OpenRecordset
13080     If rst.BOF = True And rst.EOF = True Then
13090       blnContinue = False
13100       DoCmd.Hourglass False
13110       Beep
13120       MsgBox "No changed records were found.", vbInformation + vbOKOnly, "Nothing To Do"
13130       .currratestage_changed = False
13140       rst.Close
13150     Else
13160       rst.Close
13170     End If
13180     Set rst = Nothing
13190     Set qdf = Nothing

13200     If blnContinue = True Then

            ' ** Update qryCurrency_Rate_Sub_07 (qryCurrency_Rate_Sub_06 (tblCurrency_Rate_Staging, just
            ' ** currratestage_changed = True), linked to tblCurrency, with .._new fields), for curr_rate1.
13210       Set qdf = dbs.QueryDefs("qryCurrency_Rate_Sub_08a")
13220       qdf.Execute
13230       Set qdf = Nothing

            ' ** Update qryCurrency_Rate_Sub_07 (qryCurrency_Rate_Sub_06 (tblCurrency_Rate_Staging, just
            ' ** currratestage_changed = True), linked to tblCurrency, with .._new fields), for curr_rate2.
13240       Set qdf = dbs.QueryDefs("qryCurrency_Rate_Sub_08b")
13250       qdf.Execute
13260       Set qdf = Nothing

            ' ** Update qryCurrency_Rate_Sub_07 (qryCurrency_Rate_Sub_06 (tblCurrency_Rate_Staging, just
            ' ** currratestage_changed = True), linked to tblCurrency, with .._new fields), for curr_date.
13270       Set qdf = dbs.QueryDefs("qryCurrency_Rate_Sub_08c")
13280       qdf.Execute
13290       Set qdf = Nothing

            ' ** Update qryCurrency_Rate_Sub_07 (qryCurrency_Rate_Sub_06 (tblCurrency_Rate_Staging, just
            ' ** currratestage_changed = True), linked to tblCurrency, with .._new fields), for curr_notes.
13300       Set qdf = dbs.QueryDefs("qryCurrency_Rate_Sub_08d")
13310       qdf.Execute
13320       Set qdf = Nothing

            ' ** Update qryCurrency_Rate_Sub_07 (qryCurrency_Rate_Sub_06 (tblCurrency_Rate_Staging, just
            ' ** currratestage_changed = True), linked to tblCurrency, with .._new fields), for curr_active.
13330       Set qdf = dbs.QueryDefs("qryCurrency_Rate_Sub_08e")
13340       qdf.Execute
13350       Set qdf = Nothing

13360       Beep
13370       DoCmd.Hourglass False
13380       MsgBox "Updates completed.", vbInformation + vbOKOnly, ("Update Successful" & Space(40))

13390       .currratestage_changed = False
13400       .cmdClose.SetFocus
13410       .cmdUpdate.Enabled = False

13420     End If

13430     dbs.Close

13440     DoCmd.Hourglass False

13450   End With

EXITP:
13460   Set rst = Nothing
13470   Set qdf = Nothing
13480   Set dbs = Nothing
13490   Exit Sub

ERRH:
13500   DoCmd.Hourglass False
13510   Select Case ERR.Number
        Case Else
13520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13530   End Select
13540   Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

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
          Case vbKeyTab
13670       With Me
13680         intRetVal = 0
13690         .cmdClose.SetFocus
13700       End With
13710     End Select
13720   End If

        ' ** Shift keys.
13730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13740     Select Case intRetVal
          Case vbKeyTab
13750       With Me
13760         intRetVal = 0
13770         If .cmdEdit.Enabled = True Then
13780           .cmdEdit.SetFocus
13790         ElseIf .cmdHistory.Enabled = True Then
13800           .cmdHistory.SetFocus
13810         ElseIf .cmdExcel.Enabled = True Then
13820           .cmdExcel.SetFocus
13830         ElseIf .cmdCountries.Enabled = True Then
13840           .cmdCountries.SetFocus
13850         ElseIf .cmdResetFilter.Enabled = True Then
13860           .cmdResetFilter.SetFocus
13870         Else
13880           .opgUnit.SetFocus
13890         End If
13900       End With
13910     End Select
13920   End If

EXITP:
13930   KeyCode = intRetVal
13940   Exit Sub

ERRH:
13950   intRetVal = 0
13960   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
13970   Case Else
13980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13990   End Select
14000   Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

14110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14120   Exit Sub

ERRH:
14130   Select Case ERR.Number
        Case Else
14140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14150   End Select
14160   Resume EXITP

End Sub

Private Sub cmbCurrencies_AfterUpdate()

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbCurrencies_AfterUpdate"

        Dim lngCurrID As Long

14210   With Me
14220     If IsNull(.cmbCurrencies) = False Then
14230       lngCurrID = .cmbCurrencies.Column(CBX_C_CURRID)
14240       blnFromCombo = True
14250       .frmCurrency_Rate_Sub.Form.MoveRec 0, lngCurrID  ' ** Form Procedure: frmCurrency_Rate_Sub.
14260       DoEvents
14270       .frmCurrency_Rate_Sub.SetFocus
14280     End If
14290   End With

EXITP:
14300   Exit Sub

ERRH:
14310   Select Case ERR.Number
        Case Else
14320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14330   End Select
14340   Resume EXITP

End Sub

Private Sub cmbCurrencies_KeyDown(KeyCode As Integer, Shift As Integer)

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbCurrencies_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14420   intShiftDown = (Shift And acShiftMask) > 0
14430   intAltDown = (Shift And acAltMask) > 0
14440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14470       With Me
14480         intRetVal = 0
14490         lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
14500         If lngRecsCur > 0& Then
14510           .frmCurrency_Rate_Sub.SetFocus
14520         Else
14530           .opgActive.SetFocus
14540         End If
14550       End With
14560     End Select
14570   End If

        ' ** Shift keys.
14580   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14590     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14600       With Me
14610         intRetVal = 0
14620         .cmdClose.SetFocus
14630       End With
14640     End Select
14650   End If

EXITP:
14660   KeyCode = intRetVal
14670   Exit Sub

ERRH:
14680   intRetVal = 0
14690   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
14700   Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
14710   Case Else
14720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14730   End Select
14740   Resume EXITP

End Sub

Private Sub cmbCurrencies_Exit(Cancel As Integer)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbCurrencies_Exit"

        Dim lngCurrID As Long

14810   With Me
14820     Select Case blnFromCombo
          Case True
14830       blnFromCombo = False
14840     Case False
14850       If IsNull(.cmbCurrencies) = False Then
14860         lngCurrID = .cmbCurrencies.Column(CBX_C_CURRID)
14870         .frmCurrency_Rate_Sub.Form.MoveRec 0, lngCurrID  ' ** Form Procedure: frmCurrency_Rate_Sub.
14880         DoEvents
14890         .frmCurrency_Rate_Sub.SetFocus
14900       End If
14910     End Select
14920   End With

EXITP:
14930   Exit Sub

ERRH:
14940   Select Case ERR.Number
        Case Else
14950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14960   End Select
14970   Resume EXITP

End Sub

Public Sub opgCurrencySource_AfterUpdate()

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCurrencySource_AfterUpdate"

        Dim lngCurrID As Long, lngCurrs As Long
        Dim strFilter As String
        Dim blnParams As Boolean
        Dim intActive As Integer, intFund As Integer, intMetal As Integer, intBMU As Integer, intAlt As Integer, intUnit As Integer
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim lngX As Long

15010   With Me

15020     Select Case IsNull(.cmbCurrencies)
          Case True
15030       lngCurrID = 0&
15040     Case False
15050       lngCurrID = .cmbCurrencies.Column(CBX_C_CURRID)
15060     End Select

15070     strFilter = .frmCurrency_Rate_Sub.Form.Filter
15080     If strFilter = vbNullString Then
15090       blnParams = False
15100     Else
15110       blnParams = True
15120     End If

15130     intActive = 0: intFund = 0: intMetal = 0: intBMU = 0: intAlt = 0: intUnit = 0

15140     If blnParams = True Then
15150       intPos01 = InStr(strFilter, "[curr_active]")
15160       If intPos01 > 0 Then
15170         strTmp01 = Mid(strFilter, intPos01)
15180         intPos01 = InStr(strTmp01, FAND)
15190         If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, intPos01))
15200         If InStr(strTmp01, "True") > 0 Then intActive = 1 Else intActive = 2
15210       End If
15220       intPos01 = InStr(strFilter, "[curr_fund]")
15230       If intPos01 > 0 Then
15240         strTmp01 = Mid(strFilter, intPos01)
15250         intPos01 = InStr(strTmp01, FAND)
15260         If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, intPos01))
15270         If InStr(strTmp01, "True") > 0 Then intFund = 1 Else intFund = 2
15280       End If
15290       intPos01 = InStr(strFilter, "[curr_metal]")
15300       If intPos01 > 0 Then
15310         strTmp01 = Mid(strFilter, intPos01)
15320         intPos01 = InStr(strTmp01, FAND)
15330         If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, intPos01))
15340         If InStr(strTmp01, "True") > 0 Then intMetal = 1 Else intMetal = 2
15350       End If
15360       intPos01 = InStr(strFilter, "[curr_bmu]")
15370       If intPos01 > 0 Then
15380         strTmp01 = Mid(strFilter, intPos01)
15390         intPos01 = InStr(strTmp01, FAND)
15400         If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, intPos01))
15410         If InStr(strTmp01, "True") > 0 Then intBMU = 1 Else intBMU = 2
15420       End If
15430       intPos01 = InStr(strFilter, "[curr_alt]")
15440       If intPos01 > 0 Then
15450         strTmp01 = Mid(strFilter, intPos01)
15460         intPos01 = InStr(strTmp01, FAND)
15470         If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, intPos01))
15480         If InStr(strTmp01, "True") > 0 Then intAlt = 1 Else intAlt = 2
15490       End If
15500       intPos01 = InStr(strFilter, "[curr_unit]")
15510       If intPos01 > 0 Then
15520         strTmp01 = Mid(strFilter, intPos01)
15530         intPos01 = InStr(strTmp01, FAND)
15540         If intPos01 > 0 Then strTmp01 = Trim(Left(strTmp01, intPos01))
15550         If InStr(strTmp01, "True") > 0 Then intUnit = 1 Else intUnit = 2
15560       End If

15570     End If

15580     CurrSource_Set intActive, intFund, intMetal, intBMU, intAlt, intUnit  ' ** Module Function: modCurrencyFuncs.

          ' ** No Params:
          ' **   strFilter = vbNullString
          ' ** All Others:
          ' **   None  = ([curr_active] = True Or [curr_active] = False)
          ' **   True  = ([curr_active] = True Or [curr_active] = True)
          ' **   False = ([curr_active] = False Or [curr_active] = False)

15590     Select Case .opgCurrencySource
          Case .opgCurrencySource_optCurrWord.OptionValue
15600       .opgCurrencySource_optCurrWord_lbl.FontBold = True
15610       .opgCurrencySource_optCurrCode_lbl.FontBold = False
15620       .opgCurrencySource_optCurrISO_lbl.FontBold = False
15630       .opgCurrencySource_optCurrName_lbl.FontBold = False
15640       .opgCurrencySource_optCountry_lbl.FontBold = False
15650       Select Case blnParams
            Case True
              ' ** tblCurrency, by curr_word1, by specified CurrSource_Get().
15660         .cmbCurrencies.RowSource = "qryCurrency_Rate_02f"
15670       Case False
              ' ** tblCurrency, by curr_word1, for All.
15680         .cmbCurrencies.RowSource = "qryCurrency_Rate_02a"
15690       End Select
15700     Case .opgCurrencySource_optCurrCode.OptionValue
15710       .opgCurrencySource_optCurrWord_lbl.FontBold = False
15720       .opgCurrencySource_optCurrCode_lbl.FontBold = True
15730       .opgCurrencySource_optCurrISO_lbl.FontBold = False
15740       .opgCurrencySource_optCurrName_lbl.FontBold = False
15750       .opgCurrencySource_optCountry_lbl.FontBold = False
15760       Select Case blnParams
            Case True
              ' ** tblCurrency, by curr_code, by specified CurrSource_Get().
15770         .cmbCurrencies.RowSource = "qryCurrency_Rate_02g"
15780       Case False
              ' ** tblCurrency, by curr_code, for All.
15790         .cmbCurrencies.RowSource = "qryCurrency_Rate_02b"
15800       End Select
15810     Case .opgCurrencySource_optCurrISO.OptionValue
15820       .opgCurrencySource_optCurrWord_lbl.FontBold = False
15830       .opgCurrencySource_optCurrCode_lbl.FontBold = False
15840       .opgCurrencySource_optCurrISO_lbl.FontBold = True
15850       .opgCurrencySource_optCurrName_lbl.FontBold = False
15860       .opgCurrencySource_optCountry_lbl.FontBold = False
15870       Select Case blnParams
            Case True
              ' ** tblCurrency, by curr_iso, by specified CurrSource_Get().
15880         .cmbCurrencies.RowSource = "qryCurrency_Rate_02h"
15890       Case False
              ' ** tblCurrency, by curr_iso, for All.
15900         .cmbCurrencies.RowSource = "qryCurrency_Rate_02c"
15910       End Select
15920     Case .opgCurrencySource_optCurrName.OptionValue
15930       .opgCurrencySource_optCurrWord_lbl.FontBold = False
15940       .opgCurrencySource_optCurrCode_lbl.FontBold = False
15950       .opgCurrencySource_optCurrISO_lbl.FontBold = False
15960       .opgCurrencySource_optCurrName_lbl.FontBold = True
15970       .opgCurrencySource_optCountry_lbl.FontBold = False
15980       Select Case blnParams
            Case True
              ' ** tblCurrency, by curr_name, by specified CurrSource_Get().
15990         .cmbCurrencies.RowSource = "qryCurrency_Rate_02i"
16000       Case False
              ' ** tblCurrency, by curr_name, for All.
16010         .cmbCurrencies.RowSource = "qryCurrency_Rate_02d"
16020       End Select
16030     Case .opgCurrencySource_optCountry.OptionValue
16040       .opgCurrencySource_optCurrWord_lbl.FontBold = False
16050       .opgCurrencySource_optCurrCode_lbl.FontBold = False
16060       .opgCurrencySource_optCurrISO_lbl.FontBold = False
16070       .opgCurrencySource_optCurrName_lbl.FontBold = False
16080       .opgCurrencySource_optCountry_lbl.FontBold = True
16090       Select Case blnParams
            Case True
              ' ** tblCurrency, by country_name, by specified CurrSource_Get().
16100         .cmbCurrencies.RowSource = "qryCurrency_Rate_02j"
16110       Case False
              ' ** tblCurrency, by country_name, for All.
16120         .cmbCurrencies.RowSource = "qryCurrency_Rate_02e"
16130       End Select
16140     End Select

16150     If lngCurrID > 0& Then
16160       lngCurrs = .cmbCurrencies.ListCount
16170       For lngX = 0& To (lngCurrs - 1&)
16180         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID Then
16190           .cmbCurrencies = lngCurrID
16200           Exit For
16210         End If
16220       Next
16230     End If

16240     .frmCurrency_Rate_Sub.SetFocus

16250   End With

EXITP:
16260   Exit Sub

ERRH:
16270   Select Case ERR.Number
        Case Else
16280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16290   End Select
16300   Resume EXITP

End Sub

Private Sub opgCurrencySource_optCurrWord_KeyDown(KeyCode As Integer, Shift As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCurrencySource_optCurrWord_KeyDown"

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
16490         .cmbCurrencies.SetFocus
16500       End With
16510     End Select
16520   End If

        ' ** Shift keys.
16530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16550       With Me
16560         intRetVal = 0
16570         .cmdClose.SetFocus
16580       End With
16590     End Select
16600   End If

EXITP:
16610   KeyCode = intRetVal
16620   Exit Sub

ERRH:
16630   intRetVal = 0
16640   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
16650   Case Else
16660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16670   End Select
16680   Resume EXITP

End Sub

Private Sub opgCurrencySource_optCurrCode_KeyDown(KeyCode As Integer, Shift As Integer)

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCurrencySource_optCurrCode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16720   intShiftDown = (Shift And acShiftMask) > 0
16730   intAltDown = (Shift And acAltMask) > 0
16740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16770       With Me
16780         intRetVal = 0
16790         .cmbCurrencies.SetFocus
16800       End With
16810     End Select
16820   End If

        ' ** Shift keys.
16830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16850       With Me
16860         intRetVal = 0
16870         .cmdClose.SetFocus
16880       End With
16890     End Select
16900   End If

EXITP:
16910   KeyCode = intRetVal
16920   Exit Sub

ERRH:
16930   intRetVal = 0
16940   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
16950   Case Else
16960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16970   End Select
16980   Resume EXITP

End Sub

Private Sub opgCurrencySource_optCurrISO_KeyDown(KeyCode As Integer, Shift As Integer)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCurrencySource_optCurrISO_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17020   intShiftDown = (Shift And acShiftMask) > 0
17030   intAltDown = (Shift And acAltMask) > 0
17040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17070       With Me
17080         intRetVal = 0
17090         .cmbCurrencies.SetFocus
17100       End With
17110     End Select
17120   End If

        ' ** Shift keys.
17130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17150       With Me
17160         intRetVal = 0
17170         .cmdClose.SetFocus
17180       End With
17190     End Select
17200   End If

EXITP:
17210   KeyCode = intRetVal
17220   Exit Sub

ERRH:
17230   intRetVal = 0
17240   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
17250   Case Else
17260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17270   End Select
17280   Resume EXITP

End Sub

Private Sub opgCurrencySource_optCurrName_KeyDown(KeyCode As Integer, Shift As Integer)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCurrencySource_optCurrName_KeyDown"

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
17390         .cmbCurrencies.SetFocus
17400       End With
17410     End Select
17420   End If

        ' ** Shift keys.
17430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17450       With Me
17460         intRetVal = 0
17470         .cmdClose.SetFocus
17480       End With
17490     End Select
17500   End If

EXITP:
17510   KeyCode = intRetVal
17520   Exit Sub

ERRH:
17530   intRetVal = 0
17540   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
17550   Case Else
17560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17570   End Select
17580   Resume EXITP

End Sub

Private Sub opgCurrencySource_optCountry_KeyDown(KeyCode As Integer, Shift As Integer)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCurrencySource_optCountry_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17620   intShiftDown = (Shift And acShiftMask) > 0
17630   intAltDown = (Shift And acAltMask) > 0
17640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17670       With Me
17680         intRetVal = 0
17690         .cmbCurrencies.SetFocus
17700       End With
17710     End Select
17720   End If

        ' ** Shift keys.
17730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17750       With Me
17760         intRetVal = 0
17770         .cmdClose.SetFocus
17780       End With
17790     End Select
17800   End If

EXITP:
17810   KeyCode = intRetVal
17820   Exit Sub

ERRH:
17830   intRetVal = 0
17840   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
17850   Case Else
17860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17870   End Select
17880   Resume EXITP

End Sub

Private Sub opgActive_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_box_MouseMove"

17910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17920   Exit Sub

ERRH:
17930   Select Case ERR.Number
        Case Else
17940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17950   End Select
17960   Resume EXITP

End Sub

Public Sub opgActive_AfterUpdate()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_AfterUpdate"

        Dim lngCurrID1 As Long, lngCurrID2 As Long, lngCurrs As Long
        Dim strFilter As String
        Dim blnFound1 As Boolean, blnFound2 As Boolean
        Dim lngX As Long

18010   With Me

18020     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
18030     If lngRecsCur > 0& Then
18040       lngCurrID1 = .frmCurrency_Rate_Sub.Form.curr_id
18050     Else
18060       lngCurrID1 = 0&
18070     End If

18080     lngCurrs = .cmbCurrencies.ListCount
18090     If lngCurrs > 0& Then
18100       Select Case IsNull(.cmbCurrencies)
            Case True
18110         lngCurrID2 = 0&
18120       Case False
18130         lngCurrID2 = .cmbCurrencies.Column(CBX_C_CURRID)
18140       End Select
18150     Else
18160       lngCurrID2 = 0&
18170     End If

18180     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
18190       .opgActive_optInclude_lbl.FontBold = True
18200       .opgActive_optExclude_lbl.FontBold = False
18210       .opgActive_optOnly_lbl.FontBold = False
18220       strFilter = vbNullString
18230     Case .opgActive_optExclude.OptionValue
18240       .opgActive_optInclude_lbl.FontBold = False
18250       .opgActive_optExclude_lbl.FontBold = True
18260       .opgActive_optOnly_lbl.FontBold = False
18270       strFilter = "[curr_active] = False"
18280     Case .opgActive_optOnly.OptionValue
18290       .opgActive_optInclude_lbl.FontBold = False
18300       .opgActive_optExclude_lbl.FontBold = False
18310       .opgActive_optOnly_lbl.FontBold = True
18320       strFilter = "[curr_active] = True"
18330       If .opgFund <> .opgFund_optExclude.OptionValue Then
18340         .opgFund = .opgFund_optExclude.OptionValue
18350         .opgFund_optExclude_lbl.FontBold = True
18360         .opgFund_optInclude_lbl.FontBold = False
18370         .opgFund_optOnly_lbl.FontBold = False
18380       End If
18390       If .opgBMU <> .opgBMU_optExclude.OptionValue Then
18400         .opgBMU = .opgBMU_optExclude.OptionValue
18410         .opgBMU_optExclude_lbl.FontBold = True
18420         .opgBMU_optInclude_lbl.FontBold = False
18430         .opgBMU_optOnly_lbl.FontBold = False
18440       End If
18450       If .opgMetal <> .opgMetal_optExclude.OptionValue Then
18460         .opgMetal = .opgMetal_optExclude.OptionValue
18470         .opgMetal_optExclude_lbl.FontBold = True
18480         .opgMetal_optInclude_lbl.FontBold = False
18490         .opgMetal_optOnly_lbl.FontBold = False
18500       End If
18510       If .opgAlt <> .opgAlt_optExclude.OptionValue Then
18520         .opgAlt = .opgAlt_optExclude.OptionValue
18530         .opgAlt_optExclude_lbl.FontBold = True
18540         .opgAlt_optInclude_lbl.FontBold = False
18550         .opgAlt_optOnly_lbl.FontBold = False
18560       End If
18570       If .opgUnit <> .opgUnit_optExclude.OptionValue Then
18580         .opgUnit = .opgUnit_optExclude.OptionValue
18590         .opgUnit_optExclude_lbl.FontBold = True
18600         .opgUnit_optInclude_lbl.FontBold = False
18610         .opgUnit_optOnly_lbl.FontBold = False
18620       End If
18630     End Select

18640     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
            ' ** No change.
18650     Case .opgFund_optExclude.OptionValue
18660       If strFilter <> vbNullString Then strFilter = strFilter & FAND
18670       strFilter = strFilter & "[curr_fund] = False"
18680     Case .opgFund_optOnly.OptionValue
18690       If strFilter <> vbNullString Then strFilter = strFilter & FAND
18700       strFilter = strFilter & "[curr_fund] = True"
18710     End Select

18720     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
            ' ** No change.
18730     Case .opgBMU_optExclude.OptionValue
18740       If strFilter <> vbNullString Then strFilter = strFilter & FAND
18750       strFilter = strFilter & "[curr_bmu] = False"
18760     Case .opgBMU_optOnly.OptionValue
18770       If strFilter <> vbNullString Then strFilter = strFilter & FAND
18780       strFilter = strFilter & "[curr_bmu] = True"
18790     End Select

18800     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
            ' ** No change.
18810     Case .opgMetal_optExclude.OptionValue
18820       If strFilter <> vbNullString Then strFilter = strFilter & FAND
18830       strFilter = strFilter & "[curr_metal] = False"
18840     Case .opgMetal_optOnly.OptionValue
18850       If strFilter <> vbNullString Then strFilter = strFilter & FAND
18860       strFilter = strFilter & "[curr_metal] = True"
18870     End Select

18880     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
            ' ** No change.
18890     Case .opgAlt_optExclude.OptionValue
18900       If strFilter <> vbNullString Then strFilter = strFilter & FAND
18910       strFilter = strFilter & "[curr_alt] = False"
18920     Case .opgAlt_optOnly.OptionValue
18930       If strFilter <> vbNullString Then strFilter = strFilter & FAND
18940       strFilter = strFilter & "[curr_alt] = True"
18950     End Select

18960     Select Case .opgUnit
          Case .opgUnit_optInclude.OptionValue
            ' ** No change.
18970     Case .opgUnit_optExclude.OptionValue
18980       If strFilter <> vbNullString Then strFilter = strFilter & FAND
18990       strFilter = strFilter & "[curr_unit] = False"
19000     Case .opgUnit_optOnly.OptionValue
19010       If strFilter <> vbNullString Then strFilter = strFilter & FAND
19020       strFilter = strFilter & "[curr_unit] = True"
19030     End Select

19040     .frmCurrency_Rate_Sub.Form.Filter = strFilter
19050     If strFilter <> vbNullString Then
19060       .frmCurrency_Rate_Sub.Form.FilterOn = True
19070     Else
19080       .frmCurrency_Rate_Sub.Form.FilterOn = False
19090     End If
19100     DoEvents

19110     opgCurrencySource_AfterUpdate  ' ** Procedure: Above.
19120     DoEvents

19130     blnFound1 = False: blnFound2 = False
19140     If lngCurrID1 > 0& Or lngCurrID2 > 0& Then
19150       lngCurrs = .cmbCurrencies.ListCount
19160       For lngX = 0& To (lngCurrs - 1&)
19170         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID1 Then
19180           blnFound1 = True
19190         End If
19200         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID2 Then
19210           blnFound2 = True
19220           .cmbCurrencies = lngCurrID2
19230         End If
19240       Next
19250       If blnFound1 = False Then lngCurrID1 = 0&
19260       If blnFound2 = False Then
19270         lngCurrID2 = 0&
19280         .cmbCurrencies = Null
19290       End If
19300     End If

19310     ResetFilter_Set  ' ** Procedure: Below.

19320     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
19330     If lngRecsCur > 0& Then
19340       .frmCurrency_Rate_Sub.SetFocus
19350       If blnFound1 = True Then
19360         .frmCurrency_Rate_Sub.Form.MoveRec 0, lngCurrID1  ' ** Form Procedure: frmCurrency_Rate_Sub.
19370       End If
19380     End If

19390   End With

EXITP:
19400   Exit Sub

ERRH:
19410   Select Case ERR.Number
        Case Else
19420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19430   End Select
19440   Resume EXITP

End Sub

Private Sub opgActive_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19520   intShiftDown = (Shift And acShiftMask) > 0
19530   intAltDown = (Shift And acAltMask) > 0
19540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19570       With Me
19580         intRetVal = 0
19590         If .cmdResetFilter.Enabled = True Then
19600           .cmdResetFilter.SetFocus
19610         Else
19620           .opgFund.SetFocus
19630         End If
19640       End With
19650     End Select
19660   End If

        ' ** Shift keys.
19670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19690       With Me
19700         intRetVal = 0
19710         .frmCurrency_Rate_Sub.SetFocus
19720       End With
19730     End Select
19740   End If

EXITP:
19750   KeyCode = intRetVal
19760   Exit Sub

ERRH:
19770   intRetVal = 0
19780   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
19790   Case Else
19800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19810   End Select
19820   Resume EXITP

End Sub

Private Sub opgActive_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19920   intShiftDown = (Shift And acShiftMask) > 0
19930   intAltDown = (Shift And acAltMask) > 0
19940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19970       With Me
19980         intRetVal = 0
19990         If .cmdResetFilter.Enabled = True Then
20000           .cmdResetFilter.SetFocus
20010         Else
20020           .opgFund.SetFocus
20030         End If
20040       End With
20050     End Select
20060   End If

        ' ** Shift keys.
20070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20090       With Me
20100         intRetVal = 0
20110         .frmCurrency_Rate_Sub.SetFocus
20120       End With
20130     End Select
20140   End If

EXITP:
20150   KeyCode = intRetVal
20160   Exit Sub

ERRH:
20170   intRetVal = 0
20180   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
20190   Case Else
20200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20210   End Select
20220   Resume EXITP

End Sub

Private Sub opgActive_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20320   intShiftDown = (Shift And acShiftMask) > 0
20330   intAltDown = (Shift And acAltMask) > 0
20340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20370       With Me
20380         intRetVal = 0
20390         If .cmdResetFilter.Enabled = True Then
20400           .cmdResetFilter.SetFocus
20410         Else
20420           .opgFund.SetFocus
20430         End If
20440       End With
20450     End Select
20460   End If

        ' ** Shift keys.
20470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20490       With Me
20500         intRetVal = 0
20510         .frmCurrency_Rate_Sub.SetFocus
20520       End With
20530     End Select
20540   End If

EXITP:
20550   KeyCode = intRetVal
20560   Exit Sub

ERRH:
20570   intRetVal = 0
20580   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
20590   Case Else
20600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20610   End Select
20620   Resume EXITP

End Sub

Private Sub opgActive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_MouseMove"

20710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
20720   Exit Sub

ERRH:
20730   Select Case ERR.Number
        Case Else
20740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20750   End Select
20760   Resume EXITP

End Sub

Private Sub cmdResetFilter_GotFocus()

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_GotFocus"

20810   With Me
20820     blnClearFilter_Focus = True
20830     .cmdResetFilter_raised_semifocus_dots_img.Visible = True
20840     .cmdResetFilter_raised_img.Visible = False
20850     .cmdResetFilter_raised_focus_img.Visible = False
20860     .cmdResetFilter_raised_focus_dots_img.Visible = False
20870     .cmdResetFilter_sunken_focus_dots_img.Visible = False
20880     .cmdResetFilter_raised_img_dis.Visible = False
20890   End With

EXITP:
20900   Exit Sub

ERRH:
20910   Select Case ERR.Number
        Case Else
20920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20930   End Select
20940   Resume EXITP

End Sub

Private Sub cmdResetFilter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_MouseDown"

21010   With Me
21020     blnClearFilter_MouseDown = True
21030     .cmdResetFilter_sunken_focus_dots_img.Visible = True
21040     .cmdResetFilter_raised_img.Visible = False
21050     .cmdResetFilter_raised_semifocus_dots_img.Visible = False
21060     .cmdResetFilter_raised_focus_img.Visible = False
21070     .cmdResetFilter_raised_focus_dots_img.Visible = False
21080     .cmdResetFilter_raised_img_dis.Visible = False
21090   End With

EXITP:
21100   Exit Sub

ERRH:
21110   Select Case ERR.Number
        Case Else
21120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21130   End Select
21140   Resume EXITP

End Sub

Private Sub cmdResetFilter_Click()

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_Click"

        Dim strFilter As String

21210   With Me

21220     .opgActive = .opgActive_optInclude.OptionValue
21230     .opgActive_optInclude_lbl.FontBold = True
21240     .opgActive_optExclude_lbl.FontBold = False
21250     .opgActive_optOnly_lbl.FontBold = False

21260     .opgFund = .opgFund_optInclude.OptionValue
21270     .opgFund_optInclude_lbl.FontBold = True
21280     .opgFund_optExclude_lbl.FontBold = False
21290     .opgFund_optOnly_lbl.FontBold = False

21300     .opgMetal = .opgMetal_optInclude.OptionValue
21310     .opgMetal_optInclude_lbl.FontBold = True
21320     .opgMetal_optExclude_lbl.FontBold = False
21330     .opgMetal_optOnly_lbl.FontBold = False

21340     .opgBMU = .opgBMU_optInclude.OptionValue
21350     .opgBMU_optInclude_lbl.FontBold = True
21360     .opgBMU_optExclude_lbl.FontBold = False
21370     .opgBMU_optOnly_lbl.FontBold = False

21380     .opgAlt = .opgAlt_optInclude.OptionValue
21390     .opgAlt_optInclude_lbl.FontBold = True
21400     .opgAlt_optExclude_lbl.FontBold = False
21410     .opgAlt_optOnly_lbl.FontBold = False

21420     .opgUnit = .opgUnit_optInclude.OptionValue
21430     .opgUnit_optInclude_lbl.FontBold = True
21440     .opgUnit_optExclude_lbl.FontBold = False
21450     .opgUnit_optOnly_lbl.FontBold = False

21460     strFilter = vbNullString
21470     .frmCurrency_Rate_Sub.Form.Filter = strFilter
21480     .frmCurrency_Rate_Sub.Form.FilterOn = False
21490     DoEvents

21500     opgCurrencySource_AfterUpdate  ' ** Procedure: Above.
21510     DoEvents
21520     ResetFilter_Set  ' ** Procedure: Below.

21530   End With

EXITP:
21540   Exit Sub

ERRH:
21550   Select Case ERR.Number
        Case Else
21560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21570   End Select
21580   Resume EXITP

End Sub

Private Sub cmdResetFilter_KeyDown(KeyCode As Integer, Shift As Integer)

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21620   intShiftDown = (Shift And acShiftMask) > 0
21630   intAltDown = (Shift And acAltMask) > 0
21640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21660     Select Case intRetVal
          Case vbKeyTab
21670       With Me
21680         intRetVal = 0
21690         .opgFund.SetFocus
21700       End With
21710     End Select
21720   End If

        ' ** Shift keys.
21730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21740     Select Case intRetVal
          Case vbKeyTab
21750       With Me
21760         intRetVal = 0
21770         .opgActive.SetFocus
21780       End With
21790     End Select
21800   End If

EXITP:
21810   KeyCode = intRetVal
21820   Exit Sub

ERRH:
21830   intRetVal = 0
21840   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
21850   Case Else
21860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21870   End Select
21880   Resume EXITP

End Sub

Private Sub cmdResetFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_MouseMove"

21910   With Me
21920     If blnClearFilter_MouseDown = False Then
21930       Select Case blnClearFilter_Focus
            Case True
21940         .cmdResetFilter_raised_focus_dots_img.Visible = True
21950         .cmdResetFilter_raised_focus_img.Visible = False
21960       Case False
21970         .cmdResetFilter_raised_focus_img.Visible = True
21980         .cmdResetFilter_raised_focus_dots_img.Visible = False
21990       End Select
22000       .cmdResetFilter_raised_img.Visible = False
22010       .cmdResetFilter_raised_semifocus_dots_img.Visible = False
22020       .cmdResetFilter_sunken_focus_dots_img.Visible = False
22030       .cmdResetFilter_raised_img_dis.Visible = False
22040     End If
22050   End With

EXITP:
22060   Exit Sub

ERRH:
22070   Select Case ERR.Number
        Case Else
22080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22090   End Select
22100   Resume EXITP

End Sub

Private Sub cmdResetFilter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_MouseUp"

22210   With Me
22220     .cmdResetFilter_raised_focus_dots_img.Visible = True
22230     .cmdResetFilter_raised_img.Visible = False
22240     .cmdResetFilter_raised_semifocus_dots_img.Visible = False
22250     .cmdResetFilter_raised_focus_img.Visible = False
22260     .cmdResetFilter_sunken_focus_dots_img.Visible = False
22270     .cmdResetFilter_raised_img_dis.Visible = False
22280     blnClearFilter_MouseDown = False
22290   End With

EXITP:
22300   Exit Sub

ERRH:
22310   Select Case ERR.Number
        Case Else
22320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22330   End Select
22340   Resume EXITP

End Sub

Private Sub cmdResetFilter_LostFocus()

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_LostFocus"

22410   With Me
22420     .cmdResetFilter_raised_img.Visible = False
22430     .cmdResetFilter_raised_semifocus_dots_img.Visible = False
22440     .cmdResetFilter_raised_focus_img.Visible = False
22450     .cmdResetFilter_raised_focus_dots_img.Visible = False
22460     .cmdResetFilter_sunken_focus_dots_img.Visible = False
22470     .cmdResetFilter_raised_img_dis.Visible = False
22480     blnClearFilter_Focus = False
22490   End With

EXITP:
22500   Exit Sub

ERRH:
22510   Select Case ERR.Number
        Case Else
22520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22530   End Select
22540   Resume EXITP

End Sub

Public Sub opgFund_AfterUpdate()

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFund_AfterUpdate"

        Dim lngCurrID1 As Long, lngCurrID2 As Long, lngCurrs As Long
        Dim strFilter As String
        Dim blnFound1 As Boolean, blnFound2 As Boolean
        Dim lngX As Long

22610   With Me

22620     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
22630     If lngRecsCur > 0& Then
22640       lngCurrID1 = .frmCurrency_Rate_Sub.Form.curr_id
22650     Else
22660       lngCurrID1 = 0&
22670     End If

22680     lngCurrs = .cmbCurrencies.ListCount
22690     If lngCurrs > 0& Then
22700       Select Case IsNull(.cmbCurrencies)
            Case True
22710         lngCurrID2 = 0&
22720       Case False
22730         lngCurrID2 = .cmbCurrencies.Column(CBX_C_CURRID)
22740       End Select
22750     Else
22760       lngCurrID2 = 0&
22770     End If

22780     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
22790       .opgFund_optInclude_lbl.FontBold = True
22800       .opgFund_optExclude_lbl.FontBold = False
22810       .opgFund_optOnly_lbl.FontBold = False
22820       strFilter = vbNullString
22830     Case .opgFund_optExclude.OptionValue
22840       .opgFund_optInclude_lbl.FontBold = False
22850       .opgFund_optExclude_lbl.FontBold = True
22860       .opgFund_optOnly_lbl.FontBold = False
22870       strFilter = "[curr_fund] = false"
22880     Case .opgFund_optOnly.OptionValue
22890       .opgFund_optInclude_lbl.FontBold = False
22900       .opgFund_optExclude_lbl.FontBold = False
22910       .opgFund_optOnly_lbl.FontBold = True
22920       strFilter = "[curr_fund] = True"
            ' ** Flip the other 4 to Exclude.
22930       If .opgBMU <> .opgBMU_optExclude.OptionValue Then
22940         .opgBMU = .opgBMU_optExclude.OptionValue
22950         .opgBMU_optExclude_lbl.FontBold = True
22960         .opgBMU_optInclude_lbl.FontBold = False
22970         .opgBMU_optOnly_lbl.FontBold = False
22980       End If
22990       If .opgMetal <> .opgMetal_optExclude.OptionValue Then
23000         .opgMetal = .opgMetal_optExclude.OptionValue
23010         .opgMetal_optExclude_lbl.FontBold = True
23020         .opgMetal_optInclude_lbl.FontBold = False
23030         .opgMetal_optOnly_lbl.FontBold = False
23040       End If
23050       If .opgAlt <> .opgAlt_optExclude.OptionValue Then
23060         .opgAlt = .opgAlt_optExclude.OptionValue
23070         .opgAlt_optExclude_lbl.FontBold = True
23080         .opgAlt_optInclude_lbl.FontBold = False
23090         .opgAlt_optOnly_lbl.FontBold = False
23100       End If
23110       If .opgUnit <> .opgUnit_optExclude.OptionValue Then
23120         .opgUnit = .opgUnit_optExclude.OptionValue
23130         .opgUnit_optExclude_lbl.FontBold = True
23140         .opgUnit_optInclude_lbl.FontBold = False
23150         .opgUnit_optOnly_lbl.FontBold = False
23160       End If
23170     End Select

23180     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
            ' ** No change.
23190     Case .opgActive_optExclude.OptionValue
23200       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23210       strFilter = strFilter & "[curr_active] = False"
23220     Case .opgActive_optOnly.OptionValue
23230       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23240       strFilter = strFilter & "[curr_active] = True"
23250     End Select

23260     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
            ' ** No change.
23270     Case .opgBMU_optExclude.OptionValue
23280       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23290       strFilter = strFilter & "[curr_bmu] = False"
23300     Case .opgBMU_optOnly.OptionValue
23310       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23320       strFilter = strFilter & "[curr_bmu] = True"
23330     End Select

23340     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
            ' ** No change.
23350     Case .opgMetal_optExclude.OptionValue
23360       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23370       strFilter = strFilter & "[curr_metal] = False"
23380     Case .opgMetal_optOnly.OptionValue
23390       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23400       strFilter = strFilter & "[curr_metal] = True"
23410     End Select

23420     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
            ' ** No change.
23430     Case .opgAlt_optExclude.OptionValue
23440       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23450       strFilter = strFilter & "[curr_alt] = False"
23460     Case .opgAlt_optOnly.OptionValue
23470       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23480       strFilter = strFilter & "[curr_alt] = True"
23490     End Select

23500     Select Case .opgUnit
          Case .opgUnit_optInclude.OptionValue
            ' ** No change.
23510     Case .opgUnit_optExclude.OptionValue
23520       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23530       strFilter = strFilter & "[curr_unit] = False"
23540     Case .opgUnit_optOnly.OptionValue
23550       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23560       strFilter = strFilter & "[curr_unit] = True"
23570     End Select

23580     .frmCurrency_Rate_Sub.Form.Filter = strFilter
23590     If strFilter <> vbNullString Then
23600       .frmCurrency_Rate_Sub.Form.FilterOn = True
23610     Else
23620       .frmCurrency_Rate_Sub.Form.FilterOn = False
23630     End If
23640     DoEvents

23650     opgCurrencySource_AfterUpdate  ' ** Procedure: Above.
23660     DoEvents

23670     blnFound1 = False: blnFound2 = False
23680     If lngCurrID1 > 0& Or lngCurrID2 > 0& Then
23690       lngCurrs = .cmbCurrencies.ListCount
23700       For lngX = 0& To (lngCurrs - 1&)
23710         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID1 Then
23720           blnFound1 = True
23730         End If
23740         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID2 Then
23750           blnFound2 = True
23760           .cmbCurrencies = lngCurrID2
23770         End If
23780       Next
23790       If blnFound1 = False Then lngCurrID1 = 0&
23800       If blnFound2 = False Then
23810         lngCurrID2 = 0&
23820         .cmbCurrencies = Null
23830       End If
23840     End If

23850     ResetFilter_Set  ' ** Procedure: Below.

23860     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
23870     If lngRecsCur > 0& Then
23880       .frmCurrency_Rate_Sub.SetFocus
23890       If blnFound1 = True Then
23900         .frmCurrency_Rate_Sub.Form.MoveRec 0, lngCurrID1  ' ** Form Procedure: frmCurrency_Rate_Sub.
23910       End If
23920     End If

23930   End With

EXITP:
23940   Exit Sub

ERRH:
23950   Select Case ERR.Number
        Case Else
23960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23970   End Select
23980   Resume EXITP

End Sub

Private Sub opgFund_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

24000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFund_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24020   intShiftDown = (Shift And acShiftMask) > 0
24030   intAltDown = (Shift And acAltMask) > 0
24040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24070       With Me
24080         intRetVal = 0
24090         .opgMetal.SetFocus
24100       End With
24110     End Select
24120   End If

        ' ** Shift keys.
24130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24150       With Me
24160         intRetVal = 0
24170         If .cmdResetFilter.Enabled = True Then
24180           .cmdResetFilter.SetFocus
24190         Else
24200           .opgActive.SetFocus
24210         End If
24220       End With
24230     End Select
24240   End If

EXITP:
24250   KeyCode = intRetVal
24260   Exit Sub

ERRH:
24270   intRetVal = 0
24280   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
24290   Case Else
24300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24310   End Select
24320   Resume EXITP

End Sub

Private Sub opgFund_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFund_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24420   intShiftDown = (Shift And acShiftMask) > 0
24430   intAltDown = (Shift And acAltMask) > 0
24440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24470       With Me
24480         intRetVal = 0
24490         .opgMetal.SetFocus
24500       End With
24510     End Select
24520   End If

        ' ** Shift keys.
24530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24550       With Me
24560         intRetVal = 0
24570         If .cmdResetFilter.Enabled = True Then
24580           .cmdResetFilter.SetFocus
24590         Else
24600           .opgActive.SetFocus
24610         End If
24620       End With
24630     End Select
24640   End If

EXITP:
24650   KeyCode = intRetVal
24660   Exit Sub

ERRH:
24670   intRetVal = 0
24680   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
24690   Case Else
24700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24710   End Select
24720   Resume EXITP

End Sub

Private Sub opgFund_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFund_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24820   intShiftDown = (Shift And acShiftMask) > 0
24830   intAltDown = (Shift And acAltMask) > 0
24840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24870       With Me
24880         intRetVal = 0
24890         .opgMetal.SetFocus
24900       End With
24910     End Select
24920   End If

        ' ** Shift keys.
24930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24950       With Me
24960         intRetVal = 0
24970         If .cmdResetFilter.Enabled = True Then
24980           .cmdResetFilter.SetFocus
24990         Else
25000           .opgActive.SetFocus
25010         End If
25020       End With
25030     End Select
25040   End If

EXITP:
25050   KeyCode = intRetVal
25060   Exit Sub

ERRH:
25070   intRetVal = 0
25080   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
25090   Case Else
25100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25110   End Select
25120   Resume EXITP

End Sub

Public Sub opgMetal_AfterUpdate()

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgMetal_AfterUpdate"

        Dim lngCurrID1 As Long, lngCurrID2 As Long, lngCurrs As Long
        Dim strFilter As String
        Dim blnFound1 As Boolean, blnFound2 As Boolean
        Dim lngX As Long

25210   With Me

25220     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
25230     If lngRecsCur > 0& Then
25240       lngCurrID1 = .frmCurrency_Rate_Sub.Form.curr_id
25250     Else
25260       lngCurrID1 = 0&
25270     End If

25280     lngCurrs = .cmbCurrencies.ListCount
25290     If lngCurrs > 0& Then
25300       Select Case IsNull(.cmbCurrencies)
            Case True
25310         lngCurrID2 = 0&
25320       Case False
25330         lngCurrID2 = .cmbCurrencies.Column(CBX_C_CURRID)
25340       End Select
25350     Else
25360       lngCurrID2 = 0&
25370     End If

25380     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
25390       .opgMetal_optInclude_lbl.FontBold = True
25400       .opgMetal_optExclude_lbl.FontBold = False
25410       .opgMetal_optOnly_lbl.FontBold = False
25420       strFilter = vbNullString
25430     Case .opgMetal_optExclude.OptionValue
25440       .opgMetal_optInclude_lbl.FontBold = False
25450       .opgMetal_optExclude_lbl.FontBold = True
25460       .opgMetal_optOnly_lbl.FontBold = False
25470       strFilter = "[curr_metal] = False"
25480     Case .opgMetal_optOnly.OptionValue
25490       .opgMetal_optInclude_lbl.FontBold = False
25500       .opgMetal_optExclude_lbl.FontBold = False
25510       .opgMetal_optOnly_lbl.FontBold = True
25520       strFilter = "[curr_metal] = True"
            ' ** Flip the other 4 to Exclude.
25530       If .opgFund <> .opgFund_optExclude.OptionValue Then
25540         .opgFund = .opgFund_optExclude.OptionValue
25550         .opgFund_optExclude_lbl.FontBold = True
25560         .opgFund_optInclude_lbl.FontBold = False
25570         .opgFund_optOnly_lbl.FontBold = False
25580       End If
25590       If .opgBMU <> .opgBMU_optExclude.OptionValue Then
25600         .opgBMU = .opgBMU_optExclude.OptionValue
25610         .opgBMU_optExclude_lbl.FontBold = True
25620         .opgBMU_optInclude_lbl.FontBold = False
25630         .opgBMU_optOnly_lbl.FontBold = False
25640       End If
25650       If .opgAlt <> .opgAlt_optExclude.OptionValue Then
25660         .opgAlt = .opgAlt_optExclude.OptionValue
25670         .opgAlt_optExclude_lbl.FontBold = True
25680         .opgAlt_optInclude_lbl.FontBold = False
25690         .opgAlt_optOnly_lbl.FontBold = False
25700       End If
25710       If .opgUnit <> .opgUnit_optExclude.OptionValue Then
25720         .opgUnit = .opgUnit_optExclude.OptionValue
25730         .opgUnit_optExclude_lbl.FontBold = True
25740         .opgUnit_optInclude_lbl.FontBold = False
25750         .opgUnit_optOnly_lbl.FontBold = False
25760       End If
25770     End Select

25780     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
            ' ** No change.
25790     Case .opgActive_optExclude.OptionValue
25800       If strFilter <> vbNullString Then strFilter = strFilter & FAND
25810       strFilter = strFilter & "[curr_active] = False"
25820     Case .opgActive_optOnly.OptionValue
25830       If strFilter <> vbNullString Then strFilter = strFilter & FAND
25840       strFilter = strFilter & "[curr_active] = True"
25850     End Select

25860     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
            ' ** No change.
25870     Case .opgFund_optExclude.OptionValue
25880       If strFilter <> vbNullString Then strFilter = strFilter & FAND
25890       strFilter = strFilter & "[curr_fund] = False"
25900     Case .opgFund_optOnly.OptionValue
25910       If strFilter <> vbNullString Then strFilter = strFilter & FAND
25920       strFilter = strFilter & "[curr_fund] = True"
25930     End Select

25940     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
            ' ** No change.
25950     Case .opgBMU_optExclude.OptionValue
25960       If strFilter <> vbNullString Then strFilter = strFilter & FAND
25970       strFilter = strFilter & "[curr_bmu] = False"
25980     Case .opgBMU_optOnly.OptionValue
25990       If strFilter <> vbNullString Then strFilter = strFilter & FAND
26000       strFilter = strFilter & "[curr_bmu] = True"
26010     End Select

26020     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
            ' ** No change.
26030     Case .opgAlt_optExclude.OptionValue
26040       If strFilter <> vbNullString Then strFilter = strFilter & FAND
26050       strFilter = strFilter & "[curr_alt] = False"
26060     Case .opgAlt_optOnly.OptionValue
26070       If strFilter <> vbNullString Then strFilter = strFilter & FAND
26080       strFilter = strFilter & "[curr_alt] = True"
26090     End Select

26100     Select Case .opgUnit
          Case .opgUnit_optInclude.OptionValue
            ' ** No change.
26110     Case .opgUnit_optExclude.OptionValue
26120       If strFilter <> vbNullString Then strFilter = strFilter & FAND
26130       strFilter = strFilter & "[curr_unit] = False"
26140     Case .opgUnit_optOnly.OptionValue
26150       If strFilter <> vbNullString Then strFilter = strFilter & FAND
26160       strFilter = strFilter & "[curr_unit] = True"
26170     End Select

26180     .frmCurrency_Rate_Sub.Form.Filter = strFilter
26190     If strFilter <> vbNullString Then
26200       .frmCurrency_Rate_Sub.Form.FilterOn = True
26210     Else
26220       .frmCurrency_Rate_Sub.Form.FilterOn = False
26230     End If
26240     DoEvents

26250     opgCurrencySource_AfterUpdate  ' ** Procedure: Above.
26260     DoEvents

26270     blnFound1 = False: blnFound2 = False
26280     If lngCurrID1 > 0& Or lngCurrID2 > 0& Then
26290       lngCurrs = .cmbCurrencies.ListCount
26300       For lngX = 0& To (lngCurrs - 1&)
26310         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID1 Then
26320           blnFound1 = True
26330         End If
26340         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID2 Then
26350           blnFound2 = True
26360           .cmbCurrencies = lngCurrID2
26370         End If
26380       Next
26390       If blnFound1 = False Then lngCurrID1 = 0&
26400       If blnFound2 = False Then
26410         lngCurrID2 = 0&
26420         .cmbCurrencies = Null
26430       End If
26440     End If

26450     ResetFilter_Set  ' ** Procedure: Below.

26460     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
26470     If lngRecsCur > 0& Then
26480       .frmCurrency_Rate_Sub.SetFocus
26490       If blnFound1 = True Then
26500         .frmCurrency_Rate_Sub.Form.MoveRec 0, lngCurrID1  ' ** Form Procedure: frmCurrency_Rate_Sub.
26510       End If
26520     End If

26530   End With

EXITP:
26540   Exit Sub

ERRH:
26550   Select Case ERR.Number
        Case Else
26560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26570   End Select
26580   Resume EXITP

End Sub

Private Sub opgMetal_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgMetal_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26620   intShiftDown = (Shift And acShiftMask) > 0
26630   intAltDown = (Shift And acAltMask) > 0
26640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26670       With Me
26680         intRetVal = 0
26690         .opgBMU.SetFocus
26700       End With
26710     End Select
26720   End If

        ' ** Shift keys.
26730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26750       With Me
26760         intRetVal = 0
26770         .opgFund.SetFocus
26780       End With
26790     End Select
26800   End If

EXITP:
26810   KeyCode = intRetVal
26820   Exit Sub

ERRH:
26830   intRetVal = 0
26840   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
26850   Case Else
26860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26870   End Select
26880   Resume EXITP

End Sub

Private Sub opgMetal_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgMetal_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26920   intShiftDown = (Shift And acShiftMask) > 0
26930   intAltDown = (Shift And acAltMask) > 0
26940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26970       With Me
26980         intRetVal = 0
26990         .opgBMU.SetFocus
27000       End With
27010     End Select
27020   End If

        ' ** Shift keys.
27030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27050       With Me
27060         intRetVal = 0
27070         .opgFund.SetFocus
27080       End With
27090     End Select
27100   End If

EXITP:
27110   KeyCode = intRetVal
27120   Exit Sub

ERRH:
27130   intRetVal = 0
27140   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
27150   Case Else
27160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27170   End Select
27180   Resume EXITP

End Sub

Private Sub opgMetal_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgMetal_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27220   intShiftDown = (Shift And acShiftMask) > 0
27230   intAltDown = (Shift And acAltMask) > 0
27240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27270       With Me
27280         intRetVal = 0
27290         .opgBMU.SetFocus
27300       End With
27310     End Select
27320   End If

        ' ** Shift keys.
27330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27350       With Me
27360         intRetVal = 0
27370         .opgFund.SetFocus
27380       End With
27390     End Select
27400   End If

EXITP:
27410   KeyCode = intRetVal
27420   Exit Sub

ERRH:
27430   intRetVal = 0
27440   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
27450   Case Else
27460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27470   End Select
27480   Resume EXITP

End Sub

Public Sub opgBMU_AfterUpdate()

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgBMU_AfterUpdate"

        Dim lngCurrID1 As Long, lngCurrID2 As Long, lngCurrs As Long
        Dim strFilter As String
        Dim blnFound1 As Boolean, blnFound2 As Boolean
        Dim lngX As Long

27510   With Me

27520     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
27530     If lngRecsCur > 0& Then
27540       lngCurrID1 = .frmCurrency_Rate_Sub.Form.curr_id
27550     Else
27560       lngCurrID1 = 0&
27570     End If

27580     lngCurrs = .cmbCurrencies.ListCount
27590     If lngCurrs > 0& Then
27600       Select Case IsNull(.cmbCurrencies)
            Case True
27610         lngCurrID2 = 0&
27620       Case False
27630         lngCurrID2 = .cmbCurrencies.Column(CBX_C_CURRID)
27640       End Select
27650     Else
27660       lngCurrID2 = 0&
27670     End If

27680     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
27690       .opgBMU_optInclude_lbl.FontBold = True
27700       .opgBMU_optExclude_lbl.FontBold = False
27710       .opgBMU_optOnly_lbl.FontBold = False
27720       strFilter = vbNullString
27730     Case .opgBMU_optExclude.OptionValue
27740       .opgBMU_optInclude_lbl.FontBold = False
27750       .opgBMU_optExclude_lbl.FontBold = True
27760       .opgBMU_optOnly_lbl.FontBold = False
27770       strFilter = "[curr_bmu] = False"
27780     Case .opgBMU_optOnly.OptionValue
27790       .opgBMU_optInclude_lbl.FontBold = False
27800       .opgBMU_optExclude_lbl.FontBold = False
27810       .opgBMU_optOnly_lbl.FontBold = True
27820       strFilter = "[curr_bmu] = True"
            ' ** Flip the other 4 to Exclude.
27830       If .opgFund <> .opgFund_optExclude.OptionValue Then
27840         .opgFund = .opgFund_optExclude.OptionValue
27850         .opgFund_optExclude_lbl.FontBold = True
27860         .opgFund_optInclude_lbl.FontBold = False
27870         .opgFund_optOnly_lbl.FontBold = False
27880       End If
27890       If .opgMetal <> .opgMetal_optExclude.OptionValue Then
27900         .opgMetal = .opgMetal_optExclude.OptionValue
27910         .opgMetal_optExclude_lbl.FontBold = True
27920         .opgMetal_optInclude_lbl.FontBold = False
27930         .opgMetal_optOnly_lbl.FontBold = False
27940       End If
27950       If .opgAlt <> .opgAlt_optExclude.OptionValue Then
27960         .opgAlt = .opgAlt_optExclude.OptionValue
27970         .opgAlt_optExclude_lbl.FontBold = True
27980         .opgAlt_optInclude_lbl.FontBold = False
27990         .opgAlt_optOnly_lbl.FontBold = False
28000       End If
28010       If .opgUnit <> .opgUnit_optExclude.OptionValue Then
28020         .opgUnit = .opgUnit_optExclude.OptionValue
28030         .opgUnit_optExclude_lbl.FontBold = True
28040         .opgUnit_optInclude_lbl.FontBold = False
28050         .opgUnit_optOnly_lbl.FontBold = False
28060       End If
28070     End Select

28080     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
            ' ** No change.
28090     Case .opgActive_optExclude.OptionValue
28100       If strFilter <> vbNullString Then strFilter = strFilter & FAND
28110       strFilter = strFilter & "[curr_active] = False"
28120     Case .opgActive_optOnly.OptionValue
28130       If strFilter <> vbNullString Then strFilter = strFilter & FAND
28140       strFilter = strFilter & "[curr_active] = True"
28150     End Select

28160     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
            ' ** No change.
28170     Case .opgFund_optExclude.OptionValue
28180       If strFilter <> vbNullString Then strFilter = strFilter & FAND
28190       strFilter = strFilter & "[curr_fund] = False"
28200     Case .opgFund_optOnly.OptionValue
28210       If strFilter <> vbNullString Then strFilter = strFilter & FAND
28220       strFilter = strFilter & "[curr_fund] = True"
28230     End Select

28240     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
            ' ** No change.
28250     Case .opgMetal_optExclude.OptionValue
28260       If strFilter <> vbNullString Then strFilter = strFilter & FAND
28270       strFilter = strFilter & "[curr_metal] = False"
28280     Case .opgMetal_optOnly.OptionValue
28290       If strFilter <> vbNullString Then strFilter = strFilter & FAND
28300       strFilter = strFilter & "[curr_metal] = True"
28310     End Select

28320     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
            ' ** No change.
28330     Case .opgAlt_optExclude.OptionValue
28340       If strFilter <> vbNullString Then strFilter = strFilter & FAND
28350       strFilter = strFilter & "[curr_alt] = False"
28360     Case .opgAlt_optOnly.OptionValue
28370       If strFilter <> vbNullString Then strFilter = strFilter & FAND
28380       strFilter = strFilter & "[curr_alt] = True"
28390     End Select

28400     Select Case .opgUnit
          Case .opgUnit_optInclude.OptionValue
            ' ** No change.
28410     Case .opgUnit_optExclude.OptionValue
28420       If strFilter <> vbNullString Then strFilter = strFilter & FAND
28430       strFilter = strFilter & "[curr_unit] = False"
28440     Case .opgUnit_optOnly.OptionValue
28450       If strFilter <> vbNullString Then strFilter = strFilter & FAND
28460       strFilter = strFilter & "[curr_unit] = True"
28470     End Select

28480     .frmCurrency_Rate_Sub.Form.Filter = strFilter
28490     If strFilter <> vbNullString Then
28500       .frmCurrency_Rate_Sub.Form.FilterOn = True
28510     Else
28520       .frmCurrency_Rate_Sub.Form.FilterOn = False
28530     End If
28540     DoEvents

28550     opgCurrencySource_AfterUpdate  ' ** Procedure: Above.
28560     DoEvents

28570     blnFound1 = False: blnFound2 = False
28580     If lngCurrID1 > 0& Or lngCurrID2 > 0& Then
28590       lngCurrs = .cmbCurrencies.ListCount
28600       For lngX = 0& To (lngCurrs - 1&)
28610         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID1 Then
28620           blnFound1 = True
28630         End If
28640         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID2 Then
28650           blnFound2 = True
28660           .cmbCurrencies = lngCurrID2
28670         End If
28680       Next
28690       If blnFound1 = False Then lngCurrID1 = 0&
28700       If blnFound2 = False Then
28710         lngCurrID2 = 0&
28720         .cmbCurrencies = Null
28730       End If
28740     End If

28750     ResetFilter_Set  ' ** Procedure: Below.

28760     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
28770     If lngRecsCur > 0& Then
28780       .frmCurrency_Rate_Sub.SetFocus
28790       If blnFound1 = True Then
28800         .frmCurrency_Rate_Sub.Form.MoveRec 0, lngCurrID1  ' ** Form Procedure: frmCurrency_Rate_Sub.
28810       End If
28820     End If

28830   End With

EXITP:
28840   Exit Sub

ERRH:
28850   Select Case ERR.Number
        Case Else
28860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28870   End Select
28880   Resume EXITP

End Sub

Private Sub opgBMU_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgBMU_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28920   intShiftDown = (Shift And acShiftMask) > 0
28930   intAltDown = (Shift And acAltMask) > 0
28940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28970       With Me
28980         intRetVal = 0
28990         .opgAlt.SetFocus
29000       End With
29010     End Select
29020   End If

        ' ** Shift keys.
29030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29050       With Me
29060         intRetVal = 0
29070         .opgMetal.SetFocus
29080       End With
29090     End Select
29100   End If

EXITP:
29110   KeyCode = intRetVal
29120   Exit Sub

ERRH:
29130   intRetVal = 0
29140   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
29150   Case Else
29160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29170   End Select
29180   Resume EXITP

End Sub

Private Sub opgBMU_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

29200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgBMU_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29220   intShiftDown = (Shift And acShiftMask) > 0
29230   intAltDown = (Shift And acAltMask) > 0
29240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29270       With Me
29280         intRetVal = 0
29290         .opgAlt.SetFocus
29300       End With
29310     End Select
29320   End If

        ' ** Shift keys.
29330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29350       With Me
29360         intRetVal = 0
29370         .opgMetal.SetFocus
29380       End With
29390     End Select
29400   End If

EXITP:
29410   KeyCode = intRetVal
29420   Exit Sub

ERRH:
29430   intRetVal = 0
29440   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
29450   Case Else
29460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29470   End Select
29480   Resume EXITP

End Sub

Private Sub opgBMU_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgBMU_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29520   intShiftDown = (Shift And acShiftMask) > 0
29530   intAltDown = (Shift And acAltMask) > 0
29540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29570       With Me
29580         intRetVal = 0
29590         .opgAlt.SetFocus
29600       End With
29610     End Select
29620   End If

        ' ** Shift keys.
29630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29650       With Me
29660         intRetVal = 0
29670         .opgMetal.SetFocus
29680       End With
29690     End Select
29700   End If

EXITP:
29710   KeyCode = intRetVal
29720   Exit Sub

ERRH:
29730   intRetVal = 0
29740   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
29750   Case Else
29760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29770   End Select
29780   Resume EXITP

End Sub

Public Sub opgAlt_AfterUpdate()

29800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAlt_AfterUpdate"

        Dim lngCurrID1 As Long, lngCurrID2 As Long, lngCurrs As Long
        Dim strFilter As String
        Dim blnFound1 As Boolean, blnFound2 As Boolean
        Dim lngX As Long

29810   With Me

29820     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
29830     If lngRecsCur > 0& Then
29840       lngCurrID1 = .frmCurrency_Rate_Sub.Form.curr_id
29850     Else
29860       lngCurrID1 = 0&
29870     End If

29880     lngCurrs = .cmbCurrencies.ListCount
29890     If lngCurrs > 0& Then
29900       Select Case IsNull(.cmbCurrencies)
            Case True
29910         lngCurrID2 = 0&
29920       Case False
29930         lngCurrID2 = .cmbCurrencies.Column(CBX_C_CURRID)
29940       End Select
29950     Else
29960       lngCurrID2 = 0&
29970     End If

29980     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
29990       .opgAlt_optInclude_lbl.FontBold = True
30000       .opgAlt_optExclude_lbl.FontBold = False
30010       .opgAlt_optOnly_lbl.FontBold = False
30020       strFilter = vbNullString
30030     Case .opgAlt_optExclude.OptionValue
30040       .opgAlt_optInclude_lbl.FontBold = False
30050       .opgAlt_optExclude_lbl.FontBold = True
30060       .opgAlt_optOnly_lbl.FontBold = False
30070       strFilter = "[curr_alt] = False"
30080     Case .opgAlt_optOnly.OptionValue
30090       .opgAlt_optInclude_lbl.FontBold = False
30100       .opgAlt_optExclude_lbl.FontBold = False
30110       .opgAlt_optOnly_lbl.FontBold = True
30120       strFilter = "[curr_alt] = True"
            ' ** Flip the other 4 to Exclude.
30130       If .opgFund <> .opgFund_optExclude.OptionValue Then
30140         .opgFund = .opgFund_optExclude.OptionValue
30150         .opgFund_optExclude_lbl.FontBold = True
30160         .opgFund_optInclude_lbl.FontBold = False
30170         .opgFund_optOnly_lbl.FontBold = False
30180       End If
30190       If .opgBMU <> .opgBMU_optExclude.OptionValue Then
30200         .opgBMU = .opgBMU_optExclude.OptionValue
30210         .opgBMU_optExclude_lbl.FontBold = True
30220         .opgBMU_optInclude_lbl.FontBold = False
30230         .opgBMU_optOnly_lbl.FontBold = False
30240       End If
30250       If .opgMetal <> .opgMetal_optExclude.OptionValue Then
30260         .opgMetal = .opgMetal_optExclude.OptionValue
30270         .opgMetal_optExclude_lbl.FontBold = True
30280         .opgMetal_optInclude_lbl.FontBold = False
30290         .opgMetal_optOnly_lbl.FontBold = False
30300       End If
30310       If .opgUnit <> .opgUnit_optExclude.OptionValue Then
30320         .opgUnit = .opgUnit_optExclude.OptionValue
30330         .opgUnit_optExclude_lbl.FontBold = True
30340         .opgUnit_optInclude_lbl.FontBold = False
30350         .opgUnit_optOnly_lbl.FontBold = False
30360       End If
30370     End Select

30380     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
            ' ** No change.
30390     Case .opgActive_optExclude.OptionValue
30400       If strFilter <> vbNullString Then strFilter = strFilter & FAND
30410       strFilter = strFilter & "[curr_active] = False"
30420     Case .opgActive_optOnly.OptionValue
30430       If strFilter <> vbNullString Then strFilter = strFilter & FAND
30440       strFilter = strFilter & "[curr_active] = True"
30450     End Select

30460     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
            ' ** No change.
30470     Case .opgFund_optExclude.OptionValue
30480       If strFilter <> vbNullString Then strFilter = strFilter & FAND
30490       strFilter = strFilter & "[curr_fund] = False"
30500     Case .opgFund_optOnly.OptionValue
30510       If strFilter <> vbNullString Then strFilter = strFilter & FAND
30520       strFilter = strFilter & "[curr_fund] = True"
30530     End Select

30540     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
            ' ** No change.
30550     Case .opgBMU_optExclude.OptionValue
30560       If strFilter <> vbNullString Then strFilter = strFilter & FAND
30570       strFilter = strFilter & "[curr_bmu] = False"
30580     Case .opgBMU_optOnly.OptionValue
30590       If strFilter <> vbNullString Then strFilter = strFilter & FAND
30600       strFilter = strFilter & "[curr_bmu] = True"
30610     End Select

30620     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
            ' ** No change.
30630     Case .opgMetal_optExclude.OptionValue
30640       If strFilter <> vbNullString Then strFilter = strFilter & FAND
30650       strFilter = strFilter & "[curr_metal] = False"
30660     Case .opgMetal_optOnly.OptionValue
30670       If strFilter <> vbNullString Then strFilter = strFilter & FAND
30680       strFilter = strFilter & "[curr_metal] = True"
30690     End Select

30700     Select Case .opgUnit
          Case .opgUnit_optInclude.OptionValue
            ' ** No change.
30710     Case .opgUnit_optExclude.OptionValue
30720       If strFilter <> vbNullString Then strFilter = strFilter & FAND
30730       strFilter = strFilter & "[curr_unit] = False"
30740     Case .opgUnit_optOnly.OptionValue
30750       If strFilter <> vbNullString Then strFilter = strFilter & FAND
30760       strFilter = strFilter & "[curr_unit] = True"
30770     End Select

30780     .frmCurrency_Rate_Sub.Form.Filter = strFilter
30790     If strFilter <> vbNullString Then
30800       .frmCurrency_Rate_Sub.Form.FilterOn = True
30810     Else
30820       .frmCurrency_Rate_Sub.Form.FilterOn = False
30830     End If
30840     DoEvents

30850     opgCurrencySource_AfterUpdate  ' ** Procedure: Above.
30860     DoEvents

30870     blnFound1 = False: blnFound2 = False
30880     If lngCurrID1 > 0& Or lngCurrID2 > 0& Then
30890       lngCurrs = .cmbCurrencies.ListCount
30900       For lngX = 0& To (lngCurrs - 1&)
30910         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID1 Then
30920           blnFound1 = True
30930         End If
30940         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID2 Then
30950           blnFound2 = True
30960           .cmbCurrencies = lngCurrID2
30970         End If
30980       Next
30990       If blnFound1 = False Then lngCurrID1 = 0&
31000       If blnFound2 = False Then
31010         lngCurrID2 = 0&
31020         .cmbCurrencies = Null
31030       End If
31040     End If

31050     ResetFilter_Set  ' ** Procedure: Below.

31060     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
31070     If lngRecsCur > 0& Then
31080       .frmCurrency_Rate_Sub.SetFocus
31090       If blnFound1 = True Then
31100         .frmCurrency_Rate_Sub.Form.MoveRec 0, lngCurrID1  ' ** Form Procedure: frmCurrency_Rate_Sub.
31110       End If
31120     End If

31130   End With

EXITP:
31140   Exit Sub

ERRH:
31150   Select Case ERR.Number
        Case Else
31160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31170   End Select
31180   Resume EXITP

End Sub

Private Sub opgAlt_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAlt_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
31220   intShiftDown = (Shift And acShiftMask) > 0
31230   intAltDown = (Shift And acAltMask) > 0
31240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31270       With Me
31280         intRetVal = 0
31290         .opgUnit.SetFocus
31300       End With
31310     End Select
31320   End If

        ' ** Shift keys.
31330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31350       With Me
31360         intRetVal = 0
31370         .opgBMU.SetFocus
31380       End With
31390     End Select
31400   End If

EXITP:
31410   KeyCode = intRetVal
31420   Exit Sub

ERRH:
31430   intRetVal = 0
31440   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
31450   Case Else
31460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31470   End Select
31480   Resume EXITP

End Sub

Private Sub opgAlt_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

31500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAlt_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
31520   intShiftDown = (Shift And acShiftMask) > 0
31530   intAltDown = (Shift And acAltMask) > 0
31540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31570       With Me
31580         intRetVal = 0
31590         .opgUnit.SetFocus
31600       End With
31610     End Select
31620   End If

        ' ** Shift keys.
31630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31650       With Me
31660         intRetVal = 0
31670         .opgBMU.SetFocus
31680       End With
31690     End Select
31700   End If

EXITP:
31710   KeyCode = intRetVal
31720   Exit Sub

ERRH:
31730   intRetVal = 0
31740   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
31750   Case Else
31760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31770   End Select
31780   Resume EXITP

End Sub

Private Sub opgAlt_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

31800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAlt_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
31820   intShiftDown = (Shift And acShiftMask) > 0
31830   intAltDown = (Shift And acAltMask) > 0
31840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31870       With Me
31880         intRetVal = 0
31890         .opgUnit.SetFocus
31900       End With
31910     End Select
31920   End If

        ' ** Shift keys.
31930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31950       With Me
31960         intRetVal = 0
31970         .opgBMU.SetFocus
31980       End With
31990     End Select
32000   End If

EXITP:
32010   KeyCode = intRetVal
32020   Exit Sub

ERRH:
32030   intRetVal = 0
32040   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
32050   Case Else
32060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32070   End Select
32080   Resume EXITP

End Sub

Public Sub opgUnit_AfterUpdate()

32100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnit_AfterUpdate"

        Dim lngCurrID1 As Long, lngCurrID2 As Long, lngCurrs As Long
        Dim strFilter As String
        Dim blnFound1 As Boolean, blnFound2 As Boolean
        Dim lngX As Long

32110   With Me

32120     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
32130     If lngRecsCur > 0& Then
32140       lngCurrID1 = .frmCurrency_Rate_Sub.Form.curr_id
32150     Else
32160       lngCurrID1 = 0&
32170     End If

32180     lngCurrs = .cmbCurrencies.ListCount
32190     If lngCurrs > 0& Then
32200       Select Case IsNull(.cmbCurrencies)
            Case True
32210         lngCurrID2 = 0&
32220       Case False
32230         lngCurrID2 = .cmbCurrencies.Column(CBX_C_CURRID)
32240       End Select
32250     Else
32260       lngCurrID2 = 0&
32270     End If

32280     Select Case .opgUnit
          Case .opgUnit_optInclude.OptionValue
32290       .opgUnit_optInclude_lbl.FontBold = True
32300       .opgUnit_optExclude_lbl.FontBold = False
32310       .opgUnit_optOnly_lbl.FontBold = False
32320       strFilter = vbNullString
32330     Case .opgUnit_optExclude.OptionValue
32340       .opgUnit_optInclude_lbl.FontBold = False
32350       .opgUnit_optExclude_lbl.FontBold = True
32360       .opgUnit_optOnly_lbl.FontBold = False
32370       strFilter = "[curr_unit] = False"
32380     Case .opgUnit_optOnly.OptionValue
32390       .opgUnit_optInclude_lbl.FontBold = False
32400       .opgUnit_optExclude_lbl.FontBold = False
32410       .opgUnit_optOnly_lbl.FontBold = True
32420       strFilter = "[curr_unit] = True"
            ' ** Flip the other 4 to Exclude.
32430       If .opgFund <> .opgFund_optExclude.OptionValue Then
32440         .opgFund = .opgFund_optExclude.OptionValue
32450         .opgFund_optExclude_lbl.FontBold = True
32460         .opgFund_optInclude_lbl.FontBold = False
32470         .opgFund_optOnly_lbl.FontBold = False
32480       End If
32490       If .opgBMU <> .opgBMU_optExclude.OptionValue Then
32500         .opgBMU = .opgBMU_optExclude.OptionValue
32510         .opgBMU_optExclude_lbl.FontBold = True
32520         .opgBMU_optInclude_lbl.FontBold = False
32530         .opgBMU_optOnly_lbl.FontBold = False
32540       End If
32550       If .opgMetal <> .opgMetal_optExclude.OptionValue Then
32560         .opgMetal = .opgMetal_optExclude.OptionValue
32570         .opgMetal_optExclude_lbl.FontBold = True
32580         .opgMetal_optInclude_lbl.FontBold = False
32590         .opgMetal_optOnly_lbl.FontBold = False
32600       End If
32610       If .opgAlt <> .opgAlt_optExclude.OptionValue Then
32620         .opgAlt = .opgAlt_optExclude.OptionValue
32630         .opgAlt_optExclude_lbl.FontBold = True
32640         .opgAlt_optInclude_lbl.FontBold = False
32650         .opgAlt_optOnly_lbl.FontBold = False
32660       End If
32670     End Select

32680     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
            ' ** No change.
32690     Case .opgActive_optExclude.OptionValue
32700       If strFilter <> vbNullString Then strFilter = strFilter & FAND
32710       strFilter = strFilter & "[curr_active] = False"
32720     Case .opgActive_optOnly.OptionValue
32730       If strFilter <> vbNullString Then strFilter = strFilter & FAND
32740       strFilter = strFilter & "[curr_active] = True"
32750     End Select

32760     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
            ' ** No change.
32770     Case .opgFund_optExclude.OptionValue
32780       If strFilter <> vbNullString Then strFilter = strFilter & FAND
32790       strFilter = strFilter & "[curr_fund] = False"
32800     Case .opgFund_optOnly.OptionValue
32810       If strFilter <> vbNullString Then strFilter = strFilter & FAND
32820       strFilter = strFilter & "[curr_fund] = True"
32830     End Select

32840     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
            ' ** No change.
32850     Case .opgBMU_optExclude.OptionValue
32860       If strFilter <> vbNullString Then strFilter = strFilter & FAND
32870       strFilter = strFilter & "[curr_bmu] = False"
32880     Case .opgBMU_optOnly.OptionValue
32890       If strFilter <> vbNullString Then strFilter = strFilter & FAND
32900       strFilter = strFilter & "[curr_bmu] = True"
32910     End Select

32920     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
            ' ** No change.
32930     Case .opgMetal_optExclude.OptionValue
32940       If strFilter <> vbNullString Then strFilter = strFilter & FAND
32950       strFilter = strFilter & "[curr_metal] = False"
32960     Case .opgMetal_optOnly.OptionValue
32970       If strFilter <> vbNullString Then strFilter = strFilter & FAND
32980       strFilter = strFilter & "[curr_metal] = True"
32990     End Select

33000     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
            ' ** No change.
33010     Case .opgAlt_optExclude.OptionValue
33020       If strFilter <> vbNullString Then strFilter = strFilter & FAND
33030       strFilter = strFilter & "[curr_alt] = False"
33040     Case .opgAlt_optOnly.OptionValue
33050       If strFilter <> vbNullString Then strFilter = strFilter & FAND
33060       strFilter = strFilter & "[curr_alt] = True"
33070     End Select

33080     .frmCurrency_Rate_Sub.Form.Filter = strFilter
33090     If strFilter <> vbNullString Then
33100       .frmCurrency_Rate_Sub.Form.FilterOn = True
33110     Else
33120       .frmCurrency_Rate_Sub.Form.FilterOn = False
33130     End If
33140     DoEvents

33150     opgCurrencySource_AfterUpdate  ' ** Procedure: Above.
33160     DoEvents

33170     blnFound1 = False: blnFound2 = False
33180     If lngCurrID1 > 0& Or lngCurrID2 > 0& Then
33190       lngCurrs = .cmbCurrencies.ListCount
33200       For lngX = 0& To (lngCurrs - 1&)
33210         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID1 Then
33220           blnFound1 = True
33230         End If
33240         If .cmbCurrencies.Column(CBX_C_CURRID, lngX) = lngCurrID2 Then
33250           blnFound2 = True
33260           .cmbCurrencies = lngCurrID2
33270         End If
33280       Next
33290       If blnFound1 = False Then lngCurrID1 = 0&
33300       If blnFound2 = False Then
33310         lngCurrID2 = 0&
33320         .cmbCurrencies = Null
33330       End If
33340     End If

33350     ResetFilter_Set  ' ** Procedure: Below.

33360     lngRecsCur = .frmCurrency_Rate_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Rate_Sub.
33370     If lngRecsCur > 0& Then
33380       .frmCurrency_Rate_Sub.SetFocus
33390       If blnFound1 = True Then
33400         .frmCurrency_Rate_Sub.Form.MoveRec 0, lngCurrID1  ' ** Form Procedure: frmCurrency_Rate_Sub.
33410       End If
33420     End If

33430   End With

EXITP:
33440   Exit Sub

ERRH:
33450   Select Case ERR.Number
        Case Else
33460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
33470   End Select
33480   Resume EXITP

End Sub

Private Sub opgUnit_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

33500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnit_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

33510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
33520   intShiftDown = (Shift And acShiftMask) > 0
33530   intAltDown = (Shift And acAltMask) > 0
33540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
33550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
33560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33570       With Me
33580         intRetVal = 0
33590         .cmdCountries.SetFocus
33600       End With
33610     End Select
33620   End If

        ' ** Shift keys.
33630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33650       With Me
33660         intRetVal = 0
33670         .opgAlt.SetFocus
33680       End With
33690     End Select
33700   End If

EXITP:
33710   KeyCode = intRetVal
33720   Exit Sub

ERRH:
33730   intRetVal = 0
33740   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
33750   Case Else
33760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
33770   End Select
33780   Resume EXITP

End Sub

Private Sub opgUnit_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

33800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnit_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

33810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
33820   intShiftDown = (Shift And acShiftMask) > 0
33830   intAltDown = (Shift And acAltMask) > 0
33840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
33850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
33860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33870       With Me
33880         intRetVal = 0
33890         .cmdCountries.SetFocus
33900       End With
33910     End Select
33920   End If

        ' ** Shift keys.
33930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33950       With Me
33960         intRetVal = 0
33970         .opgAlt.SetFocus
33980       End With
33990     End Select
34000   End If

EXITP:
34010   KeyCode = intRetVal
34020   Exit Sub

ERRH:
34030   intRetVal = 0
34040   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
34050   Case Else
34060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34070   End Select
34080   Resume EXITP

End Sub

Private Sub opgUnit_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

34100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnit_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

34110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
34120   intShiftDown = (Shift And acShiftMask) > 0
34130   intAltDown = (Shift And acAltMask) > 0
34140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
34150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
34170       With Me
34180         intRetVal = 0
34190         .cmdCountries.SetFocus
34200       End With
34210     End Select
34220   End If

        ' ** Shift keys.
34230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
34250       With Me
34260         intRetVal = 0
34270         .opgAlt.SetFocus
34280       End With
34290     End Select
34300   End If

EXITP:
34310   KeyCode = intRetVal
34320   Exit Sub

ERRH:
34330   intRetVal = 0
34340   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
34350   Case Else
34360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34370   End Select
34380   Resume EXITP

End Sub

Private Sub country_cnt_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

34400 On Error GoTo ERRH

        Const THIS_PROC As String = "country_cnt_box_MouseMove"

34410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
34420   Exit Sub

ERRH:
34430   Select Case ERR.Number
        Case Else
34440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34450   End Select
34460   Resume EXITP

End Sub

Private Sub country_cnt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

34500 On Error GoTo ERRH

        Const THIS_PROC As String = "country_cnt_MouseMove"

34510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
34520   Exit Sub

ERRH:
34530   Select Case ERR.Number
        Case Else
34540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34550   End Select
34560   Resume EXITP

End Sub

Private Sub cmdCountries_GotFocus()

34600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_GotFocus"

34610   With Me
34620     blnCountries_Focus = True
34630     .cmdCountries_raised_semifocus_dots_img.Visible = True
34640     .cmdCountries_raised_img.Visible = False
34650     .cmdCountries_raised_focus_img.Visible = False
34660     .cmdCountries_raised_focus_dots_img.Visible = False
34670     .cmdCountries_sunken_focus_dots_img.Visible = False
34680     .cmdCountries_raised_img_dis.Visible = False
34690   End With

EXITP:
34700   Exit Sub

ERRH:
34710   Select Case ERR.Number
        Case Else
34720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34730   End Select
34740   Resume EXITP

End Sub

Private Sub cmdCountries_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_MouseDown"

34810   With Me
34820     blnCountries_MouseDown = True
34830     .cmdCountries_sunken_focus_dots_img.Visible = True
34840     .cmdCountries_raised_img.Visible = False
34850     .cmdCountries_raised_semifocus_dots_img.Visible = False
34860     .cmdCountries_raised_focus_img.Visible = False
34870     .cmdCountries_raised_focus_dots_img.Visible = False
34880     .cmdCountries_raised_img_dis.Visible = False
34890   End With

EXITP:
34900   Exit Sub

ERRH:
34910   Select Case ERR.Number
        Case Else
34920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34930   End Select
34940   Resume EXITP

End Sub

Private Sub cmdCountries_Click()

35000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_Click"

        Dim strDocName As String

35010   strDocName = "frmCurrency_Country"
35020   Select Case IsLoaded(strDocName, acForm)
        Case True
35030     DoCmd.Close acForm, strDocName
35040   Case False
35050     gblnSetFocus = True
35060     DoCmd.OpenForm strDocName, , , , , , THIS_NAME
35070   End Select

EXITP:
35080   Exit Sub

ERRH:
35090   Select Case ERR.Number
        Case Else
35100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
35110   End Select
35120   Resume EXITP

End Sub

Private Sub cmdCountries_KeyDown(KeyCode As Integer, Shift As Integer)

35200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

35210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
35220   intShiftDown = (Shift And acShiftMask) > 0
35230   intAltDown = (Shift And acAltMask) > 0
35240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
35250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
35260     Select Case intRetVal
          Case vbKeyTab
35270       With Me
35280         intRetVal = 0
35290         If .cmdPreview.Enabled = True Then
35300           .cmdPreview.SetFocus
35310         ElseIf .cmdEdit.Enabled = True Then
35320           .cmdEdit.SetFocus
35330         ElseIf .cmdUpdate.Enabled = True Then
35340           .cmdUpdate.SetFocus
35350         Else
35360           .cmdClose.SetFocus
35370         End If
35380       End With
35390     End Select
35400   End If

        ' ** Shift keys.
35410   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35420     Select Case intRetVal
          Case vbKeyTab
35430       With Me
35440         intRetVal = 0
35450         .opgUnit.SetFocus
35460       End With
35470     End Select
35480   End If

        ' ** Ctrl keys.
35490   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35500     Select Case intRetVal
          Case vbKeyTab
35510       With Me
35520         intRetVal = 0
35530         .chkIncludeNotes.SetFocus
35540       End With
35550     End Select
35560   End If

EXITP:
35570   KeyCode = intRetVal
35580   Exit Sub

ERRH:
35590   intRetVal = 0
35600   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
35610   Case Else
35620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
35630   End Select
35640   Resume EXITP

End Sub

Private Sub cmdCountries_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_MouseMove"

35710   With Me
35720     If blnCountries_MouseDown = False Then
35730       Select Case blnCountries_Focus
            Case True
35740         .cmdCountries_raised_focus_dots_img.Visible = True
35750         .cmdCountries_raised_focus_img.Visible = False
35760       Case False
35770         .cmdCountries_raised_focus_img.Visible = True
35780         .cmdCountries_raised_focus_dots_img.Visible = False
35790       End Select
35800       .cmdCountries_raised_img.Visible = False
35810       .cmdCountries_raised_semifocus_dots_img.Visible = False
35820       .cmdCountries_sunken_focus_dots_img.Visible = False
35830       .cmdCountries_raised_img_dis.Visible = False
35840     End If
35850   End With

EXITP:
35860   Exit Sub

ERRH:
35870   Select Case ERR.Number
        Case Else
35880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
35890   End Select
35900   Resume EXITP

End Sub

Private Sub cmdCountries_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

36000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_MouseUp"

36010   With Me
36020     .cmdCountries_raised_focus_dots_img.Visible = True
36030     .cmdCountries_raised_img.Visible = False
36040     .cmdCountries_raised_semifocus_dots_img.Visible = False
36050     .cmdCountries_raised_focus_img.Visible = False
36060     .cmdCountries_sunken_focus_dots_img.Visible = False
36070     .cmdCountries_raised_img_dis.Visible = False
36080     blnCountries_MouseDown = False
36090   End With

EXITP:
36100   Exit Sub

ERRH:
36110   Select Case ERR.Number
        Case Else
36120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
36130   End Select
36140   Resume EXITP

End Sub

Private Sub cmdCountries_LostFocus()

36200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_LostFocus"

36210   With Me
36220     .cmdCountries_raised_img.Visible = True
36230     .cmdCountries_raised_semifocus_dots_img.Visible = False
36240     .cmdCountries_raised_focus_img.Visible = False
36250     .cmdCountries_raised_focus_dots_img.Visible = False
36260     .cmdCountries_sunken_focus_dots_img.Visible = False
36270     .cmdCountries_raised_img_dis.Visible = False
36280     blnCountries_Focus = False
36290   End With

EXITP:
36300   Exit Sub

ERRH:
36310   Select Case ERR.Number
        Case Else
36320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
36330   End Select
36340   Resume EXITP

End Sub

Public Sub chkIncludeNotes_AfterUpdate()

36400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeNotes_AfterUpdate"

36410   With Me
36420     Select Case .chkIncludeNotes
          Case True
36430       .chkIncludeNotes_lbl.FontBold = True
36440     Case False
36450       .chkIncludeNotes_lbl.FontBold = False
36460     End Select
36470   End With

EXITP:
36480   Exit Sub

ERRH:
36490   Select Case ERR.Number
        Case Else
36500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
36510   End Select
36520   Resume EXITP

End Sub

Private Sub chkIncludeNotes_KeyDown(KeyCode As Integer, Shift As Integer)

36600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeNotes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

36610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
36620   intShiftDown = (Shift And acShiftMask) > 0
36630   intAltDown = (Shift And acAltMask) > 0
36640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
36650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
36660     Select Case intRetVal
          Case vbKeyTab
36670       With Me
36680         intRetVal = 0
36690         .chkPageOf.SetFocus
36700       End With
36710     End Select
36720   End If

        ' ** Shift keys.
36730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
36740     Select Case intRetVal
          Case vbKeyTab
36750       With Me
36760         intRetVal = 0
36770         .cmdCountries.SetFocus
36780       End With
36790     End Select
36800   End If

EXITP:
36810   KeyCode = intRetVal
36820   Exit Sub

ERRH:
36830   intRetVal = 0
36840   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
36850   Case Else
36860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
36870   End Select
36880   Resume EXITP

End Sub

Public Sub chkPageOf_AfterUpdate()

36900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

36910   With Me
36920     Select Case .chkPageOf
          Case True
36930       .chkPageOf_lbl.FontBold = True
36940     Case False
36950       .chkPageOf_lbl.FontBold = False
36960     End Select
36970   End With

EXITP:
36980   Exit Sub

ERRH:
36990   Select Case ERR.Number
        Case Else
37000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
37010   End Select
37020   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

37100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

37110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
37120   intShiftDown = (Shift And acShiftMask) > 0
37130   intAltDown = (Shift And acAltMask) > 0
37140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
37150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
37160     Select Case intRetVal
          Case vbKeyTab
37170       With Me
37180         intRetVal = 0
37190         If .cmdPrint.Enabled = True Then
37200           .cmdPrint.SetFocus
37210         ElseIf .cmdEdit.Enabled = True Then
37220           .cmdEdit.SetFocus
37230         ElseIf .cmdUpdate.Enabled = True Then
37240           .cmdUpdate.SetFocus
37250         Else
37260           .cmdClose.SetFocus
37270         End If
37280       End With
37290     End Select
37300   End If

        ' ** Shift keys.
37310   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
37320     Select Case intRetVal
          Case vbKeyTab
37330       With Me
37340         intRetVal = 0
37350         .chkIncludeNotes.SetFocus
37360       End With
37370     End Select
37380   End If

EXITP:
37390   KeyCode = intRetVal
37400   Exit Sub

ERRH:
37410   intRetVal = 0
37420   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
37430   Case Else
37440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
37450   End Select
37460   Resume EXITP

End Sub

Private Sub ResetFilter_Set()

37500 On Error GoTo ERRH

        Const THIS_PROC As String = "ResetFilter_Set"

37510   With Me

37520     If .opgActive <> .opgActive_optInclude.OptionValue Or _
              .opgFund <> .opgFund_optInclude.OptionValue Or _
              .opgMetal <> .opgMetal_optInclude.OptionValue Or _
              .opgBMU <> .opgBMU_optInclude.OptionValue Or _
              .opgAlt <> .opgAlt_optInclude.OptionValue Or _
              .opgUnit <> .opgUnit_optInclude.OptionValue Then
37530       .cmdResetFilter.Enabled = True
37540       .cmdResetFilter_raised_img.Visible = True
37550       .cmdResetFilter_raised_semifocus_dots_img.Visible = False
37560       .cmdResetFilter_raised_focus_img.Visible = False
37570       .cmdResetFilter_raised_focus_dots_img.Visible = False
37580       .cmdResetFilter_sunken_focus_dots_img.Visible = False
37590       .cmdResetFilter_raised_img_dis.Visible = False
37600     Else
37610       .cmdResetFilter.Enabled = False
37620       .cmdResetFilter_raised_img_dis.Visible = True
37630       .cmdResetFilter_raised_img.Visible = False
37640       .cmdResetFilter_raised_semifocus_dots_img.Visible = False
37650       .cmdResetFilter_raised_focus_img.Visible = False
37660       .cmdResetFilter_raised_focus_dots_img.Visible = False
37670       .cmdResetFilter_sunken_focus_dots_img.Visible = False
37680     End If

37690   End With

EXITP:
37700   Exit Sub

ERRH:
37710   Select Case ERR.Number
        Case Else
37720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
37730   End Select
37740   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

37800 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

37810   With Me
37820     blnGTR_Emblem = False
37830     For lngX = 1& To 24&
37840       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
37850     Next
37860     If .cmdClose.Left <> lngClose_Left Then
37870       .cmdClose.Left = lngClose_Left
37880       .cmdUpdate.Left = (.cmdClose.Left - lngBtn_Offset)
37890       .cmdEdit.Left = (.cmdUpdate.Left - lngBtn_Offset)
37900       .cmdHistory.Left = (.cmdEdit.Left - lngBtn_Offset)
37910     End If
37920   End With

EXITP:
37930   Exit Sub

ERRH:
37940   Select Case ERR.Number
        Case Else
37950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
37960   End Select
37970   Resume EXITP

End Sub
