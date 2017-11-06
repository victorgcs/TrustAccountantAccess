VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency"

'VGC 10/02/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** Shortcut Alt keys responsive from this form:
' **   Active:           A {curr_active on frmCurrency_Sub}
' **   Close:            C {cmdClose}
' **   Edit:             E {cmdEdit}
' **   Excel:            L {cmdExcel}
' **   Countries:        N {cmdCountries}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Currency Word:    R {curr_word1 on frmCurrency_Sub}
' **   Reset:            S {cmdResetFilter}
' **   Rate Date:        T {curr_date on frmCurrency_Sub}
' **   Update:           U {cmdUpdate}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}
' **   Rate 1:           1 {curr_rate1 on frmCurrency_Sub}
' **   Rate 2:           2 {curr_rate2 on frmCurrency_Sub}
' **   Country:            {country_name1 on frmCurrency_Sub}

' ** Shortcut Ctrl keys responsive from this form:
' **   Exclude Active:   E {opgActive_optExclude}
' **   Page Count:       G {chkPageOf}
' **   Include Active:   I {opgActive_optInclude}
' **   Only Active:      O {opgActive_optOnly}
' **   Save:             S {cmdSave on frmCurrency_Sub}

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

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean
Private blnClearFilter_Focus As Boolean, blnClearFilter_MouseDown As Boolean
Private blnCountries_Focus As Boolean, blnCountries_MouseDown As Boolean

Private blnGoingToReport As Boolean
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

150           CLR_DISABLED_FG = CLR_DKGRY
160           CLR_DISABLED_BG = CLR_LTTEAL

170           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

      #If NoExcel Then
200           .cmdExcel.Enabled = False
      #Else
210           .cmdExcel.Enabled = True
      #End If

220           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

230           .cmdEdit.Enabled = True
240           .cmdUpdate.Enabled = False

250           .frmCurrency_Sub.Form.curr_fund.Visible = False
260           .frmCurrency_Sub.Form.curr_bmu.Visible = False
270           .frmCurrency_Sub.Form.curr_metal.Visible = False
280           .frmCurrency_Sub.Form.curr_alt.Visible = False
290           .frmCurrency_Sub.Form.curr_unit.Visible = False

300           opgActive_AfterUpdate  ' ** Procedure: Below.
310           opgFund_AfterUpdate  ' ** Procedure: Below.
320           opgMetal_AfterUpdate  ' ** Procedure: Below.
330           opgBMU_AfterUpdate  ' ** Procedure: Below.
340           opgAlt_AfterUpdate  ' ** Procedure: Below.
350           opgUnit_AfterUpdate  ' ** Procedure: Below.
360           chkIncludeNotes_AfterUpdate  ' ** Procedure: Below.
370           chkPageOf_AfterUpdate  ' ** Procedure: Below.
380           ResetFilter_Set  ' ** Procedure: Below.

390           .frmCurrency_Sub.SetFocus

400         Else
410           Cancel = -1
420         End If
430       Else
440         Cancel = -1
450       End If
460     End With

470     DoCmd.Hourglass False

480     If Cancel = -1 Then
490       Beep
500       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
510     End If

EXITP:
520     Exit Sub

ERRH:
530     DoCmd.Hourglass False
540     Select Case ERR.Number
        Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
560     End Select
570     Resume EXITP

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
        ' **   Active:           A {curr_active on frmCurrency_Sub}
        ' **   Currency Word:    R {curr_word1 on frmCurrency_Sub}
        ' **   Rate Date:        T {curr_date on frmCurrency_Sub}
        ' **   Rate 1:           1 {curr_rate1 on frmCurrency_Sub}
        ' **   Rate 2:           2 {curr_rate2 on frmCurrency_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Exclude Active:   E {opgActive_optExclude}
        ' **   Page Count:       G {chkPageOf}
        ' **   Include Active:   I {opgActive_optInclude}
        ' **   Only Active:      O {opgActive_optOnly}
        ' **   Save:             S {cmdSave on frmCurrency_Sub}

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
750           .frmCurrency_Sub.SetFocus
760           .frmCurrency_Sub.Form.curr_active.SetFocus
770           If .frmCurrency_Sub.Form.curr_active.Enabled = True And .frmCurrency_Sub.Form.curr_active.Locked = False Then
780             .frmCurrency_Sub.Form.curr_active = (Not .frmCurrency_Sub.Form.curr_active)
790             .frmCurrency_Sub.Form.curr_active_AfterUpdate  ' ** Form Procedure: frmCurrency_Sub.
800           End If
810         End With
820       Case vbKeyR
830         With Me
840           intRetVal = 0
850           .frmCurrency_Sub.SetFocus
860           .frmCurrency_Sub.Form.curr_word1.SetFocus
870         End With
880       Case vbKeyT
890         With Me
900           intRetVal = 0
910           .frmCurrency_Sub.SetFocus
920           .frmCurrency_Sub.Form.curr_date.SetFocus
930         End With
940       Case vbKeyX
950         intRetVal = 0
960         cmdClose_Click  ' ** Procedure: Below.
970       Case vbKey1
980         With Me
990           intRetVal = 0
1000          .frmCurrency_Sub.SetFocus
1010          .frmCurrency_Sub.Form.curr_rate1.SetFocus
1020        End With
1030      Case vbKey2
1040        With Me
1050          intRetVal = 0
1060          .frmCurrency_Sub.SetFocus
1070          .frmCurrency_Sub.Form.curr_rate2.SetFocus
1080        End With
1090      End Select
1100    End If

        ' ** Ctrl keys.
1110    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1120      Select Case intRetVal
          Case vbKeyE
1130        With Me
1140          intRetVal = 0
1150          .opgActive.SetFocus
1160          .opgActive = .opgActive_optExclude.OptionValue
1170          opgActive_AfterUpdate  ' ** Procedure: Below.
1180        End With
1190      Case vbKeyG
1200        With Me
1210          intRetVal = 0
1220          If .chkPageOf.Enabled = True Then
1230            .chkPageOf.SetFocus
1240            .chkPageOf = (Not .chkPageOf)
1250            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1260          End If
1270        End With
1280      Case vbKeyI
1290        With Me
1300          intRetVal = 0
1310          .opgActive.SetFocus
1320          .opgActive = .opgActive_optInclude.OptionValue
1330          opgActive_AfterUpdate  ' ** Procedure: Below.
1340        End With
1350      Case vbKeyO
1360        With Me
1370          intRetVal = 0
1380          .opgActive.SetFocus
1390          .opgActive = .opgActive_optOnly.OptionValue
1400          opgActive_AfterUpdate  ' ** Procedure: Below.
1410        End With
1420      Case vbKeyS
1430        With Me
1440          intRetVal = 0
1450          .frmCurrency_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmCurrency_Sub.
1460        End With
1470      End Select
1480    End If

        ' ** Ctrl-Shift keys.
1490    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1500      Select Case intRetVal
          Case vbKeyE
1510        With Me
1520          intRetVal = 0
1530          .opgFund.SetFocus
1540          .opgFund = .opgFund_optExclude.OptionValue
1550          opgFund_AfterUpdate  ' ** Procedure: Below.
1560        End With
1570      Case vbKeyF
1580        With Me
1590          intRetVal = 0
1600          .FocusHolder.SetFocus
1610        End With
1620      Case vbKeyI
1630        With Me
1640          intRetVal = 0
1650          .opgFund.SetFocus
1660          .opgFund = .opgFund_optInclude.OptionValue
1670          opgFund_AfterUpdate  ' ** Procedure: Below.
1680        End With
1690      Case vbKeyL
1700        With Me
1710          intRetVal = 0
1720          .opgMetal.SetFocus
1730          .opgMetal = .opgMetal_optOnly.OptionValue
1740          opgMetal_AfterUpdate  ' ** Procedure: Below.
1750        End With
1760      Case vbKeyN
1770        With Me
1780          intRetVal = 0
1790          .opgMetal.SetFocus
1800          .opgMetal = .opgMetal_optInclude.OptionValue
1810          opgMetal_AfterUpdate  ' ** Procedure: Below.
1820        End With
1830      Case vbKeyO
1840        With Me
1850          intRetVal = 0
1860          .opgFund.SetFocus
1870          .opgFund = .opgFund_optOnly.OptionValue
1880          opgFund_AfterUpdate  ' ** Procedure: Below.
1890        End With
1900      Case vbKeyX
1910        With Me
1920          intRetVal = 0
1930          .opgMetal.SetFocus
1940          .opgMetal = .opgMetal_optExclude.OptionValue
1950          opgMetal_AfterUpdate  ' ** Procedure: Below.
1960        End With
1970      End Select
1980    End If

        ' ** Ctrl-Alt keys.
1990    If intCtrlDown And intAltDown And (Not intShiftDown) Then
2000      Select Case intRetVal
          Case vbKeyD
2010        With Me
2020          intRetVal = 0
2030          .opgUnit.SetFocus
2040          .opgUnit = .opgUnit_optExclude.OptionValue
2050          opgUnit_AfterUpdate  ' ** Procedure: Below.
2060        End With
2070      Case vbKeyE
2080        With Me
2090          intRetVal = 0
2100          .opgBMU.SetFocus
2110          .opgBMU = .opgBMU_optExclude.OptionValue
2120          opgBMU_AfterUpdate  ' ** Procedure: Below.
2130        End With
2140      Case vbKeyI
2150        With Me
2160          intRetVal = 0
2170          .opgBMU.SetFocus
2180          .opgBMU = .opgBMU_optInclude.OptionValue
2190          opgBMU_AfterUpdate  ' ** Procedure: Below.
2200        End With
2210      Case vbKeyL
2220        With Me
2230          intRetVal = 0
2240          .opgAlt.SetFocus
2250          .opgAlt = .opgAlt_optOnly.OptionValue
2260          opgAlt_AfterUpdate  ' ** Procedure: Below.
2270        End With
2280      Case vbKeyN
2290        With Me
2300          intRetVal = 0
2310          .opgAlt.SetFocus
2320          .opgAlt = .opgAlt_optInclude.OptionValue
2330          opgAlt_AfterUpdate  ' ** Procedure: Below.
2340        End With
2350      Case vbKeyO
2360        With Me
2370          intRetVal = 0
2380          .opgBMU.SetFocus
2390          .opgBMU = .opgBMU_optOnly.OptionValue
2400          opgBMU_AfterUpdate  ' ** Procedure: Below.
2410        End With
2420      Case vbKeyU
2430        With Me
2440          intRetVal = 0
2450          .opgUnit.SetFocus
2460          .opgUnit = .opgUnit_optInclude.OptionValue
2470          opgUnit_AfterUpdate  ' ** Procedure: Below.
2480        End With
2490      Case vbKeyX
2500        With Me
2510          intRetVal = 0
2520          .opgAlt.SetFocus
2530          .opgAlt = .opgAlt_optExclude.OptionValue
2540          opgAlt_AfterUpdate  ' ** Procedure: Below.
2550        End With
2560      Case vbKeyY
2570        With Me
2580          intRetVal = 0
2590          .opgUnit.SetFocus
2600          .opgUnit = .opgUnit_optOnly.OptionValue
2610          opgUnit_AfterUpdate  ' ** Procedure: Below.
2620        End With
2630      End Select
2640    End If

EXITP:
2650    KeyCode = intRetVal
2660    Exit Sub

ERRH:
2670    intRetVal = 0
2680    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2690    Case Else
2700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2710    End Select
2720    Resume EXITP

End Sub

Private Sub Form_Timer()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2810    With Me
2820      .TimerInterval = 0&
2830      If gblnSetFocus = True Then
2840        gblnSetFocus = False
2850        strFormName = vbNullString
2860  On Error Resume Next
2870        Do Until strFormName = THIS_NAME
2880          DoCmd.SelectObject acForm, THIS_NAME, False
2890          strFormName = Screen.ActiveForm.Name
2900        Loop
2910  On Error GoTo ERRH
2920        .frmCurrency_Sub.SetFocus
2930      End If
2940      If gblnGoToReport = True Then
2950        Select Case blnGoingToReport
            Case True
2960          DoCmd.Hourglass False  ' ** It should already be off.
2970          gblnGoToReport = False
2980          blnGoingToReport = False
2990          DoEvents
3000        Case False
3010          DoCmd.Hourglass True  ' ** Make sure it's still running.
3020          DoEvents
3030          blnGoingToReport = True
3040          If garr_varGoToReport(GTR_FOC6) = True Then
3050            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
3060          End If
3070          Beep
3080          DoCmd.Hourglass False
3090          DoEvents
3100          .TimerInterval = (GTR_WAIT * 2&)
3110        End Select
3120      End If
3130      DoCmd.Hourglass False
3140    End With

EXITP:
3150    Exit Sub

ERRH:
3160    DoCmd.Hourglass False
3170    Select Case ERR.Number
        Case Else
3180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3190    End Select
3200    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

3310    With Me

3320      DoCmd.Hourglass True
3330      DoEvents

3340      strDocName = "frmCurrency_Country"
3350      If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
3360        DoCmd.Close acForm, strDocName
3370        DoEvents
3380      End If

          ' ** Make sure all reports are closed.
3390      Do While Reports.Count > 0
3400        DoCmd.Close acReport, Reports(0).Name
3410        DoEvents
3420      Loop

3430      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3440      Set dbs = CurrentDb
          ' ** Empty tblCurrency_Staging.
3450      Set qdf = dbs.QueryDefs("qryCurrency_Sub_05")
3460      qdf.Execute
3470      Set qdf = Nothing
3480      dbs.Close
3490      Set dbs = Nothing

3500      Select Case strCallingForm
          Case "frmMenu_ForeignExchange"
3510        gstrReturningForm = THIS_NAME
3520        gblnSetFocus = True
3530        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Maintenance"
3540      End Select

3550    End With

EXITP:
3560    Set qdf = Nothing
3570    Set dbs = Nothing
3580    Exit Sub

ERRH:
3590    DoCmd.Hourglass False
3600    Select Case ERR.Number
        Case Else
3610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3620    End Select
3630    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3710    With Me
3720      If .cmdResetFilter_raised_focus_dots_img.Visible = True Or .cmdResetFilter_raised_focus_img.Visible = True Then
3730        Select Case blnClearFilter_Focus
            Case True
3740          .cmdResetFilter_raised_semifocus_dots_img.Visible = True
3750          .cmdResetFilter_raised_img.Visible = False
3760        Case False
3770          .cmdResetFilter_raised_img.Visible = True
3780          .cmdResetFilter_raised_semifocus_dots_img.Visible = False
3790        End Select
3800        .cmdResetFilter_raised_focus_img.Visible = False
3810        .cmdResetFilter_raised_focus_dots_img.Visible = False
3820        .cmdResetFilter_sunken_focus_dots_img.Visible = False
3830        .cmdResetFilter_raised_img_dis.Visible = False
3840      End If
3850      If .cmdCountries_raised_focus_dots_img.Visible = True Or .cmdCountries_raised_focus_img.Visible = True Then
3860        Select Case blnCountries_Focus
            Case True
3870          .cmdCountries_raised_semifocus_dots_img.Visible = True
3880          .cmdCountries_raised_img.Visible = False
3890        Case False
3900          .cmdCountries_raised_img.Visible = True
3910          .cmdCountries_raised_semifocus_dots_img.Visible = False
3920        End Select
3930        .cmdCountries_raised_focus_img.Visible = False
3940        .cmdCountries_raised_focus_dots_img.Visible = False
3950        .cmdCountries_sunken_focus_dots_img.Visible = False
3960        .cmdCountries_raised_img_dis.Visible = False
3970      End If
3980    End With

EXITP:
3990    Exit Sub

ERRH:
4000    Select Case ERR.Number
        Case Else
4010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4020    End Select
4030    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

4110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub cmdClose_Click()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim msgResponse As VbMsgBoxResult

4210    With Me
4220      If .currstage_changed = True And .cmdUpdate.Enabled = True Then
4230        Beep
4240        msgResponse = MsgBox("You have made changes, but have not saved them." & vbCrLf & vbCrLf & _
              "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved")
4250      Else
4260        msgResponse = vbYes
4270      End If
4280      If msgResponse = vbYes Then
4290        DoCmd.Hourglass True
4300        DoEvents
4310        .TimerInterval = 0&
4320        DoCmd.Close acForm, THIS_NAME
4330      Else
4340        .cmdUpdate.SetFocus
4350      End If
4360    End With

EXITP:
4370    Exit Sub

ERRH:
4380    Select Case ERR.Number
        Case Else
4390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4400    End Select
4410    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
4570        With Me
4580          intRetVal = 0
4590          .frmCurrency_Sub.SetFocus
4600          .frmCurrency_Sub.Form.FocusHolder.SetFocus
4610        End With
4620      End Select
4630    End If

        ' ** Shift keys.
4640    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4650      Select Case intRetVal
          Case vbKeyTab
4660        With Me
4670          intRetVal = 0
4680          If .cmdUpdate.Enabled = True Then
4690            .cmdUpdate.SetFocus
4700          ElseIf .cmdEdit.Enabled = True Then
4710            .cmdEdit.SetFocus
4720          ElseIf .cmdExcel.Enabled = True Then
4730            .cmdExcel.SetFocus
4740          ElseIf .cmdCountries.Enabled = True Then
4750            .cmdCountries.SetFocus
4760          Else
4770            .opgUnit.SetFocus
4780          End If
4790        End With
4800      End Select
4810    End If

EXITP:
4820    KeyCode = intRetVal
4830    Exit Sub

ERRH:
4840    intRetVal = 0
4850    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4860    Case Else
4870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4880    End Select
4890    Resume EXITP

End Sub

Private Sub cmdEdit_Click()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_Click"

4910    With Me
4920      If .frmCurrency_Sub.Form.curr_rate1.Locked = True Then
            ' ** These should be Enabled, but Locked.

            ' ** These should be False and Disabled.
4930        .currstage_changed = False
4940        .cmdUpdate.Enabled = False

4950        .frmCurrency_Sub.Form.curr_rate1.Enabled = True
4960        .frmCurrency_Sub.Form.curr_rate1.Locked = False
4970        .frmCurrency_Sub.Form.curr_rate1.ForeColor = CLR_VDKGRY
4980        .frmCurrency_Sub.Form.curr_rate1.BackColor = CLR_WHT
4990        .frmCurrency_Sub.Form.curr_rate2.Enabled = True
5000        .frmCurrency_Sub.Form.curr_rate2.Locked = False
5010        .frmCurrency_Sub.Form.curr_rate2.ForeColor = CLR_VDKGRY
5020        .frmCurrency_Sub.Form.curr_rate2.BackColor = CLR_WHT
5030        .frmCurrency_Sub.Form.curr_date.Enabled = True
5040        .frmCurrency_Sub.Form.curr_date.Locked = False
5050        .frmCurrency_Sub.Form.curr_date.ForeColor = CLR_VDKGRY
5060        .frmCurrency_Sub.Form.curr_date.BackColor = CLR_WHT
5070        .frmCurrency_Sub.Form.curr_notes.Enabled = True
5080        .frmCurrency_Sub.Form.curr_notes.Locked = False
5090        .frmCurrency_Sub.Form.curr_notes.ForeColor = CLR_VDKGRY
5100        .frmCurrency_Sub.Form.curr_notes.BackColor = CLR_WHT
5110        .frmCurrency_Sub.Form.curr_active.Enabled = True
5120        .frmCurrency_Sub.Form.curr_active.Locked = False
5130        .frmCurrency_Sub.Form.curr_active_box1.ForeColor = CLR_WHT
5140        .frmCurrency_Sub.Form.curr_active_box2.ForeColor = CLR_LTBLU2
5150        .frmCurrency_Sub.Form.country_active.Enabled = True
5160        .frmCurrency_Sub.Form.country_active.Locked = False
5170        .frmCurrency_Sub.Form.country_active_box1.ForeColor = CLR_WHT
5180        .frmCurrency_Sub.Form.country_active_box2.ForeColor = CLR_LTBLU2

            ' ** Other fields   Enabled  Locked
            ' ** =============  =======  ======
            ' **   curr_code:     Yes     Yes
            ' **   curr_word1:    Yes     Yes
            ' **   curr_word2:    No      Yes
            ' **   curr_name1:    Yes     Yes
            ' **   curr_name2:    No      Yes
            ' **   curr_iso:      Yes     Yes
            ' **   curr_decimal:  Yes     Yes
            ' **   country_code3: Yes     Yes
            ' **   country_name:  Yes     Yes
            ' **   country_iso:   Yes     Yes
            ' **   curr_fund:     No      Yes
            ' **   curr_metal:    No      Yes
            ' **   curr_bmu:      No      Yes
            ' **   curr_alt:      No      Yes

5190        .frmCurrency_Sub.SetFocus
5200        .frmCurrency_Sub.Form.curr_iso.SetFocus

5210      Else
5220        If .currstage_changed = False And .cmdUpdate.Enabled = False Then
5230          .frmCurrency_Sub.Form.curr_rate1.Locked = True
5240          .frmCurrency_Sub.Form.curr_rate1.ForeColor = CLR_DISABLED_FG
5250          .frmCurrency_Sub.Form.curr_rate1.BackColor = CLR_DISABLED_BG
5260          .frmCurrency_Sub.Form.curr_rate2.Locked = True
5270          .frmCurrency_Sub.Form.curr_rate2.ForeColor = CLR_DISABLED_FG
5280          .frmCurrency_Sub.Form.curr_rate2.BackColor = CLR_DISABLED_BG
5290          .frmCurrency_Sub.Form.curr_date.Locked = True
5300          .frmCurrency_Sub.Form.curr_date.ForeColor = CLR_DISABLED_FG
5310          .frmCurrency_Sub.Form.curr_date.BackColor = CLR_DISABLED_BG
5320          .frmCurrency_Sub.Form.curr_notes.Locked = True
5330          .frmCurrency_Sub.Form.curr_notes.ForeColor = CLR_DISABLED_FG
5340          .frmCurrency_Sub.Form.curr_notes.BackColor = CLR_DISABLED_BG
5350          .frmCurrency_Sub.Form.curr_active.Locked = True
5360          .frmCurrency_Sub.Form.curr_active_box1.ForeColor = MY_CLR_BGE
5370          .frmCurrency_Sub.Form.curr_active_box2.ForeColor = WIN_CLR_DISR
5380          .frmCurrency_Sub.Form.country_active.Locked = True
5390          .frmCurrency_Sub.Form.country_active_box1.ForeColor = MY_CLR_BGE
5400          .frmCurrency_Sub.Form.country_active_box2.ForeColor = WIN_CLR_DISR
5410        Else
5420          Beep
5430        End If
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

Private Sub cmdEdit_KeyDown(KeyCode As Integer, Shift As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5620    intShiftDown = (Shift And acShiftMask) > 0
5630    intAltDown = (Shift And acAltMask) > 0
5640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5660      Select Case intRetVal
          Case vbKeyTab
5670        With Me
5680          intRetVal = 0
5690          If .cmdUpdate.Enabled = True Then
5700            .cmdUpdate.SetFocus
5710          Else
5720            .cmdClose.SetFocus
5730          End If
5740        End With
5750      End Select
5760    End If

        ' ** Shift keys.
5770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5780      Select Case intRetVal
          Case vbKeyTab
5790        With Me
5800          intRetVal = 0
5810          If .cmdExcel.Enabled = True Then
5820            .cmdExcel.SetFocus
5830          ElseIf .cmdWord.Enabled = True Then
5840            .cmdWord.SetFocus
5850          ElseIf .cmdCountries.Enabled = True Then
5860            .cmdCountries.SetFocus
5870          Else
5880            .opgUnit.SetFocus
5890          End If
5900        End With
5910      End Select
5920    End If

EXITP:
5930    KeyCode = intRetVal
5940    Exit Sub

ERRH:
5950    intRetVal = 0
5960    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5970    Case Else
5980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5990    End Select
6000    Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnContinue As Boolean

6110    blnContinue = True

6120    With Me

6130      DoCmd.Hourglass True
6140      DoEvents

6150      Set dbs = CurrentDb

          ' ** tblCurrency_Staging, just currstage_changed = True.
6160      Set qdf = dbs.QueryDefs("qryCurrency_Sub_08")
6170      Set rst = qdf.OpenRecordset
6180      If rst.BOF = True And rst.EOF = True Then
6190        blnContinue = False
6200        DoCmd.Hourglass False
6210        Beep
6220        MsgBox "No changed records were found.", vbInformation + vbOKOnly, "Nothing To Do"
6230        .currstage_changed = False
6240        rst.Close
6250      Else
6260        rst.Close
6270      End If
6280      Set rst = Nothing
6290      Set qdf = Nothing

6300      If blnContinue = True Then

            ' ** Update qryCurrency_Sub_09 (qryCurrency_Sub_08 (tblCurrency_Staging, just
            ' ** currstage_changed = True), linked to tblCurrency, with .._new fields), for curr_rate1.
6310        Set qdf = dbs.QueryDefs("qryCurrency_Sub_10a")
6320        qdf.Execute
6330        Set qdf = Nothing

            ' ** Update qryCurrency_Sub_09 (qryCurrency_Sub_08 (tblCurrency_Staging, just
            ' ** currstage_changed = True), linked to tblCurrency, with .._new fields), for curr_rate2.
6340        Set qdf = dbs.QueryDefs("qryCurrency_Sub_10b")
6350        qdf.Execute
6360        Set qdf = Nothing

            ' ** Update qryCurrency_Sub_09 (qryCurrency_Sub_08 (tblCurrency_Staging, just
            ' ** currstage_changed = True), linked to tblCurrency, with .._new fields), for curr_date.
6370        Set qdf = dbs.QueryDefs("qryCurrency_Sub_10c")
6380        qdf.Execute
6390        Set qdf = Nothing

            ' ** Update qryCurrency_Sub_09 (qryCurrency_Sub_08 (tblCurrency_Staging, just
            ' ** currstage_changed = True), linked to tblCurrency, with .._new fields), for curr_notes.
6400        Set qdf = dbs.QueryDefs("qryCurrency_Sub_10d")
6410        qdf.Execute
6420        Set qdf = Nothing

            ' ** Update qryCurrency_Sub_09 (qryCurrency_Sub_08 (tblCurrency_Staging, just
            ' ** currstage_changed = True), linked to tblCurrency, with .._new fields), for curr_active.
6430        Set qdf = dbs.QueryDefs("qryCurrency_Sub_10e")
6440        qdf.Execute
6450        Set qdf = Nothing

            ' ** Update qryCurrency_Sub_11 (qryCurrency_Sub_08 (tblCurrency_Staging, just
            ' ** currstage_changed = True), linked to tblCountry, with .._new fields), for country_active.
6460        Set qdf = dbs.QueryDefs("qryCurrency_Sub_12a")
6470        qdf.Execute
6480        Set qdf = Nothing

6490        Beep
6500        DoCmd.Hourglass False
6510        MsgBox "Updates completed.", vbInformation + vbOKOnly, ("Update Successful" & Space(40))

6520        .currstage_changed = False
6530        .cmdClose.SetFocus
6540        .cmdUpdate.Enabled = False

6550      End If  ' ** blnContinue.

6560      dbs.Close

6570      DoCmd.Hourglass False

6580    End With

EXITP:
6590    Set rst = Nothing
6600    Set qdf = Nothing
6610    Set dbs = Nothing
6620    Exit Sub

ERRH:
6630    DoCmd.Hourglass False
6640    Select Case ERR.Number
        Case Else
6650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6660    End Select
6670    Resume EXITP

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
6870          If .cmdEdit.Enabled = True Then
6880            .cmdEdit.SetFocus
6890          ElseIf .cmdExcel.Enabled = True Then
6900            .cmdExcel.SetFocus
6910          ElseIf .cmdCountries.Enabled = True Then
6920            .cmdCountries.SetFocus
6930          Else
6940            .opgUnit.SetFocus
6950          End If
6960        End With
6970      End Select
6980    End If

EXITP:
6990    KeyCode = intRetVal
7000    Exit Sub

ERRH:
7010    intRetVal = 0
7020    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7030    Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strFilter As String, strDocName As String

7110    With Me

7120      Select Case .frmCurrency_Sub.Form.FilterOn
          Case True
7130        strFilter = .frmCurrency_Sub.Form.Filter
7140      Case False
7150        strFilter = vbNullString
7160      End Select

7170      strDocName = "rptCurrencies"
7180      DoCmd.OpenReport strDocName, acViewPreview, , strFilter
7190      DoCmd.Maximize
7200      DoCmd.RunCommand acCmdFitToWindow

7210    End With

EXITP:
7220    Exit Sub

ERRH:
7230    Select Case ERR.Number
        Case Else
7240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7250    End Select
7260    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

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
7390          .cmdPrint.SetFocus
7400        End With
7410      End Select
7420    End If

        ' ** Shift keys.
7430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7440      Select Case intRetVal
          Case vbKeyTab
7450        With Me
7460          intRetVal = 0
7470          If .cmdCountries.Enabled = True Then
7480            .cmdCountries.SetFocus
7490          Else
7500            .opgUnit.SetFocus
7510          End If
7520        End With
7530      End Select
7540    End If

        ' ** Ctrl-Shift keys.
7550    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7560      Select Case intRetVal
          Case vbKeyTab
7570        With Me
7580          intRetVal = 0
7590          .chkPageOf.SetFocus
7600        End With
7610      End Select
7620    End If

EXITP:
7630    KeyCode = intRetVal
7640    Exit Sub

ERRH:
7650    intRetVal = 0
7660    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7670    Case Else
7680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7690    End Select
7700    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strFilter As String, strDocName As String

7810    With Me

7820      Select Case .frmCurrency_Sub.Form.FilterOn
          Case True
7830        strFilter = .frmCurrency_Sub.Form.Filter
7840      Case False
7850        strFilter = vbNullString
7860      End Select

7870      strDocName = "rptCurrencies"
          '##GTR_Ref: rptCurrencies
7880      DoCmd.OpenReport strDocName, acViewNormal, , strFilter

7890    End With

EXITP:
7900    Exit Sub

ERRH:
7910    Me.Visible = True
7920    Select Case ERR.Number
        Case Else
7930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7940    End Select
7950    Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8020    intShiftDown = (Shift And acShiftMask) > 0
8030    intAltDown = (Shift And acAltMask) > 0
8040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8060      Select Case intRetVal
          Case vbKeyTab
8070        With Me
8080          intRetVal = 0
8090          .cmdWord.SetFocus
8100        End With
8110      End Select
8120    End If

        ' ** Shift keys.
8130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8140      Select Case intRetVal
          Case vbKeyTab
8150        With Me
8160          intRetVal = 0
8170          .cmdPreview.SetFocus
8180        End With
8190      End Select
8200    End If

EXITP:
8210    KeyCode = intRetVal
8220    Exit Sub

ERRH:
8230    intRetVal = 0
8240    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8250    Case Else
8260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8270    End Select
8280    Resume EXITP

End Sub

Private Sub cmdWord_Click()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

8310    With Me

8320      If IsNull(.UserReportPath) = True Then
8330        strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
8340      Else
8350        strRptPath = .UserReportPath
8360      End If
8370      strRptCap = "rptCurrencies_" & Format$(Date, "yyyymmdd")

8380      strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

8390      If strRptPathFile <> vbNullString Then
8400        strRpt = "rptCurrencies"
8410        DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
8420        .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
8430      End If

8440    End With

EXITP:
8450    Exit Sub

ERRH:
8460    Select Case ERR.Number
        Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
8470      Beep
8480      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
8490    Case Else
8500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8510    End Select
8520    Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

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
          Case vbKeyTab
8670        With Me
8680          intRetVal = 0
8690          If .cmdExcel.Enabled = True Then
8700            .cmdExcel.SetFocus
8710          ElseIf .cmdEdit.Enabled = True Then
8720            .cmdEdit.SetFocus
8730          ElseIf .cmdUpdate.Enabled = True Then
8740            .cmdUpdate.SetFocus
8750          Else
8760            .cmdClose.SetFocus
8770          End If
8780        End With
8790      End Select
8800    End If

        ' ** Shift keys.
8810    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8820      Select Case intRetVal
          Case vbKeyTab
8830        With Me
8840          intRetVal = 0
8850          .cmdPrint.SetFocus
8860        End With
8870      End Select
8880    End If

EXITP:
8890    KeyCode = intRetVal
8900    Exit Sub

ERRH:
8910    intRetVal = 0
8920    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8930    Case Else
8940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8950    End Select
8960    Resume EXITP

End Sub

Private Sub cmdExcel_Click()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strQry As String, strRptCap As String, strOrderBy As String
        Dim strRptPath As String, strRptPathFile As String
        Dim intPos01 As Integer
        Dim strTmp01 As String

9010    With Me
9020      lngRecsCur = .frmCurrency_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Sub.
9030      If lngRecsCur > 0& Then

9040        If IsNull(.UserReportPath) = True Then
9050          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
9060        Else
9070          strRptPath = .UserReportPath
9080        End If
9090        strRptCap = "rptCurrencies_" & Format$(Date, "yyyymmdd")

9100        strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

9110        If strRptPathFile <> vbNullString Then

9120          strOrderBy = .frmCurrency_Sub.Form.OrderBy

9130          strQry = "qryRpt_Currencies_02"

9140          Select Case .chkIncludeNotes
              Case True
9150            strQry = strQry & "_01"
9160          Case False
9170            strQry = strQry & "_00"
9180          End Select

9190          Select Case .opgActive
              Case .opgActive_optInclude.OptionValue, .opgActive_optOnly.OptionValue
9200            strQry = strQry & "_01"
9210          Case .opgActive_optExclude.OptionValue
9220            strQry = strQry & "_00"
9230          End Select

9240          Select Case .opgFund
              Case .opgFund_optInclude.OptionValue, .opgFund_optOnly.OptionValue
9250            strQry = strQry & "_01"
9260          Case .opgFund_optExclude.OptionValue
9270            strQry = strQry & "_00"
9280          End Select

9290          Select Case .opgMetal
              Case .opgMetal_optInclude.OptionValue, .opgMetal_optOnly.OptionValue
9300            strQry = strQry & "_01"
9310          Case .opgMetal_optExclude.OptionValue
9320            strQry = strQry & "_00"
9330          End Select

9340          Select Case .opgBMU
              Case .opgBMU_optInclude.OptionValue, .opgBMU_optOnly.OptionValue
9350            strQry = strQry & "_01"
9360          Case .opgBMU_optExclude.OptionValue
9370            strQry = strQry & "_00"
9380          End Select

9390          Select Case .opgAlt
              Case .opgAlt_optInclude.OptionValue, .opgAlt_optOnly.OptionValue
9400            strQry = strQry & "_01"
9410          Case .opgAlt_optExclude.OptionValue
9420            strQry = strQry & "_00"
9430          End Select

              ' ** All original field names are present in the source query.
              ' **   [curr_code]
              ' **   [curr_word1], [country_name]
              ' **   [curr_name], [curr_word1], [country_name]
              ' **   [curr_iso]
              ' **   [curr_rate1], [country_name], [curr_word1]
              ' **   [curr_rate2], [country_name], [curr_word1]
              ' **   [curr_date], [country_name], [curr_word1]
              ' **   [curr_active], [curr_word1], [country_name]
              ' **   [country_code3], [curr_word1], [country_name]
              ' **   [country_name], [curr_word1], [country_name]
              ' **   [country_iso], [country_name], [curr_word1]
              ' **   [country_active], [country_name], [curr_word1]
              ' **   [curr_decimal], [curr_word1], [country_name]
              ' **   [curr_notes], [curr_word1], [country_name]
              ' **   [curr_fund], [curr_word1], [country_name]
              ' **   [curr_metal], [curr_word1], [country_name]
              ' **   [curr_bmu], [curr_word1], [country_name]
              ' **   [curr_alt], [curr_word1], [country_name]

9440          Set dbs = CurrentDb
9450          Set qdf = dbs.QueryDefs(strQry)
9460          strTmp01 = qdf.SQL
9470          Set qdf = Nothing

9480          intPos01 = InStr(strTmp01, ";")
9490          strTmp01 = Left(strTmp01, (intPos01 - 1)) & vbCrLf
9500          strTmp01 = strTmp01 & "ORDER BY " & strOrderBy & ";"

9510          strQry = "qryRpt_Currencies_02"

9520          Set qdf = dbs.QueryDefs(strQry)
9530          qdf.SQL = strTmp01
9540          Set qdf = Nothing
9550          dbs.Close
9560          Set dbs = Nothing

9570          DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, True
9580          .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.

9590        End If  ' ** strRptPathFile.
9600      Else
9610        Beep
9620      End If  ' ** lngRecsCur.
9630    End With

EXITP:
9640    Set qdf = Nothing
9650    Set dbs = Nothing
9660    Exit Sub

ERRH:
9670    Select Case ERR.Number
        Case Else
9680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9690    End Select
9700    Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

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
          Case vbKeyTab
9870        With Me
9880          intRetVal = 0
9890          If .cmdEdit.Enabled = True Then
9900            .cmdEdit.SetFocus
9910          ElseIf .cmdUpdate.Enabled = True Then
9920            .cmdUpdate.SetFocus
9930          Else
9940            .cmdClose.SetFocus
9950          End If
9960        End With
9970      End Select
9980    End If

        ' ** Shift keys.
9990    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10000     Select Case intRetVal
          Case vbKeyTab
10010       With Me
10020         intRetVal = 0
10030         .cmdWord.SetFocus
10040       End With
10050     End Select
10060   End If

EXITP:
10070   KeyCode = intRetVal
10080   Exit Sub

ERRH:
10090   intRetVal = 0
10100   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
10110   Case Else
10120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10130   End Select
10140   Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

10210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10220   Exit Sub

ERRH:
10230   Select Case ERR.Number
        Case Else
10240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10250   End Select
10260   Resume EXITP

End Sub

Private Sub opgActive_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_box_MouseMove"

10310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10320   Exit Sub

ERRH:
10330   Select Case ERR.Number
        Case Else
10340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10350   End Select
10360   Resume EXITP

End Sub

Public Sub opgActive_AfterUpdate()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_AfterUpdate"

        Dim strFilter As String

10410   With Me

10420     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
10430       .opgActive_optInclude_lbl.FontBold = True
10440       .opgActive_optExclude_lbl.FontBold = False
10450       .opgActive_optOnly_lbl.FontBold = False
10460       strFilter = vbNullString
10470     Case .opgActive_optExclude.OptionValue
10480       .opgActive_optInclude_lbl.FontBold = False
10490       .opgActive_optExclude_lbl.FontBold = True
10500       .opgActive_optOnly_lbl.FontBold = False
10510       strFilter = "[curr_active] = False"
10520     Case .opgActive_optOnly.OptionValue
10530       .opgActive_optInclude_lbl.FontBold = False
10540       .opgActive_optExclude_lbl.FontBold = False
10550       .opgActive_optOnly_lbl.FontBold = True
10560       If .opgFund <> .opgFund_optExclude.OptionValue Then
10570         .opgFund = .opgFund_optExclude.OptionValue
10580         .opgFund_optExclude_lbl.FontBold = True
10590         .opgFund_optInclude_lbl.FontBold = False
10600         .opgFund_optOnly_lbl.FontBold = False
10610       End If
10620       If .opgBMU <> .opgBMU_optExclude.OptionValue Then
10630         .opgBMU = .opgBMU_optExclude.OptionValue
10640         .opgBMU_optExclude_lbl.FontBold = True
10650         .opgBMU_optInclude_lbl.FontBold = False
10660         .opgBMU_optOnly_lbl.FontBold = False
10670       End If
10680       If .opgMetal <> .opgMetal_optExclude.OptionValue Then
10690         .opgMetal = .opgMetal_optExclude.OptionValue
10700         .opgMetal_optExclude_lbl.FontBold = True
10710         .opgMetal_optInclude_lbl.FontBold = False
10720         .opgMetal_optOnly_lbl.FontBold = False
10730       End If
10740       If .opgAlt <> .opgAlt_optExclude.OptionValue Then
10750         .opgAlt = .opgAlt_optExclude.OptionValue
10760         .opgAlt_optExclude_lbl.FontBold = True
10770         .opgAlt_optInclude_lbl.FontBold = False
10780         .opgAlt_optOnly_lbl.FontBold = False
10790       End If
10800       strFilter = "[curr_active] = True"
10810     End Select

10820     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
            ' ** No change.
10830     Case .opgFund_optExclude.OptionValue
10840       If strFilter <> vbNullString Then strFilter = strFilter & FAND
10850       strFilter = strFilter & "[curr_fund] = False"
10860     Case .opgFund_optOnly.OptionValue
10870       If strFilter <> vbNullString Then strFilter = strFilter & FAND
10880       strFilter = strFilter & "[curr_fund] = True"
10890     End Select

10900     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
            ' ** No change.
10910     Case .opgBMU_optExclude.OptionValue
10920       If strFilter <> vbNullString Then strFilter = strFilter & FAND
10930       strFilter = strFilter & "[curr_bmu] = False"
10940     Case .opgBMU_optOnly.OptionValue
10950       If strFilter <> vbNullString Then strFilter = strFilter & FAND
10960       strFilter = strFilter & "[curr_bmu] = True"
10970     End Select

10980     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
            ' ** No change.
10990     Case .opgMetal_optExclude.OptionValue
11000       If strFilter <> vbNullString Then strFilter = strFilter & FAND
11010       strFilter = strFilter & "[curr_metal] = False"
11020     Case .opgMetal_optOnly.OptionValue
11030       If strFilter <> vbNullString Then strFilter = strFilter & FAND
11040       strFilter = strFilter & "[curr_metal] = True"
11050     End Select

11060     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
            ' ** No change.
11070     Case .opgAlt_optExclude.OptionValue
11080       If strFilter <> vbNullString Then strFilter = strFilter & FAND
11090       strFilter = strFilter & "[curr_alt] = False"
11100     Case .opgAlt_optOnly.OptionValue
11110       If strFilter <> vbNullString Then strFilter = strFilter & FAND
11120       strFilter = strFilter & "[curr_alt] = True"
11130     End Select

11140     WhatsVis  ' ** Procedure: Below.

11150     ResetFilter_Set  ' ** Procedure: Below.

11160     .frmCurrency_Sub.Form.Filter = strFilter
11170     If strFilter <> vbNullString Then
11180       .frmCurrency_Sub.Form.FilterOn = True
11190     Else
11200       .frmCurrency_Sub.Form.FilterOn = False
11210     End If

11220     DoEvents
11230     lngRecsCur = .frmCurrency_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Sub.
11240     If lngRecsCur > 0& Then
11250       .frmCurrency_Sub.SetFocus
11260     End If

11270   End With

EXITP:
11280   Exit Sub

ERRH:
11290   Select Case ERR.Number
        Case Else
11300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11310   End Select
11320   Resume EXITP

End Sub

Private Sub opgActive_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11420   intShiftDown = (Shift And acShiftMask) > 0
11430   intAltDown = (Shift And acAltMask) > 0
11440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11460     Select Case intRetVal
          Case vbKeyTab
11470       With Me
11480         intRetVal = 0
11490         If .cmdResetFilter.Enabled = True Then
11500           .cmdResetFilter.SetFocus
11510         Else
11520           .opgFund.SetFocus
11530         End If
11540       End With
11550     End Select
11560   End If

        ' ** Shift keys.
11570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11580     Select Case intRetVal
          Case vbKeyTab
11590       With Me
11600         intRetVal = 0
11610         .frmCurrency_Sub.SetFocus
11620         .frmCurrency_Sub.Form.FocusHolder.SetFocus
11630       End With
11640     End Select
11650   End If

EXITP:
11660   KeyCode = intRetVal
11670   Exit Sub

ERRH:
11680   intRetVal = 0
11690   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
11700   Case Else
11710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11720   End Select
11730   Resume EXITP

End Sub

Private Sub opgActive_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11820   intShiftDown = (Shift And acShiftMask) > 0
11830   intAltDown = (Shift And acAltMask) > 0
11840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11860     Select Case intRetVal
          Case vbKeyTab
11870       With Me
11880         intRetVal = 0
11890         If .cmdResetFilter.Enabled = True Then
11900           .cmdResetFilter.SetFocus
11910         Else
11920           .opgFund.SetFocus
11930         End If
11940       End With
11950     End Select
11960   End If

        ' ** Shift keys.
11970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11980     Select Case intRetVal
          Case vbKeyTab
11990       With Me
12000         intRetVal = 0
12010         .frmCurrency_Sub.SetFocus
12020         .frmCurrency_Sub.Form.FocusHolder.SetFocus
12030       End With
12040     End Select
12050   End If

EXITP:
12060   KeyCode = intRetVal
12070   Exit Sub

ERRH:
12080   intRetVal = 0
12090   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
12100   Case Else
12110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12120   End Select
12130   Resume EXITP

End Sub

Private Sub opgActive_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optOnly_KeyDown"

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
          Case vbKeyTab
12270       With Me
12280         intRetVal = 0
12290         If .cmdResetFilter.Enabled = True Then
12300           .cmdResetFilter.SetFocus
12310         Else
12320           .opgFund.SetFocus
12330         End If
12340       End With
12350     End Select
12360   End If

        ' ** Shift keys.
12370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12380     Select Case intRetVal
          Case vbKeyTab
12390       With Me
12400         intRetVal = 0
12410         .frmCurrency_Sub.SetFocus
12420         .frmCurrency_Sub.Form.FocusHolder.SetFocus
12430       End With
12440     End Select
12450   End If

EXITP:
12460   KeyCode = intRetVal
12470   Exit Sub

ERRH:
12480   intRetVal = 0
12490   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
12500   Case Else
12510     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12520   End Select
12530   Resume EXITP

End Sub

Private Sub opgActive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_MouseMove"

12610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12620   Exit Sub

ERRH:
12630   Select Case ERR.Number
        Case Else
12640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12650   End Select
12660   Resume EXITP

End Sub

Private Sub cmdResetFilter_GotFocus()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_GotFocus"

12710   With Me
12720     blnClearFilter_Focus = True
12730     .cmdResetFilter_raised_semifocus_dots_img.Visible = True
12740     .cmdResetFilter_raised_img.Visible = False
12750     .cmdResetFilter_raised_focus_img.Visible = False
12760     .cmdResetFilter_raised_focus_dots_img.Visible = False
12770     .cmdResetFilter_sunken_focus_dots_img.Visible = False
12780     .cmdResetFilter_raised_img_dis.Visible = False
12790   End With

EXITP:
12800   Exit Sub

ERRH:
12810   Select Case ERR.Number
        Case Else
12820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12830   End Select
12840   Resume EXITP

End Sub

Private Sub cmdResetFilter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_MouseDown"

12910   With Me
12920     blnClearFilter_MouseDown = True
12930     .cmdResetFilter_sunken_focus_dots_img.Visible = True
12940     .cmdResetFilter_raised_img.Visible = False
12950     .cmdResetFilter_raised_semifocus_dots_img.Visible = False
12960     .cmdResetFilter_raised_focus_img.Visible = False
12970     .cmdResetFilter_raised_focus_dots_img.Visible = False
12980     .cmdResetFilter_raised_img_dis.Visible = False
12990   End With

EXITP:
13000   Exit Sub

ERRH:
13010   Select Case ERR.Number
        Case Else
13020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13030   End Select
13040   Resume EXITP

End Sub

Private Sub cmdResetFilter_Click()

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_Click"

        Dim strFilter As String

13110   With Me

13120     .opgActive = .opgActive_optInclude.OptionValue
13130     .opgActive_optInclude_lbl.FontBold = True
13140     .opgActive_optExclude_lbl.FontBold = False
13150     .opgActive_optOnly_lbl.FontBold = False

13160     .opgFund = .opgFund_optInclude.OptionValue
13170     .opgFund_optInclude_lbl.FontBold = True
13180     .opgFund_optExclude_lbl.FontBold = False
13190     .opgFund_optOnly_lbl.FontBold = False

13200     .opgMetal = .opgMetal_optInclude.OptionValue
13210     .opgMetal_optInclude_lbl.FontBold = True
13220     .opgMetal_optExclude_lbl.FontBold = False
13230     .opgMetal_optOnly_lbl.FontBold = False

13240     .opgBMU = .opgBMU_optInclude.OptionValue
13250     .opgBMU_optInclude_lbl.FontBold = True
13260     .opgBMU_optExclude_lbl.FontBold = False
13270     .opgBMU_optOnly_lbl.FontBold = False

13280     .opgAlt = .opgAlt_optInclude.OptionValue
13290     .opgAlt_optInclude_lbl.FontBold = True
13300     .opgAlt_optExclude_lbl.FontBold = False
13310     .opgAlt_optOnly_lbl.FontBold = False

13320     .opgUnit = .opgUnit_optInclude.OptionValue
13330     .opgUnit_optInclude_lbl.FontBold = True
13340     .opgUnit_optExclude_lbl.FontBold = False
13350     .opgUnit_optOnly_lbl.FontBold = False

13360     strFilter = vbNullString
13370     .frmCurrency_Sub.Form.Filter = strFilter
13380     .frmCurrency_Sub.Form.FilterOn = False
13390     DoEvents

13400     ResetFilter_Set  ' ** Procedure: Below.
13410     WhatsVis  ' ** Procedure: Below.

13420   End With

EXITP:
13430   Exit Sub

ERRH:
13440   Select Case ERR.Number
        Case Else
13450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13460   End Select
13470   Resume EXITP

End Sub

Private Sub cmdResetFilter_KeyDown(KeyCode As Integer, Shift As Integer)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13520   intShiftDown = (Shift And acShiftMask) > 0
13530   intAltDown = (Shift And acAltMask) > 0
13540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13560     Select Case intRetVal
          Case vbKeyTab
13570       With Me
13580         intRetVal = 0
13590         .opgFund.SetFocus
13600       End With
13610     End Select
13620   End If

        ' ** Shift keys.
13630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13640     Select Case intRetVal
          Case vbKeyTab
13650       With Me
13660         intRetVal = 0
13670         .opgActive.SetFocus
13680       End With
13690     End Select
13700   End If

EXITP:
13710   KeyCode = intRetVal
13720   Exit Sub

ERRH:
13730   intRetVal = 0
13740   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
13750   Case Else
13760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13770   End Select
13780   Resume EXITP

End Sub

Private Sub cmdResetFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_MouseMove"

13810   With Me
13820     If blnClearFilter_MouseDown = False Then
13830       Select Case blnClearFilter_Focus
            Case True
13840         .cmdResetFilter_raised_focus_dots_img.Visible = True
13850         .cmdResetFilter_raised_focus_img.Visible = False
13860       Case False
13870         .cmdResetFilter_raised_focus_img.Visible = True
13880         .cmdResetFilter_raised_focus_dots_img.Visible = False
13890       End Select
13900       .cmdResetFilter_raised_img.Visible = False
13910       .cmdResetFilter_raised_semifocus_dots_img.Visible = False
13920       .cmdResetFilter_sunken_focus_dots_img.Visible = False
13930       .cmdResetFilter_raised_img_dis.Visible = False
13940     End If
13950   End With

EXITP:
13960   Exit Sub

ERRH:
13970   Select Case ERR.Number
        Case Else
13980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13990   End Select
14000   Resume EXITP

End Sub

Private Sub cmdResetFilter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_MouseUp"

14110   With Me
14120     .cmdResetFilter_raised_focus_dots_img.Visible = True
14130     .cmdResetFilter_raised_img.Visible = False
14140     .cmdResetFilter_raised_semifocus_dots_img.Visible = False
14150     .cmdResetFilter_raised_focus_img.Visible = False
14160     .cmdResetFilter_sunken_focus_dots_img.Visible = False
14170     .cmdResetFilter_raised_img_dis.Visible = False
14180     blnClearFilter_MouseDown = False
14190   End With

EXITP:
14200   Exit Sub

ERRH:
14210   Select Case ERR.Number
        Case Else
14220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14230   End Select
14240   Resume EXITP

End Sub

Private Sub cmdResetFilter_LostFocus()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdResetFilter_LostFocus"

14310   With Me
14320     .cmdResetFilter_raised_img.Visible = False
14330     .cmdResetFilter_raised_semifocus_dots_img.Visible = False
14340     .cmdResetFilter_raised_focus_img.Visible = False
14350     .cmdResetFilter_raised_focus_dots_img.Visible = False
14360     .cmdResetFilter_sunken_focus_dots_img.Visible = False
14370     .cmdResetFilter_raised_img_dis.Visible = False
14380     blnClearFilter_Focus = False
14390   End With

EXITP:
14400   Exit Sub

ERRH:
14410   Select Case ERR.Number
        Case Else
14420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14430   End Select
14440   Resume EXITP

End Sub

Public Sub opgFund_AfterUpdate()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFund_AfterUpdate"

        Dim strFilter As String

14510   With Me

14520     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
14530       .opgFund_optInclude_lbl.FontBold = True
14540       .opgFund_optExclude_lbl.FontBold = False
14550       .opgFund_optOnly_lbl.FontBold = False
14560       strFilter = vbNullString
14570     Case .opgFund_optExclude.OptionValue
14580       .opgFund_optInclude_lbl.FontBold = False
14590       .opgFund_optExclude_lbl.FontBold = True
14600       .opgFund_optOnly_lbl.FontBold = False
14610       strFilter = "[curr_fund] = false"
14620     Case .opgFund_optOnly.OptionValue
14630       .opgFund_optInclude_lbl.FontBold = False
14640       .opgFund_optExclude_lbl.FontBold = False
14650       .opgFund_optOnly_lbl.FontBold = True
            ' ** Flip the other 4 to Exclude.
14660       If .opgBMU <> .opgBMU_optExclude.OptionValue Then
14670         .opgBMU = .opgBMU_optExclude.OptionValue
14680         .opgBMU_optExclude_lbl.FontBold = True
14690         .opgBMU_optInclude_lbl.FontBold = False
14700         .opgBMU_optOnly_lbl.FontBold = False
14710       End If
14720       If .opgMetal <> .opgMetal_optExclude.OptionValue Then
14730         .opgMetal = .opgMetal_optExclude.OptionValue
14740         .opgMetal_optExclude_lbl.FontBold = True
14750         .opgMetal_optInclude_lbl.FontBold = False
14760         .opgMetal_optOnly_lbl.FontBold = False
14770       End If
14780       If .opgAlt <> .opgAlt_optExclude.OptionValue Then
14790         .opgAlt = .opgAlt_optExclude.OptionValue
14800         .opgAlt_optExclude_lbl.FontBold = True
14810         .opgAlt_optInclude_lbl.FontBold = False
14820         .opgAlt_optOnly_lbl.FontBold = False
14830       End If
14840       If .opgUnit <> .opgUnit_optExclude.OptionValue Then
14850         .opgUnit = .opgUnit_optExclude.OptionValue
14860         .opgUnit_optExclude_lbl.FontBold = True
14870         .opgUnit_optInclude_lbl.FontBold = False
14880         .opgUnit_optOnly_lbl.FontBold = False
14890       End If
14900       strFilter = "[curr_fund] = True"
14910     End Select

14920     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
            ' ** No change.
14930     Case .opgActive_optExclude.OptionValue
14940       If strFilter <> vbNullString Then strFilter = strFilter & FAND
14950       strFilter = strFilter & "[curr_active] = False"
14960     Case .opgActive_optOnly.OptionValue
14970       If strFilter <> vbNullString Then strFilter = strFilter & FAND
14980       strFilter = strFilter & "[curr_active] = True"
14990     End Select

15000     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
            ' ** No change.
15010     Case .opgBMU_optExclude.OptionValue
15020       If strFilter <> vbNullString Then strFilter = strFilter & FAND
15030       strFilter = strFilter & "[curr_bmu] = False"
15040     Case .opgBMU_optOnly.OptionValue
15050       If strFilter <> vbNullString Then strFilter = strFilter & FAND
15060       strFilter = strFilter & "[curr_bmu] = True"
15070     End Select

15080     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
            ' ** No change.
15090     Case .opgMetal_optExclude.OptionValue
15100       If strFilter <> vbNullString Then strFilter = strFilter & FAND
15110       strFilter = strFilter & "[curr_metal] = False"
15120     Case .opgMetal_optOnly.OptionValue
15130       If strFilter <> vbNullString Then strFilter = strFilter & FAND
15140       strFilter = strFilter & "[curr_metal] = True"
15150     End Select

15160     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
            ' ** No change.
15170     Case .opgAlt_optExclude.OptionValue
15180       If strFilter <> vbNullString Then strFilter = strFilter & FAND
15190       strFilter = strFilter & "[curr_alt] = False"
15200     Case .opgAlt_optOnly.OptionValue
15210       If strFilter <> vbNullString Then strFilter = strFilter & FAND
15220       strFilter = strFilter & "[curr_alt] = True"
15230     End Select

15240     Select Case .opgUnit
          Case .opgUnit_optInclude.OptionValue
            ' ** No change.
15250     Case .opgUnit_optExclude.OptionValue
15260       If strFilter <> vbNullString Then strFilter = strFilter & FAND
15270       strFilter = strFilter & "[curr_unit] = False"
15280     Case .opgUnit_optOnly.OptionValue
15290       If strFilter <> vbNullString Then strFilter = strFilter & FAND
15300       strFilter = strFilter & "[curr_unit] = True"
15310     End Select

15320     WhatsVis  ' ** Procedure: Below.

15330     ResetFilter_Set  ' ** Procedure: Below.

15340     .frmCurrency_Sub.Form.Filter = strFilter
15350     If strFilter <> vbNullString Then
15360       .frmCurrency_Sub.Form.FilterOn = True
15370     Else
15380       .frmCurrency_Sub.Form.FilterOn = False
15390     End If

15400     DoEvents
15410     lngRecsCur = .frmCurrency_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Sub.
15420     If lngRecsCur > 0& Then
15430       .frmCurrency_Sub.SetFocus
15440     End If

15450   End With

EXITP:
15460   Exit Sub

ERRH:
15470   Select Case ERR.Number
        Case Else
15480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15490   End Select
15500   Resume EXITP

End Sub

Private Sub opgFund_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFund_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15620   intShiftDown = (Shift And acShiftMask) > 0
15630   intAltDown = (Shift And acAltMask) > 0
15640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15670       With Me
15680         intRetVal = 0
15690         .opgMetal.SetFocus
15700       End With
15710     End Select
15720   End If

        ' ** Shift keys.
15730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15750       With Me
15760         intRetVal = 0
15770         If .cmdResetFilter.Enabled = True Then
15780           .cmdResetFilter.SetFocus
15790         Else
15800           .opgActive.SetFocus
15810         End If
15820       End With
15830     End Select
15840   End If

EXITP:
15850   KeyCode = intRetVal
15860   Exit Sub

ERRH:
15870   intRetVal = 0
15880   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
15890   Case Else
15900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15910   End Select
15920   Resume EXITP

End Sub

Private Sub opgFund_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFund_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16020   intShiftDown = (Shift And acShiftMask) > 0
16030   intAltDown = (Shift And acAltMask) > 0
16040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16070       With Me
16080         intRetVal = 0
16090         .opgMetal.SetFocus
16100       End With
16110     End Select
16120   End If

        ' ** Shift keys.
16130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16150       With Me
16160         intRetVal = 0
16170         If .cmdResetFilter.Enabled = True Then
16180           .cmdResetFilter.SetFocus
16190         Else
16200           .opgActive.SetFocus
16210         End If
16220       End With
16230     End Select
16240   End If

EXITP:
16250   KeyCode = intRetVal
16260   Exit Sub

ERRH:
16270   intRetVal = 0
16280   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
16290   Case Else
16300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16310   End Select
16320   Resume EXITP

End Sub

Private Sub opgFund_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFund_optOnly_KeyDown"

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
16490         .opgMetal.SetFocus
16500       End With
16510     End Select
16520   End If

        ' ** Shift keys.
16530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16550       With Me
16560         intRetVal = 0
16570         If .cmdResetFilter.Enabled = True Then
16580           .cmdResetFilter.SetFocus
16590         Else
16600           .opgActive.SetFocus
16610         End If
16620       End With
16630     End Select
16640   End If

EXITP:
16650   KeyCode = intRetVal
16660   Exit Sub

ERRH:
16670   intRetVal = 0
16680   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
16690   Case Else
16700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16710   End Select
16720   Resume EXITP

End Sub

Public Sub opgMetal_AfterUpdate()

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgMetal_AfterUpdate"

        Dim strFilter As String

16810   With Me

16820     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
16830       .opgMetal_optInclude_lbl.FontBold = True
16840       .opgMetal_optExclude_lbl.FontBold = False
16850       .opgMetal_optOnly_lbl.FontBold = False
16860       strFilter = vbNullString
16870     Case .opgMetal_optExclude.OptionValue
16880       .opgMetal_optInclude_lbl.FontBold = False
16890       .opgMetal_optExclude_lbl.FontBold = True
16900       .opgMetal_optOnly_lbl.FontBold = False
16910       strFilter = "[curr_metal] = False"
16920     Case .opgMetal_optOnly.OptionValue
16930       .opgMetal_optInclude_lbl.FontBold = False
16940       .opgMetal_optExclude_lbl.FontBold = False
16950       .opgMetal_optOnly_lbl.FontBold = True
            ' ** Flip the other 4 to Exclude.
16960       If .opgFund <> .opgFund_optExclude.OptionValue Then
16970         .opgFund = .opgFund_optExclude.OptionValue
16980         .opgFund_optExclude_lbl.FontBold = True
16990         .opgFund_optInclude_lbl.FontBold = False
17000         .opgFund_optOnly_lbl.FontBold = False
17010       End If
17020       If .opgBMU <> .opgBMU_optExclude.OptionValue Then
17030         .opgBMU = .opgBMU_optExclude.OptionValue
17040         .opgBMU_optExclude_lbl.FontBold = True
17050         .opgBMU_optInclude_lbl.FontBold = False
17060         .opgBMU_optOnly_lbl.FontBold = False
17070       End If
17080       If .opgAlt <> .opgAlt_optExclude.OptionValue Then
17090         .opgAlt = .opgAlt_optExclude.OptionValue
17100         .opgAlt_optExclude_lbl.FontBold = True
17110         .opgAlt_optInclude_lbl.FontBold = False
17120         .opgAlt_optOnly_lbl.FontBold = False
17130       End If
17140       If .opgUnit <> .opgUnit_optExclude.OptionValue Then
17150         .opgUnit = .opgUnit_optExclude.OptionValue
17160         .opgUnit_optExclude_lbl.FontBold = True
17170         .opgUnit_optInclude_lbl.FontBold = False
17180         .opgUnit_optOnly_lbl.FontBold = False
17190       End If
17200       strFilter = "[curr_metal] = True"
17210     End Select

17220     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
            ' ** No change.
17230     Case .opgActive_optExclude.OptionValue
17240       If strFilter <> vbNullString Then strFilter = strFilter & FAND
17250       strFilter = strFilter & "[curr_active] = False"
17260     Case .opgActive_optOnly.OptionValue
17270       If strFilter <> vbNullString Then strFilter = strFilter & FAND
17280       strFilter = strFilter & "[curr_active] = True"
17290     End Select

17300     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
            ' ** No change.
17310     Case .opgFund_optExclude.OptionValue
17320       If strFilter <> vbNullString Then strFilter = strFilter & FAND
17330       strFilter = strFilter & "[curr_fund] = False"
17340     Case .opgFund_optOnly.OptionValue
17350       If strFilter <> vbNullString Then strFilter = strFilter & FAND
17360       strFilter = strFilter & "[curr_fund] = True"
17370     End Select

17380     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
            ' ** No change.
17390     Case .opgBMU_optExclude.OptionValue
17400       If strFilter <> vbNullString Then strFilter = strFilter & FAND
17410       strFilter = strFilter & "[curr_bmu] = False"
17420     Case .opgBMU_optOnly.OptionValue
17430       If strFilter <> vbNullString Then strFilter = strFilter & FAND
17440       strFilter = strFilter & "[curr_bmu] = True"
17450     End Select

17460     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
            ' ** No change.
17470     Case .opgAlt_optExclude.OptionValue
17480       If strFilter <> vbNullString Then strFilter = strFilter & FAND
17490       strFilter = strFilter & "[curr_alt] = False"
17500     Case .opgAlt_optOnly.OptionValue
17510       If strFilter <> vbNullString Then strFilter = strFilter & FAND
17520       strFilter = strFilter & "[curr_alt] = True"
17530     End Select

17540     Select Case .opgUnit
          Case .opgUnit_optInclude.OptionValue
            ' ** No change.
17550     Case .opgUnit_optExclude.OptionValue
17560       If strFilter <> vbNullString Then strFilter = strFilter & FAND
17570       strFilter = strFilter & "[curr_unit] = False"
17580     Case .opgUnit_optOnly.OptionValue
17590       If strFilter <> vbNullString Then strFilter = strFilter & FAND
17600       strFilter = strFilter & "[curr_unit] = True"
17610     End Select

17620     WhatsVis  ' ** Procedure: Below.

17630     ResetFilter_Set  ' ** Procedure: Below.

17640     .frmCurrency_Sub.Form.Filter = strFilter
17650     If strFilter <> vbNullString Then
17660       .frmCurrency_Sub.Form.FilterOn = True
17670     Else
17680       .frmCurrency_Sub.Form.FilterOn = False
17690     End If

17700     DoEvents
17710     lngRecsCur = .frmCurrency_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Sub.
17720     If lngRecsCur > 0& Then
17730       .frmCurrency_Sub.SetFocus
17740     End If

17750   End With

EXITP:
17760   Exit Sub

ERRH:
17770   Select Case ERR.Number
        Case Else
17780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17790   End Select
17800   Resume EXITP

End Sub

Private Sub opgMetal_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgMetal_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17920   intShiftDown = (Shift And acShiftMask) > 0
17930   intAltDown = (Shift And acAltMask) > 0
17940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17970       With Me
17980         intRetVal = 0
17990         .opgBMU.SetFocus
18000       End With
18010     End Select
18020   End If

        ' ** Shift keys.
18030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18050       With Me
18060         intRetVal = 0
18070         .opgFund.SetFocus
18080       End With
18090     End Select
18100   End If

EXITP:
18110   KeyCode = intRetVal
18120   Exit Sub

ERRH:
18130   intRetVal = 0
18140   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
18150   Case Else
18160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18170   End Select
18180   Resume EXITP

End Sub

Private Sub opgMetal_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgMetal_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18220   intShiftDown = (Shift And acShiftMask) > 0
18230   intAltDown = (Shift And acAltMask) > 0
18240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18270       With Me
18280         intRetVal = 0
18290         .opgBMU.SetFocus
18300       End With
18310     End Select
18320   End If

        ' ** Shift keys.
18330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18350       With Me
18360         intRetVal = 0
18370         .opgFund.SetFocus
18380       End With
18390     End Select
18400   End If

EXITP:
18410   KeyCode = intRetVal
18420   Exit Sub

ERRH:
18430   intRetVal = 0
18440   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
18450   Case Else
18460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18470   End Select
18480   Resume EXITP

End Sub

Private Sub opgMetal_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgMetal_optOnly_KeyDown"

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
18590         .opgBMU.SetFocus
18600       End With
18610     End Select
18620   End If

        ' ** Shift keys.
18630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18650       With Me
18660         intRetVal = 0
18670         .opgFund.SetFocus
18680       End With
18690     End Select
18700   End If

EXITP:
18710   KeyCode = intRetVal
18720   Exit Sub

ERRH:
18730   intRetVal = 0
18740   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
18750   Case Else
18760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18770   End Select
18780   Resume EXITP

End Sub

Public Sub opgBMU_AfterUpdate()

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgBMU_AfterUpdate"

        Dim strFilter As String

18810   With Me

18820     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
18830       .opgBMU_optInclude_lbl.FontBold = True
18840       .opgBMU_optExclude_lbl.FontBold = False
18850       .opgBMU_optOnly_lbl.FontBold = False
18860       strFilter = vbNullString
18870     Case .opgBMU_optExclude.OptionValue
18880       .opgBMU_optInclude_lbl.FontBold = False
18890       .opgBMU_optExclude_lbl.FontBold = True
18900       .opgBMU_optOnly_lbl.FontBold = False
18910       strFilter = "[curr_bmu] = False"
18920     Case .opgBMU_optOnly.OptionValue
18930       .opgBMU_optInclude_lbl.FontBold = False
18940       .opgBMU_optExclude_lbl.FontBold = False
18950       .opgBMU_optOnly_lbl.FontBold = True
            ' ** Flip the other 4 to Exclude.
18960       If .opgFund <> .opgFund_optExclude.OptionValue Then
18970         .opgFund = .opgFund_optExclude.OptionValue
18980         .opgFund_optExclude_lbl.FontBold = True
18990         .opgFund_optInclude_lbl.FontBold = False
19000         .opgFund_optOnly_lbl.FontBold = False
19010       End If
19020       If .opgMetal <> .opgMetal_optExclude.OptionValue Then
19030         .opgMetal = .opgMetal_optExclude.OptionValue
19040         .opgMetal_optExclude_lbl.FontBold = True
19050         .opgMetal_optInclude_lbl.FontBold = False
19060         .opgMetal_optOnly_lbl.FontBold = False
19070       End If
19080       If .opgAlt <> .opgAlt_optExclude.OptionValue Then
19090         .opgAlt = .opgAlt_optExclude.OptionValue
19100         .opgAlt_optExclude_lbl.FontBold = True
19110         .opgAlt_optInclude_lbl.FontBold = False
19120         .opgAlt_optOnly_lbl.FontBold = False
19130       End If
19140       If .opgUnit <> .opgUnit_optExclude.OptionValue Then
19150         .opgUnit = .opgUnit_optExclude.OptionValue
19160         .opgUnit_optExclude_lbl.FontBold = True
19170         .opgUnit_optInclude_lbl.FontBold = False
19180         .opgUnit_optOnly_lbl.FontBold = False
19190       End If
19200       strFilter = "[curr_bmu] = True"
19210     End Select

19220     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
            ' ** No change.
19230     Case .opgActive_optExclude.OptionValue
19240       If strFilter <> vbNullString Then strFilter = strFilter & FAND
19250       strFilter = strFilter & "[curr_active] = False"
19260     Case .opgActive_optOnly.OptionValue
19270       If strFilter <> vbNullString Then strFilter = strFilter & FAND
19280       strFilter = strFilter & "[curr_active] = True"
19290     End Select

19300     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
            ' ** No change.
19310     Case .opgFund_optExclude.OptionValue
19320       If strFilter <> vbNullString Then strFilter = strFilter & FAND
19330       strFilter = strFilter & "[curr_fund] = False"
19340     Case .opgFund_optOnly.OptionValue
19350       If strFilter <> vbNullString Then strFilter = strFilter & FAND
19360       strFilter = strFilter & "[curr_fund] = True"
19370     End Select

19380     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
            ' ** No change.
19390     Case .opgMetal_optExclude.OptionValue
19400       If strFilter <> vbNullString Then strFilter = strFilter & FAND
19410       strFilter = strFilter & "[curr_metal] = False"
19420     Case .opgMetal_optOnly.OptionValue
19430       If strFilter <> vbNullString Then strFilter = strFilter & FAND
19440       strFilter = strFilter & "[curr_metal] = True"
19450     End Select

19460     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
            ' ** No change.
19470     Case .opgAlt_optExclude.OptionValue
19480       If strFilter <> vbNullString Then strFilter = strFilter & FAND
19490       strFilter = strFilter & "[curr_alt] = False"
19500     Case .opgAlt_optOnly.OptionValue
19510       If strFilter <> vbNullString Then strFilter = strFilter & FAND
19520       strFilter = strFilter & "[curr_alt] = True"
19530     End Select

19540     Select Case .opgUnit
          Case .opgUnit_optInclude.OptionValue
            ' ** No change.
19550     Case .opgUnit_optExclude.OptionValue
19560       If strFilter <> vbNullString Then strFilter = strFilter & FAND
19570       strFilter = strFilter & "[curr_unit] = False"
19580     Case .opgUnit_optOnly.OptionValue
19590       If strFilter <> vbNullString Then strFilter = strFilter & FAND
19600       strFilter = strFilter & "[curr_unit] = True"
19610     End Select

19620     WhatsVis  ' ** Procedure: Below.

19630     ResetFilter_Set  ' ** Procedure: Below.

19640     .frmCurrency_Sub.Form.Filter = strFilter
19650     If strFilter <> vbNullString Then
19660       .frmCurrency_Sub.Form.FilterOn = True
19670     Else
19680       .frmCurrency_Sub.Form.FilterOn = False
19690     End If

19700     DoEvents
19710     lngRecsCur = .frmCurrency_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Sub.
19720     If lngRecsCur > 0& Then
19730       .frmCurrency_Sub.SetFocus
19740     End If

19750   End With

EXITP:
19760   Exit Sub

ERRH:
19770   Select Case ERR.Number
        Case Else
19780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19790   End Select
19800   Resume EXITP

End Sub

Private Sub opgBMU_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgBMU_optInclude_KeyDown"

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
19990         .opgAlt.SetFocus
20000       End With
20010     End Select
20020   End If

        ' ** Shift keys.
20030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20050       With Me
20060         intRetVal = 0
20070         .opgMetal.SetFocus
20080       End With
20090     End Select
20100   End If

EXITP:
20110   KeyCode = intRetVal
20120   Exit Sub

ERRH:
20130   intRetVal = 0
20140   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
20150   Case Else
20160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20170   End Select
20180   Resume EXITP

End Sub

Private Sub opgBMU_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgBMU_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20220   intShiftDown = (Shift And acShiftMask) > 0
20230   intAltDown = (Shift And acAltMask) > 0
20240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20270       With Me
20280         intRetVal = 0
20290         .opgAlt.SetFocus
20300       End With
20310     End Select
20320   End If

        ' ** Shift keys.
20330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20350       With Me
20360         intRetVal = 0
20370         .opgMetal.SetFocus
20380       End With
20390     End Select
20400   End If

EXITP:
20410   KeyCode = intRetVal
20420   Exit Sub

ERRH:
20430   intRetVal = 0
20440   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
20450   Case Else
20460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20470   End Select
20480   Resume EXITP

End Sub

Private Sub opgBMU_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgBMU_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20520   intShiftDown = (Shift And acShiftMask) > 0
20530   intAltDown = (Shift And acAltMask) > 0
20540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20570       With Me
20580         intRetVal = 0
20590         .opgAlt.SetFocus
20600       End With
20610     End Select
20620   End If

        ' ** Shift keys.
20630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20650       With Me
20660         intRetVal = 0
20670         .opgMetal.SetFocus
20680       End With
20690     End Select
20700   End If

EXITP:
20710   KeyCode = intRetVal
20720   Exit Sub

ERRH:
20730   intRetVal = 0
20740   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
20750   Case Else
20760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20770   End Select
20780   Resume EXITP

End Sub

Public Sub opgAlt_AfterUpdate()

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAlt_AfterUpdate"

        Dim strFilter As String

20810   With Me

20820     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
20830       .opgAlt_optInclude_lbl.FontBold = True
20840       .opgAlt_optExclude_lbl.FontBold = False
20850       .opgAlt_optOnly_lbl.FontBold = False
20860       strFilter = vbNullString
20870     Case .opgAlt_optExclude.OptionValue
20880       .opgAlt_optInclude_lbl.FontBold = False
20890       .opgAlt_optExclude_lbl.FontBold = True
20900       .opgAlt_optOnly_lbl.FontBold = False
20910       strFilter = "[curr_alt] = False"
20920     Case .opgAlt_optOnly.OptionValue
20930       .opgAlt_optInclude_lbl.FontBold = False
20940       .opgAlt_optExclude_lbl.FontBold = False
20950       .opgAlt_optOnly_lbl.FontBold = True
            ' ** Flip the other 4 to Exclude.
20960       If .opgFund <> .opgFund_optExclude.OptionValue Then
20970         .opgFund = .opgFund_optExclude.OptionValue
20980         .opgFund_optExclude_lbl.FontBold = True
20990         .opgFund_optInclude_lbl.FontBold = False
21000         .opgFund_optOnly_lbl.FontBold = False
21010       End If
21020       If .opgBMU <> .opgBMU_optExclude.OptionValue Then
21030         .opgBMU = .opgBMU_optExclude.OptionValue
21040         .opgBMU_optExclude_lbl.FontBold = True
21050         .opgBMU_optInclude_lbl.FontBold = False
21060         .opgBMU_optOnly_lbl.FontBold = False
21070       End If
21080       If .opgMetal <> .opgMetal_optExclude.OptionValue Then
21090         .opgMetal = .opgMetal_optExclude.OptionValue
21100         .opgMetal_optExclude_lbl.FontBold = True
21110         .opgMetal_optInclude_lbl.FontBold = False
21120         .opgMetal_optOnly_lbl.FontBold = False
21130       End If
21140       If .opgUnit <> .opgUnit_optExclude.OptionValue Then
21150         .opgUnit = .opgUnit_optExclude.OptionValue
21160         .opgUnit_optExclude_lbl.FontBold = True
21170         .opgUnit_optInclude_lbl.FontBold = False
21180         .opgUnit_optOnly_lbl.FontBold = False
21190       End If
21200       strFilter = "[curr_alt] = True"
21210     End Select

21220     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
            ' ** No change.
21230     Case .opgActive_optExclude.OptionValue
21240       If strFilter <> vbNullString Then strFilter = strFilter & FAND
21250       strFilter = strFilter & "[curr_active] = False"
21260     Case .opgActive_optOnly.OptionValue
21270       If strFilter <> vbNullString Then strFilter = strFilter & FAND
21280       strFilter = strFilter & "[curr_active] = True"
21290     End Select

21300     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
            ' ** No change.
21310     Case .opgFund_optExclude.OptionValue
21320       If strFilter <> vbNullString Then strFilter = strFilter & FAND
21330       strFilter = strFilter & "[curr_fund] = False"
21340     Case .opgFund_optOnly.OptionValue
21350       If strFilter <> vbNullString Then strFilter = strFilter & FAND
21360       strFilter = strFilter & "[curr_fund] = True"
21370     End Select

21380     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
            ' ** No change.
21390     Case .opgBMU_optExclude.OptionValue
21400       If strFilter <> vbNullString Then strFilter = strFilter & FAND
21410       strFilter = strFilter & "[curr_bmu] = False"
21420     Case .opgBMU_optOnly.OptionValue
21430       If strFilter <> vbNullString Then strFilter = strFilter & FAND
21440       strFilter = strFilter & "[curr_bmu] = True"
21450     End Select

21460     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
            ' ** No change.
21470     Case .opgMetal_optExclude.OptionValue
21480       If strFilter <> vbNullString Then strFilter = strFilter & FAND
21490       strFilter = strFilter & "[curr_metal] = False"
21500     Case .opgMetal_optOnly.OptionValue
21510       If strFilter <> vbNullString Then strFilter = strFilter & FAND
21520       strFilter = strFilter & "[curr_metal] = True"
21530     End Select

21540     Select Case .opgUnit
          Case .opgUnit_optInclude.OptionValue
            ' ** No change.
21550     Case .opgUnit_optExclude.OptionValue
21560       If strFilter <> vbNullString Then strFilter = strFilter & FAND
21570       strFilter = strFilter & "[curr_unit] = False"
21580     Case .opgUnit_optOnly.OptionValue
21590       If strFilter <> vbNullString Then strFilter = strFilter & FAND
21600       strFilter = strFilter & "[curr_unit] = True"
21610     End Select

21620     WhatsVis  ' ** Procedure: Below.

21630     ResetFilter_Set  ' ** Procedure: Below.

21640     .frmCurrency_Sub.Form.Filter = strFilter
21650     If strFilter <> vbNullString Then
21660       .frmCurrency_Sub.Form.FilterOn = True
21670     Else
21680       .frmCurrency_Sub.Form.FilterOn = False
21690     End If

21700     DoEvents
21710     lngRecsCur = .frmCurrency_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Sub.
21720     If lngRecsCur > 0& Then
21730       .frmCurrency_Sub.SetFocus
21740     End If

21750   End With

EXITP:
21760   Exit Sub

ERRH:
21770   Select Case ERR.Number
        Case Else
21780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21790   End Select
21800   Resume EXITP

End Sub

Private Sub opgAlt_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAlt_optInclude_KeyDown"

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
21990         .opgUnit.SetFocus
22000       End With
22010     End Select
22020   End If

        ' ** Shift keys.
22030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22050       With Me
22060         intRetVal = 0
22070         .opgBMU.SetFocus
22080       End With
22090     End Select
22100   End If

EXITP:
22110   KeyCode = intRetVal
22120   Exit Sub

ERRH:
22130   intRetVal = 0
22140   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
22150   Case Else
22160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22170   End Select
22180   Resume EXITP

End Sub

Private Sub opgAlt_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAlt_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22220   intShiftDown = (Shift And acShiftMask) > 0
22230   intAltDown = (Shift And acAltMask) > 0
22240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22270       With Me
22280         intRetVal = 0
22290         .opgUnit.SetFocus
22300       End With
22310     End Select
22320   End If

        ' ** Shift keys.
22330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22350       With Me
22360         intRetVal = 0
22370         .opgBMU.SetFocus
22380       End With
22390     End Select
22400   End If

EXITP:
22410   KeyCode = intRetVal
22420   Exit Sub

ERRH:
22430   intRetVal = 0
22440   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
22450   Case Else
22460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22470   End Select
22480   Resume EXITP

End Sub

Private Sub opgAlt_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAlt_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22520   intShiftDown = (Shift And acShiftMask) > 0
22530   intAltDown = (Shift And acAltMask) > 0
22540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22570       With Me
22580         intRetVal = 0
22590         .opgUnit.SetFocus
22600       End With
22610     End Select
22620   End If

        ' ** Shift keys.
22630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22650       With Me
22660         intRetVal = 0
22670         .opgBMU.SetFocus
22680       End With
22690     End Select
22700   End If

EXITP:
22710   KeyCode = intRetVal
22720   Exit Sub

ERRH:
22730   intRetVal = 0
22740   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
22750   Case Else
22760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22770   End Select
22780   Resume EXITP

End Sub

Public Sub opgUnit_AfterUpdate()

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnit_AfterUpdate"

        Dim strFilter As String

22810   With Me

22820     Select Case .opgUnit
          Case .opgUnit_optInclude.OptionValue
22830       .opgUnit_optInclude_lbl.FontBold = True
22840       .opgUnit_optExclude_lbl.FontBold = False
22850       .opgUnit_optOnly_lbl.FontBold = False
22860       strFilter = vbNullString
22870     Case .opgUnit_optExclude.OptionValue
22880       .opgUnit_optInclude_lbl.FontBold = False
22890       .opgUnit_optExclude_lbl.FontBold = True
22900       .opgUnit_optOnly_lbl.FontBold = False
22910       strFilter = "[curr_unit] = False"
22920     Case .opgUnit_optOnly.OptionValue
22930       .opgUnit_optInclude_lbl.FontBold = False
22940       .opgUnit_optExclude_lbl.FontBold = False
22950       .opgUnit_optOnly_lbl.FontBold = True
            ' ** Flip the other 4 to Exclude.
22960       If .opgFund <> .opgFund_optExclude.OptionValue Then
22970         .opgFund = .opgFund_optExclude.OptionValue
22980         .opgFund_optExclude_lbl.FontBold = True
22990         .opgFund_optInclude_lbl.FontBold = False
23000         .opgFund_optOnly_lbl.FontBold = False
23010       End If
23020       If .opgBMU <> .opgBMU_optExclude.OptionValue Then
23030         .opgBMU = .opgBMU_optExclude.OptionValue
23040         .opgBMU_optExclude_lbl.FontBold = True
23050         .opgBMU_optInclude_lbl.FontBold = False
23060         .opgBMU_optOnly_lbl.FontBold = False
23070       End If
23080       If .opgMetal <> .opgMetal_optExclude.OptionValue Then
23090         .opgMetal = .opgMetal_optExclude.OptionValue
23100         .opgMetal_optExclude_lbl.FontBold = True
23110         .opgMetal_optInclude_lbl.FontBold = False
23120         .opgMetal_optOnly_lbl.FontBold = False
23130       End If
23140       If .opgAlt <> .opgAlt_optExclude.OptionValue Then
23150         .opgAlt = .opgAlt_optExclude.OptionValue
23160         .opgAlt_optExclude_lbl.FontBold = True
23170         .opgAlt_optInclude_lbl.FontBold = False
23180         .opgAlt_optOnly_lbl.FontBold = False
23190       End If
23200       strFilter = "[curr_unit] = True"
23210     End Select

23220     Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
            ' ** No change.
23230     Case .opgActive_optExclude.OptionValue
23240       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23250       strFilter = strFilter & "[curr_active] = False"
23260     Case .opgActive_optOnly.OptionValue
23270       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23280       strFilter = strFilter & "[curr_active] = True"
23290     End Select

23300     Select Case .opgFund
          Case .opgFund_optInclude.OptionValue
            ' ** No change.
23310     Case .opgFund_optExclude.OptionValue
23320       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23330       strFilter = strFilter & "[curr_fund] = False"
23340     Case .opgFund_optOnly.OptionValue
23350       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23360       strFilter = strFilter & "[curr_fund] = True"
23370     End Select

23380     Select Case .opgBMU
          Case .opgBMU_optInclude.OptionValue
            ' ** No change.
23390     Case .opgBMU_optExclude.OptionValue
23400       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23410       strFilter = strFilter & "[curr_bmu] = False"
23420     Case .opgBMU_optOnly.OptionValue
23430       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23440       strFilter = strFilter & "[curr_bmu] = True"
23450     End Select

23460     Select Case .opgMetal
          Case .opgMetal_optInclude.OptionValue
            ' ** No change.
23470     Case .opgMetal_optExclude.OptionValue
23480       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23490       strFilter = strFilter & "[curr_metal] = False"
23500     Case .opgMetal_optOnly.OptionValue
23510       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23520       strFilter = strFilter & "[curr_metal] = True"
23530     End Select

23540     Select Case .opgAlt
          Case .opgAlt_optInclude.OptionValue
            ' ** No change.
23550     Case .opgAlt_optExclude.OptionValue
23560       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23570       strFilter = strFilter & "[curr_alt] = False"
23580     Case .opgAlt_optOnly.OptionValue
23590       If strFilter <> vbNullString Then strFilter = strFilter & FAND
23600       strFilter = strFilter & "[curr_alt] = True"
23610     End Select

23620     WhatsVis  ' ** Procedure: Below.

23630     ResetFilter_Set  ' ** Procedure: Below.

23640     .frmCurrency_Sub.Form.Filter = strFilter
23650     If strFilter <> vbNullString Then
23660       .frmCurrency_Sub.Form.FilterOn = True
23670     Else
23680       .frmCurrency_Sub.Form.FilterOn = False
23690     End If

23700     DoEvents
23710     lngRecsCur = .frmCurrency_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Sub.
23720     If lngRecsCur > 0& Then
23730       .frmCurrency_Sub.SetFocus
23740     End If

23750   End With

EXITP:
23760   Exit Sub

ERRH:
23770   Select Case ERR.Number
        Case Else
23780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23790   End Select
23800   Resume EXITP

End Sub

Private Sub opgUnit_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

23900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnit_optInclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23920   intShiftDown = (Shift And acShiftMask) > 0
23930   intAltDown = (Shift And acAltMask) > 0
23940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23970       With Me
23980         intRetVal = 0
23990         If .cmdCountries.Enabled = True Then
24000           .cmdCountries.SetFocus
24010         ElseIf .cmdPreview.Enabled = True Then
24020           .cmdPreview.SetFocus
24030         ElseIf .cmdEdit.Enabled = True Then
24040           .cmdEdit.SetFocus
24050         ElseIf .cmdUpdate.Enabled = True Then
24060           .cmdUpdate.SetFocus
24070         Else
24080           .cmdClose.SetFocus
24090         End If
24100       End With
24110     End Select
24120   End If

        ' ** Shift keys.
24130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24150       With Me
24160         intRetVal = 0
24170         .opgAlt.SetFocus
24180       End With
24190     End Select
24200   End If

EXITP:
24210   KeyCode = intRetVal
24220   Exit Sub

ERRH:
24230   intRetVal = 0
24240   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
24250   Case Else
24260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24270   End Select
24280   Resume EXITP

End Sub

Private Sub opgUnit_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnit_optExclude_KeyDown"

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
24390         If .cmdCountries.Enabled = True Then
24400           .cmdCountries.SetFocus
24410         ElseIf .cmdPreview.Enabled = True Then
24420           .cmdPreview.SetFocus
24430         ElseIf .cmdEdit.Enabled = True Then
24440           .cmdEdit.SetFocus
24450         ElseIf .cmdUpdate.Enabled = True Then
24460           .cmdUpdate.SetFocus
24470         Else
24480           .cmdClose.SetFocus
24490         End If
24500       End With
24510     End Select
24520   End If

        ' ** Shift keys.
24530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24550       With Me
24560         intRetVal = 0
24570         .opgAlt.SetFocus
24580       End With
24590     End Select
24600   End If

EXITP:
24610   KeyCode = intRetVal
24620   Exit Sub

ERRH:
24630   intRetVal = 0
24640   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
24650   Case Else
24660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24670   End Select
24680   Resume EXITP

End Sub

Private Sub opgUnit_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgUnit_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24720   intShiftDown = (Shift And acShiftMask) > 0
24730   intAltDown = (Shift And acAltMask) > 0
24740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24770       With Me
24780         intRetVal = 0
24790         If .cmdCountries.Enabled = True Then
24800           .cmdCountries.SetFocus
24810         ElseIf .cmdPreview.Enabled = True Then
24820           .cmdPreview.SetFocus
24830         ElseIf .cmdEdit.Enabled = True Then
24840           .cmdEdit.SetFocus
24850         ElseIf .cmdUpdate.Enabled = True Then
24860           .cmdUpdate.SetFocus
24870         Else
24880           .cmdClose.SetFocus
24890         End If
24900       End With
24910     End Select
24920   End If

        ' ** Shift keys.
24930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24950       With Me
24960         intRetVal = 0
24970         .opgAlt.SetFocus
24980       End With
24990     End Select
25000   End If

EXITP:
25010   KeyCode = intRetVal
25020   Exit Sub

ERRH:
25030   intRetVal = 0
25040   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
25050   Case Else
25060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25070   End Select
25080   Resume EXITP

End Sub

Private Sub country_cnt_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "country_cnt_box_MouseMove"

25110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
25120   Exit Sub

ERRH:
25130   Select Case ERR.Number
        Case Else
25140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25150   End Select
25160   Resume EXITP

End Sub

Private Sub country_cnt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "country_cnt_MouseMove"

25210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
25220   Exit Sub

ERRH:
25230   Select Case ERR.Number
        Case Else
25240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25250   End Select
25260   Resume EXITP

End Sub

Private Sub cmdCountries_GotFocus()

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_GotFocus"

25310   With Me
25320     blnCountries_Focus = True
25330     .cmdCountries_raised_semifocus_dots_img.Visible = True
25340     .cmdCountries_raised_img.Visible = False
25350     .cmdCountries_raised_focus_img.Visible = False
25360     .cmdCountries_raised_focus_dots_img.Visible = False
25370     .cmdCountries_sunken_focus_dots_img.Visible = False
25380     .cmdCountries_raised_img_dis.Visible = False
25390   End With

EXITP:
25400   Exit Sub

ERRH:
25410   Select Case ERR.Number
        Case Else
25420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25430   End Select
25440   Resume EXITP

End Sub

Private Sub cmdCountries_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_MouseDown"

25510   With Me
25520     blnCountries_MouseDown = True
25530     .cmdCountries_sunken_focus_dots_img.Visible = True
25540     .cmdCountries_raised_img.Visible = False
25550     .cmdCountries_raised_semifocus_dots_img.Visible = False
25560     .cmdCountries_raised_focus_img.Visible = False
25570     .cmdCountries_raised_focus_dots_img.Visible = False
25580     .cmdCountries_raised_img_dis.Visible = False
25590   End With

EXITP:
25600   Exit Sub

ERRH:
25610   Select Case ERR.Number
        Case Else
25620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25630   End Select
25640   Resume EXITP

End Sub

Private Sub cmdCountries_Click()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_Click"

        Dim strDocName As String

25710   strDocName = "frmCurrency_Country"
25720   Select Case IsLoaded(strDocName, acForm)
        Case True
25730     DoCmd.Close acForm, strDocName
25740   Case False
25750     gblnSetFocus = True
25760     DoCmd.OpenForm strDocName, , , , , , THIS_NAME
25770   End Select

EXITP:
25780   Exit Sub

ERRH:
25790   Select Case ERR.Number
        Case Else
25800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25810   End Select
25820   Resume EXITP

End Sub

Private Sub cmdCountries_KeyDown(KeyCode As Integer, Shift As Integer)

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25920   intShiftDown = (Shift And acShiftMask) > 0
25930   intAltDown = (Shift And acAltMask) > 0
25940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25960     Select Case intRetVal
          Case vbKeyTab
25970       With Me
25980         intRetVal = 0
25990         If .cmdPreview.Enabled = True Then
26000           .cmdPreview.SetFocus
26010         ElseIf .cmdEdit.Enabled = True Then
26020           .cmdEdit.SetFocus
26030         ElseIf .cmdUpdate.Enabled = True Then
26040           .cmdUpdate.SetFocus
26050         Else
26060           .cmdClose.SetFocus
26070         End If
26080       End With
26090     End Select
26100   End If

        ' ** Shift keys.
26110   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26120     Select Case intRetVal
          Case vbKeyTab
26130       With Me
26140         intRetVal = 0
26150         .opgUnit.SetFocus
26160       End With
26170     End Select
26180   End If

        ' ** Ctrl keys.
26190   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
26200     Select Case intRetVal
          Case vbKeyTab
26210       With Me
26220         intRetVal = 0
26230         .chkIncludeNotes.SetFocus
26240       End With
26250     End Select
26260   End If

EXITP:
26270   KeyCode = intRetVal
26280   Exit Sub

ERRH:
26290   intRetVal = 0
26300   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
26310   Case Else
26320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26330   End Select
26340   Resume EXITP

End Sub

Private Sub cmdCountries_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_MouseMove"

26410   With Me
26420     If blnCountries_MouseDown = False Then
26430       Select Case blnCountries_Focus
            Case True
26440         .cmdCountries_raised_focus_dots_img.Visible = True
26450         .cmdCountries_raised_focus_img.Visible = False
26460       Case False
26470         .cmdCountries_raised_focus_img.Visible = True
26480         .cmdCountries_raised_focus_dots_img.Visible = False
26490       End Select
26500       .cmdCountries_raised_img.Visible = False
26510       .cmdCountries_raised_semifocus_dots_img.Visible = False
26520       .cmdCountries_sunken_focus_dots_img.Visible = False
26530       .cmdCountries_raised_img_dis.Visible = False
26540     End If
26550   End With

EXITP:
26560   Exit Sub

ERRH:
26570   Select Case ERR.Number
        Case Else
26580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26590   End Select
26600   Resume EXITP

End Sub

Private Sub cmdCountries_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_MouseUp"

26710   With Me
26720     .cmdCountries_raised_focus_dots_img.Visible = True
26730     .cmdCountries_raised_img.Visible = False
26740     .cmdCountries_raised_semifocus_dots_img.Visible = False
26750     .cmdCountries_raised_focus_img.Visible = False
26760     .cmdCountries_sunken_focus_dots_img.Visible = False
26770     .cmdCountries_raised_img_dis.Visible = False
26780     blnCountries_MouseDown = False
26790   End With

EXITP:
26800   Exit Sub

ERRH:
26810   Select Case ERR.Number
        Case Else
26820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26830   End Select
26840   Resume EXITP

End Sub

Private Sub cmdCountries_LostFocus()

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCountries_LostFocus"

26910   With Me
26920     .cmdCountries_raised_img.Visible = True
26930     .cmdCountries_raised_semifocus_dots_img.Visible = False
26940     .cmdCountries_raised_focus_img.Visible = False
26950     .cmdCountries_raised_focus_dots_img.Visible = False
26960     .cmdCountries_sunken_focus_dots_img.Visible = False
26970     .cmdCountries_raised_img_dis.Visible = False
26980     blnCountries_Focus = False
26990   End With

EXITP:
27000   Exit Sub

ERRH:
27010   Select Case ERR.Number
        Case Else
27020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27030   End Select
27040   Resume EXITP

End Sub

Private Sub chkIncludeNotes_AfterUpdate()

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeNotes_AfterUpdate"

27110   With Me
27120     Select Case .chkIncludeNotes
          Case True
27130       .chkIncludeNotes_lbl.FontBold = True
27140     Case False
27150       .chkIncludeNotes_lbl.FontBold = False
27160     End Select
27170   End With

EXITP:
27180   Exit Sub

ERRH:
27190   Select Case ERR.Number
        Case Else
27200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27210   End Select
27220   Resume EXITP

End Sub

Private Sub chkIncludeNotes_KeyDown(KeyCode As Integer, Shift As Integer)

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeNotes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27320   intShiftDown = (Shift And acShiftMask) > 0
27330   intAltDown = (Shift And acAltMask) > 0
27340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27360     Select Case intRetVal
          Case vbKeyTab
27370       With Me
27380         intRetVal = 0
27390         .chkPageOf.SetFocus
27400       End With
27410     End Select
27420   End If

        ' ** Shift keys.
27430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27440     Select Case intRetVal
          Case vbKeyTab
27450       With Me
27460         intRetVal = 0
27470         .cmdCountries.SetFocus
27480       End With
27490     End Select
27500   End If

EXITP:
27510   KeyCode = intRetVal
27520   Exit Sub

ERRH:
27530   intRetVal = 0
27540   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
27550   Case Else
27560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27570   End Select
27580   Resume EXITP

End Sub

Private Sub chkIncludeNotes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeNotes_MouseMove"

27610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
27620   Exit Sub

ERRH:
27630   Select Case ERR.Number
        Case Else
27640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27650   End Select
27660   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

27710   With Me
27720     Select Case .chkPageOf
          Case True
27730       .chkPageOf_lbl.FontBold = True
27740     Case False
27750       .chkPageOf_lbl.FontBold = False
27760     End Select
27770   End With

EXITP:
27780   Exit Sub

ERRH:
27790   Select Case ERR.Number
        Case Else
27800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27810   End Select
27820   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27920   intShiftDown = (Shift And acShiftMask) > 0
27930   intAltDown = (Shift And acAltMask) > 0
27940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27960     Select Case intRetVal
          Case vbKeyTab
27970       With Me
27980         intRetVal = 0
27990         If .cmdPreview.Enabled = True Then
28000           .cmdPreview.SetFocus
28010         ElseIf .cmdEdit.Enabled = True Then
28020           .cmdEdit.SetFocus
28030         ElseIf .cmdUpdate.Enabled = True Then
28040           .cmdUpdate.SetFocus
28050         Else
28060           .cmdClose.SetFocus
28070         End If
28080       End With
28090     End Select
28100   End If

        ' ** Shift keys.
28110   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28120     Select Case intRetVal
          Case vbKeyTab
28130       With Me
28140         intRetVal = 0
28150         .chkIncludeNotes.SetFocus
28160       End With
28170     End Select
28180   End If

EXITP:
28190   KeyCode = intRetVal
28200   Exit Sub

ERRH:
28210   intRetVal = 0
28220   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
28230   Case Else
28240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28250   End Select
28260   Resume EXITP

End Sub

Private Sub chkPageOf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_MouseMove"

28310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
28320   Exit Sub

ERRH:
28330   Select Case ERR.Number
        Case Else
28340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28350   End Select
28360   Resume EXITP

End Sub

Private Sub ResetFilter_Set()

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "ResetFilter_Set"

28410   With Me

28420     If .opgActive <> .opgActive_optInclude.OptionValue Or _
              .opgFund <> .opgFund_optInclude.OptionValue Or _
              .opgMetal <> .opgMetal_optInclude.OptionValue Or _
              .opgBMU <> .opgBMU_optInclude.OptionValue Or _
              .opgAlt <> .opgAlt_optInclude.OptionValue Or _
              .opgUnit <> .opgUnit_optInclude.OptionValue Then
28430       .cmdResetFilter.Enabled = True
28440       .cmdResetFilter_raised_img.Visible = True
28450       .cmdResetFilter_raised_semifocus_dots_img.Visible = False
28460       .cmdResetFilter_raised_focus_img.Visible = False
28470       .cmdResetFilter_raised_focus_dots_img.Visible = False
28480       .cmdResetFilter_sunken_focus_dots_img.Visible = False
28490       .cmdResetFilter_raised_img_dis.Visible = False
28500     Else
28510       If .cmdEdit.Enabled = True Then
28520         .cmdEdit.SetFocus
28530       ElseIf .cmdUpdate.Enabled = True Then
28540         .cmdUpdate.SetFocus
28550       Else
28560         .frmCurrency_Sub.SetFocus
28570       End If
28580       DoEvents
28590       .cmdResetFilter.Enabled = False
28600       .cmdResetFilter_raised_img_dis.Visible = True
28610       .cmdResetFilter_raised_img.Visible = False
28620       .cmdResetFilter_raised_semifocus_dots_img.Visible = False
28630       .cmdResetFilter_raised_focus_img.Visible = False
28640       .cmdResetFilter_raised_focus_dots_img.Visible = False
28650       .cmdResetFilter_sunken_focus_dots_img.Visible = False
28660     End If

28670   End With

EXITP:
28680   Exit Sub

ERRH:
28690   Select Case ERR.Number
        Case Else
28700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28710   End Select
28720   Resume EXITP

End Sub

Private Sub WhatsVis()

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "WhatsVis"

28810   With Me
28820     If .opgFund = .opgFund_optOnly.OptionValue Then
28830       .frmCurrency_Sub.Form.curr_fund.Visible = True
28840       .frmCurrency_Sub.Form.curr_fund_lbl_off.Visible = True
28850       .frmCurrency_Sub.Form.curr_fund_lbl_on.Visible = True
28860       .frmCurrency_Sub.Form.curr_bmu.Visible = False
28870       .frmCurrency_Sub.Form.curr_bmu_lbl_off.Visible = False
28880       .frmCurrency_Sub.Form.curr_bmu_lbl_on.Visible = False
28890       .frmCurrency_Sub.Form.curr_metal.Visible = False
28900       .frmCurrency_Sub.Form.curr_metal_lbl_off.Visible = False
28910       .frmCurrency_Sub.Form.curr_metal_lbl_on.Visible = False
28920       .frmCurrency_Sub.Form.curr_alt.Visible = False
28930       .frmCurrency_Sub.Form.curr_alt_lbl_off.Visible = False
28940       .frmCurrency_Sub.Form.curr_alt_lbl_on.Visible = False
28950       .frmCurrency_Sub.Form.curr_unit.Visible = False
28960       .frmCurrency_Sub.Form.curr_unit_lbl_off.Visible = False
28970       .frmCurrency_Sub.Form.curr_unit_lbl_on.Visible = False
28980     ElseIf .opgBMU = .opgBMU_optOnly.OptionValue Then
28990       .frmCurrency_Sub.Form.curr_bmu.Visible = True
29000       .frmCurrency_Sub.Form.curr_bmu_lbl_off.Visible = True
29010       .frmCurrency_Sub.Form.curr_bmu_lbl_on.Visible = True
29020       .frmCurrency_Sub.Form.curr_fund.Visible = False
29030       .frmCurrency_Sub.Form.curr_fund_lbl_off.Visible = False
29040       .frmCurrency_Sub.Form.curr_fund_lbl_on.Visible = False
29050       .frmCurrency_Sub.Form.curr_metal.Visible = False
29060       .frmCurrency_Sub.Form.curr_metal_lbl_off.Visible = False
29070       .frmCurrency_Sub.Form.curr_metal_lbl_on.Visible = False
29080       .frmCurrency_Sub.Form.curr_alt.Visible = False
29090       .frmCurrency_Sub.Form.curr_alt_lbl_off.Visible = False
29100       .frmCurrency_Sub.Form.curr_alt_lbl_on.Visible = False
29110       .frmCurrency_Sub.Form.curr_unit.Visible = False
29120       .frmCurrency_Sub.Form.curr_unit_lbl_off.Visible = False
29130       .frmCurrency_Sub.Form.curr_unit_lbl_on.Visible = False
29140     ElseIf .opgMetal = .opgMetal_optOnly.OptionValue Then
29150       .frmCurrency_Sub.Form.curr_metal.Visible = True
29160       .frmCurrency_Sub.Form.curr_metal_lbl_off.Visible = True
29170       .frmCurrency_Sub.Form.curr_metal_lbl_on.Visible = True
29180       .frmCurrency_Sub.Form.curr_fund.Visible = False
29190       .frmCurrency_Sub.Form.curr_fund_lbl_off.Visible = False
29200       .frmCurrency_Sub.Form.curr_fund_lbl_on.Visible = False
29210       .frmCurrency_Sub.Form.curr_bmu.Visible = False
29220       .frmCurrency_Sub.Form.curr_bmu_lbl_off.Visible = False
29230       .frmCurrency_Sub.Form.curr_bmu_lbl_on.Visible = False
29240       .frmCurrency_Sub.Form.curr_alt.Visible = False
29250       .frmCurrency_Sub.Form.curr_alt_lbl_off.Visible = False
29260       .frmCurrency_Sub.Form.curr_alt_lbl_on.Visible = False
29270       .frmCurrency_Sub.Form.curr_unit.Visible = False
29280       .frmCurrency_Sub.Form.curr_unit_lbl_off.Visible = False
29290       .frmCurrency_Sub.Form.curr_unit_lbl_on.Visible = False
29300     ElseIf .opgAlt = .opgAlt_optOnly.OptionValue Then
29310       .frmCurrency_Sub.Form.curr_alt.Visible = True
29320       .frmCurrency_Sub.Form.curr_alt_lbl_off.Visible = True
29330       .frmCurrency_Sub.Form.curr_alt_lbl_on.Visible = True
29340       .frmCurrency_Sub.Form.curr_fund.Visible = False
29350       .frmCurrency_Sub.Form.curr_fund_lbl_off.Visible = False
29360       .frmCurrency_Sub.Form.curr_fund_lbl_on.Visible = False
29370       .frmCurrency_Sub.Form.curr_bmu.Visible = False
29380       .frmCurrency_Sub.Form.curr_bmu_lbl_off.Visible = False
29390       .frmCurrency_Sub.Form.curr_bmu_lbl_on.Visible = False
29400       .frmCurrency_Sub.Form.curr_metal.Visible = False
29410       .frmCurrency_Sub.Form.curr_metal_lbl_off.Visible = False
29420       .frmCurrency_Sub.Form.curr_metal_lbl_on.Visible = False
29430       .frmCurrency_Sub.Form.curr_unit.Visible = False
29440       .frmCurrency_Sub.Form.curr_unit_lbl_off.Visible = False
29450       .frmCurrency_Sub.Form.curr_unit_lbl_on.Visible = False
29460     ElseIf .opgUnit = .opgUnit_optOnly.OptionValue Then
29470       .frmCurrency_Sub.Form.curr_unit.Visible = True
29480       .frmCurrency_Sub.Form.curr_unit_lbl_off.Visible = True
29490       .frmCurrency_Sub.Form.curr_unit_lbl_on.Visible = True
29500       .frmCurrency_Sub.Form.curr_fund.Visible = False
29510       .frmCurrency_Sub.Form.curr_fund_lbl_off.Visible = False
29520       .frmCurrency_Sub.Form.curr_fund_lbl_on.Visible = False
29530       .frmCurrency_Sub.Form.curr_bmu.Visible = False
29540       .frmCurrency_Sub.Form.curr_bmu_lbl_off.Visible = False
29550       .frmCurrency_Sub.Form.curr_bmu_lbl_on.Visible = False
29560       .frmCurrency_Sub.Form.curr_metal.Visible = False
29570       .frmCurrency_Sub.Form.curr_metal_lbl_off.Visible = False
29580       .frmCurrency_Sub.Form.curr_metal_lbl_on.Visible = False
29590       .frmCurrency_Sub.Form.curr_alt.Visible = False
29600       .frmCurrency_Sub.Form.curr_alt_lbl_off.Visible = False
29610       .frmCurrency_Sub.Form.curr_alt_lbl_on.Visible = False
29620     Else
29630       If .opgFund = .opgFund_optExclude.OptionValue Then
29640         .frmCurrency_Sub.Form.curr_fund.Visible = False
29650         .frmCurrency_Sub.Form.curr_fund_lbl_off.Visible = False
29660         .frmCurrency_Sub.Form.curr_fund_lbl_on.Visible = False
29670       ElseIf .opgFund = .opgFund_optInclude.OptionValue Then
29680         .frmCurrency_Sub.Form.curr_fund.Visible = True
29690         .frmCurrency_Sub.Form.curr_fund_lbl_off.Visible = True
29700         .frmCurrency_Sub.Form.curr_fund_lbl_on.Visible = True
29710       End If
29720       If .opgBMU = .opgBMU_optExclude.OptionValue Then
29730         .frmCurrency_Sub.Form.curr_bmu.Visible = False
29740         .frmCurrency_Sub.Form.curr_bmu_lbl_off.Visible = False
29750         .frmCurrency_Sub.Form.curr_bmu_lbl_on.Visible = False
29760       ElseIf .opgBMU = .opgBMU_optInclude.OptionValue Then
29770         .frmCurrency_Sub.Form.curr_bmu.Visible = True
29780         .frmCurrency_Sub.Form.curr_bmu_lbl_off.Visible = True
29790         .frmCurrency_Sub.Form.curr_bmu_lbl_on.Visible = True
29800       End If
29810       If .opgMetal = .opgMetal_optExclude.OptionValue Then
29820         .frmCurrency_Sub.Form.curr_metal.Visible = False
29830         .frmCurrency_Sub.Form.curr_metal_lbl_off.Visible = False
29840         .frmCurrency_Sub.Form.curr_metal_lbl_on.Visible = False
29850       ElseIf .opgMetal = .opgMetal_optInclude.OptionValue Then
29860         .frmCurrency_Sub.Form.curr_metal.Visible = True
29870         .frmCurrency_Sub.Form.curr_metal_lbl_off.Visible = True
29880         .frmCurrency_Sub.Form.curr_metal_lbl_on.Visible = True
29890       End If
29900       If .opgAlt = .opgAlt_optExclude.OptionValue Then
29910         .frmCurrency_Sub.Form.curr_alt.Visible = False
29920         .frmCurrency_Sub.Form.curr_alt_lbl_off.Visible = False
29930         .frmCurrency_Sub.Form.curr_alt_lbl_on.Visible = False
29940       ElseIf .opgAlt = .opgAlt_optInclude.OptionValue Then
29950         .frmCurrency_Sub.Form.curr_alt.Visible = True
29960         .frmCurrency_Sub.Form.curr_alt_lbl_off.Visible = True
29970         .frmCurrency_Sub.Form.curr_alt_lbl_on.Visible = True
29980       End If
29990       If .opgUnit = .opgUnit_optExclude.OptionValue Then
30000         .frmCurrency_Sub.Form.curr_unit.Visible = False
30010         .frmCurrency_Sub.Form.curr_unit_lbl_off.Visible = False
30020         .frmCurrency_Sub.Form.curr_unit_lbl_on.Visible = False
30030       ElseIf .opgUnit = .opgUnit_optInclude.OptionValue Then
30040         .frmCurrency_Sub.Form.curr_unit.Visible = True
30050         .frmCurrency_Sub.Form.curr_unit_lbl_off.Visible = True
30060         .frmCurrency_Sub.Form.curr_unit_lbl_on.Visible = True
30070       End If
30080     End If
30090   End With

EXITP:
30100   Exit Sub

ERRH:
30110   Select Case ERR.Number
        Case Else
30120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30130   End Select
30140   Resume EXITP

End Sub
