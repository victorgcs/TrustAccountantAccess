VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Maintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Maintenance"

'VGC 10/11/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Audit Report:     A {cmdMaint01}
' **   Close:            C {cmdClose}
' **   Load Timer:       D {chkLoadTimer}
' **   Error Log:        E {cmdMaint08}
' **   Foreign Exchange: F {cmdMaint04}
' **   Shortcuts:        K {cmdMaint07}
' **   Master Balance:   M {cmdMaint02}
' **   Account Contacts: N {cmdMaint03}
' **   Form Load Mgmt:   O {cmdMaint10}
' **   Check Register:   R {cmdMaint05}
' **   Administration:   S {cmdMaint09}
' **   System Tables:    T {cmdMaint06}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Twips-per-pixel:  P {TPP}
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm}

Private blnMaint04_Focus As Boolean, blnMaint06_Focus As Boolean, blnMaint09_Focus As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long

Private blnFormOpening As Boolean, blnTimerVis As Boolean
Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnContactsCalled As Boolean, blnJustClose As Boolean, blnAuditCalled As Boolean
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long ', lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim lngTmp01 As Long

110     DoCmd.Restore

120     With Me
130       If IsNull(.OpenArgs) = False Then
140         If Trim(.OpenArgs) <> vbNullString Then

150           blnIsOpen = True

160           strCallingForm = Trim(.OpenArgs)

170           If CurrentUser <> "superuser" Then  ' ** Internal Access Function: Trust Accountant login.
180             .cmdMaint09.Enabled = False
190             .cmdMaint09.Visible = False
200           Else
210             .cmdMaint09.Enabled = True
220             .cmdMaint09.Visible = True
230           End If

240           gblnSignal = False: blnTimerVis = False

250           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

260           gstrCrtRpt_Version = vbNullString  ' ** Borrowing this variable from Court Reports.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
270           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
280           .TPP = lngTpp

              ' ** Variables are fed empty, then populated ByRef.
290           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

300           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

310           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
320           lngMonitorNum = 1&: lngTmp01 = 0&
330           EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
340           If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

350           If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
360           DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, lngFrm_Height  'lngFrm_Top
370           If lngMonitorNum > 1& Then
380             LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
390           End If

400           blnJustClose = False: blnContactsCalled = False: blnAuditCalled = False
410           blnGTR_Emblem = False: lngGTR_Stat = 0&
420           blnMaint04_Focus = False: blnMaint06_Focus = False: blnMaint09_Focus = False

430           Select Case .chkLoadTimer_vis
              Case True
440             .chkLoadTimer.Visible = True
450             .chkLoadTimer_lbl1.Visible = True
460             .chkLoadTimer_lbl2.Visible = True
470             .chkLoadTimer_lbl3.Visible = True
480             .cmdMaint10.Visible = True
490             blnTimerVis = True
500           Case False
510             .chkLoadTimer.Visible = False
520             .chkLoadTimer_lbl1.Visible = False
530             .chkLoadTimer_lbl2.Visible = False
540             .chkLoadTimer_lbl3.Visible = False
550             .cmdMaint10.Visible = False
560           End Select
570           chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
580           DoEvents

590           blnIsOpen = False

600         Else
610           Cancel = -1
620         End If
630       Else
640         Cancel = -1
650       End If
660     End With

670     If Cancel = -1 Then
680       Beep
690       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
700     End If

EXITP:
710     Exit Sub

ERRH:
720     Select Case ERR.Number
        Case Else
730       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
740     End Select
750     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strUsername As String
        Dim intRetVal As Integer, Cancel As Integer

810     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
820     intShiftDown = (Shift And acShiftMask) > 0
830     intAltDown = (Shift And acAltMask) > 0
840     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Load Timer:       D {chkLoadTimer}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Timer Chk Vis:    T {Form_KeyDown}
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyEscape
870         intRetVal = 0
880         cmdClose_Click  ' ** Procedure: Below.
890       End Select
900     End If

        ' ** Alt keys.
910     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
920       Select Case intRetVal
          Case vbKeyD
930         With Me
940           intRetVal = 0
950           If .chkLoadTimer.Visible = True Then
960             If .chkLoadTimer.Enabled = True Then
970               .chkLoadTimer.SetFocus
980             End If
990             .chkLoadTimer = (Not .chkLoadTimer)
1000            chkLoadTimer_AfterUpdate
1010          End If
1020        End With
1030      Case vbKeyX
1040        intRetVal = 0
            ' ** In case this gets stuck, using this keystroke will turn it off
1050        blnFormOpening = False
1060        cmdClose_Click  ' ** Procedure: Below.
1070      End Select
1080    End If

        ' ** Ctrl-Shift keys.
1090    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1100      Select Case intRetVal
          Case vbKeyF
1110        With Me
1120          intRetVal = 0
1130          .FocusHolder.SetFocus
1140        End With
1150      Case vbKeyP
1160        With Me
1170          intRetVal = 0
1180          .TPP.Visible = (Not .TPP.Visible)
1190        End With
1200      Case vbKeyT
1210        With Me
1220          intRetVal = 0
1230          If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
1240            FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
1250          ElseIf CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
1260            Select Case .chkLoadTimer.Visible
                Case True
1270              .FocusHolder.SetFocus
1280              DoEvents
1290              .chkLoadTimer.Visible = False
1300              .chkLoadTimer_lbl1.Visible = False
1310              .chkLoadTimer_lbl2.Visible = False
1320              .chkLoadTimer_lbl3.Visible = False
1330              .cmdMaint10.Visible = False
1340              gblnTimer = False
1350              blnTimerVis = False
1360            Case False
1370              .chkLoadTimer.Visible = True
1380              .cmdMaint10.Visible = True
1390              .chkLoadTimer_lbl1.Visible = True
1400              .chkLoadTimer_lbl2.Visible = True
1410              .chkLoadTimer_lbl3.Visible = True
1420              blnTimerVis = True
1430              chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
1440            End Select
1450            .chkLoadTimer_vis = .chkLoadTimer.Visible
1460          End If
1470        End With
1480      Case vbKeyX
1490        strUsername = GetUserName  ' ** Module Function: modFileUtilities.
1500        If strUsername = gstrDevUserName Then
1510          blnJustClose = True
1520          intRetVal = 0
1530          SetOption_Dev  ' ** Module Function: modUtilities.
1540          gblnSignal = True  ' ** Borrowing this variable.
1550          DoCmd.Close acForm, THIS_NAME
1560          DoCmd.Hourglass False
1570        End If
1580      End Select
1590    End If

EXITP:
1600    KeyCode = intRetVal
1610    Exit Sub

ERRH:
1620    intRetVal = 0
1630    Select Case ERR.Number
        Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Private Sub Form_Timer()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim strFormName As String
        Dim blnTmp01 As Boolean

        Static lngTmp01 As Long, lngTmp02 As Long

1710    With Me
1720      .TimerInterval = 0&
1730      If blnGTR_Emblem = True Then
1740        lngGTR_Stat = lngGTR_Stat + 1&
1750        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1760        lngTmp01 = lngTmp01 + 1&
1770        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1780        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1790        If lngTmp01 > 1& Then
1800          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1810        End If
1820        If lngTmp01 = 24& Then
1830          lngTmp01 = 0&
1840          lngTmp02 = lngTmp02 + 1&
1850          If lngTmp02 < 5& Then
1860            .TimerInterval = 50&
1870          Else
1880            blnGTR_Emblem = False
1890          End If
1900        Else
1910          .TimerInterval = 50&
1920        End If
1930      End If  ' ** blnGTR_Emblem.
1940      If gblnSetFocus = True Then
1950        gblnSetFocus = False
1960        strFormName = vbNullString
1970  On Error Resume Next
1980        Do Until strFormName = THIS_NAME
1990          DoCmd.SelectObject acForm, THIS_NAME, False
2000          strFormName = Screen.ActiveForm.Name
2010        Loop
2020  On Error GoTo ERRH
2030        DoEvents
2040      End If
2050      If blnGoingToReport = False Then
2060        Select Case gstrReturningForm
            Case "frmTransaction_Audit"
2070          .cmdMaint01.SetFocus
2080          blnFormOpening = False
2090          If IsNothing(gctlPBar_Box1) = False Then  ' ** Module Function: modUtilities.
2100  On Error Resume Next
2110            blnTmp01 = gctlPBar_Box1.Visible
2120  On Error GoTo ERRH
2130            If blnTmp01 = True Then
2140              blnAuditCalled = True
2150              .TimerInterval = 3000&
2160            End If
2170          End If
2180        Case "frmMasterBalance"
2190          .cmdMaint02.SetFocus
2200        Case "frmAccountContacts"
2210          .cmdMaint03.SetFocus
2220          blnContactsCalled = True
2230        Case "frmMenu_ForeignExchange"
2240          .cmdMaint04.SetFocus
2250        Case "frmCheckRegister"
2260          .cmdMaint05.SetFocus
2270        Case "frmMenu_SystemTable"
2280          .cmdMaint06.SetFocus
2290        Case "frmXAdmin_Shortcut"
2300          .cmdMaint07.SetFocus
2310        Case "frmErrorLog"
2320          .cmdMaint08.SetFocus
2330        Case "frmMenu_Admin"
2340          .cmdMaint09.SetFocus
2350        Case "frmLoadTimer_Mgmt"
2360          .cmdMaint10.SetFocus
2370        Case Else
2380          .cmdMaint01.SetFocus
2390        End Select
2400        gstrReturningForm = vbNullString
2410        DoCmd.Hourglass False
2420      End If
2430      If gblnGoToReport = True Then
2440        DoCmd.Hourglass True  ' ** Make sure it's still running.
2450        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
2460        Else
2470          If blnGTR_Emblem = False Then
2480            blnGTR_Emblem = True
2490            lngTmp01 = 0&: lngTmp02 = 0&
2500          End If
2510          DoEvents
2520          Select Case blnGoingToReport
              Case True
2530            Select Case garr_varGoToReport(GTR_INV2)
                Case True
2540              Select Case garr_varGoToReport(GTR_CTL2)
                  Case "cmdMaint01"
2550                .TimerInterval = 50&
2560                cmdMaint01_Click  ' ** Procedure: Below.
2570              Case "cmdMaint02"
2580                .TimerInterval = 50&
2590                cmdMaint02_Click  ' ** Procedure: Below.
2600              Case "cmdMaint03"
2610                .TimerInterval = 50&
2620                cmdMaint03_Click  ' ** Procedure: Below.
2630              Case "cmdMaint04"
2640                .TimerInterval = 50&
2650                cmdMaint04_Click  ' ** Procedure: Below.
2660              Case "cmdMaint06"
2670                .TimerInterval = 50&
2680                cmdMaint06_Click  ' ** Procedure: Below.
2690              Case "cmdMaint08"
2700                .TimerInterval = 50&
2710                cmdMaint08_Click  ' ** Procedure: Below.
2720              Case Else
                    ' ** Shouldn't happen.
2730                blnGoingToReport = False
2740                DoCmd.Hourglass False
2750                gblnGoToReport = False
2760                blnGTR_Emblem = False
2770                GTREmblem_Off  ' ** Procedure: Below.
2780                Beep
2790              End Select
2800            Case False
                  ' ** Just sit there, and we're done.
2810              blnGoingToReport = False
2820              DoCmd.Hourglass False
2830              gblnGoToReport = False
2840              blnGTR_Emblem = False
2850              GTREmblem_Off  ' ** Procedure: Below.
2860              Beep
2870            End Select
2880          Case False
2890            If garr_varGoToReport(GTR_FRM2) = THIS_NAME Then
                  ' ** This is second menu in path.
2900              blnGoingToReport = True
2910              Select Case IsNull(garr_varGoToReport(GTR_FRM3))
                  Case True
                    ' ** No more menus.
2920                .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
2930                lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL2)).Top
2940                lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL2)).Left
2950                .GoToReport_arw_menu_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
2960                .GoToReport_arw_menu_img.Visible = True
2970                DoEvents
2980                .TimerInterval = 50&
2990              Case False
                    ' ** Moving on.
3000                Select Case garr_varGoToReport(GTR_CTL2)
                    Case "cmdMaint04"  ' ** 4. Foreign Exchange.
3010                  .GoToReport_arw_menu_forex_img.Visible = True
3020                Case "cmdMaint06"  ' ** 6. System Tables.
3030                  .GoToReport_arw_menu_other_img.Visible = True
3040                Case "cmdMaint09"  ' ** 8. Administration.
3050                  .GoToReport_arw_menu_admin_img.Visible = True
3060                Case Else
3070                  lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL2)).Top
3080                  lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL2)).Left
3090                  .GoToReport_arw_menu_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
3100                  .GoToReport_arw_menu_img.Visible = True
3110                End Select
3120                .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
3130                Select Case garr_varGoToReport(GTR_INV2)
                    Case True
3140                  DoEvents
3150                  .TimerInterval = 50&
3160                Case False
                      ' ** Just sit there, and we're done.
3170                  blnGoingToReport = False
3180                  DoCmd.Hourglass False
3190                  gblnGoToReport = False
3200                  blnGTR_Emblem = False
3210                  GTREmblem_Off  ' ** Procedure: Below.
3220                  Beep
3230                End Select
3240              End Select
3250            Else
                  ' ** Shouldn't be here, then!
3260              blnGoingToReport = False
3270              DoCmd.Hourglass False
3280              gblnGoToReport = False
3290              blnGTR_Emblem = False
3300              GTREmblem_Off  ' ** Procedure: Below.
3310              Beep
3320            End If
3330          End Select
3340        End If
3350      Else
3360        DoCmd.Hourglass False
3370      End If
3380      If blnAuditCalled = True Then
3390        blnAuditCalled = False
3400  On Error Resume Next
3410        gctlPBar_Box1.Visible = False
3420        gctlPBar_Box2.Visible = False
3430        ProgBar_Width_Maint False, 1  ' ** Procedure: Below.
3440        gctlPBar_Lbl.Visible = False
3450        Set gctlPBar_Box1 = Nothing
3460        Set gctlPBar_Box2 = Nothing
3470        Set gctlPBar_Lbl = Nothing
3480  On Error GoTo ERRH
3490      End If
3500      If gblnTimer = True Then
3510        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
3520          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
3530          gblnTimer = False
3540          DoCmd.Hourglass False
3550        End If
3560      End If
3570    End With

EXITP:
3580    Exit Sub

ERRH:
3590    DoCmd.Hourglass False
3600    blnGoingToReport = False
3610    gblnGoToReport = False
3620    Select Case ERR.Number
        Case Else
3630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3640    End Select
3650    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

3710    With Me

3720      If gblnSignal = False Then  ' ** Borrowing this variable.
3730        DoCmd.Hourglass True
3740        DoEvents
3750      End If

          ' ** Make sure all reports are closed.
3760      Do While Reports.Count > 0
3770        DoCmd.Close acReport, Reports(0).Name
3780      Loop

3790      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3800      If blnContactsCalled = True And blnJustClose = False Then
            ' ** Only works if they haven't gone to another form since setting the variable.
3810        Set dbs = CurrentDb
            ' ** Empty tblAccount_Contact_Staging.
3820        Set qdf = dbs.QueryDefs("qryAccount_Contacts_04")
3830        qdf.Execute
3840        Set qdf = Nothing
3850        dbs.Close
3860        Set dbs = Nothing
3870      End If

3880      If blnJustClose = False Then

3890        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
3900        DoEvents

3910        gblnSetFocus = True
3920        gstrReturningForm = THIS_NAME
3930        Select Case strCallingForm
            Case "frmMenu_Utility"
3940          If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
3950            DoCmd.SelectObject acForm, strCallingForm, False
3960            Forms(strCallingForm).TimerInterval = 250&
3970          Else
3980            DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
3990          End If
4000        End Select

4010      End If

4020    End With

EXITP:
4030    Set qdf = Nothing
4040    Set dbs = Nothing
4050    Exit Sub

ERRH:
4060    DoCmd.Hourglass False
4070    Select Case ERR.Number
        Case Else
4080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4090    End Select
4100    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4210    With Me
4220      If blnMaint04_Focus = False Then
4230        If .cmdMaint04_box01.Visible = True And gblnGoToReport = False Then
4240          .cmdMaint04_box01.Visible = False
4250          .cmdMaint04_box02.Visible = False
4260        End If
4270      End If
4280      If blnMaint06_Focus = False Then
4290        If .cmdMaint06_box01.Visible = True And gblnGoToReport = False Then
4300          .cmdMaint06_box01.Visible = False
4310          .cmdMaint06_box02.Visible = False
4320        End If
4330      End If
4340      If blnMaint09_Focus = False Then
4350        If .cmdMaint09_box01.Visible = True And gblnGoToReport = False Then
4360          .cmdMaint09_box01.Visible = False
4370          .cmdMaint09_box02.Visible = False
4380        End If
4390      End If
4400      If .cmdClose.Caption <> "&Close" Then
4410        .cmdClose.Caption = "&Close"
4420      End If
4430    End With

EXITP:
4440    Exit Sub

ERRH:
4450    Select Case ERR.Number
        Case Else
4460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4470    End Select
4480    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

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

Private Sub Header_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_box_MouseMove"

4610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub Header_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_MouseMove"

4710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4720    Exit Sub

ERRH:
4730    Select Case ERR.Number
        Case Else
4740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4750    End Select
4760    Resume EXITP

End Sub

Private Sub Header_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl2_MouseMove"

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

Private Sub cmdClose_Click()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

4910    With Me
4920      If blnFormOpening = False Then

4930        DoCmd.Hourglass True
4940        DoEvents

4950        .TimerInterval = 0&

4960        blnJustClose = False
4970        DoCmd.Close acForm, THIS_NAME

4980      End If
4990    End With

EXITP:
5000    Exit Sub

ERRH:
5010    blnFormOpening = False
5020    DoCmd.Hourglass False
5030    Select Case ERR.Number
        Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
5170        With Me
5180          intRetVal = 0
5190          .cmdMaint01.SetFocus
5200        End With
5210      End Select
5220    End If

        ' ** Shift keys.
5230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5240      Select Case intRetVal
          Case vbKeyTab
5250        With Me
5260          intRetVal = 0
5270          If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
5280            .chkLoadTimer.SetFocus
5290          ElseIf .cmdMaint09.Enabled = True And .cmdMaint09.Visible = True Then
5300            .cmdMaint09.SetFocus
5310          Else
5320            .cmdMaint08.SetFocus
5330          End If
5340        End With
5350      End Select
5360    End If

EXITP:
5370    KeyCode = intRetVal
5380    Exit Sub

ERRH:
5390    intRetVal = 0
5400    Select Case ERR.Number
        Case Else
5410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5420    End Select
5430    Resume EXITP

End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_MouseMove"

5510    With Me
5520      .cmdClose.Caption = "< Ba&ck"
5530    End With

EXITP:
5540    Exit Sub

ERRH:
5550    Select Case ERR.Number
        Case Else
5560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5570    End Select
5580    Resume EXITP

End Sub

Private Sub cmdMaint_l_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint_l_box01_MouseMove"

5610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5620    Exit Sub

ERRH:
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Private Sub cmdMaint01_Click()
' ** 1. Audit Report.

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint01_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

5710    With Me
5720      If blnFormOpening = False Then

5730        strDocName = "frmTransaction_Audit"

5740        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
5750        DoEvents

5760        If .chkLoadTimer.Visible = True Then
5770          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
5780          If gblnTimer = True Then
                'gblnSetFocus = True
5790            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
5800          End If
5810        End If
5820        DoEvents

5830        DoCmd.Hourglass True
5840        DoEvents

5850        .TimerInterval = 0&

5860        blnFormOpening = True

            ' ****************************************
            ' ** Progress Bar setup.
            ' ****************************************
5870        gdblPBar_MaxWidth = .ProgBar_box.Width
5880        gdblPBar_Steps = 25#
5890        gdblPBar_Increment = (gdblPBar_MaxWidth / gdblPBar_Steps)
5900        gdblPBar_CurWidth = 0#
5910        gdblPBar_ThisStep = 0#
5920        Set gctlPBar_Box1 = .ProgBar_box
5930        Set gctlPBar_Box2 = .ProgBar_box2
5940        Set gctlPBar_Lbl = .ProgBar_lbl
            ' ****************************************

5950        If blnTimerVis = True Then
5960          .chkLoadTimer.Visible = False
5970        End If

            ' ****************************************
5980        gctlPBar_Box1.Visible = True
5990        gctlPBar_Box2.Visible = True
6000        ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Procedure: Below.
6010        ProgBar_Width_Maint True, 1  ' ** Procedure: Below.
6020        If CurrentUser = "Superuser" Then
6030          gctlPBar_Lbl.Visible = True
6040        Else
6050          gctlPBar_Lbl.Visible = False
6060        End If
6070        DoEvents
            ' ****************************************

6080        Set dbs = CurrentDb
6090        With dbs

              ' ****************************************
              ' ** Step 1. Query 1
              ' ****************************************
6100          gdblPBar_ThisStep = 1#
              ' ****************************************
              ' ** Update Ledger, where journal_USER = Null, set to 'System'.
6110          Set qdf = .QueryDefs("qryTransaction_Audit_04a")
6120          qdf.Execute
6130          Set qdf = Nothing
6140          DoEvents
              ' ****************************************
6150          gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
6160          ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Procedure: Below.
6170          gctlPBar_Lbl.Caption = "Step 1."
6180          DoEvents
              ' ****************************************

              ' ****************************************
              ' ** Step 2. Query 2
              ' ****************************************
6190          gdblPBar_ThisStep = 2#
              ' ****************************************
              ' ** Update LedgerArchive, where journal_USER = Null, set to 'System'.
6200          Set qdf = .QueryDefs("qryTransaction_Audit_04b")
6210          qdf.Execute
6220          Set qdf = Nothing
6230          DoEvents
              ' ****************************************
6240          gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
6250          ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Procedure: Below.
6260          gctlPBar_Lbl.Caption = "Step 2."
6270          DoEvents
              ' ****************************************

              ' ****************************************
              ' ** Step 3. Query 3
              ' ****************************************
6280          gdblPBar_ThisStep = 3#
              ' ****************************************
              ' ** Empty tmpAccount.
6290          Set qdf = .QueryDefs("qryTransaction_Audit_12_01")
6300          qdf.Execute
6310          Set qdf = Nothing
6320          DoEvents
              ' ** Append qryAccountMenu_01_10 (qryAccountMenu_01_01 (xx), linked to
              ' ** qryAccountMenu_01_08 (xx), just accountno, with acct_sort) to tmpAccount.
6330          Set qdf = .QueryDefs("qryTransaction_Audit_12_02")
6340          qdf.Execute
6350          Set qdf = Nothing
6360          DoEvents
              ' ****************************************
6370          gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
6380          ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Procedure: Below.
6390          gctlPBar_Lbl.Caption = "Step 3."
6400          DoEvents
              ' ****************************************

6410          .Close
6420        End With
6430        DoEvents

            ' ****************************************
            ' ** Step 4. Open form
            ' ****************************************
6440        gdblPBar_ThisStep = 4#
            ' ****************************************
6450        gblnSetFocus = True
6460        blnJustClose = True
6470        gstrReturningForm = strDocName
6480        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6490        If gblnGoToReport = True Then
6500          Forms(strDocName).TimerInterval = 50&
6510          blnGoingToReport = False
6520          blnGTR_Emblem = False
6530        End If

            ' ** frmTransaction_Audit closes the form.
            'DoCmd.Close acForm, THIS_NAME

6540        DoCmd.Hourglass False

6550      End If
6560    End With

EXITP:
6570    Set qdf = Nothing
6580    Set dbs = Nothing
6590    Exit Sub

ERRH:
6600    blnFormOpening = False
6610    DoCmd.Hourglass False
6620    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
6630    Case Else
6640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6650    End Select
6660    Resume EXITP

End Sub

Private Sub cmdMaint01_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint01_KeyDown"

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
6790          .cmdMaint02.SetFocus
6800        End With
6810      End Select
6820    End If

        ' ** Shift keys.
6830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6840      Select Case intRetVal
          Case vbKeyTab
6850        With Me
6860          intRetVal = 0
6870          .cmdClose.SetFocus
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

Private Sub cmdMaint01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint01_MouseMove"

7010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7020    Exit Sub

ERRH:
7030    Select Case ERR.Number
        Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Sub

Private Sub cmdMaint02_Click()
' ** 2. Master Balance.

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint02_Click"

        Dim strDocName As String

7110    With Me
7120      If blnFormOpening = False Then

7130        DoCmd.Hourglass True
7140        DoEvents

7150        .TimerInterval = 0&

            'gblnSetFocus = True  ' ** Leave this off so it hits the right TimerInterval code.
7160        blnJustClose = True: blnFormOpening = True
7170        strDocName = "frmMasterBalance"
7180        gstrReturningForm = strDocName
7190        gstrFormQuerySpec = strDocName  ' ** Need results of gblnAccountNoWithType before form fully loaded.
7200        DoCmd.Close acForm, THIS_NAME
7210        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
7220        If gblnGoToReport = True Then
7230          Forms(strDocName).TimerInterval = 50&
7240          blnGoingToReport = False
7250          blnGTR_Emblem = False
7260        End If

7270        DoCmd.Hourglass False

7280      End If
7290    End With

EXITP:
7300    Exit Sub

ERRH:
7310    blnFormOpening = False
7320    DoCmd.Hourglass False
7330    Select Case ERR.Number
        Case Else
7340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7350    End Select
7360    Resume EXITP

End Sub

Private Sub cmdMaint02_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint02_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7420    intShiftDown = (Shift And acShiftMask) > 0
7430    intAltDown = (Shift And acAltMask) > 0
7440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7460      Select Case intRetVal
          Case vbKeyTab
7470        With Me
7480          intRetVal = 0
7490          .cmdMaint03.SetFocus
7500        End With
7510      End Select
7520    End If

        ' ** Shift keys.
7530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7540      Select Case intRetVal
          Case vbKeyTab
7550        With Me
7560          intRetVal = 0
7570          .cmdMaint01.SetFocus
7580        End With
7590      End Select
7600    End If

EXITP:
7610    KeyCode = intRetVal
7620    Exit Sub

ERRH:
7630    intRetVal = 0
7640    Select Case ERR.Number
        Case Else
7650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7660    End Select
7670    Resume EXITP

End Sub

Private Sub cmdMaint02_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint02_MouseMove"

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

Private Sub cmdMaint03_Click()
' ** 3. Account Contacts.

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint03_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnNoContacts As Boolean
        Dim varTmp00 As Variant

7810    With Me
7820      If blnFormOpening = False Then

7830        DoCmd.Hourglass True
7840        DoEvents

7850        .TimerInterval = 0&

7860        blnFormOpening = True
7870        blnNoContacts = False

7880        Set dbs = CurrentDb
7890        With dbs

              ' ** Empty tblAccount_Contact_Staging.
7900          Set qdf = .QueryDefs("qryAccount_Contacts_04")
7910          qdf.Execute
7920          Set qdf = Nothing

              ' ** qryAccount_Contacts_03_01 (qryAccount_Contacts_02c
              ' ** (Union of qryAccount_Contacts_02a (Account, just Contact1), qryAccount_Contacts_02b
              ' ** (Account, just Contact2)), with new field names), not in qryAccount_Contacts_03_02
              ' ** (qryAccount_Contacts_03_01 (qryAccount_Contacts_02c (Union of qryAccount_Contacts_02a
              ' ** (Account, just Contact1), qryAccount_Contacts_02b (Account, just Contact2)), with new
              ' ** field names), just empty contacts), just good contacts.
7930          varTmp00 = DCount("*", "qryAccount_Contacts_03_03")
7940          Select Case IsNull(varTmp00)
              Case True
7950            blnNoContacts = True
7960          Case False
7970            If varTmp00 = 0 Then
7980              blnNoContacts = True
7990            End If
8000          End Select

8010          If blnNoContacts = False Then
                ' ** Append qryAccount_Contacts_03_03 (qryAccount_Contacts_03_01 (qryAccount_Contacts_02c
                ' ** (Union of qryAccount_Contacts_02a (Account, just Contact1), qryAccount_Contacts_02b
                ' ** (Account, just Contact2)), with new field names), not in qryAccount_Contacts_03_02
                ' ** (qryAccount_Contacts_03_01 (qryAccount_Contacts_02c (Union of qryAccount_Contacts_02a
                ' ** (Account, just Contact1), qryAccount_Contacts_02b (Account, just Contact2)), with new
                ' ** field names), just empty contacts), just good contacts) to tblAccount_Contact_Staging.
8020            Set qdf = .QueryDefs("qryAccount_Contacts_05")
8030            qdf.Execute dbFailOnError
8040            Set qdf = Nothing
8050            DoEvents
                ' ** Update qryAccount_Contacts_06a (tblAccount_Contact_Staging,
                ' ** just Contact_Zip with Contact_Zip_new, Contact_Zip_Format_new).
8060            Set qdf = .QueryDefs("qryAccount_Contacts_06b")
8070            qdf.Execute
8080            Set qdf = Nothing
8090            DoEvents
                ' ** Update qryAccount_Contacts_07a (tblAccount_Contact_Staging,
                ' ** just Contact_Phone1 with Contact_Phone1_new, Contact_Phone1_Format_new).
8100            Set qdf = .QueryDefs("qryAccount_Contacts_07b")
8110            qdf.Execute
8120            Set qdf = Nothing
8130            DoEvents
                ' ** Update qryAccount_Contacts_08a (tblAccount_Contact_Staging,
                ' ** just Contact_Phone2 with Contact_Phone2_new, Contact_Phone2_Format_new).
8140            Set qdf = .QueryDefs("qryAccount_Contacts_08b")
8150            qdf.Execute
8160            Set qdf = Nothing
8170            DoEvents
                ' ** Update qryAccount_Contacts_09a (tblAccount_Contact_Staging,
                ' ** just Contact_Fax with Contact_Fax_new, Contact_Fax_Format_new.).
8180            Set qdf = .QueryDefs("qryAccount_Contacts_09b")
8190            qdf.Execute
8200            Set qdf = Nothing
8210            DoEvents
                ' ** Update qryAccount_Contacts_10b (qryAccount_Contacts_10a
                ' ** (tblAccount_Contact_Staging, just Contact_Name with
                ' ** Contact_Name_Last_new), with Contact_Name_First_new).
8220            Set qdf = .QueryDefs("qryAccount_Contacts_10c")
8230            qdf.Execute
8240            Set qdf = Nothing
8250            DoEvents
                ' ** Update qryAccount_Contacts_11a (tblAccount_Contact_Staging, with Contact_LastFirst_new).
8260            Set qdf = .QueryDefs("qryAccount_Contacts_11b")
8270            qdf.Execute
8280            Set qdf = Nothing
8290            DoEvents
                ' ** Update qryAccount_Contacts_01_04 (tblAccount_Contact_Staging,
                ' ** just Contact_Email = Null, with .._new fields).
8300            Set qdf = .QueryDefs("qryAccount_Contacts_01_05")
8310            qdf.Execute
8320            Set qdf = Nothing
8330            DoEvents
8340          End If
8350          .Close
8360        End With  ' ** dbs.
8370        Set qdf = Nothing
8380        Set dbs = Nothing
8390        blnContactsCalled = True

8400        Select Case blnNoContacts
            Case True
8410          DoCmd.Hourglass False
8420          Beep
8430          msgResponse = MsgBox("There are no Contacts currently listed for any account." & vbCrLf & _
                "Do you wish to continue?", vbInformation + vbOKCancel + vbDefaultButton1, "No Contacts")
8440        Case False
8450          msgResponse = vbOK
8460        End Select

8470        If msgResponse = vbOK Then
8480          DoCmd.Hourglass True
8490          DoEvents
8500          gblnSetFocus = True
8510          blnJustClose = True
8520          strDocName = "frmAccountContacts"
8530          gstrReturningForm = strDocName
8540          DoCmd.Close acForm, THIS_NAME
8550          DoCmd.OpenForm strDocName, , , , , , THIS_NAME
8560          If gblnGoToReport = True Then
8570            Forms(strDocName).TimerInterval = 50&
8580            blnGoingToReport = False
8590            blnGTR_Emblem = False
8600          End If
8610        Else
8620          blnFormOpening = False
8630        End If

8640        DoCmd.Hourglass False

8650      End If
8660    End With

EXITP:
8670    Set qdf = Nothing
8680    Set dbs = Nothing
8690    Exit Sub

ERRH:
8700    blnFormOpening = False
8710    DoCmd.Hourglass False
8720    Select Case ERR.Number
        Case Else
8730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8740    End Select
8750    Resume EXITP

End Sub

Private Sub cmdMaint03_KeyDown(KeyCode As Integer, Shift As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint03_KeyDown"

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
          Case vbKeyTab
8870        With Me
8880          intRetVal = 0
8890          .cmdMaint04.SetFocus
8900        End With
8910      End Select
8920    End If

        ' ** Shift keys.
8930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8940      Select Case intRetVal
          Case vbKeyTab
8950        With Me
8960          intRetVal = 0
8970          .cmdMaint02.SetFocus
8980        End With
8990      End Select
9000    End If

EXITP:
9010    KeyCode = intRetVal
9020    Exit Sub

ERRH:
9030    intRetVal = 0
9040    Select Case ERR.Number
        Case Else
9050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9060    End Select
9070    Resume EXITP

End Sub

Private Sub cmdMaint03_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint03_MouseMove"

9110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9120    Exit Sub

ERRH:
9130    Select Case ERR.Number
        Case Else
9140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9150    End Select
9160    Resume EXITP

End Sub

Private Sub cmdMaint04_GotFocus()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint04_GotFocus"

9210    With Me
9220      blnMaint04_Focus = True
9230      .cmdMaint04_box01.Visible = True
9240      .cmdMaint04_box02.Visible = True
9250    End With

EXITP:
9260    Exit Sub

ERRH:
9270    Select Case ERR.Number
        Case Else
9280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9290    End Select
9300    Resume EXITP

End Sub

Private Sub cmdMaint04_Click()
' ** 4. Foreign Exchange.

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint04_Click"

        Dim strDocName As String

9410    With Me
9420      If blnFormOpening = False Then

9430        strDocName = "frmMenu_ForeignExchange"

9440        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
9450        DoEvents

9460        If .chkLoadTimer.Visible = True Then
9470          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
9480          If gblnTimer = True Then
                'gblnSetFocus = True
9490            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
9500          End If
9510        End If
9520        DoEvents

9530        DoCmd.Hourglass True
9540        DoEvents

9550        .TimerInterval = 0&

9560        gblnSetFocus = True
9570        blnJustClose = True: blnFormOpening = True
9580        gstrReturningForm = strDocName
9590        DoCmd.Close acForm, THIS_NAME
9600        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
9610        If gblnGoToReport = True Then
9620          Forms(strDocName).TimerInterval = 50&
9630          blnGoingToReport = False
9640          blnGTR_Emblem = False
9650        End If

9660        DoCmd.Hourglass False

9670      End If
9680    End With

EXITP:
9690    Exit Sub

ERRH:
9700    blnFormOpening = False
9710    DoCmd.Hourglass False
9720    Select Case ERR.Number
        Case Else
9730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9740    End Select
9750    Resume EXITP

End Sub

Private Sub cmdMaint04_KeyDown(KeyCode As Integer, Shift As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint04_KeyDown"

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
9890          If .chkLoadTimer.Visible = True And .cmdMaint10.Visible = True Then
9900            .cmdMaint10.SetFocus
9910          ElseIf .cmdMaint05.Enabled = True Then
9920            .cmdMaint05.SetFocus
9930          Else
9940            .cmdMaint06.SetFocus
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
10030         .cmdMaint03.SetFocus
10040       End With
10050     End Select
10060   End If

EXITP:
10070   KeyCode = intRetVal
10080   Exit Sub

ERRH:
10090   intRetVal = 0
10100   Select Case ERR.Number
        Case Else
10110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10120   End Select
10130   Resume EXITP

End Sub

Private Sub cmdMaint04_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint04_MouseMove"

10210   With Me
10220     If .cmdMaint04_box01.Visible = False Then
10230       .cmdMaint04_box01.Visible = True
10240       .cmdMaint04_box02.Visible = True
10250     End If
10260   End With

EXITP:
10270   Exit Sub

ERRH:
10280   Select Case ERR.Number
        Case Else
10290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10300   End Select
10310   Resume EXITP

End Sub

Private Sub cmdMaint04_LostFocus()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint04_LostFocus"

10410   With Me
10420     .cmdMaint04_box01.Visible = False
10430     .cmdMaint04_box02.Visible = False
10440     blnMaint04_Focus = False
10450   End With

EXITP:
10460   Exit Sub

ERRH:
10470   Select Case ERR.Number
        Case Else
10480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10490   End Select
10500   Resume EXITP

End Sub

Private Sub cmdMaint05_Click()
' ** 5. Check Register.

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint05_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim ctl As Access.Label
        Dim strDocName As String, lngChkRegCoID As Long
        Dim lngChkRegSource As Long
        Dim blnLoadI As Boolean, blnLoadTC As Boolean, blnSkip As Boolean

10610   blnSkip = True
10620   If blnSkip = False Then

10630     With Me
10640       If blnFormOpening = False Then

10650         DoCmd.Hourglass True
10660         DoEvents

10670         .TimerInterval = 0&

10680         Set ctl = .cmdMaint05_lbl
              'ctl.Visible = True

10690         Set dbs = CurrentDb
10700         With dbs

                ' ** Assure the CompanyInformation variables are set.
10710           If gstrCo_Name = vbNullString Then
10720             CoOptions_Read  ' ** Module Function: modStartupFuncs.
10730           End If
10740           DoEvents
10750           ctl.Caption = CStr(1)
10760           DoEvents

                ' ** We're going to borrow this variable for frmCheckRegister.
10770           gstrActNo = gstrCo_Name
10780           If gstrActNo = vbNullString Then gstrActNo = "{empty}"

                ' ** Check for company record.
                ' ** qryCheckRegister_09 (CompanyInformation, just needed fields), linked to tblCheckRegister_Company.
10790           Set qdf = .QueryDefs("qryCheckRegister_10")
10800           Set rst = qdf.OpenRecordset
10810           If rst.BOF = True And rst.EOF = True Then
                  ' ** Not there, so add it.
10820             rst.Close
10830             Set rst = Nothing
10840             Set qdf = Nothing
10850             DoEvents
                  ' ** Append qryCheckRegister_09 (CompanyInformation, just needed fields) to tblCheckRegister_Company.
10860             Set qdf = .QueryDefs("qryCheckRegister_11")
10870             qdf.Execute
10880             Set qdf = Nothing
10890             DoEvents
                  ' ** qryCheckRegister_09 (CompanyInformation, just needed fields), linked to tblCheckRegister_Company.
10900             Set qdf = .QueryDefs("qryCheckRegister_10")
10910             Set rst = qdf.OpenRecordset
10920             rst.MoveFirst
10930             lngChkRegCoID = rst![chkregco_id]
10940             rst.Close
10950             Set rst = Nothing
10960             Set qdf = Nothing
10970             DoEvents
                  ' ** Append qryCheckRegister_12c (qryCheckRegister_12a (tblForm_Graphics, just frmCheckRegister,
                  ' ** frmgfx_alt = 0), linked to qryCheckRegister_12b (tblForm_Graphics, just frmCheckRegister,
                  ' ** frmgfx_alt = 1)), with qryCheckRegister_12d (tblCheckRegister_Company, not in
                  ' ** tblCheckRegister_Graphics), to tblCheckRegister_Graphics; Cartesian.
10980             Set qdf = .QueryDefs("qryCheckRegister_12e")
10990             qdf.Execute
11000             Set qdf = Nothing
11010             DoEvents
11020           Else
                  ' ** Company Already present.
11030             rst.MoveFirst
11040             lngChkRegCoID = rst![chkregco_id]
11050             rst.Close
11060             Set rst = Nothing
11070             Set qdf = Nothing
11080             DoEvents
11090           End If
11100           ctl.Caption = CStr(2)
11110           DoEvents

                ' ** Check subform graphics.
                ' ** qryCheckRegister_14c (tblCheckRegister_Company, with qryCheckRegister_14b (tblForm,
                ' ** just 'frmCheckRegister_Sub'); Cartesian), not in tblCheckRegister_Sub_Graphics.
11120           Set qdf = .QueryDefs("qryCheckRegister_14d")
11130           Set rst = qdf.OpenRecordset
11140           If rst.BOF = True And rst.EOF = True Then
11150             rst.Close
11160             Set rst = Nothing
11170             Set qdf = Nothing
11180           Else
11190             rst.Close
11200             Set rst = Nothing
11210             Set qdf = Nothing
                  ' ** Append qryCheckRegister_14a (tblForm_Graphics, just frmCheckRegister_Sub),
                  ' ** linked to qryCheckRegister_14d (qryCheckRegister_14c (tblCheckRegister_Company,
                  ' ** with qryCheckRegister_14b (tblForm, just 'frmCheckRegister_Sub'); Cartesian),
                  ' ** not in tblCheckRegister_Sub_Graphics), to tblCheckRegister_Sub_Graphics.
11220             Set qdf = .QueryDefs("qryCheckRegister_14e")
11230             qdf.Execute
11240             Set qdf = Nothing
11250           End If
11260           ctl.Caption = CStr(3)
11270           DoEvents
11280           DoEvents

                ' ** Check if they've got a preference for CheckNum vs. Paid.
                ' ** qryCheckRegister_13a (tblPreference_user, just 'opgSource' on 'frmCheckRegister'),
                ' ** linked to qryCheckRegister_13b (tblForm_Control, just acOptionButton's for
                ' ** 'opgSource' on frmCheckRegister, with ctlspec_optionvalue).
11290           Set qdf = .QueryDefs("qryCheckRegister_13c")
11300           Set rst = qdf.OpenRecordset
11310           With rst
11320             If .BOF = True And .EOF = True Then
                    ' ** No preference.
11330               lngChkRegSource = CHKREG_PAID 'CHKREG_NUM
11340             Else
11350               .MoveFirst
11360               lngChkRegSource = CHKREG_PAID '![prefuser_integer]
11370             End If
11380             .Close
11390           End With
11400           Set rst = Nothing
11410           Set qdf = Nothing

                ' ** Check for default Individual accountno = '0' and Total Cash accountno = 'CRTC01'.
11420           blnLoadI = True: blnLoadTC = True

                ' ** tblCheckRegister, just needed fields, by specified [ckrgcoid].
11430           Set qdf = .QueryDefs("qryCheckRegister_05")
11440           With qdf.Parameters
11450             ![ckrgcoid] = lngChkRegCoID
11460           End With
11470           Set rst = qdf.OpenRecordset
11480           With rst
11490             If .BOF = True And .EOF = True Then
11500               blnLoadI = False
11510               blnLoadTC = False
11520               DoEvents
11530             Else
11540               .MoveFirst
11550               .FindFirst "[accountno] = '0'"
11560               blnLoadI = (Not .NoMatch)
11570               .MoveFirst
11580               .FindFirst "[accountno] = 'CRTC01'"
11590               blnLoadTC = (Not .NoMatch)
11600               DoEvents
11610             End If
11620             .Close
11630           End With
11640           Set rst = Nothing
11650           Set qdf = Nothing
11660           ctl.Caption = CStr(4)
11670           DoEvents

                ' ** Load Individual parent record for accountno = '0'.
11680           If blnLoadI = False Then
                  ' ** Append qryCheckRegister_06a (tblCheckingType, as new tblCheckRegister dummy
                  ' ** record; for Individual), linked to qryCheckRegister_06b (tblForm_Graphics,
                  ' ** just frmCheckRegister), to tblCheckRegister, by specified [ckrgcoid], [crsrc]; for Individual.
11690             Set qdf = .QueryDefs("qryCheckRegister_06c")
11700             With qdf.Parameters
11710               ![ckrgcoid] = lngChkRegCoID
11720               ![crsrc] = lngChkRegSource
11730             End With
11740             qdf.Execute dbFailOnError
11750             Set qdf = Nothing
11760           End If
11770           DoEvents
11780           ctl.Caption = CStr(5)
11790           DoEvents

                ' ** Load Total Cash parent record for accountno = 'CRTC01'.
11800           If blnLoadTC = False Then
                  ' ** Append qryCheckRegister_06d (tblCheckingType, as new tblCheckRegister 'CRTC01'
                  ' ** record; for Total Cash), linked to qryCheckRegister_06e (tblForm_Graphics,
                  ' ** just frmCheckRegister), to tblCheckRegister, by specified [ckrgcoid], [crsrc]; for Total Cash.
11810             Set qdf = .QueryDefs("qryCheckRegister_06f")
11820             With qdf.Parameters
11830               ![ckrgcoid] = lngChkRegCoID
11840               ![crsrc] = lngChkRegSource
11850             End With
11860             qdf.Execute dbFailOnError
11870             Set qdf = Nothing
11880           End If
11890           DoEvents
11900           ctl.Caption = CStr(6)
11910           DoEvents

11920           blnLoadI = False: blnLoadTC = False

                ' ** All Individuals with checks.
11930           Select Case lngChkRegSource
                Case CHKREG_NUM
                  ' ** qryCheckRegister_Sub_03a (qryCheckRegister_Sub_02a (qryCheckRegister_Sub_01c
                  ' ** (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                  ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to
                  ' ** Account, with add'l fields), grouped, with tblCheckingType; Cartesian; for
                  ' ** Individual), not in tblCheckRegister, by specified [ckrgcoid]; for Individual; CheckNum.
11940             Set qdf = .QueryDefs("qryCheckRegister_Sub_03b")
11950           Case CHKREG_PAID
                  ' ** qryCheckRegister_Sub_03e (xx), not in tblCheckRegister, by specified [ckrgcoid]; for Individual; Paid.
11960             Set qdf = .QueryDefs("qryCheckRegister_Sub_03f")
11970           End Select
11980           With qdf.Parameters
11990             ![ckrgcoid] = lngChkRegCoID
12000           End With
12010           Set rst = qdf.OpenRecordset
12020           If rst.BOF = True And rst.EOF = True Then
                  ' ** Nothing not there.
12030           Else
12040             blnLoadI = True
12050           End If
12060           rst.Close
12070           Set rst = Nothing
12080           Set qdf = Nothing
12090           DoEvents
12100           ctl.Caption = CStr(7)
12110           DoEvents

                ' ** All Total Cash with checks.
12120           Select Case lngChkRegSource
                Case CHKREG_NUM
                  ' ** qryCheckRegister_Sub_05a (qryCheckRegister_Sub_02a (qryCheckRegister_Sub_01c
                  ' ** (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                  ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to
                  ' ** Account, with add'l fields), grouped, with tblCheckingType; Cartesian;
                  ' ** for Total Cash), not in tblCheckRegister, by specified [ckrgcoid]; for Total Cash.
12130             Set qdf = .QueryDefs("qryCheckRegister_Sub_05b")
12140           Case CHKREG_PAID
                  ' ** qryCheckRegister_Sub_05e (xx), not in tblCheckRegister, by specified [ckrgcoid]; for Total Cash; Paid.
12150             Set qdf = .QueryDefs("qryCheckRegister_Sub_05f")
12160           End Select
12170           With qdf.Parameters
12180             ![ckrgcoid] = lngChkRegCoID
12190           End With
12200           Set rst = qdf.OpenRecordset
12210           If rst.BOF = True And rst.EOF = True Then
                  ' ** Nothing not there.
12220           Else
12230             blnLoadTC = True
12240           End If
12250           rst.Close
12260           Set rst = Nothing
12270           Set qdf = Nothing
12280           DoEvents
12290           ctl.Caption = CStr(8)
12300           DoEvents

                ' ** Load parent Individual records.
12310           If blnLoadI = True Then
12320             Select Case lngChkRegSource
                  Case CHKREG_NUM
                    ' ** Append qryCheckRegister_Sub_03b (qryCheckRegister_Sub_03a (qryCheckRegister_Sub_02a
                    ' ** (qryCheckRegister_Sub_01c (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                    ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to Account, with
                    ' ** add'l fields), grouped, with tblCheckingType; Cartesian; for Individual), not in tblCheckRegister,
                    ' ** by specified [ckrgcoid]; for Individual) to tblCheckRegister, by specified [crsrc]; for Individual; CheckNum.
12330               Set qdf = .QueryDefs("qryCheckRegister_Sub_03d")
12340             Case CHKREG_PAID
                    ' ** Append qryCheckRegister_Sub_03f (xx), linked to qryCheckRegister_Sub_03g (xx), to tblCheckRegister, by specified [crsrc]; for Individual; Paid.
12350               Set qdf = .QueryDefs("qryCheckRegister_Sub_03h")
12360             End Select
12370             With qdf.Parameters
12380               ![ckrgcoid] = lngChkRegCoID
12390               ![crsrc] = lngChkRegSource
12400             End With
12410             qdf.Execute
12420             Set qdf = Nothing
12430           End If
12440           DoEvents
12450           ctl.Caption = CStr(9)
12460           DoEvents

                ' ** Load parent Total Cash record.
12470           If blnLoadTC = True Then
12480             Select Case lngChkRegSource
                  Case CHKREG_NUM
                    ' ** Append qryCheckRegister_Sub_05b (qryCheckRegister_Sub_05a (qryCheckRegister_Sub_02a
                    ' ** (qryCheckRegister_Sub_01c (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                    ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to Account, with
                    ' ** add'l fields), grouped, with tblCheckingType; Cartesian; for Total Cash), not in tblCheckRegister,
                    ' ** by specified [ckrgcoid]; for Total Cash) to tblCheckRegister, by specified [crsrc]; for Total Cash; CheckNum.
12490               Set qdf = .QueryDefs("qryCheckRegister_Sub_05d")
12500             Case CHKREG_PAID
                    ' ** Append qryCheckRegister_Sub_05f (xx), linked to qryCheckRegister_Sub_05g (xx), to tblCheckRegister; for Total Cash; Paid.
12510               Set qdf = .QueryDefs("qryCheckRegister_Sub_05h")
12520             End Select
12530             With qdf.Parameters
12540               ![ckrgcoid] = lngChkRegCoID
12550               ![crsrc] = lngChkRegSource
12560             End With
12570             qdf.Execute
12580             Set qdf = Nothing
12590           End If
12600           DoEvents
12610           ctl.Caption = CStr(10)
12620           DoEvents

12630           blnLoadI = False: blnLoadTC = False

                ' ** Check for cancelled changes.
                ' ** tblCheckRegister_Detail, just chkregdet_canceled = True.
12640           Set qdf = .QueryDefs("qryCheckRegister_Sub_10a")
12650           Set rst = qdf.OpenRecordset
12660           If rst.BOF = True And rst.EOF = True Then
                  ' ** No cancels.
12670           Else
12680             blnLoadI = True
12690           End If
12700           rst.Close
12710           Set qdf = Nothing
                ' ** tblCheckRegister_Set_Detail, just chkregdet_canceled = True.
12720           Set qdf = .QueryDefs("qryCheckRegister_Sub_11a")
12730           Set rst = qdf.OpenRecordset
12740           If rst.BOF = True And rst.EOF = True Then
                  ' ** No cancels.
12750           Else
12760             blnLoadTC = True
12770           End If
12780           rst.Close
12790           Set qdf = Nothing

                ' ** Reverse any canceled changes in tblCheckRegister_Detail.
12800           If blnLoadI = True Then
                  ' ** Update qryCheckRegister_Sub_10a (tblCheckRegister_Detail, just chkregdet_canceled = True).
12810             Set qdf = .QueryDefs("qryCheckRegister_Sub_10b")
12820             qdf.Execute
12830             Set qdf = Nothing
12840           End If

                ' ** Reverse any canceled changes in tblCheckRegister_Set_Detail.
12850           If blnLoadTC = True Then
                  ' ** Update qryCheckRegister_Sub_11a (tblCheckRegister_Set_Detail, just chkregdet_canceled = True).
12860             Set qdf = .QueryDefs("qryCheckRegister_Sub_11b")
12870             qdf.Execute
12880             Set qdf = Nothing
12890           End If

12900           blnLoadI = False: blnLoadTC = False

                ' ** Check preexisting Individual checks.
12910           Select Case lngChkRegSource
                Case CHKREG_NUM
                  ' ** qryCheckRegister_Sub_04a (qryCheckRegister_Sub_02a (qryCheckRegister_Sub_01c
                  ' ** (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                  ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to
                  ' ** Account, with add'l fields), linked to tblCheckRegister, by specified [ckrgcoid]; for Individual),
                  ' ** linked to tblCheckRegister_Detail, just differing CheckNum, CheckPaid; for Individual; CheckNum.
12920             Set qdf = .QueryDefs("qryCheckRegister_Sub_04d")
12930           Case CHKREG_PAID
                  ' ** qryCheckRegister_Sub_04i (xx), linked to tblCheckRegister_Detail, just differing CheckNum, CheckPaid; for Individual; Paid.
12940             Set qdf = .QueryDefs("qryCheckRegister_Sub_04l")
12950           End Select
12960           With qdf.Parameters
12970             ![ckrgcoid] = lngChkRegCoID
12980           End With
12990           Set rst = qdf.OpenRecordset
13000           If rst.BOF = True And rst.EOF = True Then
                  ' ** No changes.
13010           Else
13020             blnLoadI = True
13030           End If
13040           rst.Close
13050           Set rst = Nothing
13060           Set qdf = Nothing
13070           DoEvents
13080           ctl.Caption = CStr(11)
13090           DoEvents

                ' ** Check preexisting Total Cash checks.
13100           Select Case lngChkRegSource
                Case CHKREG_NUM
                  ' ** qryCheckRegister_Sub_06a (qryCheckRegister_Sub_02a (qryCheckRegister_Sub_01c
                  ' ** (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                  ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to
                  ' ** Account, with add'l fields), linked to tblCheckRegister, by specified [ckrgcoid]; for Total Cash),
                  ' ** linked to tblCheckRegister_Detail, just differing CheckNum, CheckPaid; for Total Cash; CheckNum.
13110             Set qdf = .QueryDefs("qryCheckRegister_Sub_06d")
13120           Case CHKREG_PAID
                  ' ** qryCheckRegister_Sub_06i (xx), linked to tblCheckRegister_Detail, just differing CheckNum, CheckPaid; for Total Cash; Paid.
13130             Set qdf = .QueryDefs("qryCheckRegister_Sub_06l")
13140           End Select
13150           With qdf.Parameters
13160             ![ckrgcoid] = lngChkRegCoID
13170           End With
13180           Set rst = qdf.OpenRecordset
13190           If rst.BOF = True And rst.EOF = True Then
                  ' ** No changes.
13200           Else
13210             blnLoadTC = True
13220           End If
13230           rst.Close
13240           Set rst = Nothing
13250           Set qdf = Nothing
13260           DoEvents
13270           ctl.Caption = CStr(12)
13280           DoEvents

                ' ** Update preexisting Individual checks.
13290           If blnLoadI = True Then
13300             Select Case lngChkRegSource
                  Case CHKREG_NUM
                    ' ** Empty tmpCheckRegister_01; CheckNum.
13310               Set qdf = .QueryDefs("qryCheckRegister_Sub_04e")
13320             Case CHKREG_PAID
                    ' ** Empty tmpCheckRegister_01; Paid.
13330               Set qdf = .QueryDefs("qryCheckRegister_Sub_04m")
13340             End Select
13350             qdf.Execute
13360             Set qdf = Nothing
13370             Select Case lngChkRegSource
                  Case CHKREG_NUM
                    ' ** Append qryCheckRegister_Sub_04d (qryCheckRegister_Sub_04a (qryCheckRegister_Sub_02a
                    ' ** (qryCheckRegister_Sub_01c (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                    ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to Account, with
                    ' ** add'l fields), linked to tblCheckRegister, by specified [ckrgcoid]; for Individual), linked to
                    ' ** tblCheckRegister_Detail, just differing CheckNum, CheckPaid; for Individual) to tmpCheckRegister_01; CheckNum.
13380               Set qdf = .QueryDefs("qryCheckRegister_Sub_04f")
13390             Case CHKREG_PAID
                    ' ** Append qryCheckRegister_Sub_04l (xx) to tmpCheckRegister_01; Paid.
13400               Set qdf = .QueryDefs("qryCheckRegister_Sub_04n")
13410             End Select
13420             With qdf.Parameters
13430               ![ckrgcoid] = lngChkRegCoID
13440             End With
13450             qdf.Execute
13460             Set qdf = Nothing
13470             Select Case lngChkRegSource
                  Case CHKREG_NUM
                    ' ** Update qryCheckRegister_Sub_04g (tblCheckRegister_Detail, linked
                    ' ** to tmpCheckRegister_01, with checknum_new, checkpaid_new); CheckNum.
13480               Set qdf = .QueryDefs("qryCheckRegister_Sub_04h")
13490             Case CHKREG_PAID
                    ' ** Update qryCheckRegister_Sub_04o (xx); Paid.
13500               Set qdf = .QueryDefs("qryCheckRegister_Sub_04p")
13510             End Select
13520             qdf.Execute
13530             Set qdf = Nothing
13540           End If
13550           DoEvents
13560           ctl.Caption = CStr(13)
13570           DoEvents

                ' ** Update preexisting Total Cash checks.
13580           If blnLoadTC = True Then
13590             Select Case lngChkRegSource
                  Case CHKREG_NUM
                    ' ** Empty tmpCheckRegister_01; CheckNum.
13600               Set qdf = .QueryDefs("qryCheckRegister_Sub_06e")
13610             Case CHKREG_PAID
                    ' ** Empty tmpCheckRegister_01; Paid.
13620               Set qdf = .QueryDefs("qryCheckRegister_Sub_06m")
13630             End Select
13640             qdf.Execute
13650             Set qdf = Nothing
13660             Select Case lngChkRegSource
                  Case CHKREG_NUM
                    ' ** Append qryCheckRegister_Sub_06d (qryCheckRegister_Sub_06a (qryCheckRegister_Sub_02a
                    ' ** (qryCheckRegister_Sub_01c (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                    ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to Account, with
                    ' ** add'l fields), linked to tblCheckRegister, by specified [ckrgcoid]; for Total Cash), linked to
                    ' ** tblCheckRegister_Detail, just differing CheckNum, CheckPaid; for Total Cash) to tmpCheckRegister_01; CheckNum.
13670               Set qdf = .QueryDefs("qryCheckRegister_Sub_06f")
13680             Case CHKREG_PAID
                    ' ** Append qryCheckRegister_Sub_06l (xx) to tmpCheckRegister_01; Paid.
13690               Set qdf = .QueryDefs("qryCheckRegister_Sub_06n")
13700             End Select
13710             With qdf.Parameters
13720               ![ckrgcoid] = lngChkRegCoID
13730             End With
13740             qdf.Execute
13750             Set qdf = Nothing
13760             Select Case lngChkRegSource
                  Case CHKREG_NUM
                    ' ** Update qryCheckRegister_Sub_06g (tblCheckRegister_Detail, linked
                    ' ** to tmpCheckRegister_01, with checknum_new, checkpaid_new); CheckNum.
13770               Set qdf = .QueryDefs("qryCheckRegister_Sub_06h")
13780             Case CHKREG_PAID
                    ' ** Update qryCheckRegister_Sub_06o; Paid.
13790               Set qdf = .QueryDefs("qryCheckRegister_Sub_06p")
13800             End Select
13810             qdf.Execute
13820             Set qdf = Nothing
13830           End If
13840           DoEvents
13850           ctl.Caption = CStr(14)
13860           DoEvents

13870           blnLoadI = False: blnLoadTC = False

                ' ** All Individual checks.
13880           Select Case lngChkRegSource
                Case CHKREG_NUM
                  ' ** qryCheckRegister_Sub_04a (qryCheckRegister_Sub_02a (qryCheckRegister_Sub_01c
                  ' ** (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                  ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to
                  ' ** Account, with add'l fields), linked to tblCheckRegister, by specified [ckrgcoid]; for Individual),
                  ' ** not in tblCheckRegister_Detail; for Individual; CheckNum.
13890             Set qdf = .QueryDefs("qryCheckRegister_Sub_04b")
13900           Case CHKREG_PAID
                  ' ** qryCheckRegister_Sub_04i (xx), not in tblCheckRegister_Detail; for Individual; Paid.
13910             Set qdf = .QueryDefs("qryCheckRegister_Sub_04j")
13920           End Select
13930           With qdf.Parameters
13940             ![ckrgcoid] = lngChkRegCoID
13950           End With
13960           Set rst = qdf.OpenRecordset
13970           If rst.BOF = True And rst.EOF = True Then
                  ' ** Nothing not there.
13980           Else
13990             blnLoadI = True
14000           End If
14010           rst.Close
14020           Set rst = Nothing
14030           Set qdf = Nothing
14040           DoEvents
14050           ctl.Caption = CStr(15)
14060           DoEvents

                ' ** All Total Cash checks.
14070           Select Case lngChkRegSource
                Case CHKREG_NUM
                  ' ** qryCheckRegister_Sub_06a (qryCheckRegister_Sub_02a (qryCheckRegister_Sub_01c
                  ' ** (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                  ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to
                  ' ** Account, with add'l fields), linked to tblCheckRegister, by specified [ckrgcoid]; for Total Cash),
                  ' ** not in tblCheckRegister_Detail; for Total Cash; CheckNum.
14080             Set qdf = .QueryDefs("qryCheckRegister_Sub_06b")
14090           Case CHKREG_PAID
                  ' ** qryCheckRegister_Sub_06i (xx), not in tblCheckRegister_Detail; for Total Cash; Paid.
14100             Set qdf = .QueryDefs("qryCheckRegister_Sub_06j")
14110           End Select
14120           With qdf.Parameters
14130             ![ckrgcoid] = lngChkRegCoID
14140           End With
14150           Set rst = qdf.OpenRecordset
14160           If rst.BOF = True And rst.EOF = True Then
                  ' ** Nothing not there.
14170           Else
14180             blnLoadTC = True
14190           End If
14200           rst.Close
14210           Set rst = Nothing
14220           Set qdf = Nothing
14230           DoEvents
14240           ctl.Caption = CStr(16)
14250           DoEvents

                ' ** Load Individual checks.
14260           If blnLoadI = True Then
14270             Select Case lngChkRegSource
                  Case CHKREG_NUM
                    ' ** Append qryCheckRegister_Sub_04b (qryCheckRegister_Sub_04a (qryCheckRegister_Sub_02a
                    ' ** (qryCheckRegister_Sub_01c (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                    ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to Account, with
                    ' ** add'l fields), linked to tblCheckRegister; for Individual), not in tblCheckRegister_Detail, by specified [ckrgcoid];
                    ' ** for Individual) to tblCheckRegister_Detail; for Individual; CheckNum.
14280               Set qdf = .QueryDefs("qryCheckRegister_Sub_04c")
14290             Case CHKREG_PAID
                    ' ** Append qryCheckRegister_Sub_04j (xx) to tblCheckRegister_Detail; for Individual; Paid.
14300               Set qdf = .QueryDefs("qryCheckRegister_Sub_04k")
14310             End Select
14320             With qdf.Parameters
14330               ![ckrgcoid] = lngChkRegCoID
14340             End With
14350             qdf.Execute
14360             Set qdf = Nothing
14370           End If
14380           DoEvents
14390           ctl.Caption = CStr(17)
14400           DoEvents

                ' ** Load Total Cash checks.
14410           If blnLoadTC = True Then
14420             Select Case lngChkRegSource
                  Case CHKREG_NUM
                    ' ** Append qryCheckRegister_Sub_06b (qryCheckRegister_Sub_06a (qryCheckRegister_Sub_02a
                    ' ** (qryCheckRegister_Sub_01c (Union of qryCheckRegister_Sub_01a (Ledger, just CheckNum <> Null),
                    ' ** qryCheckRegister_Sub_01b (LedgerArchive, just CheckNum <> Null)), linked to Account, with
                    ' ** add'l fields), linked to tblCheckRegister; for Total Cash), not in tblCheckRegister_Detail, by specified [ckrgcoid];
                    ' ** for Total Cash) to tblCheckRegister_Detail; for Total Cash; CheckNum.
14430               Set qdf = .QueryDefs("qryCheckRegister_Sub_06c")
14440             Case CHKREG_PAID
                    ' ** Append qryCheckRegister_Sub_06j (xx) to tblCheckRegister_Detail; for Total Cash; Paid.
14450               Set qdf = .QueryDefs("qryCheckRegister_Sub_06k")
14460             End Select
14470             With qdf.Parameters
14480               ![ckrgcoid] = lngChkRegCoID
14490             End With
14500             qdf.Execute
14510             Set qdf = Nothing
14520           End If
14530           DoEvents
14540           ctl.Caption = CStr(18)
14550           DoEvents

14560           .Close
14570         End With
14580         Set dbs = Nothing
14590         DoEvents
14600         ctl.Caption = CStr(19)
14610         DoEvents

14620         gblnSetFocus = True
14630         blnJustClose = True: blnFormOpening = True
14640         strDocName = "frmCheckRegister"
14650         gstrReturningForm = strDocName
14660         DoCmd.Close acForm, THIS_NAME
14670         DoCmd.OpenForm strDocName, , , , , , THIS_NAME

14680         DoCmd.Hourglass False

14690       End If
14700     End With

14710   End If  ' ** blnSkip.

EXITP:
14720   Set ctl = Nothing
14730   Set rst = Nothing
14740   Set qdf = Nothing
14750   Set dbs = Nothing
14760   Exit Sub

ERRH:
14770   blnFormOpening = False
14780   DoCmd.Hourglass False
14790   Select Case ERR.Number
        Case Else
14800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14810   End Select
14820   Resume EXITP

End Sub

Private Sub cmdMaint05_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint05_KeyDown"

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
          Case vbKeyTab
14970       With Me
14980         intRetVal = 0
14990         .cmdMaint06.SetFocus
15000       End With
15010     End Select
15020   End If

        ' ** Shift keys.
15030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15040     Select Case intRetVal
          Case vbKeyTab
15050       With Me
15060         intRetVal = 0
15070         .cmdMaint04.SetFocus
15080       End With
15090     End Select
15100   End If

EXITP:
15110   KeyCode = intRetVal
15120   Exit Sub

ERRH:
15130   intRetVal = 0
15140   Select Case ERR.Number
        Case Else
15150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15160   End Select
15170   Resume EXITP

End Sub

Private Sub cmdMaint05_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint05_MouseMove"

15210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15220   Exit Sub

ERRH:
15230   Select Case ERR.Number
        Case Else
15240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15250   End Select
15260   Resume EXITP

End Sub

Private Sub cmdMaint06_GotFocus()

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint06_GotFocus"

15310   With Me
15320     blnMaint06_Focus = True
15330     .cmdMaint06_box01.Visible = True
15340     .cmdMaint06_box02.Visible = True
15350   End With

EXITP:
15360   Exit Sub

ERRH:
15370   Select Case ERR.Number
        Case Else
15380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15390   End Select
15400   Resume EXITP

End Sub

Private Sub cmdMaint06_Click()
' ** 6. System Tables Menu.

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint06_Click"

        Dim strDocName As String

15510   With Me
15520     If blnFormOpening = False Then

15530       strDocName = "frmMenu_SystemTable"

15540       FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
15550       DoEvents

15560       If .chkLoadTimer.Visible = True Then
15570         chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
15580         If gblnTimer = True Then
                'gblnSetFocus = True
15590           DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
15600         End If
15610       End If
15620       DoEvents

15630       DoCmd.Hourglass True
15640       DoEvents

15650       .TimerInterval = 0&

15660       gblnSetFocus = True
15670       blnJustClose = True: blnFormOpening = True
15680       gstrReturningForm = strDocName
15690       DoCmd.Close acForm, THIS_NAME
15700       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
15710       If gblnGoToReport = True Then
15720         Forms(strDocName).TimerInterval = 50&
15730         blnGoingToReport = False
15740         blnGTR_Emblem = False
15750       End If

15760       DoCmd.Hourglass False

15770     End If
15780   End With

EXITP:
15790   Exit Sub

ERRH:
15800   blnFormOpening = False
15810   DoCmd.Hourglass False
15820   Select Case ERR.Number
        Case Else
15830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15840   End Select
15850   Resume EXITP

End Sub

Private Sub cmdMaint06_KeyDown(KeyCode As Integer, Shift As Integer)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint06_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15920   intShiftDown = (Shift And acShiftMask) > 0
15930   intAltDown = (Shift And acAltMask) > 0
15940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15960     Select Case intRetVal
          Case vbKeyTab
15970       With Me
15980         intRetVal = 0
15990         .cmdMaint07.SetFocus
16000       End With
16010     End Select
16020   End If

        ' ** Shift keys.
16030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16040     Select Case intRetVal
          Case vbKeyTab
16050       With Me
16060         intRetVal = 0
16070         If .chkLoadTimer.Visible = True And .cmdMaint10.Visible = True Then
16080           .cmdMaint10.SetFocus
16090         ElseIf .cmdMaint05.Enabled = True Then
16100           .cmdMaint05.SetFocus
16110         Else
16120           .cmdMaint04.SetFocus
16130         End If
16140       End With
16150     End Select
16160   End If

EXITP:
16170   KeyCode = intRetVal
16180   Exit Sub

ERRH:
16190   intRetVal = 0
16200   Select Case ERR.Number
        Case Else
16210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16220   End Select
16230   Resume EXITP

End Sub

Private Sub cmdMaint06_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint06_MouseMove"

16310   With Me
16320     If .cmdMaint06_box01.Visible = False Then
16330       .cmdMaint06_box01.Visible = True
16340       .cmdMaint06_box02.Visible = True
16350     End If
16360   End With

EXITP:
16370   Exit Sub

ERRH:
16380   Select Case ERR.Number
        Case Else
16390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16400   End Select
16410   Resume EXITP

End Sub

Private Sub cmdMaint06_LostFocus()

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint06_LostFocus"

16510   With Me
16520     .cmdMaint06_box01.Visible = False
16530     .cmdMaint06_box02.Visible = False
16540     blnMaint06_Focus = False
16550   End With

EXITP:
16560   Exit Sub

ERRH:
16570   Select Case ERR.Number
        Case Else
16580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16590   End Select
16600   Resume EXITP

End Sub

Private Sub cmdMaint07_Click()
' ** 7. Keyboard Shortcuts.

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint07_Click"

        Dim strDocName As String

16710   With Me
16720     If blnFormOpening = False Then

16730       DoCmd.Hourglass True
16740       DoEvents

16750       .TimerInterval = 0&

16760       gblnSetFocus = True
16770       blnJustClose = True: blnFormOpening = True
16780       strDocName = "frmXAdmin_Shortcut"
16790       gstrReturningForm = strDocName
16800       DoCmd.Close acForm, THIS_NAME
16810       DoCmd.OpenForm strDocName, , , , , , THIS_NAME

16820       DoCmd.Hourglass False

16830     End If
16840   End With

EXITP:
16850   Exit Sub

ERRH:
16860   blnFormOpening = False
16870   DoCmd.Hourglass False
16880   Select Case ERR.Number
        Case Else
16890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16900   End Select
16910   Resume EXITP

End Sub

Private Sub cmdMaint07_KeyDown(KeyCode As Integer, Shift As Integer)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint07_KeyDown"

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
          Case vbKeyTab
17070       With Me
17080         intRetVal = 0
17090         .cmdMaint08.SetFocus
17100       End With
17110     End Select
17120   End If

        ' ** Shift keys.
17130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17140     Select Case intRetVal
          Case vbKeyTab
17150       With Me
17160         intRetVal = 0
17170         .cmdMaint06.SetFocus
17180       End With
17190     End Select
17200   End If

EXITP:
17210   KeyCode = intRetVal
17220   Exit Sub

ERRH:
17230   intRetVal = 0
17240   Select Case ERR.Number
        Case Else
17250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17260   End Select
17270   Resume EXITP

End Sub

Private Sub cmdMaint07_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint07_MouseMove"

17310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17320   Exit Sub

ERRH:
17330   Select Case ERR.Number
        Case Else
17340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17350   End Select
17360   Resume EXITP

End Sub

Private Sub cmdMaint08_Click()
' ** 8. Error Log.

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint08_Click"

        Dim strDocName As String

17410   With Me
17420     If blnFormOpening = False Then

17430       DoCmd.Hourglass True
17440       DoEvents

17450       .TimerInterval = 0&

17460       gblnSetFocus = True
17470       blnJustClose = True: blnFormOpening = True
17480       strDocName = "frmErrorLog"
17490       gstrReturningForm = strDocName
17500       DoCmd.Close acForm, THIS_NAME
17510       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
17520       If gblnGoToReport = True Then
17530         Forms(strDocName).TimerInterval = 50&
17540         blnGoingToReport = False
17550         blnGTR_Emblem = False
17560       End If

17570       DoCmd.Hourglass False

17580     End If
17590   End With

EXITP:
17600   Exit Sub

ERRH:
17610   blnFormOpening = False
17620   DoCmd.Hourglass False
17630   Select Case ERR.Number
        Case Else
17640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17650   End Select
17660   Resume EXITP

End Sub

Private Sub cmdMaint08_KeyDown(KeyCode As Integer, Shift As Integer)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint08_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17720   intShiftDown = (Shift And acShiftMask) > 0
17730   intAltDown = (Shift And acAltMask) > 0
17740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17760     Select Case intRetVal
          Case vbKeyTab
17770       With Me
17780         intRetVal = 0
17790         If .cmdMaint09.Visible = True And .cmdMaint09.Enabled = True Then
17800           .cmdMaint09.SetFocus
17810         ElseIf .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
17820           .chkLoadTimer.SetFocus
17830         Else
17840           .cmdClose.SetFocus
17850         End If
17860       End With
17870     End Select
17880   End If

        ' ** Shift keys.
17890   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17900     Select Case intRetVal
          Case vbKeyTab
17910       With Me
17920         intRetVal = 0
17930         .cmdMaint07.SetFocus
17940       End With
17950     End Select
17960   End If

EXITP:
17970   KeyCode = intRetVal
17980   Exit Sub

ERRH:
17990   intRetVal = 0
18000   Select Case ERR.Number
        Case Else
18010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18020   End Select
18030   Resume EXITP

End Sub

Private Sub cmdMaint08_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint08_MouseMove"

18110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
18120   Exit Sub

ERRH:
18130   Select Case ERR.Number
        Case Else
18140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18150   End Select
18160   Resume EXITP

End Sub

Private Sub cmdMaint09_GotFocus()

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint09_GotFocus"

18210   With Me
18220     blnMaint09_Focus = True
18230     .cmdMaint09_box01.Visible = True
18240     .cmdMaint09_box02.Visible = True
18250   End With

EXITP:
18260   Exit Sub

ERRH:
18270   Select Case ERR.Number
        Case Else
18280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18290   End Select
18300   Resume EXITP

End Sub

Private Sub cmdMaint09_Click()
' ** 9. Administration Menu.

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint09_Click"

        Dim strMacro As String
        Dim strPath As String, strExt As String, strVer As String
        Dim blnContinue As Boolean

18410   With Me
18420     If blnFormOpening = False Then

18430       DoCmd.Hourglass True
18440       DoEvents

18450       .TimerInterval = 0&

18460       blnContinue = True

18470       If IsTAOpen(False, "Admin") = False Then  ' ** Module Function: modStartupFuncs.
18480         PathCheck  ' ** Module Function: modXAdminFuncs.
18490         strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
18500         strExt = Parse_Ext(CurrentAppName)  ' ** Module Functions: modFileUtilities.
18510         If FileExists(strPath & LNK_SEP & "TrstXAdm." & strExt) = True Then  ' ** Module Function: modFileUtilities.
18520           strExt = "_" & strExt
18530           If InStr(strPath, gstrDir_Dev) > 0 And InStr(strPath, "VictorGCS") > 0 Then
18540             strMacro = "mcrXAdmin_Work" & strExt
18550           Else
18560             strVer = Application.Version  '11.0
18570             If InStr(strVer, ".") > 0 Then strVer = Left(strVer, (InStr(strVer, ".") - 1))
18580             strMacro = "mcrXAdmin"
18590             If InStr(strPath, "x86") > 0 Then
18600               If InStr(strPath, "Delta Data2") > 0 Then
                      ' ** mcrXAdmin_x86_DD2_Acc11
                      ' ** mcrXAdmin_x86_DD2_Acc12
                      ' ** mcrXAdmin_x86_DD2_Acc14
18610                 strMacro = strMacro & "_x86_DD2_Acc" & strVer & strExt
18620               ElseIf InStr(strPath, "Delta Data") > 0 Then
                      ' ** mcrXAdmin_x86_DD_Acc11
                      ' ** mcrXAdmin_x86_DD_Acc12
                      ' ** mcrXAdmin_x86_DD_Acc14
18630                 strMacro = strMacro & "_x86_DD_Acc" & strVer & strExt
18640               Else
18650                 Beep
18660                 MsgBox "Though TrstXAdm is present, its path is not accommodated.", vbCritical + vbOKOnly, "Trust Administration Unavailable"
18670               End If
18680             Else
18690               If InStr(strPath, "Delta Data2") > 0 Then
                      ' ** mcrXAdmin_DD2_Acc11
                      ' ** mcrXAdmin_DD2_Acc12
                      ' ** mcrXAdmin_DD2_Acc14
18700                 strMacro = strMacro & "_DD2_Acc" & strVer & strExt
18710               ElseIf InStr(strPath, "Delta Data") > 0 Then
                      ' ** mcrXAdmin_DD_Acc11
                      ' ** mcrXAdmin_DD_Acc12
                      ' ** mcrXAdmin_DD_Acc14
18720                 strMacro = strMacro & "_DD_Acc" & strVer & strExt
18730               Else
18740                 Beep
18750                 MsgBox "Though TrstXAdm is present, its path is not accommodated.", vbCritical + vbOKOnly, "Trust Administration Unavailable"
18760               End If
18770             End If
18780           End If
18790           If blnContinue = True Then
18800             DoCmd.RunMacro strMacro
18810           End If
18820         Else
18830           Beep
18840           MsgBox "Trust Administration cannot be found.", vbCritical + vbOKOnly, "File Not Found"
18850         End If
18860       Else
18870         Beep
18880         MsgBox "Trust administration is already open", vbInformation + vbOKOnly, "Trust Administration Already Open"
18890       End If

18900       DoCmd.Hourglass False

18910     End If
18920   End With

EXITP:
18930   Exit Sub

ERRH:
18940   blnFormOpening = False
18950   DoCmd.Hourglass False
18960   Select Case ERR.Number
        Case Else
18970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18980   End Select
18990   Resume EXITP

End Sub

Private Sub cmdMaint09_KeyDown(KeyCode As Integer, Shift As Integer)

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint09_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19020   intShiftDown = (Shift And acShiftMask) > 0
19030   intAltDown = (Shift And acAltMask) > 0
19040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19060     Select Case intRetVal
          Case vbKeyTab
19070       With Me
19080         intRetVal = 0
19090         If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
19100           .chkLoadTimer.SetFocus
19110         Else
19120           .cmdClose.SetFocus
19130         End If
19140       End With
19150     End Select
19160   End If

        ' ** Shift keys.
19170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19180     Select Case intRetVal
          Case vbKeyTab
19190       With Me
19200         intRetVal = 0
19210         .cmdMaint08.SetFocus
19220       End With
19230     End Select
19240   End If

EXITP:
19250   KeyCode = intRetVal
19260   Exit Sub

ERRH:
19270   intRetVal = 0
19280   Select Case ERR.Number
        Case Else
19290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19300   End Select
19310   Resume EXITP

End Sub

Private Sub cmdMaint09_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint09_MouseMove"

19410   With Me
19420     If .cmdMaint09_box01.Visible = False Then
19430       .cmdMaint09_box01.Visible = True
19440       .cmdMaint09_box02.Visible = True
19450     End If
19460   End With

EXITP:
19470   Exit Sub

ERRH:
19480   Select Case ERR.Number
        Case Else
19490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19500   End Select
19510   Resume EXITP

End Sub

Private Sub cmdMaint09_LostFocus()

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint09_LostFocus"

19610   With Me
19620     .cmdMaint09_box01.Visible = False
19630     .cmdMaint09_box02.Visible = False
19640     blnMaint09_Focus = False
19650   End With

EXITP:
19660   Exit Sub

ERRH:
19670   Select Case ERR.Number
        Case Else
19680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19690   End Select
19700   Resume EXITP

End Sub

Private Sub cmdMaint10_Click()

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint10_Click"

        Dim strDocName As String

19810   With Me
19820     If blnFormOpening = False Then

19830       DoCmd.Hourglass True
19840       DoEvents

19850       FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.

19860       .TimerInterval = 0&

19870       gblnSetFocus = True
19880       blnJustClose = True: blnFormOpening = True
19890       strDocName = "frmLoadTimer_Mgmt"
19900       gstrReturningForm = strDocName
19910       DoCmd.Close acForm, THIS_NAME
19920       DoCmd.OpenForm strDocName, , , , , , THIS_NAME

19930       DoCmd.Hourglass False

19940     End If
19950   End With

EXITP:
19960   Exit Sub

ERRH:
19970   Select Case ERR.Number
        Case Else
19980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19990   End Select
20000   Resume EXITP

End Sub

Private Sub cmdMaint10_KeyDown(KeyCode As Integer, Shift As Integer)

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaint10_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20120   intShiftDown = (Shift And acShiftMask) > 0
20130   intAltDown = (Shift And acAltMask) > 0
20140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20160     Select Case intRetVal
          Case vbKeyTab
20170       With Me
20180         intRetVal = 0
20190         .cmdMaint06.SetFocus
20200       End With
20210     End Select
20220   End If

        ' ** Shift keys.
20230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20240     Select Case intRetVal
          Case vbKeyTab
20250       With Me
20260         intRetVal = 0
20270         .cmdMaint04.SetFocus
20280       End With
20290     End Select
20300   End If

EXITP:
20310   KeyCode = intRetVal
20320   Exit Sub

ERRH:
20330   intRetVal = 0
20340   Select Case ERR.Number
        Case Else
20350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20360   End Select
20370   Resume EXITP

End Sub

Private Sub chkLoadTimer_AfterUpdate()

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_AfterUpdate"

20410   With Me
20420     If IsLoaded("frmLoadTimer", acForm) = False Then  ' ** Module Function: modFileUtilities.
            ' ** Don't turn it off if it's still playing out a timing sequence.
20430       gblnTimer = False
20440     End If
20450     Select Case .chkLoadTimer
          Case True
20460       .chkLoadTimer_lbl.FontBold = True
20470       If .chkLoadTimer.Visible = True Then  ' ** This may get turned off by the Form_Timer() sequence.
20480         gblnTimer = True
20490       End If
20500     Case False
20510       .chkLoadTimer_lbl.FontBold = False
20520     End Select
20530   End With

EXITP:
20540   Exit Sub

ERRH:
20550   Select Case ERR.Number
        Case Else
20560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20570   End Select
20580   Resume EXITP

End Sub

Private Sub chkLoadTimer_KeyDown(KeyCode As Integer, Shift As Integer)

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20620   intShiftDown = (Shift And acShiftMask) > 0
20630   intAltDown = (Shift And acAltMask) > 0
20640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20660     Select Case intRetVal
          Case vbKeyTab
20670       With Me
20680         intRetVal = 0
20690         .cmdClose.SetFocus
20700       End With
20710     End Select
20720   End If

        ' ** Shift keys.
20730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20740     Select Case intRetVal
          Case vbKeyTab
20750       With Me
20760         intRetVal = 0
20770         If .cmdMaint09.Visible = True And .cmdMaint09.Enabled = True Then
20780           .cmdMaint09.SetFocus
20790         Else
20800           .cmdMaint08.SetFocus
20810         End If
20820       End With
20830     End Select
20840   End If

EXITP:
20850   KeyCode = intRetVal
20860   Exit Sub

ERRH:
20870   intRetVal = 0
20880   Select Case ERR.Number
        Case Else
20890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20900   End Select
20910   Resume EXITP

End Sub

Public Sub RestoreForm()

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "RestoreForm"

21010   DoCmd.SelectObject acForm, THIS_NAME, False
21020   DoCmd.Restore

EXITP:
21030   Exit Sub

ERRH:
21040   Select Case ERR.Number
        Case Else
21050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21060   End Select
21070   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

21110   With Me
21120     blnGTR_Emblem = False
21130     For lngX = 1& To 24&
21140 On Error Resume Next
21150       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
21160       If ERR.Number <> 0 Then
21170 On Error GoTo ERRH
21180         Exit For
21190       Else
21200 On Error GoTo ERRH
21210       End If
21220     Next
21230   End With

EXITP:
21240   Exit Sub

ERRH:
21250   Select Case ERR.Number
        Case Else
21260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21270   End Select
21280   Resume EXITP

End Sub

Public Sub ProgBar_Width_Maint(dblWidth As Double, intMode As Integer)

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "ProgBar_Width_Maint"

        Dim strCtlName As String, blnVis As Boolean
        Dim lngX As Long

21310   With Me
21320     Select Case intMode
          Case 1
21330       blnVis = CBool(dblWidth)
21340       For lngX = 1& To 6&
21350         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
21360 On Error Resume Next
21370         .Controls(strCtlName).Visible = blnVis
21380 On Error GoTo ERRH
21390       Next
21400     Case 2
21410       For lngX = 1& To 6&
21420         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
21430         .Controls(strCtlName).Width = dblWidth
21440       Next
21450     End Select
21460   End With

EXITP:
21470   Exit Sub

ERRH:
21480   Select Case ERR.Number
        Case Else
21490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21500   End Select
21510   Resume EXITP

End Sub
