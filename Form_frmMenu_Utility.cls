VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Utility"

'VGC 10/11/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Admin Officer:    A {cmdAdministrativeOfficer}
' **   Backup:           B {cmdBackupData}
' **   Close:            C {cmdClose}
' **   Load Timer:       D {chkLoadTimer}
' **   Export:           E {cmdExport}
' **   Fee Schedules:    F {cmdFeeSchedules}
' **   Help:             H {cmdHelp}
' **   Inc/Exp Codes:    I {cmdIncomeExpenseCodes}
' **   Recurring Items:  M {cmdRecurringItems}
' **   License Update:   N {cmdUpdateLicense}
' **   Options:          O {cmdOptions}
' **   Compact:          P {cmdCompact}
' **   Restore:          R {cmdRestoreData}
' **   Special Maint:    S {cmdMaintenance}
' **   Locations:        T {cmdLocations}
' **   User Maint:       U {cmdUserMaintenance}
' **   Archive Trans:    V {cmdArchiveTransactions}
' **   Exit:             X {cmdClose}
' **   Company Info:     Y {cmdCompanyInformation}

' ** Shortcut Ctrl keys responsive from this form:
' **   Archive Reports:  V {cmdArchiveReports}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm}

' ***********************************************
' ** Easter Egg (Secret Key):
' ** cmdArchiveTransactions_KeyDown
' ***********************************************

Private blnMaintenance_Focus As Boolean

Private lngTpp As Long
Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnJustClose As Boolean, blnFormOpening As Boolean, blnNoArchive As Boolean
Private strDocName As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim grp As DAO.Group, usr As DAO.User
        Dim ctl As Access.Control
        Dim blnFound As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           gblnSignal = False

160           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

170           gstrCrtRpt_Version = vbNullString  ' ** Borrowing this variable from Court Reports.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
190           blnFound = False: blnJustClose = False: blnFormOpening = False: blnNoArchive = False
200           blnGTR_Emblem = False: lngGTR_Stat = 0&
210           blnMaintenance_Focus = False

              ' ** Moved this check from the frmMenu_Main.
220           For Each grp In DBEngine.Workspaces(0).Groups
230             If grp.Name = "ViewOnly" Or grp.Name = "DataEntry" Then
240               For Each usr In grp.Users
250                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
260                   blnFound = True
270                   Exit For
280                 End If
290               Next
300             End If
310           Next
320           If blnFound = True Then
                ' ** Limited access.
330             For Each ctl In .Controls
340               With ctl
350                 Select Case .ControlType
                    Case acCommandButton
360                   If .Name <> "cmdClose" And .Name <> "cmdUserMaintenance" And .Name <> "cmdHelp" Then
370                     .Visible = False
380                   End If
390                 End Select
400               End With
410             Next
420             .cmdHelp.Top = .cmdUserMaintenance.Top
430             .cmdHelp.Left = .cmdUserMaintenance.Left
440             .cmdUserMaintenance.Top = .cmdAdministrativeOfficer.Top
450             .cmdUserMaintenance.Left = .cmdAdministrativeOfficer.Left
460           End If

              ' ** Check the highest level group this user belongs to.
470           If CurrentGroup() = "Admins" Then  ' ** Module Function: modSecurityFunctions.
480             .cmdMaintenance.Visible = True
490           End If

500           If Len(TA_SEC) > Len(TA_SEC2) Then
510             If FileExists(CurrentAppPath & LNK_SEP & "TA Manual.pdf") = True Then  ' ** Module Function: modFileUtilities.
520               If CurrentAppName = (gstrFile_App & "." & gstrExt_AppDev) Then  ' ** Trust.mdb.
530                 .cmdHelp.Enabled = True
540               Else
550                 .cmdHelp.Enabled = False
560               End If
570             Else
580               .cmdHelp.Enabled = False
590             End If
600           End If

610           Select Case .chkLoadTimer_vis
              Case True
620             .chkLoadTimer.Visible = True
630             .chkLoadTimer_lbl1.Visible = True
640           Case False
650             .chkLoadTimer.Visible = False
660             .chkLoadTimer_lbl1.Visible = False
670           End Select
680           chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
690           DoEvents

700           DoCmd.Hourglass False

710         Else
720           Cancel = -1
730         End If
740       Else
750         Cancel = -1
760       End If
770     End With

780     If Cancel = -1 Then
790       Beep
800       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
810     End If

EXITP:
820     Set ctl = Nothing
830     Set usr = Nothing
840     Set grp = Nothing
850     Exit Sub

ERRH:
860     DoCmd.Hourglass False
870     Select Case ERR.Number
        Case 3059  ' ** Operation canceled by user.
          ' ** Ignore
880     Case Else
890       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
900     End Select
910     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer, Cancel As Integer

        Const THIS_PROC As String = "Form_KeyDown"

1010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1020    intShiftDown = (Shift And acShiftMask) > 0
1030    intAltDown = (Shift And acAltMask) > 0
1040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys responsive from this form:
        ' **   Load Timer:       D {chkLoadTimer}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys responsive from this form:
        ' **   Archive Reports:  V {cmdArchiveReports}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Timer Chk Vis:    T {Form_KeyDown}
        ' **   Close Form:       X {cmdCloseForm}

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
          Case vbKeyD
1130        With Me
1140          intRetVal = 0
1150          If .chkLoadTimer.Visible = True Then
1160            If .chkLoadTimer.Enabled = True Then
1170              .chkLoadTimer.SetFocus
1180            End If
1190            .chkLoadTimer = (Not .chkLoadTimer)
1200            chkLoadTimer_AfterUpdate
1210          End If
1220        End With
1230      Case vbKeyX
1240        intRetVal = 0
1250        cmdClose_Click  ' ** Procedure: Below.
1260      End Select
1270    End If

        ' ** Ctrl keys.
1280    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1290      Select Case intRetVal
          Case vbKeyV
1300        intRetVal = 0
1310        cmdArchiveReports_Click  ' ** Procedure: Below.
1320      End Select
1330    End If

        ' ** Ctrl-Shift keys.
1340    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1350      Select Case intRetVal
          Case vbKeyF
1360        With Me
1370          intRetVal = 0
1380          .FocusHolder.SetFocus
1390        End With
1400      Case vbKeyT
1410        With Me
1420          intRetVal = 0
1430          If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
1440            FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
1450          ElseIf CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
1460            Select Case .chkLoadTimer.Visible
                Case True
1470              .FocusHolder.SetFocus
1480              DoEvents
1490              .chkLoadTimer.Visible = False
1500              .chkLoadTimer_lbl1.Visible = False
1510              gblnTimer = False
1520            Case False
1530              .chkLoadTimer.Visible = True
1540              .chkLoadTimer_lbl1.Visible = True
1550              chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
1560            End Select
1570          End If
1580          .chkLoadTimer_vis = .chkLoadTimer.Visible
1590        End With
1600      Case vbKeyX
1610        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
1620          intRetVal = 0
1630          blnJustClose = True
1640          SetOption_Dev  ' ** Module Function: modUtilities.
1650          gblnSignal = True  ' ** Borrowing this variable.
1660          DoCmd.Close acForm, THIS_NAME
1670          DoCmd.Hourglass False
1680        End If
1690      End Select
1700    End If

EXITP:
1710    KeyCode = intRetVal
1720    Exit Sub

ERRH:
1730    intRetVal = 0
1740    Select Case ERR.Number
        Case Else
1750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1760    End Select
1770    Resume EXITP

End Sub

Private Sub Form_Timer()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim strFormName As String
        Dim blnRetVal As Boolean
        Dim strTmp01 As String

        Static lngTmp02 As Long, lngTmp03 As Long

1810    With Me
1820      .TimerInterval = 0&
1830      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
1840        lngGTR_Stat = lngGTR_Stat + 1&
1850        If lngTmp02 = 0& And lngTmp03 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1860        lngTmp02 = lngTmp02 + 1&
1870        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp02), 2) & "_img").Visible = True
1880        If lngTmp02 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1890        If lngTmp02 > 1& Then
1900          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp02 - 1&), 2) & "_img").Visible = False
1910        End If
1920        If lngTmp02 = 24& Then
1930          lngTmp02 = 0&
1940          lngTmp03 = lngTmp03 + 1&
1950          If lngTmp03 < 5& Then
1960            .TimerInterval = 50&
1970          Else
1980            blnGTR_Emblem = False
1990          End If
2000        Else
2010          .TimerInterval = 50&
2020        End If
2030      End If  ' ** blnGTR_Emblem.
2040      If gblnSetFocus = True Then
2050        gblnSetFocus = False
2060        strFormName = vbNullString
2070  On Error Resume Next
2080        Do Until Screen.ActiveForm.Name = strFormName
2090          DoCmd.SelectObject acForm, THIS_NAME, False
2100          strFormName = Screen.ActiveForm.Name
2110        Loop
2120  On Error GoTo ERRH
2130        If IsMaximized(Me) = True Then
2140          DoCmd.Restore
2150        End If
2160        DoEvents
2170      End If
2180      If blnGoingToReport = False Then
2190        Select Case gstrReturningForm
            Case "frmAdminOfficer"
2200          .cmdAdministrativeOfficer.SetFocus
2210        Case "frmUser"
2220          .cmdUserMaintenance.SetFocus
2230        Case "frmCompanyInfo"
2240          .cmdCompanyInformation.SetFocus
2250        Case "frmArchiveTransactions", "frmRpt_ArchivedTransactions"
2260          .cmdArchiveTransactions.SetFocus
2270        Case "frmLocations"
2280          .cmdLocations.SetFocus
2290        Case "frmFeeSchedules"
2300          .cmdFeeSchedules.SetFocus
2310        Case "frmRecurringItems"
2320          .cmdRecurringItems.SetFocus
2330        Case "frmIncomeExpenseCodes"
2340          .cmdIncomeExpenseCodes.SetFocus
2350        Case "frmAccountExport"
2360          .cmdExport.SetFocus
2370        Case "frmOptions"
2380          .cmdOptions.SetFocus
2390        Case "frmLicense"
2400          blnRetVal = Security_LicenseChk  ' ** Module Function: modSecurityFunctions
2410          Select Case blnRetVal
              Case True
2420            .cmdUpdateLicense.SetFocus
2430          Case False
2440            .TimerInterval = 0&
2450            QuitNow  ' ** Module Procedure: modStartupFuncs
2460          End Select
2470        Case "frmBackupRestore"
2480          Select Case gstrActNo  ' ** Borrowing this variable.
              Case "cmdBackupData_Click"
2490            .cmdBackupData.SetFocus
2500          Case "cmdRestoreData_Click"
2510            .cmdRestoreData.SetFocus
2520          End Select
2530          gstrActNo = vbNullString
2540        Case "frmLinkData"
2550          .cmdCompact.SetFocus
2560        Case "frmMenu_Maintenance"
2570          .cmdMaintenance.SetFocus
2580        Case Else
2590          If .cmdAdministrativeOfficer.Visible = True And .cmdAdministrativeOfficer.Enabled = True Then
2600            .cmdAdministrativeOfficer.SetFocus
2610          Else
2620            .cmdUserMaintenance.SetFocus
2630          End If
2640        End Select
2650        gstrReturningForm = vbNullString
2660        DoCmd.Hourglass False
2670      End If
2680      If gblnGoToReport = True Then
2690        DoCmd.Hourglass True  ' ** Make sure it's still running.
2700        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
2710        Else
2720          DoCmd.Hourglass True  ' ** Make sure it's still running.
2730          If blnGTR_Emblem = False Then
2740            blnGTR_Emblem = True
2750            lngTmp02 = 0&: lngTmp03 = 0&
2760          End If
2770          DoEvents
2780          Select Case blnGoingToReport
              Case True
2790            Select Case garr_varGoToReport(GTR_INV1)
                Case True
2800              Select Case garr_varGoToReport(GTR_CTL1)
                  Case "cmdAdministrativeOfficer"
2810                .TimerInterval = 50&
2820                cmdAdministrativeOfficer_Click  ' ** Procedure: Below.
2830              Case "cmdUserMaintenance"
2840                .TimerInterval = 50&
2850                cmdUserMaintenance_Click  ' ** Procedure: Below.
2860              Case "cmdArchiveTransactions", "cmdArchiveReports"
2870                .TimerInterval = 50&
2880                cmdArchiveTransactions_Click  ' ** Procedure: Below.
2890              Case "cmdLocations"
2900                .TimerInterval = 50&
2910                cmdLocations_Click  ' ** Procedure: Below.
2920              Case "cmdFeeSchedules"
2930                .TimerInterval = 50&
2940                cmdFeeSchedules_Click  ' ** Procedure: Below.
2950              Case "cmdRecurringItems"
2960                .TimerInterval = 50&
2970                cmdRecurringItems_Click  ' ** Procedure: Below.
2980              Case "cmdIncomeExpenseCodes"
2990                .TimerInterval = 50&
3000                cmdIncomeExpenseCodes_Click  ' ** Procedure: Below.
3010              Case "cmdMaintenance"
3020                .TimerInterval = 50&
3030                cmdMaintenance_Click  ' ** Procedure: Below.
3040              Case Else
                    ' ** Shouldn't happen.
3050                blnGoingToReport = False
3060                DoCmd.Hourglass False
3070                gblnGoToReport = False
3080                blnGTR_Emblem = False
3090                GTREmblem_Off  ' ** Procedure: Below.
3100                Beep
3110              End Select
3120            Case False
                  ' ** Just sit there, and we're done.
3130              blnGoingToReport = False
3140              DoCmd.Hourglass False
3150              gblnGoToReport = False
3160              blnGTR_Emblem = False
3170              GTREmblem_Off  ' ** Procedure: Below.
3180              Beep
3190            End Select
3200          Case False
3210            If garr_varGoToReport(GTR_FRM1) = THIS_NAME Then
                  ' ** This is first menu in path.
3220              blnGoingToReport = True
3230              Select Case IsNull(garr_varGoToReport(GTR_FRM2))
                  Case True
                    ' ** No more menus.
3240                strTmp01 = vbNullString
3250                If garr_varGoToReport(GTR_CTL1) = "cmdArchiveReports" Then
3260                  strTmp01 = "cmdArchiveTransactions"
3270                Else
3280                  strTmp01 = garr_varGoToReport(GTR_CTL1)
3290                End If
3300                .Controls(strTmp01).SetFocus
3310                lngBtn_Top = .Controls(strTmp01).Top
3320                lngBtn_Left = .Controls(strTmp01).Left
3330                Select Case lngBtn_Left
                    Case .cmdAdministrativeOfficer.Left
3340                  .GoToReport_arw_menu_l_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
3350                  .GoToReport_arw_menu_l_img.Visible = True
3360                Case .cmdExport.Left
3370                  .GoToReport_arw_menu_r_img.Top = lngBtn_Top + (3& * lngTpp)
3380                  .GoToReport_arw_menu_r_img.Visible = True
3390                End Select
3400                DoEvents
3410                .TimerInterval = 50&
3420              Case False
                    ' ** Moving on.
3430                .Controls(garr_varGoToReport(GTR_CTL1)).SetFocus
3440                lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL1)).Top
3450                lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL1)).Left
3460                If garr_varGoToReport(GTR_CTL1) = "cmdMaintenance" Then
3470                  .GoToReport_arw_menu_r_maint_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
3480                  .GoToReport_arw_menu_r_maint_img.Visible = True
3490                Else
3500                  Select Case lngBtn_Left
                      Case .cmdAdministrativeOfficer.Left
3510                    .GoToReport_arw_menu_l_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
3520                    .GoToReport_arw_menu_l_img.Visible = True
3530                  Case .cmdExport.Left
3540                    .GoToReport_arw_menu_r_img.Top = lngBtn_Top + (3& * lngTpp)
3550                    .GoToReport_arw_menu_r_img.Visible = True
3560                  End Select
3570                End If
3580                Select Case garr_varGoToReport(GTR_INV1)
                    Case True
3590                  DoEvents
3600                  If garr_varGoToReport(GTR_CTL1) = "cmdMaintenance" Then
3610                    .TimerInterval = 50&
3620                  Else
3630                    .TimerInterval = 50&
3640                  End If
3650                Case False
                      ' ** Just sit there, and we're done.
3660                  blnGoingToReport = False
3670                  DoCmd.Hourglass False
3680                  gblnGoToReport = False
3690                  blnGTR_Emblem = False
3700                  GTREmblem_Off  ' ** Procedure: Below.
3710                  Beep
3720                End Select
3730              End Select
3740            Else
                  ' ** Shouldn't be here, then!
3750              blnGoingToReport = False
3760              DoCmd.Hourglass False
3770              gblnGoToReport = False
3780              blnGTR_Emblem = False
3790              GTREmblem_Off  ' ** Procedure: Below.
3800              Beep
3810            End If
3820          End Select
3830        End If
3840      End If
3850      If gblnTimer = True Then
3860        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
3870          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
3880          gblnTimer = False
3890          DoCmd.Hourglass False
3900        End If
3910      End If
3920    End With

        'HERE! 1  Form_Open()  frmMenu_Utility
        'HERE! 2  Form_Timer()  frmMenu_Utility

EXITP:
3930    Exit Sub

ERRH:
3940    DoCmd.Hourglass False
3950    blnGoingToReport = False
3960    gblnGoToReport = False
3970    Select Case ERR.Number
        Case Else
3980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3990    End Select
4000    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

4110    With Me

4120      If gblnSignal = False Then  ' ** Borrowing this variable.
4130        DoCmd.Hourglass True
4140        DoEvents
4150      End If

          ' ** Make sure all reports are closed.
4160      Do While Reports.Count > 0
4170        DoCmd.Close acReport, Reports(0).Name
4180      Loop

4190      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

4200      blnGoingToReport = False

4210      If blnJustClose = False Then

4220        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
4230        DoEvents

4240        gblnAdmin = False
4250        gstrFormQuerySpec = vbNullString
4260        gblnSetFocus = True
4270        gstrReturningForm = THIS_NAME
4280        Select Case strCallingForm
            Case "frmMenu_Main"
4290          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Title"
4300        End Select

4310      End If  ' ** blnJustClose.

4320    End With

EXITP:
4330    Exit Sub

ERRH:
4340    DoCmd.Hourglass False
4350    Select Case ERR.Number
        Case Else
4360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4370    End Select
4380    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4410    With Me
4420      If blnMaintenance_Focus = False And gblnGoToReport = False Then
4430        If .cmdMaintenance_box01.Visible = True Then
4440          .cmdMaintenance_box01.Visible = False
4450          .cmdMaintenance_box02.Visible = False
4460        End If
4470      End If
4480      If .cmdClose.Caption <> "&Close" Then
4490        .cmdClose.Caption = "&Close"
4500      End If
4510    End With

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

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

Private Sub Header_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_box_MouseMove"

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

Private Sub Header_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_MouseMove"

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

Private Sub Header_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl2_MouseMove"

4910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub cmdClose_Click()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

5010    With Me
5020      If blnFormOpening = False Then

5030        DoCmd.Hourglass True
5040        DoEvents

5050        .TimerInterval = 0&

5060        blnJustClose = False
5070        DoCmd.Close acForm, THIS_NAME

5080      End If
5090    End With

EXITP:
5100    Exit Sub

ERRH:
5110    blnFormOpening = False
5120    DoCmd.Hourglass False
5130    Select Case ERR.Number
        Case Else
5140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5150    End Select
5160    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5220    intShiftDown = (Shift And acShiftMask) > 0
5230    intAltDown = (Shift And acAltMask) > 0
5240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5260      Select Case intRetVal
          Case vbKeyTab
5270        With Me
5280          intRetVal = 0
5290          If .cmdAdministrativeOfficer.Visible = True Then  ' ** Off if not Admins group.
5300            .cmdAdministrativeOfficer.SetFocus
5310          Else
5320            .cmdUserMaintenance.SetFocus  ' ** Everybody.
5330          End If
5340        End With
5350      End Select
5360    End If

        ' ** Shift keys.
5370    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5380      Select Case intRetVal
          Case vbKeyTab
5390        With Me
5400          intRetVal = 0
5410          If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
5420            .chkLoadTimer.SetFocus
5430          ElseIf .cmdMaintenance.Visible = True Then  ' ** Off if not logged in as Admin.
5440            .cmdMaintenance.SetFocus
5450          ElseIf .cmdHelp.Visible = True Then  ' ** Off if Demo version.
5460            .cmdHelp.SetFocus
5470          ElseIf .cmdCompact.Visible = True Then  ' ** Off if not Admins group.
5480            .cmdCompact.SetFocus
5490          Else
5500            .cmdUserMaintenance.SetFocus  ' ** Everybody.
5510          End If
5520        End With
5530      End Select
5540    End If

EXITP:
5550    KeyCode = intRetVal
5560    Exit Sub

ERRH:
5570    intRetVal = 0
5580    Select Case ERR.Number
        Case Else
5590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5600    End Select
5610    Resume EXITP

End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_MouseMove"

5710    With Me
5720      .cmdClose.Caption = "< Ba&ck"
5730    End With

EXITP:
5740    Exit Sub

ERRH:
5750    Select Case ERR.Number
        Case Else
5760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5770    End Select
5780    Resume EXITP

End Sub

Private Sub cmdUtility_l_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUtility_l_box01_MouseMove"

5810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5820    Exit Sub

ERRH:
5830    Select Case ERR.Number
        Case Else
5840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5850    End Select
5860    Resume EXITP

End Sub

Private Sub cmdAdministrativeOfficer_Click()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdministrativeOfficer_Click"

5910    With Me
5920      If blnFormOpening = False Then

5930        DoCmd.Hourglass True
5940        DoEvents

5950        .TimerInterval = 0&

5960        gblnSetFocus = True
5970        blnJustClose = True: blnFormOpening = True
5980        strDocName = "frmAdminOfficer"
5990        gstrReturningForm = strDocName
6000        DoCmd.Close acForm, THIS_NAME
6010        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6020        If gblnGoToReport = True Then
6030          Forms(strDocName).TimerInterval = 50&
6040          blnGoingToReport = False
6050          blnGTR_Emblem = False
6060        End If

6070        DoCmd.Hourglass False

6080      End If
6090    End With

EXITP:
6100    Exit Sub

ERRH:
6110    blnFormOpening = False
6120    DoCmd.Hourglass False
6130    Select Case ERR.Number
        Case Else
6140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6150    End Select
6160    Resume EXITP

End Sub

Private Sub cmdAdministrativeOfficer_KeyDown(KeyCode As Integer, Shift As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdministrativeOfficer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6220    intShiftDown = (Shift And acShiftMask) > 0
6230    intAltDown = (Shift And acAltMask) > 0
6240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6260      Select Case intRetVal
          Case vbKeyTab
6270        With Me
6280          intRetVal = 0
6290          .cmdUserMaintenance.SetFocus
6300        End With
6310      End Select
6320    End If

        ' ** Shift keys.
6330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6340      Select Case intRetVal
          Case vbKeyTab
6350        With Me
6360          intRetVal = 0
6370          .cmdClose.SetFocus
6380        End With
6390      End Select
6400    End If

EXITP:
6410    KeyCode = intRetVal
6420    Exit Sub

ERRH:
6430    intRetVal = 0
6440    Select Case ERR.Number
        Case Else
6450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6460    End Select
6470    Resume EXITP

End Sub

Private Sub cmdUserMaintenance_Click()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUserMaintenance_Click"

6510    With Me
6520      If blnFormOpening = False Then

6530        DoCmd.Hourglass True
6540        DoEvents

6550        .TimerInterval = 0&

6560        gblnSetFocus = True
6570        blnJustClose = True: blnFormOpening = True
6580        strDocName = "frmUser"
6590        gstrReturningForm = strDocName
6600        DoCmd.Close acForm, THIS_NAME
6610        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6620        If gblnGoToReport = True Then
6630          Forms(strDocName).TimerInterval = 50&
6640          blnGoingToReport = False
6650          blnGTR_Emblem = False
6660        End If

6670        DoCmd.Hourglass False

6680      End If
6690    End With

EXITP:
6700    Exit Sub

ERRH:
6710    blnFormOpening = False
6720    DoCmd.Hourglass False
6730    Select Case ERR.Number
        Case Else
6740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6750    End Select
6760    Resume EXITP

End Sub

Private Sub cmdUserMaintenance_KeyDown(KeyCode As Integer, Shift As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUserMaintenance_KeyDown"

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
6890          If .cmdCompanyInformation.Visible = True Then
6900            .cmdCompanyInformation.SetFocus
6910          ElseIf .cmdHelp.Enabled = True And .cmdHelp.Visible = True Then
6920            .cmdHelp.SetFocus
6930          Else
6940            .cmdClose.SetFocus
6950          End If
6960        End With
6970      End Select
6980    End If

        ' ** Shift keys.
6990    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7000      Select Case intRetVal
          Case vbKeyTab
7010        With Me
7020          intRetVal = 0
7030          If .cmdAdministrativeOfficer.Visible = True Then
7040            .cmdAdministrativeOfficer.SetFocus
7050          Else
7060            .cmdClose.SetFocus
7070          End If
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

Private Sub cmdCompanyInformation_Click()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCompanyInformation_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnAdd As Boolean
        Dim varTmp00 As Variant

7210    With Me
7220      If blnFormOpening = False Then

7230        DoCmd.Hourglass True
7240        DoEvents

7250        .TimerInterval = 0&

7260        blnAdd = False
7270        varTmp00 = DCount("*", "CompanyInformation")
7280        If IsNull(varTmp00) = True Then
7290          blnAdd = True
7300        Else
7310          If varTmp00 = 0 Then
7320            blnAdd = True
7330          End If
7340        End If
7350        Set dbs = CurrentDb
7360        With dbs
7370          If blnAdd = True Then
                ' ** Append first record to CompanyInformation, with name = 'Your Company'.
7380            Set qdf = .QueryDefs("qryCompanyInformation_16")
7390            qdf.Execute
7400            Set qdf = Nothing
7410          End If
7420          DoEvents
              ' ** Empty tmpCompanyInformation2.
7430          Set qdf = .QueryDefs("qryCompanyInformation_19")
7440          qdf.Execute
7450          Set qdf = Nothing
7460          DoEvents
              ' ** Append CompanyInformation to tmpCompanyInformation2.
7470          Set qdf = .QueryDefs("qryCompanyInformation_21")
7480          qdf.Execute
7490          Set qdf = Nothing
7500          DoEvents
7510          .Close
7520        End With
7530        Set dbs = Nothing

7540        blnJustClose = True: blnFormOpening = True
7550        strDocName = "frmCompanyInfo"
7560        gstrReturningForm = strDocName
7570        DoCmd.Close acForm, THIS_NAME
7580        DoCmd.OpenForm strDocName, , , , , , THIS_NAME

7590        DoCmd.Hourglass False

7600      End If
7610    End With

EXITP:
7620    Set qdf = Nothing
7630    Set dbs = Nothing
7640    Exit Sub

ERRH:
7650    blnFormOpening = False
7660    DoCmd.Hourglass False
7670    Select Case ERR.Number
        Case Else
7680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7690    End Select
7700    Resume EXITP

End Sub

Private Sub cmdArchiveTransactions_Click()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdArchiveTransactions_Click"

7810    With Me
7820      If blnFormOpening = False Then

7830        .TimerInterval = 0&

7840        If blnNoArchive = False Then

7850          DoCmd.Hourglass True
7860          DoEvents

7870          gblnSetFocus = True
7880          blnJustClose = True: blnFormOpening = True
7890          strDocName = "frmArchiveTransactions"
7900          gstrReturningForm = strDocName
7910          DoCmd.Close acForm, THIS_NAME
7920          DoCmd.OpenForm strDocName, , , , , , THIS_NAME
7930          If gblnGoToReport = True Then
7940            Forms(strDocName).TimerInterval = 50&
7950            blnGoingToReport = False
7960            blnGTR_Emblem = False
7970          End If

7980        Else
7990          blnNoArchive = False
8000          Beep
8010        End If

8020        DoCmd.Hourglass False

8030      End If
8040    End With

EXITP:
8050    Exit Sub

ERRH:
8060    blnFormOpening = False
8070    DoCmd.Hourglass False
8080    Select Case ERR.Number
        Case Else
8090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8100    End Select
8110    Resume EXITP

End Sub

Private Sub cmdArchiveTransactions_KeyDown(KeyCode As Integer, Shift As Integer)
' ** Easter Egg (Secret Key) to open frmRpt_ArchivedTransactions directly.

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdArchiveTransactions_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8220    intShiftDown = (Shift And acShiftMask) > 0
8230    intAltDown = (Shift And acAltMask) > 0
8240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Ctrl keys.
8250    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8260      Select Case intRetVal
          Case vbKeyReturn
8270        intRetVal = 0
8280        cmdArchiveReports_Click  ' ** Procedure: Below.
8290      Case Else
8300        blnNoArchive = False
8310      End Select
8320    End If

EXITP:
8330    KeyCode = intRetVal
8340    Exit Sub

ERRH:
8350    intRetVal = 0
8360    Select Case ERR.Number
        Case Else
8370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8380    End Select
8390    Resume EXITP

End Sub

Private Sub cmdArchiveTransactions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdArchiveTransactions_MouseDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer

        ' ** Use bit masks to determine which key was pressed.
8410    intShiftDown = (Shift And acShiftMask) > 0
8420    intAltDown = (Shift And acAltMask) > 0
8430    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Ctrl keys.
8440    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8450      cmdArchiveReports_Click  ' ** Procedure: Below.
8460    Else
8470      blnNoArchive = False
8480    End If

EXITP:
8490    Exit Sub

ERRH:
8500    Select Case ERR.Number
        Case Else
8510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8520    End Select
8530    Resume EXITP

End Sub

Private Sub cmdLocations_Click()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLocations_Click"

8610    With Me
8620      If blnFormOpening = False Then

8630        DoCmd.Hourglass True
8640        DoEvents

8650        .TimerInterval = 0&

8660        gblnSetFocus = True
8670        blnJustClose = True: blnFormOpening = True
8680        strDocName = "frmLocations"
8690        gstrReturningForm = strDocName
8700        DoCmd.Close acForm, THIS_NAME
8710        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
8720        If gblnGoToReport = True Then
8730          Forms(strDocName).TimerInterval = 50&
8740          blnGoingToReport = False
8750          blnGTR_Emblem = False
8760        End If

8770        DoCmd.Hourglass False

8780      End If
8790    End With

EXITP:
8800    Exit Sub

ERRH:
8810    blnFormOpening = False
8820    DoCmd.Hourglass False
8830    Select Case ERR.Number
        Case Else
8840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8850    End Select
8860    Resume EXITP

End Sub

Private Sub cmdFeeSchedules_Click()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdFeeSchedules_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngScheds As Long, arr_varSched As Variant
        Dim lngSechduleID As Long, lngOrd As Long
        Dim blnAddAll As Boolean
        Dim lngX As Long

        ' ** Array: arr_varSched().
        Const S_SID   As Integer = 0
        Const S_SDID  As Integer = 1
        Const S_ORD   As Integer = 2

8910    With Me
8920      If blnFormOpening = False Then

8930        DoCmd.Hourglass True
8940        DoEvents

8950        .TimerInterval = 0&

8960        blnAddAll = False
8970        Set dbs = CurrentDb
8980        With dbs
              ' ** Delete qryFeeSchedule_16_03 (tblScheduleDetail_Staging, not in tblScheduleDetail).
8990          Set qdf = dbs.QueryDefs("qryFeeSchedule_16_04")
9000          qdf.Execute
9010          Set qdf = Nothing
9020          Set rst = .OpenRecordset("tblScheduleDetail_Staging", dbOpenDynaset, dbReadOnly)
9030          With rst
9040            If .BOF = True And .EOF = True Then
9050              blnAddAll = True
9060            End If
9070            .Close
9080          End With
9090          Set rst = Nothing
9100          If blnAddAll = True Then
                ' ** ScheduleDetail, just needed fields, sorted.
9110            Set qdf = .QueryDefs("qryFeeSchedule_07_01")
9120            Set rst = qdf.OpenRecordset
9130            With rst
9140              If .BOF = True And .EOF = True Then
                    ' ** No current fee schedules.
9150                lngScheds = 0&
9160              Else
9170                .MoveLast
9180                lngScheds = .RecordCount
9190                .MoveFirst
9200                arr_varSched = .GetRows(lngScheds)
                    ' ******************************************************
                    ' ** Array: arr_varSched()
                    ' **
                    ' **   Field  Element  Name                 Constant
                    ' **   =====  =======  ===================  ==========
                    ' **     1       0     Schedule_ID          S_SID
                    ' **     2       1     ScheduleDetail_ID    S_SDID
                    ' **     3       2     scheddets_order      S_ORD
                    ' **
                    ' ******************************************************
9210              End If
9220              .Close
9230            End With
9240            Set rst = Nothing
9250            Set qdf = Nothing
9260            If lngScheds > 0& Then
9270              lngSechduleID = 0&: lngOrd = 0&
9280              For lngX = 0& To (lngScheds - 1&)
9290                If arr_varSched(S_SID, lngX) <> lngSechduleID Then
9300                  lngOrd = 0&
9310                  lngSechduleID = arr_varSched(S_SID, lngX)
9320                End If
9330                lngOrd = lngOrd + 1&
9340                arr_varSched(S_ORD, lngX) = lngOrd
9350              Next
9360              Set rst = .OpenRecordset("tblScheduleDetail_Staging", dbOpenDynaset, dbConsistent)
9370              With rst
9380                For lngX = 0& To (lngScheds - 1&)
9390                  .AddNew
9400                  ![Schedule_ID] = arr_varSched(S_SID, lngX)
9410                  ![ScheduleDetail_ID] = arr_varSched(S_SDID, lngX)
                      ' ** ![scheddets_id] : AutoNumber.
9420                  ![scheddets_order] = arr_varSched(S_ORD, lngX)
9430                  ![scheddets_datemodified] = Now()
9440                  .Update
9450                Next
9460                .Close
9470              End With
9480              Set rst = Nothing
9490            End If
9500          End If
9510          .Close
9520        End With
9530        Set dbs = Nothing

9540        gblnSetFocus = True
9550        blnJustClose = True: blnFormOpening = True
9560        strDocName = "frmFeeSchedules"
9570        gstrReturningForm = strDocName
9580        DoCmd.Close acForm, THIS_NAME
9590        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
9600        If gblnGoToReport = True Then
9610          Forms(strDocName).TimerInterval = 50&
9620          blnGoingToReport = False
9630          blnGTR_Emblem = False
9640        End If

9650        DoCmd.Hourglass False

9660      End If
9670    End With

EXITP:
9680    Set rst = Nothing
9690    Set qdf = Nothing
9700    Set dbs = Nothing
9710    Exit Sub

ERRH:
9720    blnFormOpening = False
9730    DoCmd.Hourglass False
9740    Select Case ERR.Number
        Case Else
9750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9760    End Select
9770    Resume EXITP

End Sub

Private Sub cmdRecurringItems_Click()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRecurringItems_Click"

9810    With Me
9820      If blnFormOpening = False Then

9830        DoCmd.Hourglass True
9840        DoEvents

9850        .TimerInterval = 0&

9860        gblnSetFocus = True
9870        blnJustClose = True: blnFormOpening = True
9880        strDocName = "frmRecurringItems"
9890        gstrReturningForm = strDocName
9900        DoCmd.Close acForm, THIS_NAME
9910        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
9920        If gblnGoToReport = True Then
9930          Forms(strDocName).TimerInterval = 50&
9940          blnGoingToReport = False
9950          blnGTR_Emblem = False
9960        End If

9970        DoCmd.Hourglass False

9980      End If
9990    End With

EXITP:
10000   Exit Sub

ERRH:
10010   blnFormOpening = False
10020   DoCmd.Hourglass False
10030   Select Case ERR.Number
        Case Else
10040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10050   End Select
10060   Resume EXITP

End Sub

Private Sub cmdIncomeExpenseCodes_Click()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdIncomeExpenseCodes_Click"

10110   With Me
10120     If blnFormOpening = False Then

10130       DoCmd.Hourglass True
10140       DoEvents

10150       .TimerInterval = 0&

10160       gblnSetFocus = True
10170       blnJustClose = True: blnFormOpening = True
10180       strDocName = "frmIncomeExpenseCodes"
10190       gstrReturningForm = strDocName
10200       DoCmd.Close acForm, THIS_NAME
10210       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
10220       If gblnGoToReport = True Then
10230         Forms(strDocName).TimerInterval = 50&
10240         blnGoingToReport = False
10250         blnGTR_Emblem = False
10260       End If

10270       DoCmd.Hourglass False

10280     End If
10290   End With

EXITP:
10300   Exit Sub

ERRH:
10310   blnFormOpening = False
10320   DoCmd.Hourglass False
10330   Select Case ERR.Number
        Case Else
10340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10350   End Select
10360   Resume EXITP

End Sub

Private Sub cmdIncomeExpenseCodes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdIncomeExpenseCodes_MouseMove"

10410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10420   Exit Sub

ERRH:
10430   Select Case ERR.Number
        Case Else
10440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10450   End Select
10460   Resume EXITP

End Sub

Private Sub cmdUtility_r_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUtility_r_box01_MouseMove"

10510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10520   Exit Sub

ERRH:
10530   Select Case ERR.Number
        Case Else
10540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10550   End Select
10560   Resume EXITP

End Sub

Private Sub cmdExport_Click()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExport_Click"

10610   With Me
10620     If blnFormOpening = False Then

10630       DoCmd.Hourglass True
10640       DoEvents

10650       .TimerInterval = 0&

10660       gblnSetFocus = True
10670       blnJustClose = True: blnFormOpening = True
10680       strDocName = "frmAccountExport"
10690       gstrReturningForm = strDocName
10700       DoCmd.Close acForm, THIS_NAME
10710       DoCmd.OpenForm strDocName, , , , , , THIS_NAME

10720       DoCmd.Hourglass False

10730     End If
10740   End With

EXITP:
10750   Exit Sub

ERRH:
10760   blnFormOpening = False
10770   DoCmd.Hourglass False
10780   Select Case ERR.Number
        Case Else
10790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10800   End Select
10810   Resume EXITP

End Sub

Private Sub cmdOptions_Click()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOptions_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim varTmp00 As Variant

10910   With Me
10920     If blnFormOpening = False Then

10930       DoCmd.Hourglass True
10940       DoEvents

10950       .TimerInterval = 0&

10960       blnFormOpening = True

10970       Set dbs = CurrentDb
10980       With dbs
              ' ** Empty tblCompanyInformation_Staging.
10990         Set qdf = .QueryDefs("qryCompanyInformation_20")
11000         qdf.Execute
11010         gblnMessage = False
11020         DoEvents
11030         Set rst = .OpenRecordset("CompanyInformation", dbOpenDynaset, dbConsistent)
11040         With rst
11050           If .BOF = True And .EOF = True Then
                  ' ** There should always be 1 record, even if it's empty.
11060             gblnMessage = True
11070             varTmp00 = DLookup("[License name]", "License Name")  ' ** 100 Characters.
11080             .AddNew
11090             ![CoInfo_Name] = Left(varTmp00, 50)  ' ** 50 Characters.
11100             ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
11110             ![CoInfo_DateModified] = Now()
11120             .Update
11130           Else
11140             .MoveFirst
11150             If IsNull(![CoInfo_Name]) = True Then
11160               gblnMessage = True
11170               varTmp00 = DLookup("[License name]", "License Name")
11180               .Edit
11190               ![CoInfo_Name] = Left(varTmp00, 50)
11200               ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
11210               ![CoInfo_DateModified] = Now()
11220               .Update
11230             End If
11240           End If
11250           DoEvents
11260           .Close
11270         End With
11280         Set rst = Nothing
              ' ** Append CompanyInformation to tblCompanyInformation_Staging.
11290         Set qdf = .QueryDefs("qryCompanyInformation_18")
11300         qdf.Execute
11310         .Close
11320       End With
11330       Set qdf = Nothing
11340       Set dbs = Nothing
11350       DoEvents

11360       gblnSetFocus = True
11370       blnJustClose = True
11380       strDocName = "frmOptions"
11390       gstrReturningForm = strDocName
11400       DoCmd.Close acForm, THIS_NAME
11410       DoCmd.OpenForm strDocName, , , , , , THIS_NAME

11420       DoCmd.Hourglass False

11430     End If
11440   End With

EXITP:
11450   Set rst = Nothing
11460   Set qdf = Nothing
11470   Set dbs = Nothing
11480   Exit Sub

ERRH:
11490   blnFormOpening = False
11500   DoCmd.Hourglass False
11510   Select Case ERR.Number
        Case Else
11520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11530   End Select
11540   Resume EXITP

End Sub

Private Sub cmdUpdateLicense_Click()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdateLicense_Click"

11610   With Me
11620     If blnFormOpening = False Then

11630       DoCmd.Hourglass True
11640       DoEvents

11650       .TimerInterval = 0&

11660       blnJustClose = True: blnFormOpening = True
11670       strDocName = "frmLicense"
11680       gstrReturningForm = strDocName
11690       DoCmd.Close acForm, THIS_NAME
11700       DoCmd.OpenForm strDocName, , , , , , THIS_NAME

11710       DoCmd.Hourglass False

11720     End If
11730   End With

EXITP:
11740   Exit Sub

ERRH:
11750   blnFormOpening = False
11760   DoCmd.Hourglass False
11770   Select Case ERR.Number
        Case Else
11780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11790   End Select
11800   Resume EXITP

End Sub

Public Sub cmdBackupData_Click()

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBackupData_Click"

11910   With Me
11920     If blnFormOpening = False Then

11930       DoCmd.Hourglass True
11940       DoEvents

11950       .TimerInterval = 0&

11960       gblnSetFocus = True
11970       blnJustClose = True: blnFormOpening = True
11980       strDocName = "frmBackupRestore"
11990       gstrActNo = THIS_PROC  ' ** Borrowing this variable.
12000       gstrReturningForm = strDocName
12010       DoCmd.Close acForm, THIS_NAME
12020       DoCmd.OpenForm strDocName, , , , , , (THIS_NAME & "~" & THIS_PROC)

12030       DoCmd.Hourglass False

12040     End If
12050   End With

EXITP:
12060   Exit Sub

ERRH:
12070   blnFormOpening = False
12080   DoCmd.Hourglass False
12090   Select Case ERR.Number
        Case Else
12100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12110   End Select
12120   Resume EXITP

End Sub

Public Sub cmdRestoreData_Click()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRestoreData_Click"

12210   With Me
12220     If blnFormOpening = False Then

12230       DoCmd.Hourglass True
12240       DoEvents

12250       .TimerInterval = 0&

12260       gblnSetFocus = True
12270       blnJustClose = True: blnFormOpening = True
12280       strDocName = "frmBackupRestore"
12290       gstrActNo = THIS_PROC  ' ** Borrowing this variable.
12300       gstrReturningForm = strDocName
12310       DoCmd.Close acForm, THIS_NAME
12320       DoCmd.OpenForm strDocName, , , , , , (THIS_NAME & "~" & THIS_PROC)

12330       DoCmd.Hourglass False

12340     End If
12350   End With

EXITP:
12360   Exit Sub

ERRH:
12370   blnFormOpening = False
12380   DoCmd.Hourglass False
12390   Select Case ERR.Number
        Case Else
12400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12410   End Select
12420   Resume EXITP

End Sub

Private Sub cmdCompact_Click()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCompact_Click"

12510   With Me
12520     If blnFormOpening = False Then

12530       DoCmd.Hourglass True
12540       DoEvents

12550       .TimerInterval = 0&

12560       blnJustClose = True: blnFormOpening = True
12570       strDocName = "frmLinkData"
12580       gstrReturningForm = strDocName
12590       DoCmd.Close acForm, THIS_NAME
12600       DoCmd.OpenForm strDocName, , , , , , THIS_NAME

12610       DoCmd.Hourglass False

12620     End If
12630   End With

EXITP:
12640   Exit Sub

ERRH:
12650   blnFormOpening = False
12660   DoCmd.Hourglass False
12670   Select Case ERR.Number
        Case Else
12680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12690   End Select
12700   Resume EXITP

End Sub

Private Sub cmdCompact_KeyDown(KeyCode As Integer, Shift As Integer)

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCompact_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12820   intShiftDown = (Shift And acShiftMask) > 0
12830   intAltDown = (Shift And acAltMask) > 0
12840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12860     Select Case intRetVal
          Case vbKeyTab
12870       With Me
12880         intRetVal = 0
12890         If .cmdHelp.Enabled = True And .cmdHelp.Visible = True Then
12900           .cmdHelp.SetFocus
12910         ElseIf .cmdMaintenance.Visible = True Then
12920           .cmdMaintenance.SetFocus
12930         Else
12940           .cmdClose.SetFocus
12950         End If
12960       End With
12970     End Select
12980   End If

        ' ** Shift keys.
12990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13000     Select Case intRetVal
          Case vbKeyTab
13010       With Me
13020         intRetVal = 0
13030         .cmdRestoreData.SetFocus
13040       End With
13050     End Select
13060   End If

EXITP:
13070   KeyCode = intRetVal
13080   Exit Sub

ERRH:
13090   intRetVal = 0
13100   Select Case ERR.Number
        Case Else
13110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13120   End Select
13130   Resume EXITP

End Sub

Private Sub cmdHelp_Click()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_Click"

13210   With Me
13220     If blnFormOpening = False Then

13230       DoCmd.Hourglass True
13240       DoEvents

13250       .TimerInterval = 0&

13260       blnFormOpening = True
13270       Help_OpenAll  ' ** Module Function: modHelpFuncs.
13280       DoEvents
13290       blnFormOpening = False

13300       DoCmd.Hourglass False

13310     End If
13320   End With

EXITP:
13330   Exit Sub

ERRH:
13340   blnFormOpening = False
13350   DoCmd.Hourglass False
13360   Select Case ERR.Number
        Case Else
13370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13380   End Select
13390   Resume EXITP

End Sub

Private Sub cmdHelp_KeyDown(KeyCode As Integer, Shift As Integer)

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13420   intShiftDown = (Shift And acShiftMask) > 0
13430   intAltDown = (Shift And acAltMask) > 0
13440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13460     Select Case intRetVal
          Case vbKeyTab
13470       With Me
13480         intRetVal = 0
13490         If .cmdMaintenance.Visible = True Then
13500           .cmdMaintenance.SetFocus
13510         ElseIf .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
13520           .chkLoadTimer.SetFocus
13530         Else
13540           .cmdClose.SetFocus
13550         End If
13560       End With
13570     End Select
13580   End If

        ' ** Shift keys.
13590   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13600     Select Case intRetVal
          Case vbKeyTab
13610       With Me
13620         intRetVal = 0
13630         If .cmdCompact.Visible = True Then
13640           .cmdCompact.SetFocus
13650         Else
13660           .cmdUserMaintenance.SetFocus
13670         End If
13680       End With
13690     End Select
13700   End If

EXITP:
13710   KeyCode = intRetVal
13720   Exit Sub

ERRH:
13730   intRetVal = 0
13740   Select Case ERR.Number
        Case Else
13750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13760   End Select
13770   Resume EXITP

End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseMove"

13810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13820   Exit Sub

ERRH:
13830   Select Case ERR.Number
        Case Else
13840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13850   End Select
13860   Resume EXITP

End Sub

Private Sub cmdMaintenance_GotFocus()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaintenance_GotFocus"

13910   With Me
13920     blnMaintenance_Focus = True
13930     .cmdMaintenance_box01.Visible = True
13940     .cmdMaintenance_box02.Visible = True
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

Private Sub cmdMaintenance_Click()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaintenance_Click"

14110   With Me
14120     If blnFormOpening = False Then

14130       strDocName = "frmMenu_Maintenance"

14140       FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
14150       DoEvents

14160       If .chkLoadTimer.Visible = True Then
14170         chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
14180         If gblnTimer = True Then
                'gblnSetFocus = True
14190           DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
14200         End If
14210       End If
14220       DoEvents

14230       DoCmd.Hourglass True
14240       DoEvents

14250       .TimerInterval = 0&

14260       gblnSetFocus = True
14270       blnJustClose = True: blnFormOpening = True
14280       gstrReturningForm = strDocName
14290       DoCmd.Close acForm, THIS_NAME
14300       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
14310       If gblnGoToReport = True Then
14320         Forms(strDocName).TimerInterval = 50&
14330         blnGoingToReport = False
14340         blnGTR_Emblem = False
14350       End If

14360       DoCmd.Hourglass False

14370     End If
14380   End With

EXITP:
14390   Exit Sub

ERRH:
14400   blnFormOpening = False
14410   DoCmd.Hourglass False
14420   Select Case ERR.Number
        Case Else
14430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14440   End Select
14450   Resume EXITP

End Sub

Private Sub cmdMaintenance_KeyDown(KeyCode As Integer, Shift As Integer)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaintenance_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14520   intShiftDown = (Shift And acShiftMask) > 0
14530   intAltDown = (Shift And acAltMask) > 0
14540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14560     Select Case intRetVal
          Case vbKeyTab
14570       With Me
14580         intRetVal = 0
14590         If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
14600           .chkLoadTimer.SetFocus
14610         Else
14620           .cmdClose.SetFocus
14630         End If
14640       End With
14650     End Select
14660   End If

        ' ** Shift keys.
14670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14680     Select Case intRetVal
          Case vbKeyTab
14690       With Me
14700         intRetVal = 0
14710         If .cmdHelp.Enabled = True And .cmdHelp.Visible = True Then
14720           .cmdHelp.SetFocus
14730         Else
14740           .cmdCompact.SetFocus
14750         End If
14760       End With
14770     End Select
14780   End If

EXITP:
14790   KeyCode = intRetVal
14800   Exit Sub

ERRH:
14810   intRetVal = 0
14820   Select Case ERR.Number
        Case Else
14830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14840   End Select
14850   Resume EXITP

End Sub

Private Sub cmdMaintenance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaintenance_MouseMove"

14910   With Me
14920     If .cmdMaintenance_box01.Visible = False Then
14930       .cmdMaintenance_box01.Visible = True
14940       .cmdMaintenance_box02.Visible = True
14950     End If
14960   End With

EXITP:
14970   Exit Sub

ERRH:
14980   Select Case ERR.Number
        Case Else
14990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15000   End Select
15010   Resume EXITP

End Sub

Private Sub cmdMaintenance_LostFocus()

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMaintenance_LostFocus"

15110   With Me
15120     .cmdMaintenance_box01.Visible = False
15130     .cmdMaintenance_box02.Visible = False
15140     blnMaintenance_Focus = False
15150   End With

EXITP:
15160   Exit Sub

ERRH:
15170   Select Case ERR.Number
        Case Else
15180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15190   End Select
15200   Resume EXITP

End Sub

Private Sub chkLoadTimer_AfterUpdate()

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_AfterUpdate"

15310   With Me
15320     If IsLoaded("frmLoadTimer", acForm) = False Then  ' ** Module Function: modFileUtilities.
            ' ** Don't turn it off if it's still playing out a timing sequence.
15330       gblnTimer = False
15340     End If
15350     Select Case .chkLoadTimer
          Case True
15360       .chkLoadTimer_lbl.FontBold = True
15370       If .chkLoadTimer.Visible = True Then  ' ** This may get turned off by the Form_Timer() sequence.
15380         gblnTimer = True
15390       End If
15400     Case False
15410       .chkLoadTimer_lbl.FontBold = False
15420     End Select
15430   End With

EXITP:
15440   Exit Sub

ERRH:
15450   Select Case ERR.Number
        Case Else
15460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15470   End Select
15480   Resume EXITP

End Sub

Private Sub chkLoadTimer_KeyDown(KeyCode As Integer, Shift As Integer)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15520   intShiftDown = (Shift And acShiftMask) > 0
15530   intAltDown = (Shift And acAltMask) > 0
15540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15560     Select Case intRetVal
          Case vbKeyTab
15570       With Me
15580         intRetVal = 0
15590         .cmdClose.SetFocus
15600       End With
15610     End Select
15620   End If

        ' ** Shift keys.
15630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15640     Select Case intRetVal
          Case vbKeyTab
15650       With Me
15660         intRetVal = 0
15670         .cmdMaintenance.SetFocus
15680       End With
15690     End Select
15700   End If

EXITP:
15710   KeyCode = intRetVal
15720   Exit Sub

ERRH:
15730   intRetVal = 0
15740   Select Case ERR.Number
        Case Else
15750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15760   End Select
15770   Resume EXITP

End Sub

Private Sub cmdArchiveReports_Click()
' ** This is a fake Click event.

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdArchiveReports_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngArchiveNum As Long, datArchiveDate As Date

15810   With Me
15820     If blnFormOpening = False Then

15830       DoCmd.Hourglass True
15840       DoEvents

15850       .TimerInterval = 0&

15860       blnFormOpening = True

15870       Set dbs = CurrentDb
15880       With dbs
              ' ** LedgerArchive, grouped, with max transdate, cnt.
15890         Set qdf = .QueryDefs("qryArchiveTransactions_02")
15900         Set rst = qdf.OpenRecordset
15910         With rst
15920           If .BOF = True And .EOF = True Then
15930             lngArchiveNum = 0&
15940             datArchiveDate = 0
15950           Else
15960             .MoveFirst
15970             lngArchiveNum = ![cnt]
15980             If ![cnt] > 0 Then
15990               datArchiveDate = ![transdate]
16000             Else
16010               datArchiveDate = 0
16020             End If
16030           End If
16040           .Close
16050         End With
16060         .Close
16070       End With

            ' ** I'm gonna go there anyway!
16080       gblnSetFocus = True
16090       blnNoArchive = True  ' ** This means it's not going through frmArchiveTransactions.
16100       blnJustClose = True
16110       strDocName = "frmRpt_ArchivedTransactions"
16120       gstrReturningForm = strDocName
16130       DoCmd.Close acForm, THIS_NAME
16140       DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & CStr(CLng(datArchiveDate))

16150       DoCmd.Hourglass False

16160     End If
16170   End With

EXITP:
16180   Exit Sub

ERRH:
16190   blnFormOpening = False
16200   DoCmd.Hourglass False
16210   Select Case ERR.Number
        Case Else
16220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16230   End Select
16240   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

16310   With Me
16320     blnGTR_Emblem = False
16330     For lngX = 1& To 24&
16340 On Error Resume Next
16350       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
16360       If ERR.Number <> 0 Then
16370 On Error GoTo ERRH
16380         Exit For
16390       Else
16400 On Error GoTo ERRH
16410       End If
16420     Next
16430   End With

EXITP:
16440   Exit Sub

ERRH:
16450   Select Case ERR.Number
        Case Else
16460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16470   End Select
16480   Resume EXITP

End Sub
